//! Best-effort Apple iWork format detection.
//!
//! Detection validates application-root evidence plus legacy bundle markers.
//! It inspects only the root archive envelope, not document content, and does
//! not expose the ZIP implementation.

use crate::archive::Archive;
use crate::registry::{Application, detect_application_from_document};
use crate::snappy::SnappyStream;
use crate::zip_utils::{is_encrypted_iwork_archive, nested_index_zip_name};
use soapberry_zip::office::{ArchiveLimits, ArchiveReader};
use std::fs::{self, File};
use std::io::{Cursor, Read, Seek, SeekFrom};
use std::path::Path;

const MAX_INPUT_BYTES: u64 = 1024 * 1024 * 1024;
const ARCHIVE_LIMITS: ArchiveLimits = ArchiveLimits {
    max_files: 100_000,
    max_entry_size: 512 * 1024 * 1024,
    max_total_size: 2 * 1024 * 1024 * 1024,
};

/// The application family of a detected iWork document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Format {
    /// Apple Pages.
    Pages,
    /// Apple Keynote.
    Keynote,
    /// Apple Numbers.
    Numbers,
}

/// Detect an iWork application from complete packaged bytes.
///
/// ZIP and Snappy metadata are validated under explicit file-count and size
/// limits. A package with conflicting application-root evidence is rejected.
pub fn bytes(value: &[u8]) -> Option<Format> {
    let archive = ArchiveReader::new_with_limits(value, ARCHIVE_LIMITS).ok()?;
    match classify_archive(&archive, true) {
        Outcome::Found(format) => Some(format),
        Outcome::None | Outcome::Conflict => None,
    }
}

fn classify_archive(archive: &ArchiveReader<'_>, allow_nested: bool) -> Outcome {
    if is_encrypted_iwork_archive(archive) {
        return Outcome::Conflict;
    }

    let marks = marks(archive.file_names());
    if marks.iwa {
        return classify_direct_archive(archive, marks);
    }
    if !allow_nested {
        return Outcome::None;
    }

    let index_name = match nested_index_zip_name(archive) {
        Ok(Some(name)) => name,
        Ok(None) => return Outcome::None,
        Err(_) => return Outcome::Conflict,
    };
    let index_data = match archive.read(&index_name) {
        Ok(data) => data,
        Err(_) => return Outcome::Conflict,
    };
    let index = match ArchiveReader::new_with_limits(&index_data, ARCHIVE_LIMITS) {
        Ok(index) => index,
        Err(_) => return Outcome::Conflict,
    };
    classify_archive(&index, false)
}

fn classify_direct_archive(archive: &ArchiveReader<'_>, marks: Marks) -> Outcome {
    let mut documents = archive
        .file_names()
        .filter(|name| index_name(name) == Some("Document.iwa"));
    let Some(document_name) = documents.next() else {
        return Outcome::None;
    };
    if documents.next().is_some() {
        return Outcome::Conflict;
    }

    let data = match archive.read(document_name) {
        Ok(data) => data,
        Err(_) => return Outcome::Conflict,
    };
    let Some(format) = root_format(&data) else {
        return Outcome::Conflict;
    };
    if marks.accepts(format) {
        Outcome::Found(format)
    } else {
        Outcome::Conflict
    }
}

fn root_format(data: &[u8]) -> Option<Format> {
    let stream = SnappyStream::decompress(&mut Cursor::new(data)).ok()?;
    let archive = Archive::parse(stream.data()).ok()?;
    let mut detected = None;

    for application in archive
        .objects
        .iter()
        .filter(|object| object.archive_info.identifier == Some(1))
        .flat_map(|object| &object.messages)
        .filter_map(|message| detect_application_from_document(&message.data))
    {
        let format = match application {
            Application::Pages => Format::Pages,
            Application::Keynote => Format::Keynote,
            Application::Numbers => Format::Numbers,
            Application::Common => continue,
        };
        if detected.is_some_and(|previous| previous != format) {
            return None;
        }
        detected = Some(format);
    }

    detected
}

/// Detect an iWork application from a seekable stream.
///
/// Detection starts at byte zero and restores the caller's original cursor on
/// every path. Streams larger than one GiB are rejected without being read.
pub fn reader<R: Read + Seek>(value: &mut R) -> Option<Format> {
    let original = value.stream_position().ok()?;
    let detected = (|| {
        let length = value.seek(SeekFrom::End(0)).ok()?;
        if length > MAX_INPUT_BYTES {
            return None;
        }

        let length = usize::try_from(length).ok()?;
        let mut data = Vec::new();
        data.try_reserve_exact(length).ok()?;
        data.resize(length, 0);

        value.seek(SeekFrom::Start(0)).ok()?;
        value.read_exact(&mut data).ok()?;
        let mut extra = [0];
        if value.read(&mut extra).ok()? != 0 {
            return None;
        }
        bytes(&data)
    })();
    value.seek(SeekFrom::Start(original)).ok()?;
    detected
}

/// Detect a packaged iWork file or a legacy directory bundle.
///
/// Symbolic links, conflicting markers, malformed `Index.zip` archives, and
/// directory traversal errors fail closed.
pub fn path(value: impl AsRef<Path>) -> Option<Format> {
    let value = value.as_ref();
    match kind(value)? {
        Kind::File => reader(&mut File::open(value).ok()?),
        Kind::Dir => directory(value),
        Kind::Missing => None,
    }
}

fn directory(root: &Path) -> Option<Format> {
    let mut evidence = classify(
        marker(root, "index.xml")?,
        marker(root, "index.apxl")?,
        marker(root, "index.numbers")?,
    );
    if evidence == Outcome::Conflict {
        return None;
    }

    let index_zip = root.join("Index.zip");
    evidence = evidence.merge(match kind(&index_zip)? {
        Kind::File => {
            reader(&mut File::open(index_zip).ok()?).map_or(Outcome::Conflict, Outcome::Found)
        },
        Kind::Dir => Outcome::Conflict,
        Kind::Missing => Outcome::None,
    });
    if evidence == Outcome::Conflict {
        return None;
    }

    let index = root.join("Index");
    evidence = evidence.merge(match kind(&index)? {
        Kind::Dir => directory_outcome(&index)?,
        Kind::File => Outcome::Conflict,
        Kind::Missing => Outcome::None,
    });

    match evidence {
        Outcome::Found(format) => Some(format),
        Outcome::None | Outcome::Conflict => None,
    }
}

fn directory_outcome(index: &Path) -> Option<Outcome> {
    let mut marks = Marks::default();
    let mut document = None;
    for entry in fs::read_dir(index).ok()? {
        let entry = entry.ok()?;
        let kind = entry.file_type().ok()?;
        if kind.is_symlink() {
            return None;
        }
        if kind.is_file() {
            let name = entry.file_name();
            let name = name.to_str()?;
            marks.see_index(name);
            if name == "Document.iwa" {
                document = Some(entry.path());
            }
        }
    }
    if !marks.iwa {
        return Some(Outcome::Conflict);
    }
    let Some(document) = document else {
        return Some(Outcome::Conflict);
    };
    if fs::metadata(&document).ok()?.len() > ARCHIVE_LIMITS.max_entry_size {
        return Some(Outcome::Conflict);
    }
    let Some(format) = fs::read(document).ok().and_then(|data| root_format(&data)) else {
        return Some(Outcome::Conflict);
    };
    Some(if marks.accepts(format) {
        Outcome::Found(format)
    } else {
        Outcome::Conflict
    })
}

fn marks<'a>(names: impl IntoIterator<Item = &'a str>) -> Marks {
    let mut marks = Marks::default();
    for name in names {
        marks.see(name);
    }
    marks
}

#[derive(Debug, Default, Clone, Copy)]
struct Marks {
    iwa: bool,
    keynote: bool,
}

impl Marks {
    fn see(&mut self, name: &str) {
        let Some(name) = index_name(name) else {
            return;
        };
        self.see_index(name);
    }

    fn see_index(&mut self, name: &str) {
        if !name.ends_with(".iwa") {
            return;
        }
        self.iwa = true;
        self.keynote |= is_component(name, "MasterSlide")
            || is_component(name, "Slide")
            || is_component(name, "TemplateSlide");
    }

    fn accepts(self, format: Format) -> bool {
        !self.keynote || format == Format::Keynote
    }
}

fn index_name(name: &str) -> Option<&str> {
    name.strip_prefix("Index/")
        .or_else(|| (!name.contains('/')).then_some(name))
}

fn is_component(name: &str, stem: &str) -> bool {
    let Some(name) = name.strip_suffix(".iwa") else {
        return false;
    };
    let Some(suffix) = name.strip_prefix(stem) else {
        return false;
    };
    suffix.is_empty()
        || suffix.strip_prefix('-').is_some_and(|version| {
            !version.is_empty()
                && version
                    .split('-')
                    .all(|part| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()))
        })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Outcome {
    None,
    Found(Format),
    Conflict,
}

impl Outcome {
    fn merge(self, other: Self) -> Self {
        match (self, other) {
            (Self::Conflict, _) | (_, Self::Conflict) => Self::Conflict,
            (Self::None, outcome) | (outcome, Self::None) => outcome,
            (Self::Found(left), Self::Found(right)) if left == right => Self::Found(left),
            (Self::Found(_), Self::Found(_)) => Self::Conflict,
        }
    }
}

fn classify(pages: bool, keynote: bool, numbers: bool) -> Outcome {
    match usize::from(pages) + usize::from(keynote) + usize::from(numbers) {
        0 => Outcome::None,
        1 if pages => Outcome::Found(Format::Pages),
        1 if keynote => Outcome::Found(Format::Keynote),
        1 => Outcome::Found(Format::Numbers),
        _ => Outcome::Conflict,
    }
}

fn marker(root: &Path, name: &str) -> Option<bool> {
    match kind(&root.join(name))? {
        Kind::File => Some(true),
        Kind::Missing => Some(false),
        Kind::Dir => None,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Missing,
    File,
    Dir,
}

fn kind(value: &Path) -> Option<Kind> {
    match fs::symlink_metadata(value) {
        Ok(metadata) => {
            let kind = metadata.file_type();
            if kind.is_symlink() {
                None
            } else if kind.is_file() {
                Some(Kind::File)
            } else if kind.is_dir() {
                Some(Kind::Dir)
            } else {
                None
            }
        },
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Some(Kind::Missing),
        Err(_) => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{ArchiveObject, RawMessage};
    use crate::protobuf::{kn, tn, tp, tsa, tsk, tsp};
    use prost::Message;
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::io::Cursor;
    use std::sync::atomic::{AtomicU64, Ordering};

    fn package(names: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        for (name, data) in names {
            writer.write_stored(name, data).unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn document(format: Format) -> Vec<u8> {
        let shared_document = || tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        };
        let reference = |identifier| tsp::Reference {
            identifier,
            ..Default::default()
        };
        let (message_type, payload) = match format {
            Format::Pages => (
                10_000,
                tp::DocumentArchive {
                    super_: shared_document(),
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            Format::Keynote => (
                1,
                kn::DocumentArchive {
                    super_: shared_document(),
                    show: reference(1),
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
            Format::Numbers => (
                6,
                tn::DocumentArchive {
                    super_: shared_document(),
                    stylesheet: reference(1),
                    sidebar_order: reference(2),
                    theme: reference(3),
                    ..Default::default()
                }
                .encode_to_vec(),
            ),
        };
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: message_type,
                        data: payload,
                    }],
                )
                .unwrap(),
            ],
        };
        SnappyStream::compress(&archive.to_bytes().unwrap()).unwrap()
    }

    fn document_package(format: Format, extra_names: &[&str]) -> Vec<u8> {
        let root = document(format);
        let mut files = vec![("Index/Document.iwa", root.as_slice())];
        files.extend(extra_names.iter().map(|name| (*name, b"iwa".as_slice())));
        package(&files)
    }

    #[test]
    fn detects_root_application_with_shared_table_components() {
        assert_eq!(
            bytes(&document_package(Format::Pages, &[])),
            Some(Format::Pages)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Keynote,
                &[
                    "Index/MasterSlide-12.iwa",
                    "Index/Slide-1.iwa",
                    "Index/TemplateSlide-31.iwa",
                    "Index/CalculationEngine-81.iwa"
                ]
            )),
            Some(Format::Keynote)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Numbers,
                &["Index/CalculationEngine-174.iwa"]
            )),
            Some(Format::Numbers)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Pages,
                &["Index/CalculationEngine.iwa"]
            )),
            Some(Format::Pages)
        );
        assert_eq!(
            bytes(&document_package(Format::Numbers, &["Index/Slide-1.iwa"])),
            None
        );
        assert_eq!(
            bytes(&document_package(
                Format::Pages,
                &["Index/MasterSlide-12.iwa"]
            )),
            None
        );
        assert_eq!(bytes(&package(&[("Index/Document.iwa", b"not iwa")])), None);
        assert_eq!(bytes(&package(&[("Index/Unknown.iwa", b"iwa")])), None);
        assert_eq!(bytes(&package(&[("Data/image.png", b"iwa")])), None);
    }

    #[test]
    fn detects_nested_legacy_indexes_and_rejects_ambiguous_or_encrypted_packages() {
        let root = document(Format::Pages);
        let index = package(&[("Document.iwa", &root)]);
        let outer = package(&[("legacy.pages/Index.zip", &index)]);
        assert_eq!(bytes(&outer), Some(Format::Pages));

        let ambiguous = package(&[("a/Index.zip", &index), ("b/Index.zip", &index)]);
        assert_eq!(bytes(&ambiguous), None);

        let root = document(Format::Pages);
        let encrypted = package(&[
            ("Index/Document.iwa", &root),
            ("Metadata/.iwpv2", b"encryption metadata"),
        ]);
        assert_eq!(bytes(&encrypted), None);
    }

    #[test]
    fn reader_restores_nonzero_cursor_on_success_and_rejection() {
        let mut valid = Cursor::new(document_package(Format::Pages, &[]));
        valid.set_position(9);
        assert_eq!(reader(&mut valid), Some(Format::Pages));
        assert_eq!(valid.position(), 9);

        let mut invalid = Cursor::new(b"not an iWork package".to_vec());
        invalid.set_position(4);
        assert_eq!(reader(&mut invalid), None);
        assert_eq!(invalid.position(), 4);
    }

    #[test]
    fn path_supports_files_legacy_bundles_and_index_zip() -> std::io::Result<()> {
        let temp = Temp::new()?;

        let packaged = temp.0.join("document.pages");
        fs::write(&packaged, document_package(Format::Pages, &[]))?;
        assert_eq!(path(&packaged), Some(Format::Pages));

        let legacy = temp.0.join("legacy.key");
        fs::create_dir(&legacy)?;
        fs::write(legacy.join("index.apxl"), [])?;
        assert_eq!(path(&legacy), Some(Format::Keynote));

        let bundle = temp.0.join("sheet.numbers");
        fs::create_dir(&bundle)?;
        fs::write(
            bundle.join("Index.zip"),
            document_package(Format::Numbers, &["Index/CalculationEngine-174.iwa"]),
        )?;
        assert_eq!(path(&bundle), Some(Format::Numbers));

        let agreeing = temp.0.join("agreeing.pages");
        fs::create_dir(&agreeing)?;
        fs::write(agreeing.join("index.xml"), [])?;
        fs::write(
            agreeing.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        assert_eq!(path(&agreeing), Some(Format::Pages));

        let unpacked = temp.0.join("unpacked.key");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(
            unpacked.join("Index/Document.iwa"),
            document(Format::Keynote),
        )?;
        fs::write(unpacked.join("Index/Slide-1.iwa"), [])?;
        assert_eq!(path(&unpacked), Some(Format::Keynote));
        Ok(())
    }

    #[test]
    fn path_rejects_generic_and_conflicting_index_directories() -> std::io::Result<()> {
        let temp = Temp::new()?;

        let generic = temp.0.join("generic");
        fs::create_dir_all(generic.join("Index"))?;
        fs::write(generic.join("Index/Unknown.iwa"), [])?;
        assert_eq!(path(&generic), None);

        let media = temp.0.join("media-only");
        fs::create_dir_all(media.join("Data"))?;
        fs::create_dir(media.join("Assets"))?;
        fs::write(media.join("theme-preview.jpg"), [])?;
        assert_eq!(path(&media), None);

        let conflict = temp.0.join("conflict");
        fs::create_dir_all(conflict.join("Index"))?;
        fs::write(conflict.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(conflict.join("Index/Slide.iwa"), [])?;
        fs::write(conflict.join("Index/CalculationEngine.iwa"), [])?;
        assert_eq!(path(&conflict), None);

        let legacy_conflict = temp.0.join("legacy-conflict");
        fs::create_dir(&legacy_conflict)?;
        fs::write(legacy_conflict.join("index.xml"), [])?;
        fs::write(
            legacy_conflict.join("Index.zip"),
            document_package(Format::Numbers, &[]),
        )?;
        assert_eq!(path(&legacy_conflict), None);

        let representation_conflict = temp.0.join("representation-conflict");
        fs::create_dir_all(representation_conflict.join("Index"))?;
        fs::write(
            representation_conflict.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        fs::write(
            representation_conflict.join("Index/Document.iwa"),
            document(Format::Keynote),
        )?;
        fs::write(representation_conflict.join("Index/Slide-1.iwa"), [])?;
        assert_eq!(path(&representation_conflict), None);
        Ok(())
    }

    struct Temp(std::path::PathBuf);

    impl Temp {
        fn new() -> std::io::Result<Self> {
            static NEXT: AtomicU64 = AtomicU64::new(0);
            let path = std::env::temp_dir().join(format!(
                "litchi-iwa-detect-{}-{}",
                std::process::id(),
                NEXT.fetch_add(1, Ordering::Relaxed)
            ));
            fs::create_dir(&path)?;
            Ok(Self(path))
        }
    }

    impl Drop for Temp {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }
}
