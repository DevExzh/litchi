//! Typed Apple iWork format detection.
//!
//! Detection validates application-root evidence plus legacy bundle markers.
//! It inspects only the root archive envelope, not document content, and does
//! not expose the ZIP implementation.

use crate::application::Application;
use crate::archive::Archive;
use crate::snappy::SnappyStream;
use litchi_iwa_archive::{self, DetectionRoot};
use litchi_iwa_common::wire::{WireField, parse_wire_fields};
use std::fs::{self, File};
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

const MAX_INPUT_BYTES: u64 = 1024 * 1024 * 1024;
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

/// Resource ceilings for one iWork detection attempt.
///
/// The defaults are conservative enough for untrusted input while allowing
/// ordinary media-heavy documents. Callers may tighten any ceiling, but the
/// checked constructor never permits a limit above the format-wide hard
/// ceiling. Detection remains fail-closed when a limit is exceeded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: u64,
    max_files: usize,
    max_entry_size: u64,
    max_total_size: u64,
    max_iwa_stream_size: usize,
}

impl Limits {
    /// Maximum input size accepted by the default safety profile.
    pub const HARD_MAX_INPUT_BYTES: u64 = MAX_INPUT_BYTES;
    /// Maximum ZIP member count accepted by the default safety profile.
    pub const HARD_MAX_FILES: usize = litchi_iwa_archive::Limits::MAX_ENTRIES;
    /// Maximum uncompressed ZIP member size accepted by the default profile.
    pub const HARD_MAX_ENTRY_SIZE: u64 = litchi_iwa_archive::Limits::MAX_ENTRY_BYTES;
    /// Maximum aggregate uncompressed ZIP size accepted by the default profile.
    pub const HARD_MAX_TOTAL_SIZE: u64 = litchi_iwa_archive::Limits::MAX_TOTAL_BYTES;
    /// Maximum decompressed size of one IWA component.
    pub const HARD_MAX_IWA_STREAM_SIZE: usize = SnappyStream::MAX_DECOMPRESSED_STREAM;

    /// Construct checked detection ceilings.
    pub fn new(
        max_input_bytes: u64,
        max_files: usize,
        max_entry_size: u64,
        max_total_size: u64,
        max_iwa_stream_size: usize,
    ) -> crate::Result<Self> {
        if max_input_bytes == 0
            || max_files == 0
            || max_entry_size == 0
            || max_total_size == 0
            || max_iwa_stream_size == 0
        {
            return Err(crate::Error::InvalidFormat(
                "iWork detection limits must be non-zero".to_owned(),
            ));
        }
        if max_input_bytes > Self::HARD_MAX_INPUT_BYTES {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork detection input limit exceeds {} bytes",
                Self::HARD_MAX_INPUT_BYTES
            )));
        }
        if max_files > Self::HARD_MAX_FILES {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork detection file limit exceeds {} entries",
                Self::HARD_MAX_FILES
            )));
        }
        if max_entry_size > Self::HARD_MAX_ENTRY_SIZE {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork detection entry limit exceeds {} bytes",
                Self::HARD_MAX_ENTRY_SIZE
            )));
        }
        if max_total_size > Self::HARD_MAX_TOTAL_SIZE {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork detection total-size limit exceeds {} bytes",
                Self::HARD_MAX_TOTAL_SIZE
            )));
        }
        if max_iwa_stream_size > Self::HARD_MAX_IWA_STREAM_SIZE {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork detection IWA limit exceeds {} bytes",
                Self::HARD_MAX_IWA_STREAM_SIZE
            )));
        }

        Ok(Self {
            max_input_bytes,
            max_files,
            max_entry_size,
            max_total_size,
            max_iwa_stream_size,
        })
    }

    /// Maximum complete input size accepted by this profile.
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }

    /// Maximum number of ZIP members indexed by one probe.
    pub const fn max_files(self) -> usize {
        self.max_files
    }

    /// Maximum declared uncompressed size of one ZIP member.
    pub const fn max_entry_size(self) -> u64 {
        self.max_entry_size
    }

    /// Maximum aggregate declared uncompressed ZIP size.
    pub const fn max_total_size(self) -> u64 {
        self.max_total_size
    }

    /// Maximum decompressed size of one IWA component.
    pub const fn max_iwa_stream_size(self) -> usize {
        self.max_iwa_stream_size
    }

    fn snappy_limits(self) -> crate::Result<crate::snappy::SnappyLimits> {
        Ok(crate::snappy::SnappyLimits::new(
            self.max_iwa_stream_size
                .min(SnappyStream::MAX_UNCOMPRESSED_CHUNK),
            self.max_iwa_stream_size,
        )?)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_input_bytes: MAX_INPUT_BYTES,
            max_files: litchi_iwa_archive::Limits::MAX_ENTRIES,
            max_entry_size: litchi_iwa_archive::Limits::MAX_ENTRY_BYTES,
            max_total_size: litchi_iwa_archive::Limits::MAX_TOTAL_BYTES,
            max_iwa_stream_size: SnappyStream::MAX_DECOMPRESSED_STREAM,
        }
    }
}

/// Detect an iWork application from complete packaged bytes.
///
/// ZIP and Snappy metadata are validated under explicit file-count and size
/// limits. A package with conflicting application-root evidence is reported as
/// a typed format error; an unrelated or unrecognized byte slice returns
/// `Ok(None)`.
pub fn bytes(value: &[u8]) -> crate::Result<Option<Format>> {
    bytes_with_limits(value, Limits::default())
}

/// Detect an iWork application using caller-selected resource ceilings.
pub fn bytes_with_limits(value: &[u8], limits: Limits) -> crate::Result<Option<Format>> {
    let input_size = u64::try_from(value.len())
        .map_err(|_| crate::Error::InvalidFormat("iWork input length exceeds u64".to_owned()))?;
    if input_size > limits.max_input_bytes {
        return Err(crate::Error::InvalidFormat(format!(
            "iWork detection input is {input_size} bytes, exceeding the {} byte limit",
            limits.max_input_bytes
        )));
    }
    if !is_zip_signature(value) {
        return Ok(None);
    }
    let root = litchi_iwa_archive::inspect_detection_root(value, archive_limits(limits)?)
        .map_err(map_archive_error)?;
    classify_root(&root, limits)
}

fn archive_limits(limits: Limits) -> crate::Result<litchi_iwa_archive::Limits> {
    litchi_iwa_archive::Limits::new(
        limits.max_input_bytes,
        limits.max_files,
        limits.max_entry_size,
        limits.max_total_size,
        limits.max_iwa_stream_size,
    )
    .map_err(map_archive_error)
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> crate::Error {
    match error {
        litchi_iwa_archive::Error::Io(error) => crate::Error::Io(error),
        litchi_iwa_archive::Error::Iwa(error) => crate::Error::from(error),
        litchi_iwa_archive::Error::Encrypted => crate::Error::InvalidFormat(
            "password-protected iWork documents are not supported".to_owned(),
        ),
        litchi_iwa_archive::Error::InvalidLimits(message) => crate::Error::InvalidFormat(message),
        litchi_iwa_archive::Error::Zip { message }
        | litchi_iwa_archive::Error::InvalidBundle(message) => {
            crate::Error::Archive(format!("iWork archive ingress: {message}"))
        },
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => crate::Error::InvalidFormat(format!(
            "iWork archive {kind} limit exceeded: observed {observed}, maximum {maximum}"
        )),
        litchi_iwa_archive::Error::Allocation { resource, amount } => {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
        },
    }
}

fn classify_root(root: &DetectionRoot, limits: Limits) -> crate::Result<Option<Format>> {
    if !root.has_iwa_components() {
        return Ok(None);
    }
    let marks = Marks {
        iwa: true,
        keynote: root.has_keynote_components(),
    };
    let Some(document) = root.document() else {
        return Ok(None);
    };
    let Some(format) = root_format_archive(document, limits)? else {
        return Err(crate::Error::InvalidFormat(
            "Document.iwa has no recognized iWork application root".to_owned(),
        ));
    };
    if marks.accepts(format) {
        Ok(Some(format))
    } else {
        Err(crate::Error::InvalidFormat(
            "iWork component markers conflict with the Document.iwa application root".to_owned(),
        ))
    }
}

fn is_zip_signature(value: &[u8]) -> bool {
    value.starts_with(b"PK\x03\x04")
        || value.starts_with(b"PK\x05\x06")
        || value.starts_with(b"PK\x07\x08")
}

/// Detect the owning iWork application from the root `DocumentArchive` payload.
///
/// Message type identifiers overlap between Pages, Numbers, and Keynote, so they
/// cannot reliably identify an application. The root protobuf schemas have
/// stable, application-specific required message shapes: Pages uses its shared
/// document at field 15, Numbers uses references at fields 4/5/6 plus its shared
/// document at field 8, and Keynote uses a reference at field 2 plus its shared
/// document at field 3. Malformed or multiply matching payloads fail closed.
pub(crate) fn detect_application_from_document(payload: &[u8]) -> Option<Application> {
    let fields = wire_fields(payload)?;
    let pages = unique_field(payload, &fields, 15, 2).is_some_and(valid_shared_document);
    let numbers = [4, 5, 6]
        .into_iter()
        .all(|field| unique_field(payload, &fields, field, 2).is_some_and(valid_reference))
        && unique_field(payload, &fields, 8, 2).is_some_and(valid_shared_document);
    let keynote = unique_field(payload, &fields, 2, 2).is_some_and(valid_reference)
        && unique_field(payload, &fields, 3, 2).is_some_and(valid_shared_document);

    match (pages, numbers, keynote) {
        (true, false, false) => Some(Application::Pages),
        (false, true, false) => Some(Application::Numbers),
        (false, false, true) => Some(Application::Keynote),
        _ => None,
    }
}

fn wire_fields(payload: &[u8]) -> Option<Vec<WireField>> {
    parse_wire_fields(payload).ok()
}

fn unique_field<'a>(
    payload: &'a [u8],
    fields: &[WireField],
    number: u32,
    wire_type: u8,
) -> Option<&'a [u8]> {
    let mut matches = fields.iter().filter(|field| field.number() == number);
    let field = matches.next()?;
    if matches.next().is_some() || field.wire_type() != wire_type {
        return None;
    }
    field.checked_payload(payload).ok()
}

fn valid_reference(payload: &[u8]) -> bool {
    wire_fields(payload)
        .and_then(|fields| unique_field(payload, &fields, 1, 0))
        .is_some()
}

fn valid_shared_document(payload: &[u8]) -> bool {
    wire_fields(payload)
        .and_then(|fields| unique_field(payload, &fields, 1, 2))
        .and_then(wire_fields)
        .is_some()
}

fn root_format(data: &[u8], limits: Limits) -> crate::Result<Option<Format>> {
    let stream = SnappyStream::decompress_with_limits(data, limits.snappy_limits()?)?;
    let archive = Archive::parse(stream.as_bytes())?;
    root_format_archive(&archive, limits)
}

fn root_format_archive(archive: &Archive, _limits: Limits) -> crate::Result<Option<Format>> {
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
        if detected.is_some() {
            return Err(crate::Error::InvalidFormat(
                "Document.iwa contains multiple application roots".to_owned(),
            ));
        }
        detected = Some(format);
    }

    Ok(detected)
}

/// Detect an iWork application from a seekable stream.
///
/// Detection starts at byte zero and restores the caller's original cursor on
/// every path. Streams larger than the selected input ceiling are rejected
/// without being read.
pub fn reader<R: Read + Seek>(value: &mut R) -> crate::Result<Option<Format>> {
    reader_with_limits(value, Limits::default())
}

/// Detect an iWork application from a seekable stream under explicit limits.
pub fn reader_with_limits<R: Read + Seek>(
    value: &mut R,
    limits: Limits,
) -> crate::Result<Option<Format>> {
    let original = value.stream_position()?;
    let detected = (|| {
        let length = value.seek(SeekFrom::End(0))?;
        if length > limits.max_input_bytes {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork detection input is {length} bytes, exceeding the {} byte limit",
                limits.max_input_bytes
            )));
        }

        let length = usize::try_from(length).map_err(|_| {
            crate::Error::InvalidFormat("iWork input length does not fit usize".to_owned())
        })?;
        let mut data = Vec::new();
        data.try_reserve_exact(length).map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "unable to reserve iWork detection input buffer: {error}"
            ))
        })?;
        data.resize(length, 0);

        value.seek(SeekFrom::Start(0))?;
        value.read_exact(&mut data)?;
        let mut extra = [0];
        if value.read(&mut extra)? != 0 {
            return Err(crate::Error::InvalidFormat(
                "iWork detection source changed while it was being read".to_owned(),
            ));
        }
        bytes_with_limits(&data, limits)
    })();
    value.seek(SeekFrom::Start(original))?;
    detected
}

/// Detect a packaged iWork file or a legacy directory bundle.
///
/// Symbolic links, conflicting markers, malformed `Index.zip` archives, and
/// directory traversal errors are typed errors.
pub fn path(value: impl AsRef<Path>) -> crate::Result<Option<Format>> {
    path_with_limits(value, Limits::default())
}

/// Detect a packaged file or legacy directory bundle under explicit limits.
pub fn path_with_limits(value: impl AsRef<Path>, limits: Limits) -> crate::Result<Option<Format>> {
    let value = value.as_ref();
    match kind(value)? {
        Kind::File => reader_with_limits(&mut File::open(value)?, limits),
        Kind::Dir => directory(value, limits),
        Kind::Missing => Ok(None),
    }
}

fn directory(root: &Path, limits: Limits) -> crate::Result<Option<Format>> {
    let mut evidence = classify(
        marker(root, "index.xml")?,
        marker(root, "index.apxl")?,
        marker(root, "index.numbers")?,
    );
    if evidence == Outcome::Conflict {
        return Err(crate::Error::InvalidFormat(
            "iWork bundle contains conflicting legacy application markers".to_owned(),
        ));
    }

    let index_zip = root.join("Index.zip");
    evidence = evidence.merge(match kind(&index_zip)? {
        Kind::File => match reader_with_limits(&mut File::open(&index_zip)?, limits)? {
            Some(format) => Outcome::Found(format),
            None => Outcome::Conflict,
        },
        Kind::Dir => Outcome::Conflict,
        Kind::Missing => Outcome::None,
    });
    if evidence == Outcome::Conflict {
        return Err(crate::Error::InvalidFormat(
            "iWork bundle contains an invalid or conflicting Index.zip".to_owned(),
        ));
    }

    let index = root.join("Index");
    evidence = evidence.merge(match kind(&index)? {
        Kind::Dir => directory_outcome(&index, limits)?,
        Kind::File => Outcome::Conflict,
        Kind::Missing => Outcome::None,
    });

    match evidence {
        Outcome::Found(format) => Ok(Some(format)),
        Outcome::None => Ok(None),
        Outcome::Conflict => Err(crate::Error::InvalidFormat(
            "iWork bundle contains conflicting application evidence".to_owned(),
        )),
    }
}

fn directory_outcome(index: &Path, limits: Limits) -> crate::Result<Outcome> {
    let mut marks = Marks::default();
    let mut document = None;
    let mut entry_count = 0usize;
    let mut total_size = 0u64;
    for entry in fs::read_dir(index)? {
        let entry = entry?;
        entry_count = entry_count.checked_add(1).ok_or_else(|| {
            crate::Error::InvalidFormat("iWork index entry count overflow".to_owned())
        })?;
        if entry_count > limits.max_files {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork index directory contains more than the {} entry limit",
                limits.max_files
            )));
        }
        let kind = entry.file_type()?;
        if kind.is_symlink() {
            return Err(crate::Error::InvalidFormat(format!(
                "iWork bundle index contains symbolic link {}",
                entry.path().display()
            )));
        }
        if kind.is_file() {
            let entry_size = entry.metadata()?.len();
            total_size = total_size.checked_add(entry_size).ok_or_else(|| {
                crate::Error::InvalidFormat("iWork index byte count overflow".to_owned())
            })?;
            if total_size > limits.max_total_size {
                return Err(crate::Error::InvalidFormat(format!(
                    "iWork index directory contains {total_size} bytes, exceeding the {} byte limit",
                    limits.max_total_size
                )));
            }
            let name = entry.file_name();
            let name = name.to_str().ok_or_else(|| {
                crate::Error::InvalidFormat(format!(
                    "iWork bundle index contains a non-UTF-8 entry: {}",
                    entry.path().display()
                ))
            })?;
            marks.see_index(name);
            if name == "Document.iwa" {
                document = Some(entry.path());
            }
        }
    }
    if !marks.iwa {
        return Ok(Outcome::None);
    }
    let Some(document) = document else {
        return Err(crate::Error::InvalidFormat(
            "iWork bundle index contains IWA components but no Document.iwa".to_owned(),
        ));
    };
    let document_size = fs::metadata(&document)?.len();
    if document_size > limits.max_input_bytes || document_size > limits.max_entry_size {
        let limit = limits.max_input_bytes.min(limits.max_entry_size);
        return Err(crate::Error::InvalidFormat(format!(
            "iWork Document.iwa is {document_size} bytes, exceeding the {limit} byte limit"
        )));
    }
    let Some(format) = root_format(&fs::read(&document)?, limits)? else {
        return Err(crate::Error::InvalidFormat(
            "Document.iwa has no recognized iWork application root".to_owned(),
        ));
    };
    Ok(if marks.accepts(format) {
        Outcome::Found(format)
    } else {
        Outcome::Conflict
    })
}

#[derive(Debug, Default, Clone, Copy)]
struct Marks {
    iwa: bool,
    keynote: bool,
}

impl Marks {
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

fn marker(root: &Path, name: &str) -> crate::Result<bool> {
    match kind(&root.join(name))? {
        Kind::File => Ok(true),
        Kind::Missing => Ok(false),
        Kind::Dir => Err(crate::Error::InvalidFormat(format!(
            "iWork marker {} is a directory",
            root.join(name).display()
        ))),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Missing,
    File,
    Dir,
}

fn kind(value: &Path) -> crate::Result<Kind> {
    match fs::symlink_metadata(value) {
        Ok(metadata) => {
            let kind = metadata.file_type();
            if kind.is_symlink() {
                Err(crate::Error::InvalidFormat(format!(
                    "iWork detection refuses symbolic link {}",
                    value.display()
                )))
            } else if kind.is_file() {
                Ok(Kind::File)
            } else if kind.is_dir() {
                Ok(Kind::Dir)
            } else {
                Err(crate::Error::InvalidFormat(format!(
                    "iWork detection refuses unsupported filesystem node {}",
                    value.display()
                )))
            }
        },
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(Kind::Missing),
        Err(error) => Err(crate::Error::Io(error)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{ArchiveObject, RawMessage};
    use crate::protobuf::{kn, tn, tp, tsa, tsk, tsp};
    use prost::Message;
    use std::io::Cursor;
    use std::sync::atomic::{AtomicU64, Ordering};

    fn shared_document() -> tsa::DocumentArchive {
        tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        }
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn document_payload(application: Application) -> Vec<u8> {
        match application {
            Application::Pages => tp::DocumentArchive {
                super_: shared_document(),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Numbers => tn::DocumentArchive {
                super_: shared_document(),
                stylesheet: reference(1),
                sidebar_order: reference(2),
                theme: reference(3),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Keynote => kn::DocumentArchive {
                super_: shared_document(),
                show: reference(1),
                ..Default::default()
            }
            .encode_to_vec(),
            Application::Common => Vec::new(),
        }
    }

    fn package(names: &[(&str, &[u8])]) -> Vec<u8> {
        litchi_iwa_archive::package::to_bytes(
            names.iter().copied(),
            litchi_iwa_archive::Limits::default(),
        )
        .unwrap()
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
    fn test_document_payload_detection() {
        assert_eq!(
            detect_application_from_document(&document_payload(Application::Pages)),
            Some(Application::Pages)
        );
        assert_eq!(
            detect_application_from_document(&document_payload(Application::Numbers)),
            Some(Application::Numbers)
        );
        assert_eq!(
            detect_application_from_document(&document_payload(Application::Keynote)),
            Some(Application::Keynote)
        );

        let pages_with_references = tp::DocumentArchive {
            super_: shared_document(),
            stylesheet: Some(reference(1)),
            floating_drawables: Some(reference(2)),
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(
            detect_application_from_document(&pages_with_references),
            Some(Application::Pages)
        );

        let mut conflicting = document_payload(Application::Pages);
        conflicting.extend(document_payload(Application::Numbers));
        assert_eq!(detect_application_from_document(&conflicting), None);

        let mut conflicting = document_payload(Application::Pages);
        conflicting.extend(document_payload(Application::Keynote));
        assert_eq!(detect_application_from_document(&conflicting), None);

        assert_eq!(detect_application_from_document(&[0x78, 0x00]), None);
        assert_eq!(detect_application_from_document(&[0x7a, 0x00]), None);
        assert_eq!(detect_application_from_document(&[0x80]), None);
    }

    #[test]
    fn detects_root_application_with_shared_table_components() {
        assert_eq!(
            bytes(&document_package(Format::Pages, &[])).unwrap(),
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
            ))
            .unwrap(),
            Some(Format::Keynote)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Numbers,
                &["Index/CalculationEngine-174.iwa"]
            ))
            .unwrap(),
            Some(Format::Numbers)
        );
        assert_eq!(
            bytes(&document_package(
                Format::Pages,
                &["Index/CalculationEngine.iwa"]
            ))
            .unwrap(),
            Some(Format::Pages)
        );
        assert!(bytes(&document_package(Format::Numbers, &["Index/Slide-1.iwa"])).is_err());
        assert!(
            bytes(&document_package(
                Format::Pages,
                &["Index/MasterSlide-12.iwa"]
            ))
            .is_err()
        );
        assert!(bytes(&package(&[("Index/Document.iwa", b"not iwa")])).is_err());
        assert_eq!(
            bytes(&package(&[("Index/Unknown.iwa", b"iwa")])).unwrap(),
            None
        );
        assert_eq!(
            bytes(&package(&[("Data/image.png", b"iwa")])).unwrap(),
            None
        );

        let duplicate_root = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: document_payload(Application::Pages),
                    }],
                )
                .unwrap(),
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 10_000,
                        data: document_payload(Application::Pages),
                    }],
                )
                .unwrap(),
            ],
        };
        assert!(root_format_archive(&duplicate_root, Limits::default()).is_err());
    }

    #[test]
    fn detects_nested_legacy_indexes_and_rejects_ambiguous_or_encrypted_packages() {
        let root = document(Format::Pages);
        let index = package(&[("Document.iwa", &root)]);
        let outer = package(&[("legacy.pages/Index.zip", &index)]);
        assert_eq!(bytes(&outer).unwrap(), Some(Format::Pages));

        let mixed = package(&[
            ("legacy.pages/Index.zip", &index),
            ("Index/CalculationEngine.iwa", b"iwa"),
        ]);
        assert!(bytes(&mixed).is_err());

        let ambiguous = package(&[("a/Index.zip", &index), ("b/Index.zip", &index)]);
        assert!(bytes(&ambiguous).is_err());

        let root = document(Format::Pages);
        let encrypted = package(&[
            ("Index/Document.iwa", &root),
            ("Metadata/.iwpv2", b"encryption metadata"),
        ]);
        assert!(bytes(&encrypted).is_err());
    }

    #[test]
    fn checked_limits_preserve_defaults_and_bound_each_layer() {
        let valid = document_package(Format::Pages, &[]);
        let defaults = Limits::default();
        assert_eq!(
            bytes_with_limits(&valid, defaults).unwrap(),
            Some(Format::Pages)
        );
        assert_eq!(defaults.max_input_bytes(), Limits::HARD_MAX_INPUT_BYTES);
        assert_eq!(defaults.max_files(), Limits::HARD_MAX_FILES);
        assert_eq!(defaults.max_entry_size(), Limits::HARD_MAX_ENTRY_SIZE);
        assert_eq!(defaults.max_total_size(), Limits::HARD_MAX_TOTAL_SIZE);
        assert_eq!(
            defaults.max_iwa_stream_size(),
            Limits::HARD_MAX_IWA_STREAM_SIZE
        );

        let input_bound = Limits::new(1, 1, 1, 1, 1).unwrap();
        assert!(bytes_with_limits(&valid, input_bound).is_err());

        let stream_bound = Limits::new(
            Limits::HARD_MAX_INPUT_BYTES,
            Limits::HARD_MAX_FILES,
            Limits::HARD_MAX_ENTRY_SIZE,
            Limits::HARD_MAX_TOTAL_SIZE,
            1,
        )
        .unwrap();
        assert!(bytes_with_limits(&valid, stream_bound).is_err());
    }

    #[test]
    fn checked_limits_reject_zero_and_hard_ceiling_escapes() {
        assert!(Limits::new(0, 1, 1, 1, 1).is_err());
        assert!(Limits::new(1, 0, 1, 1, 1).is_err());
        assert!(Limits::new(1, 1, 0, 1, 1).is_err());
        assert!(Limits::new(1, 1, 1, 0, 1).is_err());
        assert!(Limits::new(1, 1, 1, 1, 0).is_err());
        assert!(Limits::new(Limits::HARD_MAX_INPUT_BYTES + 1, 1, 1, 1, 1).is_err());
        assert!(Limits::new(1, Limits::HARD_MAX_FILES + 1, 1, 1, 1).is_err());
        assert!(Limits::new(1, 1, Limits::HARD_MAX_ENTRY_SIZE + 1, 1, 1).is_err());
        assert!(Limits::new(1, 1, 1, Limits::HARD_MAX_TOTAL_SIZE + 1, 1).is_err());
        assert!(Limits::new(1, 1, 1, 1, Limits::HARD_MAX_IWA_STREAM_SIZE + 1).is_err());
    }

    #[test]
    fn reader_restores_nonzero_cursor_on_success_and_rejection() {
        let mut valid = Cursor::new(document_package(Format::Pages, &[]));
        valid.set_position(9);
        assert_eq!(reader(&mut valid).unwrap(), Some(Format::Pages));
        assert_eq!(valid.position(), 9);

        let mut invalid = Cursor::new(b"not an iWork package".to_vec());
        invalid.set_position(4);
        assert_eq!(reader(&mut invalid).unwrap(), None);
        assert_eq!(invalid.position(), 4);
    }

    #[test]
    fn path_supports_files_legacy_bundles_and_index_zip() -> std::io::Result<()> {
        let temp = Temp::new()?;

        let packaged = temp.0.join("document.pages");
        fs::write(&packaged, document_package(Format::Pages, &[]))?;
        assert_eq!(path(&packaged).unwrap(), Some(Format::Pages));

        let legacy = temp.0.join("legacy.key");
        fs::create_dir(&legacy)?;
        fs::write(legacy.join("index.apxl"), [])?;
        assert_eq!(path(&legacy).unwrap(), Some(Format::Keynote));

        let bundle = temp.0.join("sheet.numbers");
        fs::create_dir(&bundle)?;
        fs::write(
            bundle.join("Index.zip"),
            document_package(Format::Numbers, &["Index/CalculationEngine-174.iwa"]),
        )?;
        assert_eq!(path(&bundle).unwrap(), Some(Format::Numbers));

        let agreeing = temp.0.join("agreeing.pages");
        fs::create_dir(&agreeing)?;
        fs::write(agreeing.join("index.xml"), [])?;
        fs::write(
            agreeing.join("Index.zip"),
            document_package(Format::Pages, &[]),
        )?;
        assert_eq!(path(&agreeing).unwrap(), Some(Format::Pages));

        let unpacked = temp.0.join("unpacked.key");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(
            unpacked.join("Index/Document.iwa"),
            document(Format::Keynote),
        )?;
        fs::write(unpacked.join("Index/Slide-1.iwa"), [])?;
        assert_eq!(path(&unpacked).unwrap(), Some(Format::Keynote));

        let tight = Limits::new(1, 1, 1, 1, 1).unwrap();
        assert!(path_with_limits(&packaged, tight).is_err());
        assert!(path_with_limits(&unpacked, tight).is_err());
        Ok(())
    }

    #[test]
    fn unpacked_index_entry_count_is_bounded_before_document_read() -> std::io::Result<()> {
        let temp = Temp::new()?;
        let unpacked = temp.0.join("bounded.pages");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(unpacked.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(unpacked.join("Index/Extra.iwa"), [])?;

        let limits = Limits::new(
            Limits::HARD_MAX_INPUT_BYTES,
            1,
            Limits::HARD_MAX_ENTRY_SIZE,
            Limits::HARD_MAX_TOTAL_SIZE,
            Limits::HARD_MAX_IWA_STREAM_SIZE,
        )
        .unwrap();
        let error = path_with_limits(&unpacked, limits).unwrap_err();
        assert!(error.to_string().contains("more than the 1 entry limit"));
        Ok(())
    }

    #[test]
    fn unpacked_index_total_size_is_bounded() -> std::io::Result<()> {
        let temp = Temp::new()?;
        let unpacked = temp.0.join("total-size.pages");
        fs::create_dir_all(unpacked.join("Index"))?;
        fs::write(unpacked.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(unpacked.join("Index/Extra.iwa"), b"extra bytes")?;
        let total = fs::metadata(unpacked.join("Index/Document.iwa"))?.len()
            + fs::metadata(unpacked.join("Index/Extra.iwa"))?.len();
        let limits = Limits::new(
            Limits::HARD_MAX_INPUT_BYTES,
            Limits::HARD_MAX_FILES,
            Limits::HARD_MAX_ENTRY_SIZE,
            total - 1,
            Limits::HARD_MAX_IWA_STREAM_SIZE,
        )
        .unwrap();
        let error = path_with_limits(&unpacked, limits).unwrap_err();
        assert!(error.to_string().contains("iWork index directory contains"));
        assert!(error.to_string().contains("byte limit"));
        Ok(())
    }

    #[test]
    fn path_rejects_generic_and_conflicting_index_directories() -> std::io::Result<()> {
        let temp = Temp::new()?;

        let generic = temp.0.join("generic");
        fs::create_dir_all(generic.join("Index"))?;
        fs::write(generic.join("Index/Unknown.iwa"), [])?;
        assert!(path(&generic).is_err());

        let media = temp.0.join("media-only");
        fs::create_dir_all(media.join("Data"))?;
        fs::create_dir(media.join("Assets"))?;
        fs::write(media.join("theme-preview.jpg"), [])?;
        assert_eq!(path(&media).unwrap(), None);

        let conflict = temp.0.join("conflict");
        fs::create_dir_all(conflict.join("Index"))?;
        fs::write(conflict.join("Index/Document.iwa"), document(Format::Pages))?;
        fs::write(conflict.join("Index/Slide.iwa"), [])?;
        fs::write(conflict.join("Index/CalculationEngine.iwa"), [])?;
        assert!(path(&conflict).is_err());

        let legacy_conflict = temp.0.join("legacy-conflict");
        fs::create_dir(&legacy_conflict)?;
        fs::write(legacy_conflict.join("index.xml"), [])?;
        fs::write(
            legacy_conflict.join("Index.zip"),
            document_package(Format::Numbers, &[]),
        )?;
        assert!(path(&legacy_conflict).is_err());

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
        assert!(path(&representation_conflict).is_err());
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
