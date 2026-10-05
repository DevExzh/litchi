//! Handle for PowerPoint binary (`.ppt`) inputs (record 0003 §2.3, M0a step
//! S5d).
//!
//! The facade reaches the PPT reader by two routes, and the handle has one
//! crate-internal constructor for each:
//!
//! - the filesystem route, where detection opened a positional
//!   [`SourceBackedPackage`] over the file (`Ppt::from_source_package`);
//! - the byte route, where detection parsed the input into an in-memory
//!   [`OleFile`] (`Ppt::from_ole_file`).
//!
//! Both build the same [`litchi_ppt::Presentation`] reader the facade builds,
//! with the same metadata projection. The facade then dropped the package;
//! the handle keeps it (filesystem route) or seals the parsed bytes and CFB
//! index into a [`SharedOleFile`] without a copy or a second parse (byte
//! route), so a text editor can be built at the first edit.

use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::{OleFile, SharedOleFile, SharedOleFileLimits};
use litchi_core::{Metadata, OwnedSource, ReadAt, SourceVersion};
use litchi_ole_common::property_set::PropertySetReader;
use litchi_ppt::text_edit::{self, SourceSnapshot};
use litchi_ppt::{RecordLimits, SourceBackedPackage};

use super::LegacyOpen;
use crate::{OpenOptions, Result};

/// An opened legacy PowerPoint (`.ppt`) presentation.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`. Reads go to the presentation reader built at open. The
/// handle also retains its source (unless [`OpenOptions::retain_source`] was
/// `false`), so a later edit needs no second open.
#[derive(Clone)]
pub struct Ppt(Arc<PptState>);

struct PptState {
    /// The reader every M0a read goes through, built at open exactly as the
    /// facade builds it.
    presentation: litchi_ppt::Presentation,
    /// The OLE property-set projection, read at open as the facade reads it
    /// (record 0003 §1.6); `None` when the streams are absent, unreadable or
    /// empty.
    metadata: Option<Metadata>,
    /// The record limits the reader was built under, inherited by editors.
    record_limits: RecordLimits,
    /// The source an editor is built from.
    owner: PptOwner,
}

/// The two source states of [`Ppt`] (record 0003 §2.3).
enum PptOwner {
    /// Filesystem route: the positional package detection opened.
    Source(SourceBackedPackage),
    /// Byte route: the input bytes and the CFB index detection built, sealed
    /// together.
    Eager(Arc<SharedOleFile>),
    /// The source was released after open (`retain_source(false)`).
    Released,
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Ppt>();
    assert!(size_of::<Ppt>() == size_of::<usize>());
};

impl Ppt {
    /// Opens the presentation over the positional package filesystem
    /// detection opened (the facade's `from_native_ppt_package`).
    ///
    /// The metadata is projected through the package's shared CFB view, the
    /// reader is built, and the package is kept as the source state.
    pub(crate) fn from_source_package(
        package: SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        let metadata = package
            .metadata()
            .ok()
            .map(Metadata::from)
            .filter(Metadata::has_data);
        let (presentation, record_limits) = match options.ppt_limits {
            None => (package.presentation()?, package.record_limits()),
            Some(limits) => {
                let limits = limits.constrained_by(package.record_limits());
                (package.presentation_with_limits(limits)?, limits)
            },
        };
        let owner = if options.retain_source {
            PptOwner::Source(package)
        } else {
            PptOwner::Released
        };
        Ok(Self(Arc::new(PptState {
            presentation,
            metadata,
            record_limits,
            owner,
        })))
    }

    /// Opens the presentation from the OLE file byte detection already parsed
    /// (the facade's `DetectedFormat::Ppt` arm).
    ///
    /// The PPT package adopts the parse, the metadata streams are projected,
    /// and the reader is built (with `password` when [`Locked::unlock`]
    /// supplies one). Then the parse moves into the retained source; no byte
    /// is copied and the container is not indexed again. This keeps the
    /// input alive for as long as the handle (risk R5);
    /// [`OpenOptions::retain_source`]`(false)` releases it instead.
    ///
    /// An encrypted presentation opened without a password comes back as
    /// [`LegacyOpen::PasswordRequired`] with the container unchanged.
    ///
    /// [`Locked::unlock`]: crate::Locked
    pub(crate) fn from_ole_file(
        ole: OleFile<Cursor<Vec<u8>>>,
        options: &OpenOptions,
        password: Option<&str>,
    ) -> Result<LegacyOpen<Self>> {
        let mut package = match options.ppt_limits {
            None => litchi_ppt::Package::from_ole_file(ole)?,
            Some(limits) => litchi_ppt::Package::from_ole_file_with_limits(ole, limits)?,
        };
        let metadata = package
            .ole_file()
            .get_metadata()
            .ok()
            .map(Metadata::from)
            .filter(Metadata::has_data);
        let opened = match password {
            None => package.presentation(),
            Some(password) => package.presentation_with_options(
                litchi_ppt::OpenOptions::default().with_password(password),
            ),
        };
        let presentation = match opened {
            Ok(presentation) => presentation,
            Err(litchi_ppt::Error::PasswordRequired) => {
                return Ok(LegacyOpen::PasswordRequired(Box::new(
                    package.into_ole_file(),
                )));
            },
            Err(error) => return Err(error.into()),
        };
        let record_limits = package.record_limits();
        let owner = if options.retain_source {
            let shared = SharedOleFile::from_owned_ole_file(
                package.into_ole_file(),
                sealed_version()?,
                SharedOleFileLimits::default(),
            )?;
            PptOwner::Eager(Arc::new(shared))
        } else {
            PptOwner::Released
        };
        Ok(LegacyOpen::Opened(Self(Arc::new(PptState {
            presentation,
            metadata,
            record_limits,
            owner,
        }))))
    }

    /// All text of the presentation, as the facade's `Presentation::text`
    /// returns it.
    ///
    /// # Errors
    ///
    /// Returns a typed PPT error when a slide cannot be decoded.
    pub fn text(&self) -> Result<String> {
        Ok(self.0.presentation.text()?)
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// Never fails today; the `Result` matches the unified read surface.
    pub fn slide_count(&self) -> Result<usize> {
        Ok(self.0.presentation.slide_count())
    }

    /// The presentation's OLE property-set metadata, read at open; `None`
    /// when the presentation carries none.
    ///
    /// # Errors
    ///
    /// Never fails today; the `Result` matches the unified read surface.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        Ok(self.0.metadata.clone())
    }

    /// Everything the reader repaired: the open, then every main master's,
    /// slide's and notes page's shapes, which this call parses.
    ///
    /// # Errors
    ///
    /// Returns a typed PPT error when a slide, notes page or shape tree
    /// cannot be parsed.
    pub fn tolerance_report(&self) -> Result<litchi_ppt::ToleranceReport> {
        Ok(self.0.presentation.collect_tolerance_report()?)
    }

    /// The reader the family views read slides from.
    pub(crate) fn reader(&self) -> &litchi_ppt::Presentation {
        &self.0.presentation
    }

    /// Builds the source-backed text editor over the retained source
    /// (L-PPT-1), without indexing the container again.
    ///
    /// Returns `Ok(None)` when the source was released at open. The editor's
    /// protected-container refusal and complete-artifact fingerprint run
    /// here, at the first edit. M2 maps the leaf error into the unified one.
    #[cfg_attr(not(test), expect(dead_code, reason = "edit verbs arrive in M2"))]
    pub(crate) fn text_editor(&self) -> text_edit::Result<Option<SourceSnapshot>> {
        match &self.0.owner {
            PptOwner::Source(package) => package
                .text_edit_snapshot_with_limits(self.0.record_limits)
                .map(Some),
            PptOwner::Eager(shared) => SourceSnapshot::from_shared(
                Arc::clone(shared),
                text_edit::SourceBackedOptions {
                    record_limits: self.0.record_limits,
                    ..text_edit::SourceBackedOptions::default()
                },
            )
            .map(Some),
            PptOwner::Released => Ok(None),
        }
    }

    /// The length of the retained in-memory input, if the byte route
    /// retained one.
    #[cfg(test)]
    fn retained_input_len(&self) -> Option<u64> {
        match &self.0.owner {
            PptOwner::Eager(shared) => Some(shared.file_size()),
            PptOwner::Source(_) | PptOwner::Released => None,
        }
    }
}

impl std::fmt::Debug for Ppt {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let owner = match &self.0.owner {
            PptOwner::Source(_) => "source",
            PptOwner::Eager(_) => "eager",
            PptOwner::Released => "released",
        };
        formatter
            .debug_struct("Ppt")
            .field("owner", &owner)
            .finish_non_exhaustive()
    }
}

/// A process-unique identity for a sealed in-memory source, minted by the
/// same counter as every other [`OwnedSource`].
fn sealed_version() -> Result<SourceVersion> {
    Ok(OwnedSource::new(Vec::new()).version()?)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "test assertions panic by design"
    )]

    use super::*;
    use crate::detect::{DetectedFormat, detect_format_smart};
    use litchi_core::Position;
    use litchi_ppt::Package;
    use std::path::PathBuf;

    fn fixture_path(name: &str) -> PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt")
            .join(name)
    }

    fn fixture(name: &str) -> Vec<u8> {
        std::fs::read(fixture_path(name)).expect("PPT fixture")
    }

    /// The facade's byte route: detect, then hand the parse to the handle.
    /// A password-protected presentation reports the leaf's refusal, as the
    /// facade did; the unified opener turns it into `File::Locked`.
    fn open_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Result<Ppt> {
        match detect_format_smart(bytes) {
            Some(DetectedFormat::Ppt(ole)) => match Ppt::from_ole_file(ole, options, None)? {
                LegacyOpen::Opened(ppt) => Ok(ppt),
                LegacyOpen::PasswordRequired(_) => Err(litchi_ppt::Error::PasswordRequired.into()),
            },
            other => panic!("detected {other:?}, not PPT"),
        }
    }

    /// The facade's filesystem route, through the same path detector the
    /// facade calls in this build.
    #[cfg(any(unix, windows))]
    fn open_path(name: &str, options: &OpenOptions) -> Result<Ppt> {
        open_file(&fixture_path(name), options)
    }

    #[cfg(any(unix, windows))]
    fn open_file(path: &std::path::Path, options: &OpenOptions) -> Result<Ppt> {
        let name = path.display();
        #[cfg(feature = "pptx")]
        let package = {
            use crate::detect::detected::{
                PptxSourcePathDetection, detect_pptx_source_path_with_limits,
            };
            match detect_pptx_source_path_with_limits(path, litchi_opc::ReadLimits::default()) {
                Ok(Some(PptxSourcePathDetection::Ppt(package))) => package,
                _ => panic!("{name} was not detected as a positional PPT"),
            }
        };
        #[cfg(not(feature = "pptx"))]
        let package = {
            use crate::detect::detected::{
                PresentationSourcePathDetection, detect_presentation_source_path,
            };
            match detect_presentation_source_path(path) {
                Ok(PresentationSourcePathDetection::Ppt(package)) => package,
                _ => panic!("{name} was not detected as a positional PPT"),
            }
        };
        Ppt::from_source_package(package, options)
    }

    /// The leaf pipeline the facade wraps, as an independent oracle.
    fn leaf(name: &str) -> (litchi_ppt::Presentation, Option<Metadata>) {
        let mut package = Package::from_reader(Cursor::new(fixture(name))).unwrap();
        let metadata = package
            .ole_file()
            .get_metadata()
            .ok()
            .map(Metadata::from)
            .filter(Metadata::has_data);
        (package.presentation().unwrap(), metadata)
    }

    const FIXTURES: &[&str] = &[
        "SampleShow.ppt",
        "empty.ppt",
        "empty_textbox.ppt",
        "text_shapes.ppt",
        "text-margins.ppt",
        "ppt_with_png.ppt",
    ];

    fn assert_matches_leaf(ppt: &Ppt, name: &str) {
        let (presentation, metadata) = leaf(name);
        assert_eq!(ppt.text().unwrap(), presentation.text().unwrap(), "{name}");
        assert_eq!(
            ppt.slide_count().unwrap(),
            presentation.slide_count(),
            "{name}"
        );
        assert_eq!(
            format!("{:?}", ppt.metadata().unwrap()),
            format!("{metadata:?}"),
            "{name}"
        );
        assert_eq!(
            ppt.tolerance_report().unwrap(),
            presentation.collect_tolerance_report().unwrap(),
            "{name}"
        );
        let slides = presentation.slides().unwrap();
        for (index, slide) in slides.iter().enumerate() {
            let selected = ppt.reader().slide_at(index).unwrap().expect("slide");
            assert_eq!(selected.text().unwrap(), slide.text().unwrap(), "{name}");
        }
    }

    #[test]
    fn byte_route_matches_the_leaf_reader() {
        for name in FIXTURES {
            let ppt = open_bytes(fixture(name), &OpenOptions::default()).unwrap();
            assert_matches_leaf(&ppt, name);
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn filesystem_route_matches_the_leaf_reader() {
        for name in FIXTURES {
            let ppt = open_path(name, &OpenOptions::default()).unwrap();
            assert!(matches!(ppt.0.owner, PptOwner::Source(_)));
            assert_matches_leaf(&ppt, name);
        }
    }

    #[test]
    fn record_limits_reach_the_reader() {
        let limits = RecordLimits {
            max_input_bytes: 16,
            ..RecordLimits::default()
        };
        let options = OpenOptions::default().ppt_limits(limits);
        let error = match open_bytes(fixture("SampleShow.ppt"), &options) {
            Ok(_) => panic!("a PPT larger than the package limit opened"),
            Err(error) => error,
        };
        assert_eq!(error.code(), crate::Code::LimitExceeded);

        #[cfg(any(unix, windows))]
        {
            let error = match open_path("SampleShow.ppt", &options) {
                Ok(_) => panic!("a PPT larger than the stream limit opened"),
                Err(error) => error,
            };
            assert_eq!(error.code(), crate::Code::LimitExceeded);
        }
    }

    #[test]
    fn byte_route_retains_the_input_without_a_copy_and_releases_on_request() {
        let bytes = fixture("SampleShow.ppt");
        let length = u64::try_from(bytes.len()).unwrap();
        let retained = open_bytes(bytes.clone(), &OpenOptions::default()).unwrap();
        // Risk R5: the byte route now keeps exactly the input alive.
        assert_eq!(retained.retained_input_len(), Some(length));

        let released = open_bytes(bytes, &OpenOptions::default().retain_source(false)).unwrap();
        assert_eq!(released.retained_input_len(), None);
        assert!(released.text_editor().unwrap().is_none());
        assert_eq!(released.text().unwrap(), retained.text().unwrap());
    }

    /// Every shape text the editor resolves equals the fresh positional
    /// open's, refusals included; returns how many resolved.
    fn assert_editor_reads_like(
        editor: &SourceSnapshot,
        oracle: &SourceSnapshot,
        ppt: &Ppt,
    ) -> usize {
        assert_eq!(editor.fingerprint(), oracle.fingerprint());
        assert_eq!(editor.len(), oracle.len());
        let mut resolved = 0_usize;
        for (slide, shapes) in ppt.reader().slides().unwrap().iter().enumerate() {
            for shape in 0..shapes.shapes().unwrap().len() {
                let target = text_edit::Target::new(Position::new(slide), Position::new(shape));
                let (read, expected) = (editor.read_text(target), oracle.read_text(target));
                assert_eq!(format!("{read:?}"), format!("{expected:?}"), "{target:?}");
                resolved += usize::from(read.is_ok());
            }
        }
        resolved
    }

    /// A presentation whose text boxes the narrow source-backed editor
    /// admits; the corpus fixtures carry dependency records it refuses.
    fn editable_presentation() -> Vec<u8> {
        let mut writer = litchi_ppt::writer::Writer::new();
        for slide_index in 0..2 {
            let slide = writer.add_slide().unwrap();
            for shape_index in 0..3 {
                let x = 10 + shape_index * 260;
                let text = format!("slide-{slide_index}-shape-{shape_index}");
                writer.add_textbox(slide, x, 10, 240, 40, &text).unwrap();
            }
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn editors_are_built_from_the_retained_source() {
        let bytes = editable_presentation();
        let oracle = SourceSnapshot::open(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();

        let from_bytes = open_bytes(bytes.clone(), &OpenOptions::default()).unwrap();
        let editor = from_bytes.text_editor().unwrap().expect("retained source");
        assert_eq!(assert_editor_reads_like(&editor, &oracle, &from_bytes), 6);
        // Built over the retained source, not over a copy: the sealed
        // source's process-unique identity carries over.
        let PptOwner::Eager(shared) = &from_bytes.0.owner else {
            panic!("byte route did not retain its source");
        };
        assert_eq!(editor.source_version(), shared.captured_source_version());

        // The corpus is refused by the narrow editor, as by a fresh open.
        let refused = open_bytes(fixture("text_shapes.ppt"), &OpenOptions::default()).unwrap();
        let editor = refused.text_editor().unwrap().expect("retained source");
        let fresh =
            SourceSnapshot::open(Arc::new(OwnedSource::new(fixture("text_shapes.ppt")))).unwrap();
        assert_eq!(assert_editor_reads_like(&editor, &fresh, &refused), 0);

        #[cfg(any(unix, windows))]
        {
            let file = tempfile::NamedTempFile::new().unwrap();
            std::fs::write(file.path(), &bytes).unwrap();
            let from_path = open_file(file.path(), &OpenOptions::default()).unwrap();
            let editor = from_path.text_editor().unwrap().expect("retained package");
            assert_eq!(assert_editor_reads_like(&editor, &oracle, &from_path), 6);
            let PptOwner::Source(package) = &from_path.0.owner else {
                panic!("filesystem route did not retain its package");
            };
            assert_eq!(editor.source_version(), package.source_version().unwrap());
            let released =
                open_file(file.path(), &OpenOptions::default().retain_source(false)).unwrap();
            assert!(released.text_editor().unwrap().is_none());
        }
    }

    #[test]
    fn handles_are_cheap_shared_clones() {
        let ppt = open_bytes(fixture("SampleShow.ppt"), &OpenOptions::default()).unwrap();
        let clone = ppt.clone();
        assert!(Arc::ptr_eq(&ppt.0, &clone.0));
        std::thread::scope(|scope| {
            scope.spawn(|| assert_eq!(clone.text().unwrap(), ppt.text().unwrap()));
        });
    }
}
