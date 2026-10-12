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
//! Both retain the canonical immutable PPT owner and its owned presentation
//! projection. The native owner preserves the original CFB index and source,
//! so a text editor can be built at the first edit without reopening it.

use std::io::{Cursor, Write};
use std::sync::Arc;

use litchi_cfb::OleFile;
use litchi_core::{FileFormat, Metadata, OwnedSource, ReadAt};
use litchi_ppt::text_edit::{self, SourceSnapshot};
use litchi_ppt::{Package, SourceBackedPackage, SourceLimits};

use super::LegacyOpen;
use super::signature::{self, Opened, Signature};
use crate::{OpenOptions, Result};

/// An opened legacy PowerPoint (`.ppt`) presentation.
///
/// One `Arc` wide: cloning is one atomic increment, and the handle is
/// `Send + Sync`. Reads go to the presentation reader built at open. The
/// handle retains an edit source unless [`OpenOptions::retain_source`] is false.
/// Positional reads retain their source for lazy payloads and freshness checks
/// even when editing is disabled.
#[derive(Clone)]
pub struct Ppt(Arc<Opened<PptState>>);

struct PptState {
    /// The reader every M0a read goes through, built at open exactly as the
    /// facade builds it.
    presentation: litchi_ppt::Presentation,
    /// The OLE property-set projection, read at open as the facade reads it
    /// (record 0003 §1.6); `None` when the streams are absent, unreadable or
    /// empty.
    metadata: Option<Metadata>,
    /// The source limits the reader was built under, inherited by editors.
    source_limits: SourceLimits,
    /// Whether the opened package carries password-to-open encryption.  This
    /// survives an explicit source release so capabilities and publication
    /// keep the protection policy visible without retaining the source owner.
    encrypted: bool,
    /// The source an editor is built from.
    owner: PptOwner,
}

/// The two source states of [`Ppt`] (record 0003 §2.3).
enum PptOwner {
    /// Filesystem route: the positional package detection opened.
    Source(Package),
    /// Byte route: the input bytes and the CFB index detection built, sealed
    /// together.
    Eager(Package),
    /// Edit ownership was released after open (`retain_source(false)`).
    Released,
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<Ppt>();
    assert!(size_of::<Ppt>() == size_of::<usize>());
};

/// Whether the PowerPoint reader refused an encrypted presentation for want
/// of a password, which makes the input a [`Locked`](crate::Locked) file.
///
/// PowerPoint's only password-to-open scheme is RC4 `CryptoAPI` ([MS-PPT]
/// 2.3.7). A reader built without its `encryption` feature cannot decrypt,
/// and reports that scheme as an unsupported encryption rather than asking
/// for a password; in a build without this crate's `encryption` feature that
/// refusal also means "encrypted". With the feature, the reader asks for the
/// password, and an unsupported encryption is a scheme it cannot decrypt.
pub(crate) fn password_required(error: &litchi_ppt::Error) -> bool {
    match error {
        litchi_ppt::Error::PasswordRequired => true,
        #[cfg(not(feature = "encryption"))]
        litchi_ppt::Error::UnsupportedEncryption(litchi_ppt::EncryptionKind::CryptoApi) => true,
        _ => false,
    }
}

impl Ppt {
    /// Create and adopt a blank legacy PowerPoint package directly through
    /// the native builder. The facade does not serialize and reopen the
    /// package, so the returned handle owns the same canonical package that
    /// the builder produced.
    pub(crate) fn create(options: &crate::CreateOptions) -> Result<Self> {
        use crate::create::{Conformance, Flavor};
        use crate::error::{Code, Error};

        let refusal = if options.flavor_value() != Flavor::Document {
            Some((
                "the PPT creator does not author template presentations",
                "use Flavor::Document; template authoring is not integrated",
            ))
        } else if options.conformance_value() == Conformance::Strict {
            Some((
                "legacy PPT does not support OOXML Strict conformance",
                "use Conformance::Transitional; Strict authoring is not integrated",
            ))
        } else if options.metadata_value().is_some() {
            Some((
                "initial PPT metadata authoring is not integrated",
                "omit initial metadata",
            ))
        } else {
            None
        };
        if let Some((message, suggestion)) = refusal {
            return Err(Error::new(Code::UnsupportedCapability, message)
                .with_format(FileFormat::Ppt)
                .with_suggestion(suggestion));
        }

        let package = Package::new()?;
        Self::from_package(package)
    }

    /// Adopt a newly built canonical package as an eager, editable handle.
    fn from_package(package: Package) -> Result<Self> {
        let presentation = package.presentation()?;
        let metadata = package.metadata().ok().filter(Metadata::has_data);
        let source_limits = package.source_limits();
        let encrypted = package.is_encrypted();
        Ok(Self(Opened::new(PptState {
            presentation,
            metadata,
            source_limits,
            encrypted,
            owner: PptOwner::Eager(package),
        })))
    }

    /// Whether this handle retains an edit source, without reading any payload.
    pub(crate) fn retains_source(&self) -> bool {
        !matches!(self.0.owner, PptOwner::Released)
    }

    /// Opens the presentation over the positional package filesystem
    /// detection opened (the facade's `from_native_ppt_package`).
    ///
    /// The metadata is projected through the package's shared CFB view, the
    /// reader is built, and the package is kept as the source state.
    pub(crate) fn from_source_package(
        package: SourceBackedPackage,
        options: &OpenOptions,
    ) -> Result<Self> {
        let package = Package::from_source(package, native_options(options, None))?;
        let presentation = package.presentation()?;
        let source_limits = package.source_limits();
        let metadata = package.metadata().ok().filter(Metadata::has_data);
        let encrypted = package.is_encrypted();
        let owner = if options.retain_source {
            PptOwner::Source(package)
        } else {
            PptOwner::Released
        };
        Ok(Self(Opened::new(PptState {
            presentation,
            metadata,
            source_limits,
            encrypted,
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
        let package = match Package::from_ole_file_or_password_required(
            ole,
            native_options(options, password),
        )? {
            Ok(package) => package,
            Err(ole) => return Ok(LegacyOpen::PasswordRequired(Box::new(ole))),
        };
        let presentation = package.presentation()?;
        let metadata = package.metadata().ok().filter(Metadata::has_data);
        let source_limits = package.source_limits();
        let encrypted = package.is_encrypted();
        let owner = if options.retain_source {
            PptOwner::Eager(package)
        } else {
            PptOwner::Released
        };
        Ok(LegacyOpen::Opened(Self(Opened::new(PptState {
            presentation,
            metadata,
            source_limits,
            encrypted,
            owner,
        }))))
    }

    /// The same presentation under another signature disposition (record
    /// 0013).
    pub(crate) fn with_signature(&self, signature: Signature) -> Self {
        Self(self.0.with_signature(signature))
    }

    /// This handle's signature disposition (record 0013).
    pub(crate) fn signature(&self) -> Signature {
        self.0.signature()
    }

    /// Attach the stable path used by `File::save` without changing the
    /// immutable presentation owner.
    pub(crate) fn with_origin(&self, origin: Option<Arc<std::path::Path>>) -> Self {
        Self(self.0.with_origin(origin))
    }

    /// Return the stable path retained by a path-backed open.
    pub(crate) fn origin(&self) -> Option<Arc<std::path::Path>> {
        self.0.origin()
    }

    /// Whether ordinary publication must preserve an encryption profile.
    pub(crate) fn is_encrypted(&self) -> bool {
        self.0.encrypted
    }

    /// Fence a source-backed owner before publication.  Eager owners use an
    /// immutable owned source and therefore have a stable revision too;
    /// released read projections deliberately have no publication owner.
    pub(crate) fn check_source(&self) -> Result<()> {
        match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => {
                package.source_version().map(drop).map_err(Into::into)
            },
            PptOwner::Released => Ok(()),
        }
    }

    /// Stream the exact native package.  A signed handle preserves the
    /// signature; only `Signed::unsign` reaches the explicit signature
    /// removal adapter.  A released source cannot publish because its edit
    /// owner was intentionally discarded at open.
    pub(crate) fn write_to<W: Write>(&self, sink: &mut W) -> Result<()> {
        let package = match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => package,
            PptOwner::Released => {
                return Err(crate::Error::new(
                    crate::Code::UnsupportedCapability,
                    "the PPT source was released before publication",
                )
                .with_format(FileFormat::Ppt)
                .with_suggestion(
                    "reopen the presentation without `OpenOptions::retain_source(false)`",
                ));
            },
        };
        self.check_source()?;
        if self.is_encrypted() {
            return Err(litchi_ppt::Error::EncryptionPolicy.into());
        }
        if self.signature() == Signature::RemovalAuthorized {
            let bytes = signature::unsign_compound_file(
                package.to_encrypted_bytes()?,
                litchi_sign::cfb::Format::Ppt,
            )?;
            sink.write_all(&bytes)?;
            self.check_source()?;
        } else {
            package.write_to(sink)?;
        }
        Ok(())
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
    /// here, at the first edit. After
    /// [`Signed::unsign`](crate::Signed::unsign) the editor is built over a
    /// copy of the input whose signature storages the signature leaf removed
    /// (`litchi_sign::cfb::Editor::clear`, record 0013), so a commit from it
    /// carries no signature.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// while the presentation is still signed, with a suggestion naming
    /// `Signed::unsign`, and the leaf's typed refusal otherwise.
    #[cfg_attr(not(test), expect(dead_code, reason = "edit verbs arrive in M2"))]
    pub(crate) fn text_editor(&self) -> Result<Option<SourceSnapshot>> {
        if self.0.signature() == Signature::Signed {
            return Err(signature::refusal(FileFormat::Ppt));
        }
        let options = text_edit::SourceBackedOptions {
            source_limits: self.0.source_limits,
            ..text_edit::SourceBackedOptions::default()
        };
        if self.0.signature() == Signature::RemovalAuthorized {
            let shared = match &self.0.owner {
                PptOwner::Source(package) | PptOwner::Eager(package) => package.shared_ole_file(),
                PptOwner::Released => return Ok(None),
            };
            let bytes = signature::unsign_compound_file(
                signature::compound_file_bytes(shared)?,
                litchi_sign::cfb::Format::Ppt,
            )?;
            let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
            return Ok(Some(SourceSnapshot::open_with_options(source, options)?));
        }
        Ok(match &self.0.owner {
            PptOwner::Source(package) | PptOwner::Eager(package) => Some(
                SourceSnapshot::from_shared(package.shared_ole_file_arc(), options)?,
            ),
            PptOwner::Released => None,
        })
    }

    /// The length of the retained in-memory input, if the byte route
    /// retained one.
    #[cfg(test)]
    fn retained_input_len(&self) -> Option<u64> {
        match &self.0.owner {
            PptOwner::Eager(package) => Some(package.shared_ole_file().file_size()),
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

fn native_options<'a>(
    options: &OpenOptions,
    password: Option<&'a str>,
) -> litchi_ppt::OpenOptions<'a> {
    let native = litchi_ppt::OpenOptions::new().password(password);
    match options.ppt_source_limits {
        Some(limits) => native.source_limits(limits),
        None => native,
    }
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
    use litchi_ole_common::property_set::PropertySetReader;
    use litchi_ppt::raw::Reader;
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

    /// The filesystem route: [`crate::open_with`] opens the positional
    /// PowerPoint package.
    #[cfg(any(unix, windows))]
    fn open_path(name: &str, options: &OpenOptions) -> Result<Ppt> {
        open_file(&fixture_path(name), options)
    }

    #[cfg(any(unix, windows))]
    fn open_file(path: &std::path::Path, options: &OpenOptions) -> Result<Ppt> {
        match crate::open_with(path, options)? {
            crate::File::Ppt(ppt) => Ok(ppt),
            other => panic!("{} was not opened as PPT: {other:?}", path.display()),
        }
    }

    /// The leaf pipeline the facade wraps, as an independent oracle.
    fn leaf(name: &str) -> (litchi_ppt::Presentation, Option<Metadata>) {
        let mut package = Reader::from_reader(Cursor::new(fixture(name))).unwrap();
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
    #[cfg(any(unix, windows))]
    fn positional_admission_preserves_larger_selected_limits() {
        let limits = SourceLimits::builder()
            .max_input_bytes(1024 * 1024 * 1024)
            .unwrap()
            .max_aggregate_payload_bytes(2 * 1024 * 1024 * 1024)
            .unwrap()
            .build()
            .unwrap();
        let ppt = open_path(
            "SampleShow.ppt",
            &OpenOptions::new().ppt_source_limits(limits),
        )
        .unwrap();
        assert_eq!(ppt.0.source_limits, limits);
    }

    #[test]
    fn source_limits_reach_the_reader() {
        let limits = SourceLimits::builder()
            .max_entry_bytes(16)
            .expect("a finite entry limit is valid")
            .build()
            .expect("a checked PPT source limit is valid");
        let options = OpenOptions::default().ppt_source_limits(limits);
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
        let mut writer = litchi_ppt::writer::Builder::new();
        for slide_index in 0..2 {
            let slide = writer.add_slide().unwrap();
            for shape_index in 0..3 {
                let x = 10 + shape_index * 260;
                let text = format!("slide-{slide_index}-shape-{shape_index}");
                writer.add_textbox(slide, x, 10, 240, 40, &text).unwrap();
            }
        }
        let mut output = Cursor::new(Vec::new());
        litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
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
        assert_eq!(
            editor.source_version(),
            shared.shared_ole_file().captured_source_version()
        );

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
