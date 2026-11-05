use super::model::{
    Error, OpenOptions, Package, Reader, Result, SourceLimitKind, SourceLimits,
    preflight_declared_streams,
};
use crate::document::Document;
use crate::leniency::Leniency;
use litchi_cfb::{OleError, OleFile, OleFileLimits, SharedOleFile, SharedOleFileLimits};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_ole_common::property_set::{
    PropertySetReader, Section, SharedPropertySetReader, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::fs::File;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::path::Path;
use std::sync::Arc;

impl Reader<File> {
    /// Open a .doc package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .doc file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::raw::Reader;
    ///
    /// let pkg = Reader::open("document.doc")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open a `.doc` package under the options' finite read limits.
    pub fn open_with_options<P: AsRef<Path>>(path: P, options: OpenOptions) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let path = path.as_ref();
        let file = File::open(path)?;
        let source_len = file.metadata()?.len();
        validate_package_size(source_len, source_limits)?;
        Reader::from_reader_with_options(file, options)
    }
}

impl<R: Read + Seek> Reader<R> {
    /// Create a reader from any source that implements Read + Seek.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .doc file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi_doc::raw::Reader;
    ///
    /// let file = File::open("document.doc")?;
    /// let pkg = Reader::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, OpenOptions::default())
    }

    /// Create a reader whose document reads inherit the options' finite
    /// source limits.
    pub fn from_reader_with_options(reader: R, options: OpenOptions) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let ole = Self::open_ole_file_with_limits(reader, source_limits)?;
        Self::from_ole_file_with_options(ole, options)
    }

    /// Validate a source length before entering the public DOC-reader gate.
    ///
    /// Callers that already own the source bytes can perform this check before
    /// a strict format owner parses the same source and before handing a
    /// parsed OLE context to the public reader.
    pub(crate) fn validate_source_len(
        source_len: usize,
        source_limits: SourceLimits,
    ) -> Result<()> {
        let source_len = u64::try_from(source_len).unwrap_or(u64::MAX);
        validate_package_size(source_len, source_limits)
    }

    /// Parse the physical OLE container after applying the package byte
    /// preflight, without applying DOC-specific public-reader validation.
    pub(crate) fn open_ole_file_with_limits(
        mut reader: R,
        source_limits: SourceLimits,
    ) -> Result<OleFile<R>> {
        preflight_source_len(&mut reader, source_limits)?;
        Ok(OleFile::open(reader, cfb_limits(source_limits)?)?)
    }

    /// Validate the complete public DOC reader over an already-parsed OLE
    /// container.  This keeps the package gate separate from strict format
    /// owners that need to run first on the same physical validation context.
    pub(crate) fn validate_ole_file(
        ole: &mut OleFile<R>,
        source_limits: SourceLimits,
    ) -> Result<()> {
        Self::validate_package_ole_file(ole, source_limits)?;
        Document::from_ole_with_options(ole, OpenOptions::default().source_limits(source_limits))
            .map(|_| ())
    }

    pub(crate) fn validate_package_ole_file(
        ole: &OleFile<R>,
        source_limits: SourceLimits,
    ) -> Result<()> {
        validate_package_size(ole.file_size(), source_limits)?;

        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(Error::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(())
    }

    /// Create a reader from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole` - An already-parsed OLE file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_cfb::{OleFile, OleFileLimits};
    /// use litchi_doc::raw::Reader;
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.doc")?;
    /// let ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default())?;
    /// let pkg = Reader::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        Self::from_ole_file_with_options(ole, OpenOptions::default())
    }

    /// Wrap an already-parsed OLE file with explicit document-read limits.
    ///
    /// Input and declared-stream sizes are checked before DOC payload reads.
    /// The caller controls the CFB parser's earlier directory/FAT allocation
    /// limits; adoption cannot constrain allocations that already happened.
    pub fn from_ole_file_with_options(ole: OleFile<R>, options: OpenOptions) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let data_spaces_limits = options.data_spaces_limits_value();
        Self::validate_package_ole_file(&ole, source_limits)?;
        preflight_declared_streams(&ole, source_limits)?;
        Ok(Self {
            ole,
            source_limits,
            data_spaces_limits,
        })
    }

    /// Get the main document.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, formatting, tables, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::raw::Reader;
    ///
    /// let mut pkg = Reader::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document(&mut self) -> Result<Document> {
        Document::from_ole_with_options(
            &mut self.ole,
            OpenOptions::default().source_limits(self.source_limits),
        )
    }

    /// Get the main document using explicit password-to-open options.
    pub fn document_with_options(&mut self, options: OpenOptions) -> Result<Document> {
        let source_limits = options
            .source_limits_value()
            .constrained_by(self.source_limits);
        Document::from_ole_with_options(&mut self.ole, options.source_limits(source_limits))
    }

    /// Source limits inherited by [`Self::document`].
    pub const fn source_limits(&self) -> SourceLimits {
        self.source_limits
    }

    /// DataSpaces limits retained by this raw reader.
    pub const fn data_spaces_limits(&self) -> litchi_crypto::spaces::Limits {
        self.data_spaces_limits
    }

    /// Return the complete package bytes exactly as received at ingress.
    pub fn to_bytes(&mut self) -> Result<Vec<u8>> {
        let size = usize::try_from(self.ole.file_size()).map_err(|_| {
            Error::InvalidFormat("DOC package size does not fit in memory".to_string())
        })?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(size)
            .map_err(|source| Error::Allocation {
                resource: "DOC package output",
                source,
            })?;
        self.ole.copy_source_to(&mut output)?;
        Ok(output)
    }

    /// Stream this package to a sequential sink.
    pub fn write_to(&mut self, mut writer: impl Write) -> Result<()> {
        self.ole.copy_source_to(&mut writer)?;
        Ok(())
    }

    /// Atomically save this package through a sibling temporary file.
    pub fn save(&mut self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save this package with an explicit durability level.
    pub fn save_with_durability(
        &mut self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        self.ole
            .save_source_with_durability(path, durability)
            .map_err(Error::from)
    }

    /// Inspect the package's typed MS-OFFCRYPTO `DataSpaces` graph.
    ///
    /// This is a structural, inert view of encryption and Information Rights
    /// Management metadata. It validates the CFB graph, transform headers,
    /// licenses, integrity sidecars, and custom-XML promotion markers, but it
    /// never decrypts content, evaluates a license, contacts a rights server,
    /// or activates protected content. `None` means the package has no
    /// `DataSpaces` graph.
    pub fn data_spaces(&mut self) -> Result<Option<litchi_crypto::spaces::Graph>> {
        litchi_crypto::spaces::inspect(&mut self.ole, &self.data_spaces_limits).map_err(Error::from)
    }

    /// Capture a source-preserving, validated `DataSpaces` edit owner.
    ///
    /// The returned snapshot retains only `DataSpaces` stream bytes. Its patch
    /// can later be written to a fresh destination with
    /// [`Self::write_data_spaces_patch`]; the current package is never
    /// mutated in place and no protected payload is decrypted or evaluated.
    pub fn data_spaces_snapshot(&mut self) -> Result<Option<litchi_crypto::spaces::Snapshot>> {
        litchi_crypto::spaces::Snapshot::from_ole(&mut self.ole, &self.data_spaces_limits)
            .map_err(Error::from)
    }

    /// Rebuild the package after source-checking a `DataSpaces` patch.
    ///
    /// All logical OLE streams and storages are copied, while only the patch's
    /// `DataSpaces` stream replacements are applied. The output must be a fresh
    /// seekable destination; the generic `raw::Reader<R>` reader remains
    /// intact.
    pub fn write_data_spaces_patch<W: Write + Seek>(
        &mut self,
        patch: &litchi_crypto::spaces::Patch,
        output: &mut W,
    ) -> Result<()> {
        patch.write_to(&mut self.ole, output).map_err(Error::from)
    }

    /// Get the underlying OLE file.
    ///
    /// This provides access to lower-level OLE operations and streams.
    #[inline]
    pub fn ole_file(&mut self) -> &mut OleFile<R> {
        &mut self.ole
    }

    /// Consumes the package and returns its parsed OLE file.
    ///
    /// The parsed container index is kept, so a caller that has finished
    /// reading the document can hand the same parse to another owner (for
    /// example `litchi_cfb::SharedOleFile::from_owned_ole_file`) instead of
    /// parsing the container again.
    #[must_use]
    pub fn into_ole_file(self) -> OleFile<R> {
        self.ole
    }

    /// Discover the optional MS-DOC `Macros` project storage without opening streams.
    ///
    /// This inspects CFB directory names only. It never opens, decompresses,
    /// parses, or executes the `PROJECT`, `dir`, `_VBA_PROJECT`, `__SRP_*`,
    /// or candidate module streams.
    pub fn vba_project_storages(&self) -> Vec<crate::vba::VbaProjectStorage> {
        crate::vba::discover_vba_project_storages(&self.ole.list_streams())
    }

    /// Discover the optional MS-DOC `Macros` project storage.
    pub fn vba_project_storage(&self) -> Option<crate::vba::VbaProjectStorage> {
        self.vba_project_storages().into_iter().next()
    }

    /// Parse the optional MS-DOC VBA project with explicit resource limits.
    ///
    /// Source is decompressed and decoded according to the project code page,
    /// but is never compiled, interpreted, or executed.
    pub fn vba(
        &mut self,
        limits: &litchi_vba::Limits,
    ) -> Result<Option<litchi_vba::project::Project>> {
        let Some(storage) = self.vba_project_storage() else {
            return Ok(None);
        };
        if !storage.is_structurally_complete() {
            return Ok(None);
        }
        let path: Vec<&str> = storage
            .project_root_path()
            .iter()
            .map(String::as_str)
            .collect();
        litchi_vba::project::Project::open(&mut self.ole, &path, limits)
            .map(Some)
            .map_err(Error::from)
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    pub fn custom_xml_data_store(
        &mut self,
        limits: litchi_ole_common::custom_xml::Limits,
    ) -> Result<Option<litchi_ole_common::custom_xml::Store>> {
        litchi_ole_common::custom_xml::inspect(&mut self.ole, limits).map_err(Error::from)
    }

    pub fn summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify document XML signatures with an explicit trust-neutral policy.
    pub fn signatures(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole, litchi_sign::cfb::Format::Doc, policy)
            .map_err(Error::from)
    }

    /// Read the inert VBA `DigSigBlob` stored in PIDDSI
    /// `DocumentSummaryInformation` with explicit finite blob and payload
    /// limits.
    ///
    /// This accessor validates only the `[MS-OSHARED]` container serialization.
    /// The PKCS#7 `SignedData` payload and its
    /// `SpcIndirectDataContent`/`SpcIndirectDataContentV2` `contentInfo` form
    /// remain opaque. It retains the exact blob for source-preserving edits
    /// and never verifies certificate trust, opens a VBA project, or executes
    /// code. `None` means the property stream or its `DigitalSignature`
    /// property is absent.
    pub fn vba_signature(
        &mut self,
        limits: litchi_ole_common::vba_signature::Limits,
    ) -> Result<Option<litchi_ole_common::vba_signature::Snapshot>> {
        let Some(stream) = self.document_summary_information()? else {
            return Ok(None);
        };
        let snapshot =
            litchi_ole_common::property_set::document_summary::Snapshot::from_stream(&stream)?;
        Ok(snapshot.vba_signature_with(limits)?)
    }

    pub fn document_summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    pub fn user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self
            .document_summary_information()?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }

    /// Reads inert `_PID_HLINKS` metadata under shared overlay limits
    /// ([`crate::user_defined_hyperlinks::Limits::default`] is the default
    /// profile) and discovers DOC field-begin candidates.
    ///
    /// Limits apply after generic property-set parsing and bound only the
    /// reserved metadata overlay. Stored targets and locations remain opaque
    /// strings. This method never resolves, opens, normalizes, evaluates, or
    /// executes linked content. Numeric `dwApp` matches are exposed only as
    /// [`crate::HyperlinkAssociation::FieldCandidates`]; callers must use
    /// [`crate::UserDefinedHyperlink::resolve_field`] to prove a `Field`
    /// association before it can be reordered for writing.
    pub fn user_defined_hyperlinks(
        &mut self,
        limits: crate::user_defined_hyperlinks::Limits,
    ) -> Result<Option<crate::UserDefinedHyperlinks>> {
        let Some(section) = self.user_defined_properties()? else {
            return Ok(None);
        };
        let document = self.document()?;
        crate::user_defined_hyperlinks::from_user_defined_section(
            &section,
            document.fields_table(),
            limits,
        )
    }
}

impl Reader<Cursor<Vec<u8>>> {
    /// Parse a complete in-memory DOC artifact for exact no-op publication.
    pub fn from_bytes(bytes: &[u8]) -> Result<Reader<Cursor<Vec<u8>>>> {
        let source_limits = SourceLimits::default();
        Self::from_owned_bytes(clone_bounded_source(bytes, source_limits)?, source_limits)
    }

    pub(crate) fn from_owned_bytes(
        bytes: Vec<u8>,
        source_limits: SourceLimits,
    ) -> Result<Reader<Cursor<Vec<u8>>>> {
        Self::validate_source_len(bytes.len(), source_limits)?;
        let ole = OleFile::open(Cursor::new(bytes), cfb_limits(source_limits)?)?;
        Reader::from_ole_file_with_options(ole, OpenOptions::default().source_limits(source_limits))
    }
}

/// Result of parsing a raw owned CFB at the semantic package boundary.
///
/// A password-protected source is returned with its parsed CFB intact so a
/// caller that exposes a locked state can retain the original source and ask
/// for credentials later.  Other failures consume no successful package
/// owner.  The helper is intentionally under `raw`, because handing around a
/// cursor-backed `OleFile` is an escape hatch for format-aware facades.
pub enum OpenOutcome {
    /// A fully validated immutable package owner.
    Opened(Package),
    /// The source is encrypted and needs a password before projection.
    PasswordRequired(Box<OleFile<Cursor<Vec<u8>>>>),
}

/// Validate and adopt one already parsed owned CFB under one options envelope.
pub fn open_package(
    mut ole: OleFile<Cursor<Vec<u8>>>,
    options: OpenOptions,
) -> Result<OpenOutcome> {
    let source_limits = options.source_limits_value();
    Reader::<Cursor<Vec<u8>>>::validate_package_ole_file(&ole, source_limits)?;
    let leniency = options.leniency_value();
    match Document::from_ole_with_options(&mut ole, options) {
        Ok(document) => Package::from_validated_ole(
            ole,
            document,
            source_limits,
            leniency,
            leniency == Leniency::Strict,
        )
        .map(OpenOutcome::Opened),
        Err(Error::PasswordRequired) => Ok(OpenOutcome::PasswordRequired(Box::new(ole))),
        Err(error) => Err(error),
    }
}

impl Package {
    /// Open an immutable DOC package from a path under the default policy.
    ///
    /// Ingress seals an owned snapshot. Later changes to the original path do
    /// not change this package or its document projection; source identities
    /// identify the sealed bytes, not the original filesystem object.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open an immutable DOC package under one merged options envelope.
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self> {
        let mut file = File::open(path)?;
        Self::from_reader_with_options(&mut file, options)
    }

    /// Read and validate an immutable DOC package from a reader.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, OpenOptions::default())
    }

    /// Read and validate an immutable DOC package under one options envelope.
    pub fn from_reader_with_options<R: Read>(mut reader: R, options: OpenOptions) -> Result<Self> {
        let bytes = read_bounded_source(&mut reader, options.source_limits_value())?;
        Self::from_owned_bytes(bytes, options)
    }

    /// Parse an immutable DOC package from borrowed bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Parse an immutable DOC package under one options envelope.
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions) -> Result<Self> {
        let owned = clone_bounded_source(bytes, options.source_limits_value())?;
        Self::from_owned_bytes(owned, options)
    }

    /// Adopt an already validated, owned CFB parse without re-indexing it.
    pub fn from_ole_file(ole: OleFile<Cursor<Vec<u8>>>) -> Result<Self> {
        Self::from_ole_file_with_options(ole, OpenOptions::default())
    }

    /// Adopt an already validated, owned CFB parse under one options envelope.
    ///
    /// The profile checks input and declared-stream sizes before semantic
    /// reads. The caller must bound the earlier CFB parse; this method cannot
    /// retroactively constrain its directory/FAT allocations.
    pub fn from_ole_file_with_options(
        ole: OleFile<Cursor<Vec<u8>>>,
        options: OpenOptions,
    ) -> Result<Self> {
        match open_package(ole, options)? {
            OpenOutcome::Opened(package) => Ok(package),
            OpenOutcome::PasswordRequired(_) => Err(Error::PasswordRequired),
        }
    }

    /// Construct the minimal valid DOC package through the canonical builder.
    pub fn new() -> std::result::Result<Self, crate::writer::WriteError> {
        crate::builder::Builder::new().build()
    }

    /// Return the retained immutable CFB source used by this package.
    #[must_use]
    pub fn shared_source(&self) -> Arc<SharedOleFile> {
        Arc::clone(&self.shared)
    }

    /// Consume this package and return its retained immutable CFB source.
    #[must_use]
    pub fn into_shared_source(self) -> Arc<SharedOleFile> {
        self.shared
    }

    /// Return the retained immutable parsed document projection.
    #[must_use]
    pub fn document(&self) -> Document {
        self.document.clone()
    }

    /// Return the typed package metadata projected at open.
    #[must_use]
    pub fn metadata(&self) -> &litchi_core::Metadata {
        self.metadata.as_ref()
    }

    /// Source limits retained by this immutable package.
    #[must_use]
    pub const fn source_limits(&self) -> SourceLimits {
        self.source_limits
    }

    /// Leniency policy retained by this immutable package.
    #[must_use]
    pub const fn leniency(&self) -> Leniency {
        self.leniency
    }

    /// Check and return the current source identity.
    pub fn source_version(&self) -> Result<SourceVersion> {
        Ok(self.shared.source_version()?)
    }

    /// Copy the exact retained package bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let length = self.shared.file_size();
        let capacity = usize::try_from(length)
            .map_err(|_| Error::InvalidFormat("DOC package size does not fit in memory".into()))?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "DOC package output",
                source,
            })?;
        bytes.resize(capacity, 0);
        self.read_source_into(&mut bytes)?;
        Ok(bytes)
    }

    /// Stream the exact retained package bytes to a sequential sink.
    pub fn write_to(&self, mut writer: impl Write) -> Result<()> {
        let source = self.shared.source_arc();
        let expected = self.shared.captured_source_version();
        let before = source.version()?;
        if before != expected {
            return Err(Error::Ole(OleError::SourceChanged {
                expected,
                observed: before,
            }));
        }
        stream_source(&*source, &mut writer)?;
        let after = source.version()?;
        if after != expected {
            return Err(Error::Ole(OleError::SourceChanged {
                expected,
                observed: after,
            }));
        }
        Ok(())
    }

    /// Atomically replace a destination with the exact retained package.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically replace a destination with explicit durability.
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        let source = self.shared.source_arc();
        let expected = self.shared.captured_source_version();
        litchi_cfb::atomic::save(
            path,
            durability,
            litchi_cfb::atomic::Overwrite::Replace,
            |writer| {
                let before = source.version()?;
                if before != expected {
                    return Err(Error::Ole(OleError::SourceChanged {
                        expected,
                        observed: before,
                    }));
                }
                stream_source(&*source, writer)?;
                Ok(())
            },
            || {
                let after = source.version()?;
                if after != expected {
                    return Err(Error::Ole(OleError::SourceChanged {
                        expected,
                        observed: after,
                    }));
                }
                Ok(())
            },
        )
    }

    pub(crate) fn from_owned_bytes(bytes: Vec<u8>, options: OpenOptions) -> Result<Self> {
        let source_limits = options.source_limits_value();
        Reader::<Cursor<Vec<u8>>>::validate_source_len(bytes.len(), source_limits)?;
        let ole = OleFile::open(Cursor::new(bytes), cfb_limits(source_limits)?)?;
        let raw = Reader::from_ole_file_with_options(
            ole,
            OpenOptions::default().source_limits(source_limits),
        )?;
        Self::from_raw(raw, options)
    }

    fn from_raw(mut raw: Reader<Cursor<Vec<u8>>>, options: OpenOptions) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let leniency = options.leniency_value();
        let document = Document::from_ole_with_options(&mut raw.ole, options)?;
        Self::from_validated_ole(
            raw.ole,
            document,
            source_limits,
            leniency,
            leniency == Leniency::Strict,
        )
    }

    /// Adopt an already parsed and fully validated candidate without opening
    /// its CFB container a second time.  The caller must have applied the same
    /// package and semantic validation gates used by [`Self::from_raw`].
    /// `strict_parse` says the caller parsed `document` under the default
    /// leniency and `source_limits`; with no encryption, that is the reader
    /// gate a body snapshot of these bytes applies.
    pub(crate) fn from_validated_ole(
        ole: OleFile<Cursor<Vec<u8>>>,
        document: Document,
        source_limits: SourceLimits,
        leniency: Leniency,
        strict_parse: bool,
    ) -> Result<Self> {
        let strict_document = strict_parse && !document.fib().is_encrypted();
        // The semantic document open has already charged these streams, but
        // adoption has a second metadata path through SharedOleFile. Repeat
        // the allocation-free directory preflight immediately before that
        // path so it cannot bypass the caller's entry or aggregate profile.
        preflight_declared_streams(&ole, source_limits)?;
        let version = OwnedSource::new(Vec::new()).version()?;
        let shared = Arc::new(SharedOleFile::from_owned_ole_file(
            ole,
            version,
            SharedOleFileLimits::new(
                u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX),
            )?,
        )?);
        let metadata = shared.get_metadata().map(Into::into).map_err(Error::from)?;
        Ok(Self {
            shared,
            document,
            metadata: Arc::new(metadata),
            source_limits,
            leniency,
            edit_profile: None,
            body_source: Arc::default(),
            strict_document,
        })
    }

    /// Refuse when the retained source no longer has its captured version,
    /// with the error a read of it reports.
    pub(super) fn check_source_current(&self) -> Result<()> {
        let expected = self.shared.captured_source_version();
        let observed = self.shared.source_arc().version()?;
        if observed != expected {
            return Err(Error::Ole(OleError::SourceChanged { expected, observed }));
        }
        Ok(())
    }

    fn read_source_into(&self, output: &mut [u8]) -> Result<()> {
        let source = self.shared.source_arc();
        let expected = self.shared.captured_source_version();
        let before = source.version()?;
        if before != expected {
            return Err(Error::Ole(OleError::SourceChanged {
                expected,
                observed: before,
            }));
        }
        source.read_exact_at(0, output)?;
        let after = source.version()?;
        if after != expected {
            return Err(Error::Ole(OleError::SourceChanged {
                expected,
                observed: after,
            }));
        }
        Ok(())
    }
}

fn read_bounded_source<R: Read>(reader: &mut R, source_limits: SourceLimits) -> Result<Vec<u8>> {
    let maximum = source_limits.max_input_bytes();
    let mut bytes = Vec::new();
    let mut buffer = [0_u8; 64 * 1024];
    loop {
        let count = reader.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        let next = bytes
            .len()
            .checked_add(count)
            .ok_or_else(|| Error::InvalidFormat("DOC package input length overflow".to_string()))?;
        if next > maximum {
            return Err(Error::Limit {
                kind: SourceLimitKind::InputBytes,
                observed: u64::try_from(next).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            });
        }
        bytes
            .try_reserve(count)
            .map_err(|source| Error::Allocation {
                resource: "DOC package input",
                source,
            })?;
        bytes.extend_from_slice(&buffer[..count]);
    }
    Ok(bytes)
}

fn cfb_limits(source_limits: SourceLimits) -> Result<OleFileLimits> {
    let maximum = u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX);
    Ok(OleFileLimits::new(maximum)?)
}

fn clone_bounded_source(bytes: &[u8], source_limits: SourceLimits) -> Result<Vec<u8>> {
    Reader::<Cursor<Vec<u8>>>::validate_source_len(bytes.len(), source_limits)?;
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation {
            resource: "DOC package input",
            source,
        })?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

fn stream_source(source: &dyn ReadAt, writer: &mut impl Write) -> Result<()> {
    let length = source.len()?;
    let mut offset = 0_u64;
    let mut buffer = [0_u8; 64 * 1024];
    while offset < length {
        let count = usize::try_from(length - offset)
            .unwrap_or(buffer.len())
            .min(buffer.len());
        source.read_exact_at(offset, &mut buffer[..count])?;
        writer.write_all(&buffer[..count])?;
        offset = offset.saturating_add(u64::try_from(count).unwrap_or(u64::MAX));
    }
    Ok(())
}

fn preflight_source_len<R: Read + Seek>(reader: &mut R, source_limits: SourceLimits) -> Result<()> {
    let original_position = reader.stream_position()?;
    match reader.seek(SeekFrom::End(0)) {
        Ok(source_len) => {
            reader.seek(SeekFrom::Start(original_position))?;
            validate_package_size(source_len, source_limits)
        },
        Err(_seek_error) => {
            reader.seek(SeekFrom::Start(0))?;
            let maximum = u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX);
            let probe_len = maximum.saturating_add(1);
            let mut observed = 0u64;
            let mut buffer = [0u8; 8192];
            let count_result = (|| -> std::io::Result<()> {
                while observed < probe_len {
                    let remaining = probe_len - observed;
                    let request = usize::try_from(remaining)
                        .unwrap_or(usize::MAX)
                        .min(buffer.len());
                    let read = reader.read(&mut buffer[..request])?;
                    if read == 0 {
                        break;
                    }
                    observed = observed.saturating_add(u64::try_from(read).unwrap_or(u64::MAX));
                }
                Ok(())
            })();
            let restore_result = reader.seek(SeekFrom::Start(original_position));
            count_result?;
            restore_result?;
            validate_package_size(observed, source_limits)
        },
    }
}

fn validate_package_size(actual: u64, source_limits: SourceLimits) -> Result<()> {
    let maximum = u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX);
    if actual > maximum {
        return Err(Error::Limit {
            kind: SourceLimitKind::InputBytes,
            observed: actual,
            maximum,
        });
    }
    Ok(())
}
