use super::model::{
    Error, Limits, OpenOptions, Package, PackageOpenOptions, ResourceKind, ResourceLimit, Result,
};
use crate::document::Document;
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::Path;

impl Package<File> {
    /// Open a .doc package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .doc file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let pkg = Package::open("document.doc")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with(path, PackageOpenOptions::default())
    }

    /// Open a `.doc` package using explicit package-open options.
    ///
    /// This is the ergonomic bounded alternative to [`Self::open`]. It keeps
    /// the simple default API intact while making the selected limits visible
    /// at the call site.
    pub fn open_with<P: AsRef<Path>>(path: P, options: PackageOpenOptions) -> Result<Self> {
        let file = File::open(path)?;
        Package::from_reader_with_limits(file, options.limits())
    }

    /// Open a `.doc` package with explicit finite read limits.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: Limits) -> Result<Self> {
        Self::open_with(path, PackageOpenOptions::default().with_limits(limits))
    }
}

impl<R: Read + Seek> Package<R> {
    /// Create a Package from any reader that implements Read + Seek.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .doc file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi_doc::Package;
    ///
    /// let file = File::open("document.doc")?;
    /// let pkg = Package::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, Limits::default())
    }

    /// Create a package whose document reads inherit explicit finite limits.
    pub fn from_reader_with_limits(mut reader: R, limits: Limits) -> Result<Self> {
        preflight_source_len(&mut reader, limits)?;
        let ole = OleFile::open(reader)?;

        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(Error::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole, limits })
    }

    /// Create a Package from an already-parsed OLE file.
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
    /// use litchi_cfb::OleFile;
    /// use litchi_doc::Package;
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("document.doc")?;
    /// let ole = OleFile::open(Cursor::new(bytes))?;
    /// let pkg = Package::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        Self::from_ole_file_with_limits(ole, Limits::default())
    }

    /// Wrap an already-parsed OLE file with explicit document-read limits.
    pub fn from_ole_file_with_limits(ole: OleFile<R>, limits: Limits) -> Result<Self> {
        validate_package_size(ole.file_size(), limits)?;
        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(Error::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole, limits })
    }

    /// Get the main document.
    ///
    /// Returns the `Document` object which provides access to the document's
    /// content, formatting, tables, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_doc::Package;
    ///
    /// let mut pkg = Package::open("document.doc")?;
    /// let doc = pkg.document()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn document(&mut self) -> Result<Document> {
        Document::from_ole_with_limits(&mut self.ole, self.limits)
    }

    /// Get the main document using explicit password-to-open options.
    pub fn document_with_options(&mut self, options: OpenOptions) -> Result<Document> {
        Document::from_ole_with_options(&mut self.ole, options, self.limits)
    }

    /// Open the document with explicit limits and no password.
    pub fn document_with_limits(&mut self, limits: Limits) -> Result<Document> {
        Document::from_ole_with_options(
            &mut self.ole,
            OpenOptions::default(),
            limits.constrained_by(self.limits),
        )
    }

    /// Open the document with password/leniency options and explicit limits.
    pub fn document_with_options_and_limits(
        &mut self,
        options: OpenOptions,
        limits: Limits,
    ) -> Result<Document> {
        Document::from_ole_with_options(&mut self.ole, options, limits.constrained_by(self.limits))
    }

    /// Limits inherited by [`Self::document`].
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Inspect the package's typed MS-OFFCRYPTO `DataSpaces` graph.
    ///
    /// This is a structural, inert view of encryption and Information Rights
    /// Management metadata. It validates the CFB graph, transform headers,
    /// licenses, integrity sidecars, and custom-XML promotion markers, but it
    /// never decrypts content, evaluates a license, contacts a rights server,
    /// or activates protected content. `None` means the package has no
    /// `DataSpaces` graph.
    pub fn data_spaces(
        &mut self,
    ) -> std::result::Result<Option<litchi_crypto::spaces::Graph>, litchi_crypto::spaces::Error>
    {
        litchi_crypto::spaces::inspect(&mut self.ole)
    }

    /// Capture a source-preserving, validated `DataSpaces` edit owner.
    ///
    /// The returned snapshot retains only `DataSpaces` stream bytes. Its patch
    /// can later be written to a fresh destination with
    /// [`Self::write_data_spaces_patch`]; the current package is never
    /// mutated in place and no protected payload is decrypted or evaluated.
    pub fn data_spaces_snapshot(
        &mut self,
    ) -> std::result::Result<Option<litchi_crypto::spaces::Snapshot>, litchi_crypto::spaces::Error>
    {
        litchi_crypto::spaces::Snapshot::from_ole(&mut self.ole)
    }

    /// Rebuild the package after source-checking a `DataSpaces` patch.
    ///
    /// All logical OLE streams and storages are copied, while only the patch's
    /// `DataSpaces` stream replacements are applied. The output must be a fresh
    /// seekable destination; the generic `Package<R>` reader remains intact.
    pub fn write_data_spaces_patch<W: Write + Seek>(
        &mut self,
        patch: &litchi_crypto::spaces::Patch,
        output: &mut W,
    ) -> std::result::Result<(), litchi_crypto::spaces::Error> {
        patch.write_to(&mut self.ole, output)
    }

    /// Get the underlying OLE file.
    ///
    /// This provides access to lower-level OLE operations and streams.
    #[inline]
    pub fn ole_file(&mut self) -> &mut OleFile<R> {
        &mut self.ole
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

    /// Parse the optional MS-DOC VBA project with safe default limits.
    ///
    /// Source is decompressed and decoded according to the project code page,
    /// but is never compiled, interpreted, or executed.
    pub fn vba(
        &mut self,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        self.vba_with(&litchi_vba::Limits::default())
    }

    /// Parse the optional MS-DOC VBA project with explicit resource limits.
    pub fn vba_with(
        &mut self,
        limits: &litchi_vba::Limits,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
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
        litchi_vba::project::Project::open(&mut self.ole, &path, limits).map(Some)
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    pub fn custom_xml_data_store(
        &mut self,
    ) -> litchi_ole_common::custom_xml::Result<Option<litchi_ole_common::custom_xml::Store>> {
        litchi_ole_common::custom_xml::inspect(&mut self.ole)
    }

    pub fn summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify document XML signatures with the safe strict policy, without
    /// evaluating certificate trust or opening any VBA project stream.
    pub fn signatures(&mut self) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        self.signatures_with(&litchi_sign::Policy::strict())
    }

    /// Verify document XML signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole, litchi_sign::cfb::Format::Doc, policy)
    }

    pub fn document_summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    pub fn user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self
            .document_summary_information()?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }

    /// Reads inert `_PID_HLINKS` metadata and discovers DOC field-begin candidates.
    ///
    /// Stored targets and locations remain opaque strings. This method never
    /// resolves, opens, normalizes, evaluates, or executes linked content.
    /// Numeric `dwApp` matches are exposed only as
    /// [`crate::HyperlinkAssociation::FieldCandidates`]; callers must use
    /// [`crate::UserDefinedHyperlink::resolve_field`] to prove a `Field`
    /// association before it can be reordered for writing.
    pub fn user_defined_hyperlinks(&mut self) -> Result<Option<crate::UserDefinedHyperlinks>> {
        self.user_defined_hyperlinks_with_limits(crate::user_defined_hyperlinks::Limits::default())
    }

    /// Reads inert `_PID_HLINKS` metadata with explicit shared overlay limits.
    ///
    /// Limits apply after generic property-set parsing and bound only the
    /// reserved metadata overlay. Targets and locations remain opaque strings.
    pub fn user_defined_hyperlinks_with_limits(
        &mut self,
        limits: crate::user_defined_hyperlinks::Limits,
    ) -> Result<Option<crate::UserDefinedHyperlinks>> {
        let Some(section) = self.user_defined_properties()? else {
            return Ok(None);
        };
        let document = self.document()?;
        crate::user_defined_hyperlinks::from_user_defined_section_with_limits(
            &section,
            document.fields_table(),
            limits,
        )
    }
}

fn preflight_source_len<R: Read + Seek>(reader: &mut R, limits: Limits) -> Result<()> {
    let original_position = reader.stream_position()?;
    match reader.seek(SeekFrom::End(0)) {
        Ok(source_len) => {
            reader.seek(SeekFrom::Start(original_position))?;
            validate_package_size(source_len, limits)
        },
        Err(_seek_error) => {
            reader.seek(SeekFrom::Start(0))?;
            let maximum = u64::try_from(limits.max_package_bytes()).unwrap_or(u64::MAX);
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
            validate_package_size(observed, limits)
        },
    }
}

fn validate_package_size(actual: u64, limits: Limits) -> Result<()> {
    let maximum = u64::try_from(limits.max_package_bytes()).unwrap_or(u64::MAX);
    if actual > maximum {
        return Err(Error::ResourceLimit(ResourceLimit::new(
            ResourceKind::Package,
            actual,
            maximum,
            None,
        )));
    }
    Ok(())
}
