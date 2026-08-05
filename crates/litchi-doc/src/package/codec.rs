use super::model::{Error, OpenOptions, Package, Result};
use crate::document::Document;
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::fs::File;
use std::io::{Read, Seek};
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
        let file = File::open(path)?;
        Package::from_reader(file)
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
        let ole = OleFile::open(reader)?;

        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(Error::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
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
        // Verify it's a Word document by checking for the WordDocument stream
        if !ole.exists(&["WordDocument"]) {
            return Err(Error::InvalidFormat(
                "Not a valid Word document: WordDocument stream not found".to_string(),
            ));
        }

        Ok(Self { ole })
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
        Document::from_ole(&mut self.ole)
    }

    /// Get the main document using explicit password-to-open options.
    pub fn document_with_options(&mut self, options: OpenOptions<'_>) -> Result<Document> {
        Document::from_ole_with_options(&mut self.ole, options)
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
}
