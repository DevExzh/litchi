use super::{Error, OpenOptions, Package, Result};
use crate::presentation::Presentation;
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;

const POWERPOINT_DOCUMENT_STREAM: &[&str] = &["PowerPoint Document"];

fn validate_powerpoint_document_stream<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    if ole.exists(POWERPOINT_DOCUMENT_STREAM) && !ole.directory_exists(POWERPOINT_DOCUMENT_STREAM) {
        return Ok(());
    }

    Err(Error::InvalidFormat(
        "Not a valid PowerPoint document: PowerPoint Document stream not found or is not a stream"
            .to_string(),
    ))
}

impl Package<File> {
    /// Open a .ppt package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .ppt file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
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
    /// * `reader` - A reader containing the .ppt file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi_ppt::Package;
    ///
    /// let file = File::open("presentation.ppt")?;
    /// let pkg = Package::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_reader(reader: R) -> Result<Self> {
        let ole = OleFile::open(reader)?;

        validate_powerpoint_document_stream(&ole)?;

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
    /// use litchi_ppt::Package;
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("presentation.ppt")?;
    /// let ole = OleFile::open(Cursor::new(bytes))?;
    /// let pkg = Package::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        validate_powerpoint_document_stream(&ole)?;

        Ok(Self { ole })
    }

    /// Get the main presentation.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn presentation(&mut self) -> Result<Presentation> {
        Presentation::from_ole(&mut self.ole)
    }

    /// Get the main presentation using explicit password-to-open options.
    pub fn presentation_with_options(&mut self, options: OpenOptions<'_>) -> Result<Presentation> {
        Presentation::from_ole_with_options(&mut self.ole, options)
    }

    /// Read the live document-comparison snapshot from the presentation.
    pub fn document_comparison(&mut self) -> Result<crate::document_comparison::Snapshot> {
        self.presentation()?.document_comparison()
    }

    /// Get the underlying OLE file.
    ///
    /// This provides access to lower-level OLE operations and streams.
    #[inline]
    pub fn ole_file(&mut self) -> &mut OleFile<R> {
        &mut self.ole
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

    /// Verify presentation XML signatures with the safe strict policy, without
    /// evaluating certificate trust or opening any VBA project stream.
    pub fn signatures(&mut self) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        self.signatures_with(&litchi_sign::Policy::strict())
    }

    /// Verify presentation XML signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole, litchi_sign::cfb::Format::Ppt, policy)
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
