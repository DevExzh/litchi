use super::{Error, OpenOptions, Package, RecordLimits, Result};
use crate::presentation::Presentation;
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

const POWERPOINT_DOCUMENT_STREAM: &[&str] = &["PowerPoint Document"];

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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, RecordLimits::default())
    }

    /// Open a `.ppt` package with explicit finite presentation-record limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: RecordLimits) -> Result<Self> {
        let file = File::open(path)?;
        Package::from_reader_with_limits(file, limits)
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
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, RecordLimits::default())
    }

    /// Create a package whose presentation reads inherit explicit record limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader_with_limits(mut reader: R, record_limits: RecordLimits) -> Result<Self> {
        let original_position = reader.stream_position()?;
        let input_len = reader.seek(SeekFrom::End(0))?;
        reader.seek(SeekFrom::Start(original_position))?;
        let input_bytes = usize::try_from(input_len).map_err(|_err| {
            Error::ResourceLimit("PPT package size exceeds this platform".to_string())
        })?;
        if input_bytes > record_limits.max_package_bytes {
            return Err(Error::ResourceLimit(format!(
                "PPT package size {input_bytes} exceeds limit {}",
                record_limits.max_package_bytes
            )));
        }
        let ole = OleFile::open(reader)?;

        validate_powerpoint_document_stream(&ole)?;

        Ok(Self { ole, record_limits })
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        Self::from_ole_file_with_limits(ole, RecordLimits::default())
    }

    /// Wrap an already-parsed OLE file with explicit presentation-record limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_ole_file_with_limits(ole: OleFile<R>, record_limits: RecordLimits) -> Result<Self> {
        validate_powerpoint_document_stream(&ole)?;

        Ok(Self { ole, record_limits })
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation(&mut self) -> Result<Presentation> {
        Presentation::from_ole_with_options(
            &mut self.ole,
            OpenOptions::default(),
            self.record_limits,
        )
    }

    /// Get the main presentation using explicit password-to-open options.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation_with_options(&mut self, options: OpenOptions<'_>) -> Result<Presentation> {
        Presentation::from_ole_with_options(&mut self.ole, options, self.record_limits)
    }

    /// Open the presentation with explicit record limits and no password.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation_with_limits(&mut self, limits: RecordLimits) -> Result<Presentation> {
        Presentation::from_ole_with_options(
            &mut self.ole,
            OpenOptions::default(),
            limits.constrained_by(self.record_limits),
        )
    }

    /// Open the presentation with password options and explicit record limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation_with_options_and_limits(
        &mut self,
        options: OpenOptions<'_>,
        limits: RecordLimits,
    ) -> Result<Presentation> {
        Presentation::from_ole_with_options(
            &mut self.ole,
            options,
            limits.constrained_by(self.record_limits),
        )
    }

    /// Limits inherited by [`Self::presentation`].
    pub fn record_limits(&self) -> RecordLimits {
        self.record_limits
    }

    /// Read the live document-comparison snapshot from the presentation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn custom_xml_data_store(
        &mut self,
    ) -> litchi_ole_common::custom_xml::Result<Option<litchi_ole_common::custom_xml::Store>> {
        litchi_ole_common::custom_xml::inspect(&mut self.ole)
    }

    /// Read the legacy Summary Information property set stream.
    ///
    /// Returns `None` when the document does not contain the stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying OLE container cannot be read.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "sign")]
    pub fn signatures(&mut self) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        self.signatures_with(&litchi_sign::Policy::strict())
    }

    /// Verify presentation XML signatures with an explicit trust-neutral policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "sign")]
    pub fn signatures_with(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole, litchi_sign::cfb::Format::Ppt, policy)
    }

    /// Read the legacy Document Summary Information property set stream.
    ///
    /// Returns `None` when the document does not contain the stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying OLE container cannot be read.
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

    /// Read the user-defined properties section of the Document Summary
    /// Information property set.
    ///
    /// Returns `None` when the document has no Document Summary Information
    /// stream or no user-defined properties section.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying OLE container cannot be read.
    pub fn user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self
            .document_summary_information()?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }
}

fn validate_powerpoint_document_stream<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    if ole.exists(POWERPOINT_DOCUMENT_STREAM) && !ole.directory_exists(POWERPOINT_DOCUMENT_STREAM) {
        return Ok(());
    }

    Err(Error::InvalidFormat(
        "Not a valid PowerPoint document: PowerPoint Document stream not found or is not a stream"
            .to_string(),
    ))
}
