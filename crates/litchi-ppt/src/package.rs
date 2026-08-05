/// Package implementation for legacy PowerPoint presentations (.ppt).
use super::presentation::Presentation;
use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};
use std::path::Path;

/// Options controlling how a legacy PowerPoint presentation is opened.
#[derive(Debug, Clone, Copy, Default)]
pub struct PptOpenOptions<'a> {
    /// Password used for password-to-open encryption.
    pub password: Option<&'a str>,
}

/// Password-to-open encryption schemes identified in a PPT file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptEncryptionKind {
    /// Office Binary Document RC4 CryptoAPI encryption.
    CryptoApi,
    /// An encryption version not recognized by this implementation.
    Unknown { major: u16, minor: u16 },
}

/// Error types for PPT file parsing.
#[derive(Debug)]
pub enum PptError {
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Checked OfficeArt parsing or validation error.
    OfficeArt(litchi_odraw::Error),
    /// Host-neutral Office Graph parsing or validation error.
    Graph(litchi_ograph::Error),
    /// Invalid PPT format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
    /// The presentation is encrypted and no password was supplied.
    PasswordRequired,
    /// The supplied password did not validate.
    InvalidPassword,
    /// The presentation uses a recognized but unsupported encryption scheme.
    UnsupportedEncryption(PptEncryptionKind),
    /// The clear encryption bootstrap or header is malformed.
    MalformedEncryptionHeader(String),
}

impl From<io::Error> for PptError {
    fn from(err: io::Error) -> Self {
        PptError::Io(err)
    }
}

impl From<OleError> for PptError {
    fn from(err: OleError) -> Self {
        PptError::Ole(err)
    }
}

impl From<litchi_odraw::Error> for PptError {
    fn from(err: litchi_odraw::Error) -> Self {
        PptError::OfficeArt(err)
    }
}

impl From<litchi_ograph::Error> for PptError {
    fn from(err: litchi_ograph::Error) -> Self {
        Self::Graph(err)
    }
}

impl std::fmt::Display for PptError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PptError::Io(e) => write!(f, "IO error: {}", e),
            PptError::Ole(e) => write!(f, "OLE error: {}", e),
            PptError::OfficeArt(e) => write!(f, "OfficeArt error: {e}"),
            PptError::Graph(e) => write!(f, "Office Graph error: {e}"),
            PptError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            PptError::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            PptError::Corrupted(s) => write!(f, "Corrupted file: {}", s),
            PptError::PasswordRequired => {
                write!(f, "a password is required to open this presentation")
            },
            PptError::InvalidPassword => write!(f, "the presentation password is invalid"),
            PptError::UnsupportedEncryption(kind) => {
                write!(f, "unsupported PPT encryption: {kind:?}")
            },
            PptError::MalformedEncryptionHeader(s) => {
                write!(f, "malformed PPT encryption header: {s}")
            },
        }
    }
}

impl std::error::Error for PptError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::OfficeArt(error) => Some(error),
            Self::Graph(error) => Some(error),
            Self::InvalidFormat(_)
            | Self::StreamNotFound(_)
            | Self::Corrupted(_)
            | Self::PasswordRequired
            | Self::InvalidPassword
            | Self::UnsupportedEncryption(_)
            | Self::MalformedEncryptionHeader(_) => None,
        }
    }
}

/// Result type for PPT operations.
pub type Result<T> = std::result::Result<T, PptError>;

const POWERPOINT_DOCUMENT_STREAM: &[&str] = &["PowerPoint Document"];

fn validate_powerpoint_document_stream<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    if ole.exists(POWERPOINT_DOCUMENT_STREAM) && !ole.directory_exists(POWERPOINT_DOCUMENT_STREAM) {
        return Ok(());
    }

    Err(PptError::InvalidFormat(
        "Not a valid PowerPoint document: PowerPoint Document stream not found or is not a stream"
            .to_string(),
    ))
}

/// A PowerPoint (.ppt) package.
///
/// This is the main entry point for working with legacy PowerPoint presentations.
/// It wraps an OLE file and provides PowerPoint-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ppt::Package;
///
/// // Open an existing presentation
/// let mut pkg = Package::open("presentation.ppt")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
///
/// // Extract text
/// let text = pres.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package<R: Read + Seek = File> {
    /// The underlying OLE file
    ole: OleFile<R>,
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
    pub fn presentation_with_options(
        &mut self,
        options: PptOpenOptions<'_>,
    ) -> Result<Presentation> {
        Presentation::from_ole_with_options(&mut self.ole, options)
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

    pub fn summary_information(&mut self) -> Result<Option<litchi_cfb::PropertySetStream>> {
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

    pub fn document_summary_information(
        &mut self,
    ) -> Result<Option<litchi_cfb::PropertySetStream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    pub fn user_defined_properties(&mut self) -> Result<Option<litchi_cfb::PropertySet>> {
        Ok(self.document_summary_information()?.and_then(|stream| {
            stream
                .section(litchi_cfb::USER_DEFINED_PROPERTIES_FMTID)
                .cloned()
        }))
    }
}

// `From<PptError> for litchi_core::Error` lives here (not in the umbrella) so
// the orphan rule is satisfied — both source and target crates are external
// to the umbrella.
impl From<PptError> for litchi_core::Error {
    fn from(err: PptError) -> Self {
        match err {
            PptError::Io(e) => litchi_core::Error::Io(e),
            PptError::Ole(ole_err) => litchi_core::Error::from(ole_err),
            PptError::OfficeArt(error) => {
                litchi_core::Error::CorruptedFile(format!("Invalid OfficeArt data: {error}"))
            },
            PptError::Graph(error) => {
                litchi_core::Error::CorruptedFile(format!("Invalid Office Graph data: {error}"))
            },
            PptError::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            PptError::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            PptError::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            PptError::PasswordRequired => {
                litchi_core::Error::InvalidFormat("presentation password is required".to_string())
            },
            PptError::InvalidPassword => {
                litchi_core::Error::InvalidFormat("invalid presentation password".to_string())
            },
            PptError::UnsupportedEncryption(kind) => {
                litchi_core::Error::InvalidFormat(format!("unsupported PPT encryption: {kind:?}"))
            },
            PptError::MalformedEncryptionHeader(s) => litchi_core::Error::CorruptedFile(s),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::OleWriter;
    use std::io::Cursor;
    use std::path::Path;

    fn serialize_ole(writer: &mut OleWriter) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn test_open_package() {
        let base = Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .unwrap()
            .parent()
            .unwrap();
        let result = Package::open(
            base.join("test-data")
                .join("ole")
                .join("ppt")
                .join("empty.ppt"),
        );
        assert!(result.is_ok());
    }

    #[test]
    fn rejects_powerpoint_document_storage_before_package_publication() {
        let mut writer = OleWriter::new();
        writer.create_storage(&["PowerPoint Document"]).unwrap();
        let bytes = serialize_ole(&mut writer);

        let from_reader = Package::from_reader(Cursor::new(bytes.clone()));
        assert!(matches!(
            from_reader,
            Err(PptError::InvalidFormat(message)) if message.contains("is not a stream")
        ));

        let ole = OleFile::open(Cursor::new(bytes)).unwrap();
        let from_ole_file = Package::from_ole_file(ole);
        assert!(matches!(
            from_ole_file,
            Err(PptError::InvalidFormat(message)) if message.contains("is not a stream")
        ));
    }

    #[test]
    #[ignore] // Requires test file
    fn test_invalid_file() {
        // Create a non-PPT file
        std::fs::write("test_invalid.tmp", b"Not a PPT file").unwrap();
        let result = Package::open("test_invalid.tmp");
        assert!(result.is_err());
        std::fs::remove_file("test_invalid.tmp").ok();
    }
}
