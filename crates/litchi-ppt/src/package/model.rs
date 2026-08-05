use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};

/// Options controlling how a legacy PowerPoint presentation is opened.
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions<'a> {
    /// Password used for password-to-open encryption.
    pub password: Option<&'a str>,
}

/// Password-to-open encryption schemes identified in a PPT file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionKind {
    /// Office Binary Document RC4 CryptoAPI encryption.
    CryptoApi,
    /// An encryption version not recognized by this implementation.
    Unknown { major: u16, minor: u16 },
}

/// Error types for PPT file parsing.
#[derive(Debug)]
pub enum Error {
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
    UnsupportedEncryption(EncryptionKind),
    /// The clear encryption bootstrap or header is malformed.
    MalformedEncryptionHeader(String),
}

impl From<io::Error> for Error {
    fn from(err: io::Error) -> Self {
        Error::Io(err)
    }
}

impl From<OleError> for Error {
    fn from(err: OleError) -> Self {
        Error::Ole(err)
    }
}

impl From<litchi_odraw::Error> for Error {
    fn from(err: litchi_odraw::Error) -> Self {
        Error::OfficeArt(err)
    }
}

impl From<litchi_ograph::Error> for Error {
    fn from(err: litchi_ograph::Error) -> Self {
        Self::Graph(err)
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Io(e) => write!(f, "IO error: {}", e),
            Error::Ole(e) => write!(f, "OLE error: {}", e),
            Error::OfficeArt(e) => write!(f, "OfficeArt error: {e}"),
            Error::Graph(e) => write!(f, "Office Graph error: {e}"),
            Error::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            Error::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            Error::Corrupted(s) => write!(f, "Corrupted file: {}", s),
            Error::PasswordRequired => {
                write!(f, "a password is required to open this presentation")
            },
            Error::InvalidPassword => write!(f, "the presentation password is invalid"),
            Error::UnsupportedEncryption(kind) => {
                write!(f, "unsupported PPT encryption: {kind:?}")
            },
            Error::MalformedEncryptionHeader(s) => {
                write!(f, "malformed PPT encryption header: {s}")
            },
        }
    }
}

impl std::error::Error for Error {
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
pub type Result<T> = std::result::Result<T, Error>;

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
    pub(super) ole: OleFile<R>,
}

// `From<Error> for litchi_core::Error` lives here (not in the umbrella) so
// the orphan rule is satisfied — both source and target crates are external
// to the umbrella.
impl From<Error> for litchi_core::Error {
    fn from(err: Error) -> Self {
        match err {
            Error::Io(e) => litchi_core::Error::Io(e),
            Error::Ole(ole_err) => litchi_core::Error::from(ole_err),
            Error::OfficeArt(error) => {
                litchi_core::Error::CorruptedFile(format!("Invalid OfficeArt data: {error}"))
            },
            Error::Graph(error) => {
                litchi_core::Error::CorruptedFile(format!("Invalid Office Graph data: {error}"))
            },
            Error::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            Error::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            Error::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            Error::PasswordRequired => {
                litchi_core::Error::InvalidFormat("presentation password is required".to_string())
            },
            Error::InvalidPassword => {
                litchi_core::Error::InvalidFormat("invalid presentation password".to_string())
            },
            Error::UnsupportedEncryption(kind) => {
                litchi_core::Error::InvalidFormat(format!("unsupported PPT encryption: {kind:?}"))
            },
            Error::MalformedEncryptionHeader(s) => litchi_core::Error::CorruptedFile(s),
        }
    }
}
