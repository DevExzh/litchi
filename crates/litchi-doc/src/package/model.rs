use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};

/// Options controlling how a legacy Word document is opened.
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions<'a> {
    /// Password used for password-to-open encryption.
    pub password: Option<&'a str>,
    /// How non-structural stylesheet defects are treated.
    ///
    /// Defaults to [`crate::DocLeniency::Strict`], which is the historical behaviour.
    pub leniency: crate::leniency::DocLeniency,
}

/// Password-to-open encryption schemes identified in a DOC file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionKind {
    /// Legacy Word XOR obfuscation.
    XorObfuscation,
    /// Office CryptoAPI encryption.
    CryptoApi,
    /// An encryption version not recognized by this implementation.
    Unknown {
        /// Encryption header major version.
        major: u16,
        /// Encryption header minor version.
        minor: u16,
    },
}

/// Error types for DOC file parsing.
#[derive(Debug)]
pub enum Error {
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Invalid DOC format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
    /// The document is encrypted and no password was supplied.
    PasswordRequired,
    /// The supplied password did not validate.
    InvalidPassword,
    /// The document uses a recognized but unsupported encryption scheme.
    UnsupportedEncryption(EncryptionKind),
    /// The clear encryption header is malformed.
    MalformedEncryptionHeader(String),
    /// The file predates the Word 97 binary format this reader implements.
    ///
    /// Word 6.0 and Word 95 documents keep the structures that MS-DOC places
    /// in a table stream inside `WordDocument` instead, so they cannot be read
    /// as MS-DOC no matter how tolerant the parser is.
    UnsupportedVersion {
        /// The `nFib` value found in the FIB.
        nfib: u16,
        /// Human-readable name of the originating Word release.
        name: &'static str,
    },
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

impl From<crate::sprm::Error> for Error {
    fn from(error: crate::sprm::Error) -> Self {
        Error::Corrupted(format!("malformed SPRM sequence: {error}"))
    }
}

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Error::Io(e) => write!(f, "IO error: {}", e),
            Error::Ole(e) => write!(f, "OLE error: {}", e),
            Error::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            Error::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            Error::Corrupted(s) => write!(f, "Corrupted file: {}", s),
            Error::PasswordRequired => write!(f, "a password is required to open this document"),
            Error::InvalidPassword => write!(f, "the document password is invalid"),
            Error::UnsupportedEncryption(kind) => {
                write!(f, "unsupported DOC encryption: {kind:?}")
            },
            Error::MalformedEncryptionHeader(s) => {
                write!(f, "malformed DOC encryption header: {s}")
            },
            Error::UnsupportedVersion { nfib, name } => write!(
                f,
                "{name} documents (nFib {nfib:#06x}) predate the Word 97 binary format and are not supported"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for DOC operations.
pub type Result<T> = std::result::Result<T, Error>;

/// A Word (.doc) package.
///
/// This is the main entry point for working with legacy Word documents.
/// It wraps an OLE file and provides Word-specific functionality.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_doc::Package;
///
/// // Open an existing document
/// let mut pkg = Package::open("document.doc")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
///
/// // Extract text
/// let text = doc.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package<R: Read + Seek = File> {
    /// The underlying OLE file
    pub(super) ole: OleFile<R>,
}

impl From<Error> for litchi_core::Error {
    fn from(err: Error) -> Self {
        match err {
            Error::Io(e) => litchi_core::Error::Io(e),
            Error::Ole(ole_err) => litchi_core::Error::from(ole_err),
            Error::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            Error::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            Error::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            Error::PasswordRequired => {
                litchi_core::Error::InvalidFormat("DOC password required".to_string())
            },
            Error::InvalidPassword => {
                litchi_core::Error::InvalidFormat("invalid DOC password".to_string())
            },
            Error::UnsupportedEncryption(kind) => {
                litchi_core::Error::InvalidFormat(format!("unsupported DOC encryption: {kind:?}"))
            },
            Error::MalformedEncryptionHeader(s) => litchi_core::Error::CorruptedFile(s),
            Error::UnsupportedVersion { nfib, name } => litchi_core::Error::InvalidFormat(format!(
                "{name} documents (nFib {nfib:#06x}) are not supported"
            )),
        }
    }
}
