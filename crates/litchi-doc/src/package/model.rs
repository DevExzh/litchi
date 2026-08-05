use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};

/// Options controlling how a legacy Word document is opened.
#[derive(Debug, Clone, Copy, Default)]
pub struct DocOpenOptions<'a> {
    /// Password used for password-to-open encryption.
    pub password: Option<&'a str>,
    /// How non-structural stylesheet defects are treated.
    ///
    /// Defaults to [`crate::DocLeniency::Strict`], which is the historical behaviour.
    pub leniency: crate::leniency::DocLeniency,
}

/// Password-to-open encryption schemes identified in a DOC file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocEncryptionKind {
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
pub enum DocError {
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
    UnsupportedEncryption(DocEncryptionKind),
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

impl From<io::Error> for DocError {
    fn from(err: io::Error) -> Self {
        DocError::Io(err)
    }
}

impl From<OleError> for DocError {
    fn from(err: OleError) -> Self {
        DocError::Ole(err)
    }
}

impl From<crate::sprm::Error> for DocError {
    fn from(error: crate::sprm::Error) -> Self {
        DocError::Corrupted(format!("malformed SPRM sequence: {error}"))
    }
}

impl std::fmt::Display for DocError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DocError::Io(e) => write!(f, "IO error: {}", e),
            DocError::Ole(e) => write!(f, "OLE error: {}", e),
            DocError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            DocError::StreamNotFound(s) => write!(f, "Stream not found: {}", s),
            DocError::Corrupted(s) => write!(f, "Corrupted file: {}", s),
            DocError::PasswordRequired => write!(f, "a password is required to open this document"),
            DocError::InvalidPassword => write!(f, "the document password is invalid"),
            DocError::UnsupportedEncryption(kind) => {
                write!(f, "unsupported DOC encryption: {kind:?}")
            },
            DocError::MalformedEncryptionHeader(s) => {
                write!(f, "malformed DOC encryption header: {s}")
            },
            DocError::UnsupportedVersion { nfib, name } => write!(
                f,
                "{name} documents (nFib {nfib:#06x}) predate the Word 97 binary format and are not supported"
            ),
        }
    }
}

impl std::error::Error for DocError {}

/// Result type for DOC operations.
pub type Result<T> = std::result::Result<T, DocError>;

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

impl From<DocError> for litchi_core::Error {
    fn from(err: DocError) -> Self {
        match err {
            DocError::Io(e) => litchi_core::Error::Io(e),
            DocError::Ole(ole_err) => litchi_core::Error::from(ole_err),
            DocError::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            DocError::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            DocError::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            DocError::PasswordRequired => {
                litchi_core::Error::InvalidFormat("DOC password required".to_string())
            },
            DocError::InvalidPassword => {
                litchi_core::Error::InvalidFormat("invalid DOC password".to_string())
            },
            DocError::UnsupportedEncryption(kind) => {
                litchi_core::Error::InvalidFormat(format!("unsupported DOC encryption: {kind:?}"))
            },
            DocError::MalformedEncryptionHeader(s) => litchi_core::Error::CorruptedFile(s),
            DocError::UnsupportedVersion { nfib, name } => litchi_core::Error::InvalidFormat(
                format!("{name} documents (nFib {nfib:#06x}) are not supported"),
            ),
        }
    }
}
