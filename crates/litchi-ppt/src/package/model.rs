use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};

/// Finite resource limits for legacy `PowerPoint` record ingestion.
///
/// Limits apply to the uncompressed `PowerPoint Document` stream and to every
/// record tree materialized from it. Container payloads are retained for
/// lossless access in addition to their decoded children, so copied payload
/// bytes are charged cumulatively across the whole tree.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RecordLimits {
    /// Maximum bytes in the outer CFB package when opened from a reader.
    pub max_package_bytes: usize,
    /// Maximum bytes in one input stream passed to the record parser.
    pub max_input_bytes: usize,
    /// Maximum aggregate bytes across document, Current User, and Pictures streams.
    pub max_aggregate_input_bytes: usize,
    /// Maximum encoded size of one record, including its eight-byte header.
    pub max_record_bytes: usize,
    /// Maximum number of records materialized by one parser session.
    pub max_records: usize,
    /// Maximum child nesting below a top-level record.
    pub max_depth: usize,
    /// Maximum payload bytes declared by one record.
    pub max_record_payload_bytes: usize,
    /// Maximum aggregate payload bytes copied into record-owned buffers.
    pub max_copied_payload_bytes: usize,
}

impl Default for RecordLimits {
    fn default() -> Self {
        Self {
            max_package_bytes: 1024 * 1024 * 1024,
            max_input_bytes: 512 * 1024 * 1024,
            max_aggregate_input_bytes: 768 * 1024 * 1024,
            max_record_bytes: 256 * 1024 * 1024,
            max_records: 1_000_000,
            max_depth: 128,
            max_record_payload_bytes: 256 * 1024 * 1024 - 8,
            max_copied_payload_bytes: 768 * 1024 * 1024,
        }
    }
}

impl RecordLimits {
    /// Return the component-wise stricter combination of two limit sets.
    #[must_use]
    pub const fn constrained_by(self, other: Self) -> Self {
        Self {
            max_package_bytes: min_usize(self.max_package_bytes, other.max_package_bytes),
            max_input_bytes: min_usize(self.max_input_bytes, other.max_input_bytes),
            max_aggregate_input_bytes: min_usize(
                self.max_aggregate_input_bytes,
                other.max_aggregate_input_bytes,
            ),
            max_record_bytes: min_usize(self.max_record_bytes, other.max_record_bytes),
            max_records: min_usize(self.max_records, other.max_records),
            max_depth: min_usize(self.max_depth, other.max_depth),
            max_record_payload_bytes: min_usize(
                self.max_record_payload_bytes,
                other.max_record_payload_bytes,
            ),
            max_copied_payload_bytes: min_usize(
                self.max_copied_payload_bytes,
                other.max_copied_payload_bytes,
            ),
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod limit_tests {
    use super::RecordLimits;

    #[test]
    fn constrained_limits_take_the_component_wise_minimum() {
        let package = RecordLimits {
            max_package_bytes: 100,
            max_input_bytes: 90,
            max_aggregate_input_bytes: 80,
            max_record_bytes: 70,
            max_records: 60,
            max_depth: 50,
            max_record_payload_bytes: 40,
            max_copied_payload_bytes: 30,
        };
        let request = RecordLimits {
            max_package_bytes: 10,
            max_input_bytes: 20,
            max_aggregate_input_bytes: 30,
            max_record_bytes: 40,
            max_records: 50,
            max_depth: 60,
            max_record_payload_bytes: 70,
            max_copied_payload_bytes: 80,
        };
        let limits = request.constrained_by(package);
        assert_eq!(limits.max_package_bytes, 10);
        assert_eq!(limits.max_input_bytes, 20);
        assert_eq!(limits.max_aggregate_input_bytes, 30);
        assert_eq!(limits.max_record_bytes, 40);
        assert_eq!(limits.max_records, 50);
        assert_eq!(limits.max_depth, 50);
        assert_eq!(limits.max_record_payload_bytes, 40);
        assert_eq!(limits.max_copied_payload_bytes, 30);
    }
}

/// Options controlling how a legacy `PowerPoint` presentation is opened.
#[derive(Debug, Clone, Copy, Default)]
pub struct OpenOptions<'a> {
    /// Password used for password-to-open encryption.
    pub password: Option<&'a str>,
}

/// Password-to-open encryption schemes identified in a PPT file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionKind {
    /// Office Binary Document RC4 `CryptoAPI` encryption.
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
    /// Checked `OfficeArt` parsing or validation error.
    OfficeArt(litchi_odraw::Error),
    /// Host-neutral Office Graph parsing or validation error.
    Graph(litchi_ograph::Error),
    /// Invalid PPT format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
    /// A caller-selected finite parsing limit was exceeded.
    ResourceLimit(String),
    /// A fallible parser allocation could not be satisfied.
    AllocationFailed(&'static str),
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
            Error::Io(e) => write!(f, "IO error: {e}"),
            Error::Ole(e) => write!(f, "OLE error: {e}"),
            Error::OfficeArt(e) => write!(f, "OfficeArt error: {e}"),
            Error::Graph(e) => write!(f, "Office Graph error: {e}"),
            Error::InvalidFormat(s) => write!(f, "Invalid format: {s}"),
            Error::StreamNotFound(s) => write!(f, "Stream not found: {s}"),
            Error::Corrupted(s) => write!(f, "Corrupted file: {s}"),
            Error::ResourceLimit(s) => write!(f, "Resource limit exceeded: {s}"),
            Error::AllocationFailed(context) => write!(f, "Allocation failed: {context}"),
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
            | Self::ResourceLimit(_)
            | Self::AllocationFailed(_)
            | Self::PasswordRequired
            | Self::InvalidPassword
            | Self::UnsupportedEncryption(_)
            | Self::MalformedEncryptionHeader(_) => None,
        }
    }
}

/// Result type for PPT operations.
pub type Result<T> = std::result::Result<T, Error>;

/// A `PowerPoint` (.ppt) package.
///
/// This is the main entry point for working with legacy `PowerPoint` presentations.
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
    /// Limits inherited by [`Package::presentation`](super::Package::presentation).
    pub(super) record_limits: RecordLimits,
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
            Error::Corrupted(s) | Error::MalformedEncryptionHeader(s) => {
                litchi_core::Error::CorruptedFile(s)
            },
            Error::ResourceLimit(s) => {
                litchi_core::Error::CorruptedFile(format!("PPT resource limit exceeded: {s}"))
            },
            Error::AllocationFailed(context) => {
                litchi_core::Error::CorruptedFile(format!("PPT allocation failed: {context}"))
            },
            Error::PasswordRequired => {
                litchi_core::Error::InvalidFormat("presentation password is required".to_string())
            },
            Error::InvalidPassword => {
                litchi_core::Error::InvalidFormat("invalid presentation password".to_string())
            },
            Error::UnsupportedEncryption(kind) => {
                litchi_core::Error::InvalidFormat(format!("unsupported PPT encryption: {kind:?}"))
            },
        }
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}
