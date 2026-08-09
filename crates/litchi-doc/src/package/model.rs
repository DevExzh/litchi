use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};
use zeroize::Zeroizing;

/// Finite resource limits for legacy Word package and stream ingestion.
///
/// The package limit is enforced before the CFB container is parsed. Stream
/// limits are enforced from directory metadata before a stream payload is
/// materialized, and the aggregate limit covers `WordDocument`, the selected
/// table stream, and the optional `Data` stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_package_bytes: usize,
    max_input_bytes: usize,
    max_aggregate_input_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_package_bytes: Self::DEFAULT_MAX_PACKAGE_BYTES,
            max_input_bytes: Self::DEFAULT_MAX_INPUT_BYTES,
            max_aggregate_input_bytes: Self::DEFAULT_MAX_AGGREGATE_INPUT_BYTES,
        }
    }
}

impl Limits {
    /// Default outer-package limit used by convenience open operations.
    pub const DEFAULT_MAX_PACKAGE_BYTES: usize = 128 * 1024 * 1024;
    /// Default per-stream limit used by convenience open operations.
    pub const DEFAULT_MAX_INPUT_BYTES: usize = 64 * 1024 * 1024;
    /// Default aggregate DOC-stream limit used by convenience open operations.
    pub const DEFAULT_MAX_AGGREGATE_INPUT_BYTES: usize = 96 * 1024 * 1024;
    /// Hard safety ceiling for an outer CFB package.
    pub const MAX_PACKAGE_BYTES: usize = 1024 * 1024 * 1024;
    /// Hard safety ceiling for one DOC-owned input stream.
    pub const MAX_INPUT_BYTES: usize = 512 * 1024 * 1024;
    /// Hard safety ceiling for aggregate DOC-owned stream bytes.
    pub const MAX_AGGREGATE_INPUT_BYTES: usize = 768 * 1024 * 1024;

    /// Construct an explicit limit set within the hard safety ceilings.
    ///
    /// # Errors
    ///
    /// Returns [`LimitsError`] when any requested value exceeds its ceiling.
    pub const fn try_new(
        max_package_bytes: usize,
        max_input_bytes: usize,
        max_aggregate_input_bytes: usize,
    ) -> std::result::Result<Self, LimitsError> {
        if max_package_bytes > Self::MAX_PACKAGE_BYTES {
            return Err(LimitsError::new(
                ResourceKind::Package,
                max_package_bytes,
                Self::MAX_PACKAGE_BYTES,
            ));
        }
        if max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(LimitsError::new(
                ResourceKind::Stream,
                max_input_bytes,
                Self::MAX_INPUT_BYTES,
            ));
        }
        if max_aggregate_input_bytes > Self::MAX_AGGREGATE_INPUT_BYTES {
            return Err(LimitsError::new(
                ResourceKind::Aggregate,
                max_aggregate_input_bytes,
                Self::MAX_AGGREGATE_INPUT_BYTES,
            ));
        }
        Ok(Self {
            max_package_bytes,
            max_input_bytes,
            max_aggregate_input_bytes,
        })
    }

    /// Return the outer-package byte limit.
    #[must_use]
    pub const fn max_package_bytes(self) -> usize {
        self.max_package_bytes
    }

    /// Return the per-stream byte limit.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Return the aggregate DOC-stream byte limit.
    #[must_use]
    pub const fn max_aggregate_input_bytes(self) -> usize {
        self.max_aggregate_input_bytes
    }

    /// Return a copy with a checked outer-package byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`LimitsError`] when `maximum` exceeds the hard safety ceiling.
    pub const fn with_max_package_bytes(
        self,
        maximum: usize,
    ) -> std::result::Result<Self, LimitsError> {
        Self::try_new(
            maximum,
            self.max_input_bytes,
            self.max_aggregate_input_bytes,
        )
    }

    /// Return a copy with a checked per-stream byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`LimitsError`] when `maximum` exceeds the hard safety ceiling.
    pub const fn with_max_input_bytes(
        self,
        maximum: usize,
    ) -> std::result::Result<Self, LimitsError> {
        Self::try_new(
            self.max_package_bytes,
            maximum,
            self.max_aggregate_input_bytes,
        )
    }

    /// Return a copy with a checked aggregate DOC-stream byte limit.
    ///
    /// # Errors
    ///
    /// Returns [`LimitsError`] when `maximum` exceeds the hard safety ceiling.
    pub const fn with_max_aggregate_input_bytes(
        self,
        maximum: usize,
    ) -> std::result::Result<Self, LimitsError> {
        Self::try_new(self.max_package_bytes, self.max_input_bytes, maximum)
    }

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
        }
    }
}

/// Resource dimension governed by a DOC read limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ResourceKind {
    /// Bytes in the outer CFB package.
    Package,
    /// Bytes in one named CFB stream.
    Stream,
    /// Aggregate bytes in all DOC-owned main streams.
    Aggregate,
}

impl ResourceKind {
    const fn description(self) -> &'static str {
        match self {
            Self::Package => "package bytes",
            Self::Stream => "stream bytes",
            Self::Aggregate => "aggregate stream bytes",
        }
    }
}

/// Invalid custom [`Limits`] configuration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LimitsError {
    resource: ResourceKind,
    actual: usize,
    maximum: usize,
}

impl LimitsError {
    const fn new(resource: ResourceKind, actual: usize, maximum: usize) -> Self {
        Self {
            resource,
            actual,
            maximum,
        }
    }

    /// Resource dimension whose requested value is invalid.
    #[must_use]
    pub const fn resource(&self) -> ResourceKind {
        self.resource
    }

    /// Requested limit.
    #[must_use]
    pub const fn actual(&self) -> usize {
        self.actual
    }

    /// Hard safety ceiling.
    #[must_use]
    pub const fn maximum(&self) -> usize {
        self.maximum
    }
}

impl std::fmt::Display for LimitsError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "requested DOC {} limit {} exceeds safety ceiling {}",
            self.resource.description(),
            self.actual,
            self.maximum
        )
    }
}

impl std::error::Error for LimitsError {}

/// Structured failure raised when input exceeds configured DOC read limits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ResourceLimit {
    resource: ResourceKind,
    actual: u64,
    limit: u64,
    path: Option<&'static str>,
}

impl ResourceLimit {
    pub(crate) const fn new(
        resource: ResourceKind,
        actual: u64,
        limit: u64,
        path: Option<&'static str>,
    ) -> Self {
        Self {
            resource,
            actual,
            limit,
            path,
        }
    }

    /// Resource dimension that was exceeded.
    #[must_use]
    pub const fn resource(&self) -> ResourceKind {
        self.resource
    }

    /// Observed input size.
    #[must_use]
    pub const fn actual(&self) -> u64 {
        self.actual
    }

    /// Configured maximum size.
    #[must_use]
    pub const fn limit(&self) -> u64 {
        self.limit
    }

    /// CFB stream path for stream-specific failures.
    #[must_use]
    pub const fn path(&self) -> Option<&'static str> {
        self.path
    }

    const fn scope(self) -> &'static str {
        match self.path {
            Some(path) => path,
            None => match self.resource {
                ResourceKind::Package => "DOC package",
                ResourceKind::Stream => "DOC stream",
                ResourceKind::Aggregate => "DOC document streams",
            },
        }
    }
}

impl std::fmt::Display for ResourceLimit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if let Some(path) = self.path {
            write!(
                formatter,
                "DOC {} at {path} is {} bytes, limit {}",
                self.resource.description(),
                self.actual,
                self.limit
            )
        } else {
            write!(
                formatter,
                "DOC {} are {} bytes, limit {}",
                self.resource.description(),
                self.actual,
                self.limit
            )
        }
    }
}

impl std::error::Error for ResourceLimit {}

impl From<ResourceLimit> for litchi_core::ResourceLimit {
    fn from(limit: ResourceLimit) -> Self {
        Self {
            resource: litchi_core::Resource::InputBytes,
            observed: limit.actual,
            limit: limit.limit,
            scope: std::sync::Arc::from(limit.scope()),
        }
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}

/// An owned password that is cleared on drop and redacted in diagnostics.
///
/// This type is intentionally non-`Clone`. Move it into [`OpenOptions`] and
/// the parser borrows it only for password verification; no decrypted document
/// state retains the credential.
pub struct Password(Zeroizing<String>);

impl Password {
    /// Take ownership of a password without copying its allocation.
    #[must_use]
    pub fn new(value: String) -> Self {
        Self(Zeroizing::new(value))
    }

    pub(crate) fn as_str(&self) -> &str {
        self.0.as_str()
    }
}

impl From<String> for Password {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}

impl std::fmt::Debug for Password {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("Password([REDACTED])")
    }
}

/// Options controlling how a legacy Word document is opened.
#[derive(Debug, Default)]
pub struct OpenOptions {
    password: Option<Password>,
    /// How non-structural stylesheet defects are treated.
    ///
    /// Defaults to [`crate::Leniency::Strict`], which is the historical behaviour.
    pub leniency: crate::leniency::Leniency,
}

impl OpenOptions {
    /// Provide a password-to-open credential through the non-cloneable
    /// [`Password`] boundary.
    #[must_use]
    pub fn with_password(mut self, password: Password) -> Self {
        self.password = Some(password);
        self
    }

    /// Set the stylesheet-defect policy.
    #[must_use]
    pub const fn with_leniency(mut self, leniency: crate::leniency::Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    pub(crate) fn password(&self) -> Option<&str> {
        self.password.as_ref().map(Password::as_str)
    }
}

/// Options for opening a DOC CFB package before its document streams are read.
/// [`Package::open`] remains the concise, bounded default. Use
/// [`Package::open_with`] when the input contract needs an explicit limit set.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PackageOpenOptions {
    limits: Limits,
}

impl Default for PackageOpenOptions {
    fn default() -> Self {
        Self {
            limits: Limits::default(),
        }
    }
}

impl PackageOpenOptions {
    /// Replace the finite limits used while opening the CFB package.
    #[must_use]
    pub const fn with_limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// Return the finite limits selected for this open operation.
    #[must_use]
    pub const fn limits(self) -> Limits {
        self.limits
    }
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
#[non_exhaustive]
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
    /// A caller-selected finite parsing limit was exceeded.
    ResourceLimit(ResourceLimit),
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
            Error::ResourceLimit(error) => error.fmt(f),
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
    /// Limits inherited by [`Package::document`](super::Package::document).
    pub(super) limits: Limits,
}

impl From<Error> for litchi_core::Error {
    fn from(err: Error) -> Self {
        match err {
            Error::Io(e) => litchi_core::Error::Io(e),
            Error::Ole(ole_err) => litchi_core::Error::from(ole_err),
            Error::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            Error::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            Error::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            Error::ResourceLimit(limit) => litchi_core::Error::ResourceLimit(limit.into()),
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
