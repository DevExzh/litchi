use crate::document::Document;
use litchi_cfb::{OleError, OleFile, SharedOleFile};
use std::fs::File;
use std::io::{self, Read, Seek};
use std::sync::Arc;
use zeroize::Zeroizing;

/// Finite physical ingress limits for legacy Word packages and streams.
///
/// `input_bytes` bounds the complete CFB source before it is parsed,
/// `entry_bytes` bounds each declared CFB stream before a DOC reader materializes it, and
/// `aggregate_bytes` bounds the declared CFB stream payloads together. The
/// root mini-stream backing is excluded because mini-stream payload lengths
/// are already charged through their stream entries. The defaults follow the
/// shared office source policy (512 MiB, 512 MiB, and 2 GiB).
/// Selecting a profile stores only these three integer ceilings; it does not
/// reserve any payload memory.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: usize,
    entry_bytes: usize,
    aggregate_bytes: usize,
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self {
            input_bytes: Self::DEFAULT_MAX_INPUT_BYTES,
            entry_bytes: Self::DEFAULT_MAX_ENTRY_BYTES,
            aggregate_bytes: Self::DEFAULT_MAX_AGGREGATE_BYTES,
        }
    }
}

impl SourceLimits {
    /// Shared default complete-source input limit.
    pub const DEFAULT_MAX_INPUT_BYTES: usize = 512 * 1024 * 1024;
    /// Shared default per-entry input limit.
    pub const DEFAULT_MAX_ENTRY_BYTES: usize = 512 * 1024 * 1024;
    /// Shared default aggregate input limit.
    pub const DEFAULT_MAX_AGGREGATE_BYTES: usize = 2 * 1024 * 1024 * 1024;
    /// Hard safety ceiling for the complete source.
    pub const MAX_INPUT_BYTES: usize = 1024 * 1024 * 1024;
    /// Hard safety ceiling for one declared stream entry.
    pub const MAX_ENTRY_BYTES: usize = 512 * 1024 * 1024;
    /// Hard safety ceiling for aggregate declared stream payloads.
    pub const MAX_AGGREGATE_BYTES: usize = 2 * 1024 * 1024 * 1024;

    /// Start a checked builder from the shared defaults.
    ///
    /// ```rust
    /// use litchi_doc::SourceLimits;
    ///
    /// let limits = SourceLimits::builder()
    ///     .max_input_bytes(512 * 1024 * 1024)?
    ///     .max_entry_bytes(512 * 1024 * 1024)?
    ///     .max_aggregate_bytes(2 * 1024 * 1024 * 1024)?
    ///     .build()?;
    /// assert_eq!(limits.max_entry_bytes(), 512 * 1024 * 1024);
    /// # Ok::<(), litchi_doc::Error>(())
    /// ```
    #[must_use]
    pub fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::default(),
        }
    }

    /// Return the complete-source input limit.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Return the per-entry input limit.
    #[must_use]
    pub const fn max_entry_bytes(self) -> usize {
        self.entry_bytes
    }

    /// Return the aggregate input limit.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> usize {
        self.aggregate_bytes
    }

    /// Return the component-wise stricter combination of two profiles.
    #[must_use]
    pub const fn constrained_by(self, other: Self) -> Self {
        Self {
            input_bytes: min_usize(self.input_bytes, other.input_bytes),
            entry_bytes: min_usize(self.entry_bytes, other.entry_bytes),
            aggregate_bytes: min_usize(self.aggregate_bytes, other.aggregate_bytes),
        }
    }
}

/// Consuming builder for a checked [`SourceLimits`] profile.
#[derive(Debug, Clone, Copy)]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

impl SourceLimitsBuilder {
    /// Set the complete-source input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_INPUT_BYTES`].
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            SourceLimitKind::InputBytes,
            value,
            SourceLimits::MAX_INPUT_BYTES,
        )?;
        self.limits.input_bytes = value;
        Ok(self)
    }

    /// Set the per-entry input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_ENTRY_BYTES`].
    pub fn max_entry_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            SourceLimitKind::EntryBytes,
            value,
            SourceLimits::MAX_ENTRY_BYTES,
        )?;
        self.limits.entry_bytes = value;
        Ok(self)
    }

    /// Set the aggregate input ceiling. Values must be non-zero and no
    /// greater than [`SourceLimits::MAX_AGGREGATE_BYTES`].
    pub fn max_aggregate_bytes(mut self, value: usize) -> Result<Self> {
        validate_source_limit(
            SourceLimitKind::AggregateBytes,
            value,
            SourceLimits::MAX_AGGREGATE_BYTES,
        )?;
        self.limits.aggregate_bytes = value;
        Ok(self)
    }

    /// Finish the checked profile.
    pub fn build(self) -> Result<SourceLimits> {
        Ok(self.limits)
    }
}

fn validate_source_limit(kind: SourceLimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

/// Physical resource dimension governed by a DOC source limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Complete source bytes, before the CFB container is parsed.
    InputBytes,
    /// Bytes in one named CFB entry.
    EntryBytes,
    /// Aggregate bytes in all declared CFB stream entries.
    AggregateBytes,
}

impl SourceLimitKind {
    const fn description(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
        }
    }
}

impl std::fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.description())
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
    source_limits: SourceLimits,
    /// Resource profile retained for inert MS-OFFCRYPTO `DataSpaces` reads.
    data_spaces_limits: litchi_crypto::spaces::Limits,
    /// How opt-in stylesheet-defect repairs are treated.
    ///
    /// Defaults to [`crate::Leniency::Strict`], under which those defects
    /// remain hard errors. The Word-compatible repairs listed by
    /// [`crate::SprmDefect`] are applied under every variant and recorded in
    /// [`crate::Document::tolerance_report`]; this field does not gate them.
    leniency: crate::leniency::Leniency,
}

impl OpenOptions {
    /// Select the validated physical source limits used by this open.
    ///
    /// ```rust
    /// use litchi_doc::{OpenOptions, SourceLimits};
    ///
    /// let source_limits = SourceLimits::builder()
    ///     .max_input_bytes(64 * 1024 * 1024)?
    ///     .max_entry_bytes(32 * 1024 * 1024)?
    ///     .max_aggregate_bytes(96 * 1024 * 1024)?
    ///     .build()?;
    /// let options = OpenOptions::default().source_limits(source_limits);
    /// assert_eq!(options.source_limits_value(), source_limits);
    /// # Ok::<(), litchi_doc::Error>(())
    /// ```
    #[must_use]
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Return the physical source limits selected for this open operation.
    #[must_use]
    pub const fn source_limits_value(&self) -> SourceLimits {
        self.source_limits
    }

    /// Select the resource profile retained for inert MS-OFFCRYPTO
    /// `DataSpaces` inspection and snapshot capture.
    #[must_use]
    pub const fn data_spaces_limits(mut self, limits: litchi_crypto::spaces::Limits) -> Self {
        self.data_spaces_limits = limits;
        self
    }

    /// Return the MS-OFFCRYPTO `DataSpaces` profile selected for this open.
    #[must_use]
    pub const fn data_spaces_limits_value(&self) -> litchi_crypto::spaces::Limits {
        self.data_spaces_limits
    }

    /// Provide a password-to-open credential through the non-cloneable boundary.
    #[must_use]
    pub fn password(mut self, password: Password) -> Self {
        self.password = Some(password);
        self
    }

    /// Set the stylesheet-defect policy.
    #[must_use]
    pub const fn leniency(mut self, leniency: crate::leniency::Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    /// Return the stylesheet-defect policy selected for this open operation.
    #[must_use]
    pub const fn leniency_value(&self) -> crate::leniency::Leniency {
        self.leniency
    }

    pub(crate) fn credential(&self) -> Option<&str> {
        self.password.as_ref().map(Password::as_str)
    }
}

/// Password-to-open encryption schemes identified in a DOC file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionKind {
    /// Legacy Word XOR obfuscation.
    XorObfuscation,
    /// Office `CryptoAPI` encryption.
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
#[non_exhaustive]
pub enum Error {
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Shared OLE object capture or publication failure.
    Object(litchi_ole_common::Error),
    /// Legacy Custom XML data-store parsing failure.
    CustomXml(litchi_ole_common::custom_xml::Error),
    /// Inert MS-OVBA project parsing failure.
    Vba(litchi_vba::Error),
    /// XML signature verification failure.
    Signature(litchi_sign::Error),
    /// MS-OFFCRYPTO DataSpaces inspection or patch failure.
    Crypto(litchi_crypto::spaces::Error),
    /// Invalid DOC format
    InvalidFormat(String),
    /// Stream not found
    StreamNotFound(String),
    /// Corrupted file
    Corrupted(String),
    /// The recipient does not match the patch's recorded base snapshot.
    PatchConflict,
    /// A bounded package output buffer could not be reserved.
    Allocation {
        /// Output resource whose allocation failed.
        resource: &'static str,
        /// Original allocator failure.
        source: std::collections::TryReserveError,
    },
    /// A changed DOC publication was blocked by its active protection state.
    ProtectionDenied(crate::parts::protection::EditProtection),
    /// A caller supplied an invalid physical source limit.
    InvalidLimit {
        /// Source dimension whose configured value is invalid.
        kind: SourceLimitKind,
        /// Requested value.
        value: usize,
        /// Hard safety ceiling.
        maximum: usize,
    },
    /// A caller-selected finite source limit was exceeded during ingress.
    Limit {
        /// Source dimension that exceeded its selected limit.
        kind: SourceLimitKind,
        /// Observed value.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
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

impl std::fmt::Debug for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Keep the historical native diagnostics available for the old
        // variants, while the newly exposed foreign boundaries remain
        // content-free. Their typed causes are still available through
        // `source()`.
        match self {
            Self::Io(error) => formatter.debug_tuple("Error::Io").field(error).finish(),
            Self::Ole(error) => formatter.debug_tuple("Error::Ole").field(error).finish(),
            Self::Object(error) => formatter.debug_tuple("Error::Object").field(error).finish(),
            Self::CustomXml(_) => formatter.write_str("Error::CustomXml(..)"),
            Self::Vba(_) => formatter.write_str("Error::Vba(..)"),
            Self::Signature(_) => formatter.write_str("Error::Signature(..)"),
            Self::Crypto(_) => formatter.write_str("Error::Crypto(..)"),
            Self::InvalidFormat(value) => formatter
                .debug_tuple("Error::InvalidFormat")
                .field(value)
                .finish(),
            Self::StreamNotFound(value) => formatter
                .debug_tuple("Error::StreamNotFound")
                .field(value)
                .finish(),
            Self::Corrupted(value) => formatter
                .debug_tuple("Error::Corrupted")
                .field(value)
                .finish(),
            Self::PatchConflict => formatter.write_str("Error::PatchConflict"),
            Self::Allocation { resource, source } => formatter
                .debug_struct("Error::Allocation")
                .field("resource", resource)
                .field("source", source)
                .finish(),
            Self::ProtectionDenied(state) => formatter
                .debug_tuple("Error::ProtectionDenied")
                .field(state)
                .finish(),
            Self::InvalidLimit {
                kind,
                value,
                maximum,
            } => formatter
                .debug_struct("Error::InvalidLimit")
                .field("kind", kind)
                .field("value", value)
                .field("maximum", maximum)
                .finish(),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => formatter
                .debug_struct("Error::Limit")
                .field("kind", kind)
                .field("observed", observed)
                .field("maximum", maximum)
                .finish(),
            Self::PasswordRequired => formatter.write_str("Error::PasswordRequired"),
            Self::InvalidPassword => formatter.write_str("Error::InvalidPassword"),
            Self::UnsupportedEncryption(kind) => formatter
                .debug_tuple("Error::UnsupportedEncryption")
                .field(kind)
                .finish(),
            Self::MalformedEncryptionHeader(value) => formatter
                .debug_tuple("Error::MalformedEncryptionHeader")
                .field(value)
                .finish(),
            Self::UnsupportedVersion { nfib, name } => formatter
                .debug_struct("Error::UnsupportedVersion")
                .field("nfib", nfib)
                .field("name", name)
                .finish(),
        }
    }
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
            Error::Io(e) => write!(f, "IO error: {e}"),
            Error::Ole(e) => write!(f, "OLE error: {e}"),
            Error::Object(e) => write!(f, "OLE object error: {e}"),
            Error::CustomXml(_) => f.write_str("DOC Custom XML operation failed"),
            Error::Vba(_) => f.write_str("DOC VBA operation failed"),
            Error::Signature(_) => f.write_str("DOC signature operation failed"),
            Error::Crypto(_) => f.write_str("DOC DataSpaces operation failed"),
            Error::InvalidFormat(s) => write!(f, "Invalid format: {s}"),
            Error::StreamNotFound(s) => write!(f, "Stream not found: {s}"),
            Error::Corrupted(s) => write!(f, "Corrupted file: {s}"),
            Error::PatchConflict => f.write_str("DOC patch base does not match recipient"),
            Error::Allocation { resource, source } => {
                write!(f, "could not reserve {resource}: {source}")
            },
            Error::ProtectionDenied(state) => {
                write!(
                    f,
                    "protected DOC publication requires explicit authorization ({state:?})"
                )
            },
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                f,
                "invalid DOC source limit for {kind}: {value} (maximum {maximum})"
            ),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                f,
                "DOC source limit exceeded for {}: {observed} > {maximum}",
                kind.description()
            ),
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

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::Object(error) => Some(error),
            Self::CustomXml(error) => Some(error),
            Self::Vba(error) => Some(error),
            Self::Signature(error) => Some(error),
            Self::Crypto(error) => Some(error),
            Self::Allocation { source, .. } => Some(source),
            Self::InvalidFormat(_)
            | Self::StreamNotFound(_)
            | Self::Corrupted(_)
            | Self::PatchConflict
            | Self::ProtectionDenied(_)
            | Self::InvalidLimit { .. }
            | Self::Limit { .. }
            | Self::PasswordRequired
            | Self::InvalidPassword
            | Self::UnsupportedEncryption(_)
            | Self::MalformedEncryptionHeader(_)
            | Self::UnsupportedVersion { .. } => None,
        }
    }
}

/// Result type for DOC operations.
pub type Result<T> = std::result::Result<T, Error>;

/// Preflight every declared CFB stream under the DOC physical source policy.
/// The returned aggregate includes all stream payload lengths, including
/// metadata, embedded objects, and opaque format-specific streams. The root
/// mini-stream backing is excluded because charging it as well would count
/// mini-stream payloads twice.
pub(crate) fn preflight_declared_streams<R: Read + Seek>(
    ole: &OleFile<R>,
    source_limits: SourceLimits,
) -> Result<u64> {
    let entry_maximum = u64::try_from(source_limits.max_entry_bytes()).unwrap_or(u64::MAX);
    let aggregate_maximum = u64::try_from(source_limits.max_aggregate_bytes()).unwrap_or(u64::MAX);
    let mut aggregate = 0_u64;
    for index in 0..ole.directory_entry_count() {
        let sid = u32::try_from(index).map_err(|_| {
            Error::InvalidFormat("DOC CFB directory entry index exceeds u32".to_string())
        })?;
        let Some(entry) = ole.directory_entry_by_sid(sid) else {
            continue;
        };
        if entry.entry_type != litchi_cfb::consts::STGTY_STREAM {
            continue;
        }
        if entry.size > entry_maximum {
            return Err(Error::Limit {
                kind: SourceLimitKind::EntryBytes,
                observed: entry.size,
                maximum: entry_maximum,
            });
        }
        aggregate = aggregate.checked_add(entry.size).ok_or(Error::Limit {
            kind: SourceLimitKind::AggregateBytes,
            observed: u64::MAX,
            maximum: aggregate_maximum,
        })?;
    }
    if aggregate > aggregate_maximum {
        return Err(Error::Limit {
            kind: SourceLimitKind::AggregateBytes,
            observed: aggregate,
            maximum: aggregate_maximum,
        });
    }
    Ok(aggregate)
}

/// A cursor-backed reader of one Word (.doc) compound file.
///
/// This is the low-level, mutable codec reader (record 0351 §2.1 rule 6: a
/// non-owner noun): its lazy stream and property-set readers take `&mut
/// self` because they seek the underlying OLE file. The immutable owner is
/// [`crate::Package`].
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_doc::raw::Reader;
///
/// // Open an existing document
/// let mut pkg = Reader::open("document.doc")?;
///
/// // Get the main document
/// let doc = pkg.document()?;
///
/// // Extract text
/// let text = doc.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Reader<R: Read + Seek = File> {
    /// The underlying OLE file
    pub(super) ole: OleFile<R>,
    /// Source limits inherited by [`Package::document`](super::Package::document).
    pub(super) source_limits: SourceLimits,
    /// DataSpaces limits retained by the raw reader's crypto projections.
    pub(super) data_spaces_limits: litchi_crypto::spaces::Limits,
}

impl<R: Read + Seek> std::fmt::Debug for Reader<R> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Reader")
            .field("source_limits", &self.source_limits)
            .field("data_spaces_limits", &self.data_spaces_limits)
            .finish_non_exhaustive()
    }
}

/// Private body transaction profile retained by packages produced through the
/// canonical primary edit path.  Fresh ingress keeps this absent so its
/// ordinary `Package::edit()` uses the body engine's default profile; an
/// explicit edit or patch application records the effective profile for the
/// next immutable revision.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct EditProfile {
    pub(crate) revision_limits: crate::tracked_revision::Limits,
    pub(crate) transaction_limits: crate::body_text::TransactionLimits,
    pub(crate) protection_policy: crate::parts::protection::ProtectionPolicy,
}

/// An immutable, owned DOC package snapshot.
///
/// The seek-cursor codec reader remains available as [`crate::raw::Reader`]
/// for format-specific stream and property-set work. This owner is the
/// canonical semantic package: its CFB index and parsed document projection
/// are shared by clones and are never mutated in place.
///
/// ```
/// use litchi_doc::{Builder, Package};
/// let mut builder = Builder::new();
/// builder.add_paragraph("Shared document")?;
/// let package = builder.build()?;
/// let document = package.document();
/// let bytes = package.to_bytes()?;
/// drop(package);
/// assert!(document.text()?.contains("Shared document"));
/// let reopened = Package::from_bytes(&bytes)?;
/// assert_eq!(reopened.to_bytes()?, bytes);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Package {
    pub(super) shared: Arc<SharedOleFile>,
    pub(super) document: Document,
    pub(super) metadata: Arc<litchi_core::Metadata>,
    pub(super) source_limits: SourceLimits,
    pub(super) leniency: crate::leniency::Leniency,
    pub(super) edit_profile: Option<EditProfile>,
    /// The body-text snapshot of these exact bytes the last edit opened, with
    /// the options it was opened under (record 0311), shared by clones.
    pub(super) body_source: BodySourceMemo,
    /// Whether `document` is the strict public reader's parse of these bytes
    /// under `source_limits`: the default leniency, and no credential because
    /// the document is not encrypted. A body snapshot of these bytes then
    /// needs no second parse to pass the same reader gate (record 0377a).
    pub(super) strict_document: bool,
}

/// See [`Package::body_source`]; one entry, replaced by an edit under other
/// options.
pub(super) type BodySourceMemo =
    Arc<std::sync::Mutex<Option<(crate::body_text::OpenOptions, crate::body_text::Snapshot)>>>;

impl Clone for Package {
    fn clone(&self) -> Self {
        Self {
            shared: Arc::clone(&self.shared),
            document: self.document.clone(),
            metadata: Arc::clone(&self.metadata),
            source_limits: self.source_limits,
            leniency: self.leniency,
            edit_profile: self.edit_profile.clone(),
            body_source: Arc::clone(&self.body_source),
            strict_document: self.strict_document,
        }
    }
}

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("source_limits", &self.source_limits)
            .field("leniency", &self.leniency)
            .finish_non_exhaustive()
    }
}

// The generic core projection predates the format-owned typed errors and has
// no foreign-source carrier. Keep this projection fixed and content-free;
// litchi-ops classifies the original DOC error before this conversion.
fn core_custom_xml_error(error: litchi_ole_common::custom_xml::Error) -> litchi_core::Error {
    match error {
        litchi_ole_common::custom_xml::Error::Ole(error) => {
            litchi_core::Error::from(litchi_ole_common::Error::from(error))
        },
        litchi_ole_common::custom_xml::Error::Invalid(_)
        | litchi_ole_common::custom_xml::Error::ResourceLimit(_)
        | litchi_ole_common::custom_xml::Error::Xml(_) => {
            litchi_core::Error::CorruptedFile("DOC Custom XML operation failed".to_string())
        },
        _ => litchi_core::Error::CorruptedFile("DOC Custom XML operation failed".to_string()),
    }
}

fn core_vba_error(error: litchi_vba::Error) -> litchi_core::Error {
    match error {
        litchi_vba::Error::Cfb(error) => litchi_core::Error::from(error),
        litchi_vba::Error::InvalidData(_)
        | litchi_vba::Error::Limit { .. }
        | litchi_vba::Error::UnsupportedCodePage(_) => {
            litchi_core::Error::CorruptedFile("DOC VBA operation failed".to_string())
        },
        _ => litchi_core::Error::CorruptedFile("DOC VBA operation failed".to_string()),
    }
}

fn core_signature_error(error: litchi_sign::Error) -> litchi_core::Error {
    match error {
        litchi_sign::Error::Cfb(error) => litchi_core::Error::from(error),
        litchi_sign::Error::Encrypted
        | litchi_sign::Error::Legacy
        | litchi_sign::Error::Unsupported(_) => {
            litchi_core::Error::Unsupported("DOC signature policy is unsupported".to_string())
        },
        litchi_sign::Error::Xml(_)
        | litchi_sign::Error::Limit(_)
        | litchi_sign::Error::Sha1
        | litchi_sign::Error::Key(_)
        | litchi_sign::Error::Sign(_)
        | litchi_sign::Error::Container(_) => {
            litchi_core::Error::CorruptedFile("DOC signature operation failed".to_string())
        },
        _ => litchi_core::Error::CorruptedFile("DOC signature operation failed".to_string()),
    }
}

fn core_crypto_error(error: litchi_crypto::spaces::Error) -> litchi_core::Error {
    match error {
        litchi_crypto::spaces::Error::Ole(error) => litchi_core::Error::from(error),
        litchi_crypto::spaces::Error::Invalid(_) => {
            litchi_core::Error::CorruptedFile("DOC DataSpaces operation failed".to_string())
        },
        litchi_crypto::spaces::Error::InvalidLimit {
            kind,
            value,
            maximum,
        } => litchi_core::Error::InvalidLimit {
            kind: core_crypto_resource(kind),
            value: u64::try_from(value).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            scope: "DOC DataSpaces operation",
        },
        litchi_crypto::spaces::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: core_crypto_resource(kind),
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(maximum).unwrap_or(u64::MAX),
            scope: Arc::from("DOC DataSpaces operation"),
        }),
        litchi_crypto::spaces::Error::Allocation { .. } => {
            litchi_core::Error::CorruptedFile("DOC DataSpaces allocation failed".to_string())
        },
        litchi_crypto::spaces::Error::PatchConflict => litchi_core::Error::PatchConflict,
        _ => litchi_core::Error::CorruptedFile("DOC DataSpaces operation failed".to_string()),
    }
}

fn core_crypto_resource(kind: litchi_crypto::spaces::LimitKind) -> litchi_core::Resource {
    match kind {
        litchi_crypto::spaces::LimitKind::OutputBytes => litchi_core::Resource::OutputBytes,
        litchi_crypto::spaces::LimitKind::XmlDepth => litchi_core::Resource::Depth,
        litchi_crypto::spaces::LimitKind::Entries
        | litchi_crypto::spaces::LimitKind::Components
        | litchi_crypto::spaces::LimitKind::XmlNodes => litchi_core::Resource::Objects,
        _ => litchi_core::Resource::InputBytes,
    }
}

impl From<Error> for litchi_core::Error {
    fn from(err: Error) -> Self {
        match err {
            Error::Io(e) => litchi_core::Error::Io(e),
            Error::Ole(ole_err) => litchi_core::Error::from(ole_err),
            Error::Object(error) => litchi_core::Error::from(error),
            Error::CustomXml(error) => core_custom_xml_error(error),
            Error::Vba(error) => core_vba_error(error),
            Error::Signature(error) => core_signature_error(error),
            Error::Crypto(error) => core_crypto_error(error),
            Error::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            Error::StreamNotFound(s) => litchi_core::Error::ComponentNotFound(s),
            Error::Corrupted(s) => litchi_core::Error::CorruptedFile(s),
            Error::PatchConflict => litchi_core::Error::PatchConflict,
            Error::Allocation { resource, source } => {
                litchi_core::Error::Allocation { resource, source }
            },
            Error::ProtectionDenied(state) => litchi_core::Error::InvalidFormat(format!(
                "protected DOC publication requires explicit authorization ({state:?})"
            )),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => litchi_core::Error::InvalidFormat(format!(
                "invalid DOC source limit for {kind}: {value} (maximum {maximum})"
            )),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed,
                limit: maximum,
                scope: Arc::from(format!("DOC {}", kind.description())),
            }),
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

impl From<litchi_ole_common::Error> for Error {
    fn from(error: litchi_ole_common::Error) -> Self {
        match error {
            litchi_ole_common::Error::PatchConflict => Self::PatchConflict,
            error => Self::Object(error),
        }
    }
}

impl From<litchi_ole_common::custom_xml::Error> for Error {
    fn from(error: litchi_ole_common::custom_xml::Error) -> Self {
        Self::CustomXml(error)
    }
}

impl From<litchi_vba::Error> for Error {
    fn from(error: litchi_vba::Error) -> Self {
        Self::Vba(error)
    }
}

impl From<litchi_sign::Error> for Error {
    fn from(error: litchi_sign::Error) -> Self {
        Self::Signature(error)
    }
}

impl From<litchi_crypto::spaces::Error> for Error {
    fn from(error: litchi_crypto::spaces::Error) -> Self {
        match error {
            litchi_crypto::spaces::Error::PatchConflict => Self::PatchConflict,
            error => Self::Crypto(error),
        }
    }
}

#[cfg(test)]
mod error_tests {
    use super::Error;
    use std::error::Error as _;

    #[test]
    fn new_foreign_boundaries_redact_display_and_debug_but_keep_source() {
        let payload = "doc-secret-foreign-payload";
        let errors = [
            Error::CustomXml(litchi_ole_common::custom_xml::Error::Invalid(
                payload.to_string(),
            )),
            Error::Vba(litchi_vba::Error::InvalidData(payload.to_string())),
            Error::Signature(litchi_sign::Error::Xml(payload.to_string())),
            Error::Crypto(litchi_crypto::spaces::Error::Invalid(payload.to_string())),
        ];
        for error in errors {
            assert!(error.source().is_some());
            assert!(!error.to_string().contains(payload));
            assert!(!format!("{error:?}").contains(payload));
        }
    }

    #[test]
    fn core_projection_uses_typed_or_fixed_categories_without_foreign_text() {
        let payload = "doc-secret-signature-limit";
        let core: litchi_core::Error =
            Error::Signature(litchi_sign::Error::Limit(payload.to_string())).into();
        assert!(matches!(core, litchi_core::Error::CorruptedFile(_)));
        assert!(!core.to_string().contains(payload));

        let core: litchi_core::Error =
            Error::Signature(litchi_sign::Error::Unsupported(payload.to_string())).into();
        assert!(matches!(core, litchi_core::Error::Unsupported(_)));
        assert!(!core.to_string().contains(payload));

        let core: litchi_core::Error =
            Error::Crypto(litchi_crypto::spaces::Error::PatchConflict).into();
        assert!(matches!(core, litchi_core::Error::PatchConflict));

        let core: litchi_core::Error = Error::Crypto(litchi_crypto::spaces::Error::LimitExceeded {
            kind: litchi_crypto::spaces::LimitKind::OutputBytes,
            observed: 9,
            maximum: 8,
        })
        .into();
        assert!(matches!(core, litchi_core::Error::ResourceLimit(_)));

        let core: litchi_core::Error = Error::Crypto(litchi_crypto::spaces::Error::InvalidLimit {
            kind: litchi_crypto::spaces::LimitKind::Entries,
            value: 0,
            maximum: 1,
        })
        .into();
        assert!(matches!(core, litchi_core::Error::InvalidLimit { .. }));
    }
}
