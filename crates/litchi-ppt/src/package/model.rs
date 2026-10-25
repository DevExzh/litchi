use crate::writer::WriteError;
use litchi_cfb::{OleError, OleFile};
use std::fs::File;
use std::io::{self, Read, Seek};
use std::sync::Arc;

/// Physical source dimensions bounded while a legacy `PowerPoint` package is
/// admitted and its record tree is materialized.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    // These names are retained at crate visibility so the established parser
    // code can keep its local vocabulary.  They are deliberately private to
    // this crate; callers use the canonical getters and checked builder below.
    pub(crate) max_package_bytes: usize,
    pub(crate) max_input_bytes: usize,
    pub(crate) max_aggregate_input_bytes: usize,
    pub(crate) max_record_bytes: usize,
    pub(crate) max_records: usize,
    pub(crate) max_depth: usize,
    pub(crate) max_record_payload_bytes: usize,
    pub(crate) max_copied_payload_bytes: usize,
}

impl std::fmt::Debug for SourceLimits {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("SourceLimits")
            .field("max_input_bytes", &self.max_input_bytes())
            .field("max_entry_bytes", &self.max_entry_bytes())
            .field("max_aggregate_bytes", &self.max_aggregate_bytes())
            .field("max_record_bytes", &self.max_record_bytes())
            .field("max_events", &self.max_events())
            .field("max_depth", &self.max_depth())
            .field("max_record_payload_bytes", &self.max_record_payload_bytes())
            .field(
                "max_aggregate_payload_bytes",
                &self.max_aggregate_payload_bytes(),
            )
            .finish()
    }
}

/// Physical source dimension governed by a PPT source limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Bytes in one declared CFB stream entry.
    EntryBytes,
    /// Aggregate bytes in all declared CFB stream entries.
    AggregateBytes,
    /// Encoded size of one PPT record, including its header.
    RecordBytes,
    /// Number of record events materialized by a parser session.
    Events,
    /// Record-tree nesting depth.
    Depth,
    /// Payload bytes declared by one record.
    RecordPayloadBytes,
    /// Aggregate record payload bytes materialized by a parser session.
    AggregatePayloadBytes,
}

impl SourceLimitKind {
    pub(crate) const fn description(self) -> &'static str {
        match self {
            Self::InputBytes => "input bytes",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::RecordBytes => "record bytes",
            Self::Events => "events",
            Self::Depth => "depth",
            Self::RecordPayloadBytes => "record payload bytes",
            Self::AggregatePayloadBytes => "aggregate payload bytes",
        }
    }
}

impl std::fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.description())
    }
}

impl SourceLimits {
    const MIB: usize = 1024 * 1024;
    const GIB: usize = 1024 * 1024 * 1024;

    /// Default complete-source input limit for PPT.
    pub const DEFAULT_MAX_INPUT_BYTES: usize = 512 * Self::MIB;
    /// Default per-entry stream limit for PPT.
    pub const DEFAULT_MAX_ENTRY_BYTES: usize = 512 * Self::MIB;
    /// Default aggregate stream limit for PPT.
    pub const DEFAULT_MAX_AGGREGATE_BYTES: usize = 2 * Self::GIB;
    /// Hard safety ceiling for the complete source.
    pub const MAX_INPUT_BYTES: usize = Self::GIB;
    /// Hard safety ceiling for one stream entry.
    pub const MAX_ENTRY_BYTES: usize = 512 * Self::MIB;
    /// Hard safety ceiling for aggregate stream bytes.
    pub const MAX_AGGREGATE_BYTES: usize = 2 * Self::GIB;
    /// Default one-record encoded-byte limit for PPT.
    pub const DEFAULT_MAX_RECORD_BYTES: usize = 256 * Self::MIB;
    pub const MAX_RECORD_BYTES: usize = 256 * Self::MIB;
    /// Default record-event limit for PPT.
    pub const DEFAULT_MAX_EVENTS: usize = 1_000_000;
    pub const MAX_EVENTS: usize = 1_000_000;
    /// Default record-depth limit for PPT.
    pub const DEFAULT_MAX_DEPTH: usize = 256;
    pub const MAX_DEPTH: usize = 256;
    /// Default one-record payload limit for PPT.
    pub const DEFAULT_MAX_RECORD_PAYLOAD_BYTES: usize = 256 * Self::MIB - 8;
    pub const MAX_RECORD_PAYLOAD_BYTES: usize = 256 * Self::MIB - 8;
    /// Default aggregate payload limit for PPT.
    pub const DEFAULT_MAX_AGGREGATE_PAYLOAD_BYTES: usize = 768 * Self::MIB;
    /// Hard safety ceiling for aggregate payload bytes.
    pub const MAX_AGGREGATE_PAYLOAD_BYTES: usize = 2 * Self::GIB;

    /// The checked standard profile.
    pub const DEFAULT: Self = Self {
        max_package_bytes: Self::DEFAULT_MAX_INPUT_BYTES,
        max_input_bytes: Self::DEFAULT_MAX_ENTRY_BYTES,
        max_aggregate_input_bytes: Self::DEFAULT_MAX_AGGREGATE_BYTES,
        max_record_bytes: Self::DEFAULT_MAX_RECORD_BYTES,
        max_records: Self::DEFAULT_MAX_EVENTS,
        max_depth: Self::DEFAULT_MAX_DEPTH,
        max_record_payload_bytes: Self::DEFAULT_MAX_RECORD_PAYLOAD_BYTES,
        max_copied_payload_bytes: Self::DEFAULT_MAX_AGGREGATE_PAYLOAD_BYTES,
    };

    /// Start a checked builder from the standard defaults.
    #[must_use]
    pub const fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Start a checked builder initialized from this profile.
    #[must_use]
    pub const fn to_builder(self) -> SourceLimitsBuilder {
        SourceLimitsBuilder { limits: self }
    }

    /// Maximum complete input bytes.
    #[must_use]
    #[allow(
        clippy::misnamed_getters,
        reason = "internal parser fields retain their pre-migration names"
    )]
    pub const fn max_input_bytes(self) -> usize {
        self.max_package_bytes
    }

    /// Maximum bytes in one declared stream entry.
    #[must_use]
    pub const fn max_entry_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum aggregate bytes in all declared CFB stream entries.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> usize {
        self.max_aggregate_input_bytes
    }

    /// Maximum encoded size of one PPT record.
    #[must_use]
    pub const fn max_record_bytes(self) -> usize {
        self.max_record_bytes
    }

    /// Maximum number of record events.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_records
    }

    /// Maximum record-tree nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum payload bytes declared by one record.
    #[must_use]
    pub const fn max_record_payload_bytes(self) -> usize {
        self.max_record_payload_bytes
    }

    /// Maximum aggregate record payload bytes.
    #[must_use]
    pub const fn max_aggregate_payload_bytes(self) -> usize {
        self.max_copied_payload_bytes
    }

    /// Return the component-wise stricter combination of two profiles.
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

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for a checked [`SourceLimits`] profile.
#[derive(Debug, Clone, Copy)]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

impl SourceLimitsBuilder {
    /// Set the complete-source input ceiling.
    pub fn max_input_bytes(mut self, value: usize) -> Result<Self> {
        validate_limit(
            SourceLimitKind::InputBytes,
            value,
            SourceLimits::MAX_INPUT_BYTES,
        )?;
        self.limits.max_package_bytes = value;
        Ok(self)
    }

    /// Set the per-entry stream ceiling.
    pub fn max_entry_bytes(mut self, value: usize) -> Result<Self> {
        validate_limit(
            SourceLimitKind::EntryBytes,
            value,
            SourceLimits::MAX_ENTRY_BYTES,
        )?;
        self.limits.max_input_bytes = value;
        Ok(self)
    }

    /// Set the aggregate declared-stream ceiling.
    pub fn max_aggregate_bytes(mut self, value: usize) -> Result<Self> {
        validate_limit(
            SourceLimitKind::AggregateBytes,
            value,
            SourceLimits::MAX_AGGREGATE_BYTES,
        )?;
        self.limits.max_aggregate_input_bytes = value;
        Ok(self)
    }

    /// Set the encoded one-record ceiling.
    pub fn max_record_bytes(mut self, value: usize) -> Result<Self> {
        validate_limit(
            SourceLimitKind::RecordBytes,
            value,
            SourceLimits::MAX_RECORD_BYTES,
        )?;
        self.limits.max_record_bytes = value;
        Ok(self)
    }

    /// Set the record-event ceiling.
    pub fn max_events(mut self, value: usize) -> Result<Self> {
        validate_limit(SourceLimitKind::Events, value, SourceLimits::MAX_EVENTS)?;
        self.limits.max_records = value;
        Ok(self)
    }

    /// Set the record-depth ceiling.
    pub fn max_depth(mut self, value: usize) -> Result<Self> {
        validate_limit(SourceLimitKind::Depth, value, SourceLimits::MAX_DEPTH)?;
        self.limits.max_depth = value;
        Ok(self)
    }

    /// Set the one-record payload ceiling.
    pub fn max_record_payload_bytes(mut self, value: usize) -> Result<Self> {
        validate_limit(
            SourceLimitKind::RecordPayloadBytes,
            value,
            SourceLimits::MAX_RECORD_PAYLOAD_BYTES,
        )?;
        self.limits.max_record_payload_bytes = value;
        Ok(self)
    }

    /// Set the aggregate payload ceiling.
    pub fn max_aggregate_payload_bytes(mut self, value: usize) -> Result<Self> {
        validate_limit(
            SourceLimitKind::AggregatePayloadBytes,
            value,
            SourceLimits::MAX_AGGREGATE_PAYLOAD_BYTES,
        )?;
        self.limits.max_copied_payload_bytes = value;
        Ok(self)
    }

    /// Finish the checked profile.
    pub fn build(self) -> Result<SourceLimits> {
        validate_limit(
            SourceLimitKind::InputBytes,
            self.limits.max_package_bytes,
            SourceLimits::MAX_INPUT_BYTES,
        )?;
        validate_limit(
            SourceLimitKind::EntryBytes,
            self.limits.max_input_bytes,
            SourceLimits::MAX_ENTRY_BYTES,
        )?;
        validate_limit(
            SourceLimitKind::AggregateBytes,
            self.limits.max_aggregate_input_bytes,
            SourceLimits::MAX_AGGREGATE_BYTES,
        )?;
        validate_limit(
            SourceLimitKind::RecordBytes,
            self.limits.max_record_bytes,
            SourceLimits::MAX_RECORD_BYTES,
        )?;
        validate_limit(
            SourceLimitKind::Events,
            self.limits.max_records,
            SourceLimits::MAX_EVENTS,
        )?;
        validate_limit(
            SourceLimitKind::Depth,
            self.limits.max_depth,
            SourceLimits::MAX_DEPTH,
        )?;
        validate_limit(
            SourceLimitKind::RecordPayloadBytes,
            self.limits.max_record_payload_bytes,
            SourceLimits::MAX_RECORD_PAYLOAD_BYTES,
        )?;
        validate_limit(
            SourceLimitKind::AggregatePayloadBytes,
            self.limits.max_copied_payload_bytes,
            SourceLimits::MAX_AGGREGATE_PAYLOAD_BYTES,
        )?;
        Ok(self.limits)
    }
}

fn validate_limit(kind: SourceLimitKind, value: usize, maximum: usize) -> Result<()> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind,
            value,
            maximum,
        });
    }
    Ok(())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod limit_tests {
    use super::{Error, OpenOptions, SourceLimitKind, SourceLimits};
    use std::error::Error as _;

    #[test]
    fn constrained_limits_take_the_component_wise_minimum() {
        let package = SourceLimits {
            max_package_bytes: 100,
            max_input_bytes: 90,
            max_aggregate_input_bytes: 80,
            max_record_bytes: 70,
            max_records: 60,
            max_depth: 50,
            max_record_payload_bytes: 40,
            max_copied_payload_bytes: 30,
        };
        let request = SourceLimits {
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

    #[test]
    fn checked_builder_exposes_canonical_dimensions_and_hard_bounds() {
        let defaults = SourceLimits::default();
        assert_eq!(defaults.max_input_bytes(), 512 * 1024 * 1024);
        assert_eq!(defaults.max_entry_bytes(), 512 * 1024 * 1024);
        assert_eq!(defaults.max_aggregate_bytes(), 2 * 1024 * 1024 * 1024);
        assert_eq!(defaults.max_record_bytes(), 256 * 1024 * 1024);
        assert_eq!(defaults.max_events(), 1_000_000);
        assert_eq!(defaults.max_depth(), 256);
        assert_eq!(defaults.max_record_payload_bytes(), 256 * 1024 * 1024 - 8);
        assert_eq!(defaults.max_aggregate_payload_bytes(), 768 * 1024 * 1024);

        let limits = SourceLimits::builder()
            .max_input_bytes(SourceLimits::MAX_INPUT_BYTES)
            .unwrap()
            .max_entry_bytes(SourceLimits::MAX_ENTRY_BYTES)
            .unwrap()
            .max_aggregate_bytes(SourceLimits::MAX_AGGREGATE_BYTES)
            .unwrap()
            .max_record_bytes(SourceLimits::MAX_RECORD_BYTES)
            .unwrap()
            .max_events(SourceLimits::MAX_EVENTS)
            .unwrap()
            .max_depth(SourceLimits::MAX_DEPTH)
            .unwrap()
            .max_record_payload_bytes(SourceLimits::MAX_RECORD_PAYLOAD_BYTES)
            .unwrap()
            .max_aggregate_payload_bytes(SourceLimits::MAX_AGGREGATE_PAYLOAD_BYTES)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(limits.max_input_bytes(), SourceLimits::MAX_INPUT_BYTES);
        assert_eq!(
            limits.max_aggregate_payload_bytes(),
            SourceLimits::MAX_AGGREGATE_PAYLOAD_BYTES
        );

        let error = SourceLimits::builder().max_events(0).unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidLimit {
                kind: SourceLimitKind::Events,
                value: 0,
                maximum: SourceLimits::MAX_EVENTS,
            }
        ));
    }

    #[test]
    fn open_options_redact_credentials_and_track_explicit_limits() {
        let options = OpenOptions::new().password("secret");
        assert_eq!(options.password_value(), Some("secret"));
        assert!(!format!("{options:?}").contains("secret"));
        assert!(!options.has_explicit_source_limits());

        let limits = SourceLimits::builder()
            .max_events(1)
            .unwrap()
            .build()
            .unwrap();
        let options = options.source_limits(limits);
        assert_eq!(options.source_limits_value(), limits);
        assert!(options.has_explicit_source_limits());
    }

    #[test]
    fn encryption_policy_is_a_typed_core_refusal() {
        let error = Error::EncryptionPolicy;
        assert!(error.to_string().contains("ordinary PPT publication"));
        assert!(std::error::Error::source(&error).is_none());
        let core_error: litchi_core::Error = error.into();
        assert!(matches!(
            core_error,
            litchi_core::Error::Unsupported(message)
                if message.contains("password-to-open encryption")
        ));
    }

    #[test]
    fn new_custom_xml_boundary_redacts_foreign_payload_but_keeps_source() {
        let payload = "ppt-secret-custom-xml-payload";
        let error = Error::CustomXml(litchi_ole_common::custom_xml::Error::Invalid(
            payload.to_string(),
        ));
        assert!(error.source().is_some());
        assert!(!error.to_string().contains(payload));
        assert!(!format!("{error:?}").contains(payload));
        let core: litchi_core::Error = error.into();
        assert!(matches!(core, litchi_core::Error::CorruptedFile(_)));
        assert!(!core.to_string().contains(payload));
    }

    #[cfg(feature = "sign")]
    #[test]
    fn new_signature_boundary_keeps_typed_core_categories_without_payloads() {
        let payload = "ppt-secret-signature-payload";
        let error = Error::Signature(litchi_sign::Error::Unsupported(payload.to_string()));
        assert!(error.source().is_some());
        assert!(!error.to_string().contains(payload));
        assert!(!format!("{error:?}").contains(payload));
        let core: litchi_core::Error = error.into();
        assert!(matches!(core, litchi_core::Error::Unsupported(_)));
        assert!(!core.to_string().contains(payload));
    }
}

/// Options controlling how a legacy `PowerPoint` presentation is opened.
///
/// The password is intentionally redacted from [`std::fmt::Debug`]. Setters consume
/// the envelope so callers cannot accidentally retain a mutable policy alias.
#[derive(Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub struct OpenOptions<'a> {
    /// Checked physical source limits.
    pub(crate) source_limits: SourceLimits,
    /// Whether the caller selected the source profile explicitly.
    pub(crate) source_limits_explicit: bool,
    /// Limits retained by presentations for inert VBA inspection.
    #[cfg(feature = "vba-inspection")]
    pub(crate) vba_project_limits: crate::VbaProjectLimits,
    /// Whether the caller explicitly selected the VBA inspection profile.
    #[cfg(feature = "vba-inspection")]
    pub(crate) vba_project_limits_explicit: bool,
    /// Password used for password-to-open encryption.
    pub(crate) password: Option<&'a str>,
    /// How opt-in repairs are treated.
    ///
    /// Defaults to [`crate::Leniency::Strict`], under which those defects
    /// remain hard errors. The `PowerPoint`-compatible repairs are applied
    /// under every variant and recorded in
    /// [`crate::Presentation::tolerance_report`]; this field does not gate
    /// them.
    pub(crate) leniency: crate::leniency::Leniency,
}

impl<'a> OpenOptions<'a> {
    /// Construct strict options without a password.
    #[must_use]
    pub fn new() -> Self {
        Self {
            source_limits: SourceLimits::DEFAULT,
            source_limits_explicit: false,
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: crate::VbaProjectLimits::default(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits_explicit: false,
            password: None,
            leniency: crate::leniency::Leniency::Strict,
        }
    }

    /// Select checked physical source limits.
    #[must_use]
    pub const fn source_limits(mut self, source_limits: SourceLimits) -> Self {
        self.source_limits = source_limits;
        self.source_limits_explicit = true;
        self
    }

    /// Return the selected physical source limits.
    #[must_use]
    pub const fn source_limits_value(&self) -> SourceLimits {
        self.source_limits
    }

    /// Select the limits retained by a presentation for inert VBA inspection.
    #[cfg(feature = "vba-inspection")]
    #[must_use]
    pub const fn vba_project_limits(mut self, limits: crate::VbaProjectLimits) -> Self {
        self.vba_project_limits = limits;
        self.vba_project_limits_explicit = true;
        self
    }

    /// Return the limits retained by a presentation for inert VBA inspection.
    #[cfg(feature = "vba-inspection")]
    #[must_use]
    pub const fn vba_project_limits_value(&self) -> crate::VbaProjectLimits {
        self.vba_project_limits
    }

    #[cfg(feature = "vba-inspection")]
    pub(crate) const fn inherit_vba_project_limits(
        mut self,
        limits: crate::VbaProjectLimits,
    ) -> Self {
        if !self.vba_project_limits_explicit {
            self.vba_project_limits = limits;
        }
        self
    }

    pub(crate) const fn has_explicit_source_limits(&self) -> bool {
        self.source_limits_explicit
    }

    /// Supply an optional password-to-open credential.
    #[must_use]
    pub fn password(mut self, password: impl Into<Option<&'a str>>) -> Self {
        self.password = password.into();
        self
    }

    /// Return the password when one was supplied.
    #[must_use]
    pub const fn password_value(&self) -> Option<&'a str> {
        self.password
    }

    /// Set the opt-in repair policy.
    #[must_use]
    pub const fn leniency(mut self, leniency: crate::leniency::Leniency) -> Self {
        self.leniency = leniency;
        self
    }

    /// Return the selected repair policy.
    #[must_use]
    pub const fn leniency_value(&self) -> crate::leniency::Leniency {
        self.leniency
    }
}

impl Default for OpenOptions<'_> {
    fn default() -> Self {
        Self::new()
    }
}

impl std::fmt::Debug for OpenOptions<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("OpenOptions")
            .field("source_limits", &self.source_limits)
            .field("password", &self.password.map(|_| "<redacted>"))
            .field("leniency", &self.leniency)
            .finish()
    }
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
#[non_exhaustive]
pub enum Error {
    /// Typed native Builder authoring or codec failure.
    Authoring(WriteError),
    /// IO error
    Io(io::Error),
    /// OLE file error
    Ole(OleError),
    /// Shared OLE object or Property Set error with typed classification.
    Object(litchi_ole_common::Error),
    /// Legacy Custom XML data-store parsing failure.
    CustomXml(litchi_ole_common::custom_xml::Error),
    /// XML signature verification failure.
    #[cfg(feature = "sign")]
    Signature(litchi_sign::Error),
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
    /// A caller-selected limit was invalid at policy construction.
    InvalidLimit {
        /// Dimension whose value was invalid.
        kind: SourceLimitKind,
        /// Supplied value.
        value: usize,
        /// Hard maximum permitted for the dimension.
        maximum: usize,
    },
    /// A caller-selected finite source limit was exceeded.
    Limit {
        /// Dimension whose configured limit was exceeded.
        kind: SourceLimitKind,
        /// Observed source usage.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// A limit from a genuinely separate resource domain was exceeded.
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
    /// Ordinary publication cannot preserve the active password-to-open
    /// encryption state.
    EncryptionPolicy,
}

impl std::fmt::Debug for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Keep native authoring diagnostics opaque at this public boundary.
        // The typed leaf remains available through `source()`.
        match self {
            Self::Authoring(_) => formatter.write_str("Error::Authoring(..)"),
            Self::Io(error) => formatter.debug_tuple("Error::Io").field(error).finish(),
            Self::Ole(error) => formatter.debug_tuple("Error::Ole").field(error).finish(),
            Self::Object(error) => formatter.debug_tuple("Error::Object").field(error).finish(),
            Self::CustomXml(_) => formatter.write_str("Error::CustomXml(..)"),
            #[cfg(feature = "sign")]
            Self::Signature(_) => formatter.write_str("Error::Signature(..)"),
            Self::OfficeArt(error) => formatter
                .debug_tuple("Error::OfficeArt")
                .field(error)
                .finish(),
            Self::Graph(error) => formatter.debug_tuple("Error::Graph").field(error).finish(),
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
            Self::ResourceLimit(value) => formatter
                .debug_tuple("Error::ResourceLimit")
                .field(value)
                .finish(),
            Self::AllocationFailed(context) => formatter
                .debug_tuple("Error::AllocationFailed")
                .field(context)
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
            Self::EncryptionPolicy => formatter.write_str("Error::EncryptionPolicy"),
        }
    }
}

impl From<io::Error> for Error {
    fn from(err: io::Error) -> Self {
        Error::Io(err)
    }
}

impl From<WriteError> for Error {
    fn from(error: WriteError) -> Self {
        Self::Authoring(error)
    }
}

impl From<OleError> for Error {
    fn from(err: OleError) -> Self {
        Error::Ole(err)
    }
}

impl From<litchi_ole_common::Error> for Error {
    fn from(error: litchi_ole_common::Error) -> Self {
        Self::Object(error)
    }
}

impl From<litchi_ole_common::custom_xml::Error> for Error {
    fn from(error: litchi_ole_common::custom_xml::Error) -> Self {
        Self::CustomXml(error)
    }
}

#[cfg(feature = "sign")]
impl From<litchi_sign::Error> for Error {
    fn from(error: litchi_sign::Error) -> Self {
        Self::Signature(error)
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
            Error::Authoring(_) => f.write_str("PPT authoring failed"),
            Error::Io(e) => write!(f, "IO error: {e}"),
            Error::Ole(e) => write!(f, "OLE error: {e}"),
            Error::Object(e) => write!(f, "OLE object error: {e}"),
            Error::CustomXml(_) => f.write_str("PPT Custom XML operation failed"),
            #[cfg(feature = "sign")]
            Error::Signature(_) => f.write_str("PPT signature operation failed"),
            Error::OfficeArt(e) => write!(f, "OfficeArt error: {e}"),
            Error::Graph(e) => write!(f, "Office Graph error: {e}"),
            Error::InvalidFormat(s) => write!(f, "Invalid format: {s}"),
            Error::StreamNotFound(s) => write!(f, "Stream not found: {s}"),
            Error::Corrupted(s) => write!(f, "Corrupted file: {s}"),
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => write!(
                f,
                "invalid PPT source limit for {kind}: {value} (maximum {maximum})"
            ),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                f,
                "PPT source limit exceeded for {kind}: {observed} > {maximum}"
            ),
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
            Error::EncryptionPolicy => write!(
                f,
                "ordinary PPT publication is refused for a password-to-open encrypted source"
            ),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Authoring(error) => Some(error),
            Self::Io(error) => Some(error),
            Self::Ole(error) => Some(error),
            Self::Object(error) => Some(error),
            Self::CustomXml(error) => Some(error),
            #[cfg(feature = "sign")]
            Self::Signature(error) => Some(error),
            Self::OfficeArt(error) => Some(error),
            Self::Graph(error) => Some(error),
            Self::InvalidFormat(_)
            | Self::StreamNotFound(_)
            | Self::Corrupted(_)
            | Self::InvalidLimit { .. }
            | Self::Limit { .. }
            | Self::ResourceLimit(_)
            | Self::AllocationFailed(_)
            | Self::PasswordRequired
            | Self::InvalidPassword
            | Self::UnsupportedEncryption(_)
            | Self::MalformedEncryptionHeader(_)
            | Self::EncryptionPolicy => None,
        }
    }
}

/// Result type for PPT operations.
pub type Result<T> = std::result::Result<T, Error>;

/// The low-level `PowerPoint` (.ppt) reader.
///
/// This reader is the eager, generic `Read + Seek` entry point for working with
/// legacy `PowerPoint` presentations. The source-preserving package owner is
/// exposed separately; callers that need this mutable native projection use
/// [`crate::raw::Reader`].
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ppt::raw::Reader;
///
/// // Open an existing presentation
/// let mut pkg = Reader::open("presentation.ppt")?;
///
/// // Get the main presentation
/// let pres = pkg.presentation()?;
///
/// // Extract text
/// let text = pres.text()?;
/// println!("{}", text);
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Reader<R: Read + Seek = File> {
    /// The underlying OLE file
    pub(super) ole: OleFile<R>,
    /// Limits inherited by [`Reader::presentation`](super::Reader::presentation).
    pub(super) source_limits: SourceLimits,
    /// VBA limits inherited by [`Reader::presentation`](super::Reader::presentation).
    #[cfg(feature = "vba-inspection")]
    pub(super) vba_project_limits: crate::VbaProjectLimits,
    /// Leniency inherited by [`Reader::presentation`](super::Reader::presentation).
    pub(super) leniency: crate::leniency::Leniency,
}

// `From<Error> for litchi_core::Error` lives here (not in the umbrella) so
// the orphan rule is satisfied — both source and target crates are external
// to the umbrella.
// The generic core projection has no foreign-source carrier. Keep this
// projection fixed and content-free; litchi-ops classifies the original PPT
// error before this conversion.
fn core_custom_xml_error(error: litchi_ole_common::custom_xml::Error) -> litchi_core::Error {
    match error {
        litchi_ole_common::custom_xml::Error::Ole(error) => {
            litchi_core::Error::from(litchi_ole_common::Error::from(error))
        },
        litchi_ole_common::custom_xml::Error::Invalid(_)
        | litchi_ole_common::custom_xml::Error::ResourceLimit(_)
        | litchi_ole_common::custom_xml::Error::Xml(_) => {
            litchi_core::Error::CorruptedFile("PPT Custom XML operation failed".to_string())
        },
        _ => litchi_core::Error::CorruptedFile("PPT Custom XML operation failed".to_string()),
    }
}

#[cfg(feature = "sign")]
fn core_signature_error(error: litchi_sign::Error) -> litchi_core::Error {
    match error {
        litchi_sign::Error::Cfb(error) => litchi_core::Error::from(error),
        litchi_sign::Error::Encrypted
        | litchi_sign::Error::Legacy
        | litchi_sign::Error::Unsupported(_) => {
            litchi_core::Error::Unsupported("PPT signature policy is unsupported".to_string())
        },
        litchi_sign::Error::Xml(_)
        | litchi_sign::Error::Limit(_)
        | litchi_sign::Error::Sha1
        | litchi_sign::Error::Key(_)
        | litchi_sign::Error::Sign(_)
        | litchi_sign::Error::Container(_) => {
            litchi_core::Error::CorruptedFile("PPT signature operation failed".to_string())
        },
        _ => litchi_core::Error::CorruptedFile("PPT signature operation failed".to_string()),
    }
}

impl From<Error> for litchi_core::Error {
    fn from(err: Error) -> Self {
        match err {
            Error::Io(e) => litchi_core::Error::Io(e),
            Error::Ole(ole_err) => litchi_core::Error::from(ole_err),
            Error::Object(error) => litchi_core::Error::from(error),
            Error::CustomXml(error) => core_custom_xml_error(error),
            #[cfg(feature = "sign")]
            Error::Signature(error) => core_signature_error(error),
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
            Error::InvalidLimit {
                kind,
                value,
                maximum,
            } => litchi_core::Error::InvalidFormat(format!(
                "invalid PPT source limit for {kind}: {value} (maximum {maximum})"
            )),
            Error::Limit {
                kind,
                observed,
                maximum,
            } => litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: match kind {
                    SourceLimitKind::Events => litchi_core::Resource::Objects,
                    SourceLimitKind::Depth => litchi_core::Resource::Depth,
                    _ => litchi_core::Resource::InputBytes,
                },
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                scope: Arc::from(format!("PPT {kind}")),
            }),
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
            Error::EncryptionPolicy => litchi_core::Error::Unsupported(
                "ordinary PPT publication cannot preserve password-to-open encryption".into(),
            ),
            Error::Authoring(_) => litchi_core::Error::Unsupported("PPT authoring failed".into()),
        }
    }
}

const fn min_usize(left: usize, right: usize) -> usize {
    if left < right { left } else { right }
}
