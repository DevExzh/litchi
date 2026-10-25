use std::collections::TryReserveError;
use thiserror::Error;

/// Result returned by canonical DOCX operations.
pub type Result<T> = std::result::Result<T, Error>;

/// A bounded parsing, validation, or package-graph failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The owning ingress route cannot retain a positional-source policy.
    #[error("DOCX owning ingress cannot retain {policy:?}; use positional-source ingress")]
    UnsupportedOpenPolicy {
        /// The resource policy that must not be silently discarded.
        policy: crate::OpenPolicy,
    },

    /// An output would discard encryption provenance or assumes a missing mode.
    #[cfg(feature = "encryption")]
    #[error("DOCX encryption policy rejected {operation}: {source}")]
    EncryptionPolicy {
        /// Operation rejected at the DOCX boundary.
        operation: &'static str,
        /// Shared OOXML provenance policy failure.
        #[source]
        source: litchi_ooxml_common::package_encryption::PolicyError,
    },

    /// The underlying OPC graph is malformed or could not be mutated safely.
    #[error("OPC error: {0}")]
    Opc(#[from] litchi_opc::error::OpcError),

    /// A source-backed artifact stream failed while preserving the producer's
    /// typed error and any accepted publication prefix.
    #[error(transparent)]
    ArtifactStream(Box<litchi_opc::ArtifactStreamError<Error>>),

    /// A requested OPC part is absent.
    #[error("DOCX part not found: {0}")]
    PartNotFound(String),

    /// XML syntax or encoding is invalid.
    #[error("invalid DOCX XML: {0}")]
    Xml(String),

    /// A part has a content type forbidden by the `WordprocessingML` relation.
    #[error("invalid DOCX content type: expected {expected}, got {actual}")]
    ContentType { expected: String, actual: String },

    /// A package part has an invalid content type for its relationship.
    #[error("invalid DOCX content type: expected {expected}, got {got}")]
    InvalidContentType { expected: String, got: String },

    /// A package relationship is malformed or inconsistent.
    #[error("invalid DOCX relationship: {0}")]
    InvalidRelationship(String),

    /// Parsed or requested data violates a `WordprocessingML` invariant.
    #[error("invalid DOCX data: {0}")]
    Invalid(String),

    /// A host-level DOCX operation rejected the package as invalid.
    #[error("invalid DOCX format: {0}")]
    InvalidFormat(String),

    /// `DrawingML` parsing or authoring failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),

    /// A DOCX InkML inventory exhausted a finite resource budget.
    #[error("DOCX ink {kind} limit exceeded: {actual} > {maximum}")]
    InkLimit {
        /// The exceeded budget (record 0355a).
        kind: crate::ink::LimitKind,
        actual: usize,
        maximum: usize,
    },

    /// Bounded, inert VBA parsing or authoring failed.
    #[cfg(feature = "vba-inspection")]
    #[error("VBA error: {0}")]
    Vba(#[from] litchi_vba::Error),

    /// Runtime-neutral OOXML encryption failed.
    #[cfg(feature = "encryption")]
    #[error("OOXML encryption error: {0}")]
    Encryption(#[from] litchi_crypto::ooxml::Error),

    /// A shared host-neutral OOXML service failed.
    #[error("shared OOXML error: {0}")]
    Common(#[source] litchi_ooxml_common::Error),

    /// File or stream I/O failed.
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    /// A well-formed request this route refuses (record 0351): the kind
    /// classifies it, the operation and reason are static text.
    #[error("DOCX {0}")]
    Refused(litchi_core::edit::Refusal),

    /// A patch or staged plan was made against another base (record 0351).
    #[error("DOCX {operation} patch conflict: {reason}")]
    PatchConflict {
        /// Operation whose base did not match.
        operation: &'static str,
        /// What did not match.
        reason: &'static str,
    },

    /// A changed package still carries signature infrastructure.
    #[error("DOCX signed package requires an explicit signature disposition")]
    Signed,

    /// A source-backed package changed while the operation was in progress.
    #[error("DOCX source changed while the operation was in progress")]
    SourceChanged {
        /// Version captured before the operation.
        expected: litchi_core::SourceVersion,
        /// Version observed when the operation completed.
        observed: litchi_core::SourceVersion,
    },

    /// A broken internal invariant: state the crate itself should have
    /// provided is missing or inconsistent.
    #[error("DOCX {operation} internal invariant failed: {reason}")]
    Internal {
        /// Operation that detected the broken invariant.
        operation: &'static str,
        /// The invariant that failed.
        reason: &'static str,
    },

    /// A generic bounded DOCX operation failed.
    #[error("{0}")]
    Other(String),

    /// Writing a committed paragraph-stream package failed (record 0355a):
    /// the typed stream error, carried through the package writer.
    #[error(transparent)]
    ParagraphStream(Box<crate::source_backed::tail_append_stream::Error>),

    /// A committed plain-paragraph tail append failed while its deferred
    /// source splice was being written.  Keep the operation's typed refusal,
    /// execution, limit, I/O, and incomplete-output causes intact.
    #[error(transparent)]
    TailAppend(Box<crate::source_backed::tail_append::Error>),

    /// A semantic selector matched more than one producer object.
    #[error("{object} selector '{key}' is ambiguous")]
    Ambiguous {
        /// Kind of object being selected.
        object: &'static str,
        /// User-facing semantic selector.
        key: String,
    },

    /// A checked numeric selector was outside the current collection.
    #[error("{object} index {index} is out of bounds for length {len}")]
    OutOfBounds {
        /// Kind of object being selected.
        object: &'static str,
        /// Requested zero-based index.
        index: usize,
        /// Collection length at validation time.
        len: usize,
    },

    /// An OPC part URI is invalid.
    #[error("invalid DOCX part URI: {0}")]
    Uri(String),

    /// An invalid URI was supplied to a package operation.
    #[error("invalid DOCX URI: {0}")]
    InvalidUri(String),

    /// Markup-compatibility preprocessing failed.
    #[error("DOCX markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::mce::Error),

    /// A bounded authoring operation could not reserve its planned buffer.
    #[error("DOCX allocation failed for {resource}: {source}")]
    Allocation {
        /// Buffer or collection being reserved.
        resource: &'static str,
        /// Original allocator failure.
        #[source]
        source: TryReserveError,
    },

    /// An alternative-format anchor budget is exhausted (record 0355a).
    #[error("DOCX alternative-format {kind} limit exceeded: {observed} > {maximum}")]
    AltChunkLimit {
        /// Bounded resource.
        kind: crate::alt::LimitKind,
        /// Observed or requested value.
        observed: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A glossary budget is exhausted (record 0355a).
    #[error("DOCX glossary {kind} limit exceeded: {actual} > {maximum}")]
    GlossaryLimit {
        /// The exceeded budget.
        kind: crate::glossary::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A content-control budget is exhausted (record 0355a).
    #[error("DOCX content-control {kind} limit exceeded: {actual} > {maximum}")]
    ContentControlLimit {
        /// The exceeded budget.
        kind: crate::content_control::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A conflict-markup budget is exhausted (record 0355a).
    #[error("DOCX conflict-markup {kind} limit exceeded: {actual} > {maximum}")]
    ConflictLimit {
        /// The exceeded budget.
        kind: crate::revision::conflict::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A chart budget is exhausted (record 0355a).
    #[error("DOCX chart {kind} limit exceeded: {actual} > {maximum}")]
    ChartLimit {
        /// The exceeded budget.
        kind: crate::chart::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A font-table budget is exhausted (record 0355a).
    #[error("DOCX font-table {kind} limit exceeded: {actual} > {maximum}")]
    FontLimit {
        /// The exceeded budget.
        kind: crate::font::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A bibliography-source budget is exhausted (record 0355a).
    #[error("DOCX bibliography-source {kind} limit exceeded: {actual} > {maximum}")]
    BibliographyLimit {
        /// The exceeded budget.
        kind: crate::bibliography::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A web-settings budget is exhausted (record 0355a).
    #[error("DOCX web-settings {kind} limit exceeded: {actual} > {maximum}")]
    WebLimit {
        /// The exceeded budget.
        kind: crate::web::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A story budget is exhausted (record 0355a).
    #[error("DOCX story {kind} limit exceeded: {actual} > {maximum}")]
    StoryLimit {
        /// The exceeded budget.
        kind: crate::package::story::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A numbering budget is exhausted (record 0355a).
    #[error("DOCX numbering {kind} limit exceeded: {actual} > {maximum}")]
    NumberingLimit {
        /// The exceeded budget.
        kind: crate::numbering::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A mail-merge budget is exhausted (record 0355a).
    #[error("DOCX mail-merge {kind} limit exceeded: {actual} > {maximum}")]
    MailMergeLimit {
        /// The exceeded budget.
        kind: crate::mail_merge::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A modern Word comments raw-codec budget is exhausted.
    #[error("DOCX modern comments {kind} limit exceeded: {actual} > {maximum}")]
    ModernCommentLimit {
        /// The exceeded budget.
        kind: crate::modern_comments::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A stylesWithEffects budget is exhausted (record 0355a).
    #[error("DOCX stylesWithEffects {kind} limit exceeded: {actual} > {maximum}")]
    StylesWithEffectsLimit {
        /// The exceeded budget.
        kind: crate::styles::effects::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A settings budget is exhausted (record 0355a).
    #[error("DOCX settings {kind} limit exceeded: {actual} > {maximum}")]
    SettingsLimit {
        /// The exceeded budget.
        kind: crate::settings::LimitKind,
        /// Observed or requested value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A block read exceeded a finite resource ceiling.
    #[error("DOCX block {resource} limit exceeded: {actual} > {maximum}")]
    BlockReadLimit {
        /// Bounded resource.
        resource: &'static str,
        /// Observed value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A tracked-revision read exceeded a finite resource ceiling.
    #[error("DOCX revision {kind} limit exceeded: {actual} > {maximum}")]
    RevisionLimit {
        /// The exceeded budget (record 0355a).
        kind: crate::revision::LimitKind,
        /// Observed value.
        actual: usize,
        /// Maximum accepted value.
        maximum: usize,
    },

    /// A bounded revision collection could not reserve storage.
    #[error("DOCX allocation failed for {resource}")]
    RevisionAllocation {
        /// Buffer or collection being reserved.
        resource: &'static str,
    },

    /// A bounded external hyperlink-wrapper detachment ceiling was exceeded.
    #[error("DOCX external hyperlink detachment {kind} limit exceeded: {actual} > {maximum}")]
    ExternalHyperlinkDetachmentLimit {
        /// Bounded resource.
        kind: crate::sanitize::LimitKind,
        /// Maximum accepted value.
        maximum: usize,
        /// Observed value.
        actual: usize,
    },

    /// A bounded irreversible external-hyperlink redaction ceiling was exceeded.
    #[error("DOCX external hyperlink redaction {kind} limit exceeded: {actual} > {maximum}")]
    ExternalHyperlinkRedactionLimit {
        /// Bounded resource.
        kind: crate::redact::LimitKind,
        /// Maximum accepted value.
        maximum: usize,
        /// Observed value.
        actual: usize,
    },

    /// A bounded main-document section inventory ceiling was exceeded.
    #[error("DOCX section inventory {kind} limit exceeded: {actual} > {maximum}")]
    SectionInventoryLimit {
        /// The exceeded budget (record 0355a).
        kind: crate::section::LimitKind,
        /// Maximum accepted value.
        maximum: usize,
        /// Observed value.
        actual: usize,
    },
}

impl Error {
    /// A refusal of `operation` for `reason`, classified by `kind`.
    pub(crate) const fn refused(
        kind: litchi_core::edit::RefusalKind,
        operation: &'static str,
        reason: &'static str,
    ) -> Self {
        Self::Refused(litchi_core::edit::Refusal::new(kind, operation, reason))
    }

    /// A patch or plan applied to another base.
    pub(crate) const fn conflict(operation: &'static str, reason: &'static str) -> Self {
        Self::PatchConflict { operation, reason }
    }

    /// A broken internal invariant.
    pub(crate) const fn internal(operation: &'static str, reason: &'static str) -> Self {
        Self::Internal { operation, reason }
    }

    /// The refusal this error carries, when it is one.
    #[must_use]
    pub const fn refusal(&self) -> Option<&litchi_core::edit::Refusal> {
        match self {
            Self::Refused(refusal) => Some(refusal),
            _ => None,
        }
    }
}

impl From<litchi_opc::ArtifactStreamError<Error>> for Error {
    fn from(error: litchi_opc::ArtifactStreamError<Error>) -> Self {
        Self::ArtifactStream(Box::new(error))
    }
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}

impl From<std::fmt::Error> for Error {
    fn from(error: std::fmt::Error) -> Self {
        Self::Xml(error.to_string())
    }
}

impl From<litchi_ooxml_common::XmlError> for Error {
    fn from(error: litchi_ooxml_common::XmlError) -> Self {
        Self::Common(litchi_ooxml_common::Error::Decode(error))
    }
}

impl From<litchi_ooxml_common::Error> for Error {
    fn from(error: litchi_ooxml_common::Error) -> Self {
        match error {
            litchi_ooxml_common::Error::Refused(refusal) => Self::Refused(refusal),
            litchi_ooxml_common::Error::PatchConflict => {
                Self::conflict("shared OOXML", "the selected package graph changed")
            },
            litchi_ooxml_common::Error::Signed => Self::Signed,
            litchi_ooxml_common::Error::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            litchi_ooxml_common::Error::Internal(reason) => Self::internal("shared OOXML", reason),
            other => Self::Common(other),
        }
    }
}

impl From<litchi_ooxml_common::properties::Error> for Error {
    fn from(error: litchi_ooxml_common::properties::Error) -> Self {
        match error {
            litchi_ooxml_common::properties::Error::Refused(refusal) => Self::Refused(refusal),
            litchi_ooxml_common::properties::Error::PatchConflict => {
                Self::conflict("metadata", "the selected property graph changed")
            },
            litchi_ooxml_common::properties::Error::Signed => Self::Signed,
            litchi_ooxml_common::properties::Error::SourceChanged { expected, observed } => {
                Self::SourceChanged { expected, observed }
            },
            litchi_ooxml_common::properties::Error::Internal(reason) => {
                Self::internal("metadata", reason)
            },
            litchi_ooxml_common::properties::Error::Invalid(message) => Self::Invalid(message),
            litchi_ooxml_common::properties::Error::Limit {
                kind,
                observed,
                maximum,
            } => Self::from(litchi_ooxml_common::Error::Limit {
                kind: properties_limit_kind(kind),
                observed,
                maximum,
            }),
            litchi_ooxml_common::properties::Error::Package(error) => Self::from(error),
            other => Self::Other(other.to_string()),
        }
    }
}

fn properties_limit_kind(
    kind: litchi_ooxml_common::properties::LimitKind,
) -> litchi_ooxml_common::LimitKind {
    litchi_ooxml_common::LimitKind::Properties(kind)
}
