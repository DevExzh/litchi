use thiserror::Error;

/// Result type used by this crate.
pub type Result<T> = std::result::Result<T, Error>;

/// Checked `OGraph` parsing and encoding failures.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying compound-file container is malformed.
    #[error("invalid compound file: {0}")]
    Cfb(#[from] litchi_cfb::OleError),

    /// The shared BIFF framing layer rejected the physical record stream.
    #[error("invalid BIFF framing: {0}")]
    Biff(#[from] litchi_biff::Error),

    /// A configured bound is internally inconsistent or unsupported.
    #[error("invalid {resource} limit {value}: {reason}")]
    InvalidLimit {
        /// Name of the bounded resource.
        resource: &'static str,
        /// Rejected bound.
        value: u64,
        /// Static explanation of the constraint.
        reason: &'static str,
    },

    /// Input or output crossed a configured resource bound.
    #[error("{resource} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Name of the bounded resource.
        resource: &'static str,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },

    /// A typed decoder received a different record identifier.
    #[error("expected BIFF record {expected:#06X}, found {actual:#06X}")]
    WrongRecord {
        /// Expected BIFF identifier.
        expected: u16,
        /// Actual BIFF identifier.
        actual: u16,
    },

    /// A typed BIFF record has an invalid fixed payload size.
    #[error("BIFF record {kind:#06X} has invalid length: expected {expected}, found {actual}")]
    InvalidRecordLength {
        /// BIFF record identifier.
        kind: u16,
        /// Required length.
        expected: usize,
        /// Actual length.
        actual: usize,
    },

    /// A scalar value is outside the range defined by the specification.
    #[error("BIFF record {kind:#06X} has invalid {field} value {value:#X}")]
    InvalidRecordValue {
        /// BIFF record identifier.
        kind: u16,
        /// Field name.
        field: &'static str,
        /// Rejected raw value.
        value: u64,
    },

    /// A required root stream is absent.
    #[error("missing required root stream {name:?}")]
    MissingStream {
        /// Required stream name.
        name: &'static str,
    },

    /// A root stream occurs more than once.
    #[error("duplicate root stream {name:?}")]
    DuplicateStream {
        /// Duplicated stream name.
        name: String,
    },

    /// Standalone `OGraph` packages reject unknown streams and all storages.
    #[error("unexpected root entry {name:?} with type {entry_type}")]
    UnexpectedEntry {
        /// Directory entry name.
        name: String,
        /// Raw CFB directory entry type.
        entry_type: u8,
    },

    /// The Workbook stream does not have the standalone `OGraph` substream shape.
    #[error("invalid standalone OGraph Workbook at offset {offset}: {reason}")]
    InvalidWorkbook {
        /// Record offset nearest the failure.
        offset: usize,
        /// Static structural explanation.
        reason: &'static str,
    },

    /// A BIFF chart substream is malformed or structurally incomplete.
    #[error("invalid chart substream at offset {offset}: {reason}")]
    InvalidChart {
        /// Record offset nearest the failure.
        offset: usize,
        /// Static structural explanation.
        reason: &'static str,
    },

    /// A semantic chart value cannot be represented safely.
    #[error("invalid chart {field}: {reason}")]
    InvalidModel {
        /// Semantic field nearest the failure.
        field: &'static str,
        /// Static validation explanation.
        reason: &'static str,
    },

    /// Re-encoding a parsed chart would risk losing opaque source data.
    #[error("unsafe parsed-chart edit refused: {reason}")]
    UnsafeEdit {
        /// Why a byte-preserving rewrite cannot be guaranteed.
        reason: &'static str,
    },

    /// Fresh semantic authoring is not yet backed by a complete wire grammar.
    #[error("unsupported chart authoring: {reason}")]
    UnsupportedAuthoring {
        /// Static explanation of the missing proof boundary.
        reason: &'static str,
    },

    /// A mutation is well-formed at the API boundary but its complete host
    /// ownership graph is not yet modeled safely.
    #[error("unsupported chart mutation {operation}: {reason}")]
    UnsupportedMutation {
        /// Concise operation name.
        operation: &'static str,
        /// Static explanation of the missing proof boundary.
        reason: &'static str,
    },

    /// Checked length arithmetic could not be represented.
    #[error("size overflow while processing {resource}")]
    SizeOverflow {
        /// Resource being sized.
        resource: &'static str,
    },

    /// A fallible allocation could not reserve the requested capacity.
    #[error("could not allocate storage for {resource}")]
    Allocation {
        /// Resource being allocated.
        resource: &'static str,
    },
}
