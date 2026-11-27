use thiserror::Error;

/// Resource dimension governed by Graph ingress, authoring or edit bounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete compound-file bytes.
    PackageBytes,
    /// Direct children of the compound-file root.
    RootEntries,
    /// Bytes in one root stream.
    StreamBytes,
    /// Bytes in the Workbook stream.
    WorkbookBytes,
    /// Chart substreams in a workbook.
    Charts,
    /// Records in one chart substream.
    ChartRecords,
    /// Semantic series in one chart.
    Series,
    /// Chart groups in one chart.
    Groups,
    /// Axes in one chart.
    Axes,
    /// Bytes in one inert formula token array.
    FormulaBytes,
    /// Cached values in one chart.
    CachedValues,
    /// Aggregate retained unknown-record bytes.
    UnknownBytes,
    /// Nested Begin/End chart collections.
    Nesting,
    /// Distinct staged cache replacements.
    CacheEdits,
    /// Distinct staged series-metadata replacements.
    SeriesMetadataEdits,
    /// A physical BIFF framing bound.
    Biff(litchi_biff::Resource),
}

impl LimitKind {
    pub(crate) const fn description(self) -> &'static str {
        match self {
            Self::PackageBytes => "package bytes",
            Self::RootEntries => "root entries",
            Self::StreamBytes => "stream bytes",
            Self::WorkbookBytes => "Workbook bytes",
            Self::Charts => "chart count",
            Self::ChartRecords => "chart record count",
            Self::Series => "series count",
            Self::Groups => "group count",
            Self::Axes => "axis count",
            Self::FormulaBytes => "formula bytes",
            Self::CachedValues => "cached value count",
            Self::UnknownBytes => "unknown chart bytes",
            Self::Nesting => "chart nesting",
            Self::CacheEdits => "cache edit count",
            Self::SeriesMetadataEdits => "series metadata edit count",
            Self::Biff(_) => "BIFF framing",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Biff(resource) => resource.fmt(formatter),
            kind => formatter.write_str(kind.description()),
        }
    }
}

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
    Biff(#[source] litchi_biff::Error),

    /// A configured bound is internally inconsistent or unsupported.
    #[error("invalid {kind} limit {value}; maximum is {maximum}")]
    InvalidLimit {
        /// Dimension whose configured value is invalid.
        kind: LimitKind,
        /// Rejected bound.
        value: u64,
        /// Largest supported bound; configured Graph bounds must also be nonzero.
        maximum: u64,
    },

    /// Input or output crossed a configured resource bound.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Dimension whose configured bound was crossed.
        kind: LimitKind,
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

    /// A well-formed request this route refuses (a parsed chart whose
    /// opaque records cannot be re-encoded losslessly, a fresh value outside
    /// the authoring profile, a mutation whose host ownership is not
    /// modeled), classified by [`litchi_core::edit::RefusalKind`]. The
    /// refusal is content-free: its operation and reason are static.
    #[error("{0}")]
    Refused(litchi_core::edit::Refusal),

    /// A patch or staged edit was made against another source than the one
    /// receiving it.
    #[error("{operation} patch conflict: {reason}")]
    PatchConflict {
        /// Concise operation name.
        operation: &'static str,
        /// Static explanation of the mismatch.
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

impl From<litchi_biff::Error> for Error {
    // Record loops reach this conversion only on failure. Keep the budget
    // normalization out of their successful per-record instruction paths.
    #[cold]
    #[inline(never)]
    fn from(error: litchi_biff::Error) -> Self {
        match error {
            litchi_biff::Error::LimitExceeded {
                resource,
                observed,
                maximum,
            } => Self::Limit {
                kind: LimitKind::Biff(resource),
                observed,
                maximum,
            },
            litchi_biff::Error::InvalidLimit {
                resource,
                value,
                maximum,
            } => Self::InvalidLimit {
                kind: LimitKind::Biff(resource),
                value,
                maximum,
            },
            other => Self::Biff(other),
        }
    }
}
