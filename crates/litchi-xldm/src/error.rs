//! Errors and invariant helpers for the neutral MS-XLDM owner.

use std::collections::TryReserveError;

use thiserror::Error;

/// A bounded MS-XLDM inspection or exact-source publication failure.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The source violates an MS-XLDM structural or profile invariant.
    #[error("invalid MS-XLDM structure: {0}")]
    Invalid(String),
    /// A selected XLDM raw budget was exceeded.
    #[error("MS-XLDM {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The typed resource dimension that was exceeded.
        kind: LimitKind,
        /// The first observed amount beyond the selected policy.
        observed: usize,
        /// The active upper bound.
        maximum: usize,
    },
    /// A caller supplied a raw budget outside the supported hard ceiling.
    #[error("invalid MS-XLDM {kind} limit {value} (allowed 1..={maximum})")]
    InvalidLimit {
        /// The typed resource dimension being configured.
        kind: LimitKind,
        /// The rejected caller value.
        value: usize,
        /// The format hard ceiling.
        maximum: usize,
    },
    /// The selected MS-XLDM profile is outside this owner’s supported contract.
    #[error("unsupported MS-XLDM operation: {feature}")]
    Unsupported { feature: &'static str },
    /// A bounded operation could not reserve its required memory.
    #[error("could not reserve memory for MS-XLDM {resource}: {source}")]
    Allocation {
        resource: &'static str,
        #[source]
        source: TryReserveError,
    },
    /// XML tokenization or decoding failed before the source was projected.
    #[error("MS-XLDM XML error: {0}")]
    Xml(String),
}

/// Result type for the neutral MS-XLDM owner.
pub type Result<T> = std::result::Result<T, Error>;

pub(crate) fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

pub(crate) const fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind,
        observed,
        maximum,
    }
}

#[cold]
#[track_caller]
pub(crate) fn panic_missing_invariant(message: &str) -> ! {
    panic!("MS-XLDM internal invariant failed: {message}")
}

#[cold]
#[track_caller]
pub(crate) fn panic_error_invariant(message: &str, error: impl std::fmt::Display) -> ! {
    panic!("MS-XLDM internal invariant failed: {message}: {error}")
}

/// A bounded dimension of an MS-XLDM proof or rewrite, as reported by
/// [`crate::OlapProofError::Limit`] and
/// [`crate::Xldm140RenameError::limit_kind`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes of rewritten storage or metadata payloads.
    OutputBytes,
    /// UTF-8 bytes of a table XML name.
    PathTextBytes,
    /// Visited records and XML nodes ([`crate::OlapProofLimits::max_items`]).
    Items,
    /// UTF-8 bytes examined and retained
    /// ([`crate::OlapProofLimits::max_string_bytes`]).
    StringBytes,
    /// The inspected source size ([`crate::OlapProofLimits::max_source_bytes`]).
    SourceBytes,
    /// Graph work units ([`crate::OlapProofLimits::max_work`]).
    Work,
    /// The inspected outer storage source bytes.
    XldmSourceBytes,
    /// The outer or metadata output bytes.
    XldmOutputBytes,
    /// The virtual-directory XML bytes.
    DirectoryBytes,
    /// The virtual-directory and logged-file count.
    Files,
    /// The partition count.
    Partitions,
    /// One storage path or scalar leaf's UTF-8 bytes.
    PathBytes,
    /// Outer XML node count.
    XmlNodes,
    /// Outer XML nesting depth.
    XmlDepth,
    /// Outer XML decoded text bytes.
    XmlTextBytes,
    /// Section-2.5 metadata XML bytes.
    MetadataBytes,
    /// Section-2.5 metadata XML node count.
    MetadataNodes,
    /// Section-2.5 metadata XML nesting depth.
    MetadataDepth,
    /// Section-2.5 metadata decoded text bytes.
    MetadataTextBytes,
    /// Section-2.5 metadata file count.
    MetadataFiles,
}

impl LimitKind {
    /// A short structural description of the dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::OutputBytes => "XLDM rewrite output bytes",
            Self::PathTextBytes => "XLDM path text bytes",
            Self::Items => "OLAP proof items",
            Self::StringBytes => "OLAP proof string bytes",
            Self::SourceBytes => "OLAP proof source bytes",
            Self::Work => "OLAP proof graph work",
            Self::XldmSourceBytes => "XLDM source bytes",
            Self::XldmOutputBytes => "XLDM output bytes",
            Self::DirectoryBytes => "XLDM directory bytes",
            Self::Files => "XLDM files",
            Self::Partitions => "XLDM partitions",
            Self::PathBytes => "XLDM path bytes",
            Self::XmlNodes => "XLDM XML nodes",
            Self::XmlDepth => "XLDM XML depth",
            Self::XmlTextBytes => "XLDM XML text bytes",
            Self::MetadataBytes => "XLDM metadata bytes",
            Self::MetadataNodes => "XLDM metadata XML nodes",
            Self::MetadataDepth => "XLDM metadata XML depth",
            Self::MetadataTextBytes => "XLDM metadata text bytes",
            Self::MetadataFiles => "XLDM metadata files",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
