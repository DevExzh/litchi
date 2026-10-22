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
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}
