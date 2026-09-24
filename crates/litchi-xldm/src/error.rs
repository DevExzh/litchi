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
