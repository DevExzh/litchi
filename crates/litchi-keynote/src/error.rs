//! Errors owned by the archive-free Keynote semantic layer.

use thiserror::Error;

/// A fallible semantic-model operation failed validation.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A presentation or slide size was not finite and strictly positive.
    #[error("Keynote dimensions must be finite and greater than zero")]
    InvalidDimensions,
    /// A duration was not finite and non-negative.
    #[error("Keynote duration must be finite and non-negative")]
    InvalidDuration,
    /// A known native enum value was represented by its lossless unknown form.
    #[error("Keynote mode must use its canonical variant for a known native value")]
    NonCanonicalMode,
    /// A native identifier required for lossless decoding was empty.
    #[error("Keynote animation identifier cannot be empty")]
    EmptyIdentifier,
    /// An opaque background payload cannot be empty.
    #[error("Keynote opaque background payload cannot be empty")]
    EmptyBackgroundPayload,
}

/// Result type for semantic Keynote operations.
pub type Result<T> = std::result::Result<T, Error>;
