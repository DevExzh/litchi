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
    /// A slide-audio position contained a non-finite coordinate.
    #[error("Keynote slide-audio position must have finite coordinates")]
    InvalidAudioPosition,
    /// A slide-audio duration was zero or outside finite `f32` seconds.
    #[error("Keynote slide-audio duration must be positive and fit in finite f32 seconds")]
    InvalidAudioDuration,
    /// A transition delay was not finite and non-negative.
    #[error("Keynote transition delay must be finite and non-negative")]
    InvalidDelay,
    /// A transition-specific floating-point value was not finite.
    #[error("Keynote transition custom values must be finite")]
    InvalidCustomFloat,
    /// A transition animation detail value was not finite.
    #[error("Keynote transition detail must be finite")]
    InvalidDetail,
    /// A known native enum value was represented by its lossless unknown form.
    #[error("Keynote mode must use its canonical variant for a known native value")]
    NonCanonicalMode,
    /// A known transition effect identifier was represented by its unknown form.
    #[error("Keynote transition effect must use its canonical variant for a known identifier")]
    NonCanonicalEffect,
    /// A semantic transition string contained a NUL byte.
    #[error("Keynote transition strings cannot contain NUL")]
    NulString,
    /// A transition identifier exceeded the bounded semantic storage budget.
    #[error("Keynote transition identifier exceeds the semantic byte budget")]
    IdentifierTooLarge,
    /// An opaque transition payload exceeded the bounded semantic storage budget.
    #[error("Keynote transition opaque payload exceeds the semantic byte budget")]
    PayloadTooLarge,
    /// A native identifier required for lossless decoding was empty.
    #[error("Keynote animation identifier cannot be empty")]
    EmptyIdentifier,
    /// An opaque background payload cannot be empty.
    #[error("Keynote opaque background payload cannot be empty")]
    EmptyBackgroundPayload,
}

/// Result type for semantic Keynote operations.
pub type Result<T> = std::result::Result<T, Error>;
