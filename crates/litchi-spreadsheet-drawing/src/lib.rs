//! Shared `SpreadsheetML` drawing primitives.

#![forbid(unsafe_code)]

pub mod chart;

use thiserror::Error;

/// Result of a shared `SpreadsheetML` drawing operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure while reading, validating, or writing shared `SpreadsheetML` drawing data.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The supplied chart structure violates an OOXML invariant.
    #[error("invalid SpreadsheetML drawing structure: {0}")]
    Invalid(String),
    /// XML or byte encoding could not be decoded or emitted.
    #[error("SpreadsheetML drawing encoding error: {0}")]
    Encoding(String),
    /// `DrawingML` chart decoding or validation failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),
}
