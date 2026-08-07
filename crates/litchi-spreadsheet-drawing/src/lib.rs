//! Shared `SpreadsheetML` drawing primitives.

#![forbid(unsafe_code)]

pub mod chart;
pub mod shape;

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
    /// Markup-compatibility preprocessing failed while reading drawing XML.
    #[error("SpreadsheetML drawing markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::mce::Error),
    /// Shared OOXML attribute or character-reference decoding failed.
    #[error("SpreadsheetML drawing XML error: {0}")]
    Xml(#[from] litchi_ooxml_common::XmlError),
    /// `DrawingML` chart decoding or validation failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),
}
