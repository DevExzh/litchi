//! Shared `SpreadsheetML` drawing primitives.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

pub mod chart;
pub mod objects;
pub mod shape;
pub mod vml;

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
    /// A package part or relationship could not be resolved.
    #[error("SpreadsheetML drawing package error: {0}")]
    Opc(#[from] litchi_opc::OpcError),
    /// `DrawingML` chart decoding or validation failed.
    #[error("DrawingML error: {0}")]
    Drawing(#[from] litchi_drawingml::Error),
}
