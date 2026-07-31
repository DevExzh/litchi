//! Shared DrawingML primitives independent of DOCX, PPTX, XLSX, and XLSB.

#![forbid(unsafe_code)]

pub mod blip;
pub mod ext;
pub mod fill;
pub mod xfrm;

use thiserror::Error;

/// Result of a DrawingML read operation.
pub type Result<T> = std::result::Result<T, DrawingError>;

/// Failure to parse a shared DrawingML primitive.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum DrawingError {
    /// The XML is malformed or contains an invalid encoded value.
    #[error("invalid DrawingML XML: {0}")]
    Xml(String),
}
