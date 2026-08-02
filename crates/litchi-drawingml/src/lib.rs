//! Shared DrawingML grammar independent of DOCX, PPTX, XLSX, and XLSB.
//!
//! [`chart`] owns the chart model and XML codec, while [`diagram`] owns the
//! host-neutral SmartArt model and part grammar. Concrete formats retain only
//! their package relationships and anchoring semantics.

#![forbid(unsafe_code)]

pub mod blip;
pub mod chart;
pub mod diagram;
pub mod ext;
pub mod fill;
pub mod xfrm;

use thiserror::Error;

/// Result of a DrawingML read operation.
pub type Result<T> = std::result::Result<T, Error>;

/// Failure to parse a shared DrawingML primitive.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The XML is malformed or contains an invalid encoded value.
    #[error("invalid DrawingML XML: {0}")]
    Xml(String),

    /// The document violates DrawingML structural or value constraints.
    #[error("invalid DrawingML: {0}")]
    Invalid(String),

    /// Reading a caller-provided chart stream failed.
    #[error("DrawingML input error: {0}")]
    Io(#[from] std::io::Error),

    /// Markup-compatibility preprocessing failed.
    #[error("DrawingML markup compatibility error: {0}")]
    Mce(#[from] litchi_ooxml_common::MceError),

    /// Shared OOXML entity or namespace decoding failed.
    #[error("DrawingML shared XML error: {0}")]
    Common(#[from] litchi_ooxml_common::XmlError),
}

impl From<quick_xml::Error> for Error {
    fn from(error: quick_xml::Error) -> Self {
        Self::Xml(error.to_string())
    }
}
