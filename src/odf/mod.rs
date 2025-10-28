//! OpenDocument Format (ODF) support.
//!
//! This module provides comprehensive support for parsing and working with OpenDocument
//! files, including text documents (.odt), spreadsheets (.ods), and presentations (.odp).
//!
//! # Features
//!
//! - Parse ODF files from paths or byte buffers
//! - Extract text, structured content, and metadata
//! - Support for styles and formatting
//! - Export capabilities (e.g., CSV for spreadsheets)
//!
//! # Examples
//!
//! ```no_run
//! use litchi::odf::{Document, Spreadsheet, Presentation};
//!
//! # fn main() -> litchi::Result<()> {
//! // Open a text document
//! let mut doc = Document::open("document.odt")?;
//! let text = doc.text()?;
//!
//! // Open a spreadsheet
//! let mut sheet = Spreadsheet::open("data.ods")?;
//! let csv = sheet.to_csv()?;
//!
//! // Open a presentation
//! let mut pres = Presentation::open("slides.odp")?;
//! let slide_count = pres.slide_count()?;
//!
//! # Ok(())
//! # }
//! ```

/// Core ODF parsing functionality
mod core;
/// ODF XML element classes
mod elements;
/// ODF presentation (.odp) support
mod odp;
/// ODF spreadsheet (.ods) support
mod ods;
/// ODF text document (.odt) support
mod odt;

// Re-export main types for convenience
pub use odp::Presentation;
pub use ods::{Cell as SCell, CellValue, Row as SRow, Sheet, Spreadsheet};
pub use odt::Document;

// Re-export shapes for presentations
pub use odp::{Shape, Slide};

// Re-export document element types for unified API (for ODT tables)
pub use elements::table::{Table, TableCell as Cell, TableRow as Row};
pub use elements::text::Paragraph;
pub use elements::text::Span as Run; // Span is equivalent to Run in ODF

/// ODF format types
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OdfFormat {
    /// OpenDocument Text (.odt)
    Text,
    /// OpenDocument Spreadsheet (.ods)
    Spreadsheet,
    /// OpenDocument Presentation (.odp)
    Presentation,
    /// OpenDocument Drawing (.odg)
    Drawing,
    /// OpenDocument Formula (.odf)
    Formula,
    /// OpenDocument Chart (.odc)
    Chart,
    /// OpenDocument Image (.odi)
    Image,
    /// OpenDocument Master (.odm)
    Master,
}

/// MIME types for different ODF formats
pub const ODF_MIME_TYPES: &[(&str, OdfFormat)] = &[
    ("application/vnd.oasis.opendocument.text", OdfFormat::Text),
    (
        "application/vnd.oasis.opendocument.spreadsheet",
        OdfFormat::Spreadsheet,
    ),
    (
        "application/vnd.oasis.opendocument.presentation",
        OdfFormat::Presentation,
    ),
    (
        "application/vnd.oasis.opendocument.graphics",
        OdfFormat::Drawing,
    ),
    (
        "application/vnd.oasis.opendocument.formula",
        OdfFormat::Formula,
    ),
    ("application/vnd.oasis.opendocument.chart", OdfFormat::Chart),
    ("application/vnd.oasis.opendocument.image", OdfFormat::Image),
    (
        "application/vnd.oasis.opendocument.text-master",
        OdfFormat::Master,
    ),
    // Template variants
    (
        "application/vnd.oasis.opendocument.text-template",
        OdfFormat::Text,
    ),
    (
        "application/vnd.oasis.opendocument.spreadsheet-template",
        OdfFormat::Spreadsheet,
    ),
    (
        "application/vnd.oasis.opendocument.presentation-template",
        OdfFormat::Presentation,
    ),
    (
        "application/vnd.oasis.opendocument.graphics-template",
        OdfFormat::Drawing,
    ),
    (
        "application/vnd.oasis.opendocument.formula-template",
        OdfFormat::Formula,
    ),
    (
        "application/vnd.oasis.opendocument.chart-template",
        OdfFormat::Chart,
    ),
    (
        "application/vnd.oasis.opendocument.image-template",
        OdfFormat::Image,
    ),
];

/// Detect ODF format from MIME type.
///
/// # Arguments
///
/// * `mime_type` - MIME type string to detect
///
/// # Returns
///
/// The corresponding `OdfFormat` if recognized, None otherwise.
///
/// # Examples
///
/// ```
/// use litchi::odf::{detect_format_from_mime, OdfFormat};
///
/// let format = detect_format_from_mime("application/vnd.oasis.opendocument.text");
/// assert_eq!(format, Some(OdfFormat::Text));
/// ```
pub fn detect_format_from_mime(mime_type: &str) -> Option<OdfFormat> {
    ODF_MIME_TYPES
        .iter()
        .find(|(mime, _)| *mime == mime_type)
        .map(|(_, format)| *format)
}
