
/// Core ODF parsing functionality
mod core;
/// ODF XML element classes
mod elements;
/// ODF text document (.odt) support
mod text;
/// ODF spreadsheet (.ods) support
mod spreadsheet;
/// ODF presentation (.odp) support
mod presentation;

/// Re-export the main APIs
pub use text::Document;
pub use spreadsheet::Spreadsheet;
pub use presentation::Presentation;

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
    ("application/vnd.oasis.opendocument.spreadsheet", OdfFormat::Spreadsheet),
    ("application/vnd.oasis.opendocument.presentation", OdfFormat::Presentation),
    ("application/vnd.oasis.opendocument.graphics", OdfFormat::Drawing),
    ("application/vnd.oasis.opendocument.formula", OdfFormat::Formula),
    ("application/vnd.oasis.opendocument.chart", OdfFormat::Chart),
    ("application/vnd.oasis.opendocument.image", OdfFormat::Image),
    ("application/vnd.oasis.opendocument.text-master", OdfFormat::Master),
    // Template variants
    ("application/vnd.oasis.opendocument.text-template", OdfFormat::Text),
    ("application/vnd.oasis.opendocument.spreadsheet-template", OdfFormat::Spreadsheet),
    ("application/vnd.oasis.opendocument.presentation-template", OdfFormat::Presentation),
    ("application/vnd.oasis.opendocument.graphics-template", OdfFormat::Drawing),
    ("application/vnd.oasis.opendocument.formula-template", OdfFormat::Formula),
    ("application/vnd.oasis.opendocument.chart-template", OdfFormat::Chart),
    ("application/vnd.oasis.opendocument.image-template", OdfFormat::Image),
];

/// Detect ODF format from MIME type
pub fn detect_format_from_mime(mime_type: &str) -> Option<OdfFormat> {
    ODF_MIME_TYPES
        .iter()
        .find(|(mime, _)| *mime == mime_type)
        .map(|(_, format)| *format)
}

