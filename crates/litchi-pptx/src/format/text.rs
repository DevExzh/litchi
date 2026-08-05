//! Text run styling values used by shape writers.

/// Optional styling applied to one PresentationML text run.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TextFormat {
    /// Font family.
    pub font: Option<String>,
    /// Font size in points.
    pub size: Option<f64>,
    /// Bold text.
    pub bold: Option<bool>,
    /// Italic text.
    pub italic: Option<bool>,
    /// Underlined text.
    pub underline: Option<bool>,
    /// Hexadecimal RGB text color, for example `FF0000`.
    pub color: Option<String>,
}
