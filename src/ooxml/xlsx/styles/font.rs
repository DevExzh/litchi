//! Font information and definitions.

/// Font information.
///
/// Defines the visual appearance of text in cells including
/// typeface, size, color, and text decoration.
#[derive(Debug, Clone, Default)]
pub struct Font {
    /// Font name/family (e.g., "Calibri", "Arial")
    pub name: Option<String>,
    /// Font size in points
    pub size: Option<f64>,
    /// Bold flag
    pub bold: bool,
    /// Italic flag
    pub italic: bool,
    /// Underline style
    pub underline: Option<String>,
    /// Strike-through flag
    pub strike: bool,
    /// Font color (RGB hex or theme color reference)
    pub color: Option<String>,
    /// Font charset
    pub charset: Option<u32>,
    /// Font family (1=Roman, 2=Swiss, 3=Modern, 4=Script, 5=Decorative)
    pub family: Option<u32>,
    /// Font scheme (major, minor, none)
    pub scheme: Option<String>,
}

impl Font {
    /// Create a new default font.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if the font has any special formatting.
    #[inline]
    pub fn has_formatting(&self) -> bool {
        self.bold || self.italic || self.strike || self.underline.is_some()
    }
}
