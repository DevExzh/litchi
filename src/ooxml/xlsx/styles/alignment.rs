//! Cell alignment information.

/// Alignment information for cell content.
///
/// Controls how text is positioned within a cell both horizontally
/// and vertically, as well as text wrapping and rotation.
#[derive(Debug, Clone, Default)]
pub struct Alignment {
    /// Horizontal alignment (e.g., "left", "center", "right", "fill", "justify")
    pub horizontal: Option<String>,
    /// Vertical alignment (e.g., "top", "center", "bottom", "justify")
    pub vertical: Option<String>,
    /// Text rotation (angle in degrees, 0-180, or 255 for vertical)
    pub text_rotation: Option<u32>,
    /// Wrap text flag
    pub wrap_text: bool,
    /// Indent level (for horizontal alignment)
    pub indent: Option<u32>,
    /// Shrink to fit flag
    pub shrink_to_fit: bool,
    /// Reading order (0=context, 1=LTR, 2=RTL)
    pub reading_order: Option<u32>,
}

impl Alignment {
    /// Create a new default alignment.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Create a new alignment with horizontal and vertical settings.
    #[inline]
    pub fn with_alignment(horizontal: Option<String>, vertical: Option<String>) -> Self {
        Self {
            horizontal,
            vertical,
            ..Default::default()
        }
    }

    /// Check if this alignment has any non-default settings.
    #[inline]
    pub fn has_settings(&self) -> bool {
        self.horizontal.is_some()
            || self.vertical.is_some()
            || self.text_rotation.is_some()
            || self.wrap_text
            || self.indent.is_some()
            || self.shrink_to_fit
            || self.reading_order.is_some()
    }
}
