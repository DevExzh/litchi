//! Border styles and definitions.

/// Border information for a cell.
///
/// Defines the borders on all four sides of a cell.
#[derive(Debug, Clone, Default)]
pub struct Border {
    /// Left border style
    pub left: Option<BorderStyle>,
    /// Right border style
    pub right: Option<BorderStyle>,
    /// Top border style
    pub top: Option<BorderStyle>,
    /// Bottom border style
    pub bottom: Option<BorderStyle>,
    /// Diagonal border style
    pub diagonal: Option<BorderStyle>,
    /// Diagonal direction (0=none, 1=up, 2=down, 3=both)
    pub diagonal_direction: Option<u32>,
}

/// Individual border style information.
#[derive(Debug, Clone)]
pub struct BorderStyle {
    /// Border style name (e.g., "thin", "medium", "thick", "double")
    pub style: String,
    /// Border color (RGB hex or theme color reference)
    pub color: Option<String>,
}

impl Border {
    /// Create a new empty border (no borders on any side).
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if this border has any visible borders.
    #[inline]
    pub fn has_borders(&self) -> bool {
        self.left.is_some()
            || self.right.is_some()
            || self.top.is_some()
            || self.bottom.is_some()
            || self.diagonal.is_some()
    }
}

impl BorderStyle {
    /// Create a new border style.
    #[inline]
    pub fn new(style: String, color: Option<String>) -> Self {
        Self { style, color }
    }
}
