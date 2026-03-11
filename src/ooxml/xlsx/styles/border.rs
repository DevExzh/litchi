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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_border_default() {
        let border = Border::default();
        assert!(border.left.is_none());
        assert!(border.right.is_none());
        assert!(border.top.is_none());
        assert!(border.bottom.is_none());
        assert!(border.diagonal.is_none());
        assert!(border.diagonal_direction.is_none());
    }

    #[test]
    fn test_border_new() {
        let border = Border::new();
        assert!(border.left.is_none());
        assert!(!border.has_borders());
    }

    #[test]
    fn test_border_has_borders() {
        // No borders
        let border = Border::default();
        assert!(!border.has_borders());

        // Left border
        let border = Border {
            left: Some(BorderStyle::new("thin".to_string(), None)),
            ..Default::default()
        };
        assert!(border.has_borders());

        // Right border
        let border = Border {
            right: Some(BorderStyle::new(
                "medium".to_string(),
                Some("FF0000".to_string()),
            )),
            ..Default::default()
        };
        assert!(border.has_borders());

        // Top border
        let border = Border {
            top: Some(BorderStyle::new("thick".to_string(), None)),
            ..Default::default()
        };
        assert!(border.has_borders());

        // Bottom border
        let border = Border {
            bottom: Some(BorderStyle::new("double".to_string(), None)),
            ..Default::default()
        };
        assert!(border.has_borders());

        // Diagonal border
        let border = Border {
            diagonal: Some(BorderStyle::new("dashed".to_string(), None)),
            diagonal_direction: Some(1),
            ..Default::default()
        };
        assert!(border.has_borders());
    }

    #[test]
    fn test_border_style_new() {
        let style = BorderStyle::new("thin".to_string(), Some("000000".to_string()));
        assert_eq!(style.style, "thin");
        assert_eq!(style.color, Some("000000".to_string()));

        let style = BorderStyle::new("thick".to_string(), None);
        assert_eq!(style.style, "thick");
        assert!(style.color.is_none());
    }

    #[test]
    fn test_border_style_clone() {
        let style = BorderStyle::new("medium".to_string(), Some("FF0000".to_string()));
        let style2 = style.clone();
        assert_eq!(style.style, style2.style);
        assert_eq!(style.color, style2.color);
    }

    #[test]
    fn test_border_clone() {
        let border = Border {
            left: Some(BorderStyle::new("thin".to_string(), None)),
            right: Some(BorderStyle::new("thin".to_string(), None)),
            top: None,
            bottom: None,
            diagonal: None,
            diagonal_direction: None,
        };
        let border2 = border.clone();
        assert!(border2.left.is_some());
        assert!(border2.right.is_some());
        assert!(border2.top.is_none());
    }
}
