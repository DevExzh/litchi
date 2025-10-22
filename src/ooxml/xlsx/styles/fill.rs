//! Fill patterns and colors.

/// Fill information.
///
/// Defines the background fill for cells, either as a solid color,
/// pattern, or gradient.
#[derive(Debug, Clone, Default)]
pub enum Fill {
    /// No fill
    #[default]
    None,
    /// Pattern fill with colors
    Pattern {
        /// Pattern type (e.g., "solid", "gray125", "lightGray")
        pattern_type: String,
        /// Foreground color (RGB hex or theme color reference)
        fg_color: Option<String>,
        /// Background color (RGB hex or theme color reference)
        bg_color: Option<String>,
    },
    /// Gradient fill (simplified representation)
    Gradient {
        /// Gradient type (linear or path)
        gradient_type: Option<String>,
        /// Gradient stops (position, color pairs)
        stops: Vec<(f64, String)>,
    },
}

impl Fill {
    /// Create a new solid fill with the given color.
    #[inline]
    pub fn solid(color: String) -> Self {
        Fill::Pattern {
            pattern_type: "solid".to_string(),
            fg_color: Some(color),
            bg_color: None,
        }
    }

    /// Create a new pattern fill.
    #[inline]
    pub fn pattern(
        pattern_type: String,
        fg_color: Option<String>,
        bg_color: Option<String>,
    ) -> Self {
        Fill::Pattern {
            pattern_type,
            fg_color,
            bg_color,
        }
    }

    /// Check if this is a solid fill.
    pub fn is_solid(&self) -> bool {
        matches!(self, Fill::Pattern { pattern_type, .. } if pattern_type == "solid")
    }
}
