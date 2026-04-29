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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fill_default() {
        let fill = Fill::default();
        assert!(matches!(fill, Fill::None));
    }

    #[test]
    fn test_fill_solid() {
        let fill = Fill::solid("FF0000".to_string());
        assert!(fill.is_solid());
        match &fill {
            Fill::Pattern {
                pattern_type,
                fg_color,
                bg_color,
            } => {
                assert_eq!(pattern_type, "solid");
                assert_eq!(*fg_color, Some("FF0000".to_string()));
                assert!(bg_color.is_none());
            },
            _ => panic!("Expected Pattern fill"),
        }
    }

    #[test]
    fn test_fill_pattern() {
        let fill = Fill::pattern(
            "gray125".to_string(),
            Some("CCCCCC".to_string()),
            Some("FFFFFF".to_string()),
        );
        assert!(!fill.is_solid());
        match &fill {
            Fill::Pattern {
                pattern_type,
                fg_color,
                bg_color,
            } => {
                assert_eq!(pattern_type, "gray125");
                assert_eq!(*fg_color, Some("CCCCCC".to_string()));
                assert_eq!(*bg_color, Some("FFFFFF".to_string()));
            },
            _ => panic!("Expected Pattern fill"),
        }
    }

    #[test]
    fn test_fill_gradient() {
        let fill = Fill::Gradient {
            gradient_type: Some("linear".to_string()),
            stops: vec![(0.0, "FF0000".to_string()), (1.0, "00FF00".to_string())],
        };
        assert!(!fill.is_solid());
        match fill {
            Fill::Gradient {
                gradient_type,
                stops,
            } => {
                assert_eq!(gradient_type, Some("linear".to_string()));
                assert_eq!(stops.len(), 2);
            },
            _ => panic!("Expected Gradient fill"),
        }
    }

    #[test]
    fn test_fill_none_is_not_solid() {
        let fill = Fill::None;
        assert!(!fill.is_solid());
    }

    #[test]
    fn test_fill_clone() {
        let fill = Fill::solid("00FF00".to_string());
        let fill2 = fill.clone();
        assert!(fill2.is_solid());
    }
}
