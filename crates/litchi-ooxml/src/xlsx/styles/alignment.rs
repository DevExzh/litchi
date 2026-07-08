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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_alignment_default() {
        let alignment = Alignment::default();
        assert!(alignment.horizontal.is_none());
        assert!(alignment.vertical.is_none());
        assert!(alignment.text_rotation.is_none());
        assert!(!alignment.wrap_text);
        assert!(alignment.indent.is_none());
        assert!(!alignment.shrink_to_fit);
        assert!(alignment.reading_order.is_none());
    }

    #[test]
    fn test_alignment_new() {
        let alignment = Alignment::new();
        assert!(alignment.horizontal.is_none());
        assert!(!alignment.has_settings());
    }

    #[test]
    fn test_alignment_with_alignment() {
        let alignment =
            Alignment::with_alignment(Some("center".to_string()), Some("middle".to_string()));
        assert_eq!(alignment.horizontal, Some("center".to_string()));
        assert_eq!(alignment.vertical, Some("middle".to_string()));
        assert!(!alignment.wrap_text);
    }

    #[test]
    fn test_alignment_has_settings() {
        // Default has no settings
        let alignment = Alignment::default();
        assert!(!alignment.has_settings());

        // Horizontal alignment
        let alignment = Alignment::with_alignment(Some("left".to_string()), None);
        assert!(alignment.has_settings());

        // Vertical alignment
        let alignment = Alignment {
            vertical: Some("top".to_string()),
            ..Default::default()
        };
        assert!(alignment.has_settings());

        // Text rotation
        let alignment = Alignment {
            text_rotation: Some(90),
            ..Default::default()
        };
        assert!(alignment.has_settings());

        // Wrap text
        let alignment = Alignment {
            wrap_text: true,
            ..Default::default()
        };
        assert!(alignment.has_settings());

        // Indent
        let alignment = Alignment {
            indent: Some(2),
            ..Default::default()
        };
        assert!(alignment.has_settings());

        // Shrink to fit
        let alignment = Alignment {
            shrink_to_fit: true,
            ..Default::default()
        };
        assert!(alignment.has_settings());

        // Reading order
        let alignment = Alignment {
            reading_order: Some(1),
            ..Default::default()
        };
        assert!(alignment.has_settings());
    }

    #[test]
    fn test_alignment_clone() {
        let alignment = Alignment {
            horizontal: Some("right".to_string()),
            vertical: Some("bottom".to_string()),
            text_rotation: Some(45),
            wrap_text: true,
            indent: Some(1),
            shrink_to_fit: false,
            reading_order: Some(2),
        };
        let alignment2 = alignment.clone();
        assert_eq!(alignment.horizontal, alignment2.horizontal);
        assert_eq!(alignment.text_rotation, alignment2.text_rotation);
        assert_eq!(alignment.wrap_text, alignment2.wrap_text);
    }
}
