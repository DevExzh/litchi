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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_default() {
        let font = Font::default();
        assert!(font.name.is_none());
        assert!(font.size.is_none());
        assert!(!font.bold);
        assert!(!font.italic);
        assert!(font.underline.is_none());
        assert!(!font.strike);
        assert!(font.color.is_none());
        assert!(font.charset.is_none());
        assert!(font.family.is_none());
        assert!(font.scheme.is_none());
    }

    #[test]
    fn test_font_new() {
        let font = Font::new();
        assert!(font.name.is_none());
        assert!(!font.has_formatting());
    }

    #[test]
    fn test_font_has_formatting() {
        // Default font has no formatting
        let font = Font::default();
        assert!(!font.has_formatting());

        // Bold
        let font = Font {
            bold: true,
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Italic
        let font = Font {
            italic: true,
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Strike
        let font = Font {
            strike: true,
            ..Default::default()
        };
        assert!(font.has_formatting());

        // Underline
        let font = Font {
            underline: Some("single".to_string()),
            ..Default::default()
        };
        assert!(font.has_formatting());
    }

    #[test]
    fn test_font_full() {
        let font = Font {
            name: Some("Arial".to_string()),
            size: Some(12.0),
            bold: true,
            italic: false,
            underline: Some("single".to_string()),
            strike: false,
            color: Some("FF0000".to_string()),
            charset: Some(1),
            family: Some(2),
            scheme: Some("minor".to_string()),
        };
        assert_eq!(font.name, Some("Arial".to_string()));
        assert_eq!(font.size, Some(12.0));
        assert!(font.bold);
        assert!(!font.italic);
        assert_eq!(font.underline, Some("single".to_string()));
        assert!(!font.strike);
        assert_eq!(font.color, Some("FF0000".to_string()));
        assert_eq!(font.charset, Some(1));
        assert_eq!(font.family, Some(2));
        assert_eq!(font.scheme, Some("minor".to_string()));
        assert!(font.has_formatting());
    }

    #[test]
    fn test_font_clone() {
        let font = Font {
            name: Some("Calibri".to_string()),
            size: Some(11.0),
            bold: false,
            italic: true,
            ..Default::default()
        };
        let font2 = font.clone();
        assert_eq!(font.name, font2.name);
        assert_eq!(font.italic, font2.italic);
    }
}
