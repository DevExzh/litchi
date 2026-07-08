//! Text and Paragraph Styling Information
//!
//! iWork documents support rich text with character-level and paragraph-level styling.

/// Text style properties (character-level)
#[derive(Debug, Clone, Default)]
pub struct TextStyle {
    /// Font family name
    pub font_family: Option<String>,
    /// Font size in points
    pub font_size: Option<f32>,
    /// Bold formatting
    pub bold: bool,
    /// Italic formatting
    pub italic: bool,
    /// Underline formatting
    pub underline: bool,
    /// Strikethrough formatting
    pub strikethrough: bool,
    /// Text color (RGB)
    pub color: Option<(u8, u8, u8)>,
}

impl TextStyle {
    /// Create a new default text style
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if style has any formatting applied
    pub fn has_formatting(&self) -> bool {
        self.bold
            || self.italic
            || self.underline
            || self.strikethrough
            || self.font_family.is_some()
            || self.font_size.is_some()
            || self.color.is_some()
    }
}

/// Paragraph style properties
#[derive(Debug, Clone, Default)]
pub struct ParagraphStyle {
    /// Alignment (left, center, right, justify)
    pub alignment: TextAlignment,
    /// First line indent (in points)
    pub first_line_indent: f32,
    /// Left indent (in points)
    pub left_indent: f32,
    /// Right indent (in points)  
    pub right_indent: f32,
    /// Space before paragraph (in points)
    pub space_before: f32,
    /// Space after paragraph (in points)
    pub space_after: f32,
    /// Line spacing multiplier
    pub line_spacing: f32,
}

impl ParagraphStyle {
    /// Create a new default paragraph style
    pub fn new() -> Self {
        Self {
            alignment: TextAlignment::Left,
            first_line_indent: 0.0,
            left_indent: 0.0,
            right_indent: 0.0,
            space_before: 0.0,
            space_after: 0.0,
            line_spacing: 1.0,
        }
    }
}

/// Text alignment options
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TextAlignment {
    #[default]
    Left,
    Center,
    Right,
    Justify,
}

impl TextAlignment {
    /// Parse alignment from integer value
    pub fn from_i32(value: i32) -> Self {
        match value {
            0 => Self::Left,
            1 => Self::Center,
            2 => Self::Right,
            3 => Self::Justify,
            _ => Self::Left,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_style_creation() {
        let style = TextStyle::new();
        assert!(!style.has_formatting());

        let mut styled = TextStyle::new();
        styled.bold = true;
        styled.font_size = Some(14.0);
        assert!(styled.has_formatting());
    }

    #[test]
    fn test_paragraph_style() {
        let para = ParagraphStyle::new();
        assert_eq!(para.alignment, TextAlignment::Left);
        assert_eq!(para.line_spacing, 1.0);
    }

    #[test]
    fn test_alignment_parsing() {
        assert_eq!(TextAlignment::from_i32(0), TextAlignment::Left);
        assert_eq!(TextAlignment::from_i32(1), TextAlignment::Center);
        assert_eq!(TextAlignment::from_i32(2), TextAlignment::Right);
        assert_eq!(TextAlignment::from_i32(3), TextAlignment::Justify);
        assert_eq!(TextAlignment::from_i32(999), TextAlignment::Left); // default
    }
}
