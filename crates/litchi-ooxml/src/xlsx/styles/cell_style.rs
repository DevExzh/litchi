//! Cell style format records.

use super::Alignment;

/// Cell style information.
///
/// This represents a complete cell format that references fonts, fills,
/// borders, and number formats by their IDs.
#[derive(Debug, Clone, Default)]
pub struct CellStyle {
    /// Number format ID (references built-in or custom number format)
    pub num_fmt_id: Option<u32>,
    /// Font ID (index into fonts array)
    pub font_id: Option<u32>,
    /// Fill ID (index into fills array)
    pub fill_id: Option<u32>,
    /// Border ID (index into borders array)
    pub border_id: Option<u32>,
    /// Cell style format ID (references cellStyleXfs)
    pub xf_id: Option<u32>,
    /// Alignment information
    pub alignment: Option<Alignment>,
    /// Apply number format flag
    pub apply_number_format: bool,
    /// Apply font flag
    pub apply_font: bool,
    /// Apply fill flag
    pub apply_fill: bool,
    /// Apply border flag
    pub apply_border: bool,
    /// Apply alignment flag
    pub apply_alignment: bool,
    /// Quote prefix flag (for preserving leading apostrophe)
    pub quote_prefix: bool,
}

impl CellStyle {
    /// Create a new empty cell style.
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if this style has any formatting applied.
    #[inline]
    pub fn has_formatting(&self) -> bool {
        self.num_fmt_id.is_some()
            || self.font_id.is_some()
            || self.fill_id.is_some()
            || self.border_id.is_some()
            || self.alignment.is_some()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_style_default() {
        let style = CellStyle::default();
        assert!(style.num_fmt_id.is_none());
        assert!(style.font_id.is_none());
        assert!(style.fill_id.is_none());
        assert!(style.border_id.is_none());
        assert!(style.xf_id.is_none());
        assert!(style.alignment.is_none());
        assert!(!style.apply_number_format);
        assert!(!style.apply_font);
        assert!(!style.apply_fill);
        assert!(!style.apply_border);
        assert!(!style.apply_alignment);
        assert!(!style.quote_prefix);
    }

    #[test]
    fn test_cell_style_new() {
        let style = CellStyle::new();
        assert!(style.num_fmt_id.is_none());
        assert!(!style.has_formatting());
    }

    #[test]
    fn test_cell_style_has_formatting() {
        // Empty style has no formatting
        let style = CellStyle::default();
        assert!(!style.has_formatting());

        // Number format
        let style = CellStyle {
            num_fmt_id: Some(1),
            ..Default::default()
        };
        assert!(style.has_formatting());

        // Font
        let style = CellStyle {
            font_id: Some(0),
            ..Default::default()
        };
        assert!(style.has_formatting());

        // Fill
        let style = CellStyle {
            fill_id: Some(2),
            ..Default::default()
        };
        assert!(style.has_formatting());

        // Border
        let style = CellStyle {
            border_id: Some(0),
            ..Default::default()
        };
        assert!(style.has_formatting());

        // Alignment
        let style = CellStyle {
            alignment: Some(Alignment::new()),
            ..Default::default()
        };
        assert!(style.has_formatting());
    }

    #[test]
    fn test_cell_style_clone() {
        let style = CellStyle {
            num_fmt_id: Some(14),
            font_id: Some(1),
            fill_id: Some(2),
            border_id: Some(0),
            xf_id: Some(0),
            alignment: Some(Alignment::with_alignment(Some("center".to_string()), None)),
            apply_number_format: true,
            apply_font: true,
            apply_fill: true,
            apply_border: true,
            apply_alignment: true,
            quote_prefix: false,
        };
        let style2 = style.clone();
        assert_eq!(style.num_fmt_id, style2.num_fmt_id);
        assert_eq!(style.font_id, style2.font_id);
        assert_eq!(style.apply_font, style2.apply_font);
    }
}
