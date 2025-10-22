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
