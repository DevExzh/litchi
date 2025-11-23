//! XLS cell formatting (XF records, fonts, fills, borders)
//!
//! This module implements BIFF8 formatting records for Excel 97-2003 files.
//! Based on Microsoft's "[MS-XLS]" specification and Apache POI's implementation.
//!
//! # Key Structures
//!
//! - **XF (Extended Format)**: Cell format combining font, fill, border, and alignment
//! - **FONT**: Font definition (name, size, color, style)
//! - **FORMAT**: Number format definition
//! - **PALETTE**: Color palette

use super::super::XlsResult;
use std::collections::HashMap;
use std::io::Write;

/// Font weight constants
pub const FONT_WEIGHT_NORMAL: u16 = 400;
pub const FONT_WEIGHT_BOLD: u16 = 700;

/// Default color indices
pub const COLOR_BLACK: u16 = 0x08;
pub const COLOR_WHITE: u16 = 0x09;
pub const COLOR_RED: u16 = 0x0A;
pub const COLOR_AUTOMATIC: u16 = 0x7FFF;

/// Horizontal alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalAlignment {
    General = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterAcrossSelection = 6,
}

/// Built-in number format strings as defined by BIFF8 / Excel.
///
/// These are taken from Apache POI's `BuiltinFormats` table so that
/// format indices used in `ExtendedFormat.format_index` match POI and
/// Excel expectations.
const BUILTIN_NUMBER_FORMATS: [&str; 50] = [
    "General",                              // 0x00
    "0",                                    // 0x01
    "0.00",                                 // 0x02
    "#,##0",                                // 0x03
    "#,##0.00",                             // 0x04
    "\"$\"#,##0_);(\"$\"#,##0)",            // 0x05
    "\"$\"#,##0_);[Red](\"$\"#,##0)",       // 0x06
    "\"$\"#,##0.00_);(\"$\"#,##0.00)",      // 0x07
    "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)", // 0x08
    "0%",                                   // 0x09
    "0.00%",                                // 0x0A
    "0.00E+00",                             // 0x0B
    "# ?/?",                                // 0x0C
    "# ??/??",                              // 0x0D
    "m/d/yy",                               // 0x0E
    "d-mmm-yy",                             // 0x0F
    "d-mmm",                                // 0x10
    "mmm-yy",                               // 0x11
    "h:mm AM/PM",                           // 0x12
    "h:mm:ss AM/PM",                        // 0x13
    "h:mm",                                 // 0x14
    "h:mm:ss",                              // 0x15
    "m/d/yy h:mm",                          // 0x16
    // 0x17 - 0x24 reserved for international and undocumented
    "reserved-0x17",              // 0x17
    "reserved-0x18",              // 0x18
    "reserved-0x19",              // 0x19
    "reserved-0x1A",              // 0x1A
    "reserved-0x1B",              // 0x1B
    "reserved-0x1C",              // 0x1C
    "reserved-0x1D",              // 0x1D
    "reserved-0x1E",              // 0x1E
    "reserved-0x1F",              // 0x1F
    "reserved-0x20",              // 0x20
    "reserved-0x21",              // 0x21
    "reserved-0x22",              // 0x22
    "reserved-0x23",              // 0x23
    "reserved-0x24",              // 0x24
    "#,##0_);(#,##0)",            // 0x25
    "#,##0_);[Red](#,##0)",       // 0x26
    "#,##0.00_);(#,##0.00)",      // 0x27
    "#,##0.00_);[Red](#,##0.00)", // 0x28
    "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
    "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",
    "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
    "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",
    "mm:ss",     // 0x2D
    "[h]:mm:ss", // 0x2E
    "mm:ss.0",   // 0x2F
    "##0.0E+0",  // 0x30
    "@",         // 0x31 (text)
];

/// First user-defined number format index in BIFF8 / Excel.
const FIRST_USER_DEFINED_NUMBER_FORMAT_INDEX: u16 = 164;

/// Look up the BIFF built-in number format index for a given pattern.
///
/// This mirrors Apache POI's `BuiltinFormats.getBuiltinFormat(String)`
/// in a simplified form and is used by the formatting manager to avoid
/// creating duplicate custom FORMAT records for built-in patterns.
fn builtin_number_format_index(pattern: &str) -> Option<u16> {
    BUILTIN_NUMBER_FORMATS
        .iter()
        .position(|&p| p == pattern)
        .map(|idx| idx as u16)
}

/// Vertical alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top = 0,
    Center = 1,
    Bottom = 2,
    Justify = 3,
}

/// Border style
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderStyle {
    #[default]
    None = 0,
    Thin = 1,
    Medium = 2,
    Dashed = 3,
    Dotted = 4,
    Thick = 5,
    Double = 6,
    Hair = 7,
}

/// Fill pattern
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FillPattern {
    None = 0,
    Solid = 1,
    MediumGray = 2,
    DarkGray = 3,
    LightGray = 4,
    DarkHorizontal = 5,
    DarkVertical = 6,
    DarkDown = 7,
    DarkUp = 8,
    DarkGrid = 9,
    DarkTrellis = 10,
}

/// Font definition
#[derive(Debug, Clone)]
pub struct Font {
    /// Font height in twips (1/20 of a point)
    pub height: u16,
    /// Font weight (400 = normal, 700 = bold)
    pub weight: u16,
    /// Italic flag
    pub italic: bool,
    /// Underline style (0 = none, 1 = single, 2 = double)
    pub underline: u8,
    /// Font color index
    pub color_index: u16,
    /// Font name
    pub name: String,
}

impl Default for Font {
    fn default() -> Self {
        Self {
            height: 200, // 10pt
            weight: FONT_WEIGHT_NORMAL,
            italic: false,
            underline: 0,
            color_index: COLOR_AUTOMATIC,
            name: "Arial".to_string(),
        }
    }
}

/// Cell borders
#[derive(Debug, Clone, Default)]
pub struct Borders {
    pub left_style: BorderStyle,
    pub left_color: u16,
    pub right_style: BorderStyle,
    pub right_color: u16,
    pub top_style: BorderStyle,
    pub top_color: u16,
    pub bottom_style: BorderStyle,
    pub bottom_color: u16,
}

/// Cell fill (background)
#[derive(Debug, Clone)]
pub struct Fill {
    pub pattern: FillPattern,
    pub foreground_color: u16,
    pub background_color: u16,
}

impl Default for Fill {
    fn default() -> Self {
        Self {
            pattern: FillPattern::None,
            foreground_color: COLOR_AUTOMATIC,
            background_color: COLOR_AUTOMATIC,
        }
    }
}

/// Extended Format (XF) record - combines font, fill, border, alignment
#[derive(Debug, Clone)]
pub struct ExtendedFormat {
    /// Font index
    pub font_index: u16,
    /// Number format index
    pub format_index: u16,
    /// Horizontal alignment
    pub h_align: HorizontalAlignment,
    /// Vertical alignment
    pub v_align: VerticalAlignment,
    /// Text wrap
    pub text_wrap: bool,
    /// Borders
    pub borders: Borders,
    /// Fill
    pub fill: Fill,
}

impl Default for ExtendedFormat {
    fn default() -> Self {
        Self {
            font_index: 0,
            format_index: 0,
            h_align: HorizontalAlignment::General,
            v_align: VerticalAlignment::Bottom,
            text_wrap: false,
            borders: Borders::default(),
            fill: Fill::default(),
        }
    }
}

/// High-level cell style descriptor used to build reusable styles.
///
/// This is a value-based counterpart to POI's `HSSFCellStyle`: it
/// groups together font, borders, fill, alignment, and an optional
/// number format string. `FormattingManager` converts a `CellStyle`
/// into an `ExtendedFormat` plus FONT and FORMAT records.
#[derive(Debug, Clone)]
pub struct CellStyle {
    /// Font definition used by this style.
    pub font: Font,
    /// Cell borders (styles and colors).
    pub borders: Borders,
    /// Cell fill (background pattern and colors).
    pub fill: Fill,
    /// Horizontal alignment.
    pub h_align: HorizontalAlignment,
    /// Vertical alignment.
    pub v_align: VerticalAlignment,
    /// Whether text is wrapped within the cell.
    pub text_wrap: bool,
    /// Optional number format pattern (e.g. "0.00", "yyyy-mm-dd").
    pub number_format: Option<String>,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self {
            font: Font::default(),
            borders: Borders::default(),
            fill: Fill::default(),
            h_align: HorizontalAlignment::General,
            v_align: VerticalAlignment::Bottom,
            text_wrap: false,
            number_format: None,
        }
    }
}

/// Write FONT record (0x0031)
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `font` - Font definition
pub fn write_font<W: Write>(writer: &mut W, font: &Font) -> XlsResult<()> {
    let name_bytes = font.name.as_bytes();
    let name_len = name_bytes.len().min(255);

    // Fixed payload is 14 bytes of properties:
    // - Height (2) + Attributes (2) + ColorIdx (2) + Weight (2)
    // - Escapement (2) + Underline (1) + Family (1) + Charset (1) + Reserved (1)
    // BIFF8 then stores: NameLen (1), Options (1), Name bytes (N).
    let data_len = 14 + 1 + 1 + name_len; // properties + len + options + name
    super::biff::write_record_header(writer, 0x0031, data_len as u16)?;

    // Font height in twips
    writer.write_all(&font.height.to_le_bytes())?;

    // Option flags (italic, strikeout, etc.)
    let mut flags = 0u16;
    if font.italic {
        flags |= 0x0002;
    }
    writer.write_all(&flags.to_le_bytes())?;

    // Color index
    writer.write_all(&font.color_index.to_le_bytes())?;

    // Font weight
    writer.write_all(&font.weight.to_le_bytes())?;

    // Escapement type (0 = none, 1 = superscript, 2 = subscript)
    writer.write_all(&0u16.to_le_bytes())?;

    // Underline type
    writer.write_all(&[font.underline])?;

    // Font family (0 = None)
    writer.write_all(&[0])?;

    // Character set (0 = ANSI Latin)
    writer.write_all(&[0])?;

    // Reserved
    writer.write_all(&[0])?;

    // Font name length
    writer.write_all(&[name_len as u8])?;

    // Options: 0x00 = compressed 8-bit (ASCII), 0x01 = UTF-16LE.
    // We currently treat font names as ASCII for simplicity/performance.
    writer.write_all(&[0x00])?;

    // Font name bytes
    writer.write_all(&name_bytes[..name_len])?;

    Ok(())
}

/// Write XF (Extended Format) record (0x00E0)
///
/// # Arguments
///
/// * `writer` - Output writer
/// * `xf` - Extended format definition
/// * `is_style_xf` - True for style XF, false for cell XF
pub fn write_xf<W: Write>(writer: &mut W, xf: &ExtendedFormat, is_style_xf: bool) -> XlsResult<()> {
    super::biff::write_record_header(writer, 0x00E0, 20)?;

    // Font index
    writer.write_all(&xf.font_index.to_le_bytes())?;

    // Format index
    writer.write_all(&xf.format_index.to_le_bytes())?;

    // XF type, cell protection, parent style XF
    let xf_type: u16 = if is_style_xf { 0xFFF5 } else { 0x0001 };
    writer.write_all(&xf_type.to_le_bytes())?;

    // Alignment and break
    let mut align_flags = (xf.h_align as u8) | ((xf.v_align as u8) << 4);
    if xf.text_wrap {
        align_flags |= 0x08; // Wrap text bit
    }
    writer.write_all(&[align_flags])?;

    // Rotation
    writer.write_all(&[0])?;

    // Text direction, indent
    writer.write_all(&[0])?;

    // Used attributes flags
    writer.write_all(&[0])?;

    // Border styles bitfield (field_6_border_options).
    // Matches POI's ExtendedFormatRecord: 4-bit nibbles per side.
    let border_left = (xf.borders.left_style as u16) & 0x000F;
    let border_right = ((xf.borders.right_style as u16) & 0x000F) << 4;
    let border_top = ((xf.borders.top_style as u16) & 0x000F) << 8;
    let border_bottom = ((xf.borders.bottom_style as u16) & 0x000F) << 12;
    let border_options = border_left | border_right | border_top | border_bottom;
    writer.write_all(&border_options.to_le_bytes())?;

    // Border palette indices and diagonal flags (field_7_palette_options).
    let left_idx = xf.borders.left_color & 0x007F;
    let right_idx = xf.borders.right_color & 0x007F;
    let palette_options: u16 = (left_idx & 0x007F) | ((right_idx & 0x007F) << 7);
    writer.write_all(&palette_options.to_le_bytes())?;

    // Additional palette options and fill pattern (field_8_adtl_palette_options).
    let top_idx = xf.borders.top_color & 0x007F;
    let bottom_idx = xf.borders.bottom_color & 0x007F;
    let mut adtl_palette_options: u32 = 0;
    adtl_palette_options |= (top_idx as u32) & 0x0000_007F;
    adtl_palette_options |= ((bottom_idx as u32) & 0x0000_007F) << 7;
    // Diagonal and diagonal line style are left at 0 (not used).
    let fill_pattern_bits = (xf.fill.pattern as u32) & 0x3F;
    adtl_palette_options |= fill_pattern_bits << 26;
    writer.write_all(&adtl_palette_options.to_le_bytes())?;

    // Fill foreground and background palette indices (field_9_fill_palette_options).
    let fg_idx = xf.fill.foreground_color & 0x007F;
    let bg_idx = xf.fill.background_color & 0x007F;
    let fill_palette_options: u16 = (fg_idx & 0x007F) | ((bg_idx & 0x007F) << 7);
    writer.write_all(&fill_palette_options.to_le_bytes())?;

    Ok(())
}

/// Formatting manager for tracking fonts and formats
#[derive(Debug)]
pub struct FormattingManager {
    fonts: Vec<Font>,
    formats: Vec<ExtendedFormat>,
    // Custom number formats (FORMAT records) keyed by index code.
    // Built-in formats (0x00..0x31) come from BUILTIN_NUMBER_FORMATS.
    number_formats: Vec<(u16, String)>,
    number_format_map: HashMap<String, u16>,
}

impl FormattingManager {
    /// Create a new formatting manager with default entries
    pub fn new() -> Self {
        let mut manager = Self {
            fonts: Vec::new(),
            formats: Vec::new(),
            number_formats: Vec::new(),
            number_format_map: HashMap::new(),
        };

        // Add default fonts (indices 0..3) to approximate Excel/POI defaults.
        // 0: Normal
        manager.fonts.push(Font::default());
        // 1: Bold
        manager.fonts.push(Font {
            weight: FONT_WEIGHT_BOLD,
            ..Font::default()
        });
        // 2: Italic
        manager.fonts.push(Font {
            italic: true,
            ..Font::default()
        });
        // 3: Bold + Italic
        manager.fonts.push(Font {
            weight: FONT_WEIGHT_BOLD,
            italic: true,
            ..Font::default()
        });

        // Add default format (index 0)
        manager.formats.push(ExtendedFormat::default());

        manager
    }

    /// Add a font and return its index
    pub fn add_font(&mut self, font: Font) -> u16 {
        let index = self.fonts.len() as u16;
        self.fonts.push(font);
        index
    }

    /// Add a format and return its index
    pub fn add_format(&mut self, format: ExtendedFormat) -> u16 {
        let index = self.formats.len() as u16;
        self.formats.push(format);
        index
    }

    /// Register a number format pattern and return its BIFF format index.
    ///
    /// This mirrors POI's `HSSFDataFormat.getFormat` behavior:
    /// - Built-in formats (see `BUILTIN_NUMBER_FORMATS`) return their
    ///   predefined indices.
    /// - The "TEXT" alias normalizes to "@".
    /// - Custom patterns are assigned indices starting at 164 and
    ///   written as FORMAT (0x041E) records.
    pub fn register_number_format(&mut self, pattern: &str) -> u16 {
        // Normalize "TEXT" alias used by POI to "@".
        let normalized = if pattern.eq_ignore_ascii_case("TEXT") {
            "@"
        } else {
            pattern
        };

        // Built-in lookup
        if let Some(idx) = builtin_number_format_index(normalized) {
            return idx;
        }

        // Existing custom format
        if let Some(&idx) = self.number_format_map.get(normalized) {
            return idx;
        }

        // Allocate new user-defined format index starting at 164, as in BIFF8.
        let next_index = self.next_custom_format_index();
        self.number_formats
            .push((next_index, normalized.to_string()));
        self.number_format_map
            .insert(normalized.to_string(), next_index);
        next_index
    }

    /// Register a high-level `CellStyle` and return its internal style index.
    ///
    /// This helper wires fonts, number formats, and XF properties together:
    /// - The provided font is appended to the FONT table and its index stored
    ///   in the resulting `ExtendedFormat`.
    /// - If a number format pattern is specified, it is registered via
    ///   `register_number_format` and the resulting index is stored in
    ///   `ExtendedFormat.format_index`.
    /// - Borders, fills, and alignment settings are copied into the XF.
    pub fn register_cell_style(&mut self, style: CellStyle) -> u16 {
        let CellStyle {
            font,
            borders,
            fill,
            h_align,
            v_align,
            text_wrap,
            number_format,
        } = style;

        let font_index = self.add_font(font);
        let format_index = number_format
            .as_deref()
            .map(|pattern| self.register_number_format(pattern))
            .unwrap_or(0);

        let xf = ExtendedFormat {
            font_index,
            format_index,
            h_align,
            v_align,
            text_wrap,
            borders,
            fill,
        };

        self.add_format(xf)
    }

    /// Get font by index
    pub fn get_font(&self, index: u16) -> Option<&Font> {
        self.fonts.get(index as usize)
    }

    /// Get format by index
    pub fn get_format(&self, index: u16) -> Option<&ExtendedFormat> {
        self.formats.get(index as usize)
    }

    /// Write all FONT records
    pub fn write_fonts<W: Write>(&self, writer: &mut W) -> XlsResult<()> {
        for font in &self.fonts {
            write_font(writer, font)?;
        }
        Ok(())
    }

    /// Write all FORMAT records (0x041E): built-in indices 0..7 and any
    /// registered user-defined formats.
    pub fn write_number_formats<W: Write>(&self, writer: &mut W) -> XlsResult<()> {
        // POI's InternalWorkbook.createWorkbook emits FORMAT records for
        // built-in indices 0..7. We mirror that behavior here to keep
        // record streams comparable, even though Excel does not strictly
        // require FORMAT records for built-ins.
        for (index, format_str) in BUILTIN_NUMBER_FORMATS.iter().enumerate().take(8) {
            super::biff::write_format_record(writer, index as u16, format_str)?;
        }

        for (code, pattern) in &self.number_formats {
            super::biff::write_format_record(writer, *code, pattern)?;
        }

        Ok(())
    }

    /// Write all XF records
    pub fn write_formats<W: Write>(&self, writer: &mut W) -> XlsResult<()> {
        // Base style XF (matches Excel/POI defaults: General format, font 0)
        let base = ExtendedFormat::default();

        // 0..14: default style XFs
        // Map style XFs to fonts in a POI-like way:
        //  0: font 0 (Normal)
        //  1,2: font 1 (Bold)
        //  3,4: font 2 (Italic)
        //  5..14: font 0 (Normal)
        for i in 0..15 {
            let mut xf = base.clone();
            xf.font_index = match i {
                1 | 2 => 1,
                3 | 4 => 2,
                _ => 0,
            };
            write_xf(writer, &xf, true)?;
        }

        // 15: default cell XF used by our cell records
        if let Some(default_cell_xf) = self.formats.first() {
            write_xf(writer, default_cell_xf, false)?;
        } else {
            write_xf(writer, &base, false)?;
        }

        // 16..20: built-in style XFs for common number formats.
        // These mirror POI's mapping of style XFs to built-in number formats,
        // but we keep the structure minimal while remaining BIFF8-compliant.
        //
        // Format indices are BIFF built-ins:
        //  - 0x2B, 0x29, 0x2C, 0x2A: locale-dependent currency/comma styles
        //  - 0x09: percentage
        const BUILTIN_STYLE_FORMATS: [u16; 5] = [0x002B, 0x0029, 0x002C, 0x002A, 0x0009];
        for &fmt_idx in &BUILTIN_STYLE_FORMATS {
            let mut xf = base.clone();
            xf.format_index = fmt_idx;
            write_xf(writer, &xf, true)?;
        }

        // Additional cell XFs (if user-defined formats are ever added later)
        if self.formats.len() > 1 {
            for format in &self.formats[1..] {
                write_xf(writer, format, false)?;
            }
        }

        Ok(())
    }

    /// Compute the next available user-defined number format index.
    ///
    /// BIFF8 reserves built-in indices below 164; custom formats start
    /// at `FIRST_USER_DEFINED_NUMBER_FORMAT_INDEX`. We allocate indices
    /// monotonically increasing from that base, mirroring POI's
    /// `InternalWorkbook.getFormat` behavior.
    fn next_custom_format_index(&self) -> u16 {
        self.number_formats
            .iter()
            .map(|(code, _)| *code)
            .max()
            .map(|max_code| max_code.saturating_add(1))
            .unwrap_or(FIRST_USER_DEFINED_NUMBER_FORMAT_INDEX)
    }
    pub(crate) fn cell_xf_index_for(&self, format_index: u16) -> u16 {
        const STYLE_XF_COUNT: u16 = 15;
        const BUILTIN_STYLE_XF_COUNT: u16 = 5;
        const DEFAULT_CELL_XF_INDEX: u16 = STYLE_XF_COUNT;
        const USER_CELL_XF_START_INDEX: u16 = DEFAULT_CELL_XF_INDEX + 1 + BUILTIN_STYLE_XF_COUNT;

        if format_index == 0 {
            DEFAULT_CELL_XF_INDEX
        } else {
            USER_CELL_XF_START_INDEX + (format_index - 1)
        }
    }
}

impl Default for FormattingManager {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_font_creation() {
        let font = Font {
            name: "Arial".to_string(),
            height: 240, // 12pt
            weight: FONT_WEIGHT_BOLD,
            italic: true,
            ..Default::default()
        };
        assert_eq!(font.name, "Arial");
        assert_eq!(font.height, 240);
        assert!(font.italic);
    }

    #[test]
    fn test_formatting_manager() {
        let mut mgr = FormattingManager::new();

        let font_idx = mgr.add_font(Font {
            name: "Times".to_string(),
            weight: FONT_WEIGHT_BOLD,
            ..Default::default()
        });

        assert_eq!(font_idx, 4); // Indices 0..3 are default fonts
        assert_eq!(mgr.get_font(4).unwrap().name, "Times");
    }
}
