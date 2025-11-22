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

    // Border styles and colors (simplified)
    let border_left = (xf.borders.left_style as u8) & 0x0F;
    let border_right = (xf.borders.right_style as u8) & 0x0F;
    let border_top = (xf.borders.top_style as u8) & 0x0F;
    let border_bottom = (xf.borders.bottom_style as u8) & 0x0F;

    writer.write_all(&[border_left | (border_right << 4)])?;
    writer.write_all(&[border_top | (border_bottom << 4)])?;

    // Border colors (2 bytes each, we'll use default black)
    writer.write_all(&xf.borders.left_color.to_le_bytes())?;
    writer.write_all(&xf.borders.right_color.to_le_bytes())?;

    // Fill pattern and colors (simplified - full implementation would need more fields)
    writer.write_all(&[((xf.fill.pattern as u8) & 0x3F), 0])?;
    writer.write_all(&xf.fill.foreground_color.to_le_bytes())?;

    Ok(())
}

/// Formatting manager for tracking fonts and formats
#[derive(Debug)]
pub struct FormattingManager {
    fonts: Vec<Font>,
    formats: Vec<ExtendedFormat>,
}

impl FormattingManager {
    /// Create a new formatting manager with default entries
    pub fn new() -> Self {
        let mut manager = Self {
            fonts: Vec::new(),
            formats: Vec::new(),
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
