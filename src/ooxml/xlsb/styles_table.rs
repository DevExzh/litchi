//! XLSB styles table parsing
//!
//! This module provides the main StylesTable container for Excel Binary Workbook styles,
//! including fonts, fills, borders, number formats, and cell formats.
//! Reference: [MS-XLSB] Section 2.4 - Styles Part

use crate::common::binary;
use crate::ooxml::xlsb::error::{XlsbError, XlsbResult};
use crate::ooxml::xlsb::records::{XlsbRecordIter, record_types, wide_str_with_len};
use crate::ooxml::xlsb::styles::{Alignment, Border};
use std::collections::HashMap;
use std::io::Read;

/// Font information
///
/// Fields ordered for optimal memory layout: String (24 bytes), f64 (8 bytes),
/// Option<u32> (8 bytes), then bools (1 byte each, but padded).
#[derive(Debug, Clone)]
pub struct Font {
    pub name: String,
    pub size: f64,
    pub color: Option<u32>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
}

impl Default for Font {
    fn default() -> Self {
        Font {
            name: "Calibri".to_string(),
            size: 11.0,
            color: None,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
        }
    }
}

/// Fill pattern
#[derive(Debug, Clone, Default)]
pub struct Fill {
    pub pattern_type: u32,
    pub fg_color: Option<u32>,
    pub bg_color: Option<u32>,
}

// Border types are now imported from styles module

/// Number format
#[allow(dead_code)]
#[derive(Debug, Clone)]
pub struct NumberFormat {
    pub id: u32,
    pub format_code: String,
}

/// Cell format (XF record)
#[derive(Debug, Clone)]
pub struct CellFormat {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub alignment: Option<Alignment>,
}

// Alignment type is now imported from styles module

/// Styles table container
#[derive(Debug, Clone)]
pub struct StylesTable {
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub num_fmts: HashMap<u32, String>,
    pub cell_xfs: Vec<CellFormat>,
    pub cell_style_xfs: Vec<CellFormat>,
}

impl Default for StylesTable {
    fn default() -> Self {
        StylesTable {
            fonts: vec![Font::default()],
            fills: vec![Fill::default()],
            borders: vec![Border::default()],
            num_fmts: Self::builtin_formats(),
            cell_xfs: Vec::new(),
            cell_style_xfs: Vec::new(),
        }
    }
}

impl StylesTable {
    /// Load styles from styles.bin file content
    pub fn from_reader<R: Read>(reader: R) -> XlsbResult<Self> {
        let mut styles = StylesTable::default();
        let mut iter = XlsbRecordIter::new(reader);

        let mut in_fonts = false;
        let mut in_fills = false;
        let mut in_borders = false;
        let mut in_fmts = false;
        let mut in_cell_xfs = false;
        let mut in_cell_style_xfs = false;

        for record in iter.by_ref() {
            let record = record?;
            let rec_type = record.header.record_type;
            let data = &record.data;

            match rec_type {
                record_types::BEGIN_FONTS => {
                    in_fonts = true;
                    styles.fonts.clear(); // Clear default
                },
                record_types::END_FONTS => in_fonts = false,
                record_types::FONT if in_fonts => {
                    if let Ok(font) = Self::parse_font(data) {
                        styles.fonts.push(font);
                    }
                },
                record_types::BEGIN_FILLS => {
                    in_fills = true;
                    styles.fills.clear(); // Clear default
                },
                record_types::END_FILLS => in_fills = false,
                record_types::FILL if in_fills => {
                    if let Ok(fill) = Self::parse_fill(data) {
                        styles.fills.push(fill);
                    }
                },
                record_types::BEGIN_BORDERS => {
                    in_borders = true;
                    styles.borders.clear(); // Clear default
                },
                record_types::END_BORDERS => in_borders = false,
                record_types::BORDER if in_borders => {
                    if let Ok(border) = Self::parse_border(data) {
                        styles.borders.push(border);
                    }
                },
                record_types::BEGIN_FMTS => in_fmts = true,
                record_types::END_FMTS => in_fmts = false,
                record_types::FMT if in_fmts => {
                    if let Ok((id, format_code)) = Self::parse_num_fmt(data) {
                        styles.num_fmts.insert(id, format_code);
                    }
                },
                record_types::BEGIN_CELL_XFS => in_cell_xfs = true,
                record_types::END_CELL_XFS => in_cell_xfs = false,
                record_types::XF if in_cell_xfs => {
                    if let Ok(xf) = Self::parse_xf(data) {
                        styles.cell_xfs.push(xf);
                    }
                },
                record_types::BEGIN_CELL_STYLE_XFS => in_cell_style_xfs = true,
                record_types::END_CELL_STYLE_XFS => in_cell_style_xfs = false,
                record_types::XF if in_cell_style_xfs => {
                    if let Ok(xf) = Self::parse_xf(data) {
                        styles.cell_style_xfs.push(xf);
                    }
                },
                _ => {
                    // Skip other records
                },
            }
        }

        Ok(styles)
    }

    /// Parse font record
    fn parse_font(data: &[u8]) -> XlsbResult<Font> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        // Font height in twips (1/20 of a point)
        let height = binary::read_u16_le_at(data, 0)?;
        let size = height as f64 / 20.0;

        let flags = binary::read_u16_le_at(data, 2)?;
        let bold = (flags & 0x0001) != 0;
        let italic = (flags & 0x0002) != 0;
        let underline = (flags & 0x0004) != 0;
        let strike = (flags & 0x0008) != 0;

        // Color (optional, 4 bytes ARGB)
        let color = if data.len() >= 12 {
            Some(binary::read_u32_le_at(data, 8)?)
        } else {
            None
        };

        // Font name
        let name_offset = if data.len() >= 12 { 12 } else { 8 };
        let (name, _) = if data.len() > name_offset {
            wide_str_with_len(&data[name_offset..])?
        } else {
            ("Calibri".to_string(), 0)
        };

        Ok(Font {
            size,
            name,
            bold,
            italic,
            underline,
            strike,
            color,
        })
    }

    /// Parse fill record
    fn parse_fill(data: &[u8]) -> XlsbResult<Fill> {
        if data.is_empty() {
            return Ok(Fill::default());
        }

        let pattern_type = if !data.is_empty() { data[0] as u32 } else { 0 };

        let fg_color = if data.len() >= 5 {
            Some(binary::read_u32_le_at(data, 1)?)
        } else {
            None
        };

        let bg_color = if data.len() >= 9 {
            Some(binary::read_u32_le_at(data, 5)?)
        } else {
            None
        };

        Ok(Fill {
            pattern_type,
            fg_color,
            bg_color,
        })
    }

    /// Parse border record using the dedicated border parser
    fn parse_border(data: &[u8]) -> XlsbResult<Border> {
        Border::parse(data)
    }

    /// Parse number format record
    fn parse_num_fmt(data: &[u8]) -> XlsbResult<(u32, String)> {
        if data.len() < 4 {
            return Err(XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }

        let id = binary::read_u32_le_at(data, 0)?;
        let (format_code, _) = wide_str_with_len(&data[4..])?;

        Ok((id, format_code))
    }

    /// Parse XF (cell format) record
    fn parse_xf(data: &[u8]) -> XlsbResult<CellFormat> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        // XF record structure (simplified)
        let font_id = binary::read_u16_le_at(data, 0)? as u32;
        let num_fmt_id = binary::read_u16_le_at(data, 2)? as u32;
        let fill_id = binary::read_u16_le_at(data, 4)? as u32;
        let border_id = binary::read_u16_le_at(data, 6)? as u32;

        // Parse alignment if present using the dedicated alignment parser
        let alignment = Alignment::parse(data, 0).ok().flatten();

        Ok(CellFormat {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            alignment,
        })
    }

    /// Get built-in number formats
    fn builtin_formats() -> HashMap<u32, String> {
        let mut formats = HashMap::new();
        formats.insert(0, "General".to_string());
        formats.insert(1, "0".to_string());
        formats.insert(2, "0.00".to_string());
        formats.insert(3, "#,##0".to_string());
        formats.insert(4, "#,##0.00".to_string());
        formats.insert(9, "0%".to_string());
        formats.insert(10, "0.00%".to_string());
        formats.insert(11, "0.00E+00".to_string());
        formats.insert(12, "# ?/?".to_string());
        formats.insert(13, "# ??/??".to_string());
        formats.insert(14, "mm-dd-yy".to_string());
        formats.insert(15, "d-mmm-yy".to_string());
        formats.insert(16, "d-mmm".to_string());
        formats.insert(17, "mmm-yy".to_string());
        formats.insert(18, "h:mm AM/PM".to_string());
        formats.insert(19, "h:mm:ss AM/PM".to_string());
        formats.insert(20, "h:mm".to_string());
        formats.insert(21, "h:mm:ss".to_string());
        formats.insert(22, "m/d/yy h:mm".to_string());
        formats.insert(37, "#,##0 ;(#,##0)".to_string());
        formats.insert(38, "#,##0 ;[Red](#,##0)".to_string());
        formats.insert(39, "#,##0.00;(#,##0.00)".to_string());
        formats.insert(40, "#,##0.00;[Red](#,##0.00)".to_string());
        formats.insert(45, "mm:ss".to_string());
        formats.insert(46, "[h]:mm:ss".to_string());
        formats.insert(47, "mmss.0".to_string());
        formats.insert(48, "##0.0E+0".to_string());
        formats.insert(49, "@".to_string());
        formats
    }

    /// Get cell format by index
    pub fn get_cell_format(&self, index: usize) -> Option<&CellFormat> {
        self.cell_xfs.get(index)
    }

    /// Get font by index
    pub fn get_font(&self, index: usize) -> Option<&Font> {
        self.fonts.get(index)
    }

    /// Get fill by index
    pub fn get_fill(&self, index: usize) -> Option<&Fill> {
        self.fills.get(index)
    }

    /// Get border by index
    pub fn get_border(&self, index: usize) -> Option<&Border> {
        self.borders.get(index)
    }

    /// Get number format by ID
    pub fn get_num_fmt(&self, id: u32) -> Option<&str> {
        self.num_fmts.get(&id).map(|s| s.as_str())
    }

    /// Check if a format code represents a date format
    pub fn is_date_format(&self, num_fmt_id: u32) -> bool {
        // Built-in date formats
        if matches!(num_fmt_id, 14..=22 | 27..=36 | 45..=47 | 50..=58) {
            return true;
        }

        // Custom format - check format code for date indicators
        if let Some(format_code) = self.get_num_fmt(num_fmt_id) {
            let format_lower = format_code.to_lowercase();
            // Simple heuristic: contains date/time indicators
            format_lower.contains('y')
                || format_lower.contains('m')
                || format_lower.contains('d')
                || format_lower.contains('h')
                || format_lower.contains('s')
        } else {
            false
        }
    }
}
