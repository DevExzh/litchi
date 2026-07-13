//! XLSB styles table parsing
//!
//! This module provides the main StylesTable container for Excel Binary Workbook styles,
//! including fonts, fills, borders, number formats, and cell formats.
//! Reference: [MS-XLSB] Section 2.4 - Styles Part

use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::{XlsbRecordIter, record_types, wide_str_with_len};
use crate::xlsb::styles::{Alignment, Border};
use litchi_core::binary;
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
#[derive(Debug, Clone, Default)]
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
            cell_xfs: vec![CellFormat::default()],
            cell_style_xfs: vec![CellFormat::default()],
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
                    styles.fonts.push(Self::parse_font(data)?);
                },
                record_types::BEGIN_FILLS => {
                    in_fills = true;
                    styles.fills.clear(); // Clear default
                },
                record_types::END_FILLS => in_fills = false,
                record_types::FILL if in_fills => {
                    styles.fills.push(Self::parse_fill(data)?);
                },
                record_types::BEGIN_BORDERS => {
                    in_borders = true;
                    styles.borders.clear(); // Clear default
                },
                record_types::END_BORDERS => in_borders = false,
                record_types::BORDER if in_borders => {
                    styles.borders.push(Self::parse_border(data)?);
                },
                record_types::BEGIN_FMTS => in_fmts = true,
                record_types::END_FMTS => in_fmts = false,
                record_types::FMT if in_fmts => {
                    let (id, format_code) = Self::parse_num_fmt(data)?;
                    styles.num_fmts.insert(id, format_code);
                },
                record_types::BEGIN_CELL_XFS => {
                    in_cell_xfs = true;
                    styles.cell_xfs.clear();
                },
                record_types::END_CELL_XFS => in_cell_xfs = false,
                record_types::XF if in_cell_xfs => {
                    styles.cell_xfs.push(Self::parse_xf(data)?);
                },
                record_types::BEGIN_CELL_STYLE_XFS => {
                    in_cell_style_xfs = true;
                    styles.cell_style_xfs.clear();
                },
                record_types::END_CELL_STYLE_XFS => in_cell_style_xfs = false,
                record_types::XF if in_cell_style_xfs => {
                    styles.cell_style_xfs.push(Self::parse_xf(data)?);
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
        const FONT_NAME_OFFSET: usize = 21;
        if data.len() < FONT_NAME_OFFSET {
            return Err(XlsbError::InvalidLength {
                expected: FONT_NAME_OFFSET,
                found: data.len(),
            });
        }

        // Font height in twips (1/20 of a point)
        let height = binary::read_u16_le_at(data, 0)?;
        let size = height as f64 / 20.0;

        let flags = binary::read_u16_le_at(data, 2)?;
        let italic = (flags & 0x0002) != 0;
        let strike = (flags & 0x0008) != 0;
        let bold = binary::read_u16_le_at(data, 4)? >= 0x02BC;
        let underline = data[8] != 0;

        let color = Self::parse_direct_color(data, 12)?;

        // bFontScheme is at byte 20; this compact public model does not
        // currently expose theme font schemes.
        let (name, _) = wide_str_with_len(&data[FONT_NAME_OFFSET..])?;

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
        const FIXED_FILL_SIZE: usize = 68;
        if data.len() < FIXED_FILL_SIZE {
            return Err(XlsbError::InvalidLength {
                expected: FIXED_FILL_SIZE,
                found: data.len(),
            });
        }

        let pattern_type = binary::read_u32_le_at(data, 0)?;
        let (fg_color, bg_color) = if pattern_type == 0x28 {
            // Gradient fills use the trailing GradientStop array rather than
            // the two pattern colors.
            (None, None)
        } else {
            (
                Self::parse_direct_color(data, 4)?,
                Self::parse_direct_color(data, 12)?,
            )
        };

        Ok(Fill {
            pattern_type,
            fg_color,
            bg_color,
        })
    }

    /// Decode a direct `BrtColor` into the compact ARGB representation used
    /// by the public style model. Automatic, indexed, and theme colors cannot
    /// be resolved without their palette/theme context and are left as None.
    fn parse_direct_color(data: &[u8], offset: usize) -> XlsbResult<Option<u32>> {
        if offset + 8 > data.len() {
            return Err(XlsbError::InvalidLength {
                expected: offset + 8,
                found: data.len(),
            });
        }

        let flags = data[offset];
        let valid_rgb = flags & 1 != 0;
        let color_type = flags >> 1;
        if color_type != 2 {
            return Ok(None);
        }
        if !valid_rgb {
            return Err(XlsbError::Unrecognized {
                typ: "BrtColor".to_string(),
                val: "direct RGB color is not marked valid".to_string(),
            });
        }

        let red = u32::from(data[offset + 4]);
        let green = u32::from(data[offset + 5]);
        let blue = u32::from(data[offset + 6]);
        let alpha = u32::from(data[offset + 7]);
        Ok(Some((alpha << 24) | (red << 16) | (green << 8) | blue))
    }

    /// Parse border record using the dedicated border parser
    fn parse_border(data: &[u8]) -> XlsbResult<Border> {
        Border::parse(data)
    }

    /// Parse number format record
    fn parse_num_fmt(data: &[u8]) -> XlsbResult<(u32, String)> {
        if data.len() < 8 {
            return Err(XlsbError::InvalidLength {
                expected: 8,
                found: data.len(),
            });
        }

        let id = u32::from(binary::read_u16_le_at(data, 0)?);
        if !matches!(id, 5..=8 | 23..=26 | 41..=44 | 63..=66 | 164..=382) {
            return Err(XlsbError::Unrecognized {
                typ: "BrtFmt ifmt".to_string(),
                val: id.to_string(),
            });
        }
        let (format_code, consumed) = wide_str_with_len(&data[2..])?;
        let length = format_code.encode_utf16().count();
        if !(1..=255).contains(&length) {
            return Err(XlsbError::Unrecognized {
                typ: "BrtFmt stFmtCode length".to_string(),
                val: length.to_string(),
            });
        }
        if consumed + 2 != data.len() {
            return Err(XlsbError::Unrecognized {
                typ: "BrtFmt".to_string(),
                val: format!("{} trailing bytes", data.len() - consumed - 2),
            });
        }

        Ok((id, format_code))
    }

    /// Parse XF (cell format) record
    fn parse_xf(data: &[u8]) -> XlsbResult<CellFormat> {
        if data.len() < 16 {
            return Err(XlsbError::InvalidLength {
                expected: 16,
                found: data.len(),
            });
        }

        let font_id = binary::read_u16_le_at(data, 4)? as u32;
        let num_fmt_id = binary::read_u16_le_at(data, 2)? as u32;
        let fill_id = binary::read_u16_le_at(data, 6)? as u32;
        let border_id = binary::read_u16_le_at(data, 8)? as u32;

        // Parse alignment if present using the dedicated alignment parser
        let alignment = Alignment::parse(data, 0)?;

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::styles::{HorizontalAlignment, VerticalAlignment};
    use litchi_opc::{OpcPackage, PackURI};
    use std::fs::File;

    fn direct_color(argb: u32) -> [u8; 8] {
        [
            5,
            0,
            0,
            0,
            ((argb >> 16) & 0xFF) as u8,
            ((argb >> 8) & 0xFF) as u8,
            (argb & 0xFF) as u8,
            ((argb >> 24) & 0xFF) as u8,
        ]
    }

    #[test]
    fn parses_spec_layout_font_and_direct_color() {
        let mut data = Vec::new();
        data.extend_from_slice(&240u16.to_le_bytes());
        data.extend_from_slice(&0x000Au16.to_le_bytes());
        data.extend_from_slice(&700u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[1, 2, 0, 0]);
        data.extend_from_slice(&direct_color(0x80402010));
        data.push(0);
        data.extend_from_slice(&5u32.to_le_bytes());
        for unit in "Arial".encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }

        let font = StylesTable::parse_font(&data).unwrap();
        assert_eq!(font.name, "Arial");
        assert_eq!(font.size, 12.0);
        assert_eq!(font.color, Some(0x80402010));
        assert!(font.bold);
        assert!(font.italic);
        assert!(font.underline);
        assert!(font.strike);
    }

    #[test]
    fn parses_pattern_fill_colors_at_spec_offsets() {
        let mut data = vec![0u8; 68];
        data[0..4].copy_from_slice(&1u32.to_le_bytes());
        data[4..12].copy_from_slice(&direct_color(0xFFFF0000));
        data[12..20].copy_from_slice(&direct_color(0x4000FF00));

        let fill = StylesTable::parse_fill(&data).unwrap();
        assert_eq!(fill.pattern_type, 1);
        assert_eq!(fill.fg_color, Some(0xFFFF0000));
        assert_eq!(fill.bg_color, Some(0x4000FF00));
    }

    #[test]
    fn parses_xf_indices_and_alignment_at_spec_offsets() {
        let mut data = [0u8; 16];
        data[0..2].copy_from_slice(&0x1234u16.to_le_bytes());
        data[2..4].copy_from_slice(&165u16.to_le_bytes());
        data[4..6].copy_from_slice(&2u16.to_le_bytes());
        data[6..8].copy_from_slice(&3u16.to_le_bytes());
        data[8..10].copy_from_slice(&4u16.to_le_bytes());
        data[10] = 90;
        data[11] = 6;
        data[12] = 2 | (1 << 3) | 0x40;
        data[13] = 1 | (2 << 2);

        let format = StylesTable::parse_xf(&data).unwrap();
        assert_eq!(format.num_fmt_id, 165);
        assert_eq!(format.font_id, 2);
        assert_eq!(format.fill_id, 3);
        assert_eq!(format.border_id, 4);
        let alignment = format.alignment.unwrap();
        assert_eq!(alignment.horizontal, HorizontalAlignment::Center);
        assert_eq!(alignment.vertical, VerticalAlignment::Center);
        assert_eq!(alignment.rotation, 90);
        assert_eq!(alignment.indent, 6);
        assert_eq!(alignment.text_direction, 2);
        assert!(alignment.wrap_text);
        assert!(alignment.shrink_to_fit);
    }

    #[test]
    fn parses_two_byte_number_format_identifier() {
        let mut data = 166u16.to_le_bytes().to_vec();
        data.extend_from_slice(&7u32.to_le_bytes());
        for unit in "0.00000".encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        assert_eq!(
            StylesTable::parse_num_fmt(&data).unwrap(),
            (166, "0.00000".to_string())
        );
    }

    #[test]
    fn rejects_truncated_style_records_and_invalid_direct_color() {
        assert!(matches!(
            StylesTable::parse_font(&[0; 20]),
            Err(XlsbError::InvalidLength { .. })
        ));
        assert!(matches!(
            StylesTable::parse_fill(&[0; 67]),
            Err(XlsbError::InvalidLength { .. })
        ));
        let invalid = [4, 0, 0, 0, 1, 2, 3, 4];
        assert!(matches!(
            StylesTable::parse_direct_color(&invalid, 0),
            Err(XlsbError::Unrecognized { .. })
        ));
    }

    #[test]
    fn reads_styles_from_real_xlsb_fixtures() {
        for fixture in [
            "Simple.xlsb",
            "hyperlink.xlsb",
            "date.xlsb",
            "universal-content.xlsb",
            "comments.xlsb",
            "cond_format.xlsb",
        ] {
            let path = format!(
                "{}/../../test-data/ooxml/xlsb/{fixture}",
                env!("CARGO_MANIFEST_DIR")
            );
            let package = OpcPackage::from_reader(File::open(path).unwrap()).unwrap();
            let part = package
                .get_part(&PackURI::new("/xl/styles.bin").unwrap())
                .unwrap();
            let styles = StylesTable::from_reader(part.blob()).unwrap();

            assert!(!styles.fonts.is_empty(), "{fixture}");
            assert!(styles.fills.len() >= 2, "{fixture}");
            assert!(!styles.cell_xfs.is_empty(), "{fixture}");
            for format in &styles.cell_xfs {
                assert!((format.font_id as usize) < styles.fonts.len(), "{fixture}");
                assert!((format.fill_id as usize) < styles.fills.len(), "{fixture}");
                assert!(
                    (format.border_id as usize) < styles.borders.len(),
                    "{fixture}"
                );
            }

            if fixture == "universal-content.xlsb" {
                let default_font = styles.get_font(0).unwrap();
                assert_eq!(default_font.name, "Arial");
                assert_eq!(default_font.size, 10.0);
                assert!(default_font.color.is_none());
            }
        }
    }

    #[test]
    fn reads_custom_number_formats_from_poi_fixture_when_available() {
        let path = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../3rdparty/poi/test-data/spreadsheet/62815.xlsb"
        ));
        if !path.exists() {
            return;
        }
        let package = OpcPackage::from_reader(File::open(path).unwrap()).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/styles.bin").unwrap())
            .unwrap();
        let styles = StylesTable::from_reader(part.blob()).unwrap();
        assert!(styles.num_fmts.keys().any(|id| *id >= 164));
    }
}
