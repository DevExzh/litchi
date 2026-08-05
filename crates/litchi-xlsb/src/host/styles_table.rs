//! OOXML-facing adapter for the owner XLSB styles model and codec.
//!
//! The package-neutral values and Brt* parsing live in litchi_xlsb. This
//! module adds the host-owned alignment and border representations around the
//! canonical style values.

use crate::package::error::{Error, Result};
use crate::package::styles::{
    Alignment, Border, BorderSide, BorderStyle, HorizontalAlignment, VerticalAlignment,
};
pub use crate::styles::{Fill, Font, NumberFormat};
use std::collections::HashMap;

/// Cell format with the host alignment representation.
#[derive(Debug, Clone, Default)]
pub struct CellFormat {
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
    pub num_fmt_id: u32,
    pub alignment: Option<Alignment>,
}

/// XLSB styles table.
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
        let owner = crate::styles::Table::default();
        Self {
            fonts: vec![Font::default()],
            fills: vec![Fill::default()],
            borders: vec![Border::default()],
            num_fmts: owner.num_fmts,
            cell_xfs: vec![CellFormat::default()],
            cell_style_xfs: vec![CellFormat::default()],
        }
    }
}

impl StylesTable {
    /// Load styles from a styles.bin payload.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        crate::styles::read(bytes)
            .map(Self::from_owner)
            .map_err(map_owner_error)
    }

    fn from_owner(owner: crate::styles::Table) -> Self {
        Self {
            fonts: owner.fonts,
            fills: owner.fills,
            borders: owner.borders.into_iter().map(from_owner_border).collect(),
            num_fmts: owner.num_fmts,
            cell_xfs: owner
                .cell_xfs
                .into_iter()
                .map(from_owner_cell_format)
                .collect(),
            cell_style_xfs: owner
                .cell_style_xfs
                .into_iter()
                .map(from_owner_cell_format)
                .collect(),
        }
    }

    #[allow(dead_code)]
    pub(crate) fn parse_font(data: &[u8]) -> Result<Font> {
        crate::styles::parse_font(data).map_err(map_owner_error)
    }

    #[allow(dead_code)]
    pub(crate) fn parse_fill(data: &[u8]) -> Result<Fill> {
        crate::styles::parse_fill(data).map_err(map_owner_error)
    }

    #[allow(dead_code)]
    pub(crate) fn parse_direct_color(data: &[u8], offset: usize) -> Result<Option<u32>> {
        crate::styles::parse_direct_color(data, offset).map_err(map_owner_error)
    }

    #[allow(dead_code)]
    pub(crate) fn parse_border(data: &[u8]) -> Result<Border> {
        crate::styles::parse_border(data)
            .map(from_owner_border)
            .map_err(map_owner_error)
    }

    #[allow(dead_code)]
    pub(crate) fn parse_num_fmt(data: &[u8]) -> Result<(u32, String)> {
        crate::styles::parse_num_fmt(data).map_err(map_owner_error)
    }

    #[allow(dead_code)]
    pub(crate) fn parse_xf(data: &[u8]) -> Result<CellFormat> {
        crate::styles::parse_cell_format(data)
            .map(from_owner_cell_format)
            .map_err(map_owner_error)
    }

    /// Get a cell format by zero-based index.
    pub fn get_cell_format(&self, index: usize) -> Option<&CellFormat> {
        self.cell_xfs.get(index)
    }

    /// Get a font by zero-based index.
    pub fn get_font(&self, index: usize) -> Option<&Font> {
        self.fonts.get(index)
    }

    /// Get a fill by zero-based index.
    pub fn get_fill(&self, index: usize) -> Option<&Fill> {
        self.fills.get(index)
    }

    /// Get a border by zero-based index.
    pub fn get_border(&self, index: usize) -> Option<&Border> {
        self.borders.get(index)
    }

    /// Get a number-format code by identifier.
    pub fn get_num_fmt(&self, id: u32) -> Option<&str> {
        self.num_fmts.get(&id).map(String::as_str)
    }

    /// Check whether a built-in or custom format is date-like.
    pub fn is_date_format(&self, num_fmt_id: u32) -> bool {
        if matches!(num_fmt_id, 14..=22 | 27..=36 | 45..=47 | 50..=58) {
            return true;
        }

        if let Some(format_code) = self.get_num_fmt(num_fmt_id) {
            let format_lower = format_code.to_lowercase();
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

fn from_owner_cell_format(value: crate::styles::CellFormat) -> CellFormat {
    CellFormat {
        font_id: value.font_id,
        fill_id: value.fill_id,
        border_id: value.border_id,
        num_fmt_id: value.num_fmt_id,
        alignment: value.alignment.map(from_owner_alignment),
    }
}

fn from_owner_alignment(value: crate::styles::Alignment) -> Alignment {
    Alignment {
        horizontal: match value.horizontal {
            crate::styles::HorizontalAlignment::General => HorizontalAlignment::General,
            crate::styles::HorizontalAlignment::Left => HorizontalAlignment::Left,
            crate::styles::HorizontalAlignment::Center => HorizontalAlignment::Center,
            crate::styles::HorizontalAlignment::Right => HorizontalAlignment::Right,
            crate::styles::HorizontalAlignment::Fill => HorizontalAlignment::Fill,
            crate::styles::HorizontalAlignment::Justify => HorizontalAlignment::Justify,
            crate::styles::HorizontalAlignment::CenterContinuous => {
                HorizontalAlignment::CenterContinuous
            },
            crate::styles::HorizontalAlignment::Distributed => HorizontalAlignment::Distributed,
        },
        vertical: match value.vertical {
            crate::styles::VerticalAlignment::Top => VerticalAlignment::Top,
            crate::styles::VerticalAlignment::Center => VerticalAlignment::Center,
            crate::styles::VerticalAlignment::Bottom => VerticalAlignment::Bottom,
            crate::styles::VerticalAlignment::Justify => VerticalAlignment::Justify,
            crate::styles::VerticalAlignment::Distributed => VerticalAlignment::Distributed,
        },
        rotation: value.rotation,
        indent: value.indent,
        text_direction: value.text_direction,
        wrap_text: value.wrap_text,
        shrink_to_fit: value.shrink_to_fit,
    }
}

fn from_owner_border(value: crate::styles::Border) -> Border {
    Border {
        top: value.top.map(from_owner_side),
        bottom: value.bottom.map(from_owner_side),
        left: value.left.map(from_owner_side),
        right: value.right.map(from_owner_side),
        diagonal: value.diagonal.map(from_owner_side),
        vertical: value.vertical.map(from_owner_side),
        horizontal: value.horizontal.map(from_owner_side),
        diagonal_down: value.diagonal_down,
        diagonal_up: value.diagonal_up,
    }
}

fn from_owner_side(value: crate::styles::BorderSide) -> BorderSide {
    BorderSide {
        style: match value.style {
            crate::styles::BorderStyle::None => BorderStyle::None,
            crate::styles::BorderStyle::Thin => BorderStyle::Thin,
            crate::styles::BorderStyle::Medium => BorderStyle::Medium,
            crate::styles::BorderStyle::Dashed => BorderStyle::Dashed,
            crate::styles::BorderStyle::Dotted => BorderStyle::Dotted,
            crate::styles::BorderStyle::Thick => BorderStyle::Thick,
            crate::styles::BorderStyle::Double => BorderStyle::Double,
            crate::styles::BorderStyle::Hair => BorderStyle::Hair,
            crate::styles::BorderStyle::MediumDashed => BorderStyle::MediumDashed,
            crate::styles::BorderStyle::DashDot => BorderStyle::DashDot,
            crate::styles::BorderStyle::MediumDashDot => BorderStyle::MediumDashDot,
            crate::styles::BorderStyle::DashDotDot => BorderStyle::DashDotDot,
            crate::styles::BorderStyle::MediumDashDotDot => BorderStyle::MediumDashDotDot,
            crate::styles::BorderStyle::SlantDashDot => BorderStyle::SlantDashDot,
        },
        color: value.color,
    }
}

fn map_owner_error(error: crate::styles::Error) -> Error {
    match error {
        crate::styles::Error::Wire(error) => Error::Wire(error),
        crate::styles::Error::InvalidLength { expected, found } => {
            Error::InvalidLength { expected, found }
        },
        crate::styles::Error::Unrecognized { typ, val } => Error::Unrecognized { typ, val },
        crate::styles::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::styles::{HorizontalAlignment, VerticalAlignment};
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
            Err(Error::InvalidLength { .. })
        ));
        assert!(matches!(
            StylesTable::parse_fill(&[0; 67]),
            Err(Error::InvalidLength { .. })
        ));
        let invalid = [4, 0, 0, 0, 1, 2, 3, 4];
        assert!(matches!(
            StylesTable::parse_direct_color(&invalid, 0),
            Err(Error::Unrecognized { .. })
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
            let styles = StylesTable::from_bytes(part.blob()).unwrap();

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
            "/../../test-data/poi/test-data/spreadsheet/62815.xlsb"
        ));
        if !path.exists() {
            return;
        }
        let package = OpcPackage::from_reader(File::open(path).unwrap()).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/styles.bin").unwrap())
            .unwrap();
        let styles = StylesTable::from_bytes(part.blob()).unwrap();
        assert!(styles.num_fmts.keys().any(|id| *id >= 164));
    }
}
