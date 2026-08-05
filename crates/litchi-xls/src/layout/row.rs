//! BIFF8 `ROW` records and their worksheet-facing semantic view.

use super::{ROW_RECORD_TYPE, invalid, read_u16};
use crate::error::XlsResult;

/// Formatting and display metadata for one worksheet row.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Row {
    row: u16,
    first_cell_column: u16,
    last_cell_column_exclusive: u16,
    height_twips: u16,
    outline_level: u8,
    collapsed: bool,
    hidden: bool,
    custom_height: bool,
    formatted: bool,
    format_index: Option<u16>,
    thick_top_border: bool,
    thick_bottom_border: bool,
    phonetic: bool,
}

impl Row {
    pub fn row(&self) -> u16 {
        self.row
    }

    pub fn first_cell_column(&self) -> u16 {
        self.first_cell_column
    }

    pub fn last_cell_column_exclusive(&self) -> u16 {
        self.last_cell_column_exclusive
    }

    pub fn height_twips(&self) -> u16 {
        self.height_twips
    }

    pub fn outline_level(&self) -> u8 {
        self.outline_level
    }

    pub fn is_collapsed(&self) -> bool {
        self.collapsed
    }

    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    pub fn has_custom_height(&self) -> bool {
        self.custom_height
    }

    pub fn is_formatted(&self) -> bool {
        self.formatted
    }

    pub fn format_index(&self) -> Option<u16> {
        self.format_index
    }

    pub fn has_thick_top_border(&self) -> bool {
        self.thick_top_border
    }

    pub fn has_thick_bottom_border(&self) -> bool {
        self.thick_bottom_border
    }

    pub fn has_phonetic_guide(&self) -> bool {
        self.phonetic
    }
}

pub(super) fn parse(data: &[u8]) -> XlsResult<Row> {
    if data.len() != 16 {
        return Err(invalid(
            ROW_RECORD_TYPE,
            format!("ROW payload must be 16 bytes, found {}", data.len()),
        ));
    }
    let row = read_u16(data, 0);
    let first_cell_column = read_u16(data, 2);
    let last_cell_column_exclusive = read_u16(data, 4);
    let height_twips = read_u16(data, 6);
    let reserved1 = read_u16(data, 8);
    let flags = read_u16(data, 12);
    let format_flags = read_u16(data, 14);

    if first_cell_column > 255 {
        return Err(invalid(
            ROW_RECORD_TYPE,
            "ROW first cell column exceeds 255",
        ));
    }
    if last_cell_column_exclusive > 256 || last_cell_column_exclusive < first_cell_column {
        return Err(invalid(
            ROW_RECORD_TYPE,
            "ROW last cell column is outside its valid range",
        ));
    }
    if !(2..=8192).contains(&height_twips) {
        return Err(invalid(
            ROW_RECORD_TYPE,
            "ROW height must be between 2 and 8192 twips",
        ));
    }
    if reserved1 != 0 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW reserved1 must be zero"));
    }
    if flags & 0x0008 != 0 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW reserved2 must be zero"));
    }
    // MS-XLS 2.4.221 requires the complete reserved3 byte to be 0x01.
    if flags & 0xff00 != 0x0100 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW reserved3 must be 0x01"));
    }

    let formatted = flags & 0x0080 != 0;
    Ok(Row {
        row,
        first_cell_column,
        last_cell_column_exclusive,
        height_twips,
        outline_level: (flags & 0x0007) as u8,
        collapsed: flags & 0x0010 != 0,
        hidden: flags & 0x0020 != 0,
        custom_height: flags & 0x0040 != 0,
        formatted,
        format_index: formatted.then_some(format_flags & 0x0fff),
        thick_top_border: format_flags & 0x1000 != 0,
        thick_bottom_border: format_flags & 0x2000 != 0,
        phonetic: format_flags & 0x4000 != 0,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn payload(row: u16, flags: u16, format_flags: u16) -> [u8; 16] {
        let mut data = [0u8; 16];
        data[0..2].copy_from_slice(&row.to_le_bytes());
        data[2..4].copy_from_slice(&2u16.to_le_bytes());
        data[4..6].copy_from_slice(&5u16.to_le_bytes());
        data[6..8].copy_from_slice(&300u16.to_le_bytes());
        data[12..14].copy_from_slice(&flags.to_le_bytes());
        data[14..16].copy_from_slice(&format_flags.to_le_bytes());
        data
    }

    #[test]
    fn parses_flags_and_conditional_format() {
        let row = parse(&payload(7, 0x01f5, 0x7234)).unwrap();
        assert_eq!(row.row(), 7);
        assert_eq!(row.height_twips(), 300);
        assert_eq!(row.outline_level(), 5);
        assert!(row.is_collapsed());
        assert!(row.is_hidden());
        assert!(row.has_custom_height());
        assert!(row.is_formatted());
        assert_eq!(row.format_index(), Some(0x0234));
        assert!(row.has_thick_top_border());
        assert!(row.has_thick_bottom_border());
        assert!(row.has_phonetic_guide());
    }

    #[test]
    fn rejects_malformed_records() {
        assert!(parse(&[0; 15]).is_err());
        assert!(parse(&payload(0, 0, 0)).is_err());
        assert!(parse(&payload(0, 0x0200, 0)).is_err());
    }

    #[test]
    fn accepts_unused_format_bits() {
        let row = parse(&payload(0, 0x0100, 0x8000)).unwrap();
        assert!(!row.is_formatted());
        assert_eq!(row.format_index(), None);
    }
}
