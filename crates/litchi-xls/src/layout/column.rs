//! BIFF8 `COLINFO` records and their worksheet-facing semantic view.

use super::{COLINFO_RECORD_TYPE, invalid, read_u16};
use crate::error::Result;

/// Formatting and display metadata for an inclusive worksheet column range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Column {
    first_column: u16,
    last_column: u16,
    width_256ths: u16,
    format_index: u16,
    hidden: bool,
    user_set: bool,
    best_fit: bool,
    phonetic: bool,
    outline_level: u8,
    collapsed: bool,
}

impl Column {
    #[must_use]
    pub fn first_column(&self) -> u16 {
        self.first_column
    }

    #[must_use]
    pub fn last_column(&self) -> u16 {
        self.last_column
    }

    /// Whether this range also defines formatting for newly exposed columns.
    ///
    /// BIFF8 uses column index `0x0100` as the default-column-formatting
    /// sentinel even though visible worksheet columns end at index 255.
    #[must_use]
    pub fn includes_default_column_formatting(&self) -> bool {
        self.last_column == 0x0100
    }

    #[must_use]
    pub fn width_256ths(&self) -> u16 {
        self.width_256ths
    }

    #[must_use]
    pub fn format_index(&self) -> u16 {
        self.format_index
    }

    #[must_use]
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    #[must_use]
    pub fn is_user_set(&self) -> bool {
        self.user_set
    }

    #[must_use]
    pub fn is_best_fit(&self) -> bool {
        self.best_fit
    }

    #[must_use]
    pub fn has_phonetic_guide(&self) -> bool {
        self.phonetic
    }

    #[must_use]
    pub fn outline_level(&self) -> u8 {
        self.outline_level
    }

    #[must_use]
    pub fn is_collapsed(&self) -> bool {
        self.collapsed
    }
}

pub(super) fn parse(data: &[u8]) -> Result<Column> {
    if data.len() != 12 {
        return Err(invalid(
            COLINFO_RECORD_TYPE,
            format!("COLINFO payload must be 12 bytes, found {}", data.len()),
        ));
    }
    let first_column = read_u16(data, 0);
    let last_column = read_u16(data, 2);
    let flags = read_u16(data, 8);
    if first_column > 0x0100 || last_column > 0x0100 || last_column < first_column {
        return Err(invalid(
            COLINFO_RECORD_TYPE,
            "COLINFO column range is invalid",
        ));
    }
    // MS-XLS 2.4.53 reserves bits 4..=7 and 13..=15. Bit 11 is
    // intentionally ignored because the specification marks it undefined.
    if flags & 0xe0f0 != 0 {
        return Err(invalid(
            COLINFO_RECORD_TYPE,
            "COLINFO reserved flag bits must be zero",
        ));
    }

    Ok(Column {
        first_column,
        last_column,
        width_256ths: read_u16(data, 4),
        format_index: read_u16(data, 6),
        hidden: flags & 0x0001 != 0,
        user_set: flags & 0x0002 != 0,
        best_fit: flags & 0x0004 != 0,
        phonetic: flags & 0x0008 != 0,
        outline_level: ((flags >> 8) & 0x0007) as u8,
        collapsed: flags & 0x1000 != 0,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn payload(first: u16, last: u16, flags: u16) -> [u8; 12] {
        let mut data = [0u8; 12];
        data[0..2].copy_from_slice(&first.to_le_bytes());
        data[2..4].copy_from_slice(&last.to_le_bytes());
        data[4..6].copy_from_slice(&2560u16.to_le_bytes());
        data[6..8].copy_from_slice(&0u16.to_le_bytes());
        data[8..10].copy_from_slice(&flags.to_le_bytes());
        data
    }

    #[test]
    fn parses_flags_and_range() {
        let column = parse(&payload(2, 4, 0x150f)).unwrap();
        assert_eq!((column.first_column(), column.last_column()), (2, 4));
        assert_eq!(column.width_256ths(), 2560);
        assert!(column.is_hidden());
        assert!(column.is_user_set());
        assert!(column.is_best_fit());
        assert!(column.has_phonetic_guide());
        assert_eq!(column.outline_level(), 5);
        assert!(column.is_collapsed());
    }

    #[test]
    fn accepts_default_formatting_sentinel() {
        let column = parse(&payload(0, 0x0100, 0x0002)).unwrap();
        assert_eq!((column.first_column(), column.last_column()), (0, 0x0100));
        assert!(column.includes_default_column_formatting());
    }

    #[test]
    fn rejects_malformed_records() {
        assert!(parse(&[0; 11]).is_err());
        assert!(parse(&payload(5, 4, 0)).is_err());
        assert!(parse(&payload(0, 0, 0x0010)).is_err());
    }
}
