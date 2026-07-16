//! BIFF8 worksheet row and column layout records.

use crate::xls::error::{XlsError, XlsResult};
use crate::xls::number_format::XlsFormatting;
use std::collections::BTreeMap;

/// BIFF8 `ROW` record type.
pub(crate) const ROW_RECORD_TYPE: u16 = 0x0208;
/// BIFF8 `COLINFO` record type.
pub(crate) const COLINFO_RECORD_TYPE: u16 = 0x007d;

const MAX_COLINFO_RECORDS: usize = 255;

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

/// Formatting and display metadata for one worksheet row.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsRowLayout {
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

impl XlsRowLayout {
    pub fn row(&self) -> u16 { self.row }
    pub fn first_cell_column(&self) -> u16 { self.first_cell_column }
    pub fn last_cell_column_exclusive(&self) -> u16 { self.last_cell_column_exclusive }
    pub fn height_twips(&self) -> u16 { self.height_twips }
    pub fn outline_level(&self) -> u8 { self.outline_level }
    pub fn is_collapsed(&self) -> bool { self.collapsed }
    pub fn is_hidden(&self) -> bool { self.hidden }
    pub fn has_custom_height(&self) -> bool { self.custom_height }
    pub fn is_formatted(&self) -> bool { self.formatted }
    pub fn format_index(&self) -> Option<u16> { self.format_index }
    pub fn has_thick_top_border(&self) -> bool { self.thick_top_border }
    pub fn has_thick_bottom_border(&self) -> bool { self.thick_bottom_border }
    pub fn has_phonetic_guide(&self) -> bool { self.phonetic }
}

/// Formatting and display metadata for an inclusive worksheet column range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsColumnLayout {
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

impl XlsColumnLayout {
    pub fn first_column(&self) -> u16 { self.first_column }
    pub fn last_column(&self) -> u16 { self.last_column }
    /// Whether this range also defines formatting for newly exposed columns.
    ///
    /// BIFF8 uses column index `0x0100` as the default-column-formatting
    /// sentinel even though visible worksheet columns end at index 255.
    pub fn includes_default_column_formatting(&self) -> bool { self.last_column == 0x0100 }
    pub fn width_256ths(&self) -> u16 { self.width_256ths }
    pub fn format_index(&self) -> u16 { self.format_index }
    pub fn is_hidden(&self) -> bool { self.hidden }
    pub fn is_user_set(&self) -> bool { self.user_set }
    pub fn is_best_fit(&self) -> bool { self.best_fit }
    pub fn has_phonetic_guide(&self) -> bool { self.phonetic }
    pub fn outline_level(&self) -> u8 { self.outline_level }
    pub fn is_collapsed(&self) -> bool { self.collapsed }
}

fn parse_row(data: &[u8]) -> XlsResult<XlsRowLayout> {
    if data.len() != 16 {
        return Err(invalid(ROW_RECORD_TYPE, format!("ROW payload must be 16 bytes, found {}", data.len())));
    }
    let row = read_u16(data, 0);
    let first_cell_column = read_u16(data, 2);
    let last_cell_column_exclusive = read_u16(data, 4);
    let height_twips = read_u16(data, 6);
    let reserved1 = read_u16(data, 8);
    let flags = read_u16(data, 12);
    let format_flags = read_u16(data, 14);

    if first_cell_column > 255 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW first cell column exceeds 255"));
    }
    if last_cell_column_exclusive > 256 || last_cell_column_exclusive < first_cell_column {
        return Err(invalid(ROW_RECORD_TYPE, "ROW last cell column is outside its valid range"));
    }
    if !(2..=8192).contains(&height_twips) {
        return Err(invalid(ROW_RECORD_TYPE, "ROW height must be between 2 and 8192 twips"));
    }
    if reserved1 != 0 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW reserved1 must be zero"));
    }
    if flags & 0x0008 != 0 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW reserved2 must be zero"));
    }
    if flags & 0x0100 == 0 {
        return Err(invalid(ROW_RECORD_TYPE, "ROW reserved3 must be one"));
    }

    let formatted = flags & 0x0080 != 0;
    Ok(XlsRowLayout {
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

fn parse_col_info(data: &[u8]) -> XlsResult<XlsColumnLayout> {
    if data.len() != 12 {
        return Err(invalid(COLINFO_RECORD_TYPE, format!("COLINFO payload must be 12 bytes, found {}", data.len())));
    }
    let first_column = read_u16(data, 0);
    let last_column = read_u16(data, 2);
    let flags = read_u16(data, 8);
    if first_column > 0x0100 || last_column > 0x0100 || last_column < first_column {
        return Err(invalid(COLINFO_RECORD_TYPE, "COLINFO column range is invalid"));
    }
    if flags & 0xe0f0 != 0 {
        return Err(invalid(COLINFO_RECORD_TYPE, "COLINFO reserved flag bits must be zero"));
    }

    Ok(XlsColumnLayout {
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

/// Enforces the worksheet-level record collections while preserving record order.
pub(crate) struct LayoutCollector {
    rows: BTreeMap<u16, XlsRowLayout>,
    columns: Vec<XlsColumnLayout>,
    last_row: Option<u16>,
    saw_columns: bool,
    columns_closed: bool,
}

impl LayoutCollector {
    pub(crate) fn new() -> Self {
        Self {
            rows: BTreeMap::new(),
            columns: Vec::new(),
            last_row: None,
            saw_columns: false,
            columns_closed: false,
        }
    }

    pub(crate) fn feed_record(
        &mut self,
        record_type: u16,
        data: &[u8],
        formatting: &XlsFormatting,
    ) -> XlsResult<()> {
        if self.saw_columns && record_type != COLINFO_RECORD_TYPE {
            self.columns_closed = true;
        }
        match record_type {
            COLINFO_RECORD_TYPE => {
                if self.columns_closed {
                    return Err(invalid(record_type, "COLINFO records must be contiguous"));
                }
                if self.columns.len() == MAX_COLINFO_RECORDS {
                    return Err(invalid(record_type, "worksheet contains more than 255 COLINFO records"));
                }
                let column = parse_col_info(data)?;
                formatting.validate_cell_xf(column.format_index())?;
                self.columns.push(column);
                self.saw_columns = true;
            }
            ROW_RECORD_TYPE => {
                let row = parse_row(data)?;
                if self.last_row.is_some_and(|last| row.row() <= last) {
                    return Err(invalid(record_type, "ROW records must have strictly increasing row indexes"));
                }
                if let Some(index) = row.format_index() {
                    formatting.validate_cell_xf(index)?;
                }
                self.last_row = Some(row.row());
                self.rows.insert(row.row(), row);
            }
            _ => {}
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> (BTreeMap<u16, XlsRowLayout>, Vec<XlsColumnLayout>) {
        (self.rows, self.columns)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn row_payload(row: u16, flags: u16, format_flags: u16) -> [u8; 16] {
        let mut data = [0u8; 16];
        data[0..2].copy_from_slice(&row.to_le_bytes());
        data[2..4].copy_from_slice(&2u16.to_le_bytes());
        data[4..6].copy_from_slice(&5u16.to_le_bytes());
        data[6..8].copy_from_slice(&300u16.to_le_bytes());
        data[12..14].copy_from_slice(&flags.to_le_bytes());
        data[14..16].copy_from_slice(&format_flags.to_le_bytes());
        data
    }

    fn col_payload(first: u16, last: u16, flags: u16) -> [u8; 12] {
        let mut data = [0u8; 12];
        data[0..2].copy_from_slice(&first.to_le_bytes());
        data[2..4].copy_from_slice(&last.to_le_bytes());
        data[4..6].copy_from_slice(&2560u16.to_le_bytes());
        data[6..8].copy_from_slice(&0u16.to_le_bytes());
        data[8..10].copy_from_slice(&flags.to_le_bytes());
        data
    }

    #[test]
    fn parses_row_layout_flags_and_conditional_format() {
        let row = parse_row(&row_payload(7, 0x01f5, 0x7234)).unwrap();
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
    fn parses_poi_column_layout_fixture_shape() {
        let col = parse_col_info(&col_payload(2, 4, 0x150f)).unwrap();
        assert_eq!((col.first_column(), col.last_column()), (2, 4));
        assert_eq!(col.width_256ths(), 2560);
        assert!(col.is_hidden());
        assert!(col.is_user_set());
        assert!(col.is_best_fit());
        assert!(col.has_phonetic_guide());
        assert_eq!(col.outline_level(), 5);
        assert!(col.is_collapsed());
    }

    #[test]
    fn accepts_col256u_default_formatting_sentinel() {
        let col = parse_col_info(&col_payload(0, 0x0100, 0x0002)).unwrap();
        assert_eq!((col.first_column(), col.last_column()), (0, 0x0100));
        assert!(col.includes_default_column_formatting());
    }

    #[test]
    fn rejects_malformed_layout_records() {
        assert!(parse_row(&[0; 15]).is_err());
        assert!(parse_row(&row_payload(0, 0, 0)).is_err());
        assert!(parse_col_info(&col_payload(5, 4, 0)).is_err());
        assert!(parse_col_info(&col_payload(0, 0, 0x0010)).is_err());
    }

    #[test]
    fn collector_rejects_noncontiguous_columns_and_unsorted_rows() {
        let formatting = XlsFormatting::default();
        let mut collector = LayoutCollector::new();
        collector.feed_record(COLINFO_RECORD_TYPE, &col_payload(0, 0, 0), &formatting).unwrap();
        collector.feed_record(0x0200, &[], &formatting).unwrap();
        assert!(collector.feed_record(COLINFO_RECORD_TYPE, &col_payload(1, 1, 0), &formatting).is_err());

        let mut collector = LayoutCollector::new();
        collector.feed_record(ROW_RECORD_TYPE, &row_payload(2, 0x0100, 0), &formatting).unwrap();
        assert!(collector.feed_record(ROW_RECORD_TYPE, &row_payload(1, 0x0100, 0), &formatting).is_err());
    }
}
