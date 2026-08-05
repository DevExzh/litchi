//! BIFF8 worksheet row and column layout records.

use crate::error::{Error, Result};
use crate::number_format::Formatting;
use crate::worksheet::layout::DEF_COL_WIDTH_RECORD_TYPE;
use std::collections::BTreeMap;

pub mod column;
pub mod row;

pub use column::Column;
pub use row::Row;

/// BIFF8 `ROW` record type.
pub(crate) const ROW_RECORD_TYPE: u16 = 0x0208;
/// BIFF8 `COLINFO` record type.
pub(crate) const COLINFO_RECORD_TYPE: u16 = 0x007d;

const MAX_COLINFO_RECORDS: usize = 255;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

/// Enforces the worksheet-level record collections while preserving record order.
pub(crate) struct Collector {
    rows: BTreeMap<u16, Row>,
    columns: Vec<Column>,
    last_row: Option<u16>,
    saw_default_column_width: bool,
    saw_columns: bool,
    columns_closed: bool,
    rows_started: bool,
}

impl Collector {
    pub(crate) fn new() -> Self {
        Self {
            rows: BTreeMap::new(),
            columns: Vec::new(),
            last_row: None,
            saw_default_column_width: false,
            saw_columns: false,
            columns_closed: false,
            rows_started: false,
        }
    }

    pub(crate) fn feed_record(
        &mut self,
        record_type: u16,
        data: &[u8],
        formatting: &Formatting,
    ) -> Result<()> {
        if self.saw_default_column_width
            && record_type != COLINFO_RECORD_TYPE
            && record_type != DEF_COL_WIDTH_RECORD_TYPE
        {
            self.columns_closed = true;
        }
        match record_type {
            DEF_COL_WIDTH_RECORD_TYPE => {
                self.saw_default_column_width = true;
                if self.saw_columns || self.rows_started {
                    return Err(invalid(
                        record_type,
                        "DefColWidth must begin the COLUMNS collection",
                    ));
                }
            },
            COLINFO_RECORD_TYPE => {
                if !self.saw_default_column_width {
                    return Err(invalid(
                        record_type,
                        "COLINFO records require a preceding DefColWidth",
                    ));
                }
                if self.columns_closed || self.rows_started {
                    return Err(invalid(
                        record_type,
                        "COLINFO records must remain in the COLUMNS collection",
                    ));
                }
                if self.columns.len() == MAX_COLINFO_RECORDS {
                    return Err(invalid(
                        record_type,
                        "worksheet contains more than 255 COLINFO records",
                    ));
                }
                let column = column::parse(data)?;
                formatting.validate_cell_xf(column.format_index())?;
                self.columns.push(column);
                self.saw_columns = true;
            },
            ROW_RECORD_TYPE => {
                let row = row::parse(data)?;
                if self.last_row.is_some_and(|last| row.row() <= last) {
                    return Err(invalid(
                        record_type,
                        "ROW records must have strictly increasing row indexes",
                    ));
                }
                if let Some(index) = row.format_index() {
                    formatting.validate_cell_xf(index)?;
                }
                self.last_row = Some(row.row());
                self.rows.insert(row.row(), row);
                self.rows_started = true;
            },
            _ => {},
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> (BTreeMap<u16, Row>, Vec<Column>) {
        (self.rows, self.columns)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn column_payload(first: u16, last: u16, flags: u16) -> [u8; 12] {
        let mut data = [0u8; 12];
        data[0..2].copy_from_slice(&first.to_le_bytes());
        data[2..4].copy_from_slice(&last.to_le_bytes());
        data[4..6].copy_from_slice(&2560u16.to_le_bytes());
        data[6..8].copy_from_slice(&0u16.to_le_bytes());
        data[8..10].copy_from_slice(&flags.to_le_bytes());
        data
    }

    fn row_payload(row: u16) -> [u8; 16] {
        let mut data = [0u8; 16];
        data[0..2].copy_from_slice(&row.to_le_bytes());
        data[2..4].copy_from_slice(&2u16.to_le_bytes());
        data[4..6].copy_from_slice(&5u16.to_le_bytes());
        data[6..8].copy_from_slice(&300u16.to_le_bytes());
        data[12..14].copy_from_slice(&0x0100u16.to_le_bytes());
        data
    }

    #[test]
    fn collector_enforces_columns_and_sorted_rows() {
        let formatting = Formatting::default();
        let mut collector = Collector::new();
        assert!(
            collector
                .feed_record(COLINFO_RECORD_TYPE, &column_payload(0, 0, 0), &formatting)
                .is_err()
        );
        collector
            .feed_record(DEF_COL_WIDTH_RECORD_TYPE, &8u16.to_le_bytes(), &formatting)
            .unwrap();
        collector
            .feed_record(COLINFO_RECORD_TYPE, &column_payload(0, 0, 0), &formatting)
            .unwrap();
        collector.feed_record(0x0200, &[], &formatting).unwrap();
        assert!(
            collector
                .feed_record(COLINFO_RECORD_TYPE, &column_payload(1, 1, 0), &formatting,)
                .is_err()
        );

        let mut collector = Collector::new();
        collector
            .feed_record(DEF_COL_WIDTH_RECORD_TYPE, &8u16.to_le_bytes(), &formatting)
            .unwrap();
        collector
            .feed_record(ROW_RECORD_TYPE, &row_payload(2), &formatting)
            .unwrap();
        assert!(
            collector
                .feed_record(COLINFO_RECORD_TYPE, &column_payload(1, 1, 0), &formatting)
                .is_err()
        );
        assert!(
            collector
                .feed_record(ROW_RECORD_TYPE, &row_payload(1), &formatting)
                .is_err()
        );
    }
}
