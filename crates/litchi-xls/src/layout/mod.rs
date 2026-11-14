//! BIFF8 worksheet row and column layout records.

use crate::error::{Error, Result};
use crate::leniency::ToleranceReport;
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
        tolerance: &mut ToleranceReport,
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
                formatting.check_cell_xf(
                    column.format_index(),
                    u32::from(column.first_column()),
                    tolerance,
                );
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
                    formatting.check_cell_xf(index, u32::from(row.row()), tolerance);
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
        let mut tolerance = ToleranceReport::default();
        let mut collector = Collector::new();
        assert!(
            collector
                .feed_record(
                    COLINFO_RECORD_TYPE,
                    &column_payload(0, 0, 0),
                    &formatting,
                    &mut tolerance
                )
                .is_err()
        );
        collector
            .feed_record(
                DEF_COL_WIDTH_RECORD_TYPE,
                &8u16.to_le_bytes(),
                &formatting,
                &mut tolerance,
            )
            .unwrap();
        collector
            .feed_record(
                COLINFO_RECORD_TYPE,
                &column_payload(0, 0, 0),
                &formatting,
                &mut tolerance,
            )
            .unwrap();
        collector
            .feed_record(0x0200, &[], &formatting, &mut tolerance)
            .unwrap();
        assert!(
            collector
                .feed_record(
                    COLINFO_RECORD_TYPE,
                    &column_payload(1, 1, 0),
                    &formatting,
                    &mut tolerance,
                )
                .is_err()
        );

        let mut collector = Collector::new();
        collector
            .feed_record(
                DEF_COL_WIDTH_RECORD_TYPE,
                &8u16.to_le_bytes(),
                &formatting,
                &mut tolerance,
            )
            .unwrap();
        collector
            .feed_record(
                ROW_RECORD_TYPE,
                &row_payload(2),
                &formatting,
                &mut tolerance,
            )
            .unwrap();
        assert!(
            collector
                .feed_record(
                    COLINFO_RECORD_TYPE,
                    &column_payload(1, 1, 0),
                    &formatting,
                    &mut tolerance,
                )
                .is_err()
        );
        assert!(
            collector
                .feed_record(
                    ROW_RECORD_TYPE,
                    &row_payload(1),
                    &formatting,
                    &mut tolerance
                )
                .is_err()
        );
    }
}

#[cfg(test)]
mod collection_order_tests {
    use super::*;

    /// Feeds one record per letter until the first error: `D` DefColWidth,
    /// `C` the next COLINFO, `R` the next ROW, `O` an unrelated record. The
    /// answer lists `.` per accepted record, then the error text or the
    /// collected row and column counts.
    fn outcome(letters: &str) -> String {
        let formatting = Formatting::default();
        let mut tolerance = ToleranceReport::default();
        let mut collector = Collector::new();
        let mut columns = 0u16;
        let mut rows = 0u16;
        let mut answers = String::new();
        for letter in letters.chars() {
            let (record_type, data) = match letter {
                'D' => (DEF_COL_WIDTH_RECORD_TYPE, 8u16.to_le_bytes().to_vec()),
                'C' => {
                    let mut data = [0u8; 12];
                    data[0..2].copy_from_slice(&columns.to_le_bytes());
                    data[2..4].copy_from_slice(&columns.to_le_bytes());
                    data[4..6].copy_from_slice(&2560u16.to_le_bytes());
                    columns += 1;
                    (COLINFO_RECORD_TYPE, data.to_vec())
                },
                'R' => {
                    let mut data = [0u8; 16];
                    data[0..2].copy_from_slice(&rows.to_le_bytes());
                    data[4..6].copy_from_slice(&5u16.to_le_bytes());
                    data[6..8].copy_from_slice(&300u16.to_le_bytes());
                    data[12..14].copy_from_slice(&0x0100u16.to_le_bytes());
                    rows += 1;
                    (ROW_RECORD_TYPE, data.to_vec())
                },
                'O' => (0x0200, Vec::new()),
                other => panic!("unknown test record letter {other}"),
            };
            if let Err(error) =
                collector.feed_record(record_type, &data, &formatting, &mut tolerance)
            {
                return format!("{letters} => {answers} {error}");
            }
            answers.push('.');
        }
        let (rows, columns) = collector.finish();
        format!(
            "{letters} => {answers} ok rows={} columns={}",
            rows.len(),
            columns.len()
        )
    }

    /// The COLUMNS and ROW collections are fed every order of up to three of
    /// their records, and longer orders that reopen or interleave them.
    #[test]
    fn layout_collections_accept_and_reject_record_orders() {
        let mut cases = vec![String::new()];
        let mut previous = vec![String::new()];
        for _ in 0..3 {
            let mut next = Vec::new();
            for prefix in &previous {
                for letter in ['D', 'C', 'R', 'O'] {
                    next.push(format!("{prefix}{letter}"));
                }
            }
            cases.extend(next.iter().cloned());
            previous = next;
        }
        cases.extend(
            [
                "DCCOR", "DCODC", "DODOC", "DODCR", "DDCC", "DCDC", "DCRR", "ODCOR", "RRD", "RRC",
                "RORC", "DORRO", "DDOD", "DCOOD",
            ]
            .map(str::to_string),
        );
        let actual: Vec<String> = cases.iter().map(|case| outcome(case)).collect();
        assert_eq!(actual, LAYOUT_ORDER_EXPECTED);
    }

    const LAYOUT_ORDER_EXPECTED: &[&str] = &[
        " =>  ok rows=0 columns=0",
        "D => . ok rows=0 columns=0",
        "C =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "R => . ok rows=1 columns=0",
        "O => . ok rows=0 columns=0",
        "DD => .. ok rows=0 columns=0",
        "DC => .. ok rows=0 columns=1",
        "DR => .. ok rows=1 columns=0",
        "DO => .. ok rows=0 columns=0",
        "CD =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CC =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CR =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CO =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RD => . Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RC => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RR => .. ok rows=2 columns=0",
        "RO => .. ok rows=1 columns=0",
        "OD => .. ok rows=0 columns=0",
        "OC => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "OR => .. ok rows=1 columns=0",
        "OO => .. ok rows=0 columns=0",
        "DDD => ... ok rows=0 columns=0",
        "DDC => ... ok rows=0 columns=1",
        "DDR => ... ok rows=1 columns=0",
        "DDO => ... ok rows=0 columns=0",
        "DCD => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "DCC => ... ok rows=0 columns=2",
        "DCR => ... ok rows=1 columns=1",
        "DCO => ... ok rows=0 columns=1",
        "DRD => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "DRC => .. Invalid record 0x007D: COLINFO records must remain in the COLUMNS collection",
        "DRR => ... ok rows=2 columns=0",
        "DRO => ... ok rows=1 columns=0",
        "DOD => ... ok rows=0 columns=0",
        "DOC => .. Invalid record 0x007D: COLINFO records must remain in the COLUMNS collection",
        "DOR => ... ok rows=1 columns=0",
        "DOO => ... ok rows=0 columns=0",
        "CDD =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CDC =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CDR =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CDO =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CCD =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CCC =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CCR =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CCO =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CRD =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CRC =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CRR =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "CRO =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "COD =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "COC =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "COR =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "COO =>  Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RDD => . Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RDC => . Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RDR => . Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RDO => . Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RCD => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RCC => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RCR => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RCO => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RRD => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RRC => .. Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RRR => ... ok rows=3 columns=0",
        "RRO => ... ok rows=2 columns=0",
        "ROD => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "ROC => .. Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "ROR => ... ok rows=2 columns=0",
        "ROO => ... ok rows=1 columns=0",
        "ODD => ... ok rows=0 columns=0",
        "ODC => ... ok rows=0 columns=1",
        "ODR => ... ok rows=1 columns=0",
        "ODO => ... ok rows=0 columns=0",
        "OCD => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "OCC => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "OCR => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "OCO => . Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "ORD => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "ORC => .. Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "ORR => ... ok rows=2 columns=0",
        "ORO => ... ok rows=1 columns=0",
        "OOD => ... ok rows=0 columns=0",
        "OOC => .. Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "OOR => ... ok rows=1 columns=0",
        "OOO => ... ok rows=0 columns=0",
        "DCCOR => ..... ok rows=1 columns=2",
        "DCODC => ... Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "DODOC => .... Invalid record 0x007D: COLINFO records must remain in the COLUMNS collection",
        "DODCR => ... Invalid record 0x007D: COLINFO records must remain in the COLUMNS collection",
        "DDCC => .... ok rows=0 columns=2",
        "DCDC => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "DCRR => .... ok rows=2 columns=1",
        "ODCOR => ..... ok rows=1 columns=1",
        "RRD => .. Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
        "RRC => .. Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "RORC => ... Invalid record 0x007D: COLINFO records require a preceding DefColWidth",
        "DORRO => ..... ok rows=2 columns=0",
        "DDOD => .... ok rows=0 columns=0",
        "DCOOD => .... Invalid record 0x0055: DefColWidth must begin the COLUMNS collection",
    ];
}
