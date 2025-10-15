//! Cell representation for XLSB files

use crate::sheet::{Cell, CellValue};
use crate::ooxml::xlsb::records::CellRecord;

/// XLSB cell implementation
#[derive(Debug, Clone)]
pub struct XlsbCell {
    row: u32,
    col: u32,
    value: CellValue,
}

impl XlsbCell {
    /// Create a new XLSB cell
    pub fn new(row: u32, col: u32, value: CellValue) -> Self {
        XlsbCell { row, col, value }
    }

    /// Create cell from XLSB record
    pub fn from_record(record: &CellRecord, shared_strings: Option<&Vec<String>>) -> Option<Self> {
        let value = match &record.value {
            crate::ooxml::xlsb::records::CellValue::Blank => CellValue::Empty,
            crate::ooxml::xlsb::records::CellValue::Bool(b) => CellValue::Bool(*b),
            crate::ooxml::xlsb::records::CellValue::Error(e) => CellValue::Error(format!("Error {}", e)),
            crate::ooxml::xlsb::records::CellValue::Real(f) => CellValue::Float(*f),
            crate::ooxml::xlsb::records::CellValue::String(s) => CellValue::String(s.clone()),
            crate::ooxml::xlsb::records::CellValue::Isst(idx) => {
                if let Some(sst) = shared_strings {
                    if let Some(s) = sst.get(*idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::Error("Invalid SST index".to_string())
                    }
                } else {
                    CellValue::Error("SST not available".to_string())
                }
            }
        };

        Some(XlsbCell {
            row: record.row,
            col: record.col as u32,
            value,
        })
    }
}

impl Cell for XlsbCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.col
    }

    fn coordinate(&self) -> String {
        crate::ooxml::xlsb::utils::cell_reference(self.row, self.col)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        false // XLSB formulas not implemented yet
    }
}
