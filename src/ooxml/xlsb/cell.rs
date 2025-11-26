//! Cell representation for XLSB files

use crate::ooxml::xlsb::records::CellRecord;
use crate::sheet::{Cell, CellValue};

/// XLSB cell implementation
///
/// Fields are ordered to minimize padding and optimize cache utilization.
/// Layout: CellValue (largest), u32 fields, then bool.
#[derive(Debug, Clone)]
pub struct XlsbCell {
    /// Cell value (largest field, aligned first)
    value: CellValue,
    /// Row index (0-based)
    row: u32,
    /// Column index (0-based)
    col: u32,
    /// Track whether this cell came from a formula record
    is_formula: bool,
}

impl XlsbCell {
    /// Create a new XLSB cell
    pub fn new(row: u32, col: u32, value: CellValue) -> Self {
        XlsbCell {
            row,
            col,
            value,
            is_formula: false,
        }
    }

    /// Create a new XLSB cell from a formula
    pub fn new_formula(row: u32, col: u32, value: CellValue) -> Self {
        XlsbCell {
            row,
            col,
            value,
            is_formula: true,
        }
    }

    /// Create cell from XLSB record
    pub fn from_record(record: &CellRecord, shared_strings: Option<&Vec<String>>) -> Option<Self> {
        let (value, is_formula) = match &record.value {
            crate::ooxml::xlsb::records::CellValue::Blank => (CellValue::Empty, false),
            crate::ooxml::xlsb::records::CellValue::Bool(b) => (CellValue::Bool(*b), false),
            crate::ooxml::xlsb::records::CellValue::Error(e) => {
                // Convert error code to Excel error string
                let error_str = match e {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => "#ERR!",
                };
                (CellValue::Error(error_str.to_string()), false)
            },
            crate::ooxml::xlsb::records::CellValue::Real(f) => (CellValue::Float(*f), false),
            crate::ooxml::xlsb::records::CellValue::String(s) => {
                (CellValue::String(s.clone()), false)
            },
            crate::ooxml::xlsb::records::CellValue::Isst(idx) => {
                let val = if let Some(sst) = shared_strings {
                    if let Some(s) = sst.get(*idx as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::Error("Invalid SST index".to_string())
                    }
                } else {
                    CellValue::Error("SST not available".to_string())
                };
                (val, false)
            },
            crate::ooxml::xlsb::records::CellValue::Formula { value, formula: _ } => {
                // Extract the cached value from the formula
                // Formula parsing from bytes is complex and formula bytes are in a binary RPN format
                // For now, we use the cached value which is sufficient for most use cases
                (Self::extract_formula_value(value, shared_strings), true)
            },
        };

        Some(XlsbCell {
            row: record.row,
            col: record.col as u32,
            value,
            is_formula,
        })
    }

    /// Extract value from formula cached value
    fn extract_formula_value(
        formula_value: &crate::ooxml::xlsb::records::CellValue,
        shared_strings: Option<&Vec<String>>,
    ) -> CellValue {
        match formula_value {
            crate::ooxml::xlsb::records::CellValue::Blank => CellValue::Empty,
            crate::ooxml::xlsb::records::CellValue::Bool(b) => CellValue::Bool(*b),
            crate::ooxml::xlsb::records::CellValue::Error(e) => {
                let error_str = match e {
                    0x00 => "#NULL!",
                    0x07 => "#DIV/0!",
                    0x0F => "#VALUE!",
                    0x17 => "#REF!",
                    0x1D => "#NAME?",
                    0x24 => "#NUM!",
                    0x2A => "#N/A",
                    0x2B => "#GETTING_DATA",
                    _ => "#ERR!",
                };
                CellValue::Error(error_str.to_string())
            },
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
            },
            crate::ooxml::xlsb::records::CellValue::Formula { value, formula: _ } => {
                // Recursive formula values (shouldn't happen, but handle it)
                Self::extract_formula_value(value, shared_strings)
            },
        }
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
        self.is_formula
    }
}
