//! Cell representation for XLS files

use crate::ole::xls::records::{BoolErrValue, CellRecord, FormulaValue};
use crate::ole::xls::utils;
use crate::sheet::{Cell, CellValue};

/// XLS cell implementation
#[derive(Debug, Clone)]
pub struct XlsCell {
    row: u32,
    col: u32,
    value: CellValue,
    formula: Option<String>,
}

impl XlsCell {
    /// Create a new XLS cell
    pub fn new(row: u32, col: u32, value: CellValue) -> Self {
        XlsCell {
            row,
            col,
            value,
            formula: None,
        }
    }

    /// Create a cell with formula
    pub fn with_formula(row: u32, col: u32, value: CellValue, formula: String) -> Self {
        XlsCell {
            row,
            col,
            value,
            formula: Some(formula),
        }
    }

    /// Create cell from BIFF record
    pub fn from_record(record: &CellRecord, sst: Option<&[String]>) -> Option<Self> {
        let (row, col, value, formula) = match record {
            CellRecord::Blank { row, col, .. } => {
                (*row as u32, *col as u32, CellValue::Empty, None)
            },
            CellRecord::Number {
                row, col, value, ..
            } => (*row as u32, *col as u32, CellValue::Float(*value), None),
            CellRecord::Label {
                row, col, value, ..
            } => (
                *row as u32,
                *col as u32,
                CellValue::String(value.clone()),
                None,
            ),
            CellRecord::BoolErr {
                row, col, value, ..
            } => {
                let cell_value = match value {
                    BoolErrValue::Bool(b) => CellValue::Bool(*b),
                    BoolErrValue::Error(e) => CellValue::Error(format!("Error {}", e)),
                };
                (*row as u32, *col as u32, cell_value, None)
            },
            CellRecord::Rk {
                row, col, value, ..
            } => (*row as u32, *col as u32, CellValue::Float(*value), None),
            CellRecord::LabelSst {
                row,
                col,
                sst_index,
                ..
            } => {
                let cell_value = if let Some(sst) = sst {
                    if let Some(s) = sst.get(*sst_index as usize) {
                        CellValue::String(s.clone())
                    } else {
                        CellValue::Error(format!(
                            "Invalid SST index: {} (max: {})",
                            sst_index,
                            sst.len()
                        ))
                    }
                } else {
                    CellValue::Error("SST not available".to_string())
                };
                (*row as u32, *col as u32, cell_value, None)
            },
            CellRecord::Formula {
                row,
                col,
                value,
                formula,
                xf_index: _,
            } => {
                let cell_value = match value {
                    FormulaValue::Number(n) => CellValue::Float(*n),
                    FormulaValue::String(s) => CellValue::String(s.clone()),
                    FormulaValue::Bool(b) => CellValue::Bool(*b),
                    FormulaValue::Error(e) => CellValue::Error(format!("Error {}", e)),
                    FormulaValue::Empty => CellValue::Empty,
                };

                // For now, just store the raw formula bytes as a placeholder
                // A full implementation would parse the formula
                let formula_str = format!("Formula({} bytes)", formula.len());

                (*row as u32, *col as u32, cell_value, Some(formula_str))
            },
        };

        Some(XlsCell {
            row,
            col,
            value,
            formula,
        })
    }
}

impl Cell for XlsCell {
    fn row(&self) -> u32 {
        self.row
    }

    fn column(&self) -> u32 {
        self.col
    }

    fn coordinate(&self) -> String {
        utils::cell_reference(self.row, self.col)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        self.formula.is_some()
    }
}

// Implement Cell for &XlsCell to allow zero-copy reference returns
impl Cell for &XlsCell {
    fn row(&self) -> u32 {
        (*self).row()
    }

    fn column(&self) -> u32 {
        (*self).column()
    }

    fn coordinate(&self) -> String {
        (*self).coordinate()
    }

    fn value(&self) -> &CellValue {
        (*self).value()
    }

    fn is_formula(&self) -> bool {
        (*self).is_formula()
    }
}
