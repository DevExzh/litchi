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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole::xls::records::{BoolErrValue, CellRecord, FormulaValue};

    #[test]
    fn test_new_cell() {
        let cell = XlsCell::new(0, 0, CellValue::String("Test".to_string()));
        assert_eq!(cell.row(), 0);
        assert_eq!(cell.column(), 0);
        assert_eq!(cell.coordinate(), "A1");
        assert!(!cell.is_formula());
    }

    #[test]
    fn test_cell_with_formula() {
        let cell = XlsCell::with_formula(1, 2, CellValue::Float(42.0), "=A1+B1".to_string());
        assert_eq!(cell.row(), 1);
        assert_eq!(cell.column(), 2);
        assert_eq!(cell.coordinate(), "C2");
        assert!(cell.is_formula());
    }

    #[test]
    fn test_cell_coordinate_various_positions() {
        let cell1 = XlsCell::new(0, 0, CellValue::Empty);
        assert_eq!(cell1.coordinate(), "A1");

        let cell2 = XlsCell::new(9, 25, CellValue::Empty);
        assert_eq!(cell2.coordinate(), "Z10");

        let cell3 = XlsCell::new(99, 26, CellValue::Empty);
        assert_eq!(cell3.coordinate(), "AA100");
    }

    #[test]
    fn test_from_record_blank() {
        let record = CellRecord::Blank {
            row: 5,
            col: 3,
            xf_index: 0,
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert_eq!(cell.row(), 5);
        assert_eq!(cell.column(), 3);
        assert_eq!(cell.coordinate(), "D6");
        assert!(matches!(cell.value(), CellValue::Empty));
    }

    #[test]
    fn test_from_record_number() {
        let record = CellRecord::Number {
            row: 0,
            col: 0,
            xf_index: 0,
            value: 123.456,
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert_eq!(cell.row(), 0);
        assert_eq!(cell.column(), 0);
        if let CellValue::Float(v) = cell.value() {
            assert!((v - 123.456).abs() < 0.001);
        } else {
            panic!("Expected Float value");
        }
    }

    #[test]
    fn test_from_record_label() {
        let record = CellRecord::Label {
            row: 2,
            col: 1,
            xf_index: 0,
            value: "Hello World".to_string(),
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert_eq!(cell.row(), 2);
        assert_eq!(cell.column(), 1);
        assert!(matches!(cell.value(), CellValue::String(s) if s == "Hello World"));
    }

    #[test]
    fn test_from_record_bool_err_bool() {
        let record = CellRecord::BoolErr {
            row: 0,
            col: 0,
            xf_index: 0,
            value: BoolErrValue::Bool(true),
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::Bool(true)));
    }

    #[test]
    fn test_from_record_bool_err_error() {
        let record = CellRecord::BoolErr {
            row: 0,
            col: 0,
            xf_index: 0,
            value: BoolErrValue::Error(7),
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::Error(_)));
    }

    #[test]
    fn test_from_record_rk() {
        let record = CellRecord::Rk {
            row: 10,
            col: 5,
            xf_index: 0,
            value: 100.5,
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert_eq!(cell.row(), 10);
        assert_eq!(cell.column(), 5);
        if let CellValue::Float(v) = cell.value() {
            assert!((v - 100.5).abs() < 0.001);
        } else {
            panic!("Expected Float value");
        }
    }

    #[test]
    fn test_from_record_label_sst_with_valid_sst() {
        let sst = vec![
            "First".to_string(),
            "Second".to_string(),
            "Third".to_string(),
        ];
        let record = CellRecord::LabelSst {
            row: 0,
            col: 0,
            xf_index: 0,
            sst_index: 1,
        };
        let cell = XlsCell::from_record(&record, Some(&sst)).unwrap();
        assert!(matches!(cell.value(), CellValue::String(s) if s == "Second"));
    }

    #[test]
    fn test_from_record_label_sst_invalid_index() {
        let sst = vec!["First".to_string()];
        let record = CellRecord::LabelSst {
            row: 0,
            col: 0,
            xf_index: 0,
            sst_index: 5,
        };
        let cell = XlsCell::from_record(&record, Some(&sst)).unwrap();
        assert!(matches!(cell.value(), CellValue::Error(_)));
    }

    #[test]
    fn test_from_record_label_sst_no_sst() {
        let record = CellRecord::LabelSst {
            row: 0,
            col: 0,
            xf_index: 0,
            sst_index: 0,
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::Error(_)));
    }

    #[test]
    fn test_from_record_formula_number() {
        let record = CellRecord::Formula {
            row: 0,
            col: 0,
            xf_index: 0,
            value: FormulaValue::Number(3.14),
            formula: vec![0x01, 0x02, 0x03],
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(cell.is_formula());
        if let CellValue::Float(v) = cell.value() {
            assert!((v - 3.14).abs() < 0.001);
        } else {
            panic!("Expected Float value");
        }
    }

    #[test]
    fn test_from_record_formula_string() {
        let record = CellRecord::Formula {
            row: 0,
            col: 0,
            xf_index: 0,
            value: FormulaValue::String("Result".to_string()),
            formula: vec![],
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::String(s) if s == "Result"));
    }

    #[test]
    fn test_from_record_formula_bool() {
        let record = CellRecord::Formula {
            row: 0,
            col: 0,
            xf_index: 0,
            value: FormulaValue::Bool(false),
            formula: vec![],
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::Bool(false)));
    }

    #[test]
    fn test_from_record_formula_error() {
        let record = CellRecord::Formula {
            row: 0,
            col: 0,
            xf_index: 0,
            value: FormulaValue::Error(15),
            formula: vec![],
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::Error(_)));
    }

    #[test]
    fn test_from_record_formula_empty() {
        let record = CellRecord::Formula {
            row: 0,
            col: 0,
            xf_index: 0,
            value: FormulaValue::Empty,
            formula: vec![],
        };
        let cell = XlsCell::from_record(&record, None).unwrap();
        assert!(matches!(cell.value(), CellValue::Empty));
    }

    #[test]
    fn test_cell_trait_methods_via_reference() {
        let cell = XlsCell::new(5, 10, CellValue::Float(100.0));
        let cell_ref: &XlsCell = &cell;

        // Test that &XlsCell implements Cell trait correctly
        assert_eq!(Cell::row(&cell_ref), 5);
        assert_eq!(Cell::column(&cell_ref), 10);
        assert_eq!(Cell::coordinate(&cell_ref), "K6");
        assert!(!Cell::is_formula(&cell_ref));
    }

    #[test]
    fn test_cell_clone() {
        let cell = XlsCell::with_formula(0, 0, CellValue::Float(10.0), "=SUM(A1:A10)".to_string());
        let cloned = cell.clone();
        assert_eq!(cell.row(), cloned.row());
        assert_eq!(cell.column(), cloned.column());
        assert_eq!(cell.is_formula(), cloned.is_formula());
    }

    #[test]
    fn test_cell_debug() {
        let cell = XlsCell::new(0, 0, CellValue::String("Test".to_string()));
        let debug_str = format!("{:?}", cell);
        assert!(debug_str.contains("XlsCell"));
    }

    #[test]
    fn test_cell_empty_value() {
        let cell = XlsCell::new(0, 0, CellValue::Empty);
        assert!(matches!(cell.value(), CellValue::Empty));
    }

    #[test]
    fn test_cell_int_value() {
        let cell = XlsCell::new(0, 0, CellValue::Int(42));
        assert!(matches!(cell.value(), CellValue::Int(42)));
    }
}
