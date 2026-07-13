//! Cell representation for XLSB files

use crate::xlsb::formula::{
    CellParsedFormula, FormulaConverter, FormulaGroup, FormulaGroupKind, FormulaParser,
};
use crate::xlsb::records::CellRecord;
use litchi_core::sheet::{Cell, CellValue};
use std::sync::Arc;

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
    /// Original binary formula, retained even if a token is not understood.
    formula: Option<CellParsedFormula>,
    /// Shared array/shared-formula definition.
    formula_group: Option<Arc<FormulaGroup>>,
    /// Formula calculation flags from `GrbitFmla`.
    formula_flags: u16,
}

impl XlsbCell {
    /// Create a new XLSB cell
    pub fn new(row: u32, col: u32, value: CellValue) -> Self {
        XlsbCell {
            row,
            col,
            value,
            is_formula: false,
            formula: None,
            formula_group: None,
            formula_flags: 0,
        }
    }

    /// Create a new XLSB cell from a formula
    pub fn new_formula(row: u32, col: u32, value: CellValue) -> Self {
        XlsbCell {
            row,
            col,
            value,
            is_formula: true,
            formula: None,
            formula_group: None,
            formula_flags: 0,
        }
    }

    /// Create a formula cell while preserving the exact XLSB token and
    /// ancillary streams.
    pub(crate) fn new_formula_binary(
        row: u32,
        col: u32,
        cached_value: CellValue,
        formula: CellParsedFormula,
        formula_flags: u16,
    ) -> Self {
        let value = FormulaParser::new(&formula.rgce)
            .parse()
            .and_then(|tokens| FormulaConverter::try_tokens_to_string(&tokens))
            .map(|formula_text| CellValue::Formula {
                formula: formula_text,
                cached_value: Some(Box::new(cached_value.clone())),
                is_array: false,
                array_range: None,
            })
            .unwrap_or(cached_value);

        Self {
            value,
            row,
            col,
            is_formula: true,
            formula: Some(formula),
            formula_group: None,
            formula_flags,
        }
    }

    pub(crate) fn new_grouped_formula(
        row: u32,
        col: u32,
        cached_value: CellValue,
        placeholder: CellParsedFormula,
        formula_flags: u16,
        group: Arc<FormulaGroup>,
    ) -> Self {
        let tokens = match group.kind {
            FormulaGroupKind::Array => FormulaParser::new(&group.formula.rgce).parse(),
            FormulaGroupKind::Shared => {
                FormulaParser::with_base_cell(&group.formula.rgce, row, col).parse()
            },
        };
        let is_array = group.kind == FormulaGroupKind::Array;
        let value = tokens
            .and_then(|tokens| FormulaConverter::try_tokens_to_string(&tokens))
            .map(|formula_text| CellValue::Formula {
                formula: formula_text,
                cached_value: Some(Box::new(cached_value.clone())),
                is_array,
                array_range: is_array.then(|| group.range.to_a1()),
            })
            .unwrap_or(cached_value);
        Self {
            value,
            row,
            col,
            is_formula: true,
            formula: Some(placeholder),
            formula_group: Some(group),
            formula_flags,
        }
    }

    /// Raw XLSB formula RPN token stream (`rgce`).
    pub fn formula_bytes(&self) -> Option<&[u8]> {
        self.formula.as_ref().map(|formula| formula.rgce.as_slice())
    }

    /// Raw XLSB ancillary formula stream (`rgcb`).
    pub fn formula_extra_bytes(&self) -> Option<&[u8]> {
        self.formula.as_ref().map(|formula| formula.rgcb.as_slice())
    }

    /// Actual formula definition for an array/shared formula cell.
    pub fn formula_definition_bytes(&self) -> Option<&[u8]> {
        self.formula_group
            .as_ref()
            .map(|group| group.formula.rgce.as_slice())
    }

    /// Ancillary token data for an array/shared formula definition.
    pub fn formula_definition_extra_bytes(&self) -> Option<&[u8]> {
        self.formula_group
            .as_ref()
            .map(|group| group.formula.rgcb.as_slice())
    }

    /// Whether the grouped definition requests unconditional recalculation.
    pub fn formula_group_always_calculates(&self) -> Option<bool> {
        self.formula_group
            .as_ref()
            .map(|group| group.always_calculate)
    }

    /// Whether this cell belongs to an XLSB shared-formula group.
    pub fn is_shared_formula(&self) -> bool {
        self.formula_group
            .as_ref()
            .is_some_and(|group| group.kind == FormulaGroupKind::Shared)
    }

    /// Array/shared formula range in A1 notation.
    pub fn formula_range(&self) -> Option<String> {
        self.formula_group.as_ref().map(|group| group.range.to_a1())
    }

    /// Formula recalculation flags from `GrbitFmla`.
    pub fn formula_flags(&self) -> Option<u16> {
        self.is_formula.then_some(self.formula_flags)
    }

    /// Create cell from XLSB record
    pub fn from_record(record: &CellRecord, shared_strings: Option<&Vec<String>>) -> Option<Self> {
        let (value, is_formula) = match &record.value {
            crate::xlsb::records::CellValue::Blank => (CellValue::Empty, false),
            crate::xlsb::records::CellValue::Bool(b) => (CellValue::Bool(*b), false),
            crate::xlsb::records::CellValue::Error(e) => {
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
            crate::xlsb::records::CellValue::Real(f) => (CellValue::Float(*f), false),
            crate::xlsb::records::CellValue::String(s) => (CellValue::String(s.clone()), false),
            crate::xlsb::records::CellValue::Isst(idx) => {
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
            crate::xlsb::records::CellValue::Formula { value, formula: _ } => {
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
            formula: None,
            formula_group: None,
            formula_flags: 0,
        })
    }

    /// Extract value from formula cached value
    fn extract_formula_value(
        formula_value: &crate::xlsb::records::CellValue,
        shared_strings: Option<&Vec<String>>,
    ) -> CellValue {
        match formula_value {
            crate::xlsb::records::CellValue::Blank => CellValue::Empty,
            crate::xlsb::records::CellValue::Bool(b) => CellValue::Bool(*b),
            crate::xlsb::records::CellValue::Error(e) => {
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
            crate::xlsb::records::CellValue::Real(f) => CellValue::Float(*f),
            crate::xlsb::records::CellValue::String(s) => CellValue::String(s.clone()),
            crate::xlsb::records::CellValue::Isst(idx) => {
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
            crate::xlsb::records::CellValue::Formula { value, formula: _ } => {
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
        crate::xlsb::utils::cell_reference(self.row, self.col)
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        self.is_formula
    }
}
