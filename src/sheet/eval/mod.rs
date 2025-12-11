//! Formula evaluation engine shared across spreadsheet formats.
//!
//! This module provides a small, format-agnostic evaluation layer that works
//! on top of the unified `sheet` traits. It is intentionally conservative:
//! it prefers using cached values embedded in files and can be extended
//! over time to support more Excel semantics.

pub mod engine;
pub mod parser;

use crate::sheet::{CellValue, Result, WorkbookTrait};
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::{HashMap, HashSet};

#[derive(Clone, Copy, Hash, PartialEq, Eq)]
struct CellRef {
    sheet_idx: usize,
    row: u32,
    col: u32,
}

struct EvalState {
    cache: HashMap<CellRef, CellValue>,
    visiting: HashSet<CellRef>,
}

/// Evaluation context used by the engine runtime.
pub(crate) trait EngineCtx {
    fn get_cell_value(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue>;
}

/// Simple formula evaluator operating on a `WorkbookTrait`.
///
/// The initial implementation is intentionally basic:
/// - For non-formula cells, it returns the stored value.
/// - For formula cells, it returns the cached result if present.
/// - If no cached result is available, it returns an Error cell.
pub struct FormulaEvaluator<'a, W: WorkbookTrait + ?Sized> {
    workbook: &'a W,
    sheet_index: HashMap<String, usize>,
    eval_state: RefCell<EvalState>,
}

impl<'a, W: WorkbookTrait + ?Sized> EngineCtx for FormulaEvaluator<'a, W> {
    fn get_cell_value(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        let sheet_idx = *self
            .sheet_index
            .get(sheet_name)
            .expect("Sheet name not found in index");
        let key = CellRef {
            sheet_idx,
            row,
            col,
        };

        // Fast path: cached value
        {
            let state = self.eval_state.borrow();
            if let Some(v) = state.cache.get(&key) {
                return Ok(v.clone());
            }
            if state.visiting.contains(&key) {
                // Circular reference detected.
                return Ok(CellValue::Error("Circular reference detected".to_string()));
            }
        }

        // Mark as visiting
        {
            let mut state = self.eval_state.borrow_mut();
            state.visiting.insert(key);
        }

        // Load raw value from workbook
        let sheet = self.workbook.worksheet_by_name(sheet_name)?;
        let value: Cow<'_, CellValue> = sheet.cell_value(row, col)?;
        let raw = value.into_owned();

        // Evaluate value (handles formulas and cached results)
        let result = self.evaluate_value(sheet_name, row, col, raw)?;

        // Store in cache and clear visiting
        {
            let mut state = self.eval_state.borrow_mut();
            state.visiting.remove(&key);
            state.cache.insert(key, result.clone());
        }

        Ok(result)
    }
}

impl<'a, W: WorkbookTrait + ?Sized> FormulaEvaluator<'a, W> {
    /// Create a new evaluator for the given workbook.
    pub fn new(workbook: &'a W) -> Self {
        let mut sheet_index = HashMap::new();
        for (idx, name) in workbook.worksheet_names().iter().enumerate() {
            sheet_index.insert(name.clone(), idx);
        }
        Self {
            workbook,
            sheet_index,
            eval_state: RefCell::new(EvalState {
                cache: HashMap::new(),
                visiting: HashSet::new(),
            }),
        }
    }

    /// Evaluate a single cell in the given worksheet.
    ///
    /// Row and column are 1-based, consistent with the `Worksheet` trait.
    pub fn evaluate_cell(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        self.get_cell_value(sheet_name, row, col)
    }

    /// Evaluate all cells in a worksheet and return a dense 2D grid
    /// covering the sheet's declared dimensions.
    pub fn evaluate_sheet(&self, sheet_name: &str) -> Result<Vec<Vec<CellValue>>> {
        let sheet = self.workbook.worksheet_by_name(sheet_name)?;
        let dims = match sheet.dimensions() {
            Some(d) => d,
            None => return Ok(Vec::new()),
        };

        let (min_row, min_col, max_row, max_col) = dims;
        let mut rows = Vec::new();

        for row in min_row..=max_row {
            let mut out_row = Vec::new();
            for col in min_col..=max_col {
                out_row.push(self.get_cell_value(sheet_name, row, col)?);
            }
            rows.push(out_row);
        }

        Ok(rows)
    }

    /// Core evaluation routine for a single cell value.
    ///
    /// This remains conservative and still prefers cached results when
    /// available. When no cached result is present, it performs a minimal
    /// evaluation of the formula text, currently limited to:
    ///
    /// - Literal constants (numbers, strings, booleans)
    /// - Single-cell references (same-sheet or qualified with a sheet name)
    fn evaluate_value(
        &self,
        sheet_name: &str,
        row: u32,
        col: u32,
        value: CellValue,
    ) -> Result<CellValue> {
        let result = match value {
            CellValue::Formula {
                formula,
                cached_value,
                ..
            } => {
                if let Some(cached) = cached_value {
                    // Prefer the cached result embedded in the file.
                    (*cached).clone()
                } else {
                    // No cached value – perform a minimal evaluation of the
                    // formula text. Any parsing/semantic issues are reported as
                    // CellValue::Error rather than hard failures.
                    self.evaluate_formula(sheet_name, row, col, &formula)?
                }
            },
            other => other,
        };

        Ok(result)
    }

    fn evaluate_formula(
        &self,
        sheet_name: &str,
        _row: u32,
        _col: u32,
        expr: &str,
    ) -> Result<CellValue> {
        let s = expr.trim();
        if s.is_empty() {
            return Ok(CellValue::Error("Empty formula".to_string()));
        }

        // General expression (e.g., A1+2, 1+2*3, CONCAT("a","b"),
        // TEXTJOIN("-",TRUE,A1:A3)). This uses the small expression parser
        // and runtime engine. If parsing fails, fall back to returning an
        // Error cell rather than panicking.
        if let Some(expr) = parser::parse_expression(sheet_name, s) {
            return engine::evaluate_expression(self, sheet_name, &expr);
        }

        // Unsupported or unrecognized formula in this MVP implementation.
        Ok(CellValue::Error(format!(
            "Unsupported formula for MVP evaluator: {}",
            s
        )))
    }
}
