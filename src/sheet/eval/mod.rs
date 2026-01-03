//! Formula evaluation engine shared across spreadsheet formats.
//!
//! This module provides a small, format-agnostic evaluation layer that works
//! on top of the unified `sheet` traits. It is intentionally conservative:
//! it prefers using cached values embedded in files and can be extended
//! over time to support more Excel semantics.

pub mod engine;
pub mod parser;

use self::engine::{ReferenceResolver, ResolvedName};
use self::parser::{RangeRef, parse_range_reference, parse_single_cell_reference};
use crate::sheet::{CellValue, Result, WorkbookTrait};
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
use tokio::sync::RwLock;

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

use std::future::Future;
use std::pin::Pin;

pub(crate) type BoxFuture<'a, T> = Pin<Box<dyn Future<Output = T> + Send + 'a>>;

/// Evaluation context used by the engine runtime.
pub(crate) trait EngineCtx: Send + Sync {
    fn get_cell_value<'a>(
        &'a self,
        sheet_name: &'a str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'a, Result<CellValue>>;

    /// Returns the current evaluation position (sheet, row, col) if a formula is being
    /// evaluated. This is primarily used by functions such as ROW() or COLUMN() that need
    /// to know the location of the formula cell when no explicit reference is supplied.
    fn current_position(&self) -> Option<(String, u32, u32)>;

    /// Returns the raw value stored in the workbook without triggering evaluation.
    ///
    /// This is useful for functions like ISFORMULA that need to inspect the cell's
    /// original content rather than the evaluated result.
    fn raw_cell_value<'a>(
        &'a self,
        sheet_name: &'a str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'a, Result<CellValue>>;

    /// Returns true if the workbook backing this context uses the 1904 date system.
    fn is_1904_date_system(&self) -> bool;

    /// Returns a shared HTTP client for web functions.
    #[cfg(feature = "eval_engine_web_functions")]
    fn http_client(&self) -> &reqwest::Client;

    /// Returns the index of the given sheet (0-based).
    fn get_sheet_index(&self, name: &str) -> Option<usize>;

    /// Returns the total number of sheets in the workbook.
    fn get_sheet_count(&self) -> usize;
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
    eval_state: RwLock<EvalState>,
    names: HashMap<String, String>,
    local_names: HashMap<(String, String), String>,
    tables: HashMap<String, NamedTable>,
    position_stack: RwLock<Vec<(String, u32, u32)>>,
    #[cfg(feature = "eval_engine_web_functions")]
    http_client: reqwest::Client,
}

#[derive(Clone)]
struct NamedTable {
    sheet: String,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
    headers: HashMap<String, u32>,
}

impl<'a, W: WorkbookTrait + Sync + Send + ?Sized> EngineCtx for FormulaEvaluator<'a, W> {
    fn get_cell_value<'b>(
        &'b self,
        sheet_name: &'b str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'b, Result<CellValue>> {
        Box::pin(async move {
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
                let state = self.eval_state.read().await;
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
                let mut state = self.eval_state.write().await;
                state.visiting.insert(key);
            }

            // Load raw value from workbook
            let sheet = self.workbook.worksheet_by_name(sheet_name)?;
            let value: Cow<'_, CellValue> = sheet.cell_value(row, col)?;
            let raw = value.into_owned();

            // Evaluate value (handles formulas and cached results)
            let result = self.evaluate_value(sheet_name, row, col, raw).await?;

            // Store in cache and clear visiting
            {
                let mut state = self.eval_state.write().await;
                state.visiting.remove(&key);
                state.cache.insert(key, result.clone());
            }

            Ok(result)
        })
    }

    fn current_position(&self) -> Option<(String, u32, u32)> {
        // NOTE: This now returns a default or requires a sync way to access.
        // For simplicity in this migration, we might need a sync guard or just accept that
        // functions needing position might need to be async or we use a different mechanism.
        // For now, use blocking read as it's just a stack of strings/u32.
        self.position_stack.blocking_read().last().cloned()
    }

    fn raw_cell_value<'b>(
        &'b self,
        sheet_name: &'b str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'b, Result<CellValue>> {
        Box::pin(async move {
            let sheet = self.workbook.worksheet_by_name(sheet_name)?;
            let value: Cow<'_, CellValue> = sheet.cell_value(row, col)?;
            Ok(value.into_owned())
        })
    }

    fn is_1904_date_system(&self) -> bool {
        self.workbook.is_1904_date_system()
    }

    #[cfg(feature = "eval_engine_web_functions")]
    fn http_client(&self) -> &reqwest::Client {
        &self.http_client
    }

    fn get_sheet_index(&self, name: &str) -> Option<usize> {
        self.sheet_index.get(name).copied()
    }

    fn get_sheet_count(&self) -> usize {
        self.workbook.worksheet_names().len()
    }
}

impl<'a, W: WorkbookTrait + ?Sized> ReferenceResolver for FormulaEvaluator<'a, W> {
    fn resolve_name(&self, current_sheet: &str, name: &str) -> Result<Option<ResolvedName>> {
        let trimmed = name.trim();
        if trimmed.is_empty() {
            return Ok(None);
        }

        if let Some(resolved) = self.resolve_table_reference(current_sheet, trimmed)? {
            return Ok(Some(resolved));
        }

        let norm = trimmed.to_uppercase();
        if let Some(reference) = self
            .local_names
            .get(&(current_sheet.to_string(), norm.clone()))
        {
            return Ok(self.resolve_reference_string(current_sheet, reference));
        }

        if let Some(reference) = self.names.get(&norm) {
            return Ok(self.resolve_reference_string(current_sheet, reference));
        }

        Ok(None)
    }
}

pub struct TableConfig<'a> {
    pub name: &'a str,
    pub sheet_name: &'a str,
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    pub headers: &'a [String],
}

impl<'a, W: WorkbookTrait + Sync + Send + ?Sized> FormulaEvaluator<'a, W> {
    /// Create a new evaluator for the given workbook.
    pub fn new(workbook: &'a W) -> Self {
        let mut sheet_index = HashMap::new();
        for (idx, name) in workbook.worksheet_names().iter().enumerate() {
            sheet_index.insert(name.clone(), idx);
        }
        Self {
            workbook,
            sheet_index,
            eval_state: RwLock::new(EvalState {
                cache: HashMap::new(),
                visiting: HashSet::new(),
            }),
            names: HashMap::new(),
            local_names: HashMap::new(),
            tables: HashMap::new(),
            position_stack: RwLock::new(Vec::new()),
            #[cfg(feature = "eval_engine_web_functions")]
            http_client: reqwest::Client::new(),
        }
    }

    pub fn define_name(&mut self, name: &str, reference: &str) {
        self.names
            .insert(name.trim().to_uppercase(), reference.trim().to_string());
    }

    pub fn define_name_local(&mut self, sheet_name: &str, name: &str, reference: &str) {
        self.local_names.insert(
            (sheet_name.to_string(), name.trim().to_uppercase()),
            reference.trim().to_string(),
        );
    }

    pub fn define_table(&mut self, config: TableConfig) {
        let mut header_map = HashMap::new();
        for (i, h) in config.headers.iter().enumerate() {
            let col = config.start_col + i as u32;
            if col > config.end_col {
                break;
            }
            let key = h.trim().to_uppercase();
            if !key.is_empty() {
                header_map.insert(key, col);
            }
        }
        self.tables.insert(
            config.name.trim().to_uppercase(),
            NamedTable {
                sheet: config.sheet_name.to_string(),
                start_row: config.start_row,
                start_col: config.start_col,
                end_row: config.end_row,
                end_col: config.end_col,
                headers: header_map,
            },
        );
    }

    fn resolve_reference_string(
        &self,
        current_sheet: &str,
        reference: &str,
    ) -> Option<ResolvedName> {
        if let Some(range) = parse_range_reference(current_sheet, reference) {
            return Some(ResolvedName::Range(range));
        }
        if let Some((sheet, row, col)) = parse_single_cell_reference(current_sheet, reference) {
            return Some(ResolvedName::Cell { sheet, row, col });
        }
        None
    }

    fn resolve_table_reference(
        &self,
        _current_sheet: &str,
        name: &str,
    ) -> Result<Option<ResolvedName>> {
        let (table_name, rest) = match name.split_once('[') {
            Some(v) => v,
            None => {
                let norm = name.to_uppercase();
                return Ok(self.tables.get(&norm).map(|t| {
                    ResolvedName::Range(RangeRef {
                        sheet: t.sheet.clone(),
                        start_row: t.start_row,
                        start_col: t.start_col,
                        end_row: t.end_row,
                        end_col: t.end_col,
                    })
                }));
            },
        };

        let table_norm = table_name.trim().to_uppercase();
        let table = match self.tables.get(&table_norm) {
            Some(t) => t,
            None => return Ok(None),
        };

        let spec = rest.trim_end_matches(']').trim();
        let spec = spec.trim_matches(|c| c == '[' || c == ']');
        let last = spec.split(',').next_back().unwrap_or("").trim();
        let last = last.trim_matches(|c| c == '[' || c == ']');
        let last_norm = last.to_uppercase();

        let mut out = RangeRef {
            sheet: table.sheet.clone(),
            start_row: table.start_row,
            start_col: table.start_col,
            end_row: table.end_row,
            end_col: table.end_col,
        };

        match last_norm.as_str() {
            "#ALL" => {},
            "#DATA" => {
                if out.start_row < out.end_row {
                    out.start_row += 1;
                }
            },
            "#HEADERS" => {
                out.end_row = out.start_row;
            },
            _ => {
                if let Some(col) = table.headers.get(&last_norm).copied() {
                    out.start_col = col;
                    out.end_col = col;
                    if out.start_row < out.end_row {
                        out.start_row += 1;
                    }
                } else {
                    return Ok(None);
                }
            },
        }

        Ok(Some(ResolvedName::Range(out)))
    }

    /// Evaluate a single cell in the given worksheet.
    ///
    /// Row and column are 1-based, consistent with the `Worksheet` trait.
    pub async fn evaluate_cell(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        self.get_cell_value(sheet_name, row, col).await
    }

    /// Evaluate all cells in a worksheet and return a dense 2D grid
    /// covering the sheet's declared dimensions.
    pub async fn evaluate_sheet(&self, sheet_name: &str) -> Result<Vec<Vec<CellValue>>> {
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
                out_row.push(self.get_cell_value(sheet_name, row, col).await?);
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
    async fn evaluate_value(
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
                    self.evaluate_formula(sheet_name, row, col, &formula)
                        .await?
                }
            },
            other => other,
        };

        Ok(result)
    }

    async fn evaluate_formula(
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

        let body = s.strip_prefix('=').unwrap_or(s);
        if body.is_empty() {
            return Ok(CellValue::Error("Empty formula".to_string()));
        }

        struct PositionGuard<'a> {
            stack: &'a RwLock<Vec<(String, u32, u32)>>,
        }

        impl<'a> PositionGuard<'a> {
            async fn new(
                stack: &'a RwLock<Vec<(String, u32, u32)>>,
                sheet: &str,
                row: u32,
                col: u32,
            ) -> Self {
                stack.write().await.push((sheet.to_string(), row, col));
                PositionGuard { stack }
            }
        }

        impl<'a> Drop for PositionGuard<'a> {
            fn drop(&mut self) {
                if let Ok(mut guard) = self.stack.try_write() {
                    guard.pop();
                }
            }
        }

        let _position_guard =
            PositionGuard::new(&self.position_stack, sheet_name, _row, _col).await;

        // General expression (e.g., A1+2, 1+2*3, CONCAT("a","b"),
        // TEXTJOIN("-",TRUE,A1:A3)). This uses the small expression parser
        // and runtime engine. If parsing fails, fall back to returning an
        // Error cell rather than panicking.
        if let Some(expr) = parser::parse_expression(sheet_name, body) {
            return engine::evaluate_expression(self, sheet_name, &expr).await;
        }

        // Unsupported or unrecognized formula in this MVP implementation.
        Ok(CellValue::Error(format!(
            "Unsupported formula for MVP evaluator: {}",
            s
        )))
    }
}
