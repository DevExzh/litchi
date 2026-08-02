#![allow(missing_docs)]
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
use litchi_core::sheet::{CellValue, Result, WorkbookTrait};
use parking_lot::{Mutex, RwLock};
use std::borrow::Cow;
use std::collections::{HashMap, HashSet};
#[cfg(feature = "web_functions")]
use std::error::Error;

#[derive(Clone, Copy, Hash, PartialEq, Eq)]
struct CellRef {
    sheet_idx: usize,
    row: u32,
    col: u32,
}

struct EvalState {
    cache: HashMap<CellRef, CellValue>,
}

#[derive(Default)]
struct EvalSession {
    visiting: Mutex<HashSet<CellRef>>,
}

struct Visit<'a> {
    session: &'a Mutex<HashSet<CellRef>>,
    key: CellRef,
}

impl EvalSession {
    fn enter(&self, key: CellRef) -> Option<Visit<'_>> {
        let inserted = {
            let mut visiting = self.visiting.lock();
            visiting.insert(key)
        };

        inserted.then(|| Visit {
            session: &self.visiting,
            key,
        })
    }
}

impl Drop for Visit<'_> {
    fn drop(&mut self) {
        self.session.lock().remove(&self.key);
    }
}

#[cfg(test)]
mod session_tests {
    use super::*;

    #[test]
    fn cycle_tracking_is_local_and_clears_on_drop() {
        let key = CellRef {
            sheet_idx: 0,
            row: 1,
            col: 1,
        };
        let first = EvalSession::default();
        let second = EvalSession::default();

        let visit = first.enter(key).unwrap();
        assert!(first.enter(key).is_none());
        let independent = second.enter(key).unwrap();

        drop(visit);
        assert!(first.enter(key).is_some());
        drop(independent);
    }
}

use std::future::Future;
use std::pin::Pin;

pub(crate) type BoxFuture<'a, T> = Pin<Box<dyn Future<Output = T> + Send + 'a>>;

/// Error returned by a caller-provided [`Fetch`] implementation.
#[cfg(feature = "web_functions")]
pub type FetchError = Box<dyn Error + Send + Sync + 'static>;

/// Result returned by a caller-provided [`Fetch`] implementation.
#[cfg(feature = "web_functions")]
pub type FetchResult = std::result::Result<Vec<u8>, FetchError>;

/// Runtime-neutral future returned by [`Fetch::get`].
#[cfg(feature = "web_functions")]
pub type FetchFuture<'a> = Pin<Box<dyn Future<Output = FetchResult> + Send + 'a>>;

/// Caller-owned transport for formula functions that retrieve external data.
///
/// Litchi does not select an async runtime, perform implicit network I/O, or
/// define network policy. Implementations may use any runtime and should stop
/// reading once `max` bytes have been received. The evaluator independently
/// checks that limit before accepting the response.
#[cfg(feature = "web_functions")]
pub trait Fetch: Send + Sync {
    /// Retrieve `url`, returning at most `max` response bytes.
    fn get<'a>(&'a self, url: &'a str, max: usize) -> FetchFuture<'a>;
}

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

    /// Returns the caller-provided external-data transport, when configured.
    #[cfg(feature = "web_functions")]
    fn fetch(&self) -> Option<&dyn Fetch>;

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
    #[cfg(feature = "web_functions")]
    fetch: Option<&'a dyn Fetch>,
}

/// Per-evaluation position, borrowed instead of stored in shared mutable state.
struct At<'e, 'w, W: WorkbookTrait + ?Sized> {
    evaluator: &'e FormulaEvaluator<'w, W>,
    session: &'e EvalSession,
    position: (String, u32, u32),
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

impl<'e, 'w, W: WorkbookTrait + Sync + Send + ?Sized> EngineCtx for At<'e, 'w, W> {
    fn get_cell_value<'a>(
        &'a self,
        sheet_name: &'a str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'a, Result<CellValue>> {
        self.evaluator
            .get_cell_value_in(self.session, sheet_name, row, col)
    }

    fn current_position(&self) -> Option<(String, u32, u32)> {
        Some(self.position.clone())
    }

    fn raw_cell_value<'a>(
        &'a self,
        sheet_name: &'a str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'a, Result<CellValue>> {
        Box::pin(async move {
            let sheet = self.evaluator.workbook.worksheet_by_name(sheet_name)?;
            let value: Cow<'_, CellValue> = sheet.cell_value(row, col)?;
            Ok(value.into_owned())
        })
    }

    fn is_1904_date_system(&self) -> bool {
        self.evaluator.workbook.is_1904_date_system()
    }

    #[cfg(feature = "web_functions")]
    fn fetch(&self) -> Option<&dyn Fetch> {
        self.evaluator.fetch
    }

    fn get_sheet_index(&self, name: &str) -> Option<usize> {
        self.evaluator.sheet_index.get(name).copied()
    }

    fn get_sheet_count(&self) -> usize {
        self.evaluator.workbook.worksheet_names().len()
    }
}

impl<W: WorkbookTrait + ?Sized> ReferenceResolver for At<'_, '_, W> {
    fn resolve_name(&self, current_sheet: &str, name: &str) -> Result<Option<ResolvedName>> {
        self.evaluator.resolve_name(current_sheet, name)
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
            }),
            names: HashMap::new(),
            local_names: HashMap::new(),
            tables: HashMap::new(),
            #[cfg(feature = "web_functions")]
            fetch: None,
        }
    }

    /// Attach the transport used by functions such as `WEBSERVICE`.
    ///
    /// Evaluation is network-inert until a transport is explicitly supplied.
    #[cfg(feature = "web_functions")]
    #[must_use]
    pub fn with_fetch(mut self, fetch: &'a dyn Fetch) -> Self {
        self.fetch = Some(fetch);
        self
    }

    fn get_cell_value_in<'b>(
        &'b self,
        session: &'b EvalSession,
        sheet_name: &'b str,
        row: u32,
        col: u32,
    ) -> BoxFuture<'b, Result<CellValue>> {
        Box::pin(async move {
            let sheet_idx = self.sheet_index.get(sheet_name).copied().ok_or_else(|| {
                std::io::Error::new(
                    std::io::ErrorKind::NotFound,
                    format!("worksheet '{sheet_name}' is not indexed"),
                )
            })?;
            let key = CellRef {
                sheet_idx,
                row,
                col,
            };

            if let Some(value) = self.eval_state.read().cache.get(&key) {
                return Ok(value.clone());
            }

            let Some(_visit) = session.enter(key) else {
                return Ok(CellValue::Error("Circular reference detected".to_string()));
            };

            let result = async {
                let sheet = self.workbook.worksheet_by_name(sheet_name)?;
                let value: Cow<'_, CellValue> = sheet.cell_value(row, col)?;
                self.evaluate_value(session, sheet_name, row, col, value.into_owned())
                    .await
            }
            .await;

            if let Ok(value) = &result {
                self.eval_state.write().cache.insert(key, value.clone());
            }
            result
        })
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
        use self::parser::StructuredReference;

        let structured_ref = match parser::parse_structured_reference(name) {
            Some(r) => r,
            None => return Ok(None),
        };

        let table_name = match &structured_ref {
            StructuredReference::WholeTable { table_name }
            | StructuredReference::DataOnly { table_name }
            | StructuredReference::Headers { table_name }
            | StructuredReference::Totals { table_name }
            | StructuredReference::All { table_name }
            | StructuredReference::ThisRow { table_name }
            | StructuredReference::Column { table_name, .. }
            | StructuredReference::ColumnThisRow { table_name, .. }
            | StructuredReference::ColumnRange { table_name, .. }
            | StructuredReference::HeaderColumn { table_name, .. }
            | StructuredReference::TotalsColumn { table_name, .. } => table_name,
        };

        let table = match self.tables.get(&table_name.to_uppercase()) {
            Some(t) => t,
            None => return Ok(None),
        };

        let range = match structured_ref {
            StructuredReference::WholeTable { .. } | StructuredReference::All { .. } => RangeRef {
                sheet: table.sheet.clone(),
                start_row: table.start_row,
                start_col: table.start_col,
                end_row: table.end_row,
                end_col: table.end_col,
            },
            StructuredReference::DataOnly { .. } => {
                let mut range = RangeRef {
                    sheet: table.sheet.clone(),
                    start_row: table.start_row,
                    start_col: table.start_col,
                    end_row: table.end_row,
                    end_col: table.end_col,
                };
                if range.start_row < range.end_row {
                    range.start_row += 1;
                }
                range
            },
            StructuredReference::Headers { .. } => RangeRef {
                sheet: table.sheet.clone(),
                start_row: table.start_row,
                start_col: table.start_col,
                end_row: table.start_row,
                end_col: table.end_col,
            },
            StructuredReference::Totals { .. } => RangeRef {
                sheet: table.sheet.clone(),
                start_row: table.end_row,
                start_col: table.start_col,
                end_row: table.end_row,
                end_col: table.end_col,
            },
            StructuredReference::ThisRow { .. } => {
                return Err("[@] this row references require row context".into());
            },
            StructuredReference::Column { column_name, .. } => {
                let col = table
                    .headers
                    .get(&column_name.to_uppercase())
                    .copied()
                    .ok_or_else(|| format!("Column '{}' not found in table", column_name))?;
                let mut range = RangeRef {
                    sheet: table.sheet.clone(),
                    start_row: table.start_row,
                    start_col: col,
                    end_row: table.end_row,
                    end_col: col,
                };
                if range.start_row < range.end_row {
                    range.start_row += 1;
                }
                range
            },
            StructuredReference::ColumnThisRow { column_name, .. } => {
                let _col = table
                    .headers
                    .get(&column_name.to_uppercase())
                    .copied()
                    .ok_or_else(|| format!("Column '{}' not found in table", column_name))?;
                return Err(
                    format!("[@{}] this row references require row context", column_name).into(),
                );
            },
            StructuredReference::ColumnRange {
                start_column,
                end_column,
                ..
            } => {
                let start_col = table
                    .headers
                    .get(&start_column.to_uppercase())
                    .copied()
                    .ok_or_else(|| format!("Column '{}' not found in table", start_column))?;
                let end_col = table
                    .headers
                    .get(&end_column.to_uppercase())
                    .copied()
                    .ok_or_else(|| format!("Column '{}' not found in table", end_column))?;
                let mut range = RangeRef {
                    sheet: table.sheet.clone(),
                    start_row: table.start_row,
                    start_col,
                    end_row: table.end_row,
                    end_col,
                };
                if range.start_row < range.end_row {
                    range.start_row += 1;
                }
                range
            },
            StructuredReference::HeaderColumn { column_name, .. } => {
                let col = table
                    .headers
                    .get(&column_name.to_uppercase())
                    .copied()
                    .ok_or_else(|| format!("Column '{}' not found in table", column_name))?;
                RangeRef {
                    sheet: table.sheet.clone(),
                    start_row: table.start_row,
                    start_col: col,
                    end_row: table.start_row,
                    end_col: col,
                }
            },
            StructuredReference::TotalsColumn { column_name, .. } => {
                let col = table
                    .headers
                    .get(&column_name.to_uppercase())
                    .copied()
                    .ok_or_else(|| format!("Column '{}' not found in table", column_name))?;
                RangeRef {
                    sheet: table.sheet.clone(),
                    start_row: table.end_row,
                    start_col: col,
                    end_row: table.end_row,
                    end_col: col,
                }
            },
        };

        Ok(Some(ResolvedName::Range(range)))
    }

    /// Evaluate a single cell in the given worksheet.
    ///
    /// Row and column are 1-based, consistent with the `Worksheet` trait.
    pub async fn evaluate_cell(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        let session = EvalSession::default();
        self.get_cell_value_in(&session, sheet_name, row, col).await
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
        let session = EvalSession::default();

        for row in min_row..=max_row {
            let mut out_row = Vec::new();
            for col in min_col..=max_col {
                out_row.push(
                    self.get_cell_value_in(&session, sheet_name, row, col)
                        .await?,
                );
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
        session: &EvalSession,
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
                    self.evaluate_formula(session, sheet_name, row, col, &formula)
                        .await?
                }
            },
            other => other,
        };

        Ok(result)
    }

    async fn evaluate_formula(
        &self,
        session: &EvalSession,
        sheet_name: &str,
        row: u32,
        col: u32,
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

        // General expression (e.g., A1+2, 1+2*3, CONCAT("a","b"),
        // TEXTJOIN("-",TRUE,A1:A3)). This uses the small expression parser
        // and runtime engine. If parsing fails, fall back to returning an
        // Error cell rather than panicking.
        if let Some(expr) = parser::parse_expression(sheet_name, body) {
            let at = At {
                evaluator: self,
                session,
                position: (sheet_name.to_owned(), row, col),
            };
            return engine::evaluate_expression(&at, sheet_name, &expr).await;
        }

        // Unsupported or unrecognized formula in this MVP implementation.
        Ok(CellValue::Error(format!(
            "Unsupported formula for MVP evaluator: {}",
            s
        )))
    }
}
