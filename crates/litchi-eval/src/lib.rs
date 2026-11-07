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
use litchi_core::edit::RefusalKind;
use litchi_core::sheet::{
    CellValue, Error as SheetError, LimitKind, Result, WorkbookTrait, Worksheet as _,
};
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
    /// URLs this pass needs from the caller's transport, in request order.
    #[cfg(feature = "web_functions")]
    requests: Mutex<Vec<String>>,
}

/// A transport response as the synchronous engine sees it.
#[cfg(feature = "web_functions")]
pub(crate) enum Fetched {
    /// No transport is configured.
    NoTransport,
    /// The response body, or `None` when the request failed.
    Body(Option<Vec<u8>>),
    /// Not fetched yet: recorded for the next pass.
    Pending,
}

/// A `[@…]` structured reference outside a table row.
const ROW_CONTEXT: SheetError = SheetError::refused(
    RefusalKind::UnsupportedOperation,
    "evaluate",
    "a this-row structured reference requires row context",
);

/// A structured reference naming a column the table does not have.
const MISSING_COLUMN: SheetError = SheetError::refused(
    RefusalKind::NotFound,
    "evaluate",
    "the table has no column of that name",
);

/// Explicit table syntax requires a table definition from the caller.
const MISSING_TABLE: SheetError = SheetError::refused(
    RefusalKind::UnsupportedContent,
    "evaluate",
    "a structured reference requires a registered table",
);

struct Visit<'a> {
    session: &'a Mutex<HashSet<CellRef>>,
    key: CellRef,
}

impl EvalSession {
    /// Record a URL the pass needs.
    #[cfg(feature = "web_functions")]
    fn request(&self, url: &str) {
        let mut requests = self.requests.lock();
        if !requests.iter().any(|known| known == url) {
            requests.push(url.to_owned());
        }
    }

    /// Whether the pass waits on a transport response, so values it computes
    /// may rest on a placeholder.
    fn is_suspended(&self) -> bool {
        #[cfg(feature = "web_functions")]
        {
            !self.requests.lock().is_empty()
        }
        #[cfg(not(feature = "web_functions"))]
        {
            false
        }
    }

    /// The number of formula cells on the current evaluation chain.
    fn depth(&self) -> usize {
        self.visiting.lock().len()
    }

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

#[cfg(feature = "web_functions")]
use std::future::Future;
#[cfg(feature = "web_functions")]
use std::pin::Pin;

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
    fn get_cell_value(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue>;

    /// Returns the current evaluation position (sheet, row, col) if a formula is being
    /// evaluated. This is primarily used by functions such as ROW() or COLUMN() that need
    /// to know the location of the formula cell when no explicit reference is supplied.
    fn current_position(&self) -> Option<(String, u32, u32)>;

    /// Returns the raw value stored in the workbook without triggering evaluation.
    ///
    /// This is useful for functions like ISFORMULA that need to inspect the cell's
    /// original content rather than the evaluated result.
    fn raw_cell_value(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue>;

    /// Returns true if the workbook backing this context uses the 1904 date system.
    fn is_1904_date_system(&self) -> bool;

    /// The response the caller's transport gave for `url`.
    ///
    /// The engine is synchronous: a URL not fetched yet is recorded and
    /// reported as [`Fetched::Pending`], and the async wrapper fetches it
    /// through the transport and evaluates again.
    #[cfg(feature = "web_functions")]
    fn fetched(&self, url: &str) -> Fetched;

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
    max_depth: usize,
    #[cfg(feature = "web_functions")]
    fetch: Option<&'a dyn Fetch>,
    /// Transport responses by URL (`None`: the request failed or exceeded
    /// the response bound).
    #[cfg(feature = "web_functions")]
    responses: Mutex<HashMap<String, Option<Vec<u8>>>>,
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

impl<'e, 'w, W: WorkbookTrait + Sync + Send + ?Sized> EngineCtx for At<'e, 'w, W>
where
    SheetError: From<W::Error>,
{
    fn get_cell_value(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        self.evaluator
            .get_cell_value_in(self.session, sheet_name, row, col)
    }

    fn current_position(&self) -> Option<(String, u32, u32)> {
        Some(self.position.clone())
    }

    fn raw_cell_value(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        self.evaluator.stored(sheet_name, row, col)
    }

    fn is_1904_date_system(&self) -> bool {
        self.evaluator.workbook.is_1904_date_system()
    }

    #[cfg(feature = "web_functions")]
    fn fetched(&self, url: &str) -> Fetched {
        if self.evaluator.fetch.is_none() {
            return Fetched::NoTransport;
        }
        if let Some(response) = self.evaluator.responses.lock().get(url) {
            return Fetched::Body(response.clone());
        }
        self.session.request(url);
        Fetched::Pending
    }

    fn get_sheet_index(&self, name: &str) -> Option<usize> {
        self.evaluator.sheet_index.get(name).copied()
    }

    fn get_sheet_count(&self) -> usize {
        self.evaluator.workbook.worksheet_names().len()
    }
}

impl<W: WorkbookTrait + Sync + Send + ?Sized> ReferenceResolver for At<'_, '_, W>
where
    SheetError: From<W::Error>,
{
    fn resolve_name(&self, current_sheet: &str, name: &str) -> Result<Option<ResolvedName>> {
        self.evaluator.resolve_name(current_sheet, name)
    }
}

impl<'a, W: WorkbookTrait + Sync + Send + ?Sized> ReferenceResolver for FormulaEvaluator<'a, W>
where
    SheetError: From<W::Error>,
{
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

impl<'a, W: WorkbookTrait + Sync + Send + ?Sized> FormulaEvaluator<'a, W>
where
    SheetError: From<W::Error>,
{
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
            max_depth: usize::MAX,
            #[cfg(feature = "web_functions")]
            fetch: None,
            #[cfg(feature = "web_functions")]
            responses: Mutex::new(HashMap::new()),
        }
    }

    /// Bound the chain of formula cells one evaluation follows; a longer
    /// chain fails with [`litchi_core::sheet::Error::Limit`] of kind
    /// [`LimitKind::Depth`]. Unbounded by default.
    #[must_use]
    pub fn with_max_depth(mut self, depth: usize) -> Self {
        self.max_depth = depth;
        self
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

    /// The stored value of a cell, without evaluation.
    fn stored(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        let sheet = self.workbook.worksheet_by_name(sheet_name)?;
        let value: Cow<'_, CellValue> = sheet.cell_value(row, col)?;
        Ok(value.into_owned())
    }

    fn get_cell_value_in(
        &self,
        session: &EvalSession,
        sheet_name: &str,
        row: u32,
        col: u32,
    ) -> Result<CellValue> {
        let Some(sheet_idx) = self.sheet_index.get(sheet_name).copied() else {
            return Err(SheetError::refused(
                RefusalKind::NotFound,
                "evaluate",
                "the formula names a worksheet the workbook does not list",
            ));
        };
        let key = CellRef {
            sheet_idx,
            row,
            col,
        };

        if let Some(value) = self.eval_state.read().cache.get(&key) {
            return Ok(value.clone());
        }

        let depth = session.depth();
        if depth >= self.max_depth {
            return Err(SheetError::Limit {
                kind: LimitKind::Depth,
                observed: depth + 1,
                maximum: self.max_depth,
            });
        }
        let Some(_visit) = session.enter(key) else {
            return Ok(CellValue::Error("Circular reference detected".to_string()));
        };

        let result = match self.stored(sheet_name, row, col)? {
            // Prefer the cached result embedded in the file.
            CellValue::Formula {
                cached_value: Some(cached),
                ..
            } => Ok(*cached),
            // No cached value: evaluate the formula text. Any parsing or
            // semantic issue is reported as `CellValue::Error` rather than a
            // hard failure.
            CellValue::Formula { formula, .. } => {
                self.evaluate_formula(session, sheet_name, row, col, &formula)
            },
            other => Ok(other),
        };

        // A value computed while an external fetch is outstanding may rest
        // on a placeholder; the next pass computes it again.
        if let Ok(value) = &result
            && !session.is_suspended()
        {
            self.eval_state.write().cache.insert(key, value.clone());
        }
        result
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

    pub fn define_table(&mut self, config: TableConfig<'_>) {
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
            // Returning None here turns the structured reference into an
            // unresolved scalar name. Aggregates can then discard that error
            // as nonnumeric input and manufacture a result such as SUM=0.
            // Bare identifiers must still fall through to defined-name lookup.
            None if name.contains('[') => {
                return Err(MISSING_TABLE);
            },
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
                return Err(ROW_CONTEXT);
            },
            StructuredReference::Column { column_name, .. } => {
                let col = table
                    .headers
                    .get(&column_name.to_uppercase())
                    .copied()
                    .ok_or(MISSING_COLUMN)?;
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
                    .ok_or(MISSING_COLUMN)?;
                return Err(ROW_CONTEXT);
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
                    .ok_or(MISSING_COLUMN)?;
                let end_col = table
                    .headers
                    .get(&end_column.to_uppercase())
                    .copied()
                    .ok_or(MISSING_COLUMN)?;
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
                    .ok_or(MISSING_COLUMN)?;
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
                    .ok_or(MISSING_COLUMN)?;
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
    ///
    /// Evaluation is synchronous. The function is `async` only for
    /// `WEBSERVICE`: a pass that needs a response the caller's [`Fetch`]
    /// transport has not given yet records the URL, the transport is awaited
    /// for it, and the cell is evaluated again (values already computed stay
    /// cached). Without the `web_functions` feature, or without a transport,
    /// the future completes on its first poll.
    pub async fn evaluate_cell(&self, sheet_name: &str, row: u32, col: u32) -> Result<CellValue> {
        loop {
            let session = EvalSession::default();
            let result = self.get_cell_value_in(&session, sheet_name, row, col);
            if !self.resume(session).await {
                return result;
            }
        }
    }

    /// Evaluate all cells in a worksheet and return a dense 2D grid
    /// covering the sheet's declared dimensions.
    pub async fn evaluate_sheet(&self, sheet_name: &str) -> Result<Vec<Vec<CellValue>>> {
        loop {
            let session = EvalSession::default();
            let result = self.evaluate_sheet_in(&session, sheet_name);
            if !self.resume(session).await {
                return result;
            }
        }
    }

    fn evaluate_sheet_in(
        &self,
        session: &EvalSession,
        sheet_name: &str,
    ) -> Result<Vec<Vec<CellValue>>> {
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
                out_row.push(self.get_cell_value_in(session, sheet_name, row, col)?);
            }
            rows.push(out_row);
        }

        Ok(rows)
    }

    /// Fetch the responses the finished pass of `session` asked for. Returns
    /// whether it fetched any, so the pass must run again.
    #[cfg(feature = "web_functions")]
    async fn resume(&self, session: EvalSession) -> bool {
        let requests = session.requests.into_inner();
        let Some(fetch) = self.fetch.filter(|_| !requests.is_empty()) else {
            return false;
        };
        for url in requests {
            let response = fetch
                .get(&url, engine::MAX_RESPONSE_BYTES)
                .await
                .ok()
                .filter(|bytes| bytes.len() <= engine::MAX_RESPONSE_BYTES);
            self.responses.lock().insert(url, response);
        }
        true
    }

    #[cfg(not(feature = "web_functions"))]
    #[allow(
        clippy::unused_async,
        reason = "the evaluation entry points stay async for the transport"
    )]
    async fn resume(&self, _session: EvalSession) -> bool {
        false
    }

    fn evaluate_formula(
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
            return engine::evaluate_expression(&at, sheet_name, &expr);
        }

        // Unsupported or unrecognized formula in this MVP implementation.
        Ok(CellValue::Error(format!(
            "Unsupported formula for MVP evaluator: {}",
            s
        )))
    }
}

#[cfg(all(test, feature = "web_functions"))]
mod transport_tests {
    //! `WEBSERVICE` through the evaluator: the synchronous core asks for a
    //! URL, the async entry point awaits the caller's transport, and the
    //! cell is evaluated again.

    use super::*;
    use litchi_core::sheet::{Cell, Worksheet};
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[derive(Debug)]
    struct Book {
        names: Vec<String>,
        cells: HashMap<(u32, u32), CellValue>,
    }

    struct Fixed<'a>(u32, u32, &'a CellValue);

    impl Cell for Fixed<'_> {
        fn column(&self) -> u32 {
            self.1
        }
        fn coordinate(&self) -> String {
            String::new()
        }
        fn row(&self) -> u32 {
            self.0
        }
        fn value(&self) -> &CellValue {
            self.2
        }
    }

    impl Worksheet for Book {
        type Error = SheetError;
        type Cell<'a> = Fixed<'a>;
        type Cells<'a> = std::iter::Empty<Result<Fixed<'a>>>;
        type Rows<'a> = std::iter::Empty<Result<Cow<'a, [CellValue]>>>;

        fn cell(&self, row: u32, column: u32) -> Result<Fixed<'_>> {
            Ok(Fixed(
                row,
                column,
                self.cells.get(&(row, column)).unwrap_or(CellValue::EMPTY),
            ))
        }
        fn cell_by_coordinate(&self, _: &str) -> Result<Fixed<'_>> {
            unreachable!()
        }
        fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
            Ok(Cow::Borrowed(
                self.cells.get(&(row, column)).unwrap_or(CellValue::EMPTY),
            ))
        }
        fn cells(&self) -> Self::Cells<'_> {
            std::iter::empty()
        }
        fn column_count(&self) -> usize {
            3
        }
        fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
            Some((1, 1, 1, 3))
        }
        fn name(&self) -> &str {
            &self.names[0]
        }
        fn row(&self, _: usize) -> Result<Cow<'_, [CellValue]>> {
            unreachable!()
        }
        fn row_count(&self) -> usize {
            1
        }
        fn rows(&self) -> Self::Rows<'_> {
            std::iter::empty()
        }
    }

    impl WorkbookTrait for Book {
        type Error = SheetError;
        type Worksheet<'a> = &'a Book;
        type Worksheets<'a> = std::iter::Once<Result<&'a Book>>;

        fn active_sheet_index(&self) -> usize {
            0
        }
        fn active_worksheet(&self) -> Result<&Book> {
            Ok(self)
        }
        fn worksheet_by_index(&self, _: usize) -> Result<&Book> {
            Ok(self)
        }
        fn worksheet_by_name(&self, _: &str) -> Result<&Book> {
            Ok(self)
        }
        fn worksheet_count(&self) -> usize {
            1
        }
        fn worksheet_names(&self) -> &[String] {
            &self.names
        }
        fn worksheets(&self) -> Self::Worksheets<'_> {
            std::iter::once(Ok(self))
        }
    }

    fn formula(text: &str) -> CellValue {
        CellValue::Formula {
            formula: text.to_owned(),
            cached_value: None,
            is_array: false,
            array_range: None,
        }
    }

    #[tokio::test]
    async fn missing_table_references_refuse_without_breaking_registered_tables_or_names() {
        for reference in ["Inventory[Quantity]", "Inventory[#All]", "Inventory[]"] {
            let book = Book {
                names: vec!["Sheet1".into()],
                cells: HashMap::from([((1, 3), formula(&format!("SUM({reference})")))]),
            };
            assert_eq!(
                FormulaEvaluator::new(&book)
                    .evaluate_cell("Sheet1", 1, 3)
                    .await,
                Err(MISSING_TABLE)
            );
        }
        let book = Book {
            names: vec!["Sheet1".into()],
            cells: HashMap::from([
                ((1, 3), formula("SUM(Inventory[Quantity])")),
                ((2, 3), formula("SUM(Total)")),
                ((2, 1), CellValue::String("Quantity".into())),
                ((3, 1), CellValue::Int(5)),
                ((4, 1), CellValue::Int(7)),
            ]),
        };
        let mut evaluator = FormulaEvaluator::new(&book);
        assert_eq!(
            evaluator.evaluate_cell("Sheet1", 1, 3).await,
            Err(MISSING_TABLE)
        );
        evaluator.define_table(TableConfig {
            name: "Inventory",
            sheet_name: "Sheet1",
            start_row: 2,
            start_col: 1,
            end_row: 4,
            end_col: 1,
            headers: &["Quantity".into()],
        });
        evaluator.define_name("Total", "Sheet1!A3:A4");
        for row in [1, 2] {
            assert_eq!(
                evaluator.evaluate_cell("Sheet1", row, 3).await.unwrap(),
                CellValue::Float(12.0)
            );
        }
    }

    /// Answers `https://a/` with `https://b/`, `https://b/` with `done`, and
    /// suspends once before each answer, as a real transport would.
    #[derive(Default)]
    struct Chained {
        calls: AtomicUsize,
    }

    impl Fetch for Chained {
        fn get<'a>(&'a self, url: &'a str, _max: usize) -> FetchFuture<'a> {
            self.calls.fetch_add(1, Ordering::Relaxed);
            Box::pin(async move {
                tokio::task::yield_now().await;
                match url {
                    "https://a/" => Ok(b"https://b/".to_vec()),
                    "https://b/" => Ok(b"done".to_vec()),
                    _ => Err("unknown".into()),
                }
            })
        }
    }

    #[tokio::test]
    async fn webservice_fetches_through_the_transport_and_reevaluates() {
        let book = Book {
            names: vec!["Sheet1".to_owned()],
            cells: HashMap::from([
                ((1, 1), formula("WEBSERVICE(\"https://a/\")")),
                ((1, 2), formula("WEBSERVICE(A1)")),
                ((1, 3), formula("CONCAT(B1,\"!\",A1)")),
            ]),
        };
        let transport = Chained::default();
        let evaluator = FormulaEvaluator::new(&book).with_fetch(&transport);
        assert_eq!(
            evaluator.evaluate_cell("Sheet1", 1, 3).await.unwrap(),
            CellValue::String("done!https://b/".to_owned())
        );
        // One request per distinct URL; the chained one in a second pass.
        assert_eq!(transport.calls.load(Ordering::Relaxed), 2);
        assert_eq!(
            evaluator.evaluate_sheet("Sheet1").await.unwrap(),
            vec![vec![
                CellValue::String("https://b/".to_owned()),
                CellValue::String("done".to_owned()),
                CellValue::String("done!https://b/".to_owned()),
            ]]
        );
        assert_eq!(transport.calls.load(Ordering::Relaxed), 2);
    }

    #[tokio::test]
    async fn webservice_failures_and_missing_transport_keep_their_values() {
        let book = Book {
            names: vec!["Sheet1".to_owned()],
            cells: HashMap::from([((1, 1), formula("WEBSERVICE(\"https://x/\")"))]),
        };
        let transport = Chained::default();
        let evaluator = FormulaEvaluator::new(&book).with_fetch(&transport);
        assert_eq!(
            evaluator.evaluate_cell("Sheet1", 1, 1).await.unwrap(),
            CellValue::Error("#VALUE!".to_owned())
        );
        let offline = FormulaEvaluator::new(&book);
        assert_eq!(
            offline.evaluate_cell("Sheet1", 1, 1).await.unwrap(),
            CellValue::Error("#CONNECT!".to_owned())
        );
    }
}
