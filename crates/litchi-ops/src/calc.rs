//! Formula evaluation through the `litchi-eval` engine (design §13 P0,
//! record 0266; the `eval` feature).
//!
//! [`File::evaluate`] evaluates one cell and [`File::recalculate`] evaluates
//! every formula of an XLSX, XLSB or XLS workbook. Both are reads: they never
//! change the file, and the formula caches a workbook stores are left as they
//! are. XLS and XLSB may save a blank-string placeholder when authoring a
//! formula; request recalculation to compute it instead of reading that cache.
//! The cell node view keeps reporting `cachedValue` and `evaluated`
//! (whether a cached value exists); the computed value is this module's.
//!
//! ```
//! # #[cfg(feature = "xlsx")]
//! # {
//! use litchi_core::FileFormat;
//! use litchi_ops::calc::{Options, Origin, Value};
//! use litchi_ops::props;
//!
//! let file = litchi_ops::create(FileFormat::Xlsx)?
//!     .set("/Sheet1/A1", props! { "value" => 20 })?
//!     .set("/Sheet1/A2", props! { "formula" => "SUM(A1,22)" })?;
//! let evaluated = file.evaluate("/Sheet1/A2", &Options::new())?;
//! assert_eq!(evaluated.value(), &Value::Number(42.0));
//! assert_eq!(evaluated.origin(), Origin::Computed);
//! let report = file.recalculate(&Options::new())?;
//! assert_eq!(report.uncached(), 1);
//! # }
//! # Ok::<(), litchi_ops::Error>(())
//! ```
//!
//! # Engine scope
//!
//! The engine registers about 450 Excel functions (aggregates, criteria
//! aggregates, math, trigonometry, statistics and distributions, financial,
//! date and time, text, logical, lookup, information, engineering and
//! database functions) and resolves single cells, rectangular ranges and
//! sheet-qualified references (`'My Sheet'!$A$1:$B$2`). Its expression
//! grammar has `+ - * /` and comparisons, but not `&`, `^`, `%`, unions,
//! intersections, whole-row or whole-column references, R1C1, array
//! constants, external references, structured references without a table
//! definition, or spilled results. Defined names are not resolved yet. A
//! formula outside that scope is reported as unsupported, never guessed.
//!
//! The `_xlfn.` and `_xlws.` prefixes of newer functions are removed before
//! evaluation. A formula whose result is a blank reference evaluates to 0,
//! as in Excel. Date functions use the workbook's date-system flag; evaluating
//! a 1904 workbook does not substitute 1900-based serials. Explicit structured
//! references without a registered table are unsupported, including when nested
//! inside aggregates; a missing table never supplies an empty range.
//!
//! # Volatile functions
//!
//! `NOW`, `TODAY`, `RAND`, `RANDBETWEEN` and `RANDARRAY` would consult the
//! clock or a random source, which the unified layer never does. A formula
//! that calls one keeps its cached value in every mode (reported as
//! [`Status::Volatile`]); without a cached value it cannot be evaluated.
//!
//! # Limits
//!
//! [`Options::max_cells`] bounds both the stored cells loaded from the
//! worksheets and the cell reads the engine makes (range arguments read every
//! coordinate); [`Options::max_depth`] bounds the chain of formulas one
//! evaluation follows. Exceeding either is
//! [`Code::LimitExceeded`].

use std::borrow::Cow;
use std::collections::HashMap;
use std::fmt;
use std::sync::OnceLock;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::edit::RefusalKind;
use litchi_core::sheet::{
    Cell as CoreCell, CellValue, Error as SheetError, LimitKind as SheetLimitKind,
    Result as SheetResult, WorkbookTrait, Worksheet as CoreWorksheet,
};
use litchi_eval::FormulaEvaluator;

use crate::path::IntoPath;
use crate::{Code, File, Result};

#[cfg(test)]
mod tests;

/// The functions whose result depends on the clock or a random source.
const VOLATILE: [&str; 5] = ["NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY"];

/// An evaluated or cached cell value.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Value {
    /// No value.
    Empty,
    /// A boolean.
    Bool(bool),
    /// A number (dates are serial numbers).
    Number(f64),
    /// Text.
    Text(String),
    /// A spreadsheet error value.
    Error(ErrorValue),
}

impl Value {
    /// Whether two values agree: numbers within the relative `tolerance`
    /// (`|a - b| <= tolerance * max(1, |a|, |b|)`), everything else exactly.
    /// An empty value and empty text agree.
    #[must_use]
    pub fn agrees(&self, other: &Self, tolerance: f64) -> bool {
        match (self, other) {
            (Self::Number(left), Self::Number(right)) => {
                let scale = 1f64.max(left.abs()).max(right.abs());
                (left - right).abs() <= tolerance * scale
            },
            (Self::Empty, Self::Text(text)) | (Self::Text(text), Self::Empty) => text.is_empty(),
            (left, right) => left == right,
        }
    }
}

/// A spreadsheet error value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorValue {
    /// `#NULL!`
    Null,
    /// `#DIV/0!`
    Div0,
    /// `#VALUE!`
    Value,
    /// `#REF!`
    Ref,
    /// `#NAME?`
    Name,
    /// `#NUM!`
    Num,
    /// `#N/A`
    NotAvailable,
    /// `#GETTING_DATA`
    GettingData,
    /// `#SPILL!`
    Spill,
    /// `#CALC!`
    Calc,
    /// `#CONNECT!`
    Connect,
    /// `#BLOCKED!`
    Blocked,
    /// `#FIELD!`
    Field,
    /// `#UNKNOWN!`
    Unknown,
    /// `#BUSY!`
    Busy,
}

impl ErrorValue {
    const ALL: [(Self, &'static str); 15] = [
        (Self::Null, "#NULL!"),
        (Self::Div0, "#DIV/0!"),
        (Self::Value, "#VALUE!"),
        (Self::Ref, "#REF!"),
        (Self::Name, "#NAME?"),
        (Self::Num, "#NUM!"),
        (Self::NotAvailable, "#N/A"),
        (Self::GettingData, "#GETTING_DATA"),
        (Self::Spill, "#SPILL!"),
        (Self::Calc, "#CALC!"),
        (Self::Connect, "#CONNECT!"),
        (Self::Blocked, "#BLOCKED!"),
        (Self::Field, "#FIELD!"),
        (Self::Unknown, "#UNKNOWN!"),
        (Self::Busy, "#BUSY!"),
    ];

    /// The spreadsheet spelling, for example `#DIV/0!`.
    #[must_use]
    pub fn as_str(self) -> &'static str {
        Self::ALL
            .iter()
            .find(|(value, _)| *value == self)
            .map_or("#VALUE!", |(_, text)| text)
    }

    /// Parse a spreadsheet spelling, ignoring ASCII case.
    #[must_use]
    pub fn parse(text: &str) -> Option<Self> {
        let text = text.trim();
        Self::ALL
            .iter()
            .find(|(_, spelling)| spelling.eq_ignore_ascii_case(text))
            .map(|(value, _)| *value)
    }
}

/// How [`File::evaluate`] obtained a value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Origin {
    /// The cell holds a value, not a formula.
    Literal,
    /// The formula's cached value, used because recalculation was not asked
    /// for or the formula is volatile.
    Cached,
    /// Computed by the engine.
    Computed,
}

/// Evaluation options: cached values preferred, 1,048,576 cells, a formula
/// chain of 64 and a relative tolerance of 1e-9 by default.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Options {
    recalculate: bool,
    max_cells: usize,
    max_depth: usize,
    tolerance: f64,
}

impl Default for Options {
    fn default() -> Self {
        Self::new()
    }
}

impl Options {
    /// The defaults.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            recalculate: false,
            max_cells: 1 << 20,
            max_depth: 64,
            tolerance: 1e-9,
        }
    }

    /// Ignore stored formula caches (except for volatile formulas) and
    /// compute every formula [`File::evaluate`] meets. [`File::recalculate`]
    /// always does.
    #[must_use]
    pub const fn recalculate(mut self, recalculate: bool) -> Self {
        self.recalculate = recalculate;
        self
    }

    /// Bound the stored cells loaded and the cell reads made.
    #[must_use]
    pub const fn max_cells(mut self, cells: usize) -> Self {
        self.max_cells = cells;
        self
    }

    /// Bound the chain of formula cells one evaluation follows. The engine
    /// recurses once per formula on the chain, about 2 to 2.5 KiB of stack
    /// each in an optimized build and about 10 KiB in a debug build, so raise
    /// the bound only on a thread with stack to match.
    #[must_use]
    pub const fn max_depth(mut self, depth: usize) -> Self {
        self.max_depth = depth;
        self
    }

    /// The relative tolerance [`File::recalculate`] compares numbers with.
    #[must_use]
    pub const fn tolerance(mut self, tolerance: f64) -> Self {
        self.tolerance = tolerance;
        self
    }
}

/// The result of [`File::evaluate`].
#[derive(Clone, PartialEq)]
pub struct Evaluated {
    value: Value,
    origin: Origin,
}

impl Evaluated {
    /// The value.
    #[must_use]
    pub const fn value(&self) -> &Value {
        &self.value
    }

    /// How the value was obtained.
    #[must_use]
    pub const fn origin(&self) -> Origin {
        self.origin
    }

    /// The value, consumed.
    #[must_use]
    pub fn into_value(self) -> Value {
        self.value
    }
}

// Values are authored content.
impl fmt::Debug for Evaluated {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Evaluated")
            .field("origin", &self.origin)
            .finish_non_exhaustive()
    }
}

/// How one recalculated formula compares with its cached value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Status {
    /// The computed value agrees with the cached value.
    Match,
    /// The computed value disagrees with the cached value.
    Mismatch,
    /// Computed; the workbook stores no cached value to compare with.
    Uncached,
    /// The engine cannot evaluate the formula (see the module's engine
    /// scope).
    Unsupported,
    /// The formula calls a volatile function and keeps its cached value.
    Volatile,
}

/// One formula cell of a [`Report`].
#[derive(Clone)]
pub struct Formula {
    sheet: usize,
    row: u32,
    column: u32,
    text: Box<str>,
    cached: Option<Value>,
    computed: Option<Value>,
    status: Status,
}

impl Formula {
    /// The zero-based position in the complete tab catalog, including non-worksheet tabs.
    #[must_use]
    pub const fn sheet(&self) -> usize {
        self.sheet
    }

    /// The zero-based row.
    #[must_use]
    pub const fn row(&self) -> u32 {
        self.row
    }

    /// The zero-based column.
    #[must_use]
    pub const fn column(&self) -> u32 {
        self.column
    }

    /// The formula text as stored, without a leading `=`.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The cached value the workbook stores, if any.
    #[must_use]
    pub const fn cached(&self) -> Option<&Value> {
        self.cached.as_ref()
    }

    /// The computed value, or `None` when unsupported.
    #[must_use]
    pub const fn computed(&self) -> Option<&Value> {
        self.computed.as_ref()
    }

    /// The comparison status.
    #[must_use]
    pub const fn status(&self) -> Status {
        self.status
    }
}

// Formula text and values are authored content.
impl fmt::Debug for Formula {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Formula")
            .field("sheet", &self.sheet)
            .field("row", &self.row)
            .field("column", &self.column)
            .field("status", &self.status)
            .finish_non_exhaustive()
    }
}

/// The result of [`File::recalculate`]: every formula cell in worksheet,
/// row and column order.
#[derive(Clone, Default)]
pub struct Report {
    formulas: Vec<Formula>,
}

impl Report {
    /// Every formula cell.
    #[must_use]
    pub fn formulas(&self) -> &[Formula] {
        &self.formulas
    }

    /// The number of formulas with `status`.
    #[must_use]
    pub fn count(&self, status: Status) -> usize {
        self.formulas
            .iter()
            .filter(|formula| formula.status == status)
            .count()
    }

    /// Formulas whose computed value agrees with the cache.
    #[must_use]
    pub fn matches(&self) -> usize {
        self.count(Status::Match)
    }

    /// Formulas whose computed value disagrees with the cache.
    #[must_use]
    pub fn mismatches(&self) -> usize {
        self.count(Status::Mismatch)
    }

    /// Computed formulas without a cache.
    #[must_use]
    pub fn uncached(&self) -> usize {
        self.count(Status::Uncached)
    }

    /// Formulas the engine cannot evaluate.
    #[must_use]
    pub fn unsupported(&self) -> usize {
        self.count(Status::Unsupported)
    }

    /// Volatile formulas left at their cached value.
    #[must_use]
    pub fn volatile(&self) -> usize {
        self.count(Status::Volatile)
    }
}

impl fmt::Debug for Report {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Report")
            .field("formulas", &self.formulas.len())
            .field("matches", &self.matches())
            .field("mismatches", &self.mismatches())
            .field("uncached", &self.uncached())
            .field("unsupported", &self.unsupported())
            .field("volatile", &self.volatile())
            .finish()
    }
}

/// A bounded evaluation resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Stored cells loaded and cell reads made.
    Cells,
    /// The chain of formula cells one evaluation follows.
    Depth,
}

/// A content-free evaluation failure.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The target is not one worksheet cell.
    NotACell,
    /// The engine cannot evaluate the formula.
    Unsupported,
    /// The formula is volatile and has no cached value.
    Volatile,
    /// A bound was exceeded.
    Limit {
        /// The resource.
        kind: LimitKind,
        /// The observed amount.
        observed: usize,
        /// The bound.
        maximum: usize,
    },
    /// The engine suspended, which only network functions do; they are not
    /// built.
    Suspended,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NotACell => formatter.write_str("evaluation targets one worksheet cell"),
            Self::Unsupported => formatter.write_str(
                "the formula uses syntax or a function the formula engine does not evaluate",
            ),
            Self::Volatile => formatter
                .write_str("the formula calls a clock or random function and has no cached value"),
            Self::Limit {
                kind,
                observed,
                maximum,
            } => write!(
                formatter,
                "evaluation {kind:?} limit exceeded: {observed} > {maximum}"
            ),
            Self::Suspended => formatter.write_str("the formula engine suspended"),
        }
    }
}

impl std::error::Error for Error {}

impl File {
    /// Evaluate one cell: its value, its cached formula result, or (with
    /// [`Options::recalculate`], or without a cache) a computed result.
    ///
    /// # Errors
    ///
    /// [`Code::InvalidArgument`] when the path is not a cell,
    /// [`Code::UnsupportedCapability`] for a formula the engine cannot
    /// evaluate, a volatile formula without a cache, or a non-workbook
    /// format, [`Code::LimitExceeded`] past the [`Options`] bounds, and the
    /// path and read errors of [`File::get`].
    pub fn evaluate<'p>(&self, path: impl IntoPath<'p>, options: &Options) -> Result<Evaluated> {
        let path = path.into_path()?;
        let (sheet, row, column) = cell_target(self, &path)?;
        let grid = Grid::new(self, options, options.recalculate)?;
        let stored = grid.stored(sheet, row, column)?;
        let Some(formula) = stored.formula else {
            return Ok(Evaluated {
                value: convert(&stored.value).unwrap_or(Value::Empty),
                origin: Origin::Literal,
            });
        };
        if let Some(cached) = formula
            .cached
            .filter(|_| formula.volatile || !options.recalculate)
        {
            return Ok(Evaluated {
                value: cached,
                origin: Origin::Cached,
            });
        }
        if formula.volatile {
            return Err(Error::Volatile.into());
        }
        let evaluator = grid.evaluator();
        let computed = match grid.compute(&evaluator, sheet, row, column) {
            Err(error) if is_depth(&error) => {
                // A chain that runs down the sheet resolves without deep
                // recursion when its earlier formulas are computed first.
                grid.warm(&evaluator, sheet, (row, column))?;
                grid.compute(&evaluator, sheet, row, column)?
            },
            other => other?,
        };
        match computed {
            Some(value) => Ok(Evaluated {
                value,
                origin: Origin::Computed,
            }),
            None => Err(Error::Unsupported.into()),
        }
    }

    /// Recalculate every formula of the workbook, ignoring stored caches
    /// except for volatile formulas, and compare each result with its cached
    /// value. The file is not changed.
    ///
    /// # Errors
    ///
    /// [`Code::UnsupportedCapability`] for a non-workbook format,
    /// [`Code::LimitExceeded`] past the [`Options`] bounds, and read errors.
    pub fn recalculate(&self, options: &Options) -> Result<Report> {
        let grid = Grid::new(self, options, true)?;
        let evaluator = grid.evaluator();
        let mut report = Report::default();
        for sheet in 0..grid.names.len() {
            let formulas = grid.sheet(sheet)?.formulas.clone();
            for (row, column, formula) in formulas {
                let (computed, status) = if formula.volatile {
                    (formula.cached.clone(), Status::Volatile)
                } else {
                    let computed = grid.compute(&evaluator, sheet, row, column)?;
                    let status = match (&computed, &formula.cached) {
                        (None, _) => Status::Unsupported,
                        (Some(_), None) => Status::Uncached,
                        (Some(computed), Some(cached)) => {
                            if computed.agrees(cached, options.tolerance) {
                                Status::Match
                            } else {
                                Status::Mismatch
                            }
                        },
                    };
                    (computed, status)
                };
                report
                    .formulas
                    .try_reserve(1)
                    .map_err(|source| crate::transaction::Error::Allocation { source })?;
                report.formulas.push(Formula {
                    sheet,
                    row,
                    column,
                    text: formula.text,
                    cached: formula.cached,
                    computed,
                    status,
                });
            }
        }
        Ok(report)
    }
}

fn sheet_catalog(file: &File) -> Result<Vec<(String, bool)>> {
    match file {
        #[cfg(feature = "xls")]
        File::Xls(handle) => crate::format::xls::calc::catalog(handle),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => crate::format::xlsb::calc::catalog(handle),
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => crate::format::xlsx::calc::catalog(handle),
        File::Signed(signed) => sheet_catalog(signed.file()),
        _ => Err(not_a_workbook(file)),
    }
}

/// Physical tab names for reports and native cache writes. The public family
/// worksheet view filters non-worksheets in XLS/XLSB and cannot index these
/// operations' physical tab selectors.
pub(crate) fn sheet_names(file: &File) -> Result<Vec<String>> {
    Ok(sheet_catalog(file)?
        .into_iter()
        .map(|(name, _)| name)
        .collect())
}

fn is_depth(error: &crate::Error) -> bool {
    matches!(
        error.leaf(),
        crate::error::Source::Calc(Error::Limit {
            kind: LimitKind::Depth,
            ..
        })
    )
}

/// The zero-based sheet and cell a path names.
fn cell_target(file: &File, path: &crate::path::PathArg<'_>) -> Result<(usize, u32, u32)> {
    let origin = match file {
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => crate::format::xlsx::write::import::origin(handle, path)?,
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => crate::format::xlsb::write::import::origin(handle, path)?,
        #[cfg(feature = "xls")]
        File::Xls(handle) => crate::format::xls::write::import::origin(handle, path)?,
        File::Signed(signed) => return cell_target(signed.file(), path),
        _ => {
            let _ = path;
            return Err(not_a_workbook(file));
        },
    };
    if !origin.cell {
        return Err(Error::NotACell.into());
    }
    Ok((origin.sheet, origin.row, origin.column))
}

fn not_a_workbook(file: &File) -> crate::Error {
    let error = crate::Error::new(
        Code::UnsupportedCapability,
        "formula evaluation reads worksheets of XLSX, XLSB and XLS workbooks",
    );
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}

/// Drive a future that never waits: the engine suspends only for network
/// functions, which this build does not include.
fn block_on<F: Future>(future: F) -> Result<F::Output> {
    let mut future = std::pin::pin!(future);
    let mut context = std::task::Context::from_waker(std::task::Waker::noop());
    match future.as_mut().poll(&mut context) {
        std::task::Poll::Ready(output) => Ok(output),
        std::task::Poll::Pending => Err(Error::Suspended.into()),
    }
}

/// A formula cell as loaded.
#[derive(Clone)]
struct Stored {
    text: Box<str>,
    cached: Option<Value>,
    volatile: bool,
}

/// One stored cell as [`Grid::stored`] returns it.
struct Lookup {
    value: CellValue,
    formula: Option<Stored>,
}

/// One loaded worksheet.
struct Sheet {
    name: String,
    cells: HashMap<(u32, u32), CellValue>,
    /// Formula cells in row-major order, zero-based.
    formulas: Vec<(u32, u32, Stored)>,
    dimensions: Option<(u32, u32, u32, u32)>,
}

/// The evaluation snapshot: worksheets loaded on first use, with formula
/// caches removed for recalculation. It is the engine's `WorkbookTrait`.
struct Grid<'f> {
    file: &'f File,
    names: Vec<String>,
    has_grid: Vec<bool>,
    sheets: Vec<OnceLock<Sheet>>,
    recalculate: bool,
    date1904: bool,
    max_cells: usize,
    max_depth: usize,
    cells: AtomicUsize,
    reads: AtomicUsize,
    /// The first bound exceeded, kept because the engine may fold the
    /// error into a cell value.
    tripped: OnceLock<crate::Error>,
}

impl fmt::Debug for Grid<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Grid")
            .field("sheets", &self.names.len())
            .finish_non_exhaustive()
    }
}

impl<'f> Grid<'f> {
    fn new(file: &'f File, options: &Options, recalculate: bool) -> Result<Self> {
        let inner = match file {
            File::Signed(signed) => signed.file(),
            other => other,
        };
        match inner {
            #[cfg(feature = "xlsx")]
            File::Xlsx(_) => {},
            #[cfg(feature = "xlsb")]
            File::Xlsb(_) => {},
            #[cfg(feature = "xls")]
            File::Xls(_) => {},
            _ => return Err(not_a_workbook(file)),
        }
        let view = inner.workbook()?;
        let (names, has_grid): (Vec<_>, Vec<_>) = sheet_catalog(inner)?.into_iter().unzip();
        let date1904 = view.is_1904_date_system();
        let sheets = names.iter().map(|_| OnceLock::new()).collect();
        Ok(Self {
            file: inner,
            names,
            has_grid,
            sheets,
            recalculate,
            date1904,
            max_cells: options.max_cells,
            max_depth: options.max_depth,
            cells: AtomicUsize::new(0),
            reads: AtomicUsize::new(0),
            tripped: OnceLock::new(),
        })
    }

    fn evaluator(&self) -> FormulaEvaluator<'_, Self> {
        FormulaEvaluator::new(self).with_max_depth(self.max_depth)
    }

    /// Load (once) and return a worksheet.
    fn sheet(&self, index: usize) -> Result<&Sheet> {
        let slot = self.sheets.get(index).ok_or_else(|| {
            crate::Error::new(
                Code::OutOfRange,
                "the calculation tab position is out of range",
            )
        })?;
        if let Some(sheet) = slot.get() {
            return Ok(sheet);
        }
        let sheet = self.load(index)?;
        Ok(slot.get_or_init(|| sheet))
    }

    fn load(&self, index: usize) -> Result<Sheet> {
        let mut cells = HashMap::new();
        let mut formulas = Vec::new();
        let mut dimensions: Option<(u32, u32, u32, u32)> = None;
        let mut visit = |row: u32, column: u32, value: &CellValue| -> Result<()> {
            let loaded = self.cells.fetch_add(1, Ordering::Relaxed) + 1;
            if loaded > self.max_cells {
                return Err(limit(LimitKind::Cells, loaded, self.max_cells));
            }
            let mut value = value.clone();
            if let CellValue::Formula {
                formula,
                cached_value,
                ..
            } = &mut value
            {
                let volatile = is_volatile(formula);
                let cached = cached_value.as_deref().and_then(convert);
                let text: Box<str> = formula.as_str().into();
                if let Cow::Owned(normalized) = normalize(formula) {
                    *formula = normalized;
                }
                if self.recalculate && !volatile {
                    *cached_value = None;
                }
                if volatile && cached_value.is_none() {
                    // Never let the engine call the clock or a random source.
                    value = CellValue::Error("#VOLATILE".to_owned());
                }
                formulas
                    .try_reserve(1)
                    .map_err(|source| crate::transaction::Error::Allocation { source })?;
                formulas.push((
                    row,
                    column,
                    Stored {
                        text,
                        cached,
                        volatile,
                    },
                ));
            }
            // The engine addresses cells one-based.
            let (row, column) = (row + 1, column + 1);
            dimensions = Some(match dimensions {
                None => (row, column, row, column),
                Some((top, left, bottom, right)) => (
                    top.min(row),
                    left.min(column),
                    bottom.max(row),
                    right.max(column),
                ),
            });
            cells
                .try_reserve(1)
                .map_err(|source| crate::transaction::Error::Allocation { source })?;
            cells.insert((row, column), value);
            Ok(())
        };
        match self.file {
            #[cfg(feature = "xlsx")]
            File::Xlsx(handle) => crate::format::xlsx::calc::cells(handle, index, &mut visit)?,
            #[cfg(feature = "xlsb")]
            File::Xlsb(handle) => crate::format::xlsb::calc::cells(handle, index, &mut visit)?,
            #[cfg(feature = "xls")]
            File::Xls(handle) => crate::format::xls::calc::cells(handle, index, &mut visit)?,
            _ => {},
        }
        formulas.sort_by_key(|(row, column, _)| (*row, *column));
        Ok(Sheet {
            name: self.names[index].clone(),
            cells,
            formulas,
            dimensions,
        })
    }

    /// The stored cell at zero-based coordinates, with its formula facts.
    fn stored(&self, sheet: usize, row: u32, column: u32) -> Result<Lookup> {
        let loaded = self.sheet(sheet)?;
        let value = loaded
            .cells
            .get(&(row + 1, column + 1))
            .cloned()
            .unwrap_or(CellValue::Empty);
        let formula = loaded
            .formulas
            .binary_search_by_key(&(row, column), |(row, column, _)| (*row, *column))
            .ok()
            .map(|index| loaded.formulas[index].2.clone());
        Ok(Lookup { value, formula })
    }

    /// Compute one zero-based formula cell; `None` when unsupported.
    fn compute(
        &self,
        evaluator: &FormulaEvaluator<'_, Self>,
        sheet: usize,
        row: u32,
        column: u32,
    ) -> Result<Option<Value>> {
        let name = self.names[sheet].as_str();
        let result = block_on(evaluator.evaluate_cell(name, row + 1, column + 1))?;
        if let Some(error) = self.tripped.get() {
            return Err(error.clone());
        }
        match result {
            Ok(value) => Ok(convert(&value).map(|value| match value {
                // A formula never yields a blank: a blank reference reads 0.
                Value::Empty => Value::Number(0.0),
                other => other,
            })),
            Err(SheetError::Limit {
                kind: SheetLimitKind::Depth,
                observed,
                ..
            }) => Err(limit(LimitKind::Depth, observed, self.max_depth)),
            Err(error) if error == MISSING_SHEET => Err(crate::Error::new(
                Code::NotFound,
                MISSING_SHEET
                    .refusal()
                    .map_or("", |refusal| refusal.reason()),
            )),
            // An engine refusal such as a structured reference without its
            // table: unsupported, not a failure.
            Err(_) => Ok(None),
        }
    }

    /// Compute, in row-major order, the formulas of `sheet` before `until`,
    /// filling the engine's cache. Chains that still run too deep are left
    /// for the final evaluation to report.
    fn warm(
        &self,
        evaluator: &FormulaEvaluator<'_, Self>,
        sheet: usize,
        until: (u32, u32),
    ) -> Result<()> {
        let formulas: Vec<(u32, u32)> = self
            .sheet(sheet)?
            .formulas
            .iter()
            .map(|(row, column, _)| (*row, *column))
            .take_while(|at| *at < until)
            .collect();
        for (row, column) in formulas {
            match self.compute(evaluator, sheet, row, column) {
                Err(error) if is_depth(&error) => {},
                other => {
                    other?;
                },
            }
        }
        Ok(())
    }

    /// Keep the typed error for [`Self::compute`] and hand the engine its
    /// classification.
    fn trip(&self, error: crate::Error) -> SheetError {
        let _ = self.tripped.set(error.clone());
        SheetError::from(error)
    }
}

/// The grid's refusal of a worksheet name it does not have, which
/// [`Grid::compute`] reports as [`Code::NotFound`].
const MISSING_SHEET: SheetError = SheetError::refused(
    RefusalKind::NotFound,
    "worksheet_by_name",
    "the formula names a worksheet the workbook does not have",
);

/// The grid's refusal of the reads the engine never makes.
const BY_ROW_AND_COLUMN: SheetError = SheetError::refused(
    RefusalKind::UnsupportedOperation,
    "read",
    "the evaluation grid reads cells by row and column",
);

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> crate::Error {
    Error::Limit {
        kind,
        observed,
        maximum,
    }
    .into()
}

/// Convert an engine or cached value; `None` for an error the engine raised
/// that is not a spreadsheet error value (an unsupported construct).
fn convert(value: &CellValue) -> Option<Value> {
    Some(match value {
        CellValue::Empty => Value::Empty,
        CellValue::Bool(value) => Value::Bool(*value),
        #[allow(
            clippy::cast_precision_loss,
            reason = "spreadsheet numbers are doubles"
        )]
        CellValue::Int(value) => Value::Number(*value as f64),
        CellValue::Float(value) | CellValue::DateTime(value) => Value::Number(*value),
        CellValue::String(text) => Value::Text(text.clone()),
        CellValue::Error(text) => Value::Error(ErrorValue::parse(text)?),
        CellValue::Formula { cached_value, .. } => {
            return cached_value.as_deref().and_then(convert);
        },
    })
}

/// Whether a formula calls a volatile function (outside string literals).
fn is_volatile(formula: &str) -> bool {
    let mut quoted = false;
    let bytes = formula.as_bytes();
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        if byte == b'"' {
            quoted = !quoted;
            index += 1;
            continue;
        }
        if quoted || !(byte.is_ascii_alphabetic() || byte == b'_') {
            index += 1;
            continue;
        }
        let start = index;
        while index < bytes.len()
            && (bytes[index].is_ascii_alphanumeric() || matches!(bytes[index], b'_' | b'.'))
        {
            index += 1;
        }
        if bytes.get(index) == Some(&b'(') {
            let name = &formula[start..index];
            let name = strip_prefix(name);
            if VOLATILE
                .iter()
                .any(|volatile| volatile.eq_ignore_ascii_case(name))
            {
                return true;
            }
        }
    }
    false
}

fn strip_prefix(name: &str) -> &str {
    for prefix in ["_xlfn.", "_xlws."] {
        if name.len() > prefix.len() && name[..prefix.len()].eq_ignore_ascii_case(prefix) {
            return &name[prefix.len()..];
        }
    }
    name
}

/// Remove the `_xlfn.` and `_xlws.` prefixes outside string literals.
fn normalize(formula: &str) -> Cow<'_, str> {
    if !formula.contains("_xl") && !formula.contains("_XL") {
        return Cow::Borrowed(formula);
    }
    let mut output = String::with_capacity(formula.len());
    let mut quoted = false;
    let mut rest = formula;
    while let Some(character) = rest.chars().next() {
        if character == '"' {
            quoted = !quoted;
        } else if !quoted {
            let stripped = strip_prefix(rest);
            if stripped.len() != rest.len() {
                rest = stripped;
                continue;
            }
        }
        output.push(character);
        rest = &rest[character.len_utf8()..];
    }
    Cow::Owned(output)
}

impl<'f> WorkbookTrait for Grid<'f> {
    type Error = SheetError;
    type Worksheet<'g>
        = SheetRef<'g, 'f>
    where
        Self: 'g;
    type Worksheets<'g>
        = Sheets<'g, 'f>
    where
        Self: 'g;

    fn active_sheet_index(&self) -> usize {
        0
    }

    fn active_worksheet(&self) -> SheetResult<SheetRef<'_, 'f>> {
        self.worksheet_by_index(0)
    }

    fn is_1904_date_system(&self) -> bool {
        self.date1904
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<SheetRef<'_, 'f>> {
        if index >= self.sheets.len() {
            return Err(SheetError::refused(
                RefusalKind::InvalidValue,
                "worksheet_by_index",
                "the worksheet position is out of range",
            ));
        }
        if !self.has_grid[index] {
            return Err(MISSING_SHEET);
        }
        let sheet = self.sheet(index).map_err(|error| self.trip(error))?;
        Ok(SheetRef { grid: self, sheet })
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<SheetRef<'_, 'f>> {
        let index = self
            .names
            .iter()
            .position(|candidate| candidate == name)
            .or_else(|| {
                self.names
                    .iter()
                    .position(|candidate| candidate.eq_ignore_ascii_case(name))
            })
            .ok_or(MISSING_SHEET)?;
        self.worksheet_by_index(index)
    }

    fn worksheet_count(&self) -> usize {
        self.names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        &self.names
    }

    fn worksheets(&self) -> Sheets<'_, 'f> {
        Sheets {
            grid: self,
            index: 0,
        }
    }
}

struct Sheets<'g, 'f> {
    grid: &'g Grid<'f>,
    index: usize,
}

impl<'g, 'f> Iterator for Sheets<'g, 'f> {
    type Item = SheetResult<SheetRef<'g, 'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.index >= self.grid.names.len() {
            return None;
        }
        self.index += 1;
        Some(self.grid.worksheet_by_index(self.index - 1))
    }
}

struct SheetRef<'g, 'f> {
    grid: &'g Grid<'f>,
    sheet: &'g Sheet,
}

struct GridCell<'g> {
    row: u32,
    column: u32,
    value: Cow<'g, CellValue>,
}

impl CoreCell for GridCell<'_> {
    fn column(&self) -> u32 {
        self.column
    }

    fn coordinate(&self) -> String {
        let mut label = String::new();
        let mut column = self.column;
        while column > 0 {
            let digit = u8::try_from((column - 1) % 26).unwrap_or(0);
            label.insert(0, char::from(b'A' + digit));
            column = (column - 1) / 26;
        }
        format!("{label}{}", self.row)
    }

    fn row(&self) -> u32 {
        self.row
    }

    fn value(&self) -> &CellValue {
        &self.value
    }

    fn is_formula(&self) -> bool {
        matches!(self.value.as_ref(), CellValue::Formula { .. })
    }
}

struct GridCells<'g> {
    inner: std::collections::hash_map::Iter<'g, (u32, u32), CellValue>,
}

impl<'g> Iterator for GridCells<'g> {
    type Item = SheetResult<GridCell<'g>>;

    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next().map(|((row, column), value)| {
            Ok(GridCell {
                row: *row,
                column: *column,
                value: Cow::Borrowed(value),
            })
        })
    }
}

impl SheetRef<'_, '_> {
    fn read(&self, row: u32, column: u32) -> SheetResult<Cow<'_, CellValue>> {
        let reads = self.grid.reads.fetch_add(1, Ordering::Relaxed) + 1;
        if reads > self.grid.max_cells {
            return Err(self
                .grid
                .trip(limit(LimitKind::Cells, reads, self.grid.max_cells)));
        }
        Ok(self
            .sheet
            .cells
            .get(&(row, column))
            .map_or(Cow::Borrowed(CellValue::EMPTY), Cow::Borrowed))
    }
}

impl CoreWorksheet for SheetRef<'_, '_> {
    type Error = SheetError;
    type Cell<'a>
        = GridCell<'a>
    where
        Self: 'a;
    type Cells<'a>
        = GridCells<'a>
    where
        Self: 'a;
    type Rows<'a>
        = std::iter::Empty<SheetResult<Cow<'a, [CellValue]>>>
    where
        Self: 'a;

    fn cell(&self, row: u32, column: u32) -> SheetResult<GridCell<'_>> {
        Ok(GridCell {
            row,
            column,
            value: self.read(row, column)?,
        })
    }

    fn cell_by_coordinate(&self, _coordinate: &str) -> SheetResult<GridCell<'_>> {
        Err(BY_ROW_AND_COLUMN)
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<Cow<'_, CellValue>> {
        self.read(row, column)
    }

    fn cells(&self) -> GridCells<'_> {
        GridCells {
            inner: self.sheet.cells.iter(),
        }
    }

    fn column_count(&self) -> usize {
        self.sheet
            .dimensions
            .map_or(0, |(_, left, _, right)| (right - left + 1) as usize)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.sheet.dimensions
    }

    fn name(&self) -> &str {
        &self.sheet.name
    }

    fn row(&self, _row_idx: usize) -> SheetResult<Cow<'_, [CellValue]>> {
        Err(BY_ROW_AND_COLUMN)
    }

    fn row_count(&self) -> usize {
        self.sheet
            .dimensions
            .map_or(0, |(top, _, bottom, _)| (bottom - top + 1) as usize)
    }

    fn rows(&self) -> Self::Rows<'_> {
        std::iter::empty()
    }
}
