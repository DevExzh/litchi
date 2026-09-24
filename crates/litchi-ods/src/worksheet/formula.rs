//! Borrowed worksheet values for local formula evaluation.
//!
//! Construction indexes physical row and cell runs once; it does not expand
//! repeated cells or clone worksheet text. The upfront work is proportional
//! to the stored runs, so reuse one resolver for multiple evaluations. Cell
//! reads use logical coordinates and never refresh formula caches. The index
//! retains its metadata reservation in the constructor's budget. Each lookup
//! uses the execution context supplied by the evaluation caller, so a reused
//! index does not redirect later work or cancellation to its preparation policy.
//!
//! Reuse the resolver while its borrowed sheets stay immutable:
//!
//! ```
//! use litchi_core::ExecutionContext;
//! use litchi_ods::{Cell, CellValue, Row, Sheet};
//! use litchi_ods::codec::formula::{
//!     expression::Expression,
//!     evaluation::value::{self, Context, Limits, Mode, Position, SheetExtent, Value},
//! };
//! use litchi_ods::worksheet::formula::Resolver;
//!
//! # fn example(execution: &ExecutionContext) -> Result<(), Box<dyn std::error::Error>> {
//! let mut row = Row::new();
//! row.push_cell(Cell::new(CellValue::Number(41.0), "41"))?;
//! let mut sheet = Sheet::new("Data")?;
//! sheet.push_row(row)?;
//! let sheets = [sheet];
//! let resolver = Resolver::new(&sheets, SheetExtent::new(100, 26), execution)?;
//! let expression = Expression::parse("=[.A1]+1")?;
//! let context = Context::new(execution, Position::new("Data", 0, 0)).with_mode(Mode::Scalar);
//! let result = value::evaluate(&expression, &resolver, &context, &Limits::default())?;
//! assert!(matches!(result.value(), Value::Number(42.0)));
//! # Ok(())
//! # }
//! ```

mod index;

use crate::codec::formula::evaluation::{
    EvaluationResult, ScalarError,
    value::{CellRead, Resolver as ValueResolver, SheetExtent},
};
use litchi_core::{ExecutionContext, ExecutionError, ResourceLimit};
use std::{collections::TryReserveError, fmt};

use super::{CellValue, Sheet, Snapshot};

/// Failure to prepare a borrowed worksheet resolver.
#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    /// The caller supplied a zero row or column count.
    InvalidExtent,
    /// More than one sheet has the same exact name.
    DuplicateSheetName,
    /// Physical run lengths overflow the logical coordinate domain.
    CoordinateOverflow,
    /// Caller execution policy refused preparation, including cancellation.
    Execution(ExecutionError),
    /// Index storage exceeded the caller's finite resource budget.
    ResourceLimit(ResourceLimit),
    /// Fallible allocation of index metadata failed.
    Allocation {
        /// Index storage that could not be allocated.
        resource: &'static str,
        /// Original allocator failure.
        source: TryReserveError,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidExtent => formatter.write_str("formula grid dimensions must be nonzero"),
            Self::DuplicateSheetName => {
                formatter.write_str("formula resolver sheet names are ambiguous")
            },
            Self::CoordinateOverflow => {
                formatter.write_str("formula worksheet run coordinates overflow")
            },
            Self::Execution(error) => error.fmt(formatter),
            Self::ResourceLimit(error) => error.fmt(formatter),
            Self::Allocation { resource, source } => {
                write!(formatter, "allocation failed for {resource}: {source}")
            },
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Execution(error) => Some(error),
            Self::ResourceLimit(error) => Some(error),
            Self::Allocation { source, .. } => Some(source),
            Self::InvalidExtent | Self::DuplicateSheetName | Self::CoordinateOverflow => None,
        }
    }
}

impl From<ExecutionError> for Error {
    fn from(error: ExecutionError) -> Self {
        match error {
            ExecutionError::ResourceLimit(limit) => Self::ResourceLimit(limit),
            other => Self::Execution(other),
        }
    }
}

impl From<ResourceLimit> for Error {
    fn from(error: ResourceLimit) -> Self {
        Self::ResourceLimit(error)
    }
}

/// An immutable borrowed worksheet index for local formula reads.
///
/// The caller supplies finite grid dimensions explicitly. Stored row/cell
/// extents are not treated as the spreadsheet's normative grid bounds.
/// Evaluated values may borrow worksheet text and cannot outlive the resolver:
///
/// ```compile_fail,E0515
/// use litchi_core::ExecutionContext;
/// use litchi_ods::{Sheet, worksheet::formula::Resolver};
/// use litchi_ods::codec::formula::{
///     expression::Expression,
///     evaluation::value::{evaluate, Context, Evaluated, Limits, Position, SheetExtent},
/// };
///
/// fn escape<'a>(expression: &'a Expression, execution: &ExecutionContext) -> Evaluated<'a> {
///     let sheets: [Sheet; 0] = [];
///     let resolver = Resolver::new(&sheets, SheetExtent::new(100, 26), execution).unwrap();
///     let context = Context::new(execution, Position::new("Data", 0, 0));
///     evaluate(expression, &resolver, &context, &Limits::default()).unwrap()
/// }
/// ```
pub struct Resolver<'a> {
    index: index::Index<'a>,
    extent: SheetExtent,
}

impl fmt::Debug for Resolver<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Resolver")
            .field("extent", &self.extent)
            .finish()
    }
}

impl<'a> Resolver<'a> {
    /// Build a budgeted index over immutable borrowed sheets.
    ///
    /// This accepts in-memory sheet graphs and checks the run coordinates and
    /// names needed for lookup. It does not validate worksheet serialization
    /// rules, styles, or authoring limits. Cells outside `extent` are inaccessible
    /// through this resolver even when present in the source graph.
    /// The index reservation remains charged to `execution` until this resolver
    /// is dropped. Later lookups use their supplied execution context. Share a
    /// budget and cancellation token across preparation and evaluation when both
    /// operations must follow one aggregate policy.
    ///
    /// # Errors
    /// Returns an error for empty grid dimensions, ambiguous sheet names,
    /// overflowing logical run coordinates, cancellation or resource limits.
    pub fn new(
        sheets: &'a [Sheet],
        extent: SheetExtent,
        execution: &ExecutionContext,
    ) -> Result<Self, Error> {
        execution.check()?;
        if extent.rows() == 0 || extent.columns() == 0 {
            return Err(Error::InvalidExtent);
        }
        Ok(Self {
            index: index::Index::new(sheets, execution)?,
            extent,
        })
    }

    /// Borrow an existing worksheet snapshot without reparsing its package.
    ///
    /// # Errors
    /// Returns the same index, dimension and execution errors as [`Self::new`].
    pub fn from_snapshot(
        snapshot: &'a Snapshot,
        extent: SheetExtent,
        execution: &ExecutionContext,
    ) -> Result<Self, Error> {
        Self::new(snapshot.sheets(), extent, execution)
    }

    /// Caller-selected finite grid dimensions for each sheet.
    #[must_use]
    pub const fn extent(&self) -> SheetExtent {
        self.extent
    }

    /// Retained index metadata bytes charged to the constructor's budget.
    /// Borrowed worksheet values are excluded from this amount.
    #[must_use]
    pub fn reserved_index_bytes(&self) -> u64 {
        self.index.reserved_bytes()
    }
}

impl ValueResolver for Resolver<'_> {
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SheetExtent>> {
        Ok(self
            .index
            .sheet_index(sheet, execution)?
            .map(|_| self.extent))
    }

    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> EvaluationResult<CellRead<'a>> {
        let Some(sheet) = self.index.sheet_index(sheet, execution)? else {
            return Ok(CellRead::Error(ScalarError::Reference));
        };
        if row >= self.extent.rows() || column >= self.extent.columns() {
            return Ok(CellRead::Error(ScalarError::Reference));
        }
        let Some(cell) = self.index.cell(sheet, row, column, execution)? else {
            return Ok(CellRead::Empty);
        };
        // A formula cache is not an authoritative calculated value. Keeping
        // this refusal before value dispatch covers every cached value type.
        if cell.formula.is_some() {
            return Ok(CellRead::Unsupported);
        }
        Ok(match &cell.value {
            CellValue::Empty => CellRead::Empty,
            CellValue::Text(text) => CellRead::Text(text),
            CellValue::Number(value)
            | CellValue::Percentage(value)
            | CellValue::Currency { value, .. } => {
                if value.is_finite() {
                    CellRead::Number(*value)
                } else {
                    CellRead::Error(ScalarError::Number)
                }
            },
            CellValue::Boolean(value) => CellRead::Logical(*value),
            CellValue::Date(_) | CellValue::Time(_) | CellValue::Unknown { .. } => {
                CellRead::Unsupported
            },
        })
    }

    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<usize>> {
        Ok(self.index.sheet_index(sheet, execution)?)
    }

    fn sheet_name_at(
        &self,
        index: usize,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<&str>> {
        Ok(self.index.sheet_name_at(index, execution)?)
    }

    fn sheet_count(&self, execution: &ExecutionContext) -> EvaluationResult<usize> {
        Ok(self.index.sheet_count(execution)?)
    }
}
