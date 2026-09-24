//! Bounded value evaluation for local OpenFormula references and arrays.
//!
//! The scalar evaluator in the parent module is intentionally kept as a
//! small, resolver-free hot path.  This module is the opt-in value path: it
//! adds a synchronous cell resolver, rectangular arrays, and reference
//! projection without changing the scalar evaluator's frame representation.
//!
//! A resolver is read only.  It must not perform I/O, refresh a cached value,
//! or recursively evaluate a formula cell.  Such cells are reported as
//! `CellRead::Unsupported` until a future calculation layer supplies those
//! capabilities.
//!
//! # Choosing an evaluation API
//!
//! Use [`super::evaluate_scalar`] for constant scalar expressions without a
//! worksheet. Use [`evaluate`] here for local references and arrays. Both
//! profiles implement scalar operators and the logical, bitwise, and number
//! representation function families; neither implements all OpenFormula functions.
//! This value profile also implements `MDETERM`, `MINVERSE`, `MMULT`, `MUNIT`,
//! and `TRANSPOSE`. Numeric matrix operations require finite numeric elements;
//! `TRANSPOSE` preserves element types. `MUNIT` truncates its size toward zero.
//! Singular inverses and non-finite arithmetic produce formula errors.
//! The numeric aggregate family includes `SUM`, `PRODUCT`, `SUMSQ`,
//! `SUMPRODUCT`, `SUMX2MY2`, `SUMX2PY2`, and `SUMXMY2`. `SUM` and `PRODUCT`
//! consume `NumberSequenceList` arguments, while `SUMSQ` consumes a single
//! `NumberSequence` shape and therefore rejects an explicit reference list.
//! The four remaining reducers consume forced arrays with equal row and
//! column counts; a scalar operand is a one-by-one array and is never
//! broadcast across a larger matrix.
//! The discrete reducer family adds `GCD` and `LCM` over streamed
//! `NumberSequenceList` values and `MULTINOMIAL` over a `NumberSequence`;
//! three-dimensional references remain one admissible reference while an
//! explicit reference union is rejected for `MULTINOMIAL`. `COMBIN`,
//! `COMBINA`, `FACT`, `FACTDOUBLE`, `EVEN`, `ODD`, `DELTA`, and `GESTEP` use
//! the scalar bridge and broadcast elementwise when their arguments are
//! arrays.
//! Order and rank functions include `MEDIAN`, `MODE`, `LARGE`, `SMALL`,
//! `PERCENTILE`, `PERCENTRANK`, `QUARTILE`, and `RANK`. They consume complete
//! data sequences while scalar query parameters can determine the result's
//! matrix shape. Selection retains only admitted numbers under the storage
//! budget and uses a fallible, work-charged sort.
//! Paired statistics (`CORREL`, `COVAR`, `PEARSON`, `RSQ`, `SLOPE`,
//! `INTERCEPT`, `STEYX`, and `FORECAST`) consume equal-shaped rectangular
//! data arrays without broadcasting or compacting each side independently.
//! Reference cells stream in aligned pairs into bounded exact state. Text,
//! Logical, and Empty members omit a pair; formula errors remain errors.
//! `FORECAST` lifts its scalar query in matrix mode and reuses an invariant
//! exact fit across query positions. Reference lists and 3-D data are refused
//! before scanning. INTERCEPT uses the included-constant profile documented
//! by the scalar evaluator.
//!
//! Both profiles support all 26 complex-number functions through
//! [`Value::Complex`]. `IMSUM` and `IMPRODUCT` consume arrays and ordered
//! reference sequences; referenced Empty and Logical cells are omitted.
//! `IMSUM` ignores unconvertible text. Numeric complex components remain
//! allocation-free values, including after an explicit owned conversion.
//! See [`super::complex`] for the finite representation and selected profile.
//! The scalar function bridge also covers the complete section 6.16
//! trigonometric/hyperbolic family, and matrix-mode calls apply those same
//! finite kernels elementwise to arrays and materialized references.  It also
//! covers the bounded real elementary family (`ABS`, `EXP`, `LN`, `LOG`,
//! `LOG10`, `MOD`, `POWER`, `QUOTIENT`, `SIGN`, `SQRT`, and `SQRTPI`) through
//! the same scalar kernels.
//!
//! [`Context::new`] defaults to [`Mode::Matrix`]. A bare reference in that mode
//! remains a first-class [`Value::Reference`] or [`Value::ReferenceList`] without
//! reading its cells. A consuming operation, such as adding a number to a range,
//! materializes its values. [`Mode::Scalar`] instead applies implicit intersection
//! at the caller's zero-based [`Position`]. An explicit Array rank parameter
//! to `LARGE` or `SMALL` produces a
//! same-shaped Array even in scalar mode; parentheses and a selected `IF`
//! branch preserve that result until a scalar-demanding consumer projects it.
//! Reference lists retain their distinct
//! type: they cannot be converted to a scalar or rectangular array. `AND` and `OR`
//! accept reference sequences and omit referenced text, logical, and empty cells.
//! Ordinary section 6.20 text functions apply scalar kernels at each broadcast
//! coordinate. They retain reference descriptors and borrow selected cell text
//! without materializing an input range; output arrays and constructed strings
//! remain subject to the caller's storage limits. Character positions count
//! Unicode scalars. Formatting uses the documented invariant profile.
//!
//! Matrix arithmetic broadcasts singleton dimensions. Lazy `IF`, `IFERROR`, and
//! `IFNA` evaluate selected cells only; unselected branches do not resolve missing
//! references. Formula errors are [`Value::Error`] results. Cancellation, resource
//! exhaustion, unsupported capabilities, and provider failures are
//! [`EvaluationFailure`] values and are not caught by formula error handlers.
//! Nested shape/value probes are limited to 32 active probes, further bounded
//! by the caller's Depth budget and [`Limits::with_max_stack_entries`].
//!
//! # Reusing worksheets and keeping a result
//!
//! [`crate::worksheet::formula::Resolver`] indexes physical repetition runs once.
//! Supply finite grid dimensions explicitly and reuse the index while the sheets
//! stay immutable. Borrowed results share the expression/resolver lifetime. Use
//! [`Evaluated::to_owned`] when a result must outlive them; the copy is separately
//! bounded and never reads or evaluates referenced cells.
//!
//! ```
//! use litchi_core::ExecutionContext;
//! use litchi_ods::{Cell, CellValue, Row, Sheet};
//! use litchi_ods::codec::formula::{
//!     expression::Expression,
//!     evaluation::value::{self, Context, Limits, Mode, OwnedEvaluated,
//!         OwnedValueView, Position, SheetExtent, Value},
//! };
//! use litchi_ods::worksheet::formula::Resolver;
//!
//! fn calculate(execution: &ExecutionContext) -> Result<OwnedEvaluated, Box<dyn std::error::Error>> {
//!     let owned = {
//!         let mut row = Row::new();
//!         row.push_cell(Cell::new(CellValue::Number(10.0), "10"))?;
//!         row.push_cell(Cell::new(CellValue::Number(20.0), "20"))?;
//!         let mut sheet = Sheet::new("Data")?;
//!         sheet.push_row(row)?;
//!         let sheets = [sheet];
//!         let resolver = Resolver::new(&sheets, SheetExtent::new(100, 26), execution)?;
//!         let expression = Expression::parse("=[.A1:.B1]+1")?;
//!         let context = Context::new(execution, Position::new("Data", 0, 0))
//!             .with_mode(Mode::Matrix);
//!         let limits = Limits::default();
//!         let result = value::evaluate(&expression, &resolver, &context, &limits)?;
//!         let array = result.as_array().expect("range arithmetic produces an array");
//!         assert_eq!((array.shape().rows(), array.shape().columns()), (1, 2));
//!         assert!(matches!(array.get(0), Some(Value::Number(11.0))));
//!         result.to_owned(execution, &limits)?
//!     }; // The expression, worksheet, and resolver are now dropped.
//!     assert!(matches!(owned.as_array().unwrap().get(1), Some(OwnedValueView::Number(21.0))));
//!     Ok(owned)
//! }
//! ```
//!
//! # Capability boundaries
//!
//! This API does not build a dependency graph, recalculate formula cells, spill
//! results into a sheet, or publish caches. The worksheet adapter refuses formula
//! cells even when they contain cached values. Dates, times, and unsupported cell
//! payloads also receive typed refusals. Names, labels, external sources,
//! host-defined/volatile functions, and remaining function families are not
//! resolved or executed. A custom [`Resolver`] must obey the same read-only,
//! finite, immutable-source contract. Caller cancellation and resource budgets
//! constrain preparation, evaluation, and optional owned conversion separately.

use super::{
    EvaluationContext, EvaluationFailure, EvaluationLimits, EvaluationResult, Evaluator,
    ScalarError, TextValue, WorkingValue, ensure_capacity, map_execution_error, parse_error,
};
use crate::codec::formula::{
    expression::ArrayDimensions,
    reference::{Address, EndpointValue, Reference},
};
use litchi_core::{Budget, ExecutionContext, Reservation, Resource, ResourceLimit, SourceVersion};
use std::{
    borrow::Cow,
    cmp::{max, min},
    fmt::Debug,
    num::NonZeroUsize,
    sync::Arc,
};

const VALUE_SCOPE: &str = "ods-formula-value-evaluation";
const DEFAULT_MAX_REFERENCE_CELLS: usize = 1_048_576;
const DEFAULT_MAX_REFERENCE_AREAS: usize = 65_536;
const VALUE_CHECK_CHUNK: usize = 4096;

// The scalar bridge is owned by the evaluator integration agent.  Keeping it
// private here gives the future VM one place to preserve Empty until the
// consuming operator/function chooses a target type.
mod aggregate;
mod complex;
mod conditional;
mod criteria;
mod database;
mod descriptive;
#[allow(dead_code)]
mod geometry;
mod matrix;
mod order;
mod owned;
mod paired;
mod references;
mod scalar;
mod statistical;
#[cfg(test)]
mod tests;
use geometry::Cuboid;
pub use owned::{OwnedArrayView, OwnedEvaluated, OwnedReferenceListView, OwnedValueView};

/// The caller's current cell, used for sheet-relative references and
/// implicit intersection.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Position<'a> {
    sheet: &'a str,
    row: usize,
    column: usize,
}

fn element_to_runtime(element: RuntimeElement<'_>) -> RuntimeValue<'_> {
    match element {
        RuntimeElement::Empty => RuntimeValue::Empty,
        // Missing here means an out-of-shape broadcast position, whose
        // formula value is #N/A.  A function's explicit missing parameter is
        // represented by RuntimeValue::Missing and is handled separately as
        // #VALUE! by the scalar bridge.
        RuntimeElement::Missing => {
            RuntimeValue::Scalar(WorkingValue::Error(ScalarError::NotAvailable))
        },
        RuntimeElement::Present(value) => RuntimeValue::Scalar(value),
    }
}

fn element_to_slot(element: RuntimeElement<'_>) -> scalar::Slot<'_> {
    match element {
        RuntimeElement::Empty => scalar::Slot::Empty,
        RuntimeElement::Missing => {
            scalar::Slot::Value(WorkingValue::Error(ScalarError::NotAvailable))
        },
        RuntimeElement::Present(value) => scalar::Slot::Value(value),
    }
}

fn slot_to_runtime(slot: scalar::Slot<'_>) -> RuntimeValue<'_> {
    match slot {
        scalar::Slot::Empty => RuntimeValue::Empty,
        scalar::Slot::Value(value) => RuntimeValue::Scalar(value),
    }
}

fn slot_to_element(slot: scalar::Slot<'_>) -> RuntimeElement<'_> {
    match slot {
        scalar::Slot::Empty => RuntimeElement::Empty,
        scalar::Slot::Value(value) => RuntimeElement::Present(value),
    }
}

fn broadcast_shape(left: Shape, right: Shape) -> Option<Shape> {
    Shape::new(
        max(left.rows(), right.rows()),
        max(left.columns(), right.columns()),
    )
    .ok()
}

fn is_reference_error(value: &RuntimeValue<'_>) -> bool {
    matches!(
        value,
        RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Reference))
    )
}

fn array_element_for<'a, 'value>(
    array: &'a RuntimeArrayValue<'value>,
    output: Shape,
    index: usize,
) -> Option<&'a RuntimeElement<'value>> {
    let source_index = projected_array_index(array.shape, output, index)?;
    array.cells.get(source_index)
}

fn projected_array_index(source: Shape, destination: Shape, index: usize) -> Option<usize> {
    let destination_columns = destination.columns();
    let row = index / destination_columns;
    let column = index % destination_columns;
    let source_cuboid =
        Cuboid::from_origin_extents([0, 0, 0], [1, source.rows(), source.columns()])?;
    let destination_cuboid =
        Cuboid::from_origin_extents([0, 0, 0], [1, destination.rows(), destination.columns()])?;
    let point = [0, row, column];
    let projected = if source.rows() == 1 && source.columns() == 1 {
        source_cuboid.project_2d_singleton(destination_cuboid, point)
    } else if source.rows() == 1 {
        source_cuboid.project_2d_row(destination_cuboid, point)
    } else if source.columns() == 1 {
        source_cuboid.project_2d_column(destination_cuboid, point)
    } else {
        source_cuboid.project_2d_matrix(destination_cuboid, point)
    }?;
    projected[1]
        .checked_mul(source.columns())?
        .checked_add(projected[2])
}

fn column_number(label: &str) -> EvaluationResult<usize> {
    if label.is_empty() {
        return Err(EvaluationFailure::InvalidExpression(
            "reference column is empty",
        ));
    }
    let mut value = 0usize;
    for byte in label.bytes() {
        if !byte.is_ascii_uppercase() {
            return Err(EvaluationFailure::InvalidExpression(
                "reference column is not uppercase ASCII",
            ));
        }
        value = value
            .checked_mul(26)
            .and_then(|value| value.checked_add(usize::from(byte - b'A' + 1)))
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference column overflows coordinate domain",
            ))?;
    }
    value
        .checked_sub(1)
        .ok_or(EvaluationFailure::InvalidExpression(
            "reference column is empty",
        ))
}

impl<'a> Position<'a> {
    /// Construct a zero-based current-cell position.
    #[must_use]
    pub const fn new(sheet: &'a str, row: usize, column: usize) -> Self {
        Self { sheet, row, column }
    }

    /// Current sheet name.
    #[must_use]
    pub const fn sheet(self) -> &'a str {
        self.sheet
    }

    /// Current zero-based row.
    #[must_use]
    pub const fn row(self) -> usize {
        self.row
    }

    /// Current zero-based column.
    #[must_use]
    pub const fn column(self) -> usize {
        self.column
    }
}

/// Evaluation mode for the value path.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Mode {
    /// Preserve array results.  Scalar operators broadcast over arrays.
    Matrix,
    /// Evaluate in scalar context, projecting array/reference operands through
    /// implicit intersection as operators and functions consume them.
    Scalar,
}

/// Context for a value evaluation.
#[derive(Clone, Copy, Debug)]
pub struct Context<'exec, 'position> {
    execution: &'exec ExecutionContext,
    position: Position<'position>,
    options: super::EvaluationOptions,
    mode: Mode,
}

impl<'exec, 'position> Context<'exec, 'position> {
    /// Create a matrix-mode context with the scalar evaluator's defaults.
    #[must_use]
    pub const fn new(execution: &'exec ExecutionContext, position: Position<'position>) -> Self {
        Self {
            execution,
            position,
            options: super::EvaluationOptions {
                text_case: super::TextCase::Sensitive,
            },
            mode: Mode::Matrix,
        }
    }

    /// Set matrix or scalar result mode.
    #[must_use]
    pub const fn with_mode(mut self, mode: Mode) -> Self {
        self.mode = mode;
        self
    }

    /// Set the existing scalar comparison options.
    #[must_use]
    pub const fn with_options(mut self, options: super::EvaluationOptions) -> Self {
        self.options = options;
        self
    }

    /// Caller-owned execution policy.
    #[must_use]
    pub const fn execution(self) -> &'exec ExecutionContext {
        self.execution
    }

    /// Current position.
    #[must_use]
    pub const fn position(self) -> Position<'position> {
        self.position
    }

    /// Selected value mode.
    #[must_use]
    pub const fn mode(self) -> Mode {
        self.mode
    }

    /// Scalar comparison/conversion options.
    #[must_use]
    pub const fn options(self) -> super::EvaluationOptions {
        self.options
    }
}

/// Finite worksheet dimensions used to expand whole-row and whole-column
/// references.  Both fields are logical counts, not last indices.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SheetExtent {
    rows: usize,
    columns: usize,
}

impl SheetExtent {
    /// Construct an extent from logical row and column counts.
    #[must_use]
    pub const fn new(rows: usize, columns: usize) -> Self {
        Self { rows, columns }
    }

    /// Logical row count.
    #[must_use]
    pub const fn rows(self) -> usize {
        self.rows
    }

    /// Logical column count.
    #[must_use]
    pub const fn columns(self) -> usize {
        self.columns
    }

    /// Return the checked number of logical cells in this sheet extent.
    #[must_use]
    pub const fn cell_count(self) -> Option<usize> {
        self.rows.checked_mul(self.columns)
    }
}

/// A cell value supplied by a read-only resolver.
#[derive(Clone, Copy, Debug, PartialEq)]
#[non_exhaustive]
pub enum CellRead<'a> {
    /// A physically absent or explicitly empty cell.
    Empty,
    /// A finite numeric value.
    Number(f64),
    /// A logical value.
    Logical(bool),
    /// UTF-8 text borrowed from the immutable source.
    Text(&'a str),
    /// A formula-level error value.
    Error(ScalarError),
    /// A valid cell type outside this evaluator's scalar profile.
    Unsupported,
}

/// Synchronous, immutable cell access for value evaluation.
///
/// Every operation receives the evaluation caller's execution policy. Providers
/// must use it for lookup work, allocation and cancellation; preparation may be
/// a separate operation with its own retained storage reservation.
/// Sheet names, extents and values must form one coherent immutable observation.
/// The ordered set exposed by `sheet_index`, `sheet_name_at` and `sheet_count`
/// must remain mutually consistent throughout evaluation and result borrowing.
pub trait Resolver {
    /// Return finite logical sheet dimensions.  This is required even when
    /// an evaluation only uses cell references so whole-axis references have
    /// a bounded expansion rule.
    fn sheet_extent(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SheetExtent>>;

    /// Read one zero-based logical coordinate.
    fn read_cell<'a>(
        &'a self,
        sheet: &str,
        row: usize,
        column: usize,
        execution: &ExecutionContext,
    ) -> EvaluationResult<CellRead<'a>>;

    /// Return the stable order of a sheet for a 3-D local reference.  `None`
    /// identifies a missing sheet; provider failures stay Rust-level errors.
    fn sheet_index(
        &self,
        sheet: &str,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<usize>>;

    /// Return the stable sheet name at an order position.  `None` identifies
    /// an index outside the provider's ordered sheet set.
    fn sheet_name_at(
        &self,
        index: usize,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<&str>>;

    /// Return the finite workbook sheet count used to bound ordered traversal.
    fn sheet_count(&self, execution: &ExecutionContext) -> EvaluationResult<usize>;

    /// Return the immutable source identity when the resolver has one.
    /// `None` is valid only when the provider guarantees coherent immutable
    /// observations without an identity fence, including borrowed result data.
    fn source_version(
        &self,
        execution: &ExecutionContext,
    ) -> EvaluationResult<Option<SourceVersion>> {
        execution.check().map_err(map_execution_error)?;
        Ok(None)
    }
}

/// Limits specific to the value evaluator.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    scalar: EvaluationLimits,
    max_array_cells: usize,
    max_reference_cells: usize,
    max_reference_areas: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            scalar: EvaluationLimits::default(),
            max_array_cells: super::DEFAULT_MAX_STACK_ENTRIES,
            max_reference_cells: DEFAULT_MAX_REFERENCE_CELLS,
            max_reference_areas: DEFAULT_MAX_REFERENCE_AREAS,
        }
    }
}

impl Limits {
    /// Set the aggregate work limit for one complete value evaluation.
    #[must_use]
    pub const fn with_max_steps(mut self, value: u64) -> Self {
        self.scalar = self.scalar.with_max_steps(value);
        self
    }

    /// Set the maximum bytes in one evaluated text value.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, value: usize) -> Self {
        self.scalar = self.scalar.with_max_text_bytes(value);
        self
    }

    /// Set aggregate evaluator-owned storage bytes.
    #[must_use]
    pub const fn with_max_storage_bytes(mut self, value: usize) -> Self {
        self.scalar = self.scalar.with_max_storage_bytes(value);
        self
    }

    /// Set the value/frame stack limit.
    #[must_use]
    pub const fn with_max_stack_entries(mut self, value: usize) -> Self {
        self.scalar = self.scalar.with_max_stack_entries(value);
        self
    }

    /// Set the maximum cells in one materialized rectangular array.
    #[must_use]
    pub const fn with_max_array_cells(mut self, value: usize) -> Self {
        self.max_array_cells = value;
        self
    }

    /// Bound both admitted reference geometry and cumulative provider reads.
    ///
    /// Each constructed reference result must fit this logical cell count;
    /// reference lists include duplicate areas in that count. Separately, all
    /// provider cell reads in one evaluation share this cumulative maximum.
    /// A bare matrix reference or scalar projection can therefore be refused
    /// for its full geometry before any provider cell is read.
    #[must_use]
    pub const fn with_max_reference_cells(mut self, value: usize) -> Self {
        self.max_reference_cells = value;
        self
    }

    /// Set the maximum reference records and resolved sheet planes retained
    /// at once. A 3-D reference can contain several physical sheet planes even
    /// when its public view represents them as one cuboid.
    #[must_use]
    pub const fn with_max_reference_areas(mut self, value: usize) -> Self {
        self.max_reference_areas = value;
        self
    }

    /// Scalar limits shared by individual element operations.
    #[must_use]
    pub const fn scalar_limits(self) -> EvaluationLimits {
        self.scalar
    }

    /// Maximum result-array cell count.
    #[must_use]
    pub const fn max_array_cells(self) -> usize {
        self.max_array_cells
    }

    /// Maximum admitted reference geometry and cumulative provider-read count.
    #[must_use]
    pub const fn max_reference_cells(self) -> usize {
        self.max_reference_cells
    }

    /// Maximum retained reference record or sheet-plane count.
    #[must_use]
    pub const fn max_reference_areas(self) -> usize {
        self.max_reference_areas
    }
}

/// Array dimensions exposed by an evaluated value.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Shape {
    rows: NonZeroUsize,
    columns: NonZeroUsize,
}

impl Shape {
    /// Construct checked nonempty dimensions.
    pub fn new(rows: usize, columns: usize) -> EvaluationResult<Self> {
        let rows = NonZeroUsize::new(rows)
            .ok_or(EvaluationFailure::InvalidExpression("array has zero rows"))?;
        let columns = NonZeroUsize::new(columns).ok_or(EvaluationFailure::InvalidExpression(
            "array has zero columns",
        ))?;
        Ok(Self { rows, columns })
    }

    /// Number of rows.
    #[must_use]
    pub const fn rows(self) -> usize {
        self.rows.get()
    }

    /// Number of columns.
    #[must_use]
    pub const fn columns(self) -> usize {
        self.columns.get()
    }

    /// Checked cell count.
    #[must_use]
    pub fn cell_count(self) -> Option<usize> {
        self.rows().checked_mul(self.columns())
    }
}

/// A resolved half-open sheet/row/column area.  Sheet bounds use the
/// resolver's stable ordered sheet indices; row and column bounds use
/// zero-based logical coordinates.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Area {
    sheet_start: usize,
    sheet_end: usize,
    row_start: usize,
    row_end: usize,
    column_start: usize,
    column_end: usize,
}

impl Area {
    /// Construct checked nonempty bounds from `[sheet, row, column]`
    /// starts and exclusive ends.
    #[must_use]
    pub const fn from_bounds(starts: [usize; 3], ends: [usize; 3]) -> Option<Self> {
        if starts[0] >= ends[0] || starts[1] >= ends[1] || starts[2] >= ends[2] {
            return None;
        }
        Some(Self {
            sheet_start: starts[0],
            sheet_end: ends[0],
            row_start: starts[1],
            row_end: ends[1],
            column_start: starts[2],
            column_end: ends[2],
        })
    }

    /// Inclusive starts as `[sheet, row, column]`.
    #[must_use]
    pub const fn starts(self) -> [usize; 3] {
        [self.sheet_start, self.row_start, self.column_start]
    }

    /// Exclusive ends as `[sheet, row, column]`.
    #[must_use]
    pub const fn ends(self) -> [usize; 3] {
        [self.sheet_end, self.row_end, self.column_end]
    }

    /// Checked extents as `[sheets, rows, columns]`.
    #[must_use]
    pub const fn extent(self) -> [usize; 3] {
        [
            self.sheet_end - self.sheet_start,
            self.row_end - self.row_start,
            self.column_end - self.column_start,
        ]
    }

    /// Checked number of logical cells.
    #[must_use]
    pub const fn cell_count(self) -> Option<usize> {
        let extent = self.extent();
        let sheet_rows = match extent[0].checked_mul(extent[1]) {
            Some(value) => value,
            None => return None,
        };
        sheet_rows.checked_mul(extent[2])
    }
}

/// One first-class resolved reference.  Direct AST references retain their
/// parsed lexical owner.  Range/intersection/union operations can produce a
/// view with no single owner while preserving every resolved area.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ReferenceView<'a> {
    reference: Option<&'a Reference>,
    areas: &'a [Area],
}

impl<'a> ReferenceView<'a> {
    /// Return the original parsed reference, when this view has one owner.
    #[must_use]
    pub const fn reference(self) -> Option<&'a Reference> {
        self.reference
    }

    /// Return resolved areas in source/operator order.
    #[must_use]
    pub const fn areas(self) -> &'a [Area] {
        self.areas
    }

    /// Number of resolved areas.
    #[must_use]
    pub const fn len(self) -> usize {
        self.areas.len()
    }

    /// Whether no resolved area was retained.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.areas.is_empty()
    }
}

/// Ordered first-class result for a reference-list/union expression.
#[derive(Clone, Copy, Debug)]
pub struct ReferenceListView<'a> {
    records: &'a [RuntimeReference<'a>],
}

impl PartialEq for ReferenceListView<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.iter().eq(other.iter())
    }
}

impl<'a> ReferenceListView<'a> {
    /// Number of references in source/operator order.
    #[must_use]
    pub const fn len(self) -> usize {
        self.records.len()
    }

    /// Whether the list contains no reference records.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.records.is_empty()
    }

    /// Return one reference view without cloning lexical metadata or areas.
    #[must_use]
    pub fn get(self, index: usize) -> Option<ReferenceView<'a>> {
        self.records.get(index).map(|record| ReferenceView {
            reference: record.reference,
            areas: &record.areas,
        })
    }

    /// Iterate over reference views in source/operator order.
    pub fn iter(self) -> impl Iterator<Item = ReferenceView<'a>> + 'a {
        self.records.iter().map(|record| ReferenceView {
            reference: record.reference,
            areas: &record.areas,
        })
    }
}

/// A borrowed value view.  Array elements are borrowed through
/// [`ArrayView::cell`], so inspection never clones the retained result.
/// Rust equality compares stored values and reference metadata structurally,
/// preserving array shape and record order. Array/list comparisons are linear
/// in the inspected cells and metadata, allocate nothing, and are separate
/// from OpenFormula's comparison and coercion rules.
#[derive(Clone, Copy, Debug, PartialEq)]
#[non_exhaustive]
pub enum Value<'a> {
    /// An absent or explicitly empty cell.
    Empty,
    /// A finite numeric value.
    Number(f64),
    /// A logical value.
    Logical(bool),
    /// Borrowed text.
    Text(&'a str),
    /// A formula-level error.
    Error(ScalarError),
    /// A finite OpenFormula complex number retained as a numeric pair.
    Complex(super::complex::Complex),
    /// A rectangular array view.
    Array(ArrayView<'a>),
    /// One resolved reference with its areas intact.
    Reference(ReferenceView<'a>),
    /// An ordered list of resolved references.
    ReferenceList(ReferenceListView<'a>),
}

impl Value<'_> {
    /// Return the array shape, if this is an array.
    #[must_use]
    pub const fn shape(self) -> Option<Shape> {
        match self {
            Self::Array(array) => Some(array.shape),
            Self::Empty
            | Self::Number(_)
            | Self::Logical(_)
            | Self::Text(_)
            | Self::Error(_)
            | Self::Complex(_)
            | Self::Reference(_)
            | Self::ReferenceList(_) => None,
        }
    }
}

/// Borrowed row-major view of an evaluated array.
#[derive(Clone, Copy, Debug)]
pub struct ArrayView<'a> {
    shape: Shape,
    cells: &'a [RuntimeElement<'a>],
}

impl PartialEq for ArrayView<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.shape == other.shape
            && self
                .cells
                .iter()
                .map(RuntimeElement::as_ref)
                .eq(other.cells.iter().map(RuntimeElement::as_ref))
    }
}

impl<'a> ArrayView<'a> {
    /// Array shape.
    #[must_use]
    pub const fn shape(self) -> Shape {
        self.shape
    }

    /// Number of cells.
    #[must_use]
    pub fn len(self) -> usize {
        self.cells.len()
    }

    /// Whether the array contains no cells.
    #[must_use]
    pub fn is_empty(self) -> bool {
        self.cells.is_empty()
    }

    /// Return one zero-based row-major cell.
    #[must_use]
    pub fn cell(self, row: usize, column: usize) -> Option<Value<'a>> {
        if row >= self.shape.rows() || column >= self.shape.columns() {
            return None;
        }
        let index = row.checked_mul(self.shape.columns())?.checked_add(column)?;
        self.cells.get(index).map(RuntimeElement::as_ref)
    }

    /// Return one row-major cell by linear index.
    #[must_use]
    pub fn get(self, index: usize) -> Option<Value<'a>> {
        self.cells.get(index).map(RuntimeElement::as_ref)
    }
}

/// Successful value evaluation retaining every result allocation until drop.
///
/// The expression and resolver intentionally share one borrow lifetime. This
/// permits a result to retain borrowed expression metadata and resolver cell
/// text without copying either source, and prevents the resolver from being
/// dropped while the result is live.
#[derive(Debug)]
pub struct Evaluated<'a> {
    value: RuntimeValue<'a>,
}

impl<'a> Evaluated<'a> {
    /// Borrow the complete result without cloning text or array cells.
    #[must_use]
    pub fn value(&self) -> Value<'_> {
        self.value.as_ref()
    }

    /// Return the result shape, if the result is an array.
    #[must_use]
    pub fn shape(&self) -> Option<Shape> {
        self.value.shape()
    }

    /// Return a borrowed array view, if the result is an array.
    #[must_use]
    pub fn as_array(&self) -> Option<ArrayView<'_>> {
        match &self.value {
            RuntimeValue::Array(array) => Some(ArrayView {
                shape: array.shape,
                cells: &array.cells,
            }),
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Areas(_) => None,
        }
    }

    /// Return a retained single-reference view, if the result contains one
    /// resolved reference record.
    #[must_use]
    pub fn as_reference(&self) -> Option<ReferenceView<'_>> {
        match &self.value {
            RuntimeValue::Areas(set) if !set.is_list && set.records.len() == 1 => {
                set.records.first().map(|record| ReferenceView {
                    reference: record.reference,
                    areas: &record.areas,
                })
            },
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Array(_)
            | RuntimeValue::Areas(_) => None,
        }
    }

    /// Return an ordered retained reference-list view, if the result is a
    /// reference or reference-list value.
    #[must_use]
    pub fn as_reference_list(&self) -> Option<ReferenceListView<'_>> {
        match &self.value {
            RuntimeValue::Areas(set) => Some(ReferenceListView {
                records: &set.records,
            }),
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Array(_) => None,
        }
    }

    /// Copy this result into lifetime-free, fallibly admitted storage.
    ///
    /// The returned value owns all text, array cells, and reference metadata,
    /// so it remains usable after the parsed expression and resolver are
    /// dropped.  The caller supplies the execution context and limits for
    /// the copy operation; no ambient workbook or resolver is consulted.
    /// Storage, text, array, reference-record/plane and work limits constrain
    /// the copy. Reference-cell read/admission and VM-stack limits do not apply:
    /// the copy neither reads cells nor expands or evaluates references.
    ///
    /// # Errors
    /// Returns typed cancellation, resource-limit or allocation failures.
    /// Failure drops all partial owned storage and leaves this result intact.
    pub fn to_owned(
        &self,
        execution: &ExecutionContext,
        limits: &Limits,
    ) -> EvaluationResult<OwnedEvaluated> {
        owned::from_evaluated(self, execution, limits)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Rect {
    row_start: usize,
    row_end: usize,
    column_start: usize,
    column_end: usize,
}

impl Rect {
    fn new(
        row_start: usize,
        row_end: usize,
        column_start: usize,
        column_end: usize,
    ) -> EvaluationResult<Self> {
        if row_start >= row_end || column_start >= column_end {
            return Err(EvaluationFailure::InvalidExpression(
                "reference rectangle is empty",
            ));
        }
        Ok(Self {
            row_start,
            row_end,
            column_start,
            column_end,
        })
    }

    fn cell(row: usize, column: usize) -> EvaluationResult<Self> {
        let row_end = row
            .checked_add(1)
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference row range overflow",
            ))?;
        let column_end = column
            .checked_add(1)
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference column range overflow",
            ))?;
        Self::new(row, row_end, column, column_end)
    }

    fn rows(self) -> usize {
        self.row_end - self.row_start
    }

    fn columns(self) -> usize {
        self.column_end - self.column_start
    }

    fn count(self) -> EvaluationResult<usize> {
        let count =
            self.rows()
                .checked_mul(self.columns())
                .ok_or(EvaluationFailure::InvalidExpression(
                    "reference cell count overflow",
                ))?;
        Ok(count)
    }

    fn intersection(self, other: Self) -> Option<Self> {
        let row_start = max(self.row_start, other.row_start);
        let row_end = min(self.row_end, other.row_end);
        let column_start = max(self.column_start, other.column_start);
        let column_end = min(self.column_end, other.column_end);
        (row_start < row_end && column_start < column_end).then_some(Self {
            row_start,
            row_end,
            column_start,
            column_end,
        })
    }

    fn bounding(self, other: Self) -> EvaluationResult<Self> {
        Self::new(
            min(self.row_start, other.row_start),
            max(self.row_end, other.row_end),
            min(self.column_start, other.column_start),
            max(self.column_end, other.column_end),
        )
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SheetRef<'a> {
    Current,
    Named(&'a str),
}

#[derive(Clone, Copy, Debug)]
struct RuntimeArea<'a> {
    sheet: SheetRef<'a>,
    sheet_index: usize,
    rect: Rect,
}

#[derive(Debug)]
struct RuntimeAreaSet<'a> {
    areas: Vec<RuntimeArea<'a>>,
    records: Vec<RuntimeReference<'a>>,
    /// Checked logical cells represented by the retained physical areas.
    /// Keeping this alongside the area vector lets reference union admission
    /// remain linear when a list is built left-associatively.
    cell_count: usize,
    area_reservation: Option<Reservation>,
    record_reservation: Option<Reservation>,
    is_list: bool,
}

/// Runtime ownership for the public reference-list view.  The evaluator
/// fills this record while resolving an AST reference; keeping it separate
/// from scalar/array slots prevents a reference operator from collapsing into
/// a formula error merely because it has more than one area.
#[derive(Debug)]
struct RuntimeReference<'a> {
    reference: Option<&'a Reference>,
    areas: Vec<Area>,
    reservation: Option<Reservation>,
}

impl<'a> RuntimeAreaSet<'a> {
    fn empty() -> Self {
        Self {
            areas: Vec::new(),
            records: Vec::new(),
            cell_count: 0,
            area_reservation: None,
            record_reservation: None,
            is_list: false,
        }
    }

    fn direct<R: Resolver + ?Sized>(
        reference: &'a Reference,
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<Self> {
        let mut result = Self::empty();
        result.ensure_record_capacity(1, evaluator)?;
        result.records.push(RuntimeReference {
            reference: Some(reference),
            areas: Vec::new(),
            reservation: None,
        });
        Ok(result)
    }

    fn derived<R: Resolver + ?Sized>(
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<Self> {
        let mut result = Self::empty();
        result.ensure_record_capacity(1, evaluator)?;
        result.records.push(RuntimeReference {
            reference: None,
            areas: Vec::new(),
            reservation: None,
        });
        Ok(result)
    }

    fn ensure_record_capacity<R: Resolver + ?Sized>(
        &mut self,
        additional: usize,
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<()> {
        ensure_capacity(
            &mut self.records,
            &mut self.record_reservation,
            additional,
            evaluator.limits.max_reference_areas,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula reference records",
        )
    }

    fn append_record_area<R: Resolver + ?Sized>(
        &mut self,
        area: &RuntimeArea<'a>,
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<()> {
        if self.records.is_empty() {
            self.ensure_record_capacity(1, evaluator)?;
            self.records.push(RuntimeReference {
                reference: None,
                areas: Vec::new(),
                reservation: None,
            });
        }
        let record = self
            .records
            .last_mut()
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference record missing",
            ))?;
        let public = Area::from_bounds(
            [
                area.sheet_index,
                area.rect.row_start,
                area.rect.column_start,
            ],
            [
                area.sheet_index
                    .checked_add(1)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference sheet index overflow",
                    ))?,
                area.rect.row_end,
                area.rect.column_end,
            ],
        )
        .ok_or(EvaluationFailure::InvalidExpression(
            "reference area is empty",
        ))?;
        if let Some(previous) = record.areas.last_mut() {
            let previous_end = previous.ends();
            let public_starts = public.starts();
            let public_ends = public.ends();
            if previous_end[0] == public_starts[0]
                && previous.starts()[1] == public_starts[1]
                && previous.ends()[1] == public_ends[1]
                && previous.starts()[2] == public_starts[2]
                && previous.ends()[2] == public_ends[2]
            {
                *previous = Area::from_bounds(
                    previous.starts(),
                    [public_ends[0], previous.ends()[1], previous.ends()[2]],
                )
                .ok_or(EvaluationFailure::InvalidExpression(
                    "reference area merge failed",
                ))?;
                return Ok(());
            }
        }
        ensure_capacity(
            &mut record.areas,
            &mut record.reservation,
            1,
            evaluator.limits.max_reference_areas,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula reference record areas",
        )?;
        record.areas.push(public);
        Ok(())
    }

    fn push<R: Resolver + ?Sized>(
        &mut self,
        area: RuntimeArea<'a>,
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<()> {
        self.push_raw(area, evaluator)?;
        self.append_record_area(&area, evaluator)
    }

    fn push_raw<R: Resolver + ?Sized>(
        &mut self,
        area: RuntimeArea<'a>,
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<()> {
        let area_cells = area.rect.count()?;
        let cell_count = self.cell_count.checked_add(area_cells).ok_or_else(|| {
            EvaluationFailure::ResourceLimit(evaluator.local_limit(
                Resource::Objects,
                u64::MAX,
                evaluator.limits.max_reference_cells,
            ))
        })?;
        if cell_count > evaluator.limits.max_reference_cells {
            return Err(EvaluationFailure::ResourceLimit(evaluator.local_limit(
                Resource::Objects,
                u64::try_from(cell_count).unwrap_or(u64::MAX),
                evaluator.limits.max_reference_cells,
            )));
        }
        ensure_capacity(
            &mut self.areas,
            &mut self.area_reservation,
            1,
            evaluator.limits.max_reference_areas,
            evaluator.execution,
            &evaluator.storage_budget,
            "formula reference areas",
        )?;
        self.areas.push(area);
        self.cell_count = cell_count;
        Ok(())
    }

    fn append_set<R: Resolver + ?Sized>(
        &mut self,
        mut other: RuntimeAreaSet<'a>,
        evaluator: &mut ValueEvaluator<'a, '_, '_, '_, R>,
    ) -> EvaluationResult<()> {
        for area in other.areas.drain(..) {
            evaluator.scalar.charge_work(1)?;
            self.push_raw(area, evaluator)?;
        }
        self.ensure_record_capacity(other.records.len(), evaluator)?;
        self.records.extend(other.records);
        self.is_list = true;
        // `other`'s reservations belong to the buffers that were drained or
        // moved above; dropping them here releases any capacity no longer
        // represented by this set.  Keep the count in the admission path so
        // the loop remains bounded even for a malformed provider result.
        Ok(())
    }
}

#[derive(Debug)]
struct RuntimeArrayValue<'a> {
    shape: Shape,
    cells: Vec<RuntimeElement<'a>>,
    _reservation: Option<Reservation>,
    origin: Option<RuntimeArea<'a>>,
    /// An explicitly array-valued result, rather than implicit scalar
    /// iteration. Pass-through consumers retain this publication behavior.
    preserve_scalar_result: bool,
}

#[derive(Debug)]
enum RuntimeValue<'a> {
    Empty,
    Missing,
    Scalar(WorkingValue<'a>),
    Array(RuntimeArrayValue<'a>),
    Areas(RuntimeAreaSet<'a>),
    /// A scalar-demand reference token.  The cell geometry is retained until
    /// its consuming operation projects it, which preserves sibling resolver
    /// ordering without allocating first-class reference metadata vectors.
    ScalarCell(RuntimeArea<'a>),
}

impl RuntimeValue<'_> {
    fn is_array_like(&self) -> bool {
        matches!(self, Self::Array(_) | Self::Areas(_))
    }
}

#[derive(Debug)]
enum RuntimeElement<'a> {
    Empty,
    Missing,
    Present(WorkingValue<'a>),
}

impl PartialEq for RuntimeElement<'_> {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (Self::Empty, Self::Empty) | (Self::Missing, Self::Missing) => true,
            (
                Self::Present(WorkingValue::Number(left)),
                Self::Present(WorkingValue::Number(right)),
            ) => left == right,
            (
                Self::Present(WorkingValue::Logical(left)),
                Self::Present(WorkingValue::Logical(right)),
            ) => left == right,
            (Self::Present(WorkingValue::Text(left)), Self::Present(WorkingValue::Text(right))) => {
                left.text == right.text
            },
            (
                Self::Present(WorkingValue::Error(left)),
                Self::Present(WorkingValue::Error(right)),
            ) => left == right,
            (
                Self::Present(WorkingValue::Complex(left)),
                Self::Present(WorkingValue::Complex(right)),
            ) => left == right,
            _ => false,
        }
    }
}

impl<'a> RuntimeElement<'a> {
    fn as_ref(&self) -> Value<'_> {
        match self {
            Self::Empty => Value::Empty,
            Self::Missing => Value::Error(ScalarError::NotAvailable),
            Self::Present(WorkingValue::Number(value)) => Value::Number(*value),
            Self::Present(WorkingValue::Logical(value)) => Value::Logical(*value),
            Self::Present(WorkingValue::Text(value)) => Value::Text(value.text.as_ref()),
            Self::Present(WorkingValue::Error(error)) => Value::Error(*error),
            Self::Present(WorkingValue::Complex(value)) => Value::Complex(*value),
        }
    }
}

impl<'a> RuntimeValue<'a> {
    fn as_ref(&self) -> Value<'_> {
        match self {
            Self::Empty => Value::Empty,
            Self::Missing => Value::Error(ScalarError::Value),
            Self::Scalar(WorkingValue::Number(value)) => Value::Number(*value),
            Self::Scalar(WorkingValue::Logical(value)) => Value::Logical(*value),
            Self::Scalar(WorkingValue::Text(value)) => Value::Text(value.text.as_ref()),
            Self::Scalar(WorkingValue::Error(error)) => Value::Error(*error),
            Self::Scalar(WorkingValue::Complex(value)) => Value::Complex(*value),
            // The evaluator consumes this private token before constructing
            // `Evaluated`; keep a defensive projection for any future frame
            // path that reaches the borrowed inspection boundary.
            Self::ScalarCell(_) => Value::Error(ScalarError::Value),
            Self::Array(array) => Value::Array(ArrayView {
                shape: array.shape,
                cells: &array.cells,
            }),
            Self::Areas(set) if !set.is_list && set.records.len() == 1 => {
                let record = &set.records[0];
                Value::Reference(ReferenceView {
                    reference: record.reference,
                    areas: &record.areas,
                })
            },
            Self::Areas(set) => Value::ReferenceList(ReferenceListView {
                records: &set.records,
            }),
        }
    }

    fn shape(&self) -> Option<Shape> {
        match self {
            Self::Array(array) => Some(array.shape),
            Self::Empty
            | Self::Missing
            | Self::Scalar(_)
            | Self::ScalarCell(_)
            | Self::Areas(_) => None,
        }
    }
}

/// Evaluate an already parsed expression with local cell and array support.
///
/// The expression and all resolver reads are evaluated against one immutable
/// source version.  The expression and resolver intentionally share the
/// returned borrow lifetime, so borrowed parsed metadata and cell text remain
/// valid for the whole result.
pub fn evaluate<'a, 'exec, 'position, R>(
    expression: &'a crate::codec::formula::expression::Expression,
    resolver: &'a R,
    context: &Context<'exec, 'position>,
    limits: &Limits,
) -> EvaluationResult<Evaluated<'a>>
where
    R: Resolver + ?Sized,
{
    context.execution.check().map_err(map_execution_error)?;
    let before = resolver.source_version(context.execution)?;
    let storage_budget = context.execution.budget().child(
        VALUE_SCOPE,
        litchi_core::Limits::new(
            u64::try_from(limits.scalar.max_storage_bytes).unwrap_or(u64::MAX),
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );

    // This local scalar context is borrowed only by the value VM.  Keeping
    // the established scalar Evaluator as a helper lets every scalar
    // operator/function retain one implementation and keeps it out of the
    // public value representation.
    let scalar_context = EvaluationContext::with_options(context.execution, context.options);
    let scalar = Evaluator::new(
        expression,
        &scalar_context,
        limits.scalar,
        storage_budget.clone(),
    );
    let mut evaluator = ValueEvaluator::new(
        expression,
        resolver,
        context,
        *limits,
        scalar,
        storage_budget,
    );
    let result = evaluator.run()?;
    let result = evaluator.finish_root(result)?;
    let after = resolver.source_version(context.execution)?;
    match (before, after) {
        (Some(expected), Some(observed)) if expected == observed => {},
        (Some(expected), Some(observed)) => {
            return Err(EvaluationFailure::SourceChanged { expected, observed });
        },
        (None, None) => {},
        _ => return Err(EvaluationFailure::SourceVersionAvailabilityChanged),
    }
    // Provider reads and the final source-version probe can outlive the last
    // VM frame.  Check cancellation once more before exposing the retained
    // result so a cancelled caller never receives a seemingly successful
    // evaluation after a long resolver operation.
    context.execution.check().map_err(map_execution_error)?;
    Ok(Evaluated { value: result })
}

struct ValueEvaluator<'expr, 'scalar, 'exec, 'position, R: ?Sized> {
    expression: &'expr crate::codec::formula::expression::Expression,
    resolver: &'expr R,
    execution: &'exec ExecutionContext,
    position: Position<'position>,
    mode: Mode,
    /// When a lazy matrix branch is evaluated, this is the outer matrix
    /// point currently being requested.  Scalar projection normally chooses
    /// the caller's top-left element for an unlocated array.  Carrying the
    /// requested point through a nested scalar evaluator lets an expression
    /// such as `0+[.A1:.B1]` read only the selected cell of the branch while
    /// retaining the same projection for nested arrays and operators.
    projection: Option<(Shape, usize)>,
    limits: Limits,
    scalar: Evaluator<'expr, 'scalar, 'exec>,
    storage_budget: Budget,
    frames: Vec<ValueFrame<'expr>>,
    values: Vec<RuntimeValue<'expr>>,
    frame_reservation: Option<Reservation>,
    value_reservation: Option<Reservation>,
    /// Saved evaluator contexts for array-typed arguments.  A matrix
    /// function can occur inside a lazy matrix branch, so replacing the
    /// active matrix continuation while visiting its argument must be
    /// reversible.  The entries own the saved continuation and are released
    /// before this vector's capacity reservation.
    argument_contexts: Vec<SavedValueContext<'expr, 'position>>,
    argument_context_reservation: Option<Reservation>,
    matrix: Option<MatrixState<'expr, 'position>>,
    shape_frames: Vec<ShapeFrame<'expr>>,
    shape_values: Vec<Option<Shape>>,
    shape_frame_reservation: Option<Reservation>,
    shape_value_reservation: Option<Reservation>,
    shape_masks: Vec<ShapeMask>,
    shape_mask_reservation: Option<Reservation>,
    demand_cache: Vec<DemandCacheEntry>,
    demand_cache_reservation: Option<Reservation>,
    /// Exact invariant regression fits shared by projected FORECAST queries.
    /// Drop the payload buffer before releasing its capacity reservation.
    forecast_cache: Vec<paired::ForecastCacheEntry>,
    forecast_cache_reservation: Option<Reservation>,
    /// One source shape per condition node.  Entries are keyed by source
    /// coordinates in this shape; callers with a later broadcast shape are
    /// projected back to it instead of creating one alias per output cell.
    condition_cache_shapes: Vec<ConditionCacheShape>,
    condition_cache_shape_reservation: Option<Reservation>,
    condition_cache: Vec<ConditionCacheEntry<'expr>>,
    condition_cache_reservation: Option<Reservation>,
    reference_cells_read: usize,
    probe_depth: usize,
}

#[derive(Clone, Copy)]
enum ValueFrame<'a> {
    Visit(super::Node<'a>),
    VisitScalar(super::Node<'a>),
    VisitArgument(super::Node<'a>),
    /// Visit an Array/ForceArray argument without scalar projection.
    VisitMatrixArgument(super::Node<'a>),
    /// Visit a Criterion argument without inheriting a lazy projected-cell
    /// demand. Criteria are scalar expressions even when the enclosing value
    /// evaluation is in matrix mode; direct arrays and multicell references
    /// remain first-class values and are rejected by the conditional kernel.
    VisitConditionalArgument(super::Node<'a>),
    /// Visit TRANSPOSE's array argument without inheriting a caller's
    /// per-cell projection, while retaining the caller's scalar/matrix mode.
    VisitTransposeArgument(super::Node<'a>),
    /// Visit a scalar parameter with the caller's implicit-intersection
    /// position but without an inherited matrix output projection.  MUNIT
    /// uses this to consume one `[0,0]` input element rather than invoking a
    /// separate matrix call for every output element.
    VisitScalarArgument(super::Node<'a>),
    /// Restore the context saved by one of the argument-entry frames.
    RestoreArgumentContext,
    Apply(super::Node<'a>),
    /// Finish a query using an invariant paired-data fit already retained.
    ApplyForecastQuery(super::Node<'a>),
    ApplyMatrix {
        node: super::Node<'a>,
        function: MatrixFunction,
    },
    IfAfterCondition(super::Node<'a>),
    IfErrorAfterValue(super::Node<'a>),
    MatrixStep,
    MatrixCollect,
    FinishArray {
        rows: usize,
        columns: usize,
        cells: usize,
    },
}

#[derive(Clone, Copy)]
enum MatrixFunction {
    Determinant,
    Inverse,
    Multiply,
    Unit,
    Transpose,
}

struct SavedValueContext<'expr, 'position> {
    mode: Mode,
    position: Position<'position>,
    projection: Option<(Shape, usize)>,
    matrix: Option<MatrixState<'expr, 'position>>,
}

/// A nested value probe may force array evaluation and enter another shape
/// planner. Keep the suspended planner's vectors with their reservations;
/// restoring by swapping also drops all probe-owned storage before its tokens.
struct SavedShapePlanner<'expr> {
    frames: Vec<ShapeFrame<'expr>>,
    values: Vec<Option<Shape>>,
    masks: Vec<ShapeMask>,
    frame_reservation: Option<Reservation>,
    value_reservation: Option<Reservation>,
    mask_reservation: Option<Reservation>,
}

#[derive(Clone, Copy)]
enum MatrixKind<'a> {
    If {
        then_branch: Option<super::Node<'a>>,
        else_branch: Option<super::Node<'a>>,
    },
    IfError {
        alternative: super::Node<'a>,
        catches_not_available: bool,
    },
}

struct MatrixState<'a, 'position> {
    kind: MatrixKind<'a>,
    shape: Shape,
    condition: Option<RuntimeArrayValue<'a>>,
    value: Option<RuntimeValue<'a>>,
    output: Vec<RuntimeElement<'a>>,
    output_reservation: Option<Reservation>,
    index: usize,
    restore_mode: Mode,
    restore_position: Position<'position>,
    restore_projection: Option<(Shape, usize)>,
    pending: bool,
    pending_cache: Option<MatrixCacheSlot>,
    then_cache: Option<RuntimeValue<'a>>,
    else_cache: Option<RuntimeValue<'a>>,
    alternative_cache: Option<RuntimeValue<'a>>,
    then_cacheable: bool,
    else_cacheable: bool,
    alternative_cacheable: bool,
}

#[derive(Clone, Copy)]
enum MatrixCacheSlot {
    Then,
    Else,
    Alternative,
}

struct DemandCacheEntry {
    node: usize,
    value: DemandCacheValue,
}

#[derive(Clone, Copy)]
enum DemandCacheValue {
    Empty,
    Missing,
    Number(f64),
    Logical(bool),
    Error(ScalarError),
    Complex(super::complex::Complex),
}

struct ConditionCacheEntry<'a> {
    node: usize,
    demand: Shape,
    index: usize,
    value: ConditionCacheValue<'a>,
}

struct ConditionCacheShape {
    node: usize,
    shape: Shape,
}

/// Result of looking up one condition coordinate.
///
/// `key` is the canonical source coordinate to use if the value has to be
/// evaluated and cached.  It is absent when the requested coordinate lies
/// outside the source shape already established for this condition.  Such a
/// coordinate must still be evaluated (it may be a real error value), but it
/// must not be inserted under the consumer's widened demand shape: that
/// would turn an out-of-shape `#N/A` into a cache entry for later output
/// coordinates.
struct ConditionCacheLookup<'a> {
    value: Option<RuntimeValue<'a>>,
    key: Option<(Shape, usize)>,
}

#[derive(Clone, Copy)]
enum ConditionCacheValue<'a> {
    Empty,
    Missing,
    Number(f64),
    Logical(bool),
    Text(&'a str),
    Error(ScalarError),
}

struct ShapeMask {
    shape: Shape,
    indexes: Vec<usize>,
    reservation: Option<Reservation>,
}

#[derive(Clone, Copy)]
enum ShapeFrame<'a> {
    Enter {
        node: super::Node<'a>,
        demand: ShapeDemand,
    },
    /// Resolve a contiguous reference expression in one bounded pass.  Its
    /// child operators are deliberately consumed by the reference evaluator
    /// rather than scheduled as independent shape walks; otherwise a long
    /// left-associated union would rebuild every prefix quadratically.
    Reference {
        node: super::Node<'a>,
        operator: super::InfixOperator,
    },
    /// The condition of a lazy handler is planned first.  Its shape value is
    /// left on the shape stack for this continuation, which then selects and
    /// schedules only the branches demanded by the discovered condition
    /// shape.  Keeping this as a frame makes nested handlers use the same
    /// bounded planner rather than a recursive or shallow probe.
    LazyCondition {
        node: super::Node<'a>,
        condition: super::Node<'a>,
        demand: ShapeDemand,
    },
    /// Finish planning an array-valued matrix function after its argument
    /// shapes have been collected.  The determinant is scalar and therefore
    /// bypasses this frame; the other operations have operation-specific
    /// shape rules rather than ordinary scalar broadcasting.
    MatrixExit {
        function: MatrixFunction,
        children: usize,
    },
    Exit {
        node: super::Node<'a>,
        children: usize,
        base: Option<Shape>,
    },
}

#[derive(Clone, Copy)]
enum ReferenceShapeFrame<'a> {
    Visit(super::Node<'a>),
    Apply(super::InfixOperator),
    FinishError {
        alternative: super::Node<'a>,
        catches_not_available: bool,
    },
}

/// Temporary stacks for reference geometry evaluation.
///
/// The vector is declared before its reservation so struct-field drop order
/// destroys the retained elements before releasing the corresponding budget
/// token.  This remains true for every early `?` return from the helper.
struct ReferenceShapeScratch<'a> {
    frames: Vec<ReferenceShapeFrame<'a>>,
    frame_reservation: Option<Reservation>,
    values: Vec<RuntimeValue<'a>>,
    value_reservation: Option<Reservation>,
}

impl<'a> ReferenceShapeScratch<'a> {
    fn new() -> Self {
        Self {
            frames: Vec::new(),
            frame_reservation: None,
            values: Vec::new(),
            value_reservation: None,
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ReferenceOperandKind {
    Scalar,
    Reference,
    ReferenceList,
    Array,
    Error(ScalarError),
    Unknown,
}

enum ReferenceKindValue<'a> {
    Known(ReferenceOperandKind),
    Runtime(RuntimeValue<'a>),
}

#[derive(Clone, Copy)]
enum ReferenceKindFrame<'a> {
    Visit(super::Node<'a>),
    Unary,
    Infix(super::InfixOperator),
    Function {
        node: super::Node<'a>,
        name: &'a str,
        count: usize,
    },
    IfAfterCondition {
        node: super::Node<'a>,
        condition: super::Node<'a>,
    },
    IfErrorAfterValue {
        node: super::Node<'a>,
    },
    UseChild,
}

/// The kind planner has the same ownership rule as the reference-value
/// scratch machine: each vector is declared before its matching reservation so
/// retained entries are dropped before budget tokens on every early return.
struct ReferenceKindScratch<'a> {
    frames: Vec<ReferenceKindFrame<'a>>,
    frame_reservation: Option<Reservation>,
    values: Vec<ReferenceKindValue<'a>>,
    value_reservation: Option<Reservation>,
}

impl<'a> ReferenceKindScratch<'a> {
    fn new() -> Self {
        Self {
            frames: Vec::new(),
            frame_reservation: None,
            values: Vec::new(),
            value_reservation: None,
        }
    }
}

#[derive(Clone, Copy)]
struct ShapeDemand {
    shape: Shape,
    mask: Option<usize>,
}

#[derive(Clone, Copy)]
struct LazyShapeChildren<'a> {
    first: Option<super::Node<'a>>,
    second: Option<super::Node<'a>>,
    first_mask: Option<usize>,
    second_mask: Option<usize>,
    len: usize,
    shape: Shape,
}

#[derive(Clone, Copy)]
enum UnaryOperation {
    Prefix(super::PrefixOperator),
    Postfix(super::PostfixOperator),
}

impl<'expr, 'scalar, 'exec, 'position, R> ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>
where
    R: Resolver + ?Sized,
{
    fn new(
        expression: &'expr crate::codec::formula::expression::Expression,
        resolver: &'expr R,
        context: &Context<'exec, 'position>,
        limits: Limits,
        scalar: Evaluator<'expr, 'scalar, 'exec>,
        storage_budget: Budget,
    ) -> Self {
        Self {
            expression,
            resolver,
            execution: context.execution,
            position: context.position,
            mode: context.mode,
            projection: None,
            limits,
            scalar,
            storage_budget,
            frames: Vec::new(),
            values: Vec::new(),
            frame_reservation: None,
            value_reservation: None,
            argument_contexts: Vec::new(),
            argument_context_reservation: None,
            matrix: None,
            shape_frames: Vec::new(),
            shape_values: Vec::new(),
            shape_frame_reservation: None,
            shape_value_reservation: None,
            shape_masks: Vec::new(),
            shape_mask_reservation: None,
            demand_cache: Vec::new(),
            demand_cache_reservation: None,
            forecast_cache: Vec::new(),
            forecast_cache_reservation: None,
            condition_cache_shapes: Vec::new(),
            condition_cache_shape_reservation: None,
            condition_cache: Vec::new(),
            condition_cache_reservation: None,
            reference_cells_read: 0,
            probe_depth: 0,
        }
    }

    fn run(&mut self) -> EvaluationResult<RuntimeValue<'expr>> {
        self.run_from(self.expression.root())
    }

    fn run_from(&mut self, root: super::Node<'expr>) -> EvaluationResult<RuntimeValue<'expr>> {
        self.run_from_with_demand(root, self.mode == Mode::Scalar)
    }

    fn run_from_without_scalar_demand(
        &mut self,
        root: super::Node<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        self.run_from_with_demand(root, false)
    }

    fn run_from_with_demand(
        &mut self,
        root: super::Node<'expr>,
        scalar_demand: bool,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let frame = if scalar_demand {
            ValueFrame::VisitScalar(root)
        } else {
            ValueFrame::Visit(root)
        };
        self.push_frame(frame)?;
        while let Some(frame) = self.frames.pop() {
            self.scalar.step()?;
            match frame {
                ValueFrame::Visit(node) => self.visit(node)?,
                ValueFrame::VisitScalar(node) => self.visit_scalar(node)?,
                ValueFrame::VisitArgument(node) => self.visit_argument(node)?,
                ValueFrame::VisitMatrixArgument(node) => self.visit_matrix_argument(node)?,
                ValueFrame::VisitConditionalArgument(node) => {
                    self.visit_conditional_argument(node)?
                },
                ValueFrame::VisitTransposeArgument(node) => self.visit_transpose_argument(node)?,
                ValueFrame::VisitScalarArgument(node) => self.visit_scalar_argument(node)?,
                ValueFrame::RestoreArgumentContext => self.restore_argument_context()?,
                ValueFrame::Apply(node) => self.apply(node)?,
                ValueFrame::ApplyForecastQuery(node) => {
                    let value = paired::apply_cached_forecast(self, node)?;
                    self.push_value(value)?;
                },
                ValueFrame::ApplyMatrix { node, function } => {
                    self.apply_matrix_function(node, function)?
                },
                ValueFrame::IfAfterCondition(node) => self.finish_if(node)?,
                ValueFrame::IfErrorAfterValue(node) => self.finish_if_error(node)?,
                ValueFrame::MatrixStep => self.matrix_step()?,
                ValueFrame::MatrixCollect => self.matrix_collect()?,
                ValueFrame::FinishArray {
                    rows,
                    columns,
                    cells,
                } => self.finish_array(rows, columns, cells)?,
            }
        }
        if self.values.len() != 1 {
            return Err(EvaluationFailure::InvalidExpression(
                "value evaluation did not produce exactly one value",
            ));
        }
        self.values
            .pop()
            .ok_or(EvaluationFailure::InvalidExpression(
                "value evaluation value stack was empty",
            ))
    }

    fn finish_root(&mut self, value: RuntimeValue<'expr>) -> EvaluationResult<RuntimeValue<'expr>> {
        match self.mode {
            // A bare reference is a first-class result in matrix mode.  A
            // consuming operator/function calls `materialize_for_array` (or
            // `project_scalar`) explicitly when it needs cell values.
            Mode::Matrix => match value {
                RuntimeValue::ScalarCell(_) => Err(EvaluationFailure::InvalidExpression(
                    "scalar cell token escaped matrix evaluation",
                )),
                value => Ok(value),
            },
            Mode::Scalar if matches!(&value, RuntimeValue::Array(array) if array.preserve_scalar_result) => {
                Ok(value)
            },
            Mode::Scalar => self.project_scalar(value),
        }
    }

    fn finish_array(&mut self, rows: usize, columns: usize, cells: usize) -> EvaluationResult<()> {
        let shape = Shape::new(rows, columns)?;
        if shape.cell_count() != Some(cells) || self.values.len() < cells {
            return Err(EvaluationFailure::InvalidExpression(
                "array value stack does not match its dimensions",
            ));
        }
        let (mut output, reservation) = self.new_element_vec(cells)?;
        // Children are visited in source order, but the value stack is LIFO.
        // The admitted output vector is filled backwards and reversed in
        // place, avoiding a second temporary cell allocation.
        for index in 0..cells {
            let value = self.pop_value()?;
            output.push(self.runtime_to_element(value)?);
            self.charge_cell_work(index)?;
        }
        output.reverse();
        self.push_value(RuntimeValue::Array(RuntimeArrayValue {
            shape,
            cells: output,
            _reservation: reservation,
            origin: None,
            preserve_scalar_result: false,
        }))
    }

    fn runtime_to_element(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        match value {
            RuntimeValue::Empty => Ok(RuntimeElement::Empty),
            RuntimeValue::Missing => Ok(RuntimeElement::Missing),
            RuntimeValue::Scalar(value) => Ok(RuntimeElement::Present(value)),
            RuntimeValue::ScalarCell(_) | RuntimeValue::Array(_) | RuntimeValue::Areas(_) => Err(
                EvaluationFailure::Unsupported(super::UnsupportedKind::Array),
            ),
        }
    }

    fn project_scalar(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match value {
            RuntimeValue::Scalar(_) | RuntimeValue::Empty => Ok(value),
            RuntimeValue::ScalarCell(area) => self.project_scalar_cell(area),
            RuntimeValue::Missing => Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Value,
            ))),
            RuntimeValue::Array(array) => {
                let element = self.project_array_element(&array)?;
                self.element_to_runtime(element)
            },
            RuntimeValue::Areas(areas) => self.project_area_value(&areas),
        }
    }

    fn project_scalar_cell(
        &mut self,
        area: RuntimeArea<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        // Keep the same current-sheet probe and read/conversion sequence as
        // `project_area_value`.  A scalar cell has one area, so the probe is
        // intentionally not used to reject a named endpoint on another
        // sheet; the ordinary single-area projection has the same behavior.
        let _ = self
            .resolver
            .sheet_index(self.position.sheet, self.execution)?;
        self.scalar.charge_work(1)?;
        let read =
            self.read_reference_cell(area.sheet, area.rect.row_start, area.rect.column_start)?;
        let element = self.read_to_element(read)?;
        self.element_to_runtime(element)
    }

    fn project_array_element(
        &mut self,
        array: &RuntimeArrayValue<'expr>,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        let Some(ref origin) = array.origin else {
            if let Some((output, index)) = self.projection {
                return array_element_for(array, output, index)
                    .map(|element| self.clone_element(element))
                    .transpose()?
                    .ok_or(EvaluationFailure::InvalidExpression("empty array"));
            }
            return array
                .cells
                .first()
                .map(|element| self.clone_element(element))
                .transpose()?
                .ok_or(EvaluationFailure::InvalidExpression("empty array"));
        };
        let current_sheet = self
            .resolver
            .sheet_index(self.position.sheet, self.execution)?;
        let Some(current_sheet) = current_sheet else {
            return Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::Reference,
            )));
        };
        if current_sheet != origin.sheet_index {
            return Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        }
        let row = self.position.row;
        let column = self.position.column;
        let row_offset = if array.shape.rows() == 1 {
            0
        } else if row >= origin.rect.row_start && row < origin.rect.row_end {
            row - origin.rect.row_start
        } else {
            return Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        };
        let column_offset = if array.shape.columns() == 1 {
            0
        } else if column >= origin.rect.column_start && column < origin.rect.column_end {
            column - origin.rect.column_start
        } else {
            return Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        };
        let index = row_offset
            .checked_mul(array.shape.columns())
            .and_then(|index| index.checked_add(column_offset))
            .ok_or(EvaluationFailure::InvalidExpression(
                "intersection index overflow",
            ))?;
        array
            .cells
            .get(index)
            .map(|element| self.clone_element(element))
            .transpose()?
            .ok_or(EvaluationFailure::InvalidExpression(
                "intersection index out of bounds",
            ))
    }

    fn element_to_runtime(
        &mut self,
        element: RuntimeElement<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match element {
            RuntimeElement::Empty => Ok(RuntimeValue::Empty),
            RuntimeElement::Missing => Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Value,
            ))),
            RuntimeElement::Present(value) => Ok(RuntimeValue::Scalar(value)),
        }
    }

    fn project_area_value(
        &mut self,
        areas: &RuntimeAreaSet<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let (area, row, column) = match self.scalar_reference_cell(areas)? {
            Ok(cell) => cell,
            Err(error) => return Ok(RuntimeValue::Scalar(WorkingValue::Error(error))),
        };
        let read = self.read_reference_cell(area.sheet, row, column)?;
        let element = self.read_to_element(read)?;
        self.element_to_runtime(element)
    }

    /// Select an implicitly intersected cell without reading or converting it.
    /// Consumers that distinguish source errors from generated conversion
    /// errors can then inspect the provider value before scalar normalization.
    fn scalar_reference_cell(
        &mut self,
        areas: &RuntimeAreaSet<'expr>,
    ) -> EvaluationResult<Result<(RuntimeArea<'expr>, usize, usize), ScalarError>> {
        if areas.is_list {
            // A reference-list is a distinct formula value.  Implicit
            // intersection is defined for one reference, and selecting the
            // first surviving list entry would silently change the value's
            // type (especially for `NOT`, `XOR`, and scalar functions).
            return Ok(Err(ScalarError::Value));
        }
        if areas.areas.is_empty() {
            return Ok(Err(ScalarError::Null));
        }
        let current_sheet = self
            .resolver
            .sheet_index(self.position.sheet, self.execution)?;
        let mut selected = None;
        for area in &areas.areas {
            self.scalar.charge_work(1)?;
            // A single-sheet vector intersects by its varying coordinate,
            // even when the caller is on another worksheet. A 3-D reference
            // additionally selects the caller's sheet plane.
            let matches_sheet = areas.areas.len() == 1
                || current_sheet.is_some_and(|index| index == area.sheet_index);
            if !matches_sheet {
                continue;
            }
            // The union of the caller's row and column cannot select exactly
            // one cell from a rectangle with both dimensions greater than one.
            if area.rect.rows() > 1 && area.rect.columns() > 1 {
                return Ok(Err(ScalarError::NotAvailable));
            }
            let row = if area.rect.rows() == 1 {
                Some(area.rect.row_start)
            } else if self.position.row >= area.rect.row_start
                && self.position.row < area.rect.row_end
            {
                Some(self.position.row)
            } else {
                None
            };
            let column = if area.rect.columns() == 1 {
                Some(area.rect.column_start)
            } else if self.position.column >= area.rect.column_start
                && self.position.column < area.rect.column_end
            {
                Some(self.position.column)
            } else {
                None
            };
            if let (Some(row), Some(column)) = (row, column) {
                if selected.is_some() {
                    return Ok(Err(ScalarError::NotAvailable));
                }
                selected = Some((*area, row, column));
            }
        }
        let Some((area, row, column)) = selected else {
            return Ok(Err(ScalarError::NotAvailable));
        };
        Ok(Ok((area, row, column)))
    }

    fn visit(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        self.visit_with_demand::<false>(node)
    }

    /// Visit a value that is known to be consumed as one scalar.  This demand
    /// is deliberately propagated only through grouping, unary and ordinary
    /// arithmetic nodes.  Reference operators and generic function arguments
    /// still use `visit`, because they require first-class areas/sequences.
    fn visit_scalar(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        self.visit_with_demand::<true>(node)
    }

    fn visit_with_demand<const SCALAR_DEMAND: bool>(
        &mut self,
        node: super::Node<'expr>,
    ) -> EvaluationResult<()> {
        if let Some((demand, index)) = self.projection {
            if let Some(value) = self.condition_cache_get(node, demand, index)?.value {
                return self.push_value(value);
            }
        }
        match node.kind() {
            super::Kind::Number => {
                let value = self.scalar.parse_number(node.text())?;
                self.push_scalar(value)
            },
            super::Kind::String => {
                let value = self.scalar.parse_string(node.text())?;
                self.push_scalar(value)
            },
            super::Kind::Error => {
                self.scalar.charge_bytes(node.text().len())?;
                self.push_scalar(WorkingValue::Error(parse_error(node.text())))
            },
            super::Kind::Parenthesized => {
                let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                    "parenthesized value node has no child",
                ))?;
                self.push_frame(if SCALAR_DEMAND {
                    ValueFrame::VisitScalar(child)
                } else {
                    ValueFrame::Visit(child)
                })
            },
            super::Kind::Prefix(_) | super::Kind::Postfix(_) => {
                let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                    "unary value node has no child",
                ))?;
                self.push_frame(ValueFrame::Apply(node))?;
                self.push_frame(if SCALAR_DEMAND {
                    ValueFrame::VisitScalar(child)
                } else {
                    ValueFrame::Visit(child)
                })
            },
            super::Kind::Infix(operator) => {
                let left = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                    "infix value node has no left child",
                ))?;
                let right = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                    "infix value node has no right child",
                ))?;
                self.push_frame(ValueFrame::Apply(node))?;
                let reference_operator = matches!(
                    operator,
                    super::InfixOperator::Range
                        | super::InfixOperator::Intersection
                        | super::InfixOperator::Union
                );
                let frame = |node| {
                    if SCALAR_DEMAND && !reference_operator {
                        ValueFrame::VisitScalar(node)
                    } else {
                        ValueFrame::Visit(node)
                    }
                };
                self.push_frame(frame(right))?;
                self.push_frame(frame(left))?;
                // Operators are handled in `apply`, including the three
                // reference operators.  Keep this match binding so a future
                // operation-specific admission can be added without another
                // AST walk.
                let _ = operator;
                Ok(())
            },
            super::Kind::Function { name } => self.visit_function(node, name),
            super::Kind::Reference(reference) => {
                if self.projection.is_some() && Self::cacheable_condition_reference(reference) {
                    if let Some(value) = self.demand_cache_get(node)? {
                        return self.push_value(value);
                    }
                }
                let value = if SCALAR_DEMAND && self.mode == Mode::Scalar {
                    self.reference_scalar_value(reference)?
                } else {
                    self.reference_value(reference)?
                };
                self.push_value(value)
            },
            super::Kind::Array(dimensions) => self.visit_array(node, dimensions),
            super::Kind::ArrayRow => Err(EvaluationFailure::InvalidExpression(
                "array row reached value evaluator directly",
            )),
            super::Kind::NamedExpression { .. } => Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::NamedExpression,
            )),
            super::Kind::QuotedLabel | super::Kind::AutomaticIntersection => Err(
                EvaluationFailure::Unsupported(super::UnsupportedKind::Label),
            ),
            super::Kind::Missing => self.push_value(RuntimeValue::Missing),
        }
    }

    fn visit_argument(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        if node.is_missing() {
            self.push_value(RuntimeValue::Missing)
        } else {
            self.visit(node)
        }
    }

    fn visit_matrix_argument(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        self.enter_argument_context(node, Mode::Matrix)
    }

    fn visit_conditional_argument(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        self.enter_argument_context(node, Mode::Scalar)
    }

    fn visit_transpose_argument(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        // A lazy matrix branch uses scalar execution to request one output
        // coordinate. Its Array-typed argument still belongs to the enclosing
        // matrix calculation; do not intersect the argument expression early.
        let mode = if self.projection.is_some() {
            Mode::Matrix
        } else {
            self.mode
        };
        self.enter_argument_context(node, mode)
    }

    fn visit_scalar_argument(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        // MUNIT does not implicitly iterate over its scalar parameter, but
        // the argument expression still uses the enclosing calculation mode.
        // Its first-element conversion happens after the expression returns.
        self.visit_transpose_argument(node)
    }

    fn matrix_scalar_parameter(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if self.mode == Mode::Scalar && self.projection.is_none() {
            return self.project_scalar(value);
        }
        match value {
            RuntimeValue::Array(mut array) => {
                let first = array
                    .cells
                    .first_mut()
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "empty matrix parameter",
                    ))?;
                let first = std::mem::replace(first, RuntimeElement::Empty);
                self.element_to_runtime(first)
            },
            RuntimeValue::Areas(areas) if areas.is_list => Ok(RuntimeValue::Scalar(
                WorkingValue::Error(ScalarError::Value),
            )),
            RuntimeValue::Areas(areas) => {
                if areas.areas.len() != 1 {
                    return Err(EvaluationFailure::Unsupported(
                        super::UnsupportedKind::Reference,
                    ));
                }
                let area = &areas.areas[0];
                self.scalar.charge_work(1)?;
                let read = self.read_reference_cell(
                    area.sheet,
                    area.rect.row_start,
                    area.rect.column_start,
                )?;
                let first = self.read_to_element(read)?;
                self.element_to_runtime(first)
            },
            other => Ok(other),
        }
    }

    fn enter_argument_context(
        &mut self,
        node: super::Node<'expr>,
        mode: Mode,
    ) -> EvaluationResult<()> {
        ensure_capacity(
            &mut self.argument_contexts,
            &mut self.argument_context_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value argument contexts",
        )?;
        ensure_capacity(
            &mut self.frames,
            &mut self.frame_reservation,
            2,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value argument context frames",
        )?;
        let saved = SavedValueContext {
            mode: self.mode,
            position: self.position,
            projection: self.projection,
            matrix: self.matrix.take(),
        };
        self.argument_contexts.push(saved);
        self.mode = mode;
        self.projection = None;
        self.frames.push(ValueFrame::RestoreArgumentContext);
        self.frames.push(ValueFrame::VisitArgument(node));
        Ok(())
    }

    fn restore_argument_context(&mut self) -> EvaluationResult<()> {
        let saved = self
            .argument_contexts
            .pop()
            .ok_or(EvaluationFailure::InvalidExpression(
                "matrix argument context is missing",
            ))?;
        self.mode = saved.mode;
        self.position = saved.position;
        self.projection = saved.projection;
        self.matrix = saved.matrix;
        Ok(())
    }

    fn visit_array(
        &mut self,
        node: super::Node<'expr>,
        dimensions: ArrayDimensions,
    ) -> EvaluationResult<()> {
        let columns = dimensions.columns().ok_or(EvaluationFailure::Unsupported(
            super::UnsupportedKind::Array,
        ))?;
        let rows = dimensions.rows();
        let cells = rows
            .checked_mul(columns)
            .ok_or(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::MAX,
                self.limits.max_array_cells,
            )))?;
        if cells > self.limits.max_array_cells {
            return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::try_from(cells).unwrap_or(u64::MAX),
                self.limits.max_array_cells,
            )));
        }
        if node.child_count() != rows {
            return Err(EvaluationFailure::InvalidExpression(
                "array row count disagrees with array dimensions",
            ));
        }

        // A lazy matrix branch carries one requested output point through
        // scalar evaluation. Validate and visit only that row/cell here;
        // validating every sibling for every demand would turn a linear
        // matrix into a quadratic walk before any value is needed.
        if self.mode == Mode::Scalar {
            if let Some((output, index)) = self.projection {
                let source = Shape::new(rows, columns)?;
                let Some(source_index) = projected_array_index(source, output, index) else {
                    return self.push_scalar(WorkingValue::Error(ScalarError::NotAvailable));
                };
                let row = source_index / columns;
                let column = source_index % columns;
                let row_node = node
                    .child(row)
                    .ok_or(EvaluationFailure::InvalidExpression("array row is missing"))?;
                if row_node.child_count() != columns {
                    return Err(EvaluationFailure::Unsupported(
                        super::UnsupportedKind::Array,
                    ));
                }
                return self.push_frame(ValueFrame::VisitArgument(row_node.child(column).ok_or(
                    EvaluationFailure::InvalidExpression("array cell is missing"),
                )?));
            }
        }

        for row in 0..rows {
            let row_node = node
                .child(row)
                .ok_or(EvaluationFailure::InvalidExpression("array row is missing"))?;
            if row_node.child_count() != columns {
                return Err(EvaluationFailure::Unsupported(
                    super::UnsupportedKind::Array,
                ));
            }
        }

        self.push_frame(ValueFrame::FinishArray {
            rows,
            columns,
            cells,
        })?;
        for row in (0..rows).rev() {
            let row_node = node
                .child(row)
                .ok_or(EvaluationFailure::InvalidExpression("array row is missing"))?;
            for column in (0..columns).rev() {
                self.push_frame(ValueFrame::Visit(row_node.child(column).ok_or(
                    EvaluationFailure::InvalidExpression("array cell is missing"),
                )?))?;
            }
        }
        Ok(())
    }

    fn visit_function(&mut self, node: super::Node<'expr>, name: &str) -> EvaluationResult<()> {
        self.scalar.charge_bytes(name.len())?;
        if let Some(function) = Self::matrix_function(name) {
            return self.visit_matrix_function(node, function);
        }
        if database::is_database_function(name)
            && node.child_count() != 3
            && !(node.child_count() == 2
                && (name.eq_ignore_ascii_case("DCOUNT") || name.eq_ignore_ascii_case("DCOUNTA")))
        {
            return self.push_scalar(WorkingValue::Error(ScalarError::Value));
        }
        if name.eq_ignore_ascii_case("IF") {
            let count = node.child_count();
            if !(1..=3).contains(&count) {
                return self.push_scalar(WorkingValue::Error(ScalarError::Value));
            }
            let condition = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "IF condition is missing",
            ))?;
            self.push_frame(ValueFrame::IfAfterCondition(node))?;
            return self.push_frame(ValueFrame::VisitArgument(condition));
        }
        if name.eq_ignore_ascii_case("IFERROR") || name.eq_ignore_ascii_case("IFNA") {
            if node.child_count() != 2 {
                return self.push_scalar(WorkingValue::Error(ScalarError::Value));
            }
            let value = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "error handler value is missing",
            ))?;
            self.push_frame(ValueFrame::IfErrorAfterValue(node))?;
            return self.push_frame(ValueFrame::VisitArgument(value));
        }

        // A coordinate-independent sequence aggregate may sit inside an
        // array-producing branch. Its arguments normally belong to the eager
        // function schedule, so looking in the cache only from
        // `apply_function` would resolve/read those arguments on every output
        // cell before discovering the cached result. Check before scheduling
        // the argument visits; a miss is inserted by `apply_function` after
        // its first evaluation.
        let is_sequence = name.eq_ignore_ascii_case("AND")
            || name.eq_ignore_ascii_case("OR")
            || complex::is_complex_sequence_function(name)
            || aggregate::is_aggregate_function(name)
            || statistical::is_statistical_function(name)
            || descriptive::is_descriptive_function(name)
            || paired::is_paired_function(name)
            || order::is_order_function(name)
            || conditional::is_conditional_function(name)
            || database::is_database_function(name);
        if self.projection.is_some() && is_sequence && self.cacheable_scalar_branch(node)? {
            if let Some(value) = self.demand_cache_get(node)? {
                return self.push_value(value);
            }
        }

        if paired::visit_cached_forecast(self, node, name)? {
            return Ok(());
        }

        let count = node.child_count();
        self.scalar
            .charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
        self.push_frame(ValueFrame::Apply(node))?;
        for index in (0..count).rev() {
            let child = node
                .child(index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "function argument is missing",
                ))?;
            // In a projected lazy matrix branch, a complex sequence consumes
            // the complete array/reference argument. Enter matrix context so
            // an inline array is not reduced to the one selected cell.
            let frame = if database::is_database_function(name) {
                if index == 1 && node.child_count() == 3 {
                    ValueFrame::VisitScalarArgument(child)
                } else {
                    ValueFrame::VisitMatrixArgument(child)
                }
            } else if paired::is_paired_function(name)
                && paired::matrix_argument(name, index, child)
            {
                ValueFrame::VisitMatrixArgument(child)
            } else if order::is_order_function(name) && order::matrix_argument(name, index, child) {
                ValueFrame::VisitMatrixArgument(child)
            } else if (statistical::is_statistical_function(name)
                || descriptive::is_descriptive_function(name))
                && Self::statistical_matrix_argument(child)
            {
                // A literal rectangular Array must retain all of its cells in
                // a projected branch. References and reference operators
                // already retain their area geometry through VisitArgument;
                // computed arguments stay in the enclosing projection so a
                // reducer cannot cache a position-dependent scalar.
                ValueFrame::VisitMatrixArgument(child)
            } else if conditional::is_range_argument(name, index, node.child_count()) {
                // Conditional aggregates consume ranges as references. Keep
                // their first-class area geometry intact even when the outer
                // expression is in scalar mode or is being projected one
                // cell at a time by a lazy matrix branch.
                ValueFrame::VisitMatrixArgument(child)
            } else if conditional::is_conditional_function(name) {
                ValueFrame::VisitConditionalArgument(child)
            } else if aggregate::is_matrix_aggregate_function(name) {
                ValueFrame::VisitMatrixArgument(child)
            } else if self.projection.is_some()
                && (complex::is_complex_sequence_function(name)
                    || aggregate::is_aggregate_function(name))
            {
                ValueFrame::VisitMatrixArgument(child)
            } else {
                ValueFrame::VisitArgument(child)
            };
            self.push_frame(frame)?;
        }
        Ok(())
    }

    fn matrix_function(name: &str) -> Option<MatrixFunction> {
        if !matrix::is_matrix_function(name) {
            return None;
        }
        if name.eq_ignore_ascii_case("MDETERM") {
            Some(MatrixFunction::Determinant)
        } else if name.eq_ignore_ascii_case("MINVERSE") {
            Some(MatrixFunction::Inverse)
        } else if name.eq_ignore_ascii_case("MMULT") {
            Some(MatrixFunction::Multiply)
        } else if name.eq_ignore_ascii_case("MUNIT") {
            Some(MatrixFunction::Unit)
        } else if name.eq_ignore_ascii_case("TRANSPOSE") {
            Some(MatrixFunction::Transpose)
        } else {
            None
        }
    }

    fn statistical_matrix_argument(mut node: super::Node<'expr>) -> bool {
        while matches!(node.kind(), super::Kind::Parenthesized) {
            let Some(child) = node.child(0) else {
                return false;
            };
            node = child;
        }
        matches!(node.kind(), super::Kind::Array(_))
    }

    fn visit_matrix_function(
        &mut self,
        node: super::Node<'expr>,
        function: MatrixFunction,
    ) -> EvaluationResult<()> {
        let count = node.child_count();
        self.scalar
            .charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
        self.push_frame(ValueFrame::ApplyMatrix { node, function })?;
        for index in (0..count).rev() {
            let child = node
                .child(index)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "matrix function argument is missing",
                ))?;
            let frame = match function {
                MatrixFunction::Unit => ValueFrame::VisitScalarArgument(child),
                MatrixFunction::Transpose => ValueFrame::VisitTransposeArgument(child),
                MatrixFunction::Determinant
                | MatrixFunction::Inverse
                | MatrixFunction::Multiply => ValueFrame::VisitMatrixArgument(child),
            };
            self.push_frame(frame)?;
        }
        Ok(())
    }

    fn finish_if(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        let raw_condition = self.pop_value()?;
        if self.mode == Mode::Matrix && raw_condition.is_array_like() {
            return self.start_matrix_if(node, raw_condition);
        }
        let condition = self.project_scalar(raw_condition)?;
        let condition = self.scalar_logical(condition)?;
        let condition = match condition {
            Ok(value) => value,
            Err(error) => return self.push_scalar(WorkingValue::Error(error)),
        };
        match node.child_count() {
            1 => self.push_scalar(WorkingValue::Logical(condition)),
            2 if !condition => self.push_scalar(WorkingValue::Logical(false)),
            2 => self.push_branch(node.child(1)),
            3 if condition => self.push_branch(node.child(1)),
            3 => self.push_branch(node.child(2)),
            _ => self.push_scalar(WorkingValue::Error(ScalarError::Value)),
        }
    }

    fn push_branch(&mut self, branch: Option<super::Node<'expr>>) -> EvaluationResult<()> {
        let branch = branch.ok_or(EvaluationFailure::InvalidExpression("IF branch is missing"))?;
        if branch.is_missing() {
            self.push_value(RuntimeValue::Empty)
        } else {
            self.push_frame(ValueFrame::Visit(branch))
        }
    }

    fn finish_if_error(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        let raw_value = self.pop_value()?;
        if self.mode == Mode::Matrix && raw_value.is_array_like() {
            return self.start_matrix_if_error(node, raw_value);
        }
        let value = self.project_scalar(raw_value)?;
        let catches = match &value {
            RuntimeValue::Scalar(WorkingValue::Error(error)) => {
                node.function_name().is_some_and(|name| {
                    name.eq_ignore_ascii_case("IFERROR")
                        || (name.eq_ignore_ascii_case("IFNA")
                            && *error == ScalarError::NotAvailable)
                })
            },
            _ => false,
        };
        if !catches {
            return self.push_value(value);
        }
        let alternative = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
            "error handler alternative is missing",
        ))?;
        self.push_frame(ValueFrame::VisitArgument(alternative))
    }

    fn start_matrix_if(
        &mut self,
        node: super::Node<'expr>,
        condition: RuntimeValue<'expr>,
    ) -> EvaluationResult<()> {
        let condition = self.materialize_for_array(condition)?;
        let RuntimeValue::Array(condition) = condition else {
            return self.push_scalar(WorkingValue::Error(ScalarError::Value));
        };
        if node.child_count() == 1 {
            return self.push_value(RuntimeValue::Array(condition));
        }
        let then_node = node.child(1);
        let else_node = node.child(2);
        let (then_mask, else_mask) =
            self.if_selected_branches(&condition, node.child_count() == 3)?;
        let then_selected = !then_mask.indexes.is_empty();
        let else_selected = !else_mask.indexes.is_empty();
        let shape = self.matrix_branch_shape(
            condition.shape,
            then_node,
            else_node,
            Some(then_mask),
            Some(else_mask),
        )?;
        let then_cacheable = if then_selected {
            then_node
                .map(|branch| self.cacheable_scalar_branch(branch))
                .transpose()?
                .unwrap_or(false)
        } else {
            false
        };
        let else_cacheable = if else_selected {
            else_node
                .map(|branch| self.cacheable_scalar_branch(branch))
                .transpose()?
                .unwrap_or(false)
        } else {
            false
        };
        let cells =
            shape
                .cell_count()
                .ok_or(EvaluationFailure::ResourceLimit(self.local_limit(
                    Resource::Objects,
                    u64::MAX,
                    self.limits.max_array_cells,
                )))?;
        let (output, reservation) = self.new_element_vec(cells)?;
        self.matrix = Some(MatrixState {
            kind: MatrixKind::If {
                then_branch: then_node,
                else_branch: else_node,
            },
            shape,
            condition: Some(condition),
            value: None,
            output,
            output_reservation: reservation,
            index: 0,
            restore_mode: self.mode,
            restore_position: self.position,
            restore_projection: self.projection,
            pending: false,
            pending_cache: None,
            then_cache: None,
            else_cache: None,
            alternative_cache: None,
            then_cacheable,
            else_cacheable,
            alternative_cacheable: false,
        });
        self.push_frame(ValueFrame::MatrixStep)
    }

    fn start_matrix_if_error(
        &mut self,
        node: super::Node<'expr>,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<()> {
        let value = self.materialize_for_array(value)?;
        let RuntimeValue::Array(value) = value else {
            return self.push_scalar(WorkingValue::Error(ScalarError::Value));
        };
        let value_shape = value.shape;
        let alternative = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
            "error handler alternative is missing",
        ))?;
        let catches_not_available = node
            .function_name()
            .is_some_and(|name| name.eq_ignore_ascii_case("IFNA"));
        let alternative_mask = self.array_contains_caught_error(&value, catches_not_available)?;
        let alternative_selected = !alternative_mask.indexes.is_empty();
        let shape = self.matrix_branch_shape(
            value_shape,
            Some(alternative),
            None,
            Some(alternative_mask),
            None,
        )?;
        let alternative_cacheable = if alternative_selected {
            self.cacheable_scalar_branch(alternative)?
        } else {
            false
        };
        let cells =
            shape
                .cell_count()
                .ok_or(EvaluationFailure::ResourceLimit(self.local_limit(
                    Resource::Objects,
                    u64::MAX,
                    self.limits.max_array_cells,
                )))?;
        let (output, reservation) = self.new_element_vec(cells)?;
        self.matrix = Some(MatrixState {
            kind: MatrixKind::IfError {
                alternative,
                catches_not_available,
            },
            shape,
            condition: None,
            value: Some(RuntimeValue::Array(value)),
            output,
            output_reservation: reservation,
            index: 0,
            restore_mode: self.mode,
            restore_position: self.position,
            restore_projection: self.projection,
            pending: false,
            pending_cache: None,
            then_cache: None,
            else_cache: None,
            alternative_cache: None,
            then_cacheable: false,
            else_cacheable: false,
            alternative_cacheable,
        });
        self.push_frame(ValueFrame::MatrixStep)
    }

    fn matrix_branch_shape(
        &mut self,
        initial: Shape,
        first: Option<super::Node<'expr>>,
        second: Option<super::Node<'expr>>,
        first_mask: Option<ShapeMask>,
        second_mask: Option<ShapeMask>,
    ) -> EvaluationResult<Shape> {
        // Keep both selected branches available while shape planning runs.
        // A later sibling can widen the common result, so every non-empty
        // branch is replanned at that widened demand before the shape is
        // published.  This is also what makes condition-cache source shapes
        // canonical across sibling growth rather than leaving an earlier
        // singleton result as the apparent value for new coordinates.
        let mut branches = [(first, first_mask), (second, second_mask)];
        let max_iterations = self
            .expression
            .node_count()
            .saturating_add(1)
            .min(self.limits.scalar.max_stack_entries.max(1));
        let mut shape = initial;
        for _ in 0..max_iterations {
            let mut grew = false;
            for (branch, mask_slot) in &mut branches {
                let Some(branch) = *branch else {
                    continue;
                };
                let Some(mut mask) = mask_slot.take() else {
                    continue;
                };
                if mask.indexes.is_empty() {
                    *mask_slot = Some(mask);
                    continue;
                }
                if mask.shape != shape {
                    let mask_shape = mask.shape;
                    mask = self.expand_shape_mask(mask, mask_shape, shape)?;
                }
                if mask.indexes.is_empty() {
                    *mask_slot = Some(mask);
                    continue;
                }
                let branch_shape = self.shape_hint_demand(branch, shape, Some(&mask))?;
                if let Some(branch_shape) = branch_shape {
                    let next = broadcast_shape(shape, branch_shape).ok_or(
                        EvaluationFailure::InvalidExpression("incompatible matrix branch shapes"),
                    )?;
                    if next != shape {
                        shape = next;
                        grew = true;
                    }
                }
                *mask_slot = Some(mask);
            }
            if !grew {
                return Ok(shape);
            }
        }
        Err(EvaluationFailure::ResourceLimit(self.local_limit(
            Resource::Objects,
            u64::try_from(max_iterations.saturating_add(1)).unwrap_or(u64::MAX),
            max_iterations,
        )))
    }

    fn expand_shape_mask(
        &mut self,
        mask: ShapeMask,
        old_shape: Shape,
        new_shape: Shape,
    ) -> EvaluationResult<ShapeMask> {
        if mask.shape != old_shape {
            return Err(EvaluationFailure::InvalidExpression(
                "shape demand mask has the wrong source shape",
            ));
        }
        if old_shape == new_shape {
            return Ok(mask);
        }
        let cells = new_shape
            .cell_count()
            .ok_or(EvaluationFailure::InvalidExpression(
                "expanded shape demand count overflow",
            ))?;
        let mut expanded = self.new_shape_mask(new_shape)?;
        for index in 0..cells {
            self.charge_cell_work(index)?;
            let Some(source_index) = projected_array_index(old_shape, new_shape, index) else {
                continue;
            };
            if self.shape_mask_contains(&mask, source_index)? {
                self.push_shape_mask_index(&mut expanded, index)?;
            }
        }
        Ok(expanded)
    }

    fn shape_mask_contains(&mut self, mask: &ShapeMask, index: usize) -> EvaluationResult<bool> {
        let mut first = 0;
        let mut last = mask.indexes.len();
        while first < last {
            let middle = first + (last - first) / 2;
            self.scalar.charge_work(1)?;
            match mask.indexes[middle].cmp(&index) {
                std::cmp::Ordering::Less => first = middle + 1,
                std::cmp::Ordering::Equal => return Ok(true),
                std::cmp::Ordering::Greater => last = middle,
            }
        }
        Ok(false)
    }

    fn if_selected_branches(
        &mut self,
        condition: &RuntimeArrayValue<'expr>,
        has_else: bool,
    ) -> EvaluationResult<(ShapeMask, ShapeMask)> {
        let shape = condition.shape;
        let cells =
            shape
                .cell_count()
                .ok_or(EvaluationFailure::ResourceLimit(self.local_limit(
                    Resource::Objects,
                    u64::MAX,
                    self.limits.max_array_cells,
                )))?;
        let mut then_mask = self.new_shape_mask(shape)?;
        let mut else_mask = self.new_shape_mask(shape)?;
        for index in 0..cells {
            self.charge_cell_work(index)?;
            let Some(element) = array_element_for(condition, shape, index) else {
                continue;
            };
            match self.borrowed_element_logical(element)? {
                Ok(true) => self.push_shape_mask_index(&mut then_mask, index)?,
                Ok(false) if has_else => self.push_shape_mask_index(&mut else_mask, index)?,
                Ok(false) | Err(_) => {},
            }
        }
        Ok((then_mask, else_mask))
    }

    fn array_contains_caught_error(
        &mut self,
        value: &RuntimeArrayValue<'expr>,
        catches_not_available: bool,
    ) -> EvaluationResult<ShapeMask> {
        let mut mask = self.new_shape_mask(value.shape)?;
        for (index, element) in value.cells.iter().enumerate() {
            self.charge_cell_work(index)?;
            if let RuntimeElement::Present(WorkingValue::Error(error)) = element {
                if !catches_not_available || *error == ScalarError::NotAvailable {
                    self.push_shape_mask_index(&mut mask, index)?;
                }
            }
        }
        Ok(mask)
    }

    fn borrowed_element_logical(
        &mut self,
        element: &RuntimeElement<'expr>,
    ) -> EvaluationResult<Result<bool, ScalarError>> {
        match element {
            RuntimeElement::Empty => scalar::logical(&mut self.scalar, scalar::Slot::Empty),
            RuntimeElement::Missing => Ok(Err(ScalarError::NotAvailable)),
            RuntimeElement::Present(WorkingValue::Number(value)) => scalar::logical(
                &mut self.scalar,
                scalar::Slot::Value(WorkingValue::Number(*value)),
            ),
            RuntimeElement::Present(WorkingValue::Logical(value)) => scalar::logical(
                &mut self.scalar,
                scalar::Slot::Value(WorkingValue::Logical(*value)),
            ),
            RuntimeElement::Present(WorkingValue::Error(error)) => Ok(Err(*error)),
            RuntimeElement::Present(WorkingValue::Text(value)) => scalar::logical(
                &mut self.scalar,
                scalar::Slot::Value(WorkingValue::Text(TextValue::borrowed(value.text.as_ref()))),
            ),
            RuntimeElement::Present(WorkingValue::Complex(_)) => Ok(Err(ScalarError::Value)),
        }
    }

    fn cacheable_scalar_branch(&mut self, mut node: super::Node<'expr>) -> EvaluationResult<bool> {
        // Unwrap grouping iteratively.  Shape and value traversal are both
        // flat-stack machines; cache classification must not reintroduce AST
        // recursion for a long parenthesis chain.
        loop {
            if !matches!(node.kind(), super::Kind::Parenthesized) {
                break;
            }
            self.scalar.charge_work(1)?;
            node = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "cacheable branch parentheses are empty",
            ))?;
        }
        match node.kind() {
            super::Kind::Number
            | super::Kind::String
            | super::Kind::Error
            | super::Kind::Missing => Ok(true),
            // Matrix functions consume their arguments in an explicit
            // array/scalar context, so a function whose complete subtree is
            // made only of literals is independent of the outer projected
            // cell.  MatrixState already retains one RuntimeValue per
            // selected branch; classify this case here so an array result is
            // evaluated once and indexed for every subsequent demand.
            super::Kind::Function { .. } if Self::matrix_function_node(node) => {
                self.cacheable_matrix_branch(node)
            },
            // Aggregate subtrees are reduced to one value and can be
            // invariant under an enclosing projected matrix demand when all
            // of their descendants are source-independent.  Use the
            // iterative matrix classifier here so deeply nested reductions
            // cannot recurse through the Rust call stack.
            super::Kind::Function { name } if aggregate::is_aggregate_function(name) => {
                self.cacheable_matrix_branch(node)
            },
            super::Kind::Function { name }
                if statistical::is_statistical_function(name)
                    || descriptive::is_descriptive_function(name) =>
            {
                self.cacheable_matrix_branch(node)
            },
            super::Kind::Function { name } if paired::is_paired_function(name) => {
                self.cacheable_conditional_criterion(node)
            },
            super::Kind::Function { name } if order::is_order_function(name) => {
                self.cacheable_order_branch(node)
            },
            super::Kind::Function { name } if conditional::is_conditional_function(name) => {
                self.cacheable_conditional_branch(node)
            },
            super::Kind::Function { name } if database::is_database_function(name) => {
                for index in 0..node.child_count() {
                    self.scalar.charge_work(1)?;
                    let Some(child) = node.child(index) else {
                        return Ok(false);
                    };
                    if !self.cacheable_sequence_operand(child)?
                        && !self.cacheable_matrix_branch(child)?
                    {
                        return Ok(false);
                    }
                }
                Ok(true)
            },
            super::Kind::Function { name }
                if name.eq_ignore_ascii_case("AND")
                    || name.eq_ignore_ascii_case("OR")
                    || complex::is_complex_sequence_function(name) =>
            {
                let mut cacheable = true;
                for index in 0..node.child_count() {
                    // Only a scalar literal or a fixed local reference is
                    // safe to retain across matrix demands. The sequence
                    // aggregate consumes the complete reference descriptor,
                    // rather than the projected cell at the current output.
                    self.scalar.charge_work(1)?;
                    let Some(child) = node.child(index) else {
                        return Ok(false);
                    };
                    if !self.cacheable_sequence_operand(child)? {
                        cacheable = false;
                    }
                }
                Ok(cacheable)
            },
            _ => Ok(false),
        }
    }

    fn matrix_function_node(node: super::Node<'expr>) -> bool {
        node.function_name()
            .is_some_and(|name| Self::matrix_function(name).is_some())
    }

    /// Order reducers consume complete sequence arguments, while their
    /// scalar parameters may still depend on the projected output position.
    /// Classify those two contexts independently so a reducer with a local
    /// reference parameter is not inserted into the scalar demand cache.
    fn cacheable_order_branch(&mut self, node: super::Node<'expr>) -> EvaluationResult<bool> {
        // Reuse the iterative context-aware walk: calling the scalar branch
        // classifier recursively for nested parameters would bypass the VM's
        // bounded traversal stack.
        self.cacheable_conditional_criterion(node)
    }

    /// Check that a matrix-function branch has no source-dependent operand.
    ///
    /// This is deliberately a small iterative prepass. MatrixState owns the
    /// resulting value, so the cache is safe only when every descendant is a
    /// literal or a deterministic operator/function over fixed local
    /// references. Source references, names, labels, and automatic
    /// intersections remain rejected. The local reservation is declared
    /// before the scratch vector so the vector is dropped before its budget
    /// token on every return path.
    fn cacheable_matrix_branch(&mut self, root: super::Node<'expr>) -> EvaluationResult<bool> {
        let mut reservation = None;
        let mut nodes = Vec::new();
        ensure_capacity(
            &mut nodes,
            &mut reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value matrix branch cacheability frames",
        )?;
        nodes.push(root);
        while let Some(node) = nodes.pop() {
            self.scalar.charge_work(1)?;
            match node.kind() {
                super::Kind::Reference(reference) => {
                    if !Self::cacheable_reference(reference) {
                        return Ok(false);
                    }
                },
                super::Kind::NamedExpression { .. }
                | super::Kind::AutomaticIntersection
                | super::Kind::QuotedLabel { .. } => return Ok(false),
                super::Kind::Number
                | super::Kind::String
                | super::Kind::Error
                | super::Kind::Missing => {},
                super::Kind::Parenthesized
                | super::Kind::Prefix(_)
                | super::Kind::Postfix(_)
                | super::Kind::Infix(_)
                | super::Kind::Function { .. }
                | super::Kind::Array(_)
                | super::Kind::ArrayRow => {
                    if let super::Kind::Function { name } = node.kind() {
                        if order::is_order_function(name)
                            || matches!(Self::matrix_function(name), Some(MatrixFunction::Unit))
                        {
                            // These functions have scalar parameter slots.
                            // A fixed literal parameter remains invariant,
                            // while a projected multicell reference does not.
                            if !self.cacheable_conditional_criterion(node)? {
                                return Ok(false);
                            }
                            continue;
                        }
                    }
                    let count = node.child_count();
                    ensure_capacity(
                        &mut nodes,
                        &mut reservation,
                        count,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value matrix branch cacheability frames",
                    )?;
                    for index in (0..count).rev() {
                        let Some(child) = node.child(index) else {
                            return Ok(false);
                        };
                        nodes.push(child);
                    }
                },
            }
        }
        Ok(true)
    }

    /// Check whether a conditional reducer is invariant under an enclosing
    /// projected matrix demand. Range slots are first-class matrix arguments,
    /// but a computed criterion runs in the surrounding scalar mode after its
    /// projection is cleared. A multi-cell reference inside ordinary scalar
    /// arithmetic can therefore implicitly intersect at the current output
    /// coordinate. Keep that case out of the demand cache while retaining
    /// cache reuse for literals, singleton references, and sequence reducers
    /// such as `"<"&SUM(reference)`.
    fn cacheable_conditional_branch(&mut self, node: super::Node<'expr>) -> EvaluationResult<bool> {
        let Some(name) = node.function_name() else {
            return Ok(false);
        };
        let count = node.child_count();
        for index in 0..count {
            let Some(child) = node.child(index) else {
                return Ok(false);
            };
            if conditional::is_range_argument(name, index, count) {
                if !self.cacheable_matrix_branch(child)? {
                    return Ok(false);
                }
            } else if !self.cacheable_conditional_criterion(child)? {
                return Ok(false);
            }
        }
        Ok(true)
    }

    /// Iteratively classify one computed Criterion expression. `full_reference`
    /// is set only while visiting an argument of a sequence/matrix reducer
    /// that consumes a reference's complete geometry. It deliberately does
    /// not flow through ordinary operators, whose scalar mode may project a
    /// multi-cell reference at the current output position.
    fn cacheable_conditional_criterion(
        &mut self,
        root: super::Node<'expr>,
    ) -> EvaluationResult<bool> {
        let mut reservation = None;
        let mut nodes = Vec::new();
        ensure_capacity(
            &mut nodes,
            &mut reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value conditional criterion cacheability frames",
        )?;
        nodes.push((root, true));
        while let Some((node, full_reference)) = nodes.pop() {
            self.scalar.charge_work(1)?;
            match node.kind() {
                super::Kind::Number
                | super::Kind::String
                | super::Kind::Error
                | super::Kind::Missing => {},
                super::Kind::Reference(reference) => {
                    let invariant = match reference {
                        Reference::Error => true,
                        Reference::Source { .. } => false,
                        Reference::Local(address) => match address {
                            Address::Cell(endpoint)
                                if matches!(&endpoint.value, EndpointValue::Cell(_)) =>
                            {
                                true
                            },
                            Address::Cell(_)
                            | Address::Cells(_, _)
                            | Address::Columns(_, _)
                            | Address::Rows(_, _) => full_reference,
                        },
                    };
                    if !invariant {
                        return Ok(false);
                    }
                },
                super::Kind::Parenthesized => {
                    let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                        "conditional criterion parentheses are empty",
                    ))?;
                    ensure_capacity(
                        &mut nodes,
                        &mut reservation,
                        1,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value conditional criterion cacheability frames",
                    )?;
                    nodes.push((child, full_reference));
                },
                super::Kind::Prefix(_) | super::Kind::Postfix(_) => {
                    let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                        "conditional criterion unary operand is missing",
                    ))?;
                    ensure_capacity(
                        &mut nodes,
                        &mut reservation,
                        1,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value conditional criterion cacheability frames",
                    )?;
                    nodes.push((child, false));
                },
                super::Kind::Infix(operator) => {
                    if matches!(
                        operator,
                        super::InfixOperator::Range
                            | super::InfixOperator::Intersection
                            | super::InfixOperator::Union
                    ) {
                        return Ok(false);
                    }
                    let left = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                        "conditional criterion left operand is missing",
                    ))?;
                    let right = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                        "conditional criterion right operand is missing",
                    ))?;
                    ensure_capacity(
                        &mut nodes,
                        &mut reservation,
                        2,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value conditional criterion cacheability frames",
                    )?;
                    nodes.push((right, false));
                    nodes.push((left, false));
                },
                super::Kind::Array(_) | super::Kind::ArrayRow => {
                    let count = node.child_count();
                    ensure_capacity(
                        &mut nodes,
                        &mut reservation,
                        count,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value conditional criterion cacheability frames",
                    )?;
                    for index in (0..count).rev() {
                        let Some(child) = node.child(index) else {
                            return Ok(false);
                        };
                        nodes.push((child, false));
                    }
                },
                super::Kind::Function { name } => {
                    if conditional::is_conditional_function(name) {
                        // The nested reducer has its own range/criterion
                        // argument contexts; keep this outer cache decision
                        // conservative rather than sharing a context-sensitive
                        // result across projected positions.
                        return Ok(false);
                    }
                    let full_arguments = aggregate::is_aggregate_function(name)
                        || statistical::is_statistical_function(name)
                        || descriptive::is_descriptive_function(name)
                        || complex::is_complex_sequence_function(name)
                        || name.eq_ignore_ascii_case("AND")
                        || name.eq_ignore_ascii_case("OR")
                        // MUNIT's size argument is scalar and may project a
                        // multicell reference at the current output cell.
                        // The other matrix functions consume complete arrays.
                        || Self::matrix_function(name).is_some_and(
                            |function| !matches!(function, MatrixFunction::Unit),
                        );
                    let count = node.child_count();
                    ensure_capacity(
                        &mut nodes,
                        &mut reservation,
                        count,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value conditional criterion cacheability frames",
                    )?;
                    for index in (0..count).rev() {
                        let Some(child) = node.child(index) else {
                            return Ok(false);
                        };
                        let child_full = if paired::is_paired_function(name) {
                            paired::criterion_full_argument(name, index, child)
                        } else if order::is_order_function(name) {
                            order::criterion_full_argument(name, index, child)
                        } else {
                            full_arguments
                        };
                        nodes.push((child, child_full));
                    }
                },
                super::Kind::NamedExpression { .. }
                | super::Kind::QuotedLabel { .. }
                | super::Kind::AutomaticIntersection => return Ok(false),
            }
        }
        Ok(true)
    }

    fn cacheable_sequence_operand(
        &mut self,
        mut node: super::Node<'expr>,
    ) -> EvaluationResult<bool> {
        loop {
            if !matches!(node.kind(), super::Kind::Parenthesized) {
                break;
            }
            self.scalar.charge_work(1)?;
            node = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "sequence operand parentheses are empty",
            ))?;
        }
        Ok(match node.kind() {
            super::Kind::Reference(reference) => Self::cacheable_reference(reference),
            _ => self.cacheable_scalar_leaf(node),
        })
    }

    fn cacheable_scalar_leaf(&self, node: super::Node<'expr>) -> bool {
        matches!(
            node.kind(),
            super::Kind::Number | super::Kind::String | super::Kind::Error | super::Kind::Missing
        )
    }

    fn cacheable_reference(reference: &Reference) -> bool {
        match reference {
            Reference::Error => true,
            Reference::Source { .. } => false,
            // The value evaluator resolves endpoint coordinates from the
            // parsed lexeme. In sequence-aggregate context the same area is
            // therefore invariant across matrix demands, including legacy
            // relative spellings without `$` markers.
            Reference::Local(
                Address::Cell(_)
                | Address::Cells(_, _)
                | Address::Columns(_, _)
                | Address::Rows(_, _),
            ) => true,
        }
    }

    fn cacheable_condition_reference(reference: &Reference) -> bool {
        matches!(
            reference,
            Reference::Error | Reference::Local(Address::Cell(_))
        )
    }

    fn demand_cache_position(
        &mut self,
        node: super::Node<'expr>,
    ) -> EvaluationResult<(usize, bool)> {
        let node = node.arena_index();
        let mut first = 0;
        let mut last = self.demand_cache.len();
        while first < last {
            let middle = first + (last - first) / 2;
            self.scalar.charge_work(1)?;
            match self.demand_cache[middle].node.cmp(&node) {
                std::cmp::Ordering::Less => first = middle + 1,
                std::cmp::Ordering::Equal => return Ok((middle, true)),
                std::cmp::Ordering::Greater => last = middle,
            }
        }
        Ok((first, false))
    }

    fn demand_cache_get(
        &mut self,
        node: super::Node<'expr>,
    ) -> EvaluationResult<Option<RuntimeValue<'expr>>> {
        let (index, present) = self.demand_cache_position(node)?;
        if !present {
            return Ok(None);
        }
        let value = self.demand_cache[index].value;
        Ok(Some(match value {
            DemandCacheValue::Empty => RuntimeValue::Empty,
            DemandCacheValue::Missing => RuntimeValue::Missing,
            DemandCacheValue::Number(value) => RuntimeValue::Scalar(WorkingValue::Number(value)),
            DemandCacheValue::Logical(value) => RuntimeValue::Scalar(WorkingValue::Logical(value)),
            DemandCacheValue::Error(error) => RuntimeValue::Scalar(WorkingValue::Error(error)),
            DemandCacheValue::Complex(value) => RuntimeValue::Scalar(WorkingValue::Complex(value)),
        }))
    }

    fn demand_cache_put(
        &mut self,
        node: super::Node<'expr>,
        value: &RuntimeValue<'expr>,
    ) -> EvaluationResult<()> {
        let cache_value = match value {
            RuntimeValue::Empty => DemandCacheValue::Empty,
            RuntimeValue::Missing => DemandCacheValue::Missing,
            RuntimeValue::Scalar(WorkingValue::Number(value)) => DemandCacheValue::Number(*value),
            RuntimeValue::Scalar(WorkingValue::Logical(value)) => DemandCacheValue::Logical(*value),
            RuntimeValue::Scalar(WorkingValue::Error(error)) => DemandCacheValue::Error(*error),
            RuntimeValue::Scalar(WorkingValue::Complex(value)) => DemandCacheValue::Complex(*value),
            RuntimeValue::Scalar(WorkingValue::Text(_))
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Array(_)
            | RuntimeValue::Areas(_) => return Ok(()),
        };
        let (index, present) = self.demand_cache_position(node)?;
        if present {
            self.demand_cache[index].value = cache_value;
            return Ok(());
        }
        ensure_capacity(
            &mut self.demand_cache,
            &mut self.demand_cache_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value demand cache",
        )?;
        let moved = self.demand_cache.len().saturating_sub(index);
        self.scalar
            .charge_work(u64::try_from(moved).unwrap_or(u64::MAX))?;
        self.demand_cache.insert(
            index,
            DemandCacheEntry {
                node: node.arena_index(),
                value: cache_value,
            },
        );
        Ok(())
    }

    fn condition_cache_position(
        &mut self,
        node: super::Node<'expr>,
        demand: Shape,
        index: usize,
    ) -> EvaluationResult<(usize, bool)> {
        let key = (node.arena_index(), demand.rows(), demand.columns(), index);
        let mut first = 0;
        let mut last = self.condition_cache.len();
        while first < last {
            let middle = first + (last - first) / 2;
            self.scalar.charge_work(1)?;
            let entry = &self.condition_cache[middle];
            let entry_key = (
                entry.node,
                entry.demand.rows(),
                entry.demand.columns(),
                entry.index,
            );
            match entry_key.cmp(&key) {
                std::cmp::Ordering::Less => first = middle + 1,
                std::cmp::Ordering::Equal => return Ok((middle, true)),
                std::cmp::Ordering::Greater => last = middle,
            }
        }
        Ok((first, false))
    }

    fn condition_cache_shape_position(&mut self, node: usize) -> EvaluationResult<(usize, bool)> {
        let mut first = 0;
        let mut last = self.condition_cache_shapes.len();
        while first < last {
            let middle = first + (last - first) / 2;
            self.scalar.charge_work(1)?;
            match self.condition_cache_shapes[middle].node.cmp(&node) {
                std::cmp::Ordering::Less => first = middle + 1,
                std::cmp::Ordering::Equal => return Ok((middle, true)),
                std::cmp::Ordering::Greater => last = middle,
            }
        }
        Ok((first, false))
    }

    /// Record the source shape used for a condition's cached coordinates.
    ///
    /// A lazy branch can be replanned at a larger shape after a sibling has
    /// widened the result.  Preserve entries whose coordinates remain valid
    /// under that broadcast expansion, and discard entries only when an axis
    /// shrinks or becomes incompatible.  The retained vector reservation
    /// remains owned by the cache and therefore stays charged.
    fn condition_cache_set_shape(
        &mut self,
        node: super::Node<'expr>,
        shape: Shape,
    ) -> EvaluationResult<()> {
        let node = node.arena_index();
        let (position, present) = self.condition_cache_shape_position(node)?;
        if present {
            if self.condition_cache_shapes[position].shape == shape {
                return Ok(());
            }
            let old_shape = self.condition_cache_shapes[position].shape;
            // `retain` examines every entry and may shift every surviving
            // entry.  Charge that full bounded pass before mutating the
            // cache; charging only the removed subset would make shape
            // growth an unaccounted quadratic scan.
            let scanned = self.condition_cache.len();
            self.scalar
                .charge_work(u64::try_from(scanned).unwrap_or(u64::MAX))?;
            let preserve =
                old_shape.rows() <= shape.rows() && old_shape.columns() <= shape.columns();
            self.condition_cache.retain_mut(|entry| {
                if entry.node != node {
                    return true;
                }
                if !preserve {
                    return false;
                }
                // Treat the old source shape as the destination and the new
                // canonical shape as the source.  This maps each previously
                // evaluated coordinate into the widened shape while keeping
                // singleton axes broadcastable.
                let Some(index) = projected_array_index(shape, old_shape, entry.index) else {
                    return false;
                };
                entry.demand = shape;
                entry.index = index;
                true
            });
            self.condition_cache_shapes[position].shape = shape;
            return Ok(());
        }
        ensure_capacity(
            &mut self.condition_cache_shapes,
            &mut self.condition_cache_shape_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value condition cache shapes",
        )?;
        let moved = self.condition_cache_shapes.len().saturating_sub(position);
        self.scalar
            .charge_work(u64::try_from(moved).unwrap_or(u64::MAX))?;
        self.condition_cache_shapes
            .insert(position, ConditionCacheShape { node, shape });
        Ok(())
    }

    fn condition_cache_get(
        &mut self,
        node: super::Node<'expr>,
        demand: Shape,
        index: usize,
    ) -> EvaluationResult<ConditionCacheLookup<'expr>> {
        let (shape_position, shape_present) =
            self.condition_cache_shape_position(node.arena_index())?;
        let key = if shape_present {
            let source_shape = self.condition_cache_shapes[shape_position].shape;
            let Some(source_index) = projected_array_index(source_shape, demand, index) else {
                // Do not fall back to the consumer coordinate here.  A raw
                // entry under that coordinate could be an alias created by
                // a widened demand, and an out-of-shape value is not a
                // source coordinate that can be reused safely.
                return Ok(ConditionCacheLookup {
                    value: None,
                    key: None,
                });
            };
            (source_shape, source_index)
        } else {
            (demand, index)
        };
        let (position, present) = self.condition_cache_position(node, key.0, key.1)?;
        if !present {
            return Ok(ConditionCacheLookup {
                value: None,
                key: Some(key),
            });
        }
        let value = self.condition_cache[position].value;
        Ok(ConditionCacheLookup {
            value: Some(match value {
                ConditionCacheValue::Empty => RuntimeValue::Empty,
                ConditionCacheValue::Missing => RuntimeValue::Missing,
                ConditionCacheValue::Number(value) => {
                    RuntimeValue::Scalar(WorkingValue::Number(value))
                },
                ConditionCacheValue::Logical(value) => {
                    RuntimeValue::Scalar(WorkingValue::Logical(value))
                },
                ConditionCacheValue::Text(value) => {
                    RuntimeValue::Scalar(WorkingValue::Text(TextValue::borrowed(value)))
                },
                ConditionCacheValue::Error(error) => {
                    RuntimeValue::Scalar(WorkingValue::Error(error))
                },
            }),
            key: Some(key),
        })
    }

    fn condition_cache_put(
        &mut self,
        node: super::Node<'expr>,
        demand: Shape,
        index: usize,
        value: &RuntimeValue<'expr>,
    ) -> EvaluationResult<()> {
        let cache_value = match value {
            RuntimeValue::Empty => ConditionCacheValue::Empty,
            RuntimeValue::Missing => ConditionCacheValue::Missing,
            RuntimeValue::Scalar(WorkingValue::Number(value)) => {
                ConditionCacheValue::Number(*value)
            },
            RuntimeValue::Scalar(WorkingValue::Logical(value)) => {
                ConditionCacheValue::Logical(*value)
            },
            RuntimeValue::Scalar(WorkingValue::Text(value)) => {
                let Cow::Borrowed(text) = &value.text else {
                    return Ok(());
                };
                ConditionCacheValue::Text(text)
            },
            RuntimeValue::Scalar(WorkingValue::Error(error)) => ConditionCacheValue::Error(*error),
            RuntimeValue::Scalar(WorkingValue::Complex(_)) => return Ok(()),
            RuntimeValue::ScalarCell(_) | RuntimeValue::Array(_) | RuntimeValue::Areas(_) => {
                return Ok(());
            },
        };
        self.condition_cache_set_shape(node, demand)?;
        let (position, present) = self.condition_cache_position(node, demand, index)?;
        if present {
            self.condition_cache[position].value = cache_value;
            return Ok(());
        }
        ensure_capacity(
            &mut self.condition_cache,
            &mut self.condition_cache_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value condition cache",
        )?;
        let moved = self.condition_cache.len().saturating_sub(position);
        self.scalar
            .charge_work(u64::try_from(moved).unwrap_or(u64::MAX))?;
        self.condition_cache.insert(
            position,
            ConditionCacheEntry {
                node: node.arena_index(),
                demand,
                index,
                value: cache_value,
            },
        );
        Ok(())
    }

    fn matrix_step(&mut self) -> EvaluationResult<()> {
        let mut matrix = self
            .matrix
            .take()
            .ok_or(EvaluationFailure::InvalidExpression(
                "matrix continuation is missing",
            ))?;
        if matrix.pending {
            return Err(EvaluationFailure::InvalidExpression(
                "matrix continuation has a pending branch",
            ));
        }
        let cells = matrix
            .shape
            .cell_count()
            .ok_or(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::MAX,
                self.limits.max_array_cells,
            )))?;
        if matrix.index >= cells {
            self.mode = matrix.restore_mode;
            self.position = matrix.restore_position;
            self.projection = matrix.restore_projection;
            return self.push_value(RuntimeValue::Array(RuntimeArrayValue {
                shape: matrix.shape,
                cells: matrix.output,
                _reservation: matrix.output_reservation,
                origin: None,
                preserve_scalar_result: false,
            }));
        }

        let index = matrix.index;
        self.charge_cell_work(index)?;
        let mut cache_slot = None;
        let branch =
            match matrix.kind {
                MatrixKind::If {
                    then_branch,
                    else_branch,
                } => {
                    let condition_element = {
                        let condition = matrix.condition.as_ref().ok_or(
                            EvaluationFailure::InvalidExpression("matrix IF condition is missing"),
                        )?;
                        array_element_for(condition, matrix.shape, index)
                            .map(|element| self.clone_element(element))
                            .transpose()?
                            .unwrap_or(RuntimeElement::Missing)
                    };
                    match self.scalar_logical(element_to_runtime(condition_element))? {
                        Ok(true) => {
                            if matrix.then_cacheable {
                                cache_slot = Some(MatrixCacheSlot::Then);
                            }
                            then_branch
                        },
                        Ok(false) if else_branch.is_some() => {
                            if matrix.else_cacheable {
                                cache_slot = Some(MatrixCacheSlot::Else);
                            }
                            else_branch
                        },
                        Ok(false) => {
                            matrix
                                .output
                                .push(RuntimeElement::Present(WorkingValue::Logical(false)));
                            matrix.index += 1;
                            self.matrix = Some(matrix);
                            return self.push_frame(ValueFrame::MatrixStep);
                        },
                        Err(error) => {
                            matrix
                                .output
                                .push(RuntimeElement::Present(WorkingValue::Error(error)));
                            matrix.index += 1;
                            self.matrix = Some(matrix);
                            return self.push_frame(ValueFrame::MatrixStep);
                        },
                    }
                },
                MatrixKind::IfError {
                    alternative,
                    catches_not_available,
                } => {
                    let element =
                        {
                            let value = matrix.value.as_ref().ok_or(
                                EvaluationFailure::InvalidExpression(
                                    "matrix IFERROR value is missing",
                                ),
                            )?;
                            self.select_matrix_value(value, matrix.shape, index)?
                        };
                    let catches = match element {
                        RuntimeElement::Present(WorkingValue::Error(error)) => {
                            !catches_not_available || error == ScalarError::NotAvailable
                        },
                        RuntimeElement::Empty
                        | RuntimeElement::Missing
                        | RuntimeElement::Present(WorkingValue::Number(_))
                        | RuntimeElement::Present(WorkingValue::Logical(_))
                        | RuntimeElement::Present(WorkingValue::Text(_))
                        | RuntimeElement::Present(WorkingValue::Complex(_)) => false,
                    };
                    if !catches {
                        matrix.output.push(element);
                        matrix.index += 1;
                        self.matrix = Some(matrix);
                        return self.push_frame(ValueFrame::MatrixStep);
                    }
                    if matrix.alternative_cacheable {
                        cache_slot = Some(MatrixCacheSlot::Alternative);
                    }
                    Some(alternative)
                },
            };

        let Some(branch) = branch else {
            matrix.output.push(RuntimeElement::Empty);
            matrix.index += 1;
            self.matrix = Some(matrix);
            return self.push_frame(ValueFrame::MatrixStep);
        };
        if let Some(slot) = cache_slot {
            let cached = match slot {
                MatrixCacheSlot::Then => matrix.then_cache.as_ref(),
                MatrixCacheSlot::Else => matrix.else_cache.as_ref(),
                MatrixCacheSlot::Alternative => matrix.alternative_cache.as_ref(),
            };
            if let Some(cached) = cached {
                let element = self.select_matrix_value(cached, matrix.shape, index)?;
                matrix.output.push(element);
                matrix.index += 1;
                self.matrix = Some(matrix);
                return self.push_frame(ValueFrame::MatrixStep);
            }
        }
        if branch.is_missing() {
            matrix.output.push(match matrix.kind {
                MatrixKind::IfError { .. } => RuntimeElement::Missing,
                MatrixKind::If { .. } => RuntimeElement::Empty,
            });
            matrix.index += 1;
            self.matrix = Some(matrix);
            return self.push_frame(ValueFrame::MatrixStep);
        }

        let position = self.offset_position(index, matrix.shape)?;
        matrix.pending = true;
        self.mode = Mode::Scalar;
        self.position = position;
        self.projection = Some((matrix.shape, index));
        matrix.pending_cache = cache_slot;
        self.matrix = Some(matrix);
        self.push_frame(ValueFrame::MatrixCollect)?;
        self.push_frame(ValueFrame::VisitArgument(branch))
    }

    fn matrix_collect(&mut self) -> EvaluationResult<()> {
        let mut matrix = self
            .matrix
            .take()
            .ok_or(EvaluationFailure::InvalidExpression(
                "matrix collection continuation is missing",
            ))?;
        if !matrix.pending {
            return Err(EvaluationFailure::InvalidExpression(
                "matrix collection has no pending branch",
            ));
        }
        let value = self.pop_value()?;
        let element = self.select_matrix_value(&value, matrix.shape, matrix.index)?;
        if let Some(slot) = matrix.pending_cache.take() {
            match slot {
                MatrixCacheSlot::Then => matrix.then_cache = Some(value),
                MatrixCacheSlot::Else => matrix.else_cache = Some(value),
                MatrixCacheSlot::Alternative => matrix.alternative_cache = Some(value),
            }
        }
        matrix.output.push(element);
        matrix.index += 1;
        matrix.pending = false;
        self.mode = matrix.restore_mode;
        self.position = matrix.restore_position;
        self.projection = matrix.restore_projection;
        self.matrix = Some(matrix);
        self.push_frame(ValueFrame::MatrixStep)
    }

    fn shape_hint_demand(
        &mut self,
        root: super::Node<'expr>,
        demand: Shape,
        root_mask: Option<&ShapeMask>,
    ) -> EvaluationResult<Option<Shape>> {
        self.shape_frames.clear();
        self.shape_values.clear();
        // Clearing drops per-mask storage but retains the outer allocation.
        // Keep its reservation until that capacity is released with the VM.
        self.shape_masks.clear();
        let mask = root_mask
            .map(|mask| self.copy_shape_mask(mask))
            .transpose()?
            .map(|mask| self.install_shape_mask(mask))
            .transpose()?
            .flatten();
        self.push_shape_frame(ShapeFrame::Enter {
            node: root,
            demand: ShapeDemand {
                shape: demand,
                mask,
            },
        })?;
        while let Some(frame) = self.shape_frames.pop() {
            self.scalar.step()?;
            match frame {
                ShapeFrame::Enter { node, demand } => {
                    if Self::is_lazy_handler(node) {
                        let Some(condition) = self.lazy_condition(node)? else {
                            // Invalid arity is reported by the ordinary value
                            // VM.  It has no branch whose shape can contribute.
                            self.push_shape_value(Some(Shape::new(1, 1)?))?;
                            continue;
                        };
                        // Plan the condition with the same continuation stack
                        // used for every other node.  The LazyCondition frame
                        // resumes only after that complete traversal has
                        // produced its shape value.
                        self.push_shape_frame(ShapeFrame::LazyCondition {
                            node,
                            condition,
                            demand,
                        })?;
                        self.push_shape_frame(ShapeFrame::Enter {
                            node: condition,
                            demand,
                        })?;
                        continue;
                    }
                    if let super::Kind::Function { name } = node.kind() {
                        let order_function = super::order::OrderFunction::from_name(name);
                        let paired_function = super::paired::PairedFunction::from_name(name);
                        if order_function.is_some() || paired_function.is_some() {
                            let data_argument = |index| {
                                order_function.is_some_and(|function| function.data_argument(index))
                                    || paired_function
                                        .is_some_and(|function| function.data_argument(index))
                            };
                            let count = node.child_count();
                            self.scalar
                                .charge_work(u64::try_from(count).unwrap_or(u64::MAX))?;
                            let children =
                                (0..count).filter(|index| !data_argument(*index)).count();
                            self.push_shape_frame(ShapeFrame::Exit {
                                node,
                                children,
                                base: Some(Shape::new(1, 1)?),
                            })?;
                            // Only scalar parameters contribute output shape.
                            // Walk their full expressions using the ordinary
                            // planner, including references and matrix/lazy
                            // functions; data sequences remain whole inputs.
                            for index in (0..count).rev() {
                                if data_argument(index) {
                                    continue;
                                }
                                let child = node.child(index).ok_or(
                                    EvaluationFailure::InvalidExpression(
                                        "reducer shape parameter is missing",
                                    ),
                                )?;
                                self.push_shape_frame(ShapeFrame::Enter {
                                    node: child,
                                    demand,
                                })?;
                            }
                            continue;
                        }
                        if complex::is_complex_sequence_function(name)
                            || aggregate::is_aggregate_function(name)
                            || statistical::is_statistical_function(name)
                            || descriptive::is_descriptive_function(name)
                            || conditional::is_conditional_function(name)
                            || database::is_database_function(name)
                        {
                            // Database functions, IMSUM and IMPRODUCT reduce their complete
                            // sequence arguments to one scalar value. Do not
                            // let an array/reference child widen the result
                            // shape during a projected lazy-branch probe.
                            self.push_shape_value(Some(Shape::new(1, 1)?))?;
                            continue;
                        }
                        if let Some(function) = Self::matrix_function(name) {
                            match function {
                                MatrixFunction::Determinant => {
                                    self.push_shape_value(Some(Shape::new(1, 1)?))?;
                                },
                                MatrixFunction::Unit => {
                                    let shape = self.matrix_unit_shape_hint(node)?;
                                    self.push_shape_value(shape)?;
                                },
                                MatrixFunction::Inverse
                                | MatrixFunction::Multiply
                                | MatrixFunction::Transpose => {
                                    let children = node.child_count();
                                    self.push_shape_frame(ShapeFrame::MatrixExit {
                                        function,
                                        children,
                                    })?;
                                    for index in (0..children).rev() {
                                        let child = node.child(index).ok_or(
                                            EvaluationFailure::InvalidExpression(
                                                "matrix shape argument is missing",
                                            ),
                                        )?;
                                        self.push_shape_frame(ShapeFrame::Enter {
                                            node: child,
                                            demand,
                                        })?;
                                    }
                                },
                            }
                            continue;
                        }
                    }
                    if let super::Kind::Infix(operator) = node.kind() {
                        if matches!(
                            operator,
                            super::InfixOperator::Range
                                | super::InfixOperator::Intersection
                                | super::InfixOperator::Union
                        ) {
                            self.push_shape_frame(ShapeFrame::Reference { node, operator })?;
                            continue;
                        }
                    }
                    let direct = match node.kind() {
                        super::Kind::Array(dimensions) => dimensions
                            .columns()
                            .and_then(|columns| Shape::new(dimensions.rows(), columns).ok()),
                        super::Kind::Reference(reference) => {
                            self.reference_shape_hint(reference)?
                        },
                        super::Kind::Number
                        | super::Kind::String
                        | super::Kind::Error
                        | super::Kind::Missing
                        | super::Kind::NamedExpression { .. }
                        | super::Kind::QuotedLabel
                        | super::Kind::AutomaticIntersection => Some(Shape::new(1, 1)?),
                        _ => None,
                    };
                    let children = node.child_count();
                    if let Some(shape) = direct {
                        self.push_shape_value(Some(shape))?;
                    } else if children == 0 {
                        self.push_shape_value(Some(Shape::new(1, 1)?))?;
                    } else {
                        self.push_shape_frame(ShapeFrame::Exit {
                            node,
                            children,
                            base: None,
                        })?;
                        for index in (0..children).rev() {
                            let child =
                                node.child(index)
                                    .ok_or(EvaluationFailure::InvalidExpression(
                                        "shape planner child is missing",
                                    ))?;
                            self.push_shape_frame(ShapeFrame::Enter {
                                node: child,
                                demand,
                            })?;
                        }
                    }
                },
                ShapeFrame::LazyCondition {
                    node,
                    condition,
                    demand,
                } => {
                    let condition_shape = self.pop_shape_value()?.unwrap_or(demand.shape);
                    let Some(children) =
                        self.deferred_lazy_children(node, demand, condition, condition_shape)?
                    else {
                        self.push_shape_value(Some(Shape::new(1, 1)?))?;
                        continue;
                    };
                    self.push_shape_children(node, demand, children)?;
                },
                ShapeFrame::Reference { node, operator } => {
                    let shape = self.reference_operator_shape(node, operator)?;
                    self.push_shape_value(shape)?;
                },
                ShapeFrame::MatrixExit { function, children } => {
                    let shape = self.matrix_shape_from_children(function, children)?;
                    self.push_shape_value(shape)?;
                },
                ShapeFrame::Exit {
                    node,
                    children,
                    base,
                } => {
                    let shape = self.combine_planned_children(node, children, base)?;
                    self.push_shape_value(shape)?;
                },
            }
        }
        if self.shape_values.len() != 1 {
            return Err(EvaluationFailure::InvalidExpression(
                "shape planner did not produce one result",
            ));
        }
        let result = self.shape_values.pop().flatten();
        self.shape_frames.clear();
        self.shape_values.clear();
        // Clearing drops per-mask storage but retains the outer allocation.
        // Keep its reservation until that capacity is released with the VM.
        self.shape_masks.clear();
        Ok(result)
    }

    fn lazy_condition(
        &self,
        node: super::Node<'expr>,
    ) -> EvaluationResult<Option<super::Node<'expr>>> {
        let super::Kind::Function { name } = node.kind() else {
            return Ok(None);
        };
        let count = node.child_count();
        let valid = if name.eq_ignore_ascii_case("IF") {
            (1..=3).contains(&count)
        } else if name.eq_ignore_ascii_case("IFERROR") || name.eq_ignore_ascii_case("IFNA") {
            count == 2
        } else {
            false
        };
        if !valid {
            return Ok(None);
        }
        node.child(0)
            .map(Some)
            .ok_or(EvaluationFailure::InvalidExpression(
                "lazy handler condition is missing",
            ))
    }

    fn matrix_unit_shape_hint(
        &mut self,
        node: super::Node<'expr>,
    ) -> EvaluationResult<Option<Shape>> {
        if node.child_count() != 1 {
            return Ok(Some(Shape::new(1, 1)?));
        }
        let argument = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
            "MUNIT shape argument is missing",
        ))?;
        // The scalar parameter consumes one value.  Probe that value through
        // the ordinary VM continuation so computed expressions, nested
        // parentheses, arrays, and provider-backed cells all obey the same
        // first-element and coercion rules as final evaluation.  The probe
        // uses the bounded, isolated probe path without constructing a
        // second evaluator.
        let value = match self.evaluate_matrix_value(argument) {
            Ok(value) => self.matrix_scalar_parameter(value)?,
            Err(EvaluationFailure::Unsupported(_)) => return Ok(None),
            Err(error) => return Err(error),
        };
        let RuntimeValue::Scalar(WorkingValue::Number(value)) = value else {
            return Ok(Some(Shape::new(1, 1)?));
        };
        let value = value.trunc();
        if !value.is_finite() || value <= 0.0 || value >= usize::MAX as f64 {
            return Ok(Some(Shape::new(1, 1)?));
        }
        let size = value as usize;
        Ok(Shape::new(size, size).ok())
    }

    fn matrix_shape_from_children(
        &mut self,
        function: MatrixFunction,
        children: usize,
    ) -> EvaluationResult<Option<Shape>> {
        // Shapes are pushed in source order and popped in reverse order.
        let mut first = None;
        let mut second = None;
        for index in 0..children {
            let shape = self.pop_shape_value()?;
            match index {
                0 if children == 1 => first = shape,
                0 => second = shape,
                1 => first = shape,
                _ => {},
            }
        }
        match function {
            MatrixFunction::Inverse => {
                if children != 1 {
                    return Ok(Some(Shape::new(1, 1)?));
                }
                Ok(match first {
                    Some(shape) if shape.rows() == shape.columns() => Some(shape),
                    Some(_) => Some(Shape::new(1, 1)?),
                    None => None,
                })
            },
            MatrixFunction::Multiply => {
                if children != 2 {
                    return Ok(Some(Shape::new(1, 1)?));
                }
                Ok(match (first, second) {
                    (Some(left), Some(right)) if left.columns() == right.rows() => {
                        Some(Shape::new(left.rows(), right.columns())?)
                    },
                    (Some(_), Some(_)) => Some(Shape::new(1, 1)?),
                    _ => None,
                })
            },
            MatrixFunction::Transpose => {
                if children != 1 {
                    return Ok(Some(Shape::new(1, 1)?));
                }
                first
                    .map(|shape| Shape::new(shape.columns(), shape.rows()))
                    .transpose()
            },
            MatrixFunction::Determinant | MatrixFunction::Unit => Ok(Some(Shape::new(1, 1)?)),
        }
    }

    fn push_shape_children(
        &mut self,
        node: super::Node<'expr>,
        demand: ShapeDemand,
        children: LazyShapeChildren<'expr>,
    ) -> EvaluationResult<()> {
        self.push_shape_frame(ShapeFrame::Exit {
            node,
            children: children.len,
            base: Some(children.shape),
        })?;
        if let Some(child) = children.second {
            self.push_shape_frame(ShapeFrame::Enter {
                node: child,
                demand: ShapeDemand {
                    shape: children.shape,
                    mask: children.second_mask.or(demand.mask),
                },
            })?;
        }
        if let Some(child) = children.first {
            self.push_shape_frame(ShapeFrame::Enter {
                node: child,
                demand: ShapeDemand {
                    shape: children.shape,
                    mask: children.first_mask.or(demand.mask),
                },
            })?;
        }
        Ok(())
    }

    /// Resolve geometry for a reference expression in one bounded pass.  The
    /// small frame machine accepts parenthesized references, evaluated lazy
    /// functions, and arbitrary nesting of `:`, `!`, and `~`; the runtime
    /// reference module remains the authority for their area/list semantics.
    /// Lazy conditions are scalar-probed when their branch is needed, while
    /// unselected branches are never visited.  A list or a multi-plane result
    /// has no unambiguous broadcast shape and is deliberately left unknown for
    /// the normal reference path.
    fn reference_shape_value(
        &mut self,
        root: super::Node<'expr>,
    ) -> EvaluationResult<Option<RuntimeValue<'expr>>> {
        let mut scratch = ReferenceShapeScratch::new();
        ensure_capacity(
            &mut scratch.frames,
            &mut scratch.frame_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value reference shape frames",
        )?;
        scratch.frames.push(ReferenceShapeFrame::Visit(root));

        while let Some(frame) = scratch.frames.pop() {
            self.scalar.step()?;
            match frame {
                ReferenceShapeFrame::Visit(node) => match node.kind() {
                    super::Kind::Parenthesized => {
                        let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                            "reference shape parentheses are empty",
                        ))?;
                        ensure_capacity(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            1,
                            self.limits.scalar.max_stack_entries,
                            self.execution,
                            &self.storage_budget,
                            "formula value reference shape frames",
                        )?;
                        scratch.frames.push(ReferenceShapeFrame::Visit(child));
                    },
                    super::Kind::Reference(reference) => {
                        let value = match self.reference_value(reference) {
                            Ok(value) => value,
                            Err(EvaluationFailure::Unsupported(
                                super::UnsupportedKind::Reference
                                | super::UnsupportedKind::ReferenceOperator,
                            )) => return Ok(None),
                            Err(error) => return Err(error),
                        };
                        ensure_capacity(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            1,
                            self.limits.scalar.max_stack_entries,
                            self.execution,
                            &self.storage_budget,
                            "formula value reference shape values",
                        )?;
                        scratch.values.push(value);
                    },
                    super::Kind::Function { name } if Self::is_reference_value_handler(name) => {
                        self.schedule_reference_handler(
                            node,
                            name,
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                        )?;
                    },
                    super::Kind::Infix(operator)
                        if matches!(
                            operator,
                            super::InfixOperator::Range
                                | super::InfixOperator::Intersection
                                | super::InfixOperator::Union
                        ) =>
                    {
                        let left = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                            "reference shape left operand is missing",
                        ))?;
                        let right = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                            "reference shape right operand is missing",
                        ))?;
                        ensure_capacity(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            3,
                            self.limits.scalar.max_stack_entries,
                            self.execution,
                            &self.storage_budget,
                            "formula value reference shape frames",
                        )?;
                        scratch.frames.push(ReferenceShapeFrame::Apply(operator));
                        scratch.frames.push(ReferenceShapeFrame::Visit(right));
                        scratch.frames.push(ReferenceShapeFrame::Visit(left));
                    },
                    _ => return Ok(None),
                },
                ReferenceShapeFrame::Apply(operator) => {
                    let right =
                        scratch
                            .values
                            .pop()
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "reference shape right value is missing",
                            ))?;
                    let left = scratch
                        .values
                        .pop()
                        .ok_or(EvaluationFailure::InvalidExpression(
                            "reference shape left value is missing",
                        ))?;
                    let value = match operator {
                        super::InfixOperator::Range => self.combine_range(left, right),
                        super::InfixOperator::Intersection => self.intersect_ranges(left, right),
                        super::InfixOperator::Union => self.union_ranges(left, right),
                        _ => {
                            return Err(EvaluationFailure::InvalidExpression(
                                "non-reference operator reached shape evaluator",
                            ));
                        },
                    };
                    let value = match value {
                        Ok(value) => value,
                        Err(EvaluationFailure::Unsupported(
                            super::UnsupportedKind::Reference
                            | super::UnsupportedKind::ReferenceOperator,
                        )) => return Ok(None),
                        Err(error) => return Err(error),
                    };
                    ensure_capacity(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        1,
                        self.limits.scalar.max_stack_entries,
                        self.execution,
                        &self.storage_budget,
                        "formula value reference shape values",
                    )?;
                    scratch.values.push(value);
                },
                ReferenceShapeFrame::FinishError {
                    alternative,
                    catches_not_available,
                } => {
                    let value =
                        scratch
                            .values
                            .pop()
                            .ok_or(EvaluationFailure::InvalidExpression(
                                "reference error-handler value is missing",
                            ))?;
                    // A reference list is a reference-sequence value, but it is
                    // not a scalar error-handler operand.  The scalar profile
                    // turns it into #VALUE! here so IFERROR can catch it while
                    // IFNA correctly leaves the error uncaught.
                    let value = match value {
                        RuntimeValue::Areas(areas) if areas.is_list => {
                            RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Value))
                        },
                        value => value,
                    };
                    let caught = matches!(
                        &value,
                        RuntimeValue::Scalar(WorkingValue::Error(error))
                            if !catches_not_available || *error == ScalarError::NotAvailable
                    );
                    if caught {
                        self.push_reference_value(
                            alternative,
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                        )?;
                    } else {
                        ensure_capacity(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            1,
                            self.limits.scalar.max_stack_entries,
                            self.execution,
                            &self.storage_budget,
                            "formula value reference shape values",
                        )?;
                        scratch.values.push(value);
                    }
                },
            }
        }
        if scratch.values.len() != 1 {
            return Err(EvaluationFailure::InvalidExpression(
                "reference shape evaluator did not produce one value",
            ));
        }
        Ok(scratch.values.pop())
    }

    fn is_reference_value_handler(name: &str) -> bool {
        name.eq_ignore_ascii_case("IF")
            || name.eq_ignore_ascii_case("IFERROR")
            || name.eq_ignore_ascii_case("IFNA")
    }

    fn reference_candidate(&mut self, mut node: super::Node<'expr>) -> EvaluationResult<bool> {
        loop {
            match node.kind() {
                super::Kind::Parenthesized => {
                    self.scalar.charge_work(1)?;
                    node = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                        "reference candidate parentheses are empty",
                    ))?;
                },
                super::Kind::Reference(_) => return Ok(true),
                super::Kind::Function { name } if Self::is_reference_value_handler(name) => {
                    return Ok(true);
                },
                super::Kind::Infix(
                    super::InfixOperator::Range
                    | super::InfixOperator::Intersection
                    | super::InfixOperator::Union,
                ) => {
                    return Ok(true);
                },
                _ => return Ok(false),
            }
        }
    }

    /// Find a matrix function through a bounded parenthesized wrapper.  A
    /// reference-valued lazy handler needs the complete runtime result when
    /// its selected operand is a matrix function: a successful `MINVERSE`
    /// remains an Array and must be refused by a reference operator, while a
    /// singular inverse is a scalar formula error that `IFERROR` can catch.
    /// Keeping this check iterative also avoids introducing an AST-recursive
    /// probe into the reference planner.
    fn direct_matrix_function(
        &mut self,
        mut node: super::Node<'expr>,
    ) -> EvaluationResult<Option<MatrixFunction>> {
        loop {
            match node.kind() {
                super::Kind::Parenthesized => {
                    self.scalar.charge_work(1)?;
                    node = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                        "matrix function parentheses are empty",
                    ))?;
                },
                super::Kind::Function { name } => return Ok(Self::matrix_function(name)),
                _ => return Ok(None),
            }
        }
    }

    fn reference_operand_kind_from_runtime(value: &RuntimeValue<'expr>) -> ReferenceOperandKind {
        match value {
            RuntimeValue::Areas(areas) if areas.is_list => ReferenceOperandKind::ReferenceList,
            RuntimeValue::Areas(_) => ReferenceOperandKind::Reference,
            RuntimeValue::Array(_) => ReferenceOperandKind::Array,
            RuntimeValue::Scalar(WorkingValue::Error(error)) => ReferenceOperandKind::Error(*error),
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_) => ReferenceOperandKind::Scalar,
        }
    }

    fn reference_kind_value_kind(value: &ReferenceKindValue<'expr>) -> ReferenceOperandKind {
        match value {
            ReferenceKindValue::Known(kind) => *kind,
            ReferenceKindValue::Runtime(value) => Self::reference_operand_kind_from_runtime(value),
        }
    }

    fn reference_kind_value_into_kind(value: ReferenceKindValue<'expr>) -> ReferenceOperandKind {
        match value {
            ReferenceKindValue::Known(kind) => kind,
            ReferenceKindValue::Runtime(value) => Self::reference_operand_kind_from_runtime(&value),
        }
    }

    /// Classify one handler operand without materializing an array.  The
    /// explicit frame stack carries this through arbitrary operators and
    /// nested lazy handlers. `Reference` is kept distinct from `Array` so a
    /// selected reference branch can still feed the reference geometry
    /// evaluator; either kind is array-like when it is used as a handler's
    /// condition or protected value.
    fn reference_handler_operand_kind(
        &mut self,
        root: super::Node<'expr>,
    ) -> EvaluationResult<ReferenceOperandKind> {
        let mut scratch = ReferenceKindScratch::new();
        ensure_capacity(
            &mut scratch.frames,
            &mut scratch.frame_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value reference kind frames",
        )?;
        scratch.frames.push(ReferenceKindFrame::Visit(root));

        while let Some(frame) = scratch.frames.pop() {
            self.scalar.step()?;
            match frame {
                ReferenceKindFrame::Visit(node) => match node.kind() {
                    super::Kind::Parenthesized => {
                        let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                            "reference kind parentheses are empty",
                        ))?;
                        ensure_capacity(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            1,
                            self.limits.scalar.max_stack_entries,
                            self.execution,
                            &self.storage_budget,
                            "formula value reference kind frames",
                        )?;
                        scratch.frames.push(ReferenceKindFrame::Visit(child));
                    },
                    super::Kind::Number
                    | super::Kind::String
                    | super::Kind::Error
                    | super::Kind::Missing => self.push_reference_kind_value(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        ReferenceOperandKind::Scalar,
                    )?,
                    super::Kind::Array(_) | super::Kind::ArrayRow => self
                        .push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Array,
                        )?,
                    super::Kind::Reference(reference) => {
                        let value = match self.reference_value(reference) {
                            Ok(value) => value,
                            Err(EvaluationFailure::Unsupported(
                                super::UnsupportedKind::Reference,
                            )) => {
                                self.push_reference_kind_value(
                                    &mut scratch.values,
                                    &mut scratch.value_reservation,
                                    ReferenceOperandKind::Unknown,
                                )?;
                                continue;
                            },
                            Err(error) => return Err(error),
                        };
                        self.push_reference_kind_runtime(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            value,
                        )?;
                    },
                    super::Kind::Prefix(_) | super::Kind::Postfix(_) => {
                        let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                            "reference kind unary operand is missing",
                        ))?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Unary,
                        )?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Visit(child),
                        )?;
                    },
                    super::Kind::Infix(operator) => {
                        if matches!(
                            operator,
                            super::InfixOperator::Range
                                | super::InfixOperator::Intersection
                                | super::InfixOperator::Union
                        ) {
                            let left =
                                node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                                    "reference kind left operand is missing",
                                ))?;
                            let right =
                                node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                                    "reference kind right operand is missing",
                                ))?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Infix(operator),
                            )?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Visit(right),
                            )?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Visit(left),
                            )?;
                            continue;
                        }
                        let left = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                            "reference kind left operand is missing",
                        ))?;
                        let right = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                            "reference kind right operand is missing",
                        ))?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Infix(operator),
                        )?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Visit(right),
                        )?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Visit(left),
                        )?;
                    },
                    super::Kind::Function { name } => {
                        let count = node.child_count();
                        if name.eq_ignore_ascii_case("IF") {
                            if !(1..=3).contains(&count) {
                                self.push_reference_kind_value(
                                    &mut scratch.values,
                                    &mut scratch.value_reservation,
                                    ReferenceOperandKind::Scalar,
                                )?;
                                continue;
                            }
                            let condition =
                                node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                                    "reference kind IF condition is missing",
                                ))?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::IfAfterCondition { node, condition },
                            )?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Visit(condition),
                            )?;
                            continue;
                        }
                        if name.eq_ignore_ascii_case("IFERROR") || name.eq_ignore_ascii_case("IFNA")
                        {
                            if count != 2 {
                                self.push_reference_kind_value(
                                    &mut scratch.values,
                                    &mut scratch.value_reservation,
                                    ReferenceOperandKind::Scalar,
                                )?;
                                continue;
                            }
                            let value =
                                node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                                    "reference kind error-handler value is missing",
                                ))?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::IfErrorAfterValue { node },
                            )?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Visit(value),
                            )?;
                            continue;
                        }
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Function { node, name, count },
                        )?;
                        for index in (0..count).rev() {
                            let child =
                                node.child(index)
                                    .ok_or(EvaluationFailure::InvalidExpression(
                                        "reference kind function argument is missing",
                                    ))?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Visit(child),
                            )?;
                        }
                    },
                    super::Kind::NamedExpression { .. }
                    | super::Kind::QuotedLabel
                    | super::Kind::AutomaticIntersection => self.push_reference_kind_value(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        ReferenceOperandKind::Unknown,
                    )?,
                },
                ReferenceKindFrame::Unary => {
                    let child = self.pop_reference_kind_value(&mut scratch.values)?;
                    let kind = match Self::reference_kind_value_into_kind(child) {
                        ReferenceOperandKind::Array | ReferenceOperandKind::Reference => {
                            ReferenceOperandKind::Array
                        },
                        ReferenceOperandKind::ReferenceList => {
                            ReferenceOperandKind::Error(ScalarError::Value)
                        },
                        kind => kind,
                    };
                    self.push_reference_kind_value(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        kind,
                    )?;
                },
                ReferenceKindFrame::Infix(operator) => {
                    let right_value = self.pop_reference_kind_value(&mut scratch.values)?;
                    let left_value = self.pop_reference_kind_value(&mut scratch.values)?;
                    let right = Self::reference_kind_value_kind(&right_value);
                    let left = Self::reference_kind_value_kind(&left_value);
                    if matches!(
                        operator,
                        super::InfixOperator::Range
                            | super::InfixOperator::Intersection
                            | super::InfixOperator::Union
                    ) {
                        let value = match (left, right) {
                            (ReferenceOperandKind::Error(error), _)
                            | (_, ReferenceOperandKind::Error(error)) => {
                                ReferenceKindValue::Known(ReferenceOperandKind::Error(error))
                            },
                            (
                                ReferenceOperandKind::Reference
                                | ReferenceOperandKind::ReferenceList,
                                ReferenceOperandKind::Reference
                                | ReferenceOperandKind::ReferenceList,
                            ) => match (left_value, right_value) {
                                (
                                    ReferenceKindValue::Runtime(left),
                                    ReferenceKindValue::Runtime(right),
                                ) => {
                                    let result = match operator {
                                        super::InfixOperator::Range => {
                                            self.combine_range(left, right)
                                        },
                                        super::InfixOperator::Intersection => {
                                            self.intersect_ranges(left, right)
                                        },
                                        super::InfixOperator::Union => {
                                            self.union_ranges(left, right)
                                        },
                                        _ => unreachable!("checked reference operator"),
                                    };
                                    match result {
                                        Ok(value) => ReferenceKindValue::Runtime(value),
                                        Err(EvaluationFailure::Unsupported(
                                            super::UnsupportedKind::Reference
                                            | super::UnsupportedKind::ReferenceOperator,
                                        )) => {
                                            ReferenceKindValue::Known(ReferenceOperandKind::Unknown)
                                        },
                                        Err(error) => return Err(error),
                                    }
                                },
                                _ => ReferenceKindValue::Known(ReferenceOperandKind::Unknown),
                            },
                            _ => ReferenceKindValue::Known(ReferenceOperandKind::Unknown),
                        };
                        self.push_reference_kind_entry(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            value,
                        )?;
                        continue;
                    }
                    let kind = if matches!(
                        (left, right),
                        (ReferenceOperandKind::Array, _)
                            | (_, ReferenceOperandKind::Array)
                            | (ReferenceOperandKind::Reference, _)
                            | (_, ReferenceOperandKind::Reference)
                    ) {
                        // Matrix mapping converts reference lists to a scalar
                        // error, then broadcasts that error with any array or
                        // ordinary reference operand.  Preserve that result
                        // shape even when the list appears on either side.
                        ReferenceOperandKind::Array
                    } else if matches!(
                        (left, right),
                        (ReferenceOperandKind::ReferenceList, _,)
                            | (_, ReferenceOperandKind::ReferenceList,)
                    ) {
                        ReferenceOperandKind::Error(ScalarError::Value)
                    } else if matches!(
                        (left, right),
                        (ReferenceOperandKind::Unknown, _)
                            | (_, ReferenceOperandKind::Unknown)
                            | (ReferenceOperandKind::Error(_), _)
                            | (_, ReferenceOperandKind::Error(_))
                    ) {
                        ReferenceOperandKind::Unknown
                    } else {
                        ReferenceOperandKind::Scalar
                    };
                    self.push_reference_kind_value(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        kind,
                    )?;
                },
                ReferenceKindFrame::Function { node, name, count } => {
                    let mut kind = ReferenceOperandKind::Scalar;
                    let order_function = super::order::OrderFunction::from_name(name);
                    let paired_function = super::paired::PairedFunction::from_name(name);
                    for index in (0..count).rev() {
                        let child = self.pop_reference_kind_value(&mut scratch.values)?;
                        if order_function.is_some_and(|function| function.data_argument(index))
                            || paired_function.is_some_and(|function| function.data_argument(index))
                        {
                            continue;
                        }
                        if name.eq_ignore_ascii_case("AND")
                            || name.eq_ignore_ascii_case("OR")
                            || complex::is_complex_sequence_function(name)
                            || aggregate::is_aggregate_function(name)
                            || statistical::is_statistical_function(name)
                            || descriptive::is_descriptive_function(name)
                            || conditional::is_conditional_function(name)
                            || database::is_database_function(name)
                        {
                            continue;
                        }
                        kind = match (kind, Self::reference_kind_value_into_kind(child)) {
                            (ReferenceOperandKind::Array, _)
                            | (_, ReferenceOperandKind::Array)
                            | (ReferenceOperandKind::Reference, _)
                            | (_, ReferenceOperandKind::Reference) => ReferenceOperandKind::Array,
                            (ReferenceOperandKind::ReferenceList, _)
                            | (_, ReferenceOperandKind::ReferenceList) => {
                                ReferenceOperandKind::Error(ScalarError::Value)
                            },
                            (_, ReferenceOperandKind::Error(error)) => {
                                ReferenceOperandKind::Error(error)
                            },
                            (ReferenceOperandKind::Error(error), _) => {
                                ReferenceOperandKind::Error(error)
                            },
                            (ReferenceOperandKind::Unknown, _)
                            | (_, ReferenceOperandKind::Unknown) => ReferenceOperandKind::Unknown,
                            _ => ReferenceOperandKind::Scalar,
                        };
                    }
                    if let Some(function) = Self::matrix_function(name) {
                        kind = match function {
                            MatrixFunction::Determinant => ReferenceOperandKind::Scalar,
                            MatrixFunction::Inverse
                            | MatrixFunction::Multiply
                            | MatrixFunction::Unit
                            | MatrixFunction::Transpose => {
                                self.reference_kind_matrix_value(node, function)?
                            },
                        };
                    }
                    self.push_reference_kind_value(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        kind,
                    )?;
                },
                ReferenceKindFrame::IfAfterCondition { node, condition } => {
                    let condition_value = self.pop_reference_kind_value(&mut scratch.values)?;
                    let condition_kind = Self::reference_kind_value_kind(&condition_value);
                    if condition_kind == ReferenceOperandKind::ReferenceList {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Error(ScalarError::Value),
                        )?;
                        continue;
                    }
                    if matches!(
                        condition_kind,
                        ReferenceOperandKind::Array | ReferenceOperandKind::Reference
                    ) {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Array,
                        )?;
                        continue;
                    }
                    let condition = match condition_value {
                        ReferenceKindValue::Runtime(value) => self.scalar_logical(value)?,
                        ReferenceKindValue::Known(ReferenceOperandKind::Error(error)) => Err(error),
                        ReferenceKindValue::Known(_) => {
                            let value = self.evaluate_scalar_at(condition, Shape::new(1, 1)?, 0)?;
                            self.scalar_logical(value)?
                        },
                    };
                    let condition = match condition {
                        Ok(value) => value,
                        Err(error) => {
                            self.push_reference_kind_value(
                                &mut scratch.values,
                                &mut scratch.value_reservation,
                                ReferenceOperandKind::Error(error),
                            )?;
                            continue;
                        },
                    };
                    let count = node.child_count();
                    if count == 1 {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Scalar,
                        )?;
                        continue;
                    }
                    let branch = if condition {
                        node.child(1)
                    } else if count == 3 {
                        node.child(2)
                    } else {
                        None
                    };
                    let Some(branch) = branch else {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Scalar,
                        )?;
                        continue;
                    };
                    if branch.is_missing() {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Scalar,
                        )?;
                    } else {
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::UseChild,
                        )?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Visit(branch),
                        )?;
                    }
                },
                ReferenceKindFrame::IfErrorAfterValue { node } => {
                    let value_value = self.pop_reference_kind_value(&mut scratch.values)?;
                    let value_kind = Self::reference_kind_value_kind(&value_value);
                    if value_kind == ReferenceOperandKind::ReferenceList {
                        let alternative =
                            node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                                "reference kind error-handler alternative is missing",
                            ))?;
                        if alternative.is_missing() {
                            self.push_reference_kind_value(
                                &mut scratch.values,
                                &mut scratch.value_reservation,
                                ReferenceOperandKind::Scalar,
                            )?;
                        } else if node
                            .function_name()
                            .is_some_and(|name| name.eq_ignore_ascii_case("IFERROR"))
                        {
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::UseChild,
                            )?;
                            self.push_reference_kind_frame(
                                &mut scratch.frames,
                                &mut scratch.frame_reservation,
                                ReferenceKindFrame::Visit(alternative),
                            )?;
                        } else {
                            self.push_reference_kind_value(
                                &mut scratch.values,
                                &mut scratch.value_reservation,
                                ReferenceOperandKind::Error(ScalarError::Value),
                            )?;
                        }
                        continue;
                    }
                    if matches!(
                        value_kind,
                        ReferenceOperandKind::Array | ReferenceOperandKind::Reference
                    ) {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Array,
                        )?;
                        continue;
                    }
                    let value = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                        "reference kind error-handler value is missing",
                    ))?;
                    let alternative = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                        "reference kind error-handler alternative is missing",
                    ))?;
                    let catches_not_available = node
                        .function_name()
                        .is_some_and(|name| name.eq_ignore_ascii_case("IFNA"));
                    let mut uncaught_error = None;
                    let catches = match value_kind {
                        ReferenceOperandKind::Error(error) => {
                            uncaught_error = Some(error);
                            !catches_not_available || error == ScalarError::NotAvailable
                        },
                        _ => match value_value {
                            ReferenceKindValue::Runtime(value) => match value {
                                RuntimeValue::Scalar(WorkingValue::Error(error)) => {
                                    uncaught_error = Some(error);
                                    !catches_not_available || error == ScalarError::NotAvailable
                                },
                                _ => false,
                            },
                            ReferenceKindValue::Known(_) => {
                                let value = self.evaluate_scalar_at(value, Shape::new(1, 1)?, 0)?;
                                match value {
                                    RuntimeValue::Scalar(WorkingValue::Error(error)) => {
                                        uncaught_error = Some(error);
                                        !catches_not_available || error == ScalarError::NotAvailable
                                    },
                                    _ => false,
                                }
                            },
                        },
                    };
                    if alternative.is_missing() {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            ReferenceOperandKind::Scalar,
                        )?;
                    } else if catches {
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::UseChild,
                        )?;
                        self.push_reference_kind_frame(
                            &mut scratch.frames,
                            &mut scratch.frame_reservation,
                            ReferenceKindFrame::Visit(alternative),
                        )?;
                    } else {
                        self.push_reference_kind_value(
                            &mut scratch.values,
                            &mut scratch.value_reservation,
                            uncaught_error
                                .map_or(ReferenceOperandKind::Scalar, ReferenceOperandKind::Error),
                        )?;
                    }
                },
                ReferenceKindFrame::UseChild => {
                    let value = self.pop_reference_kind_value(&mut scratch.values)?;
                    self.push_reference_kind_entry(
                        &mut scratch.values,
                        &mut scratch.value_reservation,
                        value,
                    )?;
                },
            }
        }
        if scratch.values.len() != 1 {
            return Err(EvaluationFailure::InvalidExpression(
                "reference kind evaluator did not produce one value",
            ));
        }
        scratch
            .values
            .pop()
            .map(Self::reference_kind_value_into_kind)
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference kind evaluator value is missing",
            ))
    }

    fn push_reference_kind_frame(
        &mut self,
        frames: &mut Vec<ReferenceKindFrame<'expr>>,
        reservation: &mut Option<Reservation>,
        frame: ReferenceKindFrame<'expr>,
    ) -> EvaluationResult<()> {
        ensure_capacity(
            frames,
            reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value reference kind frames",
        )?;
        frames.push(frame);
        Ok(())
    }

    fn push_reference_kind_value(
        &mut self,
        values: &mut Vec<ReferenceKindValue<'expr>>,
        reservation: &mut Option<Reservation>,
        value: ReferenceOperandKind,
    ) -> EvaluationResult<()> {
        self.push_reference_kind_entry(values, reservation, ReferenceKindValue::Known(value))
    }

    fn push_reference_kind_runtime(
        &mut self,
        values: &mut Vec<ReferenceKindValue<'expr>>,
        reservation: &mut Option<Reservation>,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<()> {
        self.push_reference_kind_entry(values, reservation, ReferenceKindValue::Runtime(value))
    }

    fn push_reference_kind_entry(
        &mut self,
        values: &mut Vec<ReferenceKindValue<'expr>>,
        reservation: &mut Option<Reservation>,
        value: ReferenceKindValue<'expr>,
    ) -> EvaluationResult<()> {
        ensure_capacity(
            values,
            reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value reference kind values",
        )?;
        values.push(value);
        Ok(())
    }

    fn pop_reference_kind_value(
        &mut self,
        values: &mut Vec<ReferenceKindValue<'expr>>,
    ) -> EvaluationResult<ReferenceKindValue<'expr>> {
        values.pop().ok_or(EvaluationFailure::InvalidExpression(
            "reference kind value stack underflow",
        ))
    }

    fn push_reference_value(
        &mut self,
        node: super::Node<'expr>,
        frames: &mut Vec<ReferenceShapeFrame<'expr>>,
        frame_reservation: &mut Option<Reservation>,
        values: &mut Vec<RuntimeValue<'expr>>,
        value_reservation: &mut Option<Reservation>,
    ) -> EvaluationResult<()> {
        if node.is_missing() {
            return self.push_reference_runtime_value(
                RuntimeValue::Empty,
                values,
                value_reservation,
            );
        }
        // Resolve a selected matrix function once in array mode.  Its
        // runtime result, rather than its declared function family, decides
        // whether the reference operator can consume it.  This avoids
        // evaluating a singular inverse once for kind discovery and again for
        // scalar error handling.
        if self.direct_matrix_function(node)?.is_some() {
            let value = self.evaluate_matrix_value(node)?;
            if value.is_array_like() {
                return Err(EvaluationFailure::Unsupported(
                    super::UnsupportedKind::ReferenceOperator,
                ));
            }
            return self.push_reference_runtime_value(value, values, value_reservation);
        }
        if matches!(
            self.reference_handler_operand_kind(node)?,
            ReferenceOperandKind::Array
        ) {
            return Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::ReferenceOperator,
            ));
        }
        if self.reference_candidate(node)? {
            ensure_capacity(
                frames,
                frame_reservation,
                1,
                self.limits.scalar.max_stack_entries,
                self.execution,
                &self.storage_budget,
                "formula value reference shape frames",
            )?;
            frames.push(ReferenceShapeFrame::Visit(node));
            return Ok(());
        }
        let value = self.evaluate_scalar_at(node, Shape::new(1, 1)?, 0)?;
        self.push_reference_runtime_value(value, values, value_reservation)
    }

    fn push_reference_runtime_value(
        &mut self,
        value: RuntimeValue<'expr>,
        values: &mut Vec<RuntimeValue<'expr>>,
        value_reservation: &mut Option<Reservation>,
    ) -> EvaluationResult<()> {
        ensure_capacity(
            values,
            value_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value reference shape values",
        )?;
        values.push(value);
        Ok(())
    }

    fn schedule_reference_handler(
        &mut self,
        node: super::Node<'expr>,
        name: &str,
        frames: &mut Vec<ReferenceShapeFrame<'expr>>,
        frame_reservation: &mut Option<Reservation>,
        values: &mut Vec<RuntimeValue<'expr>>,
        value_reservation: &mut Option<Reservation>,
    ) -> EvaluationResult<()> {
        let count = node.child_count();
        if name.eq_ignore_ascii_case("IF") {
            if !(1..=3).contains(&count) {
                return self.push_reference_runtime_value(
                    RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Value)),
                    values,
                    value_reservation,
                );
            }
            let condition = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                "reference IF condition is missing",
            ))?;
            if matches!(
                self.reference_handler_operand_kind(condition)?,
                ReferenceOperandKind::Array | ReferenceOperandKind::Reference
            ) {
                return Err(EvaluationFailure::Unsupported(
                    super::UnsupportedKind::ReferenceOperator,
                ));
            }
            let condition = self.evaluate_scalar_at(condition, Shape::new(1, 1)?, 0)?;
            let condition = match self.scalar_logical(condition)? {
                Ok(value) => value,
                Err(error) => {
                    return self.push_reference_runtime_value(
                        RuntimeValue::Scalar(WorkingValue::Error(error)),
                        values,
                        value_reservation,
                    );
                },
            };
            if count == 1 {
                return self.push_reference_runtime_value(
                    RuntimeValue::Scalar(WorkingValue::Logical(condition)),
                    values,
                    value_reservation,
                );
            }
            let branch = if condition {
                node.child(1)
            } else if count == 3 {
                node.child(2)
            } else {
                return self.push_reference_runtime_value(
                    RuntimeValue::Scalar(WorkingValue::Logical(false)),
                    values,
                    value_reservation,
                );
            };
            let branch = branch.ok_or(EvaluationFailure::InvalidExpression(
                "reference IF branch is missing",
            ))?;
            return self.push_reference_value(
                branch,
                frames,
                frame_reservation,
                values,
                value_reservation,
            );
        }

        if count != 2 {
            return self.push_reference_runtime_value(
                RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Value)),
                values,
                value_reservation,
            );
        }
        let value = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
            "reference error-handler value is missing",
        ))?;
        let alternative = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
            "reference error-handler alternative is missing",
        ))?;
        // As in `push_reference_value`, retain the one complete runtime
        // result for a selected matrix operand.  A successful array is a
        // typed reference-operator refusal; a scalar formula error reaches
        // the ordinary IFERROR/IFNA catch decision without a second matrix
        // evaluation.
        if self.direct_matrix_function(value)?.is_some() {
            let value = self.evaluate_matrix_value(value)?;
            if value.is_array_like() {
                return Err(EvaluationFailure::Unsupported(
                    super::UnsupportedKind::ReferenceOperator,
                ));
            }
            let caught = matches!(
                &value,
                RuntimeValue::Scalar(WorkingValue::Error(error))
                    if !name.eq_ignore_ascii_case("IFNA")
                        || *error == ScalarError::NotAvailable
            );
            if caught {
                return self.push_reference_value(
                    alternative,
                    frames,
                    frame_reservation,
                    values,
                    value_reservation,
                );
            }
            return self.push_reference_runtime_value(value, values, value_reservation);
        }
        if matches!(
            self.reference_handler_operand_kind(value)?,
            ReferenceOperandKind::Array | ReferenceOperandKind::Reference
        ) {
            return Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::ReferenceOperator,
            ));
        }
        if self.reference_candidate(value)? {
            ensure_capacity(
                frames,
                frame_reservation,
                2,
                self.limits.scalar.max_stack_entries,
                self.execution,
                &self.storage_budget,
                "formula value reference shape frames",
            )?;
            frames.push(ReferenceShapeFrame::FinishError {
                alternative,
                catches_not_available: name.eq_ignore_ascii_case("IFNA"),
            });
            frames.push(ReferenceShapeFrame::Visit(value));
            Ok(())
        } else {
            let value = self.evaluate_scalar_at(value, Shape::new(1, 1)?, 0)?;
            let caught = matches!(
                &value,
                RuntimeValue::Scalar(WorkingValue::Error(error))
                    if !name.eq_ignore_ascii_case("IFNA")
                        || *error == ScalarError::NotAvailable
            );
            if caught {
                self.push_reference_value(
                    alternative,
                    frames,
                    frame_reservation,
                    values,
                    value_reservation,
                )
            } else {
                self.push_reference_runtime_value(value, values, value_reservation)
            }
        }
    }

    fn reference_operator_shape(
        &mut self,
        node: super::Node<'expr>,
        operator: super::InfixOperator,
    ) -> EvaluationResult<Option<Shape>> {
        if !matches!(
            operator,
            super::InfixOperator::Range
                | super::InfixOperator::Intersection
                | super::InfixOperator::Union
        ) {
            return Ok(None);
        }
        let Some(value) = self.reference_shape_value(node)? else {
            return Ok(None);
        };
        match &value {
            RuntimeValue::Areas(areas) if !areas.is_list && areas.areas.len() == 1 => {
                let area = areas
                    .areas
                    .first()
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference operator area disappeared",
                    ))?;
                Ok(Some(Shape::new(area.rect.rows(), area.rect.columns())?))
            },
            RuntimeValue::Areas(_) => Ok(None),
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_)
            | RuntimeValue::Array(_) => Ok(Some(Shape::new(1, 1)?)),
        }
    }

    fn combine_planned_children(
        &mut self,
        node: super::Node<'expr>,
        children: usize,
        base: Option<Shape>,
    ) -> EvaluationResult<Option<Shape>> {
        let mut shapes = [None, None];
        let mut combined = base;
        for (index, shape_slot) in shapes.iter_mut().enumerate() {
            if index >= children {
                break;
            }
            let shape = self.pop_shape_value()?;
            *shape_slot = shape;
            combined = match (combined, shape) {
                (Some(left), Some(right)) => broadcast_shape(left, right),
                (left @ Some(_), None) | (None, left @ Some(_)) => left,
                (None, None) => None,
            };
        }
        for _ in 2..children {
            let shape = self.pop_shape_value()?;
            combined = match (combined, shape) {
                (Some(left), Some(right)) => broadcast_shape(left, right),
                (left @ Some(_), None) | (None, left @ Some(_)) => left,
                (None, None) => None,
            };
        }
        match node.kind() {
            super::Kind::Parenthesized | super::Kind::Prefix(_) | super::Kind::Postfix(_) => {
                Ok(shapes[0])
            },
            super::Kind::Infix(operator) => {
                if matches!(
                    operator,
                    super::InfixOperator::Range
                        | super::InfixOperator::Intersection
                        | super::InfixOperator::Union
                ) {
                    // Reference infixes are consumed by ShapeFrame::Reference
                    // before their children are scheduled.  Re-entering the
                    // complete geometry evaluator here would make a
                    // left-associated reference chain quadratic if a future
                    // frame path reached this fallback.
                    return Ok(None);
                }
                Ok(match (shapes[0], shapes[1]) {
                    (Some(left), Some(right)) => broadcast_shape(left, right),
                    (shape, None) | (None, shape) => shape,
                })
            },
            super::Kind::Function { name }
                if name.eq_ignore_ascii_case("AND")
                    || name.eq_ignore_ascii_case("OR")
                    || aggregate::is_aggregate_function(name)
                    || statistical::is_statistical_function(name)
                    || descriptive::is_descriptive_function(name)
                    || name.eq_ignore_ascii_case("TRUE")
                    || name.eq_ignore_ascii_case("FALSE") =>
            {
                Ok(Some(Shape::new(1, 1)?))
            },
            _ => Ok(combined.or(Some(Shape::new(1, 1)?))),
        }
    }

    fn demand_len(&self, demand: ShapeDemand) -> EvaluationResult<usize> {
        match demand.mask {
            Some(mask) => self
                .shape_masks
                .get(mask)
                .filter(|mask| mask.shape == demand.shape)
                .map(|mask| mask.indexes.len())
                .ok_or(EvaluationFailure::InvalidExpression(
                    "shape demand mask is missing",
                )),
            None => demand
                .shape
                .cell_count()
                .ok_or(EvaluationFailure::InvalidExpression(
                    "shape demand count overflow",
                )),
        }
    }

    fn demand_index(&self, demand: ShapeDemand, ordinal: usize) -> EvaluationResult<usize> {
        match demand.mask {
            Some(mask) => self
                .shape_masks
                .get(mask)
                .filter(|mask| mask.shape == demand.shape)
                .and_then(|mask| mask.indexes.get(ordinal).copied())
                .ok_or(EvaluationFailure::InvalidExpression(
                    "shape demand index is missing",
                )),
            None => Ok(ordinal),
        }
    }

    fn new_shape_mask(&self, shape: Shape) -> EvaluationResult<ShapeMask> {
        if shape.cell_count().is_none() {
            return Err(EvaluationFailure::InvalidExpression(
                "shape mask count overflow",
            ));
        }
        Ok(ShapeMask {
            shape,
            indexes: Vec::new(),
            reservation: None,
        })
    }

    fn copy_shape_mask(&mut self, source: &ShapeMask) -> EvaluationResult<ShapeMask> {
        let mut copy = self.new_shape_mask(source.shape)?;
        for &index in &source.indexes {
            self.push_shape_mask_index(&mut copy, index)?;
        }
        Ok(copy)
    }

    fn push_shape_mask_index(
        &mut self,
        mask: &mut ShapeMask,
        index: usize,
    ) -> EvaluationResult<()> {
        ensure_capacity(
            &mut mask.indexes,
            &mut mask.reservation,
            1,
            self.limits.max_array_cells,
            self.execution,
            &self.storage_budget,
            "formula value shape demand mask",
        )?;
        mask.indexes.push(index);
        Ok(())
    }

    fn install_shape_mask(&mut self, mask: ShapeMask) -> EvaluationResult<Option<usize>> {
        if mask.indexes.is_empty() {
            return Ok(None);
        }
        ensure_capacity(
            &mut self.shape_masks,
            &mut self.shape_mask_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value shape demand masks",
        )?;
        let index = self.shape_masks.len();
        self.shape_masks.push(mask);
        Ok(Some(index))
    }

    fn demand_mask_for_shape(
        &mut self,
        demand: ShapeDemand,
        shape: Shape,
    ) -> EvaluationResult<Option<ShapeMask>> {
        let Some(mask) = demand.mask else {
            return Ok(None);
        };
        let source = self.copy_shape_mask_id(mask)?;
        self.expand_shape_mask(source, demand.shape, shape)
            .map(Some)
    }

    fn copy_shape_mask_id(&mut self, index: usize) -> EvaluationResult<ShapeMask> {
        let (shape, length) = self
            .shape_masks
            .get(index)
            .map(|mask| (mask.shape, mask.indexes.len()))
            .ok_or(EvaluationFailure::InvalidExpression(
                "shape demand mask is missing",
            ))?;
        let mut copy = self.new_shape_mask(shape)?;
        ensure_capacity(
            &mut copy.indexes,
            &mut copy.reservation,
            length,
            self.limits.max_array_cells,
            self.execution,
            &self.storage_budget,
            "formula value shape demand mask",
        )?;
        for offset in 0..length {
            self.scalar.charge_work(1)?;
            let value = self
                .shape_masks
                .get(index)
                .and_then(|mask| mask.indexes.get(offset).copied())
                .ok_or(EvaluationFailure::InvalidExpression(
                    "shape demand mask changed during copy",
                ))?;
            copy.indexes.push(value);
        }
        Ok(copy)
    }

    fn is_lazy_handler(node: super::Node<'expr>) -> bool {
        let super::Kind::Function { name } = node.kind() else {
            return false;
        };
        name.eq_ignore_ascii_case("IF")
            || name.eq_ignore_ascii_case("IFERROR")
            || name.eq_ignore_ascii_case("IFNA")
    }

    fn deferred_lazy_children(
        &mut self,
        node: super::Node<'expr>,
        demand: ShapeDemand,
        condition: super::Node<'expr>,
        condition_shape: Shape,
    ) -> EvaluationResult<Option<LazyShapeChildren<'expr>>> {
        let super::Kind::Function { name } = node.kind() else {
            return Ok(None);
        };
        if name.eq_ignore_ascii_case("IF") {
            let count = node.child_count();
            if !(1..=3).contains(&count) {
                return Ok(None);
            }
            let effective_shape = broadcast_shape(demand.shape, condition_shape).ok_or(
                EvaluationFailure::InvalidExpression("incompatible condition shapes"),
            )?;
            let effective_mask = self.demand_mask_for_shape(demand, effective_shape)?;
            let effective_mask = effective_mask
                .map(|mask| self.install_shape_mask(mask))
                .transpose()?
                .flatten();
            let effective_demand = ShapeDemand {
                shape: effective_shape,
                mask: effective_mask,
            };
            if count == 1 {
                return Ok(Some(LazyShapeChildren {
                    first: Some(condition),
                    second: None,
                    first_mask: effective_mask,
                    second_mask: None,
                    len: 1,
                    shape: effective_shape,
                }));
            }
            let cells = self.demand_len(effective_demand)?;
            let mut first_selected = false;
            let mut second_selected = false;
            let mut first_mask = self.new_shape_mask(effective_shape)?;
            let mut second_mask = self.new_shape_mask(effective_shape)?;
            for ordinal in 0..cells {
                let index = self.demand_index(effective_demand, ordinal)?;
                self.charge_cell_work(index)?;
                let value =
                    self.condition_value_at(condition, condition_shape, effective_shape, index)?;
                match self.scalar_logical(value)? {
                    Ok(true) => {
                        first_selected = true;
                        self.push_shape_mask_index(&mut first_mask, index)?;
                    },
                    Ok(false) if count == 3 => {
                        second_selected = true;
                        self.push_shape_mask_index(&mut second_mask, index)?;
                    },
                    Ok(false) | Err(_) => {},
                }
            }
            let first = first_selected.then(|| node.child(1)).flatten();
            let second = second_selected.then(|| node.child(2)).flatten();
            let len = usize::from(first.is_some())
                .checked_add(usize::from(second.is_some()))
                .ok_or(EvaluationFailure::InvalidExpression(
                    "deferred IF child count overflow",
                ))?;
            return Ok(Some(LazyShapeChildren {
                first,
                second,
                first_mask: self.install_shape_mask(first_mask)?,
                second_mask: self.install_shape_mask(second_mask)?,
                len,
                shape: effective_shape,
            }));
        }
        if name.eq_ignore_ascii_case("IFERROR") || name.eq_ignore_ascii_case("IFNA") {
            if node.child_count() != 2 {
                return Ok(None);
            }
            let alternative = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                "error handler alternative is missing",
            ))?;
            let catches_not_available = name.eq_ignore_ascii_case("IFNA");
            let effective_shape = broadcast_shape(demand.shape, condition_shape).ok_or(
                EvaluationFailure::InvalidExpression("incompatible handler shapes"),
            )?;
            let effective_mask = self.demand_mask_for_shape(demand, effective_shape)?;
            let effective_mask = effective_mask
                .map(|mask| self.install_shape_mask(mask))
                .transpose()?
                .flatten();
            let effective_demand = ShapeDemand {
                shape: effective_shape,
                mask: effective_mask,
            };
            let cells = self.demand_len(effective_demand)?;
            let mut alternative_mask = self.new_shape_mask(effective_shape)?;
            let mut alternative_selected = false;
            for ordinal in 0..cells {
                let index = self.demand_index(effective_demand, ordinal)?;
                self.charge_cell_work(index)?;
                let evaluated =
                    self.condition_value_at(condition, condition_shape, effective_shape, index)?;
                let caught = matches!(
                    evaluated,
                    RuntimeValue::Scalar(WorkingValue::Error(error))
                        if !catches_not_available || error == ScalarError::NotAvailable
                );
                if caught {
                    alternative_selected = true;
                    self.push_shape_mask_index(&mut alternative_mask, index)?;
                }
            }
            return Ok(Some(LazyShapeChildren {
                first: None,
                second: alternative_selected.then_some(alternative),
                first_mask: None,
                second_mask: self.install_shape_mask(alternative_mask)?,
                len: usize::from(alternative_selected),
                shape: effective_shape,
            }));
        }
        Ok(None)
    }

    fn evaluate_scalar_at(
        &mut self,
        root: super::Node<'expr>,
        demand: Shape,
        index: usize,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let cache_lookup = self.condition_cache_get(root, demand, index)?;
        if let Some(value) = cache_lookup.value {
            return Ok(value);
        }
        let position = self.offset_position(index, demand)?;
        let mut probe_root = root;
        loop {
            if matches!(probe_root.kind(), super::Kind::Parenthesized) {
                self.scalar.charge_work(1)?;
                probe_root = probe_root
                    .child(0)
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "shape probe parentheses are empty",
                    ))?;
                continue;
            }
            let super::Kind::Array(dimensions) = probe_root.kind() else {
                break;
            };
            let columns = dimensions.columns().ok_or(EvaluationFailure::Unsupported(
                super::UnsupportedKind::Array,
            ))?;
            let array_shape = Shape::new(dimensions.rows(), columns)?;
            let Some(source_index) = projected_array_index(array_shape, demand, index) else {
                return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                    ScalarError::NotAvailable,
                )));
            };
            let row = source_index / columns;
            let column = source_index % columns;
            let row_node = probe_root
                .child(row)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "shape probe row is missing",
                ))?;
            probe_root = row_node
                .child(column)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "shape probe cell is missing",
                ))?;
            break;
        }
        let _probe_depth_reservation = self.enter_value_probe()?;
        let saved_shape_planner = self.suspend_shape_planner();
        let saved_frames = std::mem::take(&mut self.frames);
        let saved_values = std::mem::take(&mut self.values);
        let saved_frame_reservation = self.frame_reservation.take();
        let saved_value_reservation = self.value_reservation.take();
        let saved_argument_contexts = std::mem::take(&mut self.argument_contexts);
        let saved_argument_context_reservation = self.argument_context_reservation.take();
        let saved_matrix = self.matrix.take();
        let saved_mode = self.mode;
        let saved_position = self.position;
        let saved_projection = self.projection;

        self.mode = Mode::Scalar;
        self.position = position;
        self.projection = Some((demand, index));
        let result = self
            .run_from_without_scalar_demand(probe_root)
            .and_then(|value| self.project_scalar(value));
        let result = match result {
            Ok(value) => match cache_lookup.key {
                Some((cache_demand, cache_index)) => self
                    .condition_cache_put(root, cache_demand, cache_index, &value)
                    .map(|()| value),
                None => Ok(value),
            },
            Err(error) => Err(error),
        };

        // `run_from` is deliberately executed with an empty VM stack so a
        // shape continuation can suspend the ordinary evaluator without
        // copying or nesting a second evaluator.  Drop every probe-owned
        // frame/value and its matching reservation before restoring the
        // suspended stacks.  The scalar mode prevents a lazy handler from
        // replacing the caller's active matrix state; restore defensively in
        // case a future value path adds another matrix-producing operation.
        let probe_frames = std::mem::take(&mut self.frames);
        let probe_values = std::mem::take(&mut self.values);
        let probe_frame_reservation = self.frame_reservation.take();
        let probe_value_reservation = self.value_reservation.take();
        let probe_argument_contexts = std::mem::take(&mut self.argument_contexts);
        let probe_argument_context_reservation = self.argument_context_reservation.take();
        let probe_matrix = self.matrix.take();
        drop(probe_frames);
        drop(probe_values);
        drop(probe_frame_reservation);
        drop(probe_value_reservation);
        drop(probe_argument_contexts);
        drop(probe_argument_context_reservation);
        drop(probe_matrix);

        self.restore_shape_planner(saved_shape_planner);
        self.frames = saved_frames;
        self.values = saved_values;
        self.frame_reservation = saved_frame_reservation;
        self.value_reservation = saved_value_reservation;
        self.argument_contexts = saved_argument_contexts;
        self.argument_context_reservation = saved_argument_context_reservation;
        self.matrix = saved_matrix;
        self.mode = saved_mode;
        self.position = saved_position;
        self.projection = saved_projection;
        self.probe_depth -= 1;
        result
    }

    fn condition_value_at(
        &mut self,
        root: super::Node<'expr>,
        condition_shape: Shape,
        demand: Shape,
        index: usize,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let Some(condition_index) = projected_array_index(condition_shape, demand, index) else {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        };
        self.evaluate_scalar_at(root, condition_shape, condition_index)
    }

    /// Evaluate one selected matrix-function node while retaining its result
    /// kind.  A scalar probe would implicitly intersect a successful inverse,
    /// transpose, or product and could therefore mistake an Array for a
    /// scalar formula error in a surrounding reference handler.  This helper
    /// reuses the evaluator's flat VM stacks and restores the suspended
    /// continuation on every return.
    fn reference_kind_matrix_value(
        &mut self,
        root: super::Node<'expr>,
        _function: MatrixFunction,
    ) -> EvaluationResult<ReferenceOperandKind> {
        let value = self.evaluate_matrix_value(root)?;
        let kind = Self::reference_operand_kind_from_runtime(&value);
        drop(value);
        Ok(kind)
    }

    /// A forced-array argument can reenter shape planning from a scalar
    /// probe. Bound that Rust call-stack use independently of the heap VM
    /// stacks, and charge the caller's hierarchical Depth budget as well.
    fn enter_value_probe(&mut self) -> EvaluationResult<Reservation> {
        const MAX_VALUE_PROBE_DEPTH: usize = 32;
        let maximum = MAX_VALUE_PROBE_DEPTH.min(self.limits.scalar.max_stack_entries);
        let observed = self.probe_depth.saturating_add(1);
        if observed > maximum {
            return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Depth,
                u64::try_from(observed).unwrap_or(u64::MAX),
                maximum,
            )));
        }
        let reservation = self
            .execution
            .budget()
            .reserve(Resource::Depth, 1)
            .map_err(EvaluationFailure::ResourceLimit)?;
        self.probe_depth = observed;
        Ok(reservation)
    }

    fn suspend_shape_planner(&mut self) -> SavedShapePlanner<'expr> {
        SavedShapePlanner {
            frames: std::mem::take(&mut self.shape_frames),
            values: std::mem::take(&mut self.shape_values),
            masks: std::mem::take(&mut self.shape_masks),
            frame_reservation: self.shape_frame_reservation.take(),
            value_reservation: self.shape_value_reservation.take(),
            mask_reservation: self.shape_mask_reservation.take(),
        }
    }

    fn restore_shape_planner(&mut self, mut saved: SavedShapePlanner<'expr>) {
        std::mem::swap(&mut self.shape_frames, &mut saved.frames);
        std::mem::swap(&mut self.shape_values, &mut saved.values);
        std::mem::swap(&mut self.shape_masks, &mut saved.masks);
        std::mem::swap(
            &mut self.shape_frame_reservation,
            &mut saved.frame_reservation,
        );
        std::mem::swap(
            &mut self.shape_value_reservation,
            &mut saved.value_reservation,
        );
        std::mem::swap(
            &mut self.shape_mask_reservation,
            &mut saved.mask_reservation,
        );
    }

    fn evaluate_matrix_value(
        &mut self,
        root: super::Node<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let _probe_depth_reservation = self.enter_value_probe()?;
        let saved_shape_planner = self.suspend_shape_planner();
        let saved_frames = std::mem::take(&mut self.frames);
        let saved_values = std::mem::take(&mut self.values);
        let saved_frame_reservation = self.frame_reservation.take();
        let saved_value_reservation = self.value_reservation.take();
        let saved_argument_contexts = std::mem::take(&mut self.argument_contexts);
        let saved_argument_context_reservation = self.argument_context_reservation.take();
        let saved_matrix = self.matrix.take();
        let saved_mode = self.mode;
        let saved_position = self.position;
        let saved_projection = self.projection;

        self.mode = Mode::Matrix;
        self.projection = None;
        let result = self.run_from_without_scalar_demand(root);

        // A failed probe can leave partial values or a nested matrix
        // continuation. Drop those before returning their matching storage
        // reservation to the shared budget.
        let probe_frames = std::mem::take(&mut self.frames);
        let probe_values = std::mem::take(&mut self.values);
        let probe_frame_reservation = self.frame_reservation.take();
        let probe_value_reservation = self.value_reservation.take();
        let probe_argument_contexts = std::mem::take(&mut self.argument_contexts);
        let probe_argument_context_reservation = self.argument_context_reservation.take();
        let probe_matrix = self.matrix.take();
        drop(probe_frames);
        drop(probe_values);
        drop(probe_frame_reservation);
        drop(probe_value_reservation);
        drop(probe_argument_contexts);
        drop(probe_argument_context_reservation);
        drop(probe_matrix);

        self.restore_shape_planner(saved_shape_planner);
        self.frames = saved_frames;
        self.values = saved_values;
        self.frame_reservation = saved_frame_reservation;
        self.value_reservation = saved_value_reservation;
        self.argument_contexts = saved_argument_contexts;
        self.argument_context_reservation = saved_argument_context_reservation;
        self.matrix = saved_matrix;
        self.mode = saved_mode;
        self.position = saved_position;
        self.projection = saved_projection;
        self.probe_depth -= 1;
        result
    }

    fn reference_shape_hint(
        &mut self,
        reference: &'expr Reference,
    ) -> EvaluationResult<Option<Shape>> {
        match self.reference_shape(reference) {
            Err(EvaluationFailure::Unsupported(super::UnsupportedKind::Reference)) => Ok(None),
            result => result,
        }
    }

    fn reference_shape(&mut self, reference: &'expr Reference) -> EvaluationResult<Option<Shape>> {
        let address = match reference {
            Reference::Local(address) => address,
            Reference::Error => return Ok(Some(Shape::new(1, 1)?)),
            Reference::Source { .. } => return Ok(None),
        };
        match address {
            Address::Cell(endpoint) => {
                if self.endpoint_area(endpoint, None)?.is_none() {
                    return Ok(None);
                }
                Ok(Some(Shape::new(1, 1)?))
            },
            Address::Cells(first, second)
            | Address::Columns(first, second)
            | Address::Rows(first, second) => {
                let first_sheet = match self.endpoint_sheet(&first.sheet) {
                    Ok(sheet) => sheet,
                    Err(EvaluationFailure::Unsupported(super::UnsupportedKind::Reference)) => {
                        return Ok(None);
                    },
                    Err(error) => return Err(error),
                };
                let second_sheet =
                    match self.endpoint_sheet_with_inherited(&second.sheet, first_sheet) {
                        Ok(sheet) => sheet,
                        Err(EvaluationFailure::Unsupported(super::UnsupportedKind::Reference)) => {
                            return Ok(None);
                        },
                        Err(error) => return Err(error),
                    };
                let Some(first_index) = self.resolve_sheet(first_sheet)? else {
                    return Ok(None);
                };
                let Some(second_index) = self.resolve_sheet(second_sheet)? else {
                    return Ok(None);
                };
                if first_index != second_index {
                    return Ok(None);
                }
                let first_rect = self.endpoint_rect(first, first_sheet)?;
                let second_rect = self.endpoint_rect(second, second_sheet)?;
                let rect = first_rect.bounding(second_rect)?;
                Ok(Some(Shape::new(rect.rows(), rect.columns())?))
            },
        }
    }

    fn push_shape_frame(&mut self, frame: ShapeFrame<'expr>) -> EvaluationResult<()> {
        ensure_capacity(
            &mut self.shape_frames,
            &mut self.shape_frame_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value shape planner frames",
        )?;
        self.shape_frames.push(frame);
        Ok(())
    }

    fn push_shape_value(&mut self, value: Option<Shape>) -> EvaluationResult<()> {
        ensure_capacity(
            &mut self.shape_values,
            &mut self.shape_value_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value shape planner values",
        )?;
        self.shape_values.push(value);
        Ok(())
    }

    fn pop_shape_value(&mut self) -> EvaluationResult<Option<Shape>> {
        self.shape_values
            .pop()
            .ok_or(EvaluationFailure::InvalidExpression(
                "shape planner value stack underflow",
            ))
    }

    fn runtime_shape(&mut self, value: &RuntimeValue<'expr>) -> EvaluationResult<Shape> {
        match value {
            RuntimeValue::Array(array) => Ok(array.shape),
            RuntimeValue::Areas(areas) if !areas.is_list && areas.areas.len() == 1 => {
                let area = areas
                    .areas
                    .first()
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference area disappeared",
                    ))?;
                Shape::new(area.rect.rows(), area.rect.columns())
            },
            RuntimeValue::Areas(_) => Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::Reference,
            )),
            RuntimeValue::Empty
            | RuntimeValue::Missing
            | RuntimeValue::Scalar(_)
            | RuntimeValue::ScalarCell(_) => Shape::new(1, 1),
        }
    }

    fn offset_position(&self, index: usize, shape: Shape) -> EvaluationResult<Position<'position>> {
        let row = index / shape.columns();
        let column = index % shape.columns();
        let row =
            self.position
                .row
                .checked_add(row)
                .ok_or(EvaluationFailure::InvalidExpression(
                    "matrix row position overflow",
                ))?;
        let column = self.position.column.checked_add(column).ok_or(
            EvaluationFailure::InvalidExpression("matrix column position overflow"),
        )?;
        Ok(Position::new(self.position.sheet, row, column))
    }

    fn select_matrix_value(
        &mut self,
        value: &RuntimeValue<'expr>,
        output: Shape,
        index: usize,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        match value {
            RuntimeValue::Empty => Ok(RuntimeElement::Empty),
            RuntimeValue::Missing => Ok(RuntimeElement::Missing),
            RuntimeValue::Scalar(value) => self.clone_working(value).map(RuntimeElement::Present),
            RuntimeValue::ScalarCell(area) => {
                let value = self.project_scalar(RuntimeValue::ScalarCell(*area))?;
                self.runtime_to_element(value)
            },
            RuntimeValue::Array(array) => array_element_for(array, output, index)
                .map(|element| self.clone_element(element))
                .transpose()?
                .map_or(
                    Ok(RuntimeElement::Present(WorkingValue::Error(
                        ScalarError::NotAvailable,
                    ))),
                    Ok,
                ),
            RuntimeValue::Areas(areas) => self.select_area_element(areas, output, index),
        }
    }

    fn select_area_element(
        &mut self,
        areas: &RuntimeAreaSet<'expr>,
        output: Shape,
        index: usize,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        if areas.is_list || areas.areas.len() != 1 {
            return Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::Reference,
            ));
        }
        let area = areas
            .areas
            .first()
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference area is empty",
            ))?;
        let source =
            Cuboid::from_origin_extents([0, 0, 0], [1, area.rect.rows(), area.rect.columns()])
                .ok_or(EvaluationFailure::InvalidExpression(
                    "reference shape overflow",
                ))?;
        let destination =
            Cuboid::from_origin_extents([0, 0, 0], [1, output.rows(), output.columns()]).ok_or(
                EvaluationFailure::InvalidExpression("output shape overflow"),
            )?;
        let point = [0, index / output.columns(), index % output.columns()];
        let projected = if area.rect.rows() == 1 && area.rect.columns() == 1 {
            source.project_2d_singleton(destination, point)
        } else if area.rect.rows() == 1 {
            source.project_2d_row(destination, point)
        } else if area.rect.columns() == 1 {
            source.project_2d_column(destination, point)
        } else {
            source.project_2d_matrix(destination, point)
        };
        let Some(projected) = projected else {
            return Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        };
        self.charge_cell_work(index)?;
        let read = self.read_reference_cell(
            area.sheet,
            area.rect.row_start + projected[1],
            area.rect.column_start + projected[2],
        )?;
        self.read_to_element(read)
    }

    fn apply(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        match node.kind() {
            super::Kind::Function { name } => self.apply_function(node, name),
            super::Kind::Prefix(operator) => {
                let value = self.pop_value()?;
                let result = self.apply_unary(operator, value)?;
                self.push_value(result)
            },
            super::Kind::Postfix(operator) => {
                let value = self.pop_value()?;
                let result = self.apply_postfix(operator, value)?;
                self.push_value(result)
            },
            super::Kind::Infix(operator) => {
                let right = self.pop_value()?;
                let left = self.pop_value()?;
                let result = self.apply_infix(operator, left, right)?;
                self.push_value(result)
            },
            _ => Err(EvaluationFailure::InvalidExpression(
                "non-operator reached value apply frame",
            )),
        }
    }

    fn apply_matrix_function(
        &mut self,
        node: super::Node<'expr>,
        _function: MatrixFunction,
    ) -> EvaluationResult<()> {
        let name = node
            .function_name()
            .ok_or(EvaluationFailure::InvalidExpression(
                "matrix apply frame has no function name",
            ))?;
        let count = node.child_count();
        let mut arguments = Vec::new();
        let mut reservation = None;
        ensure_capacity(
            &mut arguments,
            &mut reservation,
            count,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value matrix function arguments",
        )?;
        for _ in 0..count {
            arguments.push(self.pop_value()?);
        }
        arguments.reverse();
        let value = self.apply_matrix_values(name, arguments)?;
        self.push_value(value)
    }

    fn apply_function(&mut self, node: super::Node<'expr>, name: &str) -> EvaluationResult<()> {
        let count = node.child_count();
        let mut reservation = None;
        let mut arguments = Vec::new();
        ensure_capacity(
            &mut arguments,
            &mut reservation,
            count,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value function arguments",
        )?;
        for _ in 0..count {
            arguments.push(self.pop_value()?);
        }
        arguments.reverse();

        if aggregate::is_aggregate_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = aggregate::apply(self, name, arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if statistical::is_statistical_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = statistical::apply(self, name, arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if descriptive::is_descriptive_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = descriptive::apply(self, name, arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if paired::is_paired_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = paired::apply(self, node, name, arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if order::is_order_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = order::apply(self, name, arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if database::is_database_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = database::apply(self, name, arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if conditional::is_conditional_function(name) {
            let cacheable = self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = conditional::apply(self, name, &arguments)?;
            if cacheable {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        if complex::is_complex_sequence_function(name) {
            let cacheable_sequence =
                self.projection.is_some() && self.cacheable_scalar_branch(node)?;
            if cacheable_sequence {
                if let Some(value) = self.demand_cache_get(node)? {
                    return self.push_value(value);
                }
            }
            let value = complex::apply_sequence(self, name, arguments)?;
            if cacheable_sequence {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        let cacheable_sequence = self.projection.is_some()
            && (name.eq_ignore_ascii_case("AND") || name.eq_ignore_ascii_case("OR"))
            && self.cacheable_scalar_branch(node)?;
        if cacheable_sequence {
            if let Some(value) = self.demand_cache_get(node)? {
                return self.push_value(value);
            }
        }
        if name.eq_ignore_ascii_case("AND") || name.eq_ignore_ascii_case("OR") {
            let value = self.apply_sequence(arguments, name.eq_ignore_ascii_case("AND"))?;
            if cacheable_sequence {
                self.demand_cache_put(node, &value)?;
            }
            return self.push_value(value);
        }

        // A list is accepted by the sequence functions above, but cannot be
        // converted to a scalar or a rectangular array for these functions.
        for argument in &mut arguments {
            if matches!(argument, RuntimeValue::Areas(areas) if areas.is_list) {
                *argument = RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Value));
            }
        }

        // XOR is a Logical function rather than a NumberSequenceList.  In
        // matrix mode, evaluate one scalar XOR per broadcast cell; in scalar
        // mode an array argument is explicitly intersected first.
        if name.eq_ignore_ascii_case("XOR")
            && self.mode == Mode::Matrix
            && arguments.iter().any(RuntimeValue::is_array_like)
        {
            let value = self.map_function(node, name, arguments)?;
            return self.push_value(value);
        }

        if arguments.iter().any(RuntimeValue::is_array_like) {
            if self.mode == Mode::Matrix {
                let value = self.map_function(node, name, arguments)?;
                return self.push_value(value);
            }
            let mut projected = Vec::new();
            let mut projected_reservation = None;
            ensure_capacity(
                &mut projected,
                &mut projected_reservation,
                count,
                self.limits.scalar.max_stack_entries,
                self.execution,
                &self.storage_budget,
                "formula value scalar arguments",
            )?;
            for argument in arguments {
                projected.push(self.project_scalar(argument)?);
            }
            let result = self.scalar_apply_function(node, name, projected)?;
            return self.push_value(result);
        }

        let mut scalar_arguments = Vec::new();
        let mut scalar_reservation = None;
        ensure_capacity(
            &mut scalar_arguments,
            &mut scalar_reservation,
            count,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value scalar arguments",
        )?;
        for argument in arguments {
            scalar_arguments.push(argument);
        }
        let result = self.scalar_apply_function(node, name, scalar_arguments)?;
        self.push_value(result)
    }

    fn apply_sequence(
        &mut self,
        arguments: Vec<RuntimeValue<'expr>>,
        conjunction: bool,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let mut result = conjunction;
        let mut error = None;
        for (argument_index, argument) in arguments.into_iter().enumerate() {
            self.scalar.charge_work(1)?;
            match argument {
                RuntimeValue::Areas(areas) => {
                    for area_index in 0..areas.areas.len() {
                        let area = &areas.areas[area_index];
                        for row in area.rect.row_start..area.rect.row_end {
                            for column in area.rect.column_start..area.rect.column_end {
                                let linear = row
                                    .saturating_sub(area.rect.row_start)
                                    .checked_mul(area.rect.columns())
                                    .and_then(|offset| {
                                        offset.checked_add(
                                            column.saturating_sub(area.rect.column_start),
                                        )
                                    })
                                    .unwrap_or(usize::MAX);
                                self.charge_cell_work(linear)?;
                                let read = self.read_reference_cell(area.sheet, row, column)?;
                                let element = self.read_to_element(read)?;
                                self.accumulate_reference_element(
                                    element,
                                    conjunction,
                                    &mut result,
                                    &mut error,
                                )?;
                            }
                        }
                    }
                },
                RuntimeValue::Array(array) => {
                    for (index, element) in array.cells.into_iter().enumerate() {
                        self.charge_cell_work(index)?;
                        self.accumulate_array_element(
                            element,
                            conjunction,
                            &mut result,
                            &mut error,
                        )?;
                    }
                },
                RuntimeValue::Empty => {},
                RuntimeValue::Missing => {
                    if error.is_none() {
                        error = Some(ScalarError::Value);
                    }
                },
                RuntimeValue::Scalar(value) => {
                    self.accumulate_scalar_sequence(
                        value,
                        argument_index,
                        conjunction,
                        &mut result,
                        &mut error,
                    )?;
                },
                RuntimeValue::ScalarCell(area) => {
                    let projected = self.project_scalar(RuntimeValue::ScalarCell(area))?;
                    match projected {
                        RuntimeValue::Scalar(value) => self.accumulate_scalar_sequence(
                            value,
                            argument_index,
                            conjunction,
                            &mut result,
                            &mut error,
                        )?,
                        RuntimeValue::Empty => {},
                        RuntimeValue::Missing => {
                            if error.is_none() {
                                error = Some(ScalarError::Value);
                            }
                        },
                        RuntimeValue::ScalarCell(_)
                        | RuntimeValue::Array(_)
                        | RuntimeValue::Areas(_) => {
                            return Err(EvaluationFailure::InvalidExpression(
                                "scalar cell projection remained non-scalar",
                            ));
                        },
                    }
                },
            }
        }
        Ok(RuntimeValue::Scalar(match error {
            Some(error) => WorkingValue::Error(error),
            None => WorkingValue::Logical(result),
        }))
    }

    fn accumulate_scalar_sequence(
        &mut self,
        value: WorkingValue<'expr>,
        _argument_index: usize,
        conjunction: bool,
        result: &mut bool,
        error: &mut Option<ScalarError>,
    ) -> EvaluationResult<()> {
        match value {
            WorkingValue::Error(value) => {
                if error.is_none() {
                    *error = Some(value);
                }
            },
            WorkingValue::Logical(value) => {
                if conjunction {
                    *result &= value;
                } else {
                    *result |= value;
                }
            },
            WorkingValue::Number(value) => {
                if conjunction {
                    *result &= value != 0.0;
                } else {
                    *result |= value != 0.0;
                }
            },
            WorkingValue::Text(value) => {
                match super::to_number(WorkingValue::Text(value), &mut self.scalar)? {
                    Ok(value) => {
                        if conjunction {
                            *result &= value != 0.0;
                        } else {
                            *result |= value != 0.0;
                        }
                    },
                    Err(value) => {
                        if error.is_none() {
                            *error = Some(value);
                        }
                    },
                }
            },
            WorkingValue::Complex(_) => {
                if error.is_none() {
                    *error = Some(ScalarError::Value);
                }
            },
        }
        Ok(())
    }

    fn accumulate_reference_element(
        &mut self,
        element: RuntimeElement<'expr>,
        conjunction: bool,
        result: &mut bool,
        error: &mut Option<ScalarError>,
    ) -> EvaluationResult<()> {
        // NumberSequenceList references contain Number and Error members.
        // Empty, Text, and distinguished Logical cells are ignored.
        if let RuntimeElement::Present(value) = element {
            match value {
                WorkingValue::Number(value) => {
                    if conjunction {
                        *result &= value != 0.0;
                    } else {
                        *result |= value != 0.0;
                    }
                },
                WorkingValue::Error(value) => {
                    if error.is_none() {
                        *error = Some(value);
                    }
                },
                WorkingValue::Logical(_) | WorkingValue::Text(_) | WorkingValue::Complex(_) => {},
            }
        }
        Ok(())
    }

    fn accumulate_array_element(
        &mut self,
        element: RuntimeElement<'expr>,
        conjunction: bool,
        result: &mut bool,
        error: &mut Option<ScalarError>,
    ) -> EvaluationResult<()> {
        match element {
            RuntimeElement::Empty => {
                let value = false;
                if conjunction {
                    *result &= value;
                }
            },
            RuntimeElement::Missing => {
                if error.is_none() {
                    *error = Some(ScalarError::Value);
                }
            },
            RuntimeElement::Present(value) => {
                match scalar::logical(&mut self.scalar, scalar::Slot::Value(value))? {
                    Ok(value) => {
                        if conjunction {
                            *result &= value;
                        } else {
                            *result |= value;
                        }
                    },
                    Err(value) => {
                        if error.is_none() {
                            *error = Some(value);
                        }
                    },
                }
            },
        }
        Ok(())
    }

    fn apply_unary(
        &mut self,
        operator: super::PrefixOperator,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        self.map_unary(value, UnaryOperation::Prefix(operator))
    }

    fn apply_postfix(
        &mut self,
        operator: super::PostfixOperator,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        self.map_unary(value, UnaryOperation::Postfix(operator))
    }

    fn apply_infix(
        &mut self,
        operator: super::InfixOperator,
        left: RuntimeValue<'expr>,
        right: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match operator {
            super::InfixOperator::Range => self.combine_range(left, right),
            super::InfixOperator::Intersection => self.intersect_ranges(left, right),
            super::InfixOperator::Union => self.union_ranges(left, right),
            _ => self.map_binary(left, right, operator),
        }
    }

    fn map_unary(
        &mut self,
        value: RuntimeValue<'expr>,
        operation: UnaryOperation,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match value {
            value @ (RuntimeValue::Array(_) | RuntimeValue::Areas(_))
                if self.mode == Mode::Scalar =>
            {
                let value = self.project_scalar(value)?;
                self.apply_scalar_unary(value, operation)
            },
            RuntimeValue::Areas(areas) => {
                let value = self.materialize_for_array(RuntimeValue::Areas(areas))?;
                self.map_unary(value, operation)
            },
            RuntimeValue::Array(array) => {
                let shape = array.shape;
                let origin = array.origin;
                let (mut output, reservation) = self.new_element_vec(shape.cell_count().ok_or(
                    EvaluationFailure::InvalidExpression("unary array cell count overflow"),
                )?)?;
                for (index, element) in array.cells.into_iter().enumerate() {
                    self.charge_cell_work(index)?;
                    output.push(self.apply_element_unary(element, operation)?);
                }
                self.make_array(shape, output, reservation, origin)
            },
            value => self.apply_scalar_unary(value, operation),
        }
    }

    fn apply_element_unary(
        &mut self,
        element: RuntimeElement<'expr>,
        operation: UnaryOperation,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        match element {
            RuntimeElement::Empty => self
                .apply_slot_unary(scalar::Slot::Empty, operation)
                .map(slot_to_element),
            RuntimeElement::Missing => Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::Value,
            ))),
            RuntimeElement::Present(value) => self
                .apply_slot_unary(scalar::Slot::Value(value), operation)
                .map(slot_to_element),
        }
    }

    fn apply_scalar_unary(
        &mut self,
        value: RuntimeValue<'expr>,
        operation: UnaryOperation,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let slot = self.value_to_slot(value)?;
        self.apply_slot_unary(slot, operation).map(slot_to_runtime)
    }

    fn apply_slot_unary(
        &mut self,
        value: scalar::Slot<'expr>,
        operation: UnaryOperation,
    ) -> EvaluationResult<scalar::Slot<'expr>> {
        match operation {
            UnaryOperation::Prefix(operator) => scalar::prefix(&mut self.scalar, operator, value),
            UnaryOperation::Postfix(operator) => scalar::postfix(&mut self.scalar, operator, value),
        }
    }

    fn map_binary(
        &mut self,
        left: RuntimeValue<'expr>,
        right: RuntimeValue<'expr>,
        operator: super::InfixOperator,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if self.mode == Mode::Scalar {
            let left = self.project_scalar(left)?;
            let right = self.project_scalar(right)?;
            let left = self.value_to_slot(left)?;
            let right = self.value_to_slot(right)?;
            return self.scalar_apply_infix(operator, left, right);
        }
        let left = self.materialize_for_array(left)?;
        let right = self.materialize_for_array(right)?;
        if !matches!(left, RuntimeValue::Array(_)) && !matches!(right, RuntimeValue::Array(_)) {
            let left = self.value_to_slot(left)?;
            let right = self.value_to_slot(right)?;
            return self.scalar_apply_infix(operator, left, right);
        }
        let left_shape = self.runtime_shape(&left)?;
        let right_shape = self.runtime_shape(&right)?;
        let shape = broadcast_shape(left_shape, right_shape).ok_or(
            EvaluationFailure::InvalidExpression("incompatible array shapes"),
        )?;
        let cells = shape
            .cell_count()
            .ok_or(EvaluationFailure::InvalidExpression(
                "binary array cell count overflow",
            ))?;
        let (mut output, reservation) = self.new_element_vec(cells)?;
        for index in 0..cells {
            self.charge_cell_work(index)?;
            let left_element = self.element_for_value(&left, shape, index)?;
            let right_element = self.element_for_value(&right, shape, index)?;
            output.push(self.apply_element_binary(left_element, right_element, operator)?);
        }
        self.make_array(shape, output, reservation, None)
    }

    fn apply_element_binary(
        &mut self,
        left: RuntimeElement<'expr>,
        right: RuntimeElement<'expr>,
        operator: super::InfixOperator,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        let left = element_to_slot(left);
        let right = element_to_slot(right);
        self.scalar_apply_infix(operator, left, right)
            .map(|value| match value {
                RuntimeValue::Empty => RuntimeElement::Empty,
                RuntimeValue::Missing => RuntimeElement::Missing,
                RuntimeValue::Scalar(value) => RuntimeElement::Present(value),
                RuntimeValue::ScalarCell(_) | RuntimeValue::Array(_) | RuntimeValue::Areas(_) => {
                    RuntimeElement::Missing
                },
            })
    }

    fn map_function(
        &mut self,
        node: super::Node<'expr>,
        name: &str,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        if super::text::is_text_function(name) {
            return self.map_text_function(node, name, arguments);
        }
        let mut arguments = arguments;
        for argument in &mut arguments {
            if matches!(argument, RuntimeValue::Areas(_)) {
                let replacement = std::mem::replace(argument, RuntimeValue::Missing);
                *argument = self.materialize_for_array(replacement)?;
            }
        }
        let mut shape = Shape::new(1, 1)?;
        for argument in &arguments {
            if let RuntimeValue::Array(array) = argument {
                shape = broadcast_shape(shape, array.shape).ok_or(
                    EvaluationFailure::InvalidExpression("incompatible function array shapes"),
                )?;
            }
        }
        let (mut output, reservation) = self.new_element_vec(shape.cell_count().ok_or(
            EvaluationFailure::InvalidExpression("function array cell count overflow"),
        )?)?;
        for index in 0..shape.cell_count().unwrap_or(0) {
            self.charge_cell_work(index)?;
            let mut scalar_arguments = Vec::new();
            let mut reservation = None;
            ensure_capacity(
                &mut scalar_arguments,
                &mut reservation,
                arguments.len(),
                self.limits.scalar.max_stack_entries,
                self.execution,
                &self.storage_budget,
                "formula value function cell arguments",
            )?;
            for argument in &arguments {
                scalar_arguments.push(element_to_runtime(
                    self.element_for_value(argument, shape, index)?,
                ));
            }
            let value = self.scalar_apply_function(node, name, scalar_arguments)?;
            output.push(self.runtime_to_element(value)?);
        }
        self.make_array(shape, output, reservation, None)
    }

    /// Map ordinary text functions without materializing reference inputs.
    /// One argument buffer is reused across output cells; its lease outlives
    /// both the buffer and any owned text still in it on a failure path.
    fn map_text_function(
        &mut self,
        node: super::Node<'expr>,
        name: &str,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let mut shape = Shape::new(1, 1)?;
        for argument in &arguments {
            shape = broadcast_shape(shape, self.runtime_shape(argument)?).ok_or(
                EvaluationFailure::InvalidExpression("incompatible function array shapes"),
            )?;
        }
        let cells = shape
            .cell_count()
            .ok_or(EvaluationFailure::InvalidExpression(
                "text function array cell count overflow",
            ))?;
        // Declare leases before their buffers so error unwinding frees the
        // allocation before returning its budget to the caller.
        let output_reservation;
        let mut output;
        (output, output_reservation) = self.new_element_vec(cells)?;
        let mut argument_reservation = None;
        let mut scalar_arguments = Vec::new();
        ensure_capacity(
            &mut scalar_arguments,
            &mut argument_reservation,
            arguments.len(),
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula text function cell arguments",
        )?;
        let rejects_empty = name.eq_ignore_ascii_case("TEXT");
        for index in 0..cells {
            self.charge_cell_work(index)?;
            let mut first_empty = false;
            for (argument_index, argument) in arguments.iter().enumerate() {
                let element = self.select_matrix_value(argument, shape, index)?;
                first_empty |= rejects_empty
                    && argument_index == 0
                    && matches!(element, RuntimeElement::Empty);
                scalar_arguments.push(scalar::argument(
                    name,
                    argument_index,
                    element_to_slot(element),
                ));
            }
            scalar::refuse_empty_text_value(&mut scalar_arguments, first_empty);
            let value = scalar::eager(
                &mut self.scalar,
                node,
                name,
                scalar_arguments.drain(..).map(Ok),
            )?;
            output.push(RuntimeElement::Present(value));
        }
        self.make_array(shape, output, output_reservation, None)
    }

    fn push_frame(&mut self, frame: ValueFrame<'expr>) -> EvaluationResult<()> {
        ensure_capacity(
            &mut self.frames,
            &mut self.frame_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value frame stack",
        )?;
        self.frames.push(frame);
        Ok(())
    }

    fn push_value(&mut self, value: RuntimeValue<'expr>) -> EvaluationResult<()> {
        ensure_capacity(
            &mut self.values,
            &mut self.value_reservation,
            1,
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value stack",
        )?;
        self.values.push(value);
        Ok(())
    }

    fn pop_value(&mut self) -> EvaluationResult<RuntimeValue<'expr>> {
        self.values
            .pop()
            .ok_or(EvaluationFailure::InvalidExpression(
                "value stack underflow",
            ))
    }

    fn push_scalar(&mut self, value: WorkingValue<'expr>) -> EvaluationResult<()> {
        self.push_value(RuntimeValue::Scalar(value))
    }

    fn local_limit(&self, resource: Resource, observed: u64, limit: usize) -> ResourceLimit {
        ResourceLimit {
            resource,
            observed,
            limit: u64::try_from(limit).unwrap_or(u64::MAX),
            scope: Arc::from(VALUE_SCOPE),
        }
    }

    fn sheet_name(&self, sheet: SheetRef<'expr>) -> &str {
        match sheet {
            SheetRef::Current => self.position.sheet,
            SheetRef::Named(name) => name,
        }
    }

    fn resolve_sheet(&self, sheet: SheetRef<'expr>) -> EvaluationResult<Option<usize>> {
        self.resolver
            .sheet_index(self.sheet_name(sheet), self.execution)
    }

    fn charge_cell_work(&mut self, index: usize) -> EvaluationResult<()> {
        self.scalar.charge_work(1)?;
        if index % VALUE_CHECK_CHUNK == 0 {
            self.execution.check().map_err(map_execution_error)?;
        }
        Ok(())
    }

    fn read_reference_cell(
        &mut self,
        sheet: SheetRef<'expr>,
        row: usize,
        column: usize,
    ) -> EvaluationResult<CellRead<'expr>> {
        let observed = self
            .reference_cells_read
            .checked_add(1)
            .ok_or_else(|| self.reference_read_limit_error(usize::MAX))?;
        if observed > self.limits.max_reference_cells {
            return Err(self.reference_read_limit_error(observed));
        }
        self.execution.check().map_err(map_execution_error)?;
        let position = self.position;
        let sheet_name = match sheet {
            SheetRef::Current => position.sheet,
            SheetRef::Named(name) => name,
        };
        let resolver = self.resolver;
        let read = resolver.read_cell(sheet_name, row, column, self.execution)?;
        self.reference_cells_read = observed;
        // A resolver may perform a synchronous provider operation while
        // returning a borrowed cell. Check again before the value enters any
        // retained array or scalar result so cancellation cannot be hidden by
        // a successful final read.
        self.execution.check().map_err(map_execution_error)?;
        Ok(read)
    }

    fn reference_read_limit_error(&self, observed: usize) -> EvaluationFailure {
        EvaluationFailure::ResourceLimit(self.local_limit(
            Resource::Objects,
            u64::try_from(observed).unwrap_or(u64::MAX),
            self.limits.max_reference_cells,
        ))
    }

    fn new_element_vec(
        &mut self,
        cells: usize,
    ) -> EvaluationResult<(Vec<RuntimeElement<'expr>>, Option<Reservation>)> {
        if cells > self.limits.max_array_cells {
            return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                Resource::Objects,
                u64::try_from(cells).unwrap_or(u64::MAX),
                self.limits.max_array_cells,
            )));
        }
        let mut values = Vec::new();
        let mut reservation = None;
        ensure_capacity(
            &mut values,
            &mut reservation,
            cells,
            self.limits.max_array_cells,
            self.execution,
            &self.storage_budget,
            "formula value array cells",
        )?;
        // Return the storage token with its vector.  Keeping them paired
        // prevents a nested materialization from overwriting a single
        // evaluator-level reservation slot while the first vector is live.
        Ok((values, reservation))
    }

    fn make_array(
        &mut self,
        shape: Shape,
        cells: Vec<RuntimeElement<'expr>>,
        reservation: Option<Reservation>,
        origin: Option<RuntimeArea<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let expected = shape
            .cell_count()
            .ok_or(EvaluationFailure::InvalidExpression(
                "array cell count overflow",
            ))?;
        if cells.len() != expected {
            return Err(EvaluationFailure::InvalidExpression(
                "array storage length disagrees with shape",
            ));
        }
        Ok(RuntimeValue::Array(RuntimeArrayValue {
            shape,
            cells,
            _reservation: reservation,
            origin,
            preserve_scalar_result: false,
        }))
    }

    fn value_to_slot(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<scalar::Slot<'expr>> {
        match value {
            RuntimeValue::Empty => Ok(scalar::Slot::Empty),
            RuntimeValue::Missing => {
                Ok(scalar::Slot::Value(WorkingValue::Error(ScalarError::Value)))
            },
            RuntimeValue::Scalar(value) => Ok(scalar::Slot::Value(value)),
            RuntimeValue::ScalarCell(area) => {
                let projected = self.project_scalar(RuntimeValue::ScalarCell(area))?;
                self.value_to_slot(projected)
            },
            RuntimeValue::Array(array) => {
                let projected = self.project_scalar(RuntimeValue::Array(array))?;
                self.value_to_slot(projected)
            },
            RuntimeValue::Areas(areas) => {
                let projected = self.project_area_value(&areas)?;
                self.value_to_slot(projected)
            },
        }
    }

    fn scalar_logical(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<Result<bool, ScalarError>> {
        match value {
            RuntimeValue::Empty => scalar::logical(&mut self.scalar, scalar::Slot::Empty),
            RuntimeValue::Missing => Ok(Err(ScalarError::Value)),
            RuntimeValue::Scalar(value) => {
                scalar::logical(&mut self.scalar, scalar::Slot::Value(value))
            },
            RuntimeValue::ScalarCell(area) => {
                let projected = self.project_scalar(RuntimeValue::ScalarCell(area))?;
                self.scalar_logical(projected)
            },
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) => Ok(Err(ScalarError::Value)),
        }
    }

    fn scalar_apply_infix(
        &mut self,
        operator: super::InfixOperator,
        left: scalar::Slot<'expr>,
        right: scalar::Slot<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        scalar::infix(&mut self.scalar, operator, left, right).map(slot_to_runtime)
    }

    fn scalar_apply_function(
        &mut self,
        node: super::Node<'expr>,
        name: &str,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        let mut reservation = None;
        let mut scalar_arguments = Vec::new();
        ensure_capacity(
            &mut scalar_arguments,
            &mut reservation,
            arguments.len(),
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value scalar arguments",
        )?;
        let mut first_empty = false;
        for (index, argument) in arguments.into_iter().enumerate() {
            let slot = self.value_to_slot(argument)?;
            first_empty |= index == 0
                && matches!(slot, scalar::Slot::Empty)
                && name.eq_ignore_ascii_case("TEXT");
            scalar_arguments.push(scalar::argument(name, index, slot));
        }
        scalar::refuse_empty_text_value(&mut scalar_arguments, first_empty);
        scalar::eager(
            &mut self.scalar,
            node,
            name,
            scalar_arguments.into_iter().map(Ok),
        )
        .map(RuntimeValue::Scalar)
    }

    fn element_for_value(
        &mut self,
        value: &RuntimeValue<'expr>,
        shape: Shape,
        index: usize,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        match value {
            RuntimeValue::Empty => Ok(RuntimeElement::Empty),
            RuntimeValue::Missing => Ok(RuntimeElement::Missing),
            RuntimeValue::Scalar(value) => Ok(RuntimeElement::Present(self.clone_working(value)?)),
            RuntimeValue::ScalarCell(area) => {
                let projected = self.project_scalar(RuntimeValue::ScalarCell(*area))?;
                self.element_for_value(&projected, shape, index)
            },
            RuntimeValue::Array(array) => array_element_for(array, shape, index)
                .map(|element| self.clone_element(element))
                .transpose()
                .map(|element| {
                    element.unwrap_or(RuntimeElement::Present(WorkingValue::Error(
                        ScalarError::NotAvailable,
                    )))
                }),
            RuntimeValue::Areas(_) => Err(EvaluationFailure::InvalidExpression(
                "area remained before function broadcast",
            )),
        }
    }

    fn clone_element(
        &mut self,
        element: &RuntimeElement<'expr>,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        match element {
            RuntimeElement::Empty => Ok(RuntimeElement::Empty),
            RuntimeElement::Missing => Ok(RuntimeElement::Missing),
            RuntimeElement::Present(value) => {
                Ok(RuntimeElement::Present(self.clone_working(value)?))
            },
        }
    }

    fn clone_working(
        &mut self,
        value: &WorkingValue<'expr>,
    ) -> EvaluationResult<WorkingValue<'expr>> {
        match value {
            WorkingValue::Number(value) => Ok(WorkingValue::Number(*value)),
            WorkingValue::Logical(value) => Ok(WorkingValue::Logical(*value)),
            WorkingValue::Error(error) => Ok(WorkingValue::Error(*error)),
            WorkingValue::Complex(value) => Ok(WorkingValue::Complex(*value)),
            WorkingValue::Text(text) => {
                if let Cow::Borrowed(value) = &text.text {
                    return Ok(WorkingValue::Text(TextValue::borrowed(value)));
                }
                let length = text.text.len();
                let reservation = self
                    .scalar
                    .reserve_storage(length, "formula value text clone")?;
                let mut output = String::new();
                if length != 0 {
                    output.try_reserve_exact(length).map_err(|source| {
                        EvaluationFailure::Allocation {
                            resource: "formula value text clone",
                            source,
                        }
                    })?;
                    output.push_str(text.text.as_ref());
                }
                Ok(WorkingValue::Text(TextValue::owned(output, reservation)))
            },
        }
    }

    fn materialize_for_array(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match value {
            RuntimeValue::Areas(areas) if areas.is_list => Ok(RuntimeValue::Scalar(
                WorkingValue::Error(ScalarError::Value),
            )),
            RuntimeValue::Areas(areas) => self.materialize_areas(areas).map(RuntimeValue::Array),
            RuntimeValue::ScalarCell(area) => self.project_scalar(RuntimeValue::ScalarCell(area)),
            other => Ok(other),
        }
    }

    fn materialize_areas(
        &mut self,
        areas: RuntimeAreaSet<'expr>,
    ) -> EvaluationResult<RuntimeArrayValue<'expr>> {
        if areas.is_list || areas.areas.len() != 1 {
            return Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::Reference,
            ));
        }
        let area = areas
            .areas
            .into_iter()
            .next()
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference area is empty",
            ))?;
        let rows = area.rect.rows();
        let columns = area.rect.columns();
        let shape = Shape::new(rows, columns)?;
        let cells = shape
            .cell_count()
            .ok_or(EvaluationFailure::InvalidExpression(
                "reference array count overflow",
            ))?;
        if cells > self.limits.max_reference_cells || cells > self.limits.max_array_cells {
            return Err(EvaluationFailure::ResourceLimit(
                self.local_limit(
                    Resource::Objects,
                    u64::try_from(cells).unwrap_or(u64::MAX),
                    self.limits
                        .max_reference_cells
                        .min(self.limits.max_array_cells),
                ),
            ));
        }
        let (mut output, reservation) = self.new_element_vec(cells)?;
        for row in area.rect.row_start..area.rect.row_end {
            for column in area.rect.column_start..area.rect.column_end {
                let index = (row - area.rect.row_start)
                    .checked_mul(columns)
                    .and_then(|index| index.checked_add(column - area.rect.column_start))
                    .ok_or(EvaluationFailure::InvalidExpression(
                        "reference array index overflow",
                    ))?;
                self.charge_cell_work(index)?;
                let read = self.read_reference_cell(area.sheet, row, column)?;
                output.push(self.read_to_element(read)?);
            }
        }
        Ok(RuntimeArrayValue {
            shape,
            cells: output,
            _reservation: reservation,
            origin: Some(area),
            preserve_scalar_result: false,
        })
    }

    fn read_to_element(
        &mut self,
        read: CellRead<'expr>,
    ) -> EvaluationResult<RuntimeElement<'expr>> {
        match read {
            CellRead::Empty => Ok(RuntimeElement::Empty),
            CellRead::Number(value) if value.is_finite() => {
                Ok(RuntimeElement::Present(WorkingValue::Number(value)))
            },
            CellRead::Number(_) => Ok(RuntimeElement::Present(WorkingValue::Error(
                ScalarError::Number,
            ))),
            CellRead::Logical(value) => Ok(RuntimeElement::Present(WorkingValue::Logical(value))),
            CellRead::Error(error) => Ok(RuntimeElement::Present(WorkingValue::Error(error))),
            CellRead::Unsupported => Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::CellValue,
            )),
            CellRead::Text(text) => {
                if text.len() > self.limits.scalar.max_text_bytes {
                    return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                        Resource::Memory,
                        u64::try_from(text.len()).unwrap_or(u64::MAX),
                        self.limits.scalar.max_text_bytes,
                    )));
                }
                // The expression and resolver share the result lifetime, so
                // resolver text remains a borrowed value view.  No copy or
                // reservation is needed for a source-backed cell string.
                Ok(RuntimeElement::Present(WorkingValue::Text(
                    TextValue::borrowed(text),
                )))
            },
        }
    }
}
