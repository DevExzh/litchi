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

use super::{
    EvaluationContext, EvaluationFailure, EvaluationLimits, EvaluationResult, Evaluator,
    ScalarError, TextValue, WorkingValue, ensure_capacity, map_execution_error, parse_error,
};
use crate::codec::formula::{
    expression::ArrayDimensions,
    reference::{Address, Reference},
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
#[allow(dead_code)]
mod geometry;
mod owned;
mod references;
mod scalar;
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
}

#[derive(Debug)]
enum RuntimeValue<'a> {
    Empty,
    Missing,
    Scalar(WorkingValue<'a>),
    Array(RuntimeArrayValue<'a>),
    Areas(RuntimeAreaSet<'a>),
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
            Self::Empty | Self::Missing | Self::Scalar(_) | Self::Areas(_) => None,
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
    matrix: Option<MatrixState<'expr, 'position>>,
    shape_frames: Vec<ShapeFrame<'expr>>,
    shape_values: Vec<Option<Shape>>,
    shape_frame_reservation: Option<Reservation>,
    shape_value_reservation: Option<Reservation>,
    shape_masks: Vec<ShapeMask>,
    shape_mask_reservation: Option<Reservation>,
    demand_cache: Vec<DemandCacheEntry>,
    demand_cache_reservation: Option<Reservation>,
    condition_cache: Vec<ConditionCacheEntry<'expr>>,
    condition_cache_reservation: Option<Reservation>,
    reference_cells_read: usize,
}

#[derive(Clone, Copy)]
enum ValueFrame<'a> {
    Visit(super::Node<'a>),
    VisitArgument(super::Node<'a>),
    Apply(super::Node<'a>),
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
}

struct ConditionCacheEntry<'a> {
    node: usize,
    demand: Shape,
    index: usize,
    value: ConditionCacheValue<'a>,
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
    Exit {
        node: super::Node<'a>,
        children: usize,
        base: Option<Shape>,
    },
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
            matrix: None,
            shape_frames: Vec::new(),
            shape_values: Vec::new(),
            shape_frame_reservation: None,
            shape_value_reservation: None,
            shape_masks: Vec::new(),
            shape_mask_reservation: None,
            demand_cache: Vec::new(),
            demand_cache_reservation: None,
            condition_cache: Vec::new(),
            condition_cache_reservation: None,
            reference_cells_read: 0,
        }
    }

    fn run(&mut self) -> EvaluationResult<RuntimeValue<'expr>> {
        self.run_from(self.expression.root())
    }

    fn run_from(&mut self, root: super::Node<'expr>) -> EvaluationResult<RuntimeValue<'expr>> {
        self.push_frame(ValueFrame::Visit(root))?;
        while let Some(frame) = self.frames.pop() {
            self.scalar.step()?;
            match frame {
                ValueFrame::Visit(node) => self.visit(node)?,
                ValueFrame::VisitArgument(node) => self.visit_argument(node)?,
                ValueFrame::Apply(node) => self.apply(node)?,
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
            Mode::Matrix => Ok(value),
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
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) => Err(EvaluationFailure::Unsupported(
                super::UnsupportedKind::Array,
            )),
        }
    }

    fn project_scalar(
        &mut self,
        value: RuntimeValue<'expr>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
        match value {
            RuntimeValue::Scalar(_) | RuntimeValue::Empty => Ok(value),
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
        if areas.is_list {
            // A reference-list is a distinct formula value.  Implicit
            // intersection is defined for one reference, and selecting the
            // first surviving list entry would silently change the value's
            // type (especially for `NOT`, `XOR`, and scalar functions).
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::Value,
            )));
        }
        if areas.areas.is_empty() {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(ScalarError::Null)));
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
                return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                    ScalarError::NotAvailable,
                )));
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
                    return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                        ScalarError::NotAvailable,
                    )));
                }
                selected = Some((area, row, column));
            }
        }
        let Some((area, row, column)) = selected else {
            return Ok(RuntimeValue::Scalar(WorkingValue::Error(
                ScalarError::NotAvailable,
            )));
        };
        let read = self.read_reference_cell(area.sheet, row, column)?;
        let element = self.read_to_element(read)?;
        self.element_to_runtime(element)
    }

    fn visit(&mut self, node: super::Node<'expr>) -> EvaluationResult<()> {
        if let Some((demand, index)) = self.projection {
            if let Some(value) = self.condition_cache_get(node, demand, index)? {
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
            super::Kind::Parenthesized => self.push_frame(ValueFrame::Visit(node.child(0).ok_or(
                EvaluationFailure::InvalidExpression("parenthesized value node has no child"),
            )?)),
            super::Kind::Prefix(_) | super::Kind::Postfix(_) => {
                let child = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                    "unary value node has no child",
                ))?;
                self.push_frame(ValueFrame::Apply(node))?;
                self.push_frame(ValueFrame::Visit(child))
            },
            super::Kind::Infix(operator) => {
                let left = node.child(0).ok_or(EvaluationFailure::InvalidExpression(
                    "infix value node has no left child",
                ))?;
                let right = node.child(1).ok_or(EvaluationFailure::InvalidExpression(
                    "infix value node has no right child",
                ))?;
                self.push_frame(ValueFrame::Apply(node))?;
                self.push_frame(ValueFrame::Visit(right))?;
                self.push_frame(ValueFrame::Visit(left))?;
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
                let value = self.reference_value(reference)?;
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
        // array-producing branch.  Its arguments normally belong to the
        // eager function schedule, so looking in the cache only from
        // `apply_function` would resolve/read those arguments on every output
        // cell before discovering the cached result.  Check before scheduling
        // the argument visits; a miss follows the ordinary eager path and is
        // inserted by `apply_function` after its first evaluation.
        if self.projection.is_some()
            && (name.eq_ignore_ascii_case("AND") || name.eq_ignore_ascii_case("OR"))
            && self.cacheable_scalar_branch(node)?
        {
            if let Some(value) = self.demand_cache_get(node)? {
                return self.push_value(value);
            }
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
            self.push_frame(ValueFrame::VisitArgument(child))?;
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
        let mut shape = initial;
        for (branch, mask) in [(first, first_mask), (second, second_mask)] {
            let Some(mut mask) = mask else {
                continue;
            };
            if mask.indexes.is_empty() {
                continue;
            }
            let Some(branch) = branch else {
                continue;
            };
            // A selected branch can itself widen a singleton condition axis.
            // Replan that branch at the widened demand so a nested lazy
            // condition is evaluated at every eventual output coordinate.
            // Each iteration is monotonic under the bounded max-shape profile;
            // the node-count cap prevents malformed cyclic growth from
            // turning planning into an unbounded loop.
            let max_iterations = self
                .expression
                .node_count()
                .saturating_add(1)
                .min(self.limits.scalar.max_stack_entries.max(1));
            let mut demand = initial;
            let mut converged = false;
            let mut unknown = false;
            for _ in 0..max_iterations {
                let Some(branch_shape) = self.shape_hint_demand(branch, demand, Some(&mask))?
                else {
                    unknown = true;
                    break;
                };
                let next = broadcast_shape(demand, branch_shape).ok_or(
                    EvaluationFailure::InvalidExpression("incompatible matrix branch shapes"),
                )?;
                if next == demand {
                    shape = broadcast_shape(shape, next).ok_or(
                        EvaluationFailure::InvalidExpression("incompatible matrix branch shapes"),
                    )?;
                    converged = true;
                    break;
                }
                mask = self.expand_shape_mask(mask, demand, next)?;
                if mask.indexes.is_empty() {
                    break;
                }
                demand = next;
            }
            if !converged && !unknown && !mask.indexes.is_empty() {
                return Err(EvaluationFailure::ResourceLimit(self.local_limit(
                    Resource::Objects,
                    u64::try_from(max_iterations.saturating_add(1)).unwrap_or(u64::MAX),
                    max_iterations,
                )));
            }
        }
        Ok(shape)
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
            super::Kind::Function { name }
                if name.eq_ignore_ascii_case("AND") || name.eq_ignore_ascii_case("OR") =>
            {
                let mut cacheable = true;
                for index in 0..node.child_count() {
                    // Only a scalar literal or a reference consumed by the
                    // sequence aggregate is safe to retain across matrix
                    // demands. A direct reference remains projection-shaped:
                    // caching it here would reuse the first cell for every
                    // output position.
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
            RuntimeValue::Scalar(WorkingValue::Text(_))
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

    fn condition_cache_get(
        &mut self,
        node: super::Node<'expr>,
        demand: Shape,
        index: usize,
    ) -> EvaluationResult<Option<RuntimeValue<'expr>>> {
        if self.condition_cache.is_empty() {
            return Ok(None);
        }
        let (position, present) = self.condition_cache_position(node, demand, index)?;
        if !present {
            return Ok(None);
        }
        let value = self.condition_cache[position].value;
        Ok(Some(match value {
            ConditionCacheValue::Empty => RuntimeValue::Empty,
            ConditionCacheValue::Missing => RuntimeValue::Missing,
            ConditionCacheValue::Number(value) => RuntimeValue::Scalar(WorkingValue::Number(value)),
            ConditionCacheValue::Logical(value) => {
                RuntimeValue::Scalar(WorkingValue::Logical(value))
            },
            ConditionCacheValue::Text(value) => {
                RuntimeValue::Scalar(WorkingValue::Text(TextValue::borrowed(value)))
            },
            ConditionCacheValue::Error(error) => RuntimeValue::Scalar(WorkingValue::Error(error)),
        }))
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
            RuntimeValue::Array(_) | RuntimeValue::Areas(_) => return Ok(()),
        };
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
                        | RuntimeElement::Present(WorkingValue::Text(_)) => false,
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
                self.condition_cache_put(condition, effective_shape, index, &value)?;
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
                self.condition_cache_put(condition, effective_shape, index, &evaluated)?;
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
        if let Some(value) = self.condition_cache_get(root, demand, index)? {
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
                let value = RuntimeValue::Scalar(WorkingValue::Error(ScalarError::NotAvailable));
                self.condition_cache_put(root, demand, index, &value)?;
                return Ok(value);
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
        let saved_frames = std::mem::take(&mut self.frames);
        let saved_values = std::mem::take(&mut self.values);
        let saved_frame_reservation = self.frame_reservation.take();
        let saved_value_reservation = self.value_reservation.take();
        let saved_matrix = self.matrix.take();
        let saved_mode = self.mode;
        let saved_position = self.position;
        let saved_projection = self.projection;

        self.mode = Mode::Scalar;
        self.position = position;
        self.projection = Some((demand, index));
        let result = self
            .run_from(probe_root)
            .and_then(|value| self.project_scalar(value));
        let result = match result {
            Ok(value) => self
                .condition_cache_put(root, demand, index, &value)
                .map(|()| value),
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
        let probe_matrix = self.matrix.take();
        drop(probe_frames);
        drop(probe_values);
        drop(probe_frame_reservation);
        drop(probe_value_reservation);
        drop(probe_matrix);

        self.frames = saved_frames;
        self.values = saved_values;
        self.frame_reservation = saved_frame_reservation;
        self.value_reservation = saved_value_reservation;
        self.matrix = saved_matrix;
        self.mode = saved_mode;
        self.position = saved_position;
        self.projection = saved_projection;
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
            let value = RuntimeValue::Scalar(WorkingValue::Error(ScalarError::NotAvailable));
            self.condition_cache_put(root, demand, index, &value)?;
            return Ok(value);
        };
        self.evaluate_scalar_at(root, condition_shape, condition_index)
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
            RuntimeValue::Empty | RuntimeValue::Missing | RuntimeValue::Scalar(_) => {
                Shape::new(1, 1)
            },
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

    fn apply_function(&mut self, node: super::Node<'expr>, name: &str) -> EvaluationResult<()> {
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
            "formula value function arguments",
        )?;
        for _ in 0..count {
            arguments.push(self.pop_value()?);
        }
        arguments.reverse();

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
                WorkingValue::Logical(_) | WorkingValue::Text(_) => {},
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
                RuntimeValue::Array(_) | RuntimeValue::Areas(_) => RuntimeElement::Missing,
            })
    }

    fn map_function(
        &mut self,
        node: super::Node<'expr>,
        name: &str,
        arguments: Vec<RuntimeValue<'expr>>,
    ) -> EvaluationResult<RuntimeValue<'expr>> {
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
        let mut scalar_arguments = Vec::new();
        let mut reservation = None;
        ensure_capacity(
            &mut scalar_arguments,
            &mut reservation,
            arguments.len(),
            self.limits.scalar.max_stack_entries,
            self.execution,
            &self.storage_budget,
            "formula value scalar arguments",
        )?;
        for (index, argument) in arguments.into_iter().enumerate() {
            let slot = self.value_to_slot(argument)?;
            scalar_arguments.push(scalar::argument(name, index, slot));
        }
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
