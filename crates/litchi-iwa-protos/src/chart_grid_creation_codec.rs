//! Bounded Buffa authoring for a fresh inline chart grid.
//!
//! `ChartGridArchive` is one of the few IWA messages that is useful to
//! construct from an archive-free semantic value.  The native message has
//! repeated row, value, and identifier-map fields, so constructing its
//! generated representation would make the peak allocation proportional to
//! the whole grid before the output buffer is even available.  This codec
//! keeps the request borrowed and emits the repeated framing directly into
//! one caller-owned output allocation.
//!
//! The only generated value used here is the private Buffa `GridValueView`.
//! It is a scalar `ViewEncode` leaf; rows, values, labels, and identifier-map
//! entries never enter generated repeated vectors.  The writer's byte order
//! and field defaults intentionally match the former Prost source builder.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The request, bounded plan, and streaming wire writer are kept together."
)]

use std::{error::Error, fmt, mem::size_of};

use buffa::ViewEncode as _;

use crate::buffa_chart_data_generated::LitchiIwaProjection as projection;

const DEFAULT_OUTPUT_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_FIELDS: usize = 4 * 1024 * 1024;
const DEFAULT_WORK_BYTES: usize = 512 * 1024 * 1024;
const DEFAULT_RETAINED_BYTES: usize = 512 * 1024 * 1024;
const DEFAULT_SCRATCH_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_CELLS: usize = 1_000_000;
const DEFAULT_LABELS: usize = 1_000_000;
const DEFAULT_TEXT_BYTES: usize = 64 * 1024 * 1024;
const OUTPUT_ALLOCATIONS: usize = 1;

const GRID_VALUE_FIELD: u32 = 1;
const GRID_ROW_NAME_FIELD: u32 = 1;
const GRID_COLUMN_NAME_FIELD: u32 = 2;
const GRID_ROW_FIELD: u32 = 3;
const GRID_ID_MAP_FIELD: u32 = 4;
const ID_MAP_ROW_FIELD: u32 = 1;
const ID_MAP_COLUMN_FIELD: u32 = 2;
const ENTRY_UNIQUE_ID_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 2;
const WIRE_VARINT: u8 = 0;
const WIRE_LENGTH_DELIMITED: u8 = 2;
const MAX_DEPTH: u32 = 3;
const UUID_BYTES: usize = 36;
const UUID_SUFFIX_MASK: u64 = 0x0000_ffff_ffff_ffff;
const UUID_PREFIX: &[u8] = b"00000000-0000-4000-8000-";

/// A borrowed semantic request for one freshly authored rectangular grid.
///
/// Labels and values remain owned by the caller for the complete preparation
/// and execution lifetime.  The codec never clones them and never exposes a
/// native row, column, or object identifier.
#[derive(Debug, Clone, Copy)]
pub struct ChartGridCreationRequest<'source> {
    row_labels: &'source [String],
    column_labels: &'source [String],
    values: &'source [Vec<Option<f64>>],
    seed: u64,
}

impl<'source> ChartGridCreationRequest<'source> {
    /// Construct a borrowed chart-grid authoring request.
    #[must_use]
    pub const fn new(
        row_labels: &'source [String],
        column_labels: &'source [String],
        values: &'source [Vec<Option<f64>>],
        seed: u64,
    ) -> Self {
        Self {
            row_labels,
            column_labels,
            values,
            seed,
        }
    }

    /// Borrow row labels in native order.
    #[must_use]
    pub const fn row_labels(self) -> &'source [String] {
        self.row_labels
    }

    /// Borrow column labels in native order.
    #[must_use]
    pub const fn column_labels(self) -> &'source [String] {
        self.column_labels
    }

    /// Borrow row-major optional numeric values.
    #[must_use]
    pub const fn values(self) -> &'source [Vec<Option<f64>>] {
        self.values
    }

    /// Return the deterministic identifier seed.
    #[must_use]
    pub const fn seed(self) -> u64 {
        self.seed
    }
}

/// Short alias for callers that use the grid terminology directly.
pub type GridCreationRequest<'source> = ChartGridCreationRequest<'source>;

/// Finite ceilings for one complete chart-grid creation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncodeOptions {
    /// Maximum encoded message bytes.
    pub max_output_bytes: usize,
    /// Maximum emitted protobuf fields.
    pub max_fields: usize,
    /// Maximum aggregate planning, Buffa, and emission work.
    pub max_work_bytes: usize,
    /// Maximum logical allocations.
    pub max_allocations: usize,
    /// Maximum request plus output bytes retained by the operation.
    pub max_retained_bytes: usize,
    /// Maximum output staging bytes.
    pub max_scratch_bytes: usize,
    /// Maximum numeric cells.
    pub max_cells: usize,
    /// Maximum row plus column labels.
    pub max_labels: usize,
    /// Maximum UTF-8 label bytes.
    pub max_text_bytes: usize,
    /// Maximum nested message depth.
    pub max_depth: u32,
}

impl EncodeOptions {
    /// Build an explicit finite policy.
    #[must_use]
    pub const fn new(
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_allocations: usize,
        max_retained_bytes: usize,
        max_scratch_bytes: usize,
        max_cells: usize,
        max_labels: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_output_bytes,
            max_fields,
            max_work_bytes,
            max_allocations,
            max_retained_bytes,
            max_scratch_bytes,
            max_cells,
            max_labels,
            max_text_bytes,
            max_depth: MAX_DEPTH,
        }
    }

    /// Build a conservative finite policy for one request.
    #[must_use]
    pub const fn for_request(_request: &ChartGridCreationRequest<'_>) -> Self {
        Self::new(
            DEFAULT_OUTPUT_BYTES,
            DEFAULT_FIELDS,
            DEFAULT_WORK_BYTES,
            OUTPUT_ALLOCATIONS,
            DEFAULT_RETAINED_BYTES,
            DEFAULT_SCRATCH_BYTES,
            DEFAULT_CELLS,
            DEFAULT_LABELS,
            DEFAULT_TEXT_BYTES,
        )
    }

    /// Return the output-byte ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Return the emitted-field ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Return the work ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Return the allocation ceiling.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    /// Return the retained-byte ceiling.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    /// Return the scratch-byte ceiling.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    /// Return the cell ceiling.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.max_cells
    }

    /// Return the label-count ceiling.
    #[must_use]
    pub const fn max_labels(self) -> usize {
        self.max_labels
    }

    /// Return the label-count ceiling using the decoder terminology.
    #[must_use]
    pub const fn max_label_count(self) -> usize {
        self.max_labels()
    }

    /// Return the aggregate text ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Return the nesting ceiling.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value;
        self
    }

    /// Replace the output-byte ceiling with concise execution terminology.
    #[must_use]
    pub const fn with_output_bytes(self, value: usize) -> Self {
        self.with_max_output_bytes(value)
    }

    /// Replace the field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, value: usize) -> Self {
        self.max_fields = value;
        self
    }

    /// Replace the field ceiling with concise execution terminology.
    #[must_use]
    pub const fn with_fields(self, value: usize) -> Self {
        self.with_max_fields(value)
    }

    /// Replace the work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, value: usize) -> Self {
        self.max_work_bytes = value;
        self
    }

    /// Replace the work ceiling with concise execution terminology.
    #[must_use]
    pub const fn with_work_bytes(self, value: usize) -> Self {
        self.with_max_work_bytes(value)
    }

    /// Replace the allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, value: usize) -> Self {
        self.max_allocations = value;
        self
    }

    /// Replace the allocation ceiling with concise execution terminology.
    #[must_use]
    pub const fn with_allocations(self, value: usize) -> Self {
        self.with_max_allocations(value)
    }

    /// Replace the retained-byte ceiling.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, value: usize) -> Self {
        self.max_retained_bytes = value;
        self
    }

    /// Replace the retained-byte ceiling with concise execution terminology.
    #[must_use]
    pub const fn with_retained_bytes(self, value: usize) -> Self {
        self.with_max_retained_bytes(value)
    }

    /// Replace the scratch-byte ceiling.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, value: usize) -> Self {
        self.max_scratch_bytes = value;
        self
    }

    /// Replace the scratch-byte ceiling with concise execution terminology.
    #[must_use]
    pub const fn with_scratch_bytes(self, value: usize) -> Self {
        self.with_max_scratch_bytes(value)
    }

    /// Replace the cell ceiling.
    #[must_use]
    pub const fn with_max_cells(mut self, value: usize) -> Self {
        self.max_cells = value;
        self
    }

    /// Replace the label-count ceiling.
    #[must_use]
    pub const fn with_max_labels(mut self, value: usize) -> Self {
        self.max_labels = value;
        self
    }

    /// Replace the label-count ceiling using the decoder terminology.
    #[must_use]
    pub const fn with_max_label_count(self, value: usize) -> Self {
        self.with_max_labels(value)
    }

    /// Replace the aggregate text ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, value: usize) -> Self {
        self.max_text_bytes = value;
        self
    }

    /// Replace the nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }

    /// Convert this preparation policy to execution ceilings.
    #[must_use]
    pub const fn execution_limits(self) -> ExecutionLimits {
        ExecutionLimits {
            output_bytes: self.max_output_bytes,
            fields: self.max_fields,
            work_bytes: self.max_work_bytes,
            allocations: self.max_allocations,
            retained_bytes: self.max_retained_bytes,
            scratch_bytes: self.max_scratch_bytes,
            cells: self.max_cells,
            labels: self.max_labels,
            text_bytes: self.max_text_bytes,
            max_depth: self.max_depth,
        }
    }
}

/// Exact resources required by one prepared chart-grid creation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExecutionRequirements {
    /// Encoded `ChartGridArchive` bytes.
    pub output_bytes: usize,
    /// Emitted protobuf field count.
    pub fields: usize,
    /// Aggregate checked planning, Buffa, and streaming work.
    pub work_bytes: usize,
    /// Maximum nested message depth.
    pub max_depth: u32,
    /// Logical allocations, currently one output vector.
    pub allocations: usize,
    /// Borrowed request representation plus candidate bytes.
    pub retained_bytes: usize,
    /// Candidate output staging bytes.
    pub scratch_bytes: usize,
    /// Numeric cell count.
    pub cells: usize,
    /// Row plus column label count.
    pub labels: usize,
    /// Aggregate UTF-8 label bytes.
    pub text_bytes: usize,
}

impl ExecutionRequirements {
    /// Return exact execution ceilings.
    #[must_use]
    pub const fn exact(self) -> ExecutionLimits {
        ExecutionLimits {
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
            cells: self.cells,
            labels: self.labels,
            text_bytes: self.text_bytes,
            max_depth: self.max_depth,
        }
    }

    /// Compatibility alias for exact limits.
    #[must_use]
    pub const fn exact_limits(self) -> ExecutionLimits {
        self.exact()
    }

    /// Return the encoded byte count.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Return the field count.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Return the aggregate work count.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Return the logical allocation count.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Return retained bytes.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Return scratch bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Return numeric cell count.
    #[must_use]
    pub const fn cells(self) -> usize {
        self.cells
    }

    /// Return row plus column labels.
    #[must_use]
    pub const fn labels(self) -> usize {
        self.labels
    }

    /// Compatibility spelling for label count.
    #[must_use]
    pub const fn label_count(self) -> usize {
        self.labels()
    }

    /// Return aggregate label text bytes.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Compatibility spelling for aggregate text.
    #[must_use]
    pub const fn text(self) -> usize {
        self.text_bytes()
    }
}

/// Caller-provided ceilings for the execution phase.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExecutionLimits {
    /// Maximum encoded bytes.
    pub output_bytes: usize,
    /// Maximum fields.
    pub fields: usize,
    /// Maximum aggregate work.
    pub work_bytes: usize,
    /// Maximum logical allocations.
    pub allocations: usize,
    /// Maximum retained bytes.
    pub retained_bytes: usize,
    /// Maximum scratch bytes.
    pub scratch_bytes: usize,
    /// Maximum cells.
    pub cells: usize,
    /// Maximum labels.
    pub labels: usize,
    /// Maximum text bytes.
    pub text_bytes: usize,
    /// Maximum nesting depth.
    pub max_depth: u32,
}

impl ExecutionLimits {
    /// Build exact limits from requirements.
    #[must_use]
    pub const fn exact(requirements: ExecutionRequirements) -> Self {
        requirements.exact()
    }

    /// Replace output ceiling.
    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    /// Replace output ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_output_bytes(self, value: usize) -> Self {
        self.with_output_bytes(value)
    }

    /// Replace field ceiling.
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }

    /// Replace field ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_fields(self, value: usize) -> Self {
        self.with_fields(value)
    }

    /// Replace work ceiling.
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }

    /// Replace work ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_work_bytes(self, value: usize) -> Self {
        self.with_work_bytes(value)
    }

    /// Replace allocation ceiling.
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }

    /// Replace allocation ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_allocations(self, value: usize) -> Self {
        self.with_allocations(value)
    }

    /// Replace retained-byte ceiling.
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }

    /// Replace retained-byte ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_retained_bytes(self, value: usize) -> Self {
        self.with_retained_bytes(value)
    }

    /// Replace scratch-byte ceiling.
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }

    /// Replace scratch-byte ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_scratch_bytes(self, value: usize) -> Self {
        self.with_scratch_bytes(value)
    }

    /// Replace cell ceiling.
    #[must_use]
    pub const fn with_cells(mut self, value: usize) -> Self {
        self.cells = value;
        self
    }

    /// Replace cell ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_cells(self, value: usize) -> Self {
        self.with_cells(value)
    }

    /// Replace label ceiling.
    #[must_use]
    pub const fn with_labels(mut self, value: usize) -> Self {
        self.labels = value;
        self
    }

    /// Replace label ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_labels(self, value: usize) -> Self {
        self.with_labels(value)
    }

    /// Replace label ceiling using decoder terminology.
    #[must_use]
    pub const fn with_max_label_count(self, value: usize) -> Self {
        self.with_max_labels(value)
    }

    /// Replace text ceiling.
    #[must_use]
    pub const fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value;
        self
    }

    /// Replace text ceiling using the options spelling.
    #[must_use]
    pub const fn with_max_text_bytes(self, value: usize) -> Self {
        self.with_text_bytes(value)
    }

    /// Replace nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }
}

/// One resource axis refused by preparation or execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeLimit {
    /// Encoded output bytes exceeded their ceiling.
    Output { observed: usize, maximum: usize },
    /// Emitted fields exceeded their ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate work exceeded its ceiling.
    Work { observed: usize, maximum: usize },
    /// Logical allocation count exceeded its ceiling.
    Allocations { observed: usize, maximum: usize },
    /// Retained bytes exceeded their ceiling.
    Retained { observed: usize, maximum: usize },
    /// Scratch bytes exceeded their ceiling.
    Scratch { observed: usize, maximum: usize },
    /// Numeric cell count exceeded its ceiling.
    Cells { observed: usize, maximum: usize },
    /// Label count exceeded its ceiling.
    Labels { observed: usize, maximum: usize },
    /// UTF-8 label bytes exceeded their ceiling.
    Text { observed: usize, maximum: usize },
    /// Message nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Shape or scalar failure found before output allocation.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum InvalidInput {
    /// No row labels were supplied.
    EmptyRows,
    /// No column labels were supplied.
    EmptyColumns,
    /// Value-row count differs from row-label count.
    RowCountMismatch { expected: usize, actual: usize },
    /// One value row has the wrong width.
    ColumnCountMismatch {
        row: usize,
        expected: usize,
        actual: usize,
    },
    /// One cell contains NaN or infinity.
    NonFiniteNumeric { row: usize, column: usize },
    /// A row or column index cannot be represented by the native uint32 map.
    IndexTooLarge { axis: &'static str, index: usize },
}

/// Failure from bounded chart-grid authoring.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncodeError {
    /// The borrowed request is not a valid rectangular finite grid.
    InvalidInput(InvalidInput),
    /// A caller-selected resource ceiling was exceeded.
    Limit(EncodeLimit),
    /// The one output allocation could not be reserved.
    Allocation { amount: usize },
    /// Buffa rejected a scalar leaf encoding.
    Buffa(buffa::EncodeError),
    /// The static plan and streaming emitter disagreed.
    Verification,
}

impl EncodeError {
    /// Return the resource limit, if this is a limit failure.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<EncodeLimit> {
        match self {
            Self::Limit(limit) => Some(*limit),
            _ => None,
        }
    }

    /// Return the output allocation amount, if present.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self {
            Self::Allocation { amount } => Some(*amount),
            _ => None,
        }
    }
}

impl fmt::Display for EncodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidInput(input) => write!(formatter, "invalid chart-grid input: {input:?}"),
            Self::Limit(EncodeLimit::Output { observed, maximum }) => write!(
                formatter,
                "chart-grid output has {observed} bytes; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "chart-grid emits {observed} fields; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "chart-grid creation requires {observed} work bytes; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "chart-grid creation requires {observed} allocations; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "chart-grid creation retains {observed} bytes; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "chart-grid creation stages {observed} bytes; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Cells { observed, maximum }) => write!(
                formatter,
                "chart-grid contains {observed} cells; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Labels { observed, maximum }) => write!(
                formatter,
                "chart-grid contains {observed} labels; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Text { observed, maximum }) => write!(
                formatter,
                "chart-grid labels contain {observed} UTF-8 bytes; maximum is {maximum}"
            ),
            Self::Limit(EncodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "chart-grid nesting {observed} exceeds maximum {maximum}"
            ),
            Self::Allocation { amount } => write!(
                formatter,
                "cannot allocate chart-grid output for {amount} bytes"
            ),
            Self::Buffa(error) => error.fmt(formatter),
            Self::Verification => formatter.write_str("chart-grid creation verification failed"),
        }
    }
}

impl Error for EncodeError {}

impl From<buffa::EncodeError> for EncodeError {
    fn from(error: buffa::EncodeError) -> Self {
        Self::Buffa(error)
    }
}

/// Encoded chart-grid bytes plus exact resource evidence.
#[derive(Debug, PartialEq, Eq)]
pub struct EncodeOutput {
    bytes: Vec<u8>,
    report: ExecutionRequirements,
}

impl EncodeOutput {
    /// Borrow the encoded `ChartGridArchive` payload.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Consume the result and return encoded bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// Return exact creation requirements.
    #[must_use]
    pub const fn report(&self) -> ExecutionRequirements {
        self.report
    }
}

/// A request whose exact output plan has been prepared without allocation.
#[derive(Clone, Copy)]
pub struct PreparedChartGridCreation<'source> {
    request: ChartGridCreationRequest<'source>,
    requirements: ExecutionRequirements,
}

impl fmt::Debug for PreparedChartGridCreation<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedChartGridCreation")
            .field("rows", &self.request.row_labels.len())
            .field("columns", &self.request.column_labels.len())
            .field("seed", &self.request.seed)
            .field("requirements", &self.requirements)
            .finish()
    }
}

impl<'source> PreparedChartGridCreation<'source> {
    /// Return the exact aggregate execution requirements.
    #[must_use]
    pub const fn execution_requirements(self) -> ExecutionRequirements {
        self.requirements
    }

    /// Encode after checking every caller-provided execution ceiling.
    pub fn execute(self, limits: ExecutionLimits) -> Result<EncodeOutput, EncodeError> {
        check_execution_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| EncodeError::Allocation {
                amount: self.requirements.output_bytes,
            })?;
        if output.capacity() < self.requirements.output_bytes {
            return Err(EncodeError::Allocation {
                amount: self.requirements.output_bytes,
            });
        }

        let mut scalar_cache = buffa::SizeCache::new();
        emit_request(
            self.request,
            self.requirements,
            &mut output,
            &mut scalar_cache,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(EncodeError::Verification);
        }
        Ok(EncodeOutput {
            bytes: output,
            report: self.requirements,
        })
    }

    /// Execute with the exact prepared ceilings.
    pub fn into_output(self) -> Result<EncodeOutput, EncodeError> {
        self.execute(self.requirements.exact())
    }
}

/// Prepare one borrowed grid without allocating output bytes.
pub fn prepare_chart_grid_creation<'source>(
    request: ChartGridCreationRequest<'source>,
    options: EncodeOptions,
) -> Result<PreparedChartGridCreation<'source>, EncodeError> {
    let plan = plan_request(request, options)?;
    check_options(plan, options)?;
    Ok(PreparedChartGridCreation {
        request,
        requirements: plan,
    })
}

/// Encode one borrowed chart grid using the exact finite policy supplied.
pub fn encode_chart_grid(
    request: ChartGridCreationRequest<'_>,
    options: EncodeOptions,
) -> Result<EncodeOutput, EncodeError> {
    prepare_chart_grid_creation(request, options)?.into_output()
}

fn plan_request(
    request: ChartGridCreationRequest<'_>,
    options: EncodeOptions,
) -> Result<ExecutionRequirements, EncodeError> {
    if options.max_depth < MAX_DEPTH {
        return Err(EncodeError::Limit(EncodeLimit::Nesting {
            observed: MAX_DEPTH,
            maximum: options.max_depth,
        }));
    }

    let row_count = request.row_labels.len();
    let column_count = request.column_labels.len();
    if row_count == 0 {
        return Err(EncodeError::InvalidInput(InvalidInput::EmptyRows));
    }
    if column_count == 0 {
        return Err(EncodeError::InvalidInput(InvalidInput::EmptyColumns));
    }
    if request.values.len() != row_count {
        return Err(EncodeError::InvalidInput(InvalidInput::RowCountMismatch {
            expected: row_count,
            actual: request.values.len(),
        }));
    }
    check_index(row_count - 1, "row")?;
    check_index(column_count - 1, "column")?;

    let labels = row_count.checked_add(column_count).ok_or_else(|| {
        overflow_limit(EncodeLimit::Labels {
            observed: usize::MAX,
            maximum: options.max_labels,
        })
    })?;
    let cells = row_count.checked_mul(column_count).ok_or_else(|| {
        overflow_limit(EncodeLimit::Cells {
            observed: usize::MAX,
            maximum: options.max_cells,
        })
    })?;
    if cells > options.max_cells {
        return Err(EncodeError::Limit(EncodeLimit::Cells {
            observed: cells,
            maximum: options.max_cells,
        }));
    }
    if labels > options.max_labels {
        return Err(EncodeError::Limit(EncodeLimit::Labels {
            observed: labels,
            maximum: options.max_labels,
        }));
    }
    if OUTPUT_ALLOCATIONS > options.max_allocations {
        return Err(EncodeError::Limit(EncodeLimit::Allocations {
            observed: OUTPUT_ALLOCATIONS,
            maximum: options.max_allocations,
        }));
    }

    // Reject zero or tiny field/work/output policies from dimensions alone.
    // This keeps a hostile request from forcing a complete label/value/map
    // walk merely to discover that it could never be admitted.
    let minimum_fields = labels
        .checked_mul(4)
        .and_then(|value| value.checked_add(row_count))
        .and_then(|value| value.checked_add(cells))
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    if minimum_fields > options.max_fields {
        return Err(EncodeError::Limit(EncodeLimit::Fields {
            observed: minimum_fields,
            maximum: options.max_fields,
        }));
    }
    let minimum_output = minimum_output_len(row_count, column_count)?;
    if minimum_output > options.max_output_bytes {
        return Err(EncodeError::Limit(EncodeLimit::Output {
            observed: minimum_output,
            maximum: options.max_output_bytes,
        }));
    }
    if minimum_output > options.max_scratch_bytes {
        return Err(EncodeError::Limit(EncodeLimit::Scratch {
            observed: minimum_output,
            maximum: options.max_scratch_bytes,
        }));
    }
    let minimum_request_bytes = size_of::<ChartGridCreationRequest<'_>>()
        .checked_add(labels.checked_mul(size_of::<String>()).ok_or_else(|| {
            overflow_limit(EncodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?)
        .and_then(|bytes| bytes.checked_add(row_count.checked_mul(size_of::<Vec<Option<f64>>>())?))
        .and_then(|bytes| bytes.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    if minimum_request_bytes
        .checked_add(minimum_output)
        .is_none_or(|bytes| bytes > options.max_retained_bytes)
    {
        let observed = minimum_request_bytes.saturating_add(minimum_output);
        return Err(EncodeError::Limit(EncodeLimit::Retained {
            observed,
            maximum: options.max_retained_bytes,
        }));
    }
    let minimum_work = minimum_request_bytes
        .checked_add(minimum_output)
        .and_then(|work| work.checked_add(minimum_fields))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    if minimum_work > options.max_work_bytes {
        return Err(EncodeError::Limit(EncodeLimit::Work {
            observed: minimum_work,
            maximum: options.max_work_bytes,
        }));
    }

    let mut text_bytes = 0usize;
    let mut output_bytes = 0usize;
    let mut fields = 0usize;
    let mut plan_work = 0usize;

    for label in request.row_labels {
        text_bytes = checked_add(
            text_bytes,
            label.len(),
            EncodeLimit::Text {
                observed: usize::MAX,
                maximum: options.max_text_bytes,
            },
        )?;
        output_bytes = checked_add(
            output_bytes,
            string_field_len(GRID_ROW_NAME_FIELD, label.len())?,
            EncodeLimit::Output {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            },
        )?;
        fields = checked_add(
            fields,
            1,
            EncodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            },
        )?;
        check_partial_limits(
            text_bytes,
            options.max_text_bytes,
            EncodeLimit::Text {
                observed: text_bytes,
                maximum: options.max_text_bytes,
            },
        )?;
        check_partial_limits(
            output_bytes,
            options.max_output_bytes,
            EncodeLimit::Output {
                observed: output_bytes,
                maximum: options.max_output_bytes,
            },
        )?;
        check_partial_limits(
            fields,
            options.max_fields,
            EncodeLimit::Fields {
                observed: fields,
                maximum: options.max_fields,
            },
        )?;
        charge_work(
            &mut plan_work,
            label
                .len()
                .checked_add(string_field_len(GRID_ROW_NAME_FIELD, label.len())?)
                .ok_or_else(|| {
                    overflow_limit(EncodeLimit::Work {
                        observed: usize::MAX,
                        maximum: options.max_work_bytes,
                    })
                })?,
            options.max_work_bytes,
        )?;
    }
    for label in request.column_labels {
        text_bytes = checked_add(
            text_bytes,
            label.len(),
            EncodeLimit::Text {
                observed: usize::MAX,
                maximum: options.max_text_bytes,
            },
        )?;
        output_bytes = checked_add(
            output_bytes,
            string_field_len(GRID_COLUMN_NAME_FIELD, label.len())?,
            EncodeLimit::Output {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            },
        )?;
        fields = checked_add(
            fields,
            1,
            EncodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            },
        )?;
        check_partial_limits(
            text_bytes,
            options.max_text_bytes,
            EncodeLimit::Text {
                observed: text_bytes,
                maximum: options.max_text_bytes,
            },
        )?;
        check_partial_limits(
            output_bytes,
            options.max_output_bytes,
            EncodeLimit::Output {
                observed: output_bytes,
                maximum: options.max_output_bytes,
            },
        )?;
        check_partial_limits(
            fields,
            options.max_fields,
            EncodeLimit::Fields {
                observed: fields,
                maximum: options.max_fields,
            },
        )?;
        charge_work(
            &mut plan_work,
            label
                .len()
                .checked_add(string_field_len(GRID_COLUMN_NAME_FIELD, label.len())?)
                .ok_or_else(|| {
                    overflow_limit(EncodeLimit::Work {
                        observed: usize::MAX,
                        maximum: options.max_work_bytes,
                    })
                })?,
            options.max_work_bytes,
        )?;
    }

    for (row_index, row) in request.values.iter().enumerate() {
        if row.len() != column_count {
            return Err(EncodeError::InvalidInput(
                InvalidInput::ColumnCountMismatch {
                    row: row_index,
                    expected: column_count,
                    actual: row.len(),
                },
            ));
        }
        let mut row_payload = 0usize;
        for (column_index, value) in row.iter().enumerate() {
            if value.is_some_and(|number| !number.is_finite()) {
                return Err(EncodeError::InvalidInput(InvalidInput::NonFiniteNumeric {
                    row: row_index,
                    column: column_index,
                }));
            }
            let scalar_bytes = scalar_value_encoded_len(*value)?;
            let value_bytes = message_field_len(GRID_VALUE_FIELD, scalar_bytes)?;
            row_payload = checked_add(
                row_payload,
                value_bytes,
                EncodeLimit::Output {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                },
            )?;
            fields = checked_add(
                fields,
                1,
                EncodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                },
            )?;
            if value.is_some() {
                fields = checked_add(
                    fields,
                    1,
                    EncodeLimit::Fields {
                        observed: usize::MAX,
                        maximum: options.max_fields,
                    },
                )?;
            }
            charge_work(
                &mut plan_work,
                size_of::<Option<f64>>()
                    .checked_add(scalar_bytes)
                    .and_then(|work| work.checked_add(value_bytes))
                    .ok_or_else(|| {
                        overflow_limit(EncodeLimit::Work {
                            observed: usize::MAX,
                            maximum: options.max_work_bytes,
                        })
                    })?,
                options.max_work_bytes,
            )?;
            check_partial_limits(
                output_bytes,
                options.max_output_bytes,
                EncodeLimit::Output {
                    observed: output_bytes,
                    maximum: options.max_output_bytes,
                },
            )?;
            check_partial_limits(
                fields,
                options.max_fields,
                EncodeLimit::Fields {
                    observed: fields,
                    maximum: options.max_fields,
                },
            )?;
        }
        output_bytes = checked_add(
            output_bytes,
            message_field_len(GRID_ROW_FIELD, row_payload)?,
            EncodeLimit::Output {
                observed: usize::MAX,
                maximum: options.max_output_bytes,
            },
        )?;
        fields = checked_add(
            fields,
            1,
            EncodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            },
        )?;
        charge_work(
            &mut plan_work,
            row_payload
                .checked_add(message_field_len(GRID_ROW_FIELD, row_payload)?)
                .ok_or_else(|| {
                    overflow_limit(EncodeLimit::Work {
                        observed: usize::MAX,
                        maximum: options.max_work_bytes,
                    })
                })?,
            options.max_work_bytes,
        )?;
        check_partial_limits(
            output_bytes,
            options.max_output_bytes,
            EncodeLimit::Output {
                observed: output_bytes,
                maximum: options.max_output_bytes,
            },
        )?;
        check_partial_limits(
            fields,
            options.max_fields,
            EncodeLimit::Fields {
                observed: fields,
                maximum: options.max_fields,
            },
        )?;
    }

    let mut map_payload = 0usize;
    for (axis_field, count) in [
        (ID_MAP_ROW_FIELD, row_count),
        (ID_MAP_COLUMN_FIELD, column_count),
    ] {
        for index in 0..count {
            let index = u32::try_from(index).map_err(|_| {
                EncodeError::InvalidInput(InvalidInput::IndexTooLarge {
                    axis: if axis_field == ID_MAP_ROW_FIELD {
                        "row"
                    } else {
                        "column"
                    },
                    index,
                })
            })?;
            let entry_payload = entry_payload_len(index as u64)?;
            map_payload = checked_add(
                map_payload,
                message_field_len(axis_field, entry_payload)?,
                EncodeLimit::Output {
                    observed: usize::MAX,
                    maximum: options.max_output_bytes,
                },
            )?;
            fields = checked_add(
                fields,
                1,
                EncodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                },
            )?;
            fields = checked_add(
                fields,
                2,
                EncodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                },
            )?;
            charge_work(
                &mut plan_work,
                entry_payload
                    .checked_add(message_field_len(axis_field, entry_payload)?)
                    .ok_or_else(|| {
                        overflow_limit(EncodeLimit::Work {
                            observed: usize::MAX,
                            maximum: options.max_work_bytes,
                        })
                    })?,
                options.max_work_bytes,
            )?;
            check_partial_limits(
                fields,
                options.max_fields,
                EncodeLimit::Fields {
                    observed: fields,
                    maximum: options.max_fields,
                },
            )?;
        }
    }
    output_bytes = checked_add(
        output_bytes,
        message_field_len(GRID_ID_MAP_FIELD, map_payload)?,
        EncodeLimit::Output {
            observed: usize::MAX,
            maximum: options.max_output_bytes,
        },
    )?;
    fields = checked_add(
        fields,
        1,
        EncodeLimit::Fields {
            observed: usize::MAX,
            maximum: options.max_fields,
        },
    )?;
    charge_work(
        &mut plan_work,
        map_payload
            .checked_add(message_field_len(GRID_ID_MAP_FIELD, map_payload)?)
            .ok_or_else(|| {
                overflow_limit(EncodeLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?,
        options.max_work_bytes,
    )?;
    check_partial_limits(
        output_bytes,
        options.max_output_bytes,
        EncodeLimit::Output {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        },
    )?;
    check_partial_limits(
        fields,
        options.max_fields,
        EncodeLimit::Fields {
            observed: fields,
            maximum: options.max_fields,
        },
    )?;

    let request_bytes = size_of::<ChartGridCreationRequest<'_>>()
        .checked_add(labels.checked_mul(size_of::<String>()).ok_or_else(|| {
            overflow_limit(EncodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?)
        .and_then(|bytes| bytes.checked_add(row_count.checked_mul(size_of::<Vec<Option<f64>>>())?))
        .and_then(|bytes| bytes.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .and_then(|bytes| bytes.checked_add(text_bytes))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    let retained_bytes = checked_add(
        request_bytes,
        output_bytes,
        EncodeLimit::Retained {
            observed: usize::MAX,
            maximum: options.max_retained_bytes,
        },
    )?;
    let scratch_bytes = output_bytes;

    // Preparation walks labels, values, and map entries once. Execution then
    // measures every row, emits every scalar leaf, and walks the map again;
    // charge those traversals explicitly before the output allocation.
    let execution_work = text_bytes
        .checked_add(labels)
        .and_then(|work| work.checked_add(labels))
        .and_then(|work| work.checked_add(labels.checked_mul(UUID_BYTES + 2)?))
        .and_then(|work| work.checked_add(cells.checked_mul(18 + 9 + 1)?))
        .and_then(|work| work.checked_add(row_count.checked_mul(2)?))
        .and_then(|work| work.checked_add(output_bytes))
        .and_then(|work| work.checked_add(fields))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let work_bytes = plan_work.checked_add(execution_work).ok_or_else(|| {
        overflow_limit(EncodeLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        })
    })?;
    if work_bytes > options.max_work_bytes {
        return Err(EncodeError::Limit(EncodeLimit::Work {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }

    Ok(ExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: MAX_DEPTH,
        allocations: OUTPUT_ALLOCATIONS,
        retained_bytes,
        scratch_bytes,
        cells,
        labels,
        text_bytes,
    })
}

fn check_index(index: usize, axis: &'static str) -> Result<(), EncodeError> {
    if u32::try_from(index).is_err() {
        return Err(EncodeError::InvalidInput(InvalidInput::IndexTooLarge {
            axis,
            index,
        }));
    }
    Ok(())
}

fn minimum_output_len(row_count: usize, column_count: usize) -> Result<usize, EncodeError> {
    let labels = row_count.checked_add(column_count).ok_or_else(|| {
        overflow_limit(EncodeLimit::Output {
            observed: usize::MAX,
            maximum: usize::MAX,
        })
    })?;
    let empty_label_bytes = labels
        .checked_mul(message_field_len(GRID_ROW_NAME_FIELD, 0)?)
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
    let empty_value_bytes = message_field_len(GRID_VALUE_FIELD, 0)?;
    let row_payload = column_count.checked_mul(empty_value_bytes).ok_or_else(|| {
        overflow_limit(EncodeLimit::Output {
            observed: usize::MAX,
            maximum: usize::MAX,
        })
    })?;
    let rows_bytes = row_count
        .checked_mul(message_field_len(GRID_ROW_FIELD, row_payload)?)
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
    let entry_bytes = message_field_len(ENTRY_UNIQUE_ID_FIELD, UUID_BYTES)?
        .checked_add(varint_field_len(ENTRY_INDEX_FIELD, 0))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
    let row_entries = row_count
        .checked_mul(message_field_len(ID_MAP_ROW_FIELD, entry_bytes)?)
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
    let column_entries = column_count
        .checked_mul(message_field_len(ID_MAP_COLUMN_FIELD, entry_bytes)?)
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
    let map_payload = row_entries.checked_add(column_entries).ok_or_else(|| {
        overflow_limit(EncodeLimit::Output {
            observed: usize::MAX,
            maximum: usize::MAX,
        })
    })?;
    let map_bytes = message_field_len(GRID_ID_MAP_FIELD, map_payload)?;
    empty_label_bytes
        .checked_add(rows_bytes)
        .and_then(|bytes| bytes.checked_add(map_bytes))
        .ok_or_else(|| {
            overflow_limit(EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })
}

fn check_partial_limits(
    observed: usize,
    maximum: usize,
    limit: EncodeLimit,
) -> Result<(), EncodeError> {
    if observed > maximum {
        return Err(EncodeError::Limit(limit));
    }
    Ok(())
}

fn charge_work(work: &mut usize, amount: usize, maximum: usize) -> Result<(), EncodeError> {
    let observed = work.checked_add(amount).ok_or_else(|| {
        overflow_limit(EncodeLimit::Work {
            observed: usize::MAX,
            maximum,
        })
    })?;
    if observed > maximum {
        return Err(EncodeError::Limit(EncodeLimit::Work { observed, maximum }));
    }
    *work = observed;
    Ok(())
}

fn check_options(
    requirements: ExecutionRequirements,
    options: EncodeOptions,
) -> Result<(), EncodeError> {
    check_execution_limits(requirements, options.execution_limits())
}

fn check_execution_limits(
    requirements: ExecutionRequirements,
    limits: ExecutionLimits,
) -> Result<(), EncodeError> {
    let checks = [
        (
            requirements.output_bytes,
            limits.output_bytes,
            EncodeLimit::Output {
                observed: requirements.output_bytes,
                maximum: limits.output_bytes,
            },
        ),
        (
            requirements.fields,
            limits.fields,
            EncodeLimit::Fields {
                observed: requirements.fields,
                maximum: limits.fields,
            },
        ),
        (
            requirements.work_bytes,
            limits.work_bytes,
            EncodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: limits.work_bytes,
            },
        ),
        (
            requirements.allocations,
            limits.allocations,
            EncodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: limits.allocations,
            },
        ),
        (
            requirements.retained_bytes,
            limits.retained_bytes,
            EncodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: limits.retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            limits.scratch_bytes,
            EncodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: limits.scratch_bytes,
            },
        ),
        (
            requirements.cells,
            limits.cells,
            EncodeLimit::Cells {
                observed: requirements.cells,
                maximum: limits.cells,
            },
        ),
        (
            requirements.labels,
            limits.labels,
            EncodeLimit::Labels {
                observed: requirements.labels,
                maximum: limits.labels,
            },
        ),
        (
            requirements.text_bytes,
            limits.text_bytes,
            EncodeLimit::Text {
                observed: requirements.text_bytes,
                maximum: limits.text_bytes,
            },
        ),
    ];
    for (observed, maximum, limit) in checks {
        if observed > maximum {
            return Err(EncodeError::Limit(limit));
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(EncodeError::Limit(EncodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    Ok(())
}

fn emit_request(
    request: ChartGridCreationRequest<'_>,
    requirements: ExecutionRequirements,
    output: &mut Vec<u8>,
    scalar_cache: &mut buffa::SizeCache,
) -> Result<(), EncodeError> {
    for label in request.row_labels {
        emit_bytes_field(GRID_ROW_NAME_FIELD, label.as_bytes(), output);
    }
    for label in request.column_labels {
        emit_bytes_field(GRID_COLUMN_NAME_FIELD, label.as_bytes(), output);
    }
    for row in request.values {
        let row_payload = row.iter().try_fold(0usize, |length, value| {
            let scalar_bytes = scalar_value_encoded_len(*value)?;
            let value_bytes = message_field_len(GRID_VALUE_FIELD, scalar_bytes)?;
            length.checked_add(value_bytes).ok_or_else(|| {
                overflow_limit(EncodeLimit::Output {
                    observed: usize::MAX,
                    maximum: requirements.output_bytes,
                })
            })
        })?;
        emit_message_prefix(GRID_ROW_FIELD, row_payload, output);
        for value in row {
            emit_grid_value(*value, output, scalar_cache)?;
        }
    }

    let mut map_payload = 0usize;
    for (field, count) in [
        (ID_MAP_ROW_FIELD, request.row_labels.len()),
        (ID_MAP_COLUMN_FIELD, request.column_labels.len()),
    ] {
        for index in 0..count {
            map_payload = map_payload
                .checked_add(message_field_len(field, entry_payload_len(index as u64)?)?)
                .ok_or_else(|| {
                    overflow_limit(EncodeLimit::Output {
                        observed: usize::MAX,
                        maximum: requirements.output_bytes,
                    })
                })?;
        }
    }
    emit_message_prefix(GRID_ID_MAP_FIELD, map_payload, output);
    for (field, count, axis_offset) in [
        (ID_MAP_ROW_FIELD, request.row_labels.len(), 0u64),
        (ID_MAP_COLUMN_FIELD, request.column_labels.len(), 1u64 << 47),
    ] {
        for index in 0..count {
            let index = u32::try_from(index).map_err(|_| {
                EncodeError::InvalidInput(InvalidInput::IndexTooLarge {
                    axis: if axis_offset == 0 { "row" } else { "column" },
                    index,
                })
            })?;
            emit_entry(field, request.seed, axis_offset, index, output)?;
        }
    }
    Ok(())
}

fn emit_grid_value(
    value: Option<f64>,
    output: &mut Vec<u8>,
    scalar_cache: &mut buffa::SizeCache,
) -> Result<(), EncodeError> {
    let scalar_bytes = scalar_value_encoded_len(value)?;
    emit_message_prefix(GRID_VALUE_FIELD, scalar_bytes, output);
    let before = output.len();
    let view: projection::GridValueView<'static> = projection::GridValueView {
        numeric_value: value,
        ..Default::default()
    };
    let encoded = view.try_encode_bounded_with_cache(9, scalar_cache, output)?;
    if usize::try_from(encoded).ok() != Some(scalar_bytes)
        || output.len().checked_sub(before) != Some(scalar_bytes)
    {
        return Err(EncodeError::Verification);
    }
    Ok(())
}

fn scalar_value_encoded_len(value: Option<f64>) -> Result<usize, EncodeError> {
    let view: projection::GridValueView<'static> = projection::GridValueView {
        numeric_value: value,
        ..Default::default()
    };
    usize::try_from(view.try_encoded_len()?).map_err(|_| EncodeError::Verification)
}

fn emit_entry(
    field: u32,
    seed: u64,
    axis_offset: u64,
    index: u32,
    output: &mut Vec<u8>,
) -> Result<(), EncodeError> {
    let payload_len = entry_payload_len(u64::from(index))?;
    emit_message_prefix(field, payload_len, output);
    emit_bytes_prefix(ENTRY_UNIQUE_ID_FIELD, UUID_BYTES, output);
    emit_uuid(
        seed.wrapping_add(axis_offset)
            .wrapping_add(u64::from(index)),
        output,
    );
    emit_varint_field(ENTRY_INDEX_FIELD, u64::from(index), output);
    Ok(())
}

fn emit_uuid(seed: u64, output: &mut Vec<u8>) {
    output.extend_from_slice(UUID_PREFIX);
    let suffix = seed & UUID_SUFFIX_MASK;
    for shift in (0..12).rev() {
        let nibble = ((suffix >> (shift * 4)) & 0x0f) as u8;
        output.push(if nibble < 10 {
            b'0' + nibble
        } else {
            b'A' + (nibble - 10)
        });
    }
}

fn emit_bytes_field(field: u32, bytes: &[u8], output: &mut Vec<u8>) {
    emit_bytes_prefix(field, bytes.len(), output);
    output.extend_from_slice(bytes);
}

fn emit_bytes_prefix(field: u32, length: usize, output: &mut Vec<u8>) {
    emit_key(field, WIRE_LENGTH_DELIMITED, output);
    emit_varint(length as u64, output);
}

fn emit_message_prefix(field: u32, length: usize, output: &mut Vec<u8>) {
    emit_bytes_prefix(field, length, output);
}

fn emit_varint_field(field: u32, value: u64, output: &mut Vec<u8>) {
    emit_key(field, WIRE_VARINT, output);
    emit_varint(value, output);
}

fn emit_key(field: u32, wire_type: u8, output: &mut Vec<u8>) {
    emit_varint((u64::from(field) << 3) | u64::from(wire_type), output);
}

fn emit_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn entry_payload_len(index: u64) -> Result<usize, EncodeError> {
    checked_add(
        string_field_len(ENTRY_UNIQUE_ID_FIELD, UUID_BYTES)?,
        varint_field_len(ENTRY_INDEX_FIELD, index),
        EncodeLimit::Output {
            observed: usize::MAX,
            maximum: usize::MAX,
        },
    )
}

fn string_field_len(field: u32, length: usize) -> Result<usize, EncodeError> {
    message_field_len(field, length)
}

fn message_field_len(field: u32, payload_length: usize) -> Result<usize, EncodeError> {
    let payload_length = u64::try_from(payload_length).map_err(|_| EncodeError::Verification)?;
    checked_add(
        checked_add(
            varint_len((u64::from(field) << 3) | u64::from(WIRE_LENGTH_DELIMITED)),
            varint_len(payload_length),
            EncodeLimit::Output {
                observed: usize::MAX,
                maximum: usize::MAX,
            },
        )?,
        payload_length as usize,
        EncodeLimit::Output {
            observed: usize::MAX,
            maximum: usize::MAX,
        },
    )
}

fn varint_field_len(field: u32, value: u64) -> usize {
    varint_len((u64::from(field) << 3) | u64::from(WIRE_VARINT)) + varint_len(value)
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn checked_add(left: usize, right: usize, limit: EncodeLimit) -> Result<usize, EncodeError> {
    left.checked_add(right).ok_or(EncodeError::Limit(limit))
}

fn overflow_limit(limit: EncodeLimit) -> EncodeError {
    EncodeError::Limit(limit)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tsch;
    use prost::Message as _;

    fn request<'source>(
        rows: &'source [String],
        columns: &'source [String],
        values: &'source [Vec<Option<f64>>],
        seed: u64,
    ) -> ChartGridCreationRequest<'source> {
        ChartGridCreationRequest::new(rows, columns, values, seed)
    }

    fn legacy(
        rows: &[String],
        columns: &[String],
        values: &[Vec<Option<f64>>],
        seed: u64,
    ) -> Vec<u8> {
        let row_id_map = rows
            .iter()
            .enumerate()
            .map(
                |(index, _)| tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                    unique_id: deterministic_uuid(seed.wrapping_add(index as u64)),
                    index: u32::try_from(index).expect("test index"),
                },
            )
            .collect();
        let column_id_map = columns
            .iter()
            .enumerate()
            .map(
                |(index, _)| tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                    unique_id: deterministic_uuid(
                        seed.wrapping_add(1u64 << 47).wrapping_add(index as u64),
                    ),
                    index: u32::try_from(index).expect("test index"),
                },
            )
            .collect();
        tsch::ChartGridArchive {
            row_name: rows.to_vec(),
            column_name: columns.to_vec(),
            grid_row: values
                .iter()
                .map(|row| tsch::GridRow {
                    value: row
                        .iter()
                        .map(|numeric_value| tsch::GridValue {
                            numeric_value: *numeric_value,
                            ..Default::default()
                        })
                        .collect(),
                })
                .collect(),
            id_map: Some(tsch::chart_grid_archive::ChartGridRowColumnIdMap {
                row_id_map,
                column_id_map,
            }),
        }
        .encode_to_vec()
    }

    fn deterministic_uuid(seed: u64) -> String {
        format!("00000000-0000-4000-8000-{:012X}", seed & UUID_SUFFIX_MASK)
    }

    fn sample() -> (Vec<String>, Vec<String>, Vec<Vec<Option<f64>>>) {
        (
            vec![String::from("North"), String::from("South")],
            vec![
                String::from("April"),
                String::from("May"),
                String::from("June"),
            ],
            vec![
                vec![Some(27.5), None, Some(-0.0)],
                vec![Some(55.0), Some(12.75), None],
            ],
        )
    }

    #[test]
    fn matches_legacy_prost_grid_bytes() {
        let (rows, columns, values) = sample();
        let request = request(&rows, &columns, &values, u64::MAX);
        let output = encode_chart_grid(request, EncodeOptions::for_request(&request))
            .expect("bounded grid encoding");
        assert_eq!(output.bytes(), legacy(&rows, &columns, &values, u64::MAX));
        assert_eq!(output.report().cells, 6);
        assert_eq!(output.report().labels, 5);
        assert_eq!(output.report().allocations, 1);
    }

    #[test]
    fn prepared_execution_is_exact_and_one_below_each_resource_refuses() {
        let (rows, columns, values) = sample();
        let request = request(&rows, &columns, &values, 17);
        let prepared = prepare_chart_grid_creation(request, EncodeOptions::for_request(&request))
            .expect("prepare");
        let exact = prepared.execution_requirements();
        assert!(prepared.execute(exact.exact()).is_ok());

        let cases = [
            (
                "output",
                exact.exact().with_output_bytes(exact.output_bytes() - 1),
            ),
            ("fields", exact.exact().with_fields(exact.fields() - 1)),
            (
                "work",
                exact.exact().with_work_bytes(exact.work_bytes() - 1),
            ),
            (
                "allocations",
                exact.exact().with_allocations(exact.allocations() - 1),
            ),
            (
                "retained",
                exact
                    .exact()
                    .with_retained_bytes(exact.retained_bytes() - 1),
            ),
            (
                "scratch",
                exact.exact().with_scratch_bytes(exact.scratch_bytes() - 1),
            ),
            ("cells", exact.exact().with_cells(exact.cells() - 1)),
            ("labels", exact.exact().with_labels(exact.labels() - 1)),
            (
                "text",
                exact.exact().with_text_bytes(exact.text_bytes() - 1),
            ),
            ("depth", exact.exact().with_max_depth(exact.max_depth - 1)),
        ];
        for (name, limits) in cases {
            assert!(
                matches!(prepared.execute(limits), Err(EncodeError::Limit(_))),
                "one-below {name} must refuse"
            );
        }
    }

    #[test]
    fn rejects_shape_and_nonfinite_values_before_output() {
        let rows = vec![String::from("row")];
        let columns = vec![String::from("value")];
        let values = vec![vec![]];
        let request_value = request(&rows, &columns, &values, 0);
        assert!(matches!(
            prepare_chart_grid_creation(request_value, EncodeOptions::for_request(&request_value)),
            Err(EncodeError::InvalidInput(
                InvalidInput::ColumnCountMismatch { .. }
            ))
        ));

        let values = vec![vec![Some(f64::NAN)]];
        let request = request(&rows, &columns, &values, 0);
        assert!(matches!(
            prepare_chart_grid_creation(request, EncodeOptions::for_request(&request)),
            Err(EncodeError::InvalidInput(InvalidInput::NonFiniteNumeric {
                row: 0,
                column: 0
            }))
        ));
    }

    #[test]
    fn preserves_empty_grid_values_as_repeated_empty_messages() {
        let rows = vec![String::from("r")];
        let columns = vec![String::from("c")];
        let values = vec![vec![None]];
        let request = request(&rows, &columns, &values, 0);
        let output =
            encode_chart_grid(request, EncodeOptions::for_request(&request)).expect("empty scalar");
        assert_eq!(output.bytes(), legacy(&rows, &columns, &values, 0));
    }
}
