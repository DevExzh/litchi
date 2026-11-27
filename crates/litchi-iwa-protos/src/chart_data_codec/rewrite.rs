//! Source-preserving numeric chart-grid rewrites.
//!
//! The decoder in the parent module remains the source of truth for the
//! semantic shape.  This module only changes the selected `GridValue` field
//! (field 1) and copies every other wire span verbatim.  In particular, the
//! native labels, id map, date/duration fields, unknown fields, and envelope
//! framing remain caller-owned bytes.

use super::{
    ChartDataSnapshot, DecodeError, DecodeOptions, DecodeReport, GRID_ROW_FIELD, GRID_VALUE_FIELD,
    MODERN_CHART_EXTENSION_FIELD, MODERN_GRID_FIELD, decode_modern_with_report, same_float,
};
use std::{error::Error, fmt};

const ROOT_DEPTH: u32 = 1;
const CHART_DEPTH: u32 = 2;
const GRID_DEPTH: u32 = 3;
const ROW_DEPTH: u32 = 4;
const VALUE_DEPTH: u32 = 5;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Exact finite resources required by one prepared chart-data rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Candidate payload bytes.
    pub output_bytes: usize,
    /// Aggregate source, emission, and candidate field visits.
    pub fields: usize,
    /// Aggregate strict, projection, planning, emission, and readback work.
    pub work_bytes: usize,
    /// Maximum nested message depth needed by the transaction.
    pub max_depth: u32,
    /// Logical candidate allocations.
    pub allocations: usize,
    /// Source plus candidate bytes retained together.
    pub retained_bytes: usize,
    /// Candidate staging bytes charged before allocation.
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Convert exact requirements to caller-enforced execution ceilings.
    #[must_use]
    pub const fn exact(self) -> RewriteExecutionLimits {
        RewriteExecutionLimits {
            output_bytes: self.output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    /// Compatibility alias for callers that spell this as exact limits.
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-provided ceilings checked before the candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    /// Candidate payload bytes.
    pub output_bytes: usize,
    /// Aggregate field visits.
    pub fields: usize,
    /// Aggregate work bytes.
    pub work_bytes: usize,
    /// Maximum nested message depth.
    pub max_depth: u32,
    /// Logical allocations.
    pub allocations: usize,
    /// Source plus candidate retained bytes.
    pub retained_bytes: usize,
    /// Candidate staging bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from one prepared requirement set.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    /// Replace the field ceiling.
    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }

    /// Replace the work-byte ceiling.
    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }

    /// Replace the nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }

    /// Replace the allocation ceiling.
    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }

    /// Replace the retained-byte ceiling.
    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }

    /// Replace the scratch-byte ceiling.
    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Resource axis rejected by a rewrite preparation or execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum RewriteLimit {
    /// Source bytes exceed the configured input ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Field visits exceed the configured ceiling.
    Fields { observed: usize, maximum: usize },
    /// Work exceeds the configured ceiling.
    Work { observed: usize, maximum: usize },
    /// Candidate output exceeds the configured ceiling.
    Output { observed: usize, maximum: usize },
    /// Nested message depth exceeds the configured ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Logical allocations exceed the configured ceiling.
    Allocations { observed: usize, maximum: usize },
    /// Source plus candidate retention exceeds the configured ceiling.
    Retained { observed: usize, maximum: usize },
    /// Candidate staging exceeds the configured ceiling.
    Scratch { observed: usize, maximum: usize },
}

/// Failure from chart-data decode, shape validation, planning, or execution.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RewriteError {
    /// The source failed strict chart-data decoding.
    Decode(DecodeError),
    /// A finite rewrite resource ceiling was exceeded.
    Limit(RewriteLimit),
    /// The request dimensions differ from the decoded rectangular grid.
    Shape {
        expected_rows: usize,
        actual_rows: usize,
        expected_columns: usize,
        actual_columns: usize,
    },
    /// A requested cell is not finite.
    NonFiniteNumeric { row: usize, column: usize },
    /// Candidate output could not be reserved exactly.
    Allocation { amount: usize },
    /// The source-preserving planner and emitter disagreed.
    Projection,
}

impl RewriteError {
    /// Return a typed rewrite resource failure, if present.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<RewriteLimit> {
        match self {
            Self::Limit(limit) => Some(*limit),
            _ => None,
        }
    }

    /// Return the wrapped source decode error, if present.
    #[must_use]
    pub const fn decode_error(&self) -> Option<&DecodeError> {
        match self {
            Self::Decode(error) => Some(error),
            _ => None,
        }
    }

    /// Return a failed candidate allocation size, if present.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self {
            Self::Allocation { amount } => Some(*amount),
            _ => None,
        }
    }

    /// Whether the requested shape differs from the source shape.
    #[must_use]
    pub const fn is_shape_mismatch(&self) -> bool {
        matches!(self, Self::Shape { .. })
    }

    /// Whether the request contains a non-finite numeric value.
    #[must_use]
    pub const fn is_non_finite_numeric(&self) -> bool {
        matches!(self, Self::NonFiniteNumeric { .. })
    }

    const fn projection() -> Self {
        Self::Projection
    }
}

impl fmt::Display for RewriteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Decode(error) => error.fmt(formatter),
            Self::Limit(RewriteLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite input has {observed} bytes; maximum is {maximum}"
            ),
            Self::Limit(RewriteLimit::Fields { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite visits {observed} fields; maximum is {maximum}"
            ),
            Self::Limit(RewriteLimit::Work { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite requires {observed} work bytes; maximum is {maximum}"
            ),
            Self::Limit(RewriteLimit::Output { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite output has {observed} bytes; maximum is {maximum}"
            ),
            Self::Limit(RewriteLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite nesting {observed} exceeds maximum {maximum}"
            ),
            Self::Limit(RewriteLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite requires {observed} allocations; maximum is {maximum}"
            ),
            Self::Limit(RewriteLimit::Retained { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite retains {observed} bytes; maximum is {maximum}"
            ),
            Self::Limit(RewriteLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "chart-data rewrite requires {observed} scratch bytes; maximum is {maximum}"
            ),
            Self::Shape {
                expected_rows,
                actual_rows,
                expected_columns,
                actual_columns,
            } => write!(
                formatter,
                "chart-data rewrite shape is {actual_rows}x{actual_columns}; expected {expected_rows}x{expected_columns}"
            ),
            Self::NonFiniteNumeric { row, column } => {
                write!(
                    formatter,
                    "chart-data rewrite cell ({row}, {column}) is not finite"
                )
            },
            Self::Allocation { amount } => write!(
                formatter,
                "cannot allocate chart-data rewrite output for {amount} bytes"
            ),
            Self::Projection => {
                formatter.write_str("chart-data rewrite source-preserving projection disagreed")
            },
        }
    }
}

impl Error for RewriteError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        match self {
            Self::Decode(error) => Some(error),
            _ => None,
        }
    }
}

impl From<DecodeError> for RewriteError {
    fn from(error: DecodeError) -> Self {
        Self::Decode(error)
    }
}

/// Candidate bytes and complete execution accounting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    output: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    /// Borrow candidate bytes.
    #[must_use]
    pub fn output(&self) -> &[u8] {
        &self.output
    }

    /// Alias for [`Self::output`].
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.output()
    }

    /// Return candidate bytes, consuming this output.
    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.output
    }

    /// Alias for [`Self::into_output`].
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.into_output()
    }

    /// Return complete execution accounting.
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// Complete accounting for one prepared rewrite execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
}

impl RewriteReport {
    /// Source payload bytes.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Candidate payload bytes.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate field visits.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate work bytes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested message depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical candidate allocations.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source plus candidate bytes retained together.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Candidate staging bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether at least one numeric field changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A rewrite prepared without allocating candidate output.
#[derive(Clone, Copy)]
pub struct PreparedChartDataRewrite<'source, 'request> {
    source: &'source [u8],
    values: &'request [Vec<Option<f64>>],
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    changed: bool,
    request_work: usize,
    emitter_fields: usize,
    emitter_work: usize,
    candidate_fields: usize,
    candidate_work: usize,
}

impl fmt::Debug for PreparedChartDataRewrite<'_, '_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedChartDataRewrite")
            .field("source_bytes", &self.source.len())
            .field("rows", &self.values.len())
            .field("changed", &self.changed)
            .field("requirements", &self.requirements)
            .finish()
    }
}

impl PreparedChartDataRewrite<'_, '_> {
    /// Return source-only decode accounting.
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }

    /// Return exact aggregate transaction requirements.
    ///
    /// These requirements include the source decode returned by
    /// [`Self::prepare_report`], request validation, both source-preserving
    /// planning/emission passes, candidate readback, and the output buffer.
    /// A package ledger that has already charged `prepare_report` should
    /// charge only the remaining axes from this value.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Allocate and validate the candidate after checking all caller limits.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, RewriteError> {
        check_execution_limits(self.requirements, limits)?;
        let mut output = reserve_output(self.requirements.output_bytes)?;

        if !self.changed {
            output.extend_from_slice(self.source);
        } else {
            let mut sink = VecSink { output };
            let mut account = Account::new(limits.fields, limits.work_bytes);
            let planned = plan_root(self.source, self.values, &mut account)?;
            if planned.output_len != self.requirements.output_bytes {
                return Err(RewriteError::projection());
            }
            emit_root(self.source, self.values, &mut sink, &mut account)?;
            output = sink.output;
            if output.len() != self.requirements.output_bytes
                || account.fields != self.emitter_fields
                || account.work != self.emitter_work
            {
                return Err(RewriteError::projection());
            }

            let readback_options = self
                .options
                .with_max_input_bytes(output.len().max(self.options.max_input_bytes()))
                .with_max_fields(self.candidate_fields.max(self.options.max_fields()))
                .with_max_work_bytes(self.candidate_work.max(self.options.max_work_bytes()));
            let (snapshot, report) = decode_modern_with_report(&output, &readback_options)?;
            let (matches, request_work) = snapshot_matches_request(snapshot, self.values);
            if report.fields() != self.candidate_fields
                || report.work_bytes() != self.candidate_work
                || request_work != self.request_work
                || !matches
            {
                return Err(RewriteError::projection());
            }
        }

        Ok(RewriteOutput {
            output,
            report: RewriteReport {
                input_bytes: self.source.len(),
                output_bytes: self.requirements.output_bytes,
                fields: self.requirements.fields,
                work_bytes: self.requirements.work_bytes,
                max_depth: self.requirements.max_depth,
                allocations: self.requirements.allocations,
                retained_bytes: self.requirements.retained_bytes,
                scratch_bytes: self.requirements.scratch_bytes,
                changed: self.changed,
            },
        })
    }
}

/// Prepare a same-shape, source-preserving numeric grid rewrite.
pub fn prepare_chart_data_rewrite<'source, 'request>(
    source: &'source [u8],
    values: &'request [Vec<Option<f64>>],
    options: DecodeOptions,
) -> Result<PreparedChartDataRewrite<'source, 'request>, RewriteError> {
    check_option_limits(options)?;
    let (snapshot, source_report) = decode_modern_with_report(source, &options)?;
    let (changed, request_cell_work) =
        validate_request(snapshot, values, options.max_work_bytes())?;
    let request_work = snapshot
        .grid_source()
        .len()
        .checked_add(request_cell_work)
        .ok_or(RewriteError::Limit(RewriteLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes(),
        }))?;

    if !changed {
        let retained_bytes =
            source
                .len()
                .checked_mul(2)
                .ok_or(RewriteError::Limit(RewriteLimit::Retained {
                    observed: usize::MAX,
                    maximum: options.max_retained_bytes(),
                }))?;
        let requirements = RewriteExecutionRequirements {
            output_bytes: source.len(),
            fields: source_report.fields(),
            work_bytes: source_report
                .work_bytes()
                .checked_add(request_work)
                .and_then(|value| value.checked_add(source.len()))
                .ok_or(RewriteError::Limit(RewriteLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes(),
                }))?,
            max_depth: source_report.max_depth(),
            allocations: usize::from(!source.is_empty()),
            retained_bytes,
            scratch_bytes: source.len(),
        };
        check_option_execution_limits(requirements, options)?;
        return Ok(PreparedChartDataRewrite {
            source,
            values,
            options,
            source_report,
            requirements,
            changed: false,
            request_work,
            emitter_fields: 0,
            emitter_work: 0,
            candidate_fields: source_report.fields(),
            candidate_work: source_report.work_bytes(),
        });
    }

    // Planning is source-only. It derives every nested candidate size and
    // readback work cost before any Vec is reserved.
    let account_fields = options.max_fields().saturating_sub(source_report.fields());
    let account_work = options
        .max_work_bytes()
        .saturating_sub(source_report.work_bytes())
        .saturating_sub(request_work);
    let mut account = Account::new(account_fields, account_work);
    let plan = plan_root(source, values, &mut account)?;
    let candidate_fields = source_report
        .fields()
        .checked_add(plan.added_fields)
        .and_then(|value| value.checked_sub(plan.removed_fields))
        .ok_or(RewriteError::Limit(RewriteLimit::Fields {
            observed: usize::MAX,
            maximum: options.max_fields(),
        }))?;
    let candidate_work = candidate_work(&plan)?;

    // This second source walk is the exact emitter operation that execute()
    // repeats. A measure sink gives its length without allocating candidate
    // bytes and keeps the source/request lifetime split intact.
    let mut sink = MeasureSink::default();
    emit_root(source, values, &mut sink, &mut account)?;
    if sink.len != plan.output_len {
        return Err(RewriteError::projection());
    }
    let emitter_fields = account.fields;
    let emitter_work = account.work;
    let retained_bytes = source
        .len()
        .checked_add(plan.output_len)
        .ok_or(RewriteError::Limit(RewriteLimit::Retained {
            observed: usize::MAX,
            maximum: options.max_retained_bytes(),
        }))?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: plan.output_len,
        fields: source_report
            .fields()
            .checked_add(emitter_fields.checked_mul(2).ok_or(RewriteError::Limit(
                RewriteLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields(),
                },
            ))?)
            .and_then(|value| value.checked_add(candidate_fields))
            .ok_or(RewriteError::Limit(RewriteLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields(),
            }))?,
        work_bytes: source_report
            .work_bytes()
            .checked_add(request_work)
            .and_then(|value| value.checked_add(emitter_work.checked_mul(2)?))
            .and_then(|value| value.checked_add(candidate_work))
            .and_then(|value| value.checked_add(plan.grid_len.checked_add(request_cell_work)?))
            .ok_or(RewriteError::Limit(RewriteLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes(),
            }))?,
        max_depth: source_report.max_depth().max(plan.max_depth),
        allocations: usize::from(plan.output_len != 0),
        retained_bytes,
        scratch_bytes: plan.output_len,
    };
    check_option_execution_limits(requirements, options)?;

    Ok(PreparedChartDataRewrite {
        source,
        values,
        options,
        source_report,
        requirements,
        changed: true,
        request_work: request_cell_work,
        emitter_fields,
        emitter_work,
        candidate_fields,
        candidate_work,
    })
}

fn validate_request(
    snapshot: ChartDataSnapshot<'_>,
    values: &[Vec<Option<f64>>],
    maximum_work: usize,
) -> Result<(bool, usize), RewriteError> {
    let expected_rows = snapshot.row_count();
    let expected_columns = snapshot.column_count();
    if values.len() != expected_rows {
        return Err(RewriteError::Shape {
            expected_rows,
            actual_rows: values.len(),
            expected_columns,
            actual_columns: values.first().map_or(0, Vec::len),
        });
    }
    let mut request_work = 0usize;
    for values in values {
        if values.len() != expected_columns {
            return Err(RewriteError::Shape {
                expected_rows,
                actual_rows: expected_rows,
                expected_columns,
                actual_columns: values.len(),
            });
        }
        request_work = request_work
            .checked_add(values.len())
            .and_then(|value| value.checked_add(values.len().checked_mul(8)?))
            .ok_or(RewriteError::Limit(RewriteLimit::Work {
                observed: usize::MAX,
                maximum: maximum_work,
            }))?;
    }
    let source_work = snapshot
        .grid_source()
        .len()
        .checked_add(request_work)
        .ok_or(RewriteError::Limit(RewriteLimit::Work {
            observed: usize::MAX,
            maximum: maximum_work,
        }))?;
    if source_work > maximum_work {
        return Err(RewriteError::Limit(RewriteLimit::Work {
            observed: source_work,
            maximum: maximum_work,
        }));
    }
    let mut changed = false;
    for (row, (source_row, requested)) in snapshot.rows().iter().zip(values).enumerate() {
        for (column, (current, requested)) in source_row.values().zip(requested).enumerate() {
            if requested.is_some_and(|value| !value.is_finite()) {
                return Err(RewriteError::NonFiniteNumeric { row, column });
            }
            changed |= !same_float(current, *requested);
        }
    }
    Ok((changed, request_work))
}

fn snapshot_matches_request(
    snapshot: ChartDataSnapshot<'_>,
    values: &[Vec<Option<f64>>],
) -> (bool, usize) {
    let mut request_work = 0usize;
    let mut matches = true;
    for (source_row, requested) in snapshot.rows().iter().zip(values) {
        request_work = request_work
            .saturating_add(requested.len())
            .saturating_add(requested.len().saturating_mul(8));
        for (current, requested) in source_row.values().zip(requested) {
            matches &= same_float(current, *requested);
        }
    }
    (matches, request_work)
}

fn check_option_limits(options: DecodeOptions) -> Result<(), RewriteError> {
    if options.max_depth() == 0 || options.max_depth() > MAX_RECURSION_LIMIT {
        return Err(RewriteError::Limit(RewriteLimit::Nesting {
            observed: options.max_depth(),
            maximum: MAX_RECURSION_LIMIT,
        }));
    }
    Ok(())
}

fn check_option_execution_limits(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), RewriteError> {
    if requirements.output_bytes > options.max_output_bytes() {
        return Err(RewriteError::Limit(RewriteLimit::Output {
            observed: requirements.output_bytes,
            maximum: options.max_output_bytes(),
        }));
    }
    if requirements.fields > options.max_fields() {
        return Err(RewriteError::Limit(RewriteLimit::Fields {
            observed: requirements.fields,
            maximum: options.max_fields(),
        }));
    }
    if requirements.work_bytes > options.max_work_bytes() {
        return Err(RewriteError::Limit(RewriteLimit::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work_bytes(),
        }));
    }
    if requirements.max_depth > options.max_depth() {
        return Err(RewriteError::Limit(RewriteLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.max_depth(),
        }));
    }
    if requirements.allocations > options.max_allocations() {
        return Err(RewriteError::Limit(RewriteLimit::Allocations {
            observed: requirements.allocations,
            maximum: options.max_allocations(),
        }));
    }
    if requirements.retained_bytes > options.max_retained_bytes() {
        return Err(RewriteError::Limit(RewriteLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: options.max_retained_bytes(),
        }));
    }
    if requirements.scratch_bytes > options.max_scratch_bytes() {
        return Err(RewriteError::Limit(RewriteLimit::Scratch {
            observed: requirements.scratch_bytes,
            maximum: options.max_scratch_bytes(),
        }));
    }
    Ok(())
}

fn check_execution_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), RewriteError> {
    if requirements.output_bytes > limits.output_bytes {
        return Err(RewriteError::Limit(RewriteLimit::Output {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    if requirements.fields > limits.fields {
        return Err(RewriteError::Limit(RewriteLimit::Fields {
            observed: requirements.fields,
            maximum: limits.fields,
        }));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(RewriteError::Limit(RewriteLimit::Work {
            observed: requirements.work_bytes,
            maximum: limits.work_bytes,
        }));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(RewriteError::Limit(RewriteLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    if requirements.allocations > limits.allocations {
        return Err(RewriteError::Limit(RewriteLimit::Allocations {
            observed: requirements.allocations,
            maximum: limits.allocations,
        }));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(RewriteError::Limit(RewriteLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: limits.retained_bytes,
        }));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(RewriteError::Limit(RewriteLimit::Scratch {
            observed: requirements.scratch_bytes,
            maximum: limits.scratch_bytes,
        }));
    }
    Ok(())
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, RewriteError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_| RewriteError::Allocation { amount })?;
    Ok(output)
}

#[derive(Debug, Clone, Copy, Default)]
struct Plan {
    output_len: usize,
    changed: bool,
    added_fields: usize,
    removed_fields: usize,
    chart_len: usize,
    grid_len: usize,
    row_bytes: usize,
    value_bytes: usize,
    max_depth: u32,
}

#[derive(Debug, Default)]
struct Account {
    fields: usize,
    work: usize,
    max_fields: usize,
    max_work: usize,
}

impl Account {
    fn new(max_fields: usize, max_work: usize) -> Self {
        Self {
            fields: 0,
            work: 0,
            max_fields,
            max_work,
        }
    }

    fn message(&mut self, bytes: usize) -> Result<(), RewriteError> {
        self.charge_work(bytes)
    }

    fn output(&mut self, bytes: usize) -> Result<(), RewriteError> {
        self.charge_work(bytes)
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), RewriteError> {
        let observed =
            self.work
                .checked_add(bytes)
                .ok_or(RewriteError::Limit(RewriteLimit::Work {
                    observed: usize::MAX,
                    maximum: self.max_work,
                }))?;
        if observed > self.max_work {
            return Err(RewriteError::Limit(RewriteLimit::Work {
                observed,
                maximum: self.max_work,
            }));
        }
        self.work = observed;
        Ok(())
    }

    fn field(&mut self) -> Result<(), RewriteError> {
        let observed =
            self.fields
                .checked_add(1)
                .ok_or(RewriteError::Limit(RewriteLimit::Fields {
                    observed: usize::MAX,
                    maximum: self.max_fields,
                }))?;
        if observed > self.max_fields {
            return Err(RewriteError::Limit(RewriteLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn fields(&mut self, amount: usize) -> Result<(), RewriteError> {
        let observed =
            self.fields
                .checked_add(amount)
                .ok_or(RewriteError::Limit(RewriteLimit::Fields {
                    observed: usize::MAX,
                    maximum: self.max_fields,
                }))?;
        if observed > self.max_fields {
            return Err(RewriteError::Limit(RewriteLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn synthetic_field(&mut self) -> Result<(), RewriteError> {
        self.field()?;
        self.charge_work(9)
    }
}

fn candidate_work(plan: &Plan) -> Result<usize, RewriteError> {
    // The decoder charges 2x for each strict message, one additional pass for
    // the chart view and grid scan, and 2x for every lazy scalar value.
    plan.output_len
        .checked_mul(2)
        .and_then(|bytes| bytes.checked_add(plan.chart_len.checked_mul(3)?))
        .and_then(|bytes| bytes.checked_add(plan.grid_len.checked_mul(3)?))
        .and_then(|bytes| bytes.checked_add(plan.row_bytes.checked_mul(2)?))
        .and_then(|bytes| bytes.checked_add(plan.value_bytes.checked_mul(4)?))
        .ok_or(RewriteError::projection())
}

fn plan_root(
    source: &[u8],
    values: &[Vec<Option<f64>>],
    account: &mut Account,
) -> Result<Plan, RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut output_len = 0usize;
    let mut chart_len = 0usize;
    let mut saw_chart = false;
    let mut plan = Plan {
        max_depth: ROOT_DEPTH,
        ..Plan::default()
    };
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != MODERN_CHART_EXTENSION_FIELD {
            output_len = add_len(output_len, field.len())?;
            continue;
        }
        if saw_chart {
            return Err(RewriteError::projection());
        }
        saw_chart = true;
        let payload = field.payload()?;
        let child = plan_chart(payload, values, account)?;
        chart_len = child.output_len;
        output_len = add_len(output_len, length_field_len(&field, child.output_len)?)?;
        plan.changed |= child.changed;
        plan.added_fields = add_len(plan.added_fields, child.added_fields)?;
        plan.removed_fields = add_len(plan.removed_fields, child.removed_fields)?;
        plan.grid_len = child.grid_len;
        plan.row_bytes = child.row_bytes;
        plan.value_bytes = child.value_bytes;
        plan.max_depth = plan.max_depth.max(child.max_depth);
    }
    if !saw_chart {
        return Err(RewriteError::projection());
    }
    plan.output_len = output_len;
    plan.chart_len = chart_len;
    Ok(plan)
}

fn plan_chart(
    source: &[u8],
    values: &[Vec<Option<f64>>],
    account: &mut Account,
) -> Result<Plan, RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut output_len = 0usize;
    let mut grid_len = 0usize;
    let mut saw_grid = false;
    let mut plan = Plan {
        max_depth: CHART_DEPTH,
        ..Plan::default()
    };
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != MODERN_GRID_FIELD {
            output_len = add_len(output_len, field.len())?;
            continue;
        }
        if saw_grid {
            return Err(RewriteError::projection());
        }
        saw_grid = true;
        let payload = field.payload()?;
        let child = plan_grid(payload, values, account)?;
        grid_len = child.output_len;
        output_len = add_len(output_len, length_field_len(&field, child.output_len)?)?;
        plan.changed |= child.changed;
        plan.added_fields = add_len(plan.added_fields, child.added_fields)?;
        plan.removed_fields = add_len(plan.removed_fields, child.removed_fields)?;
        plan.row_bytes = child.row_bytes;
        plan.value_bytes = child.value_bytes;
        plan.max_depth = plan.max_depth.max(child.max_depth);
    }
    if !saw_grid {
        return Err(RewriteError::projection());
    }
    plan.output_len = output_len;
    plan.grid_len = grid_len;
    Ok(plan)
}

fn plan_grid(
    source: &[u8],
    values: &[Vec<Option<f64>>],
    account: &mut Account,
) -> Result<Plan, RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut output_len = 0usize;
    let mut row_index = 0usize;
    let mut plan = Plan {
        max_depth: GRID_DEPTH,
        ..Plan::default()
    };
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != GRID_ROW_FIELD {
            output_len = add_len(output_len, field.len())?;
            continue;
        }
        let payload = field.payload()?;
        let request = values.get(row_index).ok_or(RewriteError::projection())?;
        let child = plan_row(payload, request, account)?;
        output_len = add_len(output_len, length_field_len(&field, child.output_len)?)?;
        plan.changed |= child.changed;
        plan.added_fields = add_len(plan.added_fields, child.added_fields)?;
        plan.removed_fields = add_len(plan.removed_fields, child.removed_fields)?;
        plan.row_bytes = add_len(plan.row_bytes, child.output_len)?;
        plan.value_bytes = add_len(plan.value_bytes, child.value_bytes)?;
        plan.max_depth = plan.max_depth.max(child.max_depth);
        row_index = row_index.checked_add(1).ok_or(RewriteError::projection())?;
    }
    if row_index != values.len() {
        return Err(RewriteError::projection());
    }
    plan.output_len = output_len;
    Ok(plan)
}

fn plan_row(
    source: &[u8],
    request: &[Option<f64>],
    account: &mut Account,
) -> Result<Plan, RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut output_len = 0usize;
    let mut column = 0usize;
    let mut plan = Plan {
        max_depth: ROW_DEPTH,
        ..Plan::default()
    };
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != GRID_VALUE_FIELD {
            output_len = add_len(output_len, field.len())?;
            continue;
        }
        let payload = field.payload()?;
        let target = *request.get(column).ok_or(RewriteError::projection())?;
        let child = plan_value(payload, target, account)?;
        output_len = add_len(output_len, length_field_len(&field, child.output_len)?)?;
        plan.changed |= child.changed;
        plan.added_fields = add_len(plan.added_fields, child.added_fields)?;
        plan.removed_fields = add_len(plan.removed_fields, child.removed_fields)?;
        plan.value_bytes = add_len(plan.value_bytes, child.output_len)?;
        plan.max_depth = plan.max_depth.max(child.max_depth);
        column = column.checked_add(1).ok_or(RewriteError::projection())?;
    }
    if column != request.len() {
        return Err(RewriteError::projection());
    }
    plan.output_len = output_len;
    Ok(plan)
}

fn plan_value(
    source: &[u8],
    target: Option<f64>,
    account: &mut Account,
) -> Result<Plan, RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut output_len = 0usize;
    let mut numeric = None;
    let mut plan = Plan {
        max_depth: VALUE_DEPTH,
        ..Plan::default()
    };
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != GRID_VALUE_FIELD {
            output_len = add_len(output_len, field.len())?;
            continue;
        }
        if field.wire_type != 1 || field.value_len() != 8 {
            return Err(RewriteError::projection());
        }
        if numeric.is_some() {
            return Err(RewriteError::projection());
        }
        let bytes = field.fixed64_bytes(source)?;
        numeric = Some(f64::from_le_bytes(bytes));
        match target {
            Some(target) => {
                output_len = add_len(output_len, field.len())?;
                if !same_float(numeric, Some(target)) {
                    plan.changed = true;
                }
            },
            None => {
                plan.changed = true;
                plan.removed_fields = 1;
            },
        }
    }
    if target.is_some() && numeric.is_none() {
        plan.changed = true;
        plan.added_fields = 1;
        account.synthetic_field()?;
        output_len = add_len(output_len, 9)?;
    }
    if target.is_none() && numeric.is_none() {
        // An absent numeric target with no numeric source is a true no-op.
        plan.changed = false;
    }
    plan.output_len = output_len;
    Ok(plan)
}

fn emit_root<S: Sink>(
    source: &[u8],
    values: &[Vec<Option<f64>>],
    sink: &mut S,
    account: &mut Account,
) -> Result<(), RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != MODERN_CHART_EXTENSION_FIELD {
            write(sink, account, &source[field.start..field.end])?;
            continue;
        }
        let payload = field.payload()?;
        let plan = plan_chart(payload, values, account)?;
        write_length_prefix(sink, account, source, &field, plan.output_len)?;
        emit_chart(payload, values, sink, account)?;
    }
    Ok(())
}

fn emit_chart<S: Sink>(
    source: &[u8],
    values: &[Vec<Option<f64>>],
    sink: &mut S,
    account: &mut Account,
) -> Result<(), RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != MODERN_GRID_FIELD {
            write(sink, account, &source[field.start..field.end])?;
            continue;
        }
        let payload = field.payload()?;
        let plan = plan_grid(payload, values, account)?;
        write_length_prefix(sink, account, source, &field, plan.output_len)?;
        emit_grid(payload, values, sink, account)?;
    }
    Ok(())
}

fn emit_grid<S: Sink>(
    source: &[u8],
    values: &[Vec<Option<f64>>],
    sink: &mut S,
    account: &mut Account,
) -> Result<(), RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut row_index = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != GRID_ROW_FIELD {
            write(sink, account, &source[field.start..field.end])?;
            continue;
        }
        let payload = field.payload()?;
        let request = values.get(row_index).ok_or(RewriteError::projection())?;
        let plan = plan_row(payload, request, account)?;
        write_length_prefix(sink, account, source, &field, plan.output_len)?;
        emit_row(payload, request, sink, account)?;
        row_index = row_index.checked_add(1).ok_or(RewriteError::projection())?;
    }
    if row_index != values.len() {
        return Err(RewriteError::projection());
    }
    Ok(())
}

fn emit_row<S: Sink>(
    source: &[u8],
    request: &[Option<f64>],
    sink: &mut S,
    account: &mut Account,
) -> Result<(), RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut column = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != GRID_VALUE_FIELD {
            write(sink, account, &source[field.start..field.end])?;
            continue;
        }
        let payload = field.payload()?;
        let target = *request.get(column).ok_or(RewriteError::projection())?;
        let plan = plan_value(payload, target, account)?;
        if !plan.changed {
            write(sink, account, &source[field.start..field.end])?;
        } else {
            write_length_prefix(sink, account, source, &field, plan.output_len)?;
            emit_value(payload, target, sink, account)?;
        }
        column = column.checked_add(1).ok_or(RewriteError::projection())?;
    }
    if column != request.len() {
        return Err(RewriteError::projection());
    }
    Ok(())
}

fn emit_value<S: Sink>(
    source: &[u8],
    target: Option<f64>,
    sink: &mut S,
    account: &mut Account,
) -> Result<(), RewriteError> {
    account.message(source.len())?;
    let mut cursor = 0usize;
    let mut numeric = false;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        account.field()?;
        account.fields(field.nested_fields)?;
        if field.number != GRID_VALUE_FIELD {
            write(sink, account, &source[field.start..field.end])?;
            continue;
        }
        if field.wire_type != 1 || field.value_len() != 8 || numeric {
            return Err(RewriteError::projection());
        }
        numeric = true;
        let current = f64::from_le_bytes(field.fixed64_bytes(source)?);
        if let Some(target) = target {
            write(sink, account, &source[field.start..field.value_start])?;
            write(sink, account, &target.to_le_bytes())?;
        }
        let _ = current;
    }
    if target.is_some() && !numeric {
        account.synthetic_field()?;
        write(sink, account, &[(GRID_VALUE_FIELD << 3 | 1) as u8])?;
        write(sink, account, &target.unwrap_or_default().to_le_bytes())?;
    }
    Ok(())
}

trait Sink {
    fn append(&mut self, bytes: &[u8]) -> Result<(), RewriteError>;
}

#[derive(Debug, Default)]
struct MeasureSink {
    len: usize,
}

impl Sink for MeasureSink {
    fn append(&mut self, bytes: &[u8]) -> Result<(), RewriteError> {
        self.len = self
            .len
            .checked_add(bytes.len())
            .ok_or(RewriteError::projection())?;
        Ok(())
    }
}

#[derive(Debug)]
struct VecSink {
    output: Vec<u8>,
}

impl Sink for VecSink {
    fn append(&mut self, bytes: &[u8]) -> Result<(), RewriteError> {
        self.output.extend_from_slice(bytes);
        Ok(())
    }
}

fn write<S: Sink>(sink: &mut S, account: &mut Account, bytes: &[u8]) -> Result<(), RewriteError> {
    sink.append(bytes)?;
    account.output(bytes.len())
}

fn add_len(left: usize, right: usize) -> Result<usize, RewriteError> {
    left.checked_add(right).ok_or(RewriteError::projection())
}

fn length_field_len(field: &RawField<'_>, payload_len: usize) -> Result<usize, RewriteError> {
    field
        .tag_len()
        .checked_add(varint_len(payload_len))
        .and_then(|length| length.checked_add(payload_len))
        .ok_or(RewriteError::projection())
}

fn write_length_prefix<S: Sink>(
    sink: &mut S,
    account: &mut Account,
    source: &[u8],
    field: &RawField<'_>,
    payload_len: usize,
) -> Result<(), RewriteError> {
    write(sink, account, &source[field.start..field.tag_end])?;
    let mut encoded = [0u8; 10];
    let length = encode_varint(payload_len as u64, &mut encoded);
    write(sink, account, &encoded[..length])
}

fn varint_len(mut value: usize) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn encode_varint(mut value: u64, output: &mut [u8; 10]) -> usize {
    let mut index = 0usize;
    while value >= 0x80 {
        output[index] = (value as u8 & 0x7f) | 0x80;
        value >>= 7;
        index += 1;
    }
    output[index] = value as u8;
    index + 1
}

#[derive(Debug, Clone, Copy)]
struct RawField<'source> {
    source: &'source [u8],
    start: usize,
    tag_end: usize,
    value_start: usize,
    payload_start: usize,
    payload_end: usize,
    end: usize,
    number: u32,
    wire_type: u32,
    nested_fields: usize,
}

impl<'source> RawField<'source> {
    fn len(self) -> usize {
        self.end - self.start
    }

    fn tag_len(self) -> usize {
        self.tag_end - self.start
    }

    fn value_len(self) -> usize {
        self.payload_end.saturating_sub(self.payload_start)
    }

    fn payload(self) -> Result<&'source [u8], RewriteError> {
        if self.wire_type != 2 {
            return Err(RewriteError::projection());
        }
        Ok(&self.source[self.payload_start..self.payload_end])
    }

    fn fixed64_bytes(self, source: &[u8]) -> Result<[u8; 8], RewriteError> {
        if self.wire_type != 1 || self.value_len() != 8 {
            return Err(RewriteError::projection());
        }
        source[self.value_start..self.end]
            .try_into()
            .map_err(|_| RewriteError::projection())
    }
}

fn next_raw_field<'source>(
    source: &'source [u8],
    cursor: &mut usize,
) -> Result<Option<RawField<'source>>, RewriteError> {
    if *cursor == source.len() {
        return Ok(None);
    }
    if *cursor > source.len() {
        return Err(RewriteError::projection());
    }
    let start = *cursor;
    let (tag, tag_end) = read_varint(source, *cursor)?;
    *cursor = tag_end;
    let number = u32::try_from(tag >> 3).map_err(|_| RewriteError::projection())?;
    let wire_type = u32::try_from(tag & 7).map_err(|_| RewriteError::projection())?;
    if number == 0 || wire_type == 4 || wire_type == 6 || wire_type == 7 {
        return Err(RewriteError::projection());
    }
    let value_start = *cursor;
    let (payload_start, payload_end, end) = match wire_type {
        0 => {
            let (_, end) = read_varint(source, *cursor)?;
            (value_start, end, end)
        },
        1 => {
            let end = checked_end(source, *cursor, 8)?;
            (value_start, end, end)
        },
        2 => {
            let (length, payload_start) = read_varint(source, *cursor)?;
            let length = usize::try_from(length).map_err(|_| RewriteError::projection())?;
            let payload_end = checked_end(source, payload_start, length)?;
            (payload_start, payload_end, payload_end)
        },
        3 => {
            let (end, nested_fields) = skip_group(source, *cursor, number)?;
            *cursor = end;
            return Ok(Some(RawField {
                source,
                start,
                tag_end,
                value_start,
                payload_start: value_start,
                payload_end: end,
                end,
                number,
                wire_type,
                nested_fields,
            }));
        },
        5 => {
            let end = checked_end(source, *cursor, 4)?;
            (value_start, end, end)
        },
        _ => return Err(RewriteError::projection()),
    };
    *cursor = end;
    Ok(Some(RawField {
        source,
        start,
        tag_end,
        value_start,
        payload_start,
        payload_end,
        end,
        number,
        wire_type,
        nested_fields: 0,
    }))
}

fn read_varint(source: &[u8], start: usize) -> Result<(u64, usize), RewriteError> {
    let mut value = 0u64;
    for index in 0..10usize {
        let position = start.checked_add(index).ok_or(RewriteError::projection())?;
        let byte = *source.get(position).ok_or(RewriteError::projection())?;
        if index == 9 && byte > 1 {
            return Err(RewriteError::projection());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            if index > 0 && value < (1u64 << (index * 7)) {
                return Err(RewriteError::projection());
            }
            return Ok((value, position + 1));
        }
    }
    Err(RewriteError::projection())
}

fn checked_end(source: &[u8], start: usize, length: usize) -> Result<usize, RewriteError> {
    let end = start
        .checked_add(length)
        .ok_or(RewriteError::projection())?;
    if end > source.len() {
        return Err(RewriteError::projection());
    }
    Ok(end)
}

fn skip_group(
    source: &[u8],
    mut cursor: usize,
    expected: u32,
) -> Result<(usize, usize), RewriteError> {
    let mut nested_fields = 0usize;
    while cursor < source.len() {
        let (tag, tag_end) = read_varint(source, cursor)?;
        cursor = tag_end;
        let number = u32::try_from(tag >> 3).map_err(|_| RewriteError::projection())?;
        let wire_type = u32::try_from(tag & 7).map_err(|_| RewriteError::projection())?;
        nested_fields = nested_fields
            .checked_add(1)
            .ok_or(RewriteError::projection())?;
        match wire_type {
            0 => cursor = read_varint(source, cursor)?.1,
            1 => cursor = checked_end(source, cursor, 8)?,
            2 => {
                let (length, payload) = read_varint(source, cursor)?;
                let length = usize::try_from(length).map_err(|_| RewriteError::projection())?;
                cursor = checked_end(source, payload, length)?;
            },
            3 => {
                let (end, child_fields) = skip_group(source, cursor, number)?;
                cursor = end;
                nested_fields = nested_fields
                    .checked_add(child_fields)
                    .ok_or(RewriteError::projection())?;
            },
            4 if number == expected => return Ok((cursor, nested_fields)),
            4 => return Err(RewriteError::projection()),
            5 => cursor = checked_end(source, cursor, 4)?,
            _ => return Err(RewriteError::projection()),
        }
    }
    Err(RewriteError::projection())
}
