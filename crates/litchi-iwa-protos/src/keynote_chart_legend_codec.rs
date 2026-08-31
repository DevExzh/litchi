//! Strict borrowed projection for Keynote chart legend visibility.
//!
//! The selected payload is only field 20 of the generated
//! `ChartNonStyleArchive` extension. The outer chart envelope, chart graph,
//! unrelated generated fields, and unknown source bytes remain caller-owned.
//! This module performs a bounded borrowed projection and a source-local wire
//! rewrite; it never performs chart lookup or archive transactions.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire preflight intentionally precedes the generated view."
)]

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_chart_legend_generated::LitchiIwaProjection as projection;

const CHART_LEGEND_VISIBLE_FIELD: u32 = 20;
const ROOT_DEPTH: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;
const LEGEND_VISIBLE_FIELD_NAME: &str =
    "TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaultshowlegend";

/// Finite limits for one generated chart-legend extension payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Build explicit source, field, work, and nesting ceilings.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_output_bytes: max_message_bytes,
            max_allocations: max_message_bytes,
            max_retained_bytes: max_message_bytes,
            max_scratch_bytes: max_message_bytes,
        }
    }

    /// Build conservative finite limits from a known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        let fields = bytes
            .checked_mul(4)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT);
        let work = bytes
            .checked_mul(16)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT);
        let retained = bytes
            .checked_mul(2)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT);
        Self::new(bytes, fields.max(1), work.max(1), 8)
            .with_max_allocations(1)
            .with_max_retained_bytes(retained.max(bytes))
            .with_max_scratch_bytes(bytes)
    }

    /// Replace the candidate output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the logical allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the source-plus-candidate retained-byte ceiling.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace the candidate scratch-byte ceiling.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Replace the field-visit ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate work-byte ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the protobuf nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Borrowed semantic facts from one generated chart non-style extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartLegendVisibilitySnapshot<'source> {
    visible: Option<bool>,
    raw: &'source [u8],
}

impl<'source> ChartLegendVisibilitySnapshot<'source> {
    /// Return the optional native field-20 value, retaining proto2 presence.
    #[must_use]
    pub const fn visible(self) -> Option<bool> {
        self.visible
    }

    /// Return the optional native field-20 value.
    #[must_use]
    pub const fn show_legend(self) -> Option<bool> {
        self.visible
    }

    /// Return the effective visibility. An absent native field is false.
    #[must_use]
    pub const fn is_visible(self) -> bool {
        match self.visible {
            Some(value) => value,
            None => false,
        }
    }

    /// Return the exact caller-owned bytes used by this snapshot.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Compatibility alias for callers that use the shorter chart-legend name.
pub type ChartLegendSnapshot<'source> = ChartLegendVisibilitySnapshot<'source>;

/// Presence-aware requested value for a chart legend rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartLegendVisibilityWrite {
    visible: Option<bool>,
    clear: bool,
    force_presence: bool,
}

impl ChartLegendVisibilityWrite {
    /// Set the effective legend visibility.
    #[must_use]
    pub const fn new(visible: bool) -> Self {
        Self {
            visible: Some(visible),
            clear: false,
            force_presence: false,
        }
    }

    /// Set an explicitly present native value, or clear field 20 with `None`.
    #[must_use]
    pub const fn explicit(visible: Option<bool>) -> Self {
        Self {
            visible,
            clear: visible.is_none(),
            force_presence: visible.is_some(),
        }
    }

    /// Remove field 20, restoring the generated default.
    #[must_use]
    pub const fn clear() -> Self {
        Self {
            visible: None,
            clear: true,
            force_presence: false,
        }
    }

    /// Return the requested explicit value, if any.
    #[must_use]
    pub const fn visible(self) -> Option<bool> {
        self.visible
    }

    /// Return the requested effective value. An absent field is false.
    #[must_use]
    pub const fn effective_visible(self) -> bool {
        match self.visible {
            Some(value) => value,
            None => false,
        }
    }
}

/// Compatibility alias for callers that use the shorter chart-legend name.
pub type ChartLegendWrite = ChartLegendVisibilityWrite;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NormalizedWrite {
    Preserve,
    Replace(bool),
    Clear,
}

/// Exact resource consumption from one successful decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Input bytes inspected by strict preflight and Buffa.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Encoded fields visited, including unknown fields and group contents.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa work charged by the decoder.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum protobuf nesting observed during strict traversal.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical allocations performed by the borrowed decoder (always zero).
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by the borrowed snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Scratch bytes allocated by the borrowed decoder (always zero).
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact prepared rewrite requirements.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Candidate payload bytes.
    pub output_bytes: usize,
    /// Aggregate source, rewrite, and candidate field visits.
    pub fields: usize,
    /// Aggregate strict-plus-Buffa work bytes.
    pub work_bytes: usize,
    /// Maximum nesting needed by the transaction.
    pub max_depth: u32,
    /// Logical candidate allocations.
    pub allocations: usize,
    /// Source plus candidate bytes retained together.
    pub retained_bytes: usize,
    /// Candidate staging bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Convert requirements to exact execution ceilings.
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

    /// Alias for [`Self::exact`].
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-provided ceilings checked before candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    /// Candidate payload bytes.
    pub output_bytes: usize,
    /// Aggregate field visits.
    pub fields: usize,
    /// Aggregate work bytes.
    pub work_bytes: usize,
    /// Maximum nesting depth.
    pub max_depth: u32,
    /// Logical allocations.
    pub allocations: usize,
    /// Source plus candidate retained bytes.
    pub retained_bytes: usize,
    /// Candidate scratch bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from requirements.
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

    /// Replace the maximum depth.
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

/// Candidate bytes and the exact report from a prepared rewrite.
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
        &self.output
    }

    /// Return candidate bytes, consuming this result.
    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.output
    }

    /// Return exact execution accounting.
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// Exact complete execution accounting for one prepared rewrite.
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

    /// Maximum nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical allocations.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source plus candidate retained bytes.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Candidate scratch bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether candidate bytes differ from the source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source or configured message bytes exceeded their ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Field visits exceeded their ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate strict-plus-Buffa work exceeded its ceiling.
    Work { observed: usize, maximum: usize },
    /// Candidate output bytes exceeded their ceiling.
    Output { observed: usize, maximum: usize },
    /// Protobuf nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Logical allocations exceeded their ceiling.
    Allocations { observed: usize, maximum: usize },
    /// Source plus candidate retained bytes exceeded their ceiling.
    Retained { observed: usize, maximum: usize },
    /// Candidate scratch bytes exceeded their ceiling.
    Scratch { observed: usize, maximum: usize },
}

/// The byte/nesting subset used by older focused codecs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// Source or configured message bytes exceeded their ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Protobuf nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Failure from strict preflight or the private Buffa projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(DecodeLimit),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Allocation { amount: usize },
    Projection,
}

impl DecodeError {
    /// Return the duplicated singular field, if any.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return the canonical-wire failure reason, if any.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    /// Return typed finite resource failure, if any.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return the byte/nesting resource failure, if any.
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        match self.kind {
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => {
                Some(WireResourceLimit::Bytes { observed, maximum })
            },
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => {
                Some(WireResourceLimit::Nesting { observed, maximum })
            },
            _ => None,
        }
    }

    /// Return the failed allocation amount, if any.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            _ => None,
        }
    }

    /// Return the field-limit observation, if any.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return the work-limit observation, if any.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return the output-limit observation, if any.
    #[must_use]
    pub const fn output_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    const fn resource(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Resource(limit),
        }
    }

    const fn duplicate() -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(LEGEND_VISIBLE_FIELD_NAME),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend projection output has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend projection nesting {observed} exceeds maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend rewrite requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend rewrite retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "Keynote chart-legend rewrite requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate Keynote chart-legend output for {amount} bytes"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote chart-legend strict preflight disagrees with the Buffa projection",
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

/// One source-preserving legend rewrite prepared without allocating output.
#[derive(Debug, Clone, Copy)]
pub struct PreparedChartLegendVisibilityRewrite<'source> {
    source: &'source [u8],
    update: NormalizedWrite,
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
}

impl PreparedChartLegendVisibilityRewrite<'_> {
    /// Return source-only preparation accounting.
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }

    /// Return exact aggregate execution requirements.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Allocate and validate the candidate after checking caller ceilings.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_execution_limits(self.requirements, limits)?;
        let output = emit_rewrite(self.source, self.update, self.requirements.output_bytes)?;
        let readback_options = DecodeOptions::new(
            output.len().max(1),
            self.candidate_fields.max(1),
            self.candidate_work.max(1),
            self.options.recursion_limit,
        )
        .with_max_output_bytes(self.options.max_output_bytes)
        .with_max_allocations(self.options.max_allocations)
        .with_max_retained_bytes(self.options.max_retained_bytes)
        .with_max_scratch_bytes(self.options.max_scratch_bytes);
        let (readback, candidate_report) =
            decode_chart_legend_visibility_with_report(&output, readback_options)?;
        if candidate_report.fields != self.candidate_fields
            || candidate_report.work_bytes != self.candidate_work
            || candidate_report.max_depth != self.candidate_depth
            || !candidate_matches(readback, self.update)
        {
            return Err(DecodeError::projection());
        }
        Ok(RewriteOutput {
            report: RewriteReport {
                input_bytes: self.source.len(),
                output_bytes: output.len(),
                fields: self.requirements.fields,
                work_bytes: self.requirements.work_bytes,
                max_depth: self.requirements.max_depth,
                allocations: self.requirements.allocations,
                retained_bytes: self.requirements.retained_bytes,
                scratch_bytes: self.requirements.scratch_bytes,
                changed: output.as_slice() != self.source,
            },
            output,
        })
    }
}

/// Decode a generated chart non-style extension.
pub fn decode_chart_legend_visibility<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ChartLegendVisibilitySnapshot<'source>, DecodeError> {
    Ok(decode_chart_legend_visibility_with_report(source, options)?.0)
}

/// Decode a generated extension and return exact resource accounting.
pub fn decode_chart_legend_visibility_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(ChartLegendVisibilitySnapshot<'source>, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = decode_with_budget(source, options, &mut budget)?;
    let report = budget.report();
    check_decode_report_limits(report, options)?;
    Ok((snapshot, report))
}

/// Compatibility spelling for callers referring to the generated extension.
pub fn decode_chart_legend_visibility_extension<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ChartLegendVisibilitySnapshot<'source>, DecodeError> {
    decode_chart_legend_visibility(source, options)
}

/// Compatibility spelling for callers using the shorter chart-legend name.
pub fn decode_chart_legend<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ChartLegendVisibilitySnapshot<'source>, DecodeError> {
    decode_chart_legend_visibility(source, options)
}

/// Prepare a source-preserving legend rewrite without allocating output.
pub fn prepare_chart_legend_visibility_rewrite<'source>(
    source: &'source [u8],
    write: ChartLegendVisibilityWrite,
    options: DecodeOptions,
) -> Result<PreparedChartLegendVisibilityRewrite<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let current = decode_with_budget(source, options, &mut budget)?;
    let source_report = budget.report();
    let update = normalize_write(current, write);
    let measure = measure_rewrite_output(source, update, options, &mut budget)?;
    let candidate_fields = candidate_field_count(source_report.fields, current.visible, update)?;
    let candidate_work = candidate_work(measure.output_bytes)?;
    let fields = budget
        .fields
        .checked_add(source_report.fields)
        .and_then(|value| value.checked_add(candidate_fields))
        .ok_or_else(|| field_size_overflow(options.max_fields))?;
    let work_bytes = source_report
        .work_bytes
        .checked_add(
            source
                .len()
                .checked_mul(2)
                .ok_or_else(|| work_size_overflow(options.max_work_bytes))?,
        )
        .and_then(|value| {
            value.checked_add(candidate_work).and_then(|value| {
                // Emission is another complete strict wire walk.
                value.checked_add(source.len().checked_mul(2)?)
            })
        })
        .ok_or_else(|| work_size_overflow(options.max_work_bytes))?;
    let allocations = usize::from(measure.output_bytes != 0);
    let retained_bytes = source
        .len()
        .checked_add(measure.output_bytes)
        .ok_or_else(|| retained_size_overflow(options.max_retained_bytes))?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: measure.output_bytes,
        fields,
        work_bytes,
        max_depth: source_report.max_depth.max(measure.max_depth),
        allocations,
        retained_bytes,
        scratch_bytes: measure.output_bytes,
    };
    check_option_execution_limits(requirements, options)?;
    Ok(PreparedChartLegendVisibilityRewrite {
        source,
        update,
        options,
        source_report,
        requirements,
        candidate_fields,
        candidate_work,
        candidate_depth: measure.max_depth,
    })
}

/// Compatibility spelling for callers using the shorter chart-legend name.
pub fn prepare_chart_legend_rewrite<'source>(
    source: &'source [u8],
    write: ChartLegendVisibilityWrite,
    options: DecodeOptions,
) -> Result<PreparedChartLegendVisibilityRewrite<'source>, DecodeError> {
    prepare_chart_legend_visibility_rewrite(source, write, options)
}

/// Rewrite field 20 while preserving all unselected source spans.
pub fn rewrite_chart_legend_visibility(
    source: &[u8],
    write: ChartLegendVisibilityWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_chart_legend_visibility_with_report(source, write, options)?.0)
}

/// Rewrite field 20 and return exact resource accounting.
pub fn rewrite_chart_legend_visibility_with_report(
    source: &[u8],
    write: ChartLegendVisibilityWrite,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_chart_legend_visibility_rewrite(source, write, options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report;
    Ok((output.into_output(), report))
}

/// Compatibility spelling for callers using the shorter chart-legend name.
pub fn rewrite_chart_legend(
    source: &[u8],
    write: ChartLegendVisibilityWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_chart_legend_visibility(source, write, options)
}

fn normalize_write(
    current: ChartLegendVisibilitySnapshot<'_>,
    write: ChartLegendVisibilityWrite,
) -> NormalizedWrite {
    if write.clear {
        return if current.visible.is_some() {
            NormalizedWrite::Clear
        } else {
            NormalizedWrite::Preserve
        };
    }
    let Some(desired) = write.visible else {
        return NormalizedWrite::Preserve;
    };
    if current.visible == Some(desired)
        || (current.visible.is_none() && !desired && !write.force_presence)
    {
        NormalizedWrite::Preserve
    } else {
        NormalizedWrite::Replace(desired)
    }
}

fn candidate_matches(snapshot: ChartLegendVisibilitySnapshot<'_>, update: NormalizedWrite) -> bool {
    match update {
        NormalizedWrite::Preserve => true,
        NormalizedWrite::Replace(value) => snapshot.visible == Some(value),
        NormalizedWrite::Clear => snapshot.visible.is_none(),
    }
}

fn candidate_field_count(
    source_fields: usize,
    current: Option<bool>,
    update: NormalizedWrite,
) -> Result<usize, DecodeError> {
    let removed = usize::from(current.is_some() && !matches!(update, NormalizedWrite::Preserve));
    let added = usize::from(matches!(update, NormalizedWrite::Replace(_)));
    source_fields
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))
}

fn decode_with_budget<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartLegendVisibilitySnapshot<'source>, DecodeError> {
    budget.message(source.len(), ROOT_DEPTH)?;
    let strict = preflight_chart_legend(source, budget)?;
    // The generated lazy view is an independent semantic check. Unknown
    // fields are skipped by Buffa and remain exclusively source-owned.
    budget.buffa(source.len())?;
    let view: projection::ChartLegendArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    if view.tschchartinfodefaultshowlegend != strict.visible {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .unwrap_or(MAX_AUTOMATIC_LIMIT)
        .min(MAX_AUTOMATIC_LIMIT);
    if options.max_message_bytes > hard_maximum {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: hard_maximum,
        }));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::resource(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.max_fields > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Fields {
            observed: options.max_fields,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_work_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Work {
            observed: options.max_work_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_output_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Output {
            observed: options.max_output_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_allocations > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Allocations {
            observed: options.max_allocations,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_retained_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Retained {
            observed: options.max_retained_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_scratch_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Scratch {
            observed: options.max_scratch_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if source.len() > options.max_retained_bytes {
        return Err(DecodeError::resource(DecodeLimit::Retained {
            observed: source.len(),
            maximum: options.max_retained_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION_LIMIT,
        }));
    }
    Ok(())
}

#[derive(Debug)]
struct Budget {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    options: DecodeOptions,
}

impl Budget {
    const fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            source_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            options,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(|| field_size_overflow(self.options.max_fields))?;
        if observed > self.options.max_fields {
            return Err(DecodeError::resource(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        self.observe_depth(depth)?;
        let cost = bytes
            .checked_mul(2)
            .ok_or_else(|| work_size_overflow(self.options.max_work_bytes))?;
        self.charge_work(cost)
    }

    fn buffa(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.charge_work(bytes)
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .work_bytes
            .checked_add(amount)
            .ok_or_else(|| work_size_overflow(self.options.max_work_bytes))?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::resource(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::resource(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: 0,
            retained_bytes: self.source_bytes,
            scratch_bytes: 0,
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct FieldSpan {
    number: u32,
    start: usize,
    end: usize,
}

#[derive(Clone, Copy, Debug)]
struct StrictSnapshot<'source> {
    visible: Option<bool>,
    raw: &'source [u8],
}

fn preflight_chart_legend<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<ChartLegendVisibilitySnapshot<'source>, DecodeError> {
    let mut visible = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, budget, ROOT_DEPTH)? {
        if field.number == CHART_LEGEND_VISIBLE_FIELD {
            if visible.is_some() {
                return Err(DecodeError::duplicate());
            }
            visible = Some(require_canonical_bool(field.varint()?)?);
        }
    }
    let strict = StrictSnapshot {
        visible,
        raw: source,
    };
    Ok(ChartLegendVisibilitySnapshot {
        visible: strict.visible,
        raw: strict.raw,
    })
}

fn measure_rewrite_output(
    source: &[u8],
    update: NormalizedWrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RewriteMeasure, DecodeError> {
    let mut output_bytes = 0usize;
    let mut saw_selected = false;
    let mut max_depth = ROOT_DEPTH;
    visit_field_spans(source, budget, |span| {
        let replacement = if span.number == CHART_LEGEND_VISIBLE_FIELD {
            saw_selected = true;
            match update {
                NormalizedWrite::Preserve => span.end - span.start,
                NormalizedWrite::Replace(value) => {
                    varint_field_len(CHART_LEGEND_VISIBLE_FIELD, u64::from(value))
                },
                NormalizedWrite::Clear => 0,
            }
        } else {
            span.end - span.start
        };
        output_bytes = output_bytes
            .checked_add(replacement)
            .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
        Ok(())
    })?;
    if !saw_selected && let NormalizedWrite::Replace(value) = update {
        output_bytes = output_bytes
            .checked_add(varint_field_len(
                CHART_LEGEND_VISIBLE_FIELD,
                u64::from(value),
            ))
            .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
    }
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::resource(DecodeLimit::Output {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    // The selected payload is scalar-only. Unknown groups are parsed by the
    // visitor and therefore still contribute to the observed depth.
    max_depth = budget.max_depth.max(max_depth);
    Ok(RewriteMeasure {
        output_bytes,
        max_depth,
    })
}

#[derive(Clone, Copy, Debug)]
struct RewriteMeasure {
    output_bytes: usize,
    max_depth: u32,
}

fn visit_field_spans(
    source: &[u8],
    budget: &mut Budget,
    mut visit: impl FnMut(FieldSpan) -> Result<(), DecodeError>,
) -> Result<(), DecodeError> {
    budget.message(source.len(), ROOT_DEPTH)?;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, budget, ROOT_DEPTH)?;
        let end = source.len() - remaining.len();
        match item {
            Some(ParseItem::Field(field)) => visit(FieldSpan {
                number: field.number,
                start,
                end,
            })?,
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => break,
        }
    }
    Ok(())
}

fn emit_rewrite(
    source: &[u8],
    update: NormalizedWrite,
    output_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(output_bytes)?;
    let mut saw_selected = false;
    let mut budget = Budget::new(
        source,
        DecodeOptions::new(
            MAX_AUTOMATIC_LIMIT,
            usize::MAX,
            usize::MAX,
            MAX_RECURSION_LIMIT,
        ),
    );
    visit_field_spans(source, &mut budget, |span| {
        let raw = &source[span.start..span.end];
        if span.number != CHART_LEGEND_VISIBLE_FIELD {
            output.extend_from_slice(raw);
            return Ok(());
        }
        saw_selected = true;
        match update {
            NormalizedWrite::Preserve => output.extend_from_slice(raw),
            NormalizedWrite::Replace(value) => {
                append_varint_field(&mut output, CHART_LEGEND_VISIBLE_FIELD, u64::from(value));
            },
            NormalizedWrite::Clear => {},
        }
        Ok(())
    })?;
    if !saw_selected && let NormalizedWrite::Replace(value) = update {
        append_varint_field(&mut output, CHART_LEGEND_VISIBLE_FIELD, u64::from(value));
    }
    if output.len() != output_bytes {
        return Err(DecodeError::projection());
    }
    Ok(output)
}

fn candidate_work(output_bytes: usize) -> Result<usize, DecodeError> {
    output_bytes
        .checked_mul(3)
        .ok_or_else(|| work_size_overflow(MAX_AUTOMATIC_LIMIT))
}

fn check_decode_report_limits(
    report: DecodeReport,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if report.retained_bytes > options.max_retained_bytes {
        return Err(DecodeError::resource(DecodeLimit::Retained {
            observed: report.retained_bytes,
            maximum: options.max_retained_bytes,
        }));
    }
    Ok(())
}

fn check_option_execution_limits(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    check_execution_limits(
        requirements,
        RewriteExecutionLimits {
            output_bytes: options.max_output_bytes,
            fields: options.max_fields,
            work_bytes: options.max_work_bytes,
            max_depth: options.recursion_limit,
            allocations: options.max_allocations,
            retained_bytes: options.max_retained_bytes,
            scratch_bytes: options.max_scratch_bytes,
        },
    )
}

fn check_execution_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    let checks = [
        (
            requirements.output_bytes,
            limits.output_bytes,
            DecodeLimit::Output {
                observed: requirements.output_bytes,
                maximum: limits.output_bytes,
            },
        ),
        (
            requirements.fields,
            limits.fields,
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: limits.fields,
            },
        ),
        (
            requirements.work_bytes,
            limits.work_bytes,
            DecodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: limits.work_bytes,
            },
        ),
        (
            requirements.allocations,
            limits.allocations,
            DecodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: limits.allocations,
            },
        ),
        (
            requirements.retained_bytes,
            limits.retained_bytes,
            DecodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: limits.retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            limits.scratch_bytes,
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: limits.scratch_bytes,
            },
        ),
    ];
    for (observed, maximum, limit) in checks {
        if observed > maximum {
            return Err(DecodeError::resource(limit));
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::resource(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    Ok(())
}

fn field_size_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::Fields {
        observed: maximum.saturating_add(1),
        maximum,
    })
}

fn work_size_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::Work {
        observed: maximum.saturating_add(1),
        maximum,
    })
}

fn output_size_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::Output {
        observed: maximum.saturating_add(1),
        maximum,
    })
}

fn retained_size_overflow(maximum: usize) -> DecodeError {
    DecodeError::resource(DecodeLimit::Retained {
        observed: maximum.saturating_add(1),
        maximum,
    })
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_allocation_error| DecodeError::allocation(amount))?;
    if output.capacity() != amount {
        return Err(DecodeError::allocation(amount));
    }
    Ok(output)
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(output, u64::from(number) << 3);
    append_varint(output, value);
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

#[derive(Clone, Copy, Debug)]
enum StrictValue {
    Varint(u64),
    Fixed64,
    LengthDelimited,
    Group,
    Fixed32,
}

#[derive(Clone, Copy, Debug)]
struct StrictField {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue,
}

impl StrictField {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if self.wire_type != expected {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: expected as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Varint)?;
        let StrictValue::Varint(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem {
    Field(StrictField),
    EndGroup(u32),
}

fn next_strict_field(
    source: &mut &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<StrictField>, DecodeError> {
    match parse_strict_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field(
    source: &mut &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    let (encoded_tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    budget.field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            StrictValue::Varint(value)
        },
        buffa::encoding::WireType::Fixed64 => {
            take_exact(source, 8)?;
            StrictValue::Fixed64
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            take_exact(source, length)?;
            StrictValue::LengthDelimited
        },
        buffa::encoding::WireType::StartGroup => {
            let child_depth = depth.checked_add(1).ok_or_else(|| {
                DecodeError::resource(DecodeLimit::Nesting {
                    observed: u32::MAX,
                    maximum: budget.options.recursion_limit,
                })
            })?;
            budget.observe_depth(child_depth)?;
            skip_strict_group(source, field_number, child_depth, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            StrictValue::Fixed32
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
    })))
}

fn skip_strict_group(
    source: &mut &[u8],
    expected_field_number: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    reason = "Focused codec tests use explicit fixtures and assertions."
)]
mod tests {
    use super::{
        ChartLegendVisibilityWrite, DecodeLimit, DecodeOptions, WireResourceLimit,
        decode_chart_legend_visibility, prepare_chart_legend_visibility_rewrite,
        rewrite_chart_legend_visibility, rewrite_chart_legend_visibility_with_report,
    };

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return output;
            }
        }
    }

    fn field_varint(number: u32, value: u64) -> Vec<u8> {
        [varint(u64::from(number) << 3), varint(value)].concat()
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    #[test]
    fn absent_legend_is_false_without_allocating_or_rewriting_source() {
        let source = field_varint(4000, 7);
        let snapshot = decode_chart_legend_visibility(&source, options(&source)).expect("legend");
        assert_eq!(snapshot.visible(), None);
        assert!(!snapshot.is_visible());
        assert_eq!(snapshot.show_legend(), None);
        assert_eq!(snapshot.raw(), source.as_slice());
        let (output, report) = rewrite_chart_legend_visibility_with_report(
            &source,
            ChartLegendVisibilityWrite::new(false),
            options(&source),
        )
        .expect("semantic no-op");
        assert_eq!(output, source);
        assert!(!report.changed());
        assert_eq!(report.allocations(), 1);
    }

    #[test]
    fn explicit_values_decode_and_preserve_presence() {
        let hidden = field_varint(20, 0);
        let hidden_snapshot =
            decode_chart_legend_visibility(&hidden, options(&hidden)).expect("hidden");
        assert_eq!(hidden_snapshot.visible(), Some(false));
        assert!(!hidden_snapshot.is_visible());

        let visible = field_varint(20, 1);
        let visible_snapshot =
            decode_chart_legend_visibility(&visible, options(&visible)).expect("visible");
        assert_eq!(visible_snapshot.visible(), Some(true));
        assert!(visible_snapshot.is_visible());
    }

    #[test]
    fn prost_chart_non_style_archive_is_the_field_20_oracle() {
        use crate::tsch::generated::ChartNonStyleArchive;
        use prost::Message as _;

        for expected in [None, Some(false), Some(true)] {
            let native = ChartNonStyleArchive {
                tschchartinfodefaultshowlegend: expected,
                ..Default::default()
            };
            let source = native.encode_to_vec();
            let expected_wire =
                expected.map_or_else(Vec::new, |value| field_varint(20, value.into()));
            assert_eq!(source, expected_wire);

            let snapshot =
                decode_chart_legend_visibility(&source, options(&source)).expect("oracle");
            assert_eq!(snapshot.visible(), expected);
            assert_eq!(snapshot.is_visible(), expected.unwrap_or(false));
            let decoded = ChartNonStyleArchive::decode(source.as_slice()).expect("prost oracle");
            assert_eq!(decoded.tschchartinfodefaultshowlegend, expected);
        }
    }

    #[test]
    fn noncanonical_field_key_is_rejected() {
        let source = [0xa0, 0x81, 0x00, 0x01];
        let error = decode_chart_legend_visibility(&source, options(&source))
            .expect_err("overlong field key");
        assert_eq!(error.noncanonical_reason(), Some("protobuf field key"));
    }

    #[test]
    fn zero_field_number_is_rejected() {
        let source = [0x00];
        let error = decode_chart_legend_visibility(&source, options(&source))
            .expect_err("field number zero");
        assert_eq!(error.to_string(), "invalid field number");
    }

    #[test]
    fn truncated_length_delimited_value_is_rejected() {
        let source = [0xa2, 0x01, 0x04, 0xde, 0xad];
        let error = decode_chart_legend_visibility(&source, options(&source))
            .expect_err("truncated length-delimited value");
        assert_eq!(error.to_string(), "unexpected end of buffer");
    }

    #[test]
    fn unterminated_and_mismatched_groups_are_rejected() {
        let unterminated = [0x53];
        let error = decode_chart_legend_visibility(&unterminated, options(&unterminated))
            .expect_err("unterminated group");
        assert_eq!(error.to_string(), "unexpected end of buffer");

        let mismatched = [0x53, 0x5c];
        let error = decode_chart_legend_visibility(&mismatched, options(&mismatched))
            .expect_err("mismatched group");
        assert_eq!(error.to_string(), "invalid end-group tag: field number 11");
    }

    #[test]
    fn duplicate_wrong_wire_noncanonical_and_truncated_fields_fail_closed() {
        let duplicate = [field_varint(20, 1), field_varint(20, 0)].concat();
        let error = decode_chart_legend_visibility(&duplicate, options(&duplicate))
            .expect_err("duplicate field");
        assert!(error.duplicate_singular_field().is_some());

        for source in [
            vec![0xa2, 0x01, 0x01, 0x00], // field 20, length-delimited wire
            vec![0xa0, 0x81, 0x00, 0x01], // overlong field key
            vec![0xa0, 0x01],             // truncated bool value
            vec![0xa0, 0x01, 0x02],       // noncanonical bool value
        ] {
            assert!(
                decode_chart_legend_visibility(&source, options(&source)).is_err(),
                "{source:?}"
            );
        }
    }

    #[test]
    fn unknown_fields_and_groups_are_checked_preserved_and_depth_limited() {
        let mut source = vec![0x53]; // field 10 start group
        source.extend(field_varint(4000, 7));
        source.push(0x54); // field 10 end group
        source.extend(field_varint(20, 1));
        let snapshot = decode_chart_legend_visibility(&source, options(&source)).expect("group");
        assert!(snapshot.is_visible());

        let rewritten = rewrite_chart_legend_visibility(
            &source,
            ChartLegendVisibilityWrite::new(false),
            DecodeOptions::new(source.len(), 16, source.len() * 16, 2)
                .with_max_output_bytes(source.len())
                .with_max_retained_bytes(source.len() * 2),
        )
        .expect("preserve group");
        assert_eq!(&rewritten[..source.len() - 2], &source[..source.len() - 2]);
        assert_eq!(rewritten.last(), Some(&0x00));

        let error = decode_chart_legend_visibility(
            &source,
            DecodeOptions::new(source.len(), 16, source.len() * 8, 1),
        )
        .expect_err("depth cap");
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn rewrite_replaces_in_place_appends_missing_and_supports_clear() {
        let source = [
            field_varint(4000, 7),
            field_varint(20, 0),
            field_varint(4001, 8),
        ]
        .concat();
        let rewritten = rewrite_chart_legend_visibility(
            &source,
            ChartLegendVisibilityWrite::new(true),
            DecodeOptions::new(1024, 128, 8192, 8).with_max_output_bytes(1024),
        )
        .expect("replace");
        assert_eq!(
            rewritten,
            [
                field_varint(4000, 7),
                field_varint(20, 1),
                field_varint(4001, 8)
            ]
            .concat()
        );

        let absent = field_varint(4000, 7);
        let appended = rewrite_chart_legend_visibility(
            &absent,
            ChartLegendVisibilityWrite::new(true),
            DecodeOptions::new(1024, 128, 8192, 8).with_max_output_bytes(1024),
        )
        .expect("append");
        assert!(appended.starts_with(&absent));
        assert!(appended.ends_with(&field_varint(20, 1)));

        let effective_false = rewrite_chart_legend_visibility(
            &absent,
            ChartLegendVisibilityWrite::new(false),
            DecodeOptions::new(1024, 128, 8192, 8).with_max_output_bytes(1024),
        )
        .expect("preserve effective false");
        assert_eq!(effective_false, absent);

        let explicit_false = rewrite_chart_legend_visibility(
            &absent,
            ChartLegendVisibilityWrite::explicit(Some(false)),
            DecodeOptions::new(1024, 128, 8192, 8).with_max_output_bytes(1024),
        )
        .expect("append explicit false");
        assert!(explicit_false.starts_with(&absent));
        assert!(explicit_false.ends_with(&field_varint(20, 0)));
        assert_eq!(
            decode_chart_legend_visibility(&explicit_false, options(&explicit_false))
                .expect("decode explicit false")
                .visible(),
            Some(false)
        );

        let cleared = rewrite_chart_legend_visibility(
            &source,
            ChartLegendVisibilityWrite::clear(),
            DecodeOptions::new(1024, 128, 8192, 8).with_max_output_bytes(1024),
        )
        .expect("clear");
        assert_eq!(
            cleared,
            [field_varint(4000, 7), field_varint(4001, 8)].concat()
        );
    }

    #[test]
    fn prepared_rewrite_reports_requirements_before_allocation() {
        let source = [field_varint(4000, 7), field_varint(20, 0)].concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8).with_max_output_bytes(1024);
        let prepared = prepare_chart_legend_visibility_rewrite(
            &source,
            ChartLegendVisibilityWrite::new(true),
            options,
        )
        .expect("prepare");
        let requirements = prepared.execution_requirements();
        assert_eq!(requirements.output_bytes, source.len());
        assert_eq!(requirements.allocations, 1);
        assert_eq!(requirements.retained_bytes, source.len() * 2);
        let output = prepared.execute(requirements.exact()).expect("execute");
        assert!(output.report().changed());
        assert_eq!(
            output.output(),
            [field_varint(4000, 7), field_varint(20, 1)].concat()
        );
    }

    #[test]
    fn work_and_output_limits_are_inclusive_and_atomic() {
        let source = [field_varint(20, 0), field_varint(4000, 7)].concat();
        let target = rewrite_chart_legend_visibility_with_report(
            &source,
            ChartLegendVisibilityWrite::new(true),
            DecodeOptions::new(1024, MAX_LIMIT, MAX_LIMIT, 8)
                .with_max_output_bytes(1024)
                .with_max_retained_bytes(1024),
        )
        .expect("unconstrained");
        let output = target.0;
        let report = target.1;
        let exact = DecodeOptions::new(1024, MAX_LIMIT, report.work_bytes(), 8)
            .with_max_output_bytes(output.len())
            .with_max_retained_bytes(source.len() + output.len());
        rewrite_chart_legend_visibility(&source, ChartLegendVisibilityWrite::new(true), exact)
            .expect("inclusive work");
        let below = DecodeOptions::new(1024, MAX_LIMIT, report.work_bytes() - 1, 8)
            .with_max_output_bytes(output.len())
            .with_max_retained_bytes(source.len() + output.len());
        let error =
            rewrite_chart_legend_visibility(&source, ChartLegendVisibilityWrite::new(true), below)
                .expect_err("work cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        assert_eq!(
            source,
            [field_varint(20, 0), field_varint(4000, 7)].concat()
        );
    }

    #[test]
    fn allocation_reservation_is_bounded() {
        let error = super::reserve_output(usize::MAX).expect_err("allocation cap");
        assert_eq!(error.allocation_amount(), Some(usize::MAX));
    }

    const MAX_LIMIT: usize = 64 * 1024 * 1024;
}
