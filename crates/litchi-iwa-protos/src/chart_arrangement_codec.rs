//! Strict borrowed projection for the shared iWork chart Arrange state.
//!
//! The selected path is `TSCH.ChartDrawableArchive.super` (field 1) to the
//! two optional `TSD.DrawableArchive` controls `locked` (field 5) and
//! `aspect_ratio_locked` (field 7). The chart envelope, every unrelated
//! drawable field, and all unknown source bytes remain caller-owned. This
//! module performs bounded strict preflight before forcing a private Buffa
//! lazy view, and rewrites only the selected source spans.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire preflight intentionally precedes the private view."
)]

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_chart_arrangement_generated::LitchiIwaProjection as projection;

const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const DRAWABLE_ASPECT_RATIO_LOCKED_FIELD: u32 = 7;
const ROOT_DEPTH: u32 = 1;
const DRAWABLE_DEPTH: u32 = 2;
const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;

const CHART_DRAWABLE_SUPER_NAME: &str = "TSCH.ChartDrawableArchive.super";
const DRAWABLE_LOCKED_NAME: &str = "TSD.DrawableArchive.locked";
const DRAWABLE_ASPECT_RATIO_LOCKED_NAME: &str = "TSD.DrawableArchive.aspect_ratio_locked";

/// Finite limits for one generated chart-drawable Arrange payload.
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

    /// Build a conservative finite policy from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        let fields = bytes
            .checked_mul(8)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .clamp(1, MAX_AUTOMATIC_LIMIT);
        let work = bytes
            .checked_mul(32)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .clamp(1, MAX_AUTOMATIC_LIMIT);
        let output = bytes
            .checked_mul(2)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .clamp(1, MAX_AUTOMATIC_LIMIT);
        let retained = bytes
            .checked_mul(4)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT)
            .max(bytes);
        let scratch = bytes
            .checked_mul(4)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT)
            .max(bytes);
        Self::new(bytes, fields, work, 8)
            .with_max_output_bytes(output)
            .with_max_allocations(1)
            .with_max_retained_bytes(retained)
            .with_max_scratch_bytes(scratch)
    }

    /// Replace the source/input-byte ceiling.
    #[must_use]
    pub const fn with_max_message_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Compatibility spelling for callers that use input terminology.
    #[must_use]
    pub const fn with_max_input_bytes(self, maximum: usize) -> Self {
        self.with_max_message_bytes(maximum)
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

/// Borrowed semantic facts from one chart drawable Arrange payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartArrangementSnapshot<'source> {
    has_super: bool,
    locked: Option<bool>,
    constrain_proportions: Option<bool>,
    drawable_raw: Option<&'source [u8]>,
    raw: &'source [u8],
}

impl<'source> ChartArrangementSnapshot<'source> {
    /// Return whether the optional chart drawable `super` envelope was present.
    #[must_use]
    pub const fn has_drawable(self) -> bool {
        self.has_super
    }

    /// Return the optional native lock value, preserving proto2 presence.
    #[must_use]
    pub const fn locked(self) -> Option<bool> {
        self.locked
    }

    /// Return the effective lock value, treating absence as unlocked.
    #[must_use]
    pub const fn is_locked(self) -> bool {
        matches!(self.locked, Some(true))
    }

    /// Return the optional native aspect-ratio constraint, preserving presence.
    #[must_use]
    pub const fn constrain_proportions(self) -> Option<bool> {
        self.constrain_proportions
    }

    /// Return the effective aspect-ratio constraint, treating absence as false.
    #[must_use]
    pub const fn is_constrained(self) -> bool {
        matches!(self.constrain_proportions, Some(true))
    }

    /// Alias for [`Self::constrain_proportions`].
    #[must_use]
    pub const fn aspect_ratio_locked(self) -> Option<bool> {
        self.constrain_proportions
    }

    /// Return the exact caller-owned bytes used by this snapshot.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Requested semantic values for a chart Arrange-state rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartArrangementWrite {
    locked: Option<bool>,
    constrain_proportions: Option<bool>,
    preserve_absent_false: bool,
}

impl ChartArrangementWrite {
    /// Build a semantic Arrange-state request.
    ///
    /// An absent false field remains absent, while an already-present field is
    /// rewritten to the requested explicit value. Use [`Self::explicit`] when
    /// exact field presence is required.
    #[must_use]
    pub const fn new(locked: bool, constrain_proportions: bool) -> Self {
        Self {
            locked: Some(locked),
            constrain_proportions: Some(constrain_proportions),
            preserve_absent_false: true,
        }
    }

    /// Build an exact presence-aware request. `None` removes its field.
    #[must_use]
    pub const fn explicit(locked: Option<bool>, constrain_proportions: Option<bool>) -> Self {
        Self {
            locked,
            constrain_proportions,
            preserve_absent_false: false,
        }
    }

    /// Remove both selected controls while preserving the drawable envelope.
    #[must_use]
    pub const fn clear() -> Self {
        Self::explicit(None, None)
    }

    /// Return the requested lock value, including requested presence.
    #[must_use]
    pub const fn locked(self) -> Option<bool> {
        self.locked
    }

    /// Return the requested aspect-ratio value, including requested presence.
    #[must_use]
    pub const fn constrain_proportions(self) -> Option<bool> {
        self.constrain_proportions
    }

    /// Return the requested effective lock value.
    #[must_use]
    pub const fn is_locked(self) -> bool {
        matches!(self.locked, Some(true))
    }

    /// Return the requested effective aspect-ratio value.
    #[must_use]
    pub const fn is_constrained(self) -> bool {
        matches!(self.constrain_proportions, Some(true))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NormalizedWrite {
    Preserve,
    Replace(bool),
    Clear,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct NormalizedArrangementWrite {
    locked: NormalizedWrite,
    constrain_proportions: NormalizedWrite,
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
    /// Aggregate source, emission, and candidate field visits.
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

    /// Alias for [`Self::into_output`].
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
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

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
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
                "chart-arrangement projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "chart-arrangement projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "chart-arrangement projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum }) => write!(
                formatter,
                "chart-arrangement projection output has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "chart-arrangement projection nesting {observed} exceeds maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "chart-arrangement rewrite requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "chart-arrangement rewrite retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "chart-arrangement rewrite requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate chart-arrangement output for {amount} bytes"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "chart-arrangement strict preflight disagrees with the Buffa projection",
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

/// One source-preserving chart-arrangement rewrite prepared without allocating
/// candidate output.
#[derive(Debug, Clone, Copy)]
pub struct PreparedChartArrangementRewrite<'source> {
    source: &'source [u8],
    update: NormalizedArrangementWrite,
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    emitter_fields: usize,
    emitter_work: usize,
    candidate_fields: usize,
    candidate_work: usize,
}

impl PreparedChartArrangementRewrite<'_> {
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
        let mut output = reserve_output(self.requirements.output_bytes)?;
        if self.requirements.output_bytes == self.source.len() && !updates_changed(self.update) {
            output.extend_from_slice(self.source);
        } else {
            let report = emit_rewrite(self.source, self.update, self.options, &mut output)?;
            if report.fields() != self.emitter_fields || report.work_bytes() != self.emitter_work {
                return Err(DecodeError::projection());
            }
        }
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::projection());
        }

        if updates_changed(self.update) {
            let readback_options = DecodeOptions {
                max_message_bytes: output.len().max(1),
                max_output_bytes: self.options.max_output_bytes.max(output.len()),
                ..self.options
            };
            validate_decode_input(&output, readback_options)?;
            let mut readback_budget = Budget::new(&output, readback_options);
            let readback = decode_with_budget(&output, readback_options, &mut readback_budget)?;
            let report = readback_budget.report();
            if report.fields() != self.candidate_fields
                || report.work_bytes() != self.candidate_work
            {
                return Err(DecodeError::projection());
            }
            if !candidate_matches(readback, self.update) {
                return Err(DecodeError::projection());
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
                changed: updates_changed(self.update),
            },
        })
    }
}

/// Decode a chart drawable Arrange payload.
pub fn decode_chart_arrangement<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ChartArrangementSnapshot<'source>, DecodeError> {
    Ok(decode_chart_arrangement_with_report(source, options)?.0)
}

/// Decode a chart drawable Arrange payload and return exact accounting.
pub fn decode_chart_arrangement_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(ChartArrangementSnapshot<'source>, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = decode_with_budget(source, options, &mut budget)?;
    let report = budget.report();
    check_decode_report_limits(report, options)?;
    Ok((snapshot, report))
}

/// Prepare a source-preserving chart Arrange rewrite without allocating output.
pub fn prepare_chart_arrangement_rewrite<'source>(
    source: &'source [u8],
    write: ChartArrangementWrite,
    options: DecodeOptions,
) -> Result<PreparedChartArrangementRewrite<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let current = decode_with_budget(source, options, &mut budget)?;
    let source_report = budget.report();
    let update = normalize_write(current, write);
    let measure = measure_rewrite(source, update, options, &mut budget)?;

    if measure.changed {
        // The emitter scans the root and measures then emits the selected
        // nested drawable. Candidate readback performs both strict and Buffa
        // traversals. Charge those exact visits before reserving the sole
        // output allocation so limits cannot fail after publication.
        budget.charge_fields(measure.emitter_fields)?;
        budget.charge_fields(measure.candidate_fields)?;
        budget.charge_work(measure.emitter_work)?;
        budget.charge_work(measure.candidate_work)?;
    }

    let retained_bytes = source
        .len()
        .checked_add(measure.output_bytes)
        .ok_or_else(|| retained_size_overflow(options.max_retained_bytes))?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: measure.output_bytes,
        fields: budget.fields,
        work_bytes: budget.work_bytes,
        max_depth: budget.max_depth.max(measure.max_depth),
        allocations: usize::from(measure.output_bytes != 0),
        retained_bytes,
        scratch_bytes: measure.output_bytes,
    };
    check_option_execution_limits(requirements, options)?;
    Ok(PreparedChartArrangementRewrite {
        source,
        update,
        options,
        source_report,
        requirements,
        emitter_fields: measure.emitter_fields,
        emitter_work: measure.emitter_work,
        candidate_fields: measure.candidate_fields,
        candidate_work: measure.candidate_work,
    })
}

/// Rewrite both selected chart Arrange controls.
pub fn rewrite_chart_arrangement(
    source: &[u8],
    write: ChartArrangementWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_chart_arrangement_with_report(source, write, options)?.0)
}

/// Rewrite both selected controls and return exact accounting.
pub fn rewrite_chart_arrangement_with_report(
    source: &[u8],
    write: ChartArrangementWrite,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_chart_arrangement_rewrite(source, write, options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report;
    Ok((output.into_output(), report))
}

fn normalize_write(
    current: ChartArrangementSnapshot<'_>,
    write: ChartArrangementWrite,
) -> NormalizedArrangementWrite {
    NormalizedArrangementWrite {
        locked: normalize_value(current.locked, write.locked, write.preserve_absent_false),
        constrain_proportions: normalize_value(
            current.constrain_proportions,
            write.constrain_proportions,
            write.preserve_absent_false,
        ),
    }
}

fn normalize_value(
    current: Option<bool>,
    desired: Option<bool>,
    preserve_absent_false: bool,
) -> NormalizedWrite {
    match desired {
        None => {
            if current.is_some() {
                NormalizedWrite::Clear
            } else {
                NormalizedWrite::Preserve
            }
        },
        Some(value) if current == Some(value) => NormalizedWrite::Preserve,
        Some(false) if current.is_none() && preserve_absent_false => NormalizedWrite::Preserve,
        Some(value) => NormalizedWrite::Replace(value),
    }
}

fn updates_changed(update: NormalizedArrangementWrite) -> bool {
    !matches!(update.locked, NormalizedWrite::Preserve)
        || !matches!(update.constrain_proportions, NormalizedWrite::Preserve)
}

fn candidate_matches(
    snapshot: ChartArrangementSnapshot<'_>,
    update: NormalizedArrangementWrite,
) -> bool {
    matches_update(snapshot.locked, update.locked)
        && matches_update(snapshot.constrain_proportions, update.constrain_proportions)
}

fn matches_update(current: Option<bool>, update: NormalizedWrite) -> bool {
    match update {
        NormalizedWrite::Preserve => true,
        NormalizedWrite::Replace(value) => current == Some(value),
        NormalizedWrite::Clear => current.is_none(),
    }
}

fn decode_with_budget<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartArrangementSnapshot<'source>, DecodeError> {
    let strict = preflight_chart_arrangement(source, budget)?;

    // The generated lazy view is an independent semantic check. The strict
    // pass above remains authoritative for canonical wire, duplicate, and
    // resource policy; Buffa only confirms selected field interpretation.
    budget.buffa(source.len())?;
    let view: projection::ChartDrawableArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = view
        .super_
        .get()
        .map_err(DecodeError::from)?
        .map(|drawable| (drawable.locked, drawable.aspect_ratio_locked));
    if let Some((locked, constrain_proportions)) = projected {
        let raw = strict.drawable_raw.ok_or_else(DecodeError::projection)?;
        budget.buffa(raw.len())?;
        if !strict.has_super
            || strict.locked != locked
            || strict.constrain_proportions != constrain_proportions
        {
            return Err(DecodeError::projection());
        }
    } else if strict.has_super {
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

    fn charge_fields(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(amount)
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
struct StrictDrawableSnapshot<'source> {
    locked: Option<bool>,
    constrain_proportions: Option<bool>,
    raw: &'source [u8],
}

fn preflight_chart_arrangement<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<ChartArrangementSnapshot<'source>, DecodeError> {
    budget.message(source.len(), ROOT_DEPTH)?;
    let mut drawable = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, budget, ROOT_DEPTH)? {
        if field.number != CHART_DRAWABLE_SUPER_FIELD {
            continue;
        }
        if drawable.is_some() {
            return Err(DecodeError::duplicate(CHART_DRAWABLE_SUPER_NAME));
        }
        let raw = field.length_delimited()?;
        drawable = Some(preflight_drawable(raw, budget)?);
    }
    Ok(match drawable {
        Some(drawable) => ChartArrangementSnapshot {
            has_super: true,
            locked: drawable.locked,
            constrain_proportions: drawable.constrain_proportions,
            drawable_raw: Some(drawable.raw),
            raw: source,
        },
        None => ChartArrangementSnapshot {
            has_super: false,
            locked: None,
            constrain_proportions: None,
            drawable_raw: None,
            raw: source,
        },
    })
}

fn preflight_drawable<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictDrawableSnapshot<'source>, DecodeError> {
    budget.message(source.len(), DRAWABLE_DEPTH)?;
    let mut locked = None;
    let mut constrain_proportions = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, budget, DRAWABLE_DEPTH)? {
        match field.number {
            DRAWABLE_LOCKED_FIELD => {
                if locked.is_some() {
                    return Err(DecodeError::duplicate(DRAWABLE_LOCKED_NAME));
                }
                locked = Some(require_canonical_bool(field.varint()?)?);
            },
            DRAWABLE_ASPECT_RATIO_LOCKED_FIELD => {
                if constrain_proportions.is_some() {
                    return Err(DecodeError::duplicate(DRAWABLE_ASPECT_RATIO_LOCKED_NAME));
                }
                constrain_proportions = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(StrictDrawableSnapshot {
        locked,
        constrain_proportions,
        raw: source,
    })
}

#[derive(Clone, Copy, Debug)]
struct RewriteMeasure {
    output_bytes: usize,
    candidate_fields: usize,
    emitter_fields: usize,
    emitter_work: usize,
    candidate_work: usize,
    max_depth: u32,
    changed: bool,
}

#[derive(Clone, Copy, Debug)]
struct DrawableRewriteMeasure {
    output_bytes: usize,
    source_bytes: usize,
    source_fields: usize,
    candidate_fields: usize,
    changed: bool,
}

fn measure_rewrite(
    source: &[u8],
    update: NormalizedArrangementWrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RewriteMeasure, DecodeError> {
    let fields_before = budget.fields;
    budget.message(source.len(), ROOT_DEPTH)?;
    let mut remaining = source;
    let mut output_bytes = 0usize;
    let mut saw_super = false;
    let mut nested_measure = None;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, budget, ROOT_DEPTH)?;
        let end = source.len() - remaining.len();
        let Some(ParseItem::Field(field)) = item else {
            return Err(match item {
                Some(ParseItem::EndGroup(number)) => {
                    buffa::DecodeError::InvalidEndGroup(number).into()
                },
                None => buffa::DecodeError::UnexpectedEof.into(),
                Some(ParseItem::Field(_)) => DecodeError::projection(),
            });
        };
        if field.number != CHART_DRAWABLE_SUPER_FIELD {
            output_bytes = output_bytes
                .checked_add(end - start)
                .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
            continue;
        }
        if saw_super {
            return Err(DecodeError::duplicate(CHART_DRAWABLE_SUPER_NAME));
        }
        saw_super = true;
        let nested = field.length_delimited()?;
        let measured = measure_drawable(nested, update, options, budget)?;
        let replacement = if measured.changed {
            length_delimited_field_len(CHART_DRAWABLE_SUPER_FIELD, measured.output_bytes)?
        } else {
            end - start
        };
        output_bytes = output_bytes
            .checked_add(replacement)
            .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
        nested_measure = Some(measured);
    }

    let source_fields = budget
        .fields
        .checked_sub(fields_before)
        .ok_or_else(|| field_size_overflow(options.max_fields))?;
    let (candidate_fields, emitter_fields, emitter_work, candidate_nested_bytes, changed) =
        match nested_measure {
            Some(nested) => {
                let root_fields = source_fields
                    .checked_sub(nested.source_fields)
                    .ok_or_else(|| field_size_overflow(options.max_fields))?;
                let candidate_fields = root_fields
                    .checked_add(nested.candidate_fields)
                    .ok_or_else(|| field_size_overflow(options.max_fields))?;
                let emitter_fields = if nested.changed {
                    source_fields
                        .checked_add(nested.source_fields)
                        .ok_or_else(|| field_size_overflow(options.max_fields))?
                } else {
                    0
                };
                let emitter_work = if nested.changed {
                    nested
                        .source_bytes
                        .checked_mul(2)
                        .ok_or_else(|| work_size_overflow(options.max_work_bytes))?
                } else {
                    0
                };
                (
                    candidate_fields,
                    emitter_fields,
                    emitter_work,
                    nested.output_bytes,
                    nested.changed,
                )
            },
            None if updates_changed(update) => {
                let (nested_bytes, nested_fields) = measure_new_drawable(update)?;
                output_bytes = output_bytes
                    .checked_add(length_delimited_field_len(
                        CHART_DRAWABLE_SUPER_FIELD,
                        nested_bytes,
                    )?)
                    .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
                let candidate_fields = source_fields
                    .checked_add(1)
                    .and_then(|value| value.checked_add(nested_fields))
                    .ok_or_else(|| field_size_overflow(options.max_fields))?;
                (candidate_fields, source_fields, 0, nested_bytes, true)
            },
            None => (source_fields, 0, 0, 0, false),
        };
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::resource(DecodeLimit::Output {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let candidate_work = if changed {
        output_bytes
            .checked_add(candidate_nested_bytes)
            .and_then(|bytes| bytes.checked_mul(3))
            .ok_or_else(|| work_size_overflow(options.max_work_bytes))?
    } else {
        0
    };
    Ok(RewriteMeasure {
        output_bytes,
        candidate_fields,
        emitter_fields,
        emitter_work,
        candidate_work,
        max_depth: budget
            .max_depth
            .max(if changed { DRAWABLE_DEPTH } else { ROOT_DEPTH }),
        changed,
    })
}

fn measure_drawable(
    source: &[u8],
    update: NormalizedArrangementWrite,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<DrawableRewriteMeasure, DecodeError> {
    let fields_before = budget.fields;
    budget.message(source.len(), DRAWABLE_DEPTH)?;
    let mut remaining = source;
    let mut output_bytes = 0usize;
    let mut locked = None;
    let mut constrain_proportions = None;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, budget, DRAWABLE_DEPTH)?;
        let end = source.len() - remaining.len();
        let Some(ParseItem::Field(field)) = item else {
            return Err(match item {
                Some(ParseItem::EndGroup(number)) => {
                    buffa::DecodeError::InvalidEndGroup(number).into()
                },
                None => buffa::DecodeError::UnexpectedEof.into(),
                Some(ParseItem::Field(_)) => DecodeError::projection(),
            });
        };
        let replacement = match field.number {
            DRAWABLE_LOCKED_FIELD => {
                if locked.is_some() {
                    return Err(DecodeError::duplicate(DRAWABLE_LOCKED_NAME));
                }
                locked = Some(require_canonical_bool(field.varint()?)?);
                replacement_len(locked, update.locked, DRAWABLE_LOCKED_FIELD, end - start)?
            },
            DRAWABLE_ASPECT_RATIO_LOCKED_FIELD => {
                if constrain_proportions.is_some() {
                    return Err(DecodeError::duplicate(DRAWABLE_ASPECT_RATIO_LOCKED_NAME));
                }
                constrain_proportions = Some(require_canonical_bool(field.varint()?)?);
                replacement_len(
                    constrain_proportions,
                    update.constrain_proportions,
                    DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
                    end - start,
                )?
            },
            _ => end - start,
        };
        output_bytes = output_bytes
            .checked_add(replacement)
            .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
    }
    let source_fields = budget
        .fields
        .checked_sub(fields_before)
        .ok_or_else(|| field_size_overflow(options.max_fields))?;
    let locked_update = normalized_append_len(locked, update.locked, DRAWABLE_LOCKED_FIELD);
    let constrain_update = normalized_append_len(
        constrain_proportions,
        update.constrain_proportions,
        DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
    );
    output_bytes = output_bytes
        .checked_add(locked_update)
        .and_then(|value| value.checked_add(constrain_update))
        .ok_or_else(|| output_size_overflow(options.max_output_bytes))?;
    let candidate_fields = adjusted_field_count(
        adjusted_field_count(source_fields, locked, update.locked)?,
        constrain_proportions,
        update.constrain_proportions,
    )?;
    let changed = updates_changed_for_values(locked, constrain_proportions, update);
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::resource(DecodeLimit::Output {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    Ok(DrawableRewriteMeasure {
        output_bytes,
        source_bytes: source.len(),
        source_fields,
        candidate_fields,
        changed,
    })
}

fn replacement_len(
    current: Option<bool>,
    update: NormalizedWrite,
    number: u32,
    original_len: usize,
) -> Result<usize, DecodeError> {
    Ok(match update {
        NormalizedWrite::Preserve => original_len,
        NormalizedWrite::Replace(value) => varint_field_len(number, u64::from(value)),
        NormalizedWrite::Clear => {
            if current.is_some() {
                0
            } else {
                original_len
            }
        },
    })
}

fn normalized_append_len(current: Option<bool>, update: NormalizedWrite, number: u32) -> usize {
    match (current, update) {
        (None, NormalizedWrite::Replace(value)) => varint_field_len(number, u64::from(value)),
        _ => 0,
    }
}

fn measure_new_drawable(update: NormalizedArrangementWrite) -> Result<(usize, usize), DecodeError> {
    let mut bytes = 0usize;
    let mut fields = 0usize;
    if let NormalizedWrite::Replace(value) = update.locked {
        bytes = bytes
            .checked_add(varint_field_len(DRAWABLE_LOCKED_FIELD, u64::from(value)))
            .ok_or_else(|| output_size_overflow(MAX_AUTOMATIC_LIMIT))?;
        fields += 1;
    }
    if let NormalizedWrite::Replace(value) = update.constrain_proportions {
        bytes = bytes
            .checked_add(varint_field_len(
                DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
                u64::from(value),
            ))
            .ok_or_else(|| output_size_overflow(MAX_AUTOMATIC_LIMIT))?;
        fields += 1;
    }
    Ok((bytes, fields))
}

fn adjusted_field_count(
    source_fields: usize,
    current: Option<bool>,
    update: NormalizedWrite,
) -> Result<usize, DecodeError> {
    let removed = usize::from(current.is_some() && matches!(update, NormalizedWrite::Clear));
    let added = usize::from(current.is_none() && matches!(update, NormalizedWrite::Replace(_)));
    source_fields
        .checked_sub(removed)
        .and_then(|value| value.checked_add(added))
        .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))
}

fn updates_changed_for_values(
    locked: Option<bool>,
    constrain_proportions: Option<bool>,
    update: NormalizedArrangementWrite,
) -> bool {
    !matches!(
        update_for_value(locked, update.locked),
        NormalizedWrite::Preserve
    ) || !matches!(
        update_for_value(constrain_proportions, update.constrain_proportions),
        NormalizedWrite::Preserve
    )
}

fn update_for_value(current: Option<bool>, update: NormalizedWrite) -> NormalizedWrite {
    match update {
        NormalizedWrite::Clear if current.is_none() => NormalizedWrite::Preserve,
        other => other,
    }
}

fn emit_rewrite(
    source: &[u8],
    update: NormalizedArrangementWrite,
    options: DecodeOptions,
    output: &mut Vec<u8>,
) -> Result<DecodeReport, DecodeError> {
    let mut budget = Budget::new(source, options);
    let mut remaining = source;
    let mut saw_super = false;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, &mut budget, ROOT_DEPTH)?;
        let end = source.len() - remaining.len();
        let Some(ParseItem::Field(field)) = item else {
            return Err(match item {
                Some(ParseItem::EndGroup(number)) => {
                    buffa::DecodeError::InvalidEndGroup(number).into()
                },
                None => buffa::DecodeError::UnexpectedEof.into(),
                Some(ParseItem::Field(_)) => DecodeError::projection(),
            });
        };
        if field.number != CHART_DRAWABLE_SUPER_FIELD {
            output.extend_from_slice(&source[start..end]);
            continue;
        }
        if saw_super {
            return Err(DecodeError::duplicate(CHART_DRAWABLE_SUPER_NAME));
        }
        saw_super = true;
        let nested = field.length_delimited()?;
        if updates_changed(update) {
            let measure = measure_drawable(
                nested,
                update,
                DecodeOptions::new(
                    MAX_AUTOMATIC_LIMIT,
                    MAX_AUTOMATIC_LIMIT,
                    MAX_AUTOMATIC_LIMIT,
                    MAX_RECURSION_LIMIT,
                ),
                &mut budget,
            )?;
            append_length_delimited_field(output, CHART_DRAWABLE_SUPER_FIELD, measure.output_bytes);
            emit_drawable(nested, update, &mut budget, output)?;
        } else {
            output.extend_from_slice(&source[start..end]);
        }
    }
    if !saw_super && updates_changed(update) {
        let (nested_bytes, _) = measure_new_drawable(update)?;
        append_length_delimited_field(output, CHART_DRAWABLE_SUPER_FIELD, nested_bytes);
        emit_new_drawable(update, output);
    }
    Ok(budget.report())
}

fn emit_drawable(
    source: &[u8],
    update: NormalizedArrangementWrite,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut remaining = source;
    let mut locked = None;
    let mut constrain_proportions = None;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, budget, DRAWABLE_DEPTH)?;
        let end = source.len() - remaining.len();
        let Some(ParseItem::Field(field)) = item else {
            return Err(match item {
                Some(ParseItem::EndGroup(number)) => {
                    buffa::DecodeError::InvalidEndGroup(number).into()
                },
                None => buffa::DecodeError::UnexpectedEof.into(),
                Some(ParseItem::Field(_)) => DecodeError::projection(),
            });
        };
        match field.number {
            DRAWABLE_LOCKED_FIELD => {
                if locked.is_some() {
                    return Err(DecodeError::duplicate(DRAWABLE_LOCKED_NAME));
                }
                locked = Some(require_canonical_bool(field.varint()?)?);
                emit_selected_bool(
                    output,
                    DRAWABLE_LOCKED_FIELD,
                    update.locked,
                    &source[start..end],
                );
            },
            DRAWABLE_ASPECT_RATIO_LOCKED_FIELD => {
                if constrain_proportions.is_some() {
                    return Err(DecodeError::duplicate(DRAWABLE_ASPECT_RATIO_LOCKED_NAME));
                }
                constrain_proportions = Some(require_canonical_bool(field.varint()?)?);
                emit_selected_bool(
                    output,
                    DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
                    update.constrain_proportions,
                    &source[start..end],
                );
            },
            _ => output.extend_from_slice(&source[start..end]),
        }
    }
    if let NormalizedWrite::Replace(value) = update.locked
        && locked.is_none()
    {
        append_varint_field(output, DRAWABLE_LOCKED_FIELD, u64::from(value));
    }
    if let NormalizedWrite::Replace(value) = update.constrain_proportions
        && constrain_proportions.is_none()
    {
        append_varint_field(output, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, u64::from(value));
    }
    Ok(())
}

fn emit_new_drawable(update: NormalizedArrangementWrite, output: &mut Vec<u8>) {
    if let NormalizedWrite::Replace(value) = update.locked {
        append_varint_field(output, DRAWABLE_LOCKED_FIELD, u64::from(value));
    }
    if let NormalizedWrite::Replace(value) = update.constrain_proportions {
        append_varint_field(output, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, u64::from(value));
    }
}

fn emit_selected_bool(output: &mut Vec<u8>, number: u32, update: NormalizedWrite, original: &[u8]) {
    match update {
        NormalizedWrite::Preserve => output.extend_from_slice(original),
        NormalizedWrite::Replace(value) => {
            append_varint_field(output, number, u64::from(value));
        },
        NormalizedWrite::Clear => {},
    }
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
        .map_err(|_error| DecodeError::allocation(amount))?;
    if output.capacity() != amount {
        return Err(DecodeError::allocation(amount));
    }
    Ok(output)
}

fn append_length_delimited_field(output: &mut Vec<u8>, number: u32, length: usize) {
    append_varint(output, u64::from(number) << 3 | 2);
    append_varint(output, length as u64);
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(output, u64::from(number) << 3);
    append_varint(output, value);
}

fn length_delimited_field_len(number: u32, length: usize) -> Result<usize, DecodeError> {
    let length =
        u64::try_from(length).map_err(|_error| output_size_overflow(MAX_AUTOMATIC_LIMIT))?;
    varint_len(u64::from(number) << 3 | 2)
        .checked_add(varint_len(length))
        .and_then(|value| value.checked_add(length as usize))
        .ok_or_else(|| output_size_overflow(MAX_AUTOMATIC_LIMIT))
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
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
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

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        let StrictValue::LengthDelimited(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
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
            let payload = take_exact(source, length)?;
            StrictValue::LengthDelimited(payload)
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
    reason = "Focused codec tests use explicit wire fixtures and assertions."
)]
mod tests {
    use super::{
        ChartArrangementWrite, DecodeLimit, DecodeOptions, WireResourceLimit,
        decode_chart_arrangement, prepare_chart_arrangement_rewrite, rewrite_chart_arrangement,
        rewrite_chart_arrangement_with_report,
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

    fn field_bytes(number: u32, value: &[u8]) -> Vec<u8> {
        [
            varint(u64::from(number) << 3 | 2),
            varint(value.len() as u64),
            value.to_vec(),
        ]
        .concat()
    }

    fn chart(drawable: &[u8]) -> Vec<u8> {
        field_bytes(1, drawable)
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    #[test]
    fn generated_chart_drawable_wire_matches_the_strict_projection() {
        use crate::{tsch, tsd};
        use prost::Message as _;

        for locked in [None, Some(false), Some(true)] {
            for constrained in [None, Some(false), Some(true)] {
                let drawable = tsd::DrawableArchive {
                    locked,
                    aspect_ratio_locked: constrained,
                    ..Default::default()
                };
                let source = tsch::ChartDrawableArchive {
                    super_: Some(drawable),
                }
                .encode_to_vec();
                let snapshot =
                    decode_chart_arrangement(&source, options(&source)).expect("arrangement");
                assert!(snapshot.has_drawable());
                assert_eq!(snapshot.locked(), locked);
                assert_eq!(snapshot.constrain_proportions(), constrained);
                assert_eq!(snapshot.is_locked(), locked.unwrap_or(false));
                assert_eq!(snapshot.is_constrained(), constrained.unwrap_or(false));
                assert_eq!(snapshot.raw(), source.as_slice());
            }
        }
    }

    #[test]
    fn absent_super_and_controls_follow_native_false_defaults() {
        let source = field_varint(4000, 9);
        let snapshot = decode_chart_arrangement(&source, options(&source)).expect("arrangement");
        assert!(!snapshot.has_drawable());
        assert_eq!(snapshot.locked(), None);
        assert_eq!(snapshot.constrain_proportions(), None);
        assert!(!snapshot.is_locked());
        assert!(!snapshot.is_constrained());

        let drawable = field_varint(4000, 10);
        let source = chart(&drawable);
        let snapshot = decode_chart_arrangement(&source, options(&source)).expect("empty drawable");
        assert!(snapshot.has_drawable());
        assert_eq!(snapshot.locked(), None);
        assert_eq!(snapshot.constrain_proportions(), None);
    }

    #[test]
    fn duplicate_root_and_nested_singular_fields_are_rejected() {
        let drawable = field_varint(5, 1);
        let duplicate_root = [chart(&drawable), chart(&drawable)].concat();
        let error = decode_chart_arrangement(&duplicate_root, options(&duplicate_root))
            .expect_err("duplicate root super");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSCH.ChartDrawableArchive.super")
        );

        let duplicate_locked = [field_varint(5, 1), field_varint(5, 0)].concat();
        let source = chart(&duplicate_locked);
        let error =
            decode_chart_arrangement(&source, options(&source)).expect_err("duplicate lock");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSD.DrawableArchive.locked")
        );
    }

    #[test]
    fn selected_wrong_wire_and_noncanonical_boolean_values_fail_closed() {
        let wrong_root_wire = field_varint(1, 1);
        assert!(decode_chart_arrangement(&wrong_root_wire, options(&wrong_root_wire)).is_err());

        let wrong_locked_wire = chart(&field_bytes(5, &[1]));
        assert!(decode_chart_arrangement(&wrong_locked_wire, options(&wrong_locked_wire)).is_err());

        for value in [2, 127, u64::MAX] {
            let source = chart(&field_varint(5, value));
            let error =
                decode_chart_arrangement(&source, options(&source)).expect_err("noncanonical bool");
            assert_eq!(
                error.noncanonical_reason(),
                Some("bool scalar is not zero or one")
            );
        }
    }

    #[test]
    fn unknown_fields_and_groups_are_checked_and_preserved() {
        let mut drawable = vec![0x53]; // unknown field 10 start-group
        drawable.extend(field_varint(4000, 77));
        drawable.push(0x54); // field 10 end-group
        drawable.extend(field_varint(5, 0));
        drawable.extend(field_varint(7, 1));
        let mut source = chart(&drawable);
        source.extend(field_varint(4001, 88));

        let snapshot = decode_chart_arrangement(&source, options(&source)).expect("unknowns");
        assert_eq!(snapshot.locked(), Some(false));
        assert_eq!(snapshot.constrain_proportions(), Some(true));

        let rewritten = rewrite_chart_arrangement(
            &source,
            ChartArrangementWrite::new(true, false),
            DecodeOptions::new(1024, 128, 16 * 1024, 3).with_max_output_bytes(1024),
        )
        .expect("rewrite");
        assert!(
            rewritten
                .windows(3)
                .any(|window| window == [0x53, 0x80, 0xfa])
        );
        assert!(
            rewritten
                .windows(3)
                .any(|window| window == [0x88, 0xfa, 0x01])
        );
        let changed = decode_chart_arrangement(&rewritten, options(&rewritten)).expect("changed");
        assert_eq!(changed.locked(), Some(true));
        assert_eq!(changed.constrain_proportions(), Some(false));

        let error =
            decode_chart_arrangement(&source, DecodeOptions::new(source.len(), 128, 16 * 1024, 1))
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
    fn rewrite_replaces_appends_clears_and_preserves_absent_false() {
        let drawable = [
            field_varint(4000, 4),
            field_varint(5, 0),
            field_varint(7, 1),
        ]
        .concat();
        let source = [chart(&drawable), field_varint(4001, 5)].concat();
        let rewritten = rewrite_chart_arrangement(
            &source,
            ChartArrangementWrite::new(true, false),
            DecodeOptions::new(1024, 128, 16 * 1024, 8).with_max_output_bytes(1024),
        )
        .expect("replace");
        let snapshot = decode_chart_arrangement(&rewritten, options(&rewritten)).expect("replace");
        assert_eq!(snapshot.locked(), Some(true));
        assert_eq!(snapshot.constrain_proportions(), Some(false));
        assert!(rewritten.ends_with(&field_varint(4001, 5)));

        let absent = field_varint(4000, 9);
        let effective_false = rewrite_chart_arrangement(
            &absent,
            ChartArrangementWrite::new(false, false),
            options(&absent),
        )
        .expect("semantic false no-op");
        assert_eq!(effective_false, absent);

        let explicit_false = rewrite_chart_arrangement(
            &absent,
            ChartArrangementWrite::explicit(Some(false), Some(false)),
            DecodeOptions::new(1024, 128, 16 * 1024, 8).with_max_output_bytes(1024),
        )
        .expect("explicit false");
        let explicit_snapshot =
            decode_chart_arrangement(&explicit_false, options(&explicit_false)).expect("explicit");
        assert_eq!(explicit_snapshot.locked(), Some(false));
        assert_eq!(explicit_snapshot.constrain_proportions(), Some(false));

        let cleared = rewrite_chart_arrangement(
            &source,
            ChartArrangementWrite::clear(),
            DecodeOptions::new(1024, 128, 16 * 1024, 8).with_max_output_bytes(1024),
        )
        .expect("clear");
        let cleared_snapshot =
            decode_chart_arrangement(&cleared, options(&cleared)).expect("clear");
        assert_eq!(cleared_snapshot.locked(), None);
        assert_eq!(cleared_snapshot.constrain_proportions(), None);
    }

    #[test]
    fn prepared_rewrite_reports_requirements_and_executes_at_exact_limits() {
        let source = chart(&[field_varint(5, 0), field_varint(4000, 2)].concat());
        let policy = DecodeOptions::new(1024, 128, 16 * 1024, 8)
            .with_max_output_bytes(1024)
            .with_max_retained_bytes(2048)
            .with_max_scratch_bytes(1024);
        let prepared = prepare_chart_arrangement_rewrite(
            &source,
            ChartArrangementWrite::new(true, true),
            policy,
        )
        .expect("prepare");
        let requirements = prepared.execution_requirements();
        assert!(requirements.output_bytes >= source.len());
        assert_eq!(requirements.allocations, 1);
        let output = prepared.execute(requirements.exact()).expect("execute");
        assert_eq!(
            decode_chart_arrangement(output.output(), options(output.output()))
                .expect("readback")
                .locked(),
            Some(true)
        );
        assert!(output.report().changed());
    }

    #[test]
    fn no_op_is_source_exact_and_rewrite_limits_fail_before_publication() {
        let source = chart(&[field_varint(5, 1), field_varint(7, 0)].concat());
        let (noop, report) = rewrite_chart_arrangement_with_report(
            &source,
            ChartArrangementWrite::new(true, false),
            options(&source),
        )
        .expect("no-op");
        assert_eq!(noop, source);
        assert!(!report.changed());

        let permissive = DecodeOptions::new(1024, 1024, 64 * 1024, 8)
            .with_max_output_bytes(1024)
            .with_max_retained_bytes(4096)
            .with_max_scratch_bytes(2048);
        let (_, report) = rewrite_chart_arrangement_with_report(
            &source,
            ChartArrangementWrite::new(false, true),
            permissive,
        )
        .expect("permissive");

        let fields_below = permissive.with_max_fields(report.fields() - 1);
        let error = rewrite_chart_arrangement(
            &source,
            ChartArrangementWrite::new(false, true),
            fields_below,
        )
        .expect_err("field cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));

        let work_below = permissive.with_max_work_bytes(report.work_bytes() - 1);
        let error =
            rewrite_chart_arrangement(&source, ChartArrangementWrite::new(false, true), work_below)
                .expect_err("work cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        assert_eq!(
            source,
            chart(&[field_varint(5, 1), field_varint(7, 0)].concat())
        );
    }

    #[test]
    fn output_and_allocation_limits_are_bounded() {
        let source = chart(&field_varint(5, 0));
        let error = rewrite_chart_arrangement(
            &source,
            ChartArrangementWrite::new(true, true),
            DecodeOptions::new(1024, 128, 16 * 1024, 8).with_max_output_bytes(source.len()),
        )
        .expect_err("output cap");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Output { .. })
        ));

        let error = super::reserve_output(usize::MAX).expect_err("allocation cap");
        assert_eq!(error.allocation_amount(), Some(usize::MAX));
    }
}
