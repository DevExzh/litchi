//! Strict borrowed projection for the Keynote chart-axis-title generated extension.
//!
//! The selected payload is only the generated `ChartAxisNonStyleArchive`
//! scalar pairs: fields 13/15 for the category axis and 14/16 for the value
//! axis. The outer
//! `TSCH.ChartAxisNonStyleArchive` envelope, chart graph, unrelated generated
//! fields, and unknown source bytes remain caller-owned. This module performs
//! only a borrowed projection and a wire-local rewrite; it never performs a
//! chart lookup or archive transaction.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict preflight intentionally precedes the low-level wire reader it consumes."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_chart_axis_title_generated::LitchiIwaProjection as projection;

const CHART_AXIS_CATEGORY_VISIBLE_FIELD: u32 = 13;
const CHART_AXIS_VALUE_VISIBLE_FIELD: u32 = 14;
const CHART_AXIS_CATEGORY_TEXT_FIELD: u32 = 15;
const CHART_AXIS_VALUE_TEXT_FIELD: u32 = 16;
const ROOT_DEPTH: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;

// The axis-pair preflight is intentionally kept wire-strict (the generated
// projection is only used as an independent semantic cross-check).
fn preflight_axis_titles<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<AxisTitleSnapshot<'source>, DecodeError> {
    preflight_axis_title(source, budget)
}

/// Finite limits for one generated chart-axis-title extension payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
    max_title_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Build an explicit bytes/fields/work/nesting policy.
    ///
    /// The output and title ceilings default to the source byte ceiling and
    /// can be replaced independently with the builder methods below. Decode
    /// output bytes count borrowed title UTF-8 bytes; rewrite output bytes
    /// count the candidate payload and preserve unknown source spans.
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
            max_title_bytes: max_message_bytes,
            max_allocations: usize::MAX,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: usize::MAX,
        }
    }

    /// Build conservative finite limits from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.checked_mul(4).unwrap_or(usize::MAX).max(1),
            bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
            8,
        )
        .with_max_allocations(1)
        .with_max_retained_bytes(bytes.saturating_mul(2))
        .with_max_scratch_bytes(bytes)
    }

    /// Replace the aggregate output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the selected title UTF-8 byte ceiling.
    #[must_use]
    pub const fn with_max_title_bytes(mut self, maximum: usize) -> Self {
        self.max_title_bytes = maximum;
        self
    }

    /// Compatibility spelling for callers that call the selected value text.
    #[must_use]
    pub const fn with_max_text_bytes(self, maximum: usize) -> Self {
        self.with_max_title_bytes(maximum)
    }

    /// Replace the logical allocation ceiling for prepared execution.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the peak retained-byte ceiling for prepared execution.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace the scratch-byte ceiling for prepared execution.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Construct all six finite ceilings explicitly.
    #[must_use]
    pub const fn with_output_and_title_limits(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_output_bytes: usize,
        max_title_bytes: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_output_bytes,
            max_title_bytes,
            max_allocations: usize::MAX,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: usize::MAX,
        }
    }

    /// Alias for [`Self::with_output_and_title_limits`] for callers that use
    /// the conventional `new_with_limits` naming.
    #[must_use]
    pub const fn new_with_limits(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_output_bytes: usize,
        max_title_bytes: usize,
    ) -> Self {
        Self::with_output_and_title_limits(
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_output_bytes,
            max_title_bytes,
        )
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Borrowed semantic facts from one generated chart-axis-title extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AxisTitleSnapshot<'source> {
    title_visible: Option<bool>,
    title: Option<&'source str>,
    value_title_visible: Option<bool>,
    value_title: Option<&'source str>,
    raw: &'source [u8],
}

impl<'source> AxisTitleSnapshot<'source> {
    /// Optional native field 13, preserving proto2 presence.
    #[must_use]
    pub const fn title_visible(self) -> Option<bool> {
        self.title_visible
    }

    /// Alias for [`Self::title_visible`].
    #[must_use]
    pub const fn show_title(self) -> Option<bool> {
        self.title_visible
    }

    /// Optional native field 15 borrowed directly from the source payload.
    #[must_use]
    pub const fn title(self) -> Option<&'source str> {
        self.title
    }

    /// Optional native field 14, preserving proto2 presence.
    #[must_use]
    pub const fn value_title_visible(self) -> Option<bool> {
        self.value_title_visible
    }

    /// Optional native field 16 borrowed directly from the source payload.
    #[must_use]
    pub const fn value_title(self) -> Option<&'source str> {
        self.value_title
    }

    /// Return the title only when field 13 explicitly enables it.
    ///
    /// A visible title with an absent field 15 follows the native empty-title
    /// default and returns `Some("")` without allocating.
    #[must_use]
    pub fn visible_title(self, kind: AxisTitleKind) -> Option<&'source str> {
        let (visible, title) = match kind {
            AxisTitleKind::Category => (self.title_visible, self.title),
            AxisTitleKind::Value => (self.value_title_visible, self.value_title),
        };
        (visible == Some(true)).then(|| title.unwrap_or_default())
    }

    /// Return the exact caller-owned bytes that were preflighted.
    ///
    /// The projection never rewrites, normalizes, or stores unknown fields;
    /// callers can retain this slice as their lossless preservation value.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisTitleKind {
    /// Category (x) axis, represented by fields 13 and 15.
    Category,
    /// Value (y) axis, represented by fields 14 and 16.
    Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AxisTitleUpdate<'title> {
    Preserve,
    Replace {
        visible: Option<bool>,
        title: Option<&'title str>,
    },
}

/// Requested presence-preserving values for a generated chart-axis-title rewrite.
///
/// This is intentionally a wire-level update only. It does not locate a
/// chart, resolve a `chart_non_style` reference, or validate a chart graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AxisTitleWrite<'title> {
    category: AxisTitleUpdate<'title>,
    value: AxisTitleUpdate<'title>,
}

impl<'title> AxisTitleWrite<'title> {
    /// Build a requested proto2 presence/value pair.
    #[must_use]
    pub const fn new(title_visible: Option<bool>, title: Option<&'title str>) -> Self {
        Self {
            category: AxisTitleUpdate::Replace {
                visible: title_visible,
                title,
            },
            value: AxisTitleUpdate::Preserve,
        }
    }

    /// Start a source-preserving rewrite with neither axis selected.
    #[must_use]
    pub const fn preserve() -> Self {
        Self {
            category: AxisTitleUpdate::Preserve,
            value: AxisTitleUpdate::Preserve,
        }
    }

    /// Set or clear one axis title while preserving the opposite axis.
    ///
    /// `Some(text)` writes an explicitly visible title and `None` writes an
    /// explicitly hidden title with its text field removed.
    #[must_use]
    pub const fn with_title(mut self, kind: AxisTitleKind, title: Option<&'title str>) -> Self {
        let update = AxisTitleUpdate::Replace {
            visible: Some(title.is_some()),
            title,
        };
        match kind {
            AxisTitleKind::Category => self.category = update,
            AxisTitleKind::Value => self.value = update,
        }
        self
    }

    /// Requested field-13 presence/value.
    #[must_use]
    pub const fn title_visible(self) -> Option<bool> {
        match self.category {
            AxisTitleUpdate::Preserve => None,
            AxisTitleUpdate::Replace { visible, .. } => visible,
        }
    }

    /// Requested field-15 UTF-8 title.
    #[must_use]
    pub const fn title(self) -> Option<&'title str> {
        match self.category {
            AxisTitleUpdate::Preserve => None,
            AxisTitleUpdate::Replace { title, .. } => title,
        }
    }
}

/// Exact finite consumption of one generated-extension rewrite.
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
    /// Input payload bytes inspected before the rewrite.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Candidate payload bytes published by the wire-level rewrite.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Strict field visits from source preflight, rewrite sizing/emission, and
    /// (when changed) candidate readback.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa work charged for the complete decode or
    /// rewrite transaction.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum source nesting observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Logical allocations required by execution.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Peak bytes retained by source plus candidate output.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Reserved candidate staging bytes.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether at least one selected field changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Exact resource consumption from one successful decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    output_bytes: usize,
    title_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    /// Input bytes inspected by the strict preflight.
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

    /// Maximum protobuf nesting depth observed during strict traversal.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Aggregate borrowed UTF-8 bytes charged for both axis titles.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate UTF-8 bytes in fields 15 and 16.
    #[must_use]
    pub const fn title_bytes(self) -> usize {
        self.title_bytes
    }

    /// Logical allocations performed by the borrowed decode.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained for the borrowed snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Scratch bytes retained by the borrowed decode.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact requirements computed by rewrite preparation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub allocations: usize,
    pub retained_bytes: usize,
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Convert exact requirements into exact execution ceilings.
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

/// Caller-provided ceilings checked before prepared output allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub allocations: usize,
    pub retained_bytes: usize,
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        requirements.exact()
    }

    #[must_use]
    pub const fn with_output_bytes(mut self, value: usize) -> Self {
        self.output_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_fields(mut self, value: usize) -> Self {
        self.fields = value;
        self
    }

    #[must_use]
    pub const fn with_work_bytes(mut self, value: usize) -> Self {
        self.work_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_max_depth(mut self, value: u32) -> Self {
        self.max_depth = value;
        self
    }

    #[must_use]
    pub const fn with_allocations(mut self, value: usize) -> Self {
        self.allocations = value;
        self
    }

    #[must_use]
    pub const fn with_retained_bytes(mut self, value: usize) -> Self {
        self.retained_bytes = value;
        self
    }

    #[must_use]
    pub const fn with_scratch_bytes(mut self, value: usize) -> Self {
        self.scratch_bytes = value;
        self
    }
}

/// Candidate bytes and exact prepared execution report.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    output: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    #[must_use]
    pub fn output(&self) -> &[u8] {
        &self.output
    }

    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.output
    }

    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.output
    }

    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source or configured Buffa message bytes exceeded the ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Encoded field visits exceeded the ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict-plus-Buffa work exceeded the ceiling.
    Work { observed: usize, maximum: usize },
    /// Borrowed semantic output exceeded the ceiling.
    Output { observed: usize, maximum: usize },
    /// Aggregate field-15/16 title bytes exceeded the ceiling.
    Title { observed: usize, maximum: usize },
    /// Configured or traversed protobuf nesting exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Logical prepared-execution allocations exceeded the ceiling.
    Allocations { observed: usize, maximum: usize },
    /// Peak retained bytes exceeded the ceiling.
    Retained { observed: usize, maximum: usize },
    /// Scratch staging bytes exceeded the ceiling.
    Scratch { observed: usize, maximum: usize },
}

/// The byte/nesting subset of [`DecodeLimit`] used by older focused codecs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// Source/configured message bytes exceeded the ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Configured/traversed nesting exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
}

/// Failure from strict chart-axis-title preflight or the private Buffa view.
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
    InvalidUtf8(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    OutputLimit { observed: usize, maximum: usize },
    TitleLimit { observed: usize, maximum: usize },
    Allocation { amount: usize },
    Projection,
}

impl DecodeError {
    /// Return the duplicated singular field, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return the stable canonical-wire failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    /// Return the rejected UTF-8 field, when applicable.
    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::InvalidUtf8(field) => Some(field),
            _ => None,
        }
    }

    /// Return the exact field-limit observation, when applicable.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::FieldLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return the exact work-limit observation, when applicable.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::WorkLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return the exact output-limit observation, when applicable.
    #[must_use]
    pub const fn output_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::OutputLimit { observed, maximum }
            | DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return the exact title-limit observation, when applicable.
    #[must_use]
    pub const fn title_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::TitleLimit { observed, maximum }
            | DecodeErrorKind::Resource(DecodeLimit::Title { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return the requested output allocation, when it could not be reserved.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            _ => None,
        }
    }

    /// Return the typed finite resource failure, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            DecodeErrorKind::FieldLimit { observed, maximum } => {
                Some(DecodeLimit::Fields { observed, maximum })
            },
            DecodeErrorKind::WorkLimit { observed, maximum } => {
                Some(DecodeLimit::Work { observed, maximum })
            },
            DecodeErrorKind::OutputLimit { observed, maximum } => {
                Some(DecodeLimit::Output { observed, maximum })
            },
            DecodeErrorKind::TitleLimit { observed, maximum } => {
                Some(DecodeLimit::Title { observed, maximum })
            },
            _ => None,
        }
    }

    /// Return the byte/nesting resource failure used by the caption codec.
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

    const fn duplicate_singular(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    const fn field_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::FieldLimit { observed, maximum },
        }
    }

    const fn work_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::WorkLimit { observed, maximum },
        }
    }

    const fn output_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::OutputLimit { observed, maximum },
        }
    }

    const fn title_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::TitleLimit { observed, maximum },
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
                "Keynote chart-axis-title projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote chart-axis-title projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum })
            | DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-axis-title projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum })
            | DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-axis-title projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum })
            | DecodeErrorKind::OutputLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-axis-title projection produced {observed} output bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Title { observed, maximum })
            | DecodeErrorKind::TitleLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-axis-title projection title has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "Keynote chart-axis-title rewrite requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "Keynote chart-axis-title rewrite retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "Keynote chart-axis-title rewrite requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate Keynote chart-axis-title output for {amount} bytes"
            ),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote chart-axis-title strict preflight disagrees with the Buffa projection",
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge => Self {
                kind: DecodeErrorKind::Resource(DecodeLimit::Bytes {
                    observed: 0,
                    maximum: 0,
                }),
            },
            buffa::DecodeError::RecursionLimitExceeded => Self {
                kind: DecodeErrorKind::Resource(DecodeLimit::Nesting {
                    observed: 0,
                    maximum: MAX_RECURSION_LIMIT,
                }),
            },
            other => Self {
                kind: DecodeErrorKind::Wire(other),
            },
        }
    }
}

/// Decode the generated extension and return its borrowed scalar snapshot.
pub fn decode_axis_title<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<AxisTitleSnapshot<'source>, DecodeError> {
    Ok(decode_axis_title_with_report(source, options)?.0)
}

/// Decode both category- and value-axis title controls.
pub fn decode_axis_titles<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<AxisTitleSnapshot<'source>, DecodeError> {
    decode_axis_title(source, options)
}

pub fn decode_axis_titles_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(AxisTitleSnapshot<'source>, DecodeReport), DecodeError> {
    decode_axis_title_with_report(source, options)
}

/// Prepared source-preserving axis-title rewrite.
///
/// Preparation validates and measures both source and candidate semantics but
/// does not allocate candidate output. Consuming [`Self::execute`] makes a
/// prepared plan one-shot.
#[derive(Debug, Clone, Copy)]
pub struct PreparedAxisTitleRewrite<'source, 'title> {
    source: &'source [u8],
    write: AxisTitleWrite<'title>,
    current: AxisTitleSnapshot<'source>,
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
}

impl PreparedAxisTitleRewrite<'_, '_> {
    /// Return source-only accounting performed during preparation.
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }

    /// Return exact aggregate preparation/execution requirements.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute once after replaying caller-provided ceilings.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_execution_limits(self.requirements, limits)?;
        let output = emit_rewrite(
            self.source,
            self.write,
            self.requirements.output_bytes,
            self.source_report,
            self.options,
        )?;
        let readback_options = DecodeOptions {
            max_message_bytes: output.len(),
            max_fields: self.candidate_fields,
            max_work_bytes: self.candidate_work,
            recursion_limit: self.options.recursion_limit,
            max_output_bytes: self.options.max_output_bytes,
            max_title_bytes: self.options.max_title_bytes,
            max_allocations: self.options.max_allocations,
            max_retained_bytes: self.options.max_retained_bytes,
            max_scratch_bytes: self.options.max_scratch_bytes,
        };
        let (readback, candidate_report) =
            decode_axis_titles_with_report(&output, readback_options)?;
        if candidate_report.fields() != self.candidate_fields
            || candidate_report.work_bytes() != self.candidate_work
            || candidate_report.max_depth() != self.candidate_depth
            || !candidate_matches(readback, self.current, self.write)
        {
            return Err(DecodeError::projection());
        }
        let report = RewriteReport {
            input_bytes: self.source.len(),
            output_bytes: output.len(),
            fields: self.requirements.fields,
            work_bytes: self.requirements.work_bytes,
            max_depth: self.requirements.max_depth,
            allocations: self.requirements.allocations,
            retained_bytes: self.requirements.retained_bytes,
            scratch_bytes: self.requirements.scratch_bytes,
            changed: output.as_slice() != self.source,
        };
        Ok(RewriteOutput { output, report })
    }
}

/// Prepare a source-preserving dual-axis rewrite without allocating output.
pub fn prepare_axis_title_rewrite<'source, 'title>(
    source: &'source [u8],
    write: AxisTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<PreparedAxisTitleRewrite<'source, 'title>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let current = decode_axis_title_with_budget(source, options, &mut budget)?;
    let source_report = budget.report();
    let write = normalize_write(current, write);
    validate_candidate_titles(current, write, options)?;
    let measure = measure_rewrite_output(source, write, options, &mut budget)?;
    let candidate_work = measure
        .bytes
        .checked_mul(2)
        .ok_or_else(|| DecodeError::work_limit(usize::MAX, options.max_work_bytes))?;
    let fields = source_report
        .fields
        .checked_mul(3)
        .and_then(|value| value.checked_add(measure.fields))
        .ok_or_else(|| DecodeError::field_limit(usize::MAX, options.max_fields))?;
    let work_bytes = source
        .len()
        .checked_mul(6)
        .and_then(|value| value.checked_add(candidate_work))
        .ok_or_else(|| DecodeError::work_limit(usize::MAX, options.max_work_bytes))?;
    if fields > options.max_fields {
        return Err(DecodeError::field_limit(fields, options.max_fields));
    }
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::work_limit(work_bytes, options.max_work_bytes));
    }
    let allocations = usize::from(measure.bytes != 0);
    let retained_bytes = source.len().checked_add(measure.bytes).ok_or(DecodeError {
        kind: DecodeErrorKind::Resource(DecodeLimit::Retained {
            observed: usize::MAX,
            maximum: options.max_retained_bytes,
        }),
    })?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: measure.bytes,
        fields,
        work_bytes,
        max_depth: source_report.max_depth,
        allocations,
        retained_bytes,
        scratch_bytes: measure.bytes,
    };
    check_option_execution_limits(requirements, options)?;
    Ok(PreparedAxisTitleRewrite {
        source,
        write,
        current,
        options,
        source_report,
        requirements,
        candidate_fields: measure.fields,
        candidate_work,
        candidate_depth: source_report.max_depth,
    })
}

/// Decode the generated extension and return exact aggregate consumption.
pub fn decode_axis_title_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(AxisTitleSnapshot<'source>, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = decode_axis_title_with_budget(source, options, &mut budget)?;
    Ok((snapshot, budget.report()))
}

/// Decode one chart-axis-title payload while charging an existing aggregate
/// budget. Rewrites use this entry point for both their source and candidate
/// readback passes so the field/work ceilings cover the complete transaction.
fn decode_axis_title_with_budget<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<AxisTitleSnapshot<'source>, DecodeError> {
    // Work and field counters are aggregate across a rewrite transaction,
    // while the selected title/output ceilings apply to each decode pass.
    // Reset the per-pass semantic counters before source and candidate reads
    // so a large replacement is not rejected merely because both snapshots
    // are charged through one shared budget.
    budget.output_bytes = 0;
    budget.title_bytes = 0;
    budget.message(source.len(), ROOT_DEPTH)?;
    let strict = preflight_axis_titles(source, budget)?;
    let view: projection::ChartAxisTitleArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = AxisTitleSnapshot {
        title_visible: view.tschchartaxiscategoryshowtitle,
        title: view.tschchartaxiscategorytitle,
        value_title_visible: view.tschchartaxisvalueshowtitle,
        value_title: view.tschchartaxisvaluetitle,
        raw: source,
    };
    if projected.title_visible != strict.title_visible
        || projected.title != strict.title
        || projected.value_title_visible != strict.value_title_visible
        || projected.value_title != strict.value_title
    {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode only the optional field-15 title text.
pub fn decode_axis_title_text(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<&str>, DecodeError> {
    Ok(decode_axis_title(source, options)?.title())
}

/// Decode the title text only when native field 13 is explicitly true.
pub fn decode_visible_axis_title(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<&str>, DecodeError> {
    Ok(decode_axis_title(source, options)?.visible_title(AxisTitleKind::Category))
}

/// Rewrite selected fields 13–16 of one generated-extension payload.
///
/// Existing unknown field spans and their order are copied byte-for-byte.
/// Existing selected fields are replaced at their original positions. A
/// requested selected field absent from `source` is appended after all source
/// spans in field-number order (13, 14, 15, then 16).
///
/// This is deliberately a source-local contract. Once a selected field has
/// been removed, its former position is no longer represented by the wire
/// payload, so a later call cannot infer where to put it back. Callers that
/// need an exact inverse across a remove/add sequence must retain the source
/// layout (or the original bytes) outside this API and replay the inverse in a
/// transaction that has that metadata. The codec never guesses a missing
/// field's historical position.
///
/// If the requested pair already matches, the returned bytes are an exact
/// source copy and the report marks the operation as unchanged.
pub fn rewrite_axis_title<'source, 'title>(
    source: &'source [u8],
    write: AxisTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_axis_title_with_report(source, write, options)?.0)
}

/// Rewrite the generated extension and return exact source/output accounting.
pub fn rewrite_axis_title_with_report<'source, 'title>(
    source: &'source [u8],
    write: AxisTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_axis_title_rewrite(source, write, options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report();
    Ok((output.into_output(), report))
}

/// Rewrite using the explicit generated-extension spelling.
pub fn rewrite_axis_title_extension<'source, 'title>(
    source: &'source [u8],
    write: AxisTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_axis_title(source, write, options)
}

/// Compatibility spelling emphasizing that `source` is the extension payload.
pub fn decode_axis_title_extension<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<AxisTitleSnapshot<'source>, DecodeError> {
    decode_axis_title(source, options)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum = hard_message_maximum(options.max_message_bytes)?;
    if options.max_message_bytes > hard_maximum {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_maximum,
            }),
        });
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }),
        });
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }),
        });
    }
    Ok(())
}

fn hard_message_maximum(observed: usize) -> Result<usize, DecodeError> {
    usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError {
        kind: DecodeErrorKind::Resource(DecodeLimit::Bytes {
            observed,
            maximum: usize::MAX,
        }),
    })
}

#[derive(Debug)]
struct Budget {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    output_bytes: usize,
    title_bytes: usize,
    options: DecodeOptions,
}

impl Budget {
    const fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            source_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            output_bytes: 0,
            title_bytes: 0,
            options,
        }
    }

    fn field(&mut self) -> Result<(), DecodeError> {
        let observed = self
            .fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::field_limit(usize::MAX, self.options.max_fields))?;
        if observed > self.options.max_fields {
            return Err(DecodeError::field_limit(observed, self.options.max_fields));
        }
        self.fields = observed;
        Ok(())
    }

    fn message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Nesting {
                    observed: depth,
                    maximum: self.options.recursion_limit,
                }),
            });
        }
        self.max_depth = self.max_depth.max(depth);
        let cost = bytes
            .checked_mul(2)
            .ok_or_else(|| DecodeError::work_limit(usize::MAX, self.options.max_work_bytes))?;
        let observed = self
            .work_bytes
            .checked_add(cost)
            .ok_or_else(|| DecodeError::work_limit(usize::MAX, self.options.max_work_bytes))?;
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::work_limit(
                observed,
                self.options.max_work_bytes,
            ));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Nesting {
                    observed: depth,
                    maximum: self.options.recursion_limit,
                }),
            });
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    fn title(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let title_bytes = self
            .title_bytes
            .checked_add(bytes)
            .ok_or_else(|| DecodeError::title_limit(usize::MAX, self.options.max_title_bytes))?;
        if title_bytes > self.options.max_title_bytes {
            return Err(DecodeError::title_limit(
                title_bytes,
                self.options.max_title_bytes,
            ));
        }
        let output = self
            .output_bytes
            .checked_add(bytes)
            .ok_or_else(|| DecodeError::output_limit(usize::MAX, self.options.max_output_bytes))?;
        if output > self.options.max_output_bytes {
            return Err(DecodeError::output_limit(
                output,
                self.options.max_output_bytes,
            ));
        }
        self.title_bytes = title_bytes;
        self.output_bytes = output;
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            source_bytes: self.source_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            output_bytes: self.output_bytes,
            title_bytes: self.title_bytes,
            allocations: 0,
            retained_bytes: self.source_bytes,
            scratch_bytes: 0,
        }
    }
}

fn preflight_axis_title<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<AxisTitleSnapshot<'source>, DecodeError> {
    let mut visible = None;
    let mut title = None;
    let mut value_visible = None;
    let mut value_title = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, budget, ROOT_DEPTH)? {
        match field.number {
            CHART_AXIS_CATEGORY_VISIBLE_FIELD => {
                if visible.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategoryshowtitle",
                    ));
                }
                visible = Some(require_canonical_bool(field.varint()?)?);
            },
            CHART_AXIS_CATEGORY_TEXT_FIELD => {
                if title.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategorytitle",
                    ));
                }
                let bytes = field.length_delimited()?;
                budget.title(bytes.len())?;
                title = Some(str::from_utf8(bytes).map_err(|_error| {
                    DecodeError::invalid_utf8(
                        "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategorytitle",
                    )
                })?);
            },
            CHART_AXIS_VALUE_VISIBLE_FIELD => {
                if value_visible.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvalueshowtitle",
                    ));
                }
                value_visible = Some(require_canonical_bool(field.varint()?)?);
            },
            CHART_AXIS_VALUE_TEXT_FIELD => {
                if value_title.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvaluetitle",
                    ));
                }
                let bytes = field.length_delimited()?;
                budget.title(bytes.len())?;
                value_title = Some(str::from_utf8(bytes).map_err(|_error| {
                    DecodeError::invalid_utf8(
                        "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvaluetitle",
                    )
                })?);
            },
            _ => {},
        }
    }
    Ok(AxisTitleSnapshot {
        title_visible: visible,
        title,
        value_title_visible: value_visible,
        value_title,
        raw: source,
    })
}

#[derive(Clone, Copy, Debug)]
struct FieldSpan {
    number: u32,
    start: usize,
    end: usize,
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

#[derive(Debug, Clone, Copy)]
struct RewriteMeasure {
    bytes: usize,
    fields: usize,
}

fn measure_rewrite_output(
    source: &[u8],
    write: AxisTitleWrite<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RewriteMeasure, DecodeError> {
    let mut length = 0usize;
    // `budget.fields` is an aggregate counter.  Keep the candidate count
    // separate: the source scan below charges every field (including group
    // contents) to the aggregate budget, while a removed selected field must
    // be subtracted only from the candidate pass.  Starting this counter at
    // the aggregate value used to make every rewrite advertise too many
    // candidate fields and caused the exact readback check to reject valid
    // output.
    let fields_before_scan = budget.fields;
    let mut removed_fields = 0usize;
    let mut saw = [false; 4];
    visit_field_spans(source, budget, |span| {
        let replacement = match selected_slot(span.number) {
            Some((slot, update, is_text)) => {
                saw[slot] = true;
                match update_from_write(write, update) {
                    AxisTitleUpdate::Preserve => span.end - span.start,
                    AxisTitleUpdate::Replace { visible, title } => {
                        let replacement = if is_text {
                            title.map_or(Ok(0), |text| {
                                text_field_len(span.number, text).ok_or_else(|| {
                                    DecodeError::output_limit(usize::MAX, options.max_output_bytes)
                                })
                            })?
                        } else {
                            visible
                                .map_or(0, |value| varint_field_len(span.number, u64::from(value)))
                        };
                        if replacement == 0 {
                            removed_fields = removed_fields
                                .checked_add(1)
                                .ok_or_else(DecodeError::projection)?;
                        }
                        replacement
                    },
                }
            },
            None => span.end - span.start,
        };
        length = length
            .checked_add(replacement)
            .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
        Ok(())
    })?;

    let scanned_fields = budget
        .fields
        .checked_sub(fields_before_scan)
        .ok_or_else(DecodeError::projection)?;
    let mut fields = scanned_fields
        .checked_sub(removed_fields)
        .ok_or_else(DecodeError::projection)?;
    for (slot, number, update, is_text) in selected_fields(write) {
        if saw[slot] {
            continue;
        }
        let appended = match update {
            AxisTitleUpdate::Preserve => 0,
            AxisTitleUpdate::Replace { title, .. } if is_text => title.map_or(Ok(0), |text| {
                text_field_len(number, text)
                    .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))
            })?,
            AxisTitleUpdate::Replace { visible, .. } => {
                visible.map_or(0, |value| varint_field_len(number, u64::from(value)))
            },
        };
        if appended != 0 {
            fields = fields.checked_add(1).ok_or_else(DecodeError::projection)?;
            length = length
                .checked_add(appended)
                .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
        }
    }
    if length > options.max_output_bytes {
        return Err(DecodeError::output_limit(length, options.max_output_bytes));
    }
    let hard_maximum = hard_message_maximum(length)?;
    if length > hard_maximum {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(DecodeLimit::Bytes {
                observed: length,
                maximum: hard_maximum,
            }),
        });
    }
    Ok(RewriteMeasure {
        bytes: length,
        fields,
    })
}

fn selected_slot(number: u32) -> Option<(usize, AxisTitleKind, bool)> {
    match number {
        CHART_AXIS_CATEGORY_VISIBLE_FIELD => Some((0, AxisTitleKind::Category, false)),
        CHART_AXIS_VALUE_VISIBLE_FIELD => Some((1, AxisTitleKind::Value, false)),
        CHART_AXIS_CATEGORY_TEXT_FIELD => Some((2, AxisTitleKind::Category, true)),
        CHART_AXIS_VALUE_TEXT_FIELD => Some((3, AxisTitleKind::Value, true)),
        _ => None,
    }
}

fn update_from_write(write: AxisTitleWrite<'_>, kind: AxisTitleKind) -> AxisTitleUpdate<'_> {
    match kind {
        AxisTitleKind::Category => write.category,
        AxisTitleKind::Value => write.value,
    }
}

fn selected_fields(write: AxisTitleWrite<'_>) -> [(usize, u32, AxisTitleUpdate<'_>, bool); 4] {
    [
        (0, CHART_AXIS_CATEGORY_VISIBLE_FIELD, write.category, false),
        (1, CHART_AXIS_VALUE_VISIBLE_FIELD, write.value, false),
        (2, CHART_AXIS_CATEGORY_TEXT_FIELD, write.category, true),
        (3, CHART_AXIS_VALUE_TEXT_FIELD, write.value, true),
    ]
}

fn normalize_write<'title>(
    current: AxisTitleSnapshot<'_>,
    write: AxisTitleWrite<'title>,
) -> AxisTitleWrite<'title> {
    AxisTitleWrite {
        category: normalize_update(write.category, current.title_visible, current.title),
        value: normalize_update(
            write.value,
            current.value_title_visible,
            current.value_title,
        ),
    }
}

fn normalize_update<'title>(
    update: AxisTitleUpdate<'title>,
    current_visible: Option<bool>,
    current_title: Option<&str>,
) -> AxisTitleUpdate<'title> {
    match update {
        AxisTitleUpdate::Replace {
            visible: Some(false),
            title: None,
        } if current_visible != Some(true) => AxisTitleUpdate::Preserve,
        AxisTitleUpdate::Replace { visible, title }
            if visible == current_visible && title == current_title =>
        {
            AxisTitleUpdate::Preserve
        },
        other => other,
    }
}

fn validate_candidate_titles(
    current: AxisTitleSnapshot<'_>,
    write: AxisTitleWrite<'_>,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let category = match write.category {
        AxisTitleUpdate::Preserve => current.title,
        AxisTitleUpdate::Replace { title, .. } => title,
    };
    let value = match write.value {
        AxisTitleUpdate::Preserve => current.value_title,
        AxisTitleUpdate::Replace { title, .. } => title,
    };
    let observed = category
        .map_or(0, str::len)
        .checked_add(value.map_or(0, str::len))
        .ok_or_else(|| DecodeError::title_limit(usize::MAX, options.max_title_bytes))?;
    if observed > options.max_title_bytes {
        return Err(DecodeError::title_limit(observed, options.max_title_bytes));
    }
    Ok(())
}

fn candidate_matches(
    candidate: AxisTitleSnapshot<'_>,
    current: AxisTitleSnapshot<'_>,
    write: AxisTitleWrite<'_>,
) -> bool {
    let category = expected_axis(write.category, current.title_visible, current.title);
    let value = expected_axis(
        write.value,
        current.value_title_visible,
        current.value_title,
    );
    candidate.title_visible == category.0
        && candidate.title == category.1
        && candidate.value_title_visible == value.0
        && candidate.value_title == value.1
}

fn expected_axis<'value>(
    update: AxisTitleUpdate<'value>,
    visible: Option<bool>,
    title: Option<&'value str>,
) -> (Option<bool>, Option<&'value str>) {
    match update {
        AxisTitleUpdate::Preserve => (visible, title),
        AxisTitleUpdate::Replace { visible, title } => (visible, title),
    }
}

fn emit_rewrite(
    source: &[u8],
    write: AxisTitleWrite<'_>,
    output_bytes: usize,
    source_report: DecodeReport,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(output_bytes)?;
    let scan_options = DecodeOptions {
        max_message_bytes: source.len(),
        max_fields: source_report.fields,
        max_work_bytes: source.len().saturating_mul(2),
        ..options
    };
    let mut budget = Budget::new(source, scan_options);
    let mut saw = [false; 4];
    visit_field_spans(source, &mut budget, |span| {
        let raw = &source[span.start..span.end];
        let Some((slot, kind, is_text)) = selected_slot(span.number) else {
            output.extend_from_slice(raw);
            return Ok(());
        };
        saw[slot] = true;
        match update_from_write(write, kind) {
            AxisTitleUpdate::Preserve => output.extend_from_slice(raw),
            AxisTitleUpdate::Replace { title, .. } if is_text => {
                if let Some(value) = title {
                    append_text_field(&mut output, span.number, value);
                }
            },
            AxisTitleUpdate::Replace { visible, .. } => {
                if let Some(value) = visible {
                    append_varint_field(&mut output, span.number, u64::from(value));
                }
            },
        }
        Ok(())
    })?;
    for (slot, number, update, is_text) in selected_fields(write) {
        if saw[slot] {
            continue;
        }
        match update {
            AxisTitleUpdate::Preserve => {},
            AxisTitleUpdate::Replace { title, .. } if is_text => {
                if let Some(value) = title {
                    append_text_field(&mut output, number, value);
                }
            },
            AxisTitleUpdate::Replace { visible, .. } => {
                if let Some(value) = visible {
                    append_varint_field(&mut output, number, u64::from(value));
                }
            },
        }
    }
    if output.len() != output_bytes {
        return Err(DecodeError::projection());
    }
    Ok(output)
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
    for (observed, maximum, limit) in [
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
    ] {
        if observed > maximum {
            return Err(DecodeError {
                kind: DecodeErrorKind::Resource(limit),
            });
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(DecodeLimit::Nesting {
                observed: requirements.max_depth,
                maximum: limits.max_depth,
            }),
        });
    }
    Ok(())
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn text_field_len(number: u32, value: &str) -> Option<usize> {
    varint_len((u64::from(number) << 3) | 2)
        .checked_add(varint_len(value.len() as u64))?
        .checked_add(value.len())
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

fn append_text_field(output: &mut Vec<u8>, number: u32, value: &str) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, value.len() as u64);
    output.extend_from_slice(value.as_bytes());
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
            StrictValue::LengthDelimited(take_exact(source, length)?)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_depth = depth.checked_add(1).ok_or(DecodeError {
                kind: DecodeErrorKind::Resource(DecodeLimit::Nesting {
                    observed: u32::MAX,
                    maximum: budget.options.recursion_limit,
                }),
            })?;
            // A group contributes one level even when it is empty, so observe
            // its child depth before scanning the first child/end tag.
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
    clippy::shadow_unrelated,
    reason = "Focused negative tests use explicit panic messages and reuse local wire helpers."
)]
mod tests {
    use super::{
        AxisTitleKind, AxisTitleSnapshot, AxisTitleWrite, DecodeLimit, DecodeOptions,
        WireResourceLimit, decode_axis_title, decode_axis_title_text, decode_visible_axis_title,
    };

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

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

    fn field_text(number: u32, value: &[u8]) -> Vec<u8> {
        [
            varint((u64::from(number) << 3) | 2),
            varint(value.len() as u64),
            value.to_vec(),
        ]
        .concat()
    }

    #[test]
    fn selected_fields_project_without_allocating_or_rewriting_source() {
        let mut source = field_varint(13, 1);
        source.extend(field_text(15, "Revenue 📈".as_bytes()));
        let unknown = field_varint(4000, 42);
        source.extend_from_slice(&unknown);
        let before = source.clone();
        let snapshot = decode_axis_title(&source, options(&source)).expect("title");
        assert_eq!(snapshot.title_visible(), Some(true));
        assert_eq!(snapshot.show_title(), Some(true));
        assert_eq!(snapshot.title(), Some("Revenue 📈"));
        assert_eq!(
            snapshot.visible_title(AxisTitleKind::Category),
            Some("Revenue 📈")
        );
        assert_eq!(snapshot.raw(), before.as_slice());
        assert_eq!(
            decode_axis_title_text(&source, options(&source)),
            Ok(Some("Revenue 📈"))
        );
        assert_eq!(
            decode_visible_axis_title(&source, options(&source)),
            Ok(Some("Revenue 📈"))
        );
        assert_eq!(source, before);
    }

    #[test]
    fn absent_fields_are_exact_no_op_and_presence_is_retained() {
        let source = Vec::new();
        let snapshot = decode_axis_title(&source, options(&source)).expect("empty");
        assert_eq!(
            snapshot,
            AxisTitleSnapshot {
                title_visible: None,
                title: None,
                value_title_visible: None,
                value_title: None,
                raw: &[]
            }
        );
        assert_eq!(snapshot.visible_title(AxisTitleKind::Category), None);

        let hidden = field_varint(13, 0);
        let snapshot = decode_axis_title(&hidden, options(&hidden)).expect("hidden");
        assert_eq!(snapshot.title_visible(), Some(false));
        assert_eq!(snapshot.title(), None);
        assert_eq!(snapshot.visible_title(AxisTitleKind::Category), None);

        let visible_without_text = field_varint(13, 1);
        let snapshot = decode_axis_title(&visible_without_text, options(&visible_without_text))
            .expect("default empty title");
        assert_eq!(snapshot.title(), None);
        assert_eq!(snapshot.visible_title(AxisTitleKind::Category), Some(""));
    }

    #[test]
    fn duplicate_wrong_wire_and_truncated_selected_fields_are_rejected() {
        let mut duplicate_visible = field_varint(13, 1);
        duplicate_visible.extend(field_varint(13, 0));
        assert_eq!(
            decode_axis_title(&duplicate_visible, options(&duplicate_visible))
                .expect_err("duplicate visible")
                .duplicate_singular_field(),
            Some("TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategoryshowtitle")
        );

        let mut duplicate_title = field_text(15, b"a");
        duplicate_title.extend(field_text(15, b"b"));
        assert_eq!(
            decode_axis_title(&duplicate_title, options(&duplicate_title))
                .expect_err("duplicate title")
                .duplicate_singular_field(),
            Some("TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategorytitle")
        );

        for source in [
            vec![0x6a, 0x01],       // field 13 with length-delimited wire type
            vec![0x7a, 0x02, b'a'], // field 15 with truncated payload
            vec![0x68],             // field 13 with truncated varint
        ] {
            assert!(
                decode_axis_title(&source, options(&source)).is_err(),
                "{source:?}"
            );
        }
    }

    #[test]
    fn bool_utf8_and_canonical_wire_rules_are_enforced() {
        let bad_bool = field_varint(13, 2);
        assert_eq!(
            decode_axis_title(&bad_bool, options(&bad_bool))
                .expect_err("noncanonical bool")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );

        let bad_utf8 = field_text(15, &[0xff]);
        assert_eq!(
            decode_axis_title(&bad_utf8, options(&bad_utf8))
                .expect_err("invalid UTF-8")
                .invalid_utf8_field(),
            Some("TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategorytitle")
        );

        let overlong_key = vec![0xe8, 0x00, 0x01];
        assert_eq!(
            decode_axis_title(&overlong_key, options(&overlong_key))
                .expect_err("overlong key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let overlong_length = vec![0x7a, 0x80, 0x00];
        assert_eq!(
            decode_axis_title(&overlong_length, options(&overlong_length))
                .expect_err("overlong length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );
    }

    #[test]
    fn finite_field_work_output_title_and_nesting_limits_are_observable() {
        let source = [field_varint(13, 1), field_text(15, b"abcd")].concat();
        assert_eq!(
            decode_axis_title(
                &source,
                DecodeOptions::new(source.len(), 1, source.len() * 8, 8),
            )
            .expect_err("field cap")
            .field_limit_values(),
            Some((2, 1))
        );
        assert_eq!(
            decode_axis_title(
                &source,
                DecodeOptions::new(source.len(), 8, source.len(), 8),
            )
            .expect_err("work cap")
            .work_limit_values(),
            Some((source.len() * 2, source.len()))
        );
        assert_eq!(
            decode_axis_title(
                &source,
                DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_output_bytes(3),
            )
            .expect_err("output cap")
            .output_limit_values(),
            Some((4, 3))
        );
        assert_eq!(
            decode_axis_title(
                &source,
                DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_title_bytes(3),
            )
            .expect_err("title cap")
            .title_limit_values(),
            Some((4, 3))
        );
        let nesting = vec![0x0b, 0x13, 0x14, 0x0c];
        let error = decode_axis_title(
            &nesting,
            DecodeOptions::new(nesting.len(), 8, nesting.len() * 8, 1),
        )
        .expect_err("nesting cap");
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn source_byte_and_configured_nesting_limits_fail_before_projection() {
        let source = field_varint(13, 1);
        assert_eq!(
            decode_axis_title(
                &source,
                DecodeOptions::new(source.len() - 1, 8, source.len() * 8, 8),
            )
            .expect_err("byte cap")
            .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
        let error = decode_axis_title(
            &source,
            DecodeOptions::new(source.len(), 8, source.len() * 8, 0),
        )
        .expect_err("zero nesting");
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 0,
                maximum: 64,
            })
        );
    }

    #[test]
    fn unknown_fields_are_wire_checked_but_not_materialized() {
        let mut source = field_varint(13, 1);
        source.extend([0x8a, 0x01, 0x01, 0xff]);
        let before = source.clone();
        let snapshot = decode_axis_title(&source, options(&source)).expect("unknown field");
        assert_eq!(snapshot.title_visible(), Some(true));
        assert_eq!(snapshot.raw(), before.as_slice());
        assert_eq!(source, before);

        let unknown_noncanonical = [0x8a, 0x81, 0x00, 0x01, 0xff];
        let error = decode_axis_title(&unknown_noncanonical, options(&unknown_noncanonical))
            .expect_err("noncanonical unknown key");
        assert_eq!(error.noncanonical_reason(), Some("protobuf field key"));
    }

    #[test]
    fn rewrite_is_wire_local_exact_noop_and_preserves_unknown_spans() {
        let mut source = field_varint(4000, 7);
        source.extend(field_varint(13, 1));
        source.extend(field_text(15, b"old"));
        source.extend(field_varint(4001, 8));
        let original_unknowns = [source[0..4].to_vec(), source[source.len() - 4..].to_vec()];
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let (noop, noop_report) = super::rewrite_axis_title_with_report(
            &source,
            AxisTitleWrite::new(Some(true), Some("old")),
            options,
        )
        .expect("exact no-op");
        assert_eq!(noop, source);
        assert!(!noop_report.changed());

        let (rewritten, report) = super::rewrite_axis_title_with_report(
            &source,
            AxisTitleWrite::new(Some(false), None),
            options,
        )
        .expect("wire-local rewrite");
        assert!(report.changed());
        assert_eq!(
            decode_axis_title(&rewritten, options)
                .expect("readback")
                .title_visible(),
            Some(false)
        );
        assert_eq!(
            decode_axis_title(&rewritten, options)
                .expect("readback")
                .title(),
            None
        );
        assert!(
            rewritten
                .windows(original_unknowns[0].len())
                .any(|window| { window == original_unknowns[0].as_slice() })
        );
        assert!(
            rewritten
                .windows(original_unknowns[1].len())
                .any(|window| { window == original_unknowns[1].as_slice() })
        );
    }

    #[test]
    fn rewrite_appends_missing_selected_fields_and_enforces_output_limit() {
        let source = field_varint(4000, 7);
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);
        let rewritten = super::rewrite_axis_title(
            &source,
            AxisTitleWrite::new(Some(true), Some("new")),
            options,
        )
        .expect("append selected fields");
        assert_eq!(
            decode_axis_title(&rewritten, options)
                .expect("readback")
                .visible_title(AxisTitleKind::Category),
            Some("new")
        );
        assert!(rewritten.starts_with(&source));

        let capped = options.with_max_output_bytes(source.len());
        let expected_output =
            source.len() + field_varint(13, 1).len() + field_text(15, b"new").len();
        assert_eq!(
            super::rewrite_axis_title(
                &source,
                AxisTitleWrite::new(Some(true), Some("new")),
                capped,
            )
            .expect_err("output cap")
            .output_limit_values(),
            Some((expected_output, source.len()))
        );
    }

    #[test]
    fn rewrite_replaces_selected_fields_in_place_and_keeps_unknown_order() {
        let source = [
            field_varint(4000, 7),
            field_text(15, b"old"),
            field_varint(4001, 8),
            field_varint(13, 0),
            field_varint(4002, 9),
        ]
        .concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let rewritten = super::rewrite_axis_title(
            &source,
            AxisTitleWrite::new(Some(true), Some("new")),
            options,
        )
        .expect("replace selected fields in place");
        let expected = [
            field_varint(4000, 7),
            field_text(15, b"new"),
            field_varint(4001, 8),
            field_varint(13, 1),
            field_varint(4002, 9),
        ]
        .concat();
        assert_eq!(rewritten, expected);
    }

    #[test]
    fn rewrite_does_not_guess_removed_field_positions() {
        // These two distinct source layouts intentionally collapse to the same
        // clear result. No source-local rewrite can know whether field 15 used
        // to precede or follow field 13 after it has been removed.
        let before_visible = [
            field_text(15, b"old"),
            field_varint(4000, 7),
            field_varint(13, 1),
        ]
        .concat();
        let after_visible = [
            field_varint(4000, 7),
            field_varint(13, 1),
            field_text(15, b"old"),
        ]
        .concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let cleared_before = super::rewrite_axis_title(
            &before_visible,
            AxisTitleWrite::new(Some(false), None),
            options,
        )
        .expect("clear title from before-visible layout");
        let cleared_after = super::rewrite_axis_title(
            &after_visible,
            AxisTitleWrite::new(Some(false), None),
            options,
        )
        .expect("clear title from after-visible layout");
        assert_eq!(cleared_before, cleared_after);
        assert_eq!(
            cleared_before,
            [field_varint(4000, 7), field_varint(13, 0)].concat()
        );

        let restored = super::rewrite_axis_title(
            &cleared_before,
            AxisTitleWrite::new(Some(true), Some("old")),
            options,
        )
        .expect("restore selected semantics");
        // Missing fields append after the source spans in field-number order;
        // they must not be guessed back into one of the historical layouts.
        assert_eq!(restored, after_visible);
        assert_ne!(restored, before_visible);
    }

    #[test]
    fn rewrite_clear_hidden_or_absent_stale_title_is_exact_noop() {
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);
        for source in [
            field_varint(13, 0),
            [field_varint(13, 0), field_text(15, b"stale")].concat(),
            field_text(15, b"stale"),
            Vec::new(),
        ] {
            let (rewritten, report) = super::rewrite_axis_title_with_report(
                &source,
                AxisTitleWrite::new(Some(false), None),
                options,
            )
            .expect("hidden/absent clear");
            assert_eq!(rewritten, source);
            assert!(!report.changed());
        }

        // A visible empty title is semantically present even when field 15 is
        // absent. Clearing and restoring that pair must retain the absence.
        let source = field_varint(13, 1);
        let cleared =
            super::rewrite_axis_title(&source, AxisTitleWrite::new(Some(false), None), options)
                .expect("clear visible title");
        let restored =
            super::rewrite_axis_title(&cleared, AxisTitleWrite::new(Some(true), None), options)
                .expect("restore visible empty title");
        assert_eq!(restored, source);
        assert_eq!(
            decode_axis_title(&restored, options)
                .expect("restored title")
                .visible_title(AxisTitleKind::Category),
            Some("")
        );
        assert_eq!(
            decode_axis_title(&restored, options)
                .expect("restored title")
                .title(),
            None
        );
    }

    #[test]
    fn rewrite_limits_are_inclusive_and_title_limit_precedes_output_allocation() {
        let source = field_varint(4000, 7);
        let write = AxisTitleWrite::new(Some(true), Some("new"));
        let base = DecodeOptions::new(1024, 128, 8192, 8).with_max_title_bytes(128);
        let expected_output =
            source.len() + field_varint(13, 1).len() + field_text(15, b"new").len();

        let (_, report) = super::rewrite_axis_title_with_report(
            &source,
            write,
            base.with_max_output_bytes(expected_output),
        )
        .expect("inclusive output limit");
        assert_eq!(report.output_bytes(), expected_output);

        let output_error = super::rewrite_axis_title(
            &source,
            write,
            base.with_max_output_bytes(expected_output - 1),
        )
        .expect_err("max-minus-one output limit");
        assert_eq!(
            output_error.output_limit_values(),
            Some((expected_output, expected_output - 1))
        );

        let title_error = super::rewrite_axis_title(&source, write, base.with_max_title_bytes(2))
            .expect_err("title limit");
        assert_eq!(title_error.title_limit_values(), Some((3, 2)));
    }

    #[test]
    fn rewrite_readback_accepts_larger_title_when_output_cap_allows_it() {
        let source = field_varint(13, 1);
        let title = "a title larger than the source";
        let write = AxisTitleWrite::new(Some(true), Some(title));
        let expected_output = [source.clone(), field_text(15, title.as_bytes())].concat();
        let options = DecodeOptions::new(source.len(), 128, 8192, 8)
            .with_max_output_bytes(expected_output.len())
            .with_max_title_bytes(title.len());

        let (rewritten, report) = super::rewrite_axis_title_with_report(&source, write, options)
            .expect("candidate larger than source fits output cap");
        assert_eq!(rewritten, expected_output);
        assert_eq!(report.output_bytes(), expected_output.len());
        let readback_options = DecodeOptions::new(expected_output.len(), 128, 8192, 8)
            .with_max_output_bytes(expected_output.len())
            .with_max_title_bytes(title.len());
        assert_eq!(
            decode_axis_title(&rewritten, readback_options)
                .expect("larger candidate readback")
                .title(),
            Some(title)
        );

        let capped = options.with_max_output_bytes(expected_output.len() - 1);
        let error = super::rewrite_axis_title(&source, write, capped)
            .expect_err("independent output cap must remain enforced");
        assert_eq!(
            error.output_limit_values(),
            Some((expected_output.len(), expected_output.len() - 1))
        );
    }

    #[test]
    fn rewrite_work_budget_covers_sizing_emission_and_readback() {
        let source = [field_varint(13, 1), field_text(15, b"old")].concat();
        let write = AxisTitleWrite::new(Some(true), Some("new title"));
        let unconstrained = DecodeOptions::new(1024, usize::MAX, usize::MAX, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);
        let (output, report) = super::rewrite_axis_title_with_report(&source, write, unconstrained)
            .expect("rewrite with an unconstrained work budget");

        // A flat payload is charged once for source decode, once for sizing,
        // once for emission, and once for candidate readback. Each pass
        // charges two bytes per source byte for strict plus Buffa traversal.
        let expected_work = source.len() * 6 + output.len() * 2;
        assert_eq!(report.work_bytes(), expected_work);
        assert_eq!(report.fields(), 8);

        let exact = DecodeOptions::new(1024, usize::MAX, expected_work, 8)
            .with_max_output_bytes(output.len())
            // Source and candidate title bytes each fit this per-pass cap;
            // they must not be accumulated as one semantic title.
            .with_max_title_bytes(9);
        super::rewrite_axis_title_with_report(&source, write, exact)
            .expect("the exact aggregate work ceiling is inclusive");

        let exact_fields = DecodeOptions::new(1024, 8, usize::MAX, 8)
            .with_max_output_bytes(output.len())
            .with_max_title_bytes(9);
        super::rewrite_axis_title_with_report(&source, write, exact_fields)
            .expect("the exact aggregate field ceiling is inclusive");

        let below_fields = DecodeOptions::new(1024, 7, usize::MAX, 8)
            .with_max_output_bytes(output.len())
            .with_max_title_bytes(9);
        let field_error = super::rewrite_axis_title_with_report(&source, write, below_fields)
            .expect_err("one field below aggregate accounting must fail");
        assert_eq!(field_error.field_limit_values(), Some((8, 7)));

        let below = DecodeOptions::new(1024, usize::MAX, expected_work - 1, 8)
            .with_max_output_bytes(output.len())
            .with_max_title_bytes(9);
        let error = super::rewrite_axis_title_with_report(&source, write, below)
            .expect_err("one byte below aggregate work must fail");
        assert_eq!(
            error.work_limit_values(),
            Some((expected_work, expected_work - 1))
        );
        assert_eq!(
            source,
            [field_varint(13, 1), field_text(15, b"old")].concat()
        );
    }

    #[test]
    fn rewrite_matching_semantics_still_preflight_duplicates_and_unknown_wire() {
        let mut duplicate = field_varint(13, 1);
        duplicate.extend(field_varint(13, 1));
        let duplicate_error = super::rewrite_axis_title(
            &duplicate,
            AxisTitleWrite::new(Some(true), None),
            options(&duplicate),
        )
        .expect_err("duplicate selected field");
        assert_eq!(
            duplicate_error.duplicate_singular_field(),
            Some("TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxiscategoryshowtitle")
        );

        let unknown_noncanonical =
            [field_varint(13, 1), vec![0x8a, 0x81, 0x00, 0x01, 0xff]].concat();
        let unknown_error = super::rewrite_axis_title(
            &unknown_noncanonical,
            AxisTitleWrite::new(Some(true), None),
            options(&unknown_noncanonical),
        )
        .expect_err("malformed unknown field");
        assert_eq!(
            unknown_error.noncanonical_reason(),
            Some("protobuf field key")
        );
    }

    #[test]
    fn unknown_groups_are_counted_preserved_and_depth_limited() {
        let mut source = vec![0x53]; // field 10, start group
        source.extend(field_varint(4000, 7));
        source.push(0x54); // field 10, end group
        source.extend(field_varint(13, 1));
        // A rewrite charges the source preflight, sizing scan, emission scan,
        // and candidate readback against one aggregate field budget. The
        // group contributes three visits plus the selected field per pass.
        let options = DecodeOptions::new(source.len(), 16, source.len() * 8, 2)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let (_, report) =
            super::decode_axis_title_with_report(&source, options).expect("matched unknown group");
        assert_eq!(report.fields(), 4);
        assert_eq!(report.max_depth(), 2);

        let rewritten =
            super::rewrite_axis_title(&source, AxisTitleWrite::new(Some(false), None), options)
                .expect("preserve unknown group");
        assert_eq!(&rewritten[..source.len() - 2], &source[..source.len() - 2]);

        let error = decode_axis_title(
            &source,
            DecodeOptions::new(source.len(), 8, source.len() * 8, 1),
        )
        .expect_err("group nesting limit");
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn output_reservation_reports_capacity_overflow_without_partial_output() {
        let error = super::reserve_output(usize::MAX).expect_err("capacity overflow");
        assert_eq!(error.allocation_amount(), Some(usize::MAX));
    }

    #[test]
    fn output_reservation_enforces_exact_capacity_or_typed_allocation_error() {
        match super::reserve_output(17) {
            Ok(output) => assert_eq!(output.capacity(), 17),
            Err(error) => assert_eq!(error.allocation_amount(), Some(17)),
        }
    }

    #[test]
    fn resource_limit_enum_maps_all_focused_classes() {
        let source = field_text(15, b"x");
        let error = decode_axis_title(
            &source,
            DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_title_bytes(0),
        )
        .expect_err("title limit");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Title {
                observed: 1,
                maximum: 0,
            })
        );
    }

    #[test]
    fn preserve_write_updates_only_the_selected_axis() {
        let source = [
            field_varint(4000, 7),
            field_varint(14, 1),
            field_text(15, b"old category"),
            field_varint(13, 1),
            field_text(16, b"untouched value"),
            field_varint(4001, 8),
        ]
        .concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);
        let rewritten = super::rewrite_axis_title(
            &source,
            AxisTitleWrite::preserve().with_title(AxisTitleKind::Category, Some("new category")),
            options,
        )
        .expect("category-only rewrite");
        let expected = [
            field_varint(4000, 7),
            field_varint(14, 1),
            field_text(15, b"new category"),
            field_varint(13, 1),
            field_text(16, b"untouched value"),
            field_varint(4001, 8),
        ]
        .concat();
        assert_eq!(rewritten, expected);
        let snapshot = decode_axis_title(&rewritten, options).expect("dual-axis readback");
        assert_eq!(
            snapshot.visible_title(AxisTitleKind::Category),
            Some("new category")
        );
        assert_eq!(
            snapshot.visible_title(AxisTitleKind::Value),
            Some("untouched value")
        );
        assert_eq!(snapshot.value_title_visible(), Some(true));
        assert_eq!(snapshot.value_title(), Some("untouched value"));
    }

    #[test]
    fn value_clear_removes_only_value_text_and_absent_clear_is_exact_noop() {
        let source = [
            field_text(15, b"category"),
            field_varint(14, 1),
            field_varint(4000, 7),
            field_text(16, b"value"),
            field_varint(13, 1),
        ]
        .concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);
        let cleared = super::rewrite_axis_title(
            &source,
            AxisTitleWrite::preserve().with_title(AxisTitleKind::Value, None),
            options,
        )
        .expect("clear value title");
        assert_eq!(
            cleared,
            [
                field_text(15, b"category"),
                field_varint(14, 0),
                field_varint(4000, 7),
                field_varint(13, 1),
            ]
            .concat()
        );
        let snapshot = decode_axis_title(&cleared, options).expect("cleared readback");
        assert_eq!(
            snapshot.visible_title(AxisTitleKind::Category),
            Some("category")
        );
        assert_eq!(snapshot.visible_title(AxisTitleKind::Value), None);

        let absent = [field_varint(13, 1), field_text(15, b"category")].concat();
        let (noop, report) = super::rewrite_axis_title_with_report(
            &absent,
            AxisTitleWrite::preserve().with_title(AxisTitleKind::Value, None),
            options,
        )
        .expect("absent value clear");
        assert_eq!(noop, absent);
        assert!(!report.changed());
    }

    #[test]
    fn prepared_rewrite_is_exact_and_each_execution_ceiling_rejects_max_minus_one() {
        let source = [
            field_varint(13, 1),
            field_text(15, b"category"),
            field_varint(14, 1),
            field_text(16, b"value"),
        ]
        .concat();
        let write =
            AxisTitleWrite::preserve().with_title(AxisTitleKind::Category, Some("longer category"));
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128)
            .with_max_allocations(8)
            .with_max_retained_bytes(2048)
            .with_max_scratch_bytes(1024);
        let prepared = super::prepare_axis_title_rewrite(&source, write, options)
            .expect("prepare bounded rewrite");
        assert_eq!(prepared.prepare_report().fields(), 4);
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(requirements.exact())
            .expect("execute exact requirements");
        assert_eq!(output.report().fields(), requirements.fields);
        assert_eq!(output.report().work_bytes(), requirements.work_bytes);
        assert_eq!(output.report().allocations(), requirements.allocations);
        assert_eq!(
            output.report().retained_bytes(),
            requirements.retained_bytes
        );
        assert_eq!(output.report().scratch_bytes(), requirements.scratch_bytes);

        let limits = [
            requirements
                .exact()
                .with_output_bytes(requirements.output_bytes - 1),
            requirements.exact().with_fields(requirements.fields - 1),
            requirements
                .exact()
                .with_work_bytes(requirements.work_bytes - 1),
            requirements
                .exact()
                .with_allocations(requirements.allocations - 1),
            requirements
                .exact()
                .with_retained_bytes(requirements.retained_bytes - 1),
            requirements
                .exact()
                .with_scratch_bytes(requirements.scratch_bytes - 1),
        ];
        for limit in limits {
            let error = super::prepare_axis_title_rewrite(&source, write, options)
                .expect("re-prepare one-shot plan")
                .execute(limit)
                .expect_err("max-minus-one execution ceiling");
            assert!(error.resource_limit().is_some());
        }
    }
}
