//! Strict source-preserving projection for Keynote chart value-axis settings.
//!
//! The native settings live in fields 4, 5, 6, 8, 17, and 18 of
//! `TSCH.Generated.ChartAxisNonStyleArchive`.  This module owns no chart
//! lookup or package transaction.  It first walks the complete protobuf wire
//! tree with a finite, canonical policy, then forces a private borrowed Buffa
//! view as an independent semantic cross-check.  Rewrites use the
//! preflighted source spans as their preservation authority; generated Buffa
//! values are never re-encoded and never escape this crate.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict wire reader intentionally precedes the generated view it cross-checks."
)]

use std::fmt;

use crate::buffa_keynote_chart_axis_value_settings_generated::LitchiIwaProjection as projection;
use buffa::DecodeOptions as BuffaDecodeOptions;

const CHART_AXIS_VALUE_DECADES_FIELD: u32 = 4;
const CHART_AXIS_VALUE_MAJOR_STEPS_FIELD: u32 = 5;
const CHART_AXIS_VALUE_MINOR_STEPS_FIELD: u32 = 6;
const CHART_AXIS_VALUE_SCALE_FIELD: u32 = 8;
const CHART_AXIS_VALUE_MAXIMUM_FIELD: u32 = 17;
const CHART_AXIS_VALUE_MINIMUM_FIELD: u32 = 18;
const MAX_RECURSION_LIMIT: u32 = 64;
const ROOT_DEPTH: u32 = 1;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;

/// A finite manual endpoint of a value-axis range.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct AxisValueBound(f64);

impl AxisValueBound {
    /// Build a finite endpoint.
    pub fn new(value: f64) -> Result<Self, ValueAxisSemanticError> {
        if value.is_finite() {
            Ok(Self(value))
        } else {
            Err(ValueAxisSemanticError::NonFiniteBound)
        }
    }

    /// Return the endpoint value.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for AxisValueBound {
    type Error = ValueAxisSemanticError;

    fn try_from(value: f64) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

/// Optional manual endpoints for a value axis.
#[derive(Debug, Clone, Copy)]
pub struct AxisValueBounds {
    minimum: Option<AxisValueBound>,
    maximum: Option<AxisValueBound>,
}

impl AxisValueBounds {
    /// Build optional endpoints, rejecting non-finite and inverted ranges.
    pub fn new(
        minimum: Option<AxisValueBound>,
        maximum: Option<AxisValueBound>,
    ) -> Result<Self, ValueAxisSemanticError> {
        if let (Some(minimum), Some(maximum)) = (minimum, maximum)
            && minimum.value() > maximum.value()
        {
            return Err(ValueAxisSemanticError::InvertedBounds);
        }
        Ok(Self { minimum, maximum })
    }

    /// Automatic lower and upper bounds.
    #[must_use]
    pub const fn automatic() -> Self {
        Self {
            minimum: None,
            maximum: None,
        }
    }

    /// Build a fully manual range.
    pub fn fixed(
        minimum: AxisValueBound,
        maximum: AxisValueBound,
    ) -> Result<Self, ValueAxisSemanticError> {
        Self::new(Some(minimum), Some(maximum))
    }

    /// Return the optional lower endpoint.
    #[must_use]
    pub const fn minimum(self) -> Option<AxisValueBound> {
        self.minimum
    }

    /// Return the optional upper endpoint.
    #[must_use]
    pub const fn maximum(self) -> Option<AxisValueBound> {
        self.maximum
    }
}

impl PartialEq for AxisValueBounds {
    // Bounds are compared by their semantic numeric values. In particular,
    // IEEE-754 positive and negative zero are equal here, so a rewrite that
    // only changes a zero's sign preserves the original fixed64 source span.
    fn eq(&self, other: &Self) -> bool {
        self.minimum.map(AxisValueBound::value) == other.minimum.map(AxisValueBound::value)
            && self.maximum.map(AxisValueBound::value) == other.maximum.map(AxisValueBound::value)
    }
}

/// A positive major-step count and non-negative minor-step count.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AxisValueSteps {
    major: Option<u32>,
    minor: Option<u32>,
}

impl AxisValueSteps {
    /// Build optional native step counts.
    pub fn new(major: Option<u32>, minor: Option<u32>) -> Result<Self, ValueAxisSemanticError> {
        if major.is_some_and(|value| value == 0) {
            return Err(ValueAxisSemanticError::MajorStepsZero);
        }
        if major.is_some_and(|value| value > i32::MAX as u32) {
            return Err(ValueAxisSemanticError::MajorStepsOutOfRange);
        }
        if minor.is_some_and(|value| value > i32::MAX as u32) {
            return Err(ValueAxisSemanticError::MinorStepsOutOfRange);
        }
        Ok(Self { major, minor })
    }

    /// Automatic major and minor steps.
    #[must_use]
    pub const fn automatic() -> Self {
        Self {
            major: None,
            minor: None,
        }
    }

    /// Return the optional major-step count.
    #[must_use]
    pub const fn major(self) -> Option<u32> {
        self.major
    }

    /// Return the optional minor-step count.
    #[must_use]
    pub const fn minor(self) -> Option<u32> {
        self.minor
    }
}

/// Native value-axis scale, retaining future enum values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Scale {
    /// Linear value-axis scale.
    #[default]
    Linear,
    /// Logarithmic value-axis scale.
    Logarithmic,
    /// A future or otherwise unrecognized native integer.
    Unsupported(i32),
}

impl Scale {
    /// Decode the native signed enum integer.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        match value {
            1 => Self::Linear,
            2 => Self::Logarithmic,
            other => Self::Unsupported(other),
        }
    }

    /// Return the native signed enum integer.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        match self {
            Self::Linear => 1,
            Self::Logarithmic => 2,
            Self::Unsupported(value) => value,
        }
    }
}

/// Semantic construction failures for this projection's value types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValueAxisSemanticError {
    /// A bound was NaN or infinite.
    NonFiniteBound,
    /// Minimum is greater than maximum.
    InvertedBounds,
    /// Major steps must be positive.
    MajorStepsZero,
    /// Major steps do not fit the native signed field.
    MajorStepsOutOfRange,
    /// Minor steps do not fit the native signed field.
    MinorStepsOutOfRange,
}

impl fmt::Display for ValueAxisSemanticError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NonFiniteBound => "chart value-axis bound must be finite",
            Self::InvertedBounds => "chart value-axis minimum exceeds maximum",
            Self::MajorStepsZero => "chart value-axis major step count must be positive",
            Self::MajorStepsOutOfRange => {
                "chart value-axis major step count exceeds the native signed range"
            },
            Self::MinorStepsOutOfRange => {
                "chart value-axis minor step count exceeds the native signed range"
            },
        })
    }
}

impl std::error::Error for ValueAxisSemanticError {}

/// The semantic value-axis settings represented by one generated extension.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AxisValueSettings {
    bounds: AxisValueBounds,
    steps: AxisValueSteps,
    scale: Scale,
}

impl AxisValueSettings {
    /// Build settings from validated semantic components.
    #[must_use]
    pub const fn new(bounds: AxisValueBounds, steps: AxisValueSteps, scale: Scale) -> Self {
        Self {
            bounds,
            steps,
            scale,
        }
    }

    /// Automatic/default settings.
    #[must_use]
    pub const fn automatic() -> Self {
        Self::new(
            AxisValueBounds::automatic(),
            AxisValueSteps::automatic(),
            Scale::Linear,
        )
    }

    /// Return bounds.
    #[must_use]
    pub const fn bounds(self) -> AxisValueBounds {
        self.bounds
    }

    /// Return step settings.
    #[must_use]
    pub const fn steps(self) -> AxisValueSteps {
        self.steps
    }

    /// Return scale, with absent native scale mapped to linear by snapshots.
    #[must_use]
    pub const fn scale(self) -> Scale {
        self.scale
    }
}

/// Borrowed semantic facts from one generated chart-axis extension.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AxisValueSettingsSnapshot<'source> {
    settings: AxisValueSettings,
    decades: Option<i32>,
    scale_present: bool,
    raw: &'source [u8],
    index: SpanIndex,
}

impl<'source> AxisValueSettingsSnapshot<'source> {
    /// Return the aggregate semantic settings.
    #[must_use]
    pub const fn settings(self) -> AxisValueSettings {
        self.settings
    }

    /// Alias useful to format adapters.
    #[must_use]
    pub const fn value_axis_settings(self) -> AxisValueSettings {
        self.settings
    }

    /// Return optional manual bounds.
    #[must_use]
    pub const fn bounds(self) -> AxisValueBounds {
        self.settings.bounds
    }

    /// Return optional step counts.
    #[must_use]
    pub const fn steps(self) -> AxisValueSteps {
        self.settings.steps
    }

    /// Return effective scale; absent native field is linear.
    #[must_use]
    pub const fn scale(self) -> Scale {
        self.settings.scale
    }

    /// Return the adjacent decades field without interpreting it.
    #[must_use]
    pub const fn decades(self) -> Option<i32> {
        self.decades
    }

    /// Whether field 8 was explicitly present in the source.
    #[must_use]
    pub const fn scale_present(self) -> bool {
        self.scale_present
    }

    /// Return the explicitly present scale, if any.
    #[must_use]
    pub fn explicit_scale(self) -> Option<Scale> {
        self.scale_present.then_some(self.settings.scale)
    }

    /// Return exact caller-owned bytes used by the snapshot.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Presence-preserving rewrite request for the five settings fields.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AxisValueSettingsWrite {
    bounds: ComponentUpdate<AxisValueBounds>,
    steps: ComponentUpdate<AxisValueSteps>,
    scale: ScaleUpdate,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum ComponentUpdate<T> {
    Preserve,
    Replace(T),
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum ScaleUpdate {
    Preserve,
    Replace { value: Scale, explicit: bool },
    Clear,
}

impl AxisValueSettingsWrite {
    /// Start a source-preserving rewrite.
    #[must_use]
    pub const fn preserve() -> Self {
        Self {
            bounds: ComponentUpdate::Preserve,
            steps: ComponentUpdate::Preserve,
            scale: ScaleUpdate::Preserve,
        }
    }

    /// Replace all semantic settings, retaining effective linear defaults.
    #[must_use]
    pub const fn new(settings: AxisValueSettings) -> Self {
        Self {
            bounds: ComponentUpdate::Replace(settings.bounds),
            steps: ComponentUpdate::Replace(settings.steps),
            scale: ScaleUpdate::Replace {
                value: settings.scale,
                explicit: true,
            },
        }
    }

    /// Replace bounds while preserving steps and scale.
    #[must_use]
    pub const fn with_bounds(mut self, bounds: AxisValueBounds) -> Self {
        self.bounds = ComponentUpdate::Replace(bounds);
        self
    }

    /// Replace steps while preserving bounds and scale.
    #[must_use]
    pub const fn with_steps(mut self, steps: AxisValueSteps) -> Self {
        self.steps = ComponentUpdate::Replace(steps);
        self
    }

    /// Set the effective scale. Linear is not forced into an absent field.
    #[must_use]
    pub const fn with_scale(mut self, scale: Scale) -> Self {
        self.scale = ScaleUpdate::Replace {
            value: scale,
            explicit: false,
        };
        self
    }

    /// Set an explicitly present native scale, including linear.
    #[must_use]
    pub const fn with_explicit_scale(mut self, scale: Scale) -> Self {
        self.scale = ScaleUpdate::Replace {
            value: scale,
            explicit: true,
        };
        self
    }

    /// Remove field 8, restoring native automatic/default scale presence.
    #[must_use]
    pub const fn clear_scale(mut self) -> Self {
        self.scale = ScaleUpdate::Clear;
        self
    }

    /// Return a request that replaces all components.
    #[must_use]
    pub const fn replace(settings: AxisValueSettings) -> Self {
        Self::new(settings)
    }
}

/// Exact resource ceilings for one decode or prepared rewrite.
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
    /// Construct explicit finite wire and execution limits.
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

    /// Build finite limits derived from one source length.
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        let fields = checked_scale(bytes, 16)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT);
        let work = checked_scale(bytes, 32)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT);
        let retained = checked_scale(bytes, 2)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT);
        Self::new(bytes, fields.max(1), work.max(1), 8)
            .with_max_allocations(
                bytes
                    .checked_add(8)
                    .unwrap_or(MAX_AUTOMATIC_LIMIT)
                    .min(MAX_AUTOMATIC_LIMIT),
            )
            .with_max_retained_bytes(retained.max(bytes))
            .with_max_scratch_bytes(bytes)
    }

    /// Replace output bytes ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace retained-byte ceiling.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace scratch-byte ceiling.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Replace field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace recursion ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(self.max_message_bytes)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Exact resource consumption from one decode.
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
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Exact prepared rewrite requirements.
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
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub allocations: usize,
    pub retained_bytes: usize,
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from requirements.
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

/// Candidate bytes produced by one prepared rewrite.
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

/// Exact complete execution accounting.
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
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Message bytes exceeded their ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Field visits exceeded their ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate strict/Buffa work exceeded its ceiling.
    Work { observed: usize, maximum: usize },
    /// Candidate bytes exceeded their ceiling.
    Output { observed: usize, maximum: usize },
    /// Protobuf nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Prepared logical allocations exceeded their ceiling.
    Allocations { observed: usize, maximum: usize },
    /// Retained source plus candidate bytes exceeded their ceiling.
    Retained { observed: usize, maximum: usize },
    /// Candidate scratch bytes exceeded their ceiling.
    Scratch { observed: usize, maximum: usize },
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
    MissingNested(&'static str),
    NonFinite(&'static str),
    InvertedBounds,
    InvalidStep(&'static str),
    Projection,
    Allocation { amount: usize },
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

    /// Return a typed resource failure, if any.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return the failed allocation size, if any.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            _ => None,
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

    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingNested(field),
        }
    }

    const fn nonfinite(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonFinite(field),
        }
    }

    const fn invalid_step(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidStep(field),
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
                "Keynote chart value-axis settings byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings output has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings nesting {observed} exceeds maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "Keynote chart value-axis settings requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::DuplicateSingular(field) => write!(formatter, "duplicate singular field {field}"),
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::MissingNested(field) => write!(formatter, "{field} is missing its numeric value"),
            DecodeErrorKind::NonFinite(field) => write!(formatter, "{field} is not finite"),
            DecodeErrorKind::InvertedBounds => formatter.write_str("chart value-axis minimum exceeds maximum"),
            DecodeErrorKind::InvalidStep(field) => write!(formatter, "invalid chart value-axis {field} step count"),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote chart value-axis settings strict preflight disagrees with the Buffa projection",
            ),
            DecodeErrorKind::Allocation { amount } => {
                write!(formatter, "cannot allocate {amount} bytes for chart value-axis settings")
            },
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

/// One source-preserving rewrite prepared without allocating its candidate.
#[derive(Debug, Clone, Copy)]
pub struct PreparedAxisValueSettingsRewrite<'source> {
    source: &'source [u8],
    current: AxisValueSettingsSnapshot<'source>,
    write: AxisValueSettingsWrite,
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
}

impl PreparedAxisValueSettingsRewrite<'_> {
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

    /// Execute once after checking caller-provided ceilings.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_execution_limits(self.requirements, limits)?;
        let output = emit_rewrite(
            self.source,
            self.current,
            self.write,
            self.current.index,
            self.requirements.output_bytes,
        )?;
        let readback_options = DecodeOptions::new(
            output.len(),
            self.candidate_fields,
            self.candidate_work,
            self.options.recursion_limit,
        )
        .with_max_output_bytes(self.options.max_output_bytes)
        .with_max_allocations(self.options.max_allocations)
        .with_max_retained_bytes(self.options.max_retained_bytes)
        .with_max_scratch_bytes(self.options.max_scratch_bytes);
        let (readback, candidate_report) =
            decode_axis_value_settings_with_report(&output, readback_options)?;
        if candidate_report.fields != self.candidate_fields
            || candidate_report.work_bytes != self.candidate_work
            || candidate_report.max_depth != self.candidate_depth
            || !candidate_matches(readback, self.current, self.write)
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

/// Decode one generated chart-axis extension.
pub fn decode_axis_value_settings<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<AxisValueSettingsSnapshot<'source>, DecodeError> {
    Ok(decode_axis_value_settings_with_report(source, options)?.0)
}

/// Decode one generated chart-axis extension and return exact accounting.
pub fn decode_axis_value_settings_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(AxisValueSettingsSnapshot<'source>, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = decode_with_budget(source, options, &mut budget)?;
    let report = budget.report();
    check_decode_report_limits(report, options)?;
    Ok((snapshot, report))
}

/// Compatibility spelling for callers that refer to the generated extension.
pub fn decode_axis_value_settings_extension<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<AxisValueSettingsSnapshot<'source>, DecodeError> {
    decode_axis_value_settings(source, options)
}

/// Prepare a source-preserving settings rewrite without allocating output.
pub fn prepare_axis_value_settings_rewrite<'source>(
    source: &'source [u8],
    write: AxisValueSettingsWrite,
    options: DecodeOptions,
) -> Result<PreparedAxisValueSettingsRewrite<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let current = decode_with_budget(source, options, &mut budget)?;
    let source_report = budget.report();
    let write = normalize_write(current, write);
    let measure = measure_rewrite_output(source, current, write, current.index)?;
    let hard_output_limit = finite_hard_message_limit();
    if measure.output_bytes > hard_output_limit {
        return Err(resource_error(DecodeLimit::Output {
            observed: measure.output_bytes,
            maximum: hard_output_limit,
        }));
    }
    let candidate_work = candidate_work_estimate(measure.output_bytes, measure.nested_bytes)?;
    let candidate_fields = candidate_field_count(current.index, write)?;
    let fields = source_report
        .fields
        .checked_add(candidate_fields)
        .and_then(|value| value.checked_add(measure.rewrite_fields))
        .ok_or_else(|| field_size_overflow(options.max_fields))?;
    let work_bytes = source_report
        .work_bytes
        .checked_add(
            measure
                .output_bytes
                .checked_mul(2)
                .ok_or_else(|| work_size_overflow(options.max_work_bytes))?,
        )
        .and_then(|value| value.checked_add(candidate_work))
        .ok_or_else(|| work_size_overflow(options.max_work_bytes))?;
    let retained_bytes = source
        .len()
        .checked_add(measure.output_bytes)
        .ok_or_else(|| retained_size_overflow(options.max_retained_bytes))?;
    let allocations = usize::from(measure.output_bytes != 0);
    let requirements = RewriteExecutionRequirements {
        output_bytes: measure.output_bytes,
        fields,
        work_bytes,
        max_depth: current.index.max_depth.max(measure.max_depth),
        allocations,
        retained_bytes,
        scratch_bytes: measure.output_bytes,
    };
    check_option_execution_limits(requirements, options)?;
    Ok(PreparedAxisValueSettingsRewrite {
        source,
        current,
        write,
        options,
        source_report,
        requirements,
        candidate_fields,
        candidate_work,
        candidate_depth: measure.max_depth,
    })
}

/// Rewrite selected settings fields while preserving all unselected source.
pub fn rewrite_axis_value_settings(
    source: &[u8],
    write: AxisValueSettingsWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_axis_value_settings_with_report(source, write, options)?.0)
}

/// Rewrite selected settings fields and return exact accounting.
pub fn rewrite_axis_value_settings_with_report(
    source: &[u8],
    write: AxisValueSettingsWrite,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_axis_value_settings_rewrite(source, write, options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report;
    Ok((output.into_output(), report))
}

/// Explicit generated-extension spelling for the source-local rewrite.
pub fn rewrite_axis_value_settings_extension(
    source: &[u8],
    write: AxisValueSettingsWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_axis_value_settings(source, write, options)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| {
        resource_error(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: finite_hard_message_limit(),
        })
    })?;
    if options.max_message_bytes > hard_maximum || source.len() > options.max_message_bytes {
        return Err(resource_error(DecodeLimit::Bytes {
            observed: source.len().max(options.max_message_bytes),
            maximum: hard_maximum.min(options.max_message_bytes),
        }));
    }
    if options.max_message_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if source.len() > options.max_retained_bytes {
        return Err(resource_error(DecodeLimit::Retained {
            observed: source.len(),
            maximum: options.max_retained_bytes,
        }));
    }
    if options.max_fields > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Fields {
            observed: options.max_fields,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_work_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Work {
            observed: options.max_work_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_output_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Output {
            observed: options.max_output_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_allocations > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Allocations {
            observed: options.max_allocations,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_retained_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Retained {
            observed: options.max_retained_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.max_scratch_bytes > MAX_AUTOMATIC_LIMIT {
        return Err(resource_error(DecodeLimit::Scratch {
            observed: options.max_scratch_bytes,
            maximum: MAX_AUTOMATIC_LIMIT,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(resource_error(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION_LIMIT,
        }));
    }
    Ok(())
}

fn finite_hard_message_limit() -> usize {
    usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .unwrap_or(MAX_AUTOMATIC_LIMIT)
        .min(MAX_AUTOMATIC_LIMIT)
}

fn output_size_overflow() -> DecodeError {
    let maximum = finite_hard_message_limit();
    resource_error(DecodeLimit::Output {
        observed: maximum.checked_add(1).unwrap_or(maximum),
        maximum,
    })
}

fn field_size_overflow(maximum: usize) -> DecodeError {
    let maximum = maximum.min(MAX_AUTOMATIC_LIMIT);
    field_limit_error(maximum.checked_add(1).unwrap_or(maximum), maximum)
}

fn work_size_overflow(maximum: usize) -> DecodeError {
    let maximum = maximum.min(MAX_AUTOMATIC_LIMIT);
    work_limit_error(maximum.checked_add(1).unwrap_or(maximum), maximum)
}

fn retained_size_overflow(maximum: usize) -> DecodeError {
    let maximum = maximum.min(MAX_AUTOMATIC_LIMIT);
    resource_error(DecodeLimit::Retained {
        observed: maximum.checked_add(1).unwrap_or(maximum),
        maximum,
    })
}

fn check_decode_report_limits(
    report: DecodeReport,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if report.allocations > options.max_allocations {
        return Err(resource_error(DecodeLimit::Allocations {
            observed: report.allocations,
            maximum: options.max_allocations,
        }));
    }
    if report.retained_bytes > options.max_retained_bytes {
        return Err(resource_error(DecodeLimit::Retained {
            observed: report.retained_bytes,
            maximum: options.max_retained_bytes,
        }));
    }
    if report.scratch_bytes > options.max_scratch_bytes {
        return Err(resource_error(DecodeLimit::Scratch {
            observed: report.scratch_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    Ok(())
}

fn resource_error(limit: DecodeLimit) -> DecodeError {
    DecodeError {
        kind: DecodeErrorKind::Resource(limit),
    }
}

fn field_limit_error(observed: usize, maximum: usize) -> DecodeError {
    resource_error(DecodeLimit::Fields { observed, maximum })
}

fn work_limit_error(observed: usize, maximum: usize) -> DecodeError {
    resource_error(DecodeLimit::Work { observed, maximum })
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
            return Err(resource_error(limit));
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(resource_error(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    Ok(())
}

fn checked_scale(value: usize, factor: usize) -> Option<usize> {
    value.checked_mul(factor)
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
            return Err(field_limit_error(observed, self.options.max_fields));
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
            return Err(work_limit_error(observed, self.options.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(resource_error(DecodeLimit::Nesting {
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct FieldSpan {
    number: u32,
    start: usize,
    end: usize,
    value_start: usize,
    value_end: usize,
    nested_inner_start: usize,
    nested_inner_end: usize,
    nested_fields: usize,
    nested_max_depth: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct SpanIndex {
    selected: [Option<FieldSpan>; INDEX_FIELD_COUNT],
    total_fields: usize,
    max_depth: u32,
    unselected_max_depth: u32,
}

impl SpanIndex {
    const EMPTY: Self = Self {
        selected: [None; INDEX_FIELD_COUNT],
        total_fields: 0,
        max_depth: 0,
        unselected_max_depth: ROOT_DEPTH,
    };

    fn set(&mut self, slot: usize, span: FieldSpan) -> Result<(), DecodeError> {
        if self.selected[slot].is_some() {
            return Err(DecodeError::duplicate(selected_field_name(slot)));
        }
        self.selected[slot] = Some(span);
        Ok(())
    }
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
    value_start: usize,
    value_end: usize,
    subtree_max_depth: u32,
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64(&'source [u8]),
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
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

    fn fixed64(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Fixed64)?;
        let StrictValue::Fixed64(value) = self.value else {
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

fn decode_with_budget<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<AxisValueSettingsSnapshot<'source>, DecodeError> {
    budget.message(source.len(), ROOT_DEPTH)?;
    let (strict, index) = preflight_axis_value_settings(source, budget)?;
    // The strict pass bounds every subtree before the generated borrowed view
    // is touched. Charge its source walk against the same budget.
    budget.buffa(source.len())?;
    let view: projection::ChartAxisValueSettingsArchiveView<'source> = options
        .buffa()
        .decode_view(source)
        .map_err(DecodeError::from)?;
    let projected_maximum = view
        .tschchartaxisdefaultusermax
        .as_option()
        .and_then(|number| number.number_archive);
    let projected_minimum = view
        .tschchartaxisdefaultusermin
        .as_option()
        .and_then(|number| number.number_archive);
    if view.tschchartaxisvaluenumberofdecades != strict.decades
        || view.tschchartaxisvaluenumberofmajorgridlines != strict.major_native
        || view.tschchartaxisvaluenumberofminorgridlines != strict.minor_native
        || view.tschchartaxisvaluescale != strict.scale_native
        || projected_maximum != strict.maximum
        || projected_minimum != strict.minimum
    {
        return Err(DecodeError::projection());
    }
    Ok(AxisValueSettingsSnapshot {
        settings: strict.settings,
        decades: strict.decades,
        scale_present: strict.scale_native.is_some(),
        raw: source,
        index,
    })
}

#[derive(Clone, Copy, Debug)]
struct StrictSnapshot {
    settings: AxisValueSettings,
    decades: Option<i32>,
    major_native: Option<i32>,
    minor_native: Option<i32>,
    scale_native: Option<i32>,
    minimum: Option<f64>,
    maximum: Option<f64>,
}

fn preflight_axis_value_settings(
    source: &[u8],
    budget: &mut Budget,
) -> Result<(StrictSnapshot, SpanIndex), DecodeError> {
    let mut index = SpanIndex::EMPTY;
    let mut decades = None;
    let mut major_native = None;
    let mut minor_native = None;
    let mut scale_native = None;
    let mut minimum = None;
    let mut maximum = None;
    let mut remaining = source;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = next_strict_field(&mut remaining, source, budget, ROOT_DEPTH)?;
        let end = source.len() - remaining.len();
        let Some(field) = item else {
            return Err(DecodeError::projection());
        };
        let slot = selected_slot(field.number);
        if let Some(slot) = slot {
            let mut span = FieldSpan {
                number: field.number,
                start,
                end,
                value_start: field.value_start,
                value_end: field.value_end,
                nested_inner_start: 0,
                nested_inner_end: 0,
                nested_fields: 1,
                nested_max_depth: ROOT_DEPTH,
            };
            match field.number {
                CHART_AXIS_VALUE_DECADES_FIELD => {
                    let value = require_canonical_i32(field.varint()?, "decades")?;
                    index.set(slot, span)?;
                    decades = Some(value);
                },
                CHART_AXIS_VALUE_MAJOR_STEPS_FIELD => {
                    let value = require_canonical_i32(field.varint()?, "major steps")?;
                    if value <= 0 {
                        return Err(DecodeError::invalid_step("major"));
                    }
                    index.set(slot, span)?;
                    major_native = Some(value);
                },
                CHART_AXIS_VALUE_MINOR_STEPS_FIELD => {
                    let value = require_canonical_i32(field.varint()?, "minor steps")?;
                    if value < 0 {
                        return Err(DecodeError::invalid_step("minor"));
                    }
                    index.set(slot, span)?;
                    minor_native = Some(value);
                },
                CHART_AXIS_VALUE_SCALE_FIELD => {
                    let value = require_canonical_i32(field.varint()?, "scale")?;
                    index.set(slot, span)?;
                    scale_native = Some(value);
                },
                CHART_AXIS_VALUE_MAXIMUM_FIELD | CHART_AXIS_VALUE_MINIMUM_FIELD => {
                    let nested = field.length_delimited()?;
                    let number = parse_number_archive(nested, budget, ROOT_DEPTH + 1)?;
                    span.nested_inner_start = number.inner_start;
                    span.nested_inner_end = number.inner_end;
                    span.nested_fields = number.fields;
                    span.nested_max_depth = number.max_depth;
                    span.nested_max_depth = number.max_depth;
                    index.set(slot, span)?;
                    if field.number == CHART_AXIS_VALUE_MAXIMUM_FIELD {
                        maximum = Some(number.value);
                    } else {
                        minimum = Some(number.value);
                    }
                },
                _ => return Err(DecodeError::projection()),
            }
        } else {
            // `field` has already been consumed recursively for groups, so
            // unknown fields and their descendants are fully budgeted while
            // their raw span remains caller-owned.
            let _ = (start, end);
            index.unselected_max_depth = index.unselected_max_depth.max(field.subtree_max_depth);
        }
    }
    if let (Some(minimum), Some(maximum)) = (minimum, maximum)
        && minimum > maximum
    {
        return Err(DecodeError {
            kind: DecodeErrorKind::InvertedBounds,
        });
    }
    let bounds = AxisValueBounds::new(minimum.map(AxisValueBound), maximum.map(AxisValueBound))
        .map_err(|error| match error {
            ValueAxisSemanticError::InvertedBounds => DecodeError {
                kind: DecodeErrorKind::InvertedBounds,
            },
            ValueAxisSemanticError::NonFiniteBound => DecodeError::nonfinite("bound"),
            _ => DecodeError::projection(),
        })?;
    let steps = AxisValueSteps::new(
        major_native.map(|value| value as u32),
        minor_native.map(|value| value as u32),
    )
    .map_err(|error| match error {
        ValueAxisSemanticError::MajorStepsZero | ValueAxisSemanticError::MajorStepsOutOfRange => {
            DecodeError::invalid_step("major")
        },
        ValueAxisSemanticError::MinorStepsOutOfRange => DecodeError::invalid_step("minor"),
        _ => DecodeError::projection(),
    })?;
    let settings = AxisValueSettings::new(
        bounds,
        steps,
        scale_native.map_or(Scale::Linear, Scale::from_native),
    );
    index.total_fields = budget.fields;
    index.max_depth = budget.max_depth;
    Ok((
        StrictSnapshot {
            settings,
            decades,
            major_native,
            minor_native,
            scale_native,
            minimum,
            maximum,
        },
        index,
    ))
}

#[derive(Clone, Copy, Debug)]
struct NumberArchive {
    value: f64,
    inner_start: usize,
    inner_end: usize,
    fields: usize,
    max_depth: u32,
}

fn parse_number_archive(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<NumberArchive, DecodeError> {
    budget.message(source.len(), depth)?;
    let mut remaining = source;
    let mut value = None;
    let mut inner_start = 0usize;
    let mut inner_end = 0usize;
    let before = budget.fields;
    let depth_before = budget.max_depth;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let Some(item) = parse_strict_field(&mut remaining, source, budget, depth)? else {
            break;
        };
        let end = source.len() - remaining.len();
        let ParseItem::Field(field) = item else {
            return Err(buffa::DecodeError::InvalidEndGroup(0).into());
        };
        if field.number != 1 {
            continue;
        }
        if value.is_some() {
            return Err(DecodeError::duplicate(
                "TSCH.ChartsNSNumberDoubleArchive.number_archive",
            ));
        }
        let bytes = field.fixed64()?;
        let bits = bytes
            .try_into()
            .map(u64::from_le_bytes)
            .map_err(|_error| DecodeError::projection())?;
        let decoded =
            require_canonical_f64(bits, "TSCH.ChartsNSNumberDoubleArchive.number_archive")?;
        value = Some(decoded);
        inner_start = start;
        inner_end = end;
    }
    let value = value
        .ok_or_else(|| DecodeError::missing("TSCH.ChartsNSNumberDoubleArchive.number_archive"))?;
    Ok(NumberArchive {
        value,
        inner_start,
        inner_end,
        fields: budget
            .fields
            .checked_sub(before)
            .ok_or_else(DecodeError::projection)?,
        max_depth: budget.max_depth.max(depth_before),
    })
}

fn selected_slot(number: u32) -> Option<usize> {
    match number {
        CHART_AXIS_VALUE_DECADES_FIELD => Some(0),
        CHART_AXIS_VALUE_MAJOR_STEPS_FIELD => Some(1),
        CHART_AXIS_VALUE_MINOR_STEPS_FIELD => Some(2),
        CHART_AXIS_VALUE_SCALE_FIELD => Some(3),
        CHART_AXIS_VALUE_MAXIMUM_FIELD => Some(4),
        CHART_AXIS_VALUE_MINIMUM_FIELD => Some(5),
        _ => None,
    }
}

const INDEX_FIELD_COUNT: usize = 6;

fn selected_field_name(slot: usize) -> &'static str {
    match slot {
        0 => "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvaluenumberofdecades",
        1 => "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvaluenumberofmajorgridlines",
        2 => "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvaluenumberofminorgridlines",
        3 => "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisvaluescale",
        4 => "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisdefaultusermax",
        5 => "TSCH.Generated.ChartAxisNonStyleArchive.tschchartaxisdefaultusermin",
        _ => "TSCH.Generated.ChartAxisNonStyleArchive.unknown",
    }
}

fn normalize_write(
    current: AxisValueSettingsSnapshot<'_>,
    write: AxisValueSettingsWrite,
) -> AxisValueSettingsWrite {
    let bounds = match write.bounds {
        ComponentUpdate::Replace(value) if value == current.bounds() => ComponentUpdate::Preserve,
        value => value,
    };
    let steps = match write.steps {
        ComponentUpdate::Replace(value) if value == current.steps() => ComponentUpdate::Preserve,
        value => value,
    };
    let scale = match write.scale {
        ScaleUpdate::Replace { value, explicit } => {
            let value = Scale::from_native(value.native_value());
            if value == current.scale() && (!explicit || current.scale_present()) {
                ScaleUpdate::Preserve
            } else {
                ScaleUpdate::Replace { value, explicit }
            }
        },
        value => value,
    };
    AxisValueSettingsWrite {
        bounds,
        steps,
        scale,
    }
}

#[derive(Clone, Copy, Debug)]
struct RewriteMeasure {
    output_bytes: usize,
    nested_bytes: usize,
    rewrite_fields: usize,
    max_depth: u32,
}

fn measure_rewrite_output(
    source: &[u8],
    _current: AxisValueSettingsSnapshot<'_>,
    write: AxisValueSettingsWrite,
    index: SpanIndex,
) -> Result<RewriteMeasure, DecodeError> {
    let mut output_bytes = source.len();
    let mut nested_bytes = 0usize;
    let mut rewrite_fields = 0usize;
    // Clearing both nested bounds may reduce the candidate depth below the
    // source depth, so derive this value from fields that survive.
    let mut max_depth = index.unselected_max_depth.max(ROOT_DEPTH);
    for slot in 1..INDEX_FIELD_COUNT {
        let Some(span) = index.selected[slot] else {
            continue;
        };
        let old_len = span.end - span.start;
        let update = update_for_index(write, slot);
        let new_len = replacement_len(source, slot, update, span)?;
        output_bytes = output_bytes
            .checked_sub(old_len)
            .and_then(|value| value.checked_add(new_len))
            .ok_or_else(output_size_overflow)?;
        if slot >= 4 && new_len != 0 {
            max_depth = max_depth.max(span.nested_max_depth);
            let payload_len = if span.nested_fields != 0 {
                let nested = &source[span.value_start..span.value_end];
                nested.len()
            } else {
                9
            };
            nested_bytes = nested_bytes
                .checked_add(payload_len)
                .ok_or_else(output_size_overflow)?;
            max_depth = max_depth.max(2);
        }
        rewrite_fields = rewrite_fields
            .checked_add(1)
            .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))?;
    }
    for slot in 1..INDEX_FIELD_COUNT {
        if index.selected[slot].is_some() {
            continue;
        }
        let update = update_for_index(write, slot);
        let appended = appended_len(slot, update)?;
        output_bytes = output_bytes
            .checked_add(appended)
            .ok_or_else(output_size_overflow)?;
        if appended != 0 {
            rewrite_fields = rewrite_fields
                .checked_add(if slot >= 4 { 2 } else { 1 })
                .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))?;
            if slot >= 4 {
                nested_bytes = nested_bytes
                    .checked_add(9)
                    .ok_or_else(output_size_overflow)?;
                max_depth = max_depth.max(2);
            }
        }
    }
    Ok(RewriteMeasure {
        output_bytes,
        nested_bytes,
        rewrite_fields,
        max_depth,
    })
}

fn update_for_slot(write: AxisValueSettingsWrite, slot: usize) -> Update {
    match slot {
        0 => Update::Preserve,
        1 => match write.steps {
            ComponentUpdate::Preserve => Update::Preserve,
            ComponentUpdate::Replace(value) => Update::Major(value.major()),
        },
        2 => match write.steps {
            ComponentUpdate::Preserve => Update::Preserve,
            ComponentUpdate::Replace(value) => Update::Minor(value.minor()),
        },
        3 => match write.scale {
            ScaleUpdate::Preserve => Update::Preserve,
            ScaleUpdate::Clear => Update::Scale(None),
            ScaleUpdate::Replace { value, .. } => Update::Scale(Some(value)),
        },
        4 => match write.bounds {
            ComponentUpdate::Preserve => Update::Preserve,
            ComponentUpdate::Replace(value) => Update::Maximum(value.maximum()),
        },
        5 => match write.bounds {
            ComponentUpdate::Preserve => Update::Preserve,
            ComponentUpdate::Replace(value) => Update::Minimum(value.minimum()),
        },
        _ => Update::Preserve,
    }
}

#[derive(Clone, Copy, Debug)]
enum Update {
    Preserve,
    Major(Option<u32>),
    Minor(Option<u32>),
    Scale(Option<Scale>),
    Maximum(Option<AxisValueBound>),
    Minimum(Option<AxisValueBound>),
}

fn update_for_index(write: AxisValueSettingsWrite, slot: usize) -> Update {
    // Field 4 is never selected by the semantic writer.
    if slot == 0 {
        Update::Preserve
    } else {
        update_for_slot(write, slot)
    }
}

fn replacement_len(
    source: &[u8],
    slot: usize,
    update: Update,
    span: FieldSpan,
) -> Result<usize, DecodeError> {
    let old = span.end - span.start;
    let value = match update {
        Update::Preserve => return Ok(old),
        Update::Major(value) | Update::Minor(value) => value.map_or(Ok(0), |value| {
            int32_field_len(field_number(slot), value as i32)
        })?,
        Update::Scale(value) => value.map_or(Ok(0), |value| {
            int32_field_len(field_number(slot), value.native_value())
        })?,
        Update::Maximum(value) | Update::Minimum(value) => {
            let Some(_bound) = value else { return Ok(0) };
            if span.nested_fields == 0 {
                bound_field_len(field_number(slot))?
            } else {
                let nested = &source[span.value_start..span.value_end];
                let old_inner_len = span
                    .nested_inner_end
                    .checked_sub(span.nested_inner_start)
                    .ok_or_else(output_size_overflow)?;
                let nested_len = nested
                    .len()
                    .checked_sub(old_inner_len)
                    .and_then(|length| length.checked_add(9))
                    .ok_or_else(output_size_overflow)?;
                length_delimited_len(field_number(slot), nested_len)?
            }
        },
    };
    Ok(value)
}

fn appended_len(slot: usize, update: Update) -> Result<usize, DecodeError> {
    Ok(match update {
        Update::Preserve => 0,
        Update::Major(value) | Update::Minor(value) => value.map_or(Ok(0), |value| {
            int32_field_len(field_number(slot), value as i32)
        })?,
        Update::Scale(value) => value.map_or(Ok(0), |value| {
            int32_field_len(field_number(slot), value.native_value())
        })?,
        Update::Maximum(value) | Update::Minimum(value) => {
            value.map_or(Ok(0), |_| bound_field_len(field_number(slot)))?
        },
    })
}

fn candidate_field_count(
    index: SpanIndex,
    write: AxisValueSettingsWrite,
) -> Result<usize, DecodeError> {
    let mut fields = index.total_fields;
    for slot in 1..INDEX_FIELD_COUNT {
        let update = update_for_index(write, slot);
        if let Some(span) = index.selected[slot] {
            let removed = match update {
                Update::Preserve
                | Update::Major(Some(_))
                | Update::Minor(Some(_))
                | Update::Scale(Some(_))
                | Update::Maximum(Some(_))
                | Update::Minimum(Some(_)) => 0,
                Update::Major(None) | Update::Minor(None) | Update::Scale(None) => 1,
                Update::Maximum(None) | Update::Minimum(None) => 1usize
                    .checked_add(span.nested_fields)
                    .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))?,
            };
            fields = fields
                .checked_sub(removed)
                .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))?;
        } else {
            let added = match update {
                Update::Preserve
                | Update::Major(None)
                | Update::Minor(None)
                | Update::Scale(None)
                | Update::Maximum(None)
                | Update::Minimum(None) => 0,
                Update::Major(Some(_)) | Update::Minor(Some(_)) | Update::Scale(Some(_)) => 1,
                Update::Maximum(Some(_)) | Update::Minimum(Some(_)) => 2,
            };
            fields = fields
                .checked_add(added)
                .ok_or_else(|| field_size_overflow(MAX_AUTOMATIC_LIMIT))?;
        }
    }
    Ok(fields)
}

fn candidate_work_estimate(output_bytes: usize, nested_bytes: usize) -> Result<usize, DecodeError> {
    output_bytes
        // Candidate strict root walk (2x) plus its private Buffa borrowed
        // projection walk (1x), followed by each nested number message (2x).
        .checked_mul(3)
        .and_then(|value| {
            nested_bytes
                .checked_mul(2)
                .and_then(|nested| value.checked_add(nested))
        })
        .ok_or_else(|| work_size_overflow(MAX_AUTOMATIC_LIMIT))
}

fn candidate_matches(
    candidate: AxisValueSettingsSnapshot<'_>,
    current: AxisValueSettingsSnapshot<'_>,
    write: AxisValueSettingsWrite,
) -> bool {
    let expected_bounds = match write.bounds {
        ComponentUpdate::Preserve => current.bounds(),
        ComponentUpdate::Replace(value) => value,
    };
    let expected_steps = match write.steps {
        ComponentUpdate::Preserve => current.steps(),
        ComponentUpdate::Replace(value) => value,
    };
    let (expected_scale, expected_presence) = match write.scale {
        ScaleUpdate::Preserve => (current.scale(), current.scale_present()),
        ScaleUpdate::Clear => (Scale::Linear, false),
        ScaleUpdate::Replace { value, .. } => (Scale::from_native(value.native_value()), true),
    };
    candidate.bounds() == expected_bounds
        && candidate.steps() == expected_steps
        && candidate.scale() == expected_scale
        && candidate.scale_present() == expected_presence
}

fn field_number(slot: usize) -> u32 {
    match slot {
        1 => CHART_AXIS_VALUE_MAJOR_STEPS_FIELD,
        2 => CHART_AXIS_VALUE_MINOR_STEPS_FIELD,
        3 => CHART_AXIS_VALUE_SCALE_FIELD,
        4 => CHART_AXIS_VALUE_MAXIMUM_FIELD,
        5 => CHART_AXIS_VALUE_MINIMUM_FIELD,
        _ => 0,
    }
}

fn int32_field_len(number: u32, value: i32) -> Result<usize, DecodeError> {
    let encoded = if value < 0 {
        u64::MAX - u64::from(value.unsigned_abs()) + 1
    } else {
        value as u64
    };
    varint_len(u64::from(number) << 3)
        .checked_add(varint_len(encoded))
        .ok_or_else(output_size_overflow)
}

fn bound_field_len(number: u32) -> Result<usize, DecodeError> {
    length_delimited_len(number, 9)
}

fn length_delimited_len(number: u32, length: usize) -> Result<usize, DecodeError> {
    let encoded_length = u64::try_from(length).map_err(|_error| output_size_overflow())?;
    varint_len((u64::from(number) << 3) | 2)
        .checked_add(varint_len(encoded_length))
        .and_then(|size| size.checked_add(length))
        .ok_or_else(output_size_overflow)
}

fn emit_rewrite(
    source: &[u8],
    current: AxisValueSettingsSnapshot<'_>,
    write: AxisValueSettingsWrite,
    index: SpanIndex,
    output_bytes: usize,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(output_bytes)?;
    let mut cursor = 0usize;
    loop {
        let next = index
            .selected
            .iter()
            .flatten()
            .filter(|span| span.start >= cursor)
            .min_by_key(|span| span.start);
        let Some(span) = next else {
            append_bytes(&mut output, &source[cursor..])?;
            break;
        };
        append_bytes(&mut output, &source[cursor..span.start])?;
        let slot = selected_slot(span.number).ok_or_else(DecodeError::projection)?;
        let update = update_for_index(write, slot);
        emit_existing_field(&mut output, source, *span, slot, update)?;
        cursor = span.end;
    }
    for slot in 1..INDEX_FIELD_COUNT {
        if index.selected[slot].is_none() {
            emit_appended_field(&mut output, slot, update_for_index(write, slot))?;
        }
    }
    let _ = current;
    if output.len() != output_bytes {
        return Err(DecodeError::projection());
    }
    Ok(output)
}

fn emit_existing_field(
    output: &mut Vec<u8>,
    source: &[u8],
    span: FieldSpan,
    slot: usize,
    update: Update,
) -> Result<(), DecodeError> {
    match update {
        Update::Preserve => append_bytes(output, &source[span.start..span.end])?,
        Update::Major(value) | Update::Minor(value) => {
            if let Some(value) = value {
                append_i32_field(output, field_number(slot), value as i32)?;
            }
        },
        Update::Scale(value) => {
            if let Some(value) = value {
                append_i32_field(output, field_number(slot), value.native_value())?;
            }
        },
        Update::Maximum(value) | Update::Minimum(value) => {
            if let Some(value) = value {
                append_bound_field(output, source, span, field_number(slot), value.value())?;
            }
        },
    }
    Ok(())
}

fn emit_appended_field(
    output: &mut Vec<u8>,
    slot: usize,
    update: Update,
) -> Result<(), DecodeError> {
    match update {
        Update::Preserve => {},
        Update::Major(value) | Update::Minor(value) => {
            if let Some(value) = value {
                append_i32_field(output, field_number(slot), value as i32)?;
            }
        },
        Update::Scale(value) => {
            if let Some(value) = value {
                append_i32_field(output, field_number(slot), value.native_value())?;
            }
        },
        Update::Maximum(value) | Update::Minimum(value) => {
            if let Some(value) = value {
                append_bound_field(
                    output,
                    &[],
                    FieldSpan::EMPTY,
                    field_number(slot),
                    value.value(),
                )?;
            }
        },
    }
    Ok(())
}

impl FieldSpan {
    const EMPTY: Self = Self {
        number: 0,
        start: 0,
        end: 0,
        value_start: 0,
        value_end: 0,
        nested_inner_start: 0,
        nested_inner_end: 0,
        nested_fields: 0,
        nested_max_depth: 0,
    };
}

fn append_bound_field(
    output: &mut Vec<u8>,
    source: &[u8],
    span: FieldSpan,
    number: u32,
    value: f64,
) -> Result<(), DecodeError> {
    append_varint(output, (u64::from(number) << 3) | 2)?;
    if span.nested_fields == 0 {
        append_varint(output, 9)?;
        append_fixed64_field(output, value.to_bits())?;
    } else {
        let nested = &source[span.value_start..span.value_end];
        let nested_length = u64::try_from(nested.len()).map_err(|_error| output_size_overflow())?;
        append_varint(output, nested_length)?;
        append_bytes(output, &nested[..span.nested_inner_start])?;
        append_fixed64_field(output, value.to_bits())?;
        append_bytes(output, &nested[span.nested_inner_end..])?;
    }
    Ok(())
}

fn append_i32_field(output: &mut Vec<u8>, number: u32, value: i32) -> Result<(), DecodeError> {
    append_varint(output, u64::from(number) << 3)?;
    let encoded = if value < 0 {
        u64::MAX - u64::from(value.unsigned_abs()) + 1
    } else {
        value as u64
    };
    append_varint(output, encoded)
}

fn append_fixed64_field(output: &mut Vec<u8>, value: u64) -> Result<(), DecodeError> {
    append_varint(output, 9)?;
    append_bytes(output, &value.to_le_bytes())
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_error| DecodeError {
            kind: DecodeErrorKind::Allocation { amount },
        })?;
    if output.capacity() != amount {
        return Err(DecodeError {
            kind: DecodeErrorKind::Allocation { amount },
        });
    }
    Ok(output)
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) -> Result<(), DecodeError> {
    while value >= 0x80 {
        append_byte(output, (value as u8 & 0x7f) | 0x80)?;
        value >>= 7;
    }
    append_byte(output, value as u8)
}

fn append_byte(output: &mut Vec<u8>, byte: u8) -> Result<(), DecodeError> {
    let required = output
        .len()
        .checked_add(1)
        .ok_or_else(output_size_overflow)?;
    if required > output.capacity() {
        return Err(DecodeError {
            kind: DecodeErrorKind::Allocation { amount: required },
        });
    }
    output.push(byte);
    Ok(())
}

fn append_bytes(output: &mut Vec<u8>, bytes: &[u8]) -> Result<(), DecodeError> {
    let required = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(output_size_overflow)?;
    if required > output.capacity() {
        return Err(DecodeError {
            kind: DecodeErrorKind::Allocation { amount: required },
        });
    }
    output.extend_from_slice(bytes);
    Ok(())
}

fn varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn require_canonical_i32(value: u64, field: &'static str) -> Result<i32, DecodeError> {
    let canonical = canonical_varint_len(value);
    let is_negative = value >= 0xffff_ffff_8000_0000;
    if (!is_negative && value > i32::MAX as u64)
        || (is_negative && value < 0xffff_ffff_8000_0000)
        || canonical
            != if is_negative {
                10
            } else {
                canonical_varint_len(value)
            }
    {
        return Err(DecodeError::noncanonical(field));
    }
    if is_negative {
        let signed = value as i64;
        if signed < i32::MIN as i64 {
            return Err(DecodeError::noncanonical(field));
        }
        Ok(signed as i32)
    } else {
        Ok(value as i32)
    }
}

fn require_canonical_f64(value: u64, field: &'static str) -> Result<f64, DecodeError> {
    let value = f64::from_bits(value);
    value
        .is_finite()
        .then_some(value)
        .ok_or_else(|| DecodeError::nonfinite(field))
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    origin: &'source [u8],
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
    let mut value_start = origin.len() - source.len();
    let (value, subtree_max_depth) = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            (StrictValue::Varint(value), depth)
        },
        buffa::encoding::WireType::Fixed64 => {
            let bytes = take_exact(source, 8)?;
            (StrictValue::Fixed64(bytes), depth)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            value_start = origin.len() - source.len();
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                depth,
            )
        },
        buffa::encoding::WireType::StartGroup => {
            let child_depth = depth.checked_add(1).ok_or_else(|| {
                resource_error(DecodeLimit::Nesting {
                    observed: u32::MAX,
                    maximum: budget.options.recursion_limit,
                })
            })?;
            budget.observe_depth(child_depth)?;
            (
                StrictValue::Group,
                skip_strict_group(source, origin, field_number, child_depth, budget)?,
            )
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            (StrictValue::Fixed32, depth)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    let value_end = origin.len() - source.len();
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
        value_start,
        value_end,
        subtree_max_depth,
    })))
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    origin: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, origin, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn skip_strict_group<'source>(
    source: &mut &'source [u8],
    origin: &'source [u8],
    expected_field_number: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<u32, DecodeError> {
    let mut max_depth = depth;
    loop {
        match parse_strict_field(source, origin, budget, depth)? {
            Some(ParseItem::Field(field)) => {
                max_depth = max_depth.max(field.subtree_max_depth);
            },
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => {
                return Ok(max_depth);
            },
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
    clippy::unwrap_used,
    reason = "Focused codec tests use small canonical wire fixtures."
)]
mod tests {
    use super::{
        AxisValueBound, AxisValueBounds, AxisValueSettings, AxisValueSettingsWrite, AxisValueSteps,
        DecodeLimit, DecodeOptions, Scale, decode_axis_value_settings,
        decode_axis_value_settings_with_report, prepare_axis_value_settings_rewrite,
        rewrite_axis_value_settings,
    };

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
        output
    }

    fn i32_value(value: i32) -> u64 {
        if value < 0 {
            u64::MAX - u64::from(value.unsigned_abs()) + 1
        } else {
            value as u64
        }
    }

    fn field_i32(number: u32, value: i32) -> Vec<u8> {
        let mut output = varint(u64::from(number) << 3);
        output.extend(varint(i32_value(value)));
        output
    }

    fn field_bytes(number: u32, value: &[u8]) -> Vec<u8> {
        let mut output = varint((u64::from(number) << 3) | 2);
        output.extend(varint(value.len() as u64));
        output.extend(value);
        output
    }

    fn field_fixed64(number: u32, value: u64) -> Vec<u8> {
        let mut output = varint(u64::from(number) << 3 | 1);
        output.extend(value.to_le_bytes());
        output
    }

    fn bound(number: u32, value: f64) -> Vec<u8> {
        bound_with_unknown(number, value, false)
    }

    fn bound_with_unknown(number: u32, value: f64, unknown: bool) -> Vec<u8> {
        let mut nested = field_fixed64(1, value.to_bits());
        if unknown {
            nested.extend(field_i32(7, 44));
        }
        field_bytes(number, &nested)
    }

    fn start_group(number: u32) -> Vec<u8> {
        varint(u64::from(number) << 3 | 3)
    }

    fn end_group(number: u32) -> Vec<u8> {
        varint(u64::from(number) << 3 | 4)
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
    }

    fn full_source() -> Vec<u8> {
        let mut source = field_i32(4, 3);
        source.extend(field_i32(5, 5));
        source.extend(field_i32(6, 2));
        source.extend(field_i32(8, 2));
        source.extend(bound(17, 30.0));
        source.extend(bound(18, -5.0));
        source.extend(field_i32(91, 7));
        source
    }

    #[test]
    fn empty_payload_maps_to_automatic_settings() {
        let source = [];
        let snapshot = decode_axis_value_settings(&source, options(&source)).expect("empty");
        assert_eq!(snapshot.bounds(), AxisValueBounds::automatic());
        assert_eq!(snapshot.steps(), AxisValueSteps::automatic());
        assert_eq!(snapshot.scale(), Scale::Linear);
        assert!(!snapshot.scale_present());
        assert_eq!(snapshot.decades(), None);
    }

    #[test]
    fn decodes_all_selected_fields_and_preserves_adjacent_decades() {
        let source = full_source();
        let (snapshot, report) =
            decode_axis_value_settings_with_report(&source, options(&source)).expect("full");
        assert_eq!(snapshot.decades(), Some(3));
        assert_eq!(snapshot.steps().major(), Some(5));
        assert_eq!(snapshot.steps().minor(), Some(2));
        assert_eq!(snapshot.scale(), Scale::Logarithmic);
        assert!(snapshot.scale_present());
        assert_eq!(snapshot.bounds().minimum().unwrap().value(), -5.0);
        assert_eq!(snapshot.bounds().maximum().unwrap().value(), 30.0);
        assert_eq!(report.source_bytes(), source.len());
        assert!(report.fields() >= 8);
        assert_eq!(snapshot.raw(), source.as_slice());
    }

    #[test]
    fn bounds_rewrite_keeps_nested_unknown_bytes_and_outer_order() {
        let mut source = field_i32(4, 11);
        source.extend(bound_with_unknown(18, -5.0, true));
        source.extend(field_i32(91, 7));
        source.extend(bound(17, 30.0));
        let replacement = AxisValueBounds::new(
            Some(AxisValueBound::new(-1.0).unwrap()),
            Some(AxisValueBound::new(40.0).unwrap()),
        )
        .unwrap();
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve().with_bounds(replacement),
            options(&source),
        )
        .expect("bounds rewrite");
        let old_unknown = field_i32(7, 44);
        assert!(
            output
                .windows(old_unknown.len())
                .any(|window| window == old_unknown)
        );
        assert!(output.windows(2).any(|window| window == [0x92, 0x01]));
        let snapshot = decode_axis_value_settings(&output, options(&output)).expect("readback");
        assert_eq!(snapshot.bounds().minimum().unwrap().value(), -1.0);
        assert_eq!(snapshot.bounds().maximum().unwrap().value(), 40.0);
        assert_eq!(snapshot.decades(), Some(11));
    }

    #[test]
    fn clear_bounds_removes_only_the_two_bound_fields() {
        let source = full_source();
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve().with_bounds(AxisValueBounds::automatic()),
            options(&source),
        )
        .expect("clear bounds");
        let snapshot = decode_axis_value_settings(&output, options(&output)).expect("readback");
        assert_eq!(snapshot.bounds(), AxisValueBounds::automatic());
        assert_eq!(snapshot.steps().major(), Some(5));
        assert_eq!(snapshot.steps().minor(), Some(2));
        assert_eq!(snapshot.scale(), Scale::Logarithmic);
        assert_eq!(snapshot.decades(), Some(3));
    }

    #[test]
    fn steps_and_scale_rewrite_atomically() {
        let source = full_source();
        let steps = AxisValueSteps::new(Some(7), Some(0)).unwrap();
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve()
                .with_steps(steps)
                .with_scale(Scale::Linear),
            options(&source),
        )
        .expect("combined rewrite");
        let snapshot = decode_axis_value_settings(&output, options(&output)).expect("readback");
        assert_eq!(snapshot.steps(), steps);
        assert_eq!(snapshot.scale(), Scale::Linear);
        assert!(snapshot.scale_present());
    }

    #[test]
    fn absent_linear_scale_stays_absent_for_effective_noop() {
        let source = field_i32(5, 4);
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve().with_scale(Scale::Linear),
            options(&source),
        )
        .expect("linear no-op");
        assert_eq!(output, source);
        assert!(
            !decode_axis_value_settings(&output, options(&output))
                .unwrap()
                .scale_present()
        );
    }

    #[test]
    fn signed_zero_bound_changes_are_semantic_noops_in_both_directions() {
        let positive_zero = AxisValueBound::new(0.0).unwrap();
        let negative_zero = AxisValueBound::new(-0.0).unwrap();
        let maximum = AxisValueBound::new(1.0).unwrap();
        let positive_source = {
            let mut source = bound(17, maximum.value());
            source.extend(bound(18, positive_zero.value()));
            source
        };
        let negative_request = AxisValueBounds::new(Some(negative_zero), Some(maximum)).unwrap();
        let positive_output = rewrite_axis_value_settings(
            &positive_source,
            AxisValueSettingsWrite::preserve().with_bounds(negative_request),
            options(&positive_source),
        )
        .expect("positive zero source");
        assert_eq!(positive_output, positive_source);

        let negative_source = {
            let mut source = bound(17, maximum.value());
            source.extend(bound(18, negative_zero.value()));
            source
        };
        let positive_request = AxisValueBounds::new(Some(positive_zero), Some(maximum)).unwrap();
        let negative_output = rewrite_axis_value_settings(
            &negative_source,
            AxisValueSettingsWrite::preserve().with_bounds(positive_request),
            options(&negative_source),
        )
        .expect("negative zero source");
        assert_eq!(negative_output, negative_source);
    }

    #[test]
    fn explicit_linear_and_unknown_negative_scale_round_trip() {
        let source = field_i32(8, -9);
        let snapshot = decode_axis_value_settings(&source, options(&source)).expect("negative");
        assert_eq!(snapshot.scale(), Scale::Unsupported(-9));
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve().with_explicit_scale(Scale::Linear),
            options(&source),
        )
        .expect("linear");
        assert_eq!(
            decode_axis_value_settings(&output, options(&output))
                .unwrap()
                .scale(),
            Scale::Linear
        );
        assert!(
            decode_axis_value_settings(&output, options(&output))
                .unwrap()
                .scale_present()
        );
        let expected = field_i32(8, 1);
        assert_eq!(output, expected);
    }

    #[test]
    fn prepared_execution_reports_exact_requirements() {
        let source = full_source();
        let write = AxisValueSettingsWrite::preserve()
            .with_steps(AxisValueSteps::new(Some(9), Some(1)).unwrap());
        let prepared =
            prepare_axis_value_settings_rewrite(&source, write, options(&source)).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).expect("execute");
        assert!(output.report().changed());
        assert_eq!(output.output().len(), requirements.output_bytes);
        assert_eq!(
            decode_axis_value_settings(output.output(), options(output.output()))
                .unwrap()
                .steps()
                .major(),
            Some(9)
        );
    }

    #[test]
    fn rejects_duplicate_wrong_wire_noncanonical_and_invalid_nested_values() {
        let duplicate = [0x28, 0x01, 0x28, 0x02];
        assert!(decode_axis_value_settings(&duplicate, options(&duplicate)).is_err());
        let wrong_wire = [0x2b, 0x01];
        assert!(decode_axis_value_settings(&wrong_wire, options(&wrong_wire)).is_err());
        let noncanonical = [0x28, 0x80, 0x00];
        assert!(decode_axis_value_settings(&noncanonical, options(&noncanonical)).is_err());
        let missing = field_bytes(17, &[]);
        assert!(decode_axis_value_settings(&missing, options(&missing)).is_err());
        for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            let nonfinite = bound(17, value);
            assert!(decode_axis_value_settings(&nonfinite, options(&nonfinite)).is_err());
        }
        let mut duplicate_inner = field_bytes(17, &{
            let mut value = field_fixed64(1, 1.0f64.to_bits());
            value.extend(field_fixed64(1, 2.0f64.to_bits()));
            value
        });
        assert!(decode_axis_value_settings(&duplicate_inner, options(&duplicate_inner)).is_err());
        let inverted = {
            let mut value = bound(18, 9.0);
            value.extend(bound(17, 1.0));
            value
        };
        assert!(decode_axis_value_settings(&inverted, options(&inverted)).is_err());
        duplicate_inner.clear();
    }

    #[test]
    fn rejects_duplicate_for_each_known_singular_field() {
        let fields = [
            field_i32(4, 1),
            field_i32(5, 1),
            field_i32(6, 0),
            field_i32(8, 1),
            bound(17, 1.0),
            bound(18, 1.0),
        ];
        for field in fields {
            let mut duplicate = field.clone();
            duplicate.extend(field);
            let error = decode_axis_value_settings(&duplicate, options(&duplicate))
                .expect_err("duplicate singular field");
            assert!(error.duplicate_singular_field().is_some());
        }
    }

    #[test]
    fn rejects_wrong_wire_and_truncation_for_each_known_field() {
        let wrong_wire = [
            field_fixed64(4, 1),
            field_fixed64(5, 1),
            field_fixed64(6, 1),
            field_fixed64(8, 1),
            field_i32(17, 1),
            field_i32(18, 1),
        ];
        for field in wrong_wire {
            assert!(decode_axis_value_settings(&field, options(&field)).is_err());
        }

        let truncated = [
            field_i32(4, 1),
            field_i32(5, 1),
            field_i32(6, 0),
            field_i32(8, 1),
            bound(17, 1.0),
            bound(18, 1.0),
        ];
        for mut field in truncated {
            field.pop();
            assert!(decode_axis_value_settings(&field, options(&field)).is_err());
        }
    }

    #[test]
    fn rejects_u32_style_negative_int32_and_preserves_nested_unknown_groups() {
        // A negative int32 is canonical only as a sign-extended ten-byte
        // varint. This five-byte u32 encoding must never be truncated/cast.
        let noncanonical_negative = [0x28, 0xff, 0xff, 0xff, 0xff, 0x0f];
        let error =
            decode_axis_value_settings(&noncanonical_negative, options(&noncanonical_negative))
                .expect_err("five-byte u32 negative");
        assert!(error.noncanonical_reason().is_some());

        let mut nested = field_fixed64(1, 10.0f64.to_bits());
        nested.extend(start_group(2));
        nested.extend(field_i32(3, 44));
        nested.extend(end_group(2));
        let mut source = field_bytes(17, &nested);
        source.extend(bound(18, 0.0));
        let replacement = AxisValueBounds::new(
            Some(AxisValueBound::new(0.0).unwrap()),
            Some(AxisValueBound::new(11.0).unwrap()),
        )
        .unwrap();
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve().with_bounds(replacement),
            options(&source),
        )
        .expect("nested unknown group preservation");
        let unknown_group = {
            let mut bytes = start_group(2);
            bytes.extend(field_i32(3, 44));
            bytes.extend(end_group(2));
            bytes
        };
        assert!(
            output
                .windows(unknown_group.len())
                .any(|window| window == unknown_group)
        );
        assert_eq!(
            decode_axis_value_settings(&output, options(&output))
                .unwrap()
                .bounds()
                .maximum()
                .unwrap()
                .value(),
            11.0
        );
    }

    #[test]
    fn rejects_unknown_group_depth_beyond_the_finite_limit() {
        let mut source = start_group(7);
        source.extend(start_group(8));
        source.extend(field_i32(9, 1));
        source.extend(end_group(8));
        source.extend(end_group(7));
        let error = decode_axis_value_settings(&source, options(&source).with_max_depth(2))
            .expect_err("nested unknown group depth");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn rejects_negative_and_zero_major_or_negative_minor_steps() {
        for source in [field_i32(5, 0), field_i32(5, -1), field_i32(6, -1)] {
            assert!(decode_axis_value_settings(&source, options(&source)).is_err());
        }
    }

    #[test]
    fn preserves_unknown_groups_and_rejects_unterminated_groups() {
        let mut source = vec![0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06];
        source.extend(field_i32(5, 2));
        let output = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve()
                .with_steps(AxisValueSteps::new(Some(3), None).unwrap()),
            options(&source),
        )
        .expect("group preservation");
        assert!(output.starts_with(&[0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06]));
        let unterminated = [0x9b, 0x06, 0x08, 0x01];
        assert!(decode_axis_value_settings(&unterminated, options(&unterminated)).is_err());
    }

    #[test]
    fn resource_limits_are_checked_before_candidate_allocation() {
        let source = full_source();
        let small = DecodeOptions::for_source(&source).with_max_output_bytes(source.len() - 1);
        let error = rewrite_axis_value_settings(
            &source,
            AxisValueSettingsWrite::preserve().with_scale(Scale::Logarithmic),
            small,
        )
        .expect_err("output limit");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Output { .. })
        ));
        let field_error = decode_axis_value_settings(
            &source,
            DecodeOptions::for_source(&source).with_max_fields(1),
        )
        .expect_err("field limit");
        assert!(matches!(
            field_error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let retained_error = decode_axis_value_settings(
            &source,
            DecodeOptions::for_source(&source).with_max_retained_bytes(0),
        )
        .expect_err("retained-byte limit");
        assert!(matches!(
            retained_error.resource_limit(),
            Some(DecodeLimit::Retained { .. })
        ));
    }

    #[test]
    fn aggregate_settings_constructor_is_copy_and_validated() {
        let bounds = AxisValueBounds::fixed(
            AxisValueBound::new(-1.0).unwrap(),
            AxisValueBound::new(1.0).unwrap(),
        )
        .unwrap();
        let steps = AxisValueSteps::new(Some(4), Some(2)).unwrap();
        let settings = AxisValueSettings::new(bounds, steps, Scale::Logarithmic);
        assert_eq!(settings.bounds(), bounds);
        assert_eq!(settings.steps(), steps);
        assert_eq!(settings.scale(), Scale::Logarithmic);
        assert!(
            AxisValueBounds::new(
                Some(AxisValueBound::new(2.0).unwrap()),
                Some(AxisValueBound::new(1.0).unwrap()),
            )
            .is_err()
        );
        assert!(AxisValueBound::new(f64::INFINITY).is_err());
    }
}
