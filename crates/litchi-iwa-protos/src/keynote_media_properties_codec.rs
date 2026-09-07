//! Strict, borrowed, source-preserving access to Keynote movie drawable
//! properties.
//!
//! The owned semantic edge is deliberately small:
//! `TSD.MovieArchive.super` -> `TSD.DrawableArchive.{hyperlink,locked,
//! aspect_ratio_locked,accessibility_description}`. The surrounding movie
//! graph, geometry, references, and every unknown field remain caller-owned
//! bytes. A private Buffa lazy view cross-checks the selected fields after the
//! strict scanner has established canonical wire and resource limits.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict source scanner is intentionally ordered before its private projection."
)]

use std::{fmt, mem::size_of};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_media_properties_generated::LitchiIwaKeynoteMoviePropertiesProjection as projection;

const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_HYPERLINK_FIELD: u32 = 4;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const DRAWABLE_ASPECT_RATIO_LOCKED_FIELD: u32 = 7;
const DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD: u32 = 8;
const MAX_RECURSION: u32 = 64;
const BUFFA_LOGICAL_ALLOCATIONS: usize = 2;
const EXECUTE_ALLOCATIONS: usize = 4;

/// A finite policy for one MovieArchive properties payload.
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
    /// Construct an explicit finite source/field/work/nesting policy.
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
            max_allocations: usize::MAX,
            max_retained_bytes: usize::MAX,
            max_scratch_bytes: usize::MAX,
        }
    }

    /// Construct a bounded convenience profile for one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(8).max(1),
            bytes.saturating_mul(32).max(1),
            16,
        )
        .with_max_output_bytes(bytes.saturating_mul(2).max(1))
        .with_max_allocations(bytes.saturating_mul(8).max(8))
        .with_max_retained_bytes(bytes.saturating_mul(16).max(1))
        .with_max_scratch_bytes(
            bytes
                .saturating_mul(size_of::<ParsedField>())
                .saturating_mul(16)
                .max(1),
        )
    }

    /// Replace the source/input ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: usize) -> Self {
        self.max_message_bytes = maximum;
        self
    }

    /// Replace the candidate-output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace field/work ceilings.
    #[must_use]
    pub const fn with_resource_limits(mut self, fields: usize, work_bytes: usize) -> Self {
        self.max_fields = fields;
        self.max_work_bytes = work_bytes;
        self
    }

    /// Replace nesting ceiling.
    #[must_use]
    pub const fn with_recursion_limit(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self
    }

    /// Replace the logical allocation ceiling.
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

    /// Return the source/input ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Return the candidate-output ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            // The handwritten scanner owns preservation of unknown fields;
            // Buffa must not materialize any of them.
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Resource axis rejected by strict decoding or prepared execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input source bytes.
    Bytes { observed: usize, maximum: usize },
    /// Strict wire-field visits.
    Fields { observed: usize, maximum: usize },
    /// Strict and projection work bytes.
    Work { observed: usize, maximum: usize },
    /// Protobuf nesting depth.
    Nesting { observed: u32, maximum: u32 },
    /// Logical staging allocations.
    Allocations { observed: usize, maximum: usize },
    /// Source and candidate bytes retained together.
    Retained { observed: usize, maximum: usize },
    /// Reserved parser/output scratch bytes.
    Scratch { observed: usize, maximum: usize },
    /// Candidate output bytes.
    Output { observed: usize, maximum: usize },
}

/// Strict properties codec failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Plain(&'static str),
    Wire(buffa::DecodeError),
    Limit(DecodeLimit),
    Duplicate(&'static str),
    WrongWire(&'static str),
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    Projection,
    Allocation(usize),
}

impl DecodeError {
    const fn plain(message: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Plain(message),
        }
    }

    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
    }

    const fn duplicate(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Duplicate(field),
        }
    }

    const fn wrong_wire(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::WrongWire(field),
        }
    }

    const fn noncanonical(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(field),
        }
    }

    const fn invalid_utf8(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::InvalidUtf8(field),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation(amount),
        }
    }

    /// Return the structured resource failure, if any.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Alias used by format-package adapters.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<DecodeLimit> {
        self.resource_limit()
    }

    /// Return the selected duplicate field name, if any.
    #[must_use]
    pub const fn duplicate_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::Duplicate(field) => Some(field),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Plain(message) => formatter.write_str(message),
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "movie properties input has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "movie properties visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "movie properties requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "movie properties nesting {observed} exceeds maximum {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Allocations { observed, maximum }) => write!(
                formatter,
                "movie properties requires {observed} allocations; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Retained { observed, maximum }) => write!(
                formatter,
                "movie properties retains {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Scratch { observed, maximum }) => write!(
                formatter,
                "movie properties requires {observed} scratch bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Output { observed, maximum }) => write!(
                formatter,
                "movie properties output has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Duplicate(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::WrongWire(field) => write!(formatter, "wrong wire type for {field}"),
            DecodeErrorKind::NonCanonical(field) => {
                write!(
                    formatter,
                    "non-canonical protobuf representation for {field}"
                )
            },
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "invalid UTF-8 in {field}"),
            DecodeErrorKind::Projection => {
                formatter.write_str("movie properties strict scan disagrees with Buffa projection")
            },
            DecodeErrorKind::Allocation(amount) => {
                write!(formatter, "movie properties cannot allocate {amount} bytes")
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

/// Borrowed semantic properties from one MovieArchive payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MoviePropertiesSnapshot<'source> {
    hyperlink: Option<&'source str>,
    locked: Option<bool>,
    aspect_ratio_locked: Option<bool>,
    accessibility_description: Option<&'source str>,
}

impl<'source> MoviePropertiesSnapshot<'source> {
    /// Return the explicitly encoded hyperlink, preserving absence.
    #[must_use]
    pub const fn hyperlink(self) -> Option<&'source str> {
        self.hyperlink
    }

    /// Alias for the semantic package vocabulary.
    #[must_use]
    pub const fn hyperlink_url(self) -> Option<&'source str> {
        self.hyperlink
    }

    /// Return the explicitly encoded lock state, preserving absence.
    #[must_use]
    pub const fn locked(self) -> Option<bool> {
        self.locked
    }

    /// Return the explicitly encoded aspect-ratio lock state.
    #[must_use]
    pub const fn aspect_ratio_locked(self) -> Option<bool> {
        self.aspect_ratio_locked
    }

    /// Return the explicitly encoded accessibility description.
    #[must_use]
    pub const fn accessibility_description(self) -> Option<&'source str> {
        self.accessibility_description
    }
}

/// Alias emphasizing that this is the selected drawable edge.
pub type MovieDrawablePropertiesSnapshot<'source> = MoviePropertiesSnapshot<'source>;

/// One optional source-preserving property update.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PropertyUpdate<T> {
    /// Preserve the complete source field span and presence.
    Preserve,
    /// Replace or append one canonical field value.
    Set(T),
    /// Remove the field when present.
    Clear,
}

impl<T> PropertyUpdate<T> {
    /// Construct a preserve update.
    #[must_use]
    pub const fn preserve() -> Self {
        Self::Preserve
    }

    /// Construct a set update.
    #[must_use]
    pub const fn set(value: T) -> Self {
        Self::Set(value)
    }

    /// Construct a clear update.
    #[must_use]
    pub const fn clear() -> Self {
        Self::Clear
    }
}

/// Requested replacement for the four selected drawable properties.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MoviePropertiesWrite<'value> {
    hyperlink: PropertyUpdate<&'value str>,
    locked: PropertyUpdate<bool>,
    aspect_ratio_locked: PropertyUpdate<bool>,
    accessibility_description: PropertyUpdate<&'value str>,
}

impl<'value> MoviePropertiesWrite<'value> {
    /// Construct an exact replacement. `None` clears a field, while
    /// `Some(false)` preserves an explicit false rather than collapsing it
    /// into absence.
    #[must_use]
    pub fn new(
        hyperlink: Option<&'value str>,
        locked: Option<bool>,
        aspect_ratio_locked: Option<bool>,
        accessibility_description: Option<&'value str>,
    ) -> Self {
        Self {
            hyperlink: option_update(hyperlink),
            locked: option_update(locked),
            aspect_ratio_locked: option_update(aspect_ratio_locked),
            accessibility_description: option_update(accessibility_description),
        }
    }

    /// Alias for [`Self::new`] using the semantic package vocabulary.
    #[must_use]
    pub fn from_values(
        hyperlink_url: Option<&'value str>,
        locked: Option<bool>,
        aspect_ratio_locked: Option<bool>,
        accessibility_description: Option<&'value str>,
    ) -> Self {
        Self::new(
            hyperlink_url,
            locked,
            aspect_ratio_locked,
            accessibility_description,
        )
    }

    /// Construct a write that preserves every selected field.
    #[must_use]
    pub const fn preserve() -> Self {
        Self {
            hyperlink: PropertyUpdate::Preserve,
            locked: PropertyUpdate::Preserve,
            aspect_ratio_locked: PropertyUpdate::Preserve,
            accessibility_description: PropertyUpdate::Preserve,
        }
    }

    /// Replace the hyperlink, or clear it when `None` is supplied.
    #[must_use]
    pub fn with_hyperlink(mut self, value: Option<&'value str>) -> Self {
        self.hyperlink = option_update(value);
        self
    }

    /// Alias using the format-package property name.
    #[must_use]
    pub fn with_hyperlink_url(self, value: Option<&'value str>) -> Self {
        self.with_hyperlink(value)
    }

    /// Replace the lock state, or clear it when `None` is supplied.
    #[must_use]
    pub fn with_locked(mut self, value: Option<bool>) -> Self {
        self.locked = option_update(value);
        self
    }

    /// Replace the aspect-ratio lock state, or clear it when `None` is supplied.
    #[must_use]
    pub fn with_aspect_ratio_locked(mut self, value: Option<bool>) -> Self {
        self.aspect_ratio_locked = option_update(value);
        self
    }

    /// Replace the accessibility description, or clear it when `None` is supplied.
    #[must_use]
    pub fn with_accessibility_description(mut self, value: Option<&'value str>) -> Self {
        self.accessibility_description = option_update(value);
        self
    }

    /// Set an explicit hyperlink.
    #[must_use]
    pub const fn set_hyperlink(mut self, value: &'value str) -> Self {
        self.hyperlink = PropertyUpdate::Set(value);
        self
    }

    /// Clear the hyperlink.
    #[must_use]
    pub const fn clear_hyperlink(mut self) -> Self {
        self.hyperlink = PropertyUpdate::Clear;
        self
    }

    /// Set an explicit lock state.
    #[must_use]
    pub const fn set_locked(mut self, value: bool) -> Self {
        self.locked = PropertyUpdate::Set(value);
        self
    }

    /// Clear the lock state.
    #[must_use]
    pub const fn clear_locked(mut self) -> Self {
        self.locked = PropertyUpdate::Clear;
        self
    }

    /// Set an explicit aspect-ratio lock state.
    #[must_use]
    pub const fn set_aspect_ratio_locked(mut self, value: bool) -> Self {
        self.aspect_ratio_locked = PropertyUpdate::Set(value);
        self
    }

    /// Clear the aspect-ratio lock state.
    #[must_use]
    pub const fn clear_aspect_ratio_locked(mut self) -> Self {
        self.aspect_ratio_locked = PropertyUpdate::Clear;
        self
    }

    /// Set an explicit accessibility description.
    #[must_use]
    pub const fn set_accessibility_description(mut self, value: &'value str) -> Self {
        self.accessibility_description = PropertyUpdate::Set(value);
        self
    }

    /// Clear the accessibility description.
    #[must_use]
    pub const fn clear_accessibility_description(mut self) -> Self {
        self.accessibility_description = PropertyUpdate::Clear;
        self
    }

    /// Set all four updates explicitly, retaining per-field presence control.
    #[must_use]
    pub const fn with_updates(
        hyperlink: PropertyUpdate<&'value str>,
        locked: PropertyUpdate<bool>,
        aspect_ratio_locked: PropertyUpdate<bool>,
        accessibility_description: PropertyUpdate<&'value str>,
    ) -> Self {
        Self {
            hyperlink,
            locked,
            aspect_ratio_locked,
            accessibility_description,
        }
    }
}

/// Alias emphasizing that this is the selected drawable edge.
pub type MovieDrawablePropertiesWrite<'value> = MoviePropertiesWrite<'value>;

const fn option_update<T: Copy>(value: Option<T>) -> PropertyUpdate<T> {
    match value {
        Some(value) => PropertyUpdate::Set(value),
        None => PropertyUpdate::Clear,
    }
}

/// Exact source accounting for a successful decode or prepare pass.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl DecodeReport {
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
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

/// Exact prepared execution requirements.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    /// Candidate output bytes.
    pub output_bytes: usize,
    /// Aggregate strict scan visits.
    pub fields: usize,
    /// Aggregate work bytes.
    pub work_bytes: usize,
    /// Maximum nesting depth.
    pub max_depth: u32,
    /// Logical allocations.
    pub allocations: usize,
    /// Source and candidate retained bytes.
    pub retained_bytes: usize,
    /// Parser/output scratch bytes.
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
    /// Convert requirements into matching execution ceilings.
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

    /// Alias used by package adapters.
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Caller-supplied ceilings replayed before candidate allocation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    /// Candidate output ceiling.
    pub output_bytes: usize,
    /// Field-visit ceiling.
    pub fields: usize,
    /// Work-byte ceiling.
    pub work_bytes: usize,
    /// Nesting ceiling.
    pub max_depth: u32,
    /// Allocation ceiling.
    pub allocations: usize,
    /// Retained-byte ceiling.
    pub retained_bytes: usize,
    /// Scratch-byte ceiling.
    pub scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact ceilings from one prepared requirement set.
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

/// Exact result accounting for one rewrite.
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

/// Owned candidate output and exact accounting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RewriteOutput {
    output: Vec<u8>,
    report: RewriteReport,
}

impl RewriteOutput {
    /// Borrow the rewritten bytes.
    #[must_use]
    pub fn output(&self) -> &[u8] {
        &self.output
    }

    /// Consume the result and return rewritten bytes.
    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.output
    }

    /// Alias returning the owned candidate bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.output
    }

    /// Borrow the candidate bytes under the common codec naming convention.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.output
    }

    /// Return exact rewrite accounting.
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// A source-bound rewrite prepared without allocating candidate output.
#[derive(Debug, Clone, Copy)]
pub struct PreparedMoviePropertiesRewrite<'source, 'value> {
    source: &'source [u8],
    update: NormalizedWrite<'value>,
    options: DecodeOptions,
    source_report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
}

impl PreparedMoviePropertiesRewrite<'_, '_> {
    /// Return source and preparation accounting.
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }

    /// Return exact requirements for candidate execution.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Allocate, emit, and candidate-readback the source-preserving rewrite.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_execution_limits(self.requirements, limits)?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| DecodeError::allocation(self.requirements.output_bytes))?;
        emit_rewrite(
            self.source,
            self.update,
            self.requirements.output_bytes,
            self.options,
            &mut output,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::projection());
        }

        let mut changed = false;
        for (left, right) in self.source.iter().zip(&output) {
            if left != right {
                changed = true;
                break;
            }
        }
        changed |= self.source.len() != output.len();

        let candidate_options = DecodeOptions {
            max_message_bytes: output.len().max(1),
            max_output_bytes: self.options.max_output_bytes.max(output.len()),
            max_fields: self.candidate_fields,
            max_work_bytes: self.candidate_work,
            max_allocations: self.requirements.allocations,
            max_retained_bytes: self.requirements.retained_bytes,
            max_scratch_bytes: self.requirements.scratch_bytes,
            ..self.options
        };
        let (candidate, report) = decode_movie_properties_with_report(&output, candidate_options)?;
        if report.fields != self.candidate_fields
            || report.work_bytes != self.candidate_work
            || report.max_depth != self.candidate_depth
            || !matches_update(candidate, self.update)
        {
            return Err(DecodeError::projection());
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
                changed,
            },
        })
    }
}

/// Decode one complete MovieArchive payload.
pub fn decode_movie_properties<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<MoviePropertiesSnapshot<'source>, DecodeError> {
    Ok(decode_movie_properties_with_report(source, options)?.0)
}

/// Decode one payload and return exact bounded accounting.
pub fn decode_movie_properties_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(MoviePropertiesSnapshot<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let scan = scan_source(source, options)?;
    let snapshot = scan.snapshot;
    let report = finish_report(source, scan.accounting, scan.max_depth, options)?;
    check_report_limits(report, options)?;
    Ok((snapshot, report))
}

/// Prepare a source-preserving rewrite without allocating candidate output.
pub fn prepare_movie_properties_rewrite<'source, 'value>(
    source: &'source [u8],
    write: MoviePropertiesWrite<'value>,
    options: DecodeOptions,
) -> Result<PreparedMoviePropertiesRewrite<'source, 'value>, DecodeError> {
    validate_input(source, options)?;
    let scan = scan_source(source, options)?;
    let source_report = finish_report(source, scan.accounting, scan.max_depth, options)?;
    let normalized = normalize_write(scan.snapshot, write);
    let measurement_options = residual_options(options, source_report)?;
    let measure = measure_rewrite(
        source,
        &scan.drawable,
        scan.super_field,
        normalized,
        scan.accounting,
        scan.max_depth,
        measurement_options,
    )?;
    let prepare_report = combine_prepare_report(source_report, measure, options)?;
    let staging_scratch = measure.output_bytes.checked_mul(2).ok_or_else(|| {
        DecodeError::limit(DecodeLimit::Scratch {
            observed: usize::MAX,
            maximum: options.max_scratch_bytes,
        })
    })?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: measure.output_bytes,
        fields: scan
            .accounting
            .fields
            .checked_add(measure.candidate_fields)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                })
            })?,
        work_bytes: scan
            .accounting
            .work_bytes
            .checked_add(measure.candidate_work)
            .and_then(|value| value.checked_add(measure.output_bytes))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?,
        max_depth: scan.max_depth,
        allocations: scan
            .accounting
            .allocations
            .checked_mul(2)
            .and_then(|value| value.checked_add(BUFFA_LOGICAL_ALLOCATIONS))
            .and_then(|value| value.checked_add(EXECUTE_ALLOCATIONS))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Allocations {
                    observed: usize::MAX,
                    maximum: options.max_allocations,
                })
            })?,
        retained_bytes: source
            .len()
            .checked_add(measure.output_bytes)
            .and_then(|value| value.checked_add(measure.new_drawable_bytes))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Retained {
                    observed: usize::MAX,
                    maximum: options.max_retained_bytes,
                })
            })?,
        scratch_bytes: scan
            .accounting
            .scratch_bytes
            .checked_add(staging_scratch)
            .and_then(|value| value.checked_add(measure.output_scratch))
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Scratch {
                    observed: usize::MAX,
                    maximum: options.max_scratch_bytes,
                })
            })?,
    };
    validate_candidate_output_size(requirements.output_bytes, options)?;
    check_requirements(requirements, options)?;
    check_combined_requirements(prepare_report, requirements, options)?;
    Ok(PreparedMoviePropertiesRewrite {
        source,
        update: normalized,
        options,
        source_report: prepare_report,
        requirements,
        candidate_fields: measure.candidate_fields,
        candidate_work: measure.candidate_work,
        candidate_depth: scan.max_depth.max(measure.measurement_depth),
    })
}

/// Rewrite selected MovieArchive drawable properties.
pub fn rewrite_movie_properties(
    source: &[u8],
    write: MoviePropertiesWrite<'_>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_movie_properties_rewrite(source, write, options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_output())
}

/// Rewrite selected properties and return exact accounting.
pub fn rewrite_movie_properties_with_report(
    source: &[u8],
    write: MoviePropertiesWrite<'_>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    let prepared = prepare_movie_properties_rewrite(source, write, options)?;
    let output = prepared.execute(prepared.execution_requirements().exact())?;
    let report = output.report;
    Ok((output.into_output(), report))
}

#[derive(Debug, Clone, Copy)]
struct NormalizedWrite<'value> {
    hyperlink: PropertyUpdate<&'value str>,
    locked: PropertyUpdate<bool>,
    aspect_ratio_locked: PropertyUpdate<bool>,
    accessibility_description: PropertyUpdate<&'value str>,
}

fn normalize_write<'value>(
    current: MoviePropertiesSnapshot<'_>,
    write: MoviePropertiesWrite<'value>,
) -> NormalizedWrite<'value> {
    NormalizedWrite {
        hyperlink: normalize_string(current.hyperlink, write.hyperlink),
        locked: normalize_bool(current.locked, write.locked),
        aspect_ratio_locked: normalize_bool(current.aspect_ratio_locked, write.aspect_ratio_locked),
        accessibility_description: normalize_string(
            current.accessibility_description,
            write.accessibility_description,
        ),
    }
}

fn normalize_string<'value>(
    current: Option<&str>,
    update: PropertyUpdate<&'value str>,
) -> PropertyUpdate<&'value str> {
    match update {
        PropertyUpdate::Set(value) if current == Some(value) => PropertyUpdate::Preserve,
        PropertyUpdate::Clear if current.is_none() => PropertyUpdate::Preserve,
        other => other,
    }
}

fn normalize_bool(current: Option<bool>, update: PropertyUpdate<bool>) -> PropertyUpdate<bool> {
    match update {
        PropertyUpdate::Set(value) if current == Some(value) => PropertyUpdate::Preserve,
        PropertyUpdate::Clear if current.is_none() => PropertyUpdate::Preserve,
        other => other,
    }
}

fn matches_update(current: MoviePropertiesSnapshot<'_>, update: NormalizedWrite<'_>) -> bool {
    matches_string(current.hyperlink, update.hyperlink)
        && matches_bool(current.locked, update.locked)
        && matches_bool(current.aspect_ratio_locked, update.aspect_ratio_locked)
        && matches_string(
            current.accessibility_description,
            update.accessibility_description,
        )
}

fn matches_string(current: Option<&str>, update: PropertyUpdate<&str>) -> bool {
    match update {
        PropertyUpdate::Preserve => true,
        PropertyUpdate::Set(value) => current == Some(value),
        PropertyUpdate::Clear => current.is_none(),
    }
}

fn matches_bool(current: Option<bool>, update: PropertyUpdate<bool>) -> bool {
    match update {
        PropertyUpdate::Preserve => true,
        PropertyUpdate::Set(value) => current == Some(value),
        PropertyUpdate::Clear => current.is_none(),
    }
}

#[derive(Debug)]
struct Scan<'source> {
    snapshot: MoviePropertiesSnapshot<'source>,
    drawable: Vec<ParsedField>,
    super_field: ParsedField,
    accounting: Accounting,
    max_depth: u32,
}

#[derive(Debug, Clone, Copy)]
struct Accounting {
    fields: usize,
    work_bytes: usize,
    allocations: usize,
    scratch_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    allocations: usize,
    scratch_bytes: usize,
    max_depth: u32,
}

impl Budget {
    const fn new() -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            allocations: 0,
            scratch_bytes: 0,
            max_depth: 0,
        }
    }

    fn visit(
        &mut self,
        fields: usize,
        work_bytes: usize,
        options: DecodeOptions,
    ) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(fields).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
        if self.fields > options.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed: self.fields,
                maximum: options.max_fields,
            }));
        }
        self.work_bytes = self.work_bytes.checked_add(work_bytes).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
        if self.work_bytes > options.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: options.max_work_bytes,
            }));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct ParsedField {
    number: u32,
    wire: u8,
    start: usize,
    key_end: usize,
    value_start: usize,
    value_end: usize,
    end: usize,
    value: Option<u64>,
}

fn scan_source<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<Scan<'source>, DecodeError> {
    let mut budget = Budget::new();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "TSD.MovieArchive.super")?;
    require_known_framing(source, super_field, 2, "TSD.MovieArchive.super")?;
    let drawable_source = &source[super_field.value_start..super_field.value_end];
    let drawable = parse_message(drawable_source, options, &mut budget, 2)?;
    let hyperlink = parse_string_field(
        drawable_source,
        &drawable,
        DRAWABLE_HYPERLINK_FIELD,
        "TSD.DrawableArchive.hyperlink",
    )?;
    let locked = parse_bool_field(
        drawable_source,
        &drawable,
        DRAWABLE_LOCKED_FIELD,
        "TSD.DrawableArchive.locked",
    )?;
    let aspect_ratio_locked = parse_bool_field(
        drawable_source,
        &drawable,
        DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
        "TSD.DrawableArchive.aspect_ratio_locked",
    )?;
    let accessibility_description = parse_string_field(
        drawable_source,
        &drawable,
        DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
        "TSD.DrawableArchive.accessibility_description",
    )?;
    let strict = MoviePropertiesSnapshot {
        hyperlink,
        locked,
        aspect_ratio_locked,
        accessibility_description,
    };
    let projected = force_buffa(source, options)?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    let accounting = Accounting {
        fields: budget.fields,
        work_bytes: budget.work_bytes,
        allocations: budget.allocations,
        scratch_bytes: budget.scratch_bytes,
    };
    Ok(Scan {
        snapshot: strict,
        drawable,
        super_field,
        accounting,
        max_depth: budget.max_depth,
    })
}

fn finish_report(
    source: &[u8],
    accounting: Accounting,
    max_depth: u32,
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    let allocations = accounting
        .allocations
        .checked_add(BUFFA_LOGICAL_ALLOCATIONS)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    let work_bytes = accounting
        .work_bytes
        .checked_add(source.len())
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let scratch_bytes = accounting
        .scratch_bytes
        .checked_add(source.len())
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let report = DecodeReport {
        input_bytes: source.len(),
        fields: accounting.fields,
        work_bytes,
        max_depth,
        allocations,
        retained_bytes: source.len(),
        scratch_bytes,
    };
    check_report_limits(report, options)?;
    Ok(report)
}

fn check_report_limits(report: DecodeReport, options: DecodeOptions) -> Result<(), DecodeError> {
    for (observed, maximum, limit) in [
        (
            report.fields,
            options.max_fields,
            DecodeLimit::Fields {
                observed: report.fields,
                maximum: options.max_fields,
            },
        ),
        (
            report.work_bytes,
            options.max_work_bytes,
            DecodeLimit::Work {
                observed: report.work_bytes,
                maximum: options.max_work_bytes,
            },
        ),
        (
            report.allocations,
            options.max_allocations,
            DecodeLimit::Allocations {
                observed: report.allocations,
                maximum: options.max_allocations,
            },
        ),
        (
            report.retained_bytes,
            options.max_retained_bytes,
            DecodeLimit::Retained {
                observed: report.retained_bytes,
                maximum: options.max_retained_bytes,
            },
        ),
        (
            report.scratch_bytes,
            options.max_scratch_bytes,
            DecodeLimit::Scratch {
                observed: report.scratch_bytes,
                maximum: options.max_scratch_bytes,
            },
        ),
    ] {
        if observed > maximum {
            return Err(DecodeError::limit(limit));
        }
    }
    if report.max_depth > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: report.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    Ok(())
}

fn check_requirements(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if requirements.output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::Output {
            observed: requirements.output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    for (observed, maximum, limit) in [
        (
            requirements.fields,
            options.max_fields,
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: options.max_fields,
            },
        ),
        (
            requirements.work_bytes,
            options.max_work_bytes,
            DecodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: options.max_work_bytes,
            },
        ),
        (
            requirements.allocations,
            options.max_allocations,
            DecodeLimit::Allocations {
                observed: requirements.allocations,
                maximum: options.max_allocations,
            },
        ),
        (
            requirements.retained_bytes,
            options.max_retained_bytes,
            DecodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: options.max_retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            options.max_scratch_bytes,
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: options.max_scratch_bytes,
            },
        ),
    ] {
        if observed > maximum {
            return Err(DecodeError::limit(limit));
        }
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    Ok(())
}

fn check_execution_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::limit(DecodeLimit::Output {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    for (observed, maximum, limit) in [
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
            return Err(DecodeError::limit(limit));
        }
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    Ok(())
}

fn residual_options(
    options: DecodeOptions,
    consumed: DecodeReport,
) -> Result<DecodeOptions, DecodeError> {
    let fields = options
        .max_fields
        .checked_sub(consumed.fields)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: options.max_fields.saturating_add(1),
                maximum: options.max_fields,
            })
        })?;
    let work_bytes = options
        .max_work_bytes
        .checked_sub(consumed.work_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: options.max_work_bytes.saturating_add(1),
                maximum: options.max_work_bytes,
            })
        })?;
    let allocations = options
        .max_allocations
        .checked_sub(consumed.allocations)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: options.max_allocations.saturating_add(1),
                maximum: options.max_allocations,
            })
        })?;
    let scratch_bytes = options
        .max_scratch_bytes
        .checked_sub(consumed.scratch_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: options.max_scratch_bytes.saturating_add(1),
                maximum: options.max_scratch_bytes,
            })
        })?;
    Ok(options
        .with_resource_limits(fields, work_bytes)
        .with_max_allocations(allocations)
        .with_max_scratch_bytes(scratch_bytes))
}

fn combine_prepare_report(
    source_report: DecodeReport,
    measure: RewriteMeasure,
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    let report = DecodeReport {
        input_bytes: source_report.input_bytes,
        fields: source_report
            .fields
            .checked_add(measure.measurement.fields)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                })
            })?,
        work_bytes: source_report
            .work_bytes
            .checked_add(measure.measurement.work_bytes)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?,
        max_depth: source_report.max_depth.max(measure.measurement_depth),
        allocations: source_report
            .allocations
            .checked_add(measure.measurement.allocations)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Allocations {
                    observed: usize::MAX,
                    maximum: options.max_allocations,
                })
            })?,
        retained_bytes: source_report.retained_bytes,
        scratch_bytes: source_report
            .scratch_bytes
            .checked_add(measure.measurement.scratch_bytes)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Scratch {
                    observed: usize::MAX,
                    maximum: options.max_scratch_bytes,
                })
            })?,
    };
    check_report_limits(report, options)?;
    Ok(report)
}

fn check_combined_requirements(
    prepare: DecodeReport,
    execution: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let fields = prepare
        .fields
        .checked_add(execution.fields)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    let work_bytes = prepare
        .work_bytes
        .checked_add(execution.work_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let allocations = prepare
        .allocations
        .checked_add(execution.allocations)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    let scratch_bytes = prepare
        .scratch_bytes
        .checked_add(execution.scratch_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let retained_bytes = prepare
        .retained_bytes
        .checked_add(execution.retained_bytes)
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    for (observed, maximum, limit) in [
        (
            fields,
            options.max_fields,
            DecodeLimit::Fields {
                observed: fields,
                maximum: options.max_fields,
            },
        ),
        (
            work_bytes,
            options.max_work_bytes,
            DecodeLimit::Work {
                observed: work_bytes,
                maximum: options.max_work_bytes,
            },
        ),
        (
            allocations,
            options.max_allocations,
            DecodeLimit::Allocations {
                observed: allocations,
                maximum: options.max_allocations,
            },
        ),
        (
            scratch_bytes,
            options.max_scratch_bytes,
            DecodeLimit::Scratch {
                observed: scratch_bytes,
                maximum: options.max_scratch_bytes,
            },
        ),
        (
            retained_bytes,
            options.max_retained_bytes,
            DecodeLimit::Retained {
                observed: retained_bytes,
                maximum: options.max_retained_bytes,
            },
        ),
    ] {
        if observed > maximum {
            return Err(DecodeError::limit(limit));
        }
    }
    if prepare.max_depth.max(execution.max_depth) > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: prepare.max_depth.max(execution.max_depth),
            maximum: options.recursion_limit,
        }));
    }
    Ok(())
}

fn validate_candidate_output_size(
    output_bytes: usize,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES).unwrap_or(usize::MAX);
    if output_bytes > hard_maximum {
        return Err(DecodeError::limit(DecodeLimit::Output {
            observed: output_bytes,
            maximum: hard_maximum,
        }));
    }
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::limit(DecodeLimit::Output {
            observed: output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    Ok(())
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES).unwrap_or(usize::MAX);
    if options.max_message_bytes > hard_maximum || source.len() > options.max_message_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: source.len().max(options.max_message_bytes),
            maximum: hard_maximum.min(options.max_message_bytes),
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}

fn force_buffa<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<MoviePropertiesSnapshot<'source>, DecodeError> {
    let view: projection::MovieArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let drawable = view
        .super_
        .get()
        .map_err(DecodeError::from)?
        .ok_or_else(|| DecodeError::plain("Buffa movie properties projection missing super"))?;
    Ok(MoviePropertiesSnapshot {
        hyperlink: drawable.hyperlink,
        locked: drawable.locked,
        aspect_ratio_locked: drawable.aspect_ratio_locked,
        accessibility_description: drawable.accessibility_description,
    })
}

fn unique_known(
    fields: &[ParsedField],
    number: u32,
    name: &'static str,
) -> Result<ParsedField, DecodeError> {
    let mut found = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number == number)
    {
        if found.is_some() {
            return Err(DecodeError::duplicate(name));
        }
        found = Some(field);
    }
    found.ok_or_else(|| DecodeError::plain(name))
}

fn optional_known(
    fields: &[ParsedField],
    number: u32,
    name: &'static str,
) -> Result<Option<ParsedField>, DecodeError> {
    let mut found = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number == number)
    {
        if found.is_some() {
            return Err(DecodeError::duplicate(name));
        }
        found = Some(field);
    }
    Ok(found)
}

fn require_known_framing(
    source: &[u8],
    field: ParsedField,
    wire: u8,
    name: &'static str,
) -> Result<(), DecodeError> {
    if field.wire != wire {
        return Err(DecodeError::wrong_wire(name));
    }
    if field.key_end - field.start != varint_len(u64::from(field.number << 3 | u32::from(wire))) {
        return Err(DecodeError::noncanonical(name));
    }
    if wire == 2 {
        let (length, length_end) = read_varint(source, field.key_end, field.value_start)?;
        if length_end - field.key_end != varint_len(length) {
            return Err(DecodeError::noncanonical(name));
        }
    }
    Ok(())
}

fn parse_string_field<'source>(
    drawable_source: &'source [u8],
    fields: &[ParsedField],
    number: u32,
    name: &'static str,
) -> Result<Option<&'source str>, DecodeError> {
    let Some(field) = optional_known(fields, number, name)? else {
        return Ok(None);
    };
    require_known_framing(drawable_source, field, 2, name)?;
    std::str::from_utf8(&drawable_source[field.value_start..field.value_end])
        .map(Some)
        .map_err(|_| DecodeError::invalid_utf8(name))
}

fn parse_bool_field(
    drawable_source: &[u8],
    fields: &[ParsedField],
    number: u32,
    name: &'static str,
) -> Result<Option<bool>, DecodeError> {
    let Some(field) = optional_known(fields, number, name)? else {
        return Ok(None);
    };
    require_known_framing(drawable_source, field, 0, name)?;
    let value = field
        .value
        .ok_or_else(|| DecodeError::plain("missing protobuf bool value"))?;
    if value > 1 || field.value_end - field.value_start != 1 {
        return Err(DecodeError::noncanonical(name));
    }
    Ok(Some(value == 1))
}

fn parse_message(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Vec<ParsedField>, DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: depth,
            maximum: options.recursion_limit,
        }));
    }
    budget.max_depth = budget.max_depth.max(depth);
    let capacity = source.len().min(options.max_fields);
    let scratch = capacity
        .checked_mul(size_of::<ParsedField>())
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    budget.scratch_bytes = budget.scratch_bytes.checked_add(scratch).ok_or_else(|| {
        DecodeError::limit(DecodeLimit::Scratch {
            observed: usize::MAX,
            maximum: options.max_scratch_bytes,
        })
    })?;
    if budget.scratch_bytes > options.max_scratch_bytes {
        return Err(DecodeError::limit(DecodeLimit::Scratch {
            observed: budget.scratch_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    if capacity != 0 {
        budget.allocations = budget.allocations.checked_add(1).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
        if budget.allocations > options.max_allocations {
            return Err(DecodeError::limit(DecodeLimit::Allocations {
                observed: budget.allocations,
                maximum: options.max_allocations,
            }));
        }
    }
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(capacity)
        .map_err(|_| DecodeError::allocation(scratch))?;
    let mut offset = 0;
    while offset < source.len() {
        let (field, next) =
            parse_field(source, offset, source.len(), options, budget, depth, true)?;
        fields.push(field);
        offset = next;
    }
    Ok(fields)
}

fn parse_field(
    source: &[u8],
    offset: usize,
    end: usize,
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
    owned_path: bool,
) -> Result<(ParsedField, usize), DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::limit(DecodeLimit::Nesting {
            observed: depth,
            maximum: options.recursion_limit,
        }));
    }
    budget.max_depth = budget.max_depth.max(depth);
    let (key, key_end) = read_varint(source, offset, end)?;
    let number_value = key >> 3;
    let wire = (key & 7) as u8;
    if number_value == 0 || number_value > 0x1fff_ffff {
        return Err(DecodeError::plain("invalid protobuf field number"));
    }
    let number = number_value as u32;
    if owned_path
        && is_known_at_depth(number, depth)
        && key_end - offset != varint_len(u64::from(number << 3 | u32::from(wire)))
    {
        return Err(DecodeError::noncanonical("known field key"));
    }
    let (value_start, value_end, next) = match wire {
        0 => {
            let (_, value_end) = read_varint(source, key_end, end)?;
            (key_end, value_end, value_end)
        },
        1 => {
            let value_end = key_end
                .checked_add(8)
                .ok_or_else(|| DecodeError::plain("fixed64 overflow"))?;
            if value_end > end {
                return Err(DecodeError::plain("truncated fixed64"));
            }
            (key_end, value_end, value_end)
        },
        2 => {
            let (length, payload_start) = read_varint(source, key_end, end)?;
            if owned_path
                && is_known_at_depth(number, depth)
                && payload_start - key_end != varint_len(length)
            {
                return Err(DecodeError::noncanonical("known length"));
            }
            let length =
                usize::try_from(length).map_err(|_| DecodeError::plain("length overflow"))?;
            let value_end = payload_start
                .checked_add(length)
                .ok_or_else(|| DecodeError::plain("length overflow"))?;
            if value_end > end {
                return Err(DecodeError::plain("truncated length-delimited field"));
            }
            (payload_start, value_end, value_end)
        },
        3 => {
            if depth >= options.recursion_limit {
                return Err(DecodeError::limit(DecodeLimit::Nesting {
                    observed: depth + 1,
                    maximum: options.recursion_limit,
                }));
            }
            let mut cursor = key_end;
            loop {
                if cursor >= end {
                    return Err(DecodeError::plain("unterminated protobuf group"));
                }
                let (child_key, child_key_end) = read_varint(source, cursor, end)?;
                let child_wire = (child_key & 7) as u8;
                let child_number_value = child_key >> 3;
                if child_number_value == 0 || child_number_value > 0x1fff_ffff {
                    return Err(DecodeError::plain("invalid protobuf group field number"));
                }
                let child_number = child_number_value as u32;
                if child_wire == 4 {
                    if child_number != number {
                        return Err(DecodeError::plain("mismatched protobuf group"));
                    }
                    budget.visit(1, child_key_end - cursor, options)?;
                    cursor = child_key_end;
                    break;
                }
                let (_, child_next) =
                    parse_field(source, cursor, end, options, budget, depth + 1, false)?;
                cursor = child_next;
            }
            (key_end, cursor, cursor)
        },
        4 => return Err(DecodeError::plain("unexpected protobuf end group")),
        5 => {
            let value_end = key_end
                .checked_add(4)
                .ok_or_else(|| DecodeError::plain("fixed32 overflow"))?;
            if value_end > end {
                return Err(DecodeError::plain("truncated fixed32"));
            }
            (key_end, value_end, value_end)
        },
        _ => return Err(DecodeError::plain("invalid protobuf wire type")),
    };
    budget.visit(1, next - offset, options)?;
    let value = if wire == 0 {
        Some(read_varint(source, value_start, value_end)?.0)
    } else {
        None
    };
    Ok((
        ParsedField {
            number,
            wire,
            start: offset,
            key_end,
            value_start,
            value_end,
            end: next,
            value,
        },
        next,
    ))
}

fn is_known_at_depth(number: u32, depth: u32) -> bool {
    match depth {
        1 => number == MOVIE_SUPER_FIELD,
        2 => matches!(
            number,
            DRAWABLE_HYPERLINK_FIELD
                | DRAWABLE_LOCKED_FIELD
                | DRAWABLE_ASPECT_RATIO_LOCKED_FIELD
                | DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD
        ),
        _ => false,
    }
}

fn read_varint(source: &[u8], mut offset: usize, end: usize) -> Result<(u64, usize), DecodeError> {
    let mut value = 0u64;
    for index in 0..10 {
        if offset >= end {
            return Err(DecodeError::plain("truncated protobuf varint"));
        }
        let byte = source[offset];
        offset += 1;
        if index == 9 && byte > 1 {
            return Err(DecodeError::plain("protobuf varint overflow"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, offset));
        }
    }
    Err(DecodeError::plain("protobuf varint too long"))
}

fn varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        ((64 - value.leading_zeros()).div_ceil(7)) as usize
    }
}

#[derive(Debug, Clone, Copy)]
struct RewriteMeasure {
    output_bytes: usize,
    candidate_fields: usize,
    candidate_work: usize,
    output_scratch: usize,
    measurement: Accounting,
    measurement_depth: u32,
    new_drawable_bytes: usize,
}

fn measure_rewrite(
    source: &[u8],
    drawable: &[ParsedField],
    super_field: ParsedField,
    update: NormalizedWrite<'_>,
    accounting: Accounting,
    depth: u32,
    options: DecodeOptions,
) -> Result<RewriteMeasure, DecodeError> {
    let mut measurement_budget = Budget::new();
    let measured_root = parse_message(source, options, &mut measurement_budget, 1)?;
    let measured_super = unique_known(&measured_root, MOVIE_SUPER_FIELD, "TSD.MovieArchive.super")?;
    require_known_framing(source, measured_super, 2, "TSD.MovieArchive.super")?;
    let measured_drawable_source = &source[measured_super.value_start..measured_super.value_end];
    let _measured_drawable = parse_message(
        measured_drawable_source,
        options,
        &mut measurement_budget,
        2,
    )?;
    let original_drawable_len = super_field.value_end - super_field.value_start;
    let replacement = replacement_lengths(drawable, update)?;
    let new_drawable_len = replace_selected_length(original_drawable_len, drawable, replacement)?;
    let original_super_span = super_field.end - super_field.start;
    let new_super_span = length_field_len(MOVIE_SUPER_FIELD, new_drawable_len)?;
    let output_bytes = source
        .len()
        .checked_sub(original_super_span)
        .and_then(|value| value.checked_add(new_super_span))
        .ok_or_else(|| DecodeError::plain("movie properties output length overflow"))?;
    let current_count = selected_count(drawable);
    let candidate_count = current_count
        .checked_add(replacement.added_fields)
        .and_then(|value| value.checked_sub(replacement.removed_fields))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    let candidate_fields = accounting
        .fields
        .checked_add(candidate_count)
        .and_then(|value| value.checked_sub(current_count))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    let drawable_delta = new_drawable_len as isize - original_drawable_len as isize;
    let super_delta = new_super_span as isize - original_super_span as isize;
    let candidate_work = adjust_signed(accounting.work_bytes, drawable_delta)
        .and_then(|value| adjust_signed(value, super_delta))
        .and_then(|value| value.checked_add(output_bytes))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let output_scratch = output_bytes
        .checked_mul(size_of::<ParsedField>())
        .and_then(|value| value.checked_mul(4))
        .ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let measurement = Accounting {
        fields: measurement_budget.fields,
        work_bytes: measurement_budget.work_bytes,
        allocations: measurement_budget.allocations,
        scratch_bytes: measurement_budget.scratch_bytes,
    };
    Ok(RewriteMeasure {
        output_bytes,
        candidate_fields,
        candidate_work,
        output_scratch,
        measurement,
        measurement_depth: measurement_budget.max_depth.max(depth),
        new_drawable_bytes: new_drawable_len,
    })
}

#[derive(Debug, Clone, Copy)]
struct ReplacementLengths {
    hyperlink: usize,
    locked: usize,
    aspect_ratio_locked: usize,
    accessibility_description: usize,
    added_fields: usize,
    removed_fields: usize,
}

fn replacement_lengths(
    fields: &[ParsedField],
    update: NormalizedWrite<'_>,
) -> Result<ReplacementLengths, DecodeError> {
    let hyperlink = replacement_string_len(fields, DRAWABLE_HYPERLINK_FIELD, update.hyperlink)?;
    let locked = replacement_bool_len(fields, DRAWABLE_LOCKED_FIELD, update.locked)?;
    let aspect_ratio_locked = replacement_bool_len(
        fields,
        DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
        update.aspect_ratio_locked,
    )?;
    let accessibility_description = replacement_string_len(
        fields,
        DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
        update.accessibility_description,
    )?;
    let mut added_fields = 0;
    let mut removed_fields = 0;
    for (field, length) in [
        (DRAWABLE_HYPERLINK_FIELD, hyperlink),
        (DRAWABLE_LOCKED_FIELD, locked),
        (DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, aspect_ratio_locked),
        (
            DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
            accessibility_description,
        ),
    ] {
        let present = fields.iter().any(|candidate| candidate.number == field);
        if !present && length != 0 {
            added_fields += 1;
        } else if present && length == 0 {
            removed_fields += 1;
        }
    }
    Ok(ReplacementLengths {
        hyperlink,
        locked,
        aspect_ratio_locked,
        accessibility_description,
        added_fields,
        removed_fields,
    })
}

fn replacement_string_len(
    fields: &[ParsedField],
    number: u32,
    update: PropertyUpdate<&str>,
) -> Result<usize, DecodeError> {
    let old = fields
        .iter()
        .find(|field| field.number == number)
        .map(|field| field.end - field.start);
    Ok(match update {
        PropertyUpdate::Preserve => old.unwrap_or(0),
        PropertyUpdate::Clear => 0,
        PropertyUpdate::Set(value) => length_field_len(number, value.len())?,
    })
}

fn replacement_bool_len(
    fields: &[ParsedField],
    number: u32,
    update: PropertyUpdate<bool>,
) -> Result<usize, DecodeError> {
    let old = fields
        .iter()
        .find(|field| field.number == number)
        .map(|field| field.end - field.start);
    Ok(match update {
        PropertyUpdate::Preserve => old.unwrap_or(0),
        PropertyUpdate::Clear => 0,
        PropertyUpdate::Set(_) => varint_len(u64::from(number << 3))
            .checked_add(1)
            .ok_or_else(|| DecodeError::plain("bool field length overflow"))?,
    })
}

fn selected_count(fields: &[ParsedField]) -> usize {
    fields
        .iter()
        .filter(|field| {
            matches!(
                field.number,
                DRAWABLE_HYPERLINK_FIELD
                    | DRAWABLE_LOCKED_FIELD
                    | DRAWABLE_ASPECT_RATIO_LOCKED_FIELD
                    | DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD
            )
        })
        .count()
}

fn replace_selected_length(
    original: usize,
    fields: &[ParsedField],
    replacement: ReplacementLengths,
) -> Result<usize, DecodeError> {
    let old = fields
        .iter()
        .filter(|field| {
            matches!(
                field.number,
                DRAWABLE_HYPERLINK_FIELD
                    | DRAWABLE_LOCKED_FIELD
                    | DRAWABLE_ASPECT_RATIO_LOCKED_FIELD
                    | DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD
            )
        })
        .map(|field| field.end - field.start)
        .try_fold(0usize, usize::checked_add)
        .ok_or_else(|| DecodeError::plain("movie properties drawable length overflow"))?;
    let replacement_total = fields
        .iter()
        .filter_map(|field| match field.number {
            DRAWABLE_HYPERLINK_FIELD => Some(replacement.hyperlink),
            DRAWABLE_LOCKED_FIELD => Some(replacement.locked),
            DRAWABLE_ASPECT_RATIO_LOCKED_FIELD => Some(replacement.aspect_ratio_locked),
            DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD => Some(replacement.accessibility_description),
            _ => None,
        })
        .try_fold(0usize, usize::checked_add)
        .ok_or_else(|| DecodeError::plain("movie properties drawable length overflow"))?
        .checked_add(
            if fields
                .iter()
                .any(|field| field.number == DRAWABLE_HYPERLINK_FIELD)
            {
                0
            } else {
                replacement.hyperlink
            },
        )
        .and_then(|value| {
            value.checked_add(
                if fields
                    .iter()
                    .any(|field| field.number == DRAWABLE_LOCKED_FIELD)
                {
                    0
                } else {
                    replacement.locked
                },
            )
        })
        .and_then(|value| {
            value.checked_add(
                if fields
                    .iter()
                    .any(|field| field.number == DRAWABLE_ASPECT_RATIO_LOCKED_FIELD)
                {
                    0
                } else {
                    replacement.aspect_ratio_locked
                },
            )
        })
        .and_then(|value| {
            value.checked_add(
                if fields
                    .iter()
                    .any(|field| field.number == DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD)
                {
                    0
                } else {
                    replacement.accessibility_description
                },
            )
        })
        .ok_or_else(|| DecodeError::plain("movie properties drawable length overflow"))?;
    original
        .checked_sub(old)
        .and_then(|value| value.checked_add(replacement_total))
        .ok_or_else(|| DecodeError::plain("movie properties drawable length overflow"))
}

fn adjust_signed(value: usize, delta: isize) -> Option<usize> {
    if delta >= 0 {
        value.checked_add(delta as usize)
    } else {
        value.checked_sub(delta.unsigned_abs())
    }
}

fn length_field_len(number: u32, payload: usize) -> Result<usize, DecodeError> {
    let key = number
        .checked_shl(3)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(|| DecodeError::plain("length-delimited field key overflow"))?;
    let payload_u64 = u64::try_from(payload)
        .map_err(|_| DecodeError::plain("length-delimited payload length overflow"))?;
    varint_len(u64::from(key))
        .checked_add(varint_len(payload_u64))
        .and_then(|value| value.checked_add(payload))
        .ok_or_else(|| DecodeError::plain("length-delimited field length overflow"))
}

fn emit_rewrite(
    source: &[u8],
    update: NormalizedWrite<'_>,
    expected: usize,
    options: DecodeOptions,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut budget = Budget::new();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "TSD.MovieArchive.super")?;
    require_known_framing(source, super_field, 2, "TSD.MovieArchive.super")?;
    let drawable_source = &source[super_field.value_start..super_field.value_end];
    let drawable = parse_message(drawable_source, options, &mut budget, 2)?;
    let replacement = replacement_lengths(&drawable, update)?;
    let new_drawable_len = replace_selected_length(drawable_source.len(), &drawable, replacement)?;
    let mut new_drawable = Vec::new();
    new_drawable
        .try_reserve_exact(new_drawable_len)
        .map_err(|_| DecodeError::allocation(new_drawable_len))?;
    emit_drawable(drawable_source, &drawable, update, &mut new_drawable)?;
    if new_drawable.len() != new_drawable_len {
        return Err(DecodeError::projection());
    }

    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::allocation(expected))?;
    for field in root.iter().copied() {
        if field.number == MOVIE_SUPER_FIELD {
            append_length_replacement(output, source, field, &new_drawable);
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    if output.len() != expected {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn emit_drawable(
    source: &[u8],
    fields: &[ParsedField],
    update: NormalizedWrite<'_>,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut emitted = [false; 4];
    for field in fields.iter().copied() {
        let index = selected_index(field.number);
        if let Some(index) = index {
            if emitted[index] {
                return Err(DecodeError::duplicate("selected drawable field"));
            }
            emitted[index] = true;
            let value = match field.number {
                DRAWABLE_HYPERLINK_FIELD => emit_update_string(
                    output,
                    source,
                    field,
                    update.hyperlink,
                    DRAWABLE_HYPERLINK_FIELD,
                )?,
                DRAWABLE_LOCKED_FIELD => {
                    emit_update_bool(output, source, field, update.locked, DRAWABLE_LOCKED_FIELD)?
                },
                DRAWABLE_ASPECT_RATIO_LOCKED_FIELD => emit_update_bool(
                    output,
                    source,
                    field,
                    update.aspect_ratio_locked,
                    DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
                )?,
                DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD => emit_update_string(
                    output,
                    source,
                    field,
                    update.accessibility_description,
                    DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
                )?,
                _ => false,
            };
            let _ = value;
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    if !emitted[0] {
        if let PropertyUpdate::Set(value) = update.hyperlink {
            append_string(output, DRAWABLE_HYPERLINK_FIELD, value)?;
        }
    }
    if !emitted[1] {
        if let PropertyUpdate::Set(value) = update.locked {
            append_bool(output, DRAWABLE_LOCKED_FIELD, value);
        }
    }
    if !emitted[2] {
        if let PropertyUpdate::Set(value) = update.aspect_ratio_locked {
            append_bool(output, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, value);
        }
    }
    if !emitted[3] {
        if let PropertyUpdate::Set(value) = update.accessibility_description {
            append_string(output, DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD, value)?;
        }
    }
    Ok(())
}

fn selected_index(number: u32) -> Option<usize> {
    match number {
        DRAWABLE_HYPERLINK_FIELD => Some(0),
        DRAWABLE_LOCKED_FIELD => Some(1),
        DRAWABLE_ASPECT_RATIO_LOCKED_FIELD => Some(2),
        DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD => Some(3),
        _ => None,
    }
}

fn emit_update_string(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    update: PropertyUpdate<&str>,
    number: u32,
) -> Result<bool, DecodeError> {
    match update {
        PropertyUpdate::Preserve => output.extend_from_slice(&source[field.start..field.end]),
        PropertyUpdate::Set(value) => append_string(output, number, value)?,
        PropertyUpdate::Clear => {},
    }
    Ok(true)
}

fn emit_update_bool(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    update: PropertyUpdate<bool>,
    number: u32,
) -> Result<bool, DecodeError> {
    match update {
        PropertyUpdate::Preserve => output.extend_from_slice(&source[field.start..field.end]),
        PropertyUpdate::Set(value) => append_bool(output, number, value),
        PropertyUpdate::Clear => {},
    }
    Ok(true)
}

fn append_string(output: &mut Vec<u8>, number: u32, value: &str) -> Result<(), DecodeError> {
    append_varint(output, u64::from(number << 3 | 2));
    append_varint(
        output,
        u64::try_from(value.len()).map_err(|_| DecodeError::plain("string length overflow"))?,
    );
    output.extend_from_slice(value.as_bytes());
    Ok(())
}

fn append_bool(output: &mut Vec<u8>, number: u32, value: bool) {
    append_varint(output, u64::from(number << 3));
    output.push(u8::from(value));
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn append_length_replacement(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    payload: &[u8],
) {
    output.extend_from_slice(&source[field.start..field.key_end]);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

#[cfg(test)]
mod tests {
    use super::*;

    fn field_bytes(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        append_varint(&mut output, u64::from(number << 3 | 2));
        append_varint(&mut output, payload.len() as u64);
        output.extend_from_slice(payload);
        output
    }

    fn field_bool(number: u32, value: bool) -> Vec<u8> {
        let mut output = Vec::new();
        append_bool(&mut output, number, value);
        output
    }

    fn source() -> Vec<u8> {
        let mut drawable = Vec::new();
        drawable.extend(field_bytes(
            DRAWABLE_HYPERLINK_FIELD,
            b"https://example.test",
        ));
        drawable.extend(field_bool(DRAWABLE_LOCKED_FIELD, false));
        drawable.extend(field_bytes(
            DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
            "Accessible ✨".as_bytes(),
        ));
        drawable.extend([0x7a, 0x03, 0x01, 0x02, 0x03]);
        let mut root = field_bytes(MOVIE_SUPER_FIELD, &drawable);
        root.extend([0x12, 0x01, 0x7f]);
        root
    }

    #[test]
    fn decode_preserves_presence_and_unknown_spans() {
        let source = source();
        let snapshot =
            decode_movie_properties(&source, DecodeOptions::for_source(&source)).unwrap();
        assert_eq!(snapshot.hyperlink(), Some("https://example.test"));
        assert_eq!(snapshot.locked(), Some(false));
        assert_eq!(snapshot.aspect_ratio_locked(), None);
        assert_eq!(snapshot.accessibility_description(), Some("Accessible ✨"));
    }

    #[test]
    fn rewrite_preserves_unknown_spans_and_supports_clear() {
        let source = source();
        let options = DecodeOptions::for_source(&source);
        let write = MoviePropertiesWrite::preserve()
            .set_hyperlink("https://changed.test")
            .set_locked(true)
            .clear_accessibility_description();
        let output = rewrite_movie_properties(&source, write, options).unwrap();
        assert!(
            output
                .windows(5)
                .any(|window| window == [0x7a, 0x03, 0x01, 0x02, 0x03])
        );
        assert!(output.ends_with(&[0x12, 0x01, 0x7f]));
        let snapshot =
            decode_movie_properties(&output, DecodeOptions::for_source(&output)).unwrap();
        assert_eq!(snapshot.hyperlink(), Some("https://changed.test"));
        assert_eq!(snapshot.locked(), Some(true));
        assert_eq!(snapshot.accessibility_description(), None);
    }

    #[test]
    fn explicit_false_is_distinct_from_absence() {
        let source = field_bytes(MOVIE_SUPER_FIELD, &[]);
        let options = DecodeOptions::for_source(&source);
        let output = rewrite_movie_properties(
            &source,
            MoviePropertiesWrite::preserve().set_locked(false),
            options,
        )
        .unwrap();
        assert_eq!(
            decode_movie_properties(&output, DecodeOptions::for_source(&output))
                .unwrap()
                .locked(),
            Some(false)
        );
    }

    #[test]
    fn malformed_owned_fields_are_rejected() {
        let duplicate = [
            field_bytes(MOVIE_SUPER_FIELD, &[]),
            field_bytes(MOVIE_SUPER_FIELD, &[]),
        ]
        .concat();
        assert!(
            decode_movie_properties(&duplicate, DecodeOptions::for_source(&duplicate)).is_err()
        );
        let wrong_wire = [0x08, 0x01, 0x01];
        assert!(
            decode_movie_properties(&wrong_wire, DecodeOptions::for_source(&wrong_wire)).is_err()
        );
        let duplicate_hyperlink = field_bytes(
            MOVIE_SUPER_FIELD,
            &[
                field_bytes(DRAWABLE_HYPERLINK_FIELD, b"a"),
                field_bytes(DRAWABLE_HYPERLINK_FIELD, b"b"),
            ]
            .concat(),
        );
        assert!(
            decode_movie_properties(
                &duplicate_hyperlink,
                DecodeOptions::for_source(&duplicate_hyperlink),
            )
            .is_err()
        );
        let wrong_hyperlink_wire = field_bytes(MOVIE_SUPER_FIELD, &[0x20, 0x01]);
        assert!(
            decode_movie_properties(
                &wrong_hyperlink_wire,
                DecodeOptions::for_source(&wrong_hyperlink_wire),
            )
            .is_err()
        );
        let overlong_key = field_bytes(MOVIE_SUPER_FIELD, &[0xa2, 0x80, 0x00, 0x01, b'x']);
        assert!(
            decode_movie_properties(&overlong_key, DecodeOptions::for_source(&overlong_key),)
                .is_err()
        );
        let overlong_length = field_bytes(MOVIE_SUPER_FIELD, &[0x22, 0x81, 0x00, b'x']);
        assert!(
            decode_movie_properties(
                &overlong_length,
                DecodeOptions::for_source(&overlong_length),
            )
            .is_err()
        );
        let invalid_utf8 = field_bytes(DRAWABLE_HYPERLINK_FIELD, &[0xff]);
        let source = field_bytes(MOVIE_SUPER_FIELD, &invalid_utf8);
        assert!(decode_movie_properties(&source, DecodeOptions::for_source(&source)).is_err());
        let noncanonical_bool = [0x0a, 0x04, 0x28, 0x80, 0x00, 0x00];
        assert!(
            decode_movie_properties(
                &noncanonical_bool,
                DecodeOptions::for_source(&noncanonical_bool)
            )
            .is_err()
        );
    }

    #[test]
    fn prepared_execution_enforces_every_reported_axis() {
        let source = source();
        let options = DecodeOptions::for_source(&source);
        let prepared = prepare_movie_properties_rewrite(
            &source,
            MoviePropertiesWrite::preserve().set_hyperlink("https://changed.example.test"),
            options,
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        let prepare_report = prepared.prepare_report();
        assert!(prepare_report.fields() > 0);
        assert!(requirements.fields > 0);
        assert!(prepare_report.work_bytes() > 0);
        assert!(requirements.work_bytes > 0);
        assert!(prepare_report.allocations() > 0);
        assert!(requirements.allocations > 0);
        assert!(requirements.retained_bytes >= source.len());
        assert!(prepare_report.scratch_bytes() > 0);
        assert!(requirements.scratch_bytes > 0);
        assert!(requirements.max_depth > 0);
        assert!(
            prepared
                .execute(requirements.exact())
                .unwrap()
                .report()
                .changed()
        );

        let checks = [
            (
                requirements.fields,
                requirements.exact().with_fields(requirements.fields - 1),
                DecodeLimit::Fields {
                    observed: requirements.fields,
                    maximum: requirements.fields - 1,
                },
            ),
            (
                requirements.work_bytes,
                requirements
                    .exact()
                    .with_work_bytes(requirements.work_bytes - 1),
                DecodeLimit::Work {
                    observed: requirements.work_bytes,
                    maximum: requirements.work_bytes - 1,
                },
            ),
            (
                requirements.output_bytes,
                requirements
                    .exact()
                    .with_output_bytes(requirements.output_bytes - 1),
                DecodeLimit::Output {
                    observed: requirements.output_bytes,
                    maximum: requirements.output_bytes - 1,
                },
            ),
            (
                requirements.allocations,
                requirements
                    .exact()
                    .with_allocations(requirements.allocations - 1),
                DecodeLimit::Allocations {
                    observed: requirements.allocations,
                    maximum: requirements.allocations - 1,
                },
            ),
            (
                requirements.retained_bytes,
                requirements
                    .exact()
                    .with_retained_bytes(requirements.retained_bytes - 1),
                DecodeLimit::Retained {
                    observed: requirements.retained_bytes,
                    maximum: requirements.retained_bytes - 1,
                },
            ),
            (
                requirements.scratch_bytes,
                requirements
                    .exact()
                    .with_scratch_bytes(requirements.scratch_bytes - 1),
                DecodeLimit::Scratch {
                    observed: requirements.scratch_bytes,
                    maximum: requirements.scratch_bytes - 1,
                },
            ),
        ];
        for (observed, limits, expected) in checks {
            assert!(observed > 0);
            assert_eq!(
                prepared.execute(limits).unwrap_err().resource_limit(),
                Some(expected)
            );
        }
        if requirements.max_depth > 0 {
            assert_eq!(
                prepared
                    .execute(
                        requirements
                            .exact()
                            .with_max_depth(requirements.max_depth - 1)
                    )
                    .unwrap_err()
                    .resource_limit(),
                Some(DecodeLimit::Nesting {
                    observed: requirements.max_depth,
                    maximum: requirements.max_depth - 1,
                })
            );
        }
    }

    #[test]
    fn exact_noop_preserves_every_byte_and_empty_presence() {
        let mut drawable = Vec::new();
        drawable.extend(field_bytes(DRAWABLE_HYPERLINK_FIELD, b""));
        drawable.extend(field_bool(DRAWABLE_LOCKED_FIELD, false));
        drawable.extend(field_bytes(
            DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
            "日本語 ✨".as_bytes(),
        ));
        drawable.extend([0x4b, 0x88, 0x80, 0x00, 0x01, 0x4c]);
        let source = field_bytes(MOVIE_SUPER_FIELD, &drawable);
        let options = DecodeOptions::for_source(&source);
        let snapshot = decode_movie_properties(&source, options).unwrap();
        assert_eq!(snapshot.hyperlink(), Some(""));
        assert_eq!(snapshot.locked(), Some(false));
        assert_eq!(snapshot.accessibility_description(), Some("日本語 ✨"));
        let output =
            rewrite_movie_properties(&source, MoviePropertiesWrite::preserve(), options).unwrap();
        assert_eq!(output, source);
        let output = rewrite_movie_properties(
            &source,
            MoviePropertiesWrite::new(Some(""), Some(false), None, Some("日本語 ✨")),
            options,
        )
        .unwrap();
        assert_eq!(output, source);
        let cleared = rewrite_movie_properties(
            &source,
            MoviePropertiesWrite::preserve().clear_hyperlink(),
            options,
        )
        .unwrap();
        assert_eq!(
            decode_movie_properties(&cleared, DecodeOptions::for_source(&cleared))
                .unwrap()
                .hyperlink(),
            None
        );
        assert!(
            cleared
                .windows(6)
                .any(|window| window == [0x4b, 0x88, 0x80, 0x00, 0x01, 0x4c])
        );
    }

    #[test]
    fn length_prefix_boundaries_are_rewritten_without_losing_opaque_fields() {
        let mut drawable = Vec::new();
        drawable.extend(field_bytes(DRAWABLE_HYPERLINK_FIELD, &vec![b'a'; 125]));
        drawable.extend([0x7a, 0x03, 0x01, 0x02, 0x03]);
        let source = field_bytes(MOVIE_SUPER_FIELD, &drawable);
        let options = DecodeOptions::for_source(&source).with_max_output_bytes(source.len() * 3);
        let output = rewrite_movie_properties(
            &source,
            MoviePropertiesWrite::preserve().set_hyperlink(&"b".repeat(128)),
            options,
        )
        .unwrap();
        assert!(
            output
                .windows(5)
                .any(|window| window == [0x7a, 0x03, 0x01, 0x02, 0x03])
        );
        assert_eq!(
            decode_movie_properties(&output, DecodeOptions::for_source(&output))
                .unwrap()
                .hyperlink()
                .map(str::len),
            Some(128)
        );
    }

    #[test]
    fn unknown_group_depth_is_bounded_and_preserved() {
        let drawable = [0x4b, 0x88, 0x80, 0x00, 0x01, 0x4c];
        let source = field_bytes(MOVIE_SUPER_FIELD, &drawable);
        assert!(decode_movie_properties(&source, DecodeOptions::for_source(&source)).is_ok());
        let shallow = DecodeOptions::for_source(&source).with_recursion_limit(2);
        assert_eq!(
            decode_movie_properties(&source, shallow)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 3,
                maximum: 2,
            })
        );
        let output = rewrite_movie_properties(
            &source,
            MoviePropertiesWrite::preserve().set_locked(true),
            DecodeOptions::for_source(&source),
        )
        .unwrap();
        assert!(output.windows(6).any(|window| window == drawable));
    }

    #[test]
    fn sizing_rejects_usize_overflow_without_allocating() {
        assert!(length_field_len(DRAWABLE_HYPERLINK_FIELD, usize::MAX).is_err());
    }

    #[test]
    fn candidate_hard_message_ceiling_is_checked_before_execution() {
        let source = source();
        let options = DecodeOptions::for_source(&source).with_max_output_bytes(usize::MAX);
        let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES).unwrap();
        let observed = hard_maximum.checked_add(1).unwrap();
        assert_eq!(
            validate_candidate_output_size(observed, options)
                .unwrap_err()
                .resource_limit(),
            Some(DecodeLimit::Output {
                observed,
                maximum: hard_maximum,
            })
        );
    }
}
