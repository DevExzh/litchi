//! Strict, generated-free projection and source-preserving rewrite for the
//! Keynote `TSD.MovieArchive` geometry edge.
//!
//! The geometry transaction owns `MovieArchive.super.geometry.position` and
//! `.size`; the additive transform transaction owns the optional `flags` and
//! `angle` scalars.  Envelope framing and unknown fields remain raw source
//! bytes.  A private Buffa lazy view is used as an ingress cross-check;
//! generated types never cross this module's public boundary.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict scanner precedes the private projection."
)]

use std::{fmt, mem::size_of};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_movie_geometry_generated::LitchiIwaKeynoteMovieGeometryProjection as projection;

const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_GEOMETRY_FIELD: u32 = 1;
const GEOMETRY_POSITION_FIELD: u32 = 1;
const GEOMETRY_SIZE_FIELD: u32 = 2;
const GEOMETRY_FLAGS_FIELD: u32 = 3;
const GEOMETRY_ANGLE_FIELD: u32 = 4;
const POINT_X_FIELD: u32 = 1;
const POINT_Y_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const MAX_RECURSION: u32 = 64;

// These are logical operation-accounting units.  The strict scanner uses one
// fallibly-reserved `Vec<ParsedField>` for each message, while Buffa's lazy
// projection and the rewrite stages have additional bounded working buffers.
// Keep these envelopes deliberately conservative: allocator/RSS telemetry is
// not available at this neutral codec boundary.
const BUFFA_LOGICAL_ALLOCATIONS: usize = 2;
const TRANSFORM_EXECUTE_ALLOCATIONS: usize = 6;
const GEOMETRY_EXECUTE_ALLOCATIONS: usize = 10;
const POSITION_EXECUTE_ALLOCATIONS: usize = 8;
const CANDIDATE_SCAN_ALLOCATIONS: usize = 5 + BUFFA_LOGICAL_ALLOCATIONS;
const TRANSFORM_OUTPUT_BUFFERS: usize = 3;
const GEOMETRY_OUTPUT_BUFFERS: usize = 5;
const POSITION_OUTPUT_BUFFERS: usize = 5;

/// Finite limits for one geometry payload.
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
    /// Construct a finite source/field/work/nesting policy.
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

    /// Construct a bounded convenience profile from one source payload.
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
        // Prepared execution retains the source envelope plus the bounded
        // geometry/drawable/output staging buffers.  This is a logical peak
        // envelope, rather than an allocator or RSS measurement.
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

    /// Replace allocation ceiling used by prepared execution.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace retained-byte ceiling used by prepared execution.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace scratch-byte ceiling used by prepared execution.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    /// Return the source ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_message_bytes
    }

    /// Return the output ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Resource axis rejected by a strict decode or prepared execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input payload bytes.
    Bytes { observed: usize, maximum: usize },
    /// Strict field visits.
    Fields { observed: usize, maximum: usize },
    /// Strict scan work.
    Work { observed: usize, maximum: usize },
    /// Nested protobuf groups/messages.
    Nesting { observed: u32, maximum: u32 },
    /// Logical execution allocations.
    Allocations { observed: usize, maximum: usize },
    /// Retained candidate bytes.
    Retained { observed: usize, maximum: usize },
    /// Reserved scratch bytes.
    Scratch { observed: usize, maximum: usize },
    /// Candidate output bytes.
    Output { observed: usize, maximum: usize },
}

/// Strict geometry codec failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    message: &'static str,
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    const fn plain(message: &'static str) -> Self {
        Self {
            message,
            limit: None,
        }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            message: "movie geometry resource limit exceeded",
            limit: Some(limit),
        }
    }

    /// Return the typed resource limit, if this error is a limit refusal.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }

    /// Alias used by package adapters.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<DecodeLimit> {
        self.limit
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.message)
    }
}

impl std::error::Error for DecodeError {}

/// Validated finite point from `TSP.Point`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    x: f32,
    y: f32,
}

impl Point {
    /// Construct a finite point. Validation occurs when the value is written.
    #[must_use]
    pub const fn new(x: f32, y: f32) -> Self {
        Self { x, y }
    }

    #[must_use]
    pub const fn x(self) -> f32 {
        self.x
    }

    #[must_use]
    pub const fn y(self) -> f32 {
        self.y
    }
}

/// Validated positive finite size from `TSP.Size`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    width: f32,
    height: f32,
}

impl Size {
    /// Construct a positive finite size. Validation occurs when written.
    #[must_use]
    pub const fn new(width: f32, height: f32) -> Self {
        Self { width, height }
    }

    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

/// Borrowed semantic geometry projection.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MovieGeometrySnapshot {
    position: Point,
    size: Size,
}

impl MovieGeometrySnapshot {
    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    #[must_use]
    pub const fn size(self) -> Size {
        self.size
    }

    #[must_use]
    pub const fn x(self) -> f32 {
        self.position.x
    }

    #[must_use]
    pub const fn y(self) -> f32 {
        self.position.y
    }

    #[must_use]
    pub const fn width(self) -> f32 {
        self.size.width
    }

    #[must_use]
    pub const fn height(self) -> f32 {
        self.size.height
    }
}

/// Archive-free projection of the optional movie transform scalars.
///
/// `None` preserves the distinction between an absent protobuf field and an
/// explicit zero value.  The angle is the native wire value in degrees; no
/// range normalization is performed by this neutral codec.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MovieTransformSnapshot {
    flags: Option<u32>,
    angle_degrees: Option<f32>,
}

impl MovieTransformSnapshot {
    /// Return the optional native reflection/transform bitfield.
    #[must_use]
    pub const fn flags(self) -> Option<u32> {
        self.flags
    }

    /// Return the optional native angle in degrees.
    #[must_use]
    pub const fn angle_degrees(self) -> Option<f32> {
        self.angle_degrees
    }

    /// Alias for callers that use the wire field's concise name.
    #[must_use]
    pub const fn angle(self) -> Option<f32> {
        self.angle_degrees
    }
}

/// One optional transform-field update.
///
/// `Preserve` copies the complete source field span, `Set` writes one
/// canonical known field, and `Clear` removes an existing known field.  A
/// clear of an absent field is an exact no-op.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TransformField<T> {
    /// Leave the source field, including its presence and raw framing, alone.
    Preserve,
    /// Set an explicit canonical value.
    Set(T),
    /// Remove the field when it is present.
    Clear,
}

impl<T> TransformField<T> {
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

/// Requested optional movie transform replacement.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MovieTransformWrite {
    flags: TransformField<u32>,
    angle_degrees: TransformField<f32>,
}

impl MovieTransformWrite {
    /// Construct an exact replacement for both optional wire fields.
    ///
    /// `Some(value)` writes a present field and `None` clears that field.
    #[must_use]
    pub const fn new(flags: Option<u32>, angle_degrees: Option<f32>) -> Self {
        Self {
            flags: match flags {
                Some(value) => TransformField::Set(value),
                None => TransformField::Clear,
            },
            angle_degrees: match angle_degrees {
                Some(value) => TransformField::Set(value),
                None => TransformField::Clear,
            },
        }
    }

    /// Construct a write that preserves both optional fields.
    #[must_use]
    pub const fn preserve() -> Self {
        Self {
            flags: TransformField::Preserve,
            angle_degrees: TransformField::Preserve,
        }
    }

    /// Construct a write from explicit preserve/set/clear updates.
    #[must_use]
    pub const fn with_updates(
        flags: TransformField<u32>,
        angle_degrees: TransformField<f32>,
    ) -> Self {
        Self {
            flags,
            angle_degrees,
        }
    }

    /// Alias for [`Self::new`].
    #[must_use]
    pub const fn from_values(flags: Option<u32>, angle_degrees: Option<f32>) -> Self {
        Self::new(flags, angle_degrees)
    }

    /// Replace the flags field, or clear it when `None` is supplied.
    #[must_use]
    pub const fn with_flags(mut self, value: Option<u32>) -> Self {
        self.flags = match value {
            Some(value) => TransformField::Set(value),
            None => TransformField::Clear,
        };
        self
    }

    /// Replace the angle field, or clear it when `None` is supplied.
    #[must_use]
    pub const fn with_angle_degrees(mut self, value: Option<f32>) -> Self {
        self.angle_degrees = match value {
            Some(value) => TransformField::Set(value),
            None => TransformField::Clear,
        };
        self
    }

    /// Alias for [`Self::with_angle_degrees`].
    #[must_use]
    pub const fn with_angle(self, value: Option<f32>) -> Self {
        self.with_angle_degrees(value)
    }

    /// Set one explicit flags value.
    #[must_use]
    pub const fn set_flags(mut self, value: u32) -> Self {
        self.flags = TransformField::Set(value);
        self
    }

    /// Clear the flags field.
    #[must_use]
    pub const fn clear_flags(mut self) -> Self {
        self.flags = TransformField::Clear;
        self
    }

    /// Set one explicit angle value.
    #[must_use]
    pub const fn set_angle_degrees(mut self, value: f32) -> Self {
        self.angle_degrees = TransformField::Set(value);
        self
    }

    /// Alias for [`Self::set_angle_degrees`].
    #[must_use]
    pub const fn set_angle(mut self, value: f32) -> Self {
        self.angle_degrees = TransformField::Set(value);
        self
    }

    /// Clear the angle field.
    #[must_use]
    pub const fn clear_angle_degrees(mut self) -> Self {
        self.angle_degrees = TransformField::Clear;
        self
    }

    /// Alias for [`Self::clear_angle_degrees`].
    #[must_use]
    pub const fn clear_angle(mut self) -> Self {
        self.angle_degrees = TransformField::Clear;
        self
    }

    #[must_use]
    pub const fn flags_update(self) -> TransformField<u32> {
        self.flags
    }

    #[must_use]
    pub const fn angle_update(self) -> TransformField<f32> {
        self.angle_degrees
    }
}

/// Requested position/size replacement. Flags and angle are preserved raw.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MovieGeometryWrite {
    position: Point,
    size: Size,
}

impl MovieGeometryWrite {
    /// Construct a replacement from semantic values.
    #[must_use]
    pub const fn new(position: Point, size: Size) -> Self {
        Self { position, size }
    }

    /// Construct a replacement from four native scalars.
    #[must_use]
    pub const fn from_values(x: f32, y: f32, width: f32, height: f32) -> Self {
        Self::new(Point::new(x, y), Size::new(width, height))
    }

    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    #[must_use]
    pub const fn size(self) -> Size {
        self.size
    }
}

/// Requested replacement for only `MovieArchive.super.geometry.position`.
///
/// The position transaction deliberately does not own the sibling `size`
/// field.  A source may omit that field, or retain an explicit zero size, and
/// the position rewrite preserves its complete source span either way.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MoviePositionWrite {
    position: Point,
}

impl MoviePositionWrite {
    /// Construct a position replacement from a validated-on-write point.
    #[must_use]
    pub const fn new(position: Point) -> Self {
        Self { position }
    }

    /// Construct a position replacement from native coordinates.
    #[must_use]
    pub const fn from_values(x: f32, y: f32) -> Self {
        Self::new(Point::new(x, y))
    }

    /// Return the requested position.
    #[must_use]
    pub const fn position(self) -> Point {
        self.position
    }

    /// Return the requested x coordinate.
    #[must_use]
    pub const fn x(self) -> f32 {
        self.position.x
    }

    /// Return the requested y coordinate.
    #[must_use]
    pub const fn y(self) -> f32 {
        self.position.y
    }
}

/// Exact successful source accounting.
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
    pub output_bytes: usize,
    pub fields: usize,
    pub work_bytes: usize,
    pub max_depth: u32,
    pub allocations: usize,
    pub retained_bytes: usize,
    pub scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
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

    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Execution ceilings replayed before output allocation.
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

/// Candidate bytes and its exact replay report.
#[derive(Debug, Clone, PartialEq)]
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

/// Exact successful rewrite accounting.
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

/// Prepared source rewrite. Preparation performs strict sizing and no output
/// allocation; execution replays the exact requirements before allocating.
#[derive(Debug, Clone, Copy)]
pub struct PreparedMovieGeometryRewrite<'source> {
    source: &'source [u8],
    write: MovieGeometryWrite,
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
    source_report: DecodeReport,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
}

impl<'source> PreparedMovieGeometryRewrite<'source> {
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute the prepared rewrite with exact residual ceilings.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_limits(self.requirements, limits)?;
        let output = emit_rewrite(
            self.source,
            self.write,
            self.requirements.output_bytes,
            self.options,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::plain(
                "geometry rewrite size disagreed with preflight",
            ));
        }
        let candidate_options = DecodeOptions::new(
            output.len(),
            self.requirements.fields,
            self.requirements.work_bytes,
            self.requirements.max_depth,
        )
        .with_max_output_bytes(self.requirements.output_bytes)
        .with_max_allocations(self.requirements.allocations)
        .with_max_retained_bytes(self.requirements.retained_bytes)
        .with_max_scratch_bytes(self.requirements.scratch_bytes);
        let (snapshot, candidate_report) =
            decode_movie_geometry_with_report(&output, candidate_options)?;
        if candidate_report.fields() != self.candidate_fields
            || candidate_report.work_bytes() != self.candidate_work
            || candidate_report.max_depth() != self.candidate_depth
            || snapshot
                != (MovieGeometrySnapshot {
                    position: self.write.position,
                    size: self.write.size,
                })
        {
            return Err(DecodeError::plain(
                "geometry rewrite candidate disagreed with preflight",
            ));
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

/// Prepared source rewrite for only the movie position.
///
/// Preparation performs strict source validation and complete output sizing;
/// execution replays the exact requirements before allocating a candidate.
/// The borrowed source remains the atomic failure boundary.
#[derive(Debug, Clone, Copy)]
pub struct PreparedMoviePositionRewrite<'source> {
    source: &'source [u8],
    write: MoviePositionWrite,
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
    source_report: DecodeReport,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
}

impl<'source> PreparedMoviePositionRewrite<'source> {
    /// Return strict source accounting captured during preparation.
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }

    /// Return the exact residual execution requirements.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute the prepared position rewrite with exact ceilings.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_limits(self.requirements, limits)?;
        let output = emit_position_rewrite(
            self.source,
            self.write,
            self.requirements.output_bytes,
            self.options,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::plain(
                "movie position rewrite size disagreed with preflight",
            ));
        }
        let candidate_options = DecodeOptions::new(
            output.len(),
            self.requirements.fields,
            self.requirements.work_bytes,
            self.requirements.max_depth,
        )
        .with_max_output_bytes(self.requirements.output_bytes)
        .with_max_allocations(self.requirements.allocations)
        .with_max_retained_bytes(self.requirements.retained_bytes)
        .with_max_scratch_bytes(self.requirements.scratch_bytes);
        let (position, candidate_report) =
            decode_movie_position_with_report(&output, candidate_options)?;
        if candidate_report.fields() != self.candidate_fields
            || candidate_report.work_bytes() != self.candidate_work
            || candidate_report.max_depth() != self.candidate_depth
            || position != self.write.position
        {
            return Err(DecodeError::plain(
                "movie position rewrite candidate disagreed with preflight",
            ));
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

/// Prepared optional transform rewrite.  Preparation performs strict source
/// validation and output sizing; execution is the single candidate-producing
/// step.
#[derive(Debug, Clone, Copy)]
pub struct PreparedMovieTransformRewrite<'source> {
    source: &'source [u8],
    write: MovieTransformWrite,
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
    source_report: DecodeReport,
    candidate_fields: usize,
    candidate_work: usize,
    candidate_depth: u32,
    candidate_transform: MovieTransformSnapshot,
}

impl<'source> PreparedMovieTransformRewrite<'source> {
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.source_report
    }

    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute the prepared optional transform rewrite with exact ceilings.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_limits(self.requirements, limits)?;
        let output = emit_transform_rewrite(
            self.source,
            self.write,
            self.requirements.output_bytes,
            self.options,
        )?;
        if output.len() != self.requirements.output_bytes {
            return Err(DecodeError::plain(
                "movie transform rewrite size disagreed with preflight",
            ));
        }
        let candidate_options = DecodeOptions::new(
            output.len(),
            self.requirements.fields,
            self.requirements.work_bytes,
            self.requirements.max_depth,
        )
        .with_max_output_bytes(self.requirements.output_bytes)
        .with_max_allocations(self.requirements.allocations)
        .with_max_retained_bytes(self.requirements.retained_bytes)
        .with_max_scratch_bytes(self.requirements.scratch_bytes);
        let (transform, candidate_report) =
            decode_movie_transform_with_report(&output, candidate_options)?;
        if candidate_report.fields() != self.candidate_fields
            || candidate_report.work_bytes() != self.candidate_work
            || candidate_report.max_depth() != self.candidate_depth
            || transform != self.candidate_transform
        {
            return Err(DecodeError::plain(
                "movie transform rewrite candidate disagreed with preflight",
            ));
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

/// Decode the strict movie geometry projection.
pub fn decode_movie_geometry(
    source: &[u8],
    options: DecodeOptions,
) -> Result<MovieGeometrySnapshot, DecodeError> {
    Ok(decode_movie_geometry_with_report(source, options)?.0)
}

/// Decode geometry and return exact strict source accounting.
pub fn decode_movie_geometry_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(MovieGeometrySnapshot, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let (snapshot, scan) = scan_document(source, options)?;
    force_buffa(source, options)?;
    Ok((snapshot, scan_report(source, scan, options)?))
}

/// Decode the strict optional movie transform projection.
pub fn decode_movie_transform(
    source: &[u8],
    options: DecodeOptions,
) -> Result<MovieTransformSnapshot, DecodeError> {
    Ok(decode_movie_transform_with_report(source, options)?.0)
}

/// Decode transform fields and return exact strict source accounting.
pub fn decode_movie_transform_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(MovieTransformSnapshot, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let (_geometry, transform, scan) = scan_document_projection(source, options)?;
    force_buffa(source, options)?;
    Ok((transform, scan_report(source, scan, options)?))
}

/// Prepare a position/size rewrite without allocating candidate output.
pub fn prepare_movie_geometry_rewrite<'source>(
    source: &'source [u8],
    write: MovieGeometryWrite,
    options: DecodeOptions,
) -> Result<PreparedMovieGeometryRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    validate_write(write)?;
    let (_snapshot, source_scan) = scan_document(source, options)?;
    let source_fields = source_scan.fields;
    force_buffa(source, options)?;
    let output_measure = measure_output(source, write, source_fields, source_scan, options)?;
    if output_measure.bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::Output {
            observed: output_measure.bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let requirements = rewrite_requirements(
        source.len(),
        source_scan,
        output_measure,
        options,
        GEOMETRY_EXECUTE_ALLOCATIONS,
    )?;
    Ok(PreparedMovieGeometryRewrite {
        source,
        write,
        options,
        requirements,
        source_report: scan_report(source, source_scan, options)?,
        candidate_fields: output_measure.fields,
        candidate_work: output_measure.work,
        candidate_depth: output_measure.max_depth,
    })
}

/// Rewrite position and size while preserving every other source span.
pub fn rewrite_movie_geometry(
    source: &[u8],
    write: MovieGeometryWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_movie_geometry_rewrite(source, write, options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_output())
}

/// Decode only the strict movie position projection.
///
/// Unlike [`decode_movie_geometry`], this read does not require the sibling
/// size field or validate its dimensions.  The native host position reader
/// owns this narrower admission: the rooted geometry and position must be
/// present and valid, while size remains source-authoritative.
pub fn decode_movie_position(source: &[u8], options: DecodeOptions) -> Result<Point, DecodeError> {
    Ok(decode_movie_position_with_report(source, options)?.0)
}

/// Decode the strict movie position and return exact source accounting.
pub fn decode_movie_position_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(Point, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let (position, scan) = scan_document_position(source, options)?;
    force_buffa(source, options)?;
    Ok((position, scan_report(source, scan, options)?))
}

/// Prepare a rewrite of only `MovieArchive.super.geometry.position`.
pub fn prepare_movie_position_rewrite<'source>(
    source: &'source [u8],
    write: MoviePositionWrite,
    options: DecodeOptions,
) -> Result<PreparedMoviePositionRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    validate_position_write(write)?;
    let (_position, source_scan) = scan_document_position(source, options)?;
    let source_report = scan_report(source, source_scan, options)?;
    force_buffa(source, options)?;
    let measurement_options = residual_options(options, source_report)?;
    let output_measure =
        measure_position_output(source, source_scan, measurement_options, options)?;
    if output_measure.bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::Output {
            observed: output_measure.bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let requirements = position_rewrite_requirements(
        source.len(),
        source_scan,
        output_measure,
        options,
        POSITION_EXECUTE_ALLOCATIONS,
    )?;
    Ok(PreparedMoviePositionRewrite {
        source,
        write,
        options,
        requirements,
        source_report: combine_position_prepare_report(source_report, output_measure, options)?,
        candidate_fields: output_measure.fields,
        candidate_work: output_measure.work,
        candidate_depth: output_measure.max_depth,
    })
}

/// Rewrite only the movie position while preserving every other source span.
pub fn rewrite_movie_position(
    source: &[u8],
    write: MoviePositionWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_movie_position_rewrite(source, write, options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_output())
}

/// Prepare an optional transform rewrite without allocating candidate output.
pub fn prepare_movie_transform_rewrite<'source>(
    source: &'source [u8],
    write: MovieTransformWrite,
    options: DecodeOptions,
) -> Result<PreparedMovieTransformRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    validate_transform_write(write)?;
    let (_geometry, transform, source_scan) = scan_document_projection(source, options)?;
    force_buffa(source, options)?;
    let output_measure = measure_transform_output(source, write, source_scan, options)?;
    if output_measure.bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::Output {
            observed: output_measure.bytes,
            maximum: options.max_output_bytes,
        }));
    }
    let requirements = rewrite_requirements(
        source.len(),
        source_scan,
        output_measure,
        options,
        TRANSFORM_EXECUTE_ALLOCATIONS,
    )?;
    Ok(PreparedMovieTransformRewrite {
        source,
        write,
        options,
        requirements,
        source_report: scan_report(source, source_scan, options)?,
        candidate_fields: output_measure.fields,
        candidate_work: output_measure.work,
        candidate_depth: output_measure.max_depth,
        candidate_transform: transform_after_write(transform, write),
    })
}

/// Rewrite optional transform fields while preserving every other source span.
pub fn rewrite_movie_transform(
    source: &[u8],
    write: MovieTransformWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_movie_transform_rewrite(source, write, options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_output())
}

#[derive(Debug, Clone, Copy)]
struct ParsedField {
    number: u32,
    wire: u8,
    start: usize,
    value_start: usize,
    value_end: usize,
    end: usize,
    value: Option<u64>,
    field_count: usize,
}

#[derive(Default)]
struct Budget {
    fields: usize,
    work: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct ScanAccounting {
    fields: usize,
    work: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct OutputMeasure {
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    measurement_fields: usize,
    measurement_work: usize,
    measurement_allocations: usize,
    measurement_scratch: usize,
}

#[derive(Debug, Clone, Copy, Default)]
struct TransformDelta {
    add_fields: usize,
    remove_fields: usize,
    add_work: usize,
    remove_work: usize,
}

fn scan_report(
    source: &[u8],
    scan: ScanAccounting,
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    let allocations = scan
        .allocations
        .checked_add(BUFFA_LOGICAL_ALLOCATIONS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    if allocations > options.max_allocations {
        return Err(DecodeError::limited(DecodeLimit::Allocations {
            observed: allocations,
            maximum: options.max_allocations,
        }));
    }
    let work_bytes = scan.work.checked_add(source.len()).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        })
    })?;
    if work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    let scratch_bytes = scan
        .scratch_bytes
        .checked_add(source.len())
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    if scratch_bytes > options.max_scratch_bytes {
        return Err(DecodeError::limited(DecodeLimit::Scratch {
            observed: scratch_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    Ok(DecodeReport {
        input_bytes: source.len(),
        fields: scan.fields,
        work_bytes,
        max_depth: scan.max_depth,
        allocations,
        retained_bytes: source.len(),
        scratch_bytes,
    })
}

fn residual_options(
    options: DecodeOptions,
    consumed: DecodeReport,
) -> Result<DecodeOptions, DecodeError> {
    let fields = options
        .max_fields
        .checked_sub(consumed.fields)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: options.max_fields.saturating_add(1),
                maximum: options.max_fields,
            })
        })?;
    let work_bytes = options
        .max_work_bytes
        .checked_sub(consumed.work_bytes)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: options.max_work_bytes.saturating_add(1),
                maximum: options.max_work_bytes,
            })
        })?;
    let allocations = options
        .max_allocations
        .checked_sub(consumed.allocations)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Allocations {
                observed: options.max_allocations.saturating_add(1),
                maximum: options.max_allocations,
            })
        })?;
    let scratch_bytes = options
        .max_scratch_bytes
        .checked_sub(consumed.scratch_bytes)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: options.max_scratch_bytes.saturating_add(1),
                maximum: options.max_scratch_bytes,
            })
        })?;
    Ok(options
        .with_resource_limits(fields, work_bytes)
        .with_max_allocations(allocations)
        .with_max_scratch_bytes(scratch_bytes))
}

fn combine_position_prepare_report(
    source_report: DecodeReport,
    measure: OutputMeasure,
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    let fields = source_report
        .fields
        .checked_add(measure.measurement_fields)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    let work_bytes = source_report
        .work_bytes
        .checked_add(measure.measurement_work)
        .and_then(|value| value.checked_add(source_report.input_bytes))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let allocations = source_report
        .allocations
        .checked_add(measure.measurement_allocations)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    let scratch_bytes = source_report
        .scratch_bytes
        .checked_add(measure.measurement_scratch)
        .and_then(|value| value.checked_add(source_report.input_bytes))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let report = DecodeReport {
        input_bytes: source_report.input_bytes,
        fields,
        work_bytes,
        max_depth: source_report.max_depth.max(measure.max_depth),
        allocations,
        retained_bytes: source_report.retained_bytes,
        scratch_bytes,
    };
    check_decode_report_limits(report, options)?;
    Ok(report)
}

fn check_decode_report_limits(
    report: DecodeReport,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
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
            report.scratch_bytes,
            options.max_scratch_bytes,
            DecodeLimit::Scratch {
                observed: report.scratch_bytes,
                maximum: options.max_scratch_bytes,
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
    ] {
        if observed > maximum {
            return Err(DecodeError::limited(limit));
        }
    }
    if report.max_depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: report.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    Ok(())
}

fn position_rewrite_requirements(
    source_bytes: usize,
    source_scan: ScanAccounting,
    output_measure: OutputMeasure,
    options: DecodeOptions,
    execute_allocations: usize,
) -> Result<RewriteExecutionRequirements, DecodeError> {
    let fields = source_scan
        .fields
        .checked_add(output_measure.fields)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    let work_bytes = source_scan
        .work
        .checked_add(output_measure.work)
        .and_then(|value| value.checked_add(source_bytes))
        .and_then(|value| value.checked_add(output_measure.bytes))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let candidate_scratch = output_measure
        .scratch_bytes
        .checked_sub(output_measure.measurement_scratch)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let scratch_bytes = source_scan
        .scratch_bytes
        .checked_add(candidate_scratch)
        .and_then(|value| value.checked_add(source_bytes))
        .and_then(|value| value.checked_add(output_measure.bytes))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let allocations = source_scan
        .allocations
        .checked_add(execute_allocations)
        .and_then(|value| value.checked_add(CANDIDATE_SCAN_ALLOCATIONS))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    let retained_bytes = source_bytes
        .checked_add(output_measure.retained_bytes)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: output_measure.bytes,
        fields,
        work_bytes,
        max_depth: output_measure.max_depth,
        allocations,
        retained_bytes,
        scratch_bytes,
    };
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
            return Err(DecodeError::limited(limit));
        }
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    Ok(requirements)
}

fn rewrite_requirements(
    source_bytes: usize,
    source_scan: ScanAccounting,
    output_measure: OutputMeasure,
    options: DecodeOptions,
    execute_allocations: usize,
) -> Result<RewriteExecutionRequirements, DecodeError> {
    let fields = source_scan
        .fields
        .checked_add(output_measure.fields)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: options.max_fields,
            })
        })?;
    // The source is parsed for admission, sizing, and emission.  The
    // candidate is emitted and then parsed again for exact readback.  Include
    // one Buffa pass for each side and the copied output bytes; these are
    // conservative logical work units, not CPU-instruction telemetry.
    let work_bytes = source_scan
        .work
        .checked_mul(3)
        .and_then(|value| value.checked_add(output_measure.work.checked_mul(2)?))
        .and_then(|value| value.checked_add(source_bytes))
        .and_then(|value| value.checked_add(output_measure.bytes))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: options.max_work_bytes,
            })
        })?;
    let scratch_bytes = source_scan
        .scratch_bytes
        .checked_add(output_measure.scratch_bytes)
        .and_then(|value| value.checked_add(source_bytes))
        .and_then(|value| value.checked_add(output_measure.bytes))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let allocations = source_scan
        .allocations
        .checked_add(BUFFA_LOGICAL_ALLOCATIONS)
        .and_then(|value| value.checked_add(output_measure.allocations))
        .and_then(|value| value.checked_add(execute_allocations))
        .and_then(|value| value.checked_add(CANDIDATE_SCAN_ALLOCATIONS))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Allocations {
                observed: usize::MAX,
                maximum: options.max_allocations,
            })
        })?;
    let retained_bytes = source_bytes
        .checked_add(output_measure.retained_bytes)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    let requirements = RewriteExecutionRequirements {
        output_bytes: output_measure.bytes,
        fields,
        work_bytes,
        max_depth: output_measure.max_depth,
        allocations,
        retained_bytes,
        scratch_bytes,
    };
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
            return Err(DecodeError::limited(limit));
        }
    }
    Ok(requirements)
}

fn transform_after_write(
    current: MovieTransformSnapshot,
    write: MovieTransformWrite,
) -> MovieTransformSnapshot {
    MovieTransformSnapshot {
        flags: match write.flags {
            TransformField::Preserve => current.flags,
            TransformField::Set(value) => Some(value),
            TransformField::Clear => None,
        },
        angle_degrees: match write.angle_degrees {
            TransformField::Preserve => current.angle_degrees,
            TransformField::Set(value) => Some(value),
            TransformField::Clear => None,
        },
    }
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: options.recursion_limit,
            maximum: MAX_RECURSION,
        }));
    }
    Ok(())
}

fn validate_write(write: MovieGeometryWrite) -> Result<(), DecodeError> {
    validate_position(write.position)?;
    if !write.size.width.is_finite()
        || !write.size.height.is_finite()
        || write.size.width <= 0.0
        || write.size.height <= 0.0
    {
        return Err(DecodeError::plain(
            "movie geometry size must be finite and positive",
        ));
    }
    Ok(())
}

fn validate_position_write(write: MoviePositionWrite) -> Result<(), DecodeError> {
    validate_position(write.position)
}

fn validate_position(position: Point) -> Result<(), DecodeError> {
    if !position.x.is_finite() || !position.y.is_finite() {
        return Err(DecodeError::plain("movie geometry position must be finite"));
    }
    Ok(())
}

fn validate_transform_write(write: MovieTransformWrite) -> Result<(), DecodeError> {
    if let TransformField::Set(value) = write.angle_degrees {
        if !value.is_finite() {
            return Err(DecodeError::plain("movie transform angle must be finite"));
        }
    }
    Ok(())
}

fn force_buffa(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let view: projection::MovieArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::plain("Buffa geometry projection rejected source"))?;
    let drawable = view
        .super_
        .get()
        .map_err(|_| DecodeError::plain("Buffa geometry projection rejected super"))?
        .ok_or_else(|| DecodeError::plain("Buffa geometry projection missing super"))?;
    let _geometry = drawable
        .geometry
        .get()
        .map_err(|_| DecodeError::plain("Buffa geometry projection rejected geometry"))?
        .ok_or_else(|| DecodeError::plain("Buffa geometry projection missing geometry"))?;
    Ok(())
}

fn scan_document(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(MovieGeometrySnapshot, ScanAccounting), DecodeError> {
    let (snapshot, _transform, scan) = scan_document_projection(source, options)?;
    Ok((snapshot, scan))
}

fn scan_document_position(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(Point, ScanAccounting), DecodeError> {
    let (position, _transform, scan) = scan_document_position_projection(source, options)?;
    Ok((position, scan))
}

fn scan_document_position_projection(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(Point, MovieTransformSnapshot, ScanAccounting), DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    require_wire(super_field, 2)?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    require_wire(geometry_field, 2)?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    require_wire(position_field, 2)?;
    // Size is deliberately outside this transaction's semantic ownership.
    // Still reject duplicate or wrong-wire known fields so the handwritten
    // admission and the private Buffa view agree on the source envelope.
    if let Some(size_field) = optional_known(&geometry_fields, GEOMETRY_SIZE_FIELD)? {
        require_wire(size_field, 2)?;
    }
    let transform = parse_transform(&geometry_fields, geometry)?;
    let position = parse_point(
        &geometry[position_field.value_start..position_field.value_end],
        options,
        &mut budget,
    )?;
    Ok((
        position,
        transform,
        ScanAccounting {
            fields: budget.fields,
            work: budget.work,
            max_depth: budget.max_depth,
            allocations: budget.allocations,
            scratch_bytes: budget.scratch_bytes,
        },
    ))
}

fn scan_document_projection(
    source: &[u8],
    options: DecodeOptions,
) -> Result<
    (
        MovieGeometrySnapshot,
        MovieTransformSnapshot,
        ScanAccounting,
    ),
    DecodeError,
> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    require_wire(super_field, 2)?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    require_wire(geometry_field, 2)?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    let size_field = unique_known(&geometry_fields, GEOMETRY_SIZE_FIELD, "Geometry.size")?;
    require_wire(position_field, 2)?;
    require_wire(size_field, 2)?;
    let transform = parse_transform(&geometry_fields, geometry)?;
    let position = parse_point(
        &geometry[position_field.value_start..position_field.value_end],
        options,
        &mut budget,
    )?;
    let size = parse_size(
        &geometry[size_field.value_start..size_field.value_end],
        options,
        &mut budget,
    )?;
    Ok((
        MovieGeometrySnapshot { position, size },
        transform,
        ScanAccounting {
            fields: budget.fields,
            work: budget.work,
            max_depth: budget.max_depth,
            allocations: budget.allocations,
            scratch_bytes: budget.scratch_bytes,
        },
    ))
}

fn parse_point(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Point, DecodeError> {
    let fields = parse_message(source, options, budget, 4)?;
    let x = unique_known(&fields, POINT_X_FIELD, "Point.x")?;
    let y = unique_known(&fields, POINT_Y_FIELD, "Point.y")?;
    require_wire(x, 5)?;
    require_wire(y, 5)?;
    let x = fixed32(source, x)?;
    let y = fixed32(source, y)?;
    if !x.is_finite() || !y.is_finite() {
        return Err(DecodeError::plain("Point coordinates must be finite"));
    }
    Ok(Point::new(x, y))
}

fn parse_size(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Size, DecodeError> {
    let fields = parse_message(source, options, budget, 4)?;
    let width = unique_known(&fields, SIZE_WIDTH_FIELD, "Size.width")?;
    let height = unique_known(&fields, SIZE_HEIGHT_FIELD, "Size.height")?;
    require_wire(width, 5)?;
    require_wire(height, 5)?;
    let width = fixed32(source, width)?;
    let height = fixed32(source, height)?;
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return Err(DecodeError::plain("Size must be finite and positive"));
    }
    Ok(Size::new(width, height))
}

fn parse_transform(
    fields: &[ParsedField],
    source: &[u8],
) -> Result<MovieTransformSnapshot, DecodeError> {
    let flags_field = optional_known(fields, GEOMETRY_FLAGS_FIELD)?;
    let flags = if let Some(field) = flags_field {
        require_wire(field, 0)?;
        let value = field
            .value
            .ok_or_else(|| DecodeError::plain("invalid geometry flags"))?;
        if varint_len(value) != field.value_end.saturating_sub(field.value_start)
            || value > u64::from(u32::MAX)
        {
            return Err(DecodeError::plain("noncanonical geometry flags"));
        }
        Some(value as u32)
    } else {
        None
    };
    let angle_field = optional_known(fields, GEOMETRY_ANGLE_FIELD)?;
    let angle_degrees = if let Some(field) = angle_field {
        require_wire(field, 5)?;
        let value = fixed32(source, field)?;
        if !value.is_finite() {
            return Err(DecodeError::plain("geometry angle must be finite"));
        }
        Some(value)
    } else {
        None
    };
    Ok(MovieTransformSnapshot {
        flags,
        angle_degrees,
    })
}

fn optional_known(
    fields: &[ParsedField],
    number: u32,
) -> Result<Option<&ParsedField>, DecodeError> {
    let mut found = None;
    for field in fields.iter().filter(|field| field.number == number) {
        if found.is_some() {
            return Err(DecodeError::plain("duplicate known geometry field"));
        }
        found = Some(field);
    }
    Ok(found)
}

fn unique_known<'a>(
    fields: &'a [ParsedField],
    number: u32,
    name: &'static str,
) -> Result<&'a ParsedField, DecodeError> {
    let mut found = None;
    for field in fields.iter().filter(|field| field.number == number) {
        if found.is_some() {
            return Err(DecodeError::plain("duplicate known geometry field"));
        }
        found = Some(field);
    }
    found.ok_or_else(|| DecodeError::plain(name))
}

fn require_wire(field: &ParsedField, wire: u8) -> Result<(), DecodeError> {
    if field.wire != wire {
        return Err(DecodeError::plain(
            "wrong wire type for known geometry field",
        ));
    }
    Ok(())
}

fn fixed32(source: &[u8], field: &ParsedField) -> Result<f32, DecodeError> {
    let bytes: [u8; 4] = source[field.value_start..field.value_end]
        .try_into()
        .map_err(|_| DecodeError::plain("invalid fixed32 geometry value"))?;
    Ok(f32::from_le_bytes(bytes))
}

fn parse_message(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<Vec<ParsedField>, DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: depth,
            maximum: options.recursion_limit,
        }));
    }
    budget.max_depth = budget.max_depth.max(depth);
    let capacity = source.len().min(options.max_fields);
    let scratch = capacity
        .checked_mul(size_of::<ParsedField>())
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let scratch_bytes = budget.scratch_bytes.checked_add(scratch).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Scratch {
            observed: usize::MAX,
            maximum: options.max_scratch_bytes,
        })
    })?;
    if scratch_bytes > options.max_scratch_bytes {
        return Err(DecodeError::limited(DecodeLimit::Scratch {
            observed: scratch_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    let allocation = usize::from(capacity != 0);
    let allocations = budget.allocations.checked_add(allocation).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Allocations {
            observed: usize::MAX,
            maximum: options.max_allocations,
        })
    })?;
    if allocations > options.max_allocations {
        return Err(DecodeError::limited(DecodeLimit::Allocations {
            observed: allocations,
            maximum: options.max_allocations,
        }));
    }
    let mut fields = Vec::new();
    fields.try_reserve_exact(capacity).map_err(|_| {
        DecodeError::limited(DecodeLimit::Scratch {
            observed: scratch_bytes,
            maximum: options.max_scratch_bytes,
        })
    })?;
    budget.allocations = allocations;
    budget.scratch_bytes = scratch_bytes;
    let mut offset = 0;
    while offset < source.len() {
        let (field, next) = parse_field(source, offset, source.len(), depth, options, budget)?;
        fields.push(field);
        offset = next;
    }
    Ok(fields)
}

fn parsed_fields_scratch_envelope(
    source_bytes: usize,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    source_bytes
        .min(options.max_fields)
        .checked_mul(size_of::<ParsedField>())
        .and_then(|value| value.checked_mul(5))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })
}

fn parse_field(
    source: &[u8],
    offset: usize,
    end: usize,
    depth: u32,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(ParsedField, usize), DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: depth,
            maximum: options.recursion_limit,
        }));
    }
    let (key, key_end) = read_varint(source, offset, end)?;
    let number_value = key >> 3;
    let wire = (key & 7) as u8;
    if number_value == 0 || number_value > 0x1fff_ffff {
        return Err(DecodeError::plain("invalid protobuf field number"));
    }
    let number = number_value as u32;
    let known = is_known_at_depth(number, depth);
    if known && key_end - offset != varint_len(key) {
        return Err(DecodeError::plain("noncanonical known geometry key"));
    }
    visit_field(budget, options, 1, key_end.saturating_sub(offset))?;
    let (value_start, value_end, next, child_count) = match wire {
        0 => {
            let (_, value_end) = read_varint(source, key_end, end)?;
            (key_end, value_end, value_end, 0)
        },
        1 => {
            let value_end = key_end
                .checked_add(8)
                .ok_or_else(|| DecodeError::plain("fixed64 overflow"))?;
            if value_end > end {
                return Err(DecodeError::plain("truncated fixed64"));
            }
            (key_end, value_end, value_end, 0)
        },
        2 => {
            let (length, payload_start) = read_varint(source, key_end, end)?;
            if known && payload_start - key_end != varint_len(length) {
                return Err(DecodeError::plain("noncanonical known length"));
            }
            let length =
                usize::try_from(length).map_err(|_| DecodeError::plain("length overflow"))?;
            let value_end = payload_start
                .checked_add(length)
                .ok_or_else(|| DecodeError::plain("length overflow"))?;
            if value_end > end {
                return Err(DecodeError::plain("truncated length-delimited field"));
            }
            (payload_start, value_end, value_end, 0)
        },
        3 => {
            if depth >= options.recursion_limit {
                return Err(DecodeError::limited(DecodeLimit::Nesting {
                    observed: depth + 1,
                    maximum: options.recursion_limit,
                }));
            }
            let mut cursor = key_end;
            let mut count = 0usize;
            loop {
                if cursor >= end {
                    return Err(DecodeError::plain("unterminated protobuf group"));
                }
                let (child_key, child_key_end) = read_varint(source, cursor, end)?;
                let child_number = (child_key >> 3) as u32;
                let child_wire = (child_key & 7) as u8;
                if child_wire == 4 {
                    if child_number != number {
                        return Err(DecodeError::plain("mismatched protobuf group"));
                    }
                    visit_field(budget, options, 1, child_key_end - cursor)?;
                    cursor = child_key_end;
                    break;
                }
                let (child, child_next) =
                    parse_field(source, cursor, end, depth + 1, options, budget)?;
                count = count.saturating_add(child.field_count);
                cursor = child_next;
            }
            (key_end, cursor, cursor, count.saturating_add(1))
        },
        4 => return Err(DecodeError::plain("unexpected protobuf end group")),
        5 => {
            let value_end = key_end
                .checked_add(4)
                .ok_or_else(|| DecodeError::plain("fixed32 overflow"))?;
            if value_end > end {
                return Err(DecodeError::plain("truncated fixed32"));
            }
            (key_end, value_end, value_end, 0)
        },
        _ => return Err(DecodeError::plain("invalid protobuf wire type")),
    };
    let value = if wire == 0 {
        Some(read_varint(source, value_start, value_end)?.0)
    } else {
        None
    };
    let span = next.saturating_sub(offset);
    visit_field(budget, options, 0, span)?;
    Ok((
        ParsedField {
            number,
            wire,
            start: offset,
            value_start,
            value_end,
            end: next,
            value,
            field_count: child_count.saturating_add(1),
        },
        next,
    ))
}

fn visit_field(
    budget: &mut Budget,
    options: DecodeOptions,
    fields: usize,
    work: usize,
) -> Result<(), DecodeError> {
    budget.fields = budget.fields.checked_add(fields).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Fields {
            observed: usize::MAX,
            maximum: options.max_fields,
        })
    })?;
    if budget.fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: budget.fields,
            maximum: options.max_fields,
        }));
    }
    budget.work = budget.work.checked_add(work).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        })
    })?;
    if budget.work > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: budget.work,
            maximum: options.max_work_bytes,
        }));
    }
    Ok(())
}

fn is_known_at_depth(number: u32, depth: u32) -> bool {
    match depth {
        1 => number == MOVIE_SUPER_FIELD,
        2 => number == DRAWABLE_GEOMETRY_FIELD,
        3 => matches!(
            number,
            GEOMETRY_POSITION_FIELD
                | GEOMETRY_SIZE_FIELD
                | GEOMETRY_FLAGS_FIELD
                | GEOMETRY_ANGLE_FIELD
        ),
        4 => matches!(number, POINT_X_FIELD | POINT_Y_FIELD),
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

fn measure_output(
    source: &[u8],
    _write: MovieGeometryWrite,
    source_fields: usize,
    source_scan: ScanAccounting,
    options: DecodeOptions,
) -> Result<OutputMeasure, DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    let size_field = unique_known(&geometry_fields, GEOMETRY_SIZE_FIELD, "Geometry.size")?;
    let point = &geometry[position_field.value_start..position_field.value_end];
    let size = &geometry[size_field.value_start..size_field.value_end];
    let point_fields = parse_message(point, options, &mut budget, 4)?;
    let size_fields = parse_message(size, options, &mut budget, 4)?;
    let point_len = encoded_fixed_pair_len(&point_fields, POINT_X_FIELD, POINT_Y_FIELD)?;
    let size_len = encoded_fixed_pair_len(&size_fields, SIZE_WIDTH_FIELD, SIZE_HEIGHT_FIELD)?;
    let new_geometry_len = geometry
        .len()
        .saturating_sub(position_field.end - position_field.start)
        .saturating_sub(size_field.end - size_field.start)
        .saturating_add(length_field_len(GEOMETRY_POSITION_FIELD, point_len))
        .saturating_add(length_field_len(GEOMETRY_SIZE_FIELD, size_len));
    let new_drawable_len = drawable
        .len()
        .saturating_sub(geometry_field.end - geometry_field.start)
        .saturating_add(length_field_len(DRAWABLE_GEOMETRY_FIELD, new_geometry_len));
    let output_len = source
        .len()
        .saturating_sub(super_field.end - super_field.start)
        .saturating_add(length_field_len(MOVIE_SUPER_FIELD, new_drawable_len));
    let candidate_fields = source_scan.fields.max(source_fields);
    let candidate_work = source_scan.work.checked_add(output_len).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        })
    })?;
    let candidate_parse_scratch = parsed_fields_scratch_envelope(output_len, options)?;
    let output_scratch = output_len
        .checked_mul(GEOMETRY_OUTPUT_BUFFERS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let scratch = budget
        .scratch_bytes
        .checked_add(candidate_parse_scratch)
        .and_then(|value| value.checked_add(output_scratch))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let retained_bytes = output_len
        .checked_mul(GEOMETRY_OUTPUT_BUFFERS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    Ok(OutputMeasure {
        bytes: output_len,
        fields: candidate_fields,
        work: candidate_work,
        max_depth: source_scan.max_depth,
        allocations: budget.allocations,
        retained_bytes,
        scratch_bytes: scratch,
        measurement_fields: budget.fields,
        measurement_work: budget.work,
        measurement_allocations: budget.allocations,
        measurement_scratch: budget.scratch_bytes,
    })
}

fn measure_position_output(
    source: &[u8],
    source_scan: ScanAccounting,
    parse_options: DecodeOptions,
    sizing_options: DecodeOptions,
) -> Result<OutputMeasure, DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, parse_options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, parse_options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, parse_options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    require_wire(position_field, 2)?;
    if let Some(size_field) = optional_known(&geometry_fields, GEOMETRY_SIZE_FIELD)? {
        require_wire(size_field, 2)?;
    }
    let _ = parse_transform(&geometry_fields, geometry)?;
    let point = &geometry[position_field.value_start..position_field.value_end];
    let point_fields = parse_message(point, parse_options, &mut budget, 4)?;
    let point_len = encoded_fixed_pair_len(&point_fields, POINT_X_FIELD, POINT_Y_FIELD)?;
    let new_geometry_len = replace_length(
        geometry.len(),
        field_span(position_field),
        length_field_len(GEOMETRY_POSITION_FIELD, point_len),
    )?;
    let new_drawable_len = replace_length(
        drawable.len(),
        field_span(geometry_field),
        length_field_len(DRAWABLE_GEOMETRY_FIELD, new_geometry_len),
    )?;
    let output_len = replace_length(
        source.len(),
        field_span(super_field),
        length_field_len(MOVIE_SUPER_FIELD, new_drawable_len),
    )?;
    let candidate_fields = source_scan.fields;
    let candidate_work = source_scan.work.checked_add(output_len).ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Work {
            observed: usize::MAX,
            maximum: sizing_options.max_work_bytes,
        })
    })?;
    let candidate_parse_scratch = parsed_fields_scratch_envelope(output_len, sizing_options)?;
    let output_scratch = output_len
        .checked_mul(POSITION_OUTPUT_BUFFERS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: sizing_options.max_scratch_bytes,
            })
        })?;
    let scratch = budget
        .scratch_bytes
        .checked_add(candidate_parse_scratch)
        .and_then(|value| value.checked_add(output_scratch))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: sizing_options.max_scratch_bytes,
            })
        })?;
    let retained_bytes = output_len
        .checked_mul(POSITION_OUTPUT_BUFFERS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Retained {
                observed: usize::MAX,
                maximum: sizing_options.max_retained_bytes,
            })
        })?;
    Ok(OutputMeasure {
        bytes: output_len,
        fields: candidate_fields,
        work: candidate_work,
        max_depth: source_scan.max_depth,
        allocations: budget.allocations,
        retained_bytes,
        scratch_bytes: scratch,
        measurement_fields: budget.fields,
        measurement_work: budget.work,
        measurement_allocations: budget.allocations,
        measurement_scratch: budget.scratch_bytes,
    })
}

fn measure_transform_output(
    source: &[u8],
    write: MovieTransformWrite,
    source_scan: ScanAccounting,
    options: DecodeOptions,
) -> Result<OutputMeasure, DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    let size_field = unique_known(&geometry_fields, GEOMETRY_SIZE_FIELD, "Geometry.size")?;
    require_wire(position_field, 2)?;
    require_wire(size_field, 2)?;
    let _ = parse_transform(&geometry_fields, geometry)?;
    let flags_field = optional_known(&geometry_fields, GEOMETRY_FLAGS_FIELD)?;
    let angle_field = optional_known(&geometry_fields, GEOMETRY_ANGLE_FIELD)?;
    let new_geometry_len = replace_optional_fields_len(
        geometry.len(),
        flags_field,
        write.flags,
        angle_field,
        write.angle_degrees,
    )?;
    let new_drawable_len = replace_length(
        drawable.len(),
        geometry_field.end - geometry_field.start,
        length_field_len(DRAWABLE_GEOMETRY_FIELD, new_geometry_len),
    )?;
    let output_len = replace_length(
        source.len(),
        super_field.end - super_field.start,
        length_field_len(MOVIE_SUPER_FIELD, new_drawable_len),
    )?;
    let flags_delta = transform_delta(
        flags_field,
        write.flags,
        encoded_flags_len(flags_field, write.flags),
    );
    let angle_delta = transform_delta(
        angle_field,
        write.angle_degrees,
        encoded_angle_len(angle_field, write.angle_degrees),
    );
    let delta = TransformDelta {
        add_fields: flags_delta
            .add_fields
            .checked_add(angle_delta.add_fields)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                })
            })?,
        remove_fields: flags_delta
            .remove_fields
            .checked_add(angle_delta.remove_fields)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                })
            })?,
        add_work: flags_delta
            .add_work
            .checked_add(angle_delta.add_work)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?,
        remove_work: flags_delta
            .remove_work
            .checked_add(angle_delta.remove_work)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?,
    };
    let mut delta = delta;
    add_work_delta(
        &mut delta,
        geometry_field.end - geometry_field.start,
        length_field_len(DRAWABLE_GEOMETRY_FIELD, new_geometry_len),
    );
    add_work_delta(
        &mut delta,
        super_field.end - super_field.start,
        length_field_len(MOVIE_SUPER_FIELD, new_drawable_len),
    );
    let candidate_fields = apply_delta(
        source_scan.fields,
        delta.add_fields,
        delta.remove_fields,
        DecodeLimit::Fields {
            observed: usize::MAX,
            maximum: options.max_fields,
        },
    )?;
    let candidate_work = apply_delta(
        source_scan.work,
        delta.add_work,
        delta.remove_work,
        DecodeLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        },
    )?
    .checked_add(output_len)
    .ok_or_else(|| {
        DecodeError::limited(DecodeLimit::Work {
            observed: usize::MAX,
            maximum: options.max_work_bytes,
        })
    })?;
    let candidate_parse_scratch = parsed_fields_scratch_envelope(output_len, options)?;
    let output_scratch = output_len
        .checked_mul(TRANSFORM_OUTPUT_BUFFERS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let scratch = budget
        .scratch_bytes
        .checked_add(candidate_parse_scratch)
        .and_then(|value| value.checked_add(output_scratch))
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Scratch {
                observed: usize::MAX,
                maximum: options.max_scratch_bytes,
            })
        })?;
    let retained_bytes = output_len
        .checked_mul(TRANSFORM_OUTPUT_BUFFERS)
        .ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Retained {
                observed: usize::MAX,
                maximum: options.max_retained_bytes,
            })
        })?;
    Ok(OutputMeasure {
        bytes: output_len,
        fields: candidate_fields,
        work: candidate_work,
        max_depth: source_scan.max_depth,
        allocations: budget.allocations,
        retained_bytes,
        scratch_bytes: scratch,
        measurement_fields: budget.fields,
        measurement_work: budget.work,
        measurement_allocations: budget.allocations,
        measurement_scratch: budget.scratch_bytes,
    })
}

fn replace_optional_fields_len(
    original: usize,
    flags_field: Option<&ParsedField>,
    flags: TransformField<u32>,
    angle_field: Option<&ParsedField>,
    angle: TransformField<f32>,
) -> Result<usize, DecodeError> {
    let after_flags = replace_length(
        original,
        flags_field.map_or(0, field_span),
        encoded_flags_len(flags_field, flags),
    )?;
    replace_length(
        after_flags,
        angle_field.map_or(0, field_span),
        encoded_angle_len(angle_field, angle),
    )
}

fn encoded_flags_len(field: Option<&ParsedField>, update: TransformField<u32>) -> usize {
    match update {
        TransformField::Preserve => field.map_or(0, field_span),
        TransformField::Set(value) => varint_len(u64::from(GEOMETRY_FLAGS_FIELD << 3))
            .saturating_add(varint_len(u64::from(value))),
        TransformField::Clear => 0,
    }
}

fn encoded_angle_len(field: Option<&ParsedField>, update: TransformField<f32>) -> usize {
    match update {
        TransformField::Preserve => field.map_or(0, field_span),
        TransformField::Set(_) => varint_len(u64::from(GEOMETRY_ANGLE_FIELD << 3 | 5)) + 4,
        TransformField::Clear => 0,
    }
}

fn field_span(field: &ParsedField) -> usize {
    field.end.saturating_sub(field.start)
}

fn replace_length(
    original: usize,
    removed: usize,
    replacement: usize,
) -> Result<usize, DecodeError> {
    original
        .checked_sub(removed)
        .and_then(|value| value.checked_add(replacement))
        .ok_or_else(|| DecodeError::plain("movie transform output length overflow"))
}

fn transform_delta<T>(
    field: Option<&ParsedField>,
    update: TransformField<T>,
    replacement_len: usize,
) -> TransformDelta {
    let Some(field) = field else {
        return match update {
            TransformField::Set(_) => TransformDelta {
                add_fields: 1,
                add_work: 1usize.saturating_add(replacement_len),
                ..TransformDelta::default()
            },
            TransformField::Preserve | TransformField::Clear => TransformDelta::default(),
        };
    };
    let old_span = field_span(field);
    let old_work = varint_len(u64::from(field.number << 3 | u32::from(field.wire))) + old_span;
    match update {
        TransformField::Preserve => TransformDelta::default(),
        TransformField::Set(_) => {
            if replacement_len >= old_span {
                TransformDelta {
                    add_work: replacement_len - old_span,
                    ..TransformDelta::default()
                }
            } else {
                TransformDelta {
                    remove_work: old_span - replacement_len,
                    ..TransformDelta::default()
                }
            }
        },
        TransformField::Clear => TransformDelta {
            remove_fields: 1,
            remove_work: old_work,
            ..TransformDelta::default()
        },
    }
}

fn add_work_delta(delta: &mut TransformDelta, original: usize, replacement: usize) {
    if replacement >= original {
        delta.add_work = delta
            .add_work
            .saturating_add(replacement.saturating_sub(original));
    } else {
        delta.remove_work = delta
            .remove_work
            .saturating_add(original.saturating_sub(replacement));
    }
}

fn apply_delta(
    current: usize,
    added: usize,
    removed: usize,
    limit: DecodeLimit,
) -> Result<usize, DecodeError> {
    current
        .checked_add(added)
        .and_then(|value| value.checked_sub(removed))
        .ok_or_else(|| DecodeError::limited(limit))
}

fn length_field_len(number: u32, payload: usize) -> usize {
    varint_len(u64::from(number << 3 | 2)) + varint_len(payload as u64) + payload
}

fn encoded_fixed_pair_len(
    fields: &[ParsedField],
    first: u32,
    second: u32,
) -> Result<usize, DecodeError> {
    let first_field = unique_known(fields, first, "fixed32 field")?;
    let second_field = unique_known(fields, second, "fixed32 field")?;
    Ok(fields
        .iter()
        .map(|field| field.end - field.start)
        .sum::<usize>()
        .saturating_sub(first_field.end - first_field.start)
        .saturating_sub(second_field.end - second_field.start)
        .saturating_add(5)
        .saturating_add(5))
}

fn emit_rewrite(
    source: &[u8],
    write: MovieGeometryWrite,
    expected: usize,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    let size_field = unique_known(&geometry_fields, GEOMETRY_SIZE_FIELD, "Geometry.size")?;
    let point = &geometry[position_field.value_start..position_field.value_end];
    let size = &geometry[size_field.value_start..size_field.value_end];
    let rewritten_point = rewrite_fixed_message(point, write.position, true, options)?;
    let rewritten_size = rewrite_fixed_message(size, write.size, false, options)?;
    let mut new_geometry = Vec::new();
    new_geometry
        .try_reserve_exact(geometry.len())
        .map_err(|_| DecodeError::plain("geometry staging allocation failed"))?;
    for field in geometry_fields.iter().copied() {
        match field.number {
            GEOMETRY_POSITION_FIELD => {
                append_length_replacement(&mut new_geometry, geometry, field, &rewritten_point)
            },
            GEOMETRY_SIZE_FIELD => {
                append_length_replacement(&mut new_geometry, geometry, field, &rewritten_size)
            },
            _ => new_geometry.extend_from_slice(&geometry[field.start..field.end]),
        }
    }
    let mut new_drawable = Vec::new();
    new_drawable
        .try_reserve_exact(drawable.len())
        .map_err(|_| DecodeError::plain("drawable staging allocation failed"))?;
    for field in drawable_fields.iter().copied() {
        if field.number == DRAWABLE_GEOMETRY_FIELD {
            append_length_replacement(&mut new_drawable, drawable, field, &new_geometry);
        } else {
            new_drawable.extend_from_slice(&drawable[field.start..field.end]);
        }
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::plain("geometry output allocation failed"))?;
    for field in root.iter().copied() {
        if field.number == MOVIE_SUPER_FIELD {
            append_length_replacement(&mut output, source, field, &new_drawable);
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    if output.len() != expected {
        return Err(DecodeError::plain(
            "geometry rewrite output sizing disagreed",
        ));
    }
    Ok(output)
}

fn emit_position_rewrite(
    source: &[u8],
    write: MoviePositionWrite,
    expected: usize,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    require_wire(position_field, 2)?;
    if let Some(size_field) = optional_known(&geometry_fields, GEOMETRY_SIZE_FIELD)? {
        require_wire(size_field, 2)?;
    }
    let _ = parse_transform(&geometry_fields, geometry)?;
    let point = &geometry[position_field.value_start..position_field.value_end];
    let rewritten_point = rewrite_fixed_message(point, write.position, true, options)?;
    let geometry_len = replace_length(
        geometry.len(),
        field_span(position_field),
        length_field_len(GEOMETRY_POSITION_FIELD, rewritten_point.len()),
    )?;
    let mut new_geometry = Vec::new();
    new_geometry
        .try_reserve_exact(geometry_len)
        .map_err(|_| DecodeError::plain("movie position geometry allocation failed"))?;
    for field in geometry_fields.iter().copied() {
        if field.number == GEOMETRY_POSITION_FIELD {
            append_length_replacement(&mut new_geometry, geometry, field, &rewritten_point);
        } else {
            new_geometry.extend_from_slice(&geometry[field.start..field.end]);
        }
    }
    if new_geometry.len() != geometry_len {
        return Err(DecodeError::plain(
            "movie position geometry sizing disagreed",
        ));
    }
    let drawable_len = replace_length(
        drawable.len(),
        field_span(geometry_field),
        length_field_len(DRAWABLE_GEOMETRY_FIELD, new_geometry.len()),
    )?;
    let mut new_drawable = Vec::new();
    new_drawable
        .try_reserve_exact(drawable_len)
        .map_err(|_| DecodeError::plain("movie position drawable allocation failed"))?;
    for field in drawable_fields.iter().copied() {
        if field.number == DRAWABLE_GEOMETRY_FIELD {
            append_length_replacement(&mut new_drawable, drawable, field, &new_geometry);
        } else {
            new_drawable.extend_from_slice(&drawable[field.start..field.end]);
        }
    }
    if new_drawable.len() != drawable_len {
        return Err(DecodeError::plain(
            "movie position drawable sizing disagreed",
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::plain("movie position output allocation failed"))?;
    for field in root.iter().copied() {
        if field.number == MOVIE_SUPER_FIELD {
            append_length_replacement(&mut output, source, field, &new_drawable);
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    if output.len() != expected {
        return Err(DecodeError::plain("movie position output sizing disagreed"));
    }
    Ok(output)
}

fn emit_transform_rewrite(
    source: &[u8],
    write: MovieTransformWrite,
    expected: usize,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let mut budget = Budget::default();
    let root = parse_message(source, options, &mut budget, 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut budget, 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut budget, 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    let size_field = unique_known(&geometry_fields, GEOMETRY_SIZE_FIELD, "Geometry.size")?;
    require_wire(position_field, 2)?;
    require_wire(size_field, 2)?;
    let _ = parse_transform(&geometry_fields, geometry)?;
    let flags_field = optional_known(&geometry_fields, GEOMETRY_FLAGS_FIELD)?;
    let angle_field = optional_known(&geometry_fields, GEOMETRY_ANGLE_FIELD)?;
    let geometry_len = replace_optional_fields_len(
        geometry.len(),
        flags_field,
        write.flags,
        angle_field,
        write.angle_degrees,
    )?;
    let mut new_geometry = Vec::new();
    new_geometry
        .try_reserve_exact(geometry_len)
        .map_err(|_| DecodeError::plain("movie transform geometry allocation failed"))?;
    let mut saw_flags = false;
    let mut saw_angle = false;
    for field in geometry_fields.iter().copied() {
        match field.number {
            GEOMETRY_FLAGS_FIELD => {
                saw_flags = true;
                append_flags_replacement(&mut new_geometry, geometry, field, write.flags);
            },
            GEOMETRY_ANGLE_FIELD => {
                saw_angle = true;
                append_angle_replacement(&mut new_geometry, geometry, field, write.angle_degrees);
            },
            _ => new_geometry.extend_from_slice(&geometry[field.start..field.end]),
        }
    }
    if !saw_flags {
        append_missing_transform(&mut new_geometry, GEOMETRY_FLAGS_FIELD, write.flags);
    }
    if !saw_angle {
        append_missing_transform(&mut new_geometry, GEOMETRY_ANGLE_FIELD, write.angle_degrees);
    }
    if new_geometry.len() != geometry_len {
        return Err(DecodeError::plain(
            "movie transform geometry sizing disagreed",
        ));
    }
    let new_drawable_len = replace_length(
        drawable.len(),
        geometry_field.end - geometry_field.start,
        length_field_len(DRAWABLE_GEOMETRY_FIELD, new_geometry.len()),
    )?;
    let mut new_drawable = Vec::new();
    new_drawable
        .try_reserve_exact(new_drawable_len)
        .map_err(|_| DecodeError::plain("movie transform drawable allocation failed"))?;
    for field in drawable_fields.iter().copied() {
        if field.number == DRAWABLE_GEOMETRY_FIELD {
            append_length_replacement(&mut new_drawable, drawable, field, &new_geometry);
        } else {
            new_drawable.extend_from_slice(&drawable[field.start..field.end]);
        }
    }
    if new_drawable.len() != new_drawable_len {
        return Err(DecodeError::plain(
            "movie transform drawable sizing disagreed",
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::plain("movie transform output allocation failed"))?;
    for field in root.iter().copied() {
        if field.number == MOVIE_SUPER_FIELD {
            append_length_replacement(&mut output, source, field, &new_drawable);
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    if output.len() != expected {
        return Err(DecodeError::plain(
            "movie transform rewrite output sizing disagreed",
        ));
    }
    Ok(output)
}

fn append_flags_replacement(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    update: TransformField<u32>,
) {
    match update {
        TransformField::Preserve => {
            output.extend_from_slice(&source[field.start..field.end]);
        },
        TransformField::Set(value) => {
            append_varint(output, u64::from(GEOMETRY_FLAGS_FIELD << 3));
            append_varint(output, u64::from(value));
        },
        TransformField::Clear => {},
    }
}

fn append_angle_replacement(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    update: TransformField<f32>,
) {
    match update {
        TransformField::Preserve => {
            output.extend_from_slice(&source[field.start..field.end]);
        },
        TransformField::Set(value) => {
            append_varint(output, u64::from(GEOMETRY_ANGLE_FIELD << 3 | 5));
            output.extend_from_slice(&value.to_bits().to_le_bytes());
        },
        TransformField::Clear => {},
    }
}

fn append_missing_transform<T>(output: &mut Vec<u8>, number: u32, update: TransformField<T>)
where
    T: TransformScalar,
{
    if let TransformField::Set(value) = update {
        append_varint(output, u64::from(number << 3 | T::WIRE));
        T::append(output, value);
    }
}

trait TransformScalar: Copy {
    const WIRE: u32;

    fn append(output: &mut Vec<u8>, value: Self);
}

impl TransformScalar for u32 {
    const WIRE: u32 = 0;

    fn append(output: &mut Vec<u8>, value: Self) {
        append_varint(output, u64::from(value));
    }
}

impl TransformScalar for f32 {
    const WIRE: u32 = 5;

    fn append(output: &mut Vec<u8>, value: Self) {
        output.extend_from_slice(&value.to_bits().to_le_bytes());
    }
}

fn rewrite_fixed_message<T: FixedPair + Copy>(
    source: &[u8],
    value: T,
    point: bool,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let fields = parse_message(source, options, &mut Budget::default(), 4)?;
    let first = if point {
        POINT_X_FIELD
    } else {
        SIZE_WIDTH_FIELD
    };
    let second = if point {
        POINT_Y_FIELD
    } else {
        SIZE_HEIGHT_FIELD
    };
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| DecodeError::plain("fixed geometry staging allocation failed"))?;
    for field in fields {
        if field.number == first {
            append_fixed_replacement(&mut output, source, field, value.first());
        } else if field.number == second {
            append_fixed_replacement(&mut output, source, field, value.second());
        } else {
            output.extend_from_slice(&source[field.start..field.end]);
        }
    }
    Ok(output)
}

trait FixedPair {
    fn first(self) -> f32;
    fn second(self) -> f32;
}

impl FixedPair for Point {
    fn first(self) -> f32 {
        self.x
    }
    fn second(self) -> f32 {
        self.y
    }
}

impl FixedPair for Size {
    fn first(self) -> f32 {
        self.width
    }
    fn second(self) -> f32 {
        self.height
    }
}

fn append_length_replacement(
    output: &mut Vec<u8>,
    source: &[u8],
    field: ParsedField,
    payload: &[u8],
) {
    let key_end = field.start + varint_len(u64::from(field.number << 3 | u32::from(field.wire)));
    output.extend_from_slice(&source[field.start..key_end]);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn append_fixed_replacement(output: &mut Vec<u8>, source: &[u8], field: ParsedField, value: f32) {
    output.extend_from_slice(&source[field.start..field.value_start]);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
}

fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn check_limits(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    if requirements.output_bytes > limits.output_bytes {
        return Err(DecodeError::limited(DecodeLimit::Output {
            observed: requirements.output_bytes,
            maximum: limits.output_bytes,
        }));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: limits.fields,
        }));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: limits.work_bytes,
        }));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::limited(DecodeLimit::Allocations {
            observed: requirements.allocations,
            maximum: limits.allocations,
        }));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::limited(DecodeLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: limits.retained_bytes,
        }));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::limited(DecodeLimit::Scratch {
            observed: requirements.scratch_bytes,
            maximum: limits.scratch_bytes,
        }));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source() -> Vec<u8> {
        source_with_transform(true)
    }

    fn source_without_transform() -> Vec<u8> {
        source_with_transform(false)
    }

    fn source_with_transform(include_transform: bool) -> Vec<u8> {
        source_with_optional_transform(
            include_transform.then_some(7),
            include_transform.then_some(1.0),
        )
    }

    fn source_with_optional_transform(flags: Option<u32>, angle: Option<f32>) -> Vec<u8> {
        fn varint(mut value: usize) -> Vec<u8> {
            let mut bytes = Vec::new();
            while value >= 0x80 {
                bytes.push((value as u8) | 0x80);
                value >>= 7;
            }
            bytes.push(value as u8);
            bytes
        }
        fn length(number: u8, payload: &[u8]) -> Vec<u8> {
            let mut bytes = vec![number << 3 | 2];
            bytes.extend(varint(payload.len()));
            bytes.extend_from_slice(payload);
            bytes
        }
        fn fixed(number: u8, value: f32) -> Vec<u8> {
            let mut bytes = vec![number << 3 | 5];
            bytes.extend_from_slice(&value.to_le_bytes());
            bytes
        }
        let mut point = fixed(1, 0.0);
        point.extend(fixed(2, 0.0));
        let mut size = fixed(1, 640.0);
        size.extend(fixed(2, 480.0));
        let mut geometry = length(1, &point);
        geometry.extend(length(2, &size));
        if let Some(flags) = flags {
            geometry.extend([0x18]);
            geometry.extend(varint(flags as usize));
        }
        if let Some(angle) = angle {
            geometry.extend(fixed(4, angle));
        }
        let drawable = length(1, &geometry);
        length(1, &drawable)
    }

    fn position_source(size: Option<(f32, f32)>) -> Vec<u8> {
        position_source_with_fields(size, true, true, true)
    }

    fn position_source_with_fields(
        size: Option<(f32, f32)>,
        include_position: bool,
        include_x: bool,
        include_y: bool,
    ) -> Vec<u8> {
        fn varint(mut value: usize) -> Vec<u8> {
            let mut bytes = Vec::new();
            while value >= 0x80 {
                bytes.push((value as u8) | 0x80);
                value >>= 7;
            }
            bytes.push(value as u8);
            bytes
        }
        fn length(number: u8, payload: &[u8]) -> Vec<u8> {
            let mut bytes = vec![number << 3 | 2];
            bytes.extend(varint(payload.len()));
            bytes.extend_from_slice(payload);
            bytes
        }
        fn fixed(number: u8, value: f32) -> Vec<u8> {
            let mut bytes = vec![number << 3 | 5];
            bytes.extend_from_slice(&value.to_le_bytes());
            bytes
        }

        let mut point = Vec::new();
        if include_x {
            point.extend(fixed(1, 3.25));
        }
        if include_y {
            point.extend(fixed(2, -4.5));
        }
        point.extend([0x98, 0x06, 0x01]);
        let mut geometry = Vec::new();
        if include_position {
            geometry.extend(length(1, &point));
        }
        if let Some((width, height)) = size {
            let mut size = fixed(1, width);
            size.extend(fixed(2, height));
            geometry.extend(length(2, &size));
        }
        geometry.extend([0x18, 0x07]);
        geometry.extend([0x25, 0x00, 0x00, 0x80, 0x3f]);
        geometry.extend([0x98, 0x06, 0x02]);
        let mut drawable = length(1, &geometry);
        drawable.extend([0x10, 0x09]);
        let mut source = length(1, &drawable);
        source.extend([0x18, 0x01]);
        source
    }

    #[test]
    fn transform_projection_preserves_optional_presence_and_values() {
        let bytes = source();
        let transform = decode_movie_transform(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(transform.flags(), Some(7));
        assert_eq!(transform.angle_degrees(), Some(1.0));
        assert_eq!(transform.angle(), Some(1.0));

        let absent = source_without_transform();
        assert_eq!(
            decode_movie_transform(&absent, DecodeOptions::for_source(&absent)).unwrap(),
            MovieTransformSnapshot {
                flags: None,
                angle_degrees: None,
            }
        );
    }

    #[test]
    fn transform_rewrite_updates_only_known_scalars_and_preserves_unknowns() {
        let mut bytes = source();
        let unknown = [
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0,
        ];
        bytes.extend_from_slice(&unknown);
        let output = rewrite_movie_transform(
            &bytes,
            MovieTransformWrite::with_updates(TransformField::set(9), TransformField::set(225.5)),
            DecodeOptions::for_source(&bytes),
        )
        .unwrap();
        assert!(output.ends_with(&unknown));
        assert_eq!(
            decode_movie_transform(&output, DecodeOptions::for_source(&output)).unwrap(),
            MovieTransformSnapshot {
                flags: Some(9),
                angle_degrees: Some(225.5),
            }
        );
        assert_eq!(
            decode_movie_geometry(&bytes, DecodeOptions::for_source(&bytes)).unwrap(),
            decode_movie_geometry(&output, DecodeOptions::for_source(&output)).unwrap()
        );
    }

    #[test]
    fn transform_rewrite_can_clear_or_append_optional_fields() {
        let bytes = source();
        let cleared = rewrite_movie_transform(
            &bytes,
            MovieTransformWrite::with_updates(TransformField::clear(), TransformField::clear()),
            DecodeOptions::for_source(&bytes),
        )
        .unwrap();
        assert_eq!(
            decode_movie_transform(&cleared, DecodeOptions::for_source(&cleared)).unwrap(),
            MovieTransformSnapshot {
                flags: None,
                angle_degrees: None,
            }
        );

        let absent = source_without_transform();
        let appended = rewrite_movie_transform(
            &absent,
            MovieTransformWrite::new(Some(4), Some(90.0)),
            DecodeOptions::for_source(&absent),
        )
        .unwrap();
        assert_eq!(
            decode_movie_transform(&appended, DecodeOptions::for_source(&appended)).unwrap(),
            MovieTransformSnapshot {
                flags: Some(4),
                angle_degrees: Some(90.0),
            }
        );
    }

    #[test]
    fn transform_rewrite_replays_candidate_accounting_for_optional_presence() {
        for (source_flags, source_angle) in [
            (None, None),
            (Some(7), None),
            (None, Some(1.0)),
            (Some(7), Some(1.0)),
        ] {
            let bytes = source_with_optional_transform(source_flags, source_angle);
            let options = DecodeOptions::for_source(&bytes);
            for flags in [
                TransformField::Preserve,
                TransformField::Set(0x24),
                TransformField::Clear,
            ] {
                for angle in [
                    TransformField::Preserve,
                    TransformField::Set(180.0),
                    TransformField::Clear,
                ] {
                    let write = MovieTransformWrite::with_updates(flags, angle);
                    let prepared = prepare_movie_transform_rewrite(&bytes, write, options).unwrap();
                    let requirements = prepared.execution_requirements();
                    let source_report = prepared.prepare_report();
                    let output = prepared.execute(requirements.exact()).unwrap();
                    let (_, candidate_report) = decode_movie_transform_with_report(
                        output.output(),
                        DecodeOptions::for_source(output.output()),
                    )
                    .unwrap();
                    assert_eq!(
                        candidate_report.fields(),
                        requirements.fields - source_report.fields()
                    );
                    assert!(
                        requirements.work_bytes
                            >= source_report
                                .work_bytes()
                                .saturating_add(candidate_report.work_bytes())
                    );
                    assert!(
                        requirements.allocations
                            >= source_report
                                .allocations()
                                .saturating_add(candidate_report.allocations())
                    );
                    assert!(requirements.retained_bytes >= output.output().len());
                    assert!(
                        requirements.scratch_bytes
                            >= source_report
                                .scratch_bytes()
                                .saturating_add(candidate_report.scratch_bytes())
                    );
                    assert_eq!(candidate_report.max_depth(), requirements.max_depth);
                }
            }
        }
    }

    #[test]
    fn transform_preserve_is_an_exact_noop_and_prepared_limits_are_replayed() {
        let bytes = source();
        let options = DecodeOptions::for_source(&bytes);
        assert_eq!(
            rewrite_movie_transform(&bytes, MovieTransformWrite::preserve(), options).unwrap(),
            bytes
        );
        let prepared = prepare_movie_transform_rewrite(
            &bytes,
            MovieTransformWrite::new(Some(11), Some(180.0)),
            options,
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        assert!(requirements.allocations > BUFFA_LOGICAL_ALLOCATIONS);
        assert!(requirements.retained_bytes > requirements.output_bytes);
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_allocations(requirements.allocations.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_retained_bytes(requirements.retained_bytes.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1))
                )
                .is_err()
        );
        let output = prepared.execute(requirements.exact()).unwrap();
        assert!(output.report().changed());
    }

    #[test]
    fn transform_known_fields_reject_duplicate_noncanonical_wrong_wire_and_nonfinite() {
        let mut duplicate = source();
        let flags = duplicate
            .windows(2)
            .position(|window| window == [0x18, 0x07])
            .unwrap();
        duplicate.splice(flags..flags, [0x18, 0x08]);
        assert!(decode_movie_transform(&duplicate, DecodeOptions::for_source(&duplicate)).is_err());

        let mut noncanonical = source();
        let flags = noncanonical
            .windows(2)
            .position(|window| window == [0x18, 0x07])
            .unwrap();
        noncanonical.splice(flags..flags + 2, [0x18, 0x87, 0x00]);
        assert!(
            decode_movie_transform(&noncanonical, DecodeOptions::for_source(&noncanonical))
                .is_err()
        );

        let mut wrong_wire = source();
        let angle = wrong_wire.iter().position(|value| *value == 0x25).unwrap();
        wrong_wire[angle] = 0x20;
        assert!(
            decode_movie_transform(&wrong_wire, DecodeOptions::for_source(&wrong_wire)).is_err()
        );

        let mut nonfinite = source();
        let angle = nonfinite.iter().position(|value| *value == 0x25).unwrap();
        nonfinite[angle + 1..angle + 5].copy_from_slice(&f32::NAN.to_le_bytes());
        assert!(decode_movie_transform(&nonfinite, DecodeOptions::for_source(&nonfinite)).is_err());
    }

    #[test]
    fn reads_and_rewrites_geometry_without_losing_unknowns() {
        let mut bytes = source();
        bytes.extend_from_slice(&[
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0,
        ]);
        let snapshot = decode_movie_geometry(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(snapshot.size(), Size::new(640.0, 480.0));
        let output = rewrite_movie_geometry(
            &bytes,
            MovieGeometryWrite::from_values(10.0, -4.0, 800.0, 600.0),
            DecodeOptions::for_source(&bytes),
        )
        .unwrap();
        assert!(output.ends_with(&[
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0
        ]));
        assert_eq!(
            decode_movie_geometry(&output, DecodeOptions::for_source(&output)).unwrap(),
            MovieGeometrySnapshot {
                position: Point::new(10.0, -4.0),
                size: Size::new(800.0, 600.0)
            }
        );
    }

    #[test]
    fn balanced_unknown_group_is_preserved() {
        let mut bytes = source();
        bytes.extend_from_slice(&[0x53, 0x08, 0x01, 0x54]);
        let output = rewrite_movie_geometry(
            &bytes,
            MovieGeometryWrite::from_values(1.0, 2.0, 3.0, 4.0),
            DecodeOptions::for_source(&bytes),
        )
        .unwrap();
        assert!(output.ends_with(&[0x53, 0x08, 0x01, 0x54]));
    }

    #[test]
    fn prepared_limits_are_checked_before_execute() {
        let bytes = source();
        let prepared = prepare_movie_geometry_rewrite(
            &bytes,
            MovieGeometryWrite::from_values(1.0, 2.0, 3.0, 4.0),
            DecodeOptions::for_source(&bytes),
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        assert!(requirements.allocations > BUFFA_LOGICAL_ALLOCATIONS);
        assert!(requirements.retained_bytes > requirements.output_bytes);
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_allocations(requirements.allocations.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_retained_bytes(requirements.retained_bytes.saturating_sub(1))
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1))
                )
                .is_err()
        );
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.report().output_bytes(), requirements.output_bytes);
    }

    #[test]
    fn rejects_partial_nonfinite_and_nonpositive_values() {
        let mut missing = source();
        missing[6] = 0x08;
        assert!(decode_movie_geometry(&missing, DecodeOptions::for_source(&missing)).is_err());
        let mut bad_size = source();
        bad_size[20..24].copy_from_slice(&0f32.to_le_bytes());
        assert!(decode_movie_geometry(&bad_size, DecodeOptions::for_source(&bad_size)).is_err());
        let nan = MovieGeometryWrite::from_values(0.0, 0.0, f32::NAN, 1.0);
        let bytes = source();
        assert!(
            prepare_movie_geometry_rewrite(&bytes, nan, DecodeOptions::for_source(&bytes)).is_err()
        );
    }

    #[test]
    fn position_read_and_rewrite_preserve_absent_or_zero_size_and_all_other_spans() {
        for size in [None, Some((0.0, 0.0))] {
            let bytes = position_source(size);
            let options = DecodeOptions::for_source(&bytes);
            assert_eq!(
                decode_movie_position(&bytes, options).unwrap(),
                Point::new(3.25, -4.5)
            );
            let output = rewrite_movie_position(
                &bytes,
                MoviePositionWrite::from_values(18.5, -22.25),
                options,
            )
            .unwrap();
            let mut expected = bytes.clone();
            let old_x = 3.25f32.to_le_bytes();
            let old_y = (-4.5f32).to_le_bytes();
            let new_x = 18.5f32.to_le_bytes();
            let new_y = (-22.25f32).to_le_bytes();
            let x = expected
                .windows(5)
                .position(|window| window == [0x0d, old_x[0], old_x[1], old_x[2], old_x[3]])
                .unwrap();
            expected[x + 1..x + 5].copy_from_slice(&new_x);
            let y = expected
                .windows(5)
                .position(|window| window == [0x15, old_y[0], old_y[1], old_y[2], old_y[3]])
                .unwrap();
            expected[y + 1..y + 5].copy_from_slice(&new_y);
            assert_eq!(output, expected);
            assert_eq!(
                decode_movie_position(&output, DecodeOptions::for_source(&output)).unwrap(),
                Point::new(18.5, -22.25)
            );
        }
    }

    #[test]
    fn position_preparation_replays_complete_limits_and_exact_noop() {
        let bytes = position_source(None);
        let options = DecodeOptions::for_source(&bytes);
        let prepared = prepare_movie_position_rewrite(
            &bytes,
            MoviePositionWrite::from_values(18.5, -22.25),
            options,
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).unwrap();
        assert!(output.report().changed());
        assert_eq!(output.report().output_bytes(), bytes.len());
        assert!(requirements.allocations > BUFFA_LOGICAL_ALLOCATIONS);
        assert!(requirements.retained_bytes > requirements.output_bytes);
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes.saturating_sub(1)),
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_fields(requirements.fields.saturating_sub(1)),
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_work_bytes(requirements.work_bytes.saturating_sub(1)),
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_max_depth(requirements.max_depth.saturating_sub(1)),
                )
                .is_err()
        );
        assert!(
            prepare_movie_position_rewrite(
                &bytes,
                MoviePositionWrite::from_values(f32::NAN, 0.0),
                options,
            )
            .is_err()
        );
        assert!(
            prepare_movie_position_rewrite(
                &bytes,
                MoviePositionWrite::from_values(18.5, -22.25),
                options,
            )
            .unwrap()
            .execute(
                requirements
                    .exact()
                    .with_allocations(requirements.allocations.saturating_sub(1)),
            )
            .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_retained_bytes(requirements.retained_bytes.saturating_sub(1)),
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1)),
                )
                .is_err()
        );

        let noop =
            rewrite_movie_position(&bytes, MoviePositionWrite::from_values(3.25, -4.5), options)
                .unwrap();
        assert_eq!(noop, bytes);
        let prepared_noop = prepare_movie_position_rewrite(
            &bytes,
            MoviePositionWrite::from_values(3.25, -4.5),
            options,
        )
        .unwrap();
        let noop_output = prepared_noop
            .execute(prepared_noop.execution_requirements().exact())
            .unwrap();
        assert!(!noop_output.report().changed());
        assert_eq!(noop_output.output(), bytes.as_slice());
    }

    #[test]
    fn position_rewrite_rejects_missing_position_without_mutating_source() {
        for (include_position, include_x, include_y) in [
            (false, true, true),
            (true, false, true),
            (true, true, false),
        ] {
            let missing = position_source_with_fields(
                Some((0.0, 0.0)),
                include_position,
                include_x,
                include_y,
            );
            let before = missing.clone();
            let options = DecodeOptions::for_source(&missing);
            assert!(decode_movie_position(&missing, options).is_err());
            assert!(
                rewrite_movie_position(
                    &missing,
                    MoviePositionWrite::from_values(1.0, 2.0),
                    options,
                )
                .is_err()
            );
            assert_eq!(missing, before);
        }
    }
}
