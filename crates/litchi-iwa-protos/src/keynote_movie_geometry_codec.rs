//! Strict, generated-free projection and source-preserving rewrite for the
//! Keynote `TSD.MovieArchive` geometry edge.
//!
//! Only `MovieArchive.super.geometry.position` and `.size` are semantic write
//! targets.  Flags, angle, envelope framing, and unknown fields remain raw
//! source bytes.  A private Buffa lazy view is used as an ingress cross-check;
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
        .with_max_retained_bytes(bytes.saturating_mul(2).max(1))
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
    if scan.scratch_bytes > options.max_scratch_bytes {
        return Err(DecodeError::limited(DecodeLimit::Scratch {
            observed: scan.scratch_bytes,
            maximum: options.max_scratch_bytes,
        }));
    }
    Ok((
        snapshot,
        DecodeReport {
            input_bytes: source.len(),
            fields: scan.fields,
            work_bytes: scan.work,
            max_depth: scan.max_depth,
            allocations: scan.allocations,
            retained_bytes: source.len(),
            scratch_bytes: scan.scratch_bytes,
        },
    ))
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
    let requirements = RewriteExecutionRequirements {
        output_bytes: output_measure.bytes,
        fields: source_scan
            .fields
            .checked_add(output_measure.fields)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Fields {
                    observed: usize::MAX,
                    maximum: options.max_fields,
                })
            })?,
        work_bytes: source_scan
            .work
            .checked_add(output_measure.work)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Work {
                    observed: usize::MAX,
                    maximum: options.max_work_bytes,
                })
            })?,
        max_depth: output_measure.max_depth,
        allocations: 3,
        retained_bytes: output_measure.bytes,
        scratch_bytes: source_scan
            .scratch_bytes
            .checked_add(output_measure.scratch_bytes)
            .ok_or_else(|| {
                DecodeError::limited(DecodeLimit::Scratch {
                    observed: usize::MAX,
                    maximum: options.max_scratch_bytes,
                })
            })?,
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
    Ok(PreparedMovieGeometryRewrite {
        source,
        write,
        options,
        requirements,
        source_report: DecodeReport {
            input_bytes: source.len(),
            fields: source_scan.fields,
            work_bytes: source_scan.work,
            max_depth: source_scan.max_depth,
            allocations: source_scan.allocations,
            retained_bytes: source.len(),
            scratch_bytes: source_scan.scratch_bytes,
        },
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
    scratch_bytes: usize,
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
    if !write.position.x.is_finite() || !write.position.y.is_finite() {
        return Err(DecodeError::plain("movie geometry position must be finite"));
    }
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
    validate_geometry_known(&geometry_fields, geometry)?;
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
        ScanAccounting {
            fields: budget.fields,
            work: budget.work,
            max_depth: budget.max_depth,
            allocations: 3,
            scratch_bytes: source
                .len()
                .min(options.max_fields)
                .saturating_mul(size_of::<ParsedField>())
                .saturating_mul(5),
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

fn validate_geometry_known(fields: &[ParsedField], source: &[u8]) -> Result<(), DecodeError> {
    for field in fields {
        match field.number {
            GEOMETRY_FLAGS_FIELD => {
                require_wire(field, 0)?;
                let value = field
                    .value
                    .ok_or_else(|| DecodeError::plain("invalid geometry flags"))?;
                if varint_len(value) != field.value_end.saturating_sub(field.value_start)
                    || value > u64::from(u32::MAX)
                {
                    return Err(DecodeError::plain("noncanonical geometry flags"));
                }
            },
            GEOMETRY_ANGLE_FIELD => {
                require_wire(field, 5)?;
                let value = fixed32(source, field)?;
                if !value.is_finite() {
                    return Err(DecodeError::plain("geometry angle must be finite"));
                }
            },
            _ => {},
        }
    }
    Ok(())
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
    if scratch > options.max_scratch_bytes {
        return Err(DecodeError::limited(DecodeLimit::Scratch {
            observed: scratch,
            maximum: options.max_scratch_bytes,
        }));
    }
    let mut fields = Vec::new();
    fields.try_reserve_exact(capacity).map_err(|_| {
        DecodeError::limited(DecodeLimit::Scratch {
            observed: scratch,
            maximum: options.max_scratch_bytes,
        })
    })?;
    let mut offset = 0;
    while offset < source.len() {
        let (field, next) = parse_field(source, offset, source.len(), depth, options, budget)?;
        fields.push(field);
        offset = next;
    }
    Ok(fields)
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
    let root = parse_message(source, options, &mut Budget::default(), 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut Budget::default(), 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut Budget::default(), 3)?;
    let position_field = unique_known(
        &geometry_fields,
        GEOMETRY_POSITION_FIELD,
        "Geometry.position",
    )?;
    let size_field = unique_known(&geometry_fields, GEOMETRY_SIZE_FIELD, "Geometry.size")?;
    let point = &geometry[position_field.value_start..position_field.value_end];
    let size = &geometry[size_field.value_start..size_field.value_end];
    let point_fields = parse_message(point, options, &mut Budget::default(), 4)?;
    let size_fields = parse_message(size, options, &mut Budget::default(), 4)?;
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
    let candidate_work = source_scan.work;
    let scratch = output_len
        .min(options.max_fields)
        .saturating_mul(size_of::<ParsedField>())
        .saturating_mul(5);
    Ok(OutputMeasure {
        bytes: output_len,
        fields: candidate_fields,
        work: candidate_work,
        max_depth: source_scan.max_depth,
        scratch_bytes: scratch,
    })
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
    let root = parse_message(source, options, &mut Budget::default(), 1)?;
    let super_field = unique_known(&root, MOVIE_SUPER_FIELD, "MovieArchive.super")?;
    let drawable = &source[super_field.value_start..super_field.value_end];
    let drawable_fields = parse_message(drawable, options, &mut Budget::default(), 2)?;
    let geometry_field = unique_known(
        &drawable_fields,
        DRAWABLE_GEOMETRY_FIELD,
        "DrawableArchive.geometry",
    )?;
    let geometry = &drawable[geometry_field.value_start..geometry_field.value_end];
    let geometry_fields = parse_message(geometry, options, &mut Budget::default(), 3)?;
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
    let mut new_geometry = Vec::with_capacity(geometry.len());
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
    let mut new_drawable = Vec::with_capacity(drawable.len());
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
    let mut output = Vec::with_capacity(source.len());
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
        geometry.extend([0x18, 0x07]);
        geometry.extend(fixed(4, 1.0));
        let drawable = length(1, &geometry);
        length(1, &drawable)
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
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes.saturating_sub(1))
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
}
