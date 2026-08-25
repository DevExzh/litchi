//! Strict, generated-free projection and source-preserving rewrite for the
//! scalar playback settings of `TSD.MovieArchive`.
//!
//! The complete movie graph is deliberately outside this module.  The raw
//! payload remains the preservation authority; the private Buffa sidecar is
//! used only as a bounded lazy ingress cross-check after the strict scanner.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict scanner is intentionally ordered before the private projection."
)]

use std::{fmt, mem::size_of};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_movie_playback_generated::LitchiIwaProjection as projection;

const SUPER_FIELD: u32 = 1;
const START_FIELD: u32 = 3;
const END_FIELD: u32 = 4;
const POSTER_FIELD: u32 = 5;
const LEGACY_LOOP_FIELD: u32 = 6;
const VOLUME_FIELD: u32 = 7;
const MODERN_LOOP_FIELD: u32 = 24;
const MAX_RECURSION: u32 = 64;

/// Finite limits for one source payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Construct a bounded payload policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_input_bytes,
            max_output_bytes: max_input_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
        }
    }

    /// Construct a conservative policy from one source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(bytes, bytes.saturating_mul(4), bytes.saturating_mul(8), 8)
    }

    /// Override the source/input ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, value: usize) -> Self {
        self.max_input_bytes = value;
        self
    }

    /// Override the candidate-output ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, value: usize) -> Self {
        self.max_output_bytes = value;
        self
    }

    /// Override the field/work ceilings.
    #[must_use]
    pub const fn with_resource_limits(mut self, fields: usize, work_bytes: usize) -> Self {
        self.max_fields = fields;
        self.max_work_bytes = work_bytes;
        self
    }

    /// Override the nesting ceiling.
    #[must_use]
    pub const fn with_recursion_limit(mut self, value: u32) -> Self {
        self.recursion_limit = value;
        self
    }

    /// Return the configured input ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Return the configured output ceiling.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            .with_unknown_field_limit(self.max_input_bytes)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Resource axis that rejected a decode or prepared execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source bytes exceeded the input ceiling.
    InputBytes,
    /// Candidate bytes exceeded the output ceiling.
    OutputBytes,
    /// Strict field visits exceeded the field ceiling.
    Fields,
    /// Strict work exceeded the work ceiling.
    Work,
    /// Protobuf nesting exceeded the nesting ceiling.
    Nesting,
    /// Candidate allocation count exceeded its ceiling.
    Allocations,
    /// Candidate scratch exceeded its ceiling.
    Scratch,
    /// Candidate retained bytes exceeded its ceiling.
    Retained,
}

/// Strict codec failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(&'static str),
    Missing(&'static str),
    Duplicate(&'static str),
    WrongWire(&'static str),
    NonCanonical(&'static str),
    Invalid(&'static str),
    Limit {
        axis: DecodeLimit,
        observed: usize,
        maximum: usize,
    },
    Nesting {
        observed: u32,
        maximum: u32,
    },
    Allocation(usize),
    Projection,
}

impl DecodeError {
    const fn missing(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Missing(field),
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

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn invalid(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Invalid(reason),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn wire(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(reason),
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation(amount),
        }
    }

    const fn limit(axis: DecodeLimit, observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Limit {
                axis,
                observed,
                maximum,
            },
        }
    }

    const fn nesting(observed: u32, maximum: u32) -> Self {
        Self {
            kind: DecodeErrorKind::Nesting { observed, maximum },
        }
    }

    /// Return the rejected resource axis, if any.
    #[must_use]
    pub const fn limit_kind(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit { axis, .. } => Some(axis),
            DecodeErrorKind::Nesting { .. } => Some(DecodeLimit::Nesting),
            _ => None,
        }
    }

    /// Return exact observed/maximum values for a byte/field/work failure.
    #[must_use]
    pub const fn limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit {
                observed, maximum, ..
            } => Some((observed, maximum)),
            _ => None,
        }
    }

    /// Return the structured resource observation used by package adapters.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<(DecodeLimit, usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit {
                axis,
                observed,
                maximum,
            } => Some((axis, observed, maximum)),
            DecodeErrorKind::Nesting { observed, maximum } => {
                Some((DecodeLimit::Nesting, observed as usize, maximum as usize))
            },
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.kind {
            DecodeErrorKind::Wire(reason) => write!(formatter, "invalid protobuf wire: {reason}"),
            DecodeErrorKind::Missing(field) => write!(formatter, "missing required field {field}"),
            DecodeErrorKind::Duplicate(field) => write!(formatter, "duplicate known field {field}"),
            DecodeErrorKind::WrongWire(field) => write!(formatter, "wrong wire type for {field}"),
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf: {reason}")
            },
            DecodeErrorKind::Invalid(reason) => formatter.write_str(reason),
            DecodeErrorKind::Limit {
                axis,
                observed,
                maximum,
            } => {
                write!(
                    formatter,
                    "movie playback {axis:?} limit exceeded: {observed} > {maximum}"
                )
            },
            DecodeErrorKind::Nesting { observed, maximum } => {
                write!(
                    formatter,
                    "movie playback Nesting limit exceeded: {observed} > {maximum}"
                )
            },
            DecodeErrorKind::Allocation(bytes) => {
                write!(
                    formatter,
                    "movie playback output allocation failed for {bytes} bytes"
                )
            },
            DecodeErrorKind::Projection => {
                formatter.write_str("Buffa playback projection disagrees with strict source scan")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

/// Borrowed semantic playback facts. Values are native finite seconds and
/// raw loop discriminants; package crates convert them to archive-free types.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MoviePlaybackSnapshot {
    /// Whether the required movie drawable envelope was present.
    pub has_super: bool,
    /// Optional native start time in seconds.
    pub start_time: Option<f32>,
    /// Required native end time in seconds.
    pub end_time: f32,
    /// Optional native poster time in seconds.
    pub poster_time: Option<f32>,
    /// Effective loop discriminant, preserving unknown future values.
    pub loop_mode: Option<i32>,
    /// Optional native volume multiplier.
    pub volume: Option<f32>,
    /// Whether legacy loop field 6 was present.
    pub has_legacy_loop: bool,
    /// Whether modern loop field 24 was present.
    pub has_modern_loop: bool,
}

impl MoviePlaybackSnapshot {
    /// Return the effective loop discriminant.
    #[must_use]
    pub const fn loop_mode(self) -> Option<i32> {
        self.loop_mode
    }
}

/// Requested playback scalar replacement. Optional values clear their native
/// fields when the source contains those fields, while required end time is
/// always emitted.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MoviePlaybackWrite {
    /// Optional native start time in seconds.
    pub start_time: Option<f32>,
    /// Required native end time in seconds.
    pub end_time: f32,
    /// Optional native poster time in seconds.
    pub poster_time: Option<f32>,
    /// Optional raw loop discriminant.
    pub loop_mode: Option<i32>,
    /// Optional native volume multiplier.
    pub volume: Option<f32>,
}

impl MoviePlaybackWrite {
    /// Construct a write with an explicit end time and omitted optionals.
    #[must_use]
    pub const fn new(end_time: f32) -> Self {
        Self {
            start_time: None,
            end_time,
            poster_time: None,
            loop_mode: None,
            volume: None,
        }
    }

    /// Construct all native values directly.
    #[must_use]
    pub const fn from_values(
        start_time: Option<f32>,
        end_time: f32,
        poster_time: Option<f32>,
        loop_mode: Option<i32>,
        volume: Option<f32>,
    ) -> Self {
        Self {
            start_time,
            end_time,
            poster_time,
            loop_mode,
            volume,
        }
    }
}

/// Exact source decode accounting.
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

/// Exact rewrite accounting returned by [`RewriteOutput`].
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

/// Prepared output and accounting requirements. Preparation performs strict
/// validation and sizing but never allocates the candidate output buffer.
///
/// Execution uses three logical allocations: the source field staging vector,
/// the candidate output buffer, and the candidate field staging vector. The
/// Buffa lazy view is borrowed and contributes no heap allocation. Scratch is
/// the exact reserved capacity of the two field vectors; the output buffer is
/// accounted separately as retained/output bytes.
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
    /// Return a matching exact execution policy.
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

    /// Alias used by package transaction owners.
    #[must_use]
    pub const fn exact_limits(self) -> RewriteExecutionLimits {
        self.exact()
    }
}

/// Limits replayed against a prepared rewrite before output allocation.
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
    /// Build exact limits from prepared requirements.
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

/// Candidate output plus exact rewrite accounting.
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
    pub fn into_bytes(self) -> Vec<u8> {
        self.output
    }
    #[must_use]
    pub const fn report(&self) -> RewriteReport {
        self.report
    }
}

/// Borrowed source rewrite prepared for exact execution-limit replay.
#[derive(Debug, Clone, Copy)]
pub struct PreparedMoviePlaybackRewrite<'source> {
    source: &'source [u8],
    write: MoviePlaybackWrite,
    options: DecodeOptions,
    report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    source_fields: usize,
    source_work_bytes: usize,
    candidate_fields: usize,
    candidate_work_bytes: usize,
}

impl<'source> PreparedMoviePlaybackRewrite<'source> {
    #[must_use]
    pub const fn prepare_report(self) -> DecodeReport {
        self.report
    }
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Execute only after all physical/resource ceilings pass.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_limits(self.requirements, limits)?;
        let (output, source_scan) = emit_rewrite(
            self.source,
            self.write,
            self.requirements.output_bytes,
            self.options,
        )?;
        if source_scan.fields != self.source_fields
            || source_scan.work_bytes != self.source_work_bytes
            || source_scan.max_depth != self.requirements.max_depth
        {
            return Err(DecodeError::projection());
        }
        let readback_options = DecodeOptions::new(
            output.len(),
            self.candidate_fields,
            self.candidate_work_bytes,
            self.requirements.max_depth,
        );
        let (readback, candidate_scan) = scan_source_with_accounting(&output, readback_options)?;
        if candidate_scan.fields != self.candidate_fields
            || candidate_scan.work_bytes != self.candidate_work_bytes
            || candidate_scan.max_depth != self.requirements.max_depth
        {
            return Err(DecodeError::projection());
        }
        force_buffa(
            &output,
            readback_options.with_recursion_limit(self.options.recursion_limit),
        )?;
        if !snapshot_matches_write(readback, self.write) {
            return Err(DecodeError::projection());
        }
        if candidate_scan.scratch_bytes
            != self
                .requirements
                .scratch_bytes
                .saturating_sub(source_scan.scratch_bytes)
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
            changed: output != self.source,
        };
        Ok(RewriteOutput { output, report })
    }
}

/// Strictly decode playback settings and exact accounting.
pub fn decode_movie_playback(
    source: &[u8],
    options: DecodeOptions,
) -> Result<MoviePlaybackSnapshot, DecodeError> {
    Ok(decode_movie_playback_with_report(source, options)?.0)
}

/// Strictly decode playback settings and return exact source accounting.
pub fn decode_movie_playback_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(MoviePlaybackSnapshot, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let (snapshot, _fields, accounting) = scan_source_parts(source, options)?;
    force_buffa(source, options)?;
    let report = DecodeReport {
        input_bytes: source.len(),
        fields: accounting.fields,
        work_bytes: accounting.work_bytes,
        max_depth: accounting.max_depth,
        allocations: accounting.allocations,
        retained_bytes: source.len(),
        scratch_bytes: accounting.scratch_bytes,
    };
    Ok((snapshot, report))
}

/// Prepare a strict playback rewrite without allocating candidate output.
pub fn prepare_movie_playback_rewrite<'source>(
    source: &'source [u8],
    write: MoviePlaybackWrite,
    options: DecodeOptions,
) -> Result<PreparedMoviePlaybackRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    let (_snapshot, source_fields, source_scan) = scan_source_parts(source, options)?;
    force_buffa(source, options)?;
    validate_write(write)?;
    let output_measure = measure_output(source, &source_fields, source_scan, write)?;
    if output_measure.bytes > options.max_output_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::OutputBytes,
            output_measure.bytes,
            options.max_output_bytes,
        ));
    }
    let requirements = RewriteExecutionRequirements {
        output_bytes: output_measure.bytes,
        fields: source_scan
            .fields
            .checked_add(output_measure.fields)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Fields, usize::MAX, options.max_fields)
            })?,
        work_bytes: source_scan
            .work_bytes
            .checked_add(output_measure.work)
            .ok_or_else(|| {
                DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes)
            })?,
        max_depth: source_scan.max_depth,
        allocations: 3,
        retained_bytes: output_measure.bytes,
        scratch_bytes: source_scan
            .scratch_bytes
            .checked_add(output_measure.scratch_bytes)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))?,
    };
    Ok(PreparedMoviePlaybackRewrite {
        source,
        write,
        options,
        report: DecodeReport {
            input_bytes: source.len(),
            fields: source_scan.fields,
            work_bytes: source_scan.work_bytes,
            max_depth: source_scan.max_depth,
            allocations: source_scan.allocations,
            retained_bytes: source.len(),
            scratch_bytes: source_scan.scratch_bytes,
        },
        requirements,
        source_fields: source_scan.fields,
        source_work_bytes: source_scan.work_bytes,
        candidate_fields: output_measure.fields,
        candidate_work_bytes: output_measure.work,
    })
}

/// Rewrite playback settings with exact prepared limits.
pub fn rewrite_movie_playback(
    source: &[u8],
    write: MoviePlaybackWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_movie_playback_rewrite(source, write, options)?;
    Ok(prepared
        .execute(prepared.execution_requirements().exact())?
        .into_output())
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    if source.len() > options.max_input_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::InputBytes,
            source.len(),
            options.max_input_bytes,
        ));
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
        return Err(DecodeError::nesting(options.recursion_limit, MAX_RECURSION));
    }
    Ok(())
}

fn validate_write(write: MoviePlaybackWrite) -> Result<(), DecodeError> {
    validate_time(write.start_time, "startTime")?;
    validate_time(Some(write.end_time), "endTime")?;
    validate_time(write.poster_time, "posterTime")?;
    if let Some(start) = write.start_time {
        if write.end_time <= start {
            return Err(DecodeError::invalid(
                "movie endTime must be later than startTime",
            ));
        }
    }
    if let Some(volume) = write.volume {
        if !volume.is_finite() || !(0.0..=1.0).contains(&volume) {
            return Err(DecodeError::invalid(
                "movie volume must be finite and in 0.0..=1.0",
            ));
        }
    }
    Ok(())
}

fn validate_time(value: Option<f32>, name: &'static str) -> Result<(), DecodeError> {
    if value.is_some_and(|x| !x.is_finite() || x < 0.0) {
        return Err(DecodeError::invalid(name));
    }
    Ok(())
}

fn force_buffa(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let view: projection::MoviePlaybackArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::projection())?;
    // Access every selected lazy scalar so Buffa validates the full sidecar
    // projection, while the strict scanner remains the semantic authority.
    let _ = (
        view.super_,
        view.start_time,
        view.end_time,
        view.poster_time,
        view.loop_option_as_integer,
        view.volume,
        view.loop_option,
    );
    Ok(())
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
    work_bytes: usize,
}

fn require_wire(field: ParsedField, wire: u8, name: &'static str) -> Result<(), DecodeError> {
    (field.wire == wire)
        .then_some(())
        .ok_or_else(|| DecodeError::wrong_wire(name))
}

fn fixed32(source: &[u8], field: ParsedField, name: &'static str) -> Result<f32, DecodeError> {
    let bytes: [u8; 4] = source[field.value_start..field.value_end]
        .try_into()
        .map_err(|_| DecodeError::wrong_wire(name))?;
    let value = f32::from_le_bytes(bytes);
    if !value.is_finite() || value < 0.0 {
        return Err(DecodeError::invalid(name));
    }
    Ok(value)
}

fn varint_known(field: ParsedField, name: &'static str) -> Result<u64, DecodeError> {
    let value = field.value.ok_or_else(|| DecodeError::wrong_wire(name))?;
    let len = field.value_end.saturating_sub(field.value_start);
    if varint_len(value) != len {
        return Err(DecodeError::noncanonical(name));
    }
    Ok(value)
}

fn legacy_loop_value(field: ParsedField) -> Result<i32, DecodeError> {
    let value = varint_known(field, "loopOptionAsInteger")?;
    u32::try_from(value)
        .map(|value| i32::from_le_bytes(value.to_le_bytes()))
        .map_err(|_| DecodeError::noncanonical("legacy loop value"))
}

fn modern_loop_value(field: ParsedField) -> Result<i32, DecodeError> {
    let value = varint_known(field, "loopOption")?;
    if value <= u64::from(u32::MAX) {
        return Ok(i32::from_le_bytes((value as u32).to_le_bytes()));
    }
    if field.value_end - field.value_start != 10 || (value >> 32) != u32::MAX as u64 {
        return Err(DecodeError::noncanonical("modern loop value"));
    }
    Ok(value as i64 as i32)
}

fn varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64 - value.leading_zeros()).div_ceil(7) as usize
    }
}

#[derive(Default)]
struct Budget {
    fields: usize,
    work: usize,
    max_depth: u32,
}

#[derive(Debug, Clone, Copy)]
struct OutputMeasure {
    bytes: usize,
    fields: usize,
    work: usize,
    scratch_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct ParseAccounting {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

fn parsed_fields_scratch(capacity: usize) -> Result<usize, DecodeError> {
    capacity
        .checked_mul(size_of::<ParsedField>())
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))
}

fn parse_root(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(Vec<ParsedField>, ParseAccounting), DecodeError> {
    let mut result = Vec::new();
    let capacity = source.len().min(options.max_fields);
    result.try_reserve(capacity).map_err(|_| {
        DecodeError::limit(DecodeLimit::Allocations, source.len(), options.max_fields)
    })?;
    let mut offset = 0;
    while offset < source.len() {
        let (field, next) = parse_field(source, offset, source.len(), 1, options, budget)?;
        result.push(field);
        offset = next;
    }
    Ok((
        result,
        ParseAccounting {
            fields: budget.fields,
            work_bytes: budget.work,
            max_depth: budget.max_depth,
            allocations: usize::from(capacity != 0),
            scratch_bytes: parsed_fields_scratch(capacity)?,
        },
    ))
}

fn scan_source_parts(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(MoviePlaybackSnapshot, Vec<ParsedField>, ParseAccounting), DecodeError> {
    let mut budget = Budget::default();
    let (fields, accounting) = parse_root(source, options, &mut budget)?;
    let mut snapshot = MoviePlaybackSnapshot {
        has_super: false,
        start_time: None,
        end_time: 0.0,
        poster_time: None,
        loop_mode: None,
        volume: None,
        has_legacy_loop: false,
        has_modern_loop: false,
    };
    let mut end_seen = false;
    let mut legacy = None;
    let mut modern = None;
    for field in fields.iter().copied() {
        match field.number {
            SUPER_FIELD => {
                require_wire(field, 2, "TSD.MovieArchive.super")?;
                if snapshot.has_super {
                    return Err(DecodeError::duplicate("TSD.MovieArchive.super"));
                }
                snapshot.has_super = true;
            },
            START_FIELD => {
                require_wire(field, 5, "startTime")?;
                if snapshot.start_time.is_some() {
                    return Err(DecodeError::duplicate("startTime"));
                }
                snapshot.start_time = Some(fixed32(source, field, "startTime")?);
            },
            END_FIELD => {
                require_wire(field, 5, "endTime")?;
                if end_seen {
                    return Err(DecodeError::duplicate("endTime"));
                }
                end_seen = true;
                snapshot.end_time = fixed32(source, field, "endTime")?;
            },
            POSTER_FIELD => {
                require_wire(field, 5, "posterTime")?;
                if snapshot.poster_time.is_some() {
                    return Err(DecodeError::duplicate("posterTime"));
                }
                snapshot.poster_time = Some(fixed32(source, field, "posterTime")?);
            },
            VOLUME_FIELD => {
                require_wire(field, 5, "volume")?;
                if snapshot.volume.is_some() {
                    return Err(DecodeError::duplicate("volume"));
                }
                let value = fixed32(source, field, "volume")?;
                if !(0.0..=1.0).contains(&value) {
                    return Err(DecodeError::invalid(
                        "movie volume must be finite and in 0.0..=1.0",
                    ));
                }
                snapshot.volume = Some(value);
            },
            LEGACY_LOOP_FIELD => {
                require_wire(field, 0, "loopOptionAsInteger")?;
                if legacy.is_some() {
                    return Err(DecodeError::duplicate("loopOptionAsInteger"));
                }
                legacy = Some(legacy_loop_value(field)?);
                snapshot.has_legacy_loop = true;
            },
            MODERN_LOOP_FIELD => {
                require_wire(field, 0, "loopOption")?;
                if modern.is_some() {
                    return Err(DecodeError::duplicate("loopOption"));
                }
                modern = Some(modern_loop_value(field)?);
                snapshot.has_modern_loop = true;
            },
            _ => {},
        }
    }
    if !snapshot.has_super {
        return Err(DecodeError::missing("TSD.MovieArchive.super"));
    }
    if !end_seen {
        return Err(DecodeError::missing("endTime"));
    }
    if snapshot.end_time <= 0.0 {
        return Err(DecodeError::invalid(
            "movie endTime must be finite and positive",
        ));
    }
    if let Some(start) = snapshot.start_time {
        if snapshot.end_time <= start {
            return Err(DecodeError::invalid(
                "movie endTime must be later than startTime",
            ));
        }
    }
    if let (Some(left), Some(right)) = (legacy, modern) {
        if left != right {
            return Err(DecodeError::invalid(
                "movie legacy and modern loop fields conflict",
            ));
        }
    }
    snapshot.loop_mode = modern.or(legacy);
    Ok((snapshot, fields, accounting))
}

fn scan_source_with_accounting(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(MoviePlaybackSnapshot, ParseAccounting), DecodeError> {
    let (snapshot, _fields, accounting) = scan_source_parts(source, options)?;
    Ok((snapshot, accounting))
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
        return Err(DecodeError::nesting(depth, options.recursion_limit));
    }
    budget.max_depth = budget.max_depth.max(depth);
    let (key, key_end) = read_varint(source, offset, end, false)?;
    let number = (key >> 3) as u32;
    let wire = (key & 7) as u8;
    if number == 0 {
        return Err(DecodeError::wire("field number zero"));
    }
    if is_known_field(number) && key_end - offset != varint_len(key) {
        return Err(DecodeError::noncanonical("known field key"));
    }
    budget.fields = budget
        .fields
        .checked_add(1)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, options.max_fields))?;
    if budget.fields > options.max_fields {
        return Err(DecodeError::limit(
            DecodeLimit::Fields,
            budget.fields,
            options.max_fields,
        ));
    }
    let (value_start, value_end, next, children_count, children_work) = match wire {
        0 => {
            let (_, value_end) = read_varint(source, key_end, end, false)?;
            (key_end, value_end, value_end, 0, 0)
        },
        1 => {
            let value_end = key_end
                .checked_add(8)
                .ok_or_else(|| DecodeError::wire("fixed64 overflow"))?;
            if value_end > end {
                return Err(DecodeError::wire("truncated fixed64"));
            }
            (key_end, value_end, value_end, 0, 0)
        },
        2 => {
            let (length, payload_start) = read_varint(source, key_end, end, false)?;
            if number == SUPER_FIELD && payload_start - key_end != varint_len(length) {
                return Err(DecodeError::noncanonical("known length framing"));
            }
            let length =
                usize::try_from(length).map_err(|_| DecodeError::wire("length overflow"))?;
            let value_end = payload_start
                .checked_add(length)
                .ok_or_else(|| DecodeError::wire("length overflow"))?;
            if value_end > end {
                return Err(DecodeError::wire("truncated length-delimited field"));
            }
            (payload_start, value_end, value_end, 0, 0)
        },
        3 => {
            if depth >= options.recursion_limit {
                return Err(DecodeError::nesting(
                    depth.saturating_add(1),
                    options.recursion_limit,
                ));
            }
            let mut cursor = key_end;
            let mut found_end = false;
            let mut count = 0usize;
            let mut child_work = 0usize;
            while cursor < end {
                let (child, next_child) = if let Ok((child_key, child_key_end)) =
                    read_varint(source, cursor, end, false)
                {
                    if (child_key & 7) as u8 == 4 {
                        let child_number = (child_key >> 3) as u32;
                        if child_number != number {
                            return Err(DecodeError::wire("mismatched end group"));
                        }
                        budget.fields = budget.fields.checked_add(1).ok_or_else(|| {
                            DecodeError::limit(DecodeLimit::Fields, usize::MAX, options.max_fields)
                        })?;
                        if budget.fields > options.max_fields {
                            return Err(DecodeError::limit(
                                DecodeLimit::Fields,
                                budget.fields,
                                options.max_fields,
                            ));
                        }
                        let end_work = child_key_end.saturating_sub(cursor);
                        budget.work = budget.work.checked_add(end_work).ok_or_else(|| {
                            DecodeError::limit(
                                DecodeLimit::Work,
                                usize::MAX,
                                options.max_work_bytes,
                            )
                        })?;
                        if budget.work > options.max_work_bytes {
                            return Err(DecodeError::limit(
                                DecodeLimit::Work,
                                budget.work,
                                options.max_work_bytes,
                            ));
                        }
                        (
                            ParsedField {
                                number: child_number,
                                wire: 4,
                                start: cursor,
                                value_start: child_key_end,
                                value_end: child_key_end,
                                end: child_key_end,
                                value: None,
                                field_count: 1,
                                work_bytes: end_work,
                            },
                            child_key_end,
                        )
                    } else {
                        parse_field(source, cursor, end, depth + 1, options, budget)?
                    }
                } else {
                    return Err(DecodeError::wire("truncated group field"));
                };
                cursor = next_child;
                count = count.saturating_add(child.field_count);
                child_work = child_work.saturating_add(child.work_bytes);
                if child.wire == 4 {
                    if child.number != number {
                        return Err(DecodeError::wire("mismatched end group"));
                    }
                    found_end = true;
                    break;
                }
            }
            if !found_end {
                return Err(DecodeError::wire("unterminated group"));
            }
            (key_end, cursor, cursor, count, child_work)
        },
        4 => return Err(DecodeError::wire("unexpected end group")),
        5 => {
            let value_end = key_end
                .checked_add(4)
                .ok_or_else(|| DecodeError::wire("fixed32 overflow"))?;
            if value_end > end {
                return Err(DecodeError::wire("truncated fixed32"));
            }
            (key_end, value_end, value_end, 0, 0)
        },
        _ => return Err(DecodeError::wire("invalid wire type")),
    };
    let value = if wire == 0 {
        Some(read_varint(source, value_start, value_end, false)?.0)
    } else {
        None
    };
    let field_work = next.saturating_sub(offset).saturating_add(children_work);
    budget.work = budget
        .work
        .checked_add(field_work)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes))?;
    if budget.work > options.max_work_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::Work,
            budget.work,
            options.max_work_bytes,
        ));
    }
    Ok((
        ParsedField {
            number,
            wire,
            start: offset,
            value_start,
            value_end,
            end: next,
            value,
            field_count: children_count.saturating_add(1),
            work_bytes: field_work,
        },
        next,
    ))
}

const fn is_known_field(number: u32) -> bool {
    matches!(
        number,
        SUPER_FIELD
            | START_FIELD
            | END_FIELD
            | POSTER_FIELD
            | LEGACY_LOOP_FIELD
            | VOLUME_FIELD
            | MODERN_LOOP_FIELD
    )
}

fn read_varint(
    source: &[u8],
    mut offset: usize,
    end: usize,
    canonical: bool,
) -> Result<(u64, usize), DecodeError> {
    let start = offset;
    let mut value = 0u64;
    for index in 0..10 {
        if offset >= end {
            return Err(DecodeError::wire("truncated varint"));
        }
        let byte = source[offset];
        offset += 1;
        if index == 9 && byte > 1 {
            return Err(DecodeError::wire("varint overflow"));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            if canonical && offset - start != varint_len(value) {
                return Err(DecodeError::noncanonical("known varint"));
            }
            return Ok((value, offset));
        }
    }
    Err(DecodeError::wire("varint too long"))
}

fn measure_output(
    source: &[u8],
    fields: &[ParsedField],
    source_scan: ParseAccounting,
    write: MoviePlaybackWrite,
) -> Result<OutputMeasure, DecodeError> {
    let mut length = source.len();
    let mut output_fields = source_scan.fields;
    let mut output_work = source_scan.work_bytes;
    let mut seen = [false; 6];
    for field in fields.iter().copied() {
        let index = match field.number {
            START_FIELD => Some(0),
            END_FIELD => Some(1),
            POSTER_FIELD => Some(2),
            LEGACY_LOOP_FIELD => Some(3),
            VOLUME_FIELD => Some(4),
            MODERN_LOOP_FIELD => Some(5),
            _ => None,
        };
        if let Some(index) = index {
            seen[index] = true;
        }
        let keep = match field.number {
            START_FIELD => write.start_time.is_some(),
            END_FIELD | SUPER_FIELD => true,
            POSTER_FIELD => write.poster_time.is_some(),
            VOLUME_FIELD => write.volume.is_some(),
            LEGACY_LOOP_FIELD | MODERN_LOOP_FIELD => write.loop_mode.is_some(),
            _ => true,
        };
        let replacement = if keep {
            match field.number {
                START_FIELD | END_FIELD | POSTER_FIELD | VOLUME_FIELD => {
                    fixed_field_len_value(field.number, 4)
                },
                LEGACY_LOOP_FIELD | MODERN_LOOP_FIELD => {
                    varint_field_len_for_loop(field.number, write.loop_mode.unwrap_or(0))
                },
                _ => field.end - field.start,
            }
        } else {
            0
        };
        let old_len = field.end - field.start;
        if keep {
            length = length
                .checked_sub(old_len)
                .and_then(|length| length.checked_add(replacement))
                .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
            if field.field_count == 1 {
                output_work = output_work
                    .checked_sub(field.work_bytes)
                    .and_then(|work| work.checked_add(replacement))
                    .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
            }
        } else {
            length = length
                .checked_sub(old_len)
                .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
            output_fields = output_fields
                .checked_sub(field.field_count)
                .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
            output_work = output_work
                .checked_sub(field.work_bytes)
                .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
        }
    }
    if !seen[0] && write.start_time.is_some() {
        let added = fixed_field_len_value(START_FIELD, 4);
        length = length
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
        output_fields = output_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
        output_work = output_work
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
    }
    if !seen[2] && write.poster_time.is_some() {
        let added = fixed_field_len_value(POSTER_FIELD, 4);
        length = length
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
        output_fields = output_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
        output_work = output_work
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
    }
    if !seen[4] && write.volume.is_some() {
        let added = fixed_field_len_value(VOLUME_FIELD, 4);
        length = length
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
        output_fields = output_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
        output_work = output_work
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
    }
    if !seen[3] && !seen[5] && write.loop_mode.is_some() {
        let added = varint_field_len_for_loop(MODERN_LOOP_FIELD, write.loop_mode.unwrap_or(0));
        length = length
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
        output_fields = output_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
        output_work = output_work
            .checked_add(added)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
    }
    let candidate_capacity = length.min(output_fields);
    Ok(OutputMeasure {
        bytes: length,
        fields: output_fields,
        work: output_work,
        scratch_bytes: parsed_fields_scratch(candidate_capacity)?,
    })
}

fn fixed_field_len_value(number: u32, bytes: usize) -> usize {
    varint_len(u64::from(number << 3 | 5)) + bytes
}
fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number << 3)) + varint_len(value)
}

fn varint_field_len_for_loop(number: u32, value: i32) -> usize {
    varint_field_len(number, loop_encoded_value(number, value))
}

fn loop_encoded_value(number: u32, value: i32) -> u64 {
    if number == LEGACY_LOOP_FIELD {
        u64::from(value as u32)
    } else {
        value as i64 as u64
    }
}

fn emit_rewrite(
    source: &[u8],
    write: MoviePlaybackWrite,
    expected: usize,
    options: DecodeOptions,
) -> Result<(Vec<u8>, ParseAccounting), DecodeError> {
    let (fields, accounting) = parse_root(
        source,
        options
            .with_max_input_bytes(source.len())
            .with_max_output_bytes(expected.max(source.len())),
        &mut Budget::default(),
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::allocation(expected))?;
    let mut seen = [false; 6];
    for field in fields.iter().copied() {
        let index = match field.number {
            START_FIELD => Some(0),
            END_FIELD => Some(1),
            POSTER_FIELD => Some(2),
            LEGACY_LOOP_FIELD => Some(3),
            VOLUME_FIELD => Some(4),
            MODERN_LOOP_FIELD => Some(5),
            _ => None,
        };
        if let Some(index) = index {
            seen[index] = true;
        }
        let keep = match field.number {
            START_FIELD => write.start_time.is_some(),
            END_FIELD | SUPER_FIELD => true,
            POSTER_FIELD => write.poster_time.is_some(),
            VOLUME_FIELD => write.volume.is_some(),
            LEGACY_LOOP_FIELD | MODERN_LOOP_FIELD => write.loop_mode.is_some(),
            _ => true,
        };
        if !keep {
            continue;
        }
        match field.number {
            START_FIELD => {
                append_fixed(&mut output, source, field, write.start_time.unwrap_or(0.0))
            },
            END_FIELD => append_fixed(&mut output, source, field, write.end_time),
            POSTER_FIELD => {
                append_fixed(&mut output, source, field, write.poster_time.unwrap_or(0.0))
            },
            VOLUME_FIELD => append_fixed(&mut output, source, field, write.volume.unwrap_or(0.0)),
            LEGACY_LOOP_FIELD | MODERN_LOOP_FIELD => {
                output.extend_from_slice(&source[field.start..field.value_start]);
                append_varint(
                    &mut output,
                    loop_encoded_value(field.number, write.loop_mode.unwrap_or(0)),
                );
            },
            _ => output.extend_from_slice(&source[field.start..field.end]),
        }
    }
    if !seen[0] {
        if let Some(value) = write.start_time {
            append_new_fixed(&mut output, START_FIELD, value);
        }
    }
    if !seen[2] {
        if let Some(value) = write.poster_time {
            append_new_fixed(&mut output, POSTER_FIELD, value);
        }
    }
    if !seen[4] {
        if let Some(value) = write.volume {
            append_new_fixed(&mut output, VOLUME_FIELD, value);
        }
    }
    if !seen[3] && !seen[5] {
        if let Some(value) = write.loop_mode {
            append_varint(&mut output, u64::from(MODERN_LOOP_FIELD << 3));
            append_varint(&mut output, loop_encoded_value(MODERN_LOOP_FIELD, value));
        }
    }
    if output.len() != expected {
        return Err(DecodeError::projection());
    }
    Ok((output, accounting))
}

fn snapshot_matches_write(snapshot: MoviePlaybackSnapshot, write: MoviePlaybackWrite) -> bool {
    snapshot.start_time == write.start_time
        && snapshot.end_time == write.end_time
        && snapshot.poster_time == write.poster_time
        && snapshot.loop_mode == write.loop_mode
        && snapshot.volume == write.volume
}

fn append_fixed(output: &mut Vec<u8>, source: &[u8], field: ParsedField, value: f32) {
    output.extend_from_slice(&source[field.start..field.value_start]);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
}
fn append_new_fixed(output: &mut Vec<u8>, number: u32, value: f32) {
    append_varint(output, u64::from(number << 3 | 5));
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
        return Err(DecodeError::limit(
            DecodeLimit::OutputBytes,
            requirements.output_bytes,
            limits.output_bytes,
        ));
    }
    if requirements.fields > limits.fields {
        return Err(DecodeError::limit(
            DecodeLimit::Fields,
            requirements.fields,
            limits.fields,
        ));
    }
    if requirements.work_bytes > limits.work_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::Work,
            requirements.work_bytes,
            limits.work_bytes,
        ));
    }
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limit(
            DecodeLimit::Nesting,
            requirements.max_depth as usize,
            limits.max_depth as usize,
        ));
    }
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::limit(
            DecodeLimit::Allocations,
            requirements.allocations,
            limits.allocations,
        ));
    }
    if requirements.retained_bytes > limits.retained_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::Retained,
            requirements.retained_bytes,
            limits.retained_bytes,
        ));
    }
    if requirements.scratch_bytes > limits.scratch_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::Scratch,
            requirements.scratch_bytes,
            limits.scratch_bytes,
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source() -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&[0x0a, 0x00]);
        bytes.extend_from_slice(&[0x1d, 0, 0, 0, 0]);
        bytes.extend_from_slice(&[0x25, 0, 0, 0xc0, 0x3f]);
        bytes.extend_from_slice(&[0x2d, 0, 0, 0x80, 0x3f]);
        bytes.extend_from_slice(&[0x3d, 0, 0, 0x80, 0x3f]);
        bytes.extend_from_slice(&[0xc0, 0x01, 0x00]);
        bytes
    }

    #[test]
    fn decodes_and_preserves_unknowns() {
        let mut bytes = source();
        bytes.extend_from_slice(&[
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x00,
        ]);
        let snapshot = decode_movie_playback(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(snapshot.end_time, 1.5);
        let replacement =
            MoviePlaybackWrite::from_values(Some(0.25), 2.0, Some(0.5), Some(1), Some(0.75));
        let changed = rewrite_movie_playback(
            &bytes,
            replacement,
            DecodeOptions::for_source(&bytes).with_max_output_bytes(1024),
        )
        .unwrap();
        assert!(changed.ends_with(&[
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x00
        ]));
        let restored = rewrite_movie_playback(
            &changed,
            MoviePlaybackWrite::from_values(Some(0.0), 1.5, Some(1.0), Some(0), Some(1.0)),
            DecodeOptions::for_source(&changed).with_max_output_bytes(1024),
        )
        .unwrap();
        assert_eq!(restored, bytes);
    }

    #[test]
    fn prepared_limits_replay_exactly() {
        let bytes = source();
        let prepared = prepare_movie_playback_rewrite(
            &bytes,
            MoviePlaybackWrite::new(2.0),
            DecodeOptions::for_source(&bytes).with_max_output_bytes(1024),
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.output().len(), requirements.output_bytes);
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes - 1)
                )
                .is_err()
        );
        assert!(
            prepared
                .execute(requirements.exact().with_fields(requirements.fields - 1))
                .is_err()
        );
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_work_bytes(requirements.work_bytes - 1)
                )
                .is_err()
        );
        let prepare_report = prepared.prepare_report();
        assert!(prepare_report.allocations() > 0);
        assert!(prepare_report.scratch_bytes() > 0);
        let (_, decode_report) =
            decode_movie_playback_with_report(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(prepare_report.fields(), decode_report.fields());
        assert_eq!(prepare_report.work_bytes(), decode_report.work_bytes());
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.report().fields(), requirements.fields);
        assert_eq!(output.report().work_bytes(), requirements.work_bytes);
        assert_eq!(output.report().max_depth(), requirements.max_depth);
        assert_eq!(output.report().allocations(), requirements.allocations);
        assert_eq!(
            output.report().retained_bytes(),
            requirements.retained_bytes
        );
        assert_eq!(output.report().scratch_bytes(), requirements.scratch_bytes);
    }

    #[test]
    fn prepared_requirements_reject_every_execution_axis_before_emission() {
        let bytes = source();
        let prepared = prepare_movie_playback_rewrite(
            &bytes,
            MoviePlaybackWrite::from_values(Some(0.25), 2.0, Some(0.5), Some(1), Some(0.75)),
            DecodeOptions::for_source(&bytes).with_max_output_bytes(1024),
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        let cases = [
            (
                "output",
                requirements
                    .exact()
                    .with_output_bytes(requirements.output_bytes.saturating_sub(1)),
                DecodeLimit::OutputBytes,
            ),
            (
                "fields",
                requirements
                    .exact()
                    .with_fields(requirements.fields.saturating_sub(1)),
                DecodeLimit::Fields,
            ),
            (
                "work",
                requirements
                    .exact()
                    .with_work_bytes(requirements.work_bytes.saturating_sub(1)),
                DecodeLimit::Work,
            ),
            (
                "depth",
                requirements
                    .exact()
                    .with_max_depth(requirements.max_depth.saturating_sub(1)),
                DecodeLimit::Nesting,
            ),
            (
                "allocations",
                requirements
                    .exact()
                    .with_allocations(requirements.allocations.saturating_sub(1)),
                DecodeLimit::Allocations,
            ),
            (
                "retained",
                requirements
                    .exact()
                    .with_retained_bytes(requirements.retained_bytes.saturating_sub(1)),
                DecodeLimit::Retained,
            ),
            (
                "scratch",
                requirements
                    .exact()
                    .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1)),
                DecodeLimit::Scratch,
            ),
        ];
        for (label, limits, expected) in cases {
            let error = prepared.execute(limits).unwrap_err();
            assert_eq!(error.limit_kind(), Some(expected), "{label}");
        }
    }

    #[test]
    fn prepared_custom_depth_and_input_limits_are_replayed_before_scan() {
        let mut bytes = source();
        bytes.extend_from_slice(&[0x53, 0x53, 0x53, 0x08, 0x01, 0x54, 0x54, 0x54]);
        let options = DecodeOptions::for_source(&bytes).with_recursion_limit(4);
        let prepared =
            prepare_movie_playback_rewrite(&bytes, MoviePlaybackWrite::new(2.0), options).unwrap();
        assert_eq!(prepared.prepare_report().max_depth(), 4);
        let requirements = prepared.execution_requirements();
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.report().max_depth(), requirements.max_depth);
        let too_shallow = DecodeOptions::for_source(&bytes).with_recursion_limit(3);
        assert_eq!(
            decode_movie_playback(&bytes, too_shallow)
                .unwrap_err()
                .limit_kind(),
            Some(DecodeLimit::Nesting)
        );
        assert_eq!(
            decode_movie_playback(
                &bytes,
                DecodeOptions::for_source(&bytes)
                    .with_max_input_bytes(bytes.len().saturating_sub(1)),
            )
            .unwrap_err()
            .limit_kind(),
            Some(DecodeLimit::InputBytes)
        );
    }

    #[test]
    fn rejects_conflicts_and_noncanonical_known_varints() {
        let mut conflict = source();
        conflict.extend_from_slice(&[0xc0, 0x01, 0x01]);
        assert!(decode_movie_playback(&conflict, DecodeOptions::for_source(&conflict)).is_err());
        let mut noncanonical = source();
        noncanonical.extend_from_slice(&[0xc0, 0x01, 0x80, 0x00]);
        assert!(
            decode_movie_playback(&noncanonical, DecodeOptions::for_source(&noncanonical)).is_err()
        );
    }

    #[test]
    fn balanced_unknown_group_at_eof_is_retained() {
        let mut bytes = source();
        bytes.extend_from_slice(&[0x53, 0x08, 0x01, 0x54]);
        let snapshot = decode_movie_playback(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        let write = MoviePlaybackWrite::from_values(
            snapshot.start_time,
            snapshot.end_time,
            snapshot.poster_time,
            snapshot.loop_mode,
            snapshot.volume,
        );
        let output = rewrite_movie_playback(
            &bytes,
            write,
            DecodeOptions::for_source(&bytes).with_max_output_bytes(bytes.len() + 64),
        )
        .unwrap();
        assert!(output.ends_with(&[0x53, 0x08, 0x01, 0x54]));
    }

    #[test]
    fn modern_negative_loop_uses_sign_extended_canonical_varint() {
        let mut bytes = vec![0x0a, 0x00, 0x25, 0x00, 0x00, 0x80, 0x3f, 0xc0, 0x01];
        bytes.extend_from_slice(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x01]);
        let snapshot = decode_movie_playback(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(snapshot.loop_mode, Some(-1));
        let output = rewrite_movie_playback(
            &bytes,
            MoviePlaybackWrite::from_values(None, 1.0, None, Some(-1), None),
            DecodeOptions::for_source(&bytes).with_max_output_bytes(bytes.len() + 32),
        )
        .unwrap();
        assert_eq!(output, bytes);
    }
}
