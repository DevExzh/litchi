//! Strict, generated-free projection and source-preserving rewrite for the
//! selected adjustment controls in `TSD.ImageArchive`.
//!
//! The complete image archive and every advanced adjustment control remain
//! caller-owned. This module validates the complete outer payload, projects
//! only image-adjustments field 14 and its public scalar controls, and keeps
//! the original bytes authoritative for rewrites.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict scanner intentionally precedes the private view."
)]

use std::{fmt, mem::size_of};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_image_adjustments_generated::LitchiIwaImageAdjustmentsProjection as projection;

const IMAGE_ADJUSTMENTS_FIELD: u32 = 14;
const EXPOSURE_FIELD: u32 = 1;
const SATURATION_FIELD: u32 = 2;
const ENHANCE_FIELD: u32 = 13;
const MAX_RECURSION: u32 = 64;

/// Finite limits for one complete `TSD.ImageArchive` payload.
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

    /// Return the configured field ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Return the configured work ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
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

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
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

    /// Return exact observed/maximum values for a resource failure.
    #[must_use]
    pub const fn limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit {
                observed, maximum, ..
            } => Some((observed, maximum)),
            DecodeErrorKind::Nesting { observed, maximum } => {
                Some((observed as usize, maximum as usize))
            },
            _ => None,
        }
    }

    /// Return the structured resource observation used by format adapters.
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
            } => write!(
                formatter,
                "image adjustments {axis:?} limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Nesting { observed, maximum } => write!(
                formatter,
                "image adjustments Nesting limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Allocation(bytes) => write!(
                formatter,
                "image adjustments output allocation failed for {bytes} bytes"
            ),
            DecodeErrorKind::Projection => formatter
                .write_str("Buffa image-adjustments projection disagrees with strict source scan"),
        }
    }
}

impl std::error::Error for DecodeError {}

/// Borrowed semantic facts from one complete `TSD.ImageArchive` payload.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageAdjustmentsSnapshot<'source> {
    exposure: Option<f32>,
    saturation: Option<f32>,
    enhance: Option<bool>,
    has_image_adjustments: bool,
    raw: &'source [u8],
}

impl<'source> ImageAdjustmentsSnapshot<'source> {
    /// Return optional native exposure, preserving proto2 presence.
    #[must_use]
    pub const fn exposure(self) -> Option<f32> {
        self.exposure
    }

    /// Return optional native saturation, preserving proto2 presence.
    #[must_use]
    pub const fn saturation(self) -> Option<f32> {
        self.saturation
    }

    /// Return optional native automatic-enhancement value.
    #[must_use]
    pub const fn enhance(self) -> Option<bool> {
        self.enhance
    }

    /// Alias for [`Self::enhance`].
    #[must_use]
    pub const fn enhancement(self) -> Option<bool> {
        self.enhance
    }

    /// Return the effective enhancement value, treating absence as false.
    #[must_use]
    pub const fn is_enhanced(self) -> bool {
        matches!(self.enhance, Some(true))
    }

    /// Return whether outer field 14 was present, including an empty message.
    #[must_use]
    pub const fn has_image_adjustments(self) -> bool {
        self.has_image_adjustments
    }

    /// Return the exact caller-owned bytes that were validated.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.raw
    }
}

/// Requested semantic values for an image-adjustments rewrite.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageAdjustmentsWrite {
    exposure: Option<f32>,
    saturation: Option<f32>,
    enhance: Option<bool>,
}

impl Default for ImageAdjustmentsWrite {
    fn default() -> Self {
        Self::new()
    }
}

impl ImageAdjustmentsWrite {
    /// Construct a request that clears all selected fields.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            exposure: None,
            saturation: None,
            enhance: None,
        }
    }

    /// Construct all selected values directly.
    #[must_use]
    pub const fn from_values(
        exposure: Option<f32>,
        saturation: Option<f32>,
        enhance: Option<bool>,
    ) -> Self {
        Self {
            exposure,
            saturation,
            enhance,
        }
    }

    /// Replace the requested exposure.
    #[must_use]
    pub const fn with_exposure(mut self, value: Option<f32>) -> Self {
        self.exposure = value;
        self
    }

    /// Replace the requested saturation.
    #[must_use]
    pub const fn with_saturation(mut self, value: Option<f32>) -> Self {
        self.saturation = value;
        self
    }

    /// Replace the requested enhancement.
    #[must_use]
    pub const fn with_enhance(mut self, value: Option<bool>) -> Self {
        self.enhance = value;
        self
    }

    /// Alias for [`Self::with_enhance`].
    #[must_use]
    pub const fn with_enhancement(self, value: Option<bool>) -> Self {
        self.with_enhance(value)
    }

    /// Return the requested exposure.
    #[must_use]
    pub const fn exposure(self) -> Option<f32> {
        self.exposure
    }

    /// Return the requested saturation.
    #[must_use]
    pub const fn saturation(self) -> Option<f32> {
        self.saturation
    }

    /// Return the requested enhancement.
    #[must_use]
    pub const fn enhance(self) -> Option<bool> {
        self.enhance
    }

    /// Alias for [`Self::enhance`].
    #[must_use]
    pub const fn enhancement(self) -> Option<bool> {
        self.enhance
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

/// Prepared output and exact execution requirements.
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
    /// Return exact limits for this prepared operation.
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
pub struct PreparedImageAdjustmentsRewrite<'source> {
    source: &'source [u8],
    write: ImageAdjustmentsWrite,
    options: DecodeOptions,
    report: DecodeReport,
    requirements: RewriteExecutionRequirements,
    source_fields: usize,
    source_work_bytes: usize,
    source_allocations: usize,
    candidate_fields: usize,
    candidate_work_bytes: usize,
    candidate_allocations: usize,
    candidate_scratch_bytes: usize,
}

impl<'source> PreparedImageAdjustmentsRewrite<'source> {
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
        check_options_limits(self.requirements, self.options)?;
        let (output, source_scan) = emit_rewrite(
            self.source,
            self.write,
            self.requirements.output_bytes,
            self.options,
        )?;
        if source_scan.fields != self.source_fields
            || source_scan.work_bytes != self.source_work_bytes
            || source_scan.allocations != self.source_allocations
            || source_scan.max_depth > self.requirements.max_depth
        {
            return Err(DecodeError::projection());
        }

        let readback_options = DecodeOptions::new(
            output.len(),
            self.candidate_fields,
            self.candidate_work_bytes,
            self.options.recursion_limit,
        );
        let (snapshot, candidate_scan) = scan_source_parts(&output, readback_options)?;
        if candidate_scan.fields != self.candidate_fields
            || candidate_scan.work_bytes != self.candidate_work_bytes
            || candidate_scan.allocations != self.candidate_allocations
            || candidate_scan.scratch_bytes != self.candidate_scratch_bytes
            || candidate_scan.max_depth > self.requirements.max_depth
        {
            return Err(DecodeError::projection());
        }
        force_buffa(
            &output,
            readback_options.with_max_input_bytes(output.len()),
            snapshot,
        )?;
        if !snapshot_matches_write(snapshot, self.write) {
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

/// Strictly decode image-adjustment controls from one complete ImageArchive.
pub fn decode_image_adjustments(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ImageAdjustmentsSnapshot<'_>, DecodeError> {
    Ok(decode_image_adjustments_with_report(source, options)?.0)
}

/// Strictly decode image adjustments and return exact source accounting.
pub fn decode_image_adjustments_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ImageAdjustmentsSnapshot<'_>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let (snapshot, scan) = scan_source_parts(source, options)?;
    force_buffa(source, options, snapshot)?;
    Ok((
        snapshot,
        DecodeReport {
            input_bytes: source.len(),
            fields: scan.fields,
            work_bytes: scan.work_bytes,
            max_depth: scan.max_depth,
            allocations: scan.allocations,
            retained_bytes: source.len(),
            scratch_bytes: scan.scratch_bytes,
        },
    ))
}

/// Prepare a strict image-adjustments rewrite without allocating output.
pub fn prepare_image_adjustments_rewrite<'source>(
    source: &'source [u8],
    write: ImageAdjustmentsWrite,
    options: DecodeOptions,
) -> Result<PreparedImageAdjustmentsRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    let scan = scan_source_parts_with_storage(source, options)?;
    force_buffa(source, options, scan.snapshot)?;
    validate_write(write)?;
    let measure = measure_output(source, &scan, write)?;
    if measure.bytes > options.max_output_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::OutputBytes,
            measure.bytes,
            options.max_output_bytes,
        ));
    }
    let requirements = RewriteExecutionRequirements {
        output_bytes: measure.bytes,
        fields: scan.fields.checked_add(measure.fields).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?,
        work_bytes: scan.work_bytes.checked_add(measure.work).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes)
        })?,
        max_depth: scan.max_depth.max(measure.max_depth),
        allocations: scan
            .allocations
            .checked_add(usize::from(measure.bytes != 0))
            .and_then(|value| value.checked_add(measure.allocations))
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Allocations, usize::MAX, usize::MAX))?,
        retained_bytes: measure.bytes,
        scratch_bytes: scan
            .scratch_bytes
            .checked_add(measure.scratch_bytes)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))?,
    };
    check_options_limits(requirements, options)?;
    Ok(PreparedImageAdjustmentsRewrite {
        source,
        write,
        options,
        report: DecodeReport {
            input_bytes: source.len(),
            fields: scan.fields,
            work_bytes: scan.work_bytes,
            max_depth: scan.max_depth,
            allocations: scan.allocations,
            retained_bytes: source.len(),
            scratch_bytes: scan.scratch_bytes,
        },
        requirements,
        source_fields: scan.fields,
        source_work_bytes: scan.work_bytes,
        source_allocations: scan.allocations,
        candidate_fields: measure.fields,
        candidate_work_bytes: measure.work,
        candidate_allocations: measure.allocations,
        candidate_scratch_bytes: measure.scratch_bytes,
    })
}

/// Rewrite image-adjustment controls with exact prepared limits.
pub fn rewrite_image_adjustments(
    source: &[u8],
    write: ImageAdjustmentsWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    let prepared = prepare_image_adjustments_rewrite(source, write, options)?;
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

fn validate_write(write: ImageAdjustmentsWrite) -> Result<(), DecodeError> {
    validate_adjustment(write.exposure, "exposure")?;
    validate_adjustment(write.saturation, "saturation")?;
    Ok(())
}

fn validate_adjustment(value: Option<f32>, name: &'static str) -> Result<(), DecodeError> {
    if value.is_some_and(|value| !value.is_finite() || !(-1.0..=1.0).contains(&value)) {
        return Err(DecodeError::invalid(name));
    }
    Ok(())
}

fn force_buffa(
    source: &[u8],
    options: DecodeOptions,
    strict: ImageAdjustmentsSnapshot<'_>,
) -> Result<(), DecodeError> {
    let view: projection::ImageArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::projection())?;
    let projected = view
        .image_adjustments
        .get()
        .map_err(|_| DecodeError::projection())?;
    let Some(projected) = projected else {
        return (!strict.has_image_adjustments)
            .then_some(())
            .ok_or_else(DecodeError::projection);
    };
    if !strict.has_image_adjustments
        || projected.exposure != strict.exposure
        || projected.saturation != strict.saturation
        || projected.enhance != strict.enhance
    {
        return Err(DecodeError::projection());
    }
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
    field_count: usize,
    work_bytes: usize,
    max_depth: u32,
    value: Option<u64>,
}

#[derive(Debug, Default)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

impl Budget {
    fn field(&mut self, options: DecodeOptions) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Fields, usize::MAX, options.max_fields)
        })?;
        if self.fields > options.max_fields {
            return Err(DecodeError::limit(
                DecodeLimit::Fields,
                self.fields,
                options.max_fields,
            ));
        }
        Ok(())
    }

    fn work(&mut self, amount: usize, options: DecodeOptions) -> Result<(), DecodeError> {
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes)
        })?;
        if self.work_bytes > options.max_work_bytes {
            return Err(DecodeError::limit(
                DecodeLimit::Work,
                self.work_bytes,
                options.max_work_bytes,
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct ParseAccounting {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

struct ArchiveScan<'source> {
    snapshot: ImageAdjustmentsSnapshot<'source>,
    outer: Vec<ParsedField>,
    nested: Option<NestedScan>,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

struct NestedScan {
    fields: Vec<ParsedField>,
    accounting: ParseAccounting,
}

fn scan_source_parts<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(ImageAdjustmentsSnapshot<'source>, ParseAccounting), DecodeError> {
    let mut budget = Budget::default();
    if 1 > options.recursion_limit {
        return Err(DecodeError::nesting(1, options.recursion_limit));
    }
    let (outer, outer_accounting) = parse_message(source, 1, options, &mut budget)?;
    let mut nested = None;
    let mut exposure = None;
    let mut saturation = None;
    let mut enhance = None;
    let mut has_image_adjustments = false;
    for field in outer.iter().copied() {
        if field.number != IMAGE_ADJUSTMENTS_FIELD {
            continue;
        }
        if has_image_adjustments {
            return Err(DecodeError::duplicate("TSD.ImageArchive.imageAdjustments"));
        }
        has_image_adjustments = true;
        if field.wire != 2 {
            return Err(DecodeError::wrong_wire("TSD.ImageArchive.imageAdjustments"));
        }
        let payload = &source[field.value_start..field.value_end];
        let (nested_fields, nested_accounting) = parse_message(payload, 2, options, &mut budget)?;
        for nested_field in nested_fields.iter().copied() {
            match nested_field.number {
                EXPOSURE_FIELD => {
                    if exposure.is_some() {
                        return Err(DecodeError::duplicate("ImageAdjustmentsArchive.exposure"));
                    }
                    require_wire(nested_field, 5, "ImageAdjustmentsArchive.exposure")?;
                    exposure = Some(fixed32(payload, nested_field, "exposure")?);
                },
                SATURATION_FIELD => {
                    if saturation.is_some() {
                        return Err(DecodeError::duplicate("ImageAdjustmentsArchive.saturation"));
                    }
                    require_wire(nested_field, 5, "ImageAdjustmentsArchive.saturation")?;
                    saturation = Some(fixed32(payload, nested_field, "saturation")?);
                },
                ENHANCE_FIELD => {
                    if enhance.is_some() {
                        return Err(DecodeError::duplicate("ImageAdjustmentsArchive.enhance"));
                    }
                    require_wire(nested_field, 0, "ImageAdjustmentsArchive.enhance")?;
                    enhance = Some(canonical_bool(nested_field, "enhance")?);
                },
                _ => {},
            }
        }
        nested = Some(NestedScan {
            fields: nested_fields,
            accounting: nested_accounting,
        });
    }
    let snapshot = ImageAdjustmentsSnapshot {
        exposure,
        saturation,
        enhance,
        has_image_adjustments,
        raw: source,
    };
    let accounting = ParseAccounting {
        fields: budget.fields,
        work_bytes: budget.work_bytes,
        max_depth: budget.max_depth,
        allocations: outer_accounting.allocations
            + nested
                .as_ref()
                .map_or(0, |nested| nested.accounting.allocations),
        scratch_bytes: outer_accounting.scratch_bytes
            + nested
                .as_ref()
                .map_or(0, |nested| nested.accounting.scratch_bytes),
    };
    // The nested parser is owned by a temporary scan only to project values;
    // execution reparses it. Keep this function's public accounting compact.
    let _ = nested;
    Ok((snapshot, accounting))
}

fn scan_source_parts_with_storage<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ArchiveScan<'source>, DecodeError> {
    let mut budget = Budget::default();
    let (outer, outer_accounting) = parse_message(source, 1, options, &mut budget)?;
    let mut nested = None;
    let mut exposure = None;
    let mut saturation = None;
    let mut enhance = None;
    let mut has_image_adjustments = false;
    for field in outer.iter().copied() {
        if field.number != IMAGE_ADJUSTMENTS_FIELD {
            continue;
        }
        if has_image_adjustments {
            return Err(DecodeError::duplicate("TSD.ImageArchive.imageAdjustments"));
        }
        has_image_adjustments = true;
        require_wire(field, 2, "TSD.ImageArchive.imageAdjustments")?;
        let payload = &source[field.value_start..field.value_end];
        let (nested_fields, nested_accounting) = parse_message(payload, 2, options, &mut budget)?;
        for nested_field in nested_fields.iter().copied() {
            match nested_field.number {
                EXPOSURE_FIELD => {
                    if exposure.is_some() {
                        return Err(DecodeError::duplicate("ImageAdjustmentsArchive.exposure"));
                    }
                    require_wire(nested_field, 5, "ImageAdjustmentsArchive.exposure")?;
                    exposure = Some(fixed32(payload, nested_field, "exposure")?);
                },
                SATURATION_FIELD => {
                    if saturation.is_some() {
                        return Err(DecodeError::duplicate("ImageAdjustmentsArchive.saturation"));
                    }
                    require_wire(nested_field, 5, "ImageAdjustmentsArchive.saturation")?;
                    saturation = Some(fixed32(payload, nested_field, "saturation")?);
                },
                ENHANCE_FIELD => {
                    if enhance.is_some() {
                        return Err(DecodeError::duplicate("ImageAdjustmentsArchive.enhance"));
                    }
                    require_wire(nested_field, 0, "ImageAdjustmentsArchive.enhance")?;
                    enhance = Some(canonical_bool(nested_field, "enhance")?);
                },
                _ => {},
            }
        }
        nested = Some(NestedScan {
            fields: nested_fields,
            accounting: nested_accounting,
        });
    }
    let nested_allocations = nested
        .as_ref()
        .map_or(0, |nested| nested.accounting.allocations);
    let nested_scratch_bytes = nested
        .as_ref()
        .map_or(0, |nested| nested.accounting.scratch_bytes);
    Ok(ArchiveScan {
        snapshot: ImageAdjustmentsSnapshot {
            exposure,
            saturation,
            enhance,
            has_image_adjustments,
            raw: source,
        },
        outer,
        nested,
        fields: budget.fields,
        work_bytes: budget.work_bytes,
        max_depth: budget.max_depth,
        allocations: outer_accounting.allocations + nested_allocations,
        scratch_bytes: outer_accounting.scratch_bytes + nested_scratch_bytes,
    })
}

fn parse_message(
    source: &[u8],
    depth: u32,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(Vec<ParsedField>, ParseAccounting), DecodeError> {
    if depth > options.recursion_limit {
        return Err(DecodeError::nesting(depth, options.recursion_limit));
    }
    budget.max_depth = budget.max_depth.max(depth);
    let mut fields = Vec::new();
    let mut allocations = 0;
    let before_fields = budget.fields;
    let before_work = budget.work_bytes;
    let mut offset = 0;
    while offset < source.len() {
        reserve_parsed_fields(&mut fields, options.max_fields, &mut allocations)?;
        let (field, next) = parse_field(source, offset, source.len(), depth, options, budget)?;
        fields.push(field);
        offset = next;
    }
    let field_count = budget.fields.saturating_sub(before_fields);
    let work_bytes = budget.work_bytes.saturating_sub(before_work);
    let max_depth = fields
        .iter()
        .map(|field| field.max_depth)
        .max()
        .unwrap_or(depth.min(options.recursion_limit));
    let scratch_bytes = parsed_fields_scratch(fields.capacity())?;
    Ok((
        fields,
        ParseAccounting {
            fields: field_count,
            work_bytes,
            max_depth,
            allocations,
            scratch_bytes,
        },
    ))
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
    let (key, key_end) = read_varint(source, offset, end)?;
    let number = u32::try_from(key >> 3).map_err(|_| DecodeError::wire("field number overflow"))?;
    let wire = (key & 7) as u8;
    if number == 0 || number > 0x1fff_ffff {
        return Err(DecodeError::wire("invalid field number"));
    }
    budget.field(options)?;
    let (value_start, value_end, next, child_count, child_depth, child_work, value) = match wire {
        0 => {
            let (value, value_end) = read_varint(source, key_end, end)?;
            (key_end, value_end, value_end, 0, depth, 0, Some(value))
        },
        1 => {
            let value_end = key_end
                .checked_add(8)
                .ok_or_else(|| DecodeError::wire("fixed64 overflow"))?;
            if value_end > end {
                return Err(DecodeError::wire("truncated fixed64"));
            }
            (key_end, value_end, value_end, 0, depth, 0, None)
        },
        2 => {
            let (length, payload_start) = read_varint(source, key_end, end)?;
            let length = usize::try_from(length)
                .map_err(|_| DecodeError::wire("length-delimited size overflow"))?;
            let value_end = payload_start
                .checked_add(length)
                .ok_or_else(|| DecodeError::wire("length-delimited size overflow"))?;
            if value_end > end {
                return Err(DecodeError::wire("truncated length-delimited field"));
            }
            (payload_start, value_end, value_end, 0, depth, 0, None)
        },
        3 => {
            if depth >= options.recursion_limit {
                return Err(DecodeError::nesting(
                    depth.saturating_add(1),
                    options.recursion_limit,
                ));
            }
            let mut cursor = key_end;
            let mut child_count = 0usize;
            let mut max_depth = depth;
            let mut child_work = 0usize;
            loop {
                if cursor >= end {
                    return Err(DecodeError::wire("unterminated group"));
                }
                let (child_key, child_key_end) = read_varint(source, cursor, end)?;
                let child_number = u32::try_from(child_key >> 3)
                    .map_err(|_| DecodeError::wire("field number overflow"))?;
                let child_wire = (child_key & 7) as u8;
                if child_number == 0 || child_number > 0x1fff_ffff {
                    return Err(DecodeError::wire("invalid field number"));
                }
                if child_wire == 4 {
                    if child_number != number {
                        return Err(DecodeError::wire("mismatched end group"));
                    }
                    budget.field(options)?;
                    let end_work = child_key_end - cursor;
                    budget.work(end_work, options)?;
                    cursor = child_key_end;
                    child_count = child_count.saturating_add(1);
                    child_work = child_work.checked_add(end_work).ok_or_else(|| {
                        DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes)
                    })?;
                    max_depth = max_depth.max(depth + 1);
                    break;
                }
                let (child, next_child) =
                    parse_field(source, cursor, end, depth + 1, options, budget)?;
                child_count = child_count.saturating_add(child.field_count);
                child_work = child_work.checked_add(child.work_bytes).ok_or_else(|| {
                    DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes)
                })?;
                max_depth = max_depth.max(child.max_depth);
                cursor = next_child;
            }
            budget.max_depth = budget.max_depth.max(max_depth);
            (
                key_end,
                cursor,
                cursor,
                child_count,
                max_depth,
                child_work,
                None,
            )
        },
        4 => return Err(DecodeError::wire("unexpected end group")),
        5 => {
            let value_end = key_end
                .checked_add(4)
                .ok_or_else(|| DecodeError::wire("fixed32 overflow"))?;
            if value_end > end {
                return Err(DecodeError::wire("truncated fixed32"));
            }
            (key_end, value_end, value_end, 0, depth, 0, None)
        },
        _ => return Err(DecodeError::wire("invalid wire type")),
    };
    let span = next.saturating_sub(offset);
    budget.work(span, options)?;
    let field_work = span
        .checked_add(child_work)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, options.max_work_bytes))?;
    Ok((
        ParsedField {
            number,
            wire,
            start: offset,
            value_start,
            value_end,
            end: next,
            field_count: child_count.saturating_add(1),
            work_bytes: field_work,
            max_depth: child_depth,
            value,
        },
        next,
    ))
}

fn parsed_fields_scratch(capacity: usize) -> Result<usize, DecodeError> {
    capacity
        .checked_mul(size_of::<ParsedField>())
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))
}

fn parsed_fields_reservation(field_count: usize, max_fields: usize) -> (usize, usize) {
    let mut capacity = 0;
    let mut allocations = 0;
    while capacity < field_count {
        let remaining = max_fields.saturating_sub(capacity);
        if remaining == 0 {
            break;
        }
        let additional = capacity.max(1).min(remaining);
        capacity += additional;
        allocations += 1;
    }
    (capacity, allocations)
}

fn reserve_parsed_fields(
    fields: &mut Vec<ParsedField>,
    max_fields: usize,
    allocations: &mut usize,
) -> Result<(), DecodeError> {
    if fields.len() != fields.capacity() {
        return Ok(());
    }
    let remaining = max_fields.saturating_sub(fields.len());
    if remaining == 0 {
        return Ok(());
    }
    let additional = fields.capacity().max(1).min(remaining);
    let allocation_bytes = additional
        .checked_mul(size_of::<ParsedField>())
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))?;
    let previous_capacity = fields.capacity();
    fields
        .try_reserve_exact(additional)
        .map_err(|_| DecodeError::allocation(allocation_bytes))?;
    if fields.capacity() > previous_capacity {
        *allocations = allocations
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Allocations, usize::MAX, usize::MAX))?;
    }
    Ok(())
}

fn read_varint(source: &[u8], mut offset: usize, end: usize) -> Result<(u64, usize), DecodeError> {
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
            if offset - start != varint_len(value) {
                return Err(DecodeError::noncanonical("varint"));
            }
            return Ok((value, offset));
        }
    }
    Err(DecodeError::wire("varint too long"))
}

fn varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64 - value.leading_zeros()).div_ceil(7) as usize
    }
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
    if !value.is_finite() || !(-1.0..=1.0).contains(&value) {
        return Err(DecodeError::invalid(name));
    }
    Ok(value)
}

fn canonical_bool(field: ParsedField, name: &'static str) -> Result<bool, DecodeError> {
    match field.value {
        Some(0) => Ok(false),
        Some(1) => Ok(true),
        Some(_) => Err(DecodeError::noncanonical(name)),
        None => Err(DecodeError::wrong_wire(name)),
    }
}

#[derive(Debug, Clone, Copy)]
struct NestedMeasure {
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    vector_fields: usize,
}

#[derive(Debug, Clone, Copy)]
struct OutputMeasure {
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    allocations: usize,
    scratch_bytes: usize,
}

fn selected_index(number: u32) -> Option<usize> {
    match number {
        EXPOSURE_FIELD => Some(0),
        SATURATION_FIELD => Some(1),
        ENHANCE_FIELD => Some(2),
        _ => None,
    }
}

fn requested(write: ImageAdjustmentsWrite, index: usize) -> bool {
    match index {
        0 => write.exposure.is_some(),
        1 => write.saturation.is_some(),
        2 => write.enhance.is_some(),
        _ => false,
    }
}

fn nested_measure(
    payload_len: usize,
    fields: Option<&NestedScan>,
    write: ImageAdjustmentsWrite,
) -> Result<NestedMeasure, DecodeError> {
    let mut bytes = payload_len;
    let mut output_fields = fields.map_or(0, |scan| scan.accounting.fields);
    let mut work = fields.map_or(0, |scan| scan.accounting.work_bytes);
    let mut max_depth = fields.map_or(0, |scan| scan.accounting.max_depth);
    let mut vector_fields = fields.map_or(0, |scan| scan.fields.len());
    let mut seen = [false; 3];
    if let Some(scan) = fields {
        for field in scan.fields.iter().copied() {
            let Some(index) = selected_index(field.number) else {
                continue;
            };
            seen[index] = true;
            if requested(write, index) {
                continue;
            }
            bytes = bytes
                .checked_sub(field.end - field.start)
                .ok_or_else(DecodeError::projection)?;
            output_fields = output_fields
                .checked_sub(field.field_count)
                .ok_or_else(DecodeError::projection)?;
            vector_fields = vector_fields
                .checked_sub(1)
                .ok_or_else(DecodeError::projection)?;
            work = work
                .checked_sub(field.work_bytes)
                .ok_or_else(DecodeError::projection)?;
        }
    }
    let additions = [
        (0, write.exposure.map(|_| fixed_field_len(EXPOSURE_FIELD))),
        (
            1,
            write.saturation.map(|_| fixed_field_len(SATURATION_FIELD)),
        ),
        (2, write.enhance.map(|_| varint_field_len(ENHANCE_FIELD, 1))),
    ];
    for (index, addition) in additions {
        if !seen[index] {
            if let Some(addition) = addition {
                bytes = bytes
                    .checked_add(addition)
                    .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
                output_fields = output_fields
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
                vector_fields = vector_fields
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
                work = work
                    .checked_add(addition)
                    .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
                max_depth = max_depth.max(2);
            }
        }
    }
    Ok(NestedMeasure {
        bytes,
        fields: output_fields,
        work,
        max_depth,
        vector_fields,
    })
}

fn measure_output(
    source: &[u8],
    scan: &ArchiveScan<'_>,
    write: ImageAdjustmentsWrite,
) -> Result<OutputMeasure, DecodeError> {
    // Semantic equality is a byte-preserving no-op, including an explicitly
    // present empty field-14 message. Do not apply the normal "empty nested
    // payload removes its outer field" rule to that path.
    if snapshot_matches_write(scan.snapshot, write) {
        let (outer_capacity, outer_allocations) =
            parsed_fields_reservation(scan.outer.len(), scan.fields);
        let nested_vector_fields = scan.nested.as_ref().map_or(0, |nested| nested.fields.len());
        let (nested_capacity, nested_allocations) =
            parsed_fields_reservation(nested_vector_fields, scan.fields);
        let scratch_bytes = parsed_fields_scratch(outer_capacity)?
            .checked_add(parsed_fields_scratch(nested_capacity)?)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))?;
        let allocations = outer_allocations
            .checked_add(nested_allocations)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Allocations, usize::MAX, usize::MAX))?;
        return Ok(OutputMeasure {
            bytes: source.len(),
            fields: scan.fields,
            work: scan.work_bytes,
            max_depth: scan.max_depth,
            allocations,
            scratch_bytes,
        });
    }
    let outer_field = scan
        .outer
        .iter()
        .copied()
        .find(|field| field.number == IMAGE_ADJUSTMENTS_FIELD);
    let nested_payload_len = outer_field.map_or(0, |field| field.value_end - field.value_start);
    let nested = nested_measure(nested_payload_len, scan.nested.as_ref(), write)?;
    let mut bytes = source.len();
    let mut outer_fields = scan
        .outer
        .iter()
        .map(|field| field.field_count)
        .sum::<usize>();
    let mut outer_vector_fields = scan.outer.len();
    let mut outer_work = scan
        .outer
        .iter()
        .map(|field| field.work_bytes)
        .sum::<usize>();
    if let Some(field) = outer_field {
        if nested.bytes == 0 {
            bytes = bytes
                .checked_sub(field.end - field.start)
                .ok_or_else(DecodeError::projection)?;
            outer_fields = outer_fields
                .checked_sub(field.field_count)
                .ok_or_else(DecodeError::projection)?;
            outer_vector_fields = outer_vector_fields
                .checked_sub(1)
                .ok_or_else(DecodeError::projection)?;
            outer_work = outer_work
                .checked_sub(field.work_bytes)
                .ok_or_else(DecodeError::projection)?;
        } else {
            let replacement = length_delimited_field_len(IMAGE_ADJUSTMENTS_FIELD, nested.bytes);
            bytes = bytes
                .checked_sub(field.end - field.start)
                .and_then(|value| value.checked_add(replacement))
                .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
            outer_work = outer_work
                .checked_sub(field.work_bytes)
                .and_then(|value| value.checked_add(replacement))
                .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
        }
    } else if nested.bytes != 0 {
        let addition = length_delimited_field_len(IMAGE_ADJUSTMENTS_FIELD, nested.bytes);
        bytes = bytes
            .checked_add(addition)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::OutputBytes, usize::MAX, 0))?;
        outer_fields = outer_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
        outer_vector_fields = outer_vector_fields
            .checked_add(1)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
        outer_work = outer_work
            .checked_add(addition)
            .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
    }
    let fields = outer_fields
        .checked_add(nested.fields)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Fields, usize::MAX, 0))?;
    let work = outer_work
        .checked_add(nested.work)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Work, usize::MAX, 0))?;
    // Execution replays the candidate with `max_fields == fields`, so these
    // capacities are the exact reservations used by the candidate pass.
    let (outer_capacity, outer_allocations) =
        parsed_fields_reservation(outer_vector_fields, fields);
    let (nested_capacity, nested_allocations) =
        parsed_fields_reservation(nested.vector_fields, fields);
    let scratch_bytes = parsed_fields_scratch(outer_capacity)?
        .checked_add(parsed_fields_scratch(nested_capacity)?)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Scratch, usize::MAX, usize::MAX))?;
    let allocations = outer_allocations
        .checked_add(nested_allocations)
        .ok_or_else(|| DecodeError::limit(DecodeLimit::Allocations, usize::MAX, usize::MAX))?;
    Ok(OutputMeasure {
        bytes,
        fields,
        work,
        max_depth: scan.max_depth.max(nested.max_depth),
        allocations,
        scratch_bytes,
    })
}

fn fixed_field_len(number: u32) -> usize {
    varint_len(u64::from((number << 3) | 5)) + 4
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number << 3)) + varint_len(value)
}

fn length_delimited_field_len(number: u32, payload_len: usize) -> usize {
    varint_len(u64::from(number << 3 | 2)) + varint_len(payload_len as u64) + payload_len
}

fn emit_rewrite(
    source: &[u8],
    write: ImageAdjustmentsWrite,
    expected: usize,
    options: DecodeOptions,
) -> Result<(Vec<u8>, ParseAccounting), DecodeError> {
    let scan = scan_source_parts_with_storage(
        source,
        options
            .with_max_input_bytes(source.len())
            .with_max_output_bytes(expected.max(source.len())),
    )?;
    let source_accounting = ParseAccounting {
        fields: scan.fields,
        work_bytes: scan.work_bytes,
        max_depth: scan.max_depth,
        allocations: scan.allocations,
        scratch_bytes: scan.scratch_bytes,
    };
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::allocation(expected))?;
    if snapshot_matches_write(scan.snapshot, write) {
        output.extend_from_slice(source);
        return Ok((output, source_accounting));
    }
    let outer_field = scan
        .outer
        .iter()
        .copied()
        .find(|field| field.number == IMAGE_ADJUSTMENTS_FIELD);
    for field in scan.outer.iter().copied() {
        if field.number != IMAGE_ADJUSTMENTS_FIELD {
            output.extend_from_slice(&source[field.start..field.end]);
            continue;
        }
        let nested_scan = scan.nested.as_ref();
        let payload = &source[field.value_start..field.value_end];
        let nested_expected = nested_measure(payload.len(), nested_scan, write)?.bytes;
        let nested_output = emit_nested(payload, nested_scan, write, nested_expected)?;
        if nested_output.is_empty() {
            continue;
        }
        append_varint(&mut output, u64::from((IMAGE_ADJUSTMENTS_FIELD << 3) | 2));
        append_varint(&mut output, nested_output.len() as u64);
        output.extend_from_slice(&nested_output);
    }
    if outer_field.is_none() {
        let nested_expected = nested_measure(0, None, write)?.bytes;
        if nested_expected != 0 {
            let nested_output = emit_nested(&[], None, write, nested_expected)?;
            append_varint(&mut output, u64::from((IMAGE_ADJUSTMENTS_FIELD << 3) | 2));
            append_varint(&mut output, nested_output.len() as u64);
            output.extend_from_slice(&nested_output);
        }
    }
    if output.len() != expected {
        return Err(DecodeError::projection());
    }
    Ok((output, source_accounting))
}

fn emit_nested(
    source: &[u8],
    scan: Option<&NestedScan>,
    write: ImageAdjustmentsWrite,
    expected: usize,
) -> Result<Vec<u8>, DecodeError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(expected)
        .map_err(|_| DecodeError::allocation(expected))?;
    let mut seen = [false; 3];
    if let Some(scan) = scan {
        for field in scan.fields.iter().copied() {
            let Some(index) = selected_index(field.number) else {
                output.extend_from_slice(&source[field.start..field.end]);
                continue;
            };
            seen[index] = true;
            let Some(value) = (match index {
                0 => write.exposure.map(|value| (value.to_bits(), 0u8)),
                1 => write.saturation.map(|value| (value.to_bits(), 0u8)),
                2 => write.enhance.map(|value| (u32::from(value), 1u8)),
                _ => None,
            }) else {
                continue;
            };
            output.extend_from_slice(&source[field.start..field.value_start]);
            if value.1 == 0 {
                output.extend_from_slice(&value.0.to_le_bytes());
            } else {
                append_varint(&mut output, u64::from(value.0));
            }
        }
    }
    if !seen[0] {
        if let Some(value) = write.exposure {
            append_fixed(&mut output, EXPOSURE_FIELD, value);
        }
    }
    if !seen[1] {
        if let Some(value) = write.saturation {
            append_fixed(&mut output, SATURATION_FIELD, value);
        }
    }
    if !seen[2] {
        if let Some(value) = write.enhance {
            append_varint(&mut output, u64::from(ENHANCE_FIELD << 3));
            append_varint(&mut output, u64::from(value));
        }
    }
    if output.len() != expected {
        return Err(DecodeError::projection());
    }
    Ok(output)
}

fn snapshot_matches_write(
    snapshot: ImageAdjustmentsSnapshot<'_>,
    write: ImageAdjustmentsWrite,
) -> bool {
    snapshot.exposure == write.exposure
        && snapshot.saturation == write.saturation
        && snapshot.enhance == write.enhance
}

fn append_fixed(output: &mut Vec<u8>, number: u32, value: f32) {
    append_varint(output, u64::from((number << 3) | 5));
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

fn check_options_limits(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    if requirements.fields > options.max_fields {
        return Err(DecodeError::limit(
            DecodeLimit::Fields,
            requirements.fields,
            options.max_fields,
        ));
    }
    if requirements.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limit(
            DecodeLimit::Work,
            requirements.work_bytes,
            options.max_work_bytes,
        ));
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::nesting(
            requirements.max_depth,
            options.recursion_limit,
        ));
    }
    Ok(())
}

#[cfg(test)]
#[path = "image_adjustments_codec/tests.rs"]
mod parity_tests;

#[cfg(test)]
mod tests {
    use super::*;

    fn source() -> Vec<u8> {
        vec![
            0x72, 0x0c, // ImageArchive.imageAdjustments
            0x0d, 0x00, 0x00, 0x00, 0x00, // exposure = 0
            0x15, 0x00, 0x00, 0x00, 0x00, // saturation = 0
            0x68, 0x00, // enhance = false
        ]
    }

    #[test]
    fn decodes_and_preserves_advanced_and_outer_unknowns() {
        let mut bytes = source();
        bytes[1] = 0x11; // the inserted advanced field adds five nested bytes
        bytes.splice(2..2, [0x1d, 0xcd, 0xcc, 0xcc, 0x3e]); // contrast
        bytes.extend_from_slice(&[0x98, 0x06, 0x80, 0x01]);
        let snapshot = decode_image_adjustments(&bytes, DecodeOptions::for_source(&bytes)).unwrap();
        assert_eq!(snapshot.exposure(), Some(0.0));
        assert_eq!(snapshot.saturation(), Some(0.0));
        assert_eq!(snapshot.enhance(), Some(false));
        let write = ImageAdjustmentsWrite::from_values(Some(0.25), Some(-0.5), Some(true));
        let changed = rewrite_image_adjustments(
            &bytes,
            write,
            DecodeOptions::for_source(&bytes).with_max_output_bytes(1024),
        )
        .unwrap();
        assert!(
            changed
                .windows(5)
                .any(|window| window == [0x1d, 0xcd, 0xcc, 0xcc, 0x3e])
        );
        assert!(changed.ends_with(&[0x98, 0x06, 0x80, 0x01]));
        let restored = rewrite_image_adjustments(
            &changed,
            ImageAdjustmentsWrite::from_values(Some(0.0), Some(0.0), Some(false)),
            DecodeOptions::for_source(&changed).with_max_output_bytes(1024),
        )
        .unwrap();
        assert_eq!(restored, bytes);
    }

    #[test]
    fn exact_noop_keeps_empty_and_unknown_payloads() {
        let bytes = vec![0x72, 0x03, 0x98, 0x06, 0x01];
        let output = rewrite_image_adjustments(
            &bytes,
            ImageAdjustmentsWrite::new(),
            DecodeOptions::for_source(&bytes),
        )
        .unwrap();
        assert_eq!(output, bytes);
    }

    #[test]
    fn rejects_duplicate_wrong_wire_nonfinite_and_noncanonical_bool() {
        let duplicate = [0x72, 0x0a, 0x0d, 0, 0, 0, 0, 0x0d, 0, 0, 0, 0];
        assert!(
            decode_image_adjustments(&duplicate, DecodeOptions::for_source(&duplicate)).is_err()
        );
        let wrong_wire = [0x72, 0x02, 0x08, 0x01];
        assert!(
            decode_image_adjustments(&wrong_wire, DecodeOptions::for_source(&wrong_wire)).is_err()
        );
        let nonfinite = [0x72, 0x05, 0x0d, 0, 0, 0xc0, 0x7f];
        assert!(
            decode_image_adjustments(&nonfinite, DecodeOptions::for_source(&nonfinite)).is_err()
        );
        let noncanonical_bool = [0x72, 0x03, 0x68, 0x80, 0x00];
        assert!(
            decode_image_adjustments(
                &noncanonical_bool,
                DecodeOptions::for_source(&noncanonical_bool)
            )
            .is_err()
        );
    }

    #[test]
    fn prepared_limits_are_replayed_before_emission() {
        let bytes = source();
        let prepared = prepare_image_adjustments_rewrite(
            &bytes,
            ImageAdjustmentsWrite::from_values(Some(0.25), Some(-0.5), Some(true)),
            DecodeOptions::for_source(&bytes).with_max_output_bytes(1024),
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        assert!(
            prepared
                .execute(
                    requirements
                        .exact()
                        .with_output_bytes(requirements.output_bytes - 1)
                )
                .unwrap_err()
                .limit_kind()
                .is_some()
        );
        let output = prepared.execute(requirements.exact()).unwrap();
        assert_eq!(output.report().output_bytes(), requirements.output_bytes);
        assert_eq!(output.report().fields(), requirements.fields);
    }
}
