//! Strict borrowed projection for the shared iWork chart-caption edge.
//!
//! The selected path is deliberately only
//! `TSCH.ChartDrawableArchive.super` -> `TSD.DrawableArchive.caption` ->
//! `TSP.Reference.identifier`.  The chart's extension fields, all unrelated
//! drawable fields, and all unknown source bytes remain caller-owned.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict preflight intentionally precedes the low-level wire reader it consumes."
)]

use std::fmt;

#[cfg(test)]
use std::cell::Cell;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_chart_caption_generated::LitchiIwaProjection as projection;

const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite limits for one chart-caption payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
}

impl DecodeOptions {
    /// Build an explicit finite bytes/fields/work/nesting policy.
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
        }
    }

    /// Build a conservative profile from one known source length.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(4).max(1),
            bytes.saturating_mul(8).max(1),
            8,
        )
    }

    /// Replace the aggregate candidate-output ceiling used by rewrites.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self, budget: &Budget) -> Result<Self, DecodeError> {
        if self.recursion_limit <= 1 {
            return Err(budget.nesting_limit());
        }
        Ok(Self {
            recursion_limit: self.recursion_limit - 1,
            ..self
        })
    }
}

/// Borrowed semantic facts from one chart drawable payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartCaptionSnapshot {
    drawable: Option<DrawableCaptionSnapshot>,
}

impl ChartCaptionSnapshot {
    /// Whether the optional chart drawable `super` envelope was present.
    #[must_use]
    pub const fn has_drawable(self) -> bool {
        self.drawable.is_some()
    }

    /// The nested caption object's identifier, preserving absent edges.
    #[must_use]
    pub fn caption_identifier(self) -> Option<u64> {
        match self.drawable {
            Some(drawable) => drawable.caption.map(|reference| reference.identifier),
            None => None,
        }
    }
}

/// Requested identifier for one chart-caption reference rewrite.
///
/// The write is deliberately limited to the selected reference edge. It does
/// not locate a chart object, inspect the caption graph, or update archive
/// metadata; those invariants remain owned by the caller.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartCaptionWrite {
    identifier: u64,
}

impl ChartCaptionWrite {
    /// Build a reference-identifier update.
    #[must_use]
    pub const fn new(identifier: u64) -> Self {
        Self { identifier }
    }

    /// Return the requested caption-object identifier.
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
}

/// Exact input/output accounting for one successful chart-caption rewrite.
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
    /// Source payload bytes inspected before the rewrite.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Candidate payload bytes produced by the rewrite.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Strict field visits across source, sizing, emission, and readback.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded work charged by the complete rewrite transaction.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum protobuf nesting depth observed by the strict scanner.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Output-buffer allocation events owned by this handwritten codec.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Candidate bytes retained by the returned output.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary scratch bytes allocated by this codec.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether the selected reference identifier changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// Exact finite consumption of one strict chart-caption decode.
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
    /// Source payload bytes inspected by the decoder.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Alias for callers that use the source terminology.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.input_bytes
    }

    /// Strict field visits, including unknown fields and group contents.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate bounded work charged by the decoder.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum protobuf nesting depth observed by the strict scanner.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Handwritten output allocations performed by the decoder.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Borrowed source bytes retained by the snapshot contract.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Temporary scratch bytes allocated by the decoder.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DrawableCaptionSnapshot {
    caption: Option<ReferenceSnapshot>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

/// A byte or nesting resource classification for [`DecodeError`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// The source or configured Buffa message-byte ceiling was exceeded.
    Bytes {
        /// Observed source/configured bytes.
        observed: usize,
        /// Applied byte ceiling.
        maximum: usize,
    },
    /// The configured or traversed protobuf nesting ceiling was exceeded.
    Nesting {
        /// Observed configured/depth value.
        observed: u32,
        /// Applied nesting ceiling.
        maximum: u32,
    },
}

/// Failure from strict chart-caption preflight or the private Buffa view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(WireResourceLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    OutputLimit { observed: usize, maximum: usize },
    Allocation { amount: usize },
    Projection,
}

impl DecodeError {
    /// Return the missing required schema field, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            _ => None,
        }
    }

    /// Return the duplicated singular schema field, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return the canonical-wire failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    /// Return exact field-limit observations, when applicable.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::FieldLimit { observed, maximum } => Some((observed, maximum)),
            _ => None,
        }
    }

    /// Return exact work-limit observations, when applicable.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::WorkLimit { observed, maximum } => Some((observed, maximum)),
            _ => None,
        }
    }

    /// Return the exact candidate-output limit observation, when applicable.
    #[must_use]
    pub const fn output_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::OutputLimit { observed, maximum } => Some((observed, maximum)),
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

    /// Return the exact byte/nesting resource failure, when applicable.
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        match self.kind {
            DecodeErrorKind::Resource(limit) => Some(limit),
            _ => None,
        }
    }

    const fn missing_required(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
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

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn output_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::OutputLimit { observed, maximum },
        }
    }

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(WireResourceLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "iWork chart-caption projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(WireResourceLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "iWork chart-caption projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "iWork chart-caption projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "iWork chart-caption projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::OutputLimit { observed, maximum } => write!(
                formatter,
                "iWork chart-caption rewrite produced {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate iWork chart-caption output for {amount} bytes"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "iWork chart-caption strict preflight disagrees with the Buffa projection",
            ),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        match error {
            buffa::DecodeError::MessageTooLarge => Self {
                kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                    observed: 0,
                    maximum: 0,
                }),
            },
            buffa::DecodeError::RecursionLimitExceeded => Self {
                kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
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

/// Decode only the optional chart-caption identifier.
pub fn decode_chart_caption_identifier(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<u64>, DecodeError> {
    Ok(decode_chart_caption(source, options)?.caption_identifier())
}

/// Strictly decode the bounded chart-caption projection.
///
/// The raw preflight runs before Buffa and is the resource/presence authority.
/// Buffa is then forced only for the selected singular envelopes, and its
/// borrowed scalar snapshot must agree with the strict result.
pub fn decode_chart_caption(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ChartCaptionSnapshot, DecodeError> {
    Ok(decode_chart_caption_with_report(source, options)?.0)
}

/// Strictly decode the bounded projection and return exact resource usage.
pub fn decode_chart_caption_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ChartCaptionSnapshot, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = decode_chart_caption_with_budget(source, options, &mut budget)?;
    Ok((snapshot, budget.decode_report()))
}

/// Decode one chart-caption payload while charging an existing aggregate
/// budget. Rewrites use this entry point for both their source and readback
/// passes so a work/field ceiling applies to the complete transaction rather
/// than independently to each pass.
fn decode_chart_caption_with_budget(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartCaptionSnapshot, DecodeError> {
    let strict = preflight_chart_caption(source, options, budget)?;
    let view: projection::ChartDrawableArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let drawable = view
        .super_
        .get()
        .map_err(DecodeError::from)?
        .map(|drawable| {
            let caption = drawable
                .caption
                .get()
                .map_err(DecodeError::from)?
                .map(|reference| {
                    if !reference.has_identifier() {
                        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
                    }
                    Ok(ReferenceSnapshot {
                        identifier: reference.identifier,
                        // The generated sidecar deliberately projects only
                        // `identifier`; selected deprecated fields are
                        // validated by the strict raw pass above and remain
                        // source-authoritative.
                        deprecated_type: None,
                        deprecated_is_external: None,
                    })
                })
                .transpose()?;
            Ok::<_, DecodeError>(DrawableCaptionSnapshot { caption })
        })
        .transpose()?;
    let projected = ChartCaptionSnapshot { drawable };
    if !same_projected_edge(projected, strict) {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn same_projected_edge(projected: ChartCaptionSnapshot, strict: ChartCaptionSnapshot) -> bool {
    match (projected.drawable, strict.drawable) {
        (None, None) => true,
        (Some(projected), Some(strict)) => match (projected.caption, strict.caption) {
            (None, None) => true,
            (Some(projected), Some(strict)) => projected.identifier == strict.identifier,
            _ => false,
        },
        _ => false,
    }
}

/// Rewrite the selected chart-caption reference identifier.
///
/// The complete source is strictly decoded before any write is considered.
/// Existing unknown fields, selected-envelope ordering, and all unrelated
/// bytes are copied byte-for-byte; only the identifier varint and the enclosing
/// length prefixes can change. A strict decode of the candidate output is
/// performed before it is returned.
pub fn rewrite_chart_caption(
    source: &[u8],
    write: ChartCaptionWrite,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_chart_caption_with_report(source, write, options)?.0)
}

/// Rewrite the selected chart-caption edge and return exact source/output
/// accounting.
pub fn rewrite_chart_caption_with_report(
    source: &[u8],
    write: ChartCaptionWrite,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_decode_input(source, options)?;
    // One budget spans source validation, sizing, emission, and candidate
    // readback so a rewrite cannot multiply the configured work ceiling by
    // the number of internal passes.
    let mut budget = Budget::new(source, options);
    let current = decode_chart_caption_with_budget(source, options, &mut budget)?;
    let current_identifier = current
        .caption_identifier()
        .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?;
    if current_identifier == write.identifier {
        if source.len() > options.max_output_bytes {
            return Err(DecodeError::output_limit(
                source.len(),
                options.max_output_bytes,
            ));
        }
        let output = clone_output(source)?;
        budget.record_allocation(output.len());
        return Ok((
            output,
            budget.rewrite_report(source.len(), source.len(), false),
        ));
    }

    let output_bytes = measure_rewrite_root(source, options, write, &mut budget)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::output_limit(
            output_bytes,
            options.max_output_bytes,
        ));
    }

    // `measure_rewrite_root` accounts for the sizing pass.  The actual
    // emission has a slightly different traversal shape (the selected
    // nested messages are measured immediately before they are emitted), and
    // the candidate readback is a further full traversal.  Meter both shapes
    // before reserving the output so field/work/nesting failures cannot occur
    // after the sole candidate allocation.
    let shape = measure_rewrite_path_lengths(source, options, write, &mut budget)?;
    preflight_rewrite_pass(source, options, write, &mut budget)?;
    let readback_options = DecodeOptions {
        max_message_bytes: options.max_message_bytes.max(output_bytes),
        max_output_bytes: options.max_output_bytes.max(output_bytes),
        ..options
    };
    preflight_candidate_readback(source, readback_options, shape, &mut budget)?;

    let mut output = reserve_output(output_bytes)?;
    budget.record_allocation(output_bytes);
    let mut emission_budget = Budget::unlimited(source, options);
    rewrite_root_into(source, options, write, &mut emission_budget, &mut output)?;
    debug_assert_eq!(output.len(), output_bytes);

    validate_decode_input(&output, readback_options)?;
    let mut readback_budget = Budget::unlimited(&output, readback_options);
    let readback =
        decode_chart_caption_with_budget(&output, readback_options, &mut readback_budget)?;
    if readback.caption_identifier() != Some(write.identifier) {
        return Err(DecodeError::projection());
    }
    Ok((
        output,
        budget.rewrite_report(source.len(), output_bytes, true),
    ))
}

fn measure_rewrite_root(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len(), 1)?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_super = false;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget, 1)?;
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
        let replacement = if field.number == CHART_DRAWABLE_SUPER_FIELD {
            if saw_super {
                return Err(DecodeError::duplicate_singular(
                    "TSCH.ChartDrawableArchive.super",
                ));
            }
            saw_super = true;
            let nested = field.length_delimited()?;
            let nested_bytes = measure_rewrite_drawable(nested, nested_options, write, budget, 2)?;
            length_delimited_field_len(CHART_DRAWABLE_SUPER_FIELD, nested_bytes)
        } else {
            end - start
        };
        output_bytes = checked_output_add(output_bytes, replacement, options)?;
    }
    if !saw_super {
        return Err(DecodeError::missing_required(
            "TSCH.ChartDrawableArchive.super",
        ));
    }
    Ok(output_bytes)
}

fn measure_rewrite_drawable(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    depth: u32,
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len(), depth)?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_caption = false;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget, depth)?;
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
        let replacement = if field.number == DRAWABLE_CAPTION_FIELD {
            if saw_caption {
                return Err(DecodeError::duplicate_singular(
                    "TSD.DrawableArchive.caption",
                ));
            }
            saw_caption = true;
            let nested = field.length_delimited()?;
            let nested_bytes =
                measure_rewrite_reference(nested, nested_options, write, budget, depth + 1)?;
            length_delimited_field_len(DRAWABLE_CAPTION_FIELD, nested_bytes)
        } else {
            end - start
        };
        output_bytes = checked_output_add(output_bytes, replacement, options)?;
    }
    if !saw_caption {
        return Err(DecodeError::missing_required("TSD.DrawableArchive.caption"));
    }
    Ok(output_bytes)
}

fn measure_rewrite_reference(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    depth: u32,
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len(), depth)?;
    let mut remaining = source;
    let mut saw_identifier = false;
    let mut saw_deprecated_type = false;
    let mut saw_deprecated_is_external = false;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget, depth)?;
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
            REFERENCE_IDENTIFIER_FIELD => {
                if saw_identifier {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                saw_identifier = true;
                field.varint()?;
                varint_field_len(REFERENCE_IDENTIFIER_FIELD, write.identifier)
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if saw_deprecated_type {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                saw_deprecated_type = true;
                require_canonical_int32(field.varint()?)?;
                end - start
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if saw_deprecated_is_external {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                saw_deprecated_is_external = true;
                require_canonical_bool(field.varint()?)?;
                end - start
            },
            _ => end - start,
        };
        output_bytes = checked_output_add(output_bytes, replacement, options)?;
    }
    if !saw_identifier {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(output_bytes)
}

#[derive(Clone, Copy, Debug)]
struct RewritePathLengths {
    root: usize,
    drawable: usize,
    reference: usize,
}

fn measure_rewrite_path_lengths(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
) -> Result<RewritePathLengths, DecodeError> {
    // This pass is used to model the path lengths consumed by candidate
    // readback.  It is still real source traversal, so it must use the
    // transaction budget rather than an unlimited detached budget; otherwise
    // exact report replay would under-account fields/work and limits could be
    // exceeded after the output allocation.
    let root = measure_rewrite_root(source, options, write, budget)?;
    let drawable_source = selected_payload(
        source,
        options,
        CHART_DRAWABLE_SUPER_FIELD,
        "TSCH.ChartDrawableArchive.super",
        budget,
        1,
    )?;
    let drawable_options = options.descend(budget)?;
    let drawable = measure_rewrite_drawable(drawable_source, drawable_options, write, budget, 2)?;
    let reference_source = selected_payload(
        drawable_source,
        drawable_options,
        DRAWABLE_CAPTION_FIELD,
        "TSD.DrawableArchive.caption",
        budget,
        2,
    )?;
    let reference_options = drawable_options.descend(budget)?;
    let reference =
        measure_rewrite_reference(reference_source, reference_options, write, budget, 3)?;
    Ok(RewritePathLengths {
        root,
        drawable,
        reference,
    })
}

fn selected_payload<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    selected_field: u32,
    selected_name: &'static str,
    budget: &mut Budget,
    depth: u32,
) -> Result<&'source [u8], DecodeError> {
    // Selecting the nested source is another complete traversal of this
    // message. Charge its bytes as work as well as its individual fields so
    // the rewrite report remains an exact upper bound for every sizing pass.
    budget.charge_message(source.len(), depth)?;
    let mut remaining = source;
    let mut payload = None;
    while let Some(field) =
        next_strict_field(&mut remaining, options.recursion_limit, budget, depth)?
    {
        if field.number != selected_field {
            continue;
        }
        if payload.is_some() {
            return Err(DecodeError::duplicate_singular(selected_name));
        }
        payload = Some(field.length_delimited()?);
    }
    payload.ok_or_else(|| DecodeError::missing_required(selected_name))
}

fn preflight_rewrite_pass(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len(), 1)?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_super = false;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget, 1)? {
        if field.number != CHART_DRAWABLE_SUPER_FIELD {
            continue;
        }
        if saw_super {
            return Err(DecodeError::duplicate_singular(
                "TSCH.ChartDrawableArchive.super",
            ));
        }
        saw_super = true;
        let nested = field.length_delimited()?;
        // This is the sizing sub-pass performed by rewrite_root_into before
        // it delegates to rewrite_drawable_into.
        measure_rewrite_drawable(nested, nested_options, write, budget, 2)?;
        preflight_rewrite_drawable_pass(nested, nested_options, write, budget, 2)?;
    }
    if !saw_super {
        return Err(DecodeError::missing_required(
            "TSCH.ChartDrawableArchive.super",
        ));
    }
    Ok(())
}

fn preflight_rewrite_drawable_pass(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len(), depth)?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_caption = false;
    while let Some(field) =
        next_strict_field(&mut remaining, options.recursion_limit, budget, depth)?
    {
        if field.number != DRAWABLE_CAPTION_FIELD {
            continue;
        }
        if saw_caption {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.caption",
            ));
        }
        saw_caption = true;
        let nested = field.length_delimited()?;
        measure_rewrite_reference(nested, nested_options, write, budget, depth + 1)?;
        preflight_rewrite_reference_pass(nested, nested_options, budget, depth + 1)?;
    }
    if !saw_caption {
        return Err(DecodeError::missing_required("TSD.DrawableArchive.caption"));
    }
    Ok(())
}

fn preflight_rewrite_reference_pass(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len(), depth)?;
    let mut remaining = source;
    let mut saw_identifier = false;
    let mut saw_deprecated_type = false;
    let mut saw_deprecated_is_external = false;
    while let Some(field) =
        next_strict_field(&mut remaining, options.recursion_limit, budget, depth)?
    {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if saw_identifier {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                saw_identifier = true;
                field.varint()?;
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if saw_deprecated_type {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                saw_deprecated_type = true;
                require_canonical_int32(field.varint()?)?;
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if saw_deprecated_is_external {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                saw_deprecated_is_external = true;
                require_canonical_bool(field.varint()?)?;
            },
            _ => {},
        }
    }
    if !saw_identifier {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(())
}

fn preflight_candidate_readback(
    source: &[u8],
    options: DecodeOptions,
    shape: RewritePathLengths,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_message(shape.root, 1)?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_super = false;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget, 1)? {
        if field.number != CHART_DRAWABLE_SUPER_FIELD {
            continue;
        }
        if saw_super {
            return Err(DecodeError::duplicate_singular(
                "TSCH.ChartDrawableArchive.super",
            ));
        }
        saw_super = true;
        let nested = field.length_delimited()?;
        preflight_candidate_drawable(nested, nested_options, shape, budget, 2)?;
    }
    if !saw_super {
        return Err(DecodeError::missing_required(
            "TSCH.ChartDrawableArchive.super",
        ));
    }
    Ok(())
}

fn preflight_candidate_drawable(
    source: &[u8],
    options: DecodeOptions,
    shape: RewritePathLengths,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge_message(shape.drawable, depth)?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_caption = false;
    while let Some(field) =
        next_strict_field(&mut remaining, options.recursion_limit, budget, depth)?
    {
        if field.number != DRAWABLE_CAPTION_FIELD {
            continue;
        }
        if saw_caption {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.caption",
            ));
        }
        saw_caption = true;
        let nested = field.length_delimited()?;
        preflight_candidate_reference(nested, nested_options, shape, budget, depth + 1)?;
    }
    if !saw_caption {
        return Err(DecodeError::missing_required("TSD.DrawableArchive.caption"));
    }
    Ok(())
}

fn preflight_candidate_reference(
    source: &[u8],
    options: DecodeOptions,
    shape: RewritePathLengths,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge_message(shape.reference, depth)?;
    let mut remaining = source;
    let mut saw_identifier = false;
    let mut saw_deprecated_type = false;
    let mut saw_deprecated_is_external = false;
    while let Some(field) =
        next_strict_field(&mut remaining, options.recursion_limit, budget, depth)?
    {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if saw_identifier {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                saw_identifier = true;
                field.varint()?;
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if saw_deprecated_type {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                saw_deprecated_type = true;
                require_canonical_int32(field.varint()?)?;
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if saw_deprecated_is_external {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                saw_deprecated_is_external = true;
                require_canonical_bool(field.varint()?)?;
            },
            _ => {},
        }
    }
    if !saw_identifier {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(())
}

fn rewrite_root_into(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let nested_options = options.descend(budget)?;
    rewrite_message_fields(
        source,
        options,
        budget,
        output,
        CHART_DRAWABLE_SUPER_FIELD,
        nested_options,
        write,
        rewrite_drawable_into,
        "TSCH.ChartDrawableArchive.super",
        1,
    )
}

fn rewrite_drawable_into(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let nested_options = options.descend(budget)?;
    rewrite_message_fields(
        source,
        options,
        budget,
        output,
        DRAWABLE_CAPTION_FIELD,
        nested_options,
        write,
        rewrite_reference_into,
        "TSD.DrawableArchive.caption",
        2,
    )
}

fn rewrite_reference_into(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len(), 3)?;
    let mut remaining = source;
    let mut saw_identifier = false;
    let mut saw_deprecated_type = false;
    let mut saw_deprecated_is_external = false;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget, 3)?;
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
            REFERENCE_IDENTIFIER_FIELD => {
                if saw_identifier {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                saw_identifier = true;
                field.varint()?;
                append_varint_field(output, REFERENCE_IDENTIFIER_FIELD, write.identifier);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if saw_deprecated_type {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                saw_deprecated_type = true;
                require_canonical_int32(field.varint()?)?;
                output.extend_from_slice(&source[start..end]);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if saw_deprecated_is_external {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                saw_deprecated_is_external = true;
                require_canonical_bool(field.varint()?)?;
                output.extend_from_slice(&source[start..end]);
            },
            _ => output.extend_from_slice(&source[start..end]),
        }
    }
    if !saw_identifier {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(())
}

fn rewrite_message_fields(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
    selected_field: u32,
    nested_options: DecodeOptions,
    write: ChartCaptionWrite,
    nested_rewrite: fn(
        &[u8],
        DecodeOptions,
        ChartCaptionWrite,
        &mut Budget,
        &mut Vec<u8>,
    ) -> Result<(), DecodeError>,
    selected_name: &'static str,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len(), depth)?;
    let mut remaining = source;
    let mut saw_selected = false;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget, depth)?;
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
        if field.number != selected_field {
            output.extend_from_slice(&source[start..end]);
            continue;
        }
        if saw_selected {
            return Err(DecodeError::duplicate_singular(selected_name));
        }
        saw_selected = true;
        let nested = field.length_delimited()?;
        let nested_len = match selected_field {
            CHART_DRAWABLE_SUPER_FIELD => {
                measure_rewrite_drawable(nested, nested_options, write, budget, depth + 1)?
            },
            DRAWABLE_CAPTION_FIELD => {
                measure_rewrite_reference(nested, nested_options, write, budget, depth + 1)?
            },
            _ => return Err(DecodeError::projection()),
        };
        append_length_delimited_field_header(output, selected_field, nested_len);
        let before = output.len();
        nested_rewrite(nested, nested_options, write, budget, output)?;
        debug_assert_eq!(output.len() - before, nested_len);
    }
    if !saw_selected {
        return Err(DecodeError::missing_required(selected_name));
    }
    Ok(())
}

fn checked_output_add(
    current: usize,
    additional: usize,
    options: DecodeOptions,
) -> Result<usize, DecodeError> {
    let total = current
        .checked_add(additional)
        .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
    if total > options.max_output_bytes {
        return Err(DecodeError::output_limit(total, options.max_output_bytes));
    }
    Ok(total)
}

fn length_delimited_field_len(number: u32, payload_len: usize) -> usize {
    varint_len((u64::from(number) << 3) | 2) + varint_len(payload_len as u64) + payload_len
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn reserve_output(amount: usize) -> Result<Vec<u8>, DecodeError> {
    #[cfg(test)]
    OUTPUT_ALLOCATIONS.with(|count| count.set(count.get().saturating_add(1)));
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_allocation_error| DecodeError::allocation(amount))?;
    if output.capacity() != amount {
        return Err(DecodeError::allocation(amount));
    }
    Ok(output)
}

#[cfg(test)]
thread_local! {
    static OUTPUT_ALLOCATIONS: Cell<usize> = const { Cell::new(0) };
}

#[cfg(test)]
fn reset_output_allocations() {
    OUTPUT_ALLOCATIONS.with(|count| count.set(0));
}

#[cfg(test)]
fn output_allocations() -> usize {
    OUTPUT_ALLOCATIONS.with(Cell::get)
}

fn clone_output(source: &[u8]) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}

fn append_length_delimited_field_header(output: &mut Vec<u8>, number: u32, payload_len: usize) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, payload_len as u64);
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    append_varint(output, u64::from(number) << 3);
    append_varint(output, value);
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
    if value > 1 {
        return Err(DecodeError::noncanonical("bool scalar is not zero or one"));
    }
    Ok(value == 1)
}

fn require_canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "The strict range check proves this is a canonical int32 representation."
    )]
    Ok(value as i32)
}

#[derive(Debug)]
struct Budget {
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: u32,
    max_depth: u32,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl Budget {
    const fn new(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_nesting: options.recursion_limit,
            max_depth: 0,
            allocations: 0,
            retained_bytes: source.len(),
            scratch_bytes: 0,
        }
    }

    const fn unlimited(source: &[u8], options: DecodeOptions) -> Self {
        Self {
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_fields: usize::MAX,
            max_work_bytes: usize::MAX,
            max_nesting: options.recursion_limit,
            max_depth: 0,
            allocations: 0,
            retained_bytes: source.len(),
            scratch_bytes: 0,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError {
                kind: DecodeErrorKind::FieldLimit {
                    observed,
                    maximum: self.max_fields,
                },
            });
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        self.observe_depth(depth)?;
        let observed = self.work_bytes.saturating_add(bytes.saturating_mul(2));
        if observed > self.max_work_bytes {
            return Err(DecodeError {
                kind: DecodeErrorKind::WorkLimit {
                    observed,
                    maximum: self.max_work_bytes,
                },
            });
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.max_nesting {
            return Err(self.nesting_limit_at(depth));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn nesting_limit(&self) -> DecodeError {
        DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: self.max_nesting.saturating_add(1),
                maximum: self.max_nesting,
            }),
        }
    }

    const fn nesting_limit_at(&self, observed: u32) -> DecodeError {
        DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed,
                maximum: self.max_nesting,
            }),
        }
    }

    fn record_allocation(&mut self, amount: usize) {
        self.allocations = self.allocations.saturating_add(1);
        self.retained_bytes = amount;
    }

    fn depth_for(&self, options: DecodeOptions) -> u32 {
        self.max_nesting
            .saturating_sub(options.recursion_limit)
            .saturating_add(1)
    }

    const fn decode_report(&self) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
        }
    }

    const fn rewrite_report(
        &self,
        input_bytes: usize,
        output_bytes: usize,
        changed: bool,
    ) -> RewriteReport {
        RewriteReport {
            input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            allocations: self.allocations,
            retained_bytes: self.retained_bytes,
            scratch_bytes: self.scratch_bytes,
            changed,
        }
    }
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError {
        kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: usize::MAX,
        }),
    })?;
    if options.max_message_bytes > hard_maximum {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_maximum,
            }),
        });
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }),
        });
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }),
        });
    }
    Ok(())
}

fn preflight_chart_caption(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartCaptionSnapshot, DecodeError> {
    budget.charge_message(source.len(), budget.depth_for(options))?;
    let nested_options = options.descend(budget)?;
    let mut drawable = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget, 1)? {
        if field.number != CHART_DRAWABLE_SUPER_FIELD {
            continue;
        }
        if drawable.is_some() {
            return Err(DecodeError::duplicate_singular(
                "TSCH.ChartDrawableArchive.super",
            ));
        }
        drawable = Some(preflight_drawable(
            field.length_delimited()?,
            nested_options,
            budget,
        )?);
    }
    Ok(ChartCaptionSnapshot { drawable })
}

fn preflight_drawable(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<DrawableCaptionSnapshot, DecodeError> {
    budget.charge_message(source.len(), budget.depth_for(options))?;
    let nested_options = options.descend(budget)?;
    let mut caption = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        budget,
        budget.depth_for(options),
    )? {
        if field.number != DRAWABLE_CAPTION_FIELD {
            continue;
        }
        if caption.is_some() {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.caption",
            ));
        }
        caption = Some(preflight_reference(
            field.length_delimited()?,
            nested_options,
            budget,
        )?);
    }
    Ok(DrawableCaptionSnapshot { caption })
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.charge_message(source.len(), budget.depth_for(options))?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        budget,
        budget.depth_for(options),
    )? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                identifier = Some(field.varint()?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                deprecated_type = Some(require_canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                deprecated_is_external = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(ReferenceSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
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
    canonical_key: bool,
    canonical_value: bool,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
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
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        let StrictValue::Varint(value) = self.value else {
            return Err(DecodeError::projection());
        };
        Ok(value)
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
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
    recursion_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let (value, canonical_value) = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            (StrictValue::Varint(value), canonical)
        },
        buffa::encoding::WireType::Fixed64 => {
            take_exact(source, 8)?;
            (StrictValue::Fixed64, true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                canonical,
            )
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(|| budget.nesting_limit())?;
            let group_depth = depth.saturating_add(1);
            budget.observe_depth(group_depth)?;
            skip_strict_group(source, field_number, child_limit, budget, group_depth)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            (StrictValue::Fixed32, true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
        canonical_key,
        canonical_value,
    })))
}

fn skip_strict_group(
    source: &mut &[u8],
    expected_field_number: u32,
    recursion_limit: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, budget, depth)? {
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
    reason = "Focused negative tests use explicit panic messages and reuse local error roles."
)]
mod tests {
    use super::{
        CHART_DRAWABLE_SUPER_FIELD, ChartCaptionSnapshot, ChartCaptionWrite,
        DRAWABLE_CAPTION_FIELD, DecodeOptions, WireResourceLimit, decode_chart_caption,
        decode_chart_caption_identifier, output_allocations, reserve_output,
        reset_output_allocations, rewrite_chart_caption_with_report,
    };

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            source.len().saturating_mul(4).max(1),
            source.len().saturating_mul(8).max(1),
            8,
        )
    }

    fn chart_with_caption(identifier: u64) -> Vec<u8> {
        let reference = [vec![0x08], varint(identifier)].concat();
        chart_with_reference(&reference)
    }

    fn chart_with_reference(reference: &[u8]) -> Vec<u8> {
        let drawable = [
            vec![0x5a],
            varint(reference.len() as u64),
            reference.to_vec(),
        ]
        .concat();
        [vec![0x0a], varint(drawable.len() as u64), drawable].concat()
    }

    fn aggregate_rewrite_work(source: &[u8], write: ChartCaptionWrite) -> (usize, usize) {
        let options = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 8)
            .with_max_output_bytes(source.len() + 32);
        let (_, report) =
            rewrite_chart_caption_with_report(source, write, options).expect("permissive rewrite");
        (report.output_bytes(), report.work_bytes())
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

    fn field_varint_raw(number: u32, value: &[u8]) -> Vec<u8> {
        [varint(u64::from(number) << 3), value.to_vec()].concat()
    }

    fn field_bytes(number: u32, value: &[u8]) -> Vec<u8> {
        [
            varint((u64::from(number) << 3) | 2),
            varint(value.len() as u64),
            value.to_vec(),
        ]
        .concat()
    }

    #[test]
    fn selected_caption_edge_matches_borrowed_projection() {
        let source = chart_with_caption(42);
        let snapshot = decode_chart_caption(&source, options(&source)).expect("caption");
        assert!(snapshot.has_drawable());
        assert_eq!(snapshot.caption_identifier(), Some(42));
        assert_eq!(
            decode_chart_caption_identifier(&source, options(&source)),
            Ok(Some(42))
        );
    }

    #[test]
    fn absent_super_and_caption_edges_remain_absent() {
        let source = Vec::new();
        let snapshot = decode_chart_caption(&source, options(&source)).expect("absence");
        assert_eq!(snapshot, ChartCaptionSnapshot { drawable: None });
        assert_eq!(snapshot.caption_identifier(), None);
        let drawable_without_caption = vec![0x0a, 0x00];
        let snapshot = decode_chart_caption(
            &drawable_without_caption,
            options(&drawable_without_caption),
        )
        .expect("empty drawable");
        assert!(snapshot.has_drawable());
        assert_eq!(
            decode_chart_caption_identifier(
                &drawable_without_caption,
                options(&drawable_without_caption)
            ),
            Ok(None)
        );
    }

    #[test]
    fn unknown_chart_extension_and_drawable_fields_are_not_materialized() {
        let mut source = chart_with_caption(42);
        source.extend([0x82, 0x01, 0x01, 0xff]); // unknown chart extension field 16
        let before = source.clone();
        assert_eq!(
            decode_chart_caption_identifier(&source, options(&source)).expect("unknowns"),
            Some(42)
        );
        assert_eq!(source, before);
    }

    #[test]
    fn unknown_overlong_scalar_values_are_accepted_and_retained_byte_for_byte() {
        // Field 16's key is canonical; its value 1 is deliberately encoded
        // with one redundant continuation byte. Unknown scalar values are
        // source-authoritative, while selected known values remain strict.
        let unknown = field_varint_raw(16, &[0x81, 0x00]);
        let mut source = chart_with_caption(7);
        source.extend_from_slice(&unknown);
        let rewrite_options =
            DecodeOptions::new(source.len(), source.len().saturating_mul(4), usize::MAX, 8)
                .with_max_output_bytes(source.len() + 32);
        assert_eq!(
            decode_chart_caption_identifier(&source, options(&source)),
            Ok(Some(7))
        );
        let (rewritten, _) = rewrite_chart_caption_with_report(
            &source,
            ChartCaptionWrite::new(300),
            rewrite_options,
        )
        .expect("unknown scalar value remains source-authoritative");
        assert!(
            rewritten
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        assert_eq!(
            decode_chart_caption_identifier(&rewritten, options(&rewritten)),
            Ok(Some(300))
        );
    }

    #[test]
    fn selected_known_scalar_values_remain_strict() {
        let source = chart_with_reference(&[0x08, 0x81, 0x00]);
        let error = decode_chart_caption(&source, options(&source))
            .expect_err("selected identifier overlong value");
        assert_eq!(error.noncanonical_reason(), Some("protobuf varint value"));
    }

    #[test]
    fn unknown_group_depth_is_reported_and_exactly_bounded() {
        let unknown_group = [0x9b, 0x03, 0x08, 0x01, 0x9c, 0x03];
        let reference = [unknown_group.to_vec(), field_varint(1, 7)].concat();
        let source = chart_with_reference(&reference);
        let (_, report) = super::decode_chart_caption_with_report(
            &source,
            DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 4),
        )
        .expect("group at the exact nesting boundary");
        assert_eq!(report.max_depth(), 4);

        let exact = DecodeOptions::new(source.len(), report.fields(), report.work_bytes(), 4);
        let (_, replay) = super::decode_chart_caption_with_report(&source, exact)
            .expect("exact group report replay");
        assert_eq!(replay.fields(), report.fields());
        assert_eq!(replay.work_bytes(), report.work_bytes());
        assert_eq!(replay.max_depth(), report.max_depth());

        let below = DecodeOptions::new(source.len(), report.fields(), report.work_bytes(), 3);
        let error = super::decode_chart_caption_with_report(&source, below)
            .expect_err("one level below group boundary");
        assert_eq!(
            error.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 4,
                maximum: 3,
            })
        );
    }

    #[test]
    fn malformed_selected_envelopes_are_rejected() {
        let malformed = [
            vec![0x0a, 0x01, 0x5a],             // truncated drawable caption
            vec![0x0a, 0x02, 0x5a, 0x00],       // missing required identifier
            vec![0x0a, 0x03, 0x5a, 0x02, 0x08], // truncated reference
            vec![0x0a, 0x05, 0x5a, 0x03, 0x08, 0x01, 0x08, 0x02], // duplicate id
            vec![0x0a, 0x02, 0x58, 0x01],       // wrong wire for caption
            vec![0x0a, 0x00, 0x0a, 0x00],       // duplicate super
        ];
        for source in malformed {
            assert!(
                decode_chart_caption(&source, options(&source)).is_err(),
                "{source:?}"
            );
        }
    }

    #[test]
    fn canonicality_and_limits_fail_before_projection() {
        let source = chart_with_caption(42);
        assert_eq!(
            decode_chart_caption(
                &source,
                DecodeOptions::new(source.len() - 1, 32, source.len() * 8, 8),
            )
            .expect_err("byte cap")
            .wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );
        assert_eq!(
            decode_chart_caption(
                &source,
                DecodeOptions::new(source.len(), 2, source.len() * 8, 8),
            )
            .expect_err("field cap")
            .field_limit_values(),
            Some((3, 2))
        );
        let work = source.len() * 2 + 1;
        assert_eq!(
            decode_chart_caption(&source, DecodeOptions::new(source.len(), 32, work, 8),)
                .expect_err("work cap")
                .work_limit_values(),
            Some((20, work))
        );
        let noncanonical = vec![0x0a, 0x80, 0x00];
        assert_eq!(
            decode_chart_caption(&noncanonical, options(&noncanonical))
                .expect_err("noncanonical length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );
        assert_eq!(
            decode_chart_caption(
                &source,
                DecodeOptions::new(source.len(), 32, source.len() * 8, 0),
            )
            .expect_err("nesting cap")
            .wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 0,
                maximum: 64,
            })
        );
    }

    #[test]
    fn source_bytes_are_immutable_on_success_and_refusal() {
        let source = chart_with_caption(7);
        let before = source.clone();
        let _ = decode_chart_caption(&source, options(&source));
        assert_eq!(source, before);
        let error = decode_chart_caption(
            &source,
            DecodeOptions::new(source.len() - 1, 32, source.len() * 8, 8),
        )
        .expect_err("refusal");
        assert!(error.wire_resource_limit().is_some());
        assert_eq!(source, before);
    }

    #[test]
    fn rewrite_changes_only_identifier_and_preserves_unknown_spans() {
        let reference = [
            field_varint(99, 1),
            field_varint(1, 7),
            field_varint(2, 1),
            field_varint(1000, 42),
        ]
        .concat();
        let drawable = [
            field_varint(2, 9),
            field_bytes(DRAWABLE_CAPTION_FIELD, &reference),
            field_varint(13, 1),
        ]
        .concat();
        let source = [
            field_varint(16, 3),
            field_bytes(CHART_DRAWABLE_SUPER_FIELD, &drawable),
            field_varint(17, 4),
        ]
        .concat();
        // Rewrites charge source validation, sizing, emission, and candidate
        // readback through one aggregate budget. The decode helper's
        // source-relative work profile is intentionally per-pass and is too
        // small for this nested rewrite; the exact inclusive aggregate
        // boundary is covered below.
        let rewrite_options = DecodeOptions::new(
            source.len(),
            source.len().saturating_mul(4).max(1),
            usize::MAX,
            8,
        )
        .with_max_output_bytes(source.len() + 32);
        let (rewritten, report) = rewrite_chart_caption_with_report(
            &source,
            ChartCaptionWrite::new(300),
            rewrite_options,
        )
        .expect("rewrite");
        assert!(report.changed());
        assert_eq!(
            decode_chart_caption_identifier(&rewritten, options(&rewritten)),
            Ok(Some(300))
        );
        assert!(
            rewritten
                .windows(field_varint(99, 1).len())
                .any(|window| { window == field_varint(99, 1).as_slice() })
        );
        assert!(
            rewritten
                .windows(field_varint(1000, 42).len())
                .any(|window| { window == field_varint(1000, 42).as_slice() })
        );
        assert!(
            rewritten
                .windows(field_varint(16, 3).len())
                .any(|window| { window == field_varint(16, 3).as_slice() })
        );
        assert!(
            rewritten
                .windows(field_varint(17, 4).len())
                .any(|window| { window == field_varint(17, 4).as_slice() })
        );
        assert_ne!(rewritten, source);
        assert_eq!(
            decode_chart_caption_identifier(&source, options(&source)),
            Ok(Some(7))
        );
    }

    #[test]
    fn selected_reference_known_optional_fields_are_strict_but_raw_preserved() {
        let unknown_group = [0x9b, 0x03, 0x08, 0x01, 0x9c, 0x03];
        let reference = [
            field_varint(2, u64::MAX),
            field_varint(1, 7),
            field_varint(3, 1),
            unknown_group.to_vec(),
        ]
        .concat();
        let source = chart_with_reference(&reference);
        let (rewritten, report) = rewrite_chart_caption_with_report(
            &source,
            ChartCaptionWrite::new(300),
            DecodeOptions::new(source.len(), source.len() * 4, usize::MAX, 8)
                .with_max_output_bytes(source.len() + 32),
        )
        .expect("canonical optional reference fields");
        assert_eq!(
            decode_chart_caption_identifier(&source, options(&source)),
            Ok(Some(7))
        );
        assert_eq!(
            decode_chart_caption_identifier(&rewritten, options(&rewritten)),
            Ok(Some(300))
        );
        let canonical_type = field_varint(2, u64::MAX);
        assert!(
            rewritten
                .windows(canonical_type.len())
                .any(|window| window == canonical_type.as_slice())
        );
        assert!(
            rewritten
                .windows(unknown_group.len())
                .any(|window| { window == unknown_group })
        );
        assert_eq!(report.allocations(), 1);
        assert_eq!(report.retained_bytes(), rewritten.len());
        assert_eq!(report.scratch_bytes(), 0);
        assert!(report.fields() > 0);
        assert!(report.work_bytes() > 0);
    }

    #[test]
    fn selected_reference_known_optional_fields_reject_wrong_wire_duplicate_and_noncanonical() {
        let malformed = [
            chart_with_reference(&[0x12, 0x01, 0x01, 0x08, 0x01]), // deprecated_type bytes
            chart_with_reference(&[0x18, 0x02, 0x08, 0x01]),       // bool value 2
            chart_with_reference(&[0x10, 0x01, 0x10, 0x01, 0x08, 0x01]), // duplicate type
            chart_with_reference(&[0x18, 0x01, 0x18, 0x01, 0x08, 0x01]), // duplicate bool
            chart_with_reference(&[0x10, 0x81, 0x00, 0x08, 0x01]), // overlong type value
            chart_with_reference(&[0x90, 0x00, 0x01, 0x08, 0x01]), // overlong type key
        ];
        for source in malformed {
            assert!(
                decode_chart_caption(&source, options(&source)).is_err(),
                "malformed selected reference accepted: {source:?}"
            );
        }
    }

    #[test]
    fn rewrite_report_replays_exact_output_and_field_limits() {
        let source = chart_with_caption(7);
        let write = ChartCaptionWrite::new(300);
        let permissive = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 8)
            .with_max_output_bytes(source.len() + 32);
        let (expected, report) = rewrite_chart_caption_with_report(&source, write, permissive)
            .expect("permissive rewrite");
        let exact = DecodeOptions::new(source.len(), report.fields(), report.work_bytes(), 8)
            .with_max_output_bytes(report.output_bytes());
        let (actual, replay) =
            rewrite_chart_caption_with_report(&source, write, exact).expect("exact report replay");
        assert_eq!(actual, expected);
        assert_eq!(replay.output_bytes(), report.output_bytes());
        assert_eq!(replay.fields(), report.fields());
        assert_eq!(replay.work_bytes(), report.work_bytes());
        let below_fields =
            DecodeOptions::new(source.len(), report.fields() - 1, report.work_bytes(), 8)
                .with_max_output_bytes(report.output_bytes());
        assert!(
            rewrite_chart_caption_with_report(&source, write, below_fields)
                .expect_err("field budget below exact report")
                .field_limit_values()
                .is_some()
        );
    }

    #[test]
    fn rewrite_limit_failures_happen_before_output_allocation() {
        let source = chart_with_caption(7);
        let write = ChartCaptionWrite::new(300);
        let permissive = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 8)
            .with_max_output_bytes(source.len() + 32);
        let (expected, report) = rewrite_chart_caption_with_report(&source, write, permissive)
            .expect("permissive rewrite");

        reset_output_allocations();
        let (actual, _) = rewrite_chart_caption_with_report(
            &source,
            write,
            DecodeOptions::new(source.len(), report.fields(), report.work_bytes(), 8)
                .with_max_output_bytes(report.output_bytes()),
        )
        .expect("exact limits");
        assert_eq!(actual, expected);
        assert_eq!(output_allocations(), 1);

        for (label, options) in [
            (
                "fields",
                DecodeOptions::new(source.len(), report.fields() - 1, report.work_bytes(), 8)
                    .with_max_output_bytes(report.output_bytes()),
            ),
            (
                "work",
                DecodeOptions::new(source.len(), report.fields(), report.work_bytes() - 1, 8)
                    .with_max_output_bytes(report.output_bytes()),
            ),
            (
                "output",
                DecodeOptions::new(source.len(), report.fields(), report.work_bytes(), 8)
                    .with_max_output_bytes(report.output_bytes() - 1),
            ),
            (
                "depth",
                DecodeOptions::new(
                    source.len(),
                    report.fields(),
                    report.work_bytes(),
                    report.max_depth() - 1,
                )
                .with_max_output_bytes(report.output_bytes()),
            ),
        ] {
            reset_output_allocations();
            let error =
                rewrite_chart_caption_with_report(&source, write, options).expect_err(label);
            assert_eq!(output_allocations(), 0, "{label} allocated output: {error}");
        }
    }

    #[test]
    fn matching_rewrite_is_an_exact_noop_after_strict_decode() {
        let source = chart_with_caption(7);
        let options = options(&source).with_max_output_bytes(source.len());
        let (rewritten, report) =
            rewrite_chart_caption_with_report(&source, ChartCaptionWrite::new(7), options)
                .expect("no-op");
        assert_eq!(rewritten, source);
        assert!(!report.changed());
    }

    #[test]
    fn output_reservation_enforces_exact_capacity_or_typed_allocation_error() {
        match reserve_output(17) {
            Ok(output) => assert_eq!(output.capacity(), 17),
            Err(error) => assert_eq!(error.allocation_amount(), Some(17)),
        }
    }

    #[test]
    fn rewrite_work_budget_is_aggregate_and_exactly_inclusive() {
        let source = chart_with_caption(7);
        let write = ChartCaptionWrite::new(300);
        let (output_bytes, exact_work) = aggregate_rewrite_work(&source, write);

        let exact_options = DecodeOptions::new(source.len(), usize::MAX, exact_work, 8)
            .with_max_output_bytes(output_bytes);
        reset_output_allocations();
        rewrite_chart_caption_with_report(&source, write, exact_options)
            .expect("the exact aggregate work ceiling is inclusive");
        assert_eq!(output_allocations(), 1);

        let below_options = DecodeOptions::new(source.len(), usize::MAX, exact_work - 1, 8)
            .with_max_output_bytes(output_bytes);
        reset_output_allocations();
        let error = rewrite_chart_caption_with_report(&source, write, below_options)
            .expect_err("one byte below aggregate work must fail");
        assert_eq!(output_allocations(), 0);
        assert_eq!(
            error.work_limit_values(),
            Some((exact_work, exact_work - 1))
        );
    }
}
