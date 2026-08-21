//! Strict borrowed projection for the Keynote chart-caption edge.
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

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_chart_caption_generated::LitchiIwaProjection as projection;

const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;

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

    /// Whether the selected reference identifier changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct DrawableCaptionSnapshot {
    caption: Option<ReferenceSnapshot>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReferenceSnapshot {
    identifier: u64,
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
                "Keynote chart-caption projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(WireResourceLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote chart-caption projection nesting limit exceeded: observed {observed}, maximum {maximum}"
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
                "Keynote chart-caption projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-caption projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::OutputLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-caption rewrite produced {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate Keynote chart-caption output for {amount} bytes"
            ),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote chart-caption strict preflight disagrees with the Buffa projection",
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
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_chart_caption(source, options, &mut budget)?;
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
                    })
                })
                .transpose()?;
            Ok::<_, DecodeError>(DrawableCaptionSnapshot { caption })
        })
        .transpose()?;
    let projected = ChartCaptionSnapshot { drawable };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
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
    let current = decode_chart_caption(source, options)?;
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
        return Ok((
            clone_output(source)?,
            RewriteReport {
                input_bytes: source.len(),
                output_bytes: source.len(),
                changed: false,
            },
        ));
    }

    let mut measure_budget = Budget::new(options);
    let output_bytes = measure_rewrite_root(source, options, write, &mut measure_budget)?;
    if output_bytes > options.max_output_bytes {
        return Err(DecodeError::output_limit(
            output_bytes,
            options.max_output_bytes,
        ));
    }
    let mut output = reserve_output(output_bytes)?;
    let mut write_budget = Budget::new(options);
    rewrite_root_into(source, options, write, &mut write_budget, &mut output)?;
    debug_assert_eq!(output.len(), output_bytes);

    let readback_options = DecodeOptions {
        max_message_bytes: options.max_message_bytes.max(output.len()),
        max_output_bytes: options.max_output_bytes.max(output.len()),
        ..options
    };
    let readback = decode_chart_caption(&output, readback_options)?;
    if readback.caption_identifier() != Some(write.identifier) {
        return Err(DecodeError::projection());
    }
    Ok((
        output,
        RewriteReport {
            input_bytes: source.len(),
            output_bytes,
            changed: true,
        },
    ))
}

fn measure_rewrite_root(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_super = false;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
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
            let nested_bytes = measure_rewrite_drawable(nested, nested_options, write, budget)?;
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
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend(budget)?;
    let mut remaining = source;
    let mut saw_caption = false;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
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
            let nested_bytes = measure_rewrite_reference(nested, nested_options, write, budget)?;
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
) -> Result<usize, DecodeError> {
    budget.charge_message(source.len())?;
    let mut remaining = source;
    let mut saw_identifier = false;
    let mut output_bytes = 0usize;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
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
        let replacement = if field.number == REFERENCE_IDENTIFIER_FIELD {
            if saw_identifier {
                return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
            }
            saw_identifier = true;
            field.varint()?;
            varint_field_len(REFERENCE_IDENTIFIER_FIELD, write.identifier)
        } else {
            end - start
        };
        output_bytes = checked_output_add(output_bytes, replacement, options)?;
    }
    if !saw_identifier {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(output_bytes)
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
    )
}

fn rewrite_reference_into(
    source: &[u8],
    options: DecodeOptions,
    write: ChartCaptionWrite,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let mut remaining = source;
    let mut saw_identifier = false;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
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
        if field.number == REFERENCE_IDENTIFIER_FIELD {
            if saw_identifier {
                return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
            }
            saw_identifier = true;
            field.varint()?;
            append_varint_field(output, REFERENCE_IDENTIFIER_FIELD, write.identifier);
        } else {
            output.extend_from_slice(&source[start..end]);
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
) -> Result<(), DecodeError> {
    budget.charge_message(source.len())?;
    let mut remaining = source;
    let mut saw_selected = false;
    while !remaining.is_empty() {
        let start = source.len() - remaining.len();
        let item = parse_strict_field(&mut remaining, options.recursion_limit, budget)?;
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
                measure_rewrite_drawable(nested, nested_options, write, budget)?
            },
            DRAWABLE_CAPTION_FIELD => {
                measure_rewrite_reference(nested, nested_options, write, budget)?
            },
            _ => return Err(DecodeError::projection()),
        };
        let mut nested_output = reserve_output(nested_len)?;
        nested_rewrite(nested, nested_options, write, budget, &mut nested_output)?;
        debug_assert_eq!(nested_output.len(), nested_len);
        append_length_delimited_field(output, selected_field, &nested_output);
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
    let mut output = Vec::new();
    output
        .try_reserve_exact(amount)
        .map_err(|_allocation_error| DecodeError::allocation(amount))?;
    Ok(output)
}

fn clone_output(source: &[u8]) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}

fn append_length_delimited_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
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

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_nesting: options.recursion_limit,
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

    fn charge_message(&mut self, bytes: usize) -> Result<(), DecodeError> {
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

    const fn nesting_limit(&self) -> DecodeError {
        DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: self.max_nesting.saturating_add(1),
                maximum: self.max_nesting,
            }),
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
    budget.charge_message(source.len())?;
    let nested_options = options.descend(budget)?;
    let mut drawable = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
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
    budget.charge_message(source.len())?;
    let nested_options = options.descend(budget)?;
    let mut caption = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
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
    budget.charge_message(source.len())?;
    let mut identifier = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        if field.number != REFERENCE_IDENTIFIER_FIELD {
            continue;
        }
        if identifier.is_some() {
            return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
        }
        identifier = Some(field.varint()?);
    }
    Ok(ReferenceSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
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
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, budget)? {
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
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
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
            skip_strict_group(source, field_number, child_limit, budget)?;
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
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, budget)? {
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
        decode_chart_caption_identifier, rewrite_chart_caption_with_report,
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
        let drawable = [vec![0x5a], varint(reference.len() as u64), reference].concat();
        [vec![0x0a], varint(drawable.len() as u64), drawable].concat()
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
        let rewrite_options = options(&source).with_max_output_bytes(source.len() + 32);
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
    fn matching_rewrite_is_an_exact_noop_after_strict_decode() {
        let source = chart_with_caption(7);
        let options = options(&source).with_max_output_bytes(source.len());
        let (rewritten, report) =
            rewrite_chart_caption_with_report(&source, ChartCaptionWrite::new(7), options)
                .expect("no-op");
        assert_eq!(rewritten, source);
        assert!(!report.changed());
    }
}
