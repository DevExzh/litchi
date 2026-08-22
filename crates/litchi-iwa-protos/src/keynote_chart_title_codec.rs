//! Strict borrowed projection for the Keynote chart-title generated extension.
//!
//! The selected payload is only the generated `ChartNonStyleArchive` scalar
//! pair: field 21 (`bool`) and field 23 (UTF-8 `string`). The outer
//! `TSCH.ChartNonStyleArchive` envelope, chart graph, unrelated generated
//! fields, and unknown source bytes remain caller-owned. This module performs
//! only a borrowed projection and a wire-local rewrite; it never performs a
//! chart lookup or archive transaction.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict preflight intentionally precedes the low-level wire reader it consumes."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_chart_title_generated::LitchiIwaProjection as projection;

const CHART_TITLE_VISIBLE_FIELD: u32 = 21;
const CHART_TITLE_TEXT_FIELD: u32 = 23;
const ROOT_DEPTH: u32 = 1;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Finite limits for one generated chart-title extension payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_output_bytes: usize,
    max_title_bytes: usize,
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

/// Borrowed semantic facts from one generated chart-title extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartTitleSnapshot<'source> {
    title_visible: Option<bool>,
    title: Option<&'source str>,
    raw: &'source [u8],
}

impl<'source> ChartTitleSnapshot<'source> {
    /// Optional native field 21, preserving proto2 presence.
    #[must_use]
    pub const fn title_visible(self) -> Option<bool> {
        self.title_visible
    }

    /// Alias for [`Self::title_visible`].
    #[must_use]
    pub const fn show_title(self) -> Option<bool> {
        self.title_visible
    }

    /// Optional native field 23 borrowed directly from the source payload.
    #[must_use]
    pub const fn title(self) -> Option<&'source str> {
        self.title
    }

    /// Return the title only when field 21 explicitly enables it.
    ///
    /// A visible title with an absent field 23 follows the native empty-title
    /// default and returns `Some("")` without allocating.
    #[must_use]
    pub fn visible_title(self) -> Option<&'source str> {
        (self.title_visible == Some(true)).then(|| self.title.unwrap_or_default())
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

/// Requested presence-preserving values for a generated chart-title rewrite.
///
/// This is intentionally a wire-level update only. It does not locate a
/// chart, resolve a `chart_non_style` reference, or validate a chart graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ChartTitleWrite<'title> {
    title_visible: Option<bool>,
    title: Option<&'title str>,
}

impl<'title> ChartTitleWrite<'title> {
    /// Build a requested proto2 presence/value pair.
    #[must_use]
    pub const fn new(title_visible: Option<bool>, title: Option<&'title str>) -> Self {
        Self {
            title_visible,
            title,
        }
    }

    /// Requested field-21 presence/value.
    #[must_use]
    pub const fn title_visible(self) -> Option<bool> {
        self.title_visible
    }

    /// Requested field-23 UTF-8 title.
    #[must_use]
    pub const fn title(self) -> Option<&'title str> {
        self.title
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

    /// Borrowed semantic output bytes charged for the selected title.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// UTF-8 bytes in field 23, when present.
    #[must_use]
    pub const fn title_bytes(self) -> usize {
        self.title_bytes
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
    /// Field 23 title bytes exceeded the ceiling.
    Title { observed: usize, maximum: usize },
    /// Configured or traversed protobuf nesting exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
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

/// Failure from strict chart-title preflight or the private Buffa view.
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
                "Keynote chart-title projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote chart-title projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Fields { observed, maximum })
            | DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-title projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Work { observed, maximum })
            | DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-title projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Output { observed, maximum })
            | DecodeErrorKind::OutputLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-title projection produced {observed} output bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Resource(DecodeLimit::Title { observed, maximum })
            | DecodeErrorKind::TitleLimit { observed, maximum } => write!(
                formatter,
                "Keynote chart-title projection title has {observed} bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate Keynote chart-title output for {amount} bytes"
            ),
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::Projection => formatter.write_str(
                "Keynote chart-title strict preflight disagrees with the Buffa projection",
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
pub fn decode_chart_title<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ChartTitleSnapshot<'source>, DecodeError> {
    Ok(decode_chart_title_with_report(source, options)?.0)
}

/// Decode the generated extension and return exact aggregate consumption.
pub fn decode_chart_title_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(ChartTitleSnapshot<'source>, DecodeReport), DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = decode_chart_title_with_budget(source, options, &mut budget)?;
    Ok((snapshot, budget.report()))
}

/// Decode one chart-title payload while charging an existing aggregate
/// budget. Rewrites use this entry point for both their source and candidate
/// readback passes so the field/work ceilings cover the complete transaction.
fn decode_chart_title_with_budget<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartTitleSnapshot<'source>, DecodeError> {
    // Work and field counters are aggregate across a rewrite transaction,
    // while the selected title/output ceilings apply to each decode pass.
    // Reset the per-pass semantic counters before source and candidate reads
    // so a large replacement is not rejected merely because both snapshots
    // are charged through one shared budget.
    budget.output_bytes = 0;
    budget.title_bytes = 0;
    budget.message(source.len(), ROOT_DEPTH)?;
    let strict = preflight_chart_title(source, budget)?;
    let view: projection::ChartTitleArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    let projected = ChartTitleSnapshot {
        title_visible: view.tschchartinfodefaultshowtitle,
        title: view.tschchartinfodefaulttitle,
        raw: source,
    };
    if projected.title_visible != strict.title_visible || projected.title != strict.title {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode only the optional field-23 title text.
pub fn decode_chart_title_text(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<&str>, DecodeError> {
    Ok(decode_chart_title(source, options)?.title())
}

/// Decode the title text only when native field 21 is explicitly true.
pub fn decode_visible_chart_title(
    source: &[u8],
    options: DecodeOptions,
) -> Result<Option<&str>, DecodeError> {
    Ok(decode_chart_title(source, options)?.visible_title())
}

/// Rewrite only fields 21 and 23 of one generated-extension payload.
///
/// Existing unknown field spans and their order are copied byte-for-byte.
/// Existing selected fields (21 and 23) are replaced at their original
/// positions. A requested selected field that is absent from `source` is
/// appended after all source spans in field-number order (21, then 23).
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
pub fn rewrite_chart_title<'source, 'title>(
    source: &'source [u8],
    write: ChartTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    Ok(rewrite_chart_title_with_report(source, write, options)?.0)
}

/// Rewrite the generated extension and return exact source/output accounting.
pub fn rewrite_chart_title_with_report<'source, 'title>(
    source: &'source [u8],
    write: ChartTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    // Strictly validate the complete source before looking at the requested
    // semantic pair. This keeps an otherwise matching write from taking a
    // fast path around duplicate selected fields or malformed unknown spans.
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let current = decode_chart_title_with_budget(source, options, &mut budget)?;
    if let Some(title) = write.title {
        if title.len() > options.max_title_bytes {
            return Err(DecodeError::title_limit(
                title.len(),
                options.max_title_bytes,
            ));
        }
    }

    // A public clear means "remove a visible title". Hidden and absent title
    // state is already clear, even when an old field-23 value remains in the
    // generated payload. Treat that request as an exact source no-op so a
    // stale hidden value is never rewritten or normalized.
    let clear_noop = write.title_visible == Some(false)
        && write.title.is_none()
        && current.title_visible != Some(true);
    let changed = !clear_noop
        && (current.title_visible != write.title_visible || current.title != write.title);
    if !changed {
        if source.len() > options.max_output_bytes {
            return Err(DecodeError::output_limit(
                source.len(),
                options.max_output_bytes,
            ));
        }
        let output = clone_output(source)?;
        return Ok((
            output,
            RewriteReport {
                input_bytes: source.len(),
                output_bytes: source.len(),
                fields: budget.fields,
                work_bytes: budget.work_bytes,
                max_depth: budget.max_depth,
                changed: false,
            },
        ));
    }

    let output_bytes = measure_rewrite_output(source, write, options, &mut budget)?;
    let mut output = reserve_output(output_bytes)?;
    let mut saw_visible = false;
    let mut saw_title = false;
    visit_field_spans(source, &mut budget, |span| {
        let raw = &source[span.start..span.end];
        match span.number {
            CHART_TITLE_VISIBLE_FIELD => {
                saw_visible = true;
                if let Some(value) = write.title_visible {
                    append_varint_field(&mut output, CHART_TITLE_VISIBLE_FIELD, u64::from(value));
                }
            },
            CHART_TITLE_TEXT_FIELD => {
                saw_title = true;
                if let Some(value) = write.title {
                    append_text_field(&mut output, CHART_TITLE_TEXT_FIELD, value);
                }
            },
            _ => output.extend_from_slice(raw),
        }
        Ok(())
    })?;
    // Append every missing selected field after all original spans in field
    // number order. The source does not carry a position for an absent field,
    // so this deterministic rule is part of the public contract.
    if !saw_visible && let Some(value) = write.title_visible {
        append_varint_field(&mut output, CHART_TITLE_VISIBLE_FIELD, u64::from(value));
    }
    if !saw_title && let Some(value) = write.title {
        append_text_field(&mut output, CHART_TITLE_TEXT_FIELD, value);
    }
    debug_assert_eq!(output.len(), output_bytes);

    // The candidate may be larger than the source when a title grows. Let its
    // message ceiling reach the larger of the caller's input ceiling and the
    // candidate length, but keep the independent output ceiling unchanged.
    // `measure_rewrite_output` has already checked that the candidate fits the
    // caller's output cap, and retaining that cap here keeps readback from
    // widening the write budget.
    let readback_options = DecodeOptions {
        max_message_bytes: options.max_message_bytes.max(output.len()),
        ..options
    };
    validate_decode_input(&output, readback_options)?;
    let readback = decode_chart_title_with_budget(&output, readback_options, &mut budget)?;
    if readback.title_visible != write.title_visible || readback.title != write.title {
        return Err(DecodeError::projection());
    }
    Ok((
        output,
        RewriteReport {
            input_bytes: source.len(),
            output_bytes,
            fields: budget.fields,
            work_bytes: budget.work_bytes,
            max_depth: budget.max_depth,
            changed: true,
        },
    ))
}

/// Rewrite using the explicit generated-extension spelling.
pub fn rewrite_chart_title_extension<'source, 'title>(
    source: &'source [u8],
    write: ChartTitleWrite<'title>,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_chart_title(source, write, options)
}

/// Compatibility spelling emphasizing that `source` is the extension payload.
pub fn decode_chart_title_extension<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<ChartTitleSnapshot<'source>, DecodeError> {
    decode_chart_title(source, options)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_maximum = usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_error| DecodeError {
        kind: DecodeErrorKind::Resource(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: usize::MAX,
        }),
    })?;
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
        if bytes > self.options.max_title_bytes {
            return Err(DecodeError::title_limit(
                bytes,
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
        self.title_bytes = self
            .title_bytes
            .checked_add(bytes)
            .ok_or_else(|| DecodeError::title_limit(usize::MAX, self.options.max_title_bytes))?;
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
        }
    }
}

fn preflight_chart_title<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<ChartTitleSnapshot<'source>, DecodeError> {
    let mut visible = None;
    let mut title = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, budget, ROOT_DEPTH)? {
        match field.number {
            CHART_TITLE_VISIBLE_FIELD => {
                if visible.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaultshowtitle",
                    ));
                }
                visible = Some(require_canonical_bool(field.varint()?)?);
            },
            CHART_TITLE_TEXT_FIELD => {
                if title.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaulttitle",
                    ));
                }
                let bytes = field.length_delimited()?;
                budget.title(bytes.len())?;
                title = Some(str::from_utf8(bytes).map_err(|_error| {
                    DecodeError::invalid_utf8(
                        "TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaulttitle",
                    )
                })?);
            },
            _ => {},
        }
    }
    Ok(ChartTitleSnapshot {
        title_visible: visible,
        title,
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

fn measure_rewrite_output(
    source: &[u8],
    write: ChartTitleWrite<'_>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut length = 0usize;
    let mut saw_visible = false;
    let mut saw_title = false;
    visit_field_spans(source, budget, |span| {
        let replacement = match span.number {
            CHART_TITLE_VISIBLE_FIELD => {
                saw_visible = true;
                write.title_visible.map_or(0, |value| {
                    varint_field_len(CHART_TITLE_VISIBLE_FIELD, u64::from(value))
                })
            },
            CHART_TITLE_TEXT_FIELD => {
                saw_title = true;
                write
                    .title
                    .map_or(0, |value| text_field_len(CHART_TITLE_TEXT_FIELD, value))
            },
            _ => span.end - span.start,
        };
        length = length
            .checked_add(replacement)
            .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
        Ok(())
    })?;
    if !saw_visible && let Some(value) = write.title_visible {
        length = length
            .checked_add(varint_field_len(
                CHART_TITLE_VISIBLE_FIELD,
                u64::from(value),
            ))
            .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
    }
    if !saw_title && let Some(value) = write.title {
        length = length
            .checked_add(text_field_len(CHART_TITLE_TEXT_FIELD, value))
            .ok_or_else(|| DecodeError::output_limit(usize::MAX, options.max_output_bytes))?;
    }
    if length > options.max_output_bytes {
        return Err(DecodeError::output_limit(length, options.max_output_bytes));
    }
    Ok(length)
}

fn varint_field_len(number: u32, value: u64) -> usize {
    varint_len(u64::from(number) << 3) + varint_len(value)
}

fn text_field_len(number: u32, value: &str) -> usize {
    varint_len((u64::from(number) << 3) | 2) + varint_len(value.len() as u64) + value.len()
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

fn clone_output(source: &[u8]) -> Result<Vec<u8>, DecodeError> {
    let mut output = reserve_output(source.len())?;
    output.extend_from_slice(source);
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
        ChartTitleSnapshot, ChartTitleWrite, DecodeLimit, DecodeOptions, WireResourceLimit,
        decode_chart_title, decode_chart_title_text, decode_visible_chart_title,
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
        let mut source = field_varint(21, 1);
        source.extend(field_text(23, "Revenue 📈".as_bytes()));
        let unknown = field_varint(4000, 42);
        source.extend_from_slice(&unknown);
        let before = source.clone();
        let snapshot = decode_chart_title(&source, options(&source)).expect("title");
        assert_eq!(snapshot.title_visible(), Some(true));
        assert_eq!(snapshot.show_title(), Some(true));
        assert_eq!(snapshot.title(), Some("Revenue 📈"));
        assert_eq!(snapshot.visible_title(), Some("Revenue 📈"));
        assert_eq!(snapshot.raw(), before.as_slice());
        assert_eq!(
            decode_chart_title_text(&source, options(&source)),
            Ok(Some("Revenue 📈"))
        );
        assert_eq!(
            decode_visible_chart_title(&source, options(&source)),
            Ok(Some("Revenue 📈"))
        );
        assert_eq!(source, before);
    }

    #[test]
    fn absent_fields_are_exact_no_op_and_presence_is_retained() {
        let source = Vec::new();
        let snapshot = decode_chart_title(&source, options(&source)).expect("empty");
        assert_eq!(
            snapshot,
            ChartTitleSnapshot {
                title_visible: None,
                title: None,
                raw: &[]
            }
        );
        assert_eq!(snapshot.visible_title(), None);

        let hidden = field_varint(21, 0);
        let snapshot = decode_chart_title(&hidden, options(&hidden)).expect("hidden");
        assert_eq!(snapshot.title_visible(), Some(false));
        assert_eq!(snapshot.title(), None);
        assert_eq!(snapshot.visible_title(), None);

        let visible_without_text = field_varint(21, 1);
        let snapshot = decode_chart_title(&visible_without_text, options(&visible_without_text))
            .expect("default empty title");
        assert_eq!(snapshot.title(), None);
        assert_eq!(snapshot.visible_title(), Some(""));
    }

    #[test]
    fn duplicate_wrong_wire_and_truncated_selected_fields_are_rejected() {
        let mut duplicate_visible = field_varint(21, 1);
        duplicate_visible.extend(field_varint(21, 0));
        assert_eq!(
            decode_chart_title(&duplicate_visible, options(&duplicate_visible))
                .expect_err("duplicate visible")
                .duplicate_singular_field(),
            Some("TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaultshowtitle")
        );

        let mut duplicate_title = field_text(23, b"a");
        duplicate_title.extend(field_text(23, b"b"));
        assert_eq!(
            decode_chart_title(&duplicate_title, options(&duplicate_title))
                .expect_err("duplicate title")
                .duplicate_singular_field(),
            Some("TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaulttitle")
        );

        for source in [
            vec![0xaa, 0x01, 0x01],       // field 21 with length-delimited wire type
            vec![0xaa, 0x01, 0x02, b'a'], // field 23 with truncated payload
            vec![0xa8, 0x01],             // field 21 with truncated varint
        ] {
            assert!(
                decode_chart_title(&source, options(&source)).is_err(),
                "{source:?}"
            );
        }
    }

    #[test]
    fn bool_utf8_and_canonical_wire_rules_are_enforced() {
        let bad_bool = field_varint(21, 2);
        assert_eq!(
            decode_chart_title(&bad_bool, options(&bad_bool))
                .expect_err("noncanonical bool")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );

        let bad_utf8 = field_text(23, &[0xff]);
        assert_eq!(
            decode_chart_title(&bad_utf8, options(&bad_utf8))
                .expect_err("invalid UTF-8")
                .invalid_utf8_field(),
            Some("TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaulttitle")
        );

        let overlong_key = vec![0xa8, 0x81, 0x00, 0x01];
        assert_eq!(
            decode_chart_title(&overlong_key, options(&overlong_key))
                .expect_err("overlong key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let overlong_length = vec![0xba, 0x01, 0x80, 0x00];
        assert_eq!(
            decode_chart_title(&overlong_length, options(&overlong_length))
                .expect_err("overlong length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );
    }

    #[test]
    fn finite_field_work_output_title_and_nesting_limits_are_observable() {
        let source = [field_varint(21, 1), field_text(23, b"abcd")].concat();
        assert_eq!(
            decode_chart_title(
                &source,
                DecodeOptions::new(source.len(), 1, source.len() * 8, 8),
            )
            .expect_err("field cap")
            .field_limit_values(),
            Some((2, 1))
        );
        assert_eq!(
            decode_chart_title(
                &source,
                DecodeOptions::new(source.len(), 8, source.len(), 8),
            )
            .expect_err("work cap")
            .work_limit_values(),
            Some((source.len() * 2, source.len()))
        );
        assert_eq!(
            decode_chart_title(
                &source,
                DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_output_bytes(3),
            )
            .expect_err("output cap")
            .output_limit_values(),
            Some((4, 3))
        );
        assert_eq!(
            decode_chart_title(
                &source,
                DecodeOptions::new(source.len(), 8, source.len() * 8, 8).with_max_title_bytes(3),
            )
            .expect_err("title cap")
            .title_limit_values(),
            Some((4, 3))
        );
        let nesting = vec![0x0b, 0x13, 0x14, 0x0c];
        let error = decode_chart_title(
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
        let source = field_varint(21, 1);
        assert_eq!(
            decode_chart_title(
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
        let error = decode_chart_title(
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
        let mut source = field_varint(21, 1);
        source.extend([0x82, 0x01, 0x01, 0xff]);
        let before = source.clone();
        let snapshot = decode_chart_title(&source, options(&source)).expect("unknown field");
        assert_eq!(snapshot.title_visible(), Some(true));
        assert_eq!(snapshot.raw(), before.as_slice());
        assert_eq!(source, before);

        let unknown_noncanonical = [0x82, 0x81, 0x00, 0x01, 0xff];
        let error = decode_chart_title(&unknown_noncanonical, options(&unknown_noncanonical))
            .expect_err("noncanonical unknown key");
        assert_eq!(error.noncanonical_reason(), Some("protobuf field key"));
    }

    #[test]
    fn rewrite_is_wire_local_exact_noop_and_preserves_unknown_spans() {
        let mut source = field_varint(4000, 7);
        source.extend(field_varint(21, 1));
        source.extend(field_text(23, b"old"));
        source.extend(field_varint(4001, 8));
        let original_unknowns = [source[0..4].to_vec(), source[source.len() - 4..].to_vec()];
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let (noop, noop_report) = super::rewrite_chart_title_with_report(
            &source,
            ChartTitleWrite::new(Some(true), Some("old")),
            options,
        )
        .expect("exact no-op");
        assert_eq!(noop, source);
        assert!(!noop_report.changed());

        let (rewritten, report) = super::rewrite_chart_title_with_report(
            &source,
            ChartTitleWrite::new(Some(false), None),
            options,
        )
        .expect("wire-local rewrite");
        assert!(report.changed());
        assert_eq!(
            decode_chart_title(&rewritten, options)
                .expect("readback")
                .title_visible(),
            Some(false)
        );
        assert_eq!(
            decode_chart_title(&rewritten, options)
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
        let rewritten = super::rewrite_chart_title(
            &source,
            ChartTitleWrite::new(Some(true), Some("new")),
            options,
        )
        .expect("append selected fields");
        assert_eq!(
            decode_chart_title(&rewritten, options)
                .expect("readback")
                .visible_title(),
            Some("new")
        );
        assert!(rewritten.starts_with(&source));

        let capped = options.with_max_output_bytes(source.len());
        let expected_output =
            source.len() + field_varint(21, 1).len() + field_text(23, b"new").len();
        assert_eq!(
            super::rewrite_chart_title(
                &source,
                ChartTitleWrite::new(Some(true), Some("new")),
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
            field_text(23, b"old"),
            field_varint(4001, 8),
            field_varint(21, 0),
            field_varint(4002, 9),
        ]
        .concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let rewritten = super::rewrite_chart_title(
            &source,
            ChartTitleWrite::new(Some(true), Some("new")),
            options,
        )
        .expect("replace selected fields in place");
        let expected = [
            field_varint(4000, 7),
            field_text(23, b"new"),
            field_varint(4001, 8),
            field_varint(21, 1),
            field_varint(4002, 9),
        ]
        .concat();
        assert_eq!(rewritten, expected);
    }

    #[test]
    fn rewrite_does_not_guess_removed_field_positions() {
        // These two distinct source layouts intentionally collapse to the same
        // clear result. No source-local rewrite can know whether field 23 used
        // to precede or follow field 21 after it has been removed.
        let before_visible = [
            field_text(23, b"old"),
            field_varint(4000, 7),
            field_varint(21, 1),
        ]
        .concat();
        let after_visible = [
            field_varint(4000, 7),
            field_varint(21, 1),
            field_text(23, b"old"),
        ]
        .concat();
        let options = DecodeOptions::new(1024, 128, 8192, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let cleared_before = super::rewrite_chart_title(
            &before_visible,
            ChartTitleWrite::new(Some(false), None),
            options,
        )
        .expect("clear title from before-visible layout");
        let cleared_after = super::rewrite_chart_title(
            &after_visible,
            ChartTitleWrite::new(Some(false), None),
            options,
        )
        .expect("clear title from after-visible layout");
        assert_eq!(cleared_before, cleared_after);
        assert_eq!(
            cleared_before,
            [field_varint(4000, 7), field_varint(21, 0)].concat()
        );

        let restored = super::rewrite_chart_title(
            &cleared_before,
            ChartTitleWrite::new(Some(true), Some("old")),
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
            field_varint(21, 0),
            [field_varint(21, 0), field_text(23, b"stale")].concat(),
            field_text(23, b"stale"),
            Vec::new(),
        ] {
            let (rewritten, report) = super::rewrite_chart_title_with_report(
                &source,
                ChartTitleWrite::new(Some(false), None),
                options,
            )
            .expect("hidden/absent clear");
            assert_eq!(rewritten, source);
            assert!(!report.changed());
        }

        // A visible empty title is semantically present even when field 23 is
        // absent. Clearing and restoring that pair must retain the absence.
        let source = field_varint(21, 1);
        let cleared =
            super::rewrite_chart_title(&source, ChartTitleWrite::new(Some(false), None), options)
                .expect("clear visible title");
        let restored =
            super::rewrite_chart_title(&cleared, ChartTitleWrite::new(Some(true), None), options)
                .expect("restore visible empty title");
        assert_eq!(restored, source);
        assert_eq!(
            decode_chart_title(&restored, options)
                .expect("restored title")
                .visible_title(),
            Some("")
        );
        assert_eq!(
            decode_chart_title(&restored, options)
                .expect("restored title")
                .title(),
            None
        );
    }

    #[test]
    fn rewrite_limits_are_inclusive_and_title_limit_precedes_output_allocation() {
        let source = field_varint(4000, 7);
        let write = ChartTitleWrite::new(Some(true), Some("new"));
        let base = DecodeOptions::new(1024, 128, 8192, 8).with_max_title_bytes(128);
        let expected_output =
            source.len() + field_varint(21, 1).len() + field_text(23, b"new").len();

        let (_, report) = super::rewrite_chart_title_with_report(
            &source,
            write,
            base.with_max_output_bytes(expected_output),
        )
        .expect("inclusive output limit");
        assert_eq!(report.output_bytes(), expected_output);

        let output_error = super::rewrite_chart_title(
            &source,
            write,
            base.with_max_output_bytes(expected_output - 1),
        )
        .expect_err("max-minus-one output limit");
        assert_eq!(
            output_error.output_limit_values(),
            Some((expected_output, expected_output - 1))
        );

        let title_error = super::rewrite_chart_title(&source, write, base.with_max_title_bytes(2))
            .expect_err("title limit");
        assert_eq!(title_error.title_limit_values(), Some((3, 2)));
    }

    #[test]
    fn rewrite_readback_accepts_larger_title_when_output_cap_allows_it() {
        let source = field_varint(21, 1);
        let title = "a title larger than the source";
        let write = ChartTitleWrite::new(Some(true), Some(title));
        let expected_output = [source.clone(), field_text(23, title.as_bytes())].concat();
        let options = DecodeOptions::new(source.len(), 128, 8192, 8)
            .with_max_output_bytes(expected_output.len())
            .with_max_title_bytes(title.len());

        let (rewritten, report) = super::rewrite_chart_title_with_report(&source, write, options)
            .expect("candidate larger than source fits output cap");
        assert_eq!(rewritten, expected_output);
        assert_eq!(report.output_bytes(), expected_output.len());
        let readback_options = DecodeOptions::new(expected_output.len(), 128, 8192, 8)
            .with_max_output_bytes(expected_output.len())
            .with_max_title_bytes(title.len());
        assert_eq!(
            decode_chart_title(&rewritten, readback_options)
                .expect("larger candidate readback")
                .title(),
            Some(title)
        );

        let capped = options.with_max_output_bytes(expected_output.len() - 1);
        let error = super::rewrite_chart_title(&source, write, capped)
            .expect_err("independent output cap must remain enforced");
        assert_eq!(
            error.output_limit_values(),
            Some((expected_output.len(), expected_output.len() - 1))
        );
    }

    #[test]
    fn rewrite_work_budget_covers_sizing_emission_and_readback() {
        let source = [field_varint(21, 1), field_text(23, b"old")].concat();
        let write = ChartTitleWrite::new(Some(true), Some("new title"));
        let unconstrained = DecodeOptions::new(1024, usize::MAX, usize::MAX, 8)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);
        let (output, report) =
            super::rewrite_chart_title_with_report(&source, write, unconstrained)
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
        super::rewrite_chart_title_with_report(&source, write, exact)
            .expect("the exact aggregate work ceiling is inclusive");

        let below = DecodeOptions::new(1024, usize::MAX, expected_work - 1, 8)
            .with_max_output_bytes(output.len())
            .with_max_title_bytes(9);
        let error = super::rewrite_chart_title_with_report(&source, write, below)
            .expect_err("one byte below aggregate work must fail");
        assert_eq!(
            error.work_limit_values(),
            Some((expected_work, expected_work - 1))
        );
        assert_eq!(
            source,
            [field_varint(21, 1), field_text(23, b"old")].concat()
        );
    }

    #[test]
    fn rewrite_matching_semantics_still_preflight_duplicates_and_unknown_wire() {
        let mut duplicate = field_varint(21, 1);
        duplicate.extend(field_varint(21, 1));
        let duplicate_error = super::rewrite_chart_title(
            &duplicate,
            ChartTitleWrite::new(Some(true), None),
            options(&duplicate),
        )
        .expect_err("duplicate selected field");
        assert_eq!(
            duplicate_error.duplicate_singular_field(),
            Some("TSCH.Generated.ChartNonStyleArchive.tschchartinfodefaultshowtitle")
        );

        let unknown_noncanonical =
            [field_varint(21, 1), vec![0x82, 0x81, 0x00, 0x01, 0xff]].concat();
        let unknown_error = super::rewrite_chart_title(
            &unknown_noncanonical,
            ChartTitleWrite::new(Some(true), None),
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
        source.extend(field_varint(21, 1));
        // A rewrite charges the source preflight, sizing scan, emission scan,
        // and candidate readback against one aggregate field budget. The
        // group contributes three visits plus the selected field per pass.
        let options = DecodeOptions::new(source.len(), 16, source.len() * 8, 2)
            .with_max_output_bytes(1024)
            .with_max_title_bytes(128);

        let (_, report) =
            super::decode_chart_title_with_report(&source, options).expect("matched unknown group");
        assert_eq!(report.fields(), 4);
        assert_eq!(report.max_depth(), 2);

        let rewritten =
            super::rewrite_chart_title(&source, ChartTitleWrite::new(Some(false), None), options)
                .expect("preserve unknown group");
        assert_eq!(&rewritten[..source.len() - 2], &source[..source.len() - 2]);

        let error = decode_chart_title(
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
        let source = field_text(23, b"x");
        let error = decode_chart_title(
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
}
