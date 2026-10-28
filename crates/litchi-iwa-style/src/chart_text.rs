//! Shared bounded chart title and axis-title payload engine.
//!
//! Format owners resolve a semantic chart and retain archive-header ownership;
//! this module owns the generated-extension wire boundary. It validates the
//! outer field-10,000 envelope, delegates scalar projection to the strict
//! generated codecs, preserves unknown bytes during rewrites, and reports the
//! complete bounded payload cost to the owner.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The payload engine keeps limits, reports, and wire phases together."
)]

use std::fmt;

use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::wire::{WireView, patch_length_delimited_field};
use litchi_iwa_common::{LimitKind as WireLimitKind, WireLimits};
use litchi_iwa_protos::{
    keynote_chart_axis_title_codec as axis_codec, keynote_chart_title_codec as title_codec,
};
use thiserror::Error as ThisError;

const GENERATED_EXTENSION_FIELD: u32 = 10_000;

/// The selected generated chart-text field family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartTextKind {
    /// The visible chart title in `ChartNonStyleArchive`.
    Title,
    /// One visible axis title in `ChartAxisNonStyleArchive`.
    Axis(Axis),
}

fn codec_axis(axis: Axis) -> axis_codec::AxisTitleKind {
    match axis {
        Axis::Category => axis_codec::AxisTitleKind::Category,
        Axis::Value => axis_codec::AxisTitleKind::Value,
    }
}

/// A finite payload budget supplied by a format owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_input_bytes: usize,
    max_fields: usize,
    max_output_bytes: usize,
    max_nesting: u32,
    max_work_bytes: usize,
    max_text_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl Limits {
    /// Build an explicit finite payload policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_output_bytes: usize,
        max_nesting: u32,
        max_work_bytes: usize,
        max_text_bytes: usize,
        max_allocations: usize,
        max_retained_bytes: usize,
        max_scratch_bytes: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_fields,
            max_output_bytes,
            max_nesting,
            max_work_bytes,
            max_text_bytes,
            max_allocations,
            max_retained_bytes,
            max_scratch_bytes,
        }
    }

    /// Build a conservative finite profile from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.saturating_mul(8).max(1),
            bytes.saturating_add(64).max(1),
            8,
            bytes.saturating_mul(16).max(1),
            bytes,
            8,
            bytes.saturating_mul(2).max(1),
            bytes,
        )
    }

    /// Maximum complete payload bytes inspected.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Maximum aggregate wire fields.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Maximum rewritten outer payload bytes.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Maximum protobuf nesting depth.
    #[must_use]
    pub const fn max_nesting(self) -> u32 {
        self.max_nesting
    }

    /// Maximum aggregate wire work.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Maximum selected UTF-8 title bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Maximum logical allocations in a prepared axis rewrite.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    /// Maximum retained source and candidate bytes.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    /// Maximum candidate scratch bytes.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }
}

/// Resource counted by one shared chart-text payload operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete outer payload bytes inspected.
    InputBytes,
    /// Strict wire fields visited.
    Fields,
    /// Rewritten outer payload bytes.
    OutputBytes,
    /// Nested protobuf depth.
    Nesting,
    /// Aggregate strict and generated-codec work.
    Work,
    /// Selected UTF-8 title bytes.
    TextBytes,
    /// Prepared rewrite allocations.
    Allocations,
    /// Source plus candidate bytes retained by a rewrite.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Fields => "wire fields",
            Self::OutputBytes => "output bytes",
            Self::Nesting => "wire nesting",
            Self::Work => "wire work",
            Self::TextBytes => "text bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
        })
    }
}

/// A typed failure from the shared chart-text payload engine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The payload or generated extension is malformed or inconsistent.
    #[error("the chart-text payload is invalid")]
    InvalidSource,
    /// A finite payload resource ceiling was exceeded.
    #[error("chart text {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded output or span allocation failed.
    #[error("could not allocate {amount} chart-text units")]
    Allocation {
        /// Requested allocation units.
        amount: usize,
    },
}

/// Borrowed semantic text selected from one generated chart-text payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Snapshot<'source> {
    text: Option<&'source str>,
}

impl<'source> Snapshot<'source> {
    /// The selected title when its native visibility field is explicitly true.
    #[must_use]
    pub const fn text(self) -> Option<&'source str> {
        self.text
    }
}

/// Exact bounded accounting for one decode or rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Report {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
    changed: bool,
}

impl Report {
    /// Complete source payload bytes inspected.
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Candidate payload bytes emitted, or zero for a read.
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    /// Aggregate strict wire fields visited.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict and generated-codec work bytes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested protobuf depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Selected UTF-8 bytes observed.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Logical allocations reported by the axis codec.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Bytes retained by the operation.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Scratch bytes reserved by the operation.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }

    /// Whether the selected payload bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
}

/// A bounded rewritten outer chart payload and its report.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite {
    output: Vec<u8>,
    report: Report,
}

impl Rewrite {
    /// Borrow the rewritten payload.
    #[must_use]
    pub fn output(&self) -> &[u8] {
        &self.output
    }

    /// Consume the result and return its payload.
    #[must_use]
    pub fn into_output(self) -> Vec<u8> {
        self.output
    }

    /// Exact bounded rewrite accounting.
    #[must_use]
    pub const fn report(&self) -> Report {
        self.report
    }
}

#[derive(Debug, Clone, Copy)]
struct LocatedExtension<'source> {
    payload: Option<&'source [u8]>,
    raw: Option<&'source [u8]>,
    fields: usize,
}

/// Decode one selected title or axis title from the complete chart payload.
pub fn decode(
    payload: &[u8],
    kind: ChartTextKind,
    limits: Limits,
) -> Result<(Snapshot<'_>, Report), Error> {
    let located = locate_extension(payload, limits)?;
    let base = base_report(payload, located.fields);
    let Some(extension) = located.payload else {
        return Ok((Snapshot { text: None }, base));
    };
    match kind {
        ChartTextKind::Title => {
            let options = title_options(extension, located.fields, limits)?;
            let (snapshot, report) =
                title_codec::decode_chart_title_with_report(extension, options)
                    .map_err(map_title_error)?;
            let text = snapshot.visible_title();
            Ok((
                Snapshot { text },
                merge_decode_report(base, report, text.map(str::len))?,
            ))
        },
        ChartTextKind::Axis(axis) => {
            let options = axis_options(extension, located.fields, limits)?;
            let (snapshot, report) = axis_codec::decode_axis_title_with_report(extension, options)
                .map_err(map_axis_error)?;
            let text = snapshot.visible_title(codec_axis(axis));
            Ok((
                Snapshot { text },
                merge_decode_report(base, report, text.map(str::len))?,
            ))
        },
    }
}

/// Rewrite one selected title or axis title while preserving every unrelated
/// outer and generated-extension byte.
pub fn rewrite(
    payload: &[u8],
    kind: ChartTextKind,
    after: Option<&str>,
    limits: Limits,
) -> Result<Rewrite, Error> {
    let located = locate_extension(payload, limits)?;
    let base = base_report(payload, located.fields);
    let extension = located.payload.unwrap_or_default();
    let (rewritten, codec_report) = match kind {
        ChartTextKind::Title => {
            let options = title_options(extension, located.fields, limits)?;
            let write = title_codec::ChartTitleWrite::new(Some(after.is_some()), after);
            let output = title_codec::rewrite_chart_title_with_report(extension, write, options)
                .map_err(map_title_error)?;
            (output.0, CodecReport::Title(output.1))
        },
        ChartTextKind::Axis(axis) => {
            let options = axis_options(extension, located.fields, limits)?;
            let write = axis_codec::AxisTitleWrite::preserve().with_title(codec_axis(axis), after);
            let output = axis_codec::rewrite_axis_title_with_report(extension, write, options)
                .map_err(map_axis_error)?;
            (output.0, CodecReport::Axis(output.1))
        },
    };
    let changed = codec_report.changed();
    let output = if !changed {
        clone_output(payload, limits.max_output_bytes())?
    } else {
        let projected = projected_output_bytes(payload, located.raw, rewritten.len())?;
        check_limit(LimitKind::OutputBytes, projected, limits.max_output_bytes())?;
        patch_length_delimited_field(
            payload,
            GENERATED_EXTENSION_FIELD,
            located.raw.is_some(),
            Some(rewritten.as_slice()),
        )
        .map_err(map_wire_error)?
    };
    let report = merge_rewrite_report(
        base,
        codec_report,
        output.len(),
        after.map(str::len).unwrap_or(0),
        changed && output.as_slice() != payload,
    )?;
    Ok(Rewrite { output, report })
}

#[derive(Debug, Clone, Copy)]
enum CodecReport {
    Title(title_codec::RewriteReport),
    Axis(axis_codec::RewriteReport),
}

impl CodecReport {
    fn changed(self) -> bool {
        match self {
            Self::Title(report) => report.changed(),
            Self::Axis(report) => report.changed(),
        }
    }

    fn fields(self) -> usize {
        match self {
            Self::Title(report) => report.fields(),
            Self::Axis(report) => report.fields(),
        }
    }

    fn work_bytes(self) -> usize {
        match self {
            Self::Title(report) => report.work_bytes(),
            Self::Axis(report) => report.work_bytes(),
        }
    }

    fn max_depth(self) -> u32 {
        match self {
            Self::Title(report) => report.max_depth(),
            Self::Axis(report) => report.max_depth(),
        }
    }

    fn allocations(self) -> usize {
        match self {
            Self::Title(_) => 0,
            Self::Axis(report) => report.allocations(),
        }
    }

    fn retained_bytes(self) -> usize {
        match self {
            Self::Title(_) => 0,
            Self::Axis(report) => report.retained_bytes(),
        }
    }

    fn scratch_bytes(self) -> usize {
        match self {
            Self::Title(_) => 0,
            Self::Axis(report) => report.scratch_bytes(),
        }
    }
}

fn base_report(payload: &[u8], fields: usize) -> Report {
    Report {
        input_bytes: payload.len(),
        output_bytes: 0,
        fields,
        work_bytes: payload.len(),
        max_depth: 1,
        text_bytes: 0,
        allocations: 0,
        retained_bytes: payload.len(),
        scratch_bytes: 0,
        changed: false,
    }
}

fn merge_decode_report<R: DecodeReportView>(
    base: Report,
    report: R,
    text_bytes: Option<usize>,
) -> Result<Report, Error> {
    Ok(Report {
        input_bytes: base.input_bytes,
        output_bytes: 0,
        fields: add(base.fields, report.fields())?,
        work_bytes: add(base.work_bytes, report.work_bytes())?,
        max_depth: base.max_depth.max(report.max_depth()),
        text_bytes: text_bytes.unwrap_or(0),
        allocations: report.allocations(),
        retained_bytes: add(base.retained_bytes, report.retained_bytes())?,
        scratch_bytes: report.scratch_bytes(),
        changed: false,
    })
}

fn merge_rewrite_report(
    base: Report,
    report: CodecReport,
    output_bytes: usize,
    text_bytes: usize,
    changed: bool,
) -> Result<Report, Error> {
    Ok(Report {
        input_bytes: base.input_bytes,
        output_bytes,
        fields: add(base.fields, report.fields())?,
        work_bytes: add(base.work_bytes, report.work_bytes())?,
        max_depth: base.max_depth.max(report.max_depth()),
        text_bytes,
        allocations: report.allocations(),
        retained_bytes: add(base.retained_bytes, report.retained_bytes())?,
        scratch_bytes: report.scratch_bytes(),
        changed,
    })
}

trait DecodeReportView {
    fn fields(self) -> usize;
    fn work_bytes(self) -> usize;
    fn max_depth(self) -> u32;
    fn allocations(self) -> usize;
    fn retained_bytes(self) -> usize;
    fn scratch_bytes(self) -> usize;
}

impl DecodeReportView for title_codec::DecodeReport {
    fn fields(self) -> usize {
        self.fields()
    }

    fn work_bytes(self) -> usize {
        self.work_bytes()
    }

    fn max_depth(self) -> u32 {
        self.max_depth()
    }

    fn allocations(self) -> usize {
        0
    }

    fn retained_bytes(self) -> usize {
        0
    }

    fn scratch_bytes(self) -> usize {
        0
    }
}

impl DecodeReportView for axis_codec::DecodeReport {
    fn fields(self) -> usize {
        self.fields()
    }

    fn work_bytes(self) -> usize {
        self.work_bytes()
    }

    fn max_depth(self) -> u32 {
        self.max_depth()
    }

    fn allocations(self) -> usize {
        self.allocations()
    }

    fn retained_bytes(self) -> usize {
        self.retained_bytes()
    }

    fn scratch_bytes(self) -> usize {
        self.scratch_bytes()
    }
}

fn locate_extension<'source>(
    payload: &'source [u8],
    limits: Limits,
) -> Result<LocatedExtension<'source>, Error> {
    let wire_limits = wire_limits(limits)?;
    let view = WireView::parse(payload, wire_limits).map_err(map_wire_error)?;
    let mut selected = None;
    for field in view.fields() {
        if field.number() != GENERATED_EXTENSION_FIELD {
            continue;
        }
        if selected.is_some() || field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        selected = Some((field.payload(), field.raw()));
    }
    Ok(LocatedExtension {
        payload: selected.map(|(payload, _)| payload),
        raw: selected.map(|(_, raw)| raw),
        fields: view.len(),
    })
}

fn wire_limits(limits: Limits) -> Result<WireLimits, Error> {
    WireLimits::DEFAULT
        .with_input_bytes(limits.max_input_bytes())
        .and_then(|value| value.with_fields(limits.max_fields()))
        .and_then(|value| value.with_output_bytes(limits.max_output_bytes()))
        .and_then(|value| {
            value.with_nesting(usize::try_from(limits.max_nesting()).unwrap_or(usize::MAX))
        })
        .and_then(|value| value.with_rewrite_work(limits.max_work_bytes()))
        .map_err(map_wire_error)
}

fn title_options(
    extension: &[u8],
    outer_fields: usize,
    limits: Limits,
) -> Result<title_codec::DecodeOptions, Error> {
    let fields = residual(limits.max_fields(), outer_fields, LimitKind::Fields)?;
    let work = residual(limits.max_work_bytes(), extension.len(), LimitKind::Work)?;
    Ok(title_codec::DecodeOptions::new(
        extension.len().max(1),
        fields.max(1),
        work.max(1),
        limits.max_nesting(),
    )
    .with_max_output_bytes(limits.max_output_bytes())
    .with_max_title_bytes(limits.max_text_bytes()))
}

fn axis_options(
    extension: &[u8],
    outer_fields: usize,
    limits: Limits,
) -> Result<axis_codec::DecodeOptions, Error> {
    let fields = residual(limits.max_fields(), outer_fields, LimitKind::Fields)?;
    let work = residual(limits.max_work_bytes(), extension.len(), LimitKind::Work)?;
    Ok(axis_codec::DecodeOptions::new(
        extension.len().max(1),
        fields.max(1),
        work.max(1),
        limits.max_nesting(),
    )
    .with_max_output_bytes(limits.max_output_bytes())
    .with_max_title_bytes(limits.max_text_bytes())
    .with_max_allocations(limits.max_allocations())
    .with_max_retained_bytes(limits.max_retained_bytes())
    .with_max_scratch_bytes(limits.max_scratch_bytes()))
}

fn projected_output_bytes(
    payload: &[u8],
    previous: Option<&[u8]>,
    replacement: usize,
) -> Result<usize, Error> {
    let key =
        litchi_iwa_common::varint::encoded_len((u64::from(GENERATED_EXTENSION_FIELD) << 3) | 2);
    let length = litchi_iwa_common::varint::encoded_len(
        u64::try_from(replacement).map_err(|_| Error::InvalidSource)?,
    );
    let field = key
        .checked_add(length)
        .and_then(|value| value.checked_add(replacement))
        .ok_or(Error::InvalidSource)?;
    payload
        .len()
        .checked_sub(previous.map_or(0, |value| value.len()))
        .and_then(|value| value.checked_add(field))
        .ok_or(Error::InvalidSource)
}

fn clone_output(payload: &[u8], maximum: usize) -> Result<Vec<u8>, Error> {
    check_limit(LimitKind::OutputBytes, payload.len(), maximum)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(payload.len())
        .map_err(|_| Error::Allocation {
            amount: payload.len(),
        })?;
    output.extend_from_slice(payload);
    Ok(output)
}

fn residual(maximum: usize, used: usize, kind: LimitKind) -> Result<usize, Error> {
    maximum.checked_sub(used).ok_or(Error::Limit {
        kind,
        observed: count(used),
        maximum: count(maximum),
    })
}

fn add(left: usize, right: usize) -> Result<usize, Error> {
    left.checked_add(right).ok_or(Error::InvalidSource)
}

fn check_limit(kind: LimitKind, observed: usize, maximum: usize) -> Result<(), Error> {
    if observed > maximum {
        return Err(Error::Limit {
            kind,
            observed: count(observed),
            maximum: count(maximum),
        });
    }
    Ok(())
}

fn count(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: map_wire_limit(kind),
            observed: count(observed),
            maximum: count(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

const fn map_wire_limit(kind: WireLimitKind) -> LimitKind {
    match kind {
        WireLimitKind::InputBytes => LimitKind::InputBytes,
        WireLimitKind::Fields => LimitKind::Fields,
        WireLimitKind::OutputBytes => LimitKind::OutputBytes,
        WireLimitKind::Nesting => LimitKind::Nesting,
        WireLimitKind::RewriteWork => LimitKind::Work,
        WireLimitKind::TableRows
        | WireLimitKind::TableColumns
        | WireLimitKind::TableCells
        | WireLimitKind::MaterializedCells => LimitKind::Work,
    }
}

fn map_title_error(error: title_codec::DecodeError) -> Error {
    if let Some(limit) = error.resource_limit() {
        return map_title_limit(limit);
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return limit(LimitKind::Fields, observed, maximum);
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return limit(LimitKind::Work, observed, maximum);
    }
    if let Some((observed, maximum)) = error.output_limit_values() {
        return limit(LimitKind::OutputBytes, observed, maximum);
    }
    if let Some((observed, maximum)) = error.title_limit_values() {
        return limit(LimitKind::TextBytes, observed, maximum);
    }
    error
        .allocation_amount()
        .map_or(Error::InvalidSource, |amount| Error::Allocation { amount })
}

fn map_axis_error(error: axis_codec::DecodeError) -> Error {
    if let Some(limit) = error.resource_limit() {
        return map_axis_limit(limit);
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return limit(LimitKind::Fields, observed, maximum);
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return limit(LimitKind::Work, observed, maximum);
    }
    if let Some((observed, maximum)) = error.output_limit_values() {
        return limit(LimitKind::OutputBytes, observed, maximum);
    }
    if let Some((observed, maximum)) = error.title_limit_values() {
        return limit(LimitKind::TextBytes, observed, maximum);
    }
    error
        .allocation_amount()
        .map_or(Error::InvalidSource, |amount| Error::Allocation { amount })
}

fn map_title_limit(limit: title_codec::DecodeLimit) -> Error {
    match limit {
        title_codec::DecodeLimit::Bytes { observed, maximum } => {
            limit(LimitKind::InputBytes, observed, maximum)
        },
        title_codec::DecodeLimit::Fields { observed, maximum } => {
            limit(LimitKind::Fields, observed, maximum)
        },
        title_codec::DecodeLimit::Work { observed, maximum } => {
            limit(LimitKind::Work, observed, maximum)
        },
        title_codec::DecodeLimit::Output { observed, maximum } => {
            limit(LimitKind::OutputBytes, observed, maximum)
        },
        title_codec::DecodeLimit::Title { observed, maximum } => {
            limit(LimitKind::TextBytes, observed, maximum)
        },
        title_codec::DecodeLimit::Nesting { observed, maximum } => Error::Limit {
            kind: LimitKind::Nesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        _ => Error::InvalidSource,
    }
}

fn map_axis_limit(limit: axis_codec::DecodeLimit) -> Error {
    match limit {
        axis_codec::DecodeLimit::Bytes { observed, maximum } => {
            limit(LimitKind::InputBytes, observed, maximum)
        },
        axis_codec::DecodeLimit::Fields { observed, maximum } => {
            limit(LimitKind::Fields, observed, maximum)
        },
        axis_codec::DecodeLimit::Work { observed, maximum } => {
            limit(LimitKind::Work, observed, maximum)
        },
        axis_codec::DecodeLimit::Output { observed, maximum } => {
            limit(LimitKind::OutputBytes, observed, maximum)
        },
        axis_codec::DecodeLimit::Title { observed, maximum } => {
            limit(LimitKind::TextBytes, observed, maximum)
        },
        axis_codec::DecodeLimit::Nesting { observed, maximum } => Error::Limit {
            kind: LimitKind::Nesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        axis_codec::DecodeLimit::Allocations { observed, maximum } => {
            limit(LimitKind::Allocations, observed, maximum)
        },
        axis_codec::DecodeLimit::Retained { observed, maximum } => {
            limit(LimitKind::RetainedBytes, observed, maximum)
        },
        axis_codec::DecodeLimit::Scratch { observed, maximum } => {
            limit(LimitKind::ScratchBytes, observed, maximum)
        },
        _ => Error::InvalidSource,
    }
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind,
        observed: count(observed),
        maximum: count(maximum),
    }
}

#[cfg(test)]
mod tests {
    use super::{ChartTextKind, Error, LimitKind, Limits, decode, rewrite};
    use litchi_iwa_common::chart::axis::Axis;

    fn varint(mut value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        loop {
            let mut byte = u8::try_from(value & 0x7f).expect("masked varint byte");
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

    fn varint_field(number: u32, value: u64) -> Vec<u8> {
        let mut output = varint((u64::from(number) << 3) | 0);
        output.extend(varint(value));
        output
    }

    fn text_field(number: u32, value: &[u8]) -> Vec<u8> {
        let mut output = varint((u64::from(number) << 3) | 2);
        output.extend(varint(u64::try_from(value.len()).expect("test length")));
        output.extend(value);
        output
    }

    fn outer(extension: &[u8]) -> Vec<u8> {
        text_field(GENERATED_EXTENSION_FIELD, extension)
    }

    #[test]
    fn title_rewrite_preserves_unknown_outer_and_extension_bytes() {
        let mut source = text_field(1, b"outer-before");
        let mut extension = text_field(2, b"extension-before");
        extension.extend(varint_field(21, 1));
        extension.extend(text_field(23, b"Before"));
        extension.extend(text_field(77, b"unknown-extension"));
        source.extend(outer(&extension));
        source.extend(varint_field(7, 3));

        let limits = Limits::for_source(&source);
        let (snapshot, report) = decode(&source, ChartTextKind::Title, limits).expect("decode");
        assert_eq!(snapshot.text(), Some("Before"));
        assert!(!report.changed());

        let output = rewrite(
            &source,
            ChartTextKind::Title,
            Some("After"),
            limits_for(&source, 5),
        )
        .expect("rewrite")
        .into_output();
        assert_eq!(
            &output[..source_field_end(&source, 1)],
            &source[..source_field_end(&source, 1)]
        );
        assert_eq!(&output[output.len() - 2..], &source[source.len() - 2..]);
        let (after, _) =
            decode(&output, ChartTextKind::Title, Limits::for_source(&output)).expect("readback");
        assert_eq!(after.text(), Some("After"));

        let noop = rewrite(
            &output,
            ChartTextKind::Title,
            Some("After"),
            limits_for(&output, 5),
        )
        .expect("no-op rewrite");
        assert_eq!(noop.output(), output.as_slice());
        assert!(!noop.report().changed());
    }

    #[test]
    fn axis_rewrite_is_local_to_selected_axis() {
        let mut extension = varint_field(13, 1);
        extension.extend(text_field(15, b"Category"));
        extension.extend(varint_field(14, 1));
        extension.extend(text_field(16, b"Value"));
        extension.extend(text_field(99, b"unknown"));
        let source = outer(&extension);
        let output = rewrite(
            &source,
            ChartTextKind::Axis(Axis::Category),
            Some("New category"),
            limits_for(&source, 12),
        )
        .expect("axis rewrite")
        .into_output();
        let (category, _) = decode(
            &output,
            ChartTextKind::Axis(Axis::Category),
            Limits::for_source(&output),
        )
        .expect("category readback");
        let (value, _) = decode(
            &output,
            ChartTextKind::Axis(Axis::Value),
            Limits::for_source(&output),
        )
        .expect("value readback");
        assert_eq!(category.text(), Some("New category"));
        assert_eq!(value.text(), Some("Value"));
    }

    #[test]
    fn malformed_duplicate_extension_and_low_text_limit_are_typed() {
        let extension = {
            let mut value = varint_field(21, 1);
            value.extend(text_field(23, b"long title"));
            value
        };
        let mut duplicate = outer(&extension);
        duplicate.extend(outer(&extension));
        assert!(matches!(
            decode(
                &duplicate,
                ChartTextKind::Title,
                Limits::for_source(&duplicate)
            ),
            Err(Error::InvalidSource)
        ));

        let source = outer(&extension);
        let mut limits = Limits::for_source(&source);
        limits = Limits::new(
            limits.max_input_bytes(),
            limits.max_fields(),
            limits.max_output_bytes(),
            limits.max_nesting(),
            limits.max_work_bytes(),
            1,
            limits.max_allocations(),
            limits.max_retained_bytes(),
            limits.max_scratch_bytes(),
        );
        assert!(matches!(
            decode(&source, ChartTextKind::Title, limits),
            Err(Error::Limit {
                kind: LimitKind::TextBytes,
                ..
            })
        ));
    }

    fn limits_for(source: &[u8], replacement: usize) -> Limits {
        let output = source
            .len()
            .checked_add(replacement)
            .and_then(|value| value.checked_add(96))
            .expect("test output size");
        Limits::new(
            source.len().max(1),
            source.len().saturating_mul(8).max(1),
            output,
            8,
            source
                .len()
                .saturating_mul(16)
                .saturating_add(output)
                .max(1),
            output,
            output,
            source.len().saturating_add(output).max(1),
            output,
        )
    }

    fn source_field_end(source: &[u8], number: usize) -> usize {
        let mut offset = 0;
        for _ in 0..number {
            let key = source[offset];
            offset += 1;
            if key & 7 == 2 {
                let length = usize::from(source[offset]);
                offset += 1 + length;
            } else {
                offset += 1;
            }
        }
        offset
    }
}
