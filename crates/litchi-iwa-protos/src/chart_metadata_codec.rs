//! Strict borrowed metadata projections for modern and legacy chart records.
//!
//! The modern chart record is the `TSCH.ChartArchive` extension (field
//! `10000`) carried by `TSCH.ChartDrawableArchive` (message route 5021). The
//! legacy record is `TSCH.PreUFF.ChartInfoArchive` (message route 5000). This
//! module scans those envelopes with one aggregate budget, then forces a
//! private Buffa lazy view over the selected scalar/bytes closure. Labels and
//! row counts stay borrowed from the original bytes; no generated message or
//! repeated-field allocation crosses this boundary.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire preflight intentionally precedes the private view."
)]

use std::{fmt, num::NonZeroU64, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_chart_metadata_generated::LitchiIwaProjection as projection;

/// Native message route for `TSCH.ChartDrawableArchive`.
pub const MODERN_CHART_DRAWABLE_MESSAGE_TYPE: u32 = 5021;
/// Native message route for `TSCH.PreUFF.ChartInfoArchive`.
pub const LEGACY_CHART_INFO_MESSAGE_TYPE: u32 = 5000;
/// Native message route for `TSCH.ChartNonStyleArchive`.
pub const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5023;

const MODERN_EXTENSION_FIELD: u32 = 10_000;
const MODERN_CHART_TYPE_FIELD: u32 = 1;
const MODERN_DEFAULT_DATA_FIELD: u32 = 6;
const MODERN_GRID_FIELD: u32 = 7;
const MODERN_NON_STYLE_FIELD: u32 = 10;
const MODERN_ROW_NAME_FIELD: u32 = 1;
const MODERN_COLUMN_NAME_FIELD: u32 = 2;
const MODERN_GRID_ROW_FIELD: u32 = 3;

const LEGACY_CHART_MODEL_FIELD: u32 = 2;
const LEGACY_CHART_TYPE_FIELD: u32 = 4;
const LEGACY_NON_STYLE_FIELD: u32 = 14;
const LEGACY_MODEL_GRID_FIELD: u32 = 2;
const LEGACY_MODEL_INLINE_GRID_FIELD: u32 = 5;
const LEGACY_ROW_NAME_FIELD: u32 = 2;
const LEGACY_COLUMN_NAME_FIELD: u32 = 3;
const LEGACY_VALUE_ROW_FIELD: u32 = 4;

const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MODERN_GRID_ROW_NAME: &str = "TSCH.ChartGridArchive.row_name";
const MODERN_GRID_COLUMN_NAME: &str = "TSCH.ChartGridArchive.column_name";
const LEGACY_GRID_ROW_NAME: &str = "TSCH.PreUFF.ChartGridArchive.row_name";
const LEGACY_GRID_COLUMN_NAME: &str = "TSCH.PreUFF.ChartGridArchive.column_name";
const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;

/// Whether a snapshot came from the modern generated chart extension or the
/// pre-UFF chart-info message.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartMetadataFormat {
    /// `TSCH.ChartDrawableArchive` extension field 10000.
    Modern,
    /// `TSCH.PreUFF.ChartInfoArchive`.
    Legacy,
}

/// Finite limits shared by the complete metadata decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_depth: u32,
    max_label_count: usize,
    max_text_bytes: usize,
}

impl DecodeOptions {
    /// Build an explicit aggregate bytes/fields/work/depth/label policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_depth: u32,
        max_label_count: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_fields,
            max_work_bytes,
            max_depth,
            max_label_count,
            max_text_bytes,
        }
    }

    /// Build a finite policy from one caller-owned archive payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        let fields = bytes
            .checked_mul(8)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .clamp(1, MAX_AUTOMATIC_LIMIT);
        let work = bytes
            .checked_mul(32)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .clamp(1, MAX_AUTOMATIC_LIMIT);
        Self::new(bytes, fields, work, 16, bytes, bytes)
    }

    /// Aggregate encoded input-byte ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Aggregate encoded field-visit ceiling.
    #[must_use]
    pub const fn max_fields(self) -> usize {
        self.max_fields
    }

    /// Aggregate strict-plus-Buffa work ceiling.
    #[must_use]
    pub const fn max_work_bytes(self) -> usize {
        self.max_work_bytes
    }

    /// Maximum protobuf nesting depth.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Maximum number of selected row and column labels.
    #[must_use]
    pub const fn max_label_count(self) -> usize {
        self.max_label_count
    }

    /// Maximum aggregate UTF-8 label bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Replace the input-byte ceiling.
    #[must_use]
    pub const fn with_max_input_bytes(mut self, maximum: usize) -> Self {
        self.max_input_bytes = maximum;
        self
    }

    /// Compatibility spelling for message-oriented callers.
    #[must_use]
    pub const fn with_max_message_bytes(self, maximum: usize) -> Self {
        self.with_max_input_bytes(maximum)
    }

    /// Replace the field-visit ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the aggregate work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the nesting ceiling.
    #[must_use]
    pub const fn with_max_depth(mut self, maximum: u32) -> Self {
        self.max_depth = maximum;
        self
    }

    /// Compatibility spelling for callers using Buffa terminology.
    #[must_use]
    pub const fn with_recursion_limit(self, maximum: u32) -> Self {
        self.with_max_depth(maximum)
    }

    /// Replace the aggregate row/column label-count ceiling.
    #[must_use]
    pub const fn with_max_label_count(mut self, maximum: usize) -> Self {
        self.max_label_count = maximum;
        self
    }

    /// Replace the aggregate UTF-8 label-byte ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.max_depth)
    }
}

/// A lazily reparsed borrowed list of repeated UTF-8 labels.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct LabelList<'source> {
    source: &'source [u8],
    field_number: u32,
    length: usize,
}

impl fmt::Debug for LabelList<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("LabelList")
            .field("field_number", &self.field_number)
            .field("length", &self.length)
            .field("source_len", &self.source.len())
            .finish()
    }
}

impl<'source> LabelList<'source> {
    const fn empty() -> Self {
        Self {
            source: &[],
            field_number: 0,
            length: 0,
        }
    }

    const fn new(source: &'source [u8], field_number: u32, length: usize) -> Self {
        Self {
            source,
            field_number,
            length,
        }
    }

    /// Number of validated labels.
    #[must_use]
    pub const fn len(self) -> usize {
        self.length
    }

    /// Whether no label was present.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.length == 0
    }

    /// Return one borrowed label by ordinal.
    #[must_use]
    pub fn get(self, index: usize) -> Option<&'source str> {
        (index < self.length)
            .then(|| self.iter().nth(index))
            .flatten()
    }

    /// Iterate labels in native wire order.
    #[must_use]
    pub fn iter(self) -> LabelIter<'source> {
        LabelIter {
            remaining: self.source,
            field_number: self.field_number,
            emitted: 0,
            length: self.length,
        }
    }
}

/// Iterator returned by [`LabelList::iter`].
#[derive(Clone, Copy)]
pub struct LabelIter<'source> {
    remaining: &'source [u8],
    field_number: u32,
    emitted: usize,
    length: usize,
}

impl fmt::Debug for LabelIter<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("LabelIter")
            .field("field_number", &self.field_number)
            .field("emitted", &self.emitted)
            .field("length", &self.length)
            .field("remaining_len", &self.remaining.len())
            .finish()
    }
}

impl<'source> Iterator for LabelIter<'source> {
    type Item = &'source str;

    fn next(&mut self) -> Option<Self::Item> {
        while self.emitted < self.length && !self.remaining.is_empty() {
            let field = parse_unchecked_field(&mut self.remaining).ok()??;
            if field.number != self.field_number {
                continue;
            }
            let WireValue::LengthDelimited(bytes) = field.value else {
                return None;
            };
            let text = str::from_utf8(bytes).ok()?;
            self.emitted += 1;
            return Some(text);
        }
        None
    }
}

/// Borrowed semantic metadata selected from one chart record.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ChartMetadataSnapshot<'source> {
    format: ChartMetadataFormat,
    chart_type: i32,
    contains_default_data: Option<bool>,
    row_labels: LabelList<'source>,
    column_labels: LabelList<'source>,
    series_count: usize,
    non_style_ref: Option<NonZeroU64>,
    source: &'source [u8],
    grid_source: Option<&'source [u8]>,
}

impl fmt::Debug for ChartMetadataSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartMetadataSnapshot")
            .field("format", &self.format)
            .field("chart_type", &self.chart_type)
            .field("contains_default_data", &self.contains_default_data)
            .field("row_label_count", &self.row_labels.len())
            .field("column_label_count", &self.column_labels.len())
            .field("series_count", &self.series_count)
            .field("non_style_ref", &self.non_style_ref)
            .field("source_len", &self.source.len())
            .field("grid_source_len", &self.grid_source.map_or(0, <[u8]>::len))
            .finish()
    }
}

impl<'source> ChartMetadataSnapshot<'source> {
    /// Source format selected by the decoder.
    #[must_use]
    pub const fn format(self) -> ChartMetadataFormat {
        self.format
    }

    /// Raw native chart-kind enum value.
    #[must_use]
    pub const fn chart_type(self) -> i32 {
        self.chart_type
    }

    /// Modern default-data presence/value. Legacy records return `None`.
    #[must_use]
    pub const fn contains_default_data(self) -> Option<bool> {
        self.contains_default_data
    }

    /// Borrowed row labels.
    #[must_use]
    pub const fn row_labels(self) -> LabelList<'source> {
        self.row_labels
    }

    /// Borrowed column labels.
    #[must_use]
    pub const fn column_labels(self) -> LabelList<'source> {
        self.column_labels
    }

    /// Number of native grid rows/value rows.
    #[must_use]
    pub const fn series_count(self) -> usize {
        self.series_count
    }

    /// Optional non-style object identifier, preserving absence.
    #[must_use]
    pub const fn non_style_ref(self) -> Option<NonZeroU64> {
        self.non_style_ref
    }

    /// Exact caller-owned source bytes.
    #[must_use]
    pub const fn source(self) -> &'source [u8] {
        self.source
    }

    /// Exact embedded grid bytes, when the source contained a grid.
    #[must_use]
    pub const fn grid_source(self) -> Option<&'source [u8]> {
        self.grid_source
    }
}

/// Aggregate accounting for one complete metadata decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    failure_work_bytes: usize,
    max_depth: u32,
    label_count: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
}

impl DecodeReport {
    /// Encoded root input bytes retained by the snapshot.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Strictly visited wire fields, including unknown fields and groups.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa work charged before success/failure.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Additional work charged while publishing a rejected decode.
    ///
    /// `work_bytes` already includes every attempted operation, including the
    /// operation that caused the failure, so this additive field is zero.
    #[must_use]
    pub const fn failure_work_bytes(self) -> usize {
        self.failure_work_bytes
    }

    /// Maximum nested message depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Selected row/column labels visited.
    #[must_use]
    pub const fn label_count(self) -> usize {
        self.label_count
    }

    /// Aggregate UTF-8 label bytes visited.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Generated-view allocations. Repeated grids remain borrowed bytes.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by the snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
}

/// Resource axis rejected by strict metadata ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Aggregate source bytes exceeded their ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Field visits exceeded their ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate strict-plus-Buffa work exceeded its ceiling.
    Work { observed: usize, maximum: usize },
    /// Message nesting exceeded its ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Selected label count exceeded its ceiling.
    Labels { observed: usize, maximum: usize },
    /// Selected UTF-8 bytes exceeded their ceiling.
    Text { observed: usize, maximum: usize },
}

/// Failure from strict metadata preflight or its private Buffa view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
    report: DecodeReport,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Limit(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    InvalidUtf8(&'static str),
    Projection,
}

impl DecodeError {
    /// Aggregate accounting observed before this failure.
    #[must_use]
    pub const fn report(&self) -> DecodeReport {
        self.report
    }

    /// Typed resource limit, when applicable.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            _ => None,
        }
    }

    /// Return the missing required field, if applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            _ => None,
        }
    }

    /// Whether the modern unity extension was absent.
    #[must_use]
    pub fn is_missing_chart_extension(&self) -> bool {
        matches!(&self.kind, DecodeErrorKind::MissingRequired(field) if *field == "TSCH.ChartDrawableArchive.unity")
    }

    /// Return the duplicated singular field, if applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Return the canonical-wire failure reason, if applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            _ => None,
        }
    }

    /// Return the invalid UTF-8 field, if applicable.
    #[must_use]
    pub const fn invalid_utf8_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::InvalidUtf8(field) => Some(field),
            _ => None,
        }
    }

    /// Return the invalid reference field, if applicable.
    #[must_use]
    pub const fn invalid_reference_field(&self) -> Option<&'static str> {
        None
    }

    /// Byte-limit values used by host adapters.
    #[must_use]
    pub const fn input_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Field-limit values used by host adapters.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Work-limit values used by host adapters.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Label-count limit values used by host adapters.
    #[must_use]
    pub const fn label_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Labels { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// UTF-8 text-limit values used by host adapters.
    #[must_use]
    pub const fn text_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Text { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Nesting-limit values used by host adapters.
    #[must_use]
    pub const fn depth_limit_values(&self) -> Option<(u32, u32)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// No allocation is performed by this borrowed projection.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        None
    }

    const fn new(kind: DecodeErrorKind, report: DecodeReport) -> Self {
        Self { kind, report }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "chart metadata input bytes {observed} exceed maximum {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "chart metadata visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "chart metadata requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "chart metadata nesting {observed} exceeds maximum {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Labels { observed, maximum }) => write!(
                formatter,
                "chart metadata contains {observed} labels; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Text { observed, maximum }) => write!(
                formatter,
                "chart metadata contains {observed} UTF-8 bytes; maximum is {maximum}"
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
            DecodeErrorKind::InvalidUtf8(field) => write!(formatter, "{field} is invalid UTF-8"),
            DecodeErrorKind::Projection => formatter
                .write_str("chart metadata strict preflight disagrees with the Buffa projection"),
        }
    }
}

impl std::error::Error for DecodeError {}

/// Decode modern chart metadata with an aggregate report.
pub fn decode_modern<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<ChartMetadataSnapshot<'source>, DecodeError> {
    Ok(decode_modern_with_report(source, options)?.0)
}

/// Decode modern chart metadata and return complete aggregate accounting.
pub fn decode_modern_with_report<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<(ChartMetadataSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(*options);
    let result = preflight_modern(source, &mut budget)
        .and_then(|strict| project_modern(strict, *options, &mut budget));
    match result {
        Ok(snapshot) => Ok((snapshot, budget.finish_success())),
        Err(kind) => Err(budget.finish_error(kind)),
    }
}

/// Decode legacy pre-UFF chart metadata.
pub fn decode_legacy<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<ChartMetadataSnapshot<'source>, DecodeError> {
    Ok(decode_legacy_with_report(source, options)?.0)
}

/// Decode legacy pre-UFF chart metadata and return complete accounting.
pub fn decode_legacy_with_report<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<(ChartMetadataSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(*options);
    let result = preflight_legacy(source, &mut budget)
        .and_then(|strict| project_legacy(strict, source, *options, &mut budget));
    match result {
        Ok(snapshot) => Ok((snapshot, budget.finish_success())),
        Err(kind) => Err(budget.finish_error(kind)),
    }
}

#[derive(Debug, Clone, Copy)]
struct StrictModern<'source> {
    chart_type: i32,
    chart_type_present: bool,
    contains_default_data: Option<bool>,
    row_labels: LabelList<'source>,
    column_labels: LabelList<'source>,
    series_count: usize,
    non_style_ref: Option<NonZeroU64>,
    grid_source: Option<&'source [u8]>,
    chart_source: &'source [u8],
    root_source: &'source [u8],
}

#[derive(Debug, Clone, Copy)]
struct StrictLegacy<'source> {
    chart_type: i32,
    row_labels: LabelList<'source>,
    column_labels: LabelList<'source>,
    series_count: usize,
    non_style_ref: Option<NonZeroU64>,
    grid_source: Option<&'source [u8]>,
}

fn preflight_modern<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictModern<'source>, DecodeErrorKind> {
    budget.root_message(source.len(), 1)?;
    let mut extension = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 1, budget)? {
        if field.number != MODERN_EXTENSION_FIELD {
            continue;
        }
        if extension.is_some() {
            return Err(DecodeErrorKind::DuplicateSingular(
                "TSCH.ChartDrawableArchive.unity",
            ));
        }
        extension = Some(field.length_delimited()?);
    }
    let chart_source = extension.ok_or(DecodeErrorKind::MissingRequired(
        "TSCH.ChartDrawableArchive.unity",
    ))?;
    budget.nested_message(chart_source.len(), 2)?;
    let mut strict = preflight_modern_chart(chart_source, budget)?;
    strict.chart_source = chart_source;
    strict.root_source = source;
    Ok(strict)
}

fn preflight_modern_chart<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictModern<'source>, DecodeErrorKind> {
    let mut chart_type = None;
    let mut contains_default_data = None;
    let mut grid = None;
    let mut row_labels = LabelList::empty();
    let mut column_labels = LabelList::empty();
    let mut series_count = 0usize;
    let mut non_style_ref = None;
    let mut non_style_present = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 2, budget)? {
        match field.number {
            MODERN_CHART_TYPE_FIELD => {
                if chart_type.is_some() {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.ChartArchive.chart_type",
                    ));
                }
                chart_type = Some(require_int32(field.varint()?)?);
            },
            MODERN_DEFAULT_DATA_FIELD => {
                if contains_default_data.is_some() {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.ChartArchive.contains_default_data",
                    ));
                }
                contains_default_data = Some(require_bool(field.varint()?)?);
            },
            MODERN_GRID_FIELD => {
                if grid.is_some() {
                    return Err(DecodeErrorKind::DuplicateSingular("TSCH.ChartArchive.grid"));
                }
                let payload = field.length_delimited()?;
                budget.nested_message(payload.len(), 3)?;
                let strict = preflight_modern_grid(payload, budget)?;
                row_labels = strict.row_labels;
                column_labels = strict.column_labels;
                series_count = strict.series_count;
                grid = Some(payload);
            },
            MODERN_NON_STYLE_FIELD => {
                if non_style_present {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.ChartArchive.chart_non_style",
                    ));
                }
                non_style_present = true;
                non_style_ref = preflight_reference(field.length_delimited()?, 3, budget)?;
            },
            _ => {},
        }
    }
    Ok(StrictModern {
        chart_type: chart_type.unwrap_or_default(),
        chart_type_present: chart_type.is_some(),
        contains_default_data,
        row_labels,
        column_labels,
        series_count,
        non_style_ref,
        grid_source: grid,
        chart_source: source,
        root_source: source,
    })
}

#[derive(Debug, Clone, Copy)]
struct StrictGrid<'source> {
    row_labels: LabelList<'source>,
    column_labels: LabelList<'source>,
    series_count: usize,
    grid_source: Option<&'source [u8]>,
}

fn preflight_modern_grid<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictGrid<'source>, DecodeErrorKind> {
    let mut row_labels = 0usize;
    let mut column_labels = 0usize;
    let mut series_count = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 3, budget)? {
        match field.number {
            MODERN_ROW_NAME_FIELD => {
                let bytes = field.length_delimited()?;
                budget.label(bytes.len(), MODERN_GRID_ROW_NAME)?;
                str::from_utf8(bytes)
                    .map_err(|_error| DecodeErrorKind::InvalidUtf8(MODERN_GRID_ROW_NAME))?;
                row_labels = row_labels.saturating_add(1);
            },
            MODERN_COLUMN_NAME_FIELD => {
                let bytes = field.length_delimited()?;
                budget.label(bytes.len(), MODERN_GRID_COLUMN_NAME)?;
                str::from_utf8(bytes)
                    .map_err(|_error| DecodeErrorKind::InvalidUtf8(MODERN_GRID_COLUMN_NAME))?;
                column_labels = column_labels.saturating_add(1);
            },
            MODERN_GRID_ROW_FIELD => {
                let payload = field.length_delimited()?;
                series_count = series_count.saturating_add(1);
                budget.nested_message(payload.len(), 4)?;
            },
            _ => {},
        }
    }
    Ok(StrictGrid {
        row_labels: LabelList::new(source, MODERN_ROW_NAME_FIELD, row_labels),
        column_labels: LabelList::new(source, MODERN_COLUMN_NAME_FIELD, column_labels),
        series_count,
        grid_source: Some(source),
    })
}

fn preflight_legacy<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictLegacy<'source>, DecodeErrorKind> {
    budget.root_message(source.len(), 1)?;
    let mut chart_model = None;
    let mut chart_type = None;
    let mut non_style_ref = None;
    let mut non_style_present = false;
    let mut row_labels = LabelList::empty();
    let mut column_labels = LabelList::empty();
    let mut grid_source = None;
    let mut series_count = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 1, budget)? {
        match field.number {
            LEGACY_CHART_MODEL_FIELD => {
                if chart_model.is_some() {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.PreUFF.ChartInfoArchive.chart_model",
                    ));
                }
                let payload = field.length_delimited()?;
                budget.nested_message(payload.len(), 2)?;
                let model = preflight_legacy_model(payload, budget)?;
                row_labels = model.row_labels;
                column_labels = model.column_labels;
                series_count = model.series_count;
                grid_source = model.grid_source;
                chart_model = Some(payload);
            },
            LEGACY_CHART_TYPE_FIELD => {
                if chart_type.is_some() {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.PreUFF.ChartInfoArchive.chart_type",
                    ));
                }
                chart_type = Some(require_int32(field.varint()?)?);
            },
            LEGACY_NON_STYLE_FIELD => {
                if non_style_present {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.PreUFF.ChartInfoArchive.non_style",
                    ));
                }
                non_style_present = true;
                non_style_ref = preflight_reference(field.length_delimited()?, 2, budget)?;
            },
            _ => {},
        }
    }
    if chart_model.is_none() {
        return Err(DecodeErrorKind::MissingRequired(
            "TSCH.PreUFF.ChartInfoArchive.chart_model",
        ));
    }
    let chart_type = chart_type.ok_or(DecodeErrorKind::MissingRequired(
        "TSCH.PreUFF.ChartInfoArchive.chart_type",
    ))?;
    Ok(StrictLegacy {
        chart_type,
        row_labels,
        column_labels,
        series_count,
        non_style_ref,
        grid_source,
    })
}

fn preflight_legacy_model<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictGrid<'source>, DecodeErrorKind> {
    let mut row_labels = LabelList::empty();
    let mut column_labels = LabelList::empty();
    let mut grid_source = None;
    let mut series_count = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 2, budget)? {
        match field.number {
            LEGACY_MODEL_GRID_FIELD => {
                let _ = field.length_delimited()?;
            },
            LEGACY_MODEL_INLINE_GRID_FIELD => {
                if grid_source.is_some() {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.PreUFF.ChartModelArchive.inline_grid",
                    ));
                }
                let payload = field.length_delimited()?;
                budget.nested_message(payload.len(), 3)?;
                let grid = preflight_legacy_grid(payload, budget)?;
                row_labels = grid.row_labels;
                column_labels = grid.column_labels;
                series_count = grid.series_count;
                grid_source = Some(payload);
            },
            _ => {},
        }
    }
    Ok(StrictGrid {
        row_labels,
        column_labels,
        series_count,
        grid_source,
    })
}

fn preflight_legacy_grid<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictGrid<'source>, DecodeErrorKind> {
    let mut row_labels = 0usize;
    let mut column_labels = 0usize;
    let mut series_count = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 3, budget)? {
        match field.number {
            1 => {
                require_int32(field.varint()?)?;
            },
            6 => {
                require_bool(field.varint()?)?;
            },
            LEGACY_ROW_NAME_FIELD => {
                let bytes = field.length_delimited()?;
                budget.label(bytes.len(), LEGACY_GRID_ROW_NAME)?;
                str::from_utf8(bytes)
                    .map_err(|_error| DecodeErrorKind::InvalidUtf8(LEGACY_GRID_ROW_NAME))?;
                row_labels = row_labels.saturating_add(1);
            },
            LEGACY_COLUMN_NAME_FIELD => {
                let bytes = field.length_delimited()?;
                budget.label(bytes.len(), LEGACY_GRID_COLUMN_NAME)?;
                str::from_utf8(bytes)
                    .map_err(|_error| DecodeErrorKind::InvalidUtf8(LEGACY_GRID_COLUMN_NAME))?;
                column_labels = column_labels.saturating_add(1);
            },
            LEGACY_VALUE_ROW_FIELD => {
                let payload = field.length_delimited()?;
                series_count = series_count.saturating_add(1);
                budget.nested_message(payload.len(), 4)?;
            },
            _ => {},
        }
    }
    Ok(StrictGrid {
        row_labels: LabelList::new(source, LEGACY_ROW_NAME_FIELD, row_labels),
        column_labels: LabelList::new(source, LEGACY_COLUMN_NAME_FIELD, column_labels),
        series_count,
        grid_source: Some(source),
    })
}

fn project_modern<'source>(
    strict: StrictModern<'source>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartMetadataSnapshot<'source>, DecodeErrorKind> {
    budget.buffa(strict.chart_source.len())?;
    let view: projection::ModernChartArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(strict.chart_source)
        .map_err(DecodeErrorKind::Wire)?;
    let projected_non_style = view
        .chart_non_style
        .get()
        .map_err(DecodeErrorKind::Wire)?
        .and_then(|reference| NonZeroU64::new(reference.identifier));
    if view.chart_type != strict.chart_type_present.then_some(strict.chart_type)
        || view.contains_default_data != strict.contains_default_data
        || !same_borrowed_bytes(view.grid, strict.grid_source)
        || projected_non_style != strict.non_style_ref
    {
        return Err(DecodeErrorKind::Projection);
    }
    Ok(ChartMetadataSnapshot {
        format: ChartMetadataFormat::Modern,
        chart_type: strict.chart_type,
        contains_default_data: strict.contains_default_data,
        row_labels: strict.row_labels,
        column_labels: strict.column_labels,
        series_count: strict.series_count,
        non_style_ref: strict.non_style_ref,
        source: strict.root_source,
        grid_source: strict.grid_source,
    })
}

fn project_legacy<'source>(
    strict: StrictLegacy<'source>,
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartMetadataSnapshot<'source>, DecodeErrorKind> {
    budget.buffa(source.len())?;
    let view: projection::LegacyChartInfoArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeErrorKind::Wire)?;
    let model = view
        .chart_model
        .get()
        .map_err(DecodeErrorKind::Wire)?
        .ok_or(DecodeErrorKind::Projection)?;
    let projected_non_style = view
        .non_style
        .get()
        .map_err(DecodeErrorKind::Wire)?
        .and_then(|reference| NonZeroU64::new(reference.identifier));
    if view.chart_type != strict.chart_type
        || !same_borrowed_bytes(model.inline_grid, strict.grid_source)
        || projected_non_style != strict.non_style_ref
    {
        return Err(DecodeErrorKind::Projection);
    }
    Ok(ChartMetadataSnapshot {
        format: ChartMetadataFormat::Legacy,
        chart_type: strict.chart_type,
        contains_default_data: None,
        row_labels: strict.row_labels,
        column_labels: strict.column_labels,
        series_count: strict.series_count,
        non_style_ref: strict.non_style_ref,
        source,
        grid_source: strict.grid_source,
    })
}

/// Compare two source-backed payloads without walking the payload a second
/// time after strict preflight.  The generated lazy view must expose the
/// exact validated span, so pointer identity and length are the projection
/// parity check.
fn same_borrowed_bytes(left: Option<&[u8]>, right: Option<&[u8]>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.as_ptr() == right.as_ptr() && left.len() == right.len(),
        (None, None) => true,
        _ => false,
    }
}

#[derive(Debug)]
struct Budget {
    options: DecodeOptions,
    report: DecodeReport,
}

impl Budget {
    fn new(options: DecodeOptions) -> Self {
        Self {
            options,
            report: DecodeReport::default(),
        }
    }

    fn root_message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeErrorKind> {
        self.report.source_bytes = bytes;
        self.report.retained_bytes = bytes;
        if bytes > self.options.max_input_bytes {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Bytes {
                observed: bytes,
                maximum: self.options.max_input_bytes,
            }));
        }
        self.nested_message(bytes, depth)
    }

    fn nested_message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeErrorKind> {
        self.observe_depth(depth)?;
        self.work(bytes.saturating_mul(2))
    }

    fn buffa(&mut self, bytes: usize) -> Result<(), DecodeErrorKind> {
        self.work(bytes)
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeErrorKind> {
        self.report.max_depth = self.report.max_depth.max(depth);
        if depth > self.options.max_depth {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.max_depth,
            }));
        }
        if depth > MAX_RECURSION_LIMIT {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: MAX_RECURSION_LIMIT,
            }));
        }
        Ok(())
    }

    fn field(&mut self) -> Result<(), DecodeErrorKind> {
        let observed = self.report.fields.saturating_add(1);
        self.report.fields = observed;
        if observed > self.options.max_fields {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }

    fn work(&mut self, bytes: usize) -> Result<(), DecodeErrorKind> {
        let observed = self.report.work_bytes.saturating_add(bytes);
        self.report.work_bytes = observed;
        if observed > self.options.max_work_bytes {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn label(&mut self, bytes: usize, _field: &'static str) -> Result<(), DecodeErrorKind> {
        let labels = self.report.label_count.saturating_add(1);
        if labels > self.options.max_label_count {
            self.report.label_count = labels;
            return Err(DecodeErrorKind::Limit(DecodeLimit::Labels {
                observed: labels,
                maximum: self.options.max_label_count,
            }));
        }
        // Publish the attempted label before checking its text budget so a
        // rejected decode reports both counters consistently.
        self.report.label_count = labels;
        let text = self.report.text_bytes.saturating_add(bytes);
        if text > self.options.max_text_bytes {
            self.report.text_bytes = text;
            return Err(DecodeErrorKind::Limit(DecodeLimit::Text {
                observed: text,
                maximum: self.options.max_text_bytes,
            }));
        }
        self.report.text_bytes = text;
        Ok(())
    }

    fn finish_success(self) -> DecodeReport {
        self.report
    }

    fn finish_error(self, kind: DecodeErrorKind) -> DecodeError {
        DecodeError::new(kind, self.report)
    }
}

#[derive(Debug, Clone, Copy)]
enum WireValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
}

#[derive(Debug, Clone, Copy)]
struct WireField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: WireValue<'source>,
}

impl<'source> WireField<'source> {
    fn varint(self) -> Result<u64, DecodeErrorKind> {
        if self.wire_type != buffa::encoding::WireType::Varint {
            return Err(DecodeErrorKind::Wire(
                buffa::DecodeError::WireTypeMismatch {
                    field_number: self.number,
                    expected: buffa::encoding::WireType::Varint as u8,
                    actual: self.wire_type as u8,
                },
            ));
        }
        let WireValue::Varint(value) = self.value else {
            return Err(DecodeErrorKind::Projection);
        };
        Ok(value)
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeErrorKind> {
        if self.wire_type != buffa::encoding::WireType::LengthDelimited {
            return Err(DecodeErrorKind::Wire(
                buffa::DecodeError::WireTypeMismatch {
                    field_number: self.number,
                    expected: buffa::encoding::WireType::LengthDelimited as u8,
                    actual: self.wire_type as u8,
                },
            ));
        }
        let WireValue::LengthDelimited(value) = self.value else {
            return Err(DecodeErrorKind::Projection);
        };
        Ok(value)
    }
}

#[derive(Debug, Clone, Copy)]
enum ParseItem<'source> {
    Field(WireField<'source>),
    EndGroup(u32),
}

fn next_field<'source>(
    source: &mut &'source [u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<WireField<'source>>, DecodeErrorKind> {
    match parse_field(source, depth, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => Err(DecodeErrorKind::Wire(
            buffa::DecodeError::InvalidEndGroup(number),
        )),
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeErrorKind> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    let (encoded_tag, canonical) = take_varint(source)?;
    if !canonical {
        return Err(DecodeErrorKind::NonCanonical("protobuf field key"));
    }
    budget.field()?;
    let raw_tag = u32::try_from(encoded_tag)
        .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::InvalidFieldNumber))?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(DecodeErrorKind::Wire(
            buffa::DecodeError::InvalidFieldNumber,
        ));
    }
    let raw_wire = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire).map_err(DecodeErrorKind::Wire)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeErrorKind::NonCanonical("protobuf varint value"));
            }
            WireValue::Varint(value)
        },
        buffa::encoding::WireType::Fixed64 => {
            take_exact(source, 8)?;
            WireValue::Fixed64
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeErrorKind::NonCanonical("length-delimited size"));
            }
            let length = usize::try_from(length)
                .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::MessageTooLarge))?;
            WireValue::LengthDelimited(take_exact(source, length)?)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_depth = depth.saturating_add(1);
            budget.observe_depth(child_depth)?;
            skip_group(source, field_number, child_depth, budget)?;
            WireValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            WireValue::Fixed32
        },
        _ => {
            return Err(DecodeErrorKind::Wire(buffa::DecodeError::InvalidWireType(
                raw_wire,
            )));
        },
    };
    Ok(Some(ParseItem::Field(WireField {
        number: field_number,
        wire_type,
        value,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<(), DecodeErrorKind> {
    loop {
        match parse_field(source, depth, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(DecodeErrorKind::Wire(buffa::DecodeError::InvalidEndGroup(
                    number,
                )));
            },
            None => return Err(DecodeErrorKind::Wire(buffa::DecodeError::UnexpectedEof)),
        }
    }
}

fn preflight_reference(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<Option<NonZeroU64>, DecodeErrorKind> {
    budget.nested_message(source.len(), depth)?;
    let mut identifier = None;
    let mut saw_identifier = false;
    let mut saw_deprecated_type = false;
    let mut saw_deprecated_external = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, depth, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if saw_identifier {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSP.Reference.identifier",
                    ));
                }
                let value = field.varint()?;
                saw_identifier = true;
                identifier = NonZeroU64::new(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if saw_deprecated_type {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                saw_deprecated_type = true;
                require_int32(field.varint()?)?;
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if saw_deprecated_external {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                saw_deprecated_external = true;
                require_bool(field.varint()?)?;
            },
            _ => {},
        }
    }
    if saw_identifier {
        Ok(identifier)
    } else {
        Err(DecodeErrorKind::MissingRequired("TSP.Reference.identifier"))
    }
}

fn require_bool(value: u64) -> Result<bool, DecodeErrorKind> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeErrorKind::NonCanonical(
            "bool scalar is not zero or one",
        )),
    }
}

fn require_int32(value: u64) -> Result<i32, DecodeErrorKind> {
    if value <= i32::MAX as u64 {
        return Ok(value as i32);
    }
    if value >= 0xffff_ffff_8000_0000 {
        return Ok(value as i64 as i32);
    }
    Err(DecodeErrorKind::NonCanonical("int32 scalar encoding"))
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeErrorKind> {
    let mut value = 0u64;
    for index in 0..10 {
        let byte = *source
            .first()
            .ok_or(DecodeErrorKind::Wire(buffa::DecodeError::UnexpectedEof))?;
        *source = &source[1..];
        if index == 9 && byte > 1 {
            return Err(DecodeErrorKind::Wire(buffa::DecodeError::VarintTooLong));
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let canonical = index == 0 || value >= (1u64 << (index * 7));
            return Ok((value, canonical));
        }
    }
    Err(DecodeErrorKind::Wire(buffa::DecodeError::VarintTooLong))
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeErrorKind> {
    if source.len() < length {
        return Err(DecodeErrorKind::Wire(buffa::DecodeError::UnexpectedEof));
    }
    let (head, tail) = source.split_at(length);
    *source = tail;
    Ok(head)
}

fn parse_unchecked_field<'source>(
    source: &mut &'source [u8],
) -> Result<Option<WireField<'source>>, DecodeErrorKind> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical) = take_varint(source)?;
    if !canonical {
        return Err(DecodeErrorKind::NonCanonical("protobuf field key"));
    }
    let raw_tag = u32::try_from(encoded_tag)
        .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::InvalidFieldNumber))?;
    let number = raw_tag >> 3;
    if number == 0 {
        return Err(DecodeErrorKind::Wire(
            buffa::DecodeError::InvalidFieldNumber,
        ));
    }
    let wire_type =
        buffa::encoding::WireType::from_u32(raw_tag & 7).map_err(DecodeErrorKind::Wire)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => WireValue::Varint(take_varint(source)?.0),
        buffa::encoding::WireType::Fixed64 => {
            take_exact(source, 8)?;
            WireValue::Fixed64
        },
        buffa::encoding::WireType::LengthDelimited => {
            let length = usize::try_from(take_varint(source)?.0)
                .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::MessageTooLarge))?;
            WireValue::LengthDelimited(take_exact(source, length)?)
        },
        buffa::encoding::WireType::StartGroup => {
            skip_unchecked_group(source, number)?;
            WireValue::Group
        },
        buffa::encoding::WireType::EndGroup => return Ok(None),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            WireValue::Fixed32
        },
        _ => {
            return Err(DecodeErrorKind::Wire(buffa::DecodeError::InvalidWireType(
                raw_tag & 7,
            )));
        },
    };
    Ok(Some(WireField {
        number,
        wire_type,
        value,
    }))
}

fn skip_unchecked_group(source: &mut &[u8], expected: u32) -> Result<(), DecodeErrorKind> {
    while !source.is_empty() {
        let (tag, _) = take_varint(source)?;
        let number = u32::try_from(tag >> 3)
            .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::InvalidFieldNumber))?;
        let wire =
            buffa::encoding::WireType::from_u32((tag & 7) as u32).map_err(DecodeErrorKind::Wire)?;
        if wire == buffa::encoding::WireType::EndGroup {
            if number == expected {
                return Ok(());
            }
            return Err(DecodeErrorKind::Wire(buffa::DecodeError::InvalidEndGroup(
                number,
            )));
        }
        match wire {
            buffa::encoding::WireType::Varint => {
                take_varint(source)?;
            },
            buffa::encoding::WireType::Fixed64 => {
                take_exact(source, 8)?;
            },
            buffa::encoding::WireType::LengthDelimited => {
                let length = usize::try_from(take_varint(source)?.0)
                    .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::MessageTooLarge))?;
                take_exact(source, length)?;
            },
            buffa::encoding::WireType::StartGroup => skip_unchecked_group(source, number)?,
            buffa::encoding::WireType::Fixed32 => {
                take_exact(source, 4)?;
            },
            buffa::encoding::WireType::EndGroup => unreachable!(),
            _ => {
                return Err(DecodeErrorKind::Wire(buffa::DecodeError::InvalidWireType(
                    (tag & 7) as u32,
                )));
            },
        }
    }
    Err(DecodeErrorKind::Wire(buffa::DecodeError::UnexpectedEof))
}

#[cfg(test)]
mod tests;
