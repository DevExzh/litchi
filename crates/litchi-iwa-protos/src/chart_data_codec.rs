//! Strict borrowed decoding of modern iWork chart grid data.
//!
//! A modern chart stores its data in the `TSCH.ChartArchive` unity extension
//! (field 10000) on a `TSCH.ChartDrawableArchive` message (route 5021).  The
//! selected `ChartGridArchive` is scanned by hand so row and value repetitions
//! remain borrowed source spans.  Buffa is used only for a bounded lazy view
//! of the selected chart `grid` span and each selected `GridValue` scalar.
//! Unknown fields, the native row/column id map, and all future data fields
//! stay source-authoritative.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire preflight intentionally precedes the private view."
)]

use std::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_chart_data_generated::LitchiIwaProjection as projection;

/// Native route for `TSCH.ChartDrawableArchive`.
pub const MODERN_CHART_DRAWABLE_MESSAGE_TYPE: u32 = 5_021;
/// Native chart unity extension field on `TSCH.ChartDrawableArchive`.
pub const MODERN_CHART_EXTENSION_FIELD: u32 = 10_000;

const MODERN_GRID_FIELD: u32 = 7;
const GRID_ROW_NAME_FIELD: u32 = 1;
const GRID_COLUMN_NAME_FIELD: u32 = 2;
const GRID_ROW_FIELD: u32 = 3;
const GRID_ID_MAP_FIELD: u32 = 4;
const GRID_VALUE_FIELD: u32 = 1;
const GRID_VALUE_DATE_1_0_FIELD: u32 = 2;
const GRID_VALUE_DURATION_FIELD: u32 = 3;
const GRID_VALUE_DATE_FIELD: u32 = 4;

const MAX_RECURSION_LIMIT: u32 = 64;
const MAX_AUTOMATIC_LIMIT: usize = 64 * 1024 * 1024;

/// Finite limits applied to one complete modern chart-data decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_input_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_depth: u32,
    max_cells: usize,
    max_label_count: usize,
    max_text_bytes: usize,
    max_output_bytes: usize,
    max_allocations: usize,
    max_retained_bytes: usize,
    max_scratch_bytes: usize,
}

impl DecodeOptions {
    /// Build an explicit bytes/fields/work/depth/cells/labels/text policy.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        max_depth: u32,
        max_cells: usize,
        max_label_count: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_input_bytes,
            max_fields,
            max_work_bytes,
            max_depth,
            max_cells,
            max_label_count,
            max_text_bytes,
            max_output_bytes: max_input_bytes,
            max_allocations: max_input_bytes,
            max_retained_bytes: max_input_bytes,
            max_scratch_bytes: max_input_bytes,
        }
    }

    /// Build finite limits from the caller-owned source length.
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
        let cells = bytes.clamp(1, MAX_AUTOMATIC_LIMIT);
        let output = bytes
            .checked_mul(2)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .clamp(1, MAX_AUTOMATIC_LIMIT);
        let retained = bytes
            .checked_mul(4)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT)
            .max(bytes);
        let scratch = bytes
            .checked_mul(4)
            .unwrap_or(MAX_AUTOMATIC_LIMIT)
            .min(MAX_AUTOMATIC_LIMIT)
            .max(bytes);
        Self::new(bytes, fields, work, 16, cells, bytes, bytes)
            .with_max_output_bytes(output)
            .with_max_allocations(1)
            .with_max_retained_bytes(retained)
            .with_max_scratch_bytes(scratch)
    }

    /// Aggregate source-byte ceiling.
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    /// Aggregate wire-field ceiling.
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

    /// Maximum numeric cell count.
    #[must_use]
    pub const fn max_cells(self) -> usize {
        self.max_cells
    }

    /// Maximum row plus column label count.
    #[must_use]
    pub const fn max_label_count(self) -> usize {
        self.max_label_count
    }

    /// Maximum aggregate UTF-8 label bytes.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Candidate output-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Logical allocation ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_allocations(self) -> usize {
        self.max_allocations
    }

    /// Source-plus-candidate retained-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_retained_bytes(self) -> usize {
        self.max_retained_bytes
    }

    /// Candidate scratch-byte ceiling used by prepared rewrites.
    #[must_use]
    pub const fn max_scratch_bytes(self) -> usize {
        self.max_scratch_bytes
    }

    /// Replace the source-byte ceiling.
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

    /// Replace the wire-field ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the work ceiling.
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

    /// Replace the cell ceiling.
    #[must_use]
    pub const fn with_max_cells(mut self, maximum: usize) -> Self {
        self.max_cells = maximum;
        self
    }

    /// Replace the row/column label-count ceiling.
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

    /// Replace the candidate output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the logical allocation ceiling.
    #[must_use]
    pub const fn with_max_allocations(mut self, maximum: usize) -> Self {
        self.max_allocations = maximum;
        self
    }

    /// Replace the source-plus-candidate retained-byte ceiling.
    #[must_use]
    pub const fn with_max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = maximum;
        self
    }

    /// Replace the candidate scratch-byte ceiling.
    #[must_use]
    pub const fn with_max_scratch_bytes(mut self, maximum: usize) -> Self {
        self.max_scratch_bytes = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_input_bytes)
            // Unknown fields are structurally bounded by the strict pass and
            // must remain ignorable for future native GridValue fields.
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.max_depth.min(MAX_RECURSION_LIMIT))
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

    /// Borrow one validated label by ordinal.
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

/// A borrowed row sequence. Every row has exactly the number of values in
/// [`ChartDataSnapshot::column_labels`].
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct GridRows<'source> {
    source: &'source [u8],
    length: usize,
    columns: usize,
}

impl fmt::Debug for GridRows<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GridRows")
            .field("length", &self.length)
            .field("columns", &self.columns)
            .field("source_len", &self.source.len())
            .finish()
    }
}

impl<'source> GridRows<'source> {
    /// Number of validated rows.
    #[must_use]
    pub const fn len(self) -> usize {
        self.length
    }

    /// Whether no row is present.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.length == 0
    }

    /// Iterate rows in native wire order.
    #[must_use]
    pub fn iter(self) -> GridRowIter<'source> {
        GridRowIter {
            remaining: self.source,
            emitted: 0,
            length: self.length,
            columns: self.columns,
        }
    }
}

/// One borrowed row from a modern chart grid.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct GridRow<'source> {
    source: &'source [u8],
    length: usize,
}

impl fmt::Debug for GridRow<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GridRow")
            .field("value_count", &self.length)
            .field("source_len", &self.source.len())
            .finish()
    }
}

impl<'source> GridRow<'source> {
    /// Number of validated values in this row.
    #[must_use]
    pub const fn len(self) -> usize {
        self.length
    }

    /// Whether this row contains no numeric cells.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.length == 0
    }

    /// Iterate values in native wire order.
    #[must_use]
    pub fn values(self) -> GridValueIter<'source> {
        GridValueIter {
            remaining: self.source,
            emitted: 0,
            length: self.length,
        }
    }
}

/// Borrowed modern chart data with no generated repeated-field allocation.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ChartDataSnapshot<'source> {
    source: &'source [u8],
    grid_source: &'source [u8],
    row_labels: LabelList<'source>,
    column_labels: LabelList<'source>,
    rows: GridRows<'source>,
}

impl fmt::Debug for ChartDataSnapshot<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartDataSnapshot")
            .field("source_len", &self.source.len())
            .field("grid_source_len", &self.grid_source.len())
            .field("row_label_count", &self.row_labels.len())
            .field("column_label_count", &self.column_labels.len())
            .field("row_count", &self.rows.len())
            .field("column_count", &self.rows.columns)
            .finish()
    }
}

impl<'source> ChartDataSnapshot<'source> {
    /// Exact caller-owned input bytes.
    #[must_use]
    pub const fn source(self) -> &'source [u8] {
        self.source
    }

    /// Exact embedded `ChartGridArchive` bytes.
    #[must_use]
    pub const fn grid_source(self) -> &'source [u8] {
        self.grid_source
    }

    /// Borrowed row labels.
    #[must_use]
    pub const fn row_labels(self) -> LabelList<'source> {
        self.row_labels
    }

    /// Compatibility spelling matching the owned chart-data facade.
    #[must_use]
    pub const fn row_names(self) -> LabelList<'source> {
        self.row_labels()
    }

    /// Borrowed column labels.
    #[must_use]
    pub const fn column_labels(self) -> LabelList<'source> {
        self.column_labels
    }

    /// Compatibility spelling matching the owned chart-data facade.
    #[must_use]
    pub const fn column_names(self) -> LabelList<'source> {
        self.column_labels()
    }

    /// Borrowed rectangular row sequence.
    #[must_use]
    pub const fn rows(self) -> GridRows<'source> {
        self.rows
    }

    /// Compatibility spelling for callers that model rows as values.
    #[must_use]
    pub const fn values(self) -> GridRows<'source> {
        self.rows()
    }

    /// Number of rows in the rectangular grid.
    #[must_use]
    pub const fn row_count(self) -> usize {
        self.rows.length
    }

    /// Number of columns in the rectangular grid.
    #[must_use]
    pub const fn column_count(self) -> usize {
        self.rows.columns
    }
}

/// Iterator returned by [`GridRows::iter`].
#[derive(Clone, Copy)]
pub struct GridRowIter<'source> {
    remaining: &'source [u8],
    emitted: usize,
    length: usize,
    columns: usize,
}

impl fmt::Debug for GridRowIter<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GridRowIter")
            .field("emitted", &self.emitted)
            .field("length", &self.length)
            .field("columns", &self.columns)
            .field("remaining_len", &self.remaining.len())
            .finish()
    }
}

impl<'source> Iterator for GridRowIter<'source> {
    type Item = GridRow<'source>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.emitted < self.length && !self.remaining.is_empty() {
            let field = parse_unchecked_field(&mut self.remaining).ok()??;
            if field.number != GRID_ROW_FIELD {
                continue;
            }
            let WireValue::LengthDelimited(source) = field.value else {
                return None;
            };
            self.emitted += 1;
            return Some(GridRow {
                source,
                length: self.columns,
            });
        }
        None
    }
}

/// Iterator returned by [`GridRow::values`].
#[derive(Clone, Copy)]
pub struct GridValueIter<'source> {
    remaining: &'source [u8],
    emitted: usize,
    length: usize,
}

impl fmt::Debug for GridValueIter<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("GridValueIter")
            .field("emitted", &self.emitted)
            .field("length", &self.length)
            .field("remaining_len", &self.remaining.len())
            .finish()
    }
}

impl<'source> Iterator for GridValueIter<'source> {
    type Item = Option<f64>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.emitted < self.length && !self.remaining.is_empty() {
            let field = parse_unchecked_field(&mut self.remaining).ok()??;
            if field.number != GRID_VALUE_FIELD {
                continue;
            }
            let WireValue::LengthDelimited(source) = field.value else {
                return None;
            };
            self.emitted += 1;
            return Some(read_unchecked_numeric(source));
        }
        None
    }
}

/// Complete accounting for one chart-data decode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DecodeReport {
    source_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    cell_count: usize,
    label_count: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
}

impl DecodeReport {
    /// Root encoded bytes retained by the snapshot.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Strictly visited fields, including ignored fields and groups.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Aggregate strict-plus-Buffa work charged before return.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Maximum nested message depth observed.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Validated numeric cell count.
    #[must_use]
    pub const fn cell_count(self) -> usize {
        self.cell_count
    }

    /// Validated row/column label count.
    #[must_use]
    pub const fn label_count(self) -> usize {
        self.label_count
    }

    /// Aggregate UTF-8 label bytes.
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Generated-view allocations. Repeated fields are borrowed.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source bytes retained by the returned snapshot.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
}

/// Resource axis rejected by chart-data ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Source bytes exceed the configured ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Visited fields exceed the configured ceiling.
    Fields { observed: usize, maximum: usize },
    /// Work exceeds the configured ceiling.
    Work { observed: usize, maximum: usize },
    /// Nested message depth exceeds the configured ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Numeric cells exceed the configured ceiling.
    Cells { observed: usize, maximum: usize },
    /// Labels exceed the configured ceiling.
    Labels { observed: usize, maximum: usize },
    /// UTF-8 label bytes exceed the configured ceiling.
    Text { observed: usize, maximum: usize },
}

/// Failure from strict chart-data preflight or its private Buffa views.
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
    NonFiniteNumeric,
    NonRectangular {
        rows: usize,
        row_labels: usize,
        columns: usize,
        column_labels: usize,
    },
    Projection,
}

impl DecodeError {
    /// Complete accounting observed before failure.
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

    /// Return the duplicated singular field, if applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            _ => None,
        }
    }

    /// Whether a finite numeric cell was rejected.
    #[must_use]
    pub const fn is_non_finite_numeric(&self) -> bool {
        matches!(self.kind, DecodeErrorKind::NonFiniteNumeric)
    }

    /// Whether row/column dimensions disagree.
    #[must_use]
    pub const fn is_non_rectangular(&self) -> bool {
        matches!(self.kind, DecodeErrorKind::NonRectangular { .. })
    }

    /// Return byte-limit values, if applicable.
    #[must_use]
    pub const fn input_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return field-limit values, if applicable.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return work-limit values, if applicable.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return cell-limit values, if applicable.
    #[must_use]
    pub const fn cell_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Cells { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return label-limit values, if applicable.
    #[must_use]
    pub const fn label_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Labels { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return text-limit values, if applicable.
    #[must_use]
    pub const fn text_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Text { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
    }

    /// Return nesting-limit values, if applicable.
    #[must_use]
    pub const fn depth_limit_values(&self) -> Option<(u32, u32)> {
        match self.kind {
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => {
                Some((observed, maximum))
            },
            _ => None,
        }
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
                "chart data input bytes {observed} exceed maximum {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "chart data visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "chart data requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "chart data nesting {observed} exceeds maximum {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Cells { observed, maximum }) => write!(
                formatter,
                "chart data contains {observed} cells; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Labels { observed, maximum }) => write!(
                formatter,
                "chart data contains {observed} labels; maximum is {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Text { observed, maximum }) => write!(
                formatter,
                "chart data contains {observed} UTF-8 bytes; maximum is {maximum}"
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
            DecodeErrorKind::NonFiniteNumeric => {
                formatter.write_str("chart numeric values must be finite")
            },
            DecodeErrorKind::NonRectangular {
                rows,
                row_labels,
                columns,
                column_labels,
            } => write!(
                formatter,
                "chart data is not rectangular: {rows} rows/{columns} values versus {row_labels} row labels/{column_labels} column labels"
            ),
            DecodeErrorKind::Projection => formatter
                .write_str("chart data strict preflight disagrees with the Buffa projection"),
        }
    }
}

impl std::error::Error for DecodeError {}

/// Decode a complete modern chart drawable and return borrowed rectangular
/// chart data.
pub fn decode_modern<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<ChartDataSnapshot<'source>, DecodeError> {
    Ok(decode_modern_with_report(source, options)?.0)
}

/// Decode a complete modern chart drawable with aggregate accounting.
pub fn decode_modern_with_report<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<(ChartDataSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(*options, source.len());
    let result = preflight_modern(source, &mut budget)
        .and_then(|strict| project_modern(strict, *options, &mut budget));
    match result {
        Ok(snapshot) => Ok((snapshot, budget.report)),
        Err(kind) => Err(budget.finish_error(kind)),
    }
}

/// Decode one modern `TSCH.ChartGridArchive` payload directly.
pub fn decode_grid<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<ChartDataSnapshot<'source>, DecodeError> {
    Ok(decode_grid_with_report(source, options)?.0)
}

/// Decode one modern `TSCH.ChartGridArchive` payload with aggregate
/// accounting. This is useful to focused package readers that already own
/// the chart extension's selected grid span.
pub fn decode_grid_with_report<'source>(
    source: &'source [u8],
    options: &DecodeOptions,
) -> Result<(ChartDataSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(*options, source.len());
    let result = budget
        .root_message()
        .and_then(|()| preflight_grid(source, &mut budget, 1))
        .and_then(|strict| project_grid(strict, source, *options, &mut budget));
    match result {
        Ok(snapshot) => Ok((snapshot, budget.report)),
        Err(kind) => Err(budget.finish_error(kind)),
    }
}

#[derive(Debug, Clone, Copy)]
struct StrictModern<'source> {
    root_source: &'source [u8],
    chart_source: &'source [u8],
    grid: StrictGrid<'source>,
}

#[derive(Debug, Clone, Copy)]
struct StrictGrid<'source> {
    row_labels: LabelList<'source>,
    column_labels: LabelList<'source>,
    rows: GridRows<'source>,
}

fn preflight_modern<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<StrictModern<'source>, DecodeErrorKind> {
    budget.root_message()?;
    let mut extension = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, 1, budget)? {
        if field.number != MODERN_CHART_EXTENSION_FIELD {
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
    let mut chart_remaining = chart_source;
    let mut grid = None;
    while let Some(field) = next_field(&mut chart_remaining, 2, budget)? {
        if field.number != MODERN_GRID_FIELD {
            continue;
        }
        if grid.is_some() {
            return Err(DecodeErrorKind::DuplicateSingular("TSCH.ChartArchive.grid"));
        }
        let payload = field.length_delimited()?;
        budget.nested_message(payload.len(), 3)?;
        grid = Some(preflight_grid(payload, budget, 3)?);
    }
    let grid = grid.ok_or(DecodeErrorKind::MissingRequired("TSCH.ChartArchive.grid"))?;
    Ok(StrictModern {
        root_source: source,
        chart_source,
        grid,
    })
}

fn preflight_grid<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<StrictGrid<'source>, DecodeErrorKind> {
    budget.observe_depth(depth)?;
    let mut row_label_count = 0usize;
    let mut column_label_count = 0usize;
    let mut rows = 0usize;
    let mut row_value_count = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, depth, budget)? {
        match field.number {
            GRID_ROW_NAME_FIELD => {
                let bytes = field.length_delimited()?;
                budget.label(bytes.len(), "TSCH.ChartGridArchive.row_name")?;
                str::from_utf8(bytes).map_err(|_error| {
                    DecodeErrorKind::InvalidUtf8("TSCH.ChartGridArchive.row_name")
                })?;
                row_label_count = row_label_count
                    .checked_add(1)
                    .ok_or(DecodeErrorKind::NonCanonical("row label count overflow"))?;
            },
            GRID_COLUMN_NAME_FIELD => {
                let bytes = field.length_delimited()?;
                budget.label(bytes.len(), "TSCH.ChartGridArchive.column_name")?;
                str::from_utf8(bytes).map_err(|_error| {
                    DecodeErrorKind::InvalidUtf8("TSCH.ChartGridArchive.column_name")
                })?;
                column_label_count = column_label_count
                    .checked_add(1)
                    .ok_or(DecodeErrorKind::NonCanonical("column label count overflow"))?;
            },
            GRID_ROW_FIELD => {
                let payload = field.length_delimited()?;
                budget.nested_message(payload.len(), depth.saturating_add(1))?;
                let values = preflight_row(payload, budget, depth.saturating_add(1))?;
                if let Some(expected) = row_value_count {
                    if expected != values {
                        let observed_rows = rows
                            .checked_add(1)
                            .ok_or(DecodeErrorKind::NonCanonical("row count overflow"))?;
                        return Err(DecodeErrorKind::NonRectangular {
                            rows: observed_rows,
                            row_labels: row_label_count,
                            columns: values,
                            column_labels: column_label_count,
                        });
                    }
                } else {
                    row_value_count = Some(values);
                }
                rows = rows
                    .checked_add(1)
                    .ok_or(DecodeErrorKind::NonCanonical("row count overflow"))?;
            },
            GRID_ID_MAP_FIELD => {
                // The complete idMap subtree remains opaque. Its framing is
                // already validated as a length-delimited field; the native
                // id strings and indexes have no semantic role here.
                let _ = field.length_delimited()?;
            },
            _ => {},
        }
    }
    let columns = row_value_count.unwrap_or_default();
    if row_label_count == 0
        || column_label_count == 0
        || rows == 0
        || row_label_count != rows
        || column_label_count != columns
    {
        return Err(DecodeErrorKind::NonRectangular {
            rows,
            row_labels: row_label_count,
            columns,
            column_labels: column_label_count,
        });
    }
    let row_labels = LabelList::new(source, GRID_ROW_NAME_FIELD, row_label_count);
    let column_labels = LabelList::new(source, GRID_COLUMN_NAME_FIELD, column_label_count);
    let rows_view = GridRows {
        source,
        length: rows,
        columns,
    };
    Ok(StrictGrid {
        row_labels,
        column_labels,
        rows: rows_view,
    })
}

fn preflight_row(source: &[u8], budget: &mut Budget, depth: u32) -> Result<usize, DecodeErrorKind> {
    let mut values = 0usize;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, depth, budget)? {
        if field.number != GRID_VALUE_FIELD {
            continue;
        }
        let payload = field.length_delimited()?;
        budget.cell()?;
        budget.nested_message(payload.len(), depth.saturating_add(1))?;
        preflight_value(payload, budget, depth.saturating_add(1))?;
        values = values
            .checked_add(1)
            .ok_or(DecodeErrorKind::NonCanonical("value count overflow"))?;
    }
    Ok(values)
}

fn preflight_value(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), DecodeErrorKind> {
    budget.observe_depth(depth)?;
    let mut date_1_0 = false;
    let mut duration = false;
    let mut date = false;
    let mut numeric = false;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, depth, budget)? {
        match field.number {
            GRID_VALUE_FIELD => {
                if numeric {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.GridValue.numeric_value",
                    ));
                }
                let raw = field.fixed64()?;
                let value = fixed64_to_f64(raw)?;
                if !value.is_finite() {
                    return Err(DecodeErrorKind::NonFiniteNumeric);
                }
                numeric = true;
            },
            GRID_VALUE_DATE_1_0_FIELD => {
                if date_1_0 {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.GridValue.date_value_1_0",
                    ));
                }
                let _ = fixed64_to_f64(field.fixed64()?)?;
                date_1_0 = true;
            },
            GRID_VALUE_DURATION_FIELD => {
                if duration {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.GridValue.duration_value",
                    ));
                }
                let _ = fixed64_to_f64(field.fixed64()?)?;
                duration = true;
            },
            GRID_VALUE_DATE_FIELD => {
                if date {
                    return Err(DecodeErrorKind::DuplicateSingular(
                        "TSCH.GridValue.date_value",
                    ));
                }
                let _ = fixed64_to_f64(field.fixed64()?)?;
                date = true;
            },
            _ => {},
        }
    }
    Ok(())
}

fn project_modern<'source>(
    strict: StrictModern<'source>,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartDataSnapshot<'source>, DecodeErrorKind> {
    budget.buffa(strict.chart_source.len())?;
    let view: projection::ChartArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(strict.chart_source)
        .map_err(DecodeErrorKind::Wire)?;
    let projected_grid = view.grid;
    if !same_borrowed_bytes(projected_grid, Some(strict.grid.rows.source)) {
        return Err(DecodeErrorKind::Projection);
    }
    project_grid_values(strict.grid.rows.source, options, budget)?;
    Ok(ChartDataSnapshot {
        source: strict.root_source,
        grid_source: strict.grid.rows.source,
        row_labels: strict.grid.row_labels,
        column_labels: strict.grid.column_labels,
        rows: strict.grid.rows,
    })
}

fn project_grid<'source>(
    strict: StrictGrid<'source>,
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ChartDataSnapshot<'source>, DecodeErrorKind> {
    project_grid_values(strict.rows.source, options, budget)?;
    Ok(ChartDataSnapshot {
        source,
        grid_source: source,
        row_labels: strict.row_labels,
        column_labels: strict.column_labels,
        rows: strict.rows,
    })
}

fn project_grid_values(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(), DecodeErrorKind> {
    // This source walk is separate from strict preflight and is charged in
    // full before publishing a borrowed snapshot. Value payloads are charged
    // again below for the raw parity lookup alongside their Buffa view.
    budget.work(source.len())?;
    let mut remaining = source;
    while let Some(field) = parse_unchecked_field(&mut remaining)? {
        if field.number != GRID_ROW_FIELD {
            continue;
        }
        let WireValue::LengthDelimited(row_source) = field.value else {
            return Err(DecodeErrorKind::Projection);
        };
        let mut row_remaining = row_source;
        while let Some(value_field) = parse_unchecked_field(&mut row_remaining)? {
            if value_field.number != GRID_VALUE_FIELD {
                continue;
            }
            let WireValue::LengthDelimited(value_source) = value_field.value else {
                return Err(DecodeErrorKind::Projection);
            };
            budget.buffa(value_source.len())?;
            let view: projection::GridValueLazyView<'_> = options
                .buffa()
                .decode_lazy_view(value_source)
                .map_err(DecodeErrorKind::Wire)?;
            budget.work(value_source.len())?;
            if !same_float(view.numeric_value, read_unchecked_numeric(value_source)) {
                return Err(DecodeErrorKind::Projection);
            }
        }
    }
    Ok(())
}

fn same_float(left: Option<f64>, right: Option<f64>) -> bool {
    left.map(f64::to_bits) == right.map(f64::to_bits)
}

fn same_borrowed_bytes(left: Option<&[u8]>, right: Option<&[u8]>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.as_ptr() == right.as_ptr() && left.len() == right.len(),
        (None, None) => true,
        _ => false,
    }
}

fn fixed64_to_f64(bytes: &[u8]) -> Result<f64, DecodeErrorKind> {
    let bytes: [u8; 8] = bytes
        .try_into()
        .map_err(|_error| DecodeErrorKind::Projection)?;
    Ok(f64::from_le_bytes(bytes))
}

fn read_unchecked_numeric(source: &[u8]) -> Option<f64> {
    let mut remaining = source;
    while let Some(field) = parse_unchecked_field(&mut remaining).ok()? {
        if field.number != GRID_VALUE_FIELD {
            continue;
        }
        return match field.value {
            WireValue::Fixed64(bytes) => fixed64_to_f64(bytes).ok(),
            _ => None,
        };
    }
    None
}

#[derive(Debug)]
struct Budget {
    options: DecodeOptions,
    report: DecodeReport,
}

impl Budget {
    const fn new(options: DecodeOptions, source_bytes: usize) -> Self {
        Self {
            options,
            report: DecodeReport {
                source_bytes,
                fields: 0,
                work_bytes: 0,
                max_depth: 0,
                cell_count: 0,
                label_count: 0,
                text_bytes: 0,
                allocations: 0,
                retained_bytes: source_bytes,
            },
        }
    }

    fn root_message(&mut self) -> Result<(), DecodeErrorKind> {
        if self.report.source_bytes > self.options.max_input_bytes {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Bytes {
                observed: self.report.source_bytes,
                maximum: self.options.max_input_bytes,
            }));
        }
        self.nested_message(self.report.source_bytes, 1)
    }

    fn nested_message(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeErrorKind> {
        self.observe_depth(depth)?;
        let amount = bytes.checked_mul(2).unwrap_or(usize::MAX);
        self.work(amount)
    }

    fn buffa(&mut self, bytes: usize) -> Result<(), DecodeErrorKind> {
        self.work(bytes)
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeErrorKind> {
        self.report.max_depth = self.report.max_depth.max(depth);
        if depth > self.options.max_depth || depth > MAX_RECURSION_LIMIT {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.max_depth.min(MAX_RECURSION_LIMIT),
            }));
        }
        Ok(())
    }

    fn field(&mut self) -> Result<(), DecodeErrorKind> {
        let observed = self.report.fields.checked_add(1).unwrap_or(usize::MAX);
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
        let observed = self
            .report
            .work_bytes
            .checked_add(bytes)
            .unwrap_or(usize::MAX);
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
        let count = self.report.label_count.checked_add(1).unwrap_or(usize::MAX);
        self.report.label_count = count;
        if count > self.options.max_label_count {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Labels {
                observed: count,
                maximum: self.options.max_label_count,
            }));
        }
        let text = self
            .report
            .text_bytes
            .checked_add(bytes)
            .unwrap_or(usize::MAX);
        self.report.text_bytes = text;
        if text > self.options.max_text_bytes {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Text {
                observed: text,
                maximum: self.options.max_text_bytes,
            }));
        }
        Ok(())
    }

    fn cell(&mut self) -> Result<(), DecodeErrorKind> {
        let count = self.report.cell_count.checked_add(1).unwrap_or(usize::MAX);
        self.report.cell_count = count;
        if count > self.options.max_cells {
            return Err(DecodeErrorKind::Limit(DecodeLimit::Cells {
                observed: count,
                maximum: self.options.max_cells,
            }));
        }
        Ok(())
    }

    fn finish_error(self, kind: DecodeErrorKind) -> DecodeError {
        DecodeError::new(kind, self.report)
    }
}

#[derive(Debug, Clone, Copy)]
enum WireValue<'source> {
    Varint,
    Fixed64(&'source [u8]),
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
    fn fixed64(self) -> Result<&'source [u8], DecodeErrorKind> {
        if self.wire_type != buffa::encoding::WireType::Fixed64 {
            return Err(DecodeErrorKind::Wire(
                buffa::DecodeError::WireTypeMismatch {
                    field_number: self.number,
                    expected: buffa::encoding::WireType::Fixed64 as u8,
                    actual: self.wire_type as u8,
                },
            ));
        }
        let WireValue::Fixed64(value) = self.value else {
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
            let (_value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeErrorKind::NonCanonical("protobuf varint value"));
            }
            WireValue::Varint
        },
        buffa::encoding::WireType::Fixed64 => WireValue::Fixed64(take_exact(source, 8)?),
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
            let _ = take_exact(source, 4)?;
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
    if number == 0 || number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(DecodeErrorKind::Wire(
            buffa::DecodeError::InvalidFieldNumber,
        ));
    }
    let wire_type =
        buffa::encoding::WireType::from_u32(raw_tag & 7).map_err(DecodeErrorKind::Wire)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let _ = take_varint(source)?;
            WireValue::Varint
        },
        buffa::encoding::WireType::Fixed64 => WireValue::Fixed64(take_exact(source, 8)?),
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
            let _ = take_exact(source, 4)?;
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
                let _ = take_varint(source)?;
            },
            buffa::encoding::WireType::Fixed64 => {
                let _ = take_exact(source, 8)?;
            },
            buffa::encoding::WireType::LengthDelimited => {
                let length = usize::try_from(take_varint(source)?.0)
                    .map_err(|_error| DecodeErrorKind::Wire(buffa::DecodeError::MessageTooLarge))?;
                let _ = take_exact(source, length)?;
            },
            buffa::encoding::WireType::StartGroup => skip_unchecked_group(source, number)?,
            buffa::encoding::WireType::Fixed32 => {
                let _ = take_exact(source, 4)?;
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

mod rewrite;

pub use rewrite::{
    PreparedChartDataRewrite, RewriteError, RewriteExecutionLimits, RewriteExecutionRequirements,
    RewriteLimit, RewriteOutput, RewriteReport, prepare_chart_data_rewrite,
};

#[cfg(test)]
mod tests;

#[cfg(test)]
mod rewrite_tests;
