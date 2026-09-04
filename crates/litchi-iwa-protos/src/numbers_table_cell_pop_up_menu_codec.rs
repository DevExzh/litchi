//! Strict generated-free projection for Numbers control-cell popup menus.
//!
//! The native `PopUpMenuModel` contains a deprecated repeated field and a
//! repeated `TSCE.CellValueArchive` field.  This module deliberately parses
//! those repeated records by borrowing the caller's bytes; the private Buffa
//! sidecar contains only singular parity envelopes and cannot materialize
//! input-width collections.  Unknown fields/groups are retained by raw source
//! spans, while every known field is checked for canonical wire shape before
//! a semantic snapshot is published.
//! Control-list refcount validation remains in
//! [`crate::numbers_table_cell_storage_codec`]; this codec owns the selected
//! `PopUpMenuModel` and `CellSpecArchive` payloads.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Strict wire helpers stay beside the snapshots they construct."
)]

use core::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_table_cell_currency_format_generated::LitchiIwaNumbersTableCellCurrencyFormatProjection as currency_projection;
use crate::buffa_numbers_table_cell_pop_up_menu_generated::LitchiIwaNumbersTableCellPopUpMenuProjection as projection;
use crate::buffa_numbers_table_cell_text_format_generated::LitchiIwaNumbersTableCellTextFormatProjection as text_projection;

const POPUP_ITEM_FIELD: u32 = 2;
const POPUP_DEPRECATED_ITEM_FIELD: u32 = 1;
const CELL_VALUE_TYPE_FIELD: u32 = 1;
const CELL_VALUE_STRING_FIELD: u32 = 5;
const STRING_VALUE_FIELD: u32 = 1;
const STRING_FORMAT_FIELD: u32 = 2;
const STRING_IMPLICIT_FIELD: u32 = 3;
const STRING_EXPLICIT_FIELD: u32 = 4;
const STRING_REGEX_FIELD: u32 = 5;
const STRING_CASE_SENSITIVE_REGEX_FIELD: u32 = 6;
const FORMAT_TYPE_FIELD: u32 = 1;
const CELL_SPEC_INTERACTION_FIELD: u32 = 1;
const CELL_SPEC_MODEL_FIELD: u32 = 6;
const CELL_SPEC_FIRST_FIELD: u32 = 7;
const CELL_SPEC_FORMULA_FIELD: u32 = 2;
const CELL_SPEC_DEPRECATED_LABEL_FIELD: u32 = 8;
const FORMAT_DECIMAL_PLACES_FIELD: u32 = 2;
const FORMAT_CURRENCY_CODE_FIELD: u32 = 3;
const FORMAT_NEGATIVE_STYLE_FIELD: u32 = 4;
const FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD: u32 = 5;
const FORMAT_USE_ACCOUNTING_STYLE_FIELD: u32 = 6;
const FORMAT_DURATION_STYLE_FIELD: u32 = 7;
const FORMAT_BASE_FIELD: u32 = 8;
const FORMAT_BASE_PLACES_FIELD: u32 = 9;
const FORMAT_BASE_USE_MINUS_SIGN_FIELD: u32 = 10;
const FORMAT_FRACTION_ACCURACY_FIELD: u32 = 11;
const FORMAT_SUPPRESS_DATE_FORMAT_FIELD: u32 = 12;
const FORMAT_SUPPRESS_TIME_FORMAT_FIELD: u32 = 13;
const FORMAT_DATE_TIME_FORMAT_FIELD: u32 = 14;
const FORMAT_DURATION_UNIT_LARGEST_FIELD: u32 = 15;
const FORMAT_DURATION_UNIT_SMALLEST_FIELD: u32 = 16;
const FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD: u32 = 20;
const FORMAT_CONTROL_MINIMUM_FIELD: u32 = 21;
const FORMAT_CONTROL_MAXIMUM_FIELD: u32 = 22;
const FORMAT_CONTROL_INCREMENT_FIELD: u32 = 23;
const FORMAT_CONTROL_FORMAT_TYPE_FIELD: u32 = 24;
const FORMAT_SLIDER_ORIENTATION_FIELD: u32 = 25;
const FORMAT_SLIDER_POSITION_FIELD: u32 = 26;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_TYPE_FIELD: u32 = 2;
const REFERENCE_EXTERNAL_FIELD: u32 = 3;
const POPUP_INTERACTION_TYPE: u32 = 7;
/// Native interaction type for a stepper control.
pub const STEPPER_INTERACTION_TYPE: u32 = 4;
/// Native interaction type for a slider control.
pub const SLIDER_INTERACTION_TYPE: u32 = 5;
/// Native interaction type for a star-rating control.
pub const STAR_RATING_INTERACTION_TYPE: u32 = 6;
/// Native interaction type for a checkbox control.
pub const CHECKBOX_INTERACTION_TYPE: u32 = 8;
const FORMAT_TYPE_TEXT: u64 = 260;
const STRING_TYPE: u64 = 5;
const NIL_TYPE: u64 = 1;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MAX_RECURSION: u32 = 64;

/// Finite aggregate limits for one popup-model or cell-spec operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_output_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    max_references: usize,
    max_items: usize,
    max_text_bytes: usize,
}

impl DecodeOptions {
    /// Construct all finite decode and rewrite ceilings explicitly.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_output_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
        max_references: usize,
        max_items: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_message_bytes,
            max_output_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
            max_references,
            max_items,
            max_text_bytes,
        }
    }

    /// Conservative finite limits derived from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.checked_mul(2).unwrap_or(usize::MAX),
            bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
            bytes.checked_mul(16).unwrap_or(usize::MAX).max(1),
            MAX_RECURSION,
            bytes.max(1),
            bytes.max(1),
            bytes,
        )
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the item-count ceiling.
    #[must_use]
    pub const fn with_max_items(mut self, maximum: usize) -> Self {
        self.max_items = maximum;
        self
    }

    /// Replace the selected text-byte ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }
}

/// Typed finite failure for strict popup decoding or prepared execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input payload exceeded the configured source ceiling.
    InputBytes { observed: usize, maximum: usize },
    /// Candidate payload exceeded the configured output ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Known and unknown field records exceeded the aggregate ceiling.
    Fields { observed: usize, maximum: usize },
    /// Aggregate inspected bytes exceeded the work ceiling.
    Work { observed: usize, maximum: usize },
    /// Nested message/group depth exceeded the ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Reference envelopes exceeded the aggregate ceiling.
    References { observed: usize, maximum: usize },
    /// Repeated popup values exceeded the aggregate ceiling.
    Items { observed: usize, maximum: usize },
    /// Selected UTF-8 text exceeded the aggregate ceiling.
    Text { observed: usize, maximum: usize },
    /// A fallible output or scratch reservation was refused.
    Allocation { requested: usize },
    /// Retained candidate bytes exceeded the execution ceiling.
    Retained { observed: usize, maximum: usize },
    /// Temporary scratch bytes exceeded the execution ceiling.
    Scratch { observed: usize, maximum: usize },
}

/// Strict popup codec error.  Native identifiers are intentionally omitted
/// from diagnostics to avoid leaking object-routing data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    pub(crate) const fn invalid() -> Self {
        Self { limit: None }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }

    /// Return a typed resource failure, when this error is a limit failure.
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }

    /// Return the refused allocation size, if applicable.
    #[must_use]
    pub const fn allocation_requested(self) -> Option<usize> {
        match self.limit {
            Some(DecodeLimit::Allocation { requested }) => Some(requested),
            _ => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid Numbers table-cell popup-menu payload")
    }
}

impl std::error::Error for DecodeError {}

/// Exact aggregate consumption for one strict operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
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
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn items(self) -> usize {
        self.items
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
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

/// A strict canonical `TSP.Reference` projection.
#[derive(Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl fmt::Debug for ReferenceSnapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ReferenceSnapshot")
            .field("identifier", &"<redacted>")
            .field("deprecated_type", &self.deprecated_type)
            .field("deprecated_is_external", &self.deprecated_is_external)
            .finish()
    }
}

impl ReferenceSnapshot {
    #[must_use]
    pub const fn identifier(self) -> u64 {
        self.identifier
    }
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// One borrowed popup string item.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuItem<'source> {
    value: &'source str,
}

impl<'source> PopUpMenuItem<'source> {
    #[must_use]
    pub const fn value(self) -> &'source str {
        self.value
    }
}

/// Borrowed semantic facts for a strict `PopUpMenuModel` payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PopUpMenuModelSnapshot<'source> {
    source: &'source [u8],
    first_nil: bool,
    item_count: usize,
    text_bytes: usize,
}

impl<'source> PopUpMenuModelSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn has_nil_sentinel(self) -> bool {
        self.first_nil
    }
    #[must_use]
    pub const fn item_count(self) -> usize {
        self.item_count
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Iterate values without allocating an input-width collection.
    pub fn items(self) -> PopUpMenuItems<'source> {
        PopUpMenuItems {
            source: self.source,
            offset: 0,
            index: 0,
        }
    }
}

/// Borrowed iterator over popup string values.
#[derive(Debug, Clone, Copy)]
pub struct PopUpMenuItems<'source> {
    source: &'source [u8],
    offset: usize,
    index: usize,
}

impl<'source> Iterator for PopUpMenuItems<'source> {
    type Item = PopUpMenuItem<'source>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.offset < self.source.len() {
            let field = parse_one_field(self.source, self.offset, 0).ok()?;
            self.offset = field.end;
            if field.number != POPUP_ITEM_FIELD {
                continue;
            }
            let payload = field.payload?;
            let value = match parse_cell_value_for_iterator(payload) {
                Ok(Some(value)) => value,
                Ok(None) => {
                    self.index = 1;
                    continue;
                },
                Err(_error) => return None,
            };
            self.index = self.index.saturating_add(1);
            return Some(PopUpMenuItem { value });
        }
        None
    }
}

/// Borrowed facts for one popup `CellSpecArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellSpecSnapshot<'source> {
    source: &'source [u8],
    interaction_type: u32,
    popup_model: ReferenceSnapshot,
    starts_with_first: bool,
}

/// Borrowed facts for one non-popup interactive `CellSpecArchive`.
///
/// The range fields are present only for slider and stepper controls.  The
/// source bytes remain authoritative, so unknown fields and groups are never
/// normalized by this projection.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ControlCellSpecSnapshot<'source> {
    source: &'source [u8],
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
}

impl<'source> ControlCellSpecSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    #[must_use]
    pub const fn interaction_type(self) -> u32 {
        self.interaction_type
    }

    #[must_use]
    pub const fn range_control_min(self) -> Option<f64> {
        self.range_control_min
    }

    #[must_use]
    pub const fn range_control_max(self) -> Option<f64> {
        self.range_control_max
    }

    #[must_use]
    pub const fn range_control_inc(self) -> Option<f64> {
        self.range_control_inc
    }

    /// Whether this snapshot describes a slider or stepper range control.
    #[must_use]
    pub const fn is_range_control(self) -> bool {
        matches!(
            self.interaction_type,
            SLIDER_INTERACTION_TYPE | STEPPER_INTERACTION_TYPE
        )
    }
}

/// Borrowed facts for a strict control-cell `FormatStructArchive` payload.
///
/// The control codec intentionally exposes only the fields that identify and
/// parameterize a native control.  Other fields are retained as raw source
/// bytes and are not silently re-encoded.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ControlFormatSnapshot<'source> {
    source: &'source [u8],
    format_type: u32,
    decimal_places: Option<u32>,
    currency_code: Option<&'source str>,
    negative_style: Option<u32>,
    show_thousands_separator: Option<bool>,
    use_accounting_style: Option<bool>,
    duration_style: Option<u32>,
    base: Option<u32>,
    base_places: Option<u32>,
    base_use_minus_sign: Option<bool>,
    fraction_accuracy: Option<u32>,
    suppress_date_format: Option<bool>,
    suppress_time_format: Option<bool>,
    date_time_format: Option<&'source str>,
    duration_unit_largest: Option<u32>,
    duration_unit_smallest: Option<u32>,
    control_minimum: Option<f64>,
    control_maximum: Option<f64>,
    control_increment: Option<f64>,
    control_format_type: Option<u32>,
    slider_orientation: Option<u32>,
    slider_position: Option<u32>,
}

impl<'source> ControlFormatSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }
    #[must_use]
    pub const fn decimal_places(self) -> Option<u32> {
        self.decimal_places
    }
    #[must_use]
    pub const fn currency_code(self) -> Option<&'source str> {
        self.currency_code
    }
    #[must_use]
    pub const fn negative_style(self) -> Option<u32> {
        self.negative_style
    }
    #[must_use]
    pub const fn show_thousands_separator(self) -> Option<bool> {
        self.show_thousands_separator
    }
    #[must_use]
    pub const fn use_accounting_style(self) -> Option<bool> {
        self.use_accounting_style
    }
    #[must_use]
    pub const fn duration_style(self) -> Option<u32> {
        self.duration_style
    }
    #[must_use]
    pub const fn base(self) -> Option<u32> {
        self.base
    }
    #[must_use]
    pub const fn base_places(self) -> Option<u32> {
        self.base_places
    }
    #[must_use]
    pub const fn base_use_minus_sign(self) -> Option<bool> {
        self.base_use_minus_sign
    }
    #[must_use]
    pub const fn fraction_accuracy(self) -> Option<u32> {
        self.fraction_accuracy
    }
    #[must_use]
    pub const fn suppress_date_format(self) -> Option<bool> {
        self.suppress_date_format
    }
    #[must_use]
    pub const fn suppress_time_format(self) -> Option<bool> {
        self.suppress_time_format
    }
    #[must_use]
    pub const fn date_time_format(self) -> Option<&'source str> {
        self.date_time_format
    }
    #[must_use]
    pub const fn duration_unit_largest(self) -> Option<u32> {
        self.duration_unit_largest
    }
    #[must_use]
    pub const fn duration_unit_smallest(self) -> Option<u32> {
        self.duration_unit_smallest
    }
    #[must_use]
    pub const fn control_minimum(self) -> Option<f64> {
        self.control_minimum
    }
    #[must_use]
    pub const fn control_maximum(self) -> Option<f64> {
        self.control_maximum
    }
    #[must_use]
    pub const fn control_increment(self) -> Option<f64> {
        self.control_increment
    }
    #[must_use]
    pub const fn control_format_type(self) -> Option<u32> {
        self.control_format_type
    }
    #[must_use]
    pub const fn slider_orientation(self) -> Option<u32> {
        self.slider_orientation
    }
    #[must_use]
    pub const fn slider_position(self) -> Option<u32> {
        self.slider_position
    }
}

impl<'source> CellSpecSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }
    #[must_use]
    pub const fn interaction_type(self) -> u32 {
        self.interaction_type
    }
    #[must_use]
    pub const fn popup_model(self) -> ReferenceSnapshot {
        self.popup_model
    }
    #[must_use]
    pub const fn starts_with_first(self) -> bool {
        self.starts_with_first
    }
}

/// Visitor for streaming popup item values without allocating a vector.
pub trait PopUpMenuVisitor {
    /// Observe one source-borrowed string item.
    fn visit_item(&mut self, item: PopUpMenuItem<'_>) -> Result<(), DecodeError>;
}

/// Decode a popup model without exposing generated Buffa values.
pub fn decode_popup_menu_model(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PopUpMenuModelSnapshot<'_>, DecodeError> {
    Ok(decode_popup_menu_model_with_report(source, options)?.0)
}

/// Decode a popup model and return its aggregate strict resource report.
pub fn decode_popup_menu_model_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(PopUpMenuModelSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = parse_popup_model_with_visitor(source, &mut budget, None)?;
    Ok((snapshot, budget.finish(0)))
}

/// Decode a popup model and stream each item to a caller-owned visitor.
pub fn decode_popup_menu_model_with_visitor<V: PopUpMenuVisitor>(
    source: &[u8],
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let _snapshot = parse_popup_model_with_visitor(source, &mut budget, Some(visitor))?;
    Ok(budget.finish(0))
}

/// Decode a popup cell-spec archive.
pub fn decode_cell_spec(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CellSpecSnapshot<'_>, DecodeError> {
    Ok(decode_cell_spec_with_report(source, options)?.0)
}

/// Decode a popup cell-spec archive with aggregate resource accounting.
pub fn decode_cell_spec_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CellSpecSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = parse_cell_spec(source, &mut budget)?;
    if !budget.unknown_fields {
        buffa_cell_spec_parity(source, &mut budget)?;
    }
    Ok((snapshot, budget.finish(0)))
}

/// Decode a strict checkbox, star-rating, slider, or stepper cell spec.
pub fn decode_control_cell_spec(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ControlCellSpecSnapshot<'_>, DecodeError> {
    Ok(decode_control_cell_spec_with_report(source, options)?.0)
}

/// Decode a strict control cell spec and return its aggregate resource report.
pub fn decode_control_cell_spec_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ControlCellSpecSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = parse_control_cell_spec(source, &mut budget)?;
    if !budget.unknown_fields {
        buffa_cell_spec_parity(source, &mut budget)?;
    }
    Ok((snapshot, budget.finish(0)))
}

/// Decode a strict `FormatStructArchive` used by an interactive cell.
pub fn decode_control_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ControlFormatSnapshot<'_>, DecodeError> {
    Ok(decode_control_format_with_report(source, options)?.0)
}

/// Decode a strict control format and return its aggregate resource report.
pub fn decode_control_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(ControlFormatSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = parse_control_format(source, &mut budget)?;
    Ok((snapshot, budget.finish(0)))
}

/// Native Numbers display-format discriminator for a plain number cell.
///
/// `FormatStructArchive` is shared by many iWork features.  This narrow
/// route intentionally accepts only the four fields that Numbers writes for
/// its plain number format; currency, date, custom, and control formats stay
/// on their respective package-owned paths.
pub const NATIVE_NUMBER_FORMAT_TYPE: u32 = 256;

/// Native Numbers display-format discriminator for a currency cell.
pub const NATIVE_CURRENCY_FORMAT_TYPE: u32 = 257;

/// Native Numbers display-format discriminator for a plain percentage cell.
pub const NATIVE_PERCENTAGE_FORMAT_TYPE: u32 = 258;

/// Native Numbers display-format discriminator for a scientific cell.
pub const NATIVE_SCIENTIFIC_FORMAT_TYPE: u32 = 259;

/// Native Numbers display-format discriminator for a Text cell.
pub const NATIVE_TEXT_FORMAT_TYPE: u32 = 260;

/// Native Numbers display-format discriminator for a fraction cell.
pub const NATIVE_FRACTION_FORMAT_TYPE: u32 = 262;

/// Native fraction accuracy for a denominator with at most one digit.
pub const NATIVE_FRACTION_UP_TO_ONE_DIGIT: u32 = u32::MAX;

/// Native fraction accuracy for a denominator with at most two digits.
pub const NATIVE_FRACTION_UP_TO_TWO_DIGITS: u32 = u32::MAX - 1;

/// Native fraction accuracy for a denominator with at most three digits.
pub const NATIVE_FRACTION_UP_TO_THREE_DIGITS: u32 = u32::MAX - 2;

/// Native fraction accuracy for halves.
pub const NATIVE_FRACTION_HALVES: u32 = 2;

/// Native fraction accuracy for quarters.
pub const NATIVE_FRACTION_QUARTERS: u32 = 4;

/// Native fraction accuracy for eighths.
pub const NATIVE_FRACTION_EIGHTHS: u32 = 8;

/// Native fraction accuracy for sixteenths.
pub const NATIVE_FRACTION_SIXTEENTHS: u32 = 16;

/// Native fraction accuracy for tenths.
pub const NATIVE_FRACTION_TENTHS: u32 = 10;

/// Native fraction accuracy for hundredths.
pub const NATIVE_FRACTION_HUNDREDTHS: u32 = 100;

/// Native negative-number style used by scientific formats.
pub const NATIVE_SCIENTIFIC_NEGATIVE_STYLE: u32 = 0;

/// Scientific formats never display a thousands separator.
pub const NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR: bool = false;

/// Native discriminator used for automatic decimal places.
pub const NATIVE_AUTOMATIC_DECIMAL_PLACES: u32 = 253;

/// Largest fixed decimal-place value accepted by Numbers' native format.
pub const MAX_NUMBER_DECIMAL_PLACES: u32 = 30;

const NUMBER_FORMAT_TYPE_FIELD: u32 = 1;
const NUMBER_FORMAT_DECIMAL_PLACES_FIELD: u32 = 2;
const NUMBER_FORMAT_NEGATIVE_STYLE_FIELD: u32 = 4;
const NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD: u32 = 5;
const NUMBER_FORMAT_MAX_KNOWN_FIELD: u32 = 45;

/// Borrowed semantic facts for a strict plain Number or Percentage
/// `FormatStructArchive`.
///
/// The complete source payload remains available through [`Self::raw`].
/// Unknown extension fields and groups are never decoded into owned storage
/// and are copied byte-for-byte by [`rewrite_number_format`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumberFormatSnapshot<'source> {
    source: &'source [u8],
    format_type: u32,
    decimal_places: u32,
    negative_style: u32,
    show_thousands_separator: bool,
}

impl<'source> NumberFormatSnapshot<'source> {
    /// Borrow the original wire payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    /// Return the native format discriminator (`256` for Number, `258` for
    /// Percentage).
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }

    /// Return `253` for automatic places, otherwise a fixed value in `0..=30`.
    #[must_use]
    pub const fn decimal_places(self) -> u32 {
        self.decimal_places
    }

    /// Return the native negative-number style (`0..=3`).
    #[must_use]
    pub const fn negative_style(self) -> u32 {
        self.negative_style
    }

    /// Whether the thousands separator is displayed.
    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.show_thousands_separator
    }
}

/// Owned scalar values accepted by the plain Number or Percentage format
/// writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumberFormatWrite {
    decimal_places: u32,
    negative_style: u32,
    show_thousands_separator: bool,
}

impl NumberFormatWrite {
    /// Construct a native plain-number format update.
    #[must_use]
    pub const fn new(
        decimal_places: u32,
        negative_style: u32,
        show_thousands_separator: bool,
    ) -> Self {
        Self {
            decimal_places,
            negative_style,
            show_thousands_separator,
        }
    }

    /// Copy the four semantic values from a decoded snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: NumberFormatSnapshot<'_>) -> Self {
        Self::new(
            snapshot.decimal_places,
            snapshot.negative_style,
            snapshot.show_thousands_separator,
        )
    }

    /// Return the native decimal-place discriminator/value.
    #[must_use]
    pub const fn decimal_places(self) -> u32 {
        self.decimal_places
    }

    /// Return the native negative-number style.
    #[must_use]
    pub const fn negative_style(self) -> u32 {
        self.negative_style
    }

    /// Return the thousands-separator setting.
    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.show_thousands_separator
    }

    /// Replace the decimal-place discriminator/value.
    #[must_use]
    pub const fn with_decimal_places(mut self, value: u32) -> Self {
        self.decimal_places = value;
        self
    }

    /// Replace the native negative-number style.
    #[must_use]
    pub const fn with_negative_style(mut self, value: u32) -> Self {
        self.negative_style = value;
        self
    }

    /// Replace the thousands-separator setting.
    #[must_use]
    pub const fn with_show_thousands_separator(mut self, value: bool) -> Self {
        self.show_thousands_separator = value;
        self
    }
}

/// Borrowed semantic facts for one strict native Text `FormatStructArchive`.
///
/// Text has no display-format scalar payload beyond its native discriminator.
/// The complete source remains authoritative so unknown extension fields and
/// groups can be retained byte-for-byte by a prepared rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TextFormatSnapshot<'source> {
    source: &'source [u8],
    format_type: u32,
}

impl<'source> TextFormatSnapshot<'source> {
    pub(crate) const fn raw(self) -> &'source [u8] {
        self.source
    }

    pub(crate) const fn format_type(self) -> u32 {
        self.format_type
    }
}

/// Scalar values accepted by the strict native Text writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TextFormatWrite;

impl TextFormatWrite {
    pub(crate) const fn new() -> Self {
        Self
    }

    pub(crate) const fn from_snapshot(_snapshot: TextFormatSnapshot<'_>) -> Self {
        Self
    }

    pub(crate) const fn format_type(self) -> u32 {
        NATIVE_TEXT_FORMAT_TYPE
    }
}

/// Prepared source-preserving native Text rewrite.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PreparedTextFormatRewrite<'source> {
    source: &'source [u8],
    layout: TextFormatLayout,
    write: TextFormatWrite,
    requirements: RewriteExecutionRequirements,
    candidate_work: usize,
    verify_options: DecodeOptions,
}

impl PreparedTextFormatRewrite<'_> {
    pub(crate) const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub(crate) fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    pub(crate) fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_text_format_rewrite(&mut bytes, self.source, self.layout)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        verify_text_format_candidate(
            &bytes,
            self.write,
            self.verify_options,
            self.layout,
            self.candidate_work,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical native Text append.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PreparedTextFormatWrite {
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedTextFormatWrite {
    pub(crate) const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub(crate) fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    pub(crate) fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_text_format_canonical(&mut bytes)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        verify_text_format_candidate(
            &bytes,
            TextFormatWrite::new(),
            self.verify_options,
            TextFormatLayout {
                fields: 1,
                max_depth: 0,
                span: None,
            },
            bytes
                .len()
                .checked_mul(2)
                .ok_or_else(DecodeError::invalid)?,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Strictly decode one native Text `FormatStructArchive`.
pub(crate) fn decode_text_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<TextFormatSnapshot<'_>, DecodeError> {
    Ok(decode_text_format_with_report(source, options)?.0)
}

/// Strictly decode one native Text format and return measured wire usage.
pub(crate) fn decode_text_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(TextFormatSnapshot<'_>, DecodeReport), DecodeError> {
    let (snapshot, _layout, report) = scan_text_format(source, options)?;
    Ok((snapshot, report))
}

/// Prepare a source-preserving native Text format rewrite.
pub(crate) fn prepare_text_format_rewrite<'source>(
    source: &'source [u8],
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedTextFormatRewrite<'source>, DecodeError> {
    validate_text_format_write(write)?;
    let (_snapshot, layout, source_report) = scan_text_format(source, options)?;
    let output_bytes = source.len();
    // Text rewrites are byte-for-byte source preserving, so the emitted
    // candidate has the same scan shape and work as the source preflight.
    // This matters for balanced unknown groups: the bounded scanner charges
    // the enclosing group record and its nested records independently, so
    // the measured work can exceed a simple `2 * output_bytes` estimate.
    let candidate_work = source_report.work_bytes();
    let fields = layout
        .fields
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = source_report
        .work_bytes()
        .checked_add(output_bytes)
        .and_then(|work| work.checked_add(candidate_work))
        .ok_or_else(DecodeError::invalid)?;
    let retained_bytes = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: layout.max_depth,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedTextFormatRewrite {
        source,
        layout,
        write,
        requirements,
        candidate_work,
        verify_options: options,
    })
}

/// Rewrite one native Text format while preserving unknown source fields.
pub(crate) fn rewrite_text_format(
    source: &[u8],
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_text_format_rewrite(source, write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Prepare a canonical native Text format payload for a new list entry.
pub(crate) fn prepare_text_format_write(
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedTextFormatWrite, DecodeError> {
    validate_text_format_write(write)?;
    let output_bytes = text_format_canonical_output_len()?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields: 1,
        work_bytes: output_bytes
            .checked_mul(3)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 0,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedTextFormatWrite {
        requirements,
        verify_options: options,
    })
}

/// Encode a canonical native Text format payload for a new list entry.
pub(crate) fn canonical_text_format(
    write: TextFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_text_format_write(write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Strictly decode one plain-number `FormatStructArchive`.
pub fn decode_number_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<NumberFormatSnapshot<'_>, DecodeError> {
    decode_decimal_format(source, NATIVE_NUMBER_FORMAT_TYPE, options)
}

/// Strictly decode one plain-number format and return measured wire usage.
pub fn decode_number_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(NumberFormatSnapshot<'_>, DecodeReport), DecodeError> {
    decode_decimal_format_with_report(source, NATIVE_NUMBER_FORMAT_TYPE, options)
}

pub(crate) fn decode_decimal_format(
    source: &[u8],
    format_type: u32,
    options: DecodeOptions,
) -> Result<NumberFormatSnapshot<'_>, DecodeError> {
    Ok(decode_decimal_format_with_report(source, format_type, options)?.0)
}

pub(crate) fn decode_decimal_format_with_report(
    source: &[u8],
    format_type: u32,
    options: DecodeOptions,
) -> Result<(NumberFormatSnapshot<'_>, DecodeReport), DecodeError> {
    let (snapshot, _layout, report) = scan_decimal_format(source, format_type, options)?;
    Ok((snapshot, report))
}

/// Prepare a source-preserving plain-number format rewrite.
///
/// Preparation performs the complete strict source preflight, records exact
/// output/work requirements, and allocates no candidate bytes.  Execution
/// must be supplied finite caller-owned limits and performs a strict
/// read-back before publishing the candidate.
pub fn prepare_number_format_rewrite<'source>(
    source: &'source [u8],
    write: NumberFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedNumberFormatRewrite<'source>, DecodeError> {
    prepare_decimal_format_rewrite(source, write, NATIVE_NUMBER_FORMAT_TYPE, options)
}

pub(crate) fn prepare_decimal_format_rewrite<'source>(
    source: &'source [u8],
    write: NumberFormatWrite,
    format_type: u32,
    options: DecodeOptions,
) -> Result<PreparedNumberFormatRewrite<'source>, DecodeError> {
    validate_decimal_format_write(write, format_type)?;
    let (_snapshot, layout, source_report) = scan_decimal_format(source, format_type, options)?;
    let output_bytes = decimal_format_rewrite_output_len(source, layout, write, format_type)?;
    let candidate_work = output_bytes
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    let fields = layout
        .fields
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    let work_bytes = source_report
        .work_bytes()
        .checked_add(output_bytes)
        .and_then(|work| work.checked_add(candidate_work))
        .ok_or_else(DecodeError::invalid)?;
    let retained_bytes = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: layout.max_depth,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedNumberFormatRewrite {
        source,
        layout,
        write,
        format_type,
        requirements,
        verify_options: options,
    })
}

/// Rewrite one plain-number format while preserving unknown source fields.
pub fn rewrite_number_format(
    source: &[u8],
    write: NumberFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    rewrite_decimal_format(source, write, NATIVE_NUMBER_FORMAT_TYPE, options)
}

pub(crate) fn rewrite_decimal_format(
    source: &[u8],
    write: NumberFormatWrite,
    format_type: u32,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_decimal_format_rewrite(source, write, format_type, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for the table-cell format route.
pub use rewrite_number_format as rewrite_table_cell_number_format;

/// Prepared source-preserving plain-number rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedNumberFormatRewrite<'source> {
    source: &'source [u8],
    layout: DecimalFormatLayout,
    write: NumberFormatWrite,
    format_type: u32,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedNumberFormatRewrite<'_> {
    /// Return the measured requirements that execution must be allowed.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Return a report-shaped view of the prepared operation.
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    /// Emit, strictly read back, and publish the source-preserving candidate.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_decimal_format_rewrite(
            &mut bytes,
            self.source,
            self.layout,
            self.write,
            self.format_type,
        )?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        verify_decimal_format_candidate(
            &bytes,
            self.write,
            self.format_type,
            self.verify_options,
            self.layout,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepare a canonical plain-number format payload for a new list entry.
///
/// Unlike [`prepare_number_format_rewrite`], this constructor has no source
/// payload and therefore has no unknown fields to preserve.  It is intended
/// for a storage append where the format entry did not previously exist.
pub fn prepare_number_format_write(
    write: NumberFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedNumberFormatWrite, DecodeError> {
    prepare_decimal_format_write(write, NATIVE_NUMBER_FORMAT_TYPE, options)
}

pub(crate) fn prepare_decimal_format_write(
    write: NumberFormatWrite,
    format_type: u32,
    options: DecodeOptions,
) -> Result<PreparedNumberFormatWrite, DecodeError> {
    validate_decimal_format_write(write, format_type)?;
    let output_bytes = decimal_format_canonical_output_len(write, format_type)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields: 4,
        work_bytes: output_bytes
            .checked_mul(3)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 0,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedNumberFormatWrite {
        write,
        format_type,
        requirements,
        verify_options: options,
    })
}

/// Prepare a canonical format append under the more explicit append spelling.
pub use prepare_number_format_write as prepare_number_format_append;

/// Encode a canonical plain-number format payload for a new list entry.
pub fn canonical_number_format(
    write: NumberFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_decimal_format(write, NATIVE_NUMBER_FORMAT_TYPE, options)
}

pub(crate) fn canonical_decimal_format(
    write: NumberFormatWrite,
    format_type: u32,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_decimal_format_write(write, format_type, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Prepared canonical plain-number format append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedNumberFormatWrite {
    write: NumberFormatWrite,
    format_type: u32,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedNumberFormatWrite {
    /// Return the finite measured requirements for execution.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Return a report-shaped view of the prepared operation.
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    /// Emit and strictly read back a canonical format payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_decimal_format_canonical(&mut bytes, self.write, self.format_type)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        let layout = DecimalFormatLayout {
            fields: 4,
            max_depth: 0,
            spans: [None, None, None, None],
        };
        verify_decimal_format_candidate(
            &bytes,
            self.write,
            self.format_type,
            self.verify_options,
            layout,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Borrowed semantic facts for one strict native Currency
/// `FormatStructArchive`.
///
/// Currency deliberately has its own nominal snapshot even though the wire
/// implementation shares the decimal scanner. This keeps a Number or
/// Percentage snapshot from being passed to a Currency writer accidentally.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CurrencyFormatSnapshot<'source> {
    source: &'source [u8],
    format_type: u32,
    decimal_places: Option<u32>,
    currency_code: Option<&'source str>,
    negative_style: Option<u32>,
    show_thousands_separator: Option<bool>,
    use_accounting_style: Option<bool>,
}

impl<'source> CurrencyFormatSnapshot<'source> {
    pub(crate) const fn raw(self) -> &'source [u8] {
        self.source
    }

    pub(crate) const fn format_type(self) -> u32 {
        self.format_type
    }

    pub(crate) const fn decimal_places(self) -> Option<u32> {
        self.decimal_places
    }

    pub(crate) const fn currency_code(self) -> Option<&'source str> {
        self.currency_code
    }

    pub(crate) const fn negative_style(self) -> Option<u32> {
        self.negative_style
    }

    pub(crate) const fn show_thousands_separator(self) -> Option<bool> {
        self.show_thousands_separator
    }

    pub(crate) const fn use_accounting_style(self) -> Option<bool> {
        self.use_accounting_style
    }
}

/// Borrowed scalar values accepted by the strict native Currency writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CurrencyFormatWrite<'source> {
    decimal_places: Option<u32>,
    currency_code: Option<&'source str>,
    negative_style: Option<u32>,
    show_thousands_separator: Option<bool>,
    use_accounting_style: Option<bool>,
}

impl<'source> CurrencyFormatWrite<'source> {
    pub(crate) const fn new(
        currency_code: &'source str,
        decimal_places: u32,
        negative_style: u32,
        show_thousands_separator: bool,
        use_accounting_style: bool,
    ) -> Self {
        Self {
            decimal_places: Some(decimal_places),
            currency_code: Some(currency_code),
            negative_style: Some(negative_style),
            show_thousands_separator: Some(show_thousands_separator),
            use_accounting_style: Some(use_accounting_style),
        }
    }

    pub(crate) const fn from_snapshot(snapshot: CurrencyFormatSnapshot<'source>) -> Self {
        Self {
            decimal_places: snapshot.decimal_places,
            currency_code: snapshot.currency_code,
            negative_style: snapshot.negative_style,
            show_thousands_separator: snapshot.show_thousands_separator,
            use_accounting_style: snapshot.use_accounting_style,
        }
    }

    pub(crate) const fn decimal_places(self) -> Option<u32> {
        self.decimal_places
    }

    pub(crate) const fn currency_code(self) -> Option<&'source str> {
        self.currency_code
    }

    pub(crate) const fn negative_style(self) -> Option<u32> {
        self.negative_style
    }

    pub(crate) const fn show_thousands_separator(self) -> Option<bool> {
        self.show_thousands_separator
    }

    pub(crate) const fn use_accounting_style(self) -> Option<bool> {
        self.use_accounting_style
    }

    pub(crate) const fn with_decimal_places(mut self, value: u32) -> Self {
        self.decimal_places = Some(value);
        self
    }

    pub(crate) const fn with_currency_code(mut self, value: &'source str) -> Self {
        self.currency_code = Some(value);
        self
    }

    pub(crate) const fn with_negative_style(mut self, value: u32) -> Self {
        self.negative_style = Some(value);
        self
    }

    pub(crate) const fn with_show_thousands_separator(mut self, value: bool) -> Self {
        self.show_thousands_separator = Some(value);
        self
    }

    pub(crate) const fn with_use_accounting_style(mut self, value: bool) -> Self {
        self.use_accounting_style = Some(value);
        self
    }
}

/// Prepared source-preserving native Currency rewrite.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PreparedCurrencyFormatRewrite<'source> {
    source: &'source [u8],
    layout: CurrencyFormatLayout,
    write: CurrencyFormatWrite<'source>,
    requirements: RewriteExecutionRequirements,
    candidate_work: usize,
    verify_options: DecodeOptions,
}

impl PreparedCurrencyFormatRewrite<'_> {
    pub(crate) const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub(crate) fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    pub(crate) fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_currency_format_rewrite(&mut bytes, self.source, self.layout, self.write)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        verify_currency_format_candidate(
            &bytes,
            self.write,
            self.verify_options,
            self.layout,
            self.requirements,
            self.candidate_work,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical native Currency append.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PreparedCurrencyFormatWrite<'source> {
    write: CurrencyFormatWrite<'source>,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedCurrencyFormatWrite<'_> {
    pub(crate) const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub(crate) fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    pub(crate) fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_currency_format_canonical(&mut bytes, self.write)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        let layout = CurrencyFormatLayout {
            max_depth: 0,
            spans: [None, None, None, None, None, None],
        };
        verify_currency_format_candidate(
            &bytes,
            self.write,
            self.verify_options,
            layout,
            self.requirements,
            bytes
                .len()
                .checked_mul(2)
                .ok_or_else(DecodeError::invalid)?,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

#[derive(Debug, Clone, Copy)]
struct CurrencyFieldSpan {
    slot: usize,
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy)]
struct CurrencyFormatLayout {
    max_depth: u32,
    spans: [Option<CurrencyFieldSpan>; 6],
}

#[derive(Debug, Clone, Copy)]
struct CurrencyFormatScanState<'source> {
    format_type: Option<u32>,
    decimal_places: Option<u32>,
    currency_code: Option<&'source str>,
    negative_style: Option<u32>,
    show_thousands_separator: Option<bool>,
    use_accounting_style: Option<bool>,
    spans: [Option<CurrencyFieldSpan>; 6],
}

impl<'source> CurrencyFormatScanState<'source> {
    const fn new() -> Self {
        Self {
            format_type: None,
            decimal_places: None,
            currency_code: None,
            negative_style: None,
            show_thousands_separator: None,
            use_accounting_style: None,
            spans: [None, None, None, None, None, None],
        }
    }
}

/// Strictly decode one native Currency `FormatStructArchive`.
pub(crate) fn decode_currency_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CurrencyFormatSnapshot<'_>, DecodeError> {
    Ok(decode_currency_format_with_report(source, options)?.0)
}

/// Strictly decode one native Currency format and return measured wire use.
pub(crate) fn decode_currency_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CurrencyFormatSnapshot<'_>, DecodeReport), DecodeError> {
    let (snapshot, _layout, report) = scan_currency_format(source, options)?;
    Ok((snapshot, report))
}

/// Prepare a source-preserving native Currency rewrite.
pub(crate) fn prepare_currency_format_rewrite<'source>(
    source: &'source [u8],
    write: CurrencyFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedCurrencyFormatRewrite<'source>, DecodeError> {
    validate_currency_format_write(write, options.max_text_bytes)?;
    let (snapshot, layout, source_report) = scan_currency_format(source, options)?;
    let output_bytes = currency_format_rewrite_output_len(source, layout, write)?;
    let old_selected_bytes = currency_format_selected_source_len(layout)?;
    let new_selected_bytes = currency_format_canonical_output_len(write)?;
    let source_parse_work = source_report
        .work_bytes()
        .checked_sub(source.len())
        .ok_or_else(DecodeError::invalid)?;
    let candidate_parse_work = source_parse_work
        .checked_sub(old_selected_bytes)
        .and_then(|work| work.checked_add(new_selected_bytes))
        .ok_or_else(DecodeError::invalid)?;
    let candidate_work = candidate_parse_work
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let fields = source_report
        .fields()
        .checked_sub(currency_format_selected_field_count(layout))
        .and_then(|fields| fields.checked_add(currency_format_field_count(write)))
        .ok_or_else(DecodeError::invalid)?;
    let old_text_bytes = snapshot.currency_code().map_or(0, str::len);
    let new_text_bytes = write.currency_code().map_or(0, str::len);
    let text_bytes = source_report
        .text_bytes()
        .checked_sub(old_text_bytes)
        .and_then(|text| text.checked_add(new_text_bytes))
        .ok_or_else(DecodeError::invalid)?;
    let retained_bytes = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: source_report
            .work_bytes()
            .checked_add(output_bytes)
            .and_then(|work| work.checked_add(candidate_work))
            .ok_or_else(DecodeError::invalid)?,
        max_depth: layout.max_depth,
        references: 0,
        items: 0,
        text_bytes,
        allocations: 1,
        retained_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedCurrencyFormatRewrite {
        source,
        layout,
        write,
        requirements,
        candidate_work,
        verify_options: options,
    })
}

/// Rewrite one native Currency format while preserving unknown source fields.
pub(crate) fn rewrite_currency_format(
    source: &[u8],
    write: CurrencyFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_currency_format_rewrite(source, write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Prepare a canonical native Currency format payload for a new list entry.
pub(crate) fn prepare_currency_format_write<'source>(
    write: CurrencyFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedCurrencyFormatWrite<'source>, DecodeError> {
    validate_currency_format_write(write, options.max_text_bytes)?;
    let output_bytes = currency_format_canonical_output_len(write)?;
    let text_bytes = write.currency_code().map_or(0, str::len);
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields: currency_format_field_count(write),
        work_bytes: output_bytes
            .checked_mul(3)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 0,
        references: 0,
        items: 0,
        text_bytes,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedCurrencyFormatWrite {
        write,
        requirements,
        verify_options: options,
    })
}

/// Encode a canonical native Currency format payload for a new list entry.
pub(crate) fn canonical_currency_format(
    write: CurrencyFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_currency_format_write(write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

fn scan_currency_format<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<
    (
        CurrencyFormatSnapshot<'source>,
        CurrencyFormatLayout,
        DecodeReport,
    ),
    DecodeError,
> {
    let mut budget = Budget::new(source, options)?;
    let mut state = CurrencyFormatScanState::new();
    let mut offset = 0usize;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        inspect_currency_root_field(offset, field, &mut budget, &mut state)?;
        offset = field.end;
    }
    let format_type = state.format_type.ok_or_else(DecodeError::invalid)?;
    if format_type != NATIVE_CURRENCY_FORMAT_TYPE
        || state.decimal_places.is_none()
        || state.currency_code.is_none()
        || state.negative_style.is_none()
        || state.show_thousands_separator.is_none()
        || state.use_accounting_style.is_none()
        || state.currency_code.is_some() != state.use_accounting_style.is_some()
        || state.decimal_places.is_some_and(|value| {
            value != NATIVE_AUTOMATIC_DECIMAL_PLACES && value > MAX_NUMBER_DECIMAL_PLACES
        })
        || state.negative_style.is_some_and(|value| value > 3)
    {
        return Err(DecodeError::invalid());
    }
    let snapshot = CurrencyFormatSnapshot {
        source,
        format_type,
        decimal_places: state.decimal_places,
        currency_code: state.currency_code,
        negative_style: state.negative_style,
        show_thousands_separator: state.show_thousands_separator,
        use_accounting_style: state.use_accounting_style,
    };
    buffa_currency_format_parity(source, snapshot, &mut budget)?;
    let report = budget.finish(0);
    let layout = CurrencyFormatLayout {
        max_depth: report.max_depth(),
        spans: state.spans,
    };
    Ok((snapshot, layout, report))
}

fn inspect_currency_root_field<'source>(
    start: usize,
    field: Field<'source>,
    budget: &mut Budget,
    state: &mut CurrencyFormatScanState<'source>,
) -> Result<(), DecodeError> {
    let Some(slot) = currency_format_slot(field.number) else {
        // TSK fields through 45 are known FormatStructArchive fields with a
        // different shape/meaning. Treating one as opaque would let a caller
        // publish another display format through the Currency API.
        if field.number <= NUMBER_FORMAT_MAX_KNOWN_FIELD {
            return Err(DecodeError::invalid());
        }
        budget.mark_unknown();
        return Ok(());
    };
    if state.spans[slot].is_some() {
        return Err(DecodeError::invalid());
    }
    let span = CurrencyFieldSpan {
        slot,
        start,
        end: field.end,
    };
    match slot {
        0 | 1 | 3 => {
            let value = u32::try_from(field.known_varint()?).map_err(|_| DecodeError::invalid())?;
            match slot {
                0 => state.format_type = Some(value),
                1 => state.decimal_places = Some(value),
                3 => state.negative_style = Some(value),
                _ => unreachable!("currency integer slot is exhaustive"),
            }
        },
        2 => {
            if field.wire != 2 {
                return Err(DecodeError::invalid());
            }
            let payload = field.payload.ok_or_else(DecodeError::invalid)?;
            let value = str::from_utf8(payload).map_err(|_| DecodeError::invalid())?;
            budget.text(payload.len())?;
            validate_currency_code(value)?;
            state.currency_code = Some(value);
        },
        4 => set_bool(&mut state.show_thousands_separator, field, false)?,
        5 => set_bool(&mut state.use_accounting_style, field, false)?,
        _ => unreachable!("currency format slot is exhaustive"),
    }
    state.spans[slot] = Some(span);
    Ok(())
}

fn buffa_currency_format_parity(
    source: &[u8],
    snapshot: CurrencyFormatSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let options = budget.options;
    let view: currency_projection::FormatStructArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.format_type != Some(snapshot.format_type)
        || view.decimal_places != snapshot.decimal_places
        || view.currency_code != snapshot.currency_code
        || view.negative_style != snapshot.negative_style
        || view.show_thousands_separator != snapshot.show_thousands_separator
        || view.use_accounting_style != snapshot.use_accounting_style
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())?;
    Ok(())
}

fn validate_currency_code(value: &str) -> Result<(), DecodeError> {
    if value.len() != 3 || !value.bytes().all(|byte| byte.is_ascii_uppercase()) {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn validate_currency_format_write(
    write: CurrencyFormatWrite<'_>,
    max_text_bytes: usize,
) -> Result<(), DecodeError> {
    if write.decimal_places.is_none()
        || write.currency_code.is_none()
        || write.negative_style.is_none()
        || write.show_thousands_separator.is_none()
        || write.use_accounting_style.is_none()
        || write.currency_code.is_some() != write.use_accounting_style.is_some()
        || write.decimal_places.is_some_and(|value| {
            value != NATIVE_AUTOMATIC_DECIMAL_PLACES && value > MAX_NUMBER_DECIMAL_PLACES
        })
        || write.negative_style.is_some_and(|value| value > 3)
    {
        return Err(DecodeError::invalid());
    }
    if let Some(value) = write.currency_code {
        validate_text(value, max_text_bytes)?;
        validate_currency_code(value)?;
    }
    Ok(())
}

fn currency_format_slot(number: u32) -> Option<usize> {
    match number {
        FORMAT_TYPE_FIELD => Some(0),
        FORMAT_DECIMAL_PLACES_FIELD => Some(1),
        FORMAT_CURRENCY_CODE_FIELD => Some(2),
        FORMAT_NEGATIVE_STYLE_FIELD => Some(3),
        FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => Some(4),
        FORMAT_USE_ACCOUNTING_STYLE_FIELD => Some(5),
        _ => None,
    }
}

fn currency_format_field_count(write: CurrencyFormatWrite<'_>) -> usize {
    1usize
        + usize::from(write.decimal_places.is_some())
        + usize::from(write.currency_code.is_some())
        + usize::from(write.negative_style.is_some())
        + usize::from(write.show_thousands_separator.is_some())
        + usize::from(write.use_accounting_style.is_some())
}

fn currency_format_selected_field_count(layout: CurrencyFormatLayout) -> usize {
    layout.spans.into_iter().filter(Option::is_some).count()
}

fn currency_format_selected_source_len(layout: CurrencyFormatLayout) -> Result<usize, DecodeError> {
    layout
        .spans
        .into_iter()
        .flatten()
        .try_fold(0usize, |length, span| {
            length
                .checked_add(
                    span.end
                        .checked_sub(span.start)
                        .ok_or_else(DecodeError::invalid)?,
                )
                .ok_or_else(DecodeError::invalid)
        })
}

fn currency_format_rewrite_output_len(
    source: &[u8],
    layout: CurrencyFormatLayout,
    write: CurrencyFormatWrite<'_>,
) -> Result<usize, DecodeError> {
    let old = currency_format_selected_source_len(layout)?;
    let new = currency_format_canonical_output_len(write)?;
    source
        .len()
        .checked_sub(old)
        .and_then(|length| length.checked_add(new))
        .ok_or_else(DecodeError::invalid)
}

fn currency_format_canonical_output_len(
    write: CurrencyFormatWrite<'_>,
) -> Result<usize, DecodeError> {
    let mut length =
        number_format_field_len(FORMAT_TYPE_FIELD, u64::from(NATIVE_CURRENCY_FORMAT_TYPE))?;
    if let Some(value) = write.decimal_places {
        length = length
            .checked_add(number_format_field_len(
                FORMAT_DECIMAL_PLACES_FIELD,
                u64::from(value),
            )?)
            .ok_or_else(DecodeError::invalid)?;
    }
    if let Some(value) = write.currency_code {
        length = length
            .checked_add(length_field_len(FORMAT_CURRENCY_CODE_FIELD, value.len())?)
            .ok_or_else(DecodeError::invalid)?;
    }
    if let Some(value) = write.negative_style {
        length = length
            .checked_add(number_format_field_len(
                FORMAT_NEGATIVE_STYLE_FIELD,
                u64::from(value),
            )?)
            .ok_or_else(DecodeError::invalid)?;
    }
    if let Some(value) = write.show_thousands_separator {
        length = length
            .checked_add(number_format_field_len(
                FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
                u64::from(value),
            )?)
            .ok_or_else(DecodeError::invalid)?;
    }
    if let Some(value) = write.use_accounting_style {
        length = length
            .checked_add(number_format_field_len(
                FORMAT_USE_ACCOUNTING_STYLE_FIELD,
                u64::from(value),
            )?)
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn currency_format_write_present(write: CurrencyFormatWrite<'_>, slot: usize) -> bool {
    match slot {
        0 => true,
        1 => write.decimal_places.is_some(),
        2 => write.currency_code.is_some(),
        3 => write.negative_style.is_some(),
        4 => write.show_thousands_separator.is_some(),
        5 => write.use_accounting_style.is_some(),
        _ => false,
    }
}

fn emit_currency_format_field(
    output: &mut Vec<u8>,
    slot: usize,
    write: CurrencyFormatWrite<'_>,
) -> Result<(), DecodeError> {
    match slot {
        0 => emit_varint_field(
            output,
            FORMAT_TYPE_FIELD,
            u64::from(NATIVE_CURRENCY_FORMAT_TYPE),
        ),
        1 => emit_varint_field(
            output,
            FORMAT_DECIMAL_PLACES_FIELD,
            u64::from(write.decimal_places.ok_or_else(DecodeError::invalid)?),
        ),
        2 => emit_len_field(
            output,
            FORMAT_CURRENCY_CODE_FIELD,
            write
                .currency_code
                .ok_or_else(DecodeError::invalid)?
                .as_bytes(),
        ),
        3 => emit_varint_field(
            output,
            FORMAT_NEGATIVE_STYLE_FIELD,
            u64::from(write.negative_style.ok_or_else(DecodeError::invalid)?),
        ),
        4 => emit_varint_field(
            output,
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(
                write
                    .show_thousands_separator
                    .ok_or_else(DecodeError::invalid)?,
            ),
        ),
        5 => emit_varint_field(
            output,
            FORMAT_USE_ACCOUNTING_STYLE_FIELD,
            u64::from(
                write
                    .use_accounting_style
                    .ok_or_else(DecodeError::invalid)?,
            ),
        ),
        _ => Err(DecodeError::invalid()),
    }
}

fn emit_currency_format_canonical(
    output: &mut Vec<u8>,
    write: CurrencyFormatWrite<'_>,
) -> Result<(), DecodeError> {
    validate_currency_format_write(write, usize::MAX)?;
    for slot in 0..6 {
        if currency_format_write_present(write, slot) {
            emit_currency_format_field(output, slot, write)?;
        }
    }
    Ok(())
}

fn emit_currency_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    layout: CurrencyFormatLayout,
    write: CurrencyFormatWrite<'_>,
) -> Result<(), DecodeError> {
    let mut ordered = [
        layout.spans[0],
        layout.spans[1],
        layout.spans[2],
        layout.spans[3],
        layout.spans[4],
        layout.spans[5],
    ];
    let mut source_offset = 0usize;
    for index in 1..ordered.len() {
        let mut cursor = index;
        while cursor > 0
            && ordered[cursor].is_some()
            && (ordered[cursor - 1].is_none()
                || ordered[cursor].as_ref().is_some_and(|current| {
                    ordered[cursor - 1]
                        .as_ref()
                        .is_some_and(|previous| current.start < previous.start)
                }))
        {
            ordered.swap(cursor, cursor - 1);
            cursor -= 1;
        }
    }
    for span in ordered.into_iter().flatten() {
        output.extend_from_slice(
            source
                .get(source_offset..span.start)
                .ok_or_else(DecodeError::invalid)?,
        );
        if currency_format_write_present(write, span.slot) {
            emit_currency_format_field(output, span.slot, write)?;
        }
        source_offset = span.end;
    }
    output.extend_from_slice(
        source
            .get(source_offset..)
            .ok_or_else(DecodeError::invalid)?,
    );
    for slot in 0..6 {
        if layout.spans[slot].is_none() && currency_format_write_present(write, slot) {
            emit_currency_format_field(output, slot, write)?;
        }
    }
    Ok(())
}

fn verify_currency_format_candidate(
    source: &[u8],
    write: CurrencyFormatWrite<'_>,
    options: DecodeOptions,
    layout: CurrencyFormatLayout,
    requirements: RewriteExecutionRequirements,
    candidate_work: usize,
) -> Result<(), DecodeError> {
    let (snapshot, report) = decode_currency_format_with_report(source, options)?;
    if report.fields() != requirements.fields()
        || report.work_bytes() != candidate_work
        || report.max_depth() != layout.max_depth
        || report.text_bytes() != requirements.text_bytes()
        || CurrencyFormatWrite::from_snapshot(snapshot) != write
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

/// Borrowed semantic facts for one strict native Fraction
/// `FormatStructArchive`.
///
/// Fraction owns only the native discriminator and denominator strategy. The
/// complete source payload remains available through [`Self::raw`], while
/// unknown extension records remain source-authoritative for rewrites.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct FractionFormatSnapshot<'source> {
    source: &'source [u8],
    format_type: u32,
    fraction_accuracy: u32,
    requires_fraction_replacement: Option<bool>,
}

impl<'source> FractionFormatSnapshot<'source> {
    pub(crate) const fn raw(self) -> &'source [u8] {
        self.source
    }

    pub(crate) const fn format_type(self) -> u32 {
        self.format_type
    }

    pub(crate) const fn fraction_accuracy(self) -> u32 {
        self.fraction_accuracy
    }

    /// Field 20 is a custom-number-format flag. An explicit canonical false
    /// is retained in the snapshot and preserved by rewrites; true is rejected
    /// because this seam cannot safely implement replacement semantics.
    pub(crate) const fn requires_fraction_replacement(self) -> Option<bool> {
        self.requires_fraction_replacement
    }
}

/// Scalar values accepted by the strict native Fraction writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct FractionFormatWrite {
    fraction_accuracy: u32,
}

impl FractionFormatWrite {
    pub(crate) const fn new(fraction_accuracy: u32) -> Self {
        Self { fraction_accuracy }
    }

    pub(crate) const fn from_snapshot(snapshot: FractionFormatSnapshot<'_>) -> Self {
        Self {
            fraction_accuracy: snapshot.fraction_accuracy,
        }
    }

    pub(crate) const fn fraction_accuracy(self) -> u32 {
        self.fraction_accuracy
    }

    pub(crate) const fn with_fraction_accuracy(mut self, value: u32) -> Self {
        self.fraction_accuracy = value;
        self
    }
}

/// Prepared source-preserving native Fraction rewrite.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PreparedFractionFormatRewrite<'source> {
    source: &'source [u8],
    layout: FractionFormatLayout,
    write: FractionFormatWrite,
    requirements: RewriteExecutionRequirements,
    candidate_work: usize,
    verify_options: DecodeOptions,
}

impl PreparedFractionFormatRewrite<'_> {
    pub(crate) const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub(crate) fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    pub(crate) fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_fraction_format_rewrite(&mut bytes, self.source, self.layout, self.write)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        verify_fraction_format_candidate(
            &bytes,
            self.write,
            self.verify_options,
            self.layout,
            self.requirements,
            self.candidate_work,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical native Fraction append.
#[derive(Debug, Clone, Copy)]
pub(crate) struct PreparedFractionFormatWrite {
    write: FractionFormatWrite,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedFractionFormatWrite {
    pub(crate) const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub(crate) fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    pub(crate) fn execute(
        self,
        limits: RewriteExecutionLimits,
    ) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_fraction_format_canonical(&mut bytes, self.write)?;
        if bytes.len() != self.requirements.output_bytes {
            return Err(DecodeError::invalid());
        }
        let layout = FractionFormatLayout {
            max_depth: 0,
            spans: [None, None],
        };
        verify_fraction_format_candidate(
            &bytes,
            self.write,
            self.verify_options,
            layout,
            self.requirements,
            bytes
                .len()
                .checked_mul(2)
                .ok_or_else(DecodeError::invalid)?,
        )?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

#[derive(Debug, Clone, Copy)]
struct FractionFieldSpan {
    slot: usize,
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy)]
struct FractionFormatLayout {
    max_depth: u32,
    spans: [Option<FractionFieldSpan>; 2],
}

#[derive(Debug, Clone, Copy)]
struct FractionFormatScanState {
    format_type: Option<u32>,
    fraction_accuracy: Option<u32>,
    requires_fraction_replacement: Option<bool>,
    spans: [Option<FractionFieldSpan>; 2],
}

impl FractionFormatScanState {
    const fn new() -> Self {
        Self {
            format_type: None,
            fraction_accuracy: None,
            requires_fraction_replacement: None,
            spans: [None, None],
        }
    }
}

/// Strictly decode one native Fraction `FormatStructArchive`.
pub(crate) fn decode_fraction_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FractionFormatSnapshot<'_>, DecodeError> {
    Ok(decode_fraction_format_with_report(source, options)?.0)
}

/// Strictly decode one native Fraction format and return measured wire use.
pub(crate) fn decode_fraction_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FractionFormatSnapshot<'_>, DecodeReport), DecodeError> {
    let (snapshot, _layout, report) = scan_fraction_format(source, options)?;
    Ok((snapshot, report))
}

/// Prepare a source-preserving native Fraction rewrite.
pub(crate) fn prepare_fraction_format_rewrite<'source>(
    source: &'source [u8],
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedFractionFormatRewrite<'source>, DecodeError> {
    validate_fraction_format_write(write)?;
    let (_snapshot, layout, source_report) = scan_fraction_format(source, options)?;
    let output_bytes = fraction_format_rewrite_output_len(source, layout, write)?;
    let old_selected_bytes = fraction_format_selected_source_len(layout)?;
    let new_selected_bytes = fraction_format_canonical_output_len(write)?;
    let source_parse_work = source_report
        .work_bytes()
        .checked_sub(source.len())
        .ok_or_else(DecodeError::invalid)?;
    let candidate_parse_work = source_parse_work
        .checked_sub(old_selected_bytes)
        .and_then(|work| work.checked_add(new_selected_bytes))
        .ok_or_else(DecodeError::invalid)?;
    let candidate_work = candidate_parse_work
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let fields = source_report
        .fields()
        .checked_sub(fraction_format_selected_field_count(layout))
        .and_then(|fields| fields.checked_add(fraction_format_field_count(write)))
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: source_report
            .work_bytes()
            .checked_add(output_bytes)
            .and_then(|work| work.checked_add(candidate_work))
            .ok_or_else(DecodeError::invalid)?,
        max_depth: layout.max_depth,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: source
            .len()
            .checked_add(output_bytes)
            .ok_or_else(DecodeError::invalid)?,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedFractionFormatRewrite {
        source,
        layout,
        write,
        requirements,
        candidate_work,
        verify_options: options,
    })
}

/// Rewrite one native Fraction format while preserving unknown source fields.
pub(crate) fn rewrite_fraction_format(
    source: &[u8],
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_fraction_format_rewrite(source, write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Prepare a canonical native Fraction format payload for a new list entry.
pub(crate) fn prepare_fraction_format_write(
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<PreparedFractionFormatWrite, DecodeError> {
    validate_fraction_format_write(write)?;
    let output_bytes = fraction_format_canonical_output_len(write)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields: fraction_format_field_count(write),
        work_bytes: output_bytes
            .checked_mul(3)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 0,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedFractionFormatWrite {
        write,
        requirements,
        verify_options: options,
    })
}

/// Encode a canonical native Fraction format payload for a new list entry.
pub(crate) fn canonical_fraction_format(
    write: FractionFormatWrite,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_fraction_format_write(write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

fn scan_fraction_format<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<
    (
        FractionFormatSnapshot<'source>,
        FractionFormatLayout,
        DecodeReport,
    ),
    DecodeError,
> {
    let mut budget = Budget::new(source, options)?;
    let mut state = FractionFormatScanState::new();
    let mut offset = 0usize;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        inspect_fraction_root_field(offset, field, &mut state)?;
        offset = field.end;
    }
    let format_type = state.format_type.ok_or_else(DecodeError::invalid)?;
    let fraction_accuracy = state.fraction_accuracy.ok_or_else(DecodeError::invalid)?;
    if format_type != NATIVE_FRACTION_FORMAT_TYPE
        || !is_valid_fraction_accuracy(fraction_accuracy)
        || state.spans[0].is_none()
        || state.spans[1].is_none()
    {
        return Err(DecodeError::invalid());
    }
    let snapshot = FractionFormatSnapshot {
        source,
        format_type,
        fraction_accuracy,
        requires_fraction_replacement: state.requires_fraction_replacement,
    };
    buffa_fraction_format_parity(source, snapshot, &mut budget)?;
    // Field 20 is reserved for custom-number formatting. An explicit false
    // is a harmless legacy marker and remains source-authoritative; true
    // requests replacement behavior that this seam cannot safely implement.
    if snapshot.requires_fraction_replacement() == Some(true) {
        return Err(DecodeError::invalid());
    }
    let report = budget.finish(0);
    let layout = FractionFormatLayout {
        max_depth: report.max_depth(),
        spans: state.spans,
    };
    Ok((snapshot, layout, report))
}

fn inspect_fraction_root_field(
    start: usize,
    field: Field<'_>,
    state: &mut FractionFormatScanState,
) -> Result<(), DecodeError> {
    match field.number {
        FORMAT_TYPE_FIELD | FORMAT_FRACTION_ACCURACY_FIELD => {
            let slot = if field.number == FORMAT_TYPE_FIELD {
                0
            } else {
                1
            };
            if field.wire != 0 || state.spans[slot].is_some() {
                return Err(DecodeError::invalid());
            }
            let value = u32::try_from(field.known_varint()?).map_err(|_| DecodeError::invalid())?;
            state.spans[slot] = Some(FractionFieldSpan {
                slot,
                start,
                end: field.end,
            });
            if slot == 0 {
                state.format_type = Some(value);
            } else {
                state.fraction_accuracy = Some(value);
            }
        },
        FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD => {
            if field.wire != 0 || state.requires_fraction_replacement.is_some() {
                return Err(DecodeError::invalid());
            }
            let value = field.known_varint()?;
            if value > 1 {
                return Err(DecodeError::invalid());
            }
            state.requires_fraction_replacement = Some(value != 0);
        },
        number if number <= NUMBER_FORMAT_MAX_KNOWN_FIELD => {
            // All other fields in TSK.FormatStructArchive belong to another
            // display family (or to custom/control metadata) and are not
            // safe to accept through the ordinary Fraction route.
            return Err(DecodeError::invalid());
        },
        _ => {},
    }
    Ok(())
}

fn buffa_fraction_format_parity(
    source: &[u8],
    snapshot: FractionFormatSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let options = budget.options;
    let view: crate::buffa_numbers_table_cell_fraction_format_generated::LitchiIwaNumbersTableCellFractionFormatProjection::FormatStructArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.format_type != Some(snapshot.format_type)
        || view.fraction_accuracy != Some(snapshot.fraction_accuracy)
        || view.requires_fraction_replacement != snapshot.requires_fraction_replacement
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())?;
    Ok(())
}

fn is_valid_fraction_accuracy(value: u32) -> bool {
    matches!(
        value,
        NATIVE_FRACTION_UP_TO_ONE_DIGIT
            | NATIVE_FRACTION_UP_TO_TWO_DIGITS
            | NATIVE_FRACTION_UP_TO_THREE_DIGITS
            | NATIVE_FRACTION_HALVES
            | NATIVE_FRACTION_QUARTERS
            | NATIVE_FRACTION_EIGHTHS
            | NATIVE_FRACTION_SIXTEENTHS
            | NATIVE_FRACTION_TENTHS
            | NATIVE_FRACTION_HUNDREDTHS
    )
}

fn validate_fraction_format_write(write: FractionFormatWrite) -> Result<(), DecodeError> {
    if is_valid_fraction_accuracy(write.fraction_accuracy) {
        Ok(())
    } else {
        Err(DecodeError::invalid())
    }
}

fn fraction_format_field_count(_write: FractionFormatWrite) -> usize {
    2
}

fn fraction_format_selected_field_count(layout: FractionFormatLayout) -> usize {
    layout.spans.into_iter().filter(Option::is_some).count()
}

fn fraction_format_selected_source_len(layout: FractionFormatLayout) -> Result<usize, DecodeError> {
    layout
        .spans
        .into_iter()
        .flatten()
        .try_fold(0usize, |length, span| {
            length
                .checked_add(
                    span.end
                        .checked_sub(span.start)
                        .ok_or_else(DecodeError::invalid)?,
                )
                .ok_or_else(DecodeError::invalid)
        })
}

fn fraction_format_rewrite_output_len(
    source: &[u8],
    layout: FractionFormatLayout,
    write: FractionFormatWrite,
) -> Result<usize, DecodeError> {
    let old = fraction_format_selected_source_len(layout)?;
    let new = fraction_format_canonical_output_len(write)?;
    source
        .len()
        .checked_sub(old)
        .and_then(|length| length.checked_add(new))
        .ok_or_else(DecodeError::invalid)
}

fn fraction_format_canonical_output_len(write: FractionFormatWrite) -> Result<usize, DecodeError> {
    validate_fraction_format_write(write)?;
    number_format_field_len(FORMAT_TYPE_FIELD, u64::from(NATIVE_FRACTION_FORMAT_TYPE))?
        .checked_add(number_format_field_len(
            FORMAT_FRACTION_ACCURACY_FIELD,
            u64::from(write.fraction_accuracy),
        )?)
        .ok_or_else(DecodeError::invalid)
}

fn emit_fraction_format_field(
    output: &mut Vec<u8>,
    slot: usize,
    write: FractionFormatWrite,
) -> Result<(), DecodeError> {
    match slot {
        0 => emit_varint_field(
            output,
            FORMAT_TYPE_FIELD,
            u64::from(NATIVE_FRACTION_FORMAT_TYPE),
        ),
        1 => emit_varint_field(
            output,
            FORMAT_FRACTION_ACCURACY_FIELD,
            u64::from(write.fraction_accuracy),
        ),
        _ => Err(DecodeError::invalid()),
    }
}

fn emit_fraction_format_canonical(
    output: &mut Vec<u8>,
    write: FractionFormatWrite,
) -> Result<(), DecodeError> {
    validate_fraction_format_write(write)?;
    emit_fraction_format_field(output, 0, write)?;
    emit_fraction_format_field(output, 1, write)?;
    Ok(())
}

fn emit_fraction_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    layout: FractionFormatLayout,
    write: FractionFormatWrite,
) -> Result<(), DecodeError> {
    validate_fraction_format_write(write)?;
    let mut ordered = [
        layout.spans[0].ok_or_else(DecodeError::invalid)?,
        layout.spans[1].ok_or_else(DecodeError::invalid)?,
    ];
    for index in 1..ordered.len() {
        let mut cursor = index;
        while cursor > 0 && ordered[cursor].start < ordered[cursor - 1].start {
            ordered.swap(cursor, cursor - 1);
            cursor -= 1;
        }
    }
    let mut source_offset = 0usize;
    for span in ordered {
        output.extend_from_slice(
            source
                .get(source_offset..span.start)
                .ok_or_else(DecodeError::invalid)?,
        );
        emit_fraction_format_field(output, span.slot, write)?;
        source_offset = span.end;
    }
    output.extend_from_slice(
        source
            .get(source_offset..)
            .ok_or_else(DecodeError::invalid)?,
    );
    Ok(())
}

fn verify_fraction_format_candidate(
    source: &[u8],
    write: FractionFormatWrite,
    options: DecodeOptions,
    layout: FractionFormatLayout,
    requirements: RewriteExecutionRequirements,
    candidate_work: usize,
) -> Result<(), DecodeError> {
    let (snapshot, report) = decode_fraction_format_with_report(source, options)?;
    if report.fields() != requirements.fields()
        || report.work_bytes() != candidate_work
        || report.max_depth() != layout.max_depth
        || FractionFormatWrite::from_snapshot(snapshot) != write
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct TextFieldSpan {
    end: usize,
}

#[derive(Debug, Clone, Copy)]
struct TextFormatLayout {
    fields: usize,
    max_depth: u32,
    span: Option<TextFieldSpan>,
}

fn scan_text_format<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(TextFormatSnapshot<'source>, TextFormatLayout, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let mut format_type = None;
    let mut span = None;
    let mut offset = 0usize;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        match field.number {
            FORMAT_TYPE_FIELD => {
                if field.wire != 0 || span.is_some() {
                    return Err(DecodeError::invalid());
                }
                format_type =
                    Some(u32::try_from(field.known_varint()?).map_err(|_| DecodeError::invalid())?);
                span = Some(TextFieldSpan { end: field.end });
            },
            number if number <= NUMBER_FORMAT_MAX_KNOWN_FIELD => {
                // Every other TSK.FormatStructArchive field belongs to a
                // different display family. Treating one as opaque would let
                // a caller publish another format through the Text route.
                return Err(DecodeError::invalid());
            },
            _ => {},
        }
        offset = field.end;
    }
    let format_type = format_type.ok_or_else(DecodeError::invalid)?;
    if format_type != NATIVE_TEXT_FORMAT_TYPE {
        return Err(DecodeError::invalid());
    }
    let snapshot = TextFormatSnapshot {
        source,
        format_type,
    };
    buffa_text_format_parity(source, snapshot, &mut budget)?;
    let report = budget.finish(0);
    let layout = TextFormatLayout {
        fields: report.fields(),
        max_depth: report.max_depth(),
        span,
    };
    Ok((snapshot, layout, report))
}

fn buffa_text_format_parity(
    source: &[u8],
    snapshot: TextFormatSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let options = budget.options;
    let view: text_projection::FormatStructArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.format_type != Some(snapshot.format_type) {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())?;
    Ok(())
}

fn validate_text_format_write(_write: TextFormatWrite) -> Result<(), DecodeError> {
    Ok(())
}

fn text_format_canonical_output_len() -> Result<usize, DecodeError> {
    number_format_field_len(FORMAT_TYPE_FIELD, u64::from(NATIVE_TEXT_FORMAT_TYPE))
}

fn emit_text_format_canonical(output: &mut Vec<u8>) -> Result<(), DecodeError> {
    emit_varint_field(
        output,
        FORMAT_TYPE_FIELD,
        u64::from(NATIVE_TEXT_FORMAT_TYPE),
    )
}

fn emit_text_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    layout: TextFormatLayout,
) -> Result<(), DecodeError> {
    let span = layout.span.ok_or_else(DecodeError::invalid)?;
    output.extend_from_slice(source.get(..span.end).ok_or_else(DecodeError::invalid)?);
    output.extend_from_slice(source.get(span.end..).ok_or_else(DecodeError::invalid)?);
    Ok(())
}

fn verify_text_format_candidate(
    source: &[u8],
    write: TextFormatWrite,
    options: DecodeOptions,
    layout: TextFormatLayout,
    candidate_work: usize,
) -> Result<(), DecodeError> {
    let (snapshot, report) = decode_text_format_with_report(source, options)?;
    if report.fields() != layout.fields
        || report.work_bytes() != candidate_work
        || report.max_depth() != layout.max_depth
        || TextFormatWrite::from_snapshot(snapshot) != write
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct DecimalFieldSpan {
    number: u32,
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy)]
struct DecimalFormatLayout {
    fields: usize,
    max_depth: u32,
    spans: [Option<DecimalFieldSpan>; 4],
}

#[derive(Debug, Clone, Copy)]
struct DecimalFormatScanState {
    format_type: Option<u32>,
    decimal_places: Option<u32>,
    negative_style: Option<u32>,
    show_thousands_separator: Option<bool>,
    spans: [Option<DecimalFieldSpan>; 4],
}

impl DecimalFormatScanState {
    const fn new() -> Self {
        Self {
            format_type: None,
            decimal_places: None,
            negative_style: None,
            show_thousands_separator: None,
            spans: [None, None, None, None],
        }
    }
}

fn scan_decimal_format<'source>(
    source: &'source [u8],
    expected_format_type: u32,
    options: DecodeOptions,
) -> Result<
    (
        NumberFormatSnapshot<'source>,
        DecimalFormatLayout,
        DecodeReport,
    ),
    DecodeError,
> {
    validate_decimal_format_type(expected_format_type)?;
    let mut budget = Budget::new(source, options)?;
    let mut state = DecimalFormatScanState::new();
    let end = scan_decimal_message(source, 0, 0, None, &mut budget, &mut state)?;
    if end != source.len() {
        return Err(DecodeError::invalid());
    }
    let format_type = state.format_type.ok_or_else(DecodeError::invalid)?;
    let decimal_places = state.decimal_places.ok_or_else(DecodeError::invalid)?;
    let negative_style = state.negative_style.ok_or_else(DecodeError::invalid)?;
    let show_thousands_separator = state
        .show_thousands_separator
        .ok_or_else(DecodeError::invalid)?;
    if format_type != expected_format_type
        || (decimal_places != NATIVE_AUTOMATIC_DECIMAL_PLACES
            && decimal_places > MAX_NUMBER_DECIMAL_PLACES)
        || negative_style > 3
        || (expected_format_type == NATIVE_SCIENTIFIC_FORMAT_TYPE
            && (decimal_places == NATIVE_AUTOMATIC_DECIMAL_PLACES
                || negative_style != NATIVE_SCIENTIFIC_NEGATIVE_STYLE
                || show_thousands_separator != NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR))
    {
        return Err(DecodeError::invalid());
    }

    let snapshot = NumberFormatSnapshot {
        source,
        format_type,
        decimal_places,
        negative_style,
        show_thousands_separator,
    };
    buffa_number_format_parity(source, snapshot, &mut budget)?;
    let report = budget.finish(0);
    let layout = DecimalFormatLayout {
        fields: report.fields(),
        max_depth: report.max_depth(),
        spans: state.spans,
    };
    Ok((snapshot, layout, report))
}

fn buffa_number_format_parity(
    source: &[u8],
    snapshot: NumberFormatSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let options = budget.options;
    let view: projection::FormatStructArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.format_type != Some(snapshot.format_type)
        || view.decimal_places != Some(snapshot.decimal_places)
        || view.negative_style != Some(snapshot.negative_style)
        || view.show_thousands_separator != Some(snapshot.show_thousands_separator)
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())?;
    Ok(())
}

fn scan_decimal_message(
    source: &[u8],
    mut cursor: usize,
    depth: u32,
    end_group: Option<u32>,
    budget: &mut Budget,
    state: &mut DecimalFormatScanState,
) -> Result<usize, DecodeError> {
    while cursor < source.len() {
        let start = cursor;
        let (key, key_len) = read_varint(source, cursor)?;
        let number = u32::try_from(key >> 3).map_err(|_| DecodeError::invalid())?;
        let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid())?;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return Err(DecodeError::invalid());
        }
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            if end_group != Some(number) {
                return Err(DecodeError::invalid());
            }
            budget.field(
                cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?,
                depth,
            )?;
            return Ok(cursor);
        }
        let start_tag_end = cursor;
        match wire {
            0 => {
                let value_len = if depth == 0 && number_format_slot(number).is_some() {
                    read_varint(source, cursor)?.1
                } else {
                    // Unknown scalar values remain source-authoritative. The
                    // field key is canonical, while an overlong value spelling
                    // is retained byte-for-byte instead of normalized.
                    read_varint_relaxed(source, cursor)?.1
                };
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
            },
            1 => {
                cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?;
                if cursor > source.len() {
                    return Err(DecodeError::invalid());
                }
            },
            2 => {
                let (length, length_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(length_len)
                    .ok_or_else(DecodeError::invalid)?;
                let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
                cursor = cursor
                    .checked_add(length)
                    .ok_or_else(DecodeError::invalid)?;
                if cursor > source.len() {
                    return Err(DecodeError::invalid());
                }
            },
            3 => {
                if depth >= budget.options.recursion_limit {
                    return Err(DecodeError::limited(DecodeLimit::Nesting {
                        observed: depth.saturating_add(1),
                        maximum: budget.options.recursion_limit,
                    }));
                }
                cursor = scan_decimal_message(
                    source,
                    cursor,
                    depth.saturating_add(1),
                    Some(number),
                    budget,
                    state,
                )?;
            },
            5 => {
                cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?;
                if cursor > source.len() {
                    return Err(DecodeError::invalid());
                }
            },
            _ => return Err(DecodeError::invalid()),
        }
        // A group start is one wire record (its key) followed by nested
        // records.  Charge each byte exactly once: nested records account for
        // the group body, while this record accounts only for its start tag.
        let record_len = if wire == 3 {
            start_tag_end
                .checked_sub(start)
                .ok_or_else(DecodeError::invalid)?
        } else {
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?
        };
        budget.field(record_len, depth)?;
        if depth == 0 {
            inspect_decimal_root_field(source, start, cursor, number, wire, state)?;
        }
    }
    if end_group.is_some() {
        return Err(DecodeError::invalid());
    }
    Ok(cursor)
}

fn inspect_decimal_root_field(
    source: &[u8],
    start: usize,
    end: usize,
    number: u32,
    wire: u8,
    state: &mut DecimalFormatScanState,
) -> Result<(), DecodeError> {
    let Some(slot) = number_format_slot(number) else {
        // Fields 3..45 are known `FormatStructArchive` fields with a
        // different shape/meaning.  Treating one as opaque would let a
        // caller accidentally publish a non-number format through this API.
        if number <= NUMBER_FORMAT_MAX_KNOWN_FIELD {
            return Err(DecodeError::invalid());
        }
        return Ok(());
    };
    if wire != 0 || state.spans[slot].is_some() {
        return Err(DecodeError::invalid());
    }
    let (_, key_len) = read_varint(source, start)?;
    let value_offset = start
        .checked_add(key_len)
        .ok_or_else(DecodeError::invalid)?;
    let (value, _) = read_varint(source, value_offset)?;
    let span = DecimalFieldSpan { number, start, end };
    state.spans[slot] = Some(span);
    match number {
        NUMBER_FORMAT_TYPE_FIELD => {
            state.format_type = Some(u32::try_from(value).map_err(|_| DecodeError::invalid())?)
        },
        NUMBER_FORMAT_DECIMAL_PLACES_FIELD => {
            state.decimal_places = Some(u32::try_from(value).map_err(|_| DecodeError::invalid())?)
        },
        NUMBER_FORMAT_NEGATIVE_STYLE_FIELD => {
            state.negative_style = Some(u32::try_from(value).map_err(|_| DecodeError::invalid())?)
        },
        NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => {
            if value > 1 {
                return Err(DecodeError::invalid());
            }
            state.show_thousands_separator = Some(value != 0);
        },
        _ => unreachable!("number format slot covers all selected fields"),
    }
    Ok(())
}

const fn number_format_slot(number: u32) -> Option<usize> {
    match number {
        NUMBER_FORMAT_TYPE_FIELD => Some(0),
        NUMBER_FORMAT_DECIMAL_PLACES_FIELD => Some(1),
        NUMBER_FORMAT_NEGATIVE_STYLE_FIELD => Some(2),
        NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => Some(3),
        _ => None,
    }
}

fn validate_decimal_format_type(format_type: u32) -> Result<(), DecodeError> {
    if matches!(
        format_type,
        NATIVE_NUMBER_FORMAT_TYPE | NATIVE_PERCENTAGE_FORMAT_TYPE | NATIVE_SCIENTIFIC_FORMAT_TYPE
    ) {
        Ok(())
    } else {
        Err(DecodeError::invalid())
    }
}

fn validate_decimal_format_write(
    write: NumberFormatWrite,
    format_type: u32,
) -> Result<(), DecodeError> {
    validate_decimal_format_type(format_type)?;
    if (write.decimal_places != NATIVE_AUTOMATIC_DECIMAL_PLACES
        && write.decimal_places > MAX_NUMBER_DECIMAL_PLACES)
        || write.negative_style > 3
        || (format_type == NATIVE_SCIENTIFIC_FORMAT_TYPE
            && (write.decimal_places == NATIVE_AUTOMATIC_DECIMAL_PLACES
                || write.negative_style != NATIVE_SCIENTIFIC_NEGATIVE_STYLE
                || write.show_thousands_separator != NATIVE_SCIENTIFIC_SHOW_THOUSANDS_SEPARATOR))
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn number_format_field_len(number: u32, value: u64) -> Result<usize, DecodeError> {
    encoded_varint_len(u64::from(number) << 3)
        .checked_add(encoded_varint_len(value))
        .ok_or_else(DecodeError::invalid)
}

fn decimal_format_rewrite_output_len(
    source: &[u8],
    layout: DecimalFormatLayout,
    write: NumberFormatWrite,
    format_type: u32,
) -> Result<usize, DecodeError> {
    let mut old = 0usize;
    for span in layout.spans.into_iter() {
        let span = span.ok_or_else(DecodeError::invalid)?;
        old = old
            .checked_add(
                span.end
                    .checked_sub(span.start)
                    .ok_or_else(DecodeError::invalid)?,
            )
            .ok_or_else(DecodeError::invalid)?;
    }
    let new = decimal_format_canonical_output_len(write, format_type)?;
    source
        .len()
        .checked_sub(old)
        .and_then(|length| length.checked_add(new))
        .ok_or_else(DecodeError::invalid)
}

fn decimal_format_canonical_output_len(
    write: NumberFormatWrite,
    format_type: u32,
) -> Result<usize, DecodeError> {
    validate_decimal_format_type(format_type)?;
    let type_len = number_format_field_len(NUMBER_FORMAT_TYPE_FIELD, u64::from(format_type))?;
    let decimal_len = number_format_field_len(
        NUMBER_FORMAT_DECIMAL_PLACES_FIELD,
        u64::from(write.decimal_places),
    )?;
    let negative_len = number_format_field_len(
        NUMBER_FORMAT_NEGATIVE_STYLE_FIELD,
        u64::from(write.negative_style),
    )?;
    let thousands_len = number_format_field_len(
        NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
        u64::from(write.show_thousands_separator),
    )?;
    type_len
        .checked_add(decimal_len)
        .and_then(|length| length.checked_add(negative_len))
        .and_then(|length| length.checked_add(thousands_len))
        .ok_or_else(DecodeError::invalid)
}

fn emit_decimal_format_canonical(
    output: &mut Vec<u8>,
    write: NumberFormatWrite,
    format_type: u32,
) -> Result<(), DecodeError> {
    validate_decimal_format_type(format_type)?;
    emit_varint_field(output, NUMBER_FORMAT_TYPE_FIELD, u64::from(format_type))?;
    emit_varint_field(
        output,
        NUMBER_FORMAT_DECIMAL_PLACES_FIELD,
        u64::from(write.decimal_places),
    )?;
    emit_varint_field(
        output,
        NUMBER_FORMAT_NEGATIVE_STYLE_FIELD,
        u64::from(write.negative_style),
    )?;
    emit_varint_field(
        output,
        NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
        u64::from(write.show_thousands_separator),
    )?;
    Ok(())
}

fn emit_decimal_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    layout: DecimalFormatLayout,
    write: NumberFormatWrite,
    format_type: u32,
) -> Result<(), DecodeError> {
    validate_decimal_format_type(format_type)?;
    let mut ordered = [
        layout.spans[0].ok_or_else(DecodeError::invalid)?,
        layout.spans[1].ok_or_else(DecodeError::invalid)?,
        layout.spans[2].ok_or_else(DecodeError::invalid)?,
        layout.spans[3].ok_or_else(DecodeError::invalid)?,
    ];
    for index in 1..ordered.len() {
        let mut cursor = index;
        while cursor > 0 && ordered[cursor].start < ordered[cursor - 1].start {
            ordered.swap(cursor, cursor - 1);
            cursor -= 1;
        }
    }
    let mut source_offset = 0usize;
    for span in ordered {
        output.extend_from_slice(
            source
                .get(source_offset..span.start)
                .ok_or_else(DecodeError::invalid)?,
        );
        let value = match span.number {
            NUMBER_FORMAT_TYPE_FIELD => u64::from(format_type),
            NUMBER_FORMAT_DECIMAL_PLACES_FIELD => u64::from(write.decimal_places),
            NUMBER_FORMAT_NEGATIVE_STYLE_FIELD => u64::from(write.negative_style),
            NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => {
                u64::from(write.show_thousands_separator)
            },
            _ => return Err(DecodeError::invalid()),
        };
        emit_varint_field(output, span.number, value)?;
        source_offset = span.end;
    }
    output.extend_from_slice(
        source
            .get(source_offset..)
            .ok_or_else(DecodeError::invalid)?,
    );
    Ok(())
}

fn verify_decimal_format_candidate(
    source: &[u8],
    write: NumberFormatWrite,
    format_type: u32,
    options: DecodeOptions,
    layout: DecimalFormatLayout,
) -> Result<(), DecodeError> {
    let (snapshot, report) = decode_decimal_format_with_report(source, format_type, options)?;
    let expected_work = source
        .len()
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields() != layout.fields
        || report.work_bytes() != expected_work
        || report.max_depth() != layout.max_depth
        || NumberFormatWrite::from_snapshot(snapshot) != write
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn buffa_cell_spec_parity(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let options = budget.options;
    let _: projection::CellSpecArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    budget.work(source.len())?;
    Ok(())
}

fn buffa_cell_value_parity(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let options = budget.options;
    let _: projection::CellValueArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    budget.work(source.len())?;
    Ok(())
}

fn buffa_string_value_parity(source: &[u8], budget: &mut Budget) -> Result<(), DecodeError> {
    let options = budget.options;
    let _: projection::StringCellValueArchiveLazyView<'_> = BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(0)
        .with_recursion_limit(options.recursion_limit)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    budget.work(source.len())?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    raw: &'source [u8],
    payload: Option<&'source [u8]>,
    varint: Option<u64>,
    varint_canonical: bool,
    nested_fields: usize,
    nested_work_bytes: usize,
    nested_max_depth: u32,
    end: usize,
}

impl Field<'_> {
    fn known_varint(self) -> Result<u64, DecodeError> {
        if self.wire != 0 || !self.varint_canonical {
            return Err(DecodeError::invalid());
        }
        self.varint.ok_or_else(DecodeError::invalid)
    }

    fn fixed64(self) -> Result<u64, DecodeError> {
        if self.wire != 1 || self.raw.len() < 8 {
            return Err(DecodeError::invalid());
        }
        let start = self
            .raw
            .len()
            .checked_sub(8)
            .ok_or_else(DecodeError::invalid)?;
        let bytes = self.raw.get(start..).ok_or_else(DecodeError::invalid)?;
        let bytes: [u8; 8] = bytes.try_into().map_err(|_| DecodeError::invalid())?;
        Ok(u64::from_le_bytes(bytes))
    }
}

#[derive(Debug, Clone, Copy)]
struct Budget {
    options: DecodeOptions,
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    unknown_fields: bool,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_message_bytes =
            usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError::invalid())?;
        if options.max_message_bytes > hard_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::InputBytes {
                observed: options.max_message_bytes,
                maximum: hard_message_bytes,
            }));
        }
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::InputBytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION.min(options.recursion_limit.max(1)),
            }));
        }
        Ok(Self {
            options,
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            items: 0,
            text_bytes: 0,
            unknown_fields: false,
        })
    }

    fn field(&mut self, amount: usize, depth: u32) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.options.max_fields,
            })
        })?;
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        self.max_depth = self.max_depth.max(depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }

    fn reference(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.references = self.references.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::References {
                observed: usize::MAX,
                maximum: self.options.max_references,
            })
        })?;
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        if self.references > self.options.max_references {
            return Err(DecodeError::limited(DecodeLimit::References {
                observed: self.references,
                maximum: self.options.max_references,
            }));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        Ok(())
    }

    fn item(&mut self, text: usize) -> Result<(), DecodeError> {
        self.items = self.items.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Items {
                observed: usize::MAX,
                maximum: self.options.max_items,
            })
        })?;
        self.text_bytes = self.text_bytes.checked_add(text).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Text {
                observed: usize::MAX,
                maximum: self.options.max_text_bytes,
            })
        })?;
        if self.items > self.options.max_items {
            return Err(DecodeError::limited(DecodeLimit::Items {
                observed: self.items,
                maximum: self.options.max_items,
            }));
        }
        if self.text_bytes > self.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: self.text_bytes,
                maximum: self.options.max_text_bytes,
            }));
        }
        Ok(())
    }

    fn text(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.text_bytes = self.text_bytes.checked_add(amount).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Text {
                observed: usize::MAX,
                maximum: self.options.max_text_bytes,
            })
        })?;
        if self.text_bytes > self.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: self.text_bytes,
                maximum: self.options.max_text_bytes,
            }));
        }
        Ok(())
    }

    fn mark_unknown(&mut self) {
        self.unknown_fields = true;
    }

    fn nested_fields(
        &mut self,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
    ) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(fields).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.options.max_fields,
            })
        })?;
        self.work_bytes = self.work_bytes.checked_add(work_bytes).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: self.options.max_work_bytes,
            })
        })?;
        self.max_depth = self.max_depth.max(max_depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        if self.work_bytes > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work_bytes,
                maximum: self.options.max_work_bytes,
            }));
        }
        if max_depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: max_depth,
                maximum: self.options.recursion_limit,
            }));
        }
        Ok(())
    }

    fn finish(self, output_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            items: self.items,
            text_bytes: self.text_bytes,
            allocations: 0,
            retained_bytes: output_bytes,
            scratch_bytes: 0,
        }
    }
}

fn parse_popup_model_with_visitor<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    mut visitor: Option<&mut dyn PopUpMenuVisitor>,
) -> Result<PopUpMenuModelSnapshot<'source>, DecodeError> {
    let mut offset = 0usize;
    let mut first_nil = false;
    let mut saw_item = false;
    let mut text_bytes = 0usize;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            POPUP_DEPRECATED_ITEM_FIELD => return Err(DecodeError::invalid()),
            POPUP_ITEM_FIELD => {
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                let (value, is_nil, value_text_bytes) = parse_cell_value(payload, budget, 1)?;
                if !saw_item {
                    if !is_nil {
                        return Err(DecodeError::invalid());
                    }
                    first_nil = true;
                    saw_item = true;
                } else if is_nil {
                    return Err(DecodeError::invalid());
                } else {
                    text_bytes = text_bytes
                        .checked_add(value_text_bytes)
                        .ok_or_else(DecodeError::invalid)?;
                    budget.item(value_text_bytes)?;
                    if value.is_none() {
                        return Err(DecodeError::invalid());
                    }
                    if let Some(visitor) = visitor.as_deref_mut() {
                        visitor.visit_item(PopUpMenuItem {
                            value: value.ok_or_else(DecodeError::invalid)?,
                        })?;
                    }
                }
            },
            _ => budget.mark_unknown(),
        }
    }
    if !saw_item || !first_nil || budget.items == 0 {
        return Err(DecodeError::invalid());
    }
    Ok(PopUpMenuModelSnapshot {
        source,
        first_nil,
        item_count: budget.items,
        text_bytes,
    })
}

fn parse_cell_value<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(Option<&'source str>, bool, usize), DecodeError> {
    let mut offset = 0usize;
    let mut value_type = None;
    let mut string_payload = None;
    let mut wrapper_seen = [false; 5];
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            CELL_VALUE_TYPE_FIELD => {
                if field.wire != 0 || value_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.known_varint()?;
                if !matches!(value, 1..=5) {
                    return Err(DecodeError::invalid());
                }
                value_type = Some(value);
            },
            2..=6 => {
                if field.wire != 2 {
                    return Err(DecodeError::invalid());
                }
                let slot = usize::try_from(field.number - 2).map_err(|_| DecodeError::invalid())?;
                if wrapper_seen[slot] {
                    return Err(DecodeError::invalid());
                }
                wrapper_seen[slot] = true;
                if field.number == CELL_VALUE_STRING_FIELD {
                    string_payload = field.payload;
                }
            },
            _ => budget.mark_unknown(),
        }
    }
    let value_type = value_type.ok_or_else(DecodeError::invalid)?;
    match value_type {
        NIL_TYPE if wrapper_seen.iter().all(|seen| !seen) => {
            if !budget.unknown_fields {
                buffa_cell_value_parity(source, budget)?;
            }
            Ok((None, true, 0))
        },
        STRING_TYPE
            if wrapper_seen[3]
                && !wrapper_seen[..3].iter().any(|seen| *seen)
                && !wrapper_seen[4] =>
        {
            let payload = string_payload.ok_or_else(DecodeError::invalid)?;
            let (value, text_bytes) = parse_string_value(payload, budget, depth + 1)?;
            if !budget.unknown_fields {
                buffa_cell_value_parity(source, budget)?;
                buffa_string_value_parity(payload, budget)?;
            }
            Ok((Some(value), false, text_bytes))
        },
        _ => Err(DecodeError::invalid()),
    }
}

fn parse_string_value<'source>(
    source: &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<(&'source str, usize), DecodeError> {
    let mut offset = 0usize;
    let mut value = None;
    let mut format = None;
    let mut explicit = None;
    let mut regex = None;
    let mut case_sensitive = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            STRING_VALUE_FIELD => {
                if field.wire != 2 || value.is_some() {
                    return Err(DecodeError::invalid());
                }
                let bytes = field.payload.ok_or_else(DecodeError::invalid)?;
                let text = str::from_utf8(bytes).map_err(|_| DecodeError::invalid())?;
                if text.chars().any(char::is_control) {
                    return Err(DecodeError::invalid());
                }
                value = Some((text, bytes.len()));
            },
            STRING_FORMAT_FIELD => {
                if field.wire != 2 || format.is_some() {
                    return Err(DecodeError::invalid());
                }
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                parse_text_format(payload, budget, depth + 1)?;
                format = Some(());
            },
            STRING_IMPLICIT_FIELD => return Err(DecodeError::invalid()),
            STRING_EXPLICIT_FIELD => set_bool(&mut explicit, field, false)?,
            STRING_REGEX_FIELD => set_bool(&mut regex, field, false)?,
            STRING_CASE_SENSITIVE_REGEX_FIELD => set_bool(&mut case_sensitive, field, false)?,
            _ => budget.mark_unknown(),
        }
    }
    let (text, text_bytes) = value.ok_or_else(DecodeError::invalid)?;
    if format.is_none()
        || explicit != Some(false)
        || regex != Some(false)
        || case_sensitive != Some(false)
    {
        return Err(DecodeError::invalid());
    }
    Ok((text, text_bytes))
}

fn parse_text_format(source: &[u8], budget: &mut Budget, depth: u32) -> Result<(), DecodeError> {
    let mut offset = 0usize;
    let mut format_type = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            FORMAT_TYPE_FIELD => {
                if field.wire != 0 || format_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                format_type = Some(field.known_varint()?);
            },
            _ => budget.mark_unknown(),
        }
    }
    if format_type != Some(FORMAT_TYPE_TEXT) {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn parse_cell_spec<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<CellSpecSnapshot<'source>, DecodeError> {
    let mut offset = 0usize;
    let mut interaction = None;
    let mut model = None;
    let mut starts = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            CELL_SPEC_INTERACTION_FIELD => {
                if field.wire != 0 || interaction.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.known_varint()?;
                interaction = Some(u32::try_from(value).map_err(|_| DecodeError::invalid())?);
            },
            CELL_SPEC_FORMULA_FIELD | 3..=5 | CELL_SPEC_DEPRECATED_LABEL_FIELD => {
                return Err(DecodeError::invalid());
            },
            CELL_SPEC_MODEL_FIELD => {
                if field.wire != 2 || model.is_some() {
                    return Err(DecodeError::invalid());
                }
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                model = Some(parse_reference(payload, budget, 1)?);
            },
            CELL_SPEC_FIRST_FIELD => set_bool(&mut starts, field, true)?,
            _ => budget.mark_unknown(),
        }
    }
    if interaction != Some(POPUP_INTERACTION_TYPE) {
        return Err(DecodeError::invalid());
    }
    Ok(CellSpecSnapshot {
        source,
        interaction_type: interaction.ok_or_else(DecodeError::invalid)?,
        popup_model: model.ok_or_else(DecodeError::invalid)?,
        starts_with_first: starts.ok_or_else(DecodeError::invalid)?,
    })
}

fn parse_control_cell_spec<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<ControlCellSpecSnapshot<'source>, DecodeError> {
    let mut offset = 0usize;
    let mut interaction = None;
    let mut minimum = None;
    let mut maximum = None;
    let mut increment = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            CELL_SPEC_INTERACTION_FIELD => {
                if field.wire != 0 || interaction.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.known_varint()?;
                if !matches!(value, 4 | 5 | 6 | 8) {
                    return Err(DecodeError::invalid());
                }
                interaction = Some(u32::try_from(value).map_err(|_| DecodeError::invalid())?);
            },
            3..=5 => {
                if field.wire != 1 {
                    return Err(DecodeError::invalid());
                }
                let value = field.fixed64()?.to_le_bytes();
                let value = f64::from_le_bytes(value);
                if !value.is_finite() {
                    return Err(DecodeError::invalid());
                }
                let target = match field.number {
                    3 => &mut minimum,
                    4 => &mut maximum,
                    5 => &mut increment,
                    _ => unreachable!("matched control range field"),
                };
                if target.is_some() {
                    return Err(DecodeError::invalid());
                }
                *target = Some(value);
            },
            // Formula, popup model, selection, and deprecated label fields
            // are all outside the four strict control projections.
            CELL_SPEC_FORMULA_FIELD
            | CELL_SPEC_MODEL_FIELD
            | CELL_SPEC_FIRST_FIELD
            | CELL_SPEC_DEPRECATED_LABEL_FIELD => return Err(DecodeError::invalid()),
            _ => budget.mark_unknown(),
        }
    }
    let interaction = interaction.ok_or_else(DecodeError::invalid)?;
    let is_range = matches!(
        interaction,
        SLIDER_INTERACTION_TYPE | STEPPER_INTERACTION_TYPE | STAR_RATING_INTERACTION_TYPE
    );
    if is_range {
        let min = minimum.ok_or_else(DecodeError::invalid)?;
        let max = maximum.ok_or_else(DecodeError::invalid)?;
        let inc = increment.ok_or_else(DecodeError::invalid)?;
        if min >= max || inc <= 0.0 || !((max - min) / inc).is_finite() {
            return Err(DecodeError::invalid());
        }
        if interaction == STAR_RATING_INTERACTION_TYPE && (min != 0.0 || max != 5.0 || inc != 1.0) {
            return Err(DecodeError::invalid());
        }
    } else if minimum.is_some() || maximum.is_some() || increment.is_some() {
        return Err(DecodeError::invalid());
    }
    Ok(ControlCellSpecSnapshot {
        source,
        interaction_type: interaction,
        range_control_min: minimum,
        range_control_max: maximum,
        range_control_inc: increment,
    })
}

fn parse_control_format<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<ControlFormatSnapshot<'source>, DecodeError> {
    let mut offset = 0usize;
    let mut format_type = None;
    let mut decimal_places = None;
    let mut currency_code = None;
    let mut negative_style = None;
    let mut show_thousands_separator = None;
    let mut use_accounting_style = None;
    let mut duration_style = None;
    let mut base = None;
    let mut base_places = None;
    let mut base_use_minus_sign = None;
    let mut fraction_accuracy = None;
    let mut suppress_date_format = None;
    let mut suppress_time_format = None;
    let mut date_time_format = None;
    let mut duration_unit_largest = None;
    let mut duration_unit_smallest = None;
    let mut control_minimum = None;
    let mut control_maximum = None;
    let mut control_increment = None;
    let mut control_format_type = None;
    let mut slider_orientation = None;
    let mut slider_position = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, 0, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), 0)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            1 => {
                if field.wire != 0 || format_type.is_some() {
                    return Err(DecodeError::invalid());
                }
                let value = field.known_varint()?;
                let value = u32::try_from(value).map_err(|_| DecodeError::invalid())?;
                if !matches!(value, 256..=263 | 267..=269) {
                    return Err(DecodeError::invalid());
                }
                format_type = Some(value);
            },
            2 | 4 | 7 | 8 | 9 | 11 | 15 | 16 => {
                if field.wire != 0 || !field.varint_canonical {
                    return Err(DecodeError::invalid());
                }
                let value =
                    u32::try_from(field.known_varint()?).map_err(|_| DecodeError::invalid())?;
                let target = match field.number {
                    FORMAT_DECIMAL_PLACES_FIELD => &mut decimal_places,
                    FORMAT_NEGATIVE_STYLE_FIELD => &mut negative_style,
                    FORMAT_DURATION_STYLE_FIELD => &mut duration_style,
                    FORMAT_BASE_FIELD => &mut base,
                    FORMAT_BASE_PLACES_FIELD => &mut base_places,
                    FORMAT_FRACTION_ACCURACY_FIELD => &mut fraction_accuracy,
                    FORMAT_DURATION_UNIT_LARGEST_FIELD => &mut duration_unit_largest,
                    FORMAT_DURATION_UNIT_SMALLEST_FIELD => &mut duration_unit_smallest,
                    _ => unreachable!("matched display-format integer field"),
                };
                if target.is_some() {
                    return Err(DecodeError::invalid());
                }
                *target = Some(value);
            },
            3 | 14 => {
                if field.wire != 2 {
                    return Err(DecodeError::invalid());
                }
                let bytes = field.payload.ok_or_else(DecodeError::invalid)?;
                let text = str::from_utf8(bytes).map_err(|_| DecodeError::invalid())?;
                if text.chars().any(char::is_control) {
                    return Err(DecodeError::invalid());
                }
                budget.text(bytes.len())?;
                let target = match field.number {
                    FORMAT_CURRENCY_CODE_FIELD => &mut currency_code,
                    FORMAT_DATE_TIME_FORMAT_FIELD => &mut date_time_format,
                    _ => unreachable!("matched display-format string field"),
                };
                if target.is_some() {
                    return Err(DecodeError::invalid());
                }
                *target = Some(text);
            },
            5 | 6 | 10 | 12 | 13 => {
                let target = match field.number {
                    FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => &mut show_thousands_separator,
                    FORMAT_USE_ACCOUNTING_STYLE_FIELD => &mut use_accounting_style,
                    FORMAT_BASE_USE_MINUS_SIGN_FIELD => &mut base_use_minus_sign,
                    FORMAT_SUPPRESS_DATE_FORMAT_FIELD => &mut suppress_date_format,
                    FORMAT_SUPPRESS_TIME_FORMAT_FIELD => &mut suppress_time_format,
                    _ => unreachable!("matched display-format bool field"),
                };
                set_bool(target, field, false)?;
            },
            FORMAT_CONTROL_MINIMUM_FIELD..=FORMAT_CONTROL_INCREMENT_FIELD => {
                if field.wire != 1 {
                    return Err(DecodeError::invalid());
                }
                let value = f64::from_le_bytes(field.fixed64()?.to_le_bytes());
                if !value.is_finite() {
                    return Err(DecodeError::invalid());
                }
                let target = match field.number {
                    FORMAT_CONTROL_MINIMUM_FIELD => &mut control_minimum,
                    FORMAT_CONTROL_MAXIMUM_FIELD => &mut control_maximum,
                    FORMAT_CONTROL_INCREMENT_FIELD => &mut control_increment,
                    _ => unreachable!("matched control format field"),
                };
                if target.is_some() {
                    return Err(DecodeError::invalid());
                }
                *target = Some(value);
            },
            FORMAT_CONTROL_FORMAT_TYPE_FIELD..=FORMAT_SLIDER_POSITION_FIELD => {
                if field.wire != 0 || !field.varint_canonical {
                    return Err(DecodeError::invalid());
                }
                let target = match field.number {
                    FORMAT_CONTROL_FORMAT_TYPE_FIELD => &mut control_format_type,
                    FORMAT_SLIDER_ORIENTATION_FIELD => &mut slider_orientation,
                    FORMAT_SLIDER_POSITION_FIELD => &mut slider_position,
                    _ => unreachable!("matched control format field"),
                };
                if target.is_some() {
                    return Err(DecodeError::invalid());
                }
                *target =
                    Some(u32::try_from(field.known_varint()?).map_err(|_| DecodeError::invalid())?);
            },
            _ => budget.mark_unknown(),
        }
    }
    let format_type = format_type.ok_or_else(DecodeError::invalid)?;
    if matches!(format_type, 263 | 267)
        && (decimal_places.is_some()
            || currency_code.is_some()
            || negative_style.is_some()
            || show_thousands_separator.is_some()
            || use_accounting_style.is_some()
            || duration_style.is_some()
            || base.is_some()
            || base_places.is_some()
            || base_use_minus_sign.is_some()
            || fraction_accuracy.is_some()
            || suppress_date_format.is_some()
            || suppress_time_format.is_some()
            || date_time_format.is_some()
            || duration_unit_largest.is_some()
            || duration_unit_smallest.is_some()
            || control_minimum.is_some()
            || control_maximum.is_some()
            || control_increment.is_some()
            || control_format_type.is_some()
            || slider_orientation.is_some()
            || slider_position.is_some())
    {
        return Err(DecodeError::invalid());
    }
    if control_minimum.is_some() != control_maximum.is_some()
        || control_minimum.is_some() != control_increment.is_some()
    {
        return Err(DecodeError::invalid());
    }
    if let (Some(min), Some(max), Some(inc)) = (control_minimum, control_maximum, control_increment)
        && (min >= max || inc <= 0.0 || !((max - min) / inc).is_finite())
    {
        return Err(DecodeError::invalid());
    }
    Ok(ControlFormatSnapshot {
        source,
        format_type,
        decimal_places,
        currency_code,
        negative_style,
        show_thousands_separator,
        use_accounting_style,
        duration_style,
        base,
        base_places,
        base_use_minus_sign,
        fraction_accuracy,
        suppress_date_format,
        suppress_time_format,
        date_time_format,
        duration_unit_largest,
        duration_unit_smallest,
        control_minimum,
        control_maximum,
        control_increment,
        control_format_type,
        slider_orientation,
        slider_position,
    })
}

fn parse_reference(
    source: &[u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.reference(source.len())?;
    let mut offset = 0usize;
    let mut identifier = None;
    while offset < source.len() {
        let field = parse_one_field_limited(source, offset, depth, budget.options.recursion_limit)?;
        budget.field(field.raw.len(), depth)?;
        budget.nested_fields(
            field.nested_fields,
            field.nested_work_bytes,
            field.nested_max_depth,
        )?;
        offset = field.end;
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if field.wire != 0 || identifier.is_some() {
                    return Err(DecodeError::invalid());
                }
                identifier = Some(field.known_varint()?);
            },
            // These legacy fields are omitted by the canonical writer.  A
            // source-preserving CellSpec rewrite has no safe way to retain
            // their exact presence/framing, so reject them even when their
            // values are the benign defaults (zero/false).
            REFERENCE_TYPE_FIELD | REFERENCE_EXTERNAL_FIELD => {
                return Err(DecodeError::invalid());
            },
            _ => budget.mark_unknown(),
        }
    }
    let identifier = identifier
        .filter(|value| *value != 0)
        .ok_or_else(DecodeError::invalid)?;
    Ok(ReferenceSnapshot {
        identifier,
        deprecated_type: None,
        deprecated_is_external: None,
    })
}

fn set_bool(
    target: &mut Option<bool>,
    field: Field<'_>,
    expected_presence: bool,
) -> Result<(), DecodeError> {
    if field.wire != 0 || target.is_some() {
        return Err(DecodeError::invalid());
    }
    let value = field.known_varint()?;
    if value > 1 {
        return Err(DecodeError::invalid());
    }
    if expected_presence && value > 1 {
        return Err(DecodeError::invalid());
    }
    *target = Some(value != 0);
    Ok(())
}

fn parse_cell_value_for_iterator(source: &[u8]) -> Result<Option<&str>, DecodeError> {
    let mut budget = Budget::new(source, DecodeOptions::for_source(source))?;
    let (value, is_nil, _) = parse_cell_value(source, &mut budget, 1)?;
    if is_nil {
        return Ok(None);
    }
    value.map(Some).ok_or_else(DecodeError::invalid)
}

fn parse_one_field<'source>(
    source: &'source [u8],
    offset: usize,
    depth: u32,
) -> Result<Field<'source>, DecodeError> {
    parse_one_field_limited(source, offset, depth, MAX_RECURSION)
}

fn parse_one_field_limited<'source>(
    source: &'source [u8],
    offset: usize,
    depth: u32,
    recursion_limit: u32,
) -> Result<Field<'source>, DecodeError> {
    let start = offset;
    let (key, key_len) = read_varint(source, offset)?;
    let number = u32::try_from(key >> 3).map_err(|_| DecodeError::invalid())?;
    let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let mut cursor = offset
        .checked_add(key_len)
        .ok_or_else(DecodeError::invalid)?;
    let mut payload = None;
    let mut varint = None;
    let mut varint_canonical = true;
    match wire {
        0 => {
            let (value, length, canonical) = read_varint_relaxed(source, cursor)?;
            varint = Some(value);
            varint_canonical = canonical;
            cursor = cursor
                .checked_add(length)
                .ok_or_else(DecodeError::invalid)?;
        },
        1 => {
            cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?;
            if cursor > source.len() {
                return Err(DecodeError::invalid());
            }
        },
        2 => {
            let (length, length_bytes) = read_varint(source, cursor)?;
            cursor = cursor
                .checked_add(length_bytes)
                .ok_or_else(DecodeError::invalid)?;
            let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
            let end = cursor
                .checked_add(length)
                .ok_or_else(DecodeError::invalid)?;
            if end > source.len() {
                return Err(DecodeError::invalid());
            }
            payload = Some(&source[cursor..end]);
            cursor = end;
        },
        3 => {
            if depth >= recursion_limit {
                return Err(DecodeError::limited(DecodeLimit::Nesting {
                    observed: depth.saturating_add(1),
                    maximum: recursion_limit,
                }));
            }
            let group = skip_group(
                source,
                cursor,
                number,
                depth.saturating_add(1),
                recursion_limit,
            )?;
            cursor = group.end;
            return Ok(Field {
                number,
                wire,
                raw: &source[start..cursor],
                payload,
                varint,
                varint_canonical,
                nested_fields: group.fields,
                nested_work_bytes: group.work_bytes,
                nested_max_depth: group.max_depth,
                end: cursor,
            });
        },
        4 => return Err(DecodeError::invalid()),
        5 => {
            cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?;
            if cursor > source.len() {
                return Err(DecodeError::invalid());
            }
        },
        _ => return Err(DecodeError::invalid()),
    }
    Ok(Field {
        number,
        wire,
        raw: &source[start..cursor],
        payload,
        varint,
        varint_canonical,
        nested_fields: 0,
        nested_work_bytes: 0,
        nested_max_depth: 0,
        end: cursor,
    })
}

#[derive(Debug, Clone, Copy)]
struct GroupScan {
    end: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
}

fn skip_group(
    source: &[u8],
    mut cursor: usize,
    root_number: u32,
    depth: u32,
    recursion_limit: u32,
) -> Result<GroupScan, DecodeError> {
    let mut stack = [0u32; 64];
    let mut stack_len = 1usize;
    stack[0] = root_number;
    let mut fields = 0usize;
    let mut work_bytes = 0usize;
    let mut max_depth = depth;
    while cursor < source.len() {
        let field_start = cursor;
        let (key, key_len) = read_varint(source, cursor)?;
        let number = u32::try_from(key >> 3).map_err(|_| DecodeError::invalid())?;
        let wire = u8::try_from(key & 7).map_err(|_| DecodeError::invalid())?;
        if number == 0 || number > MAX_FIELD_NUMBER {
            return Err(DecodeError::invalid());
        }
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        match wire {
            0 => {
                let (_, length, _) = read_varint_relaxed(source, cursor)?;
                cursor = cursor
                    .checked_add(length)
                    .ok_or_else(DecodeError::invalid)?;
            },
            1 => cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?,
            2 => {
                let (length, length_bytes) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(length_bytes)
                    .ok_or_else(DecodeError::invalid)?;
                let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
                cursor = cursor
                    .checked_add(length)
                    .ok_or_else(DecodeError::invalid)?;
            },
            3 => {
                if stack_len >= stack.len()
                    || u32::try_from(stack_len).unwrap_or(u32::MAX) >= recursion_limit
                {
                    return Err(DecodeError::limited(DecodeLimit::Nesting {
                        observed: depth.saturating_add(stack_len as u32),
                        maximum: recursion_limit,
                    }));
                }
                stack[stack_len] = number;
                stack_len += 1;
            },
            4 => {
                if stack_len == 0 || stack[stack_len - 1] != number {
                    return Err(DecodeError::invalid());
                }
                stack_len -= 1;
                if stack_len == 0 {
                    let raw_len = cursor
                        .checked_sub(field_start)
                        .ok_or_else(DecodeError::invalid)?;
                    fields = fields.checked_add(1).ok_or_else(|| {
                        DecodeError::limited(DecodeLimit::Fields {
                            observed: usize::MAX,
                            maximum: usize::MAX,
                        })
                    })?;
                    work_bytes = work_bytes.checked_add(raw_len).ok_or_else(|| {
                        DecodeError::limited(DecodeLimit::Work {
                            observed: usize::MAX,
                            maximum: usize::MAX,
                        })
                    })?;
                    return Ok(GroupScan {
                        end: cursor,
                        fields,
                        work_bytes,
                        max_depth,
                    });
                }
            },
            5 => cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?,
            _ => return Err(DecodeError::invalid()),
        }
        if cursor > source.len() {
            return Err(DecodeError::invalid());
        }
        let raw_len = cursor
            .checked_sub(field_start)
            .ok_or_else(DecodeError::invalid)?;
        fields = fields.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
        work_bytes = work_bytes.checked_add(raw_len).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Work {
                observed: usize::MAX,
                maximum: usize::MAX,
            })
        })?;
        max_depth = max_depth.max(depth.saturating_add(stack_len as u32).saturating_sub(1));
    }
    Err(DecodeError::invalid())
}

fn read_varint(source: &[u8], offset: usize) -> Result<(u64, usize), DecodeError> {
    let (value, consumed, canonical) = read_varint_relaxed(source, offset)?;
    if !canonical {
        return Err(DecodeError::invalid());
    }
    Ok((value, consumed))
}

fn read_varint_relaxed(source: &[u8], offset: usize) -> Result<(u64, usize, bool), DecodeError> {
    let mut value = 0u64;
    let mut shift = 0u32;
    let mut index = offset;
    while index < source.len() && index - offset < 10 {
        let byte = source[index];
        let part = u64::from(byte & 0x7f);
        if shift == 63 && part > 1 {
            return Err(DecodeError::invalid());
        }
        value |= part.checked_shl(shift).ok_or_else(DecodeError::invalid)?;
        index += 1;
        if byte & 0x80 == 0 {
            let consumed = index - offset;
            return Ok((value, consumed, encoded_varint_len(value) == consumed));
        }
        shift += 7;
    }
    Err(DecodeError::invalid())
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

/// Exact preflight requirements for a prepared popup write.
///
/// The field/work/depth/item/reference values are the strict candidate-decode
/// values used by `execute` after emission; execution therefore cannot publish
/// a candidate whose wire shape diverges from the prepared projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionRequirements {
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
    pub const fn references(self) -> usize {
        self.references
    }
    #[must_use]
    pub const fn items(self) -> usize {
        self.items
    }
    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
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

/// Caller-provided execution ceilings for a prepared write.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    items: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Build exact limits from a prepared requirement set.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        Self {
            output_bytes: requirements.output_bytes,
            fields: requirements.fields,
            work_bytes: requirements.work_bytes,
            max_depth: requirements.max_depth,
            references: requirements.references,
            items: requirements.items,
            text_bytes: requirements.text_bytes,
            allocations: requirements.allocations,
            retained_bytes: requirements.retained_bytes,
            scratch_bytes: requirements.scratch_bytes,
        }
    }

    /// Start with independent outer limits set to their largest values.
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            references: usize::MAX,
            items: usize::MAX,
            text_bytes: usize::MAX,
            allocations: usize::MAX,
            retained_bytes: usize::MAX,
            scratch_bytes: usize::MAX,
        }
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
    pub const fn with_references(mut self, value: usize) -> Self {
        self.references = value;
        self
    }
    #[must_use]
    pub const fn with_items(mut self, value: usize) -> Self {
        self.items = value;
        self
    }
    #[must_use]
    pub const fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value;
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

/// Output of a prepared popup write. Candidates remain private to the package
/// transaction until locality and reopen checks complete.
#[derive(Debug, PartialEq, Eq)]
pub struct RewriteOutput {
    bytes: Vec<u8>,
    report: DecodeReport,
}

impl RewriteOutput {
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
    #[must_use]
    pub const fn report(&self) -> DecodeReport {
        self.report
    }
}

/// Prepared canonical `PopUpMenuModel` write. Preparation allocates no
/// candidate bytes and only borrows caller-owned item text.
#[derive(Debug, Clone, Copy)]
pub struct PreparedPopUpMenuModelWrite<'items> {
    items: &'items [&'items str],
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl<'items> PreparedPopUpMenuModelWrite<'items> {
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_popup_model(&mut bytes, self.items)?;
        verify_popup_model_candidate(&bytes, self.verify_options, self.requirements)?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical popup `CellSpecArchive` write.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCellSpecWrite {
    model_identifier: u64,
    starts_with_first: bool,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

/// Prepared canonical write for one checkbox, star-rating, slider, or
/// stepper `CellSpecArchive`.
#[derive(Debug, Clone, Copy)]
pub struct PreparedControlCellSpecWrite {
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedControlCellSpecWrite {
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_control_cell_spec(
            &mut bytes,
            self.interaction_type,
            self.range_control_min,
            self.range_control_max,
            self.range_control_inc,
        )?;
        verify_control_cell_spec_candidate(&bytes, self.verify_options, self.requirements)?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Archive-free control-format values for a canonical `FormatStructArchive`.
///
/// The fields mirror the native display-format scalar surface (TSK fields
/// 1--16).  Control range metadata lives in `CellSpecArchive` and is not part
/// of this write bundle.  Unknown source fields are intentionally not copied
/// by the creation-only canonical writer; source-preserving package rewrites
/// must retain the original payload and patch only selected fields.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ControlFormatWrite<'source> {
    format_type: u32,
    decimal_places: Option<u32>,
    currency_code: Option<&'source str>,
    negative_style: Option<u32>,
    show_thousands_separator: Option<bool>,
    use_accounting_style: Option<bool>,
    duration_style: Option<u32>,
    base: Option<u32>,
    base_places: Option<u32>,
    base_use_minus_sign: Option<bool>,
    fraction_accuracy: Option<u32>,
    suppress_date_format: Option<bool>,
    suppress_time_format: Option<bool>,
    date_time_format: Option<&'source str>,
    duration_unit_largest: Option<u32>,
    duration_unit_smallest: Option<u32>,
}

impl<'source> ControlFormatWrite<'source> {
    #[must_use]
    pub const fn new(format_type: u32) -> Self {
        Self {
            format_type,
            decimal_places: None,
            currency_code: None,
            negative_style: None,
            show_thousands_separator: None,
            use_accounting_style: None,
            duration_style: None,
            base: None,
            base_places: None,
            base_use_minus_sign: None,
            fraction_accuracy: None,
            suppress_date_format: None,
            suppress_time_format: None,
            date_time_format: None,
            duration_unit_largest: None,
            duration_unit_smallest: None,
        }
    }
    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }
    #[must_use]
    pub const fn with_decimal_places(mut self, value: u32) -> Self {
        self.decimal_places = Some(value);
        self
    }
    #[must_use]
    pub const fn with_currency_code(mut self, value: &'source str) -> Self {
        self.currency_code = Some(value);
        self
    }
    #[must_use]
    pub const fn with_negative_style(mut self, value: u32) -> Self {
        self.negative_style = Some(value);
        self
    }
    #[must_use]
    pub const fn with_show_thousands_separator(mut self, value: bool) -> Self {
        self.show_thousands_separator = Some(value);
        self
    }
    #[must_use]
    pub const fn with_use_accounting_style(mut self, value: bool) -> Self {
        self.use_accounting_style = Some(value);
        self
    }
    #[must_use]
    pub const fn with_duration_style(mut self, value: u32) -> Self {
        self.duration_style = Some(value);
        self
    }
    #[must_use]
    pub const fn with_base(mut self, value: u32) -> Self {
        self.base = Some(value);
        self
    }
    #[must_use]
    pub const fn with_base_places(mut self, value: u32) -> Self {
        self.base_places = Some(value);
        self
    }
    #[must_use]
    pub const fn with_base_use_minus_sign(mut self, value: bool) -> Self {
        self.base_use_minus_sign = Some(value);
        self
    }
    #[must_use]
    pub const fn with_fraction_accuracy(mut self, value: u32) -> Self {
        self.fraction_accuracy = Some(value);
        self
    }
    #[must_use]
    pub const fn with_suppress_date_format(mut self, value: bool) -> Self {
        self.suppress_date_format = Some(value);
        self
    }
    #[must_use]
    pub const fn with_suppress_time_format(mut self, value: bool) -> Self {
        self.suppress_time_format = Some(value);
        self
    }
    #[must_use]
    pub const fn with_date_time_format(mut self, value: &'source str) -> Self {
        self.date_time_format = Some(value);
        self
    }
    #[must_use]
    pub const fn with_duration_unit_largest(mut self, value: u32) -> Self {
        self.duration_unit_largest = Some(value);
        self
    }
    #[must_use]
    pub const fn with_duration_unit_smallest(mut self, value: u32) -> Self {
        self.duration_unit_smallest = Some(value);
        self
    }
}

/// Prepared canonical write for one control-oriented `FormatStructArchive`.
#[derive(Debug, Clone, Copy)]
pub struct PreparedControlFormatWrite<'source> {
    write: ControlFormatWrite<'source>,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl<'source> PreparedControlFormatWrite<'source> {
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_control_format(&mut bytes, self.write)?;
        verify_control_format_candidate(&bytes, self.verify_options, self.requirements)?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

impl PreparedCellSpecWrite {
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(self.requirements.output_bytes)
            .map_err(|_| {
                DecodeError::limited(DecodeLimit::Allocation {
                    requested: self.requirements.output_bytes,
                })
            })?;
        emit_cell_spec(&mut bytes, self.model_identifier, self.starts_with_first)?;
        verify_cell_spec_candidate(&bytes, self.verify_options, self.requirements)?;
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepare a canonical popup model payload from borrowed item text.
pub fn prepare_popup_menu_model_write<'items>(
    items: &'items [&'items str],
    options: DecodeOptions,
) -> Result<PreparedPopUpMenuModelWrite<'items>, DecodeError> {
    if items.is_empty() || items.len() > options.max_items {
        return Err(DecodeError::limited(DecodeLimit::Items {
            observed: items.len(),
            maximum: options.max_items,
        }));
    }
    let mut text_bytes = 0usize;
    for text in items {
        validate_text(text, options.max_text_bytes)?;
        text_bytes = text_bytes
            .checked_add(text.len())
            .ok_or_else(DecodeError::invalid)?;
    }
    let output_bytes = popup_model_output_len(items)?;
    let fields = 2usize
        .checked_add(
            items
                .len()
                .checked_mul(9)
                .ok_or_else(DecodeError::invalid)?,
        )
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: popup_model_work_bytes(items)?
            .checked_add(output_bytes)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 3,
        references: 0,
        items: items.len(),
        text_bytes,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedPopUpMenuModelWrite {
        items,
        requirements,
        verify_options: options,
    })
}

/// One-shot canonical popup model encoding wrapper.
pub fn canonical_popup_menu_model(
    items: &[&str],
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_popup_menu_model_write(items, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for package owners that describe canonical model
/// construction as a rewrite operation.
///
/// This is a new-payload constructor, not a source patch: it intentionally
/// does not accept an existing payload and therefore does not claim to carry
/// unknown fields or deprecated-but-valid source framing into the result.
pub fn rewrite_popup_menu_model(
    items: &[&str],
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_popup_menu_model(items, options)
}

/// Prepare a canonical popup `CellSpecArchive` payload.
pub fn prepare_cell_spec_write(
    model_identifier: u64,
    starts_with_first: bool,
    options: DecodeOptions,
) -> Result<PreparedCellSpecWrite, DecodeError> {
    if model_identifier == 0 {
        return Err(DecodeError::invalid());
    }
    let output_bytes = cell_spec_output_len(model_identifier)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields: 4,
        work_bytes: output_bytes
            .checked_add(1 + encoded_varint_len(model_identifier))
            .and_then(|work| work.checked_add(1 + encoded_varint_len(model_identifier)))
            .and_then(|work| work.checked_add(output_bytes))
            .and_then(|work| work.checked_add(output_bytes))
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 1,
        references: 1,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedCellSpecWrite {
        model_identifier,
        starts_with_first,
        requirements,
        verify_options: options,
    })
}

/// One-shot canonical popup cell-spec encoding wrapper.
pub fn canonical_cell_spec(
    model_identifier: u64,
    starts_with_first: bool,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_cell_spec_write(model_identifier, starts_with_first, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for canonical control-cell spec publication.
///
/// This is likewise creation-only.  Existing `CellSpecArchive` bytes must be
/// retained by the package owner unless a future source-preserving patch API
/// is added; this function never silently drops source unknowns because it
/// never accepts source bytes.
pub fn rewrite_cell_spec(
    model_identifier: u64,
    starts_with_first: bool,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_cell_spec(model_identifier, starts_with_first, options)
}

/// Prepare a canonical control-cell spec payload.
pub fn prepare_control_cell_spec_write(
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
    options: DecodeOptions,
) -> Result<PreparedControlCellSpecWrite, DecodeError> {
    validate_control_cell_spec_values(
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
    )?;
    let range_count = if matches!(
        interaction_type,
        SLIDER_INTERACTION_TYPE | STEPPER_INTERACTION_TYPE | STAR_RATING_INTERACTION_TYPE
    ) {
        3
    } else {
        0
    };
    let output_bytes = control_cell_spec_output_len(
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
    )?;
    let fields = 1usize
        .checked_add(range_count)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: output_bytes
            .checked_add(output_bytes)
            .and_then(|work| work.checked_add(output_bytes))
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 0,
        references: 0,
        items: 0,
        text_bytes: 0,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedControlCellSpecWrite {
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
        requirements,
        verify_options: options,
    })
}

/// One-shot canonical control-cell spec encoding wrapper.
pub fn canonical_control_cell_spec(
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_control_cell_spec_write(
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
        options,
    )?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for a canonical control-cell spec constructor.
pub fn rewrite_control_cell_spec(
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_control_cell_spec(
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
        options,
    )
}

/// Prepare a canonical control-oriented format payload with only field 1.
/// This is the compact constructor used by checkbox and star-rating formats.
pub fn prepare_control_format_write(
    format_type: u32,
    options: DecodeOptions,
) -> Result<PreparedControlFormatWrite<'static>, DecodeError> {
    prepare_control_format_write_fields(ControlFormatWrite::new(format_type), options)
}

/// Prepare a canonical display-format payload carrying the supported TSK
/// fields 1--16.  The caller-owned strings are borrowed until execution.
pub fn prepare_control_format_write_fields<'source>(
    write: ControlFormatWrite<'source>,
    options: DecodeOptions,
) -> Result<PreparedControlFormatWrite<'source>, DecodeError> {
    validate_control_format_write(write, options.max_text_bytes)?;
    let output_bytes = control_format_output_len(write)?;
    let fields = control_format_field_count(write);
    let text_bytes = control_format_text_bytes(write)?;
    let requirements = RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: output_bytes
            .checked_add(output_bytes)
            .ok_or_else(DecodeError::invalid)?,
        max_depth: 0,
        references: 0,
        items: 0,
        text_bytes,
        allocations: 1,
        retained_bytes: output_bytes,
        scratch_bytes: 0,
    };
    check_options(requirements, options)?;
    Ok(PreparedControlFormatWrite {
        write,
        requirements,
        verify_options: options,
    })
}

/// One-shot canonical control-oriented format encoding wrapper.
pub fn canonical_control_format(
    format_type: u32,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_control_format_write(format_type, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// One-shot canonical display-format writer carrying fields 1--16.
pub fn canonical_control_format_fields(
    write: ControlFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_control_format_write_fields(write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for a canonical control-oriented format
/// constructor.
pub fn rewrite_control_format(
    format_type: u32,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_control_format(format_type, options)
}

/// Compatibility spelling for the full display-format constructor.
pub fn rewrite_control_format_fields(
    write: ControlFormatWrite<'_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    canonical_control_format_fields(write, options)
}

fn report_from_requirements(requirements: RewriteExecutionRequirements) -> DecodeReport {
    DecodeReport {
        input_bytes: 0,
        output_bytes: requirements.output_bytes,
        fields: requirements.fields,
        work_bytes: requirements.work_bytes,
        max_depth: requirements.max_depth,
        references: requirements.references,
        items: requirements.items,
        text_bytes: requirements.text_bytes,
        allocations: requirements.allocations,
        retained_bytes: requirements.retained_bytes,
        scratch_bytes: requirements.scratch_bytes,
    }
}

fn verify_popup_model_candidate(
    source: &[u8],
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
) -> Result<(), DecodeError> {
    let (_, report) = decode_popup_menu_model_with_report(source, options)?;
    let candidate_work = requirements
        .work_bytes
        .checked_sub(requirements.output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields != requirements.fields
        || report.work_bytes != candidate_work
        || report.max_depth != requirements.max_depth
        || report.items != requirements.items
        || report.text_bytes != requirements.text_bytes
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn verify_cell_spec_candidate(
    source: &[u8],
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
) -> Result<(), DecodeError> {
    let (_, report) = decode_cell_spec_with_report(source, options)?;
    let candidate_work = requirements
        .work_bytes
        .checked_sub(requirements.output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields != requirements.fields
        || report.work_bytes != candidate_work
        || report.max_depth != requirements.max_depth
        || report.references != requirements.references
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn verify_control_cell_spec_candidate(
    source: &[u8],
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
) -> Result<(), DecodeError> {
    let (_, report) = decode_control_cell_spec_with_report(source, options)?;
    let candidate_work = requirements
        .work_bytes
        .checked_sub(requirements.output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields != requirements.fields
        || report.work_bytes != candidate_work
        || report.max_depth != requirements.max_depth
        || report.references != requirements.references
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn verify_control_format_candidate(
    source: &[u8],
    options: DecodeOptions,
    requirements: RewriteExecutionRequirements,
) -> Result<(), DecodeError> {
    let (_, report) = decode_control_format_with_report(source, options)?;
    let candidate_work = requirements
        .work_bytes
        .checked_sub(requirements.output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    if report.fields != requirements.fields
        || report.work_bytes != candidate_work
        || report.max_depth != requirements.max_depth
        || report.text_bytes != requirements.text_bytes
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn check_options(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
    // Candidate verification decodes the emitted bytes with the original
    // caller-owned message ceiling. Keep that ceiling authoritative rather
    // than widening it to fit the candidate after preparation; otherwise a
    // one-byte varint growth could allocate and publish a message the caller
    // explicitly refused to admit. This check runs before output allocation.
    if requirements.output_bytes > options.max_message_bytes {
        return Err(DecodeError::limited(DecodeLimit::InputBytes {
            observed: requirements.output_bytes,
            maximum: options.max_message_bytes,
        }));
    }
    if requirements.output_bytes > options.max_output_bytes {
        return Err(DecodeError::limited(DecodeLimit::OutputBytes {
            observed: requirements.output_bytes,
            maximum: options.max_output_bytes,
        }));
    }
    if requirements.fields > options.max_fields {
        return Err(DecodeError::limited(DecodeLimit::Fields {
            observed: requirements.fields,
            maximum: options.max_fields,
        }));
    }
    if requirements.work_bytes > options.max_work_bytes {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: requirements.work_bytes,
            maximum: options.max_work_bytes,
        }));
    }
    if requirements.max_depth > options.recursion_limit {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: options.recursion_limit,
        }));
    }
    if requirements.references > options.max_references {
        return Err(DecodeError::limited(DecodeLimit::References {
            observed: requirements.references,
            maximum: options.max_references,
        }));
    }
    if requirements.items > options.max_items {
        return Err(DecodeError::limited(DecodeLimit::Items {
            observed: requirements.items,
            maximum: options.max_items,
        }));
    }
    if requirements.text_bytes > options.max_text_bytes {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: requirements.text_bytes,
            maximum: options.max_text_bytes,
        }));
    }
    Ok(())
}

fn check_requirements(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    macro_rules! check {
        ($observed:expr, $maximum:expr, $kind:ident) => {
            if $observed > $maximum {
                return Err(DecodeError::limited(DecodeLimit::$kind {
                    observed: $observed,
                    maximum: $maximum,
                }));
            }
        };
    }
    check!(requirements.output_bytes, limits.output_bytes, OutputBytes);
    check!(requirements.fields, limits.fields, Fields);
    check!(requirements.work_bytes, limits.work_bytes, Work);
    if requirements.max_depth > limits.max_depth {
        return Err(DecodeError::limited(DecodeLimit::Nesting {
            observed: requirements.max_depth,
            maximum: limits.max_depth,
        }));
    }
    check!(requirements.references, limits.references, References);
    check!(requirements.items, limits.items, Items);
    check!(requirements.text_bytes, limits.text_bytes, Text);
    if requirements.allocations > limits.allocations {
        return Err(DecodeError::limited(DecodeLimit::Allocation {
            requested: requirements.allocations,
        }));
    }
    check!(requirements.retained_bytes, limits.retained_bytes, Retained);
    check!(requirements.scratch_bytes, limits.scratch_bytes, Scratch);
    Ok(())
}

fn validate_text(text: &str, maximum: usize) -> Result<(), DecodeError> {
    if text.len() > maximum || text.chars().any(char::is_control) {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: text.len(),
            maximum,
        }));
    }
    Ok(())
}

fn validate_control_format_write(
    write: ControlFormatWrite<'_>,
    max_text_bytes: usize,
) -> Result<(), DecodeError> {
    if !matches!(write.format_type, 256..=263 | 267..=269) {
        return Err(DecodeError::invalid());
    }
    if let Some(value) = write.currency_code {
        validate_text(value, max_text_bytes)?;
    }
    if let Some(value) = write.date_time_format {
        validate_text(value, max_text_bytes)?;
    }
    if matches!(write.format_type, 263 | 267)
        && (write.decimal_places.is_some()
            || write.currency_code.is_some()
            || write.negative_style.is_some()
            || write.show_thousands_separator.is_some()
            || write.use_accounting_style.is_some()
            || write.duration_style.is_some()
            || write.base.is_some()
            || write.base_places.is_some()
            || write.base_use_minus_sign.is_some()
            || write.fraction_accuracy.is_some()
            || write.suppress_date_format.is_some()
            || write.suppress_time_format.is_some()
            || write.date_time_format.is_some()
            || write.duration_unit_largest.is_some()
            || write.duration_unit_smallest.is_some())
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn control_format_field_count(write: ControlFormatWrite<'_>) -> usize {
    1usize
        + usize::from(write.decimal_places.is_some())
        + usize::from(write.currency_code.is_some())
        + usize::from(write.negative_style.is_some())
        + usize::from(write.show_thousands_separator.is_some())
        + usize::from(write.use_accounting_style.is_some())
        + usize::from(write.duration_style.is_some())
        + usize::from(write.base.is_some())
        + usize::from(write.base_places.is_some())
        + usize::from(write.base_use_minus_sign.is_some())
        + usize::from(write.fraction_accuracy.is_some())
        + usize::from(write.suppress_date_format.is_some())
        + usize::from(write.suppress_time_format.is_some())
        + usize::from(write.date_time_format.is_some())
        + usize::from(write.duration_unit_largest.is_some())
        + usize::from(write.duration_unit_smallest.is_some())
}

fn control_format_text_bytes(write: ControlFormatWrite<'_>) -> Result<usize, DecodeError> {
    let mut total = 0usize;
    if let Some(value) = write.currency_code {
        total = total
            .checked_add(value.len())
            .ok_or_else(DecodeError::invalid)?;
    }
    if let Some(value) = write.date_time_format {
        total = total
            .checked_add(value.len())
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(total)
}

fn control_format_output_len(write: ControlFormatWrite<'_>) -> Result<usize, DecodeError> {
    let mut length = 0usize;
    add_control_format_varint_len(&mut length, FORMAT_TYPE_FIELD, u64::from(write.format_type))?;
    if let Some(value) = write.decimal_places {
        add_control_format_varint_len(&mut length, FORMAT_DECIMAL_PLACES_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.currency_code {
        add_control_format_string_len(&mut length, FORMAT_CURRENCY_CODE_FIELD, value)?;
    }
    if let Some(value) = write.negative_style {
        add_control_format_varint_len(&mut length, FORMAT_NEGATIVE_STYLE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.show_thousands_separator {
        add_control_format_varint_len(
            &mut length,
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.use_accounting_style {
        add_control_format_varint_len(
            &mut length,
            FORMAT_USE_ACCOUNTING_STYLE_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.duration_style {
        add_control_format_varint_len(&mut length, FORMAT_DURATION_STYLE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.base {
        add_control_format_varint_len(&mut length, FORMAT_BASE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.base_places {
        add_control_format_varint_len(&mut length, FORMAT_BASE_PLACES_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.base_use_minus_sign {
        add_control_format_varint_len(
            &mut length,
            FORMAT_BASE_USE_MINUS_SIGN_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.fraction_accuracy {
        add_control_format_varint_len(
            &mut length,
            FORMAT_FRACTION_ACCURACY_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.suppress_date_format {
        add_control_format_varint_len(
            &mut length,
            FORMAT_SUPPRESS_DATE_FORMAT_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.suppress_time_format {
        add_control_format_varint_len(
            &mut length,
            FORMAT_SUPPRESS_TIME_FORMAT_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.date_time_format {
        add_control_format_string_len(&mut length, FORMAT_DATE_TIME_FORMAT_FIELD, value)?;
    }
    if let Some(value) = write.duration_unit_largest {
        add_control_format_varint_len(
            &mut length,
            FORMAT_DURATION_UNIT_LARGEST_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.duration_unit_smallest {
        add_control_format_varint_len(
            &mut length,
            FORMAT_DURATION_UNIT_SMALLEST_FIELD,
            u64::from(value),
        )?;
    }
    Ok(length)
}

fn add_control_format_varint_len(
    length: &mut usize,
    field: u32,
    value: u64,
) -> Result<(), DecodeError> {
    *length = length
        .checked_add(encoded_varint_len(u64::from(field) << 3))
        .and_then(|length| length.checked_add(encoded_varint_len(value)))
        .ok_or_else(DecodeError::invalid)?;
    Ok(())
}

fn add_control_format_string_len(
    length: &mut usize,
    field: u32,
    value: &str,
) -> Result<(), DecodeError> {
    *length = length
        .checked_add(length_field_len(field, value.len())?)
        .ok_or_else(DecodeError::invalid)?;
    Ok(())
}

fn validate_control_cell_spec_values(
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
) -> Result<(), DecodeError> {
    if !matches!(
        interaction_type,
        STEPPER_INTERACTION_TYPE
            | SLIDER_INTERACTION_TYPE
            | STAR_RATING_INTERACTION_TYPE
            | CHECKBOX_INTERACTION_TYPE
    ) {
        return Err(DecodeError::invalid());
    }
    let requires_range = matches!(
        interaction_type,
        STEPPER_INTERACTION_TYPE | SLIDER_INTERACTION_TYPE | STAR_RATING_INTERACTION_TYPE
    );
    if requires_range {
        let (Some(min), Some(max), Some(inc)) =
            (range_control_min, range_control_max, range_control_inc)
        else {
            return Err(DecodeError::invalid());
        };
        if !min.is_finite()
            || !max.is_finite()
            || !inc.is_finite()
            || min >= max
            || inc <= 0.0
            || !((max - min) / inc).is_finite()
        {
            return Err(DecodeError::invalid());
        }
        if interaction_type == STAR_RATING_INTERACTION_TYPE
            && (min != 0.0 || max != 5.0 || inc != 1.0)
        {
            return Err(DecodeError::invalid());
        }
    } else if range_control_min.is_some()
        || range_control_max.is_some()
        || range_control_inc.is_some()
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn control_cell_spec_output_len(
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
) -> Result<usize, DecodeError> {
    validate_control_cell_spec_values(
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
    )?;
    let mut length = 1usize
        .checked_add(encoded_varint_len(u64::from(interaction_type)))
        .ok_or_else(DecodeError::invalid)?;
    if matches!(
        interaction_type,
        STEPPER_INTERACTION_TYPE | SLIDER_INTERACTION_TYPE | STAR_RATING_INTERACTION_TYPE
    ) {
        for field in 3u32..=5 {
            length = length
                .checked_add(encoded_varint_len((u64::from(field) << 3) | 1))
                .and_then(|value| value.checked_add(8))
                .ok_or_else(DecodeError::invalid)?;
        }
    }
    Ok(length)
}

fn popup_model_output_len(items: &[&str]) -> Result<usize, DecodeError> {
    let nil_cell = 2usize;
    let nil_outer = length_field_len(POPUP_ITEM_FIELD, nil_cell)?;
    let mut length = nil_outer;
    for text in items {
        let format_payload_len = 1usize
            .checked_add(encoded_varint_len(FORMAT_TYPE_TEXT))
            .ok_or_else(DecodeError::invalid)?;
        let string_len = length_field_len(STRING_VALUE_FIELD, text.len())?
            .checked_add(length_field_len(STRING_FORMAT_FIELD, format_payload_len)?)
            .and_then(|value| value.checked_add(2 + 2 + 2))
            .ok_or_else(DecodeError::invalid)?;
        let cell_len = 2usize
            .checked_add(length_field_len(CELL_VALUE_STRING_FIELD, string_len)?)
            .ok_or_else(DecodeError::invalid)?;
        length = length
            .checked_add(length_field_len(POPUP_ITEM_FIELD, cell_len)?)
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn popup_model_work_bytes(items: &[&str]) -> Result<usize, DecodeError> {
    let mut nested = 4usize; // nil payload + one Buffa parity scan
    for text in items {
        let string_len = 1usize
            .checked_add(encoded_varint_len(
                u64::try_from(text.len()).map_err(|_| DecodeError::invalid())?,
            ))
            .and_then(|value| value.checked_add(text.len()))
            .and_then(|value| value.checked_add(5 + 2 + 2 + 2))
            .ok_or_else(DecodeError::invalid)?;
        let cell_len = 2usize
            .checked_add(
                1 + encoded_varint_len(
                    u64::try_from(string_len).map_err(|_| DecodeError::invalid())?,
                ) + string_len,
            )
            .ok_or_else(DecodeError::invalid)?;
        nested = nested
            .checked_add(cell_len)
            .and_then(|value| value.checked_add(string_len))
            .and_then(|value| value.checked_add(3))
            .and_then(|value| value.checked_add(cell_len))
            .and_then(|value| value.checked_add(string_len))
            .ok_or_else(DecodeError::invalid)?;
    }
    popup_model_output_len(items)?
        .checked_add(nested)
        .ok_or_else(DecodeError::invalid)
}

fn cell_spec_output_len(model_identifier: u64) -> Result<usize, DecodeError> {
    let reference_len = 1usize
        .checked_add(encoded_varint_len(model_identifier))
        .ok_or_else(DecodeError::invalid)?;
    2usize
        .checked_add(length_field_len(CELL_SPEC_MODEL_FIELD, reference_len)?)
        .and_then(|value| value.checked_add(2))
        .ok_or_else(DecodeError::invalid)
}

fn length_field_len(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let key = encoded_varint_len(u64::from(number) << 3);
    let payload =
        encoded_varint_len(u64::try_from(payload_len).map_err(|_| DecodeError::invalid())?);
    key.checked_add(payload)
        .and_then(|value| value.checked_add(payload_len))
        .ok_or_else(DecodeError::invalid)
}

fn emit_popup_model(output: &mut Vec<u8>, items: &[&str]) -> Result<(), DecodeError> {
    emit_len_field(output, POPUP_ITEM_FIELD, &[0x08, 0x01])?;
    for text in items {
        let string_len = 1usize
            .checked_add(encoded_varint_len(
                u64::try_from(text.len()).map_err(|_| DecodeError::invalid())?,
            ))
            .and_then(|value| value.checked_add(text.len()))
            .and_then(|value| value.checked_add(5 + 2 + 2 + 2))
            .ok_or_else(DecodeError::invalid)?;
        let cell_len = 2usize
            .checked_add(
                1 + encoded_varint_len(
                    u64::try_from(string_len).map_err(|_| DecodeError::invalid())?,
                ) + string_len,
            )
            .ok_or_else(DecodeError::invalid)?;
        emit_len_field_header(output, POPUP_ITEM_FIELD, cell_len)?;
        emit_varint_field(output, CELL_VALUE_TYPE_FIELD, STRING_TYPE)?;
        emit_len_field_header(output, CELL_VALUE_STRING_FIELD, string_len)?;
        emit_len_field(output, STRING_VALUE_FIELD, text.as_bytes())?;
        emit_len_field(output, STRING_FORMAT_FIELD, &[0x08, 0x84, 0x02])?;
        emit_varint_field(output, STRING_EXPLICIT_FIELD, 0)?;
        emit_varint_field(output, STRING_REGEX_FIELD, 0)?;
        emit_varint_field(output, STRING_CASE_SENSITIVE_REGEX_FIELD, 0)?;
    }
    Ok(())
}

fn emit_cell_spec(
    output: &mut Vec<u8>,
    model_identifier: u64,
    starts_with_first: bool,
) -> Result<(), DecodeError> {
    emit_varint_field(
        output,
        CELL_SPEC_INTERACTION_FIELD,
        u64::from(POPUP_INTERACTION_TYPE),
    )?;
    let reference_len = 1usize
        .checked_add(encoded_varint_len(model_identifier))
        .ok_or_else(DecodeError::invalid)?;
    emit_len_field_header(output, CELL_SPEC_MODEL_FIELD, reference_len)?;
    emit_varint_field(output, REFERENCE_IDENTIFIER_FIELD, model_identifier)?;
    emit_varint_field(output, CELL_SPEC_FIRST_FIELD, u64::from(starts_with_first))?;
    Ok(())
}

fn emit_control_cell_spec(
    output: &mut Vec<u8>,
    interaction_type: u32,
    range_control_min: Option<f64>,
    range_control_max: Option<f64>,
    range_control_inc: Option<f64>,
) -> Result<(), DecodeError> {
    validate_control_cell_spec_values(
        interaction_type,
        range_control_min,
        range_control_max,
        range_control_inc,
    )?;
    emit_varint_field(
        output,
        CELL_SPEC_INTERACTION_FIELD,
        u64::from(interaction_type),
    )?;
    if matches!(
        interaction_type,
        STEPPER_INTERACTION_TYPE | SLIDER_INTERACTION_TYPE | STAR_RATING_INTERACTION_TYPE
    ) {
        emit_fixed64_field(
            output,
            3,
            range_control_min.ok_or_else(DecodeError::invalid)?,
        )?;
        emit_fixed64_field(
            output,
            4,
            range_control_max.ok_or_else(DecodeError::invalid)?,
        )?;
        emit_fixed64_field(
            output,
            5,
            range_control_inc.ok_or_else(DecodeError::invalid)?,
        )?;
    }
    Ok(())
}

fn emit_control_format(
    output: &mut Vec<u8>,
    write: ControlFormatWrite<'_>,
) -> Result<(), DecodeError> {
    validate_control_format_write(write, usize::MAX)?;
    emit_varint_field(output, FORMAT_TYPE_FIELD, u64::from(write.format_type))?;
    if let Some(value) = write.decimal_places {
        emit_varint_field(output, FORMAT_DECIMAL_PLACES_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.currency_code {
        emit_len_field(output, FORMAT_CURRENCY_CODE_FIELD, value.as_bytes())?;
    }
    if let Some(value) = write.negative_style {
        emit_varint_field(output, FORMAT_NEGATIVE_STYLE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.show_thousands_separator {
        emit_varint_field(
            output,
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(value),
        )?;
    }
    if let Some(value) = write.use_accounting_style {
        emit_varint_field(output, FORMAT_USE_ACCOUNTING_STYLE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.duration_style {
        emit_varint_field(output, FORMAT_DURATION_STYLE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.base {
        emit_varint_field(output, FORMAT_BASE_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.base_places {
        emit_varint_field(output, FORMAT_BASE_PLACES_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.base_use_minus_sign {
        emit_varint_field(output, FORMAT_BASE_USE_MINUS_SIGN_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.fraction_accuracy {
        emit_varint_field(output, FORMAT_FRACTION_ACCURACY_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.suppress_date_format {
        emit_varint_field(output, FORMAT_SUPPRESS_DATE_FORMAT_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.suppress_time_format {
        emit_varint_field(output, FORMAT_SUPPRESS_TIME_FORMAT_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.date_time_format {
        emit_len_field(output, FORMAT_DATE_TIME_FORMAT_FIELD, value.as_bytes())?;
    }
    if let Some(value) = write.duration_unit_largest {
        emit_varint_field(output, FORMAT_DURATION_UNIT_LARGEST_FIELD, u64::from(value))?;
    }
    if let Some(value) = write.duration_unit_smallest {
        emit_varint_field(
            output,
            FORMAT_DURATION_UNIT_SMALLEST_FIELD,
            u64::from(value),
        )?;
    }
    Ok(())
}

fn emit_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), DecodeError> {
    write_varint(output, u64::from(number) << 3)?;
    write_varint(output, value)
}

fn emit_fixed64_field(output: &mut Vec<u8>, number: u32, value: f64) -> Result<(), DecodeError> {
    if !value.is_finite() {
        return Err(DecodeError::invalid());
    }
    write_varint(output, (u64::from(number) << 3) | 1)?;
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn emit_len_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), DecodeError> {
    emit_len_field_header(output, number, payload.len())?;
    output.extend_from_slice(payload);
    Ok(())
}

fn emit_len_field_header(
    output: &mut Vec<u8>,
    number: u32,
    payload_len: usize,
) -> Result<(), DecodeError> {
    write_varint(output, u64::from(number) << 3 | 2)?;
    write_varint(
        output,
        u64::try_from(payload_len).map_err(|_| DecodeError::invalid())?,
    )
}

fn write_varint(output: &mut Vec<u8>, mut value: u64) -> Result<(), DecodeError> {
    while value >= 0x80 {
        output.push(value as u8 | 0x80);
        value >>= 7;
    }
    output.push(u8::try_from(value).map_err(|_| DecodeError::invalid())?);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers_table_cell_currency_format_codec as currency_codec;
    use crate::numbers_table_cell_percentage_format_codec as percentage_codec;

    fn options() -> DecodeOptions {
        DecodeOptions::new(16 * 1024, 16 * 1024, 16 * 1024, 64 * 1024, 64, 16, 64, 4096)
    }

    #[test]
    fn canonical_model_roundtrips_nil_and_strings() {
        let items = ["Low", "High"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let (snapshot, report) =
            decode_popup_menu_model_with_report(output.bytes(), options()).expect("decode model");
        assert!(snapshot.has_nil_sentinel());
        assert_eq!(snapshot.item_count(), 2);
        assert_eq!(
            snapshot
                .items()
                .map(PopUpMenuItem::value)
                .collect::<Vec<_>>(),
            items
        );
        assert_eq!(report.input_bytes(), output.bytes().len());
        assert_eq!(report.text_bytes(), 7);
    }

    #[test]
    fn canonical_cell_spec_roundtrips_reference_and_selection() {
        let output = canonical_cell_spec(41, true, options()).expect("canonical cell spec");
        let (snapshot, report) =
            decode_cell_spec_with_report(output.bytes(), options()).expect("decode cell spec");
        assert_eq!(snapshot.interaction_type(), POPUP_INTERACTION_TYPE);
        assert_eq!(snapshot.popup_model().identifier(), 41);
        assert!(snapshot.starts_with_first());
        assert_eq!(report.references(), 1);
    }

    #[test]
    fn cell_spec_unknown_fields_are_raw_preserved_without_buffa_rejection() {
        let output = canonical_cell_spec(41, true, options()).expect("canonical cell spec");
        let mut source = output.bytes().to_vec();
        source.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
        let snapshot = decode_cell_spec(&source, options()).expect("opaque unknown field");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.popup_model().identifier(), 41);
    }

    #[test]
    fn strict_model_rejects_deprecated_item_and_missing_sentinel() {
        let items = ["One"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let mut deprecated = output.bytes().to_vec();
        deprecated.extend_from_slice(&[0x0a, 0x00]);
        assert!(decode_popup_menu_model(&deprecated, options()).is_err());

        let mut missing = output.bytes().to_vec();
        // The first field is the two-byte nil payload wrapped in field 2.
        assert_eq!(&missing[..4], &[0x12, 0x02, 0x08, 0x01]);
        missing.drain(..4);
        assert!(decode_popup_menu_model(&missing, options()).is_err());
    }

    #[test]
    fn strict_cell_spec_rejects_zero_or_duplicate_reference() {
        let mut zero = vec![0x08, 0x07, 0x32, 0x02, 0x08, 0x00, 0x38, 0x01];
        assert!(decode_cell_spec(&zero, options()).is_err());
        zero[5] = 0x01;
        zero.extend_from_slice(&[0x32, 0x02, 0x08, 0x02]);
        assert!(decode_cell_spec(&zero, options()).is_err());
    }

    #[test]
    fn strict_cell_spec_rejects_deprecated_reference_presence_even_at_defaults() {
        let output = canonical_cell_spec(41, false, options()).expect("canonical cell spec");

        let mut deprecated_type = output.bytes().to_vec();
        deprecated_type[3] += 2;
        deprecated_type.splice(6..6, [REFERENCE_TYPE_FIELD as u8 * 2, 0]);
        assert!(decode_cell_spec(&deprecated_type, options()).is_err());

        let mut deprecated_external = output.bytes().to_vec();
        deprecated_external[3] += 2;
        deprecated_external.splice(6..6, [REFERENCE_EXTERNAL_FIELD as u8 * 2, 0]);
        assert!(decode_cell_spec(&deprecated_external, options()).is_err());
    }

    #[test]
    fn unknown_group_is_structurally_checked_and_raw_retained() {
        let items = ["One"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let mut source = output.bytes().to_vec();
        // Unknown field 100, balanced start/end group with an unknown scalar.
        source.extend_from_slice(&[
            0xa0, 0x06, 0x81, 0x00, // unknown scalar value 1, deliberately overlong
            0xa3, 0x06, 0x08, 0x81, 0x00, 0xa4, 0x06,
        ]);
        let snapshot = decode_popup_menu_model(&source, options()).expect("unknown group");
        assert_eq!(snapshot.raw(), source.as_slice());

        let mut unterminated = source;
        unterminated.pop();
        assert!(decode_popup_menu_model(&unterminated, options()).is_err());
    }

    #[test]
    fn unknown_group_fields_work_and_nesting_are_budgeted() {
        let items = ["One"];
        let output = canonical_popup_menu_model(&items, options()).expect("canonical model");
        let mut source = output.bytes().to_vec();
        source.extend_from_slice(&[
            0xa3, 0x06, // unknown group 100
            0xa3, 0x06, 0x08, 0x81, 0x00, 0xa4, 0x06, // nested group + scalar
            0xa4, 0x06,
        ]);
        let (_, report) =
            decode_popup_menu_model_with_report(&source, options()).expect("metered unknown group");
        assert!(report.fields() > 0);
        assert!(report.work_bytes() >= source.len());
        assert!(
            decode_popup_menu_model(&source, options().with_max_output_bytes(usize::MAX),).is_ok()
        );
        let field_limited = DecodeOptions::new(
            options().max_message_bytes,
            options().max_output_bytes,
            report.fields() - 1,
            options().max_work_bytes,
            options().recursion_limit,
            options().max_references,
            options().max_items,
            options().max_text_bytes,
        );
        assert!(matches!(
            decode_popup_menu_model(&source, field_limited)
                .expect_err("nested field ceiling")
                .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        let work_limited = DecodeOptions::new(
            options().max_message_bytes,
            options().max_output_bytes,
            options().max_fields,
            report.work_bytes() - 1,
            options().recursion_limit,
            options().max_references,
            options().max_items,
            options().max_text_bytes,
        );
        assert!(matches!(
            decode_popup_menu_model(&source, work_limited)
                .expect_err("nested work ceiling")
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        let nesting_limited = DecodeOptions::new(
            options().max_message_bytes,
            options().max_output_bytes,
            options().max_fields,
            options().max_work_bytes,
            1,
            options().max_references,
            options().max_items,
            options().max_text_bytes,
        );
        assert!(matches!(
            decode_popup_menu_model(&source, nesting_limited)
                .expect_err("nested depth ceiling")
                .resource_limit(),
            Some(DecodeLimit::Nesting { .. })
        ));
    }

    #[test]
    fn visitor_streams_items_without_a_second_reported_parse() {
        struct Visitor {
            values: Vec<String>,
        }

        impl PopUpMenuVisitor for Visitor {
            fn visit_item(&mut self, item: PopUpMenuItem<'_>) -> Result<(), DecodeError> {
                self.values.push(item.value().to_owned());
                Ok(())
            }
        }

        let output =
            canonical_popup_menu_model(&["Low", "High"], options()).expect("canonical model");
        let mut visitor = Visitor { values: Vec::new() };
        let visitor_report =
            decode_popup_menu_model_with_visitor(output.bytes(), options(), &mut visitor)
                .expect("stream visitor");
        let (_, report) = decode_popup_menu_model_with_report(output.bytes(), options())
            .expect("ordinary report");
        assert_eq!(visitor.values, ["Low", "High"]);
        assert_eq!(visitor_report, report);
    }

    #[test]
    fn prepared_model_exact_replay_and_minus_one_axes_fail_before_output() {
        let items = ["A", "B"];
        let prepared = prepare_popup_menu_model_write(&items, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let exact = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("exact execute");
        assert_eq!(exact.bytes().len(), requirements.output_bytes());
        assert_eq!(exact.report().fields(), requirements.fields());
        assert_eq!(exact.report().work_bytes(), requirements.work_bytes());
        let (_, decoded_report) = decode_popup_menu_model_with_report(exact.bytes(), options())
            .expect("strict candidate verification");
        assert_eq!(decoded_report.fields(), requirements.fields());
        assert_eq!(
            decoded_report.work_bytes(),
            requirements.work_bytes() - requirements.output_bytes()
        );
        assert_eq!(decoded_report.max_depth(), requirements.max_depth());
        assert_eq!(decoded_report.items(), requirements.items());
        assert_eq!(decoded_report.text_bytes(), requirements.text_bytes());
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .expect_err("output ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::OutputBytes { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1)
                )
                .expect_err("field ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Fields { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1)
                )
                .expect_err("work ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Work { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_max_depth(requirements.max_depth() - 1)
                )
                .expect_err("depth ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Nesting { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_items(requirements.items() - 1)
                )
                .expect_err("item ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Items { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_text_bytes(requirements.text_bytes() - 1)
                )
                .expect_err("text ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Text { .. }))
        );
        assert!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements).with_allocations(0))
                .expect_err("allocation ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Allocation { .. }))
        );
        assert!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_retained_bytes(requirements.retained_bytes() - 1)
                )
                .expect_err("retained ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::Retained { .. }))
        );
    }

    #[test]
    fn prepared_cell_spec_exact_replay_and_reference_limit() {
        let prepared = prepare_cell_spec_write(91, false, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute");
        let snapshot = decode_cell_spec(output.bytes(), options()).expect("decode");
        assert_eq!(snapshot.popup_model().identifier(), 91);
        assert!(!snapshot.starts_with_first());
        let (_, decoded_report) = decode_cell_spec_with_report(output.bytes(), options())
            .expect("strict candidate verification");
        assert_eq!(decoded_report.fields(), requirements.fields());
        assert_eq!(
            decoded_report.work_bytes(),
            requirements.work_bytes() - requirements.output_bytes()
        );
        assert_eq!(decoded_report.max_depth(), requirements.max_depth());
        assert_eq!(decoded_report.references(), requirements.references());
        assert!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements).with_references(0))
                .expect_err("reference ceiling")
                .resource_limit()
                .is_some_and(|limit| matches!(limit, DecodeLimit::References { .. }))
        );
    }

    #[test]
    fn strict_control_cell_specs_cover_checkbox_star_slider_and_stepper() {
        let cases = [
            (CHECKBOX_INTERACTION_TYPE, None, None, None),
            (
                STAR_RATING_INTERACTION_TYPE,
                Some(0.0),
                Some(5.0),
                Some(1.0),
            ),
            (SLIDER_INTERACTION_TYPE, Some(-1.0), Some(1.0), Some(0.25)),
            (STEPPER_INTERACTION_TYPE, Some(0.0), Some(10.0), Some(1.0)),
        ];
        for (interaction, minimum, maximum, increment) in cases {
            let output =
                canonical_control_cell_spec(interaction, minimum, maximum, increment, options())
                    .expect("canonical control cell spec");
            let (snapshot, report) =
                decode_control_cell_spec_with_report(output.bytes(), options())
                    .expect("decode control cell spec");
            assert_eq!(snapshot.interaction_type(), interaction);
            assert_eq!(snapshot.range_control_min(), minimum);
            assert_eq!(snapshot.range_control_max(), maximum);
            assert_eq!(snapshot.range_control_inc(), increment);
            assert_eq!(report.input_bytes(), output.bytes().len());
            assert_eq!(report.fields(), 1 + usize::from(minimum.is_some()) * 3);
        }
    }

    #[test]
    fn control_cell_spec_unknown_group_is_raw_preserved_and_ranges_are_strict() {
        let output = canonical_control_cell_spec(
            SLIDER_INTERACTION_TYPE,
            Some(0.0),
            Some(10.0),
            Some(1.0),
            options(),
        )
        .expect("canonical slider");
        let mut source = output.bytes().to_vec();
        source.extend_from_slice(&[
            0xa3, 0x06, 0x08, 0x81, 0x00, 0xa4, 0x06, // balanced unknown group
        ]);
        let snapshot = decode_control_cell_spec(&source, options()).expect("unknown group");
        assert_eq!(snapshot.raw(), source.as_slice());

        let missing_range = canonical_control_cell_spec(
            SLIDER_INTERACTION_TYPE,
            Some(0.0),
            Some(10.0),
            Some(1.0),
            options(),
        )
        .expect("canonical slider")
        .bytes()
        .get(..18)
        .expect("range prefix")
        .to_vec();
        assert!(decode_control_cell_spec(&missing_range, options()).is_err());

        let mut checkbox_with_range =
            canonical_control_cell_spec(CHECKBOX_INTERACTION_TYPE, None, None, None, options())
                .expect("canonical checkbox")
                .bytes()
                .to_vec();
        checkbox_with_range.extend_from_slice(&[0x19, 0, 0, 0, 0, 0, 0, 0, 0]);
        assert!(decode_control_cell_spec(&checkbox_with_range, options()).is_err());
    }

    #[test]
    fn prepared_control_cell_spec_replays_exactly_and_honors_typed_limits() {
        let prepared = prepare_control_cell_spec_write(
            STAR_RATING_INTERACTION_TYPE,
            Some(0.0),
            Some(5.0),
            Some(1.0),
            options(),
        )
        .expect("prepare star");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("exact star execute");
        assert_eq!(output.bytes().len(), requirements.output_bytes());
        assert_eq!(output.report().fields(), requirements.fields());
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .expect_err("output ceiling")
                .resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1)
                )
                .expect_err("field ceiling")
                .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1)
                )
                .expect_err("work ceiling")
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        assert!(matches!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements).with_allocations(0))
                .expect_err("allocation ceiling")
                .resource_limit(),
            Some(DecodeLimit::Allocation { .. })
        ));
    }

    #[test]
    fn control_format_preserves_full_display_fields_and_rejects_checkbox_extras() {
        let write = ControlFormatWrite::new(256)
            .with_decimal_places(3)
            .with_currency_code("USD")
            .with_negative_style(2)
            .with_show_thousands_separator(true)
            .with_use_accounting_style(false)
            .with_duration_style(4)
            .with_base(16)
            .with_base_places(2)
            .with_base_use_minus_sign(true)
            .with_fraction_accuracy(5)
            .with_suppress_date_format(false)
            .with_suppress_time_format(true)
            .with_date_time_format("yyyy-MM-dd")
            .with_duration_unit_largest(1)
            .with_duration_unit_smallest(2);
        let output = canonical_control_format_fields(write, options()).expect("format write");
        let (snapshot, report) =
            decode_control_format_with_report(output.bytes(), options()).expect("format read");
        assert_eq!(snapshot.format_type(), 256);
        assert_eq!(snapshot.decimal_places(), Some(3));
        assert_eq!(snapshot.currency_code(), Some("USD"));
        assert_eq!(snapshot.negative_style(), Some(2));
        assert_eq!(snapshot.show_thousands_separator(), Some(true));
        assert_eq!(snapshot.use_accounting_style(), Some(false));
        assert_eq!(snapshot.duration_style(), Some(4));
        assert_eq!(snapshot.base(), Some(16));
        assert_eq!(snapshot.base_places(), Some(2));
        assert_eq!(snapshot.base_use_minus_sign(), Some(true));
        assert_eq!(snapshot.fraction_accuracy(), Some(5));
        assert_eq!(snapshot.suppress_date_format(), Some(false));
        assert_eq!(snapshot.suppress_time_format(), Some(true));
        assert_eq!(snapshot.date_time_format(), Some("yyyy-MM-dd"));
        assert_eq!(snapshot.duration_unit_largest(), Some(1));
        assert_eq!(snapshot.duration_unit_smallest(), Some(2));
        assert_eq!(report.fields(), 16);
        assert_eq!(report.text_bytes(), 13);

        let mut invalid_checkbox = canonical_control_format(263, options())
            .expect("checkbox format")
            .bytes()
            .to_vec();
        invalid_checkbox.extend_from_slice(&[0x10, 0x01]);
        assert!(decode_control_format(&invalid_checkbox, options()).is_err());
    }

    fn native_number_format(decimal_places: u32, negative_style: u32, show: bool) -> Vec<u8> {
        let mut source = Vec::new();
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_TYPE_FIELD,
            u64::from(NATIVE_NUMBER_FORMAT_TYPE),
        )
        .expect("type");
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_DECIMAL_PLACES_FIELD,
            u64::from(decimal_places),
        )
        .expect("decimal places");
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_NEGATIVE_STYLE_FIELD,
            u64::from(negative_style),
        )
        .expect("negative style");
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(show),
        )
        .expect("thousands separator");
        source
    }

    #[test]
    fn number_format_native_fixture_uses_borrowed_buffa_parity_view() {
        let source = [
            0x08, 0x80, 0x02, // format_type = 256
            0x10, 0xfd, 0x01, // decimal_places = automatic (253)
            0x20, 0x02, // negative_style = red
            0x28, 0x01, // show_thousands_separator = true
        ];
        let (snapshot, report) =
            decode_number_format_with_report(&source, options()).expect("number format");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_NUMBER_FORMAT_TYPE);
        assert_eq!(snapshot.decimal_places(), NATIVE_AUTOMATIC_DECIMAL_PLACES);
        assert_eq!(snapshot.negative_style(), 2);
        assert!(snapshot.show_thousands_separator());
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.fields(), 4);
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.allocations(), 0);
    }

    #[test]
    fn number_format_accepts_native_scalar_domains_only() {
        for decimal_places in [
            0,
            MAX_NUMBER_DECIMAL_PLACES,
            NATIVE_AUTOMATIC_DECIMAL_PLACES,
        ] {
            let source = native_number_format(decimal_places, 0, false);
            assert!(decode_number_format(&source, options()).is_ok());
        }
        for negative_style in 0..=3 {
            let source = native_number_format(2, negative_style, false);
            assert!(decode_number_format(&source, options()).is_ok());
        }
        for show in [false, true] {
            let source = native_number_format(2, 0, show);
            assert!(decode_number_format(&source, options()).is_ok());
        }
        for decimal_places in [31, 254] {
            let source = native_number_format(decimal_places, 0, false);
            assert!(decode_number_format(&source, options()).is_err());
        }
        let source = native_number_format(2, 4, false);
        assert!(decode_number_format(&source, options()).is_err());
        let source = native_number_format(2, 0, false);
        let mut invalid_bool = source;
        *invalid_bool.last_mut().expect("bool byte") = 2;
        assert!(decode_number_format(&invalid_bool, options()).is_err());
    }

    #[test]
    fn number_format_rejects_missing_duplicate_wrong_and_incompatible_fields() {
        let fields = [
            &[0x08, 0x80, 0x02][..],
            &[0x10, 0x02][..],
            &[0x20, 0x00][..],
            &[0x28, 0x00][..],
        ];
        for omitted in 0..fields.len() {
            let mut source = Vec::new();
            for (index, field) in fields.iter().enumerate() {
                if index != omitted {
                    source.extend_from_slice(field);
                }
            }
            assert!(decode_number_format(&source, options()).is_err());
        }

        let mut duplicate = native_number_format(2, 0, false);
        duplicate.extend_from_slice(&[0x08, 0x80, 0x02]);
        assert!(decode_number_format(&duplicate, options()).is_err());

        let wrong_wire = [
            0x0a, 0x01, 0x00, // field 1 encoded as length-delimited
            0x10, 0x02, 0x20, 0x00, 0x28, 0x00,
        ];
        assert!(decode_number_format(&wrong_wire, options()).is_err());

        let mut noncanonical = native_number_format(2, 0, false);
        noncanonical.splice(1..3, [0x80, 0x82, 0x00]);
        assert!(decode_number_format(&noncanonical, options()).is_err());

        for incompatible in [[0x1a, 0x00], [0x30, 0x01], [0x72, 0x00]] {
            let mut source = native_number_format(2, 0, false);
            source.extend_from_slice(&incompatible);
            assert!(decode_number_format(&source, options()).is_err());
        }
    }

    #[test]
    fn number_format_rejects_malformed_groups_but_preserves_unknown_extensions() {
        let mut unterminated = native_number_format(2, 0, false);
        unterminated.extend_from_slice(&[0xa3, 0x06, 0x08, 0x01]);
        assert!(decode_number_format(&unterminated, options()).is_err());

        let mut unmatched_end = native_number_format(2, 0, false);
        unmatched_end.extend_from_slice(&[0xa4, 0x06]);
        assert!(decode_number_format(&unmatched_end, options()).is_err());

        let unknown_group = [0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06];
        let mut source = native_number_format(2, 0, false);
        source.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]); // overlong unknown scalar 1
        source.extend_from_slice(&unknown_group);
        let (snapshot, report) =
            decode_number_format_with_report(&source, options()).expect("unknown extension");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(report.work_bytes(), source.len() * 2);

        let write = NumberFormatWrite::new(30, 3, true);
        let output = rewrite_number_format(&source, write, options()).expect("rewrite");
        assert!(output.bytes().ends_with(&[
            0xa0, 0x06, 0x81, 0x00, 0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06,
        ]));
        let rewritten = decode_number_format(output.bytes(), options()).expect("rewritten");
        assert_eq!(NumberFormatWrite::from_snapshot(rewritten), write);

        let no_op = rewrite_number_format(
            output.bytes(),
            NumberFormatWrite::from_snapshot(rewritten),
            options(),
        )
        .expect("no-op rewrite");
        assert_eq!(no_op.bytes(), output.bytes());
    }

    #[test]
    fn number_format_rewrite_rejects_varint_growth_within_original_message_ceiling() {
        // `decimal_places = 0` uses a one-byte varint. Automatic decimal
        // places (`253`) uses two bytes, so the source-preserving candidate is
        // one byte larger. Keep a large opaque extension in the source so the
        // caller's message ceiling is above the small-format threshold; this
        // exercises candidate preflight rather than the source scan.
        let mut source = native_number_format(0, 0, false);
        source.extend_from_slice(&[0xa2, 0x06, 0x80, 0x02]); // unknown field 100, 256 bytes
        source.resize(source.len() + 256, 0xde);
        let mut bounded = options();
        bounded.max_message_bytes = source.len();

        assert!(decode_number_format(&source, bounded).is_ok());
        let error = prepare_number_format_rewrite(
            &source,
            NumberFormatWrite::new(NATIVE_AUTOMATIC_DECIMAL_PLACES, 0, false),
            bounded,
        )
        .expect_err("encoded growth must stay within the caller ceiling");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::InputBytes { observed, maximum })
                if observed == source.len() + 1 && maximum == source.len()
        ));
    }

    #[test]
    fn prepared_number_format_execution_is_finite_and_measured() {
        let source = native_number_format(2, 0, false);
        let write = NumberFormatWrite::new(30, 3, true);
        let prepared = prepare_number_format_rewrite(&source, write, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        assert!(requirements.output_bytes() > 0);
        assert!(requirements.fields() >= 8);
        assert!(requirements.work_bytes() >= requirements.output_bytes());
        assert_eq!(requirements.allocations(), 1);
        assert_eq!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements))
                .expect("exact execution")
                .bytes()
                .len(),
            requirements.output_bytes()
        );

        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1),
                )
                .expect_err("output minus one")
                .resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_fields(requirements.fields() - 1),
                )
                .expect_err("fields minus one")
                .resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1),
                )
                .expect_err("work minus one")
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
        assert!(matches!(
            prepared
                .execute(RewriteExecutionLimits::exact(requirements).with_allocations(0))
                .expect_err("allocation refused")
                .resource_limit(),
            Some(DecodeLimit::Allocation { .. })
        ));
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_retained_bytes(requirements.retained_bytes() - 1),
                )
                .expect_err("retained minus one")
                .resource_limit(),
            Some(DecodeLimit::Retained { .. })
        ));
    }

    #[test]
    fn canonical_number_format_append_roundtrips_through_lazy_projection() {
        let write = NumberFormatWrite::new(NATIVE_AUTOMATIC_DECIMAL_PLACES, 1, true);
        let prepared = prepare_number_format_write(write, options()).expect("prepare append");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("append");
        let snapshot = decode_number_format(output.bytes(), options()).expect("read append");
        assert_eq!(NumberFormatWrite::from_snapshot(snapshot), write);
        assert_eq!(
            output.bytes(),
            &[0x08, 0x80, 0x02, 0x10, 0xfd, 0x01, 0x20, 0x01, 0x28, 0x01]
        );
        assert_eq!(
            canonical_number_format(write, options())
                .expect("one-shot")
                .bytes(),
            output.bytes()
        );
    }

    #[test]
    fn number_format_rejects_limits_above_the_buffa_hard_ceiling() {
        let source = native_number_format(2, 0, false);
        let mut invalid = options();
        invalid.max_message_bytes = usize::MAX;
        assert!(matches!(
            decode_number_format(&source, invalid)
                .expect_err("oversized Buffa configuration")
                .resource_limit(),
            Some(DecodeLimit::InputBytes { observed, maximum })
                if observed == usize::MAX
                    && maximum == usize::try_from(buffa::MAX_MESSAGE_BYTES)
                        .expect("Buffa ceiling fits usize")
        ));
    }

    fn native_percentage_format(decimal_places: u32, negative_style: u32, show: bool) -> Vec<u8> {
        let mut source = Vec::new();
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_TYPE_FIELD,
            u64::from(NATIVE_PERCENTAGE_FORMAT_TYPE),
        )
        .expect("type");
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_DECIMAL_PLACES_FIELD,
            u64::from(decimal_places),
        )
        .expect("decimal places");
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_NEGATIVE_STYLE_FIELD,
            u64::from(negative_style),
        )
        .expect("negative style");
        emit_varint_field(
            &mut source,
            NUMBER_FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(show),
        )
        .expect("thousands separator");
        source
    }

    fn native_currency_format(
        decimal_places: u32,
        currency_code: &str,
        negative_style: u32,
        show: bool,
        accounting: bool,
    ) -> Vec<u8> {
        let mut source = Vec::new();
        emit_varint_field(
            &mut source,
            FORMAT_TYPE_FIELD,
            u64::from(NATIVE_CURRENCY_FORMAT_TYPE),
        )
        .expect("type");
        emit_varint_field(
            &mut source,
            FORMAT_DECIMAL_PLACES_FIELD,
            u64::from(decimal_places),
        )
        .expect("decimal places");
        emit_len_field(
            &mut source,
            FORMAT_CURRENCY_CODE_FIELD,
            currency_code.as_bytes(),
        )
        .expect("currency code");
        emit_varint_field(
            &mut source,
            FORMAT_NEGATIVE_STYLE_FIELD,
            u64::from(negative_style),
        )
        .expect("negative style");
        emit_varint_field(
            &mut source,
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(show),
        )
        .expect("thousands separator");
        emit_varint_field(
            &mut source,
            FORMAT_USE_ACCOUNTING_STYLE_FIELD,
            u64::from(accounting),
        )
        .expect("accounting style");
        source
    }

    #[test]
    fn currency_format_native_fixture_is_strict_and_buffa_lazy() {
        let source = native_currency_format(NATIVE_AUTOMATIC_DECIMAL_PLACES, "EUR", 3, true, true);
        let (snapshot, report) =
            currency_codec::decode_currency_format_with_report(&source, options())
                .expect("currency format");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_CURRENCY_FORMAT_TYPE);
        assert_eq!(
            snapshot.decimal_places(),
            Some(NATIVE_AUTOMATIC_DECIMAL_PLACES)
        );
        assert_eq!(snapshot.currency_code(), Some("EUR"));
        assert_eq!(snapshot.negative_style(), Some(3));
        assert_eq!(snapshot.show_thousands_separator(), Some(true));
        assert_eq!(snapshot.use_accounting_style(), Some(true));
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.fields(), 6);
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.text_bytes(), 3);
        assert_eq!(report.allocations(), 0);
    }

    #[test]
    fn currency_format_accepts_only_native_domains_and_complete_fields() {
        for decimal_places in [
            0,
            MAX_NUMBER_DECIMAL_PLACES,
            NATIVE_AUTOMATIC_DECIMAL_PLACES,
        ] {
            let source = native_currency_format(decimal_places, "USD", 0, false, false);
            assert!(currency_codec::decode_currency_format(&source, options()).is_ok());
        }
        for negative_style in 0..=3 {
            let source = native_currency_format(2, "USD", negative_style, false, false);
            assert!(currency_codec::decode_currency_format(&source, options()).is_ok());
        }
        for show in [false, true] {
            let source = native_currency_format(2, "USD", 0, show, false);
            assert!(currency_codec::decode_currency_format(&source, options()).is_ok());
        }
        for accounting in [false, true] {
            let source = native_currency_format(2, "USD", 0, false, accounting);
            assert!(currency_codec::decode_currency_format(&source, options()).is_ok());
        }
        for decimal_places in [31, 254] {
            let source = native_currency_format(decimal_places, "USD", 0, false, false);
            assert!(currency_codec::decode_currency_format(&source, options()).is_err());
        }
        let source = native_currency_format(2, "USD", 4, false, false);
        assert!(currency_codec::decode_currency_format(&source, options()).is_err());
        for code in ["usd", "US", "EURO", "€UR"] {
            let source = native_currency_format(2, code, 0, false, false);
            assert!(currency_codec::decode_currency_format(&source, options()).is_err());
        }

        let complete = native_currency_format(2, "USD", 0, false, false);
        let fields = [
            &[0x08, 0x81, 0x02][..],
            &[0x10, 0x02][..],
            &[0x1a, 0x03, b'U', b'S', b'D'][..],
            &[0x20, 0x00][..],
            &[0x28, 0x00][..],
            &[0x30, 0x00][..],
        ];
        for omitted in 0..fields.len() {
            let mut source = Vec::new();
            for (index, field) in fields.iter().enumerate() {
                if index != omitted {
                    source.extend_from_slice(field);
                }
            }
            assert!(currency_codec::decode_currency_format(&source, options()).is_err());
        }
        let mut invalid_bool = complete.clone();
        let last = invalid_bool.len() - 1;
        *invalid_bool.get_mut(last).expect("bool byte") = 2;
        assert!(currency_codec::decode_currency_format(&invalid_bool, options()).is_err());

        let mut incompatible = complete;
        incompatible.extend_from_slice(&[0x38, 0x01]);
        assert!(currency_codec::decode_currency_format(&incompatible, options()).is_err());
    }

    #[test]
    fn currency_format_rewrite_preserves_unknown_source_and_prepared_accounting() {
        let mut source = native_currency_format(2, "USD", 0, false, false);
        let unknown = [
            0xa0, 0x06, 0x81, 0x00, // unknown scalar 100, overlong value spelling
            0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06, // unknown balanced group
        ];
        source.extend_from_slice(&unknown);
        let write = currency_codec::CurrencyFormatWrite::new("JPY", 30, 3, true, true);
        let prepared = currency_codec::prepare_currency_format_rewrite(&source, write, options())
            .expect("prepare currency rewrite");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("execute currency rewrite");
        assert_eq!(output.bytes().len(), requirements.output_bytes());
        assert_eq!(output.report().fields(), requirements.fields());
        assert_eq!(output.report().work_bytes(), requirements.work_bytes());
        assert!(output.bytes().ends_with(&unknown));
        let snapshot = currency_codec::decode_currency_format(output.bytes(), options())
            .expect("rewritten currency");
        assert_eq!(
            currency_codec::CurrencyFormatWrite::from_snapshot(snapshot),
            write
        );

        let no_op = currency_codec::rewrite_currency_format(
            output.bytes(),
            currency_codec::CurrencyFormatWrite::from_snapshot(snapshot),
            options(),
        )
        .expect("currency no-op rewrite");
        assert_eq!(no_op.bytes(), output.bytes());
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1)
                )
                .expect_err("output ceiling")
                .resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
        assert!(matches!(
            prepared
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_work_bytes(requirements.work_bytes() - 1)
                )
                .expect_err("work ceiling")
                .resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
    }

    #[test]
    fn currency_format_canonical_append_has_exact_wire_and_report() {
        let write = currency_codec::CurrencyFormatWrite::new("EUR", 2, 3, true, true);
        let prepared = currency_codec::prepare_currency_format_write(write, options())
            .expect("prepare currency append");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("currency append");
        assert_eq!(
            output.bytes(),
            &[
                0x08, 0x81, 0x02, // format_type = 257
                0x10, 0x02, // decimal_places = 2
                0x1a, 0x03, b'E', b'U', b'R', // currency_code
                0x20, 0x03, // negative_style = red parentheses
                0x28, 0x01, // show_thousands_separator = true
                0x30, 0x01, // use_accounting_style = true
            ]
        );
        assert_eq!(requirements.fields(), 6);
        assert_eq!(requirements.work_bytes(), output.bytes().len() * 3);
        let (_, report) =
            currency_codec::decode_currency_format_with_report(output.bytes(), options())
                .expect("decode currency append");
        assert_eq!(report.work_bytes(), output.bytes().len() * 2);
        assert_eq!(report.fields(), requirements.fields());
        assert_eq!(report.text_bytes(), requirements.text_bytes());
        assert_eq!(
            currency_codec::canonical_currency_format(write, options())
                .expect("one-shot currency")
                .bytes(),
            output.bytes()
        );
    }

    #[test]
    fn percentage_format_is_type_258_and_isolated_from_number() {
        let source = [
            0x08, 0x82, 0x02, // format_type = 258
            0x10, 0xfd, 0x01, // decimal_places = automatic (253)
            0x20, 0x03, // negative_style = red parentheses
            0x28, 0x01, // show_thousands_separator = true
        ];
        let (snapshot, report) =
            percentage_codec::decode_percentage_format_with_report(&source, options())
                .expect("percentage format");
        assert_eq!(snapshot.raw(), source.as_slice());
        assert_eq!(snapshot.format_type(), NATIVE_PERCENTAGE_FORMAT_TYPE);
        assert_eq!(snapshot.decimal_places(), NATIVE_AUTOMATIC_DECIMAL_PLACES);
        assert_eq!(snapshot.negative_style(), 3);
        assert!(snapshot.show_thousands_separator());
        assert_eq!(report.input_bytes(), source.len());
        assert_eq!(report.fields(), 4);
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.allocations(), 0);
        assert!(decode_number_format(&source, options()).is_err());

        let number = native_number_format(2, 0, false);
        assert!(percentage_codec::decode_percentage_format(&number, options()).is_err());
    }

    #[test]
    fn percentage_format_accepts_native_domains_only() {
        for decimal_places in [
            0,
            percentage_codec::MAX_PERCENTAGE_DECIMAL_PLACES,
            NATIVE_AUTOMATIC_DECIMAL_PLACES,
        ] {
            let source = native_percentage_format(decimal_places, 0, false);
            assert!(percentage_codec::decode_percentage_format(&source, options()).is_ok());
        }
        for negative_style in 0..=3 {
            let source = native_percentage_format(2, negative_style, false);
            assert!(percentage_codec::decode_percentage_format(&source, options()).is_ok());
        }
        for show in [false, true] {
            let source = native_percentage_format(2, 0, show);
            assert!(percentage_codec::decode_percentage_format(&source, options()).is_ok());
        }
        for decimal_places in [31, 254] {
            let source = native_percentage_format(decimal_places, 0, false);
            assert!(percentage_codec::decode_percentage_format(&source, options()).is_err());
        }
        let source = native_percentage_format(2, 4, false);
        assert!(percentage_codec::decode_percentage_format(&source, options()).is_err());
        let source = native_percentage_format(2, 0, false);
        let mut invalid_bool = source;
        *invalid_bool.last_mut().expect("bool byte") = 2;
        assert!(percentage_codec::decode_percentage_format(&invalid_bool, options()).is_err());
    }

    #[test]
    fn percentage_format_rejects_malformed_and_known_fields_but_preserves_unknowns() {
        let fields = [
            &[0x08, 0x82, 0x02][..],
            &[0x10, 0x02][..],
            &[0x20, 0x00][..],
            &[0x28, 0x00][..],
        ];
        for omitted in 0..fields.len() {
            let mut source = Vec::new();
            for (index, field) in fields.iter().enumerate() {
                if index != omitted {
                    source.extend_from_slice(field);
                }
            }
            assert!(percentage_codec::decode_percentage_format(&source, options()).is_err());
        }

        let mut duplicate = native_percentage_format(2, 0, false);
        duplicate.extend_from_slice(&[0x08, 0x82, 0x02]);
        assert!(percentage_codec::decode_percentage_format(&duplicate, options()).is_err());

        let wrong_wire = [
            0x0a, 0x01, 0x00, // field 1 encoded as length-delimited
            0x10, 0x02, 0x20, 0x00, 0x28, 0x00,
        ];
        assert!(percentage_codec::decode_percentage_format(&wrong_wire, options()).is_err());

        for incompatible in [[0x1a, 0x00], [0x30, 0x01], [0x72, 0x00]] {
            let mut source = native_percentage_format(2, 0, false);
            source.extend_from_slice(&incompatible);
            assert!(percentage_codec::decode_percentage_format(&source, options()).is_err());
        }

        let mut unterminated = native_percentage_format(2, 0, false);
        unterminated.extend_from_slice(&[0xa3, 0x06, 0x08, 0x01]);
        assert!(percentage_codec::decode_percentage_format(&unterminated, options()).is_err());

        let unknown_group = [0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06];
        let mut source = native_percentage_format(2, 0, false);
        source.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
        source.extend_from_slice(&unknown_group);
        let snapshot = percentage_codec::decode_percentage_format(&source, options())
            .expect("unknown extension");
        assert_eq!(snapshot.raw(), source.as_slice());

        let write = percentage_codec::PercentageFormatWrite::new(30, 3, true);
        let output = percentage_codec::rewrite_percentage_format(&source, write, options())
            .expect("percentage rewrite");
        assert!(output.bytes().ends_with(&[
            0xa0, 0x06, 0x81, 0x00, 0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06,
        ]));
        let rewritten = percentage_codec::decode_percentage_format(output.bytes(), options())
            .expect("rewritten percentage");
        assert_eq!(
            percentage_codec::PercentageFormatWrite::from_snapshot(rewritten),
            write
        );
        assert!(
            rewrite_number_format(
                source.as_slice(),
                NumberFormatWrite::new(3, 0, false),
                options()
            )
            .is_err()
        );
    }

    #[test]
    fn prepared_percentage_format_limits_and_canonical_bytes_are_exact() {
        let write =
            percentage_codec::PercentageFormatWrite::new(NATIVE_AUTOMATIC_DECIMAL_PLACES, 1, true);
        let prepared = percentage_codec::prepare_percentage_format_write(write, options())
            .expect("prepare percentage append");
        let requirements = prepared.execution_requirements();
        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("percentage append");
        assert_eq!(
            output.bytes(),
            &[0x08, 0x82, 0x02, 0x10, 0xfd, 0x01, 0x20, 0x01, 0x28, 0x01]
        );
        let snapshot = percentage_codec::decode_percentage_format(output.bytes(), options())
            .expect("read percentage append");
        assert_eq!(
            percentage_codec::PercentageFormatWrite::from_snapshot(snapshot),
            write
        );
        assert_eq!(requirements.output_bytes(), output.bytes().len());
        assert_eq!(requirements.fields(), 4);
        assert_eq!(requirements.work_bytes(), output.bytes().len() * 3);
        assert_eq!(requirements.allocations(), 1);
        assert!(matches!(
            percentage_codec::prepare_percentage_format_write(write, options())
                .expect("prepare percentage limits")
                .execute(
                    RewriteExecutionLimits::exact(requirements)
                        .with_output_bytes(requirements.output_bytes() - 1),
                )
                .expect_err("percentage output ceiling")
                .resource_limit(),
            Some(DecodeLimit::OutputBytes { .. })
        ));
    }
}
