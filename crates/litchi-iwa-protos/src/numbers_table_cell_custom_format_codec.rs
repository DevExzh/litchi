//! Strict source-preserving Numbers custom display-format protocol layer.
//!
//! The native custom-format registry is a pair of parallel repeated fields:
//! UUIDs and `TSK.CustomFormatArchive` values.  A custom archive contains
//! nested `FormatStructArchive` values and (for custom Number formats) an
//! ordered condition list.  This module keeps those collections and nested
//! records borrowed from the caller's source bytes.  A private Buffa lazy
//! projection checks the selected scalar/bytes shapes after the handwritten
//! scanner has completed its bounded wire preflight; no generated type is
//! part of the public API or owns a rewrite.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The wire helpers are intentionally kept beside the snapshots and writers they serve."
)]
#![allow(
    clippy::module_name_repetitions,
    reason = "Public names mirror the native archive family and existing codec conventions."
)]

use core::{fmt, mem::size_of, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_numbers_table_cell_custom_format_generated::LitchiIwaNumbersTableCellCustomFormatProjection as projection;

/// Message identifier used by the Numbers custom-format registry.
pub const CUSTOM_FORMAT_REGISTRY_MESSAGE_TYPE: u32 = 222;

/// Alias retaining the archive-oriented name used by existing callers.
pub const CUSTOM_FORMAT_LIST_MESSAGE_TYPE: u32 = CUSTOM_FORMAT_REGISTRY_MESSAGE_TYPE;

/// Field on `TN.DocumentArchive` that roots the document-scoped registry.
pub const CUSTOM_FORMAT_REGISTRY_REFERENCE_FIELD: u32 = 9;

/// `TST.TableDataList.ListType::CUSTOM_FORMAT`, distinct from the registry's
/// native object/message type.
pub const CUSTOM_FORMAT_LIST_KIND: u32 = 6;

/// Native custom Number format discriminator.
pub const NATIVE_CUSTOM_NUMBER_FORMAT_TYPE: u32 = 270;

/// Native custom Text format discriminator.
pub const NATIVE_CUSTOM_TEXT_FORMAT_TYPE: u32 = 271;

/// Native custom Date & Time format discriminator.
pub const NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE: u32 = 272;

/// Native sentinel stored in `fraction_accuracy` for custom patterns.
pub const NATIVE_CUSTOM_FRACTION_SENTINEL: u32 = u32::MAX - 2;

/// Maximum UTF-8 bytes accepted for one registry name.
pub const MAX_CUSTOM_NAME_BYTES: usize = 255;

/// Maximum UTF-8 bytes accepted for one custom pattern.
pub const MAX_CUSTOM_PATTERN_BYTES: usize = 4 * 1_024;

/// Maximum ordered conditions accepted for one custom Number archive.
pub const MAX_CUSTOM_CONDITIONS: usize = 32;

const FORMAT_TYPE_FIELD: u32 = 1;
const FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD: u32 = 5;
const FORMAT_USE_ACCOUNTING_STYLE_FIELD: u32 = 6;
const FORMAT_FRACTION_ACCURACY_FIELD: u32 = 11;
const FORMAT_CUSTOM_FORMAT_STRING_FIELD: u32 = 18;
const FORMAT_SCALE_FACTOR_FIELD: u32 = 19;
const FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD: u32 = 20;
const FORMAT_DECIMAL_WIDTH_FIELD: u32 = 27;
const FORMAT_MIN_INTEGER_WIDTH_FIELD: u32 = 28;
const FORMAT_NONSPACE_INTEGER_DIGITS_FIELD: u32 = 29;
const FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD: u32 = 30;
const FORMAT_INDEX_FROM_RIGHT_FIELD: u32 = 31;
const FORMAT_HASH_DECIMAL_DIGITS_FIELD: u32 = 34;
const FORMAT_TOTAL_DECIMAL_DIGITS_FIELD: u32 = 35;
const FORMAT_IS_COMPLEX_FIELD: u32 = 36;
const FORMAT_CONTAINS_INTEGER_TOKEN_FIELD: u32 = 37;
const FORMAT_MAX_KNOWN_FIELD: u32 = 45;

const LIST_UUIDS_FIELD: u32 = 1;
const LIST_CUSTOM_FORMATS_FIELD: u32 = 2;
const ARCHIVE_NAME_FIELD: u32 = 1;
const ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD: u32 = 2;
const ARCHIVE_DEFAULT_FORMAT_FIELD: u32 = 3;
const ARCHIVE_CONDITIONS_FIELD: u32 = 4;
const ARCHIVE_FORMAT_TYPE_FIELD: u32 = 5;
const CONDITION_TYPE_FIELD: u32 = 1;
const CONDITION_VALUE_FIELD: u32 = 2;
const CONDITION_FORMAT_FIELD: u32 = 3;
const CONDITION_VALUE_DBL_FIELD: u32 = 4;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;

const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;
const MAX_RECURSION: u32 = 64;

/// Finite limits for one custom-format decode or rewrite operation.
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
    /// Construct explicit finite source, output, field, work, nesting,
    /// registry-entry, and text ceilings.
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

    /// Derive conservative finite limits from one source payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self::new(
            bytes,
            bytes.checked_mul(2).unwrap_or(usize::MAX),
            bytes.checked_mul(8).unwrap_or(usize::MAX).max(1),
            bytes.checked_mul(128).unwrap_or(usize::MAX).max(1),
            MAX_RECURSION,
            bytes,
            bytes,
            bytes,
        )
    }

    /// Replace the output-byte ceiling.
    #[must_use]
    pub const fn with_max_output_bytes(mut self, maximum: usize) -> Self {
        self.max_output_bytes = maximum;
        self
    }

    /// Replace the maximum number of UUID registry references.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    /// Replace the maximum number of custom archives and conditions.
    #[must_use]
    pub const fn with_max_items(mut self, maximum: usize) -> Self {
        self.max_items = maximum;
        self
    }

    /// Alias for [`Self::with_max_items`].
    #[must_use]
    pub const fn with_max_formats(self, maximum: usize) -> Self {
        self.with_max_items(maximum)
    }

    /// Alias for [`Self::with_max_items`].
    #[must_use]
    pub const fn with_max_conditions(self, maximum: usize) -> Self {
        self.with_max_items(maximum)
    }

    /// Replace the selected text-byte ceiling.
    #[must_use]
    pub const fn with_max_text_bytes(mut self, maximum: usize) -> Self {
        self.max_text_bytes = maximum;
        self
    }
}

/// Typed finite resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// The source or Buffa message ceiling was exceeded.
    InputBytes { observed: usize, maximum: usize },
    /// A candidate output exceeded its configured ceiling.
    OutputBytes { observed: usize, maximum: usize },
    /// Wire records exceeded the aggregate field ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus Buffa inspected bytes exceeded the work ceiling.
    Work { observed: usize, maximum: usize },
    /// Nested message/group depth exceeded the recursion ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// UUID registry entries exceeded the ceiling.
    References { observed: usize, maximum: usize },
    /// Custom archives exceeded the ceiling.
    Formats { observed: usize, maximum: usize },
    /// Conditions exceeded the ceiling.
    Conditions { observed: usize, maximum: usize },
    /// Selected UTF-8 bytes exceeded the ceiling.
    Text { observed: usize, maximum: usize },
    /// A fallible output/scratch reservation was refused.
    Allocation { requested: usize },
    /// Retained source plus candidate bytes exceeded the execution ceiling.
    Retained { observed: usize, maximum: usize },
    /// Temporary scratch bytes exceeded the execution ceiling.
    Scratch { observed: usize, maximum: usize },
}

/// Strict custom-format codec failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
}

impl DecodeError {
    const fn invalid() -> Self {
        Self { limit: None }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self { limit: Some(limit) }
    }

    /// Return the typed resource failure, when this is a limit error.
    #[must_use]
    pub const fn resource_limit(self) -> Option<DecodeLimit> {
        self.limit
    }

    /// Return a refused allocation size, if applicable.
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
        formatter.write_str("invalid Numbers custom-format payload")
    }
}

impl std::error::Error for DecodeError {}

/// Aggregate resource consumption for one strict operation.
///
/// Wire scanning itself is allocation-free, but the private Buffa lazy
/// projections use a `Vec` for each repeated bytes field.  `allocations` and
/// `scratch_bytes` include those temporary repeated-view collections using a
/// conservative backing-capacity bound.  The source and all selected text
/// remain borrowed; no generated owned message crosses this API.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    formats: usize,
    conditions: usize,
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
    pub const fn formats(self) -> usize {
        self.formats
    }

    #[must_use]
    pub const fn conditions(self) -> usize {
        self.conditions
    }

    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Logical private Buffa repeated-view allocations performed by the
    /// strict operation.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source plus candidate bytes retained across the operation.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Conservative temporary backing storage for private Buffa repeated
    /// views.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// A source-borrowed native UUID value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Uuid {
    lower: u64,
    upper: u64,
}

impl Uuid {
    /// Construct a UUID value for a canonical append or rewrite.
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Self {
        Self { lower, upper }
    }

    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// Selected, source-borrowed `FormatStructArchive` values used by a custom
/// format entry.  Every selected field is required by the native custom
/// pattern shape; unselected native fields are rejected as cross-family data.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FormatStructSnapshot<'source> {
    source: &'source [u8],
    format_type: u32,
    show_thousands_separator: bool,
    use_accounting_style: bool,
    fraction_accuracy: u32,
    custom_format_string: &'source str,
    scale_factor: f64,
    requires_fraction_replacement: bool,
    decimal_width: u32,
    min_integer_width: u32,
    num_nonspace_integer_digits: u32,
    num_nonspace_decimal_digits: u32,
    index_from_right_last_integer: u32,
    num_hash_decimal_digits: u32,
    total_num_decimal_digits: u32,
    is_complex: bool,
    contains_integer_token: bool,
}

impl<'source> FormatStructSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }

    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.show_thousands_separator
    }

    #[must_use]
    pub const fn use_accounting_style(self) -> bool {
        self.use_accounting_style
    }

    #[must_use]
    pub const fn fraction_accuracy(self) -> u32 {
        self.fraction_accuracy
    }

    #[must_use]
    pub const fn custom_format_string(self) -> &'source str {
        self.custom_format_string
    }

    #[must_use]
    pub const fn pattern(self) -> &'source str {
        self.custom_format_string
    }

    #[must_use]
    pub const fn scale_factor(self) -> f64 {
        self.scale_factor
    }

    #[must_use]
    pub const fn requires_fraction_replacement(self) -> bool {
        self.requires_fraction_replacement
    }

    #[must_use]
    pub const fn decimal_width(self) -> u32 {
        self.decimal_width
    }

    #[must_use]
    pub const fn min_integer_width(self) -> u32 {
        self.min_integer_width
    }

    #[must_use]
    pub const fn num_nonspace_integer_digits(self) -> u32 {
        self.num_nonspace_integer_digits
    }

    #[must_use]
    pub const fn num_nonspace_decimal_digits(self) -> u32 {
        self.num_nonspace_decimal_digits
    }

    #[must_use]
    pub const fn index_from_right_last_integer(self) -> u32 {
        self.index_from_right_last_integer
    }

    #[must_use]
    pub const fn num_hash_decimal_digits(self) -> u32 {
        self.num_hash_decimal_digits
    }

    #[must_use]
    pub const fn total_num_decimal_digits(self) -> u32 {
        self.total_num_decimal_digits
    }

    #[must_use]
    pub const fn is_complex(self) -> bool {
        self.is_complex
    }

    #[must_use]
    pub const fn contains_integer_token(self) -> bool {
        self.contains_integer_token
    }
}

/// Values accepted by the custom-format `FormatStructArchive` writer.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FormatStructWrite<'source> {
    format_type: u32,
    show_thousands_separator: bool,
    use_accounting_style: bool,
    fraction_accuracy: u32,
    custom_format_string: &'source str,
    scale_factor: f64,
    requires_fraction_replacement: bool,
    decimal_width: u32,
    min_integer_width: u32,
    num_nonspace_integer_digits: u32,
    num_nonspace_decimal_digits: u32,
    index_from_right_last_integer: u32,
    num_hash_decimal_digits: u32,
    total_num_decimal_digits: u32,
    is_complex: bool,
    contains_integer_token: bool,
}

impl<'source> FormatStructWrite<'source> {
    /// Construct the canonical metadata for a custom pattern.
    #[must_use]
    pub fn new(format_type: u32, pattern: &'source str) -> Self {
        let contains_integer_token = format_type == NATIVE_CUSTOM_NUMBER_FORMAT_TYPE
            && pattern
                .chars()
                .any(|character| matches!(character, '#' | '0'));
        Self {
            format_type,
            show_thousands_separator: format_type == NATIVE_CUSTOM_NUMBER_FORMAT_TYPE
                && pattern.contains(','),
            use_accounting_style: false,
            fraction_accuracy: NATIVE_CUSTOM_FRACTION_SENTINEL,
            custom_format_string: pattern,
            scale_factor: 1.0,
            requires_fraction_replacement: false,
            decimal_width: 0,
            min_integer_width: 0,
            num_nonspace_integer_digits: 0,
            num_nonspace_decimal_digits: 0,
            index_from_right_last_integer: 0,
            num_hash_decimal_digits: 0,
            total_num_decimal_digits: 0,
            is_complex: false,
            contains_integer_token,
        }
    }

    /// Copy selected semantic values from a decoded snapshot.
    #[must_use]
    pub const fn from_snapshot(snapshot: FormatStructSnapshot<'source>) -> Self {
        Self {
            format_type: snapshot.format_type,
            show_thousands_separator: snapshot.show_thousands_separator,
            use_accounting_style: snapshot.use_accounting_style,
            fraction_accuracy: snapshot.fraction_accuracy,
            custom_format_string: snapshot.custom_format_string,
            scale_factor: snapshot.scale_factor,
            requires_fraction_replacement: snapshot.requires_fraction_replacement,
            decimal_width: snapshot.decimal_width,
            min_integer_width: snapshot.min_integer_width,
            num_nonspace_integer_digits: snapshot.num_nonspace_integer_digits,
            num_nonspace_decimal_digits: snapshot.num_nonspace_decimal_digits,
            index_from_right_last_integer: snapshot.index_from_right_last_integer,
            num_hash_decimal_digits: snapshot.num_hash_decimal_digits,
            total_num_decimal_digits: snapshot.total_num_decimal_digits,
            is_complex: snapshot.is_complex,
            contains_integer_token: snapshot.contains_integer_token,
        }
    }

    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }

    #[must_use]
    pub const fn custom_format_string(self) -> &'source str {
        self.custom_format_string
    }

    #[must_use]
    pub const fn pattern(self) -> &'source str {
        self.custom_format_string
    }

    #[must_use]
    pub const fn show_thousands_separator(self) -> bool {
        self.show_thousands_separator
    }

    #[must_use]
    pub const fn use_accounting_style(self) -> bool {
        self.use_accounting_style
    }

    #[must_use]
    pub const fn fraction_accuracy(self) -> u32 {
        self.fraction_accuracy
    }

    #[must_use]
    pub const fn scale_factor(self) -> f64 {
        self.scale_factor
    }

    #[must_use]
    pub const fn requires_fraction_replacement(self) -> bool {
        self.requires_fraction_replacement
    }

    #[must_use]
    pub const fn decimal_width(self) -> u32 {
        self.decimal_width
    }

    #[must_use]
    pub const fn min_integer_width(self) -> u32 {
        self.min_integer_width
    }

    #[must_use]
    pub const fn num_nonspace_integer_digits(self) -> u32 {
        self.num_nonspace_integer_digits
    }

    #[must_use]
    pub const fn num_nonspace_decimal_digits(self) -> u32 {
        self.num_nonspace_decimal_digits
    }

    #[must_use]
    pub const fn index_from_right_last_integer(self) -> u32 {
        self.index_from_right_last_integer
    }

    #[must_use]
    pub const fn num_hash_decimal_digits(self) -> u32 {
        self.num_hash_decimal_digits
    }

    #[must_use]
    pub const fn total_num_decimal_digits(self) -> u32 {
        self.total_num_decimal_digits
    }

    #[must_use]
    pub const fn is_complex(self) -> bool {
        self.is_complex
    }

    #[must_use]
    pub const fn contains_integer_token(self) -> bool {
        self.contains_integer_token
    }

    #[must_use]
    pub const fn with_show_thousands_separator(mut self, value: bool) -> Self {
        self.show_thousands_separator = value;
        self
    }

    #[must_use]
    pub const fn with_use_accounting_style(mut self, value: bool) -> Self {
        self.use_accounting_style = value;
        self
    }

    #[must_use]
    pub const fn with_fraction_accuracy(mut self, value: u32) -> Self {
        self.fraction_accuracy = value;
        self
    }

    #[must_use]
    pub const fn with_scale_factor(mut self, value: f64) -> Self {
        self.scale_factor = value;
        self
    }

    #[must_use]
    pub const fn with_requires_fraction_replacement(mut self, value: bool) -> Self {
        self.requires_fraction_replacement = value;
        self
    }

    #[must_use]
    pub const fn with_decimal_width(mut self, value: u32) -> Self {
        self.decimal_width = value;
        self
    }

    #[must_use]
    pub const fn with_min_integer_width(mut self, value: u32) -> Self {
        self.min_integer_width = value;
        self
    }

    #[must_use]
    pub const fn with_num_nonspace_integer_digits(mut self, value: u32) -> Self {
        self.num_nonspace_integer_digits = value;
        self
    }

    #[must_use]
    pub const fn with_num_nonspace_decimal_digits(mut self, value: u32) -> Self {
        self.num_nonspace_decimal_digits = value;
        self
    }

    #[must_use]
    pub const fn with_index_from_right_last_integer(mut self, value: u32) -> Self {
        self.index_from_right_last_integer = value;
        self
    }

    #[must_use]
    pub const fn with_num_hash_decimal_digits(mut self, value: u32) -> Self {
        self.num_hash_decimal_digits = value;
        self
    }

    #[must_use]
    pub const fn with_total_num_decimal_digits(mut self, value: u32) -> Self {
        self.total_num_decimal_digits = value;
        self
    }

    #[must_use]
    pub const fn with_is_complex(mut self, value: bool) -> Self {
        self.is_complex = value;
        self
    }

    #[must_use]
    pub const fn with_contains_integer_token(mut self, value: bool) -> Self {
        self.contains_integer_token = value;
        self
    }
}

/// One source-borrowed custom Number condition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomConditionSnapshot<'source> {
    source: &'source [u8],
    condition_type: u32,
    condition_value: Option<f32>,
    condition_format: FormatStructSnapshot<'source>,
    condition_value_dbl: Option<f64>,
}

impl<'source> CustomConditionSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    #[must_use]
    pub const fn condition_type(self) -> u32 {
        self.condition_type
    }

    #[must_use]
    pub const fn condition_value(self) -> Option<f32> {
        self.condition_value
    }

    #[must_use]
    pub const fn condition_format(self) -> FormatStructSnapshot<'source> {
        self.condition_format
    }

    #[must_use]
    pub const fn condition_value_dbl(self) -> Option<f64> {
        self.condition_value_dbl
    }

    #[must_use]
    pub fn threshold(self) -> f64 {
        self.condition_value_dbl
            .unwrap_or_else(|| f64::from(self.condition_value.expect("validated condition value")))
    }
}

/// Values accepted by the custom condition writer.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomConditionWrite<'source> {
    condition_type: u32,
    condition_value: Option<f32>,
    condition_format: FormatStructWrite<'source>,
    condition_value_dbl: Option<f64>,
}

impl<'source> CustomConditionWrite<'source> {
    /// Construct one condition with exactly one optional threshold spelling.
    #[must_use]
    pub const fn new(
        condition_type: u32,
        condition_value: Option<f32>,
        condition_format: FormatStructWrite<'source>,
        condition_value_dbl: Option<f64>,
    ) -> Self {
        Self {
            condition_type,
            condition_value,
            condition_format,
            condition_value_dbl,
        }
    }

    /// Construct a canonical double-threshold condition.
    #[must_use]
    pub const fn with_double(
        condition_type: u32,
        threshold: f64,
        condition_format: FormatStructWrite<'source>,
    ) -> Self {
        Self::new(condition_type, None, condition_format, Some(threshold))
    }

    #[must_use]
    pub const fn condition_type(self) -> u32 {
        self.condition_type
    }

    #[must_use]
    pub const fn condition_value(self) -> Option<f32> {
        self.condition_value
    }

    #[must_use]
    pub const fn condition_format(self) -> FormatStructWrite<'source> {
        self.condition_format
    }

    #[must_use]
    pub const fn condition_value_dbl(self) -> Option<f64> {
        self.condition_value_dbl
    }
}

/// A source-borrowed custom-format archive.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomFormatSnapshot<'source> {
    source: &'source [u8],
    name: &'source str,
    format_type_pre_bnc: u32,
    default_format: FormatStructSnapshot<'source>,
    conditions_start: usize,
    conditions: usize,
    format_type: u32,
}

impl<'source> CustomFormatSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    #[must_use]
    pub const fn name(self) -> &'source str {
        self.name
    }

    #[must_use]
    pub const fn format_type_pre_bnc(self) -> u32 {
        self.format_type_pre_bnc
    }

    #[must_use]
    pub const fn default_format(self) -> FormatStructSnapshot<'source> {
        self.default_format
    }

    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }

    #[must_use]
    pub const fn condition_count(self) -> usize {
        self.conditions
    }

    /// Iterate conditions in their native wire order without materialising a
    /// vector.
    #[must_use]
    pub fn conditions(self) -> CustomConditionIter<'source> {
        CustomConditionIter {
            source: self.source,
            cursor: self.conditions_start,
            expected_format_type: self.format_type,
        }
    }

    /// Compatibility spelling for [`Self::conditions`].
    #[must_use]
    pub fn condition_iter(self) -> CustomConditionIter<'source> {
        self.conditions()
    }
}

/// Values accepted by the custom-format archive writer.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomFormatWrite<'source, 'conditions> {
    name: &'source str,
    format_type_pre_bnc: u32,
    default_format: FormatStructWrite<'source>,
    conditions: &'conditions [CustomConditionWrite<'source>],
    format_type: u32,
}

impl<'source, 'conditions> CustomFormatWrite<'source, 'conditions> {
    /// Construct a custom-format archive value.
    #[must_use]
    pub const fn new(
        name: &'source str,
        format_type: u32,
        default_format: FormatStructWrite<'source>,
        conditions: &'conditions [CustomConditionWrite<'source>],
    ) -> Self {
        Self {
            name,
            format_type_pre_bnc: format_type,
            default_format,
            conditions,
            format_type,
        }
    }

    /// Copy a decoded archive's semantic values.
    #[must_use]
    pub const fn from_snapshot(snapshot: CustomFormatSnapshot<'source>) -> Self {
        // The caller supplies the condition slice when it needs a complete
        // rewrite.  The empty slice is useful for the common Text/DateTime
        // shape and is rejected for a Number archive when conditions exist.
        Self::new(
            snapshot.name,
            snapshot.format_type,
            FormatStructWrite::from_snapshot(snapshot.default_format),
            &[],
        )
    }

    #[must_use]
    pub const fn name(self) -> &'source str {
        self.name
    }

    #[must_use]
    pub const fn format_type_pre_bnc(self) -> u32 {
        self.format_type_pre_bnc
    }

    #[must_use]
    pub const fn format_type(self) -> u32 {
        self.format_type
    }

    #[must_use]
    pub const fn default_format(self) -> FormatStructWrite<'source> {
        self.default_format
    }

    #[must_use]
    pub const fn conditions(self) -> &'conditions [CustomConditionWrite<'source>] {
        self.conditions
    }

    #[must_use]
    pub const fn with_conditions(
        mut self,
        conditions: &'conditions [CustomConditionWrite<'source>],
    ) -> Self {
        self.conditions = conditions;
        self
    }

    #[must_use]
    pub const fn with_name(mut self, name: &'source str) -> Self {
        self.name = name;
        self
    }
}

/// A source-borrowed custom-format registry entry.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomFormatListEntry<'source> {
    uuid: Uuid,
    custom_format: CustomFormatSnapshot<'source>,
}

impl<'source> CustomFormatListEntry<'source> {
    #[must_use]
    pub const fn uuid(self) -> Uuid {
        self.uuid
    }

    #[must_use]
    pub const fn custom_format(self) -> CustomFormatSnapshot<'source> {
        self.custom_format
    }

    #[must_use]
    pub const fn format(self) -> CustomFormatSnapshot<'source> {
        self.custom_format
    }
}

/// A source-borrowed custom-format registry.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomFormatListSnapshot<'source> {
    source: &'source [u8],
    uuid_count: usize,
    format_count: usize,
}

impl<'source> CustomFormatListSnapshot<'source> {
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    #[must_use]
    pub const fn uuid_count(self) -> usize {
        self.uuid_count
    }

    #[must_use]
    pub const fn custom_format_count(self) -> usize {
        self.format_count
    }

    #[must_use]
    pub const fn format_count(self) -> usize {
        self.format_count
    }

    /// Iterate UUIDs in native wire order without materialising a vector.
    #[must_use]
    pub fn uuids(self) -> UuidIter<'source> {
        UuidIter {
            source: self.source,
            cursor: 0,
        }
    }

    /// Iterate custom archives in native wire order without materialising a
    /// vector.
    #[must_use]
    pub fn custom_formats(self) -> CustomFormatIter<'source> {
        CustomFormatIter {
            source: self.source,
            cursor: 0,
        }
    }

    /// Compatibility spelling for [`Self::custom_formats`].
    #[must_use]
    pub fn formats(self) -> CustomFormatIter<'source> {
        self.custom_formats()
    }

    /// Iterate UUID/archive pairs by advancing two borrowed source routers.
    #[must_use]
    pub fn entries(self) -> CustomFormatListEntryIter<'source> {
        CustomFormatListEntryIter {
            uuids: self.uuids(),
            formats: self.custom_formats(),
        }
    }
}

/// A borrowed list writer.  The two slices must have equal length.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CustomFormatListWrite<'source, 'items> {
    uuids: &'items [Uuid],
    custom_formats: &'items [CustomFormatWrite<'source, 'items>],
}

impl<'source, 'items> CustomFormatListWrite<'source, 'items> {
    /// Construct a list writer from parallel UUID/archive slices.
    #[must_use]
    pub const fn new(
        uuids: &'items [Uuid],
        custom_formats: &'items [CustomFormatWrite<'source, 'items>],
    ) -> Self {
        Self {
            uuids,
            custom_formats,
        }
    }

    #[must_use]
    pub const fn uuids(self) -> &'items [Uuid] {
        self.uuids
    }

    #[must_use]
    pub const fn custom_formats(self) -> &'items [CustomFormatWrite<'source, 'items>] {
        self.custom_formats
    }
}

/// Prepared source-preserving custom-format rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCustomFormatRewrite<'source, 'conditions> {
    source: &'source [u8],
    write: CustomFormatWrite<'source, 'conditions>,
    output_bytes: usize,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedCustomFormatRewrite<'_, '_> {
    /// Return measured finite execution requirements.
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    /// Return requirements in report form.
    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    /// Emit, strictly read back, and publish a source-preserving candidate.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(self.output_bytes).map_err(|_| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: self.output_bytes,
            })
        })?;
        emit_custom_format_rewrite(&mut bytes, self.source, self.write)?;
        if bytes.len() != self.output_bytes {
            return Err(DecodeError::invalid());
        }
        let (snapshot, report) = decode_custom_format_with_report(&bytes, self.verify_options)?;
        if report.fields() != self.requirements.fields()
            || !custom_format_matches_write(snapshot, self.write)?
        {
            return Err(DecodeError::invalid());
        }
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical custom-format append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCustomFormatWrite<'source, 'conditions> {
    write: CustomFormatWrite<'source, 'conditions>,
    output_bytes: usize,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedCustomFormatWrite<'_, '_> {
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    /// Emit and strictly read back a canonical custom archive.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(self.output_bytes).map_err(|_| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: self.output_bytes,
            })
        })?;
        emit_custom_format_canonical(&mut bytes, self.write)?;
        if bytes.len() != self.output_bytes {
            return Err(DecodeError::invalid());
        }
        let (snapshot, report) = decode_custom_format_with_report(&bytes, self.verify_options)?;
        if report.fields() != self.requirements.fields()
            || !custom_format_matches_write(snapshot, self.write)?
        {
            return Err(DecodeError::invalid());
        }
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared source-preserving custom-format-list rewrite.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCustomFormatListRewrite<'source, 'items> {
    source: &'source [u8],
    write: CustomFormatListWrite<'source, 'items>,
    output_bytes: usize,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedCustomFormatListRewrite<'_, '_> {
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    /// Emit, strictly read back, and publish a registry rewrite.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(self.output_bytes).map_err(|_| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: self.output_bytes,
            })
        })?;
        emit_custom_format_list_rewrite(&mut bytes, self.source, self.write)?;
        if bytes.len() != self.output_bytes {
            return Err(DecodeError::invalid());
        }
        let (snapshot, report) =
            decode_custom_format_list_with_report(&bytes, self.verify_options)?;
        if report.fields() != self.requirements.fields()
            || !custom_format_list_matches_write(snapshot, self.write)?
        {
            return Err(DecodeError::invalid());
        }
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Prepared canonical custom-format-list append.
#[derive(Debug, Clone, Copy)]
pub struct PreparedCustomFormatListWrite<'source, 'items> {
    write: CustomFormatListWrite<'source, 'items>,
    output_bytes: usize,
    requirements: RewriteExecutionRequirements,
    verify_options: DecodeOptions,
}

impl PreparedCustomFormatListWrite<'_, '_> {
    #[must_use]
    pub const fn execution_requirements(self) -> RewriteExecutionRequirements {
        self.requirements
    }

    #[must_use]
    pub fn prepare_report(self) -> DecodeReport {
        report_from_requirements(self.requirements)
    }

    /// Emit and strictly read back a canonical registry payload.
    pub fn execute(self, limits: RewriteExecutionLimits) -> Result<RewriteOutput, DecodeError> {
        check_requirements(self.requirements, limits)?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(self.output_bytes).map_err(|_| {
            DecodeError::limited(DecodeLimit::Allocation {
                requested: self.output_bytes,
            })
        })?;
        emit_custom_format_list_canonical(&mut bytes, self.write)?;
        if bytes.len() != self.output_bytes {
            return Err(DecodeError::invalid());
        }
        let (snapshot, report) =
            decode_custom_format_list_with_report(&bytes, self.verify_options)?;
        if report.fields() != self.requirements.fields()
            || !custom_format_list_matches_write(snapshot, self.write)?
        {
            return Err(DecodeError::invalid());
        }
        Ok(RewriteOutput {
            bytes,
            report: report_from_requirements(self.requirements),
        })
    }
}

/// Requirements measured before rewrite execution.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionRequirements {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    formats: usize,
    conditions: usize,
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
    pub const fn formats(self) -> usize {
        self.formats
    }

    #[must_use]
    pub const fn conditions(self) -> usize {
        self.conditions
    }

    #[must_use]
    pub const fn text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Logical output and private Buffa repeated-view allocations required by
    /// the prepared operation.
    #[must_use]
    pub const fn allocations(self) -> usize {
        self.allocations
    }

    /// Source plus candidate bytes retained across the operation.
    #[must_use]
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }

    /// Conservative temporary backing storage for private Buffa repeated
    /// views.
    #[must_use]
    pub const fn scratch_bytes(self) -> usize {
        self.scratch_bytes
    }
}

/// Execution ceilings checked against prepared requirements.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewriteExecutionLimits {
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    formats: usize,
    conditions: usize,
    text_bytes: usize,
    allocations: usize,
    retained_bytes: usize,
    scratch_bytes: usize,
}

impl RewriteExecutionLimits {
    /// Set every execution ceiling to the measured requirements.
    #[must_use]
    pub const fn exact(requirements: RewriteExecutionRequirements) -> Self {
        Self {
            output_bytes: requirements.output_bytes,
            fields: requirements.fields,
            work_bytes: requirements.work_bytes,
            max_depth: requirements.max_depth,
            references: requirements.references,
            formats: requirements.formats,
            conditions: requirements.conditions,
            text_bytes: requirements.text_bytes,
            allocations: requirements.allocations,
            retained_bytes: requirements.retained_bytes,
            scratch_bytes: requirements.scratch_bytes,
        }
    }

    /// Set all ceilings to their largest representable values.
    #[must_use]
    pub const fn unrestricted() -> Self {
        Self {
            output_bytes: usize::MAX,
            fields: usize::MAX,
            work_bytes: usize::MAX,
            max_depth: u32::MAX,
            references: usize::MAX,
            formats: usize::MAX,
            conditions: usize::MAX,
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
    pub const fn with_formats(mut self, value: usize) -> Self {
        self.formats = value;
        self
    }

    #[must_use]
    pub const fn with_conditions(mut self, value: usize) -> Self {
        self.conditions = value;
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

/// Bytes emitted by a canonical or source-preserving rewrite.
#[derive(Debug, Clone, PartialEq, Eq)]
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

/// Decode one custom `FormatStructArchive` pattern value.
pub fn decode_format_struct(
    source: &[u8],
    options: DecodeOptions,
) -> Result<FormatStructSnapshot<'_>, DecodeError> {
    Ok(decode_format_struct_with_report(source, options)?.0)
}

/// Decode one custom format pattern and return exact resource use.
pub fn decode_format_struct_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(FormatStructSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let (snapshot, _) = scan_format_struct(source, 0, 0, None, &mut budget)?;
    buffa_format_parity(source, snapshot, &mut budget)?;
    Ok((snapshot, budget.finish(0)))
}

/// Decode one native `FormatStructArchive` used by a custom archive.
pub use decode_format_struct as decode_format_struct_archive;

/// Decode one custom-format archive.
pub fn decode_custom_format(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CustomFormatSnapshot<'_>, DecodeError> {
    Ok(decode_custom_format_with_report(source, options)?.0)
}

/// Decode one custom-format archive and return exact resource use.
pub fn decode_custom_format_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CustomFormatSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = scan_custom_format(source, 0, 0, &mut budget)?;
    buffa_custom_format_parity(source, snapshot, &mut budget)?;
    Ok((snapshot, budget.finish(0)))
}

/// Native archive spelling for [`decode_custom_format`].
pub use decode_custom_format as decode_custom_format_archive;

/// Decode one custom-format registry.
pub fn decode_custom_format_list(
    source: &[u8],
    options: DecodeOptions,
) -> Result<CustomFormatListSnapshot<'_>, DecodeError> {
    Ok(decode_custom_format_list_with_report(source, options)?.0)
}

/// Decode one custom-format registry and return exact resource use.
pub fn decode_custom_format_list_with_report(
    source: &[u8],
    options: DecodeOptions,
) -> Result<(CustomFormatListSnapshot<'_>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    let snapshot = scan_custom_format_list(source, 0, 0, &mut budget)?;
    if snapshot.uuid_count != snapshot.format_count {
        return Err(DecodeError::invalid());
    }
    validate_unique_uuids(source, snapshot.uuid_count, &mut budget)?;
    buffa_custom_format_list_parity(source, snapshot, &mut budget)?;
    Ok((snapshot, budget.finish(0)))
}

/// Native archive spelling for [`decode_custom_format_list`].
pub use decode_custom_format_list as decode_custom_format_list_archive;

/// Prepare a source-preserving archive rewrite.
pub fn prepare_custom_format_rewrite<'source, 'conditions>(
    source: &'source [u8],
    write: CustomFormatWrite<'source, 'conditions>,
    options: DecodeOptions,
) -> Result<PreparedCustomFormatRewrite<'source, 'conditions>, DecodeError> {
    validate_custom_format_write(write)?;
    let (snapshot, source_report) = decode_custom_format_with_report(source, options)?;
    if snapshot.condition_count() != write.conditions().len()
        || snapshot.format_type() != write.format_type()
    {
        return Err(DecodeError::invalid());
    }
    let output_bytes = custom_format_rewrite_len(source, write)?;
    let candidate_text_bytes = write_counts(write)?.text_bytes;
    let requirements = rewrite_requirements(
        source,
        source_report,
        output_bytes,
        source_report.formats(),
        source_report.conditions(),
        source_report.references(),
        candidate_text_bytes,
        custom_format_materialization(write.conditions().len())?,
    )?;
    check_options(requirements, options)?;
    Ok(PreparedCustomFormatRewrite {
        source,
        write,
        output_bytes,
        requirements,
        verify_options: options,
    })
}

/// Rewrite one source-preserving custom archive.
pub fn rewrite_custom_format(
    source: &[u8],
    write: CustomFormatWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_custom_format_rewrite(source, write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for table-cell custom archives.
pub use rewrite_custom_format as rewrite_table_cell_custom_format;

/// Prepare a canonical custom archive append.
pub fn prepare_custom_format_write<'source, 'conditions>(
    write: CustomFormatWrite<'source, 'conditions>,
    options: DecodeOptions,
) -> Result<PreparedCustomFormatWrite<'source, 'conditions>, DecodeError> {
    validate_custom_format_write(write)?;
    let output_bytes = custom_format_canonical_len(write)?;
    let counts = write_counts(write)?;
    let requirements = canonical_requirements(
        output_bytes,
        counts.fields,
        counts.work,
        counts.max_depth,
        0,
        1,
        counts.conditions,
        counts.text_bytes,
        output_bytes,
        custom_format_materialization(counts.conditions)?,
    )?;
    check_options(requirements, options)?;
    Ok(PreparedCustomFormatWrite {
        write,
        output_bytes,
        requirements,
        verify_options: options,
    })
}

/// Compatibility spelling for canonical custom archive append.
pub use prepare_custom_format_write as prepare_custom_format_append;

/// Emit one canonical custom archive.
pub fn canonical_custom_format(
    write: CustomFormatWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_custom_format_write(write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Prepare a source-preserving registry rewrite.  UUID/archive counts must
/// remain parallel and unchanged; append/remove operations use the canonical
/// list writer instead.
pub fn prepare_custom_format_list_rewrite<'source, 'items>(
    source: &'source [u8],
    write: CustomFormatListWrite<'source, 'items>,
    options: DecodeOptions,
) -> Result<PreparedCustomFormatListRewrite<'source, 'items>, DecodeError> {
    validate_custom_format_list_write(write)?;
    let (snapshot, source_report) = decode_custom_format_list_with_report(source, options)?;
    if snapshot.uuid_count() != write.uuids().len()
        || snapshot.format_count() != write.custom_formats().len()
    {
        return Err(DecodeError::invalid());
    }
    let output_bytes = custom_format_list_rewrite_len(source, write)?;
    let candidate_text_bytes = custom_format_list_write_text_bytes(write)?;
    let requirements = rewrite_requirements(
        source,
        source_report,
        output_bytes,
        source_report.formats(),
        source_report.conditions(),
        source_report.references(),
        candidate_text_bytes,
        custom_format_list_materialization(write)?,
    )?;
    check_options(requirements, options)?;
    Ok(PreparedCustomFormatListRewrite {
        source,
        write,
        output_bytes,
        requirements,
        verify_options: options,
    })
}

/// Rewrite one source-preserving custom-format registry.
pub fn rewrite_custom_format_list(
    source: &[u8],
    write: CustomFormatListWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_custom_format_list_rewrite(source, write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Prepare a canonical custom-format registry append/write.
pub fn prepare_custom_format_list_write<'source, 'items>(
    write: CustomFormatListWrite<'source, 'items>,
    options: DecodeOptions,
) -> Result<PreparedCustomFormatListWrite<'source, 'items>, DecodeError> {
    validate_custom_format_list_write(write)?;
    let output_bytes = custom_format_list_canonical_len(write)?;
    let mut fields = 0usize;
    let mut work = 0usize;
    let mut formats = 0usize;
    let mut conditions = 0usize;
    let mut text = 0usize;
    let mut max_depth = 0u32;
    for custom_format in write.custom_formats() {
        let counts = write_counts(*custom_format)?;
        fields = fields
            .checked_add(counts.fields)
            .ok_or_else(DecodeError::invalid)?;
        work = work
            .checked_add(counts.work)
            .ok_or_else(DecodeError::invalid)?;
        conditions = conditions
            .checked_add(counts.conditions)
            .ok_or_else(DecodeError::invalid)?;
        text = text
            .checked_add(counts.text_bytes)
            .ok_or_else(DecodeError::invalid)?;
        max_depth = max_depth.max(counts.max_depth);
        formats = formats.checked_add(1).ok_or_else(DecodeError::invalid)?;
    }
    fields = fields
        .checked_add(
            write
                .uuids()
                .len()
                .checked_mul(3)
                .ok_or_else(DecodeError::invalid)?,
        )
        .ok_or_else(DecodeError::invalid)?;
    // Each archive also has one outer list field (field 2); the per-archive
    // count returned by `write_counts` begins at the archive payload itself.
    fields = fields
        .checked_add(formats)
        .ok_or_else(DecodeError::invalid)?;
    work = work
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let requirements = canonical_requirements(
        output_bytes,
        fields,
        work,
        max_depth,
        write.uuids().len(),
        formats,
        conditions,
        text,
        output_bytes,
        custom_format_list_materialization(write)?,
    )?;
    check_options(requirements, options)?;
    Ok(PreparedCustomFormatListWrite {
        write,
        output_bytes,
        requirements,
        verify_options: options,
    })
}

/// Emit a canonical custom-format registry.
pub fn canonical_custom_format_list(
    write: CustomFormatListWrite<'_, '_>,
    options: DecodeOptions,
) -> Result<RewriteOutput, DecodeError> {
    let prepared = prepare_custom_format_list_write(write, options)?;
    prepared.execute(RewriteExecutionLimits::exact(
        prepared.execution_requirements(),
    ))
}

/// Compatibility spelling for table-cell custom registries.
pub use canonical_custom_format_list as canonical_table_cell_custom_format_list;

// ---------------------------------------------------------------------------
// Borrowed iterators
// ---------------------------------------------------------------------------

/// Iterator over source-borrowed UUIDs.
#[derive(Debug, Clone, Copy)]
pub struct UuidIter<'source> {
    source: &'source [u8],
    cursor: usize,
}

impl<'source> Iterator for UuidIter<'source> {
    type Item = Result<Uuid, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            if self.cursor == self.source.len() {
                return None;
            }
            let field = match next_raw_field(self.source, &mut self.cursor) {
                Ok(Some(field)) => field,
                Ok(None) => return None,
                Err(error) => return Some(Err(error)),
            };
            if field.number == LIST_UUIDS_FIELD {
                let Some(payload) = field.payload else {
                    return Some(Err(DecodeError::invalid()));
                };
                return Some(decode_uuid(payload));
            }
        }
    }
}

/// Iterator over source-borrowed custom archives.
#[derive(Debug, Clone, Copy)]
pub struct CustomFormatIter<'source> {
    source: &'source [u8],
    cursor: usize,
}

impl<'source> Iterator for CustomFormatIter<'source> {
    type Item = Result<CustomFormatSnapshot<'source>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            if self.cursor == self.source.len() {
                return None;
            }
            let field = match next_raw_field(self.source, &mut self.cursor) {
                Ok(Some(field)) => field,
                Ok(None) => return None,
                Err(error) => return Some(Err(error)),
            };
            if field.number == LIST_CUSTOM_FORMATS_FIELD {
                let Some(payload) = field.payload else {
                    return Some(Err(DecodeError::invalid()));
                };
                let mut budget = Budget::for_iterator(payload);
                return Some(scan_custom_format(payload, 0, 0, &mut budget));
            }
        }
    }
}

/// Iterator over source-borrowed conditions.
#[derive(Debug, Clone, Copy)]
pub struct CustomConditionIter<'source> {
    source: &'source [u8],
    cursor: usize,
    expected_format_type: u32,
}

impl<'source> Iterator for CustomConditionIter<'source> {
    type Item = Result<CustomConditionSnapshot<'source>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            if self.cursor == self.source.len() {
                return None;
            }
            let field = match next_raw_field(self.source, &mut self.cursor) {
                Ok(Some(field)) => field,
                Ok(None) => return None,
                Err(error) => return Some(Err(error)),
            };
            if field.number == ARCHIVE_CONDITIONS_FIELD {
                let Some(payload) = field.payload else {
                    return Some(Err(DecodeError::invalid()));
                };
                let mut budget = Budget::for_iterator(payload);
                return Some(scan_condition(
                    payload,
                    0,
                    0,
                    self.expected_format_type,
                    &mut budget,
                ));
            }
        }
    }
}

/// Iterator over parallel UUID/archive entries.
#[derive(Debug, Clone, Copy)]
pub struct CustomFormatListEntryIter<'source> {
    uuids: UuidIter<'source>,
    formats: CustomFormatIter<'source>,
}

impl<'source> Iterator for CustomFormatListEntryIter<'source> {
    type Item = Result<CustomFormatListEntry<'source>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        match (self.uuids.next(), self.formats.next()) {
            (None, None) => None,
            (Some(Ok(uuid)), Some(Ok(custom_format))) => Some(Ok(CustomFormatListEntry {
                uuid,
                custom_format,
            })),
            (Some(Err(error)), _) | (_, Some(Err(error))) => Some(Err(error)),
            _ => Some(Err(DecodeError::invalid())),
        }
    }
}

// ---------------------------------------------------------------------------
// Strict bounded scanners and Buffa parity
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy)]
struct Budget {
    options: DecodeOptions,
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    formats: usize,
    conditions: usize,
    text_bytes: usize,
    allocations: usize,
    scratch_bytes: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES).map_err(|_| DecodeError::invalid())?;
        if options.max_message_bytes > hard {
            return Err(DecodeError::limited(DecodeLimit::InputBytes {
                observed: options.max_message_bytes,
                maximum: hard,
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
                maximum: MAX_RECURSION,
            }));
        }
        Ok(Self {
            options,
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            formats: 0,
            conditions: 0,
            text_bytes: 0,
            allocations: 0,
            scratch_bytes: 0,
        })
    }

    fn for_iterator(source: &[u8]) -> Self {
        Self {
            options: DecodeOptions::for_source(source),
            input_bytes: source.len(),
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            references: 0,
            formats: 0,
            conditions: 0,
            text_bytes: 0,
            allocations: 0,
            scratch_bytes: 0,
        }
    }

    /// Account for one generated Buffa `RepeatedView` backing collection.
    ///
    /// Buffa 0.9.1 stores repeated borrowed byte values in a `Vec<&[u8]>`.
    /// The generated lazy decoder grows that vector with `push`, so its final
    /// capacity is not necessarily the exact item count.  A power-of-two
    /// capacity with the current non-zero minimum (four elements) is a
    /// conservative bound for the pinned Buffa/Rust implementation.  The
    /// collection is one logical allocation unit even when growth performs
    /// more than one allocator call; this is the same logical allocation
    /// accounting used by the prepared rewrite APIs.
    fn repeated_view(&mut self, count: usize) -> Result<(), DecodeError> {
        if count == 0 {
            return Ok(());
        }
        let capacity = count
            .checked_next_power_of_two()
            .ok_or_else(DecodeError::invalid)?
            .max(4);
        let bytes = capacity
            .checked_mul(size_of::<&[u8]>())
            .ok_or_else(DecodeError::invalid)?;
        self.allocations = self
            .allocations
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
        self.scratch_bytes = self
            .scratch_bytes
            .checked_add(bytes)
            .ok_or_else(DecodeError::invalid)?;
        Ok(())
    }

    fn field(&mut self, bytes: usize, depth: u32) -> Result<(), DecodeError> {
        self.fields = self.fields.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Fields {
                observed: usize::MAX,
                maximum: self.options.max_fields,
            })
        })?;
        self.work(bytes)?;
        self.max_depth = self.max_depth.max(depth);
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
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

    fn work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.work_bytes = self.work_bytes.checked_add(bytes).ok_or_else(|| {
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

    fn reference(&mut self) -> Result<(), DecodeError> {
        self.references = self.references.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::References {
                observed: usize::MAX,
                maximum: self.options.max_references,
            })
        })?;
        if self.references > self.options.max_references {
            return Err(DecodeError::limited(DecodeLimit::References {
                observed: self.references,
                maximum: self.options.max_references,
            }));
        }
        Ok(())
    }

    fn format(&mut self) -> Result<(), DecodeError> {
        self.formats = self.formats.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Formats {
                observed: usize::MAX,
                maximum: self.options.max_items,
            })
        })?;
        if self.formats > self.options.max_items {
            return Err(DecodeError::limited(DecodeLimit::Formats {
                observed: self.formats,
                maximum: self.options.max_items,
            }));
        }
        Ok(())
    }

    fn condition(&mut self) -> Result<(), DecodeError> {
        self.conditions = self.conditions.checked_add(1).ok_or_else(|| {
            DecodeError::limited(DecodeLimit::Conditions {
                observed: usize::MAX,
                maximum: self.options.max_items,
            })
        })?;
        if self.conditions > self.options.max_items {
            return Err(DecodeError::limited(DecodeLimit::Conditions {
                observed: self.conditions,
                maximum: self.options.max_items,
            }));
        }
        Ok(())
    }

    fn text(&mut self, bytes: usize) -> Result<(), DecodeError> {
        self.text_bytes = self.text_bytes.checked_add(bytes).ok_or_else(|| {
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

    fn finish(self, output_bytes: usize) -> DecodeReport {
        DecodeReport {
            input_bytes: self.input_bytes,
            output_bytes,
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            references: self.references,
            formats: self.formats,
            conditions: self.conditions,
            text_bytes: self.text_bytes,
            allocations: self.allocations,
            retained_bytes: self.input_bytes.saturating_add(output_bytes),
            scratch_bytes: self.scratch_bytes,
        }
    }
}

fn scan_custom_format_list<'source>(
    source: &'source [u8],
    mut cursor: usize,
    depth: u32,
    budget: &mut Budget,
) -> Result<CustomFormatListSnapshot<'source>, DecodeError> {
    let mut uuid_count = 0usize;
    let mut format_count = 0usize;
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            return Err(DecodeError::invalid());
        }
        let value_start = cursor;
        match number {
            LIST_UUIDS_FIELD | LIST_CUSTOM_FORMATS_FIELD if wire == 2 => {
                let (payload, end) = read_length_payload(source, cursor)?;
                cursor = end;
                if number == LIST_UUIDS_FIELD {
                    uuid_count = uuid_count.checked_add(1).ok_or_else(DecodeError::invalid)?;
                    budget.reference()?;
                    let _ = decode_uuid_with_budget(payload, depth.saturating_add(1), budget)?;
                } else {
                    format_count = format_count
                        .checked_add(1)
                        .ok_or_else(DecodeError::invalid)?;
                    budget.format()?;
                    let _ = scan_custom_format(payload, 0, depth.saturating_add(1), budget)?;
                }
            },
            LIST_UUIDS_FIELD | LIST_CUSTOM_FORMATS_FIELD => return Err(DecodeError::invalid()),
            _ => {
                if number <= 2 {
                    return Err(DecodeError::invalid());
                }
                cursor = scan_unknown_value(source, cursor, wire, number, depth, budget)?;
            },
        }
        budget.field(
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?,
            depth,
        )?;
        if value_start > cursor || cursor > source.len() {
            return Err(DecodeError::invalid());
        }
    }
    Ok(CustomFormatListSnapshot {
        source,
        uuid_count,
        format_count,
    })
}

fn scan_custom_format<'source>(
    source: &'source [u8],
    cursor: usize,
    depth: u32,
    budget: &mut Budget,
) -> Result<CustomFormatSnapshot<'source>, DecodeError> {
    let mut cursor = cursor;
    let mut name = None;
    let mut format_type_pre_bnc = None;
    let mut default_format = None;
    let mut conditions_start = source.len();
    let mut conditions = 0usize;
    let mut format_type = None;
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            return Err(DecodeError::invalid());
        }
        match number {
            ARCHIVE_NAME_FIELD if wire == 2 && name.is_none() => {
                let (payload, end) = read_length_payload(source, cursor)?;
                cursor = end;
                let value = str::from_utf8(payload).map_err(|_| DecodeError::invalid())?;
                validate_name(value)?;
                budget.text(payload.len())?;
                name = Some(value);
            },
            ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD if wire == 0 && format_type_pre_bnc.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                format_type_pre_bnc = Some(as_u32(value)?);
            },
            ARCHIVE_DEFAULT_FORMAT_FIELD if wire == 2 && default_format.is_none() => {
                let (payload, end) = read_length_payload(source, cursor)?;
                cursor = end;
                default_format =
                    Some(scan_format_struct(payload, 0, depth.saturating_add(1), None, budget)?.0);
            },
            ARCHIVE_CONDITIONS_FIELD if wire == 2 => {
                let (payload, end) = read_length_payload(source, cursor)?;
                cursor = end;
                if conditions == 0 {
                    conditions_start = start;
                }
                conditions = conditions.checked_add(1).ok_or_else(DecodeError::invalid)?;
                budget.condition()?;
                let _ = scan_condition(
                    payload,
                    0,
                    depth.saturating_add(1),
                    format_type_pre_bnc.unwrap_or(0),
                    budget,
                )?;
            },
            ARCHIVE_FORMAT_TYPE_FIELD if wire == 0 && format_type.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                format_type = Some(as_u32(value)?);
            },
            ARCHIVE_NAME_FIELD
            | ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD
            | ARCHIVE_DEFAULT_FORMAT_FIELD
            | ARCHIVE_CONDITIONS_FIELD
            | ARCHIVE_FORMAT_TYPE_FIELD => return Err(DecodeError::invalid()),
            _ => {
                if number <= ARCHIVE_FORMAT_TYPE_FIELD {
                    return Err(DecodeError::invalid());
                }
                cursor = scan_unknown_value(source, cursor, wire, number, depth, budget)?;
            },
        }
        budget.field(
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?,
            depth,
        )?;
    }
    let name = name.ok_or_else(DecodeError::invalid)?;
    let format_type_pre_bnc = format_type_pre_bnc.ok_or_else(DecodeError::invalid)?;
    let default_format = default_format.ok_or_else(DecodeError::invalid)?;
    let format_type = format_type.ok_or_else(DecodeError::invalid)?;
    if !is_custom_format_type(format_type_pre_bnc)
        || format_type != format_type_pre_bnc
        || default_format.format_type() != format_type
        || (format_type != NATIVE_CUSTOM_NUMBER_FORMAT_TYPE && conditions != 0)
        || conditions > MAX_CUSTOM_CONDITIONS
    {
        return Err(DecodeError::invalid());
    }
    if conditions != 0 {
        // Conditions may legally precede the archive's optional `format_type`
        // field.  Re-run only the bounded condition envelopes after the final
        // type is known so a mismatched nested family cannot reach Buffa.
        validate_condition_formats(source, format_type, conditions)?;
    }
    Ok(CustomFormatSnapshot {
        source,
        name,
        format_type_pre_bnc,
        default_format,
        conditions_start: if conditions == 0 {
            source.len()
        } else {
            conditions_start
        },
        conditions,
        format_type,
    })
}

fn scan_condition<'source>(
    source: &'source [u8],
    cursor: usize,
    depth: u32,
    expected_format_type: u32,
    budget: &mut Budget,
) -> Result<CustomConditionSnapshot<'source>, DecodeError> {
    let mut cursor = cursor;
    let mut condition_type = None;
    let mut condition_value = None;
    let mut condition_format = None;
    let mut condition_value_dbl = None;
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            return Err(DecodeError::invalid());
        }
        match number {
            CONDITION_TYPE_FIELD if wire == 0 && condition_type.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                let value = as_u32(value)?;
                if value > 4 {
                    return Err(DecodeError::invalid());
                }
                condition_type = Some(value);
            },
            CONDITION_VALUE_FIELD if wire == 5 && condition_value.is_none() => {
                let bytes = source
                    .get(cursor..cursor.checked_add(4).ok_or_else(DecodeError::invalid)?)
                    .ok_or_else(DecodeError::invalid)?;
                cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?;
                let value =
                    f32::from_le_bytes(bytes.try_into().map_err(|_| DecodeError::invalid())?);
                if !value.is_finite() {
                    return Err(DecodeError::invalid());
                }
                condition_value = Some(value);
            },
            CONDITION_FORMAT_FIELD if wire == 2 && condition_format.is_none() => {
                let (payload, end) = read_length_payload(source, cursor)?;
                cursor = end;
                condition_format = Some(
                    scan_format_struct(
                        payload,
                        0,
                        depth.saturating_add(1),
                        is_custom_format_type(expected_format_type).then_some(expected_format_type),
                        budget,
                    )?
                    .0,
                );
            },
            CONDITION_VALUE_DBL_FIELD if wire == 1 && condition_value_dbl.is_none() => {
                let bytes = source
                    .get(cursor..cursor.checked_add(8).ok_or_else(DecodeError::invalid)?)
                    .ok_or_else(DecodeError::invalid)?;
                cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?;
                let value =
                    f64::from_le_bytes(bytes.try_into().map_err(|_| DecodeError::invalid())?);
                if !value.is_finite() {
                    return Err(DecodeError::invalid());
                }
                condition_value_dbl = Some(value);
            },
            CONDITION_TYPE_FIELD
            | CONDITION_VALUE_FIELD
            | CONDITION_FORMAT_FIELD
            | CONDITION_VALUE_DBL_FIELD => return Err(DecodeError::invalid()),
            _ => {
                if number <= CONDITION_VALUE_DBL_FIELD {
                    return Err(DecodeError::invalid());
                }
                cursor = scan_unknown_value(source, cursor, wire, number, depth, budget)?;
            },
        }
        budget.field(
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?,
            depth,
        )?;
    }
    let condition_type = condition_type.ok_or_else(DecodeError::invalid)?;
    let condition_format = condition_format.ok_or_else(DecodeError::invalid)?;
    if condition_value.is_some() == condition_value_dbl.is_some() {
        return Err(DecodeError::invalid());
    }
    Ok(CustomConditionSnapshot {
        source,
        condition_type,
        condition_value,
        condition_format,
        condition_value_dbl,
    })
}

fn scan_format_struct<'source>(
    source: &'source [u8],
    cursor: usize,
    depth: u32,
    expected_format_type: Option<u32>,
    budget: &mut Budget,
) -> Result<(FormatStructSnapshot<'source>, usize), DecodeError> {
    let mut cursor = cursor;
    let mut format_type = None;
    let mut show_thousands_separator = None;
    let mut use_accounting_style = None;
    let mut fraction_accuracy = None;
    let mut custom_format_string = None;
    let mut scale_factor = None;
    let mut requires_fraction_replacement = None;
    let mut decimal_width = None;
    let mut min_integer_width = None;
    let mut num_nonspace_integer_digits = None;
    let mut num_nonspace_decimal_digits = None;
    let mut index_from_right_last_integer = None;
    let mut num_hash_decimal_digits = None;
    let mut total_num_decimal_digits = None;
    let mut is_complex = None;
    let mut contains_integer_token = None;
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            return Err(DecodeError::invalid());
        }
        match number {
            FORMAT_TYPE_FIELD if wire == 0 && format_type.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                let value = as_u32(value)?;
                if !is_custom_format_type(value)
                    || expected_format_type.is_some_and(|expected| expected != value)
                {
                    return Err(DecodeError::invalid());
                }
                format_type = Some(value);
            },
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD
                if wire == 0 && show_thousands_separator.is_none() =>
            {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                show_thousands_separator = Some(as_bool(value)?);
            },
            FORMAT_USE_ACCOUNTING_STYLE_FIELD if wire == 0 && use_accounting_style.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                let value = as_bool(value)?;
                if value {
                    return Err(DecodeError::invalid());
                }
                use_accounting_style = Some(value);
            },
            FORMAT_FRACTION_ACCURACY_FIELD if wire == 0 && fraction_accuracy.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                let value = as_u32(value)?;
                if value != NATIVE_CUSTOM_FRACTION_SENTINEL {
                    return Err(DecodeError::invalid());
                }
                fraction_accuracy = Some(value);
            },
            FORMAT_CUSTOM_FORMAT_STRING_FIELD if wire == 2 && custom_format_string.is_none() => {
                let (payload, end) = read_length_payload(source, cursor)?;
                cursor = end;
                let value = str::from_utf8(payload).map_err(|_| DecodeError::invalid())?;
                validate_pattern(value)?;
                budget.text(payload.len())?;
                custom_format_string = Some(value);
            },
            FORMAT_SCALE_FACTOR_FIELD if wire == 1 && scale_factor.is_none() => {
                let bytes = source
                    .get(cursor..cursor.checked_add(8).ok_or_else(DecodeError::invalid)?)
                    .ok_or_else(DecodeError::invalid)?;
                cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?;
                let value =
                    f64::from_le_bytes(bytes.try_into().map_err(|_| DecodeError::invalid())?);
                if !value.is_finite() || value != 1.0 {
                    return Err(DecodeError::invalid());
                }
                scale_factor = Some(value);
            },
            FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD
                if wire == 0 && requires_fraction_replacement.is_none() =>
            {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                let value = as_bool(value)?;
                if value {
                    return Err(DecodeError::invalid());
                }
                requires_fraction_replacement = Some(value);
            },
            FORMAT_DECIMAL_WIDTH_FIELD if wire == 0 && decimal_width.is_none() => {
                decimal_width = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_MIN_INTEGER_WIDTH_FIELD if wire == 0 && min_integer_width.is_none() => {
                min_integer_width = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_NONSPACE_INTEGER_DIGITS_FIELD
                if wire == 0 && num_nonspace_integer_digits.is_none() =>
            {
                num_nonspace_integer_digits = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD
                if wire == 0 && num_nonspace_decimal_digits.is_none() =>
            {
                num_nonspace_decimal_digits = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_INDEX_FROM_RIGHT_FIELD
                if wire == 0 && index_from_right_last_integer.is_none() =>
            {
                index_from_right_last_integer = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_HASH_DECIMAL_DIGITS_FIELD if wire == 0 && num_hash_decimal_digits.is_none() => {
                num_hash_decimal_digits = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_TOTAL_DECIMAL_DIGITS_FIELD
                if wire == 0 && total_num_decimal_digits.is_none() =>
            {
                total_num_decimal_digits = Some(read_u32_advance(source, &mut cursor)?);
            },
            FORMAT_IS_COMPLEX_FIELD if wire == 0 && is_complex.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                let value = as_bool(value)?;
                if value {
                    return Err(DecodeError::invalid());
                }
                is_complex = Some(value);
            },
            FORMAT_CONTAINS_INTEGER_TOKEN_FIELD
                if wire == 0 && contains_integer_token.is_none() =>
            {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                contains_integer_token = Some(as_bool(value)?);
            },
            FORMAT_TYPE_FIELD
            | FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD
            | FORMAT_USE_ACCOUNTING_STYLE_FIELD
            | FORMAT_FRACTION_ACCURACY_FIELD
            | FORMAT_CUSTOM_FORMAT_STRING_FIELD
            | FORMAT_SCALE_FACTOR_FIELD
            | FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD
            | FORMAT_DECIMAL_WIDTH_FIELD
            | FORMAT_MIN_INTEGER_WIDTH_FIELD
            | FORMAT_NONSPACE_INTEGER_DIGITS_FIELD
            | FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD
            | FORMAT_INDEX_FROM_RIGHT_FIELD
            | FORMAT_HASH_DECIMAL_DIGITS_FIELD
            | FORMAT_TOTAL_DECIMAL_DIGITS_FIELD
            | FORMAT_IS_COMPLEX_FIELD
            | FORMAT_CONTAINS_INTEGER_TOKEN_FIELD => return Err(DecodeError::invalid()),
            _ => {
                if number <= FORMAT_MAX_KNOWN_FIELD {
                    return Err(DecodeError::invalid());
                }
                cursor = scan_unknown_value(source, cursor, wire, number, depth, budget)?;
            },
        }
        budget.field(
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?,
            depth,
        )?;
    }
    let snapshot = FormatStructSnapshot {
        source,
        format_type: format_type.ok_or_else(DecodeError::invalid)?,
        show_thousands_separator: show_thousands_separator.ok_or_else(DecodeError::invalid)?,
        use_accounting_style: use_accounting_style.ok_or_else(DecodeError::invalid)?,
        fraction_accuracy: fraction_accuracy.ok_or_else(DecodeError::invalid)?,
        custom_format_string: custom_format_string.ok_or_else(DecodeError::invalid)?,
        scale_factor: scale_factor.ok_or_else(DecodeError::invalid)?,
        requires_fraction_replacement: requires_fraction_replacement
            .ok_or_else(DecodeError::invalid)?,
        decimal_width: decimal_width.ok_or_else(DecodeError::invalid)?,
        min_integer_width: min_integer_width.ok_or_else(DecodeError::invalid)?,
        num_nonspace_integer_digits: num_nonspace_integer_digits
            .ok_or_else(DecodeError::invalid)?,
        num_nonspace_decimal_digits: num_nonspace_decimal_digits
            .ok_or_else(DecodeError::invalid)?,
        index_from_right_last_integer: index_from_right_last_integer
            .ok_or_else(DecodeError::invalid)?,
        num_hash_decimal_digits: num_hash_decimal_digits.ok_or_else(DecodeError::invalid)?,
        total_num_decimal_digits: total_num_decimal_digits.ok_or_else(DecodeError::invalid)?,
        is_complex: is_complex.ok_or_else(DecodeError::invalid)?,
        contains_integer_token: contains_integer_token.ok_or_else(DecodeError::invalid)?,
    };
    if snapshot.format_type != NATIVE_CUSTOM_NUMBER_FORMAT_TYPE && snapshot.show_thousands_separator
    {
        return Err(DecodeError::invalid());
    }
    validate_pattern_for_type(
        snapshot.custom_format_string,
        snapshot.format_type,
        snapshot.contains_integer_token,
    )?;
    Ok((snapshot, cursor))
}

fn scan_unknown_value(
    source: &[u8],
    cursor: usize,
    wire: u8,
    number: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut cursor = cursor;
    match wire {
        0 => {
            cursor = cursor
                .checked_add(read_varint_relaxed(source, cursor)?.1)
                .ok_or_else(DecodeError::invalid)?;
        },
        1 => cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?,
        2 => cursor = read_length_payload(source, cursor)?.1,
        3 => {
            if depth >= budget.options.recursion_limit {
                return Err(DecodeError::limited(DecodeLimit::Nesting {
                    observed: depth.saturating_add(1),
                    maximum: budget.options.recursion_limit,
                }));
            }
            let nested =
                scan_unknown_group(source, cursor, number, depth.saturating_add(1), budget)?;
            cursor = nested;
        },
        5 => cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?,
        _ => return Err(DecodeError::invalid()),
    }
    if cursor > source.len() {
        return Err(DecodeError::invalid());
    }
    Ok(cursor)
}

fn scan_unknown_group(
    source: &[u8],
    mut cursor: usize,
    group: u32,
    depth: u32,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            if number != group {
                return Err(DecodeError::invalid());
            }
            budget.field(key_len, depth)?;
            return Ok(cursor);
        }
        let value_start = cursor;
        cursor = scan_unknown_value(source, cursor, wire, number, depth, budget)?;
        let bytes = if wire == 3 {
            value_start
                .checked_sub(start)
                .ok_or_else(DecodeError::invalid)?
        } else {
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?
        };
        budget.field(bytes, depth)?;
    }
    Err(DecodeError::invalid())
}

fn buffa_format_parity(
    source: &[u8],
    snapshot: FormatStructSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let options = budget.options;
    let view: projection::FormatStructArchiveLazyView<'_> = buffa_options(options, 0)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.format_type != Some(snapshot.format_type)
        || view.show_thousands_separator != Some(snapshot.show_thousands_separator)
        || view.use_accounting_style != Some(snapshot.use_accounting_style)
        || view.fraction_accuracy != Some(snapshot.fraction_accuracy)
        || view.custom_format_string != Some(snapshot.custom_format_string)
        || view.scale_factor != Some(snapshot.scale_factor)
        || view.requires_fraction_replacement != Some(snapshot.requires_fraction_replacement)
        || view.decimal_width != Some(snapshot.decimal_width)
        || view.min_integer_width != Some(snapshot.min_integer_width)
        || view.num_nonspace_integer_digits != Some(snapshot.num_nonspace_integer_digits)
        || view.num_nonspace_decimal_digits != Some(snapshot.num_nonspace_decimal_digits)
        || view.index_from_right_last_integer != Some(snapshot.index_from_right_last_integer)
        || view.num_hash_decimal_digits != Some(snapshot.num_hash_decimal_digits)
        || view.total_num_decimal_digits != Some(snapshot.total_num_decimal_digits)
        || view.is_complex != Some(snapshot.is_complex)
        || view.contains_integer_token != Some(snapshot.contains_integer_token)
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())
}

fn buffa_uuid_parity(source: &[u8], uuid: Uuid, budget: &mut Budget) -> Result<(), DecodeError> {
    let view: projection::UuidArchiveLazyView<'_> = buffa_options(budget.options, 0)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.lower != uuid.lower || view.upper != uuid.upper {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())
}

fn buffa_condition_parity(
    source: &[u8],
    snapshot: CustomConditionSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let memory = size_of::<&[u8]>();
    let view: projection::ConditionLazyView<'_> = buffa_options(budget.options, memory)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.condition_type != snapshot.condition_type
        || view.condition_value != snapshot.condition_value
        || view.condition_value_dbl != snapshot.condition_value_dbl
        || view.condition_format != snapshot.condition_format.raw()
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())
}

fn buffa_custom_format_parity(
    source: &[u8],
    snapshot: CustomFormatSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.repeated_view(snapshot.condition_count())?;
    let memory = snapshot
        .condition_count()
        .checked_mul(size_of::<&[u8]>())
        .ok_or_else(DecodeError::invalid)?;
    let view: projection::CustomFormatArchiveLazyView<'_> = buffa_options(budget.options, memory)
        .decode_lazy_view(source)
        .map_err(|_| DecodeError::invalid())?;
    if view.name != snapshot.name
        || view.format_type_pre_bnc != snapshot.format_type_pre_bnc
        || view.default_format != snapshot.default_format.raw()
        || view.format_type != Some(snapshot.format_type)
        || view.conditions.len() != snapshot.condition_count()
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())?;
    buffa_format_parity(
        snapshot.default_format.raw(),
        snapshot.default_format,
        budget,
    )?;
    for condition in snapshot.conditions() {
        let condition = condition?;
        buffa_condition_parity(condition.raw(), condition, budget)?;
        buffa_format_parity(
            condition.condition_format().raw(),
            condition.condition_format(),
            budget,
        )?;
    }
    Ok(())
}

fn buffa_custom_format_list_parity(
    source: &[u8],
    snapshot: CustomFormatListSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.repeated_view(snapshot.uuid_count())?;
    budget.repeated_view(snapshot.format_count())?;
    let pointers = snapshot
        .uuid_count()
        .checked_add(snapshot.format_count())
        .and_then(|count| count.checked_mul(size_of::<&[u8]>()))
        .ok_or_else(DecodeError::invalid)?;
    let view: projection::CustomFormatListArchiveLazyView<'_> =
        buffa_options(budget.options, pointers)
            .decode_lazy_view(source)
            .map_err(|_| DecodeError::invalid())?;
    if view.uuids.len() != snapshot.uuid_count()
        || view.custom_formats.len() != snapshot.format_count()
    {
        return Err(DecodeError::invalid());
    }
    budget.work(source.len())?;
    for field in list_uuid_payloads(source) {
        let payload = field?;
        let uuid = decode_uuid(payload)?;
        buffa_uuid_parity(payload, uuid, budget)?;
    }
    for field in list_format_payloads(source) {
        let payload = field?;
        let mut nested = Budget::for_iterator(payload);
        let custom = scan_custom_format(payload, 0, 0, &mut nested)?;
        buffa_custom_format_parity(payload, custom, budget)?;
    }
    Ok(())
}

fn buffa_options(options: DecodeOptions, element_memory: usize) -> BuffaDecodeOptions {
    BuffaDecodeOptions::new()
        .with_max_message_size(options.max_message_bytes)
        .with_unknown_field_limit(options.max_fields)
        .with_element_memory_limit(element_memory)
        .with_recursion_limit(options.recursion_limit)
}

// ---------------------------------------------------------------------------
// Wire helpers, validation, and rewrite emission
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy)]
struct RawField<'source> {
    number: u32,
    wire: u8,
    start: usize,
    end: usize,
    payload: Option<&'source [u8]>,
}

fn next_raw_field<'source>(
    source: &'source [u8],
    cursor: &mut usize,
) -> Result<Option<RawField<'source>>, DecodeError> {
    if *cursor == source.len() {
        return Ok(None);
    }
    let start = *cursor;
    let (number, wire, key_len) = read_key(source, start)?;
    *cursor = cursor
        .checked_add(key_len)
        .ok_or_else(DecodeError::invalid)?;
    if wire == 4 {
        return Err(DecodeError::invalid());
    }
    let payload = match wire {
        0 => {
            *cursor = cursor
                .checked_add(read_varint_relaxed(source, *cursor)?.1)
                .ok_or_else(DecodeError::invalid)?;
            None
        },
        1 => {
            *cursor = cursor.checked_add(8).ok_or_else(DecodeError::invalid)?;
            None
        },
        2 => {
            let (payload, end) = read_length_payload(source, *cursor)?;
            *cursor = end;
            Some(payload)
        },
        3 => {
            *cursor = skip_group(source, *cursor, number)?;
            None
        },
        5 => {
            *cursor = cursor.checked_add(4).ok_or_else(DecodeError::invalid)?;
            None
        },
        _ => return Err(DecodeError::invalid()),
    };
    if *cursor > source.len() {
        return Err(DecodeError::invalid());
    }
    Ok(Some(RawField {
        number,
        wire,
        start,
        end: *cursor,
        payload,
    }))
}

fn skip_group(source: &[u8], mut cursor: usize, group: u32) -> Result<usize, DecodeError> {
    while cursor < source.len() {
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            if number != group {
                return Err(DecodeError::invalid());
            }
            return Ok(cursor);
        }
        cursor = match wire {
            0 => cursor
                .checked_add(read_varint_relaxed(source, cursor)?.1)
                .ok_or_else(DecodeError::invalid)?,
            1 => cursor.checked_add(8).ok_or_else(DecodeError::invalid)?,
            2 => read_length_payload(source, cursor)?.1,
            3 => skip_group(source, cursor, number)?,
            5 => cursor.checked_add(4).ok_or_else(DecodeError::invalid)?,
            _ => return Err(DecodeError::invalid()),
        };
        if cursor > source.len() {
            return Err(DecodeError::invalid());
        }
    }
    Err(DecodeError::invalid())
}

fn read_key(source: &[u8], cursor: usize) -> Result<(u32, u8, usize), DecodeError> {
    let (value, length) = read_varint(source, cursor)?;
    let number = u32::try_from(value >> 3).map_err(|_| DecodeError::invalid())?;
    let wire = u8::try_from(value & 7).map_err(|_| DecodeError::invalid())?;
    // End-group (wire type 4) is parsed by the bounded group scanner; all
    // message-level callers reject it when it is not paired with a group.
    if number == 0 || number > MAX_FIELD_NUMBER || matches!(wire, 5..=7) {
        return Err(DecodeError::invalid());
    }
    Ok((number, wire, length))
}

fn read_varint(source: &[u8], cursor: usize) -> Result<(u64, usize), DecodeError> {
    let (value, length) = read_varint_relaxed(source, cursor)?;
    if varint_len(value) != length {
        return Err(DecodeError::invalid());
    }
    Ok((value, length))
}

fn read_varint_relaxed(source: &[u8], cursor: usize) -> Result<(u64, usize), DecodeError> {
    let mut value = 0u64;
    for offset in 0..10usize {
        let byte = *source
            .get(
                cursor
                    .checked_add(offset)
                    .ok_or_else(DecodeError::invalid)?,
            )
            .ok_or_else(DecodeError::invalid)?;
        if offset == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (offset * 7);
        if byte & 0x80 == 0 {
            return Ok((value, offset + 1));
        }
    }
    Err(DecodeError::invalid())
}

fn read_length_payload(source: &[u8], cursor: usize) -> Result<(&[u8], usize), DecodeError> {
    let (length, length_bytes) = read_varint(source, cursor)?;
    let length = usize::try_from(length).map_err(|_| DecodeError::invalid())?;
    let payload_start = cursor
        .checked_add(length_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let payload_end = payload_start
        .checked_add(length)
        .ok_or_else(DecodeError::invalid)?;
    let payload = source
        .get(payload_start..payload_end)
        .ok_or_else(DecodeError::invalid)?;
    Ok((payload, payload_end))
}

fn read_u32_advance(source: &[u8], cursor: &mut usize) -> Result<u32, DecodeError> {
    let (value, length) = read_varint(source, *cursor)?;
    *cursor = cursor
        .checked_add(length)
        .ok_or_else(DecodeError::invalid)?;
    as_u32(value)
}

fn as_u32(value: u64) -> Result<u32, DecodeError> {
    u32::try_from(value).map_err(|_| DecodeError::invalid())
}

fn as_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid()),
    }
}

fn decode_uuid(source: &[u8]) -> Result<Uuid, DecodeError> {
    let mut cursor = 0usize;
    let mut lower = None;
    let mut upper = None;
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            return Err(DecodeError::invalid());
        }
        match number {
            UUID_LOWER_FIELD if wire == 0 && lower.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                lower = Some(value);
            },
            UUID_UPPER_FIELD if wire == 0 && upper.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                upper = Some(value);
            },
            UUID_LOWER_FIELD | UUID_UPPER_FIELD => return Err(DecodeError::invalid()),
            _ => {
                cursor = skip_group_or_value(source, cursor, wire, number)?;
            },
        }
        if cursor <= start || cursor > source.len() {
            return Err(DecodeError::invalid());
        }
    }
    let lower = lower.ok_or_else(DecodeError::invalid)?;
    let upper = upper.ok_or_else(DecodeError::invalid)?;
    if lower == 0 || upper == 0 {
        return Err(DecodeError::invalid());
    }
    Ok(Uuid { lower, upper })
}

/// Decode a UUID and account for each nested UUID field in the caller's
/// bounded field/work budget.  The UUID itself is still returned by value;
/// its payload remains borrowed by the surrounding list snapshot.
fn decode_uuid_with_budget(
    source: &[u8],
    depth: u32,
    budget: &mut Budget,
) -> Result<Uuid, DecodeError> {
    let mut cursor = 0usize;
    let mut lower = None;
    let mut upper = None;
    while cursor < source.len() {
        let start = cursor;
        let (number, wire, key_len) = read_key(source, cursor)?;
        cursor = cursor
            .checked_add(key_len)
            .ok_or_else(DecodeError::invalid)?;
        if wire == 4 {
            return Err(DecodeError::invalid());
        }
        match number {
            UUID_LOWER_FIELD if wire == 0 && lower.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                lower = Some(value);
            },
            UUID_UPPER_FIELD if wire == 0 && upper.is_none() => {
                let (value, value_len) = read_varint(source, cursor)?;
                cursor = cursor
                    .checked_add(value_len)
                    .ok_or_else(DecodeError::invalid)?;
                upper = Some(value);
            },
            UUID_LOWER_FIELD | UUID_UPPER_FIELD => return Err(DecodeError::invalid()),
            _ => {
                cursor = scan_unknown_value(source, cursor, wire, number, depth, budget)?;
            },
        }
        budget.field(
            cursor.checked_sub(start).ok_or_else(DecodeError::invalid)?,
            depth,
        )?;
    }
    let lower = lower.ok_or_else(DecodeError::invalid)?;
    let upper = upper.ok_or_else(DecodeError::invalid)?;
    if lower == 0 || upper == 0 {
        return Err(DecodeError::invalid());
    }
    Ok(Uuid { lower, upper })
}

/// Validate all condition payloads after the archive's final format family is
/// known.  The native schema allows the optional archive `format_type` to be
/// ordered after repeated conditions, so the first pass cannot always supply
/// the expected nested type.  This second pass is source-bounded and does not
/// allocate or expose any generated message.
fn validate_condition_formats(
    source: &[u8],
    expected_format_type: u32,
    expected_conditions: usize,
) -> Result<(), DecodeError> {
    let mut budget = Budget::for_iterator(source);
    let mut cursor = 0usize;
    let mut conditions = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        if field.number == ARCHIVE_CONDITIONS_FIELD {
            if field.wire != 2 {
                return Err(DecodeError::invalid());
            }
            let payload = field.payload.ok_or_else(DecodeError::invalid)?;
            scan_condition(payload, 0, 0, expected_format_type, &mut budget)?;
            conditions = conditions.checked_add(1).ok_or_else(DecodeError::invalid)?;
        }
    }
    if conditions != expected_conditions {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

/// Reject zero UUIDs and duplicate registry keys before the lazy Buffa view is
/// forced.  The source scanner has already charged the input and field work;
/// this bounded comparison charge accounts for the additional pair checks.
fn validate_unique_uuids(
    source: &[u8],
    count: usize,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    if count < 2 {
        return Ok(());
    }
    let comparisons = count
        .checked_mul(count.saturating_sub(1))
        .and_then(|value| value.checked_div(2))
        .ok_or_else(DecodeError::invalid)?;
    budget.work(
        comparisons
            .checked_mul(size_of::<Uuid>())
            .ok_or_else(DecodeError::invalid)?,
    )?;

    let mut outer_index = 0usize;
    for payload in list_uuid_payloads(source) {
        let uuid = decode_uuid(payload?)?;
        let mut inner = list_uuid_payloads(source);
        for _ in 0..=outer_index {
            let _ = inner.next().ok_or_else(DecodeError::invalid)??;
        }
        for candidate in inner {
            if decode_uuid(candidate?)? == uuid {
                return Err(DecodeError::invalid());
            }
        }
        outer_index = outer_index
            .checked_add(1)
            .ok_or_else(DecodeError::invalid)?;
    }
    if outer_index != count {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn skip_group_or_value(
    source: &[u8],
    cursor: usize,
    wire: u8,
    number: u32,
) -> Result<usize, DecodeError> {
    match wire {
        0 => cursor
            .checked_add(read_varint_relaxed(source, cursor)?.1)
            .ok_or_else(DecodeError::invalid),
        1 => cursor.checked_add(8).ok_or_else(DecodeError::invalid),
        2 => read_length_payload(source, cursor).map(|(_, end)| end),
        3 => skip_group(source, cursor, number),
        5 => cursor.checked_add(4).ok_or_else(DecodeError::invalid),
        _ => Err(DecodeError::invalid()),
    }
}

fn validate_name(value: &str) -> Result<(), DecodeError> {
    if value.is_empty()
        || value.len() > MAX_CUSTOM_NAME_BYTES
        || value.trim() != value
        || value.chars().any(char::is_control)
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn validate_pattern(value: &str) -> Result<(), DecodeError> {
    if value.is_empty()
        || value.len() > MAX_CUSTOM_PATTERN_BYTES
        || value.chars().any(char::is_control)
    {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn validate_pattern_for_type(
    pattern: &str,
    format_type: u32,
    contains_integer_token: bool,
) -> Result<(), DecodeError> {
    match format_type {
        NATIVE_CUSTOM_NUMBER_FORMAT_TYPE => {
            if !contains_integer_token
                || !pattern
                    .chars()
                    .any(|character| matches!(character, '#' | '0'))
            {
                return Err(DecodeError::invalid());
            }
        },
        NATIVE_CUSTOM_TEXT_FORMAT_TYPE => {
            if contains_integer_token {
                return Err(DecodeError::invalid());
            }
        },
        NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE => {
            if contains_integer_token
                || !pattern.chars().any(|character| {
                    matches!(
                        character,
                        'G' | 'y'
                            | 'Y'
                            | 'M'
                            | 'L'
                            | 'w'
                            | 'W'
                            | 'D'
                            | 'd'
                            | 'F'
                            | 'E'
                            | 'e'
                            | 'a'
                            | 'h'
                            | 'H'
                            | 'K'
                            | 'k'
                            | 'm'
                            | 's'
                            | 'S'
                    )
                })
            {
                return Err(DecodeError::invalid());
            }
        },
        _ => return Err(DecodeError::invalid()),
    }
    Ok(())
}

const fn is_custom_format_type(value: u32) -> bool {
    matches!(
        value,
        NATIVE_CUSTOM_NUMBER_FORMAT_TYPE
            | NATIVE_CUSTOM_TEXT_FORMAT_TYPE
            | NATIVE_CUSTOM_DATE_TIME_FORMAT_TYPE
    )
}

fn validate_format_struct_write(write: FormatStructWrite<'_>) -> Result<(), DecodeError> {
    if !is_custom_format_type(write.format_type)
        || write.use_accounting_style
        || write.fraction_accuracy != NATIVE_CUSTOM_FRACTION_SENTINEL
        || write.scale_factor != 1.0
        || !write.scale_factor.is_finite()
        || write.requires_fraction_replacement
        || write.is_complex
        || (write.format_type != NATIVE_CUSTOM_NUMBER_FORMAT_TYPE && write.show_thousands_separator)
    {
        return Err(DecodeError::invalid());
    }
    validate_pattern(write.custom_format_string)?;
    validate_pattern_for_type(
        write.custom_format_string,
        write.format_type,
        write.contains_integer_token,
    )
}

fn validate_custom_condition_write(
    write: CustomConditionWrite<'_>,
    expected_format_type: u32,
) -> Result<(), DecodeError> {
    if write.condition_type > 4
        || write.condition_value.is_some() == write.condition_value_dbl.is_some()
        || write
            .condition_value
            .is_some_and(|value| !value.is_finite())
        || write
            .condition_value_dbl
            .is_some_and(|value| !value.is_finite())
        || write.condition_format.format_type() != expected_format_type
    {
        return Err(DecodeError::invalid());
    }
    validate_format_struct_write(write.condition_format)
}

fn validate_custom_format_write(write: CustomFormatWrite<'_, '_>) -> Result<(), DecodeError> {
    validate_name(write.name)?;
    if !is_custom_format_type(write.format_type_pre_bnc)
        || write.format_type != write.format_type_pre_bnc
        || write.default_format.format_type() != write.format_type
        || (write.format_type != NATIVE_CUSTOM_NUMBER_FORMAT_TYPE && !write.conditions.is_empty())
        || write.conditions.len() > MAX_CUSTOM_CONDITIONS
    {
        return Err(DecodeError::invalid());
    }
    validate_format_struct_write(write.default_format)?;
    for condition in write.conditions {
        validate_custom_condition_write(*condition, write.format_type)?;
    }
    Ok(())
}

fn validate_custom_format_list_write(
    write: CustomFormatListWrite<'_, '_>,
) -> Result<(), DecodeError> {
    if write.uuids.len() != write.custom_formats.len() {
        return Err(DecodeError::invalid());
    }
    for (index, uuid) in write.uuids.iter().copied().enumerate() {
        if uuid.lower == 0 || uuid.upper == 0 || write.uuids[index + 1..].contains(&uuid) {
            return Err(DecodeError::invalid());
        }
    }
    for custom_format in write.custom_formats {
        validate_custom_format_write(*custom_format)?;
    }
    Ok(())
}

fn list_uuid_payloads<'source>(source: &'source [u8]) -> UuidPayloadIter<'source> {
    UuidPayloadIter { source, cursor: 0 }
}

fn list_format_payloads<'source>(source: &'source [u8]) -> FormatPayloadIter<'source> {
    FormatPayloadIter { source, cursor: 0 }
}

#[derive(Debug, Clone, Copy)]
struct UuidPayloadIter<'source> {
    source: &'source [u8],
    cursor: usize,
}

impl<'source> Iterator for UuidPayloadIter<'source> {
    type Item = Result<&'source [u8], DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let field = match next_raw_field(self.source, &mut self.cursor) {
                Ok(Some(field)) => field,
                Ok(None) => return None,
                Err(error) => return Some(Err(error)),
            };
            if field.number == LIST_UUIDS_FIELD {
                if field.wire != 2 {
                    return Some(Err(DecodeError::invalid()));
                }
                return Some(field.payload.ok_or_else(DecodeError::invalid));
            }
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct FormatPayloadIter<'source> {
    source: &'source [u8],
    cursor: usize,
}

impl<'source> Iterator for FormatPayloadIter<'source> {
    type Item = Result<&'source [u8], DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            let field = match next_raw_field(self.source, &mut self.cursor) {
                Ok(Some(field)) => field,
                Ok(None) => return None,
                Err(error) => return Some(Err(error)),
            };
            if field.number == LIST_CUSTOM_FORMATS_FIELD {
                if field.wire != 2 {
                    return Some(Err(DecodeError::invalid()));
                }
                return Some(field.payload.ok_or_else(DecodeError::invalid));
            }
        }
    }
}

fn write_counts(write: CustomFormatWrite<'_, '_>) -> Result<WriteCounts, DecodeError> {
    let mut fields = format_struct_canonical_fields();
    let mut work = format_struct_canonical_len(write.default_format)?;
    let mut text = write.name.len();
    for condition in write.conditions {
        let condition = *condition;
        let nested = format_struct_canonical_len(condition.condition_format)?;
        let condition_len = condition_canonical_len(condition)?;
        fields = fields
            .checked_add(4)
            .and_then(|value| value.checked_add(format_struct_canonical_fields()))
            .ok_or_else(DecodeError::invalid)?;
        work = work
            .checked_add(condition_len)
            .and_then(|value| value.checked_add(nested))
            .ok_or_else(DecodeError::invalid)?;
        text = text
            .checked_add(condition.condition_format.custom_format_string().len())
            .ok_or_else(DecodeError::invalid)?;
    }
    fields = fields.checked_add(4).ok_or_else(DecodeError::invalid)?;
    let archive_len = custom_format_canonical_len(write)?;
    work = work
        .checked_add(archive_len)
        .ok_or_else(DecodeError::invalid)?;
    text = text
        .checked_add(write.default_format.custom_format_string().len())
        .ok_or_else(DecodeError::invalid)?;
    Ok(WriteCounts {
        fields,
        work,
        max_depth: if write.conditions.is_empty() { 1 } else { 2 },
        conditions: write.conditions.len(),
        text_bytes: text,
    })
}

/// Conservative backing-storage and logical-allocation requirements for one
/// private Buffa lazy repeated-bytes view.
fn repeated_view_materialization(count: usize) -> Result<(usize, usize), DecodeError> {
    if count == 0 {
        return Ok((0, 0));
    }
    let capacity = count
        .checked_next_power_of_two()
        .ok_or_else(DecodeError::invalid)?
        .max(4);
    let bytes = capacity
        .checked_mul(size_of::<&[u8]>())
        .ok_or_else(DecodeError::invalid)?;
    Ok((1, bytes))
}

/// Materialization requirements for a generated custom archive view.  Only
/// the repeated condition field owns a temporary `Vec`; singular bytes and
/// scalar fields stay borrowed or inline in the generated view.
fn custom_format_materialization(condition_count: usize) -> Result<(usize, usize), DecodeError> {
    repeated_view_materialization(condition_count)
}

/// Materialization requirements for a generated custom registry view and all
/// generated archive views forced during its parity check.
fn custom_format_list_materialization(
    write: CustomFormatListWrite<'_, '_>,
) -> Result<(usize, usize), DecodeError> {
    let mut allocations = 0usize;
    let mut scratch_bytes = 0usize;
    for count in [write.uuids().len(), write.custom_formats().len()] {
        let (count_allocations, count_scratch) = repeated_view_materialization(count)?;
        allocations = allocations
            .checked_add(count_allocations)
            .ok_or_else(DecodeError::invalid)?;
        scratch_bytes = scratch_bytes
            .checked_add(count_scratch)
            .ok_or_else(DecodeError::invalid)?;
    }
    for custom_format in write.custom_formats() {
        let (count_allocations, count_scratch) =
            custom_format_materialization(custom_format.conditions().len())?;
        allocations = allocations
            .checked_add(count_allocations)
            .ok_or_else(DecodeError::invalid)?;
        scratch_bytes = scratch_bytes
            .checked_add(count_scratch)
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok((allocations, scratch_bytes))
}

fn custom_format_list_write_text_bytes(
    write: CustomFormatListWrite<'_, '_>,
) -> Result<usize, DecodeError> {
    let mut text = 0usize;
    for custom_format in write.custom_formats() {
        text = text
            .checked_add(write_counts(*custom_format)?.text_bytes)
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(text)
}

#[derive(Debug, Clone, Copy)]
struct WriteCounts {
    fields: usize,
    work: usize,
    max_depth: u32,
    conditions: usize,
    text_bytes: usize,
}

fn custom_format_matches_write(
    snapshot: CustomFormatSnapshot<'_>,
    write: CustomFormatWrite<'_, '_>,
) -> Result<bool, DecodeError> {
    if snapshot.name() != write.name()
        || snapshot.format_type() != write.format_type()
        || !format_values_match_write(snapshot.default_format(), write.default_format())
        || snapshot.condition_count() != write.conditions().len()
    {
        return Ok(false);
    }
    for (actual, expected) in snapshot.conditions().zip(write.conditions()) {
        let actual = actual?;
        if !condition_matches_write(actual, *expected)? {
            return Ok(false);
        }
    }
    Ok(true)
}

fn condition_matches_write(
    actual: CustomConditionSnapshot<'_>,
    expected: CustomConditionWrite<'_>,
) -> Result<bool, DecodeError> {
    Ok(actual.condition_type() == expected.condition_type()
        && actual.condition_value() == expected.condition_value()
        && actual.condition_value_dbl() == expected.condition_value_dbl()
        && format_values_match_write(actual.condition_format(), expected.condition_format()))
}

fn format_values_match_write(
    snapshot: FormatStructSnapshot<'_>,
    write: FormatStructWrite<'_>,
) -> bool {
    snapshot.format_type() == write.format_type()
        && snapshot.show_thousands_separator() == write.show_thousands_separator()
        && snapshot.use_accounting_style() == write.use_accounting_style()
        && snapshot.fraction_accuracy() == write.fraction_accuracy()
        && snapshot.custom_format_string() == write.custom_format_string()
        && snapshot.scale_factor() == write.scale_factor()
        && snapshot.requires_fraction_replacement() == write.requires_fraction_replacement()
        && snapshot.decimal_width() == write.decimal_width()
        && snapshot.min_integer_width() == write.min_integer_width()
        && snapshot.num_nonspace_integer_digits() == write.num_nonspace_integer_digits()
        && snapshot.num_nonspace_decimal_digits() == write.num_nonspace_decimal_digits()
        && snapshot.index_from_right_last_integer() == write.index_from_right_last_integer()
        && snapshot.num_hash_decimal_digits() == write.num_hash_decimal_digits()
        && snapshot.total_num_decimal_digits() == write.total_num_decimal_digits()
        && snapshot.is_complex() == write.is_complex()
        && snapshot.contains_integer_token() == write.contains_integer_token()
}

fn custom_format_list_matches_write(
    snapshot: CustomFormatListSnapshot<'_>,
    write: CustomFormatListWrite<'_, '_>,
) -> Result<bool, DecodeError> {
    if snapshot.uuid_count() != write.uuids().len()
        || snapshot.format_count() != write.custom_formats().len()
    {
        return Ok(false);
    }
    for (actual, (expected_uuid, expected_format)) in snapshot.entries().zip(
        write
            .uuids()
            .iter()
            .copied()
            .zip(write.custom_formats().iter().copied()),
    ) {
        let actual = actual?;
        if actual.uuid() != expected_uuid
            || !custom_format_matches_write(actual.custom_format(), expected_format)?
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn format_struct_canonical_fields() -> usize {
    16
}

fn custom_format_canonical_len(write: CustomFormatWrite<'_, '_>) -> Result<usize, DecodeError> {
    let mut length = length_field_len(ARCHIVE_NAME_FIELD, write.name.as_bytes())?;
    length = length
        .checked_add(varint_field_len(
            ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD,
            u64::from(write.format_type_pre_bnc),
        )?)
        .ok_or_else(DecodeError::invalid)?;
    let default_length = format_struct_canonical_len(write.default_format)?;
    length = length
        .checked_add(nested_field_len(
            ARCHIVE_DEFAULT_FORMAT_FIELD,
            default_length,
        )?)
        .ok_or_else(DecodeError::invalid)?;
    for condition in write.conditions {
        let condition_length = condition_canonical_len(*condition)?;
        length = length
            .checked_add(nested_field_len(
                ARCHIVE_CONDITIONS_FIELD,
                condition_length,
            )?)
            .ok_or_else(DecodeError::invalid)?;
    }
    length
        .checked_add(varint_field_len(
            ARCHIVE_FORMAT_TYPE_FIELD,
            u64::from(write.format_type),
        )?)
        .ok_or_else(DecodeError::invalid)
}

fn custom_format_rewrite_len(
    source: &[u8],
    write: CustomFormatWrite<'_, '_>,
) -> Result<usize, DecodeError> {
    let mut cursor = 0usize;
    let mut length = 0usize;
    let mut condition = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        let field_len = match field.number {
            ARCHIVE_NAME_FIELD => length_field_len(ARCHIVE_NAME_FIELD, write.name.as_bytes())?,
            ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD => varint_field_len(
                ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD,
                u64::from(write.format_type_pre_bnc),
            )?,
            ARCHIVE_DEFAULT_FORMAT_FIELD => nested_field_len(
                ARCHIVE_DEFAULT_FORMAT_FIELD,
                format_struct_rewrite_len(
                    field.payload.ok_or_else(DecodeError::invalid)?,
                    write.default_format,
                )?,
            )?,
            ARCHIVE_CONDITIONS_FIELD => {
                let current = write
                    .conditions
                    .get(condition)
                    .ok_or_else(DecodeError::invalid)?;
                condition += 1;
                nested_field_len(
                    ARCHIVE_CONDITIONS_FIELD,
                    condition_rewrite_len(
                        field.payload.ok_or_else(DecodeError::invalid)?,
                        *current,
                    )?,
                )?
            },
            ARCHIVE_FORMAT_TYPE_FIELD => {
                varint_field_len(ARCHIVE_FORMAT_TYPE_FIELD, u64::from(write.format_type))?
            },
            _ => field
                .end
                .checked_sub(field.start)
                .ok_or_else(DecodeError::invalid)?,
        };
        length = length
            .checked_add(field_len)
            .ok_or_else(DecodeError::invalid)?;
    }
    if condition != write.conditions.len() {
        return Err(DecodeError::invalid());
    }
    Ok(length)
}

fn custom_format_list_canonical_len(
    write: CustomFormatListWrite<'_, '_>,
) -> Result<usize, DecodeError> {
    let mut length = 0usize;
    for uuid in write.uuids {
        let uuid_length = uuid_canonical_len(*uuid)?;
        length = length
            .checked_add(nested_field_len(LIST_UUIDS_FIELD, uuid_length)?)
            .ok_or_else(DecodeError::invalid)?;
    }
    for custom_format in write.custom_formats {
        let format_length = custom_format_canonical_len(*custom_format)?;
        length = length
            .checked_add(nested_field_len(LIST_CUSTOM_FORMATS_FIELD, format_length)?)
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn custom_format_list_rewrite_len(
    source: &[u8],
    write: CustomFormatListWrite<'_, '_>,
) -> Result<usize, DecodeError> {
    let mut cursor = 0usize;
    let mut length = 0usize;
    let mut uuid = 0usize;
    let mut format = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        let field_len = match field.number {
            LIST_UUIDS_FIELD => {
                let value = *write.uuids.get(uuid).ok_or_else(DecodeError::invalid)?;
                uuid += 1;
                nested_field_len(LIST_UUIDS_FIELD, uuid_canonical_len(value)?)?
            },
            LIST_CUSTOM_FORMATS_FIELD => {
                let value = *write
                    .custom_formats
                    .get(format)
                    .ok_or_else(DecodeError::invalid)?;
                format += 1;
                nested_field_len(
                    LIST_CUSTOM_FORMATS_FIELD,
                    custom_format_rewrite_len(
                        field.payload.ok_or_else(DecodeError::invalid)?,
                        value,
                    )?,
                )?
            },
            _ => field
                .end
                .checked_sub(field.start)
                .ok_or_else(DecodeError::invalid)?,
        };
        length = length
            .checked_add(field_len)
            .ok_or_else(DecodeError::invalid)?;
    }
    if uuid != write.uuids.len() || format != write.custom_formats.len() {
        return Err(DecodeError::invalid());
    }
    Ok(length)
}

fn format_struct_canonical_len(write: FormatStructWrite<'_>) -> Result<usize, DecodeError> {
    let mut length = varint_field_len(FORMAT_TYPE_FIELD, u64::from(write.format_type))?;
    for (number, value) in [
        (
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
            u64::from(write.show_thousands_separator),
        ),
        (
            FORMAT_USE_ACCOUNTING_STYLE_FIELD,
            u64::from(write.use_accounting_style),
        ),
        (
            FORMAT_FRACTION_ACCURACY_FIELD,
            u64::from(write.fraction_accuracy),
        ),
        (
            FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD,
            u64::from(write.requires_fraction_replacement),
        ),
        (FORMAT_DECIMAL_WIDTH_FIELD, u64::from(write.decimal_width)),
        (
            FORMAT_MIN_INTEGER_WIDTH_FIELD,
            u64::from(write.min_integer_width),
        ),
        (
            FORMAT_NONSPACE_INTEGER_DIGITS_FIELD,
            u64::from(write.num_nonspace_integer_digits),
        ),
        (
            FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD,
            u64::from(write.num_nonspace_decimal_digits),
        ),
        (
            FORMAT_INDEX_FROM_RIGHT_FIELD,
            u64::from(write.index_from_right_last_integer),
        ),
        (
            FORMAT_HASH_DECIMAL_DIGITS_FIELD,
            u64::from(write.num_hash_decimal_digits),
        ),
        (
            FORMAT_TOTAL_DECIMAL_DIGITS_FIELD,
            u64::from(write.total_num_decimal_digits),
        ),
        (FORMAT_IS_COMPLEX_FIELD, u64::from(write.is_complex)),
        (
            FORMAT_CONTAINS_INTEGER_TOKEN_FIELD,
            u64::from(write.contains_integer_token),
        ),
    ] {
        length = length
            .checked_add(varint_field_len(number, value)?)
            .ok_or_else(DecodeError::invalid)?;
    }
    length = length
        .checked_add(length_field_len(
            FORMAT_CUSTOM_FORMAT_STRING_FIELD,
            write.custom_format_string.as_bytes(),
        )?)
        .ok_or_else(DecodeError::invalid)?;
    length = length
        .checked_add(fixed64_field_len(FORMAT_SCALE_FACTOR_FIELD))
        .ok_or_else(DecodeError::invalid)?;
    Ok(length)
}

fn format_struct_rewrite_len(
    source: &[u8],
    write: FormatStructWrite<'_>,
) -> Result<usize, DecodeError> {
    let mut cursor = 0usize;
    let mut length = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        let value = match field.number {
            FORMAT_TYPE_FIELD => varint_field_len(FORMAT_TYPE_FIELD, u64::from(write.format_type))?,
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => varint_field_len(
                FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
                u64::from(write.show_thousands_separator),
            )?,
            FORMAT_USE_ACCOUNTING_STYLE_FIELD => varint_field_len(
                FORMAT_USE_ACCOUNTING_STYLE_FIELD,
                u64::from(write.use_accounting_style),
            )?,
            FORMAT_FRACTION_ACCURACY_FIELD => varint_field_len(
                FORMAT_FRACTION_ACCURACY_FIELD,
                u64::from(write.fraction_accuracy),
            )?,
            FORMAT_CUSTOM_FORMAT_STRING_FIELD => length_field_len(
                FORMAT_CUSTOM_FORMAT_STRING_FIELD,
                write.custom_format_string.as_bytes(),
            )?,
            FORMAT_SCALE_FACTOR_FIELD => fixed64_field_len(FORMAT_SCALE_FACTOR_FIELD),
            FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD => varint_field_len(
                FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD,
                u64::from(write.requires_fraction_replacement),
            )?,
            FORMAT_DECIMAL_WIDTH_FIELD => {
                varint_field_len(FORMAT_DECIMAL_WIDTH_FIELD, u64::from(write.decimal_width))?
            },
            FORMAT_MIN_INTEGER_WIDTH_FIELD => varint_field_len(
                FORMAT_MIN_INTEGER_WIDTH_FIELD,
                u64::from(write.min_integer_width),
            )?,
            FORMAT_NONSPACE_INTEGER_DIGITS_FIELD => varint_field_len(
                FORMAT_NONSPACE_INTEGER_DIGITS_FIELD,
                u64::from(write.num_nonspace_integer_digits),
            )?,
            FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD => varint_field_len(
                FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD,
                u64::from(write.num_nonspace_decimal_digits),
            )?,
            FORMAT_INDEX_FROM_RIGHT_FIELD => varint_field_len(
                FORMAT_INDEX_FROM_RIGHT_FIELD,
                u64::from(write.index_from_right_last_integer),
            )?,
            FORMAT_HASH_DECIMAL_DIGITS_FIELD => varint_field_len(
                FORMAT_HASH_DECIMAL_DIGITS_FIELD,
                u64::from(write.num_hash_decimal_digits),
            )?,
            FORMAT_TOTAL_DECIMAL_DIGITS_FIELD => varint_field_len(
                FORMAT_TOTAL_DECIMAL_DIGITS_FIELD,
                u64::from(write.total_num_decimal_digits),
            )?,
            FORMAT_IS_COMPLEX_FIELD => {
                varint_field_len(FORMAT_IS_COMPLEX_FIELD, u64::from(write.is_complex))?
            },
            FORMAT_CONTAINS_INTEGER_TOKEN_FIELD => varint_field_len(
                FORMAT_CONTAINS_INTEGER_TOKEN_FIELD,
                u64::from(write.contains_integer_token),
            )?,
            _ => field
                .end
                .checked_sub(field.start)
                .ok_or_else(DecodeError::invalid)?,
        };
        length = length.checked_add(value).ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn condition_canonical_len(write: CustomConditionWrite<'_>) -> Result<usize, DecodeError> {
    let mut length = varint_field_len(CONDITION_TYPE_FIELD, u64::from(write.condition_type))?;
    if let Some(value) = write.condition_value {
        length = length
            .checked_add(fixed32_field_len(CONDITION_VALUE_FIELD))
            .ok_or_else(DecodeError::invalid)?;
        let _ = value;
    }
    let format_length = format_struct_canonical_len(write.condition_format)?;
    length = length
        .checked_add(nested_field_len(CONDITION_FORMAT_FIELD, format_length)?)
        .ok_or_else(DecodeError::invalid)?;
    if write.condition_value_dbl.is_some() {
        length = length
            .checked_add(fixed64_field_len(CONDITION_VALUE_DBL_FIELD))
            .ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn condition_rewrite_len(
    source: &[u8],
    write: CustomConditionWrite<'_>,
) -> Result<usize, DecodeError> {
    let mut cursor = 0usize;
    let mut length = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        let value = match field.number {
            CONDITION_TYPE_FIELD => {
                varint_field_len(CONDITION_TYPE_FIELD, u64::from(write.condition_type))?
            },
            CONDITION_VALUE_FIELD => fixed32_field_len(CONDITION_VALUE_FIELD),
            CONDITION_FORMAT_FIELD => nested_field_len(
                CONDITION_FORMAT_FIELD,
                format_struct_rewrite_len(
                    field.payload.ok_or_else(DecodeError::invalid)?,
                    write.condition_format,
                )?,
            )?,
            CONDITION_VALUE_DBL_FIELD => fixed64_field_len(CONDITION_VALUE_DBL_FIELD),
            _ => field
                .end
                .checked_sub(field.start)
                .ok_or_else(DecodeError::invalid)?,
        };
        length = length.checked_add(value).ok_or_else(DecodeError::invalid)?;
    }
    Ok(length)
}

fn uuid_canonical_len(uuid: Uuid) -> Result<usize, DecodeError> {
    varint_field_len(UUID_LOWER_FIELD, uuid.lower).and_then(|length| {
        varint_field_len(UUID_UPPER_FIELD, uuid.upper)
            .and_then(|second| length.checked_add(second).ok_or_else(DecodeError::invalid))
    })
}

fn nested_field_len(number: u32, payload_len: usize) -> Result<usize, DecodeError> {
    let key = key_len(number)?;
    let length = varint_len_usize(payload_len)?;
    key.checked_add(length)
        .and_then(|value| value.checked_add(payload_len))
        .ok_or_else(DecodeError::invalid)
}

fn length_field_len(number: u32, payload: &[u8]) -> Result<usize, DecodeError> {
    nested_field_len(number, payload.len())
}

fn varint_field_len(number: u32, value: u64) -> Result<usize, DecodeError> {
    key_len(number)?
        .checked_add(varint_len(value))
        .ok_or_else(DecodeError::invalid)
}

const fn fixed32_field_len(number: u32) -> usize {
    // All selected field numbers fit in a one-byte key; keep the generic
    // helper arithmetic explicit for future extension fields.
    if number < 16 { 5 } else { 6 }
}

fn fixed64_field_len(number: u32) -> usize {
    key_len(number).unwrap_or(usize::MAX).saturating_add(8)
}

fn key_len(number: u32) -> Result<usize, DecodeError> {
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    Ok(varint_len(u64::from(number) << 3))
}

const fn varint_len(value: u64) -> usize {
    if value < (1u64 << 7) {
        1
    } else if value < (1u64 << 14) {
        2
    } else if value < (1u64 << 21) {
        3
    } else if value < (1u64 << 28) {
        4
    } else if value < (1u64 << 35) {
        5
    } else if value < (1u64 << 42) {
        6
    } else if value < (1u64 << 49) {
        7
    } else if value < (1u64 << 56) {
        8
    } else if value < (1u64 << 63) {
        9
    } else {
        10
    }
}

fn varint_len_usize(value: usize) -> Result<usize, DecodeError> {
    Ok(varint_len(
        u64::try_from(value).map_err(|_| DecodeError::invalid())?,
    ))
}

fn rewrite_requirements(
    source: &[u8],
    source_report: DecodeReport,
    output_bytes: usize,
    formats: usize,
    conditions: usize,
    references: usize,
    text_bytes: usize,
    candidate_materialization: (usize, usize),
) -> Result<RewriteExecutionRequirements, DecodeError> {
    let candidate_work = output_bytes
        .checked_mul(2)
        .ok_or_else(DecodeError::invalid)?;
    // A source-preserving rewrite emits one field for each preflighted field;
    // scratch/work is conservatively doubled below, but the published field
    // count must describe the candidate itself for post-write verification.
    let fields = source_report.fields();
    let work_bytes = source_report
        .work_bytes()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(candidate_work))
        .ok_or_else(DecodeError::invalid)?;
    let retained = source
        .len()
        .checked_add(output_bytes)
        .ok_or_else(DecodeError::invalid)?;
    let allocations = source_report
        .allocations()
        .checked_add(candidate_materialization.0)
        .and_then(|value| value.checked_add(1))
        .ok_or_else(DecodeError::invalid)?;
    let scratch_bytes = source_report
        .scratch_bytes()
        .checked_add(candidate_materialization.1)
        .ok_or_else(DecodeError::invalid)?;
    Ok(RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes,
        max_depth: source_report.max_depth(),
        references,
        formats,
        conditions,
        text_bytes,
        allocations,
        retained_bytes: retained,
        scratch_bytes,
    })
}

fn canonical_requirements(
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    references: usize,
    formats: usize,
    conditions: usize,
    text_bytes: usize,
    retained_bytes: usize,
    candidate_materialization: (usize, usize),
) -> Result<RewriteExecutionRequirements, DecodeError> {
    let allocations = candidate_materialization
        .0
        .checked_add(1)
        .ok_or_else(DecodeError::invalid)?;
    Ok(RewriteExecutionRequirements {
        output_bytes,
        fields,
        work_bytes: work_bytes
            .checked_add(output_bytes)
            .ok_or_else(DecodeError::invalid)?,
        max_depth,
        references,
        formats,
        conditions,
        text_bytes,
        allocations,
        retained_bytes,
        scratch_bytes: candidate_materialization.1,
    })
}

fn report_from_requirements(requirements: RewriteExecutionRequirements) -> DecodeReport {
    DecodeReport {
        input_bytes: 0,
        output_bytes: requirements.output_bytes,
        fields: requirements.fields,
        work_bytes: requirements.work_bytes,
        max_depth: requirements.max_depth,
        references: requirements.references,
        formats: requirements.formats,
        conditions: requirements.conditions,
        text_bytes: requirements.text_bytes,
        allocations: requirements.allocations,
        retained_bytes: requirements.retained_bytes,
        scratch_bytes: requirements.scratch_bytes,
    }
}

fn check_options(
    requirements: RewriteExecutionRequirements,
    options: DecodeOptions,
) -> Result<(), DecodeError> {
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
    if requirements.formats > options.max_items {
        return Err(DecodeError::limited(DecodeLimit::Formats {
            observed: requirements.formats,
            maximum: options.max_items,
        }));
    }
    if requirements.conditions > options.max_items {
        return Err(DecodeError::limited(DecodeLimit::Conditions {
            observed: requirements.conditions,
            maximum: options.max_items,
        }));
    }
    if requirements.text_bytes > options.max_text_bytes {
        return Err(DecodeError::limited(DecodeLimit::Text {
            observed: requirements.text_bytes,
            maximum: options.max_text_bytes,
        }));
    }
    if requirements.retained_bytes
        > options
            .max_output_bytes
            .saturating_add(options.max_message_bytes)
    {
        return Err(DecodeError::limited(DecodeLimit::Retained {
            observed: requirements.retained_bytes,
            maximum: options
                .max_output_bytes
                .saturating_add(options.max_message_bytes),
        }));
    }
    Ok(())
}

fn check_requirements(
    requirements: RewriteExecutionRequirements,
    limits: RewriteExecutionLimits,
) -> Result<(), DecodeError> {
    let checks = [
        (
            requirements.output_bytes,
            limits.output_bytes,
            DecodeLimit::OutputBytes {
                observed: requirements.output_bytes,
                maximum: limits.output_bytes,
            },
        ),
        (
            requirements.fields,
            limits.fields,
            DecodeLimit::Fields {
                observed: requirements.fields,
                maximum: limits.fields,
            },
        ),
        (
            requirements.work_bytes,
            limits.work_bytes,
            DecodeLimit::Work {
                observed: requirements.work_bytes,
                maximum: limits.work_bytes,
            },
        ),
        (
            usize::try_from(requirements.max_depth).unwrap_or(usize::MAX),
            usize::try_from(limits.max_depth).unwrap_or(usize::MAX),
            DecodeLimit::Nesting {
                observed: requirements.max_depth,
                maximum: limits.max_depth,
            },
        ),
        (
            requirements.references,
            limits.references,
            DecodeLimit::References {
                observed: requirements.references,
                maximum: limits.references,
            },
        ),
        (
            requirements.formats,
            limits.formats,
            DecodeLimit::Formats {
                observed: requirements.formats,
                maximum: limits.formats,
            },
        ),
        (
            requirements.conditions,
            limits.conditions,
            DecodeLimit::Conditions {
                observed: requirements.conditions,
                maximum: limits.conditions,
            },
        ),
        (
            requirements.text_bytes,
            limits.text_bytes,
            DecodeLimit::Text {
                observed: requirements.text_bytes,
                maximum: limits.text_bytes,
            },
        ),
        (
            requirements.allocations,
            limits.allocations,
            DecodeLimit::Allocation {
                requested: requirements.output_bytes,
            },
        ),
        (
            requirements.retained_bytes,
            limits.retained_bytes,
            DecodeLimit::Retained {
                observed: requirements.retained_bytes,
                maximum: limits.retained_bytes,
            },
        ),
        (
            requirements.scratch_bytes,
            limits.scratch_bytes,
            DecodeLimit::Scratch {
                observed: requirements.scratch_bytes,
                maximum: limits.scratch_bytes,
            },
        ),
    ];
    for (observed, maximum, error) in checks {
        if observed > maximum {
            return Err(DecodeError::limited(error));
        }
    }
    Ok(())
}

fn emit_custom_format_canonical(
    output: &mut Vec<u8>,
    write: CustomFormatWrite<'_, '_>,
) -> Result<(), DecodeError> {
    emit_length_field(output, ARCHIVE_NAME_FIELD, write.name.as_bytes())?;
    emit_varint_field(
        output,
        ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD,
        u64::from(write.format_type_pre_bnc),
    )?;
    emit_nested_format_field(output, ARCHIVE_DEFAULT_FORMAT_FIELD, write.default_format)?;
    for condition in write.conditions {
        emit_nested_condition_field(output, *condition)?;
    }
    emit_varint_field(
        output,
        ARCHIVE_FORMAT_TYPE_FIELD,
        u64::from(write.format_type),
    )
}

fn emit_custom_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    write: CustomFormatWrite<'_, '_>,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    let mut condition = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        match field.number {
            ARCHIVE_NAME_FIELD => {
                emit_length_field(output, ARCHIVE_NAME_FIELD, write.name.as_bytes())?
            },
            ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD => emit_varint_field(
                output,
                ARCHIVE_FORMAT_TYPE_PRE_BNC_FIELD,
                u64::from(write.format_type_pre_bnc),
            )?,
            ARCHIVE_DEFAULT_FORMAT_FIELD => emit_nested_format_rewrite(
                output,
                field.payload.ok_or_else(DecodeError::invalid)?,
                write.default_format,
            )?,
            ARCHIVE_CONDITIONS_FIELD => {
                let current = *write
                    .conditions
                    .get(condition)
                    .ok_or_else(DecodeError::invalid)?;
                condition += 1;
                emit_nested_condition_rewrite(
                    output,
                    field.payload.ok_or_else(DecodeError::invalid)?,
                    current,
                )?;
            },
            ARCHIVE_FORMAT_TYPE_FIELD => emit_varint_field(
                output,
                ARCHIVE_FORMAT_TYPE_FIELD,
                u64::from(write.format_type),
            )?,
            _ => output.extend_from_slice(
                source
                    .get(field.start..field.end)
                    .ok_or_else(DecodeError::invalid)?,
            ),
        }
    }
    if condition != write.conditions.len() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn emit_custom_format_list_canonical(
    output: &mut Vec<u8>,
    write: CustomFormatListWrite<'_, '_>,
) -> Result<(), DecodeError> {
    for uuid in write.uuids {
        emit_nested_uuid_field(output, *uuid)?;
    }
    for custom_format in write.custom_formats {
        let length = custom_format_canonical_len(*custom_format)?;
        emit_key(output, LIST_CUSTOM_FORMATS_FIELD, 2)?;
        emit_varint(
            output,
            u64::try_from(length).map_err(|_| DecodeError::invalid())?,
        );
        emit_custom_format_canonical(output, *custom_format)?;
    }
    Ok(())
}

fn emit_custom_format_list_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    write: CustomFormatListWrite<'_, '_>,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    let mut uuid = 0usize;
    let mut format = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        match field.number {
            LIST_UUIDS_FIELD => {
                let value = *write.uuids.get(uuid).ok_or_else(DecodeError::invalid)?;
                uuid += 1;
                emit_nested_uuid_field(output, value)?;
            },
            LIST_CUSTOM_FORMATS_FIELD => {
                let value = *write
                    .custom_formats
                    .get(format)
                    .ok_or_else(DecodeError::invalid)?;
                format += 1;
                let payload = field.payload.ok_or_else(DecodeError::invalid)?;
                let length = custom_format_rewrite_len(payload, value)?;
                emit_key(output, LIST_CUSTOM_FORMATS_FIELD, 2)?;
                emit_varint(
                    output,
                    u64::try_from(length).map_err(|_| DecodeError::invalid())?,
                );
                emit_custom_format_rewrite(output, payload, value)?;
            },
            _ => output.extend_from_slice(
                source
                    .get(field.start..field.end)
                    .ok_or_else(DecodeError::invalid)?,
            ),
        }
    }
    if uuid != write.uuids.len() || format != write.custom_formats.len() {
        return Err(DecodeError::invalid());
    }
    Ok(())
}

fn emit_nested_uuid_field(output: &mut Vec<u8>, uuid: Uuid) -> Result<(), DecodeError> {
    let length = uuid_canonical_len(uuid)?;
    emit_key(output, LIST_UUIDS_FIELD, 2)?;
    emit_varint(
        output,
        u64::try_from(length).map_err(|_| DecodeError::invalid())?,
    );
    emit_varint_field(output, UUID_LOWER_FIELD, uuid.lower)?;
    emit_varint_field(output, UUID_UPPER_FIELD, uuid.upper)
}

fn emit_nested_format_field(
    output: &mut Vec<u8>,
    number: u32,
    write: FormatStructWrite<'_>,
) -> Result<(), DecodeError> {
    let length = format_struct_canonical_len(write)?;
    emit_key(output, number, 2)?;
    emit_varint(
        output,
        u64::try_from(length).map_err(|_| DecodeError::invalid())?,
    );
    emit_format_canonical(output, write)
}

fn emit_nested_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    write: FormatStructWrite<'_>,
) -> Result<(), DecodeError> {
    let length = format_struct_rewrite_len(source, write)?;
    emit_key(output, ARCHIVE_DEFAULT_FORMAT_FIELD, 2)?;
    emit_varint(
        output,
        u64::try_from(length).map_err(|_| DecodeError::invalid())?,
    );
    emit_format_rewrite(output, source, write)
}

fn emit_nested_condition_field(
    output: &mut Vec<u8>,
    write: CustomConditionWrite<'_>,
) -> Result<(), DecodeError> {
    let length = condition_canonical_len(write)?;
    emit_key(output, ARCHIVE_CONDITIONS_FIELD, 2)?;
    emit_varint(
        output,
        u64::try_from(length).map_err(|_| DecodeError::invalid())?,
    );
    emit_condition_canonical(output, write)
}

fn emit_nested_condition_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    write: CustomConditionWrite<'_>,
) -> Result<(), DecodeError> {
    let length = condition_rewrite_len(source, write)?;
    emit_key(output, ARCHIVE_CONDITIONS_FIELD, 2)?;
    emit_varint(
        output,
        u64::try_from(length).map_err(|_| DecodeError::invalid())?,
    );
    emit_condition_rewrite(output, source, write)
}

fn emit_format_canonical(
    output: &mut Vec<u8>,
    write: FormatStructWrite<'_>,
) -> Result<(), DecodeError> {
    emit_varint_field(output, FORMAT_TYPE_FIELD, u64::from(write.format_type))?;
    emit_varint_field(
        output,
        FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
        u64::from(write.show_thousands_separator),
    )?;
    emit_varint_field(
        output,
        FORMAT_USE_ACCOUNTING_STYLE_FIELD,
        u64::from(write.use_accounting_style),
    )?;
    emit_varint_field(
        output,
        FORMAT_FRACTION_ACCURACY_FIELD,
        u64::from(write.fraction_accuracy),
    )?;
    emit_length_field(
        output,
        FORMAT_CUSTOM_FORMAT_STRING_FIELD,
        write.custom_format_string.as_bytes(),
    )?;
    emit_fixed64_field(
        output,
        FORMAT_SCALE_FACTOR_FIELD,
        write.scale_factor.to_bits(),
    )?;
    emit_varint_field(
        output,
        FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD,
        u64::from(write.requires_fraction_replacement),
    )?;
    emit_varint_field(
        output,
        FORMAT_DECIMAL_WIDTH_FIELD,
        u64::from(write.decimal_width),
    )?;
    emit_varint_field(
        output,
        FORMAT_MIN_INTEGER_WIDTH_FIELD,
        u64::from(write.min_integer_width),
    )?;
    emit_varint_field(
        output,
        FORMAT_NONSPACE_INTEGER_DIGITS_FIELD,
        u64::from(write.num_nonspace_integer_digits),
    )?;
    emit_varint_field(
        output,
        FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD,
        u64::from(write.num_nonspace_decimal_digits),
    )?;
    emit_varint_field(
        output,
        FORMAT_INDEX_FROM_RIGHT_FIELD,
        u64::from(write.index_from_right_last_integer),
    )?;
    emit_varint_field(
        output,
        FORMAT_HASH_DECIMAL_DIGITS_FIELD,
        u64::from(write.num_hash_decimal_digits),
    )?;
    emit_varint_field(
        output,
        FORMAT_TOTAL_DECIMAL_DIGITS_FIELD,
        u64::from(write.total_num_decimal_digits),
    )?;
    emit_varint_field(output, FORMAT_IS_COMPLEX_FIELD, u64::from(write.is_complex))?;
    emit_varint_field(
        output,
        FORMAT_CONTAINS_INTEGER_TOKEN_FIELD,
        u64::from(write.contains_integer_token),
    )
}

fn emit_format_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    write: FormatStructWrite<'_>,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        match field.number {
            FORMAT_TYPE_FIELD => {
                emit_varint_field(output, FORMAT_TYPE_FIELD, u64::from(write.format_type))?
            },
            FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD => emit_varint_field(
                output,
                FORMAT_SHOW_THOUSANDS_SEPARATOR_FIELD,
                u64::from(write.show_thousands_separator),
            )?,
            FORMAT_USE_ACCOUNTING_STYLE_FIELD => emit_varint_field(
                output,
                FORMAT_USE_ACCOUNTING_STYLE_FIELD,
                u64::from(write.use_accounting_style),
            )?,
            FORMAT_FRACTION_ACCURACY_FIELD => emit_varint_field(
                output,
                FORMAT_FRACTION_ACCURACY_FIELD,
                u64::from(write.fraction_accuracy),
            )?,
            FORMAT_CUSTOM_FORMAT_STRING_FIELD => emit_length_field(
                output,
                FORMAT_CUSTOM_FORMAT_STRING_FIELD,
                write.custom_format_string.as_bytes(),
            )?,
            FORMAT_SCALE_FACTOR_FIELD => emit_fixed64_field(
                output,
                FORMAT_SCALE_FACTOR_FIELD,
                write.scale_factor.to_bits(),
            )?,
            FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD => emit_varint_field(
                output,
                FORMAT_REQUIRES_FRACTION_REPLACEMENT_FIELD,
                u64::from(write.requires_fraction_replacement),
            )?,
            FORMAT_DECIMAL_WIDTH_FIELD => emit_varint_field(
                output,
                FORMAT_DECIMAL_WIDTH_FIELD,
                u64::from(write.decimal_width),
            )?,
            FORMAT_MIN_INTEGER_WIDTH_FIELD => emit_varint_field(
                output,
                FORMAT_MIN_INTEGER_WIDTH_FIELD,
                u64::from(write.min_integer_width),
            )?,
            FORMAT_NONSPACE_INTEGER_DIGITS_FIELD => emit_varint_field(
                output,
                FORMAT_NONSPACE_INTEGER_DIGITS_FIELD,
                u64::from(write.num_nonspace_integer_digits),
            )?,
            FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD => emit_varint_field(
                output,
                FORMAT_NONSPACE_DECIMAL_DIGITS_FIELD,
                u64::from(write.num_nonspace_decimal_digits),
            )?,
            FORMAT_INDEX_FROM_RIGHT_FIELD => emit_varint_field(
                output,
                FORMAT_INDEX_FROM_RIGHT_FIELD,
                u64::from(write.index_from_right_last_integer),
            )?,
            FORMAT_HASH_DECIMAL_DIGITS_FIELD => emit_varint_field(
                output,
                FORMAT_HASH_DECIMAL_DIGITS_FIELD,
                u64::from(write.num_hash_decimal_digits),
            )?,
            FORMAT_TOTAL_DECIMAL_DIGITS_FIELD => emit_varint_field(
                output,
                FORMAT_TOTAL_DECIMAL_DIGITS_FIELD,
                u64::from(write.total_num_decimal_digits),
            )?,
            FORMAT_IS_COMPLEX_FIELD => {
                emit_varint_field(output, FORMAT_IS_COMPLEX_FIELD, u64::from(write.is_complex))?
            },
            FORMAT_CONTAINS_INTEGER_TOKEN_FIELD => emit_varint_field(
                output,
                FORMAT_CONTAINS_INTEGER_TOKEN_FIELD,
                u64::from(write.contains_integer_token),
            )?,
            _ => output.extend_from_slice(
                source
                    .get(field.start..field.end)
                    .ok_or_else(DecodeError::invalid)?,
            ),
        }
    }
    Ok(())
}

fn emit_condition_canonical(
    output: &mut Vec<u8>,
    write: CustomConditionWrite<'_>,
) -> Result<(), DecodeError> {
    emit_varint_field(
        output,
        CONDITION_TYPE_FIELD,
        u64::from(write.condition_type),
    )?;
    if let Some(value) = write.condition_value {
        emit_fixed32_field(output, CONDITION_VALUE_FIELD, value.to_bits())?;
    }
    emit_nested_format_field(output, CONDITION_FORMAT_FIELD, write.condition_format)?;
    if let Some(value) = write.condition_value_dbl {
        emit_fixed64_field(output, CONDITION_VALUE_DBL_FIELD, value.to_bits())?;
    }
    Ok(())
}

fn emit_condition_rewrite(
    output: &mut Vec<u8>,
    source: &[u8],
    write: CustomConditionWrite<'_>,
) -> Result<(), DecodeError> {
    let mut cursor = 0usize;
    while let Some(field) = next_raw_field(source, &mut cursor)? {
        match field.number {
            CONDITION_TYPE_FIELD => emit_varint_field(
                output,
                CONDITION_TYPE_FIELD,
                u64::from(write.condition_type),
            )?,
            CONDITION_VALUE_FIELD => {
                let value = write.condition_value.ok_or_else(DecodeError::invalid)?;
                emit_fixed32_field(output, CONDITION_VALUE_FIELD, value.to_bits())?;
            },
            CONDITION_FORMAT_FIELD => emit_nested_format_rewrite(
                output,
                field.payload.ok_or_else(DecodeError::invalid)?,
                write.condition_format,
            )?,
            CONDITION_VALUE_DBL_FIELD => {
                let value = write.condition_value_dbl.ok_or_else(DecodeError::invalid)?;
                emit_fixed64_field(output, CONDITION_VALUE_DBL_FIELD, value.to_bits())?;
            },
            _ => output.extend_from_slice(
                source
                    .get(field.start..field.end)
                    .ok_or_else(DecodeError::invalid)?,
            ),
        }
    }
    Ok(())
}

fn emit_key(output: &mut Vec<u8>, number: u32, wire: u8) -> Result<(), DecodeError> {
    emit_varint(output, (u64::from(number) << 3) | u64::from(wire));
    Ok(())
}

fn emit_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn emit_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), DecodeError> {
    emit_key(output, number, 0)?;
    emit_varint(output, value);
    Ok(())
}

fn emit_length_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), DecodeError> {
    emit_key(output, number, 2)?;
    emit_varint(
        output,
        u64::try_from(payload.len()).map_err(|_| DecodeError::invalid())?,
    );
    output.extend_from_slice(payload);
    Ok(())
}

fn emit_fixed32_field(output: &mut Vec<u8>, number: u32, value: u32) -> Result<(), DecodeError> {
    emit_key(output, number, 5)?;
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

fn emit_fixed64_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), DecodeError> {
    emit_key(output, number, 1)?;
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn options() -> DecodeOptions {
        DecodeOptions::new(
            64 * 1024,
            128 * 1024,
            16 * 1024,
            512 * 1024,
            64,
            128,
            128,
            16 * 1024,
        )
    }

    fn number_write<'a>(
        pattern: &'a str,
        conditions: &'a [CustomConditionWrite<'a>],
    ) -> CustomFormatWrite<'a, 'a> {
        CustomFormatWrite::new(
            "Custom",
            NATIVE_CUSTOM_NUMBER_FORMAT_TYPE,
            FormatStructWrite::new(NATIVE_CUSTOM_NUMBER_FORMAT_TYPE, pattern),
            conditions,
        )
    }

    fn canonical_number() -> Vec<u8> {
        let conditions: &[CustomConditionWrite<'_>] = &[];
        canonical_custom_format(number_write("#,##0.00", conditions), options())
            .expect("canonical custom")
            .into_bytes()
    }

    #[test]
    fn canonical_archive_roundtrips_selected_shape_and_borrows_pattern() {
        let source = canonical_number();
        let snapshot = decode_custom_format(&source, options()).expect("custom archive");
        assert_eq!(snapshot.name(), "Custom");
        assert_eq!(snapshot.format_type(), NATIVE_CUSTOM_NUMBER_FORMAT_TYPE);
        assert_eq!(snapshot.default_format().pattern(), "#,##0.00");
        assert!(
            snapshot
                .raw()
                .as_ptr_range()
                .contains(&snapshot.name().as_ptr())
        );
        assert!(
            snapshot
                .raw()
                .as_ptr_range()
                .contains(&snapshot.default_format().pattern().as_ptr())
        );
    }

    #[test]
    fn list_streams_parallel_entries_and_rejects_count_mismatch() {
        let uuid = Uuid::new(1, 2);
        let uuids = [uuid];
        let formats = [number_write("#,##0.00", &[])];
        let list = CustomFormatListWrite::new(&uuids, &formats);
        let source = canonical_custom_format_list(list, options())
            .expect("canonical list")
            .into_bytes();
        let snapshot = decode_custom_format_list(&source, options()).expect("list");
        let entries = snapshot
            .entries()
            .collect::<Result<Vec<_>, _>>()
            .expect("entries");
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].uuid(), uuid);
        assert_eq!(
            entries[0].custom_format().default_format().pattern(),
            "#,##0.00"
        );
        let mut mismatch = source;
        mismatch.extend_from_slice(&[0x0a, 0x02, 0x08, 0x01]);
        assert!(decode_custom_format_list(&mismatch, options()).is_err());
    }

    #[test]
    fn list_rejects_zero_and_duplicate_uuid_keys() {
        let uuids = [Uuid::new(1, 2)];
        let formats = [number_write("#,##0.00", &[])];
        let source =
            canonical_custom_format_list(CustomFormatListWrite::new(&uuids, &formats), options())
                .expect("canonical list")
                .into_bytes();
        // The canonical UUID envelope is the first six bytes.  A zero key is
        // invalid even when the parallel archive remains well formed.
        let mut zero = source.clone();
        zero[3] = 0;
        assert!(decode_custom_format_list(&zero, options()).is_err());

        // Add one duplicate UUID and one parallel archive so the count check
        // cannot mask the uniqueness check.
        let mut duplicate = source.clone();
        duplicate.extend_from_slice(&source[..6]);
        duplicate.extend_from_slice(&source[6..]);
        assert!(decode_custom_format_list(&duplicate, options()).is_err());
    }

    #[test]
    fn source_rewrite_keeps_unknown_scalar_and_group_spans() {
        let mut source = canonical_number();
        source.extend_from_slice(&[
            0xa0, 0x06, 0x81, 0x00, 0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06,
        ]);
        let write = number_write("#,##0.000", &[]);
        let prepared = prepare_custom_format_rewrite(&source, write, options()).expect("prepare");
        let output = prepared
            .execute(RewriteExecutionLimits::exact(
                prepared.execution_requirements(),
            ))
            .expect("rewrite");
        assert!(
            output
                .bytes()
                .windows(4)
                .any(|window| window == [0xa0, 0x06, 0x81, 0x00])
        );
        assert!(
            output
                .bytes()
                .windows(7)
                .any(|window| window == [0xa3, 0x06, 0xa8, 0x06, 0x01, 0xa4, 0x06])
        );
        assert_eq!(
            decode_custom_format(output.bytes(), options())
                .expect("readback")
                .default_format()
                .pattern(),
            "#,##0.000"
        );
    }

    #[test]
    fn conditions_are_bounded_and_semantically_checked() {
        let format = FormatStructWrite::new(NATIVE_CUSTOM_NUMBER_FORMAT_TYPE, "#,##0");
        let condition = CustomConditionWrite::with_double(1, 5.0, format);
        let source = canonical_custom_format(number_write("#,##0", &[condition]), options())
            .expect("canonical")
            .into_bytes();
        let snapshot = decode_custom_format(&source, options()).expect("decode");
        let condition = snapshot
            .conditions()
            .next()
            .expect("condition")
            .expect("valid");
        assert_eq!(condition.condition_type(), 1);
        assert_eq!(condition.threshold(), 5.0);
        assert!(decode_custom_format(&source, options().with_max_items(0)).is_err());

        // Change only the condition's nested family (270 -> 271).  The
        // archive's optional format type appears after the condition in the
        // canonical spelling, so this exercises the second-pass family check.
        let mut mismatched = source;
        let marker = [0x08, 0x8e, 0x02];
        let mut occurrences = mismatched
            .windows(marker.len())
            .enumerate()
            .filter_map(|(index, window)| (window == marker).then_some(index));
        let _default = occurrences.next().expect("default format type");
        let condition_type = occurrences.next().expect("condition format type");
        mismatched[condition_type + 1] = 0x8f;
        assert!(decode_custom_format(&mismatched, options()).is_err());
    }

    #[test]
    fn malformed_wire_and_cross_family_fields_fail_closed() {
        let valid = canonical_number();
        let mut duplicate = valid.clone();
        duplicate.extend_from_slice(&[0x08, 0xfe, 0x01]);
        assert!(decode_custom_format(&duplicate, options()).is_err());
        let mut wrong = valid;
        // `decimal_places` is a known non-custom field of FormatStructArchive.
        wrong.extend_from_slice(&[0x12, 0x01, 0x00]);
        assert!(decode_custom_format(&wrong, options()).is_err());
        assert!(decode_custom_format(&[0x0a, 0x01, 0xff], options()).is_err());

        let mut unclosed_group = canonical_number();
        unclosed_group.extend_from_slice(&[0xa3, 0x06, 0x08, 0x01]);
        assert!(decode_custom_format(&unclosed_group, options()).is_err());
    }

    #[test]
    fn buffa_repeated_views_are_charged_in_reports() {
        let format = FormatStructWrite::new(NATIVE_CUSTOM_NUMBER_FORMAT_TYPE, "#,##0");
        let condition = CustomConditionWrite::with_double(1, 5.0, format);
        let conditions = [condition];
        let archive = number_write("#,##0", &conditions);
        let source = canonical_custom_format(archive, options())
            .expect("canonical archive")
            .into_bytes();
        let (_, report) = decode_custom_format_with_report(&source, options()).expect("archive");
        assert!(report.allocations() > 0);
        assert!(report.scratch_bytes() > 0);
        assert_eq!(report.retained_bytes(), source.len());

        let uuid = Uuid::new(1, 2);
        let uuids = [uuid];
        let formats = [archive];
        let list_source =
            canonical_custom_format_list(CustomFormatListWrite::new(&uuids, &formats), options())
                .expect("canonical list")
                .into_bytes();
        let (_, list_report) =
            decode_custom_format_list_with_report(&list_source, options()).expect("list");
        // The generated list view owns two repeated byte vectors and the
        // nested archive view owns one repeated condition vector.
        assert!(list_report.allocations() >= 3);
        assert!(list_report.scratch_bytes() >= report.scratch_bytes());
        assert_eq!(list_report.retained_bytes(), list_source.len());
    }

    #[test]
    fn prepared_requirements_enforce_buffa_materialization_limits() {
        let format = FormatStructWrite::new(NATIVE_CUSTOM_NUMBER_FORMAT_TYPE, "#,##0");
        let condition = CustomConditionWrite::with_double(1, 5.0, format);
        let conditions = [condition];
        let write = number_write("#,##0", &conditions);
        let prepared = prepare_custom_format_write(write, options()).expect("prepare");
        let requirements = prepared.execution_requirements();
        assert!(requirements.allocations() > 0);
        assert!(requirements.scratch_bytes() > 0);
        assert!(requirements.retained_bytes() > 0);

        let output = prepared
            .execute(RewriteExecutionLimits::exact(requirements))
            .expect("exact execution");
        assert_eq!(output.report().allocations(), requirements.allocations());
        assert_eq!(
            output.report().scratch_bytes(),
            requirements.scratch_bytes()
        );
        assert_eq!(
            output.report().retained_bytes(),
            requirements.retained_bytes()
        );

        let allocation_limit = requirements.allocations() - 1;
        let error = prepared
            .execute(RewriteExecutionLimits::exact(requirements).with_allocations(allocation_limit))
            .expect_err("allocation limit must fail before emission");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Allocation { .. })
        ));

        let scratch_limit = requirements.scratch_bytes() - 1;
        let error = prepared
            .execute(RewriteExecutionLimits::exact(requirements).with_scratch_bytes(scratch_limit))
            .expect_err("scratch limit must fail before emission");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Scratch { .. })
        ));

        let retained_limit = requirements.retained_bytes() - 1;
        let error = prepared
            .execute(
                RewriteExecutionLimits::exact(requirements).with_retained_bytes(retained_limit),
            )
            .expect_err("retained limit must fail before emission");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Retained { .. })
        ));
    }
}
