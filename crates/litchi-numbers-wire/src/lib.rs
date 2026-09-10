//! Low-level Binary Numbers Cell (BNC) value storage for iWork adapters.
//!
//! This adapter crate keeps the byte-preserving codec shared while the legacy
//! IWA host is migrated into the standalone Numbers owner. It is intentionally
//! excluded from the `litchi` and `litchi-numbers` facades. Applications
//! should use `litchi-numbers`; direct use of this crate opts into unstable
//! native-storage details rather than the supported semantic API.

#![forbid(unsafe_code)]

/// Source-preserving Pop-Up Menu BNC transition planner and executor.
pub mod popup_menu;

/// Borrowed parser for legacy pre-BNC Numbers cell storage.
pub mod pre_bnc;

/// Allocation-free semantic projection of one Numbers cell payload.
pub mod cell_value;

/// Shared root/segment coordinator for Numbers table-data-list projections.
pub mod table_data_list;

pub mod formula_envelope;
/// Shared generated-free Numbers formula event renderer.
pub mod formula_render;

/// Shared bounded reader for native table merge formulas.
pub mod table_merges;

/// Shared native Numbers formula function-token registry.
pub mod function_map;

use std::collections::BTreeMap;

use std::fmt;

use litchi_iwa_common::formula::FiniteF64;

const BNC_VERSION: u8 = 5;
const BNC_PREFIX_LEN: usize = 8;
const BNC_HEADER_LEN: usize = 12;
const CELL_TYPE_EMPTY: u8 = 0;
const CELL_TYPE_NUMBER: u8 = 2;
const CELL_TYPE_TEXT: u8 = 3;
const CELL_TYPE_DATE: u8 = 5;
const CELL_TYPE_BOOLEAN: u8 = 6;
const CELL_TYPE_DURATION: u8 = 7;
const CELL_TYPE_ERROR: u8 = 8;
const CELL_TYPE_RICH_TEXT_OR_NUMBER: u8 = 9;
const CELL_TYPE_ALTERNATE_NUMBER: u8 = 10;
const DECIMAL128_EXPONENT_BIAS: i32 = 0x1820;
const DECIMAL128_COEFFICIENT_BITS: u32 = 113;
const DECIMAL128_SIGN_BIT: u32 = 127;
const SECONDS_PER_DAY: f64 = 86_400.0;

pub(crate) const DECIMAL_FLAG: u32 = 0x0000_0001;
pub(crate) const NUMBER_FLAG: u32 = 0x0000_0002;
pub(crate) const DATE_FLAG: u32 = 0x0000_0004;
pub(crate) const STRING_FLAG: u32 = 0x0000_0008;
pub(crate) const RICH_TEXT_FLAG: u32 = 0x0000_0010;
pub(crate) const STYLE_FLAG: u32 = 0x0000_0020;
pub(crate) const TEXT_STYLE_FLAG: u32 = 0x0000_0040;
pub(crate) const CONDITIONAL_STYLE_FLAG: u32 = 0x0000_0080;
pub(crate) const CONDITIONAL_STYLE_APPLIED_RULE_FLAG: u32 = 0x0000_0100;
pub(crate) const FORMULA_FLAG: u32 = 0x0000_0200;
const CONTROL_CELL_SPEC_FLAG: u32 = 0x0000_0400;
pub(crate) const FORMULA_ERROR_FLAG: u32 = 0x0000_0800;
pub(crate) const COMMENT_FLAG: u32 = 0x0008_0000;
const CELL_FORMAT_KIND_FLAG: u32 = 0x0000_1000;
const CELL_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_2000;
const CURRENCY_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_4000;
const DATE_TIME_FORMAT_IDENTIFIER_FLAG: u32 = 0x0000_8000;
const DURATION_FORMAT_IDENTIFIER_FLAG: u32 = 0x0001_0000;
const TEXT_FORMAT_IDENTIFIER_FLAG: u32 = 0x0002_0000;
const CHECKBOX_FORMAT_IDENTIFIER_FLAG: u32 = 0x0004_0000;
const FORMAT_METADATA_FLAGS: u32 = CONTROL_CELL_SPEC_FLAG
    | CELL_FORMAT_KIND_FLAG
    | CELL_FORMAT_IDENTIFIER_FLAG
    | CURRENCY_FORMAT_IDENTIFIER_FLAG
    | DATE_TIME_FORMAT_IDENTIFIER_FLAG
    | DURATION_FORMAT_IDENTIFIER_FLAG
    | TEXT_FORMAT_IDENTIFIER_FLAG
    | CHECKBOX_FORMAT_IDENTIFIER_FLAG;
const EXPLICIT_FORMAT_FLAGS_START: usize = 6;
const EXPLICIT_FORMAT_FLAGS_END: usize = 8;
pub const EXPLICIT_DECIMAL_FORMAT: u16 = 1;
/// Explicit Currency metadata marker written by native Numbers.
pub const EXPLICIT_CURRENCY_FORMAT: u16 = 0x0802;
/// Explicit Currency marker when the BNC cell also retains a Number format.
pub const EXPLICIT_CURRENCY_WITH_NUMBER_FORMAT: u16 =
    EXPLICIT_CURRENCY_FORMAT | EXPLICIT_DECIMAL_FORMAT;
pub const EXPLICIT_DATE_TIME_FORMAT: u16 = 0x0008;
/// Explicit Duration metadata marker for the native primary-only shape.
pub const EXPLICIT_DURATION_FORMAT: u16 = 0x0004;
/// Explicit Duration metadata marker when the BNC cell retains a shared
/// generic Number-format identifier as a secondary reference.
pub const EXPLICIT_DURATION_WITH_NUMBER_FORMAT: u16 = 0x0005;
pub const EXPLICIT_CHECKBOX_FORMAT: u16 = 0x0020;
pub const EXPLICIT_TEXT_FORMAT: u16 = 0x0080;
pub const EXPLICIT_CONVERTED_TEXT_FORMAT: u16 = EXPLICIT_TEXT_FORMAT | EXPLICIT_DECIMAL_FORMAT;
pub const DECIMAL_CELL_FORMAT_KIND: u32 = 1;
pub const CURRENCY_CELL_FORMAT_KIND: u32 = 2;
pub const DATE_TIME_CELL_FORMAT_KIND: u32 = 3;
pub const DURATION_CELL_FORMAT_KIND: u32 = 4;
pub const CHECKBOX_CELL_FORMAT_KIND: u32 = 6;
pub const STAR_RATING_CELL_FORMAT_KIND: u32 = DECIMAL_CELL_FORMAT_KIND;
pub const TEXT_CELL_FORMAT_KIND: u32 = 5;

/// Return the exact native Currency marker for the secondary-identifier
/// shape.
#[must_use]
pub const fn explicit_currency_format_flags(has_secondary_number_format: bool) -> u16 {
    if has_secondary_number_format {
        EXPLICIT_CURRENCY_WITH_NUMBER_FORMAT
    } else {
        EXPLICIT_CURRENCY_FORMAT
    }
}

/// Return the exact native Duration marker for the secondary-identifier
/// shape.
#[must_use]
pub const fn explicit_duration_format_flags(has_secondary_number_format: bool) -> u16 {
    if has_secondary_number_format {
        EXPLICIT_DURATION_WITH_NUMBER_FORMAT
    } else {
        EXPLICIT_DURATION_FORMAT
    }
}

const VALUE_FLAGS: u32 = DECIMAL_FLAG
    | NUMBER_FLAG
    | DATE_FLAG
    | STRING_FLAG
    | RICH_TEXT_FLAG
    | FORMULA_FLAG
    | FORMULA_ERROR_FLAG;
const FORMULA_CACHE_FLAGS: u32 =
    DECIMAL_FLAG | NUMBER_FLAG | DATE_FLAG | STRING_FLAG | RICH_TEXT_FLAG;

pub(crate) const FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x0000_0001, 16),
    (0x0000_0002, 8),
    (0x0000_0004, 8),
    (0x0000_0008, 4),
    (0x0000_0010, 4),
    (0x0000_0020, 4),
    (0x0000_0040, 4),
    (0x0000_0080, 4),
    (0x0000_0100, 4),
    (0x0000_0200, 4),
    (0x0000_0400, 4),
    (0x0000_0800, 4),
    (0x0000_1000, 4),
    (0x0000_2000, 4),
    (0x0000_4000, 4),
    (0x0000_8000, 4),
    (0x0001_0000, 4),
    (0x0002_0000, 4),
    (0x0004_0000, 4),
    (0x0008_0000, 4),
    (0x0010_0000, 4),
];
const RESERVED_KNOWN_FIELD_FLAG: u32 = 0x0010_0000;
const FIELD_COUNT: usize = FIELD_LAYOUT.len();

/// Conservative allocation count for one owned [`BncCell::parse`].
///
/// Every present fixed-layout field can allocate one map node and one value
/// buffer; the opaque tail can allocate one additional buffer. Focused
/// package owners use this bound to debit their operation ledger before
/// materializing an owned cell.
pub const MAX_OWNED_BNC_PARSE_ALLOCATIONS: usize = FIELD_COUNT * 2 + 1;

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    InvalidFormat(String),
    ParseError(String),
    /// A bounded encoder's exact output would exceed its caller-selected cap.
    OutputLimitExceeded {
        /// Exact bytes required by the encoded cell.
        observed: usize,
        /// Maximum bytes authorized by the caller.
        maximum: usize,
    },
    /// A bounded encoder could not reserve its exact output allocation.
    Allocation {
        /// Exact bytes requested for the encoded cell.
        requested: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidFormat(message) | Self::ParseError(message) => {
                formatter.write_str(message)
            },
            Self::OutputLimitExceeded { observed, maximum } => write!(
                formatter,
                "Numbers BNC output limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { requested } => {
                write!(
                    formatter,
                    "Could not allocate {requested} Numbers BNC bytes"
                )
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for BNC decoding and mutation.
pub type Result<T> = std::result::Result<T, Error>;

/// Allocation-free exact output requirement for one BNC mutation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RewritePlan {
    output_len: Option<usize>,
}

impl RewritePlan {
    #[must_use]
    pub const fn output_len(self) -> Option<usize> {
        self.output_len
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellDataFormatKind {
    NumberOrPercentage,
    Currency,
    DateTime,
    Duration,
    Checkbox,
    StarRating,
    NumericControlNumberOrPercentage,
    NumericControlCurrency,
    Text,
    PopUpMenu,
}

#[derive(Debug, Clone)]
pub struct BncCell {
    prefix: [u8; BNC_PREFIX_LEN],
    fields: BTreeMap<u32, Vec<u8>>,
    tail: Vec<u8>,
}

/// Allocation-free semantic view over one encoded BNC cell.
pub struct BncCellView<'a> {
    prefix: &'a [u8],
    flags: u32,
    cell_type: u8,
    fields: [Option<&'a [u8]>; FIELD_COUNT],
    cached_scalar: Option<CachedScalar>,
    tail: &'a [u8],
}

#[derive(Clone, Copy)]
struct DecodedScalarFields {
    decimal: Option<FiniteF64>,
    number: Option<FiniteF64>,
    date: Option<FiniteF64>,
}

#[derive(Clone, Copy)]
struct DurationViewTransition<'a> {
    current_identifier: Option<u32>,
    secondary: Option<&'a [u8]>,
}

struct EncodedScalar {
    cell_type: u8,
    flag: u32,
    bytes: [u8; 16],
    length: usize,
}

fn encoded_cache_matches(value: ScalarValue, encoded: &EncodedScalar) -> Result<bool> {
    let decoded = cached_scalar_from(
        encoded.cell_type,
        decode_scalar_fields(|flag| {
            (flag == encoded.flag).then_some(&encoded.bytes[..encoded.length])
        })?,
    );
    Ok(match value {
        ScalarValue::String(_) => encoded.cell_type == CELL_TYPE_TEXT,
        ScalarValue::Number(value) => decoded == Some(CachedScalar::Number(value)),
        ScalarValue::Boolean(value) => decoded == Some(CachedScalar::Boolean(value)),
        ScalarValue::Date(value) => decoded == Some(CachedScalar::Date(value)),
        ScalarValue::Duration(value) => decoded == Some(CachedScalar::Duration(value)),
        ScalarValue::RichText(_) => false,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StoredValue {
    Empty,
    Number,
    Text(u32),
    Formula(u32),
    RichText(u32),
    Date,
    Boolean,
    Duration,
    Error,
    Unsupported(u8),
}

/// Native numeric representation of a BNC cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumericCellType {
    /// The ordinary Numbers numeric cell type.
    Number,
    /// Numbers' alternate numeric cell type used by Currency formats.
    AlternateNumber,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CachedScalar {
    Number(FiniteF64),
    Boolean(bool),
    Date(FiniteF64),
    Duration(FiniteF64),
    Unsupported(u8),
}

/// One scalar value accepted by the bounded raw BNC rewrite primitive.
///
/// `Number` follows Numbers' format-aware behavior: duration-formatted cells
/// convert spreadsheet days to seconds, while currency and date/time formats
/// retain their native numeric cell types.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ScalarValue {
    String(u32),
    RichText(u32),
    Number(FiniteF64),
    Boolean(bool),
    Date(FiniteF64),
    Duration(FiniteF64),
}

/// Result of clearing the value-bearing fields from one raw BNC cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ClearValue {
    /// The cleared bytes are exactly the canonical minimal empty cell, so the
    /// enclosing sparse slot may be deleted.
    Delete,
    /// Non-value prefix metadata, fields, or opaque tail bytes remain.
    Retain(Vec<u8>),
}

impl BncCell {
    /// Parses one Numbers BNC cell while retaining unknown trailing bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the cell is truncated, uses an unsupported
    /// version, contains an unknown field flag, or decodes a non-finite
    /// scalar.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let view = BncCellView::parse(data)?;
        let mut prefix = [0; BNC_PREFIX_LEN];
        prefix.copy_from_slice(&data[..BNC_PREFIX_LEN]);
        let mut fields = BTreeMap::new();
        for ((flag, _size), field_bytes) in FIELD_LAYOUT.iter().zip(view.fields) {
            if let Some(bytes) = field_bytes {
                fields.insert(*flag, bytes.to_vec());
            }
        }
        Ok(Self {
            prefix,
            fields,
            tail: view.tail.to_vec(),
        })
    }

    /// Creates the smallest writable BNC cell.
    #[must_use]
    pub fn minimal() -> Self {
        let mut prefix = [0; BNC_PREFIX_LEN];
        prefix[0] = BNC_VERSION;
        Self {
            prefix,
            fields: BTreeMap::new(),
            tail: Vec::new(),
        }
    }

    #[must_use]
    pub fn stored_value(&self) -> StoredValue {
        stored_value_from(self.prefix[1], |flag| self.u32_field(flag))
    }

    /// Returns the native numeric representation when this cell is numeric.
    #[must_use]
    pub fn numeric_cell_type(&self) -> Option<NumericCellType> {
        match self.prefix[1] {
            CELL_TYPE_NUMBER => Some(NumericCellType::Number),
            CELL_TYPE_ALTERNATE_NUMBER => Some(NumericCellType::AlternateNumber),
            _ => None,
        }
    }

    /// Reports whether this cell has one of the value shapes currently
    /// admitted by the focused Date & Time display-format owner.
    ///
    /// The predicate intentionally distinguishes native type 9
    /// (`rich-text-or-number`) from ordinary numeric type 2 and Currency's
    /// alternate numeric type 10. Empty cells are admitted because they have
    /// no value representation to convert. It is a wire-level admission
    /// helper; the Numbers semantic crate still owns public policy and
    /// transaction errors.
    #[must_use]
    pub fn is_date_time_format_compatible(&self) -> bool {
        self.validate_date_time_value_shape().is_ok()
    }

    /// Reports whether all present BNC format metadata belongs to the
    /// Date/Time family.
    ///
    /// Value-shape compatibility and metadata-family compatibility are kept
    /// separate so callers can diagnose an unformatted value independently
    /// from a cell that carries a generic secondary format reference. A
    /// Date/Time cell may retain ordinary style/comment fields, but it cannot
    /// carry a decimal, Currency, duration, text, control, or other format
    /// identifier alongside its primary Date/Time reference.
    #[must_use]
    pub fn has_only_date_time_format_metadata(&self) -> bool {
        self.fields.keys().all(|field| {
            FORMAT_METADATA_FLAGS & *field == 0
                || matches!(
                    *field,
                    CELL_FORMAT_KIND_FLAG | DATE_TIME_FORMAT_IDENTIFIER_FLAG
                )
        })
    }

    /// Reports whether this cell has one of the value shapes owned by the
    /// focused Duration display-format adapter.
    ///
    /// Duration values use native BNC type 7 and an eight-byte `NUMBER_FLAG`
    /// scalar. Empty cells are admitted because there is no scalar to
    /// convert. A formula may retain its typed Duration cache and the native
    /// formula display/error references.
    #[must_use]
    pub fn is_duration_format_compatible(&self) -> bool {
        self.validate_duration_value_shape().is_ok()
    }

    /// Reports whether all present BNC format metadata belongs to the
    /// Duration family.
    ///
    /// Duration has one family-specific primary identifier and may retain the
    /// shared generic Number identifier as a secondary reference. Tuple
    /// completeness, marker ownership, and non-zero identifiers are checked
    /// by the focused transition rather than this allocation-free field-family
    /// predicate.
    #[must_use]
    pub fn has_only_duration_format_metadata(&self) -> bool {
        self.fields.keys().all(|field| {
            FIELD_LAYOUT.iter().any(|(known, _)| known == field)
                && (FORMAT_METADATA_FLAGS & *field == 0
                    || matches!(
                        *field,
                        CELL_FORMAT_KIND_FLAG
                            | CELL_FORMAT_IDENTIFIER_FLAG
                            | DURATION_FORMAT_IDENTIFIER_FLAG
                    ))
        })
    }

    /// Replaces the cell value with a number.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is non-finite or cannot be encoded as a
    /// Numbers decimal value.
    pub fn set_number(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite numeric value".to_string(),
            ));
        }
        if self.cell_format_kind() == Some(DURATION_CELL_FORMAT_KIND) {
            return self.set_duration(spreadsheet_days_to_seconds(value)?);
        }
        let cell_type = match self.cell_format_kind() {
            Some(CURRENCY_CELL_FORMAT_KIND) => CELL_TYPE_ALTERNATE_NUMBER,
            Some(DATE_TIME_CELL_FORMAT_KIND) => CELL_TYPE_RICH_TEXT_OR_NUMBER,
            _ => CELL_TYPE_NUMBER,
        };
        self.replace_value(cell_type, DECIMAL_FLAG, decimal128_le(value)?.to_vec());
        Ok(())
    }

    /// Replaces the cell value with a plain numeric BNC value.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is non-finite or cannot be encoded as a
    /// Numbers decimal value.
    pub fn set_plain_number(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite numeric value".to_string(),
            ));
        }
        self.replace_value(
            CELL_TYPE_NUMBER,
            DECIMAL_FLAG,
            decimal128_le(value)?.to_vec(),
        );
        Ok(())
    }

    pub fn set_boolean(&mut self, value: bool) {
        self.replace_value(
            CELL_TYPE_BOOLEAN,
            NUMBER_FLAG,
            (if value { 1.0f64 } else { 0.0f64 }).to_le_bytes().to_vec(),
        );
    }

    /// Replaces the cell value with a duration measured in seconds.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is non-finite.
    pub fn set_duration(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite duration".to_string(),
            ));
        }
        self.replace_value(
            CELL_TYPE_DURATION,
            NUMBER_FLAG,
            value.to_le_bytes().to_vec(),
        );
        Ok(())
    }

    /// Replaces only the explicit Duration display metadata.
    ///
    /// Unlike [`Self::set_data_format_identifier`], this focused primitive
    /// never converts the stored scalar, formula cache, native cell type, or
    /// any unrelated field. An explicit Duration cell may retain the shared
    /// generic Number identifier as a secondary format-table reference; that
    /// reference is preserved while changing the primary Duration ID and is
    /// removed when the explicit Duration metadata is cleared. Native Numbers
    /// uses marker `0x0004` for the primary-only tuple and `0x0005` when this
    /// secondary reference is present. Empty cells and native type-7 Duration
    /// cells are supported.
    ///
    /// The operation is fail-closed. Automatic Duration tuples, another
    /// format family, control metadata, reserved fields, malformed fixed
    /// fields, zero IDs, and ambiguous value/cache shapes are rejected before
    /// any mutation. The prefix marker bytes are changed only when the
    /// requested explicit state changes; all other prefix, value, formula,
    /// style, comment, and opaque-tail bytes remain unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when the requested or existing Duration metadata is
    /// not a canonical explicit tuple, the value shape is not writable as a
    /// Duration cell, or a fixed-width field/scalar is malformed.
    pub fn set_duration_format_identifier_preserving_value(
        &mut self,
        identifier: Option<u32>,
    ) -> Result<()> {
        let current_identifier = self.validate_duration_transition(identifier)?;
        if current_identifier == identifier {
            return Ok(());
        }

        let has_secondary_identifier = self.cell_format_kind() == Some(DURATION_CELL_FORMAT_KIND)
            && self.fields.contains_key(&CELL_FORMAT_IDENTIFIER_FLAG);
        // Keep an existing generic Number reference in its BTreeMap node when
        // replacing the primary Duration ID.  Cloning the field here would
        // add a transient allocation to the owned path (and clearing the
        // metadata would allocate only to discard it).  Validation above has
        // already proved that this is a four-byte, canonical secondary field.
        self.fields.retain(|field, _| {
            FORMAT_METADATA_FLAGS & *field == 0
                || (identifier.is_some()
                    && has_secondary_identifier
                    && *field == CELL_FORMAT_IDENTIFIER_FLAG)
        });
        if let Some(identifier) = identifier {
            self.fields.insert(
                CELL_FORMAT_KIND_FLAG,
                DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
            );
            self.fields.insert(
                DURATION_FORMAT_IDENTIFIER_FLAG,
                identifier.to_le_bytes().to_vec(),
            );
        }
        let explicit_flags = identifier.map_or(0, |_| {
            explicit_duration_format_flags(has_secondary_identifier)
        });
        self.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&explicit_flags.to_le_bytes());
        Ok(())
    }

    /// Replaces the cell value with a date/time serial value.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is non-finite.
    pub fn set_date(&mut self, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::ParseError(
                "Numbers cells cannot store a non-finite date".to_string(),
            ));
        }
        self.replace_value(CELL_TYPE_DATE, DATE_FLAG, value.to_le_bytes().to_vec());
        Ok(())
    }

    pub fn set_string(&mut self, identifier: u32) {
        self.replace_value(
            CELL_TYPE_TEXT,
            STRING_FLAG,
            identifier.to_le_bytes().to_vec(),
        );
    }

    pub fn set_rich_text(&mut self, identifier: u32) {
        self.replace_value(
            CELL_TYPE_RICH_TEXT_OR_NUMBER,
            RICH_TEXT_FLAG,
            identifier.to_le_bytes().to_vec(),
        );
    }

    pub fn set_formula_reference(&mut self, identifier: u32) {
        // Formula references coexist with the cached result value and its cell
        // type in app-generated BNC. The caller seeds a numeric cache before
        // attaching the formula when the target cell was empty.
        if self.prefix[1] == CELL_TYPE_EMPTY {
            self.prefix[1] = CELL_TYPE_NUMBER;
            self.fields
                .insert(NUMBER_FLAG, 0.0f64.to_le_bytes().to_vec());
        }
        self.fields
            .insert(FORMULA_FLAG, identifier.to_le_bytes().to_vec());
        self.fields.remove(&FORMULA_ERROR_FLAG);
    }

    /// Reads the cached scalar value retained alongside a formula.
    ///
    /// # Errors
    ///
    /// Returns an error when a numeric field has an invalid byte width, a
    /// decimal128 field cannot be decoded, or a scalar is non-finite.
    pub fn cached_scalar(&self) -> Result<Option<CachedScalar>> {
        let fields = decode_scalar_fields(|flag| self.fields.get(&flag).map(Vec::as_slice))?;
        Ok(cached_scalar_from(self.prefix[1], fields))
    }

    /// Replaces a formula's cached result with a number.
    ///
    /// # Errors
    ///
    /// Returns an error when the cell has no formula or `value` is not a
    /// finite Numbers value.
    pub fn set_formula_cached_number(&mut self, value: f64) -> Result<()> {
        let formula = self.formula_identifier()?;
        let formula_error = self.fields.get(&FORMULA_ERROR_FLAG).cloned();
        self.set_number(value)?;
        self.fields
            .insert(FORMULA_FLAG, formula.to_le_bytes().to_vec());
        if let Some(formula_error) = formula_error {
            self.fields.insert(FORMULA_ERROR_FLAG, formula_error);
        }
        Ok(())
    }

    /// Replaces a formula's cached result with a boolean.
    ///
    /// # Errors
    ///
    /// Returns an error when the cell has no formula.
    pub fn set_formula_cached_boolean(&mut self, value: bool) -> Result<()> {
        let formula = self.formula_identifier()?;
        let formula_error = self.fields.get(&FORMULA_ERROR_FLAG).cloned();
        self.set_boolean(value);
        self.fields
            .insert(FORMULA_FLAG, formula.to_le_bytes().to_vec());
        if let Some(formula_error) = formula_error {
            self.fields.insert(FORMULA_ERROR_FLAG, formula_error);
        }
        Ok(())
    }

    #[must_use]
    pub fn formula_error_identifier(&self) -> Option<u32> {
        self.u32_field(FORMULA_ERROR_FLAG)
    }

    #[must_use]
    pub fn comment_identifier(&self) -> Option<u32> {
        self.u32_field(COMMENT_FLAG)
    }

    /// Returns the raw rich-text key carried by the cell, regardless of its
    /// native cell type. Presence is reported even when the key is zero.
    #[must_use]
    pub fn rich_text_identifier(&self) -> Option<u32> {
        self.u32_field(RICH_TEXT_FLAG)
    }

    #[must_use]
    pub fn style_identifier(&self) -> Option<u32> {
        self.u32_field(STYLE_FLAG)
    }

    #[must_use]
    pub fn text_style_identifier(&self) -> Option<u32> {
        self.u32_field(TEXT_STYLE_FLAG)
    }

    #[must_use]
    pub fn conditional_style_identifier(&self) -> Option<u32> {
        self.u32_field(CONDITIONAL_STYLE_FLAG)
    }

    #[must_use]
    pub fn conditional_style_applied_rule(&self) -> Option<u32> {
        self.u32_field(CONDITIONAL_STYLE_APPLIED_RULE_FLAG)
    }

    #[must_use]
    pub fn explicit_format_flags(&self) -> u16 {
        u16::from_le_bytes([
            self.prefix[EXPLICIT_FORMAT_FLAGS_START],
            self.prefix[EXPLICIT_FORMAT_FLAGS_START + 1],
        ])
    }

    #[must_use]
    pub fn cell_format_kind(&self) -> Option<u32> {
        self.u32_field(CELL_FORMAT_KIND_FLAG)
    }

    #[must_use]
    pub fn control_cell_spec_identifier(&self) -> Option<u32> {
        self.u32_field(CONTROL_CELL_SPEC_FLAG)
    }

    #[must_use]
    pub fn format_identifier(&self) -> Option<u32> {
        match self.cell_format_kind() {
            Some(CURRENCY_CELL_FORMAT_KIND) => self.u32_field(CURRENCY_FORMAT_IDENTIFIER_FLAG),
            Some(DATE_TIME_CELL_FORMAT_KIND) => self.u32_field(DATE_TIME_FORMAT_IDENTIFIER_FLAG),
            Some(DURATION_CELL_FORMAT_KIND) => self.u32_field(DURATION_FORMAT_IDENTIFIER_FLAG),
            Some(TEXT_CELL_FORMAT_KIND) => self.u32_field(TEXT_FORMAT_IDENTIFIER_FLAG),
            Some(CHECKBOX_CELL_FORMAT_KIND) => self.u32_field(CHECKBOX_FORMAT_IDENTIFIER_FLAG),
            _ => self.u32_field(CELL_FORMAT_IDENTIFIER_FLAG),
        }
    }

    #[must_use]
    pub fn secondary_format_identifier(&self) -> Option<u32> {
        match self.cell_format_kind() {
            Some(CURRENCY_CELL_FORMAT_KIND | DURATION_CELL_FORMAT_KIND) => {
                self.u32_field(CELL_FORMAT_IDENTIFIER_FLAG)
            },
            _ => None,
        }
    }

    /// Reports whether every present display-metadata field belongs to the
    /// ordinary decimal shape.
    ///
    /// Number, Percentage, and Scientific cells may carry only the shared
    /// format kind and primary format identifier. This check catches orphan
    /// Currency/date/duration/text/control identifiers that the kind-directed
    /// accessors intentionally do not expose.
    #[must_use]
    pub fn has_only_decimal_format_metadata(&self) -> bool {
        self.fields.keys().all(|field| {
            FORMAT_METADATA_FLAGS & field == 0
                || matches!(*field, CELL_FORMAT_KIND_FLAG | CELL_FORMAT_IDENTIFIER_FLAG)
        })
    }

    /// Applies a Numbers data format and its identifier to the cell.
    ///
    /// # Errors
    ///
    /// Returns an error when an interactive format has no control-cell
    /// identifier, when a non-interactive format has one, or when a text
    /// format would discard a non-text scalar, when Duration is requested for
    /// an incompatible value family or with a zero identifier, or when a
    /// required scalar conversion is not finite. This compatibility entry
    /// point performs the same native scalar conversions as Numbers;
    /// metadata-only application is kept as a separate internal operation so
    /// extraction never infers semantic value from display metadata. The
    /// focused Duration owner preserves a valid existing generic Number
    /// secondary and selects marker `0x0005`. This legacy compatibility entry
    /// point canonicalizes every replacement to a primary-only Duration tuple
    /// (marker `0x0004`), because its host reference bookkeeping owns the old
    /// format edges and cannot safely retain that secondary during a generic
    /// format change.
    pub fn set_data_format_identifier(
        &mut self,
        identifier: u32,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) -> Result<()> {
        self.validate_data_format_request(identifier, kind, control_identifier)?;
        self.convert_scalar_for_data_format(kind)?;
        self.set_data_format_metadata_identifier(identifier, kind, control_identifier)
    }

    /// Replaces only explicit decimal-family display metadata.
    ///
    /// Unlike [`Self::set_data_format_identifier`], this focused primitive
    /// never converts the stored value, its formula cache, or the native cell
    /// type. Passing `None` removes the explicit format while retaining every
    /// non-format field and opaque trailing byte. It exists for graph owners
    /// that have already validated the native format family and need to move
    /// a cell between entries in that same decimal format list.
    ///
    /// # Errors
    ///
    /// Returns an error when an explicit format identifier is zero.
    pub fn set_number_or_percentage_format_identifier_preserving_value(
        &mut self,
        identifier: Option<u32>,
    ) -> Result<()> {
        if identifier.is_some_and(|value| value == 0) {
            return Err(Error::InvalidFormat(
                "decimal-format identifier must be non-zero".to_owned(),
            ));
        }

        self.fields
            .retain(|field, _| FORMAT_METADATA_FLAGS & field == 0);
        let explicit_flags = if let Some(identifier) = identifier {
            self.fields.insert(
                CELL_FORMAT_KIND_FLAG,
                DECIMAL_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
            );
            self.fields.insert(
                CELL_FORMAT_IDENTIFIER_FLAG,
                identifier.to_le_bytes().to_vec(),
            );
            EXPLICIT_DECIMAL_FORMAT
        } else {
            0
        };
        self.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&explicit_flags.to_le_bytes());
        Ok(())
    }

    /// Replaces only the explicit Currency display metadata.
    ///
    /// Unlike [`Self::set_data_format_identifier`], this focused primitive
    /// never converts the stored scalar, its formula cache, or any unrelated
    /// field. A Currency format uses Numbers' alternate-number cell type, so
    /// installing an identifier changes a plain numeric cell to that type and
    /// clearing the identifier changes it back. Passing `None` also removes
    /// the optional secondary format identifier while retaining every
    /// non-format field and opaque trailing byte. When installing a Currency
    /// identifier, an existing secondary `CELL_FORMAT_IDENTIFIER_FLAG` is
    /// retained because native Currency cells may reference both entries in
    /// the format table. It exists for graph owners that have already
    /// validated the native format family and need to move a cell between
    /// entries in that Currency format list.
    ///
    /// # Errors
    ///
    /// Returns an error when an explicit format identifier is zero.
    pub fn set_currency_format_identifier_preserving_value(
        &mut self,
        identifier: Option<u32>,
    ) -> Result<()> {
        if identifier.is_some_and(|value| value == 0) {
            return Err(Error::InvalidFormat(
                "Currency format identifier must be non-zero".to_owned(),
            ));
        }

        let secondary_identifier = (self.cell_format_kind() == Some(CURRENCY_CELL_FORMAT_KIND))
            .then(|| self.fields.get(&CELL_FORMAT_IDENTIFIER_FLAG).cloned())
            .flatten();
        let has_secondary_identifier = secondary_identifier.is_some();
        self.fields
            .retain(|field, _| FORMAT_METADATA_FLAGS & field == 0);
        let explicit_flags = if let Some(identifier) = identifier {
            if self.prefix[1] == CELL_TYPE_NUMBER {
                self.prefix[1] = CELL_TYPE_ALTERNATE_NUMBER;
            }
            self.fields.insert(
                CELL_FORMAT_KIND_FLAG,
                CURRENCY_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
            );
            self.fields.insert(
                CURRENCY_FORMAT_IDENTIFIER_FLAG,
                identifier.to_le_bytes().to_vec(),
            );
            if let Some(secondary_identifier) = secondary_identifier {
                self.fields
                    .insert(CELL_FORMAT_IDENTIFIER_FLAG, secondary_identifier);
            }
            explicit_currency_format_flags(has_secondary_identifier)
        } else {
            if self.prefix[1] == CELL_TYPE_ALTERNATE_NUMBER {
                self.prefix[1] = CELL_TYPE_NUMBER;
            }
            0
        };
        self.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&explicit_flags.to_le_bytes());
        Ok(())
    }

    /// Replaces only explicit Date/Time display metadata.
    ///
    /// Date/Time is a distinct native BNC family.  Unlike the compatibility
    /// [`Self::set_data_format_identifier`] entry point, this operation never
    /// converts the cell type or rewrites a scalar, formula, cache, or any
    /// other field. It accepts the native value shapes owned by the Date/Time
    /// adapter: Empty, a type-5 date, and a type-9 numeric cell (including a
    /// numeric formula cache). Passing `None` removes the explicit Date/Time
    /// marker, kind, and identifier while retaining the complete non-format
    /// representation, including the original cell type and opaque tail.
    ///
    /// The operation is deliberately fail-closed.  An automatic Date/Time
    /// tuple, a different format family, a secondary/control reference, a
    /// reserved known field, an ambiguous value/cache shape, or a malformed
    /// reference is rejected before any mutation.  The format identifier is
    /// also required to be non-zero.
    pub fn set_date_time_format_identifier_preserving_value(
        &mut self,
        identifier: Option<u32>,
    ) -> Result<()> {
        let current_identifier = self.validate_date_time_transition(identifier)?;
        if current_identifier == identifier {
            return Ok(());
        }

        self.fields.remove(&CELL_FORMAT_KIND_FLAG);
        self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
        if let Some(identifier) = identifier {
            self.fields.insert(
                CELL_FORMAT_KIND_FLAG,
                DATE_TIME_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
            );
            self.fields.insert(
                DATE_TIME_FORMAT_IDENTIFIER_FLAG,
                identifier.to_le_bytes().to_vec(),
            );
        }
        let explicit_flags = identifier.map_or(0, |_| EXPLICIT_DATE_TIME_FORMAT);
        self.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&explicit_flags.to_le_bytes());
        Ok(())
    }

    fn validate_date_time_transition(&self, identifier: Option<u32>) -> Result<Option<u32>> {
        if identifier.is_some_and(|identifier| identifier == 0) {
            return Err(Error::InvalidFormat(
                "Date/Time format identifier must be non-zero".to_owned(),
            ));
        }
        if self.prefix[0] != BNC_VERSION {
            return Err(Error::InvalidFormat(
                "Date/Time metadata requires a Numbers BNC v5 cell".to_owned(),
            ));
        }

        for (&flag, bytes) in &self.fields {
            let Some((_, expected_width)) = FIELD_LAYOUT
                .iter()
                .find(|(candidate, _)| *candidate == flag)
            else {
                return Err(Error::InvalidFormat(
                    "Date/Time metadata encountered an unknown BNC field".to_owned(),
                ));
            };
            if bytes.len() != *expected_width {
                return Err(Error::InvalidFormat(
                    "Date/Time metadata encountered a malformed BNC field".to_owned(),
                ));
            }
        }

        if self.fields.contains_key(&RESERVED_KNOWN_FIELD_FLAG) {
            return Err(Error::InvalidFormat(
                "Date/Time metadata cannot rewrite a reserved BNC field".to_owned(),
            ));
        }

        // Date/Time has exactly one primary format reference.  In particular,
        // the shared generic identifier is not a valid secondary reference,
        // and interactive control metadata never accompanies this family.
        if self.fields.keys().any(|flag| {
            matches!(
                *flag,
                CONTROL_CELL_SPEC_FLAG
                    | CELL_FORMAT_IDENTIFIER_FLAG
                    | CURRENCY_FORMAT_IDENTIFIER_FLAG
                    | DURATION_FORMAT_IDENTIFIER_FLAG
                    | TEXT_FORMAT_IDENTIFIER_FLAG
                    | CHECKBOX_FORMAT_IDENTIFIER_FLAG
            )
        }) {
            return Err(Error::InvalidFormat(
                "Date/Time metadata has an incompatible secondary or control reference".to_owned(),
            ));
        }

        let marker = self.explicit_format_flags();
        if marker != 0 && marker != EXPLICIT_DATE_TIME_FORMAT {
            return Err(Error::InvalidFormat(
                "Date/Time metadata has an incompatible explicit-format marker".to_owned(),
            ));
        }
        let kind = self.u32_field(CELL_FORMAT_KIND_FLAG);
        if kind.is_some_and(|kind| kind != DATE_TIME_CELL_FORMAT_KIND) {
            return Err(Error::InvalidFormat(
                "Date/Time metadata has an incompatible cell-format kind".to_owned(),
            ));
        }
        let date_time_identifier = self.u32_field(DATE_TIME_FORMAT_IDENTIFIER_FLAG);
        if date_time_identifier.is_some_and(|identifier| identifier == 0) {
            return Err(Error::InvalidFormat(
                "Date/Time format identifier must be non-zero".to_owned(),
            ));
        }
        let current_identifier = match (marker, kind, date_time_identifier) {
            (0, None, None) => None,
            (EXPLICIT_DATE_TIME_FORMAT, Some(DATE_TIME_CELL_FORMAT_KIND), Some(identifier)) => {
                Some(identifier)
            },
            // Marker zero is used by Numbers for automatic metadata.  The
            // focused Date/Time owner has no proof that such a tuple is
            // explicit, so it must not silently adopt or clear it.
            (0, Some(DATE_TIME_CELL_FORMAT_KIND), Some(_)) => {
                return Err(Error::InvalidFormat(
                    "automatic Date/Time metadata is not owned by this transition".to_owned(),
                ));
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "Date/Time metadata has an incomplete explicit tuple".to_owned(),
                ));
            },
        };

        self.validate_date_time_value_shape()?;
        Ok(current_identifier)
    }

    fn validate_date_time_value_shape(&self) -> Result<()> {
        let value_flags = self
            .fields
            .keys()
            .fold(0, |flags, field| flags | (*field & VALUE_FLAGS));
        match self.prefix[1] {
            CELL_TYPE_EMPTY => {
                if value_flags != 0 || self.cached_scalar()?.is_some() {
                    return Err(Error::InvalidFormat(
                        "Date/Time empty cell has an incompatible value shape".to_owned(),
                    ));
                }
            },
            CELL_TYPE_DATE => {
                if value_flags != DATE_FLAG
                    || !matches!(self.cached_scalar()?, Some(CachedScalar::Date(_)))
                {
                    return Err(Error::InvalidFormat(
                        "Date/Time type-5 cell has an incompatible value shape".to_owned(),
                    ));
                }
            },
            CELL_TYPE_RICH_TEXT_OR_NUMBER => {
                let numeric_flags = value_flags & (DECIMAL_FLAG | NUMBER_FLAG);
                if numeric_flags != DECIMAL_FLAG && numeric_flags != NUMBER_FLAG
                    || value_flags & (DATE_FLAG | RICH_TEXT_FLAG) != 0
                    || self
                        .cached_scalar()?
                        .is_none_or(|scalar| !matches!(scalar, CachedScalar::Number(_)))
                {
                    return Err(Error::InvalidFormat(
                        "Date/Time type-9 cell has an incompatible or ambiguous value shape"
                            .to_owned(),
                    ));
                }

                let formula_identifier = self.u32_field(FORMULA_FLAG);
                if formula_identifier.is_some_and(|identifier| identifier == 0) {
                    return Err(Error::InvalidFormat(
                        "Date/Time formula identifier must be non-zero".to_owned(),
                    ));
                }
                if formula_identifier.is_none()
                    && (self.fields.contains_key(&STRING_FLAG)
                        || self.fields.contains_key(&FORMULA_ERROR_FLAG))
                {
                    return Err(Error::InvalidFormat(
                        "Date/Time type-9 cache has formula-only fields without a formula"
                            .to_owned(),
                    ));
                }
                for flag in [STRING_FLAG, FORMULA_ERROR_FLAG] {
                    if self
                        .u32_field(flag)
                        .is_some_and(|identifier| identifier == 0)
                    {
                        return Err(Error::InvalidFormat(
                            "Date/Time formula cache reference must be non-zero".to_owned(),
                        ));
                    }
                }
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "Date/Time transition requires a native Empty cell, type-5 date, or type-9 numeric cell"
                        .to_owned(),
                ));
            },
        }
        Ok(())
    }

    fn validate_duration_transition(&self, identifier: Option<u32>) -> Result<Option<u32>> {
        if self.prefix[0] != BNC_VERSION {
            return Err(Error::InvalidFormat(
                "Duration metadata requires a Numbers BNC v5 cell".to_owned(),
            ));
        }
        validate_duration_field_storage(&self.fields)?;
        if self.fields.contains_key(&RESERVED_KNOWN_FIELD_FLAG) {
            return Err(Error::InvalidFormat(
                "Duration metadata cannot rewrite a reserved BNC field".to_owned(),
            ));
        }
        if self.fields.keys().any(|field| {
            FORMAT_METADATA_FLAGS & *field != 0
                && !matches!(
                    *field,
                    CELL_FORMAT_KIND_FLAG
                        | CELL_FORMAT_IDENTIFIER_FLAG
                        | DURATION_FORMAT_IDENTIFIER_FLAG
                )
        }) {
            return Err(Error::InvalidFormat(
                "Duration metadata has an incompatible secondary or control reference".to_owned(),
            ));
        }

        let current_identifier = validate_duration_metadata_tuple(
            self.explicit_format_flags(),
            self.u32_field(CELL_FORMAT_KIND_FLAG),
            self.u32_field(DURATION_FORMAT_IDENTIFIER_FLAG),
            self.u32_field(CELL_FORMAT_IDENTIFIER_FLAG),
            identifier,
        )?;
        self.validate_duration_value_shape()?;
        Ok(current_identifier)
    }

    fn validate_duration_value_shape(&self) -> Result<()> {
        validate_duration_field_storage(&self.fields)?;
        let value_flags = self
            .fields
            .keys()
            .fold(0, |flags, field| flags | (*field & VALUE_FLAGS));
        validate_duration_value_shape_parts(
            self.prefix[1],
            value_flags,
            self.cached_scalar()?,
            self.fields.contains_key(&FORMULA_FLAG),
            self.u32_field(FORMULA_FLAG),
            self.fields.contains_key(&STRING_FLAG),
            self.u32_field(STRING_FLAG),
            self.fields.contains_key(&FORMULA_ERROR_FLAG),
            self.u32_field(FORMULA_ERROR_FLAG),
        )
    }

    /// Validate the value families accepted by the generic Duration format
    /// compatibility path. A generic numeric cell is interpreted as
    /// spreadsheet days and converted to native type-7 seconds; an existing
    /// native Duration is already in seconds and is retained as-is. Empty
    /// cells are metadata-only and remain empty. Other value families are
    /// rejected before conversion so attaching Duration metadata cannot
    /// silently relabel text, dates, booleans, errors, or rich text.
    fn validate_duration_data_format_value(&self) -> Result<()> {
        validate_duration_field_storage(&self.fields)?;
        let value_flags = self
            .fields
            .keys()
            .fold(0, |flags, field| flags | (*field & VALUE_FLAGS));
        match self.prefix[1] {
            CELL_TYPE_EMPTY => {
                if value_flags != 0 || self.cached_scalar()?.is_some() {
                    return Err(Error::InvalidFormat(
                        "Duration empty cell has an incompatible value shape".to_owned(),
                    ));
                }
            },
            CELL_TYPE_DURATION => self.validate_duration_value_shape()?,
            CELL_TYPE_NUMBER | CELL_TYPE_ALTERNATE_NUMBER | CELL_TYPE_RICH_TEXT_OR_NUMBER => {
                let numeric_flags = value_flags & (DECIMAL_FLAG | NUMBER_FLAG);
                let allowed_flags =
                    DECIMAL_FLAG | NUMBER_FLAG | FORMULA_FLAG | STRING_FLAG | FORMULA_ERROR_FLAG;
                if value_flags & !allowed_flags != 0
                    || (numeric_flags != DECIMAL_FLAG && numeric_flags != NUMBER_FLAG)
                    || !matches!(self.cached_scalar()?, Some(CachedScalar::Number(_)))
                {
                    return Err(Error::InvalidFormat(
                        "Duration format requires an unambiguous numeric value".to_owned(),
                    ));
                }

                let formula_identifier = self.u32_field(FORMULA_FLAG);
                if formula_identifier.is_some_and(|identifier| identifier == 0) {
                    return Err(Error::InvalidFormat(
                        "Duration formula identifier must be non-zero".to_owned(),
                    ));
                }
                if formula_identifier.is_none()
                    && (self.fields.contains_key(&STRING_FLAG)
                        || self.fields.contains_key(&FORMULA_ERROR_FLAG))
                {
                    return Err(Error::InvalidFormat(
                        "Duration numeric cache has formula-only fields without a formula"
                            .to_owned(),
                    ));
                }
                for flag in [STRING_FLAG, FORMULA_ERROR_FLAG] {
                    if self
                        .u32_field(flag)
                        .is_some_and(|identifier| identifier == 0)
                    {
                        return Err(Error::InvalidFormat(
                            "Duration formula cache reference must be non-zero".to_owned(),
                        ));
                    }
                }
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "Duration format requires an Empty, numeric, or native Duration cell"
                        .to_owned(),
                ));
            },
        }
        Ok(())
    }

    fn set_data_format_metadata_identifier(
        &mut self,
        identifier: u32,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) -> Result<()> {
        self.validate_data_format_request(identifier, kind, control_identifier)?;
        self.apply_data_format_identifier(identifier, kind, control_identifier);
        Ok(())
    }

    fn validate_data_format_request(
        &self,
        identifier: u32,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) -> Result<()> {
        if matches!(kind, CellDataFormatKind::Duration) && identifier == 0 {
            return Err(Error::InvalidFormat(
                "Duration format identifier must be non-zero".to_owned(),
            ));
        }
        match (kind, control_identifier) {
            (
                CellDataFormatKind::Checkbox
                | CellDataFormatKind::StarRating
                | CellDataFormatKind::NumericControlNumberOrPercentage
                | CellDataFormatKind::NumericControlCurrency
                | CellDataFormatKind::PopUpMenu,
                Some(_),
            )
            | (
                CellDataFormatKind::NumberOrPercentage
                | CellDataFormatKind::Currency
                | CellDataFormatKind::DateTime
                | CellDataFormatKind::Duration
                | CellDataFormatKind::Text,
                None,
            ) => {},
            (
                CellDataFormatKind::Checkbox
                | CellDataFormatKind::StarRating
                | CellDataFormatKind::NumericControlNumberOrPercentage
                | CellDataFormatKind::NumericControlCurrency
                | CellDataFormatKind::PopUpMenu,
                None,
            ) => {
                return Err(Error::InvalidFormat(
                    "Interactive format requires a control-cell-spec identifier".to_owned(),
                ));
            },
            (_, Some(_)) => {
                return Err(Error::InvalidFormat(
                    "Non-interactive format cannot use a control-cell-spec identifier".to_owned(),
                ));
            },
        }
        if matches!(
            kind,
            CellDataFormatKind::Text | CellDataFormatKind::PopUpMenu
        ) && !matches!(
            self.stored_value(),
            StoredValue::Empty | StoredValue::Text(_)
        ) {
            return Err(Error::InvalidFormat(
                "Text-based format can only be applied safely to an empty or text cell".to_owned(),
            ));
        }
        if matches!(kind, CellDataFormatKind::Duration) {
            self.validate_duration_data_format_value()?;
        }
        Ok(())
    }

    fn apply_data_format_identifier(
        &mut self,
        identifier: u32,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) {
        if !matches!(
            kind,
            CellDataFormatKind::Checkbox
                | CellDataFormatKind::StarRating
                | CellDataFormatKind::NumericControlNumberOrPercentage
                | CellDataFormatKind::NumericControlCurrency
                | CellDataFormatKind::PopUpMenu
        ) {
            self.fields.remove(&CONTROL_CELL_SPEC_FLAG);
        }
        let (explicit_flags, format_kind) = match kind {
            CellDataFormatKind::NumberOrPercentage
            | CellDataFormatKind::NumericControlNumberOrPercentage => {
                if self.prefix[1] == CELL_TYPE_ALTERNATE_NUMBER {
                    self.prefix[1] = CELL_TYPE_NUMBER;
                }
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    CELL_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_DECIMAL_FORMAT, DECIMAL_CELL_FORMAT_KIND)
            },
            CellDataFormatKind::Currency | CellDataFormatKind::NumericControlCurrency => {
                if self.prefix[1] == CELL_TYPE_NUMBER {
                    self.prefix[1] = CELL_TYPE_ALTERNATE_NUMBER;
                }
                self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    CURRENCY_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (
                    explicit_currency_format_flags(false),
                    CURRENCY_CELL_FORMAT_KIND,
                )
            },
            CellDataFormatKind::DateTime => {
                if matches!(
                    self.prefix[1],
                    CELL_TYPE_NUMBER | CELL_TYPE_ALTERNATE_NUMBER
                ) {
                    self.prefix[1] = CELL_TYPE_RICH_TEXT_OR_NUMBER;
                }
                self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    DATE_TIME_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_DATE_TIME_FORMAT, DATE_TIME_CELL_FORMAT_KIND)
            },
            CellDataFormatKind::Duration => {
                // The compatibility host accounts for old generic references
                // outside this cell mutation. Canonicalize this path to the
                // primary-only marker; the focused Duration owner is the
                // graph-aware path that may retain marker-5's secondary.
                self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    DURATION_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_DURATION_FORMAT, DURATION_CELL_FORMAT_KIND)
            },
            CellDataFormatKind::Checkbox => {
                self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    CHECKBOX_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_CHECKBOX_FORMAT, CHECKBOX_CELL_FORMAT_KIND)
            },
            CellDataFormatKind::StarRating => {
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    CELL_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_DECIMAL_FORMAT, STAR_RATING_CELL_FORMAT_KIND)
            },
            CellDataFormatKind::PopUpMenu => {
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    TEXT_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_TEXT_FORMAT, TEXT_CELL_FORMAT_KIND)
            },
            CellDataFormatKind::Text => {
                self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
                self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
                self.fields.insert(
                    TEXT_FORMAT_IDENTIFIER_FLAG,
                    identifier.to_le_bytes().to_vec(),
                );
                (EXPLICIT_TEXT_FORMAT, TEXT_CELL_FORMAT_KIND)
            },
        };
        if let Some(control_identifier_value) = control_identifier {
            self.fields.insert(
                CONTROL_CELL_SPEC_FLAG,
                control_identifier_value.to_le_bytes().to_vec(),
            );
        }
        self.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&explicit_flags.to_le_bytes());
        self.fields
            .insert(CELL_FORMAT_KIND_FLAG, format_kind.to_le_bytes().to_vec());
    }

    pub fn clear_explicit_format(&mut self) {
        self.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END].fill(0);
        self.fields.remove(&CELL_FORMAT_KIND_FLAG);
        self.fields.remove(&CELL_FORMAT_IDENTIFIER_FLAG);
        self.fields.remove(&CURRENCY_FORMAT_IDENTIFIER_FLAG);
        self.fields.remove(&DATE_TIME_FORMAT_IDENTIFIER_FLAG);
        self.fields.remove(&DURATION_FORMAT_IDENTIFIER_FLAG);
        self.fields.remove(&TEXT_FORMAT_IDENTIFIER_FLAG);
        self.fields.remove(&CHECKBOX_FORMAT_IDENTIFIER_FLAG);
        self.fields.remove(&CONTROL_CELL_SPEC_FLAG);
        let is_plain_numeric_rich_text_cell = self.prefix[1] == CELL_TYPE_RICH_TEXT_OR_NUMBER
            && !self.fields.contains_key(&RICH_TEXT_FLAG)
            && (self.fields.contains_key(&DECIMAL_FLAG) || self.fields.contains_key(&NUMBER_FLAG));
        if self.prefix[1] == CELL_TYPE_ALTERNATE_NUMBER || is_plain_numeric_rich_text_cell {
            self.prefix[1] = CELL_TYPE_NUMBER;
        }
    }

    fn convert_scalar_for_data_format(&mut self, kind: CellDataFormatKind) -> Result<()> {
        // Scalar replacement normally clears all value-bearing fields,
        // including formula references and FormulaError metadata. The
        // generic format path is different: it changes only the native cache
        // representation while retaining the formula graph and its error
        // display reference byte-for-byte.
        let formula = self.fields.get(&FORMULA_FLAG).cloned();
        let formula_text = formula
            .as_ref()
            .and_then(|_| self.fields.get(&STRING_FLAG).cloned());
        let formula_error = formula
            .as_ref()
            .and_then(|_| self.fields.get(&FORMULA_ERROR_FLAG).cloned());
        match (kind, self.cached_scalar()?) {
            (
                CellDataFormatKind::Checkbox,
                Some(
                    CachedScalar::Number(value)
                    | CachedScalar::Date(value)
                    | CachedScalar::Duration(value),
                ),
            ) => {
                self.set_boolean(value.get() != 0.0);
            },
            (CellDataFormatKind::Checkbox, None) => self.set_boolean(false),
            (CellDataFormatKind::StarRating, None) => self.set_number(0.0)?,
            (
                CellDataFormatKind::NumericControlNumberOrPercentage
                | CellDataFormatKind::NumericControlCurrency,
                Some(CachedScalar::Boolean(value)),
            ) => self.replace_value(
                CELL_TYPE_NUMBER,
                DECIMAL_FLAG,
                decimal128_le(if value { 1.0 } else { 0.0 })?.to_vec(),
            ),
            (
                CellDataFormatKind::NumericControlNumberOrPercentage
                | CellDataFormatKind::NumericControlCurrency,
                Some(CachedScalar::Date(value)),
            ) => self.replace_value(
                CELL_TYPE_NUMBER,
                DECIMAL_FLAG,
                decimal128_le(value.get())?.to_vec(),
            ),
            (CellDataFormatKind::Duration, Some(CachedScalar::Number(days))) => {
                self.set_duration(spreadsheet_days_to_seconds(days.get())?)?;
            },
            (
                CellDataFormatKind::NumberOrPercentage
                | CellDataFormatKind::Currency
                | CellDataFormatKind::DateTime
                | CellDataFormatKind::NumericControlNumberOrPercentage
                | CellDataFormatKind::NumericControlCurrency,
                Some(CachedScalar::Duration(seconds)),
            ) => {
                self.replace_value(
                    CELL_TYPE_NUMBER,
                    DECIMAL_FLAG,
                    decimal128_le(seconds.get() / SECONDS_PER_DAY)?.to_vec(),
                );
            },
            _ => return Ok(()),
        }
        if let Some(formula) = formula {
            self.fields.insert(FORMULA_FLAG, formula);
        }
        if let Some(formula_text) = formula_text {
            self.fields.insert(STRING_FLAG, formula_text);
        }
        if let Some(formula_error) = formula_error {
            self.fields.insert(FORMULA_ERROR_FLAG, formula_error);
        }
        Ok(())
    }

    pub fn set_style_identifier(&mut self, identifier: Option<u32>) {
        if let Some(identifier_value) = identifier {
            self.fields
                .insert(STYLE_FLAG, identifier_value.to_le_bytes().to_vec());
        } else {
            self.fields.remove(&STYLE_FLAG);
        }
    }

    pub fn set_text_style_identifier(&mut self, identifier: Option<u32>) {
        if let Some(identifier_value) = identifier {
            self.fields
                .insert(TEXT_STYLE_FLAG, identifier_value.to_le_bytes().to_vec());
        } else {
            self.fields.remove(&TEXT_STYLE_FLAG);
        }
    }

    pub fn set_comment_identifier(&mut self, identifier: Option<u32>) {
        if let Some(identifier_value) = identifier {
            self.fields
                .insert(COMMENT_FLAG, identifier_value.to_le_bytes().to_vec());
        } else {
            self.fields.remove(&COMMENT_FLAG);
        }
    }

    pub fn set_conditional_style(&mut self, identifier: Option<u32>, applied_rule: Option<u32>) {
        if let Some(identifier_value) = identifier {
            self.fields.insert(
                CONDITIONAL_STYLE_FLAG,
                identifier_value.to_le_bytes().to_vec(),
            );
        } else {
            self.fields.remove(&CONDITIONAL_STYLE_FLAG);
        }
        if let Some(applied_rule_value) = applied_rule {
            self.fields.insert(
                CONDITIONAL_STYLE_APPLIED_RULE_FLAG,
                applied_rule_value.to_le_bytes().to_vec(),
            );
        } else {
            self.fields.remove(&CONDITIONAL_STYLE_APPLIED_RULE_FLAG);
        }
    }

    pub fn clear_value_preserving_metadata(&mut self) {
        self.prefix[1] = CELL_TYPE_EMPTY;
        self.fields.retain(|field, _| VALUE_FLAGS & field == 0);
    }

    pub fn encode(&self) -> Vec<u8> {
        let flags = self.fields.keys().fold(0u32, |mask, flag| mask | flag);
        let field_len = self.fields.values().map(Vec::len).sum::<usize>();
        let mut output = Vec::with_capacity(BNC_HEADER_LEN + field_len + self.tail.len());
        output.extend_from_slice(&self.prefix);
        output.extend_from_slice(&flags.to_le_bytes());
        for (flag, _) in FIELD_LAYOUT {
            if let Some(value) = self.fields.get(flag) {
                output.extend_from_slice(value);
            }
        }
        output.extend_from_slice(&self.tail);
        output
    }

    /// Encode this owned cell with an exact, fallible output allocation.
    ///
    /// The complete encoded length is checked before allocation. This method
    /// retains the byte ordering and output of [`Self::encode`] while making
    /// output-limit and allocation failures explicit.
    ///
    /// # Errors
    ///
    /// Returns [`Error::OutputLimitExceeded`] when the exact output exceeds
    /// `max_output_bytes`, [`Error::Allocation`] when its allocation fails, or
    /// [`Error::ParseError`] if stored field lengths overflow `usize`.
    pub fn try_encode_with_limit(&self, max_output_bytes: usize) -> Result<Vec<u8>> {
        let mut output_len = BNC_HEADER_LEN;
        for value in self.fields.values() {
            output_len = output_len.checked_add(value.len()).ok_or_else(|| {
                Error::ParseError("Numbers BNC encoded length overflow".to_owned())
            })?;
        }
        output_len = output_len
            .checked_add(self.tail.len())
            .ok_or_else(|| Error::ParseError("Numbers BNC encoded length overflow".to_owned()))?;
        check_output_limit(output_len, max_output_bytes)?;

        let flags = self.fields.keys().fold(0u32, |mask, flag| mask | flag);
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_error| Error::Allocation {
                requested: output_len,
            })?;
        if output.capacity() != output_len {
            return Err(Error::Allocation {
                requested: output_len,
            });
        }
        output.extend_from_slice(&self.prefix);
        output.extend_from_slice(&flags.to_le_bytes());
        for (flag, _) in FIELD_LAYOUT {
            if let Some(value) = self.fields.get(flag) {
                output.extend_from_slice(value);
            }
        }
        output.extend_from_slice(&self.tail);
        if output.len() != output_len {
            return Err(Error::ParseError(
                "Numbers BNC encoded length changed during publication".to_owned(),
            ));
        }
        Ok(output)
    }

    fn replace_value(&mut self, cell_type: u8, flag: u32, value: Vec<u8>) {
        self.prefix[1] = cell_type;
        self.fields.retain(|field, _| VALUE_FLAGS & field == 0);
        self.fields.insert(flag, value);
    }

    fn u32_field(&self, flag: u32) -> Option<u32> {
        let bytes: [u8; 4] = self.fields.get(&flag)?.as_slice().try_into().ok()?;
        Some(u32::from_le_bytes(bytes))
    }

    fn formula_identifier(&self) -> Result<u32> {
        self.u32_field(FORMULA_FLAG).ok_or_else(|| {
            Error::InvalidFormat(
                "Numbers formula cache update targeted a cell without a formula".to_owned(),
            )
        })
    }
}

impl<'a> BncCellView<'a> {
    fn selected_fields_length(
        &self,
        removed_flags: u32,
        replacements: &[(u32, usize)],
    ) -> Result<usize> {
        let retained_flags = self.flags & !removed_flags;
        let mut output_flags = retained_flags;
        for (flag, length) in replacements {
            if output_flags & flag != 0
                || replacements
                    .iter()
                    .filter(|(candidate, _)| candidate == flag)
                    .count()
                    != 1
                || FIELD_LAYOUT
                    .iter()
                    .find(|(candidate, _)| candidate == flag)
                    .is_none_or(|(_, size)| size != length)
            {
                return Err(Error::ParseError(
                    "Numbers BNC replacement fields are invalid".to_owned(),
                ));
            }
            output_flags |= flag;
        }
        let fields = FIELD_LAYOUT
            .iter()
            .try_fold(0usize, |total, (flag, size)| {
                if output_flags & flag == 0 {
                    Ok(total)
                } else {
                    total.checked_add(*size).ok_or_else(|| {
                        Error::ParseError("Numbers BNC encoded length overflow".to_owned())
                    })
                }
            })?;
        BNC_HEADER_LEN
            .checked_add(fields)
            .and_then(|v| v.checked_add(self.tail.len()))
            .ok_or_else(|| Error::ParseError("Numbers BNC encoded length overflow".to_owned()))
    }

    /// Plan one scalar rewrite without allocating output.
    pub fn plan_scalar_rewrite(&self, value: ScalarValue) -> Result<RewritePlan> {
        let encoded = self.encode_scalar(value)?;
        Ok(RewritePlan {
            output_len: Some(
                self.selected_fields_length(VALUE_FLAGS, &[(encoded.flag, encoded.length)])?,
            ),
        })
    }

    /// Plan one formula rewrite without allocating output.
    pub fn plan_formula_rewrite(
        &self,
        identifier: u32,
        cache: Option<ScalarValue>,
    ) -> Result<RewritePlan> {
        if identifier == 0 {
            return Err(Error::InvalidFormat(
                "Numbers formula key is invalid".to_owned(),
            ));
        }
        let mut replacements = [(FORMULA_FLAG, 4usize), (0, 0)];
        let count = if let Some(cache) = cache {
            let encoded = self.encode_scalar(cache)?;
            replacements[1] = (encoded.flag, encoded.length);
            2
        } else {
            1
        };
        Ok(RewritePlan {
            output_len: Some(self.selected_fields_length(VALUE_FLAGS, &replacements[..count])?),
        })
    }

    /// Plan a supported formula cache-only rewrite without allocating output.
    pub fn plan_formula_cache_rewrite(&self, cache: CachedScalar) -> Result<RewritePlan> {
        if !matches!(self.stored_value(), StoredValue::Formula(_)) {
            return Err(Error::InvalidFormat(
                "Numbers formula cache update targeted a cell without a formula".to_owned(),
            ));
        }
        let scalar = match cache {
            CachedScalar::Number(value) => ScalarValue::Number(value),
            CachedScalar::Boolean(value) => ScalarValue::Boolean(value),
            CachedScalar::Date(_) | CachedScalar::Duration(_) | CachedScalar::Unsupported(_) => {
                return Err(Error::InvalidFormat(
                    "Numbers formula cache update supports only number and Boolean values"
                        .to_owned(),
                ));
            },
        };
        let encoded = self.encode_scalar(scalar)?;
        Ok(RewritePlan {
            output_len: Some(
                self.selected_fields_length(
                    FORMULA_CACHE_FLAGS,
                    &[(encoded.flag, encoded.length)],
                )?,
            ),
        })
    }

    /// Plan a clear without allocating output.
    pub fn plan_clear_value(&self, retain_empty: bool) -> Result<RewritePlan> {
        let retained_flags = self.flags & !VALUE_FLAGS;
        let minimal = retained_flags == 0
            && self.tail.is_empty()
            && self.prefix[0] == BNC_VERSION
            && self.prefix[2..].iter().all(|byte| *byte == 0);
        if minimal && !retain_empty {
            return Ok(RewritePlan { output_len: None });
        }
        if minimal {
            return Ok(RewritePlan {
                output_len: Some(BNC_HEADER_LEN),
            });
        }
        Ok(RewritePlan {
            output_len: Some(self.selected_fields_length(VALUE_FLAGS, &[])?),
        })
    }
    /// Parses the value-bearing portion of a BNC cell without allocating.
    ///
    /// # Errors
    ///
    /// Returns an error when the cell is truncated, uses an unsupported
    /// version, contains an unknown field flag, or decodes a non-finite
    /// scalar.
    pub fn parse(data: &'a [u8]) -> Result<Self> {
        if data.len() < BNC_HEADER_LEN {
            return Err(Error::ParseError(
                "Truncated Numbers BNC cell header".to_string(),
            ));
        }
        if data[0] != BNC_VERSION {
            return Err(Error::ParseError(format!(
                "Numbers cell storage version {} is not writable BNC v5",
                data[0]
            )));
        }

        let mut flag_bytes = [0; 4];
        flag_bytes.copy_from_slice(&data[BNC_PREFIX_LEN..BNC_HEADER_LEN]);
        let flags = u32::from_le_bytes(flag_bytes);
        let known_flags = FIELD_LAYOUT.iter().fold(0, |mask, (flag, _)| mask | flag);
        if flags & !known_flags != 0 {
            return Err(Error::ParseError(format!(
                "Numbers BNC cell uses unknown flags 0x{:08x}",
                flags & !known_flags
            )));
        }

        let mut cursor = BNC_HEADER_LEN;
        let mut fields = [None; FIELD_COUNT];
        for (index, &(flag, size)) in FIELD_LAYOUT.iter().enumerate() {
            if flags & flag == 0 {
                continue;
            }
            let end = cursor.checked_add(size).ok_or_else(|| {
                Error::ParseError("Numbers BNC field offset overflow".to_string())
            })?;
            fields[index] = Some(data.get(cursor..end).ok_or_else(|| {
                Error::ParseError(format!("Truncated Numbers BNC field 0x{flag:08x}"))
            })?);
            cursor = end;
        }

        let decoded_scalar_fields = decode_scalar_fields(|flag| field_from_layout(&fields, flag))?;
        Ok(Self {
            prefix: &data[..BNC_PREFIX_LEN],
            flags,
            cell_type: data[1],
            fields,
            cached_scalar: cached_scalar_from(data[1], decoded_scalar_fields),
            tail: &data[cursor..],
        })
    }

    /// Returns the typed value reference retained by this cell.
    #[must_use]
    pub fn stored_value(&self) -> StoredValue {
        stored_value_from(self.cell_type, |flag| self.u32_field(flag))
    }

    /// Returns the validated, allocation-free scalar cache when present.
    #[must_use]
    pub fn cached_scalar(&self) -> Option<CachedScalar> {
        self.cached_scalar
    }

    /// Returns the native numeric representation when this cell is numeric.
    ///
    /// The alternate-number representation is used by native Currency cells;
    /// it is still exposed as [`StoredValue::Number`] by the semantic value
    /// classifier.  Keeping this distinction available on the borrowed view
    /// lets a package owner validate native format/storage coherence without
    /// materializing an owned cell.
    #[must_use]
    pub fn numeric_cell_type(&self) -> Option<NumericCellType> {
        match self.cell_type {
            CELL_TYPE_NUMBER => Some(NumericCellType::Number),
            CELL_TYPE_ALTERNATE_NUMBER => Some(NumericCellType::AlternateNumber),
            _ => None,
        }
    }

    /// Return the interned string key used as a formula display cache.
    #[must_use]
    pub fn formula_text_key(&self) -> Option<u32> {
        matches!(self.stored_value(), StoredValue::Formula(_))
            .then(|| self.u32_field(STRING_FLAG))
            .flatten()
    }

    /// Returns the native cell-style table identifier, when present.
    ///
    /// The identifier is borrowed metadata: this view never resolves or
    /// allocates for the referenced style object.  A package owner must prove
    /// the referenced style table separately before treating it as admissible
    /// graph state.
    #[must_use]
    pub fn style_identifier(&self) -> Option<u32> {
        self.u32_field(STYLE_FLAG)
    }

    /// Returns the native text-style table identifier, when present.
    #[must_use]
    pub fn text_style_identifier(&self) -> Option<u32> {
        self.u32_field(TEXT_STYLE_FLAG)
    }

    /// Returns the conditional-style table identifier, when present.
    ///
    /// Conditional styles are row-affine dependencies for physical sorting;
    /// the focused Keynote owner should reject them rather than merely move
    /// this identifier with the cell.
    #[must_use]
    pub fn conditional_style_identifier(&self) -> Option<u32> {
        self.u32_field(CONDITIONAL_STYLE_FLAG)
    }

    /// Returns the conditional-style applied-rule identifier, when present.
    #[must_use]
    pub fn conditional_style_applied_rule(&self) -> Option<u32> {
        self.u32_field(CONDITIONAL_STYLE_APPLIED_RULE_FLAG)
    }

    /// Returns the explicit display-format marker stored in the BNC prefix.
    #[must_use]
    pub const fn explicit_format_flags(&self) -> u16 {
        u16::from_le_bytes([
            self.prefix[EXPLICIT_FORMAT_FLAGS_START],
            self.prefix[EXPLICIT_FORMAT_FLAGS_START + 1],
        ])
    }

    /// Returns the native cell-format family marker, when present.
    #[must_use]
    pub fn cell_format_kind(&self) -> Option<u32> {
        self.u32_field(CELL_FORMAT_KIND_FLAG)
    }

    /// Returns the interactive control-cell identifier, when present.
    ///
    /// A non-empty control identifier denotes a row-carried dependency on a
    /// control registry.  Physical row sorting should reject that graph until
    /// the corresponding registry/refcount movement is proven.
    #[must_use]
    pub fn control_cell_spec_identifier(&self) -> Option<u32> {
        self.u32_field(CONTROL_CELL_SPEC_FLAG)
    }

    /// Returns the format-table identifier selected by the native format
    /// family, when present.
    #[must_use]
    pub fn format_identifier(&self) -> Option<u32> {
        match self.cell_format_kind() {
            Some(CURRENCY_CELL_FORMAT_KIND) => self.u32_field(CURRENCY_FORMAT_IDENTIFIER_FLAG),
            Some(DATE_TIME_CELL_FORMAT_KIND) => self.u32_field(DATE_TIME_FORMAT_IDENTIFIER_FLAG),
            Some(DURATION_CELL_FORMAT_KIND) => self.u32_field(DURATION_FORMAT_IDENTIFIER_FLAG),
            Some(TEXT_CELL_FORMAT_KIND) => self.u32_field(TEXT_FORMAT_IDENTIFIER_FLAG),
            Some(CHECKBOX_CELL_FORMAT_KIND) => self.u32_field(CHECKBOX_FORMAT_IDENTIFIER_FLAG),
            _ => self.u32_field(CELL_FORMAT_IDENTIFIER_FLAG),
        }
    }

    /// Returns a secondary generic format identifier used by native Currency
    /// and Duration cells, when present.
    #[must_use]
    pub fn secondary_format_identifier(&self) -> Option<u32> {
        match self.cell_format_kind() {
            Some(CURRENCY_CELL_FORMAT_KIND | DURATION_CELL_FORMAT_KIND) => {
                self.u32_field(CELL_FORMAT_IDENTIFIER_FLAG)
            },
            _ => None,
        }
    }

    /// Reports whether every present display-metadata field belongs to the
    /// ordinary decimal shape.
    ///
    /// Number, Percentage, and Scientific cells may carry only the shared
    /// format kind and primary format identifier.  This check deliberately
    /// does not prove that an identifier resolves; the owning package must
    /// validate the format-table graph separately.
    #[must_use]
    pub fn has_only_decimal_format_metadata(&self) -> bool {
        FIELD_LAYOUT
            .iter()
            .zip(self.fields.iter())
            .all(|((field, _size), value)| {
                value.is_none()
                    || FORMAT_METADATA_FLAGS & field == 0
                    || matches!(*field, CELL_FORMAT_KIND_FLAG | CELL_FORMAT_IDENTIFIER_FLAG)
            })
    }

    /// Reports whether this borrowed cell has a value shape owned by the
    /// focused Duration display-format adapter.
    #[must_use]
    pub fn is_duration_format_compatible(&self) -> bool {
        self.validate_duration_value_shape().is_ok()
    }

    /// Reports whether all present BNC format metadata belongs to the
    /// Duration family.
    #[must_use]
    pub fn has_only_duration_format_metadata(&self) -> bool {
        FIELD_LAYOUT
            .iter()
            .zip(self.fields.iter())
            .all(|((field, _size), value)| {
                value.is_none()
                    || FORMAT_METADATA_FLAGS & field == 0
                    || matches!(
                        *field,
                        CELL_FORMAT_KIND_FLAG
                            | CELL_FORMAT_IDENTIFIER_FLAG
                            | DURATION_FORMAT_IDENTIFIER_FLAG
                    )
            })
    }

    /// Plan a Duration display-format metadata rewrite without allocating.
    ///
    /// The plan includes the optional shared generic secondary identifier that
    /// an explicit native Duration tuple already carries; execution selects
    /// marker `0x0004` without it and `0x0005` with it.
    pub fn plan_duration_format_identifier_rewrite(
        &self,
        identifier: Option<u32>,
    ) -> Result<RewritePlan> {
        let transition = self.validate_duration_transition(identifier)?;
        let removed_flags =
            CELL_FORMAT_KIND_FLAG | CELL_FORMAT_IDENTIFIER_FLAG | DURATION_FORMAT_IDENTIFIER_FLAG;
        let mut replacements = [(CELL_FORMAT_KIND_FLAG, 4usize), (0, 0), (0, 0)];
        let replacement_count = if identifier.is_some() {
            replacements[1] = (DURATION_FORMAT_IDENTIFIER_FLAG, 4);
            if transition.secondary.is_some() {
                replacements[2] = (CELL_FORMAT_IDENTIFIER_FLAG, 4);
                3
            } else {
                2
            }
        } else {
            0
        };
        Ok(RewritePlan {
            output_len: Some(
                self.selected_fields_length(removed_flags, &replacements[..replacement_count])?,
            ),
        })
    }

    /// Rewrite only explicit Duration display-format metadata with one exact,
    /// fallible output allocation.
    ///
    /// The native cell type, every value/cache/formula/style/comment byte,
    /// prefix byte outside the explicit-format marker, and opaque tail are
    /// retained exactly. Automatic Duration tuples and malformed or
    /// cross-family metadata are rejected before allocation.
    pub fn rewrite_duration_format_identifier_with_limit(
        &self,
        identifier: Option<u32>,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        let transition = self.validate_duration_transition(identifier)?;
        let removed_flags =
            CELL_FORMAT_KIND_FLAG | CELL_FORMAT_IDENTIFIER_FLAG | DURATION_FORMAT_IDENTIFIER_FLAG;
        let kind_bytes = DURATION_CELL_FORMAT_KIND.to_le_bytes();
        let identifier_bytes = identifier.map(u32::to_le_bytes);
        let mut replacements: [(u32, &[u8]); 3] = [
            (CELL_FORMAT_KIND_FLAG, &kind_bytes),
            (DURATION_FORMAT_IDENTIFIER_FLAG, &[]),
            (CELL_FORMAT_IDENTIFIER_FLAG, &[]),
        ];
        let replacement_count = if let Some(identifier_bytes) = identifier_bytes.as_ref() {
            replacements[1] = (DURATION_FORMAT_IDENTIFIER_FLAG, identifier_bytes);
            if let Some(secondary) = transition.secondary {
                replacements[2] = (CELL_FORMAT_IDENTIFIER_FLAG, secondary);
                3
            } else {
                2
            }
        } else {
            0
        };
        let mut output = self.rewrite_selected_fields_many(
            self.cell_type,
            removed_flags,
            &replacements[..replacement_count],
            max_output_bytes,
        )?;
        let explicit_flags = identifier.map_or(0, |_| {
            explicit_duration_format_flags(transition.secondary.is_some())
        });
        output[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&explicit_flags.to_le_bytes());

        let candidate = BncCellView::parse(&output)?;
        if candidate.stored_value() != self.stored_value()
            || candidate.cached_scalar() != self.cached_scalar()
            || candidate.opaque_tail() != self.opaque_tail()
            || candidate.prefix[..EXPLICIT_FORMAT_FLAGS_START]
                != self.prefix[..EXPLICIT_FORMAT_FLAGS_START]
            || candidate.prefix[EXPLICIT_FORMAT_FLAGS_END..]
                != self.prefix[EXPLICIT_FORMAT_FLAGS_END..]
            || candidate.cell_type != self.cell_type
            || candidate
                .fields
                .iter()
                .zip(self.fields.iter())
                .enumerate()
                .any(|(index, (candidate, source))| {
                    let field = FIELD_LAYOUT[index].0;
                    FORMAT_METADATA_FLAGS & field == 0 && candidate != source
                })
        {
            return Err(Error::InvalidFormat(
                "Duration metadata readback changed a non-format BNC field".to_owned(),
            ));
        }
        if candidate
            .validate_duration_transition(identifier)?
            .current_identifier
            != identifier
        {
            return Err(Error::InvalidFormat(
                "Duration metadata readback differs from the request".to_owned(),
            ));
        }
        Ok(output)
    }

    /// Returns the unparsed bytes after the final known fixed-width field.
    ///
    /// The bytes are opaque and remain borrowed from the source.  A physical
    /// row owner may preserve them by moving the complete cell/row envelope,
    /// but must not interpret them as safe row-affine state without a focused
    /// producer proof.
    #[must_use]
    pub const fn opaque_tail(&self) -> &'a [u8] {
        self.tail
    }

    /// Reports whether the cell carries the reserved known BNC v5 field.
    ///
    /// The parser retains this fixed-width field for byte preservation, but
    /// its semantics are intentionally unmodeled.  Focused physical owners
    /// should reject it unless they have a producer-specific proof.
    #[must_use]
    pub const fn has_reserved_known_field(&self) -> bool {
        self.flags & RESERVED_KNOWN_FIELD_FLAG != 0
    }

    /// Return whether applying one scalar replacement would leave the public
    /// stored value unchanged.
    ///
    /// This comparison is allocation-free and follows the same format-aware
    /// number conversion as [`Self::rewrite_scalar_with_limit`]. It compares
    /// semantic scalar state only; retained styles, formats, comments, and the
    /// opaque tail do not affect the result.
    #[must_use]
    pub fn scalar_equals(&self, expected: ScalarValue) -> bool {
        match expected {
            ScalarValue::String(identifier) => self.stored_value() == StoredValue::Text(identifier),
            ScalarValue::RichText(identifier) => {
                self.stored_value() == StoredValue::RichText(identifier)
            },
            ScalarValue::Number(value)
                if self.u32_field(CELL_FORMAT_KIND_FLAG) == Some(DURATION_CELL_FORMAT_KIND) =>
            {
                let Ok(seconds) = finite_spreadsheet_days_to_seconds(value) else {
                    return false;
                };
                self.stored_value() == StoredValue::Duration
                    && self.cached_scalar == Some(CachedScalar::Duration(seconds))
            },
            ScalarValue::Number(value) => {
                self.stored_value() == StoredValue::Number
                    && self.cached_scalar == Some(CachedScalar::Number(value))
            },
            ScalarValue::Boolean(value) => {
                self.stored_value() == StoredValue::Boolean
                    && self.cached_scalar == Some(CachedScalar::Boolean(value))
            },
            ScalarValue::Date(value) => {
                self.stored_value() == StoredValue::Date
                    && self.cached_scalar == Some(CachedScalar::Date(value))
            },
            ScalarValue::Duration(value) => {
                self.stored_value() == StoredValue::Duration
                    && self.cached_scalar == Some(CachedScalar::Duration(value))
            },
        }
    }

    /// Replace all value-bearing fields in the raw cell with one scalar.
    ///
    /// Prefix bytes other than the cell type, every non-value field byte, and
    /// the opaque tail are retained exactly. Formula and formula-error fields
    /// are value-bearing and are therefore removed. The exact encoded length
    /// is checked before one fallible output allocation.
    ///
    /// # Errors
    ///
    /// Returns an error when a format-aware number conversion or decimal128
    /// encoding fails, the exact output exceeds `max_output_bytes`, encoded
    /// length arithmetic overflows, or the output allocation fails.
    pub fn rewrite_scalar_with_limit(
        &self,
        value: ScalarValue,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        let encoded = self.encode_scalar(value)?;
        self.rewrite_value_fields(
            encoded.cell_type,
            Some((encoded.flag, &encoded.bytes[..encoded.length])),
            max_output_bytes,
        )
    }

    /// Remove all value-bearing fields while retaining raw metadata and tail
    /// bytes exactly.
    ///
    /// The minimal empty BNC representation returns [`ClearValue::Delete`]
    /// without allocating. Otherwise the retained representation uses one
    /// exact fallible allocation bounded by `max_output_bytes`.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained output exceeds `max_output_bytes`,
    /// encoded length arithmetic overflows, or allocation fails.
    pub fn clear_value_with_limit(&self, max_output_bytes: usize) -> Result<ClearValue> {
        let retained_flags = self.flags & !VALUE_FLAGS;
        if retained_flags == 0
            && self.tail.is_empty()
            && self.prefix[0] == BNC_VERSION
            && self.prefix[2..].iter().all(|byte| *byte == 0)
        {
            return Ok(ClearValue::Delete);
        }
        self.rewrite_value_fields(CELL_TYPE_EMPTY, None, max_output_bytes)
            .map(ClearValue::Retain)
    }

    /// Remove exactly one expected comment identifier while preserving every
    /// other encoded field, prefix byte, and opaque tail byte.
    ///
    /// The exact output length is checked before one fallible allocation. A
    /// missing or different comment identifier is rejected so callers cannot
    /// accidentally authorize a stale graph deletion.
    pub fn clear_comment_with_limit(
        &self,
        expected_identifier: u32,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        if expected_identifier == 0 || self.comment_identifier() != Some(expected_identifier) {
            return Err(Error::InvalidFormat(
                "Numbers comment clear targeted a different cell comment".to_owned(),
            ));
        }
        let output =
            self.rewrite_selected_fields_many(self.cell_type, COMMENT_FLAG, &[], max_output_bytes)?;
        let candidate = BncCellView::parse(&output)?;
        if candidate.comment_identifier().is_some()
            || candidate.stored_value() != self.stored_value()
            || candidate.tail != self.tail
        {
            return Err(Error::InvalidFormat(
                "Numbers comment clear readback differs from the request".to_owned(),
            ));
        }
        Ok(output)
    }

    /// Return whether a formula cell already carries the requested supported
    /// display cache.
    ///
    /// Only numeric and Boolean cache values are writable by the bounded raw
    /// cache path. Other cached scalar kinds return `false`.
    #[must_use]
    pub fn formula_cache_equals(&self, expected: CachedScalar) -> bool {
        matches!(self.stored_value(), StoredValue::Formula(_))
            && matches!(expected, CachedScalar::Number(_) | CachedScalar::Boolean(_))
            && self.cached_scalar == Some(expected)
    }

    /// Return whether this cell carries exactly the requested formula key and
    /// typed cached scalar, including finite floating-point bit equality.
    pub fn formula_value_equals(&self, identifier: u32, expected: ScalarValue) -> Result<bool> {
        Ok(self.stored_value() == StoredValue::Formula(identifier)
            && self.formula_cache_matches_scalar(expected)?)
    }

    /// Replace only a formula cell's supported display-cache fields.
    ///
    /// The formula and formula-error identifiers, format/style/comment fields,
    /// prefix bytes other than the cache type, and opaque tail are retained
    /// exactly. The exact encoded length is checked before one fallible output
    /// allocation.
    ///
    /// # Errors
    ///
    /// Returns an error when the target is not a formula cell, the requested
    /// cache is not numeric or Boolean, encoding fails, the exact output
    /// exceeds `max_output_bytes`, or allocation fails.
    pub fn rewrite_formula_cache_with_limit(
        &self,
        value: CachedScalar,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        if !matches!(self.stored_value(), StoredValue::Formula(_)) {
            return Err(Error::InvalidFormat(
                "Numbers formula cache update targeted a cell without a formula".to_owned(),
            ));
        }
        let scalar = match value {
            CachedScalar::Number(value) => ScalarValue::Number(value),
            CachedScalar::Boolean(value) => ScalarValue::Boolean(value),
            CachedScalar::Date(_) | CachedScalar::Duration(_) | CachedScalar::Unsupported(_) => {
                return Err(Error::InvalidFormat(
                    "Numbers formula cache update supports only number and Boolean values"
                        .to_owned(),
                ));
            },
        };
        let encoded = self.encode_scalar(scalar)?;
        if !matches!(
            cached_scalar_from(
                encoded.cell_type,
                decode_scalar_fields(|flag| {
                    (flag == encoded.flag).then_some(&encoded.bytes[..encoded.length])
                })?,
            ),
            Some(CachedScalar::Number(_) | CachedScalar::Boolean(_))
        ) {
            return Err(Error::InvalidFormat(
                "Numbers formula cache encoding changed the supported cache kind".to_owned(),
            ));
        }
        self.rewrite_selected_fields(
            encoded.cell_type,
            FORMULA_CACHE_FLAGS,
            Some((encoded.flag, &encoded.bytes[..encoded.length])),
            max_output_bytes,
        )
    }

    /// Replace the value fields with one typed cache and attach a formula key
    /// in the same bounded raw rewrite.
    ///
    /// The encoded cache is independently checked for exact kind and finite
    /// bits before any result is returned. In particular, a duration-formatted
    /// source cannot coerce a requested number into a duration cache.
    pub fn rewrite_formula_with_limit(
        &self,
        identifier: u32,
        cache: ScalarValue,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        if identifier == 0 || matches!(cache, ScalarValue::RichText(_)) {
            return Err(Error::InvalidFormat(
                "Numbers formula key/cache is invalid".to_owned(),
            ));
        }
        let encoded = self.encode_scalar(cache)?;
        if !encoded_cache_matches(cache, &encoded)? {
            return Err(Error::InvalidFormat(
                "Numbers formula cache encoding changed the requested kind".to_owned(),
            ));
        }
        let formula = identifier.to_le_bytes();
        let replacements = [
            (encoded.flag, &encoded.bytes[..encoded.length]),
            (FORMULA_FLAG, formula.as_slice()),
        ];
        let output = self.rewrite_selected_fields_many(
            encoded.cell_type,
            VALUE_FLAGS,
            &replacements,
            max_output_bytes,
        )?;
        let view = BncCellView::parse(&output)?;
        if view.stored_value() != StoredValue::Formula(identifier)
            || !view.formula_cache_matches_scalar(cache)?
        {
            return Err(Error::InvalidFormat(
                "Numbers formula cache readback differs from the request".to_owned(),
            ));
        }
        Ok(output)
    }

    /// Replace all value fields with a formula reference and no display cache.
    pub fn rewrite_formula_without_cache_with_limit(
        &self,
        identifier: u32,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        if identifier == 0 {
            return Err(Error::InvalidFormat(
                "Numbers formula key is invalid".to_owned(),
            ));
        }
        let formula = identifier.to_le_bytes();
        let output = self.rewrite_selected_fields_many(
            CELL_TYPE_EMPTY,
            VALUE_FLAGS,
            &[(FORMULA_FLAG, formula.as_slice())],
            max_output_bytes,
        )?;
        let view = BncCellView::parse(&output)?;
        if view.stored_value() != StoredValue::Formula(identifier) || view.cached_scalar().is_some()
        {
            return Err(Error::InvalidFormat(
                "Numbers formula readback differs from the request".to_owned(),
            ));
        }
        Ok(output)
    }

    /// Return whether the formula key and typed cache exactly match.
    pub fn formula_and_cache_equal(&self, identifier: u32, cache: ScalarValue) -> Result<bool> {
        Ok(self.stored_value() == StoredValue::Formula(identifier)
            && self.formula_cache_matches_scalar(cache)?)
    }

    /// Returns the native formula-error table identifier when present.
    #[must_use]
    pub fn formula_error_identifier(&self) -> Option<u32> {
        self.u32_field(FORMULA_ERROR_FLAG)
    }

    /// Returns the native comment table identifier when present.
    #[must_use]
    pub fn comment_identifier(&self) -> Option<u32> {
        self.u32_field(COMMENT_FLAG)
    }

    /// Returns the raw rich-text key carried by the cell, regardless of its
    /// native cell type. Presence is reported even when the key is zero.
    #[must_use]
    pub fn rich_text_identifier(&self) -> Option<u32> {
        self.u32_field(RICH_TEXT_FLAG)
    }

    fn validate_duration_transition(
        &self,
        identifier: Option<u32>,
    ) -> Result<DurationViewTransition<'a>> {
        let current_identifier = validate_duration_metadata_tuple(
            self.explicit_format_flags(),
            self.u32_field(CELL_FORMAT_KIND_FLAG),
            self.u32_field(DURATION_FORMAT_IDENTIFIER_FLAG),
            self.u32_field(CELL_FORMAT_IDENTIFIER_FLAG),
            identifier,
        )?;
        if self.flags & RESERVED_KNOWN_FIELD_FLAG != 0 {
            return Err(Error::InvalidFormat(
                "Duration metadata cannot rewrite a reserved BNC field".to_owned(),
            ));
        }
        if self.flags
            & FORMAT_METADATA_FLAGS
            & !(CELL_FORMAT_KIND_FLAG
                | CELL_FORMAT_IDENTIFIER_FLAG
                | DURATION_FORMAT_IDENTIFIER_FLAG)
            != 0
        {
            return Err(Error::InvalidFormat(
                "Duration metadata has an incompatible secondary or control reference".to_owned(),
            ));
        }
        self.validate_duration_value_shape()?;
        Ok(DurationViewTransition {
            current_identifier,
            secondary: self.field(CELL_FORMAT_IDENTIFIER_FLAG),
        })
    }

    fn validate_duration_value_shape(&self) -> Result<()> {
        validate_duration_value_shape_parts(
            self.cell_type,
            self.flags & VALUE_FLAGS,
            self.cached_scalar,
            self.field(FORMULA_FLAG).is_some(),
            self.u32_field(FORMULA_FLAG),
            self.field(STRING_FLAG).is_some(),
            self.u32_field(STRING_FLAG),
            self.field(FORMULA_ERROR_FLAG).is_some(),
            self.u32_field(FORMULA_ERROR_FLAG),
        )
    }

    fn encode_scalar(&self, value: ScalarValue) -> Result<EncodedScalar> {
        let mut encoded = EncodedScalar {
            cell_type: CELL_TYPE_EMPTY,
            flag: 0,
            bytes: [0; 16],
            length: 0,
        };
        match value {
            ScalarValue::String(identifier) => {
                encoded.cell_type = CELL_TYPE_TEXT;
                encoded.flag = STRING_FLAG;
                encoded.bytes[..4].copy_from_slice(&identifier.to_le_bytes());
                encoded.length = 4;
            },
            ScalarValue::RichText(identifier) => {
                encoded.cell_type = CELL_TYPE_RICH_TEXT_OR_NUMBER;
                encoded.flag = RICH_TEXT_FLAG;
                encoded.bytes[..4].copy_from_slice(&identifier.to_le_bytes());
                encoded.length = 4;
            },
            ScalarValue::Number(value)
                if self.u32_field(CELL_FORMAT_KIND_FLAG) == Some(DURATION_CELL_FORMAT_KIND) =>
            {
                encoded.cell_type = CELL_TYPE_DURATION;
                encoded.flag = NUMBER_FLAG;
                encoded.bytes[..8].copy_from_slice(
                    &finite_spreadsheet_days_to_seconds(value)?
                        .get()
                        .to_le_bytes(),
                );
                encoded.length = 8;
            },
            ScalarValue::Number(value) => {
                encoded.cell_type = match self.u32_field(CELL_FORMAT_KIND_FLAG) {
                    Some(CURRENCY_CELL_FORMAT_KIND) => CELL_TYPE_ALTERNATE_NUMBER,
                    Some(DATE_TIME_CELL_FORMAT_KIND) => CELL_TYPE_RICH_TEXT_OR_NUMBER,
                    _ => CELL_TYPE_NUMBER,
                };
                encoded.flag = DECIMAL_FLAG;
                encoded.bytes = decimal128_le(value.get())?;
                encoded.length = 16;
            },
            ScalarValue::Boolean(value) => {
                encoded.cell_type = CELL_TYPE_BOOLEAN;
                encoded.flag = NUMBER_FLAG;
                encoded.bytes[..8]
                    .copy_from_slice(&(if value { 1.0f64 } else { 0.0f64 }).to_le_bytes());
                encoded.length = 8;
            },
            ScalarValue::Date(value) => {
                encoded.cell_type = CELL_TYPE_DATE;
                encoded.flag = DATE_FLAG;
                encoded.bytes[..8].copy_from_slice(&value.get().to_le_bytes());
                encoded.length = 8;
            },
            ScalarValue::Duration(value) => {
                encoded.cell_type = CELL_TYPE_DURATION;
                encoded.flag = NUMBER_FLAG;
                encoded.bytes[..8].copy_from_slice(&value.get().to_le_bytes());
                encoded.length = 8;
            },
        }
        Ok(encoded)
    }

    fn rewrite_value_fields(
        &self,
        cell_type: u8,
        replacement: Option<(u32, &[u8])>,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        self.rewrite_selected_fields(cell_type, VALUE_FLAGS, replacement, max_output_bytes)
    }

    fn rewrite_selected_fields(
        &self,
        cell_type: u8,
        removed_flags: u32,
        replacement: Option<(u32, &[u8])>,
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        match replacement {
            Some(replacement) => self.rewrite_selected_fields_many(
                cell_type,
                removed_flags,
                &[replacement],
                max_output_bytes,
            ),
            None => {
                self.rewrite_selected_fields_many(cell_type, removed_flags, &[], max_output_bytes)
            },
        }
    }

    fn rewrite_selected_fields_many(
        &self,
        cell_type: u8,
        removed_flags: u32,
        replacements: &[(u32, &[u8])],
        max_output_bytes: usize,
    ) -> Result<Vec<u8>> {
        let retained_flags = self.flags & !removed_flags;
        let mut output_flags = retained_flags;
        for (flag, bytes) in replacements {
            if output_flags & flag != 0
                || replacements
                    .iter()
                    .filter(|(candidate, _)| candidate == flag)
                    .count()
                    != 1
                || FIELD_LAYOUT
                    .iter()
                    .find(|(candidate, _)| candidate == flag)
                    .is_none_or(|(_, size)| *size != bytes.len())
            {
                return Err(Error::ParseError(
                    "Numbers BNC replacement fields are invalid".to_owned(),
                ));
            }
            output_flags |= flag;
        }
        let mut output_len = BNC_HEADER_LEN;
        for (flag, size) in FIELD_LAYOUT {
            if output_flags & flag == 0 {
                continue;
            }
            let field_len = replacements
                .iter()
                .find(|(replacement_flag, _bytes)| replacement_flag == flag)
                .map_or(*size, |(_replacement_flag, bytes)| bytes.len());
            if field_len != *size {
                return Err(Error::ParseError(
                    "Numbers BNC replacement field has an invalid width".to_owned(),
                ));
            }
            output_len = output_len.checked_add(field_len).ok_or_else(|| {
                Error::ParseError("Numbers BNC encoded length overflow".to_owned())
            })?;
        }
        output_len = output_len
            .checked_add(self.tail.len())
            .ok_or_else(|| Error::ParseError("Numbers BNC encoded length overflow".to_owned()))?;
        check_output_limit(output_len, max_output_bytes)?;

        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|_error| Error::Allocation {
                requested: output_len,
            })?;
        if output.capacity() != output_len {
            return Err(Error::Allocation {
                requested: output_len,
            });
        }
        output.extend_from_slice(self.prefix);
        output[1] = cell_type;
        output.extend_from_slice(&output_flags.to_le_bytes());
        for (flag, _size) in FIELD_LAYOUT {
            if let Some((_replacement_flag, bytes)) = replacements
                .iter()
                .find(|(replacement_flag, _bytes)| replacement_flag == flag)
            {
                output.extend_from_slice(bytes);
            } else if retained_flags & flag != 0 {
                output.extend_from_slice(self.field(*flag).ok_or_else(|| {
                    Error::ParseError("Numbers BNC retained field is missing".to_owned())
                })?);
            }
        }
        output.extend_from_slice(self.tail);
        if output.len() != output_len {
            return Err(Error::ParseError(
                "Numbers BNC encoded length changed during publication".to_owned(),
            ));
        }
        Ok(output)
    }

    fn formula_cache_matches_scalar(&self, expected: ScalarValue) -> Result<bool> {
        Ok(match expected {
            ScalarValue::String(identifier) => {
                self.prefix[1] == CELL_TYPE_TEXT && self.u32_field(STRING_FLAG) == Some(identifier)
            },
            ScalarValue::Number(value) => self.cached_scalar == Some(CachedScalar::Number(value)),
            ScalarValue::Boolean(value) => self.cached_scalar == Some(CachedScalar::Boolean(value)),
            ScalarValue::Date(value) => self.cached_scalar == Some(CachedScalar::Date(value)),
            ScalarValue::Duration(value) => {
                self.cached_scalar == Some(CachedScalar::Duration(value))
            },
            ScalarValue::RichText(_) => false,
        })
    }

    fn field(&self, requested_flag: u32) -> Option<&'a [u8]> {
        field_from_layout(&self.fields, requested_flag)
    }

    fn u32_field(&self, flag: u32) -> Option<u32> {
        let bytes: [u8; 4] = self.field(flag)?.try_into().ok()?;
        Some(u32::from_le_bytes(bytes))
    }
}

fn field_from_layout<'a>(
    fields: &[Option<&'a [u8]>; FIELD_COUNT],
    requested_flag: u32,
) -> Option<&'a [u8]> {
    let index = usize::try_from(requested_flag.trailing_zeros()).ok()?;
    FIELD_LAYOUT
        .get(index)
        .filter(|(flag, _size)| *flag == requested_flag)?;
    fields.get(index).copied().flatten()
}

fn check_output_limit(observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(Error::OutputLimitExceeded { observed, maximum });
    }
    Ok(())
}

fn validate_duration_field_storage(fields: &BTreeMap<u32, Vec<u8>>) -> Result<()> {
    for (&flag, bytes) in fields {
        let Some((_, expected_width)) = FIELD_LAYOUT
            .iter()
            .find(|(candidate, _)| *candidate == flag)
        else {
            return Err(Error::InvalidFormat(
                "Duration metadata encountered an unknown BNC field".to_owned(),
            ));
        };
        if bytes.len() != *expected_width {
            return Err(Error::InvalidFormat(
                "Duration metadata encountered a malformed BNC field".to_owned(),
            ));
        }
    }
    Ok(())
}

fn validate_duration_metadata_tuple(
    marker: u16,
    kind: Option<u32>,
    primary_identifier: Option<u32>,
    secondary_identifier: Option<u32>,
    requested_identifier: Option<u32>,
) -> Result<Option<u32>> {
    if requested_identifier.is_some_and(|identifier| identifier == 0) {
        return Err(Error::InvalidFormat(
            "Duration format identifier must be non-zero".to_owned(),
        ));
    }
    if primary_identifier.is_some_and(|identifier| identifier == 0) {
        return Err(Error::InvalidFormat(
            "Duration format identifier must be non-zero".to_owned(),
        ));
    }
    if secondary_identifier.is_some_and(|identifier| identifier == 0) {
        return Err(Error::InvalidFormat(
            "Duration secondary format identifier must be non-zero".to_owned(),
        ));
    }

    match (marker, kind, primary_identifier, secondary_identifier) {
        (0, None, None, None) => Ok(None),
        // Numbers 14.4 writes the base marker for a primary-only tuple.
        (EXPLICIT_DURATION_FORMAT, Some(DURATION_CELL_FORMAT_KIND), Some(identifier), None)
            if identifier != 0 =>
        {
            Ok(Some(identifier))
        },
        // Retaining the shared generic Number reference selects the native
        // Duration-with-Number marker.
        (
            EXPLICIT_DURATION_WITH_NUMBER_FORMAT,
            Some(DURATION_CELL_FORMAT_KIND),
            Some(identifier),
            Some(_),
        ) if identifier != 0 => Ok(Some(identifier)),
        // Native Numbers uses marker zero for automatic Duration metadata.
        // This focused primitive owns only explicit transitions and therefore
        // must not silently adopt or clear an automatic tuple.
        (0, Some(DURATION_CELL_FORMAT_KIND), Some(_), _) => Err(Error::InvalidFormat(
            "automatic Duration metadata is not owned by this transition".to_owned(),
        )),
        (0, _, _, _) => Err(Error::InvalidFormat(
            "Duration metadata has an incomplete explicit tuple".to_owned(),
        )),
        _ => Err(Error::InvalidFormat(
            "Duration metadata has an incompatible explicit-format marker or tuple".to_owned(),
        )),
    }
}

fn validate_duration_value_shape_parts(
    cell_type: u8,
    value_flags: u32,
    cached_scalar: Option<CachedScalar>,
    has_formula: bool,
    formula_identifier: Option<u32>,
    has_string: bool,
    string_identifier: Option<u32>,
    has_formula_error: bool,
    formula_error_identifier: Option<u32>,
) -> Result<()> {
    match cell_type {
        CELL_TYPE_EMPTY => {
            if value_flags != 0 || cached_scalar.is_some() {
                return Err(Error::InvalidFormat(
                    "Duration empty cell has an incompatible value shape".to_owned(),
                ));
            }
        },
        CELL_TYPE_DURATION => {
            let allowed_flags = NUMBER_FLAG | FORMULA_FLAG | STRING_FLAG | FORMULA_ERROR_FLAG;
            if value_flags & !allowed_flags != 0
                || value_flags & NUMBER_FLAG == 0
                || !matches!(cached_scalar, Some(CachedScalar::Duration(_)))
            {
                return Err(Error::InvalidFormat(
                    "Duration type-7 cell has an incompatible or ambiguous value shape".to_owned(),
                ));
            }
            if has_formula {
                if formula_identifier.is_none_or(|identifier| identifier == 0) {
                    return Err(Error::InvalidFormat(
                        "Duration formula identifier must be non-zero".to_owned(),
                    ));
                }
            } else if has_string || has_formula_error {
                return Err(Error::InvalidFormat(
                    "Duration cache has formula-only fields without a formula".to_owned(),
                ));
            }
            if has_string && string_identifier.is_none_or(|identifier| identifier == 0)
                || has_formula_error
                    && formula_error_identifier.is_none_or(|identifier| identifier == 0)
            {
                return Err(Error::InvalidFormat(
                    "Duration formula cache reference must be non-zero".to_owned(),
                ));
            }
        },
        _ => {
            return Err(Error::InvalidFormat(
                "Duration transition requires a native Empty or type-7 Duration cell".to_owned(),
            ));
        },
    }
    Ok(())
}

fn stored_value_from(cell_type: u8, mut u32_field: impl FnMut(u32) -> Option<u32>) -> StoredValue {
    if let Some(identifier) = u32_field(FORMULA_FLAG) {
        return StoredValue::Formula(identifier);
    }
    match cell_type {
        CELL_TYPE_EMPTY => StoredValue::Empty,
        CELL_TYPE_NUMBER | CELL_TYPE_ALTERNATE_NUMBER => StoredValue::Number,
        CELL_TYPE_TEXT => u32_field(STRING_FLAG).map_or(StoredValue::Empty, StoredValue::Text),
        CELL_TYPE_DATE => StoredValue::Date,
        CELL_TYPE_BOOLEAN => StoredValue::Boolean,
        CELL_TYPE_DURATION => StoredValue::Duration,
        CELL_TYPE_ERROR => StoredValue::Error,
        CELL_TYPE_RICH_TEXT_OR_NUMBER => {
            if let Some(identifier) = u32_field(RICH_TEXT_FLAG) {
                StoredValue::RichText(identifier)
            } else if let Some(identifier) = u32_field(STRING_FLAG) {
                StoredValue::Text(identifier)
            } else {
                StoredValue::Number
            }
        },
        other => StoredValue::Unsupported(other),
    }
}

fn decode_scalar_fields<'a>(
    mut field: impl FnMut(u32) -> Option<&'a [u8]>,
) -> Result<DecodedScalarFields> {
    Ok(DecodedScalarFields {
        decimal: field(DECIMAL_FLAG).map(decode_decimal128_le).transpose()?,
        number: field(NUMBER_FLAG).map(read_f64_le).transpose()?,
        date: field(DATE_FLAG).map(read_f64_le).transpose()?,
    })
}

fn cached_scalar_from(cell_type: u8, fields: DecodedScalarFields) -> Option<CachedScalar> {
    match cell_type {
        CELL_TYPE_NUMBER | CELL_TYPE_RICH_TEXT_OR_NUMBER | CELL_TYPE_ALTERNATE_NUMBER => fields
            .decimal
            .or(fields.number)
            .map(CachedScalar::Number)
            .or(Some(CachedScalar::Unsupported(cell_type))),
        CELL_TYPE_TEXT | CELL_TYPE_ERROR => Some(CachedScalar::Unsupported(cell_type)),
        CELL_TYPE_DATE => fields
            .date
            .map(CachedScalar::Date)
            .or(Some(CachedScalar::Unsupported(CELL_TYPE_DATE))),
        CELL_TYPE_BOOLEAN => fields
            .number
            .map(|number| CachedScalar::Boolean(number.get() != 0.0))
            .or(Some(CachedScalar::Unsupported(CELL_TYPE_BOOLEAN))),
        CELL_TYPE_DURATION => fields
            .number
            .map(CachedScalar::Duration)
            .or(Some(CachedScalar::Unsupported(CELL_TYPE_DURATION))),
        CELL_TYPE_EMPTY => None,
        other => Some(CachedScalar::Unsupported(other)),
    }
}

fn read_f64_le(data: &[u8]) -> Result<FiniteF64> {
    let bytes: [u8; 8] = data
        .try_into()
        .map_err(|_error| Error::ParseError("Expected an eight-byte Numbers field".to_owned()))?;
    let value = f64::from_le_bytes(bytes);
    FiniteF64::new(value).map_err(|_error| {
        Error::ParseError("Numbers BNC scalar field must contain a finite value".to_owned())
    })
}

fn spreadsheet_days_to_seconds(days: f64) -> Result<f64> {
    let seconds = days * SECONDS_PER_DAY;
    if !seconds.is_finite() {
        return Err(Error::ParseError(
            "Numbers duration conversion exceeds the finite f64 range".to_owned(),
        ));
    }
    Ok(seconds)
}

fn finite_spreadsheet_days_to_seconds(days: FiniteF64) -> Result<FiniteF64> {
    FiniteF64::new(spreadsheet_days_to_seconds(days.get())?).map_err(|_error| {
        Error::ParseError("Numbers duration conversion exceeds the finite f64 range".to_owned())
    })
}

/// Decodes a little-endian IEEE 754 decimal128 value from a Numbers field.
///
/// # Errors
///
/// Returns an error when `data` is not exactly one decimal128 value or the
/// decoded result is non-finite.
pub fn read_decimal128_le(data: &[u8]) -> Result<f64> {
    decode_decimal128_le(data).map(FiniteF64::get)
}

fn decode_decimal128_le(data: &[u8]) -> Result<FiniteF64> {
    if data.len() != 16 {
        return Err(Error::ParseError(
            "Expected a sixteen-byte Numbers decimal128 field".to_owned(),
        ));
    }
    let exponent = (u16::from(data[15] & 0x7f) << 7) | u16::from(data[14] >> 1);
    let mut coefficient = f64::from(data[14] & 1);
    for byte in data[..14].iter().rev() {
        coefficient = coefficient * 256.0 + f64::from(*byte);
    }
    let signed_coefficient = if data[15] & 0x80 != 0 {
        -coefficient
    } else {
        coefficient
    };
    let value = signed_coefficient * 10f64.powi(i32::from(exponent) - DECIMAL128_EXPONENT_BIAS);
    FiniteF64::new(value).map_err(|_error| {
        Error::ParseError("Numbers BNC decimal128 field must decode to a finite value".to_owned())
    })
}

/// Encode the finite `f64`'s shortest round-tripping decimal spelling into
/// the little-endian IEEE 754 decimal128 layout used by Numbers BNC cells and
/// formula AST compatibility fields.
///
/// # Errors
///
/// Returns an error when `value` is non-finite, its coefficient exceeds the
/// decimal128 precision, or its exponent cannot be represented.
pub fn decimal128_le(value: f64) -> Result<[u8; 16]> {
    if !value.is_finite() {
        return Err(Error::ParseError(
            "Numbers cannot encode a non-finite decimal value".to_owned(),
        ));
    }
    let negative = value.is_sign_negative();
    let magnitude = value.abs();
    let mut formatting_buffer = ryu::Buffer::new();
    let spelling = if magnitude == 0.0 {
        "0"
    } else {
        formatting_buffer.format_finite(magnitude)
    };
    let (mantissa, explicit_exponent) = spelling
        .split_once(['e', 'E'])
        .map_or((spelling, 0), |(mantissa, exponent)| {
            (mantissa, exponent.parse::<i32>().unwrap_or(i32::MIN))
        });
    if explicit_exponent == i32::MIN {
        return Err(Error::ParseError(format!(
            "Could not encode Numbers decimal {spelling:?}"
        )));
    }
    let fractional_digit_count = mantissa
        .split_once('.')
        .map_or(0usize, |(_, fraction)| fraction.len());
    let mut coefficient = 0u128;
    let mut digit_count = 0usize;
    let mut trailing_zeroes = 0i32;
    for byte in mantissa.bytes() {
        if byte == b'.' {
            continue;
        }
        let digit = byte
            .checked_sub(b'0')
            .filter(|digit| *digit <= 9)
            .ok_or_else(|| {
                Error::ParseError(format!("Could not encode Numbers decimal {spelling:?}"))
            })?;
        coefficient = coefficient
            .checked_mul(10)
            .and_then(|value| value.checked_add(u128::from(digit)))
            .ok_or_else(|| {
                Error::ParseError(format!("Could not encode Numbers decimal {spelling:?}"))
            })?;
        digit_count = digit_count
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers decimal digit count overflow".to_owned()))?;
        trailing_zeroes = if digit == 0 {
            trailing_zeroes.checked_add(1).ok_or_else(|| {
                Error::ParseError("Numbers decimal trailing-zero count overflow".to_owned())
            })?
        } else {
            0
        };
    }
    if digit_count == 0 {
        return Err(Error::ParseError(format!(
            "Could not encode Numbers decimal {spelling:?}"
        )));
    }
    if coefficient == 0 {
        trailing_zeroes = 0;
    } else {
        let mut remaining_zeroes = trailing_zeroes;
        while remaining_zeroes > 0 {
            coefficient /= 10;
            remaining_zeroes -= 1;
        }
    }
    if coefficient >= (1u128 << DECIMAL128_COEFFICIENT_BITS) {
        return Err(Error::ParseError(
            "Numbers decimal coefficient exceeds 113 bits".to_owned(),
        ));
    }
    let fractional_digits_i32 = i32::try_from(fractional_digit_count)
        .map_err(|_error| Error::ParseError("Numbers decimal exponent overflow".to_owned()))?;
    let exponent = explicit_exponent
        .checked_sub(fractional_digits_i32)
        .and_then(|exponent_value| exponent_value.checked_add(trailing_zeroes))
        .ok_or_else(|| Error::ParseError("Numbers decimal exponent overflow".to_owned()))?;
    let biased_exponent = exponent
        .checked_add(DECIMAL128_EXPONENT_BIAS)
        .filter(|exponent_value| (0..=0x3fff).contains(exponent_value))
        .ok_or_else(|| Error::ParseError("Numbers decimal exponent is out of range".to_owned()))?;
    let biased_exponent_u128 = u128::try_from(biased_exponent)
        .map_err(|_error| Error::ParseError("Numbers decimal exponent is negative".to_owned()))?;
    let mut encoded = coefficient | (biased_exponent_u128 << DECIMAL128_COEFFICIENT_BITS);
    if negative {
        encoded |= 1u128 << DECIMAL128_SIGN_BIT;
    }
    Ok(encoded.to_le_bytes())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "the codec fixtures use fixed, locally constructed values whose failure should abort the test"
    )]

    use super::*;

    fn finite(value: f64) -> FiniteF64 {
        FiniteF64::new(value).expect("finite test scalar")
    }

    fn value_fields(cell: &BncCell) -> Vec<(u32, Vec<u8>)> {
        cell.fields
            .iter()
            .filter(|(flag, _value)| VALUE_FLAGS & **flag != 0)
            .map(|(flag, value)| (*flag, value.clone()))
            .collect()
    }

    #[test]
    fn changes_value_without_changing_style_fields() {
        let original = hex("050300000000000008100200040000000500000001000000");
        let mut cell = BncCell::parse(&original).unwrap();
        assert_eq!(cell.stored_value(), StoredValue::Text(4));

        cell.set_number(42.5).unwrap();
        let encoded = cell.encode();
        let reparsed = BncCell::parse(&encoded).unwrap();
        assert_eq!(reparsed.stored_value(), StoredValue::Number);
        assert_eq!(reparsed.fields[&0x0000_1000], 5u32.to_le_bytes());
        assert_eq!(reparsed.fields[&0x0002_0000], 1u32.to_le_bytes());
        assert_eq!(reparsed.fields[&DECIMAL_FLAG], decimal128_le(42.5).unwrap());
    }

    #[test]
    fn raw_scalar_rewrite_matches_owned_codec_and_preserves_non_value_bytes() {
        let mut source = BncCell::minimal();
        source.prefix[2..].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4, 0xe5, 0xf6]);
        source.set_string(17);
        source.set_style_identifier(Some(23));
        source.set_comment_identifier(Some(29));
        source.set_formula_reference(31);
        source
            .fields
            .insert(FORMULA_ERROR_FLAG, 37u32.to_le_bytes().to_vec());
        source.tail.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        let source_bytes = source.encode();
        let view = BncCellView::parse(&source_bytes).unwrap();

        for scalar in [
            ScalarValue::String(41),
            ScalarValue::RichText(43),
            ScalarValue::Number(finite(47.5)),
            ScalarValue::Boolean(true),
            ScalarValue::Date(finite(53.25)),
            ScalarValue::Duration(finite(59.75)),
        ] {
            let rewritten = view.rewrite_scalar_with_limit(scalar, usize::MAX).unwrap();
            let mut expected = source.clone();
            match scalar {
                ScalarValue::String(identifier) => expected.set_string(identifier),
                ScalarValue::RichText(identifier) => expected.set_rich_text(identifier),
                ScalarValue::Number(value) => expected.set_number(value.get()).unwrap(),
                ScalarValue::Boolean(value) => expected.set_boolean(value),
                ScalarValue::Date(value) => expected.set_date(value.get()).unwrap(),
                ScalarValue::Duration(value) => expected.set_duration(value.get()).unwrap(),
            }
            assert_eq!(rewritten, expected.encode());
            assert_eq!(&rewritten[..1], &source_bytes[..1]);
            assert_eq!(
                &rewritten[2..BNC_PREFIX_LEN],
                &source_bytes[2..BNC_PREFIX_LEN]
            );
            assert!(rewritten.ends_with(&source.tail));
        }

        let rewritten = view
            .rewrite_scalar_with_limit(ScalarValue::Boolean(true), usize::MAX)
            .unwrap();
        let reparsed = BncCell::parse(&rewritten).unwrap();
        assert_eq!(reparsed.stored_value(), StoredValue::Boolean);
        assert_eq!(
            reparsed.cached_scalar().unwrap(),
            Some(CachedScalar::Boolean(true))
        );
        assert_eq!(reparsed.style_identifier(), Some(23));
        assert_eq!(reparsed.comment_identifier(), Some(29));
        assert_eq!(reparsed.formula_error_identifier(), None);
    }

    #[test]
    fn raw_scalar_semantic_equality_is_allocation_free_and_format_aware() {
        let mut cell = BncCell::minimal();
        cell.set_data_format_identifier(7, CellDataFormatKind::Duration, None)
            .unwrap();
        cell.set_number(2.0).unwrap();
        let bytes = cell.encode();
        let view = BncCellView::parse(&bytes).unwrap();
        assert!(view.scalar_equals(ScalarValue::Number(finite(2.0))));
        assert!(view.scalar_equals(ScalarValue::Duration(finite(172_800.0))));
        assert!(!view.scalar_equals(ScalarValue::Number(finite(3.0))));
        assert!(!view.scalar_equals(ScalarValue::Boolean(true)));

        let rewritten = view
            .rewrite_scalar_with_limit(ScalarValue::Number(finite(3.0)), usize::MAX)
            .unwrap();
        let rewritten = BncCellView::parse(&rewritten).unwrap();
        assert_eq!(rewritten.stored_value(), StoredValue::Duration);
        assert_eq!(
            rewritten.cached_scalar(),
            Some(CachedScalar::Duration(finite(259_200.0)))
        );
    }

    #[test]
    fn raw_clear_distinguishes_deleted_and_retained_cells() {
        let minimal = BncCell::minimal().encode();
        let view = BncCellView::parse(&minimal).unwrap();
        assert_eq!(view.clear_value_with_limit(0).unwrap(), ClearValue::Delete);

        let mut retained = BncCell::minimal();
        retained.set_number(42.0).unwrap();
        retained.set_comment_identifier(Some(9));
        retained.tail.extend_from_slice(&[0xca, 0xfe]);
        let retained_bytes = retained.encode();
        let view = BncCellView::parse(&retained_bytes).unwrap();
        let cleared = match view.clear_value_with_limit(usize::MAX).unwrap() {
            ClearValue::Delete => panic!("metadata-bearing cell was deleted"),
            ClearValue::Retain(bytes) => bytes,
        };
        let parsed = BncCell::parse(&cleared).unwrap();
        assert_eq!(parsed.stored_value(), StoredValue::Empty);
        assert_eq!(parsed.comment_identifier(), Some(9));
        assert_eq!(parsed.tail, [0xca, 0xfe]);
    }

    #[test]
    fn bounded_raw_and_owned_encoding_use_exact_limits() {
        let mut cell = BncCell::minimal();
        cell.set_string(11);
        cell.set_style_identifier(Some(13));
        cell.tail.extend_from_slice(&[1, 2, 3]);
        let encoded = cell.encode();
        assert_eq!(cell.try_encode_with_limit(encoded.len()).unwrap(), encoded);
        assert!(matches!(
            cell.try_encode_with_limit(encoded.len() - 1),
            Err(Error::OutputLimitExceeded {
                observed,
                maximum
            }) if observed == encoded.len() && maximum + 1 == observed
        ));

        let view = BncCellView::parse(&encoded).unwrap();
        let expected = view
            .rewrite_scalar_with_limit(ScalarValue::RichText(17), usize::MAX)
            .unwrap();
        assert_eq!(
            view.rewrite_scalar_with_limit(ScalarValue::RichText(17), expected.len())
                .unwrap(),
            expected
        );
        assert!(matches!(
            view.rewrite_scalar_with_limit(ScalarValue::RichText(17), expected.len() - 1),
            Err(Error::OutputLimitExceeded {
                observed,
                maximum
            }) if observed == expected.len() && maximum + 1 == observed
        ));
    }

    #[test]
    fn rejects_unknown_flags_and_non_finite_numbers() {
        let mut data = vec![5, 2, 0, 0, 0, 0, 0, 0];
        data.extend_from_slice(&0x8000_0000u32.to_le_bytes());
        assert!(BncCell::parse(&data).is_err());
        assert!(BncCell::minimal().set_number(f64::NAN).is_err());

        let mut non_finite_binary = vec![5, 2, 0, 0, 0, 0, 0, 0];
        non_finite_binary.extend_from_slice(&NUMBER_FLAG.to_le_bytes());
        non_finite_binary.extend_from_slice(&f64::NAN.to_le_bytes());
        assert!(BncCell::parse(&non_finite_binary).is_err());

        let mut non_finite_decimal = vec![5, 2, 0, 0, 0, 0, 0, 0];
        non_finite_decimal.extend_from_slice(&DECIMAL_FLAG.to_le_bytes());
        let mut decimal128_overflow = [0; 16];
        decimal128_overflow[14] = 0xff;
        decimal128_overflow[15] = 0x7f;
        non_finite_decimal.extend_from_slice(&decimal128_overflow);
        assert!(BncCell::parse(&non_finite_decimal).is_err());
        assert!(read_decimal128_le(&decimal128_overflow).is_err());

        for (cell_type, field) in [
            (CELL_TYPE_DATE, DATE_FLAG),
            (CELL_TYPE_BOOLEAN, NUMBER_FLAG),
            (CELL_TYPE_DURATION, NUMBER_FLAG),
            (CELL_TYPE_TEXT, NUMBER_FLAG),
        ] {
            let mut non_finite = vec![5, cell_type, 0, 0, 0, 0, 0, 0];
            non_finite.extend_from_slice(&field.to_le_bytes());
            non_finite.extend_from_slice(&f64::INFINITY.to_le_bytes());
            assert!(BncCell::parse(&non_finite).is_err());
        }
    }

    #[test]
    fn value_and_formula_replacement_clear_cached_formula_error_ids() {
        let mut cell = BncCell::minimal();
        cell.prefix[1] = 8;
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 17u32.to_le_bytes().to_vec());
        assert_eq!(cell.formula_error_identifier(), Some(17));

        cell.set_number(1.0).unwrap();
        assert_eq!(cell.formula_error_identifier(), None);
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 18u32.to_le_bytes().to_vec());
        cell.set_formula_reference(3);
        assert_eq!(cell.formula_error_identifier(), None);
        assert_eq!(cell.stored_value(), StoredValue::Formula(3));
    }

    #[test]
    fn formula_cache_updates_preserve_formula_and_metadata() {
        let mut cell = BncCell::minimal();
        cell.set_comment_identifier(Some(9));
        cell.fields.insert(0x0000_1000, 5u32.to_le_bytes().to_vec());
        cell.set_number(3.0).unwrap();
        cell.set_formula_reference(17);

        cell.set_formula_cached_number(42.5).unwrap();
        assert_eq!(cell.stored_value(), StoredValue::Formula(17));
        assert_eq!(
            cell.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(42.5)))
        );
        assert_eq!(cell.comment_identifier(), Some(9));
        assert_eq!(cell.fields[&0x0000_1000], 5u32.to_le_bytes());

        cell.set_formula_cached_boolean(true).unwrap();
        assert_eq!(cell.stored_value(), StoredValue::Formula(17));
        assert_eq!(
            cell.cached_scalar().unwrap(),
            Some(CachedScalar::Boolean(true))
        );
        assert_eq!(cell.comment_identifier(), Some(9));
        assert_eq!(cell.fields[&0x0000_1000], 5u32.to_le_bytes());

        assert!(BncCell::minimal().set_formula_cached_number(1.0).is_err());
    }

    #[test]
    fn raw_formula_cache_rewrite_preserves_non_cache_bytes_and_is_bounded() {
        let mut cell = BncCell::minimal();
        cell.prefix[2] = 0x5a;
        cell.set_comment_identifier(Some(9));
        cell.fields.insert(STYLE_FLAG, 5u32.to_le_bytes().to_vec());
        cell.set_number(323.0).unwrap();
        cell.set_formula_reference(17);
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 23u32.to_le_bytes().to_vec());
        cell.tail.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        let original = cell.encode();
        let view = BncCellView::parse(&original).unwrap();
        let expected = CachedScalar::Number(finite(324.0));
        assert!(!view.formula_cache_equals(expected));

        let plan = view.plan_formula_cache_rewrite(expected).unwrap();

        let rewritten = view
            .rewrite_formula_cache_with_limit(expected, usize::MAX)
            .unwrap();
        assert_eq!(plan.output_len(), Some(rewritten.len()));
        let reparsed = BncCellView::parse(&rewritten).unwrap();
        assert!(reparsed.formula_cache_equals(expected));
        assert_eq!(reparsed.stored_value(), StoredValue::Formula(17));
        assert_eq!(reparsed.formula_error_identifier(), Some(23));
        assert_eq!(reparsed.comment_identifier(), Some(9));

        let before = BncCell::parse(&original).unwrap();
        let after = BncCell::parse(&rewritten).unwrap();
        assert_eq!(before.prefix[0], after.prefix[0]);
        assert_eq!(&before.prefix[2..], &after.prefix[2..]);
        assert_eq!(before.tail, after.tail);
        for (flag, bytes) in &before.fields {
            if FORMULA_CACHE_FLAGS & flag == 0 {
                assert_eq!(after.fields.get(flag), Some(bytes));
            }
        }

        assert!(matches!(
            view.rewrite_formula_cache_with_limit(expected, rewritten.len() - 1),
            Err(Error::OutputLimitExceeded { observed, maximum })
                if observed == rewritten.len() && maximum == rewritten.len() - 1
        ));
        let boolean = BncCellView::parse(&rewritten)
            .unwrap()
            .rewrite_formula_cache_with_limit(CachedScalar::Boolean(true), usize::MAX)
            .unwrap();
        let boolean = BncCellView::parse(&boolean).unwrap();
        assert!(boolean.formula_cache_equals(CachedScalar::Boolean(true)));
        assert_eq!(boolean.stored_value(), StoredValue::Formula(17));
        assert_eq!(boolean.formula_error_identifier(), Some(23));
        assert_eq!(boolean.comment_identifier(), Some(9));
    }

    #[test]
    fn raw_formula_rewrite_preserves_unknown_bytes_and_all_typed_caches() {
        let mut cell = BncCell::minimal();
        cell.prefix[2] = 0x6a;
        cell.set_comment_identifier(Some(9));
        cell.fields.insert(STYLE_FLAG, 5u32.to_le_bytes().to_vec());
        cell.tail.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        let source = cell.encode();

        for cache in [
            ScalarValue::String(27),
            ScalarValue::Number(finite(-0.0)),
            ScalarValue::Boolean(true),
            ScalarValue::Date(finite(789_332_889.25)),
            ScalarValue::Duration(finite(3_723.5)),
        ] {
            let source_view = BncCellView::parse(&source).unwrap();
            let plan = source_view.plan_formula_rewrite(41, Some(cache)).unwrap();
            let rewritten = source_view
                .rewrite_formula_with_limit(41, cache, usize::MAX)
                .unwrap();
            assert_eq!(plan.output_len(), Some(rewritten.len()));
            let view = BncCellView::parse(&rewritten).unwrap();
            assert!(view.formula_and_cache_equal(41, cache).unwrap());
            assert_eq!(view.comment_identifier(), Some(9));

            let before = BncCell::parse(&source).unwrap();
            let after = BncCell::parse(&rewritten).unwrap();
            assert_eq!(before.prefix[0], after.prefix[0]);
            assert_eq!(&before.prefix[2..], &after.prefix[2..]);
            assert_eq!(before.tail, after.tail);
            for (flag, bytes) in &before.fields {
                if VALUE_FLAGS & flag == 0 {
                    assert_eq!(after.fields.get(flag), Some(bytes));
                }
            }

            assert!(matches!(
                BncCellView::parse(&source)
                    .unwrap()
                    .rewrite_formula_with_limit(41, cache, rewritten.len() - 1),
                Err(Error::OutputLimitExceeded { observed, maximum })
                    if observed == rewritten.len() && maximum == rewritten.len() - 1
            ));
        }

        let view = BncCellView::parse(&source).unwrap();
        let plan = view.plan_formula_rewrite(41, None).unwrap();
        let rewritten = view
            .rewrite_formula_without_cache_with_limit(41, usize::MAX)
            .unwrap();
        assert_eq!(plan.output_len(), Some(rewritten.len()));
        assert!(matches!(
            view.rewrite_formula_without_cache_with_limit(41, rewritten.len() - 1),
            Err(Error::OutputLimitExceeded { observed, maximum })
                if observed == rewritten.len() && maximum == rewritten.len() - 1
        ));
    }

    #[test]
    fn raw_formula_number_refuses_duration_format_coercion() {
        let mut duration = BncCell::minimal();
        duration
            .set_data_format_identifier(9, CellDataFormatKind::Duration, None)
            .unwrap();
        let source = duration.encode();
        assert!(matches!(
            BncCellView::parse(&source)
                .unwrap()
                .rewrite_formula_with_limit(41, ScalarValue::Number(finite(1.0)), usize::MAX),
            Err(Error::InvalidFormat(_))
        ));
        assert_eq!(
            BncCell::parse(&source).unwrap().stored_value(),
            StoredValue::Empty
        );
    }

    #[test]
    fn comments_are_orthogonal_to_cell_values() {
        let mut cell = BncCell::minimal();
        cell.set_comment_identifier(Some(9));
        cell.set_string(3);
        assert_eq!(cell.comment_identifier(), Some(9));
        assert_eq!(cell.stored_value(), StoredValue::Text(3));

        cell.clear_value_preserving_metadata();
        assert_eq!(cell.stored_value(), StoredValue::Empty);
        assert_eq!(cell.comment_identifier(), Some(9));
        cell.set_comment_identifier(None);
        assert_eq!(cell.comment_identifier(), None);
    }

    #[test]
    fn rich_text_identifier_reports_stray_flags_for_non_rich_cell_types() {
        let mut number = BncCell::minimal();
        number.set_number(1.5).unwrap();

        let mut date = BncCell::minimal();
        date.set_date(2.5).unwrap();

        let mut boolean = BncCell::minimal();
        boolean.set_boolean(true);

        let mut duration = BncCell::minimal();
        duration.set_duration(3.5).unwrap();

        let mut text = BncCell::minimal();
        text.set_string(3);

        for (mut cell, expected) in [
            (number, StoredValue::Number),
            (date, StoredValue::Date),
            (boolean, StoredValue::Boolean),
            (duration, StoredValue::Duration),
            (text, StoredValue::Text(3)),
        ] {
            cell.fields
                .insert(RICH_TEXT_FLAG, 17u32.to_le_bytes().to_vec());
            let bytes = cell.encode();
            let owned = BncCell::parse(&bytes).unwrap();
            let view = BncCellView::parse(&bytes).unwrap();

            assert_eq!(owned.rich_text_identifier(), Some(17));
            assert_eq!(view.rich_text_identifier(), Some(17));
            assert_eq!(view.stored_value(), expected);
        }
    }

    #[test]
    fn rich_text_identifier_preserves_zero_presence() {
        let mut cell = BncCell::minimal();
        cell.set_number(1.0).unwrap();
        let without_rich_text_bytes = cell.encode();
        let without_rich_text = BncCellView::parse(&without_rich_text_bytes).unwrap();
        assert_eq!(without_rich_text.rich_text_identifier(), None);

        cell.fields
            .insert(RICH_TEXT_FLAG, 0u32.to_le_bytes().to_vec());
        let bytes = cell.encode();
        assert_eq!(
            BncCell::parse(&bytes).unwrap().rich_text_identifier(),
            Some(0)
        );
        assert_eq!(
            BncCellView::parse(&bytes).unwrap().rich_text_identifier(),
            Some(0)
        );
    }

    #[test]
    fn bounded_comment_clear_preserves_value_metadata_and_tail() {
        let mut cell = BncCell::minimal();
        cell.set_number(42.5).unwrap();
        cell.set_style_identifier(Some(17));
        cell.set_comment_identifier(Some(9));
        cell.tail.extend_from_slice(b"opaque-tail");
        let source = cell.encode();
        let view = BncCellView::parse(&source).unwrap();

        let output = view.clear_comment_with_limit(9, source.len()).unwrap();
        let candidate = BncCellView::parse(&output).unwrap();
        assert_eq!(candidate.comment_identifier(), None);
        assert_eq!(candidate.stored_value(), view.stored_value());
        assert_eq!(candidate.tail, b"opaque-tail");
        assert_eq!(output.len(), source.len() - 4);
        assert!(matches!(
            view.clear_comment_with_limit(9, output.len() - 1),
            Err(Error::OutputLimitExceeded { .. })
        ));
        assert!(view.clear_comment_with_limit(8, usize::MAX).is_err());
        let no_comment = BncCell::minimal().encode();
        assert!(
            BncCellView::parse(&no_comment)
                .unwrap()
                .clear_comment_with_limit(9, usize::MAX)
                .is_err()
        );
    }

    #[test]
    fn conditional_styles_are_orthogonal_to_cell_values() {
        let mut cell = BncCell::minimal();
        cell.set_number(42.0).unwrap();
        cell.set_conditional_style(Some(11), Some(15));

        let mut reparsed = BncCell::parse(&cell.encode()).unwrap();
        assert_eq!(reparsed.conditional_style_identifier(), Some(11));
        assert_eq!(reparsed.conditional_style_applied_rule(), Some(15));
        assert_eq!(reparsed.stored_value(), StoredValue::Number);

        reparsed.set_conditional_style(None, None);
        let cleared = BncCell::parse(&reparsed.encode()).unwrap();
        assert_eq!(cleared.conditional_style_identifier(), None);
        assert_eq!(cleared.conditional_style_applied_rule(), None);
        assert_eq!(cleared.stored_value(), StoredValue::Number);
    }

    #[test]
    fn text_and_cell_styles_use_independent_keys() {
        let mut cell = BncCell::minimal();
        cell.set_string(3);
        cell.set_style_identifier(Some(7));
        cell.set_text_style_identifier(Some(11));

        let mut reparsed = BncCell::parse(&cell.encode()).unwrap();
        assert_eq!(reparsed.style_identifier(), Some(7));
        assert_eq!(reparsed.text_style_identifier(), Some(11));
        assert_eq!(reparsed.stored_value(), StoredValue::Text(3));

        reparsed.set_text_style_identifier(None);
        let cleared = BncCell::parse(&reparsed.encode()).unwrap();
        assert_eq!(cleared.style_identifier(), Some(7));
        assert_eq!(cleared.text_style_identifier(), None);
        assert_eq!(cleared.stored_value(), StoredValue::Text(3));
    }

    #[test]
    fn borrowed_view_exposes_native_metadata_and_opaque_tail() {
        let mut cell = BncCell::minimal();
        cell.prefix[2..].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4, 0xe5, 0xf6]);
        cell.set_number(42.25).unwrap();
        cell.set_style_identifier(Some(7));
        cell.set_text_style_identifier(Some(11));
        cell.set_comment_identifier(Some(13));
        cell.set_conditional_style(Some(17), Some(19));
        cell.set_data_format_identifier(23, CellDataFormatKind::NumberOrPercentage, None)
            .unwrap();
        cell.tail.extend_from_slice(b"opaque-tail");

        let bytes = cell.encode();
        let view = BncCellView::parse(&bytes).unwrap();
        assert_eq!(view.numeric_cell_type(), Some(NumericCellType::Number));
        assert_eq!(view.style_identifier(), Some(7));
        assert_eq!(view.text_style_identifier(), Some(11));
        assert_eq!(view.comment_identifier(), Some(13));
        assert_eq!(view.conditional_style_identifier(), Some(17));
        assert_eq!(view.conditional_style_applied_rule(), Some(19));
        assert_eq!(view.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(view.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(view.format_identifier(), Some(23));
        assert_eq!(view.secondary_format_identifier(), None);
        assert_eq!(view.control_cell_spec_identifier(), None);
        assert!(view.has_only_decimal_format_metadata());
        assert_eq!(view.opaque_tail(), b"opaque-tail");
        assert_eq!(view.opaque_tail().len(), cell.tail.len());
        assert_eq!(view.stored_value(), StoredValue::Number);
    }

    #[test]
    fn borrowed_view_distinguishes_currency_secondary_and_reserved_metadata() {
        let mut cell = BncCell::minimal();
        cell.set_number(42.25).unwrap();
        cell.set_currency_format_identifier_preserving_value(Some(23))
            .unwrap();
        cell.fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 19u32.to_le_bytes().to_vec());
        cell.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_CURRENCY_WITH_NUMBER_FORMAT.to_le_bytes());
        cell.fields
            .insert(RESERVED_KNOWN_FIELD_FLAG, 29u32.to_le_bytes().to_vec());

        let bytes = cell.encode();
        let view = BncCellView::parse(&bytes).unwrap();
        assert_eq!(
            view.numeric_cell_type(),
            Some(NumericCellType::AlternateNumber)
        );
        assert_eq!(
            view.explicit_format_flags(),
            EXPLICIT_CURRENCY_WITH_NUMBER_FORMAT
        );
        assert_eq!(view.cell_format_kind(), Some(CURRENCY_CELL_FORMAT_KIND));
        assert_eq!(view.format_identifier(), Some(23));
        assert_eq!(view.secondary_format_identifier(), Some(19));
        assert!(view.has_reserved_known_field());
        assert!(!view.has_only_decimal_format_metadata());
    }

    #[test]
    fn app_authored_conditional_style_fields_decode_independently() {
        let mut data = vec![5, 2, 0, 0, 0, 0, 0, 0];
        data.extend_from_slice(
            &(DECIMAL_FLAG
                | CONDITIONAL_STYLE_FLAG
                | CONDITIONAL_STYLE_APPLIED_RULE_FLAG
                | CELL_FORMAT_KIND_FLAG
                | CELL_FORMAT_IDENTIFIER_FLAG)
                .to_le_bytes(),
        );
        data.extend_from_slice(&decimal128_le(-5.0).unwrap());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&15u32.to_le_bytes());
        data.extend_from_slice(&DECIMAL_CELL_FORMAT_KIND.to_le_bytes());
        data.extend_from_slice(&2u32.to_le_bytes());

        let cell = BncCell::parse(&data).unwrap();
        assert_eq!(cell.conditional_style_identifier(), Some(1));
        assert_eq!(cell.conditional_style_applied_rule(), Some(15));
        assert_eq!(
            cell.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(-5.0)))
        );
    }

    #[test]
    fn number_formats_are_orthogonal_to_values_and_styles() {
        let mut cell = BncCell::minimal();
        cell.set_number(1_234.5).unwrap();
        cell.set_style_identifier(Some(7));
        cell.set_data_format_identifier(2, CellDataFormatKind::NumberOrPercentage, None)
            .unwrap();

        assert_eq!(cell.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(cell.format_identifier(), Some(2));

        let reparsed = BncCell::parse(&cell.encode()).unwrap();
        assert_eq!(reparsed.stored_value(), StoredValue::Number);
        assert_eq!(reparsed.style_identifier(), Some(7));
        assert_eq!(reparsed.format_identifier(), Some(2));

        cell.clear_explicit_format();
        assert_eq!(cell.explicit_format_flags(), 0);
        assert_eq!(cell.cell_format_kind(), None);
        assert_eq!(cell.format_identifier(), None);
        assert_eq!(cell.stored_value(), StoredValue::Number);
        assert_eq!(cell.style_identifier(), Some(7));
    }

    #[test]
    fn currency_formats_use_native_alternate_number_metadata() {
        let mut cell = BncCell::minimal();
        cell.set_number(-12.345).unwrap();
        cell.set_data_format_identifier(4, CellDataFormatKind::Currency, None)
            .unwrap();

        assert_eq!(cell.explicit_format_flags(), EXPLICIT_CURRENCY_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(CURRENCY_CELL_FORMAT_KIND));
        assert_eq!(cell.format_identifier(), Some(4));
        assert_eq!(cell.stored_value(), StoredValue::Number);

        let converted_native = BncCell::parse(&hex(
            "050a0000000003080170000039300000000000000000000000003ab0020000000200000004000000",
        ))
        .unwrap();
        assert_eq!(converted_native.format_identifier(), Some(4));
        assert_eq!(converted_native.secondary_format_identifier(), Some(2));

        cell.set_number(42.0).unwrap();
        let reparsed = BncCell::parse(&cell.encode()).unwrap();
        assert_eq!(reparsed.explicit_format_flags(), EXPLICIT_CURRENCY_FORMAT);
        assert_eq!(reparsed.format_identifier(), Some(4));

        cell.clear_explicit_format();
        assert_eq!(cell.explicit_format_flags(), 0);
        assert_eq!(cell.cell_format_kind(), None);
        assert_eq!(cell.format_identifier(), None);
        assert_eq!(cell.stored_value(), StoredValue::Number);
    }

    #[test]
    fn date_time_formats_use_native_date_metadata() {
        let mut cell = BncCell::minimal();
        cell.set_date(789_332_889.0).unwrap();
        cell.set_data_format_identifier(7, CellDataFormatKind::DateTime, None)
            .unwrap();

        assert_eq!(cell.explicit_format_flags(), EXPLICIT_DATE_TIME_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(DATE_TIME_CELL_FORMAT_KIND));
        assert_eq!(cell.format_identifier(), Some(7));
        assert_eq!(cell.stored_value(), StoredValue::Date);
        assert_eq!(
            cell.encode(),
            hex("050500000000080004900000000080cc2186c7410300000007000000")
        );

        let mut number = BncCell::minimal();
        number.set_number(-1_234.5).unwrap();
        number
            .set_data_format_identifier(7, CellDataFormatKind::DateTime, None)
            .unwrap();
        assert_eq!(number.stored_value(), StoredValue::Number);
        assert_eq!(number.prefix[1], CELL_TYPE_RICH_TEXT_OR_NUMBER);
        number.clear_explicit_format();
        assert_eq!(number.prefix[1], CELL_TYPE_NUMBER);
        assert_eq!(number.stored_value(), StoredValue::Number);

        cell.clear_explicit_format();
        assert_eq!(cell.stored_value(), StoredValue::Date);
        assert_eq!(cell.explicit_format_flags(), 0);
        assert_eq!(cell.format_identifier(), None);
    }

    #[test]
    fn duration_formats_use_native_duration_metadata_and_scalar_units() {
        let native_automatic = hex("050700000000000002100100000000000017ad400400000008000000");
        let automatic = BncCell::parse(&native_automatic).unwrap();
        assert_eq!(automatic.stored_value(), StoredValue::Duration);
        assert_eq!(automatic.explicit_format_flags(), 0);
        assert_eq!(
            automatic.cell_format_kind(),
            Some(DURATION_CELL_FORMAT_KIND)
        );
        assert_eq!(automatic.format_identifier(), Some(8));
        assert_eq!(automatic.secondary_format_identifier(), None);
        assert_eq!(
            automatic.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(3_723.5)))
        );
        assert_eq!(automatic.encode(), native_automatic);

        let native_converted =
            hex("050700000000050002300100000000000f6e99c1040000000100000009000000");
        let converted = BncCell::parse(&native_converted).unwrap();
        assert_eq!(
            converted.explicit_format_flags(),
            EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(converted.format_identifier(), Some(9));
        assert_eq!(converted.secondary_format_identifier(), Some(1));
        assert_eq!(
            converted.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(-106_660_800.0)))
        );
        assert_eq!(converted.encode(), native_converted);

        let mut number = BncCell::minimal();
        number.set_number(1.5).unwrap();
        number
            .set_data_format_identifier(9, CellDataFormatKind::Duration, None)
            .unwrap();
        assert_eq!(number.stored_value(), StoredValue::Duration);
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(129_600.0)))
        );
        assert_eq!(number.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        assert_eq!(number.format_identifier(), Some(9));
        assert_eq!(number.secondary_format_identifier(), None);

        number
            .set_data_format_identifier(1, CellDataFormatKind::NumberOrPercentage, None)
            .unwrap();
        assert_eq!(number.stored_value(), StoredValue::Number);
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(1.5)))
        );

        number.set_formula_reference(12);
        number
            .set_data_format_identifier(9, CellDataFormatKind::Duration, None)
            .unwrap();
        assert_eq!(number.stored_value(), StoredValue::Formula(12));
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(129_600.0)))
        );
        number
            .set_data_format_identifier(1, CellDataFormatKind::NumberOrPercentage, None)
            .unwrap();
        assert_eq!(number.stored_value(), StoredValue::Formula(12));
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(1.5)))
        );
    }

    #[test]
    fn generic_duration_rejects_incompatible_value_families_atomically() {
        let mut text = BncCell::minimal();
        text.set_string(17);

        let mut date = BncCell::minimal();
        date.set_date(2.5).unwrap();

        let mut boolean = BncCell::minimal();
        boolean.set_boolean(true);

        let mut error = BncCell::minimal();
        error.prefix[1] = CELL_TYPE_ERROR;
        error
            .fields
            .insert(FORMULA_ERROR_FLAG, 23u32.to_le_bytes().to_vec());

        let mut rich_text = BncCell::minimal();
        rich_text.set_rich_text(29);

        for mut cell in [text, date, boolean, error, rich_text] {
            let before = cell.encode();
            assert!(
                cell.set_data_format_identifier(41, CellDataFormatKind::Duration, None)
                    .is_err()
            );
            assert_eq!(cell.encode(), before);
        }

        let mut unsupported = BncCell::minimal();
        unsupported.prefix[1] = 42;
        let before = unsupported.encode();
        assert!(
            unsupported
                .set_data_format_identifier(41, CellDataFormatKind::Duration, None)
                .is_err()
        );
        assert_eq!(unsupported.encode(), before);
    }

    #[test]
    fn generic_duration_rejects_zero_identifier_atomically() {
        let mut cell = BncCell::minimal();
        cell.set_number(1.5).unwrap();
        cell.set_style_identifier(Some(47));
        cell.set_comment_identifier(Some(53));
        cell.tail.extend_from_slice(b"duration-zero-id");
        let before = cell.encode();

        assert!(
            cell.set_data_format_identifier(0, CellDataFormatKind::Duration, None)
                .is_err()
        );
        assert_eq!(cell.encode(), before);
    }

    #[test]
    fn generic_duration_accepts_empty_and_native_duration_values() {
        let mut empty = BncCell::minimal();
        empty.set_style_identifier(Some(101));
        empty.set_comment_identifier(Some(103));
        empty.tail.extend_from_slice(b"duration-empty");
        empty
            .set_data_format_identifier(107, CellDataFormatKind::Duration, None)
            .unwrap();
        assert_eq!(empty.stored_value(), StoredValue::Empty);
        assert_eq!(empty.cached_scalar().unwrap(), None);
        assert_eq!(empty.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        assert_eq!(empty.format_identifier(), Some(107));
        assert_eq!(empty.style_identifier(), Some(101));
        assert_eq!(empty.comment_identifier(), Some(103));
        assert_eq!(empty.tail, b"duration-empty");

        let mut native = BncCell::minimal();
        native.set_duration(3_723.5).unwrap();
        let original_value = value_fields(&native);
        native
            .set_data_format_identifier(109, CellDataFormatKind::Duration, None)
            .unwrap();
        assert_eq!(native.prefix[1], CELL_TYPE_DURATION);
        assert_eq!(native.stored_value(), StoredValue::Duration);
        assert_eq!(
            native.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(3_723.5)))
        );
        assert_eq!(value_fields(&native), original_value);
        assert_eq!(native.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        assert_eq!(native.format_identifier(), Some(109));
    }

    #[test]
    fn generic_duration_conversion_preserves_formula_error_and_canonicalizes_secondary() {
        let mut number = BncCell::minimal();
        number.prefix[2..6].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4]);
        let original_prefix = number.prefix[2..6].to_vec();
        number.set_number(1.5).unwrap();
        number.set_formula_reference(61);
        number
            .fields
            .insert(STRING_FLAG, 67u32.to_le_bytes().to_vec());
        number
            .fields
            .insert(FORMULA_ERROR_FLAG, 71u32.to_le_bytes().to_vec());
        number.set_style_identifier(Some(73));
        number.set_text_style_identifier(Some(75));
        number.set_conditional_style(Some(77), Some(78));
        number.set_comment_identifier(Some(79));
        number.tail.extend_from_slice(b"duration-formula-tail");

        number
            .set_data_format_identifier(83, CellDataFormatKind::Duration, None)
            .unwrap();
        assert_eq!(number.prefix[1], CELL_TYPE_DURATION);
        assert_eq!(number.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        assert_eq!(number.cell_format_kind(), Some(DURATION_CELL_FORMAT_KIND));
        assert_eq!(number.format_identifier(), Some(83));
        assert_eq!(number.secondary_format_identifier(), None);
        assert_eq!(number.stored_value(), StoredValue::Formula(61));
        assert_eq!(number.u32_field(STRING_FLAG), Some(67));
        assert_eq!(number.formula_error_identifier(), Some(71));
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(129_600.0)))
        );
        assert_eq!(number.style_identifier(), Some(73));
        assert_eq!(number.text_style_identifier(), Some(75));
        assert_eq!(number.conditional_style_identifier(), Some(77));
        assert_eq!(number.conditional_style_applied_rule(), Some(78));
        assert_eq!(number.comment_identifier(), Some(79));
        assert_eq!(number.tail, b"duration-formula-tail");
        assert_eq!(&number.prefix[2..6], original_prefix.as_slice());

        // The legacy generic compatibility setter canonicalizes a marker-5
        // source to primary-only. The focused Duration owner is the path that
        // retains this secondary; the generic host updates format references
        // independently and must not leave a dangling edge here.
        number
            .fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 89u32.to_le_bytes().to_vec());
        number.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_WITH_NUMBER_FORMAT.to_le_bytes());
        let original_value = value_fields(&number);
        number
            .set_data_format_identifier(97, CellDataFormatKind::Duration, None)
            .unwrap();
        assert_eq!(number.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        assert_eq!(number.secondary_format_identifier(), None);
        assert_eq!(number.stored_value(), StoredValue::Formula(61));
        assert_eq!(number.u32_field(STRING_FLAG), Some(67));
        assert_eq!(number.formula_error_identifier(), Some(71));
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Duration(finite(129_600.0)))
        );
        assert_eq!(value_fields(&number), original_value);
        assert_eq!(number.style_identifier(), Some(73));
        assert_eq!(number.text_style_identifier(), Some(75));
        assert_eq!(number.conditional_style_identifier(), Some(77));
        assert_eq!(number.conditional_style_applied_rule(), Some(78));
        assert_eq!(number.comment_identifier(), Some(79));
        assert_eq!(number.tail, b"duration-formula-tail");
        assert_eq!(&number.prefix[2..6], original_prefix.as_slice());

        number
            .set_data_format_identifier(101, CellDataFormatKind::NumberOrPercentage, None)
            .unwrap();
        assert_eq!(number.prefix[1], CELL_TYPE_NUMBER);
        assert_eq!(number.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(number.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(number.formula_error_identifier(), Some(71));
        assert_eq!(number.u32_field(STRING_FLAG), Some(67));
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(1.5)))
        );
    }

    #[test]
    fn focused_duration_format_identifier_preserves_all_value_and_tail_bytes() {
        let mut duration = BncCell::minimal();
        duration.prefix[2..6].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4]);
        duration.set_duration(-98_765.25).unwrap();
        duration.set_formula_reference(71);
        duration
            .fields
            .insert(STRING_FLAG, 73u32.to_le_bytes().to_vec());
        duration
            .fields
            .insert(FORMULA_ERROR_FLAG, 79u32.to_le_bytes().to_vec());
        duration.set_style_identifier(Some(83));
        duration.set_text_style_identifier(Some(89));
        duration.set_conditional_style(Some(97), Some(101));
        duration.set_comment_identifier(Some(103));
        duration.tail.extend_from_slice(b"duration-formula-tail");

        let original_non_format = non_format_encoding(&duration);
        let original_value_fields = value_fields(&duration);
        let original_cached_scalar = duration.cached_scalar().unwrap();
        assert_eq!(duration.prefix[1], CELL_TYPE_DURATION);
        assert_eq!(duration.stored_value(), StoredValue::Formula(71));
        assert_eq!(
            original_cached_scalar,
            Some(CachedScalar::Duration(finite(-98_765.25)))
        );
        assert!(duration.is_duration_format_compatible());
        assert!(duration.has_only_duration_format_metadata());

        duration
            .set_duration_format_identifier_preserving_value(Some(41))
            .unwrap();
        assert_eq!(duration.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        // Add the shared generic Number reference carried by native explicit
        // Duration cells. It is secondary metadata, not the Duration ID.
        duration
            .fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 19u32.to_le_bytes().to_vec());
        duration.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_WITH_NUMBER_FORMAT.to_le_bytes());
        assert_eq!(
            duration.explicit_format_flags(),
            EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(duration.cell_format_kind(), Some(DURATION_CELL_FORMAT_KIND));
        assert_eq!(duration.format_identifier(), Some(41));
        assert_eq!(duration.secondary_format_identifier(), Some(19));
        assert_eq!(duration.prefix[1], CELL_TYPE_DURATION);
        assert_eq!(duration.stored_value(), StoredValue::Formula(71));
        assert_eq!(duration.formula_error_identifier(), Some(79));
        assert_eq!(duration.cached_scalar().unwrap(), original_cached_scalar);
        assert_eq!(value_fields(&duration), original_value_fields);
        assert_eq!(non_format_encoding(&duration), original_non_format);

        let explicit = duration.encode();
        duration
            .set_duration_format_identifier_preserving_value(Some(41))
            .unwrap();
        assert_eq!(duration.encode(), explicit);

        duration
            .set_duration_format_identifier_preserving_value(Some(43))
            .unwrap();
        assert_eq!(duration.format_identifier(), Some(43));
        assert_eq!(duration.secondary_format_identifier(), Some(19));
        assert_eq!(duration.prefix[1], CELL_TYPE_DURATION);
        assert_eq!(duration.stored_value(), StoredValue::Formula(71));
        assert_eq!(duration.formula_error_identifier(), Some(79));
        assert_eq!(duration.cached_scalar().unwrap(), original_cached_scalar);
        assert_eq!(value_fields(&duration), original_value_fields);
        assert_eq!(non_format_encoding(&duration), original_non_format);

        duration
            .set_duration_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(duration.explicit_format_flags(), 0);
        assert_eq!(duration.cell_format_kind(), None);
        assert_eq!(duration.format_identifier(), None);
        assert_eq!(duration.secondary_format_identifier(), None);
        assert_eq!(duration.prefix[1], CELL_TYPE_DURATION);
        assert_eq!(duration.stored_value(), StoredValue::Formula(71));
        assert_eq!(duration.formula_error_identifier(), Some(79));
        assert_eq!(duration.cached_scalar().unwrap(), original_cached_scalar);
        assert_eq!(value_fields(&duration), original_value_fields);
        assert_eq!(non_format_encoding(&duration), original_non_format);
    }

    #[test]
    fn focused_duration_format_identifier_preserves_native_number_secondary() {
        let source = hex("050700000000050002300100000000000f6e99c1040000000100000009000000");
        let mut duration = BncCell::parse(&source).unwrap();
        let original_value_fields = value_fields(&duration);
        let original_cached_scalar = duration.cached_scalar().unwrap();

        duration
            .set_duration_format_identifier_preserving_value(Some(11))
            .unwrap();
        assert_eq!(
            duration.explicit_format_flags(),
            EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(duration.cell_format_kind(), Some(DURATION_CELL_FORMAT_KIND));
        assert_eq!(duration.format_identifier(), Some(11));
        assert_eq!(duration.secondary_format_identifier(), Some(1));
        assert_eq!(value_fields(&duration), original_value_fields);
        assert_eq!(duration.cached_scalar().unwrap(), original_cached_scalar);

        let rewritten = duration.encode();
        let reparsed = BncCell::parse(&rewritten).unwrap();
        assert_eq!(
            reparsed.explicit_format_flags(),
            EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(reparsed.format_identifier(), Some(11));
        assert_eq!(reparsed.secondary_format_identifier(), Some(1));
        assert_eq!(reparsed.cached_scalar().unwrap(), original_cached_scalar);

        duration
            .set_duration_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(duration.explicit_format_flags(), 0);
        assert_eq!(duration.cell_format_kind(), None);
        assert_eq!(duration.format_identifier(), None);
        assert_eq!(duration.secondary_format_identifier(), None);
        assert_eq!(value_fields(&duration), original_value_fields);
        assert_eq!(duration.cached_scalar().unwrap(), original_cached_scalar);
    }

    #[test]
    fn focused_duration_format_identifier_borrowed_rewrite_is_bounded_and_byte_preserving() {
        let mut duration = BncCell::minimal();
        duration.prefix[2..6].copy_from_slice(&[0x51, 0x62, 0x73, 0x84]);
        duration.set_duration(3_723.5).unwrap();
        duration.set_formula_reference(107);
        duration
            .fields
            .insert(STRING_FLAG, 109u32.to_le_bytes().to_vec());
        duration
            .fields
            .insert(FORMULA_ERROR_FLAG, 113u32.to_le_bytes().to_vec());
        duration.set_style_identifier(Some(127));
        duration.set_comment_identifier(Some(131));
        duration.tail.extend_from_slice(b"duration-borrowed-tail");
        duration
            .set_duration_format_identifier_preserving_value(Some(137))
            .unwrap();
        duration
            .fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 139u32.to_le_bytes().to_vec());
        duration.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_WITH_NUMBER_FORMAT.to_le_bytes());

        let source = duration.encode();
        let source_view = BncCellView::parse(&source).unwrap();
        assert!(source_view.is_duration_format_compatible());
        assert!(source_view.has_only_duration_format_metadata());
        assert_eq!(source_view.format_identifier(), Some(137));
        assert_eq!(source_view.secondary_format_identifier(), Some(139));

        let plan = source_view
            .plan_duration_format_identifier_rewrite(Some(149))
            .unwrap();
        let rewritten = source_view
            .rewrite_duration_format_identifier_with_limit(Some(149), usize::MAX)
            .unwrap();
        assert_eq!(plan.output_len(), Some(rewritten.len()));
        let rewritten_view = BncCellView::parse(&rewritten).unwrap();
        assert_eq!(
            rewritten_view.explicit_format_flags(),
            EXPLICIT_DURATION_WITH_NUMBER_FORMAT
        );
        assert_eq!(
            rewritten_view.cell_format_kind(),
            Some(DURATION_CELL_FORMAT_KIND)
        );
        assert_eq!(rewritten_view.format_identifier(), Some(149));
        assert_eq!(rewritten_view.secondary_format_identifier(), Some(139));
        assert_eq!(rewritten_view.stored_value(), source_view.stored_value());
        assert_eq!(rewritten_view.cached_scalar(), source_view.cached_scalar());
        assert_eq!(rewritten_view.formula_error_identifier(), Some(113));
        assert_eq!(rewritten_view.style_identifier(), Some(127));
        assert_eq!(rewritten_view.comment_identifier(), Some(131));
        assert_eq!(rewritten_view.opaque_tail(), b"duration-borrowed-tail");
        assert_eq!(&rewritten[..6], &source[..6]);
        assert_eq!(&rewritten[8..12], &source[8..12]);
        assert_eq!(
            non_format_encoding(&BncCell::parse(&rewritten).unwrap()),
            non_format_encoding(&duration)
        );

        let no_op = source_view
            .rewrite_duration_format_identifier_with_limit(Some(137), source.len())
            .unwrap();
        assert_eq!(no_op, source);

        let cleared_plan = rewritten_view
            .plan_duration_format_identifier_rewrite(None)
            .unwrap();
        let cleared = rewritten_view
            .rewrite_duration_format_identifier_with_limit(None, usize::MAX)
            .unwrap();
        assert_eq!(cleared_plan.output_len(), Some(cleared.len()));
        let cleared_view = BncCellView::parse(&cleared).unwrap();
        assert_eq!(cleared_view.explicit_format_flags(), 0);
        assert_eq!(cleared_view.cell_format_kind(), None);
        assert_eq!(cleared_view.format_identifier(), None);
        assert_eq!(cleared_view.secondary_format_identifier(), None);
        assert_eq!(cleared_view.stored_value(), source_view.stored_value());
        assert_eq!(cleared_view.cached_scalar(), source_view.cached_scalar());
        assert_eq!(cleared_view.opaque_tail(), source_view.opaque_tail());
        assert_eq!(
            non_format_encoding(&BncCell::parse(&cleared).unwrap()),
            non_format_encoding(&duration)
        );

        assert!(matches!(
            source_view.rewrite_duration_format_identifier_with_limit(Some(149), rewritten.len() - 1),
            Err(Error::OutputLimitExceeded { observed, maximum })
                if observed == rewritten.len() && maximum == rewritten.len() - 1
        ));
    }

    #[test]
    fn focused_duration_format_identifier_supports_empty_cells_and_exact_no_ops() {
        let mut empty = BncCell::minimal();
        empty.prefix[2..6].copy_from_slice(&[0x11, 0x22, 0x33, 0x44]);
        empty.set_style_identifier(Some(151));
        empty.set_comment_identifier(Some(157));
        empty.tail.extend_from_slice(b"duration-empty-tail");
        let original = empty.encode();
        let original_non_format = non_format_encoding(&empty);

        assert!(empty.is_duration_format_compatible());
        assert!(empty.has_only_duration_format_metadata());
        empty
            .set_duration_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(empty.encode(), original);

        empty
            .set_duration_format_identifier_preserving_value(Some(163))
            .unwrap();
        let explicit = empty.encode();
        assert_eq!(empty.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
        assert_eq!(empty.cell_format_kind(), Some(DURATION_CELL_FORMAT_KIND));
        assert_eq!(empty.format_identifier(), Some(163));
        assert_eq!(empty.stored_value(), StoredValue::Empty);
        assert_eq!(empty.cached_scalar().unwrap(), None);
        assert_eq!(non_format_encoding(&empty), original_non_format);

        empty
            .set_duration_format_identifier_preserving_value(Some(163))
            .unwrap();
        assert_eq!(empty.encode(), explicit);
        empty
            .set_duration_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(empty.encode(), original);
        assert_eq!(non_format_encoding(&empty), original_non_format);

        let source_view = BncCellView::parse(&original).unwrap();
        assert!(source_view.is_duration_format_compatible());
        let rewritten = source_view
            .rewrite_duration_format_identifier_with_limit(Some(167), usize::MAX)
            .unwrap();
        let rewritten_view = BncCellView::parse(&rewritten).unwrap();
        assert_eq!(rewritten_view.stored_value(), StoredValue::Empty);
        assert_eq!(rewritten_view.cached_scalar(), None);
        assert_eq!(rewritten_view.format_identifier(), Some(167));
        assert_eq!(
            non_format_encoding(&BncCell::parse(&rewritten).unwrap()),
            original_non_format
        );
    }

    #[test]
    fn focused_duration_format_identifier_rejects_automatic_and_cross_family_metadata_atomically() {
        let automatic_bytes = hex("050700000000000002100100000000000017ad400400000008000000");
        let mut automatic = BncCell::parse(&automatic_bytes).unwrap();
        assert!(automatic.is_duration_format_compatible());
        assert!(automatic.has_only_duration_format_metadata());
        let before = automatic.encode();
        for identifier in [None, Some(173)] {
            assert!(
                automatic
                    .set_duration_format_identifier_preserving_value(identifier)
                    .is_err()
            );
            assert_eq!(automatic.encode(), before);
        }
        let automatic_view = BncCellView::parse(&automatic_bytes).unwrap();
        assert!(automatic_view.is_duration_format_compatible());
        assert!(
            automatic_view
                .plan_duration_format_identifier_rewrite(Some(179))
                .is_err()
        );
        assert!(
            automatic_view
                .rewrite_duration_format_identifier_with_limit(None, usize::MAX)
                .is_err()
        );

        let mut wrong_family = BncCell::minimal();
        wrong_family.set_duration(7.0).unwrap();
        wrong_family.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DATE_TIME_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        wrong_family.fields.insert(
            DATE_TIME_FORMAT_IDENTIFIER_FLAG,
            181u32.to_le_bytes().to_vec(),
        );
        wrong_family.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DATE_TIME_FORMAT.to_le_bytes());
        assert!(!wrong_family.has_only_duration_format_metadata());
        let before = wrong_family.encode();
        assert!(
            wrong_family
                .set_duration_format_identifier_preserving_value(Some(187))
                .is_err()
        );
        assert_eq!(wrong_family.encode(), before);
        let wrong_family_view = BncCellView::parse(&before).unwrap();
        assert!(!wrong_family_view.has_only_duration_format_metadata());
        assert!(
            wrong_family_view
                .rewrite_duration_format_identifier_with_limit(Some(191), usize::MAX)
                .is_err()
        );

        let mut wrong_marker = BncCell::minimal();
        wrong_marker.set_duration(7.0).unwrap();
        wrong_marker.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        wrong_marker.fields.insert(
            DURATION_FORMAT_IDENTIFIER_FLAG,
            193u32.to_le_bytes().to_vec(),
        );
        wrong_marker.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DECIMAL_FORMAT.to_le_bytes());
        let before = wrong_marker.encode();
        assert!(
            wrong_marker
                .set_duration_format_identifier_preserving_value(None)
                .is_err()
        );
        assert_eq!(wrong_marker.encode(), before);

        let mut base_with_secondary = BncCell::minimal();
        base_with_secondary.set_duration(7.0).unwrap();
        base_with_secondary.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        base_with_secondary.fields.insert(
            DURATION_FORMAT_IDENTIFIER_FLAG,
            197u32.to_le_bytes().to_vec(),
        );
        base_with_secondary
            .fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 199u32.to_le_bytes().to_vec());
        base_with_secondary.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_FORMAT.to_le_bytes());
        let before = base_with_secondary.encode();
        assert!(
            base_with_secondary
                .set_duration_format_identifier_preserving_value(Some(211))
                .is_err()
        );
        assert_eq!(base_with_secondary.encode(), before);

        let mut with_number_without_secondary = BncCell::minimal();
        with_number_without_secondary.set_duration(7.0).unwrap();
        with_number_without_secondary.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        with_number_without_secondary.fields.insert(
            DURATION_FORMAT_IDENTIFIER_FLAG,
            223u32.to_le_bytes().to_vec(),
        );
        with_number_without_secondary.prefix
            [EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_WITH_NUMBER_FORMAT.to_le_bytes());
        let before = with_number_without_secondary.encode();
        assert!(
            with_number_without_secondary
                .set_duration_format_identifier_preserving_value(None)
                .is_err()
        );
        assert_eq!(with_number_without_secondary.encode(), before);
    }

    #[test]
    fn focused_duration_format_identifier_rejects_zero_malformed_and_ambiguous_shapes() {
        let mut zero_requested = BncCell::minimal();
        zero_requested.set_duration(7.0).unwrap();
        let before = zero_requested.encode();
        assert!(
            zero_requested
                .set_duration_format_identifier_preserving_value(Some(0))
                .is_err()
        );
        assert_eq!(zero_requested.encode(), before);

        let mut zero_existing = BncCell::minimal();
        zero_existing.set_duration(7.0).unwrap();
        zero_existing.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        zero_existing
            .fields
            .insert(DURATION_FORMAT_IDENTIFIER_FLAG, 0u32.to_le_bytes().to_vec());
        zero_existing.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_FORMAT.to_le_bytes());
        let before = zero_existing.encode();
        assert!(
            zero_existing
                .set_duration_format_identifier_preserving_value(None)
                .is_err()
        );
        assert_eq!(zero_existing.encode(), before);

        let mut zero_secondary = BncCell::minimal();
        zero_secondary.set_duration(7.0).unwrap();
        zero_secondary.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        zero_secondary.fields.insert(
            DURATION_FORMAT_IDENTIFIER_FLAG,
            197u32.to_le_bytes().to_vec(),
        );
        zero_secondary
            .fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 0u32.to_le_bytes().to_vec());
        zero_secondary.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_WITH_NUMBER_FORMAT.to_le_bytes());
        let before = zero_secondary.encode();
        assert!(
            zero_secondary
                .set_duration_format_identifier_preserving_value(Some(199))
                .is_err()
        );
        assert_eq!(zero_secondary.encode(), before);

        let mut malformed = BncCell::minimal();
        malformed.set_duration(7.0).unwrap();
        malformed.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DURATION_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        malformed
            .fields
            .insert(DURATION_FORMAT_IDENTIFIER_FLAG, vec![1, 2, 3]);
        malformed.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DURATION_FORMAT.to_le_bytes());
        let before = malformed.encode();
        assert!(
            malformed
                .set_duration_format_identifier_preserving_value(Some(211))
                .is_err()
        );
        assert_eq!(malformed.encode(), before);

        let mut decimal_cache = BncCell::minimal();
        decimal_cache.set_duration(7.0).unwrap();
        decimal_cache
            .fields
            .insert(DECIMAL_FLAG, decimal128_le(7.0).unwrap().to_vec());
        let before = decimal_cache.encode();
        assert!(!decimal_cache.is_duration_format_compatible());
        assert!(
            decimal_cache
                .set_duration_format_identifier_preserving_value(Some(223))
                .is_err()
        );
        assert_eq!(decimal_cache.encode(), before);

        let mut cache_without_formula = BncCell::minimal();
        cache_without_formula.set_duration(7.0).unwrap();
        cache_without_formula
            .fields
            .insert(STRING_FLAG, 227u32.to_le_bytes().to_vec());
        let before = cache_without_formula.encode();
        assert!(!cache_without_formula.is_duration_format_compatible());
        assert!(
            cache_without_formula
                .set_duration_format_identifier_preserving_value(Some(229))
                .is_err()
        );
        assert_eq!(cache_without_formula.encode(), before);

        let mut zero_formula = BncCell::minimal();
        zero_formula.set_duration(7.0).unwrap();
        zero_formula
            .fields
            .insert(FORMULA_FLAG, 0u32.to_le_bytes().to_vec());
        let before = zero_formula.encode();
        assert!(!zero_formula.is_duration_format_compatible());
        assert!(
            zero_formula
                .set_duration_format_identifier_preserving_value(Some(233))
                .is_err()
        );
        assert_eq!(zero_formula.encode(), before);

        let mut reserved = BncCell::minimal();
        reserved.set_duration(7.0).unwrap();
        reserved
            .fields
            .insert(RESERVED_KNOWN_FIELD_FLAG, 239u32.to_le_bytes().to_vec());
        let before = reserved.encode();
        assert!(
            reserved
                .set_duration_format_identifier_preserving_value(Some(241))
                .is_err()
        );
        assert_eq!(reserved.encode(), before);
    }

    #[test]
    fn checkbox_formats_match_native_boolean_metadata_and_conversion() {
        let native_checked =
            hex("050600000000200002140400000000000000f03f01000000060000000a000000");
        let checked = BncCell::parse(&native_checked).unwrap();
        assert_eq!(checked.stored_value(), StoredValue::Boolean);
        assert_eq!(
            checked.cached_scalar().unwrap(),
            Some(CachedScalar::Boolean(true))
        );
        assert_eq!(checked.explicit_format_flags(), EXPLICIT_CHECKBOX_FORMAT);
        assert_eq!(checked.cell_format_kind(), Some(CHECKBOX_CELL_FORMAT_KIND));
        assert_eq!(checked.format_identifier(), Some(10));
        assert_eq!(checked.secondary_format_identifier(), None);
        assert_eq!(checked.encode(), native_checked);

        let native_unchecked =
            hex("050600000000200002140400000000000000000001000000060000000a000000");
        let unchecked = BncCell::parse(&native_unchecked).unwrap();
        assert_eq!(
            unchecked.cached_scalar().unwrap(),
            Some(CachedScalar::Boolean(false))
        );
        assert_eq!(unchecked.encode(), native_unchecked);

        let mut empty = BncCell::minimal();
        empty
            .set_data_format_identifier(10, CellDataFormatKind::Checkbox, Some(1))
            .unwrap();
        assert_eq!(empty.encode(), native_unchecked);
        empty.clear_explicit_format();
        assert_eq!(empty.stored_value(), StoredValue::Boolean);
        assert_eq!(empty.explicit_format_flags(), 0);
        assert_eq!(empty.format_identifier(), None);

        let mut number = BncCell::minimal();
        number.set_number(1.0).unwrap();
        number
            .set_data_format_identifier(10, CellDataFormatKind::Checkbox, Some(1))
            .unwrap();
        assert_eq!(number.encode(), native_checked);
    }

    #[test]
    fn star_rating_formats_match_native_numeric_metadata_and_conversion() {
        let native_three =
            hex("0502000000000100013400000300000000000000000000000000403002000000010000000b000000");
        let three = BncCell::parse(&native_three).unwrap();
        assert_eq!(three.stored_value(), StoredValue::Number);
        assert_eq!(
            three.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(3.0)))
        );
        assert_eq!(three.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(three.cell_format_kind(), Some(STAR_RATING_CELL_FORMAT_KIND));
        assert_eq!(three.control_cell_spec_identifier(), Some(2));
        assert_eq!(three.format_identifier(), Some(11));
        assert_eq!(three.encode(), native_three);

        let mut empty = BncCell::minimal();
        empty
            .set_data_format_identifier(11, CellDataFormatKind::StarRating, Some(2))
            .unwrap();
        assert_eq!(
            empty.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(0.0)))
        );
        assert_eq!(empty.format_identifier(), Some(11));
        empty.clear_explicit_format();
        assert_eq!(empty.stored_value(), StoredValue::Number);
        assert_eq!(empty.control_cell_spec_identifier(), None);
    }

    #[test]
    fn display_format_changes_preserve_value_encoding_through_round_trips() {
        let mut number = BncCell::minimal();
        number.set_number(1.5).unwrap();
        number.set_formula_reference(12);
        let original_number = value_fields(&number);

        for (identifier, kind) in [
            (7, CellDataFormatKind::DateTime),
            (9, CellDataFormatKind::Duration),
            (4, CellDataFormatKind::Currency),
            (2, CellDataFormatKind::NumberOrPercentage),
        ] {
            number
                .set_data_format_metadata_identifier(identifier, kind, None)
                .unwrap();
            assert_eq!(value_fields(&number), original_number);
            assert_eq!(number.stored_value(), StoredValue::Formula(12));
            assert_eq!(
                number.cached_scalar().unwrap(),
                Some(CachedScalar::Number(finite(1.5)))
            );

            let reparsed = BncCell::parse(&number.encode()).unwrap();
            assert_eq!(value_fields(&reparsed), original_number);
            number = reparsed;
        }
        number.clear_explicit_format();
        assert_eq!(value_fields(&number), original_number);

        let mut date = BncCell::minimal();
        date.set_date(789_332_889.0).unwrap();
        let original_date = value_fields(&date);
        date.set_data_format_metadata_identifier(2, CellDataFormatKind::NumberOrPercentage, None)
            .unwrap();
        assert_eq!(value_fields(&date), original_date);
        assert_eq!(date.stored_value(), StoredValue::Date);

        let mut duration = BncCell::minimal();
        duration.set_duration(3_723.5).unwrap();
        let original_duration = value_fields(&duration);
        duration
            .set_data_format_metadata_identifier(7, CellDataFormatKind::DateTime, None)
            .unwrap();
        assert_eq!(value_fields(&duration), original_duration);
        assert_eq!(duration.stored_value(), StoredValue::Duration);
    }

    #[test]
    fn focused_decimal_format_identifier_preserves_every_non_format_byte() {
        let mut cell = BncCell::minimal();
        cell.set_number(42.25).unwrap();
        cell.set_formula_reference(17);
        cell.set_style_identifier(Some(23));
        cell.set_comment_identifier(Some(29));
        cell.tail.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        cell.set_data_format_identifier(31, CellDataFormatKind::NumericControlCurrency, Some(37))
            .unwrap();

        let original_prefix = cell.prefix[..EXPLICIT_FORMAT_FLAGS_START].to_vec();
        let original_fields = cell
            .fields
            .iter()
            .filter(|(flag, _)| FORMAT_METADATA_FLAGS & **flag == 0)
            .map(|(flag, value)| (*flag, value.clone()))
            .collect::<Vec<_>>();
        let original_tail = cell.tail.clone();
        let original_value = value_fields(&cell);

        cell.set_number_or_percentage_format_identifier_preserving_value(Some(41))
            .unwrap();
        assert_eq!(cell.prefix[..EXPLICIT_FORMAT_FLAGS_START], original_prefix);
        assert_eq!(cell.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(cell.format_identifier(), Some(41));
        assert_eq!(cell.control_cell_spec_identifier(), None);
        assert_eq!(value_fields(&cell), original_value);
        assert_eq!(cell.tail, original_tail);
        assert_eq!(
            cell.fields
                .iter()
                .filter(|(flag, _)| FORMAT_METADATA_FLAGS & **flag == 0)
                .map(|(flag, value)| (*flag, value.clone()))
                .collect::<Vec<_>>(),
            original_fields
        );

        let encoded = cell.try_encode_with_limit(usize::MAX).unwrap();
        let mut reparsed = BncCell::parse(&encoded).unwrap();
        reparsed
            .set_number_or_percentage_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(
            reparsed.prefix[..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix
        );
        assert_eq!(reparsed.explicit_format_flags(), 0);
        assert_eq!(reparsed.cell_format_kind(), None);
        assert_eq!(reparsed.format_identifier(), None);
        assert_eq!(value_fields(&reparsed), original_value);
        assert_eq!(reparsed.tail, original_tail);

        let before_rejection = reparsed.encode();
        assert!(
            reparsed
                .set_number_or_percentage_format_identifier_preserving_value(Some(0))
                .is_err()
        );
        assert_eq!(reparsed.encode(), before_rejection);
    }

    #[test]
    fn focused_currency_format_identifier_preserves_value_references_and_tail() {
        assert_eq!(EXPLICIT_CURRENCY_FORMAT, 0x0802);

        let mut cell = BncCell::minimal();
        cell.prefix[2..].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4, 0xe5, 0xf6]);
        cell.set_number(42.25).unwrap();
        cell.set_formula_reference(17);
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 19u32.to_le_bytes().to_vec());
        cell.set_style_identifier(Some(23));
        cell.set_comment_identifier(Some(29));
        cell.tail.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);
        assert_eq!(cell.numeric_cell_type(), Some(NumericCellType::Number));

        let original_prefix = cell.prefix;
        let original_value = value_fields(&cell);
        let original_cache = cell.cached_scalar().unwrap();
        let original_formula = cell.formula_identifier().unwrap();
        let original_formula_error = cell.formula_error_identifier();
        let original_style = cell.style_identifier();
        let original_comment = cell.comment_identifier();
        let original_tail = cell.tail.clone();

        cell.set_currency_format_identifier_preserving_value(Some(43))
            .unwrap();
        assert_eq!(cell.prefix[0], original_prefix[0]);
        assert_eq!(
            cell.numeric_cell_type(),
            Some(NumericCellType::AlternateNumber)
        );
        assert_eq!(
            cell.prefix[2..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix[2..EXPLICIT_FORMAT_FLAGS_START]
        );
        assert_eq!(
            cell.numeric_cell_type(),
            Some(NumericCellType::AlternateNumber)
        );
        assert_eq!(cell.explicit_format_flags(), EXPLICIT_CURRENCY_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(CURRENCY_CELL_FORMAT_KIND));
        assert_eq!(cell.format_identifier(), Some(43));
        assert_eq!(cell.secondary_format_identifier(), None);
        assert_eq!(cell.control_cell_spec_identifier(), None);
        assert_eq!(value_fields(&cell), original_value);
        assert_eq!(cell.cached_scalar().unwrap(), original_cache);
        assert_eq!(cell.formula_identifier().unwrap(), original_formula);
        assert_eq!(cell.formula_error_identifier(), original_formula_error);
        assert_eq!(cell.style_identifier(), original_style);
        assert_eq!(cell.comment_identifier(), original_comment);
        assert_eq!(cell.tail, original_tail);

        cell.fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 41u32.to_le_bytes().to_vec());
        cell.set_currency_format_identifier_preserving_value(Some(47))
            .unwrap();
        assert_eq!(
            cell.explicit_format_flags(),
            EXPLICIT_CURRENCY_WITH_NUMBER_FORMAT
        );
        assert_eq!(
            cell.numeric_cell_type(),
            Some(NumericCellType::AlternateNumber)
        );
        assert_eq!(cell.format_identifier(), Some(47));
        assert_eq!(cell.secondary_format_identifier(), Some(41));
        assert_eq!(value_fields(&cell), original_value);
        assert_eq!(cell.cached_scalar().unwrap(), original_cache);
        assert_eq!(cell.formula_identifier().unwrap(), original_formula);
        assert_eq!(cell.formula_error_identifier(), original_formula_error);
        assert_eq!(cell.style_identifier(), original_style);
        assert_eq!(cell.comment_identifier(), original_comment);
        assert_eq!(cell.tail, original_tail);

        let mut reparsed = BncCell::parse(&cell.encode()).unwrap();
        reparsed
            .set_currency_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(reparsed.prefix[0], original_prefix[0]);
        assert_eq!(reparsed.numeric_cell_type(), Some(NumericCellType::Number));
        assert_eq!(
            reparsed.prefix[2..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix[2..EXPLICIT_FORMAT_FLAGS_START]
        );
        assert_eq!(reparsed.explicit_format_flags(), 0);
        assert_eq!(reparsed.cell_format_kind(), None);
        assert_eq!(reparsed.format_identifier(), None);
        assert_eq!(reparsed.secondary_format_identifier(), None);
        assert_eq!(reparsed.control_cell_spec_identifier(), None);
        assert_eq!(value_fields(&reparsed), original_value);
        assert_eq!(reparsed.cached_scalar().unwrap(), original_cache);
        assert_eq!(reparsed.formula_identifier().unwrap(), original_formula);
        assert_eq!(reparsed.formula_error_identifier(), original_formula_error);
        assert_eq!(reparsed.style_identifier(), original_style);
        assert_eq!(reparsed.comment_identifier(), original_comment);
        assert_eq!(reparsed.tail, original_tail);

        let before_rejection = reparsed.encode();
        assert!(
            reparsed
                .set_currency_format_identifier_preserving_value(Some(0))
                .is_err()
        );
        assert_eq!(reparsed.encode(), before_rejection);
    }

    #[test]
    fn scientific_format_uses_decimal_wire_metadata_and_preserves_cell_bytes() {
        // Scientific is a distinct format-list family, but its BNC cell
        // metadata is exactly the shared decimal-family shape. Keep
        // the source deliberately rich so this transition cannot silently
        // drop a formula/cache, style, comment, or opaque tail.
        let mut cell = BncCell::minimal();
        cell.prefix[2..].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4, 0xe5, 0xf6]);
        cell.set_number(42.25).unwrap();
        cell.set_formula_reference(17);
        cell.fields
            .insert(FORMULA_ERROR_FLAG, 19u32.to_le_bytes().to_vec());
        cell.set_style_identifier(Some(23));
        cell.set_text_style_identifier(Some(29));
        cell.set_conditional_style(Some(31), Some(37));
        cell.set_comment_identifier(Some(41));
        cell.tail.extend_from_slice(&[0xde, 0xad, 0xbe, 0xef]);

        let original_prefix = cell.prefix;
        let original_value = value_fields(&cell);
        let original_cache = cell.cached_scalar().unwrap();
        let original_formula = cell.formula_identifier().unwrap();
        let original_formula_error = cell.formula_error_identifier();
        let original_style = cell.style_identifier();
        let original_text_style = cell.text_style_identifier();
        let original_conditional_style = cell.conditional_style_identifier();
        let original_conditional_rule = cell.conditional_style_applied_rule();
        let original_comment = cell.comment_identifier();
        let original_tail = cell.tail.clone();

        cell.set_number_or_percentage_format_identifier_preserving_value(Some(43))
            .unwrap();

        assert_eq!(cell.prefix[0], original_prefix[0]);
        assert_eq!(
            cell.prefix[2..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix[2..EXPLICIT_FORMAT_FLAGS_START]
        );
        assert_eq!(
            cell.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END],
            EXPLICIT_DECIMAL_FORMAT.to_le_bytes()
        );
        assert_eq!(cell.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(cell.format_identifier(), Some(43));
        assert_eq!(cell.secondary_format_identifier(), None);
        assert_eq!(cell.control_cell_spec_identifier(), None);
        assert!(cell.has_only_decimal_format_metadata());

        assert_eq!(cell.stored_value(), StoredValue::Formula(original_formula));
        assert_eq!(value_fields(&cell), original_value);
        assert_eq!(cell.cached_scalar().unwrap(), original_cache);
        assert_eq!(cell.formula_error_identifier(), original_formula_error);
        assert_eq!(cell.style_identifier(), original_style);
        assert_eq!(cell.text_style_identifier(), original_text_style);
        assert_eq!(
            cell.conditional_style_identifier(),
            original_conditional_style
        );
        assert_eq!(
            cell.conditional_style_applied_rule(),
            original_conditional_rule
        );
        assert_eq!(cell.comment_identifier(), original_comment);
        assert_eq!(cell.tail, original_tail);

        let reparsed = BncCell::parse(&cell.encode()).unwrap();
        assert_eq!(reparsed.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(reparsed.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(reparsed.format_identifier(), Some(43));
        assert_eq!(
            reparsed.stored_value(),
            StoredValue::Formula(original_formula)
        );
        assert_eq!(reparsed.cached_scalar().unwrap(), original_cache);
        assert_eq!(reparsed.formula_error_identifier(), original_formula_error);
        assert_eq!(reparsed.style_identifier(), original_style);
        assert_eq!(reparsed.text_style_identifier(), original_text_style);
        assert_eq!(reparsed.comment_identifier(), original_comment);
        assert_eq!(reparsed.tail, original_tail);
    }

    #[test]
    fn scientific_format_preserves_empty_cells_and_has_no_wire_family_marker() {
        let mut scientific = BncCell::minimal();
        scientific.prefix[2..].copy_from_slice(&[0x11, 0x22, 0x33, 0x44, 0x55, 0x66]);
        scientific.set_style_identifier(Some(7));
        scientific.set_comment_identifier(Some(11));
        scientific.tail.extend_from_slice(b"scientific-tail");
        let original_prefix = scientific.prefix;
        let original_tail = scientific.tail.clone();

        scientific
            .set_number_or_percentage_format_identifier_preserving_value(Some(13))
            .unwrap();
        assert_eq!(scientific.stored_value(), StoredValue::Empty);
        assert_eq!(scientific.cached_scalar().unwrap(), None);
        assert_eq!(scientific.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(
            scientific.cell_format_kind(),
            Some(DECIMAL_CELL_FORMAT_KIND)
        );
        assert_eq!(scientific.format_identifier(), Some(13));
        assert_eq!(scientific.style_identifier(), Some(7));
        assert_eq!(scientific.comment_identifier(), Some(11));
        assert_eq!(scientific.tail, original_tail);

        let mut number_or_percentage = BncCell::minimal();
        number_or_percentage.prefix[2..].copy_from_slice(&original_prefix[2..]);
        number_or_percentage.set_style_identifier(Some(7));
        number_or_percentage.set_comment_identifier(Some(11));
        number_or_percentage
            .tail
            .extend_from_slice(b"scientific-tail");
        number_or_percentage
            .set_number_or_percentage_format_identifier_preserving_value(Some(13))
            .unwrap();
        // The format-list payload, not the BNC cell, distinguishes Scientific
        // from Number and Percentage.  A wire-only reader must not infer a
        // specific family from this identical decimal cell shape.
        assert_eq!(scientific.encode(), number_or_percentage.encode());

        let encoded = scientific.encode();
        let mut reparsed = BncCell::parse(&encoded).unwrap();
        reparsed
            .set_number_or_percentage_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(reparsed.stored_value(), StoredValue::Empty);
        assert_eq!(reparsed.cached_scalar().unwrap(), None);
        assert_eq!(reparsed.explicit_format_flags(), 0);
        assert_eq!(reparsed.cell_format_kind(), None);
        assert_eq!(reparsed.format_identifier(), None);
        assert_eq!(reparsed.prefix[0], original_prefix[0]);
        assert_eq!(
            reparsed.prefix[2..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix[2..EXPLICIT_FORMAT_FLAGS_START]
        );
        assert_eq!(reparsed.style_identifier(), Some(7));
        assert_eq!(reparsed.comment_identifier(), Some(11));
        assert_eq!(reparsed.tail, original_tail);
    }

    #[test]
    fn fraction_format_uses_decimal_wire_metadata_and_preserves_cell_bytes() {
        // Fraction is a distinct format-list family, but its BNC cell
        // metadata is exactly the shared Number-or-Percentage decimal shape.
        // Keep the source deliberately rich so the metadata transition cannot
        // drop a formula/cache, style, comment, or opaque tail.
        let mut fraction = BncCell::minimal();
        fraction.prefix[2..].copy_from_slice(&[0x91, 0xa2, 0xb3, 0xc4, 0xd5, 0xe6]);
        fraction.set_number(-0.125).unwrap();
        fraction.set_formula_reference(71);
        fraction
            .fields
            .insert(FORMULA_ERROR_FLAG, 73u32.to_le_bytes().to_vec());
        fraction.set_style_identifier(Some(79));
        fraction.set_text_style_identifier(Some(83));
        fraction.set_conditional_style(Some(89), Some(97));
        fraction.set_comment_identifier(Some(101));
        fraction.tail.extend_from_slice(b"fraction-tail");

        let original_prefix = fraction.prefix;
        let original_value = value_fields(&fraction);
        let original_cache = fraction.cached_scalar().unwrap();
        let original_formula = fraction.formula_identifier().unwrap();
        let original_formula_error = fraction.formula_error_identifier();
        let original_style = fraction.style_identifier();
        let original_text_style = fraction.text_style_identifier();
        let original_conditional_style = fraction.conditional_style_identifier();
        let original_conditional_rule = fraction.conditional_style_applied_rule();
        let original_comment = fraction.comment_identifier();
        let original_tail = fraction.tail.clone();

        fraction
            .set_number_or_percentage_format_identifier_preserving_value(Some(107))
            .unwrap();

        assert_eq!(fraction.prefix[0], original_prefix[0]);
        assert_eq!(
            fraction.prefix[2..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix[2..EXPLICIT_FORMAT_FLAGS_START]
        );
        assert_eq!(
            fraction.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END],
            EXPLICIT_DECIMAL_FORMAT.to_le_bytes()
        );
        assert_eq!(fraction.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(fraction.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(fraction.format_identifier(), Some(107));
        assert_eq!(fraction.secondary_format_identifier(), None);
        assert_eq!(fraction.control_cell_spec_identifier(), None);
        assert!(fraction.has_only_decimal_format_metadata());

        assert_eq!(
            fraction.stored_value(),
            StoredValue::Formula(original_formula)
        );
        assert_eq!(value_fields(&fraction), original_value);
        assert_eq!(fraction.cached_scalar().unwrap(), original_cache);
        assert_eq!(fraction.formula_error_identifier(), original_formula_error);
        assert_eq!(fraction.style_identifier(), original_style);
        assert_eq!(fraction.text_style_identifier(), original_text_style);
        assert_eq!(
            fraction.conditional_style_identifier(),
            original_conditional_style
        );
        assert_eq!(
            fraction.conditional_style_applied_rule(),
            original_conditional_rule
        );
        assert_eq!(fraction.comment_identifier(), original_comment);
        assert_eq!(fraction.tail, original_tail);

        let mut cleared = BncCell::parse(&fraction.encode()).unwrap();
        cleared
            .set_number_or_percentage_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(
            cleared.stored_value(),
            StoredValue::Formula(original_formula)
        );
        assert_eq!(cleared.cached_scalar().unwrap(), original_cache);
        assert_eq!(cleared.explicit_format_flags(), 0);
        assert_eq!(cleared.cell_format_kind(), None);
        assert_eq!(cleared.format_identifier(), None);
        assert_eq!(cleared.formula_error_identifier(), original_formula_error);
        assert_eq!(cleared.style_identifier(), original_style);
        assert_eq!(cleared.text_style_identifier(), original_text_style);
        assert_eq!(
            cleared.conditional_style_identifier(),
            original_conditional_style
        );
        assert_eq!(
            cleared.conditional_style_applied_rule(),
            original_conditional_rule
        );
        assert_eq!(cleared.comment_identifier(), original_comment);
        assert_eq!(cleared.tail, original_tail);
        assert_eq!(value_fields(&cleared), original_value);
    }

    #[test]
    fn fraction_format_preserves_empty_cells_and_has_no_wire_family_marker() {
        let mut fraction = BncCell::minimal();
        fraction.prefix[2..].copy_from_slice(&[0x21, 0x32, 0x43, 0x54, 0x65, 0x76]);
        fraction.set_style_identifier(Some(109));
        fraction.set_comment_identifier(Some(113));
        fraction.tail.extend_from_slice(b"empty-fraction-tail");
        let original_prefix = fraction.prefix;
        let original_tail = fraction.tail.clone();

        fraction
            .set_number_or_percentage_format_identifier_preserving_value(Some(127))
            .unwrap();
        assert_eq!(fraction.stored_value(), StoredValue::Empty);
        assert_eq!(fraction.cached_scalar().unwrap(), None);
        assert_eq!(fraction.numeric_cell_type(), None);
        assert_eq!(fraction.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(fraction.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(fraction.format_identifier(), Some(127));
        assert_eq!(fraction.style_identifier(), Some(109));
        assert_eq!(fraction.comment_identifier(), Some(113));
        assert_eq!(fraction.tail, original_tail);
        assert!(fraction.has_only_decimal_format_metadata());

        let mut number = BncCell::minimal();
        number.prefix[2..].copy_from_slice(&original_prefix[2..]);
        number.set_style_identifier(Some(109));
        number.set_comment_identifier(Some(113));
        number.tail.extend_from_slice(&original_tail);
        number
            .set_number_or_percentage_format_identifier_preserving_value(Some(127))
            .unwrap();

        let mut percentage = BncCell::minimal();
        percentage.prefix[2..].copy_from_slice(&original_prefix[2..]);
        percentage.set_style_identifier(Some(109));
        percentage.set_comment_identifier(Some(113));
        percentage.tail.extend_from_slice(&original_tail);
        percentage
            .set_number_or_percentage_format_identifier_preserving_value(Some(127))
            .unwrap();

        let mut scientific = BncCell::minimal();
        scientific.prefix[2..].copy_from_slice(&original_prefix[2..]);
        scientific.set_style_identifier(Some(109));
        scientific.set_comment_identifier(Some(113));
        scientific.tail.extend_from_slice(&original_tail);
        scientific
            .set_number_or_percentage_format_identifier_preserving_value(Some(127))
            .unwrap();

        // The format-list payload, not the BNC cell, distinguishes Fraction
        // from Number, Percentage, and Scientific. A wire-only reader must
        // not infer a specific family from this identical decimal shape.
        assert_eq!(fraction.encode(), number.encode());
        assert_eq!(fraction.encode(), percentage.encode());
        assert_eq!(fraction.encode(), scientific.encode());

        let encoded = fraction.encode();
        let mut cleared = BncCell::parse(&encoded).unwrap();
        cleared
            .set_number_or_percentage_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(cleared.stored_value(), StoredValue::Empty);
        assert_eq!(cleared.cached_scalar().unwrap(), None);
        assert_eq!(cleared.numeric_cell_type(), None);
        assert_eq!(cleared.explicit_format_flags(), 0);
        assert_eq!(cleared.cell_format_kind(), None);
        assert_eq!(cleared.format_identifier(), None);
        assert_eq!(cleared.prefix[0], original_prefix[0]);
        assert_eq!(
            cleared.prefix[2..EXPLICIT_FORMAT_FLAGS_START],
            original_prefix[2..EXPLICIT_FORMAT_FLAGS_START]
        );
        assert_eq!(cleared.style_identifier(), Some(109));
        assert_eq!(cleared.comment_identifier(), Some(113));
        assert_eq!(cleared.tail, original_tail);
    }

    #[test]
    fn slider_formats_match_native_number_and_currency_metadata() {
        let native_number =
            hex("0502000000000100013400001900000000000000000000000000403004000000010000000c000000");
        let number = BncCell::parse(&native_number).unwrap();
        assert_eq!(number.stored_value(), StoredValue::Number);
        assert_eq!(
            number.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(25.0)))
        );
        assert_eq!(number.explicit_format_flags(), EXPLICIT_DECIMAL_FORMAT);
        assert_eq!(number.cell_format_kind(), Some(DECIMAL_CELL_FORMAT_KIND));
        assert_eq!(number.control_cell_spec_identifier(), Some(4));
        assert_eq!(number.format_identifier(), Some(12));
        assert_eq!(number.secondary_format_identifier(), None);
        assert_eq!(number.encode(), native_number);

        let native_currency = hex(
            "050a000000000308017400001900000000000000000000000000403004000000020000000c0000000b000000",
        );
        let currency = BncCell::parse(&native_currency).unwrap();
        assert_eq!(currency.stored_value(), StoredValue::Number);
        assert_eq!(
            currency.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(25.0)))
        );
        assert_eq!(
            currency.explicit_format_flags(),
            EXPLICIT_CURRENCY_WITH_NUMBER_FORMAT
        );
        assert_eq!(currency.cell_format_kind(), Some(CURRENCY_CELL_FORMAT_KIND));
        assert_eq!(currency.control_cell_spec_identifier(), Some(4));
        assert_eq!(currency.format_identifier(), Some(11));
        assert_eq!(currency.secondary_format_identifier(), Some(12));
        assert!(!currency.has_only_decimal_format_metadata());
        assert_eq!(currency.encode(), native_currency);

        let mut empty = BncCell::minimal();
        empty.set_plain_number(10.0).unwrap();
        empty
            .set_data_format_identifier(
                12,
                CellDataFormatKind::NumericControlNumberOrPercentage,
                Some(4),
            )
            .unwrap();
        assert_eq!(
            empty.cached_scalar().unwrap(),
            Some(CachedScalar::Number(finite(10.0)))
        );
        assert_eq!(empty.control_cell_spec_identifier(), Some(4));
        empty
            .set_data_format_identifier(11, CellDataFormatKind::NumericControlCurrency, Some(4))
            .unwrap();
        assert_eq!(empty.cell_format_kind(), Some(CURRENCY_CELL_FORMAT_KIND));
        assert_eq!(empty.format_identifier(), Some(11));
        empty.clear_explicit_format();
        assert_eq!(empty.stored_value(), StoredValue::Number);
        assert_eq!(empty.control_cell_spec_identifier(), None);
    }

    #[test]
    fn pop_up_menu_format_matches_native_text_metadata() {
        let native = hex("0503000000008000081402000100000006000000050000000d000000");
        let cell = BncCell::parse(&native).unwrap();
        assert_eq!(cell.explicit_format_flags(), EXPLICIT_TEXT_FORMAT);
        assert_eq!(cell.cell_format_kind(), Some(TEXT_CELL_FORMAT_KIND));
        assert_eq!(cell.control_cell_spec_identifier(), Some(6));
        assert_eq!(cell.format_identifier(), Some(13));
        assert_eq!(cell.stored_value(), StoredValue::Text(1));
        assert_eq!(cell.encode(), native);

        let native_empty = hex("050000000000800000100200050000000c000000");
        let empty = BncCell::parse(&native_empty).unwrap();
        assert_eq!(empty.stored_value(), StoredValue::Empty);
        assert_eq!(empty.explicit_format_flags(), EXPLICIT_TEXT_FORMAT);
        assert_eq!(empty.cell_format_kind(), Some(TEXT_CELL_FORMAT_KIND));
        assert_eq!(empty.control_cell_spec_identifier(), None);
        assert_eq!(empty.format_identifier(), Some(12));
        assert_eq!(empty.encode(), native_empty);

        let native_converted = hex("0503000000008100083002000200000005000000010000000c000000");
        let converted = BncCell::parse(&native_converted).unwrap();
        assert_eq!(converted.stored_value(), StoredValue::Text(2));
        assert_eq!(
            converted.explicit_format_flags(),
            EXPLICIT_CONVERTED_TEXT_FORMAT
        );
        assert_eq!(converted.cell_format_kind(), Some(TEXT_CELL_FORMAT_KIND));
        assert_eq!(converted.control_cell_spec_identifier(), None);
        assert_eq!(converted.format_identifier(), Some(12));
        assert_eq!(converted.encode(), native_converted);

        let mut created = BncCell::minimal();
        created
            .set_data_format_identifier(12, CellDataFormatKind::Text, None)
            .unwrap();
        assert_eq!(created.stored_value(), StoredValue::Empty);
        assert_eq!(created.explicit_format_flags(), EXPLICIT_TEXT_FORMAT);
        assert_eq!(created.cell_format_kind(), Some(TEXT_CELL_FORMAT_KIND));
        assert_eq!(created.format_identifier(), Some(12));
        assert_eq!(created.control_cell_spec_identifier(), None);
    }

    #[test]
    fn date_time_metadata_supports_empty_cells_and_exact_no_ops() {
        let mut empty = BncCell::minimal();
        empty.prefix[2..6].copy_from_slice(&[0x11, 0x22, 0x33, 0x44]);
        empty.set_style_identifier(Some(7));
        empty.set_comment_identifier(Some(11));
        empty.tail.extend_from_slice(b"date-time-empty-tail");
        let original = empty.encode();
        let original_non_format = non_format_encoding(&empty);

        assert!(empty.is_date_time_format_compatible());
        empty
            .set_date_time_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(empty.encode(), original);

        empty
            .set_date_time_format_identifier_preserving_value(Some(13))
            .unwrap();
        let explicit = empty.encode();
        assert_eq!(empty.explicit_format_flags(), EXPLICIT_DATE_TIME_FORMAT);
        assert_eq!(empty.cell_format_kind(), Some(DATE_TIME_CELL_FORMAT_KIND));
        assert_eq!(empty.format_identifier(), Some(13));
        assert_eq!(empty.stored_value(), StoredValue::Empty);
        assert_eq!(empty.cached_scalar().unwrap(), None);
        assert_eq!(non_format_encoding(&empty), original_non_format);

        empty
            .set_date_time_format_identifier_preserving_value(Some(13))
            .unwrap();
        assert_eq!(empty.encode(), explicit);

        empty
            .set_date_time_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(empty.encode(), original);
        assert_eq!(non_format_encoding(&empty), original_non_format);
    }

    #[test]
    fn date_time_metadata_preserves_type_five_date_value_and_tail() {
        let mut date = BncCell::minimal();
        date.prefix[2..6].copy_from_slice(&[0xa1, 0xb2, 0xc3, 0xd4]);
        date.set_date(789_332_889.25).unwrap();
        date.set_style_identifier(Some(17));
        date.set_text_style_identifier(Some(19));
        date.set_conditional_style(Some(23), Some(29));
        date.set_comment_identifier(Some(31));
        date.tail.extend_from_slice(b"date-time-date-tail");
        let original_non_format = non_format_encoding(&date);
        let original_value = value_fields(&date);
        let original_cache = date.cached_scalar().unwrap();

        assert_eq!(date.prefix[1], CELL_TYPE_DATE);
        assert!(date.is_date_time_format_compatible());
        date.set_date_time_format_identifier_preserving_value(Some(41))
            .unwrap();
        assert_eq!(date.prefix[1], CELL_TYPE_DATE);
        assert_eq!(date.explicit_format_flags(), EXPLICIT_DATE_TIME_FORMAT);
        assert_eq!(date.cell_format_kind(), Some(DATE_TIME_CELL_FORMAT_KIND));
        assert_eq!(date.format_identifier(), Some(41));
        assert_eq!(date.stored_value(), StoredValue::Date);
        assert_eq!(date.cached_scalar().unwrap(), original_cache);
        assert_eq!(value_fields(&date), original_value);
        assert_eq!(non_format_encoding(&date), original_non_format);

        let explicit = date.encode();
        date.set_date_time_format_identifier_preserving_value(Some(41))
            .unwrap();
        assert_eq!(date.encode(), explicit);
        date.set_date_time_format_identifier_preserving_value(Some(43))
            .unwrap();
        assert_eq!(date.format_identifier(), Some(43));
        assert_eq!(date.prefix[1], CELL_TYPE_DATE);
        assert_eq!(date.cached_scalar().unwrap(), original_cache);
        assert_eq!(non_format_encoding(&date), original_non_format);

        date.set_date_time_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(date.explicit_format_flags(), 0);
        assert_eq!(date.cell_format_kind(), None);
        assert_eq!(date.format_identifier(), None);
        assert_eq!(date.prefix[1], CELL_TYPE_DATE);
        assert_eq!(date.stored_value(), StoredValue::Date);
        assert_eq!(date.cached_scalar().unwrap(), original_cache);
        assert_eq!(value_fields(&date), original_value);
        assert_eq!(non_format_encoding(&date), original_non_format);
    }

    #[test]
    fn date_time_metadata_preserves_type_nine_number_and_formula_cache() {
        let mut number = BncCell::minimal();
        number.prefix[2..6].copy_from_slice(&[0x91, 0xa2, 0xb3, 0xc4]);
        number.set_number(-98.25).unwrap();
        number.prefix[1] = CELL_TYPE_RICH_TEXT_OR_NUMBER;
        number.set_style_identifier(Some(47));
        number.set_comment_identifier(Some(53));
        number.tail.extend_from_slice(b"date-time-number-tail");
        let original_number_non_format = non_format_encoding(&number);
        let original_number_value = value_fields(&number);
        let original_number_cache = number.cached_scalar().unwrap();

        assert_eq!(number.stored_value(), StoredValue::Number);
        assert!(number.is_date_time_format_compatible());
        number
            .set_date_time_format_identifier_preserving_value(Some(59))
            .unwrap();
        assert_eq!(number.prefix[1], CELL_TYPE_RICH_TEXT_OR_NUMBER);
        assert_eq!(number.stored_value(), StoredValue::Number);
        assert_eq!(number.cached_scalar().unwrap(), original_number_cache);
        assert_eq!(value_fields(&number), original_number_value);
        assert_eq!(non_format_encoding(&number), original_number_non_format);

        let mut formula = BncCell::minimal();
        formula.prefix[2..6].copy_from_slice(&[0x61, 0x72, 0x83, 0x94]);
        formula.set_number(12.5).unwrap();
        formula.prefix[1] = CELL_TYPE_RICH_TEXT_OR_NUMBER;
        formula.set_formula_reference(71);
        formula
            .fields
            .insert(STRING_FLAG, 73u32.to_le_bytes().to_vec());
        formula
            .fields
            .insert(FORMULA_ERROR_FLAG, 79u32.to_le_bytes().to_vec());
        formula.set_style_identifier(Some(83));
        formula.set_comment_identifier(Some(89));
        formula.tail.extend_from_slice(b"date-time-formula-tail");
        let original_formula_non_format = non_format_encoding(&formula);
        let original_formula_value = value_fields(&formula);
        let original_formula_cache = formula.cached_scalar().unwrap();

        assert_eq!(formula.stored_value(), StoredValue::Formula(71));
        assert_eq!(
            original_formula_cache,
            Some(CachedScalar::Number(finite(12.5)))
        );
        assert!(formula.is_date_time_format_compatible());
        formula
            .set_date_time_format_identifier_preserving_value(Some(97))
            .unwrap();
        assert_eq!(formula.prefix[1], CELL_TYPE_RICH_TEXT_OR_NUMBER);
        assert_eq!(formula.stored_value(), StoredValue::Formula(71));
        assert_eq!(formula.formula_error_identifier(), Some(79));
        assert_eq!(formula.cached_scalar().unwrap(), original_formula_cache);
        assert_eq!(value_fields(&formula), original_formula_value);
        assert_eq!(non_format_encoding(&formula), original_formula_non_format);

        formula
            .set_date_time_format_identifier_preserving_value(None)
            .unwrap();
        assert_eq!(formula.explicit_format_flags(), 0);
        assert_eq!(formula.cell_format_kind(), None);
        assert_eq!(formula.format_identifier(), None);
        assert_eq!(formula.prefix[1], CELL_TYPE_RICH_TEXT_OR_NUMBER);
        assert_eq!(formula.stored_value(), StoredValue::Formula(71));
        assert_eq!(formula.formula_error_identifier(), Some(79));
        assert_eq!(formula.cached_scalar().unwrap(), original_formula_cache);
        assert_eq!(value_fields(&formula), original_formula_value);
        assert_eq!(non_format_encoding(&formula), original_formula_non_format);
    }

    #[test]
    fn date_time_metadata_rejects_zero_ids_wrong_families_and_ambiguous_shapes() {
        let mut dirty_empty = BncCell::minimal();
        dirty_empty
            .fields
            .insert(DATE_FLAG, 7.0f64.to_le_bytes().to_vec());
        assert!(!dirty_empty.is_date_time_format_compatible());
        let before = dirty_empty.encode();
        assert!(
            dirty_empty
                .set_date_time_format_identifier_preserving_value(Some(3))
                .is_err()
        );
        assert_eq!(dirty_empty.encode(), before);

        let mut zero = BncCell::minimal();
        zero.set_date(7.0).unwrap();
        let before = zero.encode();
        assert!(
            zero.set_date_time_format_identifier_preserving_value(Some(0))
                .is_err()
        );
        assert_eq!(zero.encode(), before);

        let mut ordinary_number = BncCell::minimal();
        ordinary_number.set_number(7.0).unwrap();
        let before = ordinary_number.encode();
        assert!(!ordinary_number.is_date_time_format_compatible());
        assert!(
            ordinary_number
                .set_date_time_format_identifier_preserving_value(Some(11))
                .is_err()
        );
        assert_eq!(ordinary_number.encode(), before);

        let mut wrong_family = BncCell::minimal();
        wrong_family.set_date(7.0).unwrap();
        wrong_family.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DECIMAL_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        wrong_family
            .fields
            .insert(CELL_FORMAT_IDENTIFIER_FLAG, 13u32.to_le_bytes().to_vec());
        wrong_family.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DECIMAL_FORMAT.to_le_bytes());
        let before = wrong_family.encode();
        assert!(
            wrong_family
                .set_date_time_format_identifier_preserving_value(Some(17))
                .is_err()
        );
        assert_eq!(wrong_family.encode(), before);

        let mut control = BncCell::minimal();
        control.set_date(7.0).unwrap();
        control
            .fields
            .insert(CONTROL_CELL_SPEC_FLAG, 19u32.to_le_bytes().to_vec());
        let before = control.encode();
        assert!(
            control
                .set_date_time_format_identifier_preserving_value(Some(23))
                .is_err()
        );
        assert_eq!(control.encode(), before);

        let mut automatic = BncCell::minimal();
        automatic.set_date(7.0).unwrap();
        automatic.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DATE_TIME_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        automatic.fields.insert(
            DATE_TIME_FORMAT_IDENTIFIER_FLAG,
            29u32.to_le_bytes().to_vec(),
        );
        let before = automatic.encode();
        assert!(
            automatic
                .set_date_time_format_identifier_preserving_value(None)
                .is_err()
        );
        assert_eq!(automatic.encode(), before);

        let mut incomplete = BncCell::minimal();
        incomplete.set_date(7.0).unwrap();
        incomplete.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DATE_TIME_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        incomplete.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DATE_TIME_FORMAT.to_le_bytes());
        let before = incomplete.encode();
        assert!(
            incomplete
                .set_date_time_format_identifier_preserving_value(Some(31))
                .is_err()
        );
        assert_eq!(incomplete.encode(), before);

        let mut ambiguous_value = BncCell::minimal();
        ambiguous_value.set_number(7.0).unwrap();
        ambiguous_value.prefix[1] = CELL_TYPE_RICH_TEXT_OR_NUMBER;
        ambiguous_value
            .fields
            .insert(NUMBER_FLAG, 7.0f64.to_le_bytes().to_vec());
        let before = ambiguous_value.encode();
        assert!(!ambiguous_value.is_date_time_format_compatible());
        assert!(
            ambiguous_value
                .set_date_time_format_identifier_preserving_value(Some(37))
                .is_err()
        );
        assert_eq!(ambiguous_value.encode(), before);

        let mut reserved = BncCell::minimal();
        reserved.set_date(7.0).unwrap();
        reserved
            .fields
            .insert(RESERVED_KNOWN_FIELD_FLAG, 41u32.to_le_bytes().to_vec());
        let before = reserved.encode();
        assert!(
            reserved
                .set_date_time_format_identifier_preserving_value(Some(43))
                .is_err()
        );
        assert_eq!(reserved.encode(), before);
    }

    #[test]
    fn date_time_metadata_rejects_malformed_existing_explicit_state_atomically() {
        let mut wrong_marker = BncCell::minimal();
        wrong_marker.set_date(7.0).unwrap();
        wrong_marker.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DATE_TIME_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        wrong_marker.fields.insert(
            DATE_TIME_FORMAT_IDENTIFIER_FLAG,
            47u32.to_le_bytes().to_vec(),
        );
        wrong_marker.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DECIMAL_FORMAT.to_le_bytes());
        let before = wrong_marker.encode();
        assert!(
            wrong_marker
                .set_date_time_format_identifier_preserving_value(None)
                .is_err()
        );
        assert_eq!(wrong_marker.encode(), before);

        let mut zero_existing = BncCell::minimal();
        zero_existing.set_date(7.0).unwrap();
        zero_existing.fields.insert(
            CELL_FORMAT_KIND_FLAG,
            DATE_TIME_CELL_FORMAT_KIND.to_le_bytes().to_vec(),
        );
        zero_existing.fields.insert(
            DATE_TIME_FORMAT_IDENTIFIER_FLAG,
            0u32.to_le_bytes().to_vec(),
        );
        zero_existing.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END]
            .copy_from_slice(&EXPLICIT_DATE_TIME_FORMAT.to_le_bytes());
        let before = zero_existing.encode();
        assert!(
            zero_existing
                .set_date_time_format_identifier_preserving_value(None)
                .is_err()
        );
        assert_eq!(zero_existing.encode(), before);
    }

    fn non_format_encoding(cell: &BncCell) -> Vec<u8> {
        let mut non_format = cell.clone();
        non_format.prefix[EXPLICIT_FORMAT_FLAGS_START..EXPLICIT_FORMAT_FLAGS_END].fill(0);
        non_format
            .fields
            .retain(|field, _| FORMAT_METADATA_FLAGS & field == 0);
        non_format.encode()
    }

    fn hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }
}
