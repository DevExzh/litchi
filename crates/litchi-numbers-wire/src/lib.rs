//! Low-level Binary Numbers Cell (BNC) value storage for iWork adapters.
//!
//! This adapter crate keeps the byte-preserving codec shared while the legacy
//! IWA host is migrated into the standalone Numbers owner. It is intentionally
//! excluded from the `litchi` and `litchi-numbers` facades. Applications
//! should use `litchi-numbers`; direct use of this crate opts into unstable
//! native-storage details rather than the supported semantic API.

#![forbid(unsafe_code)]

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
const EXPLICIT_FORMAT_FLAGS_START: usize = 6;
const EXPLICIT_FORMAT_FLAGS_END: usize = 8;
pub const EXPLICIT_DECIMAL_FORMAT: u16 = 1;
pub const EXPLICIT_CURRENCY_FORMAT: u16 = 0x0803;
pub const EXPLICIT_DATE_TIME_FORMAT: u16 = 0x0008;
pub const EXPLICIT_DURATION_FORMAT: u16 = 0x0005;
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

const VALUE_FLAGS: u32 = DECIMAL_FLAG
    | NUMBER_FLAG
    | DATE_FLAG
    | STRING_FLAG
    | RICH_TEXT_FLAG
    | FORMULA_FLAG
    | FORMULA_ERROR_FLAG;

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
const FIELD_COUNT: usize = FIELD_LAYOUT.len();

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    InvalidFormat(String),
    ParseError(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidFormat(message) | Self::ParseError(message) => {
                formatter.write_str(message)
            },
        }
    }
}

impl std::error::Error for Error {}

/// Result type for BNC decoding and mutation.
pub type Result<T> = std::result::Result<T, Error>;

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

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CachedScalar {
    Number(FiniteF64),
    Boolean(bool),
    Date(FiniteF64),
    Duration(FiniteF64),
    Unsupported(u8),
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
        self.set_number(value)?;
        self.fields
            .insert(FORMULA_FLAG, formula.to_le_bytes().to_vec());
        Ok(())
    }

    /// Replaces a formula's cached result with a boolean.
    ///
    /// # Errors
    ///
    /// Returns an error when the cell has no formula.
    pub fn set_formula_cached_boolean(&mut self, value: bool) -> Result<()> {
        let formula = self.formula_identifier()?;
        self.set_boolean(value);
        self.fields
            .insert(FORMULA_FLAG, formula.to_le_bytes().to_vec());
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

    /// Applies a Numbers data format and its identifier to the cell.
    ///
    /// # Errors
    ///
    /// Returns an error when an interactive format has no control-cell
    /// identifier, when a non-interactive format has one, or when a text
    /// format would discard a non-text scalar, or a required scalar conversion
    /// is not finite. This compatibility entry point performs the same native
    /// scalar conversions as Numbers; metadata-only application is kept as a
    /// separate internal operation so extraction never infers semantic value
    /// from display metadata.
    pub fn set_data_format_identifier(
        &mut self,
        identifier: u32,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) -> Result<()> {
        self.validate_data_format_request(kind, control_identifier)?;
        self.convert_scalar_for_data_format(kind)?;
        self.set_data_format_metadata_identifier(identifier, kind, control_identifier)
    }

    fn set_data_format_metadata_identifier(
        &mut self,
        identifier: u32,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) -> Result<()> {
        self.validate_data_format_request(kind, control_identifier)?;
        self.apply_data_format_identifier(identifier, kind, control_identifier);
        Ok(())
    }

    fn validate_data_format_request(
        &self,
        kind: CellDataFormatKind,
        control_identifier: Option<u32>,
    ) -> Result<()> {
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
                (EXPLICIT_CURRENCY_FORMAT, CURRENCY_CELL_FORMAT_KIND)
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
        let formula_identifier = self.u32_field(FORMULA_FLAG);
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
        if let Some(identifier) = formula_identifier {
            self.fields
                .insert(FORMULA_FLAG, identifier.to_le_bytes().to_vec());
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
    let spelling = if magnitude == 0.0 {
        "0".to_owned()
    } else {
        magnitude.to_string()
    };
    let (mantissa, explicit_exponent) = spelling
        .split_once(['e', 'E'])
        .map_or((spelling.as_str(), 0), |(mantissa, exponent)| {
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
    let mut digit_bytes = mantissa
        .bytes()
        .filter(|byte| *byte != b'.')
        .collect::<Vec<_>>();
    while digit_bytes.len() > 1 && digit_bytes.first() == Some(&b'0') {
        digit_bytes.remove(0);
    }
    let mut trailing_zeroes = 0i32;
    while digit_bytes.len() > 1 && digit_bytes.last() == Some(&b'0') {
        digit_bytes.pop();
        trailing_zeroes += 1;
    }
    let digit_text = std::str::from_utf8(&digit_bytes).map_err(|_error| {
        Error::ParseError(format!("Could not encode Numbers decimal {spelling:?}"))
    })?;
    let coefficient = digit_text.parse::<u128>().map_err(|_error| {
        Error::ParseError(format!("Could not encode Numbers decimal {spelling:?}"))
    })?;
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
        assert_eq!(converted.explicit_format_flags(), EXPLICIT_DURATION_FORMAT);
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
        assert_eq!(currency.explicit_format_flags(), EXPLICIT_CURRENCY_FORMAT);
        assert_eq!(currency.cell_format_kind(), Some(CURRENCY_CELL_FORMAT_KIND));
        assert_eq!(currency.control_cell_spec_identifier(), Some(4));
        assert_eq!(currency.format_identifier(), Some(11));
        assert_eq!(currency.secondary_format_identifier(), Some(12));
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

    fn hex(value: &str) -> Vec<u8> {
        value
            .as_bytes()
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }
}
