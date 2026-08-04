//! Typed, inert XLSB External Link data (MS-XLSB 2.1.7.25).
//!
//! These values describe stored package metadata and caches. They never open
//! an external workbook, contact DDE, instantiate OLE, refresh data, evaluate
//! formulas, or execute code.

use crate::raw::Error as WireError;
use std::collections::HashSet;
use thiserror::Error as ThisError;

#[path = "external_link_write.rs"]
mod external_link_write;
pub use external_link_write::write_external_link_stream;

/// Result type for the standalone XLSB external-link codec.
pub type Result<T> = std::result::Result<T, Error>;

/// Error returned by the standalone XLSB external-link codec.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A validated BIFF12 wire operation failed.
    #[error(transparent)]
    Wire(#[from] WireError),
    /// A modeled external-link invariant was violated.
    #[error("invalid external link: {0}")]
    InvalidFormula(String),
    /// A fixed-width or length-prefixed structure has the wrong size.
    #[error("invalid length: expected {expected}, found {found}")]
    InvalidLength { expected: usize, found: usize },
    /// A bounded collection could not reserve its validated capacity.
    #[error("allocation failed for {resource}: {source}")]
    Allocation {
        resource: &'static str,
        #[source]
        source: std::collections::TryReserveError,
    },
}

const MAX_COLLECTION_ITEMS: usize = 65_535;
const MAX_WIDE_STRING_UNITS: usize = 32_767;
/// Maximum row count accepted in an authored or parsed DDE/OLE cache.
pub const MAX_XLSB_EXTERNAL_CACHE_ROWS: u32 = 1_048_576;
/// Maximum column count accepted in an authored or parsed DDE/OLE cache.
pub const MAX_XLSB_EXTERNAL_CACHE_COLUMNS: u32 = 16_384;
/// Safety limit for materialized values across one DDE/OLE cache matrix.
pub const MAX_XLSB_EXTERNAL_CACHED_VALUES: usize = 1_048_576;

pub const EXTERNAL_REFERENCE_WORKBOOK: u16 = 0;
pub const EXTERNAL_REFERENCE_DDE: u16 = 1;
pub const EXTERNAL_REFERENCE_OLE: u16 = 2;
pub const EXTERNAL_NAME_BUILT_IN: u8 = 1;
pub const EXTERNAL_NAME_RESERVED_MASK: u8 = 0b0011_1110;
pub const DATA_ITEM_WANT_ADVISE: u8 = 1 << 1;
pub const DATA_ITEM_WANT_PICTURE: u8 = 1 << 2;
pub const DDE_ITEM_SUPPORTS_OLE: u8 = 1 << 3;
pub const DDE_ITEM_RESERVED_MASK: u8 = 0b0011_0001;
pub const OLE_ITEM_REQUIRED_CLASS_FLAG: u8 = 1 << 4;
pub const OLE_ITEM_DISPLAY_AS_ICON: u8 = 1 << 5;
pub const OLE_ITEM_RESERVED_MASK: u8 = 0b0000_1001;
pub const DATA_ITEM_REQUIRED_TRAILING_FLAG: u8 = 1;

const EXT_PTG_ERROR: u8 = 0x1C;
const EXT_PTG_REFERENCE: u8 = 0x3A;
const EXT_PTG_AREA: u8 = 0x3B;
const EXT_PTG_REFERENCE_ERROR: u8 = 0x3C;
const EXT_PTG_AREA_ERROR: u8 = 0x3D;
const REFERENCE_ERROR_CODE: u8 = 0x17;

// The owner keeps this validator local so external-name parsing does not depend on host workbook state.
fn validate_defined_name(name: &str) -> Result<()> {
    let utf16_len = name.encode_utf16().count();
    if utf16_len == 0 || utf16_len > 255 {
        return Err(Error::InvalidFormula(format!(
            "defined name length {utf16_len} is outside 1..=255"
        )));
    }
    let mut chars = name.chars();
    let first = chars.next().expect("checked non-empty defined name");
    if !is_name_start(first) || !chars.all(is_name_character) {
        return Err(Error::InvalidFormula(format!(
            "defined name {name:?} does not follow XLNameWideString grammar"
        )));
    }
    if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
        return Err(Error::InvalidFormula(format!(
            "defined name {name:?} is a reserved Boolean literal"
        )));
    }
    if is_a1_reference(name) || starts_with_r1c1_reference(name) {
        return Err(Error::InvalidFormula(format!(
            "defined name {name:?} conflicts with a cell reference"
        )));
    }
    Ok(())
}

fn is_name_start(value: char) -> bool {
    value == '_' || value == '\\' || value.is_ascii_alphabetic() || value.is_alphabetic()
}

fn is_name_character(value: char) -> bool {
    is_name_start(value) || matches!(value, '?' | '\u{061F}' | '.') || value.is_numeric()
}

fn is_a1_reference(value: &str) -> bool {
    let bytes = value.as_bytes();
    let split = bytes
        .iter()
        .position(u8::is_ascii_digit)
        .unwrap_or(bytes.len());
    if split == 0 || split > 3 || split == bytes.len() {
        return false;
    }
    if !bytes[..split].iter().all(u8::is_ascii_alphabetic)
        || !bytes[split..].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let mut column = 0u32;
    for byte in bytes[..split].iter().map(u8::to_ascii_uppercase) {
        let Some(next) = column
            .checked_mul(26)
            .and_then(|column| column.checked_add(u32::from(byte - b'A' + 1)))
        else {
            return false;
        };
        column = next;
    }
    let Some(row) = value[split..].parse::<u32>().ok() else {
        return false;
    };
    column <= 16_384 && (1..=1_048_576).contains(&row)
}

fn starts_with_r1c1_reference(value: &str) -> bool {
    let bytes = value.as_bytes();
    let Some(first) = bytes.first().copied().map(|byte| byte.to_ascii_uppercase()) else {
        return false;
    };
    match first {
        b'R' => numeric_reference_prefix(bytes, 1, 1_048_576).is_some(),
        b'C' => numeric_reference_prefix(bytes, 1, 16_384).is_some(),
        _ => false,
    }
}

fn numeric_reference_prefix(bytes: &[u8], offset: usize, maximum: u32) -> Option<usize> {
    let end = bytes[offset..]
        .iter()
        .position(|byte| !byte.is_ascii_digit())
        .map_or(bytes.len(), |position| offset + position);
    if end == offset {
        return None;
    }
    let value = std::str::from_utf8(&bytes[offset..end])
        .ok()?
        .parse::<u32>()
        .ok()?;
    (1..=maximum).contains(&value).then_some(end)
}

/// Stored kind of an XLSB External Link part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsbExternalLinkKind {
    /// A link to another workbook.
    Workbook,
    /// A link to a Dynamic Data Exchange server and topic.
    Dde,
    /// A link to an OLE data source.
    Ole,
}

/// The one token permitted in an external defined-name formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsbExternalNameFormulaKind {
    CellReference,
    AreaReference,
    CellReferenceError,
    AreaReferenceError,
    ReferenceError,
}

/// Sheet range used by an external defined-name reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsbExternalSheetRange {
    /// The referenced external sheet cannot be found.
    Missing,
    /// Inclusive zero-based external sheet indices.
    Sheets { first: u16, last: u16 },
}

impl XlsbExternalSheetRange {
    /// Create a validated inclusive range of external sheets.
    pub fn sheets(first: u16, last: u16) -> Result<Self> {
        if last < first || last > i16::MAX as u16 {
            return Err(Error::InvalidFormula(format!(
                "invalid external sheet range {first}..={last}"
            )));
        }
        Ok(Self::Sheets { first, last })
    }

    fn encode(self) -> [u8; 4] {
        match self {
            Self::Missing => [0xFF; 4],
            Self::Sheets { first, last } => {
                let first = i16::try_from(first).expect("sheet range was validated");
                let last = i16::try_from(last).expect("sheet range was validated");
                let mut encoded = [0; 4];
                encoded[..2].copy_from_slice(&first.to_le_bytes());
                encoded[2..].copy_from_slice(&last.to_le_bytes());
                encoded
            },
        }
    }
}

/// Cell coordinates and relative flags in an external-name reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbExternalCellLocation {
    row: u16,
    column: u8,
    column_relative: bool,
    row_relative: bool,
}

impl XlsbExternalCellLocation {
    /// Create an absolute cell location in the first 65,536 rows and 256 columns.
    pub const fn new(row: u16, column: u8) -> Self {
        Self {
            row,
            column,
            column_relative: false,
            row_relative: false,
        }
    }

    pub const fn with_column_relative(mut self, relative: bool) -> Self {
        self.column_relative = relative;
        self
    }

    pub const fn with_row_relative(mut self, relative: bool) -> Self {
        self.row_relative = relative;
        self
    }

    pub const fn row(self) -> u16 {
        self.row
    }

    pub const fn column(self) -> u8 {
        self.column
    }

    pub const fn is_column_relative(self) -> bool {
        self.column_relative
    }

    pub const fn is_row_relative(self) -> bool {
        self.row_relative
    }

    fn encode_column(self) -> [u8; 2] {
        let encoded = u16::from(self.column)
            | (u16::from(self.column_relative) << 14)
            | (u16::from(self.row_relative) << 15);
        encoded.to_le_bytes()
    }
}

/// A typed external cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbExternalCellReference {
    sheets: XlsbExternalSheetRange,
    location: XlsbExternalCellLocation,
}

impl XlsbExternalCellReference {
    pub const fn new(sheets: XlsbExternalSheetRange, location: XlsbExternalCellLocation) -> Self {
        Self { sheets, location }
    }

    pub const fn sheets(self) -> XlsbExternalSheetRange {
        self.sheets
    }

    pub const fn location(self) -> XlsbExternalCellLocation {
        self.location
    }
}

/// A typed external rectangular area reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbExternalAreaReference {
    sheets: XlsbExternalSheetRange,
    first: XlsbExternalCellLocation,
    last: XlsbExternalCellLocation,
}

impl XlsbExternalAreaReference {
    pub fn new(
        sheets: XlsbExternalSheetRange,
        first: XlsbExternalCellLocation,
        last: XlsbExternalCellLocation,
    ) -> Result<Self> {
        if last.row < first.row || last.column < first.column {
            return Err(Error::InvalidFormula(
                "external-name area reference is reversed".to_string(),
            ));
        }
        Ok(Self {
            sheets,
            first,
            last,
        })
    }

    pub const fn sheets(self) -> XlsbExternalSheetRange {
        self.sheets
    }

    pub const fn first(self) -> XlsbExternalCellLocation {
        self.first
    }

    pub const fn last(self) -> XlsbExternalCellLocation {
        self.last
    }
}

/// A validated external defined-name formula.
///
/// MS-XLSB restricts this formula to exactly one of five external Ptg token
/// structures. The original bytes are retained so relative-reference flags
/// and undefined bytes in error tokens round-trip losslessly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbExternalNameFormula {
    kind: XlsbExternalNameFormulaKind,
    tokens: Vec<u8>,
}

impl XlsbExternalNameFormula {
    /// Validate and retain one external-name Ptg token.
    pub fn from_tokens(tokens: Vec<u8>) -> Result<Self> {
        let kind = validate_external_name_token(&tokens, None)?;
        Ok(Self { kind, tokens })
    }

    /// Create an external cell-reference formula without raw token bytes.
    pub fn cell_reference(reference: XlsbExternalCellReference) -> Self {
        let mut tokens = Vec::with_capacity(9);
        tokens.push(EXT_PTG_REFERENCE);
        tokens.extend_from_slice(&reference.sheets.encode());
        tokens.extend_from_slice(&reference.location.row.to_le_bytes());
        tokens.extend_from_slice(&reference.location.encode_column());
        Self {
            kind: XlsbExternalNameFormulaKind::CellReference,
            tokens,
        }
    }

    /// Create an external area-reference formula without raw token bytes.
    pub fn area_reference(reference: XlsbExternalAreaReference) -> Self {
        let mut tokens = Vec::with_capacity(13);
        tokens.push(EXT_PTG_AREA);
        tokens.extend_from_slice(&reference.sheets.encode());
        tokens.extend_from_slice(&reference.first.row.to_le_bytes());
        tokens.extend_from_slice(&reference.last.row.to_le_bytes());
        tokens.extend_from_slice(&reference.first.encode_column());
        tokens.extend_from_slice(&reference.last.encode_column());
        Self {
            kind: XlsbExternalNameFormulaKind::AreaReference,
            tokens,
        }
    }

    /// Create an invalid external cell-reference token for a known sheet range.
    pub fn cell_reference_error(sheets: XlsbExternalSheetRange) -> Self {
        let mut tokens = Vec::with_capacity(9);
        tokens.push(EXT_PTG_REFERENCE_ERROR);
        tokens.extend_from_slice(&sheets.encode());
        tokens.extend_from_slice(&[0; 4]);
        Self {
            kind: XlsbExternalNameFormulaKind::CellReferenceError,
            tokens,
        }
    }

    /// Create an invalid external area-reference token for a known sheet range.
    pub fn area_reference_error(sheets: XlsbExternalSheetRange) -> Self {
        let mut tokens = Vec::with_capacity(13);
        tokens.push(EXT_PTG_AREA_ERROR);
        tokens.extend_from_slice(&sheets.encode());
        tokens.extend_from_slice(&[0; 8]);
        Self {
            kind: XlsbExternalNameFormulaKind::AreaReferenceError,
            tokens,
        }
    }

    /// Create the standalone `#REF!` external-name formula token.
    pub fn reference_error() -> Self {
        Self {
            kind: XlsbExternalNameFormulaKind::ReferenceError,
            tokens: vec![EXT_PTG_ERROR, REFERENCE_ERROR_CODE],
        }
    }

    pub const fn kind(&self) -> XlsbExternalNameFormulaKind {
        self.kind
    }

    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// Return the typed cell reference when this is `CellReference`.
    pub fn cell(&self) -> Option<XlsbExternalCellReference> {
        (self.kind == XlsbExternalNameFormulaKind::CellReference).then(|| {
            XlsbExternalCellReference::new(
                decode_sheet_range(&self.tokens[1..5]),
                decode_cell_location(&self.tokens[5..7], &self.tokens[7..9]),
            )
        })
    }

    /// Return the typed area reference when this is `AreaReference`.
    pub fn area(&self) -> Option<XlsbExternalAreaReference> {
        (self.kind == XlsbExternalNameFormulaKind::AreaReference).then(|| {
            XlsbExternalAreaReference::new(
                decode_sheet_range(&self.tokens[1..5]),
                decode_cell_location(&self.tokens[5..7], &self.tokens[9..11]),
                decode_cell_location(&self.tokens[7..9], &self.tokens[11..13]),
            )
            .expect("validated external-name area")
        })
    }

    /// Return referenced sheets for any cell/area token, including error forms.
    ///
    /// The standalone `#REF!` token has no sheet range.
    pub fn sheets(&self) -> Option<XlsbExternalSheetRange> {
        (self.kind != XlsbExternalNameFormulaKind::ReferenceError)
            .then(|| decode_sheet_range(&self.tokens[1..5]))
    }

    fn validate_for_sheets(&self, sheet_count: usize) -> Result<()> {
        let kind = validate_external_name_token(&self.tokens, Some(sheet_count))?;
        if kind != self.kind {
            return Err(Error::InvalidFormula(
                "external-name formula kind does not match its token".to_string(),
            ));
        }
        Ok(())
    }
}

/// Error value in a cached DDE/OLE matrix (`BErr`, MS-XLSB 2.5.98.2).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsbExternalErrorValue {
    Null,
    DivisionByZero,
    Value,
    Reference,
    Name,
    Number,
    NotAvailable,
    GettingData,
}

impl XlsbExternalErrorValue {
    pub const fn code(self) -> u8 {
        match self {
            Self::Null => 0x00,
            Self::DivisionByZero => 0x07,
            Self::Value => 0x0F,
            Self::Reference => 0x17,
            Self::Name => 0x1D,
            Self::Number => 0x24,
            Self::NotAvailable => 0x2A,
            Self::GettingData => 0x2B,
        }
    }

    pub fn from_code(code: u8) -> Result<Self> {
        match code {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0F => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1D => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2A => Ok(Self::NotAvailable),
            0x2B => Ok(Self::GettingData),
            _ => Err(Error::InvalidFormula(format!(
                "invalid cached external error code 0x{code:02X}"
            ))),
        }
    }
}

/// One inert cached value in a DDE/OLE matrix.
#[derive(Debug, Clone, PartialEq)]
pub enum XlsbExternalCachedValue {
    Empty,
    Number(f64),
    Boolean(bool),
    Error(XlsbExternalErrorValue),
    String(String),
}

/// A bounded, row-major cached DDE/OLE value matrix.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbExternalValueMatrix {
    rows: u32,
    columns: u32,
    values: Vec<XlsbExternalCachedValue>,
}

impl XlsbExternalValueMatrix {
    pub fn new(rows: u32, columns: u32, values: Vec<XlsbExternalCachedValue>) -> Result<Self> {
        let matrix = Self {
            rows,
            columns,
            values,
        };
        matrix.validate()?;
        Ok(matrix)
    }

    pub const fn rows(&self) -> u32 {
        self.rows
    }

    pub const fn columns(&self) -> u32 {
        self.columns
    }

    pub fn values(&self) -> &[XlsbExternalCachedValue] {
        &self.values
    }

    fn validate(&self) -> Result<()> {
        if self.rows == 0
            || self.rows > MAX_XLSB_EXTERNAL_CACHE_ROWS
            || self.columns == 0
            || self.columns > MAX_XLSB_EXTERNAL_CACHE_COLUMNS
        {
            return Err(Error::InvalidFormula(
                "external cache dimensions exceed worksheet bounds".to_string(),
            ));
        }
        let expected = usize::try_from(self.rows)
            .ok()
            .and_then(|rows| {
                usize::try_from(self.columns)
                    .ok()
                    .and_then(|columns| rows.checked_mul(columns))
            })
            .ok_or_else(|| {
                Error::InvalidFormula("external cache dimensions overflow".to_string())
            })?;
        if expected != self.values.len() || expected > MAX_XLSB_EXTERNAL_CACHED_VALUES {
            return Err(Error::InvalidLength {
                expected: expected.min(MAX_XLSB_EXTERNAL_CACHED_VALUES),
                found: self.values.len(),
            });
        }
        for value in &self.values {
            match value {
                XlsbExternalCachedValue::Number(number) => validate_external_number(*number)?,
                XlsbExternalCachedValue::String(value)
                    if value.encode_utf16().count() > MAX_WIDE_STRING_UNITS =>
                {
                    return Err(Error::InvalidFormula(
                        "external cached string exceeds 32,767 UTF-16 units".to_string(),
                    ));
                },
                _ => {},
            }
        }
        Ok(())
    }
}

/// One defined name stored by an external-workbook link.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbExternalDefinedName {
    name: String,
    formula: Option<XlsbExternalNameFormula>,
    built_in: bool,
    scope_sheet_index: Option<u16>,
}

impl XlsbExternalDefinedName {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let entry = Self {
            name: name.into(),
            formula: None,
            built_in: false,
            scope_sheet_index: None,
        };
        entry.validate(0, false)?;
        Ok(entry)
    }

    pub fn with_formula(mut self, formula: XlsbExternalNameFormula) -> Self {
        self.formula = Some(formula);
        self
    }

    pub fn with_built_in(mut self, built_in: bool) -> Self {
        self.built_in = built_in;
        self
    }

    pub fn with_sheet_scope(mut self, zero_based_sheet_index: u16) -> Self {
        self.scope_sheet_index = Some(zero_based_sheet_index);
        self
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn formula(&self) -> Option<&XlsbExternalNameFormula> {
        self.formula.as_ref()
    }

    pub const fn is_built_in(&self) -> bool {
        self.built_in
    }

    pub const fn scope_sheet_index(&self) -> Option<u16> {
        self.scope_sheet_index
    }

    fn validate(&self, sheet_count: usize, validate_scope: bool) -> Result<()> {
        validate_defined_name(&self.name)?;
        if validate_scope
            && self
                .scope_sheet_index
                .is_some_and(|index| usize::from(index) >= sheet_count)
        {
            return Err(Error::InvalidFormula(format!(
                "external defined name {:?} has an invalid sheet scope",
                self.name
            )));
        }
        if let Some(formula) = &self.formula {
            formula.validate_for_sheets(sheet_count)?;
        }
        Ok(())
    }
}

/// One DDE data item and its optional inert cache.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbDdeItem {
    name: String,
    want_advise: bool,
    want_picture: bool,
    supports_ole: bool,
    cached_values: Option<XlsbExternalValueMatrix>,
}

impl XlsbDdeItem {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let item = Self {
            name: name.into(),
            want_advise: false,
            want_picture: false,
            supports_ole: false,
            cached_values: None,
        };
        item.validate()?;
        Ok(item)
    }

    pub fn with_advise(mut self, enabled: bool) -> Self {
        self.want_advise = enabled;
        self
    }

    pub fn with_picture(mut self, enabled: bool) -> Self {
        self.want_picture = enabled;
        self
    }

    pub fn with_ole_support(mut self, enabled: bool) -> Self {
        self.supports_ole = enabled;
        self
    }

    pub fn with_cached_values(mut self, values: XlsbExternalValueMatrix) -> Self {
        self.cached_values = Some(values);
        self
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub const fn wants_advise(&self) -> bool {
        self.want_advise
    }

    pub const fn wants_picture(&self) -> bool {
        self.want_picture
    }

    pub const fn supports_ole(&self) -> bool {
        self.supports_ole
    }

    pub fn cached_values(&self) -> Option<&XlsbExternalValueMatrix> {
        self.cached_values.as_ref()
    }

    fn validate(&self) -> Result<()> {
        validate_defined_name(&self.name)?;
        if self.supports_ole && self.name != "StdDocumentName" {
            return Err(Error::InvalidFormula(
                "OLE-supporting DDE item must be named \"StdDocumentName\"".to_string(),
            ));
        }
        if let Some(values) = &self.cached_values {
            values.validate()?;
        }
        Ok(())
    }
}

/// One OLE data item and its optional inert cache.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbOleItem {
    name: String,
    want_advise: bool,
    want_picture: bool,
    display_as_icon: bool,
    cached_values: Option<XlsbExternalValueMatrix>,
}

impl XlsbOleItem {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let item = Self {
            name: name.into(),
            want_advise: false,
            want_picture: false,
            display_as_icon: false,
            cached_values: None,
        };
        item.validate()?;
        Ok(item)
    }

    pub fn with_advise(mut self, enabled: bool) -> Self {
        self.want_advise = enabled;
        self
    }

    pub fn with_picture(mut self, enabled: bool) -> Self {
        self.want_picture = enabled;
        self
    }

    pub fn with_icon(mut self, enabled: bool) -> Self {
        self.display_as_icon = enabled;
        self
    }

    pub fn with_cached_values(mut self, values: XlsbExternalValueMatrix) -> Self {
        self.cached_values = Some(values);
        self
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub const fn wants_advise(&self) -> bool {
        self.want_advise
    }

    pub const fn wants_picture(&self) -> bool {
        self.want_picture
    }

    pub const fn displays_as_icon(&self) -> bool {
        self.display_as_icon
    }

    pub fn cached_values(&self) -> Option<&XlsbExternalValueMatrix> {
        self.cached_values.as_ref()
    }

    fn validate(&self) -> Result<()> {
        validate_defined_name(&self.name)?;
        if let Some(values) = &self.cached_values {
            values.validate()?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsbExternalEntries {
    Workbook(Vec<XlsbExternalDefinedName>),
    Dde(Vec<XlsbDdeItem>),
    Ole(Vec<XlsbOleItem>),
}

/// Typed metadata and inert caches from one XLSB External Link part.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbExternalLink {
    pub kind: XlsbExternalLinkKind,
    pub source: String,
    pub detail: Option<String>,
    pub sheet_names: Vec<String>,
    pub entries: XlsbExternalEntries,
}

impl XlsbExternalLink {
    pub fn workbook(
        source: impl Into<String>,
        sheet_names: Vec<String>,
        declared_names: Vec<String>,
    ) -> Result<Self> {
        let names = declared_names
            .into_iter()
            .map(XlsbExternalDefinedName::new)
            .collect::<Result<Vec<_>>>()?;
        Self::workbook_with_defined_names(source, sheet_names, names)
    }

    pub fn workbook_with_defined_names(
        source: impl Into<String>,
        sheet_names: Vec<String>,
        defined_names: Vec<XlsbExternalDefinedName>,
    ) -> Result<Self> {
        let link = Self {
            kind: XlsbExternalLinkKind::Workbook,
            source: source.into(),
            detail: None,
            sheet_names,
            entries: XlsbExternalEntries::Workbook(defined_names),
        };
        link.validate()?;
        Ok(link)
    }

    pub fn dde(
        server: impl Into<String>,
        topic: impl Into<String>,
        item_names: Vec<String>,
    ) -> Result<Self> {
        let items = item_names
            .into_iter()
            .map(XlsbDdeItem::new)
            .collect::<Result<Vec<_>>>()?;
        Self::dde_with_items(server, topic, items)
    }

    pub fn dde_with_items(
        server: impl Into<String>,
        topic: impl Into<String>,
        items: Vec<XlsbDdeItem>,
    ) -> Result<Self> {
        let link = Self {
            kind: XlsbExternalLinkKind::Dde,
            source: server.into(),
            detail: Some(topic.into()),
            sheet_names: Vec::new(),
            entries: XlsbExternalEntries::Dde(items),
        };
        link.validate()?;
        Ok(link)
    }

    pub fn ole(
        source: impl Into<String>,
        program_id: impl Into<String>,
        item_names: Vec<String>,
    ) -> Result<Self> {
        let items = item_names
            .into_iter()
            .map(XlsbOleItem::new)
            .collect::<Result<Vec<_>>>()?;
        Self::ole_with_items(source, program_id, items)
    }

    pub fn ole_with_items(
        source: impl Into<String>,
        program_id: impl Into<String>,
        items: Vec<XlsbOleItem>,
    ) -> Result<Self> {
        let link = Self {
            kind: XlsbExternalLinkKind::Ole,
            source: source.into(),
            detail: Some(program_id.into()),
            sheet_names: Vec::new(),
            entries: XlsbExternalEntries::Ole(items),
        };
        link.validate()?;
        Ok(link)
    }

    pub const fn kind(&self) -> XlsbExternalLinkKind {
        self.kind
    }

    pub fn source(&self) -> &str {
        &self.source
    }

    pub fn dde_topic(&self) -> Option<&str> {
        match self.kind {
            XlsbExternalLinkKind::Dde => self.detail.as_deref(),
            XlsbExternalLinkKind::Workbook | XlsbExternalLinkKind::Ole => None,
        }
    }

    pub fn ole_program_id(&self) -> Option<&str> {
        match self.kind {
            XlsbExternalLinkKind::Ole => self.detail.as_deref(),
            XlsbExternalLinkKind::Workbook | XlsbExternalLinkKind::Dde => None,
        }
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    pub fn defined_names(&self) -> &[XlsbExternalDefinedName] {
        match &self.entries {
            XlsbExternalEntries::Workbook(entries) => entries,
            XlsbExternalEntries::Dde(_) | XlsbExternalEntries::Ole(_) => &[],
        }
    }

    pub fn dde_items(&self) -> &[XlsbDdeItem] {
        match &self.entries {
            XlsbExternalEntries::Dde(entries) => entries,
            XlsbExternalEntries::Workbook(_) | XlsbExternalEntries::Ole(_) => &[],
        }
    }

    pub fn ole_items(&self) -> &[XlsbOleItem] {
        match &self.entries {
            XlsbExternalEntries::Ole(entries) => entries,
            XlsbExternalEntries::Workbook(_) | XlsbExternalEntries::Dde(_) => &[],
        }
    }

    pub const fn is_workbook(&self) -> bool {
        matches!(self.kind, XlsbExternalLinkKind::Workbook)
    }

    pub fn validate(&self) -> Result<()> {
        validate_wide_string(&self.source, "external-link source")?;
        if self.sheet_names.len() > MAX_COLLECTION_ITEMS
            || self.entry_count() > MAX_COLLECTION_ITEMS
        {
            return Err(Error::InvalidFormula(
                "external-link collection exceeds 65,535 items".to_string(),
            ));
        }
        let mut folded_sheet_names = HashSet::with_capacity(self.sheet_names.len());
        for sheet_name in &self.sheet_names {
            validate_wide_string(sheet_name, "external sheet name")?;
            if !folded_sheet_names.insert(sheet_name.to_lowercase()) {
                return Err(Error::InvalidFormula(format!(
                    "duplicate external sheet name {sheet_name:?}"
                )));
            }
        }
        match (&self.kind, &self.entries) {
            (XlsbExternalLinkKind::Workbook, XlsbExternalEntries::Workbook(entries)) => {
                if self.detail.is_some() {
                    return Err(Error::InvalidFormula(
                        "external workbook link cannot have DDE/OLE detail".to_string(),
                    ));
                }
                validate_unique_entries(entries.iter().map(XlsbExternalDefinedName::name))?;
                for entry in entries {
                    entry.validate(self.sheet_names.len(), true)?;
                }
            },
            (XlsbExternalLinkKind::Dde, XlsbExternalEntries::Dde(entries)) => {
                self.validate_data_source_detail("DDE topic")?;
                validate_unique_entries(entries.iter().map(XlsbDdeItem::name))?;
                for entry in entries {
                    entry.validate()?;
                }
            },
            (XlsbExternalLinkKind::Ole, XlsbExternalEntries::Ole(entries)) => {
                self.validate_data_source_detail("OLE program ID")?;
                validate_unique_entries(entries.iter().map(XlsbOleItem::name))?;
                for entry in entries {
                    entry.validate()?;
                }
            },
            _ => {
                return Err(Error::InvalidFormula(
                    "external-link kind does not match its entry collection".to_string(),
                ));
            },
        }
        Ok(())
    }

    pub fn entry_count(&self) -> usize {
        match &self.entries {
            XlsbExternalEntries::Workbook(entries) => entries.len(),
            XlsbExternalEntries::Dde(entries) => entries.len(),
            XlsbExternalEntries::Ole(entries) => entries.len(),
        }
    }

    fn validate_data_source_detail(&self, context: &str) -> Result<()> {
        if !self.sheet_names.is_empty() {
            return Err(Error::InvalidFormula(
                "DDE/OLE links cannot declare workbook sheet names".to_string(),
            ));
        }
        validate_wide_string(self.detail.as_deref().unwrap_or_default(), context)
    }
}

/// Parsed `BrtSupBook` metadata and the relationship identifier which still
/// needs to be resolved by the OPC host.
///
/// The owner deliberately does not inspect or follow package relationships.
/// For workbook and OLE links, [`Self::relationship_id`] is the identifier
/// stored in `BrtBeginSupBook`; the host validates its type and resolves its
/// inert target before calling [`Self::resolve_source`].
#[derive(Debug, Clone, PartialEq)]
pub struct ParsedExternalLink {
    link: XlsbExternalLink,
    relationship_id: Option<String>,
}

impl ParsedExternalLink {
    /// Borrow the parsed metadata. Workbook/OLE `source()` is the unresolved
    /// relationship identifier until the host resolves it.
    pub const fn link(&self) -> &XlsbExternalLink {
        &self.link
    }

    /// Consume the parsed value and return the unresolved metadata.
    pub fn into_link(self) -> XlsbExternalLink {
        self.link
    }

    /// Return the relationship identifier, if this link has one.
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Replace the unresolved relationship identifier with its inert target.
    pub fn resolve_source(mut self, source: impl Into<String>) -> Result<XlsbExternalLink> {
        if self.relationship_id.is_none() {
            return Err(Error::InvalidFormula(
                "DDE external links do not have a relationship source".to_string(),
            ));
        }
        self.link.source = source.into();
        self.link.validate()?;
        Ok(self.link)
    }
}

/// Maximum bytes accepted or emitted for one external-link part stream.
pub const MAX_XLSB_EXTERNAL_LINK_PART_BYTES: usize = 32 * 1024 * 1024;

/// Parse one complete XLSB External Link part stream.
pub fn parse_external_link(data: &[u8]) -> Result<ParsedExternalLink> {
    if data.len() > MAX_XLSB_EXTERNAL_LINK_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_XLSB_EXTERNAL_LINK_PART_BYTES,
            found: data.len(),
        });
    }
    let limits = crate::raw::Limits::new(MAX_XLSB_EXTERNAL_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut link_type = None;
    let mut target_key = String::new();
    let mut target_detail = String::new();
    let mut sheet_names = Vec::new();
    let mut workbook_entries = Vec::new();
    let mut dde_entries = Vec::new();
    let mut ole_entries = Vec::new();
    let mut saw_sup_tabs = false;
    // 0 = outside a name, 1 = expect formula, 2 = expect bits,
    // 3 = expect end/value start, 4 = inside a cached matrix.
    let mut sup_name_state = 0u8;
    let mut current_name = None;
    let mut current_formula = None;
    let mut current_bits = None;
    let mut current_cache = None;
    let mut cache_dimensions = None;
    let mut cache_values = Vec::new();
    let mut saw_end = false;

    for record in crate::raw::Records::with_limits(data, limits) {
        let record = record?;
        if saw_end {
            return Err(invalid("external link has records after BrtEndSupBook"));
        }
        if link_type.is_none() && record.kind() != crate::raw::kind::BEGIN_SUP_BOOK {
            return Err(invalid("external link does not start with BrtBeginSupBook"));
        }
        match record.kind() {
            crate::raw::kind::BEGIN_SUP_BOOK => {
                if link_type.is_some() || record.payload().len() < 10 {
                    return Err(invalid("invalid BrtBeginSupBook framing"));
                }
                let mut cursor =
                    crate::raw::Cursor::with_limits(record.payload(), "BrtBeginSupBook", limits);
                let kind = cursor.read_u16()?;
                let first = cursor.read_wide_string()?;
                let second = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                    cursor.read_nullable_wide_string()?
                } else {
                    Some(cursor.read_wide_string()?)
                };
                cursor.finish()?;
                if kind > EXTERNAL_REFERENCE_OLE || first.is_empty() {
                    return Err(invalid("invalid BrtBeginSupBook payload"));
                }
                if kind == EXTERNAL_REFERENCE_WORKBOOK && second.is_some() {
                    return Err(invalid(
                        "external workbook BrtBeginSupBook string2 is not NULL",
                    ));
                }
                link_type = Some(kind);
                target_key = first;
                target_detail = second.unwrap_or_default();
            },
            crate::raw::kind::SUP_TABS => {
                if link_type != Some(EXTERNAL_REFERENCE_WORKBOOK)
                    || saw_sup_tabs
                    || sup_name_state != 0
                {
                    return Err(invalid("unexpected BrtSupTabs"));
                }
                sheet_names = parse_external_sheet_names(record.payload(), limits)?;
                saw_sup_tabs = true;
            },
            crate::raw::kind::SUP_NAME_START => {
                let kind =
                    link_type.ok_or_else(|| invalid("BrtSupNameStart precedes BrtBeginSupBook"))?;
                if sup_name_state != 0 || (kind == EXTERNAL_REFERENCE_WORKBOOK && !saw_sup_tabs) {
                    return Err(invalid("unexpected BrtSupNameStart"));
                }
                let mut cursor =
                    crate::raw::Cursor::with_limits(record.payload(), "BrtSupNameStart", limits);
                let name = cursor.read_wide_string()?;
                cursor.finish()?;
                validate_defined_name(&name)?;
                current_name = Some(name);
                sup_name_state = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                    1
                } else {
                    2
                };
            },
            crate::raw::kind::SUP_NAME_FORMULA => {
                if link_type != Some(EXTERNAL_REFERENCE_WORKBOOK) || sup_name_state != 1 {
                    return Err(invalid("unexpected BrtSupNameFmla"));
                }
                let mut cursor =
                    crate::raw::Cursor::with_limits(record.payload(), "BrtSupNameFmla", limits);
                if record.payload().len() < 4 {
                    return Err(Error::InvalidLength {
                        expected: 4,
                        found: record.payload().len(),
                    });
                }
                let formula_len = usize::try_from(cursor.read_u32()?)
                    .map_err(|_| invalid("BrtSupNameFmla size overflow"))?;
                let formula = cursor.read_bytes(formula_len)?.to_vec();
                cursor.finish()?;
                current_formula = if formula.is_empty() {
                    None
                } else {
                    Some(XlsbExternalNameFormula::from_tokens(formula)?)
                };
                sup_name_state = 2;
            },
            crate::raw::kind::SUP_NAME_BITS => {
                if sup_name_state != 2 || record.payload().len() != 7 {
                    return Err(invalid("unexpected BrtSupNameBits"));
                }
                let mut bits = [0u8; 7];
                bits.copy_from_slice(record.payload());
                validate_external_name_bits(
                    link_type.expect("external link kind is present"),
                    &bits,
                )?;
                current_bits = Some(bits);
                sup_name_state = 3;
            },
            crate::raw::kind::SUP_NAME_VALUE_START => {
                if !matches!(
                    link_type,
                    Some(EXTERNAL_REFERENCE_DDE | EXTERNAL_REFERENCE_OLE)
                ) || sup_name_state != 3
                    || record.payload().len() != 8
                    || current_cache.is_some()
                {
                    return Err(invalid("unexpected BrtSupNameValueStart"));
                }
                let mut cursor = crate::raw::Cursor::with_limits(
                    record.payload(),
                    "BrtSupNameValueStart",
                    limits,
                );
                let rows = cursor.read_u32()?;
                let columns = cursor.read_u32()?;
                cursor.finish()?;
                let count = usize::try_from(rows)
                    .ok()
                    .and_then(|rows| {
                        usize::try_from(columns)
                            .ok()
                            .and_then(|columns| rows.checked_mul(columns))
                    })
                    .ok_or_else(|| invalid("external cached-value dimensions overflow"))?;
                if count > MAX_XLSB_EXTERNAL_CACHED_VALUES {
                    return Err(Error::InvalidLength {
                        expected: MAX_XLSB_EXTERNAL_CACHED_VALUES,
                        found: count,
                    });
                }
                cache_values.clear();
                cache_values
                    .try_reserve(count)
                    .map_err(|source| Error::Allocation {
                        resource: "external cached values",
                        source,
                    })?;
                cache_dimensions = Some((rows, columns, count));
                sup_name_state = 4;
            },
            crate::raw::kind::SUP_NAME_NIL
            | crate::raw::kind::SUP_NAME_NUM
            | crate::raw::kind::SUP_NAME_BOOL
            | crate::raw::kind::SUP_NAME_ERROR
            | crate::raw::kind::SUP_NAME_STRING => {
                let Some((_, _, count)) = cache_dimensions else {
                    return Err(invalid("cached external value occurs outside its matrix"));
                };
                if sup_name_state != 4 || cache_values.len() >= count {
                    return Err(invalid("too many or misplaced cached external values"));
                }
                cache_values.push(parse_external_cached_value(
                    record.kind(),
                    record.payload(),
                    limits,
                )?);
            },
            crate::raw::kind::SUP_NAME_VALUE_END => {
                let Some((rows, columns, count)) = cache_dimensions.take() else {
                    return Err(invalid("unexpected BrtSupNameValueEnd"));
                };
                if sup_name_state != 4
                    || !record.payload().is_empty()
                    || cache_values.len() != count
                {
                    return Err(invalid("invalid cached external value matrix"));
                }
                current_cache = Some(XlsbExternalValueMatrix::new(
                    rows,
                    columns,
                    std::mem::take(&mut cache_values),
                )?);
                sup_name_state = 3;
            },
            crate::raw::kind::SUP_NAME_END => {
                if sup_name_state != 3 || !record.payload().is_empty() {
                    return Err(invalid("invalid BrtSupNameEnd"));
                }
                let kind = link_type.expect("external link kind is present");
                let name = current_name
                    .take()
                    .ok_or_else(|| invalid("external name block has no name"))?;
                let bits = current_bits
                    .take()
                    .ok_or_else(|| invalid("external name block has no properties"))?;
                match kind {
                    EXTERNAL_REFERENCE_WORKBOOK => {
                        let scope = u32::from_le_bytes([bits[2], bits[3], bits[4], bits[5]]);
                        let mut entry = XlsbExternalDefinedName::new(name)?
                            .with_built_in(bits[0] & EXTERNAL_NAME_BUILT_IN != 0);
                        if scope != 0 {
                            entry = entry
                                .with_sheet_scope(u16::try_from(scope - 1).map_err(|_| {
                                    invalid("external defined-name scope overflow")
                                })?);
                        }
                        if let Some(formula) = current_formula.take() {
                            entry = entry.with_formula(formula);
                        }
                        if workbook_entries.len() >= MAX_COLLECTION_ITEMS {
                            return Err(invalid(
                                "external-link entry collection exceeds 65,535 items",
                            ));
                        }
                        workbook_entries.push(entry);
                    },
                    EXTERNAL_REFERENCE_DDE => {
                        let mut item = XlsbDdeItem::new(name)?
                            .with_advise(bits[0] & DATA_ITEM_WANT_ADVISE != 0)
                            .with_picture(bits[0] & DATA_ITEM_WANT_PICTURE != 0)
                            .with_ole_support(bits[0] & DDE_ITEM_SUPPORTS_OLE != 0);
                        if let Some(cache) = current_cache.take() {
                            item = item.with_cached_values(cache);
                        }
                        if dde_entries.len() >= MAX_COLLECTION_ITEMS {
                            return Err(invalid(
                                "external-link entry collection exceeds 65,535 items",
                            ));
                        }
                        dde_entries.push(item);
                    },
                    EXTERNAL_REFERENCE_OLE => {
                        let mut item = XlsbOleItem::new(name)?
                            .with_advise(bits[0] & DATA_ITEM_WANT_ADVISE != 0)
                            .with_picture(bits[0] & DATA_ITEM_WANT_PICTURE != 0)
                            .with_icon(bits[0] & OLE_ITEM_DISPLAY_AS_ICON != 0);
                        if let Some(cache) = current_cache.take() {
                            item = item.with_cached_values(cache);
                        }
                        if ole_entries.len() >= MAX_COLLECTION_ITEMS {
                            return Err(invalid(
                                "external-link entry collection exceeds 65,535 items",
                            ));
                        }
                        ole_entries.push(item);
                    },
                    _ => unreachable!("external link kind was validated above"),
                }
                sup_name_state = 0;
            },
            crate::raw::kind::END_SUP_BOOK => {
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                if sup_name_state != 0 {
                    return Err(invalid(
                        "BrtEndSupBook occurs inside an external-name block",
                    ));
                }
                saw_end = true;
            },
            _ => {
                if sup_name_state == 4
                    || (link_type == Some(EXTERNAL_REFERENCE_WORKBOOK) && sup_name_state != 0)
                {
                    return Err(invalid(
                        "unexpected record inside an external name or cache",
                    ));
                }
            },
        }
    }

    let kind = link_type.ok_or_else(|| invalid("external link has no BrtBeginSupBook"))?;
    if !saw_end {
        return Err(invalid("external link has no BrtEndSupBook"));
    }
    if kind == EXTERNAL_REFERENCE_WORKBOOK && !saw_sup_tabs {
        return Err(invalid("external workbook link has no BrtSupTabs"));
    }
    let link_kind = match kind {
        EXTERNAL_REFERENCE_WORKBOOK => XlsbExternalLinkKind::Workbook,
        EXTERNAL_REFERENCE_DDE => XlsbExternalLinkKind::Dde,
        EXTERNAL_REFERENCE_OLE => XlsbExternalLinkKind::Ole,
        _ => unreachable!("external link kind was validated above"),
    };
    let relationship_id = match link_kind {
        XlsbExternalLinkKind::Dde => None,
        XlsbExternalLinkKind::Workbook | XlsbExternalLinkKind::Ole => Some(target_key.clone()),
    };
    let entries = match kind {
        EXTERNAL_REFERENCE_WORKBOOK => XlsbExternalEntries::Workbook(workbook_entries),
        EXTERNAL_REFERENCE_DDE => XlsbExternalEntries::Dde(dde_entries),
        EXTERNAL_REFERENCE_OLE => XlsbExternalEntries::Ole(ole_entries),
        _ => unreachable!("external link kind was validated above"),
    };
    let link = XlsbExternalLink {
        kind: link_kind,
        source: target_key,
        detail: match link_kind {
            XlsbExternalLinkKind::Dde | XlsbExternalLinkKind::Ole => Some(target_detail),
            XlsbExternalLinkKind::Workbook => None,
        },
        sheet_names,
        entries,
    };
    link.validate()?;
    Ok(ParsedExternalLink {
        link,
        relationship_id,
    })
}

/// Explicitly named alias for callers that need the unresolved relationship
/// metadata as well as the typed link.
pub fn parse_external_link_with_relationship(data: &[u8]) -> Result<ParsedExternalLink> {
    parse_external_link(data)
}

/// Parse a stream when the caller only needs its inert semantic model.
pub fn parse_external_link_model(data: &[u8]) -> Result<XlsbExternalLink> {
    parse_external_link(data).map(ParsedExternalLink::into_link)
}

fn parse_external_cached_value(
    record_type: crate::raw::Kind,
    data: &[u8],
    limits: crate::raw::Limits,
) -> Result<XlsbExternalCachedValue> {
    match record_type {
        crate::raw::kind::SUP_NAME_NIL if data.is_empty() => Ok(XlsbExternalCachedValue::Empty),
        crate::raw::kind::SUP_NAME_NUM if data.len() == 8 => {
            let number = f64::from_le_bytes(data.try_into().expect("length was checked"));
            validate_external_number(number)?;
            Ok(XlsbExternalCachedValue::Number(number))
        },
        crate::raw::kind::SUP_NAME_BOOL if data.len() == 1 && data[0] <= 1 => {
            Ok(XlsbExternalCachedValue::Boolean(data[0] != 0))
        },
        crate::raw::kind::SUP_NAME_ERROR if data.len() == 1 => Ok(XlsbExternalCachedValue::Error(
            XlsbExternalErrorValue::from_code(data[0])?,
        )),
        crate::raw::kind::SUP_NAME_STRING => {
            if data.len() < 4 {
                return Err(Error::InvalidLength {
                    expected: 4,
                    found: data.len(),
                });
            }
            let mut cursor = crate::raw::Cursor::with_limits(data, "BrtSupNameSt", limits);
            let value = cursor.read_wide_string()?;
            cursor.finish()?;
            Ok(XlsbExternalCachedValue::String(value))
        },
        _ => Err(invalid(format!(
            "invalid cached external value record {record_type}"
        ))),
    }
}

fn parse_external_sheet_names(data: &[u8], limits: crate::raw::Limits) -> Result<Vec<String>> {
    if data.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: data.len(),
        });
    }
    let mut cursor = crate::raw::Cursor::with_limits(data, "BrtSupTabs", limits);
    let count = usize::try_from(cursor.read_u32()?)
        .map_err(|_| invalid("external sheet-name count overflow"))?;
    if count >= MAX_COLLECTION_ITEMS {
        return Err(invalid(format!(
            "external sheet-name count {count} exceeds 65,534"
        )));
    }
    let mut names = Vec::new();
    names
        .try_reserve(count)
        .map_err(|source| Error::Allocation {
            resource: "external sheet names",
            source,
        })?;
    for _ in 0..count {
        let name = cursor.read_wide_string()?;
        let name_len = name.encode_utf16().count();
        if name_len == 0
            || name_len > 31
            || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
            || name.starts_with('\'')
            || name.ends_with('\'')
        {
            return Err(invalid(format!(
                "external sheet name {name:?} does not follow sheet-name grammar"
            )));
        }
        if names
            .iter()
            .any(|existing: &String| excel_name_eq(existing, &name))
        {
            return Err(invalid(format!("duplicate external sheet name {name:?}")));
        }
        names.push(name);
    }
    cursor.finish()?;
    Ok(names)
}

fn validate_external_name_bits(kind: u16, bits: &[u8; 7]) -> Result<()> {
    let reserved_word = &bits[2..6];
    let valid = match kind {
        EXTERNAL_REFERENCE_WORKBOOK => {
            bits[0] & EXTERNAL_NAME_RESERVED_MASK == 0
                && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG == 0
        },
        EXTERNAL_REFERENCE_DDE => {
            bits[0] & DDE_ITEM_RESERVED_MASK == 0
                && reserved_word == [0, 0, 0, 0]
                && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG != 0
        },
        EXTERNAL_REFERENCE_OLE => {
            bits[0] & OLE_ITEM_RESERVED_MASK == 0
                && bits[0] & OLE_ITEM_REQUIRED_CLASS_FLAG != 0
                && reserved_word == [0, 0, 0, 0]
                && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG != 0
        },
        _ => false,
    };
    if !valid {
        return Err(invalid(format!(
            "invalid BrtSupNameBits properties for external-link kind {kind}"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormula(message.into())
}

fn excel_name_eq(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

fn validate_unique_entries<'a>(names: impl Iterator<Item = &'a str>) -> Result<()> {
    let mut seen = HashSet::new();
    for name in names {
        if !seen.insert(name.to_lowercase()) {
            return Err(Error::InvalidFormula(format!(
                "duplicate external entry name {name:?}"
            )));
        }
    }
    Ok(())
}

fn validate_wide_string(value: &str, context: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > MAX_WIDE_STRING_UNITS || value.contains('\0') {
        return Err(Error::InvalidFormula(format!(
            "{context} is empty, too long, or contains NUL"
        )));
    }
    Ok(())
}

pub fn validate_external_number(number: f64) -> Result<()> {
    if !(number.is_normal() || (number == 0.0 && !number.is_sign_negative())) {
        return Err(Error::InvalidFormula(
            "external cache number is not a normalized finite Xnum".to_string(),
        ));
    }
    Ok(())
}

fn validate_external_name_token(
    token: &[u8],
    sheet_count: Option<usize>,
) -> Result<XlsbExternalNameFormulaKind> {
    let (&opcode, rest) = token
        .split_first()
        .ok_or_else(|| Error::InvalidFormula("external-name formula token is empty".to_string()))?;
    let (kind, expected) = match opcode {
        EXT_PTG_REFERENCE => (XlsbExternalNameFormulaKind::CellReference, 9),
        EXT_PTG_AREA => (XlsbExternalNameFormulaKind::AreaReference, 13),
        EXT_PTG_REFERENCE_ERROR => (XlsbExternalNameFormulaKind::CellReferenceError, 9),
        EXT_PTG_AREA_ERROR => (XlsbExternalNameFormulaKind::AreaReferenceError, 13),
        EXT_PTG_ERROR => (XlsbExternalNameFormulaKind::ReferenceError, 2),
        _ => {
            return Err(Error::InvalidFormula(format!(
                "unsupported external-name formula token 0x{opcode:02X}"
            )));
        },
    };
    if token.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: token.len(),
        });
    }
    if kind == XlsbExternalNameFormulaKind::ReferenceError {
        if rest != [REFERENCE_ERROR_CODE] {
            return Err(Error::InvalidFormula(
                "ExtPtgErr must contain #REF!".to_string(),
            ));
        }
        return Ok(kind);
    }
    validate_external_sheet_pair(&token[1..5], sheet_count)?;
    match kind {
        XlsbExternalNameFormulaKind::CellReference => {
            validate_small_column(u16::from_le_bytes([token[7], token[8]]))?;
        },
        XlsbExternalNameFormulaKind::AreaReference => {
            let first_row = u16::from_le_bytes([token[5], token[6]]);
            let last_row = u16::from_le_bytes([token[7], token[8]]);
            let first_column = validate_small_column(u16::from_le_bytes([token[9], token[10]]))?;
            let last_column = validate_small_column(u16::from_le_bytes([token[11], token[12]]))?;
            if last_row < first_row || last_column < first_column {
                return Err(Error::InvalidFormula(
                    "external-name area reference is reversed".to_string(),
                ));
            }
        },
        XlsbExternalNameFormulaKind::CellReferenceError
        | XlsbExternalNameFormulaKind::AreaReferenceError => {},
        XlsbExternalNameFormulaKind::ReferenceError => unreachable!(),
    }
    Ok(kind)
}

fn validate_external_sheet_pair(data: &[u8], sheet_count: Option<usize>) -> Result<()> {
    let first = i16::from_le_bytes([data[0], data[1]]);
    let last = i16::from_le_bytes([data[2], data[3]]);
    if first < -1 || last < -1 || (first == -1) != (last == -1) || (first >= 0 && last < first) {
        return Err(Error::InvalidFormula(format!(
            "invalid external sheet range {first}..={last}"
        )));
    }
    if let Some(sheet_count) = sheet_count
        && last >= 0
        && usize::try_from(last).map_or(true, |index| index >= sheet_count)
    {
        return Err(Error::InvalidFormula(
            "external-name formula sheet range exceeds BrtSupTabs".to_string(),
        ));
    }
    Ok(())
}

fn decode_sheet_range(data: &[u8]) -> XlsbExternalSheetRange {
    let first = i16::from_le_bytes([data[0], data[1]]);
    let last = i16::from_le_bytes([data[2], data[3]]);
    if first == -1 {
        XlsbExternalSheetRange::Missing
    } else {
        XlsbExternalSheetRange::Sheets {
            first: u16::try_from(first).expect("validated external sheet index"),
            last: u16::try_from(last).expect("validated external sheet index"),
        }
    }
}

fn decode_cell_location(row: &[u8], column: &[u8]) -> XlsbExternalCellLocation {
    let row = u16::from_le_bytes([row[0], row[1]]);
    let encoded_column = u16::from_le_bytes([column[0], column[1]]);
    XlsbExternalCellLocation {
        row,
        column: u8::try_from(encoded_column & 0x3FFF).expect("validated external-name column"),
        column_relative: encoded_column & (1 << 14) != 0,
        row_relative: encoded_column & (1 << 15) != 0,
    }
}

fn validate_small_column(encoded: u16) -> Result<u16> {
    let column = encoded & 0x3FFF;
    if column >= 256 {
        return Err(Error::InvalidFormula(format!(
            "external-name formula column {column} exceeds 255"
        )));
    }
    Ok(column)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::Writer;

    #[test]
    fn accepts_exactly_the_five_external_name_token_structures() {
        for (tokens, kind) in [
            (
                vec![EXT_PTG_REFERENCE, 0, 0, 0, 0, 3, 0, 2, 0],
                XlsbExternalNameFormulaKind::CellReference,
            ),
            (
                vec![EXT_PTG_AREA, 0, 0, 0, 0, 1, 0, 3, 0, 2, 0, 4, 0],
                XlsbExternalNameFormulaKind::AreaReference,
            ),
            (
                vec![EXT_PTG_REFERENCE_ERROR, 0, 0, 0, 0, 0, 0, 0, 0],
                XlsbExternalNameFormulaKind::CellReferenceError,
            ),
            (
                vec![EXT_PTG_AREA_ERROR, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                XlsbExternalNameFormulaKind::AreaReferenceError,
            ),
            (
                vec![EXT_PTG_ERROR, REFERENCE_ERROR_CODE],
                XlsbExternalNameFormulaKind::ReferenceError,
            ),
        ] {
            assert_eq!(
                XlsbExternalNameFormula::from_tokens(tokens).unwrap().kind(),
                kind
            );
        }
    }

    #[test]
    fn rejects_external_name_tokens_outside_the_restricted_grammar() {
        assert!(XlsbExternalNameFormula::from_tokens(vec![0x1E, 1, 0]).is_err());
        assert!(XlsbExternalNameFormula::from_tokens(vec![EXT_PTG_ERROR, 0x2A]).is_err());
        assert!(
            XlsbExternalNameFormula::from_tokens(vec![
                EXT_PTG_AREA,
                0,
                0,
                0,
                0,
                3,
                0,
                1,
                0,
                0,
                0,
                0,
                0,
            ])
            .is_err()
        );
    }

    #[test]
    fn typed_external_name_formula_constructors_emit_canonical_tokens() {
        let sheets = XlsbExternalSheetRange::sheets(0, 1).unwrap();
        let first = XlsbExternalCellLocation::new(2, 3)
            .with_column_relative(true)
            .with_row_relative(true);
        let last = XlsbExternalCellLocation::new(4, 5);
        let cell =
            XlsbExternalNameFormula::cell_reference(XlsbExternalCellReference::new(sheets, first));
        assert_eq!(
            cell.tokens(),
            [EXT_PTG_REFERENCE, 0, 0, 1, 0, 2, 0, 3, 0xC0]
        );
        assert_eq!(cell.cell().unwrap().location(), first);
        let area = XlsbExternalNameFormula::area_reference(
            XlsbExternalAreaReference::new(sheets, first, last).unwrap(),
        );
        assert_eq!(
            area.tokens(),
            [EXT_PTG_AREA, 0, 0, 1, 0, 2, 0, 4, 0, 3, 0xC0, 5, 0]
        );
        assert_eq!(area.area().unwrap().first(), first);
        assert_eq!(area.sheets(), Some(sheets));
        assert_eq!(
            XlsbExternalNameFormula::cell_reference_error(XlsbExternalSheetRange::Missing).kind(),
            XlsbExternalNameFormulaKind::CellReferenceError
        );
        assert_eq!(
            XlsbExternalNameFormula::area_reference_error(XlsbExternalSheetRange::Missing).kind(),
            XlsbExternalNameFormulaKind::AreaReferenceError
        );
        assert_eq!(
            XlsbExternalNameFormula::reference_error().tokens(),
            [EXT_PTG_ERROR, REFERENCE_ERROR_CODE]
        );
    }

    #[test]
    fn owner_stream_writer_and_parser_preserve_relationship_ids_and_metadata() {
        let sheets = XlsbExternalSheetRange::sheets(0, 1).unwrap();
        let formula = XlsbExternalNameFormula::cell_reference(XlsbExternalCellReference::new(
            sheets,
            XlsbExternalCellLocation::new(3, 2),
        ));
        let link = XlsbExternalLink::workbook_with_defined_names(
            "Book.xlsx",
            vec!["Data".to_string(), "Rates".to_string()],
            vec![
                XlsbExternalDefinedName::new("ExchangeRate")
                    .unwrap()
                    .with_formula(formula)
                    .with_built_in(true)
                    .with_sheet_scope(1),
            ],
        )
        .unwrap();

        let bytes = write_external_link_stream(&link, Some("rIdPath")).unwrap();
        let parsed = parse_external_link(&bytes).unwrap();
        assert_eq!(parsed.relationship_id(), Some("rIdPath"));
        assert_eq!(parsed.link().source(), "rIdPath");
        assert_eq!(parsed.link().sheet_names(), link.sheet_names());
        assert_eq!(parsed.link().defined_names(), link.defined_names());
        assert_eq!(parsed.resolve_source("Book.xlsx").unwrap(), link);
    }

    #[test]
    fn owner_stream_codec_preserves_inert_dde_and_ole_caches() {
        let cache = XlsbExternalValueMatrix::new(
            1,
            3,
            vec![
                XlsbExternalCachedValue::Number(7.0),
                XlsbExternalCachedValue::Boolean(true),
                XlsbExternalCachedValue::String("Ready".to_string()),
            ],
        )
        .unwrap();
        let dde = XlsbExternalLink::dde_with_items(
            "Excel",
            "System",
            vec![
                XlsbDdeItem::new("StatusItem")
                    .unwrap()
                    .with_advise(true)
                    .with_picture(true)
                    .with_cached_values(cache.clone()),
            ],
        )
        .unwrap();
        let dde_bytes = write_external_link_stream(&dde, None).unwrap();
        let parsed_dde = parse_external_link(&dde_bytes).unwrap().into_link();
        assert_eq!(parsed_dde, dde);

        let ole = XlsbExternalLink::ole_with_items(
            "Model.xlsx",
            "Acme.Server",
            vec![
                XlsbOleItem::new("ReportItem")
                    .unwrap()
                    .with_advise(true)
                    .with_picture(true)
                    .with_icon(true)
                    .with_cached_values(cache),
            ],
        )
        .unwrap();
        let ole_bytes = write_external_link_stream(&ole, Some("rIdOle")).unwrap();
        let parsed_ole = parse_external_link(&ole_bytes).unwrap();
        assert_eq!(parsed_ole.relationship_id(), Some("rIdOle"));
        assert_eq!(parsed_ole.resolve_source("Model.xlsx").unwrap(), ole);
    }

    #[test]
    fn owner_parser_rejects_trailing_records_after_end() {
        let link =
            XlsbExternalLink::dde("Excel", "System", vec!["StatusItem".to_string()]).unwrap();
        let mut bytes = write_external_link_stream(&link, None).unwrap();
        Writer::new(&mut bytes)
            .write_record(crate::raw::kind::SUP_NAME_END, &[])
            .unwrap();
        assert!(parse_external_link(&bytes).is_err());
    }
}
