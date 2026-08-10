#![allow(
    clippy::expect_used,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Semantic XLSB external-link values and invariants.
//!
//! The names in this module are contextual to the `external_link` owner.

use super::{
    Error, MAX_COLLECTION_ITEMS, MAX_WIDE_STRING_UNITS, MAX_XLSB_EXTERNAL_CACHE_COLUMNS,
    MAX_XLSB_EXTERNAL_CACHE_ROWS, MAX_XLSB_EXTERNAL_CACHED_VALUES, Result,
};
use std::collections::HashSet;
use std::sync::Arc;

pub(crate) const EXT_PTG_ERROR: u8 = 0x1C;
pub(crate) const EXT_PTG_REFERENCE: u8 = 0x3A;
pub(crate) const EXT_PTG_AREA: u8 = 0x3B;
pub(crate) const EXT_PTG_REFERENCE_ERROR: u8 = 0x3C;
pub(crate) const EXT_PTG_AREA_ERROR: u8 = 0x3D;
pub(crate) const REFERENCE_ERROR_CODE: u8 = 0x17;

// The owner keeps this validator local so external-name parsing does not depend on host workbook state.
pub(crate) fn validate_defined_name(name: &str) -> Result<()> {
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
pub enum Kind {
    /// A link to another workbook.
    Workbook,
    /// A link to a Dynamic Data Exchange server and topic.
    Dde,
    /// A link to an OLE data source.
    Ole,
}

/// The one token permitted in an external defined-name formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NameFormulaKind {
    CellReference,
    AreaReference,
    CellReferenceError,
    AreaReferenceError,
    ReferenceError,
}

/// Sheet range used by an external defined-name reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetRange {
    /// The referenced external sheet cannot be found.
    Missing,
    /// Inclusive zero-based external sheet indices.
    Sheets { first: u16, last: u16 },
}

impl SheetRange {
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
pub struct CellLocation {
    row: u16,
    column: u8,
    column_relative: bool,
    row_relative: bool,
}

impl CellLocation {
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
pub struct CellReference {
    sheets: SheetRange,
    location: CellLocation,
}

impl CellReference {
    pub const fn new(sheets: SheetRange, location: CellLocation) -> Self {
        Self { sheets, location }
    }

    pub const fn sheets(self) -> SheetRange {
        self.sheets
    }

    pub const fn location(self) -> CellLocation {
        self.location
    }
}

/// A typed external rectangular area reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct AreaReference {
    sheets: SheetRange,
    first: CellLocation,
    last: CellLocation,
}

impl AreaReference {
    pub fn new(sheets: SheetRange, first: CellLocation, last: CellLocation) -> Result<Self> {
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

    pub const fn sheets(self) -> SheetRange {
        self.sheets
    }

    pub const fn first(self) -> CellLocation {
        self.first
    }

    pub const fn last(self) -> CellLocation {
        self.last
    }
}

/// A validated external defined-name formula.
///
/// MS-XLSB restricts this formula to exactly one of five external Ptg token
/// structures. The original bytes are retained so relative-reference flags
/// and undefined bytes in error tokens round-trip losslessly.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NameFormula {
    kind: NameFormulaKind,
    tokens: Vec<u8>,
}

impl NameFormula {
    /// Validate and retain one external-name Ptg token.
    pub fn from_tokens(tokens: Vec<u8>) -> Result<Self> {
        let kind = validate_external_name_token(&tokens, None)?;
        Ok(Self { kind, tokens })
    }

    /// Create an external cell-reference formula without raw token bytes.
    pub fn cell_reference(reference: CellReference) -> Self {
        let mut tokens = Vec::with_capacity(9);
        tokens.push(EXT_PTG_REFERENCE);
        tokens.extend_from_slice(&reference.sheets.encode());
        tokens.extend_from_slice(&reference.location.row.to_le_bytes());
        tokens.extend_from_slice(&reference.location.encode_column());
        Self {
            kind: NameFormulaKind::CellReference,
            tokens,
        }
    }

    /// Create an external area-reference formula without raw token bytes.
    pub fn area_reference(reference: AreaReference) -> Self {
        let mut tokens = Vec::with_capacity(13);
        tokens.push(EXT_PTG_AREA);
        tokens.extend_from_slice(&reference.sheets.encode());
        tokens.extend_from_slice(&reference.first.row.to_le_bytes());
        tokens.extend_from_slice(&reference.last.row.to_le_bytes());
        tokens.extend_from_slice(&reference.first.encode_column());
        tokens.extend_from_slice(&reference.last.encode_column());
        Self {
            kind: NameFormulaKind::AreaReference,
            tokens,
        }
    }

    /// Create an invalid external cell-reference token for a known sheet range.
    pub fn cell_reference_error(sheets: SheetRange) -> Self {
        let mut tokens = Vec::with_capacity(9);
        tokens.push(EXT_PTG_REFERENCE_ERROR);
        tokens.extend_from_slice(&sheets.encode());
        tokens.extend_from_slice(&[0; 4]);
        Self {
            kind: NameFormulaKind::CellReferenceError,
            tokens,
        }
    }

    /// Create an invalid external area-reference token for a known sheet range.
    pub fn area_reference_error(sheets: SheetRange) -> Self {
        let mut tokens = Vec::with_capacity(13);
        tokens.push(EXT_PTG_AREA_ERROR);
        tokens.extend_from_slice(&sheets.encode());
        tokens.extend_from_slice(&[0; 8]);
        Self {
            kind: NameFormulaKind::AreaReferenceError,
            tokens,
        }
    }

    /// Create the standalone `#REF!` external-name formula token.
    pub fn reference_error() -> Self {
        Self {
            kind: NameFormulaKind::ReferenceError,
            tokens: vec![EXT_PTG_ERROR, REFERENCE_ERROR_CODE],
        }
    }

    pub const fn kind(&self) -> NameFormulaKind {
        self.kind
    }

    pub fn tokens(&self) -> &[u8] {
        &self.tokens
    }

    /// Return the typed cell reference when this is `CellReference`.
    pub fn cell(&self) -> Option<CellReference> {
        (self.kind == NameFormulaKind::CellReference).then(|| {
            CellReference::new(
                decode_sheet_range(&self.tokens[1..5]),
                decode_cell_location(&self.tokens[5..7], &self.tokens[7..9]),
            )
        })
    }

    /// Return the typed area reference when this is `AreaReference`.
    pub fn area(&self) -> Option<AreaReference> {
        (self.kind == NameFormulaKind::AreaReference).then(|| {
            AreaReference::new(
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
    pub fn sheets(&self) -> Option<SheetRange> {
        (self.kind != NameFormulaKind::ReferenceError)
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
pub enum ErrorValue {
    Null,
    DivisionByZero,
    Value,
    Reference,
    Name,
    Number,
    NotAvailable,
    GettingData,
}

impl ErrorValue {
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
pub enum CachedValue {
    Empty,
    Number(f64),
    Boolean(bool),
    Error(ErrorValue),
    String(String),
}

/// A bounded, row-major cached DDE/OLE value matrix.
#[derive(Debug, Clone, PartialEq)]
pub struct ValueMatrix {
    rows: u32,
    columns: u32,
    values: Vec<CachedValue>,
}

impl ValueMatrix {
    pub fn new(rows: u32, columns: u32, values: Vec<CachedValue>) -> Result<Self> {
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

    pub fn values(&self) -> &[CachedValue] {
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
                CachedValue::Number(number) => validate_number(*number)?,
                CachedValue::String(value)
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
pub struct DefinedName {
    name: String,
    formula: Option<NameFormula>,
    built_in: bool,
    scope_sheet_index: Option<u16>,
}

impl DefinedName {
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

    pub fn with_formula(mut self, formula: NameFormula) -> Self {
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

    pub fn formula(&self) -> Option<&NameFormula> {
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
pub struct DdeItem {
    name: String,
    want_advise: bool,
    want_picture: bool,
    supports_ole: bool,
    cached_values: Option<ValueMatrix>,
}

impl DdeItem {
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

    pub fn with_cached_values(mut self, values: ValueMatrix) -> Self {
        self.cached_values = Some(values);
        self
    }

    /// Remove the inert cache while retaining the item metadata.
    pub fn without_cached_values(mut self) -> Self {
        self.cached_values = None;
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

    pub fn cached_values(&self) -> Option<&ValueMatrix> {
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
pub struct OleItem {
    name: String,
    want_advise: bool,
    want_picture: bool,
    display_as_icon: bool,
    cached_values: Option<ValueMatrix>,
}

impl OleItem {
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

    pub fn with_cached_values(mut self, values: ValueMatrix) -> Self {
        self.cached_values = Some(values);
        self
    }

    /// Remove the inert cache while retaining the item metadata.
    pub fn without_cached_values(mut self) -> Self {
        self.cached_values = None;
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

    pub fn cached_values(&self) -> Option<&ValueMatrix> {
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
pub enum Entries {
    Workbook(Vec<DefinedName>),
    Dde(Vec<DdeItem>),
    Ole(Vec<OleItem>),
}

/// One unmodeled BIFF12 record retained by a source-bound snapshot.
///
/// The complete wire image is kept, including its original variable-length
/// record header. Edits therefore preserve future records and producer
/// extensions without interpreting, activating, or reserializing them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UnknownRecord {
    kind: u16,
    after_known: usize,
    bytes: Arc<[u8]>,
    payload_start: usize,
}

impl UnknownRecord {
    pub(crate) fn new(
        kind: u16,
        after_known: usize,
        bytes: Arc<[u8]>,
        payload_start: usize,
    ) -> Self {
        Self {
            kind,
            after_known,
            bytes,
            payload_start,
        }
    }

    /// Numeric BIFF12 record kind.
    #[must_use]
    pub const fn kind(&self) -> u16 {
        self.kind
    }

    /// Number of modeled records preceding this opaque record in source.
    #[must_use]
    pub const fn after_known(&self) -> usize {
        self.after_known
    }

    /// Complete source wire image, including the record header.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Borrow the opaque record payload without reconstructing its header.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.bytes[self.payload_start..]
    }
}

/// Typed metadata and inert caches from one XLSB External Link part.
#[derive(Debug, Clone, PartialEq)]
pub struct Link {
    pub kind: Kind,
    pub source: String,
    pub detail: Option<String>,
    pub sheet_names: Vec<String>,
    pub entries: Entries,
}

impl Link {
    pub fn workbook(
        source: impl Into<String>,
        sheet_names: Vec<String>,
        declared_names: Vec<String>,
    ) -> Result<Self> {
        let names = declared_names
            .into_iter()
            .map(DefinedName::new)
            .collect::<Result<Vec<_>>>()?;
        Self::workbook_with_defined_names(source, sheet_names, names)
    }

    pub fn workbook_with_defined_names(
        source: impl Into<String>,
        sheet_names: Vec<String>,
        defined_names: Vec<DefinedName>,
    ) -> Result<Self> {
        let link = Self {
            kind: Kind::Workbook,
            source: source.into(),
            detail: None,
            sheet_names,
            entries: Entries::Workbook(defined_names),
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
            .map(DdeItem::new)
            .collect::<Result<Vec<_>>>()?;
        Self::dde_with_items(server, topic, items)
    }

    pub fn dde_with_items(
        server: impl Into<String>,
        topic: impl Into<String>,
        items: Vec<DdeItem>,
    ) -> Result<Self> {
        let link = Self {
            kind: Kind::Dde,
            source: server.into(),
            detail: Some(topic.into()),
            sheet_names: Vec::new(),
            entries: Entries::Dde(items),
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
            .map(OleItem::new)
            .collect::<Result<Vec<_>>>()?;
        Self::ole_with_items(source, program_id, items)
    }

    pub fn ole_with_items(
        source: impl Into<String>,
        program_id: impl Into<String>,
        items: Vec<OleItem>,
    ) -> Result<Self> {
        let link = Self {
            kind: Kind::Ole,
            source: source.into(),
            detail: Some(program_id.into()),
            sheet_names: Vec::new(),
            entries: Entries::Ole(items),
        };
        link.validate()?;
        Ok(link)
    }

    pub const fn kind(&self) -> Kind {
        self.kind
    }

    pub fn source(&self) -> &str {
        &self.source
    }

    pub fn dde_topic(&self) -> Option<&str> {
        match self.kind {
            Kind::Dde => self.detail.as_deref(),
            Kind::Workbook | Kind::Ole => None,
        }
    }

    pub fn ole_program_id(&self) -> Option<&str> {
        match self.kind {
            Kind::Ole => self.detail.as_deref(),
            Kind::Workbook | Kind::Dde => None,
        }
    }

    pub fn sheet_names(&self) -> &[String] {
        &self.sheet_names
    }

    pub fn defined_names(&self) -> &[DefinedName] {
        match &self.entries {
            Entries::Workbook(entries) => entries,
            Entries::Dde(_) | Entries::Ole(_) => &[],
        }
    }

    pub fn dde_items(&self) -> &[DdeItem] {
        match &self.entries {
            Entries::Dde(entries) => entries,
            Entries::Workbook(_) | Entries::Ole(_) => &[],
        }
    }

    pub fn ole_items(&self) -> &[OleItem] {
        match &self.entries {
            Entries::Ole(entries) => entries,
            Entries::Workbook(_) | Entries::Dde(_) => &[],
        }
    }

    pub const fn is_workbook(&self) -> bool {
        matches!(self.kind, Kind::Workbook)
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
            (Kind::Workbook, Entries::Workbook(entries)) => {
                if self.detail.is_some() {
                    return Err(Error::InvalidFormula(
                        "external workbook link cannot have DDE/OLE detail".to_string(),
                    ));
                }
                validate_unique_entries(entries.iter().map(DefinedName::name))?;
                for entry in entries {
                    entry.validate(self.sheet_names.len(), true)?;
                }
            },
            (Kind::Dde, Entries::Dde(entries)) => {
                self.validate_data_source_detail("DDE topic")?;
                validate_unique_entries(entries.iter().map(DdeItem::name))?;
                for entry in entries {
                    entry.validate()?;
                }
            },
            (Kind::Ole, Entries::Ole(entries)) => {
                self.validate_data_source_detail("OLE program ID")?;
                validate_unique_entries(entries.iter().map(OleItem::name))?;
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
            Entries::Workbook(entries) => entries.len(),
            Entries::Dde(entries) => entries.len(),
            Entries::Ole(entries) => entries.len(),
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
pub struct Parsed {
    pub(super) link: Link,
    pub(super) relationship_id: Option<String>,
}

impl Parsed {
    /// Borrow the parsed metadata. Workbook/OLE `source()` is the unresolved
    /// relationship identifier until the host resolves it.
    pub const fn link(&self) -> &Link {
        &self.link
    }

    /// Consume the parsed value and return the unresolved metadata.
    pub fn into_link(self) -> Link {
        self.link
    }

    /// Return the relationship identifier, if this link has one.
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Replace the unresolved relationship identifier with its inert target.
    pub fn resolve_source(mut self, source: impl Into<String>) -> Result<Link> {
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

pub fn validate_number(number: f64) -> Result<()> {
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
) -> Result<NameFormulaKind> {
    let (&opcode, rest) = token
        .split_first()
        .ok_or_else(|| Error::InvalidFormula("external-name formula token is empty".to_string()))?;
    let (kind, expected) = match opcode {
        EXT_PTG_REFERENCE => (NameFormulaKind::CellReference, 9),
        EXT_PTG_AREA => (NameFormulaKind::AreaReference, 13),
        EXT_PTG_REFERENCE_ERROR => (NameFormulaKind::CellReferenceError, 9),
        EXT_PTG_AREA_ERROR => (NameFormulaKind::AreaReferenceError, 13),
        EXT_PTG_ERROR => (NameFormulaKind::ReferenceError, 2),
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
    if kind == NameFormulaKind::ReferenceError {
        if rest != [REFERENCE_ERROR_CODE] {
            return Err(Error::InvalidFormula(
                "ExtPtgErr must contain #REF!".to_string(),
            ));
        }
        return Ok(kind);
    }
    validate_external_sheet_pair(&token[1..5], sheet_count)?;
    match kind {
        NameFormulaKind::CellReference => {
            validate_small_column(u16::from_le_bytes([token[7], token[8]]))?;
        },
        NameFormulaKind::AreaReference => {
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
        NameFormulaKind::CellReferenceError | NameFormulaKind::AreaReferenceError => {},
        NameFormulaKind::ReferenceError => unreachable!(),
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

fn decode_sheet_range(data: &[u8]) -> SheetRange {
    let first = i16::from_le_bytes([data[0], data[1]]);
    let last = i16::from_le_bytes([data[2], data[3]]);
    if first == -1 {
        SheetRange::Missing
    } else {
        SheetRange::Sheets {
            first: u16::try_from(first).expect("validated external sheet index"),
            last: u16::try_from(last).expect("validated external sheet index"),
        }
    }
}

fn decode_cell_location(row: &[u8], column: &[u8]) -> CellLocation {
    let row = u16::from_le_bytes([row[0], row[1]]);
    let encoded_column = u16::from_le_bytes([column[0], column[1]]);
    CellLocation {
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
