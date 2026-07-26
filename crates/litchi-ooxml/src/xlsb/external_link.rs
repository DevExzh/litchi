//! Typed, inert XLSB External Link data (MS-XLSB 2.1.7.25).
//!
//! These values describe stored package metadata and caches. They never open
//! an external workbook, contact DDE, instantiate OLE, refresh data, evaluate
//! formulas, or execute code.

use crate::xlsb::error::{XlsbError, XlsbResult};
use std::collections::HashSet;

const MAX_COLLECTION_ITEMS: usize = 65_535;
const MAX_WIDE_STRING_UNITS: usize = 32_767;
/// Maximum row count accepted in an authored or parsed DDE/OLE cache.
pub const MAX_XLSB_EXTERNAL_CACHE_ROWS: u32 = 1_048_576;
/// Maximum column count accepted in an authored or parsed DDE/OLE cache.
pub const MAX_XLSB_EXTERNAL_CACHE_COLUMNS: u32 = 16_384;
/// Safety limit for materialized values across one DDE/OLE cache matrix.
pub const MAX_XLSB_EXTERNAL_CACHED_VALUES: usize = 1_048_576;

pub(crate) const EXTERNAL_REFERENCE_WORKBOOK: u16 = 0;
pub(crate) const EXTERNAL_REFERENCE_DDE: u16 = 1;
pub(crate) const EXTERNAL_REFERENCE_OLE: u16 = 2;
pub(crate) const EXTERNAL_NAME_BUILT_IN: u8 = 1;
pub(crate) const EXTERNAL_NAME_RESERVED_MASK: u8 = 0b0011_1110;
pub(crate) const DATA_ITEM_WANT_ADVISE: u8 = 1 << 1;
pub(crate) const DATA_ITEM_WANT_PICTURE: u8 = 1 << 2;
pub(crate) const DDE_ITEM_SUPPORTS_OLE: u8 = 1 << 3;
pub(crate) const DDE_ITEM_RESERVED_MASK: u8 = 0b0011_0001;
pub(crate) const OLE_ITEM_REQUIRED_CLASS_FLAG: u8 = 1 << 4;
pub(crate) const OLE_ITEM_DISPLAY_AS_ICON: u8 = 1 << 5;
pub(crate) const OLE_ITEM_RESERVED_MASK: u8 = 0b0000_1001;
pub(crate) const DATA_ITEM_REQUIRED_TRAILING_FLAG: u8 = 1;

const EXT_PTG_ERROR: u8 = 0x1C;
const EXT_PTG_REFERENCE: u8 = 0x3A;
const EXT_PTG_AREA: u8 = 0x3B;
const EXT_PTG_REFERENCE_ERROR: u8 = 0x3C;
const EXT_PTG_AREA_ERROR: u8 = 0x3D;
const REFERENCE_ERROR_CODE: u8 = 0x17;

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
    pub fn sheets(first: u16, last: u16) -> XlsbResult<Self> {
        if last < first || last > i16::MAX as u16 {
            return Err(XlsbError::InvalidFormula(format!(
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
    ) -> XlsbResult<Self> {
        if last.row < first.row || last.column < first.column {
            return Err(XlsbError::InvalidFormula(
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
    pub fn from_tokens(tokens: Vec<u8>) -> XlsbResult<Self> {
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

    fn validate_for_sheets(&self, sheet_count: usize) -> XlsbResult<()> {
        let kind = validate_external_name_token(&self.tokens, Some(sheet_count))?;
        if kind != self.kind {
            return Err(XlsbError::InvalidFormula(
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
    pub(crate) const fn code(self) -> u8 {
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

    pub(crate) fn from_code(code: u8) -> XlsbResult<Self> {
        match code {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0F => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1D => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2A => Ok(Self::NotAvailable),
            0x2B => Ok(Self::GettingData),
            _ => Err(XlsbError::InvalidFormula(format!(
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
    pub fn new(rows: u32, columns: u32, values: Vec<XlsbExternalCachedValue>) -> XlsbResult<Self> {
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

    fn validate(&self) -> XlsbResult<()> {
        if self.rows == 0
            || self.rows > MAX_XLSB_EXTERNAL_CACHE_ROWS
            || self.columns == 0
            || self.columns > MAX_XLSB_EXTERNAL_CACHE_COLUMNS
        {
            return Err(XlsbError::InvalidFormula(
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
                XlsbError::InvalidFormula("external cache dimensions overflow".to_string())
            })?;
        if expected != self.values.len() || expected > MAX_XLSB_EXTERNAL_CACHED_VALUES {
            return Err(XlsbError::InvalidLength {
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
                    return Err(XlsbError::InvalidFormula(
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
    pub fn new(name: impl Into<String>) -> XlsbResult<Self> {
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

    fn validate(&self, sheet_count: usize, validate_scope: bool) -> XlsbResult<()> {
        crate::xlsb::named_ranges::validate_defined_name(&self.name)?;
        if validate_scope
            && self
                .scope_sheet_index
                .is_some_and(|index| usize::from(index) >= sheet_count)
        {
            return Err(XlsbError::InvalidFormula(format!(
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
    pub fn new(name: impl Into<String>) -> XlsbResult<Self> {
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

    fn validate(&self) -> XlsbResult<()> {
        crate::xlsb::named_ranges::validate_defined_name(&self.name)?;
        if self.supports_ole && self.name != "StdDocumentName" {
            return Err(XlsbError::InvalidFormula(
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
    pub fn new(name: impl Into<String>) -> XlsbResult<Self> {
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

    fn validate(&self) -> XlsbResult<()> {
        crate::xlsb::named_ranges::validate_defined_name(&self.name)?;
        if let Some(values) = &self.cached_values {
            values.validate()?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum XlsbExternalEntries {
    Workbook(Vec<XlsbExternalDefinedName>),
    Dde(Vec<XlsbDdeItem>),
    Ole(Vec<XlsbOleItem>),
}

/// Typed metadata and inert caches from one XLSB External Link part.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsbExternalLink {
    pub(crate) kind: XlsbExternalLinkKind,
    pub(crate) source: String,
    pub(crate) detail: Option<String>,
    pub(crate) sheet_names: Vec<String>,
    pub(crate) entries: XlsbExternalEntries,
}

impl XlsbExternalLink {
    pub fn workbook(
        source: impl Into<String>,
        sheet_names: Vec<String>,
        declared_names: Vec<String>,
    ) -> XlsbResult<Self> {
        let names = declared_names
            .into_iter()
            .map(XlsbExternalDefinedName::new)
            .collect::<XlsbResult<Vec<_>>>()?;
        Self::workbook_with_defined_names(source, sheet_names, names)
    }

    pub fn workbook_with_defined_names(
        source: impl Into<String>,
        sheet_names: Vec<String>,
        defined_names: Vec<XlsbExternalDefinedName>,
    ) -> XlsbResult<Self> {
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
    ) -> XlsbResult<Self> {
        let items = item_names
            .into_iter()
            .map(XlsbDdeItem::new)
            .collect::<XlsbResult<Vec<_>>>()?;
        Self::dde_with_items(server, topic, items)
    }

    pub fn dde_with_items(
        server: impl Into<String>,
        topic: impl Into<String>,
        items: Vec<XlsbDdeItem>,
    ) -> XlsbResult<Self> {
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
    ) -> XlsbResult<Self> {
        let items = item_names
            .into_iter()
            .map(XlsbOleItem::new)
            .collect::<XlsbResult<Vec<_>>>()?;
        Self::ole_with_items(source, program_id, items)
    }

    pub fn ole_with_items(
        source: impl Into<String>,
        program_id: impl Into<String>,
        items: Vec<XlsbOleItem>,
    ) -> XlsbResult<Self> {
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

    pub(crate) fn validate(&self) -> XlsbResult<()> {
        validate_wide_string(&self.source, "external-link source")?;
        if self.sheet_names.len() > MAX_COLLECTION_ITEMS
            || self.entry_count() > MAX_COLLECTION_ITEMS
        {
            return Err(XlsbError::InvalidFormula(
                "external-link collection exceeds 65,535 items".to_string(),
            ));
        }
        let mut folded_sheet_names = HashSet::with_capacity(self.sheet_names.len());
        for sheet_name in &self.sheet_names {
            validate_wide_string(sheet_name, "external sheet name")?;
            if !folded_sheet_names.insert(sheet_name.to_lowercase()) {
                return Err(XlsbError::InvalidFormula(format!(
                    "duplicate external sheet name {sheet_name:?}"
                )));
            }
        }
        match (&self.kind, &self.entries) {
            (XlsbExternalLinkKind::Workbook, XlsbExternalEntries::Workbook(entries)) => {
                if self.detail.is_some() {
                    return Err(XlsbError::InvalidFormula(
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
                return Err(XlsbError::InvalidFormula(
                    "external-link kind does not match its entry collection".to_string(),
                ));
            },
        }
        Ok(())
    }

    pub(crate) fn entry_count(&self) -> usize {
        match &self.entries {
            XlsbExternalEntries::Workbook(entries) => entries.len(),
            XlsbExternalEntries::Dde(entries) => entries.len(),
            XlsbExternalEntries::Ole(entries) => entries.len(),
        }
    }

    fn validate_data_source_detail(&self, context: &str) -> XlsbResult<()> {
        if !self.sheet_names.is_empty() {
            return Err(XlsbError::InvalidFormula(
                "DDE/OLE links cannot declare workbook sheet names".to_string(),
            ));
        }
        validate_wide_string(self.detail.as_deref().unwrap_or_default(), context)
    }
}

fn validate_unique_entries<'a>(names: impl Iterator<Item = &'a str>) -> XlsbResult<()> {
    let mut seen = HashSet::new();
    for name in names {
        if !seen.insert(name.to_lowercase()) {
            return Err(XlsbError::InvalidFormula(format!(
                "duplicate external entry name {name:?}"
            )));
        }
    }
    Ok(())
}

fn validate_wide_string(value: &str, context: &str) -> XlsbResult<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > MAX_WIDE_STRING_UNITS || value.contains('\0') {
        return Err(XlsbError::InvalidFormula(format!(
            "{context} is empty, too long, or contains NUL"
        )));
    }
    Ok(())
}

pub(crate) fn validate_external_number(number: f64) -> XlsbResult<()> {
    if !(number.is_normal() || (number == 0.0 && !number.is_sign_negative())) {
        return Err(XlsbError::InvalidFormula(
            "external cache number is not a normalized finite Xnum".to_string(),
        ));
    }
    Ok(())
}

fn validate_external_name_token(
    token: &[u8],
    sheet_count: Option<usize>,
) -> XlsbResult<XlsbExternalNameFormulaKind> {
    let (&opcode, rest) = token.split_first().ok_or_else(|| {
        XlsbError::InvalidFormula("external-name formula token is empty".to_string())
    })?;
    let (kind, expected) = match opcode {
        EXT_PTG_REFERENCE => (XlsbExternalNameFormulaKind::CellReference, 9),
        EXT_PTG_AREA => (XlsbExternalNameFormulaKind::AreaReference, 13),
        EXT_PTG_REFERENCE_ERROR => (XlsbExternalNameFormulaKind::CellReferenceError, 9),
        EXT_PTG_AREA_ERROR => (XlsbExternalNameFormulaKind::AreaReferenceError, 13),
        EXT_PTG_ERROR => (XlsbExternalNameFormulaKind::ReferenceError, 2),
        _ => {
            return Err(XlsbError::InvalidFormula(format!(
                "unsupported external-name formula token 0x{opcode:02X}"
            )));
        },
    };
    if token.len() != expected {
        return Err(XlsbError::InvalidLength {
            expected,
            found: token.len(),
        });
    }
    if kind == XlsbExternalNameFormulaKind::ReferenceError {
        if rest != [REFERENCE_ERROR_CODE] {
            return Err(XlsbError::InvalidFormula(
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
                return Err(XlsbError::InvalidFormula(
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

fn validate_external_sheet_pair(data: &[u8], sheet_count: Option<usize>) -> XlsbResult<()> {
    let first = i16::from_le_bytes([data[0], data[1]]);
    let last = i16::from_le_bytes([data[2], data[3]]);
    if first < -1 || last < -1 || (first == -1) != (last == -1) || (first >= 0 && last < first) {
        return Err(XlsbError::InvalidFormula(format!(
            "invalid external sheet range {first}..={last}"
        )));
    }
    if let Some(sheet_count) = sheet_count
        && last >= 0
        && usize::try_from(last).map_or(true, |index| index >= sheet_count)
    {
        return Err(XlsbError::InvalidFormula(
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

fn validate_small_column(encoded: u16) -> XlsbResult<u16> {
    let column = encoded & 0x3FFF;
    if column >= 256 {
        return Err(XlsbError::InvalidFormula(format!(
            "external-name formula column {column} exceeds 255"
        )));
    }
    Ok(column)
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
