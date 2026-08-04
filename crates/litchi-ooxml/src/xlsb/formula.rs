//! XLSB formula parsing and generation
//!
//! Excel formulas in XLSB files are stored in a binary format using Reverse Polish Notation (RPN)
//! with Parse Tree Generators (Ptg tokens). This module provides parsing and generation of formulas.
//!
//! # Formula Token Types (Ptgs)
//!
//! Formulas are sequences of tokens that represent operands, operators, and functions:
//! - **Value tokens**: Numbers, strings, booleans, errors
//! - **Operand tokens**: Cell references, ranges, names
//! - **Operator tokens**: Add, subtract, multiply, divide, etc.
//! - **Function tokens**: SUM, IF, VLOOKUP, etc.
//!
//! # Binary Format
//!
//! Each token consists of:
//! 1. Token type byte (identifies the Ptg)
//! 2. Token data (variable length, depends on token type)
//!
//! # Reference
//!
//! - [MS-XLSB] Section 2.5.98 - Formulas
//! - [MS-XLS] Section 2.5.198 - Ptg (for token details, largely compatible)

mod function_table;

use crate::xlsb::error::{Error, Result};
use crate::xlsb::external_link::Link;
use function_table::BUILTIN_FUNCTIONS;
use litchi_xlsb::named_ranges::validate_name;

pub use litchi_xlsb::formula::ptg_types;
pub use litchi_xlsb::formula::{
    ArrayValue, BinaryOperator, ExternalTableReference, GroupKind, MAX_CELL_FORMULA_BYTES,
    MemoryKind, TableColumns, TableDataType, TableNamedColumns, TableReference, TableRowType,
    Token, UnaryOperator,
};

/// Compatibility parser whose historical methods continue to return the host
/// `Result` while the binary implementation is owned by `litchi-xlsb`.
pub struct FormulaParser<'a> {
    inner: litchi_xlsb::formula::FormulaParser<'a>,
}

impl<'a> FormulaParser<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Self {
            inner: litchi_xlsb::formula::FormulaParser::new(data),
        }
    }

    pub fn with_extra(data: &'a [u8], extra: &'a [u8]) -> Self {
        Self {
            inner: litchi_xlsb::formula::FormulaParser::with_extra(data, extra),
        }
    }

    pub fn with_base_cell(data: &'a [u8], row: u32, col: u32) -> Self {
        Self {
            inner: litchi_xlsb::formula::FormulaParser::with_base_cell(data, row, col),
        }
    }

    pub fn with_base_cell_and_extra(data: &'a [u8], extra: &'a [u8], row: u32, col: u32) -> Self {
        Self {
            inner: litchi_xlsb::formula::FormulaParser::with_base_cell_and_extra(
                data, extra, row, col,
            ),
        }
    }

    pub fn parse(&mut self) -> Result<Vec<Token>> {
        self.inner.parse().map_err(Into::into)
    }
}

impl From<litchi_xlsb::formula::Error> for Error {
    fn from(error: litchi_xlsb::formula::Error) -> Self {
        match error {
            litchi_xlsb::formula::Error::InvalidFormula(message) => Self::InvalidFormula(message),
            litchi_xlsb::formula::Error::InvalidCellReference(reference) => {
                Self::InvalidCellReference(reference)
            },
            litchi_xlsb::formula::Error::InvalidLength { expected, found } => {
                Self::InvalidLength { expected, found }
            },
            litchi_xlsb::formula::Error::UnsupportedFeature(feature) => {
                Self::UnsupportedFeature(feature)
            },
            litchi_xlsb::formula::Error::Encoding(message) => Self::Encoding(message),
            _ => Self::InvalidFormula("unknown formula codec error".to_string()),
        }
    }
}

/// Inclusive worksheet range used by array and shared formulas.
///
/// The host wrapper keeps the historical `Result` API while delegating
/// validation and binary conversion to the standalone XLSB codec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaRange {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl FormulaRange {
    fn from_owner(range: litchi_xlsb::formula::Range) -> Self {
        Self {
            row_first: range.row_first,
            row_last: range.row_last,
            col_first: range.col_first,
            col_last: range.col_last,
        }
    }

    fn into_owner(self) -> litchi_xlsb::formula::Range {
        litchi_xlsb::formula::Range {
            row_first: self.row_first,
            row_last: self.row_last,
            col_first: self.col_first,
            col_last: self.col_last,
        }
    }

    pub fn new(row_first: u32, row_last: u32, col_first: u32, col_last: u32) -> Result<Self> {
        litchi_xlsb::formula::Range::new(row_first, row_last, col_first, col_last)
            .map(Self::from_owner)
            .map_err(Error::from)
    }

    pub fn parse_a1(value: &str) -> Result<Self> {
        litchi_xlsb::formula::Range::parse_a1(value)
            .map(Self::from_owner)
            .map_err(Error::from)
    }

    pub fn parse_binary(data: &[u8]) -> Result<Self> {
        litchi_xlsb::formula::Range::parse_binary(data)
            .map(Self::from_owner)
            .map_err(Error::from)
    }

    pub fn to_binary(self) -> [u8; 16] {
        self.into_owner().to_binary()
    }

    pub fn contains(self, row: u32, col: u32) -> bool {
        self.into_owner().contains(row, col)
    }

    pub fn top_left(self) -> (u32, u32) {
        self.into_owner().top_left()
    }

    pub fn to_a1(self) -> String {
        self.into_owner().to_a1()
    }
}

/// The binary representation of a cell formula (`CellParsedFormula`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellParsedFormula {
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

impl CellParsedFormula {
    fn from_owner(formula: litchi_xlsb::formula::ParsedFormula) -> Self {
        Self {
            rgce: formula.rgce,
            rgcb: formula.rgcb,
        }
    }

    pub(crate) fn into_owner(self) -> litchi_xlsb::formula::ParsedFormula {
        litchi_xlsb::formula::ParsedFormula {
            rgce: self.rgce,
            rgcb: self.rgcb,
        }
    }

    pub fn parse(data: &[u8]) -> Result<(Self, usize)> {
        litchi_xlsb::formula::ParsedFormula::parse(data)
            .map(|(formula, consumed)| (Self::from_owner(formula), consumed))
            .map_err(Error::from)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        litchi_xlsb::formula::ParsedFormula {
            rgce: self.rgce.clone(),
            rgcb: self.rgcb.clone(),
        }
        .to_bytes()
        .map_err(Error::from)
    }

    pub fn exp(row: u32, col: u32) -> Result<Self> {
        litchi_xlsb::formula::ParsedFormula::exp(row, col)
            .map(Self::from_owner)
            .map_err(Error::from)
    }

    pub fn exp_cell(&self) -> Result<Option<(u32, u32)>> {
        litchi_xlsb::formula::ParsedFormula {
            rgce: self.rgce.clone(),
            rgcb: self.rgcb.clone(),
        }
        .exp_cell()
        .map_err(Error::from)
    }
}

/// Parsed `BrtArrFmla` or `BrtShrFmla` definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaGroup {
    pub kind: GroupKind,
    pub range: FormulaRange,
    pub formula: CellParsedFormula,
    pub always_calculate: bool,
}

impl FormulaGroup {
    fn from_owner(group: litchi_xlsb::formula::Group) -> Self {
        Self {
            kind: group.kind,
            range: FormulaRange::from_owner(group.range),
            formula: CellParsedFormula::from_owner(group.formula),
            always_calculate: group.always_calculate,
        }
    }

    pub fn parse_array(data: &[u8]) -> Result<Self> {
        litchi_xlsb::formula::Group::parse_array(data)
            .map(Self::from_owner)
            .map_err(Error::from)
    }

    pub fn parse_shared(data: &[u8]) -> Result<Self> {
        litchi_xlsb::formula::Group::parse_shared(data)
            .map(Self::from_owner)
            .map_err(Error::from)
    }

    pub fn to_record_data(&self) -> Result<Vec<u8>> {
        litchi_xlsb::formula::Group {
            kind: self.kind,
            range: self.range.into_owner(),
            formula: self.formula.clone().into_owner(),
            always_calculate: self.always_calculate,
        }
        .to_record_data()
        .map_err(Error::from)
    }
}

/// One entry from the workbook's `BrtExternSheet.rgXti` array.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaExternalSheet {
    pub external_link: u32,
    pub first_sheet: i32,
    pub last_sheet: i32,
}

/// Workbook table metadata required to resolve resident `PtgList` tokens.
/// A PivotTable view associated with a workbook PivotCache.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaPivotViewDefinition {
    cache_id: u32,
    sheet_index: usize,
    name: String,
}

impl FormulaPivotViewDefinition {
    pub fn try_new(cache_id: u32, sheet_index: usize, name: String) -> Result<Self> {
        validate_pivot_identifier(&name, "PivotTable view name", 255)?;
        Ok(Self {
            cache_id,
            sheet_index,
            name,
        })
    }

    pub fn cache_id(&self) -> u32 {
        self.cache_id
    }

    pub fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Aggregation encoded by `BrtBeginPName.ifn` for a calculated-field reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaPivotAggregation {
    Sum,
    CountA,
    Average,
    Max,
    Min,
    Product,
    Count,
    StandardDeviation,
    PopulationStandardDeviation,
    Variance,
    PopulationVariance,
}

impl FormulaPivotAggregation {
    fn formula_name(self) -> &'static str {
        match self {
            Self::Sum => "SUM",
            Self::CountA => "COUNTA",
            Self::Average => "AVERAGE",
            Self::Max => "MAX",
            Self::Min => "MIN",
            Self::Product => "PRODUCT",
            Self::Count => "COUNT",
            Self::StandardDeviation => "STDEV",
            Self::PopulationStandardDeviation => "STDEVP",
            Self::Variance => "VAR",
            Self::PopulationVariance => "VARP",
        }
    }
}

/// A calculated-item position encoded by a `BrtBeginPNPair` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormulaPivotItemReference {
    Name(String),
    AbsolutePosition(u32),
    RelativePosition(i32),
}

/// The formula text represented by one `BrtBeginPName` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FormulaPivotNameReference {
    Field {
        name: String,
        aggregation: Option<FormulaPivotAggregation>,
    },
    Item {
        field_name: String,
        item: FormulaPivotItemReference,
    },
}

impl FormulaPivotNameReference {
    fn validate(&self) -> Result<()> {
        match self {
            Self::Field { name, .. } => {
                validate_pivot_identifier(name, "pivot cache field name", 32_767)
            },
            Self::Item { field_name, item } => {
                validate_pivot_identifier(field_name, "pivot item field name", 32_767)?;
                match item {
                    FormulaPivotItemReference::Name(name) => {
                        validate_pivot_identifier(name, "pivot item name", 32_767)
                    },
                    FormulaPivotItemReference::AbsolutePosition(position) => {
                        if *position == 0 || *position > i32::MAX as u32 {
                            return Err(invalid(
                                "PtgSxName",
                                format!(
                                    "absolute pivot item position {position} is outside 1..={}",
                                    i32::MAX
                                ),
                            ));
                        }
                        Ok(())
                    },
                    FormulaPivotItemReference::RelativePosition(position) => {
                        if *position == 0 {
                            return Err(invalid(
                                "PtgSxName",
                                "relative pivot item position must not be zero",
                            ));
                        }
                        Ok(())
                    },
                }
            },
        }
    }

    fn to_formula_text(&self) -> String {
        match self {
            Self::Field { name, aggregation } => {
                let name = format_pivot_identifier(name);
                match aggregation {
                    Some(aggregation) => format!("{}({name})", aggregation.formula_name()),
                    None => name,
                }
            },
            Self::Item { field_name, item } => {
                let field_name = format_pivot_identifier(field_name);
                let item = match item {
                    FormulaPivotItemReference::Name(name) => format_pivot_identifier(name),
                    FormulaPivotItemReference::AbsolutePosition(position) => position.to_string(),
                    FormulaPivotItemReference::RelativePosition(position) if *position > 0 => {
                        format!("+{position}")
                    },
                    FormulaPivotItemReference::RelativePosition(position) => position.to_string(),
                };
                format!("{field_name}[{item}]")
            },
        }
    }
}

/// Formula-local `BrtBeginPName` collection for one PivotTable view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaPivotNameScope {
    cache_id: u32,
    sheet_index: usize,
    view_name: String,
    references: std::sync::Arc<[FormulaPivotNameReference]>,
}

impl FormulaPivotNameScope {
    pub fn try_new(
        cache_id: u32,
        sheet_index: usize,
        view_name: String,
        references: Vec<FormulaPivotNameReference>,
    ) -> Result<Self> {
        validate_pivot_identifier(&view_name, "PivotTable view name", 255)?;
        if references.len() > 16_384 {
            return Err(invalid(
                "BrtBeginPNames",
                format!(
                    "pivot calculated-name count {} exceeds 16384",
                    references.len()
                ),
            ));
        }
        for reference in &references {
            reference.validate()?;
        }
        Ok(Self {
            cache_id,
            sheet_index,
            view_name,
            references: references.into(),
        })
    }

    pub fn cache_id(&self) -> u32 {
        self.cache_id
    }

    pub fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub fn view_name(&self) -> &str {
        &self.view_name
    }

    pub fn references(&self) -> &[FormulaPivotNameReference] {
        &self.references
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormulaTableDefinition {
    table_id: u32,
    sheet_index: usize,
    display_name: String,
    columns: std::sync::Arc<[String]>,
}

impl FormulaTableDefinition {
    /// Build a validated table definition with columns in table-relative order.
    pub fn try_new(
        table_id: u32,
        sheet_index: usize,
        display_name: impl Into<String>,
        columns: Vec<String>,
    ) -> Result<Self> {
        if table_id == 0 || table_id == u32::MAX {
            return Err(Error::InvalidFormula(format!(
                "table identifier {table_id} is outside 1..=4294967294"
            )));
        }
        let display_name = display_name.into();
        validate_table_name(&display_name)?;
        if columns.is_empty() || columns.len() > 16_384 {
            return Err(Error::InvalidFormula(format!(
                "table {display_name:?} has {} columns, outside 1..=16384",
                columns.len()
            )));
        }
        for (index, column) in columns.iter().enumerate() {
            validate_table_column_name(column, index)?;
            if columns[..index]
                .iter()
                .any(|existing| excel_name_eq(existing, column))
            {
                return Err(Error::InvalidFormula(format!(
                    "table {display_name:?} contains duplicate column {column:?}"
                )));
            }
        }
        Ok(Self {
            table_id,
            sheet_index,
            display_name,
            columns: columns.into(),
        })
    }

    pub fn table_id(&self) -> u32 {
        self.table_id
    }

    pub fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub fn display_name(&self) -> &str {
        &self.display_name
    }

    pub fn columns(&self) -> &[String] {
        &self.columns
    }
}

/// Formula-resolution metadata from one XLSB External Link part.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct FormulaExternalBook {
    pub(crate) metadata: Link,
}

impl FormulaExternalBook {
    pub(crate) fn metadata(&self) -> Link {
        self.metadata.clone()
    }

    pub(crate) fn metadata_ref(&self) -> &Link {
        &self.metadata
    }
}

/// Kind of supporting link referenced by `BrtExternSheet` entries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FormulaSupportingLink {
    SelfWorkbook,
    SameSheet,
    ExternalWorkbook(u32),
    AddIn,
}

/// Workbook data required to render context-dependent XLSB formula tokens.
///
/// This context is owned once by the workbook and borrowed while worksheets
/// are decoded; it is never cloned per cell.
#[derive(Debug, Clone, Default)]
pub struct FormulaResolutionContext {
    pub(crate) worksheet_names: std::sync::Arc<[String]>,
    pub(crate) supporting_links: std::sync::Arc<[FormulaSupportingLink]>,
    pub(crate) external_sheets: std::sync::Arc<[FormulaExternalSheet]>,
    pub(crate) external_books: std::sync::Arc<[FormulaExternalBook]>,
    pub(crate) defined_names: std::sync::Arc<[String]>,
    pub(crate) tables: std::sync::Arc<[FormulaTableDefinition]>,
    pub(crate) pivot_views: std::sync::Arc<[FormulaPivotViewDefinition]>,
    pub(crate) pivot_name_scopes: std::sync::Arc<[FormulaPivotNameScope]>,
    pub(crate) active_pivot_scope: Option<(u32, usize, String)>,
    pub(crate) current_sheet: Option<usize>,
}

impl FormulaResolutionContext {
    pub(crate) fn for_sheet(&self, sheet_index: usize) -> Self {
        let mut context = self.clone();
        context.current_sheet = Some(sheet_index);
        context
    }

    /// Whether an XTI resolves to exactly one worksheet in this workbook.
    pub(crate) fn is_internal_single_sheet_xti(&self, index: u16) -> bool {
        let Some(xti) = self.external_sheets.get(usize::from(index)) else {
            return false;
        };
        let Some(FormulaSupportingLink::SelfWorkbook) =
            self.supporting_links.get(xti.external_link as usize)
        else {
            return false;
        };
        xti.first_sheet >= 0
            && xti.first_sheet == xti.last_sheet
            && usize::try_from(xti.first_sheet)
                .is_ok_and(|sheet| sheet < self.worksheet_names.len())
    }

    /// Bind formula-local `BrtBeginPName` metadata to an exact PivotTable view.
    pub fn for_pivot_formula(&self, scope: FormulaPivotNameScope) -> Result<Self> {
        let mut context = self.clone();
        context.current_sheet = Some(scope.sheet_index);
        context.active_pivot_scope =
            Some((scope.cache_id, scope.sheet_index, scope.view_name.clone()));
        context.pivot_name_scopes = vec![scope].into();
        context.validate_active_pivot_scope()?;
        Ok(context)
    }

    fn resolve_table_sheet(&self, index: u16) -> Result<usize> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "structured reference uses invalid Xti index 0xFFFF".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "structured-reference Xti index {index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link = self
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula("table external-link index overflow".to_string())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "structured-reference Xti {index} refers to missing supporting link {}",
                    xti.external_link
                ))
            })?;
        match link {
            FormulaSupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.first_sheet != xti.last_sheet {
                    return Err(Error::InvalidFormula(format!(
                        "structured-reference Xti {index} must select exactly one worksheet"
                    )));
                }
                let sheet = usize::try_from(xti.first_sheet).map_err(|_| {
                    Error::InvalidFormula("table worksheet index overflow".to_string())
                })?;
                if sheet >= self.worksheet_names.len() {
                    return Err(Error::InvalidFormula(format!(
                        "structured-reference worksheet {} exceeds {} worksheets",
                        xti.first_sheet,
                        self.worksheet_names.len()
                    )));
                }
                Ok(sheet)
            },
            FormulaSupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(format!(
                        "same-sheet structured-reference Xti {index} must use -2/-2"
                    )));
                }
                self.current_sheet.ok_or_else(|| {
                    Error::InvalidFormula(
                        "same-sheet structured reference has no consuming worksheet".to_string(),
                    )
                })
            },
            FormulaSupportingLink::ExternalWorkbook(_) => Err(Error::InvalidFormula(
                "resident structured reference points to an external workbook".to_string(),
            )),
            FormulaSupportingLink::AddIn => Err(Error::InvalidFormula(
                "structured reference points to an add-in".to_string(),
            )),
        }
    }

    fn resolve_external_table_prefix(&self, index: u16) -> Result<String> {
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external structured-reference Xti index {index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link = self
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula("table external-link index overflow".to_string())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external structured-reference Xti {index} has no supporting link"
                ))
            })?;
        if !matches!(link, FormulaSupportingLink::ExternalWorkbook(_)) {
            return Err(Error::InvalidFormula(
                "nonresident structured reference does not point to an external workbook"
                    .to_string(),
            ));
        }
        self.resolve_sheet_prefix(index)
    }

    fn resolve_table_reference(&self, reference: &TableReference) -> Result<String> {
        if let Some(external) = &reference.external {
            if reference.row_type.is_some()
                || reference.columns.is_some()
                || reference.list_index.is_some()
            {
                return Err(Error::InvalidFormula(
                    "nonresident structured reference also contains resident metadata".to_string(),
                ));
            }
            validate_table_name(&external.table)?;
            validate_named_table_columns(&external.columns)?;
            let prefix = self.resolve_external_table_prefix(reference.sheet_index)?;
            return Ok(format!(
                "{prefix}!{}",
                format_structured_reference(
                    &external.table,
                    external.row_type,
                    &external.columns,
                    reference.square_bracket_space,
                    reference.comma_space,
                )
            ));
        }

        let table_id = reference.list_index.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits table ID".to_string())
        })?;
        let row_type = reference.row_type.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits row type".to_string())
        })?;
        let columns = reference.columns.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits columns".to_string())
        })?;
        let sheet = self.resolve_table_sheet(reference.sheet_index)?;
        let mut matches = self
            .tables
            .iter()
            .filter(|table| table.table_id == table_id);
        let table = matches.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "structured reference names missing table ID {table_id}"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "structured reference table ID {table_id} is ambiguous"
            )));
        }
        if table.sheet_index != sheet {
            return Err(Error::InvalidFormula(format!(
                "structured reference locates table ID {table_id} on worksheet {sheet}, but metadata places it on {}",
                table.sheet_index
            )));
        }
        let named_columns = match columns {
            TableColumns::All => TableNamedColumns::All,
            TableColumns::One(index) => {
                let name = table.columns.get(usize::from(index)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference column {index} exceeds {} columns in table {:?}",
                        table.columns.len(),
                        table.display_name
                    ))
                })?;
                TableNamedColumns::One(name.clone())
            },
            TableColumns::Range { first, last } => {
                let first_name = table.columns.get(usize::from(first)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference first column {first} exceeds {} columns",
                        table.columns.len()
                    ))
                })?;
                let last_name = table.columns.get(usize::from(last)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference last column {last} exceeds {} columns",
                        table.columns.len()
                    ))
                })?;
                TableNamedColumns::Range {
                    first: first_name.clone(),
                    last: last_name.clone(),
                }
            },
        };
        Ok(format_structured_reference(
            &table.display_name,
            row_type,
            &named_columns,
            reference.square_bracket_space,
            reference.comma_space,
        ))
    }

    fn resolve_sheet_prefix(&self, index: u16) -> Result<String> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "3D reference uses invalid Xti index 0xFFFF".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "Xti index {index} exceeds {} extern-sheet entries",
                    self.external_sheets.len()
                ))
            })?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow".to_string()))?;
        let supporting_link = self.supporting_links.get(link_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti index {index} refers to missing supporting link {}",
                xti.external_link
            ))
        })?;
        let (first_index, last_index) = match supporting_link {
            FormulaSupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
                    return Err(Error::InvalidFormula(format!(
                        "Xti index {index} has invalid self-reference sheet range {}..={}",
                        xti.first_sheet, xti.last_sheet
                    )));
                }
                (
                    usize::try_from(xti.first_sheet).map_err(|_| {
                        Error::InvalidFormula("first sheet index overflow".to_string())
                    })?,
                    usize::try_from(xti.last_sheet).map_err(|_| {
                        Error::InvalidFormula("last sheet index overflow".to_string())
                    })?,
                )
            },
            FormulaSupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(format!(
                        "same-sheet Xti index {index} must use workbook scope -2/-2"
                    )));
                }
                let sheet = self.current_sheet.ok_or_else(|| {
                    Error::UnsupportedFeature(
                        "same-sheet reference requires a consuming worksheet".to_string(),
                    )
                })?;
                (sheet, sheet)
            },
            FormulaSupportingLink::ExternalWorkbook(book_index) => {
                return self.resolve_external_sheet_prefix(index, xti, *book_index);
            },
            FormulaSupportingLink::AddIn => {
                return Err(Error::UnsupportedFeature(format!(
                    "Xti index {index} refers to an add-in"
                )));
            },
        };
        if last_index < first_index {
            return Err(Error::InvalidFormula(format!(
                "Xti index {index} has invalid sheet range {}..={}",
                xti.first_sheet, xti.last_sheet
            )));
        }
        let first = self.worksheet_names.get(first_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti first sheet {} exceeds {} worksheets",
                xti.first_sheet,
                self.worksheet_names.len()
            ))
        })?;
        let last = self.worksheet_names.get(last_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti last sheet {} exceeds {} worksheets",
                xti.last_sheet,
                self.worksheet_names.len()
            ))
        })?;
        let unquoted = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&unquoted))
    }

    fn resolve_external_sheet_prefix(
        &self,
        xti_index: u16,
        xti: &FormulaExternalSheet,
        book_index: u32,
    ) -> Result<String> {
        let book_index = usize::try_from(book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow".to_string()))?;
        let book = self.external_books.get(book_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti index {xti_index} refers to missing external book {book_index}"
            ))
        })?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(format!(
                "Xti index {xti_index} refers to a DDE or OLE data source"
            )));
        }
        if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
            return Err(Error::InvalidFormula(format!(
                "Xti index {xti_index} has invalid external sheet range {}..={}",
                xti.first_sheet, xti.last_sheet
            )));
        }
        let first_index = usize::try_from(xti.first_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow".to_string()))?;
        let last_index = usize::try_from(xti.last_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow".to_string()))?;
        let first = book
            .metadata
            .sheet_names()
            .get(first_index)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external sheet {} exceeds {} cached names",
                    xti.first_sheet,
                    book.metadata.sheet_names().len()
                ))
            })?;
        let last = book.metadata.sheet_names().get(last_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external sheet {} exceeds {} cached names",
                xti.last_sheet,
                book.metadata.sheet_names().len()
            ))
        })?;
        let sheets = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&format!(
            "[{}]{sheets}",
            book.metadata.source()
        )))
    }

    fn resolve_external_name(&self, xti_index: u16, name_index: u32) -> Result<String> {
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgNameX name index is one-based and cannot be zero".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(xti_index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "PtgNameX Xti index {xti_index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow".to_string()))?;
        let FormulaSupportingLink::ExternalWorkbook(book_index) =
            self.supporting_links.get(link_index).ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "PtgNameX refers to missing supporting link {}",
                    xti.external_link
                ))
            })?
        else {
            return Err(Error::InvalidFormula(
                "PtgNameX does not refer to an external workbook".to_string(),
            ));
        };
        let external_book_index = usize::try_from(*book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow".to_string()))?;
        let book = self
            .external_books
            .get(external_book_index)
            .ok_or_else(|| Error::InvalidFormula(format!("missing external book {book_index}")))?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(
                "PtgNameX refers to a DDE or OLE data source".to_string(),
            ));
        }
        let index = usize::try_from(name_index - 1)
            .map_err(|_| Error::InvalidFormula("external name index overflow".to_string()))?;
        let names = book.metadata.defined_names();
        let name = names.get(index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external name index {name_index} exceeds {} names",
                names.len()
            ))
        })?;
        Ok(format!(
            "{}!{}",
            format_formula_prefix(&format!("[{}]", book.metadata.source())),
            name.name()
        ))
    }
}

fn owner_formula_resolution<T>(result: Result<T>) -> litchi_xlsb::formula::Result<T> {
    result.map_err(|error| match error {
        Error::InvalidFormula(message) => litchi_xlsb::formula::Error::InvalidFormula(message),
        Error::InvalidCellReference(reference) => {
            litchi_xlsb::formula::Error::InvalidCellReference(reference)
        },
        Error::InvalidLength { expected, found } => {
            litchi_xlsb::formula::Error::InvalidLength { expected, found }
        },
        Error::UnsupportedFeature(feature) => {
            litchi_xlsb::formula::Error::UnsupportedFeature(feature)
        },
        Error::Encoding(message) => litchi_xlsb::formula::Error::Encoding(message),
        error => litchi_xlsb::formula::Error::InvalidFormula(error.to_string()),
    })
}

impl litchi_xlsb::formula::FormulaResolution for FormulaResolutionContext {
    fn sheet_prefix(&self, index: u16) -> litchi_xlsb::formula::Result<String> {
        owner_formula_resolution(self.resolve_sheet_prefix(index))
    }

    fn defined_name(&self, index: u32) -> litchi_xlsb::formula::Result<String> {
        let index = usize::try_from(index)
            .ok()
            .and_then(|index| index.checked_sub(1))
            .ok_or_else(|| {
                litchi_xlsb::formula::Error::InvalidFormula(
                    "PtgName index is one-based and cannot be zero".to_string(),
                )
            })?;
        self.defined_names.get(index).cloned().ok_or_else(|| {
            litchi_xlsb::formula::Error::InvalidFormula(format!(
                "PtgName index {} exceeds {} workbook names",
                index + 1,
                self.defined_names.len()
            ))
        })
    }

    fn external_name(
        &self,
        sheet_index: u16,
        name_index: u32,
    ) -> litchi_xlsb::formula::Result<String> {
        owner_formula_resolution(self.resolve_external_name(sheet_index, name_index))
    }

    fn table_reference(
        &self,
        reference: &litchi_xlsb::formula::TableReference,
    ) -> litchi_xlsb::formula::Result<String> {
        owner_formula_resolution(self.resolve_table_reference(reference))
    }

    fn pivot_name(&self, index: u32) -> litchi_xlsb::formula::Result<String> {
        owner_formula_resolution(self.resolve_pivot_name(index))
    }
}

fn format_formula_prefix(value: &str) -> String {
    if value
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.'))
        && !value.chars().next().is_some_and(|ch| ch.is_ascii_digit())
    {
        value.to_string()
    } else {
        format!("'{}'", value.replace('\'', "''"))
    }
}

impl FormulaResolutionContext {
    fn validate_active_pivot_scope(&self) -> Result<&FormulaPivotNameScope> {
        let (cache_id, sheet_index, view_name) =
            self.active_pivot_scope.as_ref().ok_or_else(|| {
                Error::InvalidFormula(
                    "PtgSxName requires an explicit pivot cache, sheet, and view scope".to_string(),
                )
            })?;
        if *sheet_index >= self.worksheet_names.len() {
            return Err(Error::InvalidFormula(format!(
                "pivot sheet index {sheet_index} is outside the workbook sheet range"
            )));
        }
        if self.current_sheet != Some(*sheet_index) {
            return Err(Error::InvalidFormula(format!(
                "pivot scope sheet {sheet_index} does not match the formula sheet {:?}",
                self.current_sheet
            )));
        }

        let mut views = self.pivot_views.iter().filter(|view| {
            view.cache_id == *cache_id
                && view.sheet_index == *sheet_index
                && view.name.eq_ignore_ascii_case(view_name)
        });
        let _view = views.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "PivotTable view {view_name:?} on sheet {sheet_index} does not use cache {cache_id}"
            ))
        })?;
        if views.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "PivotTable view {view_name:?} on sheet {sheet_index} and cache {cache_id} is ambiguous"
            )));
        }

        let mut scopes = self.pivot_name_scopes.iter().filter(|scope| {
            scope.cache_id == *cache_id
                && scope.sheet_index == *sheet_index
                && scope.view_name.eq_ignore_ascii_case(view_name)
        });
        let scope = scopes.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "calculated-name metadata is missing for PivotTable view {view_name:?}"
            ))
        })?;
        if scopes.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "calculated-name metadata for PivotTable view {view_name:?} is ambiguous"
            )));
        }
        Ok(scope)
    }

    fn resolve_pivot_name(&self, index: u32) -> Result<String> {
        let scope = self.validate_active_pivot_scope()?;
        let index = usize::try_from(index).map_err(|_| {
            Error::InvalidFormula("pivot calculated-name index overflow".to_string())
        })?;
        let reference = scope.references.get(index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "pivot calculated-name index {index} is outside 0..{}",
                scope.references.len()
            ))
        })?;
        Ok(reference.to_formula_text())
    }
}

fn validate_pivot_identifier(name: &str, field: &str, max_utf16_len: usize) -> Result<()> {
    let utf16_len = name.encode_utf16().count();
    if utf16_len == 0 || utf16_len > max_utf16_len || name.contains('\0') {
        return Err(invalid(
            "PtgSxName",
            format!("{field} must contain 1..={max_utf16_len} UTF-16 code units and no NUL"),
        ));
    }
    Ok(())
}

fn invalid(typ: &'static str, value: impl Into<String>) -> Error {
    Error::InvalidFormula(format!("{typ}: {}", value.into()))
}

fn format_pivot_identifier(name: &str) -> String {
    if !name.eq_ignore_ascii_case("All")
        && !name.eq_ignore_ascii_case("Blank")
        && validate_table_name(name).is_ok()
    {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn validate_table_name(name: &str) -> Result<()> {
    validate_name(name)?;
    if name
        .get(..3)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("_xl"))
    {
        return Err(Error::InvalidFormula(format!(
            "table display name {name:?} uses reserved _xl prefix"
        )));
    }
    Ok(())
}

fn validate_table_column_name(name: &str, index: usize) -> Result<()> {
    let units = name.encode_utf16().count();
    if units == 0 || units > 255 || name.contains('\0') {
        return Err(Error::InvalidFormula(format!(
            "table column {index} has invalid name length or NUL content"
        )));
    }
    Ok(())
}

fn validate_named_table_columns(columns: &TableNamedColumns) -> Result<()> {
    match columns {
        TableNamedColumns::All => Ok(()),
        TableNamedColumns::One(name) => validate_table_column_name(name, 0),
        TableNamedColumns::Range { first, last } => {
            validate_table_column_name(first, 0)?;
            validate_table_column_name(last, 1)
        },
    }
}

fn escape_structured_column(name: &str) -> String {
    let mut escaped = String::with_capacity(name.len());
    for ch in name.chars() {
        if matches!(ch, '#' | '[' | ']' | '\'' | '@') {
            escaped.push('\'');
        }
        escaped.push(ch);
    }
    escaped
}

fn format_structured_reference(
    table: &str,
    row_type: TableRowType,
    columns: &TableNamedColumns,
    square_bracket_space: bool,
    comma_space: bool,
) -> String {
    let mut items = Vec::new();
    match row_type {
        TableRowType::Data => {},
        TableRowType::All => items.push("[#All]".to_string()),
        TableRowType::Headers => items.push("[#Headers]".to_string()),
        TableRowType::DataAlternate => items.push("[#Data]".to_string()),
        TableRowType::DataAndHeaders => {
            items.push("[#Headers]".to_string());
            items.push("[#Data]".to_string());
        },
        TableRowType::Totals => items.push("[#Totals]".to_string()),
        TableRowType::DataAndTotals => {
            items.push("[#Data]".to_string());
            items.push("[#Totals]".to_string());
        },
        TableRowType::Current => items.push("[#This Row]".to_string()),
    }
    let has_range = matches!(columns, TableNamedColumns::Range { .. });
    match columns {
        TableNamedColumns::All => {},
        TableNamedColumns::One(name) => {
            items.push(format!("[{}]", escape_structured_column(name)));
        },
        TableNamedColumns::Range { first, last } => {
            items.push(format!(
                "[{}]:[{}]",
                escape_structured_column(first),
                escape_structured_column(last)
            ));
        },
    }
    if items.is_empty() {
        return table.to_string();
    }
    let separator = if comma_space { ", " } else { "," };
    let body = items.join(separator);
    let specifiers = if items.len() == 1 && !has_range {
        if square_bracket_space {
            format!("[ {} ]", &body[1..body.len() - 1])
        } else {
            body
        }
    } else if square_bracket_space {
        format!("[ {body} ]")
    } else {
        format!("[{body}]")
    };
    format!("{table}{specifiers}")
}

/// Inclusive worksheet range used by array and shared formulas.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]

pub struct FormulaConverter;

impl FormulaConverter {
    /// Convert formula tokens to string representation
    ///
    /// Uses RPN to infix conversion with proper operator precedence.
    pub fn tokens_to_string(tokens: &[Token]) -> String {
        Self::try_tokens_to_string(tokens).unwrap_or_default()
    }

    /// Convert tokens to text, rejecting token streams that cannot be
    /// represented faithfully by this converter.
    pub fn try_tokens_to_string(tokens: &[Token]) -> Result<String> {
        Self::try_tokens_to_string_with_optional_context(tokens, None)
    }

    /// Convert formula tokens using workbook extern-sheet and name metadata.
    pub fn try_tokens_to_string_with_context(
        tokens: &[Token],
        context: &FormulaResolutionContext,
    ) -> Result<String> {
        Self::try_tokens_to_string_with_optional_context(tokens, Some(context))
    }

    fn try_tokens_to_string_with_optional_context(
        tokens: &[Token],
        context: Option<&FormulaResolutionContext>,
    ) -> Result<String> {
        let mut stack: Vec<String> = Vec::new();

        for token in tokens {
            match token {
                Token::Number(n) => stack.push(format!("{}", n)),
                Token::Int(i) => stack.push(format!("{}", i)),
                Token::MissingArg => stack.push(String::new()),
                Token::Parenthesis => {
                    let Some(expression) = stack.pop() else {
                        return Err(Error::InvalidFormula(
                            "PtgParen has no preceding expression".to_string(),
                        ));
                    };
                    stack.push(format!("({expression})"));
                },
                Token::Attribute(_) => {},
                Token::Array { rows, cols, values } => {
                    let expected = usize::try_from(u64::from(*rows) * u64::from(*cols))
                        .map_err(|_| Error::InvalidFormula("array is too large".to_string()))?;
                    if values.len() != expected {
                        return Err(Error::InvalidFormula(format!(
                            "array dimensions require {expected} values, found {}",
                            values.len()
                        )));
                    }
                    let mut text = String::from("{");
                    for row in 0..*rows {
                        if row != 0 {
                            text.push(';');
                        }
                        for col in 0..*cols {
                            if col != 0 {
                                text.push(',');
                            }
                            let index =
                                usize::try_from(u64::from(row) * u64::from(*cols) + u64::from(col))
                                    .map_err(|_| {
                                        Error::InvalidFormula("array index overflow".to_string())
                                    })?;
                            match &values[index] {
                                ArrayValue::Number(value) => {
                                    text.push_str(&value.to_string());
                                },
                                ArrayValue::String(value) => {
                                    text.push('"');
                                    text.push_str(&value.replace('"', "\"\""));
                                    text.push('"');
                                },
                                ArrayValue::Bool(value) => {
                                    text.push_str(if *value { "TRUE" } else { "FALSE" });
                                },
                                ArrayValue::Error(error) => {
                                    text.push_str(&Self::error_to_string(*error));
                                },
                            }
                        }
                    }
                    text.push('}');
                    stack.push(text);
                },
                Token::Memory { .. } => {},
                Token::String(s) => stack.push(format!("\"{}\"", s.replace('"', "\"\""))),
                Token::Bool(b) => stack.push(if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }),
                Token::Error(e) => stack.push(Self::error_to_string(*e)),
                Token::CellRef {
                    row,
                    col,
                    row_relative,
                    col_relative,
                } => {
                    let col_str = crate::xlsb::utils::column_index_to_name(*col + 1);
                    let row_str = row + 1;
                    let col_prefix = if *col_relative { "" } else { "$" };
                    let row_prefix = if *row_relative { "" } else { "$" };
                    stack.push(format!(
                        "{}{}{}{}",
                        col_prefix, col_str, row_prefix, row_str
                    ));
                },
                Token::AreaRef {
                    row_first,
                    col_first,
                    row_last,
                    col_last,
                    row_first_relative,
                    row_last_relative,
                    col_first_relative,
                    col_last_relative,
                } => {
                    let first = Self::format_reference(
                        *row_first,
                        *col_first,
                        *row_first_relative,
                        *col_first_relative,
                    );
                    let last = Self::format_reference(
                        *row_last,
                        *col_last,
                        *row_last_relative,
                        *col_last_relative,
                    );
                    stack.push(format!("{}:{}", first, last));
                },
                Token::CellRef3d {
                    sheet_index,
                    row,
                    col,
                    row_relative,
                    col_relative,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgRef3d requires workbook extern-sheet resolution".to_string(),
                        )
                    })?;
                    let prefix = context.resolve_sheet_prefix(*sheet_index)?;
                    let reference =
                        Self::format_reference(*row, *col, *row_relative, *col_relative);
                    stack.push(format!("{prefix}!{reference}"));
                },
                Token::AreaRef3d {
                    sheet_index,
                    row_first,
                    row_last,
                    col_first,
                    col_last,
                    row_first_relative,
                    row_last_relative,
                    col_first_relative,
                    col_last_relative,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgArea3d requires workbook extern-sheet resolution".to_string(),
                        )
                    })?;
                    let prefix = context.resolve_sheet_prefix(*sheet_index)?;
                    let first = Self::format_reference(
                        *row_first,
                        *col_first,
                        *row_first_relative,
                        *col_first_relative,
                    );
                    let last = Self::format_reference(
                        *row_last,
                        *col_last,
                        *row_last_relative,
                        *col_last_relative,
                    );
                    stack.push(format!("{prefix}!{first}:{last}"));
                },
                Token::ReferenceError { .. } => stack.push("#REF!".to_string()),
                Token::BinaryOp(op) => {
                    if stack.len() < 2 {
                        return Err(Error::InvalidFormula(
                            "binary operator has fewer than two operands".to_string(),
                        ));
                    }
                    let right = stack.pop().expect("length checked");
                    let left = stack.pop().expect("length checked");
                    let op_str = Self::binary_op_to_string(*op);
                    stack.push(format!("({}{}{})", left, op_str, right));
                },
                Token::UnaryOp(op) => {
                    let Some(operand) = stack.pop() else {
                        return Err(Error::InvalidFormula(
                            "unary operator has no operand".to_string(),
                        ));
                    };
                    match op {
                        UnaryOperator::Plus => stack.push(format!("+({})", operand)),
                        UnaryOperator::Minus => stack.push(format!("-({})", operand)),
                        UnaryOperator::Percent => stack.push(format!("({}%)", operand)),
                    }
                },
                Token::Function {
                    index,
                    arg_count,
                    is_command,
                } => {
                    if *is_command {
                        return Err(Error::UnsupportedFeature(format!(
                            "XLSB command function index {index}"
                        )));
                    }
                    let Some(function) = builtin_function_by_index(*index) else {
                        return Err(Error::UnsupportedFeature(format!(
                            "XLSB built-in function index {index}"
                        )));
                    };
                    let func_name = function.name;
                    if stack.len() < usize::from(*arg_count) {
                        return Err(Error::InvalidFormula(format!(
                            "function {func_name} requires {arg_count} stack operands"
                        )));
                    }
                    let mut args = Vec::new();
                    for _ in 0..*arg_count {
                        if let Some(arg) = stack.pop() {
                            args.insert(0, arg);
                        }
                    }
                    stack.push(format!("{}({})", func_name, args.join(",")));
                },
                Token::Name(idx) => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(format!(
                            "XLSB defined name index {idx} requires workbook name resolution"
                        ))
                    })?;
                    let index = usize::try_from(*idx - 1)
                        .map_err(|_| Error::InvalidFormula("PtgName index overflow".to_string()))?;
                    let name = context.defined_names.get(index).ok_or_else(|| {
                        Error::InvalidFormula(format!(
                            "PtgName index {idx} exceeds {} workbook names",
                            context.defined_names.len()
                        ))
                    })?;
                    stack.push(name.clone());
                },
                Token::ExternalName {
                    sheet_index,
                    name_index,
                } => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(
                            "PtgNameX requires workbook external-link resolution".to_string(),
                        )
                    })?;
                    stack.push(context.resolve_external_name(*sheet_index, *name_index)?);
                },
                Token::TableReference(reference) if reference.invalid => {
                    stack.push("#REF!".to_string())
                },
                Token::TableReference(reference) => {
                    let context = context.ok_or_else(|| {
                        Error::UnsupportedFeature(format!(
                            "structured table reference on Xti {} requires table-definition resolution",
                            reference.sheet_index
                        ))
                    })?;
                    stack.push(context.resolve_table_reference(reference)?);
                },
                Token::PivotName(index) => {
                    let context = context.ok_or_else(|| {
                        Error::InvalidFormula(
                            "PtgSxName requires pivot-cache calculated-name metadata".to_string(),
                        )
                    })?;
                    stack.push(context.resolve_pivot_name(*index)?);
                },
                Token::Unknown(t) => {
                    return Err(Error::UnsupportedFeature(format!(
                        "XLSB formula token 0x{t:02X}"
                    )));
                },
            }
        }

        if stack.len() != 1 {
            return Err(Error::InvalidFormula(format!(
                "formula leaves {} values on the evaluation stack",
                stack.len()
            )));
        }
        Ok(stack.pop().expect("length checked"))
    }

    fn format_reference(row: u32, col: u32, row_relative: bool, col_relative: bool) -> String {
        let col_str = crate::xlsb::utils::column_index_to_name(col + 1);
        format!(
            "{}{}{}{}",
            if col_relative { "" } else { "$" },
            col_str,
            if row_relative { "" } else { "$" },
            row + 1
        )
    }

    /// Convert binary operator to string
    fn binary_op_to_string(op: BinaryOperator) -> &'static str {
        match op {
            BinaryOperator::Add => "+",
            BinaryOperator::Subtract => "-",
            BinaryOperator::Multiply => "*",
            BinaryOperator::Divide => "/",
            BinaryOperator::Power => "^",
            BinaryOperator::Concat => "&",
            BinaryOperator::LessThan => "<",
            BinaryOperator::LessEqual => "<=",
            BinaryOperator::Equal => "=",
            BinaryOperator::GreaterEqual => ">=",
            BinaryOperator::GreaterThan => ">",
            BinaryOperator::NotEqual => "<>",
            BinaryOperator::Intersection => " ",
            BinaryOperator::Union => ",",
            BinaryOperator::Range => ":",
        }
    }

    /// Convert error code to string
    fn error_to_string(code: u8) -> String {
        match code {
            0x00 => "#NULL!".to_string(),
            0x07 => "#DIV/0!".to_string(),
            0x0F => "#VALUE!".to_string(),
            0x17 => "#REF!".to_string(),
            0x1D => "#NAME?".to_string(),
            0x24 => "#NUM!".to_string(),
            0x2A => "#N/A".to_string(),
            0x2B => "#GETTING_DATA".to_string(),
            _ => format!("#ERR{:02X}!", code),
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct BuiltinFunction {
    index: u16,
    name: &'static str,
    min_args: u8,
    max_args: u8,
}

impl BuiltinFunction {
    fn accepts_arg_count(self, count: u8) -> bool {
        if count < self.min_args || count > self.max_args {
            return false;
        }
        match self.index {
            // GETPIVOTDATA permits the two mandatory arguments, one optional
            // field, or complete field/item pairs thereafter.
            358 => count <= 3 || count.is_multiple_of(2),
            // COUNTIFS is made solely of range/criteria pairs.
            481 => count.is_multiple_of(2),
            // SUMIFS and AVERAGEIFS have one leading aggregate range followed
            // by range/criteria pairs.
            482 | 484 => !count.is_multiple_of(2),
            _ => true,
        }
    }
}

fn builtin_function_by_index(index: u16) -> Option<BuiltinFunction> {
    let position = BUILTIN_FUNCTIONS
        .binary_search_by_key(&index, |entry| entry.0)
        .ok()?;
    let (index, name, min_args, max_args) = BUILTIN_FUNCTIONS[position];
    Some(BuiltinFunction {
        index,
        name,
        min_args,
        max_args,
    })
}

fn builtin_function_by_name(name: &str) -> Option<BuiltinFunction> {
    BUILTIN_FUNCTIONS
        .iter()
        .find_map(|&(index, function_name, min_args, max_args)| {
            function_name
                .eq_ignore_ascii_case(name)
                .then_some(BuiltinFunction {
                    index,
                    name: function_name,
                    min_args,
                    max_args,
                })
        })
}

/// Compiles standards-defined Excel formula text to XLSB RPN tokens.
///
/// The compiler supports literals, A1 references and ranges, parentheses,
/// arithmetic/comparison/concatenation operators, percent, and the built-in
/// non-macro built-in functions from [MS-XLSB]'s `Ftab` table, and typed array
/// constants. Unsupported constructs return an error; they are never replaced
/// by a cached value.
pub struct FormulaCompiler<'a> {
    input: &'a str,
    offset: usize,
    context: Option<&'a FormulaCompilationContext<'a>>,
}

/// A defined name visible to the XLSB formula text compiler.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct FormulaDefinedName {
    pub(crate) name: String,
    pub(crate) sheet_id: Option<u32>,
}

/// Workbook metadata used to compile context-dependent formula operands.
#[derive(Debug, Clone, Copy)]
pub(crate) struct FormulaCompilationContext<'a> {
    pub(crate) worksheet_names: &'a [String],
    pub(crate) defined_names: &'a [FormulaDefinedName],
    pub(crate) tables: &'a [FormulaTableDefinition],
    pub(crate) supporting_links: &'a [FormulaSupportingLink],
    pub(crate) external_sheets: &'a [FormulaExternalSheet],
    pub(crate) external_books: &'a [FormulaExternalBook],
    pub(crate) sheet_ranges: &'a std::cell::RefCell<Vec<(u32, u32)>>,
    pub(crate) current_sheet: u32,
}

#[derive(Debug, Clone, Copy)]
enum FormulaEncoding {
    Cell,
    Shared { base_row: u32, base_col: u32 },
}

#[derive(Debug)]
enum CompileExpr {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
    MissingArg,
    Parenthesized(Box<CompileExpr>),
    Array {
        rows: u32,
        cols: u32,
        values: Vec<ArrayValue>,
    },
    Ref(A1Reference),
    Area(A1Reference, A1Reference),
    Ref3d(u16, A1Reference),
    Area3d(u16, A1Reference, A1Reference),
    Name(u32),
    TableReference(TableReference),
    Unary(UnaryOperator, Box<CompileExpr>),
    Binary(BinaryOperator, Box<CompileExpr>, Box<CompileExpr>),
    Function(BuiltinFunction, Vec<CompileExpr>),
}

#[derive(Debug)]
struct ParsedStructuredReference {
    row_type: TableRowType,
    columns: TableNamedColumns,
    square_bracket_space: bool,
    comma_space: bool,
}

#[derive(Debug)]
struct StructuredReferenceItem {
    text: String,
    first_character_escaped: bool,
}

#[derive(Debug, Clone, Copy)]
struct A1Reference {
    row: u32,
    col: u32,
    row_relative: bool,
    col_relative: bool,
}

impl<'a> FormulaCompiler<'a> {
    pub fn compile(formula: &'a str) -> Result<CellParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell, None)
    }

    pub(crate) fn compile_with_context(
        formula: &'a str,
        context: &'a FormulaCompilationContext<'a>,
    ) -> Result<CellParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell, Some(context))
    }

    /// Compile a shared formula, encoding relative A1 references as
    /// `PtgRefN`/`PtgAreaN` offsets from the first cell in the shared range.
    pub fn compile_shared(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
    ) -> Result<CellParsedFormula> {
        Self::compile_shared_with_optional_context(formula, base_row, base_col, None)
    }

    pub(crate) fn compile_shared_with_context(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
        context: &'a FormulaCompilationContext<'a>,
    ) -> Result<CellParsedFormula> {
        Self::compile_shared_with_optional_context(formula, base_row, base_col, Some(context))
    }

    fn compile_shared_with_optional_context(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
        context: Option<&'a FormulaCompilationContext<'a>>,
    ) -> Result<CellParsedFormula> {
        if base_row >= 1_048_576 || base_col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "shared formula base ({base_row}, {base_col})"
            )));
        }
        Self::compile_with_encoding(
            formula,
            FormulaEncoding::Shared { base_row, base_col },
            context,
        )
    }

    fn compile_with_encoding(
        formula: &'a str,
        encoding: FormulaEncoding,
        context: Option<&'a FormulaCompilationContext<'a>>,
    ) -> Result<CellParsedFormula> {
        let input = formula.strip_prefix('=').unwrap_or(formula).trim();
        if input.is_empty() {
            return Err(Error::InvalidFormula(
                "formula expression is empty".to_string(),
            ));
        }
        let mut compiler = Self {
            input,
            offset: 0,
            context,
        };
        let expression = compiler.parse_comparison()?;
        compiler.skip_spaces();
        if compiler.offset != compiler.input.len() {
            return Err(compiler.error("unexpected trailing input"));
        }

        let mut rgce = Vec::new();
        let mut rgcb = Vec::new();
        Self::emit(&expression, &mut rgce, &mut rgcb, encoding)?;
        if rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "compiled formula is {} bytes; maximum is {MAX_CELL_FORMULA_BYTES}",
                rgce.len()
            )));
        }
        Ok(CellParsedFormula { rgce, rgcb })
    }

    fn parse_comparison(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_concat()?;
        loop {
            let operator = if self.consume("<>") {
                Some(BinaryOperator::NotEqual)
            } else if self.consume("<=") {
                Some(BinaryOperator::LessEqual)
            } else if self.consume(">=") {
                Some(BinaryOperator::GreaterEqual)
            } else if self.consume("=") {
                Some(BinaryOperator::Equal)
            } else if self.consume("<") {
                Some(BinaryOperator::LessThan)
            } else if self.consume(">") {
                Some(BinaryOperator::GreaterThan)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_concat()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_concat(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_additive()?;
        while self.consume("&") {
            let right = self.parse_additive()?;
            expression = CompileExpr::Binary(
                BinaryOperator::Concat,
                Box::new(expression),
                Box::new(right),
            );
        }
        Ok(expression)
    }

    fn parse_additive(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_multiplicative()?;
        loop {
            let operator = if self.consume("+") {
                Some(BinaryOperator::Add)
            } else if self.consume("-") {
                Some(BinaryOperator::Subtract)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_multiplicative()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_multiplicative(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_power()?;
        loop {
            let operator = if self.consume("*") {
                Some(BinaryOperator::Multiply)
            } else if self.consume("/") {
                Some(BinaryOperator::Divide)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_power()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_power(&mut self) -> Result<CompileExpr> {
        let left = self.parse_unary()?;
        if self.consume("^") {
            let right = self.parse_power()?;
            Ok(CompileExpr::Binary(
                BinaryOperator::Power,
                Box::new(left),
                Box::new(right),
            ))
        } else {
            Ok(left)
        }
    }

    fn parse_unary(&mut self) -> Result<CompileExpr> {
        if self.consume("+") {
            return Ok(CompileExpr::Unary(
                UnaryOperator::Plus,
                Box::new(self.parse_unary()?),
            ));
        }
        if self.consume("-") {
            return Ok(CompileExpr::Unary(
                UnaryOperator::Minus,
                Box::new(self.parse_unary()?),
            ));
        }
        let mut expression = self.parse_primary()?;
        while self.consume("%") {
            expression = CompileExpr::Unary(UnaryOperator::Percent, Box::new(expression));
        }
        Ok(expression)
    }

    fn parse_primary(&mut self) -> Result<CompileExpr> {
        self.skip_spaces();
        if self.consume("(") {
            let expression = self.parse_comparison()?;
            if !self.consume(")") {
                return Err(self.error("expected ')'"));
            }
            return Ok(CompileExpr::Parenthesized(Box::new(expression)));
        }
        if self.consume("{") {
            return self.parse_array_constant();
        }
        if self.peek_char() == Some('"') {
            return self.parse_string().map(CompileExpr::String);
        }
        if self.peek_char() == Some('#') {
            return self.parse_error_literal().map(CompileExpr::Error);
        }
        if self
            .peek_char()
            .is_some_and(|ch| ch.is_ascii_digit() || ch == '.')
        {
            return self.parse_number().map(CompileExpr::Number);
        }

        if self.peek_char() == Some('\'') {
            let sheet_qualifier = self.parse_quoted_sheet_name()?;
            if !self.consume("!") {
                return Err(self.error("expected '!' after quoted worksheet name"));
            }
            if sheet_qualifier.starts_with('[') {
                let table = self.parse_identifier()?;
                if self.peek_char() == Some('[') || parse_a1_reference(&table).is_none() {
                    let selection = if self.peek_char() == Some('[') {
                        self.parse_structured_reference()?
                    } else {
                        ParsedStructuredReference {
                            row_type: TableRowType::Data,
                            columns: TableNamedColumns::All,
                            square_bracket_space: false,
                            comma_space: false,
                        }
                    };
                    return self.compile_external_table_reference(
                        &sheet_qualifier,
                        table,
                        selection,
                    );
                }
                return Err(self.error(
                    "external cell references are not supported by this compilation context",
                ));
            }
            let (first_sheet, last_sheet) = Self::split_sheet_qualifier(&sheet_qualifier)?;
            return self.parse_qualified_reference(first_sheet, last_sheet);
        }

        let identifier = self.parse_identifier()?;
        let sheet_range_checkpoint = self.offset;
        if self.consume(":") {
            let last_sheet = self.parse_identifier()?;
            if self.consume("!") {
                return self.parse_qualified_reference(&identifier, Some(&last_sheet));
            }
            self.offset = sheet_range_checkpoint;
        }
        if self.consume("!") {
            return self.parse_qualified_reference(&identifier, None);
        }
        if self.peek_char() == Some('[') {
            let selection = self.parse_structured_reference()?;
            return self.compile_resident_table_reference(&identifier, selection);
        }
        if self.consume("(") {
            let function = builtin_function_by_name(&identifier).ok_or_else(|| {
                Error::UnsupportedFeature(format!(
                    "XLSB formula function {identifier} is not in the supported Ftab set"
                ))
            })?;
            let mut arguments = Vec::new();
            if !self.consume(")") {
                loop {
                    if self.consume(")") {
                        arguments.push(CompileExpr::MissingArg);
                        break;
                    }
                    if self.consume(",") {
                        arguments.push(CompileExpr::MissingArg);
                        continue;
                    }
                    arguments.push(self.parse_comparison()?);
                    if self.consume(")") {
                        break;
                    }
                    if !self.consume(",") {
                        return Err(self.error("expected ',' or ')' in function call"));
                    }
                }
            }
            let argument_count = u8::try_from(arguments.len()).map_err(|_| {
                Error::InvalidFormula(format!("{} has more than 255 arguments", function.name))
            })?;
            if !function.accepts_arg_count(argument_count) {
                return Err(Error::InvalidFormula(format!(
                    "{} does not accept {} arguments (range {}..={})",
                    function.name,
                    arguments.len(),
                    function.min_args,
                    function.max_args,
                )));
            }
            return Ok(CompileExpr::Function(function, arguments));
        }
        if identifier.eq_ignore_ascii_case("TRUE") {
            return Ok(CompileExpr::Bool(true));
        }
        if identifier.eq_ignore_ascii_case("FALSE") {
            return Ok(CompileExpr::Bool(false));
        }

        if let Some(reference) = self.compile_bare_resident_table_reference(&identifier)? {
            return Ok(reference);
        }

        let Some(first) = parse_a1_reference(&identifier) else {
            return self
                .resolve_defined_name(&identifier)
                .map(CompileExpr::Name);
        };
        if self.consume(":") {
            let second_text = self.parse_identifier()?;
            let second = parse_a1_reference(&second_text)
                .ok_or_else(|| self.error("invalid range end reference"))?;
            Ok(CompileExpr::Area(first, second))
        } else {
            Ok(CompileExpr::Ref(first))
        }
    }

    fn parse_structured_reference(&mut self) -> Result<ParsedStructuredReference> {
        debug_assert_eq!(self.peek_char(), Some('['));
        self.offset += 1;
        let leading_space = self.consume_structured_space()?;
        let nested = self.peek_char() == Some('[');
        let mut items = Vec::new();
        let mut separators = Vec::new();
        let mut comma_space = None;
        let mut unwrapped_trailing_space = false;

        if nested {
            loop {
                items.push(self.parse_structured_reference_item(true)?);
                match self.peek_char() {
                    Some(',') => {
                        self.offset += 1;
                        separators.push(',');
                        let spaced = self.consume_structured_space()?;
                        if comma_space
                            .replace(spaced)
                            .is_some_and(|previous| previous != spaced)
                        {
                            return Err(self
                                .error("structured-reference commas use inconsistent whitespace"));
                        }
                    },
                    Some(':') => {
                        self.offset += 1;
                        separators.push(':');
                    },
                    _ => break,
                }
            }
        } else {
            let mut item = self.parse_structured_reference_item(false)?;
            if item.text.ends_with(char::is_whitespace) {
                if !item.text.ends_with(' ')
                    || item
                        .text
                        .strip_suffix(' ')
                        .is_some_and(|text| text.ends_with(char::is_whitespace))
                {
                    return Err(self.error(
                        "structured-reference whitespace cannot be represented by XLSB flags",
                    ));
                }
                item.text.pop();
                if item.text.is_empty() {
                    return Err(self.error("structured-reference item is empty"));
                }
                unwrapped_trailing_space = true;
            }
            items.push(item);
        }

        let trailing_space = if nested {
            self.consume_structured_space()?
        } else {
            unwrapped_trailing_space
        };
        if leading_space != trailing_space {
            return Err(self.error("structured-reference square-bracket whitespace is asymmetric"));
        }
        if self.peek_char() != Some(']') {
            return Err(self.error("expected closing structured-reference bracket"));
        }
        self.offset += 1;
        if nested && items.len() == 1 {
            return Err(self
                .error("redundant nested structured reference cannot be represented faithfully"));
        }

        let (row_type, columns) = Self::classify_structured_reference(items, &separators)?;
        Ok(ParsedStructuredReference {
            row_type,
            columns,
            square_bracket_space: leading_space,
            comma_space: comma_space.unwrap_or(false),
        })
    }

    fn parse_structured_reference_item(
        &mut self,
        bracketed: bool,
    ) -> Result<StructuredReferenceItem> {
        if bracketed {
            if self.peek_char() != Some('[') {
                return Err(self.error("expected nested structured-reference item"));
            }
            self.offset += 1;
        }
        let mut text = String::new();
        let mut first_character_escaped = false;
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated structured reference"));
            };
            if ch == ']' {
                if bracketed {
                    self.offset += 1;
                }
                break;
            }
            self.offset += ch.len_utf8();
            if ch == '\'' {
                let Some(escaped) = self.peek_char() else {
                    return Err(self.error("unterminated structured-reference escape"));
                };
                if !matches!(escaped, '#' | '[' | ']' | '\'' | '@') {
                    return Err(self.error("invalid structured-reference escape"));
                }
                if text.is_empty() {
                    first_character_escaped = true;
                }
                self.offset += escaped.len_utf8();
                text.push(escaped);
            } else {
                text.push(ch);
            }
        }
        if text.is_empty() {
            return Err(self.error("structured-reference item is empty"));
        }
        Ok(StructuredReferenceItem {
            text,
            first_character_escaped,
        })
    }

    fn consume_structured_space(&mut self) -> Result<bool> {
        let start = self.offset;
        while self.peek_char().is_some_and(char::is_whitespace) {
            self.offset += self.peek_char().expect("checked").len_utf8();
        }
        if self.offset == start {
            return Ok(false);
        }
        if &self.input[start..self.offset] != " " {
            return Err(
                self.error("structured-reference whitespace cannot be represented by XLSB flags")
            );
        }
        Ok(true)
    }

    fn classify_structured_reference(
        items: Vec<StructuredReferenceItem>,
        separators: &[char],
    ) -> Result<(TableRowType, TableNamedColumns)> {
        if separators.len() + 1 != items.len() {
            return Err(Error::InvalidFormula(
                "structured-reference separator count is invalid".to_string(),
            ));
        }

        let mut rows = Vec::new();
        let mut columns = Vec::new();
        let mut item_is_column = Vec::with_capacity(items.len());
        for item in items {
            let row = if item.first_character_escaped {
                None
            } else if item.text.eq_ignore_ascii_case("#All") {
                Some(TableRowType::All)
            } else if item.text.eq_ignore_ascii_case("#Data") {
                Some(TableRowType::DataAlternate)
            } else if item.text.eq_ignore_ascii_case("#Headers") {
                Some(TableRowType::Headers)
            } else if item.text.eq_ignore_ascii_case("#Totals") {
                Some(TableRowType::Totals)
            } else if item.text.eq_ignore_ascii_case("#This Row") {
                Some(TableRowType::Current)
            } else {
                None
            };
            if let Some(row) = row {
                rows.push(row);
                item_is_column.push(false);
            } else if !item.first_character_escaped && item.text.starts_with('#') {
                return Err(Error::InvalidFormula(format!(
                    "unknown structured-reference row selector {:?}",
                    item.text
                )));
            } else if !item.first_character_escaped && item.text.starts_with('@') {
                let column = item.text[1..].to_string();
                if column.is_empty() || !rows.is_empty() {
                    return Err(Error::InvalidFormula(
                        "invalid or duplicate current-row structured reference".to_string(),
                    ));
                }
                rows.push(TableRowType::Current);
                columns.push(column);
                item_is_column.push(true);
            } else {
                columns.push(item.text);
                item_is_column.push(true);
            }
        }

        let mut colon = None;
        for (index, separator) in separators.iter().copied().enumerate() {
            match separator {
                ':' if item_is_column[index] && item_is_column[index + 1] => {
                    if colon.replace(index).is_some() {
                        return Err(Error::InvalidFormula(
                            "structured reference has more than one column range".to_string(),
                        ));
                    }
                },
                ',' if !item_is_column[index] || !item_is_column[index + 1] => {},
                ',' => {
                    return Err(Error::InvalidFormula(
                        "disjoint structured-reference columns cannot fit one PtgList".to_string(),
                    ));
                },
                _ => {
                    return Err(Error::InvalidFormula(
                        "structured-reference separator has invalid operands".to_string(),
                    ));
                },
            }
        }

        let row_type = match rows.as_slice() {
            [] => TableRowType::Data,
            [row] => *row,
            [TableRowType::Headers, TableRowType::DataAlternate] => TableRowType::DataAndHeaders,
            [TableRowType::DataAlternate, TableRowType::Totals] => TableRowType::DataAndTotals,
            _ => {
                return Err(Error::InvalidFormula(
                    "structured-reference row union cannot fit one PtgList".to_string(),
                ));
            },
        };
        let columns = match columns.as_slice() {
            [] => TableNamedColumns::All,
            [column] if colon.is_none() => TableNamedColumns::One(column.clone()),
            [first, last] if colon.is_some() => TableNamedColumns::Range {
                first: first.clone(),
                last: last.clone(),
            },
            _ => {
                return Err(Error::InvalidFormula(
                    "structured-reference columns cannot fit one PtgList".to_string(),
                ));
            },
        };
        validate_named_table_columns(&columns)?;
        Ok((row_type, columns))
    }

    fn compile_bare_resident_table_reference(
        &self,
        table_name: &str,
    ) -> Result<Option<CompileExpr>> {
        let Some(context) = self.context else {
            return Ok(None);
        };
        if !context
            .tables
            .iter()
            .any(|table| excel_name_eq(table.display_name(), table_name))
        {
            return Ok(None);
        }
        self.compile_resident_table_reference(
            table_name,
            ParsedStructuredReference {
                row_type: TableRowType::Data,
                columns: TableNamedColumns::All,
                square_bracket_space: false,
                comma_space: false,
            },
        )
        .map(Some)
    }

    fn compile_resident_table_reference(
        &self,
        table_name: &str,
        selection: ParsedStructuredReference,
    ) -> Result<CompileExpr> {
        let context = self.context.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "structured table reference {table_name:?} requires workbook compilation context"
            ))
        })?;
        let mut matches = context
            .tables
            .iter()
            .filter(|table| excel_name_eq(table.display_name(), table_name));
        let table = matches.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "structured reference names missing table {table_name:?}"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "structured reference table name {table_name:?} is ambiguous"
            )));
        }
        let current_sheet = usize::try_from(context.current_sheet)
            .map_err(|_| Error::InvalidFormula("current worksheet index overflow".to_string()))?;
        if table.sheet_index() != current_sheet {
            return Err(Error::InvalidFormula(format!(
                "table {table_name:?} is on worksheet {}, not the formula worksheet {current_sheet}",
                table.sheet_index()
            )));
        }
        let columns = match selection.columns {
            TableNamedColumns::All => TableColumns::All,
            TableNamedColumns::One(name) => {
                TableColumns::One(Self::resolve_table_column(table, &name)?)
            },
            TableNamedColumns::Range { first, last } => {
                let first = Self::resolve_table_column(table, &first)?;
                let last = Self::resolve_table_column(table, &last)?;
                if first > last {
                    return Err(Error::InvalidFormula(
                        "structured-reference column range is reversed".to_string(),
                    ));
                }
                TableColumns::Range { first, last }
            },
        };
        let sheet_index = u16::try_from(current_sheet)
            .ok()
            .and_then(|index| index.checked_add(2))
            .ok_or_else(|| {
                Error::InvalidFormula(
                    "table worksheet cannot be represented in the extern-sheet table".to_string(),
                )
            })?;
        Ok(CompileExpr::TableReference(TableReference {
            sheet_index,
            row_type: Some(selection.row_type),
            columns: Some(columns),
            square_bracket_space: selection.square_bracket_space,
            comma_space: selection.comma_space,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: Some(table.table_id()),
            external: None,
        }))
    }

    fn resolve_table_column(table: &FormulaTableDefinition, name: &str) -> Result<u16> {
        let mut matches = table
            .columns()
            .iter()
            .enumerate()
            .filter(|(_, column)| excel_name_eq(column, name));
        let (index, _) = matches.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "structured reference names missing column {name:?} in table {:?}",
                table.display_name()
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "structured-reference column {name:?} is ambiguous"
            )));
        }
        u16::try_from(index).map_err(|_| {
            Error::InvalidFormula("structured-reference column index overflow".to_string())
        })
    }

    fn compile_external_table_reference(
        &self,
        qualifier: &str,
        table: String,
        selection: ParsedStructuredReference,
    ) -> Result<CompileExpr> {
        validate_table_name(&table)?;
        let sheet_index = self.resolve_external_table_xti(qualifier)?;
        Ok(CompileExpr::TableReference(TableReference {
            sheet_index,
            row_type: None,
            columns: None,
            square_bracket_space: selection.square_bracket_space,
            comma_space: selection.comma_space,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: None,
            external: Some(ExternalTableReference {
                table,
                row_type: selection.row_type,
                columns: selection.columns,
            }),
        }))
    }

    fn resolve_external_table_xti(&self, qualifier: &str) -> Result<u16> {
        let context = self.context.ok_or_else(|| {
            Error::UnsupportedFeature(
                "external structured reference requires workbook compilation context".to_string(),
            )
        })?;
        let close = qualifier.find(']').ok_or_else(|| {
            Error::InvalidFormula("external structured reference omits ']'".to_string())
        })?;
        if !qualifier.starts_with('[') || close == 1 || close + 1 == qualifier.len() {
            return Err(Error::InvalidFormula(format!(
                "invalid external structured-reference qualifier {qualifier:?}"
            )));
        }
        let target = &qualifier[1..close];
        let sheet = &qualifier[close + 1..];
        if sheet.contains(':') {
            return Err(Error::InvalidFormula(
                "external structured reference must select exactly one worksheet".to_string(),
            ));
        }

        let mut found = None;
        for (xti_index, xti) in context.external_sheets.iter().enumerate() {
            if xti.first_sheet < 0 || xti.first_sheet != xti.last_sheet {
                continue;
            }
            let Ok(link_index) = usize::try_from(xti.external_link) else {
                continue;
            };
            let Some(FormulaSupportingLink::ExternalWorkbook(book_index)) =
                context.supporting_links.get(link_index)
            else {
                continue;
            };
            let Ok(book_index) = usize::try_from(*book_index) else {
                continue;
            };
            let Some(book) = context.external_books.get(book_index) else {
                continue;
            };
            let Ok(sheet_index) = usize::try_from(xti.first_sheet) else {
                continue;
            };
            if !book.metadata.is_workbook()
                || !excel_name_eq(book.metadata.source(), target)
                || !book
                    .metadata
                    .sheet_names()
                    .get(sheet_index)
                    .is_some_and(|candidate| excel_name_eq(candidate, sheet))
            {
                continue;
            }
            let xti_index = u16::try_from(xti_index).map_err(|_| {
                Error::InvalidFormula("external structured-reference Xti overflow".to_string())
            })?;
            if xti_index == u16::MAX || found.replace(xti_index).is_some() {
                return Err(Error::InvalidFormula(format!(
                    "external structured-reference qualifier {qualifier:?} is ambiguous"
                )));
            }
        }
        found.ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external structured-reference qualifier {qualifier:?} is missing"
            ))
        })
    }

    fn parse_string(&mut self) -> Result<String> {
        debug_assert_eq!(self.peek_char(), Some('"'));
        self.offset += 1;
        let mut value = String::new();
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated string literal"));
            };
            self.offset += ch.len_utf8();
            if ch == '"' {
                if self.peek_char() == Some('"') {
                    self.offset += 1;
                    value.push('"');
                } else {
                    break;
                }
            } else {
                value.push(ch);
            }
        }
        if value.encode_utf16().count() > 255 {
            return Err(Error::InvalidFormula(
                "formula string literal exceeds 255 UTF-16 code units".to_string(),
            ));
        }
        Ok(value)
    }

    fn parse_array_constant(&mut self) -> Result<CompileExpr> {
        let mut values = Vec::new();
        let mut rows = 1_u32;
        let mut cols = 0_u32;
        let mut current_cols = 0_u32;
        loop {
            self.skip_spaces();
            if self.peek_char() == Some('}') {
                return Err(self.error("array rows cannot be empty"));
            }
            let value = if self.peek_char() == Some('"') {
                ArrayValue::String(self.parse_string()?)
            } else if self.peek_char() == Some('#') {
                let start = self.offset;
                while self
                    .peek_char()
                    .is_some_and(|ch| !matches!(ch, ',' | ';' | '}') && !ch.is_whitespace())
                {
                    self.offset += self.peek_char().expect("checked").len_utf8();
                }
                let error = formula_error_code(&self.input[start..self.offset])
                    .ok_or_else(|| self.error("unknown array error literal"))?;
                ArrayValue::Error(error)
            } else if self.input[self.offset..]
                .get(..4)
                .is_some_and(|value| value.eq_ignore_ascii_case("TRUE"))
            {
                self.offset += 4;
                ArrayValue::Bool(true)
            } else if self.input[self.offset..]
                .get(..5)
                .is_some_and(|value| value.eq_ignore_ascii_case("FALSE"))
            {
                self.offset += 5;
                ArrayValue::Bool(false)
            } else {
                let negative = self.consume("-");
                if !negative {
                    self.consume("+");
                }
                let mut number = self.parse_number()?;
                if negative {
                    number = -number;
                }
                ArrayValue::Number(number)
            };
            values.push(value);
            current_cols = current_cols
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormula("array column count overflow".to_string()))?;

            if self.consume(",") {
                continue;
            }
            if self.consume(";") {
                if cols == 0 {
                    cols = current_cols;
                } else if cols != current_cols {
                    return Err(self.error("array rows have different column counts"));
                }
                rows = rows
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormula("array row count overflow".to_string()))?;
                current_cols = 0;
                continue;
            }
            if self.consume("}") {
                if cols == 0 {
                    cols = current_cols;
                } else if cols != current_cols {
                    return Err(self.error("array rows have different column counts"));
                }
                break;
            }
            return Err(self.error("expected ',', ';', or '}' in array constant"));
        }
        if rows > 1_048_576 || cols == 0 || cols > 16_384 {
            return Err(self.error("array dimensions exceed worksheet limits"));
        }
        Ok(CompileExpr::Array { rows, cols, values })
    }

    fn parse_number(&mut self) -> Result<f64> {
        self.skip_spaces();
        let start = self.offset;
        let mut seen_exponent = false;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() || ch == '.' {
                self.offset += 1;
            } else if matches!(ch, 'e' | 'E') && !seen_exponent {
                seen_exponent = true;
                self.offset += 1;
                if matches!(self.peek_char(), Some('+' | '-')) {
                    self.offset += 1;
                }
            } else {
                break;
            }
        }
        self.input[start..self.offset]
            .parse::<f64>()
            .map_err(|_| self.error("invalid numeric literal"))
    }

    fn parse_error_literal(&mut self) -> Result<u8> {
        self.skip_spaces();
        let rest = &self.input[self.offset..];
        let Some((literal, code)) = FORMULA_ERRORS.iter().find(|(literal, _)| {
            rest.get(..literal.len())
                .is_some_and(|value| value.eq_ignore_ascii_case(literal))
        }) else {
            return Err(self.error("unknown formula error literal"));
        };
        self.offset += literal.len();
        Ok(*code)
    }

    fn parse_identifier(&mut self) -> Result<String> {
        self.skip_spaces();
        let start = self.offset;
        while let Some(ch) = self.peek_char() {
            if ch.is_alphanumeric() || matches!(ch, '_' | '.' | '$' | '?' | '\\' | '\u{061F}') {
                self.offset += ch.len_utf8();
            } else {
                break;
            }
        }
        if self.offset == start {
            Err(self.error("expected literal, reference, or function"))
        } else {
            Ok(self.input[start..self.offset].to_string())
        }
    }

    fn parse_quoted_sheet_name(&mut self) -> Result<String> {
        self.skip_spaces();
        debug_assert_eq!(self.peek_char(), Some('\''));
        self.offset += 1;
        let mut name = String::new();
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated quoted worksheet name"));
            };
            self.offset += ch.len_utf8();
            if ch == '\'' {
                if self.peek_char() == Some('\'') {
                    self.offset += 1;
                    name.push('\'');
                } else {
                    break;
                }
            } else {
                name.push(ch);
            }
        }
        if name.is_empty() {
            return Err(self.error("worksheet name is empty"));
        }
        Ok(name)
    }

    fn split_sheet_qualifier(value: &str) -> Result<(&str, Option<&str>)> {
        let Some((first, last)) = value.split_once(':') else {
            return Ok((value, None));
        };
        if first.is_empty() || last.is_empty() || last.contains(':') {
            return Err(Error::InvalidFormula(format!(
                "invalid worksheet range {value:?}"
            )));
        }
        Ok((first, Some(last)))
    }

    fn parse_qualified_reference(
        &mut self,
        first_sheet: &str,
        last_sheet: Option<&str>,
    ) -> Result<CompileExpr> {
        let sheet_index = self.resolve_sheet_range(first_sheet, last_sheet)?;
        let first_text = self.parse_identifier()?;
        let first = parse_a1_reference(&first_text)
            .ok_or_else(|| self.error("invalid sheet-qualified cell reference"))?;
        if self.consume(":") {
            let second_text = self.parse_identifier()?;
            let second = parse_a1_reference(&second_text)
                .ok_or_else(|| self.error("invalid sheet-qualified range end"))?;
            Ok(CompileExpr::Area3d(sheet_index, first, second))
        } else {
            Ok(CompileExpr::Ref3d(sheet_index, first))
        }
    }

    fn resolve_sheet_range(&self, first_sheet: &str, last_sheet: Option<&str>) -> Result<u16> {
        let context = self.context.ok_or_else(|| {
            Error::UnsupportedFeature(
                "sheet-qualified reference requires workbook compilation context".to_string(),
            )
        })?;
        let first_index = context
            .worksheet_names
            .iter()
            .position(|candidate| excel_name_eq(candidate, first_sheet))
            .ok_or_else(|| Error::WorksheetNotFound(first_sheet.to_string()))?;
        let last_index = if let Some(last_sheet) = last_sheet {
            context
                .worksheet_names
                .iter()
                .position(|candidate| excel_name_eq(candidate, last_sheet))
                .ok_or_else(|| Error::WorksheetNotFound(last_sheet.to_string()))?
        } else {
            first_index
        };
        if last_index < first_index {
            return Err(Error::InvalidFormula(format!(
                "worksheet range {first_sheet:?}:{last_sheet:?} is in reverse workbook order"
            )));
        }
        if first_index == last_index {
            return u16::try_from(first_index)
                .ok()
                .and_then(|index| index.checked_add(2))
                .ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "worksheet {first_sheet:?} cannot be represented in the extern-sheet table"
                    ))
                });
        }

        let first = u32::try_from(first_index)
            .map_err(|_| Error::InvalidFormula("first sheet index overflow".to_string()))?;
        let last = u32::try_from(last_index)
            .map_err(|_| Error::InvalidFormula("last sheet index overflow".to_string()))?;
        let mut ranges = context.sheet_ranges.borrow_mut();
        let range_index = if let Some(index) = ranges
            .iter()
            .position(|candidate| *candidate == (first, last))
        {
            index
        } else {
            let base_count = context
                .worksheet_names
                .len()
                .checked_add(2)
                .ok_or_else(|| Error::InvalidFormula("Xti count overflow".to_string()))?;
            if base_count
                .checked_add(ranges.len())
                .is_none_or(|count| count >= usize::from(u16::MAX))
            {
                return Err(Error::InvalidFormula(
                    "formula sheet ranges exceed the XLSB extern-sheet limit".to_string(),
                ));
            }
            ranges.push((first, last));
            ranges.len() - 1
        };
        let xti_index = context
            .worksheet_names
            .len()
            .checked_add(2)
            .and_then(|base| base.checked_add(range_index))
            .ok_or_else(|| Error::InvalidFormula("Xti index overflow".to_string()))?;
        u16::try_from(xti_index)
            .map_err(|_| Error::InvalidFormula("Xti index overflow".to_string()))
    }

    fn resolve_defined_name(&self, name: &str) -> Result<u32> {
        let context = self.context.ok_or_else(|| {
            Error::UnsupportedFeature(format!(
                "defined name {name:?} requires workbook compilation context"
            ))
        })?;
        let local = context.defined_names.iter().position(|candidate| {
            candidate.sheet_id == Some(context.current_sheet)
                && excel_name_eq(&candidate.name, name)
        });
        let index = local.or_else(|| {
            context.defined_names.iter().position(|candidate| {
                candidate.sheet_id.is_none() && excel_name_eq(&candidate.name, name)
            })
        });
        let index = index.ok_or_else(|| {
            Error::InvalidFormula(format!(
                "defined name {name:?} is not visible from worksheet {}",
                context.current_sheet
            ))
        })?;
        u32::try_from(index)
            .ok()
            .and_then(|index| index.checked_add(1))
            .ok_or_else(|| Error::InvalidFormula("defined-name index overflow".to_string()))
    }

    fn consume(&mut self, text: &str) -> bool {
        self.skip_spaces();
        if self.input[self.offset..].starts_with(text) {
            self.offset += text.len();
            true
        } else {
            false
        }
    }

    fn skip_spaces(&mut self) {
        while self.peek_char().is_some_and(char::is_whitespace) {
            self.offset += self.peek_char().expect("checked").len_utf8();
        }
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.offset..].chars().next()
    }

    fn error(&self, message: &str) -> Error {
        Error::InvalidFormula(format!("{message} at byte {}", self.offset))
    }

    fn emit(
        expression: &CompileExpr,
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        match expression {
            CompileExpr::Number(value) => {
                validate_xnum(*value, "compiled number")?;
                if value.fract() == 0.0 && *value >= 0.0 && *value <= f64::from(u16::MAX) {
                    output.push(ptg_types::PTG_INT);
                    output.extend_from_slice(&(*value as u16).to_le_bytes());
                } else {
                    output.push(ptg_types::PTG_NUM);
                    output.extend_from_slice(&value.to_le_bytes());
                }
            },
            CompileExpr::String(value) => {
                let utf16: Vec<u16> = value.encode_utf16().collect();
                output.push(ptg_types::PTG_STR);
                output.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                for unit in utf16 {
                    output.extend_from_slice(&unit.to_le_bytes());
                }
            },
            CompileExpr::Bool(value) => {
                output.push(ptg_types::PTG_BOOL);
                output.push(u8::from(*value));
            },
            CompileExpr::Error(error) => {
                output.push(ptg_types::PTG_ERR);
                output.push(*error);
            },
            CompileExpr::MissingArg => output.push(ptg_types::PTG_MISSING_ARG),
            CompileExpr::Parenthesized(expression) => {
                Self::emit(expression, output, extra, encoding)?;
                output.push(ptg_types::PTG_PAREN);
            },
            CompileExpr::Array { rows, cols, values } => {
                if matches!(encoding, FormulaEncoding::Shared { .. }) {
                    return Err(Error::InvalidFormula(
                        "shared formulas cannot contain PtgArray".to_string(),
                    ));
                }
                output.push(0x40); // PtgArray, VALUE class
                output.extend_from_slice(&[0; 14]);
                extra.extend_from_slice(&rows.to_le_bytes());
                extra.extend_from_slice(&cols.to_le_bytes());
                for value in values {
                    match value {
                        ArrayValue::Number(value) => {
                            extra.push(0x00);
                            extra.extend_from_slice(&value.to_le_bytes());
                        },
                        ArrayValue::String(value) => {
                            let utf16: Vec<u16> = value.encode_utf16().collect();
                            extra.push(0x01);
                            extra.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                            for unit in utf16 {
                                extra.extend_from_slice(&unit.to_le_bytes());
                            }
                        },
                        ArrayValue::Bool(value) => {
                            extra.extend_from_slice(&[0x02, u8::from(*value)]);
                        },
                        ArrayValue::Error(error) => {
                            extra.extend_from_slice(&[0x04, *error, 0, 0, 0]);
                        },
                    }
                }
            },
            CompileExpr::Ref(reference) => match encoding {
                FormulaEncoding::Cell => emit_reference(output, 0x44, *reference),
                FormulaEncoding::Shared { base_row, base_col } => {
                    emit_shared_reference(output, 0x4C, *reference, base_row, base_col)?
                },
            },
            CompileExpr::Area(first, last) => {
                match encoding {
                    FormulaEncoding::Cell => {
                        output.push(0x25); // PtgArea, REFERENCE class
                        output.extend_from_slice(&first.row.to_le_bytes());
                        output.extend_from_slice(&last.row.to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
                    },
                    FormulaEncoding::Shared { base_row, base_col } => {
                        output.push(0x2D); // PtgAreaN, REFERENCE class
                        let (first_row, first_col) =
                            encode_shared_reference(*first, base_row, base_col)?;
                        let (last_row, last_col) =
                            encode_shared_reference(*last, base_row, base_col)?;
                        output.extend_from_slice(&first_row.to_le_bytes());
                        output.extend_from_slice(&last_row.to_le_bytes());
                        output.extend_from_slice(&first_col.to_le_bytes());
                        output.extend_from_slice(&last_col.to_le_bytes());
                    },
                }
            },
            CompileExpr::Ref3d(sheet_index, reference) => {
                output.push(0x5A); // PtgRef3d, VALUE class
                output.extend_from_slice(&sheet_index.to_le_bytes());
                output.extend_from_slice(&reference.row.to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*reference).to_le_bytes());
            },
            CompileExpr::Area3d(sheet_index, first, last) => {
                output.push(0x5B); // PtgArea3d, VALUE class
                output.extend_from_slice(&sheet_index.to_le_bytes());
                output.extend_from_slice(&first.row.to_le_bytes());
                output.extend_from_slice(&last.row.to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
            },
            CompileExpr::Name(index) => {
                output.push(0x43); // PtgName, VALUE class
                output.extend_from_slice(&index.to_le_bytes());
            },
            CompileExpr::TableReference(reference) => {
                let (token, payload) = reference.to_extended_binary()?;
                output.extend_from_slice(&token);
                extra.extend_from_slice(&payload);
            },
            CompileExpr::Unary(operator, operand) => {
                Self::emit(operand, output, extra, encoding)?;
                output.push(match operator {
                    UnaryOperator::Plus => ptg_types::PTG_UPLUS,
                    UnaryOperator::Minus => ptg_types::PTG_UMINUS,
                    UnaryOperator::Percent => ptg_types::PTG_PERCENT,
                });
            },
            CompileExpr::Binary(operator, left, right) => {
                Self::emit(left, output, extra, encoding)?;
                Self::emit(right, output, extra, encoding)?;
                output.push(match operator {
                    BinaryOperator::Add => ptg_types::PTG_ADD,
                    BinaryOperator::Subtract => ptg_types::PTG_SUB,
                    BinaryOperator::Multiply => ptg_types::PTG_MUL,
                    BinaryOperator::Divide => ptg_types::PTG_DIV,
                    BinaryOperator::Power => ptg_types::PTG_POWER,
                    BinaryOperator::Concat => ptg_types::PTG_CONCAT,
                    BinaryOperator::LessThan => ptg_types::PTG_LT,
                    BinaryOperator::LessEqual => ptg_types::PTG_LE,
                    BinaryOperator::Equal => ptg_types::PTG_EQ,
                    BinaryOperator::GreaterEqual => ptg_types::PTG_GE,
                    BinaryOperator::GreaterThan => ptg_types::PTG_GT,
                    BinaryOperator::NotEqual => ptg_types::PTG_NE,
                    BinaryOperator::Intersection => ptg_types::PTG_ISECT,
                    BinaryOperator::Union => ptg_types::PTG_UNION,
                    BinaryOperator::Range => ptg_types::PTG_RANGE,
                });
            },
            CompileExpr::Function(function, arguments) => {
                if function.index == 1 {
                    return Self::emit_if(arguments, output, extra, encoding);
                }
                if function.index == 100 {
                    return Self::emit_choose(arguments, output, extra, encoding);
                }
                if function.index == 480 {
                    return Self::emit_iferror(arguments, output, extra, encoding);
                }
                for argument in arguments {
                    Self::emit(argument, output, extra, encoding)?;
                }
                if function.min_args == function.max_args {
                    output.push(0x41); // PtgFunc, VALUE class
                    output.extend_from_slice(&function.index.to_le_bytes());
                } else {
                    output.push(0x42); // PtgFuncVar, VALUE class
                    output.push(arguments.len() as u8);
                    output.extend_from_slice(&function.index.to_le_bytes());
                }
            },
        }
        Ok(())
    }

    fn emit_if(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert!(matches!(arguments.len(), 2 | 3));
        Self::emit(&arguments[0], output, extra, encoding)?;
        let attr_if = append_attribute(output, 0x02, 0);
        Self::emit(&arguments[1], output, extra, encoding)?;
        let goto_true = append_attribute(output, 0x08, 0);
        let goto_false = if arguments.len() == 3 {
            Self::emit(&arguments[2], output, extra, encoding)?;
            Some(append_attribute(output, 0x08, 0))
        } else {
            None
        };
        output.extend_from_slice(&[0x42, arguments.len() as u8, 0x01, 0x00]);

        patch_attribute_offset(output, attr_if, goto_true + 4 - (attr_if + 4))?;
        patch_skip_to_end(output, goto_true)?;
        if let Some(position) = goto_false {
            patch_skip_to_end(output, position)?;
        }
        Ok(())
    }

    fn emit_iferror(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert_eq!(arguments.len(), 2);
        Self::emit(&arguments[0], output, extra, encoding)?;
        let attr_if_error = append_attribute(output, 0x80, 0);
        Self::emit(&arguments[1], output, extra, encoding)?;
        let goto = append_attribute(output, 0x08, 0);
        output.extend_from_slice(&[0x41, 0xE0, 0x01]);

        patch_attribute_offset(output, attr_if_error, goto + 4 - (attr_if_error + 4))?;
        patch_skip_to_end(output, goto)?;
        Ok(())
    }

    fn emit_choose(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert!((2..=255).contains(&arguments.len()));
        Self::emit(&arguments[0], output, extra, encoding)?;
        let choice_count = arguments.len() - 1;
        let attr_choose = output.len();
        output.extend_from_slice(&[ptg_types::PTG_ATTR, 0x04]);
        output.extend_from_slice(&(choice_count as u16).to_le_bytes());
        output.resize(output.len() + (choice_count + 1) * 2, 0);
        let attr_size = output.len() - attr_choose;
        patch_u16(
            output,
            attr_choose + 4,
            attr_size - 4,
            "PtgAttrChoose first offset",
        )?;

        let mut gotos = Vec::with_capacity(choice_count);
        for (index, argument) in arguments[1..].iter().enumerate() {
            Self::emit(argument, output, extra, encoding)?;
            gotos.push(append_attribute(output, 0x08, 0));
            let cumulative = output.len() - (attr_choose + attr_size);
            patch_u16(
                output,
                attr_choose + 6 + index * 2,
                cumulative,
                "PtgAttrChoose branch offset",
            )?;
        }
        output.extend_from_slice(&[0x42, arguments.len() as u8, 0x64, 0x00]);
        for goto in gotos {
            patch_skip_to_end(output, goto)?;
        }
        Ok(())
    }
}

fn append_attribute(output: &mut Vec<u8>, selector: u8, offset: u16) -> usize {
    let position = output.len();
    output.extend_from_slice(&[ptg_types::PTG_ATTR, selector]);
    output.extend_from_slice(&offset.to_le_bytes());
    position
}

fn patch_attribute_offset(output: &mut [u8], position: usize, offset: usize) -> Result<()> {
    patch_u16(output, position + 2, offset, "PtgAttr offset")
}

fn patch_skip_to_end(output: &mut [u8], position: usize) -> Result<()> {
    let remaining = output
        .len()
        .checked_sub(position + 4)
        .ok_or_else(|| Error::InvalidFormula("PtgAttrGoTo position exceeds formula".to_string()))?;
    let offset = remaining
        .checked_sub(1)
        .ok_or_else(|| Error::InvalidFormula("PtgAttrGoTo has no following token".to_string()))?;
    patch_attribute_offset(output, position, offset)
}

fn patch_u16(output: &mut [u8], position: usize, value: usize, context: &str) -> Result<()> {
    let value = u16::try_from(value)
        .map_err(|_| Error::InvalidFormula(format!("{context} exceeds 65,535 bytes")))?;
    let target = output
        .get_mut(position..position + 2)
        .ok_or_else(|| Error::InvalidFormula(format!("{context} position is outside formula")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn validate_xnum(value: f64, context: &str) -> Result<()> {
    if !value.is_finite()
        || (value == 0.0 && value.is_sign_negative())
        || (value != 0.0 && !value.is_normal())
    {
        return Err(Error::InvalidFormula(format!(
            "{context} contains a non-finite, denormalized, or negative-zero Xnum"
        )));
    }
    Ok(())
}

const FORMULA_ERRORS: &[(&str, u8)] = &[
    ("#GETTING_DATA", 0x2B),
    ("#DIV/0!", 0x07),
    ("#VALUE!", 0x0F),
    ("#NULL!", 0x00),
    ("#NAME?", 0x1D),
    ("#REF!", 0x17),
    ("#NUM!", 0x24),
    ("#N/A", 0x2A),
];

fn formula_error_code(value: &str) -> Option<u8> {
    FORMULA_ERRORS
        .iter()
        .find_map(|(literal, code)| literal.eq_ignore_ascii_case(value).then_some(*code))
}

pub(crate) fn excel_name_eq(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

fn parse_a1_reference(value: &str) -> Option<A1Reference> {
    let bytes = value.as_bytes();
    let mut offset = 0;
    let col_relative = bytes.get(offset) != Some(&b'$');
    if !col_relative {
        offset += 1;
    }
    let col_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_alphabetic) {
        offset += 1;
    }
    if offset == col_start {
        return None;
    }
    let mut col = 0u32;
    for byte in bytes[col_start..offset].iter().map(u8::to_ascii_uppercase) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte - b'A' + 1))?;
    }
    if col == 0 || col > 16_384 {
        return None;
    }

    let row_relative = bytes.get(offset) != Some(&b'$');
    if !row_relative {
        offset += 1;
    }
    let row_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_digit) {
        offset += 1;
    }
    if offset == row_start || offset != bytes.len() {
        return None;
    }
    let row = value[row_start..offset].parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 {
        return None;
    }
    Some(A1Reference {
        row: row - 1,
        col: col - 1,
        row_relative,
        col_relative,
    })
}

fn reference_column_bits(reference: A1Reference) -> u16 {
    let mut bits = reference.col as u16;
    if reference.col_relative {
        bits |= 0x4000;
    }
    if reference.row_relative {
        bits |= 0x8000;
    }
    bits
}

fn emit_reference(output: &mut Vec<u8>, token: u8, reference: A1Reference) {
    output.push(token);
    output.extend_from_slice(&reference.row.to_le_bytes());
    output.extend_from_slice(&reference_column_bits(reference).to_le_bytes());
}

fn emit_shared_reference(
    output: &mut Vec<u8>,
    token: u8,
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> Result<()> {
    let (row, col) = encode_shared_reference(reference, base_row, base_col)?;
    output.push(token);
    output.extend_from_slice(&row.to_le_bytes());
    output.extend_from_slice(&col.to_le_bytes());
    Ok(())
}

fn encode_shared_reference(
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> Result<(u32, u16)> {
    let row = if reference.row_relative {
        let offset = i64::from(reference.row) - i64::from(base_row);
        i32::try_from(offset)
            .map_err(|_| Error::InvalidFormula("shared row offset overflow".to_string()))?
            as u32
    } else {
        reference.row
    };
    let col_value = if reference.col_relative {
        let offset = i64::from(reference.col) - i64::from(base_col);
        if !(-16_383..=16_383).contains(&offset) {
            return Err(Error::InvalidFormula(format!(
                "shared column offset {offset} is outside the XLSB range"
            )));
        }
        (offset as i32 as u16) & 0x3FFF
    } else {
        reference.col as u16
    };
    let mut col = col_value;
    if reference.col_relative {
        col |= 0x4000;
    }
    if reference.row_relative {
        col |= 0x8000;
    }
    Ok((row, col))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_operators() {
        let data = vec![0x03]; // PTG_ADD
        let mut parser = FormulaParser::new(&data);
        let tokens = parser.parse().unwrap();
        assert_eq!(tokens.len(), 1);
        match &tokens[0] {
            Token::BinaryOp(BinaryOperator::Add) => {},
            _ => panic!("Expected Add operator"),
        }
    }

    #[test]
    fn test_parse_number() {
        let mut data = vec![0x1F]; // PTG_NUM
        data.extend_from_slice(&42.5f64.to_le_bytes());
        let mut parser = FormulaParser::new(&data);
        let tokens = parser.parse().unwrap();
        assert_eq!(tokens.len(), 1);
        match &tokens[0] {
            Token::Number(n) if (*n - 42.5).abs() < 0.001 => {},
            _ => panic!("Expected number 42.5"),
        }
    }

    #[test]
    fn test_formula_converter() {
        let tokens = vec![
            Token::Number(1.0),
            Token::Number(2.0),
            Token::BinaryOp(BinaryOperator::Add),
        ];
        let formula = FormulaConverter::tokens_to_string(&tokens);
        assert_eq!(formula, "(1+2)");
    }

    #[test]
    fn parses_ms_xlsb_brt_fmla_num_example_formula() {
        // [MS-XLSB] 3.7.37: PtgRef(C13), PtgInt(2), PtgMul.
        let rgce = vec![
            0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
        ];
        let parsed = CellParsedFormula {
            rgce: rgce.clone(),
            rgcb: Vec::new(),
        };
        let bytes = parsed.to_bytes().unwrap();
        let (roundtrip, consumed) = CellParsedFormula::parse(&bytes).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(roundtrip, parsed);

        let tokens = FormulaParser::new(&rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(C13*2)"
        );
    }

    #[test]
    fn compiler_matches_ms_xlsb_reference_and_multiply_tokens() {
        let formula = FormulaCompiler::compile("=C13*2").unwrap();
        assert_eq!(
            formula.rgce,
            vec![
                0x44, 0x0C, 0x00, 0x00, 0x00, 0x02, 0xC0, 0x1E, 0x02, 0x00, 0x05,
            ]
        );
    }

    #[test]
    fn compiler_emits_conditional_control_flow_attributes() {
        assert_eq!(
            FormulaCompiler::compile("IF(TRUE,1,2)").unwrap().rgce,
            vec![
                0x1D, 0x01, 0x19, 0x02, 0x07, 0x00, 0x1E, 0x01, 0x00, 0x19, 0x08, 0x0A, 0x00, 0x1E,
                0x02, 0x00, 0x19, 0x08, 0x03, 0x00, 0x42, 0x03, 0x01, 0x00,
            ]
        );
        assert_eq!(
            FormulaCompiler::compile("IFERROR(1,2)").unwrap().rgce,
            vec![
                0x1E, 0x01, 0x00, 0x19, 0x80, 0x07, 0x00, 0x1E, 0x02, 0x00, 0x19, 0x08, 0x02, 0x00,
                0x41, 0xE0, 0x01,
            ]
        );
        assert_eq!(
            FormulaCompiler::compile("CHOOSE(2,10,20)").unwrap().rgce,
            vec![
                0x1E, 0x02, 0x00, 0x19, 0x04, 0x02, 0x00, 0x06, 0x00, 0x07, 0x00, 0x0E, 0x00, 0x1E,
                0x0A, 0x00, 0x19, 0x08, 0x0A, 0x00, 0x1E, 0x14, 0x00, 0x19, 0x08, 0x03, 0x00, 0x42,
                0x03, 0x64, 0x00,
            ]
        );

        for source in ["IF(TRUE,1,2)", "IFERROR(1,2)", "CHOOSE(2,10,20)"] {
            let compiled = FormulaCompiler::compile(source).unwrap();
            let tokens = FormulaParser::new(&compiled.rgce).parse().unwrap();
            assert_eq!(
                FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
                source
            );
        }

        let mut malformed_if = FormulaCompiler::compile("IF(TRUE,1,2)").unwrap().rgce;
        malformed_if[4] = 6;
        assert!(matches!(
            FormulaParser::new(&malformed_if).parse(),
            Err(Error::InvalidFormula(_))
        ));
        let mut malformed_choose = FormulaCompiler::compile("CHOOSE(2,10,20)").unwrap().rgce;
        malformed_choose[7] = 5;
        assert!(matches!(
            FormulaParser::new(&malformed_choose).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn compiler_supports_ranges_functions_unicode_and_absolute_refs() {
        let formula = FormulaCompiler::compile("SUM($A$1:B3)+\"荔枝\"").unwrap();
        let tokens = FormulaParser::new(&formula.rgce).parse().unwrap();
        let text = FormulaConverter::try_tokens_to_string(&tokens).unwrap();
        assert_eq!(text, "(SUM($A$1:B3)+\"荔枝\")");
    }

    #[test]
    fn compiler_emits_contextual_names_and_sheet_references() {
        let worksheet_names = vec![
            "Data".to_string(),
            "O'Brien Data".to_string(),
            "Summary".to_string(),
        ];
        let defined_names = vec![
            FormulaDefinedName {
                name: "Rate".to_string(),
                sheet_id: None,
            },
            FormulaDefinedName {
                name: "Rate".to_string(),
                sheet_id: Some(1),
            },
        ];
        let sheet_ranges = std::cell::RefCell::new(Vec::new());
        let context = FormulaCompilationContext {
            worksheet_names: &worksheet_names,
            defined_names: &defined_names,
            tables: &[],
            supporting_links: &[],
            external_sheets: &[],
            external_books: &[],
            sheet_ranges: &sheet_ranges,
            current_sheet: 1,
        };
        let compiled =
            FormulaCompiler::compile_with_context("Rate+Data!A1:B2+'O''Brien Data'!$C$3", &context)
                .unwrap();

        assert!(compiled.rgce.starts_with(&[0x43, 2, 0, 0, 0]));
        assert!(
            compiled
                .rgce
                .windows(3)
                .any(|window| window == [0x5B, 2, 0])
        );
        assert!(
            compiled
                .rgce
                .windows(3)
                .any(|window| window == [0x5A, 3, 0])
        );
        assert!(FormulaCompiler::compile("Rate").is_err());
        assert!(FormulaCompiler::compile("Data!A1").is_err());

        let span = FormulaCompiler::compile_with_context(
            "SUM('Data:Summary'!A1)+Data:Summary!$B$2",
            &context,
        )
        .unwrap();
        assert_eq!(&*sheet_ranges.borrow(), &[(0, 2)]);
        assert_eq!(
            span.rgce
                .windows(3)
                .filter(|window| *window == [0x5A, 5, 0])
                .count(),
            2
        );
        assert!(FormulaCompiler::compile_with_context("Summary:Data!A1", &context).is_err());
    }

    #[test]
    fn builtin_function_table_is_sorted_unique_and_non_macro() {
        assert_eq!(BUILTIN_FUNCTIONS.len(), 363);
        assert!(
            BUILTIN_FUNCTIONS
                .windows(2)
                .all(|entries| entries[0].0 < entries[1].0)
        );

        let mut names = std::collections::HashSet::new();
        for &(index, name, min_args, max_args) in BUILTIN_FUNCTIONS {
            assert!(min_args <= max_args, "invalid arity for {name}");
            assert!(
                names.insert(name.to_ascii_uppercase()),
                "duplicate function name {name} at index {index}"
            );
        }

        assert!(builtin_function_by_index(53).is_none()); // GOTO macro function
        assert!(builtin_function_by_index(110).is_none()); // EXEC macro function
        assert!(builtin_function_by_index(255).is_none()); // context-dependent UDF
        assert!(builtin_function_by_index(468).is_none()); // future-function CONVERT
    }

    #[test]
    fn compiler_covers_legacy_analysis_and_ooxml_function_ranges() {
        let cases: &[(&str, &str, &[u8])] = &[
            ("ROUNDUP(1.2,0)", "ROUNDUP(1.2,0)", &[0x41, 0xD4, 0x00]),
            ("MEDIAN(1,2,3)", "MEDIAN(1,2,3)", &[0x42, 0x03, 0xE3, 0x00]),
            ("CUBESETCOUNT(1)", "CUBESETCOUNT(1)", &[0x41, 0xDF, 0x01]),
        ];
        for &(source, expected, token_suffix) in cases {
            let formula = FormulaCompiler::compile(source).unwrap();
            assert!(formula.rgce.ends_with(token_suffix));
            let tokens = FormulaParser::new(&formula.rgce).parse().unwrap();
            assert_eq!(
                FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
                expected
            );
        }

        let accrint = FormulaCompiler::compile("ACCRINT(1,2,3,4,5,6,7,8)").unwrap();
        assert!(accrint.rgce.ends_with(&[0x42, 0x08, 0xD5, 0x01]));
        assert!(FormulaCompiler::compile("ACCRINT(1,2,3,4,5,6,7,8,9)").is_err());
    }

    #[test]
    fn compiler_and_parser_enforce_function_argument_grammars() {
        assert!(FormulaCompiler::compile("SUM()").is_err());
        assert!(FormulaCompiler::compile("COUNTIFS(A1,1)").is_ok());
        assert!(FormulaCompiler::compile("COUNTIFS(A1,1,B1)").is_err());
        assert!(FormulaCompiler::compile("SUMIFS(A1,B1,1)").is_ok());
        assert!(FormulaCompiler::compile("SUMIFS(A1,B1,1,C1)").is_err());
        assert!(FormulaCompiler::compile("GETPIVOTDATA(1,2,3)").is_ok());
        assert!(FormulaCompiler::compile("GETPIVOTDATA(1,2,3,4,5)").is_err());

        assert!(matches!(
            FormulaParser::new(&[0x41, 0xE3, 0x00]).parse(),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0x42, 0x02, 0xD4, 0x00]).parse(),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0x42, 0x03, 0xE1, 0x01]).parse(),
            Err(Error::InvalidFormula(_))
        ));

        assert!(matches!(
            FormulaCompiler::compile("EXEC(\"calc\")"),
            Err(Error::UnsupportedFeature(_))
        ));
        assert!(matches!(
            FormulaCompiler::compile("CONVERT(1,\"m\",\"ft\")"),
            Err(Error::UnsupportedFeature(_))
        ));
    }

    #[test]
    fn variable_functions_support_the_full_u8_argument_count() {
        let formula_255 = format!("SUM({})", vec!["1"; 255].join(","));
        let compiled = FormulaCompiler::compile(&formula_255).unwrap();
        assert!(compiled.rgce.ends_with(&[0x42, 0xFF, 0x04, 0x00]));
        let tokens = FormulaParser::new(&compiled.rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            formula_255
        );

        let formula_256 = format!("SUM({})", vec!["1"; 256].join(","));
        assert!(matches!(
            FormulaCompiler::compile(&formula_256),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn compiler_and_converter_preserve_missing_arguments_and_parentheses() {
        let missing = FormulaCompiler::compile("IF(TRUE,,0)").unwrap();
        assert!(missing.rgce.contains(&ptg_types::PTG_MISSING_ARG));
        let tokens = FormulaParser::new(&missing.rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "IF(TRUE,,0)"
        );

        let parenthesized = FormulaCompiler::compile("(1+2)*3").unwrap();
        assert!(parenthesized.rgce.contains(&ptg_types::PTG_PAREN));
        let tokens = FormulaParser::new(&parenthesized.rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(((1+2))*3)"
        );
    }

    #[test]
    fn parser_converts_binary_reference_operators() {
        let mut rgce = FormulaCompiler::compile("A1").unwrap().rgce;
        rgce.extend_from_slice(&FormulaCompiler::compile("B2").unwrap().rgce);
        rgce.push(ptg_types::PTG_UNION);
        let tokens = FormulaParser::new(&rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(A1,B2)"
        );
    }

    #[test]
    fn parser_decodes_all_reference_error_token_forms() {
        let cases = [
            (
                vec![0x4A, 1, 2, 3, 4, 5, 6],
                Token::ReferenceError {
                    is_area: false,
                    sheet_index: None,
                },
            ),
            (
                vec![0x4B, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                Token::ReferenceError {
                    is_area: true,
                    sheet_index: None,
                },
            ),
            (
                vec![0x5C, 0x34, 0x12, 1, 2, 3, 4, 5, 6],
                Token::ReferenceError {
                    is_area: false,
                    sheet_index: Some(0x1234),
                },
            ),
            (
                vec![0x7D, 0x78, 0x56, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                Token::ReferenceError {
                    is_area: true,
                    sheet_index: Some(0x5678),
                },
            ),
        ];

        for (bytes, expected) in cases {
            let tokens = FormulaParser::new(&bytes).parse().unwrap();
            assert_eq!(tokens, vec![expected]);
            assert_eq!(
                FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
                "#REF!"
            );
        }

        assert!(matches!(
            FormulaParser::new(&[0x4B; 12]).parse(),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0xAA, 0, 0, 0, 0, 0, 0]).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_resolves_internal_3d_references_and_defined_names() {
        let context = FormulaResolutionContext {
            worksheet_names: vec!["Data 1".to_string(), "Last".to_string()].into(),
            supporting_links: vec![FormulaSupportingLink::SelfWorkbook].into(),
            external_sheets: vec![
                FormulaExternalSheet {
                    external_link: 0,
                    first_sheet: 0,
                    last_sheet: 0,
                },
                FormulaExternalSheet {
                    external_link: 0,
                    first_sheet: 0,
                    last_sheet: 1,
                },
            ]
            .into(),
            external_books: Vec::new().into(),
            defined_names: vec!["Rate".to_string()].into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };

        let ref_3d = [0x5A, 0, 0, 1, 0, 0, 0, 0, 0xC0];
        let tokens = FormulaParser::new(&ref_3d).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context).unwrap(),
            "'Data 1'!A2"
        );
        assert!(FormulaConverter::try_tokens_to_string(&tokens).is_err());

        let area_3d = [0x7B, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0];
        let tokens = FormulaParser::new(&area_3d).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context).unwrap(),
            "'Data 1:Last'!$A$1:$B$2"
        );

        let name = [0x43, 1, 0, 0, 0];
        let tokens = FormulaParser::new(&name).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context).unwrap(),
            "Rate"
        );
    }

    #[test]
    fn contextual_reference_parser_rejects_invalid_indices_and_payloads() {
        let context = FormulaResolutionContext {
            worksheet_names: vec!["Sheet1".to_string()].into(),
            supporting_links: vec![FormulaSupportingLink::SelfWorkbook].into(),
            external_sheets: Vec::new().into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        let invalid_xti = [0x5A, 0, 0, 0, 0, 0, 0, 0, 0];
        let tokens = FormulaParser::new(&invalid_xti).parse().unwrap();
        assert!(matches!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0x43, 0, 0, 0, 0]).parse(),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0xDA, 0, 0, 0, 0, 0, 0, 0, 0]).parse(),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0x5B; 14]).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_rejects_reserved_class_bits_and_invalid_absolute_ranges() {
        for bytes in [
            vec![0xA4, 0, 0, 0, 0, 0, 0],
            vec![0xA1, 0, 0],
            vec![0xC3, 1, 0, 0, 0],
            vec![0xB9, 0, 0, 1, 0, 0, 0],
        ] {
            assert!(matches!(
                FormulaParser::new(&bytes).parse(),
                Err(Error::InvalidFormula(_))
            ));
        }

        let row_past_end = [0x44, 0x00, 0x00, 0x10, 0x00, 0, 0];
        assert!(matches!(
            FormulaParser::new(&row_past_end).parse(),
            Err(Error::InvalidFormula(_))
        ));

        let reversed_area = [0x45, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
        assert!(matches!(
            FormulaParser::new(&reversed_area).parse(),
            Err(Error::InvalidCellReference(_))
        ));
    }

    #[test]
    fn resolves_same_sheet_supporting_links_in_the_consuming_sheet() {
        let context = FormulaResolutionContext {
            worksheet_names: vec!["First".to_string(), "Current Sheet".to_string()].into(),
            supporting_links: vec![FormulaSupportingLink::SameSheet].into(),
            external_sheets: vec![FormulaExternalSheet {
                external_link: 0,
                first_sheet: -2,
                last_sheet: -2,
            }]
            .into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        }
        .for_sheet(1);
        let tokens = FormulaParser::new(&[0x5A, 0, 0, 0, 0, 0, 0, 0, 0])
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context).unwrap(),
            "'Current Sheet'!$A$1"
        );
    }

    #[test]
    fn parser_resolves_external_workbook_references_and_names() {
        let context = FormulaResolutionContext {
            worksheet_names: Vec::new().into(),
            supporting_links: vec![FormulaSupportingLink::ExternalWorkbook(0)].into(),
            external_sheets: vec![FormulaExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            }]
            .into(),
            external_books: vec![FormulaExternalBook {
                metadata: Link::workbook(
                    "Book.xlsx",
                    vec!["Data Sheet".to_string()],
                    vec!["Rate".to_string()],
                )
                .unwrap(),
            }]
            .into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };

        let reference = [0x5A, 0, 0, 0, 0, 0, 0, 0, 0];
        let tokens = FormulaParser::new(&reference).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context).unwrap(),
            "'[Book.xlsx]Data Sheet'!$A$1"
        );

        let name = [0x59, 0, 0, 1, 0, 0, 0];
        let tokens = FormulaParser::new(&name).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context).unwrap(),
            "'[Book.xlsx]'!Rate"
        );

        let invalid_name = [0x59, 0, 0, 2, 0, 0, 0];
        let tokens = FormulaParser::new(&invalid_name).parse().unwrap();
        assert!(matches!(
            FormulaConverter::try_tokens_to_string_with_context(&tokens, &context),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[0x59, 0, 0, 0, 0, 0, 0]).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn scalar_errors_compile_and_roundtrip_canonically() {
        for &(literal, code) in FORMULA_ERRORS {
            let compiled = FormulaCompiler::compile(&literal.to_ascii_lowercase()).unwrap();
            assert_eq!(compiled.rgce, vec![ptg_types::PTG_ERR, code]);
            let tokens = FormulaParser::new(&compiled.rgce).parse().unwrap();
            assert_eq!(
                FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
                literal
            );
        }

        let compiled = FormulaCompiler::compile("#DIV/0!+1").unwrap();
        let tokens = FormulaParser::new(&compiled.rgce).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(#DIV/0!+1)"
        );
        assert!(FormulaCompiler::compile("#SPILL!").is_err());
    }

    #[test]
    fn parser_rejects_invalid_scalar_boolean_and_error_values() {
        assert!(matches!(
            FormulaParser::new(&[ptg_types::PTG_BOOL, 2]).parse(),
            Err(Error::InvalidFormula(_))
        ));
        assert!(matches!(
            FormulaParser::new(&[ptg_types::PTG_ERR, 1]).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_consumes_attribute_payloads_and_converts_attr_sum() {
        let attr_sum = [ptg_types::PTG_INT, 1, 0, ptg_types::PTG_ATTR, 0x10, 0, 0];
        let tokens = FormulaParser::new(&attr_sum).parse().unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "SUM(1)"
        );

        let attr_choose = [ptg_types::PTG_ATTR, 0x04, 0x00, 0x00, 0x02, 0x00];
        assert_eq!(
            FormulaParser::new(&attr_choose).parse().unwrap(),
            vec![Token::Attribute(0x04)]
        );

        assert!(matches!(
            FormulaParser::new(&attr_choose[..5]).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_decodes_typed_array_ancillary_values() {
        let mut rgce = vec![0x40];
        rgce.extend_from_slice(&[0; 14]);
        let mut rgcb = Vec::new();
        rgcb.extend_from_slice(&2_u32.to_le_bytes());
        rgcb.extend_from_slice(&2_u32.to_le_bytes());
        rgcb.push(0x00);
        rgcb.extend_from_slice(&1_f64.to_le_bytes());
        rgcb.extend_from_slice(&[0x01, 0x01, 0x00, b'x', 0x00]);
        rgcb.extend_from_slice(&[0x02, 0x01]);
        rgcb.extend_from_slice(&[0x04, 0x07, 0x00, 0x00, 0x00]);

        let tokens = FormulaParser::with_extra(&rgce, &rgcb).parse().unwrap();
        assert_eq!(
            tokens,
            vec![Token::Array {
                rows: 2,
                cols: 2,
                values: vec![
                    ArrayValue::Number(1.0),
                    ArrayValue::String("x".to_string()),
                    ArrayValue::Bool(true),
                    ArrayValue::Error(0x07),
                ],
            }]
        );
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "{1,\"x\";TRUE,#DIV/0!}"
        );
    }

    #[test]
    fn parser_rejects_malformed_array_ancillary_data_without_large_allocation() {
        let mut rgce = vec![0x40];
        rgce.extend_from_slice(&[0; 14]);
        let mut impossible = Vec::new();
        impossible.extend_from_slice(&1_048_576_u32.to_le_bytes());
        impossible.extend_from_slice(&16_384_u32.to_le_bytes());
        assert!(matches!(
            FormulaParser::with_extra(&rgce, &impossible).parse(),
            Err(Error::InvalidFormula(_))
        ));

        let mut invalid_bool = Vec::new();
        invalid_bool.extend_from_slice(&1_u32.to_le_bytes());
        invalid_bool.extend_from_slice(&1_u32.to_le_bytes());
        invalid_bool.extend_from_slice(&[0x02, 0x02]);
        assert!(matches!(
            FormulaParser::with_extra(&rgce, &invalid_bool).parse(),
            Err(Error::InvalidFormula(_))
        ));

        let mut invalid_number = Vec::new();
        invalid_number.extend_from_slice(&1_u32.to_le_bytes());
        invalid_number.extend_from_slice(&1_u32.to_le_bytes());
        invalid_number.push(0x00);
        invalid_number.extend_from_slice(&f64::NEG_INFINITY.to_le_bytes());
        assert!(matches!(
            FormulaParser::with_extra(&rgce, &invalid_number).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn compiler_emits_and_roundtrips_array_constants() {
        let formula = FormulaCompiler::compile("SUM({1,\"x\";TRUE,#N/A})").unwrap();
        assert_eq!(
            &formula.rgce[..15],
            &[0x40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        );
        assert_eq!(&formula.rgcb[..8], &[2, 0, 0, 0, 2, 0, 0, 0]);
        let tokens = FormulaParser::with_extra(&formula.rgce, &formula.rgcb)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "SUM({1,\"x\";TRUE,#N/A})"
        );

        assert!(matches!(
            FormulaCompiler::compile_shared("SUM({1,2})", 0, 0),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn parser_consumes_memory_area_and_cached_ranges() {
        let left = FormulaCompiler::compile("A1").unwrap().rgce;
        let right = FormulaCompiler::compile("B2").unwrap().rgce;
        let expression_len = left.len() + right.len() + 1;
        let mut rgce = vec![0x46, 0, 0, 0, 0];
        rgce.extend_from_slice(&(expression_len as u16).to_le_bytes());
        rgce.extend_from_slice(&left);
        rgce.extend_from_slice(&right);
        rgce.push(ptg_types::PTG_UNION);

        let mut rgcb = Vec::new();
        rgcb.extend_from_slice(&1_u32.to_le_bytes());
        rgcb.extend_from_slice(&0_u32.to_le_bytes());
        rgcb.extend_from_slice(&1_u32.to_le_bytes());
        rgcb.extend_from_slice(&0_u32.to_le_bytes());
        rgcb.extend_from_slice(&1_u32.to_le_bytes());
        let tokens = FormulaParser::with_extra(&rgce, &rgcb).parse().unwrap();
        assert!(matches!(
            &tokens[0],
            Token::Memory {
                kind: MemoryKind::Area,
                expression_bytes: 15,
                cached_ranges,
            } if cached_ranges == &vec![[0, 1, 0, 1]]
        ));
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "(A1,B2)"
        );
    }

    #[test]
    fn parser_rejects_truncated_memory_metadata() {
        let rgce = [0x46, 0, 0, 0, 0, 0, 0];
        let mut rgcb = Vec::new();
        rgcb.extend_from_slice(&u32::MAX.to_le_bytes());
        assert!(matches!(
            FormulaParser::with_extra(&rgce, &rgcb).parse(),
            Err(Error::InvalidFormula(_))
        ));

        let oversized_expression = [0x49, 0x01, 0x00];
        assert!(matches!(
            FormulaParser::new(&oversized_expression).parse(),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn shared_formula_uses_relative_tokens_and_expands_per_target_cell() {
        // Real shared-formula pattern from POI bug66682.xlsb: the C3:C10
        // formula group references the cell one column earlier.
        let formula = FormulaCompiler::compile_shared("B3", 2, 2).unwrap();
        assert_eq!(formula.rgce, vec![0x4C, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF]);

        let anchor_tokens = FormulaParser::with_base_cell(&formula.rgce, 2, 2)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&anchor_tokens).unwrap(),
            "B3"
        );
        let follower_tokens = FormulaParser::with_base_cell(&formula.rgce, 3, 2)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&follower_tokens).unwrap(),
            "B4"
        );
    }

    #[test]
    fn parses_real_poi_shared_formula_definition_losslessly() {
        // BrtShrFmla from POI bug66682.xlsb: C3:C10 refers one column left.
        let bytes = [
            0x02, 0x00, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0x00,
            0x00, 0x00, 0x00,
        ];
        let group = FormulaGroup::parse_shared(&bytes).unwrap();
        assert_eq!(group.kind, GroupKind::Shared);
        assert_eq!(group.range.to_a1(), "C3:C10");
        assert_eq!(group.to_record_data().unwrap(), bytes);

        let tokens = FormulaParser::with_base_cell(&group.formula.rgce, 9, 2)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "B10"
        );
    }

    #[test]
    fn parses_real_poi_array_formula_definition_losslessly() {
        // BrtArrFmla from POI bug66682.xlsb. Its PtgName is retained even
        // when the standalone formula converter cannot resolve that name.
        let bytes = [
            0x08, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x00, 0x00, 0x01, 0x09, 0x00, 0x00, 0x00, 0x23, 0x02, 0x00, 0x00, 0x00, 0x42, 0x01,
            0xFF, 0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let group = FormulaGroup::parse_array(&bytes).unwrap();
        assert_eq!(group.kind, GroupKind::Array);
        assert_eq!(group.range.to_a1(), "C9:C9");
        assert!(group.always_calculate);
        assert_eq!(group.to_record_data().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_ptg_exp_and_array_flags() {
        let malformed = CellParsedFormula {
            rgce: vec![ptg_types::PTG_EXP, 0, 0],
            rgcb: vec![],
        };
        assert!(matches!(
            malformed.exp_cell(),
            Err(Error::InvalidFormula(_))
        ));

        let mut array = FormulaGroup {
            kind: GroupKind::Array,
            range: FormulaRange::new(0, 0, 0, 0).unwrap(),
            formula: FormulaCompiler::compile("1+1").unwrap(),
            always_calculate: false,
        }
        .to_record_data()
        .unwrap();
        array[16] = 0x80;
        assert!(matches!(
            FormulaGroup::parse_array(&array),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn shared_formula_preserves_mixed_absolute_references() {
        let formula = FormulaCompiler::compile_shared("$A1+B$2", 4, 3).unwrap();
        let tokens = FormulaParser::with_base_cell(&formula.rgce, 7, 5)
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string(&tokens).unwrap(),
            "($A4+D$2)"
        );
    }

    #[test]
    fn cell_parsed_formula_accepts_the_empty_token_streams_excel_writes() {
        // `cce == 0`, `cb == 0`: no tokens and no ancillary data.
        let empty = [0_u8; 8];
        let (formula, consumed) = CellParsedFormula::parse(&empty).unwrap();
        assert!(formula.rgce.is_empty());
        assert!(formula.rgcb.is_empty());
        assert_eq!(consumed, empty.len());
        // The empty stream must survive a write so reading cannot lose it.
        assert_eq!(formula.to_bytes().unwrap(), empty);
    }

    #[test]
    fn cell_parsed_formula_rejects_oversized_token_streams() {
        let mut oversized = Vec::new();
        oversized.extend_from_slice(&((MAX_CELL_FORMULA_BYTES as u32) + 1).to_le_bytes());
        oversized.extend_from_slice(&[0; 4]);
        assert!(matches!(
            CellParsedFormula::parse(&oversized),
            Err(Error::InvalidFormula(_))
        ));
    }

    #[test]
    fn truncated_token_is_an_error_instead_of_becoming_unknown_bytes() {
        let error = FormulaParser::new(&[0x44, 0x01]).parse().unwrap_err();
        assert!(matches!(error, Error::InvalidFormula(_)));
    }

    fn resident_table_reference(row_type: TableRowType, columns: TableColumns) -> Token {
        Token::TableReference(TableReference {
            sheet_index: 0,
            row_type: Some(row_type),
            columns: Some(columns),
            square_bracket_space: false,
            comma_space: false,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: Some(7),
            external: None,
        })
    }

    fn table_context() -> FormulaResolutionContext {
        FormulaResolutionContext {
            worksheet_names: vec!["Data".to_string()].into(),
            supporting_links: vec![FormulaSupportingLink::SelfWorkbook].into(),
            external_sheets: vec![FormulaExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            }]
            .into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: vec![
                FormulaTableDefinition::try_new(
                    7,
                    0,
                    "Sales",
                    vec![
                        "Item".to_string(),
                        "Price]Gross".to_string(),
                        "@Tag".to_string(),
                    ],
                )
                .unwrap(),
            ]
            .into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: Some(0),
        }
    }

    #[test]
    fn resolves_resident_structured_references_faithfully() {
        let context = table_context();
        for (row_type, expected) in [
            (TableRowType::Data, "Sales"),
            (TableRowType::All, "Sales[#All]"),
            (TableRowType::Headers, "Sales[#Headers]"),
            (TableRowType::DataAlternate, "Sales[#Data]"),
            (TableRowType::DataAndHeaders, "Sales[[#Headers],[#Data]]"),
            (TableRowType::Totals, "Sales[#Totals]"),
            (TableRowType::DataAndTotals, "Sales[[#Data],[#Totals]]"),
            (TableRowType::Current, "Sales[#This Row]"),
        ] {
            let token = resident_table_reference(row_type, TableColumns::All);
            assert_eq!(
                FormulaConverter::try_tokens_to_string_with_context(&[token], &context).unwrap(),
                expected
            );
        }

        let token = resident_table_reference(
            TableRowType::Current,
            TableColumns::Range { first: 1, last: 2 },
        );
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&[token], &context).unwrap(),
            "Sales[[#This Row],[Price']Gross]:['@Tag]]"
        );

        let mut spaced = resident_table_reference(TableRowType::Current, TableColumns::One(0));
        let Token::TableReference(reference) = &mut spaced else {
            unreachable!()
        };
        reference.square_bracket_space = true;
        reference.comma_space = true;
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&[spaced], &context).unwrap(),
            "Sales[ [#This Row], [Item] ]"
        );
    }

    #[test]
    fn resolves_nonresident_structured_references_with_external_prefix() {
        let context = FormulaResolutionContext {
            worksheet_names: Vec::new().into(),
            supporting_links: vec![FormulaSupportingLink::ExternalWorkbook(0)].into(),
            external_sheets: vec![FormulaExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            }]
            .into(),
            external_books: vec![FormulaExternalBook {
                metadata: Link::workbook("Book.xlsx", vec!["Data Sheet".to_string()], Vec::new())
                    .unwrap(),
            }]
            .into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        let token = Token::TableReference(TableReference {
            sheet_index: 0,
            row_type: None,
            columns: None,
            square_bracket_space: false,
            comma_space: false,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: None,
            external: Some(ExternalTableReference {
                table: "Remote".to_string(),
                row_type: TableRowType::Totals,
                columns: TableNamedColumns::One("Amount".to_string()),
            }),
        });
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&[token], &context).unwrap(),
            "'[Book.xlsx]Data Sheet'!Remote[[#Totals],[Amount]]"
        );
    }

    #[test]
    fn structured_reference_resolution_rejects_ambiguous_and_invalid_metadata() {
        assert!(FormulaTableDefinition::try_new(0, 0, "Sales", vec!["A".into()]).is_err());
        assert!(FormulaTableDefinition::try_new(1, 0, "_xlBad", vec!["A".into()]).is_err());
        assert!(FormulaTableDefinition::try_new(1, 0, "Sales", Vec::new()).is_err());
        assert!(
            FormulaTableDefinition::try_new(1, 0, "Sales", vec!["A".into(), "a".into()]).is_err()
        );

        let token = resident_table_reference(TableRowType::Data, TableColumns::One(3));
        assert!(FormulaConverter::try_tokens_to_string(std::slice::from_ref(&token)).is_err());
        assert!(
            FormulaConverter::try_tokens_to_string_with_context(&[token], &table_context())
                .is_err()
        );

        let mut missing = table_context();
        missing.tables = Vec::new().into();
        assert!(
            FormulaConverter::try_tokens_to_string_with_context(
                &[resident_table_reference(
                    TableRowType::Data,
                    TableColumns::All,
                )],
                &missing,
            )
            .is_err()
        );

        let mut ambiguous = table_context();
        ambiguous.tables = vec![
            ambiguous.tables[0].clone(),
            FormulaTableDefinition::try_new(7, 0, "Other", vec!["A".into()]).unwrap(),
        ]
        .into();
        assert!(
            FormulaConverter::try_tokens_to_string_with_context(
                &[resident_table_reference(
                    TableRowType::Data,
                    TableColumns::All,
                )],
                &ambiguous,
            )
            .is_err()
        );

        let mut wrong_sheet = table_context();
        wrong_sheet.external_sheets = vec![FormulaExternalSheet {
            external_link: 0,
            first_sheet: 0,
            last_sheet: 1,
        }]
        .into();
        assert!(
            FormulaConverter::try_tokens_to_string_with_context(
                &[resident_table_reference(
                    TableRowType::Data,
                    TableColumns::All,
                )],
                &wrong_sheet,
            )
            .is_err()
        );
    }
}

#[cfg(test)]
mod structured_reference_compiler_tests {
    use super::*;

    fn tables() -> Vec<FormulaTableDefinition> {
        vec![
            FormulaTableDefinition::try_new(
                7,
                0,
                "Sales",
                vec![
                    "Item".to_string(),
                    "Price]Gross".to_string(),
                    "@Tag".to_string(),
                    "Amount".to_string(),
                ],
            )
            .unwrap(),
        ]
    }

    #[test]
    fn compiles_parses_and_stringifies_resident_and_nonresident_structured_references() {
        let worksheet_names = vec!["Data".to_string()];
        let tables = tables();
        let defined_names = Vec::new();
        let supporting_links = vec![
            FormulaSupportingLink::ExternalWorkbook(0),
            FormulaSupportingLink::SelfWorkbook,
        ];
        let external_sheets = vec![
            FormulaExternalSheet {
                external_link: 0,
                first_sheet: 0,
                last_sheet: 0,
            },
            FormulaExternalSheet {
                external_link: 1,
                first_sheet: 0,
                last_sheet: 0,
            },
            FormulaExternalSheet {
                external_link: 1,
                first_sheet: 0,
                last_sheet: 0,
            },
        ];
        let external_books = vec![FormulaExternalBook {
            metadata: Link::workbook("Book.xlsx", vec!["Data Sheet".to_string()], Vec::new())
                .unwrap(),
        }];
        let sheet_ranges = std::cell::RefCell::new(Vec::new());
        let compile_context = FormulaCompilationContext {
            worksheet_names: &worksheet_names,
            defined_names: &defined_names,
            tables: &tables,
            supporting_links: &supporting_links,
            external_sheets: &external_sheets,
            external_books: &external_books,
            sheet_ranges: &sheet_ranges,
            current_sheet: 0,
        };
        let resolution_context = FormulaResolutionContext {
            worksheet_names: worksheet_names.clone().into(),
            supporting_links: supporting_links.clone().into(),
            external_sheets: external_sheets.clone().into(),
            external_books: external_books.clone().into(),
            defined_names: Vec::new().into(),
            tables: tables.clone().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: Some(0),
        };

        for source in [
            "Sales",
            "Sales[Item]",
            "Sales[#All]",
            "Sales[[#Headers],[#Data]]",
            "Sales[[#Data],[#Totals]]",
            "Sales[[#This Row],[Price']Gross]:['@Tag]]",
            "Sales[ [#This Row], [Item] ]",
            "'[Book.xlsx]Data Sheet'!Remote[[#Totals],[Amount]]",
        ] {
            let compiled = FormulaCompiler::compile_with_context(source, &compile_context).unwrap();
            let tokens = FormulaParser::with_extra(&compiled.rgce, &compiled.rgcb)
                .parse()
                .unwrap();
            assert_eq!(
                FormulaConverter::try_tokens_to_string_with_context(&tokens, &resolution_context,)
                    .unwrap(),
                source
            );
            assert!(matches!(tokens.as_slice(), [Token::TableReference(_)]));
        }
    }

    #[test]
    fn structured_reference_compiler_rejects_ambiguous_missing_and_unrepresentable_inputs() {
        let worksheet_names = vec!["Data".to_string(), "Other".to_string()];
        let defined_names = Vec::new();
        let supporting_links = Vec::new();
        let external_sheets = Vec::new();
        let external_books = Vec::new();
        let sheet_ranges = std::cell::RefCell::new(Vec::new());
        let base_tables = tables();
        let context = FormulaCompilationContext {
            worksheet_names: &worksheet_names,
            defined_names: &defined_names,
            tables: &base_tables,
            supporting_links: &supporting_links,
            external_sheets: &external_sheets,
            external_books: &external_books,
            sheet_ranges: &sheet_ranges,
            current_sheet: 0,
        };
        for source in [
            "Missing[Item]",
            "Sales[Missing]",
            "Sales[[Amount]:[Item]]",
            "Sales[[Item],[Amount]]",
            "Sales[[#Headers],[#Totals]]",
            "Sales[ Item]",
            "Sales[Item ]",
            "Sales[Bad'x]",
            "'[Book.xlsx]Data Sheet'!Remote[Amount]",
            "Other!Sales[Item]",
        ] {
            assert!(
                FormulaCompiler::compile_with_context(source, &context).is_err(),
                "{source} unexpectedly compiled"
            );
        }

        let ambiguous = vec![
            base_tables[0].clone(),
            FormulaTableDefinition::try_new(8, 0, "sales", vec!["Item".to_string()]).unwrap(),
        ];
        let ambiguous_context = FormulaCompilationContext {
            tables: &ambiguous,
            ..context
        };
        assert!(FormulaCompiler::compile_with_context("Sales[Item]", &ambiguous_context).is_err());

        let wrong_sheet =
            vec![FormulaTableDefinition::try_new(7, 1, "Sales", vec!["Item".to_string()]).unwrap()];
        let wrong_sheet_context = FormulaCompilationContext {
            tables: &wrong_sheet,
            ..context
        };
        assert!(
            FormulaCompiler::compile_with_context("Sales[Item]", &wrong_sheet_context).is_err()
        );
    }
}

#[cfg(test)]
mod pivot_name_resolution_tests {
    use super::*;

    fn references() -> Vec<FormulaPivotNameReference> {
        vec![
            FormulaPivotNameReference::Field {
                name: "Sales".to_string(),
                aggregation: None,
            },
            FormulaPivotNameReference::Field {
                name: "Gross Profit".to_string(),
                aggregation: Some(FormulaPivotAggregation::Average),
            },
            FormulaPivotNameReference::Item {
                field_name: "Region".to_string(),
                item: FormulaPivotItemReference::Name("North".to_string()),
            },
            FormulaPivotNameReference::Item {
                field_name: "Sales Region".to_string(),
                item: FormulaPivotItemReference::Name("O'Brien".to_string()),
            },
            FormulaPivotNameReference::Item {
                field_name: "Quarter".to_string(),
                item: FormulaPivotItemReference::AbsolutePosition(2),
            },
            FormulaPivotNameReference::Item {
                field_name: "Quarter".to_string(),
                item: FormulaPivotItemReference::RelativePosition(1),
            },
            FormulaPivotNameReference::Item {
                field_name: "Quarter".to_string(),
                item: FormulaPivotItemReference::RelativePosition(-1),
            },
        ]
    }

    fn scope() -> FormulaPivotNameScope {
        FormulaPivotNameScope::try_new(7, 1, "Sales Pivot".to_string(), references()).unwrap()
    }

    fn pivot_context() -> FormulaResolutionContext {
        FormulaResolutionContext {
            worksheet_names: vec!["Data".to_string(), "Report".to_string()].into(),
            supporting_links: Vec::new().into(),
            external_sheets: Vec::new().into(),
            external_books: Vec::new().into(),
            defined_names: Vec::new().into(),
            tables: Vec::new().into(),
            pivot_views: vec![
                FormulaPivotViewDefinition::try_new(7, 1, "Sales Pivot".to_string()).unwrap(),
            ]
            .into(),
            pivot_name_scopes: vec![scope()].into(),
            active_pivot_scope: Some((7, 1, "Sales Pivot".to_string())),
            current_sheet: Some(1),
        }
    }

    fn render(index: u32, context: &FormulaResolutionContext) -> Result<String> {
        FormulaConverter::try_tokens_to_string_with_context(&[Token::PivotName(index)], context)
    }

    #[test]
    fn resolves_pivot_names_to_faithful_field_and_item_syntax() {
        let context = pivot_context();
        assert_eq!(render(0, &context).unwrap(), "Sales");
        assert_eq!(render(1, &context).unwrap(), "AVERAGE('Gross Profit')");
        assert_eq!(render(2, &context).unwrap(), "Region[North]");
        assert_eq!(render(3, &context).unwrap(), "'Sales Region'['O''Brien']");
        assert_eq!(render(4, &context).unwrap(), "Quarter[2]");
        assert_eq!(render(5, &context).unwrap(), "Quarter[+1]");
        assert_eq!(render(6, &context).unwrap(), "Quarter[-1]");
    }

    #[test]
    fn rejects_missing_ambiguous_cross_sheet_and_out_of_range_pivot_metadata() {
        assert!(FormulaConverter::try_tokens_to_string(&[Token::PivotName(0)]).is_err());

        let mut context = pivot_context();
        assert!(render(7, &context).is_err());
        context.current_sheet = Some(0);
        assert!(render(0, &context).is_err());

        let mut context = pivot_context();
        context.pivot_views = vec![
            FormulaPivotViewDefinition::try_new(7, 1, "Sales Pivot".to_string()).unwrap(),
            FormulaPivotViewDefinition::try_new(7, 1, "sales pivot".to_string()).unwrap(),
        ]
        .into();
        assert!(render(0, &context).is_err());

        let mut context = pivot_context();
        context.pivot_name_scopes = vec![scope(), scope()].into();
        assert!(render(0, &context).is_err());

        let mut context = pivot_context();
        context.active_pivot_scope = Some((8, 1, "Sales Pivot".to_string()));
        assert!(render(0, &context).is_err());
    }

    #[test]
    fn validates_bounded_pivot_names_and_positions() {
        assert!(FormulaPivotViewDefinition::try_new(1, 0, String::new()).is_err());
        assert!(
            FormulaPivotNameScope::try_new(
                1,
                0,
                "Pivot".to_string(),
                vec![FormulaPivotNameReference::Item {
                    field_name: "Quarter".to_string(),
                    item: FormulaPivotItemReference::AbsolutePosition(0),
                }],
            )
            .is_err()
        );
        assert!(
            FormulaPivotNameScope::try_new(
                1,
                0,
                "Pivot".to_string(),
                vec![FormulaPivotNameReference::Item {
                    field_name: "Quarter".to_string(),
                    item: FormulaPivotItemReference::RelativePosition(0),
                }],
            )
            .is_err()
        );
        assert!(
            FormulaPivotNameScope::try_new(
                1,
                0,
                "Pivot".to_string(),
                vec![FormulaPivotNameReference::Field {
                    name: "bad\0field".to_string(),
                    aggregation: None,
                }],
            )
            .is_err()
        );
    }
}
