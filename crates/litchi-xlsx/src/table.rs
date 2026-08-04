//! Bounded SpreadsheetML table model and XML codec.
//!
//! Tables in Excel provide structured references and enhanced formatting for data ranges.

use std::collections::HashSet;
use std::fmt::Write as FmtWrite;

use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::conditional_formatting::IconSet;
use crate::error::{Error, Result, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;
use crate::sort::{SortBy, SortCondition, SortMethod, SortState};
use litchi_ooxml_common::mce::{MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;
const MAX_COLUMNS: usize = 16_384;
const MAX_SORT_CONDITIONS: usize = 64;
const MAX_TEXT_BYTES: usize = 1024 * 1024;
const MAX_EXCEL_COLUMN: u32 = 16_384;
const MAX_EXCEL_ROW: u32 = 1_048_576;

/// Table style information for visual formatting.
#[derive(Debug, Clone)]
pub struct TableStyleInfo {
    /// Style name (e.g., "TableStyleMedium2")
    pub name: Option<String>,
    /// Show first column with special formatting
    pub show_first_column: Option<bool>,
    /// Show last column with special formatting
    pub show_last_column: Option<bool>,
    /// Show alternating row stripes
    pub show_row_stripes: Option<bool>,
    /// Show alternating column stripes
    pub show_column_stripes: Option<bool>,
}

impl TableStyleInfo {
    pub fn new() -> Self {
        Self {
            name: None,
            show_first_column: None,
            show_last_column: None,
            show_row_stripes: None,
            show_column_stripes: None,
        }
    }

    /// Parse table style info from XML tag.
    pub fn parse(tag: &str) -> Option<Self> {
        let name = Self::extract_attribute(tag, "name");
        let show_first_column = Self::extract_attribute(tag, "showFirstColumn")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        let show_last_column = Self::extract_attribute(tag, "showLastColumn")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        let show_row_stripes = Self::extract_attribute(tag, "showRowStripes")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));
        let show_column_stripes = Self::extract_attribute(tag, "showColumnStripes")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"));

        Some(Self {
            name,
            show_first_column,
            show_last_column,
            show_row_stripes,
            show_column_stripes,
        })
    }

    fn extract_attribute(tag: &str, attr: &str) -> Option<String> {
        let search_str = format!("{}=\"", attr);
        let start = tag.find(&search_str)? + search_str.len();
        let end = tag[start..].find('"')? + start;
        Some(tag[start..end].to_string())
    }
}

impl Default for TableStyleInfo {
    fn default() -> Self {
        Self::new()
    }
}

/// Totals row function types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TotalsRowFunction {
    Sum,
    Min,
    Max,
    Average,
    Count,
    CountNums,
    StdDev,
    Var,
    Custom,
}

impl TotalsRowFunction {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Sum => "sum",
            Self::Min => "min",
            Self::Max => "max",
            Self::Average => "average",
            Self::Count => "count",
            Self::CountNums => "countNums",
            Self::StdDev => "stdDev",
            Self::Var => "var",
            Self::Custom => "custom",
        }
    }

    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "sum" => Some(Self::Sum),
            "min" => Some(Self::Min),
            "max" => Some(Self::Max),
            "average" => Some(Self::Average),
            "count" => Some(Self::Count),
            "countNums" => Some(Self::CountNums),
            "stdDev" => Some(Self::StdDev),
            "var" => Some(Self::Var),
            "custom" => Some(Self::Custom),
            _ => None,
        }
    }
}

/// A formula for a table column (calculated or totals row).
#[derive(Debug, Clone)]
pub struct TableFormula {
    /// Whether this is an array formula
    pub array: Option<bool>,
    /// Formula text
    pub text: String,
}

/// A single column in a table.
#[derive(Debug, Clone)]
pub struct TableColumn {
    /// Column ID (1-based)
    pub id: u32,
    /// Unique name (optional)
    pub unique_name: Option<String>,
    /// Display name
    pub name: String,
    /// Totals row function
    pub totals_row_function: Option<TotalsRowFunction>,
    /// Totals row label (for custom totals)
    pub totals_row_label: Option<String>,
    /// Calculated column formula
    pub calculated_column_formula: Option<TableFormula>,
    /// Totals row formula
    pub totals_row_formula: Option<TableFormula>,
}

impl TableColumn {
    /// Create a new table column.
    pub fn new(id: u32, name: impl Into<String>) -> Self {
        Self {
            id,
            unique_name: None,
            name: name.into(),
            totals_row_function: None,
            totals_row_label: None,
            calculated_column_formula: None,
            totals_row_formula: None,
        }
    }
}

/// Table type enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TableType {
    Worksheet,
    Xml,
    QueryTable,
}

impl TableType {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Worksheet => "worksheet",
            Self::Xml => "xml",
            Self::QueryTable => "queryTable",
        }
    }

    pub fn parse(s: &str) -> Option<Self> {
        match s {
            "worksheet" => Some(Self::Worksheet),
            "xml" => Some(Self::Xml),
            "queryTable" => Some(Self::QueryTable),
            _ => None,
        }
    }
}

/// An Excel table (structured data range).
///
/// Tables provide structured references in formulas and enhanced formatting.
#[derive(Debug, Clone)]
pub struct Table {
    /// Table ID (unique within workbook)
    pub id: u32,
    /// Internal name (used in formulas)
    pub name: String,
    /// Display name (shown in Excel UI)
    pub display_name: String,
    /// Comment/description
    pub comment: Option<String>,
    /// Cell range (e.g., "A1:D10")
    pub ref_range: String,
    /// Table type
    pub table_type: Option<TableType>,
    /// Number of header rows (usually 1)
    pub header_row_count: Option<u32>,
    /// Number of totals rows
    pub totals_row_count: Option<u32>,
    /// Whether totals row is shown
    pub totals_row_shown: Option<bool>,
    /// Published to server
    pub published: Option<bool>,
    /// Table columns
    pub columns: Vec<TableColumn>,
    /// Auto-filter configuration
    pub auto_filter_range: Option<String>,
    /// Sort state
    pub sort_state: Option<SortState>,
    /// Table style information
    pub style_info: Option<TableStyleInfo>,
}

impl Table {
    /// Create a new table with the given ID, name, and range.
    pub fn new(id: u32, name: impl Into<String>, ref_range: impl Into<String>) -> Self {
        let name_str = name.into();
        Self {
            id,
            name: name_str.clone(),
            display_name: name_str,
            comment: None,
            ref_range: ref_range.into(),
            table_type: None,
            header_row_count: Some(1),
            totals_row_count: None,
            totals_row_shown: None,
            published: None,
            columns: Vec::new(),
            auto_filter_range: None,
            sort_state: None,
            style_info: None,
        }
    }

    /// Initialize columns from range (creates default Column1, Column2, etc.).
    pub fn initialize_columns(&mut self) {
        if !self.columns.is_empty() {
            return;
        }

        // Parse range to determine column count
        if let Some((min_col, _min_row, max_col, _max_row)) = parse_range(&self.ref_range) {
            let Some(col_count) = max_col
                .checked_sub(min_col)
                .and_then(|width| width.checked_add(1))
                .filter(|count| *count as usize <= MAX_COLUMNS)
            else {
                return;
            };
            for i in 0..col_count {
                let Some(col_id) = min_col.checked_add(i) else {
                    return;
                };
                self.columns
                    .push(TableColumn::new(col_id, format!("Column{}", col_id)));
            }
        }

        // Set auto-filter if we have headers
        if self.header_row_count.unwrap_or(0) > 0 && self.auto_filter_range.is_none() {
            self.auto_filter_range = Some(self.ref_range.clone());
        }
    }

    /// Get column names.
    pub fn column_names(&self) -> Vec<&str> {
        self.columns.iter().map(|c| c.name.as_str()).collect()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TableContext {
    Root,
    AutoFilter,
    SortState,
    TableColumns,
    TableColumn,
    Formula(FormulaKind),
    Other,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FormulaKind {
    Calculated,
    Totals,
}

struct TableParser {
    table: Table,
    expected_columns: Option<u32>,
    pending_column: Option<TableColumn>,
    pending_formula: Option<TableFormula>,
    column_ids: HashSet<u32>,
    column_names: HashSet<String>,
    saw_auto_filter: bool,
    saw_sort_state: bool,
    saw_table_columns: bool,
    saw_style_info: bool,
    saw_calculated_formula: bool,
    saw_totals_formula: bool,
}

impl TableParser {
    fn from_root(element: &BytesStart<'_>, decoder: Decoder) -> Result<Self> {
        let id = required_u32(element, b"id", decoder, "table ID")?;
        if id == 0 {
            return Err(invalid("table ID must be positive"));
        }
        let display_name = required_string(element, b"displayName", decoder, "table displayName")?;
        if display_name.is_empty() {
            return Err(invalid("table displayName cannot be empty"));
        }
        let name =
            attribute_value(element, b"name", decoder)?.unwrap_or_else(|| display_name.clone());
        if name.is_empty() {
            return Err(invalid("table name cannot be empty"));
        }
        let ref_range = required_string(element, b"ref", decoder, "table reference")?;
        validate_table_range(&ref_range, "table reference")?;
        let table_type = attribute_value(element, b"tableType", decoder)?
            .map(|value| {
                TableType::parse(&value)
                    .ok_or_else(|| invalid(format!("invalid table type '{value}'")))
            })
            .transpose()?;
        Ok(Self {
            table: Table {
                id,
                name,
                display_name,
                comment: attribute_value(element, b"comment", decoder)?,
                ref_range,
                table_type,
                header_row_count: Some(
                    optional_u32(element, b"headerRowCount", decoder, "table headerRowCount")?
                        .unwrap_or(1),
                ),
                totals_row_count: optional_u32(
                    element,
                    b"totalsRowCount",
                    decoder,
                    "table totalsRowCount",
                )?,
                totals_row_shown: optional_bool(
                    element,
                    b"totalsRowShown",
                    decoder,
                    "table totalsRowShown",
                )?,
                published: optional_bool(element, b"published", decoder, "table published")?,
                columns: Vec::new(),
                auto_filter_range: None,
                sort_state: None,
                style_info: None,
            },
            expected_columns: None,
            pending_column: None,
            pending_formula: None,
            column_ids: HashSet::new(),
            column_names: HashSet::new(),
            saw_auto_filter: false,
            saw_sort_state: false,
            saw_table_columns: false,
            saw_style_info: false,
            saw_calculated_formula: false,
            saw_totals_formula: false,
        })
    }

    fn parse(xml: &[u8]) -> Result<Option<Table>> {
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut parser: Option<Self> = None;
        let mut stack = Vec::new();
        let mut closed_root = false;
        let mut events = 0usize;
        loop {
            events = events
                .checked_add(1)
                .ok_or_else(|| invalid("table XML event count overflow"))?;
            if events > MAX_EVENTS {
                return Err(limit("table XML event count"));
            }
            let decoder = reader.decoder();
            let event = reader.read_event().map_err(xml_error)?.into_owned();
            reject_unsafe_event(&event)?;
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root {
                        return Err(invalid("table XML contains multiple root elements"));
                    }
                    if !is_spreadsheetml_name(&namespace, element.name(), b"table") {
                        return Ok(None);
                    }
                    parser = Some(Self::from_root(&element, decoder)?);
                    if stack.len() >= MAX_DEPTH {
                        return Err(limit("table XML depth"));
                    }
                    stack.push(TableContext::Root);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"table") {
                        return Ok(None);
                    }
                    let parser = Self::from_root(&element, decoder)?;
                    parser.validate_root()?;
                    return Ok(Some(parser.table));
                },
                Event::Start(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("table XML is missing its root"))?;
                    let context = parser
                        .as_mut()
                        .ok_or_else(|| invalid("table parser is not initialized"))?
                        .start(parent, &namespace, &element, decoder)?;
                    if stack.len() >= MAX_DEPTH {
                        return Err(limit("table XML depth"));
                    }
                    stack.push(context);
                },
                Event::Empty(element) => {
                    let parent = *stack
                        .last()
                        .ok_or_else(|| invalid("table XML is missing its root"))?;
                    let parser = parser
                        .as_mut()
                        .ok_or_else(|| invalid("table parser is not initialized"))?;
                    let context = parser.start(parent, &namespace, &element, decoder)?;
                    parser.finish(context)?;
                },
                Event::Text(text) => {
                    if matches!(stack.last(), Some(TableContext::Formula(_))) {
                        parser
                            .as_mut()
                            .ok_or_else(|| invalid("table parser is not initialized"))?
                            .push_formula_text(&text.decode().map_err(xml_error)?)?;
                    }
                },
                Event::CData(text) => {
                    if matches!(stack.last(), Some(TableContext::Formula(_))) {
                        parser
                            .as_mut()
                            .ok_or_else(|| invalid("table parser is not initialized"))?
                            .push_formula_text(&text.decode().map_err(xml_error)?)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if matches!(stack.last(), Some(TableContext::Formula(_))) {
                        parser
                            .as_mut()
                            .ok_or_else(|| invalid("table parser is not initialized"))?
                            .push_formula_text(&decode_xml_reference(&reference)?)?;
                    }
                },
                Event::End(element) => {
                    let context = stack.pop().ok_or_else(|| {
                        invalid("table XML has a closing element outside its root")
                    })?;
                    parser
                        .as_mut()
                        .ok_or_else(|| invalid("table parser is not initialized"))?
                        .finish(context)?;
                    if context == TableContext::Root {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"table") {
                            return Err(invalid("table XML has an invalid root closing element"));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if parser.is_none() => return Ok(None),
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid("table XML has a missing or unterminated root"));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(parser.map(|parser| parser.table))
    }

    fn start(
        &mut self,
        parent: TableContext,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<TableContext> {
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"autoFilter")
        {
            mark_once(&mut self.saw_auto_filter, "table autoFilter")?;
            let range = required_string(element, b"ref", decoder, "table autoFilter reference")?;
            validate_table_range(&range, "table autoFilter reference")?;
            self.table.auto_filter_range = Some(range);
            return Ok(TableContext::AutoFilter);
        }
        if parent == TableContext::AutoFilter
            && is_spreadsheetml_name(namespace, element.name(), b"sortState")
        {
            mark_once(&mut self.saw_sort_state, "table sortState")?;
            self.table.sort_state = Some(parse_sort_state(element, decoder)?);
            return Ok(TableContext::SortState);
        }
        if parent == TableContext::SortState
            && is_spreadsheetml_name(namespace, element.name(), b"sortCondition")
        {
            let condition = parse_sort_condition(element, decoder)?;
            let sort_state = self
                .table
                .sort_state
                .as_mut()
                .ok_or_else(|| invalid("table sort condition outside sortState"))?;
            if sort_state.conditions.len() >= MAX_SORT_CONDITIONS {
                return Err(limit("table sort conditions"));
            }
            sort_state
                .conditions
                .try_reserve(1)
                .map_err(|source| allocation("table sort conditions", source))?;
            sort_state.conditions.push(condition);
            return Ok(TableContext::Other);
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"tableColumns")
        {
            mark_once(&mut self.saw_table_columns, "tableColumns")?;
            let count = required_u32(element, b"count", decoder, "tableColumns count")?;
            if usize::try_from(count).map_or(true, |count| count > MAX_COLUMNS) {
                return Err(limit("table columns"));
            }
            self.expected_columns = Some(count);
            return Ok(TableContext::TableColumns);
        }
        if parent == TableContext::TableColumns
            && is_spreadsheetml_name(namespace, element.name(), b"tableColumn")
        {
            if self.pending_column.is_some() {
                return Err(invalid("nested table column"));
            }
            if self.table.columns.len() >= MAX_COLUMNS {
                return Err(limit("table columns"));
            }
            let id = required_u32(element, b"id", decoder, "table column ID")?;
            if id == 0 || !self.column_ids.insert(id) {
                return Err(invalid(format!(
                    "invalid or duplicate table column ID {id}"
                )));
            }
            let name = required_string(element, b"name", decoder, "table column name")?;
            if name.is_empty() || !self.column_names.insert(name.to_ascii_lowercase()) {
                return Err(invalid(format!(
                    "empty or duplicate table column name '{name}'"
                )));
            }
            let totals_row_function = attribute_value(element, b"totalsRowFunction", decoder)?
                .map(|value| {
                    TotalsRowFunction::parse(&value).ok_or_else(|| {
                        invalid(format!("invalid table totals-row function '{value}'"))
                    })
                })
                .transpose()?;
            self.pending_column = Some(TableColumn {
                id,
                unique_name: attribute_value(element, b"uniqueName", decoder)?,
                name,
                totals_row_function,
                totals_row_label: attribute_value(element, b"totalsRowLabel", decoder)?,
                calculated_column_formula: None,
                totals_row_formula: None,
            });
            self.saw_calculated_formula = false;
            self.saw_totals_formula = false;
            return Ok(TableContext::TableColumn);
        }
        if parent == TableContext::TableColumn
            && is_spreadsheetml_name(namespace, element.name(), b"calculatedColumnFormula")
        {
            mark_once(&mut self.saw_calculated_formula, "calculatedColumnFormula")?;
            self.start_formula(element, decoder)?;
            return Ok(TableContext::Formula(FormulaKind::Calculated));
        }
        if parent == TableContext::TableColumn
            && is_spreadsheetml_name(namespace, element.name(), b"totalsRowFormula")
        {
            mark_once(&mut self.saw_totals_formula, "totalsRowFormula")?;
            self.start_formula(element, decoder)?;
            return Ok(TableContext::Formula(FormulaKind::Totals));
        }
        if parent == TableContext::Root
            && is_spreadsheetml_name(namespace, element.name(), b"tableStyleInfo")
        {
            mark_once(&mut self.saw_style_info, "tableStyleInfo")?;
            self.table.style_info = Some(TableStyleInfo {
                name: attribute_value(element, b"name", decoder)?,
                show_first_column: optional_bool(
                    element,
                    b"showFirstColumn",
                    decoder,
                    "table style showFirstColumn",
                )?,
                show_last_column: optional_bool(
                    element,
                    b"showLastColumn",
                    decoder,
                    "table style showLastColumn",
                )?,
                show_row_stripes: optional_bool(
                    element,
                    b"showRowStripes",
                    decoder,
                    "table style showRowStripes",
                )?,
                show_column_stripes: optional_bool(
                    element,
                    b"showColumnStripes",
                    decoder,
                    "table style showColumnStripes",
                )?,
            });
            return Ok(TableContext::Other);
        }
        Ok(TableContext::Other)
    }

    fn start_formula(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.pending_formula.is_some() {
            return Err(invalid("nested table formula"));
        }
        self.pending_formula = Some(TableFormula {
            array: optional_bool(element, b"array", decoder, "table formula array")?,
            text: String::new(),
        });
        Ok(())
    }

    fn push_formula_text(&mut self, text: &str) -> Result<()> {
        let formula = self
            .pending_formula
            .as_mut()
            .ok_or_else(|| invalid("table formula text outside a formula element"))?;
        let length = formula
            .text
            .len()
            .checked_add(text.len())
            .ok_or_else(|| limit("table formula text"))?;
        if length > MAX_TEXT_BYTES {
            return Err(limit("table formula text"));
        }
        formula
            .text
            .try_reserve(text.len())
            .map_err(|source| allocation("table formula text", source))?;
        formula.text.push_str(text);
        Ok(())
    }

    fn finish(&mut self, context: TableContext) -> Result<()> {
        match context {
            TableContext::Formula(kind) => {
                let formula = self
                    .pending_formula
                    .take()
                    .ok_or_else(|| invalid("missing table formula"))?;
                let column = self
                    .pending_column
                    .as_mut()
                    .ok_or_else(|| invalid("table formula outside a table column"))?;
                match kind {
                    FormulaKind::Calculated => column.calculated_column_formula = Some(formula),
                    FormulaKind::Totals => column.totals_row_formula = Some(formula),
                }
                Ok(())
            },
            TableContext::TableColumn => {
                self.table.columns.push(
                    self.pending_column
                        .take()
                        .ok_or_else(|| invalid("missing pending table column"))?,
                );
                Ok(())
            },
            TableContext::TableColumns => validate_count(
                self.expected_columns,
                self.table.columns.len(),
                "tableColumns",
            ),
            TableContext::Root => self.validate_root(),
            _ => Ok(()),
        }
    }

    fn validate_root(&self) -> Result<()> {
        if !self.saw_table_columns {
            return Err(invalid("table is missing its required tableColumns"));
        }
        let width = table_range_width(&self.table.ref_range)?;
        if usize::try_from(width) != Ok(self.table.columns.len()) {
            return Err(invalid(format!(
                "table range spans {width} columns, but {} table columns were defined",
                self.table.columns.len()
            )));
        }
        if let Some(filter_range) = self.table.auto_filter_range.as_deref() {
            ensure_range_contains(
                &self.table.ref_range,
                filter_range,
                "table autoFilter reference",
            )?;
        }
        if let Some(sort_state) = self.table.sort_state.as_ref() {
            let container = self
                .table
                .auto_filter_range
                .as_deref()
                .unwrap_or(&self.table.ref_range);
            ensure_range_contains(
                container,
                &sort_state.ref_range,
                "table sortState reference",
            )?;
            for condition in &sort_state.conditions {
                ensure_range_contains(
                    &sort_state.ref_range,
                    &condition.ref_range,
                    "table sort condition reference",
                )?;
            }
        }
        Ok(())
    }
}

pub fn parse_table_xml(xml: impl AsRef<[u8]>) -> Result<Option<Table>> {
    let xml = xml.as_ref();
    if xml.len() > MAX_XML_BYTES {
        return Err(limit("table XML bytes"));
    }
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..MceLimits::default()
    };
    let processed = process_markup_compatibility(xml, &MceCapabilities::default(), &limits)?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(limit("processed table XML bytes"));
    }
    TableParser::parse(processed.xml.as_ref())
}

/// Validate a table before it is authored as SpreadsheetML.
pub fn validate_table(table: &Table) -> Result<()> {
    if table.id == 0 {
        return Err(invalid("table ID must be positive"));
    }
    bounded_nonempty(&table.name, "table name")?;
    bounded_nonempty(&table.display_name, "table display name")?;
    bounded(&table.comment, "table comment")?;
    validate_table_range(&table.ref_range, "table reference")?;
    if table.columns.is_empty() {
        return Err(invalid("table must contain at least one column"));
    }
    if table.columns.len() > MAX_COLUMNS {
        return Err(limit("table columns"));
    }
    let width = table_range_width(&table.ref_range)?;
    if usize::try_from(width) != Ok(table.columns.len()) {
        return Err(invalid(format!(
            "table range spans {width} columns, but {} columns were provided",
            table.columns.len()
        )));
    }
    if let Some(auto_filter_range) = table.auto_filter_range.as_deref() {
        ensure_range_contains(
            &table.ref_range,
            auto_filter_range,
            "table autoFilter reference",
        )?;
    }
    if let Some(sort_state) = table.sort_state.as_ref() {
        let auto_filter_range = table
            .auto_filter_range
            .as_deref()
            .ok_or_else(|| invalid("table sort state requires an auto-filter"))?;
        ensure_range_contains(
            auto_filter_range,
            &sort_state.ref_range,
            "table sortState reference",
        )?;
        if sort_state.conditions.len() > MAX_SORT_CONDITIONS {
            return Err(limit("table sort conditions"));
        }
        bounded_text(&sort_state.ref_range, "table sortState reference")?;
        for condition in &sort_state.conditions {
            ensure_range_contains(
                &sort_state.ref_range,
                &condition.ref_range,
                "table sort condition reference",
            )?;
            condition
                .validate()
                .map_err(|error| invalid(error.to_string()))?;
            if let Some(custom_list) = condition.custom_list.as_deref() {
                bounded_text(custom_list, "table custom sort list")?;
            }
        }
    }
    let mut ids = HashSet::with_capacity(table.columns.len());
    let mut names = HashSet::with_capacity(table.columns.len());
    for column in &table.columns {
        if column.id == 0 || !ids.insert(column.id) {
            return Err(invalid(format!(
                "invalid or duplicate table column ID {}",
                column.id
            )));
        }
        bounded_nonempty(&column.name, "table column name")?;
        if !names.insert(column.name.to_ascii_lowercase()) {
            return Err(invalid(format!(
                "empty or duplicate table column name '{}'",
                column.name
            )));
        }
        bounded(&column.unique_name, "table column unique name")?;
        bounded(&column.totals_row_label, "table totals row label")?;
        validate_formula(
            column.calculated_column_formula.as_ref(),
            "calculated column",
        )?;
        validate_formula(column.totals_row_formula.as_ref(), "totals row")?;
    }
    if let Some(style) = table.style_info.as_ref() {
        bounded(&style.name, "table style name")?;
    }
    Ok(())
}

/// Write a validated table part using the transitional SpreadsheetML namespace.
pub fn write_table_xml(table: &Table) -> Result<Vec<u8>> {
    validate_table(table)?;
    let mut output = BoundedXml::new();
    output.literal(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
    )?;
    output.attribute_display("id", table.id)?;
    output.attribute("name", &table.name)?;
    output.attribute("displayName", &table.display_name)?;
    output.attribute("ref", &table.ref_range)?;
    if let Some(value) = table.comment.as_deref() {
        output.attribute("comment", value)?;
    }
    if let Some(value) = table.table_type {
        output.attribute("tableType", value.as_str())?;
    }
    if let Some(value) = table.header_row_count {
        output.attribute_display("headerRowCount", value)?;
    }
    if let Some(value) = table.totals_row_count {
        output.attribute_display("totalsRowCount", value)?;
    }
    if let Some(value) = table.totals_row_shown {
        output.attribute_bool("totalsRowShown", value)?;
    }
    if let Some(value) = table.published {
        output.attribute_bool("published", value)?;
    }
    output.literal(">")?;

    if let Some(auto_filter_range) = table.auto_filter_range.as_deref() {
        output.attribute_element_start("autoFilter", &[("ref", auto_filter_range)])?;
        if let Some(sort_state) = table.sort_state.as_ref() {
            output.literal(">")?;
            write_sort_state(&mut output, sort_state)?;
            output.literal("</autoFilter>")?;
        } else {
            output.close_empty_element()?;
        }
    }

    output.literal("<tableColumns")?;
    output.attribute_display("count", table.columns.len())?;
    output.literal(">")?;
    for column in &table.columns {
        write_table_column(&mut output, column)?;
    }
    output.literal("</tableColumns>")?;

    if let Some(style) = table.style_info.as_ref() {
        write_style_info(&mut output, style)?;
    }
    output.literal("</table>")?;
    Ok(output.finish())
}

/// Serialize a table part as UTF-8 text for the legacy host writer API.
pub fn serialize_table(table: &Table) -> Result<String> {
    String::from_utf8(write_table_xml(table)?).map_err(|error| xml_error(error.to_string()))
}

fn write_sort_state(output: &mut BoundedXml, sort_state: &SortState) -> Result<()> {
    output.literal("<sortState")?;
    output.attribute("ref", &sort_state.ref_range)?;
    if let Some(value) = sort_state.column_sort {
        output.attribute_bool("columnSort", value)?;
    }
    if let Some(value) = sort_state.case_sensitive {
        output.attribute_bool("caseSensitive", value)?;
    }
    if let Some(value) = sort_state.sort_method {
        output.attribute("sortMethod", value.as_str())?;
    }
    if sort_state.conditions.is_empty() {
        output.close_empty_element()?;
        return Ok(());
    }
    output.literal(">")?;
    for condition in &sort_state.conditions {
        condition
            .validate()
            .map_err(|error| invalid(error.to_string()))?;
        output.literal("<sortCondition")?;
        output.attribute("ref", &condition.ref_range)?;
        if let Some(value) = condition.descending {
            output.attribute_bool("descending", value)?;
        }
        if let Some(value) = condition.sort_by {
            output.attribute("sortBy", value.as_str())?;
        }
        if let Some(value) = condition.custom_list.as_deref() {
            output.attribute("customList", value)?;
        }
        if let Some(value) = condition.dxf_id {
            output.attribute_display("dxfId", value)?;
        }
        if let Some(value) = condition.icon_set {
            output.attribute("iconSet", value.as_str())?;
        }
        if let Some(value) = condition.icon_id {
            output.attribute_display("iconId", value)?;
        }
        output.close_empty_element()?;
    }
    output.literal("</sortState>")
}

fn write_table_column(output: &mut BoundedXml, column: &TableColumn) -> Result<()> {
    output.literal("<tableColumn")?;
    output.attribute_display("id", column.id)?;
    output.attribute("name", &column.name)?;
    if let Some(value) = column.unique_name.as_deref() {
        output.attribute("uniqueName", value)?;
    }
    if let Some(value) = column.totals_row_function {
        output.attribute("totalsRowFunction", value.as_str())?;
    }
    if let Some(value) = column.totals_row_label.as_deref() {
        output.attribute("totalsRowLabel", value)?;
    }
    if column.calculated_column_formula.is_none() && column.totals_row_formula.is_none() {
        return output.close_empty_element();
    }
    output.literal(">")?;
    if let Some(value) = column.calculated_column_formula.as_ref() {
        write_formula(output, "calculatedColumnFormula", value)?;
    }
    if let Some(value) = column.totals_row_formula.as_ref() {
        write_formula(output, "totalsRowFormula", value)?;
    }
    output.literal("</tableColumn>")
}

fn write_formula(output: &mut BoundedXml, name: &str, formula: &TableFormula) -> Result<()> {
    output.literal("<")?;
    output.literal(name)?;
    if let Some(value) = formula.array {
        output.attribute_bool("array", value)?;
    }
    output.literal(">")?;
    output.escaped(&formula.text)?;
    output.literal("</")?;
    output.literal(name)?;
    output.literal(">")
}

fn write_style_info(output: &mut BoundedXml, style: &TableStyleInfo) -> Result<()> {
    output.literal("<tableStyleInfo")?;
    if let Some(value) = style.name.as_deref() {
        output.attribute("name", value)?;
    }
    if let Some(value) = style.show_first_column {
        output.attribute_bool("showFirstColumn", value)?;
    }
    if let Some(value) = style.show_last_column {
        output.attribute_bool("showLastColumn", value)?;
    }
    if let Some(value) = style.show_row_stripes {
        output.attribute_bool("showRowStripes", value)?;
    }
    if let Some(value) = style.show_column_stripes {
        output.attribute_bool("showColumnStripes", value)?;
    }
    output.close_empty_element()
}

struct BoundedXml {
    bytes: Vec<u8>,
}

impl BoundedXml {
    fn new() -> Self {
        Self {
            bytes: Vec::with_capacity(2048),
        }
    }

    fn append(&mut self, bytes: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| limit("serialized table XML bytes"))?;
        if length > MAX_OUTPUT_BYTES {
            return Err(limit("serialized table XML bytes"));
        }
        self.bytes
            .try_reserve(bytes.len())
            .map_err(|source| allocation("serialized table XML bytes", source))?;
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }

    fn literal(&mut self, value: &str) -> Result<()> {
        self.append(value.as_bytes())
    }

    fn escaped(&mut self, value: &str) -> Result<()> {
        for character in value.chars() {
            match character {
                '&' => self.literal("&amp;")?,
                '<' => self.literal("&lt;")?,
                '>' => self.literal("&gt;")?,
                '"' => self.literal("&quot;")?,
                '\'' => self.literal("&apos;")?,
                _ => {
                    let mut bytes = [0; 4];
                    self.append(character.encode_utf8(&mut bytes).as_bytes())?;
                },
            }
        }
        Ok(())
    }

    fn attribute(&mut self, name: &str, value: &str) -> Result<()> {
        self.literal(" ")?;
        self.literal(name)?;
        self.literal("=\"")?;
        self.escaped(value)?;
        self.literal("\"")
    }

    fn attribute_display(&mut self, name: &str, value: impl std::fmt::Display) -> Result<()> {
        let mut text = String::new();
        write!(&mut text, "{value}").map_err(|_| invalid("failed to format table XML value"))?;
        self.attribute(name, &text)
    }

    fn attribute_bool(&mut self, name: &str, value: bool) -> Result<()> {
        self.attribute(name, if value { "1" } else { "0" })
    }

    fn attribute_element_start(&mut self, name: &str, attributes: &[(&str, &str)]) -> Result<()> {
        self.literal("<")?;
        self.literal(name)?;
        for (attribute, value) in attributes {
            self.attribute(attribute, value)?;
        }
        Ok(())
    }

    fn close_empty_element(&mut self) -> Result<()> {
        self.literal("/>")
    }

    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn validate_formula(formula: Option<&TableFormula>, description: &str) -> Result<()> {
    if let Some(formula) = formula {
        bounded_text(&formula.text, description)?;
    }
    Ok(())
}

fn bounded(value: &Option<String>, description: &str) -> Result<()> {
    if let Some(value) = value.as_deref() {
        bounded_text(value, description)?;
    }
    Ok(())
}

fn bounded_text(value: &str, description: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(limit(description));
    }
    Ok(())
}

fn bounded_nonempty(value: &str, description: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    bounded_text(value, description)
}

fn parse_sort_state(element: &BytesStart<'_>, decoder: Decoder) -> Result<SortState> {
    let range = required_string(element, b"ref", decoder, "table sortState reference")?;
    validate_table_range(&range, "table sortState reference")?;
    let sort_method = attribute_value(element, b"sortMethod", decoder)?
        .map(|value| value.parse::<SortMethod>())
        .transpose()
        .map_err(|error| invalid(error.to_string()))?;
    Ok(SortState {
        ref_range: range,
        column_sort: optional_bool(
            element,
            b"columnSort",
            decoder,
            "table sortState columnSort",
        )?,
        case_sensitive: optional_bool(
            element,
            b"caseSensitive",
            decoder,
            "table sortState caseSensitive",
        )?,
        sort_method,
        conditions: Vec::new(),
    })
}

fn parse_sort_condition(element: &BytesStart<'_>, decoder: Decoder) -> Result<SortCondition> {
    let range = required_string(element, b"ref", decoder, "table sort condition reference")?;
    validate_table_range(&range, "table sort condition reference")?;
    let sort_by = attribute_value(element, b"sortBy", decoder)?
        .map(|value| value.parse::<SortBy>())
        .transpose()
        .map_err(|error| invalid(error.to_string()))?;
    let condition = SortCondition {
        ref_range: range,
        descending: optional_bool(
            element,
            b"descending",
            decoder,
            "table sort condition descending",
        )?,
        sort_by,
        custom_list: attribute_value(element, b"customList", decoder)?,
        dxf_id: optional_u32(element, b"dxfId", decoder, "table sort condition dxfId")?,
        icon_set: attribute_value(element, b"iconSet", decoder)?
            .map(|value| value.parse::<IconSet>())
            .transpose()
            .map_err(|error| invalid(error.to_string()))?,
        icon_id: optional_u32(element, b"iconId", decoder, "table sort condition iconId")?,
    };
    condition
        .validate()
        .map_err(|error| invalid(error.to_string()))?;
    Ok(condition)
}

fn attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    let value = unqualified_attribute_value(element, name, decoder)?;
    if value
        .as_deref()
        .is_some_and(|value| value.len() > MAX_TEXT_BYTES)
    {
        return Err(limit("table attribute text"));
    }
    Ok(value)
}

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    optional_u32(element, name, decoder, description)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid {description} '{value}'"))),
        })
        .transpose()
}

fn mark_once(seen: &mut bool, description: &str) -> Result<()> {
    if *seen {
        return Err(invalid(format!("duplicate {description} element")));
    }
    *seen = true;
    Ok(())
}

fn validate_count(expected: Option<u32>, actual: usize, description: &str) -> Result<()> {
    if let Some(expected) = expected
        && usize::try_from(expected) != Ok(actual)
    {
        return Err(invalid(format!(
            "{description} count is {expected}, but {actual} elements were found"
        )));
    }
    Ok(())
}

fn validate_table_range(range: &str, description: &str) -> Result<()> {
    let mut references = range.split(':');
    let first = references
        .next()
        .ok_or_else(|| invalid(format!("empty {description}")))?;
    let (first_column, first_row) = checked_cell_ref(first, description)?;
    if let Some(second) = references.next() {
        let (second_column, second_row) = checked_cell_ref(second, description)?;
        if second_column < first_column || second_row < first_row {
            return Err(invalid(format!(
                "{description} range '{range}' is descending"
            )));
        }
    }
    if references.next().is_some() {
        return Err(invalid(format!("invalid {description} range '{range}'")));
    }
    Ok(())
}

fn ensure_range_contains(container: &str, nested: &str, description: &str) -> Result<()> {
    let (container_start, container_end) = table_range_bounds(container)?;
    let (nested_start, nested_end) = table_range_bounds(nested)?;
    if nested_start.0 < container_start.0
        || nested_start.1 < container_start.1
        || nested_end.0 > container_end.0
        || nested_end.1 > container_end.1
    {
        return Err(invalid(format!(
            "{description} '{nested}' is outside containing range '{container}'"
        )));
    }
    Ok(())
}

fn table_range_bounds(range: &str) -> Result<((u32, u32), (u32, u32))> {
    let mut references = range.split(':');
    let first = references
        .next()
        .ok_or_else(|| invalid("empty table range"))?;
    let second = references.next().unwrap_or(first);
    if references.next().is_some() {
        return Err(invalid(format!("invalid table range '{range}'")));
    }
    let start = checked_cell_ref(first, "table range")?;
    let end = checked_cell_ref(second, "table range")?;
    if end.0 < start.0 || end.1 < start.1 {
        return Err(invalid(format!("table range '{range}' is descending")));
    }
    Ok((start, end))
}

fn table_range_width(range: &str) -> Result<u32> {
    let ((first_column, _), (second_column, _)) = table_range_bounds(range)?;
    second_column
        .checked_sub(first_column)
        .and_then(|width| width.checked_add(1))
        .ok_or_else(|| invalid(format!("table range '{range}' has descending columns")))
}

fn checked_cell_ref(reference: &str, description: &str) -> Result<(u32, u32)> {
    parse_cell_ref(&reference.replace('$', "")).ok_or_else(|| {
        invalid(format!(
            "invalid {description} cell reference '{reference}'"
        ))
    })
}

/// Parse a cell range like "A1:D10" into (min_col, min_row, max_col, max_row).
/// Returns 1-based indices.
fn parse_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let normalized = range.replace('$', "");
    let parts: Vec<&str> = normalized.split(':').collect();
    if parts.len() != 2 {
        return None;
    }

    let (min_col, min_row) = parse_cell_ref(parts[0])?;
    let (max_col, max_row) = parse_cell_ref(parts[1])?;

    Some((min_col, min_row, max_col, max_row))
}

/// Parse a cell reference like "A1" into (col, row) with 1-based indices.
fn parse_cell_ref(cell_ref: &str) -> Option<(u32, u32)> {
    let bytes = cell_ref.as_bytes();
    let column_end = bytes.iter().position(u8::is_ascii_digit)?;
    if column_end == 0 || column_end == bytes.len() {
        return None;
    }
    let mut column = 0u32;
    for byte in &bytes[..column_end] {
        if !byte.is_ascii_alphabetic() {
            return None;
        }
        let digit = u32::from(byte.to_ascii_uppercase() - b'A' + 1);
        column = column.checked_mul(26)?.checked_add(digit)?;
    }
    if column == 0 || column > MAX_EXCEL_COLUMN {
        return None;
    }
    let row_bytes = &bytes[column_end..];
    if !row_bytes.iter().all(u8::is_ascii_digit) {
        return None;
    }
    let row = std::str::from_utf8(row_bytes).ok()?.parse::<u32>().ok()?;
    if row == 0 || row > MAX_EXCEL_ROW {
        return None;
    }
    Some((column, row))
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_)) {
        return Err(invalid("table XML must not contain a DOCTYPE"));
    }
    Ok(())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

fn limit(resource: &str) -> Error {
    invalid(format!("table {resource} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_style_info_new() {
        let style = TableStyleInfo::new();
        assert!(style.name.is_none());
        assert!(style.show_first_column.is_none());
        assert!(style.show_last_column.is_none());
        assert!(style.show_row_stripes.is_none());
        assert!(style.show_column_stripes.is_none());
    }

    #[test]
    fn test_table_style_info_default() {
        let style: TableStyleInfo = Default::default();
        assert!(style.name.is_none());
    }

    #[test]
    fn test_table_style_info_parse() {
        let tag = r#"name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0""#;
        let style = TableStyleInfo::parse(tag).unwrap();
        assert_eq!(style.name, Some("TableStyleMedium2".to_string()));
        assert_eq!(style.show_first_column, Some(true));
        assert_eq!(style.show_last_column, Some(false));
        assert_eq!(style.show_row_stripes, Some(true));
        assert_eq!(style.show_column_stripes, Some(false));
    }

    #[test]
    fn test_table_style_info_parse_partial() {
        let tag = r#"name="TableStyleLight1" showRowStripes="true""#;
        let style = TableStyleInfo::parse(tag).unwrap();
        assert_eq!(style.name, Some("TableStyleLight1".to_string()));
        assert_eq!(style.show_row_stripes, Some(true));
        assert!(style.show_first_column.is_none());
    }

    #[test]
    fn test_totals_row_function_as_str() {
        assert_eq!(TotalsRowFunction::Sum.as_str(), "sum");
        assert_eq!(TotalsRowFunction::Min.as_str(), "min");
        assert_eq!(TotalsRowFunction::Max.as_str(), "max");
        assert_eq!(TotalsRowFunction::Average.as_str(), "average");
        assert_eq!(TotalsRowFunction::Count.as_str(), "count");
        assert_eq!(TotalsRowFunction::CountNums.as_str(), "countNums");
        assert_eq!(TotalsRowFunction::StdDev.as_str(), "stdDev");
        assert_eq!(TotalsRowFunction::Var.as_str(), "var");
        assert_eq!(TotalsRowFunction::Custom.as_str(), "custom");
    }

    #[test]
    fn test_totals_row_function_parse() {
        assert_eq!(
            TotalsRowFunction::parse("sum"),
            Some(TotalsRowFunction::Sum)
        );
        assert_eq!(
            TotalsRowFunction::parse("min"),
            Some(TotalsRowFunction::Min)
        );
        assert_eq!(
            TotalsRowFunction::parse("max"),
            Some(TotalsRowFunction::Max)
        );
        assert_eq!(
            TotalsRowFunction::parse("average"),
            Some(TotalsRowFunction::Average)
        );
        assert_eq!(
            TotalsRowFunction::parse("count"),
            Some(TotalsRowFunction::Count)
        );
        assert_eq!(
            TotalsRowFunction::parse("countNums"),
            Some(TotalsRowFunction::CountNums)
        );
        assert_eq!(
            TotalsRowFunction::parse("stdDev"),
            Some(TotalsRowFunction::StdDev)
        );
        assert_eq!(
            TotalsRowFunction::parse("var"),
            Some(TotalsRowFunction::Var)
        );
        assert_eq!(
            TotalsRowFunction::parse("custom"),
            Some(TotalsRowFunction::Custom)
        );
        assert_eq!(TotalsRowFunction::parse("invalid"), None);
    }

    #[test]
    fn test_table_column_new() {
        let col = TableColumn::new(1u32, "Sales");
        assert_eq!(col.id, 1);
        assert_eq!(col.name, "Sales");
        assert!(col.unique_name.is_none());
        assert!(col.totals_row_function.is_none());
        assert!(col.totals_row_label.is_none());
        assert!(col.calculated_column_formula.is_none());
        assert!(col.totals_row_formula.is_none());
    }

    #[test]
    fn test_table_type_as_str() {
        assert_eq!(TableType::Worksheet.as_str(), "worksheet");
        assert_eq!(TableType::Xml.as_str(), "xml");
        assert_eq!(TableType::QueryTable.as_str(), "queryTable");
    }

    #[test]
    fn test_table_type_parse() {
        assert_eq!(TableType::parse("worksheet"), Some(TableType::Worksheet));
        assert_eq!(TableType::parse("xml"), Some(TableType::Xml));
        assert_eq!(TableType::parse("queryTable"), Some(TableType::QueryTable));
        assert_eq!(TableType::parse("invalid"), None);
    }

    #[test]
    fn test_table_new() {
        let table = Table::new(1u32, "Table1", "A1:D10");
        assert_eq!(table.id, 1);
        assert_eq!(table.name, "Table1");
        assert_eq!(table.display_name, "Table1");
        assert_eq!(table.ref_range, "A1:D10");
        assert_eq!(table.header_row_count, Some(1));
        assert!(table.columns.is_empty());
        assert!(table.comment.is_none());
        assert!(table.table_type.is_none());
    }

    #[test]
    fn test_table_initialize_columns() {
        let mut table = Table::new(1u32, "Table1", "A1:D10");
        table.initialize_columns();
        assert_eq!(table.columns.len(), 4);
        assert_eq!(table.columns[0].name, "Column1");
        assert_eq!(table.columns[3].name, "Column4");
        assert!(table.auto_filter_range.is_some());
    }

    #[test]
    fn test_table_column_names() {
        let mut table = Table::new(1u32, "Table1", "A1:C5");
        table.initialize_columns();
        let names = table.column_names();
        assert_eq!(names, vec!["Column1", "Column2", "Column3"]);
    }

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), Some((1, 1)));
        assert_eq!(parse_cell_ref("B2"), Some((2, 2)));
        assert_eq!(parse_cell_ref("Z10"), Some((26, 10)));
        assert_eq!(parse_cell_ref("AA1"), Some((27, 1)));
        assert_eq!(parse_cell_ref("AB100"), Some((28, 100)));
    }

    #[test]
    fn test_parse_range() {
        assert_eq!(parse_range("A1:D10"), Some((1, 1, 4, 10)));
        assert_eq!(parse_range("B2:C5"), Some((2, 2, 3, 5)));
        assert_eq!(parse_range("A1"), None); // Missing colon
        assert_eq!(parse_range(""), None);
    }

    #[test]
    fn parses_prefixed_strict_table_definition() {
        let xml = r#"<s:table xmlns:s="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:f="urn:foreign" id="3" name="Sales_Internal" displayName="Sales &amp; Margin"
                ref="$A$1:$B$3" tableType="worksheet" headerRowCount="1"
                totalsRowCount="1" totalsRowShown="true" published="0">
                <f:tableColumns count="1"><s:tableColumn id="99" name="Ignored"/></f:tableColumns>
                <s:autoFilter ref="$A$1:$B$3"><s:sortState ref="$A$2:$B$3" caseSensitive="0" sortMethod="pinYin">
                    <s:sortCondition ref="$B$2:$B$3" descending="1" sortBy="value" customList="High,Low"/>
                </s:sortState></s:autoFilter>
                <s:tableColumns count="2">
                    <s:tableColumn id="1" name="Region &amp; Area" uniqueName="Region" totalsRowLabel="Total"/>
                    <s:tableColumn id="2" name="Sales" totalsRowFunction="sum">
                        <s:calculatedColumnFormula array="false">SUM([Sales])&amp;1</s:calculatedColumnFormula>
                        <s:totalsRowFormula>SUBTOTAL(109,[Sales])</s:totalsRowFormula>
                    </s:tableColumn>
                </s:tableColumns>
                <s:tableStyleInfo name="TableStyleMedium2" showFirstColumn="0"
                    showLastColumn="1" showRowStripes="true" showColumnStripes="false"/>
            </s:table>"#;
        let table = parse_table_xml(xml).unwrap().unwrap();

        assert_eq!(table.id, 3);
        assert_eq!(table.display_name, "Sales & Margin");
        assert_eq!(table.columns.len(), 2);
        assert_eq!(table.columns[0].name, "Region & Area");
        assert_eq!(
            table.columns[1]
                .calculated_column_formula
                .as_ref()
                .unwrap()
                .text,
            "SUM([Sales])&1"
        );
        assert_eq!(table.sort_state.as_ref().unwrap().conditions.len(), 1);
        assert_eq!(
            table.style_info.as_ref().unwrap().show_last_column,
            Some(true)
        );
        assert_eq!(table.published, Some(false));
    }

    #[test]
    fn rejects_malformed_table_definitions() {
        const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let root = |body: &str, attributes: &str| {
            format!(
                r#"<table xmlns="{S}" id="1" displayName="Table1" ref="A1:B3" {attributes}>{body}</table>"#
            )
        };
        for xml in [
            root("", ""),
            root(
                r#"<tableColumns count="2"><tableColumn id="1" name="One"/></tableColumns>"#,
                "",
            ),
            root(
                r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="1" name="Two"/></tableColumns>"#,
                "",
            ),
            root(
                r#"<tableColumns count="2"><tableColumn id="1" name="Same"/><tableColumn id="2" name="same"/></tableColumns>"#,
                "",
            ),
            root(
                r#"<autoFilter ref="A1:C3"/><tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns>"#,
                "",
            ),
            root(
                r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two" totalsRowFunction="median"/></tableColumns>"#,
                "",
            ),
            root(
                r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns><tableStyleInfo showRowStripes="yes"/>"#,
                "",
            ),
            root(
                r#"<tableColumns count="2"><tableColumn id="1" name="One"/><tableColumn id="2" name="Two"/></tableColumns>"#,
                r#"tableType="bad""#,
            ),
        ] {
            assert!(parse_table_xml(&xml).is_err(), "accepted {xml}");
        }
        assert!(
            parse_table_xml(r#"<table xmlns="urn:foreign" id="1" displayName="T" ref="A1"/>"#)
                .unwrap()
                .is_none()
        );
    }

    fn writable_table() -> Table {
        let mut table = Table::new(1, "TestTable", "A1:B5");
        table.display_name = "Test Table".into();
        table.comment = Some("Test comment".into());
        table.header_row_count = Some(1);
        table.totals_row_count = Some(1);
        table.totals_row_shown = Some(true);
        table.columns = vec![
            TableColumn::new(1, "Column A"),
            TableColumn::new(2, "Column B"),
        ];
        table
    }

    #[test]
    fn writes_table_attributes_columns_and_style() {
        let mut table = writable_table();
        let mut style = TableStyleInfo::new();
        style.name = Some("TableStyleMedium2".into());
        style.show_first_column = Some(true);
        style.show_last_column = Some(false);
        style.show_row_stripes = Some(true);
        style.show_column_stripes = Some(false);
        table.style_info = Some(style);

        let xml = serialize_table(&table).unwrap();
        assert!(xml.contains(r#"id="1""#));
        assert!(xml.contains(r#"name="TestTable""#));
        assert!(xml.contains(r#"displayName="Test Table""#));
        assert!(xml.contains(r#"ref="A1:B5""#));
        assert!(xml.contains(r#"comment="Test comment""#));
        assert!(xml.contains(r#"<tableColumns count="2">"#));
        assert!(xml.contains(r#"<tableColumn id="1" name="Column A"/>"#));
        assert!(xml.contains(r#"<tableColumn id="2" name="Column B"/>"#));
        assert!(xml.contains(r#"<tableStyleInfo name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"#));
    }

    #[test]
    fn writes_auto_filter_sort_and_formulas_losslessly() {
        let mut table = writable_table();
        table.published = Some(true);
        table.auto_filter_range = Some("A1:B5".into());
        table.sort_state = Some(SortState {
            ref_range: "A2:B5".into(),
            column_sort: Some(false),
            case_sensitive: Some(true),
            sort_method: Some(SortMethod::PinYin),
            conditions: vec![SortCondition {
                ref_range: "B2:B5".into(),
                descending: Some(true),
                sort_by: Some(SortBy::Icon),
                custom_list: Some("High,Low".into()),
                dxf_id: None,
                icon_set: Some(IconSet::ThreeArrows),
                icon_id: Some(2),
            }],
        });
        table.columns[0].calculated_column_formula = Some(TableFormula {
            array: Some(false),
            text: "=[@[Column B]]*2".into(),
        });
        table.columns[1].totals_row_function = Some(TotalsRowFunction::Sum);
        table.columns[1].totals_row_formula = Some(TableFormula {
            array: None,
            text: "SUBTOTAL(109,[Column B])".into(),
        });

        let xml = serialize_table(&table).unwrap();
        let parsed = parse_table_xml(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(parsed.published, Some(true));
        assert_eq!(parsed.auto_filter_range.as_deref(), Some("A1:B5"));
        assert_eq!(parsed.sort_state.as_ref().unwrap().conditions.len(), 1);
        assert_eq!(
            parsed.sort_state.unwrap().sort_method,
            Some(SortMethod::PinYin)
        );
        assert_eq!(
            parsed.columns[0]
                .calculated_column_formula
                .as_ref()
                .unwrap()
                .text,
            "=[@[Column B]]*2"
        );
        assert_eq!(
            parsed.columns[1].totals_row_function,
            Some(TotalsRowFunction::Sum)
        );
        assert_eq!(
            parsed.columns[1].totals_row_formula.as_ref().unwrap().text,
            "SUBTOTAL(109,[Column B])"
        );
    }

    #[test]
    fn escapes_writable_table_values() {
        let mut table = writable_table();
        table.name = "Table<>&\"'".into();
        table.display_name = "Test <Table>".into();
        table.columns[0].name = "A & B".into();
        table.columns[0].calculated_column_formula = Some(TableFormula {
            array: None,
            text: "IF(A1<2,\"yes\",\"no\")".into(),
        });

        let xml = serialize_table(&table).unwrap();
        assert!(xml.contains("Table&lt;&gt;&amp;&quot;&apos;"));
        assert!(xml.contains("A &amp; B"));
        assert!(
            xml.contains("IF(A1&lt;2,&quot;yes&quot;,&quot;no&quot;)</calculatedColumnFormula>")
        );
        assert!(parse_table_xml(xml.as_bytes()).is_ok());
    }

    #[test]
    fn rejects_invalid_writable_table_models() {
        let mut table = writable_table();
        table.columns[1].id = 1;
        assert!(serialize_table(&table).is_err());

        let mut table = writable_table();
        table.auto_filter_range = Some("A1:C5".into());
        assert!(serialize_table(&table).is_err());

        let mut table = writable_table();
        table.sort_state = Some(SortState::new("A2:B5"));
        assert!(serialize_table(&table).is_err());

        let mut table = writable_table();
        table.columns.clear();
        assert!(serialize_table(&table).is_err());
    }
}
