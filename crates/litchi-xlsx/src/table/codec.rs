//! Bounded `SpreadsheetML` table XML codec.

use std::collections::HashSet;

use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::conditional_formatting::IconSet;
use crate::error::{Result, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;
use crate::sort::{SortBy, SortCondition, SortMethod, SortState};
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};

use super::model::{
    Table, TableColumn, TableFormula, TableStyleInfo, TableType, TotalsRowFunction,
    ensure_range_contains, table_range_width, validate_table_range,
};
use super::{
    MAX_COLUMNS, MAX_DEPTH, MAX_EVENTS, MAX_SORT_CONDITIONS, MAX_TEXT_BYTES, MAX_XML_BYTES, limit,
    xml_error,
};

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
    let limits = Limits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..Limits::default()
    };
    let processed = process_markup_compatibility(xml, &Capabilities::default(), &limits)?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(limit("processed table XML bytes"));
    }
    TableParser::parse(processed.xml.as_ref())
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

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_)) {
        return Err(invalid("table XML must not contain a DOCTYPE"));
    }
    Ok(())
}
