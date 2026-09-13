//! Streaming `SpreadsheetML` event codec.

use std::collections::HashSet;

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use litchi_sheet::{COLUMNS, Cell as Address, Column as ColumnIndex, ROWS, Rect, Row as RowIndex};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event, attributes::Attribute};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::formula::Range as FormulaRange;
use super::super::namespace::is_spreadsheetml_name;
use super::super::strings::decode_spreadsheet_text;
use super::model::{
    Context, MAX_CELL_CHARACTERS, MAX_CELL_STYLE, MAX_COLUMN_STYLE, MAX_ENCODED_CELL_BYTES,
    MAX_FORMULA_CHARACTERS, MAX_METADATA_INDEX, MAX_XML_DEPTH, MAX_XML_EVENTS, Parser, PendingCell,
    PendingRow, RawCell, RawFormula, RawFormulaKind, TextTarget, merge_successor,
};
use super::semantic::{materialize, resolve_shared_formulas};
use super::validation::{
    current, optional_bool, optional_f64, optional_u32, parse_a1, parse_defaults_element,
    parse_one_based_row, required_u32, text_target,
};
use super::x14ac;
use crate::cell::{Store, Text};
use crate::column::{self, Assignments, Flags};
use crate::error::{Result, allocation, invalid};
use crate::layout::{self, Defaults};
use crate::row;

use super::{SourceEventSpan, SourceHandoff};

pub(super) fn parse_processed_defaults(
    content: &str,
    mut descent: Option<layout::Descent>,
) -> Result<Option<Defaults>> {
    let mut reader = NsReader::from_reader(content.as_bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack = Vec::new();
    let mut closed_root = false;
    let mut defaults = None;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let decoder = reader.decoder();
        let resolver = reader.resolver();
        match event {
            Event::Start(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet") {
                    return Err(invalid(
                        "worksheet XML must have one SpreadsheetML worksheet root",
                    ));
                }
                stack.push(Context::Worksheet);
            },
            Event::Empty(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet") {
                    return Err(invalid(
                        "worksheet XML must have one SpreadsheetML worksheet root",
                    ));
                }
                return Err(invalid("worksheet root cannot be empty"));
            },
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "worksheet XML exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let parent = current(&stack)?;
                if parent == Context::SheetFormat {
                    return Err(invalid(
                        "worksheet sheetFormatPr must not have child elements",
                    ));
                }
                if parent == Context::Worksheet
                    && is_spreadsheetml_name(&namespace, element.name(), b"sheetFormatPr")
                {
                    if defaults.is_some() {
                        return Err(invalid("worksheet has duplicate sheetFormatPr elements"));
                    }
                    defaults = Some(parse_defaults_element(
                        &element,
                        decoder,
                        resolver,
                        descent.take(),
                    )?);
                    stack.push(Context::SheetFormat);
                } else {
                    stack.push(Context::Other);
                }
            },
            Event::Empty(element) => {
                let parent = current(&stack)?;
                if parent == Context::SheetFormat {
                    return Err(invalid(
                        "worksheet sheetFormatPr must not have child elements",
                    ));
                }
                if parent == Context::Worksheet
                    && is_spreadsheetml_name(&namespace, element.name(), b"sheetFormatPr")
                {
                    if defaults.is_some() {
                        return Err(invalid("worksheet has duplicate sheetFormatPr elements"));
                    }
                    defaults = Some(parse_defaults_element(
                        &element,
                        decoder,
                        resolver,
                        descent.take(),
                    )?);
                }
            },
            Event::Text(value) if stack.last() == Some(&Context::SheetFormat) => {
                if !value
                    .decode()
                    .map_err(|error| invalid(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid("worksheet sheetFormatPr cannot contain text"));
                }
            },
            Event::CData(_) if stack.last() == Some(&Context::SheetFormat) => {
                return Err(invalid("worksheet sheetFormatPr cannot contain CDATA"));
            },
            Event::GeneralRef(value) => {
                decode_xml_reference(&value)?;
                if stack.last() == Some(&Context::SheetFormat) {
                    return Err(invalid(
                        "worksheet sheetFormatPr cannot contain character references",
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::End(element) => {
                let ended = stack.pop().ok_or_else(|| {
                    invalid("worksheet XML has a closing element outside its root")
                })?;
                if ended == Context::Worksheet {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"worksheet") {
                        return Err(invalid("worksheet XML has an invalid root closing element"));
                    }
                    closed_root = true;
                } else if ended == Context::SheetFormat
                    && !is_spreadsheetml_name(&namespace, element.name(), b"sheetFormatPr")
                {
                    return Err(invalid(
                        "worksheet XML has an invalid sheetFormatPr closing element",
                    ));
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err(invalid(
                    "worksheet XML has a missing or unterminated SpreadsheetML worksheet root",
                ));
            },
            Event::Eof => break,
            Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::Decl(_) => {},
        }
    }

    Ok(defaults)
}

// Check the complete attribute list once, as the old first `r` lookup did.
// Decode `r` at encounter time, but retain the other fields undecoded so their
// errors still follow coordinate, style, cell-metadata and value-metadata checks.
#[derive(Debug)]
struct CellAttributeView<'a> {
    reference: Option<String>,
    style: Option<Attribute<'a>>,
    cell_metadata: Option<Attribute<'a>>,
    value_metadata: Option<Attribute<'a>>,
    cell_type: Option<Attribute<'a>>,
}

fn scan_cell_attributes<'a>(
    element: &'a BytesStart<'_>,
    decoder: Decoder,
) -> Result<CellAttributeView<'a>> {
    let mut attributes = CellAttributeView {
        reference: None,
        style: None,
        cell_metadata: None,
        value_metadata: None,
        cell_type: None,
    };
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| litchi_ooxml_common::XmlError::Malformed(error.to_string()))?;
        if attribute.key.prefix().is_some() {
            continue;
        }
        match attribute.key.local_name().as_ref() {
            b"r" => {
                attributes.reference = Some(decode_cell_attribute(attribute, decoder)?);
            },
            b"s" => attributes.style = Some(attribute),
            b"cm" => attributes.cell_metadata = Some(attribute),
            b"vm" => attributes.value_metadata = Some(attribute),
            b"t" => attributes.cell_type = Some(attribute),
            _ => {},
        }
    }
    Ok(attributes)
}

fn decode_cell_attribute(
    attribute: Attribute<'_>,
    decoder: Decoder,
) -> litchi_ooxml_common::xml::Result<String> {
    attribute
        .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
        .map(|value| value.into_owned())
        .map_err(|error| litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

fn parse_cell_u32(
    attribute: Option<Attribute<'_>>,
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    attribute
        .map(|attribute| {
            let value = decode_cell_attribute(attribute, decoder)?;
            value
                .parse::<u32>()
                .map_err(|_source| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

impl Parser {
    fn new(extensions: x14ac::Values) -> Self {
        Self {
            cells: Vec::new(),
            rows: Vec::new(),
            columns: None,
            defaults: None,
            extensions,
            declared_extent: None,
            row: None,
            cell: None,
            seen_rows: HashSet::new(),
            previous_row: 0,
            seen_dimension: false,
            seen_defaults: false,
            seen_columns: false,
            column_records: 0,
            seen_sheet_data: false,
            merges: Vec::new(),
            merge_count: None,
            seen_merges: false,
            merge_window_closed: false,
        }
    }

    pub(super) fn parse<'a, F>(
        content: &str,
        strings: F,
        extensions: x14ac::Values,
    ) -> Result<Store>
    where
        F: FnOnce() -> Result<Option<&'a [Text]>>,
    {
        let mut reader = NsReader::from_reader(content.as_bytes());
        reader.config_mut().check_end_names = true;
        let mut parser = Self::new(extensions);
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let event = reader
                .read_event()
                .map_err(|error| invalid(error.to_string()))?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            let decoder = reader.decoder();
            let resolver = reader.resolver();
            if parser.transition(
                &mut stack,
                &mut closed_root,
                &namespace,
                event,
                decoder,
                resolver,
            )? {
                break;
            }
        }

        parser.finish_parse(strings)
    }

    fn finish_parse<'a, F>(mut self, strings: F) -> Result<Store>
    where
        F: FnOnce() -> Result<Option<&'a [Text]>>,
    {
        resolve_shared_formulas(&mut self.cells)?;
        let needs_strings = self
            .cells
            .iter()
            .any(|cell| cell.cell_type.as_deref() == Some("s"));
        let strings = if needs_strings { strings()? } else { None };
        let mut cells = Vec::new();
        cells
            .try_reserve(self.cells.len())
            .map_err(|source| allocation("sparse worksheet cells", source))?;
        let declared_extent = self.declared_extent;
        let rows = self.rows;
        let columns = column::resolve(self.columns)?;
        let defaults = self.defaults;
        let merges = self.merges;
        for cell in self.cells {
            cells.push(materialize(cell, strings)?);
        }
        Store::from_unsorted(cells, rows, columns, defaults, merges, declared_extent)
    }

    /// Advance the ordinary worksheet parser by one XML event.
    ///
    /// The source-backed shared reader uses this same transition function so
    /// its speculative parser cannot drift from the established materialized
    /// parser. `true` is returned only for a complete end-of-file event.
    fn transition(
        &mut self,
        stack: &mut Vec<Context>,
        closed_root: &mut bool,
        namespace: &ResolveResult<'_>,
        event: Event<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<bool> {
        match event {
            Event::Start(element) if stack.is_empty() => {
                if *closed_root || !is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
                    return Err(invalid(
                        "worksheet XML must have one SpreadsheetML worksheet root",
                    ));
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("worksheet parser element stack", source))?;
                stack.push(Context::Worksheet);
            },
            Event::Empty(element) if stack.is_empty() => {
                if *closed_root || !is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
                    return Err(invalid(
                        "worksheet XML must have one SpreadsheetML worksheet root",
                    ));
                }
                *closed_root = true;
            },
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "worksheet XML exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
                let parent = current(stack)?;
                let child = self.start(parent, namespace, &element, decoder, resolver)?;
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("worksheet parser element stack", source))?;
                stack.push(child);
            },
            Event::Empty(element) => {
                let parent = current(stack)?;
                let child = self.start(parent, namespace, &element, decoder, resolver)?;
                self.finish(child)?;
            },
            Event::Text(value) => {
                if matches!(stack.last(), Some(Context::SheetFormat | Context::Merge))
                    && !value
                        .decode()
                        .map_err(|error| invalid(error.to_string()))?
                        .trim()
                        .is_empty()
                {
                    return Err(invalid("worksheet leaf property cannot contain text"));
                }
                if let Some(target) = text_target(stack) {
                    self.push_text(
                        target,
                        &value.decode().map_err(|error| invalid(error.to_string()))?,
                    )?;
                }
            },
            Event::CData(value) => {
                if matches!(stack.last(), Some(Context::SheetFormat | Context::Merge)) {
                    return Err(invalid("worksheet leaf property cannot contain CDATA"));
                }
                if let Some(target) = text_target(stack) {
                    self.push_text(
                        target,
                        &value.decode().map_err(|error| invalid(error.to_string()))?,
                    )?;
                }
            },
            Event::GeneralRef(value) => {
                if matches!(stack.last(), Some(Context::SheetFormat | Context::Merge)) {
                    return Err(invalid(
                        "worksheet leaf property cannot contain character references",
                    ));
                }
                if let Some(target) = text_target(stack) {
                    self.push_text(target, &decode_xml_reference(&value)?)?;
                }
            },
            Event::End(element) => {
                let ended = stack.pop().ok_or_else(|| {
                    invalid("worksheet XML has a closing element outside its root")
                })?;
                self.finish(ended)?;
                if ended == Context::Worksheet {
                    if !is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
                        return Err(invalid("worksheet XML has an invalid root closing element"));
                    }
                    *closed_root = true;
                }
            },
            Event::Eof if !*closed_root || !stack.is_empty() => {
                return Err(invalid(
                    "worksheet XML has a missing or unterminated SpreadsheetML worksheet root",
                ));
            },
            Event::Eof => return Ok(true),
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
        }
        Ok(false)
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        if matches!(parent, Context::SheetFormat | Context::Merge) {
            return Err(invalid(
                "worksheet leaf property must not have child elements",
            ));
        }
        if parent == Context::Worksheet && self.seen_sheet_data {
            let local = element.name().local_name();
            if is_spreadsheetml_name(namespace, element.name(), local.as_ref())
                && merge_successor(local.as_ref())
            {
                self.merge_window_closed = true;
            }
        }
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"dimension")
        {
            if self.seen_columns || self.seen_sheet_data {
                return Err(invalid(
                    "worksheet dimension appears after column or cell data",
                ));
            }
            if self.seen_dimension {
                return Err(invalid("worksheet has duplicate dimension elements"));
            }
            self.seen_dimension = true;
            let reference = unqualified_attribute_value(element, b"ref", decoder)?
                .ok_or_else(|| invalid("worksheet dimension is missing ref"))?;
            self.declared_extent = Some(Rect::from_a1(&reference).map_err(|error| {
                invalid(format!(
                    "invalid worksheet dimension '{reference}': {error}"
                ))
            })?);
            return Ok(Context::Other);
        }
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr")
        {
            self.start_defaults(element, decoder, resolver)?;
            return Ok(Context::SheetFormat);
        }
        if parent == Context::Worksheet && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            self.start_columns()?;
            return Ok(Context::Columns);
        }
        if parent == Context::Columns && is_spreadsheetml_name(namespace, element.name(), b"col") {
            self.start_column(element, decoder)?;
            return Ok(Context::Other);
        }
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            if self.seen_sheet_data {
                return Err(invalid("worksheet has duplicate sheetData"));
            }
            self.seen_sheet_data = true;
            return Ok(Context::SheetData);
        }
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"mergeCells")
        {
            if !self.seen_sheet_data {
                return Err(invalid("worksheet mergeCells appears before sheetData"));
            }
            if self.seen_merges {
                return Err(invalid("worksheet has duplicate mergeCells elements"));
            }
            if self.merge_window_closed {
                return Err(invalid(
                    "worksheet mergeCells appears after a schema successor",
                ));
            }
            self.seen_merges = true;
            self.merge_count =
                optional_u32(element, b"count", decoder, "worksheet merged-range count")?
                    .map(usize::try_from)
                    .transpose()
                    .map_err(|_source| {
                        invalid("worksheet merged-range count does not fit usize")
                    })?;
            return Ok(Context::MergeCells);
        }
        if parent == Context::MergeCells
            && is_spreadsheetml_name(namespace, element.name(), b"mergeCell")
        {
            let reference = unqualified_attribute_value(element, b"ref", decoder)?
                .ok_or_else(|| invalid("worksheet mergeCell is missing ref"))?;
            let range = Rect::from_a1(&reference)
                .map_err(|error| invalid(format!("invalid merged range '{reference}': {error}")))?;
            if range.rows() == 1 && range.columns() == 1 {
                return Err(invalid(format!(
                    "worksheet merged range '{reference}' contains only one cell"
                )));
            }
            self.merges
                .try_reserve(1)
                .map_err(|source| allocation("merged ranges", source))?;
            self.merges.push(range);
            return Ok(Context::Merge);
        }
        if parent == Context::MergeCells {
            return Err(invalid("worksheet mergeCells has an unmodeled child"));
        }
        if is_spreadsheetml_name(namespace, element.name(), b"mergeCells")
            || is_spreadsheetml_name(namespace, element.name(), b"mergeCell")
        {
            return Err(invalid(
                "worksheet merge markup appears outside its schema context",
            ));
        }
        if parent == Context::SheetData && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            self.start_row(element, decoder)?;
            return Ok(Context::Row);
        }
        if parent == Context::Row && is_spreadsheetml_name(namespace, element.name(), b"c") {
            self.start_cell(element, decoder)?;
            return Ok(Context::Cell);
        }
        if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"f") {
            self.start_formula(element, decoder)?;
            return Ok(Context::Formula);
        }
        if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"v") {
            self.start_value()?;
            return Ok(Context::Value);
        }
        if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"is") {
            self.start_inline()?;
            return Ok(Context::Inline);
        }
        if parent == Context::Inline && is_spreadsheetml_name(namespace, element.name(), b"t") {
            self.start_inline_text(false)?;
            return Ok(Context::Text(TextTarget::Inline));
        }
        if parent == Context::Inline && is_spreadsheetml_name(namespace, element.name(), b"r") {
            self.start_run()?;
            return Ok(Context::Run);
        }
        if parent == Context::Run && is_spreadsheetml_name(namespace, element.name(), b"t") {
            self.start_inline_text(true)?;
            return Ok(Context::Text(TextTarget::Inline));
        }
        Ok(Context::Other)
    }

    fn start_row(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.row.is_some() {
            return Err(invalid("nested worksheet row"));
        }
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => parse_one_based_row(&value)?,
            None => self
                .previous_row
                .checked_add(1)
                .filter(|value| *value <= ROWS)
                .ok_or_else(|| invalid("inferred worksheet row exceeds the spreadsheet grid"))?,
        };
        self.seen_rows
            .try_reserve(1)
            .map_err(|source| allocation("worksheet parser row index set", source))?;
        if !self.seen_rows.insert(number) {
            return Err(invalid(format!("duplicate worksheet row {number}")));
        }
        if self.previous_row != 0 && number < self.previous_row {
            return Err(invalid(format!(
                "worksheet row {number} appears after row {}",
                self.previous_row
            )));
        }
        self.previous_row = number;
        let height = optional_f64(element, b"ht", decoder, "worksheet row height")?
            .map(row::Height::new)
            .transpose()?;
        let style = optional_u32(element, b"s", decoder, "worksheet row style")?;
        if style.is_some_and(|style| style > MAX_CELL_STYLE) {
            return Err(invalid(format!(
                "worksheet row style exceeds {MAX_CELL_STYLE}"
            )));
        }
        let outline = row::OutlineAt::from(
            optional_u32(
                element,
                b"outlineLevel",
                decoder,
                "worksheet row outline level",
            )?
            .unwrap_or(0),
        )
        .resolve()?;
        let mut flags = row::Flags::empty();
        for (attribute, flag, field) in [
            (b"hidden".as_slice(), row::Flags::HIDDEN, "hidden"),
            (
                b"customHeight".as_slice(),
                row::Flags::CUSTOM_HEIGHT,
                "customHeight",
            ),
            (b"collapsed".as_slice(), row::Flags::COLLAPSED, "collapsed"),
            (b"thickTop".as_slice(), row::Flags::THICK_TOP, "thickTop"),
            (b"thickBot".as_slice(), row::Flags::THICK_BOTTOM, "thickBot"),
            (b"ph".as_slice(), row::Flags::PHONETIC, "ph"),
            (
                b"customFormat".as_slice(),
                row::Flags::CUSTOM_FORMAT,
                "customFormat",
            ),
        ] {
            if optional_bool(element, attribute, decoder, field)?.unwrap_or(false) {
                flags.insert(flag);
            }
        }
        self.row = Some(PendingRow {
            number,
            last_column: 0,
            properties: row::Properties {
                height,
                descent: self.extensions.rows.remove(&number),
                style,
                outline,
                flags,
            },
        });
        Ok(())
    }

    fn start_defaults(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.seen_defaults {
            return Err(invalid("worksheet has duplicate sheetFormatPr elements"));
        }
        if self.seen_columns || self.seen_sheet_data {
            return Err(invalid(
                "worksheet sheetFormatPr appears after column or cell data",
            ));
        }
        self.seen_defaults = true;

        self.defaults = Some(parse_defaults_element(
            element,
            decoder,
            resolver,
            self.extensions.defaults.take(),
        )?);
        Ok(())
    }

    fn start_columns(&mut self) -> Result<()> {
        if self.seen_columns {
            return Err(invalid("worksheet has duplicate cols elements"));
        }
        if self.seen_sheet_data {
            return Err(invalid("worksheet cols appears after sheetData"));
        }
        self.seen_columns = true;
        self.columns = Some(Assignments::new()?);
        Ok(())
    }

    fn start_column(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let min = required_u32(element, b"min", decoder, "worksheet column minimum")?;
        let max = required_u32(element, b"max", decoder, "worksheet column maximum")?;
        if min == 0 || max > COLUMNS || min > max {
            return Err(invalid(format!(
                "invalid worksheet column range '{min}:{max}'"
            )));
        }
        let width = optional_f64(element, b"width", decoder, "worksheet column width")?
            .map(column::Width::new)
            .transpose()?;
        let style = optional_u32(element, b"style", decoder, "worksheet column style")?;
        if style.is_some_and(|style| style > MAX_COLUMN_STYLE) {
            return Err(invalid(format!(
                "worksheet column style exceeds {MAX_COLUMN_STYLE}"
            )));
        }
        let outline_level = optional_u32(
            element,
            b"outlineLevel",
            decoder,
            "worksheet column outline level",
        )?
        .unwrap_or(0);
        let outline = column::OutlineAt::from(outline_level).resolve()?;
        let mut flags = Flags::empty();
        for (attribute, flag, field) in [
            (b"hidden".as_slice(), Flags::HIDDEN, "hidden"),
            (b"bestFit".as_slice(), Flags::BEST_FIT, "bestFit"),
            (
                b"customWidth".as_slice(),
                Flags::CUSTOM_WIDTH,
                "customWidth",
            ),
            (b"phonetic".as_slice(), Flags::PHONETIC, "phonetic"),
            (b"collapsed".as_slice(), Flags::COLLAPSED, "collapsed"),
        ] {
            if optional_bool(element, attribute, decoder, field)?.unwrap_or(false) {
                flags.insert(flag);
            }
        }
        let first = ColumnIndex::new(min - 1)?;
        let last = ColumnIndex::new(max - 1)?;
        self.columns
            .as_mut()
            .ok_or_else(|| invalid("worksheet col appears outside cols"))?
            .assign(
                first,
                last,
                column::Properties {
                    width,
                    style,
                    outline,
                    flags,
                },
            );
        self.column_records = self
            .column_records
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet column record count overflow"))?;
        Ok(())
    }

    fn start_cell(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.cell.is_some() {
            return Err(invalid("nested worksheet cell"));
        }
        let row = self
            .row
            .as_ref()
            .ok_or_else(|| invalid("worksheet cell outside a row"))?
            .number;
        let CellAttributeView {
            reference,
            style,
            cell_metadata,
            value_metadata,
            cell_type,
        } = scan_cell_attributes(element, decoder)?;
        let column = match reference {
            Some(reference) => {
                let (reference_row, column) = parse_a1(&reference)?;
                if reference_row != row {
                    return Err(invalid(format!(
                        "cell reference '{reference}' does not belong to row {row}"
                    )));
                }
                column
            },
            None => self
                .row
                .as_ref()
                .and_then(|row| row.last_column.checked_add(1))
                .filter(|column| *column <= COLUMNS)
                .ok_or_else(|| invalid("inferred worksheet column exceeds the grid"))?,
        };
        let pending_row = self
            .row
            .as_mut()
            .ok_or_else(|| invalid("worksheet cell outside a row"))?;
        pending_row.last_column = column;
        let style = parse_cell_u32(style, decoder, "worksheet cell style")?;
        if style.is_some_and(|style| style > MAX_CELL_STYLE) {
            return Err(invalid(format!(
                "worksheet cell style exceeds {MAX_CELL_STYLE}"
            )));
        }
        let cell_metadata = parse_cell_u32(cell_metadata, decoder, "cell metadata index")?;
        if cell_metadata.is_some_and(|index| !(1..=MAX_METADATA_INDEX).contains(&index)) {
            return Err(invalid("cell metadata index is outside Office limits"));
        }
        let value_metadata = parse_cell_u32(value_metadata, decoder, "value metadata index")?;
        if value_metadata.is_some_and(|index| !(1..=MAX_METADATA_INDEX).contains(&index)) {
            return Err(invalid("value metadata index is outside Office limits"));
        }
        self.cell = Some(PendingCell {
            row,
            column,
            style,
            cell_metadata,
            value_metadata,
            cell_type: cell_type
                .map(|attribute| decode_cell_attribute(attribute, decoder))
                .transpose()?,
            value: String::new(),
            value_bytes: 0,
            saw_value: false,
            formula: String::new(),
            formula_characters: 0,
            formula_kind: None,
            inline: String::new(),
            inline_bytes: 0,
            saw_inline: false,
            saw_inline_simple: false,
            saw_inline_run: false,
            run_has_text: false,
        });
        Ok(())
    }

    fn start_formula(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("worksheet formula outside a cell"))?;
        if cell.formula_kind.is_some() {
            return Err(invalid("duplicate worksheet formula"));
        }
        let formula_type = unqualified_attribute_value(element, b"t", decoder)?
            .unwrap_or_else(|| "normal".to_owned());
        let range = unqualified_attribute_value(element, b"ref", decoder)?;
        if let Some(value) = range.as_deref() {
            FormulaRange::parse(value)?;
        }
        if optional_bool(element, b"bx", decoder, "formula bx")?.unwrap_or(false) {
            return Err(invalid("Office requires formula bx to be false"));
        }
        cell.formula_kind = Some(match formula_type.as_str() {
            "normal" => RawFormulaKind::Scalar,
            "array" => RawFormulaKind::Array(range),
            "dataTable" => RawFormulaKind::DataTable(range),
            "shared" => RawFormulaKind::Shared {
                index: optional_u32(element, b"si", decoder, "shared formula index")?
                    .ok_or_else(|| invalid("shared formula is missing required si"))?,
                range,
            },
            _ => RawFormulaKind::Unknown(formula_type),
        });
        Ok(())
    }

    fn start_value(&mut self) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("worksheet value outside a cell"))?;
        if cell.saw_value {
            return Err(invalid("duplicate worksheet cell value"));
        }
        cell.saw_value = true;
        Ok(())
    }

    fn start_inline(&mut self) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("inline string outside a worksheet cell"))?;
        if cell.saw_inline {
            return Err(invalid("duplicate worksheet inline string"));
        }
        cell.saw_inline = true;
        Ok(())
    }

    fn start_run(&mut self) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("rich-text run outside an inline string"))?;
        if cell.saw_inline_simple {
            return Err(invalid("inline string mixes simple and rich text"));
        }
        cell.saw_inline_run = true;
        cell.run_has_text = false;
        Ok(())
    }

    fn start_inline_text(&mut self, in_run: bool) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("inline text outside a worksheet cell"))?;
        if in_run {
            if cell.run_has_text {
                return Err(invalid("rich-text run has duplicate text"));
            }
            cell.run_has_text = true;
        } else {
            if cell.saw_inline_simple || cell.saw_inline_run {
                return Err(invalid("inline string mixes or duplicates text"));
            }
            cell.saw_inline_simple = true;
        }
        Ok(())
    }

    fn push_text(&mut self, target: TextTarget, value: &str) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("worksheet cell text outside a cell"))?;
        match target {
            TextTarget::Formula => {
                cell.formula_characters = cell
                    .formula_characters
                    .checked_add(value.chars().count())
                    .filter(|length| *length <= MAX_FORMULA_CHARACTERS)
                    .ok_or_else(|| {
                        invalid(format!(
                            "worksheet formula exceeds {MAX_FORMULA_CHARACTERS} characters"
                        ))
                    })?;
                cell.formula
                    .try_reserve(value.len())
                    .map_err(|source| allocation("worksheet formula", source))?;
                cell.formula.push_str(value);
            },
            TextTarget::Value => {
                cell.value_bytes = cell
                    .value_bytes
                    .checked_add(value.len())
                    .filter(|length| *length <= MAX_ENCODED_CELL_BYTES)
                    .ok_or_else(|| invalid("worksheet value text is too large"))?;
                cell.value
                    .try_reserve(value.len())
                    .map_err(|source| allocation("worksheet value", source))?;
                cell.value.push_str(value);
            },
            TextTarget::Inline => {
                cell.inline_bytes = cell
                    .inline_bytes
                    .checked_add(value.len())
                    .filter(|length| *length <= MAX_ENCODED_CELL_BYTES)
                    .ok_or_else(|| invalid("worksheet inline text is too large"))?;
                cell.inline
                    .try_reserve(value.len())
                    .map_err(|source| allocation("worksheet inline text", source))?;
                cell.inline.push_str(value);
            },
        }
        Ok(())
    }

    fn finish(&mut self, context: Context) -> Result<()> {
        match context {
            Context::Formula => {
                // Formula character data is delivered while the Formula
                // context itself is current; no child wrapper is needed.
                Ok(())
            },
            Context::Value => Ok(()),
            Context::Cell => self.finish_cell(),
            Context::Row => self.finish_row(),
            Context::Columns if self.column_records == 0 => {
                Err(invalid("worksheet cols contains no col records"))
            },
            Context::MergeCells => {
                if self.merges.is_empty() {
                    return Err(invalid(
                        "worksheet mergeCells contains no mergeCell records",
                    ));
                }
                if self
                    .merge_count
                    .is_some_and(|count| count != self.merges.len())
                {
                    return Err(invalid(format!(
                        "worksheet merged-range count differs from {} records",
                        self.merges.len()
                    )));
                }
                Ok(())
            },
            Context::Worksheet
            | Context::SheetFormat
            | Context::Columns
            | Context::SheetData
            | Context::Merge
            | Context::Inline
            | Context::Run
            | Context::Text(_)
            | Context::Other => Ok(()),
        }
    }

    fn finish_cell(&mut self) -> Result<()> {
        let mut cell = self
            .cell
            .take()
            .ok_or_else(|| invalid("missing worksheet cell"))?;
        if cell.saw_inline && cell.saw_value {
            return Err(invalid(
                "worksheet cell contains both inline text and a value",
            ));
        }
        if cell.saw_inline && !matches!(cell.cell_type.as_deref(), None | Some("inlineStr")) {
            return Err(invalid("inline string has a non-inline cell type"));
        }
        if cell.cell_type.as_deref() == Some("inlineStr") {
            cell.saw_inline = true;
        }
        if cell.saw_value && cell.value.chars().count() > MAX_CELL_CHARACTERS {
            return Err(invalid(format!(
                "worksheet value exceeds {MAX_CELL_CHARACTERS} characters"
            )));
        }
        if let Some(kind) = cell.formula_kind.as_ref()
            && !matches!(kind, RawFormulaKind::Unknown(_))
            && cell.formula.trim_start().starts_with('=')
        {
            return Err(invalid("worksheet formula must omit the leading '='"));
        }
        if matches!(
            cell.formula_kind.as_ref(),
            Some(RawFormulaKind::Scalar | RawFormulaKind::Array(_))
        ) && cell.formula.is_empty()
        {
            return Err(invalid("worksheet formula expression is empty"));
        }
        if cell.saw_inline {
            cell.inline = decode_spreadsheet_text(&cell.inline)?;
            if cell.inline.chars().count() > MAX_CELL_CHARACTERS {
                return Err(invalid(format!(
                    "inline string exceeds {MAX_CELL_CHARACTERS} characters"
                )));
            }
        }
        let address = Address::at(cell.row - 1, cell.column - 1)?;
        let formula = cell.formula_kind.map(|kind| RawFormula {
            text: cell.formula,
            kind,
        });
        self.cells
            .try_reserve(1)
            .map_err(|source| allocation("sparse worksheet cells", source))?;
        self.cells.push(RawCell {
            address,
            style: cell.style,
            cell_metadata: cell.cell_metadata,
            value_metadata: cell.value_metadata,
            cell_type: cell.cell_type,
            value: cell.saw_value.then_some(cell.value),
            inline: cell.saw_inline.then_some(cell.inline),
            inline_rich: cell.saw_inline_run,
            formula_range: None,
            shared_formula: None,
            formula,
        });
        Ok(())
    }

    fn finish_row(&mut self) -> Result<()> {
        if self.cell.is_some() {
            return Err(invalid("unterminated worksheet cell"));
        }
        let row = self
            .row
            .take()
            .ok_or_else(|| invalid("missing worksheet row"))?;
        self.rows
            .try_reserve(1)
            .map_err(|source| allocation("sparse worksheet rows", source))?;
        self.rows.push(row::Stored::new(
            RowIndex::new(row.number - 1)?,
            row.properties,
        ));
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
enum SourceHandoffKind {
    CellStart,
    CellEmpty,
    CellEnd,
    RowStart,
    RowEmpty,
    RowEnd,
}

fn source_handoff_kind(
    stack: &[Context],
    namespace: &ResolveResult<'_>,
    event: &Event<'_>,
) -> Option<SourceHandoffKind> {
    let parent = stack.last().copied();
    match event {
        Event::Start(element)
            if parent == Some(Context::Row)
                && is_spreadsheetml_name(namespace, element.name(), b"c") =>
        {
            Some(SourceHandoffKind::CellStart)
        },
        Event::Empty(element)
            if parent == Some(Context::Row)
                && is_spreadsheetml_name(namespace, element.name(), b"c") =>
        {
            Some(SourceHandoffKind::CellEmpty)
        },
        Event::End(element)
            if parent == Some(Context::Cell) && element.local_name().as_ref() == b"c" =>
        {
            Some(SourceHandoffKind::CellEnd)
        },
        Event::Start(element)
            if parent == Some(Context::SheetData)
                && is_spreadsheetml_name(namespace, element.name(), b"row") =>
        {
            Some(SourceHandoffKind::RowStart)
        },
        Event::Empty(element)
            if parent == Some(Context::SheetData)
                && is_spreadsheetml_name(namespace, element.name(), b"row") =>
        {
            Some(SourceHandoffKind::RowEmpty)
        },
        Event::End(element)
            if parent == Some(Context::Row) && element.local_name().as_ref() == b"row" =>
        {
            Some(SourceHandoffKind::RowEnd)
        },
        _ => None,
    }
}

fn emit_source_handoff<C: ?Sized>(
    kind: SourceHandoffKind,
    span: SourceEventSpan,
    cells_before: usize,
    rows_before: usize,
    parser: &Parser,
    context: &mut C,
    handoff: &mut impl FnMut(&mut C, SourceHandoff),
) {
    match kind {
        SourceHandoffKind::CellStart => {
            let Some(cell) = parser.cell.as_ref() else {
                return;
            };
            let Ok(address) = Address::at(cell.row - 1, cell.column - 1) else {
                return;
            };
            handoff(context, SourceHandoff::CellStart { address, span });
        },
        SourceHandoffKind::CellEmpty | SourceHandoffKind::CellEnd => {
            let Some(cell) = parser.cells.get(cells_before) else {
                return;
            };
            let kind = match kind {
                SourceHandoffKind::CellEmpty => SourceHandoff::CellEmpty {
                    address: cell.address,
                    span,
                },
                SourceHandoffKind::CellEnd => SourceHandoff::CellEnd {
                    address: cell.address,
                    span,
                },
                _ => return,
            };
            handoff(context, kind);
        },
        SourceHandoffKind::RowStart => {
            let Some(row) = parser.row.as_ref() else {
                return;
            };
            handoff(
                context,
                SourceHandoff::RowStart {
                    number: row.number,
                    span,
                },
            );
        },
        SourceHandoffKind::RowEmpty | SourceHandoffKind::RowEnd => {
            let Some(row) = parser.rows.get(rows_before) else {
                return;
            };
            let number = row.index.get().saturating_add(1);
            let kind = match kind {
                SourceHandoffKind::RowEmpty => SourceHandoff::RowEmpty { number, span },
                SourceHandoffKind::RowEnd => SourceHandoff::RowEnd { number, span },
                _ => return,
            };
            handoff(context, kind);
        },
    }
}

/// Traverse one eligible source with the ordinary parser while giving a
/// caller-owned validator a borrowed view of each event first.
///
/// Early parser failures are deliberately not returned here. The source-backed
/// caller repeats the established two-pass path so the historical error text
/// and precedence remain authoritative. Returning immediately on any
/// provisional failure drops the parser before fallback starts. Once the
/// observer accepts EOF, the finalization result is carried to the raw facade
/// so its historical extension-error retry can still run.
/// Original event positions and post-transition cell/row handoffs are also supplied.
pub(super) fn parse_source_with_observer_and_handoff<'a, F, C, O, H>(
    content: &[u8],
    strings: F,
    context: &mut C,
    mut observer: O,
    mut handoff: H,
) -> super::SourceParseAttempt
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
    C: ?Sized,
    O: for<'event> FnMut(&mut C, &ResolveResult<'event>, &Event<'event>, SourceEventSpan) -> bool,
    H: FnMut(&mut C, SourceHandoff),
{
    let mut reader = NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut parser = Parser::new(x14ac::Values::default());
    let mut stack = Vec::new();
    let mut closed_root = false;
    let mut event_count = 0usize;

    loop {
        let event_start = match usize::try_from(reader.buffer_position()) {
            Ok(position) => position,
            Err(_) => return super::SourceParseAttempt::ReaderFailed,
        };
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(_) => return super::SourceParseAttempt::ReaderFailed,
        };
        let event_end = match usize::try_from(reader.buffer_position()) {
            Ok(position) => position,
            Err(_) => return super::SourceParseAttempt::ReaderFailed,
        };
        let event_limit_exceeded = match event_count.checked_add(1) {
            Some(count) => {
                event_count = count;
                count > MAX_XML_EVENTS || count > super::MAX_SHARED_PROVISIONAL_EVENTS
            },
            None => true,
        };
        let (namespace, event) = reader.resolver().resolve_event(event);
        let decoder = reader.decoder();
        let resolver = reader.resolver();
        let span = SourceEventSpan {
            start: event_start,
            end: event_end,
            decoder,
        };
        let handoff_kind = source_handoff_kind(&stack, &namespace, &event);
        let cells_before = parser.cells.len();
        let rows_before = parser.rows.len();
        let parser_allowed = observer(context, &namespace, &event, span);
        if !parser_allowed || event_limit_exceeded {
            return super::SourceParseAttempt::ProvisionalFailed;
        }

        let complete = match parser.transition(
            &mut stack,
            &mut closed_root,
            &namespace,
            event,
            decoder,
            resolver,
        ) {
            Ok(complete) => complete,
            Err(_) => return super::SourceParseAttempt::ProvisionalFailed,
        };
        if complete {
            break;
        }
        if let Some(kind) = handoff_kind {
            emit_source_handoff(
                kind,
                span,
                cells_before,
                rows_before,
                &parser,
                context,
                &mut handoff,
            );
        }
    }

    super::SourceParseAttempt::Complete(parser.finish_parse(strings))
}
