//! Streaming `SpreadsheetML` event codec.

use std::collections::HashSet;

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use litchi_sheet::{COLUMNS, Cell as Address, Column as ColumnIndex, ROWS, Rect, Row as RowIndex};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::formula::Range as FormulaRange;
use super::super::namespace::is_spreadsheetml_name;
use super::super::strings::decode_spreadsheet_text;
use super::model::{
    Context, MAX_CELL_CHARACTERS, MAX_CELL_STYLE, MAX_COLUMN_STYLE, MAX_ENCODED_CELL_BYTES,
    MAX_FORMULA_CHARACTERS, MAX_METADATA_INDEX, MAX_XML_DEPTH, Parser, PendingCell, PendingRow,
    RawCell, RawFormula, RawFormulaKind, TextTarget, merge_successor,
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
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
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
                        &resolver,
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
                        &resolver,
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
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| invalid(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet")
                    {
                        return Err(invalid(
                            "worksheet XML must have one SpreadsheetML worksheet root",
                        ));
                    }
                    stack.push(Context::Worksheet);
                },
                Event::Empty(element) if stack.is_empty() => {
                    if closed_root
                        || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet")
                    {
                        return Err(invalid(
                            "worksheet XML must have one SpreadsheetML worksheet root",
                        ));
                    }
                    closed_root = true;
                },
                Event::Start(element) => {
                    if stack.len() >= MAX_XML_DEPTH {
                        return Err(invalid(format!(
                            "worksheet XML exceeds {MAX_XML_DEPTH} levels"
                        )));
                    }
                    let parent = current(&stack)?;
                    let child = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                    stack.push(child);
                },
                Event::Empty(element) => {
                    let parent = current(&stack)?;
                    let child = parser.start(parent, &namespace, &element, decoder, &resolver)?;
                    parser.finish(child)?;
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
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(
                            target,
                            &value.decode().map_err(|error| invalid(error.to_string()))?,
                        )?;
                    }
                },
                Event::CData(value) => {
                    if matches!(stack.last(), Some(Context::SheetFormat | Context::Merge)) {
                        return Err(invalid("worksheet leaf property cannot contain CDATA"));
                    }
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(
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
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(target, &decode_xml_reference(&value)?)?;
                    }
                },
                Event::End(element) => {
                    let ended = stack.pop().ok_or_else(|| {
                        invalid("worksheet XML has a closing element outside its root")
                    })?;
                    parser.finish(ended)?;
                    if ended == Context::Worksheet {
                        if !is_spreadsheetml_name(&namespace, element.name(), b"worksheet") {
                            return Err(invalid(
                                "worksheet XML has an invalid root closing element",
                            ));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(invalid(
                        "worksheet XML has a missing or unterminated SpreadsheetML worksheet root",
                    ));
                },
                Event::Eof => break,
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
            }
        }

        resolve_shared_formulas(&mut parser.cells)?;
        let needs_strings = parser
            .cells
            .iter()
            .any(|cell| cell.cell_type.as_deref() == Some("s"));
        let strings = if needs_strings { strings()? } else { None };
        let mut cells = Vec::new();
        cells
            .try_reserve(parser.cells.len())
            .map_err(|source| allocation("sparse worksheet cells", source))?;
        let declared_extent = parser.declared_extent;
        let rows = parser.rows;
        let columns = column::resolve(parser.columns)?;
        let defaults = parser.defaults;
        let merges = parser.merges;
        for cell in parser.cells {
            cells.push(materialize(cell, strings)?);
        }
        Store::from_unsorted(cells, rows, columns, defaults, merges, declared_extent)
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
        let column = match unqualified_attribute_value(element, b"r", decoder)? {
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
        let style = optional_u32(element, b"s", decoder, "worksheet cell style")?;
        if style.is_some_and(|style| style > MAX_CELL_STYLE) {
            return Err(invalid(format!(
                "worksheet cell style exceeds {MAX_CELL_STYLE}"
            )));
        }
        let cell_metadata = optional_u32(element, b"cm", decoder, "cell metadata index")?;
        if cell_metadata.is_some_and(|index| !(1..=MAX_METADATA_INDEX).contains(&index)) {
            return Err(invalid("cell metadata index is outside Office limits"));
        }
        let value_metadata = optional_u32(element, b"vm", decoder, "value metadata index")?;
        if value_metadata.is_some_and(|index| !(1..=MAX_METADATA_INDEX).contains(&index)) {
            return Err(invalid("value metadata index is outside Office limits"));
        }
        self.cell = Some(PendingCell {
            row,
            column,
            style,
            cell_metadata,
            value_metadata,
            cell_type: unqualified_attribute_value(element, b"t", decoder)?,
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
