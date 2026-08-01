//! Namespace-aware streaming parser for sparse worksheet cell data.

pub(crate) mod edit;
mod x14ac;

use std::collections::{HashMap, HashSet};

use litchi_ooxml_common::xml::{decode_xml_reference, unqualified_attribute_value};
use litchi_sheet::{COLUMNS, Cell as Address, Column as ColumnIndex, ROWS, Rect, Row as RowIndex};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::formula::{Range as FormulaRange, translate};
use super::namespace::is_spreadsheetml_name;
use super::strings::decode_spreadsheet_text;
use crate::cell::{Cell, Date, ErrorValue, Number, Store, Stored, Text, Unknown, Value};
use crate::column::{self, Assignments, Flags};
use crate::error::{Result, invalid};
use crate::formula::{Cache, Formula, Kind};
use crate::layout::{self, Defaults};
use crate::row;

const MAX_CELL_CHARACTERS: usize = 32_767;
const MAX_FORMULA_CHARACTERS: usize = 8_192;
// A supplementary Unicode scalar can occupy two seven-byte `_xHHHH_`
// SpreadsheetML escapes before decoding.
const MAX_ENCODED_CELL_BYTES: usize = MAX_CELL_CHARACTERS * 14;
const MAX_CELL_STYLE: u32 = 65_490;
const MAX_COLUMN_STYLE: u32 = 65_429;
const MAX_METADATA_INDEX: u32 = 2_147_483_647;

fn merge_successor(local: &[u8]) -> bool {
    matches!(
        local,
        b"phoneticPr"
            | b"conditionalFormatting"
            | b"dataValidations"
            | b"hyperlinks"
            | b"printOptions"
            | b"pageMargins"
            | b"pageSetup"
            | b"headerFooter"
            | b"rowBreaks"
            | b"colBreaks"
            | b"customProperties"
            | b"cellWatches"
            | b"ignoredErrors"
            | b"smartTags"
            | b"drawing"
            | b"legacyDrawing"
            | b"legacyDrawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Worksheet,
    SheetFormat,
    Columns,
    SheetData,
    MergeCells,
    Merge,
    Row,
    Cell,
    Formula,
    Value,
    Inline,
    Run,
    Text(TextTarget),
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    Formula,
    Value,
    Inline,
}

#[derive(Debug)]
struct PendingRow {
    number: u32,
    last_column: u32,
    properties: row::Properties,
}

#[derive(Debug)]
struct PendingCell {
    row: u32,
    column: u32,
    style: Option<u32>,
    cell_metadata: Option<u32>,
    value_metadata: Option<u32>,
    cell_type: Option<String>,
    value: String,
    value_bytes: usize,
    saw_value: bool,
    formula: String,
    formula_characters: usize,
    formula_kind: Option<RawFormulaKind>,
    inline: String,
    inline_bytes: usize,
    saw_inline: bool,
    saw_inline_simple: bool,
    saw_inline_run: bool,
    run_has_text: bool,
}

#[derive(Debug)]
enum RawFormulaKind {
    Scalar,
    Array(Option<String>),
    DataTable(Option<String>),
    Shared { index: u32, range: Option<String> },
    Unknown(String),
}

#[derive(Debug)]
struct RawCell {
    address: Address,
    style: Option<u32>,
    cell_metadata: Option<u32>,
    value_metadata: Option<u32>,
    cell_type: Option<String>,
    value: Option<String>,
    inline: Option<String>,
    formula: Option<RawFormula>,
}

#[derive(Debug)]
struct RawFormula {
    text: String,
    kind: RawFormulaKind,
}

#[derive(Debug)]
struct SharedMember {
    cell_index: usize,
    row: u32,
    column: u32,
    index: u32,
    range: Option<String>,
    text: String,
}

#[derive(Debug)]
struct SharedMaster {
    row: u32,
    column: u32,
    range: FormulaRange,
    text: String,
}

#[derive(Debug)]
struct Parser {
    cells: Vec<RawCell>,
    rows: Vec<row::Stored>,
    columns: Option<Assignments<column::Properties>>,
    defaults: Option<Defaults>,
    extensions: x14ac::Values,
    declared_extent: Option<Rect>,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    seen_rows: HashSet<u32>,
    previous_row: u32,
    seen_dimension: bool,
    seen_defaults: bool,
    seen_columns: bool,
    column_records: usize,
    seen_sheet_data: bool,
    merges: Vec<Rect>,
    merge_count: Option<usize>,
    seen_merges: bool,
    merge_window_closed: bool,
}

pub(crate) fn parse<'a, F>(content: &[u8], strings: F) -> Result<Store>
where
    F: FnOnce() -> Result<Option<&'a [Text]>>,
{
    let extensions = x14ac::capture(content)?;
    let processed = litchi_ooxml_common::mce::process_ooxml(content)?;
    let content = std::str::from_utf8(processed.as_ref())
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    Parser::parse(content, strings, extensions)
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

    fn parse<'a, F>(content: &str, strings: F, extensions: x14ac::Values) -> Result<Store>
    where
        F: FnOnce() -> Result<Option<&'a [Text]>>,
    {
        let mut reader = NsReader::from_reader(content.as_bytes());
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
                    let parent = current(&stack)?;
                    let child = parser.start(parent, &namespace, &element, decoder)?;
                    stack.push(child);
                },
                Event::Empty(element) => {
                    let parent = current(&stack)?;
                    let child = parser.start(parent, &namespace, &element, decoder)?;
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
                _ => {},
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
            .map_err(|error| invalid(format!("cannot reserve sparse worksheet cells: {error}")))?;
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
            self.start_defaults(element, decoder)?;
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
                    .map_err(|_| invalid("worksheet merged-range count does not fit usize"))?;
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
                .map_err(|error| invalid(format!("cannot grow merged ranges: {error}")))?;
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

    fn start_defaults(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.seen_defaults {
            return Err(invalid("worksheet has duplicate sheetFormatPr elements"));
        }
        if self.seen_columns || self.seen_sheet_data {
            return Err(invalid(
                "worksheet sheetFormatPr appears after column or cell data",
            ));
        }
        self.seen_defaults = true;

        let base_width = optional_u32(
            element,
            b"baseColWidth",
            decoder,
            "worksheet base column width",
        )?
        .map(|value| {
            u8::try_from(value)
                .map_err(|_| invalid("worksheet base column width exceeds Office maximum 255"))
        })
        .transpose()?;
        let width = optional_f64(
            element,
            b"defaultColWidth",
            decoder,
            "worksheet default column width",
        )?
        .map(layout::Width::new)
        .transpose()?;
        let height = optional_f64(
            element,
            b"defaultRowHeight",
            decoder,
            "worksheet default row height",
        )?
        .ok_or_else(|| invalid("worksheet sheetFormatPr is missing defaultRowHeight"))
        .and_then(|value| layout::Height::new(value).map_err(Into::into))?;
        let row_outline = optional_u32(
            element,
            b"outlineLevelRow",
            decoder,
            "worksheet row outline summary",
        )?
        .map(row::OutlineAt::from)
        .map(row::OutlineAt::resolve)
        .transpose()?;
        let column_outline = optional_u32(
            element,
            b"outlineLevelCol",
            decoder,
            "worksheet column outline summary",
        )?
        .map(row::OutlineAt::from)
        .map(row::OutlineAt::resolve)
        .transpose()?;
        let mut flags = layout::Flags::empty();
        let mut present = layout::Flags::empty();
        for (attribute, flag, field) in [
            (
                b"customHeight".as_slice(),
                layout::Flags::CUSTOM_HEIGHT,
                "customHeight",
            ),
            (
                b"zeroHeight".as_slice(),
                layout::Flags::HIDDEN,
                "zeroHeight",
            ),
            (b"thickTop".as_slice(), layout::Flags::THICK_TOP, "thickTop"),
            (
                b"thickBottom".as_slice(),
                layout::Flags::THICK_BOTTOM,
                "thickBottom",
            ),
        ] {
            if let Some(value) = optional_bool(element, attribute, decoder, field)? {
                present.insert(flag);
                flags.set(flag, value);
            }
        }
        self.defaults = Some(Defaults {
            base_width,
            width,
            height,
            descent: self.extensions.defaults.take(),
            row_outline,
            column_outline,
            flags,
            present,
        });
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
                    .map_err(|error| invalid(format!("cannot grow worksheet formula: {error}")))?;
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
                    .map_err(|error| invalid(format!("cannot grow worksheet value: {error}")))?;
                cell.value.push_str(value);
            },
            TextTarget::Inline => {
                cell.inline_bytes = cell
                    .inline_bytes
                    .checked_add(value.len())
                    .filter(|length| *length <= MAX_ENCODED_CELL_BYTES)
                    .ok_or_else(|| invalid("worksheet inline text is too large"))?;
                cell.inline.try_reserve(value.len()).map_err(|error| {
                    invalid(format!("cannot grow worksheet inline text: {error}"))
                })?;
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
            _ => Ok(()),
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
            .map_err(|error| invalid(format!("cannot grow sparse worksheet cells: {error}")))?;
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
            .map_err(|error| invalid(format!("cannot grow sparse worksheet rows: {error}")))?;
        self.rows.push(row::Stored::new(
            RowIndex::new(row.number - 1)?,
            row.properties,
        ));
        Ok(())
    }
}

fn materialize(raw: RawCell, strings: Option<&[Text]>) -> Result<Stored> {
    let unknown_cell_type = raw
        .cell_type
        .as_deref()
        .filter(|kind| !matches!(*kind, "b" | "d" | "e" | "inlineStr" | "n" | "s" | "str"));
    let cell = if let Some(kind) = unknown_cell_type {
        let formula = raw.formula.map(|formula| formula.text);
        Cell::Unknown(Unknown::new(kind, raw.value, formula))
    } else if let Some(inline) = raw.inline {
        if raw.formula.is_some() {
            return Err(invalid("formula cell cannot contain an inline string"));
        }
        Cell::Value(Value::Text(inline.into()))
    } else if let Some(formula) = raw.formula {
        let RawFormula { text, kind } = formula;
        let kind = match kind {
            RawFormulaKind::Scalar => Kind::Scalar,
            RawFormulaKind::Array(range) => Kind::Array {
                range: range.map(Text::from),
            },
            RawFormulaKind::DataTable(range) => Kind::DataTable {
                range: range.map(Text::from),
            },
            RawFormulaKind::Shared { .. } => {
                return Err(invalid("unresolved shared formula storage record"));
            },
            RawFormulaKind::Unknown(value) => Kind::Unknown(value.into()),
        };
        let cached = raw
            .value
            .as_deref()
            .map(|value| parse_value(raw.cell_type.as_deref(), value, strings))
            .transpose()?
            .flatten()
            .map(Cache::stored);
        Cell::Formula(Formula::parsed(text, kind, cached))
    } else if let Some(value) = raw.value.as_deref() {
        parse_value(raw.cell_type.as_deref(), value, strings)?.map_or(Cell::Empty, Cell::Value)
    } else if let Some(kind) = raw.cell_type.as_deref()
        && !matches!(kind, "n" | "inlineStr")
    {
        if kind == "str" {
            Cell::Value(Value::Text(Text::from("")))
        } else {
            Cell::Empty
        }
    } else {
        Cell::Empty
    };
    Ok(Stored {
        address: raw.address,
        cell,
        style: raw.style,
        cell_metadata: raw.cell_metadata,
        value_metadata: raw.value_metadata,
    })
}

fn parse_value(
    cell_type: Option<&str>,
    value: &str,
    strings: Option<&[Text]>,
) -> Result<Option<Value>> {
    match cell_type {
        None | Some("n") if value.trim().is_empty() => Ok(None),
        None | Some("n") => Number::new(value).map(Value::Number).map(Some),
        Some("str") => {
            let value = decode_spreadsheet_text(value)?;
            if value.chars().count() > MAX_CELL_CHARACTERS {
                return Err(invalid(format!(
                    "worksheet string exceeds {MAX_CELL_CHARACTERS} characters"
                )));
            }
            Ok(Some(Value::Text(value.into())))
        },
        Some("d") => Date::new(value).map(Value::Date).map(Some),
        Some("s") => {
            let index = value
                .trim()
                .parse::<usize>()
                .map_err(|_| invalid(format!("invalid shared-string index '{value}'")))?;
            let strings = strings.ok_or_else(|| {
                invalid("worksheet uses shared strings but the workbook has no shared-string part")
            })?;
            let text = strings.get(index).ok_or_else(|| {
                invalid(format!(
                    "shared-string index {index} exceeds table length {}",
                    strings.len()
                ))
            })?;
            Ok(Some(Value::Text(text.clone())))
        },
        Some("b") => match value.trim() {
            "1" | "true" => Ok(Some(Value::Bool(true))),
            "0" | "false" => Ok(Some(Value::Bool(false))),
            other => Err(invalid(format!("invalid worksheet boolean '{other}'"))),
        },
        Some("e") => Ok(Some(Value::Error(ErrorValue::parse(value)))),
        Some("inlineStr") => Err(invalid(
            "inline-string cell stores text in an is element, not v",
        )),
        Some(other) => Err(invalid(format!(
            "unsupported worksheet cell type '{other}'"
        ))),
    }
}

fn resolve_shared_formulas(cells: &mut [RawCell]) -> Result<()> {
    let mut members = Vec::new();
    for (cell_index, cell) in cells.iter().enumerate() {
        let Some(RawFormula {
            text,
            kind: RawFormulaKind::Shared { index, range },
        }) = cell.formula.as_ref()
        else {
            continue;
        };
        members.push(SharedMember {
            cell_index,
            row: cell.address.row().get() + 1,
            column: cell.address.column().get() + 1,
            index: *index,
            range: range.clone(),
            text: text.clone(),
        });
    }
    if members.is_empty() {
        return Ok(());
    }

    let mut masters = HashMap::<u32, SharedMaster>::new();
    for member in &members {
        if member.range.is_none() && member.text.is_empty() {
            continue;
        }
        let range_text = member.range.as_deref().ok_or_else(|| {
            invalid(format!(
                "shared formula master at ({}, {}) is missing ref",
                member.row, member.column
            ))
        })?;
        if member.text.is_empty() {
            return Err(invalid(format!(
                "shared formula master at ({}, {}) has no expression",
                member.row, member.column
            )));
        }
        let range = FormulaRange::parse(range_text)?;
        if (member.row, member.column) != (range.first_row, range.first_column) {
            return Err(invalid(format!(
                "shared formula master at ({}, {}) is not first in '{range_text}'",
                member.row, member.column
            )));
        }
        if masters
            .insert(
                member.index,
                SharedMaster {
                    row: member.row,
                    column: member.column,
                    range,
                    text: member.text.clone(),
                },
            )
            .is_some()
        {
            return Err(invalid(format!(
                "duplicate shared formula master for si={}",
                member.index
            )));
        }
    }

    for member in members {
        let master = masters.get(&member.index).ok_or_else(|| {
            invalid(format!(
                "shared formula at ({}, {}) has no master for si={}",
                member.row, member.column, member.index
            ))
        })?;
        if !master.range.contains(member.row, member.column) {
            return Err(invalid(format!(
                "shared formula at ({}, {}) lies outside its master range",
                member.row, member.column
            )));
        }
        let is_master = (member.row, member.column) == (master.row, master.column);
        if !is_master && (!member.text.is_empty() || member.range.is_some()) {
            return Err(invalid(format!(
                "shared formula follower at ({}, {}) contains master data",
                member.row, member.column
            )));
        }
        let text = if is_master {
            master.text.clone()
        } else {
            translate(
                &master.text,
                master.row,
                master.column,
                member.row,
                member.column,
            )
        };
        let formula = cells
            .get_mut(member.cell_index)
            .and_then(|cell| cell.formula.as_mut())
            .ok_or_else(|| invalid("shared formula membership lost its cell"))?;
        formula.text = text;
        formula.kind = RawFormulaKind::Scalar;
    }
    Ok(())
}

fn parse_one_based_row(value: &str) -> Result<u32> {
    let row = value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid worksheet row '{value}'")))?;
    if !(1..=ROWS).contains(&row) {
        return Err(invalid(format!("worksheet row {row} exceeds the grid")));
    }
    Ok(row)
}

fn parse_a1(value: &str) -> Result<(u32, u32)> {
    let bytes = value.as_bytes();
    let split = bytes
        .iter()
        .position(u8::is_ascii_digit)
        .ok_or_else(|| invalid(format!("invalid cell reference '{value}'")))?;
    if split == 0 || split == bytes.len() {
        return Err(invalid(format!("invalid cell reference '{value}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[..split] {
        if !byte.is_ascii_alphabetic() {
            return Err(invalid(format!("invalid cell reference '{value}'")));
        }
        column = column
            .checked_mul(26)
            .and_then(|column| column.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| invalid(format!("cell reference '{value}' overflows")))?;
    }
    if column == 0 || column > COLUMNS {
        return Err(invalid(format!(
            "cell reference '{value}' exceeds the column grid"
        )));
    }
    let row = std::str::from_utf8(&bytes[split..])
        .ok()
        .and_then(|row| row.parse::<u32>().ok())
        .filter(|row| (1..=ROWS).contains(row))
        .ok_or_else(|| invalid(format!("cell reference '{value}' exceeds the row grid")))?;
    Ok((row, column))
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
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
        .ok_or_else(|| invalid(format!("missing {description}")))
}

fn optional_f64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<f64>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

fn optional_bool(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!("invalid {description} '{value}'"))),
        })
        .transpose()
}

fn current(stack: &[Context]) -> Result<Context> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("worksheet XML is missing its root context"))
}

fn text_target(stack: &[Context]) -> Option<TextTarget> {
    match stack.last() {
        Some(Context::Formula) => Some(TextTarget::Formula),
        Some(Context::Value) => Some(TextTarget::Value),
        Some(Context::Text(target)) => Some(*target),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn parses_exact_sparse_values_formulas_and_explicit_empty_cells() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><sheetData>
                <row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" s="3"/></row>
                <row r="3"><c r="A3"><v>-0.000</v></c><c r="B3" t="b"><v>1</v></c>
                <c r="C3" t="inlineStr"><is><r><t>Hello </t></r><r><t>world</t></r></is></c></row>
                <row r="4"><c r="A4" t="str"><f>CONCAT(A1,"!")</f><v>cached</v></c></row>
            </sheetData></worksheet>"#
        );
        let strings = [Text::from("shared")];
        let store = parse(xml.as_bytes(), || Ok(Some(&strings))).expect("valid worksheet");
        assert!(matches!(
            store.get(Address::at(0, 0).expect("address")),
            Some(Cell::Value(Value::Text(value))) if value.as_str() == "shared"
        ));
        assert!(store.get(Address::at(0, 1).expect("address")).is_none());
        assert!(matches!(
            store.get(Address::at(0, 2).expect("address")),
            Some(Cell::Empty)
        ));
        assert!(matches!(
            store.get(Address::at(2, 0).expect("address")),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "-0.000"
        ));
        assert!(matches!(
            store.get(Address::at(2, 2).expect("address")),
            Some(Cell::Value(Value::Text(value))) if value.as_str() == "Hello world"
        ));
        let Some(Cell::Formula(formula)) = store.get(Address::at(3, 0).expect("address")) else {
            panic!("expected formula")
        };
        assert_eq!(formula.text(), "CONCAT(A1,\"!\")");
        assert!(matches!(
            formula.cached().map(Cache::value),
            Some(Value::Text(value)) if value.as_str() == "cached"
        ));
    }

    #[test]
    fn parses_sparse_merges_and_rejects_ambiguous_merge_markup() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:C3"/><mergeCell ref="E5:F5"/></mergeCells></worksheet>"#
        );
        let store = parse(xml.as_bytes(), || Ok(None)).expect("valid merged ranges");
        assert_eq!(store.entries().len(), 2, "merges must stay sparse");
        assert_eq!(
            store.merges().map(Rect::a1).collect::<Vec<_>>(),
            ["A1:C3", "E5:F5"]
        );
        assert!(matches!(
            store.view(Address::from_a1("A1").expect("anchor")),
            crate::cell::View::Stored(Cell::Value(_))
        ));
        assert!(matches!(
            store.view(Address::from_a1("B2").expect("covered")),
            crate::cell::View::Covered(range) if range == Rect::from_a1("A1:C3").expect("range")
        ));
        assert!(matches!(
            store.view(Address::from_a1("D4").expect("missing")),
            crate::cell::View::Missing
        ));

        for malformed in [
            format!(
                r#"<worksheet xmlns="{S}"><sheetData/><mergeCells count="2"><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData/><mergeCells><mergeCell ref="A1:C3"/><mergeCell ref="C3:D4"/></mergeCells></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData/><mergeCells><mergeCell ref="A1"/></mergeCells></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData/><mergeCells><future/></mergeCells></worksheet>"#
            ),
            format!(r#"<worksheet xmlns="{S}"><sheetData/><mergeCell ref="A1:B2"/></worksheet>"#),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData/><hyperlinks/><mergeCells><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
            ),
        ] {
            assert!(parse(malformed.as_bytes(), || Ok(None)).is_err());
        }
    }

    #[test]
    fn expands_shared_formulas_and_preserves_cached_values() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><sheetData>
                <row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="7">B1+$C$1</f><v>1</v></c></row>
                <row r="2"><c r="A2"><f t="shared" si="7"/><v>2</v></c></row>
            </sheetData></worksheet>"#
        );
        let store = parse(xml.as_bytes(), || Ok(None)).expect("valid shared formula");
        let Some(Cell::Formula(formula)) = store.get(Address::at(1, 0).expect("address")) else {
            panic!("expected formula")
        };
        assert_eq!(formula.text(), "B2+$C$1");
        assert!(matches!(
            formula.cached().map(Cache::value),
            Some(Value::Number(number)) if number.as_str() == "2"
        ));
    }

    #[test]
    fn keeps_declared_stored_content_and_style_extents_distinct() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><dimension ref="$B$2:F9"/><sheetData><row r="2"><c r="B2"><v>1</v></c></row><row r="4" ht="30.5" s="1" customFormat="1" customHeight="true" hidden="true" outlineLevel="2" collapsed="1" thickTop="1" thickBot="true" ph="1"><c r="D4"/></row><row r="9" hidden="0"><c r="F9" s="1"/></row></sheetData></worksheet>"#
        );
        let store = parse(xml.as_bytes(), || Ok(None)).expect("valid extents");
        let extents = store.extents();
        assert_eq!(extents.declared().map(Rect::a1).as_deref(), Some("B2:F9"));
        assert_eq!(extents.stored().map(Rect::a1).as_deref(), Some("B2:F9"));
        assert_eq!(extents.content().map(Rect::a1).as_deref(), Some("B2"));
        assert_eq!(extents.styled().map(Rect::a1).as_deref(), Some("F9"));
        assert_eq!(extents.used().map(Rect::a1).as_deref(), Some("B2:F9"));
        let row = store.row(RowIndex::new(3).expect("row 4"));
        assert!(row.hidden());
        assert_eq!(row.height().map(row::Height::get), Some(30.5));
        assert!(row.custom_height());
        assert_eq!(row.outline().get(), 2);
        assert!(row.collapsed());
        assert!(row.thick_top());
        assert!(row.thick_bottom());
        assert!(row.phonetic());
        assert!(row.custom_format());
        assert_eq!(
            store.row_entry(row.index()).unwrap().properties.style,
            Some(1)
        );
        assert!(!store.row(RowIndex::new(8).expect("row 9")).hidden());
        let implicit = store.row(RowIndex::new(5).expect("row 6"));
        assert!(!implicit.stored());
        assert!(!implicit.hidden());
        assert_eq!(store.rows().count(), 3);
    }

    #[test]
    fn parses_checked_grid_defaults_and_effective_x14ac_descent() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:future="urn:future"
                xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                compat:Ignorable="ac future">
                <x:sheetFormatPr baseColWidth="10" defaultColWidth="12.5"
                    defaultRowHeight="16" customHeight="false" zeroHeight="true"
                    thickTop="1" thickBottom="true" outlineLevelRow="3"
                    outlineLevelCol="2" ac:dyDescent="0.2" future:keep="yes"/>
                <x:sheetData><x:row r="2" customHeight="0" ac:dyDescent="0.3"/></x:sheetData>
            </x:worksheet>"#
        );
        let store = parse(xml.as_bytes(), || Ok(None)).expect("valid worksheet defaults");
        let defaults = store.defaults().expect("stored defaults");
        assert_eq!(defaults.base_width(), 10);
        assert_eq!(defaults.stored_base_width(), Some(10));
        assert_eq!(defaults.width().map(layout::Width::get), Some(12.5));
        assert_eq!(defaults.height().get(), 16.0);
        assert!(defaults.custom_height());
        assert!(defaults.hidden());
        assert!(defaults.thick_top());
        assert!(defaults.thick_bottom());
        assert_eq!(defaults.row_outline().get(), 3);
        assert_eq!(defaults.column_outline().get(), 2);
        assert_eq!(defaults.descent().map(layout::Descent::get), Some(0.2));

        let row = store.row(RowIndex::new(1).expect("row 2"));
        assert_eq!(row.descent().map(layout::Descent::get), Some(0.3));
        assert!(row.custom_height());
    }

    #[test]
    fn rejects_malformed_grid_defaults_and_descent() {
        for body in [
            r#"<sheetFormatPr/>"#,
            r#"<sheetFormatPr defaultRowHeight="-1"/>"#,
            r#"<sheetFormatPr defaultRowHeight="NaN"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" baseColWidth="256"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" defaultColWidth="65536"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" outlineLevelRow="8"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" outlineLevelCol="8"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15">text</sheetFormatPr>"#,
            r#"<sheetFormatPr defaultRowHeight="15"><future/></sheetFormatPr>"#,
            r#"<sheetFormatPr defaultRowHeight="15"/><sheetFormatPr defaultRowHeight="16"/>"#,
            r#"<sheetData/><sheetFormatPr defaultRowHeight="15"/>"#,
        ] {
            let xml = format!(r#"<worksheet xmlns="{S}">{body}<sheetData/></worksheet>"#);
            assert!(
                parse(xml.as_bytes(), || Ok(None)).is_err(),
                "accepted {body}"
            );
        }

        for value in ["-0.1", "NaN", "inf"] {
            let xml = format!(
                r#"<worksheet xmlns="{S}"
                    xmlns:a="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                    mc:Ignorable="a"><sheetFormatPr defaultRowHeight="15" a:dyDescent="{value}"/><sheetData/></worksheet>"#
            );
            assert!(
                parse(xml.as_bytes(), || Ok(None)).is_err(),
                "accepted dyDescent={value}"
            );
        }
    }

    #[test]
    fn rejects_extension_xml_beyond_the_depth_limit() {
        let mut xml = format!(r#"<worksheet xmlns="{S}">"#);
        for _ in 0..256 {
            xml.push_str("<future>");
        }
        for _ in 0..256 {
            xml.push_str("</future>");
        }
        xml.push_str("</worksheet>");

        assert!(parse(xml.as_bytes(), || Ok(None)).is_err());
    }

    #[test]
    fn resolves_complete_last_matching_column_records() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><cols><col min="2" max="4" width="20" style="1" hidden="1" bestFit="true" customWidth="1" phonetic="1" outlineLevel="2" collapsed="1"/><col min="3" max="3" width="10"/></cols><sheetData/></worksheet>"#
        );
        let store = parse(xml.as_bytes(), || Ok(None)).expect("valid columns");

        let a = store.column(ColumnIndex::new(0).expect("A"));
        assert!(!a.stored());
        assert!(!a.hidden());
        let b = store.column(ColumnIndex::new(1).expect("B"));
        assert!(b.stored());
        assert!(b.hidden());
        assert_eq!(b.width().map(column::Width::get), Some(20.0));
        assert!(b.best_fit());
        assert!(b.custom_width());
        assert!(b.phonetic());
        assert_eq!(b.outline().get(), 2);
        assert!(b.collapsed());
        assert_eq!(
            store.column_entry(b.index()).unwrap().properties.style,
            Some(1)
        );

        let c = store.column(ColumnIndex::new(2).expect("C"));
        assert!(c.stored());
        assert!(!c.hidden());
        assert_eq!(c.width().map(column::Width::get), Some(10.0));
        assert!(!c.best_fit());
        assert!(!c.custom_width());
        assert!(!c.phonetic());
        assert_eq!(c.outline(), column::Outline::NONE);
        assert!(!c.collapsed());
        assert_eq!(
            store.column_entry(c.index()).unwrap().properties.style,
            None
        );

        let d = store.column(ColumnIndex::new(3).expect("D"));
        assert!(d.hidden());
        assert_eq!(
            store
                .columns()
                .map(|column| column.index())
                .collect::<Vec<_>>(),
            [
                ColumnIndex::new(1).expect("B"),
                ColumnIndex::new(2).expect("C"),
                ColumnIndex::new(3).expect("D"),
            ]
        );
    }

    #[test]
    fn rejects_malformed_column_property_records() {
        for body in [
            "<cols/>",
            "<cols></cols>",
            r#"<cols><col max="1"/></cols>"#,
            r#"<cols><col min="1"/></cols>"#,
            r#"<cols><col min="0" max="1"/></cols>"#,
            r#"<cols><col min="2" max="1"/></cols>"#,
            r#"<cols><col min="1" max="16385"/></cols>"#,
            r#"<cols><col min="1" max="1" width="256"/></cols>"#,
            r#"<cols><col min="1" max="1" width="NaN"/></cols>"#,
            r#"<cols><col min="1" max="1" style="65430"/></cols>"#,
            r#"<cols><col min="1" max="1" outlineLevel="8"/></cols>"#,
            r#"<cols><col min="1" max="1" hidden="yes"/></cols>"#,
            r#"<cols><col min="1" max="1"/></cols><cols><col min="2" max="2"/></cols>"#,
            r#"<sheetData/><cols><col min="1" max="1"/></cols>"#,
        ] {
            let xml = format!(r#"<worksheet xmlns="{S}">{body}<sheetData/></worksheet>"#);
            assert!(
                parse(xml.as_bytes(), || Ok(None)).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn rejects_malformed_dimensions_and_row_properties() {
        for body in [
            "<dimension/><sheetData/>",
            r#"<dimension ref="A0"/><sheetData/>"#,
            r#"<dimension ref="A1"/><dimension ref="B2"/><sheetData/>"#,
            r#"<sheetData/><dimension ref="A1"/>"#,
            r#"<sheetData><row r="1" hidden="yes"/></sheetData>"#,
            r#"<sheetData><row r="1" ht="NaN"/></sheetData>"#,
            r#"<sheetData><row r="1" ht="409.1"/></sheetData>"#,
            r#"<sheetData><row r="1" s="65491"/></sheetData>"#,
            r#"<sheetData><row r="1" outlineLevel="8"/></sheetData>"#,
            r#"<sheetData><row r="1" thickTop="yes"/></sheetData>"#,
        ] {
            let xml = format!(r#"<worksheet xmlns="{S}">{body}</worksheet>"#);
            assert!(
                parse(xml.as_bytes(), || Ok(None)).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn rejects_grid_escape_duplicates_and_broken_shared_groups() {
        for body in [
            r#"<row r="1048577"/>"#,
            r#"<row r="1"><c r="XFE1"/></row>"#,
            r#"<row r="1"><c r="A1"/><c r="A1"/></row>"#,
            r#"<row r="2"/><row r="1"/>"#,
            r#"<row r="1"><c r="A1" s="65491"/></row>"#,
            r#"<row r="1"><c r="A1" cm="0"/></row>"#,
            r#"<row r="1"><c r="A1" vm="2147483648"/></row>"#,
            r#"<row r="1"><c r="A1"><f bx="1">1</f></c></row>"#,
            r#"<row r="1"><c r="A1"><f t="shared" si="0"/></c></row>"#,
        ] {
            let xml =
                format!(r#"<worksheet xmlns="{S}"><sheetData>{body}</sheetData></worksheet>"#);
            assert!(
                parse(xml.as_bytes(), || Ok(None)).is_err(),
                "accepted {body}"
            );
        }
    }

    #[test]
    fn rejects_missing_shared_strings_bad_indexes_and_formula_markers() {
        let shared = format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData></worksheet>"#
        );
        assert!(parse(shared.as_bytes(), || Ok(None)).is_err());
        let strings = [Text::from("only index zero")];
        assert!(parse(shared.as_bytes(), || Ok(Some(&strings))).is_err());

        let marked = format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f>=1+1</f><v>2</v></c></row></sheetData></worksheet>"#
        );
        assert!(parse(marked.as_bytes(), || Ok(None)).is_err());
    }

    #[test]
    fn validates_typed_date_cells_without_normalizing_the_lexeme() {
        let valid = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="d"><v>2026-07-31T12:34:56.250-07:00</v></c></row></sheetData></worksheet>"#;
        let store = parse(valid, || Ok(None)).expect("valid date cell");
        assert!(matches!(
            store.get(Address::from_a1("A1").expect("address")),
            Some(Cell::Value(Value::Date(date)))
                if date.as_str() == "2026-07-31T12:34:56.250-07:00"
        ));

        let invalid = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="d"><v>2026-02-29</v></c></row></sheetData></worksheet>"#;
        assert!(parse(invalid, || Ok(None)).is_err());
    }

    #[test]
    fn numeric_sheets_do_not_load_an_unneeded_shared_string_table() {
        let xml = format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"#
        );
        let called = std::cell::Cell::new(false);
        let store = parse(xml.as_bytes(), || {
            called.set(true);
            Ok(None)
        })
        .expect("numeric worksheet");
        assert!(!called.get());
        assert!(matches!(
            store.get(Address::at(0, 0).expect("address")),
            Some(Cell::Value(Value::Number(number))) if number.as_str() == "7"
        ));
    }
}
