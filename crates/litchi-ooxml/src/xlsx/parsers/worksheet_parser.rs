//! Namespace-aware streaming parser for worksheet cell data.

use std::collections::{HashMap, HashSet};

use litchi_core::sheet::{CellValue, Result as SheetResult};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::common::xml::{decode_xml_reference, unqualified_attribute_value};
use crate::error::{OoxmlError, Result};
use crate::xlsx::RichTextRun;
use crate::xlsx::cell::Cell;
use crate::xlsx::namespace::{
    SPREADSHEETML_NAMESPACE, is_spreadsheetml_name, relationship_attribute_value,
};
use crate::xlsx::shared_strings::decode_spreadsheet_text;
use crate::xlsx::worksheet::{ColumnInfo, RowInfo};

const MAX_EXCEL_COLUMN: u32 = 16_384;

#[derive(Default)]
pub(crate) struct ParsedWorksheetData {
    pub(crate) cells: HashMap<u32, HashMap<u32, CellValue>>,
    pub(crate) cell_styles: HashMap<u32, HashMap<u32, u32>>,
    pub(crate) rows: HashMap<u32, RowInfo>,
    pub(crate) rich_text_cells: HashMap<(u32, u32), Vec<RichTextRun>>,
    pub(crate) merged_regions: Vec<(u32, u32, u32, u32)>,
    pub(crate) columns: HashMap<u32, ColumnInfo>,
    pub(crate) hyperlinks: Vec<ParsedHyperlink>,
    pub(crate) dimensions: Option<(u32, u32, u32, u32)>,
}

pub(crate) struct ParsedHyperlink {
    pub(crate) cell_ref: String,
    pub(crate) relationship_id: Option<String>,
    pub(crate) location: Option<String>,
    pub(crate) display: Option<String>,
    pub(crate) tooltip: Option<String>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Worksheet,
    Columns,
    Hyperlinks,
    MergeCells,
    SheetData,
    Row,
    Cell,
    Formula,
    Value,
    InlineString,
    RichRun,
    RunProperties,
    Text(TextTarget),
    Other,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum TextTarget {
    Formula,
    Value,
    InlineSimple,
    InlineRun,
}

struct PendingRow {
    number: u32,
    info: Option<RowInfo>,
    last_column: u32,
}

struct PendingCell {
    row: u32,
    column: u32,
    style: Option<u32>,
    cell_type: Option<String>,
    value: String,
    saw_value: bool,
    formula: String,
    saw_formula: bool,
    materialized_formula: bool,
    is_array_formula: bool,
    array_range: Option<String>,
    saw_inline_string: bool,
    inline_text: String,
    inline_runs: Vec<RichTextRun>,
    saw_inline_simple: bool,
    saw_inline_run: bool,
}

struct PendingRun {
    value: RichTextRun,
    saw_text: bool,
    saw_properties: bool,
    properties: u8,
}

impl PendingRun {
    fn new() -> Self {
        Self {
            value: RichTextRun {
                text: String::new(),
                font_name: None,
                font_size: None,
                bold: false,
                italic: false,
                underline: false,
                color: None,
            },
            saw_text: false,
            saw_properties: false,
            properties: 0,
        }
    }
}

struct Parser {
    data: ParsedWorksheetData,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    run: Option<PendingRun>,
    previous_row: u32,
    rows: HashSet<u32>,
    merged_regions: HashSet<(u32, u32, u32, u32)>,
    hyperlink_refs: HashSet<String>,
    seen_sheet_data: bool,
    seen_columns: bool,
    seen_hyperlinks: bool,
    seen_merge_cells: bool,
    expected_merge_count: Option<usize>,
    min_row: u32,
    min_column: u32,
    max_row: u32,
    max_column: u32,
}

impl Parser {
    fn new() -> Self {
        Self {
            data: ParsedWorksheetData::default(),
            row: None,
            cell: None,
            run: None,
            previous_row: 0,
            rows: HashSet::new(),
            merged_regions: HashSet::new(),
            hyperlink_refs: HashSet::new(),
            seen_sheet_data: false,
            seen_columns: false,
            seen_hyperlinks: false,
            seen_merge_cells: false,
            expected_merge_count: None,
            min_row: u32::MAX,
            min_column: u32::MAX,
            max_row: 0,
            max_column: 0,
        }
    }

    fn parse(content: &str) -> Result<ParsedWorksheetData> {
        let mut reader = NsReader::from_reader(content.as_bytes());
        let mut parser = Self::new();
        let mut stack = Vec::new();
        let mut closed_root = false;
        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    if stack.is_empty() {
                        if closed_root
                            || !is_spreadsheetml_name(&namespace, element.name(), b"worksheet")
                        {
                            return Err(invalid(
                                "worksheet XML must have one SpreadsheetML worksheet root",
                            ));
                        }
                        stack.push(Context::Worksheet);
                        continue;
                    }
                    let parent = context(&stack)?;
                    stack.push(parser.start(parent, &namespace, &element, decoder, &resolver)?);
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
                Event::Empty(element) => {
                    parser.empty(context(&stack)?, &namespace, &element, decoder, &resolver)?;
                },
                Event::Text(text) => {
                    if let Some(target) = text_target(&stack) {
                        let value = text
                            .decode()
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::CData(text) => {
                    if let Some(target) = text_target(&stack) {
                        let value = text
                            .decode()
                            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                        parser.push_text(target, &value)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if let Some(target) = text_target(&stack) {
                        parser.push_text(target, &decode_xml_reference(&reference)?)?;
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
        if parser.min_row <= parser.max_row && parser.min_column <= parser.max_column {
            parser.data.dimensions = Some((
                parser.min_row,
                parser.min_column,
                parser.max_row,
                parser.max_column,
            ));
        }
        Ok(parser.data)
    }

    fn start(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<Context> {
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            self.sheet_data()?;
            return Ok(Context::SheetData);
        }
        if parent == Context::Worksheet && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            self.start_columns()?;
            return Ok(Context::Columns);
        }
        if parent == Context::Columns && is_spreadsheetml_name(namespace, element.name(), b"col") {
            self.column(element, decoder)?;
            return Ok(Context::Other);
        }
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"mergeCells")
        {
            self.start_merge_cells(element, decoder)?;
            return Ok(Context::MergeCells);
        }
        if parent == Context::MergeCells
            && is_spreadsheetml_name(namespace, element.name(), b"mergeCell")
        {
            self.merge_cell(element, decoder)?;
            return Ok(Context::Other);
        }
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"hyperlinks")
        {
            self.start_hyperlinks()?;
            return Ok(Context::Hyperlinks);
        }
        if parent == Context::Hyperlinks
            && is_spreadsheetml_name(namespace, element.name(), b"hyperlink")
        {
            self.hyperlink(element, decoder, resolver)?;
            return Ok(Context::Other);
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
            self.start_formula(element, decoder, true)?;
            return Ok(Context::Formula);
        }
        if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"v") {
            self.start_value()?;
            return Ok(Context::Value);
        }
        if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"is") {
            self.start_inline()?;
            return Ok(Context::InlineString);
        }
        if parent == Context::InlineString && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::InlineSimple)?;
            return Ok(Context::Text(TextTarget::InlineSimple));
        }
        if parent == Context::InlineString && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            self.start_run()?;
            return Ok(Context::RichRun);
        }
        if parent == Context::RichRun && is_spreadsheetml_name(namespace, element.name(), b"rPr") {
            self.start_properties()?;
            return Ok(Context::RunProperties);
        }
        if parent == Context::RichRun && is_spreadsheetml_name(namespace, element.name(), b"t") {
            self.start_text(TextTarget::InlineRun)?;
            return Ok(Context::Text(TextTarget::InlineRun));
        }
        if parent == Context::RunProperties {
            self.run_property(namespace, element, decoder)?;
        }
        Ok(Context::Other)
    }

    fn empty(
        &mut self,
        parent: Context,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            self.sheet_data()?;
        } else if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            self.start_columns()?;
        } else if parent == Context::Columns
            && is_spreadsheetml_name(namespace, element.name(), b"col")
        {
            self.column(element, decoder)?;
        } else if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"mergeCells")
        {
            self.start_merge_cells(element, decoder)?;
            self.finish_merge_cells()?;
        } else if parent == Context::MergeCells
            && is_spreadsheetml_name(namespace, element.name(), b"mergeCell")
        {
            self.merge_cell(element, decoder)?;
        } else if parent == Context::Worksheet
            && is_spreadsheetml_name(namespace, element.name(), b"hyperlinks")
        {
            self.start_hyperlinks()?;
        } else if parent == Context::Hyperlinks
            && is_spreadsheetml_name(namespace, element.name(), b"hyperlink")
        {
            self.hyperlink(element, decoder, resolver)?;
        } else if parent == Context::SheetData
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            self.start_row(element, decoder)?;
            self.finish_row()?;
        } else if parent == Context::Row && is_spreadsheetml_name(namespace, element.name(), b"c") {
            self.start_cell(element, decoder)?;
            self.finish_cell()?;
        } else if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"f")
        {
            self.start_formula(element, decoder, false)?;
        } else if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"v")
        {
            self.start_value()?;
        } else if parent == Context::Cell && is_spreadsheetml_name(namespace, element.name(), b"is")
        {
            self.start_inline()?;
        } else if parent == Context::InlineString
            && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::InlineSimple)?;
        } else if parent == Context::InlineString
            && is_spreadsheetml_name(namespace, element.name(), b"r")
        {
            return Err(invalid("inline rich-text run is missing its text"));
        } else if parent == Context::RichRun
            && is_spreadsheetml_name(namespace, element.name(), b"rPr")
        {
            self.start_properties()?;
        } else if parent == Context::RichRun
            && is_spreadsheetml_name(namespace, element.name(), b"t")
        {
            self.start_text(TextTarget::InlineRun)?;
        } else if parent == Context::RunProperties {
            self.run_property(namespace, element, decoder)?;
        }
        Ok(())
    }

    fn sheet_data(&mut self) -> Result<()> {
        if self.seen_sheet_data {
            return Err(invalid("duplicate worksheet sheetData element"));
        }
        self.seen_sheet_data = true;
        Ok(())
    }

    fn start_columns(&mut self) -> Result<()> {
        if self.seen_columns {
            return Err(invalid("duplicate worksheet cols element"));
        }
        self.seen_columns = true;
        Ok(())
    }

    fn column(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let min = required_u32(element, b"min", decoder, "worksheet column minimum")?;
        let max = required_u32(element, b"max", decoder, "worksheet column maximum")?;
        if min == 0 || max > MAX_EXCEL_COLUMN || min > max {
            return Err(invalid(format!(
                "invalid worksheet column range '{min}:{max}'"
            )));
        }
        let width = optional_f64(element, b"width", decoder, "worksheet column width")?;
        if let Some(width) = width
            && (!width.is_finite() || width < 0.0)
        {
            return Err(invalid(format!("invalid worksheet column width '{width}'")));
        }
        let info = ColumnInfo {
            width,
            hidden: optional_bool(element, b"hidden", decoder, "worksheet column hidden")?
                .unwrap_or(false),
            custom_width: optional_bool(
                element,
                b"customWidth",
                decoder,
                "worksheet column customWidth",
            )?
            .unwrap_or(false),
        };
        for column in min..=max {
            self.data.columns.insert(column, info.clone());
        }
        Ok(())
    }

    fn start_hyperlinks(&mut self) -> Result<()> {
        if self.seen_hyperlinks {
            return Err(invalid("duplicate worksheet hyperlinks element"));
        }
        self.seen_hyperlinks = true;
        Ok(())
    }

    fn hyperlink(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let cell_ref = required_string(element, b"ref", decoder, "worksheet hyperlink reference")?;
        parse_range(&cell_ref, "worksheet hyperlink reference")?;
        if !self.hyperlink_refs.insert(cell_ref.clone()) {
            return Err(invalid(format!(
                "duplicate worksheet hyperlink reference '{cell_ref}'"
            )));
        }

        let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?;
        if relationship_id.as_deref() == Some("") {
            return Err(invalid("worksheet hyperlink has an empty relationship ID"));
        }
        let location = unqualified_attribute_value(element, b"location", decoder)?;
        if relationship_id.is_none() && location.as_deref().is_none_or(str::is_empty) {
            return Err(invalid(
                "worksheet hyperlink requires a relationship ID or location",
            ));
        }

        self.data.hyperlinks.push(ParsedHyperlink {
            cell_ref,
            relationship_id,
            location,
            display: unqualified_attribute_value(element, b"display", decoder)?,
            tooltip: unqualified_attribute_value(element, b"tooltip", decoder)?,
        });
        Ok(())
    }

    fn start_merge_cells(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.seen_merge_cells {
            return Err(invalid("duplicate worksheet mergeCells element"));
        }
        self.seen_merge_cells = true;
        self.expected_merge_count =
            optional_u32(element, b"count", decoder, "worksheet merged-cell count")?
                .map(|count| count as usize);
        Ok(())
    }

    fn merge_cell(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let reference =
            required_string(element, b"ref", decoder, "worksheet merged-cell reference")?;
        let region = parse_range(&reference, "worksheet merged-cell reference")?;
        if !self.merged_regions.insert(region) {
            return Err(invalid(format!(
                "duplicate worksheet merged-cell reference '{reference}'"
            )));
        }
        self.data.merged_regions.push(region);
        Ok(())
    }

    fn finish_merge_cells(&mut self) -> Result<()> {
        if let Some(expected) = self.expected_merge_count
            && expected != self.data.merged_regions.len()
        {
            return Err(invalid(format!(
                "worksheet mergeCells count is {expected}, but {} mergeCell elements were found",
                self.data.merged_regions.len()
            )));
        }
        Ok(())
    }

    fn start_row(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.row.is_some() {
            return Err(invalid("nested worksheet row"));
        }
        let number = match optional_u32(element, b"r", decoder, "worksheet row number")? {
            Some(number) => number,
            None => self
                .previous_row
                .checked_add(1)
                .ok_or_else(|| invalid("inferred worksheet row number overflows"))?,
        };
        validate_row(number)?;
        if !self.rows.insert(number) {
            return Err(invalid(format!("duplicate worksheet row {number}")));
        }
        let height = optional_f64(element, b"ht", decoder, "worksheet row height")?;
        if let Some(height) = height
            && (!height.is_finite() || height < 0.0)
        {
            return Err(invalid(format!("invalid worksheet row height '{height}'")));
        }
        let hidden =
            optional_bool(element, b"hidden", decoder, "worksheet row hidden")?.unwrap_or(false);
        let custom_height = optional_bool(
            element,
            b"customHeight",
            decoder,
            "worksheet row customHeight",
        )?
        .unwrap_or(false);
        let info = (height.is_some() || hidden || custom_height).then_some(RowInfo {
            height,
            hidden,
            custom_height,
        });
        self.previous_row = number;
        self.row = Some(PendingRow {
            number,
            info,
            last_column: 0,
        });
        Ok(())
    }

    fn start_cell(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.cell.is_some() {
            return Err(invalid("nested worksheet cell"));
        }
        let row_number = self
            .row
            .as_ref()
            .ok_or_else(|| invalid("worksheet cell outside a row"))?
            .number;
        let column = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(reference) => {
                let (column, reference_row) = Cell::reference_to_coords(&reference)
                    .map_err(|error| invalid(error.to_string()))?;
                if reference_row != row_number {
                    return Err(invalid(format!(
                        "cell reference '{reference}' does not belong to worksheet row {row_number}"
                    )));
                }
                column
            },
            None => self
                .row
                .as_ref()
                .ok_or_else(|| invalid("worksheet cell outside a row"))?
                .last_column
                .checked_add(1)
                .filter(|column| *column <= MAX_EXCEL_COLUMN)
                .ok_or_else(|| invalid("inferred worksheet cell column exceeds Excel limits"))?,
        };
        let row = self
            .row
            .as_mut()
            .ok_or_else(|| invalid("worksheet cell outside a row"))?;
        if self
            .data
            .cells
            .get(&row_number)
            .is_some_and(|columns| columns.contains_key(&column))
        {
            return Err(invalid(format!(
                "duplicate worksheet cell at row {row_number}, column {column}"
            )));
        }
        row.last_column = column;
        let cell_type = unqualified_attribute_value(element, b"t", decoder)?;
        if let Some(cell_type) = cell_type.as_deref()
            && !matches!(cell_type, "b" | "e" | "inlineStr" | "n" | "s" | "str" | "d")
        {
            return Err(invalid(format!(
                "invalid worksheet cell type '{cell_type}'"
            )));
        }
        self.cell = Some(PendingCell {
            row: row_number,
            column,
            style: optional_u32(element, b"s", decoder, "worksheet cell style index")?,
            cell_type,
            value: String::new(),
            saw_value: false,
            formula: String::new(),
            saw_formula: false,
            materialized_formula: false,
            is_array_formula: false,
            array_range: None,
            saw_inline_string: false,
            inline_text: String::new(),
            inline_runs: Vec::new(),
            saw_inline_simple: false,
            saw_inline_run: false,
        });
        Ok(())
    }

    fn start_formula(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        materialized: bool,
    ) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("worksheet formula outside a cell"))?;
        if cell.saw_formula {
            return Err(invalid("duplicate worksheet cell formula"));
        }
        cell.saw_formula = true;
        let formula_type =
            unqualified_attribute_value(element, b"t", decoder)?.unwrap_or_else(|| "normal".into());
        if !matches!(
            formula_type.as_str(),
            "normal" | "array" | "dataTable" | "shared"
        ) {
            return Err(invalid(format!(
                "invalid worksheet formula type '{formula_type}'"
            )));
        }
        cell.is_array_formula = formula_type == "array";
        cell.array_range = unqualified_attribute_value(element, b"ref", decoder)?;
        if let Some(range) = cell.array_range.as_deref() {
            validate_range(range, "worksheet formula range")?;
        }
        cell.materialized_formula = materialized;
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
        if cell.saw_inline_string {
            return Err(invalid("duplicate worksheet inline string"));
        }
        cell.saw_inline_string = true;
        Ok(())
    }

    fn start_text(&mut self, target: TextTarget) -> Result<()> {
        match target {
            TextTarget::InlineSimple => {
                let cell = self
                    .cell
                    .as_mut()
                    .ok_or_else(|| invalid("inline text outside a worksheet cell"))?;
                if cell.saw_inline_run {
                    return Err(invalid(
                        "worksheet inline string mixes simple text and rich-text runs",
                    ));
                }
                if cell.saw_inline_simple {
                    return Err(invalid("duplicate simple worksheet inline text"));
                }
                cell.saw_inline_simple = true;
            },
            TextTarget::InlineRun => {
                let run = self
                    .run
                    .as_mut()
                    .ok_or_else(|| invalid("inline text outside a rich-text run"))?;
                if run.saw_text {
                    return Err(invalid("duplicate text in worksheet rich-text run"));
                }
                run.saw_text = true;
            },
            TextTarget::Formula | TextTarget::Value => {},
        }
        Ok(())
    }

    fn start_run(&mut self) -> Result<()> {
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("rich-text run outside a worksheet cell"))?;
        if cell.saw_inline_simple {
            return Err(invalid(
                "worksheet inline string mixes simple text and rich-text runs",
            ));
        }
        if self.run.is_some() {
            return Err(invalid("nested worksheet rich-text run"));
        }
        cell.saw_inline_run = true;
        self.run = Some(PendingRun::new());
        Ok(())
    }

    fn start_properties(&mut self) -> Result<()> {
        let run = self
            .run
            .as_mut()
            .ok_or_else(|| invalid("run properties outside a worksheet rich-text run"))?;
        if run.saw_properties {
            return Err(invalid("duplicate worksheet rich-text run properties"));
        }
        run.saw_properties = true;
        Ok(())
    }

    fn push_text(&mut self, target: TextTarget, value: &str) -> Result<()> {
        match target {
            TextTarget::Formula => self
                .cell
                .as_mut()
                .ok_or_else(|| invalid("formula text outside a worksheet cell"))?
                .formula
                .push_str(value),
            TextTarget::Value => self
                .cell
                .as_mut()
                .ok_or_else(|| invalid("value text outside a worksheet cell"))?
                .value
                .push_str(value),
            TextTarget::InlineSimple => self
                .cell
                .as_mut()
                .ok_or_else(|| invalid("inline text outside a worksheet cell"))?
                .inline_text
                .push_str(value),
            TextTarget::InlineRun => self
                .run
                .as_mut()
                .ok_or_else(|| invalid("inline text outside a rich-text run"))?
                .value
                .text
                .push_str(value),
        }
        Ok(())
    }

    fn run_property(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let run = self
            .run
            .as_mut()
            .ok_or_else(|| invalid("run property outside a worksheet rich-text run"))?;
        if is_spreadsheetml_name(namespace, element.name(), b"rFont") {
            mark(&mut run.properties, 1, "font name")?;
            run.value.font_name = Some(required_string(
                element,
                b"val",
                decoder,
                "rich-text font name",
            )?);
        } else if is_spreadsheetml_name(namespace, element.name(), b"sz") {
            mark(&mut run.properties, 2, "font size")?;
            let value = required_string(element, b"val", decoder, "rich-text font size")?;
            let size = value
                .parse::<f64>()
                .map_err(|_| invalid(format!("invalid rich-text font size '{value}'")))?;
            if !size.is_finite() || size <= 0.0 {
                return Err(invalid(format!("invalid rich-text font size '{value}'")));
            }
            run.value.font_size = Some(size);
        } else if is_spreadsheetml_name(namespace, element.name(), b"b") {
            mark(&mut run.properties, 4, "bold property")?;
            run.value.bold = bool_property(element, decoder, "rich-text bold")?;
        } else if is_spreadsheetml_name(namespace, element.name(), b"i") {
            mark(&mut run.properties, 8, "italic property")?;
            run.value.italic = bool_property(element, decoder, "rich-text italic")?;
        } else if is_spreadsheetml_name(namespace, element.name(), b"u") {
            mark(&mut run.properties, 16, "underline property")?;
            let value = unqualified_attribute_value(element, b"val", decoder)?
                .unwrap_or_else(|| "single".to_string());
            run.value.underline = match value.as_str() {
                "none" => false,
                "single" | "double" | "singleAccounting" | "doubleAccounting" => true,
                _ => {
                    return Err(invalid(format!(
                        "invalid rich-text underline value '{value}'"
                    )));
                },
            };
        } else if is_spreadsheetml_name(namespace, element.name(), b"color") {
            mark(&mut run.properties, 32, "color property")?;
            if let Some(rgb) = unqualified_attribute_value(element, b"rgb", decoder)? {
                if !matches!(rgb.len(), 6 | 8) || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit())
                {
                    return Err(invalid(format!("invalid rich-text RGB color '{rgb}'")));
                }
                run.value.color = Some(rgb);
            }
        }
        Ok(())
    }

    fn finish(&mut self, ended: Context) -> Result<()> {
        match ended {
            Context::RichRun => self.finish_run(),
            Context::Cell => self.finish_cell(),
            Context::Row => self.finish_row(),
            Context::MergeCells => self.finish_merge_cells(),
            _ => Ok(()),
        }
    }

    fn finish_run(&mut self) -> Result<()> {
        let mut run = self
            .run
            .take()
            .ok_or_else(|| invalid("missing worksheet rich-text run"))?;
        if !run.saw_text {
            return Err(invalid("worksheet rich-text run is missing its text"));
        }
        run.value.text = decode_spreadsheet_text(&run.value.text)?;
        let cell = self
            .cell
            .as_mut()
            .ok_or_else(|| invalid("rich-text run outside a worksheet cell"))?;
        cell.inline_text.push_str(&run.value.text);
        cell.inline_runs.push(run.value);
        Ok(())
    }

    fn finish_cell(&mut self) -> Result<()> {
        if self.run.is_some() {
            return Err(invalid("unterminated worksheet rich-text run"));
        }
        let mut cell = self
            .cell
            .take()
            .ok_or_else(|| invalid("missing worksheet cell"))?;
        if cell.saw_inline_string && cell.saw_value {
            return Err(invalid(
                "worksheet cell contains both an inline string and a value",
            ));
        }
        if cell.saw_inline_string && !matches!(cell.cell_type.as_deref(), None | Some("inlineStr"))
        {
            return Err(invalid(
                "worksheet inline string has a non-inline cell type",
            ));
        }
        if cell.cell_type.as_deref() == Some("inlineStr") && !cell.saw_inline_string {
            cell.saw_inline_string = true;
        }
        let base = if cell.saw_inline_string {
            if !cell.saw_inline_run {
                cell.inline_text = decode_spreadsheet_text(&cell.inline_text)?;
            }
            CellValue::String(cell.inline_text)
        } else {
            parse_value(
                cell.cell_type.as_deref(),
                cell.saw_value.then_some(&cell.value),
            )?
        };
        let value = if cell.materialized_formula {
            let cached_value = (!matches!(base, CellValue::Empty)).then_some(Box::new(base));
            CellValue::Formula {
                formula: cell.formula,
                cached_value,
                is_array: cell.is_array_formula,
                array_range: cell.array_range,
            }
        } else {
            base
        };
        let row = cell.row;
        let column = cell.column;
        self.data
            .cells
            .entry(row)
            .or_default()
            .insert(column, value);
        if let Some(style) = cell.style {
            self.data
                .cell_styles
                .entry(row)
                .or_default()
                .insert(column, style);
        }
        if !cell.inline_runs.is_empty() {
            self.data
                .rich_text_cells
                .insert((row, column), cell.inline_runs);
        }
        self.min_row = self.min_row.min(row);
        self.max_row = self.max_row.max(row);
        self.min_column = self.min_column.min(column);
        self.max_column = self.max_column.max(column);
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
        if let Some(info) = row.info {
            self.data.rows.insert(row.number, info);
        }
        Ok(())
    }
}

pub fn parse_worksheet_xml(content: &str) -> SheetResult<HashMap<u32, HashMap<u32, CellValue>>> {
    parse_worksheet_data(content).map(|data| data.cells)
}

pub(crate) fn parse_worksheet_data(content: &str) -> SheetResult<ParsedWorksheetData> {
    Parser::parse(content)
        .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)
}

pub fn parse_sheet_data(
    sheet_data: &str,
    cells: &mut HashMap<u32, HashMap<u32, CellValue>>,
) -> SheetResult<()> {
    let fragment = if first_element_is_sheet_data(sheet_data)? {
        sheet_data.to_string()
    } else {
        format!("<sheetData>{sheet_data}</sheetData>")
    };
    let parsed = parse_worksheet_data(&wrap(&fragment))?;
    for (row, columns) in parsed.cells {
        cells.entry(row).or_default().extend(columns);
    }
    Ok(())
}

#[allow(clippy::type_complexity)]
pub fn parse_row_xml(row_content: &str) -> SheetResult<Option<(u32, Vec<(u32, CellValue)>)>> {
    let Some(row) = fragment_row(row_content, b"row")? else {
        return Ok(None);
    };
    let parsed = parse_worksheet_data(&wrap(&format!("<sheetData>{row_content}</sheetData>")))?;
    let mut values: Vec<_> = parsed
        .cells
        .get(&row)
        .into_iter()
        .flat_map(|columns| columns.iter())
        .map(|(&column, value)| (column, value.clone()))
        .collect();
    values.sort_unstable_by_key(|(column, _)| *column);
    Ok(Some((row, values)))
}

pub fn parse_cell_xml(cell_content: &str) -> SheetResult<Option<(u32, CellValue)>> {
    let Some(row) = fragment_row(cell_content, b"c")? else {
        return Ok(None);
    };
    let parsed = parse_worksheet_data(&wrap(&format!(
        "<sheetData><row r=\"{row}\">{cell_content}</row></sheetData>"
    )))?;
    Ok(parsed.cells.get(&row).and_then(|columns| {
        columns
            .iter()
            .next()
            .map(|(&column, value)| (column, value.clone()))
    }))
}

pub fn reference_to_coords(reference: &str) -> SheetResult<(u32, u32)> {
    Cell::reference_to_coords(reference)
}

fn fragment_row(content: &str, local_name: &[u8]) -> SheetResult<Option<u32>> {
    let mut reader = NsReader::from_reader(content.as_bytes());
    loop {
        let decoder = reader.decoder();
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element)
                if element.name().local_name().as_ref() == local_name =>
            {
                let reference = unqualified_attribute_value(&element, b"r", decoder)
                    .map_err(|error| Box::new(error) as Box<dyn std::error::Error + Send + Sync>)?;
                let row = match reference {
                    Some(reference) if local_name == b"row" => reference.parse::<u32>().map_err(
                        |_| -> Box<dyn std::error::Error + Send + Sync> {
                            format!("Invalid worksheet row number: {reference}").into()
                        },
                    )?,
                    Some(reference) => Cell::reference_to_coords(&reference)?.1,
                    None => 1,
                };
                return Ok(Some(row));
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

fn first_element_is_sheet_data(content: &str) -> SheetResult<bool> {
    let mut reader = NsReader::from_reader(content.as_bytes());
    loop {
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element) => {
                return Ok(element.name().local_name().as_ref() == b"sheetData");
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

fn wrap(fragment: &str) -> String {
    format!(
        r#"<worksheet xmlns="{}">{fragment}</worksheet>"#,
        String::from_utf8_lossy(SPREADSHEETML_NAMESPACE)
    )
}

fn parse_value(cell_type: Option<&str>, value: Option<&String>) -> Result<CellValue> {
    let Some(value) = value else {
        return Ok(CellValue::Empty);
    };
    match cell_type {
        Some("str" | "d") => Ok(CellValue::String(value.clone())),
        Some("s") => {
            let value = value.trim();
            let index = value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid shared-string index '{value}'")))?;
            Ok(CellValue::String(format!("SHARED_STRING_{index}")))
        },
        Some("b") => match value.trim() {
            "1" | "true" => Ok(CellValue::Bool(true)),
            "0" | "false" => Ok(CellValue::Bool(false)),
            value => Err(invalid(format!("invalid worksheet boolean '{value}'"))),
        },
        Some("e") => Ok(CellValue::Error(value.trim().to_string())),
        Some("inlineStr") => Err(invalid(
            "inline-string worksheet cell stores its text in a value element",
        )),
        None | Some("n") => parse_number(value.trim()),
        Some(other) => Err(invalid(format!("invalid worksheet cell type '{other}'"))),
    }
}

fn parse_number(value: &str) -> Result<CellValue> {
    if value.is_empty() {
        return Ok(CellValue::Empty);
    }
    if let Ok(integer) = value.parse::<i64>() {
        return Ok(CellValue::Int(integer));
    }
    fast_float2::parse(value)
        .map(CellValue::Float)
        .map_err(|_| invalid(format!("invalid worksheet numeric value '{value}'")))
}

fn validate_row(row: u32) -> Result<()> {
    Cell::reference_to_coords(&format!("A{row}"))
        .map(|_| ())
        .map_err(|error| invalid(error.to_string()))
}

fn validate_range(range: &str, description: &str) -> Result<()> {
    let mut references = range.split(':');
    let start = references
        .next()
        .ok_or_else(|| invalid(format!("empty {description}")))?;
    Cell::reference_to_coords(start).map_err(|error| invalid(error.to_string()))?;
    if let Some(end) = references.next() {
        Cell::reference_to_coords(end).map_err(|error| invalid(error.to_string()))?;
    }
    if references.next().is_some() {
        return Err(invalid(format!("invalid {description} '{range}'")));
    }
    Ok(())
}

fn context(stack: &[Context]) -> Result<Context> {
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

fn required_string(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<String> {
    unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("missing {description} attribute")))
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
                .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
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

fn parse_range(range: &str, description: &str) -> Result<(u32, u32, u32, u32)> {
    let mut references = range.split(':');
    let start = references
        .next()
        .ok_or_else(|| invalid(format!("empty {description}")))?;
    let (start_column, start_row) =
        Cell::reference_to_coords(start).map_err(|error| invalid(error.to_string()))?;
    let (end_column, end_row) = match references.next() {
        Some(end) => Cell::reference_to_coords(end).map_err(|error| invalid(error.to_string()))?,
        None => (start_column, start_row),
    };
    if references.next().is_some() || start_row > end_row || start_column > end_column {
        return Err(invalid(format!("invalid {description} '{range}'")));
    }
    Ok((start_row, start_column, end_row, end_column))
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
                .map_err(|_| invalid(format!("invalid {description} value '{value}'")))
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
            _ => Err(invalid(format!("invalid {description} value '{value}'"))),
        })
        .transpose()
}

fn bool_property(element: &BytesStart<'_>, decoder: Decoder, description: &str) -> Result<bool> {
    match unqualified_attribute_value(element, b"val", decoder)?.as_deref() {
        None | Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        Some(value) => Err(invalid(format!("invalid {description} value '{value}'"))),
    }
}

fn mark(seen: &mut u8, bit: u8, description: &str) -> Result<()> {
    if *seen & bit != 0 {
        return Err(invalid(format!(
            "duplicate worksheet rich-text {description}"
        )));
    }
    *seen |= bit;
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    #[test]
    fn parses_cell_types_formulas_rows_and_inline_text() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:f="urn:foreign">
                <f:sheetData><x:row r="1"><x:c r="A1"><x:v>999</x:v></x:c></x:row></f:sheetData>
                <x:sheetData>
                    <x:row r="2" ht="20" hidden="0" customHeight="true">
                        <x:c r="A2" s="3"><x:v>42</x:v></x:c>
                        <x:c r="B2" t="b"><x:v>true</x:v></x:c>
                        <x:c r="C2" t="e"><x:v>#DIV/0!</x:v></x:c>
                        <x:c r="D2" t="s"><x:v>7</x:v></x:c>
                        <x:c r="E2" t="str"><x:f t="array" ref="E2:E3">CONCAT(&quot;A&amp;B&quot;)</x:f>
                            <x:v>done &amp;</x:v></x:c>
                        <x:c r="F2" t="inlineStr"><x:is><x:r><x:rPr><x:b val="0"/><x:i/>
                            <x:color rgb="FF112233"/></x:rPr><x:t>Hi_x000A_</x:t></x:r>
                            <x:r><x:t><![CDATA[there]]></x:t></x:r></x:is></x:c>
                    </x:row>
                    <x:row><x:c><x:v>1.5</x:v></x:c></x:row>
                </x:sheetData>
            </x:worksheet>"#
        );
        let data = parse_worksheet_data(&xml).unwrap();
        assert_eq!(data.dimensions, Some((2, 1, 3, 6)));
        assert_eq!(data.cells[&2][&1], CellValue::Int(42));
        assert_eq!(data.cells[&2][&2], CellValue::Bool(true));
        assert_eq!(data.cells[&2][&3], CellValue::Error("#DIV/0!".to_string()));
        assert_eq!(
            data.cells[&2][&4],
            CellValue::String("SHARED_STRING_7".to_string())
        );
        match &data.cells[&2][&5] {
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => {
                assert_eq!(formula, "CONCAT(\"A&B\")");
                assert_eq!(
                    cached_value.as_deref(),
                    Some(&CellValue::String("done &".to_string()))
                );
                assert!(*is_array);
                assert_eq!(array_range.as_deref(), Some("E2:E3"));
            },
            other => panic!("unexpected formula cell {other:?}"),
        }
        assert_eq!(
            data.cells[&2][&6],
            CellValue::String("Hi\nthere".to_string())
        );
        let runs = &data.rich_text_cells[&(2, 6)];
        assert_eq!(runs.len(), 2);
        assert!(!runs[0].bold);
        assert!(runs[0].italic);
        assert_eq!(runs[0].color.as_deref(), Some("FF112233"));
        assert_eq!(data.cells[&3][&1], CellValue::Float(1.5));
        assert_eq!(data.cell_styles[&2][&1], 3);
        assert_eq!(data.rows[&2].height, Some(20.0));
        assert!(data.rows[&2].custom_height);
        assert!(!data.rows[&2].hidden);
    }

    #[test]
    fn accepts_strict_namespaces_and_ignores_foreign_lookalikes() {
        let xml = format!(
            r#"<worksheet xmlns="{STRICT_S}" xmlns:f="urn:foreign"><sheetData>
                <f:row r="1"><row r="1"><c r="A1"><v>99</v></c></row></f:row>
                <row r="4"><f:c r="A4"><c r="A4"><v>88</v></c></f:c>
                    <c r="B4" t="inlineStr"><is><f:t>ignored</f:t><t>strict</t></is></c></row>
            </sheetData></worksheet>"#
        );
        let data = parse_worksheet_data(&xml).unwrap();
        assert_eq!(data.cells.len(), 1);
        assert_eq!(data.cells[&4][&2], CellValue::String("strict".to_string()));
    }

    #[test]
    fn parses_columns_and_merged_regions_namespace_aware() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{STRICT_S}" xmlns:f="urn:foreign">
                <f:cols><x:col min="1" max="16384" hidden="1"/></f:cols>
                <x:cols>
                    <f:col min="1" max="1" width="99"/>
                    <x:col min="2" max="3" width="12.5" hidden="true" customWidth="0"/>
                    <x:col min="3" max="3" hidden="1"></x:col>
                </x:cols>
                <f:mergeCells><x:mergeCell ref="A1:XFD1048576"/></f:mergeCells>
                <x:mergeCells count="2">
                    <x:mergeCell ref="B2:C3"/>
                    <x:mergeCell ref="D4"></x:mergeCell>
                </x:mergeCells>
            </x:worksheet>"#
        );
        let data = parse_worksheet_data(&xml).unwrap();

        assert_eq!(data.columns.len(), 2);
        assert_eq!(data.columns[&2].width, Some(12.5));
        assert!(data.columns[&2].hidden);
        assert!(!data.columns[&2].custom_width);
        assert_eq!(data.columns[&3].width, None);
        assert!(data.columns[&3].hidden);
        assert_eq!(data.merged_regions, vec![(2, 2, 3, 3), (4, 4, 4, 4)]);
    }

    #[test]
    fn parses_hyperlinks_namespace_aware() {
        let xml = format!(
            r#"<x:worksheet xmlns:x="{STRICT_S}"
                    xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"
                    xmlns:f="urn:foreign">
                <f:hyperlinks><x:hyperlink ref="A1" location="Ignored"/></f:hyperlinks>
                <x:hyperlinks>
                    <f:hyperlink ref="A1" location="Ignored"/>
                    <x:hyperlink ref="A1:B2" rel:id="customRel" location="Section 1"
                        display="Example &amp; Co" tooltip="Open &quot;example&quot;"/>
                    <x:hyperlink ref="D4" location="&apos;Other Sheet&apos;!A1"></x:hyperlink>
                </x:hyperlinks>
            </x:worksheet>"#
        );
        let data = parse_worksheet_data(&xml).unwrap();

        assert_eq!(data.hyperlinks.len(), 2);
        assert_eq!(data.hyperlinks[0].cell_ref, "A1:B2");
        assert_eq!(
            data.hyperlinks[0].relationship_id.as_deref(),
            Some("customRel")
        );
        assert_eq!(data.hyperlinks[0].location.as_deref(), Some("Section 1"));
        assert_eq!(data.hyperlinks[0].display.as_deref(), Some("Example & Co"));
        assert_eq!(
            data.hyperlinks[0].tooltip.as_deref(),
            Some("Open \"example\"")
        );
        assert_eq!(
            data.hyperlinks[1].location.as_deref(),
            Some("'Other Sheet'!A1")
        );
    }

    #[test]
    fn rejects_invalid_columns_and_merged_regions() {
        let invalid_documents = [
            "<cols><col min=\"0\" max=\"1\"/></cols>",
            "<cols><col min=\"2\" max=\"1\"/></cols>",
            "<cols><col min=\"1\" max=\"16385\"/></cols>",
            "<cols><col min=\"1\" max=\"1\" width=\"NaN\"/></cols>",
            "<cols><col min=\"1\" max=\"1\" hidden=\"TRUE\"/></cols>",
            "<mergeCells count=\"2\"><mergeCell ref=\"A1:B2\"/></mergeCells>",
            "<mergeCells><mergeCell ref=\"C3:B2\"/></mergeCells>",
            "<mergeCells><mergeCell ref=\"A1:B2\"/><mergeCell ref=\"A1:B2\"/></mergeCells>",
        ];

        for fragment in invalid_documents {
            let xml = wrap(fragment);
            assert!(
                parse_worksheet_data(&xml).is_err(),
                "accepted invalid worksheet fragment: {fragment}"
            );
        }
    }

    #[test]
    fn rejects_invalid_hyperlinks() {
        let invalid_documents = [
            "<hyperlinks><hyperlink location=\"Sheet2!A1\"/></hyperlinks>",
            "<hyperlinks><hyperlink ref=\"A1\"/></hyperlinks>",
            "<hyperlinks><hyperlink ref=\"B2:A1\" location=\"Sheet2!A1\"/></hyperlinks>",
            "<hyperlinks><hyperlink ref=\"A1\" location=\"Sheet2!A1\"/><hyperlink ref=\"A1\" location=\"Sheet3!A1\"/></hyperlinks>",
            "<hyperlinks/><hyperlinks/>",
        ];

        for fragment in invalid_documents {
            let xml = wrap(fragment);
            assert!(
                parse_worksheet_data(&xml).is_err(),
                "accepted invalid worksheet fragment: {fragment}"
            );
        }
    }

    #[test]
    fn rejects_malformed_truncated_and_inconsistent_sheet_data() {
        for xml in [
            format!(r#"<worksheet xmlns="{S}"><sheetData><row r="0"/></sheetData></worksheet>"#),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="2"><c r="A3"/></row></sheetData></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1" t="bad"><v>1</v></c></row></sheetData></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>bad</v></c></row></sheetData></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><v>1</v><v>2</v></c></row></sheetData></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"/><row r="1"/></sheetData></worksheet>"#
            ),
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f/><f/></c></row></sheetData></worksheet>"#
            ),
            format!(r#"<worksheet xmlns="{S}"><sheetData/><sheetData/></worksheet>"#),
            format!(r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1">"#),
        ] {
            assert!(parse_worksheet_data(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn standalone_helpers_preserve_empty_rows_and_inner_sheet_data() {
        let row = parse_row_xml(r#"<row r="7"/>"#).unwrap().unwrap();
        assert_eq!(row, (7, Vec::new()));

        let mut cells = HashMap::new();
        parse_sheet_data(r#"<row r="2"><c r="B2"><v>4</v></c></row>"#, &mut cells).unwrap();
        assert_eq!(cells[&2][&2], CellValue::Int(4));
        assert_eq!(
            parse_cell_xml(r#"<c r="C9" t="b"><v>1</v></c>"#)
                .unwrap()
                .unwrap(),
            (3, CellValue::Bool(true))
        );

        let hidden = parse_worksheet_data(&format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1" ht="0" hidden="1"/></sheetData></worksheet>"#
        ))
        .unwrap();
        assert_eq!(hidden.rows[&1].height, Some(0.0));
        assert!(hidden.rows[&1].hidden);
    }
}
