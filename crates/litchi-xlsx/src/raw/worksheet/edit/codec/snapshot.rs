//! Exact-span worksheet scanner and XML patch writer.

use std::collections::{BTreeMap, HashMap};

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::{COLUMNS, Cell as Address, Column, ROWS, Rect, Row};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::super::{
    merge_successor, optional_bool, optional_u32, parse_a1, parse_one_based_row, x14ac,
};
use super::super::model::{
    Action, ColumnAction, DefaultsEffects, DescentEffect, HeightEffect, OptionalEffect, Payload,
    RowAction, SelectionRange, StyleEffect, WidthEffect,
};
use crate::cell::{Content, Value};
use crate::column::Assignments;
use crate::error::{ColumnEditBlock, Error, Result, allocation, invalid};
use crate::merge;
use crate::outline::Outline;
use crate::raw::namespace::is_spreadsheetml_name;
use crate::raw::strings::encode_spreadsheet_text;

use super::X14;
use super::wire::{
    column_range, is_mce_name, position, sibling_name, tag, write_attribute, write_close, write_tag,
};
#[derive(Debug, Clone, Copy)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) end: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct Attribute {
    pub(crate) name: Box<str>,
    pub(crate) value: Box<str>,
}

#[derive(Debug, Clone)]
pub(crate) struct Tag {
    pub(crate) name: Box<str>,
    pub(crate) attributes: Box<[Attribute]>,
}

#[derive(Debug)]
pub(crate) struct CellSlot {
    pub(crate) address: Address,
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) primary: Box<[Span]>,
    pub(crate) mce_payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct RowSlot {
    pub(crate) number: u32,
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) descent_attribute: Option<Box<str>>,
    pub(crate) cells: Box<[CellSlot]>,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct DefaultsSlot {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) descent_attribute: Option<Box<str>>,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct RootSlot {
    pub(crate) span: Span,
    pub(crate) tag: Tag,
}

#[derive(Debug)]
pub(crate) struct ColumnSlot {
    pub(crate) first: Column,
    pub(crate) last: Column,
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct ColumnsSlot {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) columns: Box<[ColumnSlot]>,
    pub(crate) payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct SheetData {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) rows: Box<[RowSlot]>,
    pub(crate) empty: bool,
}

#[derive(Debug)]
pub(crate) struct DimensionTag {
    pub(crate) span: Span,
    pub(crate) tag: Tag,
    pub(crate) empty: bool,
    pub(crate) declared: Rect,
}

#[derive(Debug)]
pub(crate) struct MergeSlot {
    pub(crate) range: Rect,
    pub(crate) span: Span,
}

#[derive(Debug)]
pub(crate) struct MergeCellsSlot {
    pub(crate) span: Span,
    pub(crate) tag_end: usize,
    pub(crate) close_start: usize,
    pub(crate) tag: Tag,
    pub(crate) merges: Box<[MergeSlot]>,
    pub(crate) payload: bool,
    pub(crate) empty: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Worksheet,
    Defaults,
    Columns,
    Column,
    SheetData,
    Row,
    Cell,
    Primary,
    MergeCells,
    Merge,
    Other,
}

#[derive(Debug, Clone, Copy)]
struct Frame {
    kind: FrameKind,
    start: usize,
}

#[derive(Debug)]
struct PendingCell {
    address: Address,
    start: usize,
    tag_end: usize,
    tag: Tag,
    primary: Vec<Span>,
    mce_payload: bool,
}

#[derive(Debug)]
struct PendingRow {
    number: u32,
    last_column: u32,
    start: usize,
    tag_end: usize,
    tag: Tag,
    descent_attribute: Option<Box<str>>,
    cells: Vec<CellSlot>,
}

#[derive(Debug)]
struct PendingDefaults {
    start: usize,
    tag_end: usize,
    tag: Tag,
    descent_attribute: Option<Box<str>>,
}

#[derive(Debug)]
struct PendingColumn {
    first: Column,
    last: Column,
    start: usize,
    tag_end: usize,
    tag: Tag,
    payload: bool,
}

#[derive(Debug)]
struct PendingColumns {
    start: usize,
    tag_end: usize,
    tag: Tag,
    columns: Vec<ColumnSlot>,
    payload: bool,
}

#[derive(Debug)]
struct PendingSheetData {
    start: usize,
    tag_end: usize,
    tag: Tag,
    rows: Vec<RowSlot>,
}

#[derive(Debug)]
struct PendingMergeCells {
    start: usize,
    tag_end: usize,
    tag: Tag,
    count: Option<usize>,
    merges: Vec<MergeSlot>,
    payload: bool,
}

#[derive(Debug)]
struct PendingMerge {
    range: Rect,
    start: usize,
}

fn merge_range(element: &BytesStart<'_>, decoder: Decoder) -> Result<Rect> {
    let value = unqualified_attribute_value(element, b"ref", decoder)?
        .ok_or_else(|| invalid("mergeCell is missing ref during edit"))?;
    let range = Rect::from_a1(&value).map_err(|error| {
        invalid(format!(
            "invalid merged range '{value}' during edit: {error}"
        ))
    })?;
    if range.rows() == 1 && range.columns() == 1 {
        return Err(invalid(format!(
            "merged range '{value}' contains only one cell during edit"
        )));
    }
    Ok(range)
}

fn merge_predecessor(local: &[u8]) -> bool {
    matches!(
        local,
        b"sheetData"
            | b"sheetCalcPr"
            | b"sheetProtection"
            | b"protectedRanges"
            | b"scenarios"
            | b"autoFilter"
            | b"sortState"
            | b"dataConsolidate"
            | b"customSheetViews"
    )
}

#[derive(Debug)]
struct FormulaStorage {
    address: Address,
    kind: Box<str>,
    index: Option<u32>,
    range: Option<SelectionRange>,
}

#[derive(Debug)]
pub(crate) struct Layout {
    pub(crate) root: RootSlot,
    pub(crate) defaults: Option<DefaultsSlot>,
    pub(crate) sheet_data: SheetData,
    pub(crate) columns: Option<ColumnsSlot>,
    pub(crate) dimension: Option<DimensionTag>,
    pub(crate) protected: bool,
    pub(crate) merged: Box<[SelectionRange]>,
    pub(crate) validations: Box<[SelectionRange]>,
    pub(crate) extended_validation: bool,
    pub(crate) formula_ranges: Box<[SelectionRange]>,
    pub(crate) defaults_compatibility: bool,
    pub(crate) merge_cells: Option<MergeCellsSlot>,
    pub(crate) merge_insertion: usize,
    pub(crate) merge_compatibility: bool,
}

#[derive(Debug, Default)]
struct Scanner {
    root: Option<RootSlot>,
    defaults: Option<DefaultsSlot>,
    pending_defaults: Option<PendingDefaults>,
    sheet_data: Option<SheetData>,
    columns: Option<ColumnsSlot>,
    dimension: Option<DimensionTag>,
    pending_sheet_data: Option<PendingSheetData>,
    pending_columns: Option<PendingColumns>,
    column: Option<PendingColumn>,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    previous_row: u32,
    protected: bool,
    validations: Vec<SelectionRange>,
    extended_validation: bool,
    formulas: Vec<FormulaStorage>,
    defaults_compatibility: bool,
    merge_cells: Option<MergeCellsSlot>,
    pending_merge_cells: Option<PendingMergeCells>,
    pending_merge: Option<PendingMerge>,
    merge_insertion: Option<usize>,
    merge_compatibility: bool,
    root_close_start: Option<usize>,
}

#[derive(Debug)]
pub(crate) struct RootEffect {
    pub(crate) removed: Option<Box<str>>,
    pub(crate) appended: Vec<(Box<str>, String)>,
}

pub(crate) fn scan(content: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(content);
    let mut scanner = Scanner::default();
    let mut stack = Vec::<Frame>::new();

    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let parent = stack.last().map(|frame| frame.kind);
                let kind = scanner.start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                )?;
                stack.push(Frame {
                    kind,
                    start: event_start,
                });
            },
            Event::Empty(element) => {
                let parent = stack.last().map(|frame| frame.kind);
                scanner.empty(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                )?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("worksheet edit scan has an unmatched closing tag"))?;
                scanner.finish(frame, event_start, event_end)?;
            },
            Event::Text(value) => {
                if stack.last().is_some_and(|frame| {
                    matches!(frame.kind, FrameKind::MergeCells | FrameKind::Merge)
                }) && !value
                    .decode()
                    .map_err(|error| invalid(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    scanner.mark_merge_payload();
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                if stack.last().is_some_and(|frame| {
                    matches!(frame.kind, FrameKind::MergeCells | FrameKind::Merge)
                }) {
                    scanner.mark_merge_payload();
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("worksheet edit scan ended inside an element"));
    }
    scanner.finish_layout()
}

impl Scanner {
    fn start(
        &mut self,
        parent: Option<FrameKind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        span: Span,
    ) -> Result<FrameKind> {
        let Span { start, end } = span;
        self.scan_guard(namespace, element, decoder)?;
        if parent.is_none() && is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
            if self.root.is_some() {
                return Err(invalid("worksheet edit scan found duplicate roots"));
            }
            self.root = Some(RootSlot {
                span,
                tag: tag(element, decoder)?,
            });
            return Ok(FrameKind::Worksheet);
        }
        self.observe_merge_position(parent, namespace, element, span.start);
        if is_spreadsheetml_name(namespace, element.name(), b"mergeCells") {
            if parent != Some(FrameKind::Worksheet) {
                self.merge_compatibility = true;
                return Ok(FrameKind::Other);
            }
            self.start_merge_cells(element, decoder, start, end)?;
            return Ok(FrameKind::MergeCells);
        }
        if is_spreadsheetml_name(namespace, element.name(), b"mergeCell") {
            if parent != Some(FrameKind::MergeCells) {
                self.merge_compatibility = true;
                return Ok(FrameKind::Other);
            }
            let range = merge_range(element, decoder)?;
            self.pending_merge = Some(PendingMerge { range, start });
            return Ok(FrameKind::Merge);
        }
        if matches!(parent, Some(FrameKind::MergeCells | FrameKind::Merge)) {
            self.mark_merge_payload();
            return Ok(FrameKind::Other);
        }
        if is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr") {
            if parent != Some(FrameKind::Worksheet) {
                self.defaults_compatibility = true;
                return Ok(FrameKind::Other);
            }
            if self.pending_defaults.is_some() || self.defaults.is_some() {
                return Err(invalid("worksheet has duplicate sheetFormatPr during edit"));
            }
            if self.pending_columns.is_some()
                || self.columns.is_some()
                || self.pending_sheet_data.is_some()
                || self.sheet_data.is_some()
            {
                return Err(invalid(
                    "worksheet sheetFormatPr appears after column or cell data during edit",
                ));
            }
            self.pending_defaults = Some(PendingDefaults {
                start: span.start,
                tag_end: span.end,
                tag: tag(element, decoder)?,
                descent_attribute: x14ac::attribute_name(element, resolver)?,
            });
            return Ok(FrameKind::Defaults);
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"dimension")
        {
            self.record_dimension(element, decoder, span, false)?;
            return Ok(FrameKind::Other);
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            if self.pending_columns.is_some() || self.columns.is_some() {
                return Err(invalid("worksheet has duplicate cols during edit"));
            }
            if self.sheet_data.is_some() || self.pending_sheet_data.is_some() {
                return Err(invalid(
                    "worksheet cols appears after sheetData during edit",
                ));
            }
            self.pending_columns = Some(PendingColumns {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
                columns: Vec::new(),
                payload: false,
            });
            return Ok(FrameKind::Columns);
        }
        if parent == Some(FrameKind::Columns)
            && is_spreadsheetml_name(namespace, element.name(), b"col")
        {
            let (first, last) = column_range(element, decoder)?;
            self.column = Some(PendingColumn {
                first,
                last,
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
                payload: false,
            });
            return Ok(FrameKind::Column);
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            if self.pending_sheet_data.is_some() || self.sheet_data.is_some() {
                return Err(invalid("worksheet has duplicate sheetData during edit"));
            }
            self.pending_sheet_data = Some(PendingSheetData {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
                rows: Vec::new(),
            });
            return Ok(FrameKind::SheetData);
        }
        if parent == Some(FrameKind::SheetData)
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            self.start_row(element, decoder, resolver, start, end)?;
            return Ok(FrameKind::Row);
        }
        if parent == Some(FrameKind::Row) && is_spreadsheetml_name(namespace, element.name(), b"c")
        {
            self.start_cell(element, decoder, start, end)?;
            return Ok(FrameKind::Cell);
        }
        if parent == Some(FrameKind::Cell)
            && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is")
            && is_spreadsheetml_name(
                namespace,
                element.name(),
                element.name().local_name().as_ref(),
            )
        {
            if element.name().local_name().as_ref() == b"f" {
                self.scan_formula(element, decoder)?;
            }
            return Ok(FrameKind::Primary);
        }
        if self.cell.is_some()
            && (is_mce_name(namespace, element, b"AlternateContent")
                || (parent == Some(FrameKind::Cell)
                    && !is_spreadsheetml_name(namespace, element.name(), b"extLst")))
        {
            // Unknown direct cell children may carry future value semantics.
            // Keeping them beside a replacement payload could silently create
            // two competing representations, so the ordinary editor refuses.
            if let Some(cell) = self.cell.as_mut() {
                cell.mce_payload = true;
            }
        }
        if parent == Some(FrameKind::Column)
            && let Some(column) = self.column.as_mut()
        {
            column.payload = true;
        }
        if parent == Some(FrameKind::Columns)
            && let Some(columns) = self.pending_columns.as_mut()
        {
            columns.payload = true;
        }
        Ok(FrameKind::Other)
    }

    fn empty(
        &mut self,
        parent: Option<FrameKind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        span: Span,
    ) -> Result<()> {
        self.scan_guard(namespace, element, decoder)?;
        self.observe_merge_position(parent, namespace, element, span.start);
        if is_spreadsheetml_name(namespace, element.name(), b"mergeCells") {
            if parent != Some(FrameKind::Worksheet) {
                self.merge_compatibility = true;
                return Ok(());
            }
            self.ensure_merge_cells_slot()?;
            self.merge_cells = Some(MergeCellsSlot {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                merges: Box::new([]),
                payload: false,
                empty: true,
            });
            return Ok(());
        }
        if is_spreadsheetml_name(namespace, element.name(), b"mergeCell") {
            if parent != Some(FrameKind::MergeCells) {
                self.merge_compatibility = true;
                return Ok(());
            }
            let range = merge_range(element, decoder)?;
            self.pending_merge_cells
                .as_mut()
                .ok_or_else(|| invalid("mergeCell appears outside mergeCells edit state"))?
                .merges
                .push(MergeSlot { range, span });
            return Ok(());
        }
        if matches!(parent, Some(FrameKind::MergeCells | FrameKind::Merge)) {
            self.mark_merge_payload();
            return Ok(());
        }
        if is_spreadsheetml_name(namespace, element.name(), b"sheetFormatPr") {
            if parent != Some(FrameKind::Worksheet) {
                self.defaults_compatibility = true;
                return Ok(());
            }
            if self.pending_defaults.is_some() || self.defaults.is_some() {
                return Err(invalid("worksheet has duplicate sheetFormatPr during edit"));
            }
            if self.pending_columns.is_some()
                || self.columns.is_some()
                || self.pending_sheet_data.is_some()
                || self.sheet_data.is_some()
            {
                return Err(invalid(
                    "worksheet sheetFormatPr appears after column or cell data during edit",
                ));
            }
            self.defaults = Some(DefaultsSlot {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                descent_attribute: x14ac::attribute_name(element, resolver)?,
                empty: true,
            });
            return Ok(());
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"dimension")
        {
            self.record_dimension(element, decoder, span, true)?;
            return Ok(());
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"cols")
        {
            if self.pending_columns.is_some() || self.columns.is_some() {
                return Err(invalid("worksheet has duplicate cols during edit"));
            }
            if self.sheet_data.is_some() || self.pending_sheet_data.is_some() {
                return Err(invalid(
                    "worksheet cols appears after sheetData during edit",
                ));
            }
            self.columns = Some(ColumnsSlot {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                columns: Box::new([]),
                payload: false,
                empty: true,
            });
            return Ok(());
        }
        if parent == Some(FrameKind::Columns)
            && is_spreadsheetml_name(namespace, element.name(), b"col")
        {
            let (first, last) = column_range(element, decoder)?;
            self.pending_columns
                .as_mut()
                .ok_or_else(|| invalid("empty col outside cols"))?
                .columns
                .push(ColumnSlot {
                    first,
                    last,
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    payload: false,
                    empty: true,
                });
            return Ok(());
        }
        if parent == Some(FrameKind::Worksheet)
            && is_spreadsheetml_name(namespace, element.name(), b"sheetData")
        {
            if self.pending_sheet_data.is_some() || self.sheet_data.is_some() {
                return Err(invalid("worksheet has duplicate sheetData during edit"));
            }
            self.sheet_data = Some(SheetData {
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                rows: Box::new([]),
                empty: true,
            });
            return Ok(());
        }
        if parent == Some(FrameKind::SheetData)
            && is_spreadsheetml_name(namespace, element.name(), b"row")
        {
            let (number, _) = self.row_position(element, decoder)?;
            let row = RowSlot {
                number,
                span,
                tag_end: span.end,
                close_start: span.end,
                tag: tag(element, decoder)?,
                descent_attribute: x14ac::attribute_name(element, resolver)?,
                cells: Box::new([]),
                empty: true,
            };
            self.pending_sheet_data
                .as_mut()
                .ok_or_else(|| invalid("empty row outside sheetData"))?
                .rows
                .push(row);
            return Ok(());
        }
        if parent == Some(FrameKind::Row) && is_spreadsheetml_name(namespace, element.name(), b"c")
        {
            let address = self.cell_address(element, decoder)?;
            self.row
                .as_mut()
                .ok_or_else(|| invalid("empty cell outside row"))?
                .cells
                .push(CellSlot {
                    address,
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    primary: Box::new([]),
                    mce_payload: false,
                    empty: true,
                });
            return Ok(());
        }
        if parent == Some(FrameKind::Cell)
            && matches!(element.name().local_name().as_ref(), b"f" | b"v" | b"is")
            && is_spreadsheetml_name(
                namespace,
                element.name(),
                element.name().local_name().as_ref(),
            )
        {
            if element.name().local_name().as_ref() == b"f" {
                self.scan_formula(element, decoder)?;
            }
            self.cell
                .as_mut()
                .ok_or_else(|| invalid("empty cell payload outside cell"))?
                .primary
                .push(span);
        } else if self.cell.is_some()
            && (is_mce_name(namespace, element, b"AlternateContent")
                || (parent == Some(FrameKind::Cell)
                    && !is_spreadsheetml_name(namespace, element.name(), b"extLst")))
            && let Some(cell) = self.cell.as_mut()
        {
            cell.mce_payload = true;
        }
        if parent == Some(FrameKind::Column)
            && let Some(column) = self.column.as_mut()
        {
            column.payload = true;
        }
        if parent == Some(FrameKind::Columns)
            && let Some(columns) = self.pending_columns.as_mut()
        {
            columns.payload = true;
        }
        Ok(())
    }

    fn finish(&mut self, frame: Frame, close_start: usize, end: usize) -> Result<()> {
        match frame.kind {
            FrameKind::Worksheet => {
                self.root_close_start = Some(close_start);
            },
            FrameKind::Merge => {
                let merge = self
                    .pending_merge
                    .take()
                    .ok_or_else(|| invalid("mergeCell close without edit state"))?;
                self.pending_merge_cells
                    .as_mut()
                    .ok_or_else(|| invalid("mergeCell closed outside mergeCells"))?
                    .merges
                    .push(MergeSlot {
                        range: merge.range,
                        span: Span {
                            start: merge.start,
                            end,
                        },
                    });
            },
            FrameKind::MergeCells => {
                let merges = self
                    .pending_merge_cells
                    .take()
                    .ok_or_else(|| invalid("mergeCells close without edit state"))?;
                if merges
                    .count
                    .is_some_and(|count| count != merges.merges.len())
                {
                    return Err(invalid(format!(
                        "worksheet merged-range count differs from {} records during edit",
                        merges.merges.len()
                    )));
                }
                self.merge_cells = Some(MergeCellsSlot {
                    span: Span {
                        start: merges.start,
                        end,
                    },
                    tag_end: merges.tag_end,
                    close_start,
                    tag: merges.tag,
                    merges: merges.merges.into_boxed_slice(),
                    payload: merges.payload,
                    empty: false,
                });
            },
            FrameKind::Column => {
                let column = self
                    .column
                    .take()
                    .ok_or_else(|| invalid("col close without edit state"))?;
                self.pending_columns
                    .as_mut()
                    .ok_or_else(|| invalid("col closed outside cols"))?
                    .columns
                    .push(ColumnSlot {
                        first: column.first,
                        last: column.last,
                        span: Span {
                            start: column.start,
                            end,
                        },
                        tag_end: column.tag_end,
                        close_start,
                        tag: column.tag,
                        payload: column.payload,
                        empty: false,
                    });
            },
            FrameKind::Defaults => {
                let defaults = self
                    .pending_defaults
                    .take()
                    .ok_or_else(|| invalid("sheetFormatPr close without edit state"))?;
                self.defaults = Some(DefaultsSlot {
                    span: Span {
                        start: defaults.start,
                        end,
                    },
                    tag_end: defaults.tag_end,
                    close_start,
                    tag: defaults.tag,
                    descent_attribute: defaults.descent_attribute,
                    empty: false,
                });
            },
            FrameKind::Columns => {
                let columns = self
                    .pending_columns
                    .take()
                    .ok_or_else(|| invalid("cols close without edit state"))?;
                if columns.columns.is_empty() {
                    return Err(invalid("worksheet cols contains no col during edit"));
                }
                self.columns = Some(ColumnsSlot {
                    span: Span {
                        start: columns.start,
                        end,
                    },
                    tag_end: columns.tag_end,
                    close_start,
                    tag: columns.tag,
                    columns: columns.columns.into_boxed_slice(),
                    payload: columns.payload,
                    empty: false,
                });
            },
            FrameKind::Primary => {
                self.cell
                    .as_mut()
                    .ok_or_else(|| invalid("cell payload closed outside a cell"))?
                    .primary
                    .push(Span {
                        start: frame.start,
                        end,
                    });
            },
            FrameKind::Cell => {
                let cell = self
                    .cell
                    .take()
                    .ok_or_else(|| invalid("cell close without edit state"))?;
                self.row
                    .as_mut()
                    .ok_or_else(|| invalid("cell closed outside a row"))?
                    .cells
                    .push(CellSlot {
                        address: cell.address,
                        span: Span {
                            start: cell.start,
                            end,
                        },
                        tag_end: cell.tag_end,
                        close_start,
                        tag: cell.tag,
                        primary: cell.primary.into_boxed_slice(),
                        mce_payload: cell.mce_payload,
                        empty: false,
                    });
            },
            FrameKind::Row => {
                let row = self
                    .row
                    .take()
                    .ok_or_else(|| invalid("row close without edit state"))?;
                if row
                    .cells
                    .windows(2)
                    .any(|pair| pair[0].address >= pair[1].address)
                {
                    return Err(invalid(
                        "cell edits require strictly increasing cell references within each row",
                    ));
                }
                self.pending_sheet_data
                    .as_mut()
                    .ok_or_else(|| invalid("row closed outside sheetData"))?
                    .rows
                    .push(RowSlot {
                        number: row.number,
                        span: Span {
                            start: row.start,
                            end,
                        },
                        tag_end: row.tag_end,
                        close_start,
                        tag: row.tag,
                        descent_attribute: row.descent_attribute,
                        cells: row.cells.into_boxed_slice(),
                        empty: false,
                    });
            },
            FrameKind::SheetData => {
                let data = self
                    .pending_sheet_data
                    .take()
                    .ok_or_else(|| invalid("sheetData close without edit state"))?;
                self.sheet_data = Some(SheetData {
                    span: Span {
                        start: data.start,
                        end,
                    },
                    tag_end: data.tag_end,
                    close_start,
                    tag: data.tag,
                    rows: data.rows.into_boxed_slice(),
                    empty: false,
                });
            },
            _ => {},
        }
        Ok(())
    }

    fn start_row(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let (number, last_column) = self.row_position(element, decoder)?;
        self.row = Some(PendingRow {
            number,
            last_column,
            start,
            tag_end: end,
            tag: tag(element, decoder)?,
            descent_attribute: x14ac::attribute_name(element, resolver)?,
            cells: Vec::new(),
        });
        Ok(())
    }

    fn row_position(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<(u32, u32)> {
        let number = match unqualified_attribute_value(element, b"r", decoder)? {
            Some(value) => parse_one_based_row(&value)?,
            None => self
                .previous_row
                .checked_add(1)
                .filter(|number| *number <= ROWS)
                .ok_or_else(|| invalid("inferred edit row exceeds the grid"))?,
        };
        if self.previous_row != 0 && number <= self.previous_row {
            return Err(invalid(
                "cell edits require strictly increasing worksheet rows",
            ));
        }
        self.previous_row = number;
        Ok((number, 0))
    }

    fn start_cell(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let address = self.cell_address(element, decoder)?;
        self.cell = Some(PendingCell {
            address,
            start,
            tag_end: end,
            tag: tag(element, decoder)?,
            primary: Vec::new(),
            mce_payload: false,
        });
        Ok(())
    }

    fn cell_address(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<Address> {
        let row = self
            .row
            .as_ref()
            .ok_or_else(|| invalid("cell outside edit row"))?
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
                .ok_or_else(|| invalid("inferred edit column exceeds the grid"))?,
        };
        let pending = self
            .row
            .as_mut()
            .ok_or_else(|| invalid("cell outside edit row"))?;
        pending.last_column = column;
        Address::at(row - 1, column - 1).map_err(Into::into)
    }

    fn scan_formula(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let kind = unqualified_attribute_value(element, b"t", decoder)?
            .unwrap_or_else(|| "normal".to_owned());
        if !matches!(kind.as_str(), "shared" | "array" | "dataTable") {
            return Ok(());
        }
        let address = self
            .cell
            .as_ref()
            .ok_or_else(|| invalid("formula outside edit cell"))?
            .address;
        let range = unqualified_attribute_value(element, b"ref", decoder)?
            .map(|value| SelectionRange::cell_or_area(&value))
            .transpose()?;
        let index = if kind == "shared" {
            optional_u32(element, b"si", decoder, "shared formula index")?
        } else {
            None
        };
        self.formulas.push(FormulaStorage {
            address,
            kind: kind.into_boxed_str(),
            index,
            range,
        });
        Ok(())
    }

    fn ensure_merge_cells_slot(&self) -> Result<()> {
        if self.pending_merge_cells.is_some() || self.merge_cells.is_some() {
            return Err(invalid("worksheet has duplicate mergeCells during edit"));
        }
        if self.sheet_data.is_none() {
            return Err(invalid(
                "worksheet mergeCells appears before sheetData during edit",
            ));
        }
        if self.merge_insertion.is_some() {
            return Err(invalid(
                "worksheet mergeCells appears after a schema successor during edit",
            ));
        }
        Ok(())
    }

    fn start_merge_cells(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        start: usize,
        tag_end: usize,
    ) -> Result<()> {
        self.ensure_merge_cells_slot()?;
        let count = optional_u32(element, b"count", decoder, "worksheet merged-range count")?
            .map(usize::try_from)
            .transpose()
            .map_err(|_| invalid("worksheet merged-range count does not fit usize during edit"))?;
        self.pending_merge_cells = Some(PendingMergeCells {
            start,
            tag_end,
            tag: tag(element, decoder)?,
            count,
            merges: Vec::new(),
            payload: false,
        });
        Ok(())
    }

    fn mark_merge_payload(&mut self) {
        if let Some(merges) = self.pending_merge_cells.as_mut() {
            merges.payload = true;
        } else if let Some(merges) = self.merge_cells.as_mut() {
            merges.payload = true;
        }
    }

    fn observe_merge_position(
        &mut self,
        parent: Option<FrameKind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        start: usize,
    ) {
        if parent != Some(FrameKind::Worksheet) || self.sheet_data.is_none() {
            return;
        }
        let local_name = element.name().local_name();
        let local = local_name.as_ref();
        if is_spreadsheetml_name(namespace, element.name(), local) {
            if merge_successor(local) {
                self.merge_insertion.get_or_insert(start);
            } else if !merge_predecessor(local) && local != b"mergeCells" {
                self.merge_compatibility = true;
            }
        } else {
            self.merge_compatibility = true;
        }
    }

    fn scan_guard(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        if is_spreadsheetml_name(namespace, element.name(), b"sheetProtection") {
            self.protected |= optional_bool(element, b"sheet", decoder, "sheet protection flag")?
                .unwrap_or(false);
        }
        if is_spreadsheetml_name(namespace, element.name(), b"dataValidation") {
            let value = unqualified_attribute_value(element, b"sqref", decoder)?
                .ok_or_else(|| invalid("dataValidation is missing sqref during edit"))?;
            for token in value.split_whitespace() {
                self.validations.push(SelectionRange::selection(token)?);
            }
        }
        if element.name().local_name().as_ref() == b"dataValidation"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == X14)
        {
            self.extended_validation = true;
        }
        Ok(())
    }

    fn record_dimension(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        span: Span,
        empty: bool,
    ) -> Result<()> {
        if self.dimension.is_some() {
            return Err(invalid(
                "worksheet has duplicate dimension elements during edit",
            ));
        }
        let reference = unqualified_attribute_value(element, b"ref", decoder)?
            .ok_or_else(|| invalid("worksheet dimension is missing ref during edit"))?;
        let declared = Rect::from_a1(&reference).map_err(|error| {
            invalid(format!(
                "invalid worksheet dimension '{reference}' during edit: {error}"
            ))
        })?;
        self.dimension = Some(DimensionTag {
            span,
            tag: tag(element, decoder)?,
            empty,
            declared,
        });
        Ok(())
    }

    fn finish_layout(self) -> Result<Layout> {
        let root = self
            .root
            .ok_or_else(|| invalid("worksheet edit scan requires a worksheet root"))?;
        let sheet_data = self
            .sheet_data
            .ok_or_else(|| invalid("worksheet cell edits require a direct sheetData element"))?;
        let merge_insertion = self
            .merge_insertion
            .or(self.root_close_start)
            .ok_or_else(|| invalid("worksheet edit scan did not find the root closing tag"))?;
        if self
            .defaults
            .as_ref()
            .is_some_and(|defaults| defaults.span.start >= sheet_data.span.start)
        {
            return Err(invalid(
                "worksheet sheetFormatPr must precede sheetData during edit",
            ));
        }
        if let (Some(defaults), Some(columns)) = (&self.defaults, &self.columns)
            && defaults.span.start >= columns.span.start
        {
            return Err(invalid(
                "worksheet sheetFormatPr must precede cols during edit",
            ));
        }
        if let (Some(dimension), Some(defaults)) = (&self.dimension, &self.defaults)
            && dimension.span.start >= defaults.span.start
        {
            return Err(invalid(
                "worksheet dimension must precede sheetFormatPr during edit",
            ));
        }
        if self
            .columns
            .as_ref()
            .is_some_and(|columns| columns.columns.is_empty())
        {
            return Err(invalid("worksheet cols contains no col during edit"));
        }
        if self
            .columns
            .as_ref()
            .is_some_and(|columns| columns.span.start >= sheet_data.span.start)
        {
            return Err(invalid("worksheet cols must precede sheetData during edit"));
        }
        if self
            .dimension
            .as_ref()
            .is_some_and(|dimension| dimension.span.start >= sheet_data.span.start)
        {
            return Err(invalid(
                "worksheet dimension must precede sheetData during cell edits",
            ));
        }
        let mut formula_ranges = Vec::new();
        let mut shared = HashMap::<u32, SelectionRange>::new();
        for formula in &self.formulas {
            match formula.kind.as_ref() {
                "array" | "dataTable" => {
                    formula_ranges.push(formula.range.unwrap_or(SelectionRange {
                        first_row: formula.address.row().get(),
                        first_column: formula.address.column().get(),
                        last_row: formula.address.row().get(),
                        last_column: formula.address.column().get(),
                    }))
                },
                "shared" => {
                    if let (Some(index), Some(range)) = (formula.index, formula.range) {
                        shared.insert(index, range);
                    }
                },
                _ => {},
            }
        }
        for formula in &self.formulas {
            if formula.kind.as_ref() == "shared" {
                formula_ranges.push(
                    formula
                        .index
                        .and_then(|index| shared.get(&index).copied())
                        .unwrap_or(SelectionRange {
                            first_row: formula.address.row().get(),
                            first_column: formula.address.column().get(),
                            last_row: formula.address.row().get(),
                            last_column: formula.address.column().get(),
                        }),
                );
            }
        }
        let merge_count = self
            .merge_cells
            .as_ref()
            .map_or(0, |container| container.merges.len());
        let mut merged_ranges = Vec::new();
        merged_ranges
            .try_reserve_exact(merge_count)
            .map_err(|source| allocation("scanned merged ranges", source))?;
        if let Some(container) = self.merge_cells.as_ref() {
            merged_ranges.extend(container.merges.iter().map(|merge| merge.range));
        }
        let merged_ranges = merge::Index::new(merged_ranges)?;
        let mut merged = Vec::new();
        merged
            .try_reserve_exact(merged_ranges.as_slice().len())
            .map_err(|source| allocation("merge edit guards", source))?;
        merged.extend(
            merged_ranges
                .as_slice()
                .iter()
                .copied()
                .map(SelectionRange::from_rect),
        );
        Ok(Layout {
            root,
            defaults: self.defaults,
            sheet_data,
            columns: self.columns,
            dimension: self.dimension,
            protected: self.protected,
            merged: merged.into_boxed_slice(),
            validations: self.validations.into_boxed_slice(),
            extended_validation: self.extended_validation,
            formula_ranges: formula_ranges.into_boxed_slice(),
            defaults_compatibility: self.defaults_compatibility,
            merge_cells: self.merge_cells,
            merge_insertion,
            merge_compatibility: self.merge_compatibility,
        })
    }
}

pub(crate) fn write_root(output: &mut Vec<u8>, root: &RootSlot, effect: &RootEffect) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(root.tag.name.as_bytes());
    for attribute in &root.tag.attributes {
        if effect
            .removed
            .as_deref()
            .is_some_and(|name| name == attribute.name.as_ref())
        {
            continue;
        }
        write_attribute(output, &attribute.name, &attribute.value);
    }
    for (name, value) in &effect.appended {
        write_attribute(output, name, value);
    }
    output.extend_from_slice(b">");
}

pub(crate) fn write_defaults(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &DefaultsSlot,
    effects: DefaultsEffects,
    descent_name: &str,
) {
    let stored_descent = stored.descent_attribute.as_deref().unwrap_or(descent_name);
    let mut removed = Vec::new();
    let mut appended = Vec::new();
    defaults_effect_attributes(
        effects,
        stored_descent,
        descent_name,
        &mut removed,
        &mut appended,
    );
    write_tag(output, &stored.tag, stored.empty, &removed, &appended);
    if !stored.empty {
        output.extend_from_slice(&source[stored.tag_end..stored.close_start]);
        write_close(output, &stored.tag.name);
    }
}

pub(crate) fn write_new_defaults(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    effects: DefaultsEffects,
    descent_name: &str,
) {
    let name = sibling_name(sheet_data_name, "sheetFormatPr");
    let tag = Tag {
        name: name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut removed = Vec::new();
    let mut appended = Vec::new();
    defaults_effect_attributes(
        effects,
        descent_name,
        descent_name,
        &mut removed,
        &mut appended,
    );
    write_tag(output, &tag, true, &removed, &appended);
}

fn defaults_effect_attributes<'a>(
    effects: DefaultsEffects,
    stored_descent_name: &'a str,
    appended_descent_name: &'a str,
    removed: &mut Vec<&'a str>,
    appended: &mut Vec<(&'a str, String)>,
) {
    if let Some(effect) = effects.base_width {
        removed.push("baseColWidth");
        if let OptionalEffect::Set(value) = effect {
            appended.push(("baseColWidth", value.to_string()));
        }
    }
    if let Some(effect) = effects.width {
        removed.push("defaultColWidth");
        if let OptionalEffect::Set(value) = effect {
            appended.push(("defaultColWidth", value.get().to_string()));
        }
    }
    if let Some(height) = effects.height {
        removed.extend(["defaultRowHeight", "customHeight"]);
        appended.push(("defaultRowHeight", height.get().to_string()));
        appended.push(("customHeight", "1".to_owned()));
    }
    for (value, name) in [
        (effects.hidden, "zeroHeight"),
        (effects.thick_top, "thickTop"),
        (effects.thick_bottom, "thickBottom"),
    ] {
        if let Some(value) = value {
            removed.push(name);
            if value {
                appended.push((name, "1".to_owned()));
            }
        }
    }
    if let Some(effect) = effects.descent {
        removed.push(stored_descent_name);
        if let DescentEffect::Set(value) = effect {
            appended.push((appended_descent_name, value.get().to_string()));
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum ColumnPiece {
    Keep(Column, Column),
    Edit(Column, Column, ColumnAction),
}

pub(crate) fn write_columns(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &ColumnsSlot,
    actions: BTreeMap<Column, ColumnAction>,
    sheet: &str,
) -> Result<()> {
    let mut owners = Assignments::new()?;
    for (index, column) in stored.columns.iter().enumerate() {
        owners.assign(column.first, column.last, index);
    }
    let mut by_owner = HashMap::<usize, BTreeMap<Column, ColumnAction>>::new();
    let mut implicit = BTreeMap::new();
    for (column, action) in actions {
        if let Some(owner) = owners.get(column) {
            by_owner.entry(owner).or_default().insert(column, action);
        } else if action.materializes() {
            implicit.insert(column, action);
        }
    }

    if stored.payload
        && let Some(column) = implicit.keys().next()
    {
        return Err(Error::ColumnEditBlocked {
            sheet: sheet.to_owned(),
            column: *column,
            reason: ColumnEditBlock::MarkupCompatibility,
        });
    }

    if stored.empty {
        return Err(invalid("worksheet cols contains no col during edit"));
    }
    output.extend_from_slice(&source[stored.span.start..stored.tag_end]);
    let mut cursor = stored.tag_end;
    for (index, column) in stored.columns.iter().enumerate() {
        output.extend_from_slice(&source[cursor..column.span.start]);
        if let Some(edits) = by_owner.remove(&index) {
            let pieces = column_pieces(column, &edits)?;
            if column.payload && pieces.len() > 1 {
                let edited = edits.keys().next().copied().unwrap_or(column.first);
                return Err(Error::ColumnEditBlocked {
                    sheet: sheet.to_owned(),
                    column: edited,
                    reason: ColumnEditBlock::MarkupCompatibility,
                });
            }
            for piece in pieces {
                write_column_piece(output, source, column, piece);
            }
        } else {
            output.extend_from_slice(&source[column.span.start..column.span.end]);
        }
        cursor = column.span.end;
    }
    output.extend_from_slice(&source[cursor..stored.close_start]);
    write_column_actions(output, &stored.tag.name, implicit);
    output.extend_from_slice(&source[stored.close_start..stored.span.end]);
    Ok(())
}

fn column_pieces(
    stored: &ColumnSlot,
    edits: &BTreeMap<Column, ColumnAction>,
) -> Result<Vec<ColumnPiece>> {
    let capacity = edits
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or_else(|| invalid("column edit split count overflow"))?;
    let mut pieces = Vec::new();
    pieces
        .try_reserve_exact(capacity)
        .map_err(|source| allocation("column edit splits", source))?;
    let mut next = stored.first.get();
    for (column, action) in edits {
        if column.get() > next {
            pieces.push(ColumnPiece::Keep(
                Column::new(next)?,
                Column::new(column.get() - 1)?,
            ));
        }
        if let Some(ColumnPiece::Edit(_, last, previous)) = pieces.last_mut()
            && previous == action
            && last.next() == Some(*column)
        {
            *last = *column;
        } else {
            pieces.push(ColumnPiece::Edit(*column, *column, *action));
        }
        next = column.get().saturating_add(1);
    }
    if next <= stored.last.get() {
        pieces.push(ColumnPiece::Keep(Column::new(next)?, stored.last));
    }
    Ok(pieces)
}

fn write_column_piece(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &ColumnSlot,
    piece: ColumnPiece,
) {
    let (first, last, action) = match piece {
        ColumnPiece::Keep(first, last) => (first, last, None),
        ColumnPiece::Edit(first, last, action) => (first, last, Some(action)),
    };
    let mut removed = vec!["min", "max"];
    let mut appended = vec![
        ("min", (first.get() + 1).to_string()),
        ("max", (last.get() + 1).to_string()),
    ];
    if let Some(action) = action {
        column_effect_attributes(action, &mut removed, &mut appended);
    }
    write_tag(output, &stored.tag, stored.empty, &removed, &appended);
    if !stored.empty {
        output.extend_from_slice(&source[stored.tag_end..stored.close_start]);
        write_close(output, &stored.tag.name);
    }
}

pub(crate) fn write_new_columns(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    actions: BTreeMap<Column, ColumnAction>,
) {
    if !actions.values().any(|action| action.materializes()) {
        return;
    }
    let name = sibling_name(sheet_data_name, "cols");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    write_tag(output, &tag, false, &[], &[]);
    write_column_actions(output, &name, actions);
    write_close(output, &name);
}

fn write_column_actions(
    output: &mut Vec<u8>,
    columns_name: &str,
    actions: BTreeMap<Column, ColumnAction>,
) {
    let name = sibling_name(columns_name, "col");
    let tag = Tag {
        name: name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut pending: Option<(Column, Column, ColumnAction)> = None;
    for (column, action) in actions {
        if !action.materializes() {
            continue;
        }
        match pending {
            Some((first, last, previous)) if previous == action && last.next() == Some(column) => {
                pending = Some((first, column, action));
            },
            Some((first, last, previous)) => {
                write_new_column(output, &tag, first, last, previous);
                pending = Some((column, column, action));
            },
            None => pending = Some((column, column, action)),
        }
    }
    if let Some((first, last, action)) = pending {
        write_new_column(output, &tag, first, last, action);
    }
}

fn write_new_column(
    output: &mut Vec<u8>,
    tag: &Tag,
    first: Column,
    last: Column,
    action: ColumnAction,
) {
    let mut removed = Vec::new();
    let mut appended = vec![
        ("min", (first.get() + 1).to_string()),
        ("max", (last.get() + 1).to_string()),
    ];
    column_effect_attributes(action, &mut removed, &mut appended);
    write_tag(output, tag, true, &removed, &appended);
}

fn column_effect_attributes(
    action: ColumnAction,
    removed: &mut Vec<&'static str>,
    appended: &mut Vec<(&'static str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(width) = action.width {
        removed.extend(["width", "customWidth"]);
        if let WidthEffect::Set(width) = width {
            appended.push(("width", width.get().to_string()));
            appended.push(("customWidth", "1".to_owned()));
        }
    }
    if let Some(style) = action.style {
        removed.push("style");
        if let StyleEffect::Set(key) = style {
            appended.push(("style", key.to_string()));
        }
    }
    if let Some(best_fit) = action.best_fit {
        removed.push("bestFit");
        if best_fit {
            appended.push(("bestFit", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    if let Some(collapsed) = action.collapsed {
        removed.push("collapsed");
        if collapsed {
            appended.push(("collapsed", "1".to_owned()));
        }
    }
    if let Some(phonetic) = action.phonetic {
        removed.push("phonetic");
        if phonetic {
            appended.push(("phonetic", "1".to_owned()));
        }
    }
}

pub(crate) fn write_sheet_data(
    output: &mut Vec<u8>,
    source: &[u8],
    data: &SheetData,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<Row, RowAction>,
    descent_name: &str,
) -> Result<()> {
    let mut by_row = BTreeMap::<u32, RowEdits>::new();
    for (address, action) in cells {
        by_row
            .entry(address.row().get() + 1)
            .or_default()
            .cells
            .insert(address, action);
    }
    for (row, action) in rows {
        by_row.entry(row.get() + 1).or_default().row = Some(action);
    }

    if data.empty {
        write_tag(output, &data.tag, false, &[], &[]);
        for (number, edits) in by_row {
            write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
        }
        write_close(output, &data.tag.name);
        return Ok(());
    }

    output.extend_from_slice(&source[data.span.start..data.tag_end]);
    let mut cursor = data.tag_end;
    let mut pending = by_row.into_iter().peekable();
    for row in &data.rows {
        output.extend_from_slice(&source[cursor..row.span.start]);
        while pending
            .peek()
            .is_some_and(|(number, _)| *number < row.number)
        {
            if let Some((number, edits)) = pending.next() {
                write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
            }
        }
        if pending
            .peek()
            .is_some_and(|(number, _)| *number == row.number)
        {
            let (_, edits) = pending
                .next()
                .ok_or_else(|| invalid("worksheet row edit ordering was lost"))?;
            write_row(output, source, row, &edits, descent_name)?;
        } else {
            output.extend_from_slice(&source[row.span.start..row.span.end]);
        }
        cursor = row.span.end;
    }
    output.extend_from_slice(&source[cursor..data.close_start]);
    for (number, edits) in pending {
        write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
    }
    output.extend_from_slice(&source[data.close_start..data.span.end]);
    Ok(())
}

#[derive(Debug, Default)]
struct RowEdits {
    cells: BTreeMap<Address, Action>,
    row: Option<RowAction>,
}

fn write_row(
    output: &mut Vec<u8>,
    source: &[u8],
    row: &RowSlot,
    edits: &RowEdits,
    descent_name: &str,
) -> Result<()> {
    let actions = &edits.cells;
    let membership_changed = actions.iter().any(|(address, action)| {
        let exists = row
            .cells
            .binary_search_by_key(address, |cell| cell.address)
            .is_ok();
        (!exists && action.creates_missing()) || (exists && matches!(action, Action::Remove))
    });

    if row.empty {
        let creates_cell = actions.values().any(Action::creates_missing);
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if creates_cell {
            removed.extend(["spans", "r"]);
            appended.push(("r", row.number.to_string()));
        }
        if let Some(action) = edits.row {
            row_effect_attributes(
                action,
                row.descent_attribute.as_deref().unwrap_or(descent_name),
                &mut removed,
                &mut appended,
            );
        }
        write_tag(output, &row.tag, !creates_cell, &removed, &appended);
        if !creates_cell {
            return Ok(());
        }
        for (address, action) in actions {
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        write_close(output, &row.tag.name);
        return Ok(());
    }

    if membership_changed || edits.row.is_some() {
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if membership_changed {
            removed.push("spans");
        }
        if let Some(action) = edits.row {
            row_effect_attributes(
                action,
                row.descent_attribute.as_deref().unwrap_or(descent_name),
                &mut removed,
                &mut appended,
            );
        }
        write_tag(output, &row.tag, false, &removed, &appended);
    } else {
        output.extend_from_slice(&source[row.span.start..row.tag_end]);
    }
    let mut cursor = row.tag_end;
    let mut pending = actions.iter().peekable();
    for cell in &row.cells {
        output.extend_from_slice(&source[cursor..cell.span.start]);
        while pending
            .peek()
            .is_some_and(|(address, _)| **address < cell.address)
        {
            let (address, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        if pending
            .peek()
            .is_some_and(|(address, _)| **address == cell.address)
        {
            let (_, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            match action {
                Action::Update { .. } => write_cell(output, source, cell, action)?,
                Action::Remove => {},
            }
        } else {
            output.extend_from_slice(&source[cell.span.start..cell.span.end]);
        }
        cursor = cell.span.end;
    }
    output.extend_from_slice(&source[cursor..row.close_start]);
    for (address, action) in pending {
        write_new_action(output, &row.tag.name, *address, action)?;
    }
    output.extend_from_slice(&source[row.close_start..row.span.end]);
    Ok(())
}

fn write_new_row(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    number: u32,
    edits: &RowEdits,
    descent_name: &str,
) -> Result<()> {
    let creates_cell = edits.cells.values().any(Action::creates_missing);
    let materializes = edits.row.is_some_and(RowAction::materializes);
    if !creates_cell && !materializes {
        return Ok(());
    }
    let name = sibling_name(sheet_data_name, "row");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", number.to_string())];
    let mut removed = Vec::new();
    if let Some(action) = edits.row {
        row_effect_attributes(action, descent_name, &mut removed, &mut appended);
    }
    write_tag(output, &tag, !creates_cell, &removed, &appended);
    if !creates_cell {
        return Ok(());
    }
    for (address, action) in &edits.cells {
        write_new_action(output, &name, *address, action)?;
    }
    write_close(output, &name);
    Ok(())
}

fn row_effect_attributes<'a>(
    action: RowAction,
    descent_name: &'a str,
    removed: &mut Vec<&'a str>,
    appended: &mut Vec<(&'a str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(height) = action.height {
        removed.extend(["ht", "customHeight"]);
        if let HeightEffect::Set(height) = height {
            appended.push(("ht", height.get().to_string()));
            appended.push(("customHeight", "1".to_owned()));
        }
    }
    if let Some(descent) = action.descent {
        removed.push(descent_name);
        if let DescentEffect::Set(value) = descent {
            appended.push((descent_name, value.get().to_string()));
        }
    }
    if let Some(style) = action.style {
        removed.extend(["s", "customFormat"]);
        if let StyleEffect::Set(key) = style {
            appended.push(("s", key.to_string()));
            appended.push(("customFormat", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    for (value, name) in [
        (action.collapsed, "collapsed"),
        (action.thick_top, "thickTop"),
        (action.thick_bottom, "thickBot"),
        (action.phonetic, "ph"),
    ] {
        if let Some(value) = value {
            removed.push(name);
            if value {
                appended.push((name, "1".to_owned()));
            }
        }
    }
}

fn write_cell(output: &mut Vec<u8>, source: &[u8], cell: &CellSlot, action: &Action) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Err(invalid("cannot rewrite a removed cell"));
    };
    let content = match payload.as_ref() {
        Some(Payload::Set(content)) => Some(content),
        Some(Payload::Clear | Payload::ClearIfPresent) | None => None,
    };
    let cell_type = content.and_then(content_type);
    let mut removed = vec!["r"];
    if payload.is_some() {
        removed.push("t");
    }
    if style.is_some() {
        removed.push("s");
    }
    let mut appended = vec![("r", cell.address.a1())];
    if let Some(cell_type) = cell_type {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let remains_empty = cell.empty && payload.is_none();
    write_tag(output, &cell.tag, remains_empty, &removed, &appended);
    if remains_empty {
        return Ok(());
    }
    if let Some(content) = content {
        write_content(output, &cell.tag.name, content)?;
    }
    if !cell.empty {
        if payload.is_some() {
            copy_without(
                output,
                source,
                cell.tag_end,
                cell.close_start,
                &cell.primary,
            );
        } else {
            output.extend_from_slice(&source[cell.tag_end..cell.close_start]);
        }
    }
    write_close(output, &cell.tag.name);
    Ok(())
}

fn write_new_action(
    output: &mut Vec<u8>,
    row_name: &str,
    address: Address,
    action: &Action,
) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Ok(());
    };
    if !action.creates_missing() {
        return Ok(());
    }
    let content = match payload.as_ref() {
        Some(Payload::Set(content)) => Some(content),
        Some(Payload::Clear | Payload::ClearIfPresent) | None => None,
    };
    let name = sibling_name(row_name, "c");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", address.a1())];
    if let Some(cell_type) = content.and_then(content_type) {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let empty = content.is_none();
    write_tag(output, &tag, empty, &[], &appended);
    if let Some(content) = content {
        write_content(output, &name, content)?;
        write_close(output, &name);
    }
    Ok(())
}

fn content_type(content: &Content) -> Option<&'static str> {
    match content {
        Content::Value(Value::Bool(_)) => Some("b"),
        Content::Value(Value::Text(_)) => Some("inlineStr"),
        Content::Value(Value::Date(_)) => Some("d"),
        Content::Value(Value::Error(_)) => Some("e"),
        Content::Value(Value::Number(_)) | Content::Formula(_) => None,
    }
}

fn write_content(output: &mut Vec<u8>, cell_name: &str, content: &Content) -> Result<()> {
    match content {
        Content::Value(Value::Bool(value)) => {
            write_text_element(output, cell_name, "v", if *value { "1" } else { "0" });
        },
        Content::Value(Value::Number(value)) => {
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Value(Value::Text(value)) => {
            let inline = sibling_name(cell_name, "is");
            let text = sibling_name(cell_name, "t");
            output.extend_from_slice(b"<");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b"><");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b" xml:space=\"preserve\">");
            output.extend_from_slice(escape_xml(&encode_spreadsheet_text(value)).as_bytes());
            output.extend_from_slice(b"</");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b"></");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b">");
        },
        Content::Value(Value::Date(value)) => {
            require_xml_text(value)?;
            write_text_element(output, cell_name, "v", &escape_xml(value));
        },
        Content::Value(Value::Error(value)) => {
            require_xml_text(value.as_str())?;
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Formula(formula) => {
            require_xml_text(formula.text())?;
            write_text_element(output, cell_name, "f", &escape_xml(formula.text()));
        },
    }
    Ok(())
}

fn write_text_element(output: &mut Vec<u8>, cell_name: &str, local: &str, value: &str) {
    let name = sibling_name(cell_name, local);
    output.extend_from_slice(b"<");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
    output.extend_from_slice(value.as_bytes());
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn require_xml_text(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
            || ('\u{20}'..='\u{D7FF}').contains(&character)
            || ('\u{E000}'..='\u{FFFD}').contains(&character)
            || ('\u{10000}'..='\u{10FFFF}').contains(&character)
    }) {
        Ok(())
    } else {
        Err(invalid(
            "cell content contains a character forbidden by XML 1.0",
        ))
    }
}

fn copy_without(output: &mut Vec<u8>, source: &[u8], start: usize, end: usize, removed: &[Span]) {
    let mut cursor = start;
    for span in removed {
        output.extend_from_slice(&source[cursor..span.start]);
        cursor = span.end;
    }
    output.extend_from_slice(&source[cursor..end]);
}
