//! Streaming `SpreadsheetML` scanner for worksheet snapshots.

use std::collections::HashMap;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::{COLUMNS, Cell as Address, Column, ROWS, Rect};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::X14;
use super::super::wire::{column_range, is_mce_name, position, tag};
use super::model::{
    CellSlot, ColumnSlot, ColumnsSlot, DefaultsSlot, DimensionTag, Layout, MergeCellsSlot,
    MergeSlot, RootSlot, RowSlot, SheetData, Span, Tag,
};
use crate::error::{Result, invalid};
use crate::raw::namespace::is_spreadsheetml_name;
use crate::raw::worksheet::edit::model::SelectionRange;
use crate::raw::worksheet::model::{MAX_XML_DEPTH, MAX_XML_EVENTS};
use crate::raw::worksheet::{
    merge_successor, optional_bool, optional_u32, parse_a1, parse_one_based_row, x14ac,
};
use crate::{error::allocation, merge};

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

pub(crate) fn scan(content: &[u8]) -> Result<Layout> {
    scan_with_limit(content, MAX_XML_EVENTS)
}

#[cfg(test)]
pub(crate) fn scan_with_event_limit(content: &[u8], max_events: usize) -> Result<Layout> {
    scan_with_limit(content, max_events)
}

fn scan_with_limit(content: &[u8], max_events: usize) -> Result<Layout> {
    let mut reader = NsReader::from_reader(content);
    reader.config_mut().check_end_names = true;
    let mut scanner = Scanner::default();
    let mut stack = Vec::<Frame>::new();
    let mut events = 0usize;
<<<<<<< HEAD
    let mut buffer = Vec::new();
=======
>>>>>>> agent/0232-xlsx-xml-integration

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML event count overflow"))?;
        if events > max_events {
            return Err(invalid("worksheet XML exceeds event limit"));
        }
        let event_start = position(&reader)?;
<<<<<<< HEAD
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| invalid(error.to_string()))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
=======
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(error.to_string()))?;
>>>>>>> agent/0232-xlsx-xml-integration
        let event_end = position(&reader)?;
        let decoder = reader.decoder();
        let resolver = reader.resolver();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(invalid(format!(
                        "worksheet XML exceeds {MAX_XML_DEPTH} levels"
                    )));
                }
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
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
        }
        buffer.clear();
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
            FrameKind::Other => {},
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
            .map_err(|_source| {
                invalid("worksheet merged-range count does not fit usize during edit")
            })?;
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
                    }));
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
