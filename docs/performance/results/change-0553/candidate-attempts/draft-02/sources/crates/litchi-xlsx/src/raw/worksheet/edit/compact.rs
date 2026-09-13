//! Commit-local source spans for the bounded scalar worksheet writer.
//!
//! This walker deliberately proves only the small shape consumed by the
//! compact writer.  It never materializes XML tags, values, formulas, or a
//! second semantic store.  Any namespace, structure, attribute, coordinate,
//! or resource-limit uncertainty returns `None`; the caller then uses the
//! complete provenance path.

use std::mem::size_of;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use litchi_sheet::{COLUMNS, Cell as Address, ROWS, Rect};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{
    Action, CompactCellSlot, CompactDimensionTag, CompactLayout, CompactRowSlot, CompactSheetData,
    CompactSpan, Payload,
};
use crate::cell::{Cell, Stored};
use crate::raw::namespace::is_spreadsheetml_name;
use crate::raw::worksheet::{
    MAX_SHARED_PROVISIONAL_EVENTS, MAX_SHARED_SOURCE_BYTES, parse_a1, parse_one_based_row,
    shared_event_bound_within_cap, source_stream_eligible,
};

const MAX_PROOF_BYTES: usize = 2 * 1024 * 1024;
const STACK_GROWTH: usize = 16;
const ROW_GROWTH: usize = 16;
const CELL_GROWTH: usize = 16;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Worksheet,
    Dimension,
    Defaults,
    Columns,
    Column,
    SheetData,
    Row,
    Cell,
    Value,
    Inline,
    InlineRun,
    InlineText,
}

#[derive(Debug)]
struct PendingSheetData {
    start: u32,
    tag_end: u32,
    rows: Vec<CompactRowSlot>,
}

#[derive(Debug)]
struct PendingRow {
    number: u32,
    start: u32,
    tag_end: u32,
    close_start: u32,
    last_column: u32,
    cells: Vec<CompactCellSlot>,
}

#[derive(Debug)]
struct PendingCell {
    start: u32,
    primary: Option<FrameKind>,
}

#[derive(Debug)]
struct Collector<'a> {
    source: &'a [u8],
    entries: &'a [Stored],
    stack: Vec<FrameKind>,
    sheet_data: Option<PendingSheetData>,
    finished_sheet_data: Option<CompactSheetData>,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    dimension: Option<CompactDimensionTag>,
    previous_row: Option<u32>,
    entry_index: usize,
    cell_count: u32,
    seen_dimension: bool,
    seen_defaults: bool,
    seen_columns: bool,
    columns_count: usize,
    last_column_end: u32,
    seen_sheet_data: bool,
    closed_root: bool,
    accounting: usize,
}

/// Collect the exact source spans required by the existing-cell scalar
/// writer.  The returned layout is tied to both input slices by pointer and
/// length and is intended to be dropped at the end of the commit rewrite.
pub(crate) fn collect_compact_layout(
    source: &[u8],
    entries: &[Stored],
    actions: &std::collections::BTreeMap<Address, Action>,
) -> Option<CompactLayout> {
    if actions.is_empty()
        || source.len() > MAX_SHARED_SOURCE_BYTES
        || !source_stream_eligible(source)
        || !shared_event_bound_within_cap(source)
        || source.len() > u32::MAX as usize
        || !eligible_entries(entries)
        || !eligible_actions(entries, actions)
    {
        return None;
    }

    let accounting = size_of::<Collector<'static>>()
        .checked_add(size_of::<CompactLayout>())?
        .checked_add(size_of::<NsReader<&'static [u8]>>())?;
    if accounting > MAX_PROOF_BYTES {
        return None;
    }
    Collector {
        source,
        entries,
        stack: Vec::new(),
        sheet_data: None,
        finished_sheet_data: None,
        row: None,
        cell: None,
        dimension: None,
        previous_row: None,
        entry_index: 0,
        cell_count: 0,
        seen_dimension: false,
        seen_defaults: false,
        seen_columns: false,
        columns_count: 0,
        last_column_end: 0,
        seen_sheet_data: false,
        closed_root: false,
        accounting,
    }
    .walk()
}

fn eligible_entries(entries: &[Stored]) -> bool {
    entries.iter().all(|entry| {
        entry.shared_string.is_none()
            && !entry.inline_rich
            && entry.formula_range.is_none()
            && entry.shared_formula.is_none()
            && entry.cell_metadata.is_none()
            && entry.value_metadata.is_none()
            && matches!(&entry.cell, Cell::Empty | Cell::Value(_))
    })
}

fn eligible_actions(
    entries: &[Stored],
    actions: &std::collections::BTreeMap<Address, Action>,
) -> bool {
    actions.iter().all(|(address, action)| {
        entries
            .binary_search_by_key(address, |entry| entry.address)
            .is_ok()
            && matches!(
                action,
                Action::Update {
                    payload: None
                        | Some(Payload::Set(_) | Payload::Clear | Payload::ClearIfPresent),
                    ..
                }
            )
    })
}

impl<'a> Collector<'a> {
    fn walk(mut self) -> Option<CompactLayout> {
        let mut reader = NsReader::from_reader(self.source);
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let mut events = 0usize;

        loop {
            events = events.checked_add(1)?;
            if events > MAX_SHARED_PROVISIONAL_EVENTS {
                return None;
            }
            let start = usize::try_from(reader.buffer_position()).ok()?;
            let event = reader.read_event().ok()?;
            let end = usize::try_from(reader.buffer_position()).ok()?;
            let span = compact_span(start, end)?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            let decoder = reader.decoder();
            let eof = matches!(event, Event::Eof);
            if !self.observe(&namespace, event, span, decoder) {
                return None;
            }
            if eof {
                break;
            }
        }

        self.finish()
    }

    fn observe(
        &mut self,
        namespace: &ResolveResult<'_>,
        event: Event<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        match event {
            Event::Start(element) => self.start(namespace, &element, span, decoder),
            Event::Empty(element) => self.empty(namespace, &element, span, decoder),
            Event::End(element) => self.end(namespace, &element, span),
            Event::Text(value) => {
                self.text(value.decode().ok().map(|value| value.trim().is_empty()))
            },
            Event::CData(value) => {
                self.text(value.decode().ok().map(|value| value.trim().is_empty()))
            },
            // Character references are decoded by the complete scanner.  A
            // direct walker cannot retain their error precedence without
            // retaining value state, so it conservatively declines them.
            Event::GeneralRef(_) | Event::DocType(_) => false,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => true,
            Event::Eof => self.closed_root && self.stack.is_empty(),
        }
    }

    fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        if !decode_attributes(element, decoder) {
            return false;
        }
        if self.stack.len() >= 32 || !self.reserve_stack() {
            return false;
        }
        let local = element.name().local_name();
        if self.stack.is_empty() {
            if self.closed_root || !is_spreadsheetml_name(namespace, element.name(), b"worksheet") {
                return false;
            }
            self.stack.push(FrameKind::Worksheet);
            return true;
        }
        let parent = *self.stack.last().unwrap_or(&FrameKind::Worksheet);
        if !is_core(namespace) {
            return false;
        }
        match parent {
            FrameKind::Worksheet => self.start_worksheet(local.as_ref(), element, span, decoder),
            FrameKind::SheetData if local.as_ref() == b"row" => {
                self.start_row(element, span, decoder)
            },
            FrameKind::Row if local.as_ref() == b"c" => self.start_cell(element, span, decoder),
            FrameKind::Cell => self.start_cell_child(local.as_ref()),
            FrameKind::Inline if local.as_ref() == b"t" => {
                self.stack.push(FrameKind::InlineText);
                true
            },
            FrameKind::Inline if local.as_ref() == b"r" => {
                self.stack.push(FrameKind::InlineRun);
                true
            },
            FrameKind::InlineRun if local.as_ref() == b"t" => {
                self.stack.push(FrameKind::InlineText);
                true
            },
            FrameKind::Columns if local.as_ref() == b"col" => {
                let Some((first, last)) = parse_column_range(element, decoder) else {
                    return false;
                };
                self.record_column(first, last) && self.push_frame(FrameKind::Column)
            },
            _ => false,
        }
    }

    fn start_worksheet(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        match local {
            b"dimension"
                if !self.seen_dimension
                    && !self.seen_defaults
                    && !self.seen_columns
                    && !self.seen_sheet_data =>
            {
                let Some(reference) = attribute(element, b"ref", decoder) else {
                    return false;
                };
                let Ok(declared) = Rect::from_a1(&reference) else {
                    return false;
                };
                self.dimension = Some(CompactDimensionTag {
                    span,
                    empty: false,
                    declared,
                });
                self.seen_dimension = true;
                self.stack.push(FrameKind::Dimension);
                true
            },
            b"sheetFormatPr"
                if !self.seen_defaults && !self.seen_columns && !self.seen_sheet_data =>
            {
                self.seen_defaults = true;
                self.stack.push(FrameKind::Defaults);
                true
            },
            b"cols" if !self.seen_columns && !self.seen_sheet_data && !self.seen_defaults => {
                self.seen_columns = true;
                self.columns_count = 0;
                self.last_column_end = 0;
                self.stack.push(FrameKind::Columns);
                true
            },
            b"sheetData" if !self.seen_sheet_data => {
                self.seen_sheet_data = true;
                self.sheet_data = Some(PendingSheetData {
                    start: span.start,
                    tag_end: span.end,
                    rows: Vec::new(),
                });
                self.push_frame(FrameKind::SheetData)
            },
            _ => false,
        }
    }

    fn start_row(&mut self, element: &BytesStart<'_>, span: CompactSpan, decoder: Decoder) -> bool {
        let Some(number) = inferred_row(element, decoder, self.previous_row) else {
            return false;
        };
        if self.previous_row.is_some_and(|previous| number <= previous) {
            return false;
        }
        self.previous_row = Some(number);
        self.row = Some(PendingRow {
            number,
            start: span.start,
            tag_end: span.end,
            close_start: span.end,
            last_column: 0,
            cells: Vec::new(),
        });
        self.push_frame(FrameKind::Row)
    }

    fn start_cell(
        &mut self,
        element: &BytesStart<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        let Some((row_number, previous_column)) =
            self.row.as_ref().map(|row| (row.number, row.last_column))
        else {
            return false;
        };
        let mut next_column = previous_column;
        let Some(address) = inferred_cell(element, row_number, &mut next_column, decoder) else {
            return false;
        };
        if self
            .entries
            .get(self.entry_index)
            .map(|entry| entry.address)
            != Some(address)
        {
            return false;
        }
        let Some(entry_index) = self.entry_index.checked_add(1) else {
            return false;
        };
        let Some(cell_count) = self.cell_count.checked_add(1) else {
            return false;
        };
        self.entry_index = entry_index;
        self.cell_count = cell_count;
        let Some(row) = self.row.as_mut() else {
            return false;
        };
        row.last_column = next_column;
        self.cell = Some(PendingCell {
            start: span.start,
            primary: None,
        });
        self.push_frame(FrameKind::Cell)
    }

    fn start_cell_child(&mut self, local: &[u8]) -> bool {
        if self.cell.is_none()
            || self
                .cell
                .as_ref()
                .is_some_and(|cell| cell.primary.is_some())
        {
            return false;
        }
        let kind = match local {
            b"v" => FrameKind::Value,
            b"is" => FrameKind::Inline,
            // Source formulas, including empty/shared/array forms, retain
            // semantic ownership that the scalar writer cannot prove.
            b"f" => return false,
            _ => return false,
        };
        if let Some(cell) = self.cell.as_mut() {
            cell.primary = Some(kind);
        }
        self.push_frame(kind)
    }

    fn empty(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        if !decode_attributes(element, decoder) {
            return false;
        }
        let local = element.name().local_name();
        if self.stack.is_empty() {
            return false;
        }
        if !is_core(namespace) {
            return false;
        }
        let parent = *self.stack.last().unwrap_or(&FrameKind::Worksheet);
        match parent {
            FrameKind::Worksheet => self.empty_worksheet(local.as_ref(), element, span, decoder),
            FrameKind::SheetData if local.as_ref() == b"row" => {
                self.empty_row(element, span, decoder)
            },
            FrameKind::Row if local.as_ref() == b"c" => self.empty_cell(element, span, decoder),
            FrameKind::Cell if matches!(local.as_ref(), b"v" | b"is") => {
                self.empty_cell_child(local.as_ref())
            },
            FrameKind::Inline if local.as_ref() == b"t" => true,
            FrameKind::InlineRun if local.as_ref() == b"t" => true,
            FrameKind::Columns if local.as_ref() == b"col" => {
                let Some((first, last)) = parse_column_range(element, decoder) else {
                    return false;
                };
                self.record_column(first, last)
            },
            _ => false,
        }
    }

    fn empty_worksheet(
        &mut self,
        local: &[u8],
        element: &BytesStart<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        match local {
            b"dimension"
                if !self.seen_dimension
                    && !self.seen_defaults
                    && !self.seen_columns
                    && !self.seen_sheet_data =>
            {
                let Some(reference) = attribute(element, b"ref", decoder) else {
                    return false;
                };
                let Ok(declared) = Rect::from_a1(&reference) else {
                    return false;
                };
                self.dimension = Some(CompactDimensionTag {
                    span,
                    empty: true,
                    declared,
                });
                self.seen_dimension = true;
                true
            },
            b"sheetFormatPr"
                if !self.seen_defaults && !self.seen_columns && !self.seen_sheet_data =>
            {
                self.seen_defaults = true;
                true
            },
            b"cols" if !self.seen_columns && !self.seen_sheet_data && !self.seen_defaults => {
                self.seen_columns = true;
                // The authoritative scanner rejects an empty `<cols>`
                // container after its structural pass.
                false
            },
            b"sheetData" if !self.seen_sheet_data => {
                self.seen_sheet_data = true;
                self.finished_sheet_data = Some(CompactSheetData {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    rows: Box::new([]),
                    empty: true,
                });
                true
            },
            _ => false,
        }
    }

    fn empty_row(&mut self, element: &BytesStart<'_>, span: CompactSpan, decoder: Decoder) -> bool {
        let Some(number) = inferred_row(element, decoder, self.previous_row) else {
            return false;
        };
        if self.previous_row.is_some_and(|previous| number <= previous) {
            return false;
        }
        self.previous_row = Some(number);
        if !self.reserve_rows() {
            return false;
        }
        let Some(sheet_data) = self.sheet_data.as_mut() else {
            return false;
        };
        sheet_data.rows.push(CompactRowSlot {
            number,
            span,
            tag_end: span.end,
            close_start: span.end,
            cells: Box::new([]),
            empty: true,
        });
        true
    }

    fn empty_cell(
        &mut self,
        element: &BytesStart<'_>,
        span: CompactSpan,
        decoder: Decoder,
    ) -> bool {
        let Some((row_number, previous_column)) =
            self.row.as_ref().map(|row| (row.number, row.last_column))
        else {
            return false;
        };
        let mut next_column = previous_column;
        let Some(address) = inferred_cell(element, row_number, &mut next_column, decoder) else {
            return false;
        };
        if self
            .entries
            .get(self.entry_index)
            .map(|entry| entry.address)
            != Some(address)
        {
            return false;
        }
        let Some(entry_index) = self.entry_index.checked_add(1) else {
            return false;
        };
        let Some(cell_count) = self.cell_count.checked_add(1) else {
            return false;
        };
        self.entry_index = entry_index;
        self.cell_count = cell_count;
        if !self.reserve_cells() {
            return false;
        }
        let Some(row) = self.row.as_mut() else {
            return false;
        };
        row.last_column = next_column;
        row.cells.push(CompactCellSlot { span });
        true
    }

    fn end(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &quick_xml::events::BytesEnd<'_>,
        span: CompactSpan,
    ) -> bool {
        let Some(kind) = self.stack.pop() else {
            return false;
        };
        if std::str::from_utf8(element.name().as_ref()).is_err() {
            return false;
        }
        let local = element.name().local_name();
        if !is_core(namespace) {
            return false;
        }
        match kind {
            FrameKind::Worksheet if local.as_ref() == b"worksheet" => {
                self.closed_root = true;
                true
            },
            FrameKind::Dimension if local.as_ref() == b"dimension" => self.dimension.is_some(),
            FrameKind::Defaults if local.as_ref() == b"sheetFormatPr" => true,
            FrameKind::Columns if local.as_ref() == b"cols" => self.columns_count != 0,
            FrameKind::Column if local.as_ref() == b"col" => true,
            FrameKind::Value if local.as_ref() == b"v" => true,
            FrameKind::Inline if local.as_ref() == b"is" => true,
            FrameKind::InlineRun if local.as_ref() == b"r" => true,
            FrameKind::InlineText if local.as_ref() == b"t" => true,
            FrameKind::Cell if local.as_ref() == b"c" => self.finish_cell(span.end),
            FrameKind::Row if local.as_ref() == b"row" => self.finish_row(span),
            FrameKind::SheetData if local.as_ref() == b"sheetData" => self.finish_sheet_data(span),
            _ => false,
        }
    }

    fn finish_cell(&mut self, end: u32) -> bool {
        let Some(cell) = self.cell.take() else {
            return false;
        };
        // Formula-free scalar source cells may be empty, `<v>`, or simple
        // inline text.  The authoritative Store eligibility check above
        // rejects rich inline and all formula/shared metadata.
        if !matches!(
            cell.primary,
            None | Some(FrameKind::Value) | Some(FrameKind::Inline)
        ) {
            return false;
        }
        if !self.reserve_cells() {
            return false;
        }
        let Some(row) = self.row.as_mut() else {
            return false;
        };
        row.cells.push(CompactCellSlot {
            span: CompactSpan {
                start: cell.start,
                end,
            },
        });
        true
    }

    fn empty_cell_child(&mut self, local: &[u8]) -> bool {
        let Some(cell) = self.cell.as_mut() else {
            return false;
        };
        if cell.primary.is_some() {
            return false;
        }
        match local {
            b"v" => cell.primary = Some(FrameKind::Value),
            b"is" => cell.primary = Some(FrameKind::Inline),
            _ => return false,
        }
        true
    }

    fn finish_row(&mut self, span: CompactSpan) -> bool {
        let Some(mut row) = self.row.take() else {
            return false;
        };
        row.close_start = span.start;
        let cells = row.cells.into_boxed_slice();
        if !self.reserve_rows() {
            return false;
        }
        let Some(sheet_data) = self.sheet_data.as_mut() else {
            return false;
        };
        sheet_data.rows.push(CompactRowSlot {
            number: row.number,
            span: CompactSpan {
                start: row.start,
                end: span.end,
            },
            tag_end: row.tag_end,
            close_start: row.close_start,
            cells,
            empty: false,
        });
        true
    }

    fn finish_sheet_data(&mut self, span: CompactSpan) -> bool {
        let Some(data) = self.sheet_data.take() else {
            return false;
        };
        self.finished_sheet_data = Some(CompactSheetData {
            span: CompactSpan {
                start: data.start,
                end: span.end,
            },
            tag_end: data.tag_end,
            close_start: span.start,
            rows: data.rows.into_boxed_slice(),
            empty: false,
        });
        true
    }

    fn record_column(&mut self, first: u32, last: u32) -> bool {
        if first <= self.last_column_end {
            return false;
        }
        let Some(count) = self.columns_count.checked_add(1) else {
            return false;
        };
        self.columns_count = count;
        self.last_column_end = last;
        true
    }

    fn text(&self, whitespace: Option<bool>) -> bool {
        let Some(whitespace) = whitespace else {
            return false;
        };
        whitespace
            || matches!(
                self.stack.last(),
                Some(FrameKind::Value | FrameKind::InlineText)
            )
    }

    fn finish(self) -> Option<CompactLayout> {
        if !self.closed_root
            || !self.stack.is_empty()
            || self.sheet_data.is_some()
            || self.row.is_some()
            || self.cell.is_some()
            || !self.seen_sheet_data
            || self.entry_index != self.entries.len()
            || usize::try_from(self.cell_count).ok()? != self.entries.len()
        {
            return None;
        }
        let sheet_data = self.finished_sheet_data?;
        if sheet_data
            .rows
            .iter()
            .map(|row| row.cells.len())
            .sum::<usize>()
            != self.entries.len()
        {
            return None;
        }
        Some(CompactLayout {
            source_ptr: self.source.as_ptr() as usize,
            source_len: self.source.len(),
            entries_ptr: self.entries.as_ptr() as usize,
            entries_len: self.entries.len(),
            cell_count: self.cell_count,
            sheet_data,
            dimension: self.dimension,
        })
    }

    fn push_frame(&mut self, frame: FrameKind) -> bool {
        self.stack.push(frame);
        true
    }

    fn reserve_stack(&mut self) -> bool {
        reserve_vec(&mut self.stack, STACK_GROWTH, &mut self.accounting)
    }

    fn reserve_rows(&mut self) -> bool {
        let (sheet_data, accounting) = (&mut self.sheet_data, &mut self.accounting);
        let Some(sheet_data) = sheet_data.as_mut() else {
            return false;
        };
        reserve_vec(&mut sheet_data.rows, ROW_GROWTH, accounting)
    }

    fn reserve_cells(&mut self) -> bool {
        let (row, accounting) = (&mut self.row, &mut self.accounting);
        let Some(row) = row.as_mut() else {
            return false;
        };
        reserve_vec(&mut row.cells, CELL_GROWTH, accounting)
    }
}

fn reserve_vec<T>(vector: &mut Vec<T>, growth: usize, accounting: &mut usize) -> bool {
    if vector.len() < vector.capacity() {
        return true;
    }
    let old_capacity = vector.capacity();
    let additional = if old_capacity == 0 {
        growth
    } else {
        old_capacity
    };
    let Some(bytes) = additional.checked_mul(size_of::<T>()) else {
        return false;
    };
    let Some(next) = accounting.checked_add(bytes) else {
        return false;
    };
    if next > MAX_PROOF_BYTES || vector.try_reserve_exact(additional).is_err() {
        return false;
    }
    *accounting = next;
    true
}

fn compact_span(start: usize, end: usize) -> Option<CompactSpan> {
    Some(CompactSpan {
        start: u32::try_from(start).ok()?,
        end: u32::try_from(end).ok()?,
    })
}

fn is_core(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == b"http://schemas.openxmlformats.org/spreadsheetml/2006/main" || *value == b"http://purl.oclc.org/ooxml/spreadsheetml/main")
}

fn decode_attributes(element: &BytesStart<'_>, decoder: Decoder) -> bool {
    if std::str::from_utf8(element.name().as_ref()).is_err() {
        return false;
    }
    element.attributes().with_checks(true).all(|attribute| {
        let Ok(attribute) = attribute else {
            return false;
        };
        (attribute.key.prefix().is_none()
            || attribute
                .key
                .prefix()
                .is_some_and(|prefix| prefix.as_ref() == b"xmlns")
            || attribute.key.local_name().as_ref() == b"xmlns")
            && std::str::from_utf8(attribute.key.as_ref()).is_ok()
            && attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .is_ok()
    })
}

fn attribute(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Option<String> {
    unqualified_attribute_value(element, name, decoder)
        .ok()
        .flatten()
}

fn inferred_row(element: &BytesStart<'_>, decoder: Decoder, previous: Option<u32>) -> Option<u32> {
    match attribute(element, b"r", decoder) {
        Some(value) => parse_one_based_row(&value).ok(),
        None => previous.map_or(Some(1), |row| row.checked_add(1)),
    }
    .filter(|row| (1..=ROWS).contains(row))
}

fn inferred_cell(
    element: &BytesStart<'_>,
    row: u32,
    previous_column: &mut u32,
    decoder: Decoder,
) -> Option<Address> {
    let column = if let Some(reference) = attribute(element, b"r", decoder) {
        let (reference_row, column) = parse_a1(&reference).ok()?;
        if reference_row != row {
            return None;
        }
        column
    } else {
        previous_column
            .checked_add(1)
            .or_else(|| if *previous_column == 0 { Some(1) } else { None })?
    };
    if !(1..=COLUMNS).contains(&column) || column <= *previous_column {
        return None;
    }
    *previous_column = column;
    Address::at(row - 1, column - 1).ok()
}

fn parse_column_range(element: &BytesStart<'_>, decoder: Decoder) -> Option<(u32, u32)> {
    let min = attribute(element, b"min", decoder).and_then(|value| value.parse::<u32>().ok())?;
    let max = attribute(element, b"max", decoder).and_then(|value| value.parse::<u32>().ok())?;
    (min != 0 && min <= max && max <= COLUMNS).then_some((min, max))
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::*;

    const CORE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn address(row: u32, column: u32) -> Address {
        Address::at(row, column).expect("test address is in the worksheet grid")
    }

    fn empty_entry(address: Address) -> Stored {
        Stored {
            address,
            cell: Cell::Empty,
            style: None,
            shared_string: None,
            inline_rich: false,
            formula_range: None,
            shared_formula: None,
            cell_metadata: None,
            value_metadata: None,
        }
    }

    fn clear_action(address: Address) -> BTreeMap<Address, Action> {
        let mut actions = BTreeMap::new();
        actions.insert(address, Action::clear(false));
        actions
    }

    #[test]
    fn collects_explicit_and_inferred_scalar_addresses() {
        let source = format!(
            r#"<worksheet xmlns="{CORE}"><dimension ref="A1:B2"/><sheetData><row><c/><c r="B1"/></row><row r="2"><c/></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let entries = vec![
            empty_entry(address(0, 0)),
            empty_entry(address(0, 1)),
            empty_entry(address(1, 0)),
        ];
        let actions = clear_action(address(0, 1));
        let proof = collect_compact_layout(&source, &entries, &actions)
            .expect("plain scalar worksheet has a compact layout");
        assert_eq!(proof.cell_count, 3);
        assert_eq!(proof.sheet_data.rows.len(), 2);
        assert_eq!(proof.sheet_data.rows[0].cells.len(), 2);
        assert_eq!(proof.sheet_data.rows[1].cells.len(), 1);
        assert!(proof.matches_source(&source));
        assert!(proof.matches_entries(&entries));
    }

    #[test]
    fn refuses_source_and_store_identity_changes() {
        let source = format!(
            r#"<worksheet xmlns="{CORE}"><sheetData><row r="1"><c r="A1"/></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let entries = vec![empty_entry(address(0, 0))];
        let proof = collect_compact_layout(&source, &entries, &clear_action(address(0, 0)))
            .expect("plain scalar worksheet has a compact layout");
        let source_copy = source.clone();
        let entries_copy = entries.clone();
        assert!(!proof.matches_source(&source_copy));
        assert!(!proof.matches_entries(&entries_copy));
    }

    #[test]
    fn rejects_dimension_after_columns_and_empty_columns() {
        let after_columns = format!(
            r#"<worksheet xmlns="{CORE}"><cols><col min="1" max="1"/></cols><dimension ref="A1"/><sheetData><row r="1"><c r="A1"/></row></sheetData></worksheet>"#
        );
        let empty_columns = format!(
            r#"<worksheet xmlns="{CORE}"><cols/><sheetData><row r="1"><c r="A1"/></row></sheetData></worksheet>"#
        );
        let entries = vec![empty_entry(address(0, 0))];
        let actions = clear_action(address(0, 0));
        assert!(collect_compact_layout(after_columns.as_bytes(), &entries, &actions).is_none());
        assert!(collect_compact_layout(empty_columns.as_bytes(), &entries, &actions).is_none());
    }
}

#[cfg(test)]
mod resource_tests;
