//! Compact source-bound worksheet layout proof.
//!
//! The ordinary source parser remains the semantic authority. This module
//! records only the original offsets needed by an existing-row scalar rewrite;
//! every refusal drops the scratch state and leaves the complete edit scanner
//! authoritative at commit time.

use std::borrow::Cow;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use litchi_sheet::{Cell as Address, Rect};

use crate::cell::Stored;
use crate::raw::worksheet::edit::{
    CompactCellSlot, CompactDimensionTag, CompactLayout, CompactRowSlot, CompactSheetData,
    CompactSpan,
};
use crate::raw::worksheet::{SourceEventSpan, SourceHandoff};

const MAX_PROOF_BYTES: usize = 2 * 1024 * 1024;
const CELL_BYTES: usize = std::mem::size_of::<CompactCellSlot>();
const ROW_BYTES: usize = std::mem::size_of::<CompactRowSlot>();
const SHEET_DATA_BYTES: usize = std::mem::size_of::<CompactSheetData>();
const RECORD_GROWTH: usize = 8;
const STACK_GROWTH: usize = 16;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    Worksheet,
    Defaults,
    Columns,
    Column,
    Dimension,
    SheetData,
    Row,
    Cell,
    Primary(PrimaryKind),
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PrimaryKind {
    Formula,
    Value,
    Inline,
}

#[derive(Debug)]
struct PendingCell {
    address: Address,
    start: usize,
    primary: bool,
    primary_open: Option<(PrimaryKind, usize)>,
}

#[derive(Debug)]
struct PendingRow {
    number: u32,
    start: usize,
    tag_end: usize,
    close_start: usize,
    last_column: Option<u32>,
    cells: Vec<CompactCellSlot>,
}

#[derive(Debug)]
struct PendingSheetData {
    start: usize,
    tag_end: usize,
    close_start: usize,
    rows: Vec<CompactRowSlot>,
}

/// Optional proof builder shared with the validator/parser traversal.
#[derive(Debug)]
pub(super) struct CompactProofBuilder {
    source_ptr: usize,
    source_len: usize,
    max_bytes: usize,
    bytes: usize,
    disabled: bool,
    stack: Vec<ElementKind>,
    root_seen: bool,
    root_closed: bool,
    sheet_data: Option<PendingSheetData>,
    sheet_data_finished: Option<CompactSheetData>,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    dimension: Option<CompactDimensionTag>,
    defaults_seen: bool,
    columns_seen: bool,
    column_records: usize,
    last_row_number: Option<u32>,
    cell_count: usize,
}

// The builder itself and the published compact layout contain only fixed-size
// fields. Dynamic row/cell/stack buffers are charged separately before growth.
const FIXED_METADATA_BYTES: usize =
    std::mem::size_of::<CompactProofBuilder>() + std::mem::size_of::<CompactLayout>();

impl CompactProofBuilder {
    pub(super) fn new(source: &[u8]) -> Self {
        let mut builder = Self {
            source_ptr: source.as_ptr() as usize,
            source_len: source.len(),
            max_bytes: MAX_PROOF_BYTES,
            bytes: 0,
            disabled: false,
            stack: Vec::new(),
            root_seen: false,
            root_closed: false,
            sheet_data: None,
            sheet_data_finished: None,
            row: None,
            cell: None,
            dimension: None,
            defaults_seen: false,
            columns_seen: false,
            column_records: 0,
            last_row_number: None,
            cell_count: 0,
        };
        // Charge the fixed live representation up front. This keeps the cap
        // meaningful even for a source that later declines before its first
        // row or cell allocation.
        let _ = builder.charge(FIXED_METADATA_BYTES);
        builder
    }

    #[cfg(test)]
    fn with_cap(source: &[u8], max_bytes: usize) -> Self {
        let mut builder = Self::new(source);
        builder.max_bytes = max_bytes;
        if builder.bytes > max_bytes {
            builder.disable();
        }
        builder
    }

    fn disable(&mut self) {
        self.disabled = true;
        self.bytes = 0;
        self.stack = Vec::new();
        self.sheet_data = None;
        self.sheet_data_finished = None;
        self.row = None;
        self.cell = None;
        self.dimension = None;
        self.defaults_seen = false;
        self.columns_seen = false;
        self.column_records = 0;
        self.last_row_number = None;
        self.cell_count = 0;
    }

    /// Charge metadata before attempting fallible vector growth.
    fn charge(&mut self, amount: usize) -> bool {
        if self.disabled {
            return false;
        }
        let Some(next) = self.bytes.checked_add(amount) else {
            self.disable();
            return false;
        };
        if next > self.max_bytes {
            self.disable();
            return false;
        }
        self.bytes = next;
        true
    }

    fn reserve_stack(&mut self) -> bool {
        if self.stack.len() < self.stack.capacity() {
            return true;
        }
        let Some(amount) = STACK_GROWTH.checked_mul(std::mem::size_of::<ElementKind>()) else {
            self.disable();
            return false;
        };
        if !self.charge(amount) {
            return false;
        }
        let capacity = self.stack.capacity();
        let grew = self.stack.try_reserve_exact(STACK_GROWTH).is_ok()
            && self.stack.capacity().saturating_sub(capacity) <= STACK_GROWTH;
        if !grew {
            self.disable();
            return false;
        }
        true
    }

    fn reserve_rows(&mut self) -> bool {
        let (needs_growth, capacity, length) = match self.sheet_data.as_ref() {
            Some(data) => (
                data.rows.len() == data.rows.capacity(),
                data.rows.capacity(),
                data.rows.len(),
            ),
            None => return false,
        };
        if !needs_growth {
            return true;
        }
        let target = next_record_capacity(length, capacity);
        let Some(additional) = target.checked_sub(length) else {
            self.disable();
            return false;
        };
        let Some(amount) = additional.checked_mul(ROW_BYTES) else {
            self.disable();
            return false;
        };
        if !self.charge(amount) {
            return false;
        }
        let grew = {
            let Some(data) = self.sheet_data.as_mut() else {
                return false;
            };
            data.rows.try_reserve_exact(additional).is_ok()
                && data.rows.capacity().saturating_sub(capacity) <= additional
        };
        if !grew {
            self.disable();
            return false;
        }
        true
    }

    fn reserve_cells(&mut self) -> bool {
        let (needs_growth, capacity, length) = match self.row.as_ref() {
            Some(row) => (
                row.cells.len() == row.cells.capacity(),
                row.cells.capacity(),
                row.cells.len(),
            ),
            None => return false,
        };
        if !needs_growth {
            return true;
        }
        let target = next_record_capacity(length, capacity);
        let Some(additional) = target.checked_sub(length) else {
            self.disable();
            return false;
        };
        let Some(amount) = additional.checked_mul(CELL_BYTES) else {
            self.disable();
            return false;
        };
        if !self.charge(amount) {
            return false;
        }
        let grew = {
            let Some(row) = self.row.as_mut() else {
                return false;
            };
            row.cells.try_reserve_exact(additional).is_ok()
                && row.cells.capacity().saturating_sub(capacity) <= additional
        };
        if !grew {
            self.disable();
            return false;
        }
        true
    }

    fn push_other(&mut self) -> bool {
        self.push_kind(ElementKind::Other)
    }

    fn push_kind(&mut self, kind: ElementKind) -> bool {
        if !self.reserve_stack() {
            return false;
        }
        self.stack.push(kind);
        true
    }

    fn accept_row_number(&mut self, number: u32) -> bool {
        if self
            .last_row_number
            .is_some_and(|previous| number <= previous)
        {
            return false;
        }
        self.last_row_number = Some(number);
        true
    }

    fn accept_cell_column(&mut self, address: Address) -> bool {
        let Some(row) = self.row.as_mut() else {
            return false;
        };
        if address.row().get() + 1 != row.number
            || row
                .last_column
                .is_some_and(|previous| address.column().get() <= previous)
        {
            return false;
        }
        row.last_column = Some(address.column().get());
        true
    }

    fn record_cell_count(&mut self) -> bool {
        let Some(next) = self.cell_count.checked_add(1) else {
            return false;
        };
        if next > u32::MAX as usize {
            return false;
        }
        self.cell_count = next;
        true
    }

    fn observe_start(
        &mut self,
        _namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        span: SourceEventSpan,
    ) -> bool {
        if !self.decode_tag(element, span) {
            return false;
        }
        let local = element.name().local_name();
        let parent = self.stack.last().copied();
        if parent.is_none() && local.as_ref() == b"worksheet" {
            if self.root_seen || !self.reserve_stack() {
                return false;
            }
            self.root_seen = true;
            self.stack.push(ElementKind::Worksheet);
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"sheetFormatPr" {
            if self.defaults_seen
                || self.columns_seen
                || self.sheet_data.is_some()
                || self.sheet_data_finished.is_some()
            {
                return false;
            }
            self.defaults_seen = true;
            if !self.reserve_stack() {
                return false;
            }
            self.stack.push(ElementKind::Defaults);
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"cols" {
            if self.columns_seen || self.sheet_data.is_some() || self.sheet_data_finished.is_some()
            {
                return false;
            }
            self.columns_seen = true;
            self.column_records = 0;
            if !self.reserve_stack() {
                return false;
            }
            self.stack.push(ElementKind::Columns);
            return true;
        }
        if parent == Some(ElementKind::Columns) && local.as_ref() == b"col" {
            self.column_records = self.column_records.saturating_add(1);
            if !self.reserve_stack() {
                return false;
            }
            self.stack.push(ElementKind::Column);
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"sheetData" {
            if self.sheet_data.is_some()
                || self.sheet_data_finished.is_some()
                || (self.columns_seen && self.column_records == 0)
                || !self.reserve_stack()
                || !self.charge(SHEET_DATA_BYTES)
            {
                return false;
            }
            self.sheet_data = Some(PendingSheetData {
                start: span.start,
                tag_end: span.end,
                close_start: 0,
                rows: Vec::new(),
            });
            self.stack.push(ElementKind::SheetData);
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"dimension" {
            if self.defaults_seen
                || self.columns_seen
                || self.sheet_data.is_some()
                || self.sheet_data_finished.is_some()
            {
                return false;
            }
            let Some(reference) = self.attribute_value(element, span, b"ref") else {
                return false;
            };
            let Ok(declared) = Rect::from_a1(reference.as_ref()) else {
                return false;
            };
            if self.dimension.is_some() {
                return false;
            }
            self.dimension = Some(CompactDimensionTag {
                span: model_span(span),
                empty: false,
                declared,
            });
            return self.push_kind(ElementKind::Dimension);
        }
        if parent == Some(ElementKind::SheetData) && local.as_ref() == b"row" {
            if !self.reserve_stack() {
                return false;
            }
            self.stack.push(ElementKind::Row);
            return true;
        }
        if parent == Some(ElementKind::Row) && local.as_ref() == b"c" {
            if !self.reserve_stack() {
                return false;
            }
            self.stack.push(ElementKind::Cell);
            return true;
        }
        if parent == Some(ElementKind::Cell) && matches!(local.as_ref(), b"f" | b"v" | b"is") {
            let kind = match local.as_ref() {
                b"f" => PrimaryKind::Formula,
                b"v" => PrimaryKind::Value,
                b"is" => PrimaryKind::Inline,
                _ => return false,
            };
            if kind == PrimaryKind::Formula {
                // Formula ranges and shared-group diagnostics stay owned by
                // the complete scanner. A scalar proof declines formulas.
                return false;
            }
            {
                let Some(cell) = self.cell.as_mut() else {
                    return false;
                };
                if cell.primary || cell.primary_open.is_some() {
                    return false;
                }
                cell.primary_open = Some((kind, span.start));
            }
            if !self.reserve_stack() {
                return false;
            }
            self.stack.push(ElementKind::Primary(kind));
            return true;
        }
        // The strict source validator admits only established cell children.
        // Retain a small structural stack for inline-string text wrappers.
        self.push_other()
    }

    fn observe_empty(&mut self, element: &BytesStart<'_>, span: SourceEventSpan) -> bool {
        if !self.decode_tag(element, span) {
            return false;
        }
        let local = element.name().local_name();
        let parent = self.stack.last().copied();
        if parent.is_none() && local.as_ref() == b"worksheet" {
            return false;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"sheetData" {
            if self.sheet_data.is_some()
                || self.sheet_data_finished.is_some()
                || (self.columns_seen && self.column_records == 0)
            {
                return false;
            }
            self.sheet_data_finished = Some(CompactSheetData {
                span: model_span(span),
                tag_end: model_offset(span.end),
                close_start: model_offset(span.end),
                rows: Box::new([]),
                empty: true,
            });
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"sheetFormatPr" {
            if self.defaults_seen
                || self.columns_seen
                || self.sheet_data.is_some()
                || self.sheet_data_finished.is_some()
            {
                return false;
            }
            self.defaults_seen = true;
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"cols" {
            return false;
        }
        if parent == Some(ElementKind::Columns) && local.as_ref() == b"col" {
            self.column_records = self.column_records.saturating_add(1);
            return true;
        }
        if parent == Some(ElementKind::Worksheet) && local.as_ref() == b"dimension" {
            if self.defaults_seen
                || self.columns_seen
                || self.sheet_data.is_some()
                || self.sheet_data_finished.is_some()
            {
                return false;
            }
            let Some(reference) = self.attribute_value(element, span, b"ref") else {
                return false;
            };
            let Ok(declared) = Rect::from_a1(reference.as_ref()) else {
                return false;
            };
            if self.dimension.is_some() {
                return false;
            }
            self.dimension = Some(CompactDimensionTag {
                span: model_span(span),
                empty: true,
                declared,
            });
            return true;
        }
        if parent == Some(ElementKind::Cell) && matches!(local.as_ref(), b"f" | b"v" | b"is") {
            if local.as_ref() == b"f" {
                return false;
            }
            let Some(cell) = self.cell.as_mut() else {
                return false;
            };
            if cell.primary || cell.primary_open.is_some() {
                return false;
            }
            cell.primary = true;
        }
        true
    }

    fn observe_end(&mut self, element: &[u8], span: SourceEventSpan) -> bool {
        let Some(kind) = self.stack.pop() else {
            return false;
        };
        match kind {
            ElementKind::Primary(primary_kind) if matches!(element, b"f" | b"v" | b"is") => {
                let Some(cell) = self.cell.as_mut() else {
                    return false;
                };
                let Some((open_kind, _start)) = cell.primary_open.take() else {
                    return false;
                };
                if open_kind != primary_kind || cell.primary {
                    return false;
                }
                cell.primary = true;
            },
            ElementKind::Cell if element == b"c" => {},
            ElementKind::Row if element == b"row" => {},
            ElementKind::SheetData if element == b"sheetData" => {
                let Some(mut data) = self.sheet_data.take() else {
                    return false;
                };
                data.close_start = span.start;
                self.sheet_data_finished = Some(CompactSheetData {
                    span: model_span(SourceEventSpan {
                        start: data.start,
                        end: span.end,
                        decoder: span.decoder,
                    }),
                    tag_end: model_offset(data.tag_end),
                    close_start: model_offset(data.close_start),
                    rows: std::mem::take(&mut data.rows).into_boxed_slice(),
                    empty: false,
                });
            },
            ElementKind::Columns if element == b"cols" => {
                if self.column_records == 0 {
                    return false;
                }
            },
            ElementKind::Worksheet if element == b"worksheet" => {
                self.root_closed = true;
            },
            _ => {},
        }
        true
    }

    fn decode_tag(&mut self, element: &BytesStart<'_>, span: SourceEventSpan) -> bool {
        if std::str::from_utf8(element.name().as_ref()).is_err() {
            return false;
        }
        for attribute in element.attributes().with_checks(true) {
            let Ok(attribute) = attribute else {
                return false;
            };
            if std::str::from_utf8(attribute.key.as_ref()).is_err() {
                return false;
            }
            if attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, span.decoder)
                .is_err()
            {
                return false;
            }
        }
        true
    }

    fn attribute_value<'a>(
        &self,
        element: &'a BytesStart<'a>,
        span: SourceEventSpan,
        wanted: &[u8],
    ) -> Option<Cow<'a, str>> {
        for attribute in element.attributes().with_checks(true) {
            let attribute = attribute.ok()?;
            if attribute.key.as_ref() == wanted {
                return attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, span.decoder)
                    .ok();
            }
        }
        None
    }

    pub(super) fn observe(
        &mut self,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
        span: SourceEventSpan,
    ) {
        if self.disabled {
            return;
        }
        let success = match event {
            Event::Start(element) => self.observe_start(namespace, element, span),
            Event::Empty(element) => self.observe_empty(element, span),
            Event::End(element) => self.observe_end(element.local_name().as_ref(), span),
            Event::Eof => self.root_closed && self.sheet_data_finished.is_some(),
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => true,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => true,
        };
        if !success {
            self.disable();
        }
    }

    pub(super) fn handoff(&mut self, handoff: SourceHandoff) {
        if self.disabled {
            return;
        }
        let success = match handoff {
            SourceHandoff::RowStart { number, span } => {
                if self.row.is_some()
                    || self.stack.last() != Some(&ElementKind::Row)
                    || !self.accept_row_number(number)
                {
                    false
                } else {
                    self.row = Some(PendingRow {
                        number,
                        start: span.start,
                        tag_end: span.end,
                        close_start: 0,
                        last_column: None,
                        cells: Vec::new(),
                    });
                    true
                }
            },
            SourceHandoff::RowEmpty { number, span } => {
                if !self.accept_row_number(number) || !self.reserve_rows() {
                    false
                } else {
                    let Some(data) = self.sheet_data.as_mut() else {
                        self.disable();
                        return false;
                    };
                    data.rows.push(CompactRowSlot {
                        number,
                        span: model_span(span),
                        tag_end: model_offset(span.end),
                        close_start: model_offset(span.end),
                        cells: Box::new([]),
                        empty: true,
                    });
                    true
                }
            },
            SourceHandoff::RowEnd { number, span } => {
                let Some(mut row) = self.row.take() else {
                    self.disable();
                    return false;
                };
                if row.number != number || self.stack.last() != Some(&ElementKind::SheetData) {
                    false
                } else if !self.reserve_rows() {
                    false
                } else {
                    let Some(data) = self.sheet_data.as_mut() else {
                        self.disable();
                        return false;
                    };
                    row.close_start = span.start;
                    data.rows.push(CompactRowSlot {
                        number: row.number,
                        span: model_span(SourceEventSpan {
                            start: row.start,
                            end: span.end,
                            decoder: span.decoder,
                        }),
                        tag_end: model_offset(row.tag_end),
                        close_start: model_offset(row.close_start),
                        cells: std::mem::take(&mut row.cells).into_boxed_slice(),
                        empty: false,
                    });
                    true
                }
            },
            SourceHandoff::CellStart { address, span } => {
                if !self.accept_cell_column(address) {
                    self.disable();
                    return false;
                }
                if self.cell.is_some() || self.stack.last() != Some(&ElementKind::Cell) {
                    false
                } else {
                    self.cell = Some(PendingCell {
                        address,
                        start: span.start,
                        primary: false,
                        primary_open: None,
                    });
                    true
                }
            },
            SourceHandoff::CellEmpty { address, span } => {
                if !self.accept_cell_column(address) {
                    self.disable();
                    return false;
                }
                if !self.reserve_cells() {
                    false
                } else {
                    let Some(row) = self.row.as_mut() else {
                        self.disable();
                        return false;
                    };
                    row.cells.push(CompactCellSlot {
                        span: model_span(span),
                    });
                    self.record_cell_count()
                }
            },
            SourceHandoff::CellEnd { address, span } => {
                let Some(mut cell) = self.cell.take() else {
                    self.disable();
                    return false;
                };
                if cell.address != address
                    || cell.primary_open.is_some()
                    || self.stack.last() != Some(&ElementKind::Row)
                {
                    false
                } else if !self.reserve_cells() {
                    false
                } else {
                    let Some(row) = self.row.as_mut() else {
                        self.disable();
                        return false;
                    };
                    row.cells.push(CompactCellSlot {
                        span: model_span(SourceEventSpan {
                            start: cell.start,
                            end: span.end,
                            decoder: span.decoder,
                        }),
                    });
                    self.record_cell_count()
                }
            },
        };
        if !success {
            self.disable();
        }
    }

    pub(super) fn finish(mut self, entries: &[Stored]) -> Option<CompactLayout> {
        if self.disabled
            || !self.root_seen
            || !self.root_closed
            || !self.stack.is_empty()
            || self.row.is_some()
            || self.cell.is_some()
        {
            return None;
        }
        let sheet_data = self.sheet_data_finished.take()?;
        let slot_count = sheet_data
            .rows
            .iter()
            .try_fold(0usize, |count, row| count.checked_add(row.cells.len()))?;
        if slot_count != self.cell_count {
            return None;
        }
        if entries.len() != self.cell_count {
            return None;
        }
        Some(CompactLayout {
            source_ptr: self.source_ptr,
            source_len: self.source_len,
            entries_ptr: entries.as_ptr() as usize,
            entries_len: entries.len(),
            cell_count: model_offset(self.cell_count),
            sheet_data,
            dimension: self.dimension,
        })
    }
}

fn next_record_capacity(length: usize, capacity: usize) -> usize {
    debug_assert_eq!(length, capacity);
    if capacity == 0 {
        RECORD_GROWTH
    } else {
        capacity
            .checked_mul(2)
            .unwrap_or(usize::MAX)
            .max(length.saturating_add(1))
    }
}

fn model_offset(offset: usize) -> u32 {
    debug_assert!(offset <= u32::MAX as usize);
    offset as u32
}

fn model_span(span: SourceEventSpan) -> CompactSpan {
    CompactSpan {
        start: model_offset(span.start),
        end: model_offset(span.end),
    }
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use litchi_sheet::Cell as Address;

    use super::super::validation::worksheet_xml_and_parse_source;
    use crate::raw::worksheet::edit::{
        Action, rewrite_value_only_with_compact_proof, rewrite_value_only_with_provenance,
    };

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn action(address: Address) -> BTreeMap<Address, Action> {
        let mut actions = BTreeMap::new();
        actions.insert(
            address,
            Action::set(crate::cell::Content::from(crate::cell::Value::Number(
                crate::cell::Number::from(9),
            ))),
        );
        actions
    }

    fn assert_complete_differential(source: &[u8], actions: BTreeMap<Address, Action>) {
        let (store, proof) = worksheet_xml_and_parse_source(source).expect("source parse");
        let proof = proof.expect("compact proof");
        let compact = rewrite_value_only_with_compact_proof(
            source,
            "Sheet1",
            actions.clone(),
            Some(&proof),
            store.entries(),
        )
        .expect("compact rewrite");
        let complete = rewrite_value_only_with_provenance(source, "Sheet1", actions)
            .expect("complete rewrite");
        super::super::validation::worksheet_xml(&compact.bytes)
            .expect("compact emitted worksheet validation");
        assert_eq!(compact.bytes, complete.bytes);
        assert_eq!(compact.omitted, complete.omitted);
        if !compact.omitted.is_empty() {
            let compact_reduced =
                crate::raw::worksheet::edit::reduced_readback(&compact.bytes, &compact.omitted)
                    .expect("compact reduced readback");
            let complete_reduced =
                crate::raw::worksheet::edit::reduced_readback(&complete.bytes, &complete.omitted)
                    .expect("complete reduced readback");
            assert_eq!(compact_reduced, complete_reduced);
            let compact_store = crate::raw::worksheet::parse(&compact_reduced, || Ok(None))
                .expect("compact reduced parse");
            let complete_store = crate::raw::worksheet::parse(&complete_reduced, || Ok(None))
                .expect("complete reduced parse");
            assert_eq!(
                compact_store.entries().len(),
                complete_store.entries().len()
            );
            for (compact, complete) in compact_store.entries().iter().zip(complete_store.entries())
            {
                assert_eq!(compact.address, complete.address);
                assert_eq!(compact.cell, complete.cell);
                assert_eq!(compact.style, complete.style);
            }
        }
    }

    #[test]
    fn compact_prefixed_typed_cells_match_complete_writer() {
        let source = format!(
            r#"<x:worksheet xmlns:x="{SML}"><x:dimension ref="A1:C2"/><x:sheetData><x:row r="1"><x:c r="A1" t="n"><x:v>1</x:v></x:c><x:c t="b"><x:v>0</x:v></x:c></x:row><x:row><x:c r="C2"><x:v>3</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        )
        .into_bytes();
        let (store, proof) = worksheet_xml_and_parse_source(&source).expect("source parse");
        let proof = proof.expect("compact proof");
        let compact = rewrite_value_only_with_compact_proof(
            &source,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
            Some(&proof),
            store.entries(),
        )
        .expect("compact rewrite");
        let complete = rewrite_value_only_with_provenance(
            &source,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
        )
        .expect("complete rewrite");
        assert_eq!(compact.bytes, complete.bytes);
    }

    #[test]
    fn formula_source_declines_compact_proof_and_matches_fallback() {
        let source = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><f>A2</f><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let (store, proof) = worksheet_xml_and_parse_source(&source).expect("source parse");
        assert!(proof.is_none());
        let compact = rewrite_value_only_with_compact_proof(
            &source,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
            proof.as_ref(),
            store.entries(),
        )
        .expect("fallback rewrite");
        let complete = rewrite_value_only_with_provenance(
            &source,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
        )
        .expect("complete rewrite");
        assert_eq!(compact.bytes, complete.bytes);
    }

    #[test]
    fn compact_empty_styled_inferred_cdata_and_ordered_rows_match_complete_writer() {
        let source = format!(
            r#"<worksheet xmlns="{SML}"><dimension ref="A1:C2"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="2" width="12"/></cols><sheetData><row r="1" spans="1:3"><c r="A1"/><c r="B1" s="1"/><c><v><![CDATA[3]]></v></c></row><row><c r="A2" t="str"><v>text</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let mut actions = action(Address::at(0, 1).expect("B1"));
        actions.insert(Address::at(0, 0).expect("A1"), Action::clear(false));
        actions.insert(
            Address::at(0, 2).expect("C1"),
            Action::set(crate::cell::Content::from(crate::cell::Value::Number(
                crate::cell::Number::from(8),
            ))),
        );
        actions.insert(Address::at(1, 0).expect("A2"), Action::style(3));
        assert_complete_differential(&source, actions);
    }

    #[test]
    fn compact_strict_prefixed_and_rebound_namespaces_match_complete_writer() {
        let source = format!(
            r#"<x:worksheet xmlns:x="{SML}" xmlns:y="{SML}"><x:dimension ref="A1:B1"/><x:sheetData><x:row r="1"><y:c r="A1" t="n"><y:v>1</y:v></y:c><y:c r="B1"><y:v>2</y:v></y:c></x:row></x:sheetData></x:worksheet>"#
        )
        .into_bytes();
        assert_complete_differential(&source, action(Address::at(0, 1).expect("B1")));
    }

    #[test]
    fn compact_dimension_variants_and_noop_match_complete_writer() {
        for source in [
            format!(
                r#"<worksheet xmlns="{SML}"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
            )
            .into_bytes(),
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
            )
            .into_bytes(),
        ] {
            let (store, proof) = worksheet_xml_and_parse_source(&source).expect("source parse");
            let proof = proof.expect("compact proof");
            let mut actions = action(Address::at(0, 1).expect("B1"));
            let compact = rewrite_value_only_with_compact_proof(
                &source,
                "Sheet1",
                actions.clone(),
                Some(&proof),
                store.entries(),
            )
            .expect("compact rewrite");
            let complete = rewrite_value_only_with_provenance(&source, "Sheet1", actions)
                .expect("complete rewrite");
            assert_eq!(compact.bytes, complete.bytes);
            assert_eq!(compact.omitted, complete.omitted);

            let noop = rewrite_value_only_with_compact_proof(
                &source,
                "Sheet1",
                BTreeMap::new(),
                Some(&proof),
                store.entries(),
            )
            .expect("compact no-op");
            assert_eq!(noop.bytes, source);
            assert!(noop.omitted.is_empty());
        }
    }

    #[test]
    fn compact_source_identity_mismatch_falls_back_without_byte_drift() {
        let source = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let (store, proof) = worksheet_xml_and_parse_source(&source).expect("source parse");
        let proof = proof.expect("compact proof");
        let detached = source.clone();
        let compact = rewrite_value_only_with_compact_proof(
            &detached,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
            Some(&proof),
            store.entries(),
        )
        .expect("identity fallback");
        let complete = rewrite_value_only_with_provenance(
            &detached,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
        )
        .expect("complete rewrite");
        assert_eq!(compact.bytes, complete.bytes);
        assert_eq!(compact.omitted, complete.omitted);
    }

    #[test]
    fn compact_store_identity_mismatch_falls_back_without_byte_drift() {
        let source = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let (_store, proof) = worksheet_xml_and_parse_source(&source).expect("source parse");
        let proof = proof.expect("compact proof");
        let foreign_source = source.clone();
        let (foreign_store, _) =
            worksheet_xml_and_parse_source(&foreign_source).expect("foreign source parse");
        let compact = rewrite_value_only_with_compact_proof(
            &source,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
            Some(&proof),
            foreign_store.entries(),
        )
        .expect("store identity fallback");
        let complete = rewrite_value_only_with_provenance(
            &source,
            "Sheet1",
            action(Address::at(0, 1).expect("B1")),
        )
        .expect("complete rewrite");
        assert_eq!(compact.bytes, complete.bytes);
        assert_eq!(compact.omitted, complete.omitted);
        assert!(!proof.matches_entries(foreign_store.entries()));
    }

    #[test]
    fn late_attribute_decode_refusal_preserves_complete_error() {
        let source = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1" spans="&bad;"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
        )
        .into_bytes();
        let shared = worksheet_xml_and_parse_source(&source).expect("parser ignores proof refusal");
        assert!(shared.1.is_none());
        let complete = rewrite_value_only_with_provenance(
            &source,
            "Sheet1",
            action(Address::at(0, 0).expect("A1")),
        );
        let candidate = rewrite_value_only_with_compact_proof(
            &source,
            "Sheet1",
            action(Address::at(0, 0).expect("A1")),
            shared.1.as_ref(),
            shared.0.entries(),
        );
        assert_eq!(
            candidate.as_ref().err().map(|error| format!("{error:?}")),
            complete.as_ref().err().map(|error| format!("{error:?}")),
        );
    }

    #[test]
    fn compact_cap_refusal_drops_all_builder_scratch() {
        let source = b"<worksheet/>";
        let mut builder = CompactProofBuilder::with_cap(source, 0);
        assert!(!builder.charge(1));
        assert!(builder.disabled);
        assert_eq!(builder.bytes, 0);
        assert_eq!(builder.stack.capacity(), 0);
        assert!(builder.finish(&[]).is_none());
    }

    #[test]
    fn compact_populated_cap_refusal_drops_retained_stack() {
        let source = b"<worksheet/>";
        let mut builder = CompactProofBuilder::with_cap(source, usize::MAX);
        assert!(builder.reserve_stack());
        let capacity = builder.stack.capacity();
        builder.stack.resize(capacity, ElementKind::Other);
        builder.max_bytes = builder.bytes;
        assert!(!builder.reserve_stack());
        assert!(builder.disabled);
        assert_eq!(builder.bytes, 0);
        assert!(builder.stack.is_empty());
        assert_eq!(builder.stack.capacity(), 0);
        assert!(builder.sheet_data.is_none());
        assert!(builder.row.is_none());
        assert!(builder.cell.is_none());
    }
}
