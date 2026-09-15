//! Compact source-bound worksheet facts captured during planning.
//!
//! The source-backed value editor visits every row and cell of the touched
//! worksheet twice: once while planning validates and parses the part, and
//! again while the commit reconstructs a complete [`Layout`](super::Layout)
//! from the same bytes. Change 0550 measured the second walk at 53.24-54.94%
//! of commit instructions.
//!
//! This module carries the *few* facts the value-only rewrite actually needs
//! from the first walk to the second: the `<sheetData>` envelope, one record
//! per row, and sixteen bytes per cell (its resolved address and the exact
//! source span of its `<c>` element). Nothing else is retained. Everything a
//! *changed* cell's rewrite needs beyond those bytes — the owned tag, the
//! payload spans, the empty form — is materialized from the retained span at
//! commit, and only for the at most 256 cells one commit may touch.
//!
//! The builder is deliberately timid. It observes the events of the shared
//! planning traversal, mirrors the scanner's structural decisions for the
//! vocabulary the value-only validator admits, and *declines* — dropping
//! every fact — the moment it meets anything it cannot prove the scanner
//! would accept identically. A decline is never an error: the commit then
//! runs today's scan and produces today's bytes and today's diagnostics.

use litchi_sheet::{COLUMNS, Cell as Address, Rect};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::Reader;

use super::super::wire::{cell_tag, tag};
use super::model::{CellFact, CellSlot, DimensionFact, RowFact, SourceFacts, Span, Tag};
use crate::error::{Result, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;
use crate::raw::worksheet::model::{MAX_XML_DEPTH, MAX_XML_EVENTS};
use crate::raw::worksheet::{parse_a1, parse_one_based_row};

/// Source byte range of one traversal event.
#[derive(Debug, Clone, Copy)]
pub(crate) struct EventSpan {
    pub(crate) start: usize,
    pub(crate) end: usize,
}

/// The element kinds the builder is prepared to recognize.
///
/// The names mirror the scanner's own frame kinds so the two can be read
/// side by side. Anything the builder cannot place in this vocabulary makes
/// it decline rather than guess.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Root,
    Columns,
    SheetData,
    Row,
    Cell,
    Primary,
    Other,
}

#[derive(Debug)]
struct PendingRow {
    number: u32,
    start: usize,
    tag_end: usize,
    first_cell: usize,
}

#[derive(Debug)]
struct PendingCell {
    address: Address,
    start: usize,
}

/// Collects [`SourceFacts`] alongside the planning traversal.
#[derive(Debug)]
pub(crate) struct FactsBuilder {
    declined: bool,
    events: usize,
    stack: Vec<Kind>,
    root_seen: bool,
    root_close_start: Option<usize>,
    dimension: Option<DimensionFact>,
    defaults_start: Option<usize>,
    columns_start: Option<usize>,
    column_records: usize,
    sheet_data_start: Option<usize>,
    sheet_data_tag_end: usize,
    sheet_data_close_start: usize,
    sheet_data_end: usize,
    sheet_data_empty: bool,
    sheet_data_closed: bool,
    previous_row: u32,
    row: Option<PendingRow>,
    cell: Option<PendingCell>,
    last_column: u32,
    rows: Vec<RowFact>,
    cells: Vec<CellFact>,
}

impl FactsBuilder {
    pub(crate) fn new() -> Self {
        Self {
            declined: false,
            events: 0,
            stack: Vec::new(),
            root_seen: false,
            root_close_start: None,
            dimension: None,
            defaults_start: None,
            columns_start: None,
            column_records: 0,
            sheet_data_start: None,
            sheet_data_tag_end: 0,
            sheet_data_close_start: 0,
            sheet_data_end: 0,
            sheet_data_empty: false,
            sheet_data_closed: false,
            previous_row: 0,
            row: None,
            cell: None,
            last_column: 0,
            rows: Vec::new(),
            cells: Vec::new(),
        }
    }

    /// Drop every fact and stop observing.
    ///
    /// A decline is a private, silent outcome: it never becomes a planning
    /// error and never changes the events the traversal delivers.
    fn decline(&mut self) {
        self.declined = true;
        self.stack = Vec::new();
        self.rows = Vec::new();
        self.cells = Vec::new();
        self.row = None;
        self.cell = None;
        self.dimension = None;
    }

    /// Observe one traversal event. Always returns; never aborts the caller.
    pub(crate) fn observe(
        &mut self,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
        span: EventSpan,
        content: &[u8],
    ) {
        if self.declined {
            return;
        }
        match self.observe_inner(namespace, event, span, content) {
            Some(()) => {},
            None => self.decline(),
        }
    }

    fn observe_inner(
        &mut self,
        namespace: &ResolveResult<'_>,
        event: &Event<'_>,
        span: EventSpan,
        content: &[u8],
    ) -> Option<()> {
        self.events = self.events.checked_add(1)?;
        if self.events > MAX_XML_EVENTS {
            return None;
        }
        match event {
            Event::Start(element) => {
                if self.stack.len() >= MAX_XML_DEPTH {
                    return None;
                }
                let kind = self.element(namespace, element, span, content, false)?;
                self.stack.try_reserve(1).ok()?;
                self.stack.push(kind);
                Some(())
            },
            Event::Empty(element) => {
                self.element(namespace, element, span, content, true)?;
                Some(())
            },
            Event::End(_) => {
                let kind = self.stack.pop()?;
                self.close(kind, span)
            },
            // Character data, references, comments, declarations and
            // processing instructions carry no layout the value-only rewrite
            // reads; the scanner only inspects them inside a formula or a
            // merged-range container, and the builder admits neither.
            Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Eof => Some(()),
            Event::DocType(_) => None,
        }
    }

    fn element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        span: EventSpan,
        content: &[u8],
        empty: bool,
    ) -> Option<Kind> {
        // Every tag the scanner materializes is decoded and normalized. An
        // ampersand is the only byte that can make that decode fail, so a
        // tag that carries one is not provably reproducible and the builder
        // declines rather than risk losing the scanner's diagnostic.
        if memchr::memchr(b'&', content.get(span.start..span.end)?).is_some() {
            return None;
        }
        let name = element.name();
        let local = name.local_name();
        let local = local.as_ref();
        let sml = is_spreadsheetml_name(namespace, name, local);
        if !sml {
            return None;
        }
        let parent = self.stack.last().copied();
        match parent {
            None => {
                if local != b"worksheet" || self.root_seen || empty {
                    return None;
                }
                self.root_seen = true;
                Some(Kind::Root)
            },
            Some(Kind::Root) => self.worksheet_child(element, local, span, empty),
            Some(Kind::Columns) => {
                if local != b"col" {
                    return None;
                }
                self.column(element)?;
                Some(Kind::Other)
            },
            Some(Kind::SheetData) => {
                if local != b"row" {
                    return None;
                }
                self.row(element, span, empty)?;
                Some(Kind::Row)
            },
            Some(Kind::Row) => {
                if local != b"c" {
                    return None;
                }
                self.cell(element, span, empty)?;
                Some(Kind::Cell)
            },
            Some(Kind::Cell) => {
                // `<f>` drags in shared and array formula groups, formula
                // ranges and the scanner's formula text bookkeeping. None of
                // that is compactly reproducible, so a worksheet that carries
                // one keeps the scan.
                if !matches!(local, b"v" | b"is") {
                    return None;
                }
                Some(Kind::Primary)
            },
            Some(Kind::Primary) => {
                if local != b"t" {
                    return None;
                }
                Some(Kind::Other)
            },
            Some(Kind::Other) => Some(Kind::Other),
        }
    }

    fn worksheet_child(
        &mut self,
        element: &BytesStart<'_>,
        local: &[u8],
        span: EventSpan,
        empty: bool,
    ) -> Option<Kind> {
        match local {
            b"dimension" => {
                if self.dimension.is_some() {
                    return None;
                }
                let reference = raw_attribute(element, b"ref")?;
                if !reference
                    .iter()
                    .all(|byte| byte.is_ascii_alphanumeric() || *byte == b':')
                {
                    return None;
                }
                let declared = Rect::from_a1(std::str::from_utf8(reference).ok()?).ok()?;
                self.dimension = Some(DimensionFact {
                    span: Span {
                        start: span.start,
                        end: span.end,
                    },
                    empty,
                    declared,
                });
                Some(Kind::Other)
            },
            b"sheetFormatPr" => {
                if self.defaults_start.is_some()
                    || self.columns_start.is_some()
                    || self.sheet_data_start.is_some()
                {
                    return None;
                }
                self.defaults_start = Some(span.start);
                Some(Kind::Other)
            },
            b"cols" => {
                if self.columns_start.is_some() || self.sheet_data_start.is_some() {
                    return None;
                }
                // An empty `<cols/>` has no `col` record and the scanner
                // refuses it.
                if empty {
                    return None;
                }
                self.columns_start = Some(span.start);
                Some(Kind::Columns)
            },
            b"sheetData" => {
                if self.sheet_data_start.is_some() {
                    return None;
                }
                self.sheet_data_start = Some(span.start);
                self.sheet_data_tag_end = span.end;
                self.sheet_data_empty = empty;
                if empty {
                    self.sheet_data_close_start = span.end;
                    self.sheet_data_end = span.end;
                    self.sheet_data_closed = true;
                    return Some(Kind::Other);
                }
                Some(Kind::SheetData)
            },
            b"sheetViews" => Some(Kind::Other),
            // Any other direct worksheet child is a container the value-only
            // validator does not admit, or one whose presence changes where
            // the scanner would place a merged-range insertion point.
            _ => None,
        }
    }

    fn column(&mut self, element: &BytesStart<'_>) -> Option<()> {
        let min = raw_u32(element, b"min")?;
        let max = raw_u32(element, b"max")?;
        if min == 0 || min > max || max > COLUMNS {
            return None;
        }
        self.column_records = self.column_records.checked_add(1)?;
        Some(())
    }

    fn row(&mut self, element: &BytesStart<'_>, span: EventSpan, empty: bool) -> Option<()> {
        let reference = raw_attribute(element, b"r")?;
        if !reference.iter().all(u8::is_ascii_digit) {
            return None;
        }
        let number = parse_one_based_row(std::str::from_utf8(reference).ok()?).ok()?;
        if self.previous_row != 0 && number <= self.previous_row {
            return None;
        }
        self.previous_row = number;
        self.last_column = 0;
        if empty {
            self.push_row(RowFact {
                number,
                start: u32::try_from(span.start).ok()?,
                tag_end: u32::try_from(span.end).ok()?,
                close_start: u32::try_from(span.end).ok()?,
                end: u32::try_from(span.end).ok()?,
                first_cell: u32::try_from(self.cells.len()).ok()?,
                cell_count: 0,
                empty: true,
            })?;
            return Some(());
        }
        self.row = Some(PendingRow {
            number,
            start: span.start,
            tag_end: span.end,
            first_cell: self.cells.len(),
        });
        Some(())
    }

    fn cell(&mut self, element: &BytesStart<'_>, span: EventSpan, empty: bool) -> Option<()> {
        let row = self.row.as_ref()?.number;
        let reference = raw_attribute(element, b"r")?;
        if !reference.iter().all(u8::is_ascii_alphanumeric) {
            return None;
        }
        let (reference_row, column) = parse_a1(std::str::from_utf8(reference).ok()?).ok()?;
        if reference_row != row || column <= self.last_column {
            return None;
        }
        self.last_column = column;
        let address = Address::at(row.checked_sub(1)?, column.checked_sub(1)?).ok()?;
        if empty {
            self.push_cell(CellFact {
                address,
                start: u32::try_from(span.start).ok()?,
                end: u32::try_from(span.end).ok()?,
            })?;
            return Some(());
        }
        self.cell = Some(PendingCell {
            address,
            start: span.start,
        });
        Some(())
    }

    fn close(&mut self, kind: Kind, span: EventSpan) -> Option<()> {
        match kind {
            Kind::Root => {
                self.root_close_start = Some(span.start);
            },
            Kind::SheetData => {
                self.sheet_data_close_start = span.start;
                self.sheet_data_end = span.end;
                self.sheet_data_closed = true;
            },
            Kind::Row => {
                let row = self.row.take()?;
                self.push_row(RowFact {
                    number: row.number,
                    start: u32::try_from(row.start).ok()?,
                    tag_end: u32::try_from(row.tag_end).ok()?,
                    close_start: u32::try_from(span.start).ok()?,
                    end: u32::try_from(span.end).ok()?,
                    first_cell: u32::try_from(row.first_cell).ok()?,
                    cell_count: u32::try_from(self.cells.len().checked_sub(row.first_cell)?)
                        .ok()?,
                    empty: false,
                })?;
            },
            Kind::Cell => {
                let cell = self.cell.take()?;
                self.push_cell(CellFact {
                    address: cell.address,
                    start: u32::try_from(cell.start).ok()?,
                    end: u32::try_from(span.end).ok()?,
                })?;
            },
            Kind::Columns | Kind::Primary | Kind::Other => {},
        }
        Some(())
    }

    fn push_row(&mut self, row: RowFact) -> Option<()> {
        self.rows.try_reserve(1).ok()?;
        self.rows.push(row);
        Some(())
    }

    fn push_cell(&mut self, cell: CellFact) -> Option<()> {
        self.cells.try_reserve(1).ok()?;
        self.cells.push(cell);
        Some(())
    }

    /// Publish the facts, or nothing when the builder declined.
    ///
    /// The final checks repeat [`Layout`](super::Layout) construction's own
    /// ordering rules, so a source the scanner would refuse never produces
    /// facts and always keeps its refusal.
    pub(crate) fn finish(self, content: &[u8]) -> Option<SourceFacts> {
        if self.declined || !self.stack.is_empty() || !self.root_seen {
            return None;
        }
        self.root_close_start?;
        let sheet_data_start = self.sheet_data_start?;
        if !self.sheet_data_closed || self.row.is_some() || self.cell.is_some() {
            return None;
        }
        if let Some(defaults) = self.defaults_start {
            if defaults >= sheet_data_start {
                return None;
            }
            if self
                .columns_start
                .is_some_and(|columns| defaults >= columns)
            {
                return None;
            }
            if self
                .dimension
                .as_ref()
                .is_some_and(|dimension| dimension.span.start >= defaults)
            {
                return None;
            }
        }
        if let Some(columns) = self.columns_start {
            if self.column_records == 0 || columns >= sheet_data_start {
                return None;
            }
        }
        if self
            .dimension
            .as_ref()
            .is_some_and(|dimension| dimension.span.start >= sheet_data_start)
        {
            return None;
        }
        Some(SourceFacts {
            source_ptr: content.as_ptr() as usize,
            source_len: content.len(),
            dimension: self.dimension,
            sheet_data: Span {
                start: sheet_data_start,
                end: self.sheet_data_end,
            },
            sheet_data_tag_end: self.sheet_data_tag_end,
            sheet_data_close_start: self.sheet_data_close_start,
            sheet_data_empty: self.sheet_data_empty,
            rows: self.rows.into_boxed_slice(),
            cells: self.cells.into_boxed_slice(),
        })
    }
}

/// Return the raw bytes of one unprefixed attribute, refusing duplicates.
///
/// The caller has already proved the element carries no ampersand, so these
/// bytes decode to themselves; the caller additionally restricts the value's
/// alphabet so attribute-value normalization is the identity as well.
fn raw_attribute<'a>(element: &'a BytesStart<'a>, name: &[u8]) -> Option<&'a [u8]> {
    let mut attributes = element.attributes();
    attributes.with_checks(false);
    let mut found = None;
    for attribute in attributes {
        let attribute = attribute.ok()?;
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == name {
            if found.is_some() {
                return None;
            }
            found = Some(attribute.value);
        }
    }
    match found? {
        std::borrow::Cow::Borrowed(value) => Some(value),
        std::borrow::Cow::Owned(_) => None,
    }
}

fn raw_u32(element: &BytesStart<'_>, name: &[u8]) -> Option<u32> {
    let value = raw_attribute(element, name)?;
    if !value.iter().all(u8::is_ascii_digit) {
        return None;
    }
    std::str::from_utf8(value).ok()?.parse::<u32>().ok()
}

/// Rebuild one cell's complete scanner slot from its retained source span.
///
/// This is the "changed-cell tag materialization" half of the design: at most
/// 256 cells per commit pay it, and every field it produces is the field the
/// scanner would have produced for the same bytes. Any disagreement with the
/// retained fact, and any shape the builder did not admit, returns `None` so
/// the caller can fall back to the complete scan.
pub(crate) fn materialize_cell(content: &[u8], fact: &CellFact) -> Result<Option<CellSlot>> {
    let start = fact.start as usize;
    let end = fact.end as usize;
    let Some(slice) = content.get(start..end) else {
        return Ok(None);
    };
    let mut reader = Reader::from_reader(slice);
    reader.config_mut().check_end_names = true;
    let decoder = reader.decoder();
    let Ok(first) = reader.read_event() else {
        return Ok(None);
    };
    let (element, empty) = match first {
        Event::Start(element) => (element, false),
        Event::Empty(element) => (element, true),
        _ => return Ok(None),
    };
    if element.name().local_name().as_ref() != b"c" {
        return Ok(None);
    }
    let tag_end = start + relative_position(&reader)?;
    let slot_tag = cell_tag(&element, decoder)?;
    if empty {
        if tag_end != end {
            return Ok(None);
        }
        return Ok(Some(CellSlot {
            address: fact.address,
            span: Span { start, end },
            tag_end: end,
            close_start: end,
            tag: slot_tag,
            primary: Box::new([]),
            mce_payload: false,
            empty: true,
        }));
    }

    let mut primary = Vec::new();
    let mut depth = 0usize;
    let mut child_start = 0usize;
    let close_start;
    loop {
        let cursor = start + relative_position(&reader)?;
        let Ok(event) = reader.read_event() else {
            return Ok(None);
        };
        let after = start + relative_position(&reader)?;
        match event {
            Event::Start(child) => {
                if !admitted_child(&child, depth) {
                    return Ok(None);
                }
                if depth == 0 {
                    child_start = cursor;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet cell payload depth overflow"))?;
            },
            Event::Empty(child) => {
                if !admitted_child(&child, depth) {
                    return Ok(None);
                }
                if depth == 0 {
                    push_span(
                        &mut primary,
                        Span {
                            start: cursor,
                            end: after,
                        },
                    )?;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    close_start = cursor;
                    break;
                }
                depth -= 1;
                if depth == 0 {
                    push_span(
                        &mut primary,
                        Span {
                            start: child_start,
                            end: after,
                        },
                    )?;
                }
            },
            Event::Eof => return Ok(None),
            Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
            Event::DocType(_) => return Ok(None),
        }
    }
    if start + relative_position(&reader)? != end {
        return Ok(None);
    }
    Ok(Some(CellSlot {
        address: fact.address,
        span: Span { start, end },
        tag_end,
        close_start,
        tag: slot_tag,
        primary: primary.into_boxed_slice(),
        mce_payload: false,
        empty: false,
    }))
}

/// Accept only the payload vocabulary the value-only validator admits below a
/// cell: `<v>` and `<is>` directly, and `<t>` inside `<is>`.
fn admitted_child(element: &BytesStart<'_>, depth: usize) -> bool {
    let name = element.name();
    let local = name.local_name();
    if depth == 0 {
        matches!(local.as_ref(), b"v" | b"is")
    } else {
        local.as_ref() == b"t"
    }
}

fn push_span(primary: &mut Vec<Span>, span: Span) -> Result<()> {
    primary
        .try_reserve(1)
        .map_err(|source| allocation("worksheet cell payload", source))?;
    primary.push(span);
    Ok(())
}

/// Rebuild one non-cell element's owned tag from its retained span.
pub(crate) fn materialize_tag(content: &[u8], span: Span) -> Result<Option<Tag>> {
    let Some(slice) = content.get(span.start..span.end) else {
        return Ok(None);
    };
    let mut reader = Reader::from_reader(slice);
    reader.config_mut().check_end_names = true;
    let decoder = reader.decoder();
    let Ok(event) = reader.read_event() else {
        return Ok(None);
    };
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Ok(None),
    };
    Ok(Some(tag(&element, decoder)?))
}

fn relative_position(reader: &Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source| invalid("worksheet XML position does not fit usize"))
}
