//! A streaming cursor over the stored cells of a source-backed worksheet
//! (record 0267).
//!
//! [`SourceWorksheet::grid`] parses the whole worksheet part before the first
//! cell can be read. A caller that stops after the first few cells (a query's
//! `first()` or `limit(n)`) pays for every row it never reads. The cursor
//! instead parses **prefixes** of the part: the bytes up to and including the
//! `sheetData` start tag, the first N complete rows, and the two end tags
//! that close the document. That projection is parsed with the same parser,
//! shared strings and style repair the complete grid uses, as the blocks
//! projection outside `sheetData` is (`blocks.rs`), so its records are the
//! complete grid's first records, read again in O(1) by their position.
//!
//! The first prefix holds [`FIRST_ROWS`] rows. When the caller reads past
//! it, the next prefix holds [`GROWTH`] times as many rows, so the parsing
//! repeated for a longer prefix stays a constant fraction of what was read.
//! Once a prefix would cover more than `1 / PREFIX_SHARE` of the part, or no
//! complete row is left after the current one, the cursor reads the complete
//! grid instead (the cached one when retained) and continues after the last
//! coordinate it returned.
//!
//! # What the prefixes guarantee
//!
//! Before the first prefix, one scan of the `sheetData` bytes checks that
//! every row is stored after the rows before it (`rows_ascend`): the
//! complete parse sorts rows stored out of order and merges repeated ones,
//! which no prefix could reproduce, so such a worksheet is read through the
//! complete grid from its first record. Every prefix is then a well-formed
//! worksheet document or its parse fails: a cut that does not fall between
//! two `sheetData` rows (a `</row>` inside a comment, an `AlternateContent`
//! block or markup after `sheetData`) leaves an unclosed or mismatched
//! element, so the parser refuses it, and the cursor reads the complete
//! grid, which owns the refusal or the result. A prefix whose parse records
//! any repair is not used either. Rows ascend, so the complete grid's first
//! records are a prefix's records, and the cursor returns exactly the
//! complete grid's records, in order.
//!
//! What a prefix cannot see is the content after its cut. On a worksheet
//! whose later bytes the complete parse refuses (a malformed record, a limit
//! the whole part exceeds), the cursor returns the leading records, and the
//! refusal when it reaches the complete grid; the complete grid refuses the
//! worksheet before its first record. This is the one difference, and only
//! a reader that stops early can observe it.
//!
//! # Resources
//!
//! The cursor reads the part's verified decoded payload, as the complete
//! parse does (decoding and verifying it whole), and holds it while it
//! parses prefixes. A prefix's projection
//! is a copy of at most `1 / PREFIX_SHARE` of the part; the parsed prefix is
//! offered to the workbook's weighted semantic cache under the same ceilings
//! as the grids, so a later cursor on this worksheet starts from it. Records
//! already parsed are read without source checks, as [`Grid`] reads are;
//! every stage that reads the part runs the worksheet's source and execution
//! fences.

use std::sync::Arc;

use litchi_opc::PartData;
use litchi_sheet::Cell as Address;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::grid::Grid;
use super::values::{Axis, Cached};
use super::{SourceWorksheet, WorksheetKind};
use crate::cell::{RecordIndex, Store};
use crate::error::{Error, Result, allocation};
use crate::raw;
use crate::raw::namespace::{SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE};

/// Rows in the first prefix.
const FIRST_ROWS: usize = 64;

/// How many times more rows each later prefix holds.
const GROWTH: usize = 8;

/// A prefix never covers more than this fraction of the part (as a divisor).
const PREFIX_SHARE: usize = 8;

/// Events the start-tag scan reads before it gives up and reads the complete
/// grid: the elements before `sheetData` are few.
const MAX_HEAD_EVENTS: usize = 4_096;

/// A cursor over a worksheet's stored cells, in row-major order.
///
/// Returned by [`SourceWorksheet::cell_cursor`]. Each item is a stored
/// record's coordinate and its [`RecordIndex`] in [`Self::grid`], which reads
/// it again in O(1): the grid of a leading part of the worksheet while the
/// cursor reads prefixes, then the worksheet's complete grid
/// ([`Self::is_complete`]). The records are those of
/// [`SourceWorksheet::grid`], in its order (see the module documentation for
/// a worksheet the complete parse refuses after the records returned).
///
/// ```no_run
/// use litchi_xlsx::SourceBackedWorkbook;
/// let workbook = SourceBackedWorkbook::open("report.xlsx")?;
/// let sheet = workbook.sheet("Summary")?.expect("worksheet");
/// let mut cursor = sheet.cell_cursor()?;
/// if let Some((address, index)) = cursor.next().transpose()? {
///     let grid = cursor.grid().expect("the record's grid");
///     let (_, cell) = grid.record(index).expect("the record");
///     # let _ = (address, cell);
/// }
/// # Ok::<(), litchi_xlsx::Error>(())
/// ```
pub struct CellCursor {
    sheet: SourceWorksheet,
    stage: Stage,
    /// The coordinate of the last record returned.
    last: Option<Address>,
}

enum Stage {
    Start,
    Prefix(PrefixStage),
    Complete { grid: Grid, next: usize },
    Done,
}

struct PrefixStage {
    prefix: Arc<Prefix>,
    grid: Grid,
    next: usize,
    /// The decoded part, read when the cursor first needs more rows.
    payload: Option<PartData>,
}

/// One parsed leading part of a worksheet: the cache value a later cursor
/// starts from.
pub(super) struct Prefix {
    store: Arc<Store>,
    layout: Layout,
    /// One past the last byte of the prefix's last row.
    end: usize,
    /// The complete rows the prefix holds.
    rows: usize,
}

impl Prefix {
    pub(super) fn weight(&self) -> usize {
        self.store
            .weight_bytes()
            .saturating_add(self.layout.tail.len())
            .saturating_add(self.layout.row_close.len())
    }

    pub(super) fn rows(&self) -> usize {
        self.rows
    }

    /// Whether only the cache holds this prefix's store.
    pub(super) fn is_clean(&self) -> bool {
        Arc::strong_count(&self.store) == 1
    }
}

/// Where the rows of a worksheet part are and how its documents end.
#[derive(Clone)]
struct Layout {
    /// The first byte after the `sheetData` start tag.
    body: usize,
    /// The end tags closing `sheetData` and the root, as written in the part.
    tail: Box<[u8]>,
    /// A row's end tag, in the `sheetData` element's prefix.
    row_close: Box<[u8]>,
    /// The start of a row's start tag (`<row`), in that prefix.
    row_open: Box<[u8]>,
    /// The `sheetData` end tag.
    sheet_close: Box<[u8]>,
}

impl SourceWorksheet {
    /// A cursor over this worksheet's stored cells in row-major order, which
    /// parses only as many leading rows as the caller reads.
    ///
    /// On a worksheet whose grid is resident (pinned by this handle or
    /// retained by the workbook) the cursor reads that grid. Otherwise it
    /// parses growing prefixes of the part and reads the complete grid once a
    /// prefix would cover more than an eighth of it; see [`CellCursor`].
    ///
    /// # Errors
    ///
    /// [`Error::NotWorksheet`] for a chart, dialog, or other non-grid sheet,
    /// and the source and cancellation errors of every source-backed read.
    /// Parse errors are reported by the cursor's reads.
    pub fn cell_cursor(&self) -> Result<CellCursor> {
        let checked = if self.data.kind == WorksheetKind::Worksheet {
            Ok(())
        } else {
            Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            })
        };
        self.finish_result(checked)?;
        Ok(CellCursor {
            sheet: self.clone(),
            stage: Stage::Start,
            last: None,
        })
    }

    /// The parsed grid of this worksheet if it is resident: pinned by this
    /// handle or retained by the workbook (which pins it here).
    fn resident_grid(&self) -> Option<Grid> {
        if let Some(grid) = self.grid.get() {
            return Some(grid.clone());
        }
        let store = self.cached_grid(false)?;
        let _publish_result = self.grid.set(Grid::new(store));
        self.grid.get().cloned()
    }

    /// The verified decoded worksheet part, behind the source fence.
    fn payload(&self) -> Result<PartData> {
        self.owner.execution_check()?;
        let part = self.owner.package.part(&self.data.part_uri)?;
        Ok(part.data()?)
    }

    /// Parse the prefix of `xml` ending at `end`, or `None` when it is not
    /// usable: the parse fails or repairs something.
    fn parse_prefix(&self, xml: &[u8], layout: &Layout, end: usize) -> Result<Option<Store>> {
        let mut document = Vec::new();
        document
            .try_reserve_exact(end.saturating_add(layout.tail.len()))
            .map_err(|source| allocation("worksheet cursor prefix", source))?;
        document.extend_from_slice(&xml[..end]);
        document.extend_from_slice(&layout.tail);
        let parsed = raw::worksheet::parse(&document, || self.owner.shared_strings())
            .and_then(|mut store| self.owner.repair_styles(&mut store).map(|()| store));
        self.owner.execution_check()?;
        self.owner.package.source_version()?;
        Ok(parsed
            .ok()
            .filter(|store| store.tolerance_report().is_clean()))
    }

    /// Offer a parsed prefix to the workbook's cache and return the one it
    /// keeps (the one with more rows).
    fn publish_prefix(&self, prefix: Arc<Prefix>) -> Arc<Prefix> {
        let weight = prefix.weight();
        match self.owner.values().publish(
            self.data.position,
            Cached::Prefix(Arc::clone(&prefix)),
            weight,
        ) {
            Cached::Prefix(kept) if kept.rows >= prefix.rows => kept,
            _ => prefix,
        }
    }

    fn cached_prefix(&self) -> Option<Arc<Prefix>> {
        match self
            .owner
            .values()
            .get(self.data.position, Axis::Prefix, false)?
        {
            Cached::Prefix(prefix) => Some(prefix),
            _ => None,
        }
    }
}

impl CellCursor {
    /// The grid that holds the records returned so far: a prefix's grid or
    /// the worksheet's complete one. `None` before the first record and once
    /// the cursor ended.
    #[must_use]
    pub fn grid(&self) -> Option<&Grid> {
        match &self.stage {
            Stage::Prefix(stage) => Some(&stage.grid),
            Stage::Complete { grid, .. } => Some(grid),
            Stage::Start | Stage::Done => None,
        }
    }

    /// Whether [`Self::grid`] is the worksheet's complete grid.
    #[must_use]
    pub fn is_complete(&self) -> bool {
        matches!(self.stage, Stage::Complete { .. })
    }

    fn step(&mut self) -> Result<Option<(Address, RecordIndex)>> {
        loop {
            match &mut self.stage {
                Stage::Done => return Ok(None),
                Stage::Start => self.stage = self.start()?,
                Stage::Complete { grid, next } => {
                    let index = RecordIndex::new(*next);
                    let Some((address, _)) = grid.record(index) else {
                        self.stage = Stage::Done;
                        return Ok(None);
                    };
                    *next += 1;
                    return Ok(Some((address, index)));
                },
                Stage::Prefix(stage) => {
                    let index = RecordIndex::new(stage.next);
                    if let Some((address, _)) = stage.grid.record(index) {
                        stage.next += 1;
                        return Ok(Some((address, index)));
                    }
                    self.stage = self.extend()?;
                },
            }
        }
    }

    /// The first stage: the resident grid, the cached prefix, the first
    /// prefix, or the complete grid.
    fn start(&self) -> Result<Stage> {
        let sheet = &self.sheet;
        let resident = sheet.resident_grid();
        sheet.finish_result(Ok(()))?;
        if let Some(grid) = resident {
            return Ok(Stage::Complete { grid, next: 0 });
        }
        if let Some(prefix) = sheet.cached_prefix() {
            return Ok(prefix_stage(prefix, None));
        }
        let payload = sheet.payload()?;
        let xml = payload.as_bytes();
        let Some(layout) = locate(xml) else {
            return self.complete();
        };
        if !rows_ascend(xml, &layout) {
            return self.complete();
        }
        let Some((end, rows)) = cut(xml, &layout, layout.body, FIRST_ROWS) else {
            return self.complete();
        };
        if end > xml.len() / PREFIX_SHARE {
            return self.complete();
        }
        let Some(store) = sheet.parse_prefix(xml, &layout, end)? else {
            return self.complete();
        };
        let prefix = sheet.publish_prefix(Arc::new(Prefix {
            store: Arc::new(store),
            layout,
            end,
            rows,
        }));
        Ok(prefix_stage(prefix, Some(payload)))
    }

    /// The stage after an exhausted prefix: a longer prefix, or the complete
    /// grid.
    fn extend(&mut self) -> Result<Stage> {
        let Stage::Prefix(stage) = std::mem::replace(&mut self.stage, Stage::Done) else {
            return Ok(Stage::Done);
        };
        let sheet = &self.sheet;
        let PrefixStage {
            prefix, payload, ..
        } = stage;
        let payload = match payload {
            Some(payload) => payload,
            None => sheet.payload()?,
        };
        let xml = payload.as_bytes();
        let wanted = prefix.rows.saturating_mul(GROWTH - 1);
        let Some((end, rows)) = cut(xml, &prefix.layout, prefix.end, wanted) else {
            return self.complete();
        };
        if end > xml.len() / PREFIX_SHARE {
            return self.complete();
        }
        let Some(store) = sheet.parse_prefix(xml, &prefix.layout, end)? else {
            return self.complete();
        };
        // The longer prefix starts with the records already returned; a
        // prefix that does not (it cannot, without a repair) is not used.
        let returned = prefix.store.stored_cell_count();
        let continues = returned == 0
            || store
                .record(RecordIndex::new(returned - 1))
                .map(|(address, _)| address)
                == self.last;
        if !continues {
            return self.complete();
        }
        let longer = sheet.publish_prefix(Arc::new(Prefix {
            store: Arc::new(store),
            layout: prefix.layout.clone(),
            end,
            rows: prefix.rows + rows,
        }));
        let mut stage = prefix_stage(longer, Some(payload));
        if let Stage::Prefix(stage) = &mut stage {
            stage.next = returned;
        }
        Ok(stage)
    }

    /// The complete grid, positioned after the last record returned.
    fn complete(&self) -> Result<Stage> {
        let grid = self.sheet.grid()?;
        let next = match self.last {
            None => 0,
            Some(last) => first_after(&grid, last),
        };
        Ok(Stage::Complete { grid, next })
    }
}

impl Iterator for CellCursor {
    type Item = Result<(Address, RecordIndex)>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.step() {
            Ok(Some((address, index))) => {
                self.last = Some(address);
                Some(Ok((address, index)))
            },
            Ok(None) => None,
            Err(error) => {
                self.stage = Stage::Done;
                Some(Err(error))
            },
        }
    }
}

impl std::iter::FusedIterator for CellCursor {}

impl std::fmt::Debug for CellCursor {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CellCursor")
            .field("complete", &self.is_complete())
            .finish_non_exhaustive()
    }
}

fn prefix_stage(prefix: Arc<Prefix>, payload: Option<PartData>) -> Stage {
    let grid = Grid::new(Arc::clone(&prefix.store));
    Stage::Prefix(PrefixStage {
        prefix,
        grid,
        next: 0,
        payload,
    })
}

/// The position of the first record of `grid` after `last`, row-major.
fn first_after(grid: &Grid, last: Address) -> usize {
    let (mut low, mut high) = (0, grid.stored_cell_count());
    while low < high {
        let middle = low + (high - low) / 2;
        match grid.record(RecordIndex::new(middle)) {
            Some((address, _)) if address <= last => low = middle + 1,
            _ => high = middle,
        }
    }
    low
}

/// Where the rows of a worksheet part start and how a prefix document
/// ends, or `None` when the part has no direct `sheetData` child of a
/// SpreadsheetML `worksheet` root with content, or the scan cannot read the
/// markup before it (the complete grid then owns the part).
fn locate(xml: &[u8]) -> Option<Layout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root: Option<Box<[u8]>> = None;
    for _ in 0..MAX_HEAD_EVENTS {
        let (namespace, event) = reader.read_resolved_event().ok()?;
        match event {
            Event::Start(start) => {
                let name = start.name();
                if depth == 0 {
                    if root.is_some()
                        || start.local_name().as_ref() != b"worksheet"
                        || !is_spreadsheetml(&namespace)
                    {
                        return None;
                    }
                    root = Some(name.as_ref().into());
                } else if depth == 1
                    && start.local_name().as_ref() == b"sheetData"
                    && is_spreadsheetml(&namespace)
                {
                    let root = root?;
                    let qualified = name.as_ref();
                    let prefix = &qualified[..qualified.len() - b"sheetData".len()];
                    let body = usize::try_from(reader.buffer_position()).ok()?;
                    let mut tail = Vec::with_capacity(qualified.len() + root.len() + 5);
                    tail.extend_from_slice(b"</");
                    tail.extend_from_slice(qualified);
                    tail.extend_from_slice(b"></");
                    tail.extend_from_slice(&root);
                    tail.push(b'>');
                    let mut row_close = Vec::with_capacity(prefix.len() + 6);
                    row_close.extend_from_slice(b"</");
                    row_close.extend_from_slice(prefix);
                    row_close.extend_from_slice(b"row>");
                    let mut row_open = Vec::with_capacity(prefix.len() + 4);
                    row_open.push(b'<');
                    row_open.extend_from_slice(prefix);
                    row_open.extend_from_slice(b"row");
                    let mut sheet_close = Vec::with_capacity(qualified.len() + 3);
                    sheet_close.extend_from_slice(b"</");
                    sheet_close.extend_from_slice(qualified);
                    sheet_close.push(b'>');
                    return Some(Layout {
                        body,
                        tail: tail.into(),
                        row_close: row_close.into(),
                        row_open: row_open.into(),
                        sheet_close: sheet_close.into(),
                    });
                }
                depth += 1;
            },
            Event::End(_) => depth = depth.checked_sub(1)?,
            Event::Eof => return None,
            // An empty `sheetData` has no rows to stream.
            _ => {},
        }
    }
    None
}

/// The end of the `rows`-th complete row after `from`, and the rows found:
/// `None` when fewer than `rows` row end tags follow (the rest of the part
/// is read through the complete grid).
fn cut(xml: &[u8], layout: &Layout, from: usize, rows: usize) -> Option<(usize, usize)> {
    let finder = memchr::memmem::Finder::new(&layout.row_close);
    let mut position = from;
    for _ in 0..rows {
        let found = finder.find(xml.get(position..)?)?;
        position = position
            .checked_add(found)?
            .checked_add(layout.row_close.len())?;
    }
    (rows > 0).then_some((position, rows))
}

/// Whether every row of `sheetData` is stored after the rows before it: each
/// row start tag's `r` (or, without one, the previous row's number plus one)
/// is greater than the previous row's. The complete parse sorts rows stored
/// out of order and merges repeated ones; a prefix cannot see a row stored
/// after its cut, so the cursor reads only worksheets whose rows ascend.
///
/// The scan reads markup bytes, not XML events: it takes every `row` name
/// that follows `<` (or a prefix that does) as a row start tag, so it also
/// reads rows the parser skips (in a comment, a `CDATA` section or a markup
/// compatibility branch it does not select). That only adds rows: the
/// parser's rows are a subsequence of the scanned ones, an implicit row
/// number the scan computes is at least the parser's, and so the parser's
/// rows ascend whenever the scanned ones do. A row element with another
/// prefix than `sheetData`'s, or a row number the scan cannot read the way
/// the parser does, gives `false`.
fn rows_ascend(xml: &[u8], layout: &Layout) -> bool {
    let Some(rest) = xml.get(layout.body..) else {
        return false;
    };
    let body = match memchr::memmem::find(rest, &layout.sheet_close) {
        Some(end) => &rest[..end],
        None => return false,
    };
    // `<` and the element prefix, as `row_open` spells them.
    let prefix = &layout.row_open[1..layout.row_open.len() - 3];
    // The common case: unprefixed rows, and no prefixed `row` anywhere, so
    // the row start tags are exactly the `<row` matches.
    if prefix.is_empty() && memchr::memmem::find(body, b":row").is_none() {
        let starts = memchr::memmem::find_iter(body, b"<row").map(|at| at + 4);
        return ascending(body, starts);
    }
    let finder = memchr::memmem::Finder::new(b"row");
    let mut starts = Vec::new();
    for at in finder.find_iter(body) {
        let after = at + 3;
        // The qualified name's start: `<` right before, or a prefix and its
        // colon after `<`.
        let Some(&before) = at.checked_sub(1).and_then(|index| body.get(index)) else {
            continue;
        };
        let found_prefix = match before {
            b'<' => &body[at..at],
            b':' => {
                let mut start = at - 1;
                while start > 0 && is_name_byte(body[start - 1]) {
                    start -= 1;
                }
                match start.checked_sub(1).and_then(|index| body.get(index)) {
                    Some(b'<') => &body[start..at],
                    // `</x:row` or text such as `a:row`.
                    _ => continue,
                }
            },
            _ => continue,
        };
        if body
            .get(after)
            .copied()
            .is_some_and(|byte| is_name_byte(byte) || byte == b':')
        {
            continue;
        }
        if found_prefix != prefix {
            return false;
        }
        starts.push(after);
    }
    ascending(body, starts.into_iter())
}

/// Whether the rows whose start tags' attributes begin at `starts` ascend.
fn ascending(body: &[u8], starts: impl Iterator<Item = usize>) -> bool {
    let mut previous = 0_u32;
    for start in starts {
        let number = match row_number(&body[start..]) {
            RowNumber::NotRow => continue,
            RowNumber::Implicit => previous.checked_add(1),
            RowNumber::Explicit(number) => Some(number),
            RowNumber::Unreadable => None,
        };
        match number {
            Some(number) if number > previous => previous = number,
            _ => return false,
        }
    }
    true
}

const fn is_name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.')
}

enum RowNumber {
    /// The match is a longer element name (`<rows`).
    NotRow,
    /// The row has no `r`.
    Implicit,
    Explicit(u32),
    Unreadable,
}

/// The `r` of the row start tag whose attributes start `tag`.
fn row_number(tag: &[u8]) -> RowNumber {
    let is_space = |byte: u8| matches!(byte, b' ' | b'\t' | b'\r' | b'\n');
    // What producers write: `r` first, in double quotes.
    if let Some(rest) = tag.strip_prefix(b" r=\"")
        && let Some(length) = memchr::memchr(b'"', rest)
    {
        return digits(&rest[..length]);
    }
    match tag.first() {
        Some(&byte) if is_space(byte) => {},
        Some(b'>' | b'/') => return RowNumber::Implicit,
        _ => return RowNumber::NotRow,
    }
    let mut at = 0;
    loop {
        while tag.get(at).copied().is_some_and(is_space) {
            at += 1;
        }
        match tag.get(at) {
            Some(b'>' | b'/') => return RowNumber::Implicit,
            None => return RowNumber::Unreadable,
            Some(_) => {},
        }
        let name_start = at;
        while tag
            .get(at)
            .is_some_and(|&byte| byte != b'=' && !is_space(byte) && byte != b'>')
        {
            at += 1;
        }
        let name = &tag[name_start..at];
        while tag.get(at).copied().is_some_and(is_space) {
            at += 1;
        }
        if tag.get(at) != Some(&b'=') {
            return RowNumber::Unreadable;
        }
        at += 1;
        while tag.get(at).copied().is_some_and(is_space) {
            at += 1;
        }
        let Some(&quote @ (b'"' | b'\'')) = tag.get(at) else {
            return RowNumber::Unreadable;
        };
        at += 1;
        let value_start = at;
        let Some(length) = memchr::memchr(quote, &tag[at..]) else {
            return RowNumber::Unreadable;
        };
        at += length + 1;
        if name == b"r" {
            return digits(&tag[value_start..value_start + length]);
        }
    }
}

/// A one-based row number written in decimal digits only.
fn digits(value: &[u8]) -> RowNumber {
    if value.is_empty() || value.len() > 9 {
        return RowNumber::Unreadable;
    }
    let mut number = 0_u32;
    for &byte in value {
        if !byte.is_ascii_digit() {
            return RowNumber::Unreadable;
        }
        number = number * 10 + u32::from(byte - b'0');
    }
    if number == 0 {
        RowNumber::Unreadable
    } else {
        RowNumber::Explicit(number)
    }
}

fn is_spreadsheetml(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == SPREADSHEETML_NAMESPACE || *value == STRICT_SPREADSHEETML_NAMESPACE
    )
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_core::OwnedSource;
    use litchi_opc::constants::content_type as ct;
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::super::SourceBackedWorkbook;
    use crate::Cell;

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
    const SHARED_STRINGS: &str = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>zero</t></si><si><t>one</t></si></sst>"#;
    const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/><xf numFmtId="2"/></cellXfs></styleSheet>"#;

    /// A one-sheet workbook whose worksheet part is `xml`.
    fn workbook(xml: &str) -> SourceBackedWorkbook {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        let content_types = format!(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/><Override PartName="/xl/sharedStrings.xml" ContentType="{}"/><Override PartName="/xl/styles.xml" ContentType="{}"/></Types>"#,
            ct::SML_SHEET_MAIN,
            ct::SML_WORKSHEET,
            ct::SML_SHARED_STRINGS,
            ct::SML_STYLES,
        );
        let parts = [
            ("[Content_Types].xml", content_types),
            (
                "_rels/.rels",
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                ),
            ),
            (
                "xl/workbook.xml",
                format!(
                    r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
                ),
            ),
            (
                "xl/_rels/workbook.xml.rels",
                format!(
                    r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="{REL}/styles" Target="styles.xml"/></Relationships>"#
                ),
            ),
            ("xl/sharedStrings.xml", SHARED_STRINGS.to_owned()),
            ("xl/styles.xml", STYLES.to_owned()),
        ];
        for (name, content) in parts {
            writer.write_stored(name, content.as_bytes()).unwrap();
        }
        writer
            .write_deflated("xl/worksheets/sheet1.xml", xml.as_bytes())
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
    }

    /// One row of every value kind the cursor must reproduce: a shared
    /// string, a styled number, a formula with its cached value, an inline
    /// string and a boolean.
    fn row(number: usize) -> String {
        format!(
            r#"<row r="{number}"><c r="A{number}" t="s"><v>{}</v></c><c r="B{number}" s="1"><v>{number}.5</v></c><c r="C{number}"><f>B{number}*2</f><v>{}</v></c><c r="D{number}" t="inlineStr"><is><t>text {number}</t></is></c><c r="E{number}" t="b"><v>{}</v></c></row>"#,
            number % 2,
            number * 2 + 1,
            number % 2,
        )
    }

    fn sheet(rows: impl IntoIterator<Item = usize>, tail: &str) -> String {
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{SML}" xmlns:r="{REL}"><dimension ref="A1:E9"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="2" width="12"/></cols><sheetData>"#
        );
        for number in rows {
            xml.push_str(&row(number));
        }
        xml.push_str("</sheetData>");
        xml.push_str(tail);
        xml.push_str("</worksheet>");
        xml
    }

    /// What the cursor returns, each record read again from its grid.
    fn streamed(workbook: &SourceBackedWorkbook) -> Vec<(String, Cell)> {
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        let mut cursor = sheet.cell_cursor().unwrap();
        let mut out = Vec::new();
        while let Some(item) = cursor.next() {
            let (address, index) = item.unwrap();
            let (stored, cell) = cursor.grid().unwrap().record(index).unwrap();
            assert_eq!(stored, address);
            out.push((format!("{address:?}"), cell.clone()));
        }
        assert!(cursor.next().is_none());
        out
    }

    /// Whether the first record of a fresh cursor comes from the complete
    /// grid.
    fn first_is_complete(xml: &str) -> bool {
        let workbook = workbook(xml);
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        let mut cursor = sheet.cell_cursor().unwrap();
        cursor.next().unwrap().unwrap();
        cursor.is_complete()
    }

    /// What the complete grid holds, in its order.
    fn complete(workbook: &SourceBackedWorkbook) -> Vec<(String, Cell)> {
        let sheet = workbook.sheet("Data").unwrap().unwrap();
        let grid = sheet.grid().unwrap();
        grid.cells(crate::Rect::ALL)
            .unwrap()
            .map(|(address, cell)| (format!("{address:?}"), cell.clone()))
            .collect()
    }

    #[test]
    fn every_stage_returns_the_complete_grids_records() {
        let xml = sheet(
            1..=6_000,
            r#"<mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells>"#,
        );
        let cold = workbook(&xml);
        let sheet = cold.sheet("Data").unwrap().unwrap();
        let mut cursor = sheet.cell_cursor().unwrap();
        let mut stages = Vec::new();
        let mut count = 0;
        while let Some(item) = cursor.next() {
            item.unwrap();
            count += 1;
            if stages.last() != Some(&cursor.is_complete()) {
                stages.push(cursor.is_complete());
            }
        }
        // Prefixes first, then the complete grid.
        assert_eq!(stages, [false, true]);
        assert_eq!(count, 30_000);
        assert_eq!(streamed(&workbook(&xml)), complete(&workbook(&xml)));
    }

    #[test]
    fn a_cursor_that_stops_early_parses_only_a_prefix_and_later_cursors_reuse_it() {
        let xml = sheet(1..=6_000, "");
        let cold = workbook(&xml);
        let sheet = cold.sheet("Data").unwrap().unwrap();
        let mut cursor = sheet.cell_cursor().unwrap();
        let first = cursor.next().unwrap().unwrap();
        assert!(!cursor.is_complete());
        let grid = cursor.grid().unwrap();
        assert!(grid.stored_cell_count() < 6_000 * 5 / 8);
        drop(cursor);
        let builds = cold.semantic_cache_diagnostics().builds;
        // A fresh handle starts from the retained prefix, parsing nothing.
        let again = cold.sheet("Data").unwrap().unwrap();
        let mut cursor = again.cell_cursor().unwrap();
        assert_eq!(cursor.next().unwrap().unwrap(), first);
        assert_eq!(cold.semantic_cache_diagnostics().builds, builds);
        // A resident grid is read directly.
        let _grid = again.grid().unwrap();
        let mut cursor = again.cell_cursor().unwrap();
        assert_eq!(cursor.next().unwrap().unwrap(), first);
        assert!(cursor.is_complete());
    }

    #[test]
    fn prefixed_markup_and_small_parts_stream_like_the_grid() {
        let mut xml = format!(r#"<x:worksheet xmlns:x="{SML}"><x:sheetData>"#);
        for number in 1..=3_000 {
            xml.push_str(&format!(
                r#"<x:row r="{number}"><x:c r="A{number}"><x:v>{number}</x:v></x:c><x:c r="B{number}" t="s"><x:v>1</x:v></x:c></x:row>"#
            ));
        }
        xml.push_str("</x:sheetData></x:worksheet>");
        assert_eq!(streamed(&workbook(&xml)), complete(&workbook(&xml)));
        let small = sheet(1..=3, "");
        assert_eq!(streamed(&workbook(&small)), complete(&workbook(&small)));
        let empty = format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#);
        assert!(streamed(&workbook(&empty)).is_empty());
    }

    #[test]
    fn cuts_that_are_not_between_rows_fall_back_to_the_grid() {
        // The first prefix's last row end tag sits inside a comment.
        let mut xml = sheet(1..=3_000, "");
        let at = xml.find("<row r=\"64\"").unwrap();
        xml.insert_str(at, "<!-- </row> -->");
        assert!(first_is_complete(&xml));
        assert_eq!(streamed(&workbook(&xml)), complete(&workbook(&xml)));
        // Rows stored out of order inside the first prefix: its parse
        // repairs them, so the cursor reads the sorted grid.
        let mut order: Vec<usize> = (1..=3_000).collect();
        order.swap(3, 7);
        let xml = sheet(order, "");
        assert!(first_is_complete(&xml));
        assert_eq!(streamed(&workbook(&xml)), complete(&workbook(&xml)));
        // An ordinary sheet streams its first prefix.
        assert!(!first_is_complete(&sheet(1..=3_000, "")));
    }

    #[test]
    fn a_row_stored_late_after_the_records_read_keeps_its_place() {
        // Row 2,000 is stored last: the complete grid sorts it into place,
        // and the cursor reaches it through the complete grid.
        let order = (1..=1_999).chain(2_001..=3_000).chain([2_000]);
        let xml = sheet(order, "");
        assert_eq!(streamed(&workbook(&xml)), complete(&workbook(&xml)));
        // Rows stored out of order after the first prefix: a prefix cannot
        // see them, so the cursor reads the sorted grid from the start.
        let order = (1..=1_999).chain([5]).chain(2_000..=3_000);
        let xml = sheet(order, "");
        assert!(first_is_complete(&xml));
        let repeated = (1..=3_000).chain([7]);
        assert!(first_is_complete(&sheet(repeated, "")));
    }

    #[test]
    fn a_refusal_after_the_first_prefix_is_reported_when_the_cursor_reaches_it() {
        // The cell reference of the last row is malformed: the complete
        // parse refuses the worksheet, a prefix does not reach it.
        let mut xml = sheet(1..=3_000, "");
        xml = xml.replace(r#"<c r="A3000""#, r#"<c r="A3000:B"#);
        let refused = workbook(&xml);
        let sheet = refused.sheet("Data").unwrap().unwrap();
        assert!(sheet.grid().is_err());
        let mut cursor = sheet.cell_cursor().unwrap();
        assert!(cursor.next().unwrap().is_ok());
        let error = cursor.by_ref().find_map(Result::err);
        assert!(error.is_some());
        assert!(cursor.next().is_none());
    }
}
