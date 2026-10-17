//! Lazy children of XLSB nodes (record 0010 §2.5, §4.7, §5.3).
//!
//! The adapter traverses the retained sparse worksheet grid directly. A
//! cursor copies at most [`BATCH`] coordinates from one borrowed range scan,
//! so a query that stops early does not expand a sparse BIFF12 worksheet into
//! a dense allocation. Empty-cell stubs are emitted only for an explicitly
//! budgeted dense range view.

use litchi_core::sheet::{Cell as CoreCell, CellValue};
use litchi_xlsb::sheet::Rect;

use super::Xlsb;
use super::node::{CellSource, Grid, Item, Node, Sheet, defined_names};
use super::owner::SheetHandle;
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::node::ViewOptions;
use crate::path::navigate::KindSet;

/// How many coordinates a sparse cursor retains per refill.
const BATCH: usize = 64;

/// The adapter's lazy children iterator.
pub(crate) struct Children<'f> {
    file: &'f Xlsb,
    state: State,
}

/// A zero-based `(row, column)` coordinate. The XLSB leaf's checked address
/// type remains behind the node/grid bridge; the adapter never constructs an
/// unchecked leaf coordinate.
type Address = (u32, u32);

enum State {
    Done,
    Workbook {
        next: usize,
        count: usize,
        sheets: bool,
        names: bool,
        name: usize,
    },
    Sheet(Box<SheetItems>),
    Tables {
        sheet: usize,
        parts: std::vec::IntoIter<litchi_xlsb::package::table::Table>,
        index: usize,
    },
    Cells {
        sheet: usize,
        grid: Grid,
        cursor: Box<Cursor>,
    },
    /// A source-backed worksheet's cells through the leaf's cell cursor,
    /// which decodes only as many cells as a walk that stops early reads
    /// (record 0267).
    Stream(Box<Stream>),
    Dense {
        sheet: usize,
        grid: Grid,
        rect: Rect,
        next: Option<Address>,
    },
    /// A worksheet's features (record 0322).
    Features {
        sheet: usize,
        items: std::vec::IntoIter<super::features::Feature>,
    },
    /// A worksheet's charts, or a chart's series and axes (record 0322).
    Charts {
        sheet: usize,
        items: std::vec::IntoIter<(Kind, super::chart::ChartItem)>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Want {
    columns: bool,
    rows: bool,
    cells: bool,
    tables: bool,
}

struct SheetItems {
    sheet: usize,
    grid: Grid,
    handle: SheetHandle,
    want: Want,
    /// Node-view listing filters formatting-only cells and empty rows.
    view: bool,
    phase: Phase,
}

enum Phase {
    Start,
    Columns {
        ranges: Vec<(u32, u32)>,
        range: usize,
        next: u32,
    },
    Rows {
        rows: Vec<u32>,
        next: usize,
        cells: Option<Cursor>,
    },
    Cells(Cursor),
    Tables {
        parts: std::vec::IntoIter<litchi_xlsb::package::table::Table>,
        index: usize,
    },
    Done,
}

impl<'f> Children<'f> {
    const fn new(file: &'f Xlsb, state: State) -> Self {
        Self { file, state }
    }

    /// Construct the node-view children for `node`.
    pub(super) fn view(node: &Node<'f>, options: &ViewOptions) -> Result<Self> {
        let file = node.file();
        let state = match node.item() {
            Item::Workbook => State::Workbook {
                next: 0,
                count: super::nav::tab_count(file)?,
                sheets: true,
                names: false,
                name: 0,
            },
            Item::Sheet(sheet) => sheet_items(
                sheet,
                Want {
                    columns: false,
                    rows: true,
                    cells: false,
                    tables: false,
                },
                true,
            )?,
            Item::Row(row) => State::Cells {
                sheet: node.sheet_position(),
                grid: row.grid.clone(),
                cursor: Box::new(Cursor::new(row_rect(row.index)?, row.bulk)),
            },
            Item::Column(column) => State::Cells {
                sheet: node.sheet_position(),
                grid: column.grid.clone(),
                cursor: Box::new(Cursor::new(column_rect(column.index)?, false)),
            },
            Item::Range(range) => match options.dense {
                Some(budget) => {
                    let cells = u64::from(range.rect.rows()) * u64::from(range.rect.columns());
                    if cells > u64::from(budget) {
                        return Err(dense_too_large(cells, budget));
                    }
                    let start = range.rect.start();
                    State::Dense {
                        sheet: node.sheet_position(),
                        grid: range.grid.clone(),
                        rect: range.rect,
                        next: Some((start.row().get(), start.column().get())),
                    }
                },
                None => State::Cells {
                    sheet: node.sheet_position(),
                    grid: range.grid.clone(),
                    cursor: Box::new(Cursor::new(range.rect, false)),
                },
            },
            Item::Chart(chart) => State::Charts {
                sheet: node.sheet_position(),
                items: super::chart::parts(chart, true, true)?.into_iter(),
            },
            Item::Cell(_)
            | Item::Table(_)
            | Item::NamedRange(_)
            | Item::Feature(_)
            | Item::Series(_)
            | Item::Axis(_) => State::Done,
        };
        Ok(Self::new(file, state))
    }

    /// The query walker's children of `node` for a walk that is likely to
    /// stop early: a source-backed worksheet whose cells (and tables) alone
    /// are wanted lists its cells through the leaf's cell cursor; everything
    /// else as [`Self::walk`]. A worksheet the cursor refuses before its
    /// first cell is read as [`Self::walk`] reads it.
    pub(super) fn walk_early(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        if let Item::Sheet(sheet) = node.item()
            && let SheetHandle::Source { sheet: leaf, .. } = &sheet.handle
            && sheet.handle.is_grid()
            && kinds.contains(Kind::Cell)
            && !kinds.contains(Kind::Row)
            && !kinds.contains(Kind::Column)
            && let Ok(cursor) = leaf.cell_cursor()
        {
            return Ok(Self::new(
                node.file(),
                State::Stream(Box::new(Stream {
                    sheet: sheet.position,
                    handle: sheet.handle.clone(),
                    cursor,
                    current: None,
                    yielded: false,
                    tables: kinds.contains(Kind::Table),
                })),
            ));
        }
        Self::walk(node, kinds)
    }

    /// Construct query-walk children for the requested kinds.
    pub(super) fn walk(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let file = node.file();
        let cells = kinds.contains(Kind::Cell);
        let state = match node.item() {
            Item::Workbook => State::Workbook {
                next: 0,
                count: super::nav::tab_count(file)?,
                sheets: kinds.contains(Kind::Sheet),
                names: kinds.contains(Kind::NamedRange),
                name: 0,
            },
            Item::Sheet(sheet)
                if super::features::KINDS
                    .iter()
                    .any(|kind| kinds.contains(*kind))
                    && !(cells
                        || kinds.contains(Kind::Column)
                        || kinds.contains(Kind::Row)
                        || kinds.contains(Kind::Table)) =>
            {
                // A walk scoped to worksheet features lists them (record 0322).
                let mut items = Vec::new();
                for kind in super::features::KINDS {
                    if kinds.contains(kind) {
                        items.extend(super::features::list(file, sheet, kind)?);
                    }
                }
                State::Features {
                    sheet: sheet.position,
                    items: items.into_iter(),
                }
            },
            Item::Sheet(sheet)
                if kinds.contains(Kind::Chart)
                    && !(cells
                        || kinds.contains(Kind::Column)
                        || kinds.contains(Kind::Row)
                        || kinds.contains(Kind::Table)) =>
            {
                // A walk scoped to charts lists them; a mixed walk keeps the
                // grid phases (charts are reached by path there).
                State::Charts {
                    sheet: sheet.position,
                    items: super::chart::list(sheet)?
                        .into_iter()
                        .map(|chart| (Kind::Chart, chart))
                        .collect::<Vec<_>>()
                        .into_iter(),
                }
            },
            Item::Chart(chart) => State::Charts {
                sheet: node.sheet_position(),
                items: super::chart::parts(
                    chart,
                    kinds.contains(Kind::Series),
                    kinds.contains(Kind::Axis),
                )?
                .into_iter(),
            },
            Item::Sheet(sheet) => sheet_items(
                sheet,
                Want {
                    columns: kinds.contains(Kind::Column),
                    rows: kinds.contains(Kind::Row),
                    cells,
                    tables: kinds.contains(Kind::Table),
                },
                false,
            )?,
            Item::Row(row) if cells => State::Cells {
                sheet: node.sheet_position(),
                grid: row.grid.clone(),
                cursor: Box::new(Cursor::new(row_rect(row.index)?, false)),
            },
            Item::Column(column) if cells => State::Cells {
                sheet: node.sheet_position(),
                grid: column.grid.clone(),
                cursor: Box::new(Cursor::new(column_rect(column.index)?, false)),
            },
            Item::Range(range) if cells => State::Cells {
                sheet: node.sheet_position(),
                grid: range.grid.clone(),
                cursor: Box::new(Cursor::new(range.rect, false)),
            },
            _ => State::Done,
        };
        Ok(Self::new(file, state))
    }

    pub(super) fn count(node: &Node<'f>, options: &ViewOptions) -> Result<usize> {
        let mut count = 0_usize;
        for child in Self::view(node, options)? {
            child?;
            count = count.saturating_add(1);
        }
        Ok(count)
    }

    fn step(&mut self) -> Result<Option<Node<'f>>> {
        let file = self.file;
        loop {
            match &mut self.state {
                State::Done => return Ok(None),
                State::Features { sheet, items } => {
                    return Ok(items
                        .next()
                        .map(|feature| Node::new_feature(file, *sheet, feature)));
                },
                State::Charts { sheet, items } => {
                    return Ok(items
                        .next()
                        .map(|(kind, chart)| Node::new_chart(file, *sheet, kind, chart)));
                },
                State::Workbook {
                    next,
                    count,
                    sheets,
                    names,
                    name,
                } => {
                    if *sheets && *next < *count {
                        let position = *next;
                        *next += 1;
                        let Some(handle) = SheetHandle::at(file, position)? else {
                            continue;
                        };
                        return Ok(Some(Node::new_sheet(
                            file,
                            Sheet::new(position.into(), handle),
                        )));
                    }
                    if *names && *name < defined_names(file)?.len() {
                        let index = *name;
                        *name += 1;
                        return Ok(Some(Node::new_named_range(file, index)));
                    }
                    return Ok(None);
                },
                State::Sheet(items) => return items.next(file),
                State::Stream(stream) => match stream.next(file) {
                    Ok(Some(node)) => return Ok(Some(node)),
                    Ok(None) => {
                        self.state = if stream.tables {
                            State::Tables {
                                sheet: stream.sheet,
                                parts: stream.handle.tables()?.into_iter(),
                                index: 0,
                            }
                        } else {
                            State::Done
                        };
                    },
                    // Refused before the first cell: read the worksheet as
                    // the ordinary walk does, which owns the refusal (or
                    // the cached-value fallback).
                    Err(_) if !stream.yielded => {
                        let sheet = Sheet::new(stream.sheet.into(), stream.handle.clone());
                        self.state = sheet_items(
                            &sheet,
                            Want {
                                columns: false,
                                rows: false,
                                cells: true,
                                tables: stream.tables,
                            },
                            false,
                        )?;
                    },
                    Err(error) => return Err(error),
                },
                State::Tables {
                    sheet,
                    parts,
                    index,
                } => {
                    let Some(part) = parts.next() else {
                        self.state = State::Done;
                        return Ok(None);
                    };
                    let position = *index;
                    *index += 1;
                    return Ok(Some(Node::new_table(file, *sheet, position, part)));
                },
                State::Cells {
                    sheet,
                    grid,
                    cursor,
                } => match cursor.next(grid)? {
                    Some(at) => {
                        return Ok(Some(Node::new_cell(
                            file,
                            *sheet,
                            at,
                            CellSource::Grid(grid.clone()),
                        )));
                    },
                    None => {
                        self.state = State::Done;
                        return Ok(None);
                    },
                },
                State::Dense {
                    sheet,
                    grid,
                    rect,
                    next,
                } => {
                    let Some(at) = *next else {
                        return Ok(None);
                    };
                    *next = successor(*rect, at)?;
                    let source = if grid.record(at)?.is_some() {
                        CellSource::Grid(grid.clone())
                    } else {
                        CellSource::Stub
                    };
                    return Ok(Some(Node::new_cell(file, *sheet, at, source)));
                },
            }
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if matches!(self.state, State::Done) {
            return None;
        }
        let file = self.file;
        // A source-backed operation is fenced before and after every yielded
        // child. An error terminates the iterator, so it is fused and cannot
        // expose a stale suffix after a failed read.
        match super::nav::finish(file, || self.step()) {
            Ok(Some(node)) => Some(Ok(node)),
            Ok(None) => {
                self.state = State::Done;
                None
            },
            Err(error) => {
                self.state = State::Done;
                Some(Err(error))
            },
        }
    }
}

impl std::iter::FusedIterator for Children<'_> {}

/// A source-backed worksheet's cells read through the leaf's cell cursor,
/// with the node grid of the leaf grid the cursor is reading.
struct Stream {
    sheet: usize,
    handle: SheetHandle,
    cursor: litchi_xlsb::sheet::CellCursor,
    /// The node grid of the cursor's current grid, keyed by whether it is
    /// the complete grid and its cell count.
    current: Option<((bool, usize), Grid)>,
    yielded: bool,
    tables: bool,
}

impl Stream {
    fn next<'f>(&mut self, file: &'f Xlsb) -> Result<Option<Node<'f>>> {
        let Some(at) = self.cursor.next().transpose()? else {
            return Ok(None);
        };
        let Some(leaf) = self.cursor.grid() else {
            return Err(Error::new(
                Code::InternalError,
                "the worksheet cell cursor returned a cell without its grid",
            ));
        };
        let key = (self.cursor.is_complete(), leaf.stored_cell_count());
        let grid = match &self.current {
            Some((held, grid)) if *held == key => grid.clone(),
            _ => {
                let grid = Grid::streamed(leaf.clone(), self.handle.clone(), key.0);
                self.current = Some((key, grid.clone()));
                grid
            },
        };
        self.yielded = true;
        Ok(Some(Node::new_cell(
            file,
            self.sheet,
            at,
            CellSource::Grid(grid),
        )))
    }
}

fn sheet_items(sheet: &Sheet, want: Want, view: bool) -> Result<State> {
    if !(want.columns || want.rows || want.cells || want.tables) {
        return Ok(State::Done);
    }
    if !sheet.handle.is_grid() {
        return Ok(State::Done);
    }
    if want.tables && !(want.columns || want.rows || want.cells) {
        return Ok(State::Tables {
            sheet: sheet.position,
            parts: sheet.handle.tables()?.into_iter(),
            index: 0,
        });
    }
    Ok(State::Sheet(Box::new(SheetItems {
        sheet: sheet.position,
        grid: sheet.grid()?.clone(),
        handle: sheet.handle.clone(),
        want,
        view,
        phase: Phase::Start,
    })))
}

impl SheetItems {
    fn next<'f>(&mut self, file: &'f Xlsb) -> Result<Option<Node<'f>>> {
        loop {
            let Self {
                sheet,
                grid,
                handle,
                want,
                view,
                phase,
            } = self;
            match phase {
                Phase::Start => {
                    *phase = if want.columns {
                        Phase::Columns {
                            ranges: column_ranges(grid)?,
                            range: 0,
                            next: 0,
                        }
                    } else {
                        rows_phase(grid, handle, *want)?
                    };
                },
                Phase::Columns {
                    ranges,
                    range,
                    next,
                } => {
                    let Some(&(first, last)) = ranges.get(*range) else {
                        *phase = rows_phase(grid, handle, *want)?;
                        continue;
                    };
                    let index = (*next).max(first);
                    if index <= last {
                        *next = index.saturating_add(1);
                        return Ok(Some(Node::new_column(file, *sheet, grid.clone(), index)));
                    }
                    *range += 1;
                },
                Phase::Rows { rows, next, cells } => {
                    if let Some(cursor) = cells {
                        if let Some(at) = cursor.next(grid)? {
                            return Ok(Some(Node::new_cell(
                                file,
                                *sheet,
                                at,
                                CellSource::Grid(grid.clone()),
                            )));
                        }
                        *cells = None;
                    }
                    let Some(&index) = rows.get(*next) else {
                        *phase = tables_phase(handle, *want)?;
                        continue;
                    };
                    *next += 1;
                    if *view && !listed(grid, index)? {
                        continue;
                    }
                    if want.cells {
                        *cells = Some(Cursor::new(row_rect(index)?, *view));
                    }
                    if want.rows {
                        let grid = grid.clone();
                        return Ok(Some(if *view {
                            Node::new_listed_row(file, *sheet, grid, index)
                        } else {
                            Node::new_row(file, *sheet, grid, index)
                        }));
                    }
                },
                Phase::Cells(cursor) => {
                    if let Some(at) = cursor.next(grid)? {
                        return Ok(Some(Node::new_cell(
                            file,
                            *sheet,
                            at,
                            CellSource::Grid(grid.clone()),
                        )));
                    }
                    *phase = tables_phase(handle, *want)?;
                },
                Phase::Tables { parts, index } => {
                    let Some(part) = parts.next() else {
                        *phase = Phase::Done;
                        continue;
                    };
                    let position = *index;
                    *index += 1;
                    return Ok(Some(Node::new_table(file, *sheet, position, part)));
                },
                Phase::Done => return Ok(None),
            }
        }
    }
}

fn rows_phase(grid: &Grid, handle: &SheetHandle, want: Want) -> Result<Phase> {
    if want.rows {
        let rows = row_indices(grid)?;
        return Ok(Phase::Rows {
            rows,
            next: 0,
            cells: None,
        });
    }
    if want.cells {
        if let Some(extent) = grid.extents()? {
            return Ok(Phase::Cells(Cursor::new(extent, false)));
        }
    }
    tables_phase(handle, want)
}

fn tables_phase(handle: &SheetHandle, want: Want) -> Result<Phase> {
    if !want.tables {
        return Ok(Phase::Done);
    }
    Ok(Phase::Tables {
        parts: handle.tables()?.into_iter(),
        index: 0,
    })
}

/// Return logical column records as contiguous inclusive ranges. A BIFF12
/// worksheet has explicit column metadata as well as cell columns; the node
/// bridge folds both into this allocation-free-on-the-common-path iterator.
fn column_ranges(grid: &Grid) -> Result<Vec<(u32, u32)>> {
    let mut ranges: Vec<(u32, u32)> = Vec::new();
    for (first, last) in grid.column_ranges()? {
        match ranges.last_mut() {
            Some((_, previous)) if first <= previous.saturating_add(1) => {
                *previous = (*previous).max(last)
            },
            _ => ranges.push((first, last)),
        }
    }
    Ok(ranges)
}

/// Collect occupied and explicitly formatted rows once for a worksheet walk.
fn row_indices(grid: &Grid) -> Result<Vec<u32>> {
    let mut rows = grid.row_indices()?;
    rows.sort_unstable();
    rows.dedup();
    Ok(rows)
}

/// Whether OfficeCLI's bulk listing shows a row or cell.
fn listed(grid: &Grid, index: u32) -> Result<bool> {
    if grid.row_has_meaningful_attrs(index)? {
        return Ok(true);
    }
    Ok(grid
        .cells(row_rect(index)?)?
        .any(|(_, cell)| has_content(cell)))
}

/// The complete logical row rectangle.
pub(super) fn row_rect(index: u32) -> Result<Rect> {
    let (_, columns) = Rect::ALL.end();
    Rect::at(
        index,
        0,
        index.checked_add(1).ok_or_else(|| {
            Error::new(Code::OutOfRange, "worksheet row number is outside the grid")
        })?,
        columns,
    )
    .map_err(range)
}

/// The complete logical column rectangle.
pub(super) fn column_rect(index: u32) -> Result<Rect> {
    let (rows, _) = Rect::ALL.end();
    Rect::at(
        0,
        index,
        rows,
        index.checked_add(1).ok_or_else(|| {
            Error::new(
                Code::OutOfRange,
                "worksheet column number is outside the grid",
            )
        })?,
    )
    .map_err(range)
}

fn successor(rect: Rect, at: Address) -> Result<Option<Address>> {
    let (end_row, end_column) = rect.end();
    let (row, column) = at;
    if column.saturating_add(1) < end_column {
        return Ok(Some((row, column + 1)));
    }
    if row.saturating_add(1) < end_row {
        return Ok(Some((row + 1, rect.start().column().get())));
    }
    Ok(None)
}

fn has_content(cell: &litchi_xlsb::Cell) -> bool {
    !matches!(cell.value(), CellValue::Empty)
}

/// A bounded row-major cursor over stored records in one rectangle.
pub(super) struct Cursor {
    rect: Rect,
    resume: Option<Address>,
    buffer: [Address; BATCH],
    len: usize,
    pos: usize,
    content_only: bool,
}

impl Cursor {
    pub(super) fn new(rect: Rect, content_only: bool) -> Self {
        let start = rect.start();
        let address = (start.row().get(), start.column().get());
        Self {
            rect,
            resume: Some(address),
            buffer: [address; BATCH],
            len: 0,
            pos: 0,
            content_only,
        }
    }

    pub(super) fn next(&mut self, grid: &Grid) -> Result<Option<Address>> {
        if self.pos == self.len {
            self.refill(grid)?;
            if self.len == 0 {
                return Ok(None);
            }
        }
        let at = self.buffer[self.pos];
        self.pos += 1;
        Ok(Some(at))
    }

    fn refill(&mut self, grid: &Grid) -> Result<()> {
        self.pos = 0;
        self.len = 0;
        let Some(start) = self.resume.take() else {
            return Ok(());
        };
        let (end_row, end_column) = self.rect.end();
        let first_column = self.rect.start().column().get();
        let (row, column) = start;
        let spans = if column == first_column {
            [
                Some(Rect::at(row, column, end_row, end_column).map_err(range)?),
                None,
            ]
        } else {
            let below = if row.saturating_add(1) < end_row {
                Some(Rect::at(row + 1, first_column, end_row, end_column).map_err(range)?)
            } else {
                None
            };
            [
                Some(Rect::at(row, column, row + 1, end_column).map_err(range)?),
                below,
            ]
        };
        for span in spans.into_iter().flatten() {
            for (at, cell) in grid.cells(span)? {
                if self.content_only && !has_content(cell) {
                    continue;
                }
                if self.len == BATCH {
                    self.resume = Some(at);
                    return Ok(());
                }
                self.buffer[self.len] = at;
                self.len += 1;
            }
        }
        Ok(())
    }
}

/// Convert a checked coordinate failure from the leaf bridge.
pub(super) fn coordinate<E>(error: E) -> Error
where
    litchi_xlsb::package::PackageError: From<E>,
{
    Error::from(litchi_xlsb::package::PackageError::from(error))
}

/// Convert a checked range failure from the leaf bridge.
pub(super) fn range<E>(error: E) -> Error
where
    litchi_xlsb::package::PackageError: From<E>,
{
    Error::from(litchi_xlsb::package::PackageError::from(error))
}

fn dense_too_large(cells: u64, budget: u32) -> Error {
    Error::new(
        Code::OutputTooLarge,
        format!("the range has {cells} cells, more than the dense view budget of {budget}"),
    )
    .with_suggestion("raise the `dense` budget, view a smaller range, or list stored cells only")
}
