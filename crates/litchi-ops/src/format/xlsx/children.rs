//! The children of XLSX nodes, read lazily (record 0010 §2.5, §5.3).
//!
//! One iterator serves two readers:
//!
//! - **The node view** lists OfficeCLI's shape: a workbook's sheets, a
//!   worksheet's rows, a row's cells, a column's cells, a range's cells. Like
//!   OfficeCLI's bulk listing (`GetSheetChildNodes`), a sheet lists the rows
//!   that hold content or carry row formatting, and such a row lists the
//!   records that hold content; a row resolved by path or matched by a query
//!   lists all its records, as OfficeCLI's row path does. A range lists its
//!   stored records only, unless the view asks for dense stubs (record 0010
//!   Q9).
//! - **The query walker** asks for the kinds it needs under the path tree of
//!   the spec table: a workbook's sheets and defined names; a worksheet's
//!   columns, then its rows and cells in row-major order, then its tables;
//!   the cells of a row, column or range used as a query scope.
//!
//! Cells are listed from the worksheet's pinned grid: a worksheet's cells by
//! walking its records in order, the cells of a row, column or range through
//! a cursor that copies at most [`BATCH`] coordinates, each with its
//! record's position, at a time out of one borrowed range scan. So the
//! iterator owns no borrow, a cell costs one reference-count increment, its
//! node reads its record again in O(1) (record 0262), and a walk that stops
//! early has scanned at most one batch past its last match. Row numbers are
//! collected once per worksheet walk that lists rows, because the leaf's row
//! traversal cannot resume mid-way.

use litchi_xlsx::workbook::worksheet::TablePart;
use litchi_xlsx::{self as xlsx, Address, Rect};

use super::Xlsx;
use super::node::{
    CellSource, ChartItem, Grid, Item, Node, Sheet, SheetHandle, defined_names, has_content,
};
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::node::ViewOptions;
use crate::path::navigate::KindSet;

/// How many coordinates one cursor refill copies.
const BATCH: usize = 64;

/// The children of one XLSX node, in document order.
pub(crate) struct Children<'f> {
    file: &'f Xlsx,
    state: State,
}

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
    Cells {
        sheet: usize,
        grid: Grid,
        /// Boxed: a cursor carries its batch of coordinates.
        cursor: Box<Cursor>,
    },
    Dense {
        sheet: usize,
        grid: Grid,
        rect: Rect,
        next: Option<Address>,
    },
    /// A worksheet's features (record 0320).
    Features {
        sheet: usize,
        items: std::vec::IntoIter<super::features::Feature>,
    },
    /// A worksheet's charts, or a chart's series and axes (record 0270).
    Charts {
        sheet: usize,
        items: std::vec::IntoIter<(Kind, ChartItem)>,
    },
    /// A cell's rich-text runs, for a walk that wants runs (record 0338).
    Runs {
        sheet: usize,
        at: Address,
        rich: std::sync::Arc<xlsx::RichText>,
        next: usize,
    },
}

/// What a worksheet lists.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Want {
    columns: bool,
    rows: bool,
    cells: bool,
    tables: bool,
}

/// The items of one worksheet, phase by phase.
struct SheetItems {
    sheet: usize,
    /// The worksheet's grid; `None` while the leaf's cell cursor lists the
    /// cells (it materializes the grid itself if the walk reads on).
    grid: Option<Grid>,
    handle: SheetHandle,
    want: Want,
    /// The node view's listing: only rows with content or formatting, and
    /// only cells with content.
    view: bool,
    phase: Phase,
}

#[allow(
    clippy::large_enum_variant,
    reason = "the phases live in the boxed items of one worksheet walk; a row's cell cursor keeps \
              its batch inline so listing allocates nothing per row"
)]
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
    /// Every stored record, in the grid's record order, which is row-major:
    /// the whole sheet's cells without a range scan.
    Records {
        next: usize,
    },
    /// Every stored record through the leaf's cell cursor, which parses only
    /// the leading rows a walk that stops early visits (record 0267).
    Stream(Box<Stream>),
    Tables {
        parts: std::vec::IntoIter<TablePart>,
        index: usize,
    },
    Done,
}

/// A worksheet's records read through the leaf's cell cursor, with the
/// node grid of the leaf grid the cursor is reading.
struct Stream {
    cursor: xlsx::workbook::source::CellCursor,
    sheet: xlsx::SourceWorksheet,
    /// The node grid of the cursor's current grid, keyed by what tells two
    /// of the cursor's grids apart: whether it is the complete grid, and its
    /// record count (a longer prefix that adds no record holds the same
    /// records).
    current: Option<((bool, usize), Grid)>,
}

impl Stream {
    fn next<'f>(&mut self, file: &'f Xlsx, sheet: usize) -> Result<Option<Node<'f>>> {
        let Some((at, index)) = self.cursor.next().transpose()? else {
            return Ok(None);
        };
        let Some(leaf) = self.cursor.grid() else {
            return Err(Error::new(
                Code::InternalError,
                "the worksheet cell cursor returned a record without its grid",
            ));
        };
        let key = (self.cursor.is_complete(), leaf.stored_cell_count());
        let grid = match &self.current {
            Some((held, grid)) if *held == key => grid,
            _ => {
                let grid = Grid::streamed(leaf.clone(), self.sheet.clone(), key.0);
                &self.current.insert((key, grid)).1
            },
        };
        Ok(Some(listed_cell(file, sheet, grid, at, index)))
    }
}

impl<'f> Children<'f> {
    const fn new(file: &'f Xlsx, state: State) -> Self {
        Self { file, state }
    }

    /// The query walker's children of `node` for a walk that is likely to
    /// stop early ([`crate::query::inspect::Inspect::early_children`]): a
    /// source-backed worksheet whose cells alone are wanted lists them
    /// through the leaf's cell cursor; everything else as [`Self::walk`].
    ///
    /// # Errors
    ///
    /// As [`Self::walk`].
    pub(super) fn walk_early(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let file = node.file();
        if let Item::Sheet(sheet) = node.item()
            && let SheetHandle::Source { sheet: leaf } = &sheet.handle
            && sheet.handle.is_grid()
            && kinds.contains(Kind::Cell)
            && !kinds.contains(Kind::Row)
            && !kinds.contains(Kind::Column)
        {
            let stream = Stream {
                cursor: leaf.cell_cursor()?,
                sheet: leaf.clone(),
                current: None,
            };
            return Ok(Self::new(
                file,
                State::Sheet(Box::new(SheetItems {
                    sheet: sheet.position,
                    grid: None,
                    handle: sheet.handle.clone(),
                    want: Want {
                        columns: false,
                        rows: false,
                        cells: true,
                        tables: kinds.contains(Kind::Table),
                    },
                    view: false,
                    phase: Phase::Stream(Box::new(stream)),
                })),
            ));
        }
        Self::walk(node, kinds)
    }

    /// The node view's children of `node`: OfficeCLI's listing.
    ///
    /// # Errors
    ///
    /// The leaf's error when the node's worksheet cannot be read, and
    /// `output_too_large` for a dense range larger than its cell budget.
    pub(super) fn view(node: &Node<'f>, options: &ViewOptions) -> Result<Self> {
        let file = node.file();
        let state = match node.item() {
            Item::Workbook => State::Workbook {
                next: 0,
                count: file.0.owner.len(),
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
                    State::Dense {
                        sheet: node.sheet_position(),
                        grid: range.grid.clone(),
                        rect: range.rect,
                        next: Some(range.rect.start()),
                    }
                },
                None => State::Cells {
                    sheet: node.sheet_position(),
                    grid: range.grid.clone(),
                    cursor: Box::new(Cursor::new(range.rect, false)),
                },
            },
            Item::Chart(chart) => chart_parts(node, chart, true, true)?,
            Item::Cell(_)
            | Item::Table(_)
            | Item::NamedRange(_)
            | Item::Series(_)
            | Item::Axis(_)
            | Item::Feature(_)
            | Item::Run(_) => State::Done,
        };
        Ok(Self::new(file, state))
    }

    /// The query walker's children of `node` whose kinds are in `kinds`.
    ///
    /// # Errors
    ///
    /// The leaf's error when the node's worksheet cannot be read.
    pub(super) fn walk(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let file = node.file();
        let cells = kinds.contains(Kind::Cell);
        let state = match node.item() {
            Item::Workbook => State::Workbook {
                next: 0,
                count: file.0.owner.len(),
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
                        || kinds.contains(Kind::Table)
                        || kinds.contains(Kind::Chart)) =>
            {
                // A walk scoped to worksheet features lists them (record 0320).
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
                // A walk scoped to charts lists them; a mixed walk keeps
                // the grid phases (charts are reached by path there).
                let count = file.sheet_chart_count(sheet.position)?;
                let mut items = Vec::with_capacity(count);
                for index in 0..count {
                    if let Some(part) = file.sheet_chart(sheet.position, index)? {
                        items.push((
                            Kind::Chart,
                            ChartItem {
                                chart: u32::try_from(index).unwrap_or(u32::MAX),
                                part,
                                child: 0,
                            },
                        ));
                    }
                }
                State::Charts {
                    sheet: sheet.position,
                    items: items.into_iter(),
                }
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
            Item::Chart(chart) => chart_parts(
                node,
                chart,
                kinds.contains(Kind::Series),
                kinds.contains(Kind::Axis),
            )?,
            // Rows, columns and ranges hold cells when a query is scoped to
            // them; the spec's walk tree does not enter them otherwise.
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
            // A cell's runs, only for a walk that wants them (record 0338).
            Item::Cell(_) if kinds.contains(Kind::Run) => match super::runs::cell_runs(node)? {
                Some((sheet, at, rich)) => State::Runs {
                    sheet,
                    at,
                    rich,
                    next: 0,
                },
                None => State::Done,
            },
            _ => State::Done,
        };
        Ok(Self::new(file, state))
    }

    /// The number of children the node view lists for `node`.
    ///
    /// # Errors
    ///
    /// As for [`Children::view`], and the leaf's error while counting.
    pub(super) fn count(node: &Node<'f>, options: &ViewOptions) -> Result<usize> {
        let mut count = 0_usize;
        for child in Self::view(node, options)? {
            child?;
            count += 1;
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
                State::Runs {
                    sheet,
                    at,
                    rich,
                    next,
                } => {
                    if *next >= rich.runs().len() {
                        return Ok(None);
                    }
                    *next += 1;
                    return Ok(Some(Node::new_run(
                        file,
                        *sheet,
                        *at,
                        *next - 1,
                        std::sync::Arc::clone(rich),
                    )));
                },
                State::Charts { sheet, items } => {
                    return Ok(items.next().map(|(kind, item)| match kind {
                        Kind::Chart => {
                            let index = item.chart as usize;
                            Node::new_chart(file, *sheet, index, item.part)
                        },
                        kind => Node::new_chart_part(file, *sheet, kind, item),
                    }));
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
                            Sheet::new(file, position, handle),
                        )));
                    }
                    if *names && *name < defined_names(file).len() {
                        let index = *name;
                        *name += 1;
                        return Ok(Some(Node::new_named_range(file, index)));
                    }
                    return Ok(None);
                },
                State::Sheet(items) => return items.next(file),
                State::Cells {
                    sheet,
                    grid,
                    cursor,
                } => {
                    return Ok(cursor
                        .next(grid)?
                        .map(|(at, index)| listed_cell(file, *sheet, grid, at, index)));
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
        match file.0.owner.finish(|| self.step()) {
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

/// The items of a worksheet; nothing for a chart or dialog sheet, which has
/// no grid.
fn sheet_items(sheet: &Sheet, want: Want, view: bool) -> Result<State> {
    if !sheet.handle.is_grid() || !(want.columns || want.rows || want.cells || want.tables) {
        return Ok(State::Done);
    }
    Ok(State::Sheet(Box::new(SheetItems {
        sheet: sheet.position,
        grid: Some(sheet.grid()?.clone()),
        handle: sheet.handle.clone(),
        want,
        view,
        phase: Phase::Start,
    })))
}

impl SheetItems {
    fn next<'f>(&mut self, file: &'f Xlsx) -> Result<Option<Node<'f>>> {
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
                Phase::Stream(stream) => {
                    if let Some(node) = stream.next(file, *sheet)? {
                        return Ok(Some(node));
                    }
                    *phase = tables_phase(handle, *want)?;
                    continue;
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
                Phase::Start
                | Phase::Columns { .. }
                | Phase::Rows { .. }
                | Phase::Records { .. } => {},
            }
            let Some(grid) = grid.as_ref() else {
                return Err(Error::new(
                    Code::InternalError,
                    "a worksheet listing lost its grid",
                ));
            };
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
                    if let Some(&(first, last)) = ranges.get(*range) {
                        let index = (*next).max(first);
                        if index <= last {
                            *next = index + 1;
                            return Ok(Some(Node::new_column(file, *sheet, grid.clone(), index)));
                        }
                        *range += 1;
                        continue;
                    }
                    *phase = rows_phase(grid, handle, *want)?;
                },
                Phase::Rows { rows, next, cells } => {
                    if let Some(cursor) = cells {
                        if let Some((at, index)) = cursor.next(grid)? {
                            return Ok(Some(listed_cell(file, *sheet, grid, at, index)));
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
                Phase::Records { next } => {
                    let index = xlsx::RecordIndex::new(*next);
                    if let Some((at, _)) = grid.record_at(index)? {
                        *next += 1;
                        return Ok(Some(listed_cell(file, *sheet, grid, at, index)));
                    }
                    *phase = tables_phase(handle, *want)?;
                },
                // Read above, without the grid.
                Phase::Stream(_) | Phase::Tables { .. } | Phase::Done => {},
            }
        }
    }
}

/// The rows-and-cells phase: rows (each followed by its cells when both are
/// wanted), or the stored cells alone, or straight to the tables.
fn rows_phase(grid: &Grid, handle: &SheetHandle, want: Want) -> Result<Phase> {
    if want.rows {
        let rows = grid.rows()?.map(|row| row.index().get()).collect();
        return Ok(Phase::Rows {
            rows,
            next: 0,
            cells: None,
        });
    }
    if want.cells && grid.extents()?.stored().is_some() {
        // The stored extent holds every record, so the cells are the records
        // in order.
        return Ok(Phase::Records { next: 0 });
    }
    tables_phase(handle, want)
}

/// The node of a cell a cursor listed: its record, read again by its
/// position when that fits a node's index, else by its coordinate.
fn listed_cell<'f>(
    file: &'f Xlsx,
    sheet: usize,
    grid: &Grid,
    at: Address,
    index: xlsx::RecordIndex,
) -> Node<'f> {
    let grid = grid.clone();
    let source = match u32::try_from(index.get()) {
        Ok(index) => CellSource::Record { grid, index },
        Err(_) => CellSource::Grid(grid),
    };
    Node::new_cell(file, sheet, at, source)
}

/// The tables phase, or done.
fn tables_phase(handle: &SheetHandle, want: Want) -> Result<Phase> {
    if !want.tables {
        return Ok(Phase::Done);
    }
    Ok(Phase::Tables {
        parts: handle.tables()?.into_iter(),
        index: 0,
    })
}

/// The worksheet's explicit column records, as inclusive zero-based ranges
/// of logical columns.
fn column_ranges(grid: &Grid) -> Result<Vec<(u32, u32)>> {
    let mut ranges: Vec<(u32, u32)> = Vec::new();
    for column in grid.columns()? {
        let index = column.index().get();
        match ranges.last_mut() {
            Some((_, last)) if index == *last + 1 => *last = index,
            _ => ranges.push((index, index)),
        }
    }
    Ok(ranges)
}

/// Whether OfficeCLI's bulk listing shows a row: it holds a record with
/// content, or its record carries an attribute beyond its number
/// (`RowHasMeaningfulAttrs`; the descent is `x14ac:dyDescent`).
pub(super) fn listed(grid: &Grid, index: u32) -> Result<bool> {
    let row = grid.row(index)?;
    if row.height().is_some()
        || row.descent().is_some()
        || row.hidden()
        || row.custom_height()
        || row.outline().get() > 0
        || row.collapsed()
        || row.thick_top()
        || row.thick_bottom()
        || row.phonetic()
        || row.custom_format()
    {
        return Ok(true);
    }
    Ok(grid
        .cells(row_rect(index)?)?
        .any(|(_, cell)| has_content(cell)))
}

/// The whole row `index`.
pub(super) fn row_rect(index: u32) -> Result<Rect> {
    let start = Address::at(index, 0).map_err(coordinate)?;
    Rect::new(start, index + 1, xlsx::ColumnIndex::LAST.get() + 1).map_err(range)
}

/// The whole column `index`.
pub(super) fn column_rect(index: u32) -> Result<Rect> {
    let start = Address::at(0, index).map_err(coordinate)?;
    Rect::new(start, xlsx::RowIndex::LAST.get() + 1, index + 1).map_err(range)
}

/// The coordinate after `at` in `rect`, row-major.
fn successor(rect: Rect, at: Address) -> Result<Option<Address>> {
    let (end_row, end_column) = rect.end();
    let (row, column) = (at.row().get(), at.column().get());
    if column + 1 < end_column {
        return Address::at(row, column + 1).map(Some).map_err(coordinate);
    }
    if row + 1 < end_row {
        return Address::at(row + 1, rect.start().column().get())
            .map(Some)
            .map_err(coordinate);
    }
    Ok(None)
}

/// A row-major cursor over the stored records of one rectangle.
///
/// Each refill runs one borrowed range scan of the grid from where the last
/// one stopped and copies up to [`BATCH`] coordinates, with their records'
/// positions, out of it.
pub(super) struct Cursor {
    rect: Rect,
    /// Where the next refill starts; `None` once the rectangle is exhausted.
    resume: Option<Address>,
    buffer: [(Address, xlsx::RecordIndex); BATCH],
    len: usize,
    pos: usize,
    /// Skip records without content (the node view's row listing).
    content_only: bool,
}

impl Cursor {
    pub(super) fn new(rect: Rect, content_only: bool) -> Self {
        Self {
            rect,
            resume: Some(rect.start()),
            buffer: [(rect.start(), xlsx::RecordIndex::new(0)); BATCH],
            len: 0,
            pos: 0,
            content_only,
        }
    }

    /// The next stored coordinate, row-major, with its record's position.
    pub(super) fn next(&mut self, grid: &Grid) -> Result<Option<(Address, xlsx::RecordIndex)>> {
        if self.pos == self.len {
            self.refill(grid)?;
            if self.len == 0 {
                return Ok(None);
            }
        }
        let entry = self.buffer[self.pos];
        self.pos += 1;
        Ok(Some(entry))
    }

    fn refill(&mut self, grid: &Grid) -> Result<()> {
        self.pos = 0;
        self.len = 0;
        let Some(start) = self.resume.take() else {
            return Ok(());
        };
        let (end_row, end_column) = self.rect.end();
        let first_column = self.rect.start().column().get();
        let row = start.row().get();
        // From the left edge, one span covers the rest of the rectangle;
        // otherwise the rest of the start row, then the rows below it.
        let spans = if start.column().get() == first_column {
            [
                Some(Rect::new(start, end_row, end_column).map_err(range)?),
                None,
            ]
        } else {
            let below = if row + 1 < end_row {
                let left = Address::at(row + 1, first_column).map_err(coordinate)?;
                Some(Rect::new(left, end_row, end_column).map_err(range)?)
            } else {
                None
            };
            [
                Some(Rect::new(start, row + 1, end_column).map_err(range)?),
                below,
            ]
        };
        for span in spans.into_iter().flatten() {
            for (index, at, cell) in grid.cells(span)?.indexed() {
                if self.content_only && !has_content(cell) {
                    continue;
                }
                if self.len == BATCH {
                    self.resume = Some(at);
                    return Ok(());
                }
                self.buffer[self.len] = (at, index);
                self.len += 1;
            }
        }
        Ok(())
    }
}

/// A coordinate the grid refuses, as the leaf classifies it.
pub(super) fn coordinate<E>(error: E) -> Error
where
    xlsx::Error: From<E>,
{
    Error::from(xlsx::Error::from(error))
}

/// A rectangle the grid refuses, as the leaf classifies it.
pub(super) fn range<E>(error: E) -> Error
where
    xlsx::Error: From<E>,
{
    Error::from(xlsx::Error::from(error))
}

fn dense_too_large(cells: u64, budget: u32) -> Error {
    Error::new(
        Code::OutputTooLarge,
        format!("the range has {cells} cells, more than the dense view budget of {budget}"),
    )
    .with_suggestion("raise the `dense` budget, view a smaller range, or list stored cells only")
}

/// A chart's series and axes, as wanted.
fn chart_parts(node: &Node<'_>, chart: &ChartItem, series: bool, axes: bool) -> Result<State> {
    let summary = chart.summary()?;
    let mut items = Vec::new();
    if series {
        items
            .extend((0..summary.series.len()).map(|index| (Kind::Series, chart.with_child(index))));
    }
    if axes {
        items.extend((0..summary.axes.len()).map(|index| (Kind::Axis, chart.with_child(index))));
    }
    Ok(State::Charts {
        sheet: node.sheet_position(),
        items: items.into_iter(),
    })
}
