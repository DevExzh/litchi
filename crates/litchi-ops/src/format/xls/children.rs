//! Lazy sparse children for the legacy XLS node adapter.

use super::Xls;
use super::node::{CellSource, Item, Node, Sheet};
use super::owner::{CellData, Grid, SheetHandle};
use super::spec::Kind;
use crate::error::{Code, Error, Result};
use crate::node::ViewOptions;
use crate::path::navigate::KindSet;

/// Children are cursor-backed. The grid itself is shared by `Arc`; only the
/// selected coordinate list is retained by a row/range cursor.
pub(crate) struct Children<'f> {
    file: &'f Xls,
    state: State<'f>,
    /// View children fence each step; query-walk children run inside the
    /// walker's own per-advance fence.
    fenced: bool,
}

enum State<'f> {
    EagerCells {
        sheet: usize,
        cells: litchi_xls::worksheet::StoredCells<'f>,
    },
    /// A positional worksheet's cells through the leaf's cell cursor
    /// (record 0267).
    Stream(Box<Stream<'f>>),
    Done,
    Sheets {
        next: usize,
        count: usize,
        /// Continue with the defined names once the tabs are exhausted.
        names: bool,
    },
    Names {
        next: usize,
        count: usize,
    },
    Coordinates {
        sheet: usize,
        grid: Grid,
        coords: std::vec::IntoIter<(u32, u32)>,
    },
    SheetItems(SheetItems),
    Rows {
        sheet: usize,
        grid: Grid,
        rows: std::vec::IntoIter<u32>,
    },
    /// A worksheet's features (record 0327).
    Features {
        sheet: usize,
        items: std::vec::IntoIter<super::features::Feature>,
    },
}

/// A positional worksheet's cells read through the leaf's cell cursor,
/// which scans only as far as a walk that stops early reads.
///
/// The leaf gives a cursor only for a worksheet whose stored positions
/// strictly ascend, so its records are the sorted grid's cells in order. A
/// record that did not ascend all the same would make the walk continue on
/// the sorted grid after the last coordinate it listed.
struct Stream<'f> {
    sheet: usize,
    handle: SheetHandle<'f>,
    cursor: litchi_xls::CellCursor,
    worksheet: usize,
    last: Option<(u32, u32)>,
    /// The sorted grid and its remaining coordinates, once the records
    /// stopped ascending.
    rest: Option<(Grid, std::vec::IntoIter<(u32, u32)>)>,
}

impl Stream<'_> {
    fn next<'f>(&mut self, file: &'f Xls) -> Result<Option<Node<'f>>> {
        loop {
            if let Some((grid, coords)) = &mut self.rest {
                return Ok(coords.next().map(|at| {
                    Node::new_cell(file, self.sheet, at, CellSource::Grid(grid.clone()))
                }));
            }
            let Some(cell) = self.cursor.next().transpose()? else {
                return Ok(None);
            };
            let at = (cell.row(), cell.column());
            if let Some(last) = self.last
                && at <= last
            {
                let grid = self.handle.grid()?;
                let coords = grid
                    .extents()
                    .map_or_else(Vec::new, |rect| {
                        grid.cells_in(rect)
                            .map(|(at, _)| at)
                            .filter(|at| *at > last)
                            .collect()
                    })
                    .into_iter();
                self.rest = Some((grid, coords));
                continue;
            }
            self.last = Some(at);
            let data = CellData::from_source(self.handle.package(), self.worksheet, &cell)?;
            return Ok(Some(Node::new_cell(
                file,
                self.sheet,
                at,
                CellSource::Direct(Some(data)),
            )));
        }
    }
}

#[derive(Clone, Copy)]
struct SheetWant {
    columns: bool,
    rows: bool,
    cells: bool,
}

impl SheetWant {
    const fn any(self) -> bool {
        self.columns || self.rows || self.cells
    }
}

struct SheetItems {
    sheet: usize,
    grid: Grid,
    want: SheetWant,
    columns: Vec<(u32, u32)>,
    column_range: usize,
    column_next: u32,
    rows: std::vec::IntoIter<u32>,
    cells: Option<std::vec::IntoIter<(u32, u32)>>,
    phase: SheetPhase,
}

#[derive(Clone, Copy)]
enum SheetPhase {
    Columns,
    Rows,
    Cells,
    Done,
}

impl SheetItems {
    fn new(sheet: usize, grid: Grid, want: SheetWant) -> Self {
        let phase = if want.columns {
            SheetPhase::Columns
        } else if want.rows {
            SheetPhase::Rows
        } else {
            SheetPhase::Cells
        };
        Self {
            sheet,
            columns: grid.column_ranges(),
            column_range: 0,
            column_next: 0,
            rows: grid.row_indices().into_iter(),
            cells: None,
            grid,
            want,
            phase,
        }
    }

    fn next<'f>(&mut self, file: &'f Xls) -> Result<Option<Node<'f>>> {
        loop {
            match self.phase {
                SheetPhase::Columns => {
                    let Some((first, last)) = self.columns.get(self.column_range).copied() else {
                        self.phase = if self.want.rows {
                            SheetPhase::Rows
                        } else if self.want.cells {
                            SheetPhase::Cells
                        } else {
                            SheetPhase::Done
                        };
                        continue;
                    };
                    let index = self.column_next.max(first);
                    if index <= last {
                        if index == last {
                            self.column_range += 1;
                            self.column_next = 0;
                        } else {
                            self.column_next = index + 1;
                        }
                        return Ok(Some(Node::new_column(
                            file,
                            self.sheet,
                            self.grid.clone(),
                            index,
                        )));
                    }
                    self.column_range += 1;
                    self.column_next = 0;
                },
                SheetPhase::Rows => {
                    let Some(index) = self.rows.next() else {
                        self.phase = if self.want.cells {
                            SheetPhase::Cells
                        } else {
                            SheetPhase::Done
                        };
                        continue;
                    };
                    return Ok(Some(Node::new_row(
                        file,
                        self.sheet,
                        self.grid.clone(),
                        index,
                    )));
                },
                SheetPhase::Cells => {
                    if self.cells.is_none() {
                        self.cells = Some(
                            self.grid
                                .extents()
                                .map_or_else(Vec::new, |rect| {
                                    sparse_coords(&self.grid, rect).collect()
                                })
                                .into_iter(),
                        );
                    }
                    let Some(coords) = self.cells.as_mut() else {
                        self.phase = SheetPhase::Done;
                        continue;
                    };
                    let Some(at) = coords.next() else {
                        self.phase = SheetPhase::Done;
                        continue;
                    };
                    return Ok(Some(Node::new_cell(
                        file,
                        self.sheet,
                        at,
                        CellSource::Grid(self.grid.clone()),
                    )));
                },
                SheetPhase::Done => return Ok(None),
            }
        }
    }
}

impl<'f> Children<'f> {
    fn new(file: &'f Xls, state: State<'f>) -> Self {
        Self {
            file,
            state,
            fenced: true,
        }
    }

    pub(super) fn view(node: &Node<'f>, options: &ViewOptions) -> Result<Self> {
        let file = node.file();
        let state = match node.item() {
            Item::Workbook => State::Sheets {
                next: 0,
                count: super::nav::tab_count(file)?,
                names: false,
            },
            Item::Sheet(sheet) => {
                if !sheet.handle.is_grid() {
                    return Ok(Self::new(file, State::Done));
                }
                let grid = sheet.grid()?.clone();
                State::Rows {
                    sheet: node.sheet_position(),
                    rows: grid.row_indices().into_iter(),
                    grid,
                }
            },
            Item::Row(row) => State::Coordinates {
                sheet: node.sheet_position(),
                coords: row_coords(&row.grid, row.index),
                grid: row.grid.clone(),
            },
            Item::Column(column) => State::Coordinates {
                sheet: node.sheet_position(),
                coords: column_coords(&column.grid, column.index),
                grid: column.grid.clone(),
            },
            Item::Range(range) => {
                let cells = rect_size(range.rect);
                if let Some(budget) = options.dense {
                    if cells > u64::from(budget) {
                        return Err(dense_too_large(cells, budget));
                    }
                    State::Coordinates {
                        sheet: node.sheet_position(),
                        coords: dense_coords(range.rect),
                        grid: range.grid.clone(),
                    }
                } else {
                    State::Coordinates {
                        sheet: node.sheet_position(),
                        coords: sparse_coords(&range.grid, range.rect),
                        grid: range.grid.clone(),
                    }
                }
            },
            Item::Cell(_) | Item::NamedRange(_) | Item::Feature(_) => State::Done,
        };
        Ok(Self::new(file, state))
    }

    /// The query walker's children of `node`; the walker fences each advance,
    /// so these steps add no source probe of their own.
    pub(super) fn walk(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let mut children = Self::walk_children(node, kinds)?;
        children.fenced = false;
        Ok(children)
    }

    /// The query walker's children of `node` for a walk that is likely to
    /// stop early: a positional worksheet whose cells alone are wanted lists
    /// them through the leaf's cell cursor ([`Stream`]); everything else as
    /// [`Self::walk`].
    pub(super) fn walk_early(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        // A worksheet the cursor's validating scan refuses is read through
        // the grid, which owns the refusal.
        if let Item::Sheet(sheet) = node.item()
            && sheet.handle.is_grid()
            && kinds.contains(Kind::Cell)
            && !kinds.contains(Kind::Row)
            && !kinds.contains(Kind::Column)
            && let Ok(Some((cursor, worksheet))) = sheet.handle.cell_cursor()
        {
            let mut children = Self::new(
                node.file(),
                State::Stream(Box::new(Stream {
                    sheet: node.sheet_position(),
                    handle: sheet.handle.clone(),
                    cursor,
                    worksheet,
                    last: None,
                    rest: None,
                })),
            );
            children.fenced = false;
            return Ok(children);
        }
        Self::walk(node, kinds)
    }

    fn walk_children(node: &Node<'f>, kinds: KindSet<Kind>) -> Result<Self> {
        let file = node.file();
        let state = match node.item() {
            Item::Workbook if kinds.contains(Kind::Sheet) => State::Sheets {
                next: 0,
                count: super::nav::tab_count(file)?,
                names: kinds.contains(Kind::NamedRange),
            },
            Item::Workbook if kinds.contains(Kind::NamedRange) => State::Names {
                next: 0,
                count: super::node::defined_names(file)?.len(),
            },
            Item::Sheet(sheet)
                if super::features::KINDS
                    .iter()
                    .any(|kind| kinds.contains(*kind))
                    && !(kinds.contains(Kind::Cell)
                        || kinds.contains(Kind::Column)
                        || kinds.contains(Kind::Row)) =>
            {
                // A walk scoped to worksheet features lists them (record 0327).
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
            Item::Sheet(sheet) => {
                let want = SheetWant {
                    columns: kinds.contains(Kind::Column),
                    rows: kinds.contains(Kind::Row),
                    cells: kinds.contains(Kind::Cell),
                };
                if !sheet.handle.is_grid() || !want.any() {
                    return Ok(Self::new(file, State::Done));
                }
                if want.cells && !want.rows && !want.columns {
                    if let Some(cells) = sheet.handle.eager_cells(file)? {
                        return Ok(Self::new(
                            file,
                            State::EagerCells {
                                sheet: node.sheet_position(),
                                cells,
                            },
                        ));
                    }
                }
                let grid = sheet.grid()?.clone();
                State::SheetItems(SheetItems::new(node.sheet_position(), grid, want))
            },
            Item::Row(row) if kinds.contains(Kind::Cell) => State::Coordinates {
                sheet: node.sheet_position(),
                coords: row_coords(&row.grid, row.index),
                grid: row.grid.clone(),
            },
            Item::Column(column) if kinds.contains(Kind::Cell) => State::Coordinates {
                sheet: node.sheet_position(),
                coords: column_coords(&column.grid, column.index),
                grid: column.grid.clone(),
            },
            Item::Range(range) if kinds.contains(Kind::Cell) => State::Coordinates {
                sheet: node.sheet_position(),
                coords: sparse_coords(&range.grid, range.rect),
                grid: range.grid.clone(),
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
        match &mut self.state {
            State::EagerCells { sheet, cells } => {
                let Some(cell) = cells.next() else {
                    return Ok(None);
                };
                let at = (cell.row(), cell.column());
                let data = CellData::from_eager(cell);
                Ok(Some(Node::new_cell(
                    self.file,
                    *sheet,
                    at,
                    CellSource::Direct(Some(data)),
                )))
            },
            State::Stream(stream) => stream.next(self.file),
            State::Done => Ok(None),
            State::Sheets { next, count, names } => {
                if *next >= *count {
                    self.state = if *names {
                        State::Names {
                            next: 0,
                            count: super::node::defined_names(self.file)?.len(),
                        }
                    } else {
                        State::Done
                    };
                    return self.step();
                }
                let position = *next;
                *next += 1;
                let Some(handle) = SheetHandle::at(self.file, position)? else {
                    return Ok(None);
                };
                Ok(Some(Node::new_sheet(
                    self.file,
                    Sheet::new(position, handle),
                )))
            },
            State::Names { next, count } => {
                if *next >= *count {
                    self.state = State::Done;
                    return Ok(None);
                }
                let index = *next;
                *next += 1;
                Ok(Some(Node::new_named_range(self.file, index)))
            },
            State::Rows { sheet, grid, rows } => {
                let Some(index) = rows.next() else {
                    self.state = State::Done;
                    return Ok(None);
                };
                Ok(Some(Node::new_row(self.file, *sheet, grid.clone(), index)))
            },
            State::Coordinates {
                sheet,
                grid,
                coords,
            } => {
                let Some(at) = coords.next() else {
                    self.state = State::Done;
                    return Ok(None);
                };
                let source = if grid.record(at).is_some() {
                    CellSource::Grid(grid.clone())
                } else {
                    CellSource::Stub
                };
                Ok(Some(Node::new_cell(self.file, *sheet, at, source)))
            },
            State::SheetItems(items) => items.next(self.file),
            State::Features { sheet, items } => Ok(items
                .next()
                .map(|feature| Node::new_feature(self.file, *sheet, feature))),
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        if matches!(self.state, State::Done) {
            return None;
        }
        let step = if self.fenced {
            self.file.finish(|| self.step())
        } else {
            self.step()
        };
        match step {
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

fn sparse_coords(grid: &Grid, rect: (u32, u32, u32, u32)) -> std::vec::IntoIter<(u32, u32)> {
    grid.cells_in(rect)
        .map(|(at, _)| at)
        .collect::<Vec<_>>()
        .into_iter()
}

fn dense_coords(rect: (u32, u32, u32, u32)) -> std::vec::IntoIter<(u32, u32)> {
    let (first_row, first_column, last_row, last_column) = rect;
    let mut coords = Vec::new();
    for row in first_row..=last_row {
        for column in first_column..=last_column {
            coords.push((row, column));
        }
    }
    coords.into_iter()
}

fn row_coords(grid: &Grid, row: u32) -> std::vec::IntoIter<(u32, u32)> {
    grid.extents()
        .map(|(_, first_column, _, last_column)| {
            sparse_coords(grid, (row, first_column, row, last_column))
        })
        .unwrap_or_else(|| Vec::new().into_iter())
}

fn column_coords(grid: &Grid, column: u32) -> std::vec::IntoIter<(u32, u32)> {
    grid.extents()
        .map(|(first_row, _, last_row, _)| {
            sparse_coords(grid, (first_row, column, last_row, column))
        })
        .unwrap_or_else(|| Vec::new().into_iter())
}

fn rect_size((first_row, first_column, last_row, last_column): (u32, u32, u32, u32)) -> u64 {
    u64::from(last_row.saturating_sub(first_row).saturating_add(1)).saturating_mul(u64::from(
        last_column.saturating_sub(first_column).saturating_add(1),
    ))
}

fn dense_too_large(cells: u64, budget: u32) -> Error {
    Error::new(
        Code::OutputTooLarge,
        format!("the range has {cells} cells, more than the dense view budget of {budget}"),
    )
    .with_format(litchi_core::FileFormat::Xls)
    .with_suggestion("raise the `dense` budget, use a smaller range, or list stored cells only")
}

const _: () = {
    fn assert_send_sync<T: Send + Sync>() {}
    let _ = assert_send_sync::<Children<'static>>;
};
