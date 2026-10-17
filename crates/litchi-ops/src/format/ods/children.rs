//! Lazy children of ODS workbook nodes (record 0277).
//!
//! A worksheet's walk lists its content columns, rows and cells, then its
//! annotations and merge spans, each in document order; the node view lists
//! a sheet's content rows and a row's or range's content cells. Every list
//! reads the snapshot's cached content index, so a walk decodes nothing.

use litchi_ods::Sheet;

use super::Ods;
use super::index::{Entry, RectCells, SheetIndex};
use super::node::{Item, Node, Rect, annotations, sheet_index};
use super::spec::Kind;
use crate::error::Result;
use crate::node::ViewOptions;
use crate::path::navigate::KindSet;

/// The children of one node, in document order.
pub(crate) struct Children<'f> {
    file: &'f Ods,
    state: State<'f>,
}

enum State<'f> {
    Done,
    Sheets {
        next: usize,
        sheets: &'f [Sheet],
        /// Continue with the named ranges once the sheets are exhausted.
        names: bool,
    },
    Names {
        next: usize,
        count: usize,
    },
    Sheet(Box<SheetItems<'f>>),
    Cells {
        sheet: u32,
        rows: &'f Sheet,
        cells: CellIter<'f>,
    },
    Rows {
        sheet: u32,
        rows: std::slice::Iter<'f, (u32, u32)>,
    },
}

/// The content cells of a row or a range.
enum CellIter<'f> {
    Slice(std::slice::Iter<'f, Entry>),
    Rect(RectCells<'f>),
}

impl<'f> CellIter<'f> {
    fn next(&mut self) -> Option<&'f Entry> {
        match self {
            Self::Slice(iter) => iter.next(),
            Self::Rect(iter) => iter.next(),
        }
    }
}

#[derive(Clone, Copy)]
enum Phase {
    Columns,
    Rows,
    Cells,
    Comments,
    Merges,
    Features,
    Done,
}

/// The walk over one worksheet's kinds.
struct SheetItems<'f> {
    sheet: u32,
    data: &'f Sheet,
    index: &'f SheetIndex,
    kinds: KindSet<Kind>,
    phase: Phase,
    next: usize,
    /// The per-sheet ordinal of the next annotation.
    ordinal: usize,
    /// The worksheet features still to list (record 0327).
    features: Option<std::vec::IntoIter<super::features::Feature>>,
}

impl<'f> SheetItems<'f> {
    fn advance(&mut self) {
        self.next = 0;
        self.phase = match self.phase {
            Phase::Columns => Phase::Rows,
            Phase::Rows => Phase::Cells,
            Phase::Cells => Phase::Comments,
            Phase::Comments => Phase::Merges,
            Phase::Merges => Phase::Features,
            Phase::Features | Phase::Done => Phase::Done,
        };
    }

    fn next(&mut self, file: &'f Ods) -> Result<Option<Node<'f>>> {
        loop {
            match self.phase {
                Phase::Columns => {
                    if self.kinds.contains(Kind::Column)
                        && let Some(column) = self.index.columns().get(self.next)
                    {
                        self.next += 1;
                        return Ok(Some(Node::new_column(file, self.sheet, *column)));
                    }
                    self.advance();
                },
                Phase::Rows => {
                    if self.kinds.contains(Kind::Row)
                        && let Some((row, _)) = self.index.row_entries().get(self.next)
                    {
                        self.next += 1;
                        return Ok(Some(Node::new_row(file, self.sheet, *row)));
                    }
                    self.advance();
                },
                Phase::Cells => {
                    if self.kinds.contains(Kind::Cell)
                        && let Some(entry) = self.index.cells().get(self.next)
                    {
                        self.next += 1;
                        return Ok(Some(Node::new_cell(
                            file,
                            self.sheet,
                            (entry.row, entry.column),
                            entry.cell(self.data),
                        )));
                    }
                    self.advance();
                },
                Phase::Comments => {
                    if self.kinds.contains(Kind::Comment) {
                        let entries = annotations(file)?.entries();
                        while let Some(entry) = entries.get(self.next) {
                            let index = self.next;
                            self.next += 1;
                            if entry.cell().sheet() == self.data.name {
                                let ordinal = self.ordinal;
                                self.ordinal += 1;
                                return Ok(Some(Node::new_comment(
                                    file, self.sheet, index, ordinal,
                                )));
                            }
                        }
                    }
                    self.advance();
                },
                Phase::Merges => {
                    if self.kinds.contains(Kind::MergeCell)
                        && let Some(span) = self.index.merges().get(self.next)
                    {
                        let ordinal = self.next;
                        self.next += 1;
                        return Ok(Some(Node::new_merge(file, self.sheet, *span, ordinal)));
                    }
                    self.advance();
                },
                Phase::Features => {
                    if self.features.is_none() {
                        let mut items = Vec::new();
                        for kind in super::features::KINDS {
                            if self.kinds.contains(kind) {
                                items.extend(super::features::list(file, self.sheet, kind)?);
                            }
                        }
                        self.features = Some(items.into_iter());
                    }
                    if let Some(feature) = self.features.as_mut().and_then(Iterator::next) {
                        return Ok(Some(Node::new_feature(file, self.sheet, feature)));
                    }
                    self.advance();
                },
                Phase::Done => return Ok(None),
            }
        }
    }
}

impl<'f> Children<'f> {
    const fn done(file: &'f Ods) -> Self {
        Self {
            file,
            state: State::Done,
        }
    }

    /// The query walker's children of `parent` with kinds in `kinds`.
    pub(crate) fn walk(parent: &Node<'f>, kinds: KindSet<Kind>, file: &'f Ods) -> Result<Self> {
        Ok(match parent.item() {
            Item::Workbook => {
                let sheets = kinds.contains(Kind::Sheet);
                let names = kinds.contains(Kind::NamedRange);
                if sheets {
                    Self {
                        file,
                        state: State::Sheets {
                            next: 0,
                            sheets: file.worksheets()?,
                            names,
                        },
                    }
                } else if names {
                    Self::names(file)?
                } else {
                    Self::done(file)
                }
            },
            Item::Sheet(sheet) => {
                let wanted = [
                    Kind::Column,
                    Kind::Row,
                    Kind::Cell,
                    Kind::Comment,
                    Kind::MergeCell,
                ]
                .into_iter()
                .chain(super::features::KINDS)
                .any(|kind| kinds.contains(kind));
                if !wanted {
                    return Ok(Self::done(file));
                }
                let position = parent.sheet_position();
                if [Kind::Column, Kind::Row, Kind::Comment, Kind::MergeCell]
                    .into_iter()
                    .chain(super::features::KINDS)
                    .all(|kind| !kinds.contains(kind))
                {
                    // Only cells are wanted (record 0287): the sheet walk's
                    // other phases would list nothing, so list the content
                    // cells directly, in the same order.
                    return Ok(Self {
                        file,
                        state: State::Cells {
                            sheet: position,
                            rows: sheet,
                            cells: CellIter::Slice(sheet_index(file, position)?.cells().iter()),
                        },
                    });
                }
                Self {
                    file,
                    state: State::Sheet(Box::new(SheetItems {
                        sheet: position,
                        data: sheet,
                        index: sheet_index(file, position)?,
                        kinds,
                        phase: Phase::Columns,
                        next: 0,
                        ordinal: 0,
                        features: None,
                    })),
                }
            },
            Item::Range(range) if kinds.contains(Kind::Cell) => {
                Self::cells_in(parent, range.rect, file)?
            },
            _ => Self::done(file),
        })
    }

    fn names(file: &'f Ods) -> Result<Self> {
        Ok(Self {
            file,
            state: State::Names {
                next: 0,
                count: file.definitions()?.len(),
            },
        })
    }

    fn cells_in(parent: &Node<'f>, rect: Rect, file: &'f Ods) -> Result<Self> {
        let index = parent.index()?;
        Ok(Self {
            file,
            state: State::Cells {
                sheet: parent.sheet_position(),
                rows: parent.sheet()?,
                cells: CellIter::Rect(index.cells_in(rect)),
            },
        })
    }

    /// The node view's children of `parent`.
    pub(crate) fn view(parent: &Node<'f>, _options: &ViewOptions) -> Result<Self> {
        let file = parent.file();
        Ok(match parent.item() {
            Item::Workbook => Self {
                file,
                state: State::Sheets {
                    next: 0,
                    sheets: file.worksheets()?,
                    names: false,
                },
            },
            Item::Sheet(_) => Self {
                file,
                state: State::Rows {
                    sheet: parent.sheet_position(),
                    rows: parent.index()?.row_entries().iter(),
                },
            },
            Item::Row(row) => Self {
                file,
                state: State::Cells {
                    sheet: parent.sheet_position(),
                    rows: parent.sheet()?,
                    cells: CellIter::Slice(parent.index()?.row_cells(row.index).iter()),
                },
            },
            Item::Range(range) => Self::cells_in(parent, range.rect, file)?,
            _ => Self::done(file),
        })
    }

    /// The number of children [`Self::view`] lists.
    pub(crate) fn count(parent: &Node<'f>, options: &ViewOptions) -> Result<usize> {
        Ok(match parent.item() {
            Item::Sheet(_) => parent.index()?.row_count(),
            Item::Row(row) => parent.index()?.row_cells(row.index).len(),
            _ => {
                let mut children = Self::view(parent, options)?;
                let mut count = 0_usize;
                while children.next_node()?.is_some() {
                    count += 1;
                }
                count
            },
        })
    }

    fn next_node(&mut self) -> Result<Option<Node<'f>>> {
        let file = self.file;
        match &mut self.state {
            State::Done => Ok(None),
            State::Sheets {
                next,
                sheets,
                names,
            } => {
                if let Some(sheet) = sheets.get(*next) {
                    let position = u32::try_from(*next).unwrap_or(u32::MAX);
                    *next += 1;
                    return Ok(Some(Node::new_sheet(file, position, sheet)));
                }
                if *names {
                    *self = Self::names(file)?;
                    return self.next_node();
                }
                self.state = State::Done;
                Ok(None)
            },
            State::Names { next, count } => {
                if *next < *count {
                    let index = *next;
                    *next += 1;
                    return Ok(Some(Node::new_named_range(file, index)));
                }
                self.state = State::Done;
                Ok(None)
            },
            State::Sheet(items) => {
                let node = items.next(file)?;
                if node.is_none() {
                    self.state = State::Done;
                }
                Ok(node)
            },
            State::Cells { sheet, rows, cells } => Ok(cells.next().map(|entry| {
                Node::new_cell(file, *sheet, (entry.row, entry.column), entry.cell(rows))
            })),
            State::Rows { sheet, rows } => Ok(rows
                .next()
                .map(|(row, _)| Node::new_row(file, *sheet, *row))),
        }
    }
}

impl<'f> Iterator for Children<'f> {
    type Item = Result<Node<'f>>;

    fn next(&mut self) -> Option<Self::Item> {
        match self.next_node() {
            Ok(Some(node)) => Some(Ok(node)),
            Ok(None) => None,
            Err(error) => {
                self.state = State::Done;
                Some(Err(error))
            },
        }
    }
}
