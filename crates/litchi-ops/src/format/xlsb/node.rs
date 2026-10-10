//! XLSB nodes and their unified workbook view (record 0010 R4d).
//!
//! The node layer keeps the XLSB leaf's positional catalog and immutable grid
//! snapshots.  A cell node carries only a cheap grid handle and a checked
//! coordinate; missing cells are represented by a small state variant, so a
//! dense range does not allocate a boxed leaf value for every coordinate.

use std::borrow::Cow;
use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::selector::CellRef;
use litchi_core::sheet::{Cell as CoreCell, WorkbookTrait};
use litchi_core::{FileFormat, Position};
use litchi_xlsb as xlsb;
use litchi_xlsb::sheet::Rect;

use super::children::Children;
pub(crate) use super::owner::SheetHandle;
use super::spec::{
    CELL_PROPS, COLUMN_PROPS, CellKey, ColumnKey, Kind, NAMED_RANGE_PROPS, NamedRangeKey,
    ROW_PROPS, RowKey, SHEET_PROPS, SheetKey, TABLE_PROPS, TableKey, WORKBOOK_PROPS, WorkbookKey,
};
use super::{Xlsb, XlsbOwner};
use crate::error::{Code, Error, Result};
pub(super) use crate::format::workbook::Label;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::value::Length;

/// The text of an in-grid coordinate with no stored record.
pub(super) const MISSING_TEXT: &str = "(empty)";

/// A checked zero-based `(row, column)` coordinate supplied by navigation.
type Address = (u32, u32);

/// A node of an XLSB workbook.
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Xlsb,
    item: Item,
    steps: [PathStep<'f>; 2],
    depth: u8,
    sheet: usize,
}

/// The semantic element held by a node.  Native relationship identifiers and
/// physical part names never enter this enum or its paths.
#[derive(Clone)]
pub(super) enum Item {
    Workbook,
    Sheet(Sheet),
    Cell(CellItem),
    Range(RangeItem),
    Row(RowItem),
    Column(ColumnItem),
    Table(TableItem),
    NamedRange(usize),
}

/// One complete-catalog tab and its lazily pinned worksheet grid.
#[derive(Clone)]
pub(super) struct Sheet {
    pub(super) position: usize,
    pub(super) handle: SheetHandle,
    grid: OnceLock<Grid>,
}

impl Sheet {
    pub(super) fn new(position: Position, handle: SheetHandle) -> Self {
        Self {
            position: position.get(),
            handle,
            grid: OnceLock::new(),
        }
    }

    /// Materialize a worksheet grid only when a child or cell read needs it.
    /// Public node reads and navigation bracket this retained snapshot with
    /// owner freshness and execution checks.
    pub(super) fn grid(&self) -> Result<&Grid> {
        if let Some(grid) = self.grid.get() {
            return Ok(grid);
        }
        let grid = Grid::from_handle(&self.handle)?;
        Ok(self.grid.get_or_init(|| grid))
    }
}

/// A coordinate and the grid or dense-range state from which it is read.
#[derive(Clone)]
pub(super) struct CellItem {
    pub(super) at: Address,
    pub(super) source: CellSource,
    label: Label,
    format: OnceLock<Option<StyleFormat>>,
}

/// A cell in a pinned sparse grid, or a dense-range coordinate with no record.
#[derive(Clone)]
pub(super) enum CellSource {
    Grid(Grid),
    Stub,
}

#[derive(Clone)]
pub(super) struct RangeItem {
    pub(super) grid: Grid,
    pub(super) rect: Rect,
    label: Label,
}

#[derive(Clone)]
pub(super) struct RowItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    pub(super) bulk: bool,
    label: Label,
}

#[derive(Clone)]
pub(super) struct ColumnItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    label: Label,
}

#[derive(Clone)]
pub(super) struct TableItem {
    pub(super) part: Box<xlsb::package::table::Table>,
}

/// A pinned sparse worksheet.  The grid owns an `Arc` in the leaf and the
/// handle supplies source/version fences and local style resolution.
#[derive(Clone)]
pub(super) struct Grid {
    leaf: xlsb::sheet::Grid,
    handle: SheetHandle,
    cached_only: bool,
}

/// Sparse cells selected by a checked rectangle. The leaf iterator borrows
/// cells only; this adapter adds coordinates without copying or boxing them.
pub(super) struct Cells<'a> {
    inner: xlsb::sheet::Cells<'a>,
}

impl<'a> Iterator for Cells<'a> {
    type Item = (Address, &'a xlsb::Cell);

    fn next(&mut self) -> Option<Self::Item> {
        self.inner
            .next()
            .map(|cell| ((cell.row(), cell.column()), cell))
    }
}

impl std::iter::FusedIterator for Cells<'_> {}

impl Grid {
    pub(super) fn from_handle(handle: &SheetHandle) -> Result<Self> {
        let read = handle.grid_read()?;
        Ok(Self {
            leaf: read.leaf,
            handle: handle.clone(),
            cached_only: read.cached_only,
        })
    }

    pub(super) fn cached_only(&self) -> bool {
        self.cached_only
    }

    pub(super) fn record(&self, at: Address) -> Result<Option<&xlsb::Cell>> {
        Ok(self.leaf.cell_at(at)?)
    }

    pub(super) fn cells(&self, rect: Rect) -> Result<Cells<'_>> {
        Ok(Cells {
            inner: self.leaf.cells_in(rect)?,
        })
    }

    pub(super) fn extents(&self) -> Result<Option<Rect>> {
        let Some((first_row, first_column, last_row, last_column)) = self.leaf.dimensions() else {
            return Ok(None);
        };
        let end_row = last_row
            .checked_add(1)
            .ok_or_else(|| Error::new(Code::OutOfRange, "XLSB worksheet row extent overflow"))?;
        let end_column = last_column
            .checked_add(1)
            .ok_or_else(|| Error::new(Code::OutOfRange, "XLSB worksheet column extent overflow"))?;
        Rect::at(first_row, first_column, end_row, end_column)
            .map(Some)
            .map_err(|_| Error::new(Code::InvalidPath, "invalid XLSB worksheet extent"))
    }

    pub(super) fn merge(&self, at: Address) -> Result<Option<(Rect, bool)>> {
        if self.cached_only {
            return Err(unsupported(
                "XLSB cached-value worksheet has no merged-cell metadata",
            ));
        }
        let (row, column) = at;
        for merge in self.leaf.worksheet().merged_cells() {
            if (merge.row_first..=merge.row_last).contains(&row)
                && (merge.col_first..=merge.col_last).contains(&column)
            {
                let rect = Rect::at(
                    merge.row_first,
                    merge.col_first,
                    merge.row_last.saturating_add(1),
                    merge.col_last.saturating_add(1),
                )
                .map_err(|_| Error::new(Code::InvalidPath, "invalid XLSB merged range"))?;
                return Ok(Some((
                    rect,
                    row == merge.row_first && column == merge.col_first,
                )));
            }
        }
        Ok(None)
    }

    fn format(&self, at: Address) -> Result<Option<StyleFormat>> {
        if self.cached_only {
            return Err(unsupported(
                "XLSB cached-value worksheet has no cell-format metadata",
            ));
        }
        let Some(cell) = self.record(at)? else {
            return Ok(None);
        };
        Ok(self
            .handle
            .style_for_cell(cell)?
            .map(|xf| StyleFormat { xf }))
    }

    pub(super) fn row_indices(&self) -> Result<Vec<u32>> {
        if self.cached_only {
            let mut rows = self
                .cells(Rect::ALL)?
                .map(|(at, _)| at.0)
                .collect::<Vec<_>>();
            rows.sort_unstable();
            rows.dedup();
            return Ok(rows);
        }
        let worksheet = self.leaf.worksheet();
        let mut rows = worksheet
            .row_infos()
            .iter()
            .map(|row| row.row)
            .collect::<Vec<_>>();
        for (_, cell) in self.cells(Rect::ALL)? {
            if rows.last().copied() != Some(cell.row()) {
                rows.push(cell.row());
            }
        }
        rows.sort_unstable();
        rows.dedup();
        Ok(rows)
    }

    pub(super) fn column_ranges(&self) -> Result<Vec<(u32, u32)>> {
        if self.cached_only {
            let mut columns = self
                .cells(Rect::ALL)?
                .map(|(at, _)| at.1)
                .collect::<Vec<_>>();
            columns.sort_unstable();
            columns.dedup();
            return columns
                .into_iter()
                .try_fold(Vec::<(u32, u32)>::new(), |mut ranges, column| {
                    if let Some((_, previous)) = ranges.last_mut() {
                        if column <= previous.saturating_add(1) {
                            *previous = (*previous).max(column);
                            return Ok(ranges);
                        }
                    }
                    ranges.push((column, column));
                    Ok(ranges)
                });
        }
        let mut ranges = self
            .leaf
            .worksheet()
            .column_infos()
            .iter()
            .map(|info| (info.first_column, info.last_column))
            .collect::<Vec<_>>();
        ranges.sort_unstable();
        ranges.dedup();
        ranges
            .into_iter()
            .try_fold(Vec::<(u32, u32)>::new(), |mut merged, (first, last)| {
                if let Some((_, previous_last)) = merged.last_mut() {
                    if first <= previous_last.saturating_add(1) {
                        *previous_last = (*previous_last).max(last);
                        return Ok(merged);
                    }
                }
                merged.push((first, last));
                Ok(merged)
            })
    }

    pub(super) fn row_has_meaningful_attrs(&self, index: u32) -> Result<bool> {
        if self.cached_only {
            let _ = index;
            return Ok(false);
        }
        Ok(self.row_info(index).is_some_and(|row| {
            row.height.is_some()
                || row.hidden
                || row.outline_level > 0
                || row.collapsed
                || row.show_phonetic
                || !row.column_spans.is_empty()
        }))
    }

    pub(super) fn row_info(&self, index: u32) -> Option<xlsb::sheet::RowInfo> {
        self.leaf
            .worksheet()
            .row_infos()
            .iter()
            .find(|info| info.row == index)
            .cloned()
    }

    pub(super) fn column_info(&self, index: u32) -> Option<xlsb::sheet::ColumnInfo> {
        self.leaf
            .worksheet()
            .column_infos()
            .iter()
            .find(|info| info.first_column <= index && index <= info.last_column)
            .cloned()
    }
}

const NO_STEP: PathStep<'static> = PathStep::new("");

impl<'f> Node<'f> {
    pub(super) fn new_workbook(file: &'f Xlsb) -> Self {
        Self {
            file,
            item: Item::Workbook,
            steps: [NO_STEP; 2],
            depth: 0,
            sheet: 0,
        }
    }

    pub(super) fn new_sheet(file: &'f Xlsb, sheet: Sheet) -> Self {
        let position = sheet.position;
        Self {
            file,
            item: Item::Sheet(sheet),
            steps: [PathStep::at("sheet", Position::new(position)), NO_STEP],
            depth: 1,
            sheet: position,
        }
    }

    fn in_sheet(file: &'f Xlsb, sheet: usize, item: Item, tail: PathStep<'f>) -> Self {
        Self {
            file,
            item,
            steps: [PathStep::at("sheet", Position::new(sheet)), tail],
            depth: 2,
            sheet,
        }
    }

    pub(super) fn new_cell(file: &'f Xlsb, sheet: usize, at: Address, source: CellSource) -> Self {
        let (row, column) = at;
        let item = Item::Cell(CellItem {
            at,
            source,
            label: Label::cell(column, row),
            format: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item, PathStep::cell(column, row))
    }

    pub(super) fn new_range(file: &'f Xlsb, sheet: usize, grid: Grid, rect: Rect) -> Self {
        let first = (rect.start().column().get(), rect.start().row().get());
        let last = (
            rect.end().1.saturating_sub(1),
            rect.end().0.saturating_sub(1),
        );
        let item = Item::Range(RangeItem {
            grid,
            rect,
            label: Label::range(first, last),
        });
        Self::in_sheet(file, sheet, item, PathStep::range(first, last))
    }

    pub(super) fn new_row(file: &'f Xlsb, sheet: usize, grid: Grid, index: u32) -> Self {
        Self::row(file, sheet, grid, index, false)
    }

    pub(super) fn new_listed_row(file: &'f Xlsb, sheet: usize, grid: Grid, index: u32) -> Self {
        Self::row(file, sheet, grid, index, true)
    }

    fn row(file: &'f Xlsb, sheet: usize, grid: Grid, index: u32, bulk: bool) -> Self {
        let item = Item::Row(RowItem {
            grid,
            index,
            bulk,
            label: Label::row(index),
        });
        Self::in_sheet(
            file,
            sheet,
            item,
            PathStep::at("row", Position::new(index as usize)),
        )
    }

    pub(super) fn new_column(file: &'f Xlsb, sheet: usize, grid: Grid, index: u32) -> Self {
        let item = Item::Column(ColumnItem {
            grid,
            index,
            label: Label::column(index),
        });
        Self::in_sheet(file, sheet, item, PathStep::column("col", index))
    }

    pub(super) fn new_table(
        file: &'f Xlsb,
        sheet: usize,
        index: usize,
        table: xlsb::package::table::Table,
    ) -> Self {
        Self::in_sheet(
            file,
            sheet,
            Item::Table(TableItem {
                part: Box::new(table),
            }),
            PathStep::at("table", Position::new(index)),
        )
    }

    pub(super) fn new_named_range(file: &'f Xlsb, index: usize) -> Self {
        Self {
            file,
            item: Item::NamedRange(index),
            steps: [PathStep::at("namedrange", Position::new(index)), NO_STEP],
            depth: 1,
            sheet: 0,
        }
    }

    pub(super) fn file(&self) -> &'f Xlsb {
        self.file
    }

    pub(super) fn item(&self) -> &Item {
        &self.item
    }

    pub(super) fn sheet_position(&self) -> usize {
        self.sheet
    }

    /// Zero-based worksheet position used by the native workbook editor.
    /// Complete workbook catalog positions can also refer to chart, dialog,
    /// macro, or other non-worksheet tabs, which deliberately return `None`.
    pub(super) fn worksheet_position(&self) -> Option<usize> {
        match &self.item {
            Item::Sheet(sheet) => sheet.handle.worksheet_position().map(Position::get),
            _ => None,
        }
    }

    pub(super) fn last_step(&self) -> Option<&PathStep<'f>> {
        usize::from(self.depth)
            .checked_sub(1)
            .and_then(|index| self.steps.get(index))
    }

    pub(super) fn element(&self) -> Kind {
        match self.item {
            Item::Workbook => Kind::Workbook,
            Item::Sheet(_) => Kind::Sheet,
            Item::Cell(_) => Kind::Cell,
            Item::Range(_) => Kind::Range,
            Item::Row(_) => Kind::Row,
            Item::Column(_) => Kind::Column,
            Item::Table(_) => Kind::Table,
            Item::NamedRange(_) => Kind::NamedRange,
        }
    }

    #[must_use]
    pub fn kind(&self) -> &'static str {
        super::spec::SPEC.elements[self.element() as usize].name
    }

    /// The semantic name of the node's tab, or `None` for workbook/name roots.
    /// Names are owned because the source catalog's complete tab names are
    /// deferred and may not outlive this individual fenced read.
    pub fn sheet_name(&self) -> Result<Option<String>> {
        self.file.0.owner.finish(|| match self.item {
            Item::Workbook | Item::NamedRange(_) => Ok(None),
            Item::Sheet(ref sheet) => Ok(Some(sheet.handle.name().to_owned())),
            _ => SheetHandle::at(self.file, self.sheet)
                .map(|handle| handle.map(|handle| handle.name().to_owned())),
        })
    }

    #[must_use]
    pub fn cell(&self) -> Option<(u32, u32)> {
        match &self.item {
            Item::Cell(cell) => Some((cell.at.1, cell.at.0)),
            _ => None,
        }
    }

    #[must_use]
    pub fn number(&self) -> Option<u32> {
        match &self.item {
            Item::Row(row) => row.index.checked_add(1),
            Item::Column(column) => column.index.checked_add(1),
            _ => None,
        }
    }

    pub(super) fn cell_ref(&self) -> Option<CellRef> {
        match &self.item {
            Item::Cell(cell) => CellRef::parse(cell.label.as_str()),
            _ => None,
        }
    }

    pub(super) fn cell_state(&self) -> Result<Option<CellState<'_>>> {
        self.file.0.owner.finish(|| {
            let Item::Cell(cell) = &self.item else {
                return Ok(None);
            };
            cell.state().map(Some)
        })
    }

    pub(super) fn cell_text(&self) -> Result<Option<Cow<'_, str>>> {
        Ok(self.cell_state()?.map(|state| match state {
            CellState::Stored(cell) => display_text(cell),
            CellState::Missing | CellState::Covered(_) => Cow::Borrowed(MISSING_TEXT),
            CellState::Stub => Cow::Borrowed(""),
        }))
    }

    pub(super) fn props(&self) -> Result<Format<'_>> {
        self.file.0.owner.finish(|| self.props_unfenced())
    }

    fn props_unfenced(&self) -> Result<Format<'_>> {
        let (table, props) = match &self.item {
            Item::Workbook => (
                WORKBOOK_PROPS,
                Props::Workbook {
                    file: self.file,
                    date_1904: date_1904(self.file),
                    active_tab: active_tab(self.file)?,
                },
            ),
            Item::Sheet(sheet) => (
                SHEET_PROPS,
                Props::Sheet {
                    name: sheet.handle.name().to_owned(),
                    state: sheet.handle.state(),
                },
            ),
            Item::Cell(cell) => (CELL_PROPS, Props::Cell(cell.props()?)),
            Item::Range(_) => return Ok(Format::empty()),
            Item::Row(row) => {
                if row.grid.cached_only() {
                    return Ok(Format::empty());
                }
                let info = row.grid.row_info(row.index);
                (ROW_PROPS, Props::Row(info))
            },
            Item::Column(column) => {
                if column.grid.cached_only() {
                    return Ok(Format::empty());
                }
                let info = column.grid.column_info(column.index);
                (COLUMN_PROPS, Props::Column(info))
            },
            Item::Table(table) => (TABLE_PROPS, Props::Table(&table.part)),
            Item::NamedRange(index) => {
                let Some(name) = defined_names(self.file)?.get(*index).cloned() else {
                    return Ok(Format::empty());
                };
                (NAMED_RANGE_PROPS, Props::NamedRange { name })
            },
        };
        Ok(Format::new(table, Source::Xlsb(props)))
    }
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("xlsb::Node")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

/// The cell state read by a node.
#[derive(Clone, Copy)]
pub(super) enum CellState<'a> {
    Stored(&'a xlsb::Cell),
    Missing,
    Covered(Rect),
    Stub,
}

impl CellItem {
    pub(super) fn state(&self) -> Result<CellState<'_>> {
        match &self.source {
            CellSource::Stub => Ok(CellState::Stub),
            CellSource::Grid(grid) => {
                let at = self.at;
                if let Some(cell) = grid.record(at)? {
                    return Ok(CellState::Stored(cell));
                }
                if grid.cached_only() {
                    return Ok(CellState::Missing);
                }
                Ok(grid
                    .merge(at)?
                    .map_or(CellState::Missing, |(range, _)| CellState::Covered(range)))
            },
        }
    }

    fn format(&self) -> Result<Option<&StyleFormat>> {
        if let Some(format) = self.format.get() {
            return Ok(format.as_ref());
        }
        let format = match (&self.source, self.state()?) {
            (CellSource::Grid(grid), CellState::Stored(_)) => grid.format(self.at)?,
            _ => None,
        };
        Ok(self.format.get_or_init(|| format).as_ref())
    }

    fn merge(&self, state: CellState<'_>) -> Result<Option<(Rect, bool)>> {
        match (&self.source, state) {
            (_, CellState::Covered(range)) => Ok(Some((range, false))),
            (CellSource::Grid(grid), CellState::Stored(_)) => grid.merge(self.at),
            _ => Ok(None),
        }
    }

    fn props(&self) -> Result<CellProps<'_>> {
        let state = self.state()?;
        let cached_only = matches!(&self.source, CellSource::Grid(grid) if grid.cached_only());
        let format = match state {
            CellState::Stored(_) if !cached_only => self.format()?,
            _ => None,
        };
        let merge = (!cached_only)
            .then(|| self.merge(state))
            .transpose()?
            .flatten();
        Ok(CellProps {
            state,
            format,
            merge,
            cached_only,
        })
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        Node::kind(self)
    }

    fn file_format(&self) -> FileFormat {
        FileFormat::Xlsb
    }

    fn path(&self) -> NodePath<'_> {
        NodePath::new(&self.steps[..usize::from(self.depth)])
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        self.file.0.owner.finish(|| {
            Ok(match &self.item {
                Item::Cell(_) => self.cell_text()?,
                Item::Row(row) => {
                    if row.grid.cached_only() {
                        return Ok(None);
                    }
                    row.grid
                        .row_info(row.index)
                        .is_none()
                        .then_some(Cow::Borrowed(MISSING_TEXT))
                },
                Item::Table(table) => Some(Cow::Borrowed(table_label(table))),
                Item::NamedRange(index) => defined_names(self.file)?
                    .get(*index)
                    .cloned()
                    .map(Cow::Owned),
                Item::Workbook | Item::Sheet(_) | Item::Range(_) | Item::Column(_) => None,
            })
        })
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.0.owner.finish(|| {
            Ok(match &self.item {
                Item::Workbook => None,
                Item::Sheet(sheet) => Some(Preview::Text(Cow::Borrowed(sheet.handle.name()))),
                Item::Cell(cell) => Some(Preview::Text(Cow::Borrowed(cell.label.as_str()))),
                Item::Range(range) => Some(Preview::Text(Cow::Borrowed(range.label.as_str()))),
                Item::Row(row) => Some(Preview::Text(Cow::Borrowed(row.label.as_str()))),
                Item::Column(column) => Some(Preview::Text(Cow::Borrowed(column.label.as_str()))),
                Item::Table(table) => Some(Preview::Text(Cow::Borrowed(table_label(table)))),
                Item::NamedRange(index) => defined_names(self.file)?
                    .get(*index)
                    .cloned()
                    .map(Cow::Owned)
                    .map(Preview::Text),
            })
        })
    }

    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>> {
        self.file.0.owner.finish(|| {
            let count = match &self.item {
                Item::Workbook => catalog_count(self.file)?,
                Item::Sheet(_) if options.depth == 0 => return Ok(None),
                Item::Cell(_) | Item::Table(_) | Item::NamedRange(_) => 0,
                Item::Sheet(_) | Item::Row(_) | Item::Column(_) | Item::Range(_) => {
                    Children::count(self, options)?
                },
            };
            Ok(Some(u32::try_from(count).unwrap_or(u32::MAX)))
        })
    }

    fn format(&self) -> Result<Format<'_>> {
        self.props()
    }

    fn children(&self, options: &ViewOptions) -> ViewChildren<'_, 'f> {
        if let Err(error) = self.file.0.owner.ensure_current() {
            return ViewChildren::failed(error);
        }
        match Children::view(self, options) {
            Ok(children) => ViewChildren::xlsb(children),
            Err(error) => ViewChildren::failed(error),
        }
    }
}

fn unsupported(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Xlsb)
}

fn catalog_count(file: &Xlsb) -> Result<usize> {
    match &file.0.owner {
        XlsbOwner::Source(workbook) => workbook.workbook().sheet_count().map_err(Into::into),
        XlsbOwner::Eager(workbook) => Ok(workbook.sheets().len()),
    }
}

pub(super) fn defined_names(file: &Xlsb) -> Result<Arc<[String]>> {
    file.defined_names()
}

pub(super) fn table_label(table: &TableItem) -> &str {
    table
        .part
        .display_name
        .as_deref()
        .or(table.part.name.as_deref())
        .unwrap_or_default()
}

/// A normalized local style snapshot.  The source route currently exposes
/// only the XF, while the eager route also resolves component records.
#[derive(Clone)]
struct StyleFormat {
    xf: xlsb::package::styles_table::CellFormat,
}

/// A node's property source.  Debug intentionally reports only the variant;
/// authored names, formulas and cell values stay out of diagnostics.
pub(crate) enum Props<'a> {
    Workbook {
        file: &'a Xlsb,
        date_1904: bool,
        active_tab: Option<usize>,
    },
    Sheet {
        name: String,
        state: xlsb::sheet::State,
    },
    Cell(CellProps<'a>),
    Row(Option<xlsb::sheet::RowInfo>),
    Column(Option<xlsb::sheet::ColumnInfo>),
    Table(&'a xlsb::package::table::Table),
    NamedRange {
        name: String,
    },
}

pub(crate) struct CellProps<'a> {
    state: CellState<'a>,
    format: Option<&'a StyleFormat>,
    merge: Option<(Rect, bool)>,
    cached_only: bool,
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Workbook { .. } => "Workbook",
            Self::Sheet { .. } => "Sheet",
            Self::Cell(_) => "Cell",
            Self::Row(_) => "Row",
            Self::Column(_) => "Column",
            Self::Table(_) => "Table",
            Self::NamedRange { .. } => "NamedRange",
        };
        f.debug_tuple("Props").field(&name).finish()
    }
}

impl<'a> Props<'a> {
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match self {
            Self::Workbook {
                file,
                date_1904,
                active_tab,
            } => workbook_value(file, *date_1904, *active_tab, WorkbookKey::at(index)?),
            Self::Sheet { name, state } => match SheetKey::at(index)? {
                SheetKey::Name => Some(Value::Text(Cow::Owned(name.clone()))),
                SheetKey::Hidden => state.is_hidden().then_some(Value::Bool(true)),
                SheetKey::Visibility => match state {
                    xlsb::sheet::State::Visible => None,
                    xlsb::sheet::State::Hidden => Some(Value::Token("hidden")),
                    xlsb::sheet::State::VeryHidden => Some(Value::Token("veryHidden")),
                },
            },
            Self::Cell(cell) => cell.get(CellKey::at(index)?),
            Self::Row(row) => row
                .as_ref()
                .and_then(|row| row_value(row, RowKey::at(index)?)),
            Self::Column(column) => column
                .as_ref()
                .and_then(|column| column_value(column, ColumnKey::at(index)?)),
            Self::Table(table) => table_value(table, TableKey::at(index)?),
            Self::NamedRange { name } => match NamedRangeKey::at(index)? {
                NamedRangeKey::Name => Some(Value::Text(Cow::Owned(name.clone()))),
            },
        }
    }
}

macro_rules! timestamp {
    ($utc:expr, $local:expr) => {
        match ($utc, $local) {
            (Some(utc), _) => Some(Value::Text(Cow::Owned(format!(
                "{}.{:07}Z",
                utc.format("%Y-%m-%dT%H:%M:%S"),
                utc.timestamp_subsec_nanos() / 100
            )))),
            (None, Some(local)) => Some(Value::Text(Cow::Owned(format!(
                "{}.{:07}",
                local.format("%Y-%m-%dT%H:%M:%S"),
                local.and_utc().timestamp_subsec_nanos() / 100
            )))),
            (None, None) => None,
        }
    };
}

fn workbook_value(
    file: &Xlsb,
    date_1904: bool,
    active_tab: Option<usize>,
    key: WorkbookKey,
) -> Option<Value<'_>> {
    let metadata = &file.0.metadata;
    match key {
        WorkbookKey::Title => metadata.title.as_ref().map(|value| text(value)),
        WorkbookKey::Author => metadata.author.as_ref().map(|value| text(value)),
        WorkbookKey::Subject => metadata.subject.as_ref().map(|value| text(value)),
        WorkbookKey::Keywords => metadata.keywords.as_ref().map(|value| text(value)),
        WorkbookKey::Description => metadata.description.as_ref().map(|value| text(value)),
        WorkbookKey::Category => metadata.category.as_ref().map(|value| text(value)),
        WorkbookKey::LastModifiedBy => metadata.last_modified_by.as_ref().map(|value| text(value)),
        WorkbookKey::RevisionNumber => metadata.revision.as_ref().map(|value| text(value)),
        WorkbookKey::Created => {
            timestamp!(metadata.created.as_ref(), metadata.created_local.as_ref())
        },
        WorkbookKey::Modified => {
            timestamp!(metadata.modified.as_ref(), metadata.modified_local.as_ref())
        },
        WorkbookKey::Date1904 => date_1904.then_some(Value::Bool(true)),
        WorkbookKey::ActiveTab => active_tab
            .filter(|position| *position != 0)
            .map(|position| Value::Int(i64::try_from(position).unwrap_or(i64::MAX))),
    }
}

fn date_1904(file: &Xlsb) -> bool {
    match &file.0.owner {
        XlsbOwner::Source(workbook) => workbook.date1904(),
        XlsbOwner::Eager(workbook) => workbook.is_1904_date_system(),
    }
}

fn active_tab(file: &Xlsb) -> Result<Option<usize>> {
    file.active_catalog_position()
        .map(|position| position.map(Position::get))
}

fn text(value: &str) -> Value<'_> {
    Value::Text(Cow::Borrowed(value))
}

fn row_value(row: &xlsb::sheet::RowInfo, key: RowKey) -> Option<Value<'static>> {
    match key {
        RowKey::Height => row.height.map(points),
        RowKey::Hidden => row.hidden.then_some(Value::Bool(true)),
        RowKey::OutlineLevel => {
            (row.outline_level > 0).then_some(Value::Int(i64::from(row.outline_level)))
        },
        RowKey::Collapsed => row.collapsed.then_some(Value::Bool(true)),
        RowKey::CustomHeight => row.height.is_some().then_some(Value::Bool(true)),
    }
}

fn column_value(column: &xlsb::sheet::ColumnInfo, key: ColumnKey) -> Option<Value<'static>> {
    match key {
        ColumnKey::Width => Value::number(column.width),
        ColumnKey::Hidden => column.hidden.then_some(Value::Bool(true)),
        ColumnKey::CustomWidth => column.user_set_width.then_some(Value::Bool(true)),
        ColumnKey::OutlineLevel => {
            (column.outline_level > 0).then_some(Value::Int(i64::from(column.outline_level)))
        },
        ColumnKey::Collapsed => column.collapsed.then_some(Value::Bool(true)),
    }
}

fn table_value(table: &xlsb::package::table::Table, key: TableKey) -> Option<Value<'_>> {
    let style = table.style_info.as_ref();
    let flag = |value: bool| value.then_some(Value::Bool(true));
    match key {
        TableKey::Name => table.name.as_deref().map(text),
        TableKey::DisplayName => table.display_name.as_deref().map(text),
        TableKey::Ref => Some(Value::Text(Cow::Owned(table_ref(table)))),
        TableKey::Style => style.and_then(|style| style.name.as_deref()).map(text),
        TableKey::BandedRows => style.and_then(|style| flag(style.show_row_stripes)),
        TableKey::BandedCols => style.and_then(|style| flag(style.show_column_stripes)),
        TableKey::FirstCol => style.and_then(|style| flag(style.show_first_column)),
        TableKey::LastCol => style.and_then(|style| flag(style.show_last_column)),
        TableKey::HeaderRow => Some(Value::Bool(table.header_row_count != 0)),
        TableKey::TotalRow => Some(Value::Bool(
            table.totals_row_count != 0 || table.totals_row_shown,
        )),
        TableKey::Columns => Some(Value::Text(Cow::Owned(
            table
                .columns
                .iter()
                .filter_map(|column| column.name.as_deref().or(column.caption.as_deref()))
                .collect::<Vec<_>>()
                .join(","),
        ))),
    }
}

fn table_ref(table: &xlsb::package::table::Table) -> String {
    Label::range(
        (table.range.first_column, table.range.first_row),
        (table.range.last_column, table.range.last_row),
    )
    .as_str()
    .to_owned()
}

impl<'a> CellProps<'a> {
    fn get(&self, key: CellKey) -> Option<Value<'a>> {
        if let CellKey::Merge = key {
            return self
                .merge
                .map(|(range, _)| Value::Text(Cow::Owned(range_text(range))));
        }
        if let CellKey::MergeAnchor = key {
            return self
                .merge
                .and_then(|(_, anchor)| anchor.then_some(Value::Bool(true)));
        }
        let cell = match self.state {
            CellState::Stored(cell) => cell,
            CellState::Stub => {
                return match key {
                    CellKey::Type => Some(Value::Token("Number")),
                    CellKey::Empty => Some(Value::Bool(true)),
                    _ => None,
                };
            },
            CellState::Missing | CellState::Covered(_) => return None,
        };
        let value = cell.value();
        let formula = match value {
            litchi_core::sheet::CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => Some((
                formula.as_str(),
                cached_value.as_deref(),
                *is_array,
                array_range.as_deref(),
            )),
            _ => None,
        };
        match key {
            CellKey::Type => Some(Value::Token(cell_type(value))),
            CellKey::Formula => (!self.cached_only)
                .then(|| formula.map(|(formula, _, _, _)| text(formula)))
                .flatten(),
            CellKey::CachedValue => formula
                .and_then(|(_, cached, _, _)| cached)
                .filter(|_| !self.cached_only)
                .map(stored_value),
            CellKey::Evaluated => (!self.cached_only)
                .then_some(
                    cell.is_formula()
                        .then_some(Value::Bool(cell.cached_value().is_some())),
                )
                .flatten(),
            CellKey::ArrayFormula => (!self.cached_only)
                .then(|| {
                    formula.and_then(|(_, _, is_array, _)| is_array.then_some(Value::Bool(true)))
                })
                .flatten(),
            CellKey::ArrayRef => (!self.cached_only)
                .then(|| formula.and_then(|(_, _, _, array_range)| array_range.map(text)))
                .flatten(),
            CellKey::Empty => {
                matches!(value, litchi_core::sheet::CellValue::Empty).then_some(Value::Bool(true))
            },
            key => format_value(self.format?, key),
        }
    }
}

pub(super) fn display_text(cell: &xlsb::Cell) -> Cow<'_, str> {
    display_value(cell.value())
}

fn display_value(value: &litchi_core::sheet::CellValue) -> Cow<'_, str> {
    match value {
        litchi_core::sheet::CellValue::Empty => Cow::Borrowed(""),
        litchi_core::sheet::CellValue::Bool(value) => {
            Cow::Borrowed(if *value { "TRUE" } else { "FALSE" })
        },
        litchi_core::sheet::CellValue::Int(value) => Cow::Owned(value.to_string()),
        litchi_core::sheet::CellValue::Float(value)
        | litchi_core::sheet::CellValue::DateTime(value) => Cow::Owned(value.to_string()),
        litchi_core::sheet::CellValue::String(value)
        | litchi_core::sheet::CellValue::Error(value) => Cow::Borrowed(value),
        litchi_core::sheet::CellValue::Formula { cached_value, .. } => cached_value
            .as_deref()
            .map_or(Cow::Borrowed(""), display_value),
    }
}

fn stored_value(value: &litchi_core::sheet::CellValue) -> Value<'static> {
    match value {
        litchi_core::sheet::CellValue::Bool(value) => Value::Text(Cow::Owned(if *value {
            "1".to_owned()
        } else {
            "0".to_owned()
        })),
        _ => Value::Text(Cow::Owned(display_value(value).into_owned())),
    }
}

pub(super) fn cell_type(value: &litchi_core::sheet::CellValue) -> &'static str {
    match value {
        litchi_core::sheet::CellValue::Bool(_) => "Boolean",
        litchi_core::sheet::CellValue::String(_) => "SharedString",
        litchi_core::sheet::CellValue::DateTime(_) => "Date",
        litchi_core::sheet::CellValue::Error(_) => "Error",
        litchi_core::sheet::CellValue::Formula { cached_value, .. } => {
            cached_value.as_deref().map_or("Number", cell_type)
        },
        litchi_core::sheet::CellValue::Empty
        | litchi_core::sheet::CellValue::Int(_)
        | litchi_core::sheet::CellValue::Float(_) => "Number",
    }
}

pub(super) fn is_error(cell: &xlsb::Cell) -> bool {
    fn error(value: &litchi_core::sheet::CellValue) -> bool {
        match value {
            litchi_core::sheet::CellValue::Error(_) => true,
            litchi_core::sheet::CellValue::Formula { cached_value, .. } => {
                cached_value.as_deref().is_some_and(error)
            },
            _ => false,
        }
    }
    error(cell.value())
}

pub(super) fn raw_value(cell: &xlsb::Cell) -> Value<'static> {
    fn raw(value: &litchi_core::sheet::CellValue) -> Option<Value<'static>> {
        match value {
            litchi_core::sheet::CellValue::Int(value) => Some(Value::Int(*value)),
            litchi_core::sheet::CellValue::Float(value)
            | litchi_core::sheet::CellValue::DateTime(value) => Value::number(*value),
            litchi_core::sheet::CellValue::Formula { cached_value, .. } => {
                cached_value.as_deref().and_then(raw)
            },
            _ => None,
        }
    }
    raw(cell.value()).unwrap_or_else(|| Value::Text(Cow::Owned(display_text(cell).into_owned())))
}

fn format_value(format: &StyleFormat, key: CellKey) -> Option<Value<'_>> {
    let alignment = format.xf.alignment.as_ref();
    match key {
        CellKey::WrapText => alignment
            .filter(|alignment| alignment.wrap_text)
            .map(|_| Value::Bool(true)),
        CellKey::Horizontal => alignment.map(|alignment| horizontal(alignment.horizontal)),
        CellKey::Vertical => alignment.map(|alignment| vertical(alignment.vertical)),
        CellKey::TextRotation => alignment
            .map(|alignment| i64::from(alignment.rotation))
            .filter(|rotation| *rotation != 0)
            .map(Value::Int),
        CellKey::Indent => alignment
            .map(|alignment| i64::from(alignment.indent))
            .filter(|indent| *indent != 0)
            .map(Value::Int),
        CellKey::ShrinkToFit => alignment
            .filter(|alignment| alignment.shrink_to_fit)
            .map(|_| Value::Bool(true)),
        CellKey::ReadingOrder => alignment.and_then(|alignment| match alignment.text_direction {
            1 => Some(Value::Token("ltr")),
            2 => Some(Value::Token("rtl")),
            _ => None,
        }),
        CellKey::Type
        | CellKey::Formula
        | CellKey::CachedValue
        | CellKey::Evaluated
        | CellKey::ArrayFormula
        | CellKey::ArrayRef
        | CellKey::Empty
        | CellKey::Merge
        | CellKey::MergeAnchor => None,
    }
}

fn horizontal(value: xlsb::package::styles::HorizontalAlignment) -> Value<'static> {
    Value::Token(match value {
        xlsb::package::styles::HorizontalAlignment::General => "general",
        xlsb::package::styles::HorizontalAlignment::Left => "left",
        xlsb::package::styles::HorizontalAlignment::Center => "center",
        xlsb::package::styles::HorizontalAlignment::Right => "right",
        xlsb::package::styles::HorizontalAlignment::Fill => "fill",
        xlsb::package::styles::HorizontalAlignment::Justify => "justify",
        xlsb::package::styles::HorizontalAlignment::CenterContinuous => "centerContinuous",
        xlsb::package::styles::HorizontalAlignment::Distributed => "distributed",
    })
}

fn vertical(value: xlsb::package::styles::VerticalAlignment) -> Value<'static> {
    Value::Token(match value {
        xlsb::package::styles::VerticalAlignment::Top => "top",
        xlsb::package::styles::VerticalAlignment::Center => "center",
        xlsb::package::styles::VerticalAlignment::Bottom => "bottom",
        xlsb::package::styles::VerticalAlignment::Justify => "justify",
        xlsb::package::styles::VerticalAlignment::Distributed => "distributed",
    })
}

fn points(value: f64) -> Value<'static> {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "validated font/row points fit in EMU"
    )]
    let emu = (value * 12_700.0).round() as i64;
    Value::Length(Length::from_emu(emu))
}

fn range_text(range: Rect) -> String {
    let first = (range.start().column().get(), range.start().row().get());
    let last = (
        range.end().1.saturating_sub(1),
        range.end().0.saturating_sub(1),
    );
    Label::range(first, last).as_str().to_owned()
}

const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Node<'static>>();
};
