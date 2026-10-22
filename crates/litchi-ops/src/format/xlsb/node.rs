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
    RANGE_MERGE_INDEX, RANGE_PROPS, ROW_PROPS, RowKey, SHEET_PROPS, SheetKey, TABLE_PROPS,
    TableKey, WORKBOOK_PROPS, WorkbookKey,
};
use super::{Xlsb, XlsbOwner};
use crate::error::{Code, Error, Result};
pub(super) use crate::format::workbook::Label;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::value::{Color, Length};

/// The text of an in-grid coordinate with no stored record.
pub(super) const MISSING_TEXT: &str = "(empty)";

/// A checked zero-based `(row, column)` coordinate supplied by navigation.
type Address = (u32, u32);

/// A node of an XLSB workbook.
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Xlsb,
    item: Item,
    /// The position of the node's worksheet (0 for the workbook and a
    /// defined name). The path is written from it and the item on demand
    /// (record 0267, as record 0262 did for XLSX).
    sheet: u32,
}

// A walk moves every node it visits: keep it small (record 0267; 176 bytes
// before, with its two path steps and label kept inline).
const _: () = assert!(size_of::<Node<'static>>() <= 96);

/// The semantic element held by a node.  Native relationship identifiers and
/// physical part names never enter this enum or its paths.
#[derive(Clone)]
pub(super) enum Item {
    Workbook,
    /// Boxed: a tab handle is large and sheets are few (record 0267).
    Sheet(Box<Sheet>),
    Cell(CellItem),
    Range(RangeItem),
    Row(RowItem),
    Column(ColumnItem),
    Table(TableItem),
    NamedRange(usize),
    /// A worksheet feature: a note, hyperlink, picture, shape, conditional
    /// format, validation, auto-filter or page break (record 0322).
    Feature(Arc<super::features::Feature>),
    /// A worksheet chart, its series and its axes, read-only (record 0322).
    Chart(Box<super::chart::ChartItem>),
    Series(Box<super::chart::ChartItem>),
    Axis(Box<super::chart::ChartItem>),
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

/// A direct mutation target's worksheet, resolved on first inspection.
#[derive(Clone)]
pub(super) struct DeferredCell {
    handle: SheetHandle,
    grid: OnceLock<Grid>,
}

/// A coordinate and the grid or dense-range state from which it is read.
#[derive(Clone)]
pub(super) struct CellItem {
    pub(super) at: Address,
    pub(super) source: CellSource,
    label: OnceLock<Label>,
    format: OnceLock<Option<StyleFormat>>,
}

/// A cell in a pinned sparse grid, or a dense-range coordinate with no record.
#[derive(Clone)]
pub(super) enum CellSource {
    Grid(Grid),
    /// A direct mutation target.  The worksheet grid remains lazy so target
    /// admission can validate the path without parsing the read projection;
    /// a later inspection still resolves the same source-backed grid.
    Deferred(Box<DeferredCell>),
    Stub,
}

#[derive(Clone)]
pub(super) struct RangeItem {
    pub(super) grid: Grid,
    pub(super) rect: Rect,
    label: OnceLock<Label>,
}

#[derive(Clone)]
pub(super) struct RowItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    pub(super) bulk: bool,
    label: OnceLock<Label>,
}

#[derive(Clone)]
pub(super) struct ColumnItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    label: OnceLock<Label>,
}

#[derive(Clone)]
pub(super) struct TableItem {
    pub(super) part: Box<xlsb::package::table::Table>,
    /// The table's position among its worksheet's tables, for its path step.
    index: usize,
}

/// A pinned sparse worksheet.  The grid owns an `Arc` in the leaf and the
/// handle supplies source/version fences and local style resolution.
///
/// One `Arc`, so listing a cell costs one reference-count increment instead
/// of a tab-handle clone (record 0267).
#[derive(Clone)]
pub(super) struct Grid(Arc<GridInner>);

pub(super) struct GridInner {
    leaf: xlsb::sheet::Grid,
    handle: SheetHandle,
    cached_only: bool,
    /// For a batch of the leaf's cell cursor (record 0267), `leaf` holds the
    /// batch's cells only and this the worksheet's complete grid, read on
    /// first use by a read that needs more than a cell's own record (its
    /// merge).
    complete: Option<OnceLock<xlsb::sheet::Grid>>,
}

impl std::ops::Deref for Grid {
    type Target = GridInner;

    fn deref(&self) -> &GridInner {
        &self.0
    }
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
        Ok(Self(Arc::new(GridInner {
            leaf: read.leaf,
            handle: handle.clone(),
            cached_only: read.cached_only,
            complete: None,
        })))
    }

    /// The node grid of a grid the leaf's cell cursor read: the complete
    /// grid, or a batch of it.
    pub(super) fn streamed(leaf: xlsb::sheet::Grid, handle: SheetHandle, complete: bool) -> Self {
        Self(Arc::new(GridInner {
            leaf,
            handle,
            cached_only: false,
            complete: (!complete).then(OnceLock::new),
        }))
    }

    /// The worksheet's complete grid.
    fn full(&self) -> Result<&xlsb::sheet::Grid> {
        let Some(complete) = &self.complete else {
            return Ok(&self.leaf);
        };
        if let Some(grid) = complete.get() {
            return Ok(grid);
        }
        let read = self.handle.grid_read()?;
        Ok(complete.get_or_init(|| read.leaf))
    }

    pub(super) fn cached_only(&self) -> bool {
        self.cached_only
    }

    pub(super) fn record(&self, at: Address) -> Result<Option<&xlsb::Cell>> {
        if let Some(cell) = self.leaf.cell_at(at)? {
            return Ok(Some(cell));
        }
        if self.complete.is_none() {
            return Ok(None);
        }
        Ok(self.full()?.cell_at(at)?)
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

    /// Whether `rect` is exactly one merged range.
    pub(super) fn is_merged(&self, rect: Rect) -> bool {
        let first = rect.start();
        let (end_row, end_column) = rect.end();
        self.leaf.worksheet().merged_cells().iter().any(|merge| {
            merge.row_first == first.row().get()
                && merge.col_first == first.column().get()
                && merge.row_last.saturating_add(1) == end_row
                && merge.col_last.saturating_add(1) == end_column
        })
    }

    pub(super) fn merge(&self, at: Address) -> Result<Option<(Rect, bool)>> {
        if self.cached_only {
            return Err(unsupported(
                "XLSB cached-value worksheet has no merged-cell metadata",
            ));
        }
        let (row, column) = at;
        for merge in self.full()?.worksheet().merged_cells() {
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
        Ok(self.handle.style_for_cell(cell)?.map(|style| StyleFormat {
            style: Box::new(style),
        }))
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

/// The first and last cells of `rect`, each a zero-based `(column, row)`.
fn corners(rect: Rect) -> ((u32, u32), (u32, u32)) {
    let first = (rect.start().column().get(), rect.start().row().get());
    let last = (
        rect.end().1.saturating_sub(1),
        rect.end().0.saturating_sub(1),
    );
    (first, last)
}

/// A position as stored in a node: tab positions are below the catalog's
/// length, which no workbook brings near `u32::MAX`.
fn position_u32(position: usize) -> u32 {
    u32::try_from(position).unwrap_or(u32::MAX)
}

impl<'f> Node<'f> {
    pub(super) fn new_workbook(file: &'f Xlsb) -> Self {
        Self {
            file,
            item: Item::Workbook,
            sheet: 0,
        }
    }

    pub(super) fn new_sheet(file: &'f Xlsb, sheet: Sheet) -> Self {
        let position = position_u32(sheet.position);
        Self {
            file,
            item: Item::Sheet(Box::new(sheet)),
            sheet: position,
        }
    }

    /// An element of the worksheet at `sheet`, for a sibling module.
    pub(super) fn in_sheet_item(file: &'f Xlsb, sheet: usize, item: Item) -> Self {
        Self::in_sheet(file, sheet, item)
    }

    fn in_sheet(file: &'f Xlsb, sheet: usize, item: Item) -> Self {
        Self {
            file,
            item,
            sheet: position_u32(sheet),
        }
    }

    pub(super) fn new_cell(file: &'f Xlsb, sheet: usize, at: Address, source: CellSource) -> Self {
        let item = Item::Cell(CellItem {
            at,
            source,
            label: OnceLock::new(),
            format: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
    }

    pub(super) fn new_cell_target(
        file: &'f Xlsb,
        sheet: usize,
        at: Address,
        handle: SheetHandle,
    ) -> Self {
        Self::new_cell(
            file,
            sheet,
            at,
            CellSource::Deferred(Box::new(DeferredCell {
                handle,
                grid: OnceLock::new(),
            })),
        )
    }

    pub(super) fn new_range(file: &'f Xlsb, sheet: usize, grid: Grid, rect: Rect) -> Self {
        let item = Item::Range(RangeItem {
            grid,
            rect,
            label: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
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
            label: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
    }

    pub(super) fn new_column(file: &'f Xlsb, sheet: usize, grid: Grid, index: u32) -> Self {
        let item = Item::Column(ColumnItem {
            grid,
            index,
            label: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
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
                index,
            }),
        )
    }

    pub(super) fn new_named_range(file: &'f Xlsb, index: usize) -> Self {
        Self {
            file,
            item: Item::NamedRange(index),
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
        self.sheet as usize
    }

    /// The node's own path step, written from its item; `None` for the
    /// workbook root.
    pub(super) fn last_step(&self) -> Option<PathStep<'static>> {
        Some(match &self.item {
            Item::Workbook => return None,
            Item::Sheet(_) => PathStep::at("sheet", Position::new(self.sheet_position())),
            Item::Cell(cell) => PathStep::cell(cell.at.1, cell.at.0),
            Item::Range(range) => {
                let (first, last) = corners(range.rect);
                PathStep::range(first, last)
            },
            Item::Row(row) => PathStep::at("row", Position::new(row.index as usize)),
            Item::Column(column) => PathStep::column("col", column.index),
            Item::Table(table) => PathStep::at("table", Position::new(table.index)),
            Item::NamedRange(index) => PathStep::at("namedrange", Position::new(*index)),
            Item::Feature(feature) => feature.step(),
            Item::Chart(chart) => chart.step(Kind::Chart),
            Item::Series(chart) => chart.step(Kind::Series),
            Item::Axis(chart) => chart.step(Kind::Axis),
        })
    }

    /// The A1 label of a cell, range, row or column node, written on first
    /// use.
    fn label(&self) -> Option<&Label> {
        Some(match &self.item {
            Item::Cell(cell) => cell.label.get_or_init(|| Label::cell(cell.at.1, cell.at.0)),
            Item::Range(range) => range.label.get_or_init(|| {
                let (first, last) = corners(range.rect);
                Label::range(first, last)
            }),
            Item::Row(row) => row.label.get_or_init(|| Label::row(row.index)),
            Item::Column(column) => column.label.get_or_init(|| Label::column(column.index)),
            _ => return None,
        })
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
            Item::Feature(ref feature) => feature.kind,
            Item::Chart(_) => Kind::Chart,
            Item::Series(_) => Kind::Series,
            Item::Axis(_) => Kind::Axis,
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
            _ => SheetHandle::at(self.file, self.sheet_position())
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
            Item::Cell(_) => self
                .label()
                .and_then(|label| CellRef::parse(label.as_str())),
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
            Item::Range(range) => {
                if range.grid.cached_only() {
                    return Ok(Format::empty());
                }
                (
                    RANGE_PROPS,
                    Props::Range {
                        merged: range.grid.is_merged(range.rect),
                    },
                )
            },
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
            Item::Feature(feature) => {
                // Only a note reads its worksheet's name (`anchoredTo`).
                let sheet = if feature.kind == Kind::Comment {
                    self.file
                        .catalog()?
                        .get(self.sheet_position())
                        .map(|tab| tab.name().to_owned())
                        .unwrap_or_default()
                } else {
                    String::new()
                };
                (feature.table(), Props::Feature(feature, sheet))
            },
            Item::Chart(chart) => (
                super::chart::props(Kind::Chart),
                Props::Chart(chart, Kind::Chart),
            ),
            Item::Series(chart) => (
                super::chart::props(Kind::Series),
                Props::Chart(chart, Kind::Series),
            ),
            Item::Axis(chart) => (
                super::chart::props(Kind::Axis),
                Props::Chart(chart, Kind::Axis),
            ),
            Item::NamedRange(index) => {
                let Some(entry) = self.file.defined_name_entries()?.into_iter().nth(*index) else {
                    return Ok(Format::empty());
                };
                let scope = match entry.scope {
                    Some(position) => Some(
                        self.file
                            .catalog()?
                            .get(position)
                            .map(|tab| tab.name().to_owned())
                            .ok_or_else(|| {
                                Error::new(
                                    Code::CorruptFile,
                                    "an XLSB defined name names a missing scope sheet",
                                )
                            })?,
                    ),
                    None => None,
                };
                (NAMED_RANGE_PROPS, Props::NamedRange { entry, scope })
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
    fn grid(&self) -> Result<&Grid> {
        match &self.source {
            CellSource::Grid(grid) => Ok(grid),
            CellSource::Deferred(deferred) => {
                if let Some(grid) = deferred.grid.get() {
                    return Ok(grid);
                }
                let materialized = Grid::from_handle(&deferred.handle)?;
                Ok(deferred.grid.get_or_init(|| materialized))
            },
            CellSource::Stub => Err(unsupported("the XLSB cell has no worksheet grid")),
        }
    }

    pub(super) fn state(&self) -> Result<CellState<'_>> {
        match &self.source {
            CellSource::Stub => Ok(CellState::Stub),
            CellSource::Grid(_) | CellSource::Deferred(_) => {
                let grid = self.grid()?;
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
        let format = match self.state()? {
            CellState::Stored(_) => self.grid()?.format(self.at)?,
            _ => None,
        };
        Ok(self.format.get_or_init(|| format).as_ref())
    }

    fn merge(&self, state: CellState<'_>) -> Result<Option<(Rect, bool)>> {
        match (&self.source, state) {
            (_, CellState::Covered(range)) => Ok(Some((range, false))),
            (CellSource::Grid(_) | CellSource::Deferred(_), CellState::Stored(_)) => {
                self.grid()?.merge(self.at)
            },
            _ => Ok(None),
        }
    }

    fn props(&self) -> Result<CellProps<'_>> {
        let state = self.state()?;
        let cached_only = match &self.source {
            CellSource::Grid(grid) => grid.cached_only(),
            CellSource::Deferred(_) => self.grid()?.cached_only(),
            CellSource::Stub => false,
        };
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
        match (&self.item, self.last_step()) {
            (_, None) => NodePath::ROOT,
            (Item::Sheet(_) | Item::NamedRange(_), Some(step)) => NodePath::short(step, None),
            (Item::Series(chart) | Item::Axis(chart), Some(step)) => NodePath::inline(&[
                PathStep::at("sheet", Position::new(self.sheet_position())),
                PathStep::at("chart", Position::new(chart.chart as usize)),
                step,
            ])
            .unwrap_or(NodePath::ROOT),
            (_, Some(step)) => NodePath::short(
                PathStep::at("sheet", Position::new(self.sheet_position())),
                Some(step),
            ),
        }
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
                Item::Feature(feature) => feature.text(),
                Item::Chart(chart) => chart.title()?,
                Item::Workbook
                | Item::Sheet(_)
                | Item::Range(_)
                | Item::Column(_)
                | Item::Series(_)
                | Item::Axis(_) => None,
            })
        })
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.0.owner.finish(|| {
            Ok(match &self.item {
                Item::Workbook => None,
                Item::Sheet(sheet) => Some(Preview::Text(Cow::Borrowed(sheet.handle.name()))),
                Item::Cell(_) | Item::Range(_) | Item::Row(_) | Item::Column(_) => self
                    .label()
                    .map(|label| Preview::Text(Cow::Borrowed(label.as_str()))),
                Item::Table(table) => Some(Preview::Text(Cow::Borrowed(table_label(table)))),
                Item::NamedRange(index) => defined_names(self.file)?
                    .get(*index)
                    .cloned()
                    .map(Cow::Owned)
                    .map(Preview::Text),
                Item::Feature(feature) => feature.preview(),
                Item::Chart(chart) => chart.title()?.map(Preview::Text),
                Item::Series(_) | Item::Axis(_) => None,
            })
        })
    }

    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>> {
        self.file.0.owner.finish(|| {
            let count = match &self.item {
                Item::Workbook => catalog_count(self.file)?,
                Item::Sheet(_) if options.depth == 0 => return Ok(None),
                Item::Cell(_)
                | Item::Table(_)
                | Item::NamedRange(_)
                | Item::Feature(_)
                | Item::Series(_)
                | Item::Axis(_) => 0,
                Item::Chart(chart) => {
                    let summary = chart.summary()?;
                    summary.series.len() + summary.axes.len()
                },
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
        XlsbOwner::Eager(package) => Ok(super::projection(package)?.sheets().len()),
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

/// A normalized local style snapshot: the cell XF with the font, fill,
/// border and number-format records it names, resolved through the owner's
/// stylesheet on both owner routes.
#[derive(Clone)]
struct StyleFormat {
    style: Box<super::owner::ResolvedStyle>,
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
        entry: xlsb::named_ranges::Entry,
        scope: Option<String>,
    },
    Range {
        merged: bool,
    },
    /// A worksheet feature and its worksheet's name (record 0322).
    Feature(&'a super::features::Feature, String),
    /// A worksheet chart, series or axis (record 0322).
    Chart(&'a super::chart::ChartItem, Kind),
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
            Self::Range { .. } => "Range",
            Self::Feature(..) => "Feature",
            Self::Chart(..) => "Chart",
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
            Self::Range { merged } => {
                (index == RANGE_MERGE_INDEX && *merged).then_some(Value::Bool(true))
            },
            Self::Feature(feature, sheet) => feature.value(index, sheet),
            Self::Chart(chart, kind) => chart.value(*kind, index),
            Self::NamedRange { entry, scope } => match NamedRangeKey::at(index)? {
                NamedRangeKey::Name => Some(Value::Text(Cow::Owned(entry.name.clone()))),
                NamedRangeKey::Ref => entry
                    .formula
                    .as_ref()
                    .map(|formula| Value::Text(Cow::Owned(formula.clone()))),
                NamedRangeKey::Scope => Some(Value::Text(Cow::Owned(
                    scope.clone().unwrap_or_else(|| "workbook".to_owned()),
                ))),
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
        WorkbookKey::Keywords => Value::owned_list(&metadata.keywords),
        WorkbookKey::Description => metadata.description.as_ref().map(|value| text(value)),
        WorkbookKey::Category => metadata.category.as_ref().map(|value| text(value)),
        WorkbookKey::LastModifiedBy => metadata.last_modified_by.as_ref().map(|value| text(value)),
        WorkbookKey::RevisionNumber => metadata.revision.as_ref().map(|value| text(value)),
        WorkbookKey::Created => {
            timestamp!(
                metadata
                    .created
                    .and_then(litchi_core::metadata::Timestamp::utc)
                    .as_ref(),
                metadata
                    .created
                    .and_then(litchi_core::metadata::Timestamp::unzoned)
                    .as_ref()
            )
        },
        WorkbookKey::Modified => {
            timestamp!(
                metadata
                    .modified
                    .and_then(litchi_core::metadata::Timestamp::utc)
                    .as_ref(),
                metadata
                    .modified
                    .and_then(litchi_core::metadata::Timestamp::unzoned)
                    .as_ref()
            )
        },
        WorkbookKey::ExtendedTemplate => metadata.template.as_ref().map(|value| text(value)),
        WorkbookKey::ExtendedManager => metadata.manager.as_ref().map(|value| text(value)),
        WorkbookKey::ExtendedCompany => metadata.company.as_ref().map(|value| text(value)),
        WorkbookKey::ExtendedApplication => metadata.application.as_ref().map(|value| text(value)),
        WorkbookKey::ExtendedPages => metadata
            .page_count
            .map(|value| Value::Int(i64::from(value))),
        WorkbookKey::ExtendedWords => metadata
            .word_count
            .map(|value| Value::Int(i64::from(value))),
        WorkbookKey::ExtendedCharacters => metadata
            .character_count
            .map(|value| Value::Int(i64::from(value))),
        WorkbookKey::ExtendedTotalTime => metadata
            .editing_time_minutes
            .map(|value| Value::Int(i64::from(value))),
        WorkbookKey::Date1904 => date_1904.then_some(Value::Bool(true)),
        WorkbookKey::ActiveTab => active_tab
            .filter(|position| *position != 0)
            .map(|position| Value::Int(i64::try_from(position).unwrap_or(i64::MAX))),
    }
}

fn date_1904(file: &Xlsb) -> bool {
    match &file.0.owner {
        XlsbOwner::Source(workbook) => workbook.date1904(),
        XlsbOwner::Eager(package) => {
            super::projection(package).is_ok_and(WorkbookTrait::is_1904_date_system)
        },
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

/// A key read from a cell's local formatting, with OfficeCLI's omissions:
/// the font, fill and border keys only when the format names a record other
/// than the workbook's default one, flags only when set, and colors only when
/// they are direct RGB.
fn format_value(format: &StyleFormat, key: CellKey) -> Option<Value<'_>> {
    let style = &format.style;
    let alignment = style.xf.alignment.as_ref();
    let font = style.font.as_ref();
    let border = style.border.as_ref();
    match key {
        CellKey::Bold => font?.bold.then_some(Value::Bool(true)),
        CellKey::Italic => font?.italic.then_some(Value::Bool(true)),
        CellKey::Strike => font?.strike.then_some(Value::Bool(true)),
        CellKey::Underline => font?.underline.then_some(Value::Token("single")),
        CellKey::FontColor => font?.color.map(argb_color),
        CellKey::FontSize => font.map(|font| points(font.size)),
        CellKey::FontName => font.map(|font| Value::Text(Cow::Owned(font.name.clone()))),
        CellKey::Fill => {
            let fill = style.fill.as_ref()?;
            (fill.pattern_type == 1)
                .then_some(fill.fg_color)
                .flatten()
                .map(argb_color)
        },
        CellKey::BorderLeft => side_line(border?.left.as_ref()),
        CellKey::BorderLeftColor => side_color(border?.left.as_ref()),
        CellKey::BorderRight => side_line(border?.right.as_ref()),
        CellKey::BorderRightColor => side_color(border?.right.as_ref()),
        CellKey::BorderTop => side_line(border?.top.as_ref()),
        CellKey::BorderTopColor => side_color(border?.top.as_ref()),
        CellKey::BorderBottom => side_line(border?.bottom.as_ref()),
        CellKey::BorderBottomColor => side_color(border?.bottom.as_ref()),
        CellKey::NumberFormat => style
            .number_format
            .as_ref()
            .filter(|code| code.as_str() != "General")
            .map(|code| Value::Text(Cow::Owned(code.clone()))),
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

/// A direct ARGB color as OfficeCLI reads it back.
fn argb_color(argb: u32) -> Value<'static> {
    let [alpha, red, green, blue] = argb.to_be_bytes();
    Value::Color(Color::rgba(red, green, blue, alpha))
}

fn side_line(side: Option<&xlsb::package::styles::BorderSide>) -> Option<Value<'static>> {
    use xlsb::package::styles::BorderStyle;
    Some(Value::Token(match side?.style {
        BorderStyle::None => return None,
        BorderStyle::Thin => "thin",
        BorderStyle::Medium => "medium",
        BorderStyle::Dashed => "dashed",
        BorderStyle::Dotted => "dotted",
        BorderStyle::Thick => "thick",
        BorderStyle::Double => "double",
        BorderStyle::Hair => "hair",
        BorderStyle::MediumDashed => "mediumDashed",
        BorderStyle::DashDot => "dashDot",
        BorderStyle::MediumDashDot => "mediumDashDot",
        BorderStyle::DashDotDot => "dashDotDot",
        BorderStyle::MediumDashDotDot => "mediumDashDotDot",
        BorderStyle::SlantDashDot => "slantDashDot",
    }))
}

fn side_color(side: Option<&xlsb::package::styles::BorderSide>) -> Option<Value<'static>> {
    let side = side?;
    if side.style == xlsb::package::styles::BorderStyle::None {
        return None;
    }
    side.color.map(argb_color)
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
