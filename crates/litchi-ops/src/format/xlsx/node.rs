//! XLSX typed nodes and their node view (record 0010 §2–§3, §4.7).
//!
//! A [`Node`] is one element of an open workbook: the workbook itself, a
//! sheet, a cell, a range, a row, a column, a table or a defined name. It
//! borrows the [`Xlsx`] handle and carries the leaf values it reads:
//!
//! - a sheet keeps its worksheet handle, and pins the worksheet's parsed
//!   grid once a read of its rows or cells needs it;
//! - a cell resolved by a path keeps the leaf's owned view of its
//!   coordinate, read through the leaf's own cell route (which streams a
//!   part too large for the grid cache); a cell listed under a sheet, row,
//!   column or range keeps the grid it was listed from and its coordinate,
//!   and reads its record by lookup, so listing allocates nothing per cell;
//! - rows, columns and ranges keep their worksheet's grid;
//! - a table keeps its parsed table part.
//!
//! Nothing is retained beyond the node: the handle holds no worksheet,
//! grid or parsed value for the unified layer, so the leaf's weighted
//! caches stay the only retention policy (record 0010 §3.3).
//!
//! A node is small, so a walk moves it cheaply (record 0262): its path is
//! written from its coordinates when asked for, its A1 label the first time
//! a preview needs it, and a cell read by `get` keeps the leaf's view boxed.
//!
//! The view's `format` member reads typed values straight from those leaf
//! values with OfficeCLI's keys (`spec`), local formatting only. A cell's
//! formatting is resolved once per node, through the workbook's format table
//! without re-reading the source.

use std::borrow::Cow;
use std::fmt;
use std::sync::{Arc, OnceLock};

use litchi_core::selector::CellRef;
use litchi_core::{FileFormat, Position};
use litchi_xlsx::style::Format as CellFormat;
use litchi_xlsx::style::cell_format::{
    Border, Color as StyleColor, Fill, Horizontal, Reading, Script, Side, Underline, Vertical,
};
use litchi_xlsx::workbook::worksheet::TablePart;
use litchi_xlsx::{self as xlsx, Address, Rect};

use super::children::{Children, coordinate};
use super::spec::{
    CELL_PROPS, COLUMN_PROPS, CellKey, ColumnKey, Kind, NAMED_RANGE_PROPS, NamedRangeKey,
    RANGE_PROPS, ROW_PROPS, RangeKey, RowKey, SHEET_PROPS, SheetKey, TABLE_PROPS, TableKey,
    WORKBOOK_PROPS, WorkbookKey,
};
use super::{Xlsx, XlsxOwner};
use crate::error::{Code, Error, Result};
pub(super) use crate::format::workbook::Label;
use crate::node::{
    Children as ViewChildren, Format, NodePath, PathStep, Preview, Project, Source, Value,
    ViewOptions,
};
use crate::value::{Color, Length, Rgba, ThemeSlot};

/// The text OfficeCLI shows for a coordinate with no record.
pub(super) const MISSING_TEXT: &str = "(empty)";

/// A node of an XLSX workbook (record 0010 §3.1).
///
/// Get one from the unified `get`, `find`, `query` and `children` reads; it
/// borrows the [`Xlsx`] handle it came from. The generic accessors live on
/// [`crate::node::Node`]; the typed ones here say which element the node is
/// and where it sits.
///
/// ```
/// use litchi_ops::format::xlsx::Node;
///
/// // The worksheet and coordinate of a cell node, as `Sheet1!B3`.
/// fn address(node: &Node<'_>) -> Option<String> {
///     let (column, row) = node.cell()?;
///     let sheet = match node.sheet_name() {
///         Ok(Some(sheet)) => sheet,
///         Ok(None) | Err(_) => return None,
///     };
///     Some(format!("{sheet}!{}{}", column_letters(column), row + 1))
/// }
/// # fn column_letters(column: u32) -> String { ((b'A' + column as u8) as char).to_string() }
/// # let _ = address;
/// ```
#[derive(Clone)]
pub struct Node<'f> {
    file: &'f Xlsx,
    item: Item,
    /// The position of the node's worksheet (0 for the workbook and a
    /// defined name).
    sheet: u32,
}

// A walk moves every node it visits: keep it small (record 0262; 248 bytes
// before, with its path steps and label kept inline).
const _: () = assert!(size_of::<Node<'static>>() <= 96);

/// The element a node is, with the leaf values it reads.
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
    /// A worksheet chart, its series and its axes (record 0270).
    Chart(ChartItem),
    Series(ChartItem),
    Axis(ChartItem),
    /// A worksheet feature: a note, hyperlink, picture, shape, conditional
    /// format, validation, auto-filter or page break (record 0320).
    Feature(Arc<super::features::Feature>),
    /// A rich-text run of a cell (record 0338).
    Run(RunItem),
}

/// A rich-text run: its cell, its position among the cell's runs, and the
/// cell's runs as read.
#[derive(Clone)]
pub(super) struct RunItem {
    pub(super) at: Address,
    pub(super) index: u32,
    pub(super) rich: Arc<xlsx::RichText>,
}

/// A worksheet chart read from its part: the chart's position on the sheet,
/// and, for a series or axis, the child's position in the summary.
#[derive(Clone)]
pub(crate) struct ChartItem {
    pub(super) chart: u32,
    pub(super) part: Arc<crate::format::chart::LazySummary>,
    pub(super) child: u32,
}

impl ChartItem {
    /// The chart part's projection, read on first use.
    pub(super) fn summary(&self) -> Result<&Arc<crate::format::chart::Summary>> {
        self.part.get()
    }

    pub(super) fn with_child(&self, child: usize) -> Self {
        Self {
            chart: self.chart,
            part: Arc::clone(&self.part),
            child: position_u32(child),
        }
    }
}

/// A sheet: its position, its leaf handle, and its parsed grid once a read
/// needed it.
#[derive(Clone)]
pub(super) struct Sheet {
    pub(super) position: usize,
    pub(super) handle: SheetHandle,
    grid: OnceLock<Grid>,
    /// The fence key of the handle the sheet was read from
    /// ([`super::XlsxOwner::fence_key`]). A sheet lives only inside a node
    /// that borrows that handle, so the key names a live owner.
    fence: usize,
}

impl Sheet {
    pub(super) fn new(file: &Xlsx, position: usize, handle: SheetHandle) -> Self {
        Self {
            position,
            handle,
            grid: OnceLock::new(),
            fence: file.0.owner.fence_key(),
        }
    }

    /// The sheet's parsed grid, materialized and pinned on first use.
    ///
    /// # Errors
    ///
    /// The leaf's error for a non-grid sheet, a malformed worksheet, or a
    /// changed source.
    pub(super) fn grid(&self) -> Result<&Grid> {
        if let Some(grid) = self.grid.get() {
            // A source-backed grid is an immutable snapshot. Re-enter the
            // leaf route even when the snapshot is already pinned so its
            // source-version fence still applies to cached reads, except
            // inside a resolver or query step for this sheet's handle, whose
            // closing check observes the same source after the step (record
            // 0262; record 0310: a header-row query read the source seven
            // times per probed row through this re-entry).
            if let SheetHandle::Source { sheet, .. } = &self.handle
                && !crate::path::fence::is_active(self.fence)
            {
                let _ = sheet.grid()?;
            }
            return Ok(grid);
        }
        let grid = self.handle.grid()?;
        Ok(self.grid.get_or_init(|| grid))
    }
}

/// A cell: its coordinate, where its state is read from, its A1 label once
/// a preview needed it, and its formatting once read.
#[derive(Clone)]
pub(super) struct CellItem {
    pub(super) at: Address,
    pub(super) source: CellSource,
    label: OnceLock<Label>,
    format: OnceLock<Option<Box<CellFormat>>>,
}

/// Where a cell node reads its state.
#[derive(Clone)]
pub(super) enum CellSource {
    /// The leaf's owned view of the coordinate, read by `get` through the
    /// leaf's cell route, with the sheet for formatting reads.
    View(Box<ViewSource>),
    /// A record of a pinned grid, listed under a sheet, row, column or
    /// range, at its position in the grid's record order: read in O(1).
    Record { grid: Grid, index: u32 },
    /// A record of a pinned grid, read by its coordinate.
    Grid(Grid),
    /// A coordinate with no record, listed by a dense range view.
    Stub,
}

/// What `get` read for a cell through the leaf's cell route.
#[derive(Clone)]
pub(super) struct ViewSource {
    pub(super) handle: SheetHandle,
    pub(super) view: xlsx::SourceCellView,
}

/// A range: its worksheet's grid and its normalized rectangle.
#[derive(Clone)]
pub(super) struct RangeItem {
    pub(super) grid: Grid,
    pub(super) rect: Rect,
    label: OnceLock<Label>,
}

/// A row: its worksheet's grid and its zero-based index.
#[derive(Clone)]
pub(super) struct RowItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    /// Listed by a worksheet's node view, whose bulk listing shows only the
    /// records with content, as OfficeCLI's `GetSheetChildNodes` does; a row
    /// resolved by `get` or matched by a query lists every record, as
    /// OfficeCLI's row path does.
    pub(super) bulk: bool,
    label: OnceLock<Label>,
}

/// A column: its worksheet's grid and its zero-based index.
#[derive(Clone)]
pub(super) struct ColumnItem {
    pub(super) grid: Grid,
    pub(super) index: u32,
    label: OnceLock<Label>,
}

/// A table: its parsed part and its position among its worksheet's tables.
#[derive(Clone)]
pub(super) struct TableItem {
    pub(super) part: Box<TablePart>,
    pub(super) index: usize,
}

/// One sheet of either owner.
#[derive(Clone)]
pub(super) enum SheetHandle {
    Eager {
        sheet: xlsx::Worksheet,
        workbook: Xlsx,
    },
    Source {
        sheet: xlsx::SourceWorksheet,
    },
}

impl SheetHandle {
    /// The sheet `position` of the workbook, or `None` past the last tab.
    pub(super) fn at(file: &Xlsx, position: usize) -> Result<Option<Self>> {
        Ok(match &file.0.owner {
            XlsxOwner::Source(workbook) => workbook
                .sheet(position)?
                .map(|sheet| Self::Source { sheet }),
            XlsxOwner::Eager(workbook) => {
                workbook
                    .workbook_ref()
                    .sheet(position)?
                    .map(|sheet| Self::Eager {
                        sheet,
                        workbook: file.clone(),
                    })
            },
        })
    }

    /// The sheet named `name`, matched with the leaf's canonical
    /// case-insensitive key (record 0010 §4.2, decision D6).
    pub(super) fn named(file: &Xlsx, name: &str) -> Result<Option<Self>> {
        Ok(match &file.0.owner {
            XlsxOwner::Source(workbook) => {
                workbook.sheet(name)?.map(|sheet| Self::Source { sheet })
            },
            XlsxOwner::Eager(workbook) => {
                workbook
                    .workbook_ref()
                    .sheet(name)?
                    .map(|sheet| Self::Eager {
                        sheet,
                        workbook: file.clone(),
                    })
            },
        })
    }

    pub(super) fn name(&self) -> &str {
        match self {
            Self::Eager { sheet, .. } => sheet.name(),
            Self::Source { sheet } => sheet.name(),
        }
    }

    pub(super) fn position(&self) -> usize {
        match self {
            Self::Eager { sheet, .. } => sheet.position(),
            Self::Source { sheet, .. } => sheet.position(),
        }
    }

    pub(super) fn kind(&self) -> xlsx::WorksheetKind {
        match self {
            Self::Eager { sheet, .. } => sheet.kind(),
            Self::Source { sheet, .. } => sheet.kind(),
        }
    }

    /// Whether the sheet has a cell grid (a worksheet, not a chart or
    /// dialog sheet).
    pub(super) fn is_grid(&self) -> bool {
        self.kind() == xlsx::WorksheetKind::Worksheet
    }

    fn visibility(&self) -> &xlsx::Visibility {
        match self {
            Self::Eager { sheet, .. } => sheet.visibility(),
            Self::Source { sheet, .. } => sheet.visibility(),
        }
    }

    /// The leaf's view of one coordinate, through its own cell route.
    pub(super) fn cell(&self, at: Address) -> Result<xlsx::SourceCellView> {
        match self {
            Self::Source { sheet, .. } => Ok(sheet.cell(at)?),
            Self::Eager { sheet, .. } => match sheet.cell(at)? {
                xlsx::cell::View::Missing => Ok(xlsx::SourceCellView::Missing),
                xlsx::cell::View::Covered(range) => Ok(xlsx::SourceCellView::Covered(range)),
                xlsx::cell::View::Stored(cell) => Ok(xlsx::SourceCellView::Stored(cell.clone())),
                _ => Err(Error::from(xlsx::Error::Unsupported {
                    feature: "unrecognized worksheet cell-view variant",
                })),
            },
        }
    }

    /// The worksheet's tables, in the leaf's order.
    pub(super) fn tables(&self) -> Result<Vec<TablePart>> {
        Ok(match self {
            Self::Eager { sheet, .. } => sheet.tables()?,
            Self::Source { sheet, .. } => sheet.tables()?,
        })
    }

    /// The worksheet's parsed grid. Reads the worksheet part only; the
    /// styles part is read by the first format request.
    pub(super) fn grid(&self) -> Result<Grid> {
        match self {
            Self::Eager { sheet, .. } => {
                // Parses the store once (the snapshot keeps it) and refuses a
                // non-grid sheet, as every later read would.
                sheet.extents()?;
                Ok(Grid(Arc::new(GridInner::Eager(sheet.clone()))))
            },
            Self::Source { sheet, .. } => Ok(Grid(Arc::new(GridInner::Source {
                grid: sheet.grid()?,
                sheet: sheet.clone(),
            }))),
        }
    }

    /// The effective shared format used for fill-policy admission.
    ///
    /// The ordinary node projection intentionally omits the workbook base
    /// format. Fill writes need that format because an unstored coordinate can
    /// still inherit a non-default pattern from `cellXfs[0]`; keep this lookup
    /// private to the write adapter rather than changing read semantics.
    pub(super) fn effective_format(&self, at: Address) -> Result<Option<CellFormat>> {
        match self {
            Self::Eager { sheet, workbook } => match sheet.local_style(at)? {
                Some(xlsx::LocalStyle::Shared(style)) => Ok(Some(style.format()?)),
                Some(xlsx::LocalStyle::Default) | None => base_format(workbook),
                Some(_) => Err(Error::new(
                    Code::UnsupportedCapability,
                    "the native local style kind is not supported for fill edits",
                )
                .with_format(FileFormat::Xlsx)),
            },
            Self::Source { sheet, .. } => {
                let style = sheet.style_key(at)?;
                Ok(sheet.format(style.as_ref())?)
            },
        }
    }
}

/// The parsed grid of one worksheet, pinned while a node holds it. One
/// `Arc`, so listing a cell costs one reference-count increment.
#[derive(Clone)]
pub(super) struct Grid(Arc<GridInner>);

enum GridInner {
    /// The eager owner's worksheet, whose snapshot keeps its parsed store.
    Eager(xlsx::Worksheet),
    /// The source-backed owner's pinned grid and worksheet handle. Formatting
    /// uses the handle's lineage-bound style keys instead of a native index.
    Source {
        grid: xlsx::workbook::source::Grid,
        sheet: xlsx::SourceWorksheet,
    },
    /// A leading part of a source-backed worksheet, which the leaf's cell
    /// cursor parsed (record 0267). Its records are the complete grid's
    /// first records at the same positions; everything else (a coordinate's
    /// view, a range scan, rows, merges, formatting) reads the complete
    /// grid, materialized on first use.
    Prefix {
        prefix: xlsx::workbook::source::Grid,
        sheet: xlsx::SourceWorksheet,
        complete: OnceLock<xlsx::workbook::source::Grid>,
    },
}

impl GridInner {
    /// The leaf grid that answers coordinate reads.
    fn leaf(&self) -> Result<Option<&xlsx::workbook::source::Grid>> {
        Ok(match self {
            Self::Eager(_) => None,
            Self::Source { grid, .. } => Some(grid),
            Self::Prefix {
                sheet, complete, ..
            } => Some(match complete.get() {
                Some(grid) => grid,
                None => {
                    let grid = sheet.grid()?;
                    complete.get_or_init(|| grid)
                },
            }),
        })
    }
}

impl Grid {
    /// The node grid of a leaf grid a cell cursor read: the complete grid,
    /// or a prefix of it.
    pub(super) fn streamed(
        leaf: xlsx::workbook::source::Grid,
        sheet: xlsx::SourceWorksheet,
        complete: bool,
    ) -> Self {
        Self(Arc::new(if complete {
            GridInner::Source { grid: leaf, sheet }
        } else {
            GridInner::Prefix {
                prefix: leaf,
                sheet,
                complete: OnceLock::new(),
            }
        }))
    }

    /// The record stored at `at`, as stored (a merged follower's included).
    pub(super) fn record(&self, at: Address) -> Result<Option<&xlsx::Cell>> {
        Ok(self.cells(Rect::single(at))?.next().map(|(_, cell)| cell))
    }

    /// The record at `index` of the grid's record order, with its
    /// coordinate, in O(1).
    pub(super) fn record_at(
        &self,
        index: xlsx::RecordIndex,
    ) -> Result<Option<(Address, &xlsx::Cell)>> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.record(index)?,
            GridInner::Source { grid, .. } => grid.record(index),
            GridInner::Prefix { prefix, .. } => prefix.record(index),
        })
    }

    /// The state of one coordinate: missing, covered by a merge, or stored.
    pub(super) fn view(&self, at: Address) -> Result<xlsx::cell::View<'_>> {
        Ok(match (&*self.0, self.0.leaf()?) {
            (GridInner::Eager(sheet), _) => sheet.cell(at)?,
            (_, Some(grid)) => grid.cell(at)?,
            (_, None) => return Err(lost_grid()),
        })
    }

    /// The stored records of `rect`, in row-major order.
    pub(super) fn cells(&self, rect: Rect) -> Result<xlsx::Cells<'_>> {
        Ok(match (&*self.0, self.0.leaf()?) {
            (GridInner::Eager(sheet), _) => sheet.cells(rect)?,
            (_, Some(grid)) => grid.cells(rect)?,
            (_, None) => return Err(lost_grid()),
        })
    }

    /// The row with zero-based `index`, stored or implicit.
    pub(super) fn row(&self, index: u32) -> Result<xlsx::Row<'_>> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.row(index)?,
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => {
                sheet.row(index)?
            },
        })
    }

    /// The explicit row records.
    pub(super) fn rows(&self) -> Result<xlsx::Rows<'_>> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.rows()?,
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => sheet.rows()?,
        })
    }

    /// The column with zero-based `index`, stored or implicit.
    pub(super) fn column(&self, index: u32) -> Result<xlsx::Column<'_>> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.column(index)?,
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => {
                sheet.column(index)?
            },
        })
    }

    /// The logical columns covered by explicit property records.
    pub(super) fn columns(&self) -> Result<xlsx::Columns<'_>> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.columns()?,
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => sheet.columns()?,
        })
    }

    /// The merged ranges.
    pub(super) fn merges(&self) -> Result<xlsx::merge::Merges<'_>> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.merges()?,
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => sheet.merges()?,
        })
    }

    /// Whether `rect` is one exact native merge.
    pub(super) fn is_exact_merge(&self, rect: Rect) -> Result<bool> {
        if rect.rows() == 1 && rect.columns() == 1 {
            return Ok(false);
        }

        let (end_row, end_column) = rect.end();
        let endpoint = Address::at(end_row - 1, end_column - 1).map_err(coordinate)?;
        // A non-single-cell rectangle's inclusive bottom-right is a non-anchor
        // cell when the rectangle is an actual merge, so native indexed
        // coverage returns the exact merge without scanning all merge records.
        Ok(self.view(endpoint)?.merge() == Some(rect))
    }

    /// The cell bounds.
    pub(super) fn extents(&self) -> Result<&xlsx::Extents> {
        Ok(match &*self.0 {
            GridInner::Eager(sheet) => sheet.extents()?,
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => sheet.extents()?,
        })
    }

    /// The local formatting of the record at `at`: `None` when the record
    /// uses the base format, has no record, or the workbook has no styles.
    pub(super) fn format(&self, at: Address) -> Result<Option<CellFormat>> {
        match &*self.0 {
            GridInner::Eager(sheet) => match sheet.local_style(at)? {
                Some(xlsx::LocalStyle::Shared(style)) => Ok(Some(style.format()?)),
                _ => Ok(None),
            },
            GridInner::Source { sheet, .. } | GridInner::Prefix { sheet, .. } => {
                let Some(style) = sheet.style_key(at)? else {
                    return Ok(None);
                };
                // Borrow the opaque, lineage-bound key for the leaf lookup.
                // The leaf owns the format-table cache and retains typed
                // failures instead of caching an error as `None`.
                Ok(sheet.format(Some(&style))?)
            },
        }
    }
}

fn lost_grid() -> Error {
    Error::new(Code::InternalError, "a worksheet grid lost its leaf grid")
        .with_format(FileFormat::Xlsx)
}

fn base_format(file: &Xlsx) -> Result<Option<CellFormat>> {
    match &file.0.owner {
        XlsxOwner::Source(workbook) => Ok(workbook.format(None)?),
        XlsxOwner::Eager(workbook) => workbook
            .workbook_ref()
            .styles()?
            .base()
            .map(|style| style.format())
            .transpose()
            .map_err(Error::from),
    }
}

/// A position as stored in a node: worksheet positions are below the
/// catalog's length, which no workbook brings near `u32::MAX`.
fn position_u32(position: usize) -> u32 {
    u32::try_from(position).unwrap_or(u32::MAX)
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

impl<'f> Node<'f> {
    /// The workbook root, `/`.
    pub(super) fn new_workbook(file: &'f Xlsx) -> Self {
        Self {
            file,
            item: Item::Workbook,
            sheet: 0,
        }
    }

    /// A sheet, `/sheet[N]`.
    pub(super) fn new_sheet(file: &'f Xlsx, sheet: Sheet) -> Self {
        let position = position_u32(sheet.position);
        Self {
            file,
            item: Item::Sheet(sheet),
            sheet: position,
        }
    }

    /// An element of the worksheet at `sheet`, for a sibling module.
    pub(super) fn in_sheet_item(file: &'f Xlsx, sheet: usize, item: Item) -> Self {
        Self::in_sheet(file, sheet, item)
    }

    /// An element of the worksheet at `sheet`, `/sheet[N]/<tail>`.
    fn in_sheet(file: &'f Xlsx, sheet: usize, item: Item) -> Self {
        Self {
            file,
            item,
            sheet: position_u32(sheet),
        }
    }

    /// A cell of the worksheet at `sheet`.
    pub(super) fn new_cell(file: &'f Xlsx, sheet: usize, at: Address, source: CellSource) -> Self {
        let item = Item::Cell(CellItem {
            at,
            source,
            label: OnceLock::new(),
            format: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
    }

    /// The range `rect` of the worksheet at `sheet`.
    pub(super) fn new_range(file: &'f Xlsx, sheet: usize, grid: Grid, rect: Rect) -> Self {
        let item = Item::Range(RangeItem {
            grid,
            rect,
            label: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
    }

    /// The row with zero-based `index` of the worksheet at `sheet`.
    pub(super) fn new_row(file: &'f Xlsx, sheet: usize, grid: Grid, index: u32) -> Self {
        Self::row(file, sheet, grid, index, false)
    }

    /// A row of a worksheet's node view listing (see [`RowItem::bulk`]).
    pub(super) fn new_listed_row(file: &'f Xlsx, sheet: usize, grid: Grid, index: u32) -> Self {
        Self::row(file, sheet, grid, index, true)
    }

    fn row(file: &'f Xlsx, sheet: usize, grid: Grid, index: u32, bulk: bool) -> Self {
        let item = Item::Row(RowItem {
            grid,
            index,
            bulk,
            label: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
    }

    /// The column with zero-based `index` of the worksheet at `sheet`.
    pub(super) fn new_column(file: &'f Xlsx, sheet: usize, grid: Grid, index: u32) -> Self {
        let item = Item::Column(ColumnItem {
            grid,
            index,
            label: OnceLock::new(),
        });
        Self::in_sheet(file, sheet, item)
    }

    /// The table at `index` among the tables of the worksheet at `sheet`.
    pub(super) fn new_table(file: &'f Xlsx, sheet: usize, index: usize, part: TablePart) -> Self {
        let item = Item::Table(TableItem {
            part: Box::new(part),
            index,
        });
        Self::in_sheet(file, sheet, item)
    }

    /// The chart at `index` of the worksheet at `sheet`.
    pub(super) fn new_chart(
        file: &'f Xlsx,
        sheet: usize,
        index: usize,
        part: Arc<crate::format::chart::LazySummary>,
    ) -> Self {
        let item = Item::Chart(ChartItem {
            chart: position_u32(index),
            part,
            child: 0,
        });
        Self::in_sheet(file, sheet, item)
    }

    /// A series or axis of a worksheet chart.
    pub(super) fn new_chart_part(
        file: &'f Xlsx,
        sheet: usize,
        kind: Kind,
        item: ChartItem,
    ) -> Self {
        let item = if kind == Kind::Series {
            Item::Series(item)
        } else {
            Item::Axis(item)
        };
        Self::in_sheet(file, sheet, item)
    }

    /// The run at `index` of the cell `at` of the worksheet at `sheet`
    /// (record 0338).
    pub(super) fn new_run(
        file: &'f Xlsx,
        sheet: usize,
        at: Address,
        index: usize,
        rich: Arc<xlsx::RichText>,
    ) -> Self {
        let item = Item::Run(RunItem {
            at,
            index: position_u32(index),
            rich,
        });
        Self::in_sheet(file, sheet, item)
    }

    /// The defined name at `index`, `/namedrange[N]`.
    pub(super) fn new_named_range(file: &'f Xlsx, index: usize) -> Self {
        Self {
            file,
            item: Item::NamedRange(index),
            sheet: 0,
        }
    }

    pub(super) fn file(&self) -> &'f Xlsx {
        self.file
    }

    pub(super) fn item(&self) -> &Item {
        &self.item
    }

    /// The position of the node's worksheet.
    pub(super) fn sheet_position(&self) -> usize {
        self.sheet as usize
    }

    /// The native worksheet position and address used by the cell write
    /// adapter. This deliberately succeeds only for one direct cell node;
    /// ranges, rows and other projections are admitted through their own
    /// target accessors.
    pub(super) fn write_cell_target(&self) -> Option<(usize, Address)> {
        match &self.item {
            Item::Cell(cell) => Some((self.sheet_position(), cell.at)),
            _ => None,
        }
    }

    /// The native worksheet position and rectangle used by a range write.
    ///
    /// The rectangle is normalized by path resolution and is retained by the
    /// node's grid view, so a write can include coordinates with no stored
    /// cell record (the same dense range semantics as reads).
    pub(super) fn write_range_target(&self) -> Option<(usize, Rect)> {
        match &self.item {
            Item::Range(range) => Some((self.sheet_position(), range.rect)),
            _ => None,
        }
    }

    /// The node's last path step: its own step under its parent, written
    /// from its coordinates.
    pub(super) fn last_step(&self) -> Option<PathStep<'static>> {
        Some(match &self.item {
            Item::Workbook => return None,
            Item::Sheet(_) => PathStep::at("sheet", Position::new(self.sheet_position())),
            Item::Cell(cell) => PathStep::cell(cell.at.column().get(), cell.at.row().get()),
            Item::Range(range) => {
                let (first, last) = corners(range.rect);
                PathStep::range(first, last)
            },
            Item::Row(row) => PathStep::at("row", Position::new(row.index as usize)),
            Item::Column(column) => PathStep::column("col", column.index),
            Item::Table(table) => PathStep::at("table", Position::new(table.index)),
            Item::NamedRange(index) => PathStep::at("namedrange", Position::new(*index)),
            Item::Chart(chart) => PathStep::at("chart", Position::new(chart.chart as usize)),
            Item::Series(chart) => PathStep::at("series", Position::new(chart.child as usize)),
            Item::Axis(chart) => PathStep::at("axis", Position::new(chart.child as usize)),
            Item::Feature(feature) => feature.step(),
            Item::Run(run) => PathStep::at("run", Position::new(run.index as usize)),
        })
    }

    /// The A1 label of a cell, range, row or column node, written on first
    /// use.
    fn label(&self) -> Option<&Label> {
        Some(match &self.item {
            Item::Cell(cell) => cell
                .label
                .get_or_init(|| Label::cell(cell.at.column().get(), cell.at.row().get())),
            Item::Range(range) => range.label.get_or_init(|| {
                let (first, last) = corners(range.rect);
                Label::range(first, last)
            }),
            Item::Row(row) => row.label.get_or_init(|| Label::row(row.index)),
            Item::Column(column) => column.label.get_or_init(|| Label::column(column.index)),
            _ => return None,
        })
    }

    /// The element kind.
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
            Item::Chart(_) => Kind::Chart,
            Item::Series(_) => Kind::Series,
            Item::Axis(_) => Kind::Axis,
            Item::Feature(ref feature) => feature.kind,
            Item::Run(_) => Kind::Run,
        }
    }

    /// The OfficeCLI `type` of the node: `workbook`, `sheet`, `cell`,
    /// `range`, `row`, `column`, `table` or `namedrange`.
    ///
    /// ```
    /// use litchi_ops::format::xlsx::Node;
    ///
    /// fn is_cell(node: &Node<'_>) -> bool {
    ///     node.kind() == "cell"
    /// }
    /// # let _ = is_cell;
    /// ```
    #[must_use]
    pub fn kind(&self) -> &'static str {
        self.element().spec().name
    }

    /// The name of the sheet the node belongs to (`None` for the workbook
    /// and a defined name).
    ///
    /// # Errors
    ///
    /// Returns a source-version error when a source-backed workbook changed.
    ///
    /// ```
    /// use litchi_ops::format::xlsx::Node;
    ///
    /// fn sheet_of<'f>(node: &Node<'f>) -> litchi_ops::Result<Option<&'f str>> {
    ///     node.sheet_name()
    /// }
    /// # let _ = sheet_of;
    /// ```
    pub fn sheet_name(&self) -> Result<Option<&'f str>> {
        self.file.0.owner.finish(|| {
            Ok(match self.item {
                Item::Workbook | Item::NamedRange(_) => None,
                _ => Some(sheet_name(self.file, self.sheet_position())),
            })
        })
    }

    /// The zero-based `(column, row)` of a cell node.
    ///
    /// ```
    /// use litchi_ops::format::xlsx::Node;
    ///
    /// fn in_first_row(node: &Node<'_>) -> bool {
    ///     node.cell().is_some_and(|(_, row)| row == 0)
    /// }
    /// # let _ = in_first_row;
    /// ```
    #[must_use]
    pub fn cell(&self) -> Option<(u32, u32)> {
        match &self.item {
            Item::Cell(cell) => Some((cell.at.column().get(), cell.at.row().get())),
            _ => None,
        }
    }

    /// The cell's native local formatting, borrowed from this node's cache.
    ///
    /// This exposes the leaf's typed values, including gradient geometry,
    /// ordered stop positions and theme/indexed colors that the ordinary
    /// OfficeCLI-shaped format map omits. It does not resolve inheritance or
    /// theme colors. Cells without an explicit local style, missing cells and
    /// other node kinds return `None`. An explicit reference to the base style
    /// returns its format, whose [`is_base`](CellFormat::is_base) is true.
    ///
    /// # Errors
    ///
    /// Returns the leaf's stylesheet, resource-limit or source-version error.
    /// Cached reads still check the source version.
    ///
    /// ```
    /// use litchi_ops::format::xlsx::Node;
    ///
    /// fn gradient_degree(node: &Node<'_>) -> litchi_ops::Result<Option<f64>> {
    ///     Ok(node.cell_format()?
    ///         .and_then(|format| format.fill())
    ///         .and_then(|fill| fill.gradient())
    ///         .map(|gradient| gradient.degree()))
    /// }
    /// # let _ = gradient_degree;
    /// ```
    pub fn cell_format(&self) -> Result<Option<&CellFormat>> {
        self.file.0.owner.finish(|| match &self.item {
            Item::Cell(cell) => cell.format(),
            _ => Ok(None),
        })
    }

    /// The 1-based number of a row node (`row[N]`) or a column node
    /// (`col[B]` is 2).
    ///
    /// ```
    /// use litchi_ops::format::xlsx::Node;
    ///
    /// fn is_header_row(node: &Node<'_>) -> bool {
    ///     node.kind() == "row" && node.number() == Some(1)
    /// }
    /// # let _ = is_header_row;
    /// ```
    #[must_use]
    pub fn number(&self) -> Option<u32> {
        match &self.item {
            Item::Row(row) => Some(row.index + 1),
            Item::Column(column) => Some(column.index + 1),
            _ => None,
        }
    }

    /// The selector grammar's form of a cell node's coordinate.
    pub(super) fn cell_ref(&self) -> Option<CellRef> {
        match &self.item {
            Item::Cell(_) => self
                .label()
                .and_then(|label| CellRef::parse(label.as_str())),
            _ => None,
        }
    }

    /// The cell's state as the node view and queries read it.
    pub(super) fn cell_state(&self) -> Result<Option<CellState<'_>>> {
        self.file.0.owner.finish(|| {
            let Item::Cell(cell) = &self.item else {
                return Ok(None);
            };
            cell.state().map(Some)
        })
    }

    /// The display text of a stored record, a missing coordinate or a stub.
    pub(super) fn cell_text(&self) -> Result<Option<&str>> {
        Ok(self.cell_state()?.map(|state| match state {
            CellState::Stored(cell) => display_text(cell),
            CellState::Missing | CellState::Covered(_) => MISSING_TEXT,
            CellState::Stub => "",
        }))
    }

    /// The `format` entries, reading the leaf values the node holds.
    pub(super) fn props(&self) -> Result<Format<'_>> {
        self.file.0.owner.finish(|| self.props_unfenced())
    }

    fn props_unfenced(&self) -> Result<Format<'_>> {
        let file = self.file;
        let (table, props) = match &self.item {
            Item::Workbook => (WORKBOOK_PROPS, Props::Workbook(file)),
            Item::Sheet(sheet) => (
                SHEET_PROPS,
                Props::Sheet {
                    name: sheet_name(file, sheet.position),
                    visibility: sheet.handle.visibility(),
                },
            ),
            Item::Cell(cell) => (CELL_PROPS, Props::Cell(cell.props()?)),
            Item::Range(range) => (
                RANGE_PROPS,
                Props::Range(range.grid.is_exact_merge(range.rect)?),
            ),
            Item::Row(row) => (ROW_PROPS, Props::Row(row.grid.row(row.index)?)),
            Item::Column(column) => (
                COLUMN_PROPS,
                Props::Column(column.grid.column(column.index)?),
            ),
            Item::Table(table) => (TABLE_PROPS, Props::Table(table.part.table())),
            Item::Chart(chart) => (
                super::chart::CHART_CAPABILITY_PROPS,
                Props::Chart(
                    Arc::clone(chart.summary()?),
                    file.sheet_chart_cells(self.sheet_position(), chart.chart as usize)?,
                ),
            ),
            Item::Series(chart) => (super::chart::SERIES_PROPS, Props::Series(chart.clone())),
            Item::Feature(feature) => (
                feature.table(),
                Props::Feature(feature, sheet_name(file, self.sheet_position())),
            ),
            Item::Axis(chart) => (super::chart::AXIS_PROPS, Props::Axis(chart.clone())),
            Item::Run(run) => match run.run() {
                Some(value) => (
                    super::runs::RUN_PROPS,
                    Props::Run(Box::new(value.font().map_err(Error::from)?)),
                ),
                None => return Ok(Format::empty()),
            },
            Item::NamedRange(index) => match defined_names(file).get(*index) {
                Some(name) => (
                    NAMED_RANGE_PROPS,
                    Props::NamedRange {
                        name,
                        scope: name
                            .local_sheet_id
                            .and_then(|id| file.names().get(id as usize))
                            .map(String::as_str),
                    },
                ),
                None => return Ok(Format::empty()),
            },
        };
        Ok(Format::new(table, Source::Xlsx(props)))
    }
}

impl fmt::Debug for Node<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("xlsx::Node")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

/// The sheet name at `position`, borrowed from the handle's catalog.
pub(super) fn sheet_name(file: &Xlsx, position: usize) -> &str {
    file.names().get(position).map_or("", String::as_str)
}

/// The workbook's defined names, in catalog order.
pub(super) fn defined_names(file: &Xlsx) -> &[xlsx::raw::DefinedName] {
    match &file.0.owner {
        XlsxOwner::Source(workbook) => workbook.defined_names(),
        XlsxOwner::Eager(workbook) => workbook.workbook_ref().defined_names(),
    }
}

/// The state of a cell node's coordinate.
#[derive(Clone, Copy)]
pub(super) enum CellState<'a> {
    /// A stored record.
    Stored(&'a xlsx::Cell),
    /// No record and no covering merge.
    Missing,
    /// A merged follower: the merge anchored at the range's first cell.
    Covered(Rect),
    /// A dense-listing stub.
    Stub,
}

impl CellItem {
    pub(super) fn state(&self) -> Result<CellState<'_>> {
        Ok(match &self.source {
            CellSource::View(source) => match &source.view {
                xlsx::SourceCellView::Stored(cell) => CellState::Stored(cell),
                xlsx::SourceCellView::Covered(range) => CellState::Covered(*range),
                _ => CellState::Missing,
            },
            // One point lookup in the common case; a merged follower the
            // leaf reports as covered is read as stored when it has a record.
            CellSource::Grid(grid) => match grid.view(self.at)? {
                xlsx::cell::View::Stored(cell) => CellState::Stored(cell),
                xlsx::cell::View::Covered(range) => match grid.record(self.at)? {
                    Some(cell) => CellState::Stored(cell),
                    None => CellState::Covered(range),
                },
                _ => CellState::Missing,
            },
            // A listed record is stored: the coordinate's view would read
            // it, directly or as a merged follower's own record.
            CellSource::Record { grid, index } => {
                match grid.record_at(xlsx::RecordIndex::new(*index as usize))? {
                    Some((_, cell)) => CellState::Stored(cell),
                    None => {
                        return Err(Error::new(
                            Code::InternalError,
                            "a listed cell record vanished from its grid",
                        )
                        .with_format(FileFormat::Xlsx));
                    },
                }
            },
            CellSource::Stub => CellState::Stub,
        })
    }

    /// The cell's local formatting, resolved once per node.
    fn format(&self) -> Result<Option<&CellFormat>> {
        if let Some(format) = self.format.get() {
            return Ok(format.as_deref());
        }
        let format = match &self.source {
            CellSource::View(source) => match &source.view {
                xlsx::SourceCellView::Stored(_) => source.handle.grid()?.format(self.at)?,
                _ => None,
            },
            CellSource::Grid(grid) | CellSource::Record { grid, .. } => grid.format(self.at)?,
            CellSource::Stub => None,
        };
        Ok(self.format.get_or_init(|| format.map(Box::new)).as_deref())
    }

    /// The merge the cell belongs to, and whether it is its anchor.
    fn merge(&self, state: CellState<'_>) -> Result<Option<(Rect, bool)>> {
        let grid;
        let merges = match (&self.source, state) {
            (_, CellState::Covered(range)) => return Ok(Some((range, false))),
            // A blank anchor has no cell record but still owns merge geometry.
            // Only a synthetic stub lacks a native grid to inspect.
            (_, CellState::Stub) => return Ok(None),
            (CellSource::View(source), _) => {
                grid = source.handle.grid()?;
                grid.merges()?
            },
            (CellSource::Grid(grid) | CellSource::Record { grid, .. }, _) => grid.merges()?,
            (CellSource::Stub, _) => return Ok(None),
        };
        for range in merges {
            if range.contains(self.at) {
                return Ok(Some((range, range.start() == self.at)));
            }
        }
        Ok(None)
    }

    fn props<'a>(&'a self) -> Result<CellProps<'a>> {
        let state = self.state()?;
        let format = match state {
            CellState::Stored(_) => self.format()?,
            _ => None,
        };
        let merge = self.merge(state)?;
        Ok(CellProps {
            state,
            format,
            merge,
        })
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        Node::kind(self)
    }

    fn file_format(&self) -> FileFormat {
        FileFormat::Xlsx
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
            (Item::Run(run), Some(step)) => NodePath::inline(&[
                PathStep::at("sheet", Position::new(self.sheet_position())),
                PathStep::cell(run.at.column().get(), run.at.row().get()),
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
                Item::Cell(_) => self.cell_text()?.map(Cow::Borrowed),
                Item::Row(row) => {
                    (!row.grid.row(row.index)?.stored()).then_some(Cow::Borrowed(MISSING_TEXT))
                },
                Item::Table(table) => Some(Cow::Borrowed(table_label(table))),
                Item::NamedRange(index) => defined_names(self.file).get(*index).map(|name| {
                    Cow::Borrowed(if name.reference.is_empty() {
                        name.name.as_str()
                    } else {
                        name.reference.as_str()
                    })
                }),
                Item::Feature(feature) => feature.text(),
                Item::Run(run) => run.run().map(|run| Cow::Borrowed(run.text())),
                Item::Workbook
                | Item::Sheet(_)
                | Item::Range(_)
                | Item::Column(_)
                | Item::Chart(_)
                | Item::Series(_)
                | Item::Axis(_) => None,
            })
        })
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        self.file.0.owner.finish(|| {
            Ok(match &self.item {
                Item::Workbook => None,
                Item::Sheet(_) => Some(Preview::Text(Cow::Borrowed(sheet_name(
                    self.file,
                    self.sheet_position(),
                )))),
                Item::Cell(_) | Item::Range(_) | Item::Column(_) => self
                    .label()
                    .map(|label| Preview::Text(Cow::Borrowed(label.as_str()))),
                Item::Row(row) => match row.grid.row(row.index)?.stored() {
                    true => None,
                    false => self
                        .label()
                        .map(|label| Preview::Text(Cow::Borrowed(label.as_str()))),
                },
                Item::Table(table) => {
                    let table = table.part.table();
                    Some(Preview::Text(Cow::Owned(format!(
                        "{} ({})",
                        table.name, table.ref_range
                    ))))
                },
                Item::NamedRange(index) => defined_names(self.file)
                    .get(*index)
                    .filter(|name| !name.reference.is_empty())
                    .map(|name| Preview::Text(Cow::Borrowed(name.reference.as_str()))),
                Item::Chart(chart) => chart
                    .summary()?
                    .title
                    .as_deref()
                    .map(|title| Preview::Text(Cow::Borrowed(title))),
                Item::Series(_) | Item::Axis(_) => None,
                Item::Feature(feature) => feature.preview(),
                Item::Run(run) => run
                    .run()
                    .map(|run| Preview::Text(Cow::Borrowed(run.text()))),
            })
        })
    }

    fn child_count(&self, options: &ViewOptions) -> Result<Option<u32>> {
        self.file.0.owner.finish(|| {
            let count = match &self.item {
                Item::Workbook => self.file.0.owner.len(),
                // Counting a sheet's rows decodes its part: only when the view
                // lists them anyway (M1-Δ1).
                Item::Sheet(_) if options.depth == 0 => return Ok(None),
                Item::Cell(_)
                | Item::Table(_)
                | Item::NamedRange(_)
                | Item::Series(_)
                | Item::Axis(_)
                | Item::Feature(_)
                | Item::Run(_) => 0,
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
            Ok(children) => ViewChildren::xlsx(children),
            Err(error) => ViewChildren::failed(error),
        }
    }
}

/// A table's text: its display name, else its name.
fn table_label(table: &TableItem) -> &str {
    let table = table.part.table();
    if table.display_name.is_empty() {
        &table.name
    } else {
        &table.display_name
    }
}

/// Whether a record carries content, as OfficeCLI's `CellHasContent` reads
/// it: a formula, a value, or an explicitly typed empty value. A record with
/// no payload (`<c s="1"/>`) is formatting only.
pub(super) fn has_content(cell: &xlsx::Cell) -> bool {
    !matches!(cell, xlsx::Cell::Empty)
}

/// The text OfficeCLI shows for a stored record: the value as stored,
/// booleans as `TRUE`/`FALSE`, a formula's cached value. Number formats are
/// not applied (M1-Δ5).
pub(super) fn display_text(cell: &xlsx::Cell) -> &str {
    match cell {
        xlsx::Cell::Value(value) => value_text(value),
        xlsx::Cell::Formula(formula) => formula
            .cached()
            .map_or("", |cache| value_text(cache.value())),
        xlsx::Cell::Unknown(unknown) => unknown.value().unwrap_or(""),
        _ => "",
    }
}

/// A value as OfficeCLI displays it.
fn value_text(value: &xlsx::Value) -> &str {
    match value {
        xlsx::Value::Bool(true) => "TRUE",
        xlsx::Value::Bool(false) => "FALSE",
        xlsx::Value::Number(number) => number.as_str(),
        xlsx::Value::Text(text) => text.as_str(),
        xlsx::Value::Date(date) => date.as_str(),
        xlsx::Value::Error(error) => error.as_str(),
        _ => "",
    }
}

/// A value as stored in `<v>`: a boolean is `1` or `0`.
fn stored_text(value: &xlsx::Value) -> &str {
    match value {
        xlsx::Value::Bool(true) => "1",
        xlsx::Value::Bool(false) => "0",
        value => value_text(value),
    }
}

/// The OfficeCLI cell `type` (`ExcelHandler.Helpers.Node.cs` `CellToNode`):
/// from the stored value, and for a formula from its cached value. The XLSX
/// cell model keeps shared, inline and `t="str"` text alike, so a plain text
/// value reads as `SharedString`, which is what Excel writes.
pub(super) fn cell_type(cell: &xlsx::Cell) -> &'static str {
    let value = match cell {
        xlsx::Cell::Value(value) => value,
        xlsx::Cell::Formula(formula) => {
            return match formula.cached().map(xlsx::formula::Cache::value) {
                Some(xlsx::Value::Text(_)) => "String",
                Some(xlsx::Value::Bool(_)) => "Boolean",
                Some(xlsx::Value::Error(_)) => "Error",
                Some(xlsx::Value::Date(_)) => "Date",
                _ => "Number",
            };
        },
        xlsx::Cell::Unknown(_) => return "Unknown",
        _ => return "Number",
    };
    match value {
        xlsx::Value::Bool(_) => "Boolean",
        xlsx::Value::Text(_) => "SharedString",
        xlsx::Value::Date(_) => "Date",
        xlsx::Value::Error(_) => "Error",
        _ => "Number",
    }
}

/// Whether a record is an error cell: an error value, or a formula whose
/// cached value is one (OfficeCLI `IsExcelErrorValue`).
pub(super) fn is_error(cell: &xlsx::Cell) -> bool {
    match cell {
        xlsx::Cell::Value(xlsx::Value::Error(_)) => true,
        xlsx::Cell::Formula(formula) => matches!(
            formula.cached().map(xlsx::formula::Cache::value),
            Some(xlsx::Value::Error(_))
        ),
        _ => false,
    }
}

/// The raw comparison value of a record (OfficeCLI
/// `GetCellRawComparisonValue`): a number for a numeric value or numeric
/// cached result, else the display text.
pub(super) fn raw_value(cell: &xlsx::Cell) -> Value<'_> {
    let number = match cell {
        xlsx::Cell::Value(xlsx::Value::Number(number)) => Some(number),
        xlsx::Cell::Formula(formula) => match formula.cached().map(xlsx::formula::Cache::value) {
            Some(xlsx::Value::Number(number)) => Some(number),
            _ => None,
        },
        _ => None,
    };
    number
        .and_then(xlsx::Number::as_f64)
        .and_then(Value::number)
        .unwrap_or_else(|| Value::Text(Cow::Borrowed(display_text(cell))))
}

/// The `format` source of an XLSX node: the leaf values a node's keys read.
pub(crate) enum Props<'a> {
    Workbook(&'a Xlsx),
    Sheet {
        name: &'a str,
        visibility: &'a xlsx::Visibility,
    },
    Cell(CellProps<'a>),
    Range(bool),
    Row(xlsx::Row<'a>),
    Column(xlsx::Column<'a>),
    Table(&'a xlsx::Table),
    NamedRange {
        name: &'a xlsx::raw::DefinedName,
        scope: Option<&'a str>,
    },
    /// A chart: its part's projection and its anchor's cell geometry.
    Chart(Arc<crate::format::chart::Summary>, Option<[u32; 4]>),
    Series(ChartItem),
    Axis(ChartItem),
    /// A worksheet feature and its worksheet's name (record 0320).
    Feature(&'a super::features::Feature, &'a str),
    /// A rich-text run's font (record 0338).
    Run(Box<xlsx::style::stylesheet::font::Font>),
}

/// What a cell's keys read.
pub(crate) struct CellProps<'a> {
    state: CellState<'a>,
    format: Option<&'a CellFormat>,
    merge: Option<(Rect, bool)>,
}

impl fmt::Debug for Props<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Workbook(_) => "Workbook",
            Self::Sheet { .. } => "Sheet",
            Self::Cell(_) => "Cell",
            Self::Range(_) => "Range",
            Self::Row(_) => "Row",
            Self::Column(_) => "Column",
            Self::Table(_) => "Table",
            Self::NamedRange { .. } => "NamedRange",
            Self::Chart(..) => "Chart",
            Self::Series(_) => "Series",
            Self::Axis(_) => "Axis",
            Self::Feature(..) => "Feature",
            Self::Run(_) => "Run",
        };
        f.debug_tuple("Props").field(&name).finish()
    }
}

impl<'a> Props<'a> {
    /// The value of the key at `index` of the node's property table, if the
    /// node has it. Allocates only for text OfficeCLI composes (a merge
    /// range, a table's column list, a timestamp).
    pub(crate) fn get(&self, index: usize) -> Option<Value<'a>> {
        match *self {
            Self::Workbook(file) => workbook_value(file, WorkbookKey::at(index)?),
            Self::Sheet { name, visibility } => match SheetKey::at(index)? {
                SheetKey::Name => Some(Value::Text(Cow::Borrowed(name))),
                SheetKey::Hidden => visibility.is_hidden().then_some(Value::Bool(true)),
                SheetKey::Visibility => match visibility {
                    xlsx::Visibility::Hidden => Some(Value::Token("hidden")),
                    xlsx::Visibility::VeryHidden => Some(Value::Token("veryHidden")),
                    _ => None,
                },
            },
            Self::Cell(ref cell) => cell.get(CellKey::at(index)?),
            Self::Range(merge) => match RangeKey::at(index)? {
                RangeKey::Merge => Some(Value::Bool(merge)),
            },
            Self::Row(row) => row_value(row, RowKey::at(index)?),
            Self::Column(column) => column_value(column, ColumnKey::at(index)?),
            Self::Table(table) => table_value(table, TableKey::at(index)?),
            Self::Chart(ref summary, cells) => super::chart::chart_value_at(summary, cells, index),
            Self::Series(ref item) => super::chart::series_value(item, index).map(owned),
            Self::Axis(ref item) => super::chart::axis_value(item, index).map(owned),
            Self::Feature(feature, sheet) => feature.value(index, sheet),
            Self::Run(ref font) => super::runs::value(font, index).map(owned),
            Self::NamedRange { name, scope } => match NamedRangeKey::at(index)? {
                NamedRangeKey::Name => Some(Value::Text(Cow::Borrowed(&name.name))),
                NamedRangeKey::Ref => Some(Value::Text(Cow::Borrowed(&name.reference))),
                NamedRangeKey::Scope => match name.local_sheet_id {
                    None => Some(Value::Text(Cow::Borrowed("workbook"))),
                    Some(_) => scope.map(|scope| Value::Text(Cow::Borrowed(scope))),
                },
                NamedRangeKey::Comment => name
                    .comment
                    .as_deref()
                    .filter(|comment| !comment.is_empty())
                    .map(|comment| Value::Text(Cow::Borrowed(comment))),
                NamedRangeKey::Volatile => name.function.then_some(Value::Bool(true)),
            },
        }
    }
}

fn text(text: &str) -> Value<'_> {
    Value::Text(Cow::Borrowed(text))
}

fn optional(value: Option<&String>) -> Option<Value<'_>> {
    value.map(|value| text(value))
}

/// A core-properties timestamp in .NET's round-trip form (`"o"`), as
/// OfficeCLI writes it: seven fraction digits, and `Z` for UTC.
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

fn workbook_value(file: &Xlsx, key: WorkbookKey) -> Option<Value<'_>> {
    let metadata = &file.0.metadata;
    match key {
        WorkbookKey::Title => optional(metadata.title.as_ref()),
        WorkbookKey::Author => optional(metadata.author.as_ref()),
        WorkbookKey::Subject => optional(metadata.subject.as_ref()),
        WorkbookKey::Keywords => Value::owned_list(&metadata.keywords),
        WorkbookKey::Description => optional(metadata.description.as_ref()),
        WorkbookKey::Category => optional(metadata.category.as_ref()),
        WorkbookKey::LastModifiedBy => optional(metadata.last_modified_by.as_ref()),
        WorkbookKey::RevisionNumber => optional(metadata.revision.as_ref()),
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
        WorkbookKey::Date1904 => {
            let date_system = match &file.0.owner {
                XlsxOwner::Source(workbook) => workbook.date_system(),
                XlsxOwner::Eager(workbook) => workbook.workbook_ref().date_system(),
            };
            matches!(date_system, xlsx::DateSystem::Excel1904).then_some(Value::Bool(true))
        },
        WorkbookKey::ActiveTab => {
            let active = match &file.0.owner {
                XlsxOwner::Source(workbook) => workbook.active_sheet_position(),
                XlsxOwner::Eager(workbook) => workbook
                    .workbook_ref()
                    .active_sheet()
                    .map(|sheet| sheet.position()),
            };
            active
                .filter(|position| *position != 0)
                .map(|position| Value::Int(i64::try_from(position).unwrap_or(i64::MAX)))
        },
    }
}

fn row_value(row: xlsx::Row<'_>, key: RowKey) -> Option<Value<'static>> {
    match key {
        RowKey::Height => row
            .height()
            .map(|height| Value::Length(points(height.get()))),
        RowKey::Hidden => row.hidden().then_some(Value::Bool(true)),
        RowKey::OutlineLevel => {
            let level = row.outline().get();
            (level > 0).then_some(Value::Int(i64::from(level)))
        },
        RowKey::Collapsed => row.collapsed().then_some(Value::Bool(true)),
        RowKey::CustomHeight => row.custom_height().then_some(Value::Bool(true)),
    }
}

fn column_value(column: xlsx::Column<'_>, key: ColumnKey) -> Option<Value<'static>> {
    match key {
        ColumnKey::Width => column.width().and_then(|width| Value::number(width.get())),
        ColumnKey::Hidden => column.hidden().then_some(Value::Bool(true)),
        ColumnKey::CustomWidth => column.custom_width().then_some(Value::Bool(true)),
        ColumnKey::OutlineLevel => {
            let level = column.outline().get();
            (level > 0).then_some(Value::Int(i64::from(level)))
        },
        ColumnKey::Collapsed => column.collapsed().then_some(Value::Bool(true)),
    }
}

fn table_value(table: &xlsx::Table, key: TableKey) -> Option<Value<'_>> {
    let style = table.style_info.as_ref();
    let flag = |value: Option<bool>| value.map(Value::Bool);
    match key {
        TableKey::Name => Some(text(&table.name)),
        TableKey::DisplayName => Some(text(&table.display_name)),
        TableKey::Ref => Some(text(&table.ref_range)),
        TableKey::Style => style.and_then(|style| style.name.as_deref()).map(text),
        TableKey::BandedRows => flag(style.and_then(|style| style.show_row_stripes)),
        TableKey::BandedCols => flag(style.and_then(|style| style.show_column_stripes)),
        TableKey::FirstCol => flag(style.and_then(|style| style.show_first_column)),
        TableKey::LastCol => flag(style.and_then(|style| style.show_last_column)),
        TableKey::HeaderRow => Some(Value::Bool(table.header_row_count.unwrap_or(1) != 0)),
        TableKey::TotalRow => Some(Value::Bool(
            table.totals_row_count.unwrap_or(0) > 0 || table.totals_row_shown.unwrap_or(false),
        )),
        TableKey::Columns => Some(Value::Text(Cow::Owned(table.column_names().join(",")))),
    }
}

impl<'a> CellProps<'a> {
    fn get(&self, key: CellKey) -> Option<Value<'a>> {
        match key {
            CellKey::Merge => {
                return self
                    .merge
                    .map(|(range, _)| Value::Text(Cow::Owned(range_text(range))));
            },
            CellKey::MergeAnchor => {
                return self
                    .merge
                    .and_then(|(_, anchor)| anchor.then_some(Value::Bool(true)));
            },
            _ => {},
        }
        let cell = match self.state {
            CellState::Stored(cell) => cell,
            CellState::Stub => {
                // OfficeCLI's dense-range stub: `type=Number`, `empty=true`.
                return match key {
                    CellKey::Type => Some(Value::Token("Number")),
                    CellKey::Empty => Some(Value::Bool(true)),
                    _ => None,
                };
            },
            CellState::Missing | CellState::Covered(_) => return None,
        };
        let formula = match cell {
            xlsx::Cell::Formula(formula) => Some(formula),
            _ => None,
        };
        match key {
            CellKey::Type => Some(Value::Token(cell_type(cell))),
            CellKey::Formula => formula.map(|formula| text(formula.text())),
            CellKey::CachedValue => formula
                .and_then(xlsx::Formula::cached)
                .map(|cache| stored_text(cache.value()))
                .filter(|cached| !cached.is_empty())
                .map(text),
            CellKey::Evaluated => formula.map(|formula| Value::Bool(formula.cached().is_some())),
            CellKey::ArrayFormula => formula.and_then(|formula| {
                matches!(formula.kind(), xlsx::formula::Kind::Array { .. })
                    .then_some(Value::Bool(true))
            }),
            CellKey::ArrayRef => formula.and_then(|formula| match formula.kind() {
                xlsx::formula::Kind::Array { range: Some(range) } => Some(text(range.as_str())),
                _ => None,
            }),
            CellKey::Empty => {
                (formula.is_none() && display_text(cell).is_empty()).then_some(Value::Bool(true))
            },
            key => format_value(self.format?, key),
        }
    }
}

/// A key read from a cell's local formatting, with OfficeCLI's omissions: a
/// base format reports nothing, and the font, fill and border keys only when
/// the format names a record other than the workbook's default one.
fn format_value(format: &CellFormat, key: CellKey) -> Option<Value<'_>> {
    if format.is_base() {
        return None;
    }
    let font = (!format.has_default_font())
        .then(|| format.font())
        .flatten();
    let fill = (!format.has_default_fill())
        .then(|| format.fill())
        .flatten();
    let border = (!format.has_default_border())
        .then(|| format.border())
        .flatten();
    let alignment = format.alignment();
    match key {
        CellKey::Bold => font?.bold().map(Value::Bool),
        CellKey::Italic => font?.italic().map(Value::Bool),
        CellKey::Strike => font?.strike().map(Value::Bool),
        CellKey::Underline => match font?.underline()? {
            Underline::None => None,
            Underline::Double => Some(Value::Token("double")),
            Underline::SingleAccounting => Some(Value::Token("singleAccounting")),
            Underline::DoubleAccounting => Some(Value::Token("doubleAccounting")),
            _ => Some(Value::Token("single")),
        },
        CellKey::FontColor => font?.color().and_then(style_color),
        CellKey::Superscript => {
            matches!(font?.script()?, Script::Superscript).then_some(Value::Bool(true))
        },
        CellKey::Subscript => {
            matches!(font?.script()?, Script::Subscript).then_some(Value::Bool(true))
        },
        CellKey::FontSize => font?.size().map(|size| Value::Length(points(size))),
        CellKey::FontName => font?.name().map(text),
        CellKey::FontOutline => font?.outline().map(Value::Bool),
        CellKey::FontShadow => font?.shadow().map(Value::Bool),
        CellKey::FontCondense => font?.condense().map(Value::Bool),
        CellKey::FontExtend => font?.extend().map(Value::Bool),
        CellKey::FontCharset => font?.charset().map(|value| Value::Int(i64::from(value))),
        CellKey::FontFamily => font?.family().map(|value| Value::Int(i64::from(value))),
        CellKey::FontScheme => font?.scheme().map(|value| Value::Token(value.as_str())),
        CellKey::Fill => fill_value(fill?),
        CellKey::FillPattern => match fill?.pattern()?.pattern()? {
            pattern if is_solid(pattern) => None,
            pattern => Some(Value::Token(pattern.as_str())),
        },
        CellKey::FillBg => {
            let pattern = fill?.pattern()?;
            if is_solid(pattern.pattern()?) {
                return None;
            }
            pattern.background().and_then(style_color)
        },
        CellKey::BorderLeft => side_line(border?.left.as_ref()),
        CellKey::BorderLeftColor => side_color(border?.left.as_ref()),
        CellKey::BorderRight => side_line(border?.right.as_ref()),
        CellKey::BorderRightColor => side_color(border?.right.as_ref()),
        CellKey::BorderTop => side_line(border?.top.as_ref()),
        CellKey::BorderTopColor => side_color(border?.top.as_ref()),
        CellKey::BorderBottom => side_line(border?.bottom.as_ref()),
        CellKey::BorderBottomColor => side_color(border?.bottom.as_ref()),
        CellKey::BorderDiagonal => side_line(diagonal(border?)),
        CellKey::BorderDiagonalColor => side_color(diagonal(border?)),
        CellKey::BorderDiagonalUp => border?
            .diagonal
            .as_ref()?
            .dir()?
            .is_up()
            .then_some(Value::Bool(true)),
        CellKey::BorderDiagonalDown => border?
            .diagonal
            .as_ref()?
            .dir()?
            .is_down()
            .then_some(Value::Bool(true)),
        CellKey::WrapText => alignment?.wrap_text.then_some(Value::Bool(true)),
        CellKey::Horizontal => alignment?
            .horizontal
            .map(|value| Value::Token(Horizontal::as_str(value))),
        CellKey::Vertical => alignment?
            .vertical
            .map(|value| Value::Token(Vertical::as_str(value))),
        CellKey::TextRotation => alignment?
            .text_rotation
            .map(|rotation| i64::from(rotation.get()))
            .filter(|rotation| *rotation != 0)
            .map(Value::Int),
        CellKey::Indent => alignment?
            .indent
            .map(|indent| i64::from(indent.get()))
            .filter(|indent| *indent > 0)
            .map(Value::Int),
        CellKey::RelativeIndent => alignment?
            .relative_indent
            .map(|value| Value::Int(i64::from(value))),
        CellKey::JustifyLastLine => alignment?.justify_last_line.map(Value::Bool),
        CellKey::ShrinkToFit => alignment?.shrink_to_fit.then_some(Value::Bool(true)),
        CellKey::ReadingOrder => match alignment?.reading_order? {
            Reading::Context => None,
            Reading::LeftToRight => Some(Value::Token("ltr")),
            Reading::RightToLeft => Some(Value::Token("rtl")),
        },
        CellKey::Locked => format.protection()?.locked().map(Value::Bool),
        CellKey::FormulaHidden => format.protection()?.hidden().map(Value::Bool),
        CellKey::QuotePrefix => format.quote_prefix().then_some(Value::Bool(true)),
        CellKey::NumFmtId => {
            let id = xlsx::style::cell_format::raw::number_format_id(&format.number_format());
            (id != 0).then_some(Value::Int(i64::from(id)))
        },
        CellKey::NumberFormat => {
            let number_format = format.number_format();
            (!number_format.is_general())
                .then(|| number_format.code().map(text))
                .flatten()
        },
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

/// Solid and empty patterns keep OfficeCLI's single-colour `fill` readback.
fn is_solid(pattern: xlsx::style::cell_format::Pattern) -> bool {
    matches!(
        pattern,
        xlsx::style::cell_format::Pattern::Solid | xlsx::style::cell_format::Pattern::None
    )
}

/// OfficeCLI's `fill`: a gradient's stops as `gradient;#C1;#C2;deg`, else the
/// pattern's foreground colour.
fn fill_value(fill: &Fill) -> Option<Value<'static>> {
    if let Some(gradient) = fill.gradient() {
        let mut colors = gradient
            .stops()
            .iter()
            .filter_map(|stop| match stop.color() {
                StyleColor::Rgb { value, .. } => {
                    Some(Rgba::new(value.red(), value.green(), value.blue()))
                },
                _ => None,
            });
        let mut out = String::from("gradient");
        let mut count = 0;
        for color in colors.by_ref() {
            out.push(';');
            out.push_str(&color.to_string());
            count += 1;
        }
        if count < 2 {
            return None;
        }
        // Degrees are written as a whole number, as OfficeCLI's `(int)` does.
        #[allow(
            clippy::cast_possible_truncation,
            reason = "a gradient angle is a small finite number of degrees"
        )]
        let degree = gradient.degree() as i64;
        out.push(';');
        out.push_str(&degree.to_string());
        return Some(Value::Text(Cow::Owned(out)));
    }
    fill.pattern()?.foreground().and_then(style_color)
}

/// The diagonal border's side.
fn diagonal(border: &Border) -> Option<&Side> {
    border.diagonal.as_ref()?.side()
}

fn side_line(side: Option<&Side>) -> Option<Value<'static>> {
    side.map(|side| Value::Token(side.line.as_str()))
}

/// A border colour: OfficeCLI reads RGB border colours only.
fn side_color(side: Option<&Side>) -> Option<Value<'static>> {
    match side?.color? {
        StyleColor::Rgb { value, .. } => Some(Value::Color(rgb_color(value))),
        _ => None,
    }
}

/// A SpreadsheetML colour as OfficeCLI reads it back: RGB as `#RRGGBB` (or
/// `#RRGGBBAA` when not opaque), a theme slot by name with its tint. Indexed,
/// automatic and default colours have no readback.
fn style_color(color: StyleColor) -> Option<Value<'static>> {
    match color {
        StyleColor::Rgb { value, .. } => Some(Value::Color(rgb_color(value))),
        StyleColor::Theme { index, tint } => {
            let slot = Color::theme(ThemeSlot::from_index(index)?);
            let color = match tint {
                Some(tint) => slot.with_spreadsheet_tint(tint.get())?,
                None => slot,
            };
            Some(Value::Color(color))
        },
        _ => None,
    }
}

fn rgb_color(value: xlsx::Rgb) -> Color {
    Color::rgba(value.red(), value.green(), value.blue(), value.alpha())
}

/// A run font color as a value: an RGB or theme base (record 0338).
pub(super) fn font_color_value(
    color: &xlsx::style::stylesheet::font::FontColor,
) -> Option<Value<'static>> {
    let tint = color.tint().map(|tint| tint.get());
    match color.kind() {
        xlsx::style::FontColorKind::Rgb(value) => Some(Value::Color(rgb_color(value))),
        xlsx::style::FontColorKind::Theme(index) => {
            let slot = Color::theme(ThemeSlot::from_index(index)?);
            Some(Value::Color(match tint {
                Some(tint) => slot.with_spreadsheet_tint(tint)?,
                None => slot,
            }))
        },
        _ => None,
    }
}

/// A length of `points`, rounded to the nearest EMU.
pub(super) fn points(points: f64) -> Length {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "font sizes and row heights are a few hundred points at most"
    )]
    Length::from_emu((points * 12_700.0).round() as i64)
}

/// `A1:B2`, the text of a merge.
fn range_text(range: Rect) -> String {
    let first = (range.start().column().get(), range.start().row().get());
    let last = (
        range.end().1.saturating_sub(1),
        range.end().0.saturating_sub(1),
    );
    Label::range(first, last).as_str().to_owned()
}

/// Detach a value from the node it was read from.
fn owned(value: Value<'_>) -> Value<'static> {
    match value {
        Value::Text(text) => Value::Text(Cow::Owned(text.into_owned())),
        Value::Bool(value) => Value::Bool(value),
        Value::Int(value) => Value::Int(value),
        Value::Number(value) => Value::Number(value),
        Value::Token(value) => Value::Token(value),
        Value::Length(value) => Value::Length(value),
        Value::Color(value) => Value::Color(value),
        Value::List(items) => Value::List(Cow::Owned(items.into_owned())),
    }
}
