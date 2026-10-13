//! The parsed grid of a source-backed worksheet and the reads it serves.
//!
//! A [`SourceWorksheet`] builds its sheet's parsed store on the first read
//! that needs it, in one pass over the worksheet part, and offers it to the
//! workbook's weighted clean-value cache. The handle pins the store it read,
//! so its later reads are lookups in that store: O(log n) for a cell or a
//! row, O(1) for the extents, and O(k) for k visited cells. Another handle
//! of the same workbook finds the store in the cache while it is retained.

use std::fmt;
use std::ops::ControlFlow;
use std::sync::Arc;

use litchi_sheet::{Area, At, Cell as Address, ColumnAt, RowAt};

use super::{Selection, SourceWorksheet, resolve_selected_record};
use crate::cell::{Cell, Cells, Extents, RecordIndex, Store, View};
use crate::error::{Result, invalid};
use crate::layout::Defaults;
use crate::leniency::ToleranceReport;
use crate::merge::Merges;
use crate::{Column, Columns, Row, Rows};

/// A pinned, immutable parsed grid of one source-backed worksheet.
///
/// The grid holds the worksheet's stored cells, row and column records,
/// merged ranges, grid defaults, and cell extents exactly as
/// [`crate::Worksheet`] reads them from the same part. It is returned by
/// [`SourceWorksheet::grid`] and is cheap to clone: clones share one parsed
/// store. While any clone is alive the store stays resident, whatever the
/// workbook's cache evicts; dropping the last clone and the last handle that
/// read it leaves the store to the cache's bounded retention.
///
/// The grid is a snapshot of the source version the workbook captured. It
/// does not read the source again, so its methods neither check the source
/// version nor observe cancellation, and they may outlive the workbook.
#[derive(Clone)]
pub struct Grid {
    store: Arc<Store>,
}

// Grids are shared by `Arc`-backed, `Send + Sync` node values of the unified
// operations layer.
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Grid>();
};

impl Grid {
    pub(super) const fn new(store: Arc<Store>) -> Self {
        Self { store }
    }

    pub(super) fn store(&self) -> &Store {
        &self.store
    }

    /// Look up one exact logical cell state by A1 reference or zero-based
    /// coordinate, in O(log n).
    ///
    /// [`View::Missing`], [`View::Covered`] and a stored cell stay distinct,
    /// as for [`crate::Worksheet::cell`].
    pub fn cell<'a>(&self, at: impl Into<At<'a>>) -> Result<View<'_>> {
        let address = at.into().resolve()?;
        Ok(self.store.view(address))
    }

    /// Lazily traverse the stored cells of a checked range in row-major
    /// order.
    pub fn cells<'a>(&self, area: impl Into<Area<'a>>) -> Result<Cells<'_>> {
        let range = area.into().resolve()?;
        Ok(self.store.cells(range))
    }

    /// The stored record at a position [`Cells::indexed`] reported for this
    /// grid, with its coordinate, in O(1); `None` past the last record.
    #[must_use]
    pub fn record(&self, index: RecordIndex) -> Option<(Address, &Cell)> {
        self.store.record(index)
    }

    /// Number of stored cell records.
    #[must_use]
    pub fn stored_cell_count(&self) -> usize {
        self.store.stored_cell_count()
    }

    /// Lazily traverse the validated, non-overlapping merged ranges.
    #[must_use]
    pub fn merges(&self) -> Merges<'_> {
        self.store.merges()
    }

    /// Borrow one checked logical row, including an implicit default row.
    pub fn row(&self, at: impl Into<RowAt>) -> Result<Row<'_>> {
        let index = at.into().resolve()?;
        Ok(self.store.row(index))
    }

    /// Lazily traverse only the explicit worksheet row records.
    #[must_use]
    pub fn rows(&self) -> Rows<'_> {
        self.store.rows()
    }

    /// Borrow one checked logical column, including an implicit default
    /// column.
    pub fn column<'a>(&self, at: impl Into<ColumnAt<'a>>) -> Result<Column<'_>> {
        let index = at.into().resolve()?;
        Ok(self.store.column(index))
    }

    /// Lazily traverse the logical columns covered by explicit property
    /// records.
    #[must_use]
    pub fn columns(&self) -> Columns<'_> {
        self.store.columns()
    }

    /// Stored worksheet-grid defaults, if the producer supplied them.
    #[must_use]
    pub fn defaults(&self) -> Option<&Defaults> {
        self.store.defaults()
    }

    /// Distinct declared, stored, content, and directly styled cell bounds.
    #[must_use]
    pub fn extents(&self) -> &Extents {
        self.store.extents()
    }

    /// The repairs made while this grid was parsed.
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        self.store.tolerance_report()
    }

    /// Logical retained weight of the parsed store, in bytes.
    ///
    /// This is the weight the workbook's semantic cache charges for the
    /// store: its record arrays plus the payload text the records own. It is
    /// not an allocator or resident-memory measurement.
    #[must_use]
    pub fn retained_weight(&self) -> usize {
        self.store.weight_bytes()
    }

    /// Whether two grids share one parsed store.
    #[cfg(test)]
    pub(super) fn same_grid(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.store, &other.store)
    }
}

impl fmt::Debug for Grid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Grid")
            .field("stored_cells", &self.store.stored_cell_count())
            .field("extents", self.store.extents())
            .finish_non_exhaustive()
    }
}

impl SourceWorksheet {
    /// Pin and return this worksheet's parsed grid.
    ///
    /// The first call on a worksheet whose grid is not retained parses its
    /// part once; the grid is then pinned by this handle and offered to the
    /// workbook's semantic cache, so later reads through this handle, or
    /// through another handle while the cache retains it, do not scan the
    /// part again. Unlike the cell reads, this always materializes the grid,
    /// whatever the size of the part.
    ///
    /// ```no_run
    /// use litchi_xlsx::{SourceBackedWorkbook, workbook::source::Grid};
    /// let workbook = SourceBackedWorkbook::open("report.xlsx")?;
    /// let sheet = workbook.sheet("Summary")?.expect("worksheet");
    /// let grid: Grid = sheet.grid()?;
    /// let first = grid.cell("A1")?;
    /// assert_eq!(grid.clone().stored_cell_count(), grid.stored_cell_count());
    /// # let _ = first;
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// [`crate::Error::NotWorksheet`] for a chart, dialog, or other non-grid
    /// sheet, a typed parse error for a malformed worksheet, and the source,
    /// cancellation, and limit errors of every source-backed read.
    pub fn grid(&self) -> Result<Grid> {
        let grid = self.pinned_grid();
        self.finish_result(grid)
    }

    /// Lazily traverse the validated, non-overlapping merged ranges.
    pub fn merges(&self) -> Result<Merges<'_>> {
        self.finish_result(self.store().map(Store::merges))
    }

    /// Borrow one checked logical row, including an implicit default row.
    pub fn row(&self, at: impl Into<RowAt>) -> Result<Row<'_>> {
        let row = match at.into().resolve() {
            Ok(index) => self.store().map(|store| store.row(index)),
            Err(error) => Err(error.into()),
        };
        self.finish_result(row)
    }

    /// Lazily traverse only the explicit worksheet row records.
    pub fn rows(&self) -> Result<Rows<'_>> {
        self.finish_result(self.store().map(Store::rows))
    }

    /// Borrow one checked logical column, including an implicit default
    /// column. A1 labels such as `"B"` are the primary entry; raw inputs are
    /// zero-based and validated before lookup.
    pub fn column<'a>(&self, at: impl Into<ColumnAt<'a>>) -> Result<Column<'_>> {
        let column = match at.into().resolve() {
            Ok(index) => self.store().map(|store| store.column(index)),
            Err(error) => Err(error.into()),
        };
        self.finish_result(column)
    }

    /// Lazily traverse logical columns covered by explicit property records.
    pub fn columns(&self) -> Result<Columns<'_>> {
        self.finish_result(self.store().map(Store::columns))
    }

    /// Stored worksheet-grid defaults, if the producer supplied them.
    pub fn defaults(&self) -> Result<Option<&Defaults>> {
        self.finish_result(self.store().map(Store::defaults))
    }

    /// Distinct declared, stored, content, and directly styled cell bounds.
    pub fn extents(&self) -> Result<&Extents> {
        self.finish_result(self.store().map(Store::extents))
    }

    /// Visit the stored cells of a checked range until the callback breaks.
    ///
    /// Cells are produced in row-major order from the same route and with
    /// the same values as [`Self::visit_cells`]. The callback returns
    /// [`ControlFlow::Break`] to stop the walk; the result is then
    /// `Ok(ControlFlow::Break(value))`. A walk that visits every selected
    /// cell returns `Ok(ControlFlow::Continue(count))` with the number of
    /// cells visited.
    ///
    /// On a worksheet whose parsed grid is resident (pinned by this handle or
    /// retained by the workbook), or that this read materializes, the walk
    /// stops after the breaking cell: its cost is that of the cells it
    /// visits. On a worksheet read through the bounded streaming scan (a
    /// cold eligible worksheet, see
    /// [`Self::cell`]), the scan still reaches the end of the worksheet and
    /// raises every refusal before the first callback, as for
    /// [`Self::visit_cells`]; breaking early then saves only the callbacks.
    ///
    /// Cancellation is checked before each callback. The final source and
    /// execution fences run after a break as well, so a source change or a
    /// cancellation outranks the break value.
    ///
    /// ```no_run
    /// use std::ops::ControlFlow;
    /// let workbook = litchi_xlsx::SourceBackedWorkbook::open("report.xlsx")?;
    /// let sheet = workbook.sheet("Summary")?.expect("worksheet");
    /// let first = sheet.visit_cells_until("A1:C9", |address, _| {
    ///     ControlFlow::<_, ()>::Break(address)
    /// })?;
    /// # let _ = first;
    /// # Ok::<(), litchi_xlsx::Error>(())
    /// ```
    pub fn visit_cells_until<'a, B, F>(
        &self,
        area: impl Into<Area<'a>>,
        mut visit: F,
    ) -> Result<ControlFlow<B, usize>>
    where
        F: FnMut(Address, &Cell) -> ControlFlow<B>,
    {
        let range = match area.into().resolve() {
            Ok(range) => range,
            Err(error) => return self.finish_result(Err(error.into())),
        };
        // Publish the selection through the same fences a whole-range read
        // runs before any caller can observe a cell.
        let selection = match self.finish_result(self.select_cells(range)) {
            Ok(selection) => selection,
            Err(error) => return self.finish_result(Err(error)),
        };
        let mut walk = Walk {
            visited: 0,
            outcome: Ok(ControlFlow::Continue(())),
        };
        match selection {
            Selection::Stored(store) => {
                for (address, cell) in store.cells(range) {
                    if !self.visit_one(&mut visit, address, cell, &mut walk) {
                        break;
                    }
                }
            },
            Selection::Selected { cells, shared_text } => {
                for record in cells {
                    let source_cell = match resolve_selected_record(record, shared_text.as_deref())
                    {
                        Ok(source_cell) => source_cell,
                        Err(error) => {
                            walk.outcome = Err(error);
                            break;
                        },
                    };
                    if !self.visit_one(
                        &mut visit,
                        source_cell.address,
                        &source_cell.cell,
                        &mut walk,
                    ) {
                        break;
                    }
                }
            },
        }
        let Walk { visited, outcome } = walk;
        self.finish_result(outcome.map(|flow| match flow {
            ControlFlow::Continue(()) => ControlFlow::Continue(visited),
            ControlFlow::Break(value) => ControlFlow::Break(value),
        }))
    }

    /// Hand one produced cell to a range visitor.
    ///
    /// Cancellation is checked before the callback and the visited count is
    /// bounded. `false` stops the walk, with `walk.outcome` carrying the
    /// break value or the primary error.
    fn visit_one<B, F>(
        &self,
        visit: &mut F,
        address: Address,
        cell: &Cell,
        walk: &mut Walk<B>,
    ) -> bool
    where
        F: FnMut(Address, &Cell) -> ControlFlow<B>,
    {
        if let Err(error) = self.owner.execution_check() {
            walk.outcome = Err(error);
            return false;
        }
        if let ControlFlow::Break(value) = visit(address, cell) {
            walk.outcome = Ok(ControlFlow::Break(value));
            return false;
        }
        match walk.visited.checked_add(1) {
            Some(count) => {
                walk.visited = count;
                true
            },
            None => {
                walk.outcome = Err(invalid("source-backed cell visit count overflow"));
                false
            },
        }
    }

    /// The parsed grid this handle pins, materialized on first use.
    fn pinned_grid(&self) -> Result<Grid> {
        self.store()?;
        self.grid
            .get()
            .cloned()
            .ok_or_else(|| invalid("source-backed worksheet grid was not pinned"))
    }
}

/// Progress of one range visit.
struct Walk<B> {
    visited: usize,
    outcome: Result<ControlFlow<B>>,
}
