//! Shared borrowed reader for selected iWork table-cell storage.
//!
//! Format adapters retain table selection, archive lookup, semantic sidecar
//! maps, and their aggregate budget implementation. This module owns the
//! common tile/row topology checks, sparse offset validation, and source-backed
//! cell classification. A sink must stage its side effects and publish them
//! only after [`read_table_cells`] returns successfully.

use std::collections::HashSet;

use litchi_iwa_protos::numbers_table_cell_storage_codec as storage;

use crate::cell_value::{self, CellValueSource};
pub use crate::table_data_list::Message;

/// Native message type for one table tile.
pub const TILE_MESSAGE_KIND: u32 = 6_002;

/// Dimensions and tile geometry already selected by a format adapter.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableDimensions {
    rows: u32,
    columns: u32,
    tile_size: u32,
}

impl TableDimensions {
    /// Construct a selected table geometry.
    #[must_use]
    pub const fn new(rows: u32, columns: u32, tile_size: u32) -> Self {
        Self {
            rows,
            columns,
            tile_size,
        }
    }

    /// Number of addressable rows.
    #[must_use]
    pub const fn rows(self) -> u32 {
        self.rows
    }

    /// Number of addressable columns.
    #[must_use]
    pub const fn columns(self) -> u32 {
        self.columns
    }

    /// Number of rows covered by one tile.
    #[must_use]
    pub const fn tile_size(self) -> u32 {
        self.tile_size
    }
}

/// One selected tile coordinate and its package-local object reference.
///
/// The object reference is consumed only by the resolver passed to
/// [`read_table_cells`]. It never appears in a returned cell or read report.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TileReference {
    tile_index: u32,
    object_id: u64,
}

impl TileReference {
    /// Construct a tile reference from its selected coordinate and resolver
    /// key.
    #[must_use]
    pub const fn new(tile_index: u32, object_id: u64) -> Self {
        Self {
            tile_index,
            object_id,
        }
    }

    /// Return the zero-based tile coordinate.
    #[must_use]
    pub const fn tile_index(self) -> u32 {
        self.tile_index
    }

    /// Return the resolver key for this tile.
    #[must_use]
    pub const fn object_id(self) -> u64 {
        self.object_id
    }
}

/// A source-backed semantic cell delivered to an adapter sink.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct CellSource<'source> {
    row: u32,
    column: u32,
    bytes: &'source [u8],
    value: CellValueSource,
}

impl<'source> CellSource<'source> {
    /// Return the zero-based table row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Return the zero-based table column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Return the exact borrowed cell payload.
    #[must_use]
    pub const fn bytes(self) -> &'source [u8] {
        self.bytes
    }

    /// Return the source-backed value classification and sidecar references.
    #[must_use]
    pub const fn value(self) -> CellValueSource {
        self.value
    }
}

/// Resource or shape failure observed while walking selected cell storage.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum TableCellIssue {
    /// The selected tile object could not be resolved.
    MissingTile { object_id: u64 },
    /// The selected tile object has no type-6002 payload.
    MissingTilePayload { object_id: u64 },
    /// The selected tile object has multiple type-6002 payloads.
    DuplicateTilePayload { object_id: u64 },
    /// A tile coordinate lies outside the selected table's tile grid.
    TileIndexOutOfBounds { tile_index: u32, tile_count: u32 },
    /// A tile row coordinate is outside the tile's local geometry.
    TileRowOutOfBounds {
        tile_index: u32,
        row_index: u32,
        tile_size: u32,
    },
    /// A tile row coordinate cannot be translated to a table row.
    RowCoordinateOverflow { tile_index: u32, row_index: u32 },
    /// A translated row lies outside the selected table.
    TableRowOutOfBounds { row: u32, rows: u32 },
    /// A tile repeats one local row coordinate.
    DuplicateTileRow { tile_index: u32, row_index: u32 },
    /// A counter could not be represented by the host platform.
    CounterOverflow,
    /// A sparse cell span does not borrow from the active row storage.
    CellStorageOutOfBounds { row: u32, column: u32 },
    /// The shared storage codec rejected a tile or row payload.
    StorageDecode(storage::DecodeError),
    /// The shared value classifier rejected one cell payload.
    CellValueDecode {
        row: u32,
        column: u32,
        error: cell_value::DecodeError,
    },
    /// A bounded staging allocation was refused.
    Allocation {
        /// Collection that could not grow.
        target: AllocationTarget,
        /// Requested additional elements.
        amount: usize,
    },
}

/// Collection whose staging allocation is charged by the adapter budget.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AllocationTarget {
    /// Per-tile row identity set.
    TileRows,
}

/// Aggregate counts from one successful selected-table read.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct TableCellReadReport {
    tiles: usize,
    rows: usize,
    cells: usize,
}

impl TableCellReadReport {
    /// Number of successfully decoded tile payloads.
    #[must_use]
    pub const fn tiles(self) -> usize {
        self.tiles
    }

    /// Number of validated sparse tile rows.
    #[must_use]
    pub const fn rows(self) -> usize {
        self.rows
    }

    /// Number of validated present cells.
    #[must_use]
    pub const fn cells(self) -> usize {
        self.cells
    }
}

/// Aggregate budget hooks owned by a concrete format adapter.
pub trait TableCellReadBudget {
    /// Adapter error returned for wire, classification, budget, or sink
    /// failures.
    type Error;

    /// Build the strict storage-codec options for one borrowed payload from
    /// the adapter's current residual budget.
    fn storage_options(&mut self, source: &[u8]) -> Result<storage::DecodeOptions, Self::Error>;

    /// Atomically charge one completed tile or row-span storage report.
    fn charge_storage_report(&mut self, report: storage::DecodeReport) -> Result<(), Self::Error>;

    /// Admit the cumulative number of present cells before cell values are
    /// classified or handed to the sink.
    fn check_materialized_cells(&mut self, observed: usize) -> Result<(), Self::Error>;

    /// Charge the exact borrowed cell payload before parsing it.
    fn charge_cell_source(&mut self, bytes: usize) -> Result<(), Self::Error>;

    /// Charge an auxiliary fallible staging allocation before it is attempted.
    fn charge_allocation(
        &mut self,
        _target: AllocationTarget,
        _amount: usize,
    ) -> Result<(), Self::Error> {
        Ok(())
    }

    /// Map a shared issue into the adapter's format-specific error type.
    fn map_issue(&mut self, issue: TableCellIssue) -> Self::Error;
}

/// Receives source-backed values and resolves package-owned sidecars.
pub trait CellValueSink<Budget: TableCellReadBudget> {
    /// Stage one classified cell using the reader's aggregate budget ledger.
    /// The sink must not publish irreversible mutations before the enclosing
    /// table read succeeds.
    fn visit_cell(
        &mut self,
        cell: CellSource<'_>,
        budget: &mut Budget,
    ) -> Result<(), Budget::Error>;
}

/// Read all cells from already-selected tile references.
///
/// The resolver supplies only borrowed type-6002 messages for one object. The
/// engine performs one strict tile pass and one borrowed row-span pass per
/// selected payload, then classifies every present cell with
/// [`crate::cell_value::decode_cell_value`]. It never retains native objects,
/// tile payloads, row offsets, or cell values. Duplicate tile and object
/// references must be rejected by the owner's selected topology before this
/// function is called. Both the resolver and sink receive the same mutable
/// budget instance, so package lookup and sidecar staging can consume one
/// aggregate ledger without interior mutability.
pub fn read_table_cells<'source, Tiles, Messages, Resolve, Budget, Sink>(
    dimensions: TableDimensions,
    tile_references: Tiles,
    mut resolve_tile: Resolve,
    budget: &mut Budget,
    sink: &mut Sink,
) -> Result<TableCellReadReport, Budget::Error>
where
    Tiles: IntoIterator<Item = TileReference>,
    Messages: IntoIterator<Item = Message<'source>>,
    Resolve: FnMut(u64, &mut Budget) -> Result<Option<Messages>, Budget::Error>,
    Budget: TableCellReadBudget,
    Sink: CellValueSink<Budget>,
{
    let tile_count = if dimensions.tile_size == 0 {
        return Err(budget.map_issue(TableCellIssue::CounterOverflow));
    } else {
        dimensions.rows.div_ceil(dimensions.tile_size)
    };
    let columns = usize::try_from(dimensions.columns)
        .map_err(|_| budget.map_issue(TableCellIssue::CounterOverflow))?;
    let mut report = TableCellReadReport::default();

    for tile_reference in tile_references {
        if tile_reference.tile_index >= tile_count {
            return Err(budget.map_issue(TableCellIssue::TileIndexOutOfBounds {
                tile_index: tile_reference.tile_index,
                tile_count,
            }));
        }
        let Some(messages) = resolve_tile(tile_reference.object_id, budget)? else {
            return Err(budget.map_issue(TableCellIssue::MissingTile {
                object_id: tile_reference.object_id,
            }));
        };
        let mut payload = None;
        let mut duplicate_payload = false;
        for message in messages {
            if message.kind != TILE_MESSAGE_KIND {
                continue;
            }
            if payload.replace(message.data).is_some() {
                duplicate_payload = true;
            }
        }
        if duplicate_payload {
            return Err(budget.map_issue(TableCellIssue::DuplicateTilePayload {
                object_id: tile_reference.object_id,
            }));
        }
        let Some(payload) = payload else {
            return Err(budget.map_issue(TableCellIssue::MissingTilePayload {
                object_id: tile_reference.object_id,
            }));
        };

        let options = budget.storage_options(payload)?;
        let mut visitor = TileVisitor::new(
            tile_reference.tile_index,
            dimensions,
            columns,
            report.cells,
            budget,
            sink,
        );
        let decoded = storage::decode_tile_with_visitor(payload, options, &mut visitor);
        let (_, storage_report) = match decoded {
            Ok(decoded) => decoded,
            Err(error) => {
                return Err(visitor
                    .budget
                    .map_issue(TableCellIssue::StorageDecode(error)));
            },
        };
        visitor.budget.charge_storage_report(storage_report)?;
        let (tile_report, structural_error, semantic_error) = visitor.into_parts();
        report.tiles = report
            .tiles
            .checked_add(1)
            .ok_or_else(|| budget.map_issue(TableCellIssue::CounterOverflow))?;
        report.rows = report
            .rows
            .checked_add(tile_report.rows)
            .ok_or_else(|| budget.map_issue(TableCellIssue::CounterOverflow))?;
        report.cells = report
            .cells
            .checked_add(tile_report.cells)
            .ok_or_else(|| budget.map_issue(TableCellIssue::CounterOverflow))?;
        if let Some(error) = structural_error {
            return Err(error);
        }
        if let Some(error) = semantic_error {
            return Err(error);
        }
    }

    Ok(report)
}

struct TileVisitor<'budget, 'sink, Budget, Sink>
where
    Budget: TableCellReadBudget,
{
    tile_index: u32,
    dimensions: TableDimensions,
    columns: usize,
    previous_cells: usize,
    budget: &'budget mut Budget,
    sink: &'sink mut Sink,
    seen_rows: HashSet<u32>,
    report: TableCellReadReport,
    structural_error: Option<Budget::Error>,
    semantic_error: Option<Budget::Error>,
}

impl<'budget, 'sink, Budget, Sink> TileVisitor<'budget, 'sink, Budget, Sink>
where
    Budget: TableCellReadBudget,
    Sink: CellValueSink<Budget>,
{
    fn new(
        tile_index: u32,
        dimensions: TableDimensions,
        columns: usize,
        previous_cells: usize,
        budget: &'budget mut Budget,
        sink: &'sink mut Sink,
    ) -> Self {
        Self {
            tile_index,
            dimensions,
            columns,
            previous_cells,
            budget,
            sink,
            seen_rows: HashSet::new(),
            report: TableCellReadReport::default(),
            structural_error: None,
            semantic_error: None,
        }
    }

    fn record_structural(&mut self, issue: TableCellIssue) {
        if self.structural_error.is_none() {
            self.structural_error = Some(self.budget.map_issue(issue));
        }
    }

    fn record_semantic(&mut self, issue: TableCellIssue) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(self.budget.map_issue(issue));
        }
    }

    fn record_budget_error(&mut self, error: Budget::Error) {
        if self.semantic_error.is_none() {
            self.semantic_error = Some(error);
        }
    }

    fn cumulative_cells(&self, tile_cells: usize) -> Option<usize> {
        self.previous_cells.checked_add(tile_cells)
    }

    fn into_parts(
        self,
    ) -> (
        TableCellReadReport,
        Option<Budget::Error>,
        Option<Budget::Error>,
    ) {
        (self.report, self.structural_error, self.semantic_error)
    }
}

impl<Budget, Sink> storage::StorageVisitor for TileVisitor<'_, '_, Budget, Sink>
where
    Budget: TableCellReadBudget,
    Sink: CellValueSink<Budget>,
{
    fn visit_tile_row(
        &mut self,
        row: storage::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage::DecodeError> {
        let row_index = row.tile_row_index();
        if row_index >= self.dimensions.tile_size {
            self.record_structural(TableCellIssue::TileRowOutOfBounds {
                tile_index: self.tile_index,
                row_index,
                tile_size: self.dimensions.tile_size,
            });
            return Ok(());
        }
        let Some(global_row) = self
            .tile_index
            .checked_mul(self.dimensions.tile_size)
            .and_then(|origin| origin.checked_add(row_index))
        else {
            self.record_structural(TableCellIssue::RowCoordinateOverflow {
                tile_index: self.tile_index,
                row_index,
            });
            return Ok(());
        };
        if global_row >= self.dimensions.rows {
            self.record_structural(TableCellIssue::TableRowOutOfBounds {
                row: global_row,
                rows: self.dimensions.rows,
            });
            return Ok(());
        }
        if self.seen_rows.contains(&row_index) {
            self.record_structural(TableCellIssue::DuplicateTileRow {
                tile_index: self.tile_index,
                row_index,
            });
            return Ok(());
        }
        let declared_cells = match usize::try_from(row.cell_count()) {
            Ok(value) => value,
            Err(_conversion) => {
                self.record_structural(TableCellIssue::CounterOverflow);
                return Ok(());
            },
        };

        // Reserve the row identity before charging the materialized-cell
        // count.  This preserves the legacy admission order: a refused
        // staging allocation never consumes a cell-count charge, while the
        // declared row count is still checked before offset validation.
        if let Err(error) = self.budget.charge_allocation(AllocationTarget::TileRows, 1) {
            self.record_budget_error(error);
            return Ok(());
        }
        if self.seen_rows.try_reserve(1).is_err() {
            self.record_semantic(TableCellIssue::Allocation {
                target: AllocationTarget::TileRows,
                amount: self.seen_rows.len().saturating_add(1),
            });
            return Ok(());
        }
        self.seen_rows.insert(row_index);

        let allow_semantic_work = self.semantic_error.is_none();
        if allow_semantic_work {
            let Some(declared_tile_cells) = self.report.cells.checked_add(declared_cells) else {
                self.record_structural(TableCellIssue::CounterOverflow);
                return Ok(());
            };
            let Some(declared_cells) = self.cumulative_cells(declared_tile_cells) else {
                self.record_structural(TableCellIssue::CounterOverflow);
                return Ok(());
            };
            if let Err(error) = self.budget.check_materialized_cells(declared_cells) {
                self.record_budget_error(error);
                return Ok(());
            }
        }

        let (_, offsets) = row.cell_storage_and_offsets();
        let options = match self.budget.storage_options(offsets) {
            Ok(options) => options,
            Err(error) => {
                self.record_budget_error(error);
                return Ok(());
            },
        };
        let (spans, span_report) = match row.cell_spans(self.columns, options) {
            Ok(value) => value,
            Err(error) => {
                self.record_structural(TableCellIssue::StorageDecode(error));
                return Ok(());
            },
        };
        if let Err(error) = self.budget.charge_storage_report(span_report) {
            self.record_budget_error(error);
            return Ok(());
        }
        let Some(next_cells) = self.report.cells.checked_add(spans.len()) else {
            self.record_structural(TableCellIssue::CounterOverflow);
            return Ok(());
        };
        if allow_semantic_work {
            let Some(cumulative_cells) = self.cumulative_cells(next_cells) else {
                self.record_structural(TableCellIssue::CounterOverflow);
                return Ok(());
            };
            if let Err(error) = self.budget.check_materialized_cells(cumulative_cells) {
                self.record_budget_error(error);
                return Ok(());
            }
        }
        let Some(next_rows) = self.report.rows.checked_add(1) else {
            self.record_structural(TableCellIssue::CounterOverflow);
            return Ok(());
        };
        self.report.rows = next_rows;
        self.report.cells = next_cells;
        let (storage_bytes, _) = row.cell_storage_and_offsets();
        for span in spans.iter() {
            let column = span.column();
            let Ok(column_u32) = u32::try_from(column) else {
                self.record_structural(TableCellIssue::CounterOverflow);
                continue;
            };
            let Some(cell_bytes) = span.bytes(storage_bytes) else {
                self.record_structural(TableCellIssue::CellStorageOutOfBounds {
                    row: global_row,
                    column: column_u32,
                });
                continue;
            };
            if !allow_semantic_work || self.semantic_error.is_some() {
                continue;
            }
            if let Err(error) = self.budget.charge_cell_source(cell_bytes.len()) {
                self.record_budget_error(error);
                continue;
            }
            let value = match cell_value::decode_cell_value(cell_bytes) {
                Ok(value) => value,
                Err(error) => {
                    self.record_semantic(TableCellIssue::CellValueDecode {
                        row: global_row,
                        column: column_u32,
                        error,
                    });
                    continue;
                },
            };
            if let Err(error) = self.sink.visit_cell(
                CellSource {
                    row: global_row,
                    column: column_u32,
                    bytes: cell_bytes,
                    value,
                },
                self.budget,
            ) {
                self.record_budget_error(error);
            }
        }
        Ok(())
    }
}
