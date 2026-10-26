//! The dependent formula-cache refresh of value writes into a table that
//! takes part in a formula graph, for the table owners that stage payloads
//! by object identifier (Keynote slide tables, Pages body tables; record
//! 0376).
//!
//! [`plan`](crate::formula_cache::refresh::plan) reads the document's calculation engine and formula owners, the
//! edited table's storage and every formula-owning table's values through a
//! caller-supplied object resolver, and runs the shared planner
//! ([`plan_final_cache`](crate::formula_cache::plan_final_cache)) with the strict scalar
//! evaluator over the writes as the final-state overlay. It publishes
//! nothing: the result names, per tile, the formula cells whose cached
//! results change, and [`TileCaches::apply`](crate::formula_cache::refresh::TileCaches::apply) rewrites a (staged) tile
//! payload. The policy is the Numbers table owner's, except that range
//! precedents are planned (the Numbers owner refuses a selected owner with
//! range-precedent tiles): a write that replaces or clears a formula cell, a
//! segmented formula list or a formula the evaluator does not support is
//! refused with [`Error::UnsupportedDependency`](crate::formula_cache::Error::UnsupportedDependency).

use litchi_iwa_common::formula::FormulaCachedValue;

use super::{
    CacheBaseline, CacheLimits, CacheSource, Coordinate, DependencyPayload, Error, FinalCell,
    FinalValue, FormulaCell, FormulaPayload, ScalarValue, StrictEvaluator, TableGeometry,
    TableIdentity, Unsupported, collect_formula_list, decode_options, map_decode_failure,
    plan_final_cache, validate_formula_cell_payloads,
};
use crate::{
    BncCell, BncCellView, CachedScalar, Limits, StoredValue,
    table_write::{self, CellInput, WriteError},
};
use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency;

const TABLE_OWNER_KIND: u32 = 1;
const DEFAULT_TILE_ROWS: u64 = 256;
const FORMULA_NESTING: u32 = 64;

/// The calculation graph of the document and the edited table.
#[derive(Debug, Clone, Copy)]
pub struct Source<'a> {
    /// The `TSCE.CalculationEngineArchive` payload.
    pub engine: &'a [u8],
    /// Every `TSCE.FormulaOwnerDependenciesArchive` payload of the engine,
    /// with its object identifier.
    pub owners: &'a [DependencyPayload<'a>],
    /// The edited table's drawable (`TST.TableInfoArchive`) identifier.
    pub table: u64,
}

/// One value write of the batch, at a table coordinate.
#[derive(Debug, Clone, Copy)]
pub struct Write<'a> {
    /// Table row.
    pub row: u32,
    /// Table column.
    pub column: u32,
    /// The written value.
    pub value: &'a CellInput<'a>,
}

/// The formula caches one tile's rewrite changes.
#[derive(Debug, Clone, PartialEq)]
pub struct TileCaches {
    /// The tile's object identifier.
    pub tile: u64,
    cells: Vec<(u32, u32, FormulaCachedValue)>,
}

/// The refresh a batch of writes needs.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Plan {
    /// Cache rewrites, by tile, in tile-identifier order.
    pub tiles: Vec<TileCaches>,
    /// Formula hosts whose caches the plan refreshes.
    pub refreshed: usize,
}

impl Plan {
    /// Whether the writes need no cache rewrite.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tiles.is_empty()
    }
}

impl TileCaches {
    /// Rewrite the cached results in `tile` (the tile's payload after the
    /// batch's value writes) of a table with `columns` columns.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidSource`] when a target cell is not a formula cell;
    /// [`Error::UnsupportedDependency`] for a result kind the BNC writer
    /// does not store (only number and Boolean results are written, as the
    /// host does).
    pub fn apply(&self, tile: &[u8], columns: u32, limits: &Limits) -> Result<Vec<u8>, Error> {
        let mut tile = copy(tile)?;
        for (row, column, value) in &self.cells {
            let previous = table_write::tile_cell(&tile, *row, *column, limits)
                .map_err(map_write)?
                .ok_or(Error::InvalidSource)?;
            let mut cell = BncCell::parse(&previous, limits).map_err(|_| Error::InvalidSource)?;
            if !matches!(cell.stored_value(), StoredValue::Formula(_)) {
                return Err(Error::InvalidSource);
            }
            match value {
                FormulaCachedValue::Number(number) => cell
                    .set_formula_cached_number(number.get())
                    .map_err(|_| Error::InvalidSource)?,
                FormulaCachedValue::Boolean(flag) => cell
                    .set_formula_cached_boolean(*flag)
                    .map_err(|_| Error::InvalidSource)?,
                _ => return Err(Error::UnsupportedDependency(Unsupported::CacheType)),
            }
            let encoded = cell.encode(limits).map_err(|_| Error::InvalidSource)?;
            tile =
                table_write::tile_set_cell(&tile, *row, *column, columns, Some(&encoded), limits)
                    .map_err(map_write)?
                    .tile;
        }
        Ok(tile)
    }
}

/// The planner limits derived from the owner's raw-wire [`Limits`].
#[must_use]
pub fn cache_limits(limits: &Limits) -> CacheLimits {
    let items = limits.max_items();
    CacheLimits {
        wire_bytes: limits.max_input_bytes(),
        wire_fields: limits.max_fields(),
        wire_work: limits.max_input_bytes().saturating_mul(32),
        wire_references: items,
        wire_text: limits.max_output_bytes(),
        nesting: FORMULA_NESTING,
        graph_nodes: items,
        graph_edges: items,
        cache_cells: items,
        formula_work: items,
        graph_work: items.saturating_mul(32),
        retained_bytes: limits.max_output_bytes(),
        scratch_bytes: limits.max_output_bytes(),
        allocations: items,
    }
}

/// One formula-owning table of the graph.
struct Table<'a> {
    geometry: TableGeometry,
    info: u64,
    model: &'a [u8],
}

/// Plan the cache refresh of `writes` (sorted by row, then column, without
/// repeats) into the table `source.table`. `object` resolves an object
/// identifier to its payload in the source (pre-write) state.
///
/// Returns an empty plan when the table has no formula owner.
///
/// # Errors
///
/// [`Error::UnsupportedDependency`] for a graph or a write outside the
/// supported policy (module docs); [`Error::InvalidSource`] for a malformed
/// graph or storage; [`Error::Limit`] or [`Error::Allocation`] when `limits`
/// is exceeded.
pub fn plan<'a>(
    source: Source<'a>,
    writes: &[Write<'_>],
    object: &dyn Fn(u64) -> Option<&'a [u8]>,
    limits: &Limits,
) -> Result<Plan, Error> {
    // A write no formula reads, into a cell that holds none, needs no plan:
    // the graph is not consulted (as the cell writer's admission did before
    // record 0376), so unrelated owners the planner declines do not refuse it.
    let mut owner_payloads = Vec::new();
    for owner in source.owners {
        push(&mut owner_payloads, owner.bytes)?;
    }
    let mut inert = true;
    for write in writes {
        let state = table_write::cell_formula_state(
            &owner_payloads,
            object,
            source.table,
            write.row,
            write.column,
            limits,
        )
        .map_err(map_write)?;
        inert &= state == table_write::FormulaState::Inert;
    }
    if inert {
        return Ok(Plan::default());
    }
    let cache = cache_limits(limits);
    let mut tables: Vec<Table<'a>> = Vec::new();
    let mut records: Vec<DependencyPayload<'a>> = Vec::new();
    let mut ranges: Vec<DependencyPayload<'a>> = Vec::new();
    let mut selected = None;
    for owner in source.owners {
        let (snapshot, _report) = dependency::decode_formula_owner_dependencies_with_report(
            owner.bytes,
            decode_options(cache),
        )
        .map_err(map_decode_failure)?;
        let identity = TableIdentity {
            owner: snapshot.internal_formula_owner_id(),
            uuid_lower: snapshot.formula_owner_uid().lower(),
            uuid_upper: snapshot.formula_owner_uid().upper(),
        };
        for tile in tiled_references(owner.bytes, 13, limits)? {
            let bytes = object(tile).ok_or(Error::InvalidSource)?;
            push(
                &mut records,
                DependencyPayload {
                    identifier: tile,
                    bytes,
                },
            )?;
        }
        for tile in tiled_references(owner.bytes, 15, limits)? {
            let bytes = object(tile).ok_or(Error::InvalidSource)?;
            push(
                &mut ranges,
                DependencyPayload {
                    identifier: tile,
                    bytes,
                },
            )?;
        }
        let Some(info) = snapshot
            .formula_owner()
            .map(|reference| reference.identifier())
        else {
            continue;
        };
        if snapshot.owner_kind() != Some(TABLE_OWNER_KIND) {
            continue;
        }
        if info == source.table {
            if selected.replace(identity).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        let info_payload = object(info).ok_or(Error::InvalidSource)?;
        let model_id = reference(info_payload, 2, limits)?.ok_or(Error::InvalidSource)?;
        let model = object(model_id).ok_or(Error::InvalidSource)?;
        push(
            &mut tables,
            Table {
                geometry: geometry(identity, model, limits)?,
                info,
                model,
            },
        )?;
    }
    let Some(identity) = selected else {
        return Ok(Plan::default());
    };
    tables.sort_unstable_by_key(|table| table.geometry.identity.owner);
    if tables
        .windows(2)
        .any(|pair| pair[0].geometry.identity.owner == pair[1].geometry.identity.owner)
    {
        return Err(Error::InvalidSource);
    }
    records.sort_unstable_by_key(|record| record.identifier);
    records.dedup_by_key(|record| record.identifier);
    ranges.sort_unstable_by_key(|range| range.identifier);
    ranges.dedup_by_key(|range| range.identifier);
    let selected_table = tables
        .iter()
        .find(|table| table.info == source.table)
        .ok_or(Error::InvalidSource)?;
    let geometries: Vec<TableGeometry> = tables.iter().map(|table| table.geometry).collect();

    // The selected table's formula cells and their formula-list entries.
    let storage = TableStorage::read(selected_table.model, object, limits)?;
    let list = storage.formula_list.ok_or(Error::InvalidSource)?;
    let root = object(list).ok_or(Error::InvalidSource)?;
    let root_fields = table_write::fields(root, limits).map_err(map_write)?;
    if root_fields.iter().any(|field| field.number == 4) {
        return Err(Error::UnsupportedDependency(
            Unsupported::SegmentedFormulaList,
        ));
    }
    let expected = root_fields.iter().filter(|field| field.number == 3).count();
    let (entries, _usage) = collect_formula_list(identity, root, &[], expected, cache)?;
    let mut cells = Vec::new();
    let mut payloads = Vec::new();
    for (tile, base, payload) in &storage.tiles {
        for (row, column, bytes) in table_write::stored_cells(payload, limits).map_err(map_write)? {
            let view = BncCellView::parse(bytes, limits).map_err(|_| Error::InvalidSource)?;
            let StoredValue::Formula(key) = view.stored_value() else {
                continue;
            };
            let coordinate = Coordinate {
                row: base.checked_add(row).ok_or(Error::InvalidSource)?,
                column,
            };
            let entry = entries
                .binary_search_by_key(&key, |entry| entry.key)
                .ok()
                .and_then(|index| entries.get(index))
                .ok_or(Error::InvalidSource)?;
            push(
                &mut cells,
                FormulaCell {
                    owner: identity.owner,
                    coordinate,
                    cache_object: *tile,
                },
            )?;
            push(
                &mut payloads,
                FormulaPayload {
                    owner: identity.owner,
                    coordinate,
                    key,
                    bytes: entry.bytes,
                },
            )?;
        }
    }
    cells.sort_unstable_by_key(|cell| cell.coordinate);
    payloads.sort_unstable_by_key(|payload| payload.coordinate);
    validate_formula_cell_payloads(&cells, &payloads, cache)?;

    let mut baseline = Baseline { tables: Vec::new() };
    for table in &tables {
        let values = if table.info == source.table {
            TableValues::read(&storage.tiles, limits)?
        } else {
            TableValues::read(
                &TableStorage::read(table.model, object, limits)?.tiles,
                limits,
            )?
        };
        push(
            &mut baseline.tables,
            (table.geometry.identity.owner, values),
        )?;
    }

    let (mut evaluator, _usage) = StrictEvaluator::new(&entries, &payloads, &geometries, cache)?;
    let mut overlay = Vec::new();
    for write in writes {
        let value = match write.value {
            CellInput::Empty => FinalValue::clear(),
            CellInput::Number(number) => FinalValue::number(
                litchi_iwa_common::formula::FiniteF64::new(*number)
                    .map_err(|_| Error::InvalidSource)?,
            ),
            CellInput::Boolean(flag) => FinalValue::boolean(*flag),
            CellInput::Text(_) => FinalValue::aggregate_ignored(),
        };
        push(
            &mut overlay,
            FinalCell {
                table: identity,
                coordinate: Coordinate {
                    row: write.row,
                    column: write.column,
                },
                value,
            },
        )?;
    }
    let cache_source = CacheSource {
        selected_table: identity,
        tables: &geometries,
        engine: source.engine,
        owners: source.owners,
        record_tiles: &records,
        range_tiles: &ranges,
        formulas: &cells,
        source_formulas: &cells,
    };
    let planned = plan_final_cache(cache_source, &overlay, cache, &baseline, &mut evaluator)?;
    if !planned.removals.is_empty() {
        return Err(Error::UnsupportedDependency(Unsupported::Formula));
    }
    let refreshed =
        usize::try_from(planned.usage.cache_hosts_refreshed).map_err(|_| Error::InvalidSource)?;
    let mut tiles = Vec::new();
    for rewrite in planned.rewrites {
        let (_, base, _) = storage
            .tiles
            .iter()
            .find(|(tile, _, _)| *tile == rewrite.cache_object)
            .ok_or(Error::InvalidSource)?;
        let mut changes = Vec::new();
        for cell in rewrite.cells {
            if cell.owner != identity.owner {
                return Err(Error::UnsupportedDependency(Unsupported::ExternalOwner));
            }
            let local = cell
                .coordinate
                .row
                .checked_sub(*base)
                .filter(|local| u64::from(*local) < storage.tile_rows)
                .ok_or(Error::InvalidSource)?;
            push(&mut changes, (local, cell.coordinate.column, cell.value))?;
        }
        push(
            &mut tiles,
            TileCaches {
                tile: rewrite.cache_object,
                cells: changes,
            },
        )?;
    }
    tiles.sort_unstable_by_key(|tile| tile.tile);
    Ok(Plan { tiles, refreshed })
}

/// The tiles and formula list of one table's data store.
struct TableStorage<'a> {
    /// `(tile object, first table row, payload)`.
    tiles: Vec<(u64, u32, &'a [u8])>,
    tile_rows: u64,
    formula_list: Option<u64>,
}

impl<'a> TableStorage<'a> {
    fn read(
        model: &'a [u8],
        object: &dyn Fn(u64) -> Option<&'a [u8]>,
        limits: &Limits,
    ) -> Result<Self, Error> {
        let store = table_write::single_bytes(model, 4, limits)
            .map_err(map_write)?
            .ok_or(Error::InvalidSource)?;
        let tile_storage = table_write::single_bytes(store, 3, limits)
            .map_err(map_write)?
            .ok_or(Error::InvalidSource)?;
        let tile_rows = table_write::single_varint(tile_storage, 2, limits)
            .map_err(map_write)?
            .unwrap_or(DEFAULT_TILE_ROWS)
            .max(1);
        let mut tiles = Vec::new();
        for field in table_write::fields(tile_storage, limits).map_err(map_write)? {
            if field.number != 1 {
                continue;
            }
            let key = table_write::single_varint(field.payload, 1, limits)
                .map_err(map_write)?
                .ok_or(Error::InvalidSource)?;
            let tile = reference(field.payload, 2, limits)?.ok_or(Error::InvalidSource)?;
            let base = key
                .checked_mul(tile_rows)
                .and_then(|base| u32::try_from(base).ok())
                .ok_or(Error::InvalidSource)?;
            let payload = object(tile).ok_or(Error::InvalidSource)?;
            push(&mut tiles, (tile, base, payload))?;
        }
        Ok(Self {
            tiles,
            tile_rows,
            formula_list: reference(store, 6, limits)?,
        })
    }
}

/// One table's stored values, as the planner's baseline sees them.
struct TableValues {
    values: Vec<(Coordinate, ScalarValue)>,
    aggregate_ignored: Vec<Coordinate>,
    unsupported: Vec<Coordinate>,
}

impl TableValues {
    fn read(tiles: &[(u64, u32, &[u8])], limits: &Limits) -> Result<Self, Error> {
        let mut values = Vec::new();
        let mut aggregate_ignored = Vec::new();
        let mut unsupported = Vec::new();
        for (_, base, payload) in tiles {
            for (row, column, bytes) in
                table_write::stored_cells(payload, limits).map_err(map_write)?
            {
                let coordinate = Coordinate {
                    row: base.checked_add(row).ok_or(Error::InvalidSource)?,
                    column,
                };
                let view = BncCellView::parse(bytes, limits).map_err(|_| Error::InvalidSource)?;
                match view.stored_value() {
                    StoredValue::Empty => {},
                    StoredValue::Number
                    | StoredValue::Date
                    | StoredValue::Boolean
                    | StoredValue::Duration => {
                        let value = match view.cached_scalar() {
                            Some(CachedScalar::Number(value)) => ScalarValue::Number(value),
                            Some(CachedScalar::Boolean(value)) => ScalarValue::Boolean(value),
                            Some(CachedScalar::Date(value)) => ScalarValue::Date(value),
                            Some(CachedScalar::Duration(value)) => ScalarValue::Duration(value),
                            Some(CachedScalar::Unsupported(_)) | None => {
                                push(&mut unsupported, coordinate)?;
                                continue;
                            },
                        };
                        push(&mut values, (coordinate, value))?;
                    },
                    StoredValue::Text(_) | StoredValue::RichText(_) => {
                        push(&mut aggregate_ignored, coordinate)?;
                    },
                    StoredValue::Formula(_) | StoredValue::Error | StoredValue::Unsupported(_) => {
                        push(&mut unsupported, coordinate)?;
                    },
                }
            }
        }
        values.sort_unstable_by_key(|entry| entry.0);
        aggregate_ignored.sort_unstable();
        unsupported.sort_unstable();
        Ok(Self {
            values,
            aggregate_ignored,
            unsupported,
        })
    }
}

/// Every formula-owning table's values, sorted by internal owner.
struct Baseline {
    tables: Vec<(u32, TableValues)>,
}

impl Baseline {
    fn table(&self, owner: u32) -> Result<&TableValues, Error> {
        self.tables
            .binary_search_by_key(&owner, |entry| entry.0)
            .ok()
            .and_then(|index| self.tables.get(index))
            .map(|entry| &entry.1)
            .ok_or(Error::UnsupportedDependency(Unsupported::MissingOwner))
    }
}

impl CacheBaseline for Baseline {
    fn value(&self, owner: u32, coordinate: Coordinate) -> Result<Option<ScalarValue>, Error> {
        let table = self.table(owner)?;
        if table.unsupported.binary_search(&coordinate).is_ok() {
            return Err(Error::UnsupportedDependency(Unsupported::Formula));
        }
        if table.aggregate_ignored.binary_search(&coordinate).is_ok() {
            return Err(Error::UnsupportedDependency(Unsupported::TextScalar));
        }
        Ok(table
            .values
            .binary_search_by_key(&coordinate, |entry| entry.0)
            .ok()
            .and_then(|index| table.values.get(index))
            .map(|entry| entry.1))
    }

    fn lookup_work(&self, owner: u32) -> Result<usize, Error> {
        let table = self.table(owner)?;
        super::binary_search_work(self.tables.len())
            .checked_add(super::baseline_lookup_work(
                table.values.len(),
                table.aggregate_ignored.len(),
                table.unsupported.len(),
            )?)
            .ok_or(Error::InvalidSource)
    }
}

fn geometry(
    identity: TableIdentity,
    model: &[u8],
    limits: &Limits,
) -> Result<TableGeometry, Error> {
    let field = |number| -> Result<Option<u32>, Error> {
        table_write::single_varint(model, number, limits)
            .map_err(map_write)?
            .map(|value| u32::try_from(value).map_err(|_| Error::InvalidSource))
            .transpose()
    };
    Ok(TableGeometry {
        identity,
        rows: field(6)?.ok_or(Error::InvalidSource)?,
        columns: field(7)?.ok_or(Error::InvalidSource)?,
        header_rows: field(9)?.unwrap_or(0),
        header_columns: field(10)?.unwrap_or(0),
        footer_rows: field(11)?.unwrap_or(0),
    })
}

/// The identifier of the `TSP.Reference` in field `number` of `payload`.
fn reference(payload: &[u8], number: u32, limits: &Limits) -> Result<Option<u64>, Error> {
    table_write::single_bytes(payload, number, limits)
        .map_err(map_write)?
        .map(|reference| {
            table_write::single_varint(reference, 1, limits)
                .map_err(map_write)?
                .ok_or(Error::InvalidSource)
        })
        .transpose()
}

/// The tile references of an owner's tiled dependency envelope (field
/// `number`: 13 for cell-record tiles, 15 for range-precedent tiles).
fn tiled_references(owner: &[u8], number: u32, limits: &Limits) -> Result<Vec<u64>, Error> {
    let mut references = Vec::new();
    let Some(envelope) = table_write::single_bytes(owner, number, limits).map_err(map_write)?
    else {
        return Ok(references);
    };
    for field in table_write::fields(envelope, limits).map_err(map_write)? {
        if field.number != 1 {
            continue;
        }
        let identifier = table_write::single_varint(field.payload, 1, limits)
            .map_err(map_write)?
            .ok_or(Error::InvalidSource)?;
        push(&mut references, identifier)?;
    }
    Ok(references)
}

fn push<T>(values: &mut Vec<T>, value: T) -> Result<(), Error> {
    values
        .try_reserve(1)
        .map_err(|_| Error::Allocation { amount: 1 })?;
    values.push(value);
    Ok(())
}

fn copy(bytes: &[u8]) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| Error::Allocation {
            amount: bytes.len(),
        })?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn map_write(error: WriteError) -> Error {
    match error {
        WriteError::Unsupported => Error::UnsupportedDependency(Unsupported::Storage),
        WriteError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                crate::LimitKind::Fields => super::LimitKind::WireFields,
                crate::LimitKind::OutputBytes => super::LimitKind::RetainedBytes,
                crate::LimitKind::Items => super::LimitKind::GraphNodes,
                _ => super::LimitKind::WireBytes,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        WriteError::Allocation { requested } => Error::Allocation { amount: requested },
        _ => Error::InvalidSource,
    }
}
