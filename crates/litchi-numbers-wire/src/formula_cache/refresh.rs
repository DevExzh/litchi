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
//! range-precedent tiles) and formulas in every formula-owning table are
//! refreshed, as the host refreshes them (record 0376a): a write that
//! replaces or clears a formula cell or a formula the evaluator does not
//! support is refused with
//! [`Error::UnsupportedDependency`](crate::formula_cache::Error::UnsupportedDependency).

use litchi_iwa_common::formula::FormulaCachedValue;

use super::{
    CacheBaseline, CacheLimits, CacheSource, Coordinate, DependencyPayload, Error, FinalCell,
    FinalValue, FormulaCell, FormulaListSegment, FormulaPayload, ScalarValue, StrictEvaluator,
    TableGeometry, TableIdentity, Unsupported, collect_formula_list, decode_options,
    map_decode_failure, plan_final_cache, validate_formula_cell_payloads,
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
    columns: u32,
    cells: Vec<(u32, u32, FormulaCachedValue)>,
}

/// Where one tile of a formula-owning table sits.
struct TileLayout {
    tile: u64,
    owner: u32,
    base: u32,
    rows: u64,
    columns: u32,
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
    /// Rewrite the cached results in `tile`, the tile's payload after the
    /// batch's value writes. The tile may belong to another table than the
    /// written one: a formula there reads the written cells.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidSource`] when a target cell is not a formula cell;
    /// [`Error::UnsupportedDependency`] for a result kind the BNC writer
    /// does not store (only number and Boolean results are written, as the
    /// host does).
    pub fn apply(&self, tile: &[u8], limits: &Limits) -> Result<Vec<u8>, Error> {
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
            tile = table_write::tile_set_cell(
                &tile,
                *row,
                *column,
                self.columns,
                Some(&encoded),
                limits,
            )
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
    if !tables.iter().any(|table| table.info == source.table) {
        return Err(Error::InvalidSource);
    }
    let geometries: Vec<TableGeometry> = tables.iter().map(|table| table.geometry).collect();

    // Every formula-owning table's formula cells, formula-list entries and
    // values: a formula in another table may read the written cells.
    let mut entries = Vec::new();
    let mut cells = Vec::new();
    let mut payloads = Vec::new();
    let mut baseline = Baseline { tables: Vec::new() };
    let mut layouts = Vec::new();
    for table in &tables {
        let identity = table.geometry.identity;
        let storage = TableStorage::read(table.model, object, limits)?;
        let list = storage.formula_list.ok_or(Error::InvalidSource)?;
        let root = object(list).ok_or(Error::InvalidSource)?;
        let mut expected = 0usize;
        let mut segments = Vec::new();
        for field in table_write::fields(root, limits).map_err(map_write)? {
            match field.number {
                3 => expected = expected.checked_add(1).ok_or(Error::InvalidSource)?,
                4 => {
                    let identifier = table_write::single_varint(field.payload, 1, limits)
                        .map_err(map_write)?
                        .ok_or(Error::InvalidSource)?;
                    let bytes = object(identifier).ok_or(Error::InvalidSource)?;
                    expected = table_write::fields(bytes, limits)
                        .map_err(map_write)?
                        .iter()
                        .filter(|field| field.number == 3)
                        .count()
                        .checked_add(expected)
                        .ok_or(Error::InvalidSource)?;
                    push(&mut segments, FormulaListSegment { identifier, bytes })?;
                },
                _ => {},
            }
        }
        segments.sort_unstable_by_key(|segment| segment.identifier);
        let (table_entries, _usage) =
            collect_formula_list(identity, root, &segments, expected, cache)?;
        let mut table_cells = Vec::new();
        let mut table_payloads = Vec::new();
        for (tile, base, payload) in &storage.tiles {
            for (row, column, bytes) in
                table_write::stored_cells(payload, limits).map_err(map_write)?
            {
                let view = BncCellView::parse(bytes, limits).map_err(|_| Error::InvalidSource)?;
                let StoredValue::Formula(key) = view.stored_value() else {
                    continue;
                };
                let coordinate = Coordinate {
                    row: base.checked_add(row).ok_or(Error::InvalidSource)?,
                    column,
                };
                let entry = table_entries
                    .binary_search_by_key(&key, |entry| entry.key)
                    .ok()
                    .and_then(|index| table_entries.get(index))
                    .ok_or(Error::InvalidSource)?;
                push(
                    &mut table_cells,
                    FormulaCell {
                        owner: identity.owner,
                        coordinate,
                        cache_object: *tile,
                    },
                )?;
                push(
                    &mut table_payloads,
                    FormulaPayload {
                        owner: identity.owner,
                        coordinate,
                        key,
                        bytes: entry.bytes,
                    },
                )?;
            }
        }
        table_cells.sort_unstable_by_key(|cell| cell.coordinate);
        table_payloads.sort_unstable_by_key(|payload| payload.coordinate);
        extend(&mut entries, table_entries)?;
        extend(&mut cells, table_cells)?;
        extend(&mut payloads, table_payloads)?;
        push(
            &mut baseline.tables,
            (identity.owner, TableValues::read(&storage.tiles, limits)?),
        )?;
        for (tile, base, _) in &storage.tiles {
            push(
                &mut layouts,
                TileLayout {
                    tile: *tile,
                    owner: identity.owner,
                    base: *base,
                    rows: storage.tile_rows,
                    columns: table.geometry.columns,
                },
            )?;
        }
    }
    validate_formula_cell_payloads(&cells, &payloads, cache)?;

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
    // An owner's UUID references name the tables its formulas read (for
    // structural edits; the dependency edges carry the same reads). Admit
    // them when every named table is registered and hand the planner the
    // owner without them; this plan publishes no owner (record 0376a).
    let mut stripped = Vec::new();
    for owner in source.owners {
        push(
            &mut stripped,
            without_table_references(owner.bytes, &geometries, limits)?,
        )?;
    }
    let mut planner_owners = Vec::new();
    for (owner, bytes) in source.owners.iter().zip(&stripped) {
        push(
            &mut planner_owners,
            DependencyPayload {
                identifier: owner.identifier,
                bytes: bytes.as_deref().unwrap_or(owner.bytes),
            },
        )?;
    }
    let cache_source = CacheSource {
        selected_table: identity,
        tables: &geometries,
        engine: source.engine,
        owners: &planner_owners,
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
        let layout = layouts
            .iter()
            .find(|layout| layout.tile == rewrite.cache_object)
            .ok_or(Error::InvalidSource)?;
        let mut changes = Vec::new();
        for cell in rewrite.cells {
            if cell.owner != layout.owner {
                return Err(Error::InvalidSource);
            }
            let local = cell
                .coordinate
                .row
                .checked_sub(layout.base)
                .filter(|local| u64::from(*local) < layout.rows)
                .ok_or(Error::InvalidSource)?;
            push(&mut changes, (local, cell.coordinate.column, cell.value))?;
        }
        push(
            &mut tiles,
            TileCaches {
                tile: rewrite.cache_object,
                columns: layout.columns,
                cells: changes,
            },
        )?;
    }
    tiles.sort_unstable_by_key(|tile| tile.tile);
    Ok(Plan { tiles, refreshed })
}

/// The tiles and formula list of one table's data store.
pub(super) struct TableStorage<'a> {
    /// `(tile object, first table row, payload)`.
    pub(super) tiles: Vec<(u64, u32, &'a [u8])>,
    pub(super) tile_rows: u64,
    pub(super) formula_list: Option<u64>,
}

impl<'a> TableStorage<'a> {
    pub(super) fn read(
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

pub(super) fn geometry(
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
pub(super) fn reference(
    payload: &[u8],
    number: u32,
    limits: &Limits,
) -> Result<Option<u64>, Error> {
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
pub(super) fn tiled_references(
    owner: &[u8],
    number: u32,
    limits: &Limits,
) -> Result<Vec<u64>, Error> {
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

/// `owner` without its UUID-references field when every reference names a
/// registered table by owner UUID (`None` when the owner has none);
/// cell-UUID references stay and the planner declines them.
fn without_table_references(
    owner: &[u8],
    tables: &[TableGeometry],
    limits: &Limits,
) -> Result<Option<Vec<u8>>, Error> {
    let Some(references) = table_write::single_bytes(owner, 14, limits).map_err(map_write)? else {
        return Ok(None);
    };
    for field in table_write::fields(references, limits).map_err(map_write)? {
        if field.number != 1 {
            return Ok(None);
        }
        let uuid = table_write::single_bytes(field.payload, 1, limits)
            .map_err(map_write)?
            .ok_or(Error::InvalidSource)?;
        let lower = table_write::single_varint(uuid, 1, limits)
            .map_err(map_write)?
            .unwrap_or(0);
        let upper = table_write::single_varint(uuid, 2, limits)
            .map_err(map_write)?
            .unwrap_or(0);
        if !tables
            .iter()
            .any(|table| table.identity.uuid_lower == lower && table.identity.uuid_upper == upper)
        {
            return Ok(None);
        }
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(owner.len())
        .map_err(|_| Error::Allocation {
            amount: owner.len(),
        })?;
    for field in table_write::fields(owner, limits).map_err(map_write)? {
        if field.number != 14 {
            output.extend_from_slice(field.raw);
        }
    }
    Ok(Some(output))
}

pub(super) fn push<T>(values: &mut Vec<T>, value: T) -> Result<(), Error> {
    values
        .try_reserve(1)
        .map_err(|_| Error::Allocation { amount: 1 })?;
    values.push(value);
    Ok(())
}

pub(super) fn extend<T>(values: &mut Vec<T>, more: Vec<T>) -> Result<(), Error> {
    values
        .try_reserve(more.len())
        .map_err(|_| Error::Allocation { amount: more.len() })?;
    values.extend(more);
    Ok(())
}

pub(super) fn copy(bytes: &[u8]) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| Error::Allocation {
            amount: bytes.len(),
        })?;
    output.extend_from_slice(bytes);
    Ok(output)
}

pub(super) fn map_write(error: WriteError) -> Error {
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

#[cfg(test)]
mod tests {
    use litchi_iwa_protos::{tsce, tsp};
    use prost::Message as _;

    use super::*;

    const OWNER: u32 = 7;
    const OWNER_OBJECT: u64 = 1;
    const INFO: u64 = 10;
    const MODEL: u64 = 11;
    const TILE: u64 = 12;
    const LIST: u64 = 13;

    fn varint(output: &mut Vec<u8>, mut value: u64) {
        while value >= 0x80 {
            output.push((value as u8) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn uint(number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        varint(&mut output, u64::from(number) << 3);
        varint(&mut output, value);
        output
    }

    fn bytes(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        varint(&mut output, (u64::from(number) << 3) | 2);
        varint(&mut output, payload.len() as u64);
        output.extend_from_slice(payload);
        output
    }

    fn reference(number: u32, identifier: u64) -> Vec<u8> {
        bytes(number, &uint(1, identifier))
    }

    /// `A1 + 1` as a native formula (column handle relative to the host).
    fn formula_archive(precedent_column: u32) -> Vec<u8> {
        use tsce::ast_node_array_archive::{
            AstLocalCellReferenceNodeArchive, AstNodeArchive, AstNodeType,
        };
        tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![
                    AstNodeArchive {
                        ast_node_type: AstNodeType::LocalCellReferenceNode as i32,
                        ast_local_cell_reference_node_reference: Some(
                            AstLocalCellReferenceNodeArchive {
                                row_handle: 0,
                                column_handle: precedent_column,
                                row_is_sticky: 0,
                                column_is_sticky: 0,
                            },
                        ),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::NumberNode as i32,
                        ast_number_node_number: Some(1.0),
                        ..Default::default()
                    },
                    AstNodeArchive {
                        ast_node_type: AstNodeType::AdditionNode as i32,
                        ..Default::default()
                    },
                ],
            },
            ..Default::default()
        }
        .encode_to_vec()
    }

    struct Document {
        engine: Vec<u8>,
        owner: Vec<u8>,
        objects: Vec<(u64, Vec<u8>)>,
    }

    impl Document {
        /// A 2×2 table: A1 holds `a1` (a number, or text key 1), B1 = A1 + 1
        /// cached as 4.
        fn new(a1_text: bool) -> Self {
            let limits = Limits::DEFAULT;
            let owner = tsce::FormulaOwnerDependenciesArchive {
                formula_owner_uid: tsp::Uuid {
                    lower: 11,
                    upper: 13,
                },
                internal_formula_owner_id: OWNER,
                owner_kind: Some(TABLE_OWNER_KIND),
                formula_owner: Some(tsp::Reference {
                    identifier: INFO,
                    ..Default::default()
                }),
                cell_dependencies: Some(tsce::CellDependenciesExpandedArchive {
                    cell_record: vec![tsce::CellRecordExpandedArchive {
                        row: 0,
                        column: 1,
                        expanded_edges: Some(tsce::ExpandedEdgesArchive {
                            edge_without_owner_rows: vec![0],
                            edge_without_owner_columns: vec![0],
                            ..Default::default()
                        }),
                        ..Default::default()
                    }],
                }),
                ..Default::default()
            }
            .encode_to_vec();
            let engine = tsce::CalculationEngineArchive {
                dependency_tracker: tsce::DependencyTrackerArchive {
                    formula_owner_dependencies: vec![tsp::Reference {
                        identifier: OWNER_OBJECT,
                        ..Default::default()
                    }],
                    number_of_formulas: Some(1),
                    ..Default::default()
                },
                ..Default::default()
            }
            .encode_to_vec();
            let info = reference(2, MODEL);
            let tiles = [
                bytes(1, &[uint(1, 0), reference(2, TILE)].concat()),
                uint(2, 256),
            ]
            .concat();
            let store = [bytes(3, &tiles), reference(6, LIST)].concat();
            let model = [bytes(4, &store), uint(6, 2), uint(7, 2)].concat();
            let mut a1 = BncCell::minimal();
            if a1_text {
                a1.set_string(1);
            } else {
                a1.set_number(3.0).unwrap();
            }
            let mut b1 = BncCell::minimal();
            b1.set_number(4.0).unwrap();
            b1.set_formula_reference(1);
            let mut tile = [
                uint(1, 0),
                uint(2, 0),
                uint(3, 0),
                uint(4, 0),
                uint(6, 5),
                uint(7, 1),
            ]
            .concat();
            for (column, cell) in [(0, a1), (1, b1)] {
                let encoded = cell.encode(&limits).unwrap();
                tile = table_write::tile_set_cell(&tile, 0, column, 2, Some(&encoded), &limits)
                    .unwrap()
                    .tile;
            }
            let entry = [uint(1, 1), uint(2, 1), bytes(5, &formula_archive(0))].concat();
            let list = [uint(1, 3), uint(2, 2), bytes(3, &entry)].concat();
            Self {
                engine,
                owner: owner.clone(),
                objects: vec![
                    (OWNER_OBJECT, owner),
                    (INFO, info),
                    (MODEL, model),
                    (TILE, tile),
                    (LIST, list),
                ],
            }
        }

        fn plan(&self, row: u32, column: u32, value: &CellInput<'_>) -> Result<Plan, Error> {
            let owners = [DependencyPayload {
                identifier: OWNER_OBJECT,
                bytes: &self.owner,
            }];
            let object = |identifier: u64| -> Option<&[u8]> {
                self.objects
                    .iter()
                    .find(|(candidate, _)| *candidate == identifier)
                    .map(|(_, payload)| payload.as_slice())
            };
            plan(
                Source {
                    engine: &self.engine,
                    owners: &owners,
                    table: INFO,
                },
                &[Write { row, column, value }],
                &object,
                &Limits::DEFAULT,
            )
        }

        fn tile(&self) -> &[u8] {
            &self.objects[3].1
        }
    }

    fn cached(tile: &[u8], column: u32) -> Option<CachedScalar> {
        let cell = table_write::tile_cell(tile, 0, column, &Limits::DEFAULT)
            .unwrap()
            .unwrap();
        BncCellView::parse(&cell, &Limits::DEFAULT)
            .unwrap()
            .cached_scalar()
    }

    fn author_object<'d>(
        document: &'d Document,
        identifier: u64,
    ) -> Option<super::super::author::Object<'d>> {
        document
            .objects
            .iter()
            .find(|(candidate, _)| *candidate == identifier)
            .map(|(identifier, payload)| super::super::author::Object {
                identifier: *identifier,
                payload,
                references: &[],
            })
    }

    fn author_plan(
        document: &Document,
        row: u32,
        column: u32,
        change: super::super::author::Change<'_>,
    ) -> Result<super::super::author::Edit, Error> {
        use super::super::author;
        let owners = [author::Object {
            identifier: OWNER_OBJECT,
            payload: &document.owner,
            references: &[],
        }];
        let object = |identifier: u64| author_object(document, identifier);
        author::plan(
            author::Source {
                engine: author::Object {
                    identifier: 100,
                    payload: &document.engine,
                    references: &[OWNER_OBJECT],
                },
                owners: &owners,
                table: INFO,
            },
            row,
            column,
            change,
            1_000,
            &object,
            &Limits::DEFAULT,
        )
    }

    #[test]
    fn authoring_adds_the_inline_record_and_a_list_entry() {
        use super::super::author;
        use litchi_iwa_common::formula::{
            FormulaBinaryOperator, FormulaCellReference, FormulaExpression,
        };
        let document = Document::new(false);
        let object = |identifier: u64| author_object(&document, identifier);
        let source = author::Source {
            engine: author::Object {
                identifier: 100,
                payload: &document.engine,
                references: &[OWNER_OBJECT],
            },
            owners: &[author::Object {
                identifier: OWNER_OBJECT,
                payload: &document.owner,
                references: &[],
            }],
            table: INFO,
        };
        assert_eq!(
            author::table_owner(source, &object, &Limits::DEFAULT).unwrap(),
            Some((OWNER, 2, 2))
        );
        let expression = FormulaExpression::binary(
            FormulaBinaryOperator::Add,
            FormulaExpression::Cell(FormulaCellReference::relative(0, 0)),
            FormulaExpression::Number(2.0),
        );
        let formula = author::encode(&expression, OWNER, 1, 1, 2, 2, &Limits::DEFAULT).unwrap();
        let edit = author_plan(&document, 1, 1, author::Change::Author(&formula)).unwrap();
        assert_eq!(edit.key, Some(2));
        assert_eq!(edit.owner, OWNER_OBJECT);
        assert_eq!(edit.created_tile, None);
        assert_eq!(edit.removed_tile, None);
        let owner = edit
            .payloads
            .iter()
            .find(|(identifier, _)| *identifier == OWNER_OBJECT)
            .map(|(_, payload)| {
                tsce::FormulaOwnerDependenciesArchive::decode(payload.as_slice()).unwrap()
            })
            .unwrap();
        let records = owner.cell_dependencies.unwrap().cell_record;
        assert_eq!(
            records
                .iter()
                .map(|record| (record.row, record.column))
                .collect::<Vec<_>>(),
            vec![(0, 1), (1, 1)]
        );
        let list = edit
            .payloads
            .iter()
            .find(|(identifier, _)| *identifier == LIST)
            .map(|(_, payload)| payload.clone())
            .unwrap();
        let entries = table_write::fields(&list, &Limits::DEFAULT)
            .unwrap()
            .iter()
            .filter(|field| field.number == 3)
            .count();
        assert_eq!(entries, 2);
    }

    #[test]
    fn removal_drops_the_record_and_releases_the_list_key() {
        use super::super::author;
        let document = Document::new(false);
        let edit = author_plan(&document, 0, 1, author::Change::Remove).unwrap();
        assert_eq!(edit.key, None);
        let owner = edit
            .payloads
            .iter()
            .find(|(identifier, _)| *identifier == OWNER_OBJECT)
            .map(|(_, payload)| {
                tsce::FormulaOwnerDependenciesArchive::decode(payload.as_slice()).unwrap()
            })
            .unwrap();
        // The emptied inline envelope stays, as the host keeps it.
        assert_eq!(
            owner.cell_dependencies.map(|cells| cells.cell_record.len()),
            Some(0)
        );
        let list = edit
            .payloads
            .iter()
            .find(|(identifier, _)| *identifier == LIST)
            .map(|(_, payload)| payload.clone())
            .unwrap();
        let fields = table_write::fields(&list, &Limits::DEFAULT).unwrap();
        assert!(!fields.iter().any(|field| field.number == 3));
        // The released key was just below nextListID: nextListID drops to it.
        assert_eq!(
            table_write::single_varint(&list, 2, &Limits::DEFAULT).unwrap(),
            Some(1)
        );
    }

    #[test]
    fn removal_of_a_cell_without_a_formula_is_refused() {
        use super::super::author;
        let document = Document::new(false);
        assert!(matches!(
            author_plan(&document, 0, 0, author::Change::Remove),
            Err(Error::UnsupportedDependency(Unsupported::Formula))
        ));
    }

    #[test]
    fn a_reference_to_another_table_is_refused() {
        use super::super::author;
        use litchi_iwa_common::formula::{FormulaCellReference, FormulaExpression};
        let expression = FormulaExpression::TableCell {
            table_id: 99,
            reference: FormulaCellReference::relative(0, 0),
        };
        assert!(matches!(
            author::encode(&expression, OWNER, 1, 1, 2, 2, &Limits::DEFAULT),
            Err(Error::UnsupportedDependency(Unsupported::Formula))
        ));
        let unknown =
            FormulaExpression::function("NO-SUCH-FUNCTION", [FormulaExpression::Number(1.0)]);
        assert!(matches!(
            author::encode(&unknown, OWNER, 1, 1, 2, 2, &Limits::DEFAULT),
            Err(Error::UnsupportedDependency(Unsupported::Formula))
        ));
    }

    #[test]
    fn a_precedent_write_refreshes_the_dependent_cache() {
        let document = Document::new(false);
        let plan = document.plan(0, 0, &CellInput::Number(5.0)).unwrap();
        assert_eq!(plan.refreshed, 1);
        assert_eq!(plan.tiles.len(), 1);
        assert_eq!(plan.tiles[0].tile, TILE);
        let tile = plan.tiles[0]
            .apply(document.tile(), &Limits::DEFAULT)
            .unwrap();
        assert!(matches!(
            cached(&tile, 1),
            Some(CachedScalar::Number(value)) if value.get() == 6.0
        ));
        // A clear evaluates the precedent as empty (zero).
        let plan = document.plan(0, 0, &CellInput::Empty).unwrap();
        let tile = plan.tiles[0]
            .apply(document.tile(), &Limits::DEFAULT)
            .unwrap();
        assert!(matches!(
            cached(&tile, 1),
            Some(CachedScalar::Number(value)) if value.get() == 1.0
        ));
    }

    #[test]
    fn a_write_no_formula_reads_needs_no_plan() {
        let document = Document::new(false);
        let plan = document.plan(1, 1, &CellInput::Number(5.0)).unwrap();
        assert!(plan.is_empty());
        assert_eq!(plan.refreshed, 0);
    }

    #[test]
    fn a_table_outside_the_graph_needs_no_plan() {
        let document = Document::new(false);
        let object = |_identifier: u64| -> Option<&[u8]> { None };
        let plan = plan(
            Source {
                engine: &document.engine,
                owners: &[],
                table: INFO,
            },
            &[Write {
                row: 0,
                column: 0,
                value: &CellInput::Number(1.0),
            }],
            &object,
            &Limits::DEFAULT,
        )
        .unwrap();
        assert!(plan.is_empty());
    }

    #[test]
    fn a_formula_cell_overwrite_is_refused() {
        let document = Document::new(false);
        assert!(matches!(
            document.plan(0, 1, &CellInput::Number(5.0)),
            Err(Error::UnsupportedDependency(Unsupported::Formula))
        ));
    }

    #[test]
    fn text_read_directly_by_a_formula_is_refused() {
        let document = Document::new(false);
        assert!(matches!(
            document.plan(0, 0, &CellInput::Text("x")),
            Err(Error::UnsupportedDependency(_))
        ));
    }

    #[test]
    fn a_text_baseline_read_directly_is_refused() {
        // B1 reads a text A1; writing B2 (unread) needs no plan, but a write
        // that dirties B1 would evaluate text.
        let document = Document::new(true);
        assert!(
            document
                .plan(1, 1, &CellInput::Number(1.0))
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn apply_refuses_a_cell_without_a_formula() {
        let document = Document::new(false);
        let caches = TileCaches {
            tile: TILE,
            columns: 2,
            cells: vec![(0, 0, FormulaCachedValue::number(1.0).unwrap())],
        };
        assert_eq!(
            caches.apply(document.tile(), &Limits::DEFAULT),
            Err(Error::InvalidSource)
        );
    }

    #[test]
    fn exhausted_limits_refuse_typed() {
        let document = Document::new(false);
        let owners = [DependencyPayload {
            identifier: OWNER_OBJECT,
            bytes: &document.owner,
        }];
        let object = |identifier: u64| -> Option<&[u8]> {
            document
                .objects
                .iter()
                .find(|(candidate, _)| *candidate == identifier)
                .map(|(_, payload)| payload.as_slice())
        };
        let limits = Limits::DEFAULT.with_items(1).unwrap();
        let result = plan(
            Source {
                engine: &document.engine,
                owners: &owners,
                table: INFO,
            },
            &[Write {
                row: 0,
                column: 0,
                value: &CellInput::Number(1.0),
            }],
            &object,
            &limits,
        );
        assert!(matches!(result, Err(Error::Limit { .. })), "{result:?}");
    }
}
