//! Final-overlay formula-cache admission for grouped cell transactions.

use core::mem::size_of;

use litchi_iwa_common::formula::FormulaCachedValue;

use crate::{
    Package,
    cell::Value,
    package::table_cells::{DependencyKind, Error, LimitKind, Path},
    table::{
        Table,
        cells::{Change, Input},
    },
};

use super::{budget, cache, message_object_identifier, message_payload, resolve, tile};

#[derive(Debug)]
pub(super) struct PreparedTileCache {
    pub(super) tile_id: u32,
    pub(super) changes: Vec<tile::CacheChange>,
}

/// Cache changes prepared for the grouped parent tile pass.
#[derive(Debug, Default)]
pub(super) struct PreparedCache {
    pub(super) refreshed_hosts: usize,
    pub(super) tiles: Vec<PreparedTileCache>,
}

/// Run one allocation/work phase only after its complete envelope has been
/// admitted, then replace that private reservation with the exact report.
fn authorized_phase<T>(
    budget: &mut budget::TransactionBudget,
    envelope: budget::Usage,
    operation: impl FnOnce() -> Result<(T, budget::Usage), Error>,
) -> Result<T, Error> {
    budget.authorize(envelope)?;
    match operation() {
        Ok((value, actual)) => match budget.record_authorized(actual) {
            Ok(()) => Ok(value),
            Err(error) => {
                budget.cancel_authorization();
                Err(error)
            },
        },
        Err(error) => {
            budget.cancel_authorization();
            Err(error)
        },
    }
}

fn authorized_cache_phase<T>(
    source: &Package,
    budget: &mut budget::TransactionBudget,
    path: Path,
    operation: impl FnOnce(
        cache::CacheLimits,
        budget::Remaining,
    ) -> Result<(T, cache::CacheUsage), Error>,
) -> Result<T, Error> {
    let remaining = budget.remaining()?;
    let limits = cache_limits(source, remaining, path)?;
    authorized_phase(budget, remaining, || {
        let (value, usage) = operation(limits, remaining)?;
        Ok((value, cache_usage(usage, path)?))
    })
}

/// Plan the selected dependency graph against the transaction's final overlay.
///
/// Cache rewrites are grouped by physical tile for the parent's single tile
/// pass. Formula dependency deletion remains fail-closed until its owner
/// archives can be rewritten in the same publication.
pub(super) fn prepare_final_cache(
    source: &Package,
    target: &resolve::Target,
    table: &Table,
    changes: &[Change],
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<PreparedCache, Error> {
    if target.dependencies.engine.is_none()
        && target.dependencies.formula_owners.is_empty()
        && target.dependencies.cell_record_tiles.is_empty()
        && target.dependencies.range_precedent_tiles.is_empty()
    {
        return Ok(PreparedCache::default());
    }
    if !target.dependencies.range_precedent_tiles.is_empty() {
        return Err(cache_refusal(path));
    }
    let engine_route = target
        .dependencies
        .engine
        .ok_or_else(|| cache_refusal(path))?;
    let selected = target
        .dependencies
        .selected_formula_owner
        .ok_or_else(|| cache_refusal(path))?;
    let identity = cache::TableIdentity {
        owner: selected.internal_owner_id,
        uuid_lower: selected.uid_lower,
        uuid_upper: selected.uid_upper,
    };
    let owner_count = target.dependencies.formula_owners.len();
    let record_count = target.dependencies.cell_record_tiles.len();
    let caller_elements = owner_count
        .checked_add(record_count)
        .and_then(|count| count.checked_add(changes.len()))
        .ok_or(Error::InvalidSource { path })?;
    let caller_bytes = owner_count
        .checked_mul(size_of::<cache::DependencyPayload<'_>>())
        .and_then(|bytes| {
            record_count
                .checked_mul(size_of::<cache::DependencyPayload<'_>>())
                .and_then(|record_bytes| bytes.checked_add(record_bytes))
        })
        .and_then(|bytes| {
            changes
                .len()
                .checked_mul(size_of::<cache::FinalCell>())
                .and_then(|overlay_bytes| bytes.checked_add(overlay_bytes))
        })
        .ok_or(Error::InvalidSource { path })?;
    let caller_usage = budget::Usage {
        retained_elements: as_u64(caller_elements),
        retained_bytes: as_u64(caller_bytes),
        allocation_events: as_u64(
            usize::from(owner_count != 0)
                .checked_add(usize::from(record_count != 0))
                .and_then(|count| count.checked_add(usize::from(!changes.is_empty())))
                .ok_or(Error::InvalidSource { path })?,
        ),
        lookups: as_u64(
            owner_count
                .checked_add(record_count)
                .and_then(|count| count.checked_mul(2))
                .and_then(|count| count.checked_add(1))
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(
            owner_count
                .checked_add(record_count)
                .and_then(|lookups| lookups.checked_add(caller_elements))
                .and_then(|work| work.checked_add(1))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    let (owners, records, overlay, engine) = authorized_phase(budget, caller_usage, || {
        let mut owners = Vec::new();
        owners
            .try_reserve_exact(owner_count)
            .map_err(|_error| allocation(owner_count))?;
        for route in &target.dependencies.formula_owners {
            owners.push(cache::DependencyPayload {
                identifier: message_object_identifier(source, *route, path)?,
                bytes: message_payload(source, *route, path)?,
            });
        }
        let mut records = Vec::new();
        records
            .try_reserve_exact(record_count)
            .map_err(|_error| allocation(record_count))?;
        for route in &target.dependencies.cell_record_tiles {
            records.push(cache::DependencyPayload {
                identifier: message_object_identifier(source, *route, path)?,
                bytes: message_payload(source, *route, path)?,
            });
        }
        let mut overlay = Vec::new();
        overlay
            .try_reserve_exact(changes.len())
            .map_err(|_error| allocation(changes.len()))?;
        for change in changes {
            let value = match change.input_ref() {
                None => cache::FinalValue::clear(),
                Some(Input::Number(value)) => cache::FinalValue::number(*value),
                Some(Input::Boolean(value)) => cache::FinalValue::boolean(*value),
                Some(Input::Text(_) | Input::Date(_) | Input::Duration(_)) => {
                    cache::FinalValue::unsupported()
                },
            };
            overlay.push(cache::FinalCell {
                table: identity,
                coordinate: cache::Coordinate {
                    row: change.position().row(),
                    column: change.position().column(),
                },
                value,
            });
        }
        exact_capacity(&owners, owner_count, path)?;
        exact_capacity(&records, record_count, path)?;
        exact_capacity(&overlay, changes.len(), path)?;
        let engine = message_payload(source, engine_route, path)?;
        Ok(((owners, records, overlay, engine), caller_usage))
    })?;
    let hosts = authorized_cache_phase(source, budget, path, |limits, remaining| {
        cache::collect_formula_hosts(identity, &owners, &records, limits)
            .map_err(|error| map_cache_error(error, path, remaining))
    })?;
    let (formula_cells, formula_entries, formula_payloads) =
        prepare_formula_cells(source, target, identity, &hosts, budget, path)?;
    let baseline_envelope = SemanticBaseline::envelope(table, path)?;
    let baseline = authorized_phase(budget, baseline_envelope, || {
        SemanticBaseline::build(table, path)
    })?;
    let mut evaluator = authorized_cache_phase(source, budget, path, |limits, remaining| {
        let (evaluator, usage) = cache::StrictEvaluator::new(
            &formula_entries,
            &formula_payloads,
            target.native.rows,
            target.native.columns,
            identity,
            limits,
        )
        .map_err(|error| map_cache_error(error, path, remaining))?;
        Ok((evaluator, usage))
    })?;
    let cache_source = cache::CacheSource {
        selected_table: identity,
        rows: target.native.rows,
        columns: target.native.columns,
        header_rows: u32::try_from(target.native.settings.header_row_count())
            .map_err(|_error| Error::InvalidSource { path })?,
        header_columns: u32::try_from(target.native.settings.header_column_count())
            .map_err(|_error| Error::InvalidSource { path })?,
        footer_rows: u32::try_from(target.native.settings.footer_row_count())
            .map_err(|_error| Error::InvalidSource { path })?,
        engine,
        owners: &owners,
        record_tiles: &records,
        formulas: &formula_cells,
    };
    let plan = authorized_cache_phase(source, budget, path, |limits, remaining| {
        let plan =
            cache::plan_final_cache(cache_source, &overlay, limits, &baseline, &mut evaluator)
                .map_err(|error| map_cache_error(error, path, remaining))?;
        let usage = plan.usage;
        Ok((plan, usage))
    })?;
    if !plan.removals.is_empty() {
        return Err(cache_refusal(path));
    }
    let refreshed_hosts = usize::try_from(plan.usage.cache_hosts_refreshed)
        .map_err(|_error| Error::InvalidSource { path })?;
    let materialization_envelope =
        materialization_envelope(target, &plan.rewrites, refreshed_hosts, path)?;
    let tiles = authorized_phase(budget, materialization_envelope, || {
        materialize_tile_changes(source, target, plan.rewrites, path)
    })?;
    Ok(PreparedCache {
        refreshed_hosts,
        tiles,
    })
}

struct SemanticBaseline {
    values: Vec<(cache::Coordinate, FormulaCachedValue)>,
    unsupported: Vec<cache::Coordinate>,
}

impl SemanticBaseline {
    fn envelope(table: &Table, path: Path) -> Result<budget::Usage, Error> {
        let cells = table.iter_cells().len();
        let elements = cells.checked_mul(2).ok_or(Error::InvalidSource { path })?;
        let bytes = cells
            .checked_mul(size_of::<(cache::Coordinate, FormulaCachedValue)>())
            .and_then(|bytes| {
                cells
                    .checked_mul(size_of::<cache::Coordinate>())
                    .and_then(|unsupported| bytes.checked_add(unsupported))
            })
            .ok_or(Error::InvalidSource { path })?;
        Ok(budget::Usage {
            retained_elements: as_u64(elements),
            retained_bytes: as_u64(bytes),
            allocation_events: if cells == 0 { 0 } else { 2 },
            transaction_work: as_u64(cells.checked_mul(2).ok_or(Error::InvalidSource { path })?),
            ..budget::Usage::default()
        })
    }

    fn build(table: &Table, path: Path) -> Result<(Self, budget::Usage), Error> {
        let cells = table.iter_cells();
        let cell_count = cells.len();
        let mut values = Vec::new();
        let mut unsupported = Vec::new();
        values
            .try_reserve_exact(cell_count)
            .map_err(|_error| allocation(cell_count))?;
        unsupported
            .try_reserve_exact(cell_count)
            .map_err(|_error| allocation(cell_count))?;
        exact_capacity(&values, cell_count, path)?;
        exact_capacity(&unsupported, cell_count, path)?;
        for cell in cells {
            let coordinate = cache::Coordinate {
                row: cell.position().row(),
                column: cell.position().column(),
            };
            match cell.value() {
                Value::Empty => {},
                Value::Number(value) => {
                    values.push((coordinate, FormulaCachedValue::Number(*value)))
                },
                Value::Boolean(value) => {
                    values.push((coordinate, FormulaCachedValue::Boolean(*value)))
                },
                Value::Text(_)
                | Value::Date(_)
                | Value::Duration(_)
                | Value::Formula(_)
                | Value::Error(_) => unsupported.push(coordinate),
            }
        }
        if values.windows(2).any(|pair| pair[0].0 >= pair[1].0)
            || unsupported.windows(2).any(|pair| pair[0] >= pair[1])
        {
            return Err(Error::InvalidSource { path });
        }
        // Both exactly reserved buffers remain live for every subsequent
        // cache phase. Charge their allocation shapes, including unused
        // slots, rather than only the populated semantic entries.
        let retained_elements = values
            .capacity()
            .checked_add(unsupported.capacity())
            .ok_or(Error::InvalidSource { path })?;
        let retained_bytes = values
            .capacity()
            .checked_mul(size_of::<(cache::Coordinate, FormulaCachedValue)>())
            .and_then(|bytes| {
                unsupported
                    .capacity()
                    .checked_mul(size_of::<cache::Coordinate>())
                    .and_then(|unsupported_bytes| bytes.checked_add(unsupported_bytes))
            })
            .ok_or(Error::InvalidSource { path })?;
        let validation_work = values
            .len()
            .saturating_sub(1)
            .checked_add(unsupported.len().saturating_sub(1))
            .ok_or(Error::InvalidSource { path })?;
        Ok((
            Self {
                values,
                unsupported,
            },
            budget::Usage {
                retained_elements: as_u64(retained_elements),
                retained_bytes: as_u64(retained_bytes),
                allocation_events: if cell_count == 0 { 0 } else { 2 },
                transaction_work: as_u64(
                    cell_count
                        .checked_add(validation_work)
                        .ok_or(Error::InvalidSource { path })?,
                ),
                ..budget::Usage::default()
            },
        ))
    }
}

impl cache::CacheBaseline for SemanticBaseline {
    fn value(
        &self,
        coordinate: cache::Coordinate,
    ) -> Result<Option<&FormulaCachedValue>, cache::Failure> {
        if self.unsupported.binary_search(&coordinate).is_ok() {
            return Err(cache::Failure::UnsupportedDependency(
                cache::Unsupported::Formula,
            ));
        }
        Ok(self
            .values
            .binary_search_by_key(&coordinate, |entry| entry.0)
            .ok()
            .and_then(|index| self.values.get(index))
            .map(|entry| &entry.1))
    }
}

fn prepare_formula_cells<'source>(
    source: &'source Package,
    target: &resolve::Target,
    identity: cache::TableIdentity,
    hosts: &[cache::FormulaHost],
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<
    (
        Vec<cache::FormulaCell>,
        Vec<cache::FormulaListEntry<'source>>,
        Vec<cache::FormulaPayload<'source>>,
    ),
    Error,
> {
    let formula_entries = collect_formula_entries(
        source,
        &target.storage.lists.formula,
        identity,
        budget,
        path,
    )?;
    let retained_bytes = hosts
        .len()
        .checked_mul(size_of::<cache::FormulaCell>())
        .and_then(|bytes| {
            hosts
                .len()
                .checked_mul(size_of::<cache::FormulaPayload<'_>>())
                .and_then(|payload_bytes| bytes.checked_add(payload_bytes))
        })
        .ok_or(Error::InvalidSource { path })?;
    let retained = budget::Usage {
        retained_elements: as_u64(
            hosts
                .len()
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        ),
        retained_bytes: as_u64(retained_bytes),
        allocation_events: if hosts.is_empty() { 0 } else { 2 },
        ..budget::Usage::default()
    };
    let (mut cells, mut payloads) = authorized_phase(budget, retained, || {
        let mut cells = Vec::new();
        let mut payloads = Vec::new();
        cells
            .try_reserve_exact(hosts.len())
            .map_err(|_error| allocation(hosts.len()))?;
        payloads
            .try_reserve_exact(hosts.len())
            .map_err(|_error| allocation(hosts.len()))?;
        exact_capacity(&cells, hosts.len(), path)?;
        exact_capacity(&payloads, hosts.len(), path)?;
        Ok(((cells, payloads), retained))
    })?;
    let mut start = 0usize;
    while start < hosts.len() {
        let tile_id = hosts[start].coordinate.row / target.storage.tile_size;
        let remaining_hosts = hosts.len() - start;
        let group_envelope = budget::Usage {
            transaction_work: as_u64(remaining_hosts),
            ..budget::Usage::default()
        };
        let end = authorized_phase(budget, group_envelope, || {
            let differing = hosts[start..]
                .iter()
                .position(|host| host.coordinate.row / target.storage.tile_size != tile_id);
            let end = differing.map_or(hosts.len(), |offset| start + offset);
            let work = differing.map_or(remaining_hosts, |offset| offset + 1);
            Ok((
                end,
                budget::Usage {
                    transaction_work: as_u64(work),
                    ..budget::Usage::default()
                },
            ))
        })?;
        let position_count = end - start;
        let position_work = binary_search_work(target.storage.tiles.len())
            .checked_add(position_count)
            .and_then(|work| work.checked_add(1))
            .ok_or(Error::InvalidSource { path })?;
        let position_usage = budget::Usage {
            peak_scratch_bytes: as_u64(
                position_count
                    .checked_mul(size_of::<tile::TileReadPosition>())
                    .ok_or(Error::InvalidSource { path })?,
            ),
            allocation_events: u64::from(position_count != 0),
            lookups: 2,
            transaction_work: as_u64(position_work),
            ..budget::Usage::default()
        };
        let (route, positions, tile_payload) = authorized_phase(budget, position_usage, || {
            let route = target
                .storage
                .tiles
                .binary_search_by_key(&tile_id, |route| route.tile_id)
                .ok()
                .and_then(|index| target.storage.tiles.get(index))
                .copied()
                .ok_or_else(|| cache_refusal(path))?;
            let mut positions = Vec::new();
            positions
                .try_reserve_exact(position_count)
                .map_err(|_error| allocation(position_count))?;
            for host in &hosts[start..end] {
                positions.push(tile::TileReadPosition {
                    row: host.coordinate.row % target.storage.tile_size,
                    column: host.coordinate.column,
                });
            }
            exact_capacity(&positions, position_count, path)?;
            let tile_payload = message_payload(source, route.message, path)?;
            Ok(((route, positions, tile_payload), position_usage))
        })?;
        let remaining = budget.remaining()?;
        let limits = cache_limits(source, remaining, path)?;
        let classified = authorized_phase(budget, remaining, || {
            let classified = tile::preclassify_tile(
                tile_payload,
                target.native.columns,
                &positions,
                tile::TileLimits::new(
                    limits.wire_bytes,
                    limits.retained_bytes.min(limits.wire_bytes),
                    limits.wire_fields,
                    u64::try_from(limits.wire_work)
                        .map_err(|_error| Error::InvalidSource { path })?,
                    usize::try_from(target.storage.tile_size)
                        .map_err(|_error| Error::InvalidSource { path })?,
                    limits.graph_nodes,
                ),
            )
            .map_err(|error| super::map_tile_error(error, path))?;
            let report = tile_report_usage(classified.report, path)?;
            Ok((classified, report))
        })?;
        let join_work = binary_search_work(formula_entries.len())
            .checked_add(3)
            .and_then(|per_host| per_host.checked_mul(position_count))
            .and_then(|work| work.checked_add(1))
            .ok_or(Error::InvalidSource { path })?;
        let join_usage = budget::Usage {
            lookups: as_u64(
                position_count
                    .checked_add(1)
                    .ok_or(Error::InvalidSource { path })?,
            ),
            transaction_work: as_u64(join_work),
            ..budget::Usage::default()
        };
        authorized_phase(budget, join_usage, || {
            let cache_object = message_object_identifier(source, route.message, path)?;
            if classified.cells.len() != position_count {
                return Err(Error::InvalidSource { path });
            }
            for (host, classified) in hosts[start..end].iter().zip(classified.cells) {
                let tile::CellValue::Formula { identifier, .. } = classified.before else {
                    return Err(cache_refusal(path));
                };
                let formula = formula_entries
                    .binary_search_by_key(&identifier, |entry| entry.key)
                    .ok()
                    .and_then(|index| formula_entries.get(index))
                    .ok_or(Error::InvalidSource { path })?;
                cells.push((*host).into_formula_cell(cache_object));
                payloads.push(cache::FormulaPayload {
                    owner: identity.owner,
                    coordinate: host.coordinate,
                    key: identifier,
                    bytes: formula.bytes,
                });
            }
            Ok(((), join_usage))
        })?;
        start = end;
    }
    Ok((cells, formula_entries, payloads))
}

fn collect_formula_entries<'source>(
    source: &'source Package,
    route: &resolve::ListRoute,
    identity: cache::TableIdentity,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<Vec<cache::FormulaListEntry<'source>>, Error> {
    let segment_count = route.segments.len();
    let segment_usage = budget::Usage {
        peak_scratch_bytes: as_u64(
            segment_count
                .checked_mul(size_of::<cache::FormulaListSegment<'_>>())
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(segment_count != 0),
        lookups: as_u64(
            segment_count
                .checked_mul(2)
                .and_then(|count| count.checked_add(1))
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(
            segment_count
                .checked_mul(2)
                .and_then(|work| work.checked_add(1))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    let (segments, root) = authorized_phase(budget, segment_usage, || {
        let mut segments = Vec::new();
        segments
            .try_reserve_exact(segment_count)
            .map_err(|_error| allocation(segment_count))?;
        for segment in &route.segments {
            segments.push(cache::FormulaListSegment {
                identifier: message_object_identifier(source, *segment, path)?,
                bytes: message_payload(source, *segment, path)?,
            });
        }
        exact_capacity(&segments, segment_count, path)?;
        let root = message_payload(source, route.message, path)?;
        Ok(((segments, root), segment_usage))
    })?;
    authorized_cache_phase(source, budget, path, |limits, remaining| {
        cache::collect_formula_list(identity, root, &segments, route.entries, limits)
            .map_err(|error| map_cache_error(error, path, remaining))
    })
}

fn materialization_envelope(
    target: &resolve::Target,
    rewrites: &[cache::CacheRewrite],
    changes: usize,
    path: Path,
) -> Result<budget::Usage, Error> {
    let retained_elements = rewrites
        .len()
        .checked_add(changes)
        .ok_or(Error::InvalidSource { path })?;
    let retained_bytes = rewrites
        .len()
        .checked_mul(size_of::<PreparedTileCache>())
        .and_then(|bytes| {
            changes
                .checked_mul(size_of::<tile::CacheChange>())
                .and_then(|change_bytes| bytes.checked_add(change_bytes))
        })
        .ok_or(Error::InvalidSource { path })?;
    let allocation_events = usize::from(!target.storage.tiles.is_empty())
        .checked_add(usize::from(!rewrites.is_empty()))
        .and_then(|events| events.checked_add(rewrites.len()))
        .ok_or(Error::InvalidSource { path })?;
    let scratch_bytes = target
        .storage
        .tiles
        .len()
        .checked_mul(size_of::<(u64, u32)>())
        .ok_or(Error::InvalidSource { path })?;
    let route_count = target.storage.tiles.len();
    let rewrite_count = rewrites.len();
    let transaction_work = route_count
        .checked_add(sort_work(route_count, path)?)
        .and_then(|work| work.checked_add(route_count))
        .and_then(|work| {
            binary_search_work(route_count)
                .checked_add(1)
                .and_then(|per_rewrite| per_rewrite.checked_mul(rewrite_count))
                .and_then(|rewrite_work| work.checked_add(rewrite_work))
        })
        .and_then(|work| work.checked_add(changes))
        .and_then(|work| work.checked_add(sort_work(rewrite_count, path).ok()?))
        .and_then(|work| work.checked_add(rewrite_count))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: as_u64(retained_elements),
        retained_bytes: as_u64(retained_bytes),
        peak_scratch_bytes: as_u64(scratch_bytes),
        allocation_events: as_u64(allocation_events),
        lookups: as_u64(
            route_count
                .checked_add(rewrite_count)
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn materialize_tile_changes(
    source: &Package,
    target: &resolve::Target,
    rewrites: Vec<cache::CacheRewrite>,
    path: Path,
) -> Result<(Vec<PreparedTileCache>, budget::Usage), Error> {
    let mut routes = Vec::new();
    routes
        .try_reserve_exact(target.storage.tiles.len())
        .map_err(|_error| allocation(target.storage.tiles.len()))?;
    exact_capacity(&routes, target.storage.tiles.len(), path)?;
    for route in &target.storage.tiles {
        routes.push((
            message_object_identifier(source, route.message, path)?,
            route.tile_id,
        ));
    }
    routes.sort_unstable();
    if routes.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(Error::InvalidSource { path });
    }
    let mut tiles = Vec::new();
    tiles
        .try_reserve_exact(rewrites.len())
        .map_err(|_error| allocation(rewrites.len()))?;
    exact_capacity(&tiles, rewrites.len(), path)?;
    for rewrite in rewrites {
        let tile_id = routes
            .binary_search_by_key(&rewrite.cache_object, |route| route.0)
            .ok()
            .and_then(|index| routes.get(index))
            .map(|route| route.1)
            .ok_or(Error::InvalidSource { path })?;
        let mut changes = Vec::new();
        changes
            .try_reserve_exact(rewrite.cells.len())
            .map_err(|_error| allocation(rewrite.cells.len()))?;
        exact_capacity(&changes, rewrite.cells.len(), path)?;
        for cell in rewrite.cells {
            if cell.coordinate.row / target.storage.tile_size != tile_id
                || !matches!(
                    cell.value,
                    FormulaCachedValue::Number(_) | FormulaCachedValue::Boolean(_)
                )
            {
                return Err(cache_refusal(path));
            }
            changes.push(tile::CacheChange {
                row: cell.coordinate.row % target.storage.tile_size,
                column: cell.coordinate.column,
                value: cell.value,
            });
        }
        tiles.push(PreparedTileCache { tile_id, changes });
    }
    tiles.sort_unstable_by_key(|tile| tile.tile_id);
    if tiles
        .windows(2)
        .any(|pair| pair[0].tile_id == pair[1].tile_id)
    {
        return Err(Error::InvalidSource { path });
    }
    let retained_elements = tiles
        .iter()
        .try_fold(tiles.len(), |count, tile| {
            count.checked_add(tile.changes.len())
        })
        .ok_or(Error::InvalidSource { path })?;
    let retained_bytes = tiles
        .capacity()
        .checked_mul(size_of::<PreparedTileCache>())
        .and_then(|bytes| {
            tiles.iter().try_fold(bytes, |total, tile| {
                tile.changes
                    .capacity()
                    .checked_mul(size_of::<tile::CacheChange>())
                    .and_then(|change_bytes| total.checked_add(change_bytes))
            })
        })
        .ok_or(Error::InvalidSource { path })?;
    let allocation_events = usize::from(!routes.is_empty())
        .checked_add(usize::from(!tiles.is_empty()))
        .and_then(|events| {
            events.checked_add(tiles.iter().filter(|tile| !tile.changes.is_empty()).count())
        })
        .ok_or(Error::InvalidSource { path })?;
    let scratch_bytes = routes
        .capacity()
        .checked_mul(size_of::<(u64, u32)>())
        .ok_or(Error::InvalidSource { path })?;
    let route_count = routes.len();
    let rewrite_count = tiles.len();
    let changes = retained_elements
        .checked_sub(rewrite_count)
        .ok_or(Error::InvalidSource { path })?;
    let transaction_work = route_count
        .checked_add(sort_work(route_count, path)?)
        .and_then(|work| work.checked_add(route_count))
        .and_then(|work| {
            binary_search_work(route_count)
                .checked_add(1)
                .and_then(|per_rewrite| per_rewrite.checked_mul(rewrite_count))
                .and_then(|rewrite_work| work.checked_add(rewrite_work))
        })
        .and_then(|work| work.checked_add(changes))
        .and_then(|work| work.checked_add(sort_work(rewrite_count, path).ok()?))
        .and_then(|work| work.checked_add(rewrite_count))
        .ok_or(Error::InvalidSource { path })?;
    Ok((
        tiles,
        budget::Usage {
            retained_elements: as_u64(retained_elements),
            retained_bytes: as_u64(retained_bytes),
            peak_scratch_bytes: as_u64(scratch_bytes),
            allocation_events: as_u64(allocation_events),
            lookups: as_u64(
                route_count
                    .checked_add(rewrite_count)
                    .ok_or(Error::InvalidSource { path })?,
            ),
            transaction_work: as_u64(transaction_work),
            ..budget::Usage::default()
        },
    ))
}

fn cache_limits(
    source: &Package,
    remaining: budget::Remaining,
    path: Path,
) -> Result<cache::CacheLimits, Error> {
    let archive = source.state.options.archive();
    let semantic = source.state.options.semantic();
    let wire = archive.max_iwa_stream_bytes();
    Ok(cache::CacheLimits {
        wire_bytes: bounded(remaining.wire_bytes, wire, path)?,
        wire_fields: bounded(remaining.wire_fields, wire, path)?,
        wire_work: bounded(
            remaining.wire_work,
            wire.checked_mul(32).ok_or(Error::InvalidSource { path })?,
            path,
        )?,
        wire_references: bounded(remaining.references, semantic.max_references(), path)?,
        wire_text: semantic.max_output_text_bytes(),
        nesting: 64,
        graph_nodes: bounded(
            remaining.formula_nodes,
            semantic.max_materialized_cells(),
            path,
        )?,
        graph_edges: bounded(remaining.formula_edges, semantic.max_references(), path)?,
        cache_cells: bounded(
            remaining.cache_hosts,
            semantic.max_materialized_cells(),
            path,
        )?,
        formula_work: bounded(
            remaining.formula_work,
            semantic.max_formula_render_work(),
            path,
        )?,
        graph_work: bounded(remaining.transaction_work, usize::MAX, path)?,
        retained_bytes: bounded(remaining.retained_bytes, usize::MAX, path)?,
        scratch_bytes: bounded(remaining.peak_scratch_bytes, usize::MAX, path)?,
        allocations: bounded(remaining.allocation_events, usize::MAX, path)?,
    })
}

fn cache_usage(usage: cache::CacheUsage, path: Path) -> Result<budget::Usage, Error> {
    let transaction_work = usage
        .wire_work
        .checked_add(usage.graph_work)
        .and_then(|work| work.checked_add(usage.formula_graph_builds))
        .and_then(|work| work.checked_add(usage.dependency_edges))
        .and_then(|work| work.checked_add(usage.cache_cells_read))
        .and_then(|work| work.checked_add(usage.cache_hosts_refreshed))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: usage.retained_elements,
        retained_bytes: usage.retained_bytes,
        peak_scratch_bytes: usage.peak_scratch_bytes,
        allocation_events: usage.allocations,
        wire_bytes: usage.wire_bytes,
        wire_fields: usage.wire_fields,
        wire_work: usage.wire_work,
        references: usage.wire_references,
        lookups: usage.lookup_work,
        formula_graph_builds: usage.formula_graph_builds,
        formula_nodes: usage.formula_nodes,
        formula_edges: usage.dependency_edges,
        range_candidates: usage.dependency_range_candidates,
        cache_hosts: usage.cache_hosts_refreshed,
        formula_work: usage.formula_work,
        transaction_work,
        ..budget::Usage::default()
    })
}

fn tile_report_usage(report: tile::TileReport, path: Path) -> Result<budget::Usage, Error> {
    let transaction_work = report
        .wire_work
        .checked_add(report.cell_slots_scanned)
        .and_then(|work| work.checked_add(report.cell_slots_written))
        .and_then(|work| work.checked_add(report.output_bytes))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: report.retained_elements,
        retained_bytes: report.retained_bytes,
        peak_scratch_bytes: report.peak_scratch_bytes,
        allocation_events: report.allocation_events,
        wire_bytes: report.wire_bytes,
        wire_fields: report.wire_fields,
        wire_work: report.wire_work,
        tile_reads: 1,
        row_reads: report.rows_read,
        cache_hosts: report.cache_cells_read,
        transaction_work,
        ..budget::Usage::default()
    })
}

fn map_cache_error(error: cache::Failure, path: Path, remaining: budget::Remaining) -> Error {
    match error {
        cache::Failure::InvalidSource => Error::InvalidSource { path },
        cache::Failure::UnsupportedDependency(cache::Unsupported::HeaderNameManager) => {
            Error::UnsupportedDependency {
                path,
                kind: DependencyKind::HeaderNameIndex,
            }
        },
        cache::Failure::UnsupportedDependency(_) => cache_refusal(path),
        cache::Failure::LimitExceeded { observed, .. } => Error::LimitExceeded {
            kind: LimitKind::FormulaWork,
            observed,
            maximum: remaining.formula_work,
            path,
        },
        cache::Failure::Allocation { amount } => Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount,
        },
    }
}

fn bounded(value: u64, maximum: usize, path: Path) -> Result<usize, Error> {
    usize::try_from(value)
        .map(|value| value.min(maximum))
        .map_err(|_error| Error::InvalidSource { path })
}

fn allocation(amount: usize) -> Error {
    Error::Allocation {
        kind: LimitKind::RetainedElements,
        amount,
    }
}

fn exact_capacity<T>(values: &Vec<T>, expected: usize, path: Path) -> Result<(), Error> {
    if size_of::<T>() == 0 || values.capacity() == expected {
        Ok(())
    } else {
        Err(Error::InvalidSource { path })
    }
}

fn binary_search_work(length: usize) -> usize {
    if length <= 1 {
        1
    } else {
        usize::try_from(usize::BITS - (length - 1).leading_zeros()).unwrap_or(usize::MAX)
    }
}

fn sort_work(length: usize, path: Path) -> Result<usize, Error> {
    length
        .checked_mul(binary_search_work(length))
        .ok_or(Error::InvalidSource { path })
}

fn cache_refusal(path: Path) -> Error {
    Error::UnsupportedDependency {
        path,
        kind: DependencyKind::FormulaCache,
    }
}

const fn as_u64(value: usize) -> u64 {
    value as u64
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;

    use crate::table::{CellPosition, Dimensions};

    use super::*;

    fn limits(max_retained_elements: u64) -> budget::TransactionLimits {
        budget::TransactionLimits {
            max_updates: 100,
            max_owned_value_bytes: 100,
            max_retained_elements,
            max_retained_bytes: 100_000,
            max_scratch_bytes: 100_000,
            max_allocation_events: 100,
            max_wire_bytes: 100_000,
            max_wire_fields: 100_000,
            max_wire_work: 100_000,
            max_objects: 100,
            max_references: 100_000,
            max_formula_work: 100_000,
            max_output_bytes: 100_000,
            max_reopen_work: 100_000,
            max_transaction_work: 100_000,
        }
    }

    #[test]
    fn formula_baseline_retains_both_capacities_and_next_phase_refuses_before_callback() {
        let mut builder = Table::builder("Formula", Dimensions::new(1, 2));
        builder
            .set(
                CellPosition::new(0, 0),
                Value::number(1.0).expect("finite baseline number"),
            )
            .expect("number cell is in bounds");
        builder
            .set(CellPosition::new(0, 1), Value::Formula("=A1".to_owned()))
            .expect("formula cell is in bounds");
        let table = builder.finish().expect("formula table is valid");
        let path = Path::Table { sheet: 0, table: 0 };
        let mut transaction =
            budget::TransactionBudget::from_limits(limits(4)).expect("finite limits");

        let baseline = authorized_phase(
            &mut transaction,
            SemanticBaseline::envelope(&table, path).expect("baseline envelope"),
            || SemanticBaseline::build(&table, path),
        )
        .expect("baseline fits exactly");
        assert_eq!(baseline.values.capacity(), 2);
        assert_eq!(baseline.unsupported.capacity(), 2);

        let before = transaction.remaining().expect("remaining budget");
        assert_eq!(before.retained_elements, 0);
        let callbacks = Cell::new(0usize);
        let result = authorized_phase(
            &mut transaction,
            budget::Usage {
                retained_elements: 1,
                retained_bytes: 1,
                allocation_events: 1,
                transaction_work: 1,
                ..budget::Usage::default()
            },
            || {
                callbacks.set(callbacks.get() + 1);
                Ok(((), budget::Usage::default()))
            },
        );

        assert!(matches!(
            result,
            Err(Error::LimitExceeded {
                kind: LimitKind::RetainedElements,
                observed: 5,
                maximum: 4,
                ..
            })
        ));
        assert_eq!(callbacks.get(), 0, "refused phase must not run");
        assert_eq!(
            transaction.remaining().expect("unchanged remaining budget"),
            before,
            "refusal must record no work or allocation"
        );
        assert!(!transaction.authorization_is_pending());
        drop(baseline);
    }
}
