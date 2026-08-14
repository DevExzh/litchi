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

use super::{
    budget, cache, formula_author, formula_list, message_object_identifier, message_payload,
    resolve, tile,
};

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

/// Logical post-authoring cache values. Text remains owned content here and
/// must be interned by the physical formula planner before tile publication.
#[derive(Debug)]
pub(super) struct PreparedAuthoredCache {
    pub(super) refreshed_hosts: usize,
    pub(super) rewrites: Vec<cache::CacheRewrite>,
}

/// One read-only semantic table admitted as a lazy external evaluator
/// baseline. Geometry remains in the cache-owned table registry; this route
/// supplies values only if an evaluated formula actually references `owner`.
#[derive(Debug, Clone, Copy)]
pub(super) struct ExternalBaselineTable<'source> {
    pub(super) owner: u32,
    pub(super) table: &'source Table,
}

/// Strict source formula index reusable by no-op detection and final cache
/// planning without reparsing formula lists or tile hosts.
pub(super) struct ExistingFormulaIndex<'source> {
    pub(super) identity: cache::TableIdentity,
    pub(super) cells: Vec<cache::FormulaCell>,
    pub(super) entries: Vec<cache::FormulaListEntry<'source>>,
    pub(super) payloads: Vec<cache::FormulaPayload<'source>>,
    pub(super) formula_list: PreparedFormulaListSource<'source>,
    pub(super) caches: Vec<ExistingFormulaCache>,
}

/// Complete selected-table formula state after applying one logical list
/// overlay. Unlike the authored batch, these vectors retain every untouched
/// formula which survives the transaction.
pub(super) struct PreparedFinalFormulaSet<'formula, 'cache> {
    pub(super) cells: Vec<cache::FormulaCell>,
    pub(super) entries: Vec<cache::FormulaListEntry<'formula>>,
    pub(super) payloads: Vec<cache::FormulaPayload<'formula>>,
    pub(super) authored: Vec<cache::AuthoredFormulaCache<'cache>>,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) struct ExistingFormulaCache {
    pub(super) owner: u32,
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) value: Option<tile::FormulaCacheValue>,
    pub(super) formula_error: Option<u32>,
}

#[derive(Debug)]
struct PreparedFormulaTileScan {
    tile_id: u32,
    cache_object: u64,
    scan: tile::FormulaCellScan,
}

/// One exact source message used by the physical formula-list planner.
#[derive(Debug, Clone, Copy)]
pub(super) struct PreparedFormulaListMessage<'source> {
    pub(super) object_id: u64,
    pub(super) payload: &'source [u8],
    /// Exact aggregate `MessageInfo.object_references` in source order.
    pub(super) object_references: &'source [u64],
}

/// An authoritative existing formula host joined to its strict list key.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct PreparedFormulaListHost {
    pub(super) owner: u32,
    pub(super) row: u32,
    pub(super) column: u32,
    pub(super) key: u32,
}

/// Budget-settled source route for one already-resolved involved table.
///
/// Callers flatten one value per resolver-derived `Target`; external targets
/// remain read-only, but their exact lists and hosts participate in global
/// formula/dependency proof.
#[derive(Debug)]
pub(super) struct PreparedFormulaListSource<'source> {
    pub(super) root: PreparedFormulaListMessage<'source>,
    pub(super) segments: Vec<PreparedFormulaListMessage<'source>>,
    pub(super) expected_entries: usize,
    pub(super) hosts: Vec<PreparedFormulaListHost>,
}

pub(super) fn prepare_existing_formula_index<'source>(
    source: &'source Package,
    target: &resolve::Target,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<ExistingFormulaIndex<'source>, Error> {
    let selected = target
        .dependencies
        .selected_formula_owner
        .ok_or_else(|| cache_refusal(path))?;
    let identity = cache::TableIdentity {
        owner: selected.internal_owner_id,
        uuid_lower: selected.uid_lower,
        uuid_upper: selected.uid_upper,
    };
    let (owners, records) = prepare_dependency_payloads(source, target, budget, path)?;
    let dependency_hosts = authorized_cache_phase(source, budget, path, |limits, remaining| {
        cache::collect_formula_hosts(identity, &owners, &records, limits)
            .map_err(|error| map_cache_error(error, path, remaining))
    })?;
    let (cells, entries, payloads, caches) =
        prepare_complete_formula_cells(source, target, identity, budget, path)?;
    let subset_work = dependency_hosts
        .len()
        .checked_mul(
            binary_search_work(cells.len())
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    authorized_phase(
        budget,
        budget::Usage {
            lookups: as_u64(dependency_hosts.len()),
            transaction_work: as_u64(subset_work),
            ..budget::Usage::default()
        },
        || {
            if dependency_hosts.iter().any(|host| {
                cells
                    .binary_search_by_key(
                        &(host.owner, host.coordinate.row, host.coordinate.column),
                        |cell| (cell.owner, cell.coordinate.row, cell.coordinate.column),
                    )
                    .is_err()
            }) {
                return Err(cache_refusal(path));
            }
            Ok((
                (),
                budget::Usage {
                    lookups: as_u64(dependency_hosts.len()),
                    transaction_work: as_u64(subset_work),
                    ..budget::Usage::default()
                },
            ))
        },
    )?;
    let table = [cache::TableGeometry {
        identity,
        rows: target.native.rows,
        columns: target.native.columns,
        header_rows: u32::try_from(target.native.settings.header_row_count())
            .map_err(|_| Error::InvalidSource { path })?,
        header_columns: u32::try_from(target.native.settings.header_column_count())
            .map_err(|_| Error::InvalidSource { path })?,
        footer_rows: u32::try_from(target.native.settings.footer_row_count())
            .map_err(|_| Error::InvalidSource { path })?,
    }];
    authorized_cache_phase(source, budget, path, |limits, remaining| {
        let usage = cache::validate_formula_coverage(&entries, &payloads, &table, limits)
            .map_err(|error| map_cache_error(error, path, remaining))?;
        Ok(((), usage))
    })?;
    let formula_list = prepare_formula_list_source(
        source,
        &target.storage.lists.formula,
        identity,
        &cells,
        &payloads,
        budget,
        path,
    )?;
    let index = ExistingFormulaIndex {
        identity,
        cells,
        entries,
        payloads,
        formula_list,
        caches,
    };
    if index.formula_list.expected_entries != index.entries.len() {
        return Err(Error::InvalidSource { path });
    }
    Ok(index)
}

/// Merge the authoritative source BNC index with the already-validated
/// formula-list assignments. Planning is a separate admitted linear pass so
/// no attacker-sized validation occurs before the final vector envelope is
/// known.
pub(super) fn prepare_complete_final_formula_set<'formula, 'cache>(
    source: &Package,
    target: &resolve::Target,
    existing: &ExistingFormulaIndex<'_>,
    logical: formula_list::LogicalList<'formula>,
    authored: &formula_author::PreparedFormulaBatch,
    changes: &'cache [Change],
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<PreparedFinalFormulaSet<'formula, 'cache>, Error> {
    #[derive(Clone, Copy)]
    struct MergePlan {
        final_cells: usize,
        additions: usize,
    }

    let validation_work = existing
        .cells
        .len()
        .checked_add(existing.payloads.len())
        .and_then(|work| work.checked_add(logical.assignments.len()))
        .and_then(|work| work.checked_add(logical.entries.len()))
        .and_then(|work| work.checked_add(authored.formulas.len()))
        .and_then(|work| {
            authored
                .formulas
                .len()
                .checked_mul(
                    binary_search_work(logical.assignments.len())
                        .checked_add(binary_search_work(logical.entries.len()))?
                        .checked_add(2)?,
                )
                .and_then(|authored_work| work.checked_add(authored_work))
        })
        .ok_or(Error::InvalidSource { path })?;
    let validation_usage = budget::Usage {
        lookups: as_u64(
            authored
                .formulas
                .len()
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(validation_work),
        ..budget::Usage::default()
    };
    let plan = authorized_phase(budget, validation_usage, || {
        if existing.cells.len() != existing.payloads.len()
            || existing
                .cells
                .iter()
                .zip(&existing.payloads)
                .any(|(cell, payload)| {
                    cell.owner != existing.identity.owner
                        || payload.owner != existing.identity.owner
                        || cell.coordinate != payload.coordinate
                        || cell.cache_object == 0
                        || payload.key == 0
                })
            || existing
                .cells
                .windows(2)
                .any(|pair| pair[0].coordinate >= pair[1].coordinate)
            || logical
                .assignments
                .windows(2)
                .any(|pair| (pair[0].row, pair[0].column) >= (pair[1].row, pair[1].column))
            || logical
                .entries
                .windows(2)
                .any(|pair| pair[0].key >= pair[1].key)
            || logical
                .entries
                .iter()
                .any(|entry| entry.key == 0 || entry.ref_count == 0 || entry.formula.is_empty())
            || authored.formulas.windows(2).any(|pair| {
                (pair[0].position.row(), pair[0].position.column())
                    >= (pair[1].position.row(), pair[1].position.column())
            })
        {
            return Err(Error::Verification { path });
        }
        for formula in &authored.formulas {
            let coordinate = (formula.position.row(), formula.position.column());
            let assignment = logical
                .assignments
                .binary_search_by_key(&coordinate, |assignment| {
                    (assignment.row, assignment.column)
                })
                .ok()
                .and_then(|index| logical.assignments.get(index))
                .ok_or(Error::Verification { path })?;
            let key = assignment.key.ok_or(Error::Verification { path })?;
            let entry = logical
                .entries
                .binary_search_by_key(&key, |entry| entry.key)
                .ok()
                .and_then(|index| logical.entries.get(index))
                .ok_or(Error::Verification { path })?;
            if entry.formula != formula.bytes.as_slice() {
                return Err(Error::Verification { path });
            }
        }
        let mut existing_index = 0usize;
        let mut assignment_index = 0usize;
        let mut final_cells = 0usize;
        let mut additions = 0usize;
        while existing_index < existing.cells.len() || assignment_index < logical.assignments.len()
        {
            let existing_coordinate = existing
                .cells
                .get(existing_index)
                .map(|cell| (cell.coordinate.row, cell.coordinate.column));
            let assignment = logical.assignments.get(assignment_index);
            let assignment_coordinate = assignment.map(|item| (item.row, item.column));
            match (existing_coordinate, assignment_coordinate) {
                (Some(left), Some(right)) if left < right => {
                    final_cells = final_cells
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                    existing_index += 1;
                },
                (Some(left), Some(right)) if left == right => {
                    if assignment.is_some_and(|item| item.key.is_some()) {
                        final_cells = final_cells
                            .checked_add(1)
                            .ok_or(Error::InvalidSource { path })?;
                    }
                    existing_index += 1;
                    assignment_index += 1;
                },
                (_, Some(_)) => {
                    if assignment.is_none_or(|item| item.key.is_none()) {
                        return Err(Error::Verification { path });
                    }
                    final_cells = final_cells
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                    additions = additions
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                    assignment_index += 1;
                },
                (Some(_), None) => {
                    final_cells = final_cells
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                    existing_index += 1;
                },
                (None, None) => break,
            }
        }
        Ok((
            MergePlan {
                final_cells,
                additions,
            },
            validation_usage,
        ))
    })?;

    let retained_elements = plan
        .final_cells
        .checked_mul(2)
        .and_then(|count| count.checked_add(logical.entries.len()))
        .and_then(|count| count.checked_add(authored.formulas.len()))
        .ok_or(Error::InvalidSource { path })?;
    let retained_bytes = plan
        .final_cells
        .checked_mul(size_of::<cache::FormulaCell>())
        .and_then(|bytes| {
            plan.final_cells
                .checked_mul(size_of::<cache::FormulaPayload<'_>>())
                .and_then(|payload_bytes| bytes.checked_add(payload_bytes))
        })
        .and_then(|bytes| {
            logical
                .entries
                .len()
                .checked_mul(size_of::<cache::FormulaListEntry<'_>>())
                .and_then(|entry_bytes| bytes.checked_add(entry_bytes))
        })
        .and_then(|bytes| {
            authored
                .formulas
                .len()
                .checked_mul(size_of::<cache::AuthoredFormulaCache<'_>>())
                .and_then(|authored_bytes| bytes.checked_add(authored_bytes))
        })
        .ok_or(Error::InvalidSource { path })?;
    let build_work = existing
        .cells
        .len()
        .checked_add(logical.assignments.len())
        .and_then(|work| work.checked_add(logical.entries.len()))
        .and_then(|work| {
            plan.final_cells
                .checked_mul(binary_search_work(logical.entries.len()).checked_add(2)?)
                .and_then(|join| work.checked_add(join))
        })
        .and_then(|work| {
            plan.additions
                .checked_mul(binary_search_work(target.storage.tiles.len()).checked_add(2)?)
                .and_then(|join| work.checked_add(join))
        })
        .and_then(|work| {
            authored
                .formulas
                .len()
                .checked_mul(2)
                .and_then(|authored_work| work.checked_add(authored_work))
        })
        .ok_or(Error::InvalidSource { path })?;
    let allocation_events = usize::from(plan.final_cells != 0)
        .checked_mul(2)
        .and_then(|events| events.checked_add(usize::from(!logical.entries.is_empty())))
        .and_then(|events| events.checked_add(usize::from(!authored.formulas.is_empty())))
        .ok_or(Error::InvalidSource { path })?;
    let build_usage = budget::Usage {
        retained_elements: as_u64(retained_elements),
        retained_bytes: as_u64(retained_bytes),
        allocation_events: as_u64(allocation_events),
        lookups: as_u64(
            plan.final_cells
                .checked_add(plan.additions)
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(build_work),
        ..budget::Usage::default()
    };
    authorized_phase(budget, build_usage, || {
        let mut cells = Vec::new();
        let mut entries = Vec::new();
        let mut payloads = Vec::new();
        let mut authored_cache = Vec::new();
        cells
            .try_reserve_exact(plan.final_cells)
            .map_err(|_| allocation(plan.final_cells))?;
        payloads
            .try_reserve_exact(plan.final_cells)
            .map_err(|_| allocation(plan.final_cells))?;
        entries
            .try_reserve_exact(logical.entries.len())
            .map_err(|_| allocation(logical.entries.len()))?;
        authored_cache
            .try_reserve_exact(authored.formulas.len())
            .map_err(|_| allocation(authored.formulas.len()))?;
        for entry in logical.entries {
            entries.push(cache::FormulaListEntry {
                owner: existing.identity.owner,
                key: entry.key,
                ref_count: entry.ref_count,
                bytes: entry.formula,
            });
        }
        let mut existing_index = 0usize;
        let mut assignment_index = 0usize;
        while existing_index < existing.cells.len() || assignment_index < logical.assignments.len()
        {
            let source_cell = existing.cells.get(existing_index);
            let source_payload = existing.payloads.get(existing_index);
            let source_coordinate = source_cell.map(|cell| cell.coordinate);
            let assignment = logical.assignments.get(assignment_index);
            let assignment_coordinate = assignment.map(|item| cache::Coordinate {
                row: item.row,
                column: item.column,
            });
            let (coordinate, key, cache_object, consume_source, consume_assignment) =
                match (source_coordinate, assignment_coordinate) {
                    (Some(left), Some(right)) if left < right => (
                        left,
                        source_payload.ok_or(Error::Verification { path })?.key,
                        source_cell
                            .ok_or(Error::Verification { path })?
                            .cache_object,
                        true,
                        false,
                    ),
                    (Some(left), Some(right)) if left == right => {
                        let key = assignment.and_then(|item| item.key).unwrap_or_default();
                        if key == 0 {
                            existing_index += 1;
                            assignment_index += 1;
                            continue;
                        }
                        (
                            left,
                            key,
                            source_cell
                                .ok_or(Error::Verification { path })?
                                .cache_object,
                            true,
                            true,
                        )
                    },
                    (_, Some(right)) => {
                        let key = assignment
                            .and_then(|item| item.key)
                            .ok_or(Error::Verification { path })?;
                        let tile_id = right.row / target.storage.tile_size;
                        let route = target
                            .storage
                            .tiles
                            .binary_search_by_key(&tile_id, |route| route.tile_id)
                            .ok()
                            .and_then(|index| target.storage.tiles.get(index))
                            .ok_or(Error::Verification { path })?;
                        (
                            right,
                            key,
                            message_object_identifier(source, route.message, path)?,
                            false,
                            true,
                        )
                    },
                    (Some(left), None) => (
                        left,
                        source_payload.ok_or(Error::Verification { path })?.key,
                        source_cell
                            .ok_or(Error::Verification { path })?
                            .cache_object,
                        true,
                        false,
                    ),
                    (None, None) => break,
                };
            let entry = entries
                .binary_search_by_key(&key, |entry| entry.key)
                .ok()
                .and_then(|index| entries.get(index))
                .ok_or(Error::Verification { path })?;
            cells.push(cache::FormulaCell {
                owner: existing.identity.owner,
                coordinate,
                cache_object,
            });
            payloads.push(cache::FormulaPayload {
                owner: existing.identity.owner,
                coordinate,
                key,
                bytes: entry.bytes,
            });
            existing_index += usize::from(consume_source);
            assignment_index += usize::from(consume_assignment);
        }
        for formula in &authored.formulas {
            let change = changes
                .get(formula.change_index)
                .ok_or(Error::Verification { path })?;
            let Some(Input::Formula { cached, .. }) = change.input_ref() else {
                return Err(Error::Verification { path });
            };
            authored_cache.push(cache::AuthoredFormulaCache {
                owner: existing.identity.owner,
                coordinate: cache::Coordinate {
                    row: formula.position.row(),
                    column: formula.position.column(),
                },
                supplied: cached.as_ref(),
            });
        }
        exact_capacity(&cells, plan.final_cells, path)?;
        exact_capacity(&payloads, plan.final_cells, path)?;
        exact_capacity(&entries, logical.entries.len(), path)?;
        exact_capacity(&authored_cache, authored.formulas.len(), path)?;
        if cells.len() != plan.final_cells
            || payloads.len() != plan.final_cells
            || entries.len() != logical.entries.len()
            || authored_cache.len() != authored.formulas.len()
        {
            return Err(Error::Verification { path });
        }
        Ok((
            PreparedFinalFormulaSet {
                cells,
                entries,
                payloads,
                authored: authored_cache,
                retained_elements,
                retained_bytes,
            },
            build_usage,
        ))
    })
}

fn validate_prepared_formula_list_source(
    source: &PreparedFormulaListSource<'_>,
    identity: cache::TableIdentity,
    entries: usize,
    path: Path,
) -> Result<(), Error> {
    if source.root.object_id == 0
        || source.root.payload.is_empty()
        || source.expected_entries != entries
        || source.root.object_references.len() != source.segments.len()
        || source
            .root
            .object_references
            .iter()
            .zip(&source.segments)
            .any(|(reference, segment)| *reference != segment.object_id)
        || source
            .segments
            .iter()
            .any(|segment| segment.payload.is_empty() || !segment.object_references.is_empty())
        || source
            .hosts
            .iter()
            .any(|host| host.owner != identity.owner || host.key == 0)
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "strict source route, host join, and transaction authority remain explicit"
)]
fn prepare_formula_list_source<'source>(
    source: &'source Package,
    route: &resolve::ListRoute,
    identity: cache::TableIdentity,
    cells: &[cache::FormulaCell],
    payloads: &[cache::FormulaPayload<'source>],
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<PreparedFormulaListSource<'source>, Error> {
    if cells.len() != payloads.len() {
        return Err(Error::InvalidSource { path });
    }
    let segment_count = route.segments.len();
    let host_count = cells.len();
    let elements = segment_count
        .checked_add(host_count)
        .ok_or(Error::InvalidSource { path })?;
    let retained_bytes = segment_count
        .checked_mul(size_of::<PreparedFormulaListMessage<'_>>())
        .and_then(|bytes| {
            host_count
                .checked_mul(size_of::<PreparedFormulaListHost>())
                .and_then(|host_bytes| bytes.checked_add(host_bytes))
        })
        .ok_or(Error::InvalidSource { path })?;
    let message_count = segment_count
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    let usage = budget::Usage {
        retained_elements: as_u64(elements),
        retained_bytes: as_u64(retained_bytes),
        allocation_events: u64::from(segment_count != 0) + u64::from(host_count != 0),
        lookups: as_u64(
            message_count
                .checked_mul(4)
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(
            message_count
                .checked_mul(4)
                .and_then(|work| {
                    segment_count
                        .checked_mul(3)
                        .and_then(|segment_work| work.checked_add(segment_work))
                })
                .and_then(|work| {
                    host_count
                        .checked_mul(3)
                        .and_then(|host_work| work.checked_add(host_work))
                })
                .and_then(|work| work.checked_add(1))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    authorized_phase(budget, usage, || {
        let root = formula_list_source_message(source, route.message, path)?;
        let mut segments = Vec::new();
        segments
            .try_reserve_exact(segment_count)
            .map_err(|_error| allocation(segment_count))?;
        for segment in &route.segments {
            segments.push(formula_list_source_message(source, *segment, path)?);
        }
        exact_capacity(&segments, segment_count, path)?;

        let mut hosts = Vec::new();
        hosts
            .try_reserve_exact(host_count)
            .map_err(|_error| allocation(host_count))?;
        for (cell, payload) in cells.iter().zip(payloads) {
            if cell.owner != identity.owner
                || payload.owner != identity.owner
                || cell.coordinate != payload.coordinate
            {
                return Err(Error::InvalidSource { path });
            }
            hosts.push(PreparedFormulaListHost {
                owner: identity.owner,
                row: cell.coordinate.row,
                column: cell.coordinate.column,
                key: payload.key,
            });
        }
        exact_capacity(&hosts, host_count, path)?;
        if hosts.windows(2).any(|pair| pair[0] >= pair[1]) {
            return Err(Error::InvalidSource { path });
        }
        let prepared = PreparedFormulaListSource {
            root,
            segments,
            expected_entries: route.entries,
            hosts,
        };
        validate_prepared_formula_list_source(&prepared, identity, route.entries, path)?;
        Ok((prepared, usage))
    })
}

fn formula_list_source_message(
    source: &Package,
    route: resolve::MessageRoute,
    path: Path,
) -> Result<PreparedFormulaListMessage<'_>, Error> {
    let object = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .ok_or(Error::InvalidSource { path })?;
    let message = object
        .messages
        .get(route.message_index)
        .filter(|message| message.type_ == route.message_type)
        .ok_or(Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .get(route.message_index)
        .filter(|info| info.type_ == route.message_type && info.field_infos.is_empty())
        .ok_or(Error::InvalidSource { path })?;
    let object_id = object
        .archive_info
        .identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource { path })?;
    Ok(PreparedFormulaListMessage {
        object_id,
        payload: message.data.as_slice(),
        object_references: info.object_references.as_slice(),
    })
}

pub(super) fn prepare_dependency_payloads<'source>(
    source: &'source Package,
    target: &resolve::Target,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<
    (
        Vec<cache::DependencyPayload<'source>>,
        Vec<cache::DependencyPayload<'source>>,
    ),
    Error,
> {
    let owner_count = target.dependencies.formula_owners.len();
    let record_count = target
        .dependencies
        .formula_owners
        .iter()
        .try_fold(0usize, |count, owner| {
            count.checked_add(owner.cell_record_tiles.len())
        })
        .and_then(|count| count.checked_add(target.dependencies.inert_marker_tiles.len()))
        .ok_or(Error::InvalidSource { path })?;
    let record_sort_work = sort_work(record_count, path)?;
    let usage = budget::Usage {
        retained_elements: as_u64(
            owner_count
                .checked_add(record_count)
                .ok_or(Error::InvalidSource { path })?,
        ),
        retained_bytes: as_u64(
            owner_count
                .checked_add(record_count)
                .and_then(|count| count.checked_mul(size_of::<cache::DependencyPayload<'_>>()))
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(owner_count != 0) + u64::from(record_count != 0),
        lookups: as_u64(
            owner_count
                .checked_add(record_count)
                .and_then(|count| count.checked_mul(2))
                .ok_or(Error::InvalidSource { path })?,
        ),
        transaction_work: as_u64(
            owner_count
                .checked_add(record_count)
                .and_then(|work| work.checked_add(record_sort_work))
                .and_then(|work| work.checked_add(record_count))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    authorized_phase(budget, usage, || {
        let mut owners = Vec::new();
        let mut records = Vec::new();
        owners
            .try_reserve_exact(owner_count)
            .map_err(|_error| allocation(owner_count))?;
        records
            .try_reserve_exact(record_count)
            .map_err(|_error| allocation(record_count))?;
        for route in &target.dependencies.formula_owners {
            owners.push(cache::DependencyPayload {
                identifier: message_object_identifier(source, route.message, path)?,
                bytes: message_payload(source, route.message, path)?,
            });
        }
        for owner in &target.dependencies.formula_owners {
            for route in &owner.cell_record_tiles {
                records.push(cache::DependencyPayload {
                    identifier: message_object_identifier(source, *route, path)?,
                    bytes: message_payload(source, *route, path)?,
                });
            }
        }
        for route in &target.dependencies.inert_marker_tiles {
            records.push(cache::DependencyPayload {
                identifier: message_object_identifier(source, *route, path)?,
                bytes: message_payload(source, *route, path)?,
            });
        }
        records.sort_unstable_by_key(|payload| payload.identifier);
        if records
            .windows(2)
            .any(|pair| pair[0].identifier == pair[1].identifier && pair[0].bytes != pair[1].bytes)
        {
            return Err(Error::InvalidSource { path });
        }
        records.dedup_by_key(|payload| payload.identifier);
        exact_capacity(&owners, owner_count, path)?;
        exact_capacity(&records, record_count, path)?;
        Ok(((owners, records), usage))
    })
}

/// Collect every source range-dependency tile in the calculation engine.
///
/// Range records may belong to read-only external owners.  They participate
/// in cache validation even though this transaction only rewrites the
/// selected owner's dependency messages.
pub(super) fn prepare_range_dependency_payloads<'source>(
    source: &'source Package,
    target: &resolve::Target,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<Vec<cache::DependencyPayload<'source>>, Error> {
    let count = target
        .dependencies
        .formula_owners
        .iter()
        .try_fold(0usize, |count, owner| {
            count.checked_add(owner.range_precedent_tiles.len())
        })
        .ok_or(Error::InvalidSource { path })?;
    let sort = sort_work(count, path)?;
    let usage = budget::Usage {
        retained_elements: as_u64(count),
        retained_bytes: as_u64(
            count
                .checked_mul(size_of::<cache::DependencyPayload<'_>>())
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(count != 0),
        lookups: as_u64(count.checked_mul(2).ok_or(Error::InvalidSource { path })?),
        transaction_work: as_u64(
            count
                .checked_add(sort)
                .and_then(|work| work.checked_add(count))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    authorized_phase(budget, usage, || {
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(count)
            .map_err(|_error| allocation(count))?;
        exact_capacity(&ranges, count, path)?;
        for owner in &target.dependencies.formula_owners {
            for route in &owner.range_precedent_tiles {
                ranges.push(cache::DependencyPayload {
                    identifier: message_object_identifier(source, route.message, path)?,
                    bytes: message_payload(source, route.message, path)?,
                });
            }
        }
        ranges.sort_unstable_by_key(|payload| payload.identifier);
        if ranges
            .windows(2)
            .any(|pair| pair[0].identifier == pair[1].identifier && pair[0].bytes != pair[1].bytes)
        {
            return Err(Error::InvalidSource { path });
        }
        ranges.dedup_by_key(|payload| payload.identifier);
        Ok((ranges, usage))
    })
}

/// Complete final dependency payloads prepared by the physical formula
/// transaction. These bytes must describe the post-edit graph, never an
/// intermediate sequential state.
#[derive(Debug, Clone, Copy)]
pub(super) struct FinalDependencySet<'source> {
    pub(super) engine: &'source [u8],
    pub(super) owners: &'source [cache::DependencyPayload<'source>],
    pub(super) record_tiles: &'source [cache::DependencyPayload<'source>],
    pub(super) range_tiles: &'source [cache::DependencyPayload<'source>],
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
    tables: &[cache::TableGeometry],
    baseline: &[ExternalBaselineTable<'_>],
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
    let record_count = target
        .dependencies
        .formula_owners
        .iter()
        .try_fold(0usize, |count, owner| {
            count.checked_add(owner.cell_record_tiles.len())
        })
        .and_then(|count| count.checked_add(target.dependencies.inert_marker_tiles.len()))
        .ok_or(Error::InvalidSource { path })?;
    let record_sort_work = sort_work(record_count, path)?;
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
                .and_then(|work| work.checked_add(record_sort_work))
                .and_then(|work| work.checked_add(record_count))
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
                identifier: message_object_identifier(source, route.message, path)?,
                bytes: message_payload(source, route.message, path)?,
            });
        }
        let mut records = Vec::new();
        records
            .try_reserve_exact(record_count)
            .map_err(|_error| allocation(record_count))?;
        for owner in &target.dependencies.formula_owners {
            for route in &owner.cell_record_tiles {
                records.push(cache::DependencyPayload {
                    identifier: message_object_identifier(source, *route, path)?,
                    bytes: message_payload(source, *route, path)?,
                });
            }
        }
        for route in &target.dependencies.inert_marker_tiles {
            records.push(cache::DependencyPayload {
                identifier: message_object_identifier(source, *route, path)?,
                bytes: message_payload(source, *route, path)?,
            });
        }
        records.sort_unstable_by_key(|payload| payload.identifier);
        if records
            .windows(2)
            .any(|pair| pair[0].identifier == pair[1].identifier && pair[0].bytes != pair[1].bytes)
        {
            return Err(Error::InvalidSource { path });
        }
        records.dedup_by_key(|payload| payload.identifier);
        let mut overlay = Vec::new();
        overlay
            .try_reserve_exact(changes.len())
            .map_err(|_error| allocation(changes.len()))?;
        for change in changes {
            let value = match change.input_ref() {
                None => cache::FinalValue::clear(),
                Some(Input::Number(value)) => cache::FinalValue::number(*value),
                Some(Input::Boolean(value)) => cache::FinalValue::boolean(*value),
                Some(Input::Text(_)) => cache::FinalValue::aggregate_ignored(),
                Some(Input::Date(value)) => cache::FinalValue::date(*value),
                Some(Input::Duration(value)) => cache::FinalValue::duration(*value),
                Some(Input::Formula { .. }) => return Err(cache_refusal(path)),
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
    let (formula_cells, formula_entries, formula_payloads, _formula_caches) =
        prepare_complete_formula_cells(source, target, identity, budget, path)?;
    let authored = prepare_final_cache_from_sets(
        source,
        target,
        table,
        &overlay,
        FinalDependencySet {
            engine,
            owners: &owners,
            record_tiles: &records,
            range_tiles: &[],
        },
        cache::FinalFormulaSet {
            tables,
            source_cells: &formula_cells,
            cells: &formula_cells,
            entries: &formula_entries,
            payloads: &formula_payloads,
            authored: &[],
        },
        baseline,
        None,
        budget,
        path,
    )?;
    let materialization_envelope =
        materialization_envelope(target, &authored.rewrites, authored.refreshed_hosts, path)?;
    let tiles = authorized_phase(budget, materialization_envelope, || {
        materialize_tile_changes(source, target, authored.rewrites, path)
    })?;
    Ok(PreparedCache {
        refreshed_hosts: authored.refreshed_hosts,
        tiles,
    })
}

/// Build the selected table's non-formula final scalar overlay for authored
/// formula evaluation. Formula coordinates are supplied through
/// `FinalFormulaSet` and must not also appear here.
pub(super) fn prepare_final_overlay(
    target: &resolve::Target,
    changes: &[Change],
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<Vec<cache::FinalCell>, Error> {
    let selected = target
        .dependencies
        .selected_formula_owner
        .ok_or_else(|| cache_refusal(path))?;
    let identity = cache::TableIdentity {
        owner: selected.internal_owner_id,
        uuid_lower: selected.uid_lower,
        uuid_upper: selected.uid_upper,
    };
    let capacity = changes.len();
    let usage = budget::Usage {
        retained_elements: as_u64(capacity),
        retained_bytes: as_u64(
            capacity
                .checked_mul(size_of::<cache::FinalCell>())
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(capacity != 0),
        transaction_work: as_u64(capacity),
        ..budget::Usage::default()
    };
    authorized_phase(budget, usage, || {
        let mut overlay = Vec::new();
        overlay
            .try_reserve_exact(capacity)
            .map_err(|_error| allocation(capacity))?;
        for change in changes {
            let value = match change.input_ref() {
                Some(Input::Formula { .. }) => continue,
                None => cache::FinalValue::clear(),
                Some(Input::Number(value)) => cache::FinalValue::number(*value),
                Some(Input::Boolean(value)) => cache::FinalValue::boolean(*value),
                Some(Input::Text(_)) => cache::FinalValue::aggregate_ignored(),
                Some(Input::Date(value)) => cache::FinalValue::date(*value),
                Some(Input::Duration(value)) => cache::FinalValue::duration(*value),
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
        exact_capacity(&overlay, capacity, path)?;
        Ok((overlay, usage))
    })
}

/// Preauthorize cache publication against one complete post-authoring graph.
///
/// Supplied caches for supported formulas must equal strict evaluation.
/// Structurally valid evaluator-unsupported formulas may retain a supplied
/// cache only when no impacted descendant consumes it.
#[allow(
    clippy::too_many_arguments,
    reason = "the final graph and shared transaction ledger remain explicit"
)]
pub(super) fn prepare_final_cache_from_sets<'artifact>(
    source: &Package,
    target: &resolve::Target,
    table: &Table,
    overlay: &[cache::FinalCell],
    dependencies: FinalDependencySet<'artifact>,
    formulas: cache::FinalFormulaSet<'artifact>,
    external_baseline: &[ExternalBaselineTable<'_>],
    logical: Option<super::formula_metadata::LogicalGraph<'artifact>>,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<PreparedAuthoredCache, Error> {
    let selected = target
        .dependencies
        .selected_formula_owner
        .ok_or_else(|| cache_refusal(path))?;
    let identity = cache::TableIdentity {
        owner: selected.internal_owner_id,
        uuid_lower: selected.uid_lower,
        uuid_upper: selected.uid_upper,
    };
    let baseline_envelope = SemanticBaseline::envelope(table, path)?;
    let baseline = authorized_phase(budget, baseline_envelope, || {
        SemanticBaseline::build(table, identity.owner, path)
    })?;
    let selected_tables = [cache::TableGeometry {
        identity,
        rows: target.native.rows,
        columns: target.native.columns,
        header_rows: u32::try_from(target.native.settings.header_row_count())
            .map_err(|_error| Error::InvalidSource { path })?,
        header_columns: u32::try_from(target.native.settings.header_column_count())
            .map_err(|_error| Error::InvalidSource { path })?,
        footer_rows: u32::try_from(target.native.settings.footer_row_count())
            .map_err(|_error| Error::InvalidSource { path })?,
    }];
    let tables = if formulas.tables.is_empty() {
        &selected_tables[..]
    } else {
        formulas.tables
    };
    let final_baseline = FinalBaseline {
        selected: &baseline,
        external: external_baseline,
        tables,
    };
    let external_validation = external_baseline_validation_usage(external_baseline, tables, path)?;
    authorized_phase(budget, external_validation, || {
        validate_external_baseline(external_baseline, tables, identity.owner, path)?;
        Ok(((), external_validation))
    })?;
    authorized_cache_phase(source, budget, path, |limits, remaining| {
        let usage =
            cache::validate_formula_cell_payloads(formulas.cells, formulas.payloads, limits)
                .map_err(|error| map_cache_error(error, path, remaining))?;
        Ok(((), usage))
    })?;
    let mut evaluator = authorized_cache_phase(source, budget, path, |limits, remaining| {
        let (evaluator, usage) =
            cache::StrictEvaluator::new(formulas.entries, formulas.payloads, tables, limits)
                .map_err(|error| map_cache_error(error, path, remaining))?;
        Ok((evaluator, usage))
    })?;
    let cache_source = cache::CacheSource {
        selected_table: identity,
        tables,
        engine: dependencies.engine,
        owners: dependencies.owners,
        record_tiles: dependencies.record_tiles,
        range_tiles: dependencies.range_tiles,
        formulas: formulas.cells,
        source_formulas: formulas.source_cells,
    };
    let plan = authorized_cache_phase(source, budget, path, |limits, remaining| {
        let plan = if let Some(logical) = logical {
            cache::plan_final_cache_with_logical_graph(
                cache_source,
                overlay,
                formulas.authored,
                logical,
                limits,
                &final_baseline,
                &mut evaluator,
            )
        } else if formulas.authored.is_empty() {
            cache::plan_final_cache(
                cache_source,
                overlay,
                limits,
                &final_baseline,
                &mut evaluator,
            )
        } else {
            cache::plan_final_cache_with_authored(
                cache_source,
                overlay,
                formulas.authored,
                limits,
                &final_baseline,
                &mut evaluator,
            )
        }
        .map_err(|error| map_cache_error(error, path, remaining))?;
        let usage = plan.usage;
        Ok((plan, usage))
    })?;
    if !plan.removals.is_empty() {
        return Err(cache_refusal(path));
    }
    let refreshed_hosts = usize::try_from(plan.usage.cache_hosts_refreshed)
        .map_err(|_error| Error::InvalidSource { path })?;
    Ok(PreparedAuthoredCache {
        refreshed_hosts,
        rewrites: plan.rewrites,
    })
}

struct SemanticBaseline {
    owner: u32,
    values: Vec<(cache::Coordinate, cache::ScalarValue)>,
    aggregate_ignored: Vec<cache::Coordinate>,
    unsupported: Vec<cache::Coordinate>,
}

impl SemanticBaseline {
    fn envelope(table: &Table, path: Path) -> Result<budget::Usage, Error> {
        let cells = table.iter_cells().len();
        let elements = cells.checked_mul(3).ok_or(Error::InvalidSource { path })?;
        let bytes =
            cells
                .checked_mul(size_of::<(cache::Coordinate, cache::ScalarValue)>())
                .and_then(|bytes| {
                    cells
                        .checked_mul(size_of::<cache::Coordinate>())
                        .and_then(|ignored| {
                            cells.checked_mul(size_of::<cache::Coordinate>()).and_then(
                                |unsupported| bytes.checked_add(ignored)?.checked_add(unsupported),
                            )
                        })
                })
                .ok_or(Error::InvalidSource { path })?;
        Ok(budget::Usage {
            retained_elements: as_u64(elements),
            retained_bytes: as_u64(bytes),
            allocation_events: if cells == 0 { 0 } else { 3 },
            transaction_work: as_u64(cells.checked_mul(2).ok_or(Error::InvalidSource { path })?),
            ..budget::Usage::default()
        })
    }

    fn build(table: &Table, owner: u32, path: Path) -> Result<(Self, budget::Usage), Error> {
        let cells = table.iter_cells();
        let cell_count = cells.len();
        let mut values = Vec::new();
        let mut aggregate_ignored = Vec::new();
        let mut unsupported = Vec::new();
        values
            .try_reserve_exact(cell_count)
            .map_err(|_error| allocation(cell_count))?;
        unsupported
            .try_reserve_exact(cell_count)
            .map_err(|_error| allocation(cell_count))?;
        aggregate_ignored
            .try_reserve_exact(cell_count)
            .map_err(|_error| allocation(cell_count))?;
        exact_capacity(&values, cell_count, path)?;
        exact_capacity(&aggregate_ignored, cell_count, path)?;
        exact_capacity(&unsupported, cell_count, path)?;
        for cell in cells {
            let coordinate = cache::Coordinate {
                row: cell.position().row(),
                column: cell.position().column(),
            };
            match cell.value() {
                Value::Empty => {},
                Value::Number(value) => {
                    values.push((coordinate, cache::ScalarValue::Number(*value)))
                },
                Value::Boolean(value) => {
                    values.push((coordinate, cache::ScalarValue::Boolean(*value)))
                },
                Value::Date(value) => values.push((coordinate, cache::ScalarValue::Date(*value))),
                Value::Duration(value) => {
                    values.push((coordinate, cache::ScalarValue::Duration(*value)))
                },
                Value::Text(_) => aggregate_ignored.push(coordinate),
                Value::Formula(_) | Value::Error(_) => unsupported.push(coordinate),
            }
        }
        if values.windows(2).any(|pair| pair[0].0 >= pair[1].0)
            || unsupported.windows(2).any(|pair| pair[0] >= pair[1])
            || aggregate_ignored.windows(2).any(|pair| pair[0] >= pair[1])
        {
            return Err(Error::InvalidSource { path });
        }
        // All three exactly reserved buffers remain live for every subsequent
        // cache phase. Charge their allocation shapes, including unused
        // slots, rather than only the populated semantic entries.
        let retained_elements = values
            .capacity()
            .checked_add(aggregate_ignored.capacity())
            .and_then(|count| count.checked_add(unsupported.capacity()))
            .ok_or(Error::InvalidSource { path })?;
        let retained_bytes = values
            .capacity()
            .checked_mul(size_of::<(cache::Coordinate, cache::ScalarValue)>())
            .and_then(|bytes| {
                aggregate_ignored
                    .capacity()
                    .checked_mul(size_of::<cache::Coordinate>())
                    .and_then(|ignored_bytes| {
                        unsupported
                            .capacity()
                            .checked_mul(size_of::<cache::Coordinate>())
                            .and_then(|unsupported_bytes| {
                                bytes
                                    .checked_add(ignored_bytes)?
                                    .checked_add(unsupported_bytes)
                            })
                    })
            })
            .ok_or(Error::InvalidSource { path })?;
        let validation_work = values
            .len()
            .saturating_sub(1)
            .checked_add(unsupported.len().saturating_sub(1))
            .and_then(|work| work.checked_add(aggregate_ignored.len().saturating_sub(1)))
            .ok_or(Error::InvalidSource { path })?;
        Ok((
            Self {
                owner,
                values,
                aggregate_ignored,
                unsupported,
            },
            budget::Usage {
                retained_elements: as_u64(retained_elements),
                retained_bytes: as_u64(retained_bytes),
                allocation_events: if cell_count == 0 { 0 } else { 3 },
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
        owner: u32,
        coordinate: cache::Coordinate,
    ) -> Result<Option<cache::ScalarValue>, cache::Failure> {
        if owner != self.owner {
            return Err(cache::Failure::UnsupportedDependency(
                cache::Unsupported::MissingOwner,
            ));
        }
        if self.unsupported.binary_search(&coordinate).is_ok() {
            return Err(cache::Failure::UnsupportedDependency(
                cache::Unsupported::Formula,
            ));
        }
        if self.aggregate_ignored.binary_search(&coordinate).is_ok() {
            return Err(cache::Failure::UnsupportedDependency(
                cache::Unsupported::TextScalar,
            ));
        }
        Ok(self
            .values
            .binary_search_by_key(&coordinate, |entry| entry.0)
            .ok()
            .and_then(|index| self.values.get(index))
            .map(|entry| entry.1))
    }

    fn lookup_work(&self, owner: u32) -> Result<usize, cache::Failure> {
        if owner != self.owner {
            return Err(cache::Failure::UnsupportedDependency(
                cache::Unsupported::MissingOwner,
            ));
        }
        cache::baseline_lookup_work(
            self.values.len(),
            self.aggregate_ignored.len(),
            self.unsupported.len(),
        )
    }
}

struct FinalBaseline<'source> {
    selected: &'source SemanticBaseline,
    external: &'source [ExternalBaselineTable<'source>],
    tables: &'source [cache::TableGeometry],
}

impl cache::CacheBaseline for FinalBaseline<'_> {
    fn value(
        &self,
        owner: u32,
        coordinate: cache::Coordinate,
    ) -> Result<Option<cache::ScalarValue>, cache::Failure> {
        if owner == self.selected.owner {
            return self.selected.value(owner, coordinate);
        }
        let table = self
            .tables
            .binary_search_by_key(&owner, |table| table.identity.owner)
            .map_err(|_error| {
                cache::Failure::UnsupportedDependency(cache::Unsupported::MissingOwner)
            })?;
        if coordinate.row >= self.tables[table].rows
            || coordinate.column >= self.tables[table].columns
        {
            return Err(cache::Failure::InvalidSource);
        }
        let position = self
            .external
            .binary_search_by_key(&owner, |entry| entry.owner)
            .map_err(|_| cache::Failure::UnsupportedDependency(cache::Unsupported::MissingOwner))?;
        match self.external[position]
            .table
            .view(crate::table::CellPosition::new(
                coordinate.row,
                coordinate.column,
            )) {
            crate::table::View::Missing | crate::table::View::Stored(Value::Empty) => Ok(None),
            crate::table::View::Stored(Value::Number(value)) => {
                Ok(Some(cache::ScalarValue::Number(*value)))
            },
            crate::table::View::Stored(Value::Boolean(value)) => {
                Ok(Some(cache::ScalarValue::Boolean(*value)))
            },
            crate::table::View::Stored(Value::Date(value)) => {
                Ok(Some(cache::ScalarValue::Date(*value)))
            },
            crate::table::View::Stored(Value::Duration(value)) => {
                Ok(Some(cache::ScalarValue::Duration(*value)))
            },
            crate::table::View::Stored(Value::Text(_)) => Err(
                cache::Failure::UnsupportedDependency(cache::Unsupported::TextScalar),
            ),
            crate::table::View::Stored(Value::Formula(_) | Value::Error(_))
            | crate::table::View::Covered => Err(cache::Failure::UnsupportedDependency(
                cache::Unsupported::Formula,
            )),
        }
    }

    fn lookup_work(&self, owner: u32) -> Result<usize, cache::Failure> {
        if owner == self.selected.owner {
            return self.selected.lookup_work(owner);
        }
        let position = self
            .external
            .binary_search_by_key(&owner, |entry| entry.owner)
            .map_err(|_| cache::Failure::UnsupportedDependency(cache::Unsupported::MissingOwner))?;
        binary_search_work(self.external.len())
            .checked_add(binary_search_work(
                self.external[position].table.cell_count(),
            ))
            .and_then(|work| work.checked_add(binary_search_work(self.tables.len())))
            .ok_or(cache::Failure::InvalidSource)
    }
}

fn external_baseline_validation_usage(
    external: &[ExternalBaselineTable<'_>],
    tables: &[cache::TableGeometry],
    path: Path,
) -> Result<budget::Usage, Error> {
    let per_table = binary_search_work(tables.len())
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    let work = external
        .len()
        .checked_mul(per_table)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        lookups: as_u64(external.len()),
        transaction_work: as_u64(work),
        ..budget::Usage::default()
    })
}

fn validate_external_baseline(
    external: &[ExternalBaselineTable<'_>],
    tables: &[cache::TableGeometry],
    selected_owner: u32,
    path: Path,
) -> Result<(), Error> {
    if external
        .windows(2)
        .any(|pair| pair[0].owner >= pair[1].owner)
    {
        return Err(Error::InvalidSource { path });
    }
    for baseline in external {
        let geometry = tables
            .binary_search_by_key(&baseline.owner, |table| table.identity.owner)
            .ok()
            .and_then(|index| tables.get(index))
            .ok_or(Error::InvalidSource { path })?;
        if baseline.owner == selected_owner
            || baseline.table.row_count() != geometry.rows
            || baseline.table.column_count() != geometry.columns
        {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(())
}

fn prepare_complete_formula_cells<'source>(
    source: &'source Package,
    target: &resolve::Target,
    identity: cache::TableIdentity,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<
    (
        Vec<cache::FormulaCell>,
        Vec<cache::FormulaListEntry<'source>>,
        Vec<cache::FormulaPayload<'source>>,
        Vec<ExistingFormulaCache>,
    ),
    Error,
> {
    let entries = collect_formula_entries(
        source,
        &target.storage.lists.formula,
        identity,
        budget,
        path,
    )?;
    let tile_count = target.storage.tiles.len();
    let tile_stage = budget::Usage {
        retained_elements: as_u64(tile_count),
        retained_bytes: as_u64(
            tile_count
                .checked_mul(size_of::<PreparedFormulaTileScan>())
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(tile_count != 0),
        ..budget::Usage::default()
    };
    let mut scans = authorized_phase(budget, tile_stage, || {
        let mut scans = Vec::new();
        scans
            .try_reserve_exact(tile_count)
            .map_err(|_| allocation(tile_count))?;
        exact_capacity(&scans, tile_count, path)?;
        Ok((scans, tile_stage))
    })?;
    for route in &target.storage.tiles {
        let route_usage = budget::Usage {
            lookups: 2,
            transaction_work: 2,
            ..budget::Usage::default()
        };
        let (payload, cache_object) = authorized_phase(budget, route_usage, || {
            Ok((
                (
                    message_payload(source, route.message, path)?,
                    message_object_identifier(source, route.message, path)?,
                ),
                route_usage,
            ))
        })?;
        let remaining = budget.remaining()?;
        let cache_limits = cache_limits(source, remaining, path)?;
        let retained_cell_limit = usize::try_from(remaining.retained_elements)
            .unwrap_or(usize::MAX)
            .min(
                usize::try_from(remaining.retained_bytes).unwrap_or(usize::MAX)
                    / size_of::<tile::ScannedFormulaCell>(),
            );
        let cell_limit = cache_limits
            .graph_nodes
            .min(cache_limits.cache_cells)
            .min(retained_cell_limit)
            .min(if remaining.allocation_events == 0 {
                0
            } else {
                usize::MAX
            });
        let scan_work = budget::tile_work_ceiling(
            remaining,
            u64::try_from(cache_limits.wire_work).map_err(|_| Error::InvalidSource { path })?,
            0,
        )?;
        let scan = authorized_phase(budget, remaining, || {
            let scan = tile::scan_formula_cells(
                payload,
                target.native.columns,
                tile::TileLimits::new(
                    cache_limits.wire_bytes,
                    cache_limits.retained_bytes.min(cache_limits.wire_bytes),
                    cache_limits.wire_fields,
                    scan_work,
                    usize::try_from(target.storage.tile_size)
                        .map_err(|_| Error::InvalidSource { path })?,
                    cell_limit,
                ),
            )
            .map_err(|error| super::map_tile_error(error, path))?;
            let report = tile_report_usage(scan.report, path)?;
            Ok((scan, report))
        })?;
        scans.push(PreparedFormulaTileScan {
            tile_id: route.tile_id,
            cache_object,
            scan,
        });
    }
    if scans.len() != tile_count {
        return Err(Error::InvalidSource { path });
    }
    let count = scans.iter().try_fold(0usize, |count, tile| {
        count
            .checked_add(tile.scan.cells.len())
            .ok_or(Error::InvalidSource { path })
    })?;
    let retained_bytes = count
        .checked_mul(
            size_of::<cache::FormulaCell>()
                + size_of::<cache::FormulaPayload<'_>>()
                + size_of::<ExistingFormulaCache>(),
        )
        .ok_or(Error::InvalidSource { path })?;
    let join_work = binary_search_work(entries.len())
        .checked_add(6)
        .and_then(|per| per.checked_mul(count))
        .and_then(|work| work.checked_add(tile_count))
        .ok_or(Error::InvalidSource { path })?;
    let usage = budget::Usage {
        retained_elements: as_u64(count.checked_mul(3).ok_or(Error::InvalidSource { path })?),
        retained_bytes: as_u64(retained_bytes),
        allocation_events: u64::from(count != 0) * 3,
        lookups: as_u64(count),
        transaction_work: as_u64(join_work),
        ..budget::Usage::default()
    };
    let (cells, payloads, caches) = authorized_phase(budget, usage, || {
        let mut cells = Vec::new();
        let mut payloads = Vec::new();
        let mut caches = Vec::new();
        cells
            .try_reserve_exact(count)
            .map_err(|_| allocation(count))?;
        payloads
            .try_reserve_exact(count)
            .map_err(|_| allocation(count))?;
        caches
            .try_reserve_exact(count)
            .map_err(|_| allocation(count))?;
        let mut previous = None;
        for tile in &scans {
            for scanned in &tile.scan.cells {
                let row = tile
                    .tile_id
                    .checked_mul(target.storage.tile_size)
                    .and_then(|base| base.checked_add(scanned.row))
                    .filter(|row| *row < target.native.rows)
                    .ok_or(Error::InvalidSource { path })?;
                let coordinate = cache::Coordinate {
                    row,
                    column: scanned.column,
                };
                if previous.is_some_and(|prior| prior >= coordinate) {
                    return Err(Error::InvalidSource { path });
                }
                previous = Some(coordinate);
                let entry = entries
                    .binary_search_by_key(&scanned.identifier, |entry| entry.key)
                    .ok()
                    .and_then(|position| entries.get(position))
                    .ok_or(Error::InvalidSource { path })?;
                cells.push(cache::FormulaCell {
                    owner: identity.owner,
                    coordinate,
                    cache_object: tile.cache_object,
                });
                payloads.push(cache::FormulaPayload {
                    owner: identity.owner,
                    coordinate,
                    key: scanned.identifier,
                    bytes: entry.bytes,
                });
                caches.push(ExistingFormulaCache {
                    owner: identity.owner,
                    row,
                    column: scanned.column,
                    value: scanned.cache,
                    formula_error: scanned.formula_error,
                });
            }
        }
        exact_capacity(&cells, count, path)?;
        exact_capacity(&payloads, count, path)?;
        exact_capacity(&caches, count, path)?;
        Ok(((cells, payloads, caches), usage))
    })?;
    Ok((cells, entries, payloads, caches))
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
    fn prepared_formula_list_source_requires_exact_segment_references_and_hosts() {
        let segment = PreparedFormulaListMessage {
            object_id: 20,
            payload: &[1],
            object_references: &[],
        };
        let source = PreparedFormulaListSource {
            root: PreparedFormulaListMessage {
                object_id: 10,
                payload: &[1],
                object_references: &[20],
            },
            segments: vec![segment],
            expected_entries: 1,
            hosts: vec![PreparedFormulaListHost {
                owner: 7,
                row: 2,
                column: 3,
                key: 9,
            }],
        };
        let identity = cache::TableIdentity {
            owner: 7,
            uuid_lower: 1,
            uuid_upper: 2,
        };
        validate_prepared_formula_list_source(&source, identity, 1, Path::Package)
            .expect("exact route and host authority");

        let wrong_reference = PreparedFormulaListSource {
            root: PreparedFormulaListMessage {
                object_id: 10,
                payload: &[1],
                object_references: &[21],
            },
            segments: vec![segment],
            expected_entries: 1,
            hosts: source.hosts,
        };
        assert!(matches!(
            validate_prepared_formula_list_source(&wrong_reference, identity, 1, Path::Package),
            Err(Error::InvalidSource {
                path: Path::Package
            })
        ));
    }

    #[test]
    fn formula_baseline_retains_all_capacities_and_next_phase_refuses_before_callback() {
        let mut builder = Table::builder("Formula", Dimensions::new(1, 5));
        builder
            .set(
                CellPosition::new(0, 0),
                Value::number(1.0).expect("finite baseline number"),
            )
            .expect("number cell is in bounds");
        builder
            .set(CellPosition::new(0, 1), Value::Formula("=A1".to_owned()))
            .expect("formula cell is in bounds");
        builder
            .set(
                CellPosition::new(0, 2),
                Value::Text("ignored by SUM".to_owned()),
            )
            .expect("text cell is in bounds");
        builder
            .set(
                CellPosition::new(0, 3),
                Value::date(12.0).expect("finite date"),
            )
            .expect("date cell is in bounds");
        builder
            .set(
                CellPosition::new(0, 4),
                Value::duration(3.5).expect("finite duration"),
            )
            .expect("duration cell is in bounds");
        let table = builder.finish().expect("formula table is valid");
        let path = Path::Table { sheet: 0, table: 0 };
        let mut transaction =
            budget::TransactionBudget::from_limits(limits(15)).expect("finite limits");

        let baseline = authorized_phase(
            &mut transaction,
            SemanticBaseline::envelope(&table, path).expect("baseline envelope"),
            || SemanticBaseline::build(&table, 1, path),
        )
        .expect("baseline fits exactly");
        assert_eq!(baseline.values.capacity(), 5);
        assert_eq!(baseline.aggregate_ignored.capacity(), 5);
        assert_eq!(baseline.unsupported.capacity(), 5);
        assert!(matches!(
            cache::CacheBaseline::value(
                &baseline,
                1,
                cache::Coordinate { row: 0, column: 3 },
            ),
            Ok(Some(cache::ScalarValue::Date(value))) if value.get() == 12.0
        ));
        assert!(matches!(
            cache::CacheBaseline::value(
                &baseline,
                1,
                cache::Coordinate { row: 0, column: 4 },
            ),
            Ok(Some(cache::ScalarValue::Duration(value))) if value.get() == 3.5
        ));

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
                observed: 16,
                maximum: 15,
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
