//! Exact-source selector-first Numbers table-cell mutations.

mod budget;
mod cache;
mod cache_commit;
mod lists;
mod locality;
mod metadata;
mod resolve;
mod rewrite;
mod rich;
mod rich_commit;
mod sparse;
mod sparse_commit;
mod tile;

use std::{cell::Cell, mem::size_of, sync::Arc};

use crate::{
    SheetSelector, TableSelector,
    table::{
        View,
        cells::{
            Commit, Diagnostics, DirectionalMessage, Edit, Error, EvidenceChangeKind,
            FieldReferenceRoute, Input, MessageReferenceRoute, Patch, PatchEvidence,
            PhysicalLocation, Plan, ReferenceEvidence, ReferenceSpan,
        },
    },
};

use super::{Package, table_cells::resolve_table};

impl Package {
    /// Start one bounded selector-first table-cell batch.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, ambiguity, source, or allocation failure.
    pub fn edit_table_cells<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let selected = resolve_table(self, sheet.into(), table.into())?;
        Edit::from_resolved(
            self,
            selected.path,
            selected.table.dimensions(),
            self.state.options.semantic().max_materialized_cells(),
            self.state.options.semantic().max_output_text_bytes(),
            commit_plan,
        )
    }

    /// Apply a reversible patch to its exact retained source artifact.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] unless this package is the exact
    /// retained directional source and read profile, or when directional
    /// locality verification rejects the retained target snapshot.
    pub fn apply_table_cells(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = Arc::clone(&self.state.source);
        if !patch.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit::new(
                self.snapshot(),
                patch.clone(),
                Diagnostics::unchanged(patch.len()),
            ));
        }
        let evidence = patch.evidence();
        if evidence.message_count() == 0 {
            return Err(Error::Verification { path: patch.path() });
        }
        let source_package = patch
            .source_package()
            .ok_or(Error::Verification { path: patch.path() })?;
        let target_package = patch
            .target_package()
            .ok_or(Error::Verification { path: patch.path() })?;
        if source_package.read_options() != self.read_options()
            || target_package.read_options() != self.read_options()
        {
            return Err(Error::PatchConflict);
        }
        let mut budget = budget::TransactionBudget::new(self)?;
        let locality_remaining = budget.remaining()?;
        let locality_plan_usage =
            evidence_locality_plan_usage(evidence.message_count(), patch.path())?;
        let locality_byte_bound = usize_u64(
            self.source_bytes()
                .len()
                .checked_add(patch.target_bytes().len())
                .ok_or(Error::InvalidSource { path: patch.path() })?,
        );
        if locality_byte_bound > locality_remaining.locality_bytes {
            return Err(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: locality_byte_bound,
                maximum: locality_remaining.locality_bytes,
                path: patch.path(),
            });
        }
        preflight_usage(locality_plan_usage, locality_remaining, patch.path())?;
        let admitted_apply_work = locality_plan_usage.transaction_work;
        let locality_work = locality_remaining
            .transaction_work
            .checked_sub(admitted_apply_work)
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: admitted_apply_work,
                maximum: locality_remaining.transaction_work,
                path: patch.path(),
            })?;
        let mut apply_envelope = locality_remaining;
        preflight_usage(locality_plan_usage, apply_envelope, patch.path())?;
        apply_envelope.locality_bytes = apply_envelope.locality_bytes.min(locality_byte_bound);
        let target = budget.with_apply_authorization(apply_envelope, || {
            let locality_report = verify_evidence_locality_with_report(
                self,
                target_package,
                evidence,
                locality_work,
                patch.path(),
            )?;
            let mut exact = budget::Usage::default();
            merge_usage(&mut exact, locality_plan_usage, patch.path())?;
            merge_usage(&mut exact, locality_report.usage(), patch.path())?;
            Ok((target_package.snapshot(), exact))
        })?;
        let touched_components = distinct_evidence_components(evidence, patch.path())?;
        let deleted_previews = evidence
            .source_previews()
            .saturating_sub(evidence.target_previews());
        Ok(Commit::new(
            target,
            patch.clone(),
            Diagnostics::applied(
                patch.len(),
                patch.changed_cells(),
                touched_components,
                0,
                deleted_previews,
            ),
        ))
    }
}

struct PreparedReplacement {
    route: resolve::MessageRoute,
    payload: Vec<u8>,
    references: Option<PreparedReferenceDelta>,
}

struct PreparedReferenceDelta {
    before: Vec<u64>,
    after: Vec<u64>,
}

fn commit_plan(plan: Plan<'_>) -> Result<Commit, Error> {
    let (source, path, dimensions, changes, _owned_value_bytes, staging_usage) = plan.into_parts();
    for change in &changes {
        let position = change.position();
        if position.row() >= dimensions.rows() || position.column() >= dimensions.columns() {
            return Err(Error::OutOfBounds {
                position,
                dimensions,
            });
        }
    }

    let (sheet_position, table_position) = match path {
        crate::package::table_cells::Path::Table { sheet, table } => (
            usize::try_from(sheet).map_err(|_error| Error::InvalidSource { path })?,
            usize::try_from(table).map_err(|_error| Error::InvalidSource { path })?,
        ),
        crate::package::table_cells::Path::Package => {
            return Err(Error::InvalidSource { path });
        },
    };
    let selected = resolve_table(
        source,
        SheetSelector::Index(sheet_position),
        TableSelector::Index(table_position),
    )?;
    let requested = changes.len();
    let mut changed_count = 0usize;
    let mut changed_owned_value_bytes = 0usize;
    for change in &changes {
        if !semantic_change_is_noop(selected.table, change) {
            changed_count = changed_count.checked_add(1).ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::Updates,
                observed: u64::MAX,
                maximum: u64::MAX - 1,
                path,
            })?;
            changed_owned_value_bytes = changed_owned_value_bytes
                .checked_add(change.input_ref().map_or(0, Input::owned_bytes))
                .ok_or(Error::LimitExceeded {
                    kind: crate::package::table_cells::LimitKind::OwnedValueBytes,
                    observed: u64::MAX,
                    maximum: u64::MAX - 1,
                    path,
                })?;
        }
    }
    if changed_count == 0 {
        let source_bytes = Arc::clone(&source.state.source);
        let patch = Patch::from_exact(path, requested, 0, Arc::clone(&source_bytes), source_bytes);
        return Ok(Commit::new(
            source.snapshot(),
            patch,
            Diagnostics::unchanged(requested),
        ));
    }
    let mut changed = Vec::new();
    changed
        .try_reserve_exact(changed_count)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::Updates,
            amount: changed_count,
        })?;
    for change in changes {
        if !semantic_change_is_noop(selected.table, &change) {
            changed.push(change);
        }
    }
    if changed.len() != changed_count {
        return Err(Error::Verification { path });
    }

    commit_existing_scalar_tiles(
        source,
        selected.table,
        path,
        sheet_position,
        table_position,
        changed,
        changed_owned_value_bytes,
        staging_usage,
        requested,
    )
}

fn semantic_change_is_noop(
    table: &crate::table::Table,
    change: &crate::table::cells::Change,
) -> bool {
    match change.input_ref() {
        Some(input) => match table.view(change.position()) {
            View::Stored(value) => input.matches_value(value),
            View::Missing | View::Covered => false,
        },
        None => matches!(
            table.view(change.position()),
            View::Missing | View::Stored(crate::cell::Value::Empty)
        ),
    }
}

fn commit_existing_scalar_tiles(
    source: &Package,
    table: &crate::table::Table,
    path: crate::package::table_cells::Path,
    sheet_position: usize,
    table_position: usize,
    changes: Vec<crate::table::cells::Change>,
    owned_value_bytes: usize,
    staging_usage: crate::table::cells::StagingUsage,
    requested: usize,
) -> Result<Commit, Error> {
    let mut budget = budget::TransactionBudget::new(source)?;
    budget.reserve(caller_staging_usage(
        &budget,
        changes.len(),
        owned_value_bytes,
        staging_usage,
    )?)?;
    let resolve_authorization = authorize_remaining(&mut budget)?;
    let mut positions = Vec::new();
    if let Err(_error) = positions.try_reserve_exact(changes.len()) {
        budget.cancel_authorization();
        return Err(Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: changes.len(),
        });
    }
    for change in &changes {
        positions.push(change.position());
    }
    let (target, resolve_report) = match resolve::resolve_changed_target_with_remaining(
        source,
        sheet_position,
        table_position,
        &positions,
        resolve_authorization,
    ) {
        Ok(result) => result,
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    let resolve_usage = match resolve_usage(&budget, resolve_report) {
        Ok(usage) => usage,
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    budget.record_authorized(resolve_usage)?;
    if target.path() != path {
        return Err(Error::InvalidSource { path });
    }
    let cache_prepared =
        cache_commit::prepare_final_cache(source, &target, table, &changes, &mut budget, path)?;
    if changes.iter().any(|change| {
        let tile_id = change.position().row() / target.storage.tile_size;
        target
            .storage
            .tiles
            .binary_search_by_key(&tile_id, |route| route.tile_id)
            .is_err()
    }) {
        if !cache_prepared.tiles.is_empty() {
            return Err(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::FormulaCache,
            });
        }
        return sparse_commit::commit_sparse_scalar_tiles(
            source,
            path,
            sheet_position,
            table_position,
            &target,
            budget,
            changes,
            requested,
        );
    }

    let rich_prepared = rich_commit::prepare_unique_rich_text_to_text(
        source,
        &target,
        &changes,
        &mut budget,
        path,
    )?;
    let rich_key_count = rich_prepared
        .keys
        .iter()
        .filter(|key| key.is_some())
        .count();
    if rich_prepared.owned_transition_count > rich_key_count
        || rich_prepared.owned_transition_count != rich_prepared.replacements.len()
    {
        return Err(Error::Verification { path });
    }

    let text_count = changes
        .iter()
        .enumerate()
        .filter(|(index, change)| {
            rich_prepared.keys[*index].is_none()
                && matches!(change.input_ref(), Some(Input::Text(_)))
        })
        .count();
    let mut text_requests = Vec::new();
    let mut text_change_indices = Vec::new();
    reserve_retained_vec::<lists::StringRequest<'_>>(&mut budget, text_count, path)?;
    reserve_retained_vec::<usize>(&mut budget, text_count, path)?;
    text_requests
        .try_reserve_exact(text_count)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: text_count,
        })?;
    text_change_indices
        .try_reserve_exact(text_count)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: text_count,
        })?;
    for (index, change) in changes.iter().enumerate() {
        if rich_prepared.keys[index].is_none() {
            if let Some(Input::Text(text)) = change.input_ref() {
                text_requests.push(lists::StringRequest::new(text));
                text_change_indices.push(index);
            }
        }
    }
    if !text_requests.is_empty() && !target.storage.lists.string.segments.is_empty() {
        return Err(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::SharedString,
        });
    }
    let string_list_payload = message_payload(source, target.storage.lists.string.message, path)?;
    let mut string_keys = Vec::new();
    reserve_retained_vec::<Option<u32>>(
        &mut budget,
        if text_count == 0 { 0 } else { changes.len() },
        path,
    )?;
    string_keys
        .try_reserve_exact(if text_count == 0 { 0 } else { changes.len() })
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: if text_count == 0 { 0 } else { changes.len() },
        })?;
    if text_count != 0 {
        string_keys.resize(changes.len(), None);
    }
    if !text_requests.is_empty() {
        let remaining = budget.remaining()?;
        let limits = string_list_limits(source, changes.len(), path, remaining)?;
        budget.authorize(remaining)?;
        let preliminary = match lists::preflight_string_assignments(
            string_list_payload,
            &text_requests,
            limits,
        ) {
            Ok(assignments) => assignments,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_list_error(error, path));
            },
        };
        let usage = assignment_usage(&budget, preliminary.report())?;
        budget.record_authorized(usage)?;
        for assignment in preliminary.assignments() {
            let change_index = *text_change_indices
                .get(assignment.request())
                .ok_or(Error::InvalidSource { path })?;
            *string_keys
                .get_mut(change_index)
                .ok_or(Error::InvalidSource { path })? = Some(assignment.key());
        }
    }
    let cache_row_capacity = cache_prepared
        .tiles
        .iter()
        .try_fold(0usize, |count, tile| count.checked_add(tile.changes.len()))
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            observed: u64::MAX,
            maximum: budget.limits().max_retained_elements,
            path,
        })?;
    let final_row_capacity =
        changes
            .len()
            .checked_add(cache_row_capacity)
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                observed: u64::MAX,
                maximum: budget.limits().max_retained_elements,
                path,
            })?;
    let prepared_capacity = rich_prepared
        .replacements
        .len()
        .checked_add(changes.len())
        .and_then(|count| count.checked_add(cache_prepared.tiles.len()))
        .and_then(|count| count.checked_add(final_row_capacity))
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            observed: u64::MAX,
            maximum: budget.limits().max_retained_elements,
            path,
        })?;
    let mut prepared = Vec::new();
    reserve_retained_vec::<PreparedReplacement>(&mut budget, prepared_capacity, path)?;
    prepared
        .try_reserve_exact(prepared_capacity)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: changes.len(),
        })?;
    for replacement in rich_prepared.replacements {
        let references = rich_archive_reference_delta(
            source,
            replacement.route,
            &replacement.references,
            &mut budget,
            path,
        )?;
        prepared.push(PreparedReplacement {
            route: replacement.route,
            payload: replacement.payload,
            references: Some(references),
        });
    }

    let mut releases: Vec<(u32, u32)> = Vec::new();
    let mut final_rows = Vec::new();
    reserve_retained_vec::<sparse::FinalRowCount>(&mut budget, final_row_capacity, path)?;
    final_rows
        .try_reserve_exact(final_row_capacity)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: final_row_capacity,
        })?;
    let mut start = 0usize;
    let mut cache_tile_index = 0usize;
    while start < changes.len() || cache_tile_index < cache_prepared.tiles.len() {
        let scalar_tile_id = changes
            .get(start)
            .map(|change| change.position().row() / target.storage.tile_size);
        let cache_tile_id = cache_prepared
            .tiles
            .get(cache_tile_index)
            .map(|tile| tile.tile_id);
        let tile_id = match (scalar_tile_id, cache_tile_id) {
            (Some(scalar), Some(cache)) => scalar.min(cache),
            (Some(scalar), None) => scalar,
            (None, Some(cache)) => cache,
            (None, None) => return Err(Error::InvalidSource { path }),
        };
        let end = if scalar_tile_id == Some(tile_id) {
            changes[start..]
                .iter()
                .position(|change| change.position().row() / target.storage.tile_size != tile_id)
                .map_or(changes.len(), |offset| start + offset)
        } else {
            start
        };
        let cache_changes: &[tile::CacheChange] = if cache_tile_id == Some(tile_id) {
            &cache_prepared
                .tiles
                .get(cache_tile_index)
                .ok_or(Error::InvalidSource { path })?
                .changes
        } else {
            &[]
        };
        let route = target
            .storage
            .tiles
            .binary_search_by_key(&tile_id, |route| route.tile_id)
            .ok()
            .and_then(|index| target.storage.tiles.get(index))
            .ok_or(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::CellStorage,
            })?;
        let source_payload = message_payload(source, route.message, path)?;
        let mut tile_changes = Vec::new();
        reserve_retained_vec::<tile::TileChange>(&mut budget, end - start, path)?;
        tile_changes
            .try_reserve_exact(end - start)
            .map_err(|_error| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: end - start,
            })?;
        for (offset, change) in changes[start..end].iter().enumerate() {
            tile_changes.push(tile::TileChange {
                row: change.position().row() % target.storage.tile_size,
                column: change.position().column(),
                change: scalar_change(
                    change,
                    string_keys.get(start + offset).copied().flatten(),
                    rich_prepared.keys.get(start + offset).copied().flatten(),
                    path,
                )?,
            });
        }
        let maximum_wire = source.state.options.archive().max_iwa_stream_bytes();
        let maximum_work = maximum_wire.checked_mul(32).ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::WireWork,
            observed: u64::MAX,
            maximum: usize_u64(maximum_wire),
            path,
        })?;
        let remaining = budget.remaining()?;
        let transition_bytes = tile_changes
            .len()
            .checked_mul(size_of::<tile::CellTransition>())
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::RetainedBytes,
                observed: u64::MAX,
                maximum: remaining.retained_bytes,
                path,
            })?;
        let scalar_rows = usize::from(!tile_changes.is_empty())
            .checked_add(
                tile_changes
                    .windows(2)
                    .filter(|pair| pair[0].row != pair[1].row)
                    .count(),
            )
            .ok_or(Error::InvalidSource { path })?;
        let cache_rows = usize::from(!cache_changes.is_empty())
            .checked_add(
                cache_changes
                    .windows(2)
                    .filter(|pair| pair[0].row != pair[1].row)
                    .count(),
            )
            .ok_or(Error::InvalidSource { path })?;
        let distinct_rows = scalar_rows
            .checked_add(cache_rows)
            .ok_or(Error::InvalidSource { path })?;
        let maximum_writes = tile_changes
            .len()
            .checked_add(cache_changes.len())
            .ok_or(Error::InvalidSource { path })?;
        let retained_elements = tile_changes
            .len()
            .checked_add(distinct_rows)
            .and_then(|elements| elements.checked_add(1))
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                observed: u64::MAX,
                maximum: remaining.retained_elements,
                path,
            })?;
        let allocation_events = distinct_rows
            .checked_mul(3)
            .and_then(|events| events.checked_add(3))
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: remaining.allocation_events,
                path,
            })?;
        if usize_u64(retained_elements) > remaining.retained_elements
            || usize_u64(transition_bytes) > remaining.retained_bytes
            || usize_u64(allocation_events) > remaining.allocation_events
        {
            return Err(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::RetainedBytes,
                observed: usize_u64(transition_bytes),
                maximum: remaining.retained_bytes,
                path,
            });
        }
        budget.authorize(remaining)?;
        let max_input_bytes = usize::try_from(remaining.wire_bytes)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire);
        let output_retained = remaining
            .retained_bytes
            .checked_sub(usize_u64(transition_bytes))
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::RetainedBytes,
                observed: usize_u64(transition_bytes),
                maximum: remaining.retained_bytes,
                path,
            })?;
        let max_output_bytes = usize::try_from(output_retained.min(remaining.peak_scratch_bytes))
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire);
        let max_fields = usize::try_from(remaining.wire_fields)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire);
        let transaction_wire_cap = remaining
            .transaction_work
            .checked_sub(usize_u64(maximum_writes))
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: usize_u64(maximum_writes),
                maximum: remaining.transaction_work,
                path,
            })?
            / 2;
        let max_work = remaining
            .wire_work
            .min(remaining.transaction_work)
            .min(remaining.peak_scratch_bytes)
            .min(transaction_wire_cap);
        let outcome = match tile::rewrite_tile_with_cache(
            tile::TileRewriteRequest {
                source: source_payload,
                columns: target.native.columns,
                changes: &tile_changes,
                limits: tile::TileLimits::new(
                    max_input_bytes,
                    max_output_bytes,
                    max_fields,
                    usize_u64(maximum_work).min(max_work),
                    usize::try_from(target.storage.tile_size)
                        .map_err(|_error| Error::InvalidSource { path })?,
                    source.state.options.semantic().max_materialized_cells(),
                ),
            },
            cache_changes,
        ) {
            Ok(outcome) => outcome,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_tile_error(error, path));
            },
        };
        let usage = tile_usage(&budget, outcome.report)?;
        budget.record_authorized(usage)?;
        for row in &outcome.final_rows {
            let global_row = tile_id
                .checked_mul(target.storage.tile_size)
                .and_then(|base| base.checked_add(row.row))
                .ok_or(Error::InvalidSource { path })?;
            final_rows.push(sparse::FinalRowCount {
                row: global_row,
                number_of_cells: row.cell_count,
            });
        }
        let mut transition_index = 0usize;
        for (offset, change) in changes[start..end].iter().enumerate() {
            let local_row = change.position().row() % target.storage.tile_size;
            let column = change.position().column();
            let transition = outcome
                .transitions
                .get(transition_index)
                .filter(|transition| transition.row == local_row && transition.column == column);
            let owned_rich = rich_prepared.keys[start + offset];
            let Some(transition) = transition else {
                if owned_rich.is_some() {
                    continue;
                }
                return Err(Error::Verification { path });
            };
            transition_index = transition_index
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            if let Some(identifier) = transition.before_references.string {
                if releases.capacity() == 0 {
                    reserve_retained_vec::<(u32, u32)>(&mut budget, changes.len(), path)?;
                    releases
                        .try_reserve_exact(changes.len())
                        .map_err(|_error| Error::Allocation {
                            kind: crate::package::table_cells::LimitKind::RetainedElements,
                            amount: changes.len(),
                        })?;
                }
                releases.push((identifier, 1));
            }
            if transition.before_references.rich_text.is_some()
                && !(owned_rich == transition.before_references.rich_text
                    && owned_rich == transition.after_references.rich_text)
            {
                return Err(Error::UnsupportedDependency {
                    path,
                    kind: crate::package::table_cells::DependencyKind::RichText,
                });
            }
            if transition.before_references.formula.is_some()
                || transition.before_references.formula_error.is_some()
                || matches!(
                    transition.before,
                    tile::CellValue::Formula { .. } | tile::CellValue::Error(_)
                )
            {
                return Err(Error::UnsupportedDependency {
                    path,
                    kind: crate::package::table_cells::DependencyKind::Formula,
                });
            }
        }
        if transition_index != outcome.transitions.len() {
            return Err(Error::Verification { path });
        }
        if let Some(payload) = outcome.payload {
            prepared.push(PreparedReplacement {
                route: route.message,
                payload,
                references: None,
            });
        } else if rich_prepared.keys[start..end].iter().any(Option::is_none)
            || !cache_changes.is_empty()
        {
            return Err(Error::Verification { path });
        }
        if cache_tile_id == Some(tile_id) {
            cache_tile_index = cache_tile_index
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
        }
        start = end;
    }

    final_rows.sort_unstable_by_key(|row| row.row);
    if final_rows.windows(2).any(|pair| pair[0].row == pair[1].row) {
        return Err(Error::InvalidSource { path });
    }
    let mut row_start = 0usize;
    while row_start < final_rows.len() {
        let bucket_index = final_rows[row_start].row / sparse::HEADER_BUCKET_ROWS;
        let row_end = final_rows[row_start..]
            .partition_point(|row| row.row / sparse::HEADER_BUCKET_ROWS == bucket_index)
            .checked_add(row_start)
            .ok_or(Error::InvalidSource { path })?;
        let route = *target
            .storage
            .row_headers
            .get(usize::try_from(bucket_index).map_err(|_error| Error::InvalidSource { path })?)
            .ok_or(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::CellStorage,
            })?;
        let remaining = budget.remaining()?;
        let limits = sparse_commit::sparse_limits(source, row_end - row_start, remaining, path)?;
        budget.authorize(remaining)?;
        let rewritten = sparse::rewrite_existing_header_bucket_final_rows_with_report(
            message_payload(source, route, path)?,
            bucket_index,
            target.native.columns,
            &final_rows[row_start..row_end],
            limits,
        );
        let (payload, report) = match rewritten {
            Ok(rewritten) => rewritten,
            Err(error) => {
                budget.cancel_authorization();
                return Err(sparse_commit::map_sparse_error(error, path));
            },
        };
        let usage = match sparse_commit::sparse_usage(report, path) {
            Ok(usage) => usage,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        budget.record_authorized(usage)?;
        if let Some(payload) = payload {
            prepared.push(PreparedReplacement {
                route,
                payload,
                references: None,
            });
        }
        row_start = row_end;
    }

    releases.sort_unstable_by_key(|release| release.0);
    let mut write = 0usize;
    for read in 0..releases.len() {
        if write != 0 && releases[write - 1].0 == releases[read].0 {
            releases[write - 1].1 = releases[write - 1]
                .1
                .checked_add(releases[read].1)
                .ok_or(Error::InvalidSource { path })?;
        } else {
            releases[write] = releases[read];
            write += 1;
        }
    }
    releases.truncate(write);
    if !text_requests.is_empty() || !releases.is_empty() {
        if !target.storage.lists.string.segments.is_empty() {
            return Err(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::SharedString,
            });
        }
        let remaining = budget.remaining()?;
        let limits = string_list_limits(source, changes.len(), path, remaining)?;
        budget.authorize(remaining)?;
        let final_list =
            match lists::plan_string_list(string_list_payload, &releases, &text_requests, limits) {
                Ok(plan) => plan,
                Err(error) => {
                    budget.cancel_authorization();
                    return Err(map_list_error(error, path));
                },
            };
        let usage = list_usage(&budget, final_list.report())?;
        budget.record_authorized(usage)?;
        for assignment in final_list.assignments() {
            let change_index = *text_change_indices
                .get(assignment.request())
                .ok_or(Error::InvalidSource { path })?;
            if string_keys.get(change_index).copied().flatten() != Some(assignment.key()) {
                return Err(Error::Verification { path });
            }
        }
        prepared.push(PreparedReplacement {
            route: target.storage.lists.string.message,
            payload: final_list.into_payload(),
            references: None,
        });
    }

    prepared.sort_unstable_by_key(|replacement| {
        (
            replacement.route.component_index,
            replacement.route.object_index,
            replacement.route.message_index,
        )
    });
    let reference_route_count = prepared
        .iter()
        .filter(|replacement| replacement.references.is_some())
        .count();
    let reference_identifier_count = prepared.iter().try_fold(0usize, |count, replacement| {
        replacement
            .references
            .as_ref()
            .map_or(Ok(count), |references| {
                count
                    .checked_add(references.before.len())
                    .and_then(|count| count.checked_add(references.after.len()))
                    .ok_or(Error::InvalidSource { path })
            })
    })?;
    let rewrite_remaining = authorize_remaining(&mut budget)?;
    let mut rewrite_observed = classic_staging_usage(
        prepared.len(),
        reference_route_count,
        reference_identifier_count,
        path,
    )?;
    preflight_usage(rewrite_observed, rewrite_remaining, path)?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(prepared.len())
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: prepared.len(),
        })?;
    for replacement in &prepared {
        replacements.push(rewrite::MessageReplacement {
            component_index: replacement.route.component_index,
            object_index: replacement.route.object_index,
            message_index: replacement.route.message_index,
            expected_type: replacement.route.message_type,
            payload: &replacement.payload,
            references: replacement.references.as_ref().map(|references| {
                rewrite::AggregateReferenceDelta {
                    before: &references.before,
                    after: &references.after,
                }
            }),
        });
    }
    let previews =
        rewrite::root_preview_deletions(source).map_err(|error| map_rewrite_error(error, path))?;
    let preview_membership = preview_mask(&previews, path)?;
    let mut reference_routes = Vec::new();
    reference_routes
        .try_reserve_exact(reference_route_count)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: reference_route_count,
        })?;
    let reference_fields: Vec<FieldReferenceRoute> = Vec::new();
    let mut reference_identifiers = Vec::new();
    reference_identifiers
        .try_reserve_exact(reference_identifier_count)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: reference_identifier_count,
        })?;
    let mut directional_messages = Vec::new();
    directional_messages
        .try_reserve_exact(prepared.len())
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: prepared.len(),
        })?;
    for replacement in &prepared {
        let location = PhysicalLocation {
            component: replacement.route.component_index,
            object: replacement.route.object_index,
            message: replacement.route.message_index,
        };
        let mut message = DirectionalMessage::new(
            Some(location),
            Some(location),
            message_object_identifier(source, replacement.route, path)?,
            replacement.route.message_type,
            EvidenceChangeKind::Replace,
        );
        if let Some(references) = &replacement.references {
            let source = ReferenceSpan::new(reference_identifiers.len(), references.before.len());
            reference_identifiers.extend_from_slice(&references.before);
            let target = ReferenceSpan::new(reference_identifiers.len(), references.after.len());
            reference_identifiers.extend_from_slice(&references.after);
            let route = reference_routes.len();
            reference_routes.push(MessageReferenceRoute::new(
                source,
                target,
                ReferenceSpan::new(reference_fields.len(), 0),
            ));
            message = message.with_reference_transition(route);
        }
        directional_messages.push(message);
    }
    let reference_evidence = if reference_routes.is_empty() {
        None
    } else {
        Some(ReferenceEvidence::new(
            Arc::new(reference_routes),
            Arc::new(reference_fields),
            Arc::new(reference_identifiers),
            path,
        )?)
    };
    if reference_evidence.as_ref().is_some_and(|evidence| {
        evidence.allocation_shapes()
            != (
                (reference_route_count, reference_route_count),
                (0, 0),
                (reference_identifier_count, reference_identifier_count),
            )
    }) {
        budget.cancel_authorization();
        return Err(Error::Verification { path });
    }
    let directional_messages = Arc::new(directional_messages);
    let precharge_error = Cell::new(None);
    let component_observed = Cell::new(budget::Usage::default());
    let outcome = match rewrite::rewrite_with_evidence_authorization(
        source,
        rewrite::RewritePlan {
            replacements: &replacements,
            preview_deletions: &previews,
        },
        rewrite::EvidenceRetention::Omit,
        |component_reservation| {
            let envelope = match component_reservation_usage(component_reservation, path) {
                Ok(usage) => usage,
                Err(error) => {
                    precharge_error.set(Some(error));
                    return Err(rewrite::RewriteError::Precharge);
                },
            };
            let mut combined = rewrite_observed;
            if let Err(error) = merge_usage(&mut combined, envelope, path)
                .and_then(|()| preflight_usage(combined, rewrite_remaining, path))
            {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            Ok(())
        },
        |reservation, cost| {
            if let Err(error) = budget.preauthorize_publication(reservation) {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            let usage = match component_usage(cost, path) {
                Ok(usage) => usage,
                Err(error) => {
                    precharge_error.set(Some(error));
                    return Err(rewrite::RewriteError::Precharge);
                },
            };
            let mut combined = rewrite_observed;
            if let Err(error) = merge_usage(&mut combined, usage, path)
                .and_then(|()| preflight_usage(combined, budget.authorization_remaining()?, path))
            {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            component_observed.set(usage);
            Ok(())
        },
    ) {
        Ok(outcome) => outcome,
        Err(error) => {
            budget.cancel_publication();
            budget.cancel_authorization();
            return Err(precharge_error
                .get()
                .unwrap_or_else(|| map_rewrite_error(error, path)));
        },
    };
    let evidence = match PatchEvidence::new(
        directional_messages,
        reference_evidence,
        preview_membership,
        previews.len(),
        0,
        path,
    ) {
        Ok(evidence) => evidence,
        Err(error) => {
            budget.cancel_publication();
            budget.cancel_authorization();
            return Err(error);
        },
    };
    let locality_report = match verify_evidence_locality_with_report(
        source,
        &outcome.package,
        &evidence,
        outcome.publication.locality_work,
        path,
    ) {
        Ok(report) => report,
        Err(error) => {
            budget.cancel_publication();
            budget.cancel_authorization();
            return Err(error);
        },
    };
    if locality_report.bytes > outcome.publication.locality_bytes
        || locality_report.work > outcome.publication.locality_work
    {
        budget.cancel_publication();
        budget.cancel_authorization();
        return Err(Error::Verification { path });
    }
    if let Err(error) = merge_usage(&mut rewrite_observed, component_observed.get(), path)
        .and_then(|()| merge_usage(&mut rewrite_observed, locality_report.usage(), path))
    {
        budget.cancel_publication();
        budget.cancel_authorization();
        return Err(error);
    }
    let mut publication = outcome.publication;
    publication.locality_bytes = 0;
    publication.locality_work = 0;
    budget.finish_publication(publication);
    if let Err(error) = preflight_usage(rewrite_observed, budget.authorization_remaining()?, path)
        .and_then(|()| budget.record_authorized(rewrite_observed))
    {
        budget.cancel_authorization();
        return Err(error);
    }
    verify_semantic_changes(
        &outcome.package,
        sheet_position,
        table_position,
        &changes,
        path,
    )?;
    let changed_cells = changes.len();
    for change in changes {
        let (_position, input) = change.into_parts();
        if let Some(input) = input {
            drop(input.into_value());
        }
    }

    let source_bytes = Arc::clone(&source.state.source);
    let target_bytes = Arc::clone(&outcome.package.state.source);
    let patch = Patch::from_exact_with_evidence(
        path,
        requested,
        changed_cells,
        source_bytes,
        target_bytes,
        source.snapshot(),
        outcome.package.snapshot(),
        evidence,
    )?;
    Ok(Commit::new(
        outcome.package,
        patch,
        Diagnostics::from_changed(
            requested,
            changed_cells,
            outcome.touched_components,
            cache_prepared.refreshed_hosts,
            previews.len(),
        ),
    ))
}

fn authorize_remaining(budget: &mut budget::TransactionBudget) -> Result<budget::Remaining, Error> {
    let remaining = budget.remaining()?;
    budget.authorize(remaining)?;
    Ok(remaining)
}

fn reserve_retained_vec<T>(
    budget: &mut budget::TransactionBudget,
    capacity: usize,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    let bytes = capacity
        .checked_mul(size_of::<T>())
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::RetainedBytes,
            observed: u64::MAX,
            maximum: budget.limits().max_retained_bytes,
            path,
        })?;
    budget.reserve_retained(
        usize_u64(capacity),
        usize_u64(bytes),
        u64::from(capacity != 0),
    )
}

fn evidence_locality_plan_usage(
    messages: usize,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let bytes = messages
        .checked_mul(size_of::<locality::DirectionalMessage<'_>>())
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        peak_scratch_bytes: usize_u64(bytes),
        allocation_events: u64::from(messages != 0),
        transaction_work: usize_u64(messages),
        ..budget::Usage::default()
    })
}

fn classic_staging_usage(
    messages: usize,
    reference_routes: usize,
    reference_identifiers: usize,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let mut retained = budget::arc_vec_retained_usage::<DirectionalMessage>(messages, messages)?;
    if reference_routes != 0 {
        merge_usage(
            &mut retained,
            budget::arc_vec_retained_usage::<MessageReferenceRoute>(
                reference_routes,
                reference_routes,
            )?,
            path,
        )?;
        merge_usage(
            &mut retained,
            budget::arc_vec_retained_usage::<FieldReferenceRoute>(0, 0)?,
            path,
        )?;
        merge_usage(
            &mut retained,
            budget::arc_vec_retained_usage::<u64>(reference_identifiers, reference_identifiers)?,
            path,
        )?;
        retained.references = usize_u64(reference_identifiers);
    }
    let scratch_bytes = messages
        .checked_mul(size_of::<rewrite::MessageReplacement<'_>>())
        .and_then(|bytes| {
            messages
                .checked_mul(size_of::<locality::DirectionalMessage<'_>>())
                .and_then(|locality| bytes.checked_add(locality))
        })
        .ok_or(Error::InvalidSource { path })?;
    let work = messages
        .checked_mul(8)
        .ok_or(Error::InvalidSource { path })?;
    let mut result = budget::Usage {
        peak_scratch_bytes: usize_u64(scratch_bytes),
        allocation_events: 2,
        transaction_work: usize_u64(work),
        ..budget::Usage::default()
    };
    merge_usage(&mut result, retained, path)?;
    Ok(result)
}

fn component_reservation_usage(
    reservation: rewrite::ComponentReservation,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let objects = reservation
        .appended_objects
        .checked_add(reservation.deleted_objects)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: reservation.maximum_retained_elements,
        retained_bytes: reservation.maximum_retained_evidence_bytes,
        peak_scratch_bytes: reservation.maximum_peak_bytes,
        allocation_events: reservation.maximum_allocation_events,
        wire_bytes: reservation.compressed_input_bytes,
        wire_work: reservation.work,
        objects,
        references: reservation.reference_items,
        component_encodes: reservation.components,
        transaction_work: reservation.work,
        ..budget::Usage::default()
    })
}

fn component_usage(
    cost: rewrite::ComponentCost,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let wire_work = [
        cost.compressed_input_bytes,
        cost.decoded_input_bytes,
        cost.serialized_output_bytes,
        cost.compressed_output_bytes,
    ]
    .into_iter()
    .try_fold(0u64, |total, value| {
        total
            .checked_add(value)
            .ok_or(Error::InvalidSource { path })
    })?;
    let objects = cost
        .appended_objects
        .checked_add(cost.deleted_objects)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: cost.retained_elements,
        retained_bytes: cost.retained_evidence_bytes,
        peak_scratch_bytes: cost.peak_scratch_bytes,
        allocation_events: cost.allocation_events,
        wire_bytes: cost.compressed_input_bytes,
        wire_work,
        objects,
        references: cost.reference_items,
        component_encodes: cost.components,
        transaction_work: cost.work,
        ..budget::Usage::default()
    })
}

fn merge_usage(
    base: &mut budget::Usage,
    delta: budget::Usage,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    macro_rules! add {
        ($field:ident) => {
            base.$field = base
                .$field
                .checked_add(delta.$field)
                .ok_or(Error::InvalidSource { path })?;
        };
    }
    add!(retained_elements);
    add!(retained_bytes);
    add!(allocation_events);
    add!(wire_bytes);
    add!(wire_fields);
    add!(wire_work);
    add!(objects);
    add!(references);
    add!(lookups);
    add!(component_encodes);
    add!(locality_bytes);
    add!(transaction_work);
    base.peak_scratch_bytes = base.peak_scratch_bytes.max(delta.peak_scratch_bytes);
    Ok(())
}

fn preflight_usage(
    usage: budget::Usage,
    remaining: budget::Remaining,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    macro_rules! require {
        ($field:ident, $kind:expr) => {
            if usage.$field > remaining.$field {
                return Err(Error::LimitExceeded {
                    kind: $kind,
                    observed: usage.$field,
                    maximum: remaining.$field,
                    path,
                });
            }
        };
    }
    require!(
        retained_elements,
        crate::package::table_cells::LimitKind::RetainedElements
    );
    require!(
        retained_bytes,
        crate::package::table_cells::LimitKind::RetainedBytes
    );
    require!(
        peak_scratch_bytes,
        crate::package::table_cells::LimitKind::PeakScratchBytes
    );
    require!(
        allocation_events,
        crate::package::table_cells::LimitKind::TransactionWork
    );
    require!(objects, crate::package::table_cells::LimitKind::Objects);
    require!(
        references,
        crate::package::table_cells::LimitKind::References
    );
    require!(
        locality_bytes,
        crate::package::table_cells::LimitKind::TransactionWork
    );
    require!(
        transaction_work,
        crate::package::table_cells::LimitKind::TransactionWork
    );
    Ok(())
}

fn caller_staging_usage(
    budget: &budget::TransactionBudget,
    updates: usize,
    owned_value_bytes: usize,
    staging_usage: crate::table::cells::StagingUsage,
) -> Result<budget::Usage, Error> {
    let retained_elements = staging_usage
        .change_capacity()
        .checked_add(updates)
        .and_then(|elements| elements.checked_add(updates))
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            observed: u64::MAX,
            maximum: budget.limits().max_retained_elements,
            path: crate::package::table_cells::Path::Package,
        })?;
    let retained_bytes = staging_usage
        .change_capacity()
        .checked_mul(size_of::<crate::table::cells::Change>())
        .and_then(|bytes| {
            updates
                .checked_mul(
                    size_of::<crate::table::cells::Change>()
                        .checked_add(size_of::<crate::table::CellPosition>())?,
                )
                .and_then(|live| bytes.checked_add(live))
        })
        .and_then(|bytes| bytes.checked_add(owned_value_bytes))
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::RetainedBytes,
            observed: u64::MAX,
            maximum: budget.limits().max_retained_bytes,
            path: crate::package::table_cells::Path::Package,
        })?;
    Ok(budget::Usage {
        updates: usize_u64(updates),
        input_value_bytes: usize_u64(owned_value_bytes),
        retained_elements: usize_u64(retained_elements),
        retained_bytes: usize_u64(retained_bytes),
        allocation_events: usize_u64(staging_usage.allocation_events().checked_add(2).ok_or(
            Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: budget.limits().max_allocation_events,
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        ..budget::Usage::default()
    })
}

fn resolve_usage(
    budget: &budget::TransactionBudget,
    report: resolve::ResolveReport,
) -> Result<budget::Usage, Error> {
    let path = crate::package::table_cells::Path::Package;
    let ownership_work = report
        .ownership
        .work
        .checked_add(report.ownership.transaction_work)
        .and_then(|work| work.checked_add(report.codecs.work_bytes))
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: budget.limits().max_transaction_work,
            path,
        })?;
    Ok(budget::Usage {
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocation_events),
        wire_bytes: usize_u64(report.codecs.source_bytes),
        wire_fields: usize_u64(report.codecs.fields),
        wire_work: usize_u64(report.codecs.work_bytes),
        references: usize_u64(
            report
                .codecs
                .references
                .checked_add(report.ownership.references)
                .ok_or(Error::LimitExceeded {
                    kind: crate::package::table_cells::LimitKind::References,
                    observed: u64::MAX,
                    maximum: budget.limits().max_references,
                    path,
                })?,
        ),
        transaction_work: usize_u64(ownership_work),
        ..budget::Usage::default()
    })
}

fn tile_usage(
    budget: &budget::TransactionBudget,
    report: tile::TileReport,
) -> Result<budget::Usage, Error> {
    Ok(budget::Usage {
        wire_bytes: report.wire_bytes,
        wire_fields: report.wire_fields,
        wire_work: report.wire_work,
        tile_reads: 1,
        tile_writes: u64::from(report.output_bytes != 0),
        row_reads: report.rows_read,
        row_writes: report.rows_written,
        retained_elements: report.retained_elements,
        retained_bytes: report.retained_bytes,
        peak_scratch_bytes: report.peak_scratch_bytes,
        allocation_events: report.allocation_events,
        transaction_work: report
            .wire_work
            .checked_add(report.cell_slots_scanned)
            .and_then(|work| work.checked_add(report.cell_slots_written))
            .and_then(|work| work.checked_add(report.output_bytes))
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: budget.limits().max_transaction_work,
                path: crate::package::table_cells::Path::Package,
            })?,
        ..budget::Usage::default()
    })
}

fn message_payload(
    source: &Package,
    route: resolve::MessageRoute,
    path: crate::package::table_cells::Path,
) -> Result<&[u8], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == route.message_type)
        .map(|message| message.data.as_slice())
        .ok_or(Error::InvalidSource { path })
}

fn message_object_identifier(
    source: &Package,
    route: resolve::MessageRoute,
    path: crate::package::table_cells::Path,
) -> Result<u64, Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .filter(|object| {
            object
                .messages
                .get(route.message_index)
                .is_some_and(|message| message.type_ == route.message_type)
        })
        .and_then(|object| object.archive_info.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource { path })
}

fn rich_archive_reference_delta(
    source: &Package,
    route: resolve::MessageRoute,
    rich: &rich::ReferenceDelta,
    budget: &mut budget::TransactionBudget,
    path: crate::package::table_cells::Path,
) -> Result<PreparedReferenceDelta, Error> {
    if !rich.removed_by_field.is_empty()
        || rich.removed.windows(2).any(|pair| pair[0] >= pair[1])
        || rich.removed.iter().any(|identifier| {
            rich.before
                .iter()
                .filter(|candidate| *candidate == identifier)
                .count()
                != 1
                || rich.after.contains(identifier)
        })
    {
        return Err(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::RichText,
        });
    }
    let info = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.archive_info.message_infos.get(route.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if !info.field_infos.is_empty() {
        return Err(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::RichText,
        });
    }
    for identifier in &rich.removed {
        if info
            .object_references
            .iter()
            .filter(|candidate| *candidate == identifier)
            .count()
            != 1
        {
            return Err(Error::InvalidSource { path });
        }
    }
    reserve_retained_vec::<u64>(budget, info.object_references.len(), path)?;
    let mut before = Vec::new();
    before
        .try_reserve_exact(info.object_references.len())
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::References,
            amount: info.object_references.len(),
        })?;
    before.extend_from_slice(&info.object_references);
    let after_len = before
        .len()
        .checked_sub(rich.removed.len())
        .ok_or(Error::InvalidSource { path })?;
    reserve_retained_vec::<u64>(budget, after_len, path)?;
    let mut after = Vec::new();
    after
        .try_reserve_exact(after_len)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::References,
            amount: after_len,
        })?;
    for identifier in &before {
        if !rich.removed.contains(identifier) {
            after.push(*identifier);
        }
    }
    if after.len() != after_len {
        return Err(Error::InvalidSource { path });
    }
    Ok(PreparedReferenceDelta { before, after })
}

fn scalar_change(
    change: &crate::table::cells::Change,
    string_key: Option<u32>,
    rich_key: Option<u32>,
    path: crate::package::table_cells::Path,
) -> Result<tile::BncChange, Error> {
    Ok(match change.input_ref() {
        None => tile::BncChange::Clear,
        Some(Input::Number(value)) => tile::BncChange::Set(tile::ScalarInput::Number(*value)),
        Some(Input::Boolean(value)) => tile::BncChange::Set(tile::ScalarInput::Boolean(*value)),
        Some(Input::Date(value)) => tile::BncChange::Set(tile::ScalarInput::Date(*value)),
        Some(Input::Duration(value)) => tile::BncChange::Set(tile::ScalarInput::Duration(*value)),
        Some(Input::Text(_)) => tile::BncChange::Set(match rich_key {
            Some(key) => tile::ScalarInput::RichText(key),
            None => tile::ScalarInput::String(string_key.ok_or(Error::InvalidSource { path })?),
        }),
    })
}

fn string_list_limits(
    source: &Package,
    requests: usize,
    path: crate::package::table_cells::Path,
    remaining: budget::Remaining,
) -> Result<lists::ListLimits, Error> {
    let package_wire = source.state.options.archive().max_iwa_stream_bytes();
    let package_work = package_wire.checked_mul(32).ok_or(Error::LimitExceeded {
        kind: crate::package::table_cells::LimitKind::WireWork,
        observed: u64::MAX,
        maximum: usize_u64(package_wire),
        path,
    })?;
    let maximum_wire = usize::try_from(remaining.wire_bytes)
        .map_err(|_error| Error::InvalidSource { path })?
        .min(package_wire);
    let maximum_fields = usize::try_from(remaining.wire_fields)
        .map_err(|_error| Error::InvalidSource { path })?
        .min(package_wire);
    let maximum_work = usize::try_from(remaining.wire_work.min(remaining.transaction_work))
        .map_err(|_error| Error::InvalidSource { path })?
        .min(package_work);
    let maximum_references = usize::try_from(remaining.references)
        .map_err(|_error| Error::InvalidSource { path })?
        .min(source.state.options.semantic().max_references());
    let maximum_output =
        usize::try_from(remaining.retained_bytes.min(remaining.peak_scratch_bytes))
            .map_err(|_error| Error::InvalidSource { path })?
            .min(package_wire);
    let semantic = source.state.options.semantic();
    Ok(lists::ListLimits::new(
        litchi_iwa_protos::numbers_table_cell_storage_codec::DecodeOptions::new(
            maximum_wire,
            maximum_fields,
            maximum_work,
            64,
            maximum_references,
            semantic.max_output_text_bytes(),
        ),
        maximum_references,
        requests,
        maximum_output,
    )
    .with_accounting(
        usize::try_from(remaining.retained_bytes)
            .map_err(|_error| Error::InvalidSource { path })?,
        usize::try_from(remaining.peak_scratch_bytes)
            .map_err(|_error| Error::InvalidSource { path })?,
        usize::try_from(remaining.allocation_events)
            .map_err(|_error| Error::InvalidSource { path })?,
        usize::try_from(remaining.transaction_work)
            .map_err(|_error| Error::InvalidSource { path })?,
    ))
}

fn list_usage(
    budget: &budget::TransactionBudget,
    report: lists::ListReport,
) -> Result<budget::Usage, Error> {
    let root = report.decode();
    let entries = report.entry_decode();
    Ok(budget::Usage {
        wire_bytes: usize_u64(
            root.source_bytes()
                .checked_add(entries.source_bytes())
                .ok_or(Error::LimitExceeded {
                    kind: crate::package::table_cells::LimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: budget.limits().max_wire_bytes,
                    path: crate::package::table_cells::Path::Package,
                })?,
        ),
        wire_fields: usize_u64(root.fields().checked_add(entries.fields()).ok_or(
            Error::InvalidSource {
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        wire_work: usize_u64(root.work_bytes().checked_add(entries.work_bytes()).ok_or(
            Error::InvalidSource {
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        references: usize_u64(root.references().checked_add(entries.references()).ok_or(
            Error::InvalidSource {
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        list_reads: 1,
        list_writes: u64::from(report.output_bytes() != report.input_bytes()),
        string_work: usize_u64(
            report
                .entries_scanned()
                .checked_add(report.strings_reused())
                .and_then(|value| value.checked_add(report.strings_added()))
                .ok_or(Error::InvalidSource {
                    path: crate::package::table_cells::Path::Package,
                })?,
        ),
        retained_bytes: usize_u64(report.retained_bytes()),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes()),
        allocation_events: usize_u64(report.allocations()),
        transaction_work: usize_u64(report.transaction_work()),
        ..budget::Usage::default()
    })
}

fn assignment_usage(
    budget: &budget::TransactionBudget,
    report: lists::AssignmentReport,
) -> Result<budget::Usage, Error> {
    let root = report.decode();
    let entries = report.entry_decode();
    Ok(budget::Usage {
        wire_bytes: usize_u64(
            root.source_bytes()
                .checked_add(entries.source_bytes())
                .ok_or(Error::LimitExceeded {
                    kind: crate::package::table_cells::LimitKind::WireWork,
                    observed: u64::MAX,
                    maximum: budget.limits().max_wire_bytes,
                    path: crate::package::table_cells::Path::Package,
                })?,
        ),
        wire_fields: usize_u64(root.fields().checked_add(entries.fields()).ok_or(
            Error::InvalidSource {
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        wire_work: usize_u64(root.work_bytes().checked_add(entries.work_bytes()).ok_or(
            Error::InvalidSource {
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        references: usize_u64(root.references().checked_add(entries.references()).ok_or(
            Error::InvalidSource {
                path: crate::package::table_cells::Path::Package,
            },
        )?),
        list_reads: 1,
        string_work: usize_u64(
            report
                .entries_scanned()
                .checked_add(report.requests())
                .and_then(|value| value.checked_add(report.unique_requests()))
                .ok_or(Error::InvalidSource {
                    path: crate::package::table_cells::Path::Package,
                })?,
        ),
        retained_elements: usize_u64(report.requests()),
        retained_bytes: usize_u64(report.retained_bytes()),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes()),
        allocation_events: usize_u64(report.allocations()),
        transaction_work: usize_u64(report.transaction_work()),
        ..budget::Usage::default()
    })
}

fn map_list_error(error: lists::Failure, path: crate::package::table_cells::Path) -> Error {
    match error {
        lists::Failure::Allocation { amount, .. } => Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedBytes,
            amount,
        },
        lists::Failure::LimitExceeded {
            observed, maximum, ..
        } => Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::TransactionWork,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
            path,
        },
        lists::Failure::Decode(_)
        | lists::Failure::Wire(_)
        | lists::Failure::InvalidSource(_)
        | lists::Failure::Overflow(_) => Error::InvalidSource { path },
    }
}

fn map_tile_error(error: tile::TileError, path: crate::package::table_cells::Path) -> Error {
    match error {
        tile::TileError::NeedSparse { .. } => Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::CellStorage,
        },
        tile::TileError::Allocation { amount } => Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount,
        },
        tile::TileError::LimitExceeded { observed, maximum } => Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::WireWork,
            observed,
            maximum,
            path,
        },
        tile::TileError::UnsupportedValue { .. } => Error::UnsupportedSource { path },
        tile::TileError::UnsupportedSource { .. } => Error::UnsupportedSource { path },
        tile::TileError::InvalidSource
        | tile::TileError::DuplicateOrUnsortedChange { .. }
        | tile::TileError::OutOfBounds { .. } => Error::InvalidSource { path },
    }
}

fn map_rewrite_error(
    error: rewrite::RewriteError,
    path: crate::package::table_cells::Path,
) -> Error {
    match error {
        rewrite::RewriteError::UnsupportedSource => Error::UnsupportedSource { path },
        rewrite::RewriteError::InvalidSource => Error::InvalidSource { path },
        rewrite::RewriteError::Allocation { amount } => Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount,
        },
        rewrite::RewriteError::Limit => Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path,
        },
        rewrite::RewriteError::Verification
        | rewrite::RewriteError::Precharge
        | rewrite::RewriteError::Candidate => Error::Verification { path },
    }
}

#[cfg(test)]
fn verify_evidence_locality(
    source: &Package,
    candidate: &Package,
    evidence: &PatchEvidence,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    let max_work = source
        .state
        .options
        .archive()
        .max_input_bytes()
        .checked_mul(64)
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::TransactionWork,
            observed: u64::MAX,
            maximum: source.state.options.archive().max_input_bytes(),
            path,
        })?;
    verify_evidence_locality_with_report(source, candidate, evidence, max_work, path)
        .map(|_report| ())
}

fn verify_evidence_locality_with_report(
    source: &Package,
    candidate: &Package,
    evidence: &PatchEvidence,
    max_work: u64,
    path: crate::package::table_cells::Path,
) -> Result<locality::Report, Error> {
    let count = evidence.message_count();
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(count)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: count,
        })?;
    for message in evidence.directional_messages() {
        if let Some(location) = message.source() {
            validate_evidence_endpoint(
                source,
                location,
                message.object_identifier(),
                message.expected_type(),
                path,
            )?;
        }
        let target_payload = match message.target() {
            Some(location) => {
                validate_evidence_endpoint(
                    candidate,
                    location,
                    message.object_identifier(),
                    message.expected_type(),
                    path,
                )?;
                Some(message_payload(
                    candidate,
                    resolve::MessageRoute {
                        component_index: location.component,
                        object_index: location.object,
                        message_index: location.message,
                        message_type: message.expected_type(),
                    },
                    path,
                )?)
            },
            None => None,
        };
        let mut locality_message = locality::DirectionalMessage::new(
            message.source().map(locality_message_location),
            message.target().map(locality_message_location),
            message.object_identifier(),
            message.expected_type(),
            match message.kind() {
                EvidenceChangeKind::Replace => locality::DirectionalChange::Replace,
                EvidenceChangeKind::Append => locality::DirectionalChange::Append,
                EvidenceChangeKind::Delete => locality::DirectionalChange::Delete,
            },
            target_payload,
        );
        if let Some(references) = evidence.reference_transition(message) {
            locality_message = locality_message.with_reference_transition(references);
        }
        messages.push(locality_message);
    }
    let source_catalog = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource { path })?;
    let candidate_catalog = candidate
        .state
        .components
        .physical()
        .ok_or(Error::Verification { path })?;
    let source_mask = if evidence.source_previews() == 0 {
        0
    } else {
        evidence.preview_mask()
    };
    let target_mask = if evidence.target_previews() == 0 {
        0
    } else {
        evidence.preview_mask()
    };
    locality::verify_directional(
        source_catalog,
        candidate_catalog,
        locality::DirectionalPlan {
            messages: &messages,
            previews: locality::PreviewMask {
                names: &rewrite::ROOT_PREVIEWS,
                source_mask,
                target_mask,
            },
        },
        locality::Limits { max_work },
    )
    .map_err(|_error| Error::Verification { path })
}

fn locality_message_location(location: PhysicalLocation) -> locality::MessageLocation {
    locality::MessageLocation::new(
        locality::ObjectLocation::new(
            locality::ComponentLocation::new(location.component),
            location.object,
        ),
        location.message,
    )
}

fn validate_evidence_endpoint(
    package: &Package,
    location: PhysicalLocation,
    object_identifier: u64,
    expected_type: u32,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    let object = package
        .state
        .components
        .catalog()
        .get_index(location.component)
        .and_then(|component| component.archive().objects.get(location.object))
        .filter(|object| object.archive_info.identifier == Some(object_identifier))
        .ok_or(Error::Verification { path })?;
    object
        .messages
        .get(location.message)
        .filter(|message| message.type_ == expected_type)
        .map(|_message| ())
        .ok_or(Error::Verification { path })
}

fn preview_mask(previews: &[&str], path: crate::package::table_cells::Path) -> Result<u8, Error> {
    let mut mask = 0u8;
    for preview in previews {
        let index = rewrite::ROOT_PREVIEWS
            .iter()
            .position(|candidate| candidate == preview)
            .ok_or(Error::InvalidSource { path })?;
        let bit = 1u8
            .checked_shl(u32::try_from(index).map_err(|_error| Error::InvalidSource { path })?)
            .ok_or(Error::InvalidSource { path })?;
        if mask & bit != 0 {
            return Err(Error::InvalidSource { path });
        }
        mask |= bit;
    }
    Ok(mask)
}

fn distinct_evidence_components(
    evidence: &PatchEvidence,
    path: crate::package::table_cells::Path,
) -> Result<usize, Error> {
    let mut count = 0usize;
    let mut previous = None;
    for index in 0..evidence.message_count() {
        let message = evidence
            .directional_message(index)
            .ok_or(Error::Verification { path })?;
        let location = message
            .source()
            .or_else(|| message.target())
            .ok_or(Error::Verification { path })?;
        if previous.is_some_and(|value| value > location.component) {
            return Err(Error::Verification { path });
        }
        if previous != Some(location.component) {
            count = count.checked_add(1).ok_or(Error::Verification { path })?;
            previous = Some(location.component);
        }
    }
    Ok(count)
}

fn verify_semantic_changes(
    candidate: &Package,
    sheet_position: usize,
    table_position: usize,
    changes: &[crate::table::cells::Change],
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    let selected = resolve_table(
        candidate,
        SheetSelector::Index(sheet_position),
        TableSelector::Index(table_position),
    )?;
    for change in changes {
        let valid = match change.input_ref() {
            Some(input) => match selected.table.view(change.position()) {
                View::Stored(value) => input.matches_value(value),
                View::Missing | View::Covered => false,
            },
            None => matches!(
                selected.table.view(change.position()),
                View::Missing | View::Stored(crate::cell::Value::Empty)
            ),
        };
        if !valid {
            return Err(Error::Verification { path });
        }
    }
    Ok(())
}

const fn usize_u64(value: usize) -> u64 {
    value as u64
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use litchi_iwa_common::wire::{WireView, patch_nested_varint_field, patch_varint_field};
    use litchi_iwa_core::{ArchiveObject, RawMessage};
    use litchi_iwa_protos::{tsce, tsp, tst, tswp};
    use prost::Message as _;

    use crate::{
        Package,
        cell::{
            Value,
            wire::{BncCell, CachedScalar},
        },
        package::table_cells::{DependencyKind, Error, Path},
        table::{CellPosition, cells::Input},
    };

    use super::{PhysicalLocation, preview_mask, rewrite, verify_evidence_locality};

    type TestResult = Result<(), Box<dyn std::error::Error>>;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers")
    }

    fn synthetic_513_row_source() -> Package {
        let source = Package::open(fixture()).expect("native basic fixture opens");
        let target = crate::package::table_headers::resolve::resolve_target(&source, 0, 0)
            .expect("basic table resolves");
        assert_eq!(target.rows, 22, "seed source row count is frozen");
        let component = source
            .state
            .components
            .catalog()
            .get_index(target.component_index)
            .expect("model component exists");
        let object = component
            .archive()
            .objects
            .get(target.object_index)
            .expect("model object exists");
        let object_identifier = object
            .archive_info
            .identifier
            .expect("model object has identity");
        let message = object
            .messages
            .get(target.message_index)
            .expect("model message exists");
        let payload = patch_varint_field(&message.data, 6, true, Some(513))
            .expect("raw-preserving row-count seed patch");
        let (resolved, _report) =
            super::resolve::resolve_changed_target(&source, 0, 0, &[CellPosition::new(2, 1)])
                .expect("original body coordinate resolves dependencies");
        let owner_route = resolved
            .dependencies
            .formula_owners
            .iter()
            .copied()
            .find(|route| {
                let Ok(payload) = super::message_payload(&source, *route, Path::Package) else {
                    return false;
                };
                let Ok(view) = WireView::parse(payload) else {
                    return false;
                };
                let mut owner = None;
                for field in view.fields() {
                    if field.number() == 11 && field.wire_type() == 2 {
                        if field.validate_canonical_framing().is_err() || owner.is_some() {
                            return false;
                        }
                        owner = crate::package::table_headers::resolve::local_reference_identifier(
                            field.payload(),
                        )
                        .ok();
                    }
                }
                owner == Some(target.drawable_identifier)
            })
            .expect("selected table formula owner is unique");
        let mut owner_payload = super::message_payload(&source, owner_route, Path::Package)
            .expect("selected formula-owner payload")
            .to_vec();
        for path in [[7, 2, 4], [7, 3, 4], [8, 2, 4], [8, 3, 4]] {
            owner_payload = patch_nested_varint_field(&owner_payload, &path, true, Some(512))
                .expect("raw-preserving spanning extent seed patch");
        }
        let mut replacements = vec![
            rewrite::MessageReplacement {
                component_index: target.component_index,
                object_index: target.object_index,
                message_index: target.message_index,
                expected_type: target.message_type,
                payload: &payload,
                references: None,
            },
            rewrite::MessageReplacement {
                component_index: owner_route.component_index,
                object_index: owner_route.object_index,
                message_index: owner_route.message_index,
                expected_type: owner_route.message_type,
                payload: &owner_payload,
                references: None,
            },
        ];
        replacements.sort_unstable_by_key(|replacement| {
            (
                replacement.component_index,
                replacement.object_index,
                replacement.message_index,
            )
        });
        let previews = rewrite::root_preview_deletions(&source).expect("preview set resolves");
        let outcome = rewrite::rewrite_with_precharge(
            &source,
            rewrite::RewritePlan {
                replacements: &replacements,
                preview_deletions: &previews,
            },
            |_reservation| Ok(()),
        )
        .expect("seed component publishes");

        // The synthetic provenance is itself proved as an exact one-message,
        // no-preview topology change before it is used as the sparse source.
        let mut messages = Vec::new();
        messages
            .try_reserve_exact(outcome.published_messages.len())
            .expect("seed evidence reservation");
        for message in &outcome.published_messages {
            assert_eq!(message.kind, rewrite::PublishedMessageKind::Existing);
            messages.push(super::DirectionalMessage::new(
                message.source_object_index.map(|object| PhysicalLocation {
                    component: message.component_index,
                    object,
                    message: message.message_index,
                }),
                message.target_object_index.map(|object| PhysicalLocation {
                    component: message.component_index,
                    object,
                    message: message.message_index,
                }),
                message.object_identifier,
                message.expected_type,
                super::EvidenceChangeKind::Replace,
            ));
        }
        assert!(messages.iter().any(|message| {
            message.object_identifier() == object_identifier
                && message.expected_type() == target.message_type
        }));
        let messages = std::sync::Arc::new(messages);
        let evidence = super::PatchEvidence::new(
            messages,
            None,
            preview_mask(&previews, Path::Package).expect("preview mask"),
            previews.len(),
            0,
            Path::Package,
        )
        .expect("seed evidence is canonical");
        verify_evidence_locality(&source, &outcome.package, &evidence, Path::Package)
            .expect("seed patch locality is exact");
        outcome.package
    }

    fn synthetic_wide_source(columns: u32) -> Package {
        let source = Package::open(fixture()).expect("native basic fixture opens");
        let (target, _report) =
            super::resolve::resolve_changed_target(&source, 0, 0, &[CellPosition::new(1, 1)])
                .expect("basic body coordinate resolves");
        assert!(columns > target.native.columns);

        let model_route = target.storage.model;
        let mut model = tst::TableModelArchive::decode(
            super::message_payload(&source, model_route, Path::Package)
                .expect("source model payload"),
        )
        .expect("source model decodes");
        model.number_of_columns = columns;
        let model_payload = model.encode_to_vec();

        let tile_route = target.storage.tiles.first().expect("basic has tile zero");
        let mut tile = tst::Tile::decode(
            super::message_payload(&source, tile_route.message, Path::Package)
                .expect("source tile payload"),
        )
        .expect("source tile decodes");
        for row in &mut tile.row_infos {
            let wide = row.has_wide_offsets.unwrap_or(false);
            let mut cells = decoded_test_row(
                row,
                usize::try_from(target.native.columns).expect("source width fits"),
            );
            cells.resize(
                usize::try_from(columns).expect("synthetic width fits"),
                None,
            );
            let (storage, offsets) = encoded_test_row(&cells, wide);
            row.cell_storage_buffer = Some(storage);
            row.cell_offsets = Some(offsets);
        }
        tile.max_column = columns - 1;
        let tile_payload = tile.encode_to_vec();

        let header_route = target.storage.column_headers;
        let mut headers = tst::HeaderStorageBucket::decode(
            super::message_payload(&source, header_route, Path::Package)
                .expect("source column-header payload"),
        )
        .expect("source column headers decode");
        let template = *headers.headers.last().expect("source has a column header");
        headers
            .headers
            .try_reserve_exact(
                usize::try_from(columns - target.native.columns).expect("header growth fits"),
            )
            .expect("synthetic header reservation");
        for column in target.native.columns..columns {
            headers.headers.push(tst::header_storage_bucket::Header {
                index: column,
                number_of_cells: 0,
                ..template
            });
        }
        let header_payload = headers.encode_to_vec();

        let owner_route = target
            .dependencies
            .selected_formula_owner
            .expect("basic has selected formula owner")
            .message;
        let mut owner_payload = super::message_payload(&source, owner_route, Path::Package)
            .expect("selected formula-owner payload")
            .to_vec();
        for path in [[7, 2, 3], [7, 3, 3], [8, 2, 3], [8, 3, 3]] {
            owner_payload = patch_nested_varint_field(
                &owner_payload,
                &path,
                true,
                Some(u64::from(columns - 1)),
            )
            .expect("raw-preserving spanning width patch");
        }

        let mut replacements = [
            (model_route, model_payload),
            (tile_route.message, tile_payload),
            (header_route, header_payload),
            (owner_route, owner_payload),
        ];
        replacements.sort_unstable_by_key(|replacement| {
            (
                replacement.0.component_index,
                replacement.0.object_index,
                replacement.0.message_index,
            )
        });
        let replacements = replacements
            .iter()
            .map(|(route, payload)| rewrite::MessageReplacement {
                component_index: route.component_index,
                object_index: route.object_index,
                message_index: route.message_index,
                expected_type: route.message_type,
                payload,
                references: None,
            })
            .collect::<Vec<_>>();
        let previews = rewrite::root_preview_deletions(&source).expect("preview set resolves");
        rewrite::rewrite_with_precharge(
            &source,
            rewrite::RewritePlan {
                replacements: &replacements,
                preview_deletions: &previews,
            },
            |_reservation| Ok(()),
        )
        .expect("wide synthetic source publishes")
        .package
    }

    fn decoded_test_row(row: &tst::TileRowInfo, columns: usize) -> Vec<Option<Vec<u8>>> {
        let width = if row.has_wide_offsets.unwrap_or(false) {
            4
        } else {
            1
        };
        let offsets = row.cell_offsets.as_ref().expect("seed BNC offsets");
        let storage = row.cell_storage_buffer.as_ref().expect("seed BNC storage");
        let mut cells = vec![None; columns];
        for (column, cell) in cells.iter_mut().enumerate() {
            let index = column * 2;
            let offset = u16::from_le_bytes([offsets[index], offsets[index + 1]]);
            if offset == u16::MAX {
                continue;
            }
            let start = usize::from(offset) * width;
            let end = offsets[(index + 2)..]
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .find(|offset| *offset != u16::MAX)
                .map_or(storage.len(), |offset| usize::from(offset) * width);
            *cell = Some(storage[start..end].to_vec());
        }
        cells
    }

    fn encoded_test_row(cells: &[Option<Vec<u8>>], wide: bool) -> (Vec<u8>, Vec<u8>) {
        let width = if wide { 4 } else { 1 };
        let mut storage = Vec::new();
        let mut offsets = Vec::new();
        offsets.reserve_exact(cells.len() * 2);
        for cell in cells {
            let Some(cell) = cell else {
                offsets.extend_from_slice(&u16::MAX.to_le_bytes());
                continue;
            };
            while storage.len() % width != 0 {
                storage.push(0);
            }
            let offset = u16::try_from(storage.len() / width).expect("test row offset fits");
            assert_ne!(offset, u16::MAX);
            offsets.extend_from_slice(&offset.to_le_bytes());
            storage.extend_from_slice(cell);
        }
        while storage.len() % width != 0 {
            storage.push(0);
        }
        (storage, offsets)
    }

    fn absolute_reference_node(
        row: u32,
        column: u32,
    ) -> tsce::ast_node_array_archive::AstNodeArchive {
        use tsce::ast_node_array_archive::{
            AstColumnCoordinateArchive, AstNodeArchive, AstNodeType, AstRowCoordinateArchive,
        };
        AstNodeArchive {
            ast_node_type: AstNodeType::CellReferenceNode as i32,
            ast_column: Some(AstColumnCoordinateArchive {
                column: i32::try_from(column).expect("test column fits"),
                absolute: Some(true),
            }),
            ast_row: Some(AstRowCoordinateArchive {
                row: i32::try_from(row).expect("test row fits"),
                absolute: Some(true),
            }),
            ..Default::default()
        }
    }

    fn operator_node(
        kind: tsce::ast_node_array_archive::AstNodeType,
    ) -> tsce::ast_node_array_archive::AstNodeArchive {
        tsce::ast_node_array_archive::AstNodeArchive {
            ast_node_type: kind as i32,
            ..Default::default()
        }
    }

    fn number_node(value: f64) -> tsce::ast_node_array_archive::AstNodeArchive {
        tsce::ast_node_array_archive::AstNodeArchive {
            ast_node_type: tsce::ast_node_array_archive::AstNodeType::NumberNode as i32,
            ast_number_node_number: Some(value),
            ..Default::default()
        }
    }

    fn function_node(
        identifier: u32,
        arguments: u32,
    ) -> tsce::ast_node_array_archive::AstNodeArchive {
        tsce::ast_node_array_archive::AstNodeArchive {
            ast_node_type: tsce::ast_node_array_archive::AstNodeType::FunctionNode as i32,
            ast_function_node_index: Some(identifier),
            ast_function_node_num_args: Some(arguments),
            ..Default::default()
        }
    }

    fn formula(nodes: Vec<tsce::ast_node_array_archive::AstNodeArchive>) -> tsce::FormulaArchive {
        tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive { ast_node: nodes },
            ..Default::default()
        }
    }

    fn synthetic_formula_cache_source(unsupported_formula: bool) -> Package {
        let source = Package::open(fixture()).expect("native basic fixture opens");
        let (target, _report) =
            super::resolve::resolve_changed_target(&source, 0, 0, &[CellPosition::new(2, 1)])
                .expect("basic body coordinate resolves");
        let tile_route = target.storage.tiles.first().expect("basic has tile zero");
        let mut a2 = BncCell::minimal();
        a2.set_number(120.0).expect("finite test number");
        let mut b2 = BncCell::minimal();
        b2.set_number(323.0).expect("finite test cache");
        b2.set_formula_reference(1);
        let mut c2 = BncCell::minimal();
        c2.set_number(646.0).expect("finite test cache");
        c2.set_formula_reference(2);
        let mut a3 = BncCell::minimal();
        a3.set_number(203.0).expect("finite test number");
        let mut tile = tst::Tile::decode(
            super::message_payload(&source, tile_route.message, Path::Package)
                .expect("source tile payload"),
        )
        .expect("source tile decodes");
        let row_one = tile
            .row_infos
            .iter_mut()
            .find(|row| row.tile_row_index == 1)
            .expect("source row one exists");
        let row_one_wide = row_one.has_wide_offsets.unwrap_or(false);
        let mut row_one_cells = decoded_test_row(row_one, 7);
        assert!(row_one_cells[1].is_some());
        row_one_cells[1] = Some(a2.encode());
        row_one_cells[2] = Some(b2.encode());
        row_one_cells[3] = Some(c2.encode());
        let (row_one_storage, row_one_offsets) = encoded_test_row(&row_one_cells, row_one_wide);
        row_one.cell_count =
            u32::try_from(row_one_cells.iter().flatten().count()).expect("count fits");
        row_one.cell_storage_buffer = Some(row_one_storage);
        row_one.cell_offsets = Some(row_one_offsets);
        let row_two = tile
            .row_infos
            .iter_mut()
            .find(|row| row.tile_row_index == 2)
            .expect("source row two exists");
        let row_two_wide = row_two.has_wide_offsets.unwrap_or(false);
        let mut row_two_cells = decoded_test_row(row_two, 7);
        assert!(row_two_cells[1].is_some());
        row_two_cells[1] = Some(a3.encode());
        let (row_two_storage, row_two_offsets) = encoded_test_row(&row_two_cells, row_two_wide);
        row_two.cell_count =
            u32::try_from(row_two_cells.iter().flatten().count()).expect("count fits");
        row_two.cell_storage_buffer = Some(row_two_storage);
        row_two.cell_offsets = Some(row_two_offsets);
        let tile_payload = tile.encode_to_vec();
        let formula_one = if unsupported_formula {
            formula(vec![absolute_reference_node(1, 1), function_node(999, 1)])
        } else {
            formula(vec![
                absolute_reference_node(1, 1),
                absolute_reference_node(2, 1),
                operator_node(tsce::ast_node_array_archive::AstNodeType::AdditionNode),
            ])
        };
        let formula_two = formula(vec![
            absolute_reference_node(1, 2),
            number_node(2.0),
            operator_node(tsce::ast_node_array_archive::AstNodeType::MultiplicationNode),
        ]);
        let formula_payload = tst::TableDataList {
            list_type: tst::table_data_list::ListType::Formula as i32,
            next_list_id: 3,
            entries: vec![
                tst::table_data_list::ListEntry {
                    key: 1,
                    refcount: 1,
                    formula: Some(formula_one),
                    ..Default::default()
                },
                tst::table_data_list::ListEntry {
                    key: 2,
                    refcount: 1,
                    formula: Some(formula_two),
                    ..Default::default()
                },
            ],
            segments: Vec::new(),
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let owner_route = target
            .dependencies
            .selected_formula_owner
            .expect("basic has selected formula owner")
            .message;
        let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(
            super::message_payload(&source, owner_route, Path::Package).expect("owner payload"),
        )
        .expect("owner decodes");
        owner.cell_dependencies = Some(tsce::CellDependenciesExpandedArchive {
            cell_record: vec![
                tsce::CellRecordExpandedArchive {
                    row: 1,
                    column: 2,
                    expanded_edges: Some(tsce::ExpandedEdgesArchive {
                        edge_without_owner_rows: vec![1, 2],
                        edge_without_owner_columns: vec![1, 1],
                        ..Default::default()
                    }),
                    ..Default::default()
                },
                tsce::CellRecordExpandedArchive {
                    row: 1,
                    column: 3,
                    expanded_edges: Some(tsce::ExpandedEdgesArchive {
                        edge_without_owner_rows: vec![1],
                        edge_without_owner_columns: vec![2],
                        ..Default::default()
                    }),
                    ..Default::default()
                },
            ],
        });
        owner.spanning_column_dependencies = None;
        owner.spanning_row_dependencies = None;
        owner.tiled_cell_dependencies = Some(tsce::CellDependenciesTiledArchive::default());
        let owner_payload = owner.encode_to_vec();
        let engine_route = target.dependencies.engine.expect("basic has engine");
        let mut engine = tsce::CalculationEngineArchive::decode(
            super::message_payload(&source, engine_route, Path::Package).expect("engine payload"),
        )
        .expect("engine decodes");
        engine.dependency_tracker.number_of_formulas = Some(2);
        let engine_payload = engine.encode_to_vec();
        let mut replacements = [
            (tile_route.message, tile_payload),
            (target.storage.lists.formula.message, formula_payload),
            (owner_route, owner_payload),
            (engine_route, engine_payload),
        ];
        replacements.sort_unstable_by_key(|replacement| {
            (
                replacement.0.component_index,
                replacement.0.object_index,
                replacement.0.message_index,
            )
        });
        let replacements = replacements
            .iter()
            .map(|(route, payload)| rewrite::MessageReplacement {
                component_index: route.component_index,
                object_index: route.object_index,
                message_index: route.message_index,
                expected_type: route.message_type,
                payload,
                references: None,
            })
            .collect::<Vec<_>>();
        let previews = rewrite::root_preview_deletions(&source).expect("preview set resolves");
        rewrite::rewrite_with_precharge(
            &source,
            rewrite::RewritePlan {
                replacements: &replacements,
                preview_deletions: &previews,
            },
            |_reservation| Ok(()),
        )
        .expect("formula seed publishes")
        .package
    }

    fn synthetic_formula_fanout_source(formulas: usize) -> Package {
        const HOSTS_PER_ROW: usize = 4_096;
        let columns = u32::try_from(formulas.min(HOSTS_PER_ROW) + 8).expect("formula width fits");
        let source = synthetic_wide_source(columns);
        let source_position = CellPosition::new(1, 7);
        let (target, _report) =
            super::resolve::resolve_changed_target(&source, 0, 0, &[source_position])
                .expect("wide formula source resolves");

        let tile_route = target
            .storage
            .tiles
            .first()
            .expect("wide source has tile zero");
        let mut tile = tst::Tile::decode(
            super::message_payload(&source, tile_route.message, Path::Package)
                .expect("wide tile payload"),
        )
        .expect("wide tile decodes");
        tile.should_use_wide_rows = Some(true);
        let mut entries = Vec::new();
        let mut records = Vec::new();
        entries
            .try_reserve_exact(formulas)
            .expect("formula entries reserve");
        records
            .try_reserve_exact(formulas)
            .expect("formula records reserve");
        for host_row in 0..formulas.div_ceil(HOSTS_PER_ROW) {
            let row_index = u32::try_from(host_row + 1).expect("formula row fits");
            let row = tile
                .row_infos
                .iter_mut()
                .find(|row| row.tile_row_index == row_index)
                .expect("wide source body row exists");
            let wide = true;
            row.has_wide_offsets = Some(true);
            let mut cells = decoded_test_row(
                row,
                usize::try_from(columns).expect("formula width fits usize"),
            );
            if host_row == 0 {
                let mut source_cell = BncCell::minimal();
                source_cell.set_number(1.0).expect("finite formula source");
                cells[7] = Some(source_cell.encode());
            }
            let start = host_row * HOSTS_PER_ROW;
            let end = formulas.min(start + HOSTS_PER_ROW);
            for index in start..end {
                let key = u32::try_from(index + 1).expect("formula key fits");
                let column = u32::try_from(index - start + 8).expect("formula column fits");
                let mut host = BncCell::minimal();
                host.set_number(2.0).expect("finite formula cache");
                host.set_formula_reference(key);
                cells[usize::try_from(column).expect("formula column fits usize")] =
                    Some(host.encode());
                entries.push(tst::table_data_list::ListEntry {
                    key,
                    refcount: 1,
                    formula: Some(formula(vec![
                        absolute_reference_node(1, 7),
                        number_node(1.0),
                        operator_node(tsce::ast_node_array_archive::AstNodeType::AdditionNode),
                    ])),
                    ..Default::default()
                });
                records.push(tsce::CellRecordExpandedArchive {
                    row: row_index,
                    column,
                    expanded_edges: Some(tsce::ExpandedEdgesArchive {
                        edge_without_owner_rows: vec![1],
                        edge_without_owner_columns: vec![7],
                        ..Default::default()
                    }),
                    ..Default::default()
                });
            }
            let (storage, offsets) = encoded_test_row(&cells, wide);
            row.cell_count =
                u32::try_from(cells.iter().flatten().count()).expect("cell count fits");
            row.cell_storage_buffer = Some(storage);
            row.cell_offsets = Some(offsets);
        }
        let tile_payload = tile.encode_to_vec();

        let formula_payload = tst::TableDataList {
            list_type: tst::table_data_list::ListType::Formula as i32,
            next_list_id: u32::try_from(formulas + 1).expect("next formula key fits"),
            entries,
            segments: Vec::new(),
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec();
        let owner_route = target
            .dependencies
            .selected_formula_owner
            .expect("wide source has selected formula owner")
            .message;
        let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(
            super::message_payload(&source, owner_route, Path::Package)
                .expect("formula owner payload"),
        )
        .expect("formula owner decodes");
        owner.cell_dependencies = Some(tsce::CellDependenciesExpandedArchive {
            cell_record: records,
        });
        owner.spanning_column_dependencies = None;
        owner.spanning_row_dependencies = None;
        owner.tiled_cell_dependencies = Some(tsce::CellDependenciesTiledArchive::default());
        let owner_payload = owner.encode_to_vec();
        let engine_route = target.dependencies.engine.expect("wide source has engine");
        let engine_payload = patch_nested_varint_field(
            super::message_payload(&source, engine_route, Path::Package)
                .expect("formula engine payload"),
            &[2, 5],
            true,
            Some(u64::try_from(formulas).expect("formula count fits")),
        )
        .expect("raw-preserving formula count patch");
        litchi_iwa_protos::numbers_table_cell_dependency_codec::decode_calculation_engine_with_report(
            &engine_payload,
            litchi_iwa_protos::numbers_table_cell_dependency_codec::DecodeOptions::new(
                engine_payload.len(),
                1_000_000,
                engine_payload.len() * 64,
                64,
                1_000_000,
                1_000_000,
            ),
        )
        .expect("patched formula engine passes strict projection");

        let mut replacements = [
            (tile_route.message, tile_payload),
            (target.storage.lists.formula.message, formula_payload),
            (owner_route, owner_payload),
            (engine_route, engine_payload),
        ];
        replacements.sort_unstable_by_key(|replacement| {
            (
                replacement.0.component_index,
                replacement.0.object_index,
                replacement.0.message_index,
            )
        });
        let replacements = replacements
            .iter()
            .map(|(route, payload)| rewrite::MessageReplacement {
                component_index: route.component_index,
                object_index: route.object_index,
                message_index: route.message_index,
                expected_type: route.message_type,
                payload,
                references: None,
            })
            .collect::<Vec<_>>();
        let previews = rewrite::root_preview_deletions(&source).expect("preview set resolves");
        rewrite::rewrite_with_precharge(
            &source,
            rewrite::RewritePlan {
                replacements: &replacements,
                preview_deletions: &previews,
            },
            |_reservation| Ok(()),
        )
        .expect("formula fanout source publishes")
        .package
    }

    fn synthetic_unique_rich_source() -> (Package, u64, u64) {
        let source = Package::open(fixture()).expect("native basic fixture opens");
        let position = CellPosition::from_a1("C2").expect("C2 is canonical");
        let (target, _report) = super::resolve::resolve_changed_target(&source, 0, 0, &[position])
            .expect("basic body coordinate resolves");
        let rich_index = target
            .storage
            .rich
            .as_ref()
            .expect("basic rich-text index resolves");
        assert!(rich_index.entries.is_empty());
        assert!(rich_index.pairs.is_empty());
        let rich_list = target
            .storage
            .lists
            .rich_text
            .as_ref()
            .expect("basic has a rooted rich-text list");
        assert_eq!(rich_list.entries, 0, "basic rich-text list is empty");
        assert!(
            rich_list.segments.is_empty(),
            "basic rich-text list is direct"
        );

        let first_identifier = source
            .state
            .components
            .catalog()
            .iter()
            .flat_map(|component| component.archive().objects.iter())
            .filter_map(|object| object.archive_info.identifier)
            .max()
            .expect("basic has native object identities")
            .checked_add(1)
            .expect("synthetic identifiers fit");
        let payload_identifier = first_identifier;
        let storage_identifier = first_identifier + 1;
        let style_identifier = first_identifier + 2;

        let list_route = rich_list.message;
        let mut list = tst::TableDataList::decode(
            super::message_payload(&source, list_route, Path::Package)
                .expect("source rich-list payload"),
        )
        .expect("source rich-list decodes");
        assert_eq!(
            list.list_type,
            tst::table_data_list::ListType::RichTextPayload as i32
        );
        assert!(list.entries.is_empty());
        assert!(list.segments.is_empty());
        list.next_list_id = 2;
        list.entries.push(tst::table_data_list::ListEntry {
            key: 1,
            refcount: 1,
            rich_text_payload: Some(tsp::Reference {
                identifier: payload_identifier,
                ..Default::default()
            }),
            ..Default::default()
        });
        let list_payload = list.encode_to_vec();
        let list_info = source
            .state
            .components
            .catalog()
            .get_index(list_route.component_index)
            .and_then(|component| component.archive().objects.get(list_route.object_index))
            .and_then(|object| {
                object
                    .archive_info
                    .message_infos
                    .get(list_route.message_index)
            })
            .expect("source rich-list metadata");
        let mut list_references = list_info.object_references.clone();
        assert!(!list_references.contains(&payload_identifier));
        list_references.push(payload_identifier);
        let list_delta = rewrite::ReferenceDelta {
            aggregate_before: list_info.object_references.clone(),
            aggregate_after: list_references,
            fields: Vec::new(),
        };

        let tile_route = target.storage.tiles.first().expect("basic has tile zero");
        let mut tile = tst::Tile::decode(
            super::message_payload(&source, tile_route.message, Path::Package)
                .expect("source tile payload"),
        )
        .expect("source tile decodes");
        let row = tile
            .row_infos
            .iter_mut()
            .find(|row| row.tile_row_index == position.row())
            .expect("source body row exists");
        let wide = row.has_wide_offsets.unwrap_or(false);
        let mut cells = decoded_test_row(row, 7);
        let mut rich_cell = BncCell::minimal();
        rich_cell.set_rich_text(1);
        cells[usize::try_from(position.column()).expect("column fits")] = Some(rich_cell.encode());
        let (storage, offsets) = encoded_test_row(&cells, wide);
        row.cell_count = u32::try_from(cells.iter().flatten().count()).expect("cell count fits");
        row.cell_storage_buffer = Some(storage);
        row.cell_offsets = Some(offsets);
        let tile_payload = tile.encode_to_vec();

        let rich_payload = tst::RichTextPayloadArchive {
            storage: tsp::Reference {
                identifier: storage_identifier,
                ..Default::default()
            },
            range: None,
            cellid: tst::CellId {
                packed_data: (position.row() << 16) | position.column(),
                expanded_coord: None,
            },
        }
        .encode_to_vec();
        let rich_storage = tswp::StorageArchive {
            kind: Some(tswp::storage_archive::KindType::Cell as i32),
            text: vec!["Original rich text".to_owned()],
            table_char_style: Some(tswp::ObjectAttributeTable {
                entries: vec![tswp::object_attribute_table::ObjectAttribute {
                    character_index: 4,
                    object: Some(tsp::Reference {
                        identifier: style_identifier,
                        ..Default::default()
                    }),
                }],
            }),
            ..Default::default()
        }
        .encode_to_vec();

        let mut payload_object = ArchiveObject::new(
            payload_identifier,
            vec![RawMessage {
                type_: 6_218,
                data: rich_payload,
            }],
        )
        .expect("synthetic rich payload object is valid");
        payload_object.archive_info.message_infos[0]
            .object_references
            .push(storage_identifier);
        let mut storage_object = ArchiveObject::new(
            storage_identifier,
            vec![RawMessage {
                type_: 2_001,
                data: rich_storage,
            }],
        )
        .expect("synthetic rich storage object is valid");
        storage_object.archive_info.message_infos[0]
            .object_references
            .push(style_identifier);
        let style_object = ArchiveObject::new(
            style_identifier,
            vec![RawMessage {
                type_: 2_002,
                data: Vec::new(),
            }],
        )
        .expect("synthetic style object is valid");

        let mut edits = vec![rewrite::ComponentEdit {
            component_index: list_route.component_index,
            messages: vec![rewrite::MessageEdit {
                object_index: list_route.object_index,
                message_index: list_route.message_index,
                expected_type: list_route.message_type,
                payload: list_payload,
                references: Some(list_delta),
            }],
            object_deletions: Vec::new(),
            new_objects: vec![payload_object, storage_object, style_object],
        }];
        if tile_route.message.component_index == list_route.component_index {
            edits[0].messages.push(rewrite::MessageEdit {
                object_index: tile_route.message.object_index,
                message_index: tile_route.message.message_index,
                expected_type: tile_route.message.message_type,
                payload: tile_payload,
                references: None,
            });
            edits[0]
                .messages
                .sort_unstable_by_key(|message| (message.object_index, message.message_index));
        } else {
            edits.push(rewrite::ComponentEdit {
                component_index: tile_route.message.component_index,
                messages: vec![rewrite::MessageEdit {
                    object_index: tile_route.message.object_index,
                    message_index: tile_route.message.message_index,
                    expected_type: tile_route.message.message_type,
                    payload: tile_payload,
                    references: None,
                }],
                object_deletions: Vec::new(),
                new_objects: Vec::new(),
            });
            edits.sort_unstable_by_key(|edit| edit.component_index);
        }
        let previews = rewrite::root_preview_deletions(&source).expect("preview set resolves");
        let outcome = rewrite::rewrite_staged_with_evidence_authorization(
            &source,
            rewrite::StagedRewritePlan {
                component_edits: edits,
                preview_deletions: &previews,
            },
            rewrite::EvidenceRetention::Omit,
            |_reservation| Ok(()),
            |_reservation, _cost| Ok(()),
        )
        .expect("synthetic rich source publishes");
        (outcome.package, storage_identifier, style_identifier)
    }

    fn rich_storage_text_and_references(package: &Package, identifier: u64) -> (String, Vec<u64>) {
        let resolved = package
            .state
            .index
            .resolve_ref_id(&package.state.components, identifier)
            .expect("synthetic rich identity resolves")
            .expect("synthetic rich identity exists");
        let object = package
            .state
            .components
            .catalog()
            .get_index(resolved.component_index)
            .and_then(|component| component.archive().objects.get(resolved.object_index))
            .expect("synthetic rich object exists");
        let message = object
            .messages
            .first()
            .expect("rich storage has one message");
        assert_eq!(message.type_, 2_001);
        let storage =
            tswp::StorageArchive::decode(message.data.as_slice()).expect("rich storage decodes");
        (
            storage.text.concat(),
            object.archive_info.message_infos[0]
                .object_references
                .clone(),
        )
    }

    fn formula_cache_number(package: &Package, position: CellPosition) -> f64 {
        crate::package::table_headers::resolve::resolve_target(package, 0, 0)
            .expect("formula seed native target resolves");
        let (target, _report) =
            super::resolve::resolve_changed_target(package, 0, 0, &[CellPosition::new(1, 1)])
                .expect("formula seed resolves");
        let tile_id = position.row() / target.storage.tile_size;
        let route = target
            .storage
            .tiles
            .binary_search_by_key(&tile_id, |route| route.tile_id)
            .ok()
            .and_then(|index| target.storage.tiles.get(index))
            .expect("formula tile route");
        let payload = super::message_payload(package, route.message, Path::Package)
            .expect("formula tile payload");
        let tile = tst::Tile::decode(payload).expect("formula tile decodes");
        let local_row = position.row() % target.storage.tile_size;
        let row = tile
            .row_infos
            .iter()
            .find(|row| row.tile_row_index == local_row)
            .expect("formula row exists");
        let offsets = row.cell_offsets.as_ref().expect("BNC offsets exist");
        let storage = row
            .cell_storage_buffer
            .as_ref()
            .expect("BNC storage exists");
        let column = usize::try_from(position.column()).expect("column fits");
        let offset_index = column.checked_mul(2).expect("offset index fits");
        let start = u16::from_le_bytes([
            *offsets.get(offset_index).expect("offset low byte"),
            *offsets.get(offset_index + 1).expect("offset high byte"),
        ]);
        assert_ne!(start, u16::MAX);
        let start = usize::from(start);
        let end = offsets[(offset_index + 2)..]
            .chunks_exact(2)
            .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
            .find(|offset| *offset != u16::MAX)
            .map_or(storage.len(), usize::from);
        let cell = BncCell::parse(&storage[start..end]).expect("formula BNC parses");
        match cell.cached_scalar().expect("cached scalar decodes") {
            Some(CachedScalar::Number(value)) => value.get(),
            other => panic!("expected numeric formula cache, got {other:?}"),
        }
    }

    fn observed_wide_commit(
        source: &Package,
        changes: usize,
        input: impl Fn(usize) -> Input,
    ) -> super::budget::Usage {
        let (commit, usage) = super::budget::testing::observe(None, || {
            let mut edit = source
                .edit_table_cells(0usize, 0usize)
                .expect("wide table edit starts");
            for index in 0..changes {
                edit = edit
                    .set(
                        CellPosition::new(
                            u32::try_from(index / 4_096 + 1).expect("wide row fits"),
                            u32::try_from(index % 4_096 + 7).expect("wide coordinate fits"),
                        ),
                        input(index),
                    )
                    .expect("wide change stages");
            }
            edit.commit()
        });
        let commit = commit.unwrap_or_else(|error| {
            panic!("rooted wide transaction commits: {error:?}; usage={usage:?}")
        });
        assert_eq!(commit.diagnostics().changed_cells(), changes);
        usage.expect("production budget usage is observed")
    }

    fn assert_linear_growth(small: u64, large: u64, counter: &str) {
        assert!(small != 0, "{counter} must be N-sensitive and nonzero");
        assert!(
            large > small,
            "{counter} must grow from 4K to 8K: {small} -> {large}"
        );
        assert!(
            large
                .checked_mul(100)
                .expect("scaled large counter remains representable")
                <= small
                    .checked_mul(220)
                    .expect("scaled small counter remains representable"),
            "{counter} grew too quickly: {small} -> {large}"
        );
    }

    #[test]
    fn rooted_numeric_4096_to_8192_observed_usage_is_bounded() {
        let source = synthetic_wide_source(8_199);
        let small = observed_wide_commit(&source, 4_096, |index| {
            Input::number(index as f64 + 10_000.0).expect("finite numeric input")
        });
        let large = observed_wide_commit(&source, 8_192, |index| {
            Input::number(index as f64 + 10_000.0).expect("finite numeric input")
        });
        for (name, left, right) in [
            ("updates", small.updates, large.updates),
            (
                "retained_elements",
                small.retained_elements,
                large.retained_elements,
            ),
            ("retained_bytes", small.retained_bytes, large.retained_bytes),
            ("wire_work", small.wire_work, large.wire_work),
            ("output_bytes", small.output_bytes, large.output_bytes),
            (
                "transaction_work",
                small.transaction_work,
                large.transaction_work,
            ),
        ] {
            assert_linear_growth(left, right, name);
        }
        assert_eq!((small.tile_reads, large.tile_reads), (1, 1));
        assert_eq!((small.tile_writes, large.tile_writes), (1, 1));
        assert_eq!(
            (
                small.output_artifact_allocations,
                large.output_artifact_allocations,
            ),
            (2, 2)
        );
        assert_eq!((small.candidate_reopens, large.candidate_reopens), (1, 1));
    }

    #[test]
    fn rooted_unique_text_4096_to_8192_observed_usage_is_bounded() {
        let source = synthetic_wide_source(8_199);
        let small = observed_wide_commit(&source, 4_096, |index| {
            Input::text(format!("rooted-unique-{index:05}")).expect("unique text allocates")
        });
        let large = observed_wide_commit(&source, 8_192, |index| {
            Input::text(format!("rooted-unique-{index:05}")).expect("unique text allocates")
        });
        for (name, left, right) in [
            ("updates", small.updates, large.updates),
            (
                "input_value_bytes",
                small.input_value_bytes,
                large.input_value_bytes,
            ),
            ("string_work", small.string_work, large.string_work),
            ("retained_bytes", small.retained_bytes, large.retained_bytes),
            ("output_bytes", small.output_bytes, large.output_bytes),
            (
                "transaction_work",
                small.transaction_work,
                large.transaction_work,
            ),
        ] {
            assert_linear_growth(left, right, name);
        }
        assert_eq!((small.list_writes, large.list_writes), (1, 1));
        assert_eq!(
            (
                small.output_artifact_allocations,
                large.output_artifact_allocations,
            ),
            (2, 2)
        );
        assert_eq!((small.candidate_reopens, large.candidate_reopens), (1, 1));
    }

    #[test]
    fn rooted_same_tile_4096_to_8192_stays_one_physical_tile() {
        let source = synthetic_wide_source(8_199);
        let small = observed_wide_commit(&source, 4_096, |index| Input::Boolean(index % 2 == 0));
        let large = observed_wide_commit(&source, 8_192, |index| Input::Boolean(index % 2 == 0));
        assert_linear_growth(small.updates, large.updates, "same-tile updates");
        assert_linear_growth(
            small.transaction_work,
            large.transaction_work,
            "same-tile transaction work",
        );
        assert_eq!((small.tile_reads, large.tile_reads), (1, 1));
        assert_eq!((small.tile_writes, large.tile_writes), (1, 1));
        assert_eq!((small.component_encodes, large.component_encodes), (2, 2));
        assert_eq!(
            (small.components_reassembled, large.components_reassembled),
            (2, 2)
        );
    }

    #[test]
    fn rooted_max_minus_one_refuses_before_evidence_component_output_reopen_and_snapshot() {
        let source = synthetic_wide_source(4_103);
        let mut limits = super::budget::testing::package_limits(&source);
        limits.max_updates = 4_095;
        let (result, usage) = super::budget::testing::observe(Some(limits), || {
            let mut edit = source
                .edit_table_cells(0usize, 0usize)
                .expect("rooted refusal edit starts");
            for index in 0..4_096 {
                edit = edit
                    .set(
                        CellPosition::new(
                            1,
                            u32::try_from(index + 7).expect("refusal coordinate fits"),
                        ),
                        Input::number(index as f64 + 20_000.0).expect("finite refusal input"),
                    )
                    .expect("refusal change stages");
            }
            edit.commit()
        });
        assert!(matches!(
            result,
            Err(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::Updates,
                observed: 4_096,
                maximum: 4_095,
                ..
            })
        ));
        let usage = usage.expect("refused production budget is observed");
        assert_eq!(usage.component_encodes, 0);
        assert_eq!(usage.components_reassembled, 0);
        assert_eq!(usage.output_artifact_allocations, 0);
        assert_eq!(usage.output_bytes, 0);
        assert_eq!(usage.candidate_reopens, 0);
        assert_eq!(usage.reopen_work, 0);
        assert_eq!(usage.locality_bytes, 0);
        assert_eq!(usage.retained_elements, 0);
        assert_eq!(usage.retained_bytes, 0);
        assert_eq!(usage.allocation_events, 0);
    }

    fn assert_no_publication_work(usage: super::budget::Usage) {
        assert_eq!(usage.component_encodes, 0);
        assert_eq!(usage.components_reassembled, 0);
        assert_no_output_reopen_or_locality_work(usage);
    }

    fn assert_no_output_reopen_or_locality_work(usage: super::budget::Usage) {
        assert_eq!(usage.output_artifact_allocations, 0);
        assert_eq!(usage.output_bytes, 0);
        assert_eq!(usage.candidate_reopens, 0);
        assert_eq!(usage.reopen_work, 0);
        assert_eq!(usage.locality_bytes, 0);
    }

    #[test]
    fn rooted_formula_max_minus_one_refuses_before_publication_and_reopen() {
        let source = synthetic_formula_fanout_source(4_096);
        let mut limits = super::budget::testing::package_limits(&source);
        limits.max_formula_work = 8_191;
        let (result, usage) = super::budget::testing::observe(Some(limits), || {
            source
                .edit_table_cells(0usize, 0usize)?
                .set(
                    CellPosition::new(1, 7),
                    Input::number(2.0).expect("finite formula source update"),
                )?
                .commit()
        });
        assert!(matches!(
            result,
            Err(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::FormulaWork,
                maximum: 8_191,
                ..
            })
        ));
        let usage = usage.expect("refused formula budget is observed");
        assert!(
            usage.formula_nodes != 0,
            "the rooted cache phase was entered"
        );
        assert_no_publication_work(usage);
    }

    #[test]
    fn rooted_sparse_evidence_work_max_minus_one_refuses_before_allocation_and_publication() {
        let source = synthetic_513_row_source();
        let before = source.source_bytes().to_vec();
        let run = || {
            super::budget::testing::observe(None, || {
                source
                    .edit_table_cells(0usize, 0usize)?
                    .set(
                        CellPosition::new(512, 6),
                        Input::number(513.0).expect("finite sparse evidence input"),
                    )?
                    .commit()
            })
        };
        let ((baseline, _baseline_usage), baseline_visits, requirement) =
            super::sparse_commit::testing::with_evidence_work_limit(None, run);
        baseline.expect("the rooted sparse evidence baseline commits");
        let requirement = requirement.expect("successful scan reports exact work");
        assert!(requirement > 1, "rooted shape scan has multiple work units");
        assert!(
            baseline_visits > 1,
            "rooted shape scan visits multiple items"
        );

        let maximum = requirement - 1;
        let ((result, usage), refusal_visits, refused_requirement) =
            super::sparse_commit::testing::with_evidence_work_limit(Some(maximum), run);
        assert_eq!(
            refused_requirement, None,
            "the refused scan never reports successful completion"
        );
        assert_eq!(
            refusal_visits + 1,
            baseline_visits,
            "the final required scan item is refused before its callback"
        );
        assert!(matches!(
            result,
            Err(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed,
                maximum: reported_maximum,
                ..
            }) if observed == requirement && reported_maximum == maximum
        ));
        assert_eq!(source.source_bytes(), before);
        assert_no_publication_work(usage.expect("sparse evidence refusal is observed"));
    }

    #[test]
    fn rooted_formula_fanout_4096_to_8192_observed_usage_is_bounded() {
        fn measured(formulas: usize) -> super::budget::Usage {
            let source = synthetic_formula_fanout_source(formulas);
            let (target, _report) =
                super::resolve::resolve_changed_target(&source, 0, 0, &[CellPosition::new(1, 7)])
                    .expect("rooted formula target resolves");
            assert_eq!(target.path(), Path::Table { sheet: 0, table: 0 });
            let (commit, usage) = super::budget::testing::observe(None, || {
                source
                    .edit_table_cells(0usize, 0usize)?
                    .set(
                        CellPosition::new(1, 7),
                        Input::number(2.0).expect("finite formula source update"),
                    )?
                    .commit()
            });
            let commit = commit.unwrap_or_else(|error| {
                panic!("rooted formula transaction commits: {error:?}; usage={usage:?}")
            });
            assert_eq!(commit.diagnostics().refreshed_formula_caches(), formulas);
            usage.expect("formula transaction usage is observed")
        }

        let small = measured(4_096);
        let large = measured(8_192);
        for (name, left, right) in [
            ("formula_nodes", small.formula_nodes, large.formula_nodes),
            ("formula_edges", small.formula_edges, large.formula_edges),
            (
                "range_candidates",
                small.range_candidates,
                large.range_candidates,
            ),
            ("cache_hosts", small.cache_hosts, large.cache_hosts),
            ("formula_work", small.formula_work, large.formula_work),
            (
                "transaction_work",
                small.transaction_work,
                large.transaction_work,
            ),
        ] {
            assert_linear_growth(left, right, name);
        }
        assert_eq!(
            (small.formula_graph_builds, large.formula_graph_builds,),
            (2, 2)
        );
        assert_eq!(
            (
                small.output_artifact_allocations,
                large.output_artifact_allocations,
            ),
            (2, 2)
        );
        assert_eq!((small.candidate_reopens, large.candidate_reopens), (1, 1));
    }

    fn row_header_cell_count(package: &Package, row: u32) -> Option<u32> {
        let (target, _report) =
            super::resolve::resolve_changed_target(package, 0, 0, &[CellPosition::new(row, 1)])
                .expect("row-header target resolves");
        let bucket =
            usize::try_from(row / super::sparse::HEADER_BUCKET_ROWS).expect("header bucket fits");
        let route = *target
            .storage
            .row_headers
            .get(bucket)
            .expect("row-header bucket exists");
        let payload = super::message_payload(package, route, Path::Package)
            .expect("row-header bucket payload exists");
        let bucket = tst::HeaderStorageBucket::decode(payload).expect("row-header bucket decodes");
        bucket
            .headers
            .iter()
            .find(|header| header.index == row)
            .map(|header| header.number_of_cells)
    }

    #[test]
    fn native_body_number_commit_apply_and_inverse_are_exact() -> TestResult {
        let source = Package::open(fixture())?;
        let before = source.source_bytes().to_vec();
        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set_a1("B3", Input::number(43.0)?)?
            .commit()?;
        let state = commit
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("B3")?)?;
        assert!(matches!(
            state.storage().value(),
            Some(Value::Number(value)) if value.get() == 43.0
        ));
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().changed_cells(), 1);
        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(restored.package().source_bytes(), before);
        Ok(())
    }

    #[test]
    fn native_existing_row_header_count_tracks_insert_and_delete_exactly() -> TestResult {
        let source = Package::open(fixture())?;
        let source_bytes = source.source_bytes().to_vec();
        let position = CellPosition::from_a1("C3")?;
        assert!(
            source
                .table_cell(0usize, 0usize, position)?
                .storage()
                .value()
                .is_none()
        );
        let before_count = row_header_cell_count(&source, position.row())
            .expect("existing row has a native header");

        let inserted = source
            .edit_table_cells(0usize, 0usize)?
            .set(position, Input::boolean(true))?
            .commit()?;
        assert_eq!(
            row_header_cell_count(inserted.package(), position.row()),
            Some(before_count + 1)
        );
        assert_eq!(
            source
                .apply_table_cells(inserted.patch())?
                .package()
                .source_bytes(),
            inserted.package().source_bytes()
        );
        assert_eq!(
            inserted
                .package()
                .apply_table_cells(&inserted.patch().inverse())?
                .package()
                .source_bytes(),
            source_bytes
        );

        let inserted_bytes = inserted.package().source_bytes().to_vec();
        let deleted = inserted
            .package()
            .edit_table_cells(0usize, 0usize)?
            .clear(position)?
            .commit()?;
        assert_eq!(
            row_header_cell_count(deleted.package(), position.row()),
            Some(before_count)
        );
        assert_eq!(
            inserted
                .package()
                .apply_table_cells(deleted.patch())?
                .package()
                .source_bytes(),
            deleted.package().source_bytes()
        );
        assert_eq!(
            deleted
                .package()
                .apply_table_cells(&deleted.patch().inverse())?
                .package()
                .source_bytes(),
            inserted_bytes
        );
        Ok(())
    }

    #[test]
    fn native_body_text_commit_apply_and_inverse_are_exact() -> TestResult {
        let source = Package::open(fixture())?;
        let before = source.source_bytes().to_vec();
        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set_a1("B2", Input::text("replacement")?)?
            .commit()?;
        let state = commit
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("B2")?)?;
        assert!(matches!(
            state.storage().value(),
            Some(Value::Text(value)) if value == "replacement"
        ));
        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(restored.package().source_bytes(), before);
        Ok(())
    }

    #[test]
    #[ignore = "requires external Numbers 14.4 formula-rich oracle"]
    fn formula_rich_oracle_unique_storage_commit_apply_inverse_are_exact() -> TestResult {
        let path = PathBuf::from(
            "/private/tmp/litchi-numbers-cell-batch-native.wuaiMp/oracle-preserved.numbers",
        );
        // The frozen native oracle is supplied by the external Numbers 14.4
        // evidence gate and intentionally is not copied into the repository.
        let source = Package::open(path)?;
        let before = source.source_bytes().to_vec();
        let position = CellPosition::from_a1("C2")?;
        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set(position, Input::text("CELL: Café changed")?)?
            .commit()?;
        let state = commit.package().table_cell(0usize, 0usize, position)?;
        assert!(matches!(
            state.storage().value(),
            Some(Value::Text(value)) if value == "CELL: Café changed"
        ));
        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let inverse = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(inverse.package().source_bytes(), before);
        Ok(())
    }

    #[test]
    fn synthetic_unique_rich_commit_apply_inverse_and_reference_release_are_exact() -> TestResult {
        let (source, storage_identifier, style_identifier) = synthetic_unique_rich_source();
        let before = source.source_bytes().to_vec();
        let position = CellPosition::from_a1("C2")?;
        let state = source.table_cell(0usize, 0usize, position)?;
        assert!(matches!(
            state.storage().value(),
            Some(Value::Text(value)) if value == "Original rich text"
        ));
        assert_eq!(
            rich_storage_text_and_references(&source, storage_identifier),
            ("Original rich text".to_owned(), vec![style_identifier])
        );

        let equal = source
            .edit_table_cells(0usize, 0usize)?
            .set(position, Input::text("Original rich text")?)?
            .commit()?;
        assert_eq!(equal.package().source_bytes(), source.source_bytes());
        assert!(equal.patch().is_noop());
        assert!(!equal.diagnostics().changed());
        assert_eq!(equal.diagnostics().changed_cells(), 0);
        assert_eq!(
            rich_storage_text_and_references(equal.package(), storage_identifier),
            ("Original rich text".to_owned(), vec![style_identifier])
        );

        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set(position, Input::text("Changed rich text")?)?
            .commit()?;
        assert_eq!(commit.diagnostics().changed_cells(), 1);
        let state = commit.package().table_cell(0usize, 0usize, position)?;
        assert!(matches!(
            state.storage().value(),
            Some(Value::Text(value)) if value == "Changed rich text"
        ));
        assert_eq!(
            rich_storage_text_and_references(commit.package(), storage_identifier),
            ("Changed rich text".to_owned(), Vec::new())
        );

        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(restored.package().source_bytes(), before);
        assert_eq!(
            rich_storage_text_and_references(restored.package(), storage_identifier),
            ("Original rich text".to_owned(), vec![style_identifier])
        );
        Ok(())
    }

    #[test]
    fn synthetic_body_formula_chain_refreshes_once_and_replays_exactly() -> TestResult {
        let source = synthetic_formula_cache_source(false);
        let before = source.source_bytes().to_vec();
        let precedent = CellPosition::from_a1("B2")?;
        let first_formula = CellPosition::from_a1("C2")?;
        let downstream_formula = CellPosition::from_a1("D2")?;
        assert_eq!(formula_cache_number(&source, first_formula), 323.0);
        assert_eq!(formula_cache_number(&source, downstream_formula), 646.0);

        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set(precedent, Input::number(121.0)?)?
            .commit()?;
        assert_eq!(commit.diagnostics().refreshed_formula_caches(), 2);
        assert_eq!(formula_cache_number(commit.package(), first_formula), 324.0);
        assert_eq!(
            formula_cache_number(commit.package(), downstream_formula),
            648.0
        );

        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(restored.package().source_bytes(), before);
        Ok(())
    }

    #[test]
    fn synthetic_formula_chain_uses_the_complete_batch_final_state() -> TestResult {
        let source = synthetic_formula_cache_source(false);
        let before = source.source_bytes().to_vec();
        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set_a1("B2", Input::number(121.0)?)?
            .set_a1("B3", Input::number(204.0)?)?
            .commit()?;
        assert_eq!(
            formula_cache_number(commit.package(), CellPosition::from_a1("C2")?),
            325.0
        );
        assert_eq!(
            formula_cache_number(commit.package(), CellPosition::from_a1("D2")?),
            650.0
        );
        assert_eq!(commit.diagnostics().refreshed_formula_caches(), 2);
        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(restored.package().source_bytes(), before);
        Ok(())
    }

    #[test]
    fn synthetic_unsupported_formula_refuses_before_publication() -> TestResult {
        let source = synthetic_formula_cache_source(true);
        let before = source.source_bytes().to_vec();
        let error = source
            .edit_table_cells(0usize, 0usize)?
            .set_a1("B2", Input::number(121.0)?)?
            .commit()
            .expect_err("unsupported reachable formula refuses");
        assert!(matches!(
            error,
            Error::UnsupportedDependency {
                kind: DependencyKind::FormulaCache,
                ..
            }
        ));
        assert_eq!(source.source_bytes(), before);
        Ok(())
    }

    #[test]
    fn synthetic_513_row_sparse_commit_apply_and_inverse_are_exact() -> TestResult {
        let source = synthetic_513_row_source();
        let before = source.source_bytes().to_vec();
        let position = CellPosition::new(512, 6);
        let (resolved, _report) =
            super::resolve::resolve_changed_target(&source, 0, 0, &[position])?;
        assert!(resolved.dependencies.cell_record_tiles.is_empty());
        assert!(resolved.dependencies.range_precedent_tiles.is_empty());
        let commit = source
            .edit_table_cells(0usize, 0usize)?
            .set(position, Input::number(513.0)?)?
            .commit()?;
        let state = commit.package().table_cell(0usize, 0usize, position)?;
        assert!(matches!(
            state.storage().value(),
            Some(Value::Number(value)) if value.get() == 513.0
        ));
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().changed_cells(), 1);
        let replay = source.apply_table_cells(commit.patch())?;
        assert_eq!(
            replay.package().source_bytes(),
            commit.package().source_bytes()
        );
        let restored = commit
            .package()
            .apply_table_cells(&commit.patch().inverse())?;
        assert_eq!(restored.package().source_bytes(), before);
        Ok(())
    }
}
