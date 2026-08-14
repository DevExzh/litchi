//! Exact-source selector-first Numbers table-cell mutations.

mod budget;
mod cache;
mod cache_commit;
mod formula_author;
mod formula_list;
mod formula_metadata;
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

use litchi_iwa_protos::package_metadata_codec::{
    Batch, ObjectUuidAddition, PreparedPackageMetadataRewrite, UuidBits,
    prepare_package_metadata_rewrite,
};

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

#[cfg(test)]
mod aggregate_testing {
    use core::cell::Cell;

    use super::budget::TransactionLimits;

    std::thread_local! {
        static REQUIRED_LIMITS: Cell<Option<TransactionLimits>> = const { Cell::new(None) };
        static AUTHORIZATION_LIMITS: Cell<Option<TransactionLimits>> = const { Cell::new(None) };
        static EXECUTIONS: Cell<usize> = const { Cell::new(0) };
    }

    pub(super) fn reset() {
        REQUIRED_LIMITS.with(|slot| slot.set(None));
        AUTHORIZATION_LIMITS.with(|slot| slot.set(None));
        EXECUTIONS.with(|slot| slot.set(0));
    }

    pub(super) fn record_requirement(required: TransactionLimits) {
        REQUIRED_LIMITS.with(|slot| slot.set(Some(required)));
    }

    pub(super) fn record_execution() {
        EXECUTIONS.with(|slot| slot.set(slot.get().saturating_add(1)));
    }

    pub(super) fn set_authorization_limits(limits: TransactionLimits) {
        AUTHORIZATION_LIMITS.with(|slot| slot.set(Some(limits)));
    }

    pub(super) fn authorization_limits() -> Option<TransactionLimits> {
        AUTHORIZATION_LIMITS.with(Cell::get)
    }

    pub(super) fn observation() -> (Option<TransactionLimits>, usize) {
        (REQUIRED_LIMITS.with(Cell::get), EXECUTIONS.with(Cell::get))
    }
}

struct PreparedFormulaContext {
    authored: formula_author::PreparedFormulaBatch,
    referenced_paths: Vec<formula_author::ReferencedTablePath>,
    external_targets: Vec<(formula_author::ReferencedTablePath, resolve::Target)>,
}

#[derive(Debug, Clone, Copy)]
struct FormulaDependencyTileAssignment {
    object_id: u64,
    uuid: UuidBits,
    column_begin: u32,
    row_begin: u32,
}

#[derive(Debug, Clone, Copy)]
struct ExistingFormulaDependencyTile {
    object_id: u64,
    column_begin: u32,
    row_begin: u32,
}

#[derive(Debug)]
struct PreparedFormulaDependencyTiles {
    inspection: Option<metadata::Inspection>,
    component_index: usize,
    existing: Vec<ExistingFormulaDependencyTile>,
    existing_by_object: Vec<ExistingFormulaDependencyTile>,
    assignments: Vec<FormulaDependencyTileAssignment>,
    retained_bytes: usize,
    retained_elements: usize,
}

fn prepare_formula_dependency_tiles(
    source: &Package,
    target: &resolve::Target,
    authored: &formula_author::PreparedFormulaBatch,
    budget: &mut budget::TransactionBudget,
    path: crate::package::table_cells::Path,
) -> Result<PreparedFormulaDependencyTiles, Error> {
    let selected_owner = target
        .dependencies
        .selected_formula_owner
        .ok_or(Error::InvalidSource { path })?;
    let owner_scan_count = target.dependencies.formula_owners.len();
    budget.reserve(budget::Usage {
        lookups: usize_u64(owner_scan_count),
        transaction_work: usize_u64(owner_scan_count),
        ..budget::Usage::default()
    })?;
    let source_owner = target
        .dependencies
        .formula_owners
        .iter()
        .find(|owner| owner.internal_owner_id == selected_owner.internal_owner_id)
        .ok_or(Error::InvalidSource { path })?;
    let existing_count = source_owner.cell_record_tiles.len();
    let origin_capacity = authored.formulas.len();
    let existing_index_work = sort_work(existing_count, path)?
        .checked_mul(2)
        .and_then(|work| work.checked_add(existing_count.checked_mul(2)?))
        .ok_or(Error::InvalidSource { path })?;
    let index_work = existing_index_work
        .checked_add(sort_work(origin_capacity, path)?)
        .and_then(|work| work.checked_add(origin_capacity.checked_mul(3)?))
        .and_then(|work| {
            origin_capacity
                .checked_mul(binary_search_work(existing_count))
                .and_then(|search| work.checked_add(search))
        })
        .ok_or(Error::InvalidSource { path })?;
    budget.reserve(budget::Usage {
        lookups: usize_u64(origin_capacity),
        transaction_work: usize_u64(index_work),
        ..budget::Usage::default()
    })?;
    let existing_bytes = source_owner
        .cell_record_tiles
        .len()
        .checked_mul(size_of::<ExistingFormulaDependencyTile>())
        .ok_or(Error::InvalidSource { path })?;
    let existing_retained_bytes = existing_bytes
        .checked_mul(2)
        .ok_or(Error::InvalidSource { path })?;
    budget.reserve_retained(
        usize_u64(
            source_owner
                .cell_record_tiles
                .len()
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        ),
        usize_u64(existing_retained_bytes),
        u64::from(!source_owner.cell_record_tiles.is_empty()).saturating_mul(2),
    )?;
    let mut existing = Vec::new();
    existing
        .try_reserve_exact(source_owner.cell_record_tiles.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: source_owner.cell_record_tiles.len(),
        })?;
    require_exact_capacity(&existing, source_owner.cell_record_tiles.len(), path)?;
    for route in &source_owner.cell_record_tiles {
        let payload = message_payload(source, *route, path)?;
        let remaining = budget.remaining()?;
        let max_bytes = payload.len();
        let field_limit = payload
            .len()
            .checked_mul(8)
            .ok_or(Error::InvalidSource { path })?
            .min(
                usize::try_from(remaining.wire_fields)
                    .map_err(|_| Error::InvalidSource { path })?,
            );
        let work_limit = payload
            .len()
            .checked_mul(128)
            .ok_or(Error::InvalidSource { path })?
            .min(
                usize::try_from(remaining.wire_work.min(remaining.transaction_work))
                    .map_err(|_| Error::InvalidSource { path })?,
            );
        let options = litchi_iwa_protos::numbers_table_cell_dependency_codec::DecodeOptions::new(
            max_bytes,
            field_limit,
            work_limit,
            64,
            usize::try_from(remaining.references).map_err(|_| Error::InvalidSource { path })?,
            max_bytes,
        );
        budget.authorize(remaining)?;
        let decoded =
            litchi_iwa_protos::numbers_table_cell_dependency_codec::decode_cell_record_tile_with_report(
                payload, options,
            );
        let (snapshot, report) = match decoded {
            Ok(decoded) => decoded,
            Err(_) => {
                budget.cancel_authorization();
                return Err(Error::InvalidSource { path });
            },
        };
        let usage = budget::Usage {
            wire_bytes: usize_u64(report.source_bytes()),
            wire_fields: usize_u64(report.fields()),
            wire_work: usize_u64(report.work_bytes()),
            references: usize_u64(report.references()),
            transaction_work: usize_u64(report.work_bytes()),
            ..budget::Usage::default()
        };
        budget.record_authorized(usage)?;
        if snapshot.internal_owner_id() != selected_owner.internal_owner_id {
            return Err(Error::InvalidSource { path });
        }
        existing.push(ExistingFormulaDependencyTile {
            object_id: message_object_identifier(source, *route, path)?,
            column_begin: snapshot.tile_column_begin(),
            row_begin: snapshot.tile_row_begin(),
        });
    }
    existing.sort_unstable_by_key(|tile| (tile.row_begin, tile.column_begin));
    if existing.windows(2).any(|pair| {
        (pair[0].row_begin, pair[0].column_begin) >= (pair[1].row_begin, pair[1].column_begin)
    }) {
        return Err(Error::InvalidSource { path });
    }
    let mut existing_by_object = Vec::new();
    existing_by_object
        .try_reserve_exact(existing.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: existing.len(),
        })?;
    require_exact_capacity(&existing_by_object, existing.len(), path)?;
    existing_by_object.extend_from_slice(&existing);
    existing_by_object.sort_unstable_by_key(|tile| tile.object_id);
    if existing_by_object
        .windows(2)
        .any(|pair| pair[0].object_id >= pair[1].object_id)
    {
        return Err(Error::InvalidSource { path });
    }
    let mut origins = Vec::new();
    let origin_bytes = origin_capacity
        .checked_mul(size_of::<(u32, u32)>())
        .ok_or(Error::InvalidSource { path })?;
    budget.reserve_scratch(usize_u64(origin_bytes), u64::from(origin_capacity != 0))?;
    origins
        .try_reserve_exact(origin_capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: origin_bytes,
        })?;
    require_exact_capacity(&origins, origin_capacity, path)?;
    for formula in &authored.formulas {
        origins.push((
            formula.position.row() / 128 * 128,
            formula.position.column() / 32 * 32,
        ));
    }
    origins.sort_unstable();
    origins.dedup();
    origins.retain(|&(row, column)| {
        existing
            .binary_search_by_key(&(row, column), |tile| (tile.row_begin, tile.column_begin))
            .is_err()
    });
    if origins.is_empty() {
        budget.release_scratch(usize_u64(origin_bytes));
        let retained_elements = existing
            .len()
            .checked_mul(2)
            .ok_or(Error::InvalidSource { path })?;
        return Ok(PreparedFormulaDependencyTiles {
            inspection: None,
            component_index: target
                .dependencies
                .engine
                .ok_or(Error::InvalidSource { path })?
                .component_index,
            existing,
            existing_by_object,
            assignments: Vec::new(),
            retained_bytes: existing_retained_bytes,
            retained_elements,
        });
    }
    let remaining = budget.remaining()?;
    let additions = origins
        .len()
        .checked_mul(2)
        .ok_or(Error::InvalidSource { path })?;
    let structural_work = source
        .source_bytes()
        .len()
        .checked_add(source.object_count())
        .ok_or(Error::InvalidSource { path })?;
    let structural_usage = budget::Usage {
        lookups: usize_u64(source.object_count()),
        transaction_work: usize_u64(structural_work),
        ..budget::Usage::default()
    };
    preflight_usage(structural_usage, remaining, path)?;
    let mut plan_remaining = remaining;
    plan_remaining.transaction_work = plan_remaining
        .transaction_work
        .checked_sub(structural_usage.transaction_work)
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::TransactionWork,
            observed: structural_usage.transaction_work,
            maximum: remaining.transaction_work,
            path,
        })?;
    let options = sparse_commit::metadata_options(source, plan_remaining, additions, path)?;
    budget.authorize(remaining)?;
    let plan = match metadata::plan_inspection(source, options, path) {
        Ok(plan) => plan,
        Err(error) => {
            budget.cancel_authorization();
            budget.release_scratch(usize_u64(origin_bytes));
            return Err(error);
        },
    };
    let plan_usage = sparse_commit::package_metadata_plan_usage(plan.first_report())?;
    let mut exact_plan_usage = plan_usage;
    exact_plan_usage.lookups = exact_plan_usage
        .lookups
        .checked_add(structural_usage.lookups)
        .ok_or(Error::InvalidSource { path })?;
    exact_plan_usage.transaction_work = exact_plan_usage
        .transaction_work
        .checked_add(structural_usage.transaction_work)
        .ok_or(Error::InvalidSource { path })?;
    budget.record_authorized(exact_plan_usage)?;
    let inspection_requirements = plan.requirements(path)?;
    let new_retained_bytes = origins
        .len()
        .checked_mul(size_of::<FormulaDependencyTileAssignment>())
        .ok_or(Error::InvalidSource { path })?;
    let component_index = target
        .dependencies
        .engine
        .ok_or(Error::InvalidSource { path })?
        .component_index;
    let mut collect_envelope = sparse_commit::package_metadata_collect_usage(
        plan.first_report(),
        inspection_requirements,
        origins.len(),
        source.source_bytes().len(),
        path,
    )?;
    collect_envelope.peak_scratch_bytes = collect_envelope
        .peak_scratch_bytes
        .checked_add(usize_u64(origin_bytes))
        .ok_or(Error::InvalidSource { path })?;
    collect_envelope.retained_elements = collect_envelope
        .retained_elements
        .checked_add(usize_u64(origins.len()))
        .ok_or(Error::InvalidSource { path })?;
    collect_envelope.retained_bytes = collect_envelope
        .retained_bytes
        .checked_add(usize_u64(new_retained_bytes))
        .ok_or(Error::InvalidSource { path })?;
    collect_envelope.allocation_events = collect_envelope
        .allocation_events
        .checked_add(u64::from(!origins.is_empty()))
        .ok_or(Error::InvalidSource { path })?;
    budget.authorize(collect_envelope)?;
    let (inspection, assignments) =
        match plan.collect(source, options, path).and_then(|inspection| {
            let (identifiers, uuids) =
                inspection.allocate_identifiers(origins.len(), source.source_bytes(), path)?;
            let mut assignments = Vec::new();
            assignments
                .try_reserve_exact(origins.len())
                .map_err(|_| Error::Allocation {
                    kind: crate::package::table_cells::LimitKind::RetainedElements,
                    amount: origins.len(),
                })?;
            require_exact_capacity(&assignments, origins.len(), path)?;
            for (((row_begin, column_begin), object_id), uuid) in
                origins.iter().copied().zip(identifiers).zip(uuids)
            {
                assignments.push(FormulaDependencyTileAssignment {
                    object_id,
                    uuid,
                    column_begin,
                    row_begin,
                });
            }
            Ok((inspection, assignments))
        }) {
            Ok(prepared) => prepared,
            Err(error) => {
                budget.cancel_authorization();
                budget.release_scratch(usize_u64(origin_bytes));
                return Err(error);
            },
        };
    let mut collect_usage = sparse_commit::package_metadata_collect_usage(
        inspection.second_report,
        inspection_requirements,
        origins.len(),
        source.source_bytes().len(),
        path,
    )?;
    collect_usage.peak_scratch_bytes = collect_usage
        .peak_scratch_bytes
        .checked_add(usize_u64(origin_bytes))
        .ok_or(Error::InvalidSource { path })?;
    collect_usage.retained_elements = collect_usage
        .retained_elements
        .checked_add(usize_u64(assignments.len()))
        .ok_or(Error::InvalidSource { path })?;
    collect_usage.retained_bytes = collect_usage
        .retained_bytes
        .checked_add(usize_u64(new_retained_bytes))
        .ok_or(Error::InvalidSource { path })?;
    collect_usage.allocation_events = collect_usage
        .allocation_events
        .checked_add(u64::from(!assignments.is_empty()))
        .ok_or(Error::InvalidSource { path })?;
    budget.record_authorized(collect_usage)?;
    budget.release_scratch(usize_u64(origin_bytes));
    let retained_elements = existing
        .len()
        .checked_mul(2)
        .and_then(|elements| elements.checked_add(assignments.len()))
        .and_then(|elements| elements.checked_add(inspection_requirements.retained_elements))
        .ok_or(Error::InvalidSource { path })?;
    Ok(PreparedFormulaDependencyTiles {
        inspection: Some(inspection),
        component_index,
        existing,
        existing_by_object,
        assignments,
        retained_bytes: existing_retained_bytes
            .checked_add(new_retained_bytes)
            .and_then(|bytes| bytes.checked_add(inspection_requirements.retained_bytes))
            .ok_or(Error::InvalidSource { path })?,
        retained_elements,
    })
}

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
        let source = self.state.source.clone();
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
                // The locality proof compares the same retained archive bytes
                // through ZIP entry, component, object, message and reference
                // views. Reserve the shared conservative traversal envelope,
                // while recording only the verifier's exact byte report.
                .and_then(|bytes| bytes.checked_mul(16))
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

struct PreparedTilePublication<'source> {
    route: resolve::MessageRoute,
    plan: tile::PreparedTileRewrite<'source>,
    expected_transitions: usize,
}

struct PreparedHeaderPublication<'source> {
    route: resolve::MessageRoute,
    plan: sparse::PreparedExistingHeaderBucketRewrite<'source>,
}

struct PreparedReferenceDelta {
    before: Vec<u64>,
    after: Vec<u64>,
}

struct BoundFormulaMetadata {
    replacements: Vec<PreparedReplacement>,
    new_tiles: Vec<formula_metadata::MessageEdit>,
    artifact_disposable_elements: usize,
    artifact_disposable_bytes: usize,
    binding_disposable_elements: usize,
    binding_disposable_bytes: usize,
}

struct BoundFormulaList {
    replacements: Vec<PreparedReplacement>,
    artifact_disposable_elements: usize,
    artifact_disposable_bytes: usize,
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
        let source_bytes = source.state.source.clone();
        let patch = Patch::from_exact(path, requested, 0, source_bytes.clone(), source_bytes);
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
    let has_formula_input = changes
        .iter()
        .any(|change| matches!(change.input_ref(), Some(Input::Formula { .. })));
    // Formula dependencies can be present even when this batch authors only
    // scalar inputs. Resolve the formula-capable read target for every batch
    // so a later scalar precedent edit cannot lose existing cross-owner graph
    // authority before cache refresh planning begins.
    let resolved = resolve::resolve_formula_read_target_with_remaining(
        source,
        sheet_position,
        table_position,
        &positions,
        resolve_authorization,
    );
    let (target, resolve_report) = match resolved {
        Ok(result) => result,
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    let resolved_target_usage = match resolve_usage(&budget, resolve_report) {
        Ok(usage) => usage,
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    budget.record_authorized(resolved_target_usage)?;
    if target.path() != path {
        return Err(Error::InvalidSource { path });
    }
    // Exact semantic scalar no-ops were removed before this function. The
    // formula-capable resolver deliberately permits a locked read so an exact
    // formula/cache no-op can still be proved below; every remaining
    // non-formula mutation must retain the ordinary locked-table refusal.
    if target.native.locked == crate::table::lock::State::Locked && !has_formula_input {
        return Err(Error::TableLocked { path });
    }
    let needs_formula_context = has_formula_input || target.storage.lists.formula.entries != 0;
    let mut formula_context = if needs_formula_context {
        let selected_owner =
            target
                .dependencies
                .selected_formula_owner
                .ok_or(Error::UnsupportedDependency {
                    path,
                    kind: crate::package::table_cells::DependencyKind::Formula,
                })?;
        let local_path = formula_author::ReferencedTablePath {
            sheet: u32::try_from(sheet_position).map_err(|_| Error::InvalidSource { path })?,
            table: u32::try_from(table_position).map_err(|_| Error::InvalidSource { path })?,
        };
        let referenced = formula_author::referenced_table_paths(
            &changes,
            staging_usage.formula_nodes(),
            &mut budget,
            path,
        )?;
        // Retain a read-only target for every semantic table which belongs to
        // this calculation engine, not only tables explicitly named by the
        // authored expressions. Metadata/cache publication must preserve the
        // complete engine owner registry.
        let capacity_scan = budget::Usage {
            transaction_work: usize_u64(source.sheets().len()),
            ..budget::Usage::default()
        };
        budget.reserve(capacity_scan)?;
        let external_capacity = source
            .sheets()
            .iter()
            .try_fold(0usize, |count, sheet| {
                count.checked_add(sheet.tables().len())
            })
            .and_then(|count| count.checked_sub(1))
            .ok_or(Error::InvalidSource { path })?;
        let owner_index_capacity = target
            .dependencies
            .formula_owners
            .iter()
            .filter(|owner| owner.formula_owner_object_id.is_some())
            .count();
        let owner_index_bytes = owner_index_capacity
            .checked_mul(size_of::<(u64, u32, u64, u64)>())
            .ok_or(Error::InvalidSource { path })?;
        let registry_work = owner_index_capacity
            .checked_add(sort_work(owner_index_capacity, path)?)
            .and_then(|work| {
                external_capacity
                    .checked_mul(binary_search_work(owner_index_capacity).checked_add(1)?)
                    .and_then(|discovery| work.checked_add(discovery))
            })
            .and_then(|work| {
                referenced
                    .len()
                    .checked_mul(binary_search_work(external_capacity))
                    .and_then(|coverage| work.checked_add(coverage))
            })
            .ok_or(Error::InvalidSource { path })?;
        budget.reserve(budget::Usage {
            scratch_bytes: usize_u64(owner_index_bytes),
            allocation_events: u64::from(owner_index_capacity != 0),
            lookups: usize_u64(
                external_capacity
                    .checked_add(referenced.len())
                    .ok_or(Error::InvalidSource { path })?,
            ),
            transaction_work: usize_u64(registry_work),
            ..budget::Usage::default()
        })?;
        let mut owner_index = Vec::new();
        owner_index
            .try_reserve_exact(owner_index_capacity)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
                amount: owner_index_bytes,
            })?;
        require_exact_capacity(&owner_index, owner_index_capacity, path)?;
        for owner in &target.dependencies.formula_owners {
            if let Some(object_id) = owner.formula_owner_object_id {
                owner_index.push((
                    object_id,
                    owner.internal_owner_id,
                    owner.uid_lower,
                    owner.uid_upper,
                ));
            }
        }
        owner_index.sort_unstable_by_key(|owner| owner.0);
        if owner_index
            .windows(2)
            .any(|owners| owners[0].0 >= owners[1].0)
        {
            budget.release_scratch(usize_u64(owner_index_bytes));
            return Err(Error::InvalidSource { path });
        }
        let mut external_owners = Vec::new();
        let mut external_targets = Vec::new();
        reserve_retained_vec::<formula_author::ResolvedOwner>(
            &mut budget,
            external_capacity,
            path,
        )?;
        reserve_retained_vec::<(formula_author::ReferencedTablePath, resolve::Target)>(
            &mut budget,
            external_capacity,
            path,
        )?;
        external_owners
            .try_reserve_exact(external_capacity)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: external_capacity,
            })?;
        external_targets
            .try_reserve_exact(external_capacity)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: external_capacity,
            })?;
        require_exact_capacity(&external_owners, external_capacity, path)?;
        require_exact_capacity(&external_targets, external_capacity, path)?;
        for (sheet, semantic_sheet) in source.sheets().iter().enumerate() {
            for (table_index, _) in semantic_sheet.tables().enumerate() {
                let referenced = formula_author::ReferencedTablePath {
                    sheet: u32::try_from(sheet).map_err(|_| Error::InvalidSource { path })?,
                    table: u32::try_from(table_index).map_err(|_| Error::InvalidSource { path })?,
                };
                if referenced == local_path {
                    continue;
                }
                let authorization = authorize_remaining(&mut budget)?;
                let resolved = resolve::resolve_formula_read_target_with_remaining(
                    source,
                    usize::try_from(referenced.sheet).map_err(|_| Error::InvalidSource { path })?,
                    usize::try_from(referenced.table).map_err(|_| Error::InvalidSource { path })?,
                    &[],
                    authorization,
                );
                let (external, report) = match resolved {
                    Ok(result) => result,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(error);
                    },
                };
                let usage = match resolve_usage(&budget, report) {
                    Ok(usage) => usage,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(error);
                    },
                };
                budget.record_authorized(usage)?;
                let Some(external_owner) = external.dependencies.selected_formula_owner else {
                    continue;
                };
                let owner_key = (
                    external.native.drawable_identifier,
                    external_owner.internal_owner_id,
                    external_owner.uid_lower,
                    external_owner.uid_upper,
                );
                if external.dependencies.engine != target.dependencies.engine
                    || owner_index
                        .binary_search_by_key(&owner_key.0, |candidate| candidate.0)
                        .ok()
                        .and_then(|index| owner_index.get(index))
                        != Some(&owner_key)
                {
                    continue;
                }
                external_owners.push(formula_author::ResolvedOwner {
                    path: referenced,
                    uid_lower: external_owner.uid_lower,
                    uid_upper: external_owner.uid_upper,
                    internal_owner: external_owner.internal_owner_id,
                    rows: external.native.rows,
                    columns: external.native.columns,
                });
                external_targets.push((referenced, external));
            }
        }
        budget.release_scratch(usize_u64(owner_index_bytes));
        let sort_envelope = budget::Usage {
            transaction_work: usize_u64(sort_work(external_owners.len(), path)?),
            ..budget::Usage::default()
        };
        budget.authorize(sort_envelope)?;
        external_owners.sort_unstable_by_key(|owner| (owner.uid_lower, owner.uid_upper));
        budget.record_authorized(sort_envelope)?;
        if referenced.iter().any(|referenced| {
            *referenced != local_path
                && external_targets
                    .binary_search_by_key(referenced, |(path, _)| *path)
                    .is_err()
        }) {
            return Err(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::Formula,
            });
        }
        let authored = formula_author::prepare(
            source,
            &changes,
            staging_usage.formula_nodes(),
            formula_author::LocalOwner {
                path: local_path,
                internal_owner: selected_owner.internal_owner_id,
                owners: &external_owners,
            },
            target.native.rows,
            target.native.columns,
            &mut budget,
            path,
        )?;
        PreparedFormulaContext {
            authored,
            referenced_paths: referenced,
            external_targets,
        }
    } else {
        PreparedFormulaContext {
            authored: formula_author::PreparedFormulaBatch::default(),
            referenced_paths: Vec::new(),
            external_targets: Vec::new(),
        }
    };
    let authored = &formula_context.authored;
    let mut formula_dependency_new_tiles = Vec::new();
    let mut formula_dependency_tiles = None;
    let mut formula_bnc_changes = Vec::new();
    let mut formula_cache_rewrites = Vec::new();
    let mut formula_text_change_indices = Vec::new();
    let mut formula_refreshed_hosts = 0usize;
    let mut formula_source_hosts = Vec::new();
    let mut formula_list_deltas = Vec::new();
    let mut formula_list_segments = Vec::new();
    let mut prepared_formula_list = None;
    let mut prepared_formula_list_limits = None;
    let mut prepared_formula_metadata = None;
    let mut formula_dependency_uuid_additions = Vec::new();
    let mut prepared_package_metadata: Option<PreparedPackageMetadataRewrite<'_, '_>> = None;
    // Existing formulas alone require cache refresh authority, not a formula
    // list/dependency rewrite. Only authored formulas or an ambiguous clear
    // need the existing-host index below; ordinary scalar edits proceed
    // directly to the complete nonlogical cache planner.
    let mut formula_publication =
        !authored.formulas.is_empty() || changes.iter().any(|change| change.input_ref().is_none());
    if formula_publication {
        'formula: {
            let existing =
                cache_commit::prepare_existing_formula_index(source, &target, &mut budget, path)?;
            if authored.formulas.is_empty()
                && !changes.iter().any(|change| {
                    let coordinate = (change.position().row(), change.position().column());
                    existing
                        .formula_list
                        .hosts
                        .binary_search_by_key(&coordinate, |host| (host.row, host.column))
                        .is_ok()
                })
            {
                formula_publication = false;
                break 'formula;
            }
            ensure_formula_owner_context_coverage(&target, &formula_context, &mut budget, path)?;
            let canonical_match_usage = budget::Usage {
                lookups: usize_u64(
                    authored
                        .formulas
                        .len()
                        .checked_mul(4)
                        .ok_or(Error::InvalidSource { path })?,
                ),
                transaction_work: usize_u64(
                    authored
                        .canonical_match_work(&existing, path)?
                        .checked_add(authored.supplied_cache_match_work(&existing, path)?)
                        .ok_or(Error::InvalidSource { path })?,
                ),
                ..budget::Usage::default()
            };
            budget.authorize(canonical_match_usage)?;
            let canonical_bytes_unchanged =
                match authored.canonical_bytes_match_existing(&existing, path) {
                    Ok(unchanged) => unchanged,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(error);
                    },
                };
            let supplied_caches_unchanged =
                match authored.supplied_caches_match_existing(&changes, &existing, path) {
                    Ok(unchanged) => unchanged,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(error);
                    },
                };
            budget.record_authorized(canonical_match_usage)?;
            let _existing_cache_count = existing.caches.len();
            for formula in &authored.formulas {
                if changes
                    .get(formula.change_index)
                    .map(crate::table::cells::Change::position)
                    != Some(formula.position)
                {
                    return Err(Error::Verification { path });
                }
            }
            let formula_only = authored.formulas.len() == changes.len();
            if canonical_bytes_unchanged && supplied_caches_unchanged && formula_only {
                let source_bytes = source.state.source.clone();
                let patch =
                    Patch::from_exact(path, requested, 0, source_bytes.clone(), source_bytes);
                return Ok(Commit::new(
                    source.snapshot(),
                    patch,
                    Diagnostics::unchanged(requested),
                ));
            }
            if target.native.locked == crate::table::lock::State::Locked {
                return Err(Error::TableLocked { path });
            }
            let orchestration_work = existing
                .formula_list
                .hosts
                .len()
                .checked_mul(2)
                .and_then(|work| {
                    changes
                        .len()
                        .checked_mul(
                            binary_search_work(existing.formula_list.hosts.len())
                                .checked_mul(2)?
                                .checked_add(
                                    binary_search_work(authored.formulas.len()).checked_mul(2)?,
                                )?
                                .checked_add(binary_search_work(changes.len()))?
                                .checked_add(10)?,
                        )
                        .and_then(|formula_work| work.checked_add(formula_work))
                })
                .and_then(|work| work.checked_add(existing.formula_list.segments.len()))
                .and_then(|work| work.checked_add(1))
                .ok_or(Error::InvalidSource { path })?;
            let orchestration_usage = budget::Usage {
                lookups: usize_u64(
                    changes
                        .len()
                        .checked_mul(5)
                        .ok_or(Error::InvalidSource { path })?,
                ),
                transaction_work: usize_u64(orchestration_work),
                ..budget::Usage::default()
            };
            budget.authorize(orchestration_usage)?;
            budget.record_authorized(orchestration_usage)?;
            reserve_retained_vec::<formula_list::SourceHost>(
                &mut budget,
                existing.formula_list.hosts.len(),
                path,
            )?;
            formula_source_hosts
                .try_reserve_exact(existing.formula_list.hosts.len())
                .map_err(|_| Error::Allocation {
                    kind: crate::package::table_cells::LimitKind::RetainedElements,
                    amount: existing.formula_list.hosts.len(),
                })?;
            require_exact_capacity(
                &formula_source_hosts,
                existing.formula_list.hosts.len(),
                path,
            )?;
            for host in &existing.formula_list.hosts {
                formula_source_hosts.push(formula_list::SourceHost {
                    row: host.row,
                    column: host.column,
                    key: host.key,
                });
            }
            if formula_source_hosts
                .windows(2)
                .any(|pair| pair[0] >= pair[1])
            {
                return Err(Error::InvalidSource { path });
            }
            let delta_capacity = changes.iter().try_fold(0usize, |count, change| {
                let coordinate = (change.position().row(), change.position().column());
                let old = formula_source_hosts
                    .binary_search_by_key(&coordinate, |host| (host.row, host.column))
                    .is_ok();
                let new = authored
                    .formulas
                    .binary_search_by_key(&coordinate, |formula| {
                        (formula.position.row(), formula.position.column())
                    })
                    .is_ok();
                count
                    .checked_add(usize::from(old || new))
                    .ok_or(Error::InvalidSource { path })
            })?;
            reserve_retained_vec::<formula_list::HostDelta<'_>>(&mut budget, delta_capacity, path)?;
            formula_list_deltas
                .try_reserve_exact(delta_capacity)
                .map_err(|_| Error::Allocation {
                    kind: crate::package::table_cells::LimitKind::RetainedElements,
                    amount: delta_capacity,
                })?;
            require_exact_capacity(&formula_list_deltas, delta_capacity, path)?;
            for change in &changes {
                let coordinate = (change.position().row(), change.position().column());
                let old_formula_key = formula_source_hosts
                    .binary_search_by_key(&coordinate, |host| (host.row, host.column))
                    .ok()
                    .map(|index| formula_source_hosts[index].key);
                let new_formula = authored
                    .formulas
                    .binary_search_by_key(&coordinate, |formula| {
                        (formula.position.row(), formula.position.column())
                    })
                    .ok()
                    .map(|index| authored.formulas[index].bytes.as_slice());
                if old_formula_key.is_none() && new_formula.is_none() {
                    continue;
                }
                formula_list_deltas.push(formula_list::HostDelta {
                    row: coordinate.0,
                    column: coordinate.1,
                    old_formula_key,
                    new_formula,
                });
            }
            if formula_list_deltas.len() != delta_capacity {
                return Err(Error::Verification { path });
            }
            if formula_list_deltas
                .windows(2)
                .any(|pair| (pair[0].row, pair[0].column) >= (pair[1].row, pair[1].column))
            {
                return Err(Error::Verification { path });
            }
            reserve_retained_vec::<formula_list::SourceMessage<'_>>(
                &mut budget,
                existing.formula_list.segments.len(),
                path,
            )?;
            formula_list_segments
                .try_reserve_exact(existing.formula_list.segments.len())
                .map_err(|_| Error::Allocation {
                    kind: crate::package::table_cells::LimitKind::RetainedElements,
                    amount: existing.formula_list.segments.len(),
                })?;
            require_exact_capacity(
                &formula_list_segments,
                existing.formula_list.segments.len(),
                path,
            )?;
            for message in &existing.formula_list.segments {
                formula_list_segments.push(formula_list::SourceMessage {
                    object_id: message.object_id,
                    payload: message.payload,
                    object_references: message.object_references,
                });
            }
            let remaining = budget.remaining()?;
            let package_wire = source.state.options.archive().max_iwa_stream_bytes();
            let package_work = package_wire
                .checked_mul(32)
                .ok_or(Error::InvalidSource { path })?;
            let list_limits = formula_list::Limits {
                max_input_bytes: usize::try_from(remaining.wire_bytes)
                    .map_err(|_| Error::InvalidSource { path })?
                    .min(package_wire),
                max_output_bytes: usize::try_from(remaining.output_bytes)
                    .map_err(|_| Error::InvalidSource { path })?
                    .min(package_wire),
                max_fields: usize::try_from(remaining.wire_fields)
                    .map_err(|_| Error::InvalidSource { path })?
                    .min(package_wire),
                max_work: usize::try_from(remaining.wire_work.min(remaining.transaction_work))
                    .map_err(|_| Error::InvalidSource { path })?
                    .min(package_work),
                max_entries: usize::try_from(remaining.retained_elements)
                    .map_err(|_| Error::InvalidSource { path })?,
                max_hosts: usize::try_from(remaining.retained_elements)
                    .map_err(|_| Error::InvalidSource { path })?,
                max_references: usize::try_from(remaining.references)
                    .map_err(|_| Error::InvalidSource { path })?
                    .min(source.state.options.semantic().max_references()),
                max_retained_elements: usize::try_from(remaining.retained_elements)
                    .map_err(|_| Error::InvalidSource { path })?,
                max_retained_bytes: usize::try_from(remaining.retained_bytes)
                    .map_err(|_| Error::InvalidSource { path })?,
                max_scratch_bytes: usize::try_from(remaining.peak_scratch_bytes)
                    .map_err(|_| Error::InvalidSource { path })?,
                max_allocations: usize::try_from(remaining.allocation_events)
                    .map_err(|_| Error::InvalidSource { path })?,
            };
            let list_source = formula_list::SourceList {
                root: formula_list::SourceMessage {
                    object_id: existing.formula_list.root.object_id,
                    payload: existing.formula_list.root.payload,
                    object_references: existing.formula_list.root.object_references,
                },
                segments: &formula_list_segments,
                expected_entries: existing.formula_list.expected_entries,
                source_hosts: &formula_source_hosts,
            };
            let preflight_work = formula_list::preflight_work_upper(
                list_source.segments.len(),
                list_source.source_hosts.len(),
                formula_list_deltas.len(),
            )
            .map_err(|error| map_formula_list_error(error, path))?;
            let preflight_usage = budget::Usage {
                transaction_work: usize_u64(preflight_work),
                ..budget::Usage::default()
            };
            budget.authorize(preflight_usage)?;
            let prepared_list =
                match formula_list::prepare(list_source, &formula_list_deltas, list_limits) {
                    Ok(prepared) => prepared,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(map_formula_list_error(error, path));
                    },
                };
            budget.record_authorized(preflight_usage)?;
            let requirements = prepared_list
                .requirements()
                .map_err(|error| map_formula_list_error(error, path))?;
            let list_envelope = formula_list_requirements_usage(requirements, path)?;
            budget.authorize(list_envelope)?;
            let logical_list = match prepared_list.logical() {
                Ok(logical) => logical,
                Err(error) => {
                    budget.cancel_authorization();
                    return Err(map_formula_list_error(error, path));
                },
            };
            let logical_list_view = logical_list.logical_view();
            if logical_list_view.assignments.len() != formula_list_deltas.len()
                || logical_list_view.entries.len() > requirements.retained_elements
            {
                budget.cancel_authorization();
                return Err(Error::Verification { path });
            }
            let logical_report = logical_list.prepare_report();
            let logical_usage = formula_list_logical_usage(logical_report, path)?;
            budget.record_authorized(logical_usage)?;
            let dependency_tiles =
                prepare_formula_dependency_tiles(source, &target, authored, &mut budget, path)?;
            let prepared_metadata = prepare_narrow_formula_metadata(
                source,
                &target,
                &formula_context,
                &existing,
                &formula_list_deltas,
                &dependency_tiles,
                &mut budget,
                path,
            )?;
            if dependency_tiles
                .assignments
                .iter()
                .any(|tile| tile.uuid.lower() != tile.object_id)
                || dependency_tiles.component_index
                    != target
                        .dependencies
                        .engine
                        .ok_or(Error::InvalidSource { path })?
                        .component_index
                || dependency_tiles
                    .inspection
                    .as_ref()
                    .is_some_and(|inspection| inspection.route.message_type != 11_006)
                || dependency_tiles.inspection.is_some() == dependency_tiles.assignments.is_empty()
                || dependency_tiles.retained_elements < dependency_tiles.assignments.len()
                || dependency_tiles.retained_bytes
                    < dependency_tiles
                        .assignments
                        .len()
                        .checked_mul(size_of::<FormulaDependencyTileAssignment>())
                        .ok_or(Error::InvalidSource { path })?
            {
                return Err(Error::InvalidSource { path });
            }
            let metadata_prepare_report = prepared_metadata.prepare_report();
            let metadata_execution = prepared_metadata.execution_requirements();
            let logical_metadata = prepared_metadata.logical_view();
            let final_table_formula_count = final_formula_table_count(
                existing.formula_list.hosts.len(),
                &formula_list_deltas,
                path,
            )?;
            if logical_metadata.formula_count != final_table_formula_count {
                return Err(Error::Verification { path });
            }
            if logical_list_view
                .assignments
                .iter()
                .enumerate()
                .any(|(delta, assignment)| assignment.delta != delta)
            {
                return Err(Error::Verification { path });
            }
            if !formula_context.external_targets.is_empty() {
                extend_formula_referenced_paths(
                    &mut formula_context.referenced_paths,
                    &formula_context.external_targets,
                    existing.identity.owner,
                    logical_metadata,
                    metadata_execution,
                    &mut budget,
                    path,
                )?;
            }
            let final_formulas = cache_commit::prepare_complete_final_formula_set(
                source,
                &target,
                &existing,
                logical_list_view,
                authored,
                &changes,
                &mut budget,
                path,
            )?;
            let owner = existing.identity.owner;
            let (dependency_owners, dependency_records) =
                cache_commit::prepare_dependency_payloads(source, &target, &mut budget, path)?;
            let dependency_ranges = cache_commit::prepare_range_dependency_payloads(
                source,
                &target,
                &mut budget,
                path,
            )?;
            let (formula_tables, formula_baseline) = prepare_formula_cache_registry(
                source,
                &target,
                &formula_context,
                &mut budget,
                path,
            )?;
            let final_overlay =
                cache_commit::prepare_final_overlay(&target, &changes, &mut budget, path)?;
            let engine = message_payload(
                source,
                target
                    .dependencies
                    .engine
                    .ok_or(Error::Verification { path })?,
                path,
            )?;
            let authored_cache_result = cache_commit::prepare_final_cache_from_sets(
                source,
                &target,
                table,
                &final_overlay,
                cache_commit::FinalDependencySet {
                    engine,
                    owners: &dependency_owners,
                    record_tiles: &dependency_records,
                    range_tiles: &dependency_ranges,
                },
                cache::FinalFormulaSet {
                    tables: &formula_tables,
                    source_cells: &existing.cells,
                    cells: &final_formulas.cells,
                    entries: &final_formulas.entries,
                    payloads: &final_formulas.payloads,
                    authored: &final_formulas.authored,
                },
                &formula_baseline,
                Some(logical_metadata),
                &mut budget,
                path,
            );
            let final_formula_retained_elements = final_formulas.retained_elements;
            let final_formula_retained_bytes = final_formulas.retained_bytes;
            drop(final_formulas);
            budget.release_retained(
                usize_u64(final_formula_retained_elements),
                usize_u64(final_formula_retained_bytes),
            )?;
            let authored_cache = match authored_cache_result {
                Ok(prepared) => prepared,
                Err(error) => {
                    budget.release_retained(
                        usize_u64(logical_report.retained_elements),
                        usize_u64(logical_report.retained_bytes),
                    )?;
                    budget.release_retained(
                        usize_u64(metadata_prepare_report.retained_elements),
                        usize_u64(metadata_prepare_report.retained_bytes),
                    )?;
                    return Err(error);
                },
            };
            reserve_retained_vec::<Option<tile::BncChange>>(&mut budget, changes.len(), path)?;
            reserve_retained_vec::<usize>(&mut budget, authored.formulas.len(), path)?;
            formula_bnc_changes
                .try_reserve_exact(changes.len())
                .map_err(|_| Error::Allocation {
                    kind: crate::package::table_cells::LimitKind::RetainedElements,
                    amount: changes.len(),
                })?;
            require_exact_capacity(&formula_bnc_changes, changes.len(), path)?;
            formula_text_change_indices
                .try_reserve_exact(authored.formulas.len())
                .map_err(|_| Error::Allocation {
                    kind: crate::package::table_cells::LimitKind::RetainedElements,
                    amount: authored.formulas.len(),
                })?;
            require_exact_capacity(&formula_text_change_indices, authored.formulas.len(), path)?;
            formula_bnc_changes.resize_with(changes.len(), || None);
            for assignment in logical_list_view.assignments {
                let delta = formula_list_deltas
                    .get(assignment.delta)
                    .ok_or(Error::Verification { path })?;
                let change_index = changes
                    .binary_search_by_key(&(delta.row, delta.column), |change| {
                        (change.position().row(), change.position().column())
                    })
                    .map_err(|_| Error::Verification { path })?;
                let _change = &changes[change_index];
                let Some(identifier) = assignment.key else {
                    formula_bnc_changes[change_index] = Some(tile::BncChange::FormulaClear);
                    continue;
                };
                let coordinate = cache::Coordinate {
                    row: delta.row,
                    column: delta.column,
                };
                let cache_cell = authored_cache
                    .rewrites
                    .iter()
                    .flat_map(|rewrite| &rewrite.cells)
                    .find(|cell| cell.owner == owner && cell.coordinate == coordinate);
                let cache = match cache_cell.map(|cell| &cell.value) {
                    None => None,
                    Some(litchi_iwa_common::formula::FormulaCachedValue::Number(value)) => {
                        Some(tile::ScalarInput::Number(*value))
                    },
                    Some(litchi_iwa_common::formula::FormulaCachedValue::Boolean(value)) => {
                        Some(tile::ScalarInput::Boolean(*value))
                    },
                    Some(litchi_iwa_common::formula::FormulaCachedValue::Date(value)) => {
                        Some(tile::ScalarInput::Date(*value))
                    },
                    Some(litchi_iwa_common::formula::FormulaCachedValue::Duration(value)) => {
                        Some(tile::ScalarInput::Duration(*value))
                    },
                    Some(litchi_iwa_common::formula::FormulaCachedValue::Text(_)) => {
                        formula_text_change_indices.push(change_index);
                        None
                    },
                };
                formula_bnc_changes[change_index] =
                    Some(tile::BncChange::FormulaSet { identifier, cache });
            }
            if formula_bnc_changes
                .iter()
                .filter(|change| change.is_some())
                .count()
                != logical_list_view.assignments.len()
            {
                return Err(Error::Verification { path });
            }
            prepared_formula_list_limits = Some(list_limits);
            prepared_formula_list = Some(logical_list);
            prepared_formula_metadata = Some(prepared_metadata);
            formula_dependency_tiles = Some(dependency_tiles);
            formula_refreshed_hosts = authored_cache.refreshed_hosts;
            formula_cache_rewrites = authored_cache.rewrites;
        }
    }
    if let Some(dependency_tiles) = formula_dependency_tiles
        .as_ref()
        .filter(|tiles| !tiles.assignments.is_empty())
    {
        let inspection = dependency_tiles
            .inspection
            .as_ref()
            .ok_or(Error::InvalidSource { path })?;
        let additions = dependency_tiles.assignments.len();
        reserve_retained_vec::<ObjectUuidAddition<'_>>(&mut budget, additions, path)?;
        formula_dependency_uuid_additions
            .try_reserve_exact(additions)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::References,
                amount: additions,
            })?;
        require_exact_capacity(&formula_dependency_uuid_additions, additions, path)?;
        let selector = inspection.selector(source, dependency_tiles.component_index, path)?;
        for assignment in &dependency_tiles.assignments {
            formula_dependency_uuid_additions.push(ObjectUuidAddition::new(
                selector,
                assignment.object_id,
                assignment.uuid,
            ));
        }
        let construction_usage = budget::Usage {
            transaction_work: usize_u64(additions),
            ..budget::Usage::default()
        };
        budget.reserve(construction_usage)?;
        let last = dependency_tiles
            .assignments
            .last()
            .map(|assignment| assignment.object_id)
            .ok_or(Error::InvalidSource { path })?;
        let remaining = budget.remaining()?;
        let options = sparse_commit::metadata_options(source, remaining, additions, path)?;
        budget.authorize(remaining)?;
        let prepared = match prepare_package_metadata_rewrite(
            message_payload(source, inspection.route, path)?,
            Batch::new(
                inspection.last_object_identifier,
                last,
                &formula_dependency_uuid_additions,
                &[],
            ),
            options,
        ) {
            Ok(prepared) => prepared,
            Err(error) => {
                budget.cancel_authorization();
                return Err(sparse_commit::map_metadata_error(error, path));
            },
        };
        budget.record_authorized(sparse_commit::package_metadata_plan_usage(
            prepared.prepare_report(),
        )?)?;
        prepared_package_metadata = Some(prepared);
    }
    let (scalar_formula_tables, scalar_formula_baseline) = if formula_publication
        || !needs_formula_context
    {
        (Vec::new(), Vec::new())
    } else {
        // A scalar edit can dirty formulas authored by an earlier transaction.
        // Until the existing graph has been inspected, every resolved
        // same-engine table is a possible evaluator input. Keep full geometry
        // and a lazy semantic route for that complete owner registry so the
        // scalar cache path cannot silently fall back to selected-only
        // authority or eagerly materialize unrelated tables.
        let path_count = formula_context.external_targets.len();
        reserve_retained_vec::<formula_author::ReferencedTablePath>(&mut budget, path_count, path)?;
        let mut referenced_paths = Vec::new();
        referenced_paths
            .try_reserve_exact(path_count)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: path_count,
            })?;
        require_exact_capacity(&referenced_paths, path_count, path)?;
        referenced_paths.extend(
            formula_context
                .external_targets
                .iter()
                .map(|(referenced, _)| *referenced),
        );
        if referenced_paths.windows(2).any(|pair| pair[0] >= pair[1]) {
            return Err(Error::InvalidSource { path });
        }
        let previous = core::mem::replace(&mut formula_context.referenced_paths, referenced_paths);
        let previous_bytes = previous
            .capacity()
            .checked_mul(size_of::<formula_author::ReferencedTablePath>())
            .ok_or(Error::InvalidSource { path })?;
        budget.release_retained(usize_u64(previous.capacity()), usize_u64(previous_bytes))?;
        prepare_formula_cache_registry(source, &target, &formula_context, &mut budget, path)?
    };
    let cache_prepared = if formula_publication {
        cache_commit::PreparedCache::default()
    } else {
        let prepared = cache_commit::prepare_final_cache(
            source,
            &target,
            table,
            &changes,
            &scalar_formula_tables,
            &scalar_formula_baseline,
            &mut budget,
            path,
        )?;
        let registry_elements = scalar_formula_tables
            .capacity()
            .checked_add(scalar_formula_baseline.capacity())
            .ok_or(Error::InvalidSource { path })?;
        let registry_bytes = scalar_formula_tables
            .capacity()
            .checked_mul(size_of::<cache::TableGeometry>())
            .and_then(|bytes| {
                scalar_formula_baseline
                    .capacity()
                    .checked_mul(size_of::<cache_commit::ExternalBaselineTable<'_>>())
                    .and_then(|baseline| bytes.checked_add(baseline))
            })
            .ok_or(Error::InvalidSource { path })?;
        drop(scalar_formula_tables);
        drop(scalar_formula_baseline);
        budget.release_retained(usize_u64(registry_elements), usize_u64(registry_bytes))?;
        prepared
    };
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

    let mut rich_prepared = rich_commit::prepare_unique_rich_text_to_text(
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
        || rich_prepared.owned_transition_count != usize::from(rich_prepared.transition.is_some())
    {
        return Err(Error::Verification { path });
    }

    let scalar_text_count = changes
        .iter()
        .enumerate()
        .filter(|(index, change)| {
            rich_prepared.keys[*index].is_none()
                && matches!(change.input_ref(), Some(Input::Text(_)))
        })
        .count();
    let text_count = scalar_text_count
        .checked_add(formula_text_change_indices.len())
        .ok_or(Error::InvalidSource { path })?;
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
    for &change_index in &formula_text_change_indices {
        let change = changes
            .get(change_index)
            .ok_or(Error::Verification { path })?;
        let coordinate = cache::Coordinate {
            row: change.position().row(),
            column: change.position().column(),
        };
        let text = formula_cache_rewrites
            .iter()
            .flat_map(|rewrite| &rewrite.cells)
            .find(|cell| cell.coordinate == coordinate)
            .and_then(|cell| match &cell.value {
                litchi_iwa_common::formula::FormulaCachedValue::Text(text) => Some(text.as_str()),
                _ => None,
            })
            .ok_or(Error::Verification { path })?;
        text_requests.push(lists::StringRequest::new(text));
        text_change_indices.push(change_index);
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
        let preliminary_report = preliminary.report();
        let usage = assignment_usage(&budget, preliminary_report)?;
        budget.record_authorized(usage)?;
        let assignments_match = preliminary.assignments().iter().all(|assignment| {
            let Some(&change_index) = text_change_indices.get(assignment.request()) else {
                return false;
            };
            let Some(slot) = string_keys.get_mut(change_index) else {
                return false;
            };
            *slot = Some(assignment.key());
            true
        });
        drop(preliminary);
        budget.release_retained(
            usize_u64(preliminary_report.requests()),
            usize_u64(preliminary_report.retained_bytes()),
        )?;
        if !assignments_match {
            return Err(Error::Verification { path });
        }
        for &change_index in &formula_text_change_indices {
            let key = string_keys
                .get(change_index)
                .copied()
                .flatten()
                .ok_or(Error::Verification { path })?;
            let planned = formula_bnc_changes
                .get_mut(change_index)
                .and_then(Option::as_mut)
                .ok_or(Error::Verification { path })?;
            match planned {
                tile::BncChange::FormulaSet { cache, .. } if cache.is_none() => {
                    *cache = Some(tile::ScalarInput::String(key));
                },
                _ => return Err(Error::Verification { path }),
            }
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
    let formula_list_binding_capacity = if prepared_formula_list.is_some() {
        target
            .storage
            .lists
            .formula
            .segments
            .len()
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?
    } else {
        0
    };
    let formula_metadata_binding_capacity = if prepared_formula_metadata.is_some() {
        formula_metadata_binding_capacity(
            &target,
            formula_dependency_tiles
                .as_ref()
                .ok_or(Error::Verification { path })?,
            path,
        )?
    } else {
        0
    };
    let prepared_capacity = rich_prepared
        .owned_transition_count
        .checked_add(formula_list_binding_capacity)
        .and_then(|count| count.checked_add(formula_metadata_binding_capacity))
        .and_then(|count| count.checked_add(usize::from(prepared_package_metadata.is_some())))
        .and_then(|count| count.checked_add(changes.len()))
        .and_then(|count| count.checked_add(cache_prepared.tiles.len()))
        .and_then(|count| count.checked_add(final_row_capacity))
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::LimitExceeded {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            observed: u64::MAX,
            maximum: budget.limits().max_retained_elements,
            path,
        })?;
    let mut prepared: Vec<PreparedReplacement> = Vec::new();
    reserve_retained_vec::<PreparedReplacement>(&mut budget, prepared_capacity, path)?;
    prepared
        .try_reserve_exact(prepared_capacity)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: changes.len(),
        })?;
    require_exact_capacity(&prepared, prepared_capacity, path)?;
    let mut releases: Vec<(u32, u32)> = Vec::new();
    reserve_retained_vec::<(u32, u32)>(&mut budget, changes.len(), path)?;
    releases
        .try_reserve_exact(changes.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: changes.len(),
        })?;
    require_exact_capacity(&releases, changes.len(), path)?;
    let mut final_rows = Vec::new();
    reserve_retained_vec::<sparse::FinalRowCount>(&mut budget, final_row_capacity, path)?;
    final_rows
        .try_reserve_exact(final_row_capacity)
        .map_err(|_error| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: final_row_capacity,
        })?;
    let prepared_tile_capacity = changes
        .len()
        .checked_add(cache_prepared.tiles.len())
        .ok_or(Error::InvalidSource { path })?;
    let mut prepared_tiles = Vec::new();
    reserve_retained_vec::<PreparedTilePublication<'_>>(&mut budget, prepared_tile_capacity, path)?;
    prepared_tiles
        .try_reserve_exact(prepared_tile_capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: prepared_tile_capacity,
        })?;
    require_exact_capacity(&prepared_tiles, prepared_tile_capacity, path)?;
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
            let change_index = start + offset;
            let planned_formula = formula_bnc_changes
                .get_mut(change_index)
                .and_then(Option::take);
            tile_changes.push(tile::TileChange {
                row: change.position().row() % target.storage.tile_size,
                column: change.position().column(),
                change: match planned_formula {
                    Some(change) => change,
                    None => scalar_change(
                        change,
                        string_keys.get(change_index).copied().flatten(),
                        rich_prepared.keys.get(change_index).copied().flatten(),
                        path,
                    )?,
                },
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
        let maximum_writes = tile_changes
            .len()
            .checked_add(cache_changes.len())
            .ok_or(Error::InvalidSource { path })?;
        budget.authorize(remaining)?;
        let max_input_bytes = usize::try_from(remaining.wire_bytes)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire);
        let max_output_bytes = usize::try_from(
            remaining
                .output_bytes
                .min(remaining.retained_bytes)
                .min(remaining.peak_scratch_bytes),
        )
        .map_err(|_error| Error::InvalidSource { path })?
        .min(maximum_wire);
        let max_fields = usize::try_from(remaining.wire_fields)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire);
        let max_work = budget::tile_work_ceiling(
            remaining,
            usize_u64(maximum_work),
            usize_u64(maximum_writes),
        )?;
        let tile_change_elements = tile_changes.capacity();
        let tile_change_bytes = tile_change_elements
            .checked_mul(size_of::<tile::TileChange>())
            .ok_or(Error::InvalidSource { path })?;
        let plan = match tile::prepare_tile_with_cache(
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
                )
                .with_accounting(
                    usize::try_from(remaining.retained_bytes)
                        .map_err(|_| Error::InvalidSource { path })?,
                    usize::try_from(remaining.retained_elements)
                        .map_err(|_| Error::InvalidSource { path })?,
                    usize::try_from(remaining.peak_scratch_bytes)
                        .map_err(|_| Error::InvalidSource { path })?,
                    usize::try_from(remaining.allocation_events)
                        .map_err(|_| Error::InvalidSource { path })?,
                ),
            },
            cache_changes,
        ) {
            Ok(plan) => plan,
            Err(error) => {
                budget.cancel_authorization();
                drop(tile_changes);
                budget.release_retained(
                    usize_u64(tile_change_elements),
                    usize_u64(tile_change_bytes),
                )?;
                return Err(map_tile_error(error, path));
            },
        };
        let prepare_report = plan.prepare_report().report();
        let prepare_usage = tile_usage(&budget, prepare_report, 1)?;
        budget.record_authorized(prepare_usage)?;
        drop(tile_changes);
        budget.release_retained(
            usize_u64(tile_change_elements),
            usize_u64(tile_change_bytes),
        )?;
        let visit_work = plan
            .transition_visit_work()
            .map_err(|error| map_tile_error(error, path))?;
        budget.reserve(budget::Usage {
            transaction_work: usize_u64(visit_work),
            ..budget::Usage::default()
        })?;
        for row in plan.final_rows() {
            let global_row = tile_id
                .checked_mul(target.storage.tile_size)
                .and_then(|base| base.checked_add(row.row))
                .ok_or(Error::InvalidSource { path })?;
            final_rows.push(sparse::FinalRowCount {
                row: global_row,
                number_of_cells: row.cell_count,
            });
        }
        let expected_transitions = validate_prepared_tile_transitions(
            &plan,
            &changes[start..end],
            &rich_prepared.keys[start..end],
            target.storage.tile_size,
            formula_publication,
            &mut releases,
            path,
        )?;
        if plan.execution_requirements().output_bytes() == 0
            && (rich_prepared.keys[start..end].iter().any(Option::is_none)
                || !cache_changes.is_empty())
        {
            return Err(Error::Verification { path });
        }
        prepared_tiles.push(PreparedTilePublication {
            route: route.message,
            plan,
            expected_transitions,
        });
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
    let header_plan_capacity = usize::from(!final_rows.is_empty())
        .checked_add(
            final_rows
                .windows(2)
                .filter(|pair| {
                    pair[0].row / sparse::HEADER_BUCKET_ROWS
                        != pair[1].row / sparse::HEADER_BUCKET_ROWS
                })
                .count(),
        )
        .ok_or(Error::InvalidSource { path })?;
    let mut prepared_headers = Vec::new();
    reserve_retained_vec::<PreparedHeaderPublication<'_>>(&mut budget, header_plan_capacity, path)?;
    prepared_headers
        .try_reserve_exact(header_plan_capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: header_plan_capacity,
        })?;
    require_exact_capacity(&prepared_headers, header_plan_capacity, path)?;
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
        let plan = sparse::prepare_existing_header_bucket_final_rows(
            message_payload(source, route, path)?,
            bucket_index,
            target.native.columns,
            &final_rows[row_start..row_end],
            limits,
        );
        let plan = match plan {
            Ok(plan) => plan,
            Err(error) => {
                budget.cancel_authorization();
                return Err(sparse_commit::map_sparse_error(error, path));
            },
        };
        let usage = match sparse_commit::sparse_usage(plan.prepare_report().report(), path) {
            Ok(usage) => usage,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        budget.record_authorized(usage)?;
        prepared_headers.push(PreparedHeaderPublication { route, plan });
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
    let mut prepared_string_list = None;
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
        let prepared_list = match lists::prepare_string_list(
            string_list_payload,
            &releases,
            &text_requests,
            limits,
        ) {
            Ok(plan) => plan,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_list_error(error, path));
            },
        };
        let prepare_report = prepared_list.prepare_report();
        let prepare_usage = match string_list_prepare_usage(prepare_report, path) {
            Ok(usage) => usage,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        budget.record_authorized(prepare_usage)?;
        if !prepared_list.assignments().iter().all(|assignment| {
            text_change_indices
                .get(assignment.request())
                .and_then(|change_index| string_keys.get(*change_index))
                .copied()
                .flatten()
                == Some(assignment.key())
        }) {
            budget.release_retained(
                usize_u64(prepare_report.retained_elements()),
                usize_u64(prepare_report.retained_bytes()),
            )?;
            return Err(Error::Verification { path });
        }
        prepared_string_list = Some(prepared_list);
    }

    let future_component_capacity = prepared_capacity
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    reserve_retained_vec::<usize>(&mut budget, future_component_capacity, path)?;
    let future_component_work = future_component_capacity
        .checked_add(sort_work(future_component_capacity, path)?)
        .ok_or(Error::InvalidSource { path })?;
    budget.reserve(budget::Usage {
        transaction_work: usize_u64(future_component_work),
        ..budget::Usage::default()
    })?;
    let mut future_components = Vec::new();
    future_components
        .try_reserve_exact(future_component_capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: future_component_capacity,
        })?;
    require_exact_capacity(&future_components, future_component_capacity, path)?;
    future_components.extend(
        prepared
            .iter()
            .map(|replacement| replacement.route.component_index),
    );
    if let Some(transition) = &rich_prepared.transition {
        future_components.push(transition.route().component_index);
    }
    future_components.extend(
        prepared_tiles
            .iter()
            .map(|prepared| prepared.route.component_index),
    );
    future_components.extend(
        prepared_headers
            .iter()
            .map(|prepared| prepared.route.component_index),
    );
    if prepared_string_list.is_some() {
        future_components.push(target.storage.lists.string.message.component_index);
    }
    if prepared_formula_list.is_some() {
        future_components.push(target.storage.lists.formula.message.component_index);
        future_components.extend(
            target
                .storage
                .lists
                .formula
                .segments
                .iter()
                .map(|route| route.component_index),
        );
    }
    if prepared_formula_metadata.is_some() {
        future_components.push(
            target
                .dependencies
                .engine
                .ok_or(Error::Verification { path })?
                .component_index,
        );
        let selected_owner = target
            .dependencies
            .selected_formula_owner
            .ok_or(Error::Verification { path })?;
        for owner in &target.dependencies.formula_owners {
            if owner.internal_owner_id != selected_owner.internal_owner_id {
                continue;
            }
            future_components.push(owner.message.component_index);
            future_components.extend(
                owner
                    .cell_record_tiles
                    .iter()
                    .map(|route| route.component_index),
            );
            future_components.extend(
                owner
                    .range_precedent_tiles
                    .iter()
                    .map(|route| route.message.component_index),
            );
        }
    }
    if let Some(inspection) = formula_dependency_tiles
        .as_ref()
        .and_then(|tiles| tiles.inspection.as_ref())
    {
        future_components.push(inspection.route.component_index);
    }
    if let Some(dependency_tiles) = &formula_dependency_tiles {
        if !dependency_tiles.assignments.is_empty() {
            future_components.push(dependency_tiles.component_index);
        }
    }
    if future_components.len() > future_component_capacity {
        return Err(Error::Verification { path });
    }
    future_components.sort_unstable();
    future_components.dedup();
    if future_components.is_empty() {
        return Err(Error::Verification { path });
    }
    let mut future_payload_bytes = prepared.iter().try_fold(0usize, |total, replacement| {
        total
            .checked_add(replacement.payload.len())
            .ok_or(Error::InvalidSource { path })
    })?;
    if let Some(transition) = &rich_prepared.transition {
        future_payload_bytes = future_payload_bytes
            .checked_add(transition.execution_requirements().output_bytes)
            .ok_or(Error::InvalidSource { path })?;
    }
    for prepared_tile in &prepared_tiles {
        future_payload_bytes = future_payload_bytes
            .checked_add(prepared_tile.plan.execution_requirements().output_bytes())
            .ok_or(Error::InvalidSource { path })?;
    }
    for prepared_header in &prepared_headers {
        future_payload_bytes = future_payload_bytes
            .checked_add(prepared_header.plan.execution_requirements().output_bytes)
            .ok_or(Error::InvalidSource { path })?;
    }
    if let Some(prepared_list) = &prepared_string_list {
        future_payload_bytes = future_payload_bytes
            .checked_add(prepared_list.execution_requirements().output_bytes())
            .ok_or(Error::InvalidSource { path })?;
    }
    if let Some(logical_list) = &prepared_formula_list {
        future_payload_bytes = future_payload_bytes
            .checked_add(logical_list.execution_requirements().output_bytes)
            .ok_or(Error::InvalidSource { path })?;
    }
    if let Some(prepared_metadata) = &prepared_formula_metadata {
        future_payload_bytes = future_payload_bytes
            .checked_add(prepared_metadata.execution_requirements().output_bytes)
            .ok_or(Error::InvalidSource { path })?;
    }
    if let Some(prepared_metadata) = &prepared_package_metadata {
        future_payload_bytes = future_payload_bytes
            .checked_add(prepared_metadata.execution_requirements().output_bytes())
            .ok_or(Error::InvalidSource { path })?;
    }
    let mut future_reference_items = prepared.iter().try_fold(0usize, |total, replacement| {
        replacement
            .references
            .as_ref()
            .map_or(Ok(total), |references| {
                total
                    .checked_add(references.before.len())
                    .and_then(|value| {
                        references
                            .after
                            .len()
                            .checked_mul(2)
                            .and_then(|after| value.checked_add(after))
                    })
                    .ok_or(Error::InvalidSource { path })
            })
    })?;
    if let Some(transition) = &rich_prepared.transition {
        future_reference_items = future_reference_items
            .checked_add(transition.publication_reference_items())
            .ok_or(Error::InvalidSource { path })?;
    }
    if prepared_formula_metadata.is_some() {
        let dependency_tiles = formula_dependency_tiles
            .as_ref()
            .ok_or(Error::Verification { path })?;
        let selected_owner = target
            .dependencies
            .selected_formula_owner
            .ok_or(Error::Verification { path })?;
        let selected_route = target
            .dependencies
            .formula_owners
            .iter()
            .find(|owner| owner.internal_owner_id == selected_owner.internal_owner_id)
            .ok_or(Error::Verification { path })?
            .message;
        let before = source_message_all_references(source, selected_route, path)?.len();
        future_reference_items = future_reference_items
            .checked_add(before.checked_mul(3).ok_or(Error::InvalidSource { path })?)
            .and_then(|value| {
                dependency_tiles
                    .assignments
                    .len()
                    .checked_mul(2)
                    .and_then(|assignments| value.checked_add(assignments))
            })
            .ok_or(Error::InvalidSource { path })?;
    }
    let appended_objects = formula_dependency_tiles
        .as_ref()
        .map_or(0, |tiles| tiles.assignments.len());
    let remaining = budget.remaining()?;
    budget.authorize(remaining)?;
    let future_publication = match rewrite::future_publication_reservation(
        source,
        &future_components,
        prepared_capacity,
        future_payload_bytes,
        future_reference_items,
        appended_objects,
    ) {
        Ok(requirements) => requirements,
        Err(error) => {
            budget.cancel_authorization();
            return Err(map_rewrite_error(error, path));
        },
    };
    budget.record_authorized(budget::Usage {
        lookups: future_publication.planning_lookups,
        transaction_work: future_publication.planning_work,
        ..budget::Usage::default()
    })?;
    budget.release_retained(
        usize_u64(future_components.capacity()),
        usize_u64(
            future_components
                .capacity()
                .checked_mul(size_of::<usize>())
                .ok_or(Error::InvalidSource { path })?,
        ),
    )?;

    let prepared_tile_elements = prepared_tiles.capacity();
    let prepared_tile_bytes = prepared_tile_elements
        .checked_mul(size_of::<PreparedTilePublication<'_>>())
        .ok_or(Error::InvalidSource { path })?;
    let prepared_header_elements = prepared_headers.capacity();
    let prepared_header_bytes = prepared_header_elements
        .checked_mul(size_of::<PreparedHeaderPublication<'_>>())
        .ok_or(Error::InvalidSource { path })?;
    let prepared_plan_bytes = prepared_tiles.iter().try_fold(
        prepared_tile_bytes
            .checked_add(prepared_header_bytes)
            .ok_or(Error::InvalidSource { path })?,
        |total, prepared| {
            total
                .checked_add(
                    prepared
                        .plan
                        .prepare_report()
                        .report()
                        .retained_bytes
                        .try_into()
                        .map_err(|_| Error::InvalidSource { path })?,
                )
                .ok_or(Error::InvalidSource { path })
        },
    )?;
    let prepared_plan_bytes =
        prepared_headers
            .iter()
            .try_fold(prepared_plan_bytes, |total, prepared| {
                total
                    .checked_add(prepared.plan.prepare_report().report().retained_bytes)
                    .ok_or(Error::InvalidSource { path })
            })?;
    let prepared_plan_bytes =
        prepared_string_list
            .as_ref()
            .map_or(Ok(prepared_plan_bytes), |prepared_list| {
                prepared_plan_bytes
                    .checked_add(prepared_list.prepare_report().retained_bytes())
                    .ok_or(Error::InvalidSource { path })
            })?;
    let prepared_plan_bytes =
        prepared_formula_list
            .as_ref()
            .map_or(Ok(prepared_plan_bytes), |logical_list| {
                prepared_plan_bytes
                    .checked_add(logical_list.prepare_report().retained_bytes)
                    .ok_or(Error::InvalidSource { path })
            })?;
    let prepared_plan_bytes = prepared_formula_metadata.as_ref().map_or(
        Ok(prepared_plan_bytes),
        |prepared_metadata| {
            prepared_plan_bytes
                .checked_add(prepared_metadata.prepare_report().retained_bytes)
                .ok_or(Error::InvalidSource { path })
        },
    )?;
    let prepared_plan_bytes =
        rich_prepared
            .transition
            .as_ref()
            .map_or(Ok(prepared_plan_bytes), |transition| {
                prepared_plan_bytes
                    .checked_add(transition.retained_accounting(&budget, path)?.bytes)
                    .ok_or(Error::InvalidSource { path })
            })?;
    let prepared_plan_bytes = usize_u64(prepared_plan_bytes);
    let mut execution_envelope = budget::Usage::default();
    let mut aggregate_peak = 0u64;
    if let Some(transition) = &rich_prepared.transition {
        let requirements = transition.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            rich_commit::rich_execution_requirements_usage(requirements, path)?,
            path,
        )?;
        let own_plan = usize_u64(transition.retained_accounting(&budget, path)?.bytes);
        let transient = usize_u64(requirements.peak_scratch_bytes)
            .checked_sub(own_plan)
            .ok_or(Error::InvalidSource { path })?;
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(transient)
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    for prepared_tile in &prepared_tiles {
        let requirements = prepared_tile.plan.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            tile_execution_requirements_usage(&budget, requirements)?,
            path,
        )?;
        let own_plan = prepared_tile.plan.prepare_report().report().retained_bytes;
        let transient = usize_u64(requirements.peak_scratch_bytes())
            .checked_sub(own_plan)
            .ok_or(Error::InvalidSource { path })?;
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(transient)
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    merge_usage(
        &mut execution_envelope,
        component_reservation_usage(future_publication.component, path)?,
        path,
    )?;
    merge_usage(
        &mut execution_envelope,
        budget::publication_usage_reservation(future_publication.publication)?,
        path,
    )?;
    let prepublication_usage = if appended_objects == 0 {
        classic_staging_usage(
            prepared_capacity,
            prepared_capacity,
            future_reference_items,
            path,
        )?
    } else {
        sparse_commit::formula_dependency_prepublication_usage(
            source,
            prepared_capacity,
            appended_objects,
            prepared_capacity,
            future_reference_items,
            path,
        )?
    };
    merge_usage(&mut execution_envelope, prepublication_usage, path)?;
    aggregate_peak = aggregate_peak.max(future_publication.component.maximum_peak_bytes);
    for prepared_header in &prepared_headers {
        let requirements = prepared_header.plan.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            header_execution_requirements_usage(requirements, path)?,
            path,
        )?;
        let transient = usize_u64(requirements.peak_scratch_bytes);
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(transient)
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    if let Some(prepared_list) = &prepared_string_list {
        let requirements = prepared_list.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            string_list_execution_usage(requirements, path)?,
            path,
        )?;
        let own_plan = usize_u64(prepared_list.prepare_report().retained_bytes());
        let transient = usize_u64(requirements.peak_scratch_bytes())
            .checked_sub(own_plan)
            .ok_or(Error::InvalidSource { path })?;
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(transient)
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    if let Some(logical_list) = &prepared_formula_list {
        let requirements = logical_list.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            formula_list_execution_usage(requirements, path)?,
            path,
        )?;
        merge_usage(
            &mut execution_envelope,
            formula_list_binding_usage(&target.storage.lists.formula, path)?,
            path,
        )?;
        let own_plan = usize_u64(logical_list.prepare_report().retained_bytes);
        let transient = usize_u64(requirements.peak_scratch_bytes)
            .checked_sub(own_plan)
            .ok_or(Error::InvalidSource { path })?;
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(transient)
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    if let Some(prepared_metadata) = &prepared_formula_metadata {
        let requirements = prepared_metadata.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            formula_metadata_execution_usage(requirements, path)?,
            path,
        )?;
        merge_usage(
            &mut execution_envelope,
            formula_metadata_binding_usage(
                source,
                &target,
                formula_dependency_tiles
                    .as_ref()
                    .ok_or(Error::Verification { path })?,
                path,
            )?,
            path,
        )?;
        let own_plan = usize_u64(prepared_metadata.prepare_report().retained_bytes);
        let transient = usize_u64(requirements.peak_scratch_bytes)
            .checked_sub(own_plan)
            .ok_or(Error::InvalidSource { path })?;
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(transient)
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    if let Some(prepared_metadata) = &prepared_package_metadata {
        let requirements = prepared_metadata.execution_requirements();
        merge_usage(
            &mut execution_envelope,
            sparse_commit::package_metadata_execution_usage(requirements)?,
            path,
        )?;
        aggregate_peak = aggregate_peak.max(
            prepared_plan_bytes
                .checked_add(usize_u64(requirements.scratch_bytes()))
                .ok_or(Error::InvalidSource { path })?,
        );
    }
    execution_envelope.peak_scratch_bytes = aggregate_peak;
    #[cfg(test)]
    aggregate_testing::record_requirement(budget.required_limits_for(execution_envelope)?);
    #[cfg(test)]
    if let Some(limits) = aggregate_testing::authorization_limits() {
        budget.authorize_under_limits(execution_envelope, limits)?;
    } else {
        budget.authorize(execution_envelope)?;
    }
    #[cfg(not(test))]
    budget.authorize(execution_envelope)?;
    #[cfg(test)]
    aggregate_testing::record_execution();
    let mut execution_actual = budget::Usage::default();
    let mut released_plan_elements = 0u64;
    let mut released_plan_bytes = 0u64;
    let mut released_artifact_elements = 0u64;
    let mut released_artifact_bytes = 0u64;
    if let Some(transition) = rich_prepared.transition.take() {
        let retained = match transition.retained_accounting(&budget, path) {
            Ok(retained) => retained,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        let (replacement, report) = match transition.execute(&budget, path) {
            Ok(replacement) => replacement,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        merge_usage(
            &mut execution_actual,
            rich_commit::rich_execution_usage(report, path)?,
            path,
        )?;
        let disposable_reference_elements = replacement
            .references
            .removed
            .capacity()
            .checked_add(replacement.references.removed_by_field.capacity())
            .ok_or(Error::InvalidSource { path })?;
        let disposable_reference_bytes = replacement
            .references
            .removed
            .capacity()
            .checked_mul(size_of::<u64>())
            .and_then(|bytes| {
                replacement
                    .references
                    .removed_by_field
                    .capacity()
                    .checked_mul(size_of::<(u32, u64)>())
                    .and_then(|fields| bytes.checked_add(fields))
            })
            .ok_or(Error::InvalidSource { path })?;
        let references = match rich_archive_reference_delta(
            source,
            replacement.route,
            replacement.references,
            path,
        ) {
            Ok(references) => references,
            Err(error) => {
                budget.cancel_authorization();
                return Err(error);
            },
        };
        prepared.push(PreparedReplacement {
            route: replacement.route,
            payload: replacement.payload,
            references: Some(references),
        });
        released_plan_elements = released_plan_elements
            .checked_add(usize_u64(retained.elements))
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(usize_u64(retained.bytes))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_elements = released_artifact_elements
            .checked_add(usize_u64(disposable_reference_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_bytes = released_artifact_bytes
            .checked_add(usize_u64(disposable_reference_bytes))
            .ok_or(Error::InvalidSource { path })?;
    }
    for prepared_tile in prepared_tiles {
        let requirements = prepared_tile.plan.execution_requirements();
        let prepare_report = prepared_tile.plan.prepare_report().report();
        let outcome = match prepared_tile.plan.execute(requirements.exact_limits()) {
            Ok(outcome) => outcome,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_tile_error(error, path));
            },
        };
        merge_usage(
            &mut execution_actual,
            tile_usage(&budget, outcome.report, 0)?,
            path,
        )?;
        released_plan_elements = released_plan_elements
            .checked_add(prepare_report.retained_elements)
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(prepare_report.retained_bytes)
            .ok_or(Error::InvalidSource { path })?;
        if outcome.transitions.len() != prepared_tile.expected_transitions {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        }
        let artifact_elements = outcome
            .transitions
            .capacity()
            .checked_add(outcome.final_rows.capacity())
            .ok_or(Error::InvalidSource { path })?;
        let artifact_bytes = outcome
            .transitions
            .capacity()
            .checked_mul(size_of::<tile::CellTransition>())
            .and_then(|bytes| {
                outcome
                    .final_rows
                    .capacity()
                    .checked_mul(size_of::<tile::RowCellCount>())
                    .and_then(|rows| bytes.checked_add(rows))
            })
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_elements = released_artifact_elements
            .checked_add(usize_u64(artifact_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_bytes = released_artifact_bytes
            .checked_add(usize_u64(artifact_bytes))
            .ok_or(Error::InvalidSource { path })?;
        if let Some(payload) = outcome.payload {
            prepared.push(PreparedReplacement {
                route: prepared_tile.route,
                payload,
                references: None,
            });
        } else if requirements.output_bytes() != 0 {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        }
    }
    for prepared_header in prepared_headers {
        let requirements = prepared_header.plan.execution_requirements();
        let prepare_report = prepared_header.plan.prepare_report().report();
        let (payload, report) = match prepared_header.plan.execute(requirements.exact_limits()) {
            Ok(output) => output,
            Err(error) => {
                budget.cancel_authorization();
                return Err(sparse_commit::map_sparse_error(error, path));
            },
        };
        merge_usage(
            &mut execution_actual,
            sparse_commit::sparse_usage(report, path)?,
            path,
        )?;
        released_plan_elements = released_plan_elements
            .checked_add(usize_u64(prepare_report.retained_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(usize_u64(prepare_report.retained_bytes))
            .ok_or(Error::InvalidSource { path })?;
        if let Some(payload) = payload {
            prepared.push(PreparedReplacement {
                route: prepared_header.route,
                payload,
                references: None,
            });
        } else if requirements.output_bytes() != 0 {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        }
    }
    if let Some(prepared_list) = prepared_string_list.take() {
        let prepare_report = prepared_list.prepare_report();
        let requirements = prepared_list.execution_requirements();
        let final_list = match prepared_list.execute(requirements.exact_limits()) {
            Ok(plan) => plan,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_list_error(error, path));
            },
        };
        let artifact_matches = final_list.report().output_bytes() == requirements.output_bytes()
            && final_list.report().changed() == requirements.changed();
        merge_usage(
            &mut execution_actual,
            string_list_execution_usage(requirements, path)?,
            path,
        )?;
        released_plan_elements = released_plan_elements
            .checked_add(usize_u64(prepare_report.retained_elements()))
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(usize_u64(prepare_report.retained_bytes()))
            .ok_or(Error::InvalidSource { path })?;
        let assignment_elements = requirements
            .retained_elements()
            .checked_sub(requirements.output_bytes())
            .ok_or(Error::Verification { path })?;
        let assignment_bytes = requirements
            .retained_bytes()
            .checked_sub(requirements.output_bytes())
            .ok_or(Error::Verification { path })?;
        released_artifact_elements = released_artifact_elements
            .checked_add(usize_u64(assignment_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_bytes = released_artifact_bytes
            .checked_add(usize_u64(assignment_bytes))
            .ok_or(Error::InvalidSource { path })?;
        let changed = requirements.changed();
        let payload = final_list.into_payload();
        if !artifact_matches {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        }
        if changed {
            prepared.push(PreparedReplacement {
                route: target.storage.lists.string.message,
                payload,
                references: None,
            });
        } else if payload != string_list_payload {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        } else {
            released_artifact_elements = released_artifact_elements
                .checked_add(usize_u64(requirements.output_bytes()))
                .ok_or(Error::InvalidSource { path })?;
            released_artifact_bytes = released_artifact_bytes
                .checked_add(usize_u64(requirements.output_bytes()))
                .ok_or(Error::InvalidSource { path })?;
        }
    }
    if let Some(logical_list) = prepared_formula_list.take() {
        let logical_report = logical_list.prepare_report();
        let requirements = logical_list.execution_requirements();
        let execution_limits = prepared_formula_list_limits
            .take()
            .ok_or(Error::Verification { path })?;
        let artifact = match logical_list.execute(execution_limits) {
            Ok(artifact) => artifact,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_formula_list_error(error, path));
            },
        };
        let artifact_report = artifact.report;
        merge_usage(
            &mut execution_actual,
            formula_list_usage(artifact_report, path)?,
            path,
        )?;
        let binding_usage = formula_list_binding_usage(&target.storage.lists.formula, path)?;
        let mut bound =
            match bind_formula_list_edits(source, &target.storage.lists.formula, artifact, path) {
                Ok(bound) => bound,
                Err(error) => {
                    budget.cancel_authorization();
                    return Err(error);
                },
            };
        merge_usage(&mut execution_actual, binding_usage, path)?;
        released_plan_elements = released_plan_elements
            .checked_add(usize_u64(logical_report.retained_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(usize_u64(logical_report.retained_bytes))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_elements = released_artifact_elements
            .checked_add(usize_u64(bound.artifact_disposable_elements))
            .and_then(|count| count.checked_add(binding_usage.retained_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_bytes = released_artifact_bytes
            .checked_add(usize_u64(bound.artifact_disposable_bytes))
            .and_then(|count| count.checked_add(binding_usage.retained_bytes))
            .ok_or(Error::InvalidSource { path })?;
        prepared.append(&mut bound.replacements);
        if artifact_report.output_bytes > requirements.output_bytes
            || artifact_report.retained_elements > requirements.retained_elements
            || artifact_report.retained_bytes > requirements.retained_bytes
        {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        }
    } else if prepared_formula_list_limits.take().is_some() {
        budget.cancel_authorization();
        return Err(Error::Verification { path });
    }
    if let Some(prepared_metadata) = prepared_formula_metadata.take() {
        let prepare_report = prepared_metadata.prepare_report();
        let requirements = prepared_metadata.execution_requirements();
        let package_wire = source.state.options.archive().max_iwa_stream_bytes();
        let limits = formula_metadata::Limits {
            max_source_bytes: package_wire,
            max_output_bytes: requirements.output_bytes,
            max_fields: requirements.fields,
            max_work_bytes: requirements.work_bytes,
            max_references: requirements.references,
            max_messages: requirements.objects,
            max_hosts: requirements.hosts,
            max_precedents: requirements.precedents,
            max_ranges: requirements.ranges,
            max_retained_bytes: requirements.retained_bytes,
            max_scratch_bytes: requirements.peak_scratch_bytes,
            max_allocations: requirements.allocations,
            recursion_limit: 64,
        };
        let artifact = match prepared_metadata.execute(limits) {
            Ok(artifact) => artifact,
            Err(error) => {
                budget.cancel_authorization();
                return Err(map_formula_metadata_error(error, path));
            },
        };
        let artifact_report = artifact.report;
        merge_usage(
            &mut execution_actual,
            formula_metadata_artifact_usage(artifact_report, path)?,
            path,
        )?;
        let dependency_tiles = formula_dependency_tiles
            .as_ref()
            .ok_or(Error::Verification { path })?;
        let binding_usage =
            formula_metadata_binding_usage(source, &target, dependency_tiles, path)?;
        let mut bound =
            match bind_formula_metadata_edits(source, &target, dependency_tiles, artifact, path) {
                Ok(bound) => bound,
                Err(error) => {
                    budget.cancel_authorization();
                    return Err(error);
                },
            };
        merge_usage(&mut execution_actual, binding_usage, path)?;
        released_plan_elements = released_plan_elements
            .checked_add(usize_u64(prepare_report.retained_elements))
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(usize_u64(prepare_report.retained_bytes))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_elements = released_artifact_elements
            .checked_add(usize_u64(bound.artifact_disposable_elements))
            .and_then(|count| count.checked_add(usize_u64(bound.binding_disposable_elements)))
            .ok_or(Error::InvalidSource { path })?;
        released_artifact_bytes = released_artifact_bytes
            .checked_add(usize_u64(bound.artifact_disposable_bytes))
            .and_then(|count| count.checked_add(usize_u64(bound.binding_disposable_bytes)))
            .ok_or(Error::InvalidSource { path })?;
        prepared.append(&mut bound.replacements);
        formula_dependency_new_tiles = bound.new_tiles;
        if artifact_report.output_bytes > requirements.output_bytes
            || artifact_report.retained_elements > requirements.retained_elements
            || artifact_report.retained_bytes > requirements.retained_bytes
            || formula_dependency_new_tiles.len() != dependency_tiles.assignments.len()
        {
            budget.cancel_authorization();
            return Err(Error::Verification { path });
        }
    } else if formula_dependency_tiles
        .as_ref()
        .is_some_and(|tiles| !tiles.assignments.is_empty())
    {
        budget.cancel_authorization();
        return Err(Error::Verification { path });
    }
    if let Some(prepared_metadata) = prepared_package_metadata.take() {
        let requirements = prepared_metadata.execution_requirements();
        let output = match prepared_metadata.execute(requirements.exact_limits()) {
            Ok(output) => output,
            Err(error) => {
                budget.cancel_authorization();
                return Err(sparse_commit::map_metadata_error(error, path));
            },
        };
        merge_usage(
            &mut execution_actual,
            sparse_commit::package_metadata_execution_report_usage(output.report())?,
            path,
        )?;
        let route = formula_dependency_tiles
            .as_ref()
            .and_then(|tiles| tiles.inspection.as_ref())
            .map(|inspection| inspection.route)
            .ok_or(Error::Verification { path })?;
        prepared.push(PreparedReplacement {
            route,
            payload: output.into_bytes(),
            references: None,
        });
        released_plan_elements = released_plan_elements
            .checked_add(usize_u64(formula_dependency_uuid_additions.capacity()))
            .ok_or(Error::InvalidSource { path })?;
        released_plan_bytes = released_plan_bytes
            .checked_add(usize_u64(
                formula_dependency_uuid_additions
                    .capacity()
                    .checked_mul(size_of::<ObjectUuidAddition<'_>>())
                    .ok_or(Error::InvalidSource { path })?,
            ))
            .ok_or(Error::InvalidSource { path })?;
    } else if !formula_dependency_uuid_additions.is_empty() {
        budget.cancel_authorization();
        return Err(Error::Verification { path });
    }
    execution_actual.peak_scratch_bytes = aggregate_peak;
    budget.record_authorized(execution_actual)?;
    budget.release_retained(released_plan_elements, released_plan_bytes)?;
    budget.release_retained(released_artifact_elements, released_artifact_bytes)?;
    budget.release_retained(
        usize_u64(
            prepared_tile_elements
                .checked_add(prepared_header_elements)
                .ok_or(Error::InvalidSource { path })?,
        ),
        usize_u64(
            prepared_tile_bytes
                .checked_add(prepared_header_bytes)
                .ok_or(Error::InvalidSource { path })?,
        ),
    )?;

    prepared.sort_unstable_by_key(|replacement| {
        (
            replacement.route.component_index,
            replacement.route.object_index,
            replacement.route.message_index,
        )
    });
    if formula_dependency_tiles
        .as_ref()
        .is_some_and(|tiles| !tiles.assignments.is_empty())
    {
        let publication = sparse_commit::publish_formula_dependency_objects(
            source,
            prepared,
            formula_dependency_tiles.ok_or(Error::Verification { path })?,
            formula_dependency_new_tiles,
            future_publication,
            &mut budget,
            path,
        )?;
        let changed_cells = changes.len();
        for change in changes {
            let (_position, input) = change.into_parts();
            if let Some(input) = input {
                drop(input.into_scalar_value());
            }
        }
        let source_bytes = source.state.source.clone();
        let target_bytes = publication.outcome.package.state.source.clone();
        let patch = Patch::from_exact_with_evidence(
            path,
            requested,
            changed_cells,
            source_bytes,
            target_bytes,
            source.snapshot(),
            publication.outcome.package.snapshot(),
            publication.evidence,
        )?;
        return Ok(Commit::new(
            publication.outcome.package,
            patch,
            Diagnostics::from_changed(
                requested,
                changed_cells,
                publication.outcome.touched_components,
                cache_prepared
                    .refreshed_hosts
                    .checked_add(formula_refreshed_hosts)
                    .ok_or(Error::InvalidSource { path })?,
                publication.preview_count,
            ),
        ));
    }
    if let Some(dependency_tiles) = formula_dependency_tiles {
        budget.release_retained(
            usize_u64(dependency_tiles.retained_elements),
            usize_u64(dependency_tiles.retained_bytes),
        )?;
    }
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
            if !rewrite::component_admission_shape_fits(
                component_reservation,
                future_publication.component,
            ) {
                precharge_error.set(Some(Error::Verification { path }));
                return Err(rewrite::RewriteError::Precharge);
            }
            let envelope = match component_reservation_usage(future_publication.component, path) {
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
            if !rewrite::component_cost_fits(cost, future_publication.component) {
                precharge_error.set(Some(Error::Verification { path }));
                return Err(rewrite::RewriteError::Precharge);
            }
            if !rewrite::publication_reservation_fits(reservation, future_publication.publication) {
                precharge_error.set(Some(Error::Verification { path }));
                return Err(rewrite::RewriteError::Precharge);
            }
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
    if !formula_publication {
        verify_semantic_changes(
            &outcome.package,
            sheet_position,
            table_position,
            &changes,
            path,
        )?;
    }
    let changed_cells = changes.len();
    for change in changes {
        let (_position, input) = change.into_parts();
        if let Some(input) = input {
            drop(input.into_scalar_value());
        }
    }

    let source_bytes = source.state.source.clone();
    let target_bytes = outcome.package.state.source.clone();
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
            cache_prepared
                .refreshed_hosts
                .checked_add(formula_refreshed_hosts)
                .ok_or(Error::InvalidSource { path })?,
            previews.len(),
        ),
    ))
}

fn bind_formula_list_edits(
    source: &Package,
    routes: &resolve::ListRoute,
    artifact: formula_list::Artifact,
    path: crate::package::table_cells::Path,
) -> Result<BoundFormulaList, Error> {
    if artifact.segments.len() != routes.segments.len() {
        return Err(Error::Verification { path });
    }
    let capacity = routes
        .segments
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    let artifact_disposable_elements = artifact.report.retained_elements;
    let moved_payload_bytes = artifact
        .root
        .payload
        .as_ref()
        .map_or(0, Vec::capacity)
        .checked_add(artifact.segments.iter().try_fold(0usize, |sum, edit| {
            sum.checked_add(edit.payload.as_ref().map_or(0, Vec::capacity))
                .ok_or(Error::InvalidSource { path })
        })?)
        .ok_or(Error::InvalidSource { path })?;
    let artifact_disposable_bytes = artifact
        .report
        .retained_bytes
        .checked_sub(moved_payload_bytes)
        .ok_or(Error::InvalidSource { path })?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: capacity,
        })?;
    require_exact_capacity(&replacements, capacity, path)?;
    let root_id = message_object_identifier(source, routes.message, path)?;
    if artifact.root.object_id != root_id
        || artifact.root.object_references.as_slice()
            != source_message_references(source, routes.message, path)?
    {
        return Err(Error::Verification { path });
    }
    if let Some(payload) = artifact.root.payload {
        replacements.push(PreparedReplacement {
            route: routes.message,
            payload,
            references: None,
        });
    }
    for (edit, route) in artifact.segments.into_iter().zip(&routes.segments) {
        if edit.object_id != message_object_identifier(source, *route, path)?
            || edit.object_references.as_slice() != source_message_references(source, *route, path)?
        {
            return Err(Error::Verification { path });
        }
        if let Some(payload) = edit.payload {
            replacements.push(PreparedReplacement {
                route: *route,
                payload,
                references: None,
            });
        }
    }
    Ok(BoundFormulaList {
        replacements,
        artifact_disposable_elements,
        artifact_disposable_bytes,
    })
}

fn formula_list_binding_usage(
    routes: &resolve::ListRoute,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let capacity = routes
        .segments
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    let work = routes
        .segments
        .len()
        .checked_mul(2)
        .and_then(|work| work.checked_add(2))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: usize_u64(capacity),
        retained_bytes: usize_u64(
            capacity
                .checked_mul(size_of::<PreparedReplacement>())
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(capacity != 0),
        lookups: usize_u64(work),
        transaction_work: usize_u64(work),
        ..budget::Usage::default()
    })
}

fn bind_formula_metadata_edits(
    source: &Package,
    target: &resolve::Target,
    dependency_tiles: &PreparedFormulaDependencyTiles,
    artifact: formula_metadata::Artifact,
    path: crate::package::table_cells::Path,
) -> Result<BoundFormulaMetadata, Error> {
    let engine_route = target
        .dependencies
        .engine
        .ok_or(Error::Verification { path })?;
    let cell_route_count =
        target
            .dependencies
            .formula_owners
            .iter()
            .try_fold(0usize, |count, owner| {
                count
                    .checked_add(owner.cell_record_tiles.len())
                    .ok_or(Error::InvalidSource { path })
            })?;
    let range_route_count =
        target
            .dependencies
            .formula_owners
            .iter()
            .try_fold(0usize, |count, owner| {
                count
                    .checked_add(owner.range_precedent_tiles.len())
                    .ok_or(Error::InvalidSource { path })
            })?;
    let final_cell_tile_count = cell_route_count
        .checked_add(dependency_tiles.assignments.len())
        .ok_or(Error::InvalidSource { path })?;
    if artifact.owners.len() != target.dependencies.formula_owners.len()
        || artifact.cell_tiles.len() != final_cell_tile_count
        || artifact.range_tiles.len() != range_route_count
    {
        return Err(Error::Verification { path });
    }
    let capacity = formula_metadata_binding_capacity(target, dependency_tiles, path)?;
    let artifact_report = artifact.report;
    let persistent_payload_bytes = core::iter::once(&artifact.engine)
        .chain(&artifact.owners)
        .chain(&artifact.cell_tiles)
        .chain(&artifact.range_tiles)
        .try_fold(0usize, |sum, edit| {
            sum.checked_add(edit.payload.as_ref().map_or(0, Vec::capacity))
                .ok_or(Error::InvalidSource { path })
        })?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: capacity,
        })?;
    require_exact_capacity(&replacements, capacity, path)?;
    let mut new_tiles = Vec::new();
    new_tiles
        .try_reserve_exact(dependency_tiles.assignments.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedElements,
            amount: dependency_tiles.assignments.len(),
        })?;
    require_exact_capacity(&new_tiles, dependency_tiles.assignments.len(), path)?;
    let selected_owner = target
        .dependencies
        .selected_formula_owner
        .ok_or(Error::Verification { path })?;
    let selected_route = target
        .dependencies
        .formula_owners
        .iter()
        .find(|owner| owner.internal_owner_id == selected_owner.internal_owner_id)
        .ok_or(Error::Verification { path })?
        .message;
    let selected_before = source_message_all_references(source, selected_route, path)?;
    let selected_before_capacity = selected_before.len();
    let mut copied_before = Vec::new();
    copied_before
        .try_reserve_exact(selected_before_capacity)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::References,
            amount: selected_before_capacity,
        })?;
    require_exact_capacity(&copied_before, selected_before_capacity, path)?;
    copied_before.extend_from_slice(selected_before);
    bind_formula_existing_edit(
        source,
        &mut replacements,
        engine_route,
        artifact.engine,
        path,
    )?;
    let mut selected_after_reference_elements = 0usize;
    let mut selected_after_reference_bytes = 0usize;
    let mut copied_before_moved = false;
    for (owner, edit) in target
        .dependencies
        .formula_owners
        .iter()
        .zip(artifact.owners)
    {
        if owner.internal_owner_id == selected_owner.internal_owner_id {
            let before = selected_before;
            if edit.object_id != message_object_identifier(source, owner.message, path)?
                || edit.object_references.len()
                    != before
                        .len()
                        .checked_add(dependency_tiles.assignments.len())
                        .ok_or(Error::InvalidSource { path })?
                || !edit.object_references.starts_with(before)
                || edit.object_references[before.len()..]
                    .iter()
                    .copied()
                    .ne(dependency_tiles
                        .assignments
                        .iter()
                        .map(|tile| tile.object_id))
            {
                return Err(Error::Verification { path });
            }
            if let Some(payload) = edit.payload {
                selected_after_reference_elements = edit.object_references.capacity();
                selected_after_reference_bytes = selected_after_reference_elements
                    .checked_mul(size_of::<u64>())
                    .ok_or(Error::InvalidSource { path })?;
                copied_before_moved = true;
                replacements.push(PreparedReplacement {
                    route: owner.message,
                    payload,
                    references: Some(PreparedReferenceDelta {
                        before: core::mem::take(&mut copied_before),
                        after: edit.object_references,
                    }),
                });
            }
        } else {
            bind_formula_existing_edit(source, &mut replacements, owner.message, edit, path)?;
        }
    }
    let mut cell_edits = artifact.cell_tiles.into_iter();
    for owner in &target.dependencies.formula_owners {
        for &route in &owner.cell_record_tiles {
            bind_formula_existing_edit(
                source,
                &mut replacements,
                route,
                cell_edits.next().ok_or(Error::Verification { path })?,
                path,
            )?;
        }
        if owner.internal_owner_id == selected_owner.internal_owner_id {
            for assignment in &dependency_tiles.assignments {
                let edit = cell_edits.next().ok_or(Error::Verification { path })?;
                if edit.object_id != assignment.object_id
                    || !edit.object_references.is_empty()
                    || edit.payload.is_none()
                {
                    return Err(Error::Verification { path });
                }
                new_tiles.push(edit);
            }
        }
    }
    if cell_edits.next().is_some() {
        return Err(Error::Verification { path });
    }
    for (route, edit) in target
        .dependencies
        .formula_owners
        .iter()
        .flat_map(|owner| {
            owner
                .range_precedent_tiles
                .iter()
                .map(|route| route.message)
        })
        .zip(artifact.range_tiles)
    {
        bind_formula_existing_edit(source, &mut replacements, route, edit, path)?;
    }
    let new_tile_reference_elements = new_tiles.iter().try_fold(0usize, |sum, edit| {
        sum.checked_add(edit.object_references.capacity())
            .ok_or(Error::InvalidSource { path })
    })?;
    let persistent_reference_elements = selected_after_reference_elements
        .checked_add(new_tile_reference_elements)
        .ok_or(Error::InvalidSource { path })?;
    let persistent_reference_bytes = persistent_reference_elements
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource { path })?;
    if selected_after_reference_bytes > persistent_reference_bytes {
        return Err(Error::InvalidSource { path });
    }
    let artifact_disposable_elements = artifact_report
        .retained_elements
        .checked_sub(persistent_reference_elements)
        .ok_or(Error::InvalidSource { path })?;
    let artifact_disposable_bytes = artifact_report
        .retained_bytes
        .checked_sub(persistent_payload_bytes)
        .and_then(|bytes| bytes.checked_sub(persistent_reference_bytes))
        .ok_or(Error::InvalidSource { path })?;
    let replacement_bytes = replacements
        .capacity()
        .checked_mul(size_of::<PreparedReplacement>())
        .ok_or(Error::InvalidSource { path })?;
    let unused_before_elements = if copied_before_moved {
        0
    } else {
        copied_before.capacity()
    };
    let unused_before_bytes = unused_before_elements
        .checked_mul(size_of::<u64>())
        .ok_or(Error::InvalidSource { path })?;
    Ok(BoundFormulaMetadata {
        replacements,
        new_tiles,
        artifact_disposable_elements,
        artifact_disposable_bytes,
        binding_disposable_elements: capacity
            .checked_add(unused_before_elements)
            .ok_or(Error::InvalidSource { path })?,
        binding_disposable_bytes: replacement_bytes
            .checked_add(unused_before_bytes)
            .ok_or(Error::InvalidSource { path })?,
    })
}

fn formula_metadata_binding_capacity(
    target: &resolve::Target,
    dependency_tiles: &PreparedFormulaDependencyTiles,
    path: crate::package::table_cells::Path,
) -> Result<usize, Error> {
    let cell_routes =
        target
            .dependencies
            .formula_owners
            .iter()
            .try_fold(0usize, |count, owner| {
                count
                    .checked_add(owner.cell_record_tiles.len())
                    .ok_or(Error::InvalidSource { path })
            })?;
    let range_routes =
        target
            .dependencies
            .formula_owners
            .iter()
            .try_fold(0usize, |count, owner| {
                count
                    .checked_add(owner.range_precedent_tiles.len())
                    .ok_or(Error::InvalidSource { path })
            })?;
    1usize
        .checked_add(target.dependencies.formula_owners.len())
        .and_then(|count| count.checked_add(cell_routes))
        .and_then(|count| count.checked_add(dependency_tiles.assignments.len()))
        .and_then(|count| count.checked_add(range_routes))
        .ok_or(Error::InvalidSource { path })
}

fn formula_metadata_binding_usage(
    source: &Package,
    target: &resolve::Target,
    dependency_tiles: &PreparedFormulaDependencyTiles,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let replacement_capacity = formula_metadata_binding_capacity(target, dependency_tiles, path)?;
    let new_tile_capacity = dependency_tiles.assignments.len();
    let selected_owner = target
        .dependencies
        .selected_formula_owner
        .ok_or(Error::Verification { path })?;
    let selected_route = target
        .dependencies
        .formula_owners
        .iter()
        .find(|owner| owner.internal_owner_id == selected_owner.internal_owner_id)
        .ok_or(Error::Verification { path })?
        .message;
    let before_capacity = source_message_all_references(source, selected_route, path)?.len();
    let existing_edits = 1usize
        .checked_add(target.dependencies.formula_owners.len())
        .and_then(|count| {
            target
                .dependencies
                .formula_owners
                .iter()
                .try_fold(count, |count, owner| {
                    count
                        .checked_add(owner.cell_record_tiles.len())
                        .and_then(|count| count.checked_add(owner.range_precedent_tiles.len()))
                })
        })
        .ok_or(Error::InvalidSource { path })?;
    let work = existing_edits
        .checked_mul(2)
        .and_then(|work| work.checked_add(new_tile_capacity))
        .and_then(|work| work.checked_add(before_capacity))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: usize_u64(
            replacement_capacity
                .checked_add(new_tile_capacity)
                .and_then(|count| count.checked_add(before_capacity))
                .ok_or(Error::InvalidSource { path })?,
        ),
        retained_bytes: usize_u64(
            replacement_capacity
                .checked_mul(size_of::<PreparedReplacement>())
                .and_then(|bytes| {
                    new_tile_capacity
                        .checked_mul(size_of::<formula_metadata::MessageEdit>())
                        .and_then(|tiles| bytes.checked_add(tiles))
                })
                .and_then(|bytes| {
                    before_capacity
                        .checked_mul(size_of::<u64>())
                        .and_then(|references| bytes.checked_add(references))
                })
                .ok_or(Error::InvalidSource { path })?,
        ),
        allocation_events: u64::from(replacement_capacity != 0)
            .checked_add(u64::from(new_tile_capacity != 0))
            .and_then(|count| count.checked_add(u64::from(before_capacity != 0)))
            .ok_or(Error::InvalidSource { path })?,
        lookups: usize_u64(existing_edits),
        transaction_work: usize_u64(work),
        ..budget::Usage::default()
    })
}

fn bind_formula_existing_edit(
    source: &Package,
    replacements: &mut Vec<PreparedReplacement>,
    route: resolve::MessageRoute,
    edit: formula_metadata::MessageEdit,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    if edit.object_id != message_object_identifier(source, route, path)?
        || edit.object_references.as_slice() != source_message_all_references(source, route, path)?
    {
        return Err(Error::Verification { path });
    }
    if let Some(payload) = edit.payload {
        replacements.push(PreparedReplacement {
            route,
            payload,
            references: None,
        });
    }
    Ok(())
}

fn authorize_remaining(budget: &mut budget::TransactionBudget) -> Result<budget::Remaining, Error> {
    let remaining = budget.remaining()?;
    budget.authorize(remaining)?;
    Ok(remaining)
}
fn ensure_formula_owner_context_coverage(
    selected: &resolve::Target,
    context: &PreparedFormulaContext,
    budget: &mut budget::TransactionBudget,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    let work = selected
        .dependencies
        .formula_owners
        .len()
        .checked_add(context.external_targets.len())
        .ok_or(Error::InvalidSource { path })?;
    let usage = budget::Usage {
        lookups: usize_u64(work),
        transaction_work: usize_u64(work),
        ..budget::Usage::default()
    };
    budget.authorize(usage)?;
    let selected_engine = selected.dependencies.engine;
    let external_valid = context
        .external_targets
        .iter()
        .all(|(_, target)| target.dependencies.engine == selected_engine);
    let table_owner_count = selected
        .dependencies
        .formula_owners
        .iter()
        .filter_map(|owner| owner.formula_owner_object_id)
        .count();
    let owners_covered = table_owner_count
        == context
            .external_targets
            .len()
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
    budget.record_authorized(usage)?;
    if !external_valid || !owners_covered {
        return Err(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::Formula,
        });
    }
    Ok(())
}

fn extend_formula_referenced_paths(
    referenced_paths: &mut Vec<formula_author::ReferencedTablePath>,
    external_targets: &[(formula_author::ReferencedTablePath, resolve::Target)],
    selected_owner: u32,
    logical: formula_metadata::LogicalGraph<'_>,
    requirements: formula_metadata::ExecutionRequirements,
    budget: &mut budget::TransactionBudget,
    path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    let fact_upper = requirements
        .precedents
        .checked_add(requirements.ranges)
        .ok_or(Error::InvalidSource { path })?;
    let scan_envelope = budget::Usage {
        transaction_work: usize_u64(
            requirements
                .hosts
                .checked_add(fact_upper)
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    budget.authorize(scan_envelope)?;
    let fact_counts = logical
        .hosts
        .iter()
        .try_fold((0usize, 0usize), |counts, host| {
            host.precedents
                .iter()
                .map(|precedent| precedent.target_owner)
                .chain(host.ranges.iter().map(|range| range.target_owner))
                .try_fold(counts, |(cross, total), owner| {
                    Ok((
                        cross
                            .checked_add(usize::from(owner != selected_owner))
                            .ok_or(Error::InvalidSource { path })?,
                        total.checked_add(1).ok_or(Error::InvalidSource { path })?,
                    ))
                })
        });
    let (cross_occurrences, actual_facts) = match fact_counts {
        Ok(counts) => {
            let scan_usage = budget::Usage {
                transaction_work: usize_u64(
                    logical
                        .hosts
                        .len()
                        .checked_add(counts.1)
                        .ok_or(Error::InvalidSource { path })?,
                ),
                ..budget::Usage::default()
            };
            if let Err(error) = budget.record_authorized(scan_usage) {
                budget.cancel_authorization();
                return Err(error);
            }
            counts
        },
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    let final_capacity = referenced_paths
        .len()
        .checked_add(cross_occurrences)
        .ok_or(Error::InvalidSource { path })?;
    let owner_index_bytes = external_targets
        .len()
        .checked_mul(size_of::<(u32, formula_author::ReferencedTablePath)>())
        .ok_or(Error::InvalidSource { path })?;
    let retained_bytes = final_capacity
        .checked_mul(size_of::<formula_author::ReferencedTablePath>())
        .ok_or(Error::InvalidSource { path })?;
    let build_work = external_targets
        .len()
        .checked_add(sort_work(external_targets.len(), path)?)
        .and_then(|work| work.checked_add(logical.hosts.len()))
        .and_then(|work| work.checked_add(actual_facts))
        .and_then(|work| {
            cross_occurrences
                .checked_mul(binary_search_work(external_targets.len()).checked_add(1)?)
                .and_then(|search| work.checked_add(search))
        })
        .and_then(|work| sort_work(final_capacity, path).ok()?.checked_add(work))
        .ok_or(Error::InvalidSource { path })?;
    let envelope = budget::Usage {
        retained_elements: usize_u64(final_capacity),
        retained_bytes: usize_u64(retained_bytes),
        peak_scratch_bytes: usize_u64(owner_index_bytes),
        allocation_events: u64::from(final_capacity != 0)
            .checked_add(u64::from(!external_targets.is_empty()))
            .ok_or(Error::InvalidSource { path })?,
        lookups: usize_u64(cross_occurrences),
        transaction_work: usize_u64(build_work),
        ..budget::Usage::default()
    };
    budget.authorize(envelope)?;
    let result = (|| {
        let mut owner_index = Vec::new();
        owner_index
            .try_reserve_exact(external_targets.len())
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
                amount: owner_index_bytes,
            })?;
        require_exact_capacity(&owner_index, external_targets.len(), path)?;
        for (referenced, target) in external_targets {
            let owner = target
                .dependencies
                .selected_formula_owner
                .ok_or(Error::InvalidSource { path })?;
            owner_index.push((owner.internal_owner_id, *referenced));
        }
        owner_index.sort_unstable_by_key(|entry| entry.0);
        if owner_index.windows(2).any(|pair| pair[0].0 >= pair[1].0) {
            return Err(Error::InvalidSource { path });
        }
        let mut final_paths = Vec::new();
        final_paths
            .try_reserve_exact(final_capacity)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: final_capacity,
            })?;
        require_exact_capacity(&final_paths, final_capacity, path)?;
        final_paths.extend_from_slice(referenced_paths);
        for host in logical.hosts {
            for target_owner in host
                .precedents
                .iter()
                .map(|precedent| precedent.target_owner)
                .chain(host.ranges.iter().map(|range| range.target_owner))
            {
                if target_owner == selected_owner {
                    continue;
                }
                let referenced = owner_index
                    .binary_search_by_key(&target_owner, |entry| entry.0)
                    .ok()
                    .and_then(|index| owner_index.get(index))
                    .map(|entry| entry.1)
                    .ok_or(Error::UnsupportedDependency {
                        path,
                        kind: crate::package::table_cells::DependencyKind::FormulaCache,
                    })?;
                final_paths.push(referenced);
            }
        }
        final_paths.sort_unstable();
        final_paths.dedup();
        Ok(final_paths)
    })();
    let final_paths = match result {
        Ok(paths) => {
            if let Err(error) = budget.record_authorized(envelope) {
                budget.cancel_authorization();
                return Err(error);
            }
            paths
        },
        Err(error) => {
            budget.cancel_authorization();
            return Err(error);
        },
    };
    let previous = core::mem::replace(referenced_paths, final_paths);
    let previous_bytes = previous
        .capacity()
        .checked_mul(size_of::<formula_author::ReferencedTablePath>())
        .ok_or(Error::InvalidSource { path })?;
    budget.release_retained(usize_u64(previous.capacity()), usize_u64(previous_bytes))?;
    Ok(())
}

fn prepare_formula_cache_registry<'source>(
    source: &'source Package,
    selected: &resolve::Target,
    context: &PreparedFormulaContext,
    budget: &mut budget::TransactionBudget,
    path: crate::package::table_cells::Path,
) -> Result<
    (
        Vec<cache::TableGeometry>,
        Vec<cache_commit::ExternalBaselineTable<'source>>,
    ),
    Error,
> {
    let selected_owner =
        selected
            .dependencies
            .selected_formula_owner
            .ok_or(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::FormulaCache,
            })?;
    let table_count = context
        .external_targets
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource { path })?;
    let planning_work =
        context
            .external_targets
            .iter()
            .try_fold(0usize, |work, (referenced, _)| {
                let path_work = usize::try_from(referenced.table)
                    .ok()
                    .and_then(|table| table.checked_add(1))
                    .and_then(|table| table.checked_mul(3))
                    .and_then(|table| {
                        binary_search_work(context.referenced_paths.len())
                            .checked_mul(3)
                            .and_then(|search| table.checked_add(search))
                    })
                    .ok_or(Error::InvalidSource { path })?;
                work.checked_add(path_work)
                    .ok_or(Error::InvalidSource { path })
            })?;
    let planning_usage = budget::Usage {
        lookups: usize_u64(context.external_targets.len()),
        transaction_work: usize_u64(planning_work),
        ..budget::Usage::default()
    };
    budget.reserve(planning_usage)?;
    let baseline_count = context
        .external_targets
        .iter()
        .filter(|(referenced, _)| context.referenced_paths.binary_search(referenced).is_ok())
        .count();
    let retained_elements = table_count
        .checked_add(baseline_count)
        .ok_or(Error::InvalidSource { path })?;
    let retained_bytes = table_count
        .checked_mul(size_of::<cache::TableGeometry>())
        .and_then(|bytes| {
            baseline_count
                .checked_mul(size_of::<cache_commit::ExternalBaselineTable<'_>>())
                .and_then(|baseline| bytes.checked_add(baseline))
        })
        .ok_or(Error::InvalidSource { path })?;
    let table_sort_work = sort_work(table_count, path)?;
    let baseline_sort_work = sort_work(baseline_count, path)?;
    let usage = budget::Usage {
        retained_elements: usize_u64(retained_elements),
        retained_bytes: usize_u64(retained_bytes),
        allocation_events: u64::from(table_count != 0) + u64::from(baseline_count != 0),
        lookups: usize_u64(baseline_count),
        transaction_work: usize_u64(
            table_sort_work
                .checked_add(baseline_sort_work)
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    budget.authorize(usage)?;
    let result = (|| {
        let mut tables: Vec<cache::TableGeometry> = Vec::new();
        let mut baseline: Vec<cache_commit::ExternalBaselineTable<'source>> = Vec::new();
        tables
            .try_reserve_exact(table_count)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: table_count,
            })?;
        baseline
            .try_reserve_exact(baseline_count)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::RetainedElements,
                amount: baseline_count,
            })?;
        require_exact_capacity(&tables, table_count, path)?;
        require_exact_capacity(&baseline, baseline_count, path)?;
        tables.push(cache::TableGeometry {
            identity: cache::TableIdentity {
                owner: selected_owner.internal_owner_id,
                uuid_lower: selected_owner.uid_lower,
                uuid_upper: selected_owner.uid_upper,
            },
            rows: selected.native.rows,
            columns: selected.native.columns,
            header_rows: u32::try_from(selected.native.settings.header_row_count())
                .map_err(|_| Error::InvalidSource { path })?,
            header_columns: u32::try_from(selected.native.settings.header_column_count())
                .map_err(|_| Error::InvalidSource { path })?,
            footer_rows: u32::try_from(selected.native.settings.footer_row_count())
                .map_err(|_| Error::InvalidSource { path })?,
        });
        for (referenced, target) in &context.external_targets {
            let owner =
                target
                    .dependencies
                    .selected_formula_owner
                    .ok_or(Error::UnsupportedDependency {
                        path,
                        kind: crate::package::table_cells::DependencyKind::FormulaCache,
                    })?;
            let table = source
                .sheets()
                .get(usize::try_from(referenced.sheet).map_err(|_| Error::InvalidSource { path })?)
                .and_then(|sheet| sheet.tables().nth(usize::try_from(referenced.table).ok()?))
                .ok_or(Error::InvalidSource { path })?;
            if table.row_count() != target.native.rows
                || table.column_count() != target.native.columns
            {
                return Err(Error::InvalidSource { path });
            }
            tables.push(cache::TableGeometry {
                identity: cache::TableIdentity {
                    owner: owner.internal_owner_id,
                    uuid_lower: owner.uid_lower,
                    uuid_upper: owner.uid_upper,
                },
                rows: target.native.rows,
                columns: target.native.columns,
                header_rows: u32::try_from(target.native.settings.header_row_count())
                    .map_err(|_| Error::InvalidSource { path })?,
                header_columns: u32::try_from(target.native.settings.header_column_count())
                    .map_err(|_| Error::InvalidSource { path })?,
                footer_rows: u32::try_from(target.native.settings.footer_row_count())
                    .map_err(|_| Error::InvalidSource { path })?,
            });
            if context.referenced_paths.binary_search(referenced).is_err() {
                continue;
            }
            baseline.push(cache_commit::ExternalBaselineTable {
                owner: owner.internal_owner_id,
                table,
            });
        }
        tables.sort_unstable_by_key(|table| table.identity.owner);
        baseline.sort_unstable_by_key(|entry| entry.owner);
        if tables
            .windows(2)
            .any(|pair| pair[0].identity.owner >= pair[1].identity.owner)
            || baseline
                .windows(2)
                .any(|pair| pair[0].owner >= pair[1].owner)
        {
            return Err(Error::InvalidSource { path });
        }
        Ok((tables, baseline))
    })();
    match result {
        Ok(result) => {
            budget.record_authorized(usage)?;
            Ok(result)
        },
        Err(error) => {
            budget.cancel_authorization();
            Err(error)
        },
    }
}

fn prepare_narrow_formula_metadata<'source>(
    source: &'source Package,
    target: &resolve::Target,
    context: &PreparedFormulaContext,
    existing: &cache_commit::ExistingFormulaIndex<'source>,
    list_deltas: &[formula_list::HostDelta<'_>],
    dependency_tiles: &PreparedFormulaDependencyTiles,
    budget: &mut budget::TransactionBudget,
    path: crate::package::table_cells::Path,
) -> Result<formula_metadata::PreparedGraph<'source>, Error> {
    let selected =
        target
            .dependencies
            .selected_formula_owner
            .ok_or(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::Formula,
            })?;
    let owner = target
        .dependencies
        .formula_owners
        .iter()
        .find(|candidate| {
            candidate.formula_owner_object_id == Some(target.native.drawable_identifier)
        })
        .ok_or(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::Formula,
        })?;
    let authored = context.authored.formulas.as_slice();
    for formula in authored {
        let coordinate = (formula.position.row(), formula.position.column());
        let delta = list_deltas
            .binary_search_by_key(&coordinate, |delta| (delta.row, delta.column))
            .ok()
            .and_then(|index| list_deltas.get(index))
            .ok_or(Error::Verification { path })?;
        if delta.new_formula != Some(formula.bytes.as_slice()) {
            return Err(Error::Verification { path });
        }
    }
    let engine_route = target
        .dependencies
        .engine
        .ok_or(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::Formula,
        })?;
    if selected.message != owner.message
        || selected.internal_owner_id != owner.internal_owner_id
        || selected.uid_lower != owner.uid_lower
        || selected.uid_upper != owner.uid_upper
        || owner.formula_owner_object_id != Some(target.native.drawable_identifier)
        || !target.dependencies.range_precedent_tiles.is_empty()
        || !owner.range_precedent_tiles.is_empty()
        || !existing.formula_list.segments.is_empty()
    {
        return Err(Error::UnsupportedDependency {
            path,
            kind: crate::package::table_cells::DependencyKind::Formula,
        });
    }
    let existing_host_count = existing.formula_list.hosts.len();
    let engine = metadata_source_message(source, engine_route, path)?;
    let owner_count = target.dependencies.formula_owners.len();
    let new_cell_count = dependency_tiles.assignments.len();
    let cell_count = target
        .dependencies
        .formula_owners
        .iter()
        .try_fold(0usize, |count, owner| {
            count
                .checked_add(owner.cell_record_tiles.len())
                .ok_or(Error::InvalidSource { path })
        })?
        .checked_add(new_cell_count)
        .ok_or(Error::InvalidSource { path })?;
    let range_count =
        target
            .dependencies
            .formula_owners
            .iter()
            .try_fold(0usize, |count, owner| {
                count
                    .checked_add(owner.range_precedent_tiles.len())
                    .ok_or(Error::InvalidSource { path })
            })?;
    let temporary_bytes = owner_count
        .checked_mul(
            size_of::<Vec<formula_metadata::CellTileSource<'_>>>()
                .checked_add(size_of::<Vec<formula_metadata::RangeTileSource<'_>>>())
                .and_then(|bytes| {
                    bytes.checked_add(size_of::<formula_metadata::SourceOwner<'_, '_>>())
                })
                .ok_or(Error::InvalidSource { path })?,
        )
        .and_then(|bytes| {
            cell_count
                .checked_mul(size_of::<formula_metadata::CellTileSource<'_>>())
                .and_then(|cells| bytes.checked_add(cells))
        })
        .and_then(|bytes| {
            range_count
                .checked_mul(size_of::<formula_metadata::RangeTileSource<'_>>())
                .and_then(|ranges| bytes.checked_add(ranges))
        })
        .ok_or(Error::InvalidSource { path })?;
    let temporary_allocations = owner_count
        .checked_mul(2)
        .and_then(|events| events.checked_add(3))
        .ok_or(Error::InvalidSource { path })?;
    budget.reserve_scratch(usize_u64(temporary_bytes), usize_u64(temporary_allocations))?;
    let mut cell_sets = Vec::new();
    let mut range_sets = Vec::new();
    cell_sets
        .try_reserve_exact(owner_count)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: owner_count,
        })?;
    range_sets
        .try_reserve_exact(owner_count)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: owner_count,
        })?;
    require_exact_capacity(&cell_sets, owner_count, path)?;
    require_exact_capacity(&range_sets, owner_count, path)?;
    for source_owner in &target.dependencies.formula_owners {
        let mut cells = Vec::new();
        let is_selected = source_owner.internal_owner_id == selected.internal_owner_id;
        let new_count = if is_selected { new_cell_count } else { 0 };
        let capacity = source_owner
            .cell_record_tiles
            .len()
            .checked_add(new_count)
            .ok_or(Error::InvalidSource { path })?;
        cells
            .try_reserve_exact(capacity)
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
                amount: capacity,
            })?;
        require_exact_capacity(&cells, capacity, path)?;
        for route in &source_owner.cell_record_tiles {
            let object_id = message_object_identifier(source, *route, path)?;
            let origin = is_selected
                .then(|| {
                    dependency_tiles
                        .existing_by_object
                        .binary_search_by_key(&object_id, |tile| tile.object_id)
                        .ok()
                        .map(|index| dependency_tiles.existing_by_object[index])
                })
                .flatten();
            cells.push(formula_metadata::CellTileSource {
                message: metadata_source_message(source, *route, path)?,
                source_present: true,
                tile_column_begin: origin.map_or(0, |tile| tile.column_begin),
                tile_row_begin: origin.map_or(0, |tile| tile.row_begin),
            });
        }
        if is_selected {
            for tile in &dependency_tiles.assignments {
                cells.push(formula_metadata::CellTileSource {
                    message: formula_metadata::SourceMessage {
                        object_id: tile.object_id,
                        payload: &[],
                        object_references: &[],
                    },
                    source_present: false,
                    tile_column_begin: tile.column_begin,
                    tile_row_begin: tile.row_begin,
                });
            }
        }
        cell_sets.push(cells);
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(source_owner.range_precedent_tiles.len())
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
                amount: source_owner.range_precedent_tiles.len(),
            })?;
        require_exact_capacity(&ranges, source_owner.range_precedent_tiles.len(), path)?;
        for route in &source_owner.range_precedent_tiles {
            ranges.push(formula_metadata::RangeTileSource {
                message: metadata_source_message(source, route.message, path)?,
                target_owner: route.target_owner,
            });
        }
        range_sets.push(ranges);
    }
    let mut owners = Vec::new();
    owners
        .try_reserve_exact(owner_count)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: owner_count,
        })?;
    require_exact_capacity(&owners, owner_count, path)?;
    for (index, source_owner) in target.dependencies.formula_owners.iter().enumerate() {
        owners.push(formula_metadata::SourceOwner {
            message: metadata_source_message(source, source_owner.message, path)?,
            internal_owner: source_owner.internal_owner_id,
            uid_lower: source_owner.uid_lower,
            uid_upper: source_owner.uid_upper,
            cell_tiles: &cell_sets[index],
            range_tiles: &range_sets[index],
        });
    }
    let fact_count = authored.iter().try_fold(0usize, |count, formula| {
        count
            .checked_add(formula.precedents.len())
            .and_then(|count| count.checked_add(formula.ranges.len()))
            .ok_or(Error::InvalidSource { path })
    })?;
    let fact_bytes = fact_count
        .checked_mul(size_of::<formula_metadata::Precedent>())
        .and_then(|bytes| {
            authored
                .len()
                .checked_mul(size_of::<Vec<formula_metadata::Precedent>>())
                .and_then(|sets| bytes.checked_add(sets))
        })
        .and_then(|bytes| {
            list_deltas
                .len()
                .checked_mul(size_of::<formula_metadata::HostChange<'_>>())
                .and_then(|changes| bytes.checked_add(changes))
        })
        .and_then(|bytes| {
            existing_host_count
                .checked_mul(size_of::<formula_metadata::HostKey>())
                .and_then(|hosts| bytes.checked_add(hosts))
        })
        .ok_or(Error::InvalidSource { path })?;
    let fact_allocations = authored
        .len()
        .checked_add(3)
        .ok_or(Error::InvalidSource { path })?;
    budget.reserve_scratch(usize_u64(fact_bytes), usize_u64(fact_allocations))?;
    let mut precedent_sets = Vec::new();
    let mut authored_range_sets = Vec::new();
    precedent_sets
        .try_reserve_exact(authored.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: authored.len(),
        })?;
    authored_range_sets
        .try_reserve_exact(authored.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: authored.len(),
        })?;
    for formula in authored {
        let mut facts = Vec::new();
        facts
            .try_reserve_exact(formula.precedents.len())
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
                amount: formula.precedents.len(),
            })?;
        for precedent in &formula.precedents {
            facts.push(formula_metadata::Precedent {
                target_owner: precedent.internal_owner(),
                row: precedent.row(),
                column: precedent.column(),
            });
        }
        facts.sort_unstable();
        facts.dedup();
        precedent_sets.push(facts);
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(formula.ranges.len())
            .map_err(|_| Error::Allocation {
                kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
                amount: formula.ranges.len(),
            })?;
        for range in &formula.ranges {
            ranges.push(formula_metadata::Range {
                target_owner: range.internal_owner(),
                top: range.top(),
                left: range.left(),
                bottom: range.bottom(),
                right: range.right(),
            });
        }
        ranges.sort_unstable();
        ranges.dedup();
        authored_range_sets.push(ranges);
    }
    let mut existing_hosts = Vec::new();
    existing_hosts
        .try_reserve_exact(existing_host_count)
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: existing_host_count,
        })?;
    for host in &existing.formula_list.hosts {
        existing_hosts.push(formula_metadata::HostKey {
            owner: selected.internal_owner_id,
            row: host.row,
            column: host.column,
        });
    }
    let mut changes = Vec::new();
    changes
        .try_reserve_exact(list_deltas.len())
        .map_err(|_| Error::Allocation {
            kind: crate::package::table_cells::LimitKind::PeakScratchBytes,
            amount: list_deltas.len(),
        })?;
    for delta in list_deltas {
        let old = delta.old_formula_key.map(|_| formula_metadata::HostKey {
            owner: selected.internal_owner_id,
            row: delta.row,
            column: delta.column,
        });
        let authored_index = authored
            .binary_search_by_key(&(delta.row, delta.column), |formula| {
                (formula.position.row(), formula.position.column())
            })
            .ok();
        let new = authored_index.map(|index| {
            let formula = &authored[index];
            formula_metadata::FormulaHost::authored(
                selected.internal_owner_id,
                formula.position.row(),
                formula.position.column(),
                &precedent_sets[index],
                &authored_range_sets[index],
            )
        });
        if new.is_some() != delta.new_formula.is_some() {
            return Err(Error::Verification { path });
        }
        let cell_tile_object_id = new.as_ref().and_then(|_| {
            let origin = (delta.row / 128 * 128, delta.column / 32 * 32);
            dependency_tiles
                .existing
                .binary_search_by_key(&origin, |tile| (tile.row_begin, tile.column_begin))
                .ok()
                .map(|index| dependency_tiles.existing[index].object_id)
                .or_else(|| {
                    dependency_tiles
                        .assignments
                        .binary_search_by_key(&origin, |tile| (tile.row_begin, tile.column_begin))
                        .ok()
                        .map(|index| dependency_tiles.assignments[index].object_id)
                })
        });
        changes.push(formula_metadata::HostChange {
            old,
            new,
            cell_tile_object_id,
        });
    }
    let graph = formula_metadata::SourceGraph {
        engine,
        owners: &owners,
        table_owner: selected.internal_owner_id,
        existing_formula_hosts: &existing_hosts,
        changes: &changes,
    };
    let remaining = budget.remaining()?;
    let mut limits = formula_metadata_limits(source, remaining, path)?;
    let metadata_fixed_work = owner_count
        .checked_add(cell_count)
        .and_then(|work| work.checked_add(range_count))
        .and_then(|work| work.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    limits.max_work_bytes = limits
        .max_work_bytes
        .checked_add(metadata_fixed_work)
        .ok_or(Error::InvalidSource { path })?;
    limits.max_messages = owner_count
        .checked_add(cell_count)
        .and_then(|count| count.checked_add(range_count))
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    // Source decoding and final overlay coexist during prepare.  A removal
    // lowers the final tracker count but still has to admit the source host.
    let inserted_hosts = changes
        .iter()
        .filter(|change| change.old.is_none() && change.new.is_some())
        .count();
    let source_and_final_host_upper = usize::try_from(target.dependencies.formula_count)
        .map_err(|_| Error::InvalidSource { path })?
        .checked_mul(2)
        .and_then(|count| count.checked_add(inserted_hosts))
        .ok_or(Error::InvalidSource { path })?;
    let source_fact_upper = owners.iter().try_fold(0usize, |sum, owner| {
        let owner_upper = owner
            .message
            .payload
            .len()
            .checked_div(2)
            .and_then(|count| count.checked_add(1))
            .ok_or(Error::InvalidSource { path })?;
        let cell_upper = owner.cell_tiles.iter().try_fold(0usize, |count, tile| {
            count
                .checked_add(tile.message.payload.len() / 2 + 1)
                .ok_or(Error::InvalidSource { path })
        })?;
        let range_upper = owner.range_tiles.iter().try_fold(0usize, |count, tile| {
            count
                .checked_add(tile.message.payload.len() / 2 + 1)
                .ok_or(Error::InvalidSource { path })
        })?;
        sum.checked_add(owner_upper)
            .and_then(|count| count.checked_add(cell_upper))
            .and_then(|count| count.checked_add(range_upper))
            .ok_or(Error::InvalidSource { path })
    })?;
    let precedent_upper = source_fact_upper
        .checked_add(fact_count)
        .ok_or(Error::InvalidSource { path })?;
    if source_and_final_host_upper > limits.max_hosts
        || precedent_upper > limits.max_precedents
        || source_fact_upper > limits.max_ranges
    {
        return Err(Error::LimitExceeded {
            path,
            kind: crate::package::table_cells::LimitKind::FormulaWork,
            observed: usize_u64(
                source_and_final_host_upper
                    .max(precedent_upper)
                    .max(source_fact_upper),
            ),
            maximum: usize_u64(
                limits
                    .max_hosts
                    .max(limits.max_precedents)
                    .max(limits.max_ranges),
            ),
        });
    }
    limits.max_hosts = source_and_final_host_upper;
    limits.max_precedents = precedent_upper;
    limits.max_ranges = source_fact_upper;
    let archive_reference_upper = engine
        .object_references
        .len()
        .checked_add(owners.iter().try_fold(0usize, |count, owner| {
            count
                .checked_add(owner.message.object_references.len())
                .ok_or(Error::InvalidSource { path })
        })?)
        // Source decode, source locality, candidate preflight, and strict
        // candidate reopen each visit the exact ArchiveInfo references; the
        // dependency codec also reports selected nested references.
        .and_then(|count| count.checked_mul(8))
        .ok_or(Error::InvalidSource { path })?;
    let rewrite_reference_upper = limits
        .max_messages
        .checked_mul(4)
        // A new tiled topology is visited during owner candidate construction,
        // tile encoding, ArchiveInfo validation, strict reopen, and the
        // emitted-source mirror proof. Those references are absent from the
        // source-derived message sum above, so admit them explicitly before
        // the leaf allocates any candidate output.
        .and_then(|count| {
            new_cell_count
                .checked_mul(32)
                .and_then(|new_references| count.checked_add(new_references))
        })
        .ok_or(Error::InvalidSource { path })?;
    limits.max_references = limits
        .max_references
        .min(archive_reference_upper.max(rewrite_reference_upper));
    budget.authorize(remaining)?;
    let prepared = match formula_metadata::prepare_graph(graph, limits) {
        Ok(prepared) => prepared,
        Err(error) => {
            budget.cancel_authorization();
            budget.release_scratch(usize_u64(fact_bytes));
            budget.release_scratch(usize_u64(temporary_bytes));
            return Err(map_formula_metadata_error(error, path));
        },
    };
    let report = prepared.prepare_report();
    let usage = formula_metadata_prepare_usage(report, path)?;
    budget.record_authorized(usage)?;
    budget.release_scratch(usize_u64(fact_bytes));
    budget.release_scratch(usize_u64(temporary_bytes));
    let expected_formula_count =
        final_formula_table_count(existing.formula_list.hosts.len(), list_deltas, path)?;
    if prepared.logical_view().formula_count != expected_formula_count {
        return Err(Error::Verification { path });
    }
    Ok(prepared)
}

fn final_formula_table_count(
    source_count: usize,
    deltas: &[formula_list::HostDelta<'_>],
    path: crate::package::table_cells::Path,
) -> Result<u64, Error> {
    let final_count = deltas.iter().try_fold(source_count, |count, delta| {
        match (delta.old_formula_key.is_some(), delta.new_formula.is_some()) {
            (false, true) => count.checked_add(1),
            (true, false) => count.checked_sub(1),
            (false, false) | (true, true) => Some(count),
        }
        .ok_or(Error::InvalidSource { path })
    })?;
    u64::try_from(final_count).map_err(|_| Error::InvalidSource { path })
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
    add!(updates);
    add!(input_value_bytes);
    add!(retained_elements);
    add!(retained_bytes);
    add!(scratch_bytes);
    add!(allocation_events);
    add!(wire_bytes);
    add!(wire_fields);
    add!(wire_work);
    add!(objects);
    add!(references);
    add!(lookups);
    add!(tile_reads);
    add!(tile_writes);
    add!(header_reads);
    add!(header_writes);
    add!(row_reads);
    add!(row_writes);
    add!(list_reads);
    add!(list_writes);
    add!(string_work);
    add!(rich_text_work);
    add!(formula_graph_builds);
    add!(formula_nodes);
    add!(formula_edges);
    add!(range_candidates);
    add!(cache_hosts);
    add!(authored_formula_writes);
    add!(formula_work);
    add!(component_encodes);
    add!(components_reassembled);
    add!(reassembly_bytes);
    add!(preview_bytes_deleted);
    add!(output_artifact_allocations);
    add!(output_bytes);
    add!(candidate_reopens);
    add!(reopen_references);
    add!(reopen_work);
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
        formula_nodes: usize_u64(staging_usage.formula_nodes()),
        formula_work: usize_u64(staging_usage.formula_nodes()),
        transaction_work: usize_u64(staging_usage.formula_nodes()),
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
    tile_reads: u64,
) -> Result<budget::Usage, Error> {
    Ok(budget::Usage {
        wire_bytes: report.wire_bytes,
        wire_fields: report.wire_fields,
        wire_work: report.wire_work,
        tile_reads,
        tile_writes: u64::from(report.output_bytes != 0),
        row_reads: report.rows_read,
        row_writes: report.rows_written,
        retained_elements: report.retained_elements,
        retained_bytes: report.retained_bytes,
        peak_scratch_bytes: report.peak_scratch_bytes,
        allocation_events: report.allocation_events,
        output_bytes: report.output_bytes,
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

fn tile_execution_requirements_usage(
    budget: &budget::TransactionBudget,
    requirements: tile::TileExecutionRequirements,
) -> Result<budget::Usage, Error> {
    tile_usage(
        budget,
        tile::TileReport {
            wire_bytes: usize_u64(requirements.input_bytes()),
            wire_fields: usize_u64(requirements.fields()),
            wire_work: requirements.work(),
            rows_read: usize_u64(requirements.rows_read()),
            rows_written: usize_u64(requirements.rows_written()),
            cell_slots_scanned: usize_u64(requirements.cell_slots_scanned()),
            cell_slots_written: usize_u64(requirements.cell_slots_written()),
            cache_cells_read: usize_u64(requirements.cache_cells_read()),
            cache_cells_written: usize_u64(requirements.cache_cells_written()),
            output_bytes: usize_u64(requirements.output_bytes()),
            retained_elements: usize_u64(requirements.retained_elements()),
            retained_bytes: usize_u64(requirements.retained_bytes()),
            current_scratch_bytes: 0,
            peak_scratch_bytes: usize_u64(requirements.peak_scratch_bytes()),
            allocation_events: usize_u64(requirements.allocations()),
        },
        0,
    )
}

fn header_execution_requirements_usage(
    requirements: sparse::HeaderBucketExecutionRequirements,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    sparse_commit::sparse_usage(
        sparse::SparseReport {
            input_bytes: requirements.input_bytes,
            output_bytes: requirements.output_bytes,
            fields: requirements.fields,
            work: requirements.work,
            retained_elements: requirements.retained_elements,
            retained_bytes: requirements.retained_bytes,
            peak_scratch_bytes: requirements.peak_scratch_bytes,
            allocation_events: requirements.allocation_events,
            records: requirements.headers,
            header_reads: requirements.header_reads,
            header_writes: requirements.header_writes,
            headers: requirements.headers,
            ..sparse::SparseReport::default()
        },
        path,
    )
}

fn formula_list_usage(
    report: formula_list::Report,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let transaction_work =
        report
            .work
            .checked_add(report.output_bytes)
            .ok_or(Error::LimitExceeded {
                kind: crate::package::table_cells::LimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: u64::MAX,
                path,
            })?;
    Ok(budget::Usage {
        wire_bytes: usize_u64(report.input_bytes),
        wire_fields: usize_u64(report.fields),
        wire_work: usize_u64(report.work),
        references: usize_u64(report.references),
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocations),
        list_reads: 1,
        list_writes: usize_u64(report.changed_messages),
        authored_formula_writes: usize_u64(report.assignments),
        output_bytes: usize_u64(report.output_bytes),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn formula_list_requirements_usage(
    requirements: formula_list::PlanRequirements,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let transaction_work = requirements
        .work
        .checked_add(requirements.output_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_bytes: usize_u64(requirements.input_bytes),
        wire_fields: usize_u64(requirements.fields),
        wire_work: usize_u64(requirements.work),
        references: usize_u64(requirements.references),
        retained_elements: usize_u64(requirements.retained_elements),
        retained_bytes: usize_u64(requirements.retained_bytes),
        peak_scratch_bytes: usize_u64(requirements.scratch_bytes),
        allocation_events: usize_u64(requirements.allocations),
        list_reads: 1,
        list_writes: usize_u64(requirements.changed_messages),
        authored_formula_writes: usize_u64(requirements.assignments),
        output_bytes: usize_u64(requirements.output_bytes),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn formula_list_logical_usage(
    report: formula_list::LogicalReport,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let transaction_work = report
        .work
        .checked_add(report.output_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_bytes: usize_u64(report.input_bytes),
        wire_fields: usize_u64(report.fields),
        wire_work: usize_u64(report.work),
        references: usize_u64(report.references),
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocations),
        list_reads: 1,
        output_bytes: usize_u64(report.output_bytes),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn formula_list_execution_usage(
    requirements: formula_list::ExecutionRequirements,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let transaction_work = requirements
        .work
        .checked_add(requirements.output_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_bytes: usize_u64(requirements.input_bytes),
        wire_fields: usize_u64(requirements.fields),
        wire_work: usize_u64(requirements.work),
        references: usize_u64(requirements.references),
        retained_elements: usize_u64(requirements.retained_elements),
        retained_bytes: usize_u64(requirements.retained_bytes),
        peak_scratch_bytes: usize_u64(requirements.peak_scratch_bytes),
        allocation_events: usize_u64(requirements.allocations),
        list_reads: 1,
        list_writes: usize_u64(requirements.changed_messages),
        authored_formula_writes: usize_u64(requirements.assignments),
        output_bytes: usize_u64(requirements.output_bytes),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn map_formula_list_error(
    error: formula_list::Error,
    path: crate::package::table_cells::Path,
) -> Error {
    match error {
        formula_list::Error::Limit {
            resource,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match resource {
                formula_list::Resource::InputBytes => {
                    crate::package::table_cells::LimitKind::WireBytes
                },
                formula_list::Resource::OutputBytes => {
                    crate::package::table_cells::LimitKind::OutputBytes
                },
                formula_list::Resource::Fields => {
                    crate::package::table_cells::LimitKind::WireFields
                },
                formula_list::Resource::Work => crate::package::table_cells::LimitKind::WireWork,
                formula_list::Resource::Entries | formula_list::Resource::Hosts => {
                    crate::package::table_cells::LimitKind::RetainedElements
                },
                formula_list::Resource::References => {
                    crate::package::table_cells::LimitKind::References
                },
                formula_list::Resource::RetainedElements => {
                    crate::package::table_cells::LimitKind::RetainedElements
                },
                formula_list::Resource::RetainedBytes => {
                    crate::package::table_cells::LimitKind::RetainedBytes
                },
                formula_list::Resource::ScratchBytes => {
                    crate::package::table_cells::LimitKind::PeakScratchBytes
                },
                formula_list::Resource::Allocations => {
                    crate::package::table_cells::LimitKind::AllocationEvents
                },
            },
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
            path,
        },
        formula_list::Error::Allocation { requested } => Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedBytes,
            amount: requested,
        },
        formula_list::Error::InvalidSource
        | formula_list::Error::Wire
        | formula_list::Error::Strict => Error::InvalidSource { path },
    }
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

fn source_message_references(
    source: &Package,
    route: resolve::MessageRoute,
    path: crate::package::table_cells::Path,
) -> Result<&[u64], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.archive_info.message_infos.get(route.message_index))
        .filter(|info| info.type_ == route.message_type && info.field_infos.is_empty())
        .map(|info| info.object_references.as_slice())
        .ok_or(Error::InvalidSource { path })
}

fn metadata_source_message(
    source: &Package,
    route: resolve::MessageRoute,
    path: crate::package::table_cells::Path,
) -> Result<formula_metadata::SourceMessage<'_>, Error> {
    Ok(formula_metadata::SourceMessage {
        object_id: message_object_identifier(source, route, path)?,
        payload: message_payload(source, route, path)?,
        object_references: source_message_all_references(source, route, path)?,
    })
}

fn source_message_all_references(
    source: &Package,
    route: resolve::MessageRoute,
    path: crate::package::table_cells::Path,
) -> Result<&[u64], Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.archive_info.message_infos.get(route.message_index))
        .filter(|info| info.type_ == route.message_type)
        .map(|info| info.object_references.as_slice())
        .ok_or(Error::InvalidSource { path })
}

fn formula_metadata_limits(
    source: &Package,
    remaining: budget::Remaining,
    path: crate::package::table_cells::Path,
) -> Result<formula_metadata::Limits, Error> {
    let wire = source.state.options.archive().max_iwa_stream_bytes();
    let semantic = source.state.options.semantic();
    Ok(formula_metadata::Limits {
        max_source_bytes: usize::try_from(remaining.wire_bytes)
            .map_err(|_| Error::InvalidSource { path })?
            .min(wire),
        max_output_bytes: usize::try_from(remaining.output_bytes)
            .map_err(|_| Error::InvalidSource { path })?
            .min(wire),
        max_fields: usize::try_from(remaining.wire_fields)
            .map_err(|_| Error::InvalidSource { path })?
            .min(wire),
        max_work_bytes: usize::try_from(remaining.wire_work.min(remaining.transaction_work))
            .map_err(|_| Error::InvalidSource { path })?,
        max_references: usize::try_from(remaining.references)
            .map_err(|_| Error::InvalidSource { path })?
            .min(semantic.max_references()),
        max_messages: usize::try_from(remaining.objects)
            .map_err(|_| Error::InvalidSource { path })?
            .min(semantic.max_objects()),
        max_hosts: usize::try_from(remaining.retained_elements)
            .map_err(|_| Error::InvalidSource { path })?,
        max_precedents: usize::try_from(remaining.formula_edges)
            .map_err(|_| Error::InvalidSource { path })?
            .min(semantic.max_references()),
        max_ranges: usize::try_from(remaining.range_candidates)
            .map_err(|_| Error::InvalidSource { path })?
            .min(semantic.max_references()),
        max_retained_bytes: usize::try_from(remaining.retained_bytes)
            .map_err(|_| Error::InvalidSource { path })?,
        max_scratch_bytes: usize::try_from(remaining.peak_scratch_bytes)
            .map_err(|_| Error::InvalidSource { path })?,
        max_allocations: usize::try_from(remaining.allocation_events)
            .map_err(|_| Error::InvalidSource { path })?,
        recursion_limit: 64,
    })
}

fn formula_metadata_prepare_usage(
    report: formula_metadata::Report,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    if report.output_bytes != 0 {
        return Err(Error::Verification { path });
    }
    let transaction_work = report
        .strict_work_bytes
        .checked_add(report.graph_work_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_bytes: usize_u64(report.source_bytes),
        wire_fields: usize_u64(report.fields),
        wire_work: usize_u64(report.strict_work_bytes),
        references: usize_u64(report.references),
        objects: usize_u64(report.objects),
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocations),
        formula_graph_builds: 1,
        formula_edges: usize_u64(report.precedents),
        range_candidates: usize_u64(report.ranges),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn formula_metadata_execution_usage(
    requirements: formula_metadata::ExecutionRequirements,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let transaction_work = requirements
        .work_bytes
        .checked_add(requirements.output_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_fields: usize_u64(requirements.fields),
        wire_work: usize_u64(requirements.work_bytes),
        references: usize_u64(requirements.references),
        objects: usize_u64(requirements.objects),
        retained_elements: usize_u64(requirements.retained_elements),
        retained_bytes: usize_u64(requirements.retained_bytes),
        peak_scratch_bytes: usize_u64(requirements.peak_scratch_bytes),
        allocation_events: usize_u64(requirements.allocations),
        output_bytes: usize_u64(requirements.output_bytes),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn formula_metadata_artifact_usage(
    report: formula_metadata::Report,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    let transaction_work = report
        .strict_work_bytes
        .checked_add(report.graph_work_bytes)
        .and_then(|work| work.checked_add(report.output_bytes))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_fields: usize_u64(report.fields),
        wire_work: usize_u64(report.strict_work_bytes),
        references: usize_u64(report.references),
        objects: usize_u64(report.objects),
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocations),
        output_bytes: usize_u64(report.output_bytes),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

fn map_formula_metadata_error(
    error: formula_metadata::Error,
    path: crate::package::table_cells::Path,
) -> Error {
    match error {
        formula_metadata::Error::Allocation { requested } => Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedBytes,
            amount: requested,
        },
        formula_metadata::Error::Limit {
            resource,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: if resource.contains("output") {
                crate::package::table_cells::LimitKind::OutputBytes
            } else if resource.contains("field") {
                crate::package::table_cells::LimitKind::WireFields
            } else if resource.contains("reference") {
                crate::package::table_cells::LimitKind::References
            } else if resource.contains("allocation") {
                crate::package::table_cells::LimitKind::AllocationEvents
            } else if resource.contains("scratch") {
                crate::package::table_cells::LimitKind::PeakScratchBytes
            } else if resource.contains("retained") {
                crate::package::table_cells::LimitKind::RetainedBytes
            } else {
                crate::package::table_cells::LimitKind::WireWork
            },
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
            path,
        },
        formula_metadata::Error::InvalidGraph
        | formula_metadata::Error::StrictDependency
        | formula_metadata::Error::Wire => Error::InvalidSource { path },
    }
}

fn require_exact_capacity<T>(
    values: &Vec<T>,
    expected: usize,
    _path: crate::package::table_cells::Path,
) -> Result<(), Error> {
    if core::mem::size_of::<T>() == 0 || values.capacity() == expected {
        Ok(())
    } else {
        Err(Error::Allocation {
            kind: crate::package::table_cells::LimitKind::RetainedBytes,
            amount: values.capacity(),
        })
    }
}

fn rich_archive_reference_delta(
    source: &Package,
    route: resolve::MessageRoute,
    rich: rich::ReferenceDelta,
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
    if rich.before.as_slice() != info.object_references.as_slice() {
        return Err(Error::InvalidSource { path });
    }
    let after_len = rich
        .before
        .len()
        .checked_sub(rich.removed.len())
        .ok_or(Error::InvalidSource { path })?;
    if rich.after.len() != after_len {
        return Err(Error::InvalidSource { path });
    }
    let mut after = rich.after.iter();
    for identifier in &rich.before {
        if rich.removed.contains(identifier) {
            continue;
        }
        if after.next() != Some(identifier) {
            return Err(Error::InvalidSource { path });
        }
    }
    if after.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(PreparedReferenceDelta {
        before: rich.before,
        after: rich.after,
    })
}

fn scalar_change(
    change: &crate::table::cells::Change,
    string_key: Option<u32>,
    rich_key: Option<u32>,
    path: crate::package::table_cells::Path,
) -> Result<tile::BncChange, Error> {
    Ok(match change.input_ref() {
        Some(Input::Formula { .. }) => {
            return Err(Error::UnsupportedDependency {
                path,
                kind: crate::package::table_cells::DependencyKind::Formula,
            });
        },
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

fn validate_prepared_tile_transitions(
    plan: &tile::PreparedTileRewrite<'_>,
    changes: &[crate::table::cells::Change],
    rich_keys: &[Option<u32>],
    tile_size: u32,
    formula_publication: bool,
    releases: &mut Vec<(u32, u32)>,
    path: crate::package::table_cells::Path,
) -> Result<usize, Error> {
    if changes.len() != rich_keys.len() {
        return Err(Error::Verification { path });
    }
    let mut cursor = 0usize;
    let mut validation_error = None;
    let visited = plan
        .visit_transitions(|transition| {
            if validation_error.is_some() {
                return;
            }
            while let Some(change) = changes.get(cursor) {
                let local_row = change.position().row() % tile_size;
                let column = change.position().column();
                if transition.row == local_row && transition.column == column {
                    let owned_rich = rich_keys[cursor];
                    if let Some(identifier) = transition.before_references.string {
                        if releases.len() == releases.capacity() {
                            validation_error = Some(Error::InvalidSource { path });
                            return;
                        }
                        releases.push((identifier, 1));
                    }
                    if transition.before_references.rich_text.is_some()
                        && !(owned_rich == transition.before_references.rich_text
                            && owned_rich == transition.after_references.rich_text)
                    {
                        validation_error = Some(Error::UnsupportedDependency {
                            path,
                            kind: crate::package::table_cells::DependencyKind::RichText,
                        });
                        return;
                    }
                    let authorized_formula_replacement = formula_publication
                        && matches!(change.input_ref(), Some(Input::Formula { .. }))
                        && matches!(
                            (transition.before, transition.after),
                            (
                                tile::CellValue::Formula { error: None, .. },
                                tile::CellValue::Formula {
                                    identifier,
                                    error: None,
                                },
                            ) if transition.after_references.formula == Some(identifier)
                        )
                        && transition.before_references.formula.is_some()
                        && transition.before_references.formula_error.is_none()
                        && transition.after_references.formula_error.is_none();
                    let authorized_formula_removal = formula_publication
                        && change.input_ref().is_none()
                        && matches!(
                            transition.before,
                            tile::CellValue::Formula { error: None, .. }
                        )
                        && matches!(
                            transition.after,
                            tile::CellValue::Empty | tile::CellValue::Missing
                        )
                        && transition.before_references.formula.is_some()
                        && transition.before_references.formula_error.is_none()
                        && transition.after_references.formula.is_none()
                        && transition.after_references.formula_error.is_none();
                    if !authorized_formula_replacement
                        && !authorized_formula_removal
                        && (transition.before_references.formula.is_some()
                            || transition.before_references.formula_error.is_some()
                            || matches!(
                                transition.before,
                                tile::CellValue::Formula { .. } | tile::CellValue::Error(_)
                            ))
                    {
                        validation_error = Some(Error::UnsupportedDependency {
                            path,
                            kind: crate::package::table_cells::DependencyKind::Formula,
                        });
                        return;
                    }
                    cursor += 1;
                    return;
                }
                if rich_keys[cursor].is_some() {
                    cursor += 1;
                    continue;
                }
                validation_error = Some(Error::Verification { path });
                return;
            }
            validation_error = Some(Error::Verification { path });
        })
        .map_err(|error| map_tile_error(error, path))?;
    if let Some(error) = validation_error {
        return Err(error);
    }
    while cursor < changes.len() && rich_keys[cursor].is_some() {
        cursor += 1;
    }
    if cursor != changes.len() {
        return Err(Error::Verification { path });
    }
    Ok(visited)
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
    )
    .with_retained_elements(
        usize::try_from(remaining.retained_elements)
            .map_err(|_error| Error::InvalidSource { path })?,
    ))
}

fn string_list_prepare_usage(
    report: lists::ListPrepareReport,
    path: crate::package::table_cells::Path,
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
                    maximum: u64::MAX,
                    path,
                })?,
        ),
        wire_fields: usize_u64(
            root.fields()
                .checked_add(entries.fields())
                .ok_or(Error::InvalidSource { path })?,
        ),
        wire_work: usize_u64(
            root.work_bytes()
                .checked_add(entries.work_bytes())
                .ok_or(Error::InvalidSource { path })?,
        ),
        references: usize_u64(
            root.references()
                .checked_add(entries.references())
                .ok_or(Error::InvalidSource { path })?,
        ),
        list_reads: 1,
        string_work: usize_u64(
            report
                .entries_scanned()
                .checked_add(report.strings_reused())
                .and_then(|value| value.checked_add(report.strings_added()))
                .ok_or(Error::InvalidSource { path })?,
        ),
        retained_elements: usize_u64(report.retained_elements()),
        retained_bytes: usize_u64(report.retained_bytes()),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes()),
        allocation_events: usize_u64(report.allocations()),
        transaction_work: usize_u64(report.transaction_work()),
        ..budget::Usage::default()
    })
}

fn string_list_execution_usage(
    requirements: lists::ListExecutionRequirements,
    path: crate::package::table_cells::Path,
) -> Result<budget::Usage, Error> {
    if requirements.output_bytes() > requirements.retained_bytes()
        || requirements.output_bytes() > requirements.retained_elements()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(budget::Usage {
        retained_elements: usize_u64(requirements.retained_elements()),
        retained_bytes: usize_u64(requirements.retained_bytes()),
        peak_scratch_bytes: usize_u64(requirements.peak_scratch_bytes()),
        allocation_events: usize_u64(requirements.allocations()),
        list_writes: u64::from(requirements.changed()),
        output_bytes: usize_u64(requirements.output_bytes()),
        transaction_work: usize_u64(requirements.transaction_work()),
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

fn sort_work(elements: usize, path: crate::package::table_cells::Path) -> Result<usize, Error> {
    if elements < 2 {
        return Ok(elements);
    }
    let log = usize::try_from(usize::BITS - (elements - 1).leading_zeros())
        .map_err(|_| Error::InvalidSource { path })?;
    elements
        .checked_mul(log)
        .and_then(|work| work.checked_add(elements))
        .ok_or(Error::InvalidSource { path })
}

fn binary_search_work(elements: usize) -> usize {
    if elements < 2 {
        1
    } else {
        usize::try_from(usize::BITS - (elements - 1).leading_zeros()).unwrap_or(usize::MAX)
    }
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
        formula::{BinaryOperator, CachedValue, CellReference, Expression},
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
            .find(|route| {
                let Ok(payload) = super::message_payload(&source, route.message, Path::Package)
                else {
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
        let owner_route = owner_route.message;
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
                let column = u32::try_from(index - start + 8).expect("formula column fits");
                let mut host = BncCell::minimal();
                host.set_number(2.0).expect("finite formula cache");
                host.set_formula_reference(1);
                cells[usize::try_from(column).expect("formula column fits usize")] =
                    Some(host.encode());
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
        entries.push(tst::table_data_list::ListEntry {
            key: 1,
            refcount: u32::try_from(formulas).expect("formula refcount fits"),
            formula: Some(formula(vec![
                absolute_reference_node(1, 7),
                number_node(1.0),
                operator_node(tsce::ast_node_array_archive::AstNodeType::AdditionNode),
            ])),
            ..Default::default()
        });
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
            Some(
                target
                    .dependencies
                    .formula_count
                    .checked_add(u64::try_from(formulas).expect("formula count fits"))
                    .expect("formula count addition fits"),
            ),
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
        // Prepared execution strictly reopens the candidate tile in addition
        // to the source read performed during output-free preparation.
        assert_eq!((small.tile_reads, large.tile_reads), (2, 2));
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
        assert_eq!((small.tile_reads, large.tile_reads), (2, 2));
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

    #[test]
    fn rooted_formula_aggregate_all_axes_max_minus_one_start_no_execution_or_publication() {
        let source = Package::open(fixture()).expect("native formula fixture opens");
        let expression = Expression::binary(
            BinaryOperator::Add,
            Expression::cell(CellReference::absolute(2, 1)),
            Expression::number(8.0).expect("finite formula literal"),
        )
        .expect("bounded formula expression");
        let cache = CachedValue::number(50.0).expect("finite formula cache");
        let run = || {
            source
                .edit_table_cells(0usize, 0usize)?
                .set_formula_cached(CellPosition::new(2, 2), expression.clone(), cache.clone())?
                .commit()
        };

        super::aggregate_testing::reset();
        let (baseline, _) = super::budget::testing::observe(None, run);
        baseline.expect("aggregate formula baseline commits");
        let (requirement, executions) = super::aggregate_testing::observation();
        let requirement = requirement.expect("aggregate output requirement is observed");
        assert_eq!(executions, 1, "the admitted baseline enters execution once");
        super::aggregate_testing::reset();
        super::aggregate_testing::set_authorization_limits(requirement);
        let (exact, _) = super::budget::testing::observe(None, run);
        exact.expect("the exact aggregate requirement admits the rooted transaction");
        type LimitSetter = fn(&mut super::budget::TransactionLimits, u64);
        let cases: [(
            &str,
            u64,
            LimitSetter,
            crate::package::table_cells::LimitKind,
        ); 15] = [
            (
                "updates",
                requirement.max_updates,
                |limits, maximum| limits.max_updates = maximum,
                crate::package::table_cells::LimitKind::Updates,
            ),
            (
                "owned value bytes",
                requirement.max_owned_value_bytes,
                |limits, maximum| limits.max_owned_value_bytes = maximum,
                crate::package::table_cells::LimitKind::OwnedValueBytes,
            ),
            (
                "retained elements",
                requirement.max_retained_elements,
                |limits, maximum| limits.max_retained_elements = maximum,
                crate::package::table_cells::LimitKind::RetainedElements,
            ),
            (
                "retained bytes",
                requirement.max_retained_bytes,
                |limits, maximum| limits.max_retained_bytes = maximum,
                crate::package::table_cells::LimitKind::RetainedBytes,
            ),
            (
                "scratch bytes",
                requirement.max_scratch_bytes,
                |limits, maximum| limits.max_scratch_bytes = maximum,
                crate::package::table_cells::LimitKind::PeakScratchBytes,
            ),
            (
                "allocation events",
                requirement.max_allocation_events,
                |limits, maximum| limits.max_allocation_events = maximum,
                crate::package::table_cells::LimitKind::TransactionWork,
            ),
            (
                "wire bytes",
                requirement.max_wire_bytes,
                |limits, maximum| limits.max_wire_bytes = maximum,
                crate::package::table_cells::LimitKind::WireWork,
            ),
            (
                "wire fields",
                requirement.max_wire_fields,
                |limits, maximum| limits.max_wire_fields = maximum,
                crate::package::table_cells::LimitKind::WireFields,
            ),
            (
                "wire work",
                requirement.max_wire_work,
                |limits, maximum| limits.max_wire_work = maximum,
                crate::package::table_cells::LimitKind::WireWork,
            ),
            (
                "objects",
                requirement.max_objects,
                |limits, maximum| limits.max_objects = maximum,
                crate::package::table_cells::LimitKind::Objects,
            ),
            (
                "references",
                requirement.max_references,
                |limits, maximum| limits.max_references = maximum,
                crate::package::table_cells::LimitKind::References,
            ),
            (
                "formula work",
                requirement.max_formula_work,
                |limits, maximum| limits.max_formula_work = maximum,
                crate::package::table_cells::LimitKind::FormulaWork,
            ),
            (
                "output bytes",
                requirement.max_output_bytes,
                |limits, maximum| limits.max_output_bytes = maximum,
                crate::package::table_cells::LimitKind::OutputBytes,
            ),
            (
                "reopen work",
                requirement.max_reopen_work,
                |limits, maximum| limits.max_reopen_work = maximum,
                crate::package::table_cells::LimitKind::ReopenWork,
            ),
            (
                "transaction work",
                requirement.max_transaction_work,
                |limits, maximum| limits.max_transaction_work = maximum,
                crate::package::table_cells::LimitKind::TransactionWork,
            ),
        ];

        for (name, required, set_limit, expected_kind) in cases {
            // An inactive axis is frozen at zero by the baseline requirement;
            // every active independent axis receives its own max-minus-one
            // rooted refusal below.
            if required == 0 {
                continue;
            }
            let mut limits = requirement;
            set_limit(&mut limits, required - 1);
            super::aggregate_testing::reset();
            super::aggregate_testing::set_authorization_limits(limits);
            let (result, usage) = super::budget::testing::observe(None, run);
            match result {
                Err(Error::LimitExceeded {
                    kind,
                    observed,
                    maximum,
                    ..
                }) => assert_eq!(
                    kind, expected_kind,
                    "{name}: observed {observed}, maximum {maximum}, requirement {required}"
                ),
                _ => panic!("{name} max-minus-one must refuse"),
            }
            let (refused_requirement, executions) = super::aggregate_testing::observation();
            if let Some(refused_requirement) = refused_requirement {
                assert_eq!(refused_requirement, requirement, "{name}");
            }
            assert_eq!(
                executions, 0,
                "no prepared writer executes after {name} refusal"
            );
            let usage = usage.expect("refused aggregate usage is observed");
            assert_no_publication_work(usage);
            assert_eq!(usage.output_bytes, 0, "{name}");
        }
        assert_eq!(
            source.source_bytes(),
            Package::open(fixture()).unwrap().source_bytes()
        );
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
            (1, 1)
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

    #[test]
    fn rooted_authored_formula_4096_to_8192_observed_usage_is_bounded() {
        fn measured(formulas: usize) -> super::budget::Usage {
            const HOSTS_PER_ROW: usize = 4_096;
            let source = synthetic_formula_fanout_source(formulas);
            let expression = Expression::binary(
                BinaryOperator::Add,
                Expression::cell(CellReference::absolute(1, 7)),
                Expression::number(2.0).expect("finite authored literal"),
            )
            .expect("bounded authored expression");
            let cache = CachedValue::number(3.0).expect("finite authored cache");
            let (commit, usage) = super::budget::testing::observe(None, || {
                let mut edit = source.edit_table_cells(0usize, 0usize)?;
                for index in 0..formulas {
                    edit = edit.set_formula_cached(
                        CellPosition::new(
                            u32::try_from(1 + index / HOSTS_PER_ROW).expect("formula row fits"),
                            u32::try_from(8 + index % HOSTS_PER_ROW).expect("formula column fits"),
                        ),
                        expression.clone(),
                        cache.clone(),
                    )?;
                }
                edit.commit()
            });
            let commit = commit.unwrap_or_else(|error| {
                panic!("rooted authored formula transaction commits: {error:?}; usage={usage:?}")
            });
            assert_eq!(commit.diagnostics().changed_cells(), formulas);
            assert_eq!(commit.diagnostics().refreshed_formula_caches(), formulas);
            assert!(commit.patch().authorizes_source(&source.state.source));
            assert_eq!(
                commit
                    .package()
                    .apply_table_cells(&commit.patch().inverse())
                    .expect("authored inverse applies")
                    .package()
                    .source_bytes(),
                source.source_bytes(),
            );
            usage.expect("authored formula usage is observed")
        }

        let small = measured(4_096);
        let large = measured(8_192);
        for (name, left, right) in [
            ("updates", small.updates, large.updates),
            (
                "retained_elements",
                small.retained_elements,
                large.retained_elements,
            ),
            ("retained_bytes", small.retained_bytes, large.retained_bytes),
            (
                "peak_scratch_bytes",
                small.peak_scratch_bytes,
                large.peak_scratch_bytes,
            ),
            (
                "allocation_events",
                small.allocation_events,
                large.allocation_events,
            ),
            ("wire_bytes", small.wire_bytes, large.wire_bytes),
            ("wire_fields", small.wire_fields, large.wire_fields),
            ("wire_work", small.wire_work, large.wire_work),
            ("lookups", small.lookups, large.lookups),
            ("formula_nodes", small.formula_nodes, large.formula_nodes),
            ("formula_edges", small.formula_edges, large.formula_edges),
            ("cache_hosts", small.cache_hosts, large.cache_hosts),
            (
                "authored_formula_writes",
                small.authored_formula_writes,
                large.authored_formula_writes,
            ),
            ("formula_work", small.formula_work, large.formula_work),
            ("output_bytes", small.output_bytes, large.output_bytes),
            ("reopen_work", small.reopen_work, large.reopen_work),
            ("locality_bytes", small.locality_bytes, large.locality_bytes),
            (
                "transaction_work",
                small.transaction_work,
                large.transaction_work,
            ),
        ] {
            assert_linear_growth(left, right, name);
        }
        assert_eq!(small.formula_graph_builds, large.formula_graph_builds);
        assert_eq!(small.component_encodes, large.component_encodes);
        assert_eq!(small.components_reassembled, large.components_reassembled);
        assert_eq!(
            small.output_artifact_allocations,
            large.output_artifact_allocations
        );
        assert_eq!(small.candidate_reopens, large.candidate_reopens);
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
    fn synthetic_unique_rich_output_is_behind_the_aggregate_execution_barrier() -> TestResult {
        let (source, _storage_identifier, _style_identifier) = synthetic_unique_rich_source();
        let position = CellPosition::from_a1("C2")?;
        let run = || {
            source
                .edit_table_cells(0usize, 0usize)?
                .set(position, Input::text("Changed rich text")?)?
                .commit()
        };

        super::aggregate_testing::reset();
        let (baseline, _) = super::budget::testing::observe(None, run);
        baseline?;
        let (requirement, executions) = super::aggregate_testing::observation();
        let requirement = requirement.expect("rich execution requirement is observed");
        assert_eq!(executions, 1);

        for (name, limits) in [
            ("output", {
                let mut limits = requirement;
                limits.max_output_bytes = limits
                    .max_output_bytes
                    .checked_sub(1)
                    .expect("rich execution produces output");
                limits
            }),
            ("allocation", {
                let mut limits = requirement;
                limits.max_allocation_events = limits
                    .max_allocation_events
                    .checked_sub(1)
                    .expect("rich execution allocates candidates");
                limits
            }),
        ] {
            super::aggregate_testing::reset();
            super::aggregate_testing::set_authorization_limits(limits);
            let (result, usage) = super::budget::testing::observe(None, run);
            assert!(
                matches!(result, Err(Error::LimitExceeded { .. })),
                "{name} max-minus-one refuses"
            );
            assert_eq!(super::aggregate_testing::observation().1, 0, "{name}");
            let usage = usage.expect("refused rich usage is observed");
            assert_no_publication_work(usage);
            assert_eq!(usage.output_bytes, 0, "{name}");
        }
        super::aggregate_testing::reset();
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
