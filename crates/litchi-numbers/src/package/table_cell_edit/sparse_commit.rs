//! Grouped sparse-storage publication for selector-first cell batches.

use std::{cell::Cell, mem::size_of, sync::Arc};

use litchi_iwa_core::{ArchiveObject, MessageInfo, RawMessage};
use litchi_iwa_protos::package_metadata_codec::{
    Batch, ExternalReferenceAddition, ObjectUuidAddition, RewriteExecutionRequirements,
    RewriteOptions, RewriteOutput, RewriteReport, UuidBits, rewrite_package_metadata,
};

use crate::{
    Package,
    package::table_cells::{DependencyKind, LimitKind, Path},
    table::cells::{
        Commit, Diagnostics, DirectionalMessage, Error, EvidenceChangeKind, FieldReferenceRoute,
        Input, MessageReferenceRoute, Patch, PatchEvidence, PhysicalLocation, ReferenceEvidence,
        ReferenceSpan,
    },
};

use super::{
    authorize_remaining, budget, map_rewrite_error, map_tile_error, message_object_identifier,
    message_payload, metadata, preview_mask, resolve, rewrite, scalar_change, sparse, tile,
    usize_u64, verify_evidence_locality_with_report, verify_semantic_changes,
};

const TILE_MESSAGE_TYPE: u32 = 6002;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6006;

#[cfg(test)]
pub(super) mod testing {
    use std::cell::Cell;

    std::thread_local! {
        static EVIDENCE_WORK_LIMIT: Cell<Option<u64>> = const { Cell::new(None) };
        static EVIDENCE_SHAPE_VISITS: Cell<u64> = const { Cell::new(0) };
        static EVIDENCE_SHAPE_REQUIREMENT: Cell<Option<u64>> = const { Cell::new(None) };
    }

    struct Reset;

    impl Drop for Reset {
        fn drop(&mut self) {
            EVIDENCE_WORK_LIMIT.set(None);
        }
    }

    pub(crate) fn with_evidence_work_limit<T>(
        limit: Option<u64>,
        operation: impl FnOnce() -> T,
    ) -> (T, u64, Option<u64>) {
        EVIDENCE_SHAPE_VISITS.set(0);
        EVIDENCE_SHAPE_REQUIREMENT.set(None);
        EVIDENCE_WORK_LIMIT.set(limit);
        let reset = Reset;
        let value = operation();
        let visits = EVIDENCE_SHAPE_VISITS.get();
        let requirement = EVIDENCE_SHAPE_REQUIREMENT.get();
        drop(reset);
        (value, visits, requirement)
    }

    pub(super) fn shape_work_remaining(remaining: u64) -> u64 {
        EVIDENCE_WORK_LIMIT
            .get()
            .map_or(remaining, |limit| remaining.min(limit))
    }

    pub(super) fn record_shape_visit() {
        EVIDENCE_SHAPE_VISITS.set(EVIDENCE_SHAPE_VISITS.get() + 1);
    }

    pub(super) fn record_shape_requirement(requirement: u64) {
        EVIDENCE_SHAPE_REQUIREMENT.set(Some(requirement));
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "the transaction boundary keeps selector and diagnostics facts explicit"
)]
pub(super) fn commit_sparse_scalar_tiles(
    source: &Package,
    path: Path,
    sheet_position: usize,
    table_position: usize,
    target: &resolve::Target,
    mut budget: budget::TransactionBudget,
    changes: Vec<crate::table::cells::Change>,
    requested: usize,
) -> Result<Commit, Error> {
    if changes
        .iter()
        .any(|change| matches!(change.input_ref(), Some(Input::Text(_))))
    {
        return Err(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::SharedString,
        });
    }
    let remaining = authorize_remaining(&mut budget)?;
    let limits = sparse_limits(source, changes.len(), remaining, path)?;
    let model_payload = message_payload(source, target.storage.model, path)?;
    let (data_store, mut sparse_report) =
        sparse::table_model_data_store_with_report(model_payload, limits)
            .map_err(|error| map_sparse_error(error, path))?;

    let mut cells = Vec::new();
    cells
        .try_reserve_exact(changes.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: changes.len(),
        })?;
    for change in &changes {
        cells.push(sparse::Cell {
            row: change.position().row(),
            column: change.position().column(),
        });
    }
    let mut header_sources = Vec::new();
    header_sources
        .try_reserve_exact(target.storage.row_headers.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: target.storage.row_headers.len(),
        })?;
    for &route in &target.storage.row_headers {
        header_sources.push(sparse::HeaderBucketSource {
            object_id: message_object_identifier(source, route, path)?,
            payload: message_payload(source, route, path)?,
        });
    }
    let (mut plan, plan_report) = sparse::SparsePlan::build_with_report(
        sparse::SparseRequest {
            data_store,
            cells: &cells,
            columns: target.native.columns,
            tile_size: target.storage.tile_size,
            row_header_buckets: &header_sources,
        },
        limits,
    )
    .map_err(|error| map_sparse_error(error, path))?;
    sparse_report
        .merge(plan_report)
        .map_err(|error| map_sparse_error(error, path))?;
    if plan.new_objects().is_empty() {
        budget.cancel_authorization();
        return Err(Error::InvalidSource { path });
    }

    let metadata_additions = plan
        .new_objects()
        .len()
        .checked_mul(2)
        .ok_or(Error::InvalidSource { path })?;
    let metadata_options = metadata_options(source, remaining, metadata_additions, path)?;
    let inspected = metadata::inspect(source, metadata_options, path)?;
    let (identifiers, uuids) =
        inspected.allocate_identifiers(plan.new_objects().len(), source.source_bytes(), path)?;
    let template_route = target
        .storage
        .tiles
        .first()
        .ok_or(Error::UnsupportedDependency {
            path,
            kind: DependencyKind::CellStorage,
        })?
        .message;
    let tile_component = template_route.component_index;
    let header_component = target
        .storage
        .row_headers
        .first()
        .copied()
        .unwrap_or(target.storage.column_headers)
        .component_index;
    let model_component = target.storage.model.component_index;

    let mut assignments = Vec::new();
    assignments
        .try_reserve_exact(plan.new_objects().len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::Objects,
            amount: plan.new_objects().len(),
        })?;
    let mut uuid_additions = Vec::new();
    uuid_additions
        .try_reserve_exact(plan.new_objects().len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: plan.new_objects().len(),
        })?;
    let mut external_additions = Vec::new();
    external_additions
        .try_reserve_exact(plan.new_objects().len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: plan.new_objects().len(),
        })?;
    let model_selector = inspected.selector(source, model_component, path)?;
    for ((request, &object_id), &uuid) in plan.new_objects().iter().zip(&identifiers).zip(&uuids) {
        let component = match request.kind {
            sparse::NewObjectKind::Tile { .. } => tile_component,
            sparse::NewObjectKind::RowHeaderBucket { .. } => header_component,
        };
        let selector = inspected.selector(source, component, path)?;
        assignments.push(sparse::ObjectAssignment {
            slot: request.slot,
            object_id,
            kind: request.kind,
            metadata_registered: false,
        });
        uuid_additions.push(ObjectUuidAddition::new(selector, object_id, uuid));
        if component != model_component {
            external_additions.push(ExternalReferenceAddition::new(
                model_selector,
                selector,
                object_id,
                None,
            ));
        }
    }
    let new_last = *identifiers.last().ok_or(Error::InvalidSource { path })?;
    let metadata_source = message_payload(source, inspected.route, path)?;
    let metadata_output = rewrite_package_metadata(
        metadata_source,
        Batch::new(
            inspected.last_object_identifier,
            new_last,
            &uuid_additions,
            &external_additions,
        ),
        metadata_options,
    )
    .map_err(|error| map_metadata_error(error, path))?;
    for assignment in &mut assignments {
        assignment.metadata_registered = true;
    }

    let archive_limits = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource { path })?
        .limits()
        .effective_archive_limits()
        .map_err(|_error| Error::InvalidSource { path })?;
    let tile_template = message_payload(source, template_route, path)?;
    let component_capacity = changes
        .len()
        .checked_add(4)
        .ok_or(Error::InvalidSource { path })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(component_capacity)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: component_capacity,
        })?;
    let mut final_rows = Vec::new();
    final_rows
        .try_reserve_exact(changes.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: changes.len(),
        })?;
    let mut observed = metadata_usage(&inspected, &metadata_output, &cells, &header_sources)?;

    let mut start = 0usize;
    while start < changes.len() {
        let tile_id = changes[start].position().row() / target.storage.tile_size;
        let end = changes[start..]
            .iter()
            .position(|change| change.position().row() / target.storage.tile_size != tile_id)
            .map_or(changes.len(), |offset| start + offset);
        let mut tile_changes = Vec::new();
        tile_changes
            .try_reserve_exact(end - start)
            .map_err(|_error| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: end - start,
            })?;
        for change in &changes[start..end] {
            tile_changes.push(tile::TileChange {
                row: change.position().row() % target.storage.tile_size,
                column: change.position().column(),
                change: scalar_change(change, None, None, path)?,
            });
        }
        let tile_limits = sparse_tile_limits(source, target, &tile_changes, remaining, path)?;
        let existing = target
            .storage
            .tiles
            .binary_search_by_key(&tile_id, |route| route.tile_id)
            .ok()
            .and_then(|index| target.storage.tiles.get(index));
        let outcome = match existing {
            Some(route) => tile::rewrite_tile(tile::TileRewriteRequest {
                source: message_payload(source, route.message, path)?,
                columns: target.native.columns,
                changes: &tile_changes,
                limits: tile_limits,
            }),
            None => tile::rewrite_new_tile(
                tile_template,
                target.native.columns,
                &tile_changes,
                tile_limits,
            ),
        }
        .map_err(|error| map_tile_error(error, path))?;
        merge_tile_usage(&mut observed, outcome.report, path)?;
        for transition in &outcome.transitions {
            if transition.before_references.string.is_some()
                || transition.before_references.rich_text.is_some()
                || transition.before_references.formula.is_some()
                || transition.before_references.formula_error.is_some()
            {
                budget.cancel_authorization();
                return Err(Error::UnsupportedDependency {
                    path,
                    kind: DependencyKind::CellStorage,
                });
            }
        }
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
        let payload = outcome.payload.ok_or(Error::Verification { path })?;
        match existing {
            Some(route) => push_message(
                &mut edits,
                route.message,
                payload,
                None,
                component_capacity,
                path,
            )?,
            None => {
                let assignment = assignment_for_tile(&assignments, tile_id, path)?;
                let object = one_message_object(
                    assignment.object_id,
                    TILE_MESSAGE_TYPE,
                    payload,
                    archive_limits,
                    path,
                )?;
                push_object(&mut edits, tile_component, object, component_capacity, path)?;
            },
        }
        start = end;
    }
    final_rows.sort_unstable_by_key(|row| row.row);
    if final_rows.windows(2).any(|pair| pair[0].row == pair[1].row) {
        budget.cancel_authorization();
        return Err(Error::InvalidSource { path });
    }
    let mut new_header_rows = Vec::new();
    new_header_rows
        .try_reserve_exact(plan.new_headers().len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: plan.new_headers().len(),
        })?;
    for row in &final_rows {
        if plan
            .new_headers()
            .binary_search_by_key(&row.row, |header| header.row)
            .is_ok()
        {
            new_header_rows.push(*row);
        }
    }
    let header_count_report = plan
        .synchronize_new_header_counts_with_report(&new_header_rows, limits)
        .map_err(|error| map_sparse_error(error, path))?;
    sparse_report
        .merge(header_count_report)
        .map_err(|error| map_sparse_error(error, path))?;

    for (bucket_index, &route) in target.storage.row_headers.iter().enumerate() {
        let bucket_index =
            u32::try_from(bucket_index).map_err(|_error| Error::InvalidSource { path })?;
        let start =
            final_rows.partition_point(|row| row.row / sparse::HEADER_BUCKET_ROWS < bucket_index);
        let end =
            final_rows.partition_point(|row| row.row / sparse::HEADER_BUCKET_ROWS <= bucket_index);
        let (payload, report) = sparse::rewrite_header_bucket_final_rows_with_report(
            message_payload(source, route, path)?,
            bucket_index,
            &plan,
            &final_rows[start..end],
            limits,
        )
        .map_err(|error| map_sparse_error(error, path))?;
        sparse_report
            .merge(report)
            .map_err(|error| map_sparse_error(error, path))?;
        if let Some(payload) = payload {
            push_message(&mut edits, route, payload, None, component_capacity, path)?;
        }
    }
    let (data_rewrite, data_report) =
        sparse::rewrite_data_store_with_report(data_store, &plan, &assignments, limits)
            .map_err(|error| map_sparse_error(error, path))?;
    sparse_report
        .merge(data_report)
        .map_err(|error| map_sparse_error(error, path))?;
    for (slot, payload) in data_rewrite.new_header_buckets {
        let assignment = assignments
            .iter()
            .find(|assignment| assignment.slot == slot)
            .ok_or(Error::InvalidSource { path })?;
        let object = one_message_object(
            assignment.object_id,
            HEADER_BUCKET_MESSAGE_TYPE,
            payload,
            archive_limits,
            path,
        )?;
        push_object(
            &mut edits,
            header_component,
            object,
            component_capacity,
            path,
        )?;
    }
    let (model_replacement, model_report) = sparse::rewrite_table_model_data_store_with_report(
        model_payload,
        &data_rewrite.data_store,
        limits,
    )
    .map_err(|error| map_sparse_error(error, path))?;
    sparse_report
        .merge(model_report)
        .map_err(|error| map_sparse_error(error, path))?;
    merge_usage(&mut observed, sparse_usage(sparse_report, path)?, path)?;
    let references = model_reference_delta(source, target, &assignments, path)?;
    push_message(
        &mut edits,
        target.storage.model,
        model_replacement,
        Some(references),
        component_capacity,
        path,
    )?;
    push_message(
        &mut edits,
        inspected.route,
        metadata_output.into_bytes(),
        None,
        component_capacity,
        path,
    )?;
    normalize_edits(&mut edits, path)?;

    let previews =
        rewrite::root_preview_deletions(source).map_err(|error| map_rewrite_error(error, path))?;
    let preview_membership = preview_mask(&previews, path)?;
    let evidence_count = evidence_message_count(&edits, path)?;
    let remaining = budget.authorization_remaining()?;
    let shape_work_remaining = remaining
        .transaction_work
        .checked_sub(observed.transaction_work)
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: observed.transaction_work,
            maximum: remaining.transaction_work,
            path,
        })?;
    #[cfg(test)]
    let shape_work_remaining = testing::shape_work_remaining(shape_work_remaining);
    let reference_shape = evidence_reference_shape(source, &edits, shape_work_remaining, path)?;
    let evidence_usage = patch_evidence_usage(evidence_count, reference_shape, path)?;
    let locality_plan_usage = locality_plan_usage(evidence_count, path)?;
    merge_usage(&mut observed, evidence_usage, path)?;
    merge_usage(&mut observed, locality_plan_usage, path)?;
    preflight_usage(observed, budget.authorization_remaining()?, path)?;
    let (evidence_messages, reference_evidence) =
        prebuild_evidence(source, &edits, reference_shape, path)?;
    if evidence_messages.len() != evidence_count {
        budget.cancel_authorization();
        return Err(Error::Verification { path });
    }
    let evidence_messages = Arc::new(evidence_messages);
    let precharge_error = Cell::new(None);
    let observed = Cell::new(observed);
    let component_remaining = budget.authorization_remaining()?;
    let outcome = rewrite::rewrite_staged_with_evidence_authorization(
        source,
        rewrite::StagedRewritePlan {
            component_edits: edits,
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
            let mut combined = observed.get();
            if let Err(error) = merge_usage(&mut combined, envelope, path) {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            if let Err(error) = preflight_usage(combined, component_remaining, path) {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            Ok(())
        },
        |reservation, component_cost| {
            let component_usage = match component_usage(component_cost, path) {
                Ok(usage) => usage,
                Err(error) => {
                    precharge_error.set(Some(error));
                    return Err(rewrite::RewriteError::Precharge);
                },
            };
            if let Err(error) = budget.preauthorize_publication(reservation) {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            let mut combined = observed.get();
            if let Err(error) = merge_usage(&mut combined, component_usage, path) {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            let remaining = match budget.authorization_remaining() {
                Ok(remaining) => remaining,
                Err(error) => {
                    precharge_error.set(Some(error));
                    return Err(rewrite::RewriteError::Precharge);
                },
            };
            if let Err(error) = preflight_usage(combined, remaining, path) {
                precharge_error.set(Some(error));
                return Err(rewrite::RewriteError::Precharge);
            }
            observed.set(combined);
            Ok(())
        },
    )
    .map_err(|error| {
        if budget.publication_is_authorized() {
            budget.cancel_publication();
        }
        budget.cancel_authorization();
        precharge_error
            .get()
            .unwrap_or_else(|| map_rewrite_error(error, path))
    })?;
    let evidence = match PatchEvidence::new(
        evidence_messages,
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
    let mut final_observed = observed.get();
    if let Err(error) = preflight_usage(final_observed, budget.authorization_remaining()?, path) {
        budget.cancel_publication();
        budget.cancel_authorization();
        return Err(error);
    }
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
    if let Err(error) = merge_usage(&mut final_observed, locality_report.usage(), path) {
        budget.cancel_publication();
        budget.cancel_authorization();
        return Err(error);
    }
    let mut publication = outcome.publication;
    publication.locality_bytes = 0;
    publication.locality_work = 0;
    budget.finish_publication(publication);
    if let Err(error) = preflight_usage(final_observed, budget.authorization_remaining()?, path) {
        budget.cancel_authorization();
        return Err(error);
    }
    if let Err(error) = budget.record_authorized(final_observed) {
        if budget.authorization_is_pending() {
            budget.cancel_authorization();
        }
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
            0,
            previews.len(),
        ),
    ))
}

pub(super) struct FormulaDependencyPublication {
    pub(super) outcome: rewrite::RewriteOutcome,
    pub(super) evidence: PatchEvidence,
    pub(super) preview_count: usize,
}

pub(super) fn publish_formula_dependency_objects(
    source: &Package,
    mut replacements: Vec<super::PreparedReplacement>,
    dependency_tiles: super::PreparedFormulaDependencyTiles,
    new_tiles: Vec<super::formula_metadata::MessageEdit>,
    future_publication: rewrite::FuturePublicationReservation,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<FormulaDependencyPublication, Error> {
    if dependency_tiles.assignments.is_empty()
        || dependency_tiles.assignments.len() != new_tiles.len()
    {
        return Err(Error::InvalidSource { path });
    }
    let archive_limits = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource { path })?
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let capacity = replacements
        .len()
        .checked_add(new_tiles.len())
        .and_then(|value| value.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    let message_count = replacements
        .len()
        .checked_add(new_tiles.len())
        .ok_or(Error::InvalidSource { path })?;
    let reference_identifiers = replacements.iter().try_fold(0usize, |total, replacement| {
        replacement.references.as_ref().map_or(Ok(total), |delta| {
            total
                .checked_add(delta.before.len())
                .and_then(|total| total.checked_add(delta.after.len()))
                .ok_or(Error::InvalidSource { path })
        })
    })?;
    let preview_bytes = rewrite::ROOT_PREVIEWS
        .len()
        .checked_mul(size_of::<&'static str>())
        .ok_or(Error::InvalidSource { path })?;
    let staging_bytes = capacity
        .checked_mul(size_of::<rewrite::ComponentEdit>())
        .and_then(|bytes| {
            replacements
                .len()
                .checked_mul(size_of::<rewrite::MessageEdit>())
                .and_then(|messages| bytes.checked_add(messages))
        })
        .and_then(|bytes| {
            new_tiles
                .len()
                .checked_mul(
                    size_of::<ArchiveObject>()
                        .checked_add(size_of::<RawMessage>())?
                        .checked_add(size_of::<MessageInfo>())?
                        .checked_add(3usize.checked_mul(size_of::<u32>())?)?,
                )
                .and_then(|objects| bytes.checked_add(objects))
        })
        .and_then(|bytes| {
            reference_identifiers
                .checked_mul(size_of::<u64>())
                .and_then(|references| bytes.checked_add(references))
        })
        .and_then(|bytes| bytes.checked_add(preview_bytes))
        .ok_or(Error::InvalidSource { path })?;
    let staging_elements = capacity
        .checked_add(replacements.len())
        .and_then(|elements| elements.checked_add(new_tiles.len().checked_mul(6)?))
        .and_then(|elements| elements.checked_add(reference_identifiers))
        .and_then(|elements| elements.checked_add(rewrite::ROOT_PREVIEWS.len()))
        .ok_or(Error::InvalidSource { path })?;
    let staging_allocations = u64::from(capacity != 0)
        .checked_add(usize_u64(replacements.len()))
        .and_then(|events| events.checked_add(usize_u64(new_tiles.len()).checked_mul(5)?))
        .and_then(|events| events.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    let staging_usage = budget::Usage {
        retained_elements: usize_u64(staging_elements),
        retained_bytes: usize_u64(staging_bytes),
        peak_scratch_bytes: usize_u64(staging_bytes),
        allocation_events: staging_allocations,
        objects: usize_u64(new_tiles.len()),
        references: usize_u64(reference_identifiers),
        transaction_work: usize_u64(
            message_count
                .checked_mul(8)
                .and_then(|work| {
                    source
                        .state
                        .components
                        .iter_archives()
                        .count()
                        .checked_mul(rewrite::ROOT_PREVIEWS.len())
                        .and_then(|preview_work| work.checked_add(preview_work))
                })
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    let staging_remaining = authorize_remaining(budget)?;
    preflight_usage(staging_usage, staging_remaining, path)?;
    replacements.sort_unstable_by_key(|replacement| {
        (
            replacement.route.component_index,
            replacement.route.object_index,
            replacement.route.message_index,
        )
    });
    if replacements.windows(2).any(|pair| {
        (
            pair[0].route.component_index,
            pair[0].route.object_index,
            pair[0].route.message_index,
        ) >= (
            pair[1].route.component_index,
            pair[1].route.object_index,
            pair[1].route.message_index,
        )
    }) {
        budget.cancel_authorization();
        return Err(Error::InvalidSource { path });
    }
    let component_count = replacements
        .iter()
        .map(|replacement| replacement.route.component_index)
        .chain(core::iter::once(dependency_tiles.component_index))
        .fold((None, 0usize), |(previous, count), component| {
            if previous == Some(component) {
                (previous, count)
            } else {
                (Some(component), count.saturating_add(1))
            }
        })
        .1;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(component_count)
        .map_err(|_| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: component_count,
        })?;
    let mut start = 0usize;
    while start < replacements.len() {
        let component_index = replacements[start].route.component_index;
        let end = start
            + replacements[start..].partition_point(|replacement| {
                replacement.route.component_index == component_index
            });
        let mut messages = Vec::new();
        messages
            .try_reserve_exact(end - start)
            .map_err(|_| Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: end - start,
            })?;
        let mut objects = Vec::new();
        let object_count = usize::from(component_index == dependency_tiles.component_index)
            .checked_mul(new_tiles.len())
            .ok_or(Error::InvalidSource { path })?;
        objects
            .try_reserve_exact(object_count)
            .map_err(|_| Error::Allocation {
                kind: LimitKind::Objects,
                amount: object_count,
            })?;
        edits.push(rewrite::ComponentEdit {
            component_index,
            messages,
            object_deletions: Vec::new(),
            new_objects: objects,
        });
        start = end;
    }
    if edits
        .binary_search_by_key(&dependency_tiles.component_index, |edit| {
            edit.component_index
        })
        .is_err()
    {
        let index = edits
            .binary_search_by_key(&dependency_tiles.component_index, |edit| {
                edit.component_index
            })
            .unwrap_err();
        let mut objects = Vec::new();
        objects
            .try_reserve_exact(new_tiles.len())
            .map_err(|_| Error::Allocation {
                kind: LimitKind::Objects,
                amount: new_tiles.len(),
            })?;
        edits.insert(
            index,
            rewrite::ComponentEdit {
                component_index: dependency_tiles.component_index,
                messages: Vec::new(),
                object_deletions: Vec::new(),
                new_objects: objects,
            },
        );
    }
    for replacement in replacements {
        let references = replacement
            .references
            .map(|references| rewrite::ReferenceDelta {
                aggregate_before: references.before,
                aggregate_after: references.after,
                fields: Vec::new(),
            });
        let edit = edits
            .binary_search_by_key(&replacement.route.component_index, |edit| {
                edit.component_index
            })
            .ok()
            .and_then(|index| edits.get_mut(index))
            .ok_or(Error::InvalidSource { path })?;
        if edit.messages.len() == edit.messages.capacity() {
            budget.cancel_authorization();
            return Err(Error::Allocation {
                kind: LimitKind::RetainedElements,
                amount: edit.messages.len().saturating_add(1),
            });
        }
        edit.messages.push(rewrite::MessageEdit {
            object_index: replacement.route.object_index,
            message_index: replacement.route.message_index,
            expected_type: replacement.route.message_type,
            payload: replacement.payload,
            references,
        });
    }
    for (assignment, tile) in dependency_tiles.assignments.iter().zip(new_tiles) {
        let payload = tile.payload.ok_or(Error::InvalidSource { path })?;
        if tile.object_id != assignment.object_id || !tile.object_references.is_empty() {
            return Err(Error::InvalidSource { path });
        }
        let object =
            one_message_object(assignment.object_id, 4_009, payload, archive_limits, path)?;
        let edit = edits
            .binary_search_by_key(&dependency_tiles.component_index, |edit| {
                edit.component_index
            })
            .ok()
            .and_then(|index| edits.get_mut(index))
            .ok_or(Error::InvalidSource { path })?;
        if edit.new_objects.len() == edit.new_objects.capacity() {
            budget.cancel_authorization();
            return Err(Error::Allocation {
                kind: LimitKind::Objects,
                amount: edit.new_objects.len().saturating_add(1),
            });
        }
        edit.new_objects.push(object);
    }
    let previews =
        rewrite::root_preview_deletions(source).map_err(|error| map_rewrite_error(error, path))?;
    normalize_edits(&mut edits, path)?;
    budget.record_authorized(staging_usage)?;
    let preview_membership = preview_mask(&previews, path)?;
    let evidence_count = evidence_message_count(&edits, path)?;
    let remaining = authorize_remaining(budget)?;
    let reference_shape =
        evidence_reference_shape(source, &edits, remaining.transaction_work, path)?;
    let mut observed = patch_evidence_usage(evidence_count, reference_shape, path)?;
    merge_usage(
        &mut observed,
        locality_plan_usage(evidence_count, path)?,
        path,
    )?;
    preflight_usage(observed, remaining, path)?;
    let (evidence_messages, reference_evidence) =
        prebuild_evidence(source, &edits, reference_shape, path)?;
    let evidence = PatchEvidence::new(
        Arc::new(evidence_messages),
        reference_evidence,
        preview_membership,
        previews.len(),
        0,
        path,
    )?;
    let precharge_error = Cell::new(None);
    let observed_cell = Cell::new(observed);
    let outcome = match rewrite::rewrite_staged_with_evidence_authorization(
        source,
        rewrite::StagedRewritePlan {
            component_edits: edits,
            preview_deletions: &previews,
        },
        rewrite::EvidenceRetention::Omit,
        |reservation| {
            if !rewrite::component_admission_shape_fits(reservation, future_publication.component) {
                precharge_error.set(Some(Error::Verification { path }));
                return Err(rewrite::RewriteError::Precharge);
            }
            let usage = component_reservation_usage(future_publication.component, path).map_err(
                |error| {
                    precharge_error.set(Some(error));
                    rewrite::RewriteError::Precharge
                },
            )?;
            let mut combined = observed_cell.get();
            merge_usage(&mut combined, usage, path)
                .and_then(|()| preflight_usage(combined, remaining, path))
                .map_err(|error| {
                    precharge_error.set(Some(error));
                    rewrite::RewriteError::Precharge
                })
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
            budget
                .preauthorize_publication(reservation)
                .map_err(|error| {
                    precharge_error.set(Some(error));
                    rewrite::RewriteError::Precharge
                })?;
            let usage = component_usage(cost, path).map_err(|error| {
                precharge_error.set(Some(error));
                rewrite::RewriteError::Precharge
            })?;
            let mut combined = observed_cell.get();
            merge_usage(&mut combined, usage, path)
                .and_then(|()| preflight_usage(combined, budget.authorization_remaining()?, path))
                .map_err(|error| {
                    precharge_error.set(Some(error));
                    rewrite::RewriteError::Precharge
                })?;
            observed_cell.set(combined);
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
    let locality = verify_evidence_locality_with_report(
        source,
        &outcome.package,
        &evidence,
        outcome.publication.locality_work,
        path,
    )?;
    let mut exact = observed_cell.get();
    merge_usage(&mut exact, locality.usage(), path)?;
    let mut publication = outcome.publication;
    publication.locality_bytes = 0;
    publication.locality_work = 0;
    budget.finish_publication(publication);
    exact.locality_bytes = locality.bytes;
    budget.record_authorized(exact)?;
    budget.release_retained(
        usize_u64(dependency_tiles.retained_elements),
        usize_u64(dependency_tiles.retained_bytes),
    )?;
    Ok(FormulaDependencyPublication {
        outcome,
        evidence,
        preview_count: previews.len(),
    })
}

pub(super) fn sparse_limits(
    source: &Package,
    cells: usize,
    remaining: budget::Remaining,
    path: Path,
) -> Result<sparse::SparseLimits, Error> {
    let maximum_wire = source.state.options.archive().max_iwa_stream_bytes();
    Ok(sparse::SparseLimits {
        max_fields: usize::try_from(remaining.wire_fields)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire),
        max_cells: cells,
        max_records: usize::try_from(remaining.retained_elements)
            .map_err(|_error| Error::InvalidSource { path })?,
        max_output_bytes: usize::try_from(
            remaining.retained_bytes.min(remaining.peak_scratch_bytes),
        )
        .map_err(|_error| Error::InvalidSource { path })?
        .min(maximum_wire),
        max_work: usize::try_from(remaining.transaction_work)
            .map_err(|_error| Error::InvalidSource { path })?,
        max_retained_elements: usize::try_from(remaining.retained_elements)
            .map_err(|_error| Error::InvalidSource { path })?,
        max_retained_bytes: usize::try_from(remaining.retained_bytes)
            .map_err(|_error| Error::InvalidSource { path })?,
        max_scratch_bytes: usize::try_from(remaining.peak_scratch_bytes)
            .map_err(|_error| Error::InvalidSource { path })?,
        max_allocation_events: usize::try_from(remaining.allocation_events)
            .map_err(|_error| Error::InvalidSource { path })?,
        max_references: usize::try_from(remaining.references)
            .map_err(|_error| Error::InvalidSource { path })?,
    })
}

pub(super) fn metadata_options(
    source: &Package,
    remaining: budget::Remaining,
    additions: usize,
    path: Path,
) -> Result<RewriteOptions, Error> {
    let maximum_wire = source.state.options.archive().max_iwa_stream_bytes();
    Ok(RewriteOptions::new(
        usize::try_from(remaining.wire_bytes)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire),
        usize::try_from(
            remaining
                .retained_bytes
                .min(remaining.peak_scratch_bytes)
                .min(remaining.output_bytes),
        )
        .map_err(|_error| Error::InvalidSource { path })?
        .min(maximum_wire),
        usize::try_from(remaining.wire_fields)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire),
        usize::try_from(remaining.wire_work.min(remaining.transaction_work))
            .map_err(|_error| Error::InvalidSource { path })?,
        64,
        usize::try_from(remaining.objects).map_err(|_error| Error::InvalidSource { path })?,
        usize::try_from(remaining.references).map_err(|_error| Error::InvalidSource { path })?,
        additions,
    ))
}

fn sparse_tile_limits(
    source: &Package,
    target: &resolve::Target,
    changes: &[tile::TileChange],
    remaining: budget::Remaining,
    path: Path,
) -> Result<tile::TileLimits, Error> {
    let maximum_wire = source.state.options.archive().max_iwa_stream_bytes();
    let maximum_work = maximum_wire
        .checked_mul(32)
        .ok_or(Error::InvalidSource { path })?;
    let max_work =
        budget::tile_work_ceiling(remaining, usize_u64(maximum_work), usize_u64(changes.len()))?;
    Ok(tile::TileLimits::new(
        usize::try_from(remaining.wire_bytes)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire),
        usize::try_from(remaining.retained_bytes.min(remaining.peak_scratch_bytes))
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire),
        usize::try_from(remaining.wire_fields)
            .map_err(|_error| Error::InvalidSource { path })?
            .min(maximum_wire),
        max_work,
        usize::try_from(target.storage.tile_size)
            .map_err(|_error| Error::InvalidSource { path })?,
        source.state.options.semantic().max_materialized_cells(),
    ))
}

fn metadata_usage(
    inspected: &metadata::Inspection,
    output: &RewriteOutput,
    cells: &[sparse::Cell],
    headers: &[sparse::HeaderBucketSource<'_>],
) -> Result<budget::Usage, Error> {
    let reports = [
        inspected.first_report,
        inspected.second_report,
        output.report(),
    ];
    let mut usage = budget::Usage {
        retained_elements: usize_u64(cells.len().checked_add(headers.len()).ok_or(
            Error::InvalidSource {
                path: Path::Package,
            },
        )?),
        retained_bytes: usize_u64(
            cells
                .len()
                .checked_mul(size_of::<sparse::Cell>())
                .and_then(|bytes| {
                    headers
                        .len()
                        .checked_mul(size_of::<sparse::HeaderBucketSource<'_>>())
                        .and_then(|header_bytes| bytes.checked_add(header_bytes))
                })
                .ok_or(Error::InvalidSource {
                    path: Path::Package,
                })?,
        ),
        allocation_events: 2,
        ..budget::Usage::default()
    };
    for report in reports {
        usage.wire_bytes = checked_add(usage.wire_bytes, usize_u64(report.input_bytes()))?;
        usage.wire_fields = checked_add(usage.wire_fields, usize_u64(report.fields()))?;
        usage.wire_work = checked_add(usage.wire_work, usize_u64(report.work_bytes()))?;
        usage.references = checked_add(usage.references, usize_u64(report.references_scanned()))?;
        usage.retained_bytes =
            checked_add(usage.retained_bytes, usize_u64(report.retained_bytes()))?;
        usage.peak_scratch_bytes = usage
            .peak_scratch_bytes
            .max(usize_u64(report.scratch_bytes()));
        usage.allocation_events =
            checked_add(usage.allocation_events, usize_u64(report.allocations()))?;
        usage.transaction_work =
            checked_add(usage.transaction_work, usize_u64(report.work_bytes()))?;
    }
    Ok(usage)
}

pub(super) fn package_metadata_plan_usage(report: RewriteReport) -> Result<budget::Usage, Error> {
    Ok(budget::Usage {
        wire_bytes: usize_u64(report.input_bytes()),
        wire_fields: usize_u64(report.fields()),
        wire_work: usize_u64(report.work_bytes()),
        references: usize_u64(report.references_scanned()),
        retained_bytes: usize_u64(report.retained_bytes()),
        peak_scratch_bytes: usize_u64(report.scratch_bytes()),
        allocation_events: usize_u64(report.allocations()),
        transaction_work: usize_u64(report.work_bytes()),
        ..budget::Usage::default()
    })
}

pub(super) fn package_metadata_collect_usage(
    report: RewriteReport,
    requirements: metadata::InspectionRequirements,
    identifiers: usize,
    source_bytes: usize,
    path: Path,
) -> Result<budget::Usage, Error> {
    let identifier_bytes = identifiers
        .checked_mul(
            size_of::<u64>()
                .checked_add(size_of::<UuidBits>())
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    // Every UUID attempt performs both one mix and one binary search. New
    // lower identifiers are unique, so across all additions the number of
    // exact-pair collisions is bounded by the existing UUID population.
    let attempt_work = binary_search_work(requirements.uuids)
        .checked_add(4)
        .ok_or(Error::InvalidSource { path })?;
    let search_work = identifiers
        .checked_add(requirements.uuids)
        .and_then(|attempts| attempts.checked_mul(attempt_work))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        wire_bytes: usize_u64(report.input_bytes()),
        wire_fields: usize_u64(report.fields()),
        wire_work: usize_u64(report.work_bytes()),
        references: usize_u64(report.references_scanned()),
        retained_elements: usize_u64(requirements.retained_elements),
        retained_bytes: usize_u64(requirements.retained_bytes),
        peak_scratch_bytes: usize_u64(identifier_bytes.max(report.scratch_bytes())),
        allocation_events: usize_u64(requirements.allocation_events)
            .checked_add(u64::from(identifiers != 0).saturating_mul(2))
            .ok_or(Error::InvalidSource { path })?,
        transaction_work: usize_u64(
            requirements
                .work
                .checked_add(source_bytes)
                .and_then(|work| work.checked_add(report.work_bytes()))
                .and_then(|work| work.checked_add(search_work))
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    })
}

fn binary_search_work(elements: usize) -> usize {
    if elements < 2 {
        1
    } else {
        usize::try_from(usize::BITS - (elements - 1).leading_zeros()).unwrap_or(usize::MAX)
    }
}

pub(super) fn package_metadata_execution_usage(
    execution: RewriteExecutionRequirements,
) -> Result<budget::Usage, Error> {
    Ok(budget::Usage {
        wire_fields: usize_u64(execution.fields()),
        wire_work: usize_u64(execution.work_bytes()),
        output_bytes: usize_u64(execution.output_bytes()),
        references: usize_u64(execution.references()),
        retained_elements: usize_u64(execution.output_bytes()),
        retained_bytes: usize_u64(execution.retained_bytes()),
        peak_scratch_bytes: usize_u64(execution.scratch_bytes()),
        allocation_events: usize_u64(execution.allocations()),
        transaction_work: usize_u64(execution.work_bytes()),
        ..budget::Usage::default()
    })
}

pub(super) fn package_metadata_execution_report_usage(
    report: RewriteReport,
) -> Result<budget::Usage, Error> {
    Ok(budget::Usage {
        wire_fields: usize_u64(report.fields()),
        wire_work: usize_u64(report.work_bytes()),
        output_bytes: usize_u64(report.output_bytes()),
        references: usize_u64(report.references_scanned()),
        retained_elements: usize_u64(report.output_bytes()),
        retained_bytes: usize_u64(report.retained_bytes()),
        peak_scratch_bytes: usize_u64(report.scratch_bytes()),
        allocation_events: usize_u64(report.allocations()),
        transaction_work: usize_u64(report.work_bytes()),
        ..budget::Usage::default()
    })
}

pub(super) fn formula_dependency_prepublication_usage(
    source: &Package,
    replacement_messages: usize,
    new_tiles: usize,
    reference_routes: usize,
    reference_identifiers: usize,
    path: Path,
) -> Result<budget::Usage, Error> {
    let capacity = replacement_messages
        .checked_add(new_tiles)
        .and_then(|value| value.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    let message_count = replacement_messages
        .checked_add(new_tiles)
        .ok_or(Error::InvalidSource { path })?;
    let preview_bytes = rewrite::ROOT_PREVIEWS
        .len()
        .checked_mul(size_of::<&'static str>())
        .ok_or(Error::InvalidSource { path })?;
    let staging_bytes = capacity
        .checked_mul(size_of::<rewrite::ComponentEdit>())
        .and_then(|bytes| {
            replacement_messages
                .checked_mul(size_of::<rewrite::MessageEdit>())
                .and_then(|messages| bytes.checked_add(messages))
        })
        .and_then(|bytes| {
            new_tiles
                .checked_mul(
                    size_of::<ArchiveObject>()
                        .checked_add(size_of::<RawMessage>())?
                        .checked_add(size_of::<MessageInfo>())?
                        .checked_add(3usize.checked_mul(size_of::<u32>())?)?,
                )
                .and_then(|objects| bytes.checked_add(objects))
        })
        .and_then(|bytes| {
            reference_identifiers
                .checked_mul(size_of::<u64>())
                .and_then(|references| bytes.checked_add(references))
        })
        .and_then(|bytes| bytes.checked_add(preview_bytes))
        .ok_or(Error::InvalidSource { path })?;
    let staging_elements = capacity
        .checked_add(replacement_messages)
        .and_then(|elements| elements.checked_add(new_tiles.checked_mul(6)?))
        .and_then(|elements| elements.checked_add(reference_identifiers))
        .and_then(|elements| elements.checked_add(rewrite::ROOT_PREVIEWS.len()))
        .ok_or(Error::InvalidSource { path })?;
    let staging_allocations = u64::from(capacity != 0)
        .checked_add(usize_u64(replacement_messages))
        .and_then(|events| events.checked_add(usize_u64(new_tiles).checked_mul(5)?))
        .and_then(|events| events.checked_add(1))
        .ok_or(Error::InvalidSource { path })?;
    let package_entries = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource { path })?
        .package()
        .len();
    let mut usage = budget::Usage {
        retained_elements: usize_u64(staging_elements),
        retained_bytes: usize_u64(staging_bytes),
        peak_scratch_bytes: usize_u64(staging_bytes),
        allocation_events: staging_allocations,
        objects: usize_u64(new_tiles),
        references: usize_u64(reference_identifiers),
        transaction_work: usize_u64(
            message_count
                .checked_mul(8)
                .and_then(|work| {
                    package_entries
                        .checked_mul(rewrite::ROOT_PREVIEWS.len().checked_mul(2)?)
                        .and_then(|preview_work| work.checked_add(preview_work))
                })
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    };
    merge_usage(
        &mut usage,
        patch_evidence_usage(
            message_count,
            EvidenceReferenceShape {
                routes: reference_routes,
                fields: 0,
                identifiers: reference_identifiers,
            },
            path,
        )?,
        path,
    )?;
    merge_usage(&mut usage, locality_plan_usage(message_count, path)?, path)?;
    Ok(usage)
}

pub(super) fn sparse_usage(
    report: sparse::SparseReport,
    path: Path,
) -> Result<budget::Usage, Error> {
    let wire_bytes = report
        .input_bytes
        .checked_add(report.output_bytes)
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocation_events),
        wire_bytes: usize_u64(wire_bytes),
        wire_fields: usize_u64(report.fields),
        wire_work: usize_u64(report.work),
        objects: usize_u64(report.objects),
        references: usize_u64(report.references),
        header_reads: usize_u64(report.header_reads),
        header_writes: usize_u64(report.header_writes),
        transaction_work: usize_u64(report.work),
        ..budget::Usage::default()
    })
}

fn locality_plan_usage(messages: usize, path: Path) -> Result<budget::Usage, Error> {
    let bytes = messages
        .checked_mul(size_of::<super::locality::DirectionalMessage<'_>>())
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        peak_scratch_bytes: usize_u64(bytes),
        allocation_events: u64::from(messages != 0),
        transaction_work: usize_u64(messages),
        ..budget::Usage::default()
    })
}

fn evidence_message_count(edits: &[rewrite::ComponentEdit], path: Path) -> Result<usize, Error> {
    edits.iter().try_fold(0usize, |total, edit| {
        let new_messages = edit
            .new_objects
            .iter()
            .try_fold(0usize, |messages, object| {
                messages
                    .checked_add(object.messages.len())
                    .ok_or(Error::InvalidSource { path })
            })?;
        total
            .checked_add(edit.messages.len())
            .and_then(|value| value.checked_add(new_messages))
            .ok_or(Error::InvalidSource { path })
    })
}

#[derive(Clone, Copy)]
struct EvidenceReferenceShape {
    routes: usize,
    fields: usize,
    identifiers: usize,
}

fn patch_evidence_usage(
    messages: usize,
    references: EvidenceReferenceShape,
    path: Path,
) -> Result<budget::Usage, Error> {
    let mut usage = budget::arc_vec_retained_usage::<DirectionalMessage>(messages, messages)?;
    if references.routes != 0 {
        merge_usage(
            &mut usage,
            budget::arc_vec_retained_usage::<MessageReferenceRoute>(
                references.routes,
                references.routes,
            )?,
            path,
        )?;
        merge_usage(
            &mut usage,
            budget::arc_vec_retained_usage::<FieldReferenceRoute>(
                references.fields,
                references.fields,
            )?,
            path,
        )?;
        merge_usage(
            &mut usage,
            budget::arc_vec_retained_usage::<u64>(references.identifiers, references.identifiers)?,
            path,
        )?;
        usage.references = usize_u64(references.identifiers);
    }
    // Account for the allocation-free shape pass, evidence construction,
    // `ReferenceEvidence` validation, allocation-shape validation, and the
    // four complete message passes in `PatchEvidence::new`. Field matching is
    // a linear merge below, so this is a hard work bound.
    let work = messages
        .checked_mul(8)
        .and_then(|work| {
            references
                .routes
                .checked_mul(5)
                .and_then(|value| work.checked_add(value))
        })
        .and_then(|work| {
            references
                .fields
                .checked_mul(4)
                .and_then(|value| work.checked_add(value))
        })
        .and_then(|work| {
            references
                .identifiers
                .checked_mul(4)
                .and_then(|value| work.checked_add(value))
        })
        .and_then(|work| work.checked_add(3))
        .ok_or(Error::InvalidSource { path })?;
    usage.transaction_work = usize_u64(work);
    Ok(usage)
}

fn merge_tile_usage(
    usage: &mut budget::Usage,
    report: tile::TileReport,
    path: Path,
) -> Result<(), Error> {
    merge_usage(
        usage,
        budget::Usage {
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
                .ok_or(Error::InvalidSource { path })?,
            ..budget::Usage::default()
        },
        path,
    )
}

fn component_usage(cost: rewrite::ComponentCost, _path: Path) -> Result<budget::Usage, Error> {
    let wire_work = [
        cost.compressed_input_bytes,
        cost.decoded_input_bytes,
        cost.serialized_output_bytes,
        cost.compressed_output_bytes,
    ]
    .into_iter()
    .try_fold(0u64, checked_add)?;
    let objects = checked_add(cost.appended_objects, cost.deleted_objects)?;
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

fn component_reservation_usage(
    reservation: rewrite::ComponentReservation,
    path: Path,
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

fn preflight_usage(
    usage: budget::Usage,
    remaining: budget::Remaining,
    path: Path,
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
    require!(retained_elements, LimitKind::RetainedElements);
    require!(retained_bytes, LimitKind::RetainedBytes);
    require!(peak_scratch_bytes, LimitKind::PeakScratchBytes);
    require!(allocation_events, LimitKind::TransactionWork);
    require!(wire_bytes, LimitKind::WireWork);
    require!(wire_fields, LimitKind::WireFields);
    require!(wire_work, LimitKind::WireWork);
    require!(objects, LimitKind::Objects);
    require!(references, LimitKind::References);
    require!(component_encodes, LimitKind::Objects);
    require!(transaction_work, LimitKind::TransactionWork);
    Ok(())
}

fn merge_usage(base: &mut budget::Usage, delta: budget::Usage, path: Path) -> Result<(), Error> {
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
    add!(tile_reads);
    add!(tile_writes);
    add!(header_reads);
    add!(header_writes);
    add!(row_reads);
    add!(row_writes);
    add!(component_encodes);
    add!(transaction_work);
    base.peak_scratch_bytes = base.peak_scratch_bytes.max(delta.peak_scratch_bytes);
    Ok(())
}

fn checked_add(left: u64, right: u64) -> Result<u64, Error> {
    left.checked_add(right).ok_or(Error::InvalidSource {
        path: Path::Package,
    })
}

pub(super) fn map_sparse_error(error: sparse::SparseError, path: Path) -> Error {
    match error {
        sparse::SparseError::LimitExceeded { observed, maximum } => Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
            path,
        },
        sparse::SparseError::Allocation { requested } => Error::Allocation {
            kind: LimitKind::RetainedBytes,
            amount: requested,
        },
        sparse::SparseError::InvalidAssignments
        | sparse::SparseError::InconsistentSource
        | sparse::SparseError::InvalidSource
        | sparse::SparseError::AmbiguousSource
        | sparse::SparseError::UnsortedCells
        | sparse::SparseError::ZeroTileSize
        | sparse::SparseError::Overflow => Error::InvalidSource { path },
    }
}

pub(super) fn map_metadata_error(
    error: litchi_iwa_protos::package_metadata_codec::RewriteError,
    path: Path,
) -> Error {
    if let Some(amount) = error.allocation_request() {
        Error::Allocation {
            kind: LimitKind::RetainedBytes,
            amount,
        }
    } else if error.resource_limit().is_some() {
        Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
            path,
        }
    } else {
        Error::InvalidSource { path }
    }
}

fn component_edit_mut(
    edits: &mut Vec<rewrite::ComponentEdit>,
    component_index: usize,
    capacity: usize,
    path: Path,
) -> Result<&mut rewrite::ComponentEdit, Error> {
    let index = match edits.binary_search_by_key(&component_index, |edit| edit.component_index) {
        Ok(index) => index,
        Err(index) => {
            if edits.len() == edits.capacity() {
                return Err(Error::Allocation {
                    kind: LimitKind::RetainedElements,
                    amount: capacity,
                });
            }
            edits.insert(
                index,
                rewrite::ComponentEdit {
                    component_index,
                    messages: Vec::new(),
                    object_deletions: Vec::new(),
                    new_objects: Vec::new(),
                },
            );
            index
        },
    };
    edits.get_mut(index).ok_or(Error::InvalidSource { path })
}

fn push_message(
    edits: &mut Vec<rewrite::ComponentEdit>,
    route: resolve::MessageRoute,
    payload: Vec<u8>,
    references: Option<rewrite::ReferenceDelta>,
    capacity: usize,
    path: Path,
) -> Result<(), Error> {
    let edit = component_edit_mut(edits, route.component_index, capacity, path)?;
    edit.messages
        .try_reserve(1)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: edit.messages.len().saturating_add(1),
        })?;
    edit.messages.push(rewrite::MessageEdit {
        object_index: route.object_index,
        message_index: route.message_index,
        expected_type: route.message_type,
        payload,
        references,
    });
    Ok(())
}

fn push_object(
    edits: &mut Vec<rewrite::ComponentEdit>,
    component_index: usize,
    object: ArchiveObject,
    capacity: usize,
    path: Path,
) -> Result<(), Error> {
    let edit = component_edit_mut(edits, component_index, capacity, path)?;
    edit.new_objects
        .try_reserve(1)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::Objects,
            amount: edit.new_objects.len().saturating_add(1),
        })?;
    edit.new_objects.push(object);
    Ok(())
}

pub(super) fn one_message_object(
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<ArchiveObject, Error> {
    let mut messages = Vec::new();
    messages
        .try_reserve_exact(1)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::Objects,
            amount: 1,
        })?;
    messages.push(RawMessage {
        type_: message_type,
        data: payload,
    });
    ArchiveObject::new_with_limits(identifier, messages, limits)
        .map_err(|_error| Error::InvalidSource { path })
}

fn assignment_for_tile(
    assignments: &[sparse::ObjectAssignment],
    tile_id: u32,
    path: Path,
) -> Result<sparse::ObjectAssignment, Error> {
    let mut found = None;
    for &assignment in assignments {
        if assignment.kind == (sparse::NewObjectKind::Tile { tile_id }) {
            if found.replace(assignment).is_some() {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    found.ok_or(Error::InvalidSource { path })
}

fn normalize_edits(edits: &mut [rewrite::ComponentEdit], path: Path) -> Result<(), Error> {
    if edits
        .windows(2)
        .any(|pair| pair[0].component_index >= pair[1].component_index)
    {
        return Err(Error::InvalidSource { path });
    }
    for edit in edits {
        edit.messages
            .sort_unstable_by_key(|message| (message.object_index, message.message_index));
        if edit.messages.windows(2).any(|pair| {
            (pair[0].object_index, pair[0].message_index)
                >= (pair[1].object_index, pair[1].message_index)
        }) {
            return Err(Error::InvalidSource { path });
        }
        edit.new_objects
            .sort_unstable_by_key(|object| object.archive_info.identifier);
        if edit
            .new_objects
            .iter()
            .any(|object| object.archive_info.identifier.is_none())
            || edit
                .new_objects
                .windows(2)
                .any(|pair| pair[0].archive_info.identifier >= pair[1].archive_info.identifier)
        {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(())
}

fn model_reference_delta(
    source: &Package,
    target: &resolve::Target,
    assignments: &[sparse::ObjectAssignment],
    path: Path,
) -> Result<rewrite::ReferenceDelta, Error> {
    let route = target.storage.model;
    let info = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.archive_info.message_infos.get(route.message_index))
        .filter(|info| info.type_ == route.message_type)
        .ok_or(Error::InvalidSource { path })?;
    let aggregate_before = copy_u64s(&info.object_references)?;
    let mut aggregate_after = copy_u64s(&aggregate_before)?;
    aggregate_after
        .try_reserve_exact(assignments.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: assignments.len(),
        })?;
    for assignment in assignments {
        if assignment.object_id == 0 || aggregate_after.contains(&assignment.object_id) {
            return Err(Error::InvalidSource { path });
        }
        aggregate_after.push(assignment.object_id);
    }
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(2)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: 2,
        })?;
    append_field_delta(
        &mut fields,
        info,
        &[4, 3, 1, 2],
        assignments
            .iter()
            .filter_map(|assignment| match assignment.kind {
                sparse::NewObjectKind::Tile { .. } => Some(assignment.object_id),
                sparse::NewObjectKind::RowHeaderBucket { .. } => None,
            }),
        path,
    )?;
    append_field_delta(
        &mut fields,
        info,
        &[4, 1, 2],
        assignments
            .iter()
            .filter_map(|assignment| match assignment.kind {
                sparse::NewObjectKind::RowHeaderBucket { .. } => Some(assignment.object_id),
                sparse::NewObjectKind::Tile { .. } => None,
            }),
        path,
    )?;
    fields.sort_unstable_by_key(|field| field.field_info_index);
    Ok(rewrite::ReferenceDelta {
        aggregate_before,
        aggregate_after,
        fields,
    })
}

fn append_field_delta(
    output: &mut Vec<rewrite::FieldReferenceDelta>,
    info: &MessageInfo,
    expected_path: &[u32],
    identifiers: impl Iterator<Item = u64>,
    path: Path,
) -> Result<(), Error> {
    let mut identifiers = identifiers.peekable();
    if identifiers.peek().is_none() {
        return Ok(());
    }
    let mut found = None;
    for (index, field) in info.field_infos.iter().enumerate() {
        if field.path.path == expected_path {
            if found.replace((index, field)).is_some() {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    let Some((field_info_index, field)) = found else {
        return Ok(());
    };
    let before = copy_u64s(&field.object_references)?;
    let mut after = copy_u64s(&before)?;
    for identifier in identifiers {
        if after.contains(&identifier) {
            return Err(Error::InvalidSource { path });
        }
        after.try_reserve(1).map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: after.len().saturating_add(1),
        })?;
        after.push(identifier);
    }
    let mut path_copy = Vec::new();
    path_copy
        .try_reserve_exact(expected_path.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: expected_path.len(),
        })?;
    path_copy.extend_from_slice(expected_path);
    output.push(rewrite::FieldReferenceDelta {
        field_info_index,
        expected_path: path_copy,
        before,
        after,
    });
    Ok(())
}

fn copy_u64s(source: &[u64]) -> Result<Vec<u64>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: source.len(),
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn evidence_reference_shape(
    source: &Package,
    edits: &[rewrite::ComponentEdit],
    max_work: u64,
    path: Path,
) -> Result<EvidenceReferenceShape, Error> {
    struct ShapeWork {
        used: u64,
        maximum: u64,
        path: Path,
    }

    impl ShapeWork {
        fn consume(&mut self, amount: usize) -> Result<(), Error> {
            if amount == 0 {
                return Ok(());
            }
            let amount = usize_u64(amount);
            let observed = self.used.checked_add(amount).ok_or(Error::LimitExceeded {
                kind: LimitKind::TransactionWork,
                observed: u64::MAX,
                maximum: self.maximum,
                path: self.path,
            })?;
            if observed > self.maximum {
                return Err(Error::LimitExceeded {
                    kind: LimitKind::TransactionWork,
                    observed,
                    maximum: self.maximum,
                    path: self.path,
                });
            }
            self.used = observed;
            #[cfg(test)]
            testing::record_shape_visit();
            Ok(())
        }
    }

    let mut work = ShapeWork {
        used: 0,
        maximum: max_work,
        path,
    };
    let mut shape = EvidenceReferenceShape {
        routes: 0,
        fields: 0,
        identifiers: 0,
    };
    for edit in edits {
        work.consume(1)?;
        let archive = source
            .state
            .components
            .catalog()
            .get_index(edit.component_index)
            .map(|component| component.archive())
            .ok_or(Error::InvalidSource { path })?;
        for message in &edit.messages {
            work.consume(1)?;
            let Some(delta) = &message.references else {
                continue;
            };
            let info = archive
                .objects
                .get(message.object_index)
                .and_then(|object| object.archive_info.message_infos.get(message.message_index))
                .ok_or(Error::InvalidSource { path })?;
            let route_work = delta
                .fields
                .len()
                .checked_add(delta.aggregate_before.len())
                .and_then(|amount| amount.checked_add(info.field_infos.len()))
                .ok_or(Error::InvalidSource { path })?;
            work.consume(route_work)?;
            if delta.aggregate_before != info.object_references
                || delta
                    .fields
                    .windows(2)
                    .any(|pair| pair[0].field_info_index >= pair[1].field_info_index)
            {
                return Err(Error::InvalidSource { path });
            }
            shape.routes = shape
                .routes
                .checked_add(1)
                .ok_or(Error::InvalidSource { path })?;
            shape.fields = shape
                .fields
                .checked_add(info.field_infos.len())
                .ok_or(Error::InvalidSource { path })?;
            shape.identifiers = shape
                .identifiers
                .checked_add(delta.aggregate_before.len())
                .and_then(|count| count.checked_add(delta.aggregate_after.len()))
                .ok_or(Error::InvalidSource { path })?;
            let mut changed_index = 0usize;
            for (field_index, field) in info.field_infos.iter().enumerate() {
                let changed = delta
                    .fields
                    .get(changed_index)
                    .filter(|candidate| candidate.field_info_index == field_index);
                let changed_work = match changed {
                    Some(change) => change
                        .expected_path
                        .len()
                        .checked_add(change.before.len())
                        .ok_or(Error::InvalidSource { path })?,
                    None => 0,
                };
                let comparison_work = field
                    .object_references
                    .len()
                    .checked_add(changed_work)
                    .ok_or(Error::InvalidSource { path })?;
                work.consume(comparison_work)?;
                if changed.is_some() {
                    changed_index = changed_index
                        .checked_add(1)
                        .ok_or(Error::InvalidSource { path })?;
                }
                let target_len = match changed {
                    Some(changed)
                        if changed.expected_path == field.path.path
                            && changed.before == field.object_references =>
                    {
                        changed.after.len()
                    },
                    Some(_) => return Err(Error::InvalidSource { path }),
                    None => field.object_references.len(),
                };
                shape.identifiers = shape
                    .identifiers
                    .checked_add(field.object_references.len())
                    .and_then(|count| count.checked_add(target_len))
                    .ok_or(Error::InvalidSource { path })?;
            }
            if changed_index != delta.fields.len() {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    #[cfg(test)]
    testing::record_shape_requirement(work.used);
    Ok(shape)
}

fn prebuild_evidence(
    source: &Package,
    edits: &[rewrite::ComponentEdit],
    shape: EvidenceReferenceShape,
    path: Path,
) -> Result<(Vec<DirectionalMessage>, Option<ReferenceEvidence>), Error> {
    let count = edits.iter().try_fold(0usize, |total, edit| {
        let appended = edit
            .new_objects
            .iter()
            .try_fold(0usize, |messages, object| {
                messages.checked_add(object.messages.len())
            })?;
        total
            .checked_add(edit.messages.len())?
            .checked_add(appended)
    });
    let count = count.ok_or(Error::InvalidSource { path })?;
    let mut evidence = Vec::new();
    evidence
        .try_reserve_exact(count)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::RetainedElements,
            amount: count,
        })?;
    let mut routes = Vec::new();
    routes
        .try_reserve_exact(shape.routes)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: shape.routes,
        })?;
    let mut fields = Vec::new();
    fields
        .try_reserve_exact(shape.fields)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: shape.fields,
        })?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(shape.identifiers)
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::References,
            amount: shape.identifiers,
        })?;
    for edit in edits {
        let archive = source
            .state
            .components
            .catalog()
            .get_index(edit.component_index)
            .map(|component| component.archive())
            .ok_or(Error::InvalidSource { path })?;
        for message in &edit.messages {
            let object = archive
                .objects
                .get(message.object_index)
                .filter(|object| {
                    object
                        .messages
                        .get(message.message_index)
                        .is_some_and(|source_message| source_message.type_ == message.expected_type)
                })
                .ok_or(Error::InvalidSource { path })?;
            let object_identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or(Error::InvalidSource { path })?;
            let location = PhysicalLocation {
                component: edit.component_index,
                object: message.object_index,
                message: message.message_index,
            };
            let mut evidence_message = DirectionalMessage::new(
                Some(location),
                Some(location),
                object_identifier,
                message.expected_type,
                EvidenceChangeKind::Replace,
            );
            if let Some(delta) = &message.references {
                let info = object
                    .archive_info
                    .message_infos
                    .get(message.message_index)
                    .ok_or(Error::InvalidSource { path })?;
                let aggregate_source =
                    ReferenceSpan::new(identifiers.len(), delta.aggregate_before.len());
                identifiers.extend_from_slice(&delta.aggregate_before);
                let aggregate_target =
                    ReferenceSpan::new(identifiers.len(), delta.aggregate_after.len());
                identifiers.extend_from_slice(&delta.aggregate_after);
                let field_start = fields.len();
                let mut changed_index = 0usize;
                for (field_index, field) in info.field_infos.iter().enumerate() {
                    let changed = delta
                        .fields
                        .get(changed_index)
                        .filter(|candidate| candidate.field_info_index == field_index);
                    if changed.is_some() {
                        changed_index = changed_index
                            .checked_add(1)
                            .ok_or(Error::InvalidSource { path })?;
                    }
                    let target = changed.map_or(field.object_references.as_slice(), |changed| {
                        changed.after.as_slice()
                    });
                    let source_span =
                        ReferenceSpan::new(identifiers.len(), field.object_references.len());
                    identifiers.extend_from_slice(&field.object_references);
                    let target_span = ReferenceSpan::new(identifiers.len(), target.len());
                    identifiers.extend_from_slice(target);
                    fields.push(FieldReferenceRoute::new(
                        field_index,
                        source_span,
                        target_span,
                    ));
                }
                if changed_index != delta.fields.len() {
                    return Err(Error::InvalidSource { path });
                }
                let route = routes.len();
                routes.push(MessageReferenceRoute::new(
                    aggregate_source,
                    aggregate_target,
                    ReferenceSpan::new(field_start, info.field_infos.len()),
                ));
                evidence_message = evidence_message.with_reference_transition(route);
            }
            evidence.push(evidence_message);
        }
        let base = archive.objects.len();
        for (object_offset, object) in edit.new_objects.iter().enumerate() {
            let object_index = base
                .checked_add(object_offset)
                .ok_or(Error::InvalidSource { path })?;
            let object_identifier = object
                .archive_info
                .identifier
                .filter(|identifier| *identifier != 0)
                .ok_or(Error::InvalidSource { path })?;
            for (message_index, message) in object.messages.iter().enumerate() {
                evidence.push(DirectionalMessage::new(
                    None,
                    Some(PhysicalLocation {
                        component: edit.component_index,
                        object: object_index,
                        message: message_index,
                    }),
                    object_identifier,
                    message.type_,
                    EvidenceChangeKind::Append,
                ));
            }
        }
    }
    if evidence.len() != count {
        return Err(Error::InvalidSource { path });
    }
    if routes.len() != shape.routes
        || fields.len() != shape.fields
        || identifiers.len() != shape.identifiers
    {
        return Err(Error::InvalidSource { path });
    }
    let references = if routes.is_empty() {
        None
    } else {
        let evidence = ReferenceEvidence::new(
            Arc::new(routes),
            Arc::new(fields),
            Arc::new(identifiers),
            path,
        )?;
        if evidence.allocation_shapes()
            != (
                (shape.routes, shape.routes),
                (shape.fields, shape.fields),
                (shape.identifiers, shape.identifiers),
            )
        {
            return Err(Error::InvalidSource { path });
        }
        Some(evidence)
    };
    Ok((evidence, references))
}
