//! Bounded preparation for the narrow existing-tile rich-text scalar edge.
//!
//! The grouped parent owns string assignment, final tile rewriting, staged
//! publication, inverse evidence, and semantic replay. This helper performs
//! the read-only tile preclassification needed before string assignment and
//! admits exactly one uniquely-owned rich-text cell whose user-entered text
//! can replace its existing `StorageArchive` in place.

use core::mem::size_of;

use litchi_iwa_common::WireLimits;
use litchi_iwa_text_wire::RewriteLimits;

use crate::{
    Package,
    package::table_cells::{DependencyKind, Error, LimitKind, Path},
    table::cells::{Change, Input},
};

use super::{authorize_remaining, budget, map_tile_error, message_payload, resolve, rich, tile};

/// One existing storage-message payload ready for grouped publication.
#[derive(Debug)]
pub(super) struct PreparedRichReplacement {
    pub(super) route: resolve::MessageRoute,
    pub(super) payload: Vec<u8>,
    /// Exact archive-header transition; no object deletion is implied.
    pub(super) references: rich::ReferenceDelta,
}

pub(super) struct PreparedRichTransition<'source, 'replacement> {
    pub(super) route: resolve::MessageRoute,
    expected_type: u32,
    plan: rich::PreparedPlan<'source, 'replacement>,
    publication_reference_items: usize,
}

impl PreparedRichTransition<'_, '_> {
    pub(super) const fn execution_requirements(&self) -> rich::ExecutionRequirements {
        self.plan.execution_requirements()
    }

    pub(super) const fn publication_reference_items(&self) -> usize {
        self.publication_reference_items
    }

    pub(super) const fn route(&self) -> resolve::MessageRoute {
        self.route
    }

    pub(super) fn retained_accounting(
        &self,
        budget: &budget::TransactionBudget,
        path: Path,
    ) -> Result<rich::RetainedAccounting, Error> {
        self.plan
            .retained_accounting()
            .map_err(|error| map_rich_error(error, budget, path))
    }

    pub(super) fn execute(
        self,
        budget: &budget::TransactionBudget,
        path: Path,
    ) -> Result<(PreparedRichReplacement, rich::ExecutionReport), Error> {
        let requirements = self.plan.execution_requirements();
        let plan = self
            .plan
            .execute(requirements)
            .map_err(|error| map_rich_error(error, budget, path))?;
        let execution = plan.execution_report();
        let rich::PlanParts {
            disposition,
            replacements: staged,
        } = plan.into_parts();
        if disposition != rich::Disposition::InPlace || staged.len() != 1 {
            return Err(Error::InvalidSource { path });
        }
        let replacement = staged
            .into_iter()
            .next()
            .ok_or(Error::InvalidSource { path })?;
        if replacement.location != rich_location(self.route)
            || replacement.expected_type != self.expected_type
            || replacement.kind != rich::ReplacementKind::StorageArchive
            || !replacement.references.removed_by_field.is_empty()
        {
            return Err(Error::InvalidSource { path });
        }
        Ok((
            PreparedRichReplacement {
                route: self.route,
                payload: replacement.payload,
                references: replacement.references,
            },
            execution,
        ))
    }
}

/// Rich keys aligned one-for-one with the caller's changed-cell slice.
///
/// `Some(key)` occurs only for the admitted `Input::Text` cell which was
/// preclassified as rich text. Plain text cells remain `None` for the string
/// planner, and non-text inputs are not inspected here. Exact-equality rich
/// edits retain their key without staging a storage replacement.
pub(super) struct PreparedRich<'source, 'replacement> {
    pub(super) keys: Vec<Option<u32>>,
    pub(super) transition: Option<PreparedRichTransition<'source, 'replacement>>,
    /// Rich storage messages whose reference transition is owned here.
    pub(super) owned_transition_count: usize,
}

/// Prepare one uniquely-owned RichText-to-user-text transition.
///
/// Multiple rich cells, shared ownership, field-specific metadata changes,
/// release, and non-existing tiles retain the public typed `RichText` refusal.
/// Exact aggregate reference changes travel with the staged replacement. No
/// package, archive, tile, or list payload is mutated.
pub(super) fn prepare_unique_rich_text_to_text<'source, 'replacement>(
    source: &'source Package,
    target: &resolve::Target,
    changes: &'replacement [Change],
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<PreparedRich<'source, 'replacement>, Error> {
    if target.path() != path {
        return Err(Error::InvalidSource { path });
    }
    let key_bytes =
        changes
            .len()
            .checked_mul(size_of::<Option<u32>>())
            .ok_or(Error::LimitExceeded {
                kind: LimitKind::RetainedBytes,
                observed: u64::MAX,
                maximum: budget.limits().max_retained_bytes,
                path,
            })?;
    budget.reserve_retained(
        usize_u64(changes.len()),
        usize_u64(key_bytes),
        u64::from(!changes.is_empty()),
    )?;
    let mut keys = Vec::new();
    reserve_exact(&mut keys, changes.len(), LimitKind::RetainedElements)?;
    keys.resize(changes.len(), None);

    let mut selected: Option<(usize, u32)> = None;
    let mut start = 0usize;
    while start < changes.len() {
        let tile_id = changes[start].position().row() / target.storage.tile_size;
        let end = changes[start..]
            .iter()
            .position(|change| change.position().row() / target.storage.tile_size != tile_id)
            .map_or(changes.len(), |offset| start + offset);
        let text_count = changes[start..end]
            .iter()
            .filter(|change| matches!(change.input_ref(), Some(Input::Text(_))))
            .count();
        if text_count != 0 {
            let route = target
                .storage
                .tiles
                .binary_search_by_key(&tile_id, |route| route.tile_id)
                .ok()
                .and_then(|index| target.storage.tiles.get(index))
                .ok_or(Error::UnsupportedDependency {
                    path,
                    kind: DependencyKind::CellStorage,
                })?;
            let scratch_element_bytes = size_of::<tile::TileReadPosition>()
                .checked_add(size_of::<usize>())
                .and_then(|bytes| bytes.checked_add(size_of::<tile::PreclassifiedCell>()))
                .ok_or(Error::InvalidSource { path })?;
            let scratch_bytes =
                text_count
                    .checked_mul(scratch_element_bytes)
                    .ok_or(Error::LimitExceeded {
                        kind: LimitKind::PeakScratchBytes,
                        observed: u64::MAX,
                        maximum: budget.limits().max_scratch_bytes,
                        path,
                    })?;
            let tile_payload = message_payload(source, route.message, path)?;
            budget.with_scratch(usize_u64(scratch_bytes), 2, |budget| -> Result<(), Error> {
                let mut positions = Vec::new();
                reserve_exact(&mut positions, text_count, LimitKind::PeakScratchBytes)?;
                let mut indices = Vec::new();
                reserve_exact(&mut indices, text_count, LimitKind::PeakScratchBytes)?;
                for (offset, change) in changes[start..end].iter().enumerate() {
                    if matches!(change.input_ref(), Some(Input::Text(_))) {
                        positions.push(tile::TileReadPosition {
                            row: change.position().row() % target.storage.tile_size,
                            column: change.position().column(),
                        });
                        indices.push(start + offset);
                    }
                }
                let remaining = authorize_remaining(budget)?;
                let limits = tile_limits(source, target, remaining, path)?;
                let outcome = match tile::preclassify_tile(
                    tile_payload,
                    target.native.columns,
                    &positions,
                    limits,
                ) {
                    Ok(outcome) => outcome,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(map_tile_error(error, path));
                    },
                };
                let usage = match super::tile_usage(budget, outcome.report, 1) {
                    Ok(usage) => usage,
                    Err(error) => {
                        budget.cancel_authorization();
                        return Err(error);
                    },
                };
                if let Err(error) = budget.record_authorized(usage) {
                    if budget.authorization_is_pending() {
                        budget.cancel_authorization();
                    }
                    return Err(error);
                }
                if outcome.cells.len() != indices.len() {
                    return Err(Error::InvalidSource { path });
                }
                for (change_index, cell) in indices.into_iter().zip(outcome.cells) {
                    if let Some(key) = cell.before_references.rich_text {
                        if cell.before != tile::CellValue::RichText(key)
                            || selected.replace((change_index, key)).is_some()
                        {
                            return Err(rich_refusal(path));
                        }
                    }
                }
                Ok(())
            })?;
        }
        start = end;
    }

    let Some((change_index, key)) = selected else {
        return Ok(PreparedRich {
            keys,
            transition: None,
            owned_transition_count: 0,
        });
    };
    let text = match changes.get(change_index).and_then(Change::input_ref) {
        Some(Input::Text(text)) => text.as_str(),
        _ => return Err(Error::InvalidSource { path }),
    };
    let index = target
        .storage
        .rich
        .as_ref()
        .ok_or_else(|| rich_refusal(path))?;
    let entry = index
        .entries
        .binary_search_by_key(&key, |entry| entry.key)
        .ok()
        .and_then(|position| index.entries.get(position))
        .ok_or(Error::InvalidSource { path })?;
    let pair = index
        .pairs
        .get(entry.pair_index)
        .ok_or(Error::InvalidSource { path })?;
    let payload_edges = inbound(&index.payload_list_inbound, pair.payload.object_id, path)?;
    let storage_edges = inbound(&index.storage_payload_inbound, pair.storage.object_id, path)?;
    if entry.ref_count != 1 || payload_edges != 1 || storage_edges != 1 {
        return Err(rich_refusal(path));
    }
    let rich_plan = prepare_unique(
        source,
        index,
        entry,
        pair,
        payload_edges,
        storage_edges,
        text,
        budget,
        path,
    )?;
    if rich_plan.result_key() != key {
        return Err(rich_refusal(path));
    }
    if rich_plan.disposition() == rich::Disposition::Unchanged {
        keys[change_index] = Some(key);
        return Ok(PreparedRich {
            keys,
            transition: None,
            owned_transition_count: 0,
        });
    }
    if rich_plan.disposition() != rich::Disposition::InPlace {
        return Err(rich_refusal(path));
    }
    keys[change_index] = Some(key);
    let requirements = rich_plan.execution_requirements();
    let publication_reference_items = pair
        .storage
        .object_references
        .len()
        .checked_add(
            pair.storage
                .object_references
                .len()
                .max(requirements.reference_occurrences)
                .checked_mul(2)
                .ok_or(Error::InvalidSource { path })?,
        )
        .ok_or(Error::InvalidSource { path })?;
    let transition = PreparedRichTransition {
        route: pair.storage.message,
        expected_type: pair.storage.message_type,
        plan: rich_plan,
        publication_reference_items,
    };
    Ok(PreparedRich {
        keys,
        transition: Some(transition),
        owned_transition_count: 1,
    })
}

fn prepare_unique<'source, 'target, 'replacement>(
    source: &'source Package,
    index: &'target resolve::RichRouteIndex,
    entry: &'target resolve::RichEntryRoute,
    pair: &'target resolve::RichResolvedPairRoute,
    payload_edges: u32,
    storage_edges: u32,
    text: &'replacement str,
    budget: &mut budget::TransactionBudget,
    path: Path,
) -> Result<rich::PreparedPlan<'source, 'replacement>, Error> {
    let field_count = pair
        .payload
        .field_references
        .len()
        .checked_add(pair.storage.field_references.len())
        .ok_or(Error::InvalidSource { path })?;
    let scratch_bytes = field_count
        .checked_mul(size_of::<rich::FieldReferences<'_>>())
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::PeakScratchBytes,
            observed: u64::MAX,
            maximum: budget.limits().max_scratch_bytes,
            path,
        })?;
    budget.with_scratch(
        usize_u64(scratch_bytes),
        2,
        |budget| -> Result<rich::PreparedPlan<'source, 'replacement>, Error> {
            let payload_fields = rich_fields(&pair.payload, path)?;
            let storage_fields = rich_fields(&pair.storage, path)?;
            let request = rich::Request {
                route: rich::ListRoute {
                    root_object_id: entry.root_object_id,
                    owner: match &entry.owner {
                        resolve::RichEntryOwner::Root => rich::EntryOwner::Root,
                        resolve::RichEntryOwner::Segment {
                            object_id,
                            owner_entries,
                            root_references,
                            ..
                        } => rich::EntryOwner::Segment {
                            object_id: *object_id,
                            entries: *owner_entries,
                            root_references: *root_references,
                        },
                    },
                },
                key: entry.key,
                list_ref_count: entry.ref_count,
                payload: rich_object(source, &pair.payload, &payload_fields, path)?,
                storage: rich_object(source, &pair.storage, &storage_fields, path)?,
                payload_inbound_references: payload_edges,
                storage_inbound_references: storage_edges,
                local_object_ids: &index.local_object_ids,
            };
            let remaining = authorize_remaining(budget)?;
            let limits = rich_limits(source, remaining, path)?;
            let plan = match rich::prepare_text(request, text, limits) {
                Ok(plan) => plan,
                Err(error) => {
                    budget.cancel_authorization();
                    return Err(map_rich_error(error, budget, path));
                },
            };
            let retained = plan.retained_accounting().map_err(|error| {
                budget.cancel_authorization();
                map_rich_error(error, budget, path)
            })?;
            let usage = rich_prepare_usage(plan.prepare_report(), retained, path)?;
            budget.record_authorized(usage)?;
            Ok(plan)
        },
    )
}

fn rich_fields<'a>(
    object: &'a resolve::RichObjectRoute,
    path: Path,
) -> Result<Vec<rich::FieldReferences<'a>>, Error> {
    let mut fields = Vec::new();
    reserve_exact(
        &mut fields,
        object.field_references.len(),
        LimitKind::PeakScratchBytes,
    )?;
    for field in &object.field_references {
        fields.push(rich::FieldReferences {
            root_field: *field.path.first().ok_or(Error::InvalidSource { path })?,
            references: &field.object_references,
        });
    }
    Ok(fields)
}

fn rich_object<'source, 'object, 'fields>(
    source: &'source Package,
    object: &'object resolve::RichObjectRoute,
    fields: &'fields [rich::FieldReferences<'object>],
    path: Path,
) -> Result<rich::ObjectSource<'source, 'fields, 'object>, Error> {
    Ok(rich::ObjectSource {
        location: rich_location(object.message),
        identifier: object.object_id,
        message_type: object.message_type,
        payload: message_payload(source, object.message, path)?,
        object_references: &object.object_references,
        field_references: fields,
    })
}

fn rich_location(route: resolve::MessageRoute) -> rich::MessageLocation {
    rich::MessageLocation {
        component_index: route.component_index,
        object_index: route.object_index,
        message_index: route.message_index,
    }
}

fn inbound(counts: &[(u64, u32)], identifier: u64, path: Path) -> Result<u32, Error> {
    counts
        .binary_search_by_key(&identifier, |(candidate, _count)| *candidate)
        .ok()
        .and_then(|index| counts.get(index).map(|(_identifier, count)| *count))
        .filter(|count| *count != 0)
        .ok_or(Error::InvalidSource { path })
}

fn tile_limits(
    source: &Package,
    target: &resolve::Target,
    remaining: budget::Remaining,
    path: Path,
) -> Result<tile::TileLimits, Error> {
    let maximum_wire = source.state.options.archive().max_iwa_stream_bytes();
    let maximum_work = maximum_wire
        .checked_mul(32)
        .ok_or(Error::InvalidSource { path })?;
    Ok(tile::TileLimits::new(
        bounded_usize(remaining.wire_bytes, maximum_wire, path)?,
        bounded_usize(
            remaining.retained_bytes.min(remaining.peak_scratch_bytes),
            maximum_wire,
            path,
        )?,
        bounded_usize(remaining.wire_fields, maximum_wire, path)?,
        budget::tile_work_ceiling(remaining, usize_u64(maximum_work), 0)?,
        usize::try_from(target.storage.tile_size)
            .map_err(|_error| Error::InvalidSource { path })?,
        source.state.options.semantic().max_materialized_cells(),
    ))
}

fn rich_limits(
    source: &Package,
    remaining: budget::Remaining,
    path: Path,
) -> Result<rich::Limits, Error> {
    let maximum_wire = source.state.options.archive().max_iwa_stream_bytes();
    let input = bounded_usize(remaining.wire_bytes, maximum_wire, path)?;
    let fields = bounded_usize(remaining.wire_fields, RewriteLimits::MAX_FIELDS, path)?;
    let output = bounded_usize(
        remaining.retained_bytes.min(remaining.peak_scratch_bytes),
        RewriteLimits::MAX_OUTPUT_BYTES,
        path,
    )?;
    let work = bounded_usize(
        remaining
            .wire_work
            .min(remaining.rich_text_work)
            .min(remaining.transaction_work),
        RewriteLimits::MAX_REWRITE_WORK,
        path,
    )?;
    let references = bounded_usize(
        remaining.references,
        RewriteLimits::MAX_OBJECT_REFERENCES,
        path,
    )?;
    let wire = WireLimits::default()
        .with_input_bytes(input)
        .and_then(|limits| limits.with_fields(fields.min(WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_output_bytes(output.min(WireLimits::MAX_OUTPUT_BYTES)))
        .and_then(|limits| limits.with_nesting(64))
        .and_then(|limits| limits.with_rewrite_work(work.min(WireLimits::MAX_REWRITE_WORK)))
        .map_err(|_error| Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: remaining.wire_work,
            path,
        })?;
    let text = RewriteLimits::new(
        input.min(RewriteLimits::MAX_MESSAGE_BYTES),
        fields,
        64,
        fields.min(RewriteLimits::MAX_FRAGMENTS),
        source
            .state
            .options
            .semantic()
            .max_output_text_bytes()
            .clamp(1, RewriteLimits::MAX_TEXT_BYTES),
        fields.min(RewriteLimits::MAX_TABLE_ENTRIES),
        references,
        output,
        work,
    )
    .map_err(|_error| Error::LimitExceeded {
        kind: LimitKind::WireWork,
        observed: u64::MAX,
        maximum: remaining.wire_work,
        path,
    })?;
    Ok(rich::Limits {
        wire,
        text,
        max_deltas: references,
        max_work: work,
    })
}

fn rich_prepare_usage(
    report: rich::Report,
    retained: rich::RetainedAccounting,
    path: Path,
) -> Result<budget::Usage, Error> {
    let transaction_work = report
        .work_bound
        .checked_add(report.reference_occurrences)
        .and_then(|work| work.checked_add(retained.elements))
        .ok_or(Error::InvalidSource { path })?;
    Ok(budget::Usage {
        retained_elements: usize_u64(retained.elements),
        retained_bytes: usize_u64(retained.bytes),
        allocation_events: usize_u64(retained.allocation_events),
        wire_bytes: usize_u64(report.input_bytes),
        wire_fields: usize_u64(report.wire_fields),
        wire_work: usize_u64(report.work_bound),
        references: usize_u64(report.reference_occurrences),
        rich_text_work: usize_u64(report.work_bound),
        transaction_work: usize_u64(transaction_work),
        ..budget::Usage::default()
    })
}

pub(super) fn rich_execution_usage(
    report: rich::ExecutionReport,
    path: Path,
) -> Result<budget::Usage, Error> {
    Ok(budget::Usage {
        output_bytes: usize_u64(report.output_bytes),
        retained_elements: usize_u64(report.retained_elements),
        retained_bytes: usize_u64(report.retained_bytes),
        peak_scratch_bytes: usize_u64(report.peak_scratch_bytes),
        allocation_events: usize_u64(report.allocation_events),
        references: usize_u64(report.reference_occurrences),
        rich_text_work: usize_u64(report.work_bound),
        wire_work: usize_u64(report.work_bound),
        transaction_work: usize_u64(
            report
                .work_bound
                .checked_add(report.reference_occurrences)
                .ok_or(Error::InvalidSource { path })?,
        ),
        ..budget::Usage::default()
    })
}

pub(super) fn rich_execution_requirements_usage(
    requirements: rich::ExecutionRequirements,
    path: Path,
) -> Result<budget::Usage, Error> {
    rich_execution_usage(
        rich::ExecutionReport {
            output_bytes: requirements.output_bytes,
            retained_elements: requirements.retained_elements,
            retained_bytes: requirements.retained_bytes,
            peak_scratch_bytes: requirements.peak_scratch_bytes,
            allocation_events: requirements.allocation_events,
            work_bound: requirements.work_bound,
            reference_occurrences: requirements.reference_occurrences,
        },
        path,
    )
}

fn bounded_usize(value: u64, maximum: usize, path: Path) -> Result<usize, Error> {
    let value = usize::try_from(value).map_err(|_error| Error::InvalidSource { path })?;
    let bounded = value.min(maximum);
    if bounded == 0 {
        Err(Error::LimitExceeded {
            kind: LimitKind::TransactionWork,
            observed: 1,
            maximum: 0,
            path,
        })
    } else {
        Ok(bounded)
    }
}

fn reserve_exact<T>(output: &mut Vec<T>, amount: usize, kind: LimitKind) -> Result<(), Error> {
    let expected = output
        .len()
        .checked_add(amount)
        .ok_or(Error::Allocation { kind, amount })?;
    output
        .try_reserve_exact(amount)
        .map_err(|_error| Error::Allocation { kind, amount })?;
    if size_of::<T>() != 0 && output.capacity() != expected {
        return Err(Error::Allocation { kind, amount });
    }
    Ok(())
}

fn map_rich_error(error: rich::Error, budget: &budget::TransactionBudget, path: Path) -> Error {
    match error {
        rich::Error::Allocation { amount } => Error::Allocation {
            kind: LimitKind::RetainedBytes,
            amount,
        },
        rich::Error::Limit => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: budget.limits().max_wire_work,
            path,
        },
        rich::Error::InvalidSource | rich::Error::Overflow => Error::InvalidSource { path },
    }
}

fn rich_refusal(path: Path) -> Error {
    Error::UnsupportedDependency {
        path,
        kind: DependencyKind::RichText,
    }
}

const fn usize_u64(value: usize) -> u64 {
    value as u64
}

#[cfg(test)]
mod tests {
    use std::path::{Path, PathBuf};

    use crate::{
        Package,
        package::table_cells::Path as CellPath,
        table::{
            CellPosition,
            cells::{Change, Input},
        },
    };

    use super::{budget, prepare_unique_rich_text_to_text, resolve};

    fn checked_in_fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers")
    }

    #[test]
    fn native_plain_text_preparation_defers_to_string_planner() {
        let source = Package::open(checked_in_fixture()).expect("native Numbers fixture opens");
        let position = CellPosition::from_a1("C2").expect("C2 is canonical");
        let (target, _report) = resolve::resolve_changed_target(&source, 0, 0, &[position])
            .expect("native rich route resolves");
        let mut transaction =
            budget::TransactionBudget::new(&source).expect("native budget is finite");
        let change = Change::set(
            position,
            Input::text("rich replacement").expect("replacement allocation"),
        );
        let changes = [change];
        let prepared = prepare_unique_rich_text_to_text(
            &source,
            &target,
            &changes,
            &mut transaction,
            CellPath::Table { sheet: 0, table: 0 },
        )
        .expect("native text preclassification succeeds");
        assert_eq!(prepared.keys.len(), 1);
        assert_eq!(prepared.keys[0], None);
        assert_eq!(prepared.owned_transition_count, 0);
        assert!(prepared.transition.is_none());
    }

    #[test]
    #[ignore = "requires external Numbers 14.4 oracle"]
    fn frozen_formula_rich_oracle_prepares_exact_aggregate_transition() {
        let oracle = Path::new(
            "/private/tmp/litchi-numbers-cell-batch-native.wuaiMp/oracle-preserved.numbers",
        );
        let source = Package::open(oracle).expect("frozen formula-rich oracle opens");
        let position = CellPosition::from_a1("C2").expect("C2 is canonical");
        let (target, _report) = resolve::resolve_changed_target(&source, 0, 0, &[position])
            .expect("frozen formula-rich route resolves");
        let mut transaction =
            budget::TransactionBudget::new(&source).expect("oracle budget is finite");
        let change = Change::set(
            position,
            Input::text("CELL: Café changed").expect("replacement allocation"),
        );
        let changes = [change];
        let mut prepared = prepare_unique_rich_text_to_text(
            &source,
            &target,
            &changes,
            &mut transaction,
            CellPath::Table { sheet: 0, table: 0 },
        )
        .expect("formula-rich preparation succeeds");
        assert_eq!(prepared.keys, [Some(1)]);
        assert_eq!(prepared.owned_transition_count, 1);
        let (replacement, _report) = prepared
            .transition
            .take()
            .expect("rich transition is deferred")
            .execute(&transaction, CellPath::Table { sheet: 0, table: 0 })
            .expect("rich transition executes");
        assert_eq!(replacement.references.before, [903_835, 903_815, 905_312]);
        assert_eq!(replacement.references.after, [903_835, 903_815]);
        assert_eq!(replacement.references.removed, [905_312]);
        assert!(replacement.references.removed_by_field.is_empty());
    }

    #[test]
    #[ignore = "requires external Numbers 14.4 oracle"]
    fn frozen_formula_rich_oracle_commit_replays_and_inverts_exactly() {
        let oracle = Path::new(
            "/private/tmp/litchi-numbers-cell-batch-native.wuaiMp/oracle-preserved.numbers",
        );
        let source = Package::open(oracle).expect("frozen formula-rich oracle opens");
        let mut source_bytes = Vec::new();
        source
            .write_to(&mut source_bytes)
            .expect("source package serializes");
        let position = CellPosition::from_a1("C2").expect("C2 is canonical");
        let commit = source
            .edit_table_cells(0usize, 0usize)
            .expect("table edit starts")
            .set(
                position,
                Input::text("CELL: Café changed").expect("replacement allocation"),
            )
            .expect("rich edit stages")
            .commit()
            .expect("rich edit commits");
        let mut target_bytes = Vec::new();
        commit
            .package()
            .write_to(&mut target_bytes)
            .expect("target package serializes");
        let replayed = source
            .apply_table_cells(commit.patch())
            .expect("rich patch replays");
        let mut replayed_bytes = Vec::new();
        replayed
            .package()
            .write_to(&mut replayed_bytes)
            .expect("replayed package serializes");
        assert_eq!(replayed_bytes, target_bytes);
        let restored = replayed
            .package()
            .apply_table_cells(&commit.patch().inverse())
            .expect("rich inverse applies");
        let mut restored_bytes = Vec::new();
        restored
            .package()
            .write_to(&mut restored_bytes)
            .expect("restored package serializes");
        assert_eq!(restored_bytes, source_bytes);
    }
}
