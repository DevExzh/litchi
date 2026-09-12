//! CalculationEngine ownership for Pages body-table deletion.
//!
//! This phase is deliberately a preparation-only pass.  It borrows the
//! immutable package source, validates every formula-owner and dependency
//! payload through the focused Buffa-backed codec, then returns source
//! authorized message edits and object-removal records.  The archive phase is
//! the only publisher.

use std::collections::{HashMap, HashSet};
use std::mem;
use std::num::NonZeroU64;

use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency;
use litchi_iwa_protos::numbers_table_cell_dependency_codec::{
    DecodeError, DecodeLimit, DecodeOptions, DecodeReport, DependencyVisitor,
    ExpandedEdgeComponent, ExpandedEdgeKind, FormulaDependencyFact,
    FormulaOwnerDependenciesSnapshot, ReferenceRecord,
};

use super::{
    BodyTableDeletionError, DeletionObjectLocation, FormulaPlan, GraphPlan, MessageEdit,
    ObjectRemoval, RemovalPlan,
};
use crate::package::{Package, table_lock::WireBudget};

const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;
const RANGE_PRECEDENTS_TILE_MESSAGE_TYPE: u32 = 4_010;

/// Prepare all CalculationEngine changes owned by the selected body table.
///
/// The graph phase supplies the table-context roots in
/// `graph.removals.formula_contexts`.  Formula preparation adds the owner
/// family, private dependency tiles, global dependency rewrites, and the
/// CalculationEngine replacement to a fresh removal plan.  No source archive
/// is changed here, and all allocations happen before the plan is returned.
pub(super) fn prepare(
    package: &Package,
    graph: &GraphPlan,
    budget: &mut WireBudget,
) -> Result<FormulaPlan, BodyTableDeletionError> {
    let context_ids = set_from_slice_u64(&graph.removals.formula_contexts)?;
    if context_ids.is_empty() {
        return Ok(FormulaPlan {
            removals: RemovalPlan::default(),
        });
    }

    let components = package.state.source.components();
    let already_removed = object_removal_set(&graph.removals.object_removals)?;
    let Some(engine_component_index) = calculation_engine_component_index(components)? else {
        return Ok(FormulaPlan {
            removals: RemovalPlan::default(),
        });
    };
    let owners = collect_formula_owners(components, engine_component_index, budget)?;
    if owners.is_empty() {
        return Ok(FormulaPlan {
            removals: RemovalPlan::default(),
        });
    }

    let selected = select_owner_family(&owners, &context_ids, &already_removed, budget)?;
    if selected.is_empty() {
        return Ok(FormulaPlan {
            removals: RemovalPlan::default(),
        });
    }

    let (owner_ids, internal_owner_ids, owner_uuids) =
        selected_owner_sets(&owners, &selected, budget)?;
    let owner_id_set = set_from_slice(&owner_ids)?;
    let internal_owner_id_set = set_from_slice(&internal_owner_ids)?;
    let owner_uuid_set = set_from_slice(&owner_uuids)?;

    let engine =
        find_calculation_engine(components, engine_component_index, &owner_id_set, budget)?
            .ok_or(BodyTableDeletionError::InvalidSource)?;

    let mut plan = RemovalPlan::default();
    let mut selected_cell_tiles = HashSet::new();
    let mut selected_range_tiles = HashSet::new();
    let mut surviving_cell_tiles = HashSet::new();
    let mut surviving_range_tiles = HashSet::new();

    for (index, owner) in owners.iter().enumerate() {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
        if already_removed.contains(&owner.identifier) {
            continue;
        }
        let destination = if selected.contains(&index) {
            &mut selected_cell_tiles
        } else {
            &mut surviving_cell_tiles
        };
        insert_all(destination, &owner.facts.cell_tile_ids, budget)?;
        let destination = if selected.contains(&index) {
            &mut selected_range_tiles
        } else {
            &mut surviving_range_tiles
        };
        insert_all(destination, &owner.facts.range_tile_ids, budget)?;
    }

    let private_cell_tiles = difference(&selected_cell_tiles, &surviving_cell_tiles, budget)?;
    let private_range_tiles = difference(&selected_range_tiles, &surviving_range_tiles, budget)?;

    let mut inspected_cell_tiles = HashMap::new();
    let mut inspected_range_tiles = HashMap::new();
    inspected_cell_tiles
        .try_reserve(
            selected_cell_tiles
                .len()
                .saturating_add(surviving_cell_tiles.len()),
        )
        .map_err(|_| {
            allocation(
                selected_cell_tiles
                    .len()
                    .saturating_add(surviving_cell_tiles.len()),
            )
        })?;
    inspected_range_tiles
        .try_reserve(
            selected_range_tiles
                .len()
                .saturating_add(surviving_range_tiles.len()),
        )
        .map_err(|_| {
            allocation(
                selected_range_tiles
                    .len()
                    .saturating_add(surviving_range_tiles.len()),
            )
        })?;

    for tile_id in selected_cell_tiles
        .iter()
        .chain(surviving_cell_tiles.iter())
        .copied()
    {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
        if inspected_cell_tiles.contains_key(&tile_id) {
            continue;
        }
        let tile = inspect_cell_tile(components, graph, tile_id, budget)?;
        inspected_cell_tiles.insert(tile_id, tile);
    }
    for tile_id in selected_range_tiles
        .iter()
        .chain(surviving_range_tiles.iter())
        .copied()
    {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
        if inspected_range_tiles.contains_key(&tile_id) {
            continue;
        }
        let tile = inspect_range_tile(components, graph, tile_id, budget)?;
        inspected_range_tiles.insert(tile_id, tile);
    }

    let formula_count = selected_formula_count(&owners, &selected, &inspected_cell_tiles, budget)?;

    let mut global_owner_edits = Vec::new();
    let mut global_tile_edits = Vec::new();
    let mut global_tile_edit_ids = HashSet::new();
    global_tile_edit_ids
        .try_reserve(surviving_cell_tiles.len())
        .map_err(|_| allocation(surviving_cell_tiles.len()))?;
    for (index, owner) in owners.iter().enumerate() {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
        if selected.contains(&index) || already_removed.contains(&owner.identifier) {
            continue;
        }
        if owner.snapshot.formula_owner().is_none() {
            reject_surviving_nonprunable_dependencies(
                owner,
                &internal_owner_id_set,
                &owner_uuid_set,
                &inspected_range_tiles,
                budget,
            )?;
            if owner
                .facts
                .has_removed_inline_edges(&internal_owner_id_set, budget)?
            {
                let data = rewrite_owner_edges(owner.data, &internal_owner_ids, budget)?;
                push_owner_edit(
                    &mut global_owner_edits,
                    owner,
                    data,
                    &private_cell_tiles,
                    &private_range_tiles,
                )?;
            }
            for tile_id in &owner.facts.cell_tile_ids {
                let tile = inspected_cell_tiles
                    .get(tile_id)
                    .ok_or(BodyTableDeletionError::InvalidSource)?;
                if tile.has_removed_edges(&internal_owner_id_set, budget)? {
                    if private_cell_tiles.contains(tile_id) {
                        return Err(BodyTableDeletionError::InvalidSource);
                    }
                    budget
                        .charge_payload_work(1)
                        .map_err(super::map_lock_error)?;
                    if insert_set(&mut global_tile_edit_ids, *tile_id)? {
                        let data = rewrite_cell_tile_edges(tile.data, &internal_owner_ids, budget)?;
                        push_tile_edit(&mut global_tile_edits, tile, data)?;
                    }
                }
            }
        } else {
            reject_surviving_owner(
                owner,
                &internal_owner_id_set,
                &owner_uuid_set,
                &inspected_cell_tiles,
                &inspected_range_tiles,
                budget,
            )?;
        }
    }

    append_global_owner_edits(&mut plan, global_owner_edits)?;
    append_global_tile_edits(&mut plan, global_tile_edits)?;

    let engine_data = rewrite_engine(
        engine.data,
        &owner_ids,
        &internal_owner_ids,
        formula_count,
        budget,
    )?;
    push_message_edit(
        &mut plan,
        MessageEdit {
            component_index: engine.component_index,
            object_index: engine.object_index,
            message_index: engine.message_index,
            object_identifier: engine.identifier,
            message_type: CALCULATION_ENGINE_MESSAGE_TYPE,
            data: engine_data,
            remove_object_references: nonzero_ids(&owner_ids)?,
            remove_data_references: Vec::new(),
        },
    )?;

    let mut removal_ids = already_removed;
    append_owner_removals(&mut plan, &owners, &selected, &mut removal_ids, budget)?;
    append_tile_removals(
        &mut plan,
        components,
        graph,
        &private_cell_tiles,
        &private_range_tiles,
        &mut removal_ids,
        budget,
    )?;

    Ok(FormulaPlan { removals: plan })
}

#[derive(Debug)]
struct OwnerFact<'source> {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    data: &'source [u8],
    snapshot: FormulaOwnerDependenciesSnapshot<'source>,
    facts: DependencyFacts,
}

#[derive(Debug, Default)]
struct DependencyFacts {
    cells: Vec<CellFact>,
    cell_tile_ids: Vec<u64>,
    range_tile_ids: Vec<u64>,
    range_back_owner_ids: Vec<u32>,
    opaque_internal_owner_ids: Vec<u32>,
    owner_uuid_dependencies: Vec<UuidKey>,
}

impl DependencyFacts {
    fn has_removed_inline_edges(
        &self,
        removed: &HashSet<u32>,
        budget: &mut WireBudget,
    ) -> Result<bool, BodyTableDeletionError> {
        for cell in &self.cells {
            charge_work(budget, 1)?;
            for identifier in &cell.internal_owner_ids {
                charge_work(budget, 1)?;
                if removed.contains(identifier) {
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }

    fn has_removed_opaque_dependencies(
        &self,
        removed_internal_ids: &HashSet<u32>,
        removed_uuids: &HashSet<UuidKey>,
        budget: &mut WireBudget,
    ) -> Result<bool, BodyTableDeletionError> {
        for identifier in &self.opaque_internal_owner_ids {
            charge_work(budget, 1)?;
            if removed_internal_ids.contains(identifier) {
                return Ok(true);
            }
        }
        for uuid in &self.owner_uuid_dependencies {
            charge_work(budget, 1)?;
            if removed_uuids.contains(uuid) {
                return Ok(true);
            }
        }
        Ok(false)
    }
}

#[derive(Debug)]
struct CellFact {
    row: u32,
    column: u32,
    internal_owner_ids: Vec<u32>,
}

#[derive(Debug)]
struct CellTile<'source> {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    data: &'source [u8],
    cells: Vec<CellFact>,
}

impl CellTile<'_> {
    fn has_removed_edges(
        &self,
        removed: &HashSet<u32>,
        budget: &mut WireBudget,
    ) -> Result<bool, BodyTableDeletionError> {
        for cell in &self.cells {
            charge_work(budget, 1)?;
            for identifier in &cell.internal_owner_ids {
                charge_work(budget, 1)?;
                if removed.contains(identifier) {
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }
}

#[derive(Debug)]
struct RangeTile {
    snapshot: dependency::RangePrecedentsTileSnapshot,
    has_from_to_range: bool,
}

#[derive(Debug)]
struct Engine<'source> {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    data: &'source [u8],
}

#[derive(Debug)]
struct TileEdit {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    data: Vec<u8>,
}

#[derive(Debug, Default)]
struct OwnerVisitor {
    facts: DependencyFacts,
    pending_internal_owner_ids: Vec<u32>,
}

impl DependencyVisitor for OwnerVisitor {
    fn visit_formula_dependency_fact(
        &mut self,
        fact: FormulaDependencyFact,
    ) -> Result<(), DecodeError> {
        match fact {
            FormulaDependencyFact::InternalOwner(fact) => {
                push_decode(&mut self.facts.opaque_internal_owner_ids, fact.owner_id())
            },
            FormulaDependencyFact::OwnerUuid(fact) => push_decode(
                &mut self.facts.owner_uuid_dependencies,
                uuid_key(fact.owner_uuid()),
            ),
        }
    }

    fn visit_tiled_cell_dependency(
        &mut self,
        reference: ReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        push_decode(
            &mut self.facts.cell_tile_ids,
            local_reference_id(reference)?,
        )
    }

    fn visit_tiled_range_dependency(
        &mut self,
        reference: ReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        push_decode(
            &mut self.facts.range_tile_ids,
            local_reference_id(reference)?,
        )
    }

    fn visit_range_back_dependency(
        &mut self,
        record: dependency::RangeBackDependencySnapshot<'_>,
    ) -> Result<(), DecodeError> {
        if let Some(reference) = record.decoded_internal_range_reference() {
            push_decode(&mut self.facts.range_back_owner_ids, reference.owner_id())?;
        }
        Ok(())
    }

    fn visit_expanded_edge_component(
        &mut self,
        component: ExpandedEdgeComponent,
    ) -> Result<(), DecodeError> {
        if component.kind() == ExpandedEdgeKind::InternalOwner {
            push_decode(&mut self.pending_internal_owner_ids, component.value())?;
        }
        Ok(())
    }

    fn visit_cell_record(
        &mut self,
        record: dependency::CellRecordSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        let internal_owner_ids = mem::take(&mut self.pending_internal_owner_ids);
        self.facts
            .cells
            .try_reserve(1)
            .map_err(|_| DecodeError::allocation(self.facts.cells.len().saturating_add(1)))?;
        self.facts.cells.push(CellFact {
            row: record.row(),
            column: record.column(),
            internal_owner_ids,
        });
        Ok(())
    }
}

fn local_reference_id(reference: ReferenceRecord<'_>) -> Result<u64, DecodeError> {
    let reference = reference.reference();
    if reference.identifier() == 0 || reference.deprecated_is_external() == Some(true) {
        return Err(DecodeError::invalid_visitor_result());
    }
    Ok(reference.identifier())
}

#[derive(Debug, Default)]
struct TrackerVisitor {
    owner_ids: Vec<u64>,
}

impl DependencyVisitor for TrackerVisitor {
    fn visit_formula_owner_dependency(
        &mut self,
        reference: ReferenceRecord<'_>,
    ) -> Result<(), DecodeError> {
        push_decode(&mut self.owner_ids, reference.reference().identifier())
    }
}

/// Resolve the one package component that owns the native CalculationEngine
/// archive.  The iWork host accepts the canonical name and numeric save
/// suffixes only; treating arbitrary similarly named components as formula
/// storage would mutate unrelated payloads that happen to reuse message kinds.
fn calculation_engine_component_index(
    components: &litchi_iwa_archive::ComponentCatalog,
) -> Result<Option<usize>, BodyTableDeletionError> {
    let mut found = None;
    for (index, component) in components.iter().enumerate() {
        if !is_calculation_engine_component(component.name()) {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    Ok(found)
}

fn is_calculation_engine_component(name: &str) -> bool {
    let Some(file_name) = name.rsplit('/').next() else {
        return false;
    };
    if file_name == "CalculationEngine.iwa" {
        return true;
    }
    let Some(version) = file_name
        .strip_prefix("CalculationEngine-")
        .and_then(|value| value.strip_suffix(".iwa"))
    else {
        return false;
    };
    !version.is_empty()
        && version
            .split('-')
            .all(|part| !part.is_empty() && part.bytes().all(|byte| byte.is_ascii_digit()))
}

fn collect_formula_owners<'source>(
    components: &'source litchi_iwa_archive::ComponentCatalog,
    engine_component_index: usize,
    budget: &mut WireBudget,
) -> Result<Vec<OwnerFact<'source>>, BodyTableDeletionError> {
    let mut owners = Vec::new();
    let component = components
        .get_index(engine_component_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    for (object_index, object) in component.archive().objects.iter().enumerate() {
        let identifier = object_identifier(object)?;
        let mut owner_message = None;
        for (message_index, message) in object.messages.iter().enumerate() {
            budget
                .charge_payload_work(message.data.len())
                .map_err(super::map_lock_error)?;
            if message.type_ != FORMULA_OWNER_MESSAGE_TYPE {
                continue;
            }
            if owner_message.replace((message_index, message)).is_some() {
                return Err(BodyTableDeletionError::InvalidSource);
            }
        }
        let Some((message_index, message)) = owner_message else {
            continue;
        };
        let (snapshot, facts) = decode_owner(message.data.as_slice(), budget)?;
        owners
            .try_reserve(1)
            .map_err(|_| allocation(owners.len().saturating_add(1)))?;
        owners.push(OwnerFact {
            component_index: engine_component_index,
            object_index,
            message_index,
            identifier,
            data: message.data.as_slice(),
            snapshot,
            facts,
        });
    }
    Ok(owners)
}

fn decode_owner<'source>(
    data: &'source [u8],
    budget: &mut WireBudget,
) -> Result<(FormulaOwnerDependenciesSnapshot<'source>, DependencyFacts), BodyTableDeletionError> {
    let options = codec_options(budget);
    let mut visitor = OwnerVisitor::default();
    let (snapshot, report) =
        dependency::decode_formula_owner_dependencies_with_visitor(data, options, &mut visitor)
            .map_err(map_decode_error)?;
    charge_report(report, budget)?;
    Ok((snapshot, visitor.facts))
}

fn select_owner_family(
    owners: &[OwnerFact<'_>],
    contexts: &HashSet<u64>,
    already_removed: &HashSet<NonZeroU64>,
    budget: &mut WireBudget,
) -> Result<HashSet<usize>, BodyTableDeletionError> {
    let mut selected = HashSet::new();
    selected
        .try_reserve(owners.len())
        .map_err(|_| allocation(owners.len()))?;
    for (index, owner) in owners.iter().enumerate() {
        charge_work(budget, 1)?;
        if already_removed.contains(&owner.identifier) {
            continue;
        }
        if owner
            .snapshot
            .formula_owner()
            .is_some_and(|reference| contexts.contains(&reference.identifier()))
        {
            selected.insert(index);
        }
    }
    if selected.is_empty() {
        return Ok(selected);
    }

    let mut by_base_uuid: HashMap<UuidKey, Vec<usize>> = HashMap::new();
    by_base_uuid
        .try_reserve(owners.len())
        .map_err(|_| allocation(owners.len()))?;
    for (index, owner) in owners.iter().enumerate() {
        charge_work(budget, 1)?;
        let Some(base_uuid) = owner.snapshot.base_owner_uid() else {
            continue;
        };
        by_base_uuid
            .try_reserve(1)
            .map_err(|_| allocation(by_base_uuid.len().saturating_add(1)))?;
        let children = by_base_uuid.entry(uuid_key(base_uuid)).or_default();
        children
            .try_reserve(1)
            .map_err(|_| allocation(children.len().saturating_add(1)))?;
        children.push(index);
    }

    let mut pending = Vec::new();
    pending
        .try_reserve(selected.len())
        .map_err(|_| allocation(selected.len()))?;
    pending.extend(selected.iter().copied());
    let mut cursor = 0usize;
    while cursor < pending.len() {
        let index = pending[cursor];
        cursor = cursor
            .checked_add(1)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        charge_work(budget, 1)?;
        let key = uuid_key(owners[index].snapshot.formula_owner_uid());
        let Some(children) = by_base_uuid.get(&key) else {
            continue;
        };
        for child in children {
            charge_work(budget, 1)?;
            if already_removed.contains(&owners[*child].identifier) {
                continue;
            }
            if insert_set(&mut selected, *child)? {
                pending
                    .try_reserve(1)
                    .map_err(|_| allocation(pending.len().saturating_add(1)))?;
                pending.push(*child);
            }
        }
    }
    Ok(selected)
}

fn selected_owner_sets(
    owners: &[OwnerFact<'_>],
    selected: &HashSet<usize>,
    budget: &mut WireBudget,
) -> Result<(Vec<u64>, Vec<u32>, Vec<UuidKey>), BodyTableDeletionError> {
    let mut owner_ids = Vec::new();
    let mut internal_ids = Vec::new();
    let mut uuids = Vec::new();
    owner_ids
        .try_reserve(selected.len())
        .map_err(|_| allocation(selected.len()))?;
    internal_ids
        .try_reserve(selected.len())
        .map_err(|_| allocation(selected.len()))?;
    uuids
        .try_reserve(selected.len())
        .map_err(|_| allocation(selected.len()))?;
    for index in selected {
        let owner = &owners[*index];
        owner_ids.push(owner.identifier.get());
        internal_ids.push(owner.snapshot.internal_formula_owner_id());
        uuids.push(uuid_key(owner.snapshot.formula_owner_uid()));
    }
    budget
        .charge_sort_work(owner_ids.len())
        .and_then(|_| budget.charge_sort_work(internal_ids.len()))
        .and_then(|_| budget.charge_sort_work(uuids.len()))
        .map_err(super::map_lock_error)?;
    owner_ids.sort_unstable();
    internal_ids.sort_unstable();
    uuids.sort_unstable();
    if owner_ids.windows(2).any(|pair| pair[0] == pair[1])
        || internal_ids.windows(2).any(|pair| pair[0] == pair[1])
        || uuids.windows(2).any(|pair| pair[0] == pair[1])
    {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok((owner_ids, internal_ids, uuids))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
struct UuidKey(u64, u64);

fn uuid_key(uuid: dependency::UuidSnapshot) -> UuidKey {
    UuidKey(uuid.lower(), uuid.upper())
}

fn find_calculation_engine<'source>(
    components: &'source litchi_iwa_archive::ComponentCatalog,
    engine_component_index: usize,
    selected_owner_ids: &HashSet<u64>,
    budget: &mut WireBudget,
) -> Result<Option<Engine<'source>>, BodyTableDeletionError> {
    let component = components
        .get_index(engine_component_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let mut found = None;
    let mut root_count = 0usize;
    for (object_index, object) in component.archive().objects.iter().enumerate() {
        let identifier = object_identifier(object)?;
        let mut engine_message = None;
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ != CALCULATION_ENGINE_MESSAGE_TYPE {
                continue;
            }
            if engine_message.replace((message_index, message)).is_some() {
                return Err(BodyTableDeletionError::InvalidSource);
            }
        }
        let Some((message_index, message)) = engine_message else {
            continue;
        };
        root_count = root_count
            .checked_add(1)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        budget
            .charge_payload_work(message.data.len())
            .map_err(super::map_lock_error)?;
        let mut visitor = TrackerVisitor::default();
        let (_engine, report) = dependency::decode_calculation_engine_with_visitor(
            message.data.as_slice(),
            codec_options(budget),
            &mut visitor,
        )
        .map_err(map_decode_error)?;
        charge_report(report, budget)?;
        if visitor
            .owner_ids
            .iter()
            .any(|owner_id| selected_owner_ids.contains(owner_id))
        {
            if found.is_some() {
                return Err(BodyTableDeletionError::InvalidSource);
            }
            found = Some(Engine {
                component_index: engine_component_index,
                object_index,
                message_index,
                identifier,
                data: message.data.as_slice(),
            });
            // Keep the borrowed root but continue scanning every engine
            // object so duplicate selected references are rejected.
        }
    }
    if root_count > 1 {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(found)
}

fn inspect_cell_tile<'source>(
    components: &'source litchi_iwa_archive::ComponentCatalog,
    graph: &GraphPlan,
    tile_id: u64,
    budget: &mut WireBudget,
) -> Result<CellTile<'source>, BodyTableDeletionError> {
    let location = location(graph, tile_id)?;
    let object = object_at(components, location)?;
    let (message_index, message) = unique_message(object, CELL_RECORD_TILE_MESSAGE_TYPE)?;
    let mut visitor = OwnerVisitor::default();
    let (_, report) = dependency::decode_cell_record_tile_with_visitor(
        message.data.as_slice(),
        codec_options(budget),
        &mut visitor,
    )
    .map_err(map_decode_error)?;
    charge_report(report, budget)?;
    Ok(CellTile {
        component_index: location.component_index,
        object_index: location.object_index,
        message_index,
        identifier: object_identifier(object)?,
        data: message.data.as_slice(),
        cells: visitor.facts.cells,
    })
}

fn inspect_range_tile(
    components: &litchi_iwa_archive::ComponentCatalog,
    graph: &GraphPlan,
    tile_id: u64,
    budget: &mut WireBudget,
) -> Result<RangeTile, BodyTableDeletionError> {
    let location = location(graph, tile_id)?;
    let object = object_at(components, location)?;
    let (_, message) = unique_message(object, RANGE_PRECEDENTS_TILE_MESSAGE_TYPE)?;
    let mut has_from_to_range = false;
    let mut visitor = RangeTileVisitor {
        has_from_to_range: &mut has_from_to_range,
    };
    let (snapshot, report) = dependency::decode_range_precedents_tile_with_visitor(
        message.data.as_slice(),
        codec_options(budget),
        &mut visitor,
    )
    .map_err(map_decode_error)?;
    charge_report(report, budget)?;
    Ok(RangeTile {
        snapshot,
        has_from_to_range,
    })
}

struct RangeTileVisitor<'a> {
    has_from_to_range: &'a mut bool,
}

impl DependencyVisitor for RangeTileVisitor<'_> {
    fn visit_from_to_range(
        &mut self,
        _record: dependency::FromToRangeSnapshot<'_>,
    ) -> Result<(), DecodeError> {
        if *self.has_from_to_range {
            return Err(DecodeError::invalid_visitor_result());
        }
        *self.has_from_to_range = true;
        Ok(())
    }
}

fn selected_formula_count(
    owners: &[OwnerFact<'_>],
    selected: &HashSet<usize>,
    tiles: &HashMap<u64, CellTile<'_>>,
    budget: &mut WireBudget,
) -> Result<u64, BodyTableDeletionError> {
    let mut total = 0_u64;
    for index in selected {
        charge_work(budget, 1)?;
        let owner = &owners[*index];
        let mut coords = HashSet::new();
        coords
            .try_reserve(owner.facts.cells.len())
            .map_err(|_| allocation(owner.facts.cells.len()))?;
        for cell in &owner.facts.cells {
            charge_work(budget, 1)?;
            insert_set(&mut coords, (cell.row, cell.column))?;
        }
        for tile_id in &owner.facts.cell_tile_ids {
            charge_work(budget, 1)?;
            let tile = tiles
                .get(tile_id)
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            for cell in &tile.cells {
                charge_work(budget, 1)?;
                insert_set(&mut coords, (cell.row, cell.column))?;
            }
        }
        total = total
            .checked_add(
                u64::try_from(coords.len()).map_err(|_| BodyTableDeletionError::InvalidSource)?,
            )
            .ok_or(BodyTableDeletionError::InvalidSource)?;
    }
    Ok(total)
}

fn reject_surviving_owner(
    owner: &OwnerFact<'_>,
    removed_internal_ids: &HashSet<u32>,
    removed_uuids: &HashSet<UuidKey>,
    cell_tiles: &HashMap<u64, CellTile<'_>>,
    range_tiles: &HashMap<u64, RangeTile>,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    reject_surviving_nonprunable_dependencies(
        owner,
        removed_internal_ids,
        removed_uuids,
        range_tiles,
        budget,
    )?;
    if owner
        .facts
        .has_removed_inline_edges(removed_internal_ids, budget)?
    {
        return Err(BodyTableDeletionError::UnsupportedDependency);
    }
    for tile_id in &owner.facts.cell_tile_ids {
        charge_work(budget, 1)?;
        let tile = cell_tiles
            .get(tile_id)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if tile.has_removed_edges(removed_internal_ids, budget)? {
            return Err(BodyTableDeletionError::UnsupportedDependency);
        }
    }
    Ok(())
}

fn reject_surviving_nonprunable_dependencies(
    owner: &OwnerFact<'_>,
    removed_internal_ids: &HashSet<u32>,
    removed_uuids: &HashSet<UuidKey>,
    range_tiles: &HashMap<u64, RangeTile>,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut removed_range_back = false;
    for identifier in &owner.facts.range_back_owner_ids {
        charge_work(budget, 1)?;
        removed_range_back |= removed_internal_ids.contains(identifier);
    }
    let removed_opaque =
        owner
            .facts
            .has_removed_opaque_dependencies(removed_internal_ids, removed_uuids, budget)?;
    if removed_range_back || removed_opaque {
        return Err(BodyTableDeletionError::UnsupportedDependency);
    }
    for tile_id in &owner.facts.range_tile_ids {
        charge_work(budget, 1)?;
        let tile = range_tiles
            .get(tile_id)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if removed_internal_ids.contains(&tile.snapshot.to_owner_id()) && tile.has_from_to_range {
            return Err(BodyTableDeletionError::UnsupportedDependency);
        }
    }
    Ok(())
}

fn rewrite_engine(
    data: &[u8],
    owner_ids: &[u64],
    internal_ids: &[u32],
    formula_count: u64,
    budget: &mut WireBudget,
) -> Result<Vec<u8>, BodyTableDeletionError> {
    let options = codec_options(budget);
    let prepared = dependency::prepare_calculation_engine_owner_removal(
        data,
        owner_ids,
        internal_ids,
        formula_count,
        options,
    )
    .map_err(map_decode_error)?;
    let requirements = prepared.requirements();
    charge_rewrite_requirements(requirements, budget)?;
    let (output, report) = dependency::execute_calculation_engine_owner_removal(prepared, options)
        .map_err(map_decode_error)?;
    validate_rewrite_report(requirements, report, output.len(), true)?;
    Ok(output)
}

fn rewrite_owner_edges(
    data: &[u8],
    internal_ids: &[u32],
    budget: &mut WireBudget,
) -> Result<Vec<u8>, BodyTableDeletionError> {
    let options = codec_options(budget);
    let prepared =
        dependency::prepare_formula_owner_cell_edges_removal(data, internal_ids, options)
            .map_err(map_decode_error)?;
    let requirements = prepared.requirements();
    charge_rewrite_requirements(requirements, budget)?;
    let (output, report) = dependency::execute_formula_owner_cell_edges_removal(prepared, options)
        .map_err(map_decode_error)?;
    validate_rewrite_report(requirements, report, output.len(), false)?;
    Ok(output)
}

fn rewrite_cell_tile_edges(
    data: &[u8],
    internal_ids: &[u32],
    budget: &mut WireBudget,
) -> Result<Vec<u8>, BodyTableDeletionError> {
    let options = codec_options(budget);
    let prepared = dependency::prepare_cell_record_tile_edges_removal(data, internal_ids, options)
        .map_err(map_decode_error)?;
    let requirements = prepared.requirements();
    charge_rewrite_requirements(requirements, budget)?;
    let (output, report) = dependency::execute_cell_record_tile_edges_removal(prepared, options)
        .map_err(map_decode_error)?;
    validate_rewrite_report(requirements, report, output.len(), false)?;
    Ok(output)
}

fn push_owner_edit<'source>(
    edits: &mut Vec<OwnerEdit>,
    owner: &OwnerFact<'source>,
    data: Vec<u8>,
    private_cell_tiles: &HashSet<u64>,
    private_range_tiles: &HashSet<u64>,
) -> Result<(), BodyTableDeletionError> {
    let remove_object_references = nonzero_vec(
        owner
            .facts
            .cell_tile_ids
            .iter()
            .chain(&owner.facts.range_tile_ids)
            .copied()
            .filter(|id| private_cell_tiles.contains(id) || private_range_tiles.contains(id)),
    )?;
    edits
        .try_reserve(1)
        .map_err(|_| allocation(edits.len().saturating_add(1)))?;
    edits.push(OwnerEdit {
        component_index: owner.component_index,
        object_index: owner.object_index,
        message_index: owner.message_index,
        identifier: owner.identifier,
        data,
        remove_object_references,
    });
    Ok(())
}

#[derive(Debug)]
struct OwnerEdit {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    identifier: NonZeroU64,
    data: Vec<u8>,
    remove_object_references: Vec<NonZeroU64>,
}

fn push_tile_edit(
    edits: &mut Vec<TileEdit>,
    tile: &CellTile<'_>,
    data: Vec<u8>,
) -> Result<(), BodyTableDeletionError> {
    edits
        .try_reserve(1)
        .map_err(|_| allocation(edits.len().saturating_add(1)))?;
    edits.push(TileEdit {
        component_index: tile.component_index,
        object_index: tile.object_index,
        message_index: tile.message_index,
        identifier: tile.identifier,
        data,
    });
    Ok(())
}

fn append_global_owner_edits(
    plan: &mut RemovalPlan,
    edits: Vec<OwnerEdit>,
) -> Result<(), BodyTableDeletionError> {
    plan.message_edits
        .try_reserve(edits.len())
        .map_err(|_| allocation(edits.len()))?;
    for edit in edits {
        plan.message_edits.push(MessageEdit {
            component_index: edit.component_index,
            object_index: edit.object_index,
            message_index: edit.message_index,
            object_identifier: edit.identifier,
            message_type: FORMULA_OWNER_MESSAGE_TYPE,
            data: edit.data,
            remove_object_references: edit.remove_object_references,
            remove_data_references: Vec::new(),
        });
    }
    Ok(())
}

fn append_global_tile_edits(
    plan: &mut RemovalPlan,
    edits: Vec<TileEdit>,
) -> Result<(), BodyTableDeletionError> {
    plan.message_edits
        .try_reserve(edits.len())
        .map_err(|_| allocation(edits.len()))?;
    for edit in edits {
        plan.message_edits.push(MessageEdit {
            component_index: edit.component_index,
            object_index: edit.object_index,
            message_index: edit.message_index,
            object_identifier: edit.identifier,
            message_type: CELL_RECORD_TILE_MESSAGE_TYPE,
            data: edit.data,
            remove_object_references: Vec::new(),
            remove_data_references: Vec::new(),
        });
    }
    Ok(())
}

fn append_owner_removals(
    plan: &mut RemovalPlan,
    owners: &[OwnerFact<'_>],
    selected: &HashSet<usize>,
    removal_ids: &mut HashSet<NonZeroU64>,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    plan.object_removals
        .try_reserve(selected.len())
        .map_err(|_| allocation(selected.len()))?;
    for index in selected {
        charge_work(budget, 1)?;
        let owner = &owners[*index];
        if !insert_set(removal_ids, owner.identifier)? {
            continue;
        }
        plan.object_removals.push(ObjectRemoval {
            component_index: owner.component_index,
            object_index: owner.object_index,
            identifier: owner.identifier,
        });
    }
    Ok(())
}

fn append_tile_removals(
    plan: &mut RemovalPlan,
    components: &litchi_iwa_archive::ComponentCatalog,
    graph: &GraphPlan,
    cell_tiles: &HashSet<u64>,
    range_tiles: &HashSet<u64>,
    removal_ids: &mut HashSet<NonZeroU64>,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let total = cell_tiles.len().saturating_add(range_tiles.len());
    plan.object_removals
        .try_reserve(total)
        .map_err(|_| allocation(total))?;
    for id in cell_tiles.iter().chain(range_tiles.iter()).copied() {
        charge_work(budget, 1)?;
        let location = location(graph, id)?;
        let object = object_at(components, location)?;
        let identifier = object_identifier(object)?;
        if !insert_set(removal_ids, identifier)? {
            continue;
        }
        plan.object_removals.push(ObjectRemoval {
            component_index: location.component_index,
            object_index: location.object_index,
            identifier,
        });
    }
    Ok(())
}

fn unique_message(
    object: &litchi_iwa_core::ArchiveObject,
    message_type: u32,
) -> Result<(usize, &RawMessage), BodyTableDeletionError> {
    let mut result = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if result.replace((index, message)).is_some() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    result.ok_or(BodyTableDeletionError::InvalidSource)
}

fn location(
    graph: &GraphPlan,
    identifier: u64,
) -> Result<&DeletionObjectLocation, BodyTableDeletionError> {
    let identifier = NonZeroU64::new(identifier).ok_or(BodyTableDeletionError::InvalidSource)?;
    graph
        .index
        .get(identifier)
        .ok_or(BodyTableDeletionError::InvalidSource)
}

fn object_at<'a>(
    components: &'a litchi_iwa_archive::ComponentCatalog,
    location: &DeletionObjectLocation,
) -> Result<&'a litchi_iwa_core::ArchiveObject, BodyTableDeletionError> {
    components
        .get_index(location.component_index)
        .and_then(|component| component.archive().objects.get(location.object_index))
        .ok_or(BodyTableDeletionError::InvalidSource)
}

fn object_identifier(
    object: &litchi_iwa_core::ArchiveObject,
) -> Result<NonZeroU64, BodyTableDeletionError> {
    NonZeroU64::new(
        object
            .archive_info
            .identifier
            .ok_or(BodyTableDeletionError::InvalidSource)?,
    )
    .ok_or(BodyTableDeletionError::InvalidSource)
}

fn object_removal_set(
    removals: &[ObjectRemoval],
) -> Result<HashSet<NonZeroU64>, BodyTableDeletionError> {
    let mut set = HashSet::new();
    set.try_reserve(removals.len())
        .map_err(|_| allocation(removals.len()))?;
    for removal in removals {
        insert_set(&mut set, removal.identifier)?;
    }
    Ok(set)
}

fn set_from_slice<T: Copy + Eq + std::hash::Hash>(
    values: &[T],
) -> Result<HashSet<T>, BodyTableDeletionError> {
    let mut set = HashSet::new();
    set.try_reserve(values.len())
        .map_err(|_| allocation(values.len()))?;
    for value in values {
        insert_set(&mut set, *value)?;
    }
    Ok(set)
}

fn set_from_slice_u64(values: &[NonZeroU64]) -> Result<HashSet<u64>, BodyTableDeletionError> {
    let mut set = HashSet::new();
    set.try_reserve(values.len())
        .map_err(|_| allocation(values.len()))?;
    for value in values {
        insert_set(&mut set, value.get())?;
    }
    Ok(set)
}

fn insert_set<T: Eq + std::hash::Hash>(
    set: &mut HashSet<T>,
    value: T,
) -> Result<bool, BodyTableDeletionError> {
    if set.contains(&value) {
        return Ok(false);
    }
    set.try_reserve(1)
        .map_err(|_| allocation(set.len().saturating_add(1)))?;
    Ok(set.insert(value))
}

fn insert_all<T: Copy + Eq + std::hash::Hash>(
    set: &mut HashSet<T>,
    values: &[T],
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    for value in values {
        charge_work(budget, 1)?;
        insert_set(set, *value)?;
    }
    Ok(())
}

fn difference(
    left: &HashSet<u64>,
    right: &HashSet<u64>,
    budget: &mut WireBudget,
) -> Result<HashSet<u64>, BodyTableDeletionError> {
    let mut result = HashSet::new();
    result
        .try_reserve(left.len())
        .map_err(|_| allocation(left.len()))?;
    for value in left {
        charge_work(budget, 1)?;
        if !right.contains(value) {
            insert_set(&mut result, *value)?;
        }
    }
    Ok(result)
}

fn nonzero_ids<T>(values: &[T]) -> Result<Vec<NonZeroU64>, BodyTableDeletionError>
where
    T: Copy + Into<u64>,
{
    let mut result = Vec::new();
    result
        .try_reserve_exact(values.len())
        .map_err(|_| allocation(values.len()))?;
    for value in values {
        result.push(NonZeroU64::new((*value).into()).ok_or(BodyTableDeletionError::InvalidSource)?);
    }
    Ok(result)
}

fn nonzero_vec(
    values: impl Iterator<Item = u64>,
) -> Result<Vec<NonZeroU64>, BodyTableDeletionError> {
    let mut result = Vec::new();
    for value in values {
        result
            .try_reserve(1)
            .map_err(|_| allocation(result.len().saturating_add(1)))?;
        result.push(NonZeroU64::new(value).ok_or(BodyTableDeletionError::InvalidSource)?);
    }
    Ok(result)
}

fn push_message_edit(
    plan: &mut RemovalPlan,
    edit: MessageEdit,
) -> Result<(), BodyTableDeletionError> {
    plan.message_edits
        .try_reserve(1)
        .map_err(|_| allocation(plan.message_edits.len().saturating_add(1)))?;
    plan.message_edits.push(edit);
    Ok(())
}

fn codec_options(budget: &WireBudget) -> DecodeOptions {
    let limits = budget.wire_limits();
    DecodeOptions::new(
        limits.max_input_bytes().min(limits.max_output_bytes()),
        budget.remaining_wire_fields().max(1),
        budget.remaining_wire_work().max(1),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        budget.remaining_payload_references().max(1),
        limits.max_input_bytes(),
    )
}

fn charge_report(
    report: DecodeReport,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_bytes(report.source_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references(),
            )
        })
        .and_then(|_| budget.charge_payload_work(report.reference_bytes()))
        .and_then(|_| budget.charge_payload_work(report.text_bytes()))
        .map_err(super::map_lock_error)
}

/// Charge the non-work shape of a report after a rewrite's aggregate work has
/// already been reserved from its preflight requirements.  Rewrite reports
/// include the source and result validation passes in
/// `rewrite_work_bytes`; charging their `work_bytes` again here would count
/// those passes twice and make a valid deletion depend on which codec phase
/// happened to run first.
fn charge_report_shape(
    report: DecodeReport,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_bytes(report.source_bytes())
        .and_then(|_| {
            budget.charge_codec_report(report.fields(), 0, report.max_depth(), report.references())
        })
        .and_then(|_| budget.charge_payload_work(report.reference_bytes()))
        .and_then(|_| budget.charge_payload_work(report.text_bytes()))
        .map_err(super::map_lock_error)
}

fn charge_rewrite_requirements(
    requirements: dependency::FormulaDependencyRewriteRequirements,
    budget: &mut WireBudget,
) -> Result<(), BodyTableDeletionError> {
    charge_report_shape(requirements.source(), budget)?;
    budget
        .charge_codec_report(
            requirements.result_fields(),
            0,
            requirements.result_max_depth(),
            requirements.result_references(),
        )
        .and_then(|_| budget.charge_payload_work(requirements.result_reference_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.result_text_bytes()))
        .and_then(|_| budget.charge_output_bytes(requirements.output_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.rewrite_work_bytes()))
        .map_err(super::map_lock_error)
}

fn validate_rewrite_report(
    requirements: dependency::FormulaDependencyRewriteRequirements,
    report: dependency::FormulaDependencyRewriteReport,
    output_len: usize,
    exact_output: bool,
) -> Result<(), BodyTableDeletionError> {
    let source = report.source();
    let required_source = requirements.source();
    let result = report.result();
    let output_matches = if exact_output {
        output_len == requirements.output_bytes()
    } else {
        output_len <= requirements.output_bytes()
    };
    if !output_matches
        || output_len != report.output_bytes()
        || report.rewrite_work_bytes() > requirements.rewrite_work_bytes()
        || source.source_bytes() > required_source.source_bytes()
        || source.fields() > required_source.fields()
        || source.work_bytes() > required_source.work_bytes()
        || source.max_depth() > required_source.max_depth()
        || source.references() > required_source.references()
        || source.reference_bytes() > required_source.reference_bytes()
        || source.text_bytes() > required_source.text_bytes()
        || result.fields() > requirements.result_fields()
        || result.max_depth() > requirements.result_max_depth()
        || result.references() > requirements.result_references()
        || result.reference_bytes() > requirements.result_reference_bytes()
        || result.text_bytes() > requirements.result_text_bytes()
    {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(())
}

fn map_decode_error(error: DecodeError) -> BodyTableDeletionError {
    let Some(limit) = error.resource_limit() else {
        return BodyTableDeletionError::InvalidSource;
    };
    match limit {
        DecodeLimit::Bytes { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        DecodeLimit::References { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::PayloadReferences,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        DecodeLimit::Fields { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireFields,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        DecodeLimit::Work { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireWork,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        DecodeLimit::Nesting { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        DecodeLimit::Allocation { requested } => allocation(requested),
        DecodeLimit::Text { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        DecodeLimit::Retained { observed, maximum } => BodyTableDeletionError::LimitExceeded {
            kind: super::BodyTableDeletionLimitKind::WireOutputBytes,
            observed: usize_u64(observed),
            maximum: usize_u64(maximum),
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn push_decode<T>(values: &mut Vec<T>, value: T) -> Result<(), DecodeError> {
    values
        .try_reserve(1)
        .map_err(|_| DecodeError::allocation(values.len().saturating_add(1)))?;
    values.push(value);
    Ok(())
}

fn allocation(amount: usize) -> BodyTableDeletionError {
    BodyTableDeletionError::Allocation { amount }
}

fn charge_work(budget: &mut WireBudget, amount: usize) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(amount)
        .map_err(super::map_lock_error)
}
