//! Source-authorized physical archive edits for Pages body-table deletion.
//!
//! The graph and formula phases decide which records may disappear.  This
//! module is deliberately only the physical consumer of that proof: it
//! validates every location against the retained source, rewrites selected
//! message payloads with the archive core's source-preserving primitives, and
//! removes whole objects or components in one bounded worklist.  In
//! particular, it never discovers an object by parsing a generated protobuf
//! value and it never publishes a partially edited component.

use std::collections::HashSet;

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::Entry;
use litchi_iwa_core::archive::DataReferencePruning;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};

use super::{
    ArchivePlan, BodyTableDeletionError, ComponentEdit, FormulaPlan, GraphPlan, MessageEdit,
    MetadataPlan, Package, RemovalPlan, table_lock,
};

const MAX_VARINT_BYTES: usize = 10;

/// Borrowed view of the three phase plans.  The archive phase must not clone
/// message payloads merely to concatenate plans: all records already point at
/// the same immutable source witness, and the final edit buffers are created
/// only after the physical bounds have passed.
struct RemovalView<'a> {
    graph: &'a RemovalPlan,
    formula: &'a RemovalPlan,
    metadata: &'a RemovalPlan,
    text: Option<&'a MessageEdit>,
}

impl RemovalView<'_> {
    fn message_count(&self) -> usize {
        self.graph
            .message_edits
            .len()
            .saturating_add(self.formula.message_edits.len())
            .saturating_add(self.metadata.message_edits.len())
            .saturating_add(usize::from(self.text.is_some()))
    }

    fn object_count(&self) -> usize {
        self.graph
            .object_removals
            .len()
            .saturating_add(self.formula.object_removals.len())
            .saturating_add(self.metadata.object_removals.len())
    }

    fn component_count(&self) -> usize {
        self.graph
            .component_removals
            .len()
            .saturating_add(self.formula.component_removals.len())
            .saturating_add(self.metadata.component_removals.len())
    }

    fn for_each_message(
        &self,
        mut visit: impl FnMut(&MessageEdit) -> Result<(), BodyTableDeletionError>,
    ) -> Result<(), BodyTableDeletionError> {
        if let Some(edit) = self.text {
            visit(edit)?;
        }
        for edit in &self.graph.message_edits {
            visit(edit)?;
        }
        for edit in &self.formula.message_edits {
            visit(edit)?;
        }
        for edit in &self.metadata.message_edits {
            visit(edit)?;
        }
        Ok(())
    }

    fn for_each_object(
        &self,
        mut visit: impl FnMut(&super::ObjectRemoval) -> Result<(), BodyTableDeletionError>,
    ) -> Result<(), BodyTableDeletionError> {
        for removal in &self.graph.object_removals {
            visit(removal)?;
        }
        for removal in &self.formula.object_removals {
            visit(removal)?;
        }
        for removal in &self.metadata.object_removals {
            visit(removal)?;
        }
        Ok(())
    }

    fn for_each_component(
        &self,
        mut visit: impl FnMut(&super::ComponentRemoval) -> Result<(), BodyTableDeletionError>,
    ) -> Result<(), BodyTableDeletionError> {
        for removal in &self.graph.component_removals {
            visit(removal)?;
        }
        for removal in &self.formula.component_removals {
            visit(removal)?;
        }
        for removal in &self.metadata.component_removals {
            visit(removal)?;
        }
        Ok(())
    }
}

/// Prepare every changed IWA component and return owned replacement payloads.
///
/// The returned `ArchivePlan` owns all compressed replacement streams, so the
/// parent coordinator can build borrowed `EntryEdit` values only after this
/// phase has completed.  All source archive bounds are charged before the
/// first Snappy decoder or mutable archive is allocated.
pub(super) fn prepare(
    source: &Package,
    graph: &GraphPlan,
    formula: &FormulaPlan,
    metadata: &MetadataPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<ArchivePlan, BodyTableDeletionError> {
    let source_catalog = &source.state.source;
    if !source_catalog.source_is_exact() {
        return Err(BodyTableDeletionError::UnsupportedSource);
    }

    // The body text edit is another source-authorized message replacement.
    // Keep it owned only for the duration of this preparation so the common
    // validation and archive worklist can borrow it without cloning payloads
    // into a merged `RemovalPlan`.
    let text_edit = super::text::prepare(source, graph, budget)?;
    let removals = RemovalView {
        graph: &graph.removals,
        formula: &formula.removals,
        metadata: &metadata.removals,
        text: Some(&text_edit),
    };
    validate_plan(source_catalog, &removals, budget)?;

    let mut changed_components = Vec::new();
    let reserve_hint = removals
        .message_count()
        .checked_add(removals.object_count())
        .and_then(|value| value.checked_add(removals.component_count()))
        .and_then(|value| value.checked_add(1))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(reserve_hint)
        .map_err(super::map_lock_error)?;
    changed_components.try_reserve(reserve_hint).map_err(|_| {
        BodyTableDeletionError::Allocation {
            amount: reserve_hint,
        }
    })?;
    removals.for_each_message(|edit| {
        insert_component_index(&mut changed_components, edit.component_index, budget)?;
        Ok(())
    })?;
    removals.for_each_object(|removal| {
        insert_component_index(&mut changed_components, removal.component_index, budget)?;
        Ok(())
    })?;
    removals.for_each_component(|removal| {
        insert_component_index(&mut changed_components, removal.component_index, budget)?;
        Ok(())
    })?;
    let metadata_index = source_catalog
        .components()
        .iter()
        .position(|component| component.name() == "Index/Metadata.iwa");
    let metadata_was_added = metadata_index.is_some_and(|index| {
        if changed_components.contains(&index) {
            false
        } else {
            // The actual insertion is performed below so its work is charged
            // through the common bounded helper.
            true
        }
    });
    if let Some(metadata_index) = metadata_index {
        insert_component_index(&mut changed_components, metadata_index, budget)?;
    }

    let deleted_indices = physical_deleted_indices(source_catalog, &removals, budget)?;

    // Resolve each physical component to its already materialized package
    // entry once.  The archive source index is the authority for component
    // identity; this small borrowed table avoids an unbounded repeated ZIP
    // scan while the same component is preflighted and then edited.
    let mut changed_entries: Vec<&Entry> = Vec::new();
    budget
        .charge_payload_work(changed_components.len())
        .map_err(super::map_lock_error)?;
    changed_entries
        .try_reserve_exact(changed_components.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: changed_components.len(),
        })?;
    for component_index in &changed_components {
        let component = source_catalog
            .components()
            .get_index(*component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let mut scanned = 0usize;
        let entry = source_catalog
            .package()
            .iter()
            .find(|entry| {
                scanned = scanned.saturating_add(1);
                entry.name() == component.name()
            })
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        budget
            .charge_payload_work(scanned)
            .map_err(super::map_lock_error)?;
        changed_entries.push(entry);
    }

    // Metadata preparation exposes its exact output size before reserving the
    // candidate.  Use that size to charge all selected physical components
    // before the first metadata or archive execution allocation.
    let mut physical_preflighted = false;
    let metadata_payload = super::metadata::rewrite_with_preflight(
        source,
        graph,
        formula,
        metadata,
        budget,
        |metadata_length, budget| {
            if physical_preflighted {
                return Ok(());
            }
            preflight_component_bounds(
                source_catalog,
                &changed_components,
                &changed_entries,
                &deleted_indices,
                &removals,
                Some(metadata_length),
                budget,
            )?;
            physical_preflighted = true;
            Ok(())
        },
    )?;
    if metadata_payload.is_none() && metadata_was_added {
        if let Some(metadata_index) = metadata_index {
            changed_components.retain(|index| *index != metadata_index);
            changed_entries.retain(|entry| entry.name() != "Index/Metadata.iwa");
        }
    }
    if !physical_preflighted {
        preflight_component_bounds(
            source_catalog,
            &changed_components,
            &changed_entries,
            &deleted_indices,
            &removals,
            metadata_payload.as_ref().map(Vec::len),
            budget,
        )?;
    }

    let mut plan = ArchivePlan {
        removed_object_ids: collect_removed_object_ids(&removals, budget)?,
        ..ArchivePlan::default()
    };
    budget
        .charge_payload_work(changed_components.len())
        .map_err(super::map_lock_error)?;
    plan.edits
        .try_reserve(changed_components.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: changed_components.len(),
        })?;
    budget
        .charge_payload_work(deleted_indices.len())
        .map_err(super::map_lock_error)?;
    plan.deleted_entries
        .try_reserve_exact(deleted_indices.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: deleted_indices.len(),
        })?;
    for (component_index, _component) in source_catalog.components().iter().enumerate() {
        if !deleted_indices.contains(&component_index) {
            continue;
        }
        let component = source_catalog
            .components()
            .get_index(component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        budget
            .charge_payload_work(component.name().len())
            .map_err(super::map_lock_error)?;
        let name = try_clone_str(component.name())?;
        plan.deleted_entries.push(name);
    }

    for (component_index, entry) in changed_components.into_iter().zip(changed_entries) {
        let component = source_catalog
            .components()
            .get_index(component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if deleted_indices.contains(&component_index) {
            if has_live_records(component.archive(), component_index, &removals, budget)? {
                return Err(BodyTableDeletionError::InvalidSource);
            }
            continue;
        }
        let archive_limits = source_catalog
            .limits()
            .effective_archive_limits()
            .map_err(super::map_archive_error)?;
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            source_catalog
                .limits()
                .snappy_limits()
                .map_err(super::map_archive_error)?,
        )
        .map_err(map_core_error)?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(map_core_error)?;
        apply_component_changes(
            &mut archive,
            component_index,
            &removals,
            metadata_payload
                .as_ref()
                .filter(|_| component.name() == "Index/Metadata.iwa")
                .map(Vec::as_slice),
            archive_limits,
            budget,
        )?;
        let encoded_len = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        budget
            .charge_payload_work(encoded_len)
            .map_err(super::map_lock_error)?;
        let encoded = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&encoded).map_err(map_core_error)?;
        if compressed.is_empty() {
            return Err(BodyTableDeletionError::Verification);
        }
        budget
            .charge_payload_work(component.name().len())
            .map_err(super::map_lock_error)?;
        let name = try_clone_str(component.name())?;
        plan.edits
            .try_reserve(1)
            .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
        plan.edits.push(ComponentEdit {
            component_index,
            name,
            data: compressed,
        });
        plan.touched_components = plan
            .touched_components
            .checked_add(1)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
    }

    plan.removed_objects = removals.object_count();
    plan.removed_components = plan.deleted_entries.len();
    Ok(plan)
}

fn validate_plan(
    source: &SourceCatalog,
    removals: &RemovalView<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let mut message_slots = HashSet::new();
    message_slots
        .try_reserve(removals.message_count())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removals.message_count(),
        })?;
    let mut object_slots = HashSet::new();
    object_slots
        .try_reserve(removals.object_count())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removals.object_count(),
        })?;
    let mut object_identifiers = HashSet::new();
    object_identifiers
        .try_reserve(removals.object_count())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removals.object_count(),
        })?;
    removals.for_each_message(|edit| {
        budget
            .charge_payload_work(1)
            .and_then(|_| budget.charge_payload_work(edit.data.len()))
            .map_err(super::map_lock_error)?;
        let component = source
            .components()
            .get_index(edit.component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let object = component
            .archive()
            .objects
            .get(edit.object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if object.archive_info.identifier != Some(edit.object_identifier.get())
            || edit.message_index >= object.messages.len()
            || object.messages[edit.message_index].type_ != edit.message_type
            || !message_slots.insert((edit.component_index, edit.object_index, edit.message_index))
        {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        if edit.data.len() > source.limits().max_iwa_stream_bytes() {
            return Err(BodyTableDeletionError::LimitExceeded {
                kind: super::BodyTableDeletionLimitKind::PayloadBytes,
                observed: usize_to_u64(edit.data.len()),
                maximum: usize_to_u64(source.limits().max_iwa_stream_bytes()),
            });
        }
        Ok(())
    })?;
    removals.for_each_object(|removal| {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
        let component = source
            .components()
            .get_index(removal.component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let object = component
            .archive()
            .objects
            .get(removal.object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if object.archive_info.identifier != Some(removal.identifier.get())
            || !object_slots.insert((removal.component_index, removal.object_index))
            || !object_identifiers.insert(removal.identifier)
        {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        Ok(())
    })?;
    removals.for_each_message(|edit| {
        if object_slots.contains(&(edit.component_index, edit.object_index)) {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        Ok(())
    })?;
    let mut component_slots = HashSet::new();
    component_slots
        .try_reserve(removals.component_count())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removals.component_count(),
        })?;
    removals.for_each_component(|removal| {
        if !component_slots.insert(removal.component_index) {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        if removal.identifier == 0 {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        let component = source
            .components()
            .get_index(removal.component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if component.name() != removal.name.as_ref() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        Ok(())
    })?;
    removals.for_each_message(|edit| {
        if component_slots.contains(&edit.component_index) {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        Ok(())
    })?;
    Ok(())
}

fn collect_removed_object_ids(
    removals: &RemovalView<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<std::num::NonZeroU64>, BodyTableDeletionError> {
    let count = removals.object_count();
    let sort_work = sort_work(count)?;
    budget
        .charge_payload_work(sort_work)
        .map_err(super::map_lock_error)?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(count)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: count })?;
    removals.for_each_object(|removal| {
        identifiers.push(removal.identifier);
        Ok(())
    })?;
    identifiers.sort_unstable();
    if identifiers.windows(2).any(|window| window[0] == window[1]) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(identifiers)
}

fn sort_work(count: usize) -> Result<usize, BodyTableDeletionError> {
    let steps = if count < 2 {
        1
    } else {
        (usize::BITS - (count - 1).leading_zeros()) as usize
    };
    count
        .checked_mul(steps)
        .ok_or(BodyTableDeletionError::InvalidSource)
}

fn insert_component_index(
    indices: &mut Vec<usize>,
    index: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(indices.len().saturating_add(1))
        .map_err(super::map_lock_error)?;
    if indices.contains(&index) {
        return Ok(());
    }
    indices
        .try_reserve(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    indices.push(index);
    Ok(())
}

fn physical_deleted_indices(
    source: &SourceCatalog,
    removals: &RemovalView<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<HashSet<usize>, BodyTableDeletionError> {
    let component_count = source.components().len();
    let work = component_count
        .checked_add(removals.object_count())
        .and_then(|value| value.checked_add(removals.component_count()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(work)
        .map_err(super::map_lock_error)?;
    let mut removed_counts: Vec<usize> = Vec::new();
    removed_counts
        .try_reserve_exact(component_count)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: component_count,
        })?;
    removed_counts.resize(component_count, 0usize);
    removals.for_each_object(|removal| {
        let count = removed_counts
            .get_mut(removal.component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        *count = count
            .checked_add(1)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        Ok(())
    })?;
    let mut deleted = HashSet::new();
    let reserve = component_count
        .checked_add(removals.component_count())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    deleted
        .try_reserve(reserve)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: reserve })?;
    removals.for_each_component(|removal| {
        if !deleted.insert(removal.component_index) {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        Ok(())
    })?;
    for (component_index, component) in source.components().iter().enumerate() {
        let removed_count = removed_counts
            .get(component_index)
            .copied()
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if removed_count == component.archive().objects.len()
            && (removed_count != 0 || deleted.contains(&component_index))
        {
            deleted.insert(component_index);
        }
    }
    Ok(deleted)
}

fn preflight_component_bounds(
    source: &SourceCatalog,
    changed_components: &[usize],
    changed_entries: &[&Entry],
    deleted_indices: &HashSet<usize>,
    removals: &RemovalView<'_>,
    metadata_length: Option<usize>,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    if changed_components.len() != changed_entries.len() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    // Compute every replacement bound before parsing any selected stream. A
    // deletion-only archive is omitted from this pass because reassembly will
    // remove its physical ZIP member rather than allocate a replacement.
    for (component_index, entry) in changed_components.iter().zip(changed_entries) {
        if deleted_indices.contains(component_index) {
            continue;
        }
        let component = source
            .components()
            .get_index(*component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        let source_length = archive_source_length(component.archive())?;
        let archive_bound = component_archive_bound(
            source,
            *component_index,
            source_length,
            removals,
            metadata_length.filter(|_| component.name() == "Index/Metadata.iwa"),
        )?;
        let compressed_bound = table_lock::snappy_compressed_bound(archive_bound)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let physical_bound = match entry.metadata().central().compression_method() {
            0 => compressed_bound,
            8 => table_lock::deflate_compressed_bound(compressed_bound)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
            _ => return Err(BodyTableDeletionError::UnsupportedSource),
        };
        budget
            .charge_payload_work(entry.data().len())
            .and_then(|_| budget.charge_payload_work(source_length))
            .and_then(|_| budget.charge_output_bytes(archive_bound))
            .and_then(|_| budget.charge_output_bytes(compressed_bound))
            .and_then(|_| budget.charge_output_bytes(physical_bound))
            .and_then(|_| budget.charge_payload_bytes(archive_bound))
            .and_then(|_| budget.charge_total_payload_bytes(archive_bound))
            .map_err(super::map_lock_error)?;
    }
    Ok(())
}

fn component_archive_bound(
    source: &SourceCatalog,
    component_index: usize,
    source_length: usize,
    removals: &RemovalView<'_>,
    metadata_length: Option<usize>,
) -> Result<usize, BodyTableDeletionError> {
    let component = source
        .components()
        .get_index(component_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let mut bound = source_length;
    let mut touched_objects = HashSet::new();
    touched_objects
        .try_reserve(removals.message_count())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removals.message_count(),
        })?;
    removals.for_each_message(|edit| {
        if edit.component_index != component_index {
            return Ok(());
        }
        let object = component
            .archive()
            .objects
            .get(edit.object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let old = object
            .messages
            .get(edit.message_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?
            .data
            .len();
        bound = bound
            .checked_sub(old)
            .and_then(|value| value.checked_add(edit.data.len()))
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        // A changed message can widen its length field in ArchiveInfo, and
        // the object's framing prefix can widen once per touched object.
        // Reserve both kinds of headroom before parsing the mutable archive.
        bound = bound
            .checked_add(MAX_VARINT_BYTES)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if touched_objects.insert(edit.object_index) {
            bound = bound
                .checked_add(MAX_VARINT_BYTES)
                .ok_or(BodyTableDeletionError::InvalidSource)?;
        }
        Ok(())
    })?;
    removals.for_each_object(|removal| {
        if removal.component_index != component_index {
            return Ok(());
        }
        let object = component
            .archive()
            .objects
            .get(removal.object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let old = object_encoded_len(object)?;
        bound = bound
            .checked_sub(old)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        Ok(())
    })?;
    if let Some(metadata_length) = metadata_length {
        let metadata_object = component
            .archive()
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .find(|message| message.type_ == 11_006)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        bound = bound
            .checked_sub(metadata_object.data.len())
            .and_then(|value| value.checked_add(metadata_length))
            .ok_or(BodyTableDeletionError::InvalidSource)?;
    }
    Ok(bound)
}

fn apply_component_changes(
    archive: &mut Archive,
    component_index: usize,
    removals: &RemovalView<'_>,
    metadata_payload: Option<&[u8]>,
    archive_limits: litchi_iwa_core::Limits,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    removals.for_each_message(|edit| {
        if edit.component_index != component_index {
            return Ok(());
        }
        let object = archive
            .objects
            .get(edit.object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if object.archive_info.identifier != Some(edit.object_identifier.get()) {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        let mut object_references = Vec::new();
        object_references
            .try_reserve_exact(edit.remove_object_references.len())
            .map_err(|_| BodyTableDeletionError::Allocation {
                amount: edit.remove_object_references.len(),
            })?;
        object_references.extend(
            edit.remove_object_references
                .iter()
                .map(|identifier| identifier.get()),
        );
        let mut data_references = Vec::new();
        data_references
            .try_reserve_exact(edit.remove_data_references.len())
            .map_err(|_| BodyTableDeletionError::Allocation {
                amount: edit.remove_data_references.len(),
            })?;
        data_references.extend(
            edit.remove_data_references
                .iter()
                .map(|identifier| identifier.get()),
        );
        // Charge the copy before reserving its destination.  The cumulative
        // budget therefore rejects the whole transaction before this edit can
        // consume a replacement buffer.
        budget
            .charge_payload_work(edit.data.len())
            .map_err(super::map_lock_error)?;
        let data = try_clone_bytes(&edit.data)?;
        let replacement = RawMessage {
            type_: edit.message_type,
            data,
        };
        let object = archive
            .objects
            .get_mut(edit.object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if object_references.is_empty() && data_references.is_empty() {
            object
                .replace_message_preserving_header_with_limits(
                    edit.message_index,
                    replacement,
                    archive_limits,
                )
                .map_err(map_core_error)?;
        } else {
            object
                .replace_message_pruning_references_preserving_header_with_limits(
                    edit.message_index,
                    replacement,
                    &object_references,
                    if data_references.is_empty() {
                        DataReferencePruning::None
                    } else {
                        DataReferencePruning::Selected(&data_references)
                    },
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
        Ok(())
    })?;

    if let Some(metadata_payload) = metadata_payload {
        let mut found = None;
        for (object_index, object) in archive.objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == 11_006 {
                    if found.replace((object_index, message_index)).is_some() {
                        return Err(BodyTableDeletionError::InvalidSource);
                    }
                }
            }
        }
        let (object_index, message_index) = found.ok_or(BodyTableDeletionError::InvalidSource)?;
        let object = archive
            .objects
            .get_mut(object_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        let replacement = RawMessage {
            type_: 11_006,
            data: {
                budget
                    .charge_payload_work(metadata_payload.len())
                    .map_err(super::map_lock_error)?;
                try_clone_bytes(metadata_payload)?
            },
        };
        object
            .replace_message_preserving_header_with_limits(
                message_index,
                replacement,
                archive_limits,
            )
            .map_err(map_core_error)?;
    }

    let mut removals_by_index = Vec::new();
    removals_by_index
        .try_reserve(removals.object_count())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: removals.object_count(),
        })?;
    removals.for_each_object(|removal| {
        if removal.component_index != component_index {
            return Ok(());
        }
        removals_by_index.push(removal.object_index);
        Ok(())
    })?;
    removals_by_index.sort_unstable();
    for (position, object_index) in removals_by_index.iter().enumerate() {
        if *object_index >= archive.objects.len()
            || (position > 0 && *object_index == removals_by_index[position - 1])
        {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    budget
        .charge_payload_work(
            archive
                .objects
                .len()
                .checked_add(removals_by_index.len())
                .ok_or(BodyTableDeletionError::InvalidSource)?,
        )
        .map_err(super::map_lock_error)?;
    let mut removal_position = 0usize;
    let mut original_position = 0usize;
    archive.objects.retain(|_| {
        let remove = removals_by_index
            .get(removal_position)
            .is_some_and(|index| *index == original_position);
        if remove {
            removal_position = removal_position.saturating_add(1);
        }
        original_position = original_position.saturating_add(1);
        !remove
    });
    if removal_position != removals_by_index.len() {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    archive
        .validate_with_limits(archive_limits)
        .map_err(map_core_error)?;
    Ok(())
}

fn has_live_records(
    archive: &Archive,
    component_index: usize,
    removals: &RemovalView<'_>,
    budget: &mut table_lock::WireBudget,
) -> Result<bool, BodyTableDeletionError> {
    let mut removed_count = 0usize;
    removals.for_each_object(|removal| {
        if removal.component_index == component_index {
            removed_count = removed_count
                .checked_add(1)
                .ok_or(BodyTableDeletionError::InvalidSource)?;
        }
        Ok(())
    })?;
    budget
        .charge_payload_work(
            archive
                .objects
                .len()
                .checked_add(removed_count)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
        )
        .map_err(super::map_lock_error)?;
    Ok(removed_count != archive.objects.len())
}

fn archive_source_length(archive: &Archive) -> Result<usize, BodyTableDeletionError> {
    archive
        .objects
        .last()
        .map(|object| {
            object
                .data_offset
                .checked_add(object.data_length)
                .and_then(|length| usize::try_from(length).ok())
                .ok_or(BodyTableDeletionError::InvalidSource)
        })
        .unwrap_or(Ok(0))
}

fn object_encoded_len(object: &ArchiveObject) -> Result<usize, BodyTableDeletionError> {
    let payload = object
        .archive_info
        .message_infos
        .iter()
        .try_fold(0usize, |length, info| {
            let message_length =
                usize::try_from(info.length).map_err(|_| BodyTableDeletionError::InvalidSource)?;
            length
                .checked_add(message_length)
                .ok_or(BodyTableDeletionError::InvalidSource)
        })?;
    let header_len =
        usize::try_from(object.header_length).map_err(|_| BodyTableDeletionError::InvalidSource)?;
    header_len
        .checked_add(payload)
        .ok_or(BodyTableDeletionError::InvalidSource)
}

fn try_clone_bytes(bytes: &[u8]) -> Result<Vec<u8>, BodyTableDeletionError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: bytes.len(),
        })?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn try_clone_str(value: &str) -> Result<Box<str>, BodyTableDeletionError> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: value.len(),
        })?;
    output.push_str(value);
    Ok(output.into_boxed_str())
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyTableDeletionError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableDeletionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => {
                    super::BodyTableDeletionLimitKind::PayloadObjects
                },
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    super::BodyTableDeletionLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    super::BodyTableDeletionLimitKind::PayloadItems
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    super::BodyTableDeletionLimitKind::WireNesting
                },
                _ => super::BodyTableDeletionLimitKind::PayloadBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableDeletionError::Allocation { amount: requested }
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
