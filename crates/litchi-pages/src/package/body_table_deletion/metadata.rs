//! PackageMetadata ownership preparation for Pages body-table deletion.
//!
//! This phase is a source census and an exact removal-plan builder.  It keeps
//! all component locators and registry facts private, rejects versioned or
//! unknown records when they would be claimed by the deletion, and leaves the
//! actual source-preserving codec execution to [`rewrite_with_preflight`].  The archive
//! phase invokes that helper while it is assembling the metadata component so
//! the package still has one publication boundary.

use std::collections::{HashMap, HashSet, hash_map::Entry as HashMapEntry};

use litchi_iwa_protos::package_metadata_codec::{
    self, ComponentRemovalBatch, ComponentSelector, DataMetadataMapRemoval,
    DataReferenceOwnerRemoval as CodecDataReferenceOwnerRemoval,
    ExternalReferenceRemoval as CodecExternalReferenceRemoval, ObjectUuidRemoval,
    PackageMetadataVisitor, RemovalBatch, RemovalSaveTokenBatch, RewriteError, RewriteOptions,
    SaveTokenBatch,
};

use super::{
    BodyTableDeletionError, DataReferenceOwnerRemoval, ExternalReferenceRemoval, FormulaPlan,
    GraphPlan, MetadataPlan, Package, RemovalPlan, UuidRemoval, table_lock,
};

const METADATA_MESSAGE_TYPE: u32 = 11_006;

#[derive(Debug, Clone)]
struct MetadataComponent {
    identifier: u64,
    locator: Box<str>,
    current: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataUuid {
    component: u64,
    object: u64,
    uuid: package_metadata_codec::UuidBits,
    current: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataExternal {
    source: u64,
    target: u64,
    object: Option<u64>,
    weak: Option<bool>,
    current: bool,
    unknown: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataDataOwner {
    component: u64,
    data: u64,
    object: u64,
    count: u32,
    current: bool,
    unknown: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataMap {
    object: u64,
    unknown: bool,
}

#[derive(Debug, Default)]
struct MetadataFacts {
    components: Vec<MetadataComponent>,
    uuids: Vec<MetadataUuid>,
    externals: Vec<MetadataExternal>,
    data_owners: Vec<MetadataDataOwner>,
    ambiguous: Vec<(u64, u64)>,
    maps: Vec<MetadataMap>,
}

/// Borrowed lookup tables for the ownership checks below.  PackageMetadata
/// records are keyed by the object they describe; indexing them once keeps a
/// large component census linear in the number of metadata records instead
/// of rescanning every UUID/reference vector for every removed object.
struct MetadataFactIndex<'facts> {
    component_by_locator: HashMap<&'facts str, &'facts MetadataComponent>,
    component_by_identifier: HashMap<u64, &'facts MetadataComponent>,
    versioned_component_ids: HashSet<u64>,
    uuids_by_object: HashMap<u64, Vec<MetadataUuid>>,
    externals_by_object: HashMap<u64, Vec<MetadataExternal>>,
    data_owners_by_object: HashMap<u64, Vec<MetadataDataOwner>>,
    ambiguous_by_object: HashMap<u64, Vec<u64>>,
    unknown_external_targets: HashSet<u64>,
}

impl PackageMetadataVisitor for MetadataFacts {
    fn visit_component(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        let locator = component.effective_locator();
        let mut owned = String::new();
        owned
            .try_reserve_exact(locator.len())
            .map_err(|_| RewriteError::allocation(locator.len()))?;
        owned.push_str(locator);
        self.components
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(1))?;
        self.components.push(MetadataComponent {
            identifier: component.identifier(),
            locator: owned.into_boxed_str(),
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        self.uuids
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(1))?;
        self.uuids.push(MetadataUuid {
            component: binding.component().identifier(),
            object: binding.object_identifier(),
            uuid: binding.uuid(),
            current: binding.component().is_current(),
        });
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        self.externals
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(1))?;
        self.externals.push(MetadataExternal {
            source: reference.source().identifier(),
            target: reference.target_component_identifier(),
            object: reference.object_identifier(),
            weak: reference.is_weak(),
            current: reference.source().is_current() && !reference.is_versioned(),
            unknown: reference.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), RewriteError> {
        self.data_owners
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(1))?;
        self.data_owners.push(MetadataDataOwner {
            component: owner.component().identifier(),
            data: owner.data_identifier(),
            object: owner.object_identifier(),
            count: owner.count(),
            current: owner.component().is_current(),
            unknown: owner.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), RewriteError> {
        self.ambiguous
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(1))?;
        self.ambiguous.push((component.identifier(), identifier));
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        has_unknown_fields: bool,
    ) -> Result<(), RewriteError> {
        self.maps
            .try_reserve(1)
            .map_err(|_| RewriteError::allocation(1))?;
        self.maps.push(MetadataMap {
            object: object_identifier,
            unknown: has_unknown_fields,
        });
        Ok(())
    }
}

fn build_fact_index<'facts>(
    facts: &'facts MetadataFacts,
    budget: &mut table_lock::WireBudget,
) -> Result<MetadataFactIndex<'facts>, BodyTableDeletionError> {
    let record_count = facts
        .components
        .len()
        .checked_add(facts.uuids.len())
        .and_then(|value| value.checked_add(facts.externals.len()))
        .and_then(|value| value.checked_add(facts.data_owners.len()))
        .and_then(|value| value.checked_add(facts.ambiguous.len()))
        .and_then(|value| value.checked_add(facts.maps.len()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let locator_bytes = facts
        .components
        .iter()
        .try_fold(0usize, |total, component| {
            total
                .checked_add(component.locator.len())
                .ok_or(BodyTableDeletionError::InvalidSource)
        })?;
    budget
        .charge_payload_work(
            record_count
                .checked_add(locator_bytes)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
        )
        .map_err(super::map_lock_error)?;

    let mut component_by_locator = HashMap::new();
    component_by_locator
        .try_reserve(facts.components.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.components.len(),
        })?;
    let mut component_by_identifier = HashMap::new();
    component_by_identifier
        .try_reserve(facts.components.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.components.len(),
        })?;
    let mut versioned_component_ids = HashSet::new();
    versioned_component_ids
        .try_reserve(facts.components.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.components.len(),
        })?;
    for component in facts
        .components
        .iter()
        .filter(|component| component.current)
    {
        if component_by_locator
            .insert(component.locator.as_ref(), component)
            .is_some()
        {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        if component_by_identifier
            .insert(component.identifier, component)
            .is_some()
        {
            return Err(BodyTableDeletionError::InvalidSource);
        }
    }
    for component in facts
        .components
        .iter()
        .filter(|component| !component.current)
    {
        versioned_component_ids.insert(component.identifier);
    }

    let mut uuids_by_object = HashMap::new();
    uuids_by_object
        .try_reserve(facts.uuids.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.uuids.len(),
        })?;
    for binding in facts.uuids.iter().copied() {
        push_indexed(&mut uuids_by_object, binding.object, binding)?;
    }

    let mut externals_by_object = HashMap::new();
    externals_by_object
        .try_reserve(facts.externals.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.externals.len(),
        })?;
    let mut unknown_external_targets = HashSet::new();
    unknown_external_targets
        .try_reserve(facts.externals.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.externals.len(),
        })?;
    for reference in facts.externals.iter().copied() {
        if reference.unknown {
            unknown_external_targets.insert(reference.target);
        }
        if let Some(object) = reference.object {
            push_indexed(&mut externals_by_object, object, reference)?;
        }
    }

    let mut data_owners_by_object = HashMap::new();
    data_owners_by_object
        .try_reserve(facts.data_owners.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.data_owners.len(),
        })?;
    for owner in facts.data_owners.iter().copied() {
        push_indexed(&mut data_owners_by_object, owner.object, owner)?;
    }

    let mut ambiguous_by_object = HashMap::new();
    ambiguous_by_object
        .try_reserve(facts.ambiguous.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: facts.ambiguous.len(),
        })?;
    for (component, object) in facts.ambiguous.iter().copied() {
        push_indexed(&mut ambiguous_by_object, object, component)?;
    }

    Ok(MetadataFactIndex {
        component_by_locator,
        component_by_identifier,
        versioned_component_ids,
        uuids_by_object,
        externals_by_object,
        data_owners_by_object,
        ambiguous_by_object,
        unknown_external_targets,
    })
}

fn push_indexed<T: Copy>(
    index: &mut HashMap<u64, Vec<T>>,
    key: u64,
    value: T,
) -> Result<(), BodyTableDeletionError> {
    match index.entry(key) {
        HashMapEntry::Occupied(mut entry) => {
            entry
                .get_mut()
                .try_reserve(1)
                .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
            entry.get_mut().push(value);
        },
        HashMapEntry::Vacant(entry) => {
            let mut values = Vec::new();
            values
                .try_reserve_exact(1)
                .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
            values.push(value);
            entry.insert(values);
        },
    }
    Ok(())
}

/// Build exact registry removals for the final graph/formula object set.
pub(super) fn prepare(
    source: &Package,
    graph: &GraphPlan,
    formula: &FormulaPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<MetadataPlan, BodyTableDeletionError> {
    let Some((_payload, last, facts)) = inspect(source, budget)? else {
        return Ok(MetadataPlan {
            removals: RemovalPlan::default(),
        });
    };
    let fact_index = build_fact_index(&facts, budget)?;
    let mut object_ids = Vec::new();
    let object_count = graph
        .removals
        .object_removals
        .len()
        .checked_add(formula.removals.object_removals.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    object_ids
        .try_reserve(object_count)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: object_count,
        })?;
    for removal in graph
        .removals
        .object_removals
        .iter()
        .chain(&formula.removals.object_removals)
    {
        object_ids.push(removal.identifier);
    }
    charge_sort_work(object_ids.len(), budget)?;
    object_ids.sort_unstable();
    if object_ids.windows(2).any(|window| window[0] == window[1]) {
        return Err(BodyTableDeletionError::InvalidSource);
    }

    let component_removals = graph
        .removals
        .component_removals
        .iter()
        .chain(&formula.removals.component_removals);
    let mut removed_components = HashSet::new();
    let component_removal_capacity = graph
        .removals
        .component_removals
        .len()
        .checked_add(formula.removals.component_removals.len())
        .and_then(|value| value.checked_add(source.state.source.components().len()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(component_removal_capacity)
        .map_err(super::map_lock_error)?;
    removed_components
        .try_reserve(component_removal_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: component_removal_capacity,
        })?;
    for removal in component_removals {
        if !removed_components.insert(removal.identifier) {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        let exact = fact_index
            .component_by_locator
            .get(removal.locator.as_ref())
            .is_some_and(|component| component.identifier == removal.identifier);
        let current_count = usize::from(
            fact_index
                .component_by_identifier
                .contains_key(&removal.identifier),
        );
        let versioned = fact_index
            .versioned_component_ids
            .contains(&removal.identifier);
        if versioned || !exact || current_count != 1 {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        if fact_index
            .unknown_external_targets
            .contains(&removal.identifier)
        {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
    }

    let mut removals = RemovalPlan::default();
    let component_count = source.state.source.components().len();
    budget
        .charge_payload_work(
            component_count
                .checked_add(object_count)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
        )
        .map_err(super::map_lock_error)?;
    let mut removed_counts: Vec<usize> = Vec::new();
    removed_counts
        .try_reserve_exact(component_count)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: component_count,
        })?;
    removed_counts.resize(component_count, 0usize);
    for removal in graph
        .removals
        .object_removals
        .iter()
        .chain(&formula.removals.object_removals)
    {
        let count = removed_counts
            .get_mut(removal.component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        *count = (*count)
            .checked_add(1)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
    }
    for (component_index, removed_count) in removed_counts.iter().copied().enumerate() {
        if removed_count == 0 {
            continue;
        }
        let component = source
            .state
            .source
            .components()
            .get_index(component_index)
            .ok_or(BodyTableDeletionError::InvalidSource)?;
        if removed_count != component.archive().objects.len() {
            continue;
        }
        let Some(metadata_component) =
            metadata_component_for_physical(&fact_index, component.name())?
        else {
            continue;
        };
        if fact_index
            .versioned_component_ids
            .contains(&metadata_component.identifier)
        {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        if fact_index
            .unknown_external_targets
            .contains(&metadata_component.identifier)
        {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        if !removed_components.insert(metadata_component.identifier) {
            continue;
        }
        let locator = clone_boxed_str(metadata_component.locator.as_ref())?;
        let name = clone_boxed_str(component.name())?;
        removals
            .component_removals
            .try_reserve(1)
            .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
        removals.component_removals.push(super::ComponentRemoval {
            component_index,
            identifier: metadata_component.identifier,
            locator,
            name,
        });
    }
    for identifier in &object_ids {
        let object = identifier.get();
        let owner_component = object_component_identifier(source, graph, *identifier, &fact_index)?;

        let uuid_matches = fact_index
            .uuids_by_object
            .get(&object)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        if owner_component.is_none() && !uuid_matches.is_empty() {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        for binding in uuid_matches {
            if owner_component != Some(binding.component) {
                return Err(BodyTableDeletionError::UnsupportedSource);
            }
        }
        if fact_index.ambiguous_by_object.contains_key(&object) {
            // The focused removal codec deliberately has no operation that
            // can preserve unknown/ambiguous field-20 spans while deleting a
            // selected object.  A component deletion later in the pipeline
            // cannot make the first registry scan safe, so refuse the source
            // regardless of which component owns the ambiguous record.
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        let current = uuid_matches.iter().filter(|binding| binding.current);
        let mut current_count = 0usize;
        let mut current_binding = None;
        for binding in current {
            current_count = current_count
                .checked_add(1)
                .ok_or(BodyTableDeletionError::InvalidSource)?;
            current_binding = Some(*binding);
        }
        if uuid_matches.iter().any(|binding| !binding.current) {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        if current_count > 1 {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        if let Some(binding) = current_binding {
            // UUID fields are removed in the first metadata pass.  Keep the
            // exact request even when the owning component registration is
            // scheduled for the subsequent whole-component removal.
            push_uuid_removal(&mut removals, binding, budget)?;
        }

        let external_matches = fact_index
            .externals_by_object
            .get(&object)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        if owner_component.is_none() && !external_matches.is_empty() {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        for reference in external_matches.iter().copied() {
            // In PackageMetadata an external reference's object identifier is
            // owned by its target component.  A source component may point at
            // that object, but it cannot claim ownership of the removal.
            if owner_component != Some(reference.target) {
                return Err(BodyTableDeletionError::UnsupportedSource);
            }
            // Registry field removals run before component registrations are
            // physically removed.  Authorize every matching current edge in
            // this first pass, including edges whose source or target
            // component will disappear in the subsequent registration pass;
            // otherwise the codec sees a surviving intermediate edge to the
            // just-removed object and correctly rejects it as cross-component.
            if !reference.current || reference.unknown {
                return Err(BodyTableDeletionError::UnsupportedSource);
            }
            push_external_removal(&mut removals, reference, budget)?;
        }

        let owner_matches = fact_index
            .data_owners_by_object
            .get(&object)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        if owner_component.is_none() && !owner_matches.is_empty() {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        for owner in owner_matches.iter().copied() {
            if owner_component != Some(owner.component) {
                return Err(BodyTableDeletionError::UnsupportedSource);
            }
            if !owner.current || owner.unknown {
                return Err(BodyTableDeletionError::UnsupportedSource);
            }
            // Data-owner fields are also visited by the ordinary metadata
            // pass before a later component-registration removal.  Authorize
            // the exact owner edge now so that intermediate scan cannot see a
            // dangling selected object.
            push_data_owner_removal(&mut removals, owner, budget)?;
        }
    }

    let watermark = released_watermark(graph, &object_ids, last, budget)?;
    if watermark.is_some() {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
    }
    Ok(MetadataPlan { removals })
}

/// Rewrite PackageMetadata while giving the caller a chance to preflight its
/// exact candidate size before the codec reserves the output buffer.  The
/// callback is allowed to run more than once when a registry rewrite is
/// followed by component-registration removal; callers that account a whole
/// physical archive can make the callback idempotent and use the first
/// candidate size as a conservative bound for the later deletion-only pass.
pub(super) fn rewrite_with_preflight(
    source: &Package,
    graph: &GraphPlan,
    formula: &FormulaPlan,
    metadata: &MetadataPlan,
    budget: &mut table_lock::WireBudget,
    mut preflight: impl FnMut(usize, &mut table_lock::WireBudget) -> Result<(), BodyTableDeletionError>,
) -> Result<Option<Vec<u8>>, BodyTableDeletionError> {
    let Some((payload, last, facts)) = inspect(source, budget)? else {
        return Ok(None);
    };
    let fact_index = build_fact_index(&facts, budget)?;

    let component_removals = graph
        .removals
        .component_removals
        .iter()
        .chain(&formula.removals.component_removals)
        .chain(&metadata.removals.component_removals);
    let component_selector_capacity = graph
        .removals
        .component_removals
        .len()
        .checked_add(formula.removals.component_removals.len())
        .and_then(|value| value.checked_add(metadata.removals.component_removals.len()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(component_selector_capacity)
        .map_err(super::map_lock_error)?;
    let mut component_selectors = Vec::new();
    component_selectors
        .try_reserve_exact(component_selector_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: component_selector_capacity,
        })?;
    let mut component_selector_keys = HashSet::<(u64, &str)>::new();
    component_selector_keys
        .try_reserve(component_selector_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: component_selector_capacity,
        })?;
    for removal in component_removals {
        let selector =
            component_selector(&fact_index, removal.identifier, removal.locator.as_ref())?;
        budget
            .charge_payload_work(
                selector
                    .locator()
                    .len()
                    .checked_add(1)
                    .ok_or(BodyTableDeletionError::InvalidSource)?,
            )
            .map_err(super::map_lock_error)?;
        if component_selector_keys.insert((selector.identifier(), selector.locator())) {
            component_selectors.push(selector);
        }
    }

    let mut uuid_removals = Vec::new();
    budget
        .charge_payload_work(metadata.removals.uuid_removals.len())
        .map_err(super::map_lock_error)?;
    uuid_removals
        .try_reserve(metadata.removals.uuid_removals.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: metadata.removals.uuid_removals.len(),
        })?;
    for removal in &metadata.removals.uuid_removals {
        let selector = component_selector_by_identifier(&fact_index, removal.component_identifier)?;
        uuid_removals.push(ObjectUuidRemoval::new(
            selector,
            removal.object_identifier.get(),
            removal.expected_uuid,
        ));
    }

    let mut external_removals = Vec::new();
    budget
        .charge_payload_work(metadata.removals.external_reference_removals.len())
        .map_err(super::map_lock_error)?;
    external_removals
        .try_reserve(metadata.removals.external_reference_removals.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: metadata.removals.external_reference_removals.len(),
        })?;
    for removal in &metadata.removals.external_reference_removals {
        let source_selector =
            component_selector_by_identifier(&fact_index, removal.source_component_identifier)?;
        let target_selector =
            component_selector_by_identifier(&fact_index, removal.target_component_identifier)?;
        external_removals.push(CodecExternalReferenceRemoval::new(
            source_selector,
            target_selector,
            removal.object_identifier.get(),
            removal.expected_is_weak,
        ));
    }

    let mut data_removals = Vec::new();
    budget
        .charge_payload_work(metadata.removals.data_reference_owner_removals.len())
        .map_err(super::map_lock_error)?;
    data_removals
        .try_reserve(metadata.removals.data_reference_owner_removals.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: metadata.removals.data_reference_owner_removals.len(),
        })?;
    let mut save_selectors = Vec::new();
    for removal in &metadata.removals.data_reference_owner_removals {
        let selector = component_selector_by_identifier(&fact_index, removal.component_identifier)?;
        data_removals.push(CodecDataReferenceOwnerRemoval::new(
            selector,
            removal.data_identifier.get(),
            removal.object_identifier.get(),
            removal.expected_count,
        ));
    }
    let save_selector_capacity = uuid_removals
        .len()
        .checked_add(external_removals.len())
        .and_then(|value| value.checked_add(data_removals.len()))
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(save_selector_capacity)
        .map_err(super::map_lock_error)?;
    save_selectors
        .try_reserve(save_selector_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: save_selector_capacity,
        })?;
    let mut save_selector_keys = HashSet::<(u64, &str)>::new();
    save_selector_keys
        .try_reserve(save_selector_capacity)
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: save_selector_capacity,
        })?;
    for selector in uuid_removals
        .iter()
        .map(|removal| removal.component())
        .chain(external_removals.iter().map(|removal| removal.source()))
        .chain(data_removals.iter().map(|removal| removal.component()))
    {
        budget
            .charge_payload_work(
                selector
                    .locator()
                    .len()
                    .checked_add(1)
                    .ok_or(BodyTableDeletionError::InvalidSource)?,
            )
            .map_err(super::map_lock_error)?;
        if save_selector_keys.insert((selector.identifier(), selector.locator())) {
            save_selectors.push(selector);
        }
    }

    let removed_ids = removed_object_ids(graph, formula, budget)?;
    let new_last = released_watermark(graph, &removed_ids, last, budget)?;
    let mut removal_batch =
        RemovalBatch::new(last, &uuid_removals, &external_removals, &data_removals);
    if let Some(new_last) = new_last {
        removal_batch = removal_batch.with_new_last_object_identifier(new_last);
    }
    let map_lookup_work = binary_search_work(removed_ids.len())?
        .checked_mul(facts.maps.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(map_lookup_work)
        .map_err(super::map_lock_error)?;
    for map in &facts.maps {
        if removed_ids
            .binary_search_by_key(&map.object, |identifier| identifier.get())
            .is_err()
        {
            continue;
        }
        if map.unknown {
            return Err(BodyTableDeletionError::UnsupportedSource);
        }
        if removal_batch.data_metadata_map().is_some() {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        removal_batch =
            removal_batch.with_data_metadata_map(DataMetadataMapRemoval::new(map.object));
    }

    let mut candidate = None;
    if !uuid_removals.is_empty()
        || !external_removals.is_empty()
        || !data_removals.is_empty()
        || removal_batch.data_metadata_map().is_some()
        || new_last.is_some()
    {
        let options = metadata_options(source, payload.len(), budget)?;
        if save_selectors.is_empty() {
            // The removal-only codec path can only delete registry fields or
            // lower the root watermark, so the source metadata payload is a
            // conservative output bound. Charge the physical archive before
            // the one-shot helper reserves its candidate buffer.
            preflight(payload.len(), budget)?;
            let output =
                package_metadata_codec::remove_package_metadata(payload, removal_batch, options)
                    .map_err(map_metadata_error)?;
            charge_codec_report(output.report(), budget)?;
            candidate = Some(output.into_bytes());
        } else {
            let batch =
                RemovalSaveTokenBatch::new(removal_batch, SaveTokenBatch::new(&save_selectors));
            let prepared =
                package_metadata_codec::prepare_package_metadata_removals_and_save_tokens(
                    payload, batch, options,
                )
                .map_err(map_metadata_error)?;
            charge_codec_report(prepared.prepare_report(), budget)?;
            let requirements = prepared.execution_requirements();
            charge_codec_requirements(requirements, budget)?;
            preflight(requirements.output_bytes(), budget)?;
            let output = prepared
                .execute(requirements.exact_limits())
                .map_err(map_metadata_error)?;
            verify_codec_execution_report(output.report(), requirements)?;
            candidate = Some(output.into_bytes());
        }
    }

    if !component_selectors.is_empty() {
        let batch = ComponentRemovalBatch::new(&component_selectors);
        let component_source = candidate.as_deref().unwrap_or(payload);
        let options = metadata_options(source, component_source.len(), budget)?;
        let prepared = package_metadata_codec::prepare_package_metadata_component_removals(
            component_source,
            batch,
            options,
        )
        .map_err(map_metadata_error)?;
        charge_codec_report(prepared.prepare_report(), budget)?;
        let requirements = prepared.execution_requirements();
        charge_codec_requirements(requirements, budget)?;
        preflight(requirements.output_bytes(), budget)?;
        let output = prepared
            .execute(requirements.exact_limits())
            .map_err(map_metadata_error)?;
        verify_codec_execution_report(output.report(), requirements)?;
        candidate = Some(output.into_bytes());
    }

    let Some(candidate) = candidate else {
        return Ok(None);
    };
    if candidate == payload {
        return Ok(None);
    }
    Ok(Some(candidate))
}

fn inspect<'a>(
    source: &'a Package,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<(&'a [u8], u64, MetadataFacts)>, BodyTableDeletionError> {
    let Some(payload) = metadata_payload(source, budget)? else {
        return Ok(None);
    };
    let options = metadata_options(source, payload.len(), budget)?;
    // The visitor owns locator strings and bounded record vectors while the
    // codec walks the source. Charge the source-sized staging envelope before
    // its first fallible reserve/copy; the exact codec report is charged after
    // inspection for the wire fields and reference counters.
    budget
        .charge_payload_work(payload.len())
        .map_err(super::map_lock_error)?;
    let mut facts = MetadataFacts::default();
    let inspection =
        package_metadata_codec::inspect_package_metadata_with_visitor(payload, options, &mut facts)
            .map_err(map_metadata_error)?;
    budget
        .charge_payload_work(inspection.report().input_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                inspection.report().fields(),
                inspection.report().work_bytes(),
                inspection.report().max_depth(),
                inspection.report().references_scanned(),
            )
        })
        .map_err(super::map_lock_error)?;
    Ok(Some((payload, inspection.last_object_identifier(), facts)))
}

fn metadata_payload<'a>(
    source: &'a Package,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<&'a [u8]>, BodyTableDeletionError> {
    let Some(component) = source.state.source.components().get("Index/Metadata.iwa") else {
        return Ok(None);
    };
    let mut location = None;
    for object in &component.archive().objects {
        budget
            .charge_payload_work(1)
            .map_err(super::map_lock_error)?;
        for message in &object.messages {
            budget
                .charge_payload_work(1)
                .map_err(super::map_lock_error)?;
            if message.type_ == METADATA_MESSAGE_TYPE {
                if location.replace(message.data.as_slice()).is_some() {
                    return Err(BodyTableDeletionError::InvalidSource);
                }
            }
        }
    }
    Ok(Some(location.ok_or(BodyTableDeletionError::InvalidSource)?))
}

fn metadata_options(
    source: &Package,
    payload_len: usize,
    budget: &table_lock::WireBudget,
) -> Result<RewriteOptions, BodyTableDeletionError> {
    let limits = source.state.source.limits();
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(super::map_archive_error)?;
    // PackageMetadata's prepared codecs scan the source several times
    // (source census, candidate sizing, execution, and verification).  A
    // single wire field can therefore be visited more than once even though
    // the logical payload is unchanged.  Use the *residual transaction*
    // ceilings for each codec invocation; the shared WireBudget still charges
    // every report cumulatively after the call.  This keeps the per-call
    // codec ceiling wide enough for its repeated pass while preserving the
    // finite aggregate limit.
    let work = budget.remaining_wire_work().max(1);
    let wire_items = budget.remaining_wire_fields().max(1);
    Ok(RewriteOptions::new(
        payload_len.max(1),
        limits.max_iwa_stream_bytes().max(payload_len).max(1),
        wire_items,
        work,
        u32::try_from(archive_limits.max_header_nesting()).unwrap_or(u32::MAX),
        wire_items,
        wire_items,
        wire_items,
    ))
}

fn component_selector<'a>(
    fact_index: &'a MetadataFactIndex<'a>,
    identifier: u64,
    locator: &str,
) -> Result<ComponentSelector<'a>, BodyTableDeletionError> {
    if fact_index.versioned_component_ids.contains(&identifier) {
        return Err(BodyTableDeletionError::UnsupportedSource);
    }
    let component = fact_index
        .component_by_locator
        .get(locator)
        .filter(|component| component.identifier == identifier)
        .copied()
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    Ok(ComponentSelector::new(
        component.identifier,
        component.locator.as_ref(),
    ))
}

fn component_selector_by_identifier<'a>(
    fact_index: &'a MetadataFactIndex<'a>,
    identifier: u64,
) -> Result<ComponentSelector<'a>, BodyTableDeletionError> {
    if fact_index.versioned_component_ids.contains(&identifier) {
        return Err(BodyTableDeletionError::UnsupportedSource);
    }
    let component = fact_index
        .component_by_identifier
        .get(&identifier)
        .copied()
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    Ok(ComponentSelector::new(
        component.identifier,
        component.locator.as_ref(),
    ))
}

fn metadata_component_for_physical<'a>(
    fact_index: &'a MetadataFactIndex<'a>,
    physical_name: &str,
) -> Result<Option<&'a MetadataComponent>, BodyTableDeletionError> {
    let locator = normalized_metadata_locator(physical_name)?;
    Ok(fact_index.component_by_locator.get(locator).copied())
}

fn clone_boxed_str(value: &str) -> Result<Box<str>, BodyTableDeletionError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_| BodyTableDeletionError::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn object_component_identifier(
    source: &Package,
    graph: &GraphPlan,
    identifier: std::num::NonZeroU64,
    fact_index: &MetadataFactIndex<'_>,
) -> Result<Option<u64>, BodyTableDeletionError> {
    let Some(location) = graph.index.get(identifier) else {
        return Err(BodyTableDeletionError::InvalidSource);
    };
    let component = source
        .state
        .source
        .components()
        .get_index(location.component_index)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let locator = normalized_metadata_locator(component.name())?;
    Ok(fact_index
        .component_by_locator
        .get(locator)
        .map(|component| component.identifier))
}

fn normalized_metadata_locator(physical_name: &str) -> Result<&str, BodyTableDeletionError> {
    physical_name
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or(BodyTableDeletionError::InvalidSource)
}

fn push_uuid_removal(
    plan: &mut RemovalPlan,
    binding: MetadataUuid,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(plan.uuid_removals.len())
        .map_err(super::map_lock_error)?;
    if plan.uuid_removals.iter().any(|existing| {
        existing.component_identifier == binding.component
            && existing.object_identifier.get() == binding.object
    }) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    plan.uuid_removals
        .try_reserve(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    plan.uuid_removals.push(UuidRemoval {
        component_identifier: binding.component,
        object_identifier: std::num::NonZeroU64::new(binding.object)
            .ok_or(BodyTableDeletionError::InvalidSource)?,
        expected_uuid: binding.uuid,
    });
    Ok(())
}

fn push_external_removal(
    plan: &mut RemovalPlan,
    reference: MetadataExternal,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(plan.external_reference_removals.len())
        .map_err(super::map_lock_error)?;
    if plan.external_reference_removals.iter().any(|existing| {
        existing.source_component_identifier == reference.source
            && existing.target_component_identifier == reference.target
            && existing.object_identifier.get() == reference.object.unwrap_or(0)
            && existing.expected_is_weak == reference.weak
    }) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    plan.external_reference_removals
        .try_reserve(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    plan.external_reference_removals
        .push(ExternalReferenceRemoval {
            source_component_identifier: reference.source,
            target_component_identifier: reference.target,
            object_identifier: std::num::NonZeroU64::new(
                reference
                    .object
                    .ok_or(BodyTableDeletionError::InvalidSource)?,
            )
            .ok_or(BodyTableDeletionError::InvalidSource)?,
            expected_is_weak: reference.weak,
        });
    Ok(())
}

fn push_data_owner_removal(
    plan: &mut RemovalPlan,
    owner: MetadataDataOwner,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(plan.data_reference_owner_removals.len())
        .map_err(super::map_lock_error)?;
    if plan.data_reference_owner_removals.iter().any(|existing| {
        existing.component_identifier == owner.component
            && existing.data_identifier.get() == owner.data
            && existing.object_identifier.get() == owner.object
    }) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    plan.data_reference_owner_removals
        .try_reserve(1)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: 1 })?;
    plan.data_reference_owner_removals
        .push(DataReferenceOwnerRemoval {
            component_identifier: owner.component,
            data_identifier: std::num::NonZeroU64::new(owner.data)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
            object_identifier: std::num::NonZeroU64::new(owner.object)
                .ok_or(BodyTableDeletionError::InvalidSource)?,
            expected_count: owner.count,
        });
    Ok(())
}

fn removed_object_ids(
    graph: &GraphPlan,
    formula: &FormulaPlan,
    budget: &mut table_lock::WireBudget,
) -> Result<Vec<std::num::NonZeroU64>, BodyTableDeletionError> {
    let count = graph
        .removals
        .object_removals
        .len()
        .checked_add(formula.removals.object_removals.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve(count)
        .map_err(|_| BodyTableDeletionError::Allocation { amount: count })?;
    for identifier in graph
        .removals
        .object_removals
        .iter()
        .chain(&formula.removals.object_removals)
        .map(|removal| removal.identifier)
    {
        identifiers.push(identifier);
    }
    charge_sort_work(identifiers.len(), budget)?;
    identifiers.sort_unstable();
    if identifiers.windows(2).any(|window| window[0] == window[1]) {
        return Err(BodyTableDeletionError::InvalidSource);
    }
    Ok(identifiers)
}

fn released_watermark(
    graph: &GraphPlan,
    removed: &[std::num::NonZeroU64],
    last: u64,
    budget: &mut table_lock::WireBudget,
) -> Result<Option<u64>, BodyTableDeletionError> {
    let lookup_work = binary_search_work(removed.len())?
        .checked_mul(graph.index.ordered.len())
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(lookup_work)
        .map_err(super::map_lock_error)?;
    if removed
        .binary_search_by_key(&last, |identifier| identifier.get())
        .is_err()
    {
        return Ok(None);
    }
    let mut maximum_remaining = 0u64;
    for location in graph.index.iter() {
        let identifier = location.identifier.get();
        let is_removed = removed
            .binary_search_by_key(&identifier, |removed| removed.get())
            .is_ok();
        if identifier > last && !is_removed {
            return Err(BodyTableDeletionError::InvalidSource);
        }
        if identifier <= last && !is_removed {
            maximum_remaining = maximum_remaining.max(identifier);
        }
    }
    if maximum_remaining == 0 {
        return Err(BodyTableDeletionError::UnsupportedSource);
    }
    Ok(Some(maximum_remaining))
}

fn charge_sort_work(
    count: usize,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    let work = count
        .checked_mul(binary_search_work(count)?)
        .ok_or(BodyTableDeletionError::InvalidSource)?;
    budget
        .charge_payload_work(work)
        .map_err(super::map_lock_error)
}

fn binary_search_work(count: usize) -> Result<usize, BodyTableDeletionError> {
    if count < 2 {
        Ok(1)
    } else {
        Ok((usize::BITS - (count - 1).leading_zeros()) as usize)
    }
}

fn charge_codec_report(
    report: package_metadata_codec::RewriteReport,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_payload_work(report.input_bytes())
        .and_then(|_| {
            budget.charge_codec_report(
                report.fields(),
                report.work_bytes(),
                report.max_depth(),
                report.references_scanned(),
            )
        })
        .map_err(super::map_lock_error)
}

fn charge_codec_requirements(
    requirements: package_metadata_codec::RewriteExecutionRequirements,
    budget: &mut table_lock::WireBudget,
) -> Result<(), BodyTableDeletionError> {
    budget
        .charge_output_bytes(requirements.output_bytes())
        .and_then(|_| budget.charge_payload_items(requirements.fields()))
        .and_then(|_| budget.charge_payload_references(requirements.references()))
        .and_then(|_| budget.charge_payload_work(requirements.work_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.components()))
        .and_then(|_| budget.charge_payload_work(requirements.allocations()))
        .and_then(|_| budget.charge_payload_work(requirements.retained_bytes()))
        .and_then(|_| budget.charge_payload_work(requirements.scratch_bytes()))
        .map_err(super::map_lock_error)
}

fn verify_codec_execution_report(
    report: package_metadata_codec::RewriteReport,
    requirements: package_metadata_codec::RewriteExecutionRequirements,
) -> Result<(), BodyTableDeletionError> {
    if report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.components_scanned() != requirements.components()
        || report.references_scanned() != requirements.references()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
        || report.scratch_bytes() > requirements.scratch_bytes()
    {
        return Err(BodyTableDeletionError::Verification);
    }
    Ok(())
}

fn map_metadata_error(error: RewriteError) -> BodyTableDeletionError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => (
                super::BodyTableDeletionLimitKind::WireBytes,
                observed,
                maximum,
            ),
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => (
                super::BodyTableDeletionLimitKind::WireOutputBytes,
                observed,
                maximum,
            ),
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => (
                super::BodyTableDeletionLimitKind::WireFields,
                observed,
                maximum,
            ),
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => (
                super::BodyTableDeletionLimitKind::WireWork,
                observed,
                maximum,
            ),
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
                return BodyTableDeletionError::LimitExceeded {
                    kind: super::BodyTableDeletionLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            package_metadata_codec::RewriteLimit::Components { observed, maximum } => (
                super::BodyTableDeletionLimitKind::PayloadItems,
                observed,
                maximum,
            ),
            package_metadata_codec::RewriteLimit::References { observed, maximum } => (
                super::BodyTableDeletionLimitKind::PayloadReferences,
                observed,
                maximum,
            ),
            package_metadata_codec::RewriteLimit::Additions { observed, maximum } => (
                super::BodyTableDeletionLimitKind::PayloadItems,
                observed,
                maximum,
            ),
            _ => return BodyTableDeletionError::InvalidSource,
        };
        return BodyTableDeletionError::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_request() {
        return BodyTableDeletionError::Allocation { amount };
    }
    match error.invalid_reason() {
        Some(package_metadata_codec::InvalidReason::VersionedRemoval)
        | Some(package_metadata_codec::InvalidReason::VersionedComponent)
        | Some(package_metadata_codec::InvalidReason::CrossComponentRemoval)
        | Some(package_metadata_codec::InvalidReason::ExistingReferenceCollision)
        | Some(package_metadata_codec::InvalidReason::RemovalMismatch) => {
            BodyTableDeletionError::UnsupportedSource
        },
        _ => BodyTableDeletionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
