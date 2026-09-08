//! Exact PackageMetadata object-UUID ownership transitions.
//!
//! The comment engine allocates UUIDs while cloning storage objects, but the
//! authoritative object-to-UUID registry lives in `PackageMetadata`.  This
//! adapter owns both sides of that narrow transition: it registers a fresh
//! candidate only when the replaced source object was registered, and removes
//! the exact old witness during graph garbage collection.  It resolves current
//! components and UUIDs through the parent metadata census, then delegates the
//! raw-preserving rewrite and candidate verification to the neutral identity
//! codec.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The helper keeps its source witness and one atomic identity rewrite together."
)]

use litchi_iwa_protos::package_metadata_codec as identity_codec;
use std::mem::size_of;

use super::super::super::Package;
use super::super::engine::WorkingSet;
use super::{
    Budget, ComponentRef, Error, MetadataCensus, ObjectUuidState, StorageUuid,
    charge_identity_report, metadata_options, metadata_payload, replace_metadata_payload,
};

#[derive(Debug)]
struct ExternalDependency {
    source: ComponentRef,
    target: ComponentRef,
    object_identifier: u64,
    expected_is_weak: Option<bool>,
}

#[derive(Debug)]
struct DataOwnerDependency {
    component: ComponentRef,
    data_identifier: u64,
    object_identifier: u64,
    expected_count: u32,
}

/// Reject physical deletion when PackageMetadata still owns one of the
/// candidate objects through a ComponentDataReference owner.
///
/// This check is separate from [`remove_storage_identity`]. A payload UUID can
/// exist without an ObjectUUIDMap registration, so that helper intentionally
/// returns a no-op in that case. Data-reference owners have no equivalent
/// optional payload witness: deleting the archive object while preserving any
/// owner would publish a dangling metadata graph. The engine calls this once
/// for the complete removal set before deleting identities or archive objects,
/// which keeps cleanup atomic and avoids rescanning PackageMetadata per object.
pub(in crate::package::slide_drawable_comments) fn reject_data_reference_owners_for_removed_objects(
    source: &Package,
    working: &WorkingSet<'_>,
    removed_identifiers: &[u64],
    budget: &mut Budget,
) -> Result<(), Error> {
    if removed_identifiers.is_empty() {
        return Ok(());
    }
    let duplicate_scan_width = removed_identifiers
        .len()
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    let identifier_work = removed_identifiers
        .len()
        .checked_mul(duplicate_scan_width)
        .ok_or(Error::InvalidSource)?;
    budget.charge_wire_work(identifier_work.max(1))?;
    for (index, identifier) in removed_identifiers.iter().copied().enumerate() {
        if identifier == 0 || removed_identifiers[..index].contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }

    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let owner_work = census
        .data_reference_owners
        .len()
        .checked_mul(removed_identifiers.len())
        .ok_or(Error::InvalidSource)?;
    if owner_work != 0 {
        budget.charge_wire_work(owner_work)?;
    }
    for owner in &census.data_reference_owners {
        if removed_identifiers.contains(&owner.object_identifier) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

/// Remove one exact current object-to-UUID registration.
///
/// The component locator, object identifier, and UUID are all source
/// witnesses.  A current cross-component collision, a duplicate, or a
/// selected record with hostile metadata fails before the working set is
/// changed.  A native payload UUID with no current PackageMetadata mapping is
/// an already-satisfied removal and returns without a rewrite.  Otherwise the
/// neutral codec preserves every unselected and unknown source span while
/// verifying the selected removal. The transaction advances save tokens once
/// after all object and metadata changes have been staged.
pub(in crate::package::slide_drawable_comments) fn remove_storage_identity(
    source: &Package,
    working: &mut WorkingSet<'_>,
    component_name: &str,
    object_identifier: u64,
    uuid: StorageUuid,
    budget: &mut Budget,
) -> Result<(), Error> {
    if component_name.is_empty()
        || object_identifier == 0
        || (uuid.lower() == 0 && uuid.upper() == 0)
    {
        return Err(Error::InvalidSource);
    }

    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let component = census.current_component(component_name, budget)?;
    let expected_uuid = identity_codec::UuidBits::new(uuid.lower(), uuid.upper());
    let state = census.object_uuid_state(&component, object_identifier, expected_uuid, budget)?;
    match state {
        ObjectUuidState::Current => {},
        // Native Keynote sources may carry the storage UUID in the comment
        // payload without registering that UUID in PackageMetadata. Once the
        // census proves that no current object or UUID collides with the
        // requested identity, the registry transition is already a no-op.
        ObjectUuidState::Missing => {
            return Ok(());
        },
        ObjectUuidState::Hostile => return Err(Error::InvalidSource),
    }

    let object_removals = [identity_codec::ObjectUuidRemoval::new(
        component.selector(),
        object_identifier,
        expected_uuid,
    )];

    // Removing an object UUID while leaving a current metadata edge or data
    // owner would make the neutral codec reject the candidate as a dangling
    // cross-component reference. Discover those dependencies from the same
    // source census and authorize their exact removal atomically with the
    // UUID. Historical, versioned, unknown, and unrelated records remain
    // source-authoritative; a selected record that cannot be removed without
    // losing its witness fails closed.
    let census_scan_items = census
        .external_references
        .len()
        .checked_add(census.data_reference_owners.len())
        .ok_or(Error::InvalidSource)?;
    let census_scan_work = census_scan_items
        .checked_mul(2)
        .ok_or(Error::InvalidSource)?;
    if census_scan_work != 0 {
        budget.charge_wire_work(census_scan_work)?;
    }
    let mut external_count = 0usize;
    let mut external_locator_bytes = 0usize;
    for (index, edge) in census.external_references.iter().enumerate() {
        if edge.object_identifier != Some(object_identifier) {
            continue;
        }
        if edge.versioned || edge.unknown_fields {
            return Err(Error::InvalidSource);
        }
        let source_witness =
            census.component_witness_by_identifier(edge.source_identifier, budget)?;
        let target_witness =
            census.component_witness_by_identifier(edge.target_identifier, budget)?;
        for prior in &census.external_references[..index] {
            budget.charge_wire_work(1)?;
            if prior.object_identifier == Some(object_identifier)
                && !prior.versioned
                && !prior.unknown_fields
                && prior.source_identifier == source_witness.identifier
                && prior.target_identifier == target_witness.identifier
            {
                return Err(Error::InvalidSource);
            }
        }
        external_count = external_count.checked_add(1).ok_or(Error::InvalidSource)?;
        external_locator_bytes = external_locator_bytes
            .checked_add(source_witness.locator.len())
            .and_then(|bytes| bytes.checked_add(target_witness.locator.len()))
            .ok_or(Error::InvalidSource)?;
    }

    let mut data_owner_count = 0usize;
    let mut data_owner_locator_bytes = 0usize;
    for (index, owner) in census.data_reference_owners.iter().enumerate() {
        if owner.object_identifier != object_identifier {
            continue;
        }
        if !owner.current || owner.unknown_fields || owner.count == 0 {
            return Err(Error::InvalidSource);
        }
        let component_witness =
            census.component_witness_by_identifier(owner.component_identifier, budget)?;
        for prior in &census.data_reference_owners[..index] {
            budget.charge_wire_work(1)?;
            if prior.current
                && !prior.unknown_fields
                && prior.component_identifier == component_witness.identifier
                && prior.data_identifier == owner.data_identifier
                && prior.object_identifier == object_identifier
            {
                return Err(Error::InvalidSource);
            }
        }
        data_owner_count = data_owner_count
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        data_owner_locator_bytes = data_owner_locator_bytes
            .checked_add(component_witness.locator.len())
            .ok_or(Error::InvalidSource)?;
    }

    let external_bytes = external_count
        .checked_mul(size_of::<ExternalDependency>())
        .and_then(|bytes| bytes.checked_add(external_locator_bytes))
        .ok_or(Error::InvalidSource)?;
    let data_owner_bytes = data_owner_count
        .checked_mul(size_of::<DataOwnerDependency>())
        .and_then(|bytes| bytes.checked_add(data_owner_locator_bytes))
        .ok_or(Error::InvalidSource)?;
    let external_request_bytes = external_count
        .checked_mul(size_of::<identity_codec::ExternalReferenceRemoval<'_>>())
        .ok_or(Error::InvalidSource)?;
    let data_owner_request_bytes = data_owner_count
        .checked_mul(size_of::<identity_codec::DataReferenceOwnerRemoval<'_>>())
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocations(
        external_bytes
            .checked_add(data_owner_bytes)
            .and_then(|bytes| bytes.checked_add(external_request_bytes))
            .and_then(|bytes| bytes.checked_add(data_owner_request_bytes))
            .ok_or(Error::InvalidSource)?,
    )?;

    let mut external_dependencies = Vec::new();
    external_dependencies
        .try_reserve_exact(external_count)
        .map_err(|_| Error::Allocation {
            amount: external_bytes,
        })?;
    for edge in &census.external_references {
        if edge.object_identifier != Some(object_identifier) {
            continue;
        }
        let source_witness =
            census.component_witness_by_identifier(edge.source_identifier, budget)?;
        let target_witness =
            census.component_witness_by_identifier(edge.target_identifier, budget)?;
        external_dependencies.push(ExternalDependency {
            source: ComponentRef::new(source_witness.identifier, &source_witness.locator),
            target: ComponentRef::new(target_witness.identifier, &target_witness.locator),
            object_identifier,
            expected_is_weak: edge.is_weak,
        });
    }

    let mut data_owner_dependencies = Vec::new();
    data_owner_dependencies
        .try_reserve_exact(data_owner_count)
        .map_err(|_| Error::Allocation {
            amount: data_owner_bytes,
        })?;
    for owner in &census.data_reference_owners {
        if owner.object_identifier != object_identifier {
            continue;
        }
        let component_witness =
            census.component_witness_by_identifier(owner.component_identifier, budget)?;
        data_owner_dependencies.push(DataOwnerDependency {
            component: ComponentRef::new(component_witness.identifier, &component_witness.locator),
            data_identifier: owner.data_identifier,
            object_identifier,
            expected_count: owner.count,
        });
    }

    let mut external_removals = Vec::new();
    external_removals
        .try_reserve_exact(external_dependencies.len())
        .map_err(|_| Error::Allocation {
            amount: external_request_bytes,
        })?;
    for dependency in &external_dependencies {
        external_removals.push(identity_codec::ExternalReferenceRemoval::new(
            dependency.source.selector(),
            dependency.target.selector(),
            dependency.object_identifier,
            dependency.expected_is_weak,
        ));
    }

    let mut data_owner_removals = Vec::new();
    data_owner_removals
        .try_reserve_exact(data_owner_dependencies.len())
        .map_err(|_| Error::Allocation {
            amount: data_owner_request_bytes,
        })?;
    for dependency in &data_owner_dependencies {
        data_owner_removals.push(identity_codec::DataReferenceOwnerRemoval::new(
            dependency.component.selector(),
            dependency.data_identifier,
            dependency.object_identifier,
            dependency.expected_count,
        ));
    }

    let batch = identity_codec::RemovalBatch::new(
        census.last_identifier,
        &object_removals,
        &external_removals,
        &data_owner_removals,
    );
    let output = identity_codec::remove_package_metadata(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    for dependency in &external_dependencies {
        let name = super::package_component_name_for_ref(source, &dependency.source, budget)?;
        working.mark_save_token_component(name, budget)?;
    }
    for dependency in &data_owner_dependencies {
        let name = super::package_component_name_for_ref(source, &dependency.component, budget)?;
        working.mark_save_token_component(name, budget)?;
    }
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    Ok(())
}

/// Register a copy-on-written storage object only when its source object had a
/// current PackageMetadata UUID witness.
///
/// Native packages may carry a storage UUID in the comment payload without
/// registering it in `ObjectUUIDMap`.  That optional state is preserved: a
/// missing source registration is a no-op.  A present registration must agree
/// with the source payload and the candidate UUID must remain unique; any
/// mismatch is rejected before the candidate metadata is replaced.
#[allow(
    clippy::too_many_arguments,
    reason = "Source and candidate identity witnesses remain explicit so callers cannot confuse their component or UUID roles."
)]
pub(in crate::package::slide_drawable_comments) fn register_storage_identity_if_registered(
    source: &Package,
    working: &mut WorkingSet<'_>,
    component_name: &str,
    old_identifier: u64,
    old_payload_uuid: Option<StorageUuid>,
    new_identifier: u64,
    new_payload_uuid: StorageUuid,
    budget: &mut Budget,
) -> Result<bool, Error> {
    if component_name.is_empty()
        || old_identifier == 0
        || new_identifier == 0
        || old_identifier == new_identifier
        || (new_payload_uuid.lower() == 0 && new_payload_uuid.upper() == 0)
    {
        return Err(Error::InvalidSource);
    }
    // Resolve the optional source registration from the immutable package.
    // The effective candidate may already have removed the old object and its
    // map entry, so consulting it first would lose the fact that this clone
    // must remain registered.
    let source_working = WorkingSet::new(source);
    let source_payload = metadata_payload(source, &source_working, budget)?;
    let source_options = metadata_options(source, source_payload.len(), budget)?;
    let source_census = MetadataCensus::inspect(&source_payload, source_options, budget)?;
    let source_component = source_census.current_component(component_name, budget)?;
    let Some(old_uuid) =
        source_census.object_uuid_for_object(&source_component, old_identifier, budget)?
    else {
        return Ok(false);
    };
    let Some(old_payload_uuid) = old_payload_uuid else {
        return Err(Error::InvalidSource);
    };
    let expected_old_uuid =
        identity_codec::UuidBits::new(old_payload_uuid.lower(), old_payload_uuid.upper());
    if old_uuid != expected_old_uuid {
        return Err(Error::InvalidSource);
    }
    let new_uuid =
        identity_codec::UuidBits::new(new_payload_uuid.lower(), new_payload_uuid.upper());
    if new_uuid.lower() == 0 && new_uuid.upper() == 0 {
        return Err(Error::InvalidSource);
    }
    let payload = metadata_payload(source, working, budget)?;
    let options = metadata_options(source, payload.len(), budget)?;
    let census = MetadataCensus::inspect(&payload, options, budget)?;
    let component = census.current_component(component_name, budget)?;
    let lookup_work = working
        .get(component_name)
        .map_or(0, |archive| archive.objects.len())
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    budget.charge_wire_work(lookup_work)?;
    let old_survives = working.find_object(old_identifier).is_some();
    let candidate_old_uuid = census.object_uuid_for_object(&component, old_identifier, budget)?;
    if old_survives {
        if candidate_old_uuid != Some(old_uuid) {
            return Err(Error::InvalidSource);
        }
    } else if candidate_old_uuid.is_some() {
        return Err(Error::InvalidSource);
    }
    match census.object_uuid_state(&component, new_identifier, new_uuid, budget)? {
        ObjectUuidState::Missing => {},
        ObjectUuidState::Current | ObjectUuidState::Hostile => {
            return Err(Error::InvalidSource);
        },
    }
    let addition = [identity_codec::ObjectUuidAddition::new(
        component.selector(),
        new_identifier,
        new_uuid,
    )];
    let new_last = census
        .last_identifier
        .checked_add(1)
        .ok_or(Error::InvalidSource)?
        .max(new_identifier);
    let batch = identity_codec::Batch::new(census.last_identifier, new_last, &addition, &[]);
    let output = identity_codec::rewrite_package_metadata(&payload, batch, options)
        .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    Ok(true)
}
