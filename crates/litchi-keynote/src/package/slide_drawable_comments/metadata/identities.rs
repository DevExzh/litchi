//! Exact PackageMetadata object-UUID ownership transitions.
//!
//! The comment engine allocates UUIDs while cloning storage objects, but the
//! authoritative object-to-UUID registry lives in `PackageMetadata`.  This
//! adapter owns the narrow inverse transition used by graph garbage
//! collection.  It resolves the current component and UUID through the
//! parent metadata census, then delegates the raw-preserving rewrite and
//! candidate verification to the neutral identity codec.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The helper keeps its source witness and one atomic identity rewrite together."
)]

use litchi_iwa_protos::package_metadata_codec as identity_codec;

use super::super::super::Package;
use super::super::engine::WorkingSet;
use super::{
    Budget, Error, MetadataCensus, ObjectUuidState, StorageUuid, charge_identity_report,
    metadata_options, metadata_payload, replace_metadata_payload,
};

/// Remove one exact current object-to-UUID registration.
///
/// The component locator, object identifier, and UUID are all source
/// witnesses.  A current cross-component collision, a duplicate, or a
/// selected record with hostile metadata fails before the working set is
/// changed.  A native payload UUID with no current PackageMetadata mapping is
/// an already-satisfied removal and returns without a rewrite.  Otherwise the
/// neutral codec preserves every unselected and unknown source span while
/// verifying the selected removal and the required current-component
/// save-token transition.
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
    match census.object_uuid_state(&component, object_identifier, expected_uuid) {
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

    let removals = [identity_codec::ObjectUuidRemoval::new(
        component.selector(),
        object_identifier,
        expected_uuid,
    )];
    let batch = identity_codec::RemovalBatch::new(census.last_identifier, &removals, &[], &[]);
    let selectors = [component.selector()];
    let output = identity_codec::rewrite_package_metadata_removals_and_save_tokens(
        &payload,
        identity_codec::RemovalSaveTokenBatch::new(
            batch,
            identity_codec::SaveTokenBatch::new(&selectors),
        ),
        options,
    )
    .map_err(|_| Error::InvalidSource)?;
    charge_identity_report(budget, output.report())?;
    replace_metadata_payload(working, output.into_bytes(), source, budget)?;
    Ok(())
}
