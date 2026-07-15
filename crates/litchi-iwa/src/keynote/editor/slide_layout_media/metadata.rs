//! PackageMetadata accounting for materialized layout image graphs.

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};

pub(super) fn object_data_references<'a>(
    objects: impl IntoIterator<Item = &'a ArchiveObject>,
) -> Result<Vec<(u64, u64)>> {
    let mut references = Vec::new();
    for object in objects {
        let identifier = object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("Keynote layout image object has no identifier".to_owned())
        })?;
        references.extend(object.archive_info.message_infos.iter().flat_map(|info| {
            info.data_references
                .iter()
                .chain(
                    info.field_infos
                        .iter()
                        .flat_map(|field| &field.data_references),
                )
                .map(move |data| (*data, identifier))
        }));
    }
    Ok(references)
}

pub(super) fn registered_subset(
    package: &IWorkPackage,
    archive_name: &str,
    identifiers: &[u64],
) -> Result<Vec<u64>> {
    let Some(component) = component_identifier_for_entry(package, archive_name)? else {
        return Ok(Vec::new());
    };
    let registered = component_uuid_identifiers(package, component)?.unwrap_or_default();
    Ok(identifiers
        .iter()
        .copied()
        .filter(|identifier| registered.contains(identifier))
        .collect())
}

pub(super) fn remapped_registered_subset(
    package: &IWorkPackage,
    archive_name: &str,
    identifiers: &[u64],
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u64>> {
    let Some(component) = component_identifier_for_entry(package, archive_name)? else {
        return Ok(Vec::new());
    };
    let registered = component_uuid_identifiers(package, component)?.unwrap_or_default();
    Ok(identifiers
        .iter()
        .filter(|identifier| registered.contains(identifier))
        .map(|identifier| remap[identifier])
        .collect())
}

#[allow(clippy::too_many_arguments)]
pub(super) fn update_component_metadata(
    package: &mut IWorkPackage,
    slide_archive_name: &str,
    target_archive_name: &str,
    target_slide_id: u64,
    old_uuid_ids: &[u64],
    new_uuid_ids: &[u64],
    old_data_references: &[(u64, u64)],
    new_data_references: &[(u64, u64)],
) -> Result<()> {
    let Some(slide_component) = component_identifier_for_entry(package, slide_archive_name)? else {
        return Ok(());
    };
    remove_component_object_uuids(package, slide_component, old_uuid_ids)?;
    add_component_object_uuids(package, slide_component, new_uuid_ids)?;
    for &(data, object) in old_data_references {
        remove_component_data_reference(package, slide_component, data, object)?;
    }
    for &(data, object) in new_data_references {
        add_component_data_reference(package, slide_component, data, object)?;
    }
    if let Some(target_component) = component_identifier_for_entry(package, target_archive_name)?
        && target_component != slide_component
    {
        add_component_link(package, slide_component, target_component)?;
    }

    let archive = package.archive(slide_archive_name)?;
    let internal = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .collect::<HashSet<_>>();
    let external = archive
        .objects
        .iter()
        .flat_map(|object| &object.archive_info.message_infos)
        .flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        })
        .copied()
        .filter(|identifier| !internal.contains(identifier) && *identifier != target_slide_id)
        .collect::<HashSet<_>>();
    for identifier in external {
        if let Some(target_component) = component_identifier_for_object_uuid(package, identifier)?
            && target_component != slide_component
        {
            add_component_external_reference(
                package,
                slide_component,
                target_component,
                identifier,
            )?;
        }
    }
    Ok(())
}
