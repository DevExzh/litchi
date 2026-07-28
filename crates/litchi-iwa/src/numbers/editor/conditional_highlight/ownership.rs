//! Reference validation and component-aware removal for owned style objects.

use std::collections::{HashMap, HashSet};

use super::*;

pub(super) fn ensure_children_are_private(
    package: &IWorkPackage,
    owner_identifier: u64,
    children: &[u64],
) -> Result<()> {
    if children.is_empty() {
        return Ok(());
    }
    let children = children.iter().copied().collect::<HashSet<_>>();
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            if object.archive_info.identifier == Some(owner_identifier) {
                continue;
            }
            for info in &object.archive_info.message_infos {
                let shared = info
                    .object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .find(|identifier| children.contains(identifier));
                if let Some(identifier) = shared {
                    return Err(Error::InvalidFormat(format!(
                        "conditional-highlight child {identifier} is shared by another object"
                    )));
                }
            }
        }
    }
    Ok(())
}

pub(super) fn remove_owned_objects(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifiers: &[u64],
) -> Result<()> {
    let mut by_component = HashMap::<u64, Vec<u64>>::new();
    for identifier in identifiers {
        let archive_name = locations.get(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "conditional-highlight object {identifier} is missing"
            ))
        })?;
        if let Some(component) = component_identifier_for_entry(package, archive_name)? {
            remove_component_external_references_to_object(package, component, *identifier)?;
            by_component.entry(component).or_default().push(*identifier);
        }
    }
    for (component, component_identifiers) in by_component {
        let registered = component_uuid_identifiers(package, component)?;
        let registered = component_identifiers
            .into_iter()
            .filter(|identifier| {
                registered
                    .as_ref()
                    .is_some_and(|registered| registered.contains(identifier))
            })
            .collect::<Vec<_>>();
        remove_component_object_uuids(package, component, &registered)?;
    }
    for identifier in identifiers {
        remove_object_or_empty_entry(package, locations, *identifier)?;
    }
    Ok(())
}
