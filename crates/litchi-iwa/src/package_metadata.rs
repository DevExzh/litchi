//! Wire-preserving updates to the package-wide object identifier and UUID registries.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::{Archive, RawMessage};
use crate::wire::{
    append_repeated_length_delimited_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, transform_length_delimited_fields_at_path,
};
use crate::{Error, IWorkPackage, Result};

pub(crate) const PACKAGE_METADATA_ENTRY: &str = "Index/Metadata.iwa";
pub(crate) const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

pub(crate) fn next_object_identifier(package: &IWorkPackage) -> Result<u64> {
    let mut maximum = 0u64;
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive(format!("Object in {name} has no archive identifier"))
            })?;
            maximum = maximum.max(identifier);
        }
    }
    if package.contains_entry(PACKAGE_METADATA_ENTRY) {
        let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
        let (object_index, message_index) = package_metadata_location(&archive)?;
        let metadata = crate::protobuf::tsp::PackageMetadata::decode(
            archive.objects[object_index].messages[message_index]
                .data
                .as_slice(),
        )?;
        maximum = maximum.max(package_metadata_object_identifier_maximum(&metadata));
    }
    maximum
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
}

fn package_metadata_object_identifier_maximum(
    metadata: &crate::protobuf::tsp::PackageMetadata,
) -> u64 {
    let mut maximum = metadata.last_object_identifier;
    if let Some(reference) = &metadata.data_metadata_map {
        maximum = maximum.max(reference.identifier);
    }
    for component in metadata
        .components
        .iter()
        .chain(&metadata.versioned_components)
    {
        maximum = maximum.max(component.identifier);
        for entry in &component.object_uuid_map_entries {
            maximum = maximum.max(entry.identifier);
        }
        for reference in component
            .external_references
            .iter()
            .chain(&component.versioned_external_references)
        {
            if let Some(identifier) = reference.object_identifier {
                maximum = maximum.max(identifier);
            }
        }
        for reference in &component.data_references {
            for object in &reference.object_reference_list {
                maximum = maximum.max(object.object_identifier);
            }
        }
        for &identifier in &component.ambiguous_object_identifiers {
            maximum = maximum.max(identifier);
        }
    }
    maximum
}

pub(crate) fn package_last_object_identifier(package: &IWorkPackage) -> Result<Option<u64>> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(None);
    }
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let (object_index, message_index) = package_metadata_location(&archive)?;
    let metadata = crate::protobuf::tsp::PackageMetadata::decode(
        archive.objects[object_index].messages[message_index]
            .data
            .as_slice(),
    )?;
    Ok(Some(metadata.last_object_identifier))
}

pub(crate) fn set_package_last_object_identifier(
    package: &mut IWorkPackage,
    identifier: u64,
) -> Result<()> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let data = patch_varint_field(original.data.as_slice(), 1, true, Some(identifier))?;
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        if verified.last_object_identifier != identifier {
            return Err(Error::InvalidFormat(
                "PackageMetadata last object identifier patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn component_identifier_for_entry(
    package: &IWorkPackage,
    entry_name: &str,
) -> Result<Option<u64>> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(None);
    }
    let locator = entry_name
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or_else(|| Error::InvalidFormat(format!("invalid IWA component name {entry_name}")))?;
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let (object_index, message_index) = package_metadata_location(&archive)?;
    let metadata = crate::protobuf::tsp::PackageMetadata::decode(
        archive.objects[object_index].messages[message_index]
            .data
            .as_slice(),
    )?;
    let matches = metadata
        .components
        .iter()
        .filter(|component| {
            component
                .locator
                .as_deref()
                .unwrap_or(&component.preferred_locator)
                == locator
        })
        .map(|component| component.identifier)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(Error::InvalidFormat(format!(
            "PackageMetadata contains multiple components for {entry_name}"
        ))),
    }
}

pub(crate) fn component_identifier_for_object_uuid(
    package: &IWorkPackage,
    object_identifier: u64,
) -> Result<Option<u64>> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(None);
    }
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let (object_index, message_index) = package_metadata_location(&archive)?;
    let metadata = crate::protobuf::tsp::PackageMetadata::decode(
        archive.objects[object_index].messages[message_index]
            .data
            .as_slice(),
    )?;
    let matches = metadata
        .components
        .iter()
        .filter(|component| {
            component
                .object_uuid_map_entries
                .iter()
                .any(|entry| entry.identifier == object_identifier)
        })
        .map(|component| component.identifier)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(Error::InvalidFormat(format!(
            "Object {object_identifier} is registered in multiple package components"
        ))),
    }
}

pub(crate) fn clone_component_registration(
    package: &mut IWorkPackage,
    source_component_identifier: u64,
    new_component_identifier: u64,
    preferred_locator: &str,
    object_identifier_remap: &HashMap<u64, u64>,
) -> Result<()> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    if preferred_locator.is_empty() || preferred_locator.contains('/') {
        return Err(Error::ParseError(format!(
            "Invalid iWork component locator {preferred_locator:?}"
        )));
    }
    let requested_sources = object_identifier_remap
        .keys()
        .copied()
        .collect::<HashSet<_>>();
    let requested_targets = object_identifier_remap
        .values()
        .copied()
        .collect::<HashSet<_>>();
    if requested_sources.len() != object_identifier_remap.len()
        || requested_targets.len() != object_identifier_remap.len()
    {
        return Err(Error::InvalidFormat(
            "Component clone object remap must be one-to-one".to_owned(),
        ));
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let metadata = crate::protobuf::tsp::PackageMetadata::decode(original.data.as_slice())?;
        if metadata
            .components
            .iter()
            .any(|component| component.identifier == new_component_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata already contains component {new_component_identifier}"
            )));
        }
        if metadata.components.iter().any(|component| {
            component
                .locator
                .as_deref()
                .unwrap_or(&component.preferred_locator)
                == preferred_locator
        }) {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata already contains component locator {preferred_locator}"
            )));
        }
        let sources = metadata
            .components
            .iter()
            .filter(|component| component.identifier == source_component_identifier)
            .collect::<Vec<_>>();
        let [source] = sources.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata must contain exactly one source component {source_component_identifier}"
            )));
        };
        let mut existing_uuids = metadata
            .components
            .iter()
            .flat_map(|component| &component.object_uuid_map_entries)
            .map(|entry| (entry.uuid.lower, entry.uuid.upper))
            .collect::<HashSet<_>>();
        let mut cloned = (*source).clone();
        cloned.identifier = new_component_identifier;
        cloned.preferred_locator = preferred_locator.to_owned();
        cloned.locator = None;
        cloned.object_uuid_map_entries = source
            .object_uuid_map_entries
            .iter()
            .filter_map(|entry| {
                object_identifier_remap
                    .get(&entry.identifier)
                    .map(|identifier| crate::protobuf::tsp::ObjectUuidMapEntry {
                        identifier: *identifier,
                        uuid: fresh_unique_uuid(&mut existing_uuids),
                    })
            })
            .collect();
        for reference in &mut cloned.external_references {
            if let Some(identifier) = reference.object_identifier
                && let Some(replacement) = object_identifier_remap.get(&identifier)
            {
                reference.object_identifier = Some(*replacement);
            }
        }
        cloned.data_references.retain_mut(|reference| {
            reference.object_reference_list = reference
                .object_reference_list
                .iter()
                .filter_map(|object| {
                    object_identifier_remap.get(&object.object_identifier).map(
                        |identifier| crate::protobuf::tsp::component_data_reference::ObjectReference {
                            object_identifier: *identifier,
                            count: object.count,
                        },
                    )
                })
                .collect();
            !reference.object_reference_list.is_empty()
        });
        cloned.ambiguous_object_identifiers = source
            .ambiguous_object_identifiers
            .iter()
            .filter_map(|identifier| object_identifier_remap.get(identifier).copied())
            .collect();
        cloned.save_token = metadata.save_token.or(source.save_token);

        let data = append_repeated_length_delimited_field(
            original.data.as_slice(),
            3,
            &cloned.encode_to_vec(),
        )?;
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        let matches = verified
            .components
            .iter()
            .filter(|component| component.identifier == new_component_identifier)
            .collect::<Vec<_>>();
        if matches.as_slice() != [&cloned] {
            return Err(Error::InvalidFormat(
                "Package component clone failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn remove_component_registration(
    package: &mut IWorkPackage,
    component_identifier: u64,
) -> Result<()> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let metadata = crate::protobuf::tsp::PackageMetadata::decode(original.data.as_slice())?;
        if metadata
            .components
            .iter()
            .filter(|component| component.identifier == component_identifier)
            .count()
            != 1
            || metadata
                .versioned_components
                .iter()
                .any(|component| component.identifier == component_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata must contain exactly one unversioned component {component_identifier}"
            )));
        }
        let mut data = original.data.clone();
        for field in [3, 11] {
            data = transform_length_delimited_fields_at_path(&data, &[field], |component_data| {
                let mut component_data = component_data.to_vec();
                let component = crate::protobuf::tsp::ComponentInfo::decode(
                    component_data.as_slice(),
                )?;
                for (reference_field, references) in [
                    (6, component.external_references.as_slice()),
                    (18, component.versioned_external_references.as_slice()),
                ] {
                    if !references
                        .iter()
                        .any(|reference| reference.component_identifier == component_identifier)
                    {
                        continue;
                    }
                    component_data = remove_repeated_length_delimited_field_where(
                        &component_data,
                        reference_field,
                        |payload| {
                            Ok(crate::protobuf::tsp::ComponentExternalReference::decode(payload)?
                                .component_identifier
                                == component_identifier)
                        },
                    )?;
                }
                Ok(component_data)
            })?;
        }
        data = remove_repeated_length_delimited_field_where(&data, 3, |payload| {
            Ok(crate::protobuf::tsp::ComponentInfo::decode(payload)?.identifier
                == component_identifier)
        })?;
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        if verified
            .components
            .iter()
            .chain(&verified.versioned_components)
            .any(|component| {
                component.identifier == component_identifier
                    || component
                        .external_references
                        .iter()
                        .chain(&component.versioned_external_references)
                        .any(|reference| {
                            reference.component_identifier == component_identifier
                        })
            })
        {
            return Err(Error::InvalidFormat(
                "Package component removal failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn advance_package_save_token_for_components(
    package: &mut IWorkPackage,
    component_identifiers: &[u64],
) -> Result<()> {
    if component_identifiers.is_empty() || !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    let requested = component_identifiers
        .iter()
        .copied()
        .collect::<HashSet<_>>();
    if requested.len() != component_identifiers.len() {
        return Err(Error::InvalidFormat(
            "save-token update requested duplicate component identifiers".to_owned(),
        ));
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let metadata = crate::protobuf::tsp::PackageMetadata::decode(original.data.as_slice())?;
        let save_token = metadata
            .save_token
            .unwrap_or_default()
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("package save token overflow".to_owned()))?;
        let mut matched = HashSet::new();
        let data = transform_length_delimited_fields_at_path(
            original.data.as_slice(),
            &[3],
            |component_data| {
                let component = crate::protobuf::tsp::ComponentInfo::decode(component_data)?;
                if !requested.contains(&component.identifier) {
                    return Ok(component_data.to_vec());
                }
                matched.insert(component.identifier);
                patch_varint_field(
                    component_data,
                    12,
                    component.save_token.is_some(),
                    Some(save_token),
                )
            },
        )?;
        if matched != requested {
            return Err(Error::InvalidFormat(
                "PackageMetadata is missing a component requested for save-token update".to_owned(),
            ));
        }
        let data = patch_varint_field(&data, 8, metadata.save_token.is_some(), Some(save_token))?;
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        if verified.save_token != Some(save_token)
            || verified
                .components
                .iter()
                .filter(|component| requested.contains(&component.identifier))
                .any(|component| component.save_token != Some(save_token))
        {
            return Err(Error::InvalidFormat(
                "package save-token update failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn add_component_external_reference(
    package: &mut IWorkPackage,
    source_component_identifier: u64,
    target_component_identifier: u64,
    object_identifier: u64,
) -> Result<()> {
    add_component_external_reference_value(
        package,
        source_component_identifier,
        target_component_identifier,
        Some(object_identifier),
    )
}

pub(crate) fn add_component_link(
    package: &mut IWorkPackage,
    source_component_identifier: u64,
    target_component_identifier: u64,
) -> Result<()> {
    add_component_external_reference_value(
        package,
        source_component_identifier,
        target_component_identifier,
        None,
    )
}

fn add_component_external_reference_value(
    package: &mut IWorkPackage,
    source_component_identifier: u64,
    target_component_identifier: u64,
    object_identifier: Option<u64>,
) -> Result<()> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let reference = crate::protobuf::tsp::ComponentExternalReference {
            component_identifier: target_component_identifier,
            object_identifier,
            is_weak: None,
        };
        let mut source_count = 0usize;
        let mut existing_count = 0usize;
        let data = transform_length_delimited_fields_at_path(
            original.data.as_slice(),
            &[3],
            |component_data| {
                let component = crate::protobuf::tsp::ComponentInfo::decode(component_data)?;
                if component.identifier != source_component_identifier {
                    return Ok(component_data.to_vec());
                }
                source_count += 1;
                existing_count += component
                    .external_references
                    .iter()
                    .filter(|candidate| **candidate == reference)
                    .count();
                if existing_count == 0 {
                    append_repeated_length_delimited_field(
                        component_data,
                        6,
                        &reference.encode_to_vec(),
                    )
                } else {
                    Ok(component_data.to_vec())
                }
            },
        )?;
        if source_count != 1 || existing_count > 1 {
            return Err(Error::InvalidFormat(format!(
                "component {source_component_identifier} must exist once and contain at most one matching external reference"
            )));
        }
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        let count = verified
            .components
            .iter()
            .filter(|component| component.identifier == source_component_identifier)
            .flat_map(|component| &component.external_references)
            .filter(|candidate| **candidate == reference)
            .count();
        if count != 1 {
            return Err(Error::InvalidFormat(
                "component external-reference update failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn remove_component_external_references_to_object(
    package: &mut IWorkPackage,
    target_component_identifier: u64,
    object_identifier: u64,
) -> Result<()> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let data = transform_length_delimited_fields_at_path(
            original.data.as_slice(),
            &[3],
            |component_data| {
                let component = crate::protobuf::tsp::ComponentInfo::decode(component_data)?;
                let matches = component
                    .external_references
                    .iter()
                    .filter(|reference| {
                        reference.component_identifier == target_component_identifier
                            && reference.object_identifier == Some(object_identifier)
                    })
                    .count();
                if matches == 0 {
                    return Ok(component_data.to_vec());
                }
                if matches > 1 {
                    return Err(Error::InvalidFormat(format!(
                        "component {} duplicates its external reference to object {object_identifier}",
                        component.identifier
                    )));
                }
                remove_repeated_length_delimited_field_where(component_data, 6, |payload| {
                    let reference =
                        crate::protobuf::tsp::ComponentExternalReference::decode(payload)?;
                    Ok(
                        reference.component_identifier == target_component_identifier
                            && reference.object_identifier == Some(object_identifier),
                    )
                })
            },
        )?;
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        if verified
            .components
            .iter()
            .flat_map(|component| &component.external_references)
            .any(|reference| {
                reference.component_identifier == target_component_identifier
                    && reference.object_identifier == Some(object_identifier)
            })
        {
            return Err(Error::InvalidFormat(
                "component external-reference removal failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn component_uuid_identifiers(
    package: &IWorkPackage,
    component_identifier: u64,
) -> Result<Option<HashSet<u64>>> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(None);
    }
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let (object_index, message_index) = package_metadata_location(&archive)?;
    let metadata = crate::protobuf::tsp::PackageMetadata::decode(
        archive.objects[object_index].messages[message_index]
            .data
            .as_slice(),
    )?;
    let components = metadata
        .components
        .iter()
        .filter(|component| component.identifier == component_identifier)
        .collect::<Vec<_>>();
    if components.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "PackageMetadata must contain exactly one component {component_identifier}"
        )));
    }
    let mut identifiers = HashSet::new();
    for entry in &components[0].object_uuid_map_entries {
        if !identifiers.insert(entry.identifier) {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} UUID map duplicates object {}",
                entry.identifier
            )));
        }
    }
    Ok(Some(identifiers))
}

fn fresh_unique_uuid(existing: &mut HashSet<(u64, u64)>) -> crate::protobuf::tsp::Uuid {
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let mut lower = [0u8; 8];
        lower.copy_from_slice(&bytes[..8]);
        let mut upper = [0u8; 8];
        upper.copy_from_slice(&bytes[8..]);
        let uuid = crate::protobuf::tsp::Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper),
        };
        if existing.insert((uuid.lower, uuid.upper)) {
            return uuid;
        }
    }
}

pub(crate) fn add_component_object_uuids(
    package: &mut IWorkPackage,
    component_identifier: u64,
    identifiers: &[u64],
) -> Result<()> {
    if identifiers.is_empty() || !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    let requested = identifiers.iter().copied().collect::<HashSet<_>>();
    if requested.len() != identifiers.len() {
        return Err(Error::InvalidFormat(
            "UUID allocation requested duplicate object identifiers".to_owned(),
        ));
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let metadata = crate::protobuf::tsp::PackageMetadata::decode(original.data.as_slice())?;
        let mut existing_uuids = metadata
            .components
            .iter()
            .flat_map(|component| &component.object_uuid_map_entries)
            .map(|entry| (entry.uuid.lower, entry.uuid.upper))
            .collect::<HashSet<_>>();
        let conflicting = metadata
            .components
            .iter()
            .flat_map(|component| &component.object_uuid_map_entries)
            .filter_map(|entry| {
                requested
                    .contains(&entry.identifier)
                    .then_some(entry.identifier)
            })
            .collect::<Vec<_>>();
        if !conflicting.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "UUID allocation would duplicate existing object mappings {conflicting:?}"
            )));
        }
        let entries = identifiers
            .iter()
            .map(|identifier| crate::protobuf::tsp::ObjectUuidMapEntry {
                identifier: *identifier,
                uuid: fresh_unique_uuid(&mut existing_uuids),
            })
            .collect::<Vec<_>>();
        let mut component_count = 0usize;
        let data = transform_length_delimited_fields_at_path(
            original.data.as_slice(),
            &[3],
            |component_data| {
                let component = crate::protobuf::tsp::ComponentInfo::decode(component_data)?;
                if component.identifier != component_identifier {
                    return Ok(component_data.to_vec());
                }
                component_count += 1;
                entries
                    .iter()
                    .try_fold(component_data.to_vec(), |data, entry| {
                        append_repeated_length_delimited_field(&data, 11, &entry.encode_to_vec())
                    })
            },
        )?;
        if component_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata must contain exactly one component {component_identifier}"
            )));
        }
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        let mapped = verified
            .components
            .iter()
            .filter(|component| component.identifier == component_identifier)
            .flat_map(|component| &component.object_uuid_map_entries)
            .filter(|entry| requested.contains(&entry.identifier))
            .map(|entry| entry.identifier)
            .collect::<HashSet<_>>();
        if mapped != requested {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} UUID allocation failed validation"
            )));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn remove_component_object_uuids(
    package: &mut IWorkPackage,
    component_identifier: u64,
    identifiers: &[u64],
) -> Result<()> {
    if identifiers.is_empty() || !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    let requested = identifiers.iter().copied().collect::<HashSet<_>>();
    if requested.len() != identifiers.len() {
        return Err(Error::InvalidFormat(
            "UUID removal requested duplicate object identifiers".to_owned(),
        ));
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let mut component_count = 0usize;
        let data = transform_length_delimited_fields_at_path(
            original.data.as_slice(),
            &[3],
            |component_data| {
                let component = crate::protobuf::tsp::ComponentInfo::decode(component_data)?;
                if component.identifier != component_identifier {
                    return Ok(component_data.to_vec());
                }
                component_count += 1;
                identifiers
                    .iter()
                    .try_fold(component_data.to_vec(), |data, identifier| {
                        remove_repeated_length_delimited_field_where(&data, 11, |entry| {
                            Ok(
                                crate::protobuf::tsp::ObjectUuidMapEntry::decode(entry)?.identifier
                                    == *identifier,
                            )
                        })
                    })
            },
        )?;
        if component_count != 1 {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata must contain exactly one component {component_identifier}"
            )));
        }
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        if verified
            .components
            .iter()
            .filter(|component| component.identifier == component_identifier)
            .flat_map(|component| &component.object_uuid_map_entries)
            .any(|entry| requested.contains(&entry.identifier))
        {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} UUID removal failed validation"
            )));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn release_package_identifier_suffix(
    package: &mut IWorkPackage,
    removed: &[u64],
) -> Result<()> {
    let Some(mut last) = package_last_object_identifier(package)? else {
        return Ok(());
    };
    let removed = removed.iter().copied().collect::<HashSet<_>>();
    if !removed.contains(&last) {
        return Ok(());
    }
    let mut maximum_remaining = 0u64;
    for name in package.iwa_entry_names() {
        for object in package.archive(name)?.objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive(format!("Object in {name} has no archive identifier"))
            })?;
            if identifier == last {
                return Err(Error::InvalidFormat(format!(
                    "Cannot release PackageMetadata identifier {last}: the object remains"
                )));
            }
            if identifier > last {
                return Err(Error::InvalidFormat(format!(
                    "Cannot release PackageMetadata identifier suffix: object {identifier} remains"
                )));
            }
            maximum_remaining = maximum_remaining.max(identifier);
        }
    }
    last = maximum_remaining;
    set_package_last_object_identifier(package, last)
}

fn package_metadata_location(archive: &Archive) -> Result<(usize, usize)> {
    let mut location = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                && location.replace((object_index, message_index)).is_some()
            {
                return Err(Error::InvalidFormat(
                    "Package contains multiple PackageMetadata payloads".to_owned(),
                ));
            }
        }
    }
    location.ok_or_else(|| {
        Error::InvalidFormat(
            "PackageMetadata payload is missing from Index/Metadata.iwa".to_owned(),
        )
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::ArchiveObject;
    use crate::protobuf::tsp::{
        ComponentExternalReference, ComponentInfo, ObjectUuidMapEntry, PackageMetadata, Uuid,
    };

    #[test]
    fn allocator_observes_identifiers_retained_only_by_metadata_registries() {
        let metadata = PackageMetadata {
            last_object_identifier: 10,
            components: vec![ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                object_uuid_map_entries: vec![ObjectUuidMapEntry {
                    identifier: 40,
                    uuid: Uuid { lower: 1, upper: 2 },
                }],
                external_references: vec![ComponentExternalReference {
                    component_identifier: 2,
                    object_identifier: Some(50),
                    is_weak: None,
                }],
                ambiguous_object_identifiers: vec![60],
                ..Default::default()
            }],
            ..Default::default()
        };
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                PACKAGE_METADATA_ENTRY,
                &Archive {
                    objects: vec![
                        ArchiveObject::new(
                            10,
                            vec![RawMessage {
                                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                                data: metadata.encode_to_vec(),
                            }],
                        )
                        .unwrap(),
                    ],
                },
            )
            .unwrap();

        assert_eq!(next_object_identifier(&package).unwrap(), 61);
    }

    #[test]
    fn component_clone_and_remove_restore_metadata_registration() {
        let metadata = PackageMetadata {
            last_object_identifier: 20,
            components: vec![
                ComponentInfo {
                    identifier: 1,
                    preferred_locator: "Document".to_owned(),
                    object_uuid_map_entries: vec![ObjectUuidMapEntry {
                        identifier: 1,
                        uuid: Uuid { lower: 1, upper: 2 },
                    }],
                    ..Default::default()
                },
                ComponentInfo {
                    identifier: 10,
                    preferred_locator: "TemplateSlide-10".to_owned(),
                    object_uuid_map_entries: vec![ObjectUuidMapEntry {
                        identifier: 10,
                        uuid: Uuid { lower: 3, upper: 4 },
                    }],
                    external_references: vec![ComponentExternalReference {
                        component_identifier: 1,
                        object_identifier: Some(1),
                        is_weak: None,
                    }],
                    ..Default::default()
                },
            ],
            ..Default::default()
        };
        let original = metadata.encode_to_vec();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                PACKAGE_METADATA_ENTRY,
                &Archive {
                    objects: vec![
                        ArchiveObject::new(
                            20,
                            vec![RawMessage {
                                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                                data: original.clone(),
                            }],
                        )
                        .unwrap(),
                    ],
                },
            )
            .unwrap();

        clone_component_registration(
            &mut package,
            10,
            30,
            "Slide-30",
            &HashMap::from([(10, 30), (11, 31)]),
        )
        .unwrap();
        add_component_object_uuids(&mut package, 1, &[29]).unwrap();
        add_component_link(&mut package, 1, 30).unwrap();
        let cloned = PackageMetadata::decode(
            package
                .archive(PACKAGE_METADATA_ENTRY)
                .unwrap()
                .object(20)
                .unwrap()
                .messages[0]
                .data
                .as_slice(),
        )
        .unwrap();
        let component = cloned
            .components
            .iter()
            .find(|component| component.identifier == 30)
            .unwrap();
        assert_eq!(component.preferred_locator, "Slide-30");
        assert_eq!(component.object_uuid_map_entries[0].identifier, 30);

        remove_component_registration(&mut package, 30).unwrap();
        remove_component_object_uuids(&mut package, 1, &[29]).unwrap();
        assert_eq!(
            package
                .archive(PACKAGE_METADATA_ENTRY)
                .unwrap()
                .object(20)
                .unwrap()
                .messages[0]
                .data,
            original
        );
    }
}
