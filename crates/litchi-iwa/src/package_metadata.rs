//! Wire-preserving updates to the package-wide object identifier and UUID registries.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::{Archive, RawMessage};
use crate::wire::{
    append_repeated_length_delimited_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, transform_length_delimited_fields_at_path,
};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::package_metadata_codec::{
    ComponentDescriptor, DataReferenceOwnerDescriptor, ExternalReferenceDescriptor,
    ObjectUuidDescriptor, PackageMetadataInspection, PackageMetadataVisitor, RewriteError,
    RewriteLimit, RewriteOptions, inspect_package_metadata_with_visitor,
};

pub(crate) const PACKAGE_METADATA_ENTRY: &str = "Index/Metadata.iwa";
pub(crate) const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

#[derive(Default)]
struct MetadataRootVisitor {
    maximum_identifier: u64,
    has_data_metadata_map: bool,
}

impl PackageMetadataVisitor for MetadataRootVisitor {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(component.identifier());
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(binding.object_identifier());
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: ExternalReferenceDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if let Some(identifier) = reference.object_identifier() {
            self.maximum_identifier = self.maximum_identifier.max(identifier);
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: DataReferenceOwnerDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(owner.object_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: ComponentDescriptor<'_>,
        identifier: u64,
    ) -> std::result::Result<(), RewriteError> {
        self.maximum_identifier = self.maximum_identifier.max(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> std::result::Result<(), RewriteError> {
        self.has_data_metadata_map = true;
        self.maximum_identifier = self.maximum_identifier.max(object_identifier);
        Ok(())
    }
}

pub(crate) fn next_object_identifier(package: &IWorkPackage) -> Result<u64> {
    let mut maximum = 0u64;
    for name in package.iwa_entry_names() {
        package.with_parsed_archive(name, |archive| {
            for object in &archive.objects {
                let identifier = object.archive_info.identifier.ok_or_else(|| {
                    Error::Archive(format!("Object in {name} has no archive identifier"))
                })?;
                maximum = maximum.max(identifier);
            }
            Ok(())
        })?;
    }
    let mut visitor = MetadataRootVisitor::default();
    if let Some(inspection) = inspect_package_metadata(package, &mut visitor)? {
        maximum = maximum
            .max(inspection.last_object_identifier())
            .max(visitor.maximum_identifier);
    }
    maximum
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
}

pub(crate) fn package_last_object_identifier(package: &IWorkPackage) -> Result<Option<u64>> {
    inspect_package_metadata(package, &mut MetadataRootVisitor::default())
        .map(|inspection| inspection.map(PackageMetadataInspection::last_object_identifier))
}

pub(crate) fn package_save_token(package: &IWorkPackage) -> Result<Option<u64>> {
    inspect_package_metadata(package, &mut MetadataRootVisitor::default())
        .map(|inspection| inspection.and_then(PackageMetadataInspection::save_token))
}

pub(crate) fn package_has_data_metadata_map(package: &IWorkPackage) -> Result<bool> {
    let mut visitor = MetadataRootVisitor::default();
    let _inspection = inspect_package_metadata(package, &mut visitor)?;
    Ok(visitor.has_data_metadata_map)
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
    let locator = entry_name
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or_else(|| Error::InvalidFormat(format!("invalid IWA component name {entry_name}")))?;
    with_package_metadata_payload(package, |source| {
        let mut visitor = ComponentLocatorVisitor {
            locator,
            first_match: None,
            matches: 0,
        };
        inspect_package_metadata_payload(package, source, &mut visitor)?;
        match visitor.matches {
            0 => Ok(None),
            1 => Ok(visitor.first_match),
            _ => Err(Error::InvalidFormat(format!(
                "PackageMetadata contains multiple components for {entry_name}"
            ))),
        }
    })
    .map(|value| value.flatten())
}

pub(crate) fn component_identifier_for_object_uuid(
    package: &IWorkPackage,
    object_identifier: u64,
) -> Result<Option<u64>> {
    with_package_metadata_payload(package, |source| {
        let mut visitor = ObjectUuidOwnerVisitor {
            object_identifier,
            current_component: None,
            current_component_matches: false,
            first_match: None,
            matches: 0,
        };
        inspect_package_metadata_payload(package, source, &mut visitor)?;
        match visitor.matches {
            0 => Ok(None),
            1 => Ok(visitor.first_match),
            _ => Err(Error::InvalidFormat(format!(
                "Object {object_identifier} is registered in multiple package components"
            ))),
        }
    })
    .map(|value| value.flatten())
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

/// Append a fully specified unversioned component registration.
///
/// The caller is responsible for assigning fresh object UUIDs. Existing
/// component identifiers, locators, and UUID values are rejected so malformed
/// registries cannot be created accidentally.
pub(crate) fn add_component_registration(
    package: &mut IWorkPackage,
    component: &crate::protobuf::tsp::ComponentInfo,
) -> Result<()> {
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(());
    }
    let locator = component
        .locator
        .as_deref()
        .unwrap_or(&component.preferred_locator);
    if component.identifier == 0 || locator.is_empty() || locator.contains('/') {
        return Err(Error::ParseError(format!(
            "Invalid iWork component registration {} at {locator:?}",
            component.identifier
        )));
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let metadata = crate::protobuf::tsp::PackageMetadata::decode(original.data.as_slice())?;
        if metadata
            .components
            .iter()
            .chain(&metadata.versioned_components)
            .any(|existing| {
                existing.identifier == component.identifier
                    || existing
                        .locator
                        .as_deref()
                        .unwrap_or(&existing.preferred_locator)
                        == locator
            })
        {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata already contains component {} or locator {locator}",
                component.identifier
            )));
        }
        let existing_uuids = metadata
            .components
            .iter()
            .chain(&metadata.versioned_components)
            .flat_map(|existing| &existing.object_uuid_map_entries)
            .map(|entry| (entry.uuid.lower, entry.uuid.upper))
            .collect::<HashSet<_>>();
        let mut requested_uuids = HashSet::new();
        for entry in &component.object_uuid_map_entries {
            let uuid = (entry.uuid.lower, entry.uuid.upper);
            if existing_uuids.contains(&uuid) || !requested_uuids.insert(uuid) {
                return Err(Error::InvalidFormat(
                    "Package component registration repeats an object UUID".to_owned(),
                ));
            }
        }

        let data = append_repeated_length_delimited_field(
            original.data.as_slice(),
            3,
            &component.encode_to_vec(),
        )?;
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        let matches = verified
            .components
            .iter()
            .filter(|existing| *existing == component)
            .count();
        if matches != 1 {
            return Err(Error::InvalidFormat(
                "Package component insertion failed validation".to_owned(),
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
                    for reference in references
                        .iter()
                        .filter(|reference| {
                            reference.component_identifier == component_identifier
                        })
                    {
                        component_data = remove_repeated_length_delimited_field_where(
                            &component_data,
                            reference_field,
                            |payload| {
                                Ok(
                                    crate::protobuf::tsp::ComponentExternalReference::decode(
                                        payload,
                                    )? == *reference,
                                )
                            },
                        )?;
                    }
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

pub(crate) fn remove_component_external_reference(
    package: &mut IWorkPackage,
    source_component_identifier: u64,
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
        let mut source_count = 0usize;
        let mut match_count = 0usize;
        let data = transform_length_delimited_fields_at_path(
            original.data.as_slice(),
            &[3],
            |component_data| {
                let component = crate::protobuf::tsp::ComponentInfo::decode(component_data)?;
                if component.identifier != source_component_identifier {
                    return Ok(component_data.to_vec());
                }
                source_count += 1;
                let matches = component
                    .external_references
                    .iter()
                    .filter(|reference| {
                        reference.component_identifier == target_component_identifier
                            && reference.object_identifier == Some(object_identifier)
                    })
                    .count();
                match_count += matches;
                if matches > 1 {
                    return Err(Error::InvalidFormat(format!(
                        "component {source_component_identifier} duplicates its external reference to object {object_identifier}"
                    )));
                }
                if matches == 0 {
                    return Ok(component_data.to_vec());
                }
                remove_repeated_length_delimited_field_where(component_data, 6, |payload| {
                    let reference =
                        crate::protobuf::tsp::ComponentExternalReference::decode(payload)?;
                    Ok(reference.component_identifier == target_component_identifier
                        && reference.object_identifier == Some(object_identifier))
                })
            },
        )?;
        if source_count != 1 || match_count > 1 {
            return Err(Error::InvalidFormat(format!(
                "component {source_component_identifier} must exist once and contain at most one matching external reference"
            )));
        }
        let verified = crate::protobuf::tsp::PackageMetadata::decode(data.as_slice())?;
        if verified
            .components
            .iter()
            .find(|component| component.identifier == source_component_identifier)
            .is_none_or(|component| {
                component.external_references.iter().any(|reference| {
                    reference.component_identifier == target_component_identifier
                        && reference.object_identifier == Some(object_identifier)
                })
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
    with_package_metadata_payload(package, |source| {
        let mut visitor = ComponentUuidVisitor {
            component_identifier,
            components: 0,
            identifiers: HashSet::new(),
            duplicate_identifier: None,
        };
        inspect_package_metadata_payload(package, source, &mut visitor)?;
        if visitor.components != 1 {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata must contain exactly one component {component_identifier}"
            )));
        }
        if let Some(identifier) = visitor.duplicate_identifier {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} UUID map duplicates object {identifier}"
            )));
        }
        Ok(visitor.identifiers)
    })
}

const PACKAGE_METADATA_RECURSION_LIMIT: u32 = 64;

pub(crate) fn package_metadata_read_options(package: &IWorkPackage) -> RewriteOptions {
    let package_limits = package.limits();
    let archive_limits = package_limits.archive_limits();
    let message_limit = package_limits
        .max_iwa_stream_bytes()
        .min(archive_limits.max_message_bytes())
        .max(1);
    let input_limit = message_limit.min(WireLimits::MAX_INPUT_BYTES);
    let output_limit = message_limit.min(WireLimits::MAX_OUTPUT_BYTES);
    let record_limit = WireLimits::MAX_FIELDS.min(input_limit);
    RewriteOptions::new(
        input_limit,
        output_limit,
        WireLimits::MAX_FIELDS,
        WireLimits::MAX_REWRITE_WORK,
        PACKAGE_METADATA_RECURSION_LIMIT.min(WireLimits::MAX_NESTING as u32),
        record_limit,
        record_limit,
        0,
    )
}

fn inspect_package_metadata_payload<V: PackageMetadataVisitor>(
    package: &IWorkPackage,
    source: &[u8],
    visitor: &mut V,
) -> Result<()> {
    inspect_package_metadata_source(source, package_metadata_read_options(package), visitor)
        .map(|_inspection| ())
}

pub(crate) fn inspect_package_metadata_source<V: PackageMetadataVisitor>(
    source: &[u8],
    options: RewriteOptions,
    visitor: &mut V,
) -> Result<PackageMetadataInspection> {
    inspect_package_metadata_with_visitor(source, options, visitor)
        .map_err(package_metadata_inspection_error)
}

pub(crate) fn inspect_package_metadata<V: PackageMetadataVisitor>(
    package: &IWorkPackage,
    visitor: &mut V,
) -> Result<Option<PackageMetadataInspection>> {
    with_package_metadata_payload(package, |source| {
        inspect_package_metadata_source(source, package_metadata_read_options(package), visitor)
    })
}

fn package_metadata_inspection_error(error: RewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            RewriteLimit::InputBytes { observed, maximum } => {
                (LimitKind::InputBytes, observed, maximum)
            },
            RewriteLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed, maximum)
            },
            RewriteLimit::Fields { observed, maximum } => (LimitKind::Fields, observed, maximum),
            RewriteLimit::Work { observed, maximum } => (LimitKind::RewriteWork, observed, maximum),
            RewriteLimit::Nesting { observed, maximum } => {
                (LimitKind::Nesting, observed as usize, maximum as usize)
            },
            RewriteLimit::Components { observed, maximum } => {
                return Error::InvalidFormat(format!(
                    "PackageMetadata component inspection limit exceeded: observed {observed}, limit {maximum}"
                ));
            },
            RewriteLimit::References { observed, maximum } => {
                return Error::InvalidFormat(format!(
                    "PackageMetadata reference inspection limit exceeded: observed {observed}, limit {maximum}"
                ));
            },
            RewriteLimit::Additions { observed, maximum } => {
                return Error::InvalidFormat(format!(
                    "PackageMetadata inspection unexpectedly exceeded its addition limit: observed {observed}, limit {maximum}"
                ));
            },
            _ => {
                return Error::InvalidFormat(format!("PackageMetadata inspection failed: {error}"));
            },
        };
        return Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit: maximum,
        });
    }
    if let Some(amount) = error.allocation_request() {
        return Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "PackageMetadata visitor staging",
            amount,
        });
    }
    match error.invalid_reason() {
        Some(reason) => Error::InvalidFormat(format!(
            "PackageMetadata strict inspection failed: {reason:?}"
        )),
        None => Error::InvalidFormat(format!("PackageMetadata strict inspection failed: {error}")),
    }
}

fn with_package_metadata_payload<T, F>(package: &IWorkPackage, read: F) -> Result<Option<T>>
where
    F: FnOnce(&[u8]) -> Result<T>,
{
    if !package.contains_entry(PACKAGE_METADATA_ENTRY) {
        return Ok(None);
    }
    package.with_parsed_archive(PACKAGE_METADATA_ENTRY, |archive| {
        let (object_index, message_index) = package_metadata_location(archive)?;
        read(
            archive.objects[object_index].messages[message_index]
                .data
                .as_slice(),
        )
        .map(Some)
    })
}

struct ComponentLocatorVisitor<'source> {
    locator: &'source str,
    first_match: Option<u64>,
    matches: usize,
}

impl PackageMetadataVisitor for ComponentLocatorVisitor<'_> {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if component.is_current() && component.effective_locator() == self.locator {
            self.matches = self
                .matches
                .checked_add(1)
                .ok_or_else(|| RewriteError::allocation(usize::MAX))?;
            self.first_match.get_or_insert(component.identifier());
        }
        Ok(())
    }
}

struct ObjectUuidOwnerVisitor {
    object_identifier: u64,
    current_component: Option<u64>,
    current_component_matches: bool,
    first_match: Option<u64>,
    matches: usize,
}

impl PackageMetadataVisitor for ObjectUuidOwnerVisitor {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        self.current_component = component.is_current().then_some(component.identifier());
        self.current_component_matches = false;
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        let component = binding.component();
        if component.is_current()
            && self.current_component.is_some()
            && !self.current_component_matches
            && binding.object_identifier() == self.object_identifier
        {
            self.current_component_matches = true;
            self.matches = self
                .matches
                .checked_add(1)
                .ok_or_else(|| RewriteError::allocation(usize::MAX))?;
            self.first_match.get_or_insert(component.identifier());
        }
        Ok(())
    }
}

struct ComponentUuidVisitor {
    component_identifier: u64,
    components: usize,
    identifiers: HashSet<u64>,
    duplicate_identifier: Option<u64>,
}

impl PackageMetadataVisitor for ComponentUuidVisitor {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        if component.is_current() && component.identifier() == self.component_identifier {
            self.components = self
                .components
                .checked_add(1)
                .ok_or_else(|| RewriteError::allocation(usize::MAX))?;
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: ObjectUuidDescriptor<'_>,
    ) -> std::result::Result<(), RewriteError> {
        let component = binding.component();
        if component.is_current() && component.identifier() == self.component_identifier {
            let identifier = binding.object_identifier();
            if self.identifiers.contains(&identifier) {
                self.duplicate_identifier.get_or_insert(identifier);
            } else {
                let requested = self
                    .identifiers
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| RewriteError::allocation(usize::MAX))?;
                self.identifiers
                    .try_reserve(1)
                    .map_err(|_error| RewriteError::allocation(requested))?;
                self.identifiers.insert(identifier);
            }
        }
        Ok(())
    }
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
        package.with_parsed_archive(name, |archive| {
            for object in &archive.objects {
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
            Ok(())
        })?;
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
        ComponentDataReference, ComponentExternalReference, ComponentInfo, ObjectUuidMapEntry,
        PackageMetadata, Reference, Uuid, component_data_reference,
    };

    fn package_with_metadata(metadata: PackageMetadata) -> IWorkPackage {
        package_with_metadata_data(metadata.encode_to_vec())
    }

    fn package_with_metadata_data(data: Vec<u8>) -> IWorkPackage {
        package_with_metadata_messages(vec![RawMessage {
            type_: PACKAGE_METADATA_MESSAGE_TYPE,
            data,
        }])
    }

    fn package_with_metadata_messages(messages: Vec<RawMessage>) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                PACKAGE_METADATA_ENTRY,
                &Archive {
                    objects: vec![ArchiveObject::new(10, messages).unwrap()],
                },
            )
            .unwrap();
        package
    }

    fn assert_component_queries_fail(package: &IWorkPackage) {
        assert!(component_identifier_for_entry(package, "Index/One.iwa").is_err());
        assert!(component_identifier_for_object_uuid(package, 1).is_err());
        assert!(component_uuid_identifiers(package, 1).is_err());
    }

    #[test]
    fn metadata_scalar_reads_stream_every_identifier_namespace() {
        let metadata = PackageMetadata {
            last_object_identifier: 10,
            save_token: Some(42),
            data_metadata_map: Some(Reference {
                identifier: 90,
                ..Default::default()
            }),
            components: vec![ComponentInfo {
                identifier: 11,
                preferred_locator: "Document".to_owned(),
                object_uuid_map_entries: vec![ObjectUuidMapEntry {
                    identifier: 40,
                    uuid: Uuid { lower: 1, upper: 2 },
                }],
                external_references: vec![ComponentExternalReference {
                    component_identifier: 12,
                    object_identifier: Some(50),
                    is_weak: None,
                }],
                data_references: vec![ComponentDataReference {
                    data_identifier: 70,
                    object_reference_list: vec![component_data_reference::ObjectReference {
                        object_identifier: 95,
                        count: 1,
                    }],
                }],
                ambiguous_object_identifiers: vec![80],
                ..Default::default()
            }],
            versioned_components: vec![ComponentInfo {
                identifier: 97,
                preferred_locator: "Versioned".to_owned(),
                ..Default::default()
            }],
            ..Default::default()
        };
        let source = metadata.encode_to_vec();
        let package = package_with_metadata_data(source.clone());

        assert_eq!(package_last_object_identifier(&package).unwrap(), Some(10));
        assert_eq!(package_save_token(&package).unwrap(), Some(42));
        assert!(package_has_data_metadata_map(&package).unwrap());
        assert_eq!(next_object_identifier(&package).unwrap(), 98);
        assert_eq!(
            package
                .archive(PACKAGE_METADATA_ENTRY)
                .unwrap()
                .object(10)
                .unwrap()
                .messages[0]
                .data,
            source
        );
    }

    #[test]
    fn metadata_scalar_reads_preserve_absence_and_reject_duplicate_tokens() {
        let empty = IWorkPackage::new();
        assert_eq!(package_last_object_identifier(&empty).unwrap(), None);
        assert_eq!(package_save_token(&empty).unwrap(), None);
        assert!(!package_has_data_metadata_map(&empty).unwrap());

        let package = package_with_metadata(PackageMetadata {
            last_object_identifier: 10,
            ..Default::default()
        });
        assert_eq!(package_save_token(&package).unwrap(), None);
        assert!(!package_has_data_metadata_map(&package).unwrap());

        let explicit_zero = package_with_metadata(PackageMetadata {
            last_object_identifier: 10,
            save_token: Some(0),
            ..Default::default()
        });
        assert_eq!(package_save_token(&explicit_zero).unwrap(), Some(0));

        let mut unknown_scalar = PackageMetadata {
            last_object_identifier: 10,
            save_token: Some(9),
            ..Default::default()
        }
        .encode_to_vec();
        unknown_scalar.extend_from_slice(&[0xd0, 0x05, 0x07]);
        let unknown = package_with_metadata_data(unknown_scalar.clone());
        assert_eq!(package_last_object_identifier(&unknown).unwrap(), Some(10));
        assert_eq!(package_save_token(&unknown).unwrap(), Some(9));
        assert_eq!(
            unknown
                .archive(PACKAGE_METADATA_ENTRY)
                .unwrap()
                .object(10)
                .unwrap()
                .messages[0]
                .data,
            unknown_scalar
        );

        let mut duplicate_token = PackageMetadata {
            last_object_identifier: 10,
            save_token: Some(7),
            ..Default::default()
        }
        .encode_to_vec();
        duplicate_token.extend_from_slice(&[0x40, 0x08]);
        let malformed = package_with_metadata_data(duplicate_token);
        assert!(package_last_object_identifier(&malformed).is_err());
        assert!(package_save_token(&malformed).is_err());
        assert!(package_has_data_metadata_map(&malformed).is_err());
        assert!(next_object_identifier(&malformed).is_err());
    }

    #[test]
    fn component_queries_use_locator_precedence_and_ignore_versioned_components() {
        let package = package_with_metadata(PackageMetadata {
            last_object_identifier: 10,
            components: vec![
                ComponentInfo {
                    identifier: 1,
                    preferred_locator: "Preferred".to_owned(),
                    locator: Some("Actual".to_owned()),
                    ..Default::default()
                },
                ComponentInfo {
                    identifier: 3,
                    preferred_locator: "Fallback".to_owned(),
                    object_uuid_map_entries: vec![
                        ObjectUuidMapEntry {
                            identifier: 42,
                            uuid: Uuid { lower: 3, upper: 4 },
                        },
                        ObjectUuidMapEntry {
                            identifier: 43,
                            uuid: Uuid { lower: 5, upper: 6 },
                        },
                    ],
                    ..Default::default()
                },
            ],
            versioned_components: vec![ComponentInfo {
                identifier: 2,
                preferred_locator: "VersionedPreferred".to_owned(),
                locator: Some("VersionedActual".to_owned()),
                object_uuid_map_entries: vec![ObjectUuidMapEntry {
                    identifier: 77,
                    uuid: Uuid { lower: 1, upper: 2 },
                }],
                ..Default::default()
            }],
            ..Default::default()
        });

        assert_eq!(
            component_identifier_for_entry(&package, "Index/Actual.iwa").unwrap(),
            Some(1)
        );
        assert_eq!(
            component_identifier_for_entry(&package, "Index/Preferred.iwa").unwrap(),
            None
        );
        assert_eq!(
            component_identifier_for_entry(&package, "Index/Fallback.iwa").unwrap(),
            Some(3)
        );
        assert_eq!(
            component_identifier_for_entry(&package, "Index/VersionedActual.iwa").unwrap(),
            None
        );
        assert_eq!(
            component_identifier_for_object_uuid(&package, 77).unwrap(),
            None
        );
        assert_eq!(
            component_identifier_for_object_uuid(&package, 42).unwrap(),
            Some(3)
        );
        assert_eq!(
            component_uuid_identifiers(&package, 1).unwrap(),
            Some(HashSet::new())
        );
        assert_eq!(
            component_uuid_identifiers(&package, 3).unwrap(),
            Some(HashSet::from([42, 43]))
        );
        assert!(component_uuid_identifiers(&package, 2).is_err());
    }

    #[test]
    fn component_queries_reject_duplicate_locator_and_owner_entries() {
        let package = package_with_metadata(PackageMetadata {
            last_object_identifier: 10,
            components: vec![
                ComponentInfo {
                    identifier: 1,
                    preferred_locator: "Shared".to_owned(),
                    object_uuid_map_entries: vec![ObjectUuidMapEntry {
                        identifier: 77,
                        uuid: Uuid { lower: 1, upper: 2 },
                    }],
                    ..Default::default()
                },
                ComponentInfo {
                    identifier: 2,
                    preferred_locator: "Shared".to_owned(),
                    object_uuid_map_entries: vec![ObjectUuidMapEntry {
                        identifier: 77,
                        uuid: Uuid { lower: 3, upper: 4 },
                    }],
                    ..Default::default()
                },
            ],
            ..Default::default()
        });

        assert!(component_identifier_for_entry(&package, "Index/Shared.iwa").is_err());
        assert!(component_identifier_for_object_uuid(&package, 77).is_err());
    }

    #[test]
    fn component_uuid_queries_reject_duplicate_components_and_uuid_entries() {
        let duplicate_components = package_with_metadata(PackageMetadata {
            last_object_identifier: 10,
            components: vec![
                ComponentInfo {
                    identifier: 1,
                    preferred_locator: "One".to_owned(),
                    ..Default::default()
                },
                ComponentInfo {
                    identifier: 1,
                    preferred_locator: "One-again".to_owned(),
                    ..Default::default()
                },
            ],
            versioned_components: vec![ComponentInfo {
                identifier: 1,
                preferred_locator: "Old".to_owned(),
                object_uuid_map_entries: vec![ObjectUuidMapEntry {
                    identifier: 99,
                    uuid: Uuid { lower: 5, upper: 6 },
                }],
                ..Default::default()
            }],
            ..Default::default()
        });
        assert!(component_uuid_identifiers(&duplicate_components, 1).is_err());

        let duplicate_uuid = package_with_metadata(PackageMetadata {
            last_object_identifier: 10,
            components: vec![ComponentInfo {
                identifier: 1,
                preferred_locator: "One".to_owned(),
                object_uuid_map_entries: vec![
                    ObjectUuidMapEntry {
                        identifier: 42,
                        uuid: Uuid { lower: 1, upper: 2 },
                    },
                    ObjectUuidMapEntry {
                        identifier: 42,
                        uuid: Uuid { lower: 3, upper: 4 },
                    },
                ],
                ..Default::default()
            }],
            ..Default::default()
        });
        assert_eq!(
            component_identifier_for_object_uuid(&duplicate_uuid, 42).unwrap(),
            Some(1)
        );
        assert!(component_uuid_identifiers(&duplicate_uuid, 1).is_err());
    }

    #[test]
    fn component_queries_preserve_missing_metadata_and_reject_malformed_wire() {
        let package = IWorkPackage::new();
        assert_eq!(
            component_identifier_for_entry(&package, "Index/One.iwa").unwrap(),
            None
        );
        assert_eq!(
            component_identifier_for_object_uuid(&package, 1).unwrap(),
            None
        );
        assert_eq!(component_uuid_identifiers(&package, 1).unwrap(), None);

        let missing_payload = package_with_metadata_messages(vec![RawMessage {
            type_: 1,
            data: Vec::new(),
        }]);
        assert_component_queries_fail(&missing_payload);

        let valid = PackageMetadata {
            last_object_identifier: 10,
            ..Default::default()
        }
        .encode_to_vec();
        let duplicate_payloads = package_with_metadata_messages(vec![
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: valid.clone(),
            },
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: valid,
            },
        ]);
        assert_component_queries_fail(&duplicate_payloads);

        for data in [
            vec![0x08, 0x80],
            vec![0x08, 0x00],
            vec![0x1a, 0x02, 0x08, 0x01],
            vec![0x08, 0x0a, 0x1a, 0x02, 0x08, 0x01],
            vec![0x08, 0x8a, 0x00],
            vec![0x08, 0x0a, 0x08, 0x0a],
            vec![0x0a, 0x01, 0x0a],
            vec![0x08, 0x0a, 0x1a, 0x05, 0x08, 0x01, 0x12, 0x01, 0xff],
            vec![
                0x08, 0x0a, 0x1a, 0x0b, 0x08, 0x01, 0x12, 0x01, 0x41, 0x32, 0x04, 0x08, 0x02, 0x18,
                0x02,
            ],
            vec![0x08, 0x0a, 0x5a, 0x02, 0x08, 0x01],
        ] {
            let malformed = package_with_metadata_data(data);
            assert_component_queries_fail(&malformed);
        }
    }

    #[test]
    fn component_query_inspection_honors_finite_field_limits() {
        let metadata = PackageMetadata {
            last_object_identifier: 10,
            components: vec![ComponentInfo {
                identifier: 1,
                preferred_locator: "One".to_owned(),
                ..Default::default()
            }],
            ..Default::default()
        };
        let source = metadata.encode_to_vec();
        let mut visitor = ComponentLocatorVisitor {
            locator: "One",
            first_match: None,
            matches: 0,
        };
        let options = RewriteOptions::new(
            source.len(),
            source.len(),
            1,
            WireLimits::MAX_REWRITE_WORK,
            PACKAGE_METADATA_RECURSION_LIMIT,
            source.len(),
            source.len(),
            0,
        );
        let error = inspect_package_metadata_with_visitor(&source, options, &mut visitor)
            .expect_err("the deliberately tiny field budget must fail closed");
        assert!(error.resource_limit().is_some());
        assert!(matches!(
            package_metadata_inspection_error(error),
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                ..
            })
        ));
        assert!(matches!(
            package_metadata_inspection_error(RewriteError::allocation(3)),
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "PackageMetadata visitor staging",
                amount: 3,
            })
        ));
    }

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
    fn removing_absent_component_external_reference_is_an_exact_no_op() {
        let metadata = PackageMetadata {
            components: vec![ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                external_references: vec![ComponentExternalReference {
                    component_identifier: 2,
                    object_identifier: Some(40),
                    is_weak: None,
                }],
                ..Default::default()
            }],
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
                            10,
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

        remove_component_external_reference(&mut package, 1, 3, 41).unwrap();

        assert_eq!(
            package
                .archive(PACKAGE_METADATA_ENTRY)
                .unwrap()
                .object(10)
                .unwrap()
                .messages[0]
                .data,
            original
        );
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
