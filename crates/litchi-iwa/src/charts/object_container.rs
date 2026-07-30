//! Native standalone ownership for chart-private objects.
//!
//! iWork stores newly allocated geometry-only chart overrides in a dedicated
//! `TSP.ObjectContainer` component instead of the chart's object archive. This
//! module owns that component's archive, UUID map, and external references.

use std::collections::HashSet;

use prost::Message;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::package_metadata::{
    PACKAGE_METADATA_ENTRY, add_component_external_reference, add_component_object_uuids,
    add_component_registration, component_identifier_for_entry, component_uuid_identifiers,
    package_has_data_metadata_map, package_save_token,
    remove_component_external_references_to_object, remove_component_object_uuids,
    remove_component_registration,
};
use crate::protobuf::tsp;
use crate::{Error, IWorkPackage, Result};

const OBJECT_CONTAINER_MESSAGE_TYPE: u32 = 11_008;
const OBJECT_CONTAINER_LOCATOR: &str = "ObjectContainer";
const OBJECT_CONTAINER_DOCUMENT_VERSION: &[u32] = &[2, 0, 0];
const DOCUMENT_METADATA_ENTRY: &str = "Index/DocumentMetadata.iwa";
const PRIMARY_PACKAGE_IDENTIFIER: u32 = 1;

#[derive(Debug)]
pub(crate) struct ObjectContainerAllocation {
    component_id: u64,
    archive_name: String,
    locator: String,
}

pub(crate) fn reserve_object_container(
    package: &IWorkPackage,
    next_identifier: &mut u64,
) -> Result<ObjectContainerAllocation> {
    let component_id = *next_identifier;
    *next_identifier = next_identifier
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    let mut suffix: Option<u64> = None;
    loop {
        let locator = suffix.map_or_else(
            || OBJECT_CONTAINER_LOCATOR.to_owned(),
            |value| format!("{OBJECT_CONTAINER_LOCATOR}-{value}"),
        );
        let archive_name = format!("Index/{locator}.iwa");
        if !package.contains_entry(&archive_name)
            && component_identifier_for_entry(package, &archive_name)?.is_none()
        {
            return Ok(ObjectContainerAllocation {
                component_id,
                archive_name,
                locator,
            });
        }
        suffix = Some(match suffix {
            None => component_id,
            Some(value) => value.checked_add(1).ok_or_else(|| {
                Error::ParseError("iWork object-container locator overflow".to_owned())
            })?,
        });
    }
}

pub(crate) fn insert_object_container(
    package: &mut IWorkPackage,
    source_archive_name: &str,
    allocation: ObjectContainerAllocation,
    objects: Vec<ArchiveObject>,
) -> Result<()> {
    if objects.is_empty() {
        return Err(Error::InvalidFormat(
            "cannot create an empty chart object container".to_owned(),
        ));
    }
    let identifiers = objects
        .iter()
        .map(|object| {
            object.archive_info.identifier.ok_or_else(|| {
                Error::Archive("chart object container member has no identifier".to_owned())
            })
        })
        .collect::<Result<Vec<_>>>()?;
    if identifiers.iter().copied().collect::<HashSet<_>>().len() != identifiers.len() {
        return Err(Error::InvalidFormat(
            "chart object container repeats an object identifier".to_owned(),
        ));
    }
    let container = tsp::ObjectContainer {
        identifier: Some(PRIMARY_PACKAGE_IDENTIFIER),
        objects: identifiers
            .iter()
            .map(|identifier| tsp::Reference {
                identifier: *identifier,
                ..Default::default()
            })
            .collect(),
    };
    let mut archive_objects = Vec::with_capacity(objects.len() + 1);
    archive_objects.push(ArchiveObject::new(
        allocation.component_id,
        vec![RawMessage {
            type_: OBJECT_CONTAINER_MESSAGE_TYPE,
            data: container.encode_to_vec(),
        }],
    )?);
    archive_objects.extend(objects);
    let archive = Archive {
        objects: archive_objects,
    };
    let insertion_anchor = if package.contains_entry(DOCUMENT_METADATA_ENTRY) {
        DOCUMENT_METADATA_ENTRY
    } else {
        PACKAGE_METADATA_ENTRY
    };
    package.insert_archive_before(&allocation.archive_name, &archive, insertion_anchor)?;

    add_component_registration(
        package,
        &tsp::ComponentInfo {
            identifier: allocation.component_id,
            preferred_locator: allocation.locator,
            document_read_version: OBJECT_CONTAINER_DOCUMENT_VERSION.to_vec(),
            document_write_version: OBJECT_CONTAINER_DOCUMENT_VERSION.to_vec(),
            is_stored_outside_object_archive: Some(false),
            save_token: package_save_token(package)?,
            required_package_identifier: package_has_data_metadata_map(package)?
                .then_some(PRIMARY_PACKAGE_IDENTIFIER),
            ..Default::default()
        },
    )?;
    add_component_object_uuids(package, allocation.component_id, &identifiers)?;
    if package.contains_entry(PACKAGE_METADATA_ENTRY) {
        let source_component = component_identifier_for_entry(package, source_archive_name)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart component {source_archive_name} is not registered"
                ))
            })?;
        for identifier in identifiers {
            add_component_external_reference(
                package,
                source_component,
                allocation.component_id,
                identifier,
            )?;
        }
    }
    Ok(())
}

/// Remove owned objects and return the container identifier when it became empty.
pub(crate) fn remove_object_container_objects(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifiers: &[u64],
) -> Result<Option<u64>> {
    if identifiers.is_empty() {
        return Ok(None);
    }
    let requested = identifiers.iter().copied().collect::<HashSet<_>>();
    if requested.len() != identifiers.len() {
        return Err(Error::InvalidFormat(
            "chart object-container removal repeats an identifier".to_owned(),
        ));
    }
    let archive = package.archive(archive_name)?;
    let (container_id, container_message_index, mut container) =
        object_container(&archive, archive_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!("{archive_name} is not a chart object container"))
        })?;
    let contained = container
        .objects
        .iter()
        .map(|reference| reference.identifier)
        .collect::<HashSet<_>>();
    if !requested.is_subset(&contained) {
        return Err(Error::InvalidFormat(format!(
            "{archive_name} does not own every requested chart object"
        )));
    }
    container
        .objects
        .retain(|reference| !requested.contains(&reference.identifier));
    let component_id = component_identifier_for_entry(package, archive_name)?;
    if let Some(component_id) = component_id
        && component_id != container_id
    {
        return Err(Error::InvalidFormat(format!(
            "{archive_name} component {component_id} disagrees with object container {container_id}"
        )));
    }

    if container.objects.is_empty() {
        if let Some(component_id) = component_id {
            remove_component_registration(package, component_id)?;
        }
        package.remove_entry(archive_name).ok_or_else(|| {
            Error::InvalidFormat(format!("chart object container {archive_name} disappeared"))
        })?;
        return Ok(Some(container_id));
    }

    if let Some(component_id) = component_id {
        let registered = component_uuid_identifiers(package, component_id)?.unwrap_or_default();
        let registered_removals = identifiers
            .iter()
            .copied()
            .filter(|identifier| registered.contains(identifier))
            .collect::<Vec<_>>();
        remove_component_object_uuids(package, component_id, &registered_removals)?;
        for identifier in identifiers {
            remove_component_external_references_to_object(package, component_id, *identifier)?;
        }
    }
    package.update_archive(archive_name, |archive| {
        for identifier in identifiers {
            archive.remove_object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "chart object container member {identifier} is missing"
                ))
            })?;
        }
        let object = archive.object_mut(container_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "chart object container header {container_id} is missing"
            ))
        })?;
        object.replace_message(
            container_message_index,
            RawMessage {
                type_: OBJECT_CONTAINER_MESSAGE_TYPE,
                data: container.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    Ok(None)
}

pub(crate) fn is_object_container_archive(
    package: &IWorkPackage,
    archive_name: &str,
) -> Result<bool> {
    Ok(object_container(&package.archive(archive_name)?, archive_name)?.is_some())
}

fn object_container(
    archive: &Archive,
    archive_name: &str,
) -> Result<Option<(u64, usize, tsp::ObjectContainer)>> {
    let mut matches = Vec::new();
    for object in &archive.objects {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == OBJECT_CONTAINER_MESSAGE_TYPE {
                matches.push((object, message_index, message));
            }
        }
    }
    let Some((object, message_index, message)) = matches.pop() else {
        return Ok(None);
    };
    if !matches.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{archive_name} contains multiple object-container payloads"
        )));
    }
    let object_id = object.archive_info.identifier.ok_or_else(|| {
        Error::Archive(format!(
            "object-container header in {archive_name} has no identifier"
        ))
    })?;
    let container = tsp::ObjectContainer::decode(message.data.as_slice())?;
    if container.identifier != Some(PRIMARY_PACKAGE_IDENTIFIER) {
        return Err(Error::InvalidFormat(format!(
            "{archive_name} object-container package identifier is not {PRIMARY_PACKAGE_IDENTIFIER}"
        )));
    }
    let mut identifiers = HashSet::new();
    for reference in &container.objects {
        if reference.identifier == 0
            || !identifiers.insert(reference.identifier)
            || archive.object(reference.identifier).is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "{archive_name} has an invalid object-container member {}",
                reference.identifier
            )));
        }
    }
    Ok(Some((object_id, message_index, container)))
}
