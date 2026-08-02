//! Wire-preserving `PackageMetadata` component data-reference accounting.

use std::collections::{HashMap, HashSet};

use prost::Message;

use crate::archive::RawMessage;
use crate::package_metadata::{PACKAGE_METADATA_ENTRY, PACKAGE_METADATA_MESSAGE_TYPE};
use crate::protobuf::tsp::{
    ComponentDataReference, ComponentInfo, PackageMetadata,
    component_data_reference::ObjectReference,
};
use crate::wire::{
    append_repeated_length_delimited_field, patch_varint_field,
    remove_repeated_length_delimited_field_where, transform_length_delimited_fields_at_path,
};
use crate::{Error, IWorkPackage, Result};

const COMPONENTS_FIELD: u32 = 3;
const VERSIONED_COMPONENTS_FIELD: u32 = 11;
const DATA_REFERENCES_FIELD: u32 = 7;
const DATA_IDENTIFIER_FIELD: u32 = 1;
const OBJECT_REFERENCES_FIELD: u32 = 2;
const OBJECT_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_COUNT_FIELD: u32 = 2;

#[derive(Clone, Copy)]
enum Adjustment {
    Add,
    Remove,
}

pub(crate) fn add_component_data_reference(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
) -> Result<()> {
    adjust_component_data_reference(
        package,
        component_identifier,
        data_identifier,
        object_identifier,
        Adjustment::Add,
    )
}

pub(crate) fn remove_component_data_reference(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
) -> Result<()> {
    adjust_component_data_reference(
        package,
        component_identifier,
        data_identifier,
        object_identifier,
        Adjustment::Remove,
    )
}

/// Copy every component data-reference owner covered by an object remap.
///
/// Chart and drawable graph duplication copies native style payloads verbatim.
/// PackageMetadata keeps the corresponding embedded-data ownership separately,
/// so those counts must be cloned along with the object graph.
pub(crate) fn clone_component_data_references(
    package: &mut IWorkPackage,
    component_identifier: u64,
    object_remap: &HashMap<u64, u64>,
) -> Result<()> {
    let references = component_object_data_references(package, component_identifier)?
        .into_iter()
        .filter_map(|(data_identifier, object_identifier, count)| {
            object_remap
                .get(&object_identifier)
                .copied()
                .map(|replacement| (data_identifier, replacement, count))
        })
        .collect::<Vec<_>>();
    for (data_identifier, object_identifier, count) in references {
        adjust_component_data_reference_by(
            package,
            component_identifier,
            data_identifier,
            object_identifier,
            Adjustment::Add,
            count,
        )?;
    }
    Ok(())
}

/// Remove all component data-reference owners belonging to an object set.
///
/// Returns the affected data identifiers so callers can reclaim assets that
/// became unreferenced after deleting the object graph.
pub(crate) fn remove_component_data_references_for_objects(
    package: &mut IWorkPackage,
    component_identifier: u64,
    object_identifiers: &[u64],
) -> Result<Vec<u64>> {
    let object_identifiers = object_identifiers.iter().copied().collect::<HashSet<_>>();
    let references = component_object_data_references(package, component_identifier)?
        .into_iter()
        .filter(|(_, object_identifier, _)| object_identifiers.contains(object_identifier))
        .collect::<Vec<_>>();
    let mut affected = HashSet::with_capacity(references.len());
    for (data_identifier, object_identifier, count) in references {
        adjust_component_data_reference_by(
            package,
            component_identifier,
            data_identifier,
            object_identifier,
            Adjustment::Remove,
            count,
        )?;
        affected.insert(data_identifier);
    }
    let mut affected = affected.into_iter().collect::<Vec<_>>();
    affected.sort_unstable();
    Ok(affected)
}

/// Return every embedded-data identifier registered to one package component.
///
/// Components may retain a data reference without any object owners, so this
/// deliberately reads the component-level registry rather than flattening only
/// object-reference records.
pub(crate) fn component_data_identifiers(
    package: &IWorkPackage,
    component_identifier: u64,
) -> Result<Vec<u64>> {
    let component = component_info(package, component_identifier)?;
    let mut identifiers = HashSet::with_capacity(component.data_references.len());
    for reference in component.data_references {
        if reference.data_identifier == 0 || !identifiers.insert(reference.data_identifier) {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} has an invalid or repeated data reference {}",
                reference.data_identifier
            )));
        }
    }
    let mut identifiers = identifiers.into_iter().collect::<Vec<_>>();
    identifiers.sort_unstable();
    Ok(identifiers)
}

fn adjust_component_data_reference(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    adjustment: Adjustment,
) -> Result<()> {
    adjust_component_data_reference_by(
        package,
        component_identifier,
        data_identifier,
        object_identifier,
        adjustment,
        1,
    )
}

fn adjust_component_data_reference_by(
    package: &mut IWorkPackage,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    adjustment: Adjustment,
    count: u32,
) -> Result<()> {
    if data_identifier == 0 || object_identifier == 0 {
        return Err(Error::InvalidFormat(
            "Component data and object identifiers must be non-zero".to_owned(),
        ));
    }
    if count == 0 {
        return Err(Error::InvalidFormat(
            "Component data-reference adjustment count must be non-zero".to_owned(),
        ));
    }
    package.update_archive(PACKAGE_METADATA_ENTRY, |archive| {
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
        let (object_index, message_index) = location.ok_or_else(|| {
            Error::InvalidFormat("PackageMetadata payload is missing".to_owned())
        })?;
        let object = &mut archive.objects[object_index];
        let original = &object.messages[message_index];
        let metadata = PackageMetadata::decode(original.data.as_slice())?;
        let old_count = component_reference_count(
            &metadata,
            component_identifier,
            data_identifier,
            object_identifier,
        )?;
        let expected_count = match adjustment {
            Adjustment::Add => old_count.checked_add(count).ok_or_else(|| {
                Error::InvalidFormat("Component data-reference count overflow".to_owned())
            })?,
            Adjustment::Remove => old_count.checked_sub(count).ok_or_else(|| {
                Error::InvalidFormat(
                    "Component data-reference removal has no matching reference".to_owned(),
                )
            })?,
        };

        let mut matched_components = 0usize;
        let mut data = original.data.clone();
        for field in [COMPONENTS_FIELD, VERSIONED_COMPONENTS_FIELD] {
            data = transform_length_delimited_fields_at_path(&data, &[field], |component_data| {
                let component = ComponentInfo::decode(component_data)?;
                if component.identifier != component_identifier {
                    return Ok(component_data.to_vec());
                }
                matched_components += 1;
                adjust_component_payload(
                    component_data,
                    &component,
                    data_identifier,
                    object_identifier,
                    adjustment,
                    count,
                )
            })?;
        }
        if matched_components != 1 {
            return Err(Error::InvalidFormat(format!(
                "PackageMetadata contains {matched_components} components with identifier {component_identifier}"
            )));
        }
        let verified = PackageMetadata::decode(data.as_slice())?;
        if component_reference_count(
            &verified,
            component_identifier,
            data_identifier,
            object_identifier,
        )? != expected_count
        {
            return Err(Error::InvalidFormat(
                "Component data-reference adjustment failed validation".to_owned(),
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

fn adjust_component_payload(
    data: &[u8],
    component: &ComponentInfo,
    data_identifier: u64,
    object_identifier: u64,
    adjustment: Adjustment,
    count: u32,
) -> Result<Vec<u8>> {
    let matches = component
        .data_references
        .iter()
        .filter(|reference| reference.data_identifier == data_identifier)
        .collect::<Vec<_>>();
    match (adjustment, matches.as_slice()) {
        (Adjustment::Add, []) => append_repeated_length_delimited_field(
            data,
            DATA_REFERENCES_FIELD,
            &ComponentDataReference {
                data_identifier,
                object_reference_list: vec![ObjectReference {
                    object_identifier,
                    count,
                }],
            }
            .encode_to_vec(),
        ),
        (_, [reference]) => {
            let owners = reference
                .object_reference_list
                .iter()
                .filter(|owner| owner.object_identifier == object_identifier)
                .collect::<Vec<_>>();
            let removes_entire_reference = matches!(adjustment, Adjustment::Remove)
                && owners.len() == 1
                && owners[0].count == count
                && reference.object_reference_list.len() == 1;
            if removes_entire_reference {
                return remove_repeated_length_delimited_field_where(
                    data,
                    DATA_REFERENCES_FIELD,
                    |payload| validate_data_reference_identifier(payload, data_identifier),
                );
            }
            let mut matched = 0usize;
            let patched = transform_length_delimited_fields_at_path(
                data,
                &[DATA_REFERENCES_FIELD],
                |payload| {
                    if !validate_data_reference_identifier(payload, data_identifier)? {
                        return Ok(payload.to_vec());
                    }
                    matched += 1;
                    adjust_data_reference_payload(
                        payload,
                        reference,
                        object_identifier,
                        adjustment,
                        count,
                    )
                },
            )?;
            if matched != 1 {
                return Err(Error::InvalidFormat(
                    "Component data-reference wire does not match its decoded value".to_owned(),
                ));
            }
            Ok(patched)
        },
        (Adjustment::Remove, []) => Err(Error::InvalidFormat(format!(
            "Component has no data reference {data_identifier}"
        ))),
        (_, _) => Err(Error::InvalidFormat(format!(
            "Component repeats data reference {data_identifier}"
        ))),
    }
}

fn adjust_data_reference_payload(
    data: &[u8],
    reference: &ComponentDataReference,
    object_identifier: u64,
    adjustment: Adjustment,
    count: u32,
) -> Result<Vec<u8>> {
    let owners = reference
        .object_reference_list
        .iter()
        .filter(|owner| owner.object_identifier == object_identifier)
        .collect::<Vec<_>>();
    match (adjustment, owners.as_slice()) {
        (Adjustment::Add, []) => append_repeated_length_delimited_field(
            data,
            OBJECT_REFERENCES_FIELD,
            &ObjectReference {
                object_identifier,
                count,
            }
            .encode_to_vec(),
        ),
        (_, [owner]) => {
            if matches!(adjustment, Adjustment::Remove) && owner.count == count {
                return remove_repeated_length_delimited_field_where(
                    data,
                    OBJECT_REFERENCES_FIELD,
                    |payload| validate_object_reference_identifier(payload, object_identifier),
                );
            }
            let mut matched = 0usize;
            let patched = transform_length_delimited_fields_at_path(
                data,
                &[OBJECT_REFERENCES_FIELD],
                |payload| {
                    if !validate_object_reference_identifier(payload, object_identifier)? {
                        return Ok(payload.to_vec());
                    }
                    matched += 1;
                    let count = match adjustment {
                        Adjustment::Add => owner.count.checked_add(count).ok_or_else(|| {
                            Error::InvalidFormat(
                                "Component object-reference count overflow".to_owned(),
                            )
                        })?,
                        Adjustment::Remove => owner.count.checked_sub(count).ok_or_else(|| {
                            Error::InvalidFormat(
                                "Component object-reference count underflow".to_owned(),
                            )
                        })?,
                    };
                    patch_varint_field(payload, REFERENCE_COUNT_FIELD, true, Some(u64::from(count)))
                },
            )?;
            if matched != 1 {
                return Err(Error::InvalidFormat(
                    "Component object-reference wire does not match its decoded value".to_owned(),
                ));
            }
            Ok(patched)
        },
        (Adjustment::Remove, []) => Err(Error::InvalidFormat(format!(
            "Component data reference has no object {object_identifier}"
        ))),
        (_, _) => Err(Error::InvalidFormat(format!(
            "Component data reference repeats object {object_identifier}"
        ))),
    }
}

fn component_object_data_references(
    package: &IWorkPackage,
    component_identifier: u64,
) -> Result<Vec<(u64, u64, u32)>> {
    let component = component_info(package, component_identifier)?;
    let capacity = component
        .data_references
        .iter()
        .map(|reference| reference.object_reference_list.len())
        .sum();
    let mut references = Vec::with_capacity(capacity);
    let mut data_identifiers = HashSet::with_capacity(component.data_references.len());
    for reference in &component.data_references {
        if reference.data_identifier == 0 || !data_identifiers.insert(reference.data_identifier) {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} has an invalid or repeated data reference {}",
                reference.data_identifier
            )));
        }
        let mut object_identifiers = HashSet::with_capacity(reference.object_reference_list.len());
        for owner in &reference.object_reference_list {
            if owner.object_identifier == 0
                || owner.count == 0
                || !object_identifiers.insert(owner.object_identifier)
            {
                return Err(Error::InvalidFormat(format!(
                    "Component data reference {} has an invalid or repeated object {}",
                    reference.data_identifier, owner.object_identifier
                )));
            }
            references.push((
                reference.data_identifier,
                owner.object_identifier,
                owner.count,
            ));
        }
    }
    Ok(references)
}

fn component_info(package: &IWorkPackage, component_identifier: u64) -> Result<ComponentInfo> {
    let archive = package.archive(PACKAGE_METADATA_ENTRY)?;
    let messages = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Package must contain exactly one PackageMetadata payload, found {}",
            messages.len()
        )));
    };
    let metadata = PackageMetadata::decode(message.data.as_slice())?;
    let mut components = metadata
        .components
        .into_iter()
        .chain(metadata.versioned_components)
        .filter(|component| component.identifier == component_identifier);
    let Some(component) = components.next() else {
        return Err(Error::InvalidFormat(format!(
            "PackageMetadata must contain exactly one component {component_identifier}"
        )));
    };
    if components.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "PackageMetadata must contain exactly one component {component_identifier}"
        )));
    }
    Ok(component)
}

fn validate_data_reference_identifier(data: &[u8], expected: u64) -> Result<bool> {
    let decoded = ComponentDataReference::decode(data)?;
    let _ = patch_varint_field(
        data,
        DATA_IDENTIFIER_FIELD,
        true,
        Some(decoded.data_identifier),
    )?;
    Ok(decoded.data_identifier == expected)
}

fn validate_object_reference_identifier(data: &[u8], expected: u64) -> Result<bool> {
    let decoded = ObjectReference::decode(data)?;
    let _ = patch_varint_field(
        data,
        OBJECT_IDENTIFIER_FIELD,
        true,
        Some(decoded.object_identifier),
    )?;
    let _ = patch_varint_field(
        data,
        REFERENCE_COUNT_FIELD,
        true,
        Some(u64::from(decoded.count)),
    )?;
    Ok(decoded.object_identifier == expected)
}

fn component_reference_count(
    metadata: &PackageMetadata,
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
) -> Result<u32> {
    let components = metadata
        .components
        .iter()
        .chain(&metadata.versioned_components)
        .filter(|component| component.identifier == component_identifier)
        .collect::<Vec<_>>();
    let [component] = components.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "PackageMetadata must contain exactly one component {component_identifier}"
        )));
    };
    let references = component
        .data_references
        .iter()
        .filter(|reference| reference.data_identifier == data_identifier)
        .collect::<Vec<_>>();
    let reference = match references.as_slice() {
        [] => return Ok(0),
        [reference] => *reference,
        _ => {
            return Err(Error::InvalidFormat(format!(
                "Component {component_identifier} repeats data reference {data_identifier}"
            )));
        },
    };
    let owners = reference
        .object_reference_list
        .iter()
        .filter(|owner| owner.object_identifier == object_identifier)
        .collect::<Vec<_>>();
    match owners.as_slice() {
        [] => Ok(0),
        [owner] => Ok(owner.count),
        _ => Err(Error::InvalidFormat(format!(
            "Component data reference {data_identifier} repeats object {object_identifier}"
        ))),
    }
}
