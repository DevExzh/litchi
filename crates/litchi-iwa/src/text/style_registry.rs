//! Shared registration for private text-style variations.

use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    remove_component_external_reference, remove_component_external_references_to_object,
    remove_component_object_uuids,
};
use crate::protobuf::tsp;
use crate::wire::append_repeated_length_delimited_field;
use crate::{Error, IWorkPackage, Result};
use crate::{archive::ArchiveObject, archive::RawMessage};
use prost::Message;

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLESHEET_STYLES_FIELD: u32 = 1;

pub(super) fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_some()
            && found.replace(name.to_owned()).is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork object {identifier} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork object {identifier} is missing")))
}

pub(crate) fn unregister_owner_reference_if_unused(
    package: &mut IWorkPackage,
    owner_archive_name: &str,
    style_archive_name: &str,
    style_id: u64,
) -> Result<()> {
    let Some(style_component) = component_identifier_for_entry(package, style_archive_name)? else {
        return Ok(());
    };
    let Some(owner_component) = component_identifier_for_entry(package, owner_archive_name)? else {
        return Ok(());
    };
    if owner_component == style_component {
        return Ok(());
    }
    for archive_name in package.iwa_entry_names() {
        if component_identifier_for_entry(package, archive_name)? == Some(owner_component)
            && package.archive(archive_name)?.objects.iter().any(|object| {
                object
                    .archive_info
                    .message_infos
                    .iter()
                    .any(|info| info.object_references.contains(&style_id))
            })
        {
            return Ok(());
        }
    }
    remove_component_external_reference(package, owner_component, style_component, style_id)
}

pub(crate) fn register_private_style(
    package: &mut IWorkPackage,
    owner_archive_name: &str,
    style_archive_name: &str,
    style_id: u64,
) -> Result<()> {
    let Some(style_component) = component_identifier_for_entry(package, style_archive_name)? else {
        return Ok(());
    };
    add_component_object_uuids(package, style_component, &[style_id])?;
    if let Some(owner_component) = component_identifier_for_entry(package, owner_archive_name)?
        && owner_component != style_component
    {
        add_component_external_reference(package, owner_component, style_component, style_id)?;
    }
    Ok(())
}

pub(crate) fn register_style_reference(
    package: &mut IWorkPackage,
    owner_archive_name: &str,
    style_archive_name: &str,
    style_id: u64,
) -> Result<()> {
    let Some(style_component) = component_identifier_for_entry(package, style_archive_name)? else {
        return Ok(());
    };
    if let Some(owner_component) = component_identifier_for_entry(package, owner_archive_name)?
        && owner_component != style_component
    {
        add_component_external_reference(package, owner_component, style_component, style_id)?;
    }
    Ok(())
}

pub(super) fn insert_private_style(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    style_id: u64,
    style: ArchiveObject,
) -> Result<()> {
    if style.archive_info.identifier != Some(style_id) {
        return Err(Error::InvalidFormat(format!(
            "private iWork style object does not use identifier {style_id}"
        )));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork stylesheet {stylesheet_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == STYLESHEET_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet {stylesheet_id} must have exactly one Stylesheet payload"
            )));
        };
        let original = &object.messages[*index];
        let reference = tsp::Reference {
            identifier: style_id,
            ..Default::default()
        };
        let data = append_repeated_length_delimited_field(
            &original.data,
            STYLESHEET_STYLES_FIELD,
            &reference.encode_to_vec(),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        if info.object_references.contains(&style_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet already references style {style_id}"
            )));
        }
        info.object_references.push(style_id);
        archive.insert_object(style)?;
        Ok(())
    })
}

pub(crate) fn unregister_private_style(
    package: &mut IWorkPackage,
    owner_archive_name: &str,
    style_archive_name: &str,
    style_id: u64,
    replacement_external_reference: Option<u64>,
) -> Result<()> {
    let Some(style_component) = component_identifier_for_entry(package, style_archive_name)? else {
        return Ok(());
    };
    remove_component_object_uuids(package, style_component, &[style_id])?;
    remove_component_external_references_to_object(package, style_component, style_id)?;
    if let Some(replacement) = replacement_external_reference
        && let Some(owner_component) = component_identifier_for_entry(package, owner_archive_name)?
        && owner_component != style_component
    {
        add_component_external_reference(package, owner_component, style_component, replacement)?;
    }
    Ok(())
}
