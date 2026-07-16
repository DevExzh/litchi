//! Shared registration for private text-style variations.

use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    remove_component_external_reference, remove_component_external_references_to_object,
    remove_component_object_uuids,
};
use crate::{Error, IWorkPackage, Result};

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

pub(super) fn unregister_owner_reference_if_unused(
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

pub(super) fn register_private_style(
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

pub(super) fn unregister_private_style(
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
