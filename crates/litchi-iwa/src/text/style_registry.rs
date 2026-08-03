//! Shared registration for private text-style variations.

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    remove_component_external_reference, remove_component_external_references_to_object,
    remove_component_object_uuids,
};
use crate::protobuf::tsp;
use crate::wire::append_repeated_length_delimited_field;
use crate::{Error, IWorkPackage, Result};
use prost::Message;

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLESHEET_STYLES_FIELD: u32 = 1;

pub(super) fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    object_archive(package, identifier).map(|(name, _)| name)
}

/// Locate an object while retaining the parsed archive that contains it.
///
/// Every candidate archive is still parsed before a successful result is
/// returned. This keeps malformed later archives observable and preserves the
/// duplicate-identifier check while allowing callers that need the object to
/// avoid parsing the matching archive a second time.
pub(super) fn object_archive(package: &IWorkPackage, identifier: u64) -> Result<(String, Archive)> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        if archive.object(identifier).is_none() {
            continue;
        }
        if found.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork object {identifier} occurs in multiple archives"
            )));
        }
        found = Some((name.to_owned(), archive));
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

pub(crate) fn insert_private_style(
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

#[cfg(test)]
mod tests {
    use super::*;

    fn package_with_object(name: &str, identifier: u64) -> IWorkPackage {
        let mut package = IWorkPackage::new();
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    identifier,
                    vec![RawMessage {
                        type_: 99,
                        data: vec![0],
                    }],
                )
                .unwrap(),
            ],
        };
        package.replace_archive(name, &archive).unwrap();
        package
    }

    #[test]
    fn object_archive_returns_the_parsed_matching_archive() {
        let package = package_with_object("Index/Styles.iwa", 7);

        let (name, archive) = object_archive(&package, 7).unwrap();

        assert_eq!(name, "Index/Styles.iwa");
        assert!(archive.object(7).is_some());
        assert_eq!(object_archive_name(&package, 7).unwrap(), name);
    }

    #[test]
    fn object_archive_rejects_missing_and_duplicate_identifiers() {
        assert!(object_archive(&IWorkPackage::new(), 7).is_err());

        let mut package = package_with_object("Index/One.iwa", 7);
        let second = package_with_object("Index/Two.iwa", 7)
            .entry("Index/Two.iwa")
            .unwrap()
            .to_vec();
        package.insert_entry("Index/Two.iwa", second).unwrap();

        assert!(object_archive(&package, 7).is_err());
    }

    #[test]
    fn object_archive_still_parses_later_archives_after_a_match() {
        let mut package = package_with_object("Index/One.iwa", 7);
        package.insert_entry("Index/Z-Broken.iwa", vec![0]).unwrap();

        assert!(object_archive(&package, 7).is_err());
    }
}
