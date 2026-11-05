//! Package facade for typed `cmChg` descriptors in the existing Changes
//! Information part. The surrounding document-change graph remains inert.

use super::semantic::changes::Changes;
use super::wire;
use crate::package::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI};

const CHANGES_CONTENT_TYPE: &str = "application/vnd.ms-powerpoint.changesinfo+xml";

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// A typed command with its stable location inside the Changes Information
/// part. Location fields are checked before a mutation is committed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangeCommand {
    pub part_name: String,
    pub list_index: usize,
    pub descriptor_index: usize,
    pub value: Changes,
}

/// Load every typed modern-comment command while leaving all other change XML
/// untouched and inert.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_modern_comment_changes(
    package: &OpcPackage,
    selected_limits: super::super::presentation_properties::metadata::changes::Limits,
) -> Result<Vec<ChangeCommand>> {
    let limits = selected_limits.constrained_by_source(package.read_limits());
    let Some(part) =
        super::super::presentation_properties::metadata::changes::load(package, limits)?
    else {
        return Ok(Vec::new());
    };
    let mut output = Vec::new();
    for (list_index, list) in part.changes_information.change_lists.iter().enumerate() {
        for (descriptor_index, descriptor) in list.changes.iter().enumerate() {
            for value in wire::collect_change_commands(&descriptor.xml)? {
                output.push(ChangeCommand {
                    part_name: part.part_name.clone(),
                    list_index,
                    descriptor_index,
                    value,
                });
            }
        }
    }
    Ok(output)
}

/// Update typed modern-comment commands transactionally.
///
/// The callback may edit typed values but cannot add/remove/reorder command
/// fragments. Such structural edits would risk changing unrelated future
/// command data and are rejected. No command is interpreted or executed.
///
/// A change to a signed package is refused with
/// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
/// (ADR 0006) before anything is mutated: remove the signature explicitly
/// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
/// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
/// no-op keeps the signature.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub(crate) fn update_modern_comment_changes<F>(
    package: &mut OpcPackage,
    selected_limits: super::super::presentation_properties::metadata::changes::Limits,
    update: F,
) -> Result<bool>
where
    F: FnOnce(&mut Vec<ChangeCommand>),
{
    let limits = selected_limits.constrained_by_source(package.read_limits());
    let Some(mut part) =
        super::super::presentation_properties::metadata::changes::load(package, limits)?
    else {
        return Ok(false);
    };
    if part
        .changes_information
        .change_lists
        .iter()
        .flat_map(|list| &list.changes)
        .any(|descriptor| descriptor.xml.len() > super::MAX_BYTES)
    {
        return Err(invalid(
            "Changes Information descriptor exceeds implementation limit",
        ));
    }

    let mut commands = Vec::new();
    for (list_index, list) in part.changes_information.change_lists.iter().enumerate() {
        for (descriptor_index, descriptor) in list.changes.iter().enumerate() {
            for value in wire::collect_change_commands(&descriptor.xml)? {
                commands.push(ChangeCommand {
                    part_name: part.part_name.clone(),
                    list_index,
                    descriptor_index,
                    value,
                });
            }
        }
    }
    if commands.is_empty() {
        return Ok(false);
    }
    let original = commands.clone();
    update(&mut commands);
    if commands.len() != original.len()
        || commands.iter().zip(&original).any(|(after, before)| {
            after.part_name != before.part_name
                || after.list_index != before.list_index
                || after.descriptor_index != before.descriptor_index
        })
    {
        return Err(invalid(
            "modern comment command mutation cannot change command locations",
        ));
    }

    let mut cursor = 0usize;
    for (list_index, list) in part.changes_information.change_lists.iter_mut().enumerate() {
        for (descriptor_index, descriptor) in list.changes.iter_mut().enumerate() {
            let count = wire::collect_change_commands(&descriptor.xml)?.len();
            if count == 0 {
                continue;
            }
            let end = cursor
                .checked_add(count)
                .ok_or_else(|| invalid("modern comment command count overflow"))?;
            if end > commands.len() {
                return Err(invalid("modern comment command locations are inconsistent"));
            }
            descriptor.xml = wire::replace_change_commands(
                &descriptor.xml,
                &commands[cursor..end]
                    .iter()
                    .map(|command| command.value.clone())
                    .collect::<Vec<_>>(),
            )?;
            cursor = end;
            let _ = (list_index, descriptor_index);
        }
    }
    if cursor != commands.len() {
        return Err(invalid("modern comment command locations are incomplete"));
    }
    let xml = part.changes_information.to_xml(limits)?;
    let uri = PackURI::new(&part.part_name).map_err(invalid)?;
    let stored = package.get_part(&uri)?;
    if stored.content_type() != CHANGES_CONTENT_TYPE {
        return Err(invalid(
            "Changes Information part has an unexpected content type",
        ));
    }
    // Commands that serialize to the stored bytes publish nothing, so an
    // exact no-op keeps a signed package byte for byte.
    if stored.blob() == xml.as_slice() {
        return Ok(true);
    }
    refuse_signed_change(package)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    settle_unsigned_change(package)?;
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, OpcPackage, PackURI};

    const PRESENTATION_CONTENT_TYPE: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    fn package_with_changes(xml: Vec<u8>) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            presentation_name.clone(),
            PRESENTATION_CONTENT_TYPE.into(),
            Vec::new(),
        )));
        package
            .get_part_mut(&presentation_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                super::super::super::presentation_properties::metadata::changes::RELATIONSHIP_TYPE
                    .into(),
                "changesInfos/changesInfo1.xml".into(),
                "rIdChanges".into(),
                false,
            );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/changesInfos/changesInfo1.xml").unwrap(),
            CHANGES_CONTENT_TYPE.into(),
            xml,
        )));
        package
    }

    #[test]
    fn selected_changes_profile_is_used_by_the_raw_helper() {
        let xml = br#"<pc:chgInfo xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command"><pc:docChgLst/></pc:chgInfo>"#;
        let package = package_with_changes(xml.to_vec());
        assert!(
            load_modern_comment_changes(
                &package,
                super::super::super::presentation_properties::metadata::changes::Limits::DEFAULT,
            )
            .unwrap()
            .is_empty()
        );

        let default =
            super::super::super::presentation_properties::metadata::changes::Limits::DEFAULT;
        let tight = super::super::super::presentation_properties::metadata::changes::Limits::new(
            xml.len() - 1,
            default.max_output_bytes(),
            default.max_depth(),
            default.max_nodes(),
            default.max_lists(),
            default.max_changes(),
            default.max_extensions(),
            default.max_string_bytes(),
        )
        .unwrap();
        assert!(load_modern_comment_changes(&package, tight).is_err());
    }
}
