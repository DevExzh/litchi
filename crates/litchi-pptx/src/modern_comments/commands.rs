//! Package facade for typed `cmChg` descriptors in the existing Changes
//! Information part. The surrounding document-change graph remains inert.

use super::semantic::CommentChanges;
use super::wire;
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
    pub value: CommentChanges,
}

/// Load every typed modern-comment command while leaving all other change XML
/// untouched and inert.
pub fn load_modern_comment_changes(package: &OpcPackage) -> Result<Vec<ChangeCommand>> {
    let Some(part) = super::super::presentation_properties::metadata::changes::load(package)?
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
pub fn update_modern_comment_changes<F>(package: &mut OpcPackage, update: F) -> Result<bool>
where
    F: FnOnce(&mut Vec<ChangeCommand>),
{
    let Some(mut part) = super::super::presentation_properties::metadata::changes::load(package)?
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
    let xml = part.changes_information.to_xml()?;
    let uri = PackURI::new(&part.part_name).map_err(invalid)?;
    let stored = package.get_part(&uri)?;
    if stored.content_type() != CHANGES_CONTENT_TYPE {
        return Err(invalid(
            "Changes Information part has an unexpected content type",
        ));
    }
    package.get_part_mut(&uri)?.set_blob(xml);
    package.unsign();
    Ok(true)
}
