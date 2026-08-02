//! Arbitrary logical-axis removal from a Numbers stable UID map.

use super::*;

pub(super) fn delete_row_uid(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    old_rows: usize,
    deletion: usize,
) -> Result<()> {
    update_uid_map(package, locations, identifier, |map| {
        delete_uid_axis(
            &mut map.sorted_row_uids,
            &mut map.row_index_for_uid,
            &mut map.row_uid_for_index,
            old_rows,
            deletion,
            "row",
        )
    })
}

pub(super) fn delete_column_uid(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    old_columns: usize,
    deletion: usize,
) -> Result<()> {
    update_uid_map(package, locations, identifier, |map| {
        delete_uid_axis(
            &mut map.sorted_column_uids,
            &mut map.column_index_for_uid,
            &mut map.column_uid_for_index,
            old_columns,
            deletion,
            "column",
        )
    })
}

fn update_uid_map(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    update: impl FnOnce(&mut tst::ColumnRowUidMapArchive) -> Result<()>,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                tst::ColumnRowUidMapArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Object {identifier} has no UID map payload"))
            })?;
        let original = object.messages[message_index].data.clone();
        let previous = tst::ColumnRowUidMapArchive::decode(original.as_slice())?;
        let mut current = previous.clone();
        update(&mut current)?;
        let data = rewrite_uid_map_wire(&original, &previous, &current)?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

fn delete_uid_axis(
    sorted: &mut Vec<tsp::Uuid>,
    index_for_uid: &mut Vec<u32>,
    uid_for_index: &mut Vec<u32>,
    old_length: usize,
    deletion: usize,
    axis: &str,
) -> Result<()> {
    if sorted.len() != old_length
        || index_for_uid.len() != old_length
        || uid_for_index.len() != old_length
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers {axis} UID map lengths do not match table dimensions"
        )));
    }
    let sorted_index = *uid_for_index
        .get(deletion)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers {axis} UID map is truncated")))?
        as usize;
    let deletion_u32 = u32::try_from(deletion)
        .map_err(|_| Error::ParseError(format!("Numbers {axis} exceeds u32")))?;
    if sorted_index >= sorted.len()
        || index_for_uid.get(sorted_index).copied() != Some(deletion_u32)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers {axis} UID map is inconsistent"
        )));
    }

    uid_for_index.remove(deletion);
    sorted.remove(sorted_index);
    index_for_uid.remove(sorted_index);
    let sorted_index_u32 = u32::try_from(sorted_index)
        .map_err(|_| Error::ParseError(format!("Numbers {axis} UID index exceeds u32")))?;
    for value in uid_for_index {
        if *value > sorted_index_u32 {
            *value -= 1;
        }
    }
    for index in index_for_uid {
        if *index > deletion_u32 {
            *index -= 1;
        }
    }
    Ok(())
}
