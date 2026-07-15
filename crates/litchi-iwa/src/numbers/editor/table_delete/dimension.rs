//! Table-model and stroke-sidecar dimension updates after axis deletion.

use super::*;

pub(super) fn set_stroke_dimensions(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    rows: u32,
    columns: u32,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke sidecar {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers stroke sidecar {identifier} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| tst::StrokeSidecarArchive::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Object {identifier} has no stroke sidecar payload"))
            })?;
        let original = object.messages[message_index].data.clone();
        let previous = tst::StrokeSidecarArchive::decode(original.as_slice())?;
        if !previous.left_column_stroke_layers.is_empty()
            || !previous.right_column_stroke_layers.is_empty()
            || !previous.top_row_stroke_layers.is_empty()
            || !previous.bottom_row_stroke_layers.is_empty()
        {
            return Err(Error::ParseError(
                "Cannot yet delete a Numbers table axis with explicit stroke layers".to_owned(),
            ));
        }
        let mut current = previous.clone();
        current.row_count = Some(rows);
        current.column_count = Some(columns);
        let data = rewrite_stroke_sidecar_wire(&original, &previous, &current)?;
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

pub(super) fn set_table_dimensions(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    rows: u32,
    columns: u32,
) -> Result<()> {
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers table object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table object {table_id} is missing"))
        })?;
        let message_index = find_table_model_message(object)?;
        let original = object.messages[message_index].data.clone();
        let mut data = patch_varint_field(&original, 6, true, Some(u64::from(rows)))?;
        data = patch_varint_field(&data, 7, true, Some(u64::from(columns)))?;
        let verified = TableModelArchive::decode(data.as_slice())?;
        if (verified.number_of_rows, verified.number_of_columns) != (rows, columns) {
            return Err(Error::InvalidFormat(
                "Numbers table dimensions failed wire validation".to_owned(),
            ));
        }
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
