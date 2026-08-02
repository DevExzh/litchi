//! Table-model and stroke-sidecar dimension updates after axis deletion.

use super::*;

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
