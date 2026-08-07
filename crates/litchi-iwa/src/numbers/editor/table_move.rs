//! Transactional transfer of an existing table between workbook sheets.

use super::*;
use crate::wire::{
    append_repeated_length_delimited_field, parse_wire_fields,
    remove_repeated_length_delimited_field_where,
};
use litchi_numbers::{SheetSelector, TableSelector};

impl NumbersEditor {
    /// Move an existing table to another sheet, preserving its object identity and contents.
    ///
    /// The table is appended to the destination sheet's drawable order. Its native table-info
    /// parent reference, sheet ownership lists, and IWA reference metadata are updated together.
    pub fn move_table(
        &mut self,
        selector: TableSelector,
        target: SheetSelector,
    ) -> Result<NumbersTableInfo> {
        let table_id = super::selectors::table_id(self, selector)?;
        let table = self
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| Error::ParseError(format!("Numbers table {table_id} not found")))?;
        let target_sheet_id = super::selectors::sheet_id(self, target)?;

        let owner = find_table_owner(&self.package, table_id)?;
        if owner.sheet_id == target_sheet_id {
            return Ok(table);
        }
        let locations = object_locations(&self.package)?;
        let source_sheet_archive = locations.get(&owner.sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {} is missing", owner.sheet_id))
        })?;
        let target_sheet_archive = locations.get(&target_sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet {target_sheet_id} is missing"))
        })?;
        let table_info_archive = locations.get(&owner.table_info_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table info {} is missing",
                owner.table_info_id
            ))
        })?;

        let reference_payload = sheet_drawable_reference_payload(
            &self.package,
            source_sheet_archive,
            owner.sheet_id,
            owner.table_info_id,
        )?;
        let mut staged = self.package.clone();
        remove_sheet_drawable(
            &mut staged,
            source_sheet_archive,
            owner.sheet_id,
            owner.table_info_id,
        )?;
        append_sheet_drawable(
            &mut staged,
            target_sheet_archive,
            target_sheet_id,
            owner.table_info_id,
            &reference_payload,
        )?;
        patch_table_parent(
            &mut staged,
            table_info_archive,
            owner.table_info_id,
            owner.sheet_id,
            target_sheet_id,
        )?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let verified_owner = find_table_owner(&staged, table_id)?;
        let verified_table = verified
            .tables()?
            .into_iter()
            .find(|candidate| candidate.object_id == table_id)
            .ok_or_else(|| Error::InvalidFormat("Moved Numbers table disappeared".to_owned()))?;
        if verified_owner.sheet_id != target_sheet_id || verified_table != table {
            return Err(Error::InvalidFormat(
                "Numbers table move failed validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(verified_table)
    }
}

fn sheet_drawable_reference_payload(
    package: &IWorkPackage,
    archive_name: &str,
    sheet_id: u64,
    drawable_id: u64,
) -> Result<Vec<u8>> {
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(sheet_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
    let (message_index, _) = decode_sheet(object)?;
    let message = &object.messages[message_index];
    let sheet_data = sheet_wire_payload(&message.data, message.type_)?;
    let matches = repeated_length_delimited_payloads(sheet_data, 2)?
        .into_iter()
        .filter_map(|payload| {
            crate::protobuf::tsp::Reference::decode(payload)
                .ok()
                .filter(|reference| reference.identifier == drawable_id)
                .map(|_| payload.to_vec())
        })
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [payload] => Ok(payload.clone()),
        _ => Err(Error::InvalidFormat(format!(
            "Numbers sheet {sheet_id} references drawable {drawable_id} {} times",
            matches.len()
        ))),
    }
}

fn sheet_wire_payload(data: &[u8], message_type: u32) -> Result<&[u8]> {
    if message_type != 3 {
        return Ok(data);
    }
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number() == 1 && field.wire_type() == 2)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [field] => Ok(&data[field.payload_start()..field.end()]),
        _ => Err(Error::InvalidFormat(format!(
            "Numbers form sheet must contain exactly one nested sheet payload, found {}",
            matches.len()
        ))),
    }
}

fn transform_sheet_wire<F>(data: &[u8], message_type: u32, update: F) -> Result<Vec<u8>>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    if message_type == 3 {
        transform_length_delimited_field(data, 1, update)
    } else {
        update(data)
    }
}

fn remove_sheet_drawable(
    package: &mut IWorkPackage,
    archive_name: &str,
    sheet_id: u64,
    drawable_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive
            .object_mut(sheet_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
        let (message_index, sheet) = decode_sheet(object)?;
        if sheet
            .drawable_infos
            .iter()
            .filter(|reference| reference.identifier == drawable_id)
            .count()
            != 1
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers sheet {sheet_id} does not own drawable {drawable_id} exactly once"
            )));
        }
        let message_type = object.messages[message_index].type_;
        let data = transform_sheet_wire(
            object.messages[message_index].data.as_slice(),
            message_type,
            |sheet_data| {
                remove_repeated_length_delimited_field_where(sheet_data, 2, |payload| {
                    Ok(crate::protobuf::tsp::Reference::decode(payload)?.identifier == drawable_id)
                })
            },
        )?;
        let verified = decode_sheet_data(&data, message_type)?;
        if verified
            .drawable_infos
            .iter()
            .any(|reference| reference.identifier == drawable_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers source-sheet drawable removal failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        remove_metadata_reference(object, message_index, drawable_id);
        Ok(())
    })
}

fn append_sheet_drawable(
    package: &mut IWorkPackage,
    archive_name: &str,
    sheet_id: u64,
    drawable_id: u64,
    payload: &[u8],
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive
            .object_mut(sheet_id)
            .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
        let (message_index, sheet) = decode_sheet(object)?;
        if sheet
            .drawable_infos
            .iter()
            .any(|reference| reference.identifier == drawable_id)
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers target sheet {sheet_id} already owns drawable {drawable_id}"
            )));
        }
        let previous = sheet
            .drawable_infos
            .iter()
            .map(|reference| reference.identifier)
            .collect::<HashSet<_>>();
        let message_type = object.messages[message_index].type_;
        let data = transform_sheet_wire(
            object.messages[message_index].data.as_slice(),
            message_type,
            |sheet_data| append_repeated_length_delimited_field(sheet_data, 2, payload),
        )?;
        let verified = decode_sheet_data(&data, message_type)?;
        if verified
            .drawable_infos
            .last()
            .map(|reference| reference.identifier)
            != Some(drawable_id)
        {
            return Err(Error::InvalidFormat(
                "Numbers target-sheet drawable insertion failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        add_metadata_reference(object, message_index, drawable_id, &previous);
        Ok(())
    })
}

fn decode_sheet_data(data: &[u8], message_type: u32) -> Result<tn::SheetArchive> {
    if message_type == 3 {
        Ok(tn::FormBasedSheetArchive::decode(data)?.super_)
    } else {
        Ok(tn::SheetArchive::decode(data)?)
    }
}

fn patch_table_parent(
    package: &mut IWorkPackage,
    archive_name: &str,
    table_info_id: u64,
    source_sheet_id: u64,
    target_sheet_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_info_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table info {table_info_id} is missing"))
        })?;
        let (message_index, info) = decode_table_info(object)?;
        let parent = info.super_.parent.map(|reference| reference.identifier);
        if parent.is_some_and(|identifier| identifier != source_sheet_id) {
            return Err(Error::InvalidFormat(format!(
                "Numbers table info {table_info_id} parent {parent:?} disagrees with owning sheet {source_sheet_id}"
            )));
        }
        if parent.is_none() {
            return Ok(());
        }
        let mut remap = HashMap::new();
        remap.insert(source_sheet_id, target_sheet_id);
        let message_type = object.messages[message_index].type_;
        let data = remap_numbers_reference_paths(
            object.messages[message_index].data.as_slice(),
            &[&[1, 2]],
            &remap,
        )?;
        let verified = tst::TableInfoArchive::decode(data.as_slice())?;
        if verified.super_.parent.map(|reference| reference.identifier) != Some(target_sheet_id)
            || verified.table_model != info.table_model
        {
            return Err(Error::InvalidFormat(
                "Numbers table parent patch failed validation".to_owned(),
            ));
        }
        object.replace_message(message_index, RawMessage { type_: message_type, data })?;
        replace_metadata_reference(
            object,
            message_index,
            source_sheet_id,
            target_sheet_id,
        )?;
        Ok(())
    })
}

fn remove_metadata_reference(object: &mut ArchiveObject, message_index: usize, identifier: u64) {
    let info = &mut object.archive_info.message_infos[message_index];
    info.object_references
        .retain(|candidate| *candidate != identifier);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|candidate| *candidate != identifier);
    }
}

fn add_metadata_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    identifier: u64,
    previous: &HashSet<u64>,
) {
    let info = &mut object.archive_info.message_infos[message_index];
    if !info.object_references.contains(&identifier) {
        info.object_references.push(identifier);
    }
    for field in &mut info.field_infos {
        if field
            .object_references
            .iter()
            .any(|candidate| previous.contains(candidate))
            && !field.object_references.contains(&identifier)
        {
            field.object_references.push(identifier);
        }
    }
}

fn replace_metadata_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    source: u64,
    target: u64,
) -> Result<()> {
    let info = &mut object.archive_info.message_infos[message_index];
    replace_reference_values(&mut info.object_references, source, target)?;
    for field in &mut info.field_infos {
        replace_reference_values(&mut field.object_references, source, target)?;
    }
    Ok(())
}

fn replace_reference_values(values: &mut [u64], source: u64, target: u64) -> Result<()> {
    let source_count = values.iter().filter(|value| **value == source).count();
    if source_count > 1 || (source_count == 1 && values.contains(&target)) {
        return Err(Error::InvalidFormat(format!(
            "IWA metadata cannot replace object reference {source} with {target} unambiguously"
        )));
    }
    for value in values {
        if *value == source {
            *value = target;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

    #[test]
    fn moves_table_with_name_and_sheet_selectors() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let target = editor.add_empty_sheet("Archive").unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;

        let moved = editor
            .move_table(
                TableSelector::name("Revenue"),
                SheetSelector::name("Archive"),
            )
            .unwrap();
        assert_eq!(moved.object_id, table_id);
        assert_eq!(
            find_table_owner(editor.package(), table_id)
                .unwrap()
                .sheet_id,
            target.object_id
        );
        assert!(
            editor
                .move_table(TableSelector::index(1), SheetSelector::index(0))
                .is_err()
        );
    }
}
