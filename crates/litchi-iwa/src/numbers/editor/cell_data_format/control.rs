//! Native control-cell-spec table lifecycle for interactive data formats.

use super::*;

const DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const CHECKBOX_INTERACTION_TYPE: u32 = 8;
const STAR_RATING_INTERACTION_TYPE: u32 = 6;
const STAR_RATING_MINIMUM: f64 = 0.0;
const STAR_RATING_MAXIMUM: f64 = 5.0;
const STAR_RATING_INCREMENT: f64 = 1.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ControlCellSpecKind {
    Checkbox,
    StarRating,
}

pub(super) fn acquire_spec(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    current_identifier: Option<u32>,
    kind: ControlCellSpecKind,
) -> Result<u32> {
    let expected = cell_spec(kind);
    let table_id = ensure_control_table(package, location)?;
    let locations = storage::object_locations(package)?;
    let resolved = storage::resolve_table_data_list(
        package,
        &locations,
        table_id,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    let reusable = resolved
        .entries
        .iter()
        .find(|entry| entry.entry.cell_spec.as_ref() == Some(&expected));
    if let Some(reusable) = reusable {
        if reusable.entry.refcount == 0 {
            return Err(Error::InvalidFormat(
                "Numbers control-cell-spec table contains a zero-reference entry".to_owned(),
            ));
        }
        if current_identifier != Some(reusable.entry.key) {
            storage::increment_table_data_list_entry(
                package,
                &locations,
                &resolved,
                reusable,
                tst::table_data_list::ListType::ControlCellSpec,
            )?;
        }
        return Ok(reusable.entry.key);
    }

    let key = storage::next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers control-cell-spec table object {table_id} is missing"
            ))
        })?;
        let index = storage::table_data_list_message_index(
            object,
            tst::table_data_list::ListType::ControlCellSpec,
        )
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {table_id} has no Numbers control-cell-spec list payload"
            ))
        })?;
        let previous = TableDataList::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        current.next_list_id = key.checked_add(1).ok_or_else(|| {
            Error::ParseError("Numbers control-cell-spec identifier overflow".to_owned())
        })?;
        current.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            cell_spec: Some(expected),
            ..Default::default()
        });
        let data = storage::rewrite_table_data_list_wire(
            object.messages[index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[index].type_;
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })?;
    Ok(key)
}

pub(super) fn validate_spec(
    package: &IWorkPackage,
    location: &model::CellLocation,
    identifier: u32,
    kind: ControlCellSpecKind,
) -> Result<()> {
    let label = control_label(kind);
    let table_id = location
        .descriptor
        .model
        .base_data_store
        .control_cell_spec_table
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("{label} cell has no control-cell-spec table"))
        })?;
    let resolved = storage::resolve_table_data_list(
        package,
        &location.object_locations,
        table_id,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    let entry = resolved
        .entries
        .iter()
        .find(|entry| entry.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{label} cell references missing control spec {identifier}"
            ))
        })?;
    if entry.entry.refcount == 0 || entry.entry.cell_spec.as_ref() != Some(&cell_spec(kind)) {
        return Err(Error::InvalidFormat(format!(
            "{label} control spec {identifier} is invalid"
        )));
    }
    Ok(())
}

pub(super) fn release_spec(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    identifier: u32,
) -> Result<()> {
    let table_id = location
        .descriptor
        .model
        .base_data_store
        .control_cell_spec_table
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat("Interactive cell has no control-cell-spec table".to_owned())
        })?;
    let resolved = storage::resolve_table_data_list(
        package,
        &location.object_locations,
        table_id,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    let entry = resolved
        .entries
        .iter()
        .find(|entry| entry.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Interactive cell references missing control spec {identifier}"
            ))
        })?;
    storage::decrement_table_data_list_entry(
        package,
        &location.object_locations,
        &resolved,
        entry,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    Ok(())
}

fn ensure_control_table(package: &mut IWorkPackage, location: &model::CellLocation) -> Result<u64> {
    if let Some(reference) = &location
        .descriptor
        .model
        .base_data_store
        .control_cell_spec_table
    {
        storage::resolve_table_data_list(
            package,
            &location.object_locations,
            reference.identifier,
            tst::table_data_list::ListType::ControlCellSpec,
        )?;
        return Ok(reference.identifier);
    }

    let identifier = next_object_identifier(package)?;
    let model_archive = location
        .object_locations
        .get(&location.descriptor.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table object {} is missing",
                location.descriptor.object_id
            ))
        })?
        .clone();
    package.update_archive(&model_archive, |archive| {
        archive.insert_object(ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: DATA_LIST_MESSAGE_TYPE,
                data: TableDataList {
                    list_type: tst::table_data_list::ListType::ControlCellSpec as i32,
                    next_list_id: 1,
                    entries: Vec::new(),
                    segments: Vec::new(),
                    is_new_for_bnc: Some(true),
                }
                .encode_to_vec(),
            }],
        )?)?;
        let object = archive
            .object_mut(location.descriptor.object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table object {} is missing",
                    location.descriptor.object_id
                ))
            })?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_000 || message.type_ == 6_001)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers table-model payload",
                    location.descriptor.object_id
                ))
            })?;
        let previous = TableModelArchive::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        current.base_data_store.control_cell_spec_table = Some(tsp::Reference {
            identifier,
            ..Default::default()
        });
        let data = storage::rewrite_table_model_control_cell_spec_table_wire(
            object.messages[index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[index].type_;
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        if !object.archive_info.message_infos[index]
            .object_references
            .contains(&identifier)
        {
            object.archive_info.message_infos[index]
                .object_references
                .push(identifier);
        }
        Ok(())
    })?;
    set_package_last_object_identifier(package, identifier)?;
    Ok(identifier)
}

fn cell_spec(kind: ControlCellSpecKind) -> tst::CellSpecArchive {
    match kind {
        ControlCellSpecKind::Checkbox => tst::CellSpecArchive {
            interaction_type: CHECKBOX_INTERACTION_TYPE,
            ..Default::default()
        },
        ControlCellSpecKind::StarRating => tst::CellSpecArchive {
            interaction_type: STAR_RATING_INTERACTION_TYPE,
            range_control_min: Some(STAR_RATING_MINIMUM),
            range_control_max: Some(STAR_RATING_MAXIMUM),
            range_control_inc: Some(STAR_RATING_INCREMENT),
            ..Default::default()
        },
    }
}

const fn control_label(kind: ControlCellSpecKind) -> &'static str {
    match kind {
        ControlCellSpecKind::Checkbox => "Checkbox",
        ControlCellSpecKind::StarRating => "Star Rating",
    }
}
