//! Native control-cell-spec table lifecycle for interactive data formats.

use super::*;
use crate::table_cell_data_format::{
    TableCellPopUpMenuFormat, TableCellPopUpMenuInitialSelection, TableCellSliderRange,
    TableCellStepperRange,
};

const DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const CHECKBOX_INTERACTION_TYPE: u32 = 8;
const STAR_RATING_INTERACTION_TYPE: u32 = 6;
const STAR_RATING_MINIMUM: f64 = 0.0;
const STAR_RATING_MAXIMUM: f64 = 5.0;
const STAR_RATING_INCREMENT: f64 = 1.0;
const SLIDER_INTERACTION_TYPE: u32 = 5;
const STEPPER_INTERACTION_TYPE: u32 = 4;
const POP_UP_MENU_INTERACTION_TYPE: u32 = 7;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum ControlCellSpecKind {
    Checkbox,
    StarRating,
    Slider(TableCellSliderRange),
    Stepper(TableCellStepperRange),
    PopUpMenu(TableCellPopUpMenuFormat),
}

pub(super) fn acquire_spec(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    current_identifier: Option<u32>,
    kind: ControlCellSpecKind,
) -> Result<u32> {
    let expected = cell_spec(&kind);
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

pub(super) fn acquire_pop_up_menu_spec(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    current_identifier: Option<u32>,
    format: &TableCellPopUpMenuFormat,
) -> Result<u32> {
    let table_id = ensure_control_table(package, location)?;
    let locations = storage::object_locations(package)?;
    let resolved = storage::resolve_table_data_list(
        package,
        &locations,
        table_id,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    let mut reusable_model_identifier = None;
    for entry in &resolved.entries {
        let Some(spec) = entry.entry.cell_spec.as_ref() else {
            continue;
        };
        if spec.interaction_type != POP_UP_MENU_INTERACTION_TYPE {
            continue;
        }
        if entry.entry.refcount == 0 {
            return Err(Error::InvalidFormat(
                "Numbers control-cell-spec table contains a zero-reference entry".to_owned(),
            ));
        }
        let Some(reference) = spec.chooser_control_popup_model.as_ref() else {
            continue;
        };
        let existing = pop_up_menu::read_model(
            package,
            &locations,
            reference.identifier,
            spec.chooser_control_start_w_first == Some(true),
        )?;
        if existing.items() == format.items() {
            reusable_model_identifier = Some(reference.identifier);
        }
        if &existing == format {
            if current_identifier != Some(entry.entry.key) {
                storage::increment_table_data_list_entry(
                    package,
                    &locations,
                    &resolved,
                    entry,
                    tst::table_data_list::ListType::ControlCellSpec,
                )?;
            }
            return Ok(entry.entry.key);
        }
    }

    let model_identifier = reusable_model_identifier.map_or_else(
        || pop_up_menu::create_model(package, &resolved.table_archive, format),
        Ok,
    )?;
    let spec = tst::CellSpecArchive {
        interaction_type: POP_UP_MENU_INTERACTION_TYPE,
        chooser_control_popup_model: Some(tsp::Reference {
            identifier: model_identifier,
            ..Default::default()
        }),
        chooser_control_start_w_first: Some(matches!(
            format.initial_selection(),
            TableCellPopUpMenuInitialSelection::FirstItem
        )),
        ..Default::default()
    };
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
            cell_spec: Some(spec),
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
        storage::add_message_object_reference(object, index, model_identifier, model_identifier);
        Ok(())
    })?;
    Ok(key)
}

pub(super) fn read_spec(
    package: &IWorkPackage,
    location: &model::CellLocation,
    identifier: u32,
) -> Result<ControlCellSpecKind> {
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
    if entry.entry.refcount == 0 {
        return Err(Error::InvalidFormat(format!(
            "Interactive control spec {identifier} has no references"
        )));
    }
    let spec = entry.entry.cell_spec.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Interactive control spec {identifier} has no payload"
        ))
    })?;
    if spec.interaction_type == POP_UP_MENU_INTERACTION_TYPE {
        let reference = spec.chooser_control_popup_model.as_ref().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Interactive control spec {identifier} has no Pop-Up Menu model"
            ))
        })?;
        let format = pop_up_menu::read_model(
            package,
            &location.object_locations,
            reference.identifier,
            spec.chooser_control_start_w_first == Some(true),
        )?;
        let mut expected = spec.clone();
        expected.chooser_control_popup_model = None;
        let mut canonical = tst::CellSpecArchive {
            interaction_type: POP_UP_MENU_INTERACTION_TYPE,
            chooser_control_start_w_first: spec.chooser_control_start_w_first,
            ..Default::default()
        };
        if !matches!(canonical.chooser_control_start_w_first, Some(true | false)) {
            canonical.chooser_control_start_w_first = Some(true);
        }
        if expected != canonical {
            return Err(Error::InvalidFormat(format!(
                "Interactive control spec {identifier} has non-canonical Pop-Up Menu options"
            )));
        }
        return Ok(ControlCellSpecKind::PopUpMenu(format));
    }
    parse_cell_spec(spec).map_err(|error| {
        Error::InvalidFormat(format!(
            "Interactive control spec {identifier} is invalid: {error}"
        ))
    })
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
    let popup_model = entry
        .entry
        .cell_spec
        .as_ref()
        .and_then(|spec| spec.chooser_control_popup_model.as_ref())
        .map(|reference| reference.identifier);
    let removed = storage::decrement_table_data_list_entry(
        package,
        &location.object_locations,
        &resolved,
        entry,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    if removed && let Some(identifier) = popup_model {
        pop_up_menu::remove_model_if_unreferenced(package, &location.object_locations, identifier)?;
    }
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

fn cell_spec(kind: &ControlCellSpecKind) -> tst::CellSpecArchive {
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
        ControlCellSpecKind::Slider(range) => tst::CellSpecArchive {
            interaction_type: SLIDER_INTERACTION_TYPE,
            range_control_min: Some(range.minimum()),
            range_control_max: Some(range.maximum()),
            range_control_inc: Some(range.increment()),
            ..Default::default()
        },
        ControlCellSpecKind::Stepper(range) => tst::CellSpecArchive {
            interaction_type: STEPPER_INTERACTION_TYPE,
            range_control_min: Some(range.minimum()),
            range_control_max: Some(range.maximum()),
            range_control_inc: Some(range.increment()),
            ..Default::default()
        },
        ControlCellSpecKind::PopUpMenu(_) => {
            unreachable!("Pop-Up Menu specs require a model reference")
        },
    }
}

const fn control_label(kind: &ControlCellSpecKind) -> &'static str {
    match kind {
        ControlCellSpecKind::Checkbox => "Checkbox",
        ControlCellSpecKind::StarRating => "Star Rating",
        ControlCellSpecKind::Slider(_) => "Slider",
        ControlCellSpecKind::Stepper(_) => "Stepper",
        ControlCellSpecKind::PopUpMenu(_) => "Pop-Up Menu",
    }
}

fn parse_cell_spec(
    spec: &tst::CellSpecArchive,
) -> std::result::Result<ControlCellSpecKind, String> {
    let kind = match spec.interaction_type {
        CHECKBOX_INTERACTION_TYPE => ControlCellSpecKind::Checkbox,
        STAR_RATING_INTERACTION_TYPE => ControlCellSpecKind::StarRating,
        SLIDER_INTERACTION_TYPE => {
            let minimum = spec
                .range_control_min
                .ok_or_else(|| "Slider has no minimum".to_owned())?;
            let maximum = spec
                .range_control_max
                .ok_or_else(|| "Slider has no maximum".to_owned())?;
            let increment = spec
                .range_control_inc
                .ok_or_else(|| "Slider has no increment".to_owned())?;
            ControlCellSpecKind::Slider(
                TableCellSliderRange::new(minimum, maximum, increment)
                    .map_err(|error| error.to_string())?,
            )
        },
        STEPPER_INTERACTION_TYPE => {
            let minimum = spec
                .range_control_min
                .ok_or_else(|| "Stepper has no minimum".to_owned())?;
            let maximum = spec
                .range_control_max
                .ok_or_else(|| "Stepper has no maximum".to_owned())?;
            let increment = spec
                .range_control_inc
                .ok_or_else(|| "Stepper has no increment".to_owned())?;
            ControlCellSpecKind::Stepper(
                TableCellStepperRange::new(minimum, maximum, increment)
                    .map_err(|error| error.to_string())?,
            )
        },
        value => return Err(format!("unsupported interaction type {value}")),
    };
    if spec != &cell_spec(&kind) {
        return Err(format!(
            "{} contains non-canonical options",
            control_label(&kind)
        ));
    }
    Ok(kind)
}
