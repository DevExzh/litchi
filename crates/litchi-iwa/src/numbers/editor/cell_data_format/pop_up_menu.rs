//! Native Pop-Up Menu object serialization and lifecycle.
#![allow(deprecated)] // Native archives still emit these fields; strict reads must validate them.

use super::*;
use litchi_numbers::cell::data_format::pop_up_menu::{InitialSelection, PopUpMenu};

pub(super) const MODEL_MESSAGE_TYPE: u32 = 6_206;

pub(super) fn read_model(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    starts_with_first_item: bool,
) -> Result<PopUpMenu> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Pop-Up Menu model object {identifier} is missing"))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Pop-Up Menu model object {identifier} is missing"))
    })?;
    if object.messages.len() != 1 || object.messages[0].type_ != MODEL_MESSAGE_TYPE {
        return Err(Error::InvalidFormat(format!(
            "Pop-Up Menu model object {identifier} has a non-canonical payload"
        )));
    }
    let model = tst::PopUpMenuModel::decode(object.messages[0].data.as_slice())?;
    if !model.item.is_empty() {
        return Err(Error::InvalidFormat(
            "Pop-Up Menu model uses deprecated item storage".to_owned(),
        ));
    }
    let (nil, items) = model.tsce_item.split_first().ok_or_else(|| {
        Error::InvalidFormat("Pop-Up Menu model has no blank sentinel".to_owned())
    })?;
    if nil.cell_value_type != tsce::cell_value_archive::CellValueType::NilType as i32
        || nil.boolean_value.is_some()
        || nil.date_value.is_some()
        || nil.number_value.is_some()
        || nil.string_value.is_some()
        || nil.error_value.is_some()
    {
        return Err(Error::InvalidFormat(
            "Pop-Up Menu model has an invalid blank sentinel".to_owned(),
        ));
    }
    let items = items
        .iter()
        .map(|value| {
            if value.cell_value_type != tsce::cell_value_archive::CellValueType::StringType as i32
                || value.boolean_value.is_some()
                || value.date_value.is_some()
                || value.number_value.is_some()
                || value.error_value.is_some()
            {
                return Err(Error::InvalidFormat(
                    "Pop-Up Menu contains a non-text item".to_owned(),
                ));
            }
            let string = value.string_value.as_ref().ok_or_else(|| {
                Error::InvalidFormat("Pop-Up Menu text item has no value".to_owned())
            })?;
            let expected_format = text_format_to_native();
            if string.format != expected_format
                || string.format_is_implicit.is_some()
                || string.format_is_explicit != Some(false)
                || string.is_regex != Some(false)
                || string.is_case_sensitive_regex != Some(false)
            {
                return Err(Error::InvalidFormat(
                    "Pop-Up Menu text item contains non-canonical options".to_owned(),
                ));
            }
            Ok(string.value.as_str())
        })
        .collect::<Result<Vec<_>>>()?;
    PopUpMenu::new(items)
        .map(|format| {
            format.with_initial_selection(if starts_with_first_item {
                InitialSelection::FirstItem
            } else {
                InitialSelection::Blank
            })
        })
        .map_err(Into::into)
}

pub(super) fn create_model(
    package: &mut IWorkPackage,
    archive_name: &str,
    format: &PopUpMenu,
) -> Result<u64> {
    let identifier = next_object_identifier(package)?;
    let model = tst::PopUpMenuModel {
        item: Vec::new(),
        tsce_item: std::iter::once(tsce::CellValueArchive {
            cell_value_type: tsce::cell_value_archive::CellValueType::NilType as i32,
            ..Default::default()
        })
        .chain(format.items().iter().map(|item| tsce::CellValueArchive {
            cell_value_type: tsce::cell_value_archive::CellValueType::StringType as i32,
            string_value: Some(tsce::StringCellValueArchive {
                value: item.as_str().to_owned(),
                format: text_format_to_native(),
                format_is_implicit: None,
                format_is_explicit: Some(false),
                is_regex: Some(false),
                is_case_sensitive_regex: Some(false),
            }),
            ..Default::default()
        }))
        .collect(),
    };
    package.update_archive(archive_name, |archive| {
        archive.insert_object(ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: MODEL_MESSAGE_TYPE,
                data: model.encode_to_vec(),
            }],
        )?)?;
        Ok(())
    })?;
    set_package_last_object_identifier(package, identifier)?;
    Ok(identifier)
}

pub(super) fn remove_model_if_unreferenced(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
) -> Result<()> {
    if !storage::package_references_object(package, locations, identifier)? {
        model::remove_object_or_empty_entry(package, locations, identifier)?;
    }
    Ok(())
}
