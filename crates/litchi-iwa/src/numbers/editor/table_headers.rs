//! Native conversion for archive-free Numbers table header semantics.

use super::*;

mod wire;

use litchi_numbers::table::headers::{Count, Settings};
use wire::{read_table_header_settings_wire, write_table_header_settings_wire};

fn count_from_native(count: u32, label: &str) -> Result<Count> {
    let count = usize::try_from(count).map_err(|_| {
        Error::InvalidFormat(format!(
            "Numbers table {label} count {count} does not fit in usize"
        ))
    })?;
    Count::new(count).map_err(|_| {
        Error::InvalidFormat(format!(
            "Numbers table {label} count {count} is outside the native 1..=5 range"
        ))
    })
}

fn count_as_native(count: Count) -> u32 {
    count.get() as u32
}

pub(super) fn settings_from_model(model: &TableModelArchive) -> Result<Settings> {
    Ok(Settings {
        header_rows: model
            .number_of_header_rows
            .map(|count| count_from_native(count, "header row"))
            .transpose()?,
        header_columns: model
            .number_of_header_columns
            .map(|count| count_from_native(count, "header column"))
            .transpose()?,
        footer_rows: model
            .number_of_footer_rows
            .map(|count| count_from_native(count, "footer row"))
            .transpose()?,
        header_rows_frozen: model.header_rows_frozen,
        header_columns_frozen: model.header_columns_frozen,
        repeating_header_rows_enabled: model.repeating_header_rows_enabled,
        repeating_header_columns_enabled: model.repeating_header_columns_enabled,
    })
}

fn read_table_header_settings(
    package: &IWorkPackage,
    descriptor: &TableDescriptor,
) -> Result<Settings> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let message_index = table_model_message_index(object, descriptor.object_id)?;
    let settings = read_table_header_settings_wire(
        object.messages[message_index].data.as_slice(),
        &descriptor.model,
    )?;
    validate_table_header_settings(&descriptor.model, settings).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers table model {} has invalid stored header settings: {error}",
            descriptor.object_id
        ))
    })?;
    Ok(settings)
}

pub(super) fn read_attached_table_header_settings(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Settings> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    read_table_header_settings(package, &descriptor)
}

pub(super) fn set_attached_table_header_settings(
    package: &mut IWorkPackage,
    table_id: u64,
    settings: Settings,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    validate_table_header_settings(&descriptor.model, settings)?;
    if read_table_header_settings(package, &descriptor)? == settings {
        return Ok(());
    }
    let locations = object_locations(package)?;
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
        })?;
        let message_index = table_model_message_index(object, table_id)?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let model = TableModelArchive::decode(original)?;
        let data = write_table_header_settings_wire(original, &model, settings)?;
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

fn table_model_message_index(object: &ArchiveObject, table_id: u64) -> Result<usize> {
    let matches = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            matches!(message.type_, 6000 | 6001)
                .then(|| {
                    TableModelArchive::decode(message.data.as_slice())
                        .ok()
                        .map(|_| index)
                })
                .flatten()
        })
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [index] => Ok(*index),
        _ => Err(Error::InvalidFormat(format!(
            "Numbers table model object {table_id} must contain exactly one decodable payload, found {}",
            matches.len()
        ))),
    }
}

fn validate_table_header_settings(model: &TableModelArchive, settings: Settings) -> Result<()> {
    let header_rows = settings.header_row_count();
    let footer_rows = settings.footer_row_count();
    let rows = model.number_of_rows as usize;
    if header_rows + footer_rows > rows {
        return Err(Error::ParseError(format!(
            "Numbers header rows ({header_rows}) plus footer rows ({footer_rows}) exceed the table's {rows} rows"
        )));
    }
    let header_columns = settings.header_column_count();
    let columns = model.number_of_columns as usize;
    if header_columns > columns {
        return Err(Error::ParseError(format!(
            "Numbers header columns ({header_columns}) exceed the table's {columns} columns"
        )));
    }
    Ok(())
}
