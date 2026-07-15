//! Typed title visibility and outline settings for Numbers tables.

use super::*;

mod wire;

use wire::{read_table_title_settings_wire, write_table_title_settings_wire};

const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];

/// Lossless optional title settings stored by a Numbers table model.
///
/// Optional booleans retain their native protobuf presence. Numbers normally
/// omits `visible` when the title is hidden, while documents from other
/// versions may encode an explicit `false` value.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct NumbersTableTitleSettings {
    /// Whether the table title is displayed.
    pub visible: Option<bool>,
    /// Whether the table outline includes the title region.
    pub outlined: Option<bool>,
}

impl NumbersTableTitleSettings {
    /// Return whether the title is effectively visible.
    pub fn is_visible(self) -> bool {
        self.visible.unwrap_or(false)
    }

    /// Return whether the title region is effectively outlined.
    pub fn is_outlined(self) -> bool {
        self.outlined.unwrap_or(false)
    }

    fn from_model(model: &TableModelArchive) -> Self {
        Self {
            visible: model.table_name_enabled,
            outlined: model.table_name_border_enabled,
        }
    }
}

impl NumbersEditor {
    /// Read an attached table's lossless title visibility and outline settings.
    pub fn table_title_settings(&self, table_id: u64) -> Result<NumbersTableTitleSettings> {
        let descriptor = table_descriptor(&self.package, table_id)?;
        read_table_title_settings(&self.package, &descriptor)
    }

    /// Replace an attached table's title visibility and outline settings transactionally.
    pub fn set_table_title_settings(
        &mut self,
        table_id: u64,
        settings: NumbersTableTitleSettings,
    ) -> Result<()> {
        let descriptor = table_descriptor(&self.package, table_id)?;
        if read_table_title_settings(&self.package, &descriptor)? == settings {
            return Ok(());
        }

        let locations = object_locations(&self.package)?;
        let archive_name = locations.get(&table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
        })?;
        let mut staged = self.package.clone();
        staged.update_archive(archive_name, |archive| {
            let object = archive.object_mut(table_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers table model object {table_id} is missing"))
            })?;
            let message_index = table_model_message_index(object, table_id)?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let model = TableModelArchive::decode(original)?;
            let data = write_table_title_settings_wire(original, &model, settings)?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;

        let verified = Self::from_package(staged)?;
        if verified.table_title_settings(table_id)? != settings {
            return Err(Error::InvalidFormat(
                "Numbers table title settings failed round-trip validation".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(())
    }
}

fn table_descriptor(package: &IWorkPackage, table_id: u64) -> Result<TableDescriptor> {
    table_models(package)?
        .into_iter()
        .find(|table| table.object_id == table_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table object {table_id} not found")))
}

fn read_table_title_settings(
    package: &IWorkPackage,
    descriptor: &TableDescriptor,
) -> Result<NumbersTableTitleSettings> {
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
    read_table_title_settings_wire(
        object.messages[message_index].data.as_slice(),
        &descriptor.model,
    )
}

fn table_model_message_index(object: &ArchiveObject, table_id: u64) -> Result<usize> {
    let mut matches = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            TABLE_MODEL_MESSAGE_TYPES
                .contains(&message.type_)
                .then_some(index)
        });
    let Some(index) = matches.next() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers table model object {table_id} must contain exactly one payload, found 0"
        )));
    };
    if matches.next().is_some() {
        let count = 2 + matches.count();
        return Err(Error::InvalidFormat(format!(
            "Numbers table model object {table_id} must contain exactly one payload, found {count}"
        )));
    }
    TableModelArchive::decode(object.messages[index].data.as_slice()).map_err(|error| {
        Error::InvalidFormat(format!(
            "Numbers table model object {table_id} has an invalid payload: {error}"
        ))
    })?;
    Ok(index)
}
