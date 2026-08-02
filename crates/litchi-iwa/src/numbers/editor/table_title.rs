//! Typed title visibility and outline settings for native iWork tables.

use super::*;

mod wire;

use wire::{read_table_title_settings_wire, write_table_title_settings_wire};

const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;

/// Lossless optional title settings stored by a native iWork table model.
///
/// Optional booleans retain their native protobuf presence. iWork normally
/// omits `visible` when the title is hidden, while documents from other app
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
        table_title_settings_in_package(&self.package, table_id)
    }

    /// Replace an attached table's title visibility and outline settings transactionally.
    pub fn set_table_title_settings(
        &mut self,
        table_id: u64,
        settings: NumbersTableTitleSettings,
    ) -> Result<()> {
        if table_title_settings_in_package(&self.package, table_id)? == settings {
            return Ok(());
        }
        let mut staged = self.package.clone();
        set_table_title_settings_in_package(&mut staged, table_id, settings)?;

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

pub(crate) fn table_title_settings_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<NumbersTableTitleSettings> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    read_table_title_settings(package, &descriptor)
}

pub(crate) fn set_table_title_settings_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    settings: NumbersTableTitleSettings,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    if read_table_title_settings(package, &descriptor)? == settings {
        return Ok(());
    }
    let locations = object_locations(package)?;
    if settings.is_visible() {
        validate_visible_title_prerequisites(package, &locations, &descriptor)?;
    }
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model object {table_id} is missing"))
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
    if table_title_settings_in_package(package, table_id)? != settings {
        return Err(Error::InvalidFormat(
            "iWork table title settings failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn validate_visible_title_prerequisites(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    descriptor: &TableDescriptor,
) -> Result<()> {
    let height = descriptor.model.table_name_height.ok_or_else(|| {
        Error::ParseError(format!(
            "Cannot safely enable the title for iWork table {} because its title height is absent",
            descriptor.object_id
        ))
    })?;
    if !height.is_finite() || height < 0.0 {
        return Err(Error::InvalidFormat(format!(
            "iWork table {} has invalid title height {height}",
            descriptor.object_id
        )));
    }
    let paragraph_style = descriptor
        .model
        .table_name_style
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Cannot safely enable the title for iWork table {} because its title text style is absent",
                descriptor.object_id
            ))
        })?;
    let shape_style = descriptor
        .model
        .table_name_shape_style
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "Cannot safely enable the title for iWork table {} because its title shape style is absent",
                descriptor.object_id
            ))
        })?;
    require_title_style::<tswp::ParagraphStyleArchive>(
        package,
        locations,
        paragraph_style,
        PARAGRAPH_STYLE_MESSAGE_TYPE,
        "text",
    )?;
    require_title_style::<tswp::ShapeStyleArchive>(
        package,
        locations,
        shape_style,
        SHAPE_STYLE_MESSAGE_TYPE,
        "shape",
    )?;

    let model_archive_name = locations.get(&descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let model_archive = package.archive(model_archive_name)?;
    let model_object = model_archive.object(descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let message_index = table_model_message_index(model_object, descriptor.object_id)?;
    let references = &model_object.archive_info.message_infos[message_index].object_references;
    for (identifier, label) in [(paragraph_style, "text"), (shape_style, "shape")] {
        if !references.contains(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork table {} title {label} style {identifier} is missing from its native reference metadata",
                descriptor.object_id
            )));
        }
    }
    Ok(())
}

fn require_title_style<M: Message + Default>(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    message_type: u32,
    label: &str,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table title {label} style object {identifier} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table title {label} style object {identifier} is missing"
        ))
    })?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type);
    let Some(message) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "iWork table title {label} style object {identifier} must contain exactly one native payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork table title {label} style object {identifier} must contain exactly one native payload"
        )));
    }
    M::decode(message.data.as_slice()).map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork table title {label} style object {identifier} is invalid: {error}"
        ))
    })?;
    Ok(())
}

fn read_table_title_settings(
    package: &IWorkPackage,
    descriptor: &TableDescriptor,
) -> Result<NumbersTableTitleSettings> {
    let locations = object_locations(package)?;
    let archive_name = locations.get(&descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table model object {} is missing",
            descriptor.object_id
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(descriptor.object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork table model object {} is missing",
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
            "iWork table model object {table_id} must contain exactly one payload, found 0"
        )));
    };
    if matches.next().is_some() {
        let count = 2 + matches.count();
        return Err(Error::InvalidFormat(format!(
            "iWork table model object {table_id} must contain exactly one payload, found {count}"
        )));
    }
    TableModelArchive::decode(object.messages[index].data.as_slice()).map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork table model object {table_id} has an invalid payload: {error}"
        ))
    })?;
    Ok(index)
}
