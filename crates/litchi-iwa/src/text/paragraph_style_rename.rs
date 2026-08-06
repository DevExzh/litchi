//! Lossless named paragraph-style renaming.

use crate::archive::RawMessage;
use crate::wire::patch_nested_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

use super::paragraph_alignment::native;
use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName, raw::native_id,
};

const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_NAME_FIELD: u32 = 1;

pub(super) fn rename_named_paragraph_style(
    package: &mut IWorkPackage,
    first_style_id: u64,
    target: ParagraphStyleId,
    name: ParagraphStyleName,
) -> Result<NamedParagraphStyle> {
    native::validate_named_paragraph_style(package, first_style_id, target)?;
    let current = native::named_paragraph_styles(package, first_style_id)?;
    if current
        .iter()
        .any(|style| style.id() != target && style.name().as_str() == name.as_str())
    {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style name {:?} already exists",
            name.as_str()
        )));
    }
    if let Some(existing) = current
        .iter()
        .find(|style| style.id() == target && style.name().as_str() == name.as_str())
    {
        return Ok(existing.clone());
    }

    let location = native::locate_style(package, native_id(target))?;
    let mut staged = package.clone();
    rename_at_location(&mut staged, &location, name.as_str())?;

    let renamed = NamedParagraphStyle::from_owned(target, name.as_str().to_owned())?;
    let matches = native::named_paragraph_styles(&staged, first_style_id)?
        .into_iter()
        .filter(|style| style == &renamed)
        .count();
    if matches != 1 {
        return Err(Error::InvalidFormat(
            "named iWork paragraph style rename failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(renamed)
}

pub(super) fn rename_at_location(
    package: &mut IWorkPackage,
    location: &native::ParagraphStyleLocation,
    name: &str,
) -> Result<()> {
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {} is missing",
                location.object_id
            ))
        })?;
        if object.archive_info.identifier != Some(location.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {} object identity changed unexpectedly",
                location.object_id
            )));
        }
        let Some(message) = object.messages.get(location.message_index) else {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {} anchored payload index {} is missing",
                location.object_id, location.message_index
            )));
        };
        if message.type_ != location.message_type || message.data != location.message.data {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {} anchored payload changed unexpectedly",
                location.object_id
            )));
        }
        let Some(info) = object
            .archive_info
            .message_infos
            .get(location.message_index)
        else {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {} anchored metadata index {} is missing",
                location.object_id, location.message_index
            )));
        };
        let message_length = u32::try_from(message.data.len()).map_err(|_| {
            Error::InvalidFormat("paragraph style payload exceeds u32 length".to_owned())
        })?;
        if info.type_ != location.message_type || info.length != message_length {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {} anchored metadata changed unexpectedly",
                location.object_id
            )));
        }
        let data = patch_nested_length_delimited_field(
            &message.data,
            &[STYLE_SUPER_FIELD, STYLE_NAME_FIELD],
            location.style.super_.name.is_some(),
            Some(name.as_bytes()),
        )?;
        object.replace_message(
            location.message_index,
            RawMessage {
                type_: location.message_type,
                data,
            },
        )?;
        Ok(())
    })
}
