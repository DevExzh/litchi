//! Lossless named paragraph-style renaming.

use crate::archive::RawMessage;
use crate::wire::patch_nested_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

use super::paragraph_alignment::native;
use super::paragraph_following_style::{NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
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
        .any(|style| style.id() != target && style.name() == name.as_str())
    {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style name {:?} already exists",
            name.as_str()
        )));
    }
    if let Some(existing) = current
        .iter()
        .find(|style| style.id() == target && style.name() == name.as_str())
    {
        return Ok(existing.clone());
    }

    let location = native::locate_style(package, target.get())?;
    let mut staged = package.clone();
    staged.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(target.get()).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork paragraph style {} is missing", target.get()))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {} must have exactly one paragraph-style payload",
                target.get()
            )));
        };
        let message_index = *message_index;
        let data = patch_nested_length_delimited_field(
            &object.messages[message_index].data,
            &[STYLE_SUPER_FIELD, STYLE_NAME_FIELD],
            location.style.super_.name.is_some(),
            Some(name.as_str().as_bytes()),
        )?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;

    let renamed = NamedParagraphStyle::new(target, name.as_str().to_owned())?;
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
