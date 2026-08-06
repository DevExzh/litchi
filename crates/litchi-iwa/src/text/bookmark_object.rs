//! Native bookmark-field object encoding and lossless settings mutation.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::tswp;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

use super::bookmark_types::{TextBookmarkSettings, TextBookmarkVisibility};
use super::smart_field_object::{generated_text_attribute_uuid, validate_text_attribute_uuid};
use litchi_iwa_text::bookmark::Name;

const BOOKMARK_NAME_FIELD: u32 = 2;
const BOOKMARK_HIDDEN_FIELD: u32 = 4;
const RANGED_BOOKMARK: u32 = 1;
pub(super) const BOOKMARK_MESSAGE_TYPE: u32 = 2_035;

pub(super) fn validate_bookmark_object(
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<TextBookmarkSettings>> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == BOOKMARK_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if payloads.is_empty() {
        return Ok(None);
    }
    let [message] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork bookmark object {identifier} contains multiple bookmark payloads"
        )));
    };
    if object.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork bookmark object {identifier} contains unrelated payloads"
        )));
    }
    let bookmark = tswp::BookmarkFieldArchive::decode(message.data.as_slice())?;
    let uuid = bookmark
        .super_
        .as_ref()
        .and_then(|smart_field| smart_field.text_attribute_uuid_string.as_deref())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork bookmark object {identifier} is missing its text-attribute UUID"
            ))
        })?;
    validate_text_attribute_uuid(identifier, "bookmark", uuid)?;
    if bookmark.ranged != Some(RANGED_BOOKMARK) {
        return Err(Error::InvalidFormat(format!(
            "iWork bookmark object {identifier} is not a ranged bookmark"
        )));
    }
    let name = bookmark.name.map(Name::try_from).transpose()?;
    let settings = TextBookmarkSettings::new()
        .with_visibility(visibility_from_native(bookmark.hidden.unwrap_or_default()));
    Ok(Some(match name {
        Some(name) => settings.with_name(name),
        None => settings,
    }))
}

pub(super) fn new_bookmark_object(
    identifier: u64,
    settings: &TextBookmarkSettings,
) -> Result<ArchiveObject> {
    let uuid = generated_text_attribute_uuid()?;
    let bookmark = tswp::BookmarkFieldArchive {
        super_: Some(tswp::SmartFieldArchive {
            text_attribute_uuid_string: Some(uuid),
        }),
        name: settings.name().map(|name| name.as_str().to_owned()),
        ranged: Some(RANGED_BOOKMARK),
        hidden: Some(visibility_to_native(settings.visibility())),
    };
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: BOOKMARK_MESSAGE_TYPE,
            data: bookmark.encode_to_vec(),
        }],
    )?)
}

pub(super) fn patch_bookmark_settings(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifier: u64,
    settings: &TextBookmarkSettings,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork bookmark object {identifier} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == BOOKMARK_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork bookmark object {identifier} must contain exactly one bookmark payload"
            )));
        };
        if object.messages.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork bookmark object {identifier} contains unrelated payloads"
            )));
        }
        let original = &object.messages[*index];
        let bookmark = tswp::BookmarkFieldArchive::decode(original.data.as_slice())?;
        let mut data = patch_length_delimited_field(
            &original.data,
            BOOKMARK_NAME_FIELD,
            bookmark.name.is_some(),
            settings.name().map(|name| name.as_str().as_bytes()),
        )?;
        data = patch_varint_field(
            &data,
            BOOKMARK_HIDDEN_FIELD,
            bookmark.hidden.is_some(),
            Some(u64::from(visibility_to_native(settings.visibility()))),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: BOOKMARK_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn visibility_from_native(raw: u32) -> TextBookmarkVisibility {
    match raw {
        0 => TextBookmarkVisibility::Visible,
        1 => TextBookmarkVisibility::Hidden,
        unknown => TextBookmarkVisibility::Unknown(unknown),
    }
}

fn visibility_to_native(visibility: TextBookmarkVisibility) -> u32 {
    match visibility {
        TextBookmarkVisibility::Visible => 0,
        TextBookmarkVisibility::Hidden => 1,
        TextBookmarkVisibility::Unknown(raw) => raw,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn visibility_mapping_is_lossless_at_the_native_boundary() {
        for raw in [0, 1, 7, u32::MAX] {
            assert_eq!(visibility_to_native(visibility_from_native(raw)), raw);
        }
    }

    #[test]
    fn bookmark_leaf_validation_maps_to_the_iwa_error() {
        let error: Error = litchi_iwa_text::bookmark::Error::NameTooLong.into();
        assert!(
            matches!(error, Error::InvalidFormat(message) if message.contains("bookmark name"))
        );
    }

    #[test]
    fn semantic_settings_round_trip_through_the_native_object_adapter() {
        let settings = TextBookmarkSettings::new()
            .with_name(Name::new("Methods").unwrap())
            .with_visibility(TextBookmarkVisibility::Unknown(9));
        let object = new_bookmark_object(17, &settings).unwrap();
        assert_eq!(
            validate_bookmark_object(17, &object).unwrap(),
            Some(settings)
        );
    }
}
