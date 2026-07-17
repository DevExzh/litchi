//! Native plain-highlight object graphs and ownership validation.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::comments::{fresh_comment_storage_uuid, insert_comment_storage};
use crate::protobuf::{tsd, tsp, tswp};
use crate::{Error, IWorkPackage, Result};

pub(super) const HIGHLIGHT_MESSAGE_TYPE: u32 = 2_013;
pub(super) const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;

#[derive(Debug, Clone, Copy)]
pub(super) struct PlainHighlightGraph {
    pub(super) comment_storage_id: u64,
    pub(super) author_id: Option<u64>,
}

pub(super) fn validate_highlight_object(
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<u64>> {
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == HIGHLIGHT_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if payloads.is_empty() {
        return Ok(None);
    }
    let [message] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight object {identifier} contains multiple highlight payloads"
        )));
    };
    if object.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight object {identifier} contains unrelated payloads"
        )));
    }
    let highlight = tswp::HighlightArchive::decode(message.data.as_slice())?;
    let uuid = highlight
        .text_attribute_uuid_string
        .as_deref()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork highlight object {identifier} is missing its text-attribute UUID"
            ))
        })?;
    validate_uuid(identifier, uuid)?;
    let comment_storage_id = highlight
        .comment_storage
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork highlight object {identifier} is missing its comment storage"
            ))
        })?
        .identifier;
    if comment_storage_id == 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight object {identifier} has a zero comment-storage identifier"
        )));
    }
    Ok(Some(comment_storage_id))
}

fn validate_uuid(identifier: u64, uuid: &str) -> Result<()> {
    let valid = uuid.len() == 36
        && uuid.bytes().enumerate().all(|(index, byte)| match index {
            8 | 13 | 18 | 23 => byte == b'-',
            _ => byte.is_ascii_hexdigit(),
        });
    if !valid {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight object {identifier} has an invalid text-attribute UUID"
        )));
    }
    Ok(())
}

pub(super) fn validate_plain_highlight_graph(
    package: &IWorkPackage,
    archive_name: &str,
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<PlainHighlightGraph>> {
    let Some(comment_storage_id) = validate_highlight_object(identifier, object)? else {
        return Ok(None);
    };
    let archive = package.archive(archive_name)?;
    let comment_object = archive.object(comment_storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork highlight object {identifier} references missing comment storage {comment_storage_id}"
        ))
    })?;
    if comment_object.messages.len() != 1
        || comment_object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE
    {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight {identifier} comment storage {comment_storage_id} must contain exactly one comment payload"
        )));
    }
    let comment = tsd::CommentStorageArchive::decode(comment_object.messages[0].data.as_slice())?;
    if comment.text.as_deref() != Some("") || !comment.replies.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight {identifier} is comment-backed rather than a plain highlight"
        )));
    }
    if comment
        .creation_date
        .as_ref()
        .is_some_and(|date| !date.seconds.is_finite())
    {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight {identifier} has a non-finite creation date"
        )));
    }
    let uuid = comment.storage_uuid.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork highlight {identifier} comment storage is missing its UUID"
        ))
    })?;
    if uuid.lower == 0 && uuid.upper == 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight {identifier} comment storage has a zero UUID"
        )));
    }
    let author_id = comment.author.map(|reference| reference.identifier);
    if author_id == Some(0) {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight {identifier} has a zero author identifier"
        )));
    }
    Ok(Some(PlainHighlightGraph {
        comment_storage_id,
        author_id,
    }))
}

pub(super) fn new_highlight_object(
    identifier: u64,
    comment_storage_id: u64,
) -> Result<ArchiveObject> {
    let braced = litchi_core::id::generate_guid_braced();
    let uuid = braced
        .strip_prefix('{')
        .and_then(|uuid| uuid.strip_suffix('}'))
        .ok_or_else(|| Error::InvalidFormat("generated UUID is not braced".to_owned()))?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: HIGHLIGHT_MESSAGE_TYPE,
            data: tswp::HighlightArchive {
                comment_storage: Some(tsp::Reference {
                    identifier: comment_storage_id,
                    ..Default::default()
                }),
                text_attribute_uuid_string: Some(uuid.to_owned()),
            }
            .encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0]
        .object_references
        .push(comment_storage_id);
    Ok(object)
}

pub(super) fn insert_plain_comment_storage(
    package: &mut IWorkPackage,
    archive_name: &str,
    comment_storage_id: u64,
    author_id: Option<u64>,
) -> Result<()> {
    let storage_uuid = fresh_comment_storage_uuid(package)?;
    insert_comment_storage(
        package,
        archive_name,
        comment_storage_id,
        String::new(),
        author_id,
        storage_uuid,
    )
}

pub(super) fn require_exclusive_reference(
    package: &IWorkPackage,
    expected_owner: u64,
    identifier: u64,
    label: &str,
) -> Result<()> {
    let mut owners = Vec::new();
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("object in {archive_name} has no identifier"))
            })?;
            if object
                .archive_info
                .message_infos
                .iter()
                .any(|info| info.object_references.contains(&identifier))
            {
                owners.push(object_id);
            }
        }
    }
    if owners != [expected_owner] {
        return Err(Error::InvalidFormat(format!(
            "{label} object {identifier} must be referenced only by object {expected_owner}, found {owners:?}"
        )));
    }
    Ok(())
}

pub(super) fn ensure_no_metadata_reference(package: &IWorkPackage, identifier: u64) -> Result<()> {
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            if object.archive_info.message_infos.iter().any(|info| {
                info.object_references.contains(&identifier)
                    || info
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&identifier))
            }) {
                return Err(Error::InvalidFormat(format!(
                    "highlight graph object {identifier} retains an indexed reference in {archive_name}"
                )));
            }
        }
    }
    Ok(())
}
