//! Native ranged-annotation object graphs and ownership validation.

use std::collections::HashSet;

use prost::Message;
use litchi_iwa_common::comment::Uuid;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::comments::{fresh_comment_storage_uuid, insert_comment_storage};
use crate::protobuf::{tsd, tsp, tswp};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

pub(super) const HIGHLIGHT_MESSAGE_TYPE: u32 = 2_013;
pub(super) const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;

#[derive(Debug, Clone)]
pub(super) struct AnnotationReplyGraph {
    pub(super) storage_id: u64,
    pub(super) body: String,
    pub(super) creation_date_seconds: Option<f64>,
    pub(super) author_id: Option<u64>,
    pub(super) storage_uuid: Uuid,
}

#[derive(Debug, Clone)]
pub(super) struct AnnotationGraph {
    pub(super) comment_storage_id: u64,
    pub(super) body: String,
    pub(super) creation_date_seconds: Option<f64>,
    pub(super) author_id: Option<u64>,
    pub(super) storage_uuid: Uuid,
    pub(super) replies: Vec<AnnotationReplyGraph>,
}

impl AnnotationGraph {
    pub(super) fn is_plain_highlight(&self) -> bool {
        self.body.is_empty()
    }

    pub(super) fn author_ids(&self) -> impl Iterator<Item = u64> + '_ {
        self.author_id
            .into_iter()
            .chain(self.replies.iter().filter_map(|reply| reply.author_id))
    }
}

struct CommentNode {
    body: String,
    creation_date_seconds: Option<f64>,
    author_id: Option<u64>,
    storage_uuid: Uuid,
    reply_ids: Vec<u64>,
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

#[cfg(test)]
pub(super) fn validate_annotation_graph(
    package: &IWorkPackage,
    archive_name: &str,
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<AnnotationGraph>> {
    let archive = package.archive(archive_name)?;
    validate_annotation_graph_in_archive(&archive, identifier, object)
}

pub(super) fn validate_annotation_graph_in_archive(
    archive: &Archive,
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<AnnotationGraph>> {
    let Some(comment_storage_id) = validate_highlight_object(identifier, object)? else {
        return Ok(None);
    };
    let root = validate_comment_node(archive, identifier, comment_storage_id)?;
    if root.body.is_empty() && !root.reply_ids.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "plain highlight {identifier} unexpectedly owns comment replies"
        )));
    }

    let mut seen = HashSet::with_capacity(root.reply_ids.len());
    let mut replies = Vec::with_capacity(root.reply_ids.len());
    for reply_id in root.reply_ids {
        if reply_id == comment_storage_id || !seen.insert(reply_id) {
            return Err(Error::InvalidFormat(format!(
                "text annotation {identifier} has a cyclic or duplicate reply {reply_id}"
            )));
        }
        let reply = validate_comment_node(archive, identifier, reply_id)?;
        if reply.body.is_empty() || !reply.reply_ids.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "text annotation {identifier} reply {reply_id} must be nonempty and cannot own nested replies"
            )));
        }
        replies.push(AnnotationReplyGraph {
            storage_id: reply_id,
            body: reply.body,
            creation_date_seconds: reply.creation_date_seconds,
            author_id: reply.author_id,
            storage_uuid: reply.storage_uuid,
        });
    }
    Ok(Some(AnnotationGraph {
        comment_storage_id,
        body: root.body,
        creation_date_seconds: root.creation_date_seconds,
        author_id: root.author_id,
        storage_uuid: root.storage_uuid,
        replies,
    }))
}

#[cfg(test)]
pub(super) fn validate_plain_highlight_graph(
    package: &IWorkPackage,
    archive_name: &str,
    identifier: u64,
    object: &ArchiveObject,
) -> Result<Option<AnnotationGraph>> {
    let Some(graph) = validate_annotation_graph(package, archive_name, identifier, object)? else {
        return Ok(None);
    };
    if !graph.is_plain_highlight() {
        return Err(Error::InvalidFormat(format!(
            "iWork highlight {identifier} is comment-backed rather than a plain highlight"
        )));
    }
    Ok(Some(graph))
}

fn validate_comment_node(
    archive: &Archive,
    annotation_id: u64,
    storage_id: u64,
) -> Result<CommentNode> {
    let object = archive.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} references missing comment storage {storage_id}"
        ))
    })?;
    if object.messages.len() != 1 || object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE {
        return Err(Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} comment storage {storage_id} must contain exactly one comment payload"
        )));
    }
    let comment = tsd::CommentStorageArchive::decode(object.messages[0].data.as_slice())?;
    let body = comment.text.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} comment storage {storage_id} is missing text presence"
        ))
    })?;
    let creation_date_seconds = comment.creation_date.map(|date| date.seconds);
    if creation_date_seconds.is_some_and(|seconds| !seconds.is_finite()) {
        return Err(Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} has a non-finite creation date"
        )));
    }
    let uuid = comment.storage_uuid.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} comment storage {storage_id} is missing its UUID"
        ))
    })?;
    if uuid.lower == 0 && uuid.upper == 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} comment storage {storage_id} has a zero UUID"
        )));
    }
    let author_id = comment.author.map(|reference| reference.identifier);
    if author_id == Some(0) {
        return Err(Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} has a zero author identifier"
        )));
    }
    let reply_ids = comment
        .replies
        .into_iter()
        .map(|reference| reference.identifier)
        .collect::<Vec<_>>();
    if reply_ids.contains(&0) {
        return Err(Error::InvalidFormat(format!(
            "iWork text annotation {annotation_id} has a zero reply identifier"
        )));
    }
    Ok(CommentNode {
        body,
        creation_date_seconds,
        author_id,
        storage_uuid: Uuid::from_parts(uuid.lower, uuid.upper)?,
        reply_ids,
    })
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

pub(super) fn insert_annotation_comment_storage(
    package: &mut IWorkPackage,
    archive_name: &str,
    comment_storage_id: u64,
    body: String,
    author_id: Option<u64>,
) -> Result<()> {
    let storage_uuid = fresh_comment_storage_uuid(package)?;
    insert_comment_storage(
        package,
        archive_name,
        comment_storage_id,
        body,
        author_id,
        storage_uuid,
    )
}

pub(super) fn update_annotation_comment_text(
    package: &mut IWorkPackage,
    archive_name: &str,
    annotation_id: u64,
    storage_id: u64,
    body: &str,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text annotation {annotation_id} comment storage {storage_id} is missing"
            ))
        })?;
        if object.messages.len() != 1 || object.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE {
            return Err(Error::InvalidFormat(format!(
                "iWork text annotation {annotation_id} comment storage {storage_id} must contain exactly one comment payload"
            )));
        }
        let original = &object.messages[0];
        let comment = tsd::CommentStorageArchive::decode(original.data.as_slice())?;
        let data = patch_length_delimited_field(
            &original.data,
            1,
            comment.text.is_some(),
            Some(body.as_bytes()),
        )?;
        if tsd::CommentStorageArchive::decode(data.as_slice())?
            .text
            .as_deref()
            != Some(body)
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text annotation {annotation_id} comment update failed validation"
            )));
        }
        object.replace_message(
            0,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        Ok(())
    })
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
                    "text-annotation graph object {identifier} retains an indexed reference in {archive_name}"
                )));
            }
        }
    }
    Ok(())
}
