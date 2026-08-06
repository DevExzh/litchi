//! Native ranged text-comment CRUD for TSWP text storages.

use crate::{IWorkPackage, Result};

use super::annotation::{
    AnnotationKind, AnnotationRecord, add_annotation, annotations, remove_annotation,
    update_annotation,
};
use super::annotation_reply::{
    add_annotation_reply, annotation_replies, remove_annotation_reply, update_annotation_reply,
};
use super::highlight_object::AnnotationReplyGraph;
use litchi_iwa_common::comment::AuthorId;
use litchi_iwa_text::comment::{
    Metadata, TextComment, TextCommentBody, TextCommentId, TextCommentReply, TextCommentReplyBody,
    TextCommentReplyId,
    raw::{comment_id, comment_id_value, reply_id, reply_id_value},
};
use litchi_iwa_text::date_time::Instant;
use litchi_iwa_text::position::TextRange;

/// Read every native ranged text comment in a storage, ordered by position.
pub(crate) fn text_comments(package: &IWorkPackage, storage_id: u64) -> Result<Vec<TextComment>> {
    annotations(package, storage_id)?
        .into_iter()
        .filter(|annotation| !annotation.graph.is_plain_highlight())
        .map(text_comment)
        .collect()
}

/// Create a native ranged comment over an unoccupied annotation range.
pub(crate) fn add_text_comment(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    body: TextCommentBody,
) -> Result<TextComment> {
    text_comment(add_annotation(
        package,
        storage_id,
        range,
        AnnotationKind::Comment,
        body.into_string(),
    )?)
}

/// Atomically update a comment body and range while retaining its identity.
pub(crate) fn update_text_comment(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextCommentId,
    range: TextRange,
    body: TextCommentBody,
) -> Result<TextComment> {
    text_comment(update_annotation(
        package,
        storage_id,
        comment_id_value(id),
        range,
        AnnotationKind::Comment,
        body.as_str(),
    )?)
}

/// Delete one ranged comment and its owned root/reply annotation graph.
pub(crate) fn remove_text_comment(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextCommentId,
) -> Result<TextComment> {
    text_comment(remove_annotation(
        package,
        storage_id,
        comment_id_value(id),
        AnnotationKind::Comment,
    )?)
}

/// Read every direct reply in stored order.
pub(crate) fn text_comment_replies(
    package: &IWorkPackage,
    storage_id: u64,
    comment_id: TextCommentId,
) -> Result<Vec<TextCommentReply>> {
    annotation_replies(package, storage_id, comment_id_value(comment_id))?
        .into_iter()
        .map(|reply| text_comment_reply(comment_id, reply))
        .collect()
}

/// Append one direct reply to a ranged comment.
pub(crate) fn add_text_comment_reply(
    package: &mut IWorkPackage,
    storage_id: u64,
    comment_id: TextCommentId,
    body: TextCommentReplyBody,
) -> Result<TextCommentReply> {
    text_comment_reply(
        comment_id,
        add_annotation_reply(
            package,
            storage_id,
            comment_id_value(comment_id),
            body.into_string(),
        )?,
    )
}

/// Update a direct reply while retaining its identity and metadata.
pub(crate) fn update_text_comment_reply(
    package: &mut IWorkPackage,
    storage_id: u64,
    comment_id: TextCommentId,
    reply_id: TextCommentReplyId,
    body: TextCommentReplyBody,
) -> Result<TextCommentReply> {
    text_comment_reply(
        comment_id,
        update_annotation_reply(
            package,
            storage_id,
            comment_id_value(comment_id),
            reply_id_value(reply_id),
            body.as_str(),
        )?,
    )
}

/// Delete one direct reply and its owned comment storage.
pub(crate) fn remove_text_comment_reply(
    package: &mut IWorkPackage,
    storage_id: u64,
    comment_id: TextCommentId,
    reply_id: TextCommentReplyId,
) -> Result<TextCommentReply> {
    text_comment_reply(
        comment_id,
        remove_annotation_reply(
            package,
            storage_id,
            comment_id_value(comment_id),
            reply_id_value(reply_id),
        )?,
    )
}

fn text_comment(annotation: AnnotationRecord) -> Result<TextComment> {
    let reply_count = u32::try_from(annotation.graph.replies.len()).map_err(|_| {
        crate::Error::InvalidFormat("iWork text comment reply count exceeds u32".to_owned())
    })?;
    Ok(TextComment::new(
        comment_id(annotation.object_id)?,
        annotation.range,
        TextCommentBody::from_boxed(annotation.graph.body.into_boxed_str())?,
        Metadata::new(
            annotation
                .graph
                .creation_date_seconds
                .map(Instant::from_reference_date_seconds)
                .transpose()?,
            annotation
                .graph
                .author_id
                .map(AuthorId::from_raw)
                .transpose()?,
            annotation.graph.storage_uuid,
        ),
        reply_count,
    ))
}

fn text_comment_reply(
    comment_id: TextCommentId,
    reply: AnnotationReplyGraph,
) -> Result<TextCommentReply> {
    Ok(TextCommentReply::new(
        reply_id(reply.storage_id)?,
        comment_id,
        TextCommentReplyBody::from_boxed(reply.body.into_boxed_str())?,
        Metadata::new(
            reply
                .creation_date_seconds
                .map(Instant::from_reference_date_seconds)
                .transpose()?,
            reply.author_id.map(AuthorId::from_raw).transpose()?,
            reply.storage_uuid,
        ),
    ))
}

#[cfg(test)]
#[path = "text_comment_internal_tests.rs"]
mod tests;
