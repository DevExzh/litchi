//! Native ranged text-comment CRUD for TSWP text storages.

use crate::{IWorkPackage, Result};

use super::annotation::{
    AnnotationKind, AnnotationRecord, add_annotation, annotations, remove_annotation,
    update_annotation,
};
use super::position::TextRange;
use super::text_comment_types::{TextComment, TextCommentBody, TextCommentId};

/// Read every native ranged text comment in a storage, ordered by position.
pub(crate) fn text_comments(package: &IWorkPackage, storage_id: u64) -> Result<Vec<TextComment>> {
    Ok(annotations(package, storage_id)?
        .into_iter()
        .filter(|annotation| !annotation.graph.is_plain_highlight())
        .map(text_comment)
        .collect())
}

/// Create a native ranged comment over an unoccupied annotation range.
pub(crate) fn add_text_comment(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    body: TextCommentBody,
) -> Result<TextComment> {
    Ok(text_comment(add_annotation(
        package,
        storage_id,
        range,
        AnnotationKind::Comment,
        body.into_string(),
    )?))
}

/// Atomically update a comment body and range while retaining its identity.
pub(crate) fn update_text_comment(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextCommentId,
    range: TextRange,
    body: TextCommentBody,
) -> Result<TextComment> {
    Ok(text_comment(update_annotation(
        package,
        storage_id,
        id.object_id(),
        range,
        AnnotationKind::Comment,
        body.as_str(),
    )?))
}

/// Delete one ranged comment and its owned root/reply annotation graph.
pub(crate) fn remove_text_comment(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextCommentId,
) -> Result<TextComment> {
    Ok(text_comment(remove_annotation(
        package,
        storage_id,
        id.object_id(),
        AnnotationKind::Comment,
    )?))
}

fn text_comment(annotation: AnnotationRecord) -> TextComment {
    TextComment {
        id: TextCommentId::from_native(annotation.object_id),
        range: annotation.range,
        body: TextCommentBody::from_native(annotation.graph.body),
        creation_date_seconds: annotation.graph.creation_date_seconds,
        author_object_id: annotation.graph.author_id,
        storage_uuid: annotation.graph.storage_uuid,
        reply_count: annotation.graph.replies.len(),
    }
}

#[cfg(test)]
#[path = "text_comment_internal_tests.rs"]
mod tests;
