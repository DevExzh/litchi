//! Transactional direct-reply CRUD for native ranged text annotations.

use crate::comments::{
    ensure_annotation_author, fresh_comment_storage_uuid, insert_comment_storage,
    remove_generated_annotation_author_if_unused, update_comment_reply_reference,
};
use crate::package_metadata::{
    component_identifier_for_entry, next_object_identifier, release_package_identifier_suffix,
    set_package_last_object_identifier,
};
use crate::{Error, IWorkPackage, Result};

use super::annotation::{
    AnnotationKind, add_author_external_reference, annotation_by_id, remove_registered_object,
    roundtrip,
};
use super::highlight_object::{
    AnnotationReplyGraph, ensure_no_metadata_reference, require_exclusive_reference,
    update_annotation_comment_text,
};
use super::highlight_storage::locate_storage;

pub(super) fn annotation_replies(
    package: &IWorkPackage,
    storage_id: u64,
    annotation_id: u64,
) -> Result<Vec<AnnotationReplyGraph>> {
    Ok(
        annotation_by_id(package, storage_id, annotation_id, AnnotationKind::Comment)?
            .graph
            .replies,
    )
}

pub(super) fn add_annotation_reply(
    package: &mut IWorkPackage,
    storage_id: u64,
    annotation_id: u64,
    body: String,
) -> Result<AnnotationReplyGraph> {
    require_nonempty_body(&body)?;
    let annotation = annotation_by_id(package, storage_id, annotation_id, AnnotationKind::Comment)?;
    let location = locate_storage(package, storage_id)?;
    require_exclusive_reference(
        package,
        annotation_id,
        annotation.graph.comment_storage_id,
        "text-comment root storage",
    )?;

    let mut staged = package.clone();
    let (author_id, author_entry, _) = ensure_annotation_author(&mut staged)?;
    let reply_id = next_object_identifier(&staged)?;
    let reply_uuid = fresh_comment_storage_uuid(&staged)?;
    insert_comment_storage(
        &mut staged,
        &location.archive_name,
        reply_id,
        body,
        author_id,
        reply_uuid,
    )?;
    update_comment_reply_reference(
        &mut staged,
        annotation.graph.comment_storage_id,
        None,
        Some(reply_id),
    )?;
    add_author_external_reference(
        &mut staged,
        &location.archive_name,
        author_entry.as_deref(),
        author_id,
    )?;
    set_package_last_object_identifier(&mut staged, reply_id)?;

    let verified = roundtrip(&staged)?;
    let reply = reply_by_id(
        annotation_replies(&verified, storage_id, annotation_id)?,
        reply_id,
        annotation_id,
    )?;
    *package = staged;
    Ok(reply)
}

pub(super) fn update_annotation_reply(
    package: &mut IWorkPackage,
    storage_id: u64,
    annotation_id: u64,
    reply_id: u64,
    body: &str,
) -> Result<AnnotationReplyGraph> {
    require_nonempty_body(body)?;
    let annotation = annotation_by_id(package, storage_id, annotation_id, AnnotationKind::Comment)?;
    let root_storage_id = annotation.graph.comment_storage_id;
    let current = reply_by_id(annotation.graph.replies, reply_id, annotation_id)?;
    if current.body == body {
        return Ok(current);
    }
    require_exclusive_reference(
        package,
        annotation_id,
        root_storage_id,
        "text-comment root storage",
    )?;
    require_exclusive_reference(package, root_storage_id, reply_id, "text-comment reply")?;
    let location = locate_storage(package, storage_id)?;

    let mut staged = package.clone();
    update_annotation_comment_text(
        &mut staged,
        &location.archive_name,
        annotation_id,
        reply_id,
        body,
    )?;
    let verified = roundtrip(&staged)?;
    let updated = reply_by_id(
        annotation_replies(&verified, storage_id, annotation_id)?,
        reply_id,
        annotation_id,
    )?;
    if updated.body != body {
        return Err(Error::InvalidFormat(format!(
            "iWork text-comment reply {reply_id} update failed validation"
        )));
    }
    *package = staged;
    Ok(updated)
}

pub(super) fn remove_annotation_reply(
    package: &mut IWorkPackage,
    storage_id: u64,
    annotation_id: u64,
    reply_id: u64,
) -> Result<AnnotationReplyGraph> {
    let annotation = annotation_by_id(package, storage_id, annotation_id, AnnotationKind::Comment)?;
    let root_storage_id = annotation.graph.comment_storage_id;
    let removed = reply_by_id(annotation.graph.replies, reply_id, annotation_id)?;
    require_exclusive_reference(
        package,
        annotation_id,
        root_storage_id,
        "text-comment root storage",
    )?;
    require_exclusive_reference(package, root_storage_id, reply_id, "text-comment reply")?;
    let location = locate_storage(package, storage_id)?;

    let mut staged = package.clone();
    update_comment_reply_reference(&mut staged, root_storage_id, Some(reply_id), None)?;
    ensure_no_metadata_reference(&staged, reply_id)?;
    let owning_component = component_identifier_for_entry(&staged, &location.archive_name)?;
    remove_registered_object(&mut staged, owning_component, reply_id)?;
    staged.update_archive(&location.archive_name, |archive| {
        archive.remove_object(reply_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text-comment reply storage {reply_id} is missing"
            ))
        })?;
        Ok(())
    })?;

    let mut removed_ids = vec![reply_id];
    if let Some(author_id) = removed.author_id
        && remove_generated_annotation_author_if_unused(&mut staged, author_id)?
    {
        removed_ids.push(author_id);
    }
    removed_ids.sort_unstable_by(|left, right| right.cmp(left));
    release_package_identifier_suffix(&mut staged, &removed_ids)?;

    let verified = roundtrip(&staged)?;
    if annotation_replies(&verified, storage_id, annotation_id)?
        .iter()
        .any(|reply| reply.storage_id == reply_id)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork text-comment reply {reply_id} deletion failed validation"
        )));
    }
    *package = staged;
    Ok(removed)
}

fn reply_by_id(
    replies: Vec<AnnotationReplyGraph>,
    reply_id: u64,
    annotation_id: u64,
) -> Result<AnnotationReplyGraph> {
    let mut matches = replies
        .into_iter()
        .filter(|reply| reply.storage_id == reply_id);
    let Some(reply) = matches.next() else {
        return Err(Error::ParseError(format!(
            "text comment {annotation_id} does not own reply {reply_id}"
        )));
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "text comment {annotation_id} duplicates reply {reply_id}"
        )));
    }
    Ok(reply)
}

fn require_nonempty_body(body: &str) -> Result<()> {
    if body.is_empty() {
        Err(Error::ParseError(
            "iWork ranged text-comment replies require nonempty text".to_owned(),
        ))
    } else {
        Ok(())
    }
}
