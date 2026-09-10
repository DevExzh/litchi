//! OPC graph ownership and atomic publication for legacy comment extensions.

use std::sync::Arc;

use litchi_opc::{OpcError, OpcPackage, PackURI};

use super::transaction::{
    AuthorPresenceCommit, AuthorPresencePatch, AuthorPresenceSnapshot, CommentThreadingCommit,
    CommentThreadingPatch, CommentThreadingSnapshot,
};
use super::{PresenceInfo, ThreadingInfo};
use crate::comments::{
    AUTHORS_REL, COMMENTS_REL, Comments, STRICT_AUTHORS_REL, STRICT_COMMENTS_REL,
    load_presentation_comments,
};
use crate::{Error, Result};

/// Load an inert author presence value, if the selected author has one.
pub fn load_presence(package: &OpcPackage, author_id: u32) -> Result<Option<PresenceInfo>> {
    Ok(load_presence_snapshot(package, author_id)?.and_then(|snapshot| snapshot.value().cloned()))
}

/// Capture the exact author-part source for one author, including an absent
/// typed extension when the author itself exists.
pub fn load_presence_snapshot(
    package: &OpcPackage,
    author_id: u32,
) -> Result<Option<AuthorPresenceSnapshot>> {
    let Some(graph) = validate_graph(package)? else {
        return Ok(None);
    };
    if !graph.authors.iter().any(|author| author.id == author_id) {
        return Ok(None);
    }
    let uri = PackURI::new(&graph.author_part_name).map_err(invalid)?;
    let part = package.get_part(&uri)?;
    AuthorPresenceSnapshot::from_part(graph.author_part_name, part.blob_arc(), author_id).map(Some)
}

/// Load an inert threading value, if the selected legacy comment has one.
pub fn load_threading(
    package: &OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
) -> Result<Option<ThreadingInfo>> {
    Ok(
        load_threading_snapshot(package, slide_part_name, author_id, index)?
            .and_then(|snapshot| snapshot.value().cloned()),
    )
}

/// Capture the exact comment-part source for one comment, including an absent
/// typed extension when the comment itself exists.
pub fn load_threading_snapshot(
    package: &OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
) -> Result<Option<CommentThreadingSnapshot>> {
    let Some(graph) = validate_graph(package)? else {
        return Ok(None);
    };
    let Some(slide) = graph
        .slides
        .iter()
        .find(|slide| slide.slide_part_name == slide_part_name)
    else {
        return Ok(None);
    };
    if !slide
        .comments
        .iter()
        .any(|comment| comment.author_id == author_id && comment.index == index)
    {
        return Ok(None);
    }
    let uri = PackURI::new(&slide.part_name).map_err(invalid)?;
    let part = package.get_part(&uri)?;
    CommentThreadingSnapshot::from_part(
        slide.part_name.clone(),
        part.blob_arc(),
        slide.slide_part_name.clone(),
        author_id,
        index,
    )
    .map(Some)
}

/// Apply an author-presence patch after validating its exact part source and
/// the complete legacy comment graph.
pub fn apply_presence_patch(
    package: &mut OpcPackage,
    patch: &AuthorPresencePatch,
) -> Result<AuthorPresenceSnapshot> {
    let current = load_presence_snapshot(package, patch.before().author_id())?
        .ok_or_else(|| invalid("author presence source part is absent"))?;
    ensure_same_author_source(&current, patch.before())?;
    if patch.is_empty() {
        return Ok(current);
    }
    ensure_signature_policy(package)?;
    let uri = PackURI::new(patch.before().source_part_name()).map_err(Error::Uri)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&uri)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_presence_snapshot(&candidate, patch.after().author_id())?
        .ok_or_else(|| invalid("published author presence source disappeared"))?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published author presence source differs from patch",
        ));
    }
    *package = candidate;
    Ok(result)
}

/// Apply a committed author-presence edit atomically.
pub fn apply_presence_commit(
    package: &mut OpcPackage,
    commit: AuthorPresenceCommit,
) -> Result<AuthorPresenceSnapshot> {
    apply_presence_patch(package, commit.patch())
}

/// Apply a comment-threading patch after validating its exact part source and
/// the complete legacy comment graph.
pub fn apply_threading_patch(
    package: &mut OpcPackage,
    patch: &CommentThreadingPatch,
) -> Result<CommentThreadingSnapshot> {
    let current = load_threading_snapshot(
        package,
        patch.before().slide_part_name(),
        patch.before().author_id(),
        patch.before().index(),
    )?
    .ok_or_else(|| invalid("comment threading source part is absent"))?;
    ensure_same_threading_source(&current, patch.before())?;
    if patch.is_empty() {
        return Ok(current);
    }
    ensure_signature_policy(package)?;
    let uri = PackURI::new(patch.before().source_part_name()).map_err(Error::Uri)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&uri)?
        .set_blob_shared(Arc::clone(patch.after().source_arc()));
    let result = load_threading_snapshot(
        &candidate,
        patch.after().slide_part_name(),
        patch.after().author_id(),
        patch.after().index(),
    )?
    .ok_or_else(|| invalid("published comment threading source disappeared"))?;
    if !result.same_source(patch.after()) {
        return Err(invalid(
            "published comment threading source differs from patch",
        ));
    }
    *package = candidate;
    Ok(result)
}

/// Apply a committed comment-threading edit atomically.
pub fn apply_threading_commit(
    package: &mut OpcPackage,
    commit: CommentThreadingCommit,
) -> Result<CommentThreadingSnapshot> {
    apply_threading_patch(package, commit.patch())
}

/// Replace or create the selected author presence extension.
pub fn put_presence(
    package: &mut OpcPackage,
    author_id: u32,
    value: PresenceInfo,
) -> Result<Option<PresenceInfo>> {
    let snapshot = load_presence_snapshot(package, author_id)?
        .ok_or_else(|| invalid("comment author does not exist"))?;
    let previous = snapshot.value().cloned();
    let mut edit = snapshot.edit();
    edit.set_presence(value)?;
    let commit = edit.commit()?;
    apply_presence_commit(package, commit)?;
    Ok(previous)
}

/// Remove the selected author presence extension.
pub fn remove_presence(package: &mut OpcPackage, author_id: u32) -> Result<Option<PresenceInfo>> {
    let Some(snapshot) = load_presence_snapshot(package, author_id)? else {
        return Ok(None);
    };
    let previous = snapshot.value().cloned();
    if previous.is_none() {
        return Ok(None);
    }
    let mut edit = snapshot.edit();
    edit.remove()?;
    apply_presence_commit(package, edit.commit()?)?;
    Ok(previous)
}

/// Replace or create the selected comment threading extension.
pub fn put_threading(
    package: &mut OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
    value: ThreadingInfo,
) -> Result<Option<ThreadingInfo>> {
    let snapshot = load_threading_snapshot(package, slide_part_name, author_id, index)?
        .ok_or_else(|| invalid("legacy comment does not exist"))?;
    let previous = snapshot.value().cloned();
    let mut edit = snapshot.edit();
    edit.set_threading(value)?;
    let commit = edit.commit()?;
    apply_threading_commit(package, commit)?;
    Ok(previous)
}

/// Remove the selected comment threading extension.
pub fn remove_threading(
    package: &mut OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
) -> Result<Option<ThreadingInfo>> {
    let Some(snapshot) = load_threading_snapshot(package, slide_part_name, author_id, index)?
    else {
        return Ok(None);
    };
    let previous = snapshot.value().cloned();
    if previous.is_none() {
        return Ok(None);
    }
    let mut edit = snapshot.edit();
    edit.remove()?;
    apply_threading_commit(package, edit.commit()?)?;
    Ok(previous)
}

fn validate_graph(package: &OpcPackage) -> Result<Option<Comments>> {
    validate_comment_relationship_targets(package)?;
    load_presentation_comments(package)
}

fn validate_comment_relationship_targets(package: &OpcPackage) -> Result<()> {
    if package.rels().iter().any(|relationship| {
        matches!(
            relationship.reltype(),
            AUTHORS_REL | STRICT_AUTHORS_REL | COMMENTS_REL | STRICT_COMMENTS_REL
        )
    }) {
        return Err(invalid(
            "legacy comment relationships cannot originate at the package root",
        ));
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                AUTHORS_REL | STRICT_AUTHORS_REL | COMMENTS_REL | STRICT_COMMENTS_REL
            )
        }) {
            if relationship.is_external()
                || relationship.target_query().is_some()
                || relationship.target_fragment().is_some()
            {
                return Err(invalid(format!(
                    "legacy comment relationship '{}' must target an exact internal part URI",
                    relationship.r_id()
                )));
            }
            relationship.target_partname()?;
        }
    }
    Ok(())
}

fn ensure_same_author_source(
    current: &AuthorPresenceSnapshot,
    expected: &AuthorPresenceSnapshot,
) -> Result<()> {
    if !current.same_source(expected) {
        return Err(invalid("author presence source is stale"));
    }
    Ok(())
}

fn ensure_same_threading_source(
    current: &CommentThreadingSnapshot,
    expected: &CommentThreadingSnapshot,
) -> Result<()> {
    if !current.same_source(expected) {
        return Err(invalid("comment threading source is stale"));
    }
    Ok(())
}

fn ensure_signature_policy(package: &OpcPackage) -> Result<()> {
    if package.is_signed() || package.requires_signature_edit_policy() {
        return Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
