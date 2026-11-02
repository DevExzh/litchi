//! Direct-reply transitions of the shared native comment engine.
//!
//! The host's reply writer is copy-on-write: every reply operation clones the
//! selected root comment storage under a fresh identifier, edits the clone's
//! reply references, moves the cell's comment-list entry to the clone and
//! removes the source root (with the reply a replacement or removal gives up).
//! An appended reply is a new comment storage beside the root, written by the
//! root writer; a replaced reply is a clone of the source reply with its text
//! patched.  This module reproduces that shape over the shared structural
//! working set, so Pages and Keynote share one reply implementation.
//!
//! The admitted thread is the shape the host writes: a uniquely owned root
//! whose direct replies are leaves owned only by that root.  A nested reply,
//! a shared reply or a multi-payload storage is refused as a dependency
//! before anything is staged.

use super::{
    AuthorPlan, COMMENT_STORAGE_MESSAGE_TYPE, Fault, FaultResult, LimitKind,
    MAX_COMMENT_TEXT_BYTES, Mutation, Profile, Selection, TABLE_REFERENCE_IDENTIFIER_FIELD,
    append_list, author_storage_component, comment_codec, comment_storage_message_index,
    create_root, ensure_unique_root, map_comment_decode_error, map_core_error, map_wire_error,
    materialize_author, plan_author, read_varint, register_insertions, remove_list_entry,
    rewrite_existing_text, rewrite_tile, unique_field, validate_archive_header_shape,
    validate_existing_comment_graph, validate_selection, varint_field,
};
use crate::structure::{self, TableLeaf, Working, metadata};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{self, WireView};
use litchi_iwa_core::ArchiveObject;
use std::sync::Arc;

/// The comment-storage field holding the direct reply references.
const REPLIES_FIELD: u32 = 4;

/// One direct-reply operation, addressed by its source ordinal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ReplyOperation {
    /// Append one direct reply after the source replies.
    Append(Box<str>),
    /// Replace the text of the reply at `index`.
    Replace {
        /// Source-ordered reply ordinal.
        index: usize,
        /// The replacement text.
        text: Box<str>,
    },
    /// Remove the reply at `index`.
    Remove {
        /// Source-ordered reply ordinal.
        index: usize,
    },
}

/// Execute one bounded copy-on-write reply transition through the shared
/// structural working set.
pub fn rewrite_reply<P: TableLeaf>(
    source: &P,
    selection: &Selection,
    operation: &ReplyOperation,
    profile: Profile,
) -> FaultResult<Mutation<P>> {
    validate_selection(source, selection, profile)?;
    let mut working = Working::new(source)?;
    let census_admission = structure::collect::Census::admit(&working)?;
    let census = structure::collect::Census::from_admission(&working, census_admission)?;
    let author_storage_component = author_storage_component(&working)?;
    let registry = metadata::registry(&working)?.ok_or(Fault::Unsupported)?;
    let root = selection.root_identifier.ok_or(Fault::Invalid)?;
    let list = selection.list_identifier.ok_or(Fault::Invalid)?;
    let key = selection.comment_key.ok_or(Fault::Invalid)?;
    ensure_unique_root(&working, &census, selection, root)?;
    validate_existing_comment_graph(
        &working,
        root,
        selection.model_component.as_ref(),
        author_storage_component.as_deref(),
    )?;
    let root_component = working.locate(root)?.ok_or(Fault::Invalid)?;
    let replies = direct_replies(&working, &census, root)?;
    let selected = match operation {
        ReplyOperation::Append(_) => None,
        ReplyOperation::Replace { index, .. } | ReplyOperation::Remove { index } => {
            Some(*replies.get(*index).ok_or(Fault::Invalid)?)
        },
    };
    if let ReplyOperation::Append(text) | ReplyOperation::Replace { text, .. } = operation {
        if text.len() > MAX_COMMENT_TEXT_BYTES {
            return Err(Fault::Limit {
                kind: LimitKind::TextBytes,
                observed: text.len() as u64,
                maximum: MAX_COMMENT_TEXT_BYTES as u64,
            });
        }
    }

    let mut author: Option<AuthorPlan> = match operation {
        ReplyOperation::Append(_) => Some(plan_author(
            &working,
            selection,
            profile.generate_author_if_missing,
            profile.update_author_storage,
        )?),
        _ => None,
    };
    let writes_reply = !matches!(operation, ReplyOperation::Remove { .. });
    let allocation_count = 1_usize
        .saturating_add(usize::from(writes_reply))
        .saturating_add(
            author
                .as_ref()
                .map_or(0, |plan| usize::from(plan.new_author)),
        )
        .saturating_add(
            author
                .as_ref()
                .map_or(0, |plan| usize::from(plan.new_storage)),
        );
    let mut allocated = working.allocate(allocation_count)?.into_iter();
    let new_root = allocated.next().ok_or(Fault::Invalid)?;
    let new_reply = if writes_reply {
        Some(allocated.next().ok_or(Fault::Invalid)?)
    } else {
        None
    };
    if let Some(plan) = author.as_mut() {
        if plan.new_author {
            plan.author_identifier = allocated.next();
        }
        if plan.new_storage {
            plan.storage_identifier = allocated.next();
        }
    }

    let mut transition = metadata::Transition::default();
    let mut inserted = Vec::new();
    structure::reserve(&mut inserted, 4)?;
    let mut removed = Vec::new();
    structure::reserve(&mut removed, 2)?;

    clone_storage(
        &mut working,
        &root_component,
        root,
        new_root,
        &registry,
        &mut transition,
    )?;
    inserted.push(new_root);
    match operation {
        ReplyOperation::Append(text) => {
            let reply = new_reply.ok_or(Fault::Invalid)?;
            let plan = author.as_ref().ok_or(Fault::Invalid)?;
            create_root(
                &mut working,
                &root_component,
                reply,
                text,
                plan.author_identifier,
                &registry,
                &mut transition,
            )?;
            inserted.push(reply);
            materialize_author(
                &mut working,
                plan,
                profile,
                &registry,
                &mut transition,
                &mut inserted,
            )?;
        },
        ReplyOperation::Replace { text, .. } => {
            let reply = new_reply.ok_or(Fault::Invalid)?;
            let source_reply = selected.ok_or(Fault::Invalid)?;
            let reply_component = working.locate(source_reply)?.ok_or(Fault::Invalid)?;
            clone_storage(
                &mut working,
                &reply_component,
                source_reply,
                reply,
                &registry,
                &mut transition,
            )?;
            rewrite_existing_text(&mut working, reply, text)?;
            inserted.push(reply);
        },
        ReplyOperation::Remove { .. } => {},
    }
    retarget_replies(&mut working, new_root, &replies, selected, new_reply)?;

    remove_list_entry(&mut working, list, key, root)?;
    let new_key = append_list(&mut working, list, new_root, profile.list_keys)?;
    rewrite_tile(&mut working, selection, Some(key), Some(new_key))?;

    working.remove(root)?;
    removed.push(root);
    if let Some(source_reply) = selected {
        working.remove(source_reply)?;
        removed.push(source_reply);
    }
    transition.remove_objects(&registry, &removed)?;
    transition.release_suffix(&removed)?;

    let mut stamped = vec![
        selection.tile_component.to_string(),
        root_component,
        working.locate(list)?.ok_or(Fault::Invalid)?,
    ];
    if author.as_ref().is_some_and(|plan| plan.new_author) {
        if let Some(component) = author_storage_component {
            stamped.push(component);
        }
    }
    for component in stamped {
        let component_id = registry.component_for_entry(&component)?.identifier;
        transition.stamp(component_id)?;
    }
    register_insertions(&working, &registry, &inserted, &mut transition)?;
    if let Some(token) = working.save_token()? {
        working.set_save_token(token.checked_add(1).ok_or(Fault::Limit {
            kind: LimitKind::References,
            observed: u64::MAX,
            maximum: u64::MAX,
        })?);
    }
    if profile.keep_previews {
        working.keep_previews();
    }
    let published = working.publish(&transition)?;
    let target: Arc<[u8]> = Arc::from(published.package.source_bytes());
    Ok(Mutation {
        package: published.package,
        target,
        touched_components: published.touched_components,
        deleted_previews: published.deleted_previews,
    })
}

/// The source-ordered direct replies of `root`, each admitted as a leaf
/// storage owned only by `root`.
pub(super) fn direct_replies<P: TableLeaf>(
    working: &Working<'_, P>,
    census: &structure::collect::Census,
    root: u64,
) -> FaultResult<Vec<u64>> {
    let replies = storage_replies(working, root)?;
    for (index, reply) in replies.iter().enumerate() {
        if replies[..index].contains(reply) || *reply == root {
            return Err(Fault::Invalid);
        }
        let referrers = census.referrers(*reply);
        if referrers.is_empty() || referrers.iter().any(|referrer| *referrer != root) {
            return Err(Fault::Dependency);
        }
        let object = working.object(*reply)?;
        if object.messages.len() != 1 {
            return Err(Fault::Dependency);
        }
        validate_archive_header_shape(object)?;
        if !storage_replies(working, *reply)?.is_empty() {
            return Err(Fault::Dependency);
        }
    }
    Ok(replies)
}

/// Decode the reply references of one comment storage.
fn storage_replies<P: TableLeaf>(working: &Working<'_, P>, storage: u64) -> FaultResult<Vec<u64>> {
    let object = working.object(storage)?;
    let index = comment_storage_message_index(object)?;
    let data = object
        .messages
        .get(index)
        .ok_or(Fault::Invalid)?
        .data
        .as_slice();
    let maximum = working.source().max_references();
    let options = comment_codec::DecodeOptions::new(
        data.len().max(1),
        data.len().saturating_mul(4).max(1),
        data.len().saturating_mul(64).max(1),
        64,
        maximum,
        MAX_COMMENT_TEXT_BYTES,
    );
    let mut replies = Vec::new();
    let mut fault = None;
    comment_codec::visit_comment_storage_replies(data, options, &mut |reply| {
        if fault.is_some() {
            return Ok(());
        }
        if replies.len() >= maximum {
            fault = Some(Fault::Limit {
                kind: LimitKind::Replies,
                observed: replies.len().saturating_add(1) as u64,
                maximum: maximum as u64,
            });
        } else if replies.try_reserve(1).is_err() {
            fault = Some(Fault::Allocation {
                amount: replies.len().saturating_add(1),
            });
        } else {
            replies.push(reply.identifier());
        }
        Ok(())
    })
    .map_err(map_comment_decode_error)?;
    if let Some(fault) = fault {
        return Err(fault);
    }
    Ok(replies)
}

/// Clone one single-payload comment storage exactly (payload and header)
/// under `identifier`, as the host's `clone_comment_storage_exact` does,
/// and register the clone's object UUID.
fn clone_storage<P: TableLeaf>(
    working: &mut Working<'_, P>,
    component: &str,
    source: u64,
    identifier: u64,
    registry: &metadata::Registry,
    transition: &mut metadata::Transition,
) -> FaultResult<()> {
    let object = working.object(source)?;
    if object.messages.len() != 1
        || object
            .messages
            .first()
            .is_none_or(|message| message.type_ != COMMENT_STORAGE_MESSAGE_TYPE)
    {
        return Err(Fault::Dependency);
    }
    validate_archive_header_shape(object)?;
    let limits = working.archive_limits();
    let mut clone =
        ArchiveObject::new(identifier, object.messages.clone(), limits).map_err(map_core_error)?;
    clone.archive_info.should_merge = object.archive_info.should_merge;
    clone
        .archive_info
        .message_infos
        .clone_from(&object.archive_info.message_infos);
    working.insert(component, clone)?;
    let seed = metadata::identity_seed(working, identifier)?;
    let component_id = registry.component_for_entry(component)?.identifier;
    transition.register_uuid(registry, component_id, identifier, seed)?;
    Ok(())
}

/// Rewrite the clone's reply references: `old -> new` (replace),
/// `None -> new` (append) or `old -> None` (remove), then mirror the host's
/// header update (aggregate references follow; field-level mirrors only
/// drop the old reply).
fn retarget_replies<P: TableLeaf>(
    working: &mut Working<'_, P>,
    clone: u64,
    before: &[u64],
    old: Option<u64>,
    new: Option<u64>,
) -> FaultResult<()> {
    let component = working.locate(clone)?.ok_or(Fault::Invalid)?;
    let index = comment_storage_message_index(working.object(clone)?)?;
    let source = structure::copy_bytes(
        working
            .object(clone)?
            .messages
            .get(index)
            .ok_or(Fault::Invalid)?
            .data
            .as_slice(),
    )?;
    let output = match (old, new) {
        (None, Some(new)) => {
            let reference = varint_field(TABLE_REFERENCE_IDENTIFIER_FIELD, new)?;
            let mut output = source;
            wire::append_length_delimited_field(
                &mut output,
                REPLIES_FIELD,
                &reference,
                WireLimits::DEFAULT,
            )
            .map_err(map_wire_error)?;
            output
        },
        (Some(old), Some(new)) => {
            wire::transform_repeated_length_delimited_fields(&source, REPLIES_FIELD, |payload| {
                if reference_identifier(payload).map_err(ReplyWireError)? == old {
                    wire::patch_varint_field(
                        payload,
                        TABLE_REFERENCE_IDENTIFIER_FIELD,
                        true,
                        Some(new),
                    )
                    .map_err(ReplyWireError::from)
                } else {
                    structure::copy_bytes(payload)
                        .map_err(|fault| ReplyWireError(Fault::from(fault)))
                }
            })
            .map_err(|error: ReplyWireError| error.0)?
        },
        (Some(old), None) => {
            wire::remove_repeated_length_delimited_field_where(&source, REPLIES_FIELD, |payload| {
                Ok(reference_identifier(payload).map_err(ReplyWireError)? == old)
            })
            .map_err(|error: ReplyWireError| error.0)?
        },
        (None, None) => return Err(Fault::Invalid),
    };
    let limits = working.archive_limits();
    {
        let object = working
            .archive_mut(&component)?
            .object_mut(clone)
            .ok_or(Fault::Invalid)?;
        object
            .replace_message_preserving_header(
                index,
                litchi_iwa_core::RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data: output,
                },
                limits,
            )
            .map_err(map_core_error)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Fault::Invalid)?;
        if let Some(old) = old {
            info.object_references.retain(|reference| *reference != old);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|reference| *reference != old);
            }
        }
        if let Some(new) = new {
            if !info.object_references.contains(&new) {
                info.object_references.push(new);
            }
        }
    }
    let expected: Vec<u64> = match (old, new) {
        (None, Some(new)) => before.iter().copied().chain([new]).collect(),
        (Some(old), new) => before
            .iter()
            .filter_map(|reply| if *reply == old { new } else { Some(*reply) })
            .collect(),
        (None, None) => return Err(Fault::Invalid),
    };
    if storage_replies(working, clone)? != expected {
        return Err(Fault::Verification);
    }
    Ok(())
}

fn reference_identifier(payload: &[u8]) -> FaultResult<u64> {
    let view = WireView::parse(payload, WireLimits::DEFAULT).map_err(map_wire_error)?;
    let field = unique_field(&view, TABLE_REFERENCE_IDENTIFIER_FIELD)?.ok_or(Fault::Invalid)?;
    if field.wire_type() != 0 {
        return Err(Fault::Invalid);
    }
    read_varint(field.payload())
}

/// A wire-callback error carrying the engine's fault.
struct ReplyWireError(Fault);

impl From<litchi_iwa_common::Error> for ReplyWireError {
    fn from(error: litchi_iwa_common::Error) -> Self {
        Self(map_wire_error(error))
    }
}
