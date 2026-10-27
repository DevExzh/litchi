//! Highlights, text comments and replies of a text box or shape text
//! (record 0372; the Keynote leaf of the host's
//! `*_slide_text_box_{highlight,comment,comment_reply}*` families over the
//! shared W1-D annotation engine, `litchi_iwa_text_wire::annotation`).
//!
//! A highlight and a comment are both a `TSWP.HighlightArchive` covering a
//! range of the storage's highlight table and owning a
//! `TSD.CommentStorageArchive` (empty for a highlight); replies are further
//! comment storages the root lists. New storages name the document's
//! generated annotation author, created on first use as the host does.
//! Annotations are addressed by position (highlights and comments
//! separately, in text order), replies by position within their comment.

use litchi_iwa_common::comment::{AuthorId, Uuid};
use litchi_iwa_core::{ArchiveObject, FieldInfo, FieldPath, RawMessage, UnknownFieldRule};
use litchi_iwa_protos::annotation_author_codec;
use litchi_iwa_text::date_time::Instant;
use litchi_iwa_text::position::TextRange;
use litchi_iwa_text_wire::annotation::{
    self, Annotation, COMMENT_STORAGE_MESSAGE_TYPE, HIGHLIGHT_MESSAGE_TYPE, Objects, ReplyChange,
};
use litchi_iwa_text_wire::ranged::{self, Change, Table};

use super::super::structure::wire::reference;
use super::super::structure::workspace::Workspace;
use super::text_hyperlinks::attribute_uuid;
use super::text_lists::{add_edge, component_of, remove_edges_to};
use super::text_properties::storage_payload;
use super::{Error, Refusal, Unsupported, lift};

const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const AUTHOR_MESSAGE_TYPE: u32 = 212;
const AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const AUTHOR_NAME: &str = "litchi-iwa";
const AUTHOR_PUBLIC_ID: &str = "4C495443-4849-4957-8100-000000000001:058e44481db1c6fdeeac88af010136d7f8949f54bde61ef9af3e078562b968b6";

/// One staged annotation change of one storage.
#[derive(Clone)]
pub(super) enum CommentChange {
    AddHighlight(TextRange, Instant),
    UpdateHighlight(usize, TextRange),
    RemoveHighlight(usize),
    AddComment(TextRange, String, Instant),
    UpdateComment(usize, TextRange, String),
    RemoveComment(usize),
    AddReply(usize, String, Instant),
    UpdateReply(usize, usize, String),
    RemoveReply(usize, usize),
}

fn refused() -> Error {
    Error::Refused(Refusal::Unsupported {
        reason: Unsupported::TextStyling,
    })
}

fn from_engine<E>(_error: E) -> Error {
    refused()
}

struct WorkspaceObjects<'w, 'p>(&'w Workspace<'p>);

impl Objects for WorkspaceObjects<'_, '_> {
    fn messages(&self, identifier: u64) -> Option<Vec<(u32, &[u8])>> {
        self.0.object(identifier).map(|object| {
            object
                .messages
                .iter()
                .map(|message| (message.type_, message.data.as_slice()))
                .collect()
        })
    }
}

/// Every source object holding `message_type`, staged.
fn messages(workspace: &Workspace<'_>, message_type: u32) -> Vec<(u64, usize, Vec<u8>)> {
    let mut found = Vec::new();
    for component in workspace.package().state.source.components().iter() {
        for source in &component.archive().objects {
            let Some(identifier) = source.archive_info.identifier else {
                continue;
            };
            let Some(object) = workspace.object(identifier) else {
                continue;
            };
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == message_type {
                    found.push((identifier, index, message.data.clone()));
                }
            }
        }
    }
    found
}

/// The generated author's canonical payload (the host's
/// `generated_annotation_author(true)`).
fn generated_author() -> Result<Vec<u8>, Error> {
    let color = annotation_author_codec::AuthorColorWrite::new(
        1,
        Some(0.368_627_46),
        Some(0.568_627_5),
        Some(0.937_254_9),
        Some(1.0),
        None,
        None,
        None,
        None,
        None,
        Some(1),
    );
    let public_ids = [AUTHOR_PUBLIC_ID];
    let write = annotation_author_codec::AnnotationAuthorWrite::new(
        Some(AUTHOR_NAME),
        Some(color),
        Some(AUTHOR_PUBLIC_ID),
        Some(false),
        &public_ids,
    );
    annotation_author_codec::encode_annotation_author_with_report(
        &write,
        annotation_author_codec::EncodeOptions::for_author(&write),
    )
    .map(|(bytes, _)| bytes)
    .map_err(|_| Error::InvalidSource)
}

fn author_storage(workspace: &Workspace<'_>) -> Result<Option<(u64, usize, Vec<u8>)>, Error> {
    let mut found = messages(workspace, AUTHOR_STORAGE_MESSAGE_TYPE);
    if found.len() > 1 {
        return Err(refused());
    }
    Ok(found.pop())
}

fn author_ids(payload: &[u8], wire: litchi_iwa_common::WireLimits) -> Result<Vec<u64>, Error> {
    let fields = super::super::structure::wire::Fields::parse(payload, wire).map_err(lift)?;
    fields
        .all(1)
        .map(|field| {
            super::super::structure::wire::reference_identifier(
                field.message().map_err(lift)?,
                wire,
            )
            .map_err(lift)
        })
        .collect()
}

fn is_generated(workspace: &Workspace<'_>, author: u64) -> Result<bool, Error> {
    let generated = generated_author()?;
    Ok(workspace.object(author).is_some_and(|object| {
        object
            .messages
            .iter()
            .any(|message| message.type_ == AUTHOR_MESSAGE_TYPE && message.data == generated)
    }))
}

/// The host's `ensure_annotation_author`: the generated public author,
/// appended to the author storage when missing.
fn ensure_author(workspace: &mut Workspace<'_>) -> Result<Option<u64>, Error> {
    let wire = workspace.wire;
    let Some((storage, index, payload)) = author_storage(workspace)? else {
        return Ok(None);
    };
    for author in author_ids(&payload, wire)? {
        if is_generated(workspace, author)? {
            return Ok(Some(author));
        }
    }
    let author = workspace.allocate().map_err(lift)?;
    let data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
        &payload,
        1,
        &reference(author),
    )
    .map_err(|_| Error::InvalidSource)?;
    let limits = workspace.archive_limits();
    let mut object = ArchiveObject::new(
        author,
        vec![RawMessage {
            type_: AUTHOR_MESSAGE_TYPE,
            data: generated_author()?,
        }],
        limits,
    )
    .map_err(|_| Error::InvalidSource)?;
    let mut infos = Vec::new();
    for field in [4, 3] {
        let mut info = FieldInfo::new(
            FieldPath::new(vec![field], limits).map_err(|_| Error::InvalidSource)?,
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
        info.unknown_field_rule = Some(UnknownFieldRule::IgnoreAndPreserve);
        infos.push(info);
    }
    object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?
        .field_infos = infos;
    workspace
        .object_mut(storage)
        .map_err(lift)?
        .replace_message(
            index,
            RawMessage {
                type_: AUTHOR_STORAGE_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let member = workspace
        .component_of(storage)
        .ok_or(Error::InvalidSource)?;
    workspace.insert(&member, object).map_err(lift)?;
    let component = component_of(workspace, storage)?;
    let uuid = workspace.clone_uuid(None, author).map_err(lift)?;
    workspace
        .registry
        .component_mut(component)
        .ok_or(Error::InvalidSource)?
        .add_object_uuid(author, uuid);
    Ok(Some(author))
}

/// The host's `remove_generated_annotation_author_if_unused`.
fn remove_author_if_unused(workspace: &mut Workspace<'_>, author: u64) -> Result<(), Error> {
    let wire = workspace.wire;
    if !is_generated(workspace, author)? {
        return Ok(());
    }
    for (_, _, payload) in messages(workspace, COMMENT_STORAGE_MESSAGE_TYPE) {
        if annotation::comment_storage(&payload, wire)
            .ok()
            .and_then(|storage| storage.metadata().author())
            .map(AuthorId::get)
            == Some(author)
        {
            return Ok(());
        }
    }
    let (storage, index, payload) = author_storage(workspace)?.ok_or(Error::InvalidSource)?;
    let component = component_of(workspace, storage)?;
    remove_edges_to(workspace, component, author, None)?;
    workspace
        .registry
        .component_mut(component)
        .ok_or(Error::InvalidSource)?
        .remove_object_uuids(&[author]);
    let ids = author_ids(&payload, wire)?;
    let kept = ids
        .iter()
        .filter(|id| **id != author)
        .map(|id| reference(*id))
        .collect::<Vec<_>>();
    let mut fields = super::super::structure::wire::Fields::parse(&payload, wire).map_err(lift)?;
    fields.replace_all(1, super::super::structure::wire::LENGTH_DELIMITED, &kept);
    let limits = workspace.archive_limits();
    workspace
        .object_mut(storage)
        .map_err(lift)?
        .replace_message(
            index,
            RawMessage {
                type_: AUTHOR_STORAGE_MESSAGE_TYPE,
                data: fields.encode(),
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    workspace.remove_object(author).map_err(lift)?;
    Ok(())
}

/// A deterministic storage UUID no comment storage uses yet.
fn fresh_uuid(workspace: &Workspace<'_>, seed: u64) -> Result<Uuid, Error> {
    let wire = workspace.wire;
    let used = messages(workspace, COMMENT_STORAGE_MESSAGE_TYPE)
        .into_iter()
        .filter_map(|(_, _, payload)| annotation::comment_storage(&payload, wire).ok())
        .map(|storage| storage.metadata().storage_uuid())
        .collect::<Vec<_>>();
    let mut lower = 0x4c49_5443_4b4e_0000 | (seed & 0xffff);
    loop {
        let uuid =
            Uuid::from_parts(lower, 0x6c69_7463_6869_6b6e).map_err(|_| Error::InvalidSource)?;
        if !used.contains(&uuid) {
            return Ok(uuid);
        }
        lower = lower.wrapping_add(1);
    }
}

fn insert_comment_storage(
    workspace: &mut Workspace<'_>,
    member: &str,
    identifier: u64,
    body: &str,
    created: Instant,
    author: Option<u64>,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let uuid = fresh_uuid(workspace, identifier)?;
    let payload = annotation::encode_comment_storage(
        body,
        created,
        author
            .map(AuthorId::from_raw)
            .transpose()
            .map_err(|_| Error::InvalidSource)?,
        uuid,
        wire,
    )
    .map_err(from_engine)?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
        workspace.archive_limits(),
    )
    .map_err(|_| Error::InvalidSource)?;
    if let Some(author) = author {
        object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or(Error::InvalidSource)?
            .object_references
            .push(author);
    }
    workspace.insert(member, object).map_err(lift)?;
    let component = workspace
        .registry
        .component_for_member(member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    let uuid = workspace.clone_uuid(None, identifier).map_err(lift)?;
    workspace
        .registry
        .component_mut(component)
        .ok_or(Error::InvalidSource)?
        .add_object_uuid(identifier, uuid);
    Ok(())
}

fn add_author_edge(
    workspace: &mut Workspace<'_>,
    storage: u64,
    author: Option<u64>,
) -> Result<(), Error> {
    let Some(author) = author else {
        return Ok(());
    };
    let (source, target) = (
        component_of(workspace, storage)?,
        component_of(workspace, author)?,
    );
    add_edge(workspace, source, target, author)
}

/// Remove the edges and UUID registration of a removed object (the host's
/// `remove_registered_object`).
fn remove_registered(workspace: &mut Workspace<'_>, owner: u64, object: u64) -> Result<(), Error> {
    remove_edges_to(workspace, owner, object, None)?;
    let registered = workspace
        .registry
        .components()
        .find(|candidate| candidate.uuid_of(object).is_some())
        .map(|candidate| candidate.identifier);
    if let Some(registered) = registered {
        workspace
            .registry
            .component_mut(registered)
            .ok_or(Error::InvalidSource)?
            .remove_object_uuids(&[object]);
    }
    Ok(())
}

fn replace_message(
    workspace: &mut Workspace<'_>,
    object: u64,
    message_type: u32,
    payload: Vec<u8>,
    added: Option<u64>,
    removed: Option<u64>,
) -> Result<(), Error> {
    let limits = workspace.archive_limits();
    let target = workspace.object_mut(object).map_err(lift)?;
    let index = target
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    target
        .replace_message(
            index,
            RawMessage {
                type_: message_type,
                data: payload,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    let info = target
        .archive_info
        .message_infos
        .get_mut(index)
        .ok_or(Error::InvalidSource)?;
    if let Some(added) = added {
        if info.object_references.contains(&added) {
            return Err(Error::InvalidSource);
        }
        info.object_references.push(added);
    }
    if let Some(removed) = removed {
        info.object_references
            .retain(|reference| *reference != removed);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|reference| *reference != removed);
        }
    }
    Ok(())
}

/// The storage's annotations of one kind, in text order.
pub(super) fn annotations(
    workspace: &Workspace<'_>,
    storage: u64,
    highlights: bool,
) -> Result<Vec<Annotation>, Error> {
    let (_, payload) = storage_payload(workspace, storage)?;
    Ok(
        annotation::annotations(&payload, &WorkspaceObjects(workspace), workspace.wire)
            .map_err(from_engine)?
            .into_iter()
            .filter(|item| item.is_highlight() == highlights)
            .collect(),
    )
}

fn select(
    workspace: &Workspace<'_>,
    storage: u64,
    highlights: bool,
    position: usize,
) -> Result<Annotation, Error> {
    annotations(workspace, storage, highlights)?
        .into_iter()
        .nth(position)
        .ok_or(refused())
}

fn add(
    workspace: &mut Workspace<'_>,
    storage: u64,
    range: TextRange,
    body: &str,
    created: Instant,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let author = ensure_author(workspace)?;
    let object = workspace.allocate().map_err(lift)?;
    let comment_storage = workspace.allocate().map_err(lift)?;
    let (_, payload) = storage_payload(workspace, storage)?;
    let rewrite = ranged::rewrite(
        &payload,
        Table::Highlight,
        &Change::Add { object, range },
        wire,
    )
    .map_err(from_engine)?;
    let (added, removed) = (rewrite.added(), rewrite.removed());
    replace_message(
        workspace,
        storage,
        STORAGE_MESSAGE_TYPE,
        rewrite.into_payload(),
        added,
        removed,
    )?;
    let uuid = attribute_uuid(workspace, object)?;
    let mut highlight = ArchiveObject::new(
        object,
        vec![RawMessage {
            type_: HIGHLIGHT_MESSAGE_TYPE,
            data: annotation::encode_highlight(comment_storage, &uuid, wire)
                .map_err(from_engine)?,
        }],
        workspace.archive_limits(),
    )
    .map_err(|_| Error::InvalidSource)?;
    highlight
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?
        .object_references
        .push(comment_storage);
    let member = workspace
        .component_of(storage)
        .ok_or(Error::InvalidSource)?;
    workspace.insert(&member, highlight).map_err(lift)?;
    insert_comment_storage(workspace, &member, comment_storage, body, created, author)?;
    add_author_edge(workspace, storage, author)
}

fn update(
    workspace: &mut Workspace<'_>,
    storage: u64,
    current: &Annotation,
    range: TextRange,
    body: &str,
) -> Result<(), Error> {
    let wire = workspace.wire;
    if current.range() != range {
        let (_, payload) = storage_payload(workspace, storage)?;
        let rewrite = ranged::rewrite(
            &payload,
            Table::Highlight,
            &Change::Move {
                object: current.object(),
                range,
            },
            wire,
        )
        .map_err(from_engine)?;
        replace_message(
            workspace,
            storage,
            STORAGE_MESSAGE_TYPE,
            rewrite.into_payload(),
            None,
            None,
        )?;
    }
    if current.root().body() != body {
        let root = current.root().storage();
        let payload = workspace
            .payload(root, COMMENT_STORAGE_MESSAGE_TYPE)
            .map_err(lift)?
            .to_vec();
        let payload =
            annotation::comment_storage_with_body(&payload, body, wire).map_err(from_engine)?;
        replace_message(
            workspace,
            root,
            COMMENT_STORAGE_MESSAGE_TYPE,
            payload,
            None,
            None,
        )?;
    }
    Ok(())
}

fn remove(workspace: &mut Workspace<'_>, storage: u64, current: &Annotation) -> Result<(), Error> {
    let wire = workspace.wire;
    let (_, payload) = storage_payload(workspace, storage)?;
    let rewrite = ranged::rewrite(
        &payload,
        Table::Highlight,
        &Change::Remove {
            object: current.object(),
        },
        wire,
    )
    .map_err(from_engine)?;
    let (added, removed) = (rewrite.added(), rewrite.removed());
    replace_message(
        workspace,
        storage,
        STORAGE_MESSAGE_TYPE,
        rewrite.into_payload(),
        added,
        removed,
    )?;
    let owner = component_of(workspace, storage)?;
    let mut objects = vec![current.object(), current.root().storage()];
    objects.extend(current.replies().iter().map(annotation::Node::storage));
    for identifier in objects {
        remove_registered(workspace, owner, identifier)?;
        workspace.remove_object(identifier).map_err(lift)?;
    }
    let mut authors = current.authors().collect::<Vec<_>>();
    authors.sort_unstable();
    authors.dedup();
    for author in authors {
        remove_author_if_unused(workspace, author)?;
    }
    workspace.release_removed_identifier_suffix();
    Ok(())
}

/// Publish one annotation change of `storage`.
pub(super) fn apply(
    workspace: &mut Workspace<'_>,
    storage: u64,
    change: &CommentChange,
) -> Result<(), Error> {
    let wire = workspace.wire;
    match change {
        CommentChange::AddHighlight(range, created) => {
            add(workspace, storage, *range, "", *created)
        },
        CommentChange::UpdateHighlight(position, range) => {
            let current = select(workspace, storage, true, *position)?;
            update(workspace, storage, &current, *range, "")
        },
        CommentChange::RemoveHighlight(position) => {
            let current = select(workspace, storage, true, *position)?;
            remove(workspace, storage, &current)
        },
        CommentChange::AddComment(range, body, created) => {
            if body.is_empty() {
                return Err(refused());
            }
            add(workspace, storage, *range, body, *created)
        },
        CommentChange::UpdateComment(position, range, body) => {
            if body.is_empty() {
                return Err(refused());
            }
            let current = select(workspace, storage, false, *position)?;
            update(workspace, storage, &current, *range, body)
        },
        CommentChange::RemoveComment(position) => {
            let current = select(workspace, storage, false, *position)?;
            remove(workspace, storage, &current)
        },
        CommentChange::AddReply(position, body, created) => {
            let current = select(workspace, storage, false, *position)?;
            let author = ensure_author(workspace)?;
            let reply = workspace.allocate().map_err(lift)?;
            let member = workspace
                .component_of(storage)
                .ok_or(Error::InvalidSource)?;
            insert_comment_storage(workspace, &member, reply, body, *created, author)?;
            let root = current.root().storage();
            let payload = workspace
                .payload(root, COMMENT_STORAGE_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let payload =
                annotation::comment_storage_with_reply(&payload, ReplyChange::Add(reply), wire)
                    .map_err(from_engine)?;
            replace_message(
                workspace,
                root,
                COMMENT_STORAGE_MESSAGE_TYPE,
                payload,
                Some(reply),
                None,
            )?;
            add_author_edge(workspace, storage, author)
        },
        CommentChange::UpdateReply(position, reply, body) => {
            let current = select(workspace, storage, false, *position)?;
            let node = current.replies().get(*reply).ok_or(refused())?;
            if node.body() == body {
                return Ok(());
            }
            let storage_id = node.storage();
            let payload = workspace
                .payload(storage_id, COMMENT_STORAGE_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let payload =
                annotation::comment_storage_with_body(&payload, body, wire).map_err(from_engine)?;
            replace_message(
                workspace,
                storage_id,
                COMMENT_STORAGE_MESSAGE_TYPE,
                payload,
                None,
                None,
            )
        },
        CommentChange::RemoveReply(position, reply) => {
            let current = select(workspace, storage, false, *position)?;
            let node = current.replies().get(*reply).ok_or(refused())?.clone();
            let root = current.root().storage();
            let payload = workspace
                .payload(root, COMMENT_STORAGE_MESSAGE_TYPE)
                .map_err(lift)?
                .to_vec();
            let payload = annotation::comment_storage_with_reply(
                &payload,
                ReplyChange::Remove(node.storage()),
                wire,
            )
            .map_err(from_engine)?;
            replace_message(
                workspace,
                root,
                COMMENT_STORAGE_MESSAGE_TYPE,
                payload,
                None,
                Some(node.storage()),
            )?;
            let owner = component_of(workspace, storage)?;
            remove_registered(workspace, owner, node.storage())?;
            workspace.remove_object(node.storage()).map_err(lift)?;
            if let Some(author) = node.metadata().author() {
                remove_author_if_unused(workspace, author.get())?;
            }
            workspace.release_removed_identifier_suffix();
            Ok(())
        },
    }
}
