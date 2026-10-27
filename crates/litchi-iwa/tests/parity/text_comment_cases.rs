//! Highlight and text-comment parity (wave W1-D, record 0374): the host
//! `text::Draft` highlight, comment and reply families against
//! `litchi_iwa_text_wire::annotation` and `ranged` published by
//! [`text_glue`](crate::text_glue).
//!
//! The host stamps new highlights with a random text-attribute UUID and new
//! comment storages with a random storage UUID and the current time, so
//! every case that creates one is tier B with those fields masked; moves,
//! body updates and removals are tier A.

use std::collections::BTreeMap;

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_archive::iwa::{ArchiveObject, FieldInfo, FieldPath, RawMessage, UnknownFieldRule};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::comment::{AuthorId, Uuid};
use litchi_iwa_protos::{tsd, tsk, tsp};
use litchi_iwa_text::comment::raw::{comment_id, comment_id_value, reply_id, reply_id_value};
use litchi_iwa_text::comment::{TextCommentBody, TextCommentReplyBody};
use litchi_iwa_text::date_time::Instant;
use litchi_iwa_text::highlight::raw::{from_object_id, object_id};
use litchi_iwa_text::position::TextRange;
use litchi_iwa_text_wire::annotation::{
    self, Annotation, COMMENT_STORAGE_MESSAGE_TYPE, HIGHLIGHT_MESSAGE_TYPE, Objects, ReplyChange,
};
use litchi_iwa_text_wire::ranged::{self, Change, Table};
use litchi_iwa_text_wire::smart_field::TextAttributeUuid;
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::{Decoded, KnownFields, decode_as, top_level_fields};
use crate::text_annotation_cases::text_storages;
use crate::text_glue::{Located, Workspace};
use crate::text_style_cases::FIXTURES;

const LIMITS: WireLimits = WireLimits::DEFAULT;
const LEAF_UUID: &str = "4c495443-4849-5731-442d-686c69676874";
const AUTHOR_MESSAGE_TYPE: u32 = 212;
const AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const AUTHOR_NAME: &str = "litchi-iwa";
const AUTHOR_PUBLIC_ID: &str = "4C495443-4849-4957-8100-000000000001:058e44481db1c6fdeeac88af010136d7f8949f54bde61ef9af3e078562b968b6";
const APPLE_EPOCH_UNIX_OFFSET_SECONDS: f64 = 978_307_200.0;

#[derive(Clone, Debug)]
pub enum Operation {
    AddHighlight(TextRange),
    UpdateHighlight(u64, TextRange),
    RemoveHighlight(u64),
    AddComment(TextRange, &'static str),
    UpdateComment(u64, TextRange, &'static str),
    RemoveComment(u64),
    AddReply(u64, &'static str),
    UpdateReply(u64, u64, &'static str),
    RemoveReply(u64, u64),
}

struct ObjectMap(BTreeMap<u64, Vec<(u32, Vec<u8>)>>);

impl ObjectMap {
    fn load(workspace: &Workspace) -> TestResult<Self> {
        let mut objects = BTreeMap::new();
        for (_, archive) in workspace.archives()? {
            for object in &archive.objects {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                objects.insert(
                    identifier,
                    object
                        .messages
                        .iter()
                        .map(|message| (message.type_, message.data.clone()))
                        .collect(),
                );
            }
        }
        Ok(Self(objects))
    }
}

impl Objects for ObjectMap {
    fn messages(&self, identifier: u64) -> Option<Vec<(u32, &[u8])>> {
        self.0.get(&identifier).map(|messages| {
            messages
                .iter()
                .map(|(message_type, payload)| (*message_type, payload.as_slice()))
                .collect()
        })
    }
}

fn generated_author(public: bool) -> tsk::AnnotationAuthorArchive {
    tsk::AnnotationAuthorArchive {
        name: Some(AUTHOR_NAME.to_owned()),
        color: Some(tsp::Color {
            model: tsp::color::ColorModel::Rgb as i32,
            r: Some(0.368_627_46),
            g: Some(0.568_627_5),
            b: Some(0.937_254_9),
            rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
            a: Some(1.0),
            ..Default::default()
        }),
        public_id: public.then(|| AUTHOR_PUBLIC_ID.to_owned()),
        is_public_author: Some(false),
        public_ids: if public {
            vec![AUTHOR_PUBLIC_ID.to_owned()]
        } else {
            Vec::new()
        },
    }
}

fn is_generated(workspace: &Workspace, author: u64, public: bool) -> TestResult<bool> {
    let located = workspace.message(author, AUTHOR_MESSAGE_TYPE)?;
    Ok(
        tsk::AnnotationAuthorArchive::decode(located.payload.as_slice())?
            == generated_author(public),
    )
}

fn author_storage(workspace: &Workspace) -> TestResult<Option<Located>> {
    let mut found = workspace.messages(AUTHOR_STORAGE_MESSAGE_TYPE)?;
    if found.len() > 1 {
        return Err("several annotation-author storages".into());
    }
    Ok(found.pop())
}

fn author_ids(storage: &Located) -> TestResult<Vec<u64>> {
    Ok(
        tsk::AnnotationAuthorStorageArchive::decode(storage.payload.as_slice())?
            .annotation_author
            .iter()
            .map(|reference| reference.identifier)
            .collect(),
    )
}

/// The host's `ensure_annotation_author` (public generated author).
fn ensure_author(workspace: &mut Workspace) -> TestResult<Option<(u64, String)>> {
    let Some(storage) = author_storage(workspace)? else {
        return Ok(None);
    };
    for author in author_ids(&storage)? {
        if is_generated(workspace, author, true)? {
            return Ok(Some((author, storage.entry)));
        }
    }
    let author = workspace.next_identifier()?;
    let reference = tsp::Reference {
        identifier: author,
        ..Default::default()
    }
    .encode_to_vec();
    let data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
        &storage.payload,
        1,
        &reference,
    )?;
    let mut object = ArchiveObject::new(
        author,
        vec![RawMessage {
            type_: AUTHOR_MESSAGE_TYPE,
            data: generated_author(true).encode_to_vec(),
        }],
        Default::default(),
    )?;
    object.archive_info.message_infos.as_mut_slice()[0].field_infos = [4, 3]
        .into_iter()
        .map(|field| FieldInfo {
            path: FieldPath { path: vec![field] },
            unknown_field_rule: Some(UnknownFieldRule::IgnoreAndPreserve),
            ..Default::default()
        })
        .collect();
    workspace.update(&storage.entry, |archive| {
        let storage_object = archive
            .object_mut(storage.object)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("author storage".to_owned()))?;
        storage_object.replace_message(
            storage.index,
            RawMessage {
                type_: AUTHOR_STORAGE_MESSAGE_TYPE,
                data,
            },
            Default::default(),
        )?;
        archive.insert_object(object, Default::default())?;
        Ok(())
    })?;
    add_uuid_for_entry(workspace, &storage.entry, author)?;
    Ok(Some((author, storage.entry)))
}

fn add_uuid_for_entry(workspace: &mut Workspace, entry: &str, object: u64) -> TestResult {
    if workspace.metadata()?.is_none() {
        return Ok(());
    }
    let component = workspace
        .component(entry)?
        .ok_or("no component for the archive")?;
    workspace.add_object_uuid(component, object)
}

/// The host's `remove_generated_annotation_author_if_unused`.
fn remove_author_if_unused(
    workspace: &mut Workspace,
    owner_entry: &str,
    author: u64,
) -> TestResult<bool> {
    if !is_generated(workspace, author, true)? && !is_generated(workspace, author, false)? {
        return Ok(false);
    }
    for located in workspace.messages(COMMENT_STORAGE_MESSAGE_TYPE)? {
        let decoded = tsd::CommentStorageArchive::decode(located.payload.as_slice())?;
        if decoded
            .author
            .is_some_and(|reference| reference.identifier == author)
        {
            return Ok(false);
        }
    }
    let storage = author_storage(workspace)?.ok_or("no author storage")?;
    if let Some(component) = workspace.component(&storage.entry)? {
        workspace.remove_external_references_to(component, author)?;
    }
    let _ = owner_entry;
    if workspace.metadata()?.is_some() {
        let component = workspace
            .component(&storage.entry)?
            .ok_or("no component for the author archive")?;
        if workspace.uuid_component(author)? == Some(component) {
            workspace.remove_object_uuid(component, author)?;
        }
    }
    let data = litchi_iwa_common::wire::remove_repeated_length_delimited_field_where(
        &storage.payload,
        1,
        |payload| {
            Ok::<bool, litchi_iwa_common::Error>(
                tsp::Reference::decode(payload)
                    .is_ok_and(|reference| reference.identifier == author),
            )
        },
    )?;
    workspace.update(&storage.entry, |archive| {
        let storage_object = archive
            .object_mut(storage.object)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("author storage".to_owned()))?;
        storage_object.replace_message(
            storage.index,
            RawMessage {
                type_: AUTHOR_STORAGE_MESSAGE_TYPE,
                data,
            },
            Default::default(),
        )?;
        archive
            .remove_object(author)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("author".to_owned()))?;
        Ok(())
    })?;
    Ok(true)
}

fn add_author_reference(
    workspace: &mut Workspace,
    source_entry: &str,
    author: Option<&(u64, String)>,
) -> TestResult {
    let Some((author, author_entry)) = author else {
        return Ok(());
    };
    if let (Some(source), Some(target)) = (
        workspace.component(source_entry)?,
        workspace.component(author_entry)?,
    ) && source != target
    {
        workspace.add_external_reference(source, target, *author)?;
    }
    Ok(())
}

/// The host's `remove_registered_object`.
fn remove_registered(workspace: &mut Workspace, owner: Option<u64>, object: u64) -> TestResult {
    if let Some(component) = owner {
        workspace.remove_external_references_to(component, object)?;
    }
    if let Some(component) = workspace.uuid_component(object)? {
        workspace.remove_object_uuid(component, object)?;
    }
    Ok(())
}

fn now() -> TestResult<Instant> {
    let unix = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)?
        .as_secs_f64();
    Ok(Instant::from_reference_date_seconds(
        unix - APPLE_EPOCH_UNIX_OFFSET_SECONDS,
    )?)
}

/// A deterministic storage UUID no comment storage uses yet.
fn fresh_uuid(workspace: &Workspace, seed: u64) -> TestResult<Uuid> {
    let mut used = Vec::new();
    for located in workspace.messages(COMMENT_STORAGE_MESSAGE_TYPE)? {
        if let Ok(decoded) = annotation::comment_storage(&located.payload, LIMITS) {
            used.push(decoded.metadata().storage_uuid());
        }
    }
    let mut lower = 0x4c49_5443_4849_0000 | seed;
    loop {
        let uuid = Uuid::from_parts(lower, 0x6c69_7463_6869_7577)?;
        if !used.contains(&uuid) {
            return Ok(uuid);
        }
        lower += 1;
    }
}

fn insert_comment_storage(
    workspace: &mut Workspace,
    entry: &str,
    identifier: u64,
    body: &str,
    author: Option<&(u64, String)>,
) -> TestResult {
    let uuid = fresh_uuid(workspace, identifier)?;
    let payload = annotation::encode_comment_storage(
        body,
        now()?,
        author.map(|(id, _)| AuthorId::from_raw(*id)).transpose()?,
        uuid,
        LIMITS,
    )?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
        Default::default(),
    )?;
    if let Some((author, _)) = author {
        object.archive_info.message_infos.as_mut_slice()[0]
            .object_references
            .push(*author);
    }
    workspace.update(entry, |archive| {
        archive.insert_object(object, Default::default())?;
        Ok(())
    })?;
    add_uuid_for_entry(workspace, entry, identifier)
}

fn annotation_of(
    workspace: &Workspace,
    storage: &Located,
    object: u64,
    highlight: bool,
) -> TestResult<Annotation> {
    Ok(annotation::annotation(
        &storage.payload,
        &ObjectMap::load(workspace)?,
        object,
        highlight,
        LIMITS,
    )?)
}

fn add(workspace: &mut Workspace, storage: u64, range: TextRange, body: &str) -> TestResult {
    let author = ensure_author(workspace)?;
    let object = workspace.next_identifier()?;
    let comment_storage = object + 1;
    let located = workspace.storage(storage)?;
    let rewrite = ranged::rewrite(
        &located.payload,
        Table::Highlight,
        &Change::Add { object, range },
        LIMITS,
    )?;
    let (added, removed) = (rewrite.added(), rewrite.removed());
    workspace.replace_with_references(&located, rewrite.into_payload(), added, removed)?;
    let mut highlight = ArchiveObject::new(
        object,
        vec![RawMessage {
            type_: HIGHLIGHT_MESSAGE_TYPE,
            data: annotation::encode_highlight(
                comment_storage,
                &TextAttributeUuid::new(LEAF_UUID)?,
                LIMITS,
            )?,
        }],
        Default::default(),
    )?;
    highlight.archive_info.message_infos.as_mut_slice()[0]
        .object_references
        .push(comment_storage);
    workspace.update(&located.entry, |archive| {
        archive.insert_object(highlight, Default::default())?;
        Ok(())
    })?;
    insert_comment_storage(
        workspace,
        &located.entry,
        comment_storage,
        body,
        author.as_ref(),
    )?;
    add_author_reference(workspace, &located.entry, author.as_ref())?;
    workspace.raise_watermark(comment_storage)
}

fn update(
    workspace: &mut Workspace,
    storage: u64,
    object: u64,
    range: TextRange,
    body: &str,
) -> TestResult {
    let located = workspace.storage(storage)?;
    let current = annotation_of(workspace, &located, object, body.is_empty())?;
    if current.range() != range {
        let rewrite = ranged::rewrite(
            &located.payload,
            Table::Highlight,
            &Change::Move { object, range },
            LIMITS,
        )?;
        workspace.replace(&located, rewrite.into_payload(), None)?;
    }
    if current.root().body() != body {
        let root = workspace.message(current.root().storage(), COMMENT_STORAGE_MESSAGE_TYPE)?;
        let payload = annotation::comment_storage_with_body(&root.payload, body, LIMITS)?;
        workspace.replace(&root, payload, None)?;
    }
    Ok(())
}

fn remove(workspace: &mut Workspace, storage: u64, object: u64, highlight: bool) -> TestResult {
    let located = workspace.storage(storage)?;
    let current = annotation_of(workspace, &located, object, highlight)?;
    let rewrite = ranged::rewrite(
        &located.payload,
        Table::Highlight,
        &Change::Remove { object },
        LIMITS,
    )?;
    let (added, removed) = (rewrite.added(), rewrite.removed());
    workspace.replace_with_references(&located, rewrite.into_payload(), added, removed)?;
    let owner = workspace.component(&located.entry)?;
    let mut removed_ids = Vec::new();
    let mut objects = vec![object, current.root().storage()];
    objects.extend(current.replies().iter().map(annotation::Node::storage));
    for identifier in objects {
        remove_registered(workspace, owner, identifier)?;
        workspace.remove_object(&located.entry, identifier)?;
        removed_ids.push(identifier);
    }
    let mut authors: Vec<u64> = current.authors().collect();
    authors.sort_unstable();
    authors.dedup();
    for author in authors {
        if remove_author_if_unused(workspace, &located.entry, author)? {
            removed_ids.push(author);
        }
    }
    workspace.release_watermark(&removed_ids)
}

fn add_reply(workspace: &mut Workspace, storage: u64, comment: u64, body: &str) -> TestResult {
    let located = workspace.storage(storage)?;
    let current = annotation_of(workspace, &located, comment, false)?;
    let author = ensure_author(workspace)?;
    let reply = workspace.next_identifier()?;
    insert_comment_storage(workspace, &located.entry, reply, body, author.as_ref())?;
    let root = workspace.message(current.root().storage(), COMMENT_STORAGE_MESSAGE_TYPE)?;
    let payload =
        annotation::comment_storage_with_reply(&root.payload, ReplyChange::Add(reply), LIMITS)?;
    workspace.replace_with_references(&root, payload, Some(reply), None)?;
    add_author_reference(workspace, &located.entry, author.as_ref())?;
    workspace.raise_watermark(reply)
}

fn update_reply(
    workspace: &mut Workspace,
    storage: u64,
    comment: u64,
    reply: u64,
    body: &str,
) -> TestResult {
    let located = workspace.storage(storage)?;
    let current = annotation_of(workspace, &located, comment, false)?;
    if current.reply(reply)?.body() == body {
        return Ok(());
    }
    let node = workspace.message(reply, COMMENT_STORAGE_MESSAGE_TYPE)?;
    let payload = annotation::comment_storage_with_body(&node.payload, body, LIMITS)?;
    workspace.replace(&node, payload, None)
}

fn remove_reply(workspace: &mut Workspace, storage: u64, comment: u64, reply: u64) -> TestResult {
    let located = workspace.storage(storage)?;
    let current = annotation_of(workspace, &located, comment, false)?;
    let author = current.reply(reply)?.metadata().author().map(AuthorId::get);
    let root = workspace.message(current.root().storage(), COMMENT_STORAGE_MESSAGE_TYPE)?;
    let payload =
        annotation::comment_storage_with_reply(&root.payload, ReplyChange::Remove(reply), LIMITS)?;
    workspace.replace_with_references(&root, payload, None, Some(reply))?;
    let owner = workspace.component(&located.entry)?;
    remove_registered(workspace, owner, reply)?;
    workspace.remove_object(&located.entry, reply)?;
    let mut removed = vec![reply];
    if let Some(author) = author
        && remove_author_if_unused(workspace, &located.entry, author)?
    {
        removed.push(author);
    }
    workspace.release_watermark(&removed)
}

pub fn leaf_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    let ws = &mut workspace;
    match operation {
        Operation::AddHighlight(range) => add(ws, storage, *range, "")?,
        Operation::UpdateHighlight(id, range) => update(ws, storage, *id, *range, "")?,
        Operation::RemoveHighlight(id) => remove(ws, storage, *id, true)?,
        Operation::AddComment(range, body) => add(ws, storage, *range, body)?,
        Operation::UpdateComment(id, range, body) => update(ws, storage, *id, *range, body)?,
        Operation::RemoveComment(id) => remove(ws, storage, *id, false)?,
        Operation::AddReply(comment, body) => add_reply(ws, storage, *comment, body)?,
        Operation::UpdateReply(comment, reply, body) => {
            update_reply(ws, storage, *comment, *reply, body)?;
        },
        Operation::RemoveReply(comment, reply) => remove_reply(ws, storage, *comment, *reply)?,
    }
    workspace.into_bytes()
}

pub fn host_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match operation {
        Operation::AddHighlight(range) => {
            draft.add_text_highlight(id, *range)?;
        },
        Operation::UpdateHighlight(highlight, range) => {
            draft.update_text_highlight(id, from_object_id(*highlight)?, *range)?;
        },
        Operation::RemoveHighlight(highlight) => {
            draft.remove_text_highlight(id, from_object_id(*highlight)?)?;
        },
        Operation::AddComment(range, body) => {
            draft.add_text_comment(id, *range, TextCommentBody::new(body)?)?;
        },
        Operation::UpdateComment(comment, range, body) => {
            draft.update_text_comment(
                id,
                comment_id(*comment)?,
                *range,
                TextCommentBody::new(body)?,
            )?;
        },
        Operation::RemoveComment(comment) => {
            draft.remove_text_comment(id, comment_id(*comment)?)?;
        },
        Operation::AddReply(comment, body) => {
            draft.add_text_comment_reply(
                id,
                comment_id(*comment)?,
                TextCommentReplyBody::new(body)?,
            )?;
        },
        Operation::UpdateReply(comment, reply, body) => {
            draft.update_text_comment_reply(
                id,
                comment_id(*comment)?,
                reply_id(*reply)?,
                TextCommentReplyBody::new(body)?,
            )?;
        },
        Operation::RemoveReply(comment, reply) => {
            draft.remove_text_comment_reply(id, comment_id(*comment)?, reply_id(*reply)?)?;
        },
    }
    Ok(draft.to_bytes()?)
}

pub type Reads = Vec<String>;

pub fn host_reads(bytes: &[u8], storage: u64) -> TestResult<Reads> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    let mut reads = Vec::new();
    for highlight in draft.text_highlights(id)? {
        reads.push(format!(
            "highlight {} {:?}",
            object_id(highlight.id),
            highlight.range
        ));
    }
    for comment in draft.text_comments(id)? {
        reads.push(format!(
            "comment {} {:?} {:?} {:?} {}",
            comment_id_value(comment.id()),
            comment.range(),
            comment.body(),
            comment.author(),
            comment.reply_count()
        ));
        for reply in draft.text_comment_replies(id, comment.id())? {
            reads.push(format!(
                "reply {} {:?} {:?}",
                reply_id_value(reply.id()),
                reply.body(),
                reply.author()
            ));
        }
    }
    Ok(reads)
}

pub fn leaf_reads(bytes: &[u8], storage: u64) -> TestResult<Reads> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let found = annotation::annotations(&located.payload, &ObjectMap::load(&workspace)?, LIMITS)?;
    let mut reads = Vec::new();
    for item in found.iter().filter(|item| item.is_highlight()) {
        let highlight = item.highlight()?;
        reads.push(format!(
            "highlight {} {:?}",
            object_id(highlight.id),
            highlight.range
        ));
    }
    for item in found.iter().filter(|item| !item.is_highlight()) {
        let comment = item.comment()?;
        reads.push(format!(
            "comment {} {:?} {:?} {:?} {}",
            comment_id_value(comment.id()),
            comment.range(),
            comment.body(),
            comment.author(),
            comment.reply_count()
        ));
        for reply in item.comment_replies()? {
            reads.push(format!(
                "reply {} {:?} {:?}",
                reply_id_value(reply.id()),
                reply.body(),
                reply.author()
            ));
        }
    }
    Ok(reads)
}

fn masked_highlight(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = litchi_iwa_protos::tswp::HighlightArchive::decode(data)?;
    if message
        .text_attribute_uuid_string
        .as_deref()
        .is_some_and(|uuid| TextAttributeUuid::new(uuid).is_ok())
    {
        message.text_attribute_uuid_string = Some("<random uuid>".to_owned());
    }
    let mut decoded =
        decode_as::<litchi_iwa_protos::tswp::HighlightArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

fn masked_comment_storage(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut message = tsd::CommentStorageArchive::decode(data)?;
    if message.creation_date.is_some() {
        message.creation_date = Some(tsp::Date { seconds: 0.0 });
    }
    if message.storage_uuid.is_some() {
        message.storage_uuid = Some(tsp::Uuid { lower: 1, upper: 1 });
    }
    let mut decoded = decode_as::<tsd::CommentStorageArchive>(&message.encode_to_vec())?;
    decoded.unknown_fields = top_level_fields(data)
        .difference(&top_level_fields(&decoded.canonical))
        .copied()
        .collect();
    Ok(decoded)
}

pub(crate) fn masked(known: KnownFields) -> KnownFields {
    known
        .with(
            HIGHLIGHT_MESSAGE_TYPE,
            "TSWP.HighlightArchive (text-attribute UUID masked)",
            masked_highlight,
        )
        .with(
            COMMENT_STORAGE_MESSAGE_TYPE,
            "TSD.CommentStorageArchive (creation date and storage UUID masked)",
            masked_comment_storage,
        )
}

fn run(
    name: &'static str,
    format: Format,
    tier: Tier,
    input: &Fixture,
    storage: u64,
    operation: &Operation,
) -> TestResult {
    let mut case = Case::new(name, format, tier, Evidence::Queued, input)
        .host(|source| host_route(source, storage, operation))
        .leaf(|source| Ok(LeafRun::output(leaf_route(source, storage, operation)?)))
        .read("host annotation reads", move |bytes| {
            host_reads(bytes, storage)
        })
        .read("leaf annotation reads", move |bytes| {
            leaf_reads(bytes, storage)
        });
    if tier == Tier::B {
        case = case.known_fields(masked).leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn with(input: &Fixture, storage: u64, operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, storage, operation)?,
    })
}

fn first_highlight(bytes: &[u8], storage: u64) -> TestResult<u64> {
    let draft = Draft::from_bytes(bytes)?;
    let highlight = draft
        .text_highlights(TextStorageId::try_from(storage)?)?
        .into_iter()
        .next()
        .ok_or("no highlight")?;
    Ok(object_id(highlight.id))
}

fn first_comment(bytes: &[u8], storage: u64) -> TestResult<(u64, Option<u64>)> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    let comment = draft
        .text_comments(id)?
        .into_iter()
        .next()
        .ok_or("no comment")?;
    let reply = draft
        .text_comment_replies(id, comment.id())?
        .first()
        .map(|reply| reply_id_value(reply.id()));
    Ok((comment_id_value(comment.id()), reply))
}

/// Run tier-B creation cases: `create` builds the operation from the
/// storage's leading two-character range and its base fixture.
fn creation_cases(
    name: &'static str,
    create: impl Fn(&Fixture, u64, TextRange) -> TestResult<(Fixture, Operation)>,
) -> TestResult<usize> {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, first, _) in text_storages(&input.bytes, 2)? {
            let (base, operation) = create(&input, storage, first)?;
            run(name, format, Tier::B, &base, storage, &operation)
                .map_err(|e| format!("{path} storage {storage} {operation:?}: {e}"))?;
            cases += 1;
        }
    }
    Ok(cases)
}

/// Run tier-A edit cases on a base fixture `prepare` builds; `edits`
/// lists the operations from the base, the storage's two ranges.
fn edit_cases(
    name: &'static str,
    prepare: impl Fn(&Fixture, u64, TextRange) -> TestResult<Fixture>,
    edits: impl Fn(&Fixture, u64, TextRange, TextRange) -> TestResult<Vec<Operation>>,
) -> TestResult<usize> {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, first, wide) in text_storages(&input.bytes, 2)? {
            let base = prepare(&input, storage, first)?;
            for operation in edits(&base, storage, first, wide)? {
                run(name, format, Tier::A, &base, storage, &operation)
                    .map_err(|e| format!("{path} storage {storage} {operation:?}: {e}"))?;
                cases += 1;
            }
        }
    }
    Ok(cases)
}

fn commented(input: &Fixture, storage: u64, first: TextRange) -> TestResult<Fixture> {
    with(input, storage, &Operation::AddComment(first, "Check this"))
}

fn replied(input: &Fixture, storage: u64, first: TextRange) -> TestResult<Fixture> {
    let base = commented(input, storage, first)?;
    let (id, _) = first_comment(&base.bytes, storage)?;
    with(&base, storage, &Operation::AddReply(id, "Agreed"))
}

#[test]
fn text_highlight_add_parity() -> TestResult {
    let cases = creation_cases("text_highlight_add_parity", |input, _, first| {
        Ok((copy(input), Operation::AddHighlight(first)))
    })?;
    let _ = (Tier::B, Evidence::Queued);
    assert!(cases >= 5, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_highlight_edit_parity() -> TestResult {
    let cases = edit_cases(
        "text_highlight_edit_parity",
        |input, storage, first| with(input, storage, &Operation::AddHighlight(first)),
        |base, storage, _, wide| {
            let id = first_highlight(&base.bytes, storage)?;
            Ok(vec![
                Operation::UpdateHighlight(id, wide),
                Operation::RemoveHighlight(id),
            ])
        },
    )?;
    let _ = (Tier::A, Evidence::Queued);
    assert!(cases >= 10, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_comment_add_parity() -> TestResult {
    let cases = creation_cases("text_comment_add_parity", |input, _, first| {
        Ok((copy(input), Operation::AddComment(first, "Check this")))
    })?;
    let _ = (Tier::B, Evidence::Queued);
    assert!(cases >= 5, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_comment_edit_parity() -> TestResult {
    let cases = edit_cases(
        "text_comment_edit_parity",
        commented,
        |base, storage, first, wide| {
            let (id, _) = first_comment(&base.bytes, storage)?;
            Ok(vec![
                Operation::UpdateComment(id, wide, "Check this"),
                Operation::UpdateComment(id, first, "Reworded"),
                Operation::UpdateComment(id, wide, "Both"),
                Operation::RemoveComment(id),
            ])
        },
    )?;
    let _ = (Tier::A, Evidence::Queued);
    assert!(cases >= 20, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_comment_reply_add_parity() -> TestResult {
    let cases = creation_cases("text_comment_reply_add_parity", |input, storage, first| {
        let base = commented(input, storage, first)?;
        let (id, _) = first_comment(&base.bytes, storage)?;
        Ok((base, Operation::AddReply(id, "Agreed")))
    })?;
    let _ = (Tier::B, Evidence::Queued);
    assert!(cases >= 5, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_comment_reply_edit_parity() -> TestResult {
    let cases = edit_cases(
        "text_comment_reply_edit_parity",
        replied,
        |base, storage, _, _| {
            let (id, reply) = first_comment(&base.bytes, storage)?;
            let reply = reply.ok_or("no reply")?;
            Ok(vec![
                Operation::UpdateReply(id, reply, "Disagreed"),
                Operation::RemoveReply(id, reply),
                Operation::RemoveComment(id),
            ])
        },
    )?;
    let _ = (Tier::A, Evidence::Queued);
    assert!(cases >= 15, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_annotation_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, first, wide) in text_storages(&input.bytes, 2)? {
            let highlighted = with(&input, storage, &Operation::AddHighlight(first))?;
            let commented = with(&input, storage, &Operation::AddComment(wide, "Note"))?;
            let (id, _) = first_comment(&commented.bytes, storage)?;
            let replied = with(&commented, storage, &Operation::AddReply(id, "Reply"))?;
            for bytes in [
                &input.bytes,
                &highlighted.bytes,
                &commented.bytes,
                &replied.bytes,
            ] {
                assert_eq!(
                    host_reads(bytes, storage)?,
                    leaf_reads(bytes, storage)?,
                    "{path} storage {storage}"
                );
                reads += 1;
            }
        }
    }
    assert!(reads >= 20, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

fn copy(input: &Fixture) -> Fixture {
    Fixture {
        path: input.path,
        bytes: input.bytes.clone(),
    }
}
