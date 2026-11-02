//! Implementation of the public [`crate::drawable::comment`]
//! transaction: the comment attached directly to a body drawable (a text
//! box, shape, image, movie or audio clip) and its replies (record 0413,
//! the port of the Numbers family of record 0378a).
//!
//! The comment and its replies are `TSD.CommentStorageArchive` objects the
//! drawable's comment reference roots. The publication follows the host
//! editor's copy-on-write comment graph:
//!
//! - setting the body of a comment no other object can observe rewrites it
//!   in place; a shared comment is copied under a fresh identity and the
//!   drawable repointed;
//! - a reply change copies the root (and a rewritten reply) under fresh
//!   identifiers, repoints the drawable, and removes every comment storage
//!   the change leaves unreferenced;
//! - clearing detaches the comment and removes its unreferenced graph;
//! - new comment storages are authored by the generated annotation author,
//!   which is created on first use and removed with its last comment.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::BTreeSet;
use std::fmt;

use litchi_iwa_common::comment::AuthorId;
use litchi_iwa_common::wire::{
    patch_nested_length_delimited_field, patch_nested_varint_field, patch_varint_field,
    repeated_length_delimited_payloads, transform_length_delimited_fields_at_path,
};
use litchi_iwa_core::{ArchiveObject, MessageInfo, RawMessage};
use litchi_iwa_text_wire::annotation::{self as engine, COMMENT_STORAGE_MESSAGE_TYPE, ReplyChange};
use thiserror::Error as ThisError;

use super::Package;
use super::text_box_comment_graph::{Graph, comment_storage, engine_fault, replace};
use super::text_box_structure::{
    self as structure, Fault, FaultResult, Published, Working,
    metadata::{self, Transition},
    references,
};
use crate::selector::DrawableSelector;

structure::text_box_family!("drawable comment", Wire(litchi_iwa_common::LimitKind));

/// `TSD.CommentStorageArchive.replies`.
const REPLIES_FIELD: u32 = 4;
/// `TSP.Reference.identifier`.
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

/// A typed, content-redacted drawable comment failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted drawable comment source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} drawable comment items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("drawable comment candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("drawable comment patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No drawable matches the selector.
    #[error("the body has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The drawable takes no direct comment (tables keep cell comments).
    #[error("a Pages body {kind} takes no direct comment")]
    NotCommentable {
        /// The selected kind.
        kind: &'static str,
    },
    /// A reply change was staged on a drawable without a comment.
    #[error("the drawable has no comment")]
    NoComment,
    /// No reply exists at the position.
    #[error("the comment has no reply at position {position}")]
    ReplyNotFound {
        /// The requested position.
        position: usize,
    },
    /// Commit was requested before staging a change.
    #[error("no drawable comment change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten comment graph.
    #[error("a surviving object depends on the drawable comment graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support drawable comment edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::DrawableNotFound { .. } | Self::NoComment | Self::ReplyNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::NotCommentable { .. } => RefusalKind::UnsupportedTarget,
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// The comment attached directly to a body drawable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawableComment {
    body: String,
    replies: Vec<String>,
}

impl DrawableComment {
    /// The comment's body.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }

    /// The bodies of its replies, oldest first.
    #[must_use]
    pub fn replies(&self) -> &[String] {
        &self.replies
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Staged {
    Set(String),
    Clear,
    AddReply(String),
    SetReply(usize, String),
    RemoveReply(usize),
}

/// Comment changes staged against one body text box or shape.
pub struct Edit<'a> {
    source: &'a Package,
    selector: DrawableSelector,
    changes: Vec<Staged>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

fn owned(text: &str) -> Result<String, Error> {
    let mut value = String::new();
    value
        .try_reserve_exact(text.len())
        .map_err(|_error| Error::Allocation { amount: text.len() })?;
    value.push_str(text);
    Ok(value)
}

impl Edit<'_> {
    fn stage(&mut self, change: Staged) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change);
        Ok(self)
    }

    /// Stage setting the comment's body, creating the comment when the
    /// drawable has none.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set(&mut self, body: &str) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::Set(body))
    }

    /// Stage removing the comment and its replies (a no-op without one).
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        self.stage(Staged::Clear)
    }

    /// Stage appending a reply.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn add_reply(&mut self, body: &str) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::AddReply(body))
    }

    /// Stage setting the body of reply `position` (oldest first).
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn set_reply(&mut self, position: usize, body: &str) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::SetReply(position, body))
    }

    /// Stage removing reply `position` (oldest first).
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove_reply(&mut self, position: usize) -> Result<&mut Self, Error> {
        self.stage(Staged::RemoveReply(position))
    }

    /// Validate and atomically publish the staged changes, in order; an
    /// edit that changes nothing is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], [`Refusal::NoComment`],
    /// [`Refusal::ReplyNotFound`], source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            let before = read(base, self.selector)?;
            if let Some(published) = apply_one(base, self.selector, change)? {
                let after = read(&published.package, self.selector)?;
                if !holds(before.as_ref(), after.as_ref(), change) {
                    return Err(Error::Verification);
                }
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Ok(Commit::unchanged(self.source));
        };
        Ok(Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        ))
    }
}

/// Whether `after` is `before` with `change` applied.
fn holds(
    before: Option<&DrawableComment>,
    after: Option<&DrawableComment>,
    change: &Staged,
) -> bool {
    match change {
        Staged::Set(body) => after.is_some_and(|comment| {
            comment.body == *body
                && before.map_or(&[][..], |comment| comment.replies.as_slice()) == comment.replies
        }),
        Staged::Clear => after.is_none(),
        Staged::AddReply(body) => match (before, after) {
            (Some(before), Some(after)) => {
                after.body == before.body
                    && after.replies.len() == before.replies.len() + 1
                    && after.replies[..before.replies.len()] == before.replies[..]
                    && after.replies.last() == Some(body)
            },
            _ => false,
        },
        Staged::SetReply(position, body) => match (before, after) {
            (Some(before), Some(after)) => {
                let mut expected = before.replies.clone();
                if let Some(reply) = expected.get_mut(*position) {
                    reply.clone_from(body);
                }
                after.body == before.body && after.replies == expected
            },
            _ => false,
        },
        Staged::RemoveReply(position) => match (before, after) {
            (Some(before), Some(after)) => {
                let mut expected = before.replies.clone();
                if *position < expected.len() {
                    expected.remove(*position);
                }
                after.body == before.body && after.replies == expected
            },
            _ => false,
        },
    }
}

/// The comment-reference route of a drawable payload (the host editor's
/// Numbers `drawable_comment_wire_path`).
const fn route(message_type: u32) -> Option<&'static [u32]> {
    match message_type {
        3_002 => Some(&[6]),
        3_004 | 3_005 | 3_006 | 3_007 | 3_008 | 5_021 => Some(&[1, 6]),
        3_009 | 2_011 => Some(&[1, 1, 6]),
        2_014 | 7 => Some(&[1, 1, 1, 6]),
        _ => None,
    }
}

/// The comment storage a drawable payload's route names, if any.
fn comment_reference(payload: &[u8], path: &[u32]) -> FaultResult<Option<u64>> {
    let mut current = payload;
    for (depth, number) in path.iter().enumerate() {
        let found = repeated_length_delimited_payloads(current, *number)
            .map_err(|error| structure::map_wire(&error))?;
        match found.as_slice() {
            [] => return Ok(None),
            [only] if depth + 1 == path.len() => {
                return Ok(Some(references::reference_identifier(only)?));
            },
            [only] => current = only,
            _ => return Err(Fault::Invalid),
        }
    }
    Err(Fault::Invalid)
}

/// The selected drawable: (object, message index, route).
struct Target {
    drawable: u64,
    index: usize,
    path: &'static [u32],
}

fn target(package: &Package, selector: DrawableSelector) -> Result<Target, Error> {
    use super::body_drawable_graph::{self as graph, Budget, ObjectIndex};
    let fault = structure::graph_fault;
    let mut budget = Budget::new(package).map_err(fault)?;
    let index = ObjectIndex::new(package, &mut budget).map_err(fault)?;
    let body = graph::resolve_body(&index, &mut budget).map_err(fault)?;
    let selected = graph::select_drawable(&index, &body, selector, &mut budget)
        .map_err(fault)?
        .ok_or(Refusal::DrawableNotFound { selector })?;
    let path = route(graph::message_type(selector)).ok_or(Refusal::NotCommentable {
        kind: graph::kind_name(selector),
    })?;
    Ok(Target {
        drawable: selected.anchor.drawable.get(),
        index: selected.message_index,
        path,
    })
}

fn current_comment(working: &Working<'_>, target: &Target) -> FaultResult<Option<u64>> {
    let object = working.object(target.drawable)?;
    comment_reference(&object.messages[target.index].data, target.path)
}

/// The decoded direct-reply graph of `root`: its replies, each distinct and
/// decodable (the host's `validate_direct_reply_graph`).
fn replies_of(working: &Working<'_>, root: u64) -> FaultResult<Vec<u64>> {
    let (_, storage) = comment_storage(working, root)?;
    let replies = storage.replies().to_vec();
    for (position, reply) in replies.iter().enumerate() {
        if *reply == root || replies[..position].contains(reply) {
            return Err(Fault::Invalid);
        }
        comment_storage(working, *reply)?;
    }
    Ok(replies)
}

fn read(package: &Package, selector: DrawableSelector) -> Result<Option<DrawableComment>, Error> {
    let target = target(package, selector)?;
    let working = Working::new(package)?;
    let Some(root) = current_comment(&working, &target)? else {
        return Ok(None);
    };
    let (_, storage) = comment_storage(&working, root)?;
    let mut replies = Vec::new();
    for reply in replies_of(&working, root)? {
        structure::reserve(&mut replies, 1)?;
        replies.push(owned(comment_storage(&working, reply)?.1.body())?);
    }
    Ok(Some(DrawableComment {
        body: owned(storage.body())?,
        replies,
    }))
}

/// Remove `old` and push `new` into the message's aggregate references
/// (the host's `update_reference_metadata`).
fn update_references(info: &mut MessageInfo, old: Option<u64>, new: Option<u64>) {
    if let Some(old) = old {
        info.object_references.retain(|reference| *reference != old);
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|reference| *reference != old);
        }
    }
    if let Some(new) = new
        && !info.object_references.contains(&new)
    {
        info.object_references.push(new);
    }
}

/// Repoint the drawable's comment reference from `old` to `new` (the host's
/// `replace_drawable_comment_reference`).
fn repoint(
    working: &mut Working<'_>,
    target: &Target,
    old: Option<u64>,
    new: Option<u64>,
) -> FaultResult<()> {
    let limits = working.archive_limits();
    let object = working.object_mut(target.drawable)?;
    let message = object.messages.get(target.index).ok_or(Fault::Invalid)?;
    if comment_reference(&message.data, target.path)? != old {
        return Err(Fault::Invalid);
    }
    let wire = |error: litchi_iwa_common::Error| structure::map_wire(&error);
    let data = match (old, new) {
        (Some(_), Some(identifier)) => {
            let mut path = target.path.to_vec();
            path.push(REFERENCE_IDENTIFIER_FIELD);
            patch_nested_varint_field(&message.data, &path, true, Some(identifier)).map_err(wire)?
        },
        (None, Some(identifier)) => patch_nested_length_delimited_field(
            &message.data,
            target.path,
            false,
            Some(&references::reference_payload(identifier)?),
        )
        .map_err(wire)?,
        (Some(_) | None, None) => {
            patch_nested_length_delimited_field(&message.data, target.path, old.is_some(), None)
                .map_err(wire)?
        },
    };
    if comment_reference(&data, target.path)? != new {
        return Err(Fault::Invalid);
    }
    let message_type = message.type_;
    let mut staged = object.clone();
    staged
        .replace_message(
            target.index,
            RawMessage {
                type_: message_type,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    update_references(
        staged
            .archive_info
            .message_infos
            .get_mut(target.index)
            .ok_or(Fault::Invalid)?,
        old,
        new,
    );
    staged
        .validate(limits)
        .map_err(|error| structure::map_core(&error))?;
    *object = staged;
    Ok(())
}

/// Patch reply reference `old` of `root` into `new` in place, keeping the
/// order (the host's `update_comment_reply_reference` replace arm).
fn replace_reply(working: &mut Working<'_>, root: u64, old: u64, new: u64) -> FaultResult<()> {
    let (index, _) = comment_storage(working, root)?;
    let original = working.object(root)?.messages[index].data.clone();
    let data = transform_length_delimited_fields_at_path(&original, &[REPLIES_FIELD], |payload| {
        if references::reference_identifier(payload).ok() == Some(old) {
            patch_varint_field(payload, REFERENCE_IDENTIFIER_FIELD, true, Some(new))
        } else {
            Ok(payload.to_vec())
        }
    })
    .map_err(|error: litchi_iwa_common::Error| structure::map_wire(&error))?;
    let limits = working.archive_limits();
    let object = working.object_mut(root)?;
    let mut staged = object.clone();
    staged
        .replace_message(
            index,
            RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(|error| structure::map_core(&error))?;
    update_references(
        &mut staged.archive_info.message_infos[index],
        Some(old),
        Some(new),
    );
    *object = staged;
    Ok(())
}

impl Graph<'_, '_> {
    /// Stamp the component of `member`.
    fn stamp(&mut self, member: &str) -> FaultResult<()> {
        let component = self.component(member)?;
        self.transition.stamp(component)
    }

    /// Copy comment storage `source` exactly as `identifier` beside it (the
    /// host's `clone_comment_storage_exact`); its member.
    fn clone_exact(&mut self, source: u64, identifier: u64) -> FaultResult<String> {
        let member = self.working.locate(source)?.ok_or(Fault::Invalid)?;
        let original = self.working.object(source)?;
        if original.messages.len() != 1
            || original.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE
        {
            return Err(Fault::Unsupported);
        }
        comment_storage(self.working, source)?;
        let mut object = ArchiveObject::new(
            identifier,
            original.messages.clone(),
            self.working.archive_limits(),
        )
        .map_err(|error| structure::map_core(&error))?;
        object.archive_info.should_merge = original.archive_info.should_merge;
        object.archive_info.message_infos = original.archive_info.message_infos.clone();
        self.working.insert(&member, object)?;
        self.register_uuid(&member, identifier)?;
        Ok(member)
    }

    /// Copy comment storage `source` as `identifier` with `body` and a
    /// fresh storage UUID (the host's `clone_comment_storage`).
    fn clone_with_body(&mut self, source: u64, identifier: u64, body: &str) -> FaultResult<String> {
        let member = self.working.locate(source)?.ok_or(Fault::Invalid)?;
        let original = self.working.object(source)?.clone();
        if original.messages.len() != 1
            || original.messages[0].type_ != COMMENT_STORAGE_MESSAGE_TYPE
        {
            return Err(Fault::Unsupported);
        }
        let uuid = self.fresh_uuid(identifier)?;
        let mut encoded = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut encoded, 1, uuid.lower())
            .and_then(|()| {
                litchi_iwa_common::wire::append_varint_field(&mut encoded, 2, uuid.upper())
            })
            .map_err(|error| structure::map_wire(&error))?;
        let data = engine::comment_storage_with_body(
            &original.messages[0].data,
            body,
            litchi_iwa_common::WireLimits::DEFAULT,
        )
        .map_err(|error| engine_fault(&error))?;
        let data =
            litchi_iwa_common::wire::patch_length_delimited_field(&data, 5, true, Some(&encoded))
                .map_err(|error| structure::map_wire(&error))?;
        let mut object = ArchiveObject::new(
            identifier,
            vec![RawMessage {
                type_: COMMENT_STORAGE_MESSAGE_TYPE,
                data,
            }],
            self.working.archive_limits(),
        )
        .map_err(|error| structure::map_core(&error))?;
        object.archive_info.should_merge = original.archive_info.should_merge;
        let length = object.archive_info.message_infos[0].length;
        object.archive_info.message_infos[0] = original.archive_info.message_infos[0].clone();
        object.archive_info.message_infos[0].length = length;
        self.working.insert(&member, object)?;
        self.register_uuid(&member, identifier)?;
        Ok(member)
    }

    /// Whether anything but its own identity still names comment storage
    /// `identifier` (the host's `comment_object_is_referenced`).
    fn is_referenced(&self, identifier: u64) -> FaultResult<bool> {
        let owning = match self.working.locate(identifier)? {
            Some(member) => Some(self.component(&member)?),
            None => None,
        };
        for name in self.working.component_names() {
            for object in &self.working.archive(&name)?.objects {
                for (message, info) in object
                    .messages
                    .iter()
                    .zip(&object.archive_info.message_infos)
                {
                    if info.object_references.contains(&identifier)
                        || info
                            .field_infos
                            .iter()
                            .any(|field| field.object_references.contains(&identifier))
                    {
                        return Ok(true);
                    }
                    if let Some(path) = route(message.type_)
                        && comment_reference(&message.data, path)? == Some(identifier)
                    {
                        return Ok(true);
                    }
                    if message.type_ == COMMENT_STORAGE_MESSAGE_TYPE
                        && engine::comment_storage(
                            &message.data,
                            litchi_iwa_common::WireLimits::DEFAULT,
                        )
                        .map_err(|error| engine_fault(&error))?
                        .replies()
                        .contains(&identifier)
                    {
                        return Ok(true);
                    }
                }
            }
        }
        let registry = &self.registry;
        Ok(registry
            .uuids
            .iter()
            .any(|binding| binding.object == identifier && Some(binding.component) != owning)
            || registry.externals.iter().any(|external| {
                external.object == identifier
                    && !self.transition.external_removals.contains(external)
            })
            || registry.data_owners.contains(&identifier))
    }

    /// Remove the comment storages rooted at `root` that nothing references
    /// any more, then the generated authors they leave unused (the host's
    /// `remove_unreferenced_comment_graph` and
    /// `cleanup_removed_comment_graph`).
    fn remove_unreferenced(&mut self, root: u64) -> FaultResult<()> {
        let mut pending = vec![root];
        let mut visited = BTreeSet::new();
        let mut authors = BTreeSet::new();
        while let Some(identifier) = pending.pop() {
            if !visited.insert(identifier) || self.is_referenced(identifier)? {
                continue;
            }
            let Some(member) = self.working.locate(identifier)? else {
                continue;
            };
            let (_, storage) = comment_storage(self.working, identifier)?;
            if let Some(author) = storage.metadata().author() {
                authors.insert(AuthorId::get(author));
            }
            self.stamp(&member)?;
            self.forget(identifier)?;
            pending.extend_from_slice(storage.replies());
        }
        for author in authors {
            let member = self.working.locate(author)?;
            if self.remove_author_if_unused(author)?
                && let Some(member) = member
            {
                self.stamp(&member)?;
            }
        }
        Ok(())
    }

    /// Ensure the generated author for a new comment storage, stamping the
    /// author storage's component when the author is created.
    fn author(&mut self) -> FaultResult<Option<u64>> {
        let Some((author, created)) = self.ensure_author()? else {
            return Ok(None);
        };
        if created && let Some(author_member) = self.author_member()? {
            self.stamp(&author_member)?;
        }
        Ok(Some(author))
    }
}

#[allow(
    clippy::too_many_lines,
    reason = "one arm per host drawable comment route"
)]
fn apply_one(
    source: &Package,
    selector: DrawableSelector,
    change: &Staged,
) -> Result<Option<Published>, Error> {
    let target = target(source, selector)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let drawable_member = working
        .locate(target.drawable)?
        .ok_or(Error::InvalidSource)?;
    let current = current_comment(&working, &target)?;
    let mut graph = Graph {
        working: &mut working,
        registry,
        transition: Transition::default(),
    };
    let limits = litchi_iwa_common::WireLimits::DEFAULT;
    match change {
        Staged::Set(body) => {
            if let Some(storage) = current {
                let (index, old) = comment_storage(graph.working, storage)?;
                if old.body() == body {
                    return Ok(None);
                }
                replies_of(graph.working, storage)?;
                let users = drawable_users(graph.working, storage)?;
                let owned = users == 1 && {
                    // Prove no other owner observes the storage: detach the
                    // selected edge in a scratch copy and take the census.
                    let mut scratch = Working::new(source)?;
                    repoint(&mut scratch, &target, Some(storage), None)?;
                    let registry = metadata::registry(&scratch)?
                        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
                    let scratch_graph = Graph {
                        working: &mut scratch,
                        registry,
                        transition: Transition::default(),
                    };
                    !scratch_graph.is_referenced(storage)?
                };
                let storage_member = graph.working.locate(storage)?.ok_or(Error::InvalidSource)?;
                if owned {
                    let data = engine::comment_storage_with_body(
                        &graph.working.object(storage)?.messages[index].data,
                        body,
                        limits,
                    )
                    .map_err(|error| engine_fault(&error))?;
                    replace(graph.working, storage, index, data, None, None)?;
                    graph.stamp(&storage_member)?;
                } else {
                    let copy = allocate(graph.working)?;
                    graph.clone_with_body(storage, copy, body)?;
                    repoint(graph.working, &target, Some(storage), Some(copy))?;
                    graph.stamp(&storage_member)?;
                    graph.stamp(&drawable_member)?;
                }
            } else {
                let author = graph.author()?;
                let storage = allocate(graph.working)?;
                graph.insert_comment_storage(&drawable_member, storage, body, author)?;
                repoint(graph.working, &target, None, Some(storage))?;
                graph.add_author_reference(&drawable_member, author)?;
                graph.stamp(&drawable_member)?;
            }
        },
        Staged::Clear => {
            let Some(storage) = current else {
                return Ok(None);
            };
            replies_of(graph.working, storage)?;
            repoint(graph.working, &target, Some(storage), None)?;
            graph.stamp(&drawable_member)?;
            graph.remove_unreferenced(storage)?;
        },
        Staged::AddReply(body) => {
            let root = current.ok_or(Refusal::NoComment)?;
            replies_of(graph.working, root)?;
            let author = graph.author()?;
            let copy = allocate(graph.working)?;
            let root_member = graph.clone_exact(root, copy)?;
            repoint(graph.working, &target, Some(root), Some(copy))?;
            let reply = allocate(graph.working)?;
            graph.insert_comment_storage(&root_member, reply, body, author)?;
            let (index, _) = comment_storage(graph.working, copy)?;
            let data = engine::comment_storage_with_reply(
                &graph.working.object(copy)?.messages[index].data,
                ReplyChange::Add(reply),
                limits,
            )
            .map_err(|error| engine_fault(&error))?;
            replace(graph.working, copy, index, data, Some(reply), None)?;
            graph.add_author_reference(&root_member, author)?;
            graph.stamp(&drawable_member)?;
            graph.stamp(&root_member)?;
            graph.remove_unreferenced(root)?;
        },
        Staged::SetReply(position, body) => {
            let root = current.ok_or(Refusal::NoComment)?;
            let replies = replies_of(graph.working, root)?;
            let reply = *replies.get(*position).ok_or(Refusal::ReplyNotFound {
                position: *position,
            })?;
            if comment_storage(graph.working, reply)?.1.body() == body {
                return Ok(None);
            }
            let copy = allocate(graph.working)?;
            let root_member = graph.clone_exact(root, copy)?;
            repoint(graph.working, &target, Some(root), Some(copy))?;
            let new_reply = allocate(graph.working)?;
            let reply_member = graph.clone_exact(reply, new_reply)?;
            let (index, _) = comment_storage(graph.working, new_reply)?;
            let data = engine::comment_storage_with_body(
                &graph.working.object(new_reply)?.messages[index].data,
                body,
                limits,
            )
            .map_err(|error| engine_fault(&error))?;
            replace(graph.working, new_reply, index, data, None, None)?;
            replace_reply(graph.working, copy, reply, new_reply)?;
            graph.stamp(&drawable_member)?;
            graph.stamp(&root_member)?;
            graph.stamp(&reply_member)?;
            graph.remove_unreferenced(root)?;
        },
        Staged::RemoveReply(position) => {
            let root = current.ok_or(Refusal::NoComment)?;
            let replies = replies_of(graph.working, root)?;
            let reply = *replies.get(*position).ok_or(Refusal::ReplyNotFound {
                position: *position,
            })?;
            let copy = allocate(graph.working)?;
            let root_member = graph.clone_exact(root, copy)?;
            repoint(graph.working, &target, Some(root), Some(copy))?;
            let (index, _) = comment_storage(graph.working, copy)?;
            let data = engine::comment_storage_with_reply(
                &graph.working.object(copy)?.messages[index].data,
                ReplyChange::Remove(reply),
                limits,
            )
            .map_err(|error| engine_fault(&error))?;
            replace(graph.working, copy, index, data, None, Some(reply))?;
            graph.stamp(&drawable_member)?;
            graph.stamp(&root_member)?;
            graph.remove_unreferenced(root)?;
        },
    }
    let transition = graph.transition;
    Ok(Some(working.publish(&transition)?))
}

fn allocate(working: &mut Working<'_>) -> FaultResult<u64> {
    working.allocate(1)?.first().copied().ok_or(Fault::Invalid)
}

/// How many drawable payloads name comment storage `storage`.
fn drawable_users(working: &Working<'_>, storage: u64) -> FaultResult<usize> {
    let mut users = 0_usize;
    for object in structure::source_objects(working.source()) {
        let identifier = object.archive_info.identifier.ok_or(Fault::Invalid)?;
        let Ok(object) = working.object(identifier) else {
            continue;
        };
        for message in &object.messages {
            if let Some(path) = route(message.type_)
                && comment_reference(&message.data, path)? == Some(storage)
            {
                users = users.saturating_add(1);
            }
        }
    }
    Ok(users)
}

impl Package {
    /// The comment attached directly to one body text box or shape, with its
    /// replies, or `None` without one.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NotCommentable`], source errors.
    pub fn body_drawable_comment(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Option<DrawableComment>, Error> {
        read(self, drawable.into())
    }

    /// Start an immutable comment edit of one body text box or shape.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NotCommentable`], source errors.
    pub fn edit_body_drawable_comment(
        &self,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = drawable.into();
        target(self, selector)?;
        Ok(Edit {
            source: self,
            selector,
            changes: Vec::new(),
        })
    }

    /// Apply a drawable comment patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_body_drawable_comment(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
