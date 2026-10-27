//! Implementation of the public [`crate::sheet::drawable::annotation`]
//! transaction: highlights, text comments and replies of the text a sheet
//! text box or shape owns, on W1-D's annotation and ranged engines
//! ([`litchi_iwa_text_wire::annotation`], record 0374), record 0378.
//!
//! A highlight or comment is a `TSWP.HighlightArchive` object named by a
//! range of the storage's highlight table; it points at a
//! `TSD.CommentStorageArchive` holding the body (empty for a highlight),
//! the author and the replies (each one more comment storage). Creating a
//! comment or reply makes sure the document's generated annotation author
//! exists (the host editor's `ensure_annotation_author`); removing the last
//! use of that author removes it again. The publication follows the host
//! text editor step by step.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::fmt;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::comment::AuthorId;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_text::position::TextRange;
use litchi_iwa_text_wire::annotation::{
    self as engine, Annotation, COMMENT_STORAGE_MESSAGE_TYPE, HIGHLIGHT_MESSAGE_TYPE, Objects,
    ReplyChange,
};
use litchi_iwa_text_wire::ranged::{self, Change as RangedChange, Table};
use litchi_iwa_text_wire::smart_field::TextAttributeUuid;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_comment_graph::{Graph, replace};
use super::sheet_drawable_graph::{self as graph, TextTarget};
use super::structure::{
    self, Published, Working,
    metadata::{self, Transition},
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::DrawableSelector;

structure::structural_family!("text annotation", Wire(litchi_iwa_common::LimitKind));

/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;

/// A typed, content-redacted annotation failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted text-annotation source is invalid")]
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
    #[error("could not allocate {amount} text-annotation items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("text-annotation candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("text-annotation patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The drawable owns no text (it is not a text box or shape).
    #[error("a Numbers sheet {kind} owns no text")]
    NoText {
        /// The selected kind.
        kind: &'static str,
    },
    /// No highlight, comment or reply is at the selected position.
    #[error("no annotation at position {position}")]
    AnnotationNotFound {
        /// The selected zero-based position.
        position: usize,
    },
    /// A range bound is past the text or splits a surrogate pair.
    #[error("the range is not on text boundaries")]
    NotABoundary,
    /// The range overlaps another annotation's range.
    #[error("the range overlaps another annotation")]
    Occupied,
    /// A comment or reply body is empty (an empty body is a highlight).
    #[error("a comment body must not be empty")]
    EmptyBody,
    /// Commit was requested before staging a change.
    #[error("no text-annotation change is staged")]
    NoStagedOperation,
    /// A surviving object depends on the rewritten graph (overlapping
    /// highlights, a shared comment storage).
    #[error("a surviving object depends on the annotation graph")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support text-annotation edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound
            | Self::DrawableNotFound { .. }
            | Self::AnnotationNotFound { .. } => RefusalKind::NotFound,
            Self::NoText { .. } => RefusalKind::UnsupportedTarget,
            Self::NotABoundary | Self::Occupied | Self::EmptyBody => RefusalKind::InvalidValue,
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

fn from_ranged(error: ranged::Error) -> Error {
    match error {
        ranged::Error::Refused(ranged::Refusal::Occupied) => Refusal::Occupied.into(),
        ranged::Error::Refused(ranged::Refusal::NotABoundary) => Refusal::NotABoundary.into(),
        ranged::Error::Refused(_) => Refusal::UnsupportedDependency.into(),
        ranged::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn from_engine(error: engine::Error) -> Error {
    match error {
        engine::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

impl From<litchi_iwa_common::Error> for Error {
    fn from(error: litchi_iwa_common::Error) -> Self {
        structure::map_wire(&error).into()
    }
}

/// One highlight of a drawable's text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Highlight {
    range: TextRange,
}

impl Highlight {
    /// The highlighted text range.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }
}

/// One text comment of a drawable's text with its replies.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    range: TextRange,
    body: String,
    replies: Vec<String>,
}

impl Comment {
    /// The commented text range.
    #[must_use]
    pub const fn range(&self) -> TextRange {
        self.range
    }

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
    AddHighlight(TextRange),
    UpdateHighlight(usize, TextRange),
    RemoveHighlight(usize),
    AddComment(TextRange, String),
    UpdateComment(usize, TextRange, String),
    RemoveComment(usize),
    AddReply(usize, String),
    UpdateReply(usize, usize, String),
    RemoveReply(usize, usize),
}

/// Annotation changes staged against the text of one sheet drawable.
pub struct Edit<'a> {
    source: &'a Package,
    sheet: usize,
    selector: DrawableSelector,
    changes: Vec<Staged>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.sheet)
            .field("selector", &self.selector)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

fn owned(text: &str) -> Result<String, Error> {
    if text.is_empty() {
        return Err(Refusal::EmptyBody.into());
    }
    let mut output = String::new();
    output
        .try_reserve_exact(text.len())
        .map_err(|_error| Error::Allocation { amount: text.len() })?;
    output.push_str(text);
    Ok(output)
}

impl Edit<'_> {
    fn stage(&mut self, change: Staged) -> Result<&mut Self, Error> {
        structure::reserve(&mut self.changes, 1)?;
        self.changes.push(change);
        Ok(self)
    }

    /// Stage a new highlight over `range`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn add_highlight(&mut self, range: TextRange) -> Result<&mut Self, Error> {
        self.stage(Staged::AddHighlight(range))
    }

    /// Stage moving the highlight at zero-based `position` (in text order,
    /// after the earlier changes) to `range`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn update_highlight(
        &mut self,
        position: usize,
        range: TextRange,
    ) -> Result<&mut Self, Error> {
        self.stage(Staged::UpdateHighlight(position, range))
    }

    /// Stage removing the highlight at zero-based `position`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove_highlight(&mut self, position: usize) -> Result<&mut Self, Error> {
        self.stage(Staged::RemoveHighlight(position))
    }

    /// Stage a new comment over `range` with `body`.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyBody`], [`Error::Allocation`].
    pub fn add_comment(&mut self, range: TextRange, body: &str) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::AddComment(range, body))
    }

    /// Stage moving the comment at zero-based `position` to `range` with
    /// `body`.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyBody`], [`Error::Allocation`].
    pub fn update_comment(
        &mut self,
        position: usize,
        range: TextRange,
        body: &str,
    ) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::UpdateComment(position, range, body))
    }

    /// Stage removing the comment at zero-based `position` and its replies.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove_comment(&mut self, position: usize) -> Result<&mut Self, Error> {
        self.stage(Staged::RemoveComment(position))
    }

    /// Stage a new reply to the comment at zero-based `comment`.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyBody`], [`Error::Allocation`].
    pub fn add_reply(&mut self, comment: usize, body: &str) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::AddReply(comment, body))
    }

    /// Stage a new body for reply `reply` of comment `comment`.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyBody`], [`Error::Allocation`].
    pub fn update_reply(
        &mut self,
        comment: usize,
        reply: usize,
        body: &str,
    ) -> Result<&mut Self, Error> {
        let body = owned(body)?;
        self.stage(Staged::UpdateReply(comment, reply, body))
    }

    /// Stage removing reply `reply` of comment `comment`.
    ///
    /// # Errors
    ///
    /// [`Error::Allocation`].
    pub fn remove_reply(&mut self, comment: usize, reply: usize) -> Result<&mut Self, Error> {
        self.stage(Staged::RemoveReply(comment, reply))
    }

    /// Validate and atomically publish the staged changes, in order.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`], the range and selection refusals,
    /// source and verification errors.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for change in &self.changes {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = apply_one(base, self.sheet, self.selector, change)? {
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            return Commit::unchanged(self.source);
        };
        Commit::new(
            self.source,
            Published {
                package,
                touched_components: touched,
                deleted_previews: 0,
            },
        )
    }
}

fn storage_of(package: &Package, sheet: usize, selector: DrawableSelector) -> Result<u64, Error> {
    match graph::text_storage(package, sheet, selector)? {
        TextTarget::Storage { storage } => Ok(storage),
        TextTarget::NoText(kind) => Err(Refusal::NoText { kind: kind.name() }.into()),
        TextTarget::NotFound => Err(Refusal::DrawableNotFound { selector }.into()),
    }
}

/// The working set's objects, as the annotation engine reads them.
struct WorkingObjects<'w, 'p>(&'w Working<'p>);

impl Objects for WorkingObjects<'_, '_> {
    fn messages(&self, identifier: u64) -> Option<Vec<(u32, &[u8])>> {
        let object = self.0.object(identifier).ok()?;
        Some(
            object
                .messages
                .iter()
                .map(|message| (message.type_, message.data.as_slice()))
                .collect(),
        )
    }
}

/// The storage's annotations in text order: highlights, then comments.
fn split(annotations: Vec<Annotation>) -> (Vec<Annotation>, Vec<Annotation>) {
    annotations.into_iter().partition(Annotation::is_highlight)
}

fn annotations_of(working: &Working<'_>, storage: u64) -> Result<Vec<Annotation>, Error> {
    let object = working.object(storage)?;
    let index = structure::unique_message(object, &[STORAGE_MESSAGE_TYPE])?;
    engine::annotations(
        &object.messages[index].data,
        &WorkingObjects(working),
        WireLimits::DEFAULT,
    )
    .map_err(from_engine)
}

fn storage_payload(working: &Working<'_>, storage: u64) -> Result<(usize, Vec<u8>), Error> {
    let object = working.object(storage)?;
    let index = structure::unique_message(object, &[STORAGE_MESSAGE_TYPE])?;
    Ok((index, object.messages[index].data.clone()))
}

/// A deterministic text-attribute UUID for a new highlight.
fn highlight_uuid(working: &Working<'_>, object: u64) -> Result<TextAttributeUuid, Error> {
    let seed = metadata::identity_seed(working, object)?;
    let low = metadata::mix(seed ^ object);
    let high = metadata::mix(low ^ 0x6c69_7463_6869_a378);
    let text = format!(
        "{:08x}-{:04x}-4{:03x}-{:04x}-{:012x}",
        high >> 32,
        (high >> 16) & 0xffff,
        high & 0x0fff,
        0x8000 | ((low >> 48) & 0x3fff),
        low & 0xffff_ffff_ffff
    );
    TextAttributeUuid::new(&text).map_err(|_error| Error::InvalidSource)
}

fn pick(list: &[Annotation], position: usize) -> Result<Annotation, Error> {
    list.get(position)
        .cloned()
        .ok_or(Error::Refused(Refusal::AnnotationNotFound { position }))
}

#[allow(clippy::too_many_lines, reason = "one arm per host annotation route")]
fn apply_one(
    source: &Package,
    sheet: usize,
    selector: DrawableSelector,
    change: &Staged,
) -> Result<Option<Published>, Error> {
    let storage = storage_of(source, sheet, selector)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let registry =
        metadata::registry(&working)?.ok_or(Error::Refused(Refusal::UnsupportedSource))?;
    let member = working.locate(storage)?.ok_or(Error::InvalidSource)?;
    let (highlights, comments) = split(annotations_of(&working, storage)?);
    let limits = WireLimits::DEFAULT;
    let mut context = Graph {
        working: &mut working,
        registry,
        transition: Transition::default(),
    };
    match change {
        Staged::AddHighlight(range) | Staged::AddComment(range, _) => {
            let body = match change {
                Staged::AddComment(_, body) => body.as_str(),
                _ => "",
            };
            let author = context.ensure_author()?.map(|(author, _)| author);
            let ids = context.working.allocate(2)?;
            let (object, comment_storage) = (ids[0], ids[1]);
            let (index, payload) = storage_payload(context.working, storage)?;
            let rewrite = ranged::rewrite(
                &payload,
                Table::Highlight,
                &RangedChange::Add {
                    object,
                    range: *range,
                },
                limits,
            )
            .map_err(from_ranged)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            replace(
                context.working,
                storage,
                index,
                rewrite.into_payload(),
                added,
                removed,
            )?;
            let archive_limits = context.working.archive_limits();
            let mut highlight = ArchiveObject::new(
                object,
                vec![RawMessage {
                    type_: HIGHLIGHT_MESSAGE_TYPE,
                    data: engine::encode_highlight(
                        comment_storage,
                        &highlight_uuid(context.working, object)?,
                        limits,
                    )
                    .map_err(from_engine)?,
                }],
                archive_limits,
            )
            .map_err(|error| structure::map_core(&error))?;
            highlight.archive_info.message_infos[0]
                .object_references
                .push(comment_storage);
            context.working.insert(&member, highlight)?;
            context.insert_comment_storage(&member, comment_storage, body, author)?;
            context.add_author_reference(&member, author)?;
        },
        Staged::UpdateHighlight(position, range) | Staged::UpdateComment(position, range, _) => {
            let (current, body) = match change {
                Staged::UpdateComment(_, _, body) => (pick(&comments, *position)?, body.as_str()),
                _ => (pick(&highlights, *position)?, ""),
            };
            let mut changed = false;
            if current.range() != *range {
                let (index, payload) = storage_payload(context.working, storage)?;
                let rewrite = ranged::rewrite(
                    &payload,
                    Table::Highlight,
                    &RangedChange::Move {
                        object: current.object(),
                        range: *range,
                    },
                    limits,
                )
                .map_err(from_ranged)?;
                replace(
                    context.working,
                    storage,
                    index,
                    rewrite.into_payload(),
                    None,
                    None,
                )?;
                changed = true;
            }
            if current.root().body() != body {
                let root = current.root().storage();
                let object = context.working.object(root)?;
                let index = structure::unique_message(object, &[COMMENT_STORAGE_MESSAGE_TYPE])?;
                let payload =
                    engine::comment_storage_with_body(&object.messages[index].data, body, limits)
                        .map_err(from_engine)?;
                replace(context.working, root, index, payload, None, None)?;
                changed = true;
            }
            if !changed {
                return Ok(None);
            }
        },
        Staged::RemoveHighlight(position) | Staged::RemoveComment(position) => {
            let current = match change {
                Staged::RemoveComment(_) => pick(&comments, *position)?,
                _ => pick(&highlights, *position)?,
            };
            let (index, payload) = storage_payload(context.working, storage)?;
            let rewrite = ranged::rewrite(
                &payload,
                Table::Highlight,
                &RangedChange::Remove {
                    object: current.object(),
                },
                limits,
            )
            .map_err(from_ranged)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            replace(
                context.working,
                storage,
                index,
                rewrite.into_payload(),
                added,
                removed,
            )?;
            context.forget(current.object())?;
            context.forget(current.root().storage())?;
            for reply in current.replies() {
                context.forget(reply.storage())?;
            }
            let mut authors: Vec<u64> = current.authors().collect();
            authors.sort_unstable();
            authors.dedup();
            for author in authors {
                context.remove_author_if_unused(author)?;
            }
        },
        Staged::AddReply(comment, body) => {
            let current = pick(&comments, *comment)?;
            let author = context.ensure_author()?.map(|(author, _)| author);
            let reply = *context
                .working
                .allocate(1)?
                .first()
                .ok_or(Error::InvalidSource)?;
            context.insert_comment_storage(&member, reply, body, author)?;
            let root = current.root().storage();
            let object = context.working.object(root)?;
            let index = structure::unique_message(object, &[COMMENT_STORAGE_MESSAGE_TYPE])?;
            let payload = engine::comment_storage_with_reply(
                &object.messages[index].data,
                ReplyChange::Add(reply),
                limits,
            )
            .map_err(from_engine)?;
            replace(context.working, root, index, payload, Some(reply), None)?;
            context.add_author_reference(&member, author)?;
        },
        Staged::UpdateReply(comment, reply, body) => {
            let current = pick(&comments, *comment)?;
            let node = current.replies().get(*reply).ok_or(Error::Refused(
                Refusal::AnnotationNotFound { position: *reply },
            ))?;
            if node.body() == body {
                return Ok(None);
            }
            let object = context.working.object(node.storage())?;
            let index = structure::unique_message(object, &[COMMENT_STORAGE_MESSAGE_TYPE])?;
            let payload =
                engine::comment_storage_with_body(&object.messages[index].data, body, limits)
                    .map_err(from_engine)?;
            replace(context.working, node.storage(), index, payload, None, None)?;
        },
        Staged::RemoveReply(comment, reply) => {
            let current = pick(&comments, *comment)?;
            let node = current
                .replies()
                .get(*reply)
                .ok_or(Error::Refused(Refusal::AnnotationNotFound {
                    position: *reply,
                }))?
                .clone();
            let author = node.metadata().author().map(AuthorId::get);
            let root = current.root().storage();
            let object = context.working.object(root)?;
            let index = structure::unique_message(object, &[COMMENT_STORAGE_MESSAGE_TYPE])?;
            let payload = engine::comment_storage_with_reply(
                &object.messages[index].data,
                ReplyChange::Remove(node.storage()),
                limits,
            )
            .map_err(from_engine)?;
            replace(
                context.working,
                root,
                index,
                payload,
                None,
                Some(node.storage()),
            )?;
            context.forget(node.storage())?;
            if let Some(author) = author {
                context.remove_author_if_unused(author)?;
            }
        },
    }
    let transition = context.transition;
    Ok(Some(working.publish(&transition)?))
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

fn read(
    package: &Package,
    sheet: usize,
    selector: DrawableSelector,
) -> Result<(Vec<Annotation>, Vec<Annotation>), Error> {
    let storage = storage_of(package, sheet, selector)?;
    let working = Working::new(package)?;
    Ok(split(annotations_of(&working, storage)?))
}

impl Package {
    /// The highlights of the text a sheet text box or shape owns, in text
    /// order.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn sheet_drawable_highlights<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<Highlight>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        Ok(read(self, sheet, drawable.into())?
            .0
            .into_iter()
            .map(|item| Highlight {
                range: item.range(),
            })
            .collect())
    }

    /// The comments of the text a sheet text box or shape owns, in text
    /// order, with their replies.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn sheet_drawable_comments<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Vec<Comment>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        Ok(read(self, sheet, drawable.into())?
            .1
            .into_iter()
            .map(|item| Comment {
                range: item.range(),
                body: item.root().body().to_owned(),
                replies: item
                    .replies()
                    .iter()
                    .map(|node| node.body().to_owned())
                    .collect(),
            })
            .collect())
    }

    /// Start an immutable highlight and comment edit of the text of one
    /// sheet text box or shape.
    ///
    /// # Errors
    ///
    /// Selector errors, [`Refusal::NoText`], source errors.
    pub fn edit_sheet_drawable_annotations<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        drawable: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let sheet = resolve_sheet(self, sheet.into())?;
        let selector = drawable.into();
        storage_of(self, sheet, selector)?;
        Ok(Edit {
            source: self,
            sheet,
            selector,
            changes: Vec::new(),
        })
    }

    /// Apply an annotation patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_drawable_annotations(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
