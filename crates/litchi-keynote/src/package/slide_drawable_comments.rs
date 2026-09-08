//! Selector-first Keynote direct-drawable comment transactions.
//!
//! This owner is the public semantic boundary for drawable comments.  The
//! graph and engine children retain native object identities and archive
//! locations, while this module owns source profiles, candidate reopening,
//! exact-source patches, and the immutable commit result.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The transaction keeps public selection, candidate verification, and publication together."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::ExactArtifacts};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::slide::comment::{Comment, Reply};
use crate::{DrawableKind, DrawableSelector, ReplySelector, SlideSelector};

mod engine;
mod graph;
mod metadata;

use engine::EngineOutput;
use graph::Selection;

/// Reuse the lifecycle owner’s bounded ledger for this transaction family.
/// The type remains crate-private; callers only see the format-owned limit
/// variants below.
type Budget = super::slide_media_lifecycle::LifecycleBudget;
type Error = SlideDrawableCommentError;

const MAX_COMMENT_TEXT_BYTES: usize = 64 * 1024 * 1024;
// An add-reply operation may create a cloned root, one reply, and the
// generated author object when the source has no reusable author binding.
// This is used only to bound the candidate object-index reservation; the
// engine still validates every actual object during candidate reopen.
const MAX_CREATED_OBJECTS_PER_OPERATION: usize = 3;

/// The one semantic operation staged by a drawable-comment edit.
#[derive(Debug, Clone, PartialEq, Eq)]
enum Operation {
    Set {
        text: Box<str>,
    },
    Clear,
    AddReply {
        text: Box<str>,
    },
    SetReply {
        selector: ReplySelector,
        text: Box<str>,
    },
    RemoveReply {
        selector: ReplySelector,
    },
}

/// Finite resource categories reported by a drawable-comment transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideDrawableCommentLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Aggregate media/comment bytes.
    MediaBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
    /// Bounded transaction allocations.
    Allocations,
    /// Bytes in one semantic comment text value.
    CommentBytes,
}

impl fmt::Display for SlideDrawableCommentLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::MediaBytes => "comment bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::CommentBytes => "comment bytes",
        })
    }
}

/// Content-redacted failure raised by a direct-drawable comment operation.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideDrawableCommentError {
    /// The source was not retained as an exact physical Keynote package.
    #[error("this Keynote source does not support direct-drawable comment edits")]
    UnsupportedSource,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched an exact-name selector.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// An exact-name selector matched more than one slide.
    #[error("the Keynote direct-drawable comment selector is ambiguous")]
    AmbiguousSelector,
    /// A checked slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// A checked drawable position does not exist in the selected slide.
    #[error("the selected Keynote slide has no direct drawable at position {position:?}")]
    DrawablePositionNotFound { position: Position },
    /// A checked reply position does not exist in the selected thread.
    #[error("the selected Keynote comment has no reply at position {position:?}")]
    ReplyPositionNotFound { position: Position },
    /// The selected source-order drawable has an unknown native kind.
    #[error("the selected Keynote drawable kind is unsupported")]
    UnsupportedDrawable,
    /// The selected drawable has no direct comment.
    #[error("the selected Keynote drawable has no direct comment")]
    CommentNotFound,
    /// The selected graph is outside this owner’s safe mutation profile.
    #[error("the requested Keynote direct-drawable comment graph is unsupported")]
    UnsupportedDependency,
    /// The source graph or payload is malformed or inconsistent.
    #[error("the Keynote direct-drawable comment source is invalid")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote direct-drawable comment {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its configured ceiling.
        kind: SlideDrawableCommentLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote comment transaction")]
    Allocation { amount: usize },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Keynote direct-drawable comment failed semantic verification")]
    Verification,
    /// A patch was presented to a different exact source or stale profile.
    #[error("the Keynote direct-drawable comment patch does not match the exact source package")]
    PatchConflict,
    /// A lower-level package read failed at the semantic boundary.
    #[error("the Keynote direct-drawable comment source could not be read")]
    Read,
}

impl From<super::slide_media_lifecycle::SlideMediaLifecycleError> for SlideDrawableCommentError {
    fn from(error: super::slide_media_lifecycle::SlideMediaLifecycleError) -> Self {
        map_lifecycle_error(error)
    }
}

/// A compact, public inventory item for one supported direct drawable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawableSummary {
    selector: DrawableSelector,
    kind: DrawableKind,
    has_comment: bool,
    reply_count: usize,
}

type ThreadSnapshot = graph::ThreadSnapshot;

fn read_snapshot(
    package: &Package,
    selection: &Selection,
    budget: &mut Budget,
) -> Result<ThreadSnapshot, Error> {
    match graph::read_comment_graph(package, selection, budget)? {
        Some(thread) => Ok(thread.snapshot()),
        None => Ok(ThreadSnapshot::default()),
    }
}

impl DrawableSummary {
    /// Construct an inventory item from private graph facts.
    pub(super) const fn new(
        selector: DrawableSelector,
        kind: DrawableKind,
        has_comment: bool,
        reply_count: usize,
    ) -> Self {
        Self {
            selector,
            kind,
            has_comment,
            reply_count,
        }
    }

    /// Return the checked source-order selector.
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// Return the semantic drawable kind.
    #[must_use]
    pub const fn kind(&self) -> DrawableKind {
        self.kind
    }

    /// Return whether the drawable has a direct comment.
    #[must_use]
    pub const fn has_comment(&self) -> bool {
        self.has_comment
    }

    /// Return the number of ordered direct replies, without exposing IDs.
    #[must_use]
    pub const fn reply_count(&self) -> usize {
        self.reply_count
    }
}

/// One mutable direct-drawable comment edit staged against an immutable
/// package snapshot.
pub struct SlideDrawableCommentEdit<'a> {
    source: &'a Package,
    selection: Selection,
    before: ThreadSnapshot,
    operation: Option<Operation>,
    budget: Budget,
}

impl fmt::Debug for SlideDrawableCommentEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideDrawableCommentEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("drawable_position", &self.selection.drawable_position)
            .field("kind", &self.selection.kind)
            .field("has_before_comment", &self.before.comment.is_some())
            .field("has_operation", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> SlideDrawableCommentEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        drawable_selector: impl Into<DrawableSelector>,
    ) -> Result<Self, Error> {
        let mut budget = budget_for(source)?;
        let selection = graph::select_drawable(
            source,
            slide_selector.into(),
            drawable_selector.into(),
            &mut budget,
        )?;
        let before = read_snapshot(source, &selection, &mut budget)?;
        Ok(Self {
            source,
            selection,
            before,
            operation: None,
            budget,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic drawable position.
    #[must_use]
    pub const fn drawable_position(&self) -> Position {
        self.selection.drawable_position
    }

    /// Return the selected drawable kind.
    #[must_use]
    pub const fn kind(&self) -> DrawableKind {
        self.selection.kind
    }

    /// Borrow the root comment observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&Comment> {
        self.before.comment.as_deref()
    }

    /// Stage replacement or creation of the selected root comment.
    pub fn set(mut self, text: impl AsRef<str>) -> Result<Self, Error> {
        self.operation = Some(Operation::Set {
            text: copy_text(text.as_ref(), &mut self.budget)?,
        });
        Ok(self)
    }

    /// Stage removal of the selected root comment and its direct replies.
    pub fn clear(mut self) -> Result<Self, Error> {
        self.operation = Some(Operation::Clear);
        Ok(self)
    }

    /// Stage appending one ordered reply to the selected root comment.
    pub fn add_reply(mut self, text: impl AsRef<str>) -> Result<Self, Error> {
        if self.before.comment.is_none() {
            return Err(Error::CommentNotFound);
        }
        self.operation = Some(Operation::AddReply {
            text: copy_text(text.as_ref(), &mut self.budget)?,
        });
        Ok(self)
    }

    /// Stage replacement of one checked reply ordinal.
    pub fn set_reply(
        mut self,
        selector: impl Into<ReplySelector>,
        text: impl AsRef<str>,
    ) -> Result<Self, Error> {
        let selector = selector.into();
        self.ensure_reply(selector)?;
        self.operation = Some(Operation::SetReply {
            selector,
            text: copy_text(text.as_ref(), &mut self.budget)?,
        });
        Ok(self)
    }

    /// Stage removal of one checked reply ordinal.
    pub fn remove_reply(mut self, selector: impl Into<ReplySelector>) -> Result<Self, Error> {
        let selector = selector.into();
        self.ensure_reply(selector)?;
        self.operation = Some(Operation::RemoveReply { selector });
        Ok(self)
    }

    fn ensure_reply(&self, selector: ReplySelector) -> Result<(), Error> {
        if self.before.comment.is_none() {
            return Err(Error::CommentNotFound);
        }
        if self
            .before
            .replies
            .get(selector.as_position().get())
            .is_none()
        {
            return Err(Error::ReplyPositionNotFound {
                position: selector.as_position(),
            });
        }
        Ok(())
    }

    /// Validate and publish the staged immutable candidate.
    pub fn commit(self) -> Result<SlideDrawableCommentCommit, Error> {
        let mut budget = self.budget;
        let current = graph::select_drawable(
            self.source,
            SlideSelector::position(self.selection.slide_position),
            DrawableSelector::position(self.selection.drawable_position),
            &mut budget,
        )?;
        let current_thread = read_snapshot(self.source, &current, &mut budget)?;
        if !same_selection(&self.selection, &current) || current_thread != self.before {
            return Err(Error::InvalidSource);
        }

        let operation = self.operation;
        let source_catalog = physical_catalog(self.source)?;
        let source_bytes = source_catalog.shared_source();
        if operation
            .as_ref()
            .is_none_or(|operation| operation_is_noop(operation, &self.before))
        {
            self.source.validate().map_err(map_read_error)?;
            return Ok(SlideDrawableCommentCommit {
                package: self.source.snapshot(),
                patch: SlideDrawableCommentPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position: self.selection.slide_position,
                    drawable_position: self.selection.drawable_position,
                    kind: self.selection.kind,
                    before: self.before.clone(),
                    after: self.before,
                    touched_components: 0,
                    deleted_previews: 0,
                },
                diagnostics: SlideDrawableCommentDiagnostics::unchanged(),
            });
        }
        self.source.validate().map_err(map_read_error)?;
        let operation = operation.ok_or(Error::InvalidSource)?;
        let output: EngineOutput =
            engine::execute(self.source, &self.selection, &operation, &mut budget)?;
        charge_candidate_reopen(&mut budget, self.source, output.bytes.len(), false)?;
        let target = Arc::<[u8]>::from(output.bytes);
        let candidate =
            Package::from_source_with_options(target.clone(), self.source.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let target_selection = graph::select_drawable(
            &candidate,
            SlideSelector::position(self.selection.slide_position),
            DrawableSelector::position(self.selection.drawable_position),
            &mut budget,
        )?;
        if !same_selection(&self.selection, &target_selection) {
            return Err(Error::Verification);
        }
        let target_thread = read_snapshot(&candidate, &target_selection, &mut budget)?;
        verify_operation(&self.before, &target_thread, &operation)?;
        let diagnostics = map_engine_diagnostics(&output.diagnostics);
        Ok(SlideDrawableCommentCommit {
            package: candidate,
            patch: SlideDrawableCommentPatch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                slide_position: self.selection.slide_position,
                drawable_position: self.selection.drawable_position,
                kind: self.selection.kind,
                before: self.before,
                after: target_thread,
                touched_components: diagnostics.touched_components,
                deleted_previews: diagnostics.deleted_previews,
            },
            diagnostics,
        })
    }
}

/// An exact-source-checked reversible direct-drawable comment patch.
#[derive(Clone, PartialEq)]
pub struct SlideDrawableCommentPatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    drawable_position: Position,
    kind: DrawableKind,
    before: ThreadSnapshot,
    after: ThreadSnapshot,
    touched_components: usize,
    deleted_previews: usize,
}

impl fmt::Debug for SlideDrawableCommentPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideDrawableCommentPatch")
            .field("slide_position", &self.slide_position)
            .field("drawable_position", &self.drawable_position)
            .field("kind", &self.kind)
            .field("before_has_comment", &self.before.comment.is_some())
            .field("after_has_comment", &self.after.comment.is_some())
            .finish_non_exhaustive()
    }
}

impl SlideDrawableCommentPatch {
    /// Return the selected slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected drawable position.
    #[must_use]
    pub const fn drawable_position(&self) -> Position {
        self.drawable_position
    }

    /// Return the selected drawable kind.
    #[must_use]
    pub const fn kind(&self) -> DrawableKind {
        self.kind
    }

    /// Return the compact source artifact fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the compact target artifact fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Borrow the root comment required by this patch's source profile.
    #[must_use]
    pub fn before(&self) -> Option<&Comment> {
        self.before.comment.as_deref()
    }

    /// Borrow the root comment produced by this patch.
    #[must_use]
    pub fn after(&self) -> Option<&Comment> {
        self.after.comment.as_deref()
    }

    /// Return whether this patch preserves exact bytes and semantic state.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact reversible patch from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            drawable_position: self.drawable_position,
            kind: self.kind,
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
        }
    }
}

/// Compact evidence describing one direct-drawable comment commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideDrawableCommentDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideDrawableCommentDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    /// Construct diagnostics for a changed engine candidate.
    pub(super) const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of stale root previews deleted by the engine.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the candidate was reopened and fully parsed.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable direct-drawable comment edit.
#[must_use = "a Keynote direct-drawable comment commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideDrawableCommentCommit {
    package: Package,
    patch: SlideDrawableCommentPatch,
    diagnostics: SlideDrawableCommentDiagnostics,
}

impl SlideDrawableCommentCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideDrawableCommentPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideDrawableCommentDiagnostics {
        &self.diagnostics
    }
}

fn map_engine_diagnostics(
    diagnostics: &engine::EngineDiagnostics,
) -> SlideDrawableCommentDiagnostics {
    SlideDrawableCommentDiagnostics::published(diagnostics.touched_components, 0)
}

impl Package {
    /// List supported direct drawables owned by one selected slide.
    pub fn slide_drawables<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
    ) -> Result<Box<[DrawableSummary]>, SlideDrawableCommentError> {
        let mut budget = budget_for(self)?;
        graph::inventory_drawables(self, slide_selector.into(), &mut budget)
    }

    /// Read the root comment attached directly to one selected drawable.
    pub fn slide_drawable_comment<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        drawable_selector: impl Into<DrawableSelector>,
    ) -> Result<Option<Comment>, SlideDrawableCommentError> {
        let mut budget = budget_for(self)?;
        let selection = graph::select_drawable(
            self,
            slide_selector.into(),
            drawable_selector.into(),
            &mut budget,
        )?;
        Ok(read_snapshot(self, &selection, &mut budget)?
            .comment
            .map(|comment| (*comment).clone()))
    }

    /// Read ordered direct replies attached to one selected drawable comment.
    pub fn slide_drawable_comment_replies<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        drawable_selector: impl Into<DrawableSelector>,
    ) -> Result<Box<[Reply]>, SlideDrawableCommentError> {
        let mut budget = budget_for(self)?;
        let selection = graph::select_drawable(
            self,
            slide_selector.into(),
            drawable_selector.into(),
            &mut budget,
        )?;
        Ok(read_snapshot(self, &selection, &mut budget)?.replies)
    }

    /// Start an exact-source edit of one selected direct-drawable comment.
    pub fn edit_slide_drawable_comment<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        drawable_selector: impl Into<DrawableSelector>,
    ) -> Result<SlideDrawableCommentEdit<'_>, SlideDrawableCommentError> {
        SlideDrawableCommentEdit::new(self, slide_selector, drawable_selector)
    }

    /// Apply an exact-source-checked direct-drawable comment patch.
    pub fn apply_slide_drawable_comment(
        &self,
        patch: &SlideDrawableCommentPatch,
    ) -> Result<SlideDrawableCommentCommit, SlideDrawableCommentError> {
        let source_catalog = physical_catalog(self)?;
        let source = source_catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let mut budget = budget_for(self)?;
        let current = graph::select_drawable(
            self,
            SlideSelector::position(patch.slide_position),
            DrawableSelector::position(patch.drawable_position),
            &mut budget,
        )?;
        let current_thread = read_snapshot(self, &current, &mut budget)?;
        if !same_selection_profile(patch, &current) || current_thread != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideDrawableCommentCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideDrawableCommentDiagnostics::unchanged(),
            });
        }
        charge_candidate_reopen(&mut budget, self, patch.artifacts.target().len(), true)?;
        let target = patch.artifacts.target().clone();
        let candidate = Package::from_source_with_options(target, self.state.options)
            .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_selection = graph::select_drawable(
            &candidate,
            SlideSelector::position(patch.slide_position),
            DrawableSelector::position(patch.drawable_position),
            &mut budget,
        )?;
        if !same_selection_profile(patch, &candidate_selection) {
            return Err(Error::Verification);
        }
        let candidate_thread = read_snapshot(&candidate, &candidate_selection, &mut budget)?;
        if candidate_thread != patch.after {
            return Err(Error::Verification);
        }
        Ok(SlideDrawableCommentCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideDrawableCommentDiagnostics::published(
                patch.touched_components,
                patch.deleted_previews,
            ),
        })
    }
}

fn budget_for(package: &Package) -> Result<Budget, Error> {
    Budget::for_package(package).map_err(map_lifecycle_error)
}

/// Admit the package-sized allocations owned by candidate reopening before
/// converting or parsing the target bytes.  `Package::from_source_with_options`
/// retains the target `Arc` and builds a fresh object index; the physical
/// catalog applies its own ZIP/IWA limits, while the selected graph and
/// semantic readback charge their bounded wire work through `budget` after
/// this envelope has been admitted.
fn charge_candidate_reopen(
    budget: &mut Budget,
    source: &Package,
    target_bytes: usize,
    charge_output: bool,
) -> Result<(), Error> {
    if charge_output {
        budget.charge_output(target_bytes)?;
    }
    let candidate_objects = source
        .state
        .total_objects
        .checked_add(MAX_CREATED_OBJECTS_PER_OPERATION)
        .ok_or(Error::InvalidSource)?;
    let object_index_bytes = candidate_objects
        .checked_mul(size_of::<(u64, usize, usize)>())
        .ok_or(Error::InvalidSource)?;
    let retained_bytes = target_bytes
        .checked_add(object_index_bytes)
        .ok_or(Error::InvalidSource)?;
    budget.charge_allocation_plan(retained_bytes, 2)?;
    Ok(())
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(catalog) if catalog.source_is_exact() => Ok(catalog),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

fn copy_text(text: &str, budget: &mut Budget) -> Result<Box<str>, Error> {
    if text.len() > MAX_COMMENT_TEXT_BYTES {
        return Err(Error::LimitExceeded {
            kind: SlideDrawableCommentLimitKind::CommentBytes,
            observed: text.len() as u64,
            maximum: MAX_COMMENT_TEXT_BYTES as u64,
        });
    }
    budget.charge_allocations(text.len())?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(text.len())
        .map_err(|_| Error::Allocation { amount: text.len() })?;
    owned.push_str(text);
    Ok(owned.into_boxed_str())
}

fn same_selection(left: &Selection, right: &Selection) -> bool {
    left.slide_position == right.slide_position
        && left.drawable_position == right.drawable_position
        && left.slide_identifier == right.slide_identifier
        && left.drawable_identifier == right.drawable_identifier
        && left.kind == right.kind
        && left.message_type == right.message_type
        && left.component_name == right.component_name
        && left.comment_wire_path == right.comment_wire_path
}

fn same_selection_profile(patch: &SlideDrawableCommentPatch, selection: &Selection) -> bool {
    patch.slide_position == selection.slide_position
        && patch.drawable_position == selection.drawable_position
        && patch.kind == selection.kind
}

fn operation_is_noop(operation: &Operation, before: &ThreadSnapshot) -> bool {
    match operation {
        Operation::Set { text } => before
            .comment
            .as_deref()
            .is_some_and(|comment| comment.text() == text.as_ref()),
        Operation::Clear => before.comment.is_none(),
        Operation::AddReply { .. } => false,
        Operation::SetReply { selector, text } => before
            .replies
            .get(selector.as_position().get())
            .is_some_and(|reply| reply.text() == text.as_ref()),
        Operation::RemoveReply { .. } => false,
    }
}

fn verify_operation(
    before: &ThreadSnapshot,
    after: &ThreadSnapshot,
    operation: &Operation,
) -> Result<(), Error> {
    match operation {
        Operation::Set { text } => {
            let Some(comment) = after.comment.as_deref() else {
                return Err(Error::Verification);
            };
            let metadata_preserved = before.comment.as_deref().is_none_or(|before_comment| {
                comment.timestamp() == before_comment.timestamp()
                    && comment.author() == before_comment.author()
            });
            if comment.text() != text.as_ref()
                || !metadata_preserved
                || after.replies != before.replies
            {
                return Err(Error::Verification);
            }
        },
        Operation::Clear => {
            if after.comment.is_some() || !after.replies.is_empty() {
                return Err(Error::Verification);
            }
        },
        Operation::AddReply { text } => {
            let (Some(before_comment), Some(after_comment)) =
                (before.comment.as_deref(), after.comment.as_deref())
            else {
                return Err(Error::Verification);
            };
            if before_comment != after_comment
                || after.replies.len() != before.replies.len().saturating_add(1)
                || after
                    .replies
                    .last()
                    .is_none_or(|reply| reply.text() != text.as_ref())
                || after
                    .replies
                    .iter()
                    .take(before.replies.len())
                    .ne(before.replies.iter())
            {
                return Err(Error::Verification);
            }
        },
        Operation::SetReply { selector, text } => {
            let (Some(before_comment), Some(after_comment)) =
                (before.comment.as_deref(), after.comment.as_deref())
            else {
                return Err(Error::Verification);
            };
            let index = selector.as_position().get();
            let metadata_preserved = before
                .replies
                .get(index)
                .zip(after.replies.get(index))
                .is_some_and(|(before_reply, after_reply)| {
                    before_reply.timestamp() == after_reply.timestamp()
                        && before_reply.author() == after_reply.author()
                });
            if before_comment != after_comment
                || before.replies.len() != after.replies.len()
                || after
                    .replies
                    .get(index)
                    .is_none_or(|reply| reply.text() != text.as_ref())
                || !metadata_preserved
                || !before
                    .replies
                    .iter()
                    .zip(after.replies.iter())
                    .enumerate()
                    .all(|(position, (before, after))| position == index || before == after)
            {
                return Err(Error::Verification);
            }
        },
        Operation::RemoveReply { selector } => {
            let (Some(before_comment), Some(after_comment)) =
                (before.comment.as_deref(), after.comment.as_deref())
            else {
                return Err(Error::Verification);
            };
            let index = selector.as_position().get();
            if before_comment != after_comment
                || before.replies.len() != after.replies.len().saturating_add(1)
                || after.replies.iter().enumerate().any(|(position, reply)| {
                    let source_position = if position < index {
                        position
                    } else {
                        position + 1
                    };
                    before.replies.get(source_position) != Some(reply)
                })
            {
                return Err(Error::Verification);
            }
        },
    }
    Ok(())
}

fn map_lifecycle_error(error: super::slide_media_lifecycle::SlideMediaLifecycleError) -> Error {
    use super::slide_media_lifecycle::SlideMediaLifecycleError as SourceError;
    match error {
        SourceError::UnsupportedSource => Error::UnsupportedSource,
        SourceError::EmptySlideName => Error::EmptySlideName,
        SourceError::SlideNameNotFound => Error::SlideNameNotFound,
        SourceError::AmbiguousSelector => Error::AmbiguousSelector,
        SourceError::SlidePositionNotFound { position } => {
            Error::SlidePositionNotFound { position }
        },
        SourceError::MoviePositionNotFound { position } => {
            Error::DrawablePositionNotFound { position }
        },
        SourceError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        SourceError::Allocation { amount } => Error::Allocation { amount },
        SourceError::Read => Error::Read,
        SourceError::Verification => Error::Verification,
        SourceError::PatchConflict => Error::PatchConflict,
        SourceError::KindMismatch { .. }
        | SourceError::UnsupportedComment
        | SourceError::AudioPoster
        | SourceError::InvalidSource => Error::InvalidSource,
    }
}

fn map_limit_kind(
    kind: super::slide_media_lifecycle::SlideMediaLifecycleLimitKind,
) -> SlideDrawableCommentLimitKind {
    use super::slide_media_lifecycle::SlideMediaLifecycleLimitKind as SourceKind;
    match kind {
        SourceKind::InputBytes => SlideDrawableCommentLimitKind::InputBytes,
        SourceKind::OutputBytes => SlideDrawableCommentLimitKind::OutputBytes,
        SourceKind::Entries => SlideDrawableCommentLimitKind::Entries,
        SourceKind::EntryBytes => SlideDrawableCommentLimitKind::EntryBytes,
        SourceKind::TotalBytes => SlideDrawableCommentLimitKind::TotalBytes,
        SourceKind::Slides => SlideDrawableCommentLimitKind::Slides,
        SourceKind::References => SlideDrawableCommentLimitKind::References,
        SourceKind::MediaBytes => SlideDrawableCommentLimitKind::MediaBytes,
        SourceKind::WireFields => SlideDrawableCommentLimitKind::WireFields,
        SourceKind::WireNesting => SlideDrawableCommentLimitKind::WireNesting,
        SourceKind::WireWork => SlideDrawableCommentLimitKind::WireWork,
        SourceKind::Allocations => SlideDrawableCommentLimitKind::Allocations,
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Slides => SlideDrawableCommentLimitKind::Slides,
                SemanticLimitKind::References => SlideDrawableCommentLimitKind::References,
                _ => SlideDrawableCommentLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            observed,
            maximum,
            kind,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideDrawableCommentLimitKind::OutputBytes,
                super::PayloadLimitKind::Fields => SlideDrawableCommentLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideDrawableCommentLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideDrawableCommentLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::Read,
    }
}
