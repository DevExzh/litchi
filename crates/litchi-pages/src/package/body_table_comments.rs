//! Selector-first Pages body-table cell comment lifecycle.
//!
//! The semantic surface uses only a body-table selector, a checked cell
//! position, and reply ordinals.  Native table-list keys, storage identities,
//! UUID bindings, tile routes, and archive metadata stay inside the package
//! adapter and the shared native root-comment graph engine.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::ExactArtifacts};
use litchi_iwa_common::wire::WireView;
use litchi_iwa_common::{decode_varint_from_bytes, table::coordinate::At};
use litchi_iwa_core::ArchiveObject;
use thiserror::Error;

use super::{Package, PackageError, table_cells, table_lock};
use crate::ReplySelector;
use crate::selector::BodyTableSelector;
use litchi_iwa_common::table::read::{Comment, CommentReply};

type CellPosition = At;

const TILE_MESSAGE_TYPE: u32 = 6_002;
const LIST_MESSAGE_TYPE: u32 = 6_005;
const COMMENT_LIST_TYPE: u64 = 10;
const TILE_ROWS_FIELD: u32 = 5;
const ROW_INDEX_FIELD: u32 = 1;
const ROW_CELL_COUNT_FIELD: u32 = 2;
const ROW_STORAGE_FIELD: u32 = 6;
const ROW_OFFSETS_FIELD: u32 = 7;
const ROW_WIDE_FIELD: u32 = 8;
const LIST_TYPE_FIELD: u32 = 1;
const LIST_NEXT_ID_FIELD: u32 = 2;
const LIST_ENTRY_FIELD: u32 = 3;
const ENTRY_KEY_FIELD: u32 = 1;
const ENTRY_REFCOUNT_FIELD: u32 = 2;
const ENTRY_STORAGE_FIELD: u32 = 10;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
pub(crate) const MAX_COMMENT_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Finite resources charged by a focused Pages table-cell comment operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyTableCommentLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate output bytes.
    OutputBytes,
    /// ZIP entries inspected.
    Entries,
    /// Bytes retained by one ZIP entry.
    EntryBytes,
    /// Aggregate bytes retained by ZIP entries.
    TotalEntryBytes,
    /// ZIP names and structural metadata bytes.
    PackageBytes,
    /// Native payload bytes inspected.
    PayloadBytes,
    /// Aggregate native payload bytes inspected.
    TotalPayloadBytes,
    /// Native payload objects inspected.
    PayloadObjects,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native payload items inspected.
    PayloadItems,
    /// Native object and data references inspected.
    PayloadReferences,
    /// Wire fields inspected.
    WireFields,
    /// Wire bytes inspected.
    WireBytes,
    /// Wire output bytes produced by a focused codec.
    WireOutputBytes,
    /// Wire nesting inspected.
    WireNesting,
    /// Aggregate wire work.
    WireWork,
    /// Native references inspected.
    References,
    /// Replies inspected.
    Replies,
    /// Semantic text bytes retained.
    TextBytes,
    /// Bounded scratch bytes.
    ScratchBytes,
    /// Logical allocations requested.
    Allocations,
    /// Candidate bytes retained by the codec.
    RetainedBytes,
}

impl fmt::Display for BodyTableCommentLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireFields => "wire fields",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::References => "references",
            Self::Replies => "replies",
            Self::TextBytes => "text bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
        })
    }
}

/// Failure from a Pages body-table cell comment read or transaction.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum BodyTableCommentError {
    /// No rooted table matched the selector.
    #[error("the Pages body has no table matching the requested selector")]
    TableNotFound,
    /// More than one table matched an exact name.
    #[error("the Pages body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted selector or graph was ambiguous.
    #[error("the Pages body-table comment selector is ambiguous")]
    AmbiguousSelector,
    /// The source is outside the exact writer profile.
    #[error("the Pages package source does not support exact body-table comment editing")]
    UnsupportedSource,
    /// The selected rooted graph is malformed or unsupported.
    #[error("the selected Pages body-table comment source is invalid")]
    InvalidSource,
    /// No root comment is attached to the selected cell.
    #[error("the selected Pages body-table cell has no comment")]
    CommentNotFound,
    /// The checked reply ordinal is absent.
    #[error("the Pages body-table comment reply position {position:?} does not exist")]
    ReplyPositionNotFound { position: litchi_core::Position },
    /// A valid graph dependency is outside the focused writer profile.
    #[error("the Pages body-table comment graph has an unsupported dependency")]
    UnsupportedDependency,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Pages body-table comments {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: BodyTableCommentLimitKind,
        /// Observed amount.
        observed: u64,
        /// Maximum amount.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} units for Pages body-table comments")]
    Allocation { amount: usize },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Pages body-table comment failed semantic verification")]
    Verification,
    /// The patch was created from another exact source or stale semantic state.
    #[error("the Pages body-table comment patch does not match the exact source package")]
    PatchConflict,
    /// A lower-level focused package read failed.
    #[error("the Pages body-table comment source could not be read")]
    Read,
}

/// Semantic path carried by a table-cell comment patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyTableCommentPath {
    table: usize,
    cell: CellPosition,
}

impl BodyTableCommentPath {
    /// Return the selected zero-based body-table position.
    #[must_use]
    pub const fn table(self) -> usize {
        self.table
    }

    /// Return the selected zero-based cell coordinate.
    #[must_use]
    pub const fn cell(self) -> CellPosition {
        self.cell
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct TableCommentSnapshot {
    comment: Option<Comment>,
}

#[allow(
    dead_code,
    reason = "reply variants remain explicit host-only refusals in this slice"
)]
#[derive(Debug, Clone)]
enum Operation {
    Set(Box<str>),
    Clear,
    AddReply(Box<str>),
    SetReply { index: usize, text: Box<str> },
    RemoveReply { index: usize },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CandidateAdmission;

impl CandidateAdmission {
    const EMPTY: Self = Self;

    fn charge(self, budget: &mut ReadBudget, bytes: usize) -> Result<(), BodyTableCommentError> {
        budget.wire_work(bytes)
    }
}

#[derive(Debug, Clone, Copy)]
struct ReadBudget {
    text_bytes: usize,
    wire_work: usize,
    max_text_bytes: usize,
    max_wire_work: usize,
}

impl ReadBudget {
    fn for_package(package: &Package) -> Result<Self, BodyTableCommentError> {
        let limits = package.state.source.limits();
        let maximum = limits.max_iwa_stream_bytes();
        let source_bytes = package.source_bytes().len();
        if source_bytes > usize::try_from(limits.max_input_bytes()).unwrap_or(usize::MAX) {
            return Err(BodyTableCommentError::Limit {
                kind: BodyTableCommentLimitKind::InputBytes,
                observed: source_bytes as u64,
                maximum: limits.max_input_bytes(),
            });
        }
        Ok(Self {
            text_bytes: 0,
            wire_work: 0,
            max_text_bytes: MAX_COMMENT_TEXT_BYTES,
            max_wire_work: maximum.saturating_mul(64).max(source_bytes),
        })
    }

    fn copy_text(&mut self, bytes: usize) -> Result<(), BodyTableCommentError> {
        let observed = self
            .text_bytes
            .checked_add(bytes)
            .ok_or(BodyTableCommentError::Limit {
                kind: BodyTableCommentLimitKind::TextBytes,
                observed: u64::MAX,
                maximum: self.max_text_bytes as u64,
            })?;
        if observed > self.max_text_bytes {
            return Err(BodyTableCommentError::Limit {
                kind: BodyTableCommentLimitKind::TextBytes,
                observed: observed as u64,
                maximum: self.max_text_bytes as u64,
            });
        }
        self.text_bytes = observed;
        Ok(())
    }

    fn wire_work(&mut self, bytes: usize) -> Result<(), BodyTableCommentError> {
        let observed = self
            .wire_work
            .checked_add(bytes)
            .ok_or(BodyTableCommentError::Limit {
                kind: BodyTableCommentLimitKind::WireWork,
                observed: u64::MAX,
                maximum: self.max_wire_work as u64,
            })?;
        if observed > self.max_wire_work {
            return Err(BodyTableCommentError::Limit {
                kind: BodyTableCommentLimitKind::WireWork,
                observed: observed as u64,
                maximum: self.max_wire_work as u64,
            });
        }
        self.wire_work = observed;
        Ok(())
    }
}

#[derive(Debug, Clone)]
struct TableCommentSelection {
    component: Box<str>,
    tile_component: Box<str>,
    model_identifier: u64,
    model_message_index: usize,
    model_message_type: u32,
    tile_identifier: u64,
    tile_message_index: usize,
    tile_message_type: u32,
    tile_row: u32,
    column: u32,
    table_columns: u32,
    comment_key: Option<u32>,
    list_identifier: Option<u64>,
    root_identifier: Option<u64>,
}

/// A mutable selector-first body-table cell comment edit.
pub struct BodyTableCommentEdit<'a> {
    source: &'a Package,
    table: usize,
    cell: CellPosition,
    selection: TableCommentSelection,
    before: TableCommentSnapshot,
    budget: ReadBudget,
    operation: Option<Operation>,
}

impl fmt::Debug for BodyTableCommentEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableCommentEdit")
            .field("table", &self.table)
            .field("cell", &self.cell)
            .field("has_comment", &self.before.comment.is_some())
            .field("has_operation", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl BodyTableCommentEdit<'_> {
    /// Return the selected body-table path.
    #[must_use]
    pub const fn path(&self) -> BodyTableCommentPath {
        BodyTableCommentPath {
            table: self.table,
            cell: self.cell,
        }
    }

    /// Borrow the source comment observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&Comment> {
        self.before.comment.as_ref()
    }

    /// Stage replacement or creation of the root comment.
    pub fn set(mut self, text: impl AsRef<str>) -> Result<Self, BodyTableCommentError> {
        self.operation = Some(Operation::Set(copy_text(text.as_ref(), &mut self.budget)?));
        Ok(self)
    }

    /// Stage removal of the root comment and its reply thread.
    pub fn clear(mut self) -> Result<Self, BodyTableCommentError> {
        self.operation = Some(Operation::Clear);
        Ok(self)
    }

    /// Stage appending one direct reply.
    pub fn add_reply(mut self, text: impl AsRef<str>) -> Result<Self, BodyTableCommentError> {
        if self.before.comment.is_none() {
            return Err(BodyTableCommentError::CommentNotFound);
        }
        self.operation = Some(Operation::AddReply(copy_text(
            text.as_ref(),
            &mut self.budget,
        )?));
        Ok(self)
    }

    /// Stage replacement of one checked reply ordinal.
    pub fn set_reply(
        mut self,
        selector: impl Into<ReplySelector>,
        text: impl AsRef<str>,
    ) -> Result<Self, BodyTableCommentError> {
        let index = selector.into().as_index();
        self.ensure_reply(index)?;
        self.operation = Some(Operation::SetReply {
            index,
            text: copy_text(text.as_ref(), &mut self.budget)?,
        });
        Ok(self)
    }

    /// Stage removal of one checked reply ordinal.
    pub fn remove_reply(
        mut self,
        selector: impl Into<ReplySelector>,
    ) -> Result<Self, BodyTableCommentError> {
        let index = selector.into().as_index();
        self.ensure_reply(index)?;
        self.operation = Some(Operation::RemoveReply { index });
        Ok(self)
    }

    fn ensure_reply(&self, index: usize) -> Result<(), BodyTableCommentError> {
        let Some(comment) = self.before.comment.as_ref() else {
            return Err(BodyTableCommentError::CommentNotFound);
        };
        if comment
            .replies()
            .is_none_or(|replies| index >= replies.len())
        {
            return Err(BodyTableCommentError::ReplyPositionNotFound {
                position: litchi_core::Position::new(index),
            });
        }
        Ok(())
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        commit_edit(self)
    }
}

/// Exact-source reversible body-table comment patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyTableCommentPatch {
    artifacts: ExactArtifacts,
    path: BodyTableCommentPath,
    before: TableCommentSnapshot,
    after: TableCommentSnapshot,
    touched_components: usize,
    source_admission: CandidateAdmission,
    target_admission: CandidateAdmission,
}

impl fmt::Debug for BodyTableCommentPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyTableCommentPatch")
            .field("path", &self.path)
            .field("before_has_comment", &self.before.comment.is_some())
            .field("after_has_comment", &self.after.comment.is_some())
            .finish_non_exhaustive()
    }
}

impl BodyTableCommentPatch {
    /// Return the selected semantic path.
    #[must_use]
    pub const fn path(&self) -> BodyTableCommentPath {
        self.path
    }

    /// Borrow the source root comment.
    #[must_use]
    pub fn before(&self) -> Option<&Comment> {
        self.before.comment.as_ref()
    }

    /// Borrow the target root comment.
    #[must_use]
    pub fn after(&self) -> Option<&Comment> {
        self.after.comment.as_ref()
    }

    /// Return source fingerprint used for exact conflict checks.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return target fingerprint used for exact conflict checks.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether semantic and physical artifacts are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
            source_admission: self.target_admission,
            target_admission: self.source_admission,
        }
    }
}

/// Publication diagnostics for one table-cell comment edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyTableCommentDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl BodyTableCommentDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Return whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened result of one table-cell comment transaction.
#[must_use = "a Pages table-cell comment commit contains the verified package"]
#[derive(Debug)]
pub struct BodyTableCommentCommit {
    package: Package,
    patch: BodyTableCommentPatch,
    diagnostics: BodyTableCommentDiagnostics,
}

impl BodyTableCommentCommit {
    /// Borrow the verified package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the verified commit and return its package and exact patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, BodyTableCommentPatch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyTableCommentPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyTableCommentDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read a semantic comment from one selected body-table cell.
    pub fn body_table_cell_comment<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
    ) -> Result<Option<Comment>, BodyTableCommentError> {
        read_comment(self, selector.into(), cell)
    }

    /// Read source-ordered replies from one selected body-table cell.
    pub fn body_table_cell_comment_replies<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
    ) -> Result<Box<[CommentReply]>, BodyTableCommentError> {
        Ok(read_comment(self, selector.into(), cell)?
            .and_then(|comment| comment.replies().map(ToOwned::to_owned))
            .unwrap_or_default()
            .into())
    }

    /// Start a selector-first body-table cell comment edit.
    pub fn edit_body_table_cell_comment<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
    ) -> Result<BodyTableCommentEdit<'_>, BodyTableCommentError> {
        let selector = selector.into();
        let target = resolve_selection(self, selector, cell)?;
        let before = read_comment(self, selector, cell)?;
        let budget = ReadBudget::for_package(self).map_err(map_writer_error)?;
        Ok(BodyTableCommentEdit {
            source: self,
            table: target.table_position,
            cell,
            selection: target.selection,
            before: TableCommentSnapshot { comment: before },
            budget,
            operation: None,
        })
    }

    /// Apply an exact-source-checked reversible table-cell comment patch.
    pub fn apply_body_table_cell_comment(
        &self,
        patch: &BodyTableCommentPatch,
    ) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        if !self.state.source.source_is_exact() {
            return Err(BodyTableCommentError::UnsupportedSource);
        }
        let mut budget = ReadBudget::for_package(self).map_err(map_writer_error)?;
        let source = self.state.source.shared_source();
        if !authorizes_source(&patch.artifacts, &source, &mut budget)? {
            return Err(BodyTableCommentError::PatchConflict);
        }
        let current = read_comment(
            self,
            BodyTableSelector::index(patch.path.table),
            patch.path.cell,
        )?;
        if (TableCommentSnapshot { comment: current }) != patch.before {
            return Err(BodyTableCommentError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyTableCommentCommit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: BodyTableCommentDiagnostics::unchanged(),
            });
        }
        let target = patch.artifacts.target();
        patch
            .target_admission
            .charge(&mut budget, target.len())
            .map_err(map_writer_error)?;
        let candidate = reopen(self, target)?;
        candidate.validate().map_err(map_reopen_package_error)?;
        let after = read_comment(
            &candidate,
            BodyTableSelector::index(patch.path.table),
            patch.path.cell,
        )?;
        if (TableCommentSnapshot { comment: after }) != patch.after {
            return Err(BodyTableCommentError::Verification);
        }
        Ok(BodyTableCommentCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyTableCommentDiagnostics::published(patch.touched_components),
        })
    }

    /// Convenience root-comment set/create operation.
    pub fn set_body_table_cell_comment<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
        text: impl AsRef<str>,
    ) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        self.edit_body_table_cell_comment(selector, cell)?
            .set(text)?
            .commit()
    }

    /// Convenience root-comment removal operation.
    pub fn clear_body_table_cell_comment<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
    ) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        self.edit_body_table_cell_comment(selector, cell)?
            .clear()?
            .commit()
    }

    /// Convenience reply append operation.
    pub fn add_body_table_cell_comment_reply<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
        text: impl AsRef<str>,
    ) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        self.edit_body_table_cell_comment(selector, cell)?
            .add_reply(text)?
            .commit()
    }

    /// Convenience reply replacement operation.
    pub fn set_body_table_cell_comment_reply<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
        reply: impl Into<ReplySelector>,
        text: impl AsRef<str>,
    ) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        self.edit_body_table_cell_comment(selector, cell)?
            .set_reply(reply, text)?
            .commit()
    }

    /// Convenience reply removal operation.
    pub fn remove_body_table_cell_comment_reply<'table>(
        &self,
        selector: impl Into<BodyTableSelector<'table>>,
        cell: CellPosition,
        reply: impl Into<ReplySelector>,
    ) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
        self.edit_body_table_cell_comment(selector, cell)?
            .remove_reply(reply)?
            .commit()
    }
}

fn commit_edit(
    edit: BodyTableCommentEdit<'_>,
) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
    let source = edit.source;
    let before = edit.before.clone();
    let Some(operation) = edit.operation.clone() else {
        return Ok(noop_commit(source, edit.table, edit.cell, before));
    };
    if operation_is_noop(&before, &operation) {
        return Ok(noop_commit(source, edit.table, edit.cell, before));
    }
    if matches!(operation, Operation::Set(_) | Operation::Clear) {
        return commit_shared_root(edit);
    }
    // Direct replies remain host-only until their shared graph transition is
    // lifted; do not route them through the former Pages-local writer.
    return Err(BodyTableCommentError::UnsupportedDependency);
}

fn commit_shared_root(
    edit: BodyTableCommentEdit<'_>,
) -> Result<BodyTableCommentCommit, BodyTableCommentError> {
    let source = edit.source;
    let before = edit.before.clone();
    let operation = edit.operation.ok_or(BodyTableCommentError::Verification)?;
    let native_operation = match &operation {
        Operation::Set(text) => litchi_iwa_table::comments_native::RootOperation::Set(text.clone()),
        Operation::Clear => litchi_iwa_table::comments_native::RootOperation::Clear,
        Operation::AddReply(_) | Operation::SetReply { .. } | Operation::RemoveReply { .. } => {
            return Err(BodyTableCommentError::UnsupportedDependency);
        },
    };
    let selection = litchi_iwa_table::comments_native::Selection {
        model_component: edit.selection.component.clone(),
        tile_component: edit.selection.tile_component.clone(),
        model_identifier: edit.selection.model_identifier,
        model_message_index: edit.selection.model_message_index,
        model_message_type: edit.selection.model_message_type,
        tile_identifier: edit.selection.tile_identifier,
        tile_message_index: edit.selection.tile_message_index,
        tile_message_type: edit.selection.tile_message_type,
        tile_row: edit.selection.tile_row,
        column: edit.selection.column,
        table_columns: edit.selection.table_columns,
        comment_key: edit.selection.comment_key,
        list_identifier: edit.selection.list_identifier,
        root_identifier: edit.selection.root_identifier,
    };
    let native = litchi_iwa_table::comments_native::rewrite_root(
        source,
        &selection,
        &native_operation,
        litchi_iwa_table::comments_native::Profile::PAGES,
    )
    .map_err(map_native_fault)?;
    let target = native.target();
    let mut budget = edit.budget;
    budget.wire_work(target.len()).map_err(map_writer_error)?;
    let touched_components = native.touched_components();
    let candidate = native.into_package();
    candidate.validate().map_err(map_reopen_package_error)?;
    let after = read_comment(&candidate, BodyTableSelector::index(edit.table), edit.cell)?;
    let expected = match operation {
        Operation::Set(text) => Some(Comment::new(text.as_ref())),
        Operation::Clear => None,
        Operation::AddReply(_) | Operation::SetReply { .. } | Operation::RemoveReply { .. } => {
            return Err(BodyTableCommentError::UnsupportedDependency);
        },
    };
    if !snapshot_matches(&after, &TableCommentSnapshot { comment: expected }) {
        return Err(BodyTableCommentError::Verification);
    }
    let source_bytes = source.state.source.shared_source();
    Ok(BodyTableCommentCommit {
        package: candidate,
        patch: BodyTableCommentPatch {
            artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
            path: BodyTableCommentPath {
                table: edit.table,
                cell: edit.cell,
            },
            before,
            after: TableCommentSnapshot { comment: after },
            touched_components,
            source_admission: CandidateAdmission::EMPTY,
            target_admission: CandidateAdmission::EMPTY,
        },
        diagnostics: BodyTableCommentDiagnostics::published(touched_components),
    })
}

fn map_native_fault(fault: litchi_iwa_table::comments_native::Fault) -> BodyTableCommentError {
    use litchi_iwa_table::comments_native::{Fault, LimitKind};
    match fault {
        Fault::Unsupported | Fault::Dependency => BodyTableCommentError::UnsupportedDependency,
        Fault::Invalid => BodyTableCommentError::InvalidSource,
        Fault::Verification => BodyTableCommentError::Verification,
        Fault::Allocation { amount } => BodyTableCommentError::Allocation { amount },
        Fault::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableCommentError::Limit {
            kind: match kind {
                LimitKind::PackageBytes => BodyTableCommentLimitKind::PackageBytes,
                LimitKind::PayloadBytes => BodyTableCommentLimitKind::PayloadBytes,
                LimitKind::References => BodyTableCommentLimitKind::References,
                LimitKind::Replies => BodyTableCommentLimitKind::Replies,
                LimitKind::TextBytes => BodyTableCommentLimitKind::TextBytes,
                LimitKind::WireFields => BodyTableCommentLimitKind::WireFields,
                LimitKind::WireBytes => BodyTableCommentLimitKind::WireBytes,
                LimitKind::WireOutputBytes => BodyTableCommentLimitKind::WireOutputBytes,
                LimitKind::WireNesting => BodyTableCommentLimitKind::WireNesting,
                LimitKind::WireWork => BodyTableCommentLimitKind::WireWork,
                LimitKind::Allocations => BodyTableCommentLimitKind::Allocations,
                LimitKind::ScratchBytes => BodyTableCommentLimitKind::ScratchBytes,
                LimitKind::RetainedBytes => BodyTableCommentLimitKind::RetainedBytes,
            },
            observed,
            maximum,
        },
    }
}

fn noop_commit(
    source: &Package,
    table: usize,
    cell: CellPosition,
    before: TableCommentSnapshot,
) -> BodyTableCommentCommit {
    let bytes = source.state.source.shared_source();
    BodyTableCommentCommit {
        package: source.clone(),
        patch: BodyTableCommentPatch {
            artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
            path: BodyTableCommentPath { table, cell },
            after: before.clone(),
            before,
            touched_components: 0,
            source_admission: CandidateAdmission::EMPTY,
            target_admission: CandidateAdmission::EMPTY,
        },
        diagnostics: BodyTableCommentDiagnostics::unchanged(),
    }
}

fn operation_is_noop(before: &TableCommentSnapshot, operation: &Operation) -> bool {
    match (before.comment.as_ref(), operation) {
        (None, Operation::Clear) => true,
        (Some(comment), Operation::Set(text)) => comment.text() == text.as_ref(),
        (Some(_), Operation::AddReply(_)) => false,
        (Some(comment), Operation::SetReply { index, text }) => comment
            .replies()
            .and_then(|replies| replies.get(*index))
            .is_some_and(|reply| reply.text() == text.as_ref()),
        _ => false,
    }
}

fn snapshot_matches(actual: &Option<Comment>, expected: &TableCommentSnapshot) -> bool {
    match (actual.as_ref(), expected.comment.as_ref()) {
        (None, None) => true,
        (Some(actual), Some(expected)) => {
            actual.text() == expected.text()
                && expected
                    .timestamp()
                    .is_none_or(|timestamp| actual.timestamp() == Some(timestamp))
                && expected
                    .author()
                    .is_none_or(|author| actual.author() == Some(author))
                && expected.replies().is_none_or(|expected_replies| {
                    actual.replies().is_some_and(|actual_replies| {
                        actual_replies.len() == expected_replies.len()
                            && actual_replies.iter().zip(expected_replies).all(
                                |(actual, expected)| {
                                    actual.text() == expected.text()
                                        && expected.timestamp().is_none_or(|timestamp| {
                                            actual.timestamp() == Some(timestamp)
                                        })
                                        && expected
                                            .author()
                                            .is_none_or(|author| actual.author() == Some(author))
                                },
                            )
                    })
                })
        },
        _ => false,
    }
}

fn read_comment(
    package: &Package,
    selector: BodyTableSelector<'_>,
    cell: CellPosition,
) -> Result<Option<Comment>, BodyTableCommentError> {
    let read = package
        .body_table_cells(selector)
        .map_err(map_cells_error)?;
    let dimensions = read.table().dimensions();
    if cell.row() >= dimensions.rows() || cell.column() >= dimensions.columns() {
        return Err(BodyTableCommentError::InvalidSource);
    }
    Ok(read.get_comment(cell).cloned())
}

#[derive(Debug, Clone)]
struct ResolvedSelection {
    table_position: usize,
    selection: TableCommentSelection,
}

fn resolve_selection(
    package: &Package,
    selector: BodyTableSelector<'_>,
    cell: CellPosition,
) -> Result<ResolvedSelection, BodyTableCommentError> {
    let mut budget =
        table_lock::WireBudget::new(package.state.source.limits()).map_err(map_lock_error)?;
    let target = package
        .resolve_body_table_with_budget(selector, &mut budget)
        .map_err(map_lock_error)?;
    table_lock::validate_body_table_target(package, &target, &mut budget)
        .map_err(map_lock_error)?;
    if cell.row() >= target.table_rows || cell.column() >= target.table_columns {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let selected = table_cells::selected_body_table_storage(package, selector, &mut budget)
        .map_err(map_cells_error)?;
    let tile_index = cell.row() / selected.dimensions().tile_size();
    let tile = selected
        .tiles()
        .iter()
        .find(|tile| tile.tile_index() == tile_index)
        .copied()
        .ok_or(BodyTableCommentError::InvalidSource)?;
    let (tile_component, _tile_object_index, tile_object) = find_object(package, tile.object_id())?;
    let model_component = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .ok_or(BodyTableCommentError::InvalidSource)?;
    let tile_messages = tile_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == TILE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if tile_messages.len() != 1 {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let (tile_message_index, tile_message) = tile_messages[0];
    let key = read_tile_comment_key(
        &tile_message.data,
        cell.row() % selected.dimensions().tile_size(),
        cell.column(),
        target.table_columns,
    )?;
    let list_identifier = selected
        .data_store()
        .comment_storage_table()
        .map(|reference| reference.identifier());
    if let Some(list_identifier) = list_identifier {
        validate_comment_list(package, list_identifier)?;
    }
    let root_identifier = match (key, list_identifier) {
        (Some(key), Some(list)) => Some(read_list_entry(package, list, key)?),
        (None, _) => None,
        (Some(_), None) => return Err(BodyTableCommentError::InvalidSource),
    };
    let model_object = package
        .state
        .source
        .components()
        .get_index(target.model_component_index)
        .and_then(|component| component.archive().objects.get(target.model_object_index))
        .ok_or(BodyTableCommentError::InvalidSource)?;
    let _model_message = model_object
        .messages
        .get(target.model_message_index)
        .filter(|message| message.type_ == target.model_message_type)
        .ok_or(BodyTableCommentError::InvalidSource)?;
    Ok(ResolvedSelection {
        table_position: target.table_position,
        selection: TableCommentSelection {
            component: model_component.name().into(),
            tile_component: package
                .state
                .source
                .components()
                .get_index(tile_component)
                .ok_or(BodyTableCommentError::InvalidSource)?
                .name()
                .into(),
            model_identifier: target.model_identifier.get(),
            model_message_index: target.model_message_index,
            model_message_type: target.model_message_type,
            tile_identifier: tile.object_id(),
            tile_message_index,
            tile_message_type: tile_message.type_,
            tile_row: cell.row() % selected.dimensions().tile_size(),
            column: cell.column(),
            table_columns: target.table_columns,
            comment_key: key,
            list_identifier,
            root_identifier,
        },
    })
}

fn find_object(
    package: &Package,
    identifier: u64,
) -> Result<(usize, usize, &ArchiveObject), BodyTableCommentError> {
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        if let Some((object_index, object)) = component
            .archive()
            .objects
            .iter()
            .enumerate()
            .find(|(_, object)| object.archive_info.identifier == Some(identifier))
        {
            return Ok((component_index, object_index, object));
        }
    }
    Err(BodyTableCommentError::InvalidSource)
}

fn read_tile_comment_key(
    source: &[u8],
    row_index: u32,
    column: u32,
    table_columns: u32,
) -> Result<Option<u32>, BodyTableCommentError> {
    let view = WireView::parse(source, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|_| BodyTableCommentError::InvalidSource)?;
    let mut found = None;
    for field in view
        .fields()
        .filter(|field| field.number() == TILE_ROWS_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(BodyTableCommentError::InvalidSource);
        }
        let row = WireView::parse(field.payload(), litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_| BodyTableCommentError::InvalidSource)?;
        let row_value = unique_varint(&row, ROW_INDEX_FIELD)?;
        if row_value != u64::from(row_index) {
            continue;
        }
        if found
            .replace(read_tile_row_comment_key(
                field.payload(),
                column,
                table_columns,
            )?)
            .is_some()
        {
            return Err(BodyTableCommentError::InvalidSource);
        }
    }
    Ok(found.flatten())
}

fn read_tile_row_comment_key(
    source: &[u8],
    column: u32,
    table_columns: u32,
) -> Result<Option<u32>, BodyTableCommentError> {
    let view = WireView::parse(source, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|_| BodyTableCommentError::InvalidSource)?;
    let storage = unique_length(&view, ROW_STORAGE_FIELD)?;
    let offsets = unique_length(&view, ROW_OFFSETS_FIELD)?;
    let cell_count = usize::try_from(unique_varint(&view, ROW_CELL_COUNT_FIELD)?)
        .map_err(|_| BodyTableCommentError::InvalidSource)?;
    let wide = view
        .fields()
        .filter(|field| field.number() == ROW_WIDE_FIELD)
        .map(|field| unique_field_varint(field))
        .collect::<Result<Vec<_>, _>>()?;
    if wide.len() > 1 {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let unit = if wide.first().copied().unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    let column = usize::try_from(column).map_err(|_| BodyTableCommentError::InvalidSource)?;
    let table_columns =
        usize::try_from(table_columns).map_err(|_| BodyTableCommentError::InvalidSource)?;
    if !offsets.payload().len().is_multiple_of(2) || column >= table_columns {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let mut starts = Vec::new();
    for chunk in offsets.payload().chunks_exact(2) {
        let value = u16::from_le_bytes([chunk[0], chunk[1]]);
        starts.push(if value == u16::MAX {
            None
        } else {
            Some(
                usize::from(value)
                    .checked_mul(unit)
                    .ok_or(BodyTableCommentError::InvalidSource)?,
            )
        });
    }
    if starts
        .iter()
        .enumerate()
        .any(|(index, start)| index >= table_columns && start.is_some())
        || starts.iter().flatten().count() != cell_count
    {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let Some(start) = starts.get(column).copied().flatten() else {
        return Ok(None);
    };
    let end = starts
        .iter()
        .skip(column + 1)
        .flatten()
        .next()
        .copied()
        .unwrap_or(storage.payload().len());
    if start >= end || end > storage.payload().len() {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let cell = litchi_numbers_wire::BncCell::parse(
        &storage.payload()[start..end],
        &litchi_numbers_wire::Limits::DEFAULT,
    )
    .map_err(|_| BodyTableCommentError::InvalidSource)?;
    Ok(cell.comment_identifier())
}

fn read_list_entry(
    package: &Package,
    identifier: u64,
    key: u32,
) -> Result<u64, BodyTableCommentError> {
    let (_, _, object) = find_object(package, identifier)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == LIST_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let view = WireView::parse(&messages[0].data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|_| BodyTableCommentError::InvalidSource)?;
    let list_type = unique_varint(&view, LIST_TYPE_FIELD)?;
    if list_type != COMMENT_LIST_TYPE {
        return Err(BodyTableCommentError::UnsupportedDependency);
    }
    let mut result = None;
    for field in view
        .fields()
        .filter(|field| field.number() == LIST_ENTRY_FIELD)
    {
        let entry = WireView::parse(field.payload(), litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(|_| BodyTableCommentError::InvalidSource)?;
        if entry.fields().any(|entry_field| {
            !matches!(
                entry_field.number(),
                ENTRY_KEY_FIELD | ENTRY_REFCOUNT_FIELD | ENTRY_STORAGE_FIELD
            )
        }) {
            return Err(BodyTableCommentError::UnsupportedDependency);
        }
        let entry_key = u32::try_from(unique_varint(&entry, ENTRY_KEY_FIELD)?)
            .map_err(|_| BodyTableCommentError::InvalidSource)?;
        if entry_key != key {
            continue;
        }
        if result.is_some() {
            return Err(BodyTableCommentError::InvalidSource);
        }
        let reference_count = unique_varint(&entry, 2)?;
        if reference_count == 0 {
            return Err(BodyTableCommentError::InvalidSource);
        }
        let reference = unique_length(&entry, ENTRY_STORAGE_FIELD)?;
        result = Some(unique_varint(
            &WireView::parse(reference.payload(), litchi_iwa_common::WireLimits::DEFAULT)
                .map_err(|_| BodyTableCommentError::InvalidSource)?,
            REFERENCE_IDENTIFIER_FIELD,
        )?);
    }
    let result = result.ok_or(BodyTableCommentError::InvalidSource)?;
    if result == 0 {
        return Err(BodyTableCommentError::InvalidSource);
    }
    Ok(result)
}

fn validate_comment_list(package: &Package, identifier: u64) -> Result<(), BodyTableCommentError> {
    let (_, _, object) = find_object(package, identifier)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == LIST_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(BodyTableCommentError::InvalidSource);
    }
    let view = WireView::parse(&messages[0].data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(|_| BodyTableCommentError::InvalidSource)?;
    for field in view.fields() {
        if !matches!(
            field.number(),
            LIST_TYPE_FIELD | LIST_NEXT_ID_FIELD | LIST_ENTRY_FIELD
        ) {
            return Err(BodyTableCommentError::UnsupportedDependency);
        }
        if field.number() == LIST_ENTRY_FIELD && field.wire_type() != 2 {
            return Err(BodyTableCommentError::InvalidSource);
        }
    }
    if unique_varint(&view, LIST_TYPE_FIELD)? != COMMENT_LIST_TYPE {
        return Err(BodyTableCommentError::UnsupportedDependency);
    }
    unique_varint(&view, LIST_NEXT_ID_FIELD)?;
    Ok(())
}

fn unique_length<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<litchi_iwa_common::wire::WireFieldView<'a>, BodyTableCommentError> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 || found.replace(field).is_some() {
            return Err(BodyTableCommentError::InvalidSource);
        }
    }
    found.ok_or(BodyTableCommentError::InvalidSource)
}

fn unique_varint(view: &WireView<'_>, number: u32) -> Result<u64, BodyTableCommentError> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 0 || found.is_some() {
            return Err(BodyTableCommentError::InvalidSource);
        }
        found = Some(unique_field_varint(field)?);
    }
    found.ok_or(BodyTableCommentError::InvalidSource)
}

fn unique_field_varint(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<u64, BodyTableCommentError> {
    let (value, used) = decode_varint_from_bytes(field.payload())
        .map_err(|_| BodyTableCommentError::InvalidSource)?;
    if used != field.payload().len() {
        return Err(BodyTableCommentError::InvalidSource);
    }
    Ok(value)
}

fn copy_text(text: &str, budget: &mut ReadBudget) -> Result<Box<str>, BodyTableCommentError> {
    budget.copy_text(text.len()).map_err(map_writer_error)?;
    let mut copied = String::new();
    copied
        .try_reserve_exact(text.len())
        .map_err(|_| BodyTableCommentError::Allocation { amount: text.len() })?;
    copied.push_str(text);
    Ok(copied.into_boxed_str())
}

fn reopen(source: &Package, target: Arc<[u8]>) -> Result<Package, BodyTableCommentError> {
    let catalog = SourceCatalog::from_shared_bytes(target, source.state.source.limits())
        .map_err(map_reopen_archive_error)?;
    Package::from_source_catalog(catalog).map_err(map_reopen_package_error)
}

fn map_reopen_archive_error(error: litchi_iwa_archive::Error) -> BodyTableCommentError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableCommentError::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => BodyTableCommentLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyTableCommentLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => BodyTableCommentLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    BodyTableCommentLimitKind::PackageBytes
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    BodyTableCommentLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    BodyTableCommentLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyTableCommentLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    BodyTableCommentLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyTableCommentError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_reopen_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => BodyTableCommentError::UnsupportedSource,
        _ => BodyTableCommentError::Read,
    }
}

fn map_reopen_core_error(error: litchi_iwa_core::Error) -> BodyTableCommentError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableCommentError::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => BodyTableCommentLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyTableCommentLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    BodyTableCommentLimitKind::PayloadItems
                },
                litchi_iwa_core::LimitKind::HeaderNesting => BodyTableCommentLimitKind::WireNesting,
                _ => BodyTableCommentLimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyTableCommentError::Allocation { amount: requested }
        },
        _ => BodyTableCommentError::Read,
    }
}

fn map_reopen_package_error(error: PackageError) -> BodyTableCommentError {
    match error {
        PackageError::Archive(error) => map_reopen_archive_error(error),
        PackageError::Limit {
            observed, maximum, ..
        } => BodyTableCommentError::Limit {
            kind: BodyTableCommentLimitKind::WireWork,
            observed,
            maximum,
        },
        PackageError::PayloadLimit {
            observed,
            limit: maximum,
        }
        | PackageError::SectionNamesTooLarge {
            observed,
            limit: maximum,
        } => BodyTableCommentError::Limit {
            kind: BodyTableCommentLimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        PackageError::ObjectLimit {
            observed,
            limit: maximum,
        } => BodyTableCommentError::Limit {
            kind: BodyTableCommentLimitKind::PayloadObjects,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        PackageError::Semantic(crate::Error::TooManySections {
            actual: observed,
            limit: maximum,
        })
        | PackageError::Semantic(crate::Error::TooManyBodyStorages {
            actual: observed,
            limit: maximum,
        }) => BodyTableCommentError::Limit {
            kind: BodyTableCommentLimitKind::PayloadItems,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        PackageError::Semantic(crate::Error::TextTooLarge {
            observed,
            limit: maximum,
        }) => BodyTableCommentError::Limit {
            kind: BodyTableCommentLimitKind::PayloadBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        PackageError::Allocation { amount } => BodyTableCommentError::Allocation { amount },
        _ => BodyTableCommentError::Read,
    }
}

fn authorizes_source(
    artifacts: &ExactArtifacts,
    candidate: &Arc<[u8]>,
    budget: &mut ReadBudget,
) -> Result<bool, BodyTableCommentError> {
    budget
        .wire_work(candidate.len())
        .map_err(map_writer_error)?;
    Ok(artifacts.authorizes_source(candidate))
}

fn map_cells_error(error: table_cells::Error) -> BodyTableCommentError {
    match error {
        table_cells::Error::TableNotFound => BodyTableCommentError::TableNotFound,
        table_cells::Error::AmbiguousTableName => BodyTableCommentError::AmbiguousTableName,
        table_cells::Error::AmbiguousSelector => BodyTableCommentError::AmbiguousSelector,
        table_cells::Error::UnsupportedSource => BodyTableCommentError::UnsupportedSource,
        table_cells::Error::InvalidSource => BodyTableCommentError::InvalidSource,
        table_cells::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableCommentError::Limit {
            kind: map_cells_limit(kind),
            observed,
            maximum,
        },
        table_cells::Error::Allocation { amount } => BodyTableCommentError::Allocation { amount },
    }
}

fn map_lock_error(error: table_lock::Error) -> BodyTableCommentError {
    match error {
        table_lock::Error::Refused(table_lock::Refusal::TableNotFound) => {
            BodyTableCommentError::TableNotFound
        },
        table_lock::Error::Refused(table_lock::Refusal::AmbiguousTableName) => {
            BodyTableCommentError::AmbiguousTableName
        },
        table_lock::Error::Refused(table_lock::Refusal::AmbiguousSelector) => {
            BodyTableCommentError::AmbiguousSelector
        },
        table_lock::Error::Refused(table_lock::Refusal::UnsupportedSource) => {
            BodyTableCommentError::UnsupportedSource
        },
        table_lock::Error::InvalidSource => BodyTableCommentError::InvalidSource,
        table_lock::Error::Verification => BodyTableCommentError::Verification,
        table_lock::Error::PatchConflict => BodyTableCommentError::PatchConflict,
        table_lock::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyTableCommentError::Limit {
            kind: map_lock_limit(kind),
            observed,
            maximum,
        },
        table_lock::Error::Allocation { amount } => BodyTableCommentError::Allocation { amount },
    }
}

fn map_writer_error(error: BodyTableCommentError) -> BodyTableCommentError {
    error
}

const fn map_cells_limit(kind: table_cells::LimitKind) -> BodyTableCommentLimitKind {
    use table_cells::LimitKind as Cells;
    match kind {
        Cells::InputBytes => BodyTableCommentLimitKind::InputBytes,
        Cells::OutputBytes => BodyTableCommentLimitKind::OutputBytes,
        Cells::Entries => BodyTableCommentLimitKind::Entries,
        Cells::EntryBytes => BodyTableCommentLimitKind::EntryBytes,
        Cells::TotalEntryBytes => BodyTableCommentLimitKind::TotalEntryBytes,
        Cells::PackageBytes => BodyTableCommentLimitKind::PackageBytes,
        Cells::PayloadBytes => BodyTableCommentLimitKind::PayloadBytes,
        Cells::TotalPayloadBytes => BodyTableCommentLimitKind::TotalPayloadBytes,
        Cells::PayloadObjects => BodyTableCommentLimitKind::PayloadObjects,
        Cells::PayloadMessages => BodyTableCommentLimitKind::PayloadMessages,
        Cells::PayloadItems => BodyTableCommentLimitKind::PayloadItems,
        Cells::PayloadReferences => BodyTableCommentLimitKind::PayloadReferences,
        Cells::WireBytes => BodyTableCommentLimitKind::WireBytes,
        Cells::WireOutputBytes => BodyTableCommentLimitKind::WireOutputBytes,
        Cells::WireFields => BodyTableCommentLimitKind::WireFields,
        Cells::WireNesting => BodyTableCommentLimitKind::WireNesting,
        Cells::WireWork => BodyTableCommentLimitKind::WireWork,
    }
}

const fn map_lock_limit(kind: table_lock::LimitKind) -> BodyTableCommentLimitKind {
    use table_lock::LimitKind as Lock;
    match kind {
        Lock::InputBytes => BodyTableCommentLimitKind::InputBytes,
        Lock::OutputBytes => BodyTableCommentLimitKind::OutputBytes,
        Lock::Entries => BodyTableCommentLimitKind::Entries,
        Lock::EntryBytes => BodyTableCommentLimitKind::EntryBytes,
        Lock::TotalEntryBytes => BodyTableCommentLimitKind::TotalEntryBytes,
        Lock::PackageBytes => BodyTableCommentLimitKind::PackageBytes,
        Lock::PayloadBytes => BodyTableCommentLimitKind::PayloadBytes,
        Lock::TotalPayloadBytes => BodyTableCommentLimitKind::TotalPayloadBytes,
        Lock::PayloadObjects => BodyTableCommentLimitKind::PayloadObjects,
        Lock::PayloadMessages => BodyTableCommentLimitKind::PayloadMessages,
        Lock::PayloadItems => BodyTableCommentLimitKind::PayloadItems,
        Lock::PayloadReferences => BodyTableCommentLimitKind::PayloadReferences,
        Lock::WireBytes => BodyTableCommentLimitKind::WireBytes,
        Lock::WireFields => BodyTableCommentLimitKind::WireFields,
        Lock::WireNesting => BodyTableCommentLimitKind::WireNesting,
        Lock::WireWork => BodyTableCommentLimitKind::WireWork,
    }
}
