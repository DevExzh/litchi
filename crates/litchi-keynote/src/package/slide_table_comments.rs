//! Selector-first Keynote body-table cell comment lifecycle.
//!
//! The semantic surface uses only a body-table selector, a checked cell
//! position, and reply ordinals.  Native table-list keys, storage identities,
//! UUID bindings, tile routes, and archive metadata stay inside the package
//! adapter and its shared bounded comment writer.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::wire::WireView;
use litchi_iwa_common::{decode_varint_from_bytes, table::coordinate::At};
use litchi_iwa_core::ArchiveObject;
use thiserror::Error;

use super::{
    Package, PayloadLimitKind, ReadError, SemanticLimitKind, slide_table_cells, slide_table_core,
};
use crate::selector::SlideSelector;
use crate::slide::table::TableSelector;
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

/// Finite resources charged by a focused Keynote table-cell comment operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableCommentLimitKind {
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
    /// Slides inspected by the semantic reader.
    Slides,
    /// Native payload messages inspected.
    PayloadMessages,
    /// Native messages per object inspected.
    PayloadMessagesPerObject,
    /// Native payload items inspected.
    PayloadItems,
    /// Native metadata items inspected.
    PayloadMetadataItems,
    /// Text-storage objects inspected by the semantic reader.
    TextStorages,
    /// Rich-text fragments retained by the semantic reader.
    TextFragments,
    /// Snappy frames inspected.
    PayloadSnappyFrames,
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
    /// Addressable table rows inspected.
    WireTableRows,
    /// Addressable table columns inspected.
    WireTableColumns,
    /// Addressable table cells inspected.
    WireTableCells,
    /// Materialized sparse table cells inspected.
    WireMaterializedCells,
    /// Repeated Numbers-wire items inspected.
    WireItems,
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

impl fmt::Display for SlideTableCommentLimitKind {
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
            Self::Slides => "slides",
            Self::PayloadMessages => "payload messages",
            Self::PayloadMessagesPerObject => "payload messages per object",
            Self::PayloadItems => "payload items",
            Self::PayloadMetadataItems => "payload metadata items",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::PayloadSnappyFrames => "payload Snappy frames",
            Self::PayloadReferences => "payload references",
            Self::WireFields => "wire fields",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::WireTableRows => "wire table rows",
            Self::WireTableColumns => "wire table columns",
            Self::WireTableCells => "wire table cells",
            Self::WireMaterializedCells => "wire materialized cells",
            Self::WireItems => "wire items",
            Self::References => "references",
            Self::Replies => "replies",
            Self::TextBytes => "text bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
        })
    }
}

/// Failure from a Keynote body-table cell comment read or transaction.
#[derive(Debug, Error, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlideTableCommentError {
    /// No rooted table matched the selector.
    #[error("the Keynote body has no table matching the requested selector")]
    TableNotFound,
    /// More than one table matched an exact name.
    #[error("the Keynote body has more than one table with the requested name")]
    AmbiguousTableName,
    /// The rooted selector or graph was ambiguous.
    #[error("the Keynote body-table comment selector is ambiguous")]
    AmbiguousSelector,
    /// The source is outside the exact writer profile.
    #[error("the Keynote package source does not support exact body-table comment editing")]
    UnsupportedSource,
    /// The selected rooted graph is malformed or unsupported.
    #[error("the selected Keynote body-table comment source is invalid")]
    InvalidSource,
    /// No root comment is attached to the selected cell.
    #[error("the selected Keynote body-table cell has no comment")]
    CommentNotFound,
    /// A valid graph dependency is outside the focused writer profile.
    #[error("the Keynote body-table comment graph has an unsupported dependency")]
    UnsupportedDependency,
    /// The selected table is explicitly locked for mutation.
    #[error("the selected Keynote table is locked")]
    TableLocked,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote body-table comments {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: SlideTableCommentLimitKind,
        /// Observed amount.
        observed: u64,
        /// Maximum amount.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} units for Keynote body-table comments")]
    Allocation { amount: usize },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Keynote body-table comment failed semantic verification")]
    Verification,
    /// The patch was created from another exact source or stale semantic state.
    #[error("the Keynote body-table comment patch does not match the exact source package")]
    PatchConflict,
    /// A lower-level focused package read failed.
    #[error("the Keynote body-table comment source could not be read")]
    Read,
}

/// Semantic path carried by a table-cell comment patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideTableCommentPath {
    slide: litchi_core::Position,
    table: usize,
    cell: CellPosition,
}

impl SlideTableCommentPath {
    /// Return the selected zero-based slide position.
    #[must_use]
    pub const fn slide(self) -> litchi_core::Position {
        self.slide
    }

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

#[derive(Debug, Clone)]
enum Operation {
    Set(Box<str>),
    Clear,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CandidateAdmission;

impl CandidateAdmission {
    const EMPTY: Self = Self;

    fn charge(self, budget: &mut ReadBudget, bytes: usize) -> Result<(), SlideTableCommentError> {
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
    fn for_package(package: &Package) -> Result<Self, SlideTableCommentError> {
        let limits = package.state.options.archive();
        let maximum = limits.max_iwa_stream_bytes();
        let source_bytes = package.source_bytes().len();
        if source_bytes > usize::try_from(limits.max_input_bytes()).unwrap_or(usize::MAX) {
            return Err(SlideTableCommentError::Limit {
                kind: SlideTableCommentLimitKind::InputBytes,
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

    fn copy_text(&mut self, bytes: usize) -> Result<(), SlideTableCommentError> {
        let observed = self
            .text_bytes
            .checked_add(bytes)
            .ok_or(SlideTableCommentError::Limit {
                kind: SlideTableCommentLimitKind::TextBytes,
                observed: u64::MAX,
                maximum: self.max_text_bytes as u64,
            })?;
        if observed > self.max_text_bytes {
            return Err(SlideTableCommentError::Limit {
                kind: SlideTableCommentLimitKind::TextBytes,
                observed: observed as u64,
                maximum: self.max_text_bytes as u64,
            });
        }
        self.text_bytes = observed;
        Ok(())
    }

    fn wire_work(&mut self, bytes: usize) -> Result<(), SlideTableCommentError> {
        let observed = self
            .wire_work
            .checked_add(bytes)
            .ok_or(SlideTableCommentError::Limit {
                kind: SlideTableCommentLimitKind::WireWork,
                observed: u64::MAX,
                maximum: self.max_wire_work as u64,
            })?;
        if observed > self.max_wire_work {
            return Err(SlideTableCommentError::Limit {
                kind: SlideTableCommentLimitKind::WireWork,
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
pub struct SlideTableCommentEdit<'a> {
    source: &'a Package,
    slide: litchi_core::Position,
    table: usize,
    cell: CellPosition,
    selection: TableCommentSelection,
    before: TableCommentSnapshot,
    budget: ReadBudget,
    operation: Option<Operation>,
}

impl fmt::Debug for SlideTableCommentEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableCommentEdit")
            .field("slide", &self.slide)
            .field("table", &self.table)
            .field("cell", &self.cell)
            .field("has_comment", &self.before.comment.is_some())
            .field("has_operation", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

impl SlideTableCommentEdit<'_> {
    /// Return the selected body-table path.
    #[must_use]
    pub const fn path(&self) -> SlideTableCommentPath {
        SlideTableCommentPath {
            slide: self.slide,
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
    pub fn set(mut self, text: impl AsRef<str>) -> Result<Self, SlideTableCommentError> {
        self.operation = Some(Operation::Set(copy_text(text.as_ref(), &mut self.budget)?));
        Ok(self)
    }

    /// Stage removal of a root comment with no replies.
    ///
    /// Reply-bearing roots are refused by this root-only route.
    pub fn clear(mut self) -> Result<Self, SlideTableCommentError> {
        self.operation = Some(Operation::Clear);
        Ok(self)
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<SlideTableCommentCommit, SlideTableCommentError> {
        commit_edit(self)
    }
}

/// Exact-source reversible body-table comment patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableCommentPatch {
    artifacts: ExactArtifacts,
    path: SlideTableCommentPath,
    before: TableCommentSnapshot,
    after: TableCommentSnapshot,
    touched_components: usize,
    source_admission: CandidateAdmission,
    target_admission: CandidateAdmission,
}

impl fmt::Debug for SlideTableCommentPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableCommentPatch")
            .field("path", &self.path)
            .field("before_has_comment", &self.before.comment.is_some())
            .field("after_has_comment", &self.after.comment.is_some())
            .finish_non_exhaustive()
    }
}

impl SlideTableCommentPatch {
    /// Return the selected semantic path.
    #[must_use]
    pub const fn path(&self) -> SlideTableCommentPath {
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
pub struct SlideTableCommentDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SlideTableCommentDiagnostics {
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
#[must_use = "a Keynote table-cell comment commit contains the verified package"]
#[derive(Debug)]
pub struct SlideTableCommentCommit {
    package: Package,
    patch: SlideTableCommentPatch,
    diagnostics: SlideTableCommentDiagnostics,
}

impl SlideTableCommentCommit {
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
    pub fn into_parts(self) -> (Package, SlideTableCommentPatch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideTableCommentPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableCommentDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read a semantic comment from one selected slide-table cell.
    pub fn slide_table_cell_comment<'table>(
        &self,
        slide: impl Into<SlideSelector<'table>>,
        table: impl Into<TableSelector>,
        cell: CellPosition,
    ) -> Result<Option<Comment>, SlideTableCommentError> {
        read_comment(self, slide.into(), table.into(), cell)
    }

    /// Read source-ordered replies from one selected slide-table cell.
    pub fn slide_table_cell_comment_replies<'table>(
        &self,
        slide: impl Into<SlideSelector<'table>>,
        table: impl Into<TableSelector>,
        cell: CellPosition,
    ) -> Result<Box<[CommentReply]>, SlideTableCommentError> {
        Ok(read_comment(self, slide.into(), table.into(), cell)?
            .and_then(|comment| comment.replies().map(ToOwned::to_owned))
            .unwrap_or_default()
            .into())
    }

    /// Start a selector-first slide-table cell comment edit.
    pub fn edit_slide_table_cell_comment<'table>(
        &self,
        slide: impl Into<SlideSelector<'table>>,
        table: impl Into<TableSelector>,
        cell: CellPosition,
    ) -> Result<SlideTableCommentEdit<'_>, SlideTableCommentError> {
        let slide = slide.into();
        let table = table.into();
        let target = resolve_selection(self, slide, table, cell)?;
        let before = read_comment(self, slide, table, cell)?;
        let budget = ReadBudget::for_package(self).map_err(map_writer_error)?;
        Ok(SlideTableCommentEdit {
            source: self,
            slide: target.slide_position,
            table: target.table_position.get(),
            cell,
            selection: target.selection,
            before: TableCommentSnapshot { comment: before },
            budget,
            operation: None,
        })
    }

    /// Apply an exact-source-checked reversible table-cell comment patch.
    pub fn apply_slide_table_cell_comment(
        &self,
        patch: &SlideTableCommentPatch,
    ) -> Result<SlideTableCommentCommit, SlideTableCommentError> {
        if !matches!(&self.state.source, super::PhysicalSource::Package(source) if source.source_is_exact())
        {
            return Err(SlideTableCommentError::UnsupportedSource);
        }
        let mut budget = ReadBudget::for_package(self).map_err(map_writer_error)?;
        let source = self.state.source.shared_source();
        if !authorizes_source(&patch.artifacts, &source, &mut budget)? {
            return Err(SlideTableCommentError::PatchConflict);
        }
        let current = read_comment(
            self,
            SlideSelector::position(patch.path.slide),
            TableSelector::index(patch.path.table),
            patch.path.cell,
        )?;
        if (TableCommentSnapshot { comment: current }) != patch.before {
            return Err(SlideTableCommentError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(SlideTableCommentCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableCommentDiagnostics::unchanged(),
            });
        }
        let target = patch.artifacts.target();
        patch
            .target_admission
            .charge(&mut budget, target.len())
            .map_err(map_writer_error)?;
        let candidate = reopen(self, target)?;
        candidate.validate().map_err(map_reopen_error)?;
        let after = read_comment(
            &candidate,
            SlideSelector::position(patch.path.slide),
            TableSelector::index(patch.path.table),
            patch.path.cell,
        )?;
        if (TableCommentSnapshot { comment: after }) != patch.after {
            return Err(SlideTableCommentError::Verification);
        }
        Ok(SlideTableCommentCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideTableCommentDiagnostics::published(patch.touched_components),
        })
    }

    /// Convenience root-comment set/create operation.
    pub fn set_slide_table_cell_comment<'table>(
        &self,
        slide: impl Into<SlideSelector<'table>>,
        table: impl Into<TableSelector>,
        cell: CellPosition,
        text: impl AsRef<str>,
    ) -> Result<SlideTableCommentCommit, SlideTableCommentError> {
        self.edit_slide_table_cell_comment(slide, table, cell)?
            .set(text)?
            .commit()
    }

    /// Convenience root-comment removal operation; reply-bearing roots are refused.
    pub fn clear_slide_table_cell_comment<'table>(
        &self,
        slide: impl Into<SlideSelector<'table>>,
        table: impl Into<TableSelector>,
        cell: CellPosition,
    ) -> Result<SlideTableCommentCommit, SlideTableCommentError> {
        self.edit_slide_table_cell_comment(slide, table, cell)?
            .clear()?
            .commit()
    }
}

fn commit_edit(
    edit: SlideTableCommentEdit<'_>,
) -> Result<SlideTableCommentCommit, SlideTableCommentError> {
    let source = edit.source;
    let before = edit.before.clone();
    let Some(operation) = edit.operation.clone() else {
        return Ok(noop_commit(
            source, edit.slide, edit.table, edit.cell, before,
        ));
    };
    if operation_is_noop(&before, &operation) {
        return Ok(noop_commit(
            source, edit.slide, edit.table, edit.cell, before,
        ));
    }
    if matches!(operation, Operation::Set(_))
        && before
            .comment
            .as_ref()
            .and_then(Comment::replies)
            .is_some_and(|replies| !replies.is_empty())
    {
        // The first shared slice owns root-only threads.  Replacing the text
        // of a reply-bearing root is deferred until reply metadata and
        // topology transitions are lifted into the shared engine.
        return Err(SlideTableCommentError::UnsupportedDependency);
    }
    if matches!(operation, Operation::Set(_) | Operation::Clear) {
        return commit_shared_root(edit);
    }
    unreachable!("root-only operation set is exhaustive")
}

fn commit_shared_root(
    edit: SlideTableCommentEdit<'_>,
) -> Result<SlideTableCommentCommit, SlideTableCommentError> {
    let source = edit.source;
    let before = edit.before.clone();
    let operation = edit.operation.ok_or(SlideTableCommentError::Verification)?;
    let native_operation = match &operation {
        Operation::Set(text) => litchi_iwa_table::comments_native::RootOperation::Set(text.clone()),
        Operation::Clear => litchi_iwa_table::comments_native::RootOperation::Clear,
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
        litchi_iwa_table::comments_native::Profile::KEYNOTE,
    )
    .map_err(map_native_fault)?;
    let target = native.target();
    let mut budget = edit.budget;
    budget.wire_work(target.len()).map_err(map_writer_error)?;
    let touched_components = native.touched_components();
    let candidate = native.into_package();
    candidate.validate().map_err(map_reopen_error)?;
    let after = read_comment(
        &candidate,
        SlideSelector::position(edit.slide),
        TableSelector::index(edit.table),
        edit.cell,
    )?;
    let expected = match operation {
        Operation::Set(text) => Some(Comment::new(text.as_ref())),
        Operation::Clear => None,
    };
    if !snapshot_matches(&after, &TableCommentSnapshot { comment: expected }) {
        return Err(SlideTableCommentError::Verification);
    }
    let source_bytes = source.state.source.shared_source();
    Ok(SlideTableCommentCommit {
        package: candidate,
        patch: SlideTableCommentPatch {
            artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
            path: SlideTableCommentPath {
                slide: edit.slide,
                table: edit.table,
                cell: edit.cell,
            },
            before,
            after: TableCommentSnapshot { comment: after },
            touched_components,
            source_admission: CandidateAdmission::EMPTY,
            target_admission: CandidateAdmission::EMPTY,
        },
        diagnostics: SlideTableCommentDiagnostics::published(touched_components),
    })
}

fn map_native_fault(fault: litchi_iwa_table::comments_native::Fault) -> SlideTableCommentError {
    use litchi_iwa_table::comments_native::{Fault, LimitKind};
    match fault {
        Fault::Unsupported | Fault::Dependency => SlideTableCommentError::UnsupportedDependency,
        Fault::Invalid => SlideTableCommentError::InvalidSource,
        Fault::Verification => SlideTableCommentError::Verification,
        Fault::Allocation { amount } => SlideTableCommentError::Allocation { amount },
        Fault::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCommentError::Limit {
            kind: match kind {
                LimitKind::InputBytes => SlideTableCommentLimitKind::InputBytes,
                LimitKind::PackageBytes => SlideTableCommentLimitKind::PackageBytes,
                LimitKind::PayloadBytes => SlideTableCommentLimitKind::PayloadBytes,
                LimitKind::References => SlideTableCommentLimitKind::References,
                LimitKind::Objects => SlideTableCommentLimitKind::PayloadObjects,
                LimitKind::Messages => SlideTableCommentLimitKind::PayloadMessages,
                LimitKind::MessagesPerObject => {
                    SlideTableCommentLimitKind::PayloadMessagesPerObject
                },
                LimitKind::MetadataItems => SlideTableCommentLimitKind::PayloadMetadataItems,
                LimitKind::SnappyFrames => SlideTableCommentLimitKind::PayloadSnappyFrames,
                LimitKind::Replies => SlideTableCommentLimitKind::Replies,
                LimitKind::TextBytes => SlideTableCommentLimitKind::TextBytes,
                LimitKind::WireFields => SlideTableCommentLimitKind::WireFields,
                LimitKind::WireBytes => SlideTableCommentLimitKind::WireBytes,
                LimitKind::WireOutputBytes => SlideTableCommentLimitKind::WireOutputBytes,
                LimitKind::WireNesting => SlideTableCommentLimitKind::WireNesting,
                LimitKind::WireWork => SlideTableCommentLimitKind::WireWork,
                LimitKind::TableRows => SlideTableCommentLimitKind::WireTableRows,
                LimitKind::TableColumns => SlideTableCommentLimitKind::WireTableColumns,
                LimitKind::TableCells => SlideTableCommentLimitKind::WireTableCells,
                LimitKind::MaterializedCells => SlideTableCommentLimitKind::WireMaterializedCells,
                LimitKind::Items => SlideTableCommentLimitKind::WireItems,
                LimitKind::Allocations => SlideTableCommentLimitKind::Allocations,
                LimitKind::ScratchBytes => SlideTableCommentLimitKind::ScratchBytes,
                LimitKind::RetainedBytes => SlideTableCommentLimitKind::RetainedBytes,
            },
            observed,
            maximum,
        },
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideTableCommentError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideTableCommentError::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideTableCommentLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::Fields => SlideTableCommentLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    SlideTableCommentLimitKind::WireOutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => SlideTableCommentLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideTableCommentLimitKind::WireWork,
                litchi_iwa_common::LimitKind::TableRows => {
                    SlideTableCommentLimitKind::WireTableRows
                },
                litchi_iwa_common::LimitKind::TableColumns => {
                    SlideTableCommentLimitKind::WireTableColumns
                },
                litchi_iwa_common::LimitKind::TableCells => {
                    SlideTableCommentLimitKind::WireTableCells
                },
                litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideTableCommentLimitKind::WireMaterializedCells
                },
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideTableCommentError::Allocation { amount }
        },
        _ => SlideTableCommentError::InvalidSource,
    }
}

fn map_bnc_error(error: litchi_numbers_wire::Error) -> SlideTableCommentError {
    match error {
        litchi_numbers_wire::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SlideTableCommentError::Limit {
            kind: match kind {
                litchi_numbers_wire::LimitKind::InputBytes => SlideTableCommentLimitKind::WireBytes,
                litchi_numbers_wire::LimitKind::OutputBytes => {
                    SlideTableCommentLimitKind::WireOutputBytes
                },
                litchi_numbers_wire::LimitKind::Fields => SlideTableCommentLimitKind::WireFields,
                litchi_numbers_wire::LimitKind::TailBytes => {
                    SlideTableCommentLimitKind::PayloadBytes
                },
                litchi_numbers_wire::LimitKind::Items => SlideTableCommentLimitKind::WireItems,
                _ => return SlideTableCommentError::InvalidSource,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_numbers_wire::Error::Allocation { requested } => {
            SlideTableCommentError::Allocation { amount: requested }
        },
        _ => SlideTableCommentError::InvalidSource,
    }
}

fn noop_commit(
    source: &Package,
    slide: litchi_core::Position,
    table: usize,
    cell: CellPosition,
    before: TableCommentSnapshot,
) -> SlideTableCommentCommit {
    let bytes = source.state.source.shared_source();
    SlideTableCommentCommit {
        package: source.snapshot(),
        patch: SlideTableCommentPatch {
            artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
            path: SlideTableCommentPath { slide, table, cell },
            after: before.clone(),
            before,
            touched_components: 0,
            source_admission: CandidateAdmission::EMPTY,
            target_admission: CandidateAdmission::EMPTY,
        },
        diagnostics: SlideTableCommentDiagnostics::unchanged(),
    }
}

fn operation_is_noop(before: &TableCommentSnapshot, operation: &Operation) -> bool {
    match (before.comment.as_ref(), operation) {
        (None, Operation::Clear) => true,
        (Some(comment), Operation::Set(text)) => comment.text() == text.as_ref(),
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
    slide: SlideSelector<'_>,
    table: TableSelector,
    cell: CellPosition,
) -> Result<Option<Comment>, SlideTableCommentError> {
    let read = package
        .slide_table_cells(slide, table)
        .map_err(map_cells_error)?;
    let dimensions = read.table().dimensions();
    if cell.row() >= dimensions.rows() || cell.column() >= dimensions.columns() {
        return Err(SlideTableCommentError::InvalidSource);
    }
    Ok(read.get_comment(cell).cloned())
}

#[derive(Debug, Clone)]
struct ResolvedSelection {
    slide_position: litchi_core::Position,
    table_position: litchi_core::Position,
    selection: TableCommentSelection,
}

fn resolve_selection(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    cell: CellPosition,
) -> Result<ResolvedSelection, SlideTableCommentError> {
    let mut budget = slide_table_core::Budget::new(package).map_err(map_core_error)?;
    let target = slide_table_core::select_table(package, slide, table, &mut budget)
        .map_err(map_core_error)?;
    if target.locked {
        return Err(SlideTableCommentError::TableLocked);
    }
    if cell.row() >= target.rows || cell.column() >= target.columns {
        return Err(SlideTableCommentError::InvalidSource);
    }
    let selected =
        slide_table_cells::selected_slide_table_storage(package, slide, table, &mut budget)
            .map_err(map_core_error)?;
    let tile_index = cell.row() / selected.dimensions().tile_size();
    let tile = selected
        .tiles()
        .iter()
        .find(|tile| tile.tile_index() == tile_index)
        .copied()
        .ok_or(SlideTableCommentError::InvalidSource)?;
    let (tile_component, tile_object) =
        slide_table_core::object_with_component(package, tile.object_id())
            .map_err(map_core_error)?;
    let tile_messages = tile_object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == TILE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if tile_messages.len() != 1 {
        return Err(SlideTableCommentError::InvalidSource);
    }
    let (tile_message_index, tile_message) = tile_messages[0];
    let key = read_tile_comment_key(
        &tile_message.data,
        cell.row() % selected.dimensions().tile_size(),
        cell.column(),
        target.columns,
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
        (Some(_), None) => return Err(SlideTableCommentError::InvalidSource),
    };
    let model_object =
        slide_table_core::object_at(package, &target.model).map_err(map_core_error)?;
    let model_message = model_object
        .messages
        .get(target.model_message_index)
        .ok_or(SlideTableCommentError::InvalidSource)?;
    Ok(ResolvedSelection {
        slide_position: target.slide_position,
        table_position: target.table_position,
        selection: TableCommentSelection {
            component: target.model.component.as_ref().into(),
            tile_component: tile_component.into(),
            model_identifier: target.model.identifier,
            model_message_index: target.model_message_index,
            model_message_type: model_message.type_,
            tile_identifier: tile.object_id(),
            tile_message_index,
            tile_message_type: tile_message.type_,
            tile_row: cell.row() % selected.dimensions().tile_size(),
            column: cell.column(),
            table_columns: target.columns,
            comment_key: key,
            list_identifier,
            root_identifier,
        },
    })
}

fn find_object(
    package: &Package,
    identifier: u64,
) -> Result<(usize, usize, &ArchiveObject), SlideTableCommentError> {
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
    Err(SlideTableCommentError::InvalidSource)
}

fn read_tile_comment_key(
    source: &[u8],
    row_index: u32,
    column: u32,
    table_columns: u32,
) -> Result<Option<u32>, SlideTableCommentError> {
    let view =
        WireView::parse(source, litchi_iwa_common::WireLimits::DEFAULT).map_err(map_wire_error)?;
    let mut found = None;
    for field in view
        .fields()
        .filter(|field| field.number() == TILE_ROWS_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(SlideTableCommentError::InvalidSource);
        }
        let row = WireView::parse(field.payload(), litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(map_wire_error)?;
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
            return Err(SlideTableCommentError::InvalidSource);
        }
    }
    Ok(found.flatten())
}

fn read_tile_row_comment_key(
    source: &[u8],
    column: u32,
    table_columns: u32,
) -> Result<Option<u32>, SlideTableCommentError> {
    let view =
        WireView::parse(source, litchi_iwa_common::WireLimits::DEFAULT).map_err(map_wire_error)?;
    let storage = unique_length(&view, ROW_STORAGE_FIELD)?;
    let offsets = unique_length(&view, ROW_OFFSETS_FIELD)?;
    let cell_count = usize::try_from(unique_varint(&view, ROW_CELL_COUNT_FIELD)?)
        .map_err(|_| SlideTableCommentError::InvalidSource)?;
    let wide = view
        .fields()
        .filter(|field| field.number() == ROW_WIDE_FIELD)
        .map(|field| unique_field_varint(field))
        .collect::<Result<Vec<_>, _>>()?;
    if wide.len() > 1 {
        return Err(SlideTableCommentError::InvalidSource);
    }
    let unit = if wide.first().copied().unwrap_or(0) != 0 {
        4
    } else {
        1
    };
    let column = usize::try_from(column).map_err(|_| SlideTableCommentError::InvalidSource)?;
    let table_columns =
        usize::try_from(table_columns).map_err(|_| SlideTableCommentError::InvalidSource)?;
    if !offsets.payload().len().is_multiple_of(2) || column >= table_columns {
        return Err(SlideTableCommentError::InvalidSource);
    }
    let offset_count = offsets.payload().len() / 2;
    if offset_count > litchi_iwa_common::WireLimits::MAX_FIELDS {
        return Err(SlideTableCommentError::Limit {
            kind: SlideTableCommentLimitKind::WireFields,
            observed: offset_count as u64,
            maximum: litchi_iwa_common::WireLimits::MAX_FIELDS as u64,
        });
    }
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(offset_count)
        .map_err(|_| SlideTableCommentError::Allocation {
            amount: offset_count,
        })?;
    for chunk in offsets.payload().chunks_exact(2) {
        let value = u16::from_le_bytes([chunk[0], chunk[1]]);
        starts.push(if value == u16::MAX {
            None
        } else {
            Some(
                usize::from(value)
                    .checked_mul(unit)
                    .ok_or(SlideTableCommentError::InvalidSource)?,
            )
        });
    }
    if starts
        .iter()
        .enumerate()
        .any(|(index, start)| index >= table_columns && start.is_some())
        || starts.iter().flatten().count() != cell_count
    {
        return Err(SlideTableCommentError::InvalidSource);
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
        return Err(SlideTableCommentError::InvalidSource);
    }
    let cell = litchi_numbers_wire::BncCell::parse(
        &storage.payload()[start..end],
        &litchi_numbers_wire::Limits::DEFAULT,
    )
    .map_err(map_bnc_error)?;
    Ok(cell.comment_identifier())
}

fn read_list_entry(
    package: &Package,
    identifier: u64,
    key: u32,
) -> Result<u64, SlideTableCommentError> {
    let (_, _, object) = find_object(package, identifier)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == LIST_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(SlideTableCommentError::InvalidSource);
    }
    let view = WireView::parse(&messages[0].data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    let list_type = unique_varint(&view, LIST_TYPE_FIELD)?;
    if list_type != COMMENT_LIST_TYPE {
        return Err(SlideTableCommentError::UnsupportedDependency);
    }
    let mut result = None;
    for field in view
        .fields()
        .filter(|field| field.number() == LIST_ENTRY_FIELD)
    {
        let entry = WireView::parse(field.payload(), litchi_iwa_common::WireLimits::DEFAULT)
            .map_err(map_wire_error)?;
        if entry.fields().any(|entry_field| {
            !matches!(
                entry_field.number(),
                ENTRY_KEY_FIELD | ENTRY_REFCOUNT_FIELD | ENTRY_STORAGE_FIELD
            )
        }) {
            return Err(SlideTableCommentError::UnsupportedDependency);
        }
        let entry_key = u32::try_from(unique_varint(&entry, ENTRY_KEY_FIELD)?)
            .map_err(|_| SlideTableCommentError::InvalidSource)?;
        if entry_key != key {
            continue;
        }
        if result.is_some() {
            return Err(SlideTableCommentError::InvalidSource);
        }
        let reference_count = unique_varint(&entry, 2)?;
        if reference_count == 0 {
            return Err(SlideTableCommentError::InvalidSource);
        }
        let reference = unique_length(&entry, ENTRY_STORAGE_FIELD)?;
        result = Some(unique_varint(
            &WireView::parse(reference.payload(), litchi_iwa_common::WireLimits::DEFAULT)
                .map_err(map_wire_error)?,
            REFERENCE_IDENTIFIER_FIELD,
        )?);
    }
    let result = result.ok_or(SlideTableCommentError::InvalidSource)?;
    if result == 0 {
        return Err(SlideTableCommentError::InvalidSource);
    }
    Ok(result)
}

fn validate_comment_list(package: &Package, identifier: u64) -> Result<(), SlideTableCommentError> {
    let (_, _, object) = find_object(package, identifier)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == LIST_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(SlideTableCommentError::InvalidSource);
    }
    let view = WireView::parse(&messages[0].data, litchi_iwa_common::WireLimits::DEFAULT)
        .map_err(map_wire_error)?;
    for field in view.fields() {
        if !matches!(
            field.number(),
            LIST_TYPE_FIELD | LIST_NEXT_ID_FIELD | LIST_ENTRY_FIELD
        ) {
            return Err(SlideTableCommentError::UnsupportedDependency);
        }
        if field.number() == LIST_ENTRY_FIELD && field.wire_type() != 2 {
            return Err(SlideTableCommentError::InvalidSource);
        }
    }
    if unique_varint(&view, LIST_TYPE_FIELD)? != COMMENT_LIST_TYPE {
        return Err(SlideTableCommentError::UnsupportedDependency);
    }
    unique_varint(&view, LIST_NEXT_ID_FIELD)?;
    Ok(())
}

fn unique_length<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<litchi_iwa_common::wire::WireFieldView<'a>, SlideTableCommentError> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 || found.replace(field).is_some() {
            return Err(SlideTableCommentError::InvalidSource);
        }
    }
    found.ok_or(SlideTableCommentError::InvalidSource)
}

fn unique_varint(view: &WireView<'_>, number: u32) -> Result<u64, SlideTableCommentError> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 0 || found.is_some() {
            return Err(SlideTableCommentError::InvalidSource);
        }
        found = Some(unique_field_varint(field)?);
    }
    found.ok_or(SlideTableCommentError::InvalidSource)
}

fn unique_field_varint(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<u64, SlideTableCommentError> {
    let (value, used) = decode_varint_from_bytes(field.payload())
        .map_err(|_| SlideTableCommentError::InvalidSource)?;
    if used != field.payload().len() {
        return Err(SlideTableCommentError::InvalidSource);
    }
    Ok(value)
}

fn copy_text(text: &str, budget: &mut ReadBudget) -> Result<Box<str>, SlideTableCommentError> {
    budget.copy_text(text.len()).map_err(map_writer_error)?;
    let mut copied = String::new();
    copied
        .try_reserve_exact(text.len())
        .map_err(|_| SlideTableCommentError::Allocation { amount: text.len() })?;
    copied.push_str(text);
    Ok(copied.into_boxed_str())
}

fn reopen(source: &Package, target: Arc<[u8]>) -> Result<Package, SlideTableCommentError> {
    source
        .structure_reopen_shared(target)
        .map_err(map_reopen_error)
}

fn map_reopen_error(error: ReadError) -> SlideTableCommentError {
    match error {
        ReadError::Archive(error) => map_reopen_archive_error(error),
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableCommentError::Limit {
            kind: match kind {
                PayloadLimitKind::Bytes => SlideTableCommentLimitKind::PayloadBytes,
                PayloadLimitKind::Fields => SlideTableCommentLimitKind::WireFields,
                PayloadLimitKind::Nesting => SlideTableCommentLimitKind::WireNesting,
                PayloadLimitKind::Work => SlideTableCommentLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableCommentError::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => SlideTableCommentLimitKind::PayloadObjects,
                SemanticLimitKind::Slides => SlideTableCommentLimitKind::Slides,
                SemanticLimitKind::References => SlideTableCommentLimitKind::References,
                SemanticLimitKind::TextStorages => SlideTableCommentLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => SlideTableCommentLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => SlideTableCommentLimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableCommentError::Allocation { amount },
        _ => SlideTableCommentError::Read,
    }
}

fn map_reopen_archive_error(error: litchi_iwa_archive::Error) -> SlideTableCommentError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCommentError::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTableCommentLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideTableCommentLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideTableCommentLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    SlideTableCommentLimitKind::PackageBytes
                },
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => {
                    SlideTableCommentLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideTableCommentLimitKind::TotalEntryBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    SlideTableCommentLimitKind::PayloadBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideTableCommentLimitKind::TotalPayloadBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableCommentError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_reopen_core_error(error),
        litchi_iwa_archive::Error::Reassembly(_) => SlideTableCommentError::UnsupportedSource,
        _ => SlideTableCommentError::Read,
    }
}

fn map_reopen_core_error(error: litchi_iwa_core::Error) -> SlideTableCommentError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCommentError::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableCommentLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages => SlideTableCommentLimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableCommentLimitKind::PayloadMessagesPerObject
                },
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    SlideTableCommentLimitKind::PayloadBytes
                },
                litchi_iwa_core::LimitKind::ArchiveBytes => {
                    SlideTableCommentLimitKind::PackageBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => SlideTableCommentLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideTableCommentLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::HeaderMemoryBytes => {
                    SlideTableCommentLimitKind::ScratchBytes
                },
                litchi_iwa_core::LimitKind::MetadataItems => {
                    SlideTableCommentLimitKind::PayloadMetadataItems
                },
                litchi_iwa_core::LimitKind::SnappyFrames => {
                    SlideTableCommentLimitKind::PayloadSnappyFrames
                },
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableCommentError::Allocation { amount: requested }
        },
        _ => SlideTableCommentError::Read,
    }
}

fn authorizes_source(
    artifacts: &ExactArtifacts,
    candidate: &Arc<[u8]>,
    budget: &mut ReadBudget,
) -> Result<bool, SlideTableCommentError> {
    budget
        .wire_work(candidate.len())
        .map_err(map_writer_error)?;
    Ok(artifacts.authorizes_source(candidate))
}

fn map_cells_error(error: slide_table_cells::Error) -> SlideTableCommentError {
    match error {
        slide_table_cells::Error::UnsupportedSource => SlideTableCommentError::UnsupportedSource,
        slide_table_cells::Error::UnsupportedDependency
        | slide_table_cells::Error::UnsupportedTopology => {
            SlideTableCommentError::UnsupportedDependency
        },
        slide_table_cells::Error::AmbiguousSelector => SlideTableCommentError::AmbiguousSelector,
        slide_table_cells::Error::EmptySlideName
        | slide_table_cells::Error::SlideNameNotFound
        | slide_table_cells::Error::SlidePositionNotFound { .. }
        | slide_table_cells::Error::TablePositionNotFound { .. } => {
            SlideTableCommentError::TableNotFound
        },
        slide_table_cells::Error::InvalidSource => SlideTableCommentError::InvalidSource,
        slide_table_cells::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCommentError::Limit {
            kind: map_cells_limit(kind),
            observed,
            maximum,
        },
        slide_table_cells::Error::Allocation { amount } => {
            SlideTableCommentError::Allocation { amount }
        },
    }
}

fn map_core_error(error: slide_table_core::Error) -> SlideTableCommentError {
    match error {
        slide_table_core::Error::UnsupportedSource => SlideTableCommentError::UnsupportedSource,
        slide_table_core::Error::UnsupportedDependency
        | slide_table_core::Error::UnsupportedTopology => {
            SlideTableCommentError::UnsupportedDependency
        },
        slide_table_core::Error::AmbiguousSelector => SlideTableCommentError::AmbiguousSelector,
        slide_table_core::Error::EmptySlideName
        | slide_table_core::Error::SlideNameNotFound
        | slide_table_core::Error::SlidePositionNotFound(_)
        | slide_table_core::Error::TablePositionNotFound(_) => {
            SlideTableCommentError::TableNotFound
        },
        slide_table_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableCommentError::Limit {
            kind: map_core_limit(kind),
            observed,
            maximum,
        },
        slide_table_core::Error::Allocation(amount) => {
            SlideTableCommentError::Allocation { amount }
        },
        slide_table_core::Error::InvalidSource
        | slide_table_core::Error::Read
        | slide_table_core::Error::Wire
        | slide_table_core::Error::Codec
        | slide_table_core::Error::Archive => SlideTableCommentError::InvalidSource,
        slide_table_core::Error::Verification => SlideTableCommentError::Verification,
    }
}

fn map_writer_error(error: SlideTableCommentError) -> SlideTableCommentError {
    error
}

const fn map_cells_limit(kind: slide_table_cells::LimitKind) -> SlideTableCommentLimitKind {
    use slide_table_cells::LimitKind as Cells;
    match kind {
        Cells::InputBytes => SlideTableCommentLimitKind::InputBytes,
        Cells::OutputBytes => SlideTableCommentLimitKind::OutputBytes,
        Cells::Entries => SlideTableCommentLimitKind::Entries,
        Cells::EntryBytes => SlideTableCommentLimitKind::EntryBytes,
        Cells::TotalBytes => SlideTableCommentLimitKind::TotalEntryBytes,
        Cells::PayloadObjects => SlideTableCommentLimitKind::PayloadObjects,
        Cells::PayloadMessages => SlideTableCommentLimitKind::PayloadMessages,
        Cells::PayloadItems => SlideTableCommentLimitKind::PayloadItems,
        Cells::WireFields => SlideTableCommentLimitKind::WireFields,
        Cells::WireNesting => SlideTableCommentLimitKind::WireNesting,
        Cells::WireWork => SlideTableCommentLimitKind::WireWork,
        Cells::Allocations => SlideTableCommentLimitKind::Allocations,
        Cells::Retained | Cells::Scratch => SlideTableCommentLimitKind::ScratchBytes,
        Cells::Components => SlideTableCommentLimitKind::PayloadItems,
        _ => SlideTableCommentLimitKind::WireWork,
    }
}

const fn map_core_limit(kind: slide_table_core::LimitKind) -> SlideTableCommentLimitKind {
    use slide_table_core::LimitKind as Core;
    match kind {
        Core::InputBytes => SlideTableCommentLimitKind::InputBytes,
        Core::OutputBytes => SlideTableCommentLimitKind::OutputBytes,
        Core::Entries => SlideTableCommentLimitKind::Entries,
        Core::EntryBytes => SlideTableCommentLimitKind::EntryBytes,
        Core::TotalBytes => SlideTableCommentLimitKind::TotalEntryBytes,
        Core::PayloadObjects => SlideTableCommentLimitKind::PayloadObjects,
        Core::PayloadMessages => SlideTableCommentLimitKind::PayloadMessages,
        Core::References => SlideTableCommentLimitKind::PayloadReferences,
        Core::WireFields => SlideTableCommentLimitKind::WireFields,
        Core::WireNesting => SlideTableCommentLimitKind::WireNesting,
        Core::WireWork => SlideTableCommentLimitKind::WireWork,
        Core::Allocations => SlideTableCommentLimitKind::Allocations,
        Core::Retained | Core::Scratch => SlideTableCommentLimitKind::ScratchBytes,
        Core::Components => SlideTableCommentLimitKind::PayloadItems,
    }
}
