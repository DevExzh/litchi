//! Selector-first direct-reply transactions for Numbers cell comments.
//!
//! This module is the package-facing half of the direct-reply owner.  It keeps
//! the public transaction vocabulary independent from the native comment
//! storage graph and gives the native transition engine one narrow call
//! boundary.  The first admitted write scope is intentionally conservative:
//! one rooted, same-member comment list with no segments or opaque inbound
//! owners.  A source outside that scope is rejected before a candidate is
//! allocated.

use std::{fmt, mem::size_of, sync::Arc};

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_core::{
    Archive, ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, SnappyStream,
};
use litchi_iwa_protos::package_metadata_codec::{
    CombinedBatch, CombinedSaveTokenBatch, RewriteOptions as MetadataRewriteOptions, SaveTokenBatch,
};
use litchi_iwa_protos::tst;
use thiserror::Error as ThisError;

use crate::{SheetSelector, TableSelector, table::CellPosition};

use super::super::comments_reply_native::{
    NativeReplyMember, NativeReplyObjectRoute, NativeReplyOperation, NativeReplyRequest,
    NativeReplyStorageIdentity, rewrite_native_comment_reply,
};
use super::super::table_cell_pop_up_menu::{
    Error as BudgetError, LimitKind as BudgetLimitKind, Path as BudgetPath, TransactionBudget,
};
use super::{
    CommentReply, Error as RootError, LimitKind as RootLimitKind, Located, Package,
    Path as RootPath, census_comment_ownership, physical_source, resolve_comment,
};

/// A checked zero-based ordinal in one comment's direct-reply list.
///
/// The ordinal is deliberately distinct from a native reply object
/// identifier.  It remains stable for one source snapshot and is resolved to
/// a native identity only inside the private package transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CommentReplyIndex(u32);

impl CommentReplyIndex {
    /// Construct an ordinal from its bounded wire-sized representation.
    #[must_use]
    pub const fn new(index: u32) -> Self {
        Self(index)
    }

    /// Return the zero-based ordinal.
    #[must_use]
    pub const fn index(self) -> u32 {
        self.0
    }

    /// Return the zero-based ordinal.
    #[must_use]
    pub const fn get(self) -> u32 {
        self.index()
    }

    /// Convert a platform-sized index without truncating it.
    pub fn try_from_usize(index: usize) -> Result<Self, CommentReplyError> {
        u32::try_from(index)
            .map(Self)
            .map_err(|_| CommentReplyError::IndexOverflow { observed: index })
    }
}

/// A content-free location associated with a direct-reply operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CommentReplyPath {
    /// The complete Numbers package.
    Package,
    /// A rooted table cell used by an append operation or collection edit.
    Cell {
        /// Zero-based sheet position.
        sheet: usize,
        /// Zero-based table position.
        table: usize,
        /// Zero-based row.
        row: usize,
        /// Zero-based column.
        column: usize,
    },
    /// One selected direct reply in a rooted table cell.
    Reply {
        /// Zero-based sheet position.
        sheet: usize,
        /// Zero-based table position.
        table: usize,
        /// Zero-based row.
        row: usize,
        /// Zero-based column.
        column: usize,
        /// Zero-based direct-reply ordinal.
        index: CommentReplyIndex,
    },
}

impl CommentReplyPath {
    const fn cell_parts(self) -> (usize, usize, usize, usize) {
        match self {
            Self::Package => (0, 0, 0, 0),
            Self::Cell {
                sheet,
                table,
                row,
                column,
            }
            | Self::Reply {
                sheet,
                table,
                row,
                column,
                ..
            } => (sheet, table, row, column),
        }
    }

    const fn as_cell(self) -> Self {
        let (sheet, table, row, column) = self.cell_parts();
        Self::Cell {
            sheet,
            table,
            row,
            column,
        }
    }

    const fn as_root(self) -> RootPath {
        let (sheet, table, row, column) = self.cell_parts();
        match self {
            Self::Package => RootPath::Package,
            Self::Cell { .. } | Self::Reply { .. } => RootPath::Cell {
                sheet,
                table,
                row,
                column,
            },
        }
    }

    const fn reply(self, index: CommentReplyIndex) -> Self {
        let (sheet, table, row, column) = self.cell_parts();
        Self::Reply {
            sheet,
            table,
            row,
            column,
            index,
        }
    }
}

/// Resource axes exposed by the direct-reply transaction boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CommentReplyLimitKind {
    /// Encoded package input bytes inspected.
    InputBytes,
    /// Candidate package output bytes emitted.
    OutputBytes,
    /// ZIP entries inspected or emitted.
    Entries,
    /// One ZIP entry's bytes.
    EntryBytes,
    /// Aggregate ZIP entry bytes.
    TotalEntryBytes,
    /// Native payload bytes.
    PayloadBytes,
    /// Native objects inspected or retained.
    PayloadObjects,
    /// Native messages inspected or retained.
    PayloadMessages,
    /// Native list items inspected or retained.
    PayloadItems,
    /// Native references inspected or emitted.
    References,
    /// Encoded native fields inspected.
    Fields,
    /// Encoded native bytes inspected.
    EncodedBytes,
    /// Encoded traversal work.
    Work,
    /// Encoded nesting depth.
    Nesting,
    /// Encoded reference payload bytes.
    ReferenceBytes,
    /// Reply text bytes retained or emitted.
    TextBytes,
    /// Temporary plan scratch bytes.
    ScratchBytes,
    /// Candidate-retained bytes.
    RetainedBytes,
    /// Compressed member bytes.
    CompressedBytes,
    /// Aggregate transaction work.
    TransactionWork,
    /// Fallible allocation units.
    Allocations,
}

/// Failure from a selector-first direct-reply operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum CommentReplyError {
    /// An A1 address was malformed.
    #[error("invalid Numbers comment-reply cell address")]
    InvalidAddress,
    /// A sheet selector did not resolve.
    #[error("the Numbers workbook has no sheet matching the comment-reply selector")]
    SheetNotFound,
    /// A table selector did not resolve.
    #[error("the selected Numbers sheet has no table matching the comment-reply selector")]
    TableNotFound,
    /// A selected cell was outside the table.
    #[error("the selected Numbers comment-reply cell is outside the table")]
    OutOfBounds { path: CommentReplyPath },
    /// The source graph or its native framing was not authoritative.
    #[error("the Numbers comment-reply source is invalid at {path:?}")]
    InvalidSource { path: CommentReplyPath },
    /// The package does not retain exact physical provenance.
    #[error("this Numbers source does not support exact comment-reply editing")]
    UnsupportedSource,
    /// The selected cell has no rooted comment.
    #[error("the selected Numbers cell has no rooted comment")]
    CommentNotFound { path: CommentReplyPath },
    /// The selected direct-reply ordinal does not exist.
    #[error("the selected Numbers direct-reply ordinal does not exist")]
    ReplyNotFound { path: CommentReplyPath },
    /// The source graph needs a broader owner than this bounded slice.
    #[error("the Numbers comment-reply source has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: CommentReplyPath },
    /// The selected table is locked against comment-reply mutation.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: CommentReplyPath },
    /// An operation-local resource ceiling was exceeded.
    #[error(
        "Numbers comment-reply {kind:?} limit exceeded: observed {observed}, maximum {maximum} at {path:?}"
    )]
    LimitExceeded {
        /// Resource axis.
        kind: CommentReplyLimitKind,
        /// Observed or requested quantity.
        observed: usize,
        /// Configured ceiling.
        maximum: usize,
        /// Semantic operation path.
        path: CommentReplyPath,
    },
    /// A bounded plan allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers comment-reply operation")]
    Allocation {
        /// Requested allocation units.
        amount: usize,
        /// Semantic operation path.
        path: CommentReplyPath,
    },
    /// Candidate reopen did not reproduce the requested state.
    #[error("the edited Numbers comment reply failed semantic verification")]
    Verification,
    /// A patch was applied to a different exact source.
    #[error("the Numbers comment-reply patch does not match the exact source package")]
    PatchConflict,
    /// A platform-sized ordinal could not be represented by the semantic
    /// index type.
    #[error("the comment-reply ordinal is too large")]
    IndexOverflow { observed: usize },
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CommentReplyDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl CommentReplyDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of changed native members/components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of removed canonical preview members.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the private candidate was reopened and verified.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The operation selected by a collection edit.
#[derive(Clone)]
enum ReplyOperation {
    Append {
        text: Arc<str>,
    },
    Set {
        index: CommentReplyIndex,
        text: Arc<str>,
    },
    Remove {
        index: CommentReplyIndex,
    },
}

/// A selector-first direct-reply edit staged against one immutable package.
pub struct CommentReplyEdit<'a> {
    source: &'a Package,
    target: Located,
    path: CommentReplyPath,
    before: Box<[CommentReply]>,
    after: Vec<CommentReply>,
    operation: Option<ReplyOperation>,
    staging_error: Option<CommentReplyError>,
}

impl fmt::Debug for CommentReplyEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CommentReplyEdit")
            .field("path", &self.path)
            .field("reply_count", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl CommentReplyEdit<'_> {
    /// Return the rooted cell path selected for this edit.
    #[must_use]
    pub const fn path(&self) -> CommentReplyPath {
        self.path
    }

    /// Return the source-ordered replies observed when this edit was opened.
    #[must_use]
    pub fn before(&self) -> &[CommentReply] {
        &self.before
    }

    /// Return the currently staged source-ordered replies.
    #[must_use]
    pub fn after(&self) -> &[CommentReply] {
        &self.after
    }

    /// Append one direct reply. Appending never selects by text.
    #[must_use]
    pub fn append(mut self, text: impl AsRef<str>) -> Self {
        match reply_from_text(self.source, text.as_ref(), self.path) {
            Ok(reply) => {
                let Ok(after) = clone_replies(&self.before, self.path) else {
                    self.staging_error = Some(CommentReplyError::Allocation {
                        amount: self.before.len(),
                        path: self.path,
                    });
                    return self;
                };
                self.after = after;
                self.after.push(reply.clone());
                self.operation = Some(ReplyOperation::Append {
                    text: reply.text.clone(),
                });
                self.staging_error = None;
            },
            Err(error) => self.staging_error = Some(error),
        }
        self
    }

    /// Replace one direct reply by ordinal.
    #[must_use]
    pub fn set(mut self, index: CommentReplyIndex, text: impl AsRef<str>) -> Self {
        let ordinal = index.index() as usize;
        if ordinal >= self.before.len() {
            self.staging_error = Some(CommentReplyError::ReplyNotFound {
                path: self.path.reply(index),
            });
            return self;
        }
        match reply_from_text(self.source, text.as_ref(), self.path.reply(index)) {
            Ok(reply) => {
                let Ok(after) = clone_replies(&self.before, self.path) else {
                    self.staging_error = Some(CommentReplyError::Allocation {
                        amount: self.before.len(),
                        path: self.path,
                    });
                    return self;
                };
                self.after = after;
                self.after[ordinal] = reply.clone();
                self.operation = Some(ReplyOperation::Set {
                    index,
                    text: reply.text.clone(),
                });
                self.staging_error = None;
            },
            Err(error) => self.staging_error = Some(error),
        }
        self
    }

    /// Remove one direct reply by ordinal.
    #[must_use]
    pub fn remove(mut self, index: CommentReplyIndex) -> Self {
        let ordinal = index.index() as usize;
        if ordinal >= self.before.len() {
            self.staging_error = Some(CommentReplyError::ReplyNotFound {
                path: self.path.reply(index),
            });
            return self;
        }
        let Ok(after) = clone_replies(&self.before, self.path) else {
            self.staging_error = Some(CommentReplyError::Allocation {
                amount: self.before.len(),
                path: self.path,
            });
            return self;
        };
        self.after = after;
        self.after.remove(ordinal);
        self.operation = Some(ReplyOperation::Remove { index });
        self.staging_error = None;
        self
    }

    /// Validate and atomically publish this edit.
    pub fn commit(self) -> Result<CommentReplyCommit, CommentReplyError> {
        if let Some(error) = self.staging_error {
            return Err(error);
        }
        commit_edit(self)
    }
}

/// A reversible exact-source direct-reply patch.
#[derive(Clone, PartialEq, Eq)]
pub struct CommentReplyPatch {
    artifacts: OwnedExactArtifacts,
    path: CommentReplyPath,
    before: Box<[CommentReply]>,
    after: Box<[CommentReply]>,
    source_cell: Arc<[u8]>,
    target_cell: Arc<[u8]>,
    source_previews: usize,
    target_previews: usize,
    touched_components: usize,
}

impl fmt::Debug for CommentReplyPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("CommentReplyPatch")
            .field("path", &self.path)
            .field("reply_count", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl CommentReplyPatch {
    /// Return the semantic path represented by this patch.
    #[must_use]
    pub const fn path(&self) -> CommentReplyPath {
        self.path
    }

    /// Return the source-ordered replies before this patch.
    #[must_use]
    pub fn before(&self) -> &[CommentReply] {
        &self.before
    }

    /// Return the source-ordered replies after this patch.
    #[must_use]
    pub fn after(&self) -> &[CommentReply] {
        &self.after
    }

    /// Return a compact source fingerprint for diagnostics only.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return a compact target fingerprint for diagnostics only.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Whether this patch retains an exact byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
            source_cell: Arc::clone(&self.target_cell),
            target_cell: Arc::clone(&self.source_cell),
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            touched_components: self.touched_components,
        }
    }
}

/// One fully validated immutable direct-reply publication.
#[must_use = "a comment-reply commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct CommentReplyCommit {
    package: Package,
    patch: CommentReplyPatch,
    diagnostics: CommentReplyDiagnostics,
}

impl CommentReplyCommit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &CommentReplyPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &CommentReplyDiagnostics {
        &self.diagnostics
    }
}

struct ReplyState {
    located: Located,
    replies: Box<[CommentReply]>,
}

fn map_root_error(error: RootError, path: CommentReplyPath) -> CommentReplyError {
    match error {
        RootError::InvalidAddress => CommentReplyError::InvalidAddress,
        RootError::SheetNotFound => CommentReplyError::SheetNotFound,
        RootError::TableNotFound => CommentReplyError::TableNotFound,
        RootError::OutOfBounds { .. } => CommentReplyError::OutOfBounds { path },
        RootError::UnsupportedSource => CommentReplyError::UnsupportedSource,
        RootError::CommentNotFound { .. } => CommentReplyError::CommentNotFound { path },
        RootError::UnsupportedDependency { .. } => {
            CommentReplyError::UnsupportedDependency { path }
        },
        RootError::LimitExceeded {
            kind,
            observed,
            maximum,
            ..
        } => CommentReplyError::LimitExceeded {
            kind: match kind {
                RootLimitKind::InputBytes => CommentReplyLimitKind::InputBytes,
                RootLimitKind::OutputBytes => CommentReplyLimitKind::OutputBytes,
                RootLimitKind::WireFields => CommentReplyLimitKind::Fields,
                RootLimitKind::WireBytes => CommentReplyLimitKind::EncodedBytes,
                RootLimitKind::WireWork => CommentReplyLimitKind::Work,
                RootLimitKind::TextBytes => CommentReplyLimitKind::TextBytes,
                RootLimitKind::References => CommentReplyLimitKind::References,
            },
            observed,
            maximum,
            path,
        },
        RootError::Allocation { amount, .. } => CommentReplyError::Allocation { amount, path },
        RootError::Verification => CommentReplyError::Verification,
        RootError::PatchConflict => CommentReplyError::PatchConflict,
        RootError::InvalidSource { .. } => CommentReplyError::InvalidSource { path },
    }
}

fn map_budget_error(error: BudgetError, path: CommentReplyPath) -> CommentReplyError {
    match error {
        BudgetError::LimitExceeded {
            kind,
            observed,
            maximum,
            ..
        } => CommentReplyError::LimitExceeded {
            kind: match kind {
                BudgetLimitKind::InputBytes => CommentReplyLimitKind::InputBytes,
                BudgetLimitKind::OutputBytes | BudgetLimitKind::PackageBytes => {
                    CommentReplyLimitKind::OutputBytes
                },
                BudgetLimitKind::Entries => CommentReplyLimitKind::Entries,
                BudgetLimitKind::EntryBytes => CommentReplyLimitKind::EntryBytes,
                BudgetLimitKind::TotalEntryBytes => CommentReplyLimitKind::TotalEntryBytes,
                BudgetLimitKind::PayloadBytes | BudgetLimitKind::TotalPayloadBytes => {
                    CommentReplyLimitKind::PayloadBytes
                },
                BudgetLimitKind::PayloadObjects => CommentReplyLimitKind::PayloadObjects,
                BudgetLimitKind::PayloadMessages => CommentReplyLimitKind::PayloadMessages,
                BudgetLimitKind::PayloadItems => CommentReplyLimitKind::PayloadItems,
                BudgetLimitKind::PayloadReferences => CommentReplyLimitKind::References,
                BudgetLimitKind::WireBytes => CommentReplyLimitKind::EncodedBytes,
                BudgetLimitKind::WireOutputBytes => CommentReplyLimitKind::OutputBytes,
                BudgetLimitKind::WireReferenceBytes => CommentReplyLimitKind::ReferenceBytes,
                BudgetLimitKind::WireTextBytes => CommentReplyLimitKind::TextBytes,
                BudgetLimitKind::WireFields => CommentReplyLimitKind::Fields,
                BudgetLimitKind::WireNesting => CommentReplyLimitKind::Nesting,
                BudgetLimitKind::WireWork => CommentReplyLimitKind::Work,
                BudgetLimitKind::ScratchBytes => CommentReplyLimitKind::ScratchBytes,
                BudgetLimitKind::RetainedBytes => CommentReplyLimitKind::RetainedBytes,
                BudgetLimitKind::Allocations => CommentReplyLimitKind::Allocations,
                BudgetLimitKind::CompressedBytes => CommentReplyLimitKind::CompressedBytes,
                BudgetLimitKind::TransactionWork => CommentReplyLimitKind::TransactionWork,
            },
            observed: usize::try_from(observed).unwrap_or(usize::MAX),
            maximum: usize::try_from(maximum).unwrap_or(usize::MAX),
            path,
        },
        BudgetError::Allocation { amount, .. } => CommentReplyError::Allocation { amount, path },
        BudgetError::UnsupportedSource => CommentReplyError::UnsupportedSource,
        BudgetError::UnsupportedDependency { .. } => {
            CommentReplyError::UnsupportedDependency { path }
        },
        BudgetError::Verification => CommentReplyError::Verification,
        BudgetError::PatchConflict => CommentReplyError::PatchConflict,
        BudgetError::SheetNotFound
        | BudgetError::TableNotFound
        | BudgetError::CellNotFound
        | BudgetError::TableLocked { .. }
        | BudgetError::InvalidSource { .. } => CommentReplyError::InvalidSource { path },
    }
}

fn metadata_rewrite_options(payload_len: usize, operation_count: usize) -> MetadataRewriteOptions {
    let source = payload_len.max(1);
    MetadataRewriteOptions::new(
        source,
        source.saturating_mul(2).max(1),
        source.saturating_mul(16).max(1),
        source.saturating_mul(64).max(1),
        64,
        source.saturating_mul(2).max(1),
        source.saturating_mul(2).max(1),
        operation_count.max(1),
    )
}

fn charge_metadata_report(
    budget: &mut TransactionBudget,
    report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
    path: CommentReplyPath,
) -> Result<(), CommentReplyError> {
    budget
        .charge_wire_bytes(report.input_bytes(), BudgetPath::Package)
        .and_then(|_| budget.charge_wire_fields(report.fields(), BudgetPath::Package))
        .and_then(|_| budget.charge_wire_work(report.work_bytes(), BudgetPath::Package))
        .and_then(|_| budget.charge_wire_nesting(report.max_depth(), BudgetPath::Package))
        .and_then(|_| budget.charge_payload_items(report.components_scanned(), BudgetPath::Package))
        .and_then(|_| {
            budget.charge_payload_references(report.references_scanned(), BudgetPath::Package)
        })
        .and_then(|_| {
            budget.charge_transaction_work(
                report.input_bytes().saturating_add(report.output_bytes()),
                BudgetPath::Package,
            )
        })
        .map_err(|error| map_budget_error(error, path))
}

fn charge_metadata_requirements(
    budget: &mut TransactionBudget,
    requirements: litchi_iwa_protos::package_metadata_codec::RewriteExecutionRequirements,
    path: CommentReplyPath,
) -> Result<(), CommentReplyError> {
    budget
        .charge_wire_fields(requirements.fields(), BudgetPath::Package)
        .and_then(|_| budget.charge_wire_work(requirements.work_bytes(), BudgetPath::Package))
        .and_then(|_| budget.charge_payload_items(requirements.components(), BudgetPath::Package))
        .and_then(|_| {
            budget.charge_payload_references(requirements.references(), BudgetPath::Package)
        })
        .and_then(|_| budget.charge_allocations(requirements.allocations(), BudgetPath::Package))
        .and_then(|_| {
            budget.charge_retained_bytes(requirements.retained_bytes(), BudgetPath::Package)
        })
        .and_then(|_| {
            budget.charge_scratch_bytes(requirements.scratch_bytes(), BudgetPath::Package)
        })
        .and_then(|_| budget.charge_output(requirements.output_bytes(), BudgetPath::Package))
        .and_then(|_| {
            budget.charge_transaction_work(requirements.output_bytes(), BudgetPath::Package)
        })
        .map_err(|error| map_budget_error(error, path))
}

fn map_metadata_error(
    error: super::super::comments_metadata::MetadataError,
    path: CommentReplyPath,
) -> CommentReplyError {
    match error.kind {
        super::super::comments_metadata::FailureKind::Unsupported => {
            CommentReplyError::UnsupportedDependency { path }
        },
        super::super::comments_metadata::FailureKind::Limit => CommentReplyError::LimitExceeded {
            kind: CommentReplyLimitKind::Work,
            observed: error.observed,
            maximum: error.maximum,
            path,
        },
        super::super::comments_metadata::FailureKind::Allocation => CommentReplyError::Allocation {
            amount: error.allocation,
            path,
        },
        super::super::comments_metadata::FailureKind::MissingRoute
        | super::super::comments_metadata::FailureKind::AmbiguousRoute
        | super::super::comments_metadata::FailureKind::InvalidSource
        | super::super::comments_metadata::FailureKind::Conflict
        | super::super::comments_metadata::FailureKind::VersionedOwnership => {
            CommentReplyError::InvalidSource { path }
        },
    }
}

fn map_native_error(
    error: super::super::comments_reply_native::NativeReplyError,
    path: CommentReplyPath,
) -> CommentReplyError {
    match error {
        super::super::comments_reply_native::NativeReplyError::UnsupportedDependency => {
            CommentReplyError::UnsupportedDependency { path }
        },
        super::super::comments_reply_native::NativeReplyError::Allocation => {
            CommentReplyError::Allocation { amount: 1, path }
        },
        super::super::comments_reply_native::NativeReplyError::Limit => {
            CommentReplyError::LimitExceeded {
                kind: CommentReplyLimitKind::Work,
                observed: 1,
                maximum: 0,
                path,
            }
        },
        super::super::comments_reply_native::NativeReplyError::InvalidSource
        | super::super::comments_reply_native::NativeReplyError::Codec
        | super::super::comments_reply_native::NativeReplyError::Archive => {
            CommentReplyError::InvalidSource { path }
        },
    }
}

fn reply_from_text(
    source: &Package,
    text: &str,
    path: CommentReplyPath,
) -> Result<CommentReply, CommentReplyError> {
    let maximum = source.state.options.semantic().max_output_text_bytes();
    if text.len() > maximum {
        return Err(CommentReplyError::LimitExceeded {
            kind: CommentReplyLimitKind::TextBytes,
            observed: text.len(),
            maximum,
            path,
        });
    }
    let mut retained = String::new();
    retained
        .try_reserve_exact(text.len())
        .map_err(|_| CommentReplyError::Allocation {
            amount: text.len(),
            path,
        })?;
    retained.push_str(text);
    Ok(CommentReply {
        text: Arc::from(retained.into_boxed_str()),
    })
}

fn clone_replies(
    source: &[CommentReply],
    path: CommentReplyPath,
) -> Result<Vec<CommentReply>, CommentReplyError> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(source.len())
        .map_err(|_| CommentReplyError::Allocation {
            amount: source.len(),
            path,
        })?;
    cloned.extend(source.iter().cloned());
    Ok(cloned)
}

fn resolve_reply_state<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    position: CellPosition,
) -> Result<ReplyState, CommentReplyError> {
    let located = resolve_comment(source, sheet, table, position).map_err(|error| {
        let path = match error {
            RootError::OutOfBounds { path }
            | RootError::InvalidSource { path }
            | RootError::CommentNotFound { path }
            | RootError::UnsupportedDependency { path }
            | RootError::LimitExceeded { path, .. }
            | RootError::Allocation { path, .. } => root_to_reply_path(path),
            _ => CommentReplyPath::Package,
        };
        map_root_error(error, path)
    })?;
    let path = root_to_reply_path(located.target.path);
    if located.comment.is_none() {
        return Err(CommentReplyError::CommentNotFound { path });
    }
    let replies = source
        .table_cell_comment_replies(
            SheetSelector::index(path.cell_parts().0),
            TableSelector::index(path.cell_parts().1),
            position,
        )
        .map_err(|error| map_root_error(error, path))?;
    if replies.len() != located.reply_ids.len() {
        return Err(CommentReplyError::InvalidSource { path });
    }
    let entry = located
        .entry
        .as_ref()
        .ok_or(CommentReplyError::InvalidSource { path })?;
    if !entry.owner.supports_text_rewrite() {
        return Err(CommentReplyError::UnsupportedDependency { path });
    }
    Ok(ReplyState { located, replies })
}

fn root_to_reply_path(path: RootPath) -> CommentReplyPath {
    match path {
        RootPath::Package => CommentReplyPath::Package,
        RootPath::Cell {
            sheet,
            table,
            row,
            column,
        } => CommentReplyPath::Cell {
            sheet,
            table,
            row,
            column,
        },
    }
}

fn no_op_commit(
    source: &Package,
    path: CommentReplyPath,
    before: &[CommentReply],
    after: &[CommentReply],
) -> Result<CommentReplyCommit, CommentReplyError> {
    let catalog = physical_source(source).map_err(|error| map_root_error(error, path))?;
    let owner = catalog.__source_owner();
    let before = before.to_vec().into_boxed_slice();
    let after = after.to_vec().into_boxed_slice();
    Ok(CommentReplyCommit {
        package: source.snapshot(),
        patch: CommentReplyPatch {
            artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
            path,
            before,
            after,
            source_cell: Arc::from([]),
            target_cell: Arc::from([]),
            source_previews: 0,
            target_previews: 0,
            touched_components: 0,
        },
        diagnostics: CommentReplyDiagnostics::unchanged(),
    })
}

struct CrossMemberReferenceProbe<'a> {
    identifiers: &'a [u64],
    found: bool,
}

impl ArchiveReferenceVisitor for CrossMemberReferenceProbe<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object
            && self.identifiers.contains(&occurrence.referenced_identifier)
        {
            self.found = true;
        }
        Ok(())
    }
}

/// A reply write is currently a single-member transition.  Reject any
/// recognized BNC/list ownership or ArchiveInfo edge in another current
/// member before native candidates are allocated.  The strict archive
/// policy also makes opaque metadata fail closed rather than allowing an
/// unknown inbound owner to survive a root/reply cull.
fn reject_cross_member_references(
    source: &Package,
    census: &super::CommentOwnershipCensus,
    selected_component_index: usize,
    comment_key: u32,
    identifiers: &[u64],
    limits: litchi_iwa_core::Limits,
    path: CommentReplyPath,
) -> Result<(), CommentReplyError> {
    if census
        .cell_comment_owners
        .iter()
        .any(|(component, key)| *component != selected_component_index && *key == comment_key)
        || census.list_entries.iter().any(|entry| {
            entry.component_index != selected_component_index
                && entry.list_type == tst::table_data_list::ListType::CommentStorage as i32
                && (entry.key == comment_key
                    || entry
                        .storage_id
                        .is_some_and(|identifier| identifiers.contains(&identifier)))
        })
    {
        return Err(CommentReplyError::UnsupportedDependency { path });
    }

    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        if component_index == selected_component_index {
            continue;
        }
        for object in &component.archive().objects {
            let mut probe = CrossMemberReferenceProbe {
                identifiers,
                found: false,
            };
            object
                .inspect_references_with_policy_and_limits(
                    &mut probe,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    limits,
                )
                .map_err(|_| CommentReplyError::UnsupportedDependency { path })?;
            if probe.found {
                return Err(CommentReplyError::UnsupportedDependency { path });
            }
        }
    }
    Ok(())
}

fn commit_edit(edit: CommentReplyEdit<'_>) -> Result<CommentReplyCommit, CommentReplyError> {
    let path = edit.path;
    if edit.before.as_ref() == edit.after.as_slice() {
        return no_op_commit(edit.source, path, &edit.before, &edit.after);
    }
    if edit.target.target.native.locked == crate::table::lock::State::Locked {
        return Err(CommentReplyError::TableLocked { path });
    }
    // The reply transaction owns one ledger from source ingress through the
    // private candidate reopen.  Start it before resolving the physical
    // catalog so every subsequent census and codec stage shares the same
    // residual ceilings.
    let mut budget = TransactionBudget::new(edit.source);
    let catalog = physical_source(edit.source).map_err(|error| map_root_error(error, path))?;
    if !catalog.source_is_exact() {
        return Err(CommentReplyError::UnsupportedSource);
    }
    budget
        .charge_package_source(catalog, BudgetPath::Package)
        .map_err(|error| map_budget_error(error, path))?;
    let catalog_bytes = catalog.source_bytes().len();
    let catalog_objects = catalog
        .components()
        .iter()
        .map(|component| component.archive().objects.len())
        .sum::<usize>();
    let catalog_messages = catalog
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .map(|object| object.messages.len())
        .sum::<usize>();
    budget
        .charge_payload_objects(catalog_objects, BudgetPath::Package)
        .and_then(|_| budget.charge_payload_messages(catalog_messages, BudgetPath::Package))
        .and_then(|_| {
            budget.charge_allocations(
                catalog.components().len().saturating_add(4),
                BudgetPath::Package,
            )
        })
        .and_then(|_| budget.charge_scratch_bytes(catalog_bytes, BudgetPath::Package))
        .and_then(|_| budget.charge_transaction_work(catalog_bytes, BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let entry = edit
        .target
        .entry
        .as_ref()
        .ok_or(CommentReplyError::InvalidSource { path })?;
    if !matches!(entry.owner, super::EntryOwner::Root)
        || edit
            .target
            .entry
            .as_ref()
            .is_some_and(|entry| entry.storage_occurrences != 1)
    {
        return Err(CommentReplyError::UnsupportedDependency { path });
    }
    // The complete global graph census is deliberately performed before any
    // native candidate allocation. It validates direct leaves, duplicate IDs,
    // storage ownership, and known inbound list edges. The native transition
    // engine repeats the same census with its BNC/refcount view before it
    // publishes a member edit.
    let census = census_comment_ownership(edit.source, path.as_root())
        .map_err(|error| map_root_error(error, path))?;
    let metadata = crate::package::comments_metadata::strict_source(edit.source)
        .map_err(|error| map_metadata_error(error, path))?;
    budget
        .charge_allocations(8, BudgetPath::Package)
        .and_then(|_| budget.charge_scratch_bytes(metadata.payload.len(), BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let metadata_facts = crate::package::comments_metadata::inspect(
        edit.source,
        metadata_rewrite_options(metadata.payload.len(), 2),
    )
    .map_err(|error| map_metadata_error(error, path))?;
    charge_metadata_report(&mut budget, metadata_facts.report(), path)?;
    crate::package::comments_metadata::reject_unknown_archive_metadata(
        edit.source,
        catalog.limits().archive_limits(),
    )
    .map_err(|error| map_metadata_error(error, path))?;
    budget
        .charge_transaction_work(catalog_bytes, BudgetPath::Package)
        .and_then(|_| {
            budget.charge_allocations(catalog_objects.saturating_add(1), BudgetPath::Package)
        })
        .map_err(|error| map_budget_error(error, path))?;

    // Keep the call boundary explicit while the dedicated native helper is
    // landing. This branch is intentionally fail-closed: no archive, ZIP, or
    // metadata candidate is allocated until the helper supplies a complete
    // COW/refcount/ownership transition.
    let entry = entry.route;
    let storage = edit
        .target
        .storage
        .ok_or(CommentReplyError::InvalidSource { path })?;
    let tile = edit
        .target
        .tile
        .ok_or(CommentReplyError::InvalidSource { path })?;
    let model = edit.target.target.native;
    let component_index = model.component_index;
    let component = catalog
        .components()
        .get_index(component_index)
        .ok_or(CommentReplyError::InvalidSource { path })?;
    let member = NativeReplyMember {
        archive: component.archive(),
        component_index,
        member_name: component.name(),
    };
    let route = |route: super::MessageRoute| -> Result<NativeReplyObjectRoute, CommentReplyError> {
        if route.component_index != component_index {
            return Err(CommentReplyError::UnsupportedDependency { path });
        }
        let object = component
            .archive()
            .objects
            .get(route.object_index)
            .ok_or(CommentReplyError::InvalidSource { path })?;
        let identifier = object
            .archive_info
            .identifier
            .ok_or(CommentReplyError::InvalidSource { path })?;
        Ok(NativeReplyObjectRoute {
            member_index: 0,
            identifier,
            message_index: route.message_index,
        })
    };
    let model = NativeReplyObjectRoute {
        member_index: 0,
        identifier: model.model_identifier,
        message_index: model.message_index,
    };
    let tile = route(tile)?;
    let list = route(entry)?;
    let root = route(storage)?;
    let mut replies = Vec::new();
    budget
        .charge_allocations(
            edit.target.reply_ids.len().saturating_add(1),
            BudgetPath::Package,
        )
        .and_then(|_| {
            budget.charge_scratch_bytes(
                edit.target.reply_ids.len().saturating_mul(size_of::<u64>()),
                BudgetPath::Package,
            )
        })
        .map_err(|error| map_budget_error(error, path))?;
    replies
        .try_reserve_exact(edit.target.reply_ids.len())
        .map_err(|_| CommentReplyError::Allocation {
            amount: edit.target.reply_ids.len(),
            path,
        })?;
    for identifier in &edit.target.reply_ids {
        let object = component
            .archive()
            .objects
            .iter()
            .enumerate()
            .find(|(_, object)| object.archive_info.identifier == Some(*identifier))
            .ok_or(CommentReplyError::InvalidSource { path })?;
        let message_index = object
            .1
            .messages
            .iter()
            .position(|message| message.type_ == 3_056)
            .ok_or(CommentReplyError::InvalidSource { path })?;
        replies.push(NativeReplyObjectRoute {
            member_index: 0,
            identifier: *identifier,
            message_index,
        });
    }
    let comment_key = edit
        .target
        .comment_key
        .ok_or(CommentReplyError::InvalidSource { path })?;
    let mut protected_identifiers = Vec::new();
    budget
        .charge_allocations(1, BudgetPath::Package)
        .and_then(|_| {
            budget.charge_scratch_bytes(
                1usize
                    .saturating_add(replies.len())
                    .saturating_mul(size_of::<u64>()),
                BudgetPath::Package,
            )
        })
        .map_err(|error| map_budget_error(error, path))?;
    protected_identifiers
        .try_reserve_exact(1usize.saturating_add(replies.len()))
        .map_err(|_| CommentReplyError::Allocation {
            amount: 1usize.saturating_add(replies.len()),
            path,
        })?;
    protected_identifiers.push(root.identifier);
    protected_identifiers.extend(replies.iter().map(|reply| reply.identifier));
    reject_cross_member_references(
        edit.source,
        &census,
        component_index,
        comment_key,
        &protected_identifiers,
        catalog.limits().archive_limits(),
        path,
    )?;
    let fresh_count = match edit.operation.as_ref() {
        Some(ReplyOperation::Remove { .. }) => 1,
        Some(ReplyOperation::Append { .. } | ReplyOperation::Set { .. }) => 2,
        None => return no_op_commit(edit.source, path, &edit.before, &edit.after),
    };
    budget
        .charge_allocations(1, BudgetPath::Package)
        .and_then(|_| budget.charge_transaction_work(fresh_count, BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let fresh = metadata_facts
        .allocate_identifiers(fresh_count)
        .map_err(|error| map_metadata_error(error, path))?;
    let root_identity = NativeReplyStorageIdentity {
        identifier: fresh[0].identifier,
        uuid_lower: fresh[0].uuid.lower(),
        uuid_upper: fresh[0].uuid.upper(),
    };
    let new_reply = fresh.get(1).map(|identity| NativeReplyStorageIdentity {
        identifier: identity.identifier,
        uuid_lower: identity.uuid.lower(),
        uuid_upper: identity.uuid.upper(),
    });
    let operation = match edit.operation.as_ref() {
        Some(ReplyOperation::Append { text }) => NativeReplyOperation::Append {
            text: text.as_ref(),
        },
        Some(ReplyOperation::Set { index, text }) => NativeReplyOperation::Replace {
            ordinal: index.index() as usize,
            expected_reply_identifier: edit.target.reply_ids[index.index() as usize],
            text: text.as_ref(),
        },
        Some(ReplyOperation::Remove { index }) => NativeReplyOperation::Remove {
            ordinal: index.index() as usize,
            expected_reply_identifier: edit.target.reply_ids[index.index() as usize],
        },
        None => return no_op_commit(edit.source, path, &edit.before, &edit.after),
    };
    let request = NativeReplyRequest {
        members: std::slice::from_ref(&member),
        model,
        tile,
        list,
        root,
        replies: &replies,
        tile_row: u32::try_from(edit.target.target.row)
            .map_err(|_| CommentReplyError::OutOfBounds { path })?,
        tile_column: u32::try_from(edit.target.target.column)
            .map_err(|_| CommentReplyError::OutOfBounds { path })?,
        comment_key,
        operation,
        new_root: root_identity,
        new_reply,
        author_identifier: None,
        limits: catalog.limits().archive_limits(),
        path,
    };
    let native = rewrite_native_comment_reply(request, &mut budget)
        .map_err(|error| map_native_error(error, path))?;
    let component_indices = native.touched_components.clone();
    let mut additions = Vec::new();
    for (ordinal, identifier) in native.added_object_identifiers.iter().enumerate() {
        let identity = fresh
            .iter()
            .find(|item| item.identifier == *identifier)
            .copied()
            .ok_or(CommentReplyError::InvalidSource { path })?;
        let component = *component_indices
            .first()
            .ok_or(CommentReplyError::InvalidSource { path })?;
        additions.push(
            metadata_facts
                .uuid_addition(component, identity)
                .map_err(|_| CommentReplyError::UnsupportedDependency { path })?,
        );
        let _ = ordinal;
    }
    let mut removals = Vec::new();
    for identifier in &native.removed_object_identifiers {
        removals.push(
            metadata_facts
                .uuid_removal(component_index, *identifier)
                .map_err(|_| CommentReplyError::UnsupportedDependency { path })?,
        );
    }
    let selectors = metadata_facts
        .touched_selectors(&component_indices)
        .map_err(|_| CommentReplyError::UnsupportedDependency { path })?;
    let transition = CombinedBatch::new(
        metadata_facts.last_object_identifier(),
        fresh
            .last()
            .map_or(metadata_facts.last_object_identifier(), |item| {
                item.identifier
            }),
        &additions,
        &[],
        &removals,
        &[],
        &[],
    );
    budget
        .charge_allocations(8, BudgetPath::Package)
        .and_then(|_| budget.charge_scratch_bytes(metadata.payload.len(), BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let prepared = crate::package::comments_metadata::prepare_combined(
        &metadata_facts,
        CombinedSaveTokenBatch::new(transition, SaveTokenBatch::new(&selectors)),
        metadata_rewrite_options(
            metadata.payload.len(),
            additions.len().saturating_add(removals.len()),
        ),
    )
    .map_err(|error| map_metadata_error(error, path))?;
    charge_metadata_report(&mut budget, prepared.prepare_report(), path)?;
    let requirements = prepared.execution_requirements();
    charge_metadata_requirements(&mut budget, requirements, path)?;
    let metadata_limits = requirements.exact_limits();
    let metadata_payload = prepared
        .execute(metadata_limits)
        .map_err(|_| CommentReplyError::Verification)?
        .into_bytes();
    let metadata_bytes = pack_reply_metadata(
        edit.source,
        metadata.route,
        metadata_payload,
        path,
        &mut budget,
    )?;
    let physical = edit
        .source
        .state
        .components
        .physical()
        .ok_or(CommentReplyError::UnsupportedSource)?;
    let edit_count = native.member_edits.len().saturating_add(1);
    budget
        .charge_allocations(edit_count.saturating_add(2), BudgetPath::Package)
        .and_then(|_| budget.charge_transaction_work(edit_count, BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let mut compressed_members = Vec::new();
    compressed_members
        .try_reserve_exact(native.member_edits.len())
        .map_err(|_| CommentReplyError::Allocation {
            amount: native.member_edits.len(),
            path,
        })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(edit_count)
        .map_err(|_| CommentReplyError::Allocation {
            amount: edit_count,
            path,
        })?;
    for member_edit in &native.member_edits {
        let maximum_compressed =
            SnappyStream::maximum_compressed_len(member_edit.member_bytes.len())
                .map_err(|_| CommentReplyError::InvalidSource { path })?;
        budget
            .charge_compressed_bytes(maximum_compressed, BudgetPath::Package)
            .and_then(|_| budget.charge_retained_bytes(maximum_compressed, BudgetPath::Package))
            .and_then(|_| {
                budget.charge_transaction_work(
                    member_edit
                        .member_bytes
                        .len()
                        .saturating_add(maximum_compressed),
                    BudgetPath::Package,
                )
            })
            .map_err(|error| map_budget_error(error, path))?;
        let bytes = SnappyStream::compress(&member_edit.member_bytes)
            .map_err(|_| CommentReplyError::Verification)?;
        if bytes.len() > maximum_compressed {
            return Err(CommentReplyError::Verification);
        }
        compressed_members.push((member_edit.member_name.clone(), bytes));
    }
    for (name, bytes) in &compressed_members {
        edits.push(EntryEdit::new(name.as_str(), bytes.as_slice()));
    }
    edits.push(EntryEdit::new(
        crate::package::metadata::ENTRY_NAME,
        metadata_bytes.as_slice(),
    ));
    let previews: [&str; 0] = [];
    let reassembly_scratch = catalog_bytes
        .checked_add(metadata_bytes.len())
        .and_then(|total| {
            compressed_members
                .iter()
                .try_fold(total, |total, (_, bytes)| total.checked_add(bytes.len()))
        })
        .ok_or(CommentReplyError::InvalidSource { path })?;
    budget
        .charge_allocations(2, BudgetPath::Package)
        .and_then(|_| budget.charge_scratch_bytes(reassembly_scratch, BudgetPath::Package))
        .and_then(|_| budget.charge_transaction_work(reassembly_scratch, BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let prepared_zip = physical
        .package()
        .prepare_reassembly_with_deletions(&edits, &previews, physical.limits())
        .map_err(|_| CommentReplyError::Verification)?;
    let zip_requirements = prepared_zip.execution_requirements();
    budget
        .preflight_reassembly(zip_requirements, BudgetPath::Package)
        .map_err(|error| map_budget_error(error, path))?;
    budget
        .charge_compressed_bytes(zip_requirements.output_bytes(), BudgetPath::Package)
        .and_then(|_| {
            budget
                .charge_candidate_input_bytes(zip_requirements.output_bytes(), BudgetPath::Package)
        })
        .and_then(|_| {
            budget.charge_transaction_work(zip_requirements.output_bytes(), BudgetPath::Package)
        })
        .map_err(|error| map_budget_error(error, path))?;
    let zip_limits = zip_requirements.exact_limits();
    let bytes = prepared_zip
        .execute(zip_limits)
        .map_err(|_| CommentReplyError::Verification)?;
    budget
        .charge_allocations(1, BudgetPath::Package)
        .and_then(|_| budget.charge_scratch_bytes(bytes.len(), BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let candidate = Package::from_owned_bytes_with_options(bytes, edit.source.state.options)
        .map_err(|_| CommentReplyError::Verification)?;
    let candidate_catalog =
        physical_source(&candidate).map_err(|_| CommentReplyError::Verification)?;
    budget
        .charge_candidate_reopen(candidate_catalog, BudgetPath::Package)
        .map_err(|error| map_budget_error(error, path))?;
    let (sheet, table, row, column) = path.cell_parts();
    let position =
        CellPosition::try_from_usize(row, column).map_err(|_| CommentReplyError::Verification)?;
    let actual = candidate
        .table_cell_comment_replies(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )
        .map_err(|_| CommentReplyError::Verification)?;
    if actual.as_ref() != edit.after.as_slice() {
        return Err(CommentReplyError::Verification);
    }
    let source_owner = physical.__source_owner();
    let artifacts = OwnedExactArtifacts::new(source_owner, candidate_catalog.__source_owner());
    Ok(CommentReplyCommit {
        package: candidate,
        patch: CommentReplyPatch {
            artifacts,
            path,
            before: edit.before,
            after: edit.after.into_boxed_slice(),
            source_cell: native.source_cell.into(),
            target_cell: native.target_cell.into(),
            source_previews: 0,
            target_previews: 0,
            touched_components: native.touched_components.len(),
        },
        diagnostics: CommentReplyDiagnostics::published(native.touched_components.len(), 0),
    })
}

fn pack_reply_metadata(
    source: &Package,
    route: crate::package::metadata::MessageRoute,
    payload: Vec<u8>,
    path: CommentReplyPath,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, CommentReplyError> {
    let physical = source
        .state
        .components
        .physical()
        .ok_or(CommentReplyError::UnsupportedSource)?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == crate::package::metadata::ENTRY_NAME)
        .ok_or(CommentReplyError::InvalidSource { path })?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    let snappy_limits = physical
        .limits()
        .snappy_limits()
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    // Decompression and archive parsing allocate private buffers before their
    // exact report is available. Reserve that scratch envelope first; the
    // decoded archive and rewritten output are charged below from measured
    // lengths and bounded encoded sizes.
    budget
        .charge_allocations(3, BudgetPath::Package)
        .and_then(|_| budget.charge_scratch_bytes(entry.data().len(), BudgetPath::Package))
        .and_then(|_| budget.charge_transaction_work(entry.data().len(), BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    let message_count = archive
        .objects
        .iter()
        .map(|object| object.messages.len())
        .sum::<usize>();
    budget
        .charge_payload_bytes(stream.as_bytes().len(), BudgetPath::Package)
        .and_then(|_| budget.charge_payload_objects(archive.objects.len(), BudgetPath::Package))
        .and_then(|_| budget.charge_payload_messages(message_count, BudgetPath::Package))
        .map_err(|error| map_budget_error(error, path))?;
    let message = archive
        .objects
        .get_mut(route.object_index)
        .and_then(|object| object.messages.get_mut(route.message_index))
        .ok_or(CommentReplyError::InvalidSource { path })?;
    message.data = payload;
    let estimated = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    let maximum_compressed = SnappyStream::maximum_compressed_len(estimated)
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    budget
        .charge_output(estimated, BudgetPath::Package)
        .and_then(|_| budget.charge_retained_bytes(estimated, BudgetPath::Package))
        .and_then(|_| budget.charge_compressed_bytes(maximum_compressed, BudgetPath::Package))
        .and_then(|_| {
            budget.charge_transaction_work(
                estimated.saturating_add(maximum_compressed),
                BudgetPath::Package,
            )
        })
        .map_err(|error| map_budget_error(error, path))?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| CommentReplyError::InvalidSource { path })?;
    if bytes.len() > estimated {
        return Err(CommentReplyError::Verification);
    }
    let compressed =
        SnappyStream::compress(&bytes).map_err(|_| CommentReplyError::InvalidSource { path })?;
    if compressed.len() > maximum_compressed
        || compressed.len() > snappy_limits.max_compressed_stream()
    {
        return Err(CommentReplyError::LimitExceeded {
            kind: CommentReplyLimitKind::CompressedBytes,
            observed: compressed.len(),
            maximum: maximum_compressed.min(snappy_limits.max_compressed_stream()),
            path,
        });
    }
    Ok(compressed)
}

impl Package {
    /// Read one direct reply by ordinal.
    pub fn table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
        index: CommentReplyIndex,
    ) -> Result<CommentReply, CommentReplyError> {
        let state = resolve_reply_state(self, sheet, table, position)?;
        let path = root_to_reply_path(state.located.target.path);
        let replies = state.replies;
        replies
            .get(index.index() as usize)
            .cloned()
            .ok_or(CommentReplyError::ReplyNotFound {
                path: path.reply(index),
            })
    }

    /// Read one direct reply using an A1 address.
    pub fn table_cell_comment_reply_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
        index: CommentReplyIndex,
    ) -> Result<CommentReply, CommentReplyError> {
        let position =
            CellPosition::from_a1(address).map_err(|_| CommentReplyError::InvalidAddress)?;
        self.table_cell_comment_reply(sheet, table, position, index)
    }

    /// Start a collection edit for all direct replies in one rooted cell.
    pub fn edit_table_cell_comment_replies<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<CommentReplyEdit<'_>, CommentReplyError> {
        let state = resolve_reply_state(self, sheet, table, position)?;
        let path = root_to_reply_path(state.located.target.path).as_cell();
        let after = clone_replies(&state.replies, path)?;
        Ok(CommentReplyEdit {
            source: self,
            target: state.located,
            path,
            before: state.replies,
            after,
            operation: None,
            staging_error: None,
        })
    }

    /// Start a collection edit using an A1 address.
    pub fn edit_table_cell_comment_replies_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
    ) -> Result<CommentReplyEdit<'_>, CommentReplyError> {
        let position =
            CellPosition::from_a1(address).map_err(|_| CommentReplyError::InvalidAddress)?;
        self.edit_table_cell_comment_replies(sheet, table, position)
    }

    /// Replace one direct reply by ordinal.
    pub fn set_table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
        index: CommentReplyIndex,
        text: impl AsRef<str>,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        self.edit_table_cell_comment_replies(sheet, table, position)?
            .set(index, text)
            .commit()
    }

    /// Replace one direct reply using an A1 address.
    pub fn set_table_cell_comment_reply_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
        index: CommentReplyIndex,
        text: impl AsRef<str>,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        let position =
            CellPosition::from_a1(address).map_err(|_| CommentReplyError::InvalidAddress)?;
        self.set_table_cell_comment_reply(sheet, table, position, index, text)
    }

    /// Remove one direct reply by ordinal.
    pub fn remove_table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
        index: CommentReplyIndex,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        self.edit_table_cell_comment_replies(sheet, table, position)?
            .remove(index)
            .commit()
    }

    /// Remove one direct reply using an A1 address.
    pub fn remove_table_cell_comment_reply_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
        index: CommentReplyIndex,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        let position =
            CellPosition::from_a1(address).map_err(|_| CommentReplyError::InvalidAddress)?;
        self.remove_table_cell_comment_reply(sheet, table, position, index)
    }

    /// Append one direct reply after the existing reply list.
    pub fn add_table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
        text: impl AsRef<str>,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        self.edit_table_cell_comment_replies(sheet, table, position)?
            .append(text)
            .commit()
    }

    /// Append one direct reply using an A1 address.
    pub fn add_table_cell_comment_reply_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
        text: impl AsRef<str>,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        let position =
            CellPosition::from_a1(address).map_err(|_| CommentReplyError::InvalidAddress)?;
        self.add_table_cell_comment_reply(sheet, table, position, text)
    }

    /// Apply a reversible exact-source direct-reply patch.
    pub fn apply_table_cell_comment_reply(
        &self,
        patch: &CommentReplyPatch,
    ) -> Result<CommentReplyCommit, CommentReplyError> {
        let catalog = physical_source(self).map_err(|error| map_root_error(error, patch.path))?;
        let owner = catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&owner) {
            return Err(CommentReplyError::PatchConflict);
        }
        let (sheet, table, row, column) = patch.path.cell_parts();
        let position = CellPosition::try_from_usize(row, column)
            .map_err(|_| CommentReplyError::PatchConflict)?;
        let current = self
            .table_cell_comment_replies(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(|error| map_root_error(error, patch.path))?;
        if current.as_ref() != patch.before.as_ref() {
            return Err(CommentReplyError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(CommentReplyCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: CommentReplyDiagnostics::unchanged(),
            });
        }
        let target_owner = patch.artifacts.target_owner();
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| CommentReplyError::Verification)?;
        let after = candidate
            .table_cell_comment_replies(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(|error| map_root_error(error, patch.path))?;
        if after.as_ref() != patch.after.as_ref() {
            return Err(CommentReplyError::Verification);
        }
        Ok(CommentReplyCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: CommentReplyDiagnostics::published(
                patch.touched_components,
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }
}
