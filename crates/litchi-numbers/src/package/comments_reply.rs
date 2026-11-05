//! Selector-first direct-reply transactions for Numbers cell comments.
//!
//! This module is the package-facing half of the direct-reply owner.  It keeps
//! the public transaction vocabulary independent from the native comment
//! storage graph; the native transition is the shared copy-on-write reply
//! engine (`litchi_iwa_table::comments_native::rewrite_reply`), the host
//! writer's shape, which Pages and Keynote use as well.  The engine admits a
//! uniquely owned root whose direct replies are leaves owned only by it.

use std::{fmt, sync::Arc};

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_protos::package_metadata_codec::RewriteOptions as MetadataRewriteOptions;
use thiserror::Error as ThisError;

use crate::{SheetSelector, TableSelector, table::At};

use super::{
    CommentAuthor, CommentReply, CommentTimestamp, Error as RootError, LimitKind as RootLimitKind,
    Located, Package, Path as RootPath, physical_source, resolve_comment,
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
    pub fn try_from_usize(index: usize) -> Result<Self, Error> {
        u32::try_from(index)
            .map(Self)
            .map_err(|_| Error::Refused(Refusal::IndexOverflow { observed: index }))
    }
}

/// A content-free location associated with a direct-reply operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
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

impl Path {
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
pub enum LimitKind {
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
    /// Native messages inspected per owning container.
    MessagesPerContainer,
    /// Native list items inspected or retained.
    PayloadItems,
    /// Native metadata items inspected.
    MetadataItems,
    /// Native Snappy frames inspected.
    SnappyFrames,
    /// Addressable table rows inspected.
    TableRows,
    /// Addressable table columns inspected.
    TableColumns,
    /// Addressable table cells inspected.
    TableCells,
    /// Materialized sparse table cells inspected.
    MaterializedCells,
    /// Repeated Numbers-wire items inspected.
    Items,
    /// Comment replies inspected or emitted.
    Replies,
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
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The source graph or its native framing was not authoritative.
    #[error("the Numbers comment-reply source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// An operation-local resource ceiling was exceeded.
    #[error(
        "Numbers comment-reply {kind:?} limit exceeded: observed {observed}, maximum {maximum} at {path:?}"
    )]
    Limit {
        /// Resource axis.
        kind: LimitKind,
        /// Observed or requested quantity.
        observed: usize,
        /// Configured ceiling.
        maximum: usize,
        /// Semantic operation path.
        path: Path,
    },
    /// A bounded plan allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers comment-reply operation")]
    Allocation {
        /// Requested allocation units.
        amount: usize,
        /// Semantic operation path.
        path: Path,
    },
    /// Candidate reopen did not reproduce the requested state.
    #[error("the edited Numbers comment reply failed semantic verification")]
    Verification,
    /// A patch was applied to a different exact source.
    #[error("the Numbers comment-reply patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
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
    OutOfBounds { path: Path },
    /// The package does not retain exact physical provenance.
    #[error("this Numbers source does not support exact comment-reply editing")]
    UnsupportedSource,
    /// The selected cell has no rooted comment.
    #[error("the selected Numbers cell has no rooted comment")]
    CommentNotFound { path: Path },
    /// The selected direct-reply ordinal does not exist.
    #[error("the selected Numbers direct-reply ordinal does not exist")]
    ReplyNotFound { path: Path },
    /// The source graph needs a broader owner than this bounded slice.
    #[error("the Numbers comment-reply source has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: Path },
    /// The selected table is locked against comment-reply mutation.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    /// A platform-sized ordinal could not be represented by the semantic
    /// index type.
    #[error("the comment-reply ordinal is too large")]
    IndexOverflow { observed: usize },
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::InvalidAddress { .. } | Self::OutOfBounds { .. } | Self::IndexOverflow { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SheetNotFound { .. }
            | Self::TableNotFound { .. }
            | Self::CommentNotFound { .. }
            | Self::ReplyNotFound { .. } => RefusalKind::NotFound,
            Self::UnsupportedSource { .. } | Self::UnsupportedDependency { .. } => {
                RefusalKind::UnsupportedContent
            },
            Self::TableLocked { .. } => RefusalKind::Protected,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
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
pub struct Edit<'a> {
    source: &'a Package,
    target: Located,
    path: Path,
    before: Box<[CommentReply]>,
    after: Vec<CommentReply>,
    operation: Option<ReplyOperation>,
    staging_error: Option<Error>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path)
            .field("reply_count", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the rooted cell path selected for this edit.
    #[must_use]
    pub const fn path(&self) -> Path {
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
    ///
    /// The new reply is written by the package's commenting author (or the
    /// generated `litchi-iwa` author), as the host's reply writer does; the
    /// staged reply carries no metadata until the commit reads it back.
    #[must_use]
    pub fn append(mut self, text: impl AsRef<str>) -> Self {
        match reply_from_text(self.source, text.as_ref(), self.path, None, None) {
            Ok(reply) => {
                let Ok(after) = clone_replies(&self.before, self.path) else {
                    self.staging_error = Some(Error::Allocation {
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
            self.staging_error = Some(Error::Refused(Refusal::ReplyNotFound {
                path: self.path.reply(index),
            }));
            return self;
        }
        let (timestamp, author) = self
            .before
            .get(ordinal)
            .map(|reply| (reply.timestamp, reply.author.clone()))
            .unwrap_or((None, None));
        match reply_from_text(
            self.source,
            text.as_ref(),
            self.path.reply(index),
            timestamp,
            author,
        ) {
            Ok(reply) => {
                let Ok(after) = clone_replies(&self.before, self.path) else {
                    self.staging_error = Some(Error::Allocation {
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
            self.staging_error = Some(Error::Refused(Refusal::ReplyNotFound {
                path: self.path.reply(index),
            }));
            return self;
        }
        let Ok(after) = clone_replies(&self.before, self.path) else {
            self.staging_error = Some(Error::Allocation {
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
    pub fn commit(self) -> Result<Commit, Error> {
        if let Some(error) = self.staging_error {
            return Err(error);
        }
        commit_edit(self)
    }
}

/// A reversible exact-source direct-reply patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Box<[CommentReply]>,
    after: Box<[CommentReply]>,
    source_cell: Arc<[u8]>,
    target_cell: Arc<[u8]>,
    source_previews: usize,
    target_previews: usize,
    touched_components: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path)
            .field("reply_count", &self.after.len())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic path represented by this patch.
    #[must_use]
    pub const fn path(&self) -> Path {
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
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

struct ReplyState {
    located: Located,
    replies: Box<[CommentReply]>,
}

fn map_root_error(error: RootError, path: Path) -> Error {
    match error {
        RootError::Refused(crate::package::comments::Refusal::InvalidAddress) => {
            Error::Refused(Refusal::InvalidAddress)
        },
        RootError::Refused(crate::package::comments::Refusal::SheetNotFound) => {
            Error::Refused(Refusal::SheetNotFound)
        },
        RootError::Refused(crate::package::comments::Refusal::TableNotFound) => {
            Error::Refused(Refusal::TableNotFound)
        },
        RootError::Refused(crate::package::comments::Refusal::OutOfBounds { .. }) => {
            Error::Refused(Refusal::OutOfBounds { path })
        },
        RootError::Refused(crate::package::comments::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        RootError::Refused(crate::package::comments::Refusal::CommentNotFound { .. }) => {
            Error::Refused(Refusal::CommentNotFound { path })
        },
        RootError::Refused(crate::package::comments::Refusal::UnsupportedDependency { .. }) => {
            Error::Refused(Refusal::UnsupportedDependency { path })
        },
        RootError::Limit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                RootLimitKind::InputBytes => LimitKind::InputBytes,
                RootLimitKind::OutputBytes => LimitKind::OutputBytes,
                RootLimitKind::WireFields => LimitKind::Fields,
                RootLimitKind::WireBytes => LimitKind::EncodedBytes,
                RootLimitKind::WireWork => LimitKind::Work,
                RootLimitKind::TextBytes => LimitKind::TextBytes,
                RootLimitKind::References => LimitKind::References,
                RootLimitKind::ReferenceBytes => LimitKind::ReferenceBytes,
                RootLimitKind::CompressedBytes => LimitKind::CompressedBytes,
                RootLimitKind::Objects => LimitKind::PayloadObjects,
                RootLimitKind::Messages => LimitKind::PayloadMessages,
                RootLimitKind::MessagesPerObject => LimitKind::MessagesPerContainer,
                RootLimitKind::MetadataItems => LimitKind::MetadataItems,
                RootLimitKind::SnappyFrames => LimitKind::SnappyFrames,
                RootLimitKind::TableRows => LimitKind::TableRows,
                RootLimitKind::TableColumns => LimitKind::TableColumns,
                RootLimitKind::TableCells => LimitKind::TableCells,
                RootLimitKind::MaterializedCells => LimitKind::MaterializedCells,
                RootLimitKind::Items => LimitKind::Items,
                RootLimitKind::Replies => LimitKind::Replies,
                RootLimitKind::WireNesting => LimitKind::Nesting,
                RootLimitKind::ScratchBytes => LimitKind::ScratchBytes,
                RootLimitKind::RetainedBytes => LimitKind::RetainedBytes,
                RootLimitKind::Allocations => LimitKind::Allocations,
            },
            observed,
            maximum,
            path,
        },
        RootError::Allocation { amount, .. } => Error::Allocation { amount, path },
        RootError::Verification => Error::Verification,
        RootError::PatchConflict => Error::PatchConflict,
        RootError::InvalidSource { .. } => Error::InvalidSource { path },
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

fn map_metadata_error(error: super::super::comments_metadata::MetadataError, path: Path) -> Error {
    match error.kind {
        super::super::comments_metadata::FailureKind::Unsupported => {
            Error::Refused(Refusal::UnsupportedDependency { path })
        },
        super::super::comments_metadata::FailureKind::Limit => Error::Limit {
            kind: LimitKind::Work,
            observed: error.observed,
            maximum: error.maximum,
            path,
        },
        super::super::comments_metadata::FailureKind::Allocation => Error::Allocation {
            amount: error.allocation,
            path,
        },
        super::super::comments_metadata::FailureKind::MissingRoute
        | super::super::comments_metadata::FailureKind::AmbiguousRoute
        | super::super::comments_metadata::FailureKind::InvalidSource
        | super::super::comments_metadata::FailureKind::Conflict
        | super::super::comments_metadata::FailureKind::VersionedOwnership => {
            Error::InvalidSource { path }
        },
    }
}

fn reply_from_text(
    source: &Package,
    text: &str,
    path: Path,
    timestamp: Option<CommentTimestamp>,
    author: Option<CommentAuthor>,
) -> Result<CommentReply, Error> {
    let maximum = source.state.options.semantic().max_output_text_bytes();
    if text.len() > maximum {
        return Err(Error::Limit {
            kind: LimitKind::TextBytes,
            observed: text.len(),
            maximum,
            path,
        });
    }
    let mut retained = String::new();
    retained
        .try_reserve_exact(text.len())
        .map_err(|_| Error::Allocation {
            amount: text.len(),
            path,
        })?;
    retained.push_str(text);
    Ok(CommentReply {
        text: Arc::from(retained.into_boxed_str()),
        timestamp,
        author,
    })
}

fn clone_replies(source: &[CommentReply], path: Path) -> Result<Vec<CommentReply>, Error> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation {
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
    position: At,
) -> Result<ReplyState, Error> {
    let located = resolve_comment(source, sheet, table, position).map_err(|error| {
        let path = match error {
            RootError::Refused(crate::package::comments::Refusal::OutOfBounds { path })
            | RootError::InvalidSource { path }
            | RootError::Refused(crate::package::comments::Refusal::CommentNotFound { path })
            | RootError::Refused(crate::package::comments::Refusal::UnsupportedDependency {
                path,
            })
            | RootError::Limit { path, .. }
            | RootError::Allocation { path, .. } => root_to_reply_path(path),
            _ => Path::Package,
        };
        map_root_error(error, path)
    })?;
    let path = root_to_reply_path(located.target.path);
    if located.comment.is_none() {
        return Err(Error::Refused(Refusal::CommentNotFound { path }));
    }
    let replies = source
        .table_cell_comment_replies(
            SheetSelector::index(path.cell_parts().0),
            TableSelector::index(path.cell_parts().1),
            position,
        )
        .map_err(|error| map_root_error(error, path))?;
    if replies.len() != located.reply_ids.len() {
        return Err(Error::InvalidSource { path });
    }
    let entry = located
        .entry
        .as_ref()
        .ok_or(Error::InvalidSource { path })?;
    if !entry.owner.supports_text_rewrite() {
        return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
    }
    Ok(ReplyState { located, replies })
}

fn root_to_reply_path(path: RootPath) -> Path {
    match path {
        RootPath::Package => Path::Package,
        RootPath::Cell {
            sheet,
            table,
            row,
            column,
        } => Path::Cell {
            sheet,
            table,
            row,
            column,
        },
    }
}

fn no_op_commit(
    source: &Package,
    path: Path,
    before: &[CommentReply],
    after: &[CommentReply],
) -> Result<Commit, Error> {
    let catalog = physical_source(source).map_err(|error| map_root_error(error, path))?;
    let owner = catalog.source_owner();
    let before = before.to_vec().into_boxed_slice();
    let after = after.to_vec().into_boxed_slice();
    Ok(Commit {
        package: source.snapshot(),
        patch: Patch {
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
        diagnostics: Diagnostics::unchanged(),
    })
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let path = edit.path;
    if edit.before.as_ref() == edit.after.as_slice() {
        return no_op_commit(edit.source, path, &edit.before, &edit.after);
    }
    if edit.target.target.native.locked == crate::table::lock::State::Locked {
        return Err(Error::Refused(Refusal::TableLocked { path }));
    }
    let catalog = physical_source(edit.source).map_err(|error| map_root_error(error, path))?;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    // Numbers keeps its strict package-metadata admission ahead of the
    // shared engine: a reply write must not publish over an unowned,
    // versioned or ambiguous metadata registry.
    let entry = edit
        .target
        .entry
        .as_ref()
        .ok_or(Error::InvalidSource { path })?;
    if !matches!(entry.owner, super::EntryOwner::Root) || entry.storage_occurrences != 1 {
        return Err(Error::Refused(Refusal::UnsupportedDependency { path }));
    }
    let metadata = crate::package::comments_metadata::strict_source(edit.source)
        .map_err(|error| map_metadata_error(error, path))?;
    let metadata_facts = crate::package::comments_metadata::inspect(
        edit.source,
        metadata_rewrite_options(metadata.payload.len(), 2),
    )
    .map_err(|error| map_metadata_error(error, path))?;
    // Every storage the copy-on-write write may release (the root and its
    // direct replies) must own exactly one registry UUID binding.
    let components = edit.source.state.components.catalog();
    for identifier in
        std::iter::once(entry.entry.storage_id).chain(edit.target.reply_ids.iter().copied())
    {
        let component_index = components
            .iter()
            .position(|component| {
                component
                    .archive()
                    .objects
                    .iter()
                    .any(|object| object.archive_info.identifier == Some(identifier))
            })
            .ok_or(Error::InvalidSource { path })?;
        metadata_facts
            .uuid_removal(component_index, identifier)
            .map_err(|_| Error::Refused(Refusal::UnsupportedDependency { path }))?;
    }
    crate::package::comments_metadata::reject_unknown_archive_metadata(
        edit.source,
        catalog.limits().archive_limits(),
    )
    .map_err(|error| map_metadata_error(error, path))?;
    let operation = match edit.operation.as_ref() {
        Some(ReplyOperation::Append { text }) => {
            litchi_iwa_table::comments_native::ReplyOperation::Append(Box::from(text.as_ref()))
        },
        Some(ReplyOperation::Set { index, text }) => {
            litchi_iwa_table::comments_native::ReplyOperation::Replace {
                index: index.index() as usize,
                text: Box::from(text.as_ref()),
            }
        },
        Some(ReplyOperation::Remove { index }) => {
            litchi_iwa_table::comments_native::ReplyOperation::Remove {
                index: index.index() as usize,
            }
        },
        None => return no_op_commit(edit.source, path, &edit.before, &edit.after),
    };
    // The shared copy-on-write reply engine (the host's writer shape) owns
    // the native transition, its metadata registry transition and the
    // candidate reopen; this facade keeps selection and semantic readback.
    let root_path = edit.target.target.path;
    let selection = super::native_selection(edit.source, &edit.target)
        .map_err(|error| map_root_error(error, path))?;
    let mutation = litchi_iwa_table::comments_native::rewrite_reply(
        edit.source,
        &selection,
        &operation,
        litchi_iwa_table::comments_native::Profile::NUMBERS,
    )
    .map_err(|fault| map_root_error(super::map_native_fault(fault, root_path), path))?;
    let touched_components = mutation.touched_components();
    let deleted_previews = mutation.deleted_previews();
    let candidate = mutation.into_package();
    let (sheet, table, row, column) = path.cell_parts();
    let position = At::try_from_usize(row, column).map_err(|_| Error::Verification)?;
    let actual = candidate
        .table_cell_comment_replies(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )
        .map_err(|_| Error::Verification)?;
    // A new reply carries the engine's author and date, so the staged
    // thread is checked by text; the patch records the published thread.
    if actual.len() != edit.after.len()
        || actual
            .iter()
            .zip(&edit.after)
            .any(|(actual, expected)| actual.text() != expected.text())
    {
        return Err(Error::Verification);
    }
    let readback = resolve_comment(
        &candidate,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )
    .map_err(|error| map_root_error(error, path))?;
    let source_cell = edit
        .target
        .cell_bytes
        .clone()
        .ok_or(Error::InvalidSource { path })?;
    let target_cell = readback.cell_bytes.ok_or(Error::Verification)?;
    let source_previews = super::root_preview_deletions(catalog)
        .map_err(|error| map_root_error(error, path))?
        .len();
    let source_owner = catalog.source_owner();
    let target_owner = physical_source(&candidate)
        .map_err(|_| Error::Verification)?
        .source_owner();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            path,
            before: edit.before,
            after: actual,
            source_cell,
            target_cell,
            source_previews,
            target_previews: source_previews.saturating_sub(deleted_previews),
            touched_components,
        },
        diagnostics: Diagnostics::published(touched_components, deleted_previews),
    })
}

impl Package {
    /// Read one direct reply by ordinal.
    pub fn table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        index: CommentReplyIndex,
    ) -> Result<CommentReply, Error> {
        let state = resolve_reply_state(self, sheet, table, position)?;
        let path = root_to_reply_path(state.located.target.path);
        let replies = state.replies;
        replies
            .get(index.index() as usize)
            .cloned()
            .ok_or(Error::Refused(Refusal::ReplyNotFound {
                path: path.reply(index),
            }))
    }

    /// Read one direct reply using an A1 address.
    pub fn table_cell_comment_reply_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
        index: CommentReplyIndex,
    ) -> Result<CommentReply, Error> {
        let position = At::from_a1(address).map_err(|_| Error::Refused(Refusal::InvalidAddress))?;
        self.table_cell_comment_reply(sheet, table, position, index)
    }

    /// Start a collection edit for all direct replies in one rooted cell.
    pub fn edit_table_cell_comment_replies<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Edit<'_>, Error> {
        let state = resolve_reply_state(self, sheet, table, position)?;
        let path = root_to_reply_path(state.located.target.path).as_cell();
        let after = clone_replies(&state.replies, path)?;
        Ok(Edit {
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
    ) -> Result<Edit<'_>, Error> {
        let position = At::from_a1(address).map_err(|_| Error::Refused(Refusal::InvalidAddress))?;
        self.edit_table_cell_comment_replies(sheet, table, position)
    }

    /// Replace one direct reply by ordinal.
    pub fn set_table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        index: CommentReplyIndex,
        text: impl AsRef<str>,
    ) -> Result<Commit, Error> {
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
    ) -> Result<Commit, Error> {
        let position = At::from_a1(address).map_err(|_| Error::Refused(Refusal::InvalidAddress))?;
        self.set_table_cell_comment_reply(sheet, table, position, index, text)
    }

    /// Remove one direct reply by ordinal.
    pub fn remove_table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        index: CommentReplyIndex,
    ) -> Result<Commit, Error> {
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
    ) -> Result<Commit, Error> {
        let position = At::from_a1(address).map_err(|_| Error::Refused(Refusal::InvalidAddress))?;
        self.remove_table_cell_comment_reply(sheet, table, position, index)
    }

    /// Append one direct reply after the existing reply list.
    pub fn add_table_cell_comment_reply<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
        text: impl AsRef<str>,
    ) -> Result<Commit, Error> {
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
    ) -> Result<Commit, Error> {
        let position = At::from_a1(address).map_err(|_| Error::Refused(Refusal::InvalidAddress))?;
        self.add_table_cell_comment_reply(sheet, table, position, text)
    }

    /// Apply a reversible exact-source direct-reply patch.
    pub fn apply_table_cell_comment_reply(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_source(self).map_err(|error| map_root_error(error, patch.path))?;
        let owner = catalog.source_owner();
        if !patch.artifacts.authorizes_owner(&owner) {
            return Err(Error::PatchConflict);
        }
        let (sheet, table, row, column) = patch.path.cell_parts();
        let position = At::try_from_usize(row, column).map_err(|_| Error::PatchConflict)?;
        let current = self
            .table_cell_comment_replies(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(|error| map_root_error(error, patch.path))?;
        if current.as_ref() != patch.before.as_ref() {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        require_exact_apply_source(catalog.source_is_exact())?;
        let target_owner = patch.artifacts.target_owner();
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let after = candidate
            .table_cell_comment_replies(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(|error| map_root_error(error, patch.path))?;
        if after.as_ref() != patch.after.as_ref() {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.touched_components,
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }
}

fn require_exact_apply_source(source_is_exact: bool) -> Result<(), Error> {
    if source_is_exact {
        Ok(())
    } else {
        Err(Error::Refused(Refusal::UnsupportedSource))
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, LimitKind, Path, Refusal, map_root_error, require_exact_apply_source};

    #[test]
    fn root_limit_axes_reach_public_error_without_coercion() {
        let cases = [
            (
                crate::package::comments::LimitKind::InputBytes,
                LimitKind::InputBytes,
            ),
            (
                crate::package::comments::LimitKind::OutputBytes,
                LimitKind::OutputBytes,
            ),
            (
                crate::package::comments::LimitKind::WireFields,
                LimitKind::Fields,
            ),
            (
                crate::package::comments::LimitKind::WireBytes,
                LimitKind::EncodedBytes,
            ),
            (
                crate::package::comments::LimitKind::WireWork,
                LimitKind::Work,
            ),
            (
                crate::package::comments::LimitKind::Objects,
                LimitKind::PayloadObjects,
            ),
            (
                crate::package::comments::LimitKind::Messages,
                LimitKind::PayloadMessages,
            ),
            (
                crate::package::comments::LimitKind::MessagesPerObject,
                LimitKind::MessagesPerContainer,
            ),
            (
                crate::package::comments::LimitKind::MetadataItems,
                LimitKind::MetadataItems,
            ),
            (
                crate::package::comments::LimitKind::SnappyFrames,
                LimitKind::SnappyFrames,
            ),
            (
                crate::package::comments::LimitKind::TableRows,
                LimitKind::TableRows,
            ),
            (
                crate::package::comments::LimitKind::TableColumns,
                LimitKind::TableColumns,
            ),
            (
                crate::package::comments::LimitKind::TableCells,
                LimitKind::TableCells,
            ),
            (
                crate::package::comments::LimitKind::MaterializedCells,
                LimitKind::MaterializedCells,
            ),
            (crate::package::comments::LimitKind::Items, LimitKind::Items),
            (
                crate::package::comments::LimitKind::TextBytes,
                LimitKind::TextBytes,
            ),
            (
                crate::package::comments::LimitKind::References,
                LimitKind::References,
            ),
            (
                crate::package::comments::LimitKind::ReferenceBytes,
                LimitKind::ReferenceBytes,
            ),
            (
                crate::package::comments::LimitKind::Replies,
                LimitKind::Replies,
            ),
            (
                crate::package::comments::LimitKind::WireNesting,
                LimitKind::Nesting,
            ),
            (
                crate::package::comments::LimitKind::ScratchBytes,
                LimitKind::ScratchBytes,
            ),
            (
                crate::package::comments::LimitKind::RetainedBytes,
                LimitKind::RetainedBytes,
            ),
            (
                crate::package::comments::LimitKind::Allocations,
                LimitKind::Allocations,
            ),
            (
                crate::package::comments::LimitKind::CompressedBytes,
                LimitKind::CompressedBytes,
            ),
        ];
        for (kind, expected) in cases {
            let error = map_root_error(
                crate::package::comments::Error::Limit {
                    kind,
                    observed: 53,
                    maximum: 31,
                    path: crate::package::comments::Path::Package,
                },
                Path::Package,
            );
            assert_eq!(
                error,
                Error::Limit {
                    kind: expected,
                    observed: 53,
                    maximum: 31,
                    path: Path::Package,
                }
            );
        }
    }

    #[test]
    fn changed_direct_apply_requires_exact_physical_provenance() {
        assert_eq!(
            require_exact_apply_source(false),
            Err(Error::Refused(Refusal::UnsupportedSource))
        );
        assert_eq!(require_exact_apply_source(true), Ok(()));
    }
}
