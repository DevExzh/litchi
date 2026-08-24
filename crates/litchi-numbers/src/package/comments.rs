//! Selector-first semantic comments for rooted Numbers table cells.
//!
//! This module is deliberately separate from the scalar-cell and formula
//! editors.  A cell comment is a format-owned annotation graph, not a cell
//! value.  The public boundary therefore exposes only a small text-bearing
//! semantic value; native object identifiers, table-list keys, protobuf
//! messages, and archive member names remain private to this adapter.
//!
//! The exact-source write seam is intentionally narrow: it can replace the
//! text of an existing, unshared root comment whose table-list entry and cell
//! key are globally unique and whose storage has no replies. Creating
//! comments and clearing comments are refused; both operations require
//! ownership-graph mutations that this adapter does not perform.

use std::{collections::HashSet, fmt, sync::Arc};

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{WireView, patch_length_delimited_field};
use litchi_iwa_core::{
    Archive, ArchiveReferenceKind, ArchiveReferenceOccurrence, ArchiveReferencePolicy,
    ArchiveReferenceVisitor, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{
    comment_storage_codec, numbers_table_cell_storage_codec,
    package_metadata_codec::{
        RewriteOptions as MetadataRewriteOptions, inspect_package_metadata_with_visitor,
    },
    tst,
};
use thiserror::Error as ThisError;

use crate::{SheetSelector, TableSelector, table::CellPosition};

use super::Package;

const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPES: [u32; 2] = [6_005, 6_201];
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const DEFAULT_TILE_SIZE: usize = 256;
const COMMENT_MAX_NESTING: u32 = 64;
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// A semantic comment attached to one table cell.
///
/// Numbers comment metadata (authors, reply identities, and storage UUIDs)
/// is format-owned.  The focused package seam intentionally publishes only
/// the editable text, so callers cannot depend on native identifiers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    text: Arc<str>,
}

impl Comment {
    /// Construct a semantic comment value.
    ///
    /// This convenience constructor has no limit-bearing `Result` return
    /// type.  Source-backed edits use [`Edit::set`], which validates the
    /// caller's text against the package's semantic limit and copies it with
    /// a fallible reservation before staging it.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: Arc::from(text.into().into_boxed_str()),
        }
    }

    /// Fallibly construct a bounded semantic comment value.
    ///
    /// This constructor applies the hard Numbers semantic text ceiling. A
    /// package edit may impose a lower configured ceiling; [`Edit::set`]
    /// performs that package-specific validation before staging. The text is
    /// copied with a fallible reservation, so oversized or allocation-failed
    /// values are rejected without publishing a partial comment value.
    pub fn try_new(text: impl AsRef<str>) -> Result<Self, Error> {
        Self::try_from_text(
            text.as_ref(),
            super::SemanticLimits::MAX_OUTPUT_TEXT_BYTES,
            Path::Package,
        )
    }

    /// Borrow the comment text.
    #[must_use]
    pub fn text(&self) -> &str {
        self.text.as_ref()
    }

    fn try_from_text(text: &str, maximum: usize, path: Path) -> Result<Self, Error> {
        if text.len() > maximum {
            return Err(Error::LimitExceeded {
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
        Ok(Self {
            text: Arc::from(retained.into_boxed_str()),
        })
    }
}

/// A content-free location associated with a comment operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One selected table and cell, in zero-based semantic positions.
    Cell {
        /// Rooted sheet position.
        sheet: usize,
        /// Table position within the sheet.
        table: usize,
        /// Zero-based row.
        row: usize,
        /// Zero-based column.
        column: usize,
    },
}

impl Path {
    const fn cell_parts(self) -> (usize, usize, usize, usize) {
        match self {
            Self::Cell {
                sheet,
                table,
                row,
                column,
            } => (sheet, table, row, column),
            Self::Package => (0, 0, 0, 0),
        }
    }

    const fn sheet(self) -> usize {
        self.cell_parts().0
    }

    const fn table(self) -> usize {
        self.cell_parts().1
    }
}

fn cell_position_for_path(path: Path) -> Result<CellPosition, Error> {
    match path {
        Path::Cell { row, column, .. } => {
            CellPosition::try_from_usize(row, column).map_err(|_| Error::InvalidSource { path })
        },
        Path::Package => Err(Error::PatchConflict),
    }
}

/// A finite resource governed by a comment transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Encoded source bytes inspected.
    InputBytes,
    /// Candidate package bytes emitted.
    OutputBytes,
    /// Protobuf fields inspected.
    WireFields,
    /// Protobuf bytes inspected.
    WireBytes,
    /// Protobuf rewrite work.
    WireWork,
    /// Comment text bytes retained or requested.
    TextBytes,
    /// Native references inspected.
    References,
}

/// Failure from a selector-first Numbers comment operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A relative or absolute A1 address is malformed.
    #[error("invalid Numbers comment cell address")]
    InvalidAddress,
    /// The sheet selector did not resolve.
    #[error("the Numbers workbook has no sheet matching the comment selector")]
    SheetNotFound,
    /// The table selector did not resolve.
    #[error("the selected Numbers sheet has no table matching the comment selector")]
    TableNotFound,
    /// The requested cell is outside the selected table.
    #[error("the Numbers comment cell is outside the selected table")]
    OutOfBounds { path: Path },
    /// The source has an ambiguous or malformed native ownership graph.
    #[error("the Numbers comment source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// The source does not retain an exact physical package suitable for edit.
    #[error("this Numbers source does not support exact comment editing")]
    UnsupportedSource,
    /// The selected cell has no existing comment for an operation that needs one.
    #[error("the selected Numbers cell has no comment")]
    CommentNotFound { path: Path },
    /// Creating a new native comment would require an unsupported ownership graph.
    #[error("creating a new Numbers cell comment is not supported for this source")]
    UnsupportedDependency { path: Path },
    /// A bounded resource ceiling was exceeded.
    #[error(
        "Numbers comment {kind:?} limit exceeded: observed {observed}, maximum {maximum} at {path:?}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
        /// Semantic location.
        path: Path,
    },
    /// A fallible allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers comment operation at {path:?}")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
        /// Semantic location.
        path: Path,
    },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Numbers comment failed semantic verification")]
    Verification,
    /// A patch was applied to another source package.
    #[error("the Numbers comment patch does not match the exact source package")]
    PatchConflict,
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

    /// Number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of canonical previews removed by the changed operation.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate was reopened and verified.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// A selector-first comment edit staged against one immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    before: Option<Comment>,
    after: Option<Comment>,
    staging_error: Option<Error>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.target.path)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected semantic cell.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.target.path
    }

    /// Return the staged comment, if any.
    #[must_use]
    pub fn comment(&self) -> Option<&Comment> {
        self.after.as_ref()
    }

    /// Stage a replacement for the selected existing comment.
    ///
    /// New native comments are not part of this seam. Text is copied with a
    /// bounded, fallible allocation before the edit is staged.
    pub fn set(mut self, text: impl AsRef<str>) -> Self {
        let maximum = self.source.state.options.semantic().max_output_text_bytes();
        match Comment::try_from_text(text.as_ref(), maximum, self.target.path) {
            Ok(comment) => {
                self.after = Some(comment);
                self.staging_error = None;
            },
            Err(error) => {
                self.staging_error = Some(error);
            },
        }
        self
    }

    /// Stage removal of the selected comment.
    ///
    /// A changed clear is rejected by [`Edit::commit`].  The operation is
    /// retained as a staging convenience so callers can detect the typed
    /// [`Error::UnsupportedDependency`] result without any partial rewrite.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
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

/// A reversible process-local exact-source comment patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Option<Comment>,
    after: Option<Comment>,
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
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected semantic cell.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the source comment state.
    #[must_use]
    pub fn before(&self) -> Option<&Comment> {
        self.before.as_ref()
    }

    /// Return the target comment state.
    #[must_use]
    pub fn after(&self) -> Option<&Comment> {
        self.after.as_ref()
    }

    /// Return the source diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether the patch is an exact no-op.
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
            source_cell: Arc::clone(&self.target_cell),
            target_cell: Arc::clone(&self.source_cell),
            source_previews: self.target_previews,
            target_previews: self.source_previews,
            touched_components: self.touched_components,
        }
    }
}

/// One fully validated immutable comment publication.
#[must_use = "a comment commit contains the validated package snapshot"]
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

    /// Consume this publication and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
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

#[derive(Debug, Clone)]
struct Target {
    path: Path,
    native: super::table_headers::Target,
    row: usize,
    column: usize,
    tile_size: usize,
}

#[derive(Debug, Clone)]
struct Located {
    target: Target,
    comment: Option<Comment>,
    entry: Option<CommentEntryLocation>,
    storage: Option<MessageRoute>,
    cell_bytes: Option<Arc<[u8]>>,
    comment_key: Option<u32>,
    comment_table_id: Option<u64>,
    replies: usize,
}

#[derive(Debug, Clone)]
struct CommentStorageFact {
    object_id: u64,
    author_id: Option<u64>,
    replies: usize,
    reply_ids: Vec<u64>,
    storage_uuid: Option<(u64, u64)>,
}

#[derive(Debug, Clone, Copy)]
struct CommentListFact {
    table_id: u64,
    list_type: i32,
    key: u32,
    storage_id: Option<u64>,
    root: bool,
}

#[derive(Debug, Default)]
struct CommentOwnershipCensus {
    comment_list_ids: Vec<u64>,
    rooted_segment_ids: Vec<u64>,
    list_entries: Vec<CommentListFact>,
    table_references: Vec<u64>,
    cell_comment_keys: Vec<u32>,
    storages: Vec<CommentStorageFact>,
    author_ids: Vec<u64>,
    reply_ids: Vec<u64>,
    entry_storage_ids: Vec<u64>,
    uuids: Vec<(u64, u64)>,
}

#[derive(Debug, Clone, Copy)]
struct MessageRoute {
    component_index: usize,
    object_index: usize,
    message_index: usize,
    message_type: u32,
}

#[derive(Debug, Clone)]
struct Mutation {
    route: MessageRoute,
    data: Vec<u8>,
}

#[derive(Debug, Clone)]
struct CommentEntryLocation {
    owner: EntryOwner,
    route: MessageRoute,
    entry: EntryFact,
    storage_occurrences: usize,
}

#[derive(Debug, Clone, Copy)]
enum EntryOwner {
    Root,
    Segment,
}

impl EntryOwner {
    // Segment-backed entries are readable but deliberately outside the text
    // rewrite seam: publishing one would require updating the owning root
    // list's segment graph, not just the storage payload.
    const fn supports_text_rewrite(self) -> bool {
        matches!(self, Self::Root)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct EntryFact {
    refcount: u32,
    storage_id: u64,
}

impl Package {
    /// Read one semantic comment from a selector-first table cell.
    pub fn table_cell_comment<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<Comment>, Error> {
        let located = resolve_comment(self, sheet, table, position)?;
        Ok(located.comment)
    }

    /// Read one semantic comment using a relative or absolute A1 address.
    pub fn table_cell_comment_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
    ) -> Result<Option<Comment>, Error> {
        let position = CellPosition::from_a1(address).map_err(|_| Error::InvalidAddress)?;
        self.table_cell_comment(sheet, table, position)
    }

    /// Start a selector-first immutable comment edit.
    pub fn edit_table_cell_comment<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Edit<'_>, Error> {
        let located = resolve_comment(self, sheet, table, position)?;
        Ok(Edit {
            source: self,
            target: located.target,
            before: located.comment.clone(),
            after: located.comment,
            staging_error: None,
        })
    }

    /// Replace one existing selector-first table-cell comment.
    pub fn set_table_cell_comment<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
        text: impl AsRef<str>,
    ) -> Result<Commit, Error> {
        self.edit_table_cell_comment(sheet, table, position)?
            .set(text)
            .commit()
    }

    /// Remove a table-cell comment.
    ///
    /// Changed clears are intentionally refused with
    /// [`Error::UnsupportedDependency`]. Removing the cell pointer requires
    /// owner-aware table-list and comment-storage graph cleanup, so this
    /// adapter never publishes a cell-only rewrite that could leave an
    /// orphaned comment graph. Clearing an already-empty cell remains the
    /// exact no-op supported by the edit transaction.
    pub fn clear_table_cell_comment<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Commit, Error> {
        self.edit_table_cell_comment(sheet, table, position)?
            .clear()
            .commit()
    }

    /// Replace an existing table-cell comment using an A1 address.
    pub fn set_table_cell_comment_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
        text: impl AsRef<str>,
    ) -> Result<Commit, Error> {
        let position = CellPosition::from_a1(address).map_err(|_| Error::InvalidAddress)?;
        self.set_table_cell_comment(sheet, table, position, text)
    }

    /// Clear a table-cell comment using an A1 address.
    pub fn clear_table_cell_comment_a1<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        address: &str,
    ) -> Result<Commit, Error> {
        let position = CellPosition::from_a1(address).map_err(|_| Error::InvalidAddress)?;
        self.clear_table_cell_comment(sheet, table, position)
    }

    /// Apply a reversible exact-source comment patch.
    pub fn apply_table_cell_comment(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_catalog = physical_source(self)?;
        let source = source_catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let position = cell_position_for_path(patch.path)?;
        let located = resolve_comment_at_target(self, patch.path, position, patch.before.as_ref())?;
        if located.comment != patch.before
            || located.cell_bytes.as_deref() != Some(patch.source_cell.as_ref())
        {
            return Err(Error::PatchConflict);
        }
        if patch.before.is_some() {
            let Some(entry) = located.entry.as_ref() else {
                return Err(Error::PatchConflict);
            };
            if !entry.owner.supports_text_rewrite()
                || entry.entry.refcount != 1
                || entry.storage_occurrences != 1
                || located.replies != 0
            {
                return Err(Error::UnsupportedDependency { path: patch.path });
            }
            prove_global_comment_ownership(self, &located)?;
        }
        let source_previews = root_preview_deletions(source_catalog)?;
        if source_previews.len() != patch.source_previews {
            return Err(Error::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        let package = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let target_previews = root_preview_deletions(physical_source(&package)?)?;
        if target_previews.len() != patch.target_previews {
            return Err(Error::Verification);
        }
        if package.table_cell_comment(
            SheetSelector::index(patch.path.sheet()),
            TableSelector::index(patch.path.table()),
            position,
        )? != patch.after
        {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.touched_components,
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }
}

fn resolve_comment<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    position: CellPosition,
) -> Result<Located, Error> {
    let selected_sheet = source
        .state
        .document
        .sheet(sheet)
        .map_err(|_| Error::InvalidSource {
            path: Path::Package,
        })?
        .ok_or(Error::SheetNotFound)?;
    if selected_sheet
        .tables()
        .enumerate()
        .any(|(position, candidate)| {
            selected_sheet
                .tables()
                .skip(position.saturating_add(1))
                .any(|later| later.name() == candidate.name())
        })
    {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    let table_position = match table.into() {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => {
            let mut matches = selected_sheet
                .tables()
                .enumerate()
                .filter(|(_, candidate)| candidate.name() == name);
            let selected = matches.next();
            if selected.is_some() && matches.next().is_some() {
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            }
            selected.map(|(index, _)| index)
        },
    }
    .ok_or(Error::TableNotFound)?;
    let table = selected_sheet
        .tables()
        .nth(table_position)
        .ok_or(Error::TableNotFound)?;
    let dimensions = table.dimensions();
    let row = usize::try_from(position.row()).map_err(|_| Error::InvalidSource {
        path: Path::Package,
    })?;
    let column = usize::try_from(position.column()).map_err(|_| Error::InvalidSource {
        path: Path::Package,
    })?;
    if position.row() >= dimensions.rows() || position.column() >= dimensions.columns() {
        return Err(Error::OutOfBounds {
            path: Path::Cell {
                sheet: selected_sheet.index(),
                table: table_position,
                row,
                column,
            },
        });
    }
    let path = Path::Cell {
        sheet: selected_sheet.index(),
        table: table_position,
        row,
        column,
    };
    let native = super::table_headers::resolve::resolve_target(
        source,
        selected_sheet.index(),
        table_position,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    resolve_comment_native(
        source,
        Target {
            path,
            native,
            row,
            column,
            tile_size: DEFAULT_TILE_SIZE,
        },
    )
}

fn resolve_comment_at_target(
    source: &Package,
    path: Path,
    position: CellPosition,
    expected: Option<&Comment>,
) -> Result<Located, Error> {
    let Path::Cell { sheet, table, .. } = path else {
        return Err(Error::PatchConflict);
    };
    let native = super::table_headers::resolve::resolve_target(source, sheet, table)
        .map_err(|_| Error::InvalidSource { path })?;
    let row = usize::try_from(position.row()).map_err(|_| Error::PatchConflict)?;
    let column = usize::try_from(position.column()).map_err(|_| Error::PatchConflict)?;
    let located = resolve_comment_native(
        source,
        Target {
            path,
            native,
            row,
            column,
            tile_size: DEFAULT_TILE_SIZE,
        },
    )?;
    if expected.is_some() != located.comment.is_some() {
        return Err(Error::PatchConflict);
    }
    Ok(located)
}

#[derive(Clone, Copy)]
struct DecodedTableStorage<'source> {
    data_store: numbers_table_cell_storage_codec::DataStoreSnapshot<'source>,
    tile_storage: numbers_table_cell_storage_codec::TileStorageSnapshot,
}

fn decode_table_storage<'source>(
    source: &'source Package,
    target: &Target,
) -> Result<DecodedTableStorage<'source>, Error> {
    let model_route = MessageRoute {
        component_index: target.native.component_index,
        object_index: target.native.object_index,
        message_index: target.native.message_index,
        message_type: target.native.message_type,
    };
    let model_message = message_at_route(source, model_route, target.path)?;
    let model_options = table_cell_decode_options(
        source,
        model_message.data.len(),
        source.state.options.semantic().max_references(),
    );
    let (model, compatibility_data_store) =
        match numbers_table_cell_storage_codec::decode_table_model_with_report(
            model_message.data.as_slice(),
            model_options,
        ) {
            Ok((model, _report)) => (model, false),
            Err(error) if error.resource_limit().is_none() => {
                let (model, _report) = numbers_table_cell_storage_codec::decode_table_model_with_compatibility_data_store_with_report(
                    model_message.data.as_slice(),
                    model_options,
                )
                .map_err(|fallback| map_table_codec_error(fallback, target.path))?;
                (model, true)
            },
            Err(error) => return Err(map_table_codec_error(error, target.path)),
        };
    let data_store_options = table_cell_decode_options(
        source,
        model.base_data_store().len(),
        source.state.options.semantic().max_references(),
    );
    let (data_store, _report) = if compatibility_data_store {
        numbers_table_cell_storage_codec::decode_data_store_compatibility_with_report(
            model.base_data_store(),
            data_store_options,
        )
    } else {
        numbers_table_cell_storage_codec::decode_data_store_with_report(
            model.base_data_store(),
            data_store_options,
        )
    }
    .map_err(|error| map_table_codec_error(error, target.path))?;
    let (tile_storage, _report) =
        numbers_table_cell_storage_codec::decode_tile_storage_with_report(
            data_store.tiles(),
            data_store_options,
        )
        .map_err(|error| map_table_codec_error(error, target.path))?;
    Ok(DecodedTableStorage {
        data_store,
        tile_storage,
    })
}

fn resolve_comment_native(source: &Package, mut target: Target) -> Result<Located, Error> {
    let decoded = decode_table_storage(source, &target)?;
    let data_store = decoded.data_store;
    let tile_storage = decoded.tile_storage;
    let tile_size = usize::try_from(
        tile_storage
            .tile_size()
            .unwrap_or(u32::try_from(DEFAULT_TILE_SIZE).unwrap_or(u32::MAX)),
    )
    .map_err(|_| Error::InvalidSource { path: target.path })?;
    if tile_size == 0 {
        return Err(Error::InvalidSource { path: target.path });
    }
    target.tile_size = tile_size;
    let tile_key = target.row / tile_size;
    let mut tile_finder = TileFinder {
        target: u32::try_from(tile_key).map_err(|_| Error::InvalidSource { path: target.path })?,
        ..TileFinder::default()
    };
    let (_, _report) = numbers_table_cell_storage_codec::decode_tile_storage_with_visitor(
        data_store.tiles(),
        table_cell_decode_options(
            source,
            data_store.tiles().len(),
            source.state.options.semantic().max_references(),
        ),
        &mut tile_finder,
    )
    .map_err(|error| map_table_codec_error(error, target.path))?;
    if tile_finder.duplicate {
        return Err(Error::InvalidSource { path: target.path });
    }
    let Some(tile_id) = tile_finder.identifier else {
        return Ok(Located {
            target,
            comment: None,
            entry: None,
            storage: None,
            cell_bytes: None,
            comment_key: None,
            comment_table_id: None,
            replies: 0,
        });
    };
    let tile_resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, tile_id)
        .map_err(|_| Error::InvalidSource { path: target.path })?
        .ok_or(Error::InvalidSource { path: target.path })?;
    let tile_route = MessageRoute {
        component_index: tile_resolved.component_index,
        object_index: tile_resolved.object_index,
        message_index: unique_message_index(
            tile_resolved.messages,
            TILE_MESSAGE_TYPE,
            target.path,
        )?,
        message_type: TILE_MESSAGE_TYPE,
    };
    let tile_message = message_at_route(source, tile_route, target.path)?;
    let mut row_finder = RowFinder {
        target: u32::try_from(target.row % tile_size)
            .map_err(|_| Error::InvalidSource { path: target.path })?,
        ..RowFinder::default()
    };
    let (_, _report) = numbers_table_cell_storage_codec::decode_tile_with_visitor(
        tile_message.data.as_slice(),
        table_cell_decode_options(
            source,
            tile_message.data.len(),
            source.state.options.semantic().max_references(),
        ),
        &mut row_finder,
    )
    .map_err(|error| map_table_codec_error(error, target.path))?;
    if row_finder.duplicate {
        return Err(Error::InvalidSource { path: target.path });
    }
    let Some(row_occurrence) = row_finder.selected_occurrence else {
        return Ok(Located {
            target,
            comment: None,
            entry: None,
            storage: None,
            cell_bytes: None,
            comment_key: None,
            comment_table_id: None,
            replies: 0,
        });
    };
    let tile_view = WireView::parse_with_limits(
        tile_message.data.as_slice(),
        wire_limits_for(tile_message.data.len(), 0, target.path)?,
    )
    .map_err(|error| map_wire_error(error, target.path))?;
    let row_field = tile_view
        .fields()
        .filter(|field| field.number() == 5)
        .nth(row_occurrence)
        .ok_or(Error::InvalidSource { path: target.path })?;
    let row_info = numbers_table_cell_storage_codec::decode_tile_row_info(
        row_field.payload(),
        table_cell_decode_options(
            source,
            row_field.payload().len(),
            source.state.options.semantic().max_references(),
        ),
    )
    .map_err(|error| map_table_codec_error(error, target.path))?;
    // Package ingress intentionally falls back to the pre-BNC pair whenever
    // either modern row buffer is absent. Keep this route identical so a
    // comment read/edit cannot disagree with the semantic table projection.
    let (storage_buffer, offsets) = match (row_info.cell_storage_buffer(), row_info.cell_offsets())
    {
        (Some(storage), Some(offsets)) => (storage, offsets),
        _ => (
            row_info.cell_storage_buffer_pre_bnc(),
            row_info.cell_offsets_pre_bnc(),
        ),
    };
    let cell_range = cell_range(
        offsets,
        storage_buffer.len(),
        target.column,
        usize::try_from(row_info.cell_count())
            .map_err(|_| Error::InvalidSource { path: target.path })?,
        row_info.has_wide_offsets().unwrap_or(false),
        target.native.columns as usize,
        target.path,
    )?;
    let Some(cell_range) = cell_range else {
        return Ok(Located {
            target,
            comment: None,
            entry: None,
            storage: None,
            cell_bytes: None,
            comment_key: None,
            comment_table_id: None,
            replies: 0,
        });
    };
    let cell_source = storage_buffer
        .get(cell_range)
        .ok_or(Error::InvalidSource { path: target.path })?;
    let cell = crate::cell::wire::BncCell::parse(cell_source)
        .map_err(|_| Error::InvalidSource { path: target.path })?;
    let Some(comment_key) = cell.comment_identifier() else {
        let path = target.path;
        return Ok(Located {
            target,
            comment: None,
            entry: None,
            storage: None,
            cell_bytes: Some(copy_bytes(cell_source, path)?),
            comment_key: None,
            comment_table_id: None,
            replies: 0,
        });
    };
    if comment_key == 0 {
        return Err(Error::InvalidSource { path: target.path });
    }
    let comment_table_id = data_store
        .comment_storage_table()
        .map(|reference| reference.identifier())
        .ok_or(Error::InvalidSource { path: target.path })?;
    if comment_table_id == 0 {
        return Err(Error::InvalidSource { path: target.path });
    }
    let list_resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, comment_table_id)
        .map_err(|_| Error::InvalidSource { path: target.path })?
        .ok_or(Error::InvalidSource { path: target.path })?;
    let list_location = comment_table_entry(source, list_resolved, comment_key, target.path)?;
    let storage_resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, list_location.storage_id)
        .map_err(|_| Error::InvalidSource { path: target.path })?
        .ok_or(Error::InvalidSource { path: target.path })?;
    let storage_message_index = unique_message_index(
        storage_resolved.messages,
        COMMENT_STORAGE_MESSAGE_TYPE,
        target.path,
    )?;
    let storage_route = MessageRoute {
        component_index: storage_resolved.component_index,
        object_index: storage_resolved.object_index,
        message_index: storage_message_index,
        message_type: COMMENT_STORAGE_MESSAGE_TYPE,
    };
    let storage_message = &storage_resolved.messages[storage_message_index];
    let details = decode_comment_storage(source, storage_message.data.as_slice(), target.path)?;
    let entry = CommentEntryLocation {
        owner: list_location.owner,
        route: list_location.route,
        entry: list_location.entry,
        storage_occurrences: list_location.storage_occurrences,
    };
    let path = target.path;
    Ok(Located {
        target,
        comment: Some(details.comment),
        entry: Some(entry),
        storage: Some(storage_route),
        cell_bytes: Some(copy_bytes(cell_source, path)?),
        comment_key: Some(comment_key),
        comment_table_id: Some(comment_table_id),
        replies: details.replies,
    })
}

#[derive(Debug, Default)]
struct TileFinder {
    target: u32,
    identifier: Option<u64>,
    duplicate: bool,
}

impl numbers_table_cell_storage_codec::StorageVisitor for TileFinder {
    fn visit_tile_reference(
        &mut self,
        record: numbers_table_cell_storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if record.tile_id() != self.target {
            return Ok(());
        }
        if self
            .identifier
            .replace(record.reference().identifier())
            .is_some()
        {
            self.duplicate = true;
        }
        Ok(())
    }
}

#[derive(Debug, Default)]
struct RowFinder {
    target: u32,
    occurrence: usize,
    selected_occurrence: Option<usize>,
    duplicate: bool,
}

impl numbers_table_cell_storage_codec::StorageVisitor for RowFinder {
    fn visit_tile_row(
        &mut self,
        row: numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if row.tile_row_index() == self.target {
            if self.selected_occurrence.is_some() {
                self.duplicate = true;
            } else {
                self.selected_occurrence = Some(self.occurrence);
            }
        }
        self.occurrence = self.occurrence.saturating_add(1);
        Ok(())
    }
}

fn table_cell_decode_options(
    source: &Package,
    payload_bytes: usize,
    max_references: usize,
) -> numbers_table_cell_storage_codec::DecodeOptions {
    let fields = payload_bytes.clamp(1, WireLimits::MAX_FIELDS);
    let work = payload_bytes
        .saturating_mul(32)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    numbers_table_cell_storage_codec::DecodeOptions::new(
        payload_bytes.clamp(1, WireLimits::MAX_INPUT_BYTES),
        fields,
        work,
        COMMENT_MAX_NESTING,
        max_references.max(1),
        source
            .state
            .options
            .semantic()
            .max_output_text_bytes()
            .max(1),
    )
}

fn map_table_codec_error(
    error: numbers_table_cell_storage_codec::DecodeError,
    path: Path,
) -> Error {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            numbers_table_cell_storage_codec::DecodeLimit::Bytes { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireBytes,
                    observed,
                    maximum,
                    path,
                }
            },
            numbers_table_cell_storage_codec::DecodeLimit::References { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::References,
                    observed,
                    maximum,
                    path,
                }
            },
            numbers_table_cell_storage_codec::DecodeLimit::Text { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::TextBytes,
                    observed,
                    maximum,
                    path,
                }
            },
            numbers_table_cell_storage_codec::DecodeLimit::Fields { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireFields,
                    observed,
                    maximum,
                    path,
                }
            },
            numbers_table_cell_storage_codec::DecodeLimit::Work { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireWork,
                    observed,
                    maximum,
                    path,
                }
            },
            numbers_table_cell_storage_codec::DecodeLimit::Nesting { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireWork,
                    observed: usize::try_from(observed).unwrap_or(usize::MAX),
                    maximum: usize::try_from(maximum).unwrap_or(usize::MAX),
                    path,
                }
            },
            _ => Error::InvalidSource { path },
        };
    }
    if let Some(amount) = error.allocation_requested() {
        return Error::Allocation { amount, path };
    }
    Error::InvalidSource { path }
}

#[derive(Debug)]
struct CommentStorageDetails {
    comment: Comment,
    replies: usize,
    author_id: Option<u64>,
    storage_uuid: Option<(u64, u64)>,
    reply_ids: Vec<u64>,
}

#[derive(Debug, Default)]
struct ReplyIdCollector {
    ids: Vec<u64>,
    allocation_failed: Option<usize>,
}

impl comment_storage_codec::CommentStorageVisitor for ReplyIdCollector {
    fn visit_reply(
        &mut self,
        reply: comment_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), comment_storage_codec::DecodeError> {
        if self.ids.try_reserve(1).is_err() {
            self.allocation_failed = Some(self.ids.len().saturating_add(1));
        } else {
            self.ids.push(reply.identifier());
        }
        Ok(())
    }
}

fn decode_comment_storage(
    source: &Package,
    data: &[u8],
    path: Path,
) -> Result<CommentStorageDetails, Error> {
    let max_text = source.state.options.semantic().max_output_text_bytes();
    let options = comment_storage_codec::DecodeOptions::new(
        data.len().clamp(1, WireLimits::MAX_INPUT_BYTES),
        data.len().clamp(1, WireLimits::MAX_FIELDS),
        data.len()
            .saturating_mul(16)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        COMMENT_MAX_NESTING,
        source.state.options.semantic().max_references().max(1),
        max_text.max(1),
    );
    let mut replies = ReplyIdCollector::default();
    let (snapshot, report) = comment_storage_codec::decode_comment_storage_archive_with_visitor(
        data,
        options,
        &mut replies,
    )
    .map_err(|error| map_comment_codec_error(error, path))?;
    if let Some(amount) = replies.allocation_failed {
        return Err(Error::Allocation { amount, path });
    }
    let text = snapshot.text().unwrap_or_default();
    let comment = Comment::try_from_text(text, max_text, path)?;
    Ok(CommentStorageDetails {
        comment,
        replies: report.replies(),
        author_id: snapshot.author().map(|reference| reference.identifier()),
        storage_uuid: snapshot
            .storage_uuid()
            .map(|uuid| (uuid.lower(), uuid.upper())),
        reply_ids: replies.ids,
    })
}

fn map_comment_codec_error(error: comment_storage_codec::DecodeError, path: Path) -> Error {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            comment_storage_codec::DecodeLimit::Bytes { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireBytes,
                    observed,
                    maximum,
                    path,
                }
            },
            comment_storage_codec::DecodeLimit::References { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::References,
                    observed,
                    maximum,
                    path,
                }
            },
            comment_storage_codec::DecodeLimit::Text { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::TextBytes,
                    observed,
                    maximum,
                    path,
                }
            },
            comment_storage_codec::DecodeLimit::Fields { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireFields,
                    observed,
                    maximum,
                    path,
                }
            },
            comment_storage_codec::DecodeLimit::Work { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireWork,
                    observed,
                    maximum,
                    path,
                }
            },
            comment_storage_codec::DecodeLimit::Nesting { observed, maximum } => {
                Error::LimitExceeded {
                    kind: LimitKind::WireWork,
                    observed: usize::try_from(observed).unwrap_or(usize::MAX),
                    maximum: usize::try_from(maximum).unwrap_or(usize::MAX),
                    path,
                }
            },
            _ => Error::InvalidSource { path },
        };
    }
    Error::InvalidSource { path }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source_catalog = physical_source(edit.source)?;
    let source_owner = source_catalog.__source_owner();
    if edit.before == edit.after {
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                path: edit.target.path,
                before: edit.before,
                after: edit.after,
                source_cell: Arc::from([]),
                target_cell: Arc::from([]),
                source_previews: 0,
                target_previews: 0,
                touched_components: 0,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    // A changed clear is deliberately unsupported.  Do this before native
    // graph inspection or archive mutation so segmented/shared entries are
    // rejected uniformly and publication can never become a cell-only clear.
    if edit.before.is_some() && edit.after.is_none() {
        return Err(Error::UnsupportedDependency {
            path: edit.target.path,
        });
    }
    if !source_catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let located = resolve_comment_native(edit.source, edit.target.clone())?;
    if located.comment != edit.before {
        return Err(Error::InvalidSource {
            path: edit.target.path,
        });
    }
    if edit.after.is_some() && located.comment.is_none() {
        return Err(Error::UnsupportedDependency {
            path: edit.target.path,
        });
    }
    if edit.before.is_some() {
        let Some(entry) = located.entry.as_ref() else {
            return Err(Error::InvalidSource {
                path: edit.target.path,
            });
        };
        if !entry.owner.supports_text_rewrite()
            || entry.entry.refcount != 1
            || entry.storage_occurrences != 1
            || edit.after.is_none()
        {
            return Err(Error::UnsupportedDependency {
                path: edit.target.path,
            });
        }
        if !matches!(entry.owner, EntryOwner::Root) || located.replies != 0 {
            return Err(Error::UnsupportedDependency {
                path: edit.target.path,
            });
        }
        prove_global_comment_ownership(edit.source, &located)?;
    }
    let after = edit.after.clone();
    let source_cell = located.cell_bytes.clone().ok_or(Error::InvalidSource {
        path: edit.target.path,
    })?;
    let previews = root_preview_deletions(source_catalog)?;
    let (package, target_cell, touched, deleted_previews) =
        rewrite_existing(edit.source, &located, after.as_ref())?;
    let target_cell_for_patch = target_cell.clone();
    if package.table_cell_comment(
        SheetSelector::index(edit.target.path.sheet()),
        TableSelector::index(edit.target.path.table()),
        cell_position_for_path(edit.target.path)?,
    )? != after
    {
        return Err(Error::Verification);
    }
    let target_owner = physical_source(&package)?.__source_owner();
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            path: edit.target.path,
            before: edit.before,
            after,
            source_cell,
            target_cell: target_cell_for_patch,
            source_previews: previews.len(),
            target_previews: previews.len().saturating_sub(deleted_previews),
            touched_components: touched,
        },
        diagnostics: Diagnostics::published(touched, deleted_previews),
    })
}

fn rewrite_existing(
    source: &Package,
    located: &Located,
    after: Option<&Comment>,
) -> Result<(Package, Arc<[u8]>, usize, usize), Error> {
    let source_catalog = physical_source(source)?;
    if !source_catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let Some(comment) = after else {
        // This guard is intentionally duplicated at the transaction boundary
        // so this private helper can never grow a cell-only clear path by
        // accident in a later refactor.
        return Err(Error::UnsupportedDependency {
            path: located.target.path,
        });
    };
    let storage = located.storage.ok_or(Error::InvalidSource {
        path: located.target.path,
    })?;
    let original = message_at_route(source, storage, located.target.path)?
        .data
        .as_slice();
    let before = decode_comment_storage(source, original, located.target.path)?.comment;
    let cell = located.cell_bytes.as_ref().ok_or(Error::InvalidSource {
        path: located.target.path,
    })?;
    if before.text == comment.text {
        return Ok((source.snapshot(), Arc::clone(cell), 0, 0));
    }
    let data = patch_length_delimited_field_checked(
        original,
        1,
        !before.text.is_empty() || comment_field_present(original, 1)?,
        Some(comment.text.as_bytes()),
        located.target.path,
    )?;
    let verified = decode_comment_storage(source, &data, located.target.path)?.comment;
    if verified != *comment {
        return Err(Error::Verification);
    }
    let mut mutations = Vec::new();
    mutations.try_reserve(1).map_err(|_| Error::Allocation {
        amount: 1,
        path: located.target.path,
    })?;
    mutations.push(Mutation {
        route: storage,
        data,
    });
    let (bytes, touched) = rewrite_archives(source, &mut mutations, located.target.path)?;
    let deleted = root_preview_deletions(source_catalog)?;
    let edits = bytes
        .iter()
        .map(|(name, data)| EntryEdit::new(name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    let deleted_names = deleted.iter().map(String::as_str).collect::<Vec<_>>();
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, &deleted_names, source_catalog.limits())
        .map_err(|_| Error::InvalidSource {
            path: located.target.path,
        })?;
    let package = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(|_| Error::Verification)?;
    Ok((package, Arc::clone(cell), touched, deleted.len()))
}

fn rewrite_archives(
    source: &Package,
    mutations: &mut [Mutation],
    path: Path,
) -> Result<(Vec<(String, Vec<u8>)>, usize), Error> {
    mutations.sort_by_key(|mutation| {
        (
            mutation.route.component_index,
            mutation.route.object_index,
            mutation.route.message_index,
        )
    });
    let mut outputs = Vec::new();
    let mut touched = 0usize;
    let mut mutation_cursor = 0usize;
    while mutation_cursor < mutations.len() {
        let component_index = mutations[mutation_cursor].route.component_index;
        let component = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::InvalidSource { path })?;
        let name = component.name().to_owned();
        let entry = physical_source(source)
            .map_err(|_| Error::UnsupportedSource)?
            .package()
            .iter()
            .find(|entry| entry.name() == name)
            .ok_or(Error::InvalidSource { path })?;
        if entry.is_opaque() {
            return Err(Error::UnsupportedSource);
        }
        let physical = physical_source(source).map_err(|_| Error::UnsupportedSource)?;
        let archive_limits = physical
            .limits()
            .effective_archive_limits()
            .map_err(|_| Error::InvalidSource { path })?;
        let stream = SnappyStream::decompress_with_limits(
            entry.data(),
            physical
                .limits()
                .snappy_limits()
                .map_err(|_| Error::InvalidSource { path })?,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(|_| Error::InvalidSource { path })?;
        let start_mutations = mutation_cursor;
        while mutation_cursor < mutations.len()
            && mutations[mutation_cursor].route.component_index == component_index
        {
            let mutation = &mut mutations[mutation_cursor];
            let object = archive
                .objects
                .get_mut(mutation.route.object_index)
                .ok_or(Error::InvalidSource { path })?;
            if object.archive_info.identifier.is_none() {
                return Err(Error::InvalidSource { path });
            }
            object
                .replace_message_preserving_header_with_limits(
                    mutation.route.message_index,
                    RawMessage {
                        type_: mutation.route.message_type,
                        data: std::mem::take(&mut mutation.data),
                    },
                    archive_limits,
                )
                .map_err(|_| Error::InvalidSource { path })?;
            mutation_cursor += 1;
        }
        if mutation_cursor == start_mutations {
            return Err(Error::InvalidSource { path });
        }
        if archive.objects.is_empty() {
            return Err(Error::InvalidSource { path });
        }
        let rewritten = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(|_| Error::InvalidSource { path })?;
        let compressed =
            SnappyStream::compress(&rewritten).map_err(|_| Error::InvalidSource { path })?;
        outputs.try_reserve(1).map_err(|_| Error::Allocation {
            amount: outputs.len().saturating_add(1),
            path,
        })?;
        outputs.push((name, compressed));
        touched += 1;
    }
    Ok((outputs, touched))
}

fn patch_length_delimited_field_checked(
    source: &[u8],
    field_number: u32,
    expected_present: bool,
    replacement: Option<&[u8]>,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let replacement_bytes = replacement.map_or(0, <[u8]>::len);
    let _ = wire_limits_for(source.len(), replacement_bytes, path)?;
    let output = patch_length_delimited_field(source, field_number, expected_present, replacement)
        .map_err(|error| map_wire_error(error, path))?;
    if output.len() > WireLimits::MAX_OUTPUT_BYTES {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: output.len(),
            maximum: WireLimits::MAX_OUTPUT_BYTES,
            path,
        });
    }
    Ok(output)
}

fn wire_limits_for(
    source_bytes: usize,
    replacement_bytes: usize,
    path: Path,
) -> Result<WireLimits, Error> {
    let output = source_bytes
        .checked_add(replacement_bytes)
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
            observed: usize::MAX,
            maximum: WireLimits::MAX_OUTPUT_BYTES,
            path,
        })?;
    WireLimits::default()
        .with_input_bytes(source_bytes.clamp(1, WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| limits.with_fields(source_bytes.clamp(1, WireLimits::MAX_FIELDS)))
        .and_then(|limits| limits.with_output_bytes(output.clamp(1, WireLimits::MAX_OUTPUT_BYTES)))
        .and_then(|limits| limits.with_nesting(WireLimits::MAX_NESTING))
        .and_then(|limits| {
            limits.with_rewrite_work(
                source_bytes
                    .saturating_add(replacement_bytes)
                    .saturating_mul(4)
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(|error| map_wire_error(error, path))
}

fn map_wire_error(error: litchi_iwa_common::Error, path: Path) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::RewriteWork
                | litchi_iwa_common::LimitKind::Nesting
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireWork,
            },
            observed,
            maximum: limit,
            path,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount, path },
        _ => Error::InvalidSource { path },
    }
}

// The generated storage decoders charge their own wire work, but the native
// range and uniqueness scans below run in this adapter. Keep their cost on
// the same hard ceiling so malformed sparse rows cannot hide an unbounded
// amount of adapter-side work.
fn charge_scan_work(work: &mut usize, amount: usize, path: Path) -> Result<(), Error> {
    let observed = work.checked_add(amount).ok_or(Error::LimitExceeded {
        kind: LimitKind::WireWork,
        observed: usize::MAX,
        maximum: WireLimits::MAX_REWRITE_WORK,
        path,
    })?;
    if observed > WireLimits::MAX_REWRITE_WORK {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed,
            maximum: WireLimits::MAX_REWRITE_WORK,
            path,
        });
    }
    *work = observed;
    Ok(())
}

fn cell_ranges(
    offsets: &[u8],
    storage_length: usize,
    expected_cells: usize,
    wide: bool,
    column_count: usize,
    path: Path,
    scan_work: &mut usize,
) -> Result<Vec<Option<std::ops::Range<usize>>>, Error> {
    if !offsets.len().is_multiple_of(2)
        || column_count > offsets.len() / 2
        || expected_cells > offsets.len() / 2
        || expected_cells > column_count
    {
        return Err(Error::InvalidSource { path });
    }
    let slot_count = offsets.len() / 2;
    // Native producers may pad the offset table beyond the semantic table
    // width, but those slots must remain missing-column sentinels. Inspect
    // the complete table so a comment scan cannot silently accept a cell
    // outside the selected table's declared bounds.
    charge_scan_work(scan_work, slot_count, path)?;
    let width = if wide { 4 } else { 1 };
    let mut populated = 0usize;
    let mut ranges = Vec::new();
    ranges
        .try_reserve_exact(column_count)
        .map_err(|_| Error::Allocation {
            amount: column_count,
            path,
        })?;
    let mut previous = None;
    for (index, bytes) in offsets.chunks_exact(2).enumerate() {
        let raw = u16::from_le_bytes([bytes[0], bytes[1]]);
        if index >= column_count {
            if raw != u16::MAX {
                return Err(Error::InvalidSource { path });
            }
            continue;
        }
        if raw == u16::MAX {
            ranges.push(None);
            continue;
        }
        populated = populated
            .checked_add(1)
            .ok_or(Error::InvalidSource { path })?;
        let offset = usize::from(raw)
            .checked_mul(width)
            .ok_or(Error::InvalidSource { path })?;
        if offset >= storage_length {
            return Err(Error::InvalidSource { path });
        }
        if let Some((previous_index, previous_offset)) = previous {
            if previous_offset >= offset {
                return Err(Error::InvalidSource { path });
            }
            ranges[previous_index] = Some(previous_offset..offset);
        }
        previous = Some((index, offset));
        ranges.push(None);
    }
    if populated != expected_cells {
        return Err(Error::InvalidSource { path });
    }
    if let Some((previous_index, previous_offset)) = previous {
        if previous_offset >= storage_length {
            return Err(Error::InvalidSource { path });
        }
        ranges[previous_index] = Some(previous_offset..storage_length);
    }
    Ok(ranges)
}

fn cell_range(
    offsets: &[u8],
    storage_length: usize,
    column: usize,
    expected_cells: usize,
    wide: bool,
    column_count: usize,
    path: Path,
) -> Result<Option<std::ops::Range<usize>>, Error> {
    if column >= column_count {
        return Err(Error::InvalidSource { path });
    }
    let mut scan_work = 0;
    let ranges = cell_ranges(
        offsets,
        storage_length,
        expected_cells,
        wide,
        column_count,
        path,
        &mut scan_work,
    )?;
    ranges
        .get(column)
        .cloned()
        .ok_or(Error::InvalidSource { path })
}

fn copy_bytes(bytes: &[u8], path: Path) -> Result<Arc<[u8]>, Error> {
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(bytes.len())
        .map_err(|_| Error::Allocation {
            amount: bytes.len(),
            path,
        })?;
    retained.extend_from_slice(bytes);
    Ok(retained.into())
}

fn message_at_route(
    source: &Package,
    route: MessageRoute,
    path: Path,
) -> Result<&RawMessage, Error> {
    source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == route.message_type)
        .ok_or(Error::InvalidSource { path })
}

fn unique_message_index(
    messages: &[RawMessage],
    message_type: u32,
    path: Path,
) -> Result<usize, Error> {
    let mut index = None;
    for (candidate, message) in messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if index.replace(candidate).is_some() {
            return Err(Error::InvalidSource { path });
        }
    }
    index.ok_or(Error::InvalidSource { path })
}

#[derive(Debug)]
struct ListProbe {
    target: u32,
    target_storage: Option<u64>,
    target_entry: Option<EntryFact>,
    target_missing_storage: bool,
    target_key_occurrences: usize,
    target_storage_occurrences: usize,
    segment_ids: Vec<(u64, usize)>,
    storage_ids: Vec<u64>,
    entries: Vec<ListEntryFact>,
    segment_key_min: Option<u32>,
    segment_key_max: Option<u32>,
    allocation_failed: Option<usize>,
}

impl ListProbe {
    fn new(target: u32, target_storage: Option<u64>) -> Self {
        Self {
            target,
            target_storage,
            target_entry: None,
            target_missing_storage: false,
            target_key_occurrences: 0,
            target_storage_occurrences: 0,
            segment_ids: Vec::new(),
            storage_ids: Vec::new(),
            entries: Vec::new(),
            segment_key_min: None,
            segment_key_max: None,
            allocation_failed: None,
        }
    }

    fn push_bounded<T>(items: &mut Vec<T>, value: T, failed: &mut Option<usize>) {
        if failed.is_some() {
            return;
        }
        if items.try_reserve(1).is_err() {
            *failed = Some(items.len().saturating_add(1));
            return;
        }
        items.push(value);
    }
}

impl numbers_table_cell_storage_codec::StorageVisitor for ListProbe {
    fn visit_list_entry(
        &mut self,
        entry: numbers_table_cell_storage_codec::TableDataListEntrySnapshot<'_>,
    ) -> Result<(), numbers_table_cell_storage_codec::DecodeError> {
        if self.entries.try_reserve(1).is_err() {
            self.allocation_failed = Some(self.entries.len().saturating_add(1));
        } else {
            self.entries.push(ListEntryFact {
                key: entry.key(),
                refcount: entry.ref_count(),
                storage_id: entry
                    .comment_storage()
                    .map(|reference| reference.identifier()),
            });
        }
        match self.segment_key_min {
            Some(minimum) => self.segment_key_min = Some(minimum.min(entry.key())),
            None => self.segment_key_min = Some(entry.key()),
        }
        match self.segment_key_max {
            Some(maximum) => self.segment_key_max = Some(maximum.max(entry.key())),
            None => self.segment_key_max = Some(entry.key()),
        }
        if let Some(storage) = entry.comment_storage() {
            if self.target_storage == Some(storage.identifier()) {
                self.target_storage_occurrences = self.target_storage_occurrences.saturating_add(1);
            }
            Self::push_bounded(
                &mut self.storage_ids,
                storage.identifier(),
                &mut self.allocation_failed,
            );
        }
        if entry.key() != self.target {
            return Ok(());
        }
        self.target_key_occurrences = self.target_key_occurrences.saturating_add(1);
        if self.target_entry.is_some() {
            self.target_missing_storage = true;
            return Ok(());
        }
        let Some(storage) = entry.comment_storage() else {
            self.target_missing_storage = true;
            return Ok(());
        };
        self.target_entry = Some(EntryFact {
            refcount: entry.ref_count(),
            storage_id: storage.identifier(),
        });
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        reference: numbers_table_cell_storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), numbers_table_cell_storage_codec::DecodeError> {
        let occurrence = self.segment_ids.len();
        Self::push_bounded(
            &mut self.segment_ids,
            (reference.reference().identifier(), occurrence),
            &mut self.allocation_failed,
        );
        Ok(())
    }
}

fn list_decode_options(
    source: &Package,
    payload_bytes: usize,
) -> numbers_table_cell_storage_codec::DecodeOptions {
    table_cell_decode_options(
        source,
        payload_bytes,
        source.state.options.semantic().max_references(),
    )
}

fn decode_list_probe(
    source: &Package,
    data: &[u8],
    key: u32,
    path: Path,
    segment: bool,
    target_storage: Option<u64>,
) -> Result<(ListProbe, i32), Error> {
    let mut probe = ListProbe::new(key, target_storage);
    let options = list_decode_options(source, data.len());
    let (list_type, key_range) = if segment {
        let (snapshot, _report) =
            numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
                data, options, &mut probe,
            )
            .map_err(|error| map_table_codec_error(error, path))?;
        (
            snapshot.list_type(),
            Some((snapshot.key_range_location(), snapshot.key_range_length())),
        )
    } else {
        let (snapshot, _report) =
            numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
                data, options, &mut probe,
            )
            .map_err(|error| map_table_codec_error(error, path))?;
        (snapshot.list_type(), None)
    };
    if let Some(amount) = probe.allocation_failed {
        return Err(Error::Allocation { amount, path });
    }
    if let Some((location, length)) = key_range {
        validate_comment_segment_key_range(
            location,
            length,
            probe.segment_key_min,
            probe.segment_key_max,
            path,
        )?;
    }
    Ok((probe, list_type))
}

/// Validate a segmented comment-list envelope before the requested key is
/// projected or its storage payload is opened. The min/max pair proves every
/// entry visited by the strict decoder is inside the half-open key range.
fn validate_comment_segment_key_range(
    location: u32,
    length: u32,
    minimum: Option<u32>,
    maximum: Option<u32>,
    path: Path,
) -> Result<(), Error> {
    if length == 0 {
        return Err(Error::InvalidSource { path });
    }
    let end = location
        .checked_add(length)
        .ok_or(Error::InvalidSource { path })?;
    match (minimum, maximum) {
        (Some(minimum), Some(maximum)) => {
            if minimum > maximum || minimum < location || maximum >= end {
                return Err(Error::InvalidSource { path });
            }
        },
        (None, None) => {},
        _ => return Err(Error::InvalidSource { path }),
    }
    Ok(())
}

#[derive(Debug, Clone)]
struct ListLocation {
    owner: EntryOwner,
    route: MessageRoute,
    entry: EntryFact,
    storage_id: u64,
    storage_occurrences: usize,
}

#[derive(Debug, Clone, Copy)]
struct ListEntryFact {
    key: u32,
    refcount: u32,
    storage_id: Option<u64>,
}

fn comment_table_entry(
    source: &Package,
    resolved: super::Resolved<'_>,
    key: u32,
    path: Path,
) -> Result<ListLocation, Error> {
    let expected = tst::table_data_list::ListType::CommentStorage as i32;
    let mut selected: Option<(MessageRoute, ListProbe)> = None;
    for (message_index, message) in resolved.messages.iter().enumerate() {
        if !TABLE_DATA_LIST_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }
        let (probe, list_type) =
            decode_list_probe(source, message.data.as_slice(), key, path, false, None)?;
        if list_type != expected {
            continue;
        }
        if selected.is_some() {
            return Err(Error::InvalidSource { path });
        }
        selected = Some((
            MessageRoute {
                component_index: resolved.component_index,
                object_index: resolved.object_index,
                message_index,
                message_type: message.type_,
            },
            probe,
        ));
    }
    let (table_route, mut root_probe) = selected.ok_or(Error::InvalidSource { path })?;
    let mut selected_entry = root_probe
        .target_entry
        .map(|entry| (entry, EntryOwner::Root));
    let mut storage_ids = std::mem::take(&mut root_probe.storage_ids);
    for (segment_id, _segment_occurrence) in root_probe.segment_ids {
        let segment = source
            .state
            .index
            .resolve_ref_id(&source.state.components, segment_id)
            .map_err(|_| Error::InvalidSource { path })?
            .ok_or(Error::InvalidSource { path })?;
        let segment_message_index = unique_message_index(segment.messages, 6_011, path)?;
        let segment_message = &segment.messages[segment_message_index];
        let (segment_probe, segment_type) = decode_list_probe(
            source,
            segment_message.data.as_slice(),
            key,
            path,
            true,
            None,
        )?;
        if segment_type != expected {
            return Err(Error::InvalidSource { path });
        }
        storage_ids
            .try_reserve(segment_probe.storage_ids.len())
            .map_err(|_| Error::Allocation {
                amount: storage_ids
                    .len()
                    .saturating_add(segment_probe.storage_ids.len()),
                path,
            })?;
        storage_ids.extend(segment_probe.storage_ids);
        if segment_probe.target_missing_storage {
            return Err(Error::InvalidSource { path });
        }
        if let Some(entry) = segment_probe.target_entry {
            if selected_entry.is_some() {
                return Err(Error::InvalidSource { path });
            }
            selected_entry = Some((entry, EntryOwner::Segment));
        }
    }
    if root_probe.target_missing_storage {
        return Err(Error::InvalidSource { path });
    }
    let (entry, owner) = selected_entry.ok_or(Error::InvalidSource { path })?;
    if entry.refcount == 0 {
        return Err(Error::InvalidSource { path });
    }
    let storage_occurrences = storage_ids
        .into_iter()
        .filter(|identifier| *identifier == entry.storage_id)
        .count();
    if storage_occurrences == 0 {
        return Err(Error::InvalidSource { path });
    }
    Ok(ListLocation {
        owner,
        route: table_route,
        entry,
        storage_id: entry.storage_id,
        storage_occurrences,
    })
}

#[derive(Debug, Default)]
struct GlobalTileProbe {
    tiles: Vec<(u32, u64)>,
    allocation_failed: Option<usize>,
}

impl numbers_table_cell_storage_codec::StorageVisitor for GlobalTileProbe {
    fn visit_tile_reference(
        &mut self,
        record: numbers_table_cell_storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), numbers_table_cell_storage_codec::DecodeError> {
        ListProbe::push_bounded(
            &mut self.tiles,
            (record.tile_id(), record.reference().identifier()),
            &mut self.allocation_failed,
        );
        Ok(())
    }
}

#[derive(Debug, Default)]
struct GlobalRowProbe {
    rows: Vec<(u32, usize)>,
    allocation_failed: Option<usize>,
}

impl numbers_table_cell_storage_codec::StorageVisitor for GlobalRowProbe {
    fn visit_tile_row(
        &mut self,
        row: numbers_table_cell_storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), numbers_table_cell_storage_codec::DecodeError> {
        let occurrence = self.rows.len();
        ListProbe::push_bounded(
            &mut self.rows,
            (row.tile_row_index(), occurrence),
            &mut self.allocation_failed,
        );
        Ok(())
    }
}

fn scan_global_comment_cells(
    source: &Package,
    column_count: usize,
    storage: DecodedTableStorage<'_>,
    path: Path,
    scan_work: &mut usize,
) -> Result<Vec<u32>, Error> {
    let tile_size = usize::try_from(
        storage
            .tile_storage
            .tile_size()
            .unwrap_or(u32::try_from(DEFAULT_TILE_SIZE).unwrap_or(u32::MAX)),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    if tile_size == 0 {
        return Err(Error::InvalidSource { path });
    }
    let options = table_cell_decode_options(
        source,
        storage.data_store.tiles().len(),
        source.state.options.semantic().max_references(),
    );
    let mut tiles = GlobalTileProbe::default();
    let (_, _report) = numbers_table_cell_storage_codec::decode_tile_storage_with_visitor(
        storage.data_store.tiles(),
        options,
        &mut tiles,
    )
    .map_err(|error| map_table_codec_error(error, path))?;
    if let Some(amount) = tiles.allocation_failed {
        return Err(Error::Allocation { amount, path });
    }
    let mut seen_tile_keys = HashSet::new();
    seen_tile_keys
        .try_reserve(tiles.tiles.len())
        .map_err(|_| Error::Allocation {
            amount: tiles.tiles.len(),
            path,
        })?;
    let mut comment_keys = Vec::new();
    for (tile_key, tile_id) in tiles.tiles.iter().copied() {
        charge_scan_work(scan_work, 1, path)?;
        if tile_id == 0 || !seen_tile_keys.insert(tile_key) {
            return Err(Error::InvalidSource { path });
        }
        let tile_resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, tile_id)
            .map_err(|_| Error::InvalidSource { path })?
            .ok_or(Error::InvalidSource { path })?;
        let tile_route = MessageRoute {
            component_index: tile_resolved.component_index,
            object_index: tile_resolved.object_index,
            message_index: unique_message_index(tile_resolved.messages, TILE_MESSAGE_TYPE, path)?,
            message_type: TILE_MESSAGE_TYPE,
        };
        let tile_message = message_at_route(source, tile_route, path)?;
        let mut rows = GlobalRowProbe::default();
        let (_, _report) = numbers_table_cell_storage_codec::decode_tile_with_visitor(
            tile_message.data.as_slice(),
            table_cell_decode_options(
                source,
                tile_message.data.len(),
                source.state.options.semantic().max_references(),
            ),
            &mut rows,
        )
        .map_err(|error| map_table_codec_error(error, path))?;
        if let Some(amount) = rows.allocation_failed {
            return Err(Error::Allocation { amount, path });
        }
        let mut seen_row_keys = HashSet::new();
        seen_row_keys
            .try_reserve(rows.rows.len())
            .map_err(|_| Error::Allocation {
                amount: rows.rows.len(),
                path,
            })?;
        let tile_view = WireView::parse_with_limits(
            tile_message.data.as_slice(),
            wire_limits_for(tile_message.data.len(), 0, path)?,
        )
        .map_err(|error| map_wire_error(error, path))?;
        for (row_key, row_occurrence) in rows.rows.iter().copied() {
            charge_scan_work(scan_work, 1, path)?;
            if !seen_row_keys.insert(row_key) {
                return Err(Error::InvalidSource { path });
            }
            let row_field = tile_view
                .fields()
                .filter(|field| field.number() == 5)
                .nth(row_occurrence)
                .ok_or(Error::InvalidSource { path })?;
            let row_info = numbers_table_cell_storage_codec::decode_tile_row_info(
                row_field.payload(),
                table_cell_decode_options(
                    source,
                    row_field.payload().len(),
                    source.state.options.semantic().max_references(),
                ),
            )
            .map_err(|error| map_table_codec_error(error, path))?;
            let (storage_buffer, offsets) =
                match (row_info.cell_storage_buffer(), row_info.cell_offsets()) {
                    (Some(storage), Some(offsets)) => (storage, offsets),
                    _ => (
                        row_info.cell_storage_buffer_pre_bnc(),
                        row_info.cell_offsets_pre_bnc(),
                    ),
                };
            let expected_cells = usize::try_from(row_info.cell_count())
                .map_err(|_| Error::InvalidSource { path })?;
            let ranges = cell_ranges(
                offsets,
                storage_buffer.len(),
                expected_cells,
                row_info.has_wide_offsets().unwrap_or(false),
                column_count,
                path,
                scan_work,
            )?;
            charge_scan_work(scan_work, expected_cells, path)?;
            for cell_range in ranges.into_iter().flatten() {
                let cell_source = storage_buffer
                    .get(cell_range)
                    .ok_or(Error::InvalidSource { path })?;
                let cell = crate::cell::wire::BncCell::parse(cell_source)
                    .map_err(|_| Error::InvalidSource { path })?;
                let Some(comment_key) = cell.comment_identifier() else {
                    continue;
                };
                if comment_key == 0 {
                    return Err(Error::InvalidSource { path });
                }
                comment_keys.try_reserve(1).map_err(|_| Error::Allocation {
                    amount: comment_keys.len().saturating_add(1),
                    path,
                })?;
                comment_keys.push(comment_key);
            }
        }
    }
    Ok(comment_keys)
}

fn prove_global_comment_ownership(source: &Package, selected: &Located) -> Result<(), Error> {
    let entry = selected.entry.as_ref().ok_or(Error::InvalidSource {
        path: selected.target.path,
    })?;
    if !entry.owner.supports_text_rewrite()
        || entry.entry.refcount != 1
        || entry.storage_occurrences != 1
        || selected.replies != 0
    {
        return Err(Error::UnsupportedDependency {
            path: selected.target.path,
        });
    }
    let root_list = message_at_route(source, entry.route, selected.target.path)?;
    if !TABLE_DATA_LIST_MESSAGE_TYPES.contains(&root_list.type_) {
        return Err(Error::InvalidSource {
            path: selected.target.path,
        });
    }
    let (root_probe, root_type) = decode_list_probe(
        source,
        root_list.data.as_slice(),
        selected.comment_key.unwrap_or(0),
        selected.target.path,
        false,
        None,
    )?;
    if root_type != tst::table_data_list::ListType::CommentStorage as i32
        || root_probe.target_entry != Some(entry.entry)
    {
        return Err(Error::InvalidSource {
            path: selected.target.path,
        });
    }
    let census = census_comment_ownership(source, selected.target.path)?;
    let selected_key = selected.comment_key.ok_or(Error::InvalidSource {
        path: selected.target.path,
    })?;
    let selected_table_id = selected.comment_table_id.ok_or(Error::InvalidSource {
        path: selected.target.path,
    })?;
    let selected_storage_id = entry.entry.storage_id;
    if selected_key == 0 || selected_table_id == 0 || selected_storage_id == 0 {
        return Err(Error::InvalidSource {
            path: selected.target.path,
        });
    }
    let selected_list_occurrences = census
        .list_entries
        .iter()
        .filter(|fact| {
            fact.list_type == tst::table_data_list::ListType::CommentStorage as i32
                && fact.key == selected_key
        })
        .filter(|fact| match entry.owner {
            EntryOwner::Root => fact.table_id == selected_table_id && fact.root,
            EntryOwner::Segment => fact.storage_id == Some(selected_storage_id),
        })
        .count();
    if census
        .table_references
        .iter()
        .filter(|identifier| **identifier == selected_table_id)
        .count()
        != 1
        || census
            .list_entries
            .iter()
            .filter(|fact| fact.list_type == tst::table_data_list::ListType::CommentStorage as i32)
            .filter(|fact| fact.key == selected_key)
            .count()
            != 1
        || selected_list_occurrences != 1
        || census
            .list_entries
            .iter()
            .filter(|fact| fact.storage_id == Some(selected_storage_id))
            .count()
            != 1
        || census
            .cell_comment_keys
            .iter()
            .filter(|key| **key == selected_key)
            .count()
            != 1
    {
        return Err(Error::InvalidSource {
            path: selected.target.path,
        });
    }
    let storage = census
        .storages
        .iter()
        .find(|storage| storage.object_id == selected_storage_id)
        .ok_or(Error::InvalidSource {
            path: selected.target.path,
        })?;
    if storage.replies != 0
        || !storage.reply_ids.is_empty()
        || storage.storage_uuid.is_none()
        || census
            .entry_storage_ids
            .iter()
            .filter(|identifier| **identifier == selected_storage_id)
            .count()
            != 1
    {
        return Err(Error::UnsupportedDependency {
            path: selected.target.path,
        });
    }
    prove_archive_reference_ownership(source, selected_storage_id, selected.target.path)?;
    prove_metadata_ownership(source, selected_storage_id, selected.target.path)?;
    Ok(())
}

fn prove_metadata_ownership(
    source: &Package,
    deleted_object: u64,
    path: Path,
) -> Result<(), Error> {
    let route = super::metadata::unique_message_route(source)
        .ok_or(Error::UnsupportedDependency { path })?;
    let message = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .and_then(|component| component.archive().objects.get(route.object_index))
        .and_then(|object| object.messages.get(route.message_index))
        .filter(|message| message.type_ == super::metadata::MESSAGE_TYPE)
        .ok_or(Error::InvalidSource { path })?;
    let maximum_wire = source.state.options.archive().max_iwa_stream_bytes().min(
        source
            .state
            .options
            .archive()
            .archive_limits()
            .max_archive_bytes(),
    );
    if message.data.is_empty() || message.data.len() > maximum_wire {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: message.data.len(),
            maximum: maximum_wire,
            path,
        });
    }
    let fields = message
        .data
        .len()
        .saturating_mul(8)
        .clamp(1, WireLimits::MAX_FIELDS);
    let work = message
        .data
        .len()
        .saturating_mul(128)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let components = source
        .state
        .components
        .catalog()
        .len()
        .max(message.data.len())
        .max(1);
    let references = source.state.options.semantic().max_references().max(1);
    let options = MetadataRewriteOptions::new(
        maximum_wire,
        maximum_wire,
        fields,
        work,
        COMMENT_MAX_NESTING,
        components,
        references,
        0,
    );
    let mut census = super::metadata::ObjectOwnershipVisitor::new(deleted_object);
    inspect_package_metadata_with_visitor(message.data.as_slice(), options, &mut census)
        .map_err(|_| Error::UnsupportedDependency { path })?;
    if census.is_owned() {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(())
}

struct DeletedObjectReferenceCensus {
    deleted: u64,
    occurrences: usize,
}

impl ArchiveReferenceVisitor for DeletedObjectReferenceCensus {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        if occurrence.kind == ArchiveReferenceKind::Object
            && occurrence.referenced_identifier == self.deleted
        {
            self.occurrences = self.occurrences.saturating_add(1);
        }
        Ok(())
    }
}

fn prove_archive_reference_ownership(
    source: &Package,
    deleted_object: u64,
    path: Path,
) -> Result<(), Error> {
    let physical = physical_source(source).map_err(|_| Error::UnsupportedSource)?;
    let limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let mut census = DeletedObjectReferenceCensus {
        deleted: deleted_object,
        occurrences: 0,
    };
    for component in source.state.components.catalog().iter() {
        for object in &component.archive().objects {
            object
                .inspect_references_with_policy_and_limits(
                    &mut census,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    limits,
                )
                .map_err(|_| Error::UnsupportedDependency { path })?;
        }
    }
    if census.occurrences != 0 {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok(())
}

fn census_comment_ownership(source: &Package, path: Path) -> Result<CommentOwnershipCensus, Error> {
    let mut census = CommentOwnershipCensus::default();
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            let object_id = object
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource { path })?;
            if object_id == 0 {
                return Err(Error::InvalidSource { path });
            }
            for (message_index, message) in object.messages.iter().enumerate() {
                let route = MessageRoute {
                    component_index,
                    object_index,
                    message_index,
                    message_type: message.type_,
                };
                match message.type_ {
                    6_005 | 6_201 => {
                        let (probe, list_type) = decode_list_probe(
                            source,
                            message.data.as_slice(),
                            u32::MAX,
                            path,
                            false,
                            None,
                        )?;
                        let is_comment =
                            list_type == tst::table_data_list::ListType::CommentStorage as i32;
                        if is_comment {
                            let mut expected = probe.storage_ids.clone();
                            expected.extend(
                                probe.segment_ids.iter().map(|(identifier, _)| *identifier),
                            );
                            validate_comment_message_metadata(
                                object,
                                message_index,
                                &expected,
                                path,
                            )?;
                        } else {
                            validate_comment_message_metadata(object, message_index, &[], path)?;
                        }
                        if is_comment {
                            if census.comment_list_ids.contains(&object_id) {
                                return Err(Error::InvalidSource { path });
                            }
                            push_nonzero(&mut census.comment_list_ids, object_id, path)?;
                        }
                        append_list_facts(&mut census, object_id, list_type, &probe, true, path)?;
                        for (segment_position, (segment_id, _)) in
                            probe.segment_ids.iter().copied().enumerate()
                        {
                            if segment_id == 0
                                || (is_comment
                                    && (segment_id == object_id
                                        || probe.segment_ids[..segment_position]
                                            .iter()
                                            .any(|(prior, _)| *prior == segment_id)))
                            {
                                return Err(Error::InvalidSource { path });
                            }
                            let segment = source
                                .state
                                .index
                                .resolve_ref_id(&source.state.components, segment_id)
                                .map_err(|_| Error::InvalidSource { path })?
                                .ok_or(Error::InvalidSource { path })?;
                            let segment_message_index =
                                unique_message_index(segment.messages, 6_011, path)?;
                            let segment_message = message_at_route(
                                source,
                                MessageRoute {
                                    component_index: segment.component_index,
                                    object_index: segment.object_index,
                                    message_index: segment_message_index,
                                    message_type: 6_011,
                                },
                                path,
                            )?;
                            let segment_object = source
                                .state
                                .components
                                .catalog()
                                .get_index(segment.component_index)
                                .and_then(|component| {
                                    component.archive().objects.get(segment.object_index)
                                })
                                .ok_or(Error::InvalidSource { path })?;
                            if is_comment {
                                let segment_object_id = segment_object
                                    .archive_info
                                    .identifier
                                    .ok_or(Error::InvalidSource { path })?;
                                if segment_object_id == 0
                                    || census.rooted_segment_ids.contains(&segment_object_id)
                                {
                                    return Err(Error::InvalidSource { path });
                                }
                                push_nonzero(
                                    &mut census.rooted_segment_ids,
                                    segment_object_id,
                                    path,
                                )?;
                            }
                            let (segment_probe, segment_type) = decode_list_probe(
                                source,
                                segment_message.data.as_slice(),
                                u32::MAX,
                                path,
                                true,
                                None,
                            )?;
                            if segment_type != list_type {
                                return Err(Error::InvalidSource { path });
                            }
                            let expected = segment_probe.storage_ids.clone();
                            validate_comment_message_metadata(
                                segment_object,
                                segment_message_index,
                                &expected,
                                path,
                            )?;
                        }
                        let _ = route;
                    },
                    6_011 => {
                        if unique_message_index(object.messages.as_slice(), 6_011, path)?
                            != message_index
                        {
                            return Err(Error::InvalidSource { path });
                        }
                        let (probe, list_type) = decode_list_probe(
                            source,
                            message.data.as_slice(),
                            u32::MAX,
                            path,
                            true,
                            None,
                        )?;
                        let expected = probe.storage_ids.clone();
                        validate_comment_message_metadata(object, message_index, &expected, path)?;
                        append_list_facts(&mut census, object_id, list_type, &probe, false, path)?;
                    },
                    COMMENT_STORAGE_MESSAGE_TYPE => {
                        let details =
                            decode_comment_storage(source, message.data.as_slice(), path)?;
                        let mut expected = details.reply_ids.clone();
                        if let Some(author) = details.author_id {
                            expected.push(author);
                        }
                        validate_comment_message_metadata(object, message_index, &expected, path)?;
                        let fact = CommentStorageFact {
                            object_id,
                            author_id: details.author_id,
                            replies: details.replies,
                            reply_ids: details.reply_ids.clone(),
                            storage_uuid: details.storage_uuid,
                        };
                        if details.author_id == Some(0)
                            || details.storage_uuid == Some((0, 0))
                            || census
                                .storages
                                .iter()
                                .any(|item| item.object_id == object_id)
                        {
                            return Err(Error::InvalidSource { path });
                        }
                        census
                            .storages
                            .try_reserve(1)
                            .map_err(|_| Error::Allocation { amount: 1, path })?;
                        census.storages.push(fact);
                        if let Some(author) = details.author_id {
                            if author == object_id {
                                return Err(Error::InvalidSource { path });
                            }
                            push_nonzero(&mut census.author_ids, author, path)?;
                        }
                        for reply in &details.reply_ids {
                            if *reply == object_id {
                                return Err(Error::InvalidSource { path });
                            }
                            push_nonzero(&mut census.reply_ids, *reply, path)?;
                        }
                        if let Some(uuid) = details.storage_uuid {
                            census
                                .uuids
                                .try_reserve(1)
                                .map_err(|_| Error::Allocation { amount: 1, path })?;
                            census.uuids.push(uuid);
                        }
                        let _ = route;
                    },
                    _ => {},
                }
            }
        }
    }
    census_models_and_cells(source, &mut census, path)?;
    census_alias_checks(&census, path)?;
    Ok(census)
}

fn append_list_facts(
    census: &mut CommentOwnershipCensus,
    table_id: u64,
    list_type: i32,
    probe: &ListProbe,
    root: bool,
    path: Path,
) -> Result<(), Error> {
    if table_id == 0 {
        return Err(Error::InvalidSource { path });
    }
    for entry in &probe.entries {
        if list_type == tst::table_data_list::ListType::CommentStorage as i32 && entry.key == 0 {
            return Err(Error::InvalidSource { path });
        }
        if list_type == tst::table_data_list::ListType::CommentStorage as i32 && entry.refcount == 0
        {
            return Err(Error::InvalidSource { path });
        }
        if list_type == tst::table_data_list::ListType::CommentStorage as i32
            && entry.storage_id.is_none()
        {
            return Err(Error::InvalidSource { path });
        }
        if let Some(storage_id) = entry.storage_id {
            push_nonzero(&mut census.entry_storage_ids, storage_id, path)?;
        }
        census
            .list_entries
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1, path })?;
        census.list_entries.push(CommentListFact {
            table_id,
            list_type,
            key: entry.key,
            storage_id: entry.storage_id,
            root,
        });
    }
    Ok(())
}

fn push_nonzero(values: &mut Vec<u64>, value: u64, path: Path) -> Result<(), Error> {
    if value == 0 {
        return Err(Error::InvalidSource { path });
    }
    values.try_reserve(1).map_err(|_| Error::Allocation {
        amount: values.len().saturating_add(1),
        path,
    })?;
    values.push(value);
    Ok(())
}

fn census_alias_checks(census: &CommentOwnershipCensus, path: Path) -> Result<(), Error> {
    let mut storage_ids = Vec::new();
    storage_ids
        .try_reserve(census.storages.len())
        .map_err(|_| Error::Allocation {
            amount: census.storages.len(),
            path,
        })?;
    storage_ids.extend(census.storages.iter().map(|storage| storage.object_id));
    for author in &census.author_ids {
        if census.reply_ids.contains(author)
            || census.entry_storage_ids.contains(author)
            || storage_ids.contains(author)
        {
            return Err(Error::InvalidSource { path });
        }
    }
    for table in &census.table_references {
        if !census.comment_list_ids.contains(table) {
            return Err(Error::InvalidSource { path });
        }
    }
    for storage in &census.entry_storage_ids {
        if !storage_ids.contains(storage) {
            return Err(Error::InvalidSource { path });
        }
    }
    for storage in &census.storages {
        if storage.author_id == Some(storage.object_id)
            || storage
                .author_id
                .is_some_and(|author| census.reply_ids.contains(&author))
            || storage
                .reply_ids
                .iter()
                .any(|reply| *reply == storage.object_id)
            || storage.replies != storage.reply_ids.len()
        {
            return Err(Error::InvalidSource { path });
        }
    }
    for key in &census.cell_comment_keys {
        if !census.list_entries.iter().any(|entry| {
            entry.list_type == tst::table_data_list::ListType::CommentStorage as i32
                && entry.key == *key
        }) {
            return Err(Error::InvalidSource { path });
        }
    }
    for reply in &census.reply_ids {
        if census
            .reply_ids
            .iter()
            .filter(|candidate| *candidate == reply)
            .count()
            != 1
            || census.entry_storage_ids.contains(reply)
            || !storage_ids.contains(reply)
        {
            return Err(Error::InvalidSource { path });
        }
    }
    for uuid in &census.uuids {
        if *uuid == (0, 0)
            || census
                .uuids
                .iter()
                .filter(|candidate| *candidate == uuid)
                .count()
                != 1
        {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(())
}

fn validate_comment_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    message_index: usize,
    expected: &[u64],
    path: Path,
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource { path });
    }
    if !info.object_references.is_empty() {
        for reference in &info.object_references {
            if *reference == 0
                || info
                    .object_references
                    .iter()
                    .filter(|candidate| *candidate == reference)
                    .count()
                    != 1
                || (!expected.is_empty() && !expected.contains(reference))
            {
                return Err(Error::InvalidSource { path });
            }
        }
        if !expected.is_empty()
            && expected
                .iter()
                .any(|reference| !info.object_references.contains(reference))
        {
            return Err(Error::InvalidSource { path });
        }
    }
    for field in &info.field_infos {
        for reference in &field.object_references {
            if *reference == 0
                || (!expected.is_empty() && !expected.contains(reference))
                || (!info.object_references.is_empty()
                    && !info.object_references.contains(reference))
            {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    Ok(())
}

fn census_models_and_cells(
    source: &Package,
    census: &mut CommentOwnershipCensus,
    path: Path,
) -> Result<(), Error> {
    let mut scan_work = 0;
    for component in source.state.components.catalog().iter() {
        for object in component.archive().objects.iter() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if !matches!(message.type_, 6_000 | 6_001) {
                    continue;
                }
                if message.type_ == 6_001
                    && unique_message_index(object.messages.as_slice(), 6_001, path)?
                        != message_index
                {
                    return Err(Error::InvalidSource { path });
                }
                validate_comment_message_metadata(object, message_index, &[], path)?;
                let options = table_cell_decode_options(
                    source,
                    message.data.len(),
                    source.state.options.semantic().max_references(),
                );
                let model = match numbers_table_cell_storage_codec::decode_table_model_with_report(
                    message.data.as_slice(),
                    options,
                ) {
                    Ok((model, _report)) => model,
                    Err(error) if error.resource_limit().is_none() => {
                        match numbers_table_cell_storage_codec::decode_table_model_with_compatibility_data_store_with_report(
                            message.data.as_slice(),
                            options,
                        ) {
                            Ok((model, _report)) => model,
                            Err(_fallback) if message.type_ == 6_000 => continue,
                            Err(fallback) => {
                                return Err(map_table_codec_error(fallback, path));
                            },
                        }
                    },
                    Err(_error) if message.type_ == 6_000 => continue,
                    Err(error) => return Err(map_table_codec_error(error, path)),
                };
                let data_options = table_cell_decode_options(
                    source,
                    model.base_data_store().len(),
                    source.state.options.semantic().max_references(),
                );
                let store = match numbers_table_cell_storage_codec::decode_data_store_with_report(
                    model.base_data_store(),
                    data_options,
                ) {
                    Ok((store, _report)) => store,
                    Err(error) if error.resource_limit().is_none() => {
                        numbers_table_cell_storage_codec::decode_data_store_compatibility_with_report(
                            model.base_data_store(),
                            data_options,
                        )
                        .map(|(store, _report)| store)
                        .map_err(|error| map_table_codec_error(error, path))?
                    },
                    Err(error) => return Err(map_table_codec_error(error, path)),
                };
                let comment_table = store.comment_storage_table();
                if let Some(reference) = comment_table {
                    push_nonzero(&mut census.table_references, reference.identifier(), path)?;
                }
                let decoded = DecodedTableStorage {
                    data_store: store,
                    tile_storage:
                        numbers_table_cell_storage_codec::decode_tile_storage_with_report(
                            store.tiles(),
                            data_options,
                        )
                        .map_err(|_| Error::InvalidSource { path })?
                        .0,
                };
                let keys = scan_global_comment_cells(
                    source,
                    usize::try_from(model.number_of_columns())
                        .map_err(|_| Error::InvalidSource { path })?,
                    decoded,
                    path,
                    &mut scan_work,
                )?;
                if comment_table.is_none() && !keys.is_empty() {
                    return Err(Error::InvalidSource { path });
                }
                census
                    .cell_comment_keys
                    .try_reserve(keys.len())
                    .map_err(|_| Error::Allocation {
                        amount: keys.len(),
                        path,
                    })?;
                census.cell_comment_keys.extend(keys);
            }
        }
    }
    Ok(())
}

fn comment_field_present(source: &[u8], field_number: u32) -> Result<bool, Error> {
    let view =
        WireView::parse_with_limits(source, wire_limits_for(source.len(), 0, Path::Package)?)
            .map_err(|error| map_wire_error(error, Path::Package))?;
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    if count > 1 {
        return Err(Error::InvalidSource {
            path: Path::Package,
        });
    }
    Ok(count == 1)
}

fn physical_source(source: &Package) -> Result<&SourceCatalog, Error> {
    source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)
}

fn root_preview_deletions(source: &SourceCatalog) -> Result<Vec<String>, Error> {
    let mut deletions = Vec::new();
    for name in ROOT_PREVIEWS {
        let count = source
            .package()
            .iter()
            .filter(|entry| entry.name() == name)
            .count();
        match count {
            0 => {},
            1 => {
                deletions.try_reserve(1).map_err(|_| Error::Allocation {
                    amount: deletions.len().saturating_add(1),
                    path: Path::Package,
                })?;
                deletions.push(name.to_owned());
            },
            _ => {
                // Preview names are root-owned singleton entries. Treat a
                // duplicate as an ambiguous source rather than deleting an
                // arbitrary matching entry during publication.
                return Err(Error::InvalidSource {
                    path: Path::Package,
                });
            },
        }
    }
    Ok(deletions)
}

#[cfg(test)]
mod tests {
    use super::{Comment, Error, LimitKind, Path, WireLimits, cell_ranges};

    #[test]
    fn cell_ranges_precompute_sparse_offsets_in_one_pass() {
        let offsets = [0, 0, u8::MAX, u8::MAX, 3, 0, u8::MAX, u8::MAX];
        let mut scan_work = 0;
        let ranges = cell_ranges(&offsets, 8, 2, false, 4, Path::Package, &mut scan_work)
            .expect("sparse offsets should decode");

        assert_eq!(scan_work, 4);
        assert_eq!(ranges, vec![Some(0..3), None, Some(3..8), None]);
    }

    #[test]
    fn cell_ranges_charges_offset_scan_work() {
        let mut scan_work = WireLimits::MAX_REWRITE_WORK - 3;
        let error = cell_ranges(
            &[0, 0, 1, 0, 2, 0, 3, 0],
            4,
            4,
            false,
            4,
            Path::Package,
            &mut scan_work,
        )
        .expect_err("range scan should charge its work");

        assert_eq!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::WireWork,
                observed: WireLimits::MAX_REWRITE_WORK + 1,
                maximum: WireLimits::MAX_REWRITE_WORK,
                path: Path::Package,
            }
        );
    }

    #[test]
    fn cell_ranges_reject_descending_offsets() {
        let mut scan_work = 0;
        let error = cell_ranges(&[3, 0, 1, 0], 4, 2, false, 2, Path::Package, &mut scan_work)
            .expect_err("offsets must be strictly increasing");

        assert!(matches!(
            error,
            Error::InvalidSource {
                path: Path::Package
            }
        ));
    }

    #[test]
    fn cell_ranges_reject_non_sentinel_padding() {
        let mut scan_work = 0;
        let error = cell_ranges(
            &[0, 0, u8::MAX, u8::MAX, 1, 0],
            4,
            1,
            false,
            2,
            Path::Package,
            &mut scan_work,
        )
        .expect_err("padded slots must remain missing-column sentinels");

        assert!(matches!(
            error,
            Error::InvalidSource {
                path: Path::Package
            }
        ));
        assert_eq!(scan_work, 3);
    }

    #[test]
    fn cell_ranges_accept_sentinel_padding() {
        let mut scan_work = 0;
        let ranges = cell_ranges(
            &[0, 0, 3, 0, u8::MAX, u8::MAX, u8::MAX, u8::MAX],
            8,
            2,
            false,
            2,
            Path::Package,
            &mut scan_work,
        )
        .expect("sentinel padding should be accepted");

        assert_eq!(ranges, vec![Some(0..3), Some(3..8)]);
        assert_eq!(scan_work, 4);
    }

    #[test]
    fn comment_clone_shares_arc_backed_text() {
        let original = Comment::new("shared text");
        let cloned = original.clone();

        assert!(std::sync::Arc::ptr_eq(&original.text, &cloned.text));
        assert_eq!(cloned.text(), "shared text");
    }

    #[test]
    fn comment_try_new_copies_bounded_text() {
        let comment = Comment::try_new("fallible text").expect("bounded comment allocation");

        assert_eq!(comment.text(), "fallible text");
    }

    #[test]
    fn comment_try_from_text_rejects_selected_limit_before_allocation() {
        let error = Comment::try_from_text("too long", 3, Path::Package)
            .expect_err("the selected text ceiling must be enforced");

        assert_eq!(
            error,
            Error::LimitExceeded {
                kind: LimitKind::TextBytes,
                observed: 8,
                maximum: 3,
                path: Path::Package,
            }
        );
    }
}
