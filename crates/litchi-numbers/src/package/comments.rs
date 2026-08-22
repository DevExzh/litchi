//! Selector-first semantic comments for rooted Numbers table cells.
//!
//! This module is deliberately separate from the scalar-cell and formula
//! editors.  A cell comment is a format-owned annotation graph, not a cell
//! value.  The public boundary therefore exposes only a small text-bearing
//! semantic value; native object identifiers, table-list keys, protobuf
//! messages, and archive member names remain private to this adapter.
//!
//! The exact-source write seam is intentionally narrow: it can replace the
//! text of an existing, unshared comment whose table-list entry is rooted in
//! either the table object or one of its bounded list segments. Creating
//! comments and clearing comments are refused; both operations require
//! ownership-graph mutations that this adapter does not perform.

use std::{fmt, sync::Arc};

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{WireView, patch_length_delimited_field};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{comment_storage_codec, numbers_table_cell_storage_codec, tst};
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
    text: String,
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
        Self { text: text.into() }
    }

    /// Borrow the comment text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
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
        Ok(Self { text: retained })
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
    entry: EntryFact,
    storage_occurrences: usize,
}

#[derive(Debug, Clone, Copy)]
enum EntryOwner {
    Root,
    Segment,
}

impl EntryOwner {
    const fn supports_text_rewrite(self) -> bool {
        matches!(self, Self::Root | Self::Segment)
    }
}

#[derive(Debug, Clone, Copy)]
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
    let table_position = match table.into() {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => selected_sheet
            .tables()
            .enumerate()
            .find(|(_, candidate)| candidate.name() == name)
            .map(|(index, _)| index),
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

fn resolve_comment_native(source: &Package, mut target: Target) -> Result<Located, Error> {
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
        data_store_options,
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
        });
    };
    let comment_table_id = data_store
        .comment_storage_table()
        .map(|reference| reference.identifier())
        .ok_or(Error::InvalidSource { path: target.path })?;
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
    let snapshot = comment_storage_codec::decode_comment_storage_archive(data, options)
        .map_err(|error| map_comment_codec_error(error, path))?;
    let text = snapshot.text().unwrap_or_default();
    let comment = Comment::try_from_text(text, max_text, path)?;
    Ok(CommentStorageDetails { comment })
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
    }
    let maximum = edit.source.state.options.semantic().max_output_text_bytes();
    let after = edit
        .after
        .as_ref()
        .map(|comment| Comment::try_from_text(comment.text.as_str(), maximum, edit.target.path))
        .transpose()?;
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

fn cell_range(
    offsets: &[u8],
    storage_length: usize,
    column: usize,
    expected_cells: usize,
    wide: bool,
    column_count: usize,
    path: Path,
) -> Result<Option<std::ops::Range<usize>>, Error> {
    if !offsets.len().is_multiple_of(2)
        || column >= column_count
        || column_count > offsets.len() / 2
        || expected_cells > offsets.len() / 2
    {
        return Err(Error::InvalidSource { path });
    }
    let width = if wide { 4 } else { 1 };
    let mut populated = 0usize;
    let mut selected = None;
    let mut end = None;
    for (index, bytes) in offsets.chunks_exact(2).take(column_count).enumerate() {
        let raw = u16::from_le_bytes([bytes[0], bytes[1]]);
        if raw == u16::MAX {
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
        if index == column {
            selected = Some(offset);
        } else if selected.is_some() && end.is_none() {
            end = Some(offset);
        }
    }
    if populated != expected_cells {
        return Err(Error::InvalidSource { path });
    }
    let Some(start) = selected else {
        return Ok(None);
    };
    let end = end.unwrap_or(storage_length);
    (start < end)
        .then_some(start..end)
        .ok_or(Error::InvalidSource { path })
        .map(Some)
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
    target_entry: Option<EntryFact>,
    target_missing_storage: bool,
    segment_ids: Vec<(u64, usize)>,
    storage_ids: Vec<u64>,
    allocation_failed: Option<usize>,
}

impl ListProbe {
    fn new(target: u32) -> Self {
        Self {
            target,
            target_entry: None,
            target_missing_storage: false,
            segment_ids: Vec::new(),
            storage_ids: Vec::new(),
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
        if let Some(storage) = entry.comment_storage() {
            Self::push_bounded(
                &mut self.storage_ids,
                storage.identifier(),
                &mut self.allocation_failed,
            );
        }
        if entry.key() != self.target {
            return Ok(());
        }
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
) -> Result<(ListProbe, i32), Error> {
    let mut probe = ListProbe::new(key);
    let options = list_decode_options(source, data.len());
    let (snapshot, _report) = if segment {
        numbers_table_cell_storage_codec::decode_table_data_list_segment_with_visitor(
            data, options, &mut probe,
        )
        .map(|(snapshot, report)| (snapshot.list_type(), report))
    } else {
        numbers_table_cell_storage_codec::decode_table_data_list_with_visitor(
            data, options, &mut probe,
        )
        .map(|(snapshot, report)| (snapshot.list_type(), report))
    }
    .map_err(|error| map_table_codec_error(error, path))?;
    if let Some(amount) = probe.allocation_failed {
        return Err(Error::Allocation { amount, path });
    }
    Ok((probe, snapshot))
}

#[derive(Debug, Clone)]
struct ListLocation {
    owner: EntryOwner,
    entry: EntryFact,
    storage_id: u64,
    storage_occurrences: usize,
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
            decode_list_probe(source, message.data.as_slice(), key, path, false)?;
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
    let (_table_route, mut root_probe) = selected.ok_or(Error::InvalidSource { path })?;
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
        let (segment_probe, segment_type) =
            decode_list_probe(source, segment_message.data.as_slice(), key, path, true)?;
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
        entry,
        storage_id: entry.storage_id,
        storage_occurrences,
    })
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
