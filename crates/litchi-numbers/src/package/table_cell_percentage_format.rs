//! Selector-first transactions for an existing table cell's decimal
//! Percentage format.
//!
//! The public values in this module are archive-free. Native BNC records,
//! format-table keys, IWA members, and strict wire views remain below the
//! package boundary. The native implementation is shared with the focused
//! display-format owner so graph routing, bounded staging, and candidate
//! locality checks stay in one place.

use std::fmt;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_core::SnappyStream;
use thiserror::Error as ThisError;

use super::{
    Package,
    physical_entry_index::{Error as PhysicalEntryIndexError, PhysicalEntryIndex},
    table_cell_control_native as control_native, table_cell_display_format_native as native,
    table_cell_pop_up_menu as popup,
};
use crate::{CellPosition, SheetSelector, TableSelector, cell::data_format::Percentage};

/// A content-free location associated with a Percentage-format operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One selected rooted table cell.
    Cell {
        /// Zero-based rooted sheet position.
        sheet: usize,
        /// Zero-based table position within the sheet.
        table: usize,
        /// Zero-based semantic cell coordinate.
        position: CellPosition,
    },
}

/// A finite resource governed by a Percentage-format transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical entry bytes.
    TotalEntryBytes,
    /// Physical package/container metadata bytes.
    PackageBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing/items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict codec input bytes.
    WireBytes,
    /// Strict codec output bytes.
    WireOutputBytes,
    /// Strict codec reference-envelope bytes.
    WireReferenceBytes,
    /// Strict codec selected-text bytes.
    WireTextBytes,
    /// Strict codec fields inspected.
    WireFields,
    /// Strict codec nesting depth.
    WireNesting,
    /// Strict codec work.
    WireWork,
    /// Private codec/reassembly scratch bytes.
    ScratchBytes,
    /// Private candidate bytes retained during execution.
    RetainedBytes,
    /// Bounded allocation units reserved by a phase.
    Allocations,
    /// Compressed physical member bytes.
    CompressedBytes,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted Percentage-format transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// The selected semantic cell was not present in the rooted table.
    #[error("the selected Numbers table cell was not found")]
    CellNotFound,
    /// A changed operation targeted an effectively locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    /// The selected cell has an explicit non-Percentage display/control family.
    #[error("the selected Numbers cell does not contain an explicit Percentage format at {path:?}")]
    WrongFormatFamily { path: Path },
    /// A native dependency is not safe to rewrite in this owner.
    #[error("the selected Numbers Percentage format has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: Path },
    /// The exact native owner is not available for this source.
    #[error("this Numbers source does not support exact Percentage-format editing")]
    UnsupportedSource,
    /// Rooted ownership or wire framing is invalid.
    #[error("the Numbers Percentage-format source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Numbers Percentage format {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free failure location.
        path: Path,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Numbers Percentage-format transaction")]
    Allocation { amount: usize, path: Path },
    /// Candidate reopening or semantic locality verification failed.
    #[error("the edited Numbers Percentage format failed semantic verification")]
    Verification,
    /// A patch was created for another exact package artifact.
    #[error("the Percentage-format patch does not match the exact source package")]
    PatchConflict,
}

impl From<popup::Error> for Error {
    fn from(error: popup::Error) -> Self {
        map_popup_error(error)
    }
}

/// A selector-first Percentage-format edit staged against one exact package
/// snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    path: Path,
    before: Option<Percentage>,
    after: Option<Percentage>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected semantic cell path.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the staged Percentage format, if the cell has an explicit one.
    #[must_use]
    pub const fn percentage(&self) -> Option<&Percentage> {
        self.after.as_ref()
    }

    /// Alias useful to generic data-format callers.
    #[must_use]
    pub const fn format(&self) -> Option<&Percentage> {
        self.percentage()
    }

    /// Return the value observed when this edit was opened.
    #[must_use]
    pub const fn before(&self) -> Option<&Percentage> {
        self.before.as_ref()
    }

    /// Return the currently staged semantic value.
    #[must_use]
    pub const fn after(&self) -> Option<&Percentage> {
        self.after.as_ref()
    }

    /// Stage a Percentage format without changing package bytes.
    #[must_use]
    pub fn set(mut self, value: Percentage) -> Self {
        self.after = Some(value);
        self
    }

    /// Stage the inherited/automatic (no explicit Percentage) state.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    /// Alias for [`Self::clear`].
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    /// Validate and publish this transaction.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.before == self.after {
            return no_op_commit(self.source, self.path, self.before, self.after);
        }
        rewrite(self.source, self.path, self.before, self.after)
    }
}

/// A reversible exact-source Percentage-format patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Option<Percentage>,
    after: Option<Percentage>,
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
    /// Return the selected path.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the exact source semantic value.
    #[must_use]
    pub const fn before(&self) -> Option<&Percentage> {
        self.before.as_ref()
    }

    /// Return the exact target semantic value.
    #[must_use]
    pub const fn after(&self) -> Option<&Percentage> {
        self.after.as_ref()
    }

    /// Return the target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after,
            after: self.before,
        }
    }

    /// Whether this patch is both semantically and byte-wise unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
}

/// Content-free diagnostics for one Percentage-format publication.
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

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of changed physical components/members.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of deleted preview members (always zero for Percentage formats).
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the private candidate was reopened and read back.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One validated package publication.
#[must_use = "a Percentage-format commit contains the validated package snapshot"]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("patch", &self.patch)
            .field("diagnostics", &self.diagnostics)
            .finish_non_exhaustive()
    }
}

impl Commit {
    /// Borrow the candidate package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one existing cell's explicit Percentage format.
    ///
    /// `None` denotes the inherited/no-explicit-Percentage state. An
    /// explicit Number, currency, date, duration, text, or interactive format
    /// is a typed [`Error::WrongFormatFamily`] rather than an automatic value.
    pub fn table_cell_percentage_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<Percentage>, Error> {
        let target = popup::resolve_cell(self, sheet, table, position).map_err(map_popup_error)?;
        let path = Path::Cell {
            sheet: target.sheet_position,
            table: target.table_position,
            position,
        };
        native::read_percentage_format(self, target, native_path(path))
            .map_err(|error| map_percentage_read_error(error, path))
    }

    /// Start a selector-first Percentage-format edit for one existing cell.
    pub fn edit_table_cell_percentage_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Edit<'_>, Error> {
        let target = popup::resolve_cell(self, sheet, table, position).map_err(map_popup_error)?;
        let path = Path::Cell {
            sheet: target.sheet_position,
            table: target.table_position,
            position,
        };
        let before = native::read_percentage_format(self, target, native_path(path))
            .map_err(|error| map_percentage_read_error(error, path))?;
        Ok(Edit {
            source: self,
            path,
            before,
            after: before,
        })
    }

    /// Apply a previously-created exact Percentage-format patch.
    pub fn apply_table_cell_percentage_format(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = super::table_headers::rewrite::physical_source(self)
            .map_err(|_| Error::UnsupportedSource)?;
        let owner = catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&owner) {
            return Err(Error::PatchConflict);
        }
        let Path::Cell {
            sheet,
            table,
            position,
        } = patch.path
        else {
            return Err(Error::PatchConflict);
        };
        let current = self.table_cell_percentage_format(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        if current != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }

        let target_owner = patch.artifacts.target_owner();
        let mut budget = popup::TransactionBudget::for_cell_control(self);
        budget.charge_package_source(catalog, native_path(patch.path))?;
        budget.charge_output(target_owner.as_ref().len(), native_path(patch.path))?;
        budget.charge_allocations(2, native_path(patch.path))?;
        budget.charge_transaction_work(
            target_owner
                .as_ref()
                .len()
                .saturating_mul(2)
                .saturating_add(catalog.package().iter().count().saturating_mul(1024)),
            native_path(patch.path),
        )?;
        budget
            .charge_candidate_input_bytes(target_owner.as_ref().len(), native_path(patch.path))?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
            .map_err(|_| Error::Verification)?;
        budget.charge_candidate_reopen(candidate_catalog, native_path(patch.path))?;
        let candidate_target = popup::resolve_cell(
            &candidate,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )
        .map_err(map_popup_error)?;
        let after = native::read_percentage_format_with_budget(
            &candidate,
            candidate_target,
            native_path(patch.path),
            &mut budget,
        )
        .map_err(|error| map_percentage_read_error(error, patch.path))?;
        if after != patch.after {
            return Err(Error::Verification);
        }
        let touched_components = verify_percentage_package_locality(self, &candidate)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components,
                deleted_previews: 0,
                full_reparse_performed: true,
            },
        })
    }
}

fn no_op_commit(
    source: &Package,
    path: Path,
    before: Option<Percentage>,
    after: Option<Percentage>,
) -> Result<Commit, Error> {
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::UnsupportedSource)?;
    let owner = catalog.__source_owner();
    Ok(Commit {
        package: source.snapshot(),
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
            path,
            before,
            after,
        },
        diagnostics: Diagnostics::unchanged(),
    })
}

fn rewrite(
    source: &Package,
    path: Path,
    before: Option<Percentage>,
    after: Option<Percentage>,
) -> Result<Commit, Error> {
    let Path::Cell {
        sheet,
        table,
        position,
    } = path
    else {
        return Err(Error::InvalidSource { path });
    };
    let source_catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::UnsupportedSource)?;
    if !source_catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let source_owner = source_catalog.__source_owner();
    let mut budget = popup::TransactionBudget::for_cell_control(source);
    budget.charge_package_source(source_catalog, native_path(path))?;
    let target = popup::resolve_cell(
        source,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )
    .map_err(map_popup_error)?;
    let current =
        native::read_percentage_format_with_budget(source, target, native_path(path), &mut budget)
            .map_err(|error| map_percentage_read_error(error, path))?;
    if current != before {
        return Err(Error::PatchConflict);
    }
    let native_bound = control_native::preflight_copy_on_write_scalar_control(
        source,
        target,
        native_path(path),
        &mut budget,
    )
    .map_err(map_popup_error)?;
    let native_output = native::rewrite_percentage_format(
        source,
        target,
        before.as_ref(),
        after.as_ref(),
        native_path(path),
        &mut budget,
    )
    .map_err(map_popup_error)?;
    let member_count = native_output.members.len();
    if !(1..=2).contains(&member_count) {
        return Err(Error::Verification);
    }
    let mut maximum_compressed_bytes = 0usize;
    let mut compression_allocations = 3usize;
    for member in &native_output.members {
        if member.archive_bytes.len() > native_bound.archive_bytes {
            return Err(Error::LimitExceeded {
                kind: LimitKind::PayloadBytes,
                observed: u64::try_from(member.archive_bytes.len()).unwrap_or(u64::MAX),
                maximum: u64::try_from(native_bound.archive_bytes).unwrap_or(u64::MAX),
                path,
            });
        }
        let maximum_compressed = SnappyStream::maximum_compressed_len(member.archive_bytes.len())
            .map_err(|_| Error::Verification)?;
        if maximum_compressed > native_bound.compressed_bytes {
            return Err(Error::LimitExceeded {
                kind: LimitKind::CompressedBytes,
                observed: u64::try_from(maximum_compressed).unwrap_or(u64::MAX),
                maximum: u64::try_from(native_bound.compressed_bytes).unwrap_or(u64::MAX),
                path,
            });
        }
        maximum_compressed_bytes = maximum_compressed_bytes
            .checked_add(maximum_compressed)
            .ok_or(Error::Verification)?;
        let frames = member
            .archive_bytes
            .len()
            .checked_add(SnappyStream::WRITE_CHUNK_SIZE - 1)
            .ok_or(Error::Verification)?
            / SnappyStream::WRITE_CHUNK_SIZE;
        compression_allocations = compression_allocations
            .checked_add(frames)
            .and_then(|amount| amount.checked_add(1))
            .ok_or(Error::Verification)?;
    }

    // Percentage-format edits do not manufacture UUIDs, save tokens, or
    // previews. Only the tile and shared format-list members returned by the
    // native owner enter reassembly; metadata and every preview remain exact.
    budget
        .charge_compressed_bytes(maximum_compressed_bytes, native_path(path))
        .map_err(map_popup_error)?;
    budget
        .charge_allocations(compression_allocations, native_path(path))
        .map_err(map_popup_error)?;
    budget
        .charge_scratch_bytes(maximum_compressed_bytes, native_path(path))
        .map_err(map_popup_error)?;
    budget
        .charge_retained_bytes(maximum_compressed_bytes, native_path(path))
        .map_err(map_popup_error)?;
    let compression_work = source_catalog
        .source_bytes()
        .len()
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(maximum_compressed_bytes))
        .ok_or(Error::Verification)?;
    budget
        .charge_transaction_work(compression_work, native_path(path))
        .map_err(map_popup_error)?;
    let mut compressed_members = Vec::new();
    compressed_members
        .try_reserve_exact(member_count)
        .map_err(|_| Error::Allocation {
            amount: member_count,
            path,
        })?;
    for member in &native_output.members {
        let compressed =
            SnappyStream::compress(&member.archive_bytes).map_err(|_| Error::Verification)?;
        let maximum = SnappyStream::maximum_compressed_len(member.archive_bytes.len())
            .map_err(|_| Error::Verification)?;
        if compressed.len() > maximum {
            return Err(Error::Verification);
        }
        compressed_members.push(compressed);
    }
    let mut changed_names = Vec::new();
    changed_names
        .try_reserve_exact(member_count)
        .map_err(|_| Error::Allocation {
            amount: member_count,
            path,
        })?;
    changed_names.extend(
        native_output
            .members
            .iter()
            .map(|member| member.member_name.as_str()),
    );
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(member_count)
        .map_err(|_| Error::Allocation {
            amount: member_count,
            path,
        })?;
    edits.extend(
        native_output
            .members
            .iter()
            .zip(&compressed_members)
            .map(|(member, bytes)| EntryEdit::new(&member.member_name, bytes)),
    );
    // The reassembly planner owns two small edit-index collections. Debit
    // them before entering the planner; its exact byte/work requirements are
    // then checked before execution below.
    budget
        .charge_allocations(2, native_path(path))
        .map_err(map_popup_error)?;
    let prepared = source_catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], source_catalog.limits())
        .map_err(|_| Error::Verification)?;
    let requirements = prepared.execution_requirements();
    budget
        .preflight_reassembly(requirements, native_path(path))
        .map_err(map_popup_error)?;
    let publication_work = requirements
        .output_bytes()
        .checked_mul(2)
        .and_then(|amount| {
            source_catalog
                .package()
                .iter()
                .count()
                .checked_mul(1024)
                .and_then(|catalog_work| amount.checked_add(catalog_work))
        })
        .ok_or(Error::Verification)?;
    budget
        .charge_transaction_work(publication_work, native_path(path))
        .map_err(map_popup_error)?;
    let limits = requirements.exact_limits();
    budget
        .charge_candidate_input_bytes(requirements.output_bytes(), native_path(path))
        .map_err(map_popup_error)?;
    let bytes = prepared.execute(limits).map_err(|_| Error::Verification)?;
    let candidate = Package::from_owned_bytes_with_options(bytes, source.state.options)
        .map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?;
    budget
        .charge_candidate_reopen(candidate_catalog, native_path(path))
        .map_err(map_popup_error)?;
    let candidate_target = popup::resolve_cell(
        &candidate,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )
    .map_err(map_popup_error)?;
    let candidate_value = native::read_percentage_format_with_budget(
        &candidate,
        candidate_target,
        native_path(path),
        &mut budget,
    )
    .map_err(|error| map_percentage_read_error(error, path))?;
    if candidate_value != after {
        return Err(Error::Verification);
    }
    let touched_components =
        verify_percentage_package_locality_with_members(source, &candidate, &changed_names)?;
    let target_owner = candidate_catalog.__source_owner();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            path,
            before,
            after,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components,
            deleted_previews: 0,
            full_reparse_performed: true,
        },
    })
}

fn native_path(path: Path) -> popup::Path {
    match path {
        Path::Package => popup::Path::Package,
        Path::Cell {
            sheet,
            table,
            position,
        } => popup::Path::Cell {
            sheet,
            table,
            position,
        },
    }
}

fn from_native_path(path: popup::Path) -> Path {
    match path {
        popup::Path::Package => Path::Package,
        popup::Path::Cell {
            sheet,
            table,
            position,
        } => Path::Cell {
            sheet,
            table,
            position,
        },
    }
}

fn map_percentage_read_error(error: native::PercentageFormatReadError, path: Path) -> Error {
    match error {
        native::PercentageFormatReadError::WrongFormatFamily => Error::WrongFormatFamily { path },
        native::PercentageFormatReadError::Native(error) => map_popup_error(error),
    }
}

fn map_popup_error(error: popup::Error) -> Error {
    match error {
        popup::Error::SheetNotFound => Error::SheetNotFound,
        popup::Error::TableNotFound => Error::TableNotFound,
        popup::Error::CellNotFound => Error::CellNotFound,
        popup::Error::TableLocked { path } => Error::TableLocked {
            path: from_native_path(path),
        },
        popup::Error::UnsupportedDependency { path } => Error::UnsupportedDependency {
            path: from_native_path(path),
        },
        popup::Error::UnsupportedSource => Error::UnsupportedSource,
        popup::Error::InvalidSource { path } => Error::InvalidSource {
            path: from_native_path(path),
        },
        popup::Error::LimitExceeded {
            kind,
            observed,
            maximum,
            path,
        } => Error::LimitExceeded {
            kind: map_limit_kind(kind),
            observed,
            maximum,
            path: from_native_path(path),
        },
        popup::Error::Allocation { amount, path } => Error::Allocation {
            amount,
            path: from_native_path(path),
        },
        popup::Error::Verification => Error::Verification,
        popup::Error::PatchConflict => Error::PatchConflict,
    }
}

const fn map_limit_kind(kind: popup::LimitKind) -> LimitKind {
    match kind {
        popup::LimitKind::InputBytes => LimitKind::InputBytes,
        popup::LimitKind::OutputBytes => LimitKind::OutputBytes,
        popup::LimitKind::Entries => LimitKind::Entries,
        popup::LimitKind::EntryBytes => LimitKind::EntryBytes,
        popup::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
        popup::LimitKind::PackageBytes => LimitKind::PackageBytes,
        popup::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
        popup::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        popup::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        popup::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        popup::LimitKind::PayloadItems => LimitKind::PayloadItems,
        popup::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
        popup::LimitKind::WireBytes => LimitKind::WireBytes,
        popup::LimitKind::WireOutputBytes => LimitKind::WireOutputBytes,
        popup::LimitKind::WireReferenceBytes => LimitKind::WireReferenceBytes,
        popup::LimitKind::WireTextBytes => LimitKind::WireTextBytes,
        popup::LimitKind::WireFields => LimitKind::WireFields,
        popup::LimitKind::WireNesting => LimitKind::WireNesting,
        popup::LimitKind::WireWork => LimitKind::WireWork,
        popup::LimitKind::ScratchBytes => LimitKind::ScratchBytes,
        popup::LimitKind::RetainedBytes => LimitKind::RetainedBytes,
        popup::LimitKind::Allocations => LimitKind::Allocations,
        popup::LimitKind::CompressedBytes => LimitKind::CompressedBytes,
        popup::LimitKind::TransactionWork => LimitKind::TransactionWork,
    }
}

fn verify_percentage_package_locality(
    source: &Package,
    candidate: &Package,
) -> Result<usize, Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    let source_index = physical_entry_index(source_catalog)?;
    let candidate_index = physical_entry_index(candidate_catalog)?;
    let changed = changed_member_names(
        source_catalog,
        candidate_catalog,
        &source_index,
        &candidate_index,
    )?;
    verify_percentage_package_locality_with_indexes(
        source,
        candidate,
        &changed,
        source_catalog,
        candidate_catalog,
        &source_index,
        &candidate_index,
    )
}

fn verify_percentage_package_locality_with_members(
    source: &Package,
    candidate: &Package,
    changed_members: &[&str],
) -> Result<usize, Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    let source_index = physical_entry_index(source_catalog)?;
    let candidate_index = physical_entry_index(candidate_catalog)?;
    verify_percentage_package_locality_with_indexes(
        source,
        candidate,
        changed_members,
        source_catalog,
        candidate_catalog,
        &source_index,
        &candidate_index,
    )
}

fn verify_percentage_package_locality_with_indexes(
    source: &Package,
    candidate: &Package,
    changed_members: &[&str],
    source_catalog: &litchi_iwa_archive::SourceCatalog,
    candidate_catalog: &litchi_iwa_archive::SourceCatalog,
    source_index: &PhysicalEntryIndex<'_>,
    candidate_index: &PhysicalEntryIndex<'_>,
) -> Result<usize, Error> {
    if changed_members
        .iter()
        .any(|name| *name == super::metadata::ENTRY_NAME || name.starts_with("preview"))
    {
        return Err(Error::Verification);
    }
    popup::verify_package_locality_for_members(source, candidate, changed_members)
        .map_err(|_| Error::Verification)?;
    for source_entry in source_catalog.package().iter().filter(|entry| {
        entry.name() == super::metadata::ENTRY_NAME || entry.name().starts_with("preview")
    }) {
        let candidate_entry = candidate_index
            .get(source_entry.name())
            .ok_or(Error::Verification)?;
        if !super::table_headers::rewrite::package_member_preserved(source_entry, candidate_entry) {
            return Err(Error::Verification);
        }
    }
    for candidate_entry in candidate_catalog.package().iter().filter(|entry| {
        entry.name() == super::metadata::ENTRY_NAME || entry.name().starts_with("preview")
    }) {
        let source_entry = source_index
            .get(candidate_entry.name())
            .ok_or(Error::Verification)?;
        if !super::table_headers::rewrite::package_member_preserved(source_entry, candidate_entry) {
            return Err(Error::Verification);
        }
    }
    Ok(changed_member_count(source_catalog, candidate_index))
}

fn changed_member_names<'a>(
    source: &'a litchi_iwa_archive::SourceCatalog,
    candidate: &'a litchi_iwa_archive::SourceCatalog,
    source_index: &PhysicalEntryIndex<'a>,
    candidate_index: &PhysicalEntryIndex<'a>,
) -> Result<Vec<&'a str>, Error> {
    // At most one borrowed name is retained for each non-preview member in
    // either catalog. Reserve that finite upper bound before the comparison
    // walk so hostile package entry counts cannot trigger infallible Vec
    // growth. The returned names borrow the immutable source/candidate
    // catalogs and need no per-name String allocation.
    let capacity = source
        .package()
        .iter()
        .count()
        .checked_add(candidate.package().iter().count())
        .ok_or(Error::Verification)?;
    let mut changed = Vec::new();
    changed
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Verification)?;
    for source_entry in source.package().iter() {
        if source_entry.name().starts_with("preview") {
            continue;
        }
        let differs = candidate_index
            .get(source_entry.name())
            .is_none_or(|entry| entry.data() != source_entry.data());
        if differs {
            changed.push(source_entry.name());
        }
    }
    for candidate_entry in candidate.package().iter() {
        if candidate_entry.name().starts_with("preview") {
            continue;
        }
        if source_index.get(candidate_entry.name()).is_none() {
            changed.push(candidate_entry.name());
        }
    }
    changed.sort_unstable();
    changed.dedup();
    Ok(changed)
}

fn physical_entry_index(
    catalog: &litchi_iwa_archive::SourceCatalog,
) -> Result<PhysicalEntryIndex<'_>, Error> {
    PhysicalEntryIndex::new(catalog.package()).map_err(|error| match error {
        PhysicalEntryIndexError::Allocation { .. } | PhysicalEntryIndexError::Duplicate => {
            Error::Verification
        },
    })
}

fn changed_member_count(
    source: &litchi_iwa_archive::SourceCatalog,
    candidate_index: &PhysicalEntryIndex<'_>,
) -> usize {
    source
        .package()
        .iter()
        .filter(|entry| !entry.name().starts_with("preview"))
        .filter(|entry| {
            candidate_index
                .get(entry.name())
                .is_some_and(|candidate| candidate.data() != entry.data())
        })
        .count()
}
