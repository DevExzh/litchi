//! Selector-first transactions for an existing cell's document-scoped
//! Custom display format.
//!
//! The public value is archive-free.  Native format-list keys, UUIDs, IWA
//! member names, and wire payloads remain private to the native owner.  A
//! changed transaction is candidate-reopened and source-bound in the same
//! way as the other focused cell-format owners.
//!
//! Transaction [`Debug`](fmt::Debug) output delegates to the semantic
//! redacted views and never includes the staged user content or native
//! artifacts.

use std::fmt;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_core::SnappyStream;
use thiserror::Error as ThisError;

use super::{
    Package, table_cell_control_native as control_native,
    table_cell_display_format_native as native, table_cell_pop_up_menu as popup,
};
use crate::{CellPosition, SheetSelector, TableSelector, cell::data_format::custom::Custom};

/// A content-free location associated with a Custom-format operation.
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

/// A finite resource governed by a Custom-format transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalEntryBytes,
    PackageBytes,
    PayloadBytes,
    TotalPayloadBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    PayloadReferences,
    WireBytes,
    WireOutputBytes,
    WireReferenceBytes,
    WireTextBytes,
    WireFields,
    WireNesting,
    WireWork,
    ScratchBytes,
    RetainedBytes,
    Allocations,
    CompressedBytes,
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted Custom-format transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    #[error("the selected Numbers table cell was not found")]
    CellNotFound,
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    #[error("the selected Numbers cell does not contain an explicit Custom format at {path:?}")]
    WrongFormatFamily { path: Path },
    #[error("the selected Numbers Custom format has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: Path },
    #[error("this Numbers source does not support exact Custom-format editing")]
    UnsupportedSource,
    #[error("the Numbers Custom-format source is invalid at {path:?}")]
    InvalidSource { path: Path },
    #[error("Numbers Custom format {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
        path: Path,
    },
    #[error("could not allocate {amount} units for the Numbers Custom-format transaction")]
    Allocation { amount: usize, path: Path },
    #[error("the edited Numbers Custom format failed semantic verification")]
    Verification,
    #[error("the Custom-format patch does not match the exact source package")]
    PatchConflict,
}

impl From<popup::Error> for Error {
    fn from(error: popup::Error) -> Self {
        map_popup_error(error)
    }
}

/// A selector-first Custom-format edit staged against one exact source.
pub struct Edit<'a> {
    source: &'a Package,
    path: Path,
    before: Option<Custom>,
    after: Option<Custom>,
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
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the currently staged Custom value.
    #[must_use]
    pub const fn custom_format(&self) -> Option<&Custom> {
        self.after.as_ref()
    }

    /// Alias for [`Self::custom_format`].
    #[must_use]
    pub const fn format(&self) -> Option<&Custom> {
        self.custom_format()
    }

    #[must_use]
    pub const fn before(&self) -> Option<&Custom> {
        self.before.as_ref()
    }

    #[must_use]
    pub const fn after(&self) -> Option<&Custom> {
        self.after.as_ref()
    }

    #[must_use]
    pub fn set(mut self, value: Custom) -> Self {
        self.after = Some(value);
        self
    }

    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    pub fn commit(self) -> Result<Commit, Error> {
        if self.before == self.after {
            return no_op_commit(self.source, self.path, self.before, self.after);
        }
        rewrite(self.source, self.path, self.before, self.after)
    }
}

/// A reversible exact-source Custom-format patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Option<Custom>,
    after: Option<Custom>,
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
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    #[must_use]
    pub const fn before(&self) -> Option<&Custom> {
        self.before.as_ref()
    }

    #[must_use]
    pub const fn after(&self) -> Option<&Custom> {
        self.after.as_ref()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
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

    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One fully validated Custom-format publication.
#[must_use = "a Custom-format commit contains the validated package snapshot"]
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
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one existing cell's explicit document-scoped Custom format.
    pub fn table_cell_custom_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<Custom>, Error> {
        let target = popup::resolve_cell(self, sheet, table, position).map_err(map_popup_error)?;
        let path = Path::Cell {
            sheet: target.sheet_position,
            table: target.table_position,
            position,
        };
        native::read_custom_format(self, target, native_path(path))
            .map_err(|error| map_custom_read_error(error, path))
    }

    /// Start a selector-first Custom-format edit for one existing cell.
    pub fn edit_table_cell_custom_format<'sheet, 'table>(
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
        let before = native::read_custom_format(self, target, native_path(path))
            .map_err(|error| map_custom_read_error(error, path))?;
        Ok(Edit {
            source: self,
            path,
            before: before.clone(),
            after: before,
        })
    }

    /// Apply a previously-created exact Custom-format patch.
    pub fn apply_table_cell_custom_format(&self, patch: &Patch) -> Result<Commit, Error> {
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
        let current = self.table_cell_custom_format(
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
        let application_work = target_owner
            .as_ref()
            .len()
            .checked_mul(2)
            .and_then(|amount| {
                catalog
                    .package()
                    .iter()
                    .count()
                    .checked_mul(1024)
                    .and_then(|catalog_work| amount.checked_add(catalog_work))
            })
            .ok_or(Error::Verification)?;
        budget.charge_transaction_work(application_work, native_path(patch.path))?;
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
        let after = native::read_custom_format_with_budget(
            &candidate,
            candidate_target,
            native_path(patch.path),
            &mut budget,
        )
        .map_err(|error| map_custom_read_error(error, patch.path))?;
        if after != patch.after {
            return Err(Error::Verification);
        }
        let changed_members = changed_member_names(self, &candidate)?;
        verify_custom_package_locality(self, &candidate, &changed_members)?;
        let touched_components = changed_members.len();
        drop(changed_members);
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
    before: Option<Custom>,
    after: Option<Custom>,
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
    before: Option<Custom>,
    after: Option<Custom>,
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
        native::read_custom_format_with_budget(source, target, native_path(path), &mut budget)
            .map_err(|error| map_custom_read_error(error, path))?;
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
    let native_output = native::rewrite_custom_format(
        source,
        target,
        before.as_ref(),
        after.as_ref(),
        native_path(path),
        &mut budget,
    )
    .map_err(map_popup_error)?;
    let member_count = native_output.members.len();
    if !(1..=3).contains(&member_count) {
        return Err(Error::Verification);
    }

    let mut maximum_compressed_bytes = 0usize;
    let mut compression_allocations = 3usize;
    for member in &native_output.members {
        if member.archive_bytes.len() > native_bound.archive_bytes {
            return Err(Error::LimitExceeded {
                kind: LimitKind::PayloadBytes,
                observed: member.archive_bytes.len() as u64,
                maximum: native_bound.archive_bytes as u64,
                path,
            });
        }
        let maximum = SnappyStream::maximum_compressed_len(member.archive_bytes.len())
            .map_err(|_| Error::Verification)?;
        if maximum > native_bound.compressed_bytes {
            return Err(Error::LimitExceeded {
                kind: LimitKind::CompressedBytes,
                observed: maximum as u64,
                maximum: native_bound.compressed_bytes as u64,
                path,
            });
        }
        maximum_compressed_bytes = maximum_compressed_bytes
            .checked_add(maximum)
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
    budget.charge_compressed_bytes(maximum_compressed_bytes, native_path(path))?;
    budget.charge_allocations(compression_allocations, native_path(path))?;
    budget.charge_scratch_bytes(maximum_compressed_bytes, native_path(path))?;
    budget.charge_retained_bytes(maximum_compressed_bytes, native_path(path))?;
    let compression_work = source_catalog
        .source_bytes()
        .len()
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(maximum_compressed_bytes))
        .ok_or(Error::Verification)?;
    budget.charge_transaction_work(compression_work, native_path(path))?;

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
        if compressed.len()
            > SnappyStream::maximum_compressed_len(member.archive_bytes.len())
                .map_err(|_| Error::Verification)?
        {
            return Err(Error::Verification);
        }
        compressed_members.push(compressed);
    }
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
    budget.charge_allocations(2, native_path(path))?;
    let prepared = source_catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], source_catalog.limits())
        .map_err(|_| Error::Verification)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_reassembly(requirements, native_path(path))?;
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
    budget.charge_transaction_work(publication_work, native_path(path))?;
    budget.charge_candidate_input_bytes(requirements.output_bytes(), native_path(path))?;
    let bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::Verification)?;
    let candidate = Package::from_owned_bytes_with_options(bytes, source.state.options)
        .map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?;
    budget.charge_candidate_reopen(candidate_catalog, native_path(path))?;
    let candidate_target = popup::resolve_cell(
        &candidate,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )
    .map_err(map_popup_error)?;
    let candidate_value = native::read_custom_format_with_budget(
        &candidate,
        candidate_target,
        native_path(path),
        &mut budget,
    )
    .map_err(|error| map_custom_read_error(error, path))?;
    if candidate_value != after {
        return Err(Error::Verification);
    }
    let changed_names = native_output
        .members
        .iter()
        .map(|member| member.member_name.as_str())
        .collect::<Vec<_>>();
    verify_custom_package_locality(source, &candidate, &changed_names)?;
    let touched_components = changed_member_count(source, &candidate)?;
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

fn map_custom_read_error(error: native::CustomFormatReadError, path: Path) -> Error {
    match error {
        native::CustomFormatReadError::WrongFormatFamily => Error::WrongFormatFamily { path },
        native::CustomFormatReadError::Native(error) => map_popup_error(error),
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

fn changed_member_names<'a>(
    source: &'a Package,
    candidate: &'a Package,
) -> Result<Vec<&'a str>, Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(
            source_catalog
                .package()
                .iter()
                .count()
                .checked_add(candidate_catalog.package().iter().count())
                .ok_or(Error::Verification)?,
        )
        .map_err(|_| Error::Verification)?;
    for entry in source_catalog.package().iter() {
        if entry.name().starts_with("preview") {
            continue;
        }
        if candidate_catalog
            .package()
            .iter()
            .find(|other| other.name() == entry.name())
            .is_none_or(|other| other.data() != entry.data())
        {
            names.push(entry.name());
        }
    }
    for entry in candidate_catalog.package().iter() {
        if !entry.name().starts_with("preview")
            && source_catalog
                .package()
                .iter()
                .all(|other| other.name() != entry.name())
        {
            names.push(entry.name());
        }
    }
    names.sort_unstable();
    names.dedup();
    Ok(names)
}

fn changed_member_count(source: &Package, candidate: &Package) -> Result<usize, Error> {
    Ok(changed_member_names(source, candidate)?.len())
}

fn verify_custom_package_locality(
    source: &Package,
    candidate: &Package,
    changed_members: &[&str],
) -> Result<(), Error> {
    if changed_members
        .iter()
        .any(|name| *name == super::metadata::ENTRY_NAME || name.starts_with("preview"))
    {
        return Err(Error::Verification);
    }
    popup::verify_package_locality_for_members(source, candidate, changed_members)
        .map_err(|_| Error::Verification)?;
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    let source_metadata = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == super::metadata::ENTRY_NAME)
        .ok_or(Error::Verification)?;
    let candidate_metadata = candidate_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == super::metadata::ENTRY_NAME)
        .ok_or(Error::Verification)?;
    if source_metadata.data() != candidate_metadata.data() {
        return Err(Error::Verification);
    }
    for source_preview in source_catalog
        .package()
        .iter()
        .filter(|entry| entry.name().starts_with("preview"))
    {
        let candidate_preview = candidate_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == source_preview.name())
            .ok_or(Error::Verification)?;
        if source_preview.data() != candidate_preview.data() {
            return Err(Error::Verification);
        }
    }
    Ok(())
}
