//! Selector-first Pop-Up Menu cell-control transactions.
//!
//! This is the package boundary for the Numbers Pop-Up Menu owner. The
//! native `TST.PopUpMenuModel` graph is handled through strict, generated-free
//! codecs plus a refcount/ownership census and metadata transition. Semantic
//! callers therefore do not need native identifiers or row/column integer
//! pairs.

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_table::cell_format as table;
use litchi_iwa_table::cell_format::pop_up_menu::{
    changed_member_count, read_popup, read_popup_with_budget,
};
pub(super) use litchi_iwa_table::cell_format::pop_up_menu::{
    read_cell_control, rewrite_component_save_tokens,
};
pub(super) use litchi_iwa_table::cell_format::{Target as CellTarget, TransactionBudget};
use thiserror::Error as ThisError;

use super::Package;
use crate::{SheetSelector, TableSelector, cell::data_format::PopUpMenu, table::At};

/// A content-free location associated with a Pop-Up Menu operation.
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
        position: At,
    },
}

/// A finite resource governed by a Pop-Up Menu transaction.
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

/// A content-redacted Pop-Up Menu transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// Rooted ownership or wire framing is invalid.
    #[error("the Numbers Pop-Up Menu source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers Pop-Up Menu {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
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
    #[error("could not allocate {amount} units for the Numbers Pop-Up Menu transaction")]
    Allocation { amount: usize, path: Path },
    /// Candidate reopening or semantic locality verification failed.
    #[error("the edited Numbers Pop-Up Menu failed semantic verification")]
    Verification,
    /// A patch was created for another exact package artifact.
    #[error("the Pop-Up Menu patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
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
    /// A native dependency is not safe to rewrite in this owner.
    #[error("the selected Numbers Pop-Up Menu has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: Path },
    /// The exact native Pop-Up Menu owner is not available for this source.
    #[error("this Numbers source does not support exact Pop-Up Menu editing")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound { .. } | Self::TableNotFound { .. } | Self::CellNotFound { .. } => {
                RefusalKind::NotFound
            },
            Self::TableLocked { .. } => RefusalKind::Protected,
            Self::UnsupportedDependency { .. } | Self::UnsupportedSource { .. } => {
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

// The shared table engine (`litchi_iwa_table::cell_format`, record 0379)
// reports the same failures with a leaf-neutral location: its `scope` is the
// rooted sheet position.
impl From<table::Path> for Path {
    fn from(path: table::Path) -> Self {
        match path {
            table::Path::Package => Self::Package,
            table::Path::Cell {
                scope,
                table,
                position,
            } => Self::Cell {
                sheet: scope,
                table,
                position,
            },
        }
    }
}

impl From<Path> for table::Path {
    fn from(path: Path) -> Self {
        match path {
            Path::Package => Self::Package,
            Path::Cell {
                sheet,
                table,
                position,
            } => Self::Cell {
                scope: sheet,
                table,
                position,
            },
        }
    }
}

impl From<table::LimitKind> for LimitKind {
    fn from(kind: table::LimitKind) -> Self {
        match kind {
            table::LimitKind::InputBytes => Self::InputBytes,
            table::LimitKind::OutputBytes => Self::OutputBytes,
            table::LimitKind::Entries => Self::Entries,
            table::LimitKind::EntryBytes => Self::EntryBytes,
            table::LimitKind::TotalEntryBytes => Self::TotalEntryBytes,
            table::LimitKind::PackageBytes => Self::PackageBytes,
            table::LimitKind::PayloadBytes => Self::PayloadBytes,
            table::LimitKind::TotalPayloadBytes => Self::TotalPayloadBytes,
            table::LimitKind::PayloadObjects => Self::PayloadObjects,
            table::LimitKind::PayloadMessages => Self::PayloadMessages,
            table::LimitKind::PayloadItems => Self::PayloadItems,
            table::LimitKind::PayloadReferences => Self::PayloadReferences,
            table::LimitKind::WireBytes => Self::WireBytes,
            table::LimitKind::WireOutputBytes => Self::WireOutputBytes,
            table::LimitKind::WireReferenceBytes => Self::WireReferenceBytes,
            table::LimitKind::WireTextBytes => Self::WireTextBytes,
            table::LimitKind::WireFields => Self::WireFields,
            table::LimitKind::WireNesting => Self::WireNesting,
            table::LimitKind::WireWork => Self::WireWork,
            table::LimitKind::ScratchBytes => Self::ScratchBytes,
            table::LimitKind::RetainedBytes => Self::RetainedBytes,
            table::LimitKind::Allocations => Self::Allocations,
            table::LimitKind::CompressedBytes => Self::CompressedBytes,
            table::LimitKind::TransactionWork => Self::TransactionWork,
        }
    }
}

impl From<table::Error> for Error {
    fn from(error: table::Error) -> Self {
        match error {
            table::Error::Refused(refusal) => Self::Refused(match refusal {
                table::Refusal::CellNotFound => Refusal::CellNotFound,
                table::Refusal::TableLocked { path } => Refusal::TableLocked { path: path.into() },
                table::Refusal::UnsupportedDependency { path } => {
                    Refusal::UnsupportedDependency { path: path.into() }
                },
                table::Refusal::UnsupportedSource => Refusal::UnsupportedSource,
            }),
            table::Error::InvalidSource { path } => Self::InvalidSource { path: path.into() },
            table::Error::Limit {
                kind,
                observed,
                maximum,
                path,
            } => Self::Limit {
                kind: kind.into(),
                observed,
                maximum,
                path: path.into(),
            },
            table::Error::Allocation { amount, path } => Self::Allocation {
                amount,
                path: path.into(),
            },
            table::Error::Verification => Self::Verification,
            table::Error::PatchConflict => Self::PatchConflict,
        }
    }
}

/// Preserve resource and allocation failures from the bounded Snappy
/// encoder at the Numbers transaction boundary.
pub(super) fn map_core_error(error: litchi_iwa_core::Error, path: table::Path) -> Error {
    table::map_core_error(error, path).into()
}

/// A selector-first Pop-Up Menu edit staged against one package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected semantic cell.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the value staged for publication, if any.
    #[must_use]
    pub fn pop_up_menu(&self) -> Option<&PopUpMenu> {
        self.after.as_ref()
    }

    /// Return the value staged for publication, if any.
    #[must_use]
    pub fn format(&self) -> Option<&PopUpMenu> {
        self.pop_up_menu()
    }

    /// Return the value observed when this edit was created, if one existed.
    #[must_use]
    pub fn before(&self) -> Option<&PopUpMenu> {
        self.before.as_ref()
    }

    /// Stage a Pop-Up Menu value without touching package bytes.
    #[must_use]
    pub fn set(mut self, value: PopUpMenu) -> Self {
        self.after = Some(value);
        self
    }

    /// Stage the automatic/no-menu state for the selected cell.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    /// Alias for [`Self::clear`] used by cell data-format callers.
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    /// Validate and publish the staged edit.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.before == self.after {
            return no_op_commit(self.source, self.path, self.before, self.after);
        }
        rewrite_transaction(self.source, self.path, self.before, self.after)
    }
}

/// A reversible exact-source Pop-Up Menu patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
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

    /// Return the exact source value.
    #[must_use]
    pub const fn before(&self) -> Option<&PopUpMenu> {
        self.before.as_ref()
    }

    /// Return the exact target value.
    #[must_use]
    pub const fn after(&self) -> Option<&PopUpMenu> {
        self.after.as_ref()
    }

    /// Return the target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return whether this patch represents an unchanged semantic value.
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

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of previews removed by the changed operation.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether a complete candidate package was reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One fully validated Pop-Up Menu publication.
#[must_use = "a Pop-Up Menu commit contains the validated package snapshot"]
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

    /// Consume the publication and return its package snapshot.
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

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the selected cell's Pop-Up Menu, if one is present.
    ///
    /// The strict native owner admits only rooted same-component control
    /// graphs whose table lists, model ownership, and metadata agree exactly.
    pub fn table_cell_pop_up_menu_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Option<PopUpMenu>, Error> {
        let target = resolve_cell(self, sheet, table, position)?;
        Ok(read_popup(self, target)?)
    }

    /// Start a selector-first Pop-Up Menu edit.
    pub fn edit_table_cell_pop_up_menu_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: At,
    ) -> Result<Edit<'_>, Error> {
        let target = resolve_cell(self, sheet, table, position)?;
        let before = read_popup(self, target)?;
        Ok(Edit {
            source: self,
            path: Path::Cell {
                sheet: target.scope,
                table: target.table,
                position,
            },
            before: before.clone(),
            after: before,
        })
    }

    /// Apply a reversible exact-source Pop-Up Menu patch.
    pub fn apply_table_cell_pop_up_menu_format(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = super::table_headers::rewrite::physical_source(self)
            .map_err(|_| Error::Refused(Refusal::UnsupportedSource))?;
        let source = catalog.source_owner();
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
        let Path::Cell {
            sheet,
            table,
            position,
        } = patch.path
        else {
            return Err(Error::PatchConflict);
        };
        let current = resolve_cell(
            self,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        if read_popup(self, current)? != patch.before {
            return Err(Error::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        let mut budget = TransactionBudget::new(self);
        let source_catalog = super::table_headers::rewrite::physical_source(self)
            .map_err(|_| Error::Refused(Refusal::UnsupportedSource))?;
        budget.charge_package_source(source_catalog, patch.path.into())?;
        budget.charge_output(target_owner.as_ref().len(), patch.path.into())?;
        budget.charge_transaction_work(
            target_owner.as_ref().len().saturating_mul(2),
            patch.path.into(),
        )?;
        budget.charge_allocations(2, patch.path.into())?;
        budget.charge_candidate_input_bytes(target_owner.as_ref().len(), patch.path.into())?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
            .map_err(|_| Error::Verification)?;
        budget.charge_candidate_reopen(candidate_catalog, patch.path.into())?;
        let candidate_target = resolve_cell(
            &candidate,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        let after = read_popup_with_budget(&candidate, candidate_target, &mut budget, true)?;
        if after != patch.after {
            return Err(Error::Verification);
        }
        let touched_components = {
            let target_catalog = super::table_headers::rewrite::physical_source(&candidate)
                .map_err(|_| Error::Verification)?;
            changed_member_count(source_catalog, target_catalog)
        };
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
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
) -> Result<Commit, Error> {
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::Refused(Refusal::UnsupportedSource))?;
    let owner = catalog.source_owner();
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

/// First lifecycle checkpoint: rewrite a same-component document graph into a
/// private candidate, then reopen it before publication.  The metadata sidecar
/// is updated in the same ZIP reassembly; later hardening can replace the
/// generated-free list surgery below with the prepared storage transition
/// without changing the public transaction or patch contract.
fn rewrite_transaction(
    source: &Package,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
) -> Result<Commit, Error> {
    let Path::Cell {
        sheet,
        table,
        position,
    } = path
    else {
        return Err(Error::InvalidSource { path });
    };
    let target = resolve_cell(
        source,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )?;
    if target.locked {
        return Err(Error::Refused(Refusal::TableLocked { path }));
    }
    let source_owner = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::Refused(Refusal::UnsupportedSource))?
        .source_owner();
    let published = table::pop_up_publish::publish(
        source,
        target,
        before.as_ref(),
        after.as_ref(),
        true,
        table::ListKeys::Reuse,
        true,
        |bytes| {
            Package::from_owned_bytes_with_options(bytes, source.state.options)
                .map_err(|_| table::Error::Verification)
        },
        |candidate| {
            resolve_cell(
                candidate,
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )
            .map_err(|_| table::Error::Verification)
        },
    )?;
    let candidate = published.package;
    let target_owner = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?
        .source_owner();
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
            touched_components: published.touched_components,
            deleted_previews: published.deleted_previews,
            full_reparse_performed: true,
        },
    })
}

/// Verify package-level locality for a native transaction that rewrites one
/// or more physical members.  All non-authorized members must remain exact,
/// including opaque/versioned entries and metadata; preview deletions are
/// intentionally handled by the caller because their policy is operation
/// specific.
pub(super) fn verify_package_locality_for_members(
    source: &Package,
    candidate: &Package,
    native_members: &[&str],
) -> Result<(), Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    Ok(table::pop_up_publish::verify_locality(
        source_catalog,
        candidate_catalog,
        native_members,
        table::Path::Package,
    )?)
}

pub(super) fn resolve_cell<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    position: At,
) -> Result<CellTarget, Error> {
    let sheet_selector = sheet.into();
    let selected_sheet = source
        .state
        .document
        .sheet(sheet_selector)
        .map_err(|_| Error::InvalidSource {
            path: Path::Package,
        })?
        .ok_or(Error::Refused(Refusal::SheetNotFound))?;
    let table_selector = table.into();
    let table_position = match table_selector {
        TableSelector::Position(index) => {
            let index = index.get();
            selected_sheet.tables().nth(index).map(|_| index)
        },
        TableSelector::Name(name) => {
            let mut matches = selected_sheet
                .tables()
                .enumerate()
                .filter(|(_, table)| table.name() == name);
            let first = matches.next().map(|(index, _)| index);
            if matches.next().is_some() {
                return Err(Error::InvalidSource {
                    path: Path::Cell {
                        sheet: selected_sheet.index(),
                        table: 0,
                        position,
                    },
                });
            }
            first
        },
    }
    .ok_or(Error::Refused(Refusal::TableNotFound))?;
    let native = super::table_headers::resolve::resolve_target(
        source,
        selected_sheet.index(),
        table_position,
    )
    .map_err(|_| Error::InvalidSource {
        path: Path::Cell {
            sheet: selected_sheet.index(),
            table: table_position,
            position,
        },
    })?;
    if position.row() >= native.rows || position.column() >= native.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    Ok(CellTarget {
        scope: selected_sheet.index(),
        table: table_position,
        position,
        model_identifier: native.model_identifier,
        component_index: native.component_index,
        object_index: native.object_index,
        message_index: native.message_index,
        message_type: native.message_type,
        locked: native.locked == crate::table::lock::State::Locked,
    })
}
