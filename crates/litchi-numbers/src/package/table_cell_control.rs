//! Selector-first ownership for all interactive Numbers cell controls.
//!
//! This module is the single public transaction boundary for checkbox,
//! star-rating, slider, stepper, and Pop-Up Menu cells.  The Pop-Up Menu
//! implementation remains the audited native graph engine; this facade wraps
//! it for compatibility while the scalar-control native path is kept in the
//! sibling private module.  No native identifiers or generated messages are
//! present in the public transaction values.
//!
//! Shared native control entries are updated with copy-on-write semantics;
//! the exact patch/inverse transaction remains the publication boundary.

use std::fmt;

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_core::SnappyStream;

use super::table_cell_control_native as native;
use super::{Package, table_cell_pop_up_menu as popup};
use crate::{CellPosition, SheetSelector, TableSelector, cell::data_format::control::CellControl};

pub use popup::{Error, LimitKind, Path};

/// A selector-first edit staged against one exact package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    path: Path,
    before: Option<CellControl>,
    after: Option<CellControl>,
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
    /// Return the selected semantic cell path.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the staged control, if any.
    #[must_use]
    pub fn control(&self) -> Option<&CellControl> {
        self.after.as_ref()
    }

    /// Alias useful to generic data-format callers.
    #[must_use]
    pub fn format(&self) -> Option<&CellControl> {
        self.control()
    }

    /// Return the value observed when the edit was opened.
    #[must_use]
    pub fn before(&self) -> Option<&CellControl> {
        self.before.as_ref()
    }

    /// Stage a control value without changing package bytes.
    #[must_use]
    pub fn set(mut self, value: CellControl) -> Self {
        self.after = Some(value);
        self
    }

    /// Stage the automatic/no-control state.
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
        let Path::Cell {
            sheet,
            table,
            position,
        } = self.path
        else {
            return Err(Error::InvalidSource { path: self.path });
        };
        let target = popup::resolve_cell(
            self.source,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        if target.locked {
            return Err(Error::TableLocked { path: self.path });
        }
        if self.before == self.after {
            return no_op_commit(self.source, self.path, self.before, self.after);
        }
        rewrite(self.source, self.path, self.before, self.after)
    }
}

/// A reversible exact-source control patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Option<CellControl>,
    after: Option<CellControl>,
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
    pub const fn before(&self) -> Option<&CellControl> {
        self.before.as_ref()
    }

    /// Return the exact target semantic value.
    #[must_use]
    pub const fn after(&self) -> Option<&CellControl> {
        self.after.as_ref()
    }

    /// Return the target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Whether this patch is a semantic and byte-level no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
}

/// Content-free diagnostics for one publication.
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

    /// Number of changed native/metadata components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of deleted preview members.
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
#[must_use = "a cell-control commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
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
    /// Read the selected rooted cell's complete interactive control.
    pub fn table_cell_control_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<CellControl>, Error> {
        let target = popup::resolve_cell(self, sheet, table, position)?;
        popup::read_cell_control(self, target)
    }

    /// Start a selector-first complete-control edit.
    pub fn edit_table_cell_control_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Edit<'_>, Error> {
        let target = popup::resolve_cell(self, sheet, table, position)?;
        let before = popup::read_cell_control(self, target)?;
        Ok(Edit {
            source: self,
            path: Path::Cell {
                sheet: target.sheet_position,
                table: target.table_position,
                position,
            },
            before: before.clone(),
            after: before,
        })
    }

    /// Apply a previously-created exact control patch.
    pub fn apply_table_cell_control_format(&self, patch: &Patch) -> Result<Commit, Error> {
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
        let current = self.table_cell_control_format(
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
        let mut budget = popup::TransactionBudget::new(self);
        budget.charge_package_source(catalog, patch.path)?;
        budget.charge_output(target_owner.as_ref().len(), patch.path)?;
        budget.charge_allocations(2, patch.path)?;
        budget.charge_transaction_work(
            target_owner
                .as_ref()
                .len()
                .saturating_mul(2)
                .saturating_add(catalog.package().iter().count().saturating_mul(1024)),
            patch.path,
        )?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let after = candidate.table_cell_control_format(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        if after != patch.after {
            return Err(Error::Verification);
        }
        let target_catalog = super::table_headers::rewrite::physical_source(&candidate)
            .map_err(|_| Error::Verification)?;
        budget.charge_payload_items(target_catalog.package().iter().count(), patch.path)?;
        budget.charge_transaction_work(
            target_catalog.source_bytes().len().saturating_mul(2),
            patch.path,
        )?;
        let touched_components = catalog
            .package()
            .iter()
            .filter(|source_entry| {
                target_catalog
                    .package()
                    .iter()
                    .find(|target_entry| target_entry.name() == source_entry.name())
                    .is_none_or(|target_entry| target_entry.data() != source_entry.data())
            })
            .count();
        let deleted_previews = catalog
            .package()
            .iter()
            .filter(|entry| entry.name().starts_with("preview"))
            .filter(|source_entry| {
                target_catalog
                    .package()
                    .iter()
                    .all(|target_entry| target_entry.name() != source_entry.name())
            })
            .count();
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components,
                deleted_previews,
                full_reparse_performed: true,
            },
        })
    }
}

fn no_op_commit(
    source: &Package,
    path: Path,
    before: Option<CellControl>,
    after: Option<CellControl>,
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
    before: Option<CellControl>,
    after: Option<CellControl>,
) -> Result<Commit, Error> {
    let Path::Cell {
        sheet,
        table,
        position,
    } = path
    else {
        return Err(Error::InvalidSource { path });
    };
    let source_owner = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::UnsupportedSource)?
        .__source_owner();
    let (target, popup_before, popup_after) = match (&before, &after) {
        (Some(CellControl::PopUpMenu(old)), Some(CellControl::PopUpMenu(new))) => {
            (true, Some(old.clone()), Some(new.clone()))
        },
        (Some(CellControl::PopUpMenu(old)), None) => (true, Some(old.clone()), None),
        (None, Some(CellControl::PopUpMenu(new))) => (true, None, Some(new.clone())),
        _ => (false, None, None),
    };
    if target {
        let edit = source.edit_table_cell_pop_up_menu_format(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        let commit = match &popup_after {
            Some(value) => edit.set(value.clone()).commit()?,
            None => edit.clear().commit()?,
        };
        if commit.patch().before() != popup_before.as_ref()
            || commit.patch().after() != popup_after.as_ref()
        {
            return Err(Error::Verification);
        }
        let diagnostics = *commit.diagnostics();
        let package = commit.into_package();
        let target_owner = super::table_headers::rewrite::physical_source(&package)
            .map_err(|_| Error::Verification)?
            .__source_owner();
        return Ok(Commit {
            package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                path,
                before,
                after,
            },
            diagnostics: Diagnostics {
                changed: true,
                touched_components: diagnostics.touched_components(),
                deleted_previews: diagnostics.deleted_previews(),
                full_reparse_performed: diagnostics.full_reparse_performed(),
            },
        });
    }
    let target = popup::resolve_cell(
        source,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )?;
    popup::reject_cross_component_write(source, target, path)?;
    let mut budget = popup::TransactionBudget::for_cell_control(source);
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::UnsupportedSource)?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    // Charge the complete source ZIP and native archive before the scalar
    // owner clones any archive or stages a list candidate.  The native helper
    // returns a conservative private-candidate bound; its exact serialized
    // size is checked again before compression.
    budget.charge_package_source(catalog, path)?;
    let native_bound =
        native::preflight_copy_on_write_scalar_control(source, target, path, &mut budget)?;
    let native_output = native::rewrite_scalar_control(
        source,
        target,
        before.as_ref(),
        after.as_ref(),
        path,
        &mut budget,
    )?;
    if native_output.archive_bytes.len() > native_bound.archive_bytes
        || SnappyStream::maximum_compressed_len(native_output.archive_bytes.len())
            .map_err(|_| Error::Verification)?
            > native_bound.compressed_bytes
    {
        return Err(Error::LimitExceeded {
            kind: LimitKind::PayloadBytes,
            observed: u64::try_from(native_output.archive_bytes.len()).unwrap_or(u64::MAX),
            maximum: u64::try_from(native_bound.archive_bytes).unwrap_or(u64::MAX),
            path,
        });
    }
    let previews = super::table_headers::rewrite::root_preview_deletions(catalog)
        .map_err(|_| Error::InvalidSource { path })?;
    // Reassembly preparation allocates its plan/index scratch.  Precharge a
    // bounded private staging allowance before asking the archive package to
    // build that plan; the exact requirements are charged immediately after
    // preparation and before its output buffer is allocated.
    budget.charge_allocations(2, path)?;
    budget.charge_transaction_work(catalog.source_bytes().len().saturating_mul(2), path)?;
    let compressed =
        SnappyStream::compress(&native_output.archive_bytes).map_err(|_| Error::Verification)?;
    let edits = [EntryEdit::new(&native_output.member_name, &compressed)];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &previews, catalog.limits())
        .map_err(|_| Error::Verification)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_reassembly(requirements, path)?;
    // Candidate parsing/readback is a second private phase.  It is charged
    // before execute so an allocation or semantic reread limit cannot occur
    // after the final ZIP buffer has been published.
    budget.charge_allocations(2, path)?;
    budget.charge_transaction_work(
        requirements
            .output_bytes()
            .saturating_mul(2)
            .saturating_add(catalog.package().iter().count().saturating_mul(1024)),
        path,
    )?;
    let limits = requirements.exact_limits();
    let bytes = prepared.execute(limits).map_err(|_| Error::Verification)?;
    let candidate = Package::from_owned_bytes_with_options(bytes, source.state.options)
        .map_err(|_| Error::Verification)?;
    let candidate_value = candidate.table_cell_control_format(
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )?;
    if candidate_value != after {
        return Err(Error::Verification);
    }
    verify_scalar_package_locality(source, &candidate, &native_output.member_name, &previews)?;
    let target_owner = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?
        .__source_owner();
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
            touched_components: 1,
            deleted_previews: previews.len(),
            full_reparse_performed: true,
        },
    })
}

fn verify_scalar_package_locality(
    source: &Package,
    candidate: &Package,
    changed_member: &str,
    deleted_previews: &[&str],
) -> Result<(), Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    for source_entry in source_catalog.package().iter() {
        if source_entry.name() == changed_member || deleted_previews.contains(&source_entry.name())
        {
            continue;
        }
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or(Error::Verification)?;
        if candidate_entry.data() != source_entry.data() {
            return Err(Error::Verification);
        }
    }
    for candidate_entry in candidate_catalog.package().iter() {
        if candidate_entry.name() == changed_member {
            continue;
        }
        let source_entry = source_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == candidate_entry.name());
        if source_entry.is_none() && !deleted_previews.contains(&candidate_entry.name()) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}
