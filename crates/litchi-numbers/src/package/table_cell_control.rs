//! Selector-first ownership for all interactive Numbers cell controls.
//!
//! This module is the single public transaction boundary for checkbox,
//! star-rating, slider, stepper, and menu cells.  The specialized menu graph
//! implementation remains the audited native owner; this facade wraps it
//! alongside the scalar-control native path kept in the sibling private
//! module.  No native identifiers or generated messages are present in the
//! public transaction values.
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

/// Content-redacted errors shared by every cell-control transaction.
///
/// The historical error variants remain reexported for source compatibility;
/// their values never contain authored labels, package bytes, or native IDs.
pub use popup::Error;

/// Resource axes shared by all cell-control transactions.
pub use popup::LimitKind;

/// Selector path shared by all cell-control transactions.
pub use popup::Path;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RedactedControlKind {
    None,
    Checkbox,
    StarRating,
    Slider,
    Stepper,
    PopUpMenu,
}

fn redacted_control_kind(value: Option<&CellControl>) -> RedactedControlKind {
    match value {
        None => RedactedControlKind::None,
        Some(CellControl::Checkbox(_)) => RedactedControlKind::Checkbox,
        Some(CellControl::StarRating(_)) => RedactedControlKind::StarRating,
        Some(CellControl::Slider(_)) => RedactedControlKind::Slider,
        Some(CellControl::Stepper(_)) => RedactedControlKind::Stepper,
        Some(CellControl::PopUpMenu(_)) => RedactedControlKind::PopUpMenu,
    }
}

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
            .field("before", &redacted_control_kind(self.before.as_ref()))
            .field("after", &redacted_control_kind(self.after.as_ref()))
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
            .field("before", &redacted_control_kind(self.before.as_ref()))
            .field("after", &redacted_control_kind(self.after.as_ref()))
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
        budget.charge_candidate_input_bytes(target_owner.as_ref().len(), patch.path)?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
            .map_err(|_| Error::Verification)?;
        budget.charge_candidate_reopen(candidate_catalog, patch.path)?;
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
        let touched_components = changed_component_count(self, &candidate)?;
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
    let target = popup::resolve_cell(
        source,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )?;
    // Re-check the complete semantic control at publication time.  The
    // focused popup transaction only observes the popup projection, so this
    // guard also protects scalar -> popup transitions from publishing over a
    // stale scalar edit.
    if popup::read_cell_control(source, target)? != before {
        return Err(Error::PatchConflict);
    }
    let popup_before = match &before {
        Some(CellControl::PopUpMenu(value)) => Some(value.clone()),
        _ => None,
    };
    let popup_after = match &after {
        Some(CellControl::PopUpMenu(value)) => Some(value.clone()),
        _ => None,
    };
    if popup_before.is_some() || popup_after.is_some() {
        // A popup -> scalar transition has two distinct graph owners: first
        // release the popup model through its focused transaction, then run
        // the already-audited scalar writer against that private candidate.
        // Both candidates remain private; the outer patch is built only from
        // the original source and final package, preserving atomicity and an
        // exact inverse for the unified facade.
        if popup_before.is_some() && popup_after.is_none() && after.is_some() {
            let cleared = source
                .edit_table_cell_pop_up_menu_format(
                    SheetSelector::index(sheet),
                    TableSelector::index(table),
                    position,
                )?
                .clear()
                .commit()?;
            let cleared_package = cleared.into_package();
            let scalar_commit = rewrite(&cleared_package, path, None, after.clone())?;
            let package = scalar_commit.into_package();
            if package.table_cell_control_format(
                SheetSelector::index(sheet),
                TableSelector::index(table),
                position,
            )? != after
            {
                return Err(Error::Verification);
            }
            let diagnostics = package_diff_diagnostics(source, &package)?;
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
                diagnostics,
            });
        }
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
        if package.table_cell_control_format(
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )? != after
        {
            return Err(Error::Verification);
        }
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
    for member in &native_output.members {
        if member.archive_bytes.len() > native_bound.archive_bytes
            || SnappyStream::maximum_compressed_len(member.archive_bytes.len())
                .map_err(|_| Error::Verification)?
                > native_bound.compressed_bytes
        {
            return Err(Error::LimitExceeded {
                kind: LimitKind::PayloadBytes,
                observed: u64::try_from(member.archive_bytes.len()).unwrap_or(u64::MAX),
                maximum: u64::try_from(native_bound.archive_bytes).unwrap_or(u64::MAX),
                path,
            });
        }
        let maximum_compressed = SnappyStream::maximum_compressed_len(member.archive_bytes.len())
            .map_err(|_| Error::Verification)?;
        budget.charge_compressed_bytes(maximum_compressed, path)?;
    }
    let metadata_bytes = popup::rewrite_component_save_tokens(
        source,
        &native_output.component_indices,
        path,
        &mut budget,
    )?;
    let previews = super::table_headers::rewrite::root_preview_deletions(catalog)
        .map_err(|_| Error::InvalidSource { path })?;
    // Reassembly preparation allocates its plan/index scratch.  Precharge a
    // bounded private staging allowance before asking the archive package to
    // build that plan; the exact requirements are charged immediately after
    // preparation and before its output buffer is allocated.
    budget.charge_allocations(2, path)?;
    budget.charge_transaction_work(catalog.source_bytes().len().saturating_mul(2), path)?;
    let compressed_members = native_output
        .members
        .iter()
        .map(|member| {
            SnappyStream::compress(&member.archive_bytes).map_err(|_| Error::Verification)
        })
        .collect::<Result<Vec<_>, _>>()?;
    let mut changed_names = native_output
        .members
        .iter()
        .map(|member| member.member_name.clone())
        .collect::<Vec<_>>();
    changed_names.push(super::metadata::ENTRY_NAME.to_owned());
    let mut edit_buffers = native_output
        .members
        .iter()
        .zip(compressed_members)
        .map(|(member, compressed)| (member.member_name.clone(), compressed))
        .collect::<Vec<_>>();
    edit_buffers.push((super::metadata::ENTRY_NAME.to_owned(), metadata_bytes));
    let edits = edit_buffers
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes))
        .collect::<Vec<_>>();
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
    budget.charge_candidate_input_bytes(requirements.output_bytes(), path)?;
    let bytes = prepared.execute(limits).map_err(|_| Error::Verification)?;
    let candidate = Package::from_owned_bytes_with_options(bytes, source.state.options)
        .map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?;
    budget.charge_candidate_reopen(candidate_catalog, path)?;
    let candidate_value = candidate.table_cell_control_format(
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )?;
    if candidate_value != after {
        return Err(Error::Verification);
    }
    verify_scalar_package_locality(
        source,
        &candidate,
        &changed_names,
        &previews,
        &native_output.changed_objects,
    )?;
    let target_owner = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?
        .__source_owner();
    let touched_components = changed_component_count(source, &candidate)?;
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
            deleted_previews: previews.len(),
            full_reparse_performed: true,
        },
    })
}

fn verify_scalar_package_locality(
    source: &Package,
    candidate: &Package,
    changed_members: &[String],
    deleted_previews: &[&str],
    changed_objects: &[(usize, u64)],
) -> Result<(), Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    for source_entry in source_catalog.package().iter() {
        if changed_members
            .iter()
            .any(|name| name == source_entry.name())
            || deleted_previews.contains(&source_entry.name())
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
        if changed_members
            .iter()
            .any(|name| name == candidate_entry.name())
        {
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
    for &(component_index, _) in changed_objects {
        let source_component = source
            .state
            .components
            .catalog()
            .get_index(component_index)
            .ok_or(Error::Verification)?;
        let candidate_component_index = candidate
            .state
            .components
            .catalog()
            .iter()
            .position(|component| component.name() == source_component.name())
            .ok_or(Error::Verification)?;
        let candidate_component = candidate
            .state
            .components
            .catalog()
            .get_index(candidate_component_index)
            .ok_or(Error::Verification)?;
        let changed_in_component = changed_objects
            .iter()
            .filter_map(|(owner, identifier)| (*owner == component_index).then_some(*identifier))
            .collect::<Vec<_>>();
        for source_object in &source_component.archive().objects {
            let identifier = source_object
                .archive_info
                .identifier
                .ok_or(Error::Verification)?;
            if changed_in_component.contains(&identifier) {
                continue;
            }
            let candidate_object = candidate_component
                .archive()
                .objects
                .iter()
                .find(|object| object.archive_info.identifier == Some(identifier))
                .ok_or(Error::Verification)?;
            if !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(Error::Verification);
            }
        }
    }
    Ok(())
}

fn package_diff_diagnostics(source: &Package, candidate: &Package) -> Result<Diagnostics, Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    let changed_members = changed_member_names(source_catalog, candidate_catalog);
    let deleted_previews = source_catalog
        .package()
        .iter()
        .filter(|entry| entry.name().starts_with("preview"))
        .filter(|source_entry| {
            candidate_catalog
                .package()
                .iter()
                .all(|candidate_entry| candidate_entry.name() != source_entry.name())
        })
        .count();
    Ok(Diagnostics {
        changed: !changed_members.is_empty() || deleted_previews != 0,
        touched_components: changed_component_count(source, candidate)?,
        deleted_previews,
        full_reparse_performed: true,
    })
}

fn changed_member_names(
    source: &litchi_iwa_archive::SourceCatalog,
    candidate: &litchi_iwa_archive::SourceCatalog,
) -> Vec<String> {
    let mut changed_members = Vec::new();
    for source_entry in source.package().iter() {
        let changed = candidate
            .package()
            .iter()
            .find(|candidate_entry| candidate_entry.name() == source_entry.name())
            .is_none_or(|candidate_entry| candidate_entry.data() != source_entry.data());
        if changed && !source_entry.name().starts_with("preview") {
            changed_members.push(source_entry.name().to_owned());
        }
    }
    for candidate_entry in candidate.package().iter() {
        let added = source
            .package()
            .iter()
            .all(|source_entry| source_entry.name() != candidate_entry.name());
        if added && !candidate_entry.name().starts_with("preview") {
            changed_members.push(candidate_entry.name().to_owned());
        }
    }
    changed_members.sort_unstable();
    changed_members.dedup();
    changed_members
}

fn changed_component_count(source: &Package, candidate: &Package) -> Result<usize, Error> {
    let source_components = source.state.components.catalog();
    let candidate_components = candidate.state.components.catalog();
    let mut changed = 0usize;
    for source_component in source_components.iter() {
        let candidate_component = candidate_components
            .iter()
            .find(|component| component.name() == source_component.name());
        let differs = candidate_component.is_none_or(|candidate_component| {
            source_component.archive().objects.len() != candidate_component.archive().objects.len()
                || source_component
                    .archive()
                    .objects
                    .iter()
                    .zip(candidate_component.archive().objects.iter())
                    .any(|(source_object, candidate_object)| {
                        !source_object.same_content_ignoring_offsets(candidate_object)
                    })
        });
        if differs {
            changed = changed.saturating_add(1);
        }
    }
    changed = changed.saturating_add(
        candidate_components
            .iter()
            .filter(|candidate_component| {
                source_components
                    .iter()
                    .all(|source_component| source_component.name() != candidate_component.name())
            })
            .count(),
    );
    // Metadata.iwa is normally part of the parsed component catalog.  Keep a
    // physical fallback for semantic/legacy sources that retain the metadata
    // sidecar outside that catalog so its token transition is never omitted
    // from the diagnostic count.
    let metadata_in_catalog = source_components
        .iter()
        .chain(candidate_components.iter())
        .any(|component| component.name() == super::metadata::ENTRY_NAME);
    if !metadata_in_catalog {
        let source_metadata = super::table_headers::rewrite::physical_source(source)
            .map_err(|_| Error::Verification)?
            .package()
            .iter()
            .find(|entry| entry.name() == super::metadata::ENTRY_NAME);
        let candidate_metadata = super::table_headers::rewrite::physical_source(candidate)
            .map_err(|_| Error::Verification)?
            .package()
            .iter()
            .find(|entry| entry.name() == super::metadata::ENTRY_NAME);
        if source_metadata.map(|entry| entry.data()) != candidate_metadata.map(|entry| entry.data())
        {
            changed = changed.saturating_add(1);
        }
    }
    Ok(changed)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::data_format::PopUpMenu;

    #[test]
    fn redacted_control_debug_does_not_include_authored_menu_items() {
        let menu = PopUpMenu::new(["private-control-label"]).expect("valid menu");
        let control = CellControl::PopUpMenu(menu);
        let rendered = format!("{:?}", redacted_control_kind(Some(&control)));
        assert!(rendered.contains("PopUpMenu"));
        assert!(!rendered.contains("private-control-label"));
    }
}
