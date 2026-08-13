use super::super::{Error as ReadError, Package};
use super::error::{map_candidate_read_error, map_read_error};
use super::resolve::{resolve_target, settings_at_target};
use super::rewrite::{
    physical_source, preflight_transaction_work, root_preview_deletions, selected_payload,
    verify_exact_locality,
};
use super::{Commit, Diagnostics, Edit, Error, Patch};
use crate::{
    selector::{SheetSelector, TableSelector},
    table::headers::Settings,
};

impl Package {
    /// Read one rooted table's lossless header and footer settings.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, allocation, or resource error.
    ///
    /// # Costs
    ///
    /// Uses the retained semantic and native indexes and strictly scans the
    /// selected model payload once.
    pub fn table_header_settings<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Settings, Error> {
        let (sheet_position, table_position) = self.resolve_header_selectors(sheet, table)?;
        Ok(resolve_target(self, sheet_position, table_position)?.settings)
    }

    /// Start a selector-first immutable header and footer settings edit.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, allocation, or resource error.
    ///
    /// # Costs
    ///
    /// Borrows this package and strictly scans the selected model once; it
    /// does not copy package bytes.
    pub fn edit_table_headers<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let (sheet_position, table_position) = self.resolve_header_selectors(sheet, table)?;
        let target = resolve_target(self, sheet_position, table_position)?;
        let before = target.settings;
        Ok(Edit {
            source: self,
            sheet_position,
            table_position,
            before,
            settings: before,
            target,
        })
    }

    /// Apply an exact-source-checked reversible header settings patch.
    ///
    /// # Errors
    ///
    /// Returns a conflict when this is not the retained exact source. A valid
    /// changed target must reopen and reproduce its requested semantic state.
    ///
    /// # Costs
    ///
    /// A no-op shares this snapshot. A changed patch fully reopens one retained
    /// target artifact and verifies its semantic state and physical locality.
    pub fn apply_table_headers(&self, patch: &Patch) -> Result<Commit, Error> {
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
        if settings_at_target(self, patch.target)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.source_payload.as_deref() != Some(selected_payload(self, patch.target)?) {
            return Err(Error::PatchConflict);
        }
        if !source_catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let target_bytes = patch.artifacts.target_owner();
        preflight_transaction_work(self, Some(target_bytes.as_ref()))?;
        let candidate = Package::from_source_owner_with_options(target_bytes, self.state.options)
            .map_err(map_candidate_read_error)?;
        if settings_at_target(&candidate, patch.target)? != patch.after {
            return Err(Error::Verification);
        }
        let source_previews = root_preview_deletions(source_catalog)?;
        if source_previews.len() != patch.source_previews {
            return Err(Error::PatchConflict);
        }
        verify_exact_locality(
            self,
            &candidate,
            patch.target,
            &source_previews,
            patch.target_previews,
            patch
                .target_payload
                .as_deref()
                .ok_or(Error::PatchConflict)?,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }

    fn resolve_header_selectors<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<(usize, usize), Error> {
        let selected_sheet = self
            .state
            .document
            .sheet(sheet)
            .map_err(|error| map_read_error(ReadError::Semantic(error)))?
            .ok_or(Error::SheetNotFound)?;
        let table_position = match table.into() {
            TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
            TableSelector::Name(name) => selected_sheet
                .tables()
                .position(|candidate_table| candidate_table.name() == name),
        }
        .ok_or(Error::TableNotFound)?;
        Ok((selected_sheet.index(), table_position))
    }
}
