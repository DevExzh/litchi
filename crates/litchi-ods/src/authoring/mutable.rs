//! Transactional editing of an existing spreadsheet package.

use crate::{
    Spreadsheet,
    model::names::{Definition, Expression, Range, Scope},
    worksheet::{Cell, CellView, Sheet},
};
use litchi_core::{Error, Result};
use litchi_odf_common::rdf::Triple;
use std::path::Path;

/// Legacy mutable ODS compatibility facade.
///
/// Every package-level edit is validated and atomically replaces the owned
/// immutable snapshot. Failed edits leave the document unchanged. New code
/// should prefer the source-checked `Snapshot`/`Transaction`/`Commit`/`Patch`
/// APIs of individual feature owners, or `FlatSpreadsheet` for flat-ODS cell
/// edits; this attached compatibility facade cannot safely edit every
/// source-level worksheet extension.
pub struct MutableSpreadsheet {
    spreadsheet: Spreadsheet,
}

impl MutableSpreadsheet {
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Spreadsheet::open(path).map(Self::from_spreadsheet)
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Spreadsheet::from_bytes(bytes).map(Self::from_spreadsheet)
    }

    #[must_use]
    pub fn from_spreadsheet(spreadsheet: Spreadsheet) -> Self {
        Self { spreadsheet }
    }

    #[must_use]
    pub fn spreadsheet(&self) -> &Spreadsheet {
        &self.spreadsheet
    }

    /// Capture the exact package as the unified immutable transaction owner.
    ///
    /// # Errors
    ///
    /// Returns an error when package bounds or complete facade readback fail.
    pub fn document_snapshot(&self) -> Result<crate::document::Snapshot> {
        self.spreadsheet.document_snapshot()
    }

    /// Apply one durable exact-source unified package patch.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, security refusal, package bounds, or candidate
    /// readback failure.
    pub fn apply_document_patch(&mut self, patch: &crate::document::Patch) -> Result<()> {
        self.spreadsheet.apply_document_patch(patch)
    }

    /// Clone-stage all supported ODS owners and publish them as one immutable package commit.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, security checks, durable patch, package rebuild, or
    /// complete typed readback fails. The attached facade remains unchanged on failure.
    pub fn edit_document<F>(&mut self, update: F) -> Result<crate::document::Patch>
    where
        F: FnOnce(&mut crate::document::Edit) -> Result<()>,
    {
        let snapshot = self.document_snapshot()?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        let patch = commit.patch().clone();
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(patch)
    }

    /// Borrow the compact cross-format metadata projection.
    #[must_use]
    pub fn metadata(&self) -> &litchi_core::Metadata {
        self.spreadsheet.metadata()
    }

    /// Borrow the complete typed ODF metadata model.
    #[must_use]
    pub fn odf_metadata(&self) -> &crate::metadata::Metadata {
        self.spreadsheet.odf_metadata()
    }

    /// Replace the supported metadata projection atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_metadata(&mut self, metadata: litchi_core::Metadata) -> Result<()> {
        self.spreadsheet.publish_metadata(metadata)
    }

    /// Apply a short-lived metadata update transactionally.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update_metadata<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut litchi_core::Metadata) -> Result<()>,
    {
        let mut metadata = self.metadata().clone();
        update(&mut metadata)?;
        self.set_metadata(metadata)
    }

    /// Remove the physical `meta.xml` part atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_metadata(&mut self) -> Result<()> {
        self.spreadsheet.remove_metadata()
    }

    /// Borrow spreadsheet calculation settings, if present.
    #[must_use]
    pub fn settings(&self) -> Option<&crate::settings::Settings> {
        self.spreadsheet.settings()
    }

    /// Replace or remove calculation settings atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_settings(&mut self, settings: Option<crate::settings::Settings>) -> Result<()> {
        if let Some(settings) = &settings {
            settings.validate()?;
        }
        self.spreadsheet.publish_settings(settings)
    }

    /// Apply a typed calculation-settings update, creating the owner when it
    /// is absent.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update_settings<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::settings::Settings) -> Result<()>,
    {
        let mut settings = self.settings().cloned().unwrap_or_default();
        update(&mut settings)?;
        self.set_settings(Some(settings))
    }

    /// Remove the calculation-settings element from `content.xml`.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_settings(&mut self) -> Result<()> {
        self.set_settings(None)
    }

    /// Capture the source-checked cell-annotation owner for this snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn annotations(&self) -> Result<crate::annotations::Snapshot> {
        self.spreadsheet.annotations()
    }

    /// Apply one failure-atomic cell-annotation transaction.
    ///
    /// The transaction resolves cells by exact sheet name and zero-based
    /// logical coordinates.  If the closure or commit fails, this facade and
    /// its package bytes remain unchanged; an empty commit does not rebuild
    /// the archive.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn edit_annotations<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut crate::annotations::Transaction) -> Result<()>,
    {
        let snapshot = self.spreadsheet.annotations()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        if commit.changed() {
            let content_xml = commit.content_xml().to_owned();
            self.spreadsheet.publish_annotations(&content_xml)?;
        }
        Ok(())
    }

    /// Return the typed worksheet graph in document order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        self.spreadsheet.sheets()
    }

    /// Capture the exact-source worksheet transaction owner.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package or worksheet graph is invalid.
    pub fn worksheet_snapshot(&self) -> Result<crate::worksheet::Snapshot> {
        self.spreadsheet.worksheet_snapshot()
    }

    /// Apply an exact-source reversible worksheet patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package.
    pub fn apply_worksheet_patch(&mut self, patch: &crate::worksheet::Patch) -> Result<()> {
        self.spreadsheet.apply_worksheet_patch(patch)
    }

    /// Clone-stage worksheet structure and cell CRUD as one package edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, validation, compactness check, rebuild, or readback
    /// fails.
    pub fn edit_worksheets<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut crate::worksheet::Edit) -> Result<()>,
    {
        let snapshot = self.worksheet_snapshot()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Discover embedded charts in the current immutable package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn charts(&self) -> Result<crate::charts::Inventory<'_>> {
        self.spreadsheet.charts()
    }

    /// Capture embedded charts as an explicit immutable transaction snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart_snapshot(&self) -> Result<crate::charts::Snapshot> {
        self.spreadsheet.chart_snapshot()
    }

    /// Apply an exact-source embedded-chart patch and rehydrate the spreadsheet.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply_chart_patch(&mut self, patch: &crate::charts::Patch) -> Result<()> {
        let commit = patch.apply(&self.chart_snapshot()?)?;
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Clone-stage chart replacements and publish them as one package edit.
    ///
    /// A failed closure or commit leaves this mutable facade unchanged. An
    /// empty transaction does not rebuild the archive, preserving exact bytes.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn edit_charts<F>(&mut self, edit: F) -> Result<()>
    where
        F: for<'source> FnOnce(&mut crate::charts::Transaction<'source>) -> Result<()>,
    {
        let inventory = self.spreadsheet.charts()?;
        let mut transaction = inventory.transaction();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.into_owned_bytes())?;
        }
        Ok(())
    }

    /// Discover the typed `DataPilot` catalog in the current package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilots(&self) -> Result<crate::data_pilot::Catalog<'_>> {
        self.spreadsheet.data_pilots()
    }

    /// Capture an immutable, exact-source `DataPilot` snapshot for explicit
    /// `Snapshot` → `Edit` → `Commit` → `Patch` workflows.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilot_snapshot(&self) -> Result<crate::data_pilot::Snapshot> {
        self.spreadsheet.data_pilot_snapshot()
    }

    /// Apply an exact-source `DataPilot` patch and rehydrate the full spreadsheet.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn apply_data_pilot_patch(&mut self, patch: &crate::data_pilot::Patch) -> Result<()> {
        let commit = patch.apply(&self.data_pilot_snapshot()?)?;
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Clone-stage `DataPilot` CRUD and publish it as one package edit.
    ///
    /// Unknown markup in the owned XML is retained by no-op transactions and
    /// causes a changed transaction to fail before package bytes are rebuilt.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn edit_data_pilots<F>(&mut self, edit: F) -> Result<()>
    where
        F: for<'source> FnOnce(&mut crate::data_pilot::Editor<'_, 'source>) -> Result<()>,
    {
        let commit = {
            let catalog = self.spreadsheet.data_pilots()?;
            let mut transaction = catalog.transaction();
            edit(&mut transaction.editor())?;
            transaction.commit()?
        };
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.into_owned_bytes())?;
        }
        Ok(())
    }

    /// Find a worksheet by its exact ODF name.
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.spreadsheet.sheet(name)
    }

    /// Look up a logical cell in the current immutable snapshot.
    #[must_use]
    pub fn cell(&self, sheet_name: &str, row: usize, column: usize) -> Option<CellView<'_>> {
        self.spreadsheet.cell(sheet_name, row, column)
    }

    /// Atomically replace all worksheets in the current package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_sheets(&mut self, sheets: Vec<Sheet>) -> Result<()> {
        self.spreadsheet.publish_sheets(sheets)
    }

    /// Append one worksheet while preserving document order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_sheet(&mut self, sheet: Sheet) -> Result<()> {
        self.edit_sheets(move |sheets| {
            sheets.push(sheet);
            Ok(())
        })
    }

    /// Remove one worksheet by its exact ODF name.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove_sheet(&mut self, name: &str) -> Result<Sheet> {
        let mut removed = None;
        self.edit_sheets(|sheets| {
            let index = sheets
                .iter()
                .position(|sheet| sheet.name == name)
                .ok_or_else(|| Error::InvalidFormat(format!("ODS sheet '{name}' was not found")))?;
            removed = Some(sheets.remove(index));
            Ok(())
        })?;
        removed
            .ok_or_else(|| Error::InvalidFormat("ODS sheet removal was not committed".to_string()))
    }

    /// Atomically replace one logical cell in a named worksheet.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell(
        &mut self,
        sheet_name: &str,
        row: usize,
        column: usize,
        cell: Cell,
    ) -> Result<()> {
        self.edit_sheet(sheet_name, |sheet| sheet.set_cell(row, column, cell))
    }

    /// Clear one logical cell while retaining its direct style, if any.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_cell(&mut self, sheet_name: &str, row: usize, column: usize) -> Result<()> {
        self.edit_sheet(sheet_name, |sheet| sheet.clear_cell(row, column))
    }

    /// Set an inert formula on one logical cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_formula(
        &mut self,
        sheet_name: &str,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<()> {
        let formula = formula.into();
        self.edit_sheet(sheet_name, move |sheet| {
            sheet.set_formula(row, column, formula)
        })
    }

    /// Set a direct cell style reference on one logical cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell_style(
        &mut self,
        sheet_name: &str,
        row: usize,
        column: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let style_name = style_name.into();
        self.edit_sheet(sheet_name, move |sheet| {
            sheet.set_cell_style(row, column, style_name)
        })
    }

    fn edit_sheets<F>(&mut self, operation: F) -> Result<()>
    where
        F: FnOnce(&mut Vec<Sheet>) -> Result<()>,
    {
        let candidate = crate::worksheet::transaction::edit(self.spreadsheet.sheets(), operation)?;
        self.spreadsheet.publish_sheets(candidate)
    }

    fn edit_sheet<F>(&mut self, sheet_name: &str, operation: F) -> Result<()>
    where
        F: FnOnce(&mut Sheet) -> Result<()>,
    {
        self.edit_sheets(|sheets| {
            let sheet = sheets
                .iter_mut()
                .find(|sheet| sheet.name == sheet_name)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!("ODS sheet '{sheet_name}' was not found"))
                })?;
            operation(sheet)
        })
    }

    /// Return the ordered named-definition catalog of the edited snapshot.
    #[must_use]
    pub fn definitions(&self) -> &[Definition] {
        self.spreadsheet.definitions()
    }

    /// Capture the exact-source named-definition transaction owner.
    ///
    /// # Errors
    ///
    /// Returns an error when the retained package cannot be reparsed.
    pub fn definitions_snapshot(&self) -> Result<crate::definitions::Snapshot> {
        self.spreadsheet.definitions_snapshot()
    }

    /// Apply an exact-source reversible named-definition patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package.
    pub fn apply_definitions_patch(&mut self, patch: &crate::definitions::Patch) -> Result<()> {
        self.spreadsheet.apply_definitions_patch(patch)
    }

    /// Clone-stage ordered definition CRUD and publish it as one package edit.
    ///
    /// An unchanged edit retains the original package bytes. A changed edit must produce compact
    /// `content.xml`, pass a full package reopen, and match the staged typed catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, validation, compactness check, rebuild, or readback fails.
    pub fn edit_definitions<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut crate::definitions::Edit) -> Result<()>,
    {
        let snapshot = self.definitions_snapshot()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        if commit.changed() {
            self.spreadsheet = Spreadsheet::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Append a validated named range atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_range(&mut self, range: Range) -> Result<()> {
        self.spreadsheet.add_range(range)
    }

    /// Append a validated named expression atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_expression(&mut self, expression: Expression) -> Result<()> {
        self.spreadsheet.add_expression(expression)
    }

    /// Append a validated named definition atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_definition(&mut self, definition: Definition) -> Result<()> {
        self.spreadsheet.add_definition(definition)
    }

    /// Replace the complete ordered named-definition catalog atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_definitions(&mut self, definitions: Vec<Definition>) -> Result<()> {
        self.spreadsheet.set_definitions(definitions)
    }

    /// Find a named range in the current snapshot.
    #[must_use]
    pub fn range(&self, name: &str, scope: &Scope) -> Option<&Range> {
        self.spreadsheet.range(name, scope)
    }

    /// Find a named expression in the current snapshot.
    #[must_use]
    pub fn expression(&self, name: &str, scope: &Scope) -> Option<&Expression> {
        self.spreadsheet.expression(name, scope)
    }

    /// Add a validated RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the path, triples, compact XML, or rebuilt package is invalid.
    pub fn add_rdf_graph(&mut self, path: Option<&str>, triples: &[Triple]) -> Result<String> {
        self.spreadsheet.add_rdf_graph(path, triples)
    }

    /// Capture RDF metadata as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the package, manifest, or a declared graph is invalid.
    pub fn rdf_snapshot(&self) -> Result<crate::metadata_graphs::Snapshot> {
        self.spreadsheet.rdf_snapshot()
    }

    /// Apply an exact-source reversible RDF graph patch.
    ///
    /// # Errors
    ///
    /// Returns an error for a stale patch or invalid candidate package.
    pub fn apply_rdf_patch(&mut self, patch: &crate::metadata_graphs::Patch) -> Result<()> {
        self.spreadsheet.apply_rdf_patch(patch)
    }

    /// Clone-stage RDF graph and triple CRUD and publish it as one package edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or a staged package rebuild fails.
    pub fn edit_rdf<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut crate::metadata_graphs::Edit) -> Result<()>,
    {
        let snapshot = self.rdf_snapshot()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        self.spreadsheet
            .apply_rdf_patch(transaction.commit().patch())
    }

    /// Replace one complete RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, triples, compact XML, or rebuilt package is invalid.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> Result<()> {
        self.spreadsheet.replace_rdf_graph(path, triples)
    }

    /// Remove one unreferenced RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is missing, referenced, or package rebuilding fails.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        self.spreadsheet.remove_rdf_graph(path)
    }

    #[must_use]
    pub fn to_bytes(self) -> Vec<u8> {
        self.spreadsheet.into_bytes()
    }

    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn save(self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.to_bytes()).map_err(Into::into)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Builder;

    #[test]
    fn preserves_owned_package_without_edits() {
        let bytes = Builder::new().build().unwrap();
        let mutable = MutableSpreadsheet::from_bytes(bytes).unwrap();
        let reopened = Spreadsheet::from_bytes(mutable.to_bytes()).unwrap();
        assert!(reopened.content_xml().contains("office:spreadsheet"));
    }
}
