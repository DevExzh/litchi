//! A builder draft over an existing spreadsheet (record 0358).
//!
//! [`Workbook`] starts from a read-only [`crate::Workbook`] projection, stages
//! every change against an owned candidate and [`Workbook::build`]s an
//! immutable [`crate::Package`]. Opened packages edit through
//! [`crate::Package::edit`]; this draft keeps the closure-scoped part
//! editors (`edit_worksheets`, `edit_charts`, `update_protection`, ...) for
//! callers that assemble a workbook before it is published.

use crate::{
    Workbook as Projection,
    model::names::{Definition, Expression, Range, Scope},
    worksheet::{Cell, CellView, Sheet},
};
use litchi_core::{Error as CoreError, Result as CoreResult};
use litchi_odf_common::rdf::Triple;

type FamilyResult<T> = Result<T, litchi_odf_common::Error>;

/// A builder draft over an existing spreadsheet.
///
/// Every staged change is validated and atomically replaces the owned
/// candidate; a failed change leaves the draft unchanged. [`Self::build`]
/// publishes the candidate as an immutable [`crate::Package`]. Opened
/// packages edit through [`crate::Package::edit`], flat documents through
/// [`crate::flat::Workbook`].
pub struct Workbook {
    spreadsheet: Projection,
}

impl Workbook {
    /// Start a draft from a read-only projection.
    #[must_use]
    pub fn from_workbook(workbook: Projection) -> Self {
        Self {
            spreadsheet: workbook,
        }
    }

    /// Borrow the current read-only projection.
    #[must_use]
    pub fn workbook(&self) -> &Projection {
        &self.spreadsheet
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
    pub fn set_metadata(&mut self, metadata: litchi_core::Metadata) -> FamilyResult<()> {
        Ok(self.spreadsheet.publish_metadata(metadata)?)
    }

    /// Apply a short-lived metadata update transactionally.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update_metadata<F>(&mut self, update: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut litchi_core::Metadata) -> CoreResult<()>,
    {
        let mut metadata = self.metadata().clone();
        update(&mut metadata)?;
        self.set_metadata(metadata)
    }

    /// Remove the physical `meta.xml` part atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_metadata(&mut self) -> FamilyResult<()> {
        Ok(self.spreadsheet.remove_metadata()?)
    }

    /// Capture the standalone table-template catalog.
    pub fn table_templates(&self) -> FamilyResult<crate::styles::table_template::Snapshot> {
        self.spreadsheet.table_templates()
    }

    /// Stage and publish one source-checked table-template edit.
    pub fn edit_table_templates<F>(&mut self, update: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::styles::table_template::Edit) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.edit_table_templates(update)?)
    }

    /// Capture inert DDE source declarations and cached tables.
    pub fn dde(&self) -> FamilyResult<crate::dde::Snapshot> {
        self.spreadsheet.dde()
    }

    /// Capture DDE metadata under explicit limits and execution policy.
    pub fn dde_with(
        &self,
        limits: crate::dde::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> FamilyResult<crate::dde::Snapshot> {
        self.spreadsheet.dde_with(limits, context)
    }

    /// Publish a failure-atomic inert DDE metadata edit.
    pub fn edit_dde<F>(&mut self, update: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::dde::Edit) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.edit_dde(update)?)
    }

    /// Edit DDE metadata with explicit parsing, staging, and readback budgets.
    /// Package replacement and rehydration retain the package policy.
    pub fn edit_dde_with_context<F>(
        &mut self,
        limits: crate::dde::Limits,
        context: &litchi_core::ExecutionContext,
        update: F,
    ) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::dde::Edit) -> CoreResult<()>,
    {
        Ok(self
            .spreadsheet
            .edit_dde_with_context(limits, context, update)?)
    }

    /// Capture the inert, source-bound scenario metadata catalog.
    pub fn scenarios(&self) -> FamilyResult<crate::scenario::Snapshot> {
        self.spreadsheet.scenarios()
    }

    /// Stage and publish one failure-atomic scenario metadata edit.
    pub fn edit_scenarios<F>(&mut self, update: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::scenario::Edit) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.edit_scenarios(update)?)
    }

    /// Capture the source-backed sheet metadata catalog.
    pub fn sheet_metadata(&self) -> FamilyResult<crate::sheet_metadata::Snapshot> {
        self.spreadsheet.sheet_metadata()
    }

    /// Capture sheet metadata under explicit finite limits and context.
    pub fn sheet_metadata_with(
        &self,
        limits: crate::sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
    ) -> FamilyResult<crate::sheet_metadata::Snapshot> {
        self.spreadsheet.sheet_metadata_with(limits, context)
    }

    /// Stage and publish one failure-atomic sheet metadata edit.
    pub fn edit_sheet_metadata<F>(&mut self, update: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::sheet_metadata::Edit) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.edit_sheet_metadata(update)?)
    }

    /// Stage and publish sheet metadata under explicit limits and context.
    ///
    /// The supplied context governs metadata parsing, staging, candidate
    /// rendering, and target readback. Owned package replacement and facade
    /// rehydration use the package's own bounded policy; the context is
    /// checked again before the attached spreadsheet is replaced.
    pub fn edit_sheet_metadata_with_context<F>(
        &mut self,
        limits: crate::sheet_metadata::Limits,
        context: &litchi_core::ExecutionContext,
        update: F,
    ) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::sheet_metadata::Edit) -> CoreResult<()>,
    {
        Ok(self
            .spreadsheet
            .edit_sheet_metadata_with_context(limits, context, update)?)
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
    pub fn set_settings(
        &mut self,
        settings: Option<crate::settings::Settings>,
    ) -> FamilyResult<()> {
        if let Some(settings) = &settings {
            settings.validate()?;
        }
        Ok(self.spreadsheet.publish_settings(settings)?)
    }

    /// Apply a typed calculation-settings update, creating the owner when it
    /// is absent.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn update_settings<F>(&mut self, update: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::settings::Settings) -> CoreResult<()>,
    {
        let mut settings = self.settings().cloned().unwrap_or_default();
        update(&mut settings)?;
        self.set_settings(Some(settings))
    }

    /// Remove the calculation-settings element from `content.xml`.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_settings(&mut self) -> FamilyResult<()> {
        self.set_settings(None)
    }

    /// Capture the source-checked cell-annotation owner for this snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn annotations(&self) -> FamilyResult<crate::annotations::Snapshot> {
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
    pub fn edit_annotations<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::annotations::Edit) -> CoreResult<()>,
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
    pub fn worksheet_snapshot(&self) -> FamilyResult<crate::worksheet::Snapshot> {
        self.spreadsheet.worksheet_snapshot()
    }

    /// Clone-stage worksheet structure and cell CRUD as one package edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, validation, compactness check, rebuild, or readback
    /// fails.
    pub fn edit_worksheets<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::worksheet::Edit) -> CoreResult<()>,
    {
        let snapshot = self.worksheet_snapshot()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        if commit.changed() {
            self.spreadsheet = Projection::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Discover embedded charts in the current immutable package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn charts(&self) -> FamilyResult<crate::charts::Inventory<'_>> {
        self.spreadsheet.charts()
    }

    /// Capture embedded charts as an explicit immutable transaction snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn chart_snapshot(&self) -> FamilyResult<crate::charts::Snapshot> {
        self.spreadsheet.chart_snapshot()
    }

    /// Clone-stage chart replacements and publish them as one package edit.
    ///
    /// A failed closure or commit leaves this mutable facade unchanged. An
    /// empty transaction does not rebuild the archive, preserving exact bytes.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn edit_charts<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: for<'source> FnOnce(&mut crate::charts::Editor<'_, 'source>) -> CoreResult<()>,
    {
        let inventory = self.spreadsheet.charts()?;
        let mut transaction = inventory.transaction();
        edit(&mut transaction.editor())?;
        let commit = transaction.commit()?;
        if commit.changed() {
            self.spreadsheet = Projection::from_bytes(commit.into_owned_bytes())?;
        }
        Ok(())
    }

    /// Discover the typed `DataPilot` catalog in the current package snapshot.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilots(&self) -> FamilyResult<crate::data_pilot::Catalog<'_>> {
        self.spreadsheet.data_pilots()
    }

    /// Capture an immutable, exact-source `DataPilot` snapshot for explicit
    /// `Snapshot` → `Edit` → `Commit` → `Patch` workflows.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn data_pilot_snapshot(&self) -> FamilyResult<crate::data_pilot::Snapshot> {
        self.spreadsheet.data_pilot_snapshot()
    }

    /// Clone-stage `DataPilot` CRUD and publish it as one package edit.
    ///
    /// Unknown markup in the owned XML is retained by no-op transactions and
    /// causes a changed transaction to fail before package bytes are rebuilt.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn edit_data_pilots<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: for<'source> FnOnce(&mut crate::data_pilot::Editor<'_, 'source>) -> CoreResult<()>,
    {
        let commit = {
            let catalog = self.spreadsheet.data_pilots()?;
            let mut transaction = catalog.transaction();
            edit(&mut transaction.editor())?;
            transaction.commit()?
        };
        if commit.changed() {
            self.spreadsheet = Projection::from_bytes(commit.into_owned_bytes())?;
        }
        Ok(())
    }

    /// Discover inert database-range declarations in the current package.
    ///
    /// No database, query, refresh, or filter execution occurs.
    pub fn database_ranges(&self) -> FamilyResult<crate::database_range::Catalog<'_>> {
        self.spreadsheet.database_ranges()
    }

    /// Capture an immutable, exact-source database-range snapshot.
    pub fn database_range_snapshot(&self) -> FamilyResult<crate::database_range::Snapshot> {
        self.spreadsheet.database_range_snapshot()
    }

    /// Alias for [`Self::database_range_snapshot`].
    pub fn database_ranges_snapshot(&self) -> FamilyResult<crate::database_range::Snapshot> {
        self.database_range_snapshot()
    }

    /// Clone-stage inert database-range CRUD and publish one atomic package
    /// edit.
    pub fn edit_database_ranges<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: for<'source> FnOnce(&mut crate::database_range::Editor<'_, 'source>) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.edit_database_ranges(edit)?)
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
    pub fn set_sheets(&mut self, sheets: Vec<Sheet>) -> FamilyResult<()> {
        Ok(self.spreadsheet.publish_sheets(sheets)?)
    }

    /// Append one worksheet while preserving document order.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_sheet(&mut self, sheet: Sheet) -> FamilyResult<()> {
        Ok(self.edit_sheets(move |sheets| {
            sheets.push(sheet);
            Ok(())
        })?)
    }

    /// Remove one worksheet by its exact ODF name.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn remove_sheet(&mut self, name: &str) -> FamilyResult<Sheet> {
        let mut removed = None;
        self.edit_sheets(|sheets| {
            let index = sheets
                .iter()
                .position(|sheet| sheet.name == name)
                .ok_or_else(|| {
                    CoreError::InvalidFormat(format!("ODS sheet '{name}' was not found"))
                })?;
            removed = Some(sheets.remove(index));
            Ok(())
        })?;
        Ok(removed.ok_or_else(|| {
            CoreError::InvalidFormat("ODS sheet removal was not committed".to_string())
        })?)
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
    ) -> FamilyResult<()> {
        Ok(self.edit_sheet(sheet_name, |sheet| sheet.set_cell(row, column, cell))?)
    }

    /// Clear one logical cell while retaining its direct style, if any.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_cell(&mut self, sheet_name: &str, row: usize, column: usize) -> FamilyResult<()> {
        Ok(self.edit_sheet(sheet_name, |sheet| sheet.clear_cell(row, column))?)
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
    ) -> FamilyResult<()> {
        let formula = formula.into();
        Ok(self.edit_sheet(sheet_name, move |sheet| {
            sheet.set_formula(row, column, formula)
        })?)
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
    ) -> FamilyResult<()> {
        let style_name = style_name.into();
        Ok(self.edit_sheet(sheet_name, move |sheet| {
            sheet.set_cell_style(row, column, style_name)
        })?)
    }

    /// Set or clear one worksheet's direct table title.
    pub fn set_sheet_title(&mut self, sheet_name: &str, title: Option<String>) -> FamilyResult<()> {
        Ok(self.edit_sheet(sheet_name, |sheet| sheet.set_title(title))?)
    }

    /// Set or clear one worksheet's direct table description.
    pub fn set_sheet_description(
        &mut self,
        sheet_name: &str,
        description: Option<String>,
    ) -> FamilyResult<()> {
        Ok(self.edit_sheet(sheet_name, |sheet| sheet.set_description(description))?)
    }

    fn edit_sheets<F>(&mut self, operation: F) -> CoreResult<()>
    where
        F: FnOnce(&mut Vec<Sheet>) -> CoreResult<()>,
    {
        let candidate = crate::worksheet::transaction::edit(self.spreadsheet.sheets(), operation)?;
        self.spreadsheet.publish_sheets(candidate)
    }

    fn edit_sheet<F>(&mut self, sheet_name: &str, operation: F) -> CoreResult<()>
    where
        F: FnOnce(&mut Sheet) -> CoreResult<()>,
    {
        self.edit_sheets(|sheets| {
            let sheet = sheets
                .iter_mut()
                .find(|sheet| sheet.name == sheet_name)
                .ok_or_else(|| {
                    CoreError::InvalidFormat(format!("ODS sheet '{sheet_name}' was not found"))
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
    pub fn definitions_snapshot(&self) -> FamilyResult<crate::definitions::Snapshot> {
        self.spreadsheet.definitions_snapshot()
    }

    /// Clone-stage ordered definition CRUD and publish it as one package edit.
    ///
    /// An unchanged edit retains the original package bytes. A changed edit must produce compact
    /// `content.xml`, pass a full package reopen, and match the staged typed catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, validation, compactness check, rebuild, or readback fails.
    pub fn edit_definitions<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::definitions::Edit) -> CoreResult<()>,
    {
        let snapshot = self.definitions_snapshot()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        if commit.changed() {
            self.spreadsheet = Projection::from_bytes(commit.snapshot().as_bytes().to_vec())?;
        }
        Ok(())
    }

    /// Append a validated named range atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_range(&mut self, range: Range) -> FamilyResult<()> {
        Ok(self.spreadsheet.add_range(range)?)
    }

    /// Append a validated named expression atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_expression(&mut self, expression: Expression) -> FamilyResult<()> {
        Ok(self.spreadsheet.add_expression(expression)?)
    }

    /// Append a validated named definition atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn add_definition(&mut self, definition: Definition) -> FamilyResult<()> {
        Ok(self.spreadsheet.add_definition(definition)?)
    }

    /// Replace the complete ordered named-definition catalog atomically.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_definitions(&mut self, definitions: Vec<Definition>) -> FamilyResult<()> {
        Ok(self.spreadsheet.set_definitions(definitions)?)
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
    pub fn add_rdf_graph(
        &mut self,
        path: Option<&str>,
        triples: &[Triple],
    ) -> FamilyResult<String> {
        Ok(self.spreadsheet.add_rdf_graph(path, triples)?)
    }

    /// Capture RDF metadata as an immutable, exact-package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the package, manifest, or a declared graph is invalid.
    pub fn rdf_snapshot(&self) -> FamilyResult<crate::metadata_graphs::Snapshot> {
        self.spreadsheet.rdf_snapshot()
    }

    /// Clone-stage RDF graph and triple CRUD and publish it as one package edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or a staged package rebuild fails.
    pub fn edit_rdf<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::metadata_graphs::Edit) -> CoreResult<()>,
    {
        let snapshot = self.rdf_snapshot()?;
        let mut transaction = snapshot.edit();
        edit(&mut transaction)?;
        Ok(self
            .spreadsheet
            .apply_rdf_patch(transaction.commit().patch())?)
    }

    /// Replace one complete RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph, triples, compact XML, or rebuilt package is invalid.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> FamilyResult<()> {
        Ok(self.spreadsheet.replace_rdf_graph(path, triples)?)
    }

    /// Remove one unreferenced RDF metadata graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is missing, referenced, or package rebuilding fails.
    pub fn remove_rdf_graph(&mut self, path: &str) -> FamilyResult<()> {
        Ok(self.spreadsheet.remove_rdf_graph(path)?)
    }

    /// Append one RDF triple and return its position in the graph.
    ///
    /// # Errors
    ///
    /// Returns an error when the graph is missing or package rebuilding fails.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &Triple) -> FamilyResult<usize> {
        Ok(self.spreadsheet.add_rdf_triple(path, triple)?)
    }

    /// Replace one RDF triple while preserving its description subject.
    ///
    /// # Errors
    ///
    /// Returns an error when the triple is missing or package rebuilding fails.
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &Triple,
    ) -> FamilyResult<()> {
        Ok(self.spreadsheet.replace_rdf_triple(path, index, triple)?)
    }

    /// Remove one RDF triple.
    ///
    /// # Errors
    ///
    /// Returns an error when the triple is missing or package rebuilding fails.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> FamilyResult<()> {
        Ok(self.spreadsheet.remove_rdf_triple(path, index)?)
    }

    /// Move one RDF triple within its description.
    ///
    /// # Errors
    ///
    /// Returns an error when a position is invalid or package rebuilding fails.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> FamilyResult<()> {
        Ok(self.spreadsheet.move_rdf_triple(path, from, to)?)
    }

    /// Apply a short-lived protection transaction.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or package rebuilding fails.
    pub fn update_protection<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::protection::Edit) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.update_protection(edit)?)
    }

    /// Apply a short-lived tracked-change transaction.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or package rebuilding fails.
    pub fn update_tracked_changes<F>(&mut self, edit: F) -> FamilyResult<()>
    where
        F: FnOnce(&mut crate::tracked_changes::Edit) -> CoreResult<()>,
    {
        Ok(self.spreadsheet.update_tracked_changes(edit)?)
    }

    /// Publish the draft as an immutable [`crate::Package`].
    ///
    /// # Errors
    ///
    /// Returns an archive-validation failure.
    pub fn build(self) -> FamilyResult<crate::Package> {
        Ok(crate::Package::from_authored(
            self.spreadsheet.into_bytes(),
        )?)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Builder;

    #[test]
    fn preserves_owned_package_without_edits() {
        let bytes = Builder::new()
            .build()
            .expect("test fixture or operation should succeed")
            .to_bytes()
            .expect("test fixture or operation should succeed");
        let draft = Projection::from_bytes(bytes.clone())
            .map(Workbook::from_workbook)
            .expect("test fixture or operation should succeed");
        let built = draft.build().expect("an unchanged draft builds");
        assert_eq!(built.as_bytes(), bytes.as_slice());
        let reopened =
            Projection::from_bytes(bytes).expect("test fixture or operation should succeed");
        assert!(reopened.content_xml().contains("office:spreadsheet"));
    }
}
