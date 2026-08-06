//! Transactional editing of an existing spreadsheet package.

use crate::{
    Spreadsheet,
    model::names::{Definition, Expression, Range, Scope},
    worksheet::{Cell, CellView, Sheet},
};
use litchi_core::{Error, Result};
use litchi_odf_common::rdf::Triple;
use std::path::Path;

/// Mutable ODS snapshot.
///
/// Every package-level edit is validated and atomically replaces the owned
/// immutable snapshot. Failed edits leave the document unchanged.
pub struct MutableSpreadsheet {
    spreadsheet: Spreadsheet,
}

impl MutableSpreadsheet {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Spreadsheet::open(path).map(Self::from_spreadsheet)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Spreadsheet::from_bytes(bytes).map(Self::from_spreadsheet)
    }

    pub fn from_spreadsheet(spreadsheet: Spreadsheet) -> Self {
        Self { spreadsheet }
    }

    pub fn spreadsheet(&self) -> &Spreadsheet {
        &self.spreadsheet
    }

    /// Borrow the compact cross-format metadata projection.
    pub fn metadata(&self) -> &litchi_core::Metadata {
        self.spreadsheet.metadata()
    }

    /// Borrow the complete typed ODF metadata model.
    pub fn odf_metadata(&self) -> &crate::metadata::Metadata {
        self.spreadsheet.odf_metadata()
    }

    /// Replace the supported metadata projection atomically.
    pub fn set_metadata(&mut self, metadata: litchi_core::Metadata) -> Result<()> {
        self.spreadsheet.publish_metadata(metadata)
    }

    /// Apply a short-lived metadata update transactionally.
    pub fn update_metadata<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut litchi_core::Metadata) -> Result<()>,
    {
        let mut metadata = self.metadata().clone();
        update(&mut metadata)?;
        self.set_metadata(metadata)
    }

    /// Remove the physical `meta.xml` part atomically.
    pub fn clear_metadata(&mut self) -> Result<()> {
        self.spreadsheet.remove_metadata()
    }

    /// Borrow spreadsheet calculation settings, if present.
    pub fn settings(&self) -> Option<&crate::settings::Settings> {
        self.spreadsheet.settings()
    }

    /// Replace or remove calculation settings atomically.
    pub fn set_settings(&mut self, settings: Option<crate::settings::Settings>) -> Result<()> {
        if let Some(settings) = &settings {
            settings.validate()?;
        }
        self.spreadsheet.publish_settings(settings)
    }

    /// Apply a typed calculation-settings update, creating the owner when it
    /// is absent.
    pub fn update_settings<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::settings::Settings) -> Result<()>,
    {
        let mut settings = self.settings().cloned().unwrap_or_default();
        update(&mut settings)?;
        self.set_settings(Some(settings))
    }

    /// Remove the calculation-settings element from `content.xml`.
    pub fn clear_settings(&mut self) -> Result<()> {
        self.set_settings(None)
    }

    /// Capture the source-checked cell-annotation owner for this snapshot.
    pub fn annotations(&self) -> Result<crate::annotations::Snapshot> {
        self.spreadsheet.annotations()
    }

    /// Apply one failure-atomic cell-annotation transaction.
    ///
    /// The transaction resolves cells by exact sheet name and zero-based
    /// logical coordinates.  If the closure or commit fails, this facade and
    /// its package bytes remain unchanged; an empty commit does not rebuild
    /// the archive.
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
    pub fn sheets(&self) -> &[Sheet] {
        self.spreadsheet.sheets()
    }

    /// Discover embedded charts in the current immutable package snapshot.
    pub fn charts(&self) -> Result<crate::charts::Inventory<'_>> {
        self.spreadsheet.charts()
    }

    /// Clone-stage chart replacements and publish them as one package edit.
    ///
    /// A failed closure or commit leaves this mutable facade unchanged. An
    /// empty transaction does not rebuild the archive, preserving exact bytes.
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

    /// Discover the typed DataPilot catalog in the current package snapshot.
    pub fn data_pilots(&self) -> Result<crate::data_pilot::Catalog<'_>> {
        self.spreadsheet.data_pilots()
    }

    /// Clone-stage DataPilot CRUD and publish it as one package edit.
    ///
    /// Unknown markup in the owned XML is retained by no-op transactions and
    /// causes a changed transaction to fail before package bytes are rebuilt.
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
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.spreadsheet.sheet(name)
    }

    /// Look up a logical cell in the current immutable snapshot.
    pub fn cell(&self, sheet_name: &str, row: usize, column: usize) -> Option<CellView<'_>> {
        self.spreadsheet.cell(sheet_name, row, column)
    }

    /// Atomically replace all worksheets in the current package snapshot.
    pub fn set_sheets(&mut self, sheets: Vec<Sheet>) -> Result<()> {
        self.spreadsheet.publish_sheets(sheets)
    }

    /// Append one worksheet while preserving document order.
    pub fn add_sheet(&mut self, sheet: Sheet) -> Result<()> {
        self.edit_sheets(move |sheets| {
            sheets.push(sheet);
            Ok(())
        })
    }

    /// Remove one worksheet by its exact ODF name.
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
    pub fn clear_cell(&mut self, sheet_name: &str, row: usize, column: usize) -> Result<()> {
        self.edit_sheet(sheet_name, |sheet| sheet.clear_cell(row, column))
    }

    /// Set an inert formula on one logical cell.
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
    pub fn definitions(&self) -> &[Definition] {
        self.spreadsheet.definitions()
    }

    /// Append a validated named range atomically.
    pub fn add_range(&mut self, range: Range) -> Result<()> {
        self.spreadsheet.add_range(range)
    }

    /// Append a validated named expression atomically.
    pub fn add_expression(&mut self, expression: Expression) -> Result<()> {
        self.spreadsheet.add_expression(expression)
    }

    /// Append a validated named definition atomically.
    pub fn add_definition(&mut self, definition: Definition) -> Result<()> {
        self.spreadsheet.add_definition(definition)
    }

    /// Replace the complete ordered named-definition catalog atomically.
    pub fn set_definitions(&mut self, definitions: Vec<Definition>) -> Result<()> {
        self.spreadsheet.set_definitions(definitions)
    }

    /// Find a named range in the current snapshot.
    pub fn range(&self, name: &str, scope: &Scope) -> Option<&Range> {
        self.spreadsheet.range(name, scope)
    }

    /// Find a named expression in the current snapshot.
    pub fn expression(&self, name: &str, scope: &Scope) -> Option<&Expression> {
        self.spreadsheet.expression(name, scope)
    }

    pub fn add_rdf_graph(&mut self, path: Option<&str>, triples: &[Triple]) -> Result<String> {
        self.spreadsheet.add_rdf_graph(path, triples)
    }

    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> Result<()> {
        self.spreadsheet.replace_rdf_graph(path, triples)
    }

    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        self.spreadsheet.remove_rdf_graph(path)
    }

    pub fn to_bytes(self) -> Vec<u8> {
        self.spreadsheet.into_bytes()
    }

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
