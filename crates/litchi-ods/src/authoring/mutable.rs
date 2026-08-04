//! Transactional editing of an existing spreadsheet package.

use crate::{
    Spreadsheet,
    model::{NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange, Triple},
};
use litchi_core::Result;
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

    /// Return the ordered named-definition catalog of the edited snapshot.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        self.spreadsheet.named_definitions()
    }

    /// Append a validated named range atomically.
    pub fn add_named_range(&mut self, range: NamedRange) -> Result<()> {
        self.spreadsheet.add_named_range(range)
    }

    /// Append a validated named expression atomically.
    pub fn add_named_expression(&mut self, expression: NamedExpression) -> Result<()> {
        self.spreadsheet.add_named_expression(expression)
    }

    /// Append a validated named definition atomically.
    pub fn add_named_definition(&mut self, definition: NamedDefinition) -> Result<()> {
        self.spreadsheet.add_named_definition(definition)
    }

    /// Replace the complete ordered named-definition catalog atomically.
    pub fn set_named_definitions(&mut self, definitions: Vec<NamedDefinition>) -> Result<()> {
        self.spreadsheet.set_named_definitions(definitions)
    }

    /// Find a named range in the current snapshot.
    pub fn named_range(&self, name: &str, scope: &NamedDefinitionScope) -> Option<&NamedRange> {
        self.spreadsheet.named_range(name, scope)
    }

    /// Find a named expression in the current snapshot.
    pub fn named_expression(
        &self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<&NamedExpression> {
        self.spreadsheet.named_expression(name, scope)
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
