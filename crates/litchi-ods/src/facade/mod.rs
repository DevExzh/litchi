//! Concise user-facing ODS entry points.

use litchi_core::Result;
use std::path::Path;

pub use crate::authoring::{MutableSpreadsheet, SpreadsheetBuilder};
pub use crate::model::{
    NamedDefinition, NamedDefinitionScope, NamedExpression, NamedRange, NamedRangeUsage,
};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};

/// Immutable ODS document facade.
pub struct Spreadsheet {
    package: crate::package::SpreadsheetPackage,
    named_definitions: Vec<NamedDefinition>,
}

impl Spreadsheet {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = crate::package::SpreadsheetPackage::open(path)?;
        Self::from_package(package)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Self::from_package(package)
    }

    fn from_package(package: crate::package::SpreadsheetPackage) -> Result<Self> {
        let named_definitions = package.named_definitions()?;
        Ok(Self {
            package,
            named_definitions,
        })
    }

    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Return all global and sheet-local named definitions in document order.
    pub fn named_definitions(&self) -> &[NamedDefinition] {
        &self.named_definitions
    }

    /// Return named ranges in their document order.
    pub fn named_ranges(&self) -> impl Iterator<Item = &NamedRange> {
        self.named_definitions
            .iter()
            .filter_map(|definition| match definition {
                NamedDefinition::Range(range) => Some(range),
                NamedDefinition::Expression(_) => None,
            })
    }

    /// Return named expressions in their document order.
    pub fn named_expressions(&self) -> impl Iterator<Item = &NamedExpression> {
        self.named_definitions
            .iter()
            .filter_map(|definition| match definition {
                NamedDefinition::Range(_) => None,
                NamedDefinition::Expression(expression) => Some(expression),
            })
    }

    /// Find a named range by its exact name and visibility scope.
    pub fn named_range(&self, name: &str, scope: &NamedDefinitionScope) -> Option<&NamedRange> {
        self.named_ranges()
            .find(|range| range.name == name && &range.scope == scope)
    }

    /// Find a named expression by its exact name and visibility scope.
    pub fn named_expression(
        &self,
        name: &str,
        scope: &NamedDefinitionScope,
    ) -> Option<&NamedExpression> {
        self.named_expressions()
            .find(|expression| expression.name == name && &expression.scope == scope)
    }

    /// Atomically append a validated named range.
    pub fn add_named_range(&mut self, range: NamedRange) -> Result<()> {
        self.add_named_definition(range.into())
    }

    /// Atomically append a validated named expression.
    pub fn add_named_expression(&mut self, expression: NamedExpression) -> Result<()> {
        self.add_named_definition(expression.into())
    }

    /// Atomically append a validated named definition while preserving catalog order.
    pub fn add_named_definition(&mut self, definition: NamedDefinition) -> Result<()> {
        let mut candidate = self.named_definitions.clone();
        candidate.push(definition);
        self.set_named_definitions(candidate)
    }

    /// Atomically replace the complete ordered named-definition catalog.
    pub fn set_named_definitions(&mut self, definitions: Vec<NamedDefinition>) -> Result<()> {
        let updated =
            crate::codec::named_expression::replace(self.package.content_xml(), &definitions)?;
        let package = self.package.replace_content_xml(&updated)?;
        self.package = package;
        self.named_definitions = definitions;
        Ok(())
    }

    /// Read all inert RDF metadata graphs in package order.
    pub fn rdf_graphs(&self) -> Result<Vec<Graph>> {
        litchi_odf_common::rdf::graphs(self.package.package())
    }

    /// Add a graph and atomically replace this snapshot with the rebuilt package.
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[Triple],
    ) -> Result<String> {
        let (bytes, path) =
            litchi_odf_common::rdf::add_graph(self.package.package(), preferred_path, triples)?;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(path)
    }

    /// Replace one complete RDF graph and atomically publish the result.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> Result<()> {
        let bytes = litchi_odf_common::rdf::replace_graph(self.package.package(), path, triples)?;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one RDF graph after validating that no remaining graph references it.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = litchi_odf_common::rdf::remove_graph(self.package.package(), path)?;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(())
    }

    /// Append one triple to an existing graph and return its committed index.
    pub fn add_rdf_triple(&mut self, path: &str, triple: &Triple) -> Result<usize> {
        let index = self
            .rdf_graphs()?
            .into_iter()
            .find(|graph| graph.path == path)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!("RDF graph '{path}' was not found"))
            })?
            .triples
            .len();
        let bytes = litchi_odf_common::rdf::add_triple(self.package.package(), path, triple)?.0;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(index)
    }

    /// Replace one triple while preserving its description subject.
    pub fn replace_rdf_triple(&mut self, path: &str, index: usize, triple: &Triple) -> Result<()> {
        let bytes =
            litchi_odf_common::rdf::replace_triple(self.package.package(), path, index, triple)?;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one triple from a graph.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = litchi_odf_common::rdf::remove_triple(self.package.package(), path, index)?;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(())
    }

    /// Move one triple within its RDF description.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = litchi_odf_common::rdf::move_triple(self.package.package(), path, from, to)?;
        self.package = crate::package::SpreadsheetPackage::from_bytes(bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builder_round_trips_through_facade() {
        let bytes = SpreadsheetBuilder::new().build().unwrap();
        let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
        assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    }
}
