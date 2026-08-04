//! Concise user-facing ODS entry points.

use litchi_core::Result;
use std::path::Path;

pub use crate::authoring::SpreadsheetBuilder;
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};

/// Immutable ODS document facade.
pub struct Spreadsheet {
    package: crate::package::SpreadsheetPackage,
}

impl Spreadsheet {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::SpreadsheetPackage::open(path).map(|package| Self { package })
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::SpreadsheetPackage::from_bytes(bytes).map(|package| Self { package })
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
