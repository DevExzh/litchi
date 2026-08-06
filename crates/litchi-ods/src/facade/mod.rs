//! Concise user-facing ODS entry points.

use litchi_core::Result;
use std::path::Path;

pub use crate::authoring::{Builder, MutableSpreadsheet};
use crate::model::names::{Definition, Expression, Range, Scope};
pub use litchi_odf_common::rdf::{Graph, Object, Subject, Triple};

/// Immutable ODS document facade.
pub struct Spreadsheet {
    package: crate::package::Package,
    definitions: Vec<Definition>,
    sheets: Vec<crate::worksheet::Sheet>,
}

impl Spreadsheet {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = crate::package::Package::open(path)?;
        Self::from_package(package)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = crate::package::Package::from_bytes(bytes)?;
        Self::from_package(package)
    }

    fn from_package(package: crate::package::Package) -> Result<Self> {
        let definitions = package.definitions()?;
        let sheets = package.sheets()?;
        Ok(Self {
            package,
            definitions,
            sheets,
        })
    }

    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Return the typed worksheet graph in document order.
    pub fn sheets(&self) -> &[crate::worksheet::Sheet] {
        &self.sheets
    }

    /// Find a worksheet by its exact ODF name.
    pub fn sheet(&self, name: &str) -> Option<&crate::worksheet::Sheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }

    /// Look up a logical cell while retaining the distinction between a
    /// missing coordinate and a physical repeated cell run.
    pub fn cell(
        &self,
        sheet_name: &str,
        row: usize,
        column: usize,
    ) -> Option<crate::worksheet::CellView<'_>> {
        self.sheet(sheet_name)
            .map(|sheet| sheet.cell_view(row, column))
    }

    /// Discover package, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::media::Image>> {
        let package = self.package.package().package()?;
        crate::media::scan_package(
            self.package.content_xml(),
            self.package.styles_xml(),
            &package,
        )
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::embedded::Object>> {
        let package = self.package.package().package()?;
        crate::embedded::scan_package(
            self.package.content_xml(),
            self.package.styles_xml(),
            &package,
        )
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked and missing images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::media::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            crate::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::media::Source::PackagePart { path, .. } => {
                self.package.package().get_file(path).map(Some)
            },
            _ => Ok(None),
        }
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Return all global and sheet-local named definitions in document order.
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    /// Return named ranges in their document order.
    pub fn ranges(&self) -> impl Iterator<Item = &Range> {
        self.definitions
            .iter()
            .filter_map(|definition| match definition {
                Definition::Range(range) => Some(range),
                Definition::Expression(_) => None,
            })
    }

    /// Return named expressions in their document order.
    pub fn expressions(&self) -> impl Iterator<Item = &Expression> {
        self.definitions
            .iter()
            .filter_map(|definition| match definition {
                Definition::Range(_) => None,
                Definition::Expression(expression) => Some(expression),
            })
    }

    /// Find a named range by its exact name and visibility scope.
    pub fn range(&self, name: &str, scope: &Scope) -> Option<&Range> {
        self.ranges()
            .find(|range| range.name == name && &range.scope == scope)
    }

    /// Find a named expression by its exact name and visibility scope.
    pub fn expression(&self, name: &str, scope: &Scope) -> Option<&Expression> {
        self.expressions()
            .find(|expression| expression.name == name && &expression.scope == scope)
    }

    /// Atomically append a validated named range.
    pub fn add_range(&mut self, range: Range) -> Result<()> {
        self.add_definition(range.into())
    }

    /// Atomically append a validated named expression.
    pub fn add_expression(&mut self, expression: Expression) -> Result<()> {
        self.add_definition(expression.into())
    }

    /// Atomically append a validated named definition while preserving catalog order.
    pub fn add_definition(&mut self, definition: Definition) -> Result<()> {
        let mut candidate = self.definitions.clone();
        candidate.push(definition);
        self.set_definitions(candidate)
    }

    /// Atomically replace the complete ordered named-definition catalog.
    pub fn set_definitions(&mut self, definitions: Vec<Definition>) -> Result<()> {
        let updated = crate::codec::names::replace(self.package.content_xml(), &definitions)?;
        let package = self.package.replace_content_xml(&updated)?;
        self.package = package;
        self.definitions = definitions;
        Ok(())
    }

    /// Publish a validated worksheet snapshot as one package transaction.
    pub(crate) fn publish_sheets(&mut self, sheets: Vec<crate::worksheet::Sheet>) -> Result<()> {
        let package = self.package.replace_sheets(&sheets)?;
        self.package = package;
        self.sheets = sheets;
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
        self.package = crate::package::Package::from_bytes(bytes)?;
        Ok(path)
    }

    /// Replace one complete RDF graph and atomically publish the result.
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[Triple]) -> Result<()> {
        let bytes = litchi_odf_common::rdf::replace_graph(self.package.package(), path, triples)?;
        self.package = crate::package::Package::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one RDF graph after validating that no remaining graph references it.
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = litchi_odf_common::rdf::remove_graph(self.package.package(), path)?;
        self.package = crate::package::Package::from_bytes(bytes)?;
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
        self.package = crate::package::Package::from_bytes(bytes)?;
        Ok(index)
    }

    /// Replace one triple while preserving its description subject.
    pub fn replace_rdf_triple(&mut self, path: &str, index: usize, triple: &Triple) -> Result<()> {
        let bytes =
            litchi_odf_common::rdf::replace_triple(self.package.package(), path, index, triple)?;
        self.package = crate::package::Package::from_bytes(bytes)?;
        Ok(())
    }

    /// Remove one triple from a graph.
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = litchi_odf_common::rdf::remove_triple(self.package.package(), path, index)?;
        self.package = crate::package::Package::from_bytes(bytes)?;
        Ok(())
    }

    /// Move one triple within its RDF description.
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = litchi_odf_common::rdf::move_triple(self.package.package(), path, from, to)?;
        self.package = crate::package::Package::from_bytes(bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builder_round_trips_through_facade() {
        let bytes = Builder::new().build().unwrap();
        let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();
        assert!(spreadsheet.content_xml().contains("office:spreadsheet"));
    }

    #[test]
    fn shared_resource_inventory_is_available_from_spreadsheet() {
        let bytes = Builder::new()
            .content_xml(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3">
  <office:body><office:spreadsheet>
    <draw:frame draw:name="Photo">
      <draw:image><office:binary-data>AQID</office:binary-data></draw:image>
    </draw:frame>
    <draw:object xlink:href="https://example.invalid/object" xlink:type="simple"/>
  </office:spreadsheet></office:body>
</office:document-content>"#,
            )
            .build()
            .unwrap();
        let spreadsheet = Spreadsheet::from_bytes(bytes).unwrap();

        let images = spreadsheet.images().unwrap();
        assert_eq!(images.len(), 1);
        assert_eq!(images[0].inline_bytes(), Some(&[1, 2, 3][..]));
        assert_eq!(
            spreadsheet.image_bytes(&images[0]).unwrap(),
            Some(vec![1, 2, 3])
        );

        let objects = spreadsheet.embedded_objects().unwrap();
        assert_eq!(objects.len(), 1);
        assert!(matches!(
            objects[0].source,
            crate::embedded::Source::Linked { ref href }
                if href == "https://example.invalid/object"
        ));
    }
}
