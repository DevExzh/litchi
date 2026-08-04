//! Transactional editing of an existing spreadsheet package.

use crate::{Spreadsheet, model::Triple};
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
    use crate::SpreadsheetBuilder;

    #[test]
    fn preserves_owned_package_without_edits() {
        let bytes = SpreadsheetBuilder::new().build().unwrap();
        let mutable = MutableSpreadsheet::from_bytes(bytes).unwrap();
        let reopened = Spreadsheet::from_bytes(mutable.to_bytes()).unwrap();
        assert!(reopened.content_xml().contains("office:spreadsheet"));
    }
}
