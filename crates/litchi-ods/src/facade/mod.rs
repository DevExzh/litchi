//! Concise user-facing ODS entry points.

use litchi_core::Result;
use std::path::Path;

pub use crate::authoring::SpreadsheetBuilder;

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
