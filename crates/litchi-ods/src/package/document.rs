use litchi_core::Result;
use litchi_odf_common::core::{FamilyPackage, OwnedPackage};
use std::path::Path;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const BODY_MARKER: &str = "<office:spreadsheet";

/// Validated ownership boundary for an ODS package.
pub struct SpreadsheetPackage(FamilyPackage);

impl SpreadsheetPackage {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        FamilyPackage::open(path, MIMETYPE, BODY_MARKER, "ODS").map(Self)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        FamilyPackage::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS").map(Self)
    }

    pub fn content_xml(&self) -> &str {
        self.0.content_xml()
    }

    pub fn styles_xml(&self) -> Option<&str> {
        self.0.styles_xml()
    }

    pub fn package(&self) -> &OwnedPackage {
        self.0.package()
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.0.into_bytes()
    }
}
