use litchi_core::Result;
use litchi_odf_common::core::{FamilyPackage, OwnedPackage};
use std::path::Path;

use crate::model::NamedDefinition;

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

    /// Read the ordered named-definition catalog from `content.xml`.
    pub fn named_definitions(&self) -> Result<Vec<NamedDefinition>> {
        crate::codec::named_expression::parse(self.content_xml())
    }

    /// Rebuild this package with a replacement `content.xml`.
    pub(crate) fn replace_content_xml(&self, content_xml: &str) -> Result<Self> {
        let bytes = litchi_odf_common::package::rebuild_package(
            self.package(),
            content_xml,
            Vec::new(),
            Vec::new(),
            Vec::new(),
            Vec::new(),
        )?;
        Self::from_bytes(bytes)
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.0.into_bytes()
    }
}
