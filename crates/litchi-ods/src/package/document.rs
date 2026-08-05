use litchi_core::Result;
use litchi_odf_common::core::{OwnedPackage, family};
use std::path::Path;

use crate::model::names::Definition;

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";
const BODY_MARKER: &str = "<office:spreadsheet";

/// Validated ownership boundary for an ODS package.
pub struct Package(family::Package);

impl Package {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        family::Package::open(path, MIMETYPE, BODY_MARKER, "ODS").map(Self)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        family::Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS").map(Self)
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
    pub fn definitions(&self) -> Result<Vec<Definition>> {
        crate::codec::names::parse(self.content_xml())
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
