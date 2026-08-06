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
        let package = family::Package::open(path, MIMETYPE, BODY_MARKER, "ODS")?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self(package))
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = family::Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODS")?;
        crate::authoring::validate_content_xml(package.content_xml())?;
        Ok(Self(package))
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

    /// Decode the public typed worksheet graph without expanding repetition
    /// runs into one object per logical row or cell.
    pub(crate) fn sheets(&self) -> Result<Vec<crate::worksheet::Sheet>> {
        crate::worksheet::codec::parse(self.content_xml())
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

    /// Rebuild the package after an atomic semantic worksheet edit.
    pub(crate) fn replace_sheets(&self, sheets: &[crate::worksheet::Sheet]) -> Result<Self> {
        let content_xml = crate::worksheet::package::replace_tables(self.content_xml(), sheets)?;
        self.replace_content_xml(&content_xml)
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.0.into_bytes()
    }
}
