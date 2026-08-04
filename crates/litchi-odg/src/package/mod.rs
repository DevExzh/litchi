//! Validated package ownership for this family.

use litchi_core::{Metadata, Result};
use litchi_odf_common::core::FamilyPackage;
use std::path::Path;

pub const MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics";
const BODY_MARKER: &str = "<office:drawing";

/// An immutable, validated package snapshot.
pub struct Package(FamilyPackage);

impl Package {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        FamilyPackage::open(path, MIMETYPE, BODY_MARKER, "ODG").map(Self)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        FamilyPackage::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODG").map(Self)
    }

    pub fn content_xml(&self) -> &str {
        self.0.content_xml()
    }

    pub fn styles_xml(&self) -> Option<&str> {
        self.0.styles_xml()
    }

    pub fn metadata(&self) -> Option<&Metadata> {
        self.0.metadata()
    }

    pub fn files(&self) -> Result<Vec<String>> {
        self.0.files()
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.0.into_bytes()
    }
}
