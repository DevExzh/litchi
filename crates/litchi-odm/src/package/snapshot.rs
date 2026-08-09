//! Immutable master-document package ownership.

use litchi_core::{Metadata, Result};
use litchi_odf_common::core::family::Package;
use std::path::Path;

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-master";
const BODY_MARKER: &str = "<office:text";

/// An immutable, validated package snapshot.
pub(crate) struct Snapshot {
    package: Package,
    semantics: crate::codec::Semantics,
}

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path, MIMETYPE, BODY_MARKER, "ODM")?;
        let semantics = crate::codec::parse(package.content_xml())?;
        Ok(Self { package, semantics })
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODM")?;
        let semantics = crate::codec::parse(package.content_xml())?;
        Ok(Self { package, semantics })
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    pub(crate) fn references(&self) -> &[crate::model::subdocument::Reference] {
        self.semantics.references()
    }
}
