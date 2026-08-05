//! Immutable database package ownership.

use litchi_core::{Metadata, Result};
use litchi_odf_common::core::family::Package;
use std::path::Path;

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.database";
const BODY_MARKER: &str = "<office:database";

/// An immutable, validated package snapshot.
pub(crate) struct Snapshot(Package);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        Package::open(path, MIMETYPE, BODY_MARKER, "ODB").map(Self)
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODB").map(Self)
    }

    pub(crate) fn content_xml(&self) -> &str {
        self.0.content_xml()
    }

    pub(crate) fn styles_xml(&self) -> Option<&str> {
        self.0.styles_xml()
    }

    pub(crate) fn metadata(&self) -> Option<&Metadata> {
        self.0.metadata()
    }

    pub(crate) fn as_bytes(&self) -> &[u8] {
        self.0.as_bytes()
    }

    pub(crate) fn files(&self) -> Result<Vec<String>> {
        self.0.files()
    }

    pub(crate) fn into_bytes(self) -> Vec<u8> {
        self.0.into_bytes()
    }
}
