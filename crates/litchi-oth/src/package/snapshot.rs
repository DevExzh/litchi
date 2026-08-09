//! Immutable web-template package ownership.

use litchi_core::{Metadata, Result};
use litchi_odf_common::core::family::Package;
use std::path::Path;

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.text-web";
// The common package shell requires a cheap marker before this family applies
// its namespace-aware structural contract below. A literal prefix is not a
// valid ODF namespace check, so it must remain deliberately non-semantic.
const PRELIMINARY_XML_MARKER: &str = "<";

/// An immutable, validated package snapshot.
pub(crate) struct Snapshot(Package);

impl Snapshot {
    pub(crate) fn open(path: impl AsRef<Path>) -> Result<Self> {
        let package = Package::open(path, MIMETYPE, PRELIMINARY_XML_MARKER, "OTH")?;
        crate::codec::validate(package.content_xml())?;
        Ok(Self(package))
    }

    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = Package::from_bytes(bytes, MIMETYPE, PRELIMINARY_XML_MARKER, "OTH")?;
        crate::codec::validate(package.content_xml())?;
        Ok(Self(package))
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

    pub(crate) fn paragraphs(&self) -> Result<Vec<crate::paragraph::Paragraph>> {
        crate::codec::paragraphs(self.content_xml())
    }
}
