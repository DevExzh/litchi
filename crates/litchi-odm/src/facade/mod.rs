//! Concise family entry points.

use litchi_core::{Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
pub struct Master {
    package: crate::package::Snapshot,
}

impl Master {
    pub(crate) fn from_snapshot(package: crate::package::Snapshot) -> Self {
        Self { package }
    }

    pub(crate) fn from_shared_bytes(bytes: std::sync::Arc<Vec<u8>>) -> Result<Self> {
        crate::package::Snapshot::from_shared_bytes(bytes).map(Self::from_snapshot)
    }

    pub(crate) fn shared_bytes(&self) -> std::sync::Arc<Vec<u8>> {
        self.package.shared_bytes()
    }

    pub(crate) fn meta_xml(&self) -> Result<Option<String>> {
        self.package.meta_xml()
    }

    pub(crate) fn with_meta_xml(&self, meta_xml: &str) -> Result<Self> {
        self.package
            .with_meta_xml(meta_xml)
            .map(Self::from_snapshot)
    }

    /// Opens a master-document package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a master-document package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

    /// Returns the `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns the `styles.xml` document, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns the document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns the projected document title, if the package has metadata.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.metadata()
            .and_then(|metadata| metadata.title.as_deref())
    }

    /// Starts a source-checked transaction for the projected document title.
    ///
    /// # Errors
    ///
    /// Returns an error when the source metadata part cannot be read as UTF-8.
    pub fn edit_title(&self) -> Result<crate::title::Edit<'_>> {
        crate::title::Edit::new(self)
    }

    /// Applies a title patch only when this is its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when this package does not match the patch's source.
    pub fn apply_title_patch(&self, patch: &crate::title::Patch) -> Result<Self> {
        patch.apply(self)
    }

    /// Returns subdocument references in document order.
    ///
    /// Targets are classified but never opened, resolved, fetched, or
    /// executed. The snapshot is intentionally read-only because the current
    /// archive adapter cannot rewrite this XML dependency closure without
    /// normalizing unknown package records.
    #[must_use]
    pub fn subdocuments(&self) -> &[crate::model::subdocument::Reference] {
        self.package.references()
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]
mod tests {
    use super::{Builder, Master};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Master::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:text"));
        assert!(!document.as_bytes().is_empty());
    }
}
