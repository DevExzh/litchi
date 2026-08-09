//! Concise family entry points.

use litchi_core::{Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Database {
    pub(crate) package: crate::package::Snapshot,
}

impl Database {
    /// Opens a database package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a database package from in-memory bytes.
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

    /// Reads a bounded, source-bound inert schema and query catalog.
    ///
    /// This never opens a connection, resolves a driver, or executes a query.
    /// Unknown source markup remains preserved in the immutable package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the content part is structurally invalid or
    /// exceeds the default semantic discovery limits.
    pub fn catalog(&self) -> Result<crate::Catalog<'_>> {
        self.catalog_with(crate::Limits::default())
    }

    /// Reads an inert schema and query catalog with caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the content part is structurally invalid or
    /// exceeds `limits`.
    pub fn catalog_with(&self, limits: crate::Limits) -> Result<crate::Catalog<'_>> {
        crate::Catalog::parse(self.content_xml(), limits)
    }

    /// Inventories opaque direct producer-extension subtrees without
    /// interpreting or activating them.
    ///
    /// # Errors
    ///
    /// Returns an error when the content XML is malformed or an extension
    /// range cannot be represented safely.
    pub fn producer_extensions(&self) -> Result<Vec<crate::ProducerExtension>> {
        crate::authoring::producer_extensions(self.content_xml())
    }

    /// Inventories signatures and manifest-declared encryption without
    /// validating, decrypting, or executing protected content.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be inspected.
    pub fn protection_status(&self) -> Result<crate::ProtectionStatus> {
        self.package.protection_status()
    }

    /// Starts a source-bound unified database-front-end transaction.
    ///
    /// Connection targets, query commands, component links, and producer
    /// extensions remain inert data and are never opened or executed.
    #[must_use]
    pub fn edit(&self) -> crate::Edit<'_> {
        crate::Edit::new(self, crate::EditPolicy::default())
    }

    /// Starts a unified transaction with explicit protected-package policy.
    #[must_use]
    pub fn edit_with_policy(&self, policy: crate::EditPolicy) -> crate::Edit<'_> {
        crate::Edit::new(self, policy)
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
    use super::{Builder, Database};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Database::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:database"));
        assert!(!document.as_bytes().is_empty());
        assert!(document.catalog().unwrap().tables().is_empty());
    }
}
