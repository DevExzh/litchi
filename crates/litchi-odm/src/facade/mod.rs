//! Concise family entry points.

use litchi_core::{Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
#[derive(Clone)]
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

    pub(crate) fn with_content_xml(&self, content_xml: &str) -> Result<Self> {
        self.package
            .with_content_xml(content_xml)
            .map(Self::from_snapshot)
    }

    pub(crate) fn with_transaction_parts(
        &self,
        content_xml: &str,
        styles_xml: Option<&str>,
        meta_xml: Option<&str>,
        removed_resources: &[String],
        resource_writes: &[crate::package::ResourceWrite],
    ) -> Result<Self> {
        self.package
            .with_transaction_parts(
                content_xml,
                styles_xml,
                meta_xml,
                removed_resources,
                resource_writes,
            )
            .map(Self::from_snapshot)
    }

    pub(crate) fn resource_bytes(&self, path: &str) -> Result<Vec<u8>> {
        self.package.resource_bytes(path)
    }

    pub(crate) fn local_section_references(&self) -> &[(String, std::ops::Range<usize>)] {
        self.package.local_section_references()
    }

    pub(crate) fn href_span(&self, reference: usize) -> Option<&std::ops::Range<usize>> {
        self.package.href_span(reference)
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

    /// Opens a password-encrypted master-document package from bytes.
    ///
    /// The resulting snapshot remains readable, but any changed transaction
    /// is refused because this crate does not publish or retain credentials.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid package, password, or encrypted entry.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        crate::package::Snapshot::from_bytes_with_password(bytes, password).map(Self::from_snapshot)
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

    /// Starts a source-checked transaction for one linked section target.
    ///
    /// The selector is the checked zero-based position returned by
    /// [`Self::subdocuments`]. It is resolved immediately against this
    /// immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds or the exact source
    /// attribute cannot be addressed losslessly.
    pub fn edit_link<'source, 'selector>(
        &'source self,
        selector: impl Into<crate::link::Selector<'selector>>,
    ) -> Result<crate::link::Edit<'source>> {
        crate::link::Edit::new(self, selector.into())
    }

    /// Applies a linked-section patch only when this is its exact source.
    ///
    /// # Errors
    ///
    /// Returns an error when this package does not match the patch's source.
    pub fn apply_link_patch(&self, patch: &crate::link::Patch) -> Result<Self> {
        patch.apply(self)
    }

    /// Returns subdocument references in document order.
    ///
    /// Targets are classified but never opened, resolved, fetched, or
    /// executed. Unified transactions can retarget existing references or add
    /// and transfer linked sections with explicit package dependency closure.
    #[must_use]
    pub fn subdocuments(&self) -> &[crate::model::subdocument::Reference] {
        self.package.references()
    }

    /// Returns the bounded section tree projected from `content.xml`.
    #[must_use]
    pub fn section_tree(&self) -> &crate::section::Tree {
        self.package.section_tree()
    }

    /// Returns common direct master-body structures in authored order.
    #[must_use]
    pub fn structure(&self) -> &crate::structure::Structure {
        self.package.structure()
    }

    /// Returns named style definitions from `content.xml` and `styles.xml`.
    #[must_use]
    pub fn styles(&self) -> &[crate::style::Definition] {
        self.package.styles()
    }

    /// Returns the inert package-resource graph.
    #[must_use]
    pub fn resources(&self) -> &crate::resource::Graph {
        self.package.resources()
    }

    /// Returns signatures, encryption, and inert active-content inventory.
    #[must_use]
    pub fn security(&self) -> &crate::security::State {
        self.package.security()
    }

    /// Starts one atomic metadata, section, style, and resource transaction.
    #[must_use]
    pub fn edit(&self) -> crate::transaction::Edit<'_> {
        crate::transaction::Edit::new(self)
    }

    /// Starts one transaction under an explicit external-link/resource policy.
    #[must_use]
    pub fn edit_with_policy(
        &self,
        policy: crate::transaction::SecurityPolicy,
    ) -> crate::transaction::Edit<'_> {
        crate::transaction::Edit::with_policy(self, policy)
    }

    /// Applies a unified patch only to its exact source package.
    ///
    /// # Errors
    ///
    /// Returns an error when this package is not the patch source.
    pub fn apply_patch(&self, patch: &crate::transaction::Patch) -> Result<Self> {
        patch.apply(self)
    }

    /// Creates explicit bounded undo/redo history at this snapshot.
    #[must_use]
    pub fn history(&self, limits: litchi_core::HistoryLimits) -> crate::transaction::History {
        crate::transaction::History::new(self.clone(), limits)
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
