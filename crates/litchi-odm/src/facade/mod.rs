//! The immutable master-document owner and its read-only projection.

use litchi_core::{Metadata, Result as CoreResult};
use litchi_odf_common::{Error as PackageError, OpenOptions};
use std::path::Path;

pub use crate::authoring::Builder;

type FamilyResult<T> = Result<T, PackageError>;

/// The immutable, cheap-to-clone master-document owner (ADR 0034 §2).
///
/// It owns every state verb: opening, the transactions (`edit`,
/// `edit_title`, `edit_link`), patch application and publication. Reads go
/// through the [`Document`] projection returned by [`Package::document`].
#[derive(Clone)]
pub struct Package {
    package: crate::package::Snapshot,
}

/// The read-only semantic projection of a master document.
///
/// Obtained from [`Package::document`]; cloning shares the owner's
/// immutable state. It has no state verbs.
#[derive(Clone)]
pub struct Document {
    package: crate::package::Snapshot,
}

fn copy_input(bytes: &[u8]) -> Result<Vec<u8>, PackageError> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| litchi_core::Error::Allocation {
            resource: "ODM package bytes",
            source,
        })?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

impl Package {
    /// The retained physical package (metadata transactions, record 0288).
    pub(crate) fn owned_package(&self) -> &litchi_odf_common::OwnedPackage {
        self.package.owned_package()
    }

    pub(crate) fn xml_limits_value(&self) -> litchi_odf_common::XmlStreamLimits {
        self.package.xml_limits_value()
    }

    pub(crate) fn from_snapshot(package: crate::package::Snapshot) -> Self {
        Self { package }
    }

    pub(crate) fn from_shared_bytes(bytes: std::sync::Arc<Vec<u8>>) -> CoreResult<Self> {
        crate::package::Snapshot::from_shared_bytes(bytes).map(Self::from_snapshot)
    }

    pub(crate) fn from_vec(bytes: Vec<u8>) -> CoreResult<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(Self::from_snapshot)
    }

    pub(crate) fn shared_bytes(&self) -> std::sync::Arc<Vec<u8>> {
        self.package.shared_bytes()
    }

    pub(crate) fn meta_xml(&self) -> CoreResult<Option<String>> {
        self.package.meta_xml()
    }

    pub(crate) fn with_meta_xml(&self, meta_xml: &str) -> CoreResult<Self> {
        self.package
            .with_meta_xml(meta_xml)
            .map(Self::from_snapshot)
    }

    pub(crate) fn with_content_xml(&self, content_xml: &str) -> CoreResult<Self> {
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
    ) -> CoreResult<Self> {
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

    pub(crate) fn resource_bytes(&self, path: &str) -> CoreResult<Vec<u8>> {
        self.package.resource_bytes(path)
    }

    pub(crate) fn local_section_references(&self) -> &[(String, std::ops::Range<usize>)] {
        self.package.local_section_references()
    }

    pub(crate) fn href_span(&self, reference: usize) -> Option<&std::ops::Range<usize>> {
        self.package.href_span(reference)
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

    pub(crate) fn title(&self) -> Option<&str> {
        self.metadata()
            .and_then(|metadata| metadata.title.as_deref())
    }

    pub(crate) fn subdocuments(&self) -> &[crate::model::subdocument::Reference] {
        self.package.references()
    }

    pub(crate) fn section_tree(&self) -> &crate::section::Tree {
        self.package.section_tree()
    }

    pub(crate) fn structure(&self) -> &crate::structure::Structure {
        self.package.structure()
    }

    pub(crate) fn styles(&self) -> &[crate::style::Definition] {
        self.package.styles()
    }

    pub(crate) fn resources(&self) -> &crate::resource::Graph {
        self.package.resources()
    }

    pub(crate) fn security(&self) -> &crate::security::State {
        self.package.security()
    }

    /// Creates the minimal valid master document (an empty text body).
    ///
    /// Deterministic: no clock or randomness is consulted, so two calls
    /// produce identical bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if package writing fails.
    pub fn new() -> Result<Self, PackageError> {
        Builder::new().build()
    }

    /// Opens a master-document package from a file path with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Limit`] when a selected limit is exceeded and
    /// [`PackageError::Core`] for I/O, archive or family-validation failures.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, PackageError> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens a master-document package from a file path under `options`.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> Result<Self, PackageError> {
        crate::package::Snapshot::open_with_options(path, options).map(Self::from_snapshot)
    }

    /// Opens a password-encrypted master-document package from a file path.
    ///
    /// The resulting package remains readable, but any changed transaction is
    /// refused because this crate does not publish or retain credentials.
    ///
    /// # Errors
    ///
    /// Returns an error for filesystem, password, encryption, limit or
    /// package validation failures.
    pub fn open_with_password(
        path: impl AsRef<Path>,
        password: impl Into<String>,
    ) -> Result<Self, PackageError> {
        Self::open_with_password_and_options(path, password, OpenOptions::new())
    }

    /// Opens a password-encrypted master-document package from a file path
    /// under `options`.
    ///
    /// # Errors
    ///
    /// As [`Self::open_with_password`].
    pub fn open_with_password_and_options(
        path: impl AsRef<Path>,
        password: impl Into<String>,
        options: OpenOptions,
    ) -> Result<Self, PackageError> {
        crate::package::Snapshot::open_with_password_and_options(path, password, options)
            .map(Self::from_snapshot)
    }

    /// Opens a master-document package from in-memory bytes (copied once).
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, PackageError> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens a master-document package from in-memory bytes under `options`.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> Result<Self, PackageError> {
        crate::package::Snapshot::from_bytes_with_options(copy_input(bytes)?, options)
            .map(Self::from_snapshot)
    }

    /// Opens a password-encrypted master-document package from bytes.
    ///
    /// The resulting package remains readable, but any changed transaction is
    /// refused because this crate does not publish or retain credentials.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid package, password, encrypted entry or
    /// exceeded limit.
    pub fn from_bytes_with_password(
        bytes: &[u8],
        password: impl Into<String>,
    ) -> Result<Self, PackageError> {
        Self::from_bytes_with_password_and_options(bytes, password, OpenOptions::new())
    }

    /// Opens a password-encrypted master-document package from bytes under
    /// `options`.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_password`].
    pub fn from_bytes_with_password_and_options(
        bytes: &[u8],
        password: impl Into<String>,
        options: OpenOptions,
    ) -> Result<Self, PackageError> {
        crate::package::Snapshot::from_bytes_with_password_and_options(
            copy_input(bytes)?,
            password,
            options,
        )
        .map(Self::from_snapshot)
    }

    /// Reads a master-document package from a stream.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader(reader: impl std::io::Read) -> Result<Self, PackageError> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Reads a master-document package from a stream under `options`.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader_with_options(
        reader: impl std::io::Read,
        options: OpenOptions,
    ) -> Result<Self, PackageError> {
        crate::package::Snapshot::from_reader_with_options(reader, options).map(Self::from_snapshot)
    }

    /// Adopts the package smart detection prepared, reusing its archive
    /// index (no second ZIP parse).
    ///
    /// # Errors
    ///
    /// Returns [`litchi_core::Error::UnexpectedFormat`] (as
    /// [`PackageError::Core`]) when the prepared package is another family,
    /// and otherwise as [`Self::open`].
    pub fn from_prepared_package(
        prepared: litchi_odf_common::PreparedPackage,
    ) -> Result<Self, PackageError> {
        Self::from_prepared_package_with_options(prepared, OpenOptions::new())
    }

    /// Adopts a prepared package under the semantic limits of `options`
    /// (its source limits were applied when it was indexed).
    ///
    /// # Errors
    ///
    /// As [`Self::from_prepared_package`].
    pub fn from_prepared_package_with_options(
        prepared: litchi_odf_common::PreparedPackage,
        options: OpenOptions,
    ) -> Result<Self, PackageError> {
        let detected = prepared.format();
        if detected != litchi_core::detection::FileFormat::Odm {
            return Err(litchi_core::Error::UnexpectedFormat { detected }.into());
        }
        crate::package::Snapshot::from_owned_package_with_options(prepared.into_package(), options)
            .map(Self::from_snapshot)
    }

    /// Returns the read-only semantic projection.
    #[must_use]
    pub fn document(&self) -> Document {
        Document {
            package: self.package.clone(),
        }
    }

    /// Starts a source-checked transaction for the projected document title.
    ///
    /// # Errors
    ///
    /// Returns an error when the source metadata part cannot be read as UTF-8.
    pub fn edit_title(&self) -> FamilyResult<crate::title::Edit<'_>> {
        crate::title::Edit::new(self).map_err(PackageError::from)
    }

    /// Applies a title patch only when this is its exact source package.
    ///
    /// # Errors
    ///
    /// Returns an error when this package does not match the patch's source.
    pub fn apply_title(
        &self,
        patch: &crate::title::Patch,
    ) -> Result<crate::title::Commit, PackageError> {
        patch.apply(self)
    }

    /// Starts a source-checked transaction for one linked section target.
    ///
    /// The selector is the checked zero-based position returned by
    /// [`Document::subdocuments`]. It is resolved immediately against this
    /// immutable package.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds or the exact source
    /// attribute cannot be addressed losslessly.
    pub fn edit_link<'source, 'selector>(
        &'source self,
        selector: impl Into<crate::link::Selector<'selector>>,
    ) -> FamilyResult<crate::link::Edit<'source>> {
        crate::link::Edit::new(self, selector.into()).map_err(PackageError::from)
    }

    /// Applies a linked-section patch only when this is its exact source.
    ///
    /// # Errors
    ///
    /// Returns an error when this package does not match the patch's source.
    pub fn apply_link(
        &self,
        patch: &crate::link::Patch,
    ) -> Result<crate::link::Commit, PackageError> {
        patch.apply(self)
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
    pub fn apply(
        &self,
        patch: &crate::transaction::Patch,
    ) -> Result<crate::transaction::Commit, PackageError> {
        if !patch.is_applicable_to(self) {
            return Err(PackageError::patch_conflict(
                "odm.apply",
                "the package is not the patch source",
            ));
        }
        let package = patch.apply(self)?;
        Ok(crate::transaction::Commit::new(
            package,
            patch.clone(),
            litchi_core::edit::Diagnostics::new(!patch.is_noop()),
        ))
    }

    /// Creates explicit bounded undo/redo history at this package.
    #[must_use]
    pub fn history(&self, limits: litchi_core::HistoryLimits) -> crate::transaction::History {
        crate::transaction::History::new(self.clone(), limits)
    }

    /// Atomically replaces `path` with this package (`Durability::Full`):
    /// a sibling temporary file renamed over the destination.
    ///
    /// # Errors
    ///
    /// Returns an I/O or publication failure; the destination keeps its
    /// previous bytes on failure.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<(), PackageError> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically replaces `path` with this package under explicit
    /// durability.
    ///
    /// # Errors
    ///
    /// As [`Self::save`].
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<(), PackageError> {
        litchi_odf_common::publish::save_bytes(path.as_ref(), self.as_bytes(), durability)
    }

    /// Copies the exact package bytes.
    ///
    /// # Errors
    ///
    /// Returns an allocation failure.
    pub fn to_bytes(&self) -> Result<Vec<u8>, PackageError> {
        copy_input(self.as_bytes())
    }

    /// Writes the exact package bytes to a stream.
    ///
    /// # Errors
    ///
    /// Returns the stream's I/O failure.
    pub fn write_to(&self, mut writer: impl std::io::Write) -> Result<(), PackageError> {
        writer.write_all(self.as_bytes())?;
        Ok(())
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }
}

impl std::fmt::Debug for Package {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Package")
            .field("bytes", &self.as_bytes().len())
            .finish_non_exhaustive()
    }
}

impl Document {
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

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> FamilyResult<Vec<String>> {
        self.package.files().map_err(PackageError::from)
    }
}

impl std::fmt::Debug for Document {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Document")
            .field("subdocuments", &self.subdocuments().len())
            .finish_non_exhaustive()
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]
mod tests {
    use super::{Builder, Package};

    #[test]
    fn builder_opens_as_validated_package() {
        let package = Builder::new().build().unwrap();
        let reopened = Package::from_bytes(package.as_bytes()).unwrap();
        assert!(reopened.document().content_xml().contains("<office:text"));
        assert!(!reopened.as_bytes().is_empty());
    }
}
