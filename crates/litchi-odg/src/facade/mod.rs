//! Concise family entry points.

use litchi_core::{CompositionLimits, Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;
pub use crate::package::{
    Change, Commit, ControlReferenceChange, DurablePatch, GeometryChange, History, HistoryLimits,
    JoinedEdits, LayerChange, Lineage, MergePlan, NameChange, PageNameChange, PageStyleChange,
    Patch, PathChange, PreparedEdit, ResourceChange, SecurityStatus, ShapeTransfer, Snapshot,
    StructureChange, StyleChange, TextChange, Transaction, TransferResource,
};

/// Immutable source-owning drawing facade.
pub struct Drawing {
    package: Snapshot,
}

impl Drawing {
    /// Opens a drawing package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package cannot be read or is not a structurally valid ODG.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens an OTG drawing template from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is unreadable or is not a structurally valid OTG.
    pub fn open_template(path: impl AsRef<Path>) -> Result<Self> {
        Snapshot::open_template(path).map(|package| Self { package })
    }

    /// Opens a drawing package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a structurally valid ODG.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

    /// Opens password-protected ODG bytes for inert inspection.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid encryption metadata, a bad password, or invalid ODG content.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Snapshot::from_bytes_with_password(bytes, password).map(|package| Self { package })
    }

    /// Opens a password-protected ODG file for inert inspection.
    ///
    /// # Errors
    ///
    /// Returns an error for filesystem, password, encryption, or ODG validation failures.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        Snapshot::open_with_password(path, password).map(|package| Self { package })
    }

    /// Opens an OTG drawing template from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a structurally valid OTG.
    pub fn from_template_bytes(bytes: Vec<u8>) -> Result<Self> {
        Snapshot::from_template_bytes(bytes).map(|package| Self { package })
    }

    /// Returns the exact `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns exact `styles.xml`, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns bounded semantic pages in source order.
    #[must_use]
    pub fn pages(&self) -> &[crate::page::Page] {
        self.package.pages()
    }

    /// Selects one page by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact-name selector is ambiguous.
    pub fn page<'selector>(
        &self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Option<&crate::page::Page>> {
        self.package.page(selector)
    }

    /// Returns global `styles.xml` drawing layers in source order.
    #[must_use]
    pub fn layers(&self) -> &[crate::layer::Layer] {
        self.package.layers()
    }

    /// Returns original package bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Whether this drawing is an OTG template.
    #[must_use]
    pub fn is_template(&self) -> bool {
        self.package.is_template()
    }

    /// Inert signature/encryption state and rewrite policy.
    #[must_use]
    pub fn security(&self) -> SecurityStatus {
        self.package.security()
    }

    /// Lists safe package member names.
    ///
    /// # Errors
    ///
    /// Returns an error when package member validation fails.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Returns package-local image resources referenced by drawing XML.
    #[must_use]
    pub fn resources(&self) -> &[crate::resource::Resource] {
        self.package.resources()
    }

    /// Returns inert form elements carrying `form:id` in source order.
    #[must_use]
    pub fn form_controls(&self) -> &[crate::FormControl] {
        self.package.form_controls()
    }

    /// Reads one package-local resource without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable member.
    pub fn resource_bytes(&self, resource: usize) -> Result<Option<Vec<u8>>> {
        self.package.resource_bytes(resource)
    }

    /// Starts a source-bound package semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.package.edit()
    }

    /// Prepares a provenance-bound shape or complete group subtree for cross-drawing transfer.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, noncompact source, or unreadable dependencies.
    pub fn prepare_shape_transfer(&self, page: usize, shape: usize) -> Result<ShapeTransfer> {
        self.package.prepare_shape_transfer(page, shape)
    }

    /// Starts deterministic exact-lineage sub-edit composition.
    #[must_use]
    pub fn joined_edits(&self, limits: CompositionLimits) -> JoinedEdits {
        self.package.joined_edits(limits)
    }

    /// Applies joined disjoint work atomically and returns a new drawing facade.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, conflicts, policy refusal, or failed full reopen.
    pub fn apply_joined(&self, joined: JoinedEdits) -> Result<Self> {
        self.package
            .apply_joined(joined)
            .map(|package| Self { package })
    }

    /// Starts explicit bounded undo/redo history at this drawing snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> History {
        self.package.history(limits)
    }

    /// Returns the source-owning package snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.package
    }

    /// Consumes the facade and returns source bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}
