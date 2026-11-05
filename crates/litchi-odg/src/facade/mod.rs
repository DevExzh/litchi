//! Concise family entry points.

use crate::package::{
    ActiveContentStatus, Package, SecurityCapabilities, SecurityStatus, ShapeTransfer,
};
use litchi_core::Metadata;

pub use crate::authoring::Builder;

type FamilyResult<T> = Result<T, litchi_odf_common::Error>;

/// Source-checked ODG package transactions, patches and history.
///
/// The owner is [`crate::Package`]; [`edit::Edit`](crate::edit::Edit) is started with
/// [`crate::Package::edit`] and published with [`crate::edit::Edit::commit`].
pub mod edit {
    pub use crate::package::{
        ActiveContentStatus, ActiveContentWritePolicy, Change, Commit, ControlReferenceChange,
        Diagnostics, DurablePatch, Edit, GeometryChange, History, HistoryLimits, JoinError,
        JoinedEdits, LayerChange, Lineage, MergePlan, NameChange, PageNameChange, PageStyleChange,
        PageTransitionChange, Patch, PathChange, ResourceChange, SecurityCapabilities,
        SecurityLifecycleDisposition, SecurityLifecycleOperation, SecurityLifecycleRefusal,
        SecurityStatus, SecurityWritePolicy, ShapeTransfer, StructureChange, StyleChange, SubEdit,
        TextChange, TransferControl, TransferResource, TransferStyle, TransferStyleResource,
    };
    pub use litchi_odf_common::{Error, LimitKind};
}

/// The read-only drawing projection of one [`Package`].
///
/// A cheap clone (one reference-count increment). It owns no state verbs:
/// ingress, editing, patch application and output live on [`Package`].
#[derive(Clone)]
pub struct Drawing {
    package: Package,
}

impl std::fmt::Debug for Drawing {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Drawing")
            .field("pages", &self.package.pages().len())
            .finish_non_exhaustive()
    }
}

impl Drawing {
    pub(crate) fn from_package(package: Package) -> Self {
        Self { package }
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
    ) -> FamilyResult<Option<&crate::page::Page>> {
        self.package
            .page(selector)
            .map_err(litchi_odf_common::Error::from)
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

    /// Returns the supported password and signature lifecycle for this source.
    #[must_use]
    pub fn security_capabilities(&self) -> SecurityCapabilities {
        self.package.security_capabilities()
    }

    /// Reads inert document and macro signature metadata.
    ///
    /// # Errors
    ///
    /// Returns an error when signature XML cannot be decoded safely.
    pub fn digital_signatures(
        &self,
    ) -> FamilyResult<litchi_odf_common::signature::DigitalSignatures> {
        self.package
            .digital_signatures()
            .map_err(litchi_odf_common::Error::from)
    }

    /// Verifies signature math without making a certificate trust decision.
    ///
    /// # Errors
    ///
    /// Returns an error when signature metadata or referenced bytes cannot be verified.
    pub fn verify_document_signatures(
        &self,
    ) -> FamilyResult<Vec<litchi_odf_common::signature::SignatureVerification>> {
        self.package
            .verify_document_signatures()
            .map_err(litchi_odf_common::Error::from)
    }

    /// Returns inert script, event, action, link, and embedded-object inventory.
    #[must_use]
    pub fn active_content(&self) -> ActiveContentStatus {
        self.package.active_content()
    }

    /// Lists safe package member names.
    ///
    /// # Errors
    ///
    /// Returns an error when package member validation fails.
    pub fn files(&self) -> FamilyResult<Vec<String>> {
        self.package.files().map_err(litchi_odf_common::Error::from)
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

    /// Returns inert drawing style definitions from content and styles parts.
    #[must_use]
    pub fn style_definitions(&self) -> &[crate::style::Style] {
        self.package.style_definitions()
    }

    /// Returns inert named drawing resources referenced by automatic styles.
    #[must_use]
    pub fn style_resources(&self) -> &[crate::style_resource::StyleResource] {
        self.package.style_resources()
    }

    /// Resolves a checked group subtree and its nested flattened descendants.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid or non-group selector.
    pub fn group(&self, page: usize, shape: usize) -> FamilyResult<crate::group::Group> {
        self.package
            .group(page, shape)
            .map_err(litchi_odf_common::Error::from)
    }

    /// Reads one package-local resource without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable member.
    pub fn resource_bytes(&self, resource: usize) -> FamilyResult<Option<Vec<u8>>> {
        self.package
            .resource_bytes(resource)
            .map_err(litchi_odf_common::Error::from)
    }

    /// Prepares a provenance-bound shape or complete group subtree for cross-drawing transfer.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, noncompact source, or unreadable dependencies.
    pub fn prepare_shape_transfer(&self, page: usize, shape: usize) -> FamilyResult<ShapeTransfer> {
        self.package
            .prepare_shape_transfer(page, shape)
            .map_err(litchi_odf_common::Error::from)
    }
}
