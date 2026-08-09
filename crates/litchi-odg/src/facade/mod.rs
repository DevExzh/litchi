//! Concise family entry points.

use litchi_core::{Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;
pub use crate::package::{
    Change, Commit, GeometryChange, History, HistoryLimits, LayerChange, NameChange, Patch,
    ResourceChange, Snapshot, StructureChange, StyleChange, TextChange, Transaction,
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

    /// Opens a drawing package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a structurally valid ODG.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Snapshot::from_bytes(bytes).map(|package| Self { package })
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
