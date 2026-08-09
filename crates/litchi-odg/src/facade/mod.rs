//! Concise family entry points.

use litchi_core::{Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;
pub use crate::package::{Commit, Patch, Snapshot, TextChange, Transaction};

/// Immutable source-owning drawing facade.
pub struct Drawing {
    package: Snapshot,
}

impl Drawing {
    /// Opens a drawing package from a file path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a drawing package from in-memory bytes.
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

    /// Returns declared drawing layers in source order.
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
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
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
