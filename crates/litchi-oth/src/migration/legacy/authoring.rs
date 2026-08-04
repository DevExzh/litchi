//! Authoring support for legacy OpenDocument web templates (`.oth`).
//!
//! Both types are thin wrappers over the ODT authoring machinery
//! ([`DocumentBuilder`] and [`MutableDocument`]) that pin the legacy
//! `application/vnd.oasis.opendocument.text-web` root MIME type. Every
//! produced package is re-validated through [`WebDocument`] before it is
//! returned, so authoring can never emit a package the reader rejects.

use super::WebDocument;
use crate::constants;
use crate::odt::{Document, DocumentBuilder, MutableDocument};
use litchi_core::Result;
use std::path::Path;

/// Builder for creating new `.oth` Writer/Web templates from scratch.
///
/// This wraps [`DocumentBuilder`] and emits the legacy
/// `application/vnd.oasis.opendocument.text-web` MIME type instead of the
/// standard text MIME type. All content APIs are available through
/// [`Self::builder_mut`].
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::WebDocumentBuilder;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut builder = WebDocumentBuilder::new();
/// builder.builder_mut().add_heading("Portal template", 1)?;
/// builder.builder_mut().add_paragraph("Reusable web body text.")?;
/// builder.save("portal.oth")?;
/// # Ok(())
/// # }
/// ```
pub struct WebDocumentBuilder {
    builder: DocumentBuilder,
}

impl WebDocumentBuilder {
    /// Create a new empty web-template builder.
    pub fn new() -> Self {
        Self {
            builder: DocumentBuilder::new(),
        }
    }

    /// Borrow the wrapped text document builder.
    pub fn builder(&self) -> &DocumentBuilder {
        &self.builder
    }

    /// Borrow the wrapped text document builder mutably.
    pub fn builder_mut(&mut self) -> &mut DocumentBuilder {
        &mut self.builder
    }

    /// Build the `.oth` package bytes, validated through [`WebDocument`].
    pub fn build(self) -> Result<Vec<u8>> {
        let bytes = self.builder.build_package(constants::ODF_WEB)?;
        WebDocument::from_bytes(bytes.clone())?;
        Ok(bytes)
    }

    /// Build the template and return it as a validated [`WebDocument`].
    pub fn build_document(self) -> Result<WebDocument> {
        WebDocument::from_bytes(self.build()?)
    }

    /// Build and save the template to a file.
    pub fn save(self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.build()?)?;
        Ok(())
    }
}

impl Default for WebDocumentBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Mutable semantic editor for `.oth` Writer/Web templates.
///
/// Obtained from [`WebDocument::into_mutable`], [`WebDocument::to_mutable`],
/// or [`Self::from_document`] for save-as-template conversion of an existing
/// text document. Edits go through the packaged
/// [`MutableDocument`](crate::MutableDocument) authoring APIs;
/// `to_bytes`/`save` emit the web-template MIME type and are re-validated
/// through [`WebDocument`].
pub struct MutableWebDocument {
    document: MutableDocument,
}

impl MutableWebDocument {
    fn with_web_mimetype(mut document: MutableDocument) -> Self {
        document.set_mimetype(constants::ODF_WEB);
        Self { document }
    }

    /// Convert a packaged text document into a mutable `.oth` template
    /// (save-as-template). Package parts the editor does not model are
    /// carried over unchanged by the wrapped editor.
    pub fn from_document(document: Document) -> Result<Self> {
        Ok(Self::with_web_mimetype(MutableDocument::from_document(
            document,
        )?))
    }

    /// Borrow the wrapped text document editor.
    pub fn document(&self) -> &MutableDocument {
        &self.document
    }

    /// Borrow the wrapped text document editor mutably.
    pub fn document_mut(&mut self) -> &mut MutableDocument {
        &mut self.document
    }

    /// Serialize the edited template as `.oth` package bytes, validated
    /// through [`WebDocument`].
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let bytes = self.document.to_bytes()?;
        WebDocument::from_bytes(bytes.clone())?;
        Ok(bytes)
    }

    /// Save the edited template to a file.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.to_bytes()?)?;
        Ok(())
    }
}
