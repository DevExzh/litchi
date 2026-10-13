//! Typed document authoring, finalized as a canonical package.
//!
//! The authoring verbs of the legacy writer (paragraphs, tables, images,
//! headers and the rest of [`MutableDocument`]) live here, on fresh content
//! only. An opened package never hydrates into this writer: it changes
//! through [`Package::edit`] and its scoped `edit_*` families (ADR 0034 §9).

use crate::writer::MutableDocument;
use crate::{Package, Result};

/// A native authoring owner whose output is an immutable-editable package.
///
/// Publication lives on the package returned by [`Self::build`].
///
/// ```
/// use litchi_docx::Builder;
///
/// let mut builder = Builder::new()?;
/// let document = builder.document()?;
/// document.add_paragraph_with_text("Hello, World!");
/// let table = document.add_table(2, 2);
/// if let Some(cell) = table.cell(0, 0) {
///     cell.set_text("Header");
/// }
/// let mut package = builder.build()?;
/// assert!(package.document()?.text()?.contains("Hello, World!"));
/// let bytes = package.to_bytes()?;
/// # let _ = bytes;
/// # Ok::<(), litchi_docx::Error>(())
/// ```
pub struct Builder {
    package: Package,
}

impl Builder {
    /// Start the native minimal empty document.
    ///
    /// # Errors
    ///
    /// Returns native construction failures.
    pub fn new() -> Result<Self> {
        Ok(Self {
            package: Package::new()?,
        })
    }

    /// Start a macro-free Word template (`.dotx`).
    ///
    /// # Errors
    ///
    /// Returns native construction failures.
    pub fn new_template() -> Result<Self> {
        Ok(Self {
            package: Package::new_template()?,
        })
    }

    /// The typed document being authored.
    ///
    /// # Errors
    ///
    /// Returns native authoring admission failures.
    pub fn document(&mut self) -> Result<&mut MutableDocument> {
        self.package.authoring_document()
    }

    /// The package under construction, for the package-level authoring
    /// verbs (glossary entries, alternative chunks, custom XML and similar).
    ///
    /// Its content is finalized by [`Self::build`].
    pub fn package(&mut self) -> &mut Package {
        &mut self.package
    }

    /// Select automatic font embedding during finalization.
    ///
    /// # Errors
    ///
    /// Returns native font-policy admission failures. Font discovery and
    /// embedding failures are returned by [`Self::build`].
    #[cfg(feature = "automatic-fonts")]
    pub fn font_embedding(&mut self, mode: litchi_fonts::embedding::Mode) -> Result<&mut Self> {
        self.package.set_font_embedding(mode)?;
        Ok(self)
    }

    /// Finalize the authored document into the package graph without
    /// serializing and reopening a ZIP.
    ///
    /// # Errors
    ///
    /// Returns native graph construction and validation failures.
    pub fn build(self) -> Result<Package> {
        self.package.finish_authoring()
    }
}

impl std::fmt::Debug for Builder {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Builder").finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn build_finalizes_authored_content_for_immutable_edits() -> Result<()> {
        let mut builder = Builder::new()?;
        builder.document()?.add_paragraph_with_text("before");
        let package = builder.build()?;
        assert_eq!(package.document()?.text()?, "before");

        let mut edit = package
            .edit()
            .map_err(|error| crate::Error::Other(error.to_string()))?;
        edit.replace_paragraph_text(litchi_core::Position::new(0), "after")
            .map_err(|error| crate::Error::Other(error.to_string()))?;
        let changed = edit
            .commit()
            .map_err(|error| crate::Error::Other(error.to_string()))?
            .into_package();
        assert_eq!(changed.document()?.text()?, "after");
        assert_eq!(package.document()?.text()?, "before");

        let mut copy = changed.clone();
        let reopened = Package::from_bytes(&copy.to_bytes()?)?;
        assert_eq!(reopened.document()?.text()?, "after");
        Ok(())
    }
}
