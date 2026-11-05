//! Typed document authoring, finalized as a canonical package.
//!
//! [`Builder`] authors fresh content only: [`Builder::document`] is the draft
//! [`Document`] whose paragraphs, tables, images, headers and the rest of its
//! verbs are staged here and materialized into the package graph by
//! [`Builder::build`]. The drafts reached from it live in this module without
//! a prefix ([`Document`], [`Paragraph`], [`Run`], [`Table`], [`Row`],
//! [`Cell`], [`Style`], …; record 0351 §2.1 rule 7). An opened package never
//! hydrates into this writer: it changes through [`Package::edit`] and its
//! scoped `edit_*` families (ADR 0034 §9).

pub use crate::writer::bookmark::Bookmark;
pub use crate::writer::comment::Comment;
pub use crate::writer::content_control::ContentControl;
pub use crate::writer::doc::Document;
pub use crate::writer::field::Field;
pub use crate::writer::hyperlink::Hyperlink;
pub use crate::writer::image::InlineImage;
pub use crate::writer::ole_object::OleObject;
pub use crate::writer::paragraph::Paragraph;
pub use crate::writer::revision::{Conflict, CustomXmlConflictRange, Revision};
pub use crate::writer::run::Run;
pub use crate::writer::section::borders;
pub use crate::writer::smart_tag::{SmartTag, SmartTagAttribute};
pub use crate::writer::smartart::SmartArt;
pub use crate::writer::style::Style;
pub use crate::writer::table::{Cell, Row, Table};
pub use crate::writer::textbox::TextBox;
pub use crate::writer::theme::Theme;
pub use crate::writer::vml_shape::VmlShape;
pub use crate::writer::{
    BibliographyFieldSpec, BibliographyFilter, CellProperties, CellRevisionKind, ChapterSep,
    Checksum, CitationFieldSpec, CitationSource, ColorScheme, ConflictKind, ContentControlType,
    DataBinding, Date, DateError, EndnotePos, Endnotes, FootnotePos, Footnotes, FormattingAllowed,
    GridType, ImageFormat, ImageWatermark, ImageWatermarkAnchor, LineNumberRestart, ListType,
    MAX_DATE_BYTES, MAX_OLE_PAYLOAD_BYTES, MAX_SMART_ARTS, MAX_WATERMARK_IMAGE_BYTES, Note,
    NoteNumberRestart, PageNumberFormat, PageOrientation, Protection, RevisionContentControl,
    RevisionKind, RevisionMetadata, RowRevisionKind, RunContent, SectionColumn, SectionColumns,
    SectionDocumentGrid, SectionHeaderFooterPart, SectionHeaderFooterReference,
    SectionLineNumbering, SectionPageNumbering, SectionPaperSource, SectionProperties,
    SectionTextDirection, SectionVerticalAlignment, TableBorder, TableBorders,
    TableCellMergeRevisionState, TableOfContents, VmlShapeKind, VmlShapePosition, Watermark,
    WatermarkLayout, generate_styles_xml,
};

use crate::alt::{Chunk, Import};
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
/// let package = builder.build()?;
/// assert!(package.document()?.text()?.contains("Hello, World!"));
/// let bytes = package.to_bytes()?;
/// # let _ = bytes;
/// # Ok::<(), litchi_docx::Error>(())
/// ```
pub struct Builder {
    package: Package,
    document: Document,
    #[cfg(feature = "automatic-fonts")]
    font_embedding: Option<litchi_fonts::embedding::Mode>,
}

impl Builder {
    fn with_package(package: Package) -> Self {
        Self {
            package,
            document: Document::new(),
            #[cfg(feature = "automatic-fonts")]
            font_embedding: None,
        }
    }

    /// Start the native minimal empty document.
    ///
    /// # Errors
    ///
    /// Returns native construction failures.
    pub fn new() -> Result<Self> {
        Ok(Self::with_package(Package::new()?))
    }

    /// Start a macro-free Word template (`.dotx`).
    ///
    /// Template packages are the native container for reusable `AutoText`
    /// and other building blocks, authored after [`Self::build`] through
    /// [`Package::edit_glossary`].
    ///
    /// # Errors
    ///
    /// Returns native construction failures.
    pub fn new_template() -> Result<Self> {
        Ok(Self::with_package(Package::new_template()?))
    }

    /// The package under construction, for crate tests that stage package
    /// state between authoring steps.
    #[cfg(test)]
    pub(crate) fn package_mut(&mut self) -> &mut Package {
        &mut self.package
    }

    /// The draft document being authored.
    ///
    /// # Errors
    ///
    /// Returns native authoring admission failures.
    pub fn document(&mut self) -> Result<&mut Document> {
        Ok(&mut self.document)
    }

    /// Append a package-backed alternative-format import to the document
    /// body.
    ///
    /// # Errors
    ///
    /// Returns an error if the import's target part cannot be installed.
    pub fn add_alt(&mut self, import: Import, match_source: Option<bool>) -> Result<Chunk> {
        let index = self.document.alts().len();
        self.insert_alt(index, import, match_source)
    }

    /// Insert a package-backed alternative-format import by anchor-relative
    /// index. Part, relationship, and body changes are undone together on
    /// error.
    ///
    /// # Errors
    ///
    /// Returns an error for an index out of range, an exhausted anchor limit
    /// or a target part that cannot be installed.
    pub fn insert_alt(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        self.package
            .insert_alt(&mut self.document, index, import, match_source)
    }

    /// Replace an anchor and its relationship as one change; returns the
    /// replaced anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for an index out of range or a target part that
    /// cannot be installed.
    pub fn replace_alt(
        &mut self,
        index: usize,
        import: Import,
        match_source: Option<bool>,
    ) -> Result<Chunk> {
        self.package
            .replace_alt(&mut self.document, index, import, match_source)
    }

    /// Remove an anchor, its relationship, and an unreachable internal
    /// payload; returns the removed anchor.
    ///
    /// # Errors
    ///
    /// Returns an error for an index out of range.
    pub fn remove_alt(&mut self, index: usize) -> Result<Chunk> {
        self.package.remove_alt(&mut self.document, index)
    }

    /// Reorder body anchors without changing their package relationships.
    ///
    /// # Errors
    ///
    /// Returns an error for an index out of range.
    pub fn move_alt(&mut self, from: usize, to: usize) -> Result<()> {
        self.document.move_alt(from, to)
    }

    /// Select automatic font embedding during finalization.
    ///
    /// # Errors
    ///
    /// Returns native font-policy admission failures. Font discovery and
    /// embedding failures are returned by [`Self::build`].
    #[cfg(feature = "automatic-fonts")]
    pub fn font_embedding(&mut self, mode: litchi_fonts::embedding::Mode) -> Result<&mut Self> {
        if !self.document.glyphs_are_complete() {
            return Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "font_embedding",
                "font discovery requires a complete draft document model",
            ));
        }
        self.font_embedding = Some(mode);
        Ok(self)
    }

    /// Bound automatic font preparation for the package finalized by this
    /// builder. The policy is retained by the package's semantic limits.
    #[cfg(feature = "automatic-fonts")]
    #[must_use]
    pub fn font_preparation_limits(
        &mut self,
        limits: litchi_fonts::PreparationLimits,
    ) -> &mut Self {
        let semantic_limits = self.package.semantic_limits().clone().fonts(limits);
        self.package = self.package.with_semantic_limits(semantic_limits);
        self
    }

    /// Finalize the authored document into the package graph without
    /// serializing and reopening a ZIP.
    ///
    /// # Errors
    ///
    /// Returns native graph construction and validation failures.
    pub fn build(self) -> Result<Package> {
        let Self {
            package,
            mut document,
            #[cfg(feature = "automatic-fonts")]
            font_embedding,
        } = self;
        package.finish_authoring(
            &mut document,
            #[cfg(feature = "automatic-fonts")]
            font_embedding,
        )
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

        let reopened = Package::from_bytes(&changed.to_bytes()?)?;
        assert_eq!(reopened.document()?.text()?, "after");
        Ok(())
    }
}
