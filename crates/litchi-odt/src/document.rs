//! OpenDocument Text document structure and API.

use crate::core::{Content, Meta, OwnedPackage, Styles};
use crate::elements::style::{StyleElements, StyleRegistry};
use crate::elements::table::Table as ElementTable;
use crate::elements::text::{Paragraph as ElementParagraph, TextElements};
use crate::elements::xml::{
    DRAW_NAMESPACE, TEXT_NAMESPACE, XLINK_NAMESPACE, append_checked, append_text_control,
    decode_reference, is_bound, namespaced_attribute,
};
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use std::collections::HashSet;
use std::path::Path;

use super::header_footer::{MasterPage, parse_master_pages};
use super::page_layout::{PageLayout, parse_page_layouts};
use super::page_sequence::{OdtPageSequence, parse_page_sequence};

const MAX_REFERENCE_DEPTH: usize = 4_096;
const MAX_REFERENCES: usize = 1_000_000;

/// An OpenDocument text document (.odt).
///
/// This struct represents a complete ODT document and provides methods to access
/// its content, structure, styles, and metadata. Documents are immutable after loading
/// to ensure thread safety and performance.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::Document;
///
/// # fn main() -> litchi_core::Result<()> {
/// // Open a document
/// let mut doc = Document::open("document.odt")?;
///
/// // Extract text
/// let text = doc.text()?;
/// println!("Text: {}", text);
///
/// // Get metadata
/// let metadata = doc.metadata()?;
/// if let Some(title) = &metadata.title {
///     println!("Title: {}", title);
/// }
///
/// // Access structured elements
/// let paragraphs = doc.paragraphs()?;
/// let tables = doc.tables()?;
///
/// println!("Paragraphs: {}, Tables: {}", paragraphs.len(), tables.len());
/// # Ok(())
/// # }
/// ```
#[allow(dead_code)]
pub struct Document {
    /// ZIP package containing all document files
    package: OwnedPackage,
    /// Parsed content.xml (main document content)
    content: Content,
    /// Parsed styles.xml (document styles), if present
    styles: Option<Styles>,
    /// Parsed meta.xml (document metadata), if present
    meta: Option<Meta>,
    /// Registry of all styles in the document
    style_registry: StyleRegistry,
}

impl Document {
    crate::script_package::script_facade_methods!();
    crate::annotation_package::annotation_facade_methods!(Text);

    pub(crate) fn into_package(self) -> OwnedPackage {
        self.package
    }

    /// Open an ODT document from a file path.
    ///
    /// This method reads the entire file into memory and parses it. For large files,
    /// consider using `from_bytes` with a streaming reader if memory is constrained.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .odt file
    ///
    /// # Errors
    ///
    /// Returns an error if:
    /// - The file cannot be read
    /// - The file is not a valid ZIP archive
    /// - The file is not a valid ODT document
    /// - Required XML components are malformed
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("my_document.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
    }

    /// Open a password-encrypted ODT document.
    pub fn open_with_password<P: AsRef<Path>>(
        path: P,
        password: impl Into<String>,
    ) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes_with_password(bytes, password)
    }

    /// Create a Document from a byte buffer.
    ///
    /// This is useful when you have the document data in memory already,
    /// such as from network transfers or embedded resources.
    ///
    /// # Arguments
    ///
    /// * `bytes` - Complete ODT file contents as bytes
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes do not represent a valid ODT document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let bytes = std::fs::read("document.odt")?;
    /// let doc = Document::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let owned_package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned_package(owned_package)
    }

    /// Create a document from password-encrypted ODT bytes.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_owned_package(OwnedPackage::from_bytes_with_password(bytes, password)?)
    }

    pub(crate) fn from_owned_package(owned_package: OwnedPackage) -> Result<Self> {
        let package = owned_package.package()?;

        // Verify this is a text document
        let mime_type = package.mimetype();
        if !mime_type.contains("opendocument.text") {
            return Err(Error::InvalidFormat(format!(
                "Not an ODT file: MIME type is {}",
                mime_type
            )));
        }

        // Parse core components
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_bytes(&styles_bytes)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_bytes(&meta_bytes)?)
        } else {
            None
        };

        // Initialize style registry
        let mut style_registry = StyleRegistry::default();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles
            && let Ok(registry) = StyleElements::parse_styles(styles_part.xml_content())
        {
            style_registry = registry;
        }

        // Also parse styles from content.xml (automatic styles)
        if let Ok(content_registry) = StyleElements::parse_styles(content.xml_content()) {
            // Merge content styles into main registry (content styles take precedence)
            for (_name, style) in content_registry.styles {
                style_registry.add_style(style);
            }
        }

        Ok(Self {
            package: owned_package,
            content,
            styles,
            meta,
            style_registry,
        })
    }

    pub fn original_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Create an ODT document from raw bytes (ZIP archive data).
    ///
    /// This is used for single-pass parsing where the ZIP archive has already
    /// been validated during format detection. It avoids double-parsing.
    pub fn from_archive_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes(bytes)
    }

    /// Extract all text content from the document.
    ///
    /// This method extracts plain text from all paragraphs, headings, and text elements
    /// in the document, preserving paragraph breaks. Formatting, styles, and non-text
    /// elements are omitted.
    ///
    /// # Performance
    ///
    /// This method parses the XML on each call. For repeated access, consider caching
    /// the result.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let text = doc.text()?;
    /// println!("Text content:\n{}", text);
    /// # Ok(())
    /// # }
    /// ```
    pub fn text(&self) -> Result<String> {
        TextElements::extract_text(self.content.xml_content())
    }

    /// Get the number of paragraphs in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let count = doc.paragraph_count()?;
    /// println!("Paragraph count: {}", count);
    /// # Ok(())
    /// # }
    /// ```
    pub fn paragraph_count(&self) -> Result<usize> {
        Ok(self.paragraphs()?.len())
    }

    /// Get all paragraphs in the document as structured elements.
    ///
    /// Returns a vector of `Paragraph` elements that can be used to access
    /// individual paragraph content, styles, and attributes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let paragraphs = doc.paragraphs()?;
    ///
    /// for para in paragraphs {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn paragraphs(&self) -> Result<Vec<ElementParagraph>> {
        TextElements::parse_paragraphs(self.content.xml_content())
    }

    /// Get all tables in the document.
    ///
    /// Returns a vector of `Table` elements representing all tables found
    /// in the document body.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let tables = doc.tables()?;
    ///
    /// println!("Found {} tables", tables.len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables(&self) -> Result<Vec<ElementTable>> {
        use crate::elements::table::TableElements;
        TableElements::parse_tables_from_content(self.content.xml_content())
    }

    /// Get all document elements (paragraphs, headings, and tables) in document order.
    ///
    /// This method extracts both paragraphs (including headings) and tables, interleaved
    /// in the order they appear in the document. This provides a more efficient way to
    /// iterate through document content than calling `paragraphs()` and `tables()` separately,
    /// and preserves the exact document order.
    ///
    /// # Returns
    ///
    /// A vector of `DocumentOrderElement` containing all paragraphs, headings, tables, and
    /// lists in the order they appear in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    /// use litchi_odf::elements::parser::DocumentOrderElement;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let elements = doc.elements()?;
    ///
    /// for element in elements {
    ///     match element {
    ///         DocumentOrderElement::Paragraph(para) => {
    ///             println!("Paragraph: {}", para.text()?);
    ///         },
    ///         DocumentOrderElement::NumberedParagraph(para) => {
    ///             println!("Numbered paragraph: {}", para.text()?);
    ///         },
    ///         DocumentOrderElement::Heading(heading) => {
    ///             println!("Heading: {}", heading.text()?);
    ///         },
    ///         DocumentOrderElement::Table(table) => {
    ///             println!("Table with {} rows", table.row_count()?);
    ///         },
    ///         DocumentOrderElement::List(_) => {
    ///             println!("List element");
    ///         },
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn elements(&self) -> Result<Vec<crate::elements::parser::DocumentOrderElement>> {
        use crate::elements::parser::DocumentParser;

        // Parse all elements in document order using the generic ODF parser
        DocumentParser::parse_elements_in_order(self.content.xml_content())
    }

    /// Get document metadata.
    ///
    /// Extracts metadata from the meta.xml file, including title, author,
    /// creation date, modification date, word count, and other document properties.
    ///
    /// # Returns
    ///
    /// A `Metadata` struct containing all available metadata fields. Fields that
    /// are not present in the document will be `None`.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let metadata = doc.metadata()?;
    ///
    /// if let Some(title) = &metadata.title {
    ///     println!("Title: {}", title);
    /// }
    /// if let Some(author) = &metadata.author {
    ///     println!("Author: {}", author);
    /// }
    /// if let Some(word_count) = metadata.word_count {
    ///     println!("Words: {}", word_count);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn metadata(&self) -> Result<Metadata> {
        if let Some(meta) = &self.meta {
            meta.try_extract_metadata()
        } else {
            Ok(Metadata::default())
        }
    }

    /// Get the complete format-specific OpenDocument metadata model.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.meta.as_ref().map(Meta::odf_metadata).transpose()
    }

    /// Get the style registry for this document.
    ///
    /// The style registry contains all styles defined in the document,
    /// including both automatic styles and named styles.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let styles = doc.styles();
    /// // Use the style registry to query styles...
    /// # Ok(())
    /// # }
    /// ```
    pub fn styles(&self) -> &StyleRegistry {
        &self.style_registry
    }

    /// Get typed named ruby style definitions from `styles.xml`.
    pub fn ruby_styles(&self) -> Result<crate::RubyStyles> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::parse_ruby_styles(styles.xml_content()),
        )
    }

    /// Return font-face declarations stored in `content.xml`.
    ///
    /// This preserves stored metadata only. It does not fetch linked font
    /// resources, load fonts, or inspect embedded font data.
    pub fn content_font_face_declarations(&self) -> Result<Option<crate::font_face::Faces>> {
        crate::font_face::parse_content_font_face_declarations(self.content.xml_content())
    }

    /// Return font-face declarations stored in `styles.xml`.
    ///
    /// This preserves stored metadata only. It does not fetch linked font
    /// resources, load fonts, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<crate::font_face::Faces>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| crate::font_face::parse_styles_font_face_declarations(styles.xml_content()),
        )
    }

    /// Return named legacy and SVG drawing gradients from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites, load external data, or render gradients.
    pub fn drawing_gradients(&self) -> Result<crate::drawing_gradient::OdfDrawingGradients> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_gradient::parse_drawing_gradients(styles.xml_content()),
        )
    }

    /// Return named drawing hatch resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing_hatch::OdfDrawingHatches> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_hatch::parse_drawing_hatches(styles.xml_content()),
        )
    }

    /// Return named drawing stroke-dash resources from `styles.xml`.
    ///
    /// This exposes stored common-style resources only. It does not resolve
    /// style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing_stroke_dash::OdfDrawingStrokeDashes> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_stroke_dash::parse_drawing_stroke_dashes(styles.xml_content()),
        )
    }

    /// Return named drawing fill-image definitions from `styles.xml`.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites, follow links, load linked resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing_fill_image::OdfDrawingFillImages> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_fill_image::parse_drawing_fill_images(styles.xml_content()),
        )
    }

    /// Return named drawing marker definitions from `styles.xml`.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing_marker::OdfDrawingMarkers> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_marker::parse_drawing_markers(styles.xml_content()),
        )
    }

    /// Return named drawing opacity definitions from `styles.xml`.
    ///
    /// This exposes stored common-style metadata only. It does not resolve
    /// style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing_opacity::OdfDrawingOpacities> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::drawing_opacity::parse_drawing_opacities(styles.xml_content()),
        )
    }

    /// Parse master pages and their losslessly retained headers and footers.
    pub fn master_pages(&self) -> Result<Vec<MasterPage>> {
        self.styles.as_ref().map_or_else(
            || Ok(Vec::new()),
            |styles| parse_master_pages(styles.xml_content()),
        )
    }

    /// Parse automatic page layouts, their properties, and header/footer styles.
    pub fn page_layouts(&self) -> Result<Vec<PageLayout>> {
        self.styles.as_ref().map_or_else(
            || Ok(Vec::new()),
            |styles| parse_page_layouts(styles.xml_content()),
        )
    }

    /// The unnamed fallback page layout (`style:default-page-layout`), when
    /// the document declares one.
    pub fn default_page_layout(&self) -> Result<Option<PageLayout>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| super::page_layout::parse_default_page_layout(styles.xml_content()),
        )
    }

    /// Return the stored footnote and endnote presentation configurations.
    ///
    /// These style declarations are retained as metadata only. This API does
    /// not renumber notes, resolve style references, or render continuation
    /// notices.
    pub fn notes_configurations(&self) -> Result<crate::OdfNotesConfigurations> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::parse_notes_configurations(styles.xml_content()),
        )
    }

    /// Return stored outline numbering styles.
    ///
    /// These are styles metadata only. This API does not apply styles to
    /// headings, generate labels, or update tables of contents.
    pub fn outline_styles(&self) -> Result<crate::OdfOutlineStyles> {
        self.styles.as_ref().map_or_else(
            || Ok(Default::default()),
            |styles| crate::parse_outline_styles(styles.xml_content()),
        )
    }

    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// This styles metadata remains inert: the API does not generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(
        &self,
    ) -> Result<Option<crate::OdfBibliographyConfiguration>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| {
                crate::bibliography_configuration::parse_bibliography_configuration(
                    styles.xml_content(),
                )
            },
        )
    }

    /// Return the explicit master-page assignments from `text:page-sequence`, if present.
    ///
    /// This exposes stored page metadata only. It does not resolve master-page
    /// styles, calculate page breaks, or paginate the document.
    pub fn page_sequence(&self) -> Result<Option<OdtPageSequence>> {
        parse_page_sequence(self.content.xml_content())
    }

    /// Return the stored document line-numbering configuration.
    ///
    /// This is presentation metadata from styles.xml only. It is never used to
    /// paginate the document or generate line numbers.
    pub fn line_numbering_configuration(
        &self,
    ) -> Result<Option<crate::OdfLineNumberingConfiguration>> {
        self.styles.as_ref().map_or_else(
            || Ok(None),
            |styles| crate::parse_line_numbering_configuration(styles.xml_content()),
        )
    }

    /// Get resolved style properties for a given style name.
    ///
    /// This method resolves style inheritance to provide the complete set of
    /// properties that apply to elements using the specified style.
    ///
    /// # Arguments
    ///
    /// * `style_name` - Name of the style to resolve
    ///
    /// # Returns
    ///
    /// A `StyleProperties` struct containing all resolved properties for the style.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let props = doc.get_style_properties("Heading1");
    ///
    /// if let Some(font_size) = &props.text.font_size {
    ///     println!("Heading 1 font size: {}", font_size);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_style_properties(
        &self,
        style_name: &str,
    ) -> crate::elements::style::StyleProperties<'_> {
        self.style_registry.get_resolved_properties(style_name)
    }

    /// Get all tracked changes in the document.
    ///
    /// Tracked changes include insertions, deletions, and format changes made
    /// by document collaborators.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let changes = doc.track_changes()?;
    ///
    /// for change in changes {
    ///     println!("Change by {:?}: {:?}", change.author, change.change_type);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn track_changes(&self) -> Result<Vec<super::parser::TrackChange>> {
        super::parser::OdtParser::parse_track_changes(self.content.xml_content())
    }

    /// Get tracked changes together with their inert container policy metadata.
    ///
    /// Protection-key material and digest identifiers are retained for round-trip and
    /// inspection only; this method never unlocks, accepts, rejects, or evaluates changes.
    pub fn tracked_changes(&self) -> Result<super::parser::TrackedChanges> {
        super::parser::OdtParser::parse_tracked_changes(self.content.xml_content())
    }

    /// Get all comments/annotations in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let comments = doc.comments()?;
    ///
    /// for comment in comments {
    ///     println!("Comment by {:?}: {}", comment.author, comment.content);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn comments(&self) -> Result<Vec<super::parser::Comment>> {
        super::parser::OdtParser::parse_comments(self.content.xml_content())
    }

    /// Get all sections in the document.
    ///
    /// Sections are document subdivisions that can have protected content,
    /// different formatting, or special layout properties.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let sections = doc.sections()?;
    ///
    /// for section in sections {
    ///     println!("Section '{}': protected={}", section.name, section.protected);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn sections(&self) -> Result<Vec<super::parser::Section>> {
        super::parser::OdtParser::parse_sections(self.content.xml_content())
    }

    /// Get all generated text indexes, including their inert source definitions and cached bodies.
    ///
    /// This covers tables of contents and illustration, table, object, user, alphabetical,
    /// and bibliography indexes. Stored source declarations are never evaluated, and external
    /// alphabetical auto-mark files are never fetched.
    pub fn text_indexes(&self) -> Result<Vec<super::TextIndex>> {
        super::index::parse_text_indexes(self.content.xml_content())
    }

    /// Get point and resolved range marks that contribute entries to generated text indexes.
    pub fn text_index_marks(&self) -> Result<Vec<super::TextIndexMark>> {
        super::index_mark::parse_text_index_marks(self.content.xml_content())
    }

    /// Get point and range targets used by `text:reference-ref` fields.
    pub fn reference_marks(&self) -> Result<Vec<super::ReferenceMark>> {
        super::reference_mark::parse_reference_marks(self.content.xml_content())
    }

    /// Get all semantic footnotes and endnotes in document order.
    pub fn notes(&self) -> Result<Vec<super::Note>> {
        super::note::parse_notes(self.content.xml_content())
    }

    pub fn footnotes(&self) -> Result<Vec<super::Note>> {
        Ok(self
            .notes()?
            .into_iter()
            .filter(|note| note.class() == super::NoteClass::Footnote)
            .collect())
    }

    pub fn endnotes(&self) -> Result<Vec<super::Note>> {
        Ok(self
            .notes()?
            .into_iter()
            .filter(|note| note.class() == super::NoteClass::Endnote)
            .collect())
    }

    /// Get structure-preserving ruby annotations in document order.
    pub fn ruby_annotations(&self) -> Result<crate::RubyAnnotations> {
        crate::parse_ruby_annotations(self.content.xml_content())
    }

    /// Get simplified ruby base/pronunciation pairs in document order.
    pub fn rubies(&self) -> Result<Vec<super::Ruby>> {
        super::ruby::parse_rubies(self.content.xml_content())
    }

    /// Get all bookmarks in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let bookmarks = doc.bookmarks()?;
    ///
    /// for bookmark in bookmarks {
    ///     if let Some(name) = bookmark.name() {
    ///         println!("Bookmark: {}", name);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn bookmarks(&self) -> Result<Vec<crate::elements::bookmark::Bookmark>> {
        use crate::elements::bookmark::BookmarkParser;
        BookmarkParser::parse_bookmarks(self.content.xml_content())
    }

    /// Get all bookmark ranges in the document.
    ///
    /// Bookmark ranges span multiple paragraphs or sections.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let ranges = doc.bookmark_ranges()?;
    ///
    /// for range in ranges {
    ///     if range.is_complete() {
    ///         println!("Complete bookmark range: {}", range.name);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn bookmark_ranges(&self) -> Result<Vec<crate::elements::bookmark::BookmarkRange>> {
        use crate::elements::bookmark::BookmarkParser;
        BookmarkParser::parse_bookmark_ranges(self.content.xml_content())
    }

    /// Get all fields in the document.
    ///
    /// Fields are dynamic content elements like page numbers, dates, and references.
    /// Returned values are the document's cached display text; expressions, database
    /// fields, and other dynamic sources are never evaluated or executed.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let fields = doc.fields()?;
    ///
    /// for field in fields {
    ///     println!("Field type: {}, value: {}", field.field_type(), field.value());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn fields(&self) -> Result<Vec<crate::elements::field::Field>> {
        use crate::elements::field::FieldParser;
        FieldParser::parse_fields(self.content.xml_content())
    }

    /// Get conditional, hidden, and placeholder text fields in document order.
    ///
    /// Conditions and formulas are returned as inert strings. This method never
    /// evaluates them and returns only the cached display text stored in the file.
    pub fn dynamic_text_fields(&self) -> Result<Vec<crate::elements::field::OdfDynamicTextField>> {
        use crate::elements::field::FieldParser;
        FieldParser::parse_dynamic_text_fields(self.content.xml_content())
    }

    /// Get all tables with repeated cells and rows expanded.
    ///
    /// ODF files can store repeated cells/rows compactly. This method expands
    /// them into their full representation.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let expanded_tables = doc.tables_expanded()?;
    ///
    /// for table in expanded_tables {
    ///     println!("Expanded table has {} rows", table.row_count()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables_expanded(&self) -> Result<Vec<crate::elements::table::Table>> {
        use crate::elements::table_expansion::TableExpander;
        let tables = self.tables()?;
        TableExpander::expand_tables(tables)
    }

    // Note: For document modification operations, see `MutableDocument` which provides
    // full CRUD operations (Create, Read, Update, Delete) on document content including
    // adding, updating, and removing paragraphs and tables while preserving insertion order.

    /// Save the document to a new file.
    ///
    /// This method saves the current document state to a new file. Note that this
    /// creates a copy of the original document; modifications are not yet supported.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("input.odt")?;
    /// doc.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Note
    ///
    /// Full document modification support is planned for future releases. For now,
    /// to modify a document, use `DocumentBuilder` to create a new document with
    /// the desired content.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the document to bytes.
    ///
    /// This method serializes the document to an ODF-compliant ZIP archive.
    /// All embedded media files (images, etc.) are automatically copied.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let bytes = doc.to_bytes()?;
    /// // Use bytes for network transfer, etc.
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(self.package.as_bytes().to_vec())
    }

    /// Extract all hyperlinks from the document
    ///
    /// Returns a vector of tuples containing (link text, URL).
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let hyperlinks = doc.hyperlinks()?;
    /// for (text, url) in hyperlinks {
    ///     println!("{}: {}", text, url);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn hyperlinks(&self) -> Result<Vec<(String, String)>> {
        parse_hyperlinks(self.content.xml_content())
    }

    /// Extract all bookmark names from the document
    ///
    /// Returns a vector of bookmark names.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let bookmark_names = doc.bookmark_names()?;
    /// for bookmark in bookmark_names {
    ///     println!("Bookmark: {}", bookmark);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn bookmark_names(&self) -> Result<Vec<String>> {
        parse_bookmark_names(self.content.xml_content())
    }

    /// Extract all linked image references from the document
    ///
    /// Returns package-local and external `xlink:href` values in document order.
    /// Images stored inline as `office:binary-data` have no path and are omitted.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let images = doc.image_paths()?;
    /// for img_path in images {
    ///     println!("Image: {}", img_path);
    ///     // You can extract the image bytes with:
    ///     // let bytes = doc.get_file(&img_path)?;
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn image_paths(&self) -> Result<Vec<String>> {
        parse_image_references(self.content.xml_content())
    }

    /// Discover referenced, inline, missing, and inert linked images.
    pub fn images(&self) -> Result<Vec<crate::Image>> {
        let package = self.package.package()?;
        crate::media::scan_packaged_images(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    /// Inspect classic forms without executing bindings, events, or external resources.
    pub fn forms(&self) -> Result<crate::OdfForms> {
        let mut parts = vec![(self.content.xml_content(), crate::OdfFormPart::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::OdfFormPart::Styles));
        }
        crate::form::parse_form_parts(&parts)
    }

    pub fn rdf_graphs(&self) -> Result<Vec<crate::rdf::Graph>> {
        crate::rdf::graphs(&self.package)
    }
    pub fn add_rdf_graph(
        &mut self,
        preferred_path: Option<&str>,
        triples: &[crate::rdf::Triple],
    ) -> Result<String> {
        let (bytes, path) = crate::rdf::add_graph(&self.package, preferred_path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(path)
    }
    pub fn replace_rdf_graph(&mut self, path: &str, triples: &[crate::rdf::Triple]) -> Result<()> {
        let bytes = crate::rdf::replace_graph(&self.package, path, triples)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_rdf_graph(&mut self, path: &str) -> Result<()> {
        let bytes = crate::rdf::remove_graph(&self.package, path)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn add_rdf_triple(&mut self, path: &str, triple: &crate::rdf::Triple) -> Result<usize> {
        let index = self
            .rdf_graphs()?
            .into_iter()
            .find(|graph| graph.path == path)
            .ok_or_else(|| Error::InvalidFormat(format!("RDF graph '{path}' was not found")))?
            .triples
            .len();
        let (bytes, _) = crate::rdf::add_triple(&self.package, path, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_rdf_triple(
        &mut self,
        path: &str,
        index: usize,
        triple: &crate::rdf::Triple,
    ) -> Result<()> {
        let bytes = crate::rdf::replace_triple(&self.package, path, index, triple)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_rdf_triple(&mut self, path: &str, index: usize) -> Result<()> {
        let bytes = crate::rdf::remove_triple(&self.package, path, index)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_rdf_triple(&mut self, path: &str, from: usize, to: usize) -> Result<()> {
        let bytes = crate::rdf::move_triple(&self.package, path, from, to)?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn add_form(&mut self, group_index: usize, form: &crate::OdfAuthoredForm) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::form_package::FormHost::Text,
            group_index,
            None,
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn add_nested_form(
        &mut self,
        parent_form: usize,
        form: &crate::OdfAuthoredForm,
    ) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::form_package::FormHost::Text,
            0,
            Some(parent_form),
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_form(&mut self, index: usize, form: &crate::OdfAuthoredForm) -> Result<()> {
        let bytes = crate::form_package::replace_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            form,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_form(&mut self, index: usize) -> Result<()> {
        let bytes = crate::form_package::remove_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_form(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::form_package::move_form(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn add_form_control(
        &mut self,
        form_index: usize,
        control: &crate::OdfAuthoredFormControl,
    ) -> Result<usize> {
        let (bytes, index) = crate::form_package::add_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            form_index,
            control,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }
    pub fn replace_form_control(
        &mut self,
        index: usize,
        control: &crate::OdfAuthoredFormControl,
    ) -> Result<()> {
        let bytes = crate::form_package::replace_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            control,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn remove_form_control(&mut self, index: usize) -> Result<()> {
        let bytes = crate::form_package::remove_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }
    pub fn move_form_control(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::form_package::move_control(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::OdfVariableDeclarations> {
        let mut parts = vec![(self.content.xml_content(), crate::OdfVariablePart::Content)];
        if let Some(styles) = self.styles.as_ref().map(Styles::xml_content) {
            parts.push((styles, crate::OdfVariablePart::Styles));
        }
        crate::variable_declaration::parse_variable_declaration_parts(&parts)
    }

    /// Discover package, inline, missing, and inert linked embedded objects.
    pub fn embedded_objects(&self) -> Result<Vec<crate::OdfEmbeddedObject>> {
        let package = self.package.package()?;
        crate::embedded_object::scan_packaged_objects(
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            |path| package.has_file(path),
            |path| package.manifest().get_media_type(path).map(str::to_string),
        )
    }

    /// Open one inert embedded chart as a standalone chart document.
    pub fn embedded_chart(&self, index: usize) -> Result<crate::ChartDocument> {
        crate::embedded_chart::open_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )
    }

    /// Append a packaged chart object to the text body.
    pub fn add_embedded_chart(&mut self, definition: &crate::ChartDefinition) -> Result<usize> {
        self.add_embedded_chart_with_storage(
            definition,
            crate::OdfEmbeddedChartStorage::PackageSubdocument,
        )
    }

    /// Append a chart object using an explicit storage form.
    pub fn add_embedded_chart_with_storage(
        &mut self,
        definition: &crate::ChartDefinition,
        storage: crate::OdfEmbeddedChartStorage,
    ) -> Result<usize> {
        let (bytes, index) = crate::embedded_chart::add_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::embedded_chart::EmbeddedChartHost::Text,
            storage,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_chart(
        &mut self,
        index: usize,
        definition: &crate::ChartDefinition,
    ) -> Result<()> {
        let bytes = crate::embedded_chart::replace_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            definition,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_chart(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_chart::remove_embedded_chart(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Append an inert embedded object or image to the text body.
    pub fn add_embedded_resource(
        &mut self,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<usize> {
        let (bytes, index) = crate::embedded_package::add(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            crate::embedded_chart::EmbeddedChartHost::Text,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(index)
    }

    pub fn replace_embedded_object(
        &mut self,
        index: usize,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::embedded_package::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Object,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn replace_embedded_image(
        &mut self,
        index: usize,
        resource: &crate::OdfEmbeddedResource,
    ) -> Result<()> {
        let bytes = crate::embedded_package::replace(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Image,
            resource,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_object(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_package::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn remove_embedded_image(&mut self, index: usize) -> Result<()> {
        let bytes = crate::embedded_package::remove(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            index,
            crate::embedded_package::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_object(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::embedded_package::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::embedded_package::ResourceTarget::Object,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    pub fn move_embedded_image(&mut self, from: usize, to: usize) -> Result<()> {
        let bytes = crate::embedded_package::reorder(
            &self.package,
            self.content.xml_content(),
            self.styles.as_ref().map(Styles::xml_content),
            from,
            to,
            crate::embedded_package::ResourceTarget::Image,
        )?;
        *self = Self::from_bytes(bytes)?;
        Ok(())
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            crate::ImageSource::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            crate::ImageSource::PackagePart { path, .. } => self.package.get_file(path).map(Some),
            _ => Ok(None),
        }
    }

    /// Get a file from the ODF package (useful for extracting images)
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the file within the package (e.g., "Pictures/image1.png")
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Document;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let images = doc.image_paths()?;
    /// if let Some(first_image) = images.first() {
    ///     let image_bytes = doc.get_file(first_image)?;
    ///     std::fs::write("extracted_image.png", image_bytes)?;
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        self.package.get_file(path)
    }

    // Note: DELETE operations are available via `MutableDocument`. To modify this document:
    //   1. Convert to MutableDocument:  `let mut mutable = MutableDocument::from_document(doc)?`
    //   2. Perform modifications: `mutable.remove_paragraph(0)?`, `mutable.remove_table(1)?`, etc.
    //   3. Save: `mutable.save("output.odt")?`
    // Available methods: remove_paragraph, remove_table, update_paragraph, clear_content, etc.
}

struct ActiveHyperlink {
    href: Option<String>,
    text: String,
    depth: usize,
}

fn parse_hyperlinks(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active: Option<ActiveHyperlink> = None;
    let mut links = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid hyperlink XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_reference_depth(document_depth)?;
                if let Some(link) = active.as_mut() {
                    if text_element && element.local_name().as_ref() == b"a" {
                        return Err(Error::InvalidFormat(
                            "nested text:a hyperlinks are not allowed".to_string(),
                        ));
                    }
                    link.depth += 1;
                    if text_element {
                        append_text_control(&reader, element, &mut link.text)?;
                    }
                } else if text_element && element.local_name().as_ref() == b"a" {
                    active = Some(ActiveHyperlink {
                        href: namespaced_attribute(
                            &reader,
                            element,
                            XLINK_NAMESPACE,
                            b"href",
                            "text:a",
                        )?,
                        text: String::new(),
                        depth: 1,
                    });
                }
            },
            Event::Empty(ref element) => {
                if let Some(link) = active.as_mut() {
                    if text_element && element.local_name().as_ref() == b"a" {
                        return Err(Error::InvalidFormat(
                            "nested text:a hyperlinks are not allowed".to_string(),
                        ));
                    }
                    if text_element {
                        append_text_control(&reader, element, &mut link.text)?;
                    }
                } else if text_element
                    && element.local_name().as_ref() == b"a"
                    && let Some(href) =
                        namespaced_attribute(&reader, element, XLINK_NAMESPACE, b"href", "text:a")?
                {
                    ensure_reference_capacity(links.len(), "hyperlinks")?;
                    links.push((String::new(), href));
                }
            },
            Event::Text(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid hyperlink text: {error}"))
                    })?;
                append_checked(
                    &mut active.as_mut().expect("checked hyperlink").text,
                    &value,
                )?;
            },
            Event::CData(ref value) if active.is_some() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid hyperlink CDATA: {error}"))
                    })?;
                append_checked(
                    &mut active.as_mut().expect("checked hyperlink").text,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if active.is_some() => {
                let value = decode_reference(reference, "hyperlink")?;
                append_checked(
                    &mut active.as_mut().expect("checked hyperlink").text,
                    &value,
                )?;
            },
            Event::End(_) => {
                document_depth = document_depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("hyperlink XML stack underflow".to_string())
                })?;
                if let Some(link) = active.as_mut() {
                    link.depth = link.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("hyperlink element stack underflow".to_string())
                    })?;
                    if link.depth == 0 {
                        let link = active.take().expect("checked hyperlink");
                        if let Some(href) = link.href {
                            ensure_reference_capacity(links.len(), "hyperlinks")?;
                            links.push((link.text, href));
                        }
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if document_depth != 0 || active.is_some() {
        return Err(Error::InvalidFormat(
            "incomplete hyperlink XML structure".to_string(),
        ));
    }
    Ok(links)
}

fn parse_bookmark_names(xml: &str) -> Result<Vec<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut names = Vec::new();
    let mut unique_names = HashSet::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid bookmark XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                depth = checked_reference_depth(depth)?;
                collect_bookmark_name(
                    &reader,
                    text_element,
                    element,
                    &mut names,
                    &mut unique_names,
                )?;
            },
            Event::Empty(ref element) => {
                collect_bookmark_name(
                    &reader,
                    text_element,
                    element,
                    &mut names,
                    &mut unique_names,
                )?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("bookmark XML stack underflow".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete bookmark XML structure".to_string(),
        ));
    }
    Ok(names)
}

fn collect_bookmark_name(
    reader: &NsReader<&[u8]>,
    text_element: bool,
    element: &quick_xml::events::BytesStart<'_>,
    names: &mut Vec<String>,
    unique_names: &mut HashSet<String>,
) -> Result<()> {
    if text_element
        && matches!(
            element.local_name().as_ref(),
            b"bookmark" | b"bookmark-start" | b"bookmark-end"
        )
        && let Some(name) =
            namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "bookmark")?
        && unique_names.insert(name.clone())
    {
        ensure_reference_capacity(names.len(), "bookmark names")?;
        names.push(name);
    }
    Ok(())
}

fn parse_image_references(xml: &str) -> Result<Vec<String>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut references = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid image XML: {error}")))?;
        let draw_element = is_bound(&namespace, DRAW_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                depth = checked_reference_depth(depth)?;
                collect_image_reference(&reader, draw_element, element, &mut references)?;
            },
            Event::Empty(ref element) => {
                collect_image_reference(&reader, draw_element, element, &mut references)?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("image XML stack underflow".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::InvalidFormat(
            "incomplete image XML structure".to_string(),
        ));
    }
    Ok(references)
}

fn collect_image_reference(
    reader: &NsReader<&[u8]>,
    draw_element: bool,
    element: &quick_xml::events::BytesStart<'_>,
    references: &mut Vec<String>,
) -> Result<()> {
    if draw_element
        && element.local_name().as_ref() == b"image"
        && let Some(href) =
            namespaced_attribute(reader, element, XLINK_NAMESPACE, b"href", "draw:image")?
    {
        ensure_reference_capacity(references.len(), "image references")?;
        references.push(href);
    }
    Ok(())
}

fn checked_reference_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODF reference nesting depth overflow".to_string()))?;
    if depth > MAX_REFERENCE_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "ODF reference nesting exceeds {MAX_REFERENCE_DEPTH} levels"
        )));
    }
    Ok(depth)
}

fn ensure_reference_capacity(length: usize, kind: &str) -> Result<()> {
    if length >= MAX_REFERENCES {
        return Err(Error::InvalidFormat(format!(
            "document exceeds {MAX_REFERENCES} {kind}"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod text_model_tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use crate::elements::parser::DocumentOrderElement;

    fn document(content: &str) -> Document {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_TEXT).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        Document::from_bytes(writer.finish_to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn text_model_accepts_arbitrary_prefixes_and_decodes_mixed_text() {
        let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:h t:outline-level="2">Title &amp; More</t:h><t:p t:style-name="Body">A<t:span>B</t:span>C<t:s t:c="2"/>D<![CDATA[!]]></t:p></o:text></o:body></o:document-content>"#;
        let document = document(content);

        assert_eq!(document.text().unwrap(), "Title & More\nABC  D!");
        assert_eq!(document.paragraph_count().unwrap(), 1);
        let paragraph = document.paragraphs().unwrap().remove(0);
        assert_eq!(paragraph.style_name(), Some("Body"));
        assert_eq!(paragraph.text().unwrap(), "ABC  D!");

        let elements = document.elements().unwrap();
        assert_eq!(elements.len(), 2);
        let DocumentOrderElement::Heading(heading) = &elements[0] else {
            panic!("first document element is not a heading");
        };
        assert_eq!(heading.level(), Some(2));
        assert_eq!(heading.text().unwrap(), "Title & More");
        let DocumentOrderElement::Paragraph(paragraph) = &elements[1] else {
            panic!("second document element is not a paragraph");
        };
        assert_eq!(paragraph.text().unwrap(), "ABC  D!");
    }

    #[test]
    fn references_fields_and_images_are_namespace_aware_and_decoded() {
        let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><o:body><o:text><t:p><t:bookmark t:name="point &amp; one"/><t:bookmark-start t:name="range"/>ab<t:s t:c="2"/>c<t:bookmark-end t:name="range"/></t:p><t:p><t:a x:type="simple" x:href="https://example.invalid/?a=1&amp;b=2">A<t:span>B &amp; C</t:span><t:s t:c="2"/>D<![CDATA[!]]></t:a><t:date s:data-style-name="N1" t:fixed="true" t:date-value="2026-07-16">July &amp; 16</t:date><t:word-count>42</t:word-count><d:frame><d:image x:href="Pictures/a&amp;b.png"/><d:image x:href="https://example.invalid/image.png"/><d:image><o:binary-data>AA==</o:binary-data></d:image></d:frame></t:p></o:text></o:body></o:document-content>"#;
        let document = document(content);

        assert_eq!(
            document.hyperlinks().unwrap(),
            vec![(
                "AB & C  D!".to_string(),
                "https://example.invalid/?a=1&b=2".to_string()
            )]
        );
        assert_eq!(
            document.bookmark_names().unwrap(),
            vec!["point & one".to_string(), "range".to_string()]
        );
        let bookmarks = document.bookmarks().unwrap();
        assert_eq!(bookmarks.len(), 1);
        assert_eq!(bookmarks[0].name(), Some("point & one"));
        let ranges = document.bookmark_ranges().unwrap();
        assert_eq!(ranges.len(), 1);
        assert_eq!(ranges[0].name, "range");
        assert_eq!(ranges[0].start, Some((0, 0)));
        assert_eq!(ranges[0].end, Some((0, 5)));
        assert!(ranges[0].is_complete());

        let fields = document.fields().unwrap();
        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].field_type(), "text:date");
        assert_eq!(fields[0].value(), "July & 16");
        assert_eq!(fields[0].format(), Some("N1"));
        assert_eq!(fields[1].field_type(), "text:word-count");
        assert_eq!(fields[1].value(), "42");
        assert_eq!(
            document.image_paths().unwrap(),
            vec![
                "Pictures/a&b.png".to_string(),
                "https://example.invalid/image.png".to_string()
            ]
        );
    }

    #[test]
    fn reference_readers_reject_malformed_xml_and_duplicate_expanded_attributes() {
        let duplicate = r#"<t:p xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:y="http://www.w3.org/1999/xlink"><t:a x:href="a" y:href="b">bad</t:a></t:p>"#;
        assert!(parse_hyperlinks(duplicate).is_err());
        let missing_name =
            r#"<t:p xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:bookmark/></t:p>"#;
        assert!(crate::elements::bookmark::BookmarkParser::parse_bookmarks(missing_name).is_err());
        let nonempty = r#"<t:p xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><t:bookmark t:name="bad">content</t:bookmark></t:p>"#;
        assert!(crate::elements::bookmark::BookmarkParser::parse_bookmarks(nonempty).is_err());
        assert!(parse_image_references("<d:image").is_err());
        assert!(crate::elements::field::FieldParser::parse_fields("<t:date>").is_err());
    }
}
