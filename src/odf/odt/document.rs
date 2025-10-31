//! OpenDocument Text document structure and API.

use crate::common::{Error, Metadata, Result};
use crate::odf::core::{Content, Meta, Package, Styles};
use crate::odf::elements::style::{StyleElements, StyleRegistry};
use crate::odf::elements::table::Table as ElementTable;
use crate::odf::elements::text::{Paragraph as ElementParagraph, TextElements};
use std::io::Cursor;
use std::path::Path;

/// An OpenDocument text document (.odt).
///
/// This struct represents a complete ODT document and provides methods to access
/// its content, structure, styles, and metadata. Documents are immutable after loading
/// to ensure thread safety and performance.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::Document;
///
/// # fn main() -> litchi::Result<()> {
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
    package: Package<Cursor<Vec<u8>>>,
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("my_document.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let bytes = std::fs::read("document.odt")?;
    /// let doc = Document::from_bytes(bytes)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let cursor = Cursor::new(bytes);
        let package = Package::from_reader(cursor)?;

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
            package,
            content,
            styles,
            meta,
            style_registry,
        })
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let tables = doc.tables()?;
    ///
    /// println!("Found {} tables", tables.len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables(&self) -> Result<Vec<ElementTable>> {
        use crate::odf::elements::table::TableElements;
        TableElements::parse_tables_from_content(self.content.xml_content())
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method extracts both paragraphs and tables, interleaved in the order
    /// they appear in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let elements = doc.elements()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn elements(&self) -> Result<Vec<crate::document::DocumentElement>> {
        use crate::odf::elements::parser::{DocumentOrderElement, DocumentParser};

        // Parse all elements in document order using the generic ODF parser
        let ordered_elements = DocumentParser::parse_elements_in_order(self.content.xml_content())?;
        let mut elements = Vec::new();

        for element in ordered_elements {
            match element {
                DocumentOrderElement::Paragraph(para) => {
                    elements.push(crate::document::DocumentElement::Paragraph(
                        crate::document::Paragraph::Odt(para),
                    ));
                },
                DocumentOrderElement::Heading(heading) => {
                    // Convert heading to paragraph for unified API
                    if let Ok(text) = heading.text() {
                        let mut para = ElementParagraph::new();
                        para.set_text(&text);
                        if let Some(style) = heading.style_name() {
                            para.set_style_name(style);
                        }
                        elements.push(crate::document::DocumentElement::Paragraph(
                            crate::document::Paragraph::Odt(para),
                        ));
                    }
                },
                DocumentOrderElement::Table(table) => {
                    elements.push(crate::document::DocumentElement::Table(
                        crate::document::Table::Odt(table),
                    ));
                },
                DocumentOrderElement::List(_list) => {
                    // Lists could be converted to paragraphs or handled separately
                    // For now, skip lists in the unified document element API
                    // as they are typically expanded to paragraphs by text extraction
                },
            }
        }

        Ok(elements)
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
            Ok(meta.extract_metadata())
        } else {
            Ok(Metadata::default())
        }
    }

    /// Get the style registry for this document.
    ///
    /// The style registry contains all styles defined in the document,
    /// including both automatic styles and named styles.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let styles = doc.styles();
    /// // Use the style registry to query styles...
    /// # Ok(())
    /// # }
    /// ```
    pub fn styles(&self) -> &StyleRegistry {
        &self.style_registry
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    ) -> crate::odf::elements::style::StyleProperties<'_> {
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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

    /// Get all comments/annotations in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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

    /// Get all bookmarks in the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    pub fn bookmarks(&self) -> Result<Vec<crate::odf::elements::bookmark::Bookmark>> {
        use crate::odf::elements::bookmark::BookmarkParser;
        BookmarkParser::parse_bookmarks(self.content.xml_content())
    }

    /// Get all bookmark ranges in the document.
    ///
    /// Bookmark ranges span multiple paragraphs or sections.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
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
    pub fn bookmark_ranges(&self) -> Result<Vec<crate::odf::elements::bookmark::BookmarkRange>> {
        use crate::odf::elements::bookmark::BookmarkParser;
        BookmarkParser::parse_bookmark_ranges(self.content.xml_content())
    }

    /// Get all fields in the document.
    ///
    /// Fields are dynamic content elements like page numbers, dates, and references.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let fields = doc.fields()?;
    ///
    /// for field in fields {
    ///     println!("Field type: {}, value: {}", field.field_type(), field.value());
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn fields(&self) -> Result<Vec<crate::odf::elements::field::Field>> {
        use crate::odf::elements::field::FieldParser;
        FieldParser::parse_fields(self.content.xml_content())
    }

    /// Get all tables with repeated cells and rows expanded.
    ///
    /// ODF files can store repeated cells/rows compactly. This method expands
    /// them into their full representation.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::Document;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let expanded_tables = doc.tables_expanded()?;
    ///
    /// for table in expanded_tables {
    ///     println!("Expanded table has {} rows", table.row_count()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables_expanded(&self) -> Result<Vec<crate::odf::elements::table::Table>> {
        use crate::odf::elements::table_expansion::TableExpander;
        let tables = self.tables()?;
        TableExpander::expand_tables(tables)
    }
}
