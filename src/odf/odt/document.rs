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
        let mut package = Package::from_reader(cursor)?;

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
        let mut style_registry = StyleRegistry::new();

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
    /// let mut doc = Document::open("document.odt")?;
    /// let text = doc.text()?;
    /// println!("Text content:\n{}", text);
    /// # Ok(())
    /// # }
    /// ```
    pub fn text(&mut self) -> Result<String> {
        TextElements::extract_text(self.content.xml_content())
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
    /// let mut doc = Document::open("document.odt")?;
    /// let paragraphs = doc.paragraphs()?;
    ///
    /// for para in paragraphs {
    ///     println!("Paragraph: {}", para.text()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn paragraphs(&mut self) -> Result<Vec<ElementParagraph>> {
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
    /// let mut doc = Document::open("document.odt")?;
    /// let tables = doc.tables()?;
    ///
    /// println!("Found {} tables", tables.len());
    /// # Ok(())
    /// # }
    /// ```
    pub fn tables(&mut self) -> Result<Vec<ElementTable>> {
        use crate::odf::elements::table::TableElements;
        TableElements::parse_tables_from_content(self.content.xml_content())
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
}
