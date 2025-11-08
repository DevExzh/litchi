//! Mutable document structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODT documents that allows
//! for in-place modification of content, styles, and metadata.

use crate::common::{Metadata, Result};
use crate::odf::core::{OdfStructure, PackageWriter};
use crate::odf::elements::table::Table;
use crate::odf::elements::text::Paragraph;
use crate::odf::odt::Document;
use std::path::Path;

/// Document element type for tracking insertion order
#[derive(Debug, Clone)]
enum DocumentElement {
    /// A paragraph element
    Paragraph(Paragraph),
    /// A table element
    Table(Table),
}

/// A mutable ODT document that supports in-place modifications.
///
/// This struct wraps an ODT document and provides methods to modify its content,
/// including adding, updating, and removing paragraphs, tables, and other elements.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::{Document, MutableDocument};
///
/// # fn main() -> litchi::Result<()> {
/// // Open an existing document
/// let doc = Document::open("input.odt")?;
/// let mut mutable_doc = MutableDocument::from_document(doc)?;
///
/// // Modify the document
/// mutable_doc.add_paragraph("New paragraph")?;
/// mutable_doc.remove_paragraph(0)?;
///
/// // Save the modified document
/// mutable_doc.save("output.odt")?;
/// # Ok(())
/// # }
/// ```
pub struct MutableDocument {
    /// Document elements in insertion order (paragraphs and tables mixed)
    elements: Vec<DocumentElement>,
    /// Document metadata (mutable)
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is for now)
    styles_xml: Option<String>,
    /// Original meta XML (will be regenerated)
    _original_meta: Option<String>,
}

impl MutableDocument {
    /// Create a mutable document from an existing Document.
    ///
    /// This parses the document structure into mutable elements.
    ///
    /// # Arguments
    ///
    /// * `doc` - The document to make mutable
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{Document, MutableDocument};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let mut mutable_doc = MutableDocument::from_document(doc)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_document(doc: Document) -> Result<Self> {
        let paragraphs = doc.paragraphs()?;
        let tables = doc.tables()?;
        let metadata = doc.metadata()?;

        // Get MIME type from package
        let mimetype = "application/vnd.oasis.opendocument.text".to_string();

        // Extract styles XML from the document's package
        let styles_xml = doc
            .get_file("styles.xml")
            .ok()
            .and_then(|bytes| String::from_utf8(bytes).ok());

        // Create elements vector with proper ordering
        // For now, add paragraphs first, then tables (preserving original behavior)
        // Future: Use Document::elements() to preserve exact order
        let mut elements = Vec::new();
        for para in paragraphs {
            elements.push(DocumentElement::Paragraph(para));
        }
        for table in tables {
            elements.push(DocumentElement::Table(table));
        }

        Ok(Self {
            elements,
            metadata,
            mimetype,
            styles_xml,
            _original_meta: None,
        })
    }

    /// Create a new empty mutable document.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::MutableDocument;
    ///
    /// let doc = MutableDocument::new();
    /// ```
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.text".to_string(),
            styles_xml: None,
            _original_meta: None,
        }
    }

    /// Get all paragraphs in the document.
    pub fn paragraphs(&self) -> Vec<&Paragraph> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Paragraph(p) = elem {
                    Some(p)
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all paragraphs (owned) in the document.
    pub fn paragraphs_owned(&self) -> Vec<Paragraph> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Paragraph(p) = elem {
                    Some(p.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all tables in the document.
    pub fn tables(&self) -> Vec<&Table> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Table(t) = elem {
                    Some(t)
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get all tables (owned) in the document.
    pub fn tables_owned(&self) -> Vec<Table> {
        self.elements
            .iter()
            .filter_map(|elem| {
                if let DocumentElement::Table(t) = elem {
                    Some(t.clone())
                } else {
                    None
                }
            })
            .collect()
    }

    /// Get the document metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get a mutable reference to the document metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Add a new paragraph to the end of the document.
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello, World!")?;
    /// doc.add_paragraph("Second paragraph")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph(&mut self, text: &str) -> Result<()> {
        let mut para = Paragraph::new();
        para.set_text(text);
        self.elements.push(DocumentElement::Paragraph(para));
        Ok(())
    }

    /// Insert a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Position to insert at (0-based)
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("First")?;
    /// doc.add_paragraph("Third")?;
    /// doc.insert_paragraph(1, "Second")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        let mut para = Paragraph::new();
        para.set_text(text);

        if index <= self.elements.len() {
            self.elements
                .insert(index, DocumentElement::Paragraph(para));
            Ok(())
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.elements.len()
            )))
        }
    }

    /// Remove a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the paragraph to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Paragraph 1")?;
    /// doc.add_paragraph("Paragraph 2")?;
    /// doc.remove_paragraph(0)?; // Remove first paragraph
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_paragraph(&mut self, index: usize) -> Result<Paragraph> {
        // Find the index of the nth paragraph
        let mut para_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Paragraph(_)) {
                if para_count == index {
                    element_index = Some(i);
                    break;
                }
                para_count += 1;
            }
        }

        if let Some(idx) = element_index {
            if let DocumentElement::Paragraph(para) = self.elements.remove(idx) {
                Ok(para)
            } else {
                unreachable!()
            }
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Paragraph index {} out of bounds (found {} paragraphs)",
                index, para_count
            )))
        }
    }

    /// Update a paragraph at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the paragraph to update (0-based)
    /// * `text` - New text content
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Old text")?;
    /// doc.update_paragraph(0, "New text")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn update_paragraph(&mut self, index: usize, text: &str) -> Result<()> {
        // Find the index of the nth paragraph
        let mut para_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Paragraph(_)) {
                if para_count == index {
                    element_index = Some(i);
                    break;
                }
                para_count += 1;
            }
        }

        if let Some(idx) = element_index {
            if let DocumentElement::Paragraph(ref mut para) = self.elements[idx] {
                para.set_text(text);
                Ok(())
            } else {
                unreachable!()
            }
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Paragraph index {} out of bounds (found {} paragraphs)",
                index, para_count
            )))
        }
    }

    /// Clear all paragraphs from the document.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Paragraph 1")?;
    /// doc.add_paragraph("Paragraph 2")?;
    /// doc.clear_paragraphs();
    /// assert_eq!(doc.paragraphs().len(), 0);
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_paragraphs(&mut self) {
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Paragraph(_)));
    }

    /// Add a table to the document.
    ///
    /// # Arguments
    ///
    /// * `table` - Table to add
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{MutableDocument, Table};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// let mut table = Table::new();
    /// table.set_name("Table1");
    /// doc.add_table(table)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        self.elements.push(DocumentElement::Table(table));
        Ok(())
    }

    /// Remove a table at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the table to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::{MutableDocument, Table};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_table(Table::new())?;
    /// doc.remove_table(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_table(&mut self, index: usize) -> Result<Table> {
        // Find the index of the nth table
        let mut table_count = 0;
        let mut element_index = None;

        for (i, elem) in self.elements.iter().enumerate() {
            if matches!(elem, DocumentElement::Table(_)) {
                if table_count == index {
                    element_index = Some(i);
                    break;
                }
                table_count += 1;
            }
        }

        if let Some(idx) = element_index {
            if let DocumentElement::Table(table) = self.elements.remove(idx) {
                Ok(table)
            } else {
                unreachable!()
            }
        } else {
            Err(crate::common::Error::InvalidFormat(format!(
                "Table index {} out of bounds (found {} tables)",
                index, table_count
            )))
        }
    }

    /// Clear all tables from the document.
    pub fn clear_tables(&mut self) {
        self.elements
            .retain(|elem| !matches!(elem, DocumentElement::Table(_)));
    }

    /// Clear all content (paragraphs and tables) from the document.
    pub fn clear_content(&mut self) {
        self.elements.clear();
    }

    /// Generate content.xml from the current mutable state.
    fn generate_content_xml(&self) -> String {
        let mut body = String::new();

        // Add elements in their insertion order (paragraphs and tables mixed)
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::odf::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::odf::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
            }
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        let mut meta_fields = String::new();

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = Self::escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = Self::escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        if let Some(ref subject) = self.metadata.subject {
            let escaped_subject = Self::escape_xml(subject);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:subject>{}</dc:subject>"#,
                escaped_subject
            ));
        }

        if let Some(ref description) = self.metadata.description {
            let escaped_description = Self::escape_xml(description);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:description>{}</dc:description>"#,
                escaped_description
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            let escaped_keywords = Self::escape_xml(keywords);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<meta:keyword>{}</meta:keyword>"#,
                escaped_keywords
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date>{}</office:meta></office:document-meta>"#,
            now,
            meta_fields
        )
    }

    /// Escape XML special characters.
    fn escape_xml(text: &str) -> String {
        text.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }

    /// Save the modified document to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// doc.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the document to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::MutableDocument;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut doc = MutableDocument::new();
    /// doc.add_paragraph("Hello!")?;
    /// let bytes = doc.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml (preserved or default)
        let default_styles = OdfStructure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml (regenerated with current metadata)
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        writer.finish_to_bytes()
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}
