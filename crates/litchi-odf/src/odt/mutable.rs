//! Mutable document structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODT documents that allows
//! for in-place modification of content, styles, and metadata.

use crate::core::{OdfStructure, PackageWriter};
use crate::elements::parser::DocumentOrderElement;
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use crate::odt::Document;
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::path::Path;

/// Document element type for tracking insertion order
#[derive(Debug, Clone)]
enum DocumentElement {
    /// A paragraph element
    Paragraph(Paragraph),
    /// A heading element
    Heading(Heading),
    /// A table element
    Table(Table),
    /// A text list element
    List(List),
}

/// A mutable ODT document that supports in-place modifications.
///
/// This struct wraps an ODT document and provides methods to modify its content,
/// including adding, updating, and removing paragraphs, tables, and other elements.
///
/// # Examples
///
/// ```no_run
/// use litchi_odf::{Document, MutableDocument};
///
/// # fn main() -> litchi_core::Result<()> {
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
    /// use litchi_odf::{Document, MutableDocument};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Document::open("document.odt")?;
    /// let mut mutable_doc = MutableDocument::from_document(doc)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_document(doc: Document) -> Result<Self> {
        let source_elements = doc.elements()?;
        let metadata = doc.metadata()?;

        // Get MIME type from package
        let mimetype = "application/vnd.oasis.opendocument.text".to_string();

        // Extract styles XML from the document's package
        let styles_xml = doc
            .get_file("styles.xml")
            .ok()
            .and_then(|bytes| String::from_utf8(bytes).ok());

        let elements = source_elements
            .into_iter()
            .map(|element| match element {
                DocumentOrderElement::Paragraph(paragraph) => DocumentElement::Paragraph(paragraph),
                DocumentOrderElement::Heading(heading) => DocumentElement::Heading(heading),
                DocumentOrderElement::Table(table) => DocumentElement::Table(table),
                DocumentOrderElement::List(list) => DocumentElement::List(list),
            })
            .collect();

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
    /// use litchi_odf::MutableDocument;
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

    /// Get all headings in the document.
    pub fn headings(&self) -> Vec<&Heading> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::Heading(heading) => Some(heading),
                _ => None,
            })
            .collect()
    }

    /// Get all lists in the document.
    pub fn lists(&self) -> Vec<&List> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                DocumentElement::List(list) => Some(list),
                _ => None,
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
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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

    /// Add a heading to the end of the document.
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<()> {
        if !(1..=6).contains(&level) {
            return Err(litchi_core::Error::InvalidFormat(
                "Heading level must be between 1 and 6".to_string(),
            ));
        }

        let mut heading = Heading::new(level);
        heading.set_text(text);
        self.elements.push(DocumentElement::Heading(heading));
        Ok(())
    }

    /// Add an existing list element to the end of the document.
    pub fn add_list(&mut self, list: List) -> Result<()> {
        self.elements.push(DocumentElement::List(list));
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
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
    /// use litchi_odf::{MutableDocument, Table};
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
    /// use litchi_odf::{MutableDocument, Table};
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
            Err(litchi_core::Error::InvalidFormat(format!(
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
        let mut estimated = 256usize;
        estimated += self.elements.len() * 96;
        estimated += self
            .elements
            .iter()
            .map(|e| match e {
                DocumentElement::Paragraph(p) => p.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Heading(h) => h.text().map(|t| t.len()).unwrap_or(0),
                DocumentElement::Table(_) => 256,
                DocumentElement::List(_) => 256,
            })
            .sum::<usize>();
        let mut body = String::with_capacity(estimated);

        // Add elements in their insertion order (paragraphs and tables mixed)
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::elements::element::Element = list.clone().into();
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
        let mut estimated = 64usize;
        estimated += self.metadata.title.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.author.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.subject.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self
            .metadata
            .description
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        estimated += self
            .metadata
            .keywords
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        let mut meta_fields = String::with_capacity(estimated);

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        if let Some(ref subject) = self.metadata.subject {
            let escaped_subject = escape_xml(subject);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:subject>{}</dc:subject>"#,
                escaped_subject
            ));
        }

        if let Some(ref description) = self.metadata.description {
            let escaped_description = escape_xml(description);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:description>{}</dc:description>"#,
                escaped_description
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            let escaped_keywords = escape_xml(keywords);
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

    /// Save the modified document to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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
    /// use litchi_odf::MutableDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::elements::parser::DocumentOrderElement;
    use crate::elements::table::{TableCell, TableRow};
    use crate::elements::text::{ListItem, Paragraph};
    use crate::odt::DocumentBuilder;

    fn source_document() -> Document {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Before table").unwrap();

        let mut table = Table::new();
        table.set_name("Data");
        let mut row = TableRow::new();
        let mut cell = TableCell::new();
        cell.set_text("Cell content");
        row.add_cell(cell);
        table.add_row(row);
        builder.add_table(table).unwrap();

        builder.add_heading("After table", 2).unwrap();
        builder
            .add_bulleted_list(vec!["First item", "Second item"])
            .unwrap();
        builder.add_paragraph("Document end").unwrap();

        Document::from_bytes(builder.build().unwrap()).unwrap()
    }

    fn element_kinds(document: &Document) -> Vec<&'static str> {
        document
            .elements()
            .unwrap()
            .iter()
            .map(|element| match element {
                DocumentOrderElement::Paragraph(_) => "paragraph",
                DocumentOrderElement::Heading(_) => "heading",
                DocumentOrderElement::Table(_) => "table",
                DocumentOrderElement::List(_) => "list",
            })
            .collect()
    }

    #[test]
    fn conversion_preserves_top_level_order_without_nested_paragraph_duplicates() {
        let mutable = MutableDocument::from_document(source_document()).unwrap();

        assert_eq!(mutable.paragraphs().len(), 2);
        assert_eq!(mutable.tables().len(), 1);
        assert_eq!(mutable.headings().len(), 1);
        assert_eq!(mutable.lists().len(), 1);
        assert_eq!(mutable.headings()[0].text().unwrap(), "After table");
        assert_eq!(mutable.headings()[0].level(), Some(2));
        assert_eq!(mutable.lists()[0].items().unwrap().len(), 2);
    }

    #[test]
    fn read_modify_write_keeps_paragraph_table_heading_and_list_order() {
        let mutable = MutableDocument::from_document(source_document()).unwrap();
        let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();

        assert_eq!(
            element_kinds(&round_trip),
            ["paragraph", "table", "heading", "list", "paragraph"]
        );

        let elements = round_trip.elements().unwrap();
        let DocumentOrderElement::Table(table) = &elements[1] else {
            panic!("second element should remain a table");
        };
        assert_eq!(table.name(), Some("Data"));
        assert_eq!(
            table
                .row(0)
                .unwrap()
                .unwrap()
                .cell(0)
                .unwrap()
                .unwrap()
                .text()
                .unwrap(),
            "Cell content"
        );

        let DocumentOrderElement::List(list) = &elements[3] else {
            panic!("fourth element should remain a list");
        };
        let items = list.items().unwrap();
        assert_eq!(items[0].text().unwrap(), "First item");
        assert_eq!(items[1].text().unwrap(), "Second item");
    }

    #[test]
    fn mutable_document_can_add_headings_and_lists() {
        let mut document = MutableDocument::new();
        assert!(document.add_heading("Invalid", 0).is_err());
        document.add_heading("Title", 1).unwrap();

        let mut list = List::new();
        let mut item = ListItem::new();
        let mut paragraph = Paragraph::new();
        paragraph.set_text("Item");
        item.add_paragraph(paragraph);
        list.add_item(item);
        document.add_list(list).unwrap();

        let round_trip = Document::from_bytes(document.to_bytes().unwrap()).unwrap();
        assert_eq!(element_kinds(&round_trip), ["heading", "list"]);
    }
}
