//! OpenDocument Text document builder.
//!
//! This module provides a builder pattern for creating new ODT documents from scratch.

use crate::common::{Metadata, Result, xml::escape_xml};
use crate::odf::core::PackageWriter;
use crate::odf::elements::table::Table;
use crate::odf::elements::text::{Heading, List, ListItem, Paragraph, Span};
use std::path::Path;

/// Builder for creating new ODT documents.
///
/// This builder allows you to create ODT documents programmatically by adding
/// paragraphs, tables, and other elements, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi::odf::DocumentBuilder;
///
/// # fn main() -> litchi::Result<()> {
/// let mut builder = DocumentBuilder::new();
/// builder.add_paragraph("Hello, World!")?;
/// builder.add_paragraph("This is a new document.")?;
/// builder.save("document.odt")?;
/// # Ok(())
/// # }
/// ```
/// Document element - can be paragraph, heading, table, or list
#[derive(Debug, Clone)]
enum DocumentElement {
    Paragraph(Paragraph),
    Heading(Heading),
    Table(Table),
    List(List),
}

pub struct DocumentBuilder {
    elements: Vec<DocumentElement>,
    metadata: Metadata,
}

impl Default for DocumentBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl DocumentBuilder {
    /// Create a new document builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    ///
    /// let builder = DocumentBuilder::new();
    /// ```
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            metadata: Metadata::default(),
        }
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    /// use litchi::common::Metadata;
    ///
    /// let mut builder = DocumentBuilder::new();
    /// let mut metadata = Metadata::default();
    /// metadata.title = Some("My Document".to_string());
    /// metadata.author = Some("John Doe".to_string());
    /// builder.set_metadata(metadata);
    /// ```
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Add a paragraph with text
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the paragraph
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph(&mut self, text: &str) -> Result<&mut Self> {
        let mut para = Paragraph::new();
        para.set_text(text);
        self.elements.push(DocumentElement::Paragraph(para));
        Ok(self)
    }

    /// Add a heading
    ///
    /// # Arguments
    ///
    /// * `text` - Heading text
    /// * `level` - Heading level (1-6)
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_heading("Chapter 1", 1)?;
    /// builder.add_heading("Section 1.1", 2)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_heading(&mut self, text: &str, level: u8) -> Result<&mut Self> {
        if !(1..=6).contains(&level) {
            return Err(crate::Error::Other(
                "Heading level must be between 1 and 6".to_string(),
            ));
        }
        let mut heading = Heading::new(level);
        heading.set_text(text);
        self.elements.push(DocumentElement::Heading(heading));
        Ok(self)
    }

    /// Add a paragraph with rich text formatting
    ///
    /// # Arguments
    ///
    /// * `spans` - Vector of (text, style_name) tuples for formatted text
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_rich_paragraph(vec![
    ///     ("This is ", None),
    ///     ("bold", Some("Bold")),
    ///     (" and ", None),
    ///     ("italic", Some("Italic")),
    ///     (" text.", None),
    /// ])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_rich_paragraph(&mut self, spans: Vec<(&str, Option<&str>)>) -> Result<&mut Self> {
        let mut para = Paragraph::new();

        for (text, style) in spans {
            let mut span = Span::new();
            span.set_text(text);
            if let Some(style_name) = style {
                span.set_style_name(style_name);
            }
            para.add_span(span);
        }

        self.elements.push(DocumentElement::Paragraph(para));
        Ok(self)
    }

    /// Add a bulleted list
    ///
    /// # Arguments
    ///
    /// * `items` - Vector of list item texts
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_bulleted_list(vec!["Item 1", "Item 2", "Item 3"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_bulleted_list(&mut self, items: Vec<&str>) -> Result<&mut Self> {
        let mut list = List::new();

        for item_text in items {
            let mut item = ListItem::new();
            let mut para = Paragraph::new();
            para.set_text(item_text);
            item.add_paragraph(para);
            list.add_item(item);
        }

        self.elements.push(DocumentElement::List(list));
        Ok(self)
    }

    /// Add a numbered list
    ///
    /// # Arguments
    ///
    /// * `items` - Vector of list item texts
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_numbered_list(vec!["First", "Second", "Third"])?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_numbered_list(&mut self, items: Vec<&str>) -> Result<&mut Self> {
        let mut list = List::new();
        // Set the numbered list style
        list.set_style_name("L1");

        for item_text in items {
            let mut item = ListItem::new();
            let mut para = Paragraph::new();
            para.set_text(item_text);
            item.add_paragraph(para);
            list.add_item(item);
        }

        self.elements.push(DocumentElement::List(list));
        Ok(self)
    }

    /// Add a paragraph element
    ///
    /// # Arguments
    ///
    /// * `paragraph` - A `Paragraph` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::{DocumentBuilder, Paragraph};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// let mut para = Paragraph::new();
    /// para.set_text("Styled paragraph");
    /// para.set_style_name("Heading1");
    /// builder.add_paragraph_element(para)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_paragraph_element(&mut self, paragraph: Paragraph) -> Result<&mut Self> {
        self.elements.push(DocumentElement::Paragraph(paragraph));
        Ok(self)
    }

    /// Add a heading element
    ///
    /// # Arguments
    ///
    /// * `heading` - A `Heading` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    /// use litchi::odf::elements::text::Heading;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// let mut heading = Heading::new(1);
    /// heading.set_text("Chapter Title");
    /// builder.add_heading_element(heading)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_heading_element(&mut self, heading: Heading) -> Result<&mut Self> {
        self.elements.push(DocumentElement::Heading(heading));
        Ok(self)
    }

    /// Add a list element
    ///
    /// # Arguments
    ///
    /// * `list` - A `List` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::DocumentBuilder;
    /// use litchi::odf::elements::text::{List, ListItem};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// let mut list = List::new();
    /// let mut item = ListItem::new();
    /// item.set_text("First item");
    /// list.add_item(item);
    /// builder.add_list_element(list)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_list_element(&mut self, list: List) -> Result<&mut Self> {
        self.elements.push(DocumentElement::List(list));
        Ok(self)
    }

    /// Add a table to the document
    ///
    /// # Arguments
    ///
    /// * `table` - A `Table` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi::odf::{DocumentBuilder, Table};
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// let mut table = Table::new();
    /// table.set_name("Table1");
    /// builder.add_table(table)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_table(&mut self, table: Table) -> Result<&mut Self> {
        self.elements.push(DocumentElement::Table(table));
        Ok(self)
    }

    /// Generate the content.xml body
    fn generate_content_body(&self) -> String {
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

        // Add all elements in order they were added
        for element in &self.elements {
            match element {
                DocumentElement::Paragraph(para) => {
                    let elem: crate::odf::elements::element::Element = para.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Heading(heading) => {
                    let elem: crate::odf::elements::element::Element = heading.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::Table(table) => {
                    let elem: crate::odf::elements::element::Element = table.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
                DocumentElement::List(list) => {
                    let elem: crate::odf::elements::element::Element = list.clone().into();
                    body.push_str(&elem.to_xml_string());
                },
            }
        }

        body
    }

    /// Generate the complete content.xml
    fn generate_content_xml(&self) -> String {
        let body = self.generate_content_body();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:text>{}</office:text></office:body></office:document-content>"#,
            body
        )
    }

    /// Generate meta.xml with metadata
    fn generate_meta_xml(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();

        let mut meta = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><meta:creation-date>{}</meta:creation-date><dc:date>{}</dc:date>"#,
            now, now
        );

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str(&format!("<dc:title>{}</dc:title>", escape_xml(title)));
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str(&format!("<dc:creator>{}</dc:creator>", escape_xml(author)));
        }

        if let Some(ref subject) = self.metadata.subject {
            meta.push_str(&format!("<dc:subject>{}</dc:subject>", escape_xml(subject)));
        }

        if let Some(ref description) = self.metadata.description {
            meta.push_str(&format!(
                "<dc:description>{}</dc:description>",
                escape_xml(description)
            ));
        }

        if let Some(ref keywords) = self.metadata.keywords {
            meta.push_str(&format!(
                "<meta:keyword>{}</meta:keyword>",
                escape_xml(keywords)
            ));
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Generate styles.xml with list styles
    fn generate_styles_xml(&self) -> String {
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:version="1.3"><office:font-face-decls/><office:styles><!-- Numbered list style --><text:list-style style:name="L1"><text:list-level-style-number text:level="1" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.27cm" fo:text-indent="-0.635cm" fo:margin-left="1.27cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="2" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1.905cm" fo:text-indent="-0.635cm" fo:margin-left="1.905cm"/></style:list-level-properties></text:list-level-style-number><text:list-level-style-number text:level="3" text:style-name="Numbering_20_Symbols" style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="2.54cm" fo:text-indent="-0.635cm" fo:margin-left="2.54cm"/></style:list-level-properties></text:list-level-style-number></text:list-style></office:styles><office:automatic-styles/><office:master-styles/></office:document-styles>"#.to_string()
    }

    /// Build the document and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn build(self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype("application/vnd.oasis.opendocument.text")?;

        // Add content.xml
        let content_xml = self.generate_content_xml();
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml with list styles
        let styles_xml = self.generate_styles_xml();
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml
        let meta_xml = self.generate_meta_xml();
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        // Finish and return bytes
        writer.finish_to_bytes()
    }

    /// Build and save the document to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODT file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi::odf::DocumentBuilder;
    ///
    /// # fn main() -> litchi::Result<()> {
    /// let mut builder = DocumentBuilder::new();
    /// builder.add_paragraph("Hello, World!")?;
    /// builder.save("output.odt")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn test_document_builder_new() {
        let builder = DocumentBuilder::new();
        assert!(builder.elements.is_empty());
    }

    #[test]
    fn test_document_builder_default() {
        let builder: DocumentBuilder = Default::default();
        assert!(builder.elements.is_empty());
    }

    #[test]
    fn test_add_paragraph() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Hello, World!").unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_heading() {
        let mut builder = DocumentBuilder::new();
        builder.add_heading("Chapter 1", 1).unwrap();
        builder.add_heading("Section 1.1", 2).unwrap();
        assert_eq!(builder.elements.len(), 2);
    }

    #[test]
    fn test_add_heading_invalid_level() {
        let mut builder = DocumentBuilder::new();
        let result = builder.add_heading("Invalid", 0);
        assert!(result.is_err());

        let result = builder.add_heading("Invalid", 7);
        assert!(result.is_err());
    }

    #[test]
    fn test_add_rich_paragraph() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_rich_paragraph(vec![
                ("This is ", None),
                ("bold", Some("Bold")),
                (" text.", None),
            ])
            .unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_bulleted_list() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_bulleted_list(vec!["Item 1", "Item 2", "Item 3"])
            .unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_numbered_list() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_numbered_list(vec!["First", "Second", "Third"])
            .unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_paragraph_element() {
        let mut builder = DocumentBuilder::new();
        let mut para = Paragraph::new();
        para.set_text("Custom paragraph");
        builder.add_paragraph_element(para).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_heading_element() {
        let mut builder = DocumentBuilder::new();
        let mut heading = Heading::new(1);
        heading.set_text("Custom heading");
        builder.add_heading_element(heading).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_list_element() {
        let mut builder = DocumentBuilder::new();
        let mut list = List::new();
        let mut item = ListItem::new();
        item.set_text("Item");
        list.add_item(item);
        builder.add_list_element(list).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_add_table() {
        let mut builder = DocumentBuilder::new();
        let mut table = Table::new();
        table.set_name("Table1");
        builder.add_table(table).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }

    #[test]
    fn test_set_metadata() {
        let mut builder = DocumentBuilder::new();
        let mut metadata = Metadata::default();
        metadata.title = Some("Test Title".to_string());
        metadata.author = Some("Test Author".to_string());
        metadata.subject = Some("Test Subject".to_string());
        metadata.description = Some("Test Description".to_string());
        metadata.keywords = Some("test, keywords".to_string());
        builder.set_metadata(metadata);

        assert_eq!(builder.metadata.title, Some("Test Title".to_string()));
        assert_eq!(builder.metadata.author, Some("Test Author".to_string()));
    }

    #[test]
    fn test_generate_content_body() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Paragraph 1").unwrap();
        builder.add_heading("Heading", 1).unwrap();

        let body = builder.generate_content_body();
        assert!(body.contains("Paragraph 1"));
        assert!(body.contains("Heading"));
    }

    #[test]
    fn test_generate_content_xml() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test").unwrap();

        let xml = builder.generate_content_xml();
        assert!(xml.starts_with(r#"<?xml version="1.0" encoding="UTF-8"?"#));
        assert!(xml.contains("office:document-content"));
        assert!(xml.contains("office:text"));
        assert!(xml.contains("Test"));
    }

    #[test]
    fn test_generate_meta_xml() {
        let mut builder = DocumentBuilder::new();
        builder.metadata.title = Some("My Title".to_string());
        builder.metadata.author = Some("My Author".to_string());
        builder.metadata.subject = Some("My Subject".to_string());
        builder.metadata.description = Some("My Description".to_string());
        builder.metadata.keywords = Some("my, keywords".to_string());

        let meta_xml = builder.generate_meta_xml();
        assert!(meta_xml.contains("office:document-meta"));
        assert!(meta_xml.contains("Litchi/"));
        assert!(meta_xml.contains("My Title"));
        assert!(meta_xml.contains("My Author"));
        assert!(meta_xml.contains("My Subject"));
        assert!(meta_xml.contains("My Description"));
        assert!(meta_xml.contains("my, keywords"));
    }

    #[test]
    fn test_generate_styles_xml() {
        let builder = DocumentBuilder::new();
        let styles_xml = builder.generate_styles_xml();
        assert!(styles_xml.contains("office:document-styles"));
        assert!(styles_xml.contains("L1")); // Numbered list style
    }

    #[test]
    fn test_build() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test content").unwrap();

        let result = builder.build();
        assert!(result.is_ok());
        let bytes = result.unwrap();
        assert!(!bytes.is_empty());
        // Check it's a valid ZIP (starts with PK)
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn test_save() {
        let dir = tempdir().unwrap();
        let path = dir.path().join("test.odt");

        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test content").unwrap();

        let result = builder.save(&path);
        assert!(result.is_ok());
        assert!(path.exists());

        // Verify the file is a valid ZIP
        let bytes = std::fs::read(&path).unwrap();
        assert_eq!(&bytes[0..2], b"PK");
    }

    #[test]
    fn test_chained_builder_api() {
        let mut builder = DocumentBuilder::new();
        builder
            .add_heading("Title", 1)
            .unwrap()
            .add_paragraph("Introduction")
            .unwrap()
            .add_bulleted_list(vec!["Point 1", "Point 2"])
            .unwrap()
            .add_numbered_list(vec!["Step 1", "Step 2"])
            .unwrap();

        assert_eq!(builder.elements.len(), 4);
    }

    #[test]
    fn test_document_element_clone() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test").unwrap();

        let cloned = builder.elements[0].clone();
        match (&builder.elements[0], &cloned) {
            (DocumentElement::Paragraph(_), DocumentElement::Paragraph(_)) => {},
            _ => panic!("Clone mismatch"),
        }
    }

    #[test]
    fn test_document_element_debug() {
        let mut builder = DocumentBuilder::new();
        builder.add_paragraph("Test").unwrap();

        let debug_str = format!("{:?}", builder.elements[0]);
        assert!(debug_str.contains("Paragraph"));
    }

    #[test]
    fn test_complete_document() {
        let mut builder = DocumentBuilder::new();

        // Set metadata
        let mut metadata = Metadata::default();
        metadata.title = Some("Complete Document".to_string());
        metadata.author = Some("Test Author".to_string());
        builder.set_metadata(metadata);

        // Add various elements
        builder.add_heading("Title", 1).unwrap();
        builder.add_paragraph("This is a paragraph.").unwrap();
        builder
            .add_rich_paragraph(vec![
                ("Normal ", None),
                ("styled", Some("Emphasis")),
                (" text", None),
            ])
            .unwrap();
        builder
            .add_bulleted_list(vec!["Bullet 1", "Bullet 2"])
            .unwrap();
        builder
            .add_numbered_list(vec!["Number 1", "Number 2"])
            .unwrap();

        // Build and verify
        let result = builder.build();
        assert!(result.is_ok());
    }

    #[test]
    fn test_empty_document_build() {
        let builder = DocumentBuilder::new();
        let result = builder.build();
        assert!(result.is_ok());
    }

    #[test]
    fn test_heading_levels() {
        let mut builder = DocumentBuilder::new();
        for level in 1..=6 {
            builder
                .add_heading(&format!("Level {}", level), level)
                .unwrap();
        }
        assert_eq!(builder.elements.len(), 6);
    }

    #[test]
    fn test_list_with_empty_items() {
        let mut builder = DocumentBuilder::new();
        builder.add_bulleted_list(vec![]).unwrap();
        assert_eq!(builder.elements.len(), 1);
    }
}
