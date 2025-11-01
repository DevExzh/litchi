//! Common document properties for OOXML formats.
//!
//! This module provides document metadata properties that are shared across
//! DOCX, XLSX, and PPTX formats.

use chrono::{DateTime, Utc};

/// Document core properties (metadata).
///
/// These properties are stored in the `docProps/core.xml` file in the OPC package.
#[derive(Debug, Clone, Default)]
pub struct DocumentProperties {
    /// Document title
    pub title: Option<String>,
    /// Document subject
    pub subject: Option<String>,
    /// Document creator/author
    pub creator: Option<String>,
    /// Document keywords (comma-separated)
    pub keywords: Option<String>,
    /// Document description
    pub description: Option<String>,
    /// Last modified by
    pub last_modified_by: Option<String>,
    /// Document category
    pub category: Option<String>,
    /// Content status (e.g., "Draft", "Final")
    pub content_status: Option<String>,
    /// Document language
    pub language: Option<String>,
    /// Creation date
    pub created: Option<DateTime<Utc>>,
    /// Last modification date
    pub modified: Option<DateTime<Utc>>,
}

impl DocumentProperties {
    /// Create a new empty document properties.
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the document title.
    pub fn title(mut self, title: &str) -> Self {
        self.title = Some(title.to_string());
        self
    }

    /// Set the document subject.
    pub fn subject(mut self, subject: &str) -> Self {
        self.subject = Some(subject.to_string());
        self
    }

    /// Set the document creator/author.
    pub fn creator(mut self, creator: &str) -> Self {
        self.creator = Some(creator.to_string());
        self
    }

    /// Set the document keywords.
    pub fn keywords(mut self, keywords: &str) -> Self {
        self.keywords = Some(keywords.to_string());
        self
    }

    /// Set the document description.
    pub fn description(mut self, description: &str) -> Self {
        self.description = Some(description.to_string());
        self
    }

    /// Set who last modified the document.
    pub fn last_modified_by(mut self, name: &str) -> Self {
        self.last_modified_by = Some(name.to_string());
        self
    }

    /// Set the document category.
    pub fn category(mut self, category: &str) -> Self {
        self.category = Some(category.to_string());
        self
    }

    /// Set the content status.
    pub fn content_status(mut self, status: &str) -> Self {
        self.content_status = Some(status.to_string());
        self
    }

    /// Set the document language.
    pub fn language(mut self, language: &str) -> Self {
        self.language = Some(language.to_string());
        self
    }

    /// Generate core.xml content for this properties set.
    pub fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#);

        // Title
        if let Some(ref title) = self.title {
            xml.push_str("<dc:title>");
            xml.push_str(&escape_xml(title));
            xml.push_str("</dc:title>");
        }

        // Subject
        if let Some(ref subject) = self.subject {
            xml.push_str("<dc:subject>");
            xml.push_str(&escape_xml(subject));
            xml.push_str("</dc:subject>");
        }

        // Creator
        if let Some(ref creator) = self.creator {
            xml.push_str("<dc:creator>");
            xml.push_str(&escape_xml(creator));
            xml.push_str("</dc:creator>");
        }

        // Keywords
        if let Some(ref keywords) = self.keywords {
            xml.push_str("<cp:keywords>");
            xml.push_str(&escape_xml(keywords));
            xml.push_str("</cp:keywords>");
        }

        // Description
        if let Some(ref description) = self.description {
            xml.push_str("<dc:description>");
            xml.push_str(&escape_xml(description));
            xml.push_str("</dc:description>");
        }

        // Last modified by
        if let Some(ref last_modified_by) = self.last_modified_by {
            xml.push_str("<cp:lastModifiedBy>");
            xml.push_str(&escape_xml(last_modified_by));
            xml.push_str("</cp:lastModifiedBy>");
        }

        // Category
        if let Some(ref category) = self.category {
            xml.push_str("<cp:category>");
            xml.push_str(&escape_xml(category));
            xml.push_str("</cp:category>");
        }

        // Content status
        if let Some(ref status) = self.content_status {
            xml.push_str("<cp:contentStatus>");
            xml.push_str(&escape_xml(status));
            xml.push_str("</cp:contentStatus>");
        }

        // Language
        if let Some(ref language) = self.language {
            xml.push_str("<dc:language>");
            xml.push_str(&escape_xml(language));
            xml.push_str("</dc:language>");
        }

        // Created date
        if let Some(ref created) = self.created {
            xml.push_str("<dcterms:created xsi:type=\"dcterms:W3CDTF\">");
            xml.push_str(&created.to_rfc3339());
            xml.push_str("</dcterms:created>");
        }

        // Modified date
        if let Some(ref modified) = self.modified {
            xml.push_str("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">");
            xml.push_str(&modified.to_rfc3339());
            xml.push_str("</dcterms:modified>");
        }

        xml.push_str("</cp:coreProperties>");
        xml
    }
}

/// Escape XML special characters.
fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_document_properties_builder() {
        let props = DocumentProperties::new()
            .title("Test Document")
            .creator("John Doe")
            .subject("Testing")
            .keywords("test, document, rust");

        assert_eq!(props.title, Some("Test Document".to_string()));
        assert_eq!(props.creator, Some("John Doe".to_string()));
        assert_eq!(props.subject, Some("Testing".to_string()));
        assert_eq!(props.keywords, Some("test, document, rust".to_string()));
    }

    #[test]
    fn test_xml_generation() {
        let props = DocumentProperties::new()
            .title("My Document")
            .creator("Test Author");

        let xml = props.to_xml();
        assert!(xml.contains("<dc:title>My Document</dc:title>"));
        assert!(xml.contains("<dc:creator>Test Author</dc:creator>"));
    }

    #[test]
    fn test_xml_escaping() {
        let props = DocumentProperties::new().title("Test & <Special> \"Characters\"");

        let xml = props.to_xml();
        assert!(xml.contains("&amp;"));
        assert!(xml.contains("&lt;"));
        assert!(xml.contains("&gt;"));
        assert!(xml.contains("&quot;"));
    }
}
