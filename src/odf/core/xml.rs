//! XML parsing utilities for ODF files.
//!
//! This module provides common XML parsing functionality used across
//! different ODF document types.

use crate::common::{Error, Result};

/// XML content parser for ODF parts
#[derive(Debug)]
pub struct XmlPart {
    content: String,
}

impl XmlPart {
    /// Parse XML content from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let content = String::from_utf8(bytes.to_vec())
            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in XML content".to_string()))?;
        Ok(Self { content })
    }

    /// Get the raw XML content
    pub fn content(&self) -> &str {
        &self.content
    }

    /// Get the XML content as bytes (zero-copy)
    #[allow(dead_code)]
    pub fn as_bytes(&self) -> &[u8] {
        self.content.as_bytes()
    }
}

/// Parsed content.xml part
#[derive(Debug)]
pub struct Content {
    xml: XmlPart,
}

impl Content {
    /// Parse content from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let xml = XmlPart::from_bytes(bytes)?;
        Ok(Self { xml })
    }

    /// Get the raw XML content
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }

    /// Extract paragraphs from the document body
    #[allow(dead_code)]
    pub fn extract_paragraphs(&self) -> Result<Vec<crate::odf::elements::text::Paragraph>> {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(self.xml.content());
        let mut buf = Vec::new();
        let mut paragraphs = Vec::new();
        let mut in_body = false;
        let mut in_paragraph = false;
        let mut current_para_text = String::new();
        let mut is_current_heading = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let name_bytes = name.as_ref();

                    // Check if we're entering the body
                    if name_bytes == b"office:body" {
                        in_body = true;
                    }

                    // If we're in the body, check for text elements
                    if in_body && (name_bytes == b"text:p" || name_bytes == b"text:h") {
                        in_paragraph = true;
                        is_current_heading = name_bytes == b"text:h";
                        current_para_text.clear();
                    }
                }
                Ok(Event::Text(ref t)) => {
                    if in_paragraph {
                        let text_content = String::from_utf8(t.to_vec())
                            .unwrap_or_default();
                        current_para_text.push_str(&text_content);
                    }
                }
                Ok(Event::End(ref e)) => {
                    // Copy the name bytes to avoid lifetime issues
                    let name_bytes = e.name().as_ref().to_vec();

                    if name_bytes == b"office:body" {
                        in_body = false;
                    }

                    // Check if we're ending a paragraph element
                    if in_paragraph {
                        let is_ending_para = (is_current_heading && name_bytes == b"text:h") ||
                                           (!is_current_heading && name_bytes == b"text:p");

                        if is_ending_para {
                            in_paragraph = false;
                            let trimmed_text = current_para_text.trim().to_string();
                            if !trimmed_text.is_empty() {
                                let mut para = crate::odf::elements::text::Paragraph::new();
                                para.set_text(&trimmed_text);
                                if is_current_heading {
                                    // For headings, we could set a style or attribute here
                                    // For now, we'll just use regular paragraphs
                                }
                                paragraphs.push(para);
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                _ => {}
            }
            buf.clear();
        }

        Ok(paragraphs)
    }
}


/// Parsed styles.xml part
#[derive(Debug)]
pub struct Styles {
    xml: XmlPart,
}

impl Styles {
    /// Parse styles from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let xml = XmlPart::from_bytes(bytes)?;
        Ok(Self { xml })
    }

    /// Get the raw XML content
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }
}

/// Parsed meta.xml part
#[derive(Debug)]
pub struct Meta {
    xml: XmlPart,
}

impl Meta {
    /// Parse meta from bytes
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let xml = XmlPart::from_bytes(bytes)?;
        Ok(Self { xml })
    }

    /// Get the raw XML content
    #[allow(dead_code)]
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }

    /// Extract basic metadata
    pub fn extract_metadata(&self) -> crate::common::Metadata {
        // Parse ODF metadata from meta.xml content
        match crate::odf::core::metadata::OdfMetadata::from_xml(self.xml.content()) {
            Ok(odf_meta) => odf_meta.into(),
            Err(_) => crate::common::Metadata::default(), // Fall back to default on parse error
        }
    }
}


