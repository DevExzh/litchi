//! XML parsing utilities for ODF files.
//!
//! This module provides common XML parsing functionality used across
//! different ODF document types.

use litchi_core::{Error, Result};

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

    /// Parse the complete format-specific OpenDocument metadata model.
    pub fn odf_metadata(&self) -> Result<crate::core::metadata::Metadata> {
        crate::core::metadata::Metadata::from_xml(self.xml.content())
    }

    /// Extract common metadata while preserving parse failures.
    pub fn try_extract_metadata(&self) -> Result<litchi_core::Metadata> {
        Ok(self.odf_metadata()?.into())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xml_part_from_bytes() {
        let xml = b"<?xml version=\"1.0\"?><root><child>text</child></root>";
        let part = XmlPart::from_bytes(xml).unwrap();
        assert_eq!(
            part.content(),
            "<?xml version=\"1.0\"?><root><child>text</child></root>"
        );
    }

    #[test]
    fn test_xml_part_from_bytes_invalid_utf8() {
        // Invalid UTF-8 sequence
        let invalid = vec![0x80, 0x81, 0x82];
        assert!(XmlPart::from_bytes(&invalid).is_err());
    }

    #[test]
    fn test_xml_part_as_bytes() {
        let xml = b"<root/>";
        let part = XmlPart::from_bytes(xml).unwrap();
        assert_eq!(part.as_bytes(), b"<root/>");
    }

    #[test]
    fn test_content_from_bytes() {
        let xml = b"<?xml version=\"1.0\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><office:body><office:text><text:p>Hello World</text:p></office:text></office:body></office:document-content>";
        let content = Content::from_bytes(xml).unwrap();
        assert!(content.xml_content().contains("Hello World"));
    }

    #[test]
    fn test_styles_from_bytes() {
        let xml = b"<?xml version=\"1.0\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-styles>";
        let styles = Styles::from_bytes(xml).unwrap();
        assert!(styles.xml_content().contains("document-styles"));
    }

    #[test]
    fn test_meta_from_bytes() {
        let xml = b"<?xml version=\"1.0\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"></office:document-meta>";
        let meta = Meta::from_bytes(xml).unwrap();
        assert!(meta.xml_content().contains("document-meta"));
    }

    #[test]
    fn test_meta_extract_metadata_empty() {
        let xml = r#"<?xml version="1.0"?>
        <office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                              xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                              xmlns:dc="http://purl.org/dc/elements/1.1/">
        </office:document-meta>"#;
        let meta = Meta::from_bytes(xml.as_bytes()).unwrap();
        let metadata = meta.try_extract_metadata().unwrap();
        // Default metadata should be returned
        assert!(metadata.title.is_none());
    }
}
