//! ODF metadata parsing functionality.
//!
//! This module provides comprehensive parsing of ODF metadata from meta.xml,
//! including document properties, statistics, and user information.

use crate::common::{Error, Metadata, Result};
use chrono::{DateTime, Utc};
use quick_xml::events::Event;
use std::collections::HashMap;

/// Comprehensive ODF metadata
#[derive(Debug, Clone, Default)]
pub struct OdfMetadata {
    /// Document title
    pub title: Option<String>,
    /// Document description
    pub description: Option<String>,
    /// Document subject
    pub subject: Option<String>,
    /// Document keywords
    pub keywords: Vec<String>,
    /// Document creator/author
    pub creator: Option<String>,
    /// Document language
    pub language: Option<String>,
    /// Creation date
    pub creation_date: Option<String>,
    /// Last modification date
    pub modification_date: Option<String>,
    /// Generator application
    pub generator: Option<String>,
    /// Document statistics
    pub statistics: DocumentStatistics,
    /// Custom properties
    pub custom_properties: HashMap<String, String>,
}

/// Document statistics from metadata
#[derive(Debug, Clone, Default)]
pub struct DocumentStatistics {
    /// Number of pages
    pub page_count: Option<u32>,
    /// Number of paragraphs
    pub paragraph_count: Option<u32>,
    /// Number of words
    pub word_count: Option<u32>,
    /// Number of characters
    pub character_count: Option<u32>,
    /// Number of tables
    pub table_count: Option<u32>,
    /// Number of images
    pub image_count: Option<u32>,
    /// Number of objects
    pub object_count: Option<u32>,
}

impl OdfMetadata {
    /// Parse metadata from meta.xml content
    pub fn from_xml(xml_content: &str) -> Result<Self> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut metadata = OdfMetadata::default();
        let mut current_element = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let name_str = String::from_utf8(name.as_ref().to_vec()).unwrap_or_default();

                    current_element.push(name_str);

                    match name.as_ref() {
                        b"dc:title" => {
                            metadata.title = Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"dc:description" => {
                            metadata.description =
                                Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"dc:subject" => {
                            metadata.subject = Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"meta:keyword" => {
                            if let Some(keyword) =
                                Self::extract_text_content(&mut reader, &mut buf)?
                            {
                                metadata.keywords.push(keyword);
                            }
                        },
                        b"dc:creator" => {
                            metadata.creator = Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"dc:language" => {
                            metadata.language = Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"meta:creation-date" => {
                            metadata.creation_date =
                                Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"dc:date" => {
                            metadata.modification_date =
                                Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"meta:generator" => {
                            metadata.generator = Self::extract_text_content(&mut reader, &mut buf)?;
                        },
                        b"meta:document-statistic" => {
                            metadata.statistics = Self::parse_document_statistics(e)?;
                        },
                        b"meta:user-defined" => {
                            let mut temp_buf = Vec::new();
                            if let Some((key, value)) =
                                Self::parse_user_defined_property(e, &mut reader, &mut temp_buf)?
                            {
                                metadata.custom_properties.insert(key, value);
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::End(ref e)) => {
                    if let Some(last) = current_element.last()
                        && last.as_bytes() == e.name().as_ref()
                    {
                        current_element.pop();
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(Error::InvalidFormat(format!(
                        "XML parsing error in metadata: {}",
                        e
                    )));
                },
                _ => {},
            }
            buf.clear();
        }

        Ok(metadata)
    }

    /// Extract text content from current element
    fn extract_text_content(
        reader: &mut quick_xml::Reader<&[u8]>,
        buf: &mut Vec<u8>,
    ) -> Result<Option<String>> {
        let mut content = String::new();
        let mut depth = 0;

        loop {
            match reader.read_event_into(buf) {
                Ok(Event::Start(_)) => {
                    depth += 1;
                },
                Ok(Event::Text(ref t)) => {
                    if depth == 0 {
                        content.push_str(&String::from_utf8(t.to_vec()).unwrap_or_default());
                    }
                },
                Ok(Event::End(_)) => {
                    if depth == 0 {
                        break;
                    }
                    depth -= 1;
                },
                Ok(Event::Eof) => break,
                _ => {},
            }
        }

        let trimmed = content.trim();
        if trimmed.is_empty() {
            Ok(None)
        } else {
            Ok(Some(trimmed.to_string()))
        }
    }

    /// Parse document statistics
    fn parse_document_statistics(e: &quick_xml::events::BytesStart) -> Result<DocumentStatistics> {
        let mut stats = DocumentStatistics::default();

        for attr_result in e.attributes() {
            let attr = attr_result.map_err(|_| {
                Error::InvalidFormat("Invalid attribute in document statistics".to_string())
            })?;
            let value_str = String::from_utf8(attr.value.to_vec()).map_err(|_| {
                Error::InvalidFormat("Invalid UTF-8 in document statistics".to_string())
            })?;

            if let Ok(value) = value_str.parse::<u32>() {
                match attr.key.as_ref() {
                    b"meta:page-count" => stats.page_count = Some(value),
                    b"meta:paragraph-count" => stats.paragraph_count = Some(value),
                    b"meta:word-count" => stats.word_count = Some(value),
                    b"meta:character-count" => stats.character_count = Some(value),
                    b"meta:table-count" => stats.table_count = Some(value),
                    b"meta:image-count" => stats.image_count = Some(value),
                    b"meta:object-count" => stats.object_count = Some(value),
                    _ => {},
                }
            }
        }

        Ok(stats)
    }

    /// Parse user-defined property
    fn parse_user_defined_property(
        e: &quick_xml::events::BytesStart,
        reader: &mut quick_xml::Reader<&[u8]>,
        buf: &mut Vec<u8>,
    ) -> Result<Option<(String, String)>> {
        let mut name = None;

        // Get property name from attributes
        for attr_result in e.attributes() {
            let attr = attr_result.map_err(|_| {
                Error::InvalidFormat("Invalid attribute in user-defined property".to_string())
            })?;
            if attr.key.as_ref() == b"meta:name" {
                name = Some(String::from_utf8(attr.value.to_vec()).map_err(|_| {
                    Error::InvalidFormat("Invalid UTF-8 in property name".to_string())
                })?);
                break;
            }
        }

        if let Some(name) = name {
            if let Some(value) = Self::extract_text_content(reader, buf)? {
                Ok(Some((name, value)))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }
}

impl OdfMetadata {
    /// Parse a date string into DateTime<Utc>
    fn parse_date(date_str: Option<String>) -> Option<DateTime<Utc>> {
        date_str.and_then(|s| {
            // Try different date formats that ODF might use
            // ISO 8601 format: 2023-10-15T14:30:00Z or 2023-10-15T14:30:00.000Z
            if let Ok(dt) = DateTime::parse_from_rfc3339(&s) {
                Some(dt.into())
            } else if let Ok(dt) = DateTime::parse_from_str(&s, "%Y-%m-%dT%H:%M:%S%.fZ") {
                Some(dt.into())
            } else if let Ok(dt) = DateTime::parse_from_str(&s, "%Y-%m-%dT%H:%M:%SZ") {
                Some(dt.into())
            } else {
                // Try simpler date format
                DateTime::parse_from_str(&s, "%Y-%m-%d")
                    .ok()
                    .map(|dt| dt.into())
            }
        })
    }
}

impl From<OdfMetadata> for Metadata {
    fn from(odf_meta: OdfMetadata) -> Self {
        Metadata {
            title: odf_meta.title,
            author: odf_meta.creator,
            subject: odf_meta.subject,
            keywords: if odf_meta.keywords.is_empty() {
                None
            } else {
                Some(odf_meta.keywords.join(", "))
            },
            description: odf_meta.description,
            created: OdfMetadata::parse_date(odf_meta.creation_date),
            modified: OdfMetadata::parse_date(odf_meta.modification_date),
            page_count: odf_meta.statistics.page_count,
            word_count: odf_meta.statistics.word_count,
            character_count: odf_meta.statistics.character_count,
            application: odf_meta.generator,
            ..Default::default()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_odf_metadata_default() {
        let meta = OdfMetadata::default();
        assert!(meta.title.is_none());
        assert!(meta.description.is_none());
        assert!(meta.subject.is_none());
        assert!(meta.keywords.is_empty());
        assert!(meta.creator.is_none());
        assert!(meta.language.is_none());
        assert!(meta.creation_date.is_none());
        assert!(meta.modification_date.is_none());
        assert!(meta.generator.is_none());
        assert!(meta.custom_properties.is_empty());
    }

    #[test]
    fn test_odf_metadata_from_xml_empty() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                      xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert!(meta.title.is_none());
        assert!(meta.creator.is_none());
    }

    #[test]
    fn test_odf_metadata_from_xml_title() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:title>Test Document</dc:title>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.title, Some("Test Document".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_creator() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:creator>John Doe</dc:creator>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.creator, Some("John Doe".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_description() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:description>This is a test document</dc:description>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(
            meta.description,
            Some("This is a test document".to_string())
        );
    }

    #[test]
    fn test_odf_metadata_from_xml_subject() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:subject>Testing</dc:subject>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.subject, Some("Testing".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_keywords() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:keyword>rust</meta:keyword>
        <meta:keyword>odf</meta:keyword>
        <meta:keyword>testing</meta:keyword>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.keywords, vec!["rust", "odf", "testing"]);
    }

    #[test]
    fn test_odf_metadata_from_xml_language() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/">
    <office:meta>
        <dc:language>en-US</dc:language>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.language, Some("en-US".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_dates() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:creation-date>2024-01-15T10:30:00Z</meta:creation-date>
        <dc:date>2024-03-20T14:45:00Z</dc:date>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.creation_date, Some("2024-01-15T10:30:00Z".to_string()));
        assert_eq!(
            meta.modification_date,
            Some("2024-03-20T14:45:00Z".to_string())
        );
    }

    #[test]
    fn test_odf_metadata_from_xml_generator() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:generator>LibreOffice/7.0</meta:generator>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.generator, Some("LibreOffice/7.0".to_string()));
    }

    #[test]
    fn test_odf_metadata_from_xml_statistics() {
        // Note: The parser handles empty document-statistic elements
        // Statistics are parsed from attributes on the Start event
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:document-statistic meta:page-count="5"
                                 meta:paragraph-count="42"
                                 meta:word-count="350"
                                 meta:character-count="2100"
                                 meta:table-count="3"
                                 meta:image-count="2"
                                 meta:object-count="1"></meta:document-statistic>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        // The statistics parsing happens on Start event with attributes
        assert_eq!(meta.statistics.page_count, Some(5));
        assert_eq!(meta.statistics.paragraph_count, Some(42));
        assert_eq!(meta.statistics.word_count, Some(350));
        assert_eq!(meta.statistics.character_count, Some(2100));
        assert_eq!(meta.statistics.table_count, Some(3));
        assert_eq!(meta.statistics.image_count, Some(2));
        assert_eq!(meta.statistics.object_count, Some(1));
    }

    #[test]
    fn test_odf_metadata_from_xml_user_defined() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <meta:user-defined meta:name="Department">Engineering</meta:user-defined>
        <meta:user-defined meta:name="Project">Alpha</meta:user-defined>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(
            meta.custom_properties.get("Department"),
            Some(&"Engineering".to_string())
        );
        assert_eq!(
            meta.custom_properties.get("Project"),
            Some(&"Alpha".to_string())
        );
    }

    #[test]
    fn test_odf_metadata_from_xml_full() {
        let xml = r#"<?xml version="1.0"?>
<office:document-meta xmlns:dc="http://purl.org/dc/elements/1.1/"
                      xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0">
    <office:meta>
        <dc:title>Full Test Document</dc:title>
        <dc:description>A comprehensive test</dc:description>
        <dc:subject>Testing</dc:subject>
        <dc:creator>Test Author</dc:creator>
        <dc:language>en</dc:language>
        <meta:creation-date>2024-01-01T00:00:00Z</meta:creation-date>
        <dc:date>2024-06-01T00:00:00Z</dc:date>
        <meta:generator>Test Generator</meta:generator>
        <meta:keyword>test</meta:keyword>
        <meta:document-statistic meta:page-count="10"></meta:document-statistic>
    </office:meta>
</office:document-meta>"#;

        let meta = OdfMetadata::from_xml(xml).unwrap();
        assert_eq!(meta.title, Some("Full Test Document".to_string()));
        assert_eq!(meta.description, Some("A comprehensive test".to_string()));
        assert_eq!(meta.subject, Some("Testing".to_string()));
        assert_eq!(meta.creator, Some("Test Author".to_string()));
        assert_eq!(meta.language, Some("en".to_string()));
        assert_eq!(meta.creation_date, Some("2024-01-01T00:00:00Z".to_string()));
        assert_eq!(
            meta.modification_date,
            Some("2024-06-01T00:00:00Z".to_string())
        );
        assert_eq!(meta.generator, Some("Test Generator".to_string()));
        assert_eq!(meta.keywords, vec!["test"]);
        assert_eq!(meta.statistics.page_count, Some(10));
    }

    #[test]
    fn test_document_statistics_default() {
        let stats = DocumentStatistics::default();
        assert!(stats.page_count.is_none());
        assert!(stats.paragraph_count.is_none());
        assert!(stats.word_count.is_none());
        assert!(stats.character_count.is_none());
        assert!(stats.table_count.is_none());
        assert!(stats.image_count.is_none());
        assert!(stats.object_count.is_none());
    }

    #[test]
    fn test_parse_date_iso8601() {
        let date = OdfMetadata::parse_date(Some("2024-03-15T14:30:00Z".to_string()));
        assert!(date.is_some());
    }

    #[test]
    fn test_parse_date_rfc3339() {
        let date = OdfMetadata::parse_date(Some("2024-03-15T00:00:00+00:00".to_string()));
        assert!(date.is_some());
    }

    #[test]
    fn test_parse_date_none() {
        let date = OdfMetadata::parse_date(None);
        assert!(date.is_none());
    }

    #[test]
    fn test_parse_date_invalid() {
        let date = OdfMetadata::parse_date(Some("not-a-date".to_string()));
        assert!(date.is_none());
    }

    #[test]
    fn test_into_metadata_empty() {
        let odf = OdfMetadata::default();
        let meta: Metadata = odf.into();
        assert!(meta.title.is_none());
        assert!(meta.author.is_none());
        assert!(meta.keywords.is_none());
    }

    #[test]
    fn test_into_metadata_with_data() {
        let odf = OdfMetadata {
            title: Some("Title".to_string()),
            creator: Some("Author".to_string()),
            subject: Some("Subject".to_string()),
            keywords: vec!["a".to_string(), "b".to_string()],
            description: Some("Desc".to_string()),
            creation_date: Some("2024-01-01T00:00:00Z".to_string()),
            modification_date: Some("2024-06-01T00:00:00Z".to_string()),
            generator: Some("App".to_string()),
            statistics: DocumentStatistics {
                page_count: Some(5),
                word_count: Some(100),
                character_count: Some(500),
                ..Default::default()
            },
            ..Default::default()
        };

        let meta: Metadata = odf.into();
        assert_eq!(meta.title, Some("Title".to_string()));
        assert_eq!(meta.author, Some("Author".to_string()));
        assert_eq!(meta.subject, Some("Subject".to_string()));
        assert_eq!(meta.keywords, Some("a, b".to_string()));
        assert_eq!(meta.description, Some("Desc".to_string()));
        assert_eq!(meta.page_count, Some(5));
        assert_eq!(meta.word_count, Some(100));
        assert_eq!(meta.character_count, Some(500));
        assert_eq!(meta.application, Some("App".to_string()));
        assert!(meta.created.is_some());
        assert!(meta.modified.is_some());
    }

    #[test]
    fn test_into_metadata_no_keywords() {
        let odf = OdfMetadata {
            keywords: vec![],
            ..Default::default()
        };

        let meta: Metadata = odf.into();
        assert!(meta.keywords.is_none());
    }
}
