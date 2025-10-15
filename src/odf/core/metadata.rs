//! ODF metadata parsing functionality.
//!
//! This module provides comprehensive parsing of ODF metadata from meta.xml,
//! including document properties, statistics, and user information.

use crate::common::{Error, Result, Metadata};
use std::collections::HashMap;
use quick_xml::events::Event;
use chrono::{DateTime, Utc};

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
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut metadata = OdfMetadata::default();
        let mut current_element = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    let name = e.name();
                    let name_str = String::from_utf8(name.as_ref().to_vec())
                        .unwrap_or_default();

                    current_element.push(name_str.clone());

                    match name.as_ref() {
                        b"dc:title" => {
                            metadata.title = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"dc:description" => {
                            metadata.description = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"dc:subject" => {
                            metadata.subject = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"meta:keyword" => {
                            if let Some(keyword) = Self::extract_text_content(&mut reader, &mut buf)? {
                                metadata.keywords.push(keyword);
                            }
                        }
                        b"dc:creator" => {
                            metadata.creator = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"dc:language" => {
                            metadata.language = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"meta:creation-date" => {
                            metadata.creation_date = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"dc:date" => {
                            metadata.modification_date = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"meta:generator" => {
                            metadata.generator = Self::extract_text_content(&mut reader, &mut buf)?;
                        }
                        b"meta:document-statistic" => {
                            metadata.statistics = Self::parse_document_statistics(e)?;
                        }
                        b"meta:user-defined" => {
                            let mut temp_buf = Vec::new();
                            if let Some((key, value)) = Self::parse_user_defined_property(e, &mut reader, &mut temp_buf)? {
                                metadata.custom_properties.insert(key, value);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(ref e)) => {
                    if let Some(last) = current_element.last()
                        && last.as_bytes() == e.name().as_ref() {
                            current_element.pop();
                        }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::InvalidFormat(format!("XML parsing error in metadata: {}", e))),
                _ => {}
            }
            buf.clear();
        }

        Ok(metadata)
    }

    /// Extract text content from current element
    fn extract_text_content(reader: &mut quick_xml::Reader<&[u8]>, buf: &mut Vec<u8>) -> Result<Option<String>> {
        let mut content = String::new();
        let mut depth = 0;

        loop {
            match reader.read_event_into(buf) {
                Ok(Event::Start(_)) => {
                    depth += 1;
                }
                Ok(Event::Text(ref t)) => {
                    if depth == 0 {
                        content.push_str(&String::from_utf8(t.to_vec())
                            .unwrap_or_default());
                    }
                }
                Ok(Event::End(_)) => {
                    if depth == 0 {
                        break;
                    }
                    depth -= 1;
                }
                Ok(Event::Eof) => break,
                _ => {}
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
            let attr = attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute in document statistics".to_string()))?;
            let value_str = String::from_utf8(attr.value.to_vec())
                .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in document statistics".to_string()))?;

            if let Ok(value) = value_str.parse::<u32>() {
                match attr.key.as_ref() {
                    b"meta:page-count" => stats.page_count = Some(value),
                    b"meta:paragraph-count" => stats.paragraph_count = Some(value),
                    b"meta:word-count" => stats.word_count = Some(value),
                    b"meta:character-count" => stats.character_count = Some(value),
                    b"meta:table-count" => stats.table_count = Some(value),
                    b"meta:image-count" => stats.image_count = Some(value),
                    b"meta:object-count" => stats.object_count = Some(value),
                    _ => {}
                }
            }
        }

        Ok(stats)
    }

    /// Parse user-defined property
    fn parse_user_defined_property(
        e: &quick_xml::events::BytesStart,
        reader: &mut quick_xml::Reader<&[u8]>,
        buf: &mut Vec<u8>
    ) -> Result<Option<(String, String)>> {
        let mut name = None;

        // Get property name from attributes
        for attr_result in e.attributes() {
            let attr = attr_result.map_err(|_| Error::InvalidFormat("Invalid attribute in user-defined property".to_string()))?;
            if attr.key.as_ref() == b"meta:name" {
                name = Some(String::from_utf8(attr.value.to_vec())
                    .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in property name".to_string()))?);
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
                DateTime::parse_from_str(&s, "%Y-%m-%d").ok().map(|dt| dt.into())
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
