/// OOXML core properties/metadata extraction.
///
/// This module provides functionality to extract document metadata from
/// Office Open XML (OOXML) documents, including Word (.docx), Excel (.xlsx),
/// and PowerPoint (.pptx) files.
///
/// Core properties are stored in the "docProps/core.xml" part of OOXML packages
/// and contain standard document metadata like title, author, creation date, etc.
use crate::common::Metadata;
use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::{OpcPackage, PackURI};
use chrono::{DateTime, Utc};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::io::BufRead;

/// Extract metadata from an OOXML package.
///
/// This function looks for the core properties part in the OOXML package
/// and extracts standard document metadata like title, author, creation date, etc.
///
/// # Arguments
///
/// * `package` - The OOXML package to extract metadata from
///
/// # Returns
///
/// A `Metadata` struct containing the extracted document properties
pub fn extract_metadata(package: &OpcPackage) -> Result<Metadata> {
    // Find the core properties part
    let core_part = find_core_properties_part(package)?;

    // Parse the core properties XML
    let xml_content = std::str::from_utf8(core_part.blob())
        .map_err(|e| OoxmlError::Xml(format!("Invalid UTF-8 in core properties: {}", e)))?;

    parse_core_properties_xml(xml_content)
}

/// Find the core properties part in an OOXML package.
///
/// Core properties are typically located at "/docProps/core.xml" and have
/// the content type "application/vnd.openxmlformats-package.core-properties+xml".
fn find_core_properties_part(package: &OpcPackage) -> Result<&dyn crate::ooxml::opc::part::Part> {
    // Try the standard location first
    let standard_uri = PackURI::new("/docProps/core.xml")
        .map_err(|e| OoxmlError::Other(format!("Invalid core properties URI: {}", e)))?;

    if let Ok(part) = package.get_part(&standard_uri)
        && part.content_type() == ct::OPC_CORE_PROPERTIES
    {
        return Ok(part);
    }

    // Fallback: search through all parts for core properties content type
    for part in package.iter_parts() {
        if part.content_type() == ct::OPC_CORE_PROPERTIES {
            return Ok(part);
        }
    }

    Err(OoxmlError::PartNotFound(
        "Core properties part not found".to_string(),
    ))
}

/// Parse core properties XML and extract metadata.
///
/// The core properties XML follows the Dublin Core metadata standard
/// and OPC-specific extensions.
fn parse_core_properties_xml(xml: &str) -> Result<Metadata> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut metadata = Metadata::default();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.name().as_ref() {
                    b"dc:title" | b"cp:title" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.title = Some(text);
                        }
                    },
                    b"dc:subject" | b"cp:subject" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.subject = Some(text);
                        }
                    },
                    b"dc:creator" | b"cp:creator" | b"dc:author" | b"cp:author" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.author = Some(text);
                        }
                    },
                    b"cp:keywords" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.keywords = Some(text);
                        }
                    },
                    b"dc:description" | b"cp:description" | b"cp:comment" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.description = Some(text);
                        }
                    },
                    b"cp:lastModifiedBy" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.last_modified_by = Some(text);
                        }
                    },
                    b"cp:revision" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)?
                            && let Ok(rev) = text.parse::<u32>()
                        {
                            metadata.revision = Some(rev.to_string());
                        }
                    },
                    b"cp:category" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.category = Some(text);
                        }
                    },
                    b"cp:contentStatus" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)? {
                            metadata.content_status = Some(text);
                        }
                    },
                    b"dcterms:created" | b"cp:created" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)?
                            && let Ok(dt) = parse_datetime(&text)
                        {
                            metadata.created = Some(dt);
                        }
                    },
                    b"dcterms:modified" | b"cp:modified" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)?
                            && let Ok(dt) = parse_datetime(&text)
                        {
                            metadata.modified = Some(dt);
                        }
                    },
                    b"cp:lastPrinted" => {
                        if let Some(text) = read_text_element(&mut reader, &mut buf)?
                            && let Ok(dt) = parse_datetime(&text)
                        {
                            metadata.last_printed_time = Some(dt);
                        }
                    },
                    _ => {
                        // Skip unknown elements
                    },
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML parsing error: {}", e))),
            _ => {
                // Skip other events
            },
        }
        buf.clear();
    }

    Ok(metadata)
}

/// Read the text content of an XML element.
fn read_text_element<B: BufRead>(
    reader: &mut Reader<B>,
    buf: &mut Vec<u8>,
) -> Result<Option<String>> {
    let mut text = String::new();

    loop {
        match reader.read_event_into(buf) {
            Ok(Event::Text(e)) => {
                // Convert bytes to string (XML should be UTF-8)
                let text_content = std::str::from_utf8(e.as_ref()).map_err(|e| {
                    OoxmlError::Xml(format!("Invalid UTF-8 in text content: {}", e))
                })?;
                text.push_str(text_content);
            },
            Ok(Event::End(_)) => break,
            Ok(Event::Eof) => break,
            Err(e) => return Err(OoxmlError::Xml(format!("XML parsing error: {}", e))),
            _ => {
                // Skip other events
            },
        }
    }

    if text.trim().is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}

/// Parse an ISO 8601 datetime string into a DateTime<Utc>.
///
/// Supports formats like:
/// - 2023-10-10T14:30:00Z
/// - 2023-10-10T14:30:00.1234567Z
/// - 2023-10-10T14:30:00
fn parse_datetime(s: &str) -> Result<DateTime<Utc>> {
    // Try parsing with different formats
    if let Ok(dt) = DateTime::parse_from_rfc3339(s) {
        return Ok(dt.with_timezone(&Utc));
    }

    // Try parsing as naive datetime and assume UTC
    if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S%.fZ") {
        return Ok(DateTime::from_naive_utc_and_offset(dt, Utc));
    }

    if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%SZ") {
        return Ok(DateTime::from_naive_utc_and_offset(dt, Utc));
    }

    if let Ok(dt) = chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S") {
        return Ok(DateTime::from_naive_utc_and_offset(dt, Utc));
    }

    Err(OoxmlError::InvalidFormat(format!(
        "Invalid datetime format: {}",
        s
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::Datelike;

    #[test]
    #[ignore] // Requires test file
    fn test_extract_metadata() {
        // This would require a test .docx file with core properties
        // let package = OpcPackage::open("test.docx").unwrap();
        // let metadata = extract_metadata(&package).unwrap();
        // assert!(metadata.title.is_some() || metadata.author.is_some());
    }

    #[test]
    fn test_parse_datetime() {
        // Test RFC3339 format
        let dt = parse_datetime("2023-10-10T14:30:00Z").unwrap();
        assert_eq!(dt.year(), 2023);
        assert_eq!(dt.month(), 10);
        assert_eq!(dt.day(), 10);

        // Test with microseconds
        let dt = parse_datetime("2023-10-10T14:30:00.123456Z").unwrap();
        assert_eq!(dt.year(), 2023);

        // Test without Z
        let dt = parse_datetime("2023-10-10T14:30:00").unwrap();
        assert_eq!(dt.year(), 2023);
    }

    #[test]
    fn test_parse_core_properties_xml() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>Test Document</dc:title>
    <dc:subject>Test Subject</dc:subject>
    <dc:creator>Test Author</dc:creator>
    <cp:keywords>test, document</cp:keywords>
    <dc:description>Test Description</dc:description>
    <cp:lastModifiedBy>Test Modifier</cp:lastModifiedBy>
    <cp:revision>5</cp:revision>
    <cp:category>Test Category</cp:category>
    <dcterms:created>2023-10-10T14:30:00Z</dcterms:created>
    <dcterms:modified>2023-10-10T15:30:00Z</dcterms:modified>
</cp:coreProperties>"#;

        let metadata = parse_core_properties_xml(xml).unwrap();
        assert_eq!(metadata.title, Some("Test Document".to_string()));
        assert_eq!(metadata.subject, Some("Test Subject".to_string()));
        assert_eq!(metadata.author, Some("Test Author".to_string()));
        assert_eq!(metadata.keywords, Some("test, document".to_string()));
        assert_eq!(metadata.description, Some("Test Description".to_string()));
        assert_eq!(metadata.last_modified_by, Some("Test Modifier".to_string()));
        assert_eq!(metadata.revision, Some("5".to_string()));
        assert_eq!(metadata.category, Some("Test Category".to_string()));
        assert!(metadata.created.is_some());
        assert!(metadata.modified.is_some());
    }
}
