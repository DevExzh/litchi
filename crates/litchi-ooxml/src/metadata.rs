use crate::common::xml::decode_xml_reference;
use crate::error::{OoxmlError, Result};
use chrono::{DateTime, Utc};
/// OOXML core properties/metadata extraction.
///
/// This module provides functionality to extract document metadata from
/// Office Open XML (OOXML) documents, including Word (.docx), Excel (.xlsx),
/// and PowerPoint (.pptx) files.
///
/// Core properties are stored in the "docProps/core.xml" part of OOXML packages
/// and contain standard document metadata like title, author, creation date, etc.
use litchi_core::Metadata;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const STRICT_CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/package/metadata/core-properties";
const DUBLIN_CORE_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
const DUBLIN_CORE_TERMS_NAMESPACE: &[u8] = b"http://purl.org/dc/terms/";

#[derive(Clone, Copy)]
enum CoreProperty {
    Title,
    Subject,
    Author,
    Keywords,
    Description,
    LastModifiedBy,
    Revision,
    Category,
    ContentStatus,
    Created,
    Modified,
    LastPrinted,
}

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
fn find_core_properties_part(package: &OpcPackage) -> Result<&dyn litchi_opc::part::Part> {
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
    let mut reader = NsReader::from_reader(xml.as_bytes());

    let mut metadata = Metadata::default();

    loop {
        let property = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| OoxmlError::Xml(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(element) => core_property(&namespace, element.name()),
                Event::Eof => break,
                _ => None,
            }
        };
        let Some(property) = property else {
            continue;
        };
        let Some(text) = read_text_element(&mut reader)? else {
            continue;
        };

        match property {
            CoreProperty::Title => metadata.title = Some(text),
            CoreProperty::Subject => metadata.subject = Some(text),
            CoreProperty::Author => metadata.author = Some(text),
            CoreProperty::Keywords => metadata.keywords = Some(text),
            CoreProperty::Description => metadata.description = Some(text),
            CoreProperty::LastModifiedBy => metadata.last_modified_by = Some(text),
            CoreProperty::Revision => {
                if let Ok(revision) = text.parse::<u32>() {
                    metadata.revision = Some(revision.to_string());
                }
            },
            CoreProperty::Category => metadata.category = Some(text),
            CoreProperty::ContentStatus => metadata.content_status = Some(text),
            CoreProperty::Created => {
                if let Ok(created) = parse_datetime(&text) {
                    metadata.created = Some(created);
                }
            },
            CoreProperty::Modified => {
                if let Ok(modified) = parse_datetime(&text) {
                    metadata.modified = Some(modified);
                }
            },
            CoreProperty::LastPrinted => {
                if let Ok(last_printed) = parse_datetime(&text) {
                    metadata.last_printed_time = Some(last_printed);
                }
            },
        }
    }

    Ok(metadata)
}

fn core_property(namespace: &ResolveResult<'_>, name: QName<'_>) -> Option<CoreProperty> {
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return None;
    };
    let local_name = name.local_name();
    let local_name = local_name.as_ref();

    if *namespace == DUBLIN_CORE_NAMESPACE {
        return match local_name {
            b"title" => Some(CoreProperty::Title),
            b"subject" => Some(CoreProperty::Subject),
            b"creator" | b"author" => Some(CoreProperty::Author),
            b"description" => Some(CoreProperty::Description),
            _ => None,
        };
    }
    if *namespace == DUBLIN_CORE_TERMS_NAMESPACE {
        return match local_name {
            b"created" => Some(CoreProperty::Created),
            b"modified" => Some(CoreProperty::Modified),
            _ => None,
        };
    }
    if *namespace != CORE_PROPERTIES_NAMESPACE && *namespace != STRICT_CORE_PROPERTIES_NAMESPACE {
        return None;
    }
    match local_name {
        // Some producers use the core-properties namespace for fields that the standard
        // places in Dublin Core. Retain that compatibility without depending on prefixes.
        b"title" => Some(CoreProperty::Title),
        b"subject" => Some(CoreProperty::Subject),
        b"creator" | b"author" => Some(CoreProperty::Author),
        b"keywords" => Some(CoreProperty::Keywords),
        b"description" | b"comment" => Some(CoreProperty::Description),
        b"lastModifiedBy" => Some(CoreProperty::LastModifiedBy),
        b"revision" => Some(CoreProperty::Revision),
        b"category" => Some(CoreProperty::Category),
        b"contentStatus" => Some(CoreProperty::ContentStatus),
        b"created" => Some(CoreProperty::Created),
        b"modified" => Some(CoreProperty::Modified),
        b"lastPrinted" => Some(CoreProperty::LastPrinted),
        _ => None,
    }
}

/// Read the text content of an XML element.
fn read_text_element(reader: &mut NsReader<&[u8]>) -> Result<Option<String>> {
    let mut text = String::new();
    let mut depth = 1usize;

    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("core properties nesting is too deep".to_string())
                })?;
            },
            Ok(Event::Text(e)) => {
                let decoded = e
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                text.push_str(
                    &quick_xml::escape::unescape(&decoded)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Ok(Event::CData(e)) => {
                text.push_str(
                    &e.xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| OoxmlError::Xml(error.to_string()))?,
                );
            },
            Ok(Event::GeneralRef(reference)) => {
                text.push_str(&decode_xml_reference(&reference)?);
            },
            Ok(Event::End(_)) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("invalid core properties nesting".to_string())
                })?;
                if depth == 0 {
                    break;
                }
            },
            Ok(Event::Eof) => {
                return Err(OoxmlError::InvalidFormat(
                    "unterminated core property element".to_string(),
                ));
            },
            Err(e) => return Err(OoxmlError::Xml(format!("XML parsing error: {}", e))),
            _ => {},
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
    fn extracts_aliased_core_properties_and_decodes_xml_text() {
        use litchi_opc::part::BlobPart;

        let xml = br#"<props xmlns:core="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
            xmlns:d="http://purl.org/dc/elements/1.1/"
            xmlns:terms="http://purl.org/dc/terms/"
            xmlns:false="urn:not-core-properties">
            <d:title>R&amp;D <![CDATA[<notes>]]></d:title>
            <d:creator>Alice</d:creator>
            <core:revision>7</core:revision>
            <terms:created>2023-10-10T14:30:00Z</terms:created>
            <false:title>ignored</false:title>
        </props>"#;
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/unusual/core-properties.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_string(),
            xml.to_vec(),
        )));

        let metadata = extract_metadata(&package).unwrap();
        assert_eq!(metadata.title.as_deref(), Some("R&D <notes>"));
        assert_eq!(metadata.author.as_deref(), Some("Alice"));
        assert_eq!(metadata.revision.as_deref(), Some("7"));
        assert_eq!(metadata.created.unwrap().year(), 2023);
    }

    #[test]
    fn extracts_strict_core_properties_and_rejects_truncation() {
        let strict = r#"<s:coreProperties xmlns:s="http://purl.oclc.org/ooxml/package/metadata/core-properties"><s:category>Strict</s:category></s:coreProperties>"#;
        assert_eq!(
            parse_core_properties_xml(strict)
                .unwrap()
                .category
                .as_deref(),
            Some("Strict")
        );
        assert!(parse_core_properties_xml(
            r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:keywords>broken"#
        )
        .is_err());
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
