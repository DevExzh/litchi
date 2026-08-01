//! Strict OPC core-properties metadata extraction.

use crate::error::{OoxmlError, Result};
use chrono::{DateTime, Utc};
use litchi_core::Metadata;
use litchi_ooxml_common::xml::decode_xml_reference;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

const CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const STRICT_CORE_PROPERTIES_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/package/metadata/core-properties";
const STRICT_CORE_PROPERTIES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
const DUBLIN_CORE_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
const DUBLIN_CORE_TERMS_NAMESPACE: &[u8] = b"http://purl.org/dc/terms/";
const XSI_NAMESPACE: &[u8] = b"http://www.w3.org/2001/XMLSchema-instance";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MARKUP_COMPATIBILITY_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_PROPERTY_TEXT: usize = 1_048_576;

#[derive(Clone, Copy, Debug)]
enum CoreProperty {
    Title,
    Subject,
    Author,
    Keywords,
    Description,
    Identifier,
    Language,
    LastModifiedBy,
    Revision,
    Category,
    ContentStatus,
    ContentType,
    Version,
    Created,
    Modified,
    LastPrinted,
}

impl CoreProperty {
    const COUNT: usize = 16;

    fn index(self) -> usize {
        self as usize
    }

    fn requires_w3cdtf(self) -> bool {
        matches!(self, Self::Created | Self::Modified)
    }
}

/// Extract relationship-selected core metadata from an OOXML package.
///
/// Absence is valid and returns empty metadata. A present relationship or part
/// that violates OPC M4.1-M4.5 fails closed.
pub fn extract_metadata(package: &OpcPackage) -> Result<Metadata> {
    let Some(core_part) = find_core_properties_part(package)? else {
        return Ok(Metadata::default());
    };
    let xml = std::str::from_utf8(core_part.blob())
        .map_err(|error| OoxmlError::Xml(format!("invalid UTF-8 in core properties: {error}")))?;
    parse_core_properties_xml(xml)
}

fn find_core_properties_part(package: &OpcPackage) -> Result<Option<&dyn litchi_opc::part::Part>> {
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::CORE_PROPERTIES | STRICT_CORE_PROPERTIES_RELATIONSHIP
        )
    });
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(OoxmlError::InvalidRelationship(
            "OPC M4.1 permits at most one core-properties relationship".to_string(),
        ));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidRelationship(
            "core-properties relationship has an external target".to_string(),
        ));
    }
    let target = relationship.target_partname().map_err(OoxmlError::Opc)?;
    let part = package.get_part(&target).map_err(|_| {
        OoxmlError::PartNotFound(format!(
            "core-properties relationship target '{}' does not exist",
            target.as_str()
        ))
    })?;
    if part.content_type() != ct::OPC_CORE_PROPERTIES {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::OPC_CORE_PROPERTIES.to_string(),
            got: part.content_type().to_string(),
        });
    }
    Ok(Some(part))
}

fn parse_core_properties_xml(xml: &str) -> Result<Metadata> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let mut metadata = Metadata::default();
    let mut seen = [false; CoreProperty::COUNT];
    let mut root_namespace: Option<Vec<u8>> = None;
    let mut root_closed = false;
    let mut active: Option<(CoreProperty, String)> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| OoxmlError::Xml(format!("core properties XML error: {error}")))?;
        let namespace = bound_namespace(&namespace).map(<[u8]>::to_vec);
        let namespace = namespace.as_deref();
        let event = event.into_owned();
        reject_markup_compatibility(namespace)?;

        match event {
            Event::Start(element) if root_namespace.is_none() => {
                validate_root(namespace, &element, &reader)?;
                root_namespace = Some(namespace.expect("validated root namespace").to_vec());
            },
            Event::Empty(element) if root_namespace.is_none() => {
                validate_root(namespace, &element, &reader)?;
                root_namespace = Some(namespace.expect("validated root namespace").to_vec());
                root_closed = true;
            },
            Event::Start(element) if !root_closed && active.is_none() => {
                let property = parse_property(namespace, element.local_name().as_ref())?;
                mark_seen(&mut seen, property)?;
                validate_property_attributes(&reader, &element, property)?;
                active = Some((property, String::new()));
            },
            Event::Empty(element) if !root_closed && active.is_none() => {
                let property = parse_property(namespace, element.local_name().as_ref())?;
                mark_seen(&mut seen, property)?;
                validate_property_attributes(&reader, &element, property)?;
                apply_property(&mut metadata, property, String::new())?;
            },
            Event::Start(_) | Event::Empty(_) if active.is_some() => {
                return Err(OoxmlError::InvalidFormat(
                    "core property values must not contain child elements".to_string(),
                ));
            },
            Event::Start(_) | Event::Empty(_) => {
                return Err(OoxmlError::InvalidFormat(
                    "coreProperties must be the only document element".to_string(),
                ));
            },
            Event::Text(value) => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                push_property_text(&mut active, &decoded)?;
            },
            Event::CData(value) => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                push_property_text(&mut active, &decoded)?;
            },
            Event::GeneralRef(reference) => {
                push_property_text(&mut active, &decode_xml_reference(&reference)?)?;
            },
            Event::End(element) if active.is_some() => {
                let (property, value) = active.take().expect("checked active property");
                let expected = property_name(property);
                if element.local_name().as_ref() != expected {
                    return Err(OoxmlError::InvalidFormat(
                        "malformed core property element nesting".to_string(),
                    ));
                }
                apply_property(&mut metadata, property, value)?;
            },
            Event::End(element) if !root_closed => {
                if element.local_name().as_ref() != b"coreProperties"
                    || namespace != root_namespace.as_deref()
                {
                    return Err(OoxmlError::InvalidFormat(
                        "malformed coreProperties root element".to_string(),
                    ));
                }
                root_closed = true;
            },
            Event::End(_) => {
                return Err(OoxmlError::InvalidFormat(
                    "unexpected element after coreProperties".to_string(),
                ));
            },
            Event::DocType(_) => {
                return Err(OoxmlError::InvalidFormat(
                    "DTD is not allowed in core properties".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if root_namespace.is_none() || !root_closed || active.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "unterminated or missing coreProperties root element".to_string(),
        ));
    }
    Ok(metadata)
}

fn validate_root(
    namespace: Option<&[u8]>,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    if element.local_name().as_ref() != b"coreProperties"
        || !matches!(
            namespace,
            Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE)
        )
    {
        return Err(OoxmlError::InvalidFormat(
            "expected coreProperties in the OPC core-properties namespace".to_string(),
        ));
    }
    validate_attributes(reader, element, None)
}

fn parse_property(namespace: Option<&[u8]>, local: &[u8]) -> Result<CoreProperty> {
    let property = match (namespace, local) {
        (Some(DUBLIN_CORE_NAMESPACE), b"title") => CoreProperty::Title,
        (Some(DUBLIN_CORE_NAMESPACE), b"subject") => CoreProperty::Subject,
        (Some(DUBLIN_CORE_NAMESPACE), b"creator") => CoreProperty::Author,
        (Some(DUBLIN_CORE_NAMESPACE), b"description") => CoreProperty::Description,
        (Some(DUBLIN_CORE_NAMESPACE), b"identifier") => CoreProperty::Identifier,
        (Some(DUBLIN_CORE_NAMESPACE), b"language") => CoreProperty::Language,
        (Some(DUBLIN_CORE_TERMS_NAMESPACE), b"created") => CoreProperty::Created,
        (Some(DUBLIN_CORE_TERMS_NAMESPACE), b"modified") => CoreProperty::Modified,
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"keywords") => {
            CoreProperty::Keywords
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"lastModifiedBy") => {
            CoreProperty::LastModifiedBy
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"revision") => {
            CoreProperty::Revision
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"category") => {
            CoreProperty::Category
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"contentStatus") => {
            CoreProperty::ContentStatus
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"contentType") => {
            CoreProperty::ContentType
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"version") => {
            CoreProperty::Version
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"lastPrinted") => {
            CoreProperty::LastPrinted
        },
        (Some(DUBLIN_CORE_TERMS_NAMESPACE), _) => {
            return Err(OoxmlError::InvalidFormat(
                "OPC M4.3 permits only dcterms:created and dcterms:modified".to_string(),
            ));
        },
        _ => {
            return Err(OoxmlError::InvalidFormat(format!(
                "unexpected core property element '{}'",
                String::from_utf8_lossy(local)
            )));
        },
    };
    Ok(property)
}

fn validate_property_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    property: CoreProperty,
) -> Result<()> {
    validate_attributes(reader, element, Some(property))
}

fn validate_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    property: Option<CoreProperty>,
) -> Result<()> {
    let mut has_w3cdtf = false;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            OoxmlError::Xml(format!("invalid core property attribute: {error}"))
        })?;
        let raw_key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if matches!(raw_key, b"xmlns") || raw_key.starts_with(b"xmlns:") {
            if value.as_bytes() == MARKUP_COMPATIBILITY_NAMESPACE {
                return Err(OoxmlError::InvalidFormat(
                    "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
                ));
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = bound_namespace(&namespace);
        reject_markup_compatibility(namespace)?;
        if namespace == Some(XML_NAMESPACE) && local.as_ref() == b"lang" {
            return Err(OoxmlError::InvalidFormat(
                "OPC M4.4 forbids xml:lang in core properties".to_string(),
            ));
        }
        if namespace == Some(XSI_NAMESPACE) && local.as_ref() == b"type" {
            let Some(property) = property else {
                return Err(OoxmlError::InvalidFormat(
                    "OPC M4.5 forbids xsi:type on coreProperties".to_string(),
                ));
            };
            if !property.requires_w3cdtf() || has_w3cdtf {
                return Err(OoxmlError::InvalidFormat(
                    "OPC M4.5 permits one xsi:type only on created or modified".to_string(),
                ));
            }
            let (value_namespace, value_local) =
                reader.resolver().resolve_element(QName(value.as_bytes()));
            if bound_namespace(&value_namespace) != Some(DUBLIN_CORE_TERMS_NAMESPACE)
                || value_local.as_ref() != b"W3CDTF"
            {
                return Err(OoxmlError::InvalidFormat(
                    "OPC M4.5 requires xsi:type=\"dcterms:W3CDTF\"".to_string(),
                ));
            }
            has_w3cdtf = true;
            continue;
        }
        return Err(OoxmlError::InvalidFormat(format!(
            "attribute '{}' is not allowed on core properties",
            String::from_utf8_lossy(raw_key)
        )));
    }
    if property.is_some_and(CoreProperty::requires_w3cdtf) && !has_w3cdtf {
        return Err(OoxmlError::InvalidFormat(
            "OPC M4.5 requires xsi:type on created and modified".to_string(),
        ));
    }
    Ok(())
}

fn reject_markup_compatibility(namespace: Option<&[u8]>) -> Result<()> {
    if namespace == Some(MARKUP_COMPATIBILITY_NAMESPACE) {
        return Err(OoxmlError::InvalidFormat(
            "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
        ));
    }
    Ok(())
}

fn mark_seen(seen: &mut [bool; CoreProperty::COUNT], property: CoreProperty) -> Result<()> {
    if std::mem::replace(&mut seen[property.index()], true) {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate core property '{}'",
            String::from_utf8_lossy(property_name(property))
        )));
    }
    Ok(())
}

fn push_property_text(active: &mut Option<(CoreProperty, String)>, value: &str) -> Result<()> {
    let Some((_, text)) = active.as_mut() else {
        if value.trim().is_empty() {
            return Ok(());
        }
        return Err(OoxmlError::InvalidFormat(
            "non-whitespace text outside a core property".to_string(),
        ));
    };
    let length = text
        .len()
        .checked_add(value.len())
        .ok_or_else(|| OoxmlError::InvalidFormat("core property is too large".to_string()))?;
    if length > MAX_PROPERTY_TEXT {
        return Err(OoxmlError::InvalidFormat(
            "core property exceeds the text safety limit".to_string(),
        ));
    }
    text.push_str(value);
    Ok(())
}

fn apply_property(metadata: &mut Metadata, property: CoreProperty, value: String) -> Result<()> {
    match property {
        CoreProperty::Title => metadata.title = Some(value),
        CoreProperty::Subject => metadata.subject = Some(value),
        CoreProperty::Author => metadata.author = Some(value),
        CoreProperty::Keywords => metadata.keywords = Some(value),
        CoreProperty::Description => metadata.description = Some(value),
        CoreProperty::Identifier => metadata.identifier = Some(value),
        CoreProperty::Language => metadata.language = Some(value),
        CoreProperty::LastModifiedBy => metadata.last_modified_by = Some(value),
        CoreProperty::Revision => {
            value.trim().parse::<u32>().map_err(|_| {
                OoxmlError::InvalidFormat(format!("invalid core revision '{value}'"))
            })?;
            metadata.revision = Some(value);
        },
        CoreProperty::Category => metadata.category = Some(value),
        CoreProperty::ContentStatus => metadata.content_status = Some(value),
        CoreProperty::ContentType => metadata.content_type = Some(value),
        CoreProperty::Version => metadata.version = Some(value),
        CoreProperty::Created => metadata.created = Some(parse_datetime(value.trim())?),
        CoreProperty::Modified => metadata.modified = Some(parse_datetime(value.trim())?),
        CoreProperty::LastPrinted => {
            metadata.last_printed_time = Some(parse_datetime(value.trim())?)
        },
    }
    Ok(())
}

fn property_name(property: CoreProperty) -> &'static [u8] {
    match property {
        CoreProperty::Title => b"title",
        CoreProperty::Subject => b"subject",
        CoreProperty::Author => b"creator",
        CoreProperty::Keywords => b"keywords",
        CoreProperty::Description => b"description",
        CoreProperty::Identifier => b"identifier",
        CoreProperty::Language => b"language",
        CoreProperty::LastModifiedBy => b"lastModifiedBy",
        CoreProperty::Revision => b"revision",
        CoreProperty::Category => b"category",
        CoreProperty::ContentStatus => b"contentStatus",
        CoreProperty::ContentType => b"contentType",
        CoreProperty::Version => b"version",
        CoreProperty::Created => b"created",
        CoreProperty::Modified => b"modified",
        CoreProperty::LastPrinted => b"lastPrinted",
    }
}

fn bound_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        _ => None,
    }
}

fn parse_datetime(value: &str) -> Result<DateTime<Utc>> {
    if let Ok(datetime) = DateTime::parse_from_rfc3339(value) {
        return Ok(datetime.with_timezone(&Utc));
    }
    if let Ok(datetime) = chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f") {
        return Ok(DateTime::from_naive_utc_and_offset(datetime, Utc));
    }
    Err(OoxmlError::InvalidFormat(format!(
        "invalid core property datetime '{value}'"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::{Datelike, Timelike};
    use litchi_ooxml_common::DocumentProperties;
    use litchi_opc::{PackURI, part::BlobPart};

    fn package_with_core(path: &str, content_type: &str, xml: &[u8]) -> OpcPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(path).unwrap(),
            content_type.to_string(),
            xml.to_vec(),
        )));
        package.relate_to(path.trim_start_matches('/'), rt::CORE_PROPERTIES);
        package
    }

    fn poi_package(name: &str) -> OpcPackage {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/openxml4j")
            .join(name);
        OpcPackage::from_bytes(&std::fs::read(path).unwrap()).unwrap()
    }

    #[test]
    fn selects_only_relationship_target_and_allows_absence() {
        let valid = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Target</dc:title></cp:coreProperties>"#;
        let decoy = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Decoy</dc:title></cp:coreProperties>"#;
        let mut package = package_with_core("/custom/real.xml", ct::OPC_CORE_PROPERTIES, valid);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_string(),
            decoy.to_vec(),
        )));
        assert_eq!(
            extract_metadata(&package).unwrap().title.as_deref(),
            Some("Target")
        );
        assert!(!extract_metadata(&OpcPackage::new()).unwrap().has_data());
    }

    #[test]
    fn rejects_external_dangling_and_wrong_content_type_relationships() {
        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.test/core.xml", rt::CORE_PROPERTIES);
        assert!(extract_metadata(&external).is_err());

        let mut dangling = OpcPackage::new();
        dangling.relate_to("missing.xml", rt::CORE_PROPERTIES);
        assert!(extract_metadata(&dangling).is_err());

        let wrong = package_with_core(
            "/core.xml",
            ct::WML_DOCUMENT,
            br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>"#,
        );
        assert!(extract_metadata(&wrong).is_err());
    }

    #[test]
    fn parses_all_schema_fields_and_normalizes_offsets() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>T</dc:title><dc:subject>S</dc:subject><dc:creator>A</dc:creator><dc:description>D</dc:description><dc:identifier>I</dc:identifier><dc:language>en-GB</dc:language><cp:keywords>K</cp:keywords><cp:lastModifiedBy>M</cp:lastModifiedBy><cp:revision>7</cp:revision><cp:category>C</cp:category><cp:contentStatus>Draft</cp:contentStatus><cp:contentType>text</cp:contentType><cp:version>2.0</cp:version><dcterms:created xsi:type="dcterms:W3CDTF">2006-10-13T18:06:00.123+03:00</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2007-06-20T07:59:00-13:00</dcterms:modified><cp:lastPrinted>2008-01-02T03:04:05Z</cp:lastPrinted></cp:coreProperties>"#;
        let metadata = parse_core_properties_xml(xml).unwrap();
        assert_eq!(metadata.identifier.as_deref(), Some("I"));
        assert_eq!(metadata.language.as_deref(), Some("en-GB"));
        assert_eq!(metadata.version.as_deref(), Some("2.0"));
        assert_eq!(metadata.content_type.as_deref(), Some("text"));
        assert_eq!(metadata.created.unwrap().hour(), 15);
        assert_eq!(metadata.modified.unwrap().day(), 20);
        assert!(metadata.last_printed_time.is_some());
    }

    #[test]
    fn shared_writer_round_trips_through_strict_parser() {
        let created = DateTime::parse_from_rfc3339("2024-03-04T05:06:07+02:00")
            .unwrap()
            .with_timezone(&Utc);
        let properties = DocumentProperties::new()
            .title("A & B")
            .creator("Writer")
            .identifier("urn:test:42")
            .language("en-US")
            .content_type("application/test")
            .revision(12)
            .version("3.0");
        let properties = DocumentProperties {
            created: Some(created),
            last_printed: Some(created),
            ..properties
        };
        let parsed = parse_core_properties_xml(&properties.to_xml()).unwrap();
        assert_eq!(parsed.title.as_deref(), Some("A & B"));
        assert_eq!(parsed.author.as_deref(), Some("Writer"));
        assert_eq!(parsed.identifier.as_deref(), Some("urn:test:42"));
        assert_eq!(parsed.language.as_deref(), Some("en-US"));
        assert_eq!(parsed.content_type.as_deref(), Some("application/test"));
        assert_eq!(parsed.revision.as_deref(), Some("12"));
        assert_eq!(parsed.version.as_deref(), Some("3.0"));
        assert_eq!(parsed.created, Some(created));
        assert_eq!(parsed.last_printed_time, Some(created));
    }

    #[test]
    fn rejects_wrong_root_duplicates_invalid_values_dtd_and_entities() {
        assert!(parse_core_properties_xml("<props/>").is_err());
        let duplicate = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>A</dc:title><dc:title>B</dc:title></cp:coreProperties>"#;
        assert!(parse_core_properties_xml(duplicate).is_err());
        let bad_revision = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:revision>seven</cp:revision></cp:coreProperties>"#;
        assert!(parse_core_properties_xml(bad_revision).is_err());
        assert!(parse_core_properties_xml("<!DOCTYPE x><x/>").is_err());
        let package = poi_package("CorePropertiesHasEntities.ooxml");
        assert!(extract_metadata(&package).is_err());
    }

    #[test]
    fn matches_poi_success_timezone_and_m4_conformance_fixtures() {
        let success =
            extract_metadata(&poi_package("OPCCompliance_CoreProperties_SUCCESS.docx")).unwrap();
        assert_eq!(success.title.as_deref(), Some("MyTitle"));
        assert_eq!(success.revision.as_deref(), Some("2"));

        let timezone = extract_metadata(&poi_package(
            "OPCCompliance_CoreProperties_AlternateTimezones.docx",
        ))
        .unwrap();
        assert_eq!(timezone.created.unwrap().hour(), 15);
        assert_eq!(timezone.modified.unwrap().hour(), 20);

        for name in [
            "OPCCompliance_CoreProperties_DoNotUseCompatibilityMarkupFAIL.docx",
            "OPCCompliance_CoreProperties_DCTermsNamespaceLimitedUseFAIL.docx",
            "OPCCompliance_CoreProperties_UnauthorizedXMLLangAttributeFAIL.docx",
            "OPCCompliance_CoreProperties_LimitedXSITypeAttribute_NotPresentFAIL.docx",
            "OPCCompliance_CoreProperties_LimitedXSITypeAttribute_PresentWithUnauthorizedValueFAIL.docx",
        ] {
            assert!(
                extract_metadata(&poi_package(name)).is_err(),
                "accepted {name}"
            );
        }
    }

    #[test]
    fn poi_no_core_is_empty_and_multiple_relationships_fail_during_package_load() {
        let empty = extract_metadata(&poi_package("OPCCompliance_NoCoreProperties.xlsx")).unwrap();
        assert!(!empty.has_data());
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../test-data/poi/test-data/openxml4j/OPCCompliance_CoreProperties_OnlyOneCorePropertiesPartFAIL.docx",
        );
        assert!(OpcPackage::from_bytes(&std::fs::read(path).unwrap()).is_err());
    }
}
