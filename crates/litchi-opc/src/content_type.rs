//! Validated OPC content types and `[Content_Types].xml` handling.

use crate::constants::namespace;
use crate::error::{OpcError, Result};
use crate::leniency::{CDATA_NODE, ContentTypeDefect, TEXT_NODE, ToleranceReport, bounded_name};
use crate::limits::{ReadLimits, ReadResource};
use crate::packuri::{CONTENT_TYPES_URI, PackURI};
use crate::xml_attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Decoder, XmlVersion};
use std::collections::{HashMap, TryReserveError};
use std::fmt;

/// A content type conforming to the MIME grammar required by OPC.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ContentType(String);

impl ContentType {
    /// Validate and construct a content type.
    ///
    /// # Errors
    /// Returns an error if the value does not conform to the MIME grammar
    /// required by OPC.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let text = value.into();
        validate_content_type(&text).map_err(|reason| OpcError::InvalidContentType {
            value: text.clone(),
            reason,
        })?;
        Ok(Self(text))
    }

    /// Return the serialized content type.
    #[inline]
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for ContentType {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Display for ContentType {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl TryFrom<&str> for ContentType {
    type Error = OpcError;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// Parsed content type mappings used by the package reader.
#[derive(Debug, PartialEq, Eq)]
pub(crate) struct ContentTypeMap {
    defaults: HashMap<String, ContentType>,
    overrides: HashMap<String, (PackURI, ContentType)>,
}

impl ContentTypeMap {
    pub(crate) fn empty() -> Self {
        Self {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
        }
    }

    pub(crate) fn mapping_count(&self) -> usize {
        self.defaults.len().saturating_add(self.overrides.len())
    }

    /// Parse a content-types manifest, discarding the record of any
    /// reference-compatible repair.
    ///
    /// Used where the bytes were already admitted (and their repairs
    /// recorded) at ingress, or where an edit re-derives a manifest: the
    /// mapping semantics are identical to [`Self::from_xml_reporting`].
    pub(crate) fn from_xml(xml: &[u8], limits: ReadLimits) -> Result<Self> {
        Self::from_xml_reporting(xml, limits, &mut ToleranceReport::default())
    }

    /// Parse a content-types manifest, recording every reference-compatible
    /// repair into `report` (see [`crate::leniency`]).
    pub(crate) fn from_xml_reporting(
        xml: &[u8],
        limits: ReadLimits,
        report: &mut ToleranceReport,
    ) -> Result<Self> {
        limits.check(
            ReadResource::ContentTypesBytes,
            xml.len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        let mut map = Self {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
        };
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(true);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut root_seen = false;
        let mut entry_count = 0usize;
        let mut events = 0usize;
        // While an unknown element's subtree is skipped, the depth of that
        // element; everything nested in it is ignored with it.
        let mut skip_from: Option<usize> = None;
        // The key of the open mapping element, for locating defects inside it.
        let mut open_mapping: Option<String> = None;

        loop {
            events = checked_increment(events, limits.max_xml_events(), ReadResource::XmlEvents)?;
            let decoder = reader.decoder();
            let (resolved_namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) => {
                    let next_depth =
                        checked_increment(depth, limits.max_xml_depth(), ReadResource::XmlDepth)?;
                    if skip_from.is_none() {
                        match inspect_element(
                            &mut map,
                            &resolved_namespace,
                            &element,
                            decoder,
                            depth,
                            &mut root_seen,
                            &mut entry_count,
                            open_mapping.as_deref(),
                            limits,
                            report,
                        )? {
                            Inspected::Root => {},
                            Inspected::Mapping(key) => open_mapping = Some(key),
                            Inspected::Skipped => skip_from = Some(depth),
                        }
                    }
                    depth = next_depth;
                },
                Event::Empty(element) => {
                    if skip_from.is_none() {
                        let _ = inspect_element(
                            &mut map,
                            &resolved_namespace,
                            &element,
                            decoder,
                            depth,
                            &mut root_seen,
                            &mut entry_count,
                            open_mapping.as_deref(),
                            limits,
                            report,
                        )?;
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OpcError::InvalidContentTypesManifest(
                            "unmatched closing element".to_string(),
                        )
                    })?;
                    if skip_from == Some(depth) {
                        skip_from = None;
                    } else if skip_from.is_none() && depth == 1 {
                        open_mapping = None;
                    }
                },
                // Character data is ignored by Apache POI and LibreOffice
                // alike (LENIENCY_AUDIT 8.11); it is recorded, not stored.
                // Whitespace-only text never reaches here (`trim_text`).
                Event::Text(text) if !text.as_ref().is_empty() => {
                    if skip_from.is_none() {
                        report.record(
                            ContentTypeDefect::UnexpectedContent,
                            CONTENT_TYPES_URI,
                            open_mapping.as_deref(),
                            Some(TEXT_NODE),
                        );
                    }
                },
                Event::CData(data) => {
                    if skip_from.is_none() && !data.as_ref().iter().all(u8::is_ascii_whitespace) {
                        report.record(
                            ContentTypeDefect::UnexpectedContent,
                            CONTENT_TYPES_URI,
                            open_mapping.as_deref(),
                            Some(CDATA_NODE),
                        );
                    }
                },
                // A DTD stays fatal: Apache POI disallows DOCTYPE outright,
                // and entity declarations are a parser-security surface.
                Event::DocType(_) => {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "DTDs are not permitted in the content types manifest".to_string(),
                    ));
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
            }
        }

        if !root_seen || depth != 0 {
            return Err(OpcError::InvalidContentTypesManifest(
                "missing or unclosed Types root element".to_string(),
            ));
        }
        Ok(map)
    }

    pub(crate) fn get(&self, pack_uri: &PackURI) -> Result<String> {
        self.lookup(pack_uri)
            .map(|content_type| content_type.as_str().to_string())
            .ok_or_else(|| OpcError::ContentTypeNotFound(pack_uri.to_string()))
    }

    /// Return the parsed content type without allocating. This is used when a
    /// source manifest may be reused after a package edit; the caller can
    /// compare the source declaration to the current part directly.
    pub(crate) fn lookup(&self, pack_uri: &PackURI) -> Option<&ContentType> {
        if let Some((_, content_type)) = self.overrides.get(&pack_uri.as_str().to_ascii_lowercase())
        {
            return Some(content_type);
        }

        let extension = pack_uri.ext().to_ascii_lowercase();
        self.defaults.get(&extension)
    }

    /// Return all explicit source overrides, including their original URI
    /// spelling. Unused defaults are intentionally not exposed: OPC permits
    /// a default declaration to remain in a manifest when no current part uses
    /// its extension.
    pub(crate) fn overrides(&self) -> impl Iterator<Item = (&PackURI, &ContentType)> {
        self.overrides
            .values()
            .map(|(partname, content_type)| (partname, content_type))
    }

    /// Return an explicitly declared Override mapping without consulting a
    /// Default extension mapping. Source-backed topology publication uses
    /// this distinction to reject a conflicting existing override while still
    /// allowing a new override to supersede a default mapping.
    pub(crate) fn override_for(&self, pack_uri: &PackURI) -> Option<&ContentType> {
        self.overrides
            .get(&pack_uri.as_str().to_ascii_lowercase())
            .map(|(_, content_type)| content_type)
    }

    /// Add a Default mapping. ECMA-376 Part 2 §9.1.2.2 allows one Default
    /// per extension as a producer rule; a later duplicate (compared ASCII
    /// case-insensitively) is dropped and recorded, keeping the first mapping
    /// as LibreOffice's first-match lookup does (LENIENCY_AUDIT 8.9).
    fn add_default(
        &mut self,
        extension: String,
        content_type: ContentType,
        report: &mut ToleranceReport,
    ) -> Result<()> {
        validate_extension(&extension)?;
        let key = extension.to_ascii_lowercase();
        if self.defaults.contains_key(&key) {
            report.record(
                ContentTypeDefect::DuplicateDefault,
                CONTENT_TYPES_URI,
                Some(&extension),
                Some(content_type.as_str()),
            );
            return Ok(());
        }
        self.defaults
            .try_reserve(1)
            .map_err(|source| allocation("OPC default content types", source))?;
        self.defaults.insert(key, content_type);
        Ok(())
    }

    /// Add an Override mapping; a later duplicate for an equivalent part
    /// name is dropped and recorded, as for [`Self::add_default`].
    fn add_override(
        &mut self,
        partname: PackURI,
        content_type: ContentType,
        report: &mut ToleranceReport,
    ) -> Result<()> {
        let key = partname.as_str().to_ascii_lowercase();
        if self.overrides.contains_key(&key) {
            report.record(
                ContentTypeDefect::DuplicateOverride,
                CONTENT_TYPES_URI,
                Some(partname.as_str()),
                Some(content_type.as_str()),
            );
            return Ok(());
        }
        self.overrides
            .try_reserve(1)
            .map_err(|source| allocation("OPC content type overrides", source))?;
        self.overrides.insert(key, (partname, content_type));
        Ok(())
    }
}

pub(crate) fn validate_content_type(value: &str) -> std::result::Result<(), String> {
    let mut components = value.split(';');
    let media_type = components.next().unwrap_or_default();
    let Some((type_name, subtype)) = media_type.split_once('/') else {
        return Err("missing type/subtype separator".to_string());
    };
    if !is_token(type_name) || !is_token(subtype) {
        return Err("type and subtype must be non-empty ASCII tokens".to_string());
    }

    for parameter in components {
        let Some((name, raw_value)) = parameter.split_once('=') else {
            return Err("content type parameter is missing '='".to_string());
        };
        if !is_token(name) {
            return Err("content type parameter name is not an ASCII token".to_string());
        }
        let parameter_value = if raw_value.starts_with('"') && raw_value.ends_with('"') {
            raw_value
                .strip_prefix('"')
                .and_then(|quoted| quoted.strip_suffix('"'))
                .unwrap_or_default()
        } else {
            raw_value
        };
        if !is_token(parameter_value)
            || (raw_value.contains('"')
                && !(raw_value.starts_with('"') && raw_value.ends_with('"')))
        {
            return Err("content type parameter value is not an ASCII token".to_string());
        }
    }

    Ok(())
}

fn is_token(value: &str) -> bool {
    !value.is_empty()
        && value.bytes().all(|byte| {
            matches!(byte, 0x21..=0x7e)
                && !matches!(
                    byte,
                    b'(' | b')'
                        | b'<'
                        | b'>'
                        | b'@'
                        | b','
                        | b';'
                        | b':'
                        | b'\\'
                        | b'"'
                        | b'/'
                        | b'['
                        | b']'
                        | b'?'
                        | b'='
                        | b'{'
                        | b'}'
                )
        })
}

/// What [`inspect_element`] found.
enum Inspected {
    /// The `Types` document element.
    Root,
    /// A `Default` or `Override` mapping, with its key.
    Mapping(String),
    /// An unknown element, recorded; its subtree must be skipped.
    Skipped,
}

fn inspect_element(
    map: &mut ContentTypeMap,
    resolved_namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    depth: usize,
    root_seen: &mut bool,
    entry_count: &mut usize,
    open_mapping: Option<&str>,
    limits: ReadLimits,
    report: &mut ToleranceReport,
) -> Result<Inspected> {
    let in_content_types_namespace = matches!(
        resolved_namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == namespace::OPC_CONTENT_TYPES.as_bytes()
    );

    if depth == 0 {
        if *root_seen {
            return Err(OpcError::InvalidContentTypesManifest(
                "multiple root elements".to_string(),
            ));
        }
        *root_seen = true;
        if element.local_name().as_ref() != b"Types" || !in_content_types_namespace {
            return Err(OpcError::InvalidContentTypesManifest(
                "root must be Types in the OPC content-types namespace".to_string(),
            ));
        }
        return Ok(Inspected::Root);
    }

    if let ResolveResult::Unknown(prefix) = resolved_namespace {
        // An unbound prefix breaks namespace well-formedness: still fatal.
        return Err(OpcError::InvalidContentTypesManifest(format!(
            "unbound namespace prefix '{}' in content types manifest",
            String::from_utf8_lossy(prefix)
        )));
    }
    let local_name = element.local_name();
    let is_mapping = matches!(local_name.as_ref(), b"Default" | b"Override");
    if depth != 1 || !in_content_types_namespace || !is_mapping {
        // Annex C closes the content model as a producer rule. Apache POI
        // skips every element it does not query; LibreOffice's SAX handler
        // throws. Skipping cannot invent a mapping, so the element and its
        // subtree are skipped and recorded (LENIENCY_AUDIT 8.11).
        let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
        report.record(
            ContentTypeDefect::UnknownElement,
            CONTENT_TYPES_URI,
            open_mapping,
            Some(bounded_name(&name)),
        );
        return Ok(Inspected::Skipped);
    }
    *entry_count = checked_increment(
        *entry_count,
        limits.max_content_type_mappings(),
        ReadResource::ContentTypeMappings,
    )?;

    if local_name.as_ref() == b"Default" {
        let (extension, content_type, unknown) =
            required_attributes(element, decoder, "Extension", "Default", limits)?;
        record_unknown_attributes(report, &extension, &unknown);
        map.add_default(extension.clone(), ContentType::new(content_type)?, report)?;
        Ok(Inspected::Mapping(extension))
    } else {
        let (partname, content_type, unknown) =
            required_attributes(element, decoder, "PartName", "Override", limits)?;
        record_unknown_attributes(report, &partname, &unknown);
        let part_uri = PackURI::new(partname.clone()).map_err(OpcError::InvalidPackUri)?;
        map.add_override(part_uri, ContentType::new(content_type)?, report)?;
        Ok(Inspected::Mapping(partname))
    }
}

fn record_unknown_attributes(report: &mut ToleranceReport, key: &str, unknown: &[String]) {
    for name in unknown {
        report.record(
            ContentTypeDefect::UnknownAttribute,
            CONTENT_TYPES_URI,
            Some(key),
            Some(bounded_name(name)),
        );
    }
}

/// Read a mapping's key and `ContentType`, returning the qualified names of
/// any other non-namespace attributes for the caller to record.
fn required_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    key_name: &str,
    element_name: &str,
    limits: ReadLimits,
) -> Result<(String, String, Vec<String>)> {
    let mut key_value = None;
    let mut content_type = None;
    let mut unknown = Vec::new();
    for attribute_result in element.checked_attributes() {
        let attribute = attribute_result?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            attribute.value.as_ref().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        let raw_name = attribute.key.as_ref();
        if raw_name == key_name.as_bytes() {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
                .to_string();
            limits.check(
                ReadResource::XmlAttributeBytes,
                value.len() as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            key_value = Some(value);
        } else if raw_name == b"ContentType" {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
                .to_string();
            limits.check(
                ReadResource::XmlAttributeBytes,
                value.len() as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            content_type = Some(value);
        } else if raw_name != b"xmlns" && !raw_name.starts_with(b"xmlns:") {
            // Apache POI and LibreOffice both ignore attributes they do not
            // query (LENIENCY_AUDIT 8.11); the caller records it.
            unknown
                .try_reserve(1)
                .map_err(|source| allocation("OPC unknown content-type attributes", source))?;
            unknown.push(String::from_utf8_lossy(raw_name).into_owned());
        }
    }

    match (key_value, content_type) {
        (Some(key), Some(content_type_value)) => Ok((key, content_type_value, unknown)),
        _ => Err(OpcError::InvalidContentTypesManifest(format!(
            "{element_name} requires {key_name} and ContentType attributes"
        ))),
    }
}

fn validate_extension(extension: &str) -> Result<()> {
    if extension.is_empty()
        || extension.bytes().any(|byte| {
            !byte.is_ascii()
                || byte.is_ascii_control()
                || byte.is_ascii_whitespace()
                || matches!(byte, b'.' | b'/' | b'\\' | b'?' | b'#')
        })
    {
        return Err(OpcError::InvalidContentTypeExtension(extension.to_string()));
    }
    Ok(())
}

fn allocation(resource: &'static str, source: TryReserveError) -> OpcError {
    OpcError::Allocation { resource, source }
}

fn checked_increment(current: usize, maximum: usize, resource: ReadResource) -> Result<usize> {
    let actual = current.checked_add(1).ok_or(OpcError::ReadLimit {
        resource,
        actual: u64::MAX,
        maximum: maximum as u64,
    })?;
    if actual > maximum {
        return Err(OpcError::ReadLimit {
            resource,
            actual: actual as u64,
            maximum: maximum as u64,
        });
    }
    Ok(actual)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;
    use crate::leniency::Defect;

    const NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

    fn manifest(children: &str) -> String {
        format!(r#"<Types xmlns="{NS}">{children}</Types>"#)
    }

    #[test]
    fn validates_mime_grammar_without_whitespace_or_comments() {
        assert!(ContentType::new("application/xml;charset=us-ascii").is_ok());
        assert!(ContentType::new("application/xml;charset=\"us-ascii\"").is_ok());
        for invalid in [
            "application",
            "application/",
            "/xml",
            "application /xml",
            "application/xml; charset=utf-8",
            "application/xml (comment)",
            "application/xml;charset=",
        ] {
            assert!(matches!(
                ContentType::new(invalid),
                Err(OpcError::InvalidContentType { .. })
            ));
        }
    }

    #[test]
    fn requires_the_content_types_root_and_required_attributes() {
        for xml in [
            "<Types/>",
            r#"<Types xmlns="urn:wrong"/>"#,
            &manifest(r#"<Default Extension="xml"/>"#),
            &manifest(r#"<Default ContentType="application/xml"/>"#),
        ] {
            assert!(matches!(
                ContentTypeMap::from_xml(xml.as_bytes(), ReadLimits::default()),
                Err(OpcError::InvalidContentTypesManifest(_))
            ));
        }
    }

    fn recorded(report: &ToleranceReport) -> Vec<(Defect, Option<&str>, Option<&str>)> {
        report
            .iter()
            .map(|tolerated| {
                assert_eq!(tolerated.part(), CONTENT_TYPES_URI);
                (tolerated.defect(), tolerated.item(), tolerated.value())
            })
            .collect()
    }

    /// LENIENCY_AUDIT 8.11: unknown elements (with their subtrees), unknown
    /// mapping attributes, character data, and CDATA are ignored and
    /// recorded; the known mappings around them are read normally.
    #[test]
    fn manifest_shape_deviations_are_ignored_and_recorded() {
        let xml = format!(
            r#"<Types xmlns="{NS}" xmlns:x="urn:vendor">stray<x:Extension><Default Extension="bin" ContentType="application/x-hidden"/></x:Extension><Default Extension="xml" ContentType="application/xml" x:hint="1" Producer="tool"><x:note>inner</x:note></Default><Override PartName="/a.xml" ContentType="text/xml"><![CDATA[data]]></Override><Unknown/><![CDATA[ ]]></Types>"#
        );
        let mut report = ToleranceReport::default();
        let map =
            ContentTypeMap::from_xml_reporting(xml.as_bytes(), ReadLimits::default(), &mut report)
                .unwrap();
        assert_eq!(map.mapping_count(), 2);
        assert!(map.get(&PackURI::new("/b.bin").unwrap()).is_err());
        assert_eq!(
            map.get(&PackURI::new("/a.xml").unwrap()).unwrap(),
            "text/xml"
        );
        assert_eq!(
            recorded(&report),
            vec![
                (
                    ContentTypeDefect::UnexpectedContent.into(),
                    None,
                    Some(TEXT_NODE)
                ),
                (
                    ContentTypeDefect::UnknownElement.into(),
                    None,
                    Some("x:Extension")
                ),
                (
                    ContentTypeDefect::UnknownAttribute.into(),
                    Some("xml"),
                    Some("x:hint")
                ),
                (
                    ContentTypeDefect::UnknownAttribute.into(),
                    Some("xml"),
                    Some("Producer")
                ),
                (
                    ContentTypeDefect::UnknownElement.into(),
                    Some("xml"),
                    Some("x:note")
                ),
                (
                    ContentTypeDefect::UnexpectedContent.into(),
                    Some("/a.xml"),
                    Some(CDATA_NODE)
                ),
                (
                    ContentTypeDefect::UnknownElement.into(),
                    None,
                    Some("Unknown")
                ),
            ]
        );

        // The same manifest read through the reportless entry point yields
        // the same mappings.
        assert_eq!(
            ContentTypeMap::from_xml(xml.as_bytes(), ReadLimits::default()).unwrap(),
            map
        );
    }

    /// Well-formedness, the document element, DTDs, and the depth limit stay
    /// fatal even inside a skipped subtree.
    #[test]
    fn manifest_structure_stays_fatal() {
        for xml in [
            manifest("<x:Unbound/>"),
            manifest(
                r#"<Extension><Default Extension="xml" ContentType="application/xml"></Extension>"#,
            ),
            format!(r#"<!DOCTYPE Types><Types xmlns="{NS}"/>"#),
        ] {
            assert!(
                ContentTypeMap::from_xml(xml.as_bytes(), ReadLimits::default()).is_err(),
                "{xml}"
            );
        }
        let deep = manifest("<a><b><c></c></b></a>");
        assert!(matches!(
            ContentTypeMap::from_xml(
                deep.as_bytes(),
                ReadLimits::builder()
                    .max_xml_depth(3)
                    .unwrap()
                    .build()
                    .unwrap()
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlDepth,
                ..
            })
        ));
    }

    #[test]
    fn case_equivalent_defaults_and_overrides_keep_the_first_and_are_recorded() {
        // LENIENCY_AUDIT 8.9: first-wins, as LibreOffice's first-match lookup.
        let duplicates = manifest(
            r#"<Default Extension="xml" ContentType="application/xml"/><Default Extension="XML" ContentType="text/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.first+xml"/><Override PartName="/WORD/DOCUMENT.XML" ContentType="text/xml"/>"#,
        );
        let mut report = ToleranceReport::default();
        let map = ContentTypeMap::from_xml_reporting(
            duplicates.as_bytes(),
            ReadLimits::default(),
            &mut report,
        )
        .unwrap();
        assert_eq!(map.mapping_count(), 2);
        assert_eq!(
            map.get(&PackURI::new("/word/document.xml").unwrap())
                .unwrap(),
            "application/vnd.first+xml"
        );
        assert_eq!(
            map.get(&PackURI::new("/word/other.xml").unwrap()).unwrap(),
            "application/xml"
        );
        let recorded: Vec<_> = report
            .iter()
            .map(|tolerated| {
                (
                    tolerated.defect(),
                    tolerated.part(),
                    tolerated.item(),
                    tolerated.value(),
                )
            })
            .collect();
        assert_eq!(
            recorded,
            vec![
                (
                    ContentTypeDefect::DuplicateDefault.into(),
                    CONTENT_TYPES_URI,
                    Some("XML"),
                    Some("text/xml"),
                ),
                (
                    ContentTypeDefect::DuplicateOverride.into(),
                    CONTENT_TYPES_URI,
                    Some("/WORD/DOCUMENT.XML"),
                    Some("text/xml"),
                ),
            ]
        );
    }

    #[test]
    fn override_precedes_case_insensitive_extension_default() {
        let xml = manifest(
            r#"<Default Extension="XML" ContentType="application/xml"/><Override PartName="/word/document.XML" ContentType="text/xml"/>"#,
        );
        let map = ContentTypeMap::from_xml(xml.as_bytes(), ReadLimits::default()).unwrap();
        assert_eq!(
            map.get(&PackURI::new("/custom/data.xml").unwrap()).unwrap(),
            "application/xml"
        );
        assert_eq!(
            map.get(&PackURI::new("/word/document.XML").unwrap())
                .unwrap(),
            "text/xml"
        );
    }

    #[test]
    fn rejects_invalid_override_part_name() {
        let xml =
            manifest(r#"<Override PartName="word/document.xml" ContentType="application/xml"/>"#);
        assert!(matches!(
            ContentTypeMap::from_xml(xml.as_bytes(), ReadLimits::default()),
            Err(OpcError::InvalidPackUri(_))
        ));
    }

    #[test]
    fn honors_content_type_resource_limits_at_exact_boundaries() {
        let valid = manifest(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        let exact = ReadLimits::builder()
            .max_content_types_bytes(valid.len())
            .unwrap()
            .max_content_type_mappings(1)
            .unwrap()
            .max_xml_events(4)
            .unwrap()
            .max_xml_depth(1)
            .unwrap()
            .max_xml_attribute_bytes("application/xml".len())
            .unwrap()
            .max_relationship_target_bytes("application/xml".len())
            .unwrap()
            .build()
            .unwrap();
        assert!(ContentTypeMap::from_xml(valid.as_bytes(), exact).is_ok());

        let oversized = vec![b' '; valid.len() + 1];
        assert!(matches!(
            ContentTypeMap::from_xml(&oversized, exact),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ContentTypesBytes,
                ..
            })
        ));

        let two_mappings = manifest(
            r#"<Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        );
        assert!(matches!(
            ContentTypeMap::from_xml(
                two_mappings.as_bytes(),
                ReadLimits::builder()
                    .max_content_type_mappings(1)
                    .unwrap()
                    .build()
                    .unwrap()
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ContentTypeMappings,
                actual: 2,
                maximum: 1,
            })
        ));

        let events = two_mappings;
        assert!(matches!(
            ContentTypeMap::from_xml(
                events.as_bytes(),
                ReadLimits::builder()
                    .max_xml_events(4)
                    .unwrap()
                    .build()
                    .unwrap()
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                ..
            })
        ));

        let nested = manifest("<Nested></Nested>");
        assert!(matches!(
            ContentTypeMap::from_xml(
                nested.as_bytes(),
                ReadLimits::builder()
                    .max_xml_depth(1)
                    .unwrap()
                    .build()
                    .unwrap()
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlDepth,
                actual: 2,
                maximum: 1,
            })
        ));

        let encoded_attribute =
            manifest(r#"<Default Extension="xml" ContentType="application&#x2f;xml"/>"#);
        assert!(matches!(
            ContentTypeMap::from_xml(
                encoded_attribute.as_bytes(),
                ReadLimits::builder()
                    .max_xml_attribute_bytes("application/xml".len())
                    .unwrap()
                    .max_relationship_target_bytes("application/xml".len())
                    .unwrap()
                    .build()
                    .unwrap(),
            ),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlAttributeBytes,
                ..
            })
        ));
    }
}
