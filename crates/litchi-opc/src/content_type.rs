//! Validated OPC content types and `[Content_Types].xml` handling.

use crate::constants::namespace;
use crate::error::{OpcError, Result};
use crate::packuri::PackURI;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Decoder, XmlVersion};
use std::collections::{HashMap, TryReserveError};
use std::fmt;

const MAX_CONTENT_TYPE_ENTRIES: usize = 65_536;
const MAX_CONTENT_TYPES_XML_BYTES: usize = 8 * 1024 * 1024;
const MAX_XML_EVENTS: usize = 1_000_000;
const MAX_XML_DEPTH: usize = 256;

/// A content type conforming to the MIME grammar required by OPC.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ContentType(String);

impl ContentType {
    /// Validate and construct a content type.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_content_type(&value).map_err(|reason| OpcError::InvalidContentType {
            value: value.clone(),
            reason,
        })?;
        Ok(Self(value))
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

fn validate_content_type(value: &str) -> std::result::Result<(), String> {
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
                .and_then(|value| value.strip_suffix('"'))
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

/// Parsed content type mappings used by the package reader.
pub(crate) struct ContentTypeMap {
    defaults: HashMap<String, ContentType>,
    overrides: HashMap<String, (PackURI, ContentType)>,
}

impl ContentTypeMap {
    pub(crate) fn from_xml(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX_CONTENT_TYPES_XML_BYTES {
            return Err(OpcError::InvalidContentTypesManifest(format!(
                "manifest exceeds {MAX_CONTENT_TYPES_XML_BYTES} bytes"
            )));
        }
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

        loop {
            events = events.checked_add(1).ok_or_else(|| {
                OpcError::InvalidContentTypesManifest("XML event count overflow".to_string())
            })?;
            if events > MAX_XML_EVENTS {
                return Err(OpcError::InvalidContentTypesManifest(format!(
                    "manifest exceeds {MAX_XML_EVENTS} XML events"
                )));
            }
            let decoder = reader.decoder();
            let (resolved_namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) => {
                    inspect_element(
                        &mut map,
                        &resolved_namespace,
                        &element,
                        decoder,
                        depth,
                        &mut root_seen,
                        &mut entry_count,
                    )?;
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OpcError::InvalidContentTypesManifest(
                            "XML nesting depth overflow".to_string(),
                        )
                    })?;
                    if depth > MAX_XML_DEPTH {
                        return Err(OpcError::InvalidContentTypesManifest(format!(
                            "XML nesting exceeds {MAX_XML_DEPTH} levels"
                        )));
                    }
                },
                Event::Empty(element) => inspect_element(
                    &mut map,
                    &resolved_namespace,
                    &element,
                    decoder,
                    depth,
                    &mut root_seen,
                    &mut entry_count,
                )?,
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OpcError::InvalidContentTypesManifest(
                            "unmatched closing element".to_string(),
                        )
                    })?;
                },
                Event::Text(text) if !text.as_ref().is_empty() => {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "text is not permitted in the content types manifest".to_string(),
                    ));
                },
                Event::CData(_) | Event::DocType(_) => {
                    return Err(OpcError::InvalidContentTypesManifest(
                        "CDATA and DTDs are not permitted in the content types manifest"
                            .to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
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
        if let Some((_, content_type)) = self.overrides.get(&pack_uri.as_str().to_ascii_lowercase())
        {
            return Ok(content_type.as_str().to_string());
        }

        let extension = pack_uri.ext().to_ascii_lowercase();
        self.defaults
            .get(&extension)
            .map(|content_type| content_type.as_str().to_string())
            .ok_or_else(|| OpcError::ContentTypeNotFound(pack_uri.to_string()))
    }

    fn add_default(&mut self, extension: String, content_type: ContentType) -> Result<()> {
        validate_extension(&extension)?;
        let key = extension.to_ascii_lowercase();
        if self.defaults.contains_key(&key) {
            return Err(OpcError::DuplicateContentTypeDefault(extension));
        }
        self.defaults
            .try_reserve(1)
            .map_err(|source| allocation("OPC default content types", source))?;
        self.defaults.insert(key, content_type);
        Ok(())
    }

    fn add_override(&mut self, partname: PackURI, content_type: ContentType) -> Result<()> {
        let key = partname.as_str().to_ascii_lowercase();
        if let Some((existing, _)) = self.overrides.get(&key) {
            return Err(OpcError::DuplicateContentTypeOverride {
                existing: existing.to_string(),
                candidate: partname.to_string(),
            });
        }
        self.overrides
            .try_reserve(1)
            .map_err(|source| allocation("OPC content type overrides", source))?;
        self.overrides.insert(key, (partname, content_type));
        Ok(())
    }
}

fn inspect_element(
    map: &mut ContentTypeMap,
    resolved_namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    depth: usize,
    root_seen: &mut bool,
    entry_count: &mut usize,
) -> Result<()> {
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
        return Ok(());
    }

    if depth != 1 || !in_content_types_namespace {
        return Err(OpcError::InvalidContentTypesManifest(
            "unexpected element or namespace in content types manifest".to_string(),
        ));
    }
    *entry_count = entry_count.checked_add(1).ok_or_else(|| {
        OpcError::InvalidContentTypesManifest("content type entry count overflow".to_string())
    })?;
    if *entry_count > MAX_CONTENT_TYPE_ENTRIES {
        return Err(OpcError::InvalidContentTypesManifest(format!(
            "content types manifest exceeds {MAX_CONTENT_TYPE_ENTRIES} entries"
        )));
    }

    match element.local_name().as_ref() {
        b"Default" => {
            let (extension, content_type) =
                required_attributes(element, decoder, "Extension", "Default")?;
            map.add_default(extension, ContentType::new(content_type)?)
        },
        b"Override" => {
            let (partname, content_type) =
                required_attributes(element, decoder, "PartName", "Override")?;
            let partname = PackURI::new(&partname).map_err(OpcError::InvalidPackUri)?;
            map.add_override(partname, ContentType::new(content_type)?)
        },
        _ => Err(OpcError::InvalidContentTypesManifest(
            "only Default and Override children are permitted".to_string(),
        )),
    }
}

fn required_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    key_name: &str,
    element_name: &str,
) -> Result<(String, String)> {
    let mut key_value = None;
    let mut content_type = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        let raw_name = attribute.key.as_ref();
        if raw_name == key_name.as_bytes() {
            key_value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
                    .to_string(),
            );
        } else if raw_name == b"ContentType" {
            content_type = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
                    .to_string(),
            );
        } else if raw_name != b"xmlns" && !raw_name.starts_with(b"xmlns:") {
            return Err(OpcError::InvalidContentTypesManifest(format!(
                "unexpected attribute on {element_name}"
            )));
        }
    }

    match (key_value, content_type) {
        (Some(key), Some(content_type)) => Ok((key, content_type)),
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

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

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
            &manifest(r#"<Other Extension="xml" ContentType="application/xml"/>"#),
        ] {
            assert!(matches!(
                ContentTypeMap::from_xml(xml.as_bytes()),
                Err(OpcError::InvalidContentTypesManifest(_))
            ));
        }
    }

    #[test]
    fn rejects_case_equivalent_defaults_and_overrides() {
        let duplicate_defaults = manifest(
            r#"<Default Extension="xml" ContentType="application/xml"/><Default Extension="XML" ContentType="text/xml"/>"#,
        );
        assert!(matches!(
            ContentTypeMap::from_xml(duplicate_defaults.as_bytes()),
            Err(OpcError::DuplicateContentTypeDefault(_))
        ));

        let duplicate_overrides = manifest(
            r#"<Override PartName="/word/document.xml" ContentType="application/xml"/><Override PartName="/WORD/DOCUMENT.XML" ContentType="text/xml"/>"#,
        );
        assert!(matches!(
            ContentTypeMap::from_xml(duplicate_overrides.as_bytes()),
            Err(OpcError::DuplicateContentTypeOverride { .. })
        ));
    }

    #[test]
    fn override_precedes_case_insensitive_extension_default() {
        let xml = manifest(
            r#"<Default Extension="XML" ContentType="application/xml"/><Override PartName="/word/document.XML" ContentType="text/xml"/>"#,
        );
        let map = ContentTypeMap::from_xml(xml.as_bytes()).unwrap();
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
            ContentTypeMap::from_xml(xml.as_bytes()),
            Err(OpcError::InvalidPackUri(_))
        ));
    }

    #[test]
    fn rejects_oversized_and_event_bomb_manifests() {
        let oversized = vec![b' '; MAX_CONTENT_TYPES_XML_BYTES + 1];
        assert!(matches!(
            ContentTypeMap::from_xml(&oversized),
            Err(OpcError::InvalidContentTypesManifest(_))
        ));

        let mut bomb = format!(r#"<Types xmlns="{NS}">"#);
        for _ in 0..MAX_XML_EVENTS {
            bomb.push_str("<!--x-->");
        }
        bomb.push_str("</Types>");
        assert!(matches!(
            ContentTypeMap::from_xml(bomb.as_bytes()),
            Err(OpcError::InvalidContentTypesManifest(_))
        ));
    }
}
