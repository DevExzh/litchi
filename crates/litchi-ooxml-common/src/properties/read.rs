//! Strict OPC core-properties metadata extraction.

use super::{
    CORE_PROPERTIES_NAMESPACE, Dialect, Keywords, MAX_PROPERTY_TEXT, MAX_XML_BYTES, MAX_XML_EVENTS,
    Props, STRICT_CORE_PROPERTIES_NAMESPACE, graph,
    keyword::{self, Item},
    time,
};
use crate::{Error, Result};
use litchi_opc::{OpcPackage, SourceBackedPackage};

use crate::xml::decode_xml_reference;
use quick_xml::{
    XmlVersion,
    events::{BytesEnd, BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};

const DUBLIN_CORE_NAMESPACE: &[u8] = b"http://purl.org/dc/elements/1.1/";
const DUBLIN_CORE_TERMS_NAMESPACE: &[u8] = b"http://purl.org/dc/terms/";
const XSI_NAMESPACE: &[u8] = b"http://www.w3.org/2001/XMLSchema-instance";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MARKUP_COMPATIBILITY_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
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
    Version,
    Created,
    Modified,
    LastPrinted,
}

#[derive(Debug)]
enum Active {
    Text(CoreProperty, String),
    Keywords {
        value: Keywords,
        child: Option<keyword::Value>,
        bytes: usize,
    },
}

impl CoreProperty {
    const COUNT: usize = Self::LastPrinted as usize + 1;

    fn index(self) -> usize {
        self as usize
    }

    fn requires_w3cdtf(self) -> bool {
        matches!(self, Self::Created | Self::Modified)
    }
}

/// Extract relationship-selected core metadata from an OOXML package.
///
/// Absence is valid and remains distinguishable from a present empty document.
/// A present relationship or part that violates OPC M4.1-M4.5 fails closed.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn read(package: &OpcPackage) -> Result<Option<Props>> {
    let graph = graph::inspect(package)?;
    let Some(part_name) = graph.part else {
        return Ok(None);
    };
    let core_part = package.get_part(&part_name)?;
    let xml = std::str::from_utf8(core_part.blob())
        .map_err(|error| Error::Xml(format!("invalid UTF-8 in core properties: {error}")))?;
    let (props, dialect) = decode(xml)?;
    if Some(dialect) != graph.dialect {
        return Err(Error::Invalid(
            "core-properties XML namespace does not match its relationship dialect".to_owned(),
        ));
    }
    Ok(Some(props))
}

/// Extract relationship-selected core metadata from a source-backed OPC
/// package.
///
/// The catalog and relationship graph remain metadata-only until the
/// relationship-selected core-properties part is read. This is the
/// positional equivalent of [`read`]: it applies the same strict graph and
/// dialect checks, while retaining the source package's version and
/// execution-policy checks around the selected payload read.
///
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a
/// configured bound, the source changes, execution is cancelled, or an
/// underlying source/package operation fails.
pub fn read_source_backed(package: &SourceBackedPackage) -> Result<Option<Props>> {
    package.check_execution()?;
    package.source_version()?;
    let graph = graph::inspect_source(package)?;
    let Some(part_name) = graph.part else {
        package.check_execution()?;
        package.source_version()?;
        return Ok(None);
    };
    let core_part = package.part(&part_name)?;
    let data = core_part.data()?;
    package.check_execution()?;
    let xml = std::str::from_utf8(data.as_bytes())
        .map_err(|error| Error::Xml(format!("invalid UTF-8 in core properties: {error}")))?;
    let (props, dialect) = decode(xml)?;
    if Some(dialect) != graph.dialect {
        return Err(Error::Invalid(
            "core-properties XML namespace does not match its relationship dialect".to_owned(),
        ));
    }
    package.check_execution()?;
    package.source_version()?;
    Ok(Some(props))
}

pub(super) fn decode(xml: &str) -> Result<(Props, Dialect)> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::Limit {
            resource: "core-properties XML bytes",
            max: MAX_XML_BYTES,
            actual: xml.len(),
        });
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut buffer = Vec::new();
    let mut props = Props::new();
    let mut seen = [false; CoreProperty::COUNT];
    let mut root_namespace: Option<Vec<u8>> = None;
    let mut root_closed = false;
    let mut active: Option<Active> = None;
    let mut events = 0usize;

    loop {
        events = events.checked_add(1).ok_or(Error::Limit {
            resource: "core-properties XML events",
            max: MAX_XML_EVENTS,
            actual: usize::MAX,
        })?;
        if events > MAX_XML_EVENTS {
            return Err(Error::Limit {
                resource: "core-properties XML events",
                max: MAX_XML_EVENTS,
                actual: events,
            });
        }
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::Xml(format!("core properties XML error: {error}")))?;
        let namespace = bound_namespace(&namespace).map(<[u8]>::to_vec);
        let namespace = namespace.as_deref();
        let event = event.into_owned();
        reject_markup_compatibility(namespace)?;

        match event {
            Event::Start(element) if root_namespace.is_none() => {
                validate_root(namespace, &element, &reader)?;
                root_namespace = Some(
                    namespace
                        .ok_or_else(|| Error::Invalid("coreProperties has no namespace".into()))?
                        .to_vec(),
                );
            },
            Event::Empty(element) if root_namespace.is_none() => {
                validate_root(namespace, &element, &reader)?;
                root_namespace = Some(
                    namespace
                        .ok_or_else(|| Error::Invalid("coreProperties has no namespace".into()))?
                        .to_vec(),
                );
                root_closed = true;
            },
            Event::Start(element) if !root_closed && active.is_none() => {
                let property = parse_property(namespace, element.local_name().as_ref())?;
                mark_seen(&mut seen, property)?;
                active = Some(if property == CoreProperty::Keywords {
                    Active::Keywords {
                        value: Keywords {
                            lang: keyword_lang(&reader, &element)?,
                            items: Vec::new(),
                        },
                        child: None,
                        bytes: 0,
                    }
                } else {
                    validate_property_attributes(&reader, &element, property)?;
                    Active::Text(property, String::new())
                });
            },
            Event::Empty(element) if !root_closed && active.is_none() => {
                let property = parse_property(namespace, element.local_name().as_ref())?;
                mark_seen(&mut seen, property)?;
                if property == CoreProperty::Keywords {
                    props.keywords = Some(Keywords {
                        lang: keyword_lang(&reader, &element)?,
                        items: Vec::new(),
                    });
                } else {
                    validate_property_attributes(&reader, &element, property)?;
                    apply_property(&mut props, property, String::new())?;
                }
            },
            Event::Start(element) if active.is_some() => {
                start_keyword_value(
                    &mut active,
                    namespace,
                    root_namespace.as_deref(),
                    &reader,
                    &element,
                    false,
                )?;
            },
            Event::Empty(element) if active.is_some() => {
                start_keyword_value(
                    &mut active,
                    namespace,
                    root_namespace.as_deref(),
                    &reader,
                    &element,
                    true,
                )?;
            },
            Event::Start(_) | Event::Empty(_) => {
                return Err(Error::Invalid(
                    "coreProperties must be the only document element".to_string(),
                ));
            },
            Event::Text(value) => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let decoded = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                push_property_text(&mut active, &decoded)?;
            },
            Event::CData(value) => {
                let decoded = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                push_property_text(&mut active, &decoded)?;
            },
            Event::GeneralRef(reference) => {
                push_property_text(&mut active, &decode_xml_reference(&reference)?)?;
            },
            Event::End(element) if active.is_some() => {
                close_active(
                    &mut active,
                    namespace,
                    root_namespace.as_deref(),
                    &element,
                    &mut props,
                )?;
            },
            Event::End(element) if !root_closed => {
                if element.local_name().as_ref() != b"coreProperties"
                    || namespace != root_namespace.as_deref()
                {
                    return Err(Error::Invalid(
                        "malformed coreProperties root element".to_string(),
                    ));
                }
                root_closed = true;
            },
            Event::End(_) => {
                return Err(Error::Invalid(
                    "unexpected element after coreProperties".to_string(),
                ));
            },
            Event::DocType(_) => {
                return Err(Error::Invalid(
                    "DTD is not allowed in core properties".to_string(),
                ));
            },
            Event::Eof => break,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
        buffer.clear();
    }
    if root_namespace.is_none() || !root_closed || active.is_some() {
        return Err(Error::Invalid(
            "unterminated or missing coreProperties root element".to_string(),
        ));
    }
    let dialect = match root_namespace.as_deref() {
        Some(CORE_PROPERTIES_NAMESPACE) => Dialect::Transitional,
        Some(STRICT_CORE_PROPERTIES_NAMESPACE) => Dialect::Strict,
        _ => {
            return Err(Error::Invalid(
                "unsupported core-properties namespace".to_owned(),
            ));
        },
    };
    Ok((props, dialect))
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
        return Err(Error::Invalid(
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
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"version") => {
            CoreProperty::Version
        },
        (Some(CORE_PROPERTIES_NAMESPACE | STRICT_CORE_PROPERTIES_NAMESPACE), b"lastPrinted") => {
            CoreProperty::LastPrinted
        },
        (Some(DUBLIN_CORE_TERMS_NAMESPACE), _) => {
            return Err(Error::Invalid(
                "OPC M4.3 permits only dcterms:created and dcterms:modified".to_string(),
            ));
        },
        _ => {
            return Err(Error::Invalid(format!(
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

fn keyword_lang(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<Option<keyword::Lang>> {
    let mut language = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| Error::Xml(format!("invalid keyword attribute: {error}")))?;
        let raw_key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if matches!(raw_key, b"xmlns") || raw_key.starts_with(b"xmlns:") {
            if value.as_bytes() == MARKUP_COMPATIBILITY_NAMESPACE {
                return Err(Error::Invalid(
                    "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
                ));
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = bound_namespace(&namespace);
        reject_markup_compatibility(namespace)?;
        if namespace == Some(XML_NAMESPACE) && local.as_ref() == b"lang" {
            if language.is_some() {
                return Err(Error::Invalid(
                    "keyword contains duplicate xml:lang attributes".to_owned(),
                ));
            }
            language = Some(keyword::Lang::new(value.into_owned())?);
            continue;
        }
        return Err(Error::Invalid(format!(
            "attribute '{}' is not allowed on keywords",
            String::from_utf8_lossy(raw_key)
        )));
    }
    Ok(language)
}

fn start_keyword_value(
    active: &mut Option<Active>,
    namespace: Option<&[u8]>,
    root_namespace: Option<&[u8]>,
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    empty: bool,
) -> Result<()> {
    let Some(Active::Keywords {
        value,
        child,
        bytes: _,
    }) = active.as_mut()
    else {
        return Err(Error::Invalid(
            "only cp:keywords may contain child elements".to_owned(),
        ));
    };
    if child.is_some() || namespace != root_namespace || element.local_name().as_ref() != b"value" {
        return Err(Error::Invalid(
            "cp:keywords permits only non-nested cp:value children".to_owned(),
        ));
    }
    let child_value = keyword::Value {
        text: String::new(),
        lang: keyword_lang(reader, element)?,
    };
    if empty {
        value.items.push(Item::Value(child_value));
    } else {
        *child = Some(child_value);
    }
    Ok(())
}

fn close_active(
    active: &mut Option<Active>,
    namespace: Option<&[u8]>,
    root_namespace: Option<&[u8]>,
    element: &BytesEnd<'_>,
    props: &mut Props,
) -> Result<()> {
    if matches!(active, Some(Active::Keywords { child: Some(_), .. })) {
        if namespace != root_namespace || element.local_name().as_ref() != b"value" {
            return Err(Error::Invalid(
                "malformed cp:value element nesting".to_owned(),
            ));
        }
        let Some(Active::Keywords { value, child, .. }) = active.as_mut() else {
            return Err(Error::Invalid("missing keyword state".to_owned()));
        };
        let child = child
            .take()
            .ok_or_else(|| Error::Invalid("missing cp:value state".to_owned()))?;
        value.items.push(Item::Value(child));
        return Ok(());
    }

    let finished = active
        .take()
        .ok_or_else(|| Error::Invalid("core property has no active state".to_owned()))?;
    match finished {
        Active::Text(property, value) => {
            if parse_property(namespace, element.local_name().as_ref())? != property {
                return Err(Error::Invalid(
                    "malformed core property element nesting".to_owned(),
                ));
            }
            apply_property(props, property, value)
        },
        Active::Keywords {
            value,
            child: None,
            bytes: _,
        } => {
            if namespace != root_namespace || element.local_name().as_ref() != b"keywords" {
                return Err(Error::Invalid(
                    "malformed cp:keywords element nesting".to_owned(),
                ));
            }
            props.keywords = Some(value);
            Ok(())
        },
        Active::Keywords { child: Some(_), .. } => Err(Error::Invalid(
            "unterminated cp:value in cp:keywords".to_owned(),
        )),
    }
}

fn validate_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    property: Option<CoreProperty>,
) -> Result<()> {
    let mut has_w3cdtf = false;
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::Xml(format!("invalid core property attribute: {error}")))?;
        let raw_key = attribute.key.as_ref();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if matches!(raw_key, b"xmlns") || raw_key.starts_with(b"xmlns:") {
            if value.as_bytes() == MARKUP_COMPATIBILITY_NAMESPACE {
                return Err(Error::Invalid(
                    "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
                ));
            }
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = bound_namespace(&namespace);
        reject_markup_compatibility(namespace)?;
        if namespace == Some(XML_NAMESPACE) && local.as_ref() == b"lang" {
            return Err(Error::Invalid(
                "OPC M4.4 forbids xml:lang in core properties".to_string(),
            ));
        }
        if namespace == Some(XSI_NAMESPACE) && local.as_ref() == b"type" {
            let Some(property) = property else {
                return Err(Error::Invalid(
                    "OPC M4.5 forbids xsi:type on coreProperties".to_string(),
                ));
            };
            if !property.requires_w3cdtf() || has_w3cdtf {
                return Err(Error::Invalid(
                    "OPC M4.5 permits one xsi:type only on created or modified".to_string(),
                ));
            }
            let (value_namespace, value_local) =
                reader.resolver().resolve_element(QName(value.as_bytes()));
            if value.as_ref() != "dcterms:W3CDTF"
                || bound_namespace(&value_namespace) != Some(DUBLIN_CORE_TERMS_NAMESPACE)
                || value_local.as_ref() != b"W3CDTF"
            {
                return Err(Error::Invalid(
                    "OPC M4.5 requires xsi:type=\"dcterms:W3CDTF\"".to_string(),
                ));
            }
            has_w3cdtf = true;
            continue;
        }
        return Err(Error::Invalid(format!(
            "attribute '{}' is not allowed on core properties",
            String::from_utf8_lossy(raw_key)
        )));
    }
    if property.is_some_and(CoreProperty::requires_w3cdtf) && !has_w3cdtf {
        return Err(Error::Invalid(
            "OPC M4.5 requires xsi:type on created and modified".to_string(),
        ));
    }
    Ok(())
}

fn reject_markup_compatibility(namespace: Option<&[u8]>) -> Result<()> {
    if namespace == Some(MARKUP_COMPATIBILITY_NAMESPACE) {
        return Err(Error::Invalid(
            "OPC M4.2 forbids the Markup Compatibility namespace".to_string(),
        ));
    }
    Ok(())
}

fn mark_seen(seen: &mut [bool; CoreProperty::COUNT], property: CoreProperty) -> Result<()> {
    if std::mem::replace(&mut seen[property.index()], true) {
        return Err(Error::Invalid(format!(
            "duplicate core property '{}'",
            String::from_utf8_lossy(property_name(property))
        )));
    }
    Ok(())
}

fn push_property_text(active: &mut Option<Active>, value: &str) -> Result<()> {
    let Some(active) = active.as_mut() else {
        if value.trim().is_empty() {
            return Ok(());
        }
        return Err(Error::Invalid(
            "non-whitespace text outside a core property".to_string(),
        ));
    };
    match active {
        Active::Text(_, text) => append_bounded(text, value),
        Active::Keywords {
            value: keywords,
            child,
            bytes,
        } => {
            *bytes = bytes.checked_add(value.len()).ok_or(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual: usize::MAX,
            })?;
            if *bytes > MAX_PROPERTY_TEXT {
                return Err(Error::Limit {
                    resource: "core property text bytes",
                    max: MAX_PROPERTY_TEXT,
                    actual: *bytes,
                });
            }
            if let Some(child) = child.as_mut() {
                child.text.push_str(value);
            } else {
                keywords.append_text(value.to_owned());
            }
            Ok(())
        },
    }
}

fn append_bounded(text: &mut String, value: &str) -> Result<()> {
    let length = text.len().checked_add(value.len()).ok_or(Error::Limit {
        resource: "core property text bytes",
        max: MAX_PROPERTY_TEXT,
        actual: usize::MAX,
    })?;
    if length > MAX_PROPERTY_TEXT {
        return Err(Error::Limit {
            resource: "core property text bytes",
            max: MAX_PROPERTY_TEXT,
            actual: length,
        });
    }
    text.push_str(value);
    Ok(())
}

fn apply_property(props: &mut Props, property: CoreProperty, value: String) -> Result<()> {
    match property {
        CoreProperty::Title => props.title = Some(value),
        CoreProperty::Subject => props.subject = Some(value),
        CoreProperty::Author => props.creator = Some(value),
        CoreProperty::Keywords => {
            return Err(Error::Invalid(
                "keyword mixed content bypassed its typed parser".to_owned(),
            ));
        },
        CoreProperty::Description => props.description = Some(value),
        CoreProperty::Identifier => props.identifier = Some(value),
        CoreProperty::Language => props.language = Some(value),
        CoreProperty::LastModifiedBy => props.last_modified_by = Some(value),
        CoreProperty::Revision => props.revision = Some(value),
        CoreProperty::Category => props.category = Some(value),
        CoreProperty::ContentStatus => props.content_status = Some(value),
        CoreProperty::Version => props.version = Some(value),
        CoreProperty::Created => props.created = Some(time::W3c::new(value)?),
        CoreProperty::Modified => props.modified = Some(time::W3c::new(value)?),
        CoreProperty::LastPrinted => props.last_printed = Some(time::DateTime::new(value)?),
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
        CoreProperty::Version => b"version",
        CoreProperty::Created => b"created",
        CoreProperty::Modified => b"modified",
        CoreProperty::LastPrinted => b"lastPrinted",
    }
}

fn bound_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Props;
    use crate::properties::STRICT_CORE_PROPERTIES_RELATIONSHIP;
    use chrono::{DateTime, Utc};
    use litchi_core::OwnedSource;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{PackURI, SourceBackedPackage, part::BlobPart};
    use std::sync::Arc;

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

    fn source_package_with_core(xml: &[u8], relationship: &str) -> SourceBackedPackage {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            ct::OPC_CORE_PROPERTIES.to_owned(),
            xml.to_vec(),
        )));
        package.relate_to("docProps/core.xml", relationship);
        let mut bytes = Vec::new();
        package.to_stream(&mut bytes).unwrap();
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
    }

    fn source_package(package: &OpcPackage) -> SourceBackedPackage {
        let mut bytes = Vec::new();
        package.to_stream(&mut bytes).unwrap();
        SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
    }

    fn error_family(error: &Error) -> &'static str {
        match error {
            Error::Opc(_) => "opc",
            Error::Xml(_) => "xml",
            Error::Missing(_) => "missing",
            Error::ContentType { .. } => "content-type",
            Error::Relationship(_) => "relationship",
            Error::Invalid(_) => "invalid",
            Error::Limit { .. } => "limit",
            Error::Uri(_) => "uri",
            #[cfg(feature = "vba-inspection")]
            Error::Vba(_) => "vba",
            Error::SpreadsheetXmlMaps(_) => "spreadsheet-xml-maps",
            Error::Mce(_) => "mce",
            Error::Decode(_) => "decode",
            Error::Io(_) => "io",
            Error::Fmt(_) => "fmt",
        }
    }

    #[test]
    fn selects_only_relationship_target_and_allows_absence() {
        let valid = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Target</dc:title></cp:coreProperties>"#;
        let decoy = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Decoy</dc:title></cp:coreProperties>"#;
        let mut package = package_with_core("/custom/real.xml", ct::OPC_CORE_PROPERTIES, valid);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/docProps/core.xml").unwrap(),
            "application/xml".to_owned(),
            decoy.to_vec(),
        )));
        assert_eq!(
            read(&package).unwrap().unwrap().title.as_deref(),
            Some("Target")
        );
        assert!(read(&OpcPackage::new()).unwrap().is_none());
    }

    #[test]
    fn source_reader_matches_eager_for_present_absent_and_strict_properties() {
        let valid = br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Target</dc:title></cp:coreProperties>"#;
        let eager = package_with_core("/docProps/core.xml", ct::OPC_CORE_PROPERTIES, valid);
        let mut bytes = Vec::new();
        eager.to_stream(&mut bytes).unwrap();
        let source = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert_eq!(read_source_backed(&source).unwrap(), read(&eager).unwrap());

        let empty = OpcPackage::new();
        let mut bytes = Vec::new();
        empty.to_stream(&mut bytes).unwrap();
        let source = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert_eq!(read_source_backed(&source).unwrap(), read(&empty).unwrap());

        let strict_xml = br#"<cp:coreProperties xmlns:cp="http://purl.oclc.org/ooxml/package/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Strict</dc:title></cp:coreProperties>"#;
        let source = source_package_with_core(strict_xml, STRICT_CORE_PROPERTIES_RELATIONSHIP);
        assert_eq!(
            read_source_backed(&source)
                .unwrap()
                .unwrap()
                .title
                .as_deref(),
            Some("Strict")
        );
    }

    #[test]
    fn source_reader_matches_eager_graph_failures() {
        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.test/core.xml", rt::CORE_PROPERTIES);

        let mut dangling = OpcPackage::new();
        dangling.relate_to("missing.xml", rt::CORE_PROPERTIES);

        let wrong = package_with_core(
            "/docProps/core.xml",
            ct::WML_DOCUMENT,
            br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>"#,
        );

        for package in [&external, &dangling, &wrong] {
            let eager = read(package).unwrap_err();
            let source = source_package(package);
            let deferred = read_source_backed(&source).unwrap_err();
            assert_eq!(error_family(&eager), error_family(&deferred));
        }
    }

    #[test]
    fn rejects_external_dangling_and_wrong_content_type_relationships() {
        let mut external = OpcPackage::new();
        external.relate_to_external("https://example.test/core.xml", rt::CORE_PROPERTIES);
        assert!(matches!(read(&external), Err(Error::Relationship(_))));

        let mut dangling = OpcPackage::new();
        dangling.relate_to("missing.xml", rt::CORE_PROPERTIES);
        assert!(matches!(read(&dangling), Err(Error::Missing(_))));

        let wrong = package_with_core(
            "/core.xml",
            ct::WML_DOCUMENT,
            br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>"#,
        );
        assert!(matches!(read(&wrong), Err(Error::ContentType { .. })));
    }

    #[test]
    fn parses_all_schema_fields_without_normalizing_lexical_values() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>T</dc:title><dc:subject>S</dc:subject><dc:creator>A</dc:creator><dc:description>D</dc:description><dc:identifier>I</dc:identifier><dc:language>en-GB</dc:language><cp:keywords>K</cp:keywords><cp:lastModifiedBy>M</cp:lastModifiedBy><cp:revision>007</cp:revision><cp:category>C</cp:category><cp:contentStatus>Draft</cp:contentStatus><cp:version>2.0</cp:version><dcterms:created xsi:type="dcterms:W3CDTF">2006-10-13T18:06:00.123+03:00</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2007-06-20T07:59:00-13:00</dcterms:modified><cp:lastPrinted>2008-01-02T03:04:05Z</cp:lastPrinted></cp:coreProperties>"#;
        let (props, dialect) = decode(xml).unwrap();
        assert_eq!(dialect, Dialect::Transitional);
        assert_eq!(props.identifier.as_deref(), Some("I"));
        assert_eq!(props.language.as_deref(), Some("en-GB"));
        assert_eq!(props.version.as_deref(), Some("2.0"));
        assert_eq!(props.keywords.as_ref().and_then(Keywords::plain), Some("K"));
        assert_eq!(props.revision.as_deref(), Some("007"));
        assert_eq!(
            props.created.as_ref().map(time::W3c::as_str),
            Some("2006-10-13T18:06:00.123+03:00")
        );
        assert_eq!(
            props.modified.as_ref().map(time::W3c::as_str),
            Some("2007-06-20T07:59:00-13:00")
        );
        assert_eq!(
            props.last_printed.as_ref().map(time::DateTime::as_str),
            Some("2008-01-02T03:04:05Z")
        );
    }

    #[test]
    fn shared_writer_round_trips_through_strict_parser() {
        let created = DateTime::parse_from_rfc3339("2024-03-04T05:06:07+02:00")
            .unwrap()
            .with_timezone(&Utc);
        let properties = Props::new()
            .title("A & B")
            .creator("Writer")
            .identifier("urn:test:42")
            .language("en-US")
            .revision("12")
            .version("3.0")
            .created(created)
            .last_printed(created);
        let xml = properties.xml().unwrap();
        let (parsed, dialect) = decode(&xml).unwrap();
        assert_eq!(dialect, Dialect::Transitional);
        assert_eq!(parsed.title.as_deref(), Some("A & B"));
        assert_eq!(parsed.creator.as_deref(), Some("Writer"));
        assert_eq!(parsed.identifier.as_deref(), Some("urn:test:42"));
        assert_eq!(parsed.language.as_deref(), Some("en-US"));
        assert_eq!(parsed.revision.as_deref(), Some("12"));
        assert_eq!(parsed.version.as_deref(), Some("3.0"));
        assert_eq!(
            parsed.created.as_ref().map(time::W3c::as_str),
            Some("2024-03-04T03:06:07Z")
        );
        assert_eq!(
            parsed.last_printed.as_ref().map(time::DateTime::as_str),
            Some("2024-03-04T03:06:07Z")
        );
    }

    #[test]
    fn rejects_wrong_root_duplicates_invalid_values_dtd_and_entities() {
        assert!(decode("<props/>").is_err());
        let duplicate = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>A</dc:title><dc:title>B</dc:title></cp:coreProperties>"#;
        assert!(decode(duplicate).is_err());
        let non_schema = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:contentType>text/plain</cp:contentType></cp:coreProperties>"#;
        assert!(decode(non_schema).is_err());
        let bad_last_printed = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:lastPrinted>2026-08-03</cp:lastPrinted></cp:coreProperties>"#;
        assert!(decode(bad_last_printed).is_err());
        assert!(decode("<!DOCTYPE x><x/>").is_err());
        let package = poi_package("CorePropertiesHasEntities.ooxml");
        assert!(read(&package).is_err());
    }

    #[test]
    fn revision_is_a_lossless_string_including_empty_values() {
        for value in ["", "seven", "0007", " 7 "] {
            let xml = Props::new().revision(value).xml().unwrap();
            let (parsed, _) = decode(&xml).unwrap();
            assert_eq!(parsed.revision.as_deref(), Some(value));
        }
    }

    #[test]
    fn keywords_preserve_mixed_content_and_language_annotations() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:keywords xml:lang="">lead<cp:value xml:lang="en-CA">colour</cp:value>;<cp:value xml:lang="fr">couleur</cp:value>tail</cp:keywords></cp:coreProperties>"#;
        let (props, _) = decode(xml).unwrap();
        let expected = Keywords::new()
            .lang(keyword::Lang::new("").unwrap())
            .text("lead")
            .value(keyword::Value::new("colour").lang(keyword::Lang::new("en-CA").unwrap()))
            .text(";")
            .value(keyword::Value::new("couleur").lang(keyword::Lang::new("fr").unwrap()))
            .text("tail");
        assert_eq!(props.keywords.as_ref(), Some(&expected));
        assert_eq!(
            props.keywords.as_ref().unwrap().joined(),
            "leadcolour;couleurtail"
        );

        let encoded = Props::new().keywords(expected.clone()).xml().unwrap();
        assert!(encoded.contains(r#"<cp:keywords xml:lang="">"#));
        assert!(encoded.contains(r#"<cp:value xml:lang="en-CA">colour</cp:value>"#));
        assert_eq!(decode(&encoded).unwrap().0.keywords, Some(expected));
    }

    #[test]
    fn rejects_invalid_keyword_children_and_language_placement() {
        let nested = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:keywords><cp:value><cp:value>x</cp:value></cp:value></cp:keywords></cp:coreProperties>"#;
        assert!(decode(nested).is_err());
        let wrong_child = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:x="urn:test"><cp:keywords><x:value>x</x:value></cp:keywords></cp:coreProperties>"#;
        assert!(decode(wrong_child).is_err());
        let illegal_language = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:subject xml:lang="en">x</dc:subject></cp:coreProperties>"#;
        assert!(decode(illegal_language).is_err());
    }

    #[test]
    fn w3cdtf_supports_partial_and_unzoned_forms_but_xsi_type_is_literal() {
        let xml = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">2026</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-08-03T04:05:06</dcterms:modified><cp:lastPrinted>2026-08-03T04:05:06</cp:lastPrinted></cp:coreProperties>"#;
        let (props, _) = decode(xml).unwrap();
        assert_eq!(props.created.as_ref().map(time::W3c::as_str), Some("2026"));
        assert_eq!(
            props.modified.as_ref().map(time::W3c::as_str),
            Some("2026-08-03T04:05:06")
        );
        assert_eq!(
            props.last_printed.as_ref().map(time::DateTime::as_str),
            Some("2026-08-03T04:05:06")
        );

        let alias = r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:t="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="t:W3CDTF">2026</dcterms:created></cp:coreProperties>"#;
        assert!(decode(alias).is_err());
    }

    #[test]
    fn rejects_property_text_beyond_the_declared_budget() {
        let value = "x".repeat(MAX_PROPERTY_TEXT + 1);
        let xml = format!(
            r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{value}</dc:title></cp:coreProperties>"#
        );
        assert!(matches!(
            decode(&xml),
            Err(Error::Limit {
                resource: "core property text bytes",
                max: MAX_PROPERTY_TEXT,
                actual,
            }) if actual == MAX_PROPERTY_TEXT + 1
        ));
    }

    #[test]
    fn matches_poi_success_timezone_and_m4_conformance_fixtures() {
        let success = read(&poi_package("OPCCompliance_CoreProperties_SUCCESS.docx"))
            .unwrap()
            .unwrap();
        assert_eq!(success.title.as_deref(), Some("MyTitle"));
        assert_eq!(success.revision.as_deref(), Some("2"));

        let timezone = read(&poi_package(
            "OPCCompliance_CoreProperties_AlternateTimezones.docx",
        ))
        .unwrap()
        .unwrap();
        assert_eq!(
            timezone.created.as_ref().map(time::W3c::as_str),
            Some("2006-10-13T18:06:00.123+03:00")
        );
        assert_eq!(
            timezone.modified.as_ref().map(time::W3c::as_str),
            Some("2007-06-20T07:59:00-13:00")
        );

        for name in [
            "OPCCompliance_CoreProperties_DoNotUseCompatibilityMarkupFAIL.docx",
            "OPCCompliance_CoreProperties_DCTermsNamespaceLimitedUseFAIL.docx",
            "OPCCompliance_CoreProperties_UnauthorizedXMLLangAttributeFAIL.docx",
            "OPCCompliance_CoreProperties_LimitedXSITypeAttribute_NotPresentFAIL.docx",
            "OPCCompliance_CoreProperties_LimitedXSITypeAttribute_PresentWithUnauthorizedValueFAIL.docx",
        ] {
            assert!(read(&poi_package(name)).is_err(), "accepted {name}");
        }
    }

    #[test]
    fn poi_no_core_is_empty_and_multiple_relationships_fail_during_package_load() {
        let empty = read(&poi_package("OPCCompliance_NoCoreProperties.xlsx")).unwrap();
        assert!(empty.is_none());
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../test-data/poi/test-data/openxml4j/OPCCompliance_CoreProperties_OnlyOneCorePropertiesPartFAIL.docx",
        );
        assert!(OpcPackage::from_bytes(&std::fs::read(path).unwrap()).is_err());
    }
}
