use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(crate) fn is_presentationml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == PRESENTATIONML_NAMESPACE || *value == STRICT_PRESENTATIONML_NAMESPACE
        },
        // Shape fragments normally inherit the conventional prefix from a slide root.
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

pub(crate) fn relationship_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == RELATIONSHIPS_NAMESPACE || value == STRICT_RELATIONSHIPS_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r");
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate relationship attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

pub(crate) fn presentation_name(xml_bytes: &[u8]) -> Result<String> {
    let mut reader = NsReader::from_reader(xml_bytes);
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_presentationml_name(&namespace, element.name(), b"cSld") =>
            {
                return Ok(
                    unqualified_attribute_value(&element, b"name", decoder)?.unwrap_or_default()
                );
            },
            Event::Eof => return Ok(String::new()),
            _ => {},
        }
    }
}
