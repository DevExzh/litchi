use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub(crate) const PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/presentationml/2006/main";
pub(crate) const STRICT_PRESENTATIONML_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) use litchi_ooxml_common::relationships::{
    STRICT_NAMESPACE as STRICT_RELATIONSHIPS_NAMESPACE,
    TRANSITIONAL_NAMESPACE as RELATIONSHIPS_NAMESPACE,
};

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
    litchi_ooxml_common::relationships::attribute_value(element, name, decoder, resolver)
        .map_err(OoxmlError::from)
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
