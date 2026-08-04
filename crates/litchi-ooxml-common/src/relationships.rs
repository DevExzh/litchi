//! Shared OOXML relationship-namespace attribute decoding.

use crate::XmlError;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};

/// Transitional OOXML relationships namespace.
pub const TRANSITIONAL_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// Strict OOXML relationships namespace.
pub const STRICT_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// Decode one attribute in the OOXML relationships namespace.
///
/// Unresolved `r:` fragments are accepted because a caller may pass an XML
/// slice whose namespace declaration belongs to an omitted ancestor. Duplicate
/// semantic attributes are rejected before any value is returned.
pub fn attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>, XmlError> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| XmlError::Malformed(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_relationship = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if value == TRANSITIONAL_NAMESPACE || value == STRICT_NAMESPACE
        ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r");
        if !is_relationship {
            continue;
        }
        if value.is_some() {
            return Err(XmlError::Invalid(format!(
                "duplicate relationship attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| XmlError::Malformed(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::reader::NsReader;

    #[test]
    fn accepts_transitional_and_strict_relationship_attributes() {
        for (xml, expected) in [
            (
                format!(
                    r#"<p:item xmlns:p="urn:test" xmlns:r="{}" r:id="rId1"/>"#,
                    String::from_utf8_lossy(TRANSITIONAL_NAMESPACE)
                ),
                "rId1",
            ),
            (
                format!(
                    r#"<p:item xmlns:p="urn:test" xmlns:r="{}" r:id="rId2"/>"#,
                    String::from_utf8_lossy(STRICT_NAMESPACE)
                ),
                "rId2",
            ),
        ] {
            let mut reader = NsReader::from_reader(xml.as_bytes());
            let (_, event) = reader.read_resolved_event().expect("item");
            let quick_xml::events::Event::Empty(element) = event else {
                panic!("expected empty item");
            };
            assert_eq!(
                attribute_value(&element, b"id", reader.decoder(), reader.resolver())
                    .expect("relationship attribute")
                    .as_deref(),
                Some(expected)
            );
        }
    }

    #[test]
    fn rejects_duplicate_relationship_attributes() {
        let xml = format!(
            r#"<p:item xmlns:p="urn:test" xmlns:r="{}" xmlns:q="{}" r:id="one" q:id="two"/>"#,
            String::from_utf8_lossy(TRANSITIONAL_NAMESPACE),
            String::from_utf8_lossy(TRANSITIONAL_NAMESPACE)
        );
        let mut reader = NsReader::from_reader(xml.as_bytes());
        let (_, event) = reader.read_resolved_event().expect("item");
        let quick_xml::events::Event::Empty(element) = event else {
            panic!("expected empty item");
        };
        assert!(matches!(
            attribute_value(&element, b"id", reader.decoder(), reader.resolver()),
            Err(XmlError::Invalid(message)) if message.contains("duplicate")
        ));
    }
}
