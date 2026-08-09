//! XML codecs for OOXML relationship attributes.

use crate::XmlError;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};

use super::Id;

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
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
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

/// Decode one relationship attribute into a checked [`Id`].
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn attribute_id(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<Id>, XmlError> {
    attribute_value(element, name, decoder, resolver)?
        .map(Id::new)
        .transpose()
}
