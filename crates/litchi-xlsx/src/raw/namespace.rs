//! SpreadsheetML namespace helpers.

use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};

use crate::error::{Result, invalid};

/// Transitional SpreadsheetML main namespace.
pub const SPREADSHEETML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/// Strict SpreadsheetML main namespace.
pub const STRICT_SPREADSHEETML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";

/// Test an expanded element name against either SpreadsheetML dialect.
pub fn is_spreadsheetml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == SPREADSHEETML_NAMESPACE
                    || *value == STRICT_SPREADSHEETML_NAMESPACE
        )
}

/// Decode one relationship-namespace attribute while rejecting duplicates.
pub fn relationship_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
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
            return Err(invalid(format!(
                "duplicate relationship attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| invalid(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}
