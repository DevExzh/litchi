//! `SpreadsheetML` namespace helpers.

use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};

use crate::error::Result;
use litchi_ooxml_common::relationships::attribute_value;

/// Transitional `SpreadsheetML` main namespace.
pub const SPREADSHEETML_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/// Strict `SpreadsheetML` main namespace.
pub const STRICT_SPREADSHEETML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";

/// Test an expanded element name against either `SpreadsheetML` dialect.
#[must_use]
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
    attribute_value(element, name, decoder, resolver).map_err(crate::error::Error::from)
}
