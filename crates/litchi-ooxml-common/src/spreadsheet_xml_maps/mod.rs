//! Format-neutral `SpreadsheetML` Custom XML Maps model and bounded XML codec.

mod codec;
mod model;
mod validation;

pub use codec::{
    parse_xml_map_info, parse_xml_map_info_with_conformance, patch_xml_map_info_source,
    patch_xml_map_info_source_ref, serialize_xml_map_info, serialize_xml_map_info_ref,
};
pub use model::{
    CONTENT_TYPE, DataBinding, DataBindingRef, MAX_DEPTH, MAX_EVENTS, MAX_MAPS, MAX_OPAQUE_BYTES,
    MAX_PART_BYTES, MAX_SCHEMAS, MAX_STRING_BYTES, NS, NS_TEXT, ParsedXmlMapInfo, REL, STRICT_NS,
    STRICT_NS_TEXT, STRICT_REL, XmlMap, XmlMapConformance, XmlMapInfo, XmlMapInfoRef, XmlMapLimits,
    XmlMapRef, XmlSchema, XmlSchemaRef,
};
pub use validation::{validate_xml_map_info, validate_xml_map_info_ref};

fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::SpreadsheetXmlMaps(message.into())
}

#[cfg(test)]
mod tests;
