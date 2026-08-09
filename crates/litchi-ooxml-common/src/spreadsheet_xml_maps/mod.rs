//! Format-neutral `SpreadsheetML` Custom XML Maps model and bounded XML codec.

mod codec;
mod model;
mod validation;

pub use codec::{
    parse_xml_map_info, parse_xml_map_info_with_conformance,
    parse_xml_map_info_with_conformance_and_limits, parse_xml_map_info_with_limits,
    patch_xml_map_info_source, patch_xml_map_info_source_ref,
    patch_xml_map_info_source_ref_with_limits, patch_xml_map_info_source_with_limits,
    serialize_xml_map_info, serialize_xml_map_info_ref, serialize_xml_map_info_ref_with_limits,
    serialize_xml_map_info_with_limits,
};
pub use codec::{
    parse_xml_map_info as parse_map_info, patch_xml_map_info_source as patch_map_info_source,
    serialize_xml_map_info as serialize_map_info,
};
pub use model::{
    CONTENT_TYPE, DataBinding, DataBindingRef, MAX_DEPTH, MAX_EVENTS, MAX_MAPS, MAX_OPAQUE_BYTES,
    MAX_PART_BYTES, MAX_SCHEMAS, MAX_STRING_BYTES, NS, NS_TEXT, ParsedXmlMapInfo, REL, STRICT_NS,
    STRICT_NS_TEXT, STRICT_REL, XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo,
    XmlMapInfoRef, XmlMapLimits, XmlMapRef, XmlMapSchema, XmlSchema, XmlSchemaRef,
};
pub use validation::validate_xml_map_info as validate_map_info;
pub use validation::{
    validate_xml_map_info, validate_xml_map_info_ref, validate_xml_map_info_ref_with_limits,
    validate_xml_map_info_with_limits,
};

fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::SpreadsheetXmlMaps(message.into())
}

#[cfg(test)]
mod tests;
