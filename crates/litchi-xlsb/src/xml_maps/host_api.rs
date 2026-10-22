//! XLSB-contextual adapters for the shared SpreadsheetML MapInfo codec.
//!
//! MapInfo is format-neutral OOXML, so its model and parser live in
//! `litchi-ooxml-common`. XLSB callers, however, should not need to handle a
//! second error domain for a feature exposed by this crate. These adapters
//! retain the shared values while attaching the XLSB Custom XML Maps context
//! to every shared-codec failure.

use litchi_ooxml_common::spreadsheet_xml_maps as common;

use super::{Error, Result, XmlMapConformance, XmlMapInfo, XmlMapLimits};

fn contextualize(_operation: &str, _error: litchi_ooxml_common::Error) -> Error {
    Error::InvalidFormat("XLSB Custom XML Maps")
}

/// Parse a bounded, namespace-aware XLSB Custom XML Maps catalog under
/// `limits` ([`XmlMapLimits::DEFAULT`] for the default ceilings).
pub fn parse_xml_map_info(xml: &[u8], limits: &XmlMapLimits) -> Result<XmlMapInfo> {
    common::parse_xml_map_info(xml, limits).map_err(|error| contextualize("parse", error))
}

/// Parse a catalog under `limits` and report its SpreadsheetML namespace
/// family.
pub fn parse_xml_map_info_with_conformance(
    xml: &[u8],
    limits: &XmlMapLimits,
) -> Result<common::ParsedXmlMapInfo> {
    common::parse_xml_map_info_with_conformance(xml, limits)
        .map_err(|error| contextualize("parse", error))
}

/// Serialize a catalog canonically for the selected OOXML conformance family,
/// under `limits`.
pub fn serialize_xml_map_info(
    info: &XmlMapInfo,
    conformance: XmlMapConformance,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    common::serialize_xml_map_info(info, conformance, limits)
        .map_err(|error| contextualize("serialize", error))
}

/// Validate a catalog against the shared bounded SpreadsheetML vocabulary,
/// under `limits`.
pub fn validate_xml_map_info(info: &XmlMapInfo, limits: &XmlMapLimits) -> Result<()> {
    common::validate_xml_map_info(info, limits).map_err(|error| contextualize("validate", error))
}

/// Patch modeled catalog fields while preserving unaffected source markup,
/// under `limits` (validation and output ceilings).
pub fn patch_xml_map_info_source(
    source: &[u8],
    before: &XmlMapInfo,
    after: &XmlMapInfo,
    before_conformance: XmlMapConformance,
    after_conformance: XmlMapConformance,
    limits: &XmlMapLimits,
) -> Result<Vec<u8>> {
    common::patch_xml_map_info_source(
        source,
        before,
        after,
        before_conformance,
        after_conformance,
        limits,
    )
    .map_err(|error| contextualize("patch", error))
}
