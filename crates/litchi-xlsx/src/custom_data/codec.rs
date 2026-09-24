//! XLSX error-mapping facade for the shared Custom Data XML codec.

use super::model::{ExtensionList, Properties};
use crate::error::{Error, Result};
use std::sync::Arc;

/// XML-only resource profile used by the package-neutral Custom Data codec.
pub use litchi_ooxml_common::custom_data::codec::Limits as XmlLimits;

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "the shared error is non-exhaustive; host limits and allocations are mapped explicitly"
)]
fn map_common_error(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Limit {
            resource,
            max,
            actual,
        } => Error::ResourceLimit(litchi_core::ResourceLimit {
            // Keep byte and scalar-size limits as byte admission failures.
            // Only structural cardinalities are object charges: a namespace
            // byte ceiling or an escaped attribute length is still a byte
            // ceiling, even though its diagnostic label contains
            // `namespace` or `attribute`.
            resource: match resource {
                "XML depth" => litchi_core::Resource::Depth,
                "XML node count"
                | "XML event count"
                | "XML attribute count"
                | "XML namespace binding count" => litchi_core::Resource::Objects,
                "serialized properties XML bytes" | "properties XML output" => {
                    litchi_core::Resource::OutputBytes
                },
                _ => litchi_core::Resource::InputBytes,
            },
            observed: actual as u64,
            limit: max as u64,
            scope: Arc::<str>::from(format!("XLSX Custom Data {resource}")),
        }),
        litchi_ooxml_common::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
        other => Error::Common(other),
    }
}

pub fn parse_properties(xml: &[u8]) -> Result<Properties> {
    litchi_ooxml_common::custom_data::parse_properties(xml).map_err(Error::from)
}

/// Parse Custom Data Properties under an explicit XML resource profile.
pub fn parse_properties_with_limits(xml: &[u8], limits: &XmlLimits) -> Result<Properties> {
    litchi_ooxml_common::custom_data::parse_properties_with_limits(xml, limits)
        .map_err(map_common_error)
}

pub fn validate_workbook_root(xml: &[u8]) -> Result<()> {
    litchi_ooxml_common::custom_data::validate_workbook_root(xml).map_err(Error::from)
}

/// Validate a workbook root under an explicit XML resource profile.
pub fn validate_workbook_root_with_limits(xml: &[u8], limits: &XmlLimits) -> Result<()> {
    litchi_ooxml_common::custom_data::validate_workbook_root_with_limits(xml, limits)
        .map_err(map_common_error)
}

pub fn write_properties(value: &Properties) -> Result<Vec<u8>> {
    litchi_ooxml_common::custom_data::write_properties(value).map_err(Error::from)
}

/// Serialize Custom Data Properties under an explicit XML resource profile.
pub fn write_properties_with_limits(value: &Properties, limits: &XmlLimits) -> Result<Vec<u8>> {
    litchi_ooxml_common::custom_data::write_properties_with_limits(value, limits)
        .map_err(map_common_error)
}

pub(crate) fn validate_source_properties_with_limits(
    value: &Properties,
    limits: &XmlLimits,
) -> Result<()> {
    litchi_ooxml_common::custom_data::codec::validate_source_properties_with_limits(value, limits)
        .map_err(map_common_error)
}

pub(crate) fn rewrite_id_with_limits_and_output(
    source: &litchi_opc::OwnedXmlPart,
    id: &str,
    limits: &XmlLimits,
    max_output_bytes: usize,
) -> Result<litchi_opc::OwnedXmlPart> {
    litchi_ooxml_common::custom_data::codec::rewrite_id_with_limits_and_output(
        source,
        id,
        limits,
        max_output_bytes,
    )
    .map_err(map_common_error)
}

pub(crate) fn rewrite_extension_list_with_limits_and_output(
    source: &litchi_opc::OwnedXmlPart,
    extension: Option<&ExtensionList>,
    limits: &XmlLimits,
    max_output_bytes: usize,
) -> Result<litchi_opc::OwnedXmlPart> {
    litchi_ooxml_common::custom_data::codec::rewrite_extension_list_with_limits_and_output(
        source,
        extension,
        limits,
        max_output_bytes,
    )
    .map_err(map_common_error)
}

pub(crate) fn canonical_extension_with_limits(
    extension: Option<&ExtensionList>,
    limits: &XmlLimits,
) -> Result<Option<Vec<u8>>> {
    litchi_ooxml_common::custom_data::codec::canonical_extension_with_limits(extension, limits)
        .map_err(map_common_error)
}
