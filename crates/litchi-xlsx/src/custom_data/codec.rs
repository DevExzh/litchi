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
    use litchi_ooxml_common::LimitKind as CommonLimit;
    match error {
        litchi_ooxml_common::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::ResourceLimit(litchi_core::ResourceLimit {
            // Keep byte and scalar-size limits as byte admission failures.
            // Only structural cardinalities are object charges: a namespace
            // byte ceiling or an escaped attribute length is still a byte
            // ceiling, even though its diagnostic label contains
            // `namespace` or `attribute`.
            resource: match kind {
                CommonLimit::XmlDepth => litchi_core::Resource::Depth,
                CommonLimit::XmlNodeCount
                | CommonLimit::XmlEventCount
                | CommonLimit::XmlAttributeCount
                | CommonLimit::XmlNamespaceBindingCount => litchi_core::Resource::Objects,
                CommonLimit::SerializedPropertiesXmlBytes | CommonLimit::PropertiesXmlOutput => {
                    litchi_core::Resource::OutputBytes
                },
                _ => litchi_core::Resource::InputBytes,
            },
            observed: observed as u64,
            limit: maximum as u64,
            scope: Arc::<str>::from(format!("XLSX Custom Data {kind}")),
        }),
        litchi_ooxml_common::Error::Allocation { resource, source } => {
            Error::Allocation { resource, source }
        },
        other => Error::Common(other),
    }
}

/// Parse Custom Data Properties under the XML resource profile `limits`
/// ([`XmlLimits::standard`] for the default ceilings).
pub fn parse_properties(xml: &[u8], limits: &XmlLimits) -> Result<Properties> {
    litchi_ooxml_common::custom_data::parse_properties(xml, limits).map_err(map_common_error)
}

/// Validate a workbook root under the XML resource profile `limits`.
pub fn validate_workbook_root(xml: &[u8], limits: &XmlLimits) -> Result<()> {
    litchi_ooxml_common::custom_data::validate_workbook_root(xml, limits).map_err(map_common_error)
}

/// Serialize Custom Data Properties under the XML resource profile `limits`.
pub fn write_properties(value: &Properties, limits: &XmlLimits) -> Result<Vec<u8>> {
    litchi_ooxml_common::custom_data::write_properties(value, limits).map_err(map_common_error)
}

pub(crate) fn validate_source_properties(value: &Properties, limits: &XmlLimits) -> Result<()> {
    litchi_ooxml_common::custom_data::codec::validate_source_properties(value, limits)
        .map_err(map_common_error)
}

pub(crate) fn rewrite_id(
    source: &litchi_opc::OwnedXmlPart,
    id: &str,
    limits: &XmlLimits,
    max_output_bytes: usize,
) -> Result<litchi_opc::OwnedXmlPart> {
    litchi_ooxml_common::custom_data::codec::rewrite_id(source, id, limits, max_output_bytes)
        .map_err(map_common_error)
}

pub(crate) fn rewrite_extension_list(
    source: &litchi_opc::OwnedXmlPart,
    extension: Option<&ExtensionList>,
    limits: &XmlLimits,
    max_output_bytes: usize,
) -> Result<litchi_opc::OwnedXmlPart> {
    litchi_ooxml_common::custom_data::codec::rewrite_extension_list(
        source,
        extension,
        limits,
        max_output_bytes,
    )
    .map_err(map_common_error)
}

pub(crate) fn canonical_extension(
    extension: Option<&ExtensionList>,
    limits: &XmlLimits,
) -> Result<Option<Vec<u8>>> {
    litchi_ooxml_common::custom_data::codec::canonical_extension(extension, limits)
        .map_err(map_common_error)
}
