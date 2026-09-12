//! XLSX error-mapping facade for the shared Custom Data XML codec.

use super::model::{ExtensionList, Properties};
use crate::error::{Error, Result};

/// XML-only resource profile used by the package-neutral Custom Data codec.
pub use litchi_ooxml_common::custom_data::codec::Limits as XmlLimits;

pub fn parse_properties(xml: &[u8]) -> Result<Properties> {
    litchi_ooxml_common::custom_data::parse_properties(xml).map_err(Error::from)
}

/// Parse Custom Data Properties under an explicit XML resource profile.
pub fn parse_properties_with_limits(xml: &[u8], limits: &XmlLimits) -> Result<Properties> {
    litchi_ooxml_common::custom_data::parse_properties_with_limits(xml, limits).map_err(Error::from)
}

pub fn validate_workbook_root(xml: &[u8]) -> Result<()> {
    litchi_ooxml_common::custom_data::validate_workbook_root(xml).map_err(Error::from)
}

/// Validate a workbook root under an explicit XML resource profile.
pub fn validate_workbook_root_with_limits(xml: &[u8], limits: &XmlLimits) -> Result<()> {
    litchi_ooxml_common::custom_data::validate_workbook_root_with_limits(xml, limits)
        .map_err(Error::from)
}

pub fn write_properties(value: &Properties) -> Result<Vec<u8>> {
    litchi_ooxml_common::custom_data::write_properties(value).map_err(Error::from)
}

/// Serialize Custom Data Properties under an explicit XML resource profile.
pub fn write_properties_with_limits(value: &Properties, limits: &XmlLimits) -> Result<Vec<u8>> {
    litchi_ooxml_common::custom_data::write_properties_with_limits(value, limits)
        .map_err(Error::from)
}

pub(crate) fn validate_source_properties(value: &Properties) -> Result<()> {
    litchi_ooxml_common::custom_data::codec::validate_source_properties(value).map_err(Error::from)
}

pub(crate) fn rewrite_id(
    source: &litchi_opc::OwnedXmlPart,
    id: &str,
) -> Result<litchi_opc::OwnedXmlPart> {
    litchi_ooxml_common::custom_data::codec::rewrite_id(source, id).map_err(Error::from)
}

pub(crate) fn rewrite_extension_list(
    source: &litchi_opc::OwnedXmlPart,
    extension: Option<&ExtensionList>,
) -> Result<litchi_opc::OwnedXmlPart> {
    litchi_ooxml_common::custom_data::codec::rewrite_extension_list(source, extension)
        .map_err(Error::from)
}

pub(crate) fn canonical_extension(extension: Option<&ExtensionList>) -> Result<Option<Vec<u8>>> {
    litchi_ooxml_common::custom_data::codec::canonical_extension(extension).map_err(Error::from)
}
