//! Native-identity values for typed slicer graph edits.
//!
//! The ordinary slicer model generates relationship IDs and part names. This
//! module is the explicit raw escape for callers that must preserve those
//! physical identities. Its values are accepted only by the package-scoped
//! [`super::Edit`] insertion methods, which still validate and publish the
//! complete XLSX package through the ordinary edit commit.

use super::{Definition, Slicers};
use crate::error::Result;
use litchi_opc::SourceLimits;

/// A Slicer Cache definition with caller-selected OPC identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Cache {
    /// The workbook relationship ID that points at this cache.
    pub relationship_id: String,
    /// The absolute OPC part name for this cache.
    pub part_name: String,
    /// The typed cache definition.
    pub definition: Definition,
}

impl Cache {
    pub(crate) fn into_model(self) -> crate::slicer_cache::Cache {
        crate::slicer_cache::Cache {
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            definition: self.definition,
        }
    }
}

pub(crate) fn preflight_cache(
    value: &Cache,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<()> {
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::XmlAttributeBytes,
        value.relationship_id.len() as u64,
        limits.max_xml_attribute_bytes() as u64,
    )?;
    crate::raw::check_source_bytes(value.part_name.len(), limits)?;
    let part_uri = litchi_opc::PackURI::new(&value.part_name).map_err(|error| {
        crate::Error::Invalid(format!("invalid Slicer Cache part URI: {error}"))
    })?;
    crate::raw::check_relationship_target(&part_uri, owner_base_uri, limits)?;
    let maximum = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
    let output = super::codec::write_limited(&value.definition, maximum)?;
    crate::raw::check_output_bytes(output.len(), limits)?;
    Ok(())
}

/// A worksheet Slicers part with caller-selected OPC identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    /// The worksheet relationship ID that points at this part.
    pub relationship_id: String,
    /// The absolute OPC part name for this worksheet Slicers part.
    pub part_name: String,
    /// The typed Slicers XML value.
    pub slicers: Slicers,
}

impl Part {
    pub(crate) fn into_model(self) -> crate::slicer_cache::views::SlicerPart {
        crate::slicer_cache::views::SlicerPart {
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            slicers: self.slicers,
        }
    }
}

pub(crate) fn preflight_part(
    value: &Part,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<()> {
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::XmlAttributeBytes,
        value.relationship_id.len() as u64,
        limits.max_xml_attribute_bytes() as u64,
    )?;
    crate::raw::check_source_bytes(value.part_name.len(), limits)?;
    let part_uri = litchi_opc::PackURI::new(&value.part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid Slicers part URI: {error}")))?;
    crate::raw::check_relationship_target(&part_uri, owner_base_uri, limits)?;
    let maximum = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
    let output = super::codec::write_views_limited(&value.slicers, maximum)?;
    crate::raw::check_output_bytes(output.len(), limits)?;
    Ok(())
}

/// Parse one native Slicer Cache definition under the caller's raw limits.
pub fn parse_cache(xml: &[u8], limits: &SourceLimits) -> Result<Definition> {
    crate::raw::check_xml_limits(xml, limits)?;
    super::codec::read(xml)
}

/// Parse one native worksheet Slicers part under the caller's raw limits.
pub fn parse_part(xml: &[u8], limits: &SourceLimits) -> Result<Slicers> {
    crate::raw::check_xml_limits(xml, limits)?;
    super::codec::read_views(xml)
}
