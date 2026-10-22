//! Native-identity values for typed timeline graph edits.
//!
//! The ordinary timeline model generates relationship IDs and part names.
//! This module is the explicit raw escape for callers that must preserve
//! those physical identities. Its values are accepted only by the
//! package-scoped [`super::Edit`] insertion methods, which still validate and
//! publish the complete XLSX package through the ordinary edit commit.

use super::{CacheDefinition, Views};
use crate::error::Result;
use litchi_opc::SourceLimits;

/// A View Cache definition with caller-selected OPC identity.
#[derive(Debug, Clone, PartialEq)]
pub struct Cache {
    /// The workbook relationship ID that points at this cache.
    pub relationship_id: String,
    /// The absolute OPC part name for this cache.
    pub part_name: String,
    /// The typed cache definition.
    pub definition: CacheDefinition,
}

impl Cache {
    pub(crate) fn into_model(self) -> crate::timelines::Cache {
        crate::timelines::Cache {
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
    let part_uri = litchi_opc::PackURI::new(&value.part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid View Cache part URI: {error}")))?;
    crate::raw::check_relationship_target(&part_uri, owner_base_uri, limits)?;
    let maximum = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
    let output = super::codec::write_cache_limited(&value.definition, maximum)?;
    crate::raw::check_output_bytes(output.len(), limits)?;
    Ok(())
}

/// A worksheet Timelines part with caller-selected worksheet and OPC identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Part {
    /// The worksheet part that owns the relationship.
    pub worksheet_part_name: String,
    /// The worksheet relationship ID that points at this part.
    pub relationship_id: String,
    /// The absolute OPC part name for this worksheet Timelines part.
    pub part_name: String,
    /// The typed Timelines XML value.
    pub timelines: Views,
}

impl Part {
    pub(crate) fn into_model(self) -> crate::timelines::Part {
        crate::timelines::Part {
            worksheet_part_name: self.worksheet_part_name,
            relationship_id: self.relationship_id,
            part_name: self.part_name,
            timelines: self.timelines,
        }
    }
}

pub(crate) fn preflight_part(value: &Part, limits: &SourceLimits) -> Result<()> {
    crate::raw::check_source_bytes(value.worksheet_part_name.len(), limits)?;
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::XmlAttributeBytes,
        value.relationship_id.len() as u64,
        limits.max_xml_attribute_bytes() as u64,
    )?;
    crate::raw::check_source_bytes(value.part_name.len(), limits)?;
    let part_uri = litchi_opc::PackURI::new(&value.part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid Timelines part URI: {error}")))?;
    let worksheet_uri = litchi_opc::PackURI::new(&value.worksheet_part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid worksheet part URI: {error}")))?;
    crate::raw::check_relationship_target(&part_uri, worksheet_uri.base_uri(), limits)?;
    let maximum = usize::try_from(limits.max_part_bytes()).unwrap_or(usize::MAX);
    let output = super::codec::write_views_limited(&value.timelines, maximum)?;
    crate::raw::check_output_bytes(output.len(), limits)?;
    Ok(())
}

/// Parse one native View Cache definition under the caller's raw limits.
pub fn parse_cache(xml: &[u8], limits: &SourceLimits) -> Result<CacheDefinition> {
    crate::raw::check_xml_limits(xml, limits)?;
    super::codec::read_cache(xml)
}

/// Parse one native worksheet Timelines part under the caller's raw limits.
pub fn parse_part(xml: &[u8], limits: &SourceLimits) -> Result<Views> {
    crate::raw::check_xml_limits(xml, limits)?;
    super::codec::read_views(xml)
}
