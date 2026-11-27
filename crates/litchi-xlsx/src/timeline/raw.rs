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

pub(crate) struct PreparedCache {
    pub(crate) value: crate::timelines::Cache,
    pub(crate) uri: litchi_opc::PackURI,
    pub(crate) xml: Vec<u8>,
}

pub(crate) struct PreparedPart {
    pub(crate) value: crate::timelines::Part,
    pub(crate) uri: litchi_opc::PackURI,
    pub(crate) worksheet: litchi_opc::PackURI,
    pub(crate) xml: Vec<u8>,
}

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

pub(crate) fn prepare_cache(
    value: Cache,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<PreparedCache> {
    let uri = validate_cache_identity(&value, owner_base_uri, limits)?;
    let xml =
        super::codec::write_cache_limited(&value.definition, crate::raw::output_ceiling(limits))?;
    crate::raw::check_output_bytes(xml.len(), limits)?;
    Ok(PreparedCache {
        value: value.into_model(),
        uri,
        xml,
    })
}

pub(crate) fn validate_cache_identity(
    value: &Cache,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<litchi_opc::PackURI> {
    validate_cache_identity_fields(
        &value.relationship_id,
        &value.part_name,
        owner_base_uri,
        limits,
    )
}

pub(crate) fn validate_cache_identity_fields(
    relationship_id: &str,
    part_name: &str,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<litchi_opc::PackURI> {
    crate::raw::validate_relationship_id(relationship_id)?;
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::XmlAttributeBytes,
        relationship_id.len() as u64,
        limits.max_xml_attribute_bytes() as u64,
    )?;
    let part_uri = litchi_opc::PackURI::new(part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid View Cache part URI: {error}")))?;
    if !part_uri.as_str().starts_with("/xl/timelineCaches/") || !part_uri.as_str().ends_with(".xml")
    {
        return Err(crate::Error::Invalid(
            "View Cache part must be under /xl/timelineCaches and end in .xml".into(),
        ));
    }
    crate::raw::check_part_member_name(&part_uri, limits)?;
    crate::raw::check_relationship_target(&part_uri, owner_base_uri, limits)?;
    Ok(part_uri)
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

pub(crate) fn prepare_part(value: Part, limits: &SourceLimits) -> Result<PreparedPart> {
    let (uri, worksheet) = validate_part_identity(&value, limits)?;
    let xml =
        super::codec::write_views_limited(&value.timelines, crate::raw::output_ceiling(limits))?;
    crate::raw::check_output_bytes(xml.len(), limits)?;
    Ok(PreparedPart {
        value: value.into_model(),
        uri,
        worksheet,
        xml,
    })
}

pub(crate) fn validate_part_identity(
    value: &Part,
    limits: &SourceLimits,
) -> Result<(litchi_opc::PackURI, litchi_opc::PackURI)> {
    validate_part_identity_fields(
        &value.worksheet_part_name,
        &value.relationship_id,
        &value.part_name,
        limits,
    )
}

pub(crate) fn validate_part_identity_fields(
    worksheet_part_name: &str,
    relationship_id: &str,
    part_name: &str,
    limits: &SourceLimits,
) -> Result<(litchi_opc::PackURI, litchi_opc::PackURI)> {
    crate::raw::validate_relationship_id(relationship_id)?;
    crate::raw::check_limit(
        litchi_opc::SourceLimitKind::XmlAttributeBytes,
        relationship_id.len() as u64,
        limits.max_xml_attribute_bytes() as u64,
    )?;
    let part_uri = litchi_opc::PackURI::new(part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid Timelines part URI: {error}")))?;
    if !part_uri.as_str().starts_with("/xl/timelines/") || !part_uri.as_str().ends_with(".xml") {
        return Err(crate::Error::Invalid(
            "Timelines part must be under /xl/timelines and end in .xml".into(),
        ));
    }
    crate::raw::check_part_member_name(&part_uri, limits)?;
    let worksheet = litchi_opc::PackURI::new(worksheet_part_name)
        .map_err(|error| crate::Error::Invalid(format!("invalid worksheet part URI: {error}")))?;
    crate::raw::check_relationship_target(&part_uri, worksheet.base_uri(), limits)?;
    Ok((part_uri, worksheet))
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

#[cfg(test)]
mod tests {
    use super::validate_cache_identity_fields;
    use litchi_opc::{PackURI, SourceLimits};

    #[test]
    fn cache_identity_does_not_charge_part_name_to_payload_limit() {
        let part_name = format!(
            "/xl/timelineCaches/{}.xml",
            "long-native-identity-".repeat(64)
        );
        let uri = PackURI::new(&part_name).unwrap();
        let limits = SourceLimits::builder()
            .max_part_bytes(1)
            .unwrap()
            .max_archive_member_name_bytes(uri.membername().len() as u64)
            .unwrap()
            .max_relationship_target_bytes(uri.membername().len())
            .unwrap()
            .build()
            .unwrap();
        validate_cache_identity_fields("rIdNativeIdentity", &part_name, "/xl/", &limits).unwrap();
    }
}
