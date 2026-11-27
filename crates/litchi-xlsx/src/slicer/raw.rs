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

pub(crate) struct PreparedCache {
    pub(crate) value: crate::slicer_cache::Cache,
    pub(crate) uri: litchi_opc::PackURI,
    pub(crate) xml: Vec<u8>,
}

pub(crate) struct PreparedPart {
    pub(crate) value: crate::slicer_cache::views::SlicerPart,
    pub(crate) uri: litchi_opc::PackURI,
    pub(crate) xml: Vec<u8>,
}

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

pub(crate) fn prepare_cache(
    value: Cache,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<PreparedCache> {
    let uri = validate_cache_identity(&value, owner_base_uri, limits)?;
    let xml = super::codec::write_limited(&value.definition, crate::raw::output_ceiling(limits))?;
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
    let part_uri = litchi_opc::PackURI::new(part_name).map_err(|error| {
        crate::Error::Invalid(format!("invalid Slicer Cache part URI: {error}"))
    })?;
    if !part_uri.as_str().starts_with("/xl/slicerCaches/") || !part_uri.as_str().ends_with(".xml") {
        return Err(crate::Error::Invalid(
            "Slicer Cache part must be under /xl/slicerCaches and end in .xml".into(),
        ));
    }
    crate::raw::check_part_member_name(&part_uri, limits)?;
    crate::raw::check_relationship_target(&part_uri, owner_base_uri, limits)?;
    Ok(part_uri)
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

pub(crate) fn prepare_part(
    value: Part,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<PreparedPart> {
    let uri = validate_part_identity(&value, owner_base_uri, limits)?;
    let xml =
        super::codec::write_views_limited(&value.slicers, crate::raw::output_ceiling(limits))?;
    crate::raw::check_output_bytes(xml.len(), limits)?;
    Ok(PreparedPart {
        value: value.into_model(),
        uri,
        xml,
    })
}

pub(crate) fn validate_part_identity(
    value: &Part,
    owner_base_uri: &str,
    limits: &SourceLimits,
) -> Result<litchi_opc::PackURI> {
    validate_part_identity_fields(
        &value.relationship_id,
        &value.part_name,
        owner_base_uri,
        limits,
    )
}

pub(crate) fn validate_part_identity_fields(
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
        .map_err(|error| crate::Error::Invalid(format!("invalid Slicers part URI: {error}")))?;
    if !part_uri.as_str().starts_with("/xl/slicers/") || !part_uri.as_str().ends_with(".xml") {
        return Err(crate::Error::Invalid(
            "Slicers part must be under /xl/slicers and end in .xml".into(),
        ));
    }
    crate::raw::check_part_member_name(&part_uri, limits)?;
    crate::raw::check_relationship_target(&part_uri, owner_base_uri, limits)?;
    Ok(part_uri)
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

#[cfg(test)]
mod tests {
    use super::validate_cache_identity_fields;
    use litchi_opc::{PackURI, SourceLimits};

    #[test]
    fn cache_identity_does_not_charge_part_name_to_payload_limit() {
        let part_name = format!(
            "/xl/slicerCaches/{}.xml",
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
