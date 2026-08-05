//! ChartEx OPC part boundary and relationship validation.

use super::super::super::style::package::discover;
use super::super::model::*;
use super::super::package::Part;
use super::CONTENT_TYPE;
use super::limits::*;
use super::xml::*;
use crate::{Error, Result};
use litchi_opc::OpcPackage;
use litchi_opc::part::Part as OpcPart;

impl<'a> Part<'a> {
    pub fn from_part(part: &'a dyn OpcPart) -> Result<Self> {
        if part.content_type() != CONTENT_TYPE {
            return invalid("ChartEx part has the wrong content type");
        }
        Ok(Self { part })
    }

    pub fn parse(&self) -> Result<Document> {
        parse_document(self.part.blob())
    }

    /// Parse and validate referenced package resources without opening their bytes.
    pub fn parse_in_package(&self, package: &OpcPackage) -> Result<Document> {
        let mut document = self.parse()?;
        if let Some(external) = &document.info.external_data {
            if external.auto_update {
                return invalid("auto-updating  external data is rejected");
            }
            document.external_data_target = Some(validate_external_data(
                package,
                self.part,
                &external.relationship_id,
            )?);
        }
        if let Some(id) = &document.info.fallback_image_relationship_id {
            document.fallback_image_part_name =
                Some(validate_fallback_image(package, self.part, id)?);
        }
        let (chart_style, chart_color_style) = discover(package, self.part)?;
        document.chart_style = chart_style;
        document.chart_color_style = chart_color_style;
        Ok(document)
    }

    pub fn part(&self) -> &'a dyn OpcPart {
        self.part
    }
}

fn validate_external_data(
    package: &OpcPackage,
    part: &dyn OpcPart,
    id: &str,
) -> Result<ExternalDataTarget> {
    let relationship = internal_relationship(part, id)?;
    reject_target(relationship.target_ref())?;
    let target = relationship.target_partname().map_err(Error::Opc)?;
    if !target.as_str().starts_with("/ppt/embeddings/") || target.as_str().ends_with('/') {
        return invalid(" external data escapes /ppt/embeddings/");
    }
    let target_part = package
        .get_part(&target)
        .map_err(|_| invalid_error(" external data target is missing"))?;
    if PACKAGE_REL.contains(&relationship.reltype()) {
        if !WORKBOOK_CONTENT_TYPES.contains(&target_part.content_type()) {
            return invalid(" package relationship targets a non-workbook part");
        }
        Ok(ExternalDataTarget::EmbeddedPackage {
            part_name: target.as_str().to_owned(),
            content_type: target_part.content_type().to_owned(),
        })
    } else if OLE_REL.contains(&relationship.reltype()) {
        if target_part.content_type() != OLE_CONTENT_TYPE {
            return invalid(" OLE relationship has mismatched content type");
        }
        Ok(ExternalDataTarget::OleObject {
            part_name: target.as_str().to_owned(),
        })
    } else {
        invalid(" externalData relationship has the wrong type")
    }
}

fn validate_fallback_image(package: &OpcPackage, part: &dyn OpcPart, id: &str) -> Result<String> {
    let relationship = internal_relationship(part, id)?;
    if !IMAGE_REL.contains(&relationship.reltype()) {
        return invalid(" fallback image relationship has the wrong type");
    }
    reject_target(relationship.target_ref())?;
    let target = relationship.target_partname().map_err(Error::Opc)?;
    if !target.as_str().starts_with("/ppt/media/") || target.as_str().ends_with('/') {
        return invalid(" fallback image escapes /ppt/media/");
    }
    let target_part = package
        .get_part(&target)
        .map_err(|_| invalid_error(" fallback image target is missing"))?;
    if !target_part.content_type().starts_with("image/") {
        return invalid(" fallback image has a non-image content type");
    }
    Ok(target.as_str().to_owned())
}

fn internal_relationship<'a>(
    part: &'a dyn OpcPart,
    id: &str,
) -> Result<&'a litchi_opc::Relationship> {
    validate_id(id)?;
    let relationship = part
        .rels()
        .get(id)
        .ok_or_else(|| invalid_error(format!("missing  relationship '{id}'")))?;
    if relationship.is_external() {
        return invalid("external  relationships are not loaded");
    }
    Ok(relationship)
}

fn reject_target(value: &str) -> Result<()> {
    let lower = value.to_ascii_lowercase();
    if value.is_empty()
        || value.contains(['?', '#', '\\'])
        || lower.contains("%2e")
        || lower.contains("%2f")
        || lower.contains("%5c")
    {
        return invalid("ambiguous or encoded  relationship target");
    }
    Ok(())
}
