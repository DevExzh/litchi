//! OPC relationship and part lifecycle for presentation properties.

use super::model::Properties;
use super::{CONTENT_TYPE, REL, STRICT_REL};
use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI};

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub fn load_from_package(package: &OpcPackage) -> Result<Option<Properties>> {
    let presentation = package.main_document_part()?;
    let mut found = presentation
        .rels()
        .iter()
        .filter(|r| matches!(r.reltype(), REL | STRICT_REL));
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "presentation has multiple presentation-properties relationships",
        ));
    }
    if rel.is_external() {
        return Err(invalid(
            "presentation-properties relationship cannot be external",
        ));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "presentation-properties part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    let mut value = Properties::parse(part.blob())?;
    if let Some(html) = value.html_publish.as_mut() {
        let target = part
            .rels()
            .get(&html.target.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "missing HTML publish relationship '{}'",
                    html.target.relationship_id
                ))
            })?;
        html.target.target = Some(target.target_ref().to_string());
        html.target.relationship_type = Some(target.reltype().to_string());
        html.target.external = Some(target.is_external());
    }
    Ok(Some(value))
}
