//! Presentation package relationship loading for view properties.

use super::model::ViewProperties;
use crate::{Error, Result};
use litchi_opc::{OpcPackage, PackURI};

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps";
const CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml";

/// Loads the presentation view-properties part through its package relationship.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<ViewProperties>> {
    let presentation = package.main_document_part()?;
    let mut found = presentation
        .rels()
        .iter()
        .filter(|x| matches!(x.reltype(), REL | STRICT_REL));
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "presentation has multiple view-properties relationships",
        ));
    }
    if rel.is_external() {
        return Err(invalid("view-properties relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CT {
        return Err(invalid(format!(
            "view-properties part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    let mut value = ViewProperties::parse(part.blob())?;
    if let Some(outline) = value.outline.as_mut() {
        for slide in &mut outline.slides {
            let relationship = part.rels().get(&slide.relationship_id).ok_or_else(|| {
                invalid(format!(
                    "missing outline slide relationship '{}'",
                    slide.relationship_id
                ))
            })?;
            if relationship.is_external() {
                return Err(invalid("outline slide relationship cannot be external"));
            }
            slide.target = Some(relationship.target_ref().to_string());
        }
    }
    Ok(Some(value))
}

fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}
