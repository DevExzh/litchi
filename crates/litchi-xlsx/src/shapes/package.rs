//! Package relationship services for worksheet shape inventories.
//!
//! This layer resolves worksheet and drawing parts only. It never follows OLE
//! payloads, external links, or chart/image targets.

use crate::error::{Error, Result};
use crate::raw::Sheet;
use crate::raw::parse_catalog;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};

use super::codec::parse_drawing_shapes;
use super::model::*;

/// Load the shape inventory of every worksheet in a workbook package.
///
/// One entry is returned per worksheet that anchors at least one shape-like
/// object; worksheets without shapes are omitted.
pub fn load_shapes(package: &OpcPackage) -> Result<Vec<Shapes>> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let mut output = Vec::new();
    for sheet in &sheets {
        let shapes = load_shapes_for_sheet(package, workbook_part, sheet)?;
        if !shapes.objects.is_empty() {
            output.push(shapes);
        }
    }
    Ok(output)
}

/// Load the shape inventory of one worksheet, addressed by sheet name.
pub fn load_sheet_shapes(package: &OpcPackage, sheet_name: &str) -> Result<Shapes> {
    let workbook_part = package.main_document_part()?;
    let sheets = parse_workbook_sheets(workbook_part.blob())?;
    let sheet = sheets
        .iter()
        .find(|sheet| sheet.name == sheet_name)
        .ok_or_else(|| invalid(format!("worksheet '{sheet_name}' not found")))?;
    load_shapes_for_sheet(package, workbook_part, sheet)
}

fn parse_workbook_sheets(xml: &[u8]) -> Result<Vec<Sheet>> {
    if xml.len() > MAX_WORKBOOK_BYTES {
        return Err(limit("workbook XML bytes"));
    }
    Ok(parse_catalog(xml)?.sheets)
}

fn load_shapes_for_sheet(
    package: &OpcPackage,
    workbook_part: &dyn Part,
    sheet: &Sheet,
) -> Result<Shapes> {
    let relationship = workbook_part
        .rels()
        .get(&sheet.relationship_id)
        .ok_or_else(|| {
            invalid(format!(
                "worksheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
    if !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET) {
        return Err(invalid(format!(
            "worksheet '{}' relationship has invalid type '{}'",
            sheet.name,
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(invalid(format!(
            "worksheet '{}' relationship cannot be external",
            sheet.name
        )));
    }
    let sheet_uri = relationship.target_partname()?;
    let sheet_part = package.get_part(&sheet_uri)?;
    if sheet_part.content_type() != ct::SML_WORKSHEET {
        return Err(invalid(format!(
            "worksheet part has content type '{}', expected '{}'",
            sheet_part.content_type(),
            ct::SML_WORKSHEET
        )));
    }
    let mut objects = Vec::new();
    let drawings: Vec<_> = sheet_part
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING))
        .collect();
    if drawings.len() > MAX_DRAWINGS_PER_WORKSHEET {
        return Err(limit("drawings per worksheet"));
    }
    for drawing_relationship in drawings {
        if drawing_relationship.is_external() {
            return Err(invalid("worksheet drawing relationship cannot be external"));
        }
        let drawing_uri = drawing_relationship.target_partname()?;
        let drawing_part = package.get_part(&drawing_uri)?;
        if drawing_part.content_type() != ct::OFC_DRAWING {
            return Err(invalid(format!(
                "drawing part has content type '{}', expected '{}'",
                drawing_part.content_type(),
                ct::OFC_DRAWING
            )));
        }
        if drawing_part.blob().len() > MAX_DRAWING_PART_BYTES {
            return Err(limit("drawing part bytes"));
        }
        let drawing_xml = std::str::from_utf8(drawing_part.blob())
            .map_err(|error| Error::Invalid(error.to_string()))?;
        let Some(anchored) = parse_drawing_shapes(drawing_xml)? else {
            continue;
        };
        for object in anchored {
            if objects.len() >= MAX_ANCHORS_PER_DRAWING {
                return Err(limit("shapes per worksheet"));
            }
            objects.push(object);
        }
    }
    Ok(Shapes {
        worksheet_name: sheet.name.clone(),
        worksheet_part_name: sheet_uri.to_string(),
        objects,
    })
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(name: &str) -> Error {
    invalid(format!("XLSX drawing shape {name} limit exceeded"))
}
