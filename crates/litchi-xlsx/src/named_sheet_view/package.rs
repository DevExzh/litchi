//! Worksheet relationship and OPC package operations for named-sheet-view parts.

use crate::error::Result;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships};
use std::collections::HashSet;

use super::codec::{parse_named_sheet_views, write_named_sheet_views};
use super::model::Views;
use super::{CONTENT_TYPE, RELATIONSHIP, content_type_mismatch, invalid, uri_error};

pub fn discover_named_sheet_views(
    package: &OpcPackage,
    relationships: &Relationships,
) -> Result<Option<Views>> {
    let mut found = relationships.iter().filter(|r| r.reltype() == RELATIONSHIP);
    let Some(relationship) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "worksheet has multiple Named Sheet Views relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("Named Sheet Views relationship cannot be external"));
    }
    let part = package.get_part(&relationship.target_partname()?)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(content_type_mismatch(CONTENT_TYPE, part.content_type()));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "Named Sheet Views part must not have relationships",
        ));
    }
    parse_named_sheet_views(part.blob()).map(Some)
}

/// Load the optional Named Sheet Views part owned by one worksheet.
///
/// Filters, sort metadata, and retained extensions are parsed only. This does
/// not apply a view, evaluate formulas, or fetch any external resource.
pub fn load_worksheet_named_sheet_views(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Option<Views>> {
    require_worksheet(package.get_part(worksheet_part)?)?;
    validate_named_sheet_views_graph(package)?;
    let Some(relationship) = named_sheet_views_relationship(package, worksheet_part)? else {
        return Ok(None);
    };
    let part = package.get_part(&relationship.part_name)?;
    parse_named_sheet_views(part.blob()).map(Some)
}

/// Store a parsed Named Sheet Views value as the worksheet's sole modern view
/// part.
///
/// The value is serialized as inert filter and sort metadata. Existing package
/// graph violations are rejected before any part is changed.
pub fn store_worksheet_named_sheet_views(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
    value: &Views,
) -> Result<()> {
    let xml = write_named_sheet_views(value)?;
    require_worksheet(package.get_part(worksheet_part)?)?;
    validate_named_sheet_views_graph(package)?;

    if let Some(relationship) = named_sheet_views_relationship(package, worksheet_part)? {
        package.get_part_mut(&relationship.part_name)?.set_blob(xml);
    } else {
        let part_name = next_named_sheet_views_part_name(package)?;
        let relationship_id = next_named_sheet_views_relationship_id(package, worksheet_part)?;
        let target = part_name.relative_ref(worksheet_part.base_uri());
        package.try_add_part(Box::new(BlobPart::new(part_name, CONTENT_TYPE.into(), xml)))?;
        package
            .get_part_mut(worksheet_part)?
            .rels_mut()
            .add_relationship(RELATIONSHIP.into(), target, relationship_id, false);
    }

    package.unsign();
    Ok(())
}

/// Remove a worksheet's Named Sheet Views relationship and unreferenced part.
///
/// The worksheet data and its ordinary active view are left unchanged.
pub fn remove_worksheet_named_sheet_views(
    package: &mut OpcPackage,
    worksheet_part: &PackURI,
) -> Result<bool> {
    require_worksheet(package.get_part(worksheet_part)?)?;
    validate_named_sheet_views_graph(package)?;
    let Some(relationship) = named_sheet_views_relationship(package, worksheet_part)? else {
        return Ok(false);
    };

    package
        .get_part_mut(worksheet_part)?
        .rels_mut()
        .remove(&relationship.relationship_id);
    if !package_part_is_referenced(package, &relationship.part_name) {
        package.remove_part(&relationship.part_name);
    }
    package.unsign();
    Ok(true)
}

#[derive(Debug, Clone)]
struct Relationship {
    relationship_id: String,
    part_name: PackURI,
}

fn named_sheet_views_relationship(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<Option<Relationship>> {
    let worksheet = package.get_part(worksheet_part)?;
    require_worksheet(worksheet)?;
    let mut matches = worksheet
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == RELATIONSHIP);
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(invalid(
            "worksheet has multiple Named Sheet Views relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("Named Sheet Views relationship cannot be external"));
    }
    let part_name = relationship.target_partname()?;
    validate_named_sheet_views_part(package, &part_name)?;
    Ok(Some(Relationship {
        relationship_id: relationship.r_id().to_owned(),
        part_name,
    }))
}

fn validate_named_sheet_views_graph(package: &OpcPackage) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == RELATIONSHIP)
    {
        return Err(invalid(
            "package root cannot source a Named Sheet Views relationship",
        ));
    }

    let mut targets = HashSet::new();
    for source in package.iter_parts() {
        let mut relationships = source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == RELATIONSHIP);
        let Some(relationship) = relationships.next() else {
            continue;
        };
        require_worksheet(source)?;
        if relationships.next().is_some() {
            return Err(invalid(format!(
                "worksheet '{}' has multiple Named Sheet Views relationships",
                source.partname()
            )));
        }
        if relationship.is_external() {
            return Err(invalid("Named Sheet Views relationship cannot be external"));
        }
        let target = relationship.target_partname()?;
        if !targets.insert(target.clone()) {
            return Err(invalid(format!(
                "Named Sheet Views part '{target}' is targeted more than once"
            )));
        }
        validate_named_sheet_views_part(package, &target)?;
    }

    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
    {
        if !targets.contains(part.partname()) {
            return Err(invalid(format!(
                "Named Sheet Views part '{}' has no worksheet relationship",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn validate_named_sheet_views_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(content_type_mismatch(CONTENT_TYPE, part.content_type()));
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!(
            "Named Sheet Views part '{part_name}' must not have relationships"
        )));
    }
    Ok(())
}

fn next_named_sheet_views_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=65_536u32 {
        let candidate = PackURI::new(format!("/xl/namedSheetViews/namedSheetView{suffix}.xml"))
            .map_err(uri_error)?;
        if package
            .iter_parts()
            .all(|part| part.partname() != &candidate)
        {
            return Ok(candidate);
        }
    }
    Err(invalid("no free Named Sheet Views part name"))
}

fn next_named_sheet_views_relationship_id(
    package: &OpcPackage,
    worksheet_part: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(worksheet_part)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdNamedSheetView{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free Named Sheet Views relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|part_name| part_name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|part_name| part_name == *target)
    })
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == ct::SML_WORKSHEET {
        Ok(())
    } else {
        Err(invalid(format!(
            "part '{}' is not a worksheet",
            part.partname()
        )))
    }
}
