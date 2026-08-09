//! Workbook relationship services for Custom XML Maps.

use super::invalid;
use super::model::{CONTENT_TYPE, MAX_PART_BYTES, REL, STRICT_REL, XmlMapConformance, XmlMapInfo};
use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};

/// Discovers and parses the single Custom XML Maps part related to the workbook.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<XmlMapInfo>> {
    Ok(load_from_package_with_conformance(package)?.map(|(value, _)| value))
}

/// Discovers the workbook's Custom XML Maps part together with its namespace family.
///
/// The schema payload and data-binding payloads remain opaque. This function does
/// not resolve schema locations, open bound files, or import/export mapped data.
pub fn load_from_package_with_conformance(
    package: &OpcPackage,
) -> Result<Option<(XmlMapInfo, XmlMapConformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Store caller-authored Custom XML Maps metadata in a `SpreadsheetML` package.
///
/// Existing malformed XML Maps relationships are rejected before mutation. The
/// writer never resolves inline schema references, opens bound files, or applies
/// a mapping to worksheet cells.
pub fn store_in_package(
    package: &mut OpcPackage,
    value: &XmlMapInfo,
    conformance: XmlMapConformance,
) -> Result<()> {
    if let Some((existing, existing_conformance)) = load_from_package_with_conformance(package)?
        && existing_conformance == conformance
        && existing == *value
    {
        // Retain a byte-identical source-loaded XML Maps part when the caller
        // requests no semantic or conformance change. This is the package
        // writer's deliberately narrow source-preservation exemption and also
        // avoids rewriting opaque inline schema payloads.
        return Ok(());
    }
    let xml = value.to_xml(conformance.is_strict())?;
    store_xml_in_package(package, &xml, conformance)
}

/// Publish already-validated XML for a Custom XML Maps part.
///
/// Transactions use this narrow service after patching the original source
/// bytes. Keeping the package graph orchestration here means source edits and
/// ordinary callers share the same bounded relationship validation.
pub(super) fn store_xml_in_package(
    package: &mut OpcPackage,
    xml: &[u8],
    conformance: XmlMapConformance,
) -> Result<()> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("custom XML maps part exceeds 32 MiB"));
    }
    let workbook_uri = main_workbook_uri(package)?;
    let existing = xml_maps_relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_xml_maps_graph(package, &workbook_uri, Some(&existing))?;
        validate_xml_maps_part(package, &existing.part_name)?;
        package
            .get_part_mut(&existing.part_name)?
            .set_blob(xml.to_vec());
        if existing.conformance != conformance {
            let workbook = package.get_part_mut(&workbook_uri)?;
            workbook.rels_mut().remove(&existing.relationship_id);
            workbook.rels_mut().add_relationship(
                conformance.relationship_type().into(),
                existing.target_reference,
                existing.relationship_id,
                false,
            );
        }
    } else {
        validate_xml_maps_graph(package, &workbook_uri, None)?;
        let part_name = next_xml_maps_part_name(package)?;
        let relationship_id = next_xml_maps_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name,
            CONTENT_TYPE.into(),
            xml.to_vec(),
        )))?;
        package
            .get_part_mut(&workbook_uri)?
            .rels_mut()
            .add_relationship(
                conformance.relationship_type().into(),
                target,
                relationship_id,
                false,
            );
    }

    package.unsign();
    Ok(())
}

/// Remove the workbook's Custom XML Maps relationship and its unreferenced part.
///
/// No mapping is applied to worksheet data. A target that remains referenced by
/// another package part is retained.
pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = xml_maps_relationship(package, &workbook_uri)? else {
        validate_xml_maps_graph(package, &workbook_uri, None)?;
        return Ok(false);
    };
    validate_xml_maps_graph(package, &workbook_uri, Some(&existing))?;
    validate_xml_maps_part(package, &existing.part_name)?;

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !package_part_is_referenced(package, &existing.part_name) {
        package.remove_part(&existing.part_name);
    }
    package.unsign();
    Ok(true)
}

fn load_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(XmlMapInfo, XmlMapConformance)>> {
    let Some(relationship) = xml_maps_relationship(package, workbook_uri)? else {
        validate_xml_maps_graph(package, workbook_uri, None)?;
        return Ok(None);
    };
    validate_xml_maps_graph(package, workbook_uri, Some(&relationship))?;
    validate_xml_maps_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((
        XmlMapInfo::parse(part.blob())?,
        relationship.conformance,
    )))
}

#[derive(Clone, Debug)]
struct XmlMapsRelationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: XmlMapConformance,
}

fn xml_maps_relationship(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<XmlMapsRelationship>> {
    let workbook = package.get_part(workbook_uri)?;
    let mut relationships = workbook
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook has multiple custom XML maps relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("custom XML maps relationship cannot be external"));
    }
    let conformance = if relationship.reltype() == REL {
        XmlMapConformance::Transitional
    } else {
        XmlMapConformance::Strict
    };
    Ok(Some(XmlMapsRelationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_xml_maps_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "custom XML maps part '{part_name}' has content type '{}', expected '{CONTENT_TYPE}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("custom XML maps part must not have relationships"));
    }
    Ok(())
}

fn validate_xml_maps_graph(
    package: &OpcPackage,
    workbook_uri: &PackURI,
    expected: Option<&XmlMapsRelationship>,
) -> Result<()> {
    let mut found = 0usize;
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            if part.partname() != workbook_uri {
                return Err(invalid(
                    "custom XML maps relationships may only originate from the workbook",
                ));
            }
            if relationship.is_external() {
                return Err(invalid("custom XML maps relationship cannot be external"));
            }
            let target = relationship.target_partname()?;
            let Some(expected) = expected else {
                return Err(invalid(
                    "workbook has an unexpected custom XML maps relationship",
                ));
            };
            if relationship.r_id() != expected.relationship_id || target != expected.part_name {
                return Err(invalid(
                    "custom XML maps relationship graph is inconsistent",
                ));
            }
            found += 1;
        }
    }
    if package
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
    {
        return Err(invalid(
            "custom XML maps relationships may not originate from the package root",
        ));
    }
    match (expected, found) {
        (None, 0) | (Some(_), 1) => {},
        (None, _) => {
            return Err(invalid(
                "workbook has an unexpected custom XML maps relationship",
            ));
        },
        (Some(_), _) => {
            return Err(invalid(
                "workbook custom XML maps relationship graph is incomplete",
            ));
        },
    }
    Ok(())
}

fn main_workbook_uri(package: &OpcPackage) -> Result<PackURI> {
    use litchi_opc::constants::content_type as ct;

    let workbook = package.main_document_part()?;
    if !matches!(
        workbook.content_type(),
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid(format!(
            "main document part '{}' is not an XML workbook",
            workbook.partname()
        )));
    }
    Ok(workbook.partname().clone())
}

fn next_xml_maps_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/xmlMaps.xml".to_string()
        } else {
            format!("/xl/xmlMaps{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free custom XML maps part name"))
}

fn next_xml_maps_relationship_id(package: &OpcPackage, workbook_uri: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdXmlMaps{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free custom XML maps relationship ID"))
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
