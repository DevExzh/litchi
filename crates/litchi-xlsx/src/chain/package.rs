//! OPC package ownership for the workbook calculation-chain part.

use crate::error::{Result, invalid};
use litchi_opc::{OpcPackage, PackURI};

use super::codec::{read, write};
use super::model::{CONTENT_TYPE, Chain, Conformance, RELATIONSHIP, STRICT_RELATIONSHIP};

/// Load the optional inert calculation chain and its relationship conformance.
/// Formula cells are parsed as metadata only; no formula is evaluated.
pub fn load(package: &OpcPackage) -> Result<Option<(Chain, Conformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Validate the package topology for the optional calculation chain without
/// decoding its inert XML payload.
pub(crate) fn validate_package(package: &OpcPackage) -> Result<()> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(relationship) = relationship(package, &workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(());
    };
    validate_part_set(package, Some(&relationship.part_name))?;
    validate_part(package, &relationship.part_name)
}

/// Store a caller-authored inert calculation chain in a SpreadsheetML package.
///
/// The supplied order is serialized without recalculating formulas or inferring
/// dependencies. Existing calculation-chain graph violations are rejected
/// before any package part is changed. The requested conformance is applied to
/// both the part XML and its workbook relationship.
pub fn put(package: &mut OpcPackage, chain: &Chain, conformance: Conformance) -> Result<bool> {
    let xml = write(chain, conformance)?;
    let workbook_uri = main_workbook_uri(package)?;
    let existing = relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_part_set(package, Some(&existing.part_name))?;
        validate_part(package, &existing.part_name)?;
        let bytes_changed = package.get_part(&existing.part_name)?.blob() != xml;
        let relationship_changed = existing.conformance != conformance;
        if !bytes_changed && !relationship_changed {
            return Ok(false);
        }
        if bytes_changed {
            package.get_part_mut(&existing.part_name)?.set_blob(xml);
        }
        if relationship_changed {
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
        validate_part_set(package, None)?;
        let part_name = next_part_name(package)?;
        let relationship_id = next_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CONTENT_TYPE.into(),
            xml,
        )))?;
        let workbook = match package.get_part_mut(&workbook_uri) {
            Ok(workbook) => workbook,
            Err(error) => {
                package.remove_part(&part_name);
                return Err(error.into());
            },
        };
        workbook.rels_mut().add_relationship(
            conformance.relationship_type().into(),
            target,
            relationship_id,
            false,
        );
    }

    package.unsign();
    Ok(true)
}

/// Remove the workbook's calculation-chain relationship and its unreferenced part.
///
/// No formulas are changed. A target that is also referenced elsewhere in the
/// package is retained.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = relationship(package, &workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(false);
    };
    validate_part_set(package, Some(&existing.part_name))?;
    validate_part(package, &existing.part_name)?;
    let retain_part = part_is_referenced_elsewhere(
        package,
        &existing.part_name,
        &workbook_uri,
        &existing.relationship_id,
    )?;

    package
        .get_part_mut(&workbook_uri)?
        .rels_mut()
        .remove(&existing.relationship_id);
    if !retain_part {
        package.remove_part(&existing.part_name);
    }
    package.unsign();
    Ok(true)
}

fn load_for_workbook(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<(Chain, Conformance)>> {
    let Some(relationship) = relationship(package, workbook_uri)? else {
        validate_part_set(package, None)?;
        return Ok(None);
    };
    validate_part_set(package, Some(&relationship.part_name))?;
    validate_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((read(part.blob())?, relationship.conformance)))
}

#[derive(Debug, Clone)]
struct Relationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: Conformance,
}

fn relationship(package: &OpcPackage, workbook_uri: &PackURI) -> Result<Option<Relationship>> {
    let workbook = package.get_part(workbook_uri)?;
    let mut relationships = workbook.rels().iter().filter(|relationship| {
        matches!(relationship.reltype(), RELATIONSHIP | STRICT_RELATIONSHIP)
    });
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let conformance = if relationship.reltype() == RELATIONSHIP {
        Conformance::Transitional
    } else {
        Conformance::Strict
    };
    Ok(Some(Relationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "calculation-chain part '{part_name}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("calculation-chain part cannot have relationships"));
    }
    Ok(())
}

fn validate_part_set(package: &OpcPackage, relationship_target: Option<&PackURI>) -> Result<()> {
    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let part_name = parts.next().map(|part| part.partname());
    if parts.next().is_some() {
        return Err(invalid(
            "package contains more than one calculation-chain part",
        ));
    }
    match (relationship_target, part_name) {
        (None, None) => Ok(()),
        (None, Some(_)) => Err(invalid(
            "package contains a calculation-chain part without a workbook relationship",
        )),
        (Some(_), None) => Ok(()),
        (Some(target), Some(part_name)) if part_name == target => Ok(()),
        (Some(_), Some(_)) => Err(invalid(
            "workbook calculation-chain relationship does not target the calculation-chain part",
        )),
    }
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

fn next_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/calcChain.xml".to_string()
        } else {
            format!("/xl/calcChain{suffix}.xml")
        };
        let candidate = PackURI::new(&name).map_err(invalid)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain part name"))
}

fn next_relationship_id(package: &OpcPackage, workbook_uri: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdCalcChain{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free calculation-chain relationship ID"))
}

fn part_is_referenced_elsewhere(
    package: &OpcPackage,
    target: &PackURI,
    owner: &PackURI,
    owner_relationship: &str,
) -> Result<bool> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if part.partname() == owner && relationship.r_id() == owner_relationship {
                continue;
            }
            if !relationship.is_external() && relationship.target_partname()? == *target {
                return Ok(true);
            }
        }
    }
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()? == *target {
            return Ok(true);
        }
    }
    Ok(false)
}
