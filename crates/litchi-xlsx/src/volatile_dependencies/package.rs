//! Workbook relationship services for volatile-dependencies metadata.
//!
//! This layer owns only the inert OPC graph around the `SpreadsheetML` part.
//! It never contacts RTD servers, opens OLAP connections, or evaluates
//! formulas.

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};

use super::invalid;
use super::model::{
    CONTENT_TYPE, MAX_PART_BYTES, REL, STRICT_REL, VolatileDependencies,
    VolatileDependenciesConformance,
};

/// Loads the single volatile-dependencies part related to the package workbook.
pub fn load_from_package(package: &OpcPackage) -> Result<Option<VolatileDependencies>> {
    Ok(load_from_package_with_conformance(package)?.map(|(value, _)| value))
}

/// Load volatile-dependencies metadata with the XML/relationship namespace family.
///
/// Dependency records are metadata only: this never contacts RTD servers, opens
/// OLAP connections, or evaluates/recalculates workbook formulas.
pub fn load_from_package_with_conformance(
    package: &OpcPackage,
) -> Result<Option<(VolatileDependencies, VolatileDependenciesConformance)>> {
    let workbook_uri = main_workbook_uri(package)?;
    load_for_workbook(package, &workbook_uri)
}

/// Store caller-authored inert volatile-dependencies metadata in a `SpreadsheetML` package.
///
/// Existing invalid package graphs are rejected before mutation. The writer only
/// persists the supplied dependency records and never performs RTD, cube, or
/// formula evaluation work.
pub fn store_in_package(
    package: &mut OpcPackage,
    value: &VolatileDependencies,
    conformance: VolatileDependenciesConformance,
) -> Result<()> {
    let xml = value.to_xml(conformance.is_strict())?;
    let workbook_uri = main_workbook_uri(package)?;
    let existing = volatile_dependencies_relationship(package, &workbook_uri)?;

    if let Some(existing) = existing {
        validate_volatile_dependencies_graph(package, &workbook_uri, Some(&existing))?;
        validate_volatile_dependencies_part(package, &existing.part_name)?;
        package.get_part_mut(&existing.part_name)?.set_blob(xml);
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
        validate_volatile_dependencies_graph(package, &workbook_uri, None)?;
        let part_name = next_volatile_dependencies_part_name(package)?;
        let relationship_id = next_volatile_dependencies_relationship_id(package, &workbook_uri)?;
        let target = part_name.relative_ref(workbook_uri.base_uri());
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name,
            CONTENT_TYPE.into(),
            xml,
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

/// Remove the workbook volatile-dependencies relationship and its unreferenced part.
///
/// No RTD, cube, or formula work is performed. A target retained by another
/// relationship is left in the package.
pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    let workbook_uri = main_workbook_uri(package)?;
    let Some(existing) = volatile_dependencies_relationship(package, &workbook_uri)? else {
        validate_volatile_dependencies_graph(package, &workbook_uri, None)?;
        return Ok(false);
    };
    validate_volatile_dependencies_graph(package, &workbook_uri, Some(&existing))?;
    validate_volatile_dependencies_part(package, &existing.part_name)?;

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
) -> Result<Option<(VolatileDependencies, VolatileDependenciesConformance)>> {
    let Some(relationship) = volatile_dependencies_relationship(package, workbook_uri)? else {
        validate_volatile_dependencies_graph(package, workbook_uri, None)?;
        return Ok(None);
    };
    validate_volatile_dependencies_graph(package, workbook_uri, Some(&relationship))?;
    validate_volatile_dependencies_part(package, &relationship.part_name)?;
    let part = package.get_part(&relationship.part_name)?;
    Ok(Some((
        VolatileDependencies::parse(part.blob())?,
        relationship.conformance,
    )))
}

#[derive(Clone, Debug)]
struct VolatileDependenciesRelationship {
    relationship_id: String,
    part_name: PackURI,
    target_reference: String,
    conformance: VolatileDependenciesConformance,
}

fn volatile_dependencies_relationship(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<Option<VolatileDependenciesRelationship>> {
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
            "workbook has multiple volatile-dependencies relationships",
        ));
    }
    if relationship.is_external() {
        return Err(invalid(
            "volatile-dependencies relationship cannot be external",
        ));
    }
    let conformance = if relationship.reltype() == REL {
        VolatileDependenciesConformance::Transitional
    } else {
        VolatileDependenciesConformance::Strict
    };
    Ok(Some(VolatileDependenciesRelationship {
        relationship_id: relationship.r_id().to_string(),
        part_name: relationship.target_partname()?,
        target_reference: relationship.target_ref().to_string(),
        conformance,
    }))
}

fn validate_volatile_dependencies_part(package: &OpcPackage, part_name: &PackURI) -> Result<()> {
    let part = package.get_part(part_name)?;
    if part.content_type() != CONTENT_TYPE {
        return Err(invalid(format!(
            "volatile-dependencies part '{part_name}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid(
            "volatile-dependencies part must not have relationships",
        ));
    }
    if part.blob().len() > MAX_PART_BYTES {
        return Err(invalid("volatile-dependencies part exceeds 8 MiB"));
    }
    Ok(())
}

fn validate_volatile_dependencies_graph(
    package: &OpcPackage,
    workbook_uri: &PackURI,
    expected: Option<&VolatileDependenciesRelationship>,
) -> Result<()> {
    validate_volatile_dependencies_part_set(package, expected.map(|value| &value.part_name))?;

    let mut found = 0usize;
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            if part.partname() != workbook_uri {
                return Err(invalid(
                    "volatile-dependencies relationships may only originate from the workbook",
                ));
            }
            if relationship.is_external() {
                return Err(invalid(
                    "volatile-dependencies relationship cannot be external",
                ));
            }
            let target = relationship.target_partname()?;
            let Some(expected) = expected else {
                return Err(invalid(
                    "workbook has an unexpected volatile-dependencies relationship",
                ));
            };
            if relationship.r_id() != expected.relationship_id || target != expected.part_name {
                return Err(invalid(
                    "volatile-dependencies relationship graph is inconsistent",
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
            "volatile-dependencies relationships may not originate from the package root",
        ));
    }
    match (expected, found) {
        (None, 0) | (Some(_), 1) => Ok(()),
        (None, _) => Err(invalid(
            "workbook has an unexpected volatile-dependencies relationship",
        )),
        (Some(_), _) => Err(invalid(
            "workbook volatile-dependencies relationship graph is incomplete",
        )),
    }
}

fn validate_volatile_dependencies_part_set(
    package: &OpcPackage,
    relationship_target: Option<&PackURI>,
) -> Result<()> {
    let part_names = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE)
        .map(|part| part.partname().clone())
        .collect::<Vec<_>>();
    if part_names.len() > 1 {
        return Err(invalid(
            "package contains more than one volatile-dependencies part",
        ));
    }
    match (relationship_target, part_names.as_slice()) {
        (None, []) => Ok(()),
        (None, _) => Err(invalid(
            "package contains a volatile-dependencies part without a workbook relationship",
        )),
        (Some(_), []) => Err(invalid(
            "workbook volatile-dependencies relationship targets a missing part",
        )),
        (Some(target), [part_name]) if part_name == target => Ok(()),
        (Some(_), _) => Err(invalid(
            "workbook volatile-dependencies relationship does not target the volatile-dependencies part",
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

fn next_volatile_dependencies_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/volatileDependencies.xml".to_string()
        } else {
            format!("/xl/volatileDependencies{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free volatile-dependencies part name"))
}

fn next_volatile_dependencies_relationship_id(
    package: &OpcPackage,
    workbook_uri: &PackURI,
) -> Result<String> {
    let relationships = package.get_part(workbook_uri)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdVolatileDependencies{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free volatile-dependencies relationship ID"))
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
