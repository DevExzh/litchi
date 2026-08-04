//! OPC ownership and failure-atomic package service for web settings.

use super::codec::{read, write};
use super::model::{Conformance, Settings};
use super::{
    CONTENT_TYPE, STRICT_OFFICE_DOCUMENT_RELATIONSHIP, invalid, is_web_settings_relationship,
};
use crate::{Error, Result};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

#[derive(Debug, Clone)]
struct Owner {
    main: PackURI,
    target: PackURI,
    relationship_id: String,
    conformance: Conformance,
}

/// Load the document-owned web-settings model, if present.
pub fn load(package: &OpcPackage) -> Result<Option<(Settings, Conformance)>> {
    let Some(owner) = locate(package)? else {
        return Ok(None);
    };
    let part = package.get_part(&owner.target)?;
    let (settings, conformance) = read(part)?;
    if conformance != owner.conformance {
        return Err(invalid(
            "web-settings relationship and XML use different conformance families",
        ));
    }
    Ok(Some((settings, conformance)))
}

/// Move a complete model into package ownership.
///
/// Serialization, graph validation, and a semantic round trip complete before
/// signatures or package members are changed. Semantic/conformance equality is
/// a no-op that retains the producer's exact original bytes and signatures.
/// A different requested conformance is rejected before mutation. A real
/// semantic edit writes canonical modeled XML; ignored or unknown extension
/// markup is not retained source-surgically.
pub fn put(package: &mut OpcPackage, value: Settings, conformance: Conformance) -> Result<bool> {
    let xml = write(&value, conformance)?;
    let package_conformance = package_conformance(package)?;
    if package_conformance != conformance {
        return Err(invalid(
            "web-settings conformance does not match the document package",
        ));
    }

    let existing = locate(package)?;
    if let Some(owner) = &existing {
        let part = package.get_part(&owner.target)?;
        let (current, parsed_conformance) = read(part)?;
        if parsed_conformance != owner.conformance {
            return Err(invalid(
                "web-settings relationship and XML use different conformance families",
            ));
        }
        if owner.conformance == conformance && current == value {
            return Ok(false);
        }
        if has_other_inbound(package, owner)? {
            return Err(invalid(format!(
                "shared web-settings part '{}' cannot be overwritten",
                owner.target
            )));
        }

        let mut staged = BlobPart::new(owner.target.clone(), CONTENT_TYPE.to_owned(), xml);
        copy_relationships(part, &mut staged);
        let (round_trip, staged_conformance) = read(&staged)?;
        if round_trip != value || staged_conformance != conformance {
            return Err(invalid("staged web-settings XML did not round-trip"));
        }
        validate_internal_targets(package)?;

        package.add_part(Box::new(staged));
        package.unsign();
        return Ok(true);
    }

    let main = package.main_document_part()?.partname().clone();
    let target = next_part_name(package)?;
    let staged = BlobPart::new(target.clone(), CONTENT_TYPE.to_owned(), xml);
    let (round_trip, staged_conformance) = read(&staged)?;
    if round_trip != value || staged_conformance != conformance {
        return Err(invalid("staged web-settings XML did not round-trip"));
    }
    validate_internal_targets(package)?;
    let target_ref = target.relative_ref(main.base_uri());
    let relationship_type = conformance.relationship();

    package
        .get_part_mut(&main)?
        .rels_mut()
        .get_or_add(relationship_type, &target_ref);
    package.add_part(Box::new(staged));
    package.unsign();
    Ok(true)
}

/// Remove the document-owned web-settings part.
///
/// A part shared by another relationship is rejected rather than silently
/// detached or deleted. Absence is an exact, signature-preserving no-op.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    let Some(owner) = locate(package)? else {
        return Ok(false);
    };
    let part = package.get_part(&owner.target)?;
    let (_, conformance) = read(part)?;
    if conformance != owner.conformance {
        return Err(invalid(
            "web-settings relationship and XML use different conformance families",
        ));
    }
    if has_other_inbound(package, &owner)? {
        return Err(invalid(format!(
            "shared web-settings part '{}' cannot be removed",
            owner.target
        )));
    }
    validate_internal_targets(package)?;

    package
        .get_part_mut(&owner.main)?
        .rels_mut()
        .remove(&owner.relationship_id);
    package.remove_part(&owner.target);
    package.unsign();
    Ok(true)
}

fn locate(package: &OpcPackage) -> Result<Option<Owner>> {
    use litchi_opc::constants::content_type as ct;

    let main_part = package.main_document_part()?;
    if !matches!(
        main_part.content_type(),
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(invalid("main document is not a WordprocessingML document"));
    }
    let main = main_part.partname().clone();
    let expected_conformance = package_conformance(package)?;

    if package
        .rels()
        .iter()
        .any(|relationship| is_web_settings_relationship(relationship.reltype()))
    {
        return Err(invalid(
            "package root cannot own a web-settings relationship",
        ));
    }
    for part in package.iter_parts() {
        if part.partname() != &main
            && part
                .rels()
                .iter()
                .any(|relationship| is_web_settings_relationship(relationship.reltype()))
        {
            return Err(invalid(format!(
                "web-settings relationship has invalid source '{}'",
                part.partname()
            )));
        }
    }

    let mut relationships = main_part
        .rels()
        .iter()
        .filter(|relationship| is_web_settings_relationship(relationship.reltype()));
    let relationship = relationships.next();
    if relationships.next().is_some() {
        return Err(invalid("document has multiple web-settings relationships"));
    }

    let mut parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let part_name = parts.next().map(|part| part.partname());
    if parts.next().is_some() {
        return Err(invalid("package has multiple web-settings parts"));
    }

    let Some(relationship) = relationship else {
        if part_name.is_some() {
            return Err(invalid(
                "package contains a web-settings part without document ownership",
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid("web-settings relationship cannot be external"));
    }
    let conformance = Conformance::from_relationship(relationship.reltype())
        .ok_or_else(|| invalid("web-settings relationship type is unsupported"))?;
    if conformance != expected_conformance {
        return Err(invalid(
            "web-settings relationship conformance does not match the package",
        ));
    }
    let target = relationship.target_partname()?;
    let Some(part_name) = part_name else {
        let part = package.get_part(&target)?;
        return Err(Error::ContentType {
            expected: CONTENT_TYPE.to_owned(),
            actual: part.content_type().to_owned(),
        });
    };
    if part_name != &target {
        return Err(invalid(
            "web-settings relationship does not target the web-settings part",
        ));
    }

    Ok(Some(Owner {
        main,
        target,
        relationship_id: relationship.r_id().to_owned(),
        conformance,
    }))
}

fn package_conformance(package: &OpcPackage) -> Result<Conformance> {
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            litchi_opc::constants::relationship_type::OFFICE_DOCUMENT
                | STRICT_OFFICE_DOCUMENT_RELATIONSHIP
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(
        if relationship.reltype() == STRICT_OFFICE_DOCUMENT_RELATIONSHIP {
            Conformance::Strict
        } else {
            Conformance::Transitional
        },
    )
}

fn copy_relationships(source: &dyn Part, target: &mut dyn Part) {
    for relationship in source.rels().iter() {
        target.rels_mut().add_relationship(
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.r_id().to_owned(),
            relationship.is_external(),
        );
    }
}

fn has_other_inbound(package: &OpcPackage, owner: &Owner) -> Result<bool> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        if relationship.target_partname()? == owner.target {
            return Ok(true);
        }
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            if relationship.target_partname()? == owner.target
                && (part.partname() != &owner.main || relationship.r_id() != owner.relationship_id)
            {
                return Ok(true);
            }
        }
    }
    Ok(false)
}

fn validate_internal_targets(package: &OpcPackage) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        relationship.target_partname()?;
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            relationship.target_partname()?;
        }
    }
    Ok(())
}

fn next_part_name(package: &OpcPackage) -> Result<PackURI> {
    for index in 1..=4096 {
        let name = if index == 1 {
            "/word/webSettings.xml".to_owned()
        } else {
            format!("/word/webSettings{index}.xml")
        };
        let candidate = PackURI::new(name).map_err(Error::Uri)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    Err(invalid("no bounded web-settings part name is available"))
}
