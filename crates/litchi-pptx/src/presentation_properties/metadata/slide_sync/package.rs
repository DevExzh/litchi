//! OPC graph lifecycle for slide synchronization metadata.

use litchi_opc::OpcPackage;
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;

use super::codec::Limits;
use super::model::Part;
use crate::{Error, Result};

/// Content type of a Slide Synchronization Data part.
pub const CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slideSyncData+xml";

/// Relationship type from a slide to its synchronization data part.
pub const RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideSyncData";

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

/// Load every synchronization part and validate its slide relationship graph.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, limits: Limits) -> Result<Vec<Part>> {
    if package
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "slide synchronization relationship cannot originate at the package root",
        ));
    }

    let mut loaded = Vec::new();
    let mut total_xml_bytes = 0usize;
    for source in package.iter_parts() {
        let relationships: Vec<_> = source
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
            .collect();
        if relationships.is_empty() {
            continue;
        }
        if source.content_type() != ct::PML_SLIDE {
            return Err(invalid(
                "slide synchronization relationship must originate at a slide part",
            ));
        }
        if relationships.len() > 1 {
            return Err(invalid(
                "slide part has multiple slide synchronization relationships",
            ));
        }
        let relationship = relationships[0];
        if relationship.is_external() {
            return Err(invalid(
                "slide synchronization relationship cannot be external",
            ));
        }
        if loaded.len() >= limits.max_parts() {
            return Err(limit(
                crate::error::CodecLimitKind::Parts,
                loaded.len().checked_add(1).unwrap_or(usize::MAX),
                limits.max_parts(),
            ));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type() != CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: CONTENT_TYPE.to_owned(),
                actual: part.content_type().to_owned(),
            });
        }
        if !part.rels().is_empty() {
            return Err(invalid(
                "slide synchronization part cannot have outbound relationships",
            ));
        }
        total_xml_bytes =
            total_xml_bytes
                .checked_add(part.blob().len())
                .ok_or(Error::CapacityOverflow {
                    resource: "total slide synchronization XML bytes",
                })?;
        if total_xml_bytes > limits.max_total_input_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                total_xml_bytes,
                limits.max_total_input_bytes(),
            ));
        }
        loaded.push(Part::new(
            relationship.r_id().to_owned(),
            source.partname().clone(),
            target,
            super::model::Properties::parse(part.blob(), limits)?,
        ));
    }

    for part in package.iter_parts() {
        if part.content_type() != CONTENT_TYPE {
            continue;
        }
        let references = loaded
            .iter()
            .filter(|entry| entry.part_name == *part.partname())
            .count();
        match references {
            1 => {},
            0 => {
                return Err(invalid(
                    "package contains an orphan slide synchronization part",
                ));
            },
            _ => {
                return Err(invalid(
                    "slide synchronization part is referenced by multiple slides",
                ));
            },
        }
    }
    Ok(loaded)
}

/// Attach one synchronization part to its source slide.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub(crate) fn store(package: &mut OpcPackage, value: &Part, limits: Limits) -> Result<()> {
    validate_ncname(&value.relationship_id)?;
    let slide_name = &value.slide_part_name;
    let part_name = &value.part_name;
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(Error::ContentType {
            expected: ct::PML_SLIDE.to_owned(),
            actual: slide.content_type().to_owned(),
        });
    }
    if slide
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == RELATIONSHIP_TYPE)
    {
        return Err(invalid(
            "slide part already has a slide synchronization relationship",
        ));
    }
    if slide.rels().get(&value.relationship_id).is_some() {
        return Err(invalid(
            "slide synchronization relationship ID already exists",
        ));
    }
    if package
        .iter_parts()
        .any(|part| part.partname() == part_name)
    {
        return Err(invalid(format!(
            "part '{part_name}' already exists in the package"
        )));
    }

    let xml = value.properties.to_xml(limits)?;
    let target = part_name.relative_ref(slide_name.base_uri());
    package.try_add_part(Box::new(BlobPart::new(
        part_name.clone(),
        CONTENT_TYPE.to_owned(),
        xml,
    )))?;
    package
        .get_part_mut(slide_name)?
        .rels_mut()
        .add_relationship(
            RELATIONSHIP_TYPE.to_owned(),
            target,
            value.relationship_id.clone(),
            false,
        );
    Ok(())
}

fn validate_ncname(value: &str) -> Result<()> {
    let mut characters = value.chars();
    let valid = characters
        .next()
        .is_some_and(|character| character == '_' || character.is_ascii_alphabetic())
        && characters.all(|character| {
            character == '_'
                || character == '-'
                || character == '.'
                || character.is_ascii_alphanumeric()
        });
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "slide synchronization relationship ID is not an XML NCName",
        ))
    }
}
