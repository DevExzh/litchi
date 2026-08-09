//! Semantic and OPC graph validation for inert content-part transactions.

use std::collections::HashSet;

use litchi_opc::{OpcPackage, PackURI, Part, TargetMode};

use super::codec;
use super::model::{ContentPart, Payload, Relationship, RelationshipMetadata, Target};
use crate::presentation::embedded::{MAX_ATTRIBUTE_BYTES, invalid, limit};
use crate::{Error, Result};

pub(crate) const MAX_PAYLOAD_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_TOTAL_PAYLOAD_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_PAYLOAD_RELATIONSHIPS: usize = 4_096;
pub(crate) const MAX_RELATIONSHIP_FIELD_BYTES: usize = 16 * 1024;

pub(crate) fn limit_xml() -> Error {
    limit(
        "content-part slide XML bytes",
        crate::presentation::embedded::MAX_XML_BYTES,
    )
}

pub(crate) fn limit_total_payloads() -> Error {
    limit("total content-part payload bytes", MAX_TOTAL_PAYLOAD_BYTES)
}

pub(crate) fn validate_id(value: &str, label: &'static str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(label, MAX_ATTRIBUTE_BYTES));
    }
    if value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}'
                | '\u{000B}'..='\u{000C}'
                | '\u{000E}'..='\u{001F}'
                | '<'
                | '>'
                | '"'
                | '\''
        ) || character.is_whitespace()
    }) {
        return Err(invalid(format!(
            "{label} contains an invalid XML character"
        )));
    }
    Ok(())
}

pub(crate) fn validate_anchor(anchor: &super::model::Anchor) -> Result<()> {
    let parsed = codec::validate_anchor_xml(anchor.xml())?;
    if parsed != anchor.relationship_id() {
        return Err(invalid(
            "content-part anchor relationship ID does not match its XML",
        ));
    }
    validate_id(anchor.relationship_id(), "content-part relationship ID")
}

pub(crate) fn validate_relationship(relationship: &Relationship) -> Result<()> {
    validate_id(relationship.id(), "content-part relationship ID")?;
    validate_field(
        relationship.relationship_type(),
        "content-part relationship type",
    )?;
    validate_field(
        relationship.target_ref(),
        "content-part relationship target",
    )?;
    if relationship.target_ref().is_empty() {
        return Err(invalid("content-part relationship target cannot be empty"));
    }
    match (&relationship.target, relationship.target_mode()) {
        (Target::Internal(payload), TargetMode::Internal) => validate_payload(payload),
        (Target::External { target_ref }, TargetMode::External) => {
            if target_ref != relationship.target_ref() {
                return Err(invalid(
                    "external content-part target metadata is inconsistent",
                ));
            }
            Ok(())
        },
        (Target::Internal(_), TargetMode::External)
        | (Target::External { .. }, TargetMode::Internal) => Err(invalid(
            "content-part target mode does not match its target kind",
        )),
    }
}

pub(crate) fn validate_payload(payload: &Payload) -> Result<()> {
    if payload.bytes().len() > MAX_PAYLOAD_BYTES {
        return Err(limit("content-part payload bytes", MAX_PAYLOAD_BYTES));
    }
    if payload.part_name().as_str().len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit("content-part payload part name", MAX_ATTRIBUTE_BYTES));
    }
    validate_field(payload.content_type(), "content-part payload content type")?;
    if payload.relationships().len() > MAX_PAYLOAD_RELATIONSHIPS {
        return Err(limit(
            "content-part payload relationships",
            MAX_PAYLOAD_RELATIONSHIPS,
        ));
    }
    let mut ids = HashSet::with_capacity(payload.relationships().len());
    for relationship in payload.relationships() {
        validate_relationship_metadata(relationship)?;
        if !ids.insert(relationship.id()) {
            return Err(invalid("duplicate content-part payload relationship ID"));
        }
    }
    Ok(())
}

pub(crate) fn validate_parts(parts: &[ContentPart]) -> Result<()> {
    if parts.len() > 4_096 {
        return Err(limit("content-part count", 4_096));
    }
    let mut ids = HashSet::with_capacity(parts.len());
    let mut payload_names = HashSet::new();
    let mut payloads = Vec::new();
    for part in parts {
        validate_anchor(part.anchor())?;
        validate_relationship(part.relationship())?;
        if part.relationship_id() != part.relationship().id() {
            return Err(invalid(
                "content-part anchor and relationship IDs are inconsistent",
            ));
        }
        if !ids.insert(part.relationship_id()) {
            return Err(invalid(
                "content-part relationship IDs must be unique within a slide",
            ));
        }
        if let Some(payload) = part.payload() {
            if payload_names.insert(payload.part_name().clone()) {
                payloads.push(payload.clone());
            } else if payloads
                .iter()
                .any(|value: &Payload| value.part_name() == payload.part_name() && value != payload)
            {
                return Err(invalid(
                    "content-part payload declarations conflict for one part",
                ));
            }
        }
    }
    Ok(())
}

pub(crate) fn validate_field(value: &str, label: &'static str) -> Result<()> {
    if value.len() > MAX_RELATIONSHIP_FIELD_BYTES {
        return Err(limit(label, MAX_RELATIONSHIP_FIELD_BYTES));
    }
    if value.chars().any(|character| {
        matches!(
            character,
            '\u{0000}'..='\u{0008}'
                | '\u{000B}'..='\u{000C}'
                | '\u{000E}'..='\u{001F}'
        )
    }) {
        return Err(invalid(format!(
            "{label} contains an invalid XML character"
        )));
    }
    Ok(())
}

pub(crate) fn validate_relationship_metadata(value: &RelationshipMetadata) -> Result<()> {
    validate_id(value.id(), "content-part payload relationship ID")?;
    validate_field(
        value.relationship_type(),
        "content-part payload relationship type",
    )?;
    validate_field(
        value.target_ref(),
        "content-part payload relationship target",
    )?;
    if value.target_ref().is_empty() {
        return Err(invalid(
            "content-part payload relationship target cannot be empty",
        ));
    }
    Ok(())
}

/// Validate the selected owner against the actual package graph after a
/// staged publication. Payload vocabularies remain opaque; only OPC edges
/// and byte ownership are checked.
pub(crate) fn validate_package_graph(
    package: &OpcPackage,
    slide: &dyn Part,
    parts: &[ContentPart],
) -> Result<()> {
    validate_parts(parts)?;
    let mut referenced = HashSet::<PackURI>::new();
    for part in parts {
        let relationship = slide
            .rels()
            .get(part.relationship_id())
            .ok_or_else(|| Error::Relationship("content-part relationship disappeared".into()))?;
        if relationship.target_ref() != part.relationship().target_ref()
            || relationship.reltype() != part.relationship().relationship_type()
            || relationship.target_mode() != part.relationship().target_mode()
        {
            return Err(Error::Relationship(
                "content-part relationship metadata differs from the package graph".into(),
            ));
        }
        if let Target::Internal(payload) = part.target() {
            let target = relationship.target_partname()?;
            if !target.is_equivalent_to(payload.part_name()) {
                return Err(Error::Relationship(
                    "content-part relationship target differs from its payload".into(),
                ));
            }
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != payload.content_type()
                || target_part.blob() != payload.bytes()
            {
                return Err(Error::Relationship(
                    "content-part payload differs from the package target".into(),
                ));
            }
            referenced.insert(target);
            for nested in payload.relationships() {
                if nested.target_mode() == TargetMode::Internal {
                    let relationship = litchi_opc::Relationship::new_with_mode(
                        nested.id().to_owned(),
                        nested.relationship_type().to_owned(),
                        nested.target_ref().to_owned(),
                        payload.part_name().base_uri().to_owned(),
                        TargetMode::Internal,
                    );
                    package.get_part(&relationship.target_partname()?)?;
                }
            }
        }
    }
    for target in referenced {
        let has_inbound = package.iter_parts().any(|source| {
            source.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|value| value.is_equivalent_to(&target))
            })
        });
        if !has_inbound {
            return Err(Error::Relationship(format!(
                "content-part payload '{}' has no inbound relationship",
                target.as_str()
            )));
        }
    }
    Ok(())
}

pub(crate) fn invalid_revision() -> Error {
    invalid("content-part patch source is stale")
}
