//! OPC relationship and part orchestration for external links.

use crate::error::Result;
use litchi_ooxml_common::external_link::{
    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES, is_external_workbook_relationship,
};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::part::BlobPart;
use litchi_opc::{PackURI, Part};

use super::codec::parse_external_link;
use super::model::*;
use super::{invalid, limit};

/// One external-link part together with its workbook package relationship.
///
/// This physical identity belongs to the OPC package layer; the typed link
/// models remain usable without a package or relationship catalog.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Entry {
    pub index: u32,
    pub relationship_id: String,
    pub part_uri: PackURI,
    pub link: Link,
}

pub fn build_external_link_part(part_uri: PackURI, kind: &Link) -> Result<BlobPart> {
    build_external_link_part_with_conformance(part_uri, kind, Conformance::Transitional)
}

pub fn build_external_link_part_with_conformance(
    part_uri: PackURI,
    kind: &Link,
    conformance: Conformance,
) -> Result<BlobPart> {
    let xml = kind.to_xml_with_conformance(conformance)?;
    let mut part = BlobPart::new(
        part_uri,
        litchi_opc::constants::content_type::SML_EXTERNAL_LINK.into(),
        xml,
    );
    match kind {
        Link::Workbook(link) => add_external_target_relationship(
            &mut part,
            &link.target,
            EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES,
            "external workbook",
        )?,
        Link::Ole(link) => add_external_target_relationship(
            &mut part,
            &link.target,
            &[rt::OLE_OBJECT, rt::STRICT_OLE_OBJECT],
            "OLE",
        )?,
        Link::Dde(_) => {},
    }
    Ok(part)
}

trait TargetMetadata {
    fn relationship_id(&self) -> &str;
    fn target(&self) -> &str;
    fn relationship_type(&self) -> &str;
}

impl TargetMetadata for Target {
    fn relationship_id(&self) -> &str {
        &self.relationship_id
    }
    fn target(&self) -> &str {
        &self.target
    }
    fn relationship_type(&self) -> &str {
        &self.relationship_type
    }
}

fn add_external_target_relationship(
    part: &mut BlobPart,
    target: &impl TargetMetadata,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    validate_external_target(target, allowed_types, description)?;
    part.rels_mut().add_relationship(
        target.relationship_type().to_string(),
        target.target().to_string(),
        target.relationship_id().to_string(),
        true,
    );
    Ok(())
}

fn validate_external_target(
    target: &impl TargetMetadata,
    allowed_types: &[&str],
    description: &str,
) -> Result<()> {
    if target.relationship_id().is_empty() {
        return Err(invalid(format!(
            "{description} relationship ID must not be empty"
        )));
    }
    if target.target().is_empty() {
        return Err(invalid(format!("{description} target must not be empty")));
    }
    if target.target().len() > MAX_EXTERNAL_TARGET_BYTES {
        return Err(limit(&format!("{description} target URI")));
    }
    if target.target().chars().any(|character| {
        character.is_control() || character == '\u{fffe}' || character == '\u{ffff}'
    }) {
        return Err(invalid(format!(
            "{description} target URI contains an invalid character"
        )));
    }
    if target.relationship_id().len() > 1024
        || target.relationship_id().chars().any(char::is_control)
    {
        return Err(invalid(format!("{description} relationship ID is invalid")));
    }
    if !allowed_types.contains(&target.relationship_type()) {
        return Err(invalid(format!(
            "{description} has invalid relationship type '{}'",
            target.relationship_type()
        )));
    }
    Ok(())
}

pub fn load_external_link(
    part: &dyn Part,
    workbook_relationship_id: String,
    index: u32,
) -> Result<Entry> {
    let mut kind = parse_external_link(part.blob())?;
    match &mut kind {
        Link::Workbook(book) => {
            let relationship = part
                .rels()
                .get(&book.target.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "externalBook references missing relationship '{}'",
                        book.target.relationship_id
                    ))
                })?;
            if !relationship.is_external() {
                return Err(invalid("externalBook target relationship must be external"));
            }
            if !is_external_workbook_relationship(relationship.reltype()) {
                return Err(invalid(format!(
                    "externalBook target has invalid relationship type '{}'",
                    relationship.reltype()
                )));
            }
            book.target.target = relationship.target_ref().to_string();
            book.target.relationship_type = relationship.reltype().to_string();
        },
        Link::Dde(_) => {},
        Link::Ole(link) => {
            let relationship = part
                .rels()
                .get(&link.target.relationship_id)
                .ok_or_else(|| {
                    invalid(format!(
                        "oleLink references missing relationship '{}'",
                        link.target.relationship_id
                    ))
                })?;
            if !relationship.is_external() {
                return Err(invalid("oleLink target relationship must be external"));
            }
            if !matches!(
                relationship.reltype(),
                rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT
            ) {
                return Err(invalid(format!(
                    "oleLink target has invalid relationship type '{}'",
                    relationship.reltype()
                )));
            }
            link.target.target = relationship.target_ref().to_string();
            link.target.relationship_type = relationship.reltype().to_string();
        },
    }
    Ok(Entry {
        index,
        relationship_id: workbook_relationship_id,
        part_uri: part.partname().clone(),
        link: kind,
    })
}
