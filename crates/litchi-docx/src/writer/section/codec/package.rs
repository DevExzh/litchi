#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use crate::error::{Error, Result};
use litchi_ooxml_common::xml_name::is_ncname;
use std::fmt::Write;

use super::super::super::relmap::RelationshipMapper;
use super::super::model::SectionHeaderFooterReference;
use super::xml::{NamespacePlan, escape};

pub(super) fn write_references(
    xml: &mut String,
    element: &str,
    references: &[SectionHeaderFooterReference],
    rels: Option<&RelationshipMapper>,
    header: bool,
    plan: &NamespacePlan,
) -> Result<()> {
    if references.is_empty() {
        let managed = rels.and_then(|rels| {
            if header {
                rels.get_header_id()
            } else {
                rels.get_footer_id()
            }
        });
        if let Some(id) = managed {
            validate_relationship_id(id, element)?;
            write!(
                xml,
                "<{} {}=\"default\" {}=\"{}\"/>",
                plan.word_qname(element),
                plan.word_attribute_qname("type"),
                plan.relationship_qname("id"),
                escape(id)
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        }
        return Ok(());
    }
    for reference in references {
        write_reference(xml, element, reference, rels, header, plan, &[])?;
    }
    Ok(())
}

pub(super) fn write_reference(
    xml: &mut String,
    element: &str,
    reference: &SectionHeaderFooterReference,
    rels: Option<&RelationshipMapper>,
    header: bool,
    plan: &NamespacePlan,
    extra_attributes: &[String],
) -> Result<()> {
    let managed = rels.and_then(|rels| {
        if header {
            rels.get_header_id()
        } else {
            rels.get_footer_id()
        }
    });
    let id = if let Some(part) = reference.part.as_ref() {
        rels.and_then(|rels| rels.get_section_header_footer_id(&part.key))
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "section {element} owned part has no relationship ID"
                ))
            })?
    } else {
        reference
            .relationship_id
            .as_deref()
            .or(managed)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("section {element} has no relationship ID"))
            })?
    };
    validate_relationship_id(id, element)?;
    for attribute in extra_attributes {
        if plan
            .word_prefix
            .as_deref()
            .is_some_and(|prefix| plan.shadows_word_prefix(attribute, prefix))
            || plan.shadows_word_prefix(attribute, &plan.word_attribute_prefix)
            || plan.shadows_relationship_prefix(attribute)
        {
            return Err(Error::InvalidFormat(format!(
                "section {element} shadows a generated namespace prefix"
            )));
        }
    }
    write!(
        xml,
        "<{} {}=\"{}\" {}=\"{}\"",
        plan.word_qname(element),
        plan.word_attribute_qname("type"),
        reference.kind.to_xml(),
        plan.relationship_qname("id"),
        escape(id)
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    for attribute in extra_attributes {
        write!(xml, " {attribute}").map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

fn validate_relationship_id(id: &str, element: &str) -> Result<()> {
    if !is_ncname(id) {
        return Err(Error::InvalidFormat(format!(
            "section {element} relationship ID is not an XML NCName"
        )));
    }
    Ok(())
}
