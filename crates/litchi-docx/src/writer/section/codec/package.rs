use crate::error::{Error, Result};
use std::fmt::Write;

use super::super::super::relmap::RelationshipMapper;
use super::super::model::SectionHeaderFooterReference;
use super::xml::escape;

pub(super) fn write_references(
    xml: &mut String,
    element: &str,
    references: &[SectionHeaderFooterReference],
    rels: Option<&RelationshipMapper>,
    header: bool,
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
            write!(
                xml,
                "<w:{element} w:type=\"default\" r:id=\"{}\"/>",
                escape(id)
            )
            .map_err(|error| Error::Xml(error.to_string()))?;
        }
        return Ok(());
    }
    for reference in references {
        let managed = rels.and_then(|rels| {
            if header {
                rels.get_header_id()
            } else {
                rels.get_footer_id()
            }
        });
        let owned = reference
            .part
            .as_ref()
            .and_then(|part| rels.and_then(|rels| rels.get_section_header_footer_id(&part.key)));
        let id = reference
            .relationship_id
            .as_deref()
            .or(owned)
            .or(managed)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("section {element} has no relationship ID"))
            })?;
        write!(
            xml,
            "<w:{element} w:type=\"{}\" r:id=\"{}\"/>",
            reference.kind.to_xml(),
            escape(id)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    Ok(())
}
