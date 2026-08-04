//! OPC relationship loading for inert action settings.

use super::codec::{self, MAX_ATTRIBUTE_BYTES, MAX_SLIDE_XML_BYTES};
use super::model::{self, Setting, Target};
use super::{invalid, limit};
use crate::{Error, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, Part};

/// Per-presentation safety accounting for action-setting discovery.
#[derive(Default)]
pub struct Limits {
    pub(super) total_slide_xml_bytes: usize,
    pub(super) action_count: usize,
}

/// Load bounded, inert action settings from one PresentationML slide.
pub fn load_slide_action_settings(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Vec<Setting>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "action-setting discovery requires a PresentationML slide part",
        ));
    }
    if slide.blob().len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes", MAX_SLIDE_XML_BYTES));
    }
    limits.add_slide_xml(slide.blob().len())?;

    codec::scan(slide.blob(), limits)?
        .into_iter()
        .enumerate()
        .map(|(action_index, parsed)| {
            let target = parsed
                .relationship_id
                .as_deref()
                .map(|relationship_id| resolve_target(package, slide_index, slide, relationship_id))
                .transpose()?;
            let kind = model::classify(parsed.action.as_deref(), parsed.relationship_id.is_some());
            Ok(Setting {
                slide_index,
                action_index,
                trigger: parsed.trigger,
                kind,
                action: parsed.action,
                relationship_id: parsed.relationship_id,
                target,
                tooltip: parsed.tooltip,
                target_frame: parsed.target_frame,
            })
        })
        .collect()
}

fn resolve_target(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    relationship_id: &str,
) -> Result<Target> {
    let relationship = slide.rels().get(relationship_id).ok_or_else(|| {
        Error::Relationship(format!(
            "slide {slide_index} action setting references missing relationship '{relationship_id}'"
        ))
    })?;
    let relationship_type = relationship.reltype().to_owned();
    if relationship.is_external() {
        let target = relationship.target_ref().to_owned();
        if target.is_empty() {
            return Err(Error::Relationship(format!(
                "slide {slide_index} action relationship '{relationship_id}' has an empty external target"
            )));
        }
        bounded(&target, "external action target")?;
        return Ok(Target::External {
            target,
            relationship_type,
        });
    }

    let part_name = relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "slide {slide_index} action relationship '{relationship_id}' has an invalid target: {error}"
        ))
    })?;
    package.get_part(&part_name).map_err(|error| {
        Error::PartNotFound(format!(
            "slide {slide_index} action relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    Ok(Target::Internal {
        part_name,
        relationship_type,
    })
}

fn bounded(value: &str, what: &'static str) -> Result<()> {
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(what, MAX_ATTRIBUTE_BYTES));
    }
    Ok(())
}
