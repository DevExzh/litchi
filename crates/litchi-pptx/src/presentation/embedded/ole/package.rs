use super::codec::{inventory, parse_tree};
use super::model::{Kind, Object, Target};
use crate::presentation::embedded::{invalid, limit};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, Part};

const MAX_SLIDE_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_BYTES: usize = 256 * 1024 * 1024;

/// Shared OLE inventory bounds across slides.
#[derive(Debug, Default)]
pub struct Limits {
    total_slide_bytes: usize,
    object_count: usize,
}

impl Limits {
    fn add_slide(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_BYTES {
            return Err(limit("OLE slide XML bytes", MAX_SLIDE_BYTES));
        }
        self.total_slide_bytes = self
            .total_slide_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("OLE slide XML bytes", MAX_TOTAL_SLIDE_BYTES))?;
        if self.total_slide_bytes > MAX_TOTAL_SLIDE_BYTES {
            return Err(limit("OLE slide XML bytes", MAX_TOTAL_SLIDE_BYTES));
        }
        Ok(())
    }

    fn add_objects(&mut self, count: usize) -> Result<()> {
        self.object_count = self
            .object_count
            .checked_add(count)
            .ok_or_else(|| limit("OLE object count", super::codec::MAX_OBJECTS))?;
        if self.object_count > super::codec::MAX_OBJECTS {
            return Err(limit("OLE object count", super::codec::MAX_OBJECTS));
        }
        Ok(())
    }
}

/// Load the inert OLE inventory for one PresentationML slide.
pub fn load_slide(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut Limits,
) -> Result<Vec<Object>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid("OLE discovery requires a PresentationML slide"));
    }
    limits.add_slide(slide.blob().len())?;
    let root = parse_tree(slide.blob())?;
    let parsed = inventory(&root)?;
    limits.add_objects(parsed.len())?;
    parsed
        .into_iter()
        .enumerate()
        .map(|(index, value)| {
            let (kind, target) = value
                .relationship_id
                .as_deref()
                .map(|id| resolve_target(package, slide, id))
                .transpose()?
                .map_or((None, None), |(kind, target)| (Some(kind), Some(target)));
            Ok(Object {
                slide_index,
                index,
                shape_id: value.shape_id,
                shape_name: value.shape_name,
                legacy_shape_id: value.legacy_shape_id,
                name: value.name,
                program_id: value.program_id,
                show_as_icon: value.show_as_icon,
                preview_width: value.preview_width,
                preview_height: value.preview_height,
                anchor: value.anchor,
                mode: value.mode,
                relationship_id: value.relationship_id,
                kind,
                target,
                preview_relationship_id: value.preview_relationship_id,
            })
        })
        .collect()
}

fn resolve_target(package: &OpcPackage, slide: &dyn Part, id: &str) -> Result<(Kind, Target)> {
    let relationship = slide
        .rels()
        .get(id)
        .ok_or_else(|| Error::Relationship(format!("OLE relationship '{id}' is missing")))?;
    let kind = match relationship.reltype() {
        rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT => Kind::OleObject,
        rt::PACKAGE | rt::STRICT_PACKAGE => Kind::Package,
        other => {
            return Err(Error::Relationship(format!(
                "OLE relationship '{id}' has unsupported type '{other}'"
            )));
        },
    };
    if relationship.is_external() {
        return Ok((
            kind,
            Target::External {
                target: relationship.target_ref().to_string(),
                relationship_type: relationship.reltype().to_string(),
            },
        ));
    }
    let part_name = relationship.target_partname()?;
    let part = package.get_part(&part_name)?;
    let expected = match kind {
        Kind::OleObject => ct::OFC_OLE_OBJECT,
        Kind::Package => ct::OFC_PACKAGE,
    };
    if part.content_type() != expected {
        return Err(Error::ContentType {
            expected: expected.to_string(),
            actual: part.content_type().to_string(),
        });
    }
    Ok((
        kind,
        Target::Internal {
            part_name,
            content_type: part.content_type().to_string(),
            relationship_type: relationship.reltype().to_string(),
        },
    ))
}
