//! Finding the equation objects of a presentation's slides.

use super::content::Content;
use super::tree::{self, Node};
use super::{Bounds, Equation};
use crate::embedded::object::{Collection, ContainerKind, ExternalObject, ObjectSubtype};
use crate::package::Result;

pub(super) const PP_DRAWING: u16 = 0x040C;
pub(super) const DG_CONTAINER: u16 = 0xF002;
pub(super) const SPGR_CONTAINER: u16 = 0xF003;
pub(super) const SP_CONTAINER: u16 = 0xF004;
pub(super) const FSP: u16 = 0xF00A;
pub(super) const FOPT: u16 = 0xF00B;
pub(super) const CHILD_ANCHOR: u16 = 0xF00F;
pub(super) const CLIENT_ANCHOR: u16 = 0xF010;
pub(super) const CLIENT_DATA: u16 = 0xF011;
pub(super) const EX_OBJ_REF_ATOM: u16 = 0x0BC1;
/// `ExOleObjStg`.
const EX_OLE_OBJ_STG: u16 = 0x1011;
/// The `pib` property (MS-ODRAW 2.3.23.4), without its `fBid` flag.
pub(super) const PIB: u16 = 0x0104;

/// The stream of an OLE equation object that holds its MTEF.
pub(super) const EQUATION_NATIVE: &str = "Equation Native";
/// The largest decompressed equation storage read.
pub(super) const MAX_STORAGE_BYTES: usize = 64 * 1024 * 1024;

/// One OLE-object shape of a drawing: where it is and what it shows.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct ObjectShape {
    pub(super) shape_id: u32,
    pub(super) object_id: u32,
    pub(super) bounds: Option<Bounds>,
    pub(super) grouped: bool,
    /// The one-based picture-store entry of its preview, when it has one.
    pub(super) picture: Option<u32>,
}

/// Every shape of `slide` (a complete `SlideContainer`) that references an
/// external object, in drawing order.
pub(super) fn object_shapes(slide: &[u8]) -> Result<Vec<ObjectShape>> {
    let mut slide = tree::parse(slide)?;
    slide.expand_all(0)?;
    let mut out = Vec::new();
    if let Some(group) = patriarch(&slide) {
        collect(group, false, &mut out);
    }
    Ok(out)
}

/// The patriarch group of a slide's drawing.
pub(super) fn patriarch(slide: &Node) -> Option<&Node> {
    slide
        .child(PP_DRAWING)?
        .child(DG_CONTAINER)?
        .child(SPGR_CONTAINER)
}

fn collect(group: &Node, grouped: bool, out: &mut Vec<ObjectShape>) {
    for child in group.children() {
        match child.record_type() {
            SP_CONTAINER => {
                if let Some(shape) = object_shape(child, grouped) {
                    out.push(shape);
                }
            },
            SPGR_CONTAINER => collect(child, true, out),
            _ => {},
        }
    }
}

pub(super) fn object_shape(shape: &Node, grouped: bool) -> Option<ObjectShape> {
    let shape_id = shape_id(shape)?;
    let object_id = object_reference(shape)?;
    Some(ObjectShape {
        shape_id,
        object_id,
        bounds: bounds(shape),
        grouped,
        picture: picture(shape),
    })
}

pub(super) fn shape_id(shape: &Node) -> Option<u32> {
    let payload = shape.child(FSP)?.payload()?;
    Some(u32::from_le_bytes(payload.get(..4)?.try_into().ok()?))
}

pub(super) fn object_reference(shape: &Node) -> Option<u32> {
    let atom = shape
        .child(CLIENT_DATA)?
        .child(EX_OBJ_REF_ATOM)?
        .payload()?;
    Some(u32::from_le_bytes(atom.get(..4)?.try_into().ok()?))
}

fn picture(shape: &Node) -> Option<u32> {
    let options = shape.child(FOPT)?;
    let payload = options.payload()?;
    (0..usize::from(options.instance())).find_map(|index| {
        let entry = payload.get(index * 6..index * 6 + 6)?;
        let id = u16::from_le_bytes([entry[0], entry[1]]);
        (id & 0x3FFF == PIB).then(|| u32::from_le_bytes([entry[2], entry[3], entry[4], entry[5]]))
    })
}

/// A shape's anchor: its client anchor (master units) or, inside a group,
/// its child anchor (group coordinates).
pub(super) fn bounds(shape: &Node) -> Option<Bounds> {
    if let Some(anchor) = shape.child(CLIENT_ANCHOR).and_then(Node::payload) {
        return match anchor.len() {
            8 => {
                let at = |index: usize| {
                    i32::from(i16::from_le_bytes([anchor[index], anchor[index + 1]]))
                };
                Some(Bounds::from_edges(at(2), at(0), at(4), at(6)))
            },
            16 => {
                let at = |index: usize| {
                    i32::from_le_bytes([
                        anchor[index],
                        anchor[index + 1],
                        anchor[index + 2],
                        anchor[index + 3],
                    ])
                };
                Some(Bounds::from_edges(at(4), at(0), at(8), at(12)))
            },
            _ => None,
        };
    }
    let anchor = shape.child(CHILD_ANCHOR).and_then(Node::payload)?;
    let at = |index: usize| -> Option<i32> {
        Some(i32::from_le_bytes(
            anchor.get(index..index + 4)?.try_into().ok()?,
        ))
    };
    Some(Bounds::from_edges(at(0)?, at(4)?, at(8)?, at(12)?))
}

/// The embedded OLE definition of `object_id` when it is an equation
/// object: an Equation-subtype object or one whose ProgID is
/// `Equation.*`.
pub(super) fn equation_definition(
    objects: &Collection,
    object_id: u32,
) -> Option<(u32, Option<&str>)> {
    let ExternalObject::Object(definition) = objects.get(object_id)? else {
        return None;
    };
    if !matches!(definition.kind, ContainerKind::Embedded(_)) {
        return None;
    }
    let program = definition.program_id.as_deref();
    let equation = definition.object.subtype == ObjectSubtype::Equation
        || program.is_some_and(|program| program.starts_with("Equation."));
    equation.then_some((definition.object.persist_id, program))
}

/// Whether `record` is an `ExOleObjStg` record ([MS-PPT] 2.10.34).
pub(super) fn is_storage_record(record: &[u8]) -> bool {
    record.len() >= 8 && u16::from_le_bytes([record[2], record[3]]) == EX_OLE_OBJ_STG
}

/// Whether an `ExOleObjStg` record is compressed (record instance 1).
pub(super) fn is_compressed_storage(record: &[u8]) -> bool {
    record.len() >= 2 && u16::from_le_bytes([record[0], record[1]]) >> 4 == 1
}

/// The equations of one slide; `entry` gives each equation object's
/// storage compression and content by persist identifier.
pub(super) fn slide_equations(
    slide: usize,
    record: &[u8],
    objects: &Collection,
    entry: &mut dyn FnMut(u32) -> (bool, Content),
) -> Result<Vec<Equation>> {
    let mut out = Vec::new();
    for shape in object_shapes(record)? {
        let Some((persist_id, program)) = equation_definition(objects, shape.object_id) else {
            continue;
        };
        let (compressed, content) = entry(persist_id);
        out.push(Equation {
            slide,
            shape_id: shape.shape_id,
            object_id: shape.object_id,
            persist_id,
            program_id: program.map(str::to_owned),
            bounds: shape.bounds,
            grouped: shape.grouped,
            compressed,
            content,
        });
    }
    Ok(out)
}
