use std::io::Write;

use crate::writer::shape::Anchor;
use crate::writer::{XlsShapeFill, XlsShapeLine, XlsShapeText, XlsShapeTextRun, XlsShapeWrite};
use crate::{XlsError, XlsResult};
use litchi_odraw::{
    prop::Id,
    shape::{Flags, Native},
    write::{
        Atom as WriteAtom, Container as WriteContainer, PropertyBuilder, ShapeBuilder,
        atom as write_escher_atom, atom_header as write_atom_header, container as write_container,
        container_header as write_container_header,
    },
};

use super::comment::{self, CommentConfig};
use super::drawing_group::{GroupFragmentObj, GroupShapeConfig, group_fragments};
use super::pivot;
use super::write_record_header;

const MSODRAWING: u16 = 0x00EC;
const OBJ: u16 = 0x005D;
const TXO: u16 = 0x01B6;
const CONTINUE: u16 = 0x003C;

pub(crate) struct PrimitiveShapeConfig<'a> {
    pub shape: &'a XlsShapeWrite,
    pub object_id: u16,
}

enum DrawingObject<'a> {
    Pivot {
        object_id: u16,
        escher: Vec<u8>,
    },
    Primitive {
        config: &'a PrimitiveShapeConfig<'a>,
        escher: Vec<u8>,
        has_textbox: bool,
    },
    Group {
        obj: GroupFragmentObj<'a>,
        escher: Vec<u8>,
        has_textbox: bool,
    },
    Comment {
        config: &'a CommentConfig<'a>,
        escher: Vec<u8>,
    },
}

impl DrawingObject<'_> {
    fn escher(&self) -> &[u8] {
        match self {
            Self::Pivot { escher, .. }
            | Self::Primitive { escher, .. }
            | Self::Group { escher, .. }
            | Self::Comment { escher, .. } => escher,
        }
    }

    fn has_textbox(&self) -> bool {
        matches!(
            self,
            Self::Primitive {
                has_textbox: true,
                ..
            } | Self::Group {
                has_textbox: true,
                ..
            }
        )
    }

    fn officeart_len(&self) -> usize {
        self.escher().len() + usize::from(self.has_textbox()) * 8
    }
}

fn write_mso<W: Write>(writer: &mut W, data: &[u8]) -> XlsResult<()> {
    let length = u16::try_from(data.len())
        .map_err(|_| XlsError::InvalidData("MsoDrawing record is too large".to_string()))?;
    if length > 8224 {
        return Err(XlsError::InvalidData(
            "MsoDrawing record exceeds 8224 bytes".to_string(),
        ));
    }
    write_record_header(writer, MSODRAWING, length)?;
    writer.write_all(data)?;
    Ok(())
}

fn group_prefix(drawing_id: u32, object_count: usize, shapes_size: usize) -> XlsResult<Vec<u8>> {
    if drawing_id == 0 || drawing_id > 0x0FFF || object_count == 0 || object_count > 1022 {
        return Err(XlsError::InvalidData(
            "worksheet drawing ID or object count is outside OfficeArt limits".to_string(),
        ));
    }
    let shape_count = u32::try_from(object_count + 1)
        .map_err(|_| XlsError::InvalidData("drawing shape count overflows".to_string()))?;
    let spgr_length =
        48u32
            .checked_add(u32::try_from(shapes_size).map_err(|_| {
                XlsError::InvalidData("worksheet drawing size exceeds u32".to_string())
            })?)
            .ok_or_else(|| XlsError::InvalidData("worksheet drawing size overflows".to_string()))?;
    let dg_length = 24u32
        .checked_add(spgr_length)
        .ok_or_else(|| XlsError::InvalidData("worksheet drawing size overflows".to_string()))?;
    let patriarch = drawing_id << 10;
    let mut out = Vec::with_capacity(80);
    write_container_header(&mut out, 0, WriteContainer::Dg, dg_length)?;
    write_atom_header(&mut out, drawing_id as u16, WriteAtom::Dg, 8)?;
    out.extend_from_slice(&shape_count.to_le_bytes());
    out.extend_from_slice(&(patriarch + object_count as u32).to_le_bytes());
    write_container_header(&mut out, 0, WriteContainer::Spgr, spgr_length)?;
    let mut patriarch_children = Vec::with_capacity(40);
    write_atom_header(&mut patriarch_children, 0, WriteAtom::Spgr, 16)?;
    patriarch_children.extend_from_slice(&[0; 16]);
    ShapeBuilder::new(Native::FREEFORM, patriarch)
        .with_flags(Flags::GROUP | Flags::PATRIARCH)
        .write(&mut patriarch_children)?;
    write_container(&mut out, 0, WriteContainer::Sp, &patriarch_children)?;
    Ok(out)
}

pub(super) fn write_xls_anchor<W: Write>(writer: &mut W, anchor: &Anchor) -> XlsResult<()> {
    write_atom_header(writer, 0, WriteAtom::ClientAnchor, 18)?;
    for value in anchor.fields() {
        writer.write_all(&value.to_le_bytes())?;
    }
    Ok(())
}

/// Build the shared primitive style OPT properties (protection, fill, line, visibility).
pub(super) fn style_properties(
    locked: bool,
    fill: XlsShapeFill,
    line: XlsShapeLine,
    visible: bool,
) -> PropertyBuilder<'static> {
    let mut properties = PropertyBuilder::new();
    properties.add_simple(
        Id::LockAgainstGrouping,
        if locked { 0x0104_0104 } else { 0x0104_0000 },
    );
    match fill {
        XlsShapeFill::None => {
            properties.add_simple(Id::FillColor, 0);
            properties.add_simple(Id::NoFillHitTest, 0x0010_0000);
        },
        XlsShapeFill::Solid(color) => {
            properties.add_simple(Id::FillColor, color.officeart_color() as i32);
            properties.add_simple(Id::NoFillHitTest, 0x0015_0011);
        },
    }
    match line {
        XlsShapeLine::None => {
            properties.add_simple(Id::LineColor, 0);
            properties.add_simple(Id::LineWidth, 0);
            properties.add_simple(Id::NoLineDrawDash, 0x0008_0000);
        },
        XlsShapeLine::Solid { color, width_emu } => {
            properties.add_simple(Id::LineColor, color.officeart_color() as i32);
            properties.add_simple(Id::LineWidth, width_emu as i32);
            properties.add_simple(Id::NoLineDrawDash, 0x0008_0008);
        },
    }
    properties.add_simple(
        Id::from(0x03BF),
        if visible { 0x0002_0000 } else { 0x0002_0002 },
    );
    properties
}

pub(super) fn split_client_textbox(
    mut escher: Vec<u8>,
    has_textbox: bool,
) -> XlsResult<(Vec<u8>, bool)> {
    if has_textbox {
        let split = escher.len().checked_sub(8).ok_or_else(|| {
            XlsError::InvalidData("OfficeArt textbox shape is truncated".to_string())
        })?;
        escher.truncate(split);
    }
    Ok((escher, has_textbox))
}

fn primitive_shape(shape: &XlsShapeWrite, shape_id: u32) -> XlsResult<(Vec<u8>, bool)> {
    let mut children = Vec::with_capacity(112);
    ShapeBuilder::new(Native::from_raw(shape.kind.officeart_type()), shape_id)
        .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
        .write(&mut children)?;
    style_properties(shape.locked, shape.fill, shape.line, shape.visible).write(&mut children)?;
    write_xls_anchor(&mut children, &shape.anchor)?;
    let has_textbox = shape.kind == crate::writer::XlsShapeKind::TextBox || shape.text.is_some();
    write_escher_atom(&mut children, 0, WriteAtom::ClientData, &[])?;
    if has_textbox {
        write_escher_atom(&mut children, 0, WriteAtom::ClientTextbox, &[])?;
    }
    let mut out = Vec::with_capacity(children.len() + 8);
    write_container(&mut out, 0, WriteContainer::Sp, &children)?;
    split_client_textbox(out, has_textbox)
}

fn pivot_shape(shape_id: u32) -> XlsResult<Vec<u8>> {
    let mut children = Vec::with_capacity(82);
    ShapeBuilder::new(Native::from_raw(201), shape_id)
        .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
        .write(&mut children)?;
    let mut properties = PropertyBuilder::new();
    properties.add_simple(Id::LockAgainstGrouping, 0x0104_0104);
    properties.add_simple(Id::TextBooleanProperties, 0x0008_0008);
    properties.add_simple(Id::NoLineDrawDash, 0x0008_0000);
    properties.add_simple(Id::from(0x03BF), 0x0002_0000);
    properties.write(&mut children)?;
    write_atom_header(&mut children, 0, WriteAtom::ClientAnchor, 18)?;
    for field in [1u16, 1, 0, 0, 0, 2, 0, 1, 0] {
        children.extend_from_slice(&field.to_le_bytes());
    }
    write_escher_atom(&mut children, 0, WriteAtom::ClientData, &[])?;
    let mut out = Vec::with_capacity(children.len() + 8);
    write_container(&mut out, 0, WriteContainer::Sp, &children)?;
    Ok(out)
}

/// ftCmo feature type and payload size (MS-XLS 2.5.143 FtCmo).
const FT_CMO: u16 = 0x0015;
const FT_CMO_SIZE: u16 = 0x0012;
/// ftGmo group marker feature type and payload size (MS-XLS 2.5.148 FtGmo).
const FT_GMO: u16 = 0x0006;
const FT_GMO_SIZE: u16 = 0x0002;
/// ftCmo object type for group objects (MS-XLS 2.5.143).
const OBJECT_TYPE_GROUP: u16 = 0x0000;

fn write_ft_cmo<W: Write>(
    writer: &mut W,
    object_type: u16,
    object_id: u16,
    locked: bool,
    visible: bool,
) -> XlsResult<()> {
    writer.write_all(&FT_CMO.to_le_bytes())?;
    writer.write_all(&FT_CMO_SIZE.to_le_bytes())?;
    writer.write_all(&object_type.to_le_bytes())?;
    writer.write_all(&object_id.to_le_bytes())?;
    let flags = 0x6000u16 | u16::from(locked) | if visible { 0x0010 } else { 0 };
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&[0; 12])?;
    Ok(())
}

fn write_ft_end<W: Write>(writer: &mut W) -> XlsResult<()> {
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
}

fn write_shape_obj<W: Write>(
    writer: &mut W,
    object_type: u16,
    object_id: u16,
    locked: bool,
    visible: bool,
) -> XlsResult<()> {
    write_record_header(writer, OBJ, 26)?;
    write_ft_cmo(writer, object_type, object_id, locked, visible)?;
    write_ft_end(writer)
}

/// Write the OBJ record of a group object: ftCmo + the mandatory ftGmo marker.
fn write_group_obj<W: Write>(
    writer: &mut W,
    object_id: u16,
    locked: bool,
    visible: bool,
) -> XlsResult<()> {
    write_record_header(writer, OBJ, 32)?;
    write_ft_cmo(writer, OBJECT_TYPE_GROUP, object_id, locked, visible)?;
    writer.write_all(&FT_GMO.to_le_bytes())?;
    writer.write_all(&FT_GMO_SIZE.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    write_ft_end(writer)
}

fn write_continue<W: Write>(writer: &mut W, data: &[u8]) -> XlsResult<()> {
    let length = u16::try_from(data.len())
        .map_err(|_| XlsError::InvalidData("shape CONTINUE record is too large".to_string()))?;
    if length > 8224 {
        return Err(XlsError::InvalidData(
            "shape CONTINUE record exceeds 8224 bytes".to_string(),
        ));
    }
    write_record_header(writer, CONTINUE, length)?;
    writer.write_all(data)?;
    Ok(())
}

fn write_shape_txo<W: Write>(writer: &mut W, text: Option<&XlsShapeText>) -> XlsResult<()> {
    let value = text.map_or("", |text| text.value.as_str());
    let units = value.encode_utf16().collect::<Vec<_>>();
    let runs = match text {
        _ if units.is_empty() => Vec::new(),
        Some(text) if !text.runs.is_empty() => text.runs.clone(),
        _ => vec![XlsShapeTextRun {
            character_index: 0,
            font_index: 0,
        }],
    };
    let run_bytes = if units.is_empty() {
        0
    } else {
        (runs.len() + 1) * 8
    };
    write_record_header(writer, TXO, 18)?;
    writer.write_all(&0x0212u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&[0; 6])?;
    writer.write_all(&(units.len() as u16).to_le_bytes())?;
    writer.write_all(&(run_bytes as u16).to_le_bytes())?;
    writer.write_all(&text.map_or(0, |text| text.font_when_empty).to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;

    let mut start = 0usize;
    while start < units.len() {
        let mut end = (start + 4111).min(units.len());
        if end < units.len() && (0xD800..=0xDBFF).contains(&units[end - 1]) {
            end -= 1;
        }
        let mut segment = Vec::with_capacity(1 + (end - start) * 2);
        segment.push(1);
        for unit in &units[start..end] {
            segment.extend_from_slice(&unit.to_le_bytes());
        }
        write_continue(writer, &segment)?;
        start = end;
    }
    if !runs.is_empty() {
        let mut bytes = Vec::with_capacity(run_bytes);
        for run in runs {
            bytes.extend_from_slice(&run.character_index.to_le_bytes());
            bytes.extend_from_slice(&run.font_index.to_le_bytes());
            bytes.extend_from_slice(&[0; 4]);
        }
        bytes.extend_from_slice(&(units.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&[0; 6]);
        for chunk in bytes.chunks(8224) {
            write_continue(writer, chunk)?;
        }
    }
    Ok(())
}

fn write_client_textbox<W: Write>(writer: &mut W) -> XlsResult<()> {
    let mut textbox = Vec::with_capacity(8);
    write_escher_atom(&mut textbox, 0, WriteAtom::ClientTextbox, &[])?;
    write_mso(writer, &textbox)
}

pub(crate) fn write_worksheet_drawing<W: Write>(
    writer: &mut W,
    drawing_id: u32,
    pivot_object_ids: &[u16],
    primitives: &[PrimitiveShapeConfig<'_>],
    groups: &[GroupShapeConfig<'_>],
    comments: &[CommentConfig<'_>],
) -> XlsResult<()> {
    let group_object_count = groups
        .iter()
        .map(|config| 1 + config.group.children.len())
        .sum::<usize>();
    let object_count =
        pivot_object_ids.len() + primitives.len() + group_object_count + comments.len();
    if object_count == 0 {
        return Ok(());
    }
    let mut objects = Vec::with_capacity(object_count);
    let mut offset = 1u32;
    for &object_id in pivot_object_ids {
        objects.push(DrawingObject::Pivot {
            object_id,
            escher: pivot_shape((drawing_id << 10) + offset)?,
        });
        offset += 1;
    }
    for config in primitives {
        let (escher, has_textbox) = primitive_shape(config.shape, (drawing_id << 10) + offset)?;
        objects.push(DrawingObject::Primitive {
            config,
            escher,
            has_textbox,
        });
        offset += 1;
    }
    for config in groups {
        let fragments = group_fragments(config, (drawing_id << 10) + offset)?;
        offset += fragments.len() as u32;
        objects.extend(fragments.into_iter().map(|fragment| DrawingObject::Group {
            obj: fragment.obj,
            escher: fragment.escher,
            has_textbox: fragment.has_textbox,
        }));
    }
    for config in comments {
        objects.push(DrawingObject::Comment {
            config,
            escher: comment::comment_shape(config, (drawing_id << 10) + offset),
        });
        offset += 1;
    }
    let shapes_size = objects.iter().map(DrawingObject::officeart_len).sum();
    let prefix = group_prefix(drawing_id, object_count, shapes_size)?;
    for (index, object) in objects.iter().enumerate() {
        let mut drawing = if index == 0 {
            prefix.clone()
        } else {
            Vec::new()
        };
        drawing.extend_from_slice(object.escher());
        write_mso(writer, &drawing)?;
        match object {
            DrawingObject::Pivot { object_id, .. } => {
                pivot::write_pivot_page_obj(writer, *object_id)?;
            },
            DrawingObject::Primitive {
                config,
                has_textbox,
                ..
            } => {
                write_shape_obj(
                    writer,
                    config.shape.kind.object_type(),
                    config.object_id,
                    config.shape.locked,
                    config.shape.visible,
                )?;
                if *has_textbox {
                    write_client_textbox(writer)?;
                    write_shape_txo(writer, config.shape.text.as_ref())?;
                }
            },
            DrawingObject::Group {
                obj, has_textbox, ..
            } => match obj {
                GroupFragmentObj::Header {
                    object_id,
                    locked,
                    visible,
                } => {
                    write_group_obj(writer, *object_id, *locked, *visible)?;
                },
                GroupFragmentObj::Child { child, object_id } => {
                    write_shape_obj(
                        writer,
                        child.kind.object_type(),
                        *object_id,
                        child.locked,
                        child.visible,
                    )?;
                    if *has_textbox {
                        write_client_textbox(writer)?;
                        write_shape_txo(writer, child.text.as_ref())?;
                    }
                },
            },
            DrawingObject::Comment { config, .. } => {
                comment::write_obj(writer, config)?;
                write_client_textbox(writer)?;
                comment::write_txo(writer, config)?;
            },
        }
    }
    for comment in comments {
        comment::write_note(writer, comment)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::writer::shape::{Behavior, Point};

    #[test]
    fn checked_anchor_serializes_exact_client_anchor_bytes() {
        let anchor = Anchor::new(
            Point::new(1, 1).unwrap().offset(20, 10).unwrap(),
            Point::new(6, 4).unwrap().offset(200, 900).unwrap(),
            Behavior::MoveAndSize,
        )
        .unwrap();
        let mut bytes = Vec::new();
        write_xls_anchor(&mut bytes, &anchor).unwrap();
        assert_eq!(
            bytes,
            vec![
                0x00, 0x00, 0x10, 0xF0, 0x12, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x00, 0x0A, 0x00,
                0x01, 0x00, 0x14, 0x00, 0x04, 0x00, 0x84, 0x03, 0x06, 0x00, 0xC8, 0x00,
            ]
        );
    }
}
