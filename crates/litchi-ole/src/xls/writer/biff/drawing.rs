use std::io::Write;

use crate::escher::writer::{
    PropertyBuilder, ShapeBuilder, ShapeFlags, record_type, write_atom, write_container,
    write_record_header as write_escher_header,
};
use crate::xls::writer::{XlsShapeFill, XlsShapeLine, XlsShapeText, XlsShapeTextRun, XlsShapeWrite};
use crate::xls::{XlsError, XlsResult};

use super::comment::{self, CommentConfig};
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
    Pivot { object_id: u16, escher: Vec<u8> },
    Primitive {
        config: &'a PrimitiveShapeConfig<'a>,
        escher: Vec<u8>,
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
            | Self::Comment { escher, .. } => escher,
        }
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
    let spgr_length = 48u32
        .checked_add(u32::try_from(shapes_size).map_err(|_| {
            XlsError::InvalidData("worksheet drawing size exceeds u32".to_string())
        })?)
        .ok_or_else(|| XlsError::InvalidData("worksheet drawing size overflows".to_string()))?;
    let dg_length = 24u32
        .checked_add(spgr_length)
        .ok_or_else(|| XlsError::InvalidData("worksheet drawing size overflows".to_string()))?;
    let patriarch = drawing_id << 10;
    let mut out = Vec::with_capacity(80);
    write_escher_header(&mut out, 0x0F, 0, record_type::DG_CONTAINER, dg_length)?;
    write_escher_header(
        &mut out,
        0,
        drawing_id as u16,
        record_type::DG,
        8,
    )?;
    out.extend_from_slice(&shape_count.to_le_bytes());
    out.extend_from_slice(&(patriarch + object_count as u32).to_le_bytes());
    write_escher_header(&mut out, 0x0F, 0, record_type::SPGR_CONTAINER, spgr_length)?;
    let mut patriarch_children = Vec::with_capacity(40);
    write_escher_header(&mut patriarch_children, 1, 0, record_type::SPGR, 16)?;
    patriarch_children.extend_from_slice(&[0; 16]);
    ShapeBuilder::new(0, patriarch)
        .with_flags((ShapeFlags::GROUP | ShapeFlags::PATRIARCH).bits())
        .write(&mut patriarch_children)?;
    write_container(
        &mut out,
        0,
        record_type::SP_CONTAINER,
        &patriarch_children,
    )?;
    Ok(out)
}

fn write_xls_anchor<W: Write>(writer: &mut W, shape: &XlsShapeWrite) -> XlsResult<()> {
    write_escher_header(writer, 0, 0, record_type::CLIENT_ANCHOR, 18)?;
    let anchor = shape.anchor;
    let flags = u16::from(anchor.move_with_cells) | (u16::from(anchor.size_with_cells) << 1);
    for value in [
        flags,
        anchor.first_column,
        anchor.first_column_offset,
        anchor.first_row as u16,
        anchor.first_row_offset,
        anchor.last_column,
        anchor.last_column_offset,
        anchor.last_row as u16,
        anchor.last_row_offset,
    ] {
        writer.write_all(&value.to_le_bytes())?;
    }
    Ok(())
}

fn primitive_shape(shape: &XlsShapeWrite, shape_id: u32) -> XlsResult<Vec<u8>> {
    let mut children = Vec::with_capacity(112);
    ShapeBuilder::new(shape.kind.officeart_type(), shape_id)
        .with_flags((ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT).bits())
        .write(&mut children)?;

    let mut properties = PropertyBuilder::new();
    properties.add_simple(0x007F, if shape.locked { 0x0104_0104 } else { 0x0104_0000 });
    match shape.fill {
        XlsShapeFill::None => {
            properties.add_simple(0x0181, 0);
            properties.add_simple(0x01BF, 0x0010_0000);
        },
        XlsShapeFill::Solid(color) => {
            properties.add_simple(0x0181, color.officeart_color() as i32);
            properties.add_simple(0x01BF, 0x0015_0011);
        },
    }
    match shape.line {
        XlsShapeLine::None => {
            properties.add_simple(0x01C0, 0);
            properties.add_simple(0x01CB, 0);
            properties.add_simple(0x01FF, 0x0008_0000);
        },
        XlsShapeLine::Solid { color, width_emu } => {
            properties.add_simple(0x01C0, color.officeart_color() as i32);
            properties.add_simple(0x01CB, width_emu as i32);
            properties.add_simple(0x01FF, 0x0008_0008);
        },
    }
    properties.add_simple(0x03BF, if shape.visible { 0x0002_0000 } else { 0x0002_0002 });
    properties.write(&mut children)?;
    write_xls_anchor(&mut children, shape)?;
    let has_textbox = shape.kind == crate::xls::writer::XlsShapeKind::TextBox || shape.text.is_some();
    write_atom(
        &mut children,
        0,
        0,
        if has_textbox {
            record_type::CLIENT_TEXTBOX
        } else {
            record_type::CLIENT_DATA
        },
        &[],
    )?;
    let mut out = Vec::with_capacity(children.len() + 8);
    write_container(&mut out, 0, record_type::SP_CONTAINER, &children)?;
    Ok(out)
}

fn pivot_shape(shape_id: u32) -> XlsResult<Vec<u8>> {
    let mut children = Vec::with_capacity(82);
    ShapeBuilder::new(201, shape_id)
        .with_flags((ShapeFlags::HAVE_ANCHOR | ShapeFlags::HAVE_SPT).bits())
        .write(&mut children)?;
    let mut properties = PropertyBuilder::new();
    properties.add_simple(0x007F, 0x0104_0104);
    properties.add_simple(0x00BF, 0x0008_0008);
    properties.add_simple(0x01FF, 0x0008_0000);
    properties.add_simple(0x03BF, 0x0002_0000);
    properties.write(&mut children)?;
    write_escher_header(&mut children, 0, 0, record_type::CLIENT_ANCHOR, 18)?;
    for field in [1u16, 1, 0, 0, 0, 2, 0, 1, 0] {
        children.extend_from_slice(&field.to_le_bytes());
    }
    write_atom(&mut children, 0, 0, record_type::CLIENT_DATA, &[])?;
    let mut out = Vec::with_capacity(children.len() + 8);
    write_container(&mut out, 0, record_type::SP_CONTAINER, &children)?;
    Ok(out)
}

fn write_primitive_obj<W: Write>(
    writer: &mut W,
    config: &PrimitiveShapeConfig<'_>,
) -> XlsResult<()> {
    write_record_header(writer, OBJ, 26)?;
    writer.write_all(&0x0015u16.to_le_bytes())?;
    writer.write_all(&0x0012u16.to_le_bytes())?;
    writer.write_all(&config.shape.kind.object_type().to_le_bytes())?;
    writer.write_all(&config.object_id.to_le_bytes())?;
    let flags = 0x6000u16
        | u16::from(config.shape.locked)
        | if config.shape.visible { 0x0010 } else { 0 };
    writer.write_all(&flags.to_le_bytes())?;
    writer.write_all(&[0; 12])?;
    writer.write_all(&0u16.to_le_bytes())?;
    writer.write_all(&0u16.to_le_bytes())?;
    Ok(())
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
    let runs = if units.is_empty() {
        Vec::new()
    } else if text.is_none_or(|text| text.runs.is_empty()) {
        vec![XlsShapeTextRun {
            character_index: 0,
            font_index: 0,
        }]
    } else {
        text.unwrap().runs.clone()
    };
    let run_bytes = if units.is_empty() { 0 } else { (runs.len() + 1) * 8 };
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

pub(crate) fn write_worksheet_drawing<W: Write>(
    writer: &mut W,
    drawing_id: u32,
    pivot_object_ids: &[u16],
    primitives: &[PrimitiveShapeConfig<'_>],
    comments: &[CommentConfig<'_>],
) -> XlsResult<()> {
    let object_count = pivot_object_ids.len() + primitives.len() + comments.len();
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
        objects.push(DrawingObject::Primitive {
            config,
            escher: primitive_shape(config.shape, (drawing_id << 10) + offset)?,
        });
        offset += 1;
    }
    for config in comments {
        objects.push(DrawingObject::Comment {
            config,
            escher: comment::comment_shape(config, (drawing_id << 10) + offset),
        });
        offset += 1;
    }
    let shapes_size = objects.iter().map(|object| object.escher().len()).sum();
    let prefix = group_prefix(drawing_id, object_count, shapes_size)?;
    for (index, object) in objects.iter().enumerate() {
        let mut drawing = if index == 0 { prefix.clone() } else { Vec::new() };
        drawing.extend_from_slice(object.escher());
        write_mso(writer, &drawing)?;
        match object {
            DrawingObject::Pivot { object_id, .. } => {
                pivot::write_pivot_page_obj(writer, *object_id)?;
            },
            DrawingObject::Primitive { config, .. } => {
                write_primitive_obj(writer, config)?;
                if config.shape.kind == crate::xls::writer::XlsShapeKind::TextBox
                    || config.shape.text.is_some()
                {
                    write_shape_txo(writer, config.shape.text.as_ref())?;
                }
            },
            DrawingObject::Comment { config, .. } => {
                comment::write_obj(writer, config)?;
                let mut textbox = Vec::with_capacity(8);
                write_atom(&mut textbox, 0, 0, record_type::CLIENT_TEXTBOX, &[])?;
                write_mso(writer, &textbox)?;
                comment::write_txo(writer, config)?;
            },
        }
    }
    for comment in comments {
        comment::write_note(writer, comment)?;
    }
    Ok(())
}
