use std::io::Write;

use crate::xls::writer::{XlsCommentAnchor, XlsCommentTextRunWrite};
use crate::xls::{XlsError, XlsResult};

use super::write_record_header;

const MSODRAWING: u16 = 0x00EC;
const OBJ: u16 = 0x005D;
const TXO: u16 = 0x01B6;
const CONTINUE: u16 = 0x003C;
const NOTE: u16 = 0x001C;

pub(crate) struct CommentConfig<'a> {
    pub row: u16,
    pub column: u8,
    pub author: &'a str,
    pub text: &'a str,
    pub visible: bool,
    pub shared: bool,
    pub anchor: XlsCommentAnchor,
    pub text_runs: &'a [XlsCommentTextRunWrite],
    pub font_when_empty: u16,
    pub guid: [u8; 16],
    pub object_id: u16,
}

fn escher_header(out: &mut Vec<u8>, options: u16, record_type: u16, length: u32) {
    out.extend_from_slice(&options.to_le_bytes());
    out.extend_from_slice(&record_type.to_le_bytes());
    out.extend_from_slice(&length.to_le_bytes());
}

fn write_mso<W: Write>(writer: &mut W, data: &[u8]) -> XlsResult<()> {
    let length = u16::try_from(data.len())
        .map_err(|_| XlsError::InvalidData("comment MsoDrawing record is too large".to_string()))?;
    if length > 8224 {
        return Err(XlsError::InvalidData(
            "comment MsoDrawing record exceeds 8224 bytes".to_string(),
        ));
    }
    write_record_header(writer, MSODRAWING, length)?;
    writer.write_all(data)?;
    Ok(())
}

fn group_prefix(drawing_id: u32, comment_count: usize) -> XlsResult<Vec<u8>> {
    let shape_count = u32::try_from(comment_count + 1)
        .map_err(|_| XlsError::InvalidData("comment shape count overflows".to_string()))?;
    let shapes_size = u32::try_from(comment_count).unwrap() * 134;
    let spgr_length = 48u32
        .checked_add(shapes_size)
        .ok_or_else(|| XlsError::InvalidData("comment drawing size overflows".to_string()))?;
    let dg_length = 16u32
        .checked_add(8 + spgr_length)
        .ok_or_else(|| XlsError::InvalidData("comment drawing size overflows".to_string()))?;
    let patriarch = drawing_id << 10;
    let mut out = Vec::with_capacity(80);
    escher_header(&mut out, 0x000F, 0xF002, dg_length);
    escher_header(&mut out, (drawing_id as u16) << 4, 0xF008, 8);
    out.extend_from_slice(&shape_count.to_le_bytes());
    out.extend_from_slice(&(patriarch + shape_count - 1).to_le_bytes());
    escher_header(&mut out, 0x000F, 0xF003, spgr_length);
    escher_header(&mut out, 0x000F, 0xF004, 40);
    escher_header(&mut out, 0x0001, 0xF009, 16);
    out.extend_from_slice(&[0; 16]);
    escher_header(&mut out, 0x0002, 0xF00A, 8);
    out.extend_from_slice(&patriarch.to_le_bytes());
    out.extend_from_slice(&5u32.to_le_bytes());
    Ok(out)
}

fn comment_shape(config: &CommentConfig<'_>, shape_id: u32) -> Vec<u8> {
    let mut out = Vec::with_capacity(126);
    escher_header(&mut out, 0x000F, 0xF004, 126);
    escher_header(&mut out, 0x0CA2, 0xF00A, 8);
    out.extend_from_slice(&shape_id.to_le_bytes());
    out.extend_from_slice(&0x0000_0A00u32.to_le_bytes());
    escher_header(&mut out, 0x00A3, 0xF00B, 60);
    for (id, value) in [
        (0x0080u16, 0x1FF4_0000u32),
        (0x0141, 0x0002_8B01),
        (0x00BF, 0x0008_0008),
        (0x0158, 0),
        (0x0181, 0x0800_0050),
        (0x0183, 0x0800_0050),
        (0x01BF, 0x0010_0010),
        (0x0201, 0),
        (0x023F, 0x0003_0003),
        (
            0x03BF,
            if config.visible {
                0x0002_0000
            } else {
                0x0002_0002
            },
        ),
    ] {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&value.to_le_bytes());
    }
    escher_header(&mut out, 0, 0xF010, 18);
    let anchor = config.anchor;
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
        out.extend_from_slice(&value.to_le_bytes());
    }
    escher_header(&mut out, 0, 0xF011, 0);
    out
}

fn write_obj<W: Write>(writer: &mut W, config: &CommentConfig<'_>) -> XlsResult<()> {
    write_record_header(writer, OBJ, 52)?;
    writer.write_all(&0x0015u16.to_le_bytes())?;
    writer.write_all(&0x0012u16.to_le_bytes())?;
    writer.write_all(&0x0019u16.to_le_bytes())?;
    writer.write_all(&config.object_id.to_le_bytes())?;
    writer.write_all(&0x4011u16.to_le_bytes())?;
    writer.write_all(&[0; 12])?;
    writer.write_all(&0x000Du16.to_le_bytes())?;
    writer.write_all(&0x0016u16.to_le_bytes())?;
    writer.write_all(&config.guid)?;
    writer.write_all(&u16::from(config.shared).to_le_bytes())?;
    writer.write_all(&[0; 4])?;
    writer.write_all(&[0; 4])?;
    Ok(())
}

fn write_continue<W: Write>(writer: &mut W, data: &[u8]) -> XlsResult<()> {
    let length = u16::try_from(data.len())
        .map_err(|_| XlsError::InvalidData("comment CONTINUE record is too large".to_string()))?;
    if length > 8224 {
        return Err(XlsError::InvalidData(
            "comment CONTINUE exceeds 8224 bytes".to_string(),
        ));
    }
    write_record_header(writer, CONTINUE, length)?;
    writer.write_all(data)?;
    Ok(())
}

fn write_txo<W: Write>(writer: &mut W, config: &CommentConfig<'_>) -> XlsResult<()> {
    let units: Vec<u16> = config.text.encode_utf16().collect();
    let runs: Vec<XlsCommentTextRunWrite> = if units.is_empty() {
        Vec::new()
    } else if config.text_runs.is_empty() {
        vec![XlsCommentTextRunWrite {
            character_index: 0,
            font_index: 0,
        }]
    } else {
        config.text_runs.to_vec()
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
    writer.write_all(&config.font_when_empty.to_le_bytes())?;
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

fn write_note<W: Write>(writer: &mut W, config: &CommentConfig<'_>) -> XlsResult<()> {
    let author: Vec<u16> = config.author.encode_utf16().collect();
    let compressed = author.iter().all(|unit| *unit <= 0x00FF);
    let byte_count = author.len() * if compressed { 1 } else { 2 };
    write_record_header(writer, NOTE, (12 + byte_count) as u16)?;
    writer.write_all(&config.row.to_le_bytes())?;
    writer.write_all(&u16::from(config.column).to_le_bytes())?;
    writer.write_all(&(if config.visible { 2u16 } else { 0 }).to_le_bytes())?;
    writer.write_all(&config.object_id.to_le_bytes())?;
    writer.write_all(&(author.len() as u16).to_le_bytes())?;
    writer.write_all(&[u8::from(!compressed)])?;
    if compressed {
        for unit in author {
            writer.write_all(&[unit as u8])?;
        }
    } else {
        for unit in author {
            writer.write_all(&unit.to_le_bytes())?;
        }
    }
    writer.write_all(&[0])?;
    Ok(())
}

pub(crate) fn write_comments<W: Write>(
    writer: &mut W,
    drawing_id: u32,
    comments: &[CommentConfig<'_>],
) -> XlsResult<()> {
    if comments.is_empty() {
        return Ok(());
    }
    let prefix = group_prefix(drawing_id, comments.len())?;
    for (index, comment) in comments.iter().enumerate() {
        let mut drawing = if index == 0 {
            prefix.clone()
        } else {
            Vec::new()
        };
        drawing.extend_from_slice(&comment_shape(
            comment,
            (drawing_id << 10) + index as u32 + 1,
        ));
        write_mso(writer, &drawing)?;
        write_obj(writer, comment)?;
        let mut textbox = Vec::with_capacity(8);
        escher_header(&mut textbox, 0, 0xF00D, 0);
        write_mso(writer, &textbox)?;
        write_txo(writer, comment)?;
    }
    for comment in comments {
        write_note(writer, comment)?;
    }
    Ok(())
}
