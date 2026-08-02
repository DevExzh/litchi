use std::io::Write;

use crate::writer::XlsCommentTextRunWrite;
use crate::writer::shape::Anchor;
use crate::{XlsError, XlsResult};

use super::write_record_header;

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
    pub anchor: Anchor,
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

pub(super) fn comment_shape(config: &CommentConfig<'_>, shape_id: u32) -> Vec<u8> {
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
    for value in config.anchor.fields() {
        out.extend_from_slice(&value.to_le_bytes());
    }
    escher_header(&mut out, 0, 0xF011, 0);
    out
}

pub(super) fn write_obj<W: Write>(writer: &mut W, config: &CommentConfig<'_>) -> XlsResult<()> {
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

pub(super) fn write_txo<W: Write>(writer: &mut W, config: &CommentConfig<'_>) -> XlsResult<()> {
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

pub(super) fn write_note<W: Write>(writer: &mut W, config: &CommentConfig<'_>) -> XlsResult<()> {
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
