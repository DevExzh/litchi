//! DOC and `ObjInfo` wire parsing.

use super::{
    CLX, MAX_FIELDS, MAX_PIECES, OBJ_INFO_STREAM, ODTPERSIST1_MUST_BE_ZERO, ODTPERSIST1_RESERVED,
    ODTPERSIST2_MUST_BE_ZERO, ODTPERSIST2_RESERVED, PLCFFLD_MOM, array_at, corrupted, fib_pair,
    slice, u32_at, word,
};
use crate::package::Result;
use litchi_ole_common::object::Object;

use super::super::model::{FieldMarker, Info, RawPiece, Reference};

impl Info {
    pub fn read(data: &[u8]) -> Result<Self> {
        if data.len() != 4 && data.len() != 6 {
            return Err(corrupted("ObjInfo ODT must be 4 or 6 bytes"));
        }
        let first = word(data, 0)?;
        if first & ODTPERSIST1_MUST_BE_ZERO != 0 {
            return Err(corrupted("ObjInfo ODTPersist1 MUST-be-zero bits are set"));
        }
        let activex = first & (1 << 12) != 0;
        let stream_control = first & (1 << 13) != 0;
        if stream_control && !activex {
            return Err(corrupted("ObjInfo stream control requires ActiveX"));
        }
        let second = if data.len() == 6 { word(data, 4)? } else { 0 };
        if second & ODTPERSIST2_MUST_BE_ZERO != 0 {
            return Err(corrupted("ObjInfo ODTPersist2 MUST-be-zero bit is set"));
        }
        Ok(Self {
            persist2_present: data.len() == 6,
            default_handler: first & (1 << 1) != 0,
            linked: first & (1 << 4) != 0,
            display_as_icon: first & (1 << 6) != 0,
            ole1: first & (1 << 7) != 0,
            manual_update: first & (1 << 8) != 0,
            recompose_on_resize: first & (1 << 9) != 0,
            activex,
            stream_control,
            view_object: first & (1 << 15) != 0,
            enhanced_metafile: second & 1 != 0,
            queried_enhanced_metafile: second & 4 != 0,
            stored_as_enhanced_metafile: second & 8 != 0,
            clipboard_format: word(data, 2)?,
            reserved_persist1: first & ODTPERSIST1_RESERVED,
            reserved_persist2: second & ODTPERSIST2_RESERVED,
        })
    }

    /// Reads this metadata from the opaque DOC `ObjInfo` stream of an object.
    pub fn of(object: &Object) -> Result<Option<Self>> {
        object
            .stream(&[OBJ_INFO_STREAM])
            .map(Self::read)
            .transpose()
    }
}

pub(in crate::embedded_object) fn managed_objects(
    word: &[u8],
    pieces: &[RawPiece],
    fields: &[FieldMarker],
) -> Result<Vec<Reference>> {
    let mut stack: Vec<(u32, u8, Option<u32>)> = Vec::new();
    let mut output = Vec::new();
    for marker in fields {
        match marker.descriptor[0] & 0x1F {
            0x13 => stack.push((marker.cp, marker.descriptor[1], None)),
            0x14 => {
                if let Some(value) = stack.last_mut() {
                    value.2 = Some(marker.cp);
                }
            },
            0x15 => {
                let Some((start, kind, Some(separator))) = stack.pop() else {
                    continue;
                };
                if kind != 0x3A || !stack.is_empty() {
                    continue;
                }
                let code = text_range(word, pieces, start + 1, separator)?;
                let Some(id_text) = [" EMBED LITCHI_OBJECT _", " EMBED Equation.3 _"]
                    .into_iter()
                    .find_map(|prefix| code.strip_prefix(prefix).and_then(|v| v.strip_suffix(' ')))
                else {
                    continue;
                };
                let Ok(storage_id) = id_text.parse::<u32>() else {
                    continue;
                };
                let end = marker.cp;
                if pieces
                    .iter()
                    .any(|piece| piece.start == start && piece.end == end + 1)
                {
                    output.push(Reference {
                        storage_id,
                        storage_name: format!("_{storage_id}"),
                        start_cp: start,
                        separator_cp: separator,
                        end_cp: end,
                        data_offset: 0,
                    });
                }
            },
            _ => {},
        }
    }
    Ok(output)
}

pub(in crate::embedded_object) fn text_range(
    word: &[u8],
    pieces: &[RawPiece],
    start: u32,
    end: u32,
) -> Result<String> {
    let mut units = Vec::new();
    let mut cp = start;
    while cp < end {
        let piece = pieces
            .iter()
            .find(|piece| cp >= piece.start && cp < piece.end)
            .ok_or_else(|| corrupted("field code CP has no text piece"))?;
        let take_end = end.min(piece.end);
        let count = (take_end - cp) as usize;
        let relative = (cp - piece.start) as usize;
        if piece.unicode {
            let offset = piece.fc as usize + relative * 2;
            let bytes = word
                .get(offset..offset + count * 2)
                .ok_or_else(|| corrupted("field code exceeds WordDocument"))?;
            for pair in bytes.chunks_exact(2) {
                units.push(u16::from_le_bytes([pair[0], pair[1]]));
            }
        } else {
            let offset = piece.fc as usize + relative;
            let bytes = word
                .get(offset..offset + count)
                .ok_or_else(|| corrupted("field code exceeds WordDocument"))?;
            units.extend(bytes.iter().map(|byte| u16::from(*byte)));
        }
        cp = take_end;
    }
    String::from_utf16(&units).map_err(|_| corrupted("field instruction contains invalid UTF-16"))
}

pub(in crate::embedded_object) fn parse_clx(word: &[u8], table: &[u8]) -> Result<Vec<RawPiece>> {
    let (offset, length) = fib_pair(word, CLX)?;
    let data = slice(table, offset, length, "CLX")?;
    if data.first() != Some(&2) {
        return Err(corrupted("fast-save RgPrc CLX is unsupported"));
    }
    let size = u32_at(data, 1)? as usize;
    if size + 5 != data.len() || size < 4 || !(size - 4).is_multiple_of(12) {
        return Err(corrupted("CLX PlcPcd size is invalid"));
    }
    let count = (size - 4) / 12;
    if count == 0 || count > MAX_PIECES {
        return Err(corrupted("piece count is unsupported"));
    }
    let cps = &data[5..5 + (count + 1) * 4];
    let pcds = &data[5 + (count + 1) * 4..];
    let mut pieces = Vec::with_capacity(count);
    for index in 0..count {
        let start = u32_at(cps, index * 4)?;
        let end = u32_at(cps, (index + 1) * 4)?;
        if start >= end
            || pieces
                .last()
                .is_some_and(|last: &RawPiece| last.end != start)
        {
            return Err(corrupted("piece CPs overlap or contain gaps"));
        }
        let pcd = pcds
            .get(index * 8..index * 8 + 8)
            .ok_or_else(|| corrupted("piece descriptor is truncated"))?;
        let raw_fc = u32_at(pcd, 2)?;
        let unicode = raw_fc & 0x4000_0000 == 0;
        let fc = if unicode {
            raw_fc & 0x3FFF_FFFF
        } else {
            (raw_fc & 0x3FFF_FFFF) / 2
        };
        let byte_len = (end - start)
            .checked_mul(if unicode { 2 } else { 1 })
            .ok_or_else(|| corrupted("piece length overflow"))?;
        if fc
            .checked_add(byte_len)
            .is_none_or(|end| end as usize > word.len())
        {
            return Err(corrupted("piece text exceeds WordDocument"));
        }
        pieces.push(RawPiece {
            start,
            end,
            fc,
            unicode,
            pcd_prefix: array_at(pcd, 0, "piece descriptor prefix")?,
            prm: array_at(pcd, 6, "piece descriptor PRM")?,
        });
    }
    if pieces.iter().any(|piece| piece.prm != [0, 0]) {
        return Err(corrupted("piece-level fast-save SPRMs are unsupported"));
    }
    Ok(pieces)
}

pub(in crate::embedded_object) fn parse_fields(
    word: &[u8],
    table: &[u8],
    main_ccp: u32,
) -> Result<Vec<FieldMarker>> {
    let (offset, length) = fib_pair(word, PLCFFLD_MOM)?;
    if length == 0 {
        return Ok(Vec::new());
    }
    let data = slice(table, offset, length, "PlcfFldMom")?;
    if data.len() < 4 || (data.len() - 4) % 6 != 0 {
        return Err(corrupted("PlcfFldMom length is invalid"));
    }
    let count = (data.len() - 4) / 6;
    if count > MAX_FIELDS {
        return Err(corrupted("field count exceeds limit"));
    }
    let cp_bytes = (count + 1) * 4;
    let mut output = Vec::with_capacity(count);
    for index in 0..count {
        let cp = u32_at(data, index * 4)?;
        if cp >= main_ccp
            || output
                .last()
                .is_some_and(|last: &FieldMarker| last.cp >= cp)
        {
            return Err(corrupted("field marker CPs are invalid"));
        }
        output.push(FieldMarker {
            cp,
            descriptor: array_at(data, cp_bytes + index * 2, "field descriptor")?,
        });
    }
    Ok(output)
}

pub(in crate::embedded_object) fn parse_bte(
    word: &[u8],
    table: &[u8],
    index: usize,
) -> Result<(Vec<u32>, Vec<u32>)> {
    let (offset, length) = fib_pair(word, index)?;
    let data = slice(table, offset, length, "PlcBteChpx")?;
    if data.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }
    if data.len() < 4 || (data.len() - 4) % 8 != 0 {
        return Err(corrupted(format!(
            "PlcBteChpx length {} is invalid",
            data.len()
        )));
    }
    let count = (data.len() - 4) / 8;
    let mut fc = Vec::with_capacity(count + 1);
    let mut pages = Vec::with_capacity(count);
    for i in 0..=count {
        fc.push(u32_at(data, i * 4)?);
    }
    for i in 0..count {
        pages.push(u32_at(data, (count + 1) * 4 + i * 4)?);
    }
    if fc.windows(2).any(|v| v[0] >= v[1])
        || pages
            .iter()
            .any(|pn| (*pn as usize) * 512 + 512 > word.len())
    {
        return Err(corrupted("PlcBteChpx references invalid FKPs"));
    }
    Ok((fc, pages))
}
