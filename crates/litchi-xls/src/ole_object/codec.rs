//! Lossless BIFF8 Obj and XLUnicodeString codecs.

use super::model::*;
use super::*;
use crate::error::{Error, Result};

fn parse_formula(body: &[u8]) -> Result<FtPictFmla> {
    if body.len() < 2 {
        return Err(invalid(OBJ, "FtPictFmla is truncated"));
    }
    let len = usize::from(u16::from_le_bytes([body[0], body[1]]));
    let end = 2usize
        .checked_add(len)
        .ok_or_else(|| invalid(OBJ, "formula overflow"))?;
    let formula = body
        .get(2..end)
        .ok_or_else(|| invalid(OBJ, "formula is truncated"))?
        .to_vec();
    let tail = &body[end..];
    let (storage_position, control_buffer_size) = match tail.len() {
        0 => (None, None),
        8 => (
            Some(u32_at(tail, 0).ok_or_else(|| invalid(OBJ, "storage position is truncated"))?),
            Some(u32_at(tail, 4).ok_or_else(|| invalid(OBJ, "control buffer size is truncated"))?),
        ),
        _ => return Err(invalid(OBJ, "unsupported FtPictFmla trailing layout")),
    };
    Ok(FtPictFmla {
        formula,
        storage_position,
        control_buffer_size,
    })
}

pub(super) fn parse_subrecords(data: &[u8]) -> Result<Vec<ObjSubrecord>> {
    let mut offset = 0usize;
    let mut control_type = None;
    let mut subrecords = Vec::new();
    while offset < data.len() {
        let header = data
            .get(offset..offset + 4)
            .ok_or_else(|| invalid(OBJ, "truncated Obj subrecord header"))?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        offset += 4;
        let end = offset
            .checked_add(len)
            .ok_or_else(|| invalid(OBJ, "Obj subrecord overflow"))?;
        let body = data
            .get(offset..end)
            .ok_or_else(|| invalid(OBJ, "truncated Obj subrecord"))?;
        let value =
            match (kind, len) {
                (FT_CMO, 18) => ObjSubrecord::Common(FtCmo {
                    object_type: u16::from_le_bytes([body[0], body[1]]),
                    object_id: u16::from_le_bytes([body[2], body[3]]),
                    flags: u16::from_le_bytes([body[4], body[5]]),
                    reserved: array_at(body, 6)
                        .ok_or_else(|| invalid(OBJ, "FtCmo reserved bytes are truncated"))?,
                }),
                (FT_CMO, _) => return Err(invalid(OBJ, "FtCmo must contain 18 bytes")),
                (FT_CF, _) => ObjSubrecord::ClipboardFormat(body.to_vec()),
                (FT_PIO, 2) => ObjSubrecord::PictureFlags(FtPioGrbit {
                    raw: u16::from_le_bytes([body[0], body[1]]),
                }),
                (FT_PIO, _) => return Err(invalid(OBJ, "FtPioGrbit must contain 2 bytes")),
                (FT_PICT_FMLA, _) => ObjSubrecord::PictureFormula(parse_formula(body)?),
                // Form-control data subrecords fall back to raw preservation when
                // their contents do not match the MS-XLS layout.
                (FT_CBLS_DATA, _) => parse_cbls_data(body)
                    .map_or_else(|| unknown(kind, body), ObjSubrecord::CheckBoxData),
                (FT_RBO_DATA, _) => parse_rbo_data(body)
                    .map_or_else(|| unknown(kind, body), ObjSubrecord::RadioButtonData),
                (FT_EDO_DATA, _) => parse_edo_data(body)
                    .map_or_else(|| unknown(kind, body), ObjSubrecord::EditBoxData),
                (FT_GBO_DATA, _) => parse_gbo_data(body)
                    .map_or_else(|| unknown(kind, body), ObjSubrecord::GroupBoxData),
                (FT_SBS, _) => {
                    parse_sbs(body).map_or_else(|| unknown(kind, body), ObjSubrecord::ScrollBarData)
                },
                (FT_LBS_DATA, _) => parse_lbs_data(body, control_type)
                    .map_or_else(|| unknown(kind, body), ObjSubrecord::ListBoxData),
                (FT_END, 0) => ObjSubrecord::End,
                (FT_END, _) => return Err(invalid(OBJ, "FtEnd must be empty")),
                _ => unknown(kind, body),
            };
        if let ObjSubrecord::Common(common) = &value {
            control_type = Some(common.object_type);
        }
        subrecords.push(value);
        offset = end;
    }
    Ok(subrecords)
}

fn unknown(kind: u16, body: &[u8]) -> ObjSubrecord {
    ObjSubrecord::Unknown {
        kind,
        data: body.to_vec(),
    }
}

pub(super) fn u16_at(data: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(array_at(data, offset)?))
}

pub(super) fn u32_at(data: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(array_at(data, offset)?))
}

fn array_at<const N: usize>(data: &[u8], offset: usize) -> Option<[u8; N]> {
    let end = offset.checked_add(N)?;
    data.get(offset..end)?.try_into().ok()
}

fn bool_at(data: &[u8], offset: usize) -> Option<bool> {
    match u16_at(data, offset)? {
        0 => Some(false),
        1 => Some(true),
        _ => None,
    }
}

fn parse_cbls_data(body: &[u8]) -> Option<FtCblsData> {
    if body.len() != 8 {
        return None;
    }
    Some(FtCblsData {
        state: CheckState::from_code(u16_at(body, 0)?)?,
        accelerator: u16_at(body, 2)?,
        reserved: u16_at(body, 4)?,
        flags: u16_at(body, 6)?,
    })
}

fn parse_rbo_data(body: &[u8]) -> Option<FtRboData> {
    if body.len() != 4 {
        return None;
    }
    Some(FtRboData {
        next_radio_button_id: u16_at(body, 0)?,
        first_in_group: bool_at(body, 2)?,
    })
}

fn parse_edo_data(body: &[u8]) -> Option<FtEdoData> {
    if body.len() != 8 {
        return None;
    }
    Some(FtEdoData {
        validation: EditBoxValidation::from_code(u16_at(body, 0)?)?,
        multi_line: bool_at(body, 2)?,
        vertical_scroll_bar: bool_at(body, 4)?,
        list_control_id: u16_at(body, 6)?,
    })
}

fn parse_gbo_data(body: &[u8]) -> Option<FtGboData> {
    if body.len() != 6 {
        return None;
    }
    Some(FtGboData {
        accelerator: u16_at(body, 0)?,
        reserved: u16_at(body, 2)?,
        flags: u16_at(body, 4)?,
    })
}

fn parse_sbs(body: &[u8]) -> Option<FtSbs> {
    if body.len() != 20 {
        return None;
    }
    let value = FtSbs {
        reserved: array_at(body, 0)?,
        value: i16::from_le_bytes(array_at(body, 4)?),
        minimum: i16::from_le_bytes(array_at(body, 6)?),
        maximum: i16::from_le_bytes(array_at(body, 8)?),
        increment: i16::from_le_bytes(array_at(body, 10)?),
        page_increment: i16::from_le_bytes(array_at(body, 12)?),
        horizontal: bool_at(body, 14)?,
        scroll_width: i16::from_le_bytes(array_at(body, 16)?),
        flags: u16_at(body, 18)?,
    };
    value.validate().ok()?;
    Some(value)
}

fn parse_lbs_data(body: &[u8], control_type: Option<u16>) -> Option<FtLbsData> {
    if body.is_empty() {
        return Some(FtLbsData::default());
    }
    let formula_len = usize::from(u16_at(body, 0)?);
    let formula_end = 2usize.checked_add(formula_len)?;
    let formula = body.get(2..formula_end)?.to_vec();
    let header_end = formula_end.checked_add(8)?;
    if body.len() < header_end {
        return None;
    }
    let mut data = FtLbsData {
        formula,
        entry_count: u16_at(body, formula_end)?,
        selected_index: u16_at(body, formula_end + 2)?,
        flags: u16_at(body, formula_end + 4)?,
        edit_box_id: u16_at(body, formula_end + 6)?,
        ..FtLbsData::default()
    };
    data.validate().ok()?;
    let mut offset = header_end;
    if control_type.and_then(ObjectType::from_code) == Some(ObjectType::DropDown) {
        let drop_header_end = offset.checked_add(6)?;
        if body.len() < drop_header_end {
            return None;
        }
        let flags = u16_at(body, offset)?;
        let line_count = u16_at(body, offset + 2)?;
        let min_width = u16_at(body, offset + 4)?;
        offset = drop_header_end;
        let text_len = xl_unicode_string_size(body.get(offset..)?)?;
        let text = LbsItem::parse(body.get(offset..offset + text_len)?.to_vec())?;
        offset += text_len;
        let padding = if text_len % 2 == 1 {
            let value = *body.get(offset)?;
            offset += 1;
            Some(value)
        } else {
            None
        };
        data.drop_down = Some(LbsDropData {
            flags,
            line_count,
            min_width,
            text,
            padding,
        });
        data.validate().ok()?;
    }
    // rgLines: parse up to `entry_count` item strings. A record continued into
    // Continue records holds fewer strings here; a defective string stops the
    // walk and its bytes are preserved verbatim as trailing data.
    let mut items = Vec::new();
    while items.len() < usize::from(data.entry_count) && offset < body.len() {
        match xl_unicode_string_size(&body[offset..]) {
            Some(size) if offset + size <= body.len() => {
                items.push(LbsItem::parse(body[offset..offset + size].to_vec())?);
                offset += size;
            },
            _ => break,
        }
    }
    if offset < body.len() {
        // bsels: one selection byte per entry for multiple-selection lists.
        let multiple = (data.flags >> LBS_SELECTION_TYPE_SHIFT) & LBS_SELECTION_TYPE_MASK != 0;
        if multiple {
            let count = usize::from(data.entry_count).min(body.len() - offset);
            let selection = &body[offset..offset + count];
            if selection.iter().all(|value| *value <= 1) {
                data.multi_selection = selection.iter().map(|value| *value != 0).collect();
                offset += count;
            }
        }
        data.trailing = body[offset..].to_vec();
    }
    data.set_items(items);
    data.validate().ok()?;
    Some(data)
}

/// Total byte size of the XLUnicodeString (MS-XLS 2.5.294) starting at
/// `data`, including formatting runs and extension data, or `None` when the
/// framing is truncated or inconsistent.
fn xl_unicode_string_size(data: &[u8]) -> Option<usize> {
    if data.len() < 3 {
        return None;
    }
    let character_count = usize::from(u16_at(data, 0)?);
    let options = *data.get(2)?;
    let mut offset = 3usize;
    let formatting_runs = if options & XL_STRING_RICH != 0 {
        let count = usize::from(u16_at(data, offset)?);
        offset += 2;
        count
    } else {
        0
    };
    let extension_size = if options & XL_STRING_EXT != 0 {
        let size = u32::from_le_bytes(data.get(offset..offset + 4)?.try_into().ok()?) as usize;
        offset += 4;
        size
    } else {
        0
    };
    let character_bytes = character_count.checked_mul(if options & XL_STRING_HIGH_BYTE != 0 {
        2
    } else {
        1
    })?;
    let total = offset
        .checked_add(character_bytes)?
        .checked_add(formatting_runs.checked_mul(FORMATTING_RUN_SIZE)?)?
        .checked_add(extension_size)?;
    if total > data.len() {
        return None;
    }
    Some(total)
}

/// Decode the text of an exact-size XLUnicodeString, ignoring formatting runs
/// and extension data.
pub(super) fn decode_xl_unicode_string(encoded: &[u8]) -> Option<String> {
    if xl_unicode_string_size(encoded)? != encoded.len() {
        return None;
    }
    let character_count = usize::from(u16_at(encoded, 0)?);
    let options = *encoded.get(2)?;
    let mut offset = 3usize;
    if options & XL_STRING_RICH != 0 {
        offset += 2;
    }
    if options & XL_STRING_EXT != 0 {
        offset += 4;
    }
    if options & XL_STRING_HIGH_BYTE != 0 {
        let bytes = encoded.get(offset..offset + character_count * 2)?;
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        String::from_utf16(&units.collect::<Vec<_>>()).ok()
    } else {
        let bytes = encoded.get(offset..offset + character_count)?;
        Some(bytes.iter().map(|value| char::from(*value)).collect())
    }
}

pub(super) fn serialize_subrecords(subrecords: &[ObjSubrecord]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for value in subrecords {
        let (kind, body) = serialize_subrecord(value)?;
        let len =
            u16::try_from(body.len()).map_err(|_| invalid(kind, "Obj subrecord exceeds u16"))?;
        output.extend_from_slice(&kind.to_le_bytes());
        output.extend_from_slice(&len.to_le_bytes());
        output.extend_from_slice(&body);
    }
    Ok(output)
}

fn serialize_subrecord(value: &ObjSubrecord) -> Result<(u16, Vec<u8>)> {
    Ok(match value {
        ObjSubrecord::Common(value) => {
            let mut body = Vec::with_capacity(18);
            body.extend_from_slice(&value.object_type.to_le_bytes());
            body.extend_from_slice(&value.object_id.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            body.extend_from_slice(&value.reserved);
            (FT_CMO, body)
        },
        ObjSubrecord::ClipboardFormat(data) => (FT_CF, data.clone()),
        ObjSubrecord::PictureFlags(value) => (FT_PIO, value.raw.to_le_bytes().to_vec()),
        ObjSubrecord::PictureFormula(value) => {
            let len = u16::try_from(value.formula.len())
                .map_err(|_| invalid(OBJ, "formula exceeds u16"))?;
            let mut body = len.to_le_bytes().to_vec();
            body.extend_from_slice(&value.formula);
            match (value.storage_position, value.control_buffer_size) {
                (Some(position), Some(size)) => {
                    body.extend_from_slice(&position.to_le_bytes());
                    body.extend_from_slice(&size.to_le_bytes());
                },
                (None, None) => {},
                _ => {
                    return Err(invalid(
                        OBJ,
                        "FtPictFmla optional fields must occur together",
                    ));
                },
            }
            (FT_PICT_FMLA, body)
        },
        ObjSubrecord::CheckBoxData(value) => {
            let mut body = Vec::with_capacity(8);
            body.extend_from_slice(&value.state.code().to_le_bytes());
            body.extend_from_slice(&value.accelerator.to_le_bytes());
            body.extend_from_slice(&value.reserved.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            (FT_CBLS_DATA, body)
        },
        ObjSubrecord::RadioButtonData(value) => {
            let mut body = Vec::with_capacity(4);
            body.extend_from_slice(&value.next_radio_button_id.to_le_bytes());
            body.extend_from_slice(&u16::from(value.first_in_group).to_le_bytes());
            (FT_RBO_DATA, body)
        },
        ObjSubrecord::EditBoxData(value) => {
            let mut body = Vec::with_capacity(8);
            body.extend_from_slice(&value.validation.code().to_le_bytes());
            body.extend_from_slice(&u16::from(value.multi_line).to_le_bytes());
            body.extend_from_slice(&u16::from(value.vertical_scroll_bar).to_le_bytes());
            body.extend_from_slice(&value.list_control_id.to_le_bytes());
            (FT_EDO_DATA, body)
        },
        ObjSubrecord::GroupBoxData(value) => {
            let mut body = Vec::with_capacity(6);
            body.extend_from_slice(&value.accelerator.to_le_bytes());
            body.extend_from_slice(&value.reserved.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            (FT_GBO_DATA, body)
        },
        ObjSubrecord::ScrollBarData(value) => {
            value.validate()?;
            let mut body = Vec::with_capacity(20);
            body.extend_from_slice(&value.reserved);
            body.extend_from_slice(&value.value.to_le_bytes());
            body.extend_from_slice(&value.minimum.to_le_bytes());
            body.extend_from_slice(&value.maximum.to_le_bytes());
            body.extend_from_slice(&value.increment.to_le_bytes());
            body.extend_from_slice(&value.page_increment.to_le_bytes());
            body.extend_from_slice(&u16::from(value.horizontal).to_le_bytes());
            body.extend_from_slice(&value.scroll_width.to_le_bytes());
            body.extend_from_slice(&value.flags.to_le_bytes());
            (FT_SBS, body)
        },
        ObjSubrecord::ListBoxData(value) => {
            value.validate()?;
            if value.is_vacant() {
                (FT_LBS_DATA, Vec::new())
            } else {
                let len = u16::try_from(value.formula.len())
                    .map_err(|_| invalid(OBJ, "ObjFmla exceeds u16"))?;
                let mut body = len.to_le_bytes().to_vec();
                body.extend_from_slice(&value.formula);
                body.extend_from_slice(&value.entry_count.to_le_bytes());
                body.extend_from_slice(&value.selected_index.to_le_bytes());
                body.extend_from_slice(&value.flags.to_le_bytes());
                body.extend_from_slice(&value.edit_box_id.to_le_bytes());
                if let Some(drop_down) = &value.drop_down {
                    body.extend_from_slice(&drop_down.flags.to_le_bytes());
                    body.extend_from_slice(&drop_down.line_count.to_le_bytes());
                    body.extend_from_slice(&drop_down.min_width.to_le_bytes());
                    body.extend_from_slice(&drop_down.text.encoded);
                    if drop_down.text.encoded.len() % 2 == 1 {
                        body.push(drop_down.padding.unwrap_or(0));
                    }
                }
                for item in value.items() {
                    body.extend_from_slice(item.encoded());
                }
                body.extend(
                    value
                        .multi_selection
                        .iter()
                        .map(|selected| u8::from(*selected)),
                );
                body.extend_from_slice(&value.trailing);
                (FT_LBS_DATA, body)
            }
        },
        ObjSubrecord::Unknown { kind, data } => (*kind, data.clone()),
        ObjSubrecord::End => (FT_END, Vec::new()),
    })
}

#[allow(clippy::type_complexity)]
pub(super) fn ranges(input: &[u8]) -> Result<Vec<(usize, usize, u16, usize, usize)>> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    while offset < input.len() {
        let header = input.get(offset..offset + 4).ok_or(Error::InvalidLength {
            expected: offset + 4,
            found: input.len(),
        })?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = offset
            .checked_add(4 + len)
            .ok_or_else(|| invalid(kind, "record size overflow"))?;
        if end > input.len() {
            return Err(Error::InvalidLength {
                expected: end,
                found: input.len(),
            });
        }
        output.push((offset, end, kind, offset + 4, end));
        offset = end;
    }
    Ok(output)
}

pub(super) fn record(kind: u16, body: &[u8]) -> Result<Vec<u8>> {
    if body.len() > 8_224 {
        return Err(invalid(kind, "record exceeds BIFF8 limit"));
    }
    let mut output = Vec::with_capacity(body.len() + 4);
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(body.len() as u16).to_le_bytes());
    output.extend_from_slice(body);
    Ok(output)
}
