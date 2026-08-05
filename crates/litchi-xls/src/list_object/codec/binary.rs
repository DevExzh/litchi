//! Bounded BIFF8 primitives for worksheet-table payloads.

use super::super::model::*;
use super::super::{FEATURE11_RECORD_TYPE, invalid};
use crate::Result;

pub(in crate::list_object) struct PendingFeature {
    pub(in crate::list_object) record_type: u16,
    pub(in crate::list_object) base: Vec<u8>,
    pub(in crate::list_object) continuations: Vec<Vec<u8>>,
    pub(in crate::list_object) combined: Vec<u8>,
}
pub(in crate::list_object) fn u16_at(
    data: &[u8],
    offset: usize,
    rt: u16,
    field: &str,
) -> Result<u16> {
    data.get(offset..offset + 2)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))
}
pub(in crate::list_object) fn u32_at(
    data: &[u8],
    offset: usize,
    rt: u16,
    field: &str,
) -> Result<u32> {
    data.get(offset..offset + 4)
        .map(|v| u32::from_le_bytes(v.try_into().unwrap()))
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))
}
pub(in crate::list_object) fn append_range(out: &mut Vec<u8>, range: ListObjectRange) {
    out.extend_from_slice(&range.first_row.to_le_bytes());
    out.extend_from_slice(&range.last_row.to_le_bytes());
    out.extend_from_slice(&range.first_column.to_le_bytes());
    out.extend_from_slice(&range.last_column.to_le_bytes());
}
pub(in crate::list_object) fn parse_range(
    data: &[u8],
    offset: usize,
    rt: u16,
) -> Result<ListObjectRange> {
    ListObjectRange::try_new(
        u16_at(data, offset, rt, "rwFirst")?,
        u16_at(data, offset + 2, rt, "rwLast")?,
        u16_at(data, offset + 4, rt, "colFirst")?,
        u16_at(data, offset + 6, rt, "colLast")?,
    )
}
pub(in crate::list_object) fn append_frt(
    out: &mut Vec<u8>,
    rt: u16,
    range: Option<ListObjectRange>,
) {
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&u16::from(range.is_some()).to_le_bytes());
    if let Some(range) = range {
        append_range(out, range);
    } else {
        out.extend_from_slice(&[0; 8]);
    }
}
pub(in crate::list_object) fn validate_frt(data: &[u8], rt: u16, reference: bool) -> Result<()> {
    if u16_at(data, 0, rt, "frt.rt")? != rt
        || u16_at(data, 2, rt, "frt.flags")? != u16::from(reference)
    {
        return Err(invalid(rt, "future-record header is invalid"));
    }
    if !reference && data.get(4..12).is_none_or(|v| v.iter().any(|b| *b != 0)) {
        return Err(invalid(rt, "future-record reserved bytes must be zero"));
    }
    Ok(())
}
pub(in crate::list_object) fn validate_frt_any(data: &[u8], rt: u16) -> Result<()> {
    if u16_at(data, 0, rt, "frt.rt")? != rt {
        return Err(invalid(rt, "future-record type echo is invalid"));
    }
    let flags = u16_at(data, 2, rt, "frt.flags")?;
    if flags & 0x0002 != 0 {
        return Err(invalid(rt, "future-record alert flag must be zero"));
    }
    if flags & 0x0001 == 0 && data.get(4..12).is_none_or(|v| v.iter().any(|b| *b != 0)) {
        return Err(invalid(
            rt,
            "future-record reference is present without fFrtRef",
        ));
    }
    Ok(())
}
pub(in crate::list_object) fn record(rt: u16, payload: Vec<u8>) -> Result<Vec<u8>> {
    let len =
        u16::try_from(payload.len()).map_err(|_| invalid(rt, "payload exceeds BIFF8 length"))?;
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&payload);
    Ok(out)
}
pub(in crate::list_object) fn append_string(out: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().collect::<Vec<_>>();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    if value.is_ascii() {
        out.push(0);
        out.extend_from_slice(value.as_bytes());
    } else {
        out.push(1);
        out.extend(units.into_iter().flat_map(u16::to_le_bytes));
    }
}
pub(in crate::list_object) fn parse_string(
    data: &[u8],
    offset: usize,
    rt: u16,
    field: &str,
) -> Result<(String, usize)> {
    let count = usize::from(u16_at(data, offset, rt, field)?);
    let flags = *data
        .get(offset + 2)
        .ok_or_else(|| invalid(rt, format!("truncated {field} flags")))?;
    if flags & !1 != 0 {
        return Err(invalid(rt, format!("{field} flags are unsupported")));
    }
    let width = if flags == 0 { 1 } else { 2 };
    let end = offset
        .checked_add(3)
        .and_then(|v| count.checked_mul(width).and_then(|n| v.checked_add(n)))
        .ok_or_else(|| invalid(rt, format!("{field} length overflows")))?;
    let bytes = data
        .get(offset + 3..end)
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))?;
    let value = if width == 1 {
        bytes.iter().map(|b| char::from(*b)).collect()
    } else {
        char::decode_utf16(
            bytes
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]])),
        )
        .collect::<Result<String, _>>()
        .map_err(|_| invalid(rt, format!("invalid UTF-16 in {field}")))?
    };
    Ok((value, end))
}

#[derive(Clone, Copy)]
pub(super) enum FormulaExtraKind {
    Array,
    Memory,
}

pub(super) fn parse_list_formula_extra_end(
    data: &[u8],
    tokens: &[u8],
    mut offset: usize,
    rt: u16,
) -> Result<usize> {
    let mut extras = Vec::new();
    let mut position = 0usize;
    while position < tokens.len() {
        let opcode = tokens[position];
        let base = if opcode < 0x20 {
            opcode
        } else {
            (opcode & 0x1f) | 0x20
        };
        let size = match base {
            0x03..=0x16 => 1,
            0x17 => {
                let count = usize::from(
                    *tokens
                        .get(position + 1)
                        .ok_or_else(|| invalid(rt, "truncated formula string token"))?,
                );
                let flags = *tokens
                    .get(position + 2)
                    .ok_or_else(|| invalid(rt, "truncated formula string flags"))?;
                if flags & !1 != 0 {
                    return Err(invalid(rt, "unsupported formula string flags"));
                }
                3usize
                    .checked_add(
                        count
                            .checked_mul(if flags == 0 { 1 } else { 2 })
                            .ok_or_else(|| invalid(rt, "formula string length overflows"))?,
                    )
                    .ok_or_else(|| invalid(rt, "formula string length overflows"))?
            },
            0x19 => {
                let header = tokens
                    .get(position..position + 4)
                    .ok_or_else(|| invalid(rt, "truncated Attr token"))?;
                if header[1] & 0x04 != 0 {
                    4usize
                        .checked_add(
                            (usize::from(u16::from_le_bytes([header[2], header[3]])) + 1) * 2,
                        )
                        .ok_or_else(|| invalid(rt, "Attr token length overflows"))?
                } else {
                    4
                }
            },
            0x1c | 0x1d => 2,
            0x1e => 3,
            0x1f => 9,
            0x20 => {
                extras.push(FormulaExtraKind::Array);
                8
            },
            0x21 => 3,
            0x22 => 4,
            0x23 | 0x24 | 0x2a | 0x2c => 5,
            0x25 | 0x2b | 0x2d => 9,
            0x26 => {
                extras.push(FormulaExtraKind::Memory);
                7
            },
            0x27 => 7,
            0x29 => 3,
            0x39 | 0x3a | 0x3c => 7,
            0x3b | 0x3d => 11,
            _ => {
                return Err(invalid(
                    rt,
                    "invalid or forbidden token in list array formula",
                ));
            },
        };
        position = position
            .checked_add(size)
            .ok_or_else(|| invalid(rt, "formula token length overflows"))?;
        if position > tokens.len() {
            return Err(invalid(rt, "truncated formula token"));
        }
    }
    for extra in extras {
        match extra {
            FormulaExtraKind::Memory => {
                let count = usize::from(u16_at(data, offset, rt, "PtgExtraMem count")?);
                offset = offset
                    .checked_add(2)
                    .and_then(|value| value.checked_add(count.checked_mul(8)?))
                    .ok_or_else(|| invalid(rt, "PtgExtraMem length overflows"))?;
                data.get(..offset)
                    .ok_or_else(|| invalid(rt, "truncated PtgExtraMem"))?;
            },
            FormulaExtraKind::Array => {
                let dimensions = data
                    .get(offset..offset + 3)
                    .ok_or_else(|| invalid(rt, "truncated PtgExtraArray dimensions"))?;
                let count = (usize::from(dimensions[0]) + 1)
                    .checked_mul(
                        usize::from(u16::from_le_bytes([dimensions[1], dimensions[2]])) + 1,
                    )
                    .ok_or_else(|| invalid(rt, "PtgExtraArray dimensions overflow"))?;
                offset += 3;
                for _ in 0..count {
                    let kind = *data
                        .get(offset)
                        .ok_or_else(|| invalid(rt, "truncated PtgExtraArray value"))?;
                    offset += 1;
                    match kind {
                        0 | 1 | 4 | 16 => {
                            offset = offset
                                .checked_add(8)
                                .ok_or_else(|| invalid(rt, "PtgExtraArray length overflows"))?;
                            data.get(..offset)
                                .ok_or_else(|| invalid(rt, "truncated PtgExtraArray value"))?;
                        },
                        2 => {
                            offset = parse_string(data, offset, rt, "PtgExtraArray string")?.1;
                        },
                        _ => return Err(invalid(rt, "invalid PtgExtraArray value type")),
                    }
                }
            },
        }
    }
    Ok(offset)
}

pub(super) fn append_formula(out: &mut Vec<u8>, tokens: &[u8]) -> Result<()> {
    let len = u16::try_from(tokens.len())
        .map_err(|_| invalid(FEATURE11_RECORD_TYPE, "formula token length exceeds 65535"))?;
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(tokens);
    Ok(())
}
pub(super) fn parse_formula(
    data: &[u8],
    offset: &mut usize,
    rt: u16,
    field: &str,
) -> Result<Vec<u8>> {
    let len = usize::from(u16_at(data, *offset, rt, field)?);
    if len == 0 {
        return Err(invalid(rt, format!("empty {field}")));
    }
    let end = (*offset)
        .checked_add(2 + len)
        .ok_or_else(|| invalid(rt, format!("{field} length overflows")))?;
    let value = data
        .get(*offset + 2..end)
        .ok_or_else(|| invalid(rt, format!("truncated {field}")))?
        .to_vec();
    *offset = end;
    Ok(value)
}
pub(super) fn append_web_info(out: &mut Vec<u8>, info: &WebFieldInfo) -> Result<()> {
    out.extend_from_slice(&info.locale.to_le_bytes());
    out.extend_from_slice(&info.decimal_places.to_le_bytes());
    let flags1 = u32::from(info.percent)
        | (u32::from(info.fixed_decimal) << 1)
        | (u32::from(info.date_only) << 2)
        | (info.reading_order.code() << 3)
        | (u32::from(info.rich_text) << 5)
        | (u32::from(info.unknown_rich_text) << 6)
        | (u32::from(info.alert_unknown_rich_text) << 7)
        | info.ignored_display_flags;
    out.extend_from_slice(&flags1.to_le_bytes());
    let default_type = match info.default_value {
        None => 0,
        Some(WebDefaultValue::String(_)) => 1,
        Some(WebDefaultValue::Boolean(_)) => 2,
        Some(WebDefaultValue::Number(_) | WebDefaultValue::DateTime(_)) => 3,
    };
    let flags2 = u32::from(info.read_only)
        | (u32::from(info.required) << 1)
        | (u32::from(info.minimum_set) << 2)
        | (u32::from(info.maximum_set) << 3)
        | (u32::from(info.default_value.is_some()) << 4)
        | (u32::from(info.default_today) << 5)
        | (u32::from(info.validation_formula.is_some()) << 6)
        | (u32::from(info.allow_fill_in) << 7)
        | (default_type << 8)
        | info.ignored_validation_flags;
    out.extend_from_slice(&flags2.to_le_bytes());
    if let Some(value) = &info.default_value {
        match value {
            WebDefaultValue::String(v) => append_string(out, v),
            WebDefaultValue::Boolean(v) => out.extend_from_slice(&u32::from(*v).to_le_bytes()),
            WebDefaultValue::Number(v) | WebDefaultValue::DateTime(v) => {
                out.extend_from_slice(&v.to_le_bytes())
            },
        }
    }
    if let Some(v) = &info.validation_formula {
        append_string(out, v)
    }
    out.extend_from_slice(&0u32.to_le_bytes());
    Ok(())
}
pub(super) fn parse_web_info(
    data: &[u8],
    offset: &mut usize,
    kind: WebColumnType,
    rt: u16,
) -> Result<WebFieldInfo> {
    let locale = u32_at(data, *offset, rt, "Web LCID")?;
    let decimal_places = u32_at(data, *offset + 4, rt, "Web cDec")?;
    let a = u32_at(data, *offset + 8, rt, "Web display flags")?;
    let b = u32_at(data, *offset + 12, rt, "Web validation flags")?;
    let reading_order = WebReadingOrder::from_code((a >> 3) & 3)?;
    let default_set = b & 0x10 != 0;
    let default_type = ((b >> 8) & 0xff) as u8;
    *offset += 16;
    let default_value = if default_set {
        Some(match (default_type, kind) {
            (1, WebColumnType::Text | WebColumnType::Choice | WebColumnType::MultipleChoices) => {
                let (v, end) = parse_string(data, *offset, rt, "Web default string")?;
                *offset = end;
                WebDefaultValue::String(v)
            },
            (2, WebColumnType::Boolean) => {
                let v = u32_at(data, *offset, rt, "Web default boolean")?;
                if v > 1 {
                    return Err(invalid(rt, "invalid Web default boolean"));
                }
                *offset += 4;
                WebDefaultValue::Boolean(v != 0)
            },
            (3, WebColumnType::Number | WebColumnType::Currency | WebColumnType::DateTime) => {
                let bytes = data
                    .get(*offset..*offset + 8)
                    .ok_or_else(|| invalid(rt, "truncated Web default number"))?;
                *offset += 8;
                let v = f64::from_le_bytes(bytes.try_into().unwrap());
                if kind == WebColumnType::DateTime {
                    WebDefaultValue::DateTime(v)
                } else {
                    WebDefaultValue::Number(v)
                }
            },
            _ => return Err(invalid(rt, "Web default type does not match column type")),
        })
    } else {
        if default_type != 0 {
            return Err(invalid(rt, "Web default type exists without a default"));
        }
        None
    };
    let validation_formula = if b & 0x40 != 0 {
        let (v, end) = parse_string(data, *offset, rt, "Web validation formula")?;
        *offset = end;
        Some(v)
    } else {
        None
    };
    if u32_at(data, *offset, rt, "Web reserved")? != 0 {
        return Err(invalid(rt, "Web field-info reserved value must be zero"));
    }
    *offset += 4;
    let value = WebFieldInfo {
        locale,
        decimal_places,
        percent: a & 1 != 0,
        fixed_decimal: a & 2 != 0,
        date_only: a & 4 != 0,
        reading_order,
        rich_text: a & 0x20 != 0,
        unknown_rich_text: a & 0x40 != 0,
        alert_unknown_rich_text: a & 0x80 != 0,
        read_only: b & 1 != 0,
        required: b & 2 != 0,
        minimum_set: b & 4 != 0,
        maximum_set: b & 8 != 0,
        default_today: b & 0x20 != 0,
        allow_fill_in: b & 0x80 != 0,
        default_value,
        validation_formula,
        ignored_display_flags: a & !0xff,
        ignored_validation_flags: b & 0xffff_0000,
    };
    value.validate(kind)?;
    Ok(value)
}
