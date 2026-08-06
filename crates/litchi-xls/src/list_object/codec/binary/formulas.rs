//! Formula-token and array-extra wire primitives.

use super::primitives::u16_at;
use super::strings::parse_string;
use crate::Result;
use crate::list_object::{FEATURE11_RECORD_TYPE, invalid};

#[derive(Clone, Copy)]
pub(in crate::list_object) enum FormulaExtraKind {
    Array,
    Memory,
}

pub(in crate::list_object) fn parse_list_formula_extra_end(
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

pub(in crate::list_object) fn append_formula(out: &mut Vec<u8>, tokens: &[u8]) -> Result<()> {
    let len = u16::try_from(tokens.len())
        .map_err(|_| invalid(FEATURE11_RECORD_TYPE, "formula token length exceeds 65535"))?;
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(tokens);
    Ok(())
}

pub(in crate::list_object) fn parse_formula(
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
