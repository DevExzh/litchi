//! Formula-token and array-extra wire primitives.

use super::primitives::u16_at;
use super::strings::parse_string;
use crate::Result;
use crate::formula_metadata::{FormulaExtraKind, scan_list_extra_kinds};
use crate::list_object::{FEATURE11_RECORD_TYPE, invalid};

pub(in crate::list_object) fn parse_list_formula_extra_end(
    data: &[u8],
    tokens: &[u8],
    mut offset: usize,
    rt: u16,
) -> Result<usize> {
    let extras = scan_list_extra_kinds(tokens, rt).map_err(|error| match error {
        crate::Error::InvalidRecord { message, .. } => {
            if message.contains("unframeable token") {
                invalid(rt, "invalid or forbidden token in list array formula")
            } else {
                invalid(
                    rt,
                    message.replace("Formula RgbExtra", "list array formula"),
                )
            }
        },
        other => other,
    })?;
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
        .map_err(|_error| invalid(FEATURE11_RECORD_TYPE, "formula token length exceeds 65535"))?;
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
