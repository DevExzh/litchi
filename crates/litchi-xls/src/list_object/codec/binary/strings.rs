//! BIFF8 compressed and UTF-16 string primitives.

use super::primitives::u16_at;
use crate::Result;
use crate::list_object::invalid;

pub(in crate::list_object) fn append_string(out: &mut Vec<u8>, value: &str) {
    let units = value.encode_utf16().collect::<Vec<_>>();
    out.extend_from_slice(&crate::utils::truncate_usize_to_u16(units.len()).to_le_bytes());
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
        .map_err(|_error| invalid(rt, format!("invalid UTF-16 in {field}")))?
    };
    Ok((value, end))
}
