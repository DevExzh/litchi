//! Fixed-width BIFF and future-record wire primitives.

use crate::Result;
use crate::list_object::invalid;
use crate::list_object::model::ListObjectRange;

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
    let len = u16::try_from(payload.len())
        .map_err(|_error| invalid(rt, "payload exceeds BIFF8 length"))?;
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&rt.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&payload);
    Ok(out)
}
