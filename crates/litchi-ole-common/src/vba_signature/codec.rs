//! Wire decoding for `[MS-OSHARED]` VBA signature containers.

use std::ops::Range;

use super::model::{Error, Kind, Limits};
use super::validation;

pub(crate) const INFO_HEADER_SIZE: usize = 36;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Layout {
    pub(crate) info: Range<usize>,
    pub(crate) signature: Range<usize>,
    pub(crate) certificate_store: Range<usize>,
    pub(crate) project_name: Range<usize>,
    pub(crate) timestamp_url: Range<usize>,
    pub(crate) timestamp_marker: u32,
    pub(crate) padding: Range<usize>,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct Header {
    pub(crate) signature_size: u32,
    pub(crate) signature_offset: u32,
    pub(crate) certificate_store_size: u32,
    pub(crate) certificate_store_offset: u32,
    pub(crate) project_name_size: u32,
    pub(crate) project_name_offset: u32,
    pub(crate) timestamp_marker: u32,
    pub(crate) timestamp_url_size: u32,
    pub(crate) timestamp_url_offset: u32,
}

#[derive(Debug, Clone)]
pub(crate) struct Outer {
    pub(crate) base: usize,
    pub(crate) info: Range<usize>,
    pub(crate) total_end: usize,
}

pub(crate) fn parse(source: &[u8], kind: Kind, limits: Limits) -> Result<Layout, Error> {
    validation::blob_size(source.len(), limits)?;

    let outer = match kind {
        Kind::Property => property_outer(source)?,
        Kind::Word => word_outer(source)?,
    };
    let header = header(source, outer.info.start)?;
    validation::layout(source, kind, &outer, header, limits)
}

/// Rewrites only the two opaque payload ranges while retaining every other
/// source byte. The returned candidate is reparsed before publication, so
/// size, offset, alignment, reserved-field, and resource invariants are
/// checked by the same decoder used for the original source.
pub(crate) fn rewrite(
    source: &[u8],
    kind: Kind,
    layout: &Layout,
    signature: &[u8],
    certificate_store: &[u8],
    limits: Limits,
) -> Result<Vec<u8>, Error> {
    check_payload_size(signature.len(), "signature", limits.max_signature_bytes)?;
    check_payload_size(
        certificate_store.len(),
        "certificate-store",
        limits.max_certificate_store_bytes,
    )?;
    if signature == &source[layout.signature.clone()]
        && certificate_store == &source[layout.certificate_store.clone()]
    {
        return Ok(source.to_vec());
    }

    let old_ranges = [layout.signature.clone(), layout.certificate_store.clone()];
    let replacements = [signature, certificate_store];
    reject_ambiguous_insertions(&old_ranges, &replacements)?;

    let mut order = [0usize, 1usize];
    order.sort_by_key(|index| (old_ranges[*index].start, old_ranges[*index].end, *index));

    let mut candidate = Vec::with_capacity(source.len());
    let mut cursor = 0usize;
    let mut new_ranges = [0..0, 0..0];
    for index in order {
        let old = &old_ranges[index];
        if old.start < cursor {
            if !old.is_empty() {
                return Err(Error::invalid("VBA signature payload ranges overlap"));
            }
            new_ranges[index] = candidate.len()..candidate.len();
            continue;
        }
        candidate.extend_from_slice(
            source.get(cursor..old.start).ok_or_else(|| {
                Error::invalid("VBA signature payload range is outside its source")
            })?,
        );
        let start = candidate.len();
        candidate.extend_from_slice(replacements[index]);
        new_ranges[index] = start..candidate.len();
        cursor = old.end;
    }
    candidate.extend_from_slice(
        source
            .get(cursor..)
            .ok_or_else(|| Error::invalid("VBA signature payload range is outside its source"))?,
    );

    let project_name = map_position(layout.project_name.start, &old_ranges, &replacements)?;
    let timestamp_url = map_position(layout.timestamp_url.start, &old_ranges, &replacements)?;
    let info_end = map_position(layout.info.end, &old_ranges, &replacements)?;
    let alignment = match kind {
        Kind::Property => 4,
        Kind::Word => 2,
    };
    normalize_padding(&mut candidate, info_end, alignment)?;
    let info_start = layout.info.start;
    let base = match kind {
        Kind::Property => 0,
        Kind::Word => 2,
    };
    put_u32(
        &mut candidate,
        info_start,
        u32_from_len(new_ranges[0].len(), "signature size")?,
    )?;
    put_u32(
        &mut candidate,
        info_start + 4,
        u32_from_offset(new_ranges[0].start, base, "signature offset")?,
    )?;
    put_u32(
        &mut candidate,
        info_start + 8,
        u32_from_len(new_ranges[1].len(), "certificate-store size")?,
    )?;
    put_u32(
        &mut candidate,
        info_start + 12,
        u32_from_offset(new_ranges[1].start, base, "certificate-store offset")?,
    )?;
    put_u32(
        &mut candidate,
        info_start + 20,
        u32_from_offset(project_name, base, "reserved project-name offset")?,
    )?;
    put_u32(
        &mut candidate,
        info_start + 32,
        u32_from_offset(timestamp_url, base, "reserved timestamp-URL offset")?,
    )?;

    match kind {
        Kind::Property => {
            let declared = candidate
                .len()
                .checked_sub(8)
                .ok_or_else(|| Error::invalid("DigSigBlob size underflows"))?;
            put_u32(
                &mut candidate,
                0,
                u32_from_len(declared, "DigSigBlob size")?,
            )?;
        },
        Kind::Word => {
            let info_size = info_end
                .checked_sub(10)
                .ok_or_else(|| Error::invalid("WordSigBlob information size underflows"))?;
            let payload_size = candidate
                .len()
                .checked_sub(2)
                .ok_or_else(|| Error::invalid("WordSigBlob size underflows"))?;
            if payload_size % 2 != 0 {
                return Err(Error::invalid(
                    "WordSigBlob replacement is not representable in UTF-16 code units",
                ));
            }
            let code_units = payload_size / 2;
            let code_units_u16 = u16::try_from(code_units)
                .map_err(|_error| Error::invalid("WordSigBlob character count overflows u16"))?;
            candidate[0..2].copy_from_slice(&code_units_u16.to_le_bytes());
            put_u32(
                &mut candidate,
                2,
                u32_from_len(info_size, "WordSigBlob information size")?,
            )?;
        },
    }

    parse(&candidate, kind, limits)?;
    Ok(candidate)
}

pub(crate) fn check_payload_size(
    size: usize,
    field: &'static str,
    maximum: usize,
) -> Result<(), Error> {
    if size > maximum {
        return Err(Error::Limit(field));
    }
    if u32::try_from(size).is_err() {
        return Err(Error::invalid(format!("{field} size overflows u32")));
    }
    Ok(())
}

fn normalize_padding(
    candidate: &mut Vec<u8>,
    info_end: usize,
    alignment: usize,
) -> Result<(), Error> {
    let expected = padding_for(info_end, alignment);
    let existing = candidate
        .get(info_end..)
        .ok_or_else(|| Error::invalid("VBA signature information end is outside its source"))?;
    if existing.len() > expected {
        if existing[expected..].iter().any(|byte| *byte != 0) {
            return Err(Error::invalid(
                "VBA signature edit would discard nonzero undefined padding",
            ));
        }
        candidate.truncate(info_end + expected);
    } else if existing.len() < expected {
        candidate.resize(info_end + expected, 0);
    }
    Ok(())
}

fn reject_ambiguous_insertions(
    old_ranges: &[Range<usize>; 2],
    replacements: &[&[u8]; 2],
) -> Result<(), Error> {
    for (index, old) in old_ranges.iter().enumerate() {
        if !old.is_empty() || replacements[index].is_empty() {
            continue;
        }
        if old_ranges.iter().enumerate().any(|(other_index, other)| {
            other_index != index
                && !other.is_empty()
                && old.start > other.start
                && old.start < other.end
        }) {
            return Err(Error::invalid(
                "VBA signature insertion point is inside another payload",
            ));
        }
    }
    Ok(())
}

fn map_position(
    position: usize,
    old_ranges: &[Range<usize>; 2],
    replacements: &[&[u8]; 2],
) -> Result<usize, Error> {
    let mut mapped = position;
    let mut order = [0usize, 1usize];
    order.sort_by_key(|index| (old_ranges[*index].start, old_ranges[*index].end, *index));
    for index in order {
        let old = &old_ranges[index];
        let replacement = replacements[index];
        if position < old.start {
            break;
        }
        if position < old.end {
            return Err(Error::invalid(
                "VBA signature field boundary is inside a payload",
            ));
        }
        if replacement.len() >= old.len() {
            mapped = mapped
                .checked_add(replacement.len() - old.len())
                .ok_or_else(|| Error::invalid("VBA signature offset overflows usize"))?;
        } else {
            mapped = mapped
                .checked_sub(old.len() - replacement.len())
                .ok_or_else(|| Error::invalid("VBA signature offset underflows usize"))?;
        }
    }
    Ok(mapped)
}

fn u32_from_len(value: usize, field: &'static str) -> Result<u32, Error> {
    u32::try_from(value).map_err(|_error| Error::invalid(format!("{field} overflows u32")))
}

fn u32_from_offset(offset: usize, base: usize, field: &'static str) -> Result<u32, Error> {
    let relative = offset
        .checked_sub(base)
        .ok_or_else(|| Error::invalid(format!("{field} underflows its container base")))?;
    u32_from_len(relative, field)
}

fn put_u32(bytes: &mut [u8], offset: usize, value: u32) -> Result<(), Error> {
    let target = bytes
        .get_mut(offset..offset + 4)
        .ok_or(Error::Truncated("VBA signature header"))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn property_outer(source: &[u8]) -> Result<Outer, Error> {
    let declared = usize::try_from(read_u32(source, 0, "DigSigBlob size")?)
        .map_err(|_error| Error::invalid("DigSigBlob size overflows usize"))?;
    let pointer = read_u32(source, 4, "DigSigBlob serialized pointer")?;
    if pointer != 8 {
        return Err(Error::invalid("DigSigBlob serialized pointer must equal 8"));
    }
    let total_end = 8usize
        .checked_add(declared)
        .ok_or_else(|| Error::invalid("DigSigBlob size overflows usize"))?;
    if total_end > source.len() {
        return Err(Error::Truncated("DigSigBlob payload"));
    }
    if total_end != source.len() {
        return Err(Error::invalid("DigSigBlob has trailing bytes"));
    }
    Ok(Outer {
        base: 0,
        info: 8..total_end,
        total_end,
    })
}

fn word_outer(source: &[u8]) -> Result<Outer, Error> {
    let code_units = usize::from(read_u16(source, 0, "WordSigBlob character count")?);
    let info_size = usize::try_from(read_u32(source, 2, "WordSigBlob info size")?)
        .map_err(|_error| Error::invalid("WordSigBlob info size overflows usize"))?;
    let pointer = read_u32(source, 6, "WordSigBlob serialized pointer")?;
    if pointer != 8 {
        return Err(Error::invalid(
            "WordSigBlob serialized pointer must equal 8",
        ));
    }
    let total_end = 2usize
        .checked_add(
            code_units
                .checked_mul(2)
                .ok_or_else(|| Error::invalid("WordSigBlob size overflows usize"))?,
        )
        .ok_or_else(|| Error::invalid("WordSigBlob size overflows usize"))?;
    if total_end > source.len() {
        return Err(Error::Truncated("WordSigBlob payload"));
    }
    if total_end != source.len() {
        return Err(Error::invalid("WordSigBlob has trailing bytes"));
    }
    let info_end = 10usize
        .checked_add(info_size)
        .ok_or_else(|| Error::invalid("WordSigBlob info size overflows usize"))?;
    if info_end > total_end {
        return Err(Error::Truncated("WordSigBlob signature info"));
    }
    Ok(Outer {
        base: 2,
        info: 10..info_end,
        total_end,
    })
}

fn header(source: &[u8], offset: usize) -> Result<Header, Error> {
    Ok(Header {
        signature_size: read_u32(source, offset, "signature size")?,
        signature_offset: read_u32(source, offset + 4, "signature offset")?,
        certificate_store_size: read_u32(source, offset + 8, "certificate-store size")?,
        certificate_store_offset: read_u32(source, offset + 12, "certificate-store offset")?,
        project_name_size: read_u32(source, offset + 16, "project-name size")?,
        project_name_offset: read_u32(source, offset + 20, "project-name offset")?,
        timestamp_marker: read_u32(source, offset + 24, "timestamp marker")?,
        timestamp_url_size: read_u32(source, offset + 28, "timestamp-URL size")?,
        timestamp_url_offset: read_u32(source, offset + 32, "timestamp-URL offset")?,
    })
}

fn read_u16(source: &[u8], offset: usize, field: &'static str) -> Result<u16, Error> {
    let bytes = source
        .get(offset..offset + 2)
        .ok_or(Error::Truncated(field))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(source: &[u8], offset: usize, field: &'static str) -> Result<u32, Error> {
    let bytes = source
        .get(offset..offset + 4)
        .ok_or(Error::Truncated(field))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn padding_for(offset: usize, alignment: usize) -> usize {
    (alignment - (offset % alignment)) % alignment
}
