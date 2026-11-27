//! Fixed-range binary codecs for the five MS-DOC smart-tag structures.

use super::semantic::{TableKind, Topology};
use super::validation;
use super::{DocumentSmartTags, SmartTagOrigin, SmartTagRecognizerState};
use crate::package::{Error as PackageError, Result};
use litchi_ole_common::smart_tags::Limits;

/// Read the stable start-to-end bookmark links from `PlcfBkfFactoid`.
pub(super) fn bookmark_links(topology: &Topology, table_stream: &[u8]) -> Result<Vec<u16>> {
    let Some(data) = topology.range_bytes(table_stream, TableKind::BookmarkStarts)? else {
        return Ok(Vec::new());
    };
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(10) {
        return Err(corrupted("PlcfBkfFactoid has an invalid byte length"));
    }
    let count = (data.len() - 4) / 10;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfFactoid position bytes overflow"))?;
    (0..count)
        .map(|index| super::read_u16(data, properties + index * 6, "smart-tag end index"))
        .collect()
}

/// Encode a candidate over the exact source table stream.
///
/// Every replacement is length-checked against its original FIB range. The
/// function only writes recognized fields, leaving all other table bytes and
/// ignored fields in place.
pub(super) fn encode(
    source_table: &[u8],
    topology: &Topology,
    source: &DocumentSmartTags,
    candidate: &DocumentSmartTags,
    links: &[u16],
    limits: Limits,
) -> Result<Vec<u8>> {
    validation::candidate(topology, source, candidate, links, limits)?;
    if source == candidate {
        return Ok(source_table.to_vec());
    }
    let mut output = source_table.to_vec();
    write_infos(&mut output, topology, candidate)?;
    write_starts(&mut output, topology, source, candidate, links)?;
    write_ends(&mut output, topology, candidate, links)?;
    write_factoid_data(&mut output, topology, source, candidate, limits)?;
    write_recognizer(&mut output, topology, candidate)?;
    Ok(output)
}

fn write_infos(
    output: &mut [u8],
    topology: &Topology,
    candidate: &DocumentSmartTags,
) -> Result<()> {
    let Some(range) = topology.range(TableKind::BookmarkInfo) else {
        return Ok(());
    };
    let range = range.as_usize(output.len())?;
    let bytes = &mut output[range];
    for (index, tag) in candidate.tags.iter().enumerate() {
        let entry = 6usize
            .checked_add(
                index
                    .checked_mul(14)
                    .ok_or_else(|| corrupted("FACTOIDINFO offset overflows"))?,
            )
            .ok_or_else(|| corrupted("FACTOIDINFO offset overflows"))?;
        let id_end = entry
            .checked_add(6)
            .ok_or_else(|| corrupted("FACTOIDINFO offset overflows"))?;
        let flags_end = entry
            .checked_add(8)
            .ok_or_else(|| corrupted("FACTOIDINFO offset overflows"))?;
        let origin_end = entry
            .checked_add(10)
            .ok_or_else(|| corrupted("FACTOIDINFO offset overflows"))?;
        if origin_end > bytes.len() {
            return Err(corrupted("SttbfBkmkFactoid is truncated"));
        }
        bytes[entry + 2..id_end].copy_from_slice(&tag.info.id.to_le_bytes());
        let old_flags = super::read_u16(bytes, entry + 6, "FACTOIDINFO flags")?;
        let flags = (old_flags & !1) | u16::from(tag.info.is_sub_entity);
        bytes[entry + 6..flags_end].copy_from_slice(&flags.to_le_bytes());
        bytes[entry + 8..origin_end].copy_from_slice(&origin_code(tag.info.origin).to_le_bytes());
    }
    Ok(())
}

fn write_starts(
    output: &mut [u8],
    topology: &Topology,
    source: &DocumentSmartTags,
    candidate: &DocumentSmartTags,
    links: &[u16],
) -> Result<()> {
    let Some(range) = topology.range(TableKind::BookmarkStarts) else {
        return Ok(());
    };
    let range = range.as_usize(output.len())?;
    let bytes = &mut output[range];
    let count = candidate.tags.len();
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfFactoid position bytes overflow"))?;
    for (index, tag) in candidate.tags.iter().enumerate() {
        let cp_end = index
            .checked_mul(4)
            .and_then(|offset| offset.checked_add(4))
            .ok_or_else(|| corrupted("PlcfBkfFactoid CP offset overflows"))?;
        bytes[index * 4..cp_end].copy_from_slice(&tag.start.to_le_bytes());

        let property = properties
            .checked_add(
                index
                    .checked_mul(6)
                    .ok_or_else(|| corrupted("PlcfBkfFactoid property offset overflows"))?,
            )
            .ok_or_else(|| corrupted("PlcfBkfFactoid property offset overflows"))?;
        let property_end = property
            .checked_add(6)
            .ok_or_else(|| corrupted("PlcfBkfFactoid property offset overflows"))?;
        if property_end > bytes.len() {
            return Err(corrupted("PlcfBkfFactoid is truncated"));
        }
        bytes[property..property + 2].copy_from_slice(&links[index].to_le_bytes());
        let old_bkc = super::read_u16(bytes, property + 2, "smart-tag BKC")?;
        let bkc = bkc(old_bkc, tag.is_native, tag.column_range);
        bytes[property + 2..property + 4].copy_from_slice(&bkc.to_le_bytes());
        bytes[property + 4..property_end].copy_from_slice(&tag.start_depth.to_le_bytes());
    }
    if source.tags.len() != count {
        return Err(corrupted(
            "smart-tag bookmark count changed during encoding",
        ));
    }
    Ok(())
}

fn write_ends(
    output: &mut [u8],
    topology: &Topology,
    candidate: &DocumentSmartTags,
    links: &[u16],
) -> Result<()> {
    let Some(range) = topology.range(TableKind::BookmarkEnds) else {
        return Ok(());
    };
    let range = range.as_usize(output.len())?;
    let bytes = &mut output[range];
    let count = candidate.tags.len();
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBklFactoid position bytes overflow"))?;
    for (start_index, tag) in candidate.tags.iter().enumerate() {
        let end_index = usize::from(links[start_index]);
        let cp_end = end_index
            .checked_mul(4)
            .and_then(|offset| offset.checked_add(4))
            .ok_or_else(|| corrupted("PlcfBklFactoid CP offset overflows"))?;
        let property = properties
            .checked_add(
                end_index
                    .checked_mul(4)
                    .ok_or_else(|| corrupted("PlcfBklFactoid property offset overflows"))?,
            )
            .ok_or_else(|| corrupted("PlcfBklFactoid property offset overflows"))?;
        let property_end = property
            .checked_add(4)
            .ok_or_else(|| corrupted("PlcfBklFactoid property offset overflows"))?;
        if property_end > bytes.len() {
            return Err(corrupted("PlcfBklFactoid is truncated"));
        }
        bytes[end_index * 4..cp_end].copy_from_slice(&tag.end.to_le_bytes());
        bytes[property..property + 2].copy_from_slice(&(start_index as u16).to_le_bytes());
        bytes[property + 2..property_end].copy_from_slice(&tag.end_depth.to_le_bytes());
    }
    Ok(())
}

fn write_factoid_data(
    output: &mut [u8],
    topology: &Topology,
    source: &DocumentSmartTags,
    candidate: &DocumentSmartTags,
    limits: Limits,
) -> Result<()> {
    let Some(range) = topology.range(TableKind::PropertyBags) else {
        return Ok(());
    };
    let range = range.as_usize(output.len())?;
    if source.store == candidate.store
        && source
            .tags
            .iter()
            .map(|tag| &tag.property_bag)
            .eq(candidate.tags.iter().map(|tag| &tag.property_bag))
    {
        return Ok(());
    }
    let store = candidate
        .store
        .as_ref()
        .ok_or_else(|| corrupted("FactoidData has no PropertyBagStore"))?;
    let bags = candidate
        .tags
        .iter()
        .map(|tag| tag.property_bag.clone())
        .collect::<Vec<_>>();
    let encoded = store
        .to_bytes_with_bags(&bags)
        .map_err(|error| corrupted(format!("invalid PropertyBagStore: {error}")))?;
    if encoded.len() > limits.max_bytes {
        return Err(corrupted("FactoidData exceeds the configured size limit"));
    }
    if encoded.len() != range.len() {
        return Err(corrupted(
            "property-bag edits would change the FIB FactoidData range length",
        ));
    }
    output[range].copy_from_slice(&encoded);
    Ok(())
}

fn write_recognizer(
    output: &mut [u8],
    topology: &Topology,
    candidate: &DocumentSmartTags,
) -> Result<()> {
    let Some(range) = topology.range(TableKind::Recognizer) else {
        return Ok(());
    };
    let range = range.as_usize(output.len())?;
    let bytes = &mut output[range];
    let count = candidate.recognizer_ranges.len();
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("Plcffactoid position bytes overflow"))?;
    for (index, value) in candidate.recognizer_ranges.iter().enumerate() {
        let start = index * 4;
        bytes[start..start + 4].copy_from_slice(&value.start.to_le_bytes());
        bytes[(index + 1) * 4..(index + 2) * 4].copy_from_slice(&value.end.to_le_bytes());
        let code = state_code(value.state);
        let old = super::read_u16(bytes, properties + index * 2, "FactoidSpls")?;
        bytes[properties + index * 2..properties + index * 2 + 2]
            .copy_from_slice(&((old & !0x000f) | code).to_le_bytes());
    }
    Ok(())
}

fn bkc(old: u16, is_native: bool, column_range: Option<(u8, u8)>) -> u16 {
    const NATIVE: u16 = 0x4000;
    const COLUMN: u16 = 0x8000;
    const FIRST: u16 = 0x007f;
    const LIMIT: u16 = 0x3f00;
    let mut value = old & !NATIVE;
    if is_native {
        value |= NATIVE;
    }
    if let Some((first, limit)) = column_range {
        value = (value & !(COLUMN | FIRST | LIMIT))
            | COLUMN
            | u16::from(first)
            | (u16::from(limit) << 8);
    } else {
        value &= !COLUMN;
    }
    value
}

fn origin_code(origin: SmartTagOrigin) -> u16 {
    match origin {
        SmartTagOrigin::Unknown => 0,
        SmartTagOrigin::GrammarChecker => 1,
        SmartTagOrigin::ExternalRecognizer => 2,
        SmartTagOrigin::VisualBasic => 3,
    }
}

fn state_code(state: SmartTagRecognizerState) -> u16 {
    match state {
        SmartTagRecognizerState::Pending => 1,
        SmartTagRecognizerState::MaybeDirty => 2,
        SmartTagRecognizerState::Dirty => 3,
        SmartTagRecognizerState::Edit => 4,
        SmartTagRecognizerState::Clean => 7,
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
