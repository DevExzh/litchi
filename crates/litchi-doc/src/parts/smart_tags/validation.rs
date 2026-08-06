//! Bounded semantic and topology validation for legacy smart-tag edits.

use super::semantic::{TableKind, Topology};
use super::{DocumentSmartTags, SmartTagRecognizerRange};
use crate::package::{Error as PackageError, Result};
use litchi_ole_common::smart_tags::{Limits, PropertyBagStore};
use std::collections::BTreeMap;
use std::collections::HashSet;

const MAX_BOOKMARKS: usize = 0x7ff0;

/// Validate a parsed source against its retained FIB/table topology.
pub(super) fn source(
    topology: &Topology,
    metadata: &DocumentSmartTags,
    table_stream: &[u8],
    links: &[u16],
    limits: Limits,
) -> Result<()> {
    validate_presence(topology, metadata)?;
    validate_metadata(metadata, topology.document_end(), links, limits)?;
    validate_ranges(topology, metadata, table_stream, limits)
}

/// Validate a candidate before it can replace the immutable source snapshot.
pub(super) fn candidate(
    topology: &Topology,
    source: &DocumentSmartTags,
    metadata: &DocumentSmartTags,
    links: &[u16],
    limits: Limits,
) -> Result<()> {
    validate_presence(topology, metadata)?;
    if metadata.tags.len() != source.tags.len() {
        return Err(corrupted(
            "smart-tag bookmark count changes would alter FIB/table topology",
        ));
    }
    if metadata.recognizer_ranges.len() != source.recognizer_ranges.len() {
        return Err(corrupted(
            "smart-tag recognizer range count changes would alter FIB/table topology",
        ));
    }
    validate_metadata(metadata, topology.document_end(), links, limits)
}

fn validate_presence(topology: &Topology, metadata: &DocumentSmartTags) -> Result<()> {
    let bookmark_ranges = [
        topology.range(TableKind::BookmarkInfo),
        topology.range(TableKind::BookmarkStarts),
        topology.range(TableKind::BookmarkEnds),
    ];
    let bookmark_present = bookmark_ranges.map(|range| range.is_some());
    if bookmark_present.iter().any(|present| *present)
        && bookmark_present.iter().any(|present| !present)
    {
        return Err(corrupted(
            "the three parallel smart-tag bookmark tables must be present together",
        ));
    }
    if !bookmark_present[0] && !metadata.tags.is_empty() {
        return Err(corrupted(
            "smart-tag bookmarks are present without their FIB tables",
        ));
    }
    if topology.range(TableKind::PropertyBags).is_some() != metadata.store.is_some() {
        return Err(corrupted(
            "FactoidData presence does not match the parsed PropertyBagStore",
        ));
    }
    Ok(())
}

fn validate_metadata(
    metadata: &DocumentSmartTags,
    document_end: u32,
    links: &[u16],
    limits: Limits,
) -> Result<()> {
    if metadata.tags.len() > MAX_BOOKMARKS || metadata.tags.len() > limits.max_bags {
        return Err(corrupted(
            "smart-tag bookmark count exceeds the configured limit",
        ));
    }
    if links.len() != metadata.tags.len() {
        return Err(corrupted(
            "smart-tag bookmark link count does not match the metadata",
        ));
    }

    let mut ids = HashSet::with_capacity(metadata.tags.len());
    for tag in &metadata.tags {
        if tag.start > tag.end || tag.end > document_end {
            return Err(corrupted(
                "smart-tag bookmark range is outside the document parts",
            ));
        }
        if !ids.insert(tag.info.id) {
            return Err(corrupted("FACTOIDINFO ids must be unique"));
        }
        if let Some((first, limit)) = tag.column_range
            && (first >= limit || first > 0x7f || limit > 0x3f)
        {
            return Err(corrupted("smart-tag BKC column range is invalid"));
        }
    }
    if metadata
        .tags
        .windows(2)
        .any(|pair| pair[0].start > pair[1].start)
    {
        return Err(corrupted("smart-tag bookmark start CPs are not monotonic"));
    }
    let mut seen_links = HashSet::with_capacity(links.len());
    let mut ends = vec![0u32; links.len()];
    for (start_index, &end_index) in links.iter().enumerate() {
        let end_index = usize::from(end_index);
        if end_index >= ends.len() || !seen_links.insert(end_index) {
            return Err(corrupted(
                "smart-tag bookmark end indexes must be unique and in range",
            ));
        }
        ends[end_index] = metadata.tags[start_index].end;
    }
    if ends.windows(2).any(|pair| pair[0] > pair[1]) {
        return Err(corrupted("smart-tag bookmark end CPs are not monotonic"));
    }
    super::validate_depths(&metadata.tags)?;

    if metadata.recognizer_ranges.len() > MAX_BOOKMARKS {
        return Err(corrupted(
            "smart-tag recognizer range count exceeds the encoded limit",
        ));
    }
    validate_recognizer_ranges(&metadata.recognizer_ranges, document_end)?;

    if let Some(store) = &metadata.store {
        let bags = metadata
            .tags
            .iter()
            .map(|tag| tag.property_bag.clone())
            .collect::<Vec<_>>();
        if bags.len() > limits.max_bags || store.types.len() > limits.max_types {
            return Err(corrupted(
                "smart-tag PropertyBagStore exceeds the configured limit",
            ));
        }
        if store.strings.len() > limits.max_strings {
            return Err(corrupted(
                "smart-tag string table exceeds the configured limit",
            ));
        }
        let property_count = bags.iter().try_fold(0usize, |count, bag| {
            count
                .checked_add(bag.properties.len())
                .ok_or_else(|| corrupted("smart-tag property count overflows"))
        })?;
        if property_count > limits.max_properties {
            return Err(corrupted(
                "smart-tag property count exceeds the configured limit",
            ));
        }
        let encoded = store
            .to_bytes_with_bags(&bags)
            .map_err(|error| corrupted(format!("invalid PropertyBagStore: {error}")))?;
        if encoded.len() > limits.max_bytes {
            return Err(corrupted(
                "smart-tag property-bag payload exceeds the configured limit",
            ));
        }
    }
    Ok(())
}

fn validate_ranges(
    topology: &Topology,
    metadata: &DocumentSmartTags,
    table_stream: &[u8],
    limits: Limits,
) -> Result<()> {
    let count = metadata.tags.len();
    check_length(
        topology,
        table_stream,
        TableKind::BookmarkInfo,
        6usize
            .checked_add(
                count
                    .checked_mul(14)
                    .ok_or_else(|| corrupted("SttbfBkmkFactoid size overflows"))?,
            )
            .ok_or_else(|| corrupted("SttbfBkmkFactoid size overflows"))?,
    )?;
    check_length(
        topology,
        table_stream,
        TableKind::BookmarkStarts,
        plc_length(count, 6, "PlcfBkfFactoid")?,
    )?;
    check_length(
        topology,
        table_stream,
        TableKind::BookmarkEnds,
        plc_length(count, 4, "PlcfBklFactoid")?,
    )?;
    let recognizer_count = metadata.recognizer_ranges.len();
    check_length(
        topology,
        table_stream,
        TableKind::Recognizer,
        plc_length(recognizer_count, 2, "Plcffactoid")?,
    )?;

    if let Some(data) = topology.range_bytes(table_stream, TableKind::PropertyBags)? {
        let store = metadata
            .store
            .as_ref()
            .ok_or_else(|| corrupted("FactoidData has no PropertyBagStore"))?;
        if data.len() > limits.max_bytes {
            return Err(corrupted("FactoidData exceeds the configured size limit"));
        }
        let (parsed_store, consumed) = PropertyBagStore::parse_prefix(data, store.ansi, limits)
            .map_err(|error| corrupted(format!("invalid PropertyBagStore: {error}")))?;
        let parsed_bags = parsed_store
            .parse_bags_to_end(&data[consumed..], limits)
            .map_err(|error| corrupted(format!("invalid SmartTagData property bags: {error}")))?;
        let expected_bags = metadata
            .tags
            .iter()
            .map(|tag| tag.property_bag.clone())
            .collect::<Vec<_>>();
        if parsed_store != *store || parsed_bags != expected_bags {
            return Err(corrupted(
                "FactoidData does not match the typed smart-tag projection",
            ));
        }
    }
    Ok(())
}

fn validate_recognizer_ranges(ranges: &[SmartTagRecognizerRange], document_end: u32) -> Result<()> {
    for range in ranges {
        if range.start > range.end || range.start > document_end || range.end > document_end {
            return Err(corrupted("Plcffactoid range is outside the document parts"));
        }
    }
    Ok(())
}

pub(super) fn recompute_depths(tags: &mut [super::DocumentSmartTag]) -> Result<()> {
    #[derive(Default)]
    struct Events {
        starts: usize,
        empty: usize,
        ends: usize,
    }

    let mut events = BTreeMap::<u32, Events>::new();
    for tag in tags.iter() {
        if tag.start == tag.end {
            events.entry(tag.start).or_default().empty += 1;
        } else {
            events.entry(tag.start).or_default().starts += 1;
            events.entry(tag.end).or_default().ends += 1;
        }
    }
    let mut active = 0usize;
    let mut depths = BTreeMap::<u32, (u16, u16)>::new();
    for (cp, event) in events {
        active = active
            .checked_sub(event.ends)
            .ok_or_else(|| corrupted("smart-tag bookmark depth events underflow"))?;
        active = active
            .checked_add(event.starts)
            .ok_or_else(|| corrupted("smart-tag bookmark depth events overflow"))?;
        let start = active
            .checked_add(event.empty)
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| corrupted("smart-tag start depth exceeds 0xFFFF"))?;
        let end =
            u16::try_from(active).map_err(|_| corrupted("smart-tag end depth exceeds 0xFFFF"))?;
        depths.insert(cp, (start, end));
    }
    if active != 0 {
        return Err(corrupted(
            "smart-tag bookmark depth events leave unterminated ranges",
        ));
    }
    for tag in tags {
        let start = depths
            .get(&tag.start)
            .ok_or_else(|| corrupted("smart-tag start depth event is missing"))?;
        let end = depths
            .get(&tag.end)
            .ok_or_else(|| corrupted("smart-tag end depth event is missing"))?;
        tag.start_depth = start.0;
        tag.end_depth = end.1;
    }
    Ok(())
}

fn check_length(
    topology: &Topology,
    table_stream: &[u8],
    kind: TableKind,
    expected: usize,
) -> Result<()> {
    let Some(range) = topology.range(kind) else {
        if expected != 0 {
            return Err(corrupted(format!(
                "{} is missing its FIB range",
                kind_name(kind)
            )));
        }
        return Ok(());
    };
    let bytes = range.as_usize(table_stream.len())?;
    if bytes.len() != expected {
        return Err(corrupted(format!(
            "{} byte length does not match its typed count",
            kind_name(kind)
        )));
    }
    Ok(())
}

fn plc_length(count: usize, property_size: usize, name: &str) -> Result<usize> {
    count
        .checked_add(1)
        .and_then(|value| value.checked_mul(4))
        .and_then(|positions| {
            count
                .checked_mul(property_size)
                .and_then(|properties| positions.checked_add(properties))
        })
        .ok_or_else(|| corrupted(format!("{name} size overflows")))
}

fn kind_name(kind: TableKind) -> &'static str {
    match kind {
        TableKind::BookmarkInfo => "SttbfBkmkFactoid",
        TableKind::BookmarkStarts => "PlcfBkfFactoid",
        TableKind::BookmarkEnds => "PlcfBklFactoid",
        TableKind::PropertyBags => "FactoidData",
        TableKind::Recognizer => "Plcffactoid",
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
