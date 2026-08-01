//! Legacy Word smart-tag bookmarks and recognizer-state ranges.
//!
//! These structures are parsed as inert metadata. No recognizer, VBA callback,
//! download URL, or external schema is executed or contacted.

use super::fib::FileInformationBlock;
use crate::doc::package::{DocError, Result};
use litchi_ole_common::smart_tags::{PropertyBagStore, SmartTagLimits, SmartTagPropertyBag};
use std::collections::{BTreeMap, HashSet};

const STTBF_BKMK_FACTOID: usize = 114;
const PLCF_BKF_FACTOID: usize = 115;
const PLCF_BKL_FACTOID: usize = 117;
const FACTOID_DATA: usize = 118;
const PLCF_FACTOID: usize = 132;

/// Producer that originally created a smart-tag factoid.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SmartTagOrigin {
    Unknown,
    GrammarChecker,
    ExternalRecognizer,
    VisualBasic,
}

/// The `FACTOIDINFO` stored parallel to one smart-tag bookmark.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SmartTagBookmarkInfo {
    pub id: u32,
    pub is_sub_entity: bool,
    pub origin: SmartTagOrigin,
}

/// A validated Word smart-tag bookmark and its corresponding property bag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentSmartTag {
    pub start: u32,
    pub end: u32,
    pub start_depth: u16,
    pub end_depth: u16,
    pub is_native: bool,
    pub column_range: Option<(u8, u8)>,
    pub info: SmartTagBookmarkInfo,
    pub property_bag: SmartTagPropertyBag,
}

/// Recognizer processing state for one CP range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SmartTagRecognizerState {
    Pending,
    MaybeDirty,
    Dirty,
    Edit,
    Clean,
}

/// One range from `Plcffactoid`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SmartTagRecognizerRange {
    pub start: u32,
    pub end: u32,
    pub state: SmartTagRecognizerState,
}

/// Complete legacy Word smart-tag metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentSmartTags {
    /// Shared types and strings, when `FactoidData` is present.
    pub store: Option<PropertyBagStore>,
    pub tags: Vec<DocumentSmartTag>,
    pub recognizer_ranges: Vec<SmartTagRecognizerRange>,
}

impl DocumentSmartTags {
    /// Parse all smart-tag structures addressed by the FIB.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentSmartTags>> {
        let bookmark_lengths = [STTBF_BKMK_FACTOID, PLCF_BKF_FACTOID, PLCF_BKL_FACTOID]
            .map(|index| fib.get_table_pointer(index).map_or(0, |(_, length)| length));
        let factoid_data = optional_slice(fib, table_stream, FACTOID_DATA, "FactoidData")?;
        let recognizer_data = optional_slice(fib, table_stream, PLCF_FACTOID, "Plcffactoid")?;
        if bookmark_lengths.iter().all(|&length| length == 0)
            && factoid_data.is_none()
            && recognizer_data.is_none()
        {
            return Ok(None);
        }
        let has_bookmark_tables = bookmark_lengths.iter().all(|&length| length != 0);
        if !has_bookmark_tables && !bookmark_lengths.iter().all(|&length| length == 0) {
            return Err(corrupted(
                "the three parallel smart-tag bookmark tables must be present together",
            ));
        }

        let (infos, starts, ends) = if has_bookmark_tables {
            (
                parse_infos(required_slice(
                    fib,
                    table_stream,
                    STTBF_BKMK_FACTOID,
                    "SttbfBkmkFactoid",
                )?)?,
                parse_start_plcf(required_slice(
                    fib,
                    table_stream,
                    PLCF_BKF_FACTOID,
                    "PlcfBkfFactoid",
                )?)?,
                parse_end_plcf(required_slice(
                    fib,
                    table_stream,
                    PLCF_BKL_FACTOID,
                    "PlcfBklFactoid",
                )?)?,
            )
        } else {
            (Vec::new(), Vec::new(), Vec::new())
        };
        if infos.len() != starts.len() || starts.len() != ends.len() {
            return Err(corrupted(
                "smart-tag info, start, and end table counts do not match",
            ));
        }

        let document_end = fib
            .get_document_parts_end()
            .ok_or_else(|| corrupted("document-part character counts overflow"))?;
        validate_positions(&starts, document_end, "PlcfBkfFactoid")?;
        validate_positions(&ends, document_end, "PlcfBklFactoid")?;

        let codepage = ansi_codepage_for_lcid(fib.language_id());
        let limits = SmartTagLimits::default();
        let (store, bags) = if let Some(factoid_data) = factoid_data {
            let (store, consumed) = PropertyBagStore::parse_prefix(factoid_data, codepage, limits)
                .map_err(|error| corrupted(format!("invalid PropertyBagStore: {error}")))?;
            let bags = store
                .parse_bags_to_end(&factoid_data[consumed..], limits)
                .map_err(|error| {
                    corrupted(format!("invalid SmartTagData property bags: {error}"))
                })?;
            (Some(store), bags)
        } else {
            (None, Vec::new())
        };
        if bags.len() != starts.len() {
            return Err(corrupted(
                "SmartTagData property-bag count does not match the bookmark count",
            ));
        }

        let mut used_end_indexes = HashSet::with_capacity(starts.len());
        let mut tags = Vec::with_capacity(starts.len());
        for (start_index, (((start, start_data), info), property_bag)) in
            starts.iter().zip(&infos).zip(bags).enumerate()
        {
            let end_index = usize::from(start_data.end_index);
            if end_index >= ends.len() || !used_end_indexes.insert(end_index) {
                return Err(corrupted(
                    "smart-tag start end-index values must be unique and in range",
                ));
            }
            let (end, end_data) = &ends[end_index];
            if usize::from(end_data.start_index) != start_index {
                return Err(corrupted(
                    "smart-tag start and end bookmark indexes are not reciprocal",
                ));
            }
            if start > end {
                return Err(corrupted("smart-tag start CP exceeds its end CP"));
            }
            tags.push(DocumentSmartTag {
                start: *start,
                end: *end,
                start_depth: start_data.depth,
                end_depth: end_data.depth,
                is_native: start_data.is_native,
                column_range: start_data.column_range,
                info: *info,
                property_bag,
            });
        }
        validate_depths(&tags)?;

        let recognizer_ranges = recognizer_data
            .map(|data| parse_recognizer_ranges(data, document_end))
            .transpose()?
            .unwrap_or_default();
        Ok(Some(Self {
            store,
            tags,
            recognizer_ranges,
        }))
    }

    /// Resolve the type declaration associated with a tag.
    pub fn tag_type(
        &self,
        tag: &DocumentSmartTag,
    ) -> Option<&litchi_ole_common::smart_tags::SmartTagType> {
        self.store.as_ref()?.tag_type(tag.property_bag.type_id)
    }
}

#[derive(Debug)]
struct StartData {
    end_index: u16,
    depth: u16,
    is_native: bool,
    column_range: Option<(u8, u8)>,
}

#[derive(Debug)]
struct EndData {
    start_index: u16,
    depth: u16,
}

fn parse_infos(data: &[u8]) -> Result<Vec<SmartTagBookmarkInfo>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbfBkmkFactoid fExtend")? != 0xffff
        || read_u16(data, 4, "SttbfBkmkFactoid cbExtra")? != 0
    {
        return Err(corrupted("SttbfBkmkFactoid has an invalid header"));
    }
    let count = usize::from(read_u16(data, 2, "SttbfBkmkFactoid count")?);
    if count > 0x7ff0 {
        return Err(corrupted("SttbfBkmkFactoid contains too many entries"));
    }
    let expected = 6usize
        .checked_add(
            count
                .checked_mul(14)
                .ok_or_else(|| corrupted("SttbfBkmkFactoid size overflows"))?,
        )
        .ok_or_else(|| corrupted("SttbfBkmkFactoid size overflows"))?;
    if data.len() != expected {
        return Err(corrupted(
            "SttbfBkmkFactoid byte length does not match its count",
        ));
    }
    let mut infos = Vec::with_capacity(count);
    let mut ids = HashSet::with_capacity(count);
    let mut offset = 6usize;
    for _ in 0..count {
        if read_u16(data, offset, "FACTOIDINFO character count")? != 6 {
            return Err(corrupted(
                "SttbfBkmkFactoid entries must contain six UTF-16 code units",
            ));
        }
        offset += 2;
        let id = read_u32(data, offset, "FACTOIDINFO id")?;
        let flags = read_u16(data, offset + 4, "FACTOIDINFO flags")?;
        let origin = match read_u16(data, offset + 6, "FACTOIDINFO origin")? {
            0 => SmartTagOrigin::Unknown,
            1 => SmartTagOrigin::GrammarChecker,
            2 => SmartTagOrigin::ExternalRecognizer,
            3 => SmartTagOrigin::VisualBasic,
            _ => return Err(corrupted("FACTOIDINFO contains an invalid origin")),
        };
        // The final four bytes are pfpb and are explicitly ignored by consumers.
        let _ = read_u32(data, offset + 8, "FACTOIDINFO ignored pointer")?;
        if !ids.insert(id) {
            return Err(corrupted("FACTOIDINFO ids must be unique"));
        }
        infos.push(SmartTagBookmarkInfo {
            id,
            is_sub_entity: flags & 1 != 0,
            origin,
        });
        offset += 12;
    }
    Ok(infos)
}

fn parse_start_plcf(data: &[u8]) -> Result<Vec<(u32, StartData)>> {
    let count = plcf_count(data, 6, "PlcfBkfFactoid")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBkfFactoid position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 6;
        let bkc = read_u16(data, property + 2, "smart-tag BKC")?;
        if bkc & 0x0080 != 0 {
            return Err(corrupted("smart-tag BKC fPub must be zero"));
        }
        let column_range = if bkc & 0x8000 != 0 {
            let first = (bkc & 0x007f) as u8;
            let limit = ((bkc >> 8) & 0x003f) as u8;
            if first >= limit {
                return Err(corrupted("smart-tag BKC column range is empty or reversed"));
            }
            Some((first, limit))
        } else {
            None
        };
        values.push((
            read_u32(data, index * 4, "smart-tag start CP")?,
            StartData {
                end_index: read_u16(data, property, "smart-tag end index")?,
                depth: read_u16(data, property + 4, "smart-tag start depth")?,
                is_native: bkc & 0x4000 != 0,
                column_range,
            },
        ));
    }
    Ok(values)
}

fn parse_end_plcf(data: &[u8]) -> Result<Vec<(u32, EndData)>> {
    let count = plcf_count(data, 4, "PlcfBklFactoid")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfBklFactoid position bytes overflow"))?;
    let mut values = Vec::with_capacity(count);
    for index in 0..count {
        let property = properties + index * 4;
        values.push((
            read_u32(data, index * 4, "smart-tag end CP")?,
            EndData {
                start_index: read_u16(data, property, "smart-tag start index")?,
                depth: read_u16(data, property + 2, "smart-tag end depth")?,
            },
        ));
    }
    Ok(values)
}

fn parse_recognizer_ranges(data: &[u8], document_end: u32) -> Result<Vec<SmartTagRecognizerRange>> {
    let count = plcf_count(data, 2, "Plcffactoid")?;
    let properties = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("Plcffactoid position bytes overflow"))?;
    let mut ranges = Vec::with_capacity(count);
    for index in 0..count {
        let start = read_u32(data, index * 4, "Plcffactoid start CP")?;
        let end = read_u32(data, (index + 1) * 4, "Plcffactoid end CP")?;
        // As with bookmark PLCs, producers disagree on whether the final CP
        // includes inter-story paragraph marks. It is only the terminal CP;
        // all actual range starts and nonterminal ends remain bounded.
        if start > end || start > document_end || (index + 1 < count && end > document_end) {
            return Err(corrupted(format!(
                "Plcffactoid range {index} is {start}..{end}, outside document CP 0..{document_end}"
            )));
        }
        let raw = read_u16(data, properties + index * 2, "FactoidSpls")?;
        if raw & !0x000f != 0 {
            return Err(corrupted("FactoidSpls contains nonzero reserved flags"));
        }
        let state = match raw & 0x000f {
            1 => SmartTagRecognizerState::Pending,
            2 => SmartTagRecognizerState::MaybeDirty,
            3 => SmartTagRecognizerState::Dirty,
            4 => SmartTagRecognizerState::Edit,
            7 => SmartTagRecognizerState::Clean,
            _ => return Err(corrupted("FactoidSpls contains an invalid state")),
        };
        ranges.push(SmartTagRecognizerRange { start, end, state });
    }
    Ok(ranges)
}

fn plcf_count(data: &[u8], property_size: usize, name: &str) -> Result<usize> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4 + property_size) {
        return Err(corrupted(format!("{name} has an invalid byte length")));
    }
    Ok((data.len() - 4) / (4 + property_size))
}

fn validate_positions<T>(values: &[(u32, T)], document_end: u32, name: &str) -> Result<()> {
    if values.iter().any(|(cp, _)| *cp > document_end)
        || values.windows(2).any(|pair| pair[0].0 > pair[1].0)
    {
        return Err(corrupted(format!(
            "{name} contains out-of-range or non-monotonic CPs"
        )));
    }
    Ok(())
}

#[derive(Default)]
struct DepthEvents {
    nonempty_starts: usize,
    empty_ranges: usize,
    nonempty_ends: usize,
}

fn validate_depths(tags: &[DocumentSmartTag]) -> Result<()> {
    let mut events = BTreeMap::<u32, DepthEvents>::new();
    for tag in tags {
        if tag.start == tag.end {
            events.entry(tag.start).or_default().empty_ranges += 1;
        } else {
            events.entry(tag.start).or_default().nonempty_starts += 1;
            events.entry(tag.end).or_default().nonempty_ends += 1;
        }
    }

    let mut active = 0usize;
    let mut expected = BTreeMap::<u32, (u16, u16)>::new();
    for (cp, event) in events {
        active = active
            .checked_sub(event.nonempty_ends)
            .ok_or_else(|| corrupted("smart-tag bookmark depth events are inconsistent"))?;
        active = active
            .checked_add(event.nonempty_starts)
            .ok_or_else(|| corrupted("smart-tag bookmark depth overflows"))?;
        let start_depth = active
            .checked_add(event.empty_ranges)
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| corrupted("smart-tag start depth exceeds 0xFFFF"))?;
        let end_depth =
            u16::try_from(active).map_err(|_| corrupted("smart-tag end depth exceeds 0xFFFF"))?;
        expected.insert(cp, (start_depth, end_depth));
    }
    if active != 0 {
        return Err(corrupted(
            "smart-tag bookmark depth events leave unterminated ranges",
        ));
    }

    for tag in tags {
        let expected_start = expected
            .get(&tag.start)
            .ok_or_else(|| corrupted("smart-tag start depth event is missing"))?
            .0;
        let expected_end = expected
            .get(&tag.end)
            .ok_or_else(|| corrupted("smart-tag end depth event is missing"))?
            .1;
        if tag.start_depth != expected_start || tag.end_depth != expected_end {
            return Err(corrupted(format!(
                "smart-tag bookmark {}..{} has depths {}/{}, expected {}/{}",
                tag.start, tag.end, tag.start_depth, tag.end_depth, expected_start, expected_end
            )));
        }
    }
    Ok(())
}

fn required_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<&'a [u8]> {
    optional_slice(fib, table_stream, index, name)?
        .ok_or_else(|| corrupted(format!("{name} is missing")))
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, name: &str) -> Result<u16> {
    let bytes = data
        .get(offset..offset.saturating_add(2))
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, name: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset.saturating_add(4))
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

fn ansi_codepage_for_lcid(lcid: u16) -> u32 {
    match lcid & 0x03ff {
        0x01 | 0x20 | 0x29 => 1256,
        0x02 | 0x19 | 0x22 | 0x23 | 0x28 | 0x2f | 0x3f | 0x40 | 0x44 | 0x50 => 1251,
        0x04 => match lcid {
            0x0804 | 0x1004 => 936,
            _ => 950,
        },
        0x05 | 0x0e | 0x15 | 0x18 | 0x1b | 0x1c | 0x24 | 0x2e => 1250,
        0x08 => 1253,
        0x0d => 1255,
        0x11 => 932,
        0x12 => 949,
        0x1a => match lcid {
            0x0c1a | 0x201a => 1251,
            _ => 1250,
        },
        0x1e => 874,
        0x1f => 1254,
        0x25..=0x27 => 1257,
        0x2a => 1258,
        0x2c => match lcid {
            0x082c => 1251,
            _ => 1254,
        },
        0x43 => match lcid {
            0x0843 => 1251,
            _ => 1254,
        },
        _ => 1252,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::Package;
    use std::path::PathBuf;

    #[test]
    fn maps_representative_lcids_to_ansi_codepages() {
        assert_eq!(ansi_codepage_for_lcid(0x0409), 1252);
        assert_eq!(ansi_codepage_for_lcid(0x0411), 932);
        assert_eq!(ansi_codepage_for_lcid(0x0804), 936);
        assert_eq!(ansi_codepage_for_lcid(0x0419), 1251);
    }

    #[test]
    fn rejects_invalid_recognizer_states_and_reserved_bits() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        assert!(parse_recognizer_ranges(&data, 10).is_err());
        data[8..10].copy_from_slice(&0x0011u16.to_le_bytes());
        assert!(parse_recognizer_ranges(&data, 10).is_err());
        data[8..10].copy_from_slice(&7u16.to_le_bytes());
        assert_eq!(
            parse_recognizer_ranges(&data, 10).unwrap()[0].state,
            SmartTagRecognizerState::Clean
        );
    }

    #[test]
    fn accepts_independent_recognizer_ranges() {
        let mut fib_bytes = vec![0u8; 154 + 136 * 8];
        fib_bytes[..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        fib_bytes[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        fib_bytes[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
        fib_bytes[76..80].copy_from_slice(&10u32.to_le_bytes());
        fib_bytes[152..154].copy_from_slice(&136u16.to_le_bytes());

        let mut table = Vec::new();
        table.extend_from_slice(&2u32.to_le_bytes());
        table.extend_from_slice(&7u32.to_le_bytes());
        table.extend_from_slice(&7u16.to_le_bytes());
        let pointer = 154 + PLCF_FACTOID * 8;
        fib_bytes[pointer..pointer + 4].copy_from_slice(&0u32.to_le_bytes());
        fib_bytes[pointer + 4..pointer + 8].copy_from_slice(&(table.len() as u32).to_le_bytes());

        let fib = FileInformationBlock::parse(&fib_bytes).unwrap();
        let parsed = DocumentSmartTags::parse(&fib, &table).unwrap().unwrap();
        assert!(parsed.store.is_none());
        assert!(parsed.tags.is_empty());
        assert_eq!(
            parsed.recognizer_ranges,
            vec![SmartTagRecognizerRange {
                start: 2,
                end: 7,
                state: SmartTagRecognizerState::Clean,
            }]
        );
    }

    fn depth_tag(start: u32, end: u32, start_depth: u16, end_depth: u16) -> DocumentSmartTag {
        DocumentSmartTag {
            start,
            end,
            start_depth,
            end_depth,
            is_native: false,
            column_range: None,
            info: SmartTagBookmarkInfo {
                id: start,
                is_sub_entity: false,
                origin: SmartTagOrigin::Unknown,
            },
            property_bag: SmartTagPropertyBag {
                type_id: 0,
                properties: Vec::new(),
            },
        }
    }

    #[test]
    fn validates_overlapping_and_empty_bookmark_depths() {
        let tags = [
            depth_tag(0, 10, 1, 1),
            depth_tag(5, 15, 3, 0),
            depth_tag(5, 5, 3, 2),
        ];
        validate_depths(&tags).unwrap();
        let mut bad = tags;
        bad[1].start_depth = 1;
        assert!(validate_depths(&bad).is_err());
    }

    fn fixture(path: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(path)
    }

    #[test]
    fn parses_real_word_smart_tag_fixtures() {
        let cases = [
            "test-data/libreoffice-core/sw/qa/core/doc/data/bookmark-delete-redline.doc",
            "test-data/ole/doc/FloatingPictures.doc",
        ];
        for path in cases {
            let mut package = Package::open(fixture(path)).unwrap();
            let document = package.document().unwrap();
            let smart_tags = document
                .smart_tags()
                .unwrap_or_else(|| panic!("{path} should contain smart tags"));
            assert!(!smart_tags.tags.is_empty(), "{path}");
            let store = smart_tags
                .store
                .as_ref()
                .unwrap_or_else(|| panic!("{path} should contain FactoidData"));
            assert!(!store.types.is_empty(), "{path}");
            for tag in &smart_tags.tags {
                assert!(tag.start <= tag.end, "{path}");
                assert!(smart_tags.tag_type(tag).is_some(), "{path}");
                for property in &tag.property_bag.properties {
                    assert!(store.resolve_property(*property).is_some(), "{path}");
                }
            }
        }
    }
}
