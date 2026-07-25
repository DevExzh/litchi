//! Typed authoring for legacy Word smart tags and recognizer-state ranges.

use super::core::DocWriteError;
use crate::doc::{SmartTagOrigin, SmartTagRecognizerRange, SmartTagRecognizerState};
use crate::smart_tags::{
    PropertyBagStore, PropertyBagString, PropertyBagStringEncoding, SmartTagProperty,
    SmartTagPropertyBag, SmartTagType,
};
use std::collections::{BTreeMap, HashMap};

const MAX_SMART_TAGS: usize = 0x7ff0;

/// One smart-tag bookmark to embed in a generated DOC file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocSmartTagEntry {
    pub start: u32,
    pub end: u32,
    pub namespace_uri: String,
    pub tag_name: String,
    pub download_url: String,
    pub origin: SmartTagOrigin,
    pub is_sub_entity: bool,
    pub is_native: bool,
    pub column_range: Option<(u8, u8)>,
    pub properties: Vec<(String, String)>,
}

impl DocSmartTagEntry {
    /// Create a smart tag with an empty property bag.
    pub fn new(
        start: u32,
        end: u32,
        namespace_uri: impl Into<String>,
        tag_name: impl Into<String>,
    ) -> Self {
        Self {
            start,
            end,
            namespace_uri: namespace_uri.into(),
            tag_name: tag_name.into(),
            download_url: String::new(),
            origin: SmartTagOrigin::Unknown,
            is_sub_entity: false,
            is_native: false,
            column_range: None,
            properties: Vec::new(),
        }
    }

    pub fn with_download_url(mut self, value: impl Into<String>) -> Self {
        self.download_url = value.into();
        self
    }

    pub fn with_origin(mut self, value: SmartTagOrigin) -> Self {
        self.origin = value;
        self
    }

    pub fn with_sub_entity(mut self, value: bool) -> Self {
        self.is_sub_entity = value;
        self
    }

    pub fn with_native_export(mut self, value: bool) -> Self {
        self.is_native = value;
        self
    }

    pub fn with_column_range(mut self, first: u8, limit: u8) -> Self {
        self.column_range = Some((first, limit));
        self
    }

    pub fn with_property(mut self, key: impl Into<String>, value: impl Into<String>) -> Self {
        self.properties.push((key.into(), value.into()));
        self
    }

    pub fn add_property(&mut self, key: impl Into<String>, value: impl Into<String>) {
        self.properties.push((key.into(), value.into()));
    }
}

#[derive(Debug)]
pub(crate) struct SmartTagTableData {
    pub(crate) infos: Option<Vec<u8>>,
    pub(crate) starts: Option<Vec<u8>>,
    pub(crate) ends: Option<Vec<u8>>,
    pub(crate) factoid_data: Option<Vec<u8>>,
    pub(crate) recognizer_ranges: Option<Vec<u8>>,
}

impl SmartTagTableData {
    pub(crate) fn is_empty(&self) -> bool {
        self.infos.is_none() && self.recognizer_ranges.is_none()
    }
}

#[derive(Default)]
struct DepthEvents {
    starts: usize,
    empty: usize,
    ends: usize,
}

pub(crate) fn build_tables(
    entries: &[DocSmartTagEntry],
    recognizer_ranges: &[SmartTagRecognizerRange],
    document_end: u32,
) -> Result<SmartTagTableData, DocWriteError> {
    if entries.len() > MAX_SMART_TAGS {
        return invalid("DOC smart-tag table exceeds 0x7FF0 entries");
    }
    let mut events = BTreeMap::<u32, DepthEvents>::new();
    for entry in entries {
        validate_entry(entry, document_end)?;
        if entry.start == entry.end {
            events.entry(entry.start).or_default().empty += 1;
        } else {
            events.entry(entry.start).or_default().starts += 1;
            events.entry(entry.end).or_default().ends += 1;
        }
    }
    let depths = calculate_depths(events)?;

    let mut start_order = (0..entries.len()).collect::<Vec<_>>();
    start_order.sort_by_key(|&index| (entries[index].start, index));
    let mut end_order = (0..entries.len()).collect::<Vec<_>>();
    end_order.sort_by_key(|&index| (entries[index].end, index));
    let start_indexes = start_order
        .iter()
        .enumerate()
        .map(|(position, index)| (*index, position as u16))
        .collect::<HashMap<_, _>>();
    let end_indexes = end_order
        .iter()
        .enumerate()
        .map(|(position, index)| (*index, position as u16))
        .collect::<HashMap<_, _>>();

    let (infos, starts, ends, factoid_data) = if entries.is_empty() {
        (None, None, None, None)
    } else {
        let infos = build_infos(entries, &start_order)?;
        let starts = build_starts(entries, &start_order, &end_indexes, &depths, document_end)?;
        let ends = build_ends(entries, &end_order, &start_indexes, &depths, document_end)?;
        let factoid_data = build_factoid_data(entries, &start_order)?;
        (Some(infos), Some(starts), Some(ends), Some(factoid_data))
    };
    let recognizer_ranges = build_recognizer_ranges(recognizer_ranges, document_end)?;
    Ok(SmartTagTableData {
        infos,
        starts,
        ends,
        factoid_data,
        recognizer_ranges,
    })
}

fn validate_entry(entry: &DocSmartTagEntry, document_end: u32) -> Result<(), DocWriteError> {
    if entry.start > entry.end || entry.end > document_end {
        return invalid("DOC smart-tag range must be ordered and inside the document parts");
    }
    if entry.namespace_uri.contains('\0')
        || entry.tag_name.contains('\0')
        || entry.download_url.contains('\0')
        || entry
            .properties
            .iter()
            .any(|(key, value)| key.contains('\0') || value.contains('\0'))
    {
        return invalid("DOC smart-tag strings cannot contain NUL characters");
    }
    if entry.properties.len() > usize::from(u16::MAX) {
        return invalid("DOC smart-tag property bag exceeds 65535 properties");
    }
    if let Some((first, limit)) = entry.column_range
        && (first >= limit || first > 0x7f || limit > 0x3f)
    {
        return invalid("DOC smart-tag column range exceeds BKC limits");
    }
    Ok(())
}

fn calculate_depths(
    events: BTreeMap<u32, DepthEvents>,
) -> Result<BTreeMap<u32, (u16, u16)>, DocWriteError> {
    let mut active = 0usize;
    let mut depths = BTreeMap::new();
    for (cp, event) in events {
        active = active
            .checked_sub(event.ends)
            .ok_or_else(|| DocWriteError::InvalidData("smart-tag depth underflow".to_string()))?;
        active = active
            .checked_add(event.starts)
            .ok_or_else(|| DocWriteError::InvalidData("smart-tag depth overflow".to_string()))?;
        let start = active
            .checked_add(event.empty)
            .and_then(|value| u16::try_from(value).ok())
            .ok_or_else(|| {
                DocWriteError::InvalidData("smart-tag start depth exceeds 0xFFFF".to_string())
            })?;
        let end = u16::try_from(active).map_err(|_| {
            DocWriteError::InvalidData("smart-tag end depth exceeds 0xFFFF".to_string())
        })?;
        depths.insert(cp, (start, end));
    }
    if active != 0 {
        return invalid("DOC smart-tag ranges contain unterminated depth events");
    }
    Ok(depths)
}

fn build_infos(
    entries: &[DocSmartTagEntry],
    start_order: &[usize],
) -> Result<Vec<u8>, DocWriteError> {
    let count = u16::try_from(start_order.len())
        .map_err(|_| DocWriteError::InvalidData("smart-tag count exceeds u16".to_string()))?;
    let mut output = Vec::with_capacity(6 + start_order.len() * 14);
    output.extend_from_slice(&0xffffu16.to_le_bytes());
    output.extend_from_slice(&count.to_le_bytes());
    output.extend_from_slice(&0u16.to_le_bytes());
    for (id, &index) in start_order.iter().enumerate() {
        output.extend_from_slice(&6u16.to_le_bytes());
        output.extend_from_slice(&(id as u32).to_le_bytes());
        output.extend_from_slice(&u16::from(entries[index].is_sub_entity).to_le_bytes());
        output.extend_from_slice(&origin_code(entries[index].origin).to_le_bytes());
        output.extend_from_slice(&0u32.to_le_bytes());
    }
    Ok(output)
}

fn build_starts(
    entries: &[DocSmartTagEntry],
    start_order: &[usize],
    end_indexes: &HashMap<usize, u16>,
    depths: &BTreeMap<u32, (u16, u16)>,
    document_end: u32,
) -> Result<Vec<u8>, DocWriteError> {
    let mut output = Vec::with_capacity(4 + start_order.len() * 10);
    for &index in start_order {
        output.extend_from_slice(&entries[index].start.to_le_bytes());
    }
    output.extend_from_slice(&document_end.to_le_bytes());
    for &index in start_order {
        let entry = &entries[index];
        output.extend_from_slice(
            &end_indexes
                .get(&index)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("smart-tag end index is missing".to_string())
                })?
                .to_le_bytes(),
        );
        output.extend_from_slice(&bkc(entry)?.to_le_bytes());
        output.extend_from_slice(
            &depths
                .get(&entry.start)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("smart-tag start depth is missing".to_string())
                })?
                .0
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn build_ends(
    entries: &[DocSmartTagEntry],
    end_order: &[usize],
    start_indexes: &HashMap<usize, u16>,
    depths: &BTreeMap<u32, (u16, u16)>,
    document_end: u32,
) -> Result<Vec<u8>, DocWriteError> {
    let mut output = Vec::with_capacity(4 + end_order.len() * 8);
    for &index in end_order {
        output.extend_from_slice(&entries[index].end.to_le_bytes());
    }
    output.extend_from_slice(&document_end.to_le_bytes());
    for &index in end_order {
        let entry = &entries[index];
        output.extend_from_slice(
            &start_indexes
                .get(&index)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("smart-tag start index is missing".to_string())
                })?
                .to_le_bytes(),
        );
        output.extend_from_slice(
            &depths
                .get(&entry.end)
                .ok_or_else(|| {
                    DocWriteError::InvalidData("smart-tag end depth is missing".to_string())
                })?
                .1
                .to_le_bytes(),
        );
    }
    Ok(output)
}

fn build_factoid_data(
    entries: &[DocSmartTagEntry],
    start_order: &[usize],
) -> Result<Vec<u8>, DocWriteError> {
    let mut type_ids = HashMap::<(&str, &str, &str), u16>::new();
    let mut types = Vec::new();
    let mut string_indexes = HashMap::<&str, u32>::new();
    let mut strings = Vec::new();
    let mut bags = Vec::with_capacity(entries.len());

    for &index in start_order {
        let entry = &entries[index];
        let type_key = (
            entry.namespace_uri.as_str(),
            entry.tag_name.as_str(),
            entry.download_url.as_str(),
        );
        let type_id = if let Some(id) = type_ids.get(&type_key) {
            *id
        } else {
            let id = u16::try_from(types.len() + 1).map_err(|_| {
                DocWriteError::InvalidData("DOC smart-tag type count exceeds 65535".to_string())
            })?;
            type_ids.insert(type_key, id);
            types.push(SmartTagType {
                id,
                namespace_uri: unicode_string(&entry.namespace_uri),
                tag_name: unicode_string(&entry.tag_name),
                download_url: unicode_string(&entry.download_url),
            });
            id
        };
        let mut properties = Vec::with_capacity(entry.properties.len());
        for (key, value) in &entry.properties {
            let key_index = intern_string(key, &mut string_indexes, &mut strings)?;
            let value_index = intern_string(value, &mut string_indexes, &mut strings)?;
            properties.push(SmartTagProperty {
                key_index,
                value_index,
            });
        }
        bags.push(SmartTagPropertyBag {
            type_id,
            properties,
        });
    }
    PropertyBagStore {
        ansi_codepage: 1252,
        reserved_factoid_count: 0,
        types,
        strings,
    }
    .to_bytes_with_bags(&bags)
    .map_err(|error| DocWriteError::InvalidData(error.to_string()))
}

fn intern_string<'a>(
    value: &'a str,
    indexes: &mut HashMap<&'a str, u32>,
    strings: &mut Vec<PropertyBagString>,
) -> Result<u32, DocWriteError> {
    if let Some(index) = indexes.get(value) {
        return Ok(*index);
    }
    let index = u32::try_from(strings.len()).map_err(|_| {
        DocWriteError::InvalidData("DOC smart-tag string table exceeds u32".to_string())
    })?;
    indexes.insert(value, index);
    strings.push(unicode_string(value));
    Ok(index)
}

fn unicode_string(value: &str) -> PropertyBagString {
    PropertyBagString {
        value: value.to_string(),
        encoding: PropertyBagStringEncoding::Utf16,
    }
}

fn build_recognizer_ranges(
    ranges: &[SmartTagRecognizerRange],
    document_end: u32,
) -> Result<Option<Vec<u8>>, DocWriteError> {
    if ranges.is_empty() {
        return Ok(None);
    }
    let mut output = Vec::with_capacity(4 + ranges.len() * 6);
    let mut previous_end = None;
    for range in ranges {
        if range.start > range.end || range.end > document_end {
            return invalid(
                "DOC smart-tag recognizer range must be ordered and inside the document parts",
            );
        }
        if previous_end.is_some_and(|end| end != range.start) {
            return invalid("DOC smart-tag recognizer ranges must be contiguous and ordered");
        }
        if previous_end.is_none() {
            output.extend_from_slice(&range.start.to_le_bytes());
        }
        output.extend_from_slice(&range.end.to_le_bytes());
        previous_end = Some(range.end);
    }
    for range in ranges {
        output.extend_from_slice(&recognizer_state_code(range.state).to_le_bytes());
    }
    Ok(Some(output))
}

fn bkc(entry: &DocSmartTagEntry) -> Result<u16, DocWriteError> {
    let mut value = u16::from(entry.is_native) << 14;
    if let Some((first, limit)) = entry.column_range {
        if first >= limit || first > 0x7f || limit > 0x3f {
            return invalid("DOC smart-tag column range exceeds BKC limits");
        }
        value |= 0x8000 | u16::from(first) | (u16::from(limit) << 8);
    }
    Ok(value)
}

const fn origin_code(origin: SmartTagOrigin) -> u16 {
    match origin {
        SmartTagOrigin::Unknown => 0,
        SmartTagOrigin::GrammarChecker => 1,
        SmartTagOrigin::ExternalRecognizer => 2,
        SmartTagOrigin::VisualBasic => 3,
    }
}

const fn recognizer_state_code(state: SmartTagRecognizerState) -> u16 {
    match state {
        SmartTagRecognizerState::Pending => 1,
        SmartTagRecognizerState::MaybeDirty => 2,
        SmartTagRecognizerState::Dirty => 3,
        SmartTagRecognizerState::Edit => 4,
        SmartTagRecognizerState::Clean => 7,
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T, DocWriteError> {
    Err(DocWriteError::InvalidData(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builds_overlapping_empty_and_recognizer_tables() {
        let entries = vec![
            DocSmartTagEntry::new(0, 10, "urn:test", "outer").with_property("a", "1"),
            DocSmartTagEntry::new(5, 15, "urn:test", "inner").with_property("b", "2"),
            DocSmartTagEntry::new(5, 5, "urn:test", "point"),
        ];
        let ranges = vec![
            SmartTagRecognizerRange {
                start: 0,
                end: 5,
                state: SmartTagRecognizerState::Dirty,
            },
            SmartTagRecognizerRange {
                start: 5,
                end: 15,
                state: SmartTagRecognizerState::Clean,
            },
        ];
        let data = build_tables(&entries, &ranges, 20).unwrap();
        assert_eq!(data.infos.as_ref().unwrap().len(), 48);
        assert_eq!(data.starts.as_ref().unwrap().len(), 34);
        assert_eq!(data.ends.as_ref().unwrap().len(), 28);
        assert!(data.factoid_data.as_ref().unwrap().len() > 12);
        assert_eq!(data.recognizer_ranges.as_ref().unwrap().len(), 16);
    }

    #[test]
    fn rejects_invalid_ranges_columns_and_recognizer_gaps() {
        assert!(build_tables(&[DocSmartTagEntry::new(2, 1, "u", "t")], &[], 10).is_err());
        assert!(
            build_tables(
                &[DocSmartTagEntry::new(0, 1, "u", "t").with_column_range(3, 3)],
                &[],
                10,
            )
            .is_err()
        );
        assert!(
            build_tables(
                &[],
                &[
                    SmartTagRecognizerRange {
                        start: 0,
                        end: 2,
                        state: SmartTagRecognizerState::Pending,
                    },
                    SmartTagRecognizerRange {
                        start: 3,
                        end: 4,
                        state: SmartTagRecognizerState::Clean,
                    },
                ],
                10,
            )
            .is_err()
        );
    }
}
