//! Strict, inert PowerPoint document bookmark-summary metadata.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;
use std::collections::HashSet;

const NAME_BYTES: usize = 64;
const ENTITY_BYTES: usize = 68;
const MAX_VALUE_BYTES: usize = 510;
const MAX_BOOKMARKS: usize = 4_096;

/// One summary bookmark and its link to a text bookmark.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Bookmark {
    /// Source `BookmarkEntityAtomContainer` record instance.
    pub container_instance: u16,
    pub id: u32,
    pub name: String,
    pub value: String,
}

/// The bookmark collection in the optional document `SummaryContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct BookmarkSummary {
    pub id_seed: u32,
    pub bookmarks: Vec<Bookmark>,
}

impl BookmarkSummary {
    pub fn parse(document: &Record) -> Result<Option<Self>> {
        let summaries = document
            .children
            .iter()
            .filter(|record| record.record_type_raw == RecordType::Summary.as_u16())
            .collect::<Vec<_>>();
        if summaries.len() > 1 {
            return corrupted("DocumentContainer contains duplicate SummaryContainer records");
        }
        let Some(summary) = summaries.first() else {
            return Ok(None);
        };
        require_header(summary, 0x0f, 0, RecordType::Summary, "SummaryContainer")?;
        let summary_children = Record::parse_sequence_strict(&summary.data, "SummaryContainer")?;
        if summary_children.len() != 1 {
            return corrupted("SummaryContainer must contain one BookmarkCollectionContainer");
        }
        let collection = &summary_children[0];
        require_header(
            collection,
            0x0f,
            0,
            RecordType::BookmarkCollection,
            "BookmarkCollectionContainer",
        )?;
        let children =
            Record::parse_sequence_strict(&collection.data, "BookmarkCollectionContainer")?;
        let Some(seed) = children.first() else {
            return corrupted("BookmarkCollectionContainer is missing BookmarkSeedAtom");
        };
        require_header(seed, 0, 2, RecordType::BookmarkSeedAtom, "BookmarkSeedAtom")?;
        if seed.data.len() != 4 || seed.data_length != 4 {
            return corrupted("BookmarkSeedAtom must contain four bytes");
        }
        if children.len().saturating_sub(1) > MAX_BOOKMARKS {
            return corrupted(format!(
                "bookmark collection exceeds {MAX_BOOKMARKS} entries"
            ));
        }
        let id_seed = u32::from_le_bytes(seed.data.as_slice().try_into().expect("four bytes"));
        let mut bookmarks = Vec::with_capacity(children.len().saturating_sub(1));
        let mut ids = HashSet::with_capacity(children.len().saturating_sub(1));
        for child in &children[1..] {
            let bookmark = parse_bookmark(child)?;
            if !ids.insert(bookmark.id) {
                return corrupted(format!("duplicate PowerPoint bookmark ID {}", bookmark.id));
            }
            bookmarks.push(bookmark);
        }
        let parsed = Self { id_seed, bookmarks };
        parsed.validate_seed(std::iter::empty())?;
        Ok(Some(parsed))
    }

    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_record_bytes()?;
        let (record, end) = Record::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical SummaryContainer did not consume its bytes");
        }
        Ok(record)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate_seed(std::iter::empty())?;
        if self.bookmarks.len() > MAX_BOOKMARKS {
            return corrupted(format!(
                "bookmark collection exceeds {MAX_BOOKMARKS} entries"
            ));
        }
        let mut collection = record_bytes(
            0,
            2,
            RecordType::BookmarkSeedAtom.as_u16(),
            &self.id_seed.to_le_bytes(),
        )?;
        let mut ids = HashSet::with_capacity(self.bookmarks.len());
        for bookmark in &self.bookmarks {
            if bookmark.container_instance > 0x0fff {
                return corrupted("bookmark container instance exceeds 12 bits");
            }
            if !ids.insert(bookmark.id) {
                return corrupted(format!("duplicate PowerPoint bookmark ID {}", bookmark.id));
            }
            let mut entity_data = Vec::with_capacity(ENTITY_BYTES);
            entity_data.extend_from_slice(&bookmark.id.to_le_bytes());
            entity_data.extend_from_slice(&encode_name(&bookmark.name)?);
            let mut entity_children =
                record_bytes(0, 0, RecordType::BookmarkEntityAtom.as_u16(), &entity_data)?;
            entity_children.extend_from_slice(&record_bytes(
                0,
                1,
                RecordType::CString.as_u16(),
                &encode_value(&bookmark.value)?,
            )?);
            collection.extend_from_slice(&record_bytes(
                0x0f,
                bookmark.container_instance,
                RecordType::BookmarkEntityAtom.as_u16(),
                &entity_children,
            )?);
        }
        let collection = record_bytes(
            0x0f,
            0,
            RecordType::BookmarkCollection.as_u16(),
            &collection,
        )?;
        record_bytes(0x0f, 0, RecordType::Summary.as_u16(), &collection)
    }

    /// Validate summary IDs against document `TextBookMarkAtom` identifiers.
    pub fn validate_text_bookmark_ids(
        &self,
        text_bookmark_ids: impl IntoIterator<Item = u32>,
    ) -> Result<()> {
        let mut text_ids = HashSet::new();
        for id in text_bookmark_ids {
            if text_ids.len() >= MAX_BOOKMARKS {
                return corrupted(format!(
                    "text bookmark collection exceeds {MAX_BOOKMARKS} entries"
                ));
            }
            if !text_ids.insert(id) {
                return corrupted(format!("duplicate TextBookMarkAtom ID {id}"));
            }
        }
        for bookmark in &self.bookmarks {
            if !text_ids.contains(&bookmark.id) {
                return corrupted(format!(
                    "summary bookmark ID {} has no TextBookMarkAtom",
                    bookmark.id
                ));
            }
        }
        if text_ids.len() != self.bookmarks.len() {
            return corrupted("a TextBookMarkAtom has no summary bookmark entity");
        }
        self.validate_seed(text_ids)
    }

    fn validate_seed(&self, other_ids: impl IntoIterator<Item = u32>) -> Result<()> {
        let max_id = self
            .bookmarks
            .iter()
            .map(|bookmark| bookmark.id)
            .chain(other_ids)
            .max();
        if max_id.is_some_and(|id| self.id_seed <= id) {
            return corrupted("bookmark ID seed must exceed every existing bookmark ID");
        }
        Ok(())
    }
}

fn parse_bookmark(container: &Record) -> Result<Bookmark> {
    if container.version != 0x0f
        || container.record_type_raw != RecordType::BookmarkEntityAtom.as_u16()
    {
        return corrupted("invalid BookmarkEntityAtomContainer header");
    }
    let children = Record::parse_sequence_strict(&container.data, "BookmarkEntityAtomContainer")?;
    if children.len() != 2 {
        return corrupted("BookmarkEntityAtomContainer must contain entity and value atoms");
    }
    let entity = &children[0];
    require_header(
        entity,
        0,
        0,
        RecordType::BookmarkEntityAtom,
        "BookmarkEntityAtom",
    )?;
    if entity.data.len() != ENTITY_BYTES || entity.data_length != ENTITY_BYTES as u32 {
        return corrupted("BookmarkEntityAtom must contain 68 bytes");
    }
    let id = u32::from_le_bytes(entity.data[..4].try_into().expect("four bytes"));
    let name = parse_name(&entity.data[4..])?;
    let value_atom = &children[1];
    require_header(value_atom, 0, 1, RecordType::CString, "BookmarkValueAtom")?;
    if value_atom.data.is_empty()
        || value_atom.data.len() > MAX_VALUE_BYTES
        || value_atom.data.len() % 2 != 0
    {
        return corrupted("BookmarkValueAtom length must be even and between 2 and 510 bytes");
    }
    let value = parse_printable(&value_atom.data, "BookmarkValueAtom")?;
    Ok(Bookmark {
        container_instance: container.instance,
        id,
        name,
        value,
    })
}

fn parse_name(data: &[u8]) -> Result<String> {
    if data.len() != NAME_BYTES {
        return corrupted("bookmarkName must occupy 64 bytes");
    }
    let name = parse_utf16_terminated(data, "bookmarkName")?;
    if name.is_empty() {
        return corrupted("bookmarkName cannot be empty");
    }
    Ok(name)
}

fn parse_printable(data: &[u8], context: &str) -> Result<String> {
    let value = parse_utf16_terminated(data, context)?;
    if value
        .encode_utf16()
        .any(|unit| matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f))
    {
        return corrupted(format!("{context} contains a non-printable character"));
    }
    Ok(value)
}

fn parse_utf16_terminated(data: &[u8], context: &str) -> Result<String> {
    let mut units = Vec::with_capacity(data.len() / 2);
    for bytes in data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| Error::Corrupted(format!("{context} contains invalid UTF-16")))
}

fn encode_name(name: &str) -> Result<[u8; NAME_BYTES]> {
    let units = name.encode_utf16().collect::<Vec<_>>();
    if units.is_empty() || units.len() > NAME_BYTES / 2 || units.contains(&0) {
        return corrupted("bookmarkName must contain 1 through 32 non-null UTF-16 code units");
    }
    let mut data = [0; NAME_BYTES];
    for (slot, unit) in data.chunks_exact_mut(2).zip(units) {
        slot.copy_from_slice(&unit.to_le_bytes());
    }
    Ok(data)
}

fn encode_value(value: &str) -> Result<Vec<u8>> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.contains(&0)
        || units
            .iter()
            .any(|unit| matches!(*unit, 0x0001..=0x001f | 0x007f..=0x009f))
    {
        return corrupted("BookmarkValueAtom contains a non-printable character");
    }
    let encoded_units = if units.is_empty() { 1 } else { units.len() };
    if encoded_units > MAX_VALUE_BYTES / 2 {
        return corrupted("BookmarkValueAtom exceeds 255 UTF-16 code units");
    }
    if units.is_empty() {
        return Ok(vec![0, 0]);
    }
    Ok(units.into_iter().flat_map(u16::to_le_bytes).collect())
}

fn require_header(
    record: &Record,
    version: u16,
    instance: u16,
    record_type: RecordType,
    context: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.record_type_raw != record_type.as_u16()
    {
        return corrupted(format!("invalid {context} record header"));
    }
    Ok(())
}

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| Error::Corrupted("PowerPoint record payload exceeds u32".to_string()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn root(children: Vec<Record>) -> Record {
        Record {
            version: 0x0f,
            instance: 0,
            record_type: RecordType::Document,
            record_type_raw: RecordType::Document.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    fn summary() -> BookmarkSummary {
        BookmarkSummary {
            id_seed: 43,
            bookmarks: vec![
                Bookmark {
                    container_instance: 7,
                    id: 41,
                    name: "Revenue".into(),
                    value: "FY 2026".into(),
                },
                Bookmark {
                    container_instance: 0,
                    id: 42,
                    name: "EmptyValue".into(),
                    value: String::new(),
                },
            ],
        }
    }

    #[test]
    fn protocol_shaped_bookmark_summary_roundtrips() {
        let expected = summary();
        let parsed = BookmarkSummary::parse(&root(vec![expected.to_record().unwrap()]))
            .unwrap()
            .unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
        parsed.validate_text_bookmark_ids([41, 42]).unwrap();
    }

    #[test]
    fn rejects_hostile_ids_names_values_and_seed() {
        let record = summary().to_record().unwrap();
        assert!(BookmarkSummary::parse(&root(vec![record.clone(), record])).is_err());
        let mut value = summary();
        value.id_seed = 42;
        assert!(value.to_record_bytes().is_err());
        value = summary();
        value.bookmarks[1].id = 41;
        assert!(value.to_record_bytes().is_err());
        value = summary();
        value.bookmarks[0].name.clear();
        assert!(value.to_record_bytes().is_err());
        value = summary();
        value.bookmarks[0].value = "bad\nvalue".into();
        assert!(value.to_record_bytes().is_err());
        value = summary();
        assert!(value.validate_text_bookmark_ids([41]).is_err());
        assert!(value.validate_text_bookmark_ids([41, 42, 99]).is_err());
    }
}
