//! Binary record codec for bookmark summaries.

use super::model::{Bookmark, Summary};
use super::validation::{
    ENTITY_BYTES, MAX_BOOKMARKS, MAX_VALUE_BYTES, NAME_BYTES, corrupted, require_header,
    validate_bookmark, validate_summary,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

impl Summary {
    /// Parse the optional document `SummaryContainer`.
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
        let id_seed = u32::from_le_bytes(
            seed.data
                .as_slice()
                .try_into()
                .map_err(|_| Error::Corrupted("BookmarkSeedAtom is truncated".to_string()))?,
        );
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
        validate_summary(&parsed)?;
        Ok(Some(parsed))
    }

    /// Serialize the summary into a validated `SummaryContainer` record.
    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_record_bytes()?;
        let (record, end) = Record::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical SummaryContainer did not consume its bytes");
        }
        Ok(record)
    }

    /// Serialize the summary into canonical PowerPoint record bytes.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        validate_summary(self)?;
        let mut collection = record_bytes(
            0,
            2,
            RecordType::BookmarkSeedAtom.as_u16(),
            &self.id_seed.to_le_bytes(),
        )?;
        for bookmark in &self.bookmarks {
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
    let id = u32::from_le_bytes(
        entity.data[..4]
            .try_into()
            .map_err(|_| Error::Corrupted("BookmarkEntityAtom is truncated".to_string()))?,
    );
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
    let bookmark = Bookmark {
        container_instance: container.instance,
        id,
        name,
        value,
    };
    validate_bookmark(&bookmark)?;
    Ok(bookmark)
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

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    if version > 0x0f || instance > 0x0fff {
        return corrupted("PowerPoint record header fields exceed their wire widths");
    }
    let length = u32::try_from(data.len())
        .map_err(|_| Error::Corrupted("PowerPoint record payload exceeds u32".to_string()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}
