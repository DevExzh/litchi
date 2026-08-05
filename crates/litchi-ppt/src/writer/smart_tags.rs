//! PowerPoint 11 smart-tag store and text-run authoring.

use super::core::PptWriteError;
use super::records::{RecordBuilder, record_type};
use crate::consts::PptRecordType;
use litchi_codepage::Ansi;
use litchi_ole_common::smart_tags::{
    Property, PropertyBag, PropertyBagStore, PropertyBagString, PropertyBagStringEncoding, Type,
};
use std::collections::HashMap;

const PPT11_TAG_NAME: &str = "___PPT11";
const MAX_SMART_TAGS: usize = 1_000_000;
const MAX_STYLE_RUNS: usize = 64 * 1024;
const MAX_STYLE_PAYLOAD_BYTES: usize = 1024 * 1024;
const SPECIAL_INFO_PP10_EXTENSION: u32 = 0x20;
const SPECIAL_INFO_SMART_TAG: u32 = 0x0200;

/// Typed zero-based reference into a PowerPoint 11 smart-tag store.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PowerPointSmartTagIndex(u32);

impl PowerPointSmartTagIndex {
    pub(crate) const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Return the encoded zero-based `SmartTagIndex`.
    pub const fn as_u32(self) -> u32 {
        self.0
    }
}

/// One document-wide smart tag that can be referenced by text runs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSmartTagDefinition {
    pub namespace_uri: String,
    pub tag_name: String,
    pub download_url: String,
    pub properties: Vec<(String, String)>,
}

impl PowerPointSmartTagDefinition {
    /// Create a smart tag with an empty property bag.
    pub fn new(namespace_uri: impl Into<String>, tag_name: impl Into<String>) -> Self {
        Self {
            namespace_uri: namespace_uri.into(),
            tag_name: tag_name.into(),
            download_url: String::new(),
            properties: Vec::new(),
        }
    }

    /// Set the inert recognizer download URL retained in the type declaration.
    pub fn with_download_url(mut self, value: impl Into<String>) -> Self {
        self.download_url = value.into();
        self
    }

    /// Append a key/value property.
    pub fn with_property(mut self, key: impl Into<String>, value: impl Into<String>) -> Self {
        self.properties.push((key.into(), value.into()));
        self
    }

    /// Append a key/value property in place.
    pub fn add_property(&mut self, key: impl Into<String>, value: impl Into<String>) {
        self.properties.push((key.into(), value.into()));
    }
}

pub(crate) fn build_document_binary_tag(
    entries: &[PowerPointSmartTagDefinition],
) -> Result<Option<Vec<u8>>, PptWriteError> {
    if entries.is_empty() {
        return Ok(None);
    }
    if entries.len() > MAX_SMART_TAGS {
        return invalid("PowerPoint smart-tag store exceeds one million entries");
    }

    let mut type_ids = HashMap::<(&str, &str, &str), u16>::new();
    let mut types = Vec::new();
    let mut string_indexes = HashMap::<&str, u32>::new();
    let mut strings = Vec::new();
    let mut bags = Vec::with_capacity(entries.len());

    for entry in entries {
        validate_definition(entry)?;
        let type_key = (
            entry.namespace_uri.as_str(),
            entry.tag_name.as_str(),
            entry.download_url.as_str(),
        );
        let type_id = if let Some(id) = type_ids.get(&type_key) {
            *id
        } else {
            let id = u16::try_from(types.len() + 1).map_err(|_| {
                PptWriteError::InvalidData(
                    "PowerPoint smart-tag type count exceeds 65535".to_string(),
                )
            })?;
            type_ids.insert(type_key, id);
            types.push(Type {
                id,
                namespace_uri: unicode_string(&entry.namespace_uri),
                tag_name: unicode_string(&entry.tag_name),
                download_url: unicode_string(&entry.download_url),
            });
            id
        };

        let mut properties = Vec::with_capacity(entry.properties.len());
        for (key, value) in &entry.properties {
            properties.push(Property {
                key_index: intern_string(key, &mut string_indexes, &mut strings)?,
                value_index: intern_string(value, &mut string_indexes, &mut strings)?,
            });
        }
        bags.push(PropertyBag {
            type_id,
            properties,
        });
    }

    let shared = PropertyBagStore {
        ansi: Ansi::WINDOWS_1252,
        reserved_factoid_count: 0,
        types,
        strings,
    };
    let mut store_payload = u32::try_from(bags.len())
        .map_err(|_| {
            PptWriteError::InvalidData("PowerPoint smart-tag count exceeds u32".to_string())
        })?
        .to_le_bytes()
        .to_vec();
    store_payload.extend(
        shared
            .to_bytes_with_bags(&bags)
            .map_err(|error| PptWriteError::InvalidData(error.to_string()))?,
    );

    let mut store = RecordBuilder::new(0x0f, 0, PptRecordType::SmartTagStore11.as_u16());
    store.write_data(&store_payload);
    Ok(Some(build_binary_tag(PPT11_TAG_NAME, &[store.build()?])?))
}

pub(crate) fn build_shape_programmable_tags(
    runs: &[Vec<u32>],
) -> Result<Option<Vec<u8>>, PptWriteError> {
    if runs.iter().all(Vec::is_empty) {
        return Ok(None);
    }
    if runs.len() > MAX_STYLE_RUNS {
        return invalid("PowerPoint smart-tag text mapping exceeds 65536 runs");
    }
    let style9_len = runs
        .len()
        .checked_mul(16)
        .filter(|length| *length <= MAX_STYLE_PAYLOAD_BYTES)
        .ok_or_else(|| {
            PptWriteError::InvalidData("PowerPoint 9 text-style mapping exceeds 1 MiB".to_string())
        })?;
    let style11_len = runs.iter().try_fold(0usize, |total, smart_tags| {
        let entry = if smart_tags.is_empty() {
            4
        } else {
            smart_tags
                .len()
                .checked_mul(4)
                .and_then(|bytes| bytes.checked_add(8))
                .ok_or_else(|| {
                    PptWriteError::InvalidData(
                        "PowerPoint 11 smart-tag mapping size overflows".to_string(),
                    )
                })?
        };
        total.checked_add(entry).ok_or_else(|| {
            PptWriteError::InvalidData("PowerPoint 11 smart-tag mapping size overflows".to_string())
        })
    })?;
    if style11_len > MAX_STYLE_PAYLOAD_BYTES {
        return invalid("PowerPoint 11 smart-tag mapping exceeds 1 MiB");
    }
    let mut style9 = Vec::with_capacity(style9_len);
    let mut style11 = Vec::with_capacity(style11_len);
    for (index, smart_tags) in runs.iter().enumerate() {
        style9.extend_from_slice(&0u32.to_le_bytes());
        style9.extend_from_slice(&0u32.to_le_bytes());
        style9.extend_from_slice(&SPECIAL_INFO_PP10_EXTENSION.to_le_bytes());
        style9.extend_from_slice(
            &u32::try_from(index % 16)
                .expect("modulo 16 always fits u32")
                .to_le_bytes(),
        );

        if smart_tags.is_empty() {
            style11.extend_from_slice(&0u32.to_le_bytes());
        } else {
            style11.extend_from_slice(&SPECIAL_INFO_SMART_TAG.to_le_bytes());
            style11.extend_from_slice(
                &u32::try_from(smart_tags.len())
                    .map_err(|_| {
                        PptWriteError::InvalidData(
                            "PowerPoint text run smart-tag count exceeds u32".to_string(),
                        )
                    })?
                    .to_le_bytes(),
            );
            for index in smart_tags {
                style11.extend_from_slice(&index.to_le_bytes());
            }
        }
    }

    let style9_record = atom(PptRecordType::StyleTextProp9Atom, &style9)?;
    let style11_record = atom(PptRecordType::StyleTextProp11Atom, &style11)?;
    let ppt9 = build_binary_tag("___PPT9", &[style9_record])?;
    let ppt11 = build_binary_tag(PPT11_TAG_NAME, &[style11_record])?;
    let mut tags = RecordBuilder::new(0x0f, 0, record_type::PROG_TAGS);
    tags.write_child(&ppt9);
    tags.write_child(&ppt11);
    Ok(Some(tags.build()?))
}

fn build_binary_tag(name: &str, records: &[Vec<u8>]) -> Result<Vec<u8>, PptWriteError> {
    let mut binary_tag = RecordBuilder::new(0x0f, 0, record_type::PROG_BINARY_TAG);
    let mut name_atom = RecordBuilder::new(0, 0, record_type::CSTRING);
    let name_bytes = name
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    name_atom.write_data(&name_bytes);
    binary_tag.write_child(&name_atom.build()?);

    let mut blob = RecordBuilder::new(0, 0, record_type::BINARY_TAG_DATA);
    for record in records {
        blob.write_child(record);
    }
    binary_tag.write_child(&blob.build()?);
    Ok(binary_tag.build()?)
}

fn atom(kind: PptRecordType, data: &[u8]) -> Result<Vec<u8>, PptWriteError> {
    let mut atom = RecordBuilder::new(0, 0, kind.as_u16());
    atom.write_data(data);
    Ok(atom.build()?)
}

pub(crate) fn validate_definition(
    entry: &PowerPointSmartTagDefinition,
) -> Result<(), PptWriteError> {
    if entry.namespace_uri.contains('\0')
        || entry.tag_name.contains('\0')
        || entry.download_url.contains('\0')
        || entry
            .properties
            .iter()
            .any(|(key, value)| key.contains('\0') || value.contains('\0'))
    {
        return invalid("PowerPoint smart-tag strings cannot contain NUL characters");
    }
    if entry.properties.len() > usize::from(u16::MAX) {
        return invalid("PowerPoint smart-tag property bag exceeds 65535 properties");
    }
    if [
        entry.namespace_uri.as_str(),
        entry.tag_name.as_str(),
        entry.download_url.as_str(),
    ]
    .into_iter()
    .chain(
        entry
            .properties
            .iter()
            .flat_map(|(key, value)| [key.as_str(), value.as_str()]),
    )
    .any(|value| value.encode_utf16().count() > 0x7fff)
    {
        return invalid("PowerPoint smart-tag string exceeds 32767 UTF-16 code units");
    }
    Ok(())
}

fn intern_string<'a>(
    value: &'a str,
    indexes: &mut HashMap<&'a str, u32>,
    strings: &mut Vec<PropertyBagString>,
) -> Result<u32, PptWriteError> {
    if let Some(index) = indexes.get(value) {
        return Ok(*index);
    }
    let index = u32::try_from(strings.len()).map_err(|_| {
        PptWriteError::InvalidData("PowerPoint smart-tag string table exceeds u32".to_string())
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

fn invalid<T>(message: impl Into<String>) -> Result<T, PptWriteError> {
    Err(PptWriteError::InvalidData(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::records::PptRecord;
    use crate::{
        PowerPointShapeProgrammableTagLimits, PowerPointShapeProgrammableTags,
        PowerPointSmartTagStore,
    };

    fn root_with_document_tag(binary_tag: &[u8]) -> PptRecord {
        let mut tags = RecordBuilder::new(0x0f, 0, record_type::PROG_TAGS);
        tags.write_child(binary_tag);
        let bytes = tags.build().unwrap();
        let (programmable_tags, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: PptRecordType::Document.as_u16(),
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(bytes.len()).unwrap(),
            data: Vec::new(),
            children: vec![programmable_tags],
        }
    }

    #[test]
    fn document_store_deduplicates_types_and_round_trips_unicode() {
        let entries = vec![
            PowerPointSmartTagDefinition::new("urn:geo", "place").with_property("city", "東京"),
            PowerPointSmartTagDefinition::new("urn:geo", "place").with_property("city", "Paris"),
        ];
        let binary_tag = build_document_binary_tag(&entries).unwrap().unwrap();
        let store = PowerPointSmartTagStore::parse(&root_with_document_tag(&binary_tag))
            .unwrap()
            .unwrap();
        assert_eq!(store.types.len(), 1);
        assert_eq!(store.tags.len(), 2);
        assert_eq!(store.tags[0].properties[0].value, "東京");
    }

    #[test]
    fn shape_run_mappings_round_trip_and_reject_invalid_strings() {
        let bytes = build_shape_programmable_tags(&[vec![0, 1], Vec::new(), vec![1]])
            .unwrap()
            .unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        let tags = PowerPointShapeProgrammableTags::parse(
            &record,
            PowerPointShapeProgrammableTagLimits::default(),
        )
        .unwrap();
        assert_eq!(tags.powerpoint9().unwrap().runs.len(), 3);
        assert_eq!(
            tags.powerpoint11().unwrap().runs[0].smart_tag_indices,
            vec![0, 1]
        );
        assert!(
            tags.powerpoint11().unwrap().runs[1]
                .smart_tag_indices
                .is_empty()
        );

        let invalid = PowerPointSmartTagDefinition::new("urn:test", "bad\0name");
        assert!(build_document_binary_tag(&[invalid]).is_err());
    }
}
