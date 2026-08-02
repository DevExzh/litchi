//! PowerPoint 11 smart-tag store parsing.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use litchi_codepage::Ansi;
use litchi_ole_common::smart_tags::{PropertyBagStore, SmartTagLimits};

/// One smart-tag type declared by the shared property-bag store.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSmartTagType {
    pub id: u16,
    pub namespace_uri: String,
    pub tag_name: String,
    pub download_url: String,
}

/// One resolved key/value pair attached to a smart tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSmartTagProperty {
    pub key_index: u32,
    pub value_index: u32,
    pub key: String,
    pub value: String,
}

/// One smart tag referenced by zero-based `SmartTagIndex` values in text runs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSmartTag {
    pub type_id: u16,
    pub properties: Vec<PowerPointSmartTagProperty>,
}

/// PowerPoint 11 smart-tag types, shared strings, and property bags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSmartTagStore {
    /// ANSI code page used to decode ANSI `PBString` values.
    pub ansi: Ansi,
    pub types: Vec<PowerPointSmartTagType>,
    pub string_table: Vec<String>,
    pub tags: Vec<PowerPointSmartTag>,
}

impl PowerPointSmartTagStore {
    /// Parse the optional `___PPT11` smart-tag store using Windows-1252 for ANSI strings.
    pub fn parse(root: &PptRecord) -> Result<Option<Self>> {
        Self::parse_with(root, Ansi::WINDOWS_1252)
    }

    /// Parse with an explicitly validated ANSI page.
    pub fn parse_with(root: &PptRecord, ansi: Ansi) -> Result<Option<Self>> {
        let mut store = None;
        for record in root.versioned_binary_tag_records(11)? {
            if record.record_type != PptRecordType::SmartTagStore11 {
                continue;
            }
            if store.is_some() {
                return Err(PptError::Corrupted(
                    "Record tree contains multiple PowerPoint 11 smart-tag stores".to_string(),
                ));
            }
            store = Some(parse_store(&record, ansi)?);
        }
        Ok(store)
    }

    /// Validate a raw ANSI page identifier and parse the optional store.
    pub fn parse_page(root: &PptRecord, page: u32) -> Result<Option<Self>> {
        let ansi = Ansi::require(page).map_err(|error| PptError::Corrupted(error.to_string()))?;
        Self::parse_with(root, ansi)
    }

    /// Resolve a zero-based smart-tag index from a text run.
    pub fn get(&self, index: u32) -> Option<&PowerPointSmartTag> {
        self.tags.get(usize::try_from(index).ok()?)
    }

    /// Resolve a smart tag's declared type.
    pub fn tag_type(&self, tag: &PowerPointSmartTag) -> Option<&PowerPointSmartTagType> {
        self.types.iter().find(|kind| kind.id == tag.type_id)
    }
}

fn parse_store(record: &PptRecord, ansi: Ansi) -> Result<PowerPointSmartTagStore> {
    if record.record_type != PptRecordType::SmartTagStore11
        || record.version != 0x0f
        || record.instance != 0
    {
        return Err(PptError::Corrupted(
            "SmartTagStore11Container has an invalid record header".to_string(),
        ));
    }
    let data = &record.data;
    let bag_count_bytes = data.get(..4).ok_or_else(|| {
        PptError::Corrupted("SmartTagStore11Container is missing its bag count".to_string())
    })?;
    let bag_count = usize::try_from(u32::from_le_bytes([
        bag_count_bytes[0],
        bag_count_bytes[1],
        bag_count_bytes[2],
        bag_count_bytes[3],
    ]))
    .map_err(|_| PptError::Corrupted("smart-tag bag count overflows usize".to_string()))?;
    let limits = SmartTagLimits::default();
    let (shared, consumed) = PropertyBagStore::parse_prefix(&data[4..], ansi, limits)
        .map_err(|error| PptError::Corrupted(error.to_string()))?;
    let bags_start = 4usize
        .checked_add(consumed)
        .ok_or_else(|| PptError::Corrupted("smart-tag store offset overflows".to_string()))?;
    let bags = shared
        .parse_bags(&data[bags_start..], bag_count, limits)
        .map_err(|error| PptError::Corrupted(error.to_string()))?;

    let types = shared
        .types
        .iter()
        .map(|kind| PowerPointSmartTagType {
            id: kind.id,
            namespace_uri: kind.namespace_uri.value.clone(),
            tag_name: kind.tag_name.value.clone(),
            download_url: kind.download_url.value.clone(),
        })
        .collect();
    let string_table = shared
        .strings
        .iter()
        .map(|value| value.value.clone())
        .collect();
    let mut tags = Vec::with_capacity(bags.len());
    for bag in bags {
        let mut properties = Vec::with_capacity(bag.properties.len());
        for property in bag.properties {
            let (key, value) = shared.resolve_property(property).ok_or_else(|| {
                PptError::Corrupted(
                    "shared smart-tag property index validation was inconsistent".to_string(),
                )
            })?;
            properties.push(PowerPointSmartTagProperty {
                key_index: property.key_index,
                value_index: property.value_index,
                key: key.to_string(),
                value: value.to_string(),
            });
        }
        tags.push(PowerPointSmartTag {
            type_id: bag.type_id,
            properties,
        });
    }
    Ok(PowerPointSmartTagStore {
        ansi,
        types,
        string_table,
        tags,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn pb_ansi(bytes: &[u8]) -> Vec<u8> {
        let mut data = (0x8000 | u16::try_from(bytes.len()).unwrap())
            .to_le_bytes()
            .to_vec();
        data.extend_from_slice(bytes);
        data
    }

    fn pb_unicode(value: &str) -> Vec<u8> {
        let units: Vec<u16> = value.encode_utf16().collect();
        let mut data = u16::try_from(units.len()).unwrap().to_le_bytes().to_vec();
        for unit in units {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    fn store_payload(bag_type: u16, key_index: u32, reserved: u16) -> Vec<u8> {
        let mut data = 1u32.to_le_bytes().to_vec();
        data.extend_from_slice(&1u32.to_le_bytes());

        let mut factoid = 2u32.to_le_bytes().to_vec();
        factoid.extend_from_slice(&pb_ansi(b"urn:example:smarttags"));
        factoid.extend_from_slice(&pb_ansi(b"place"));
        factoid.extend_from_slice(&pb_ansi(b""));
        data.extend_from_slice(&u32::try_from(factoid.len()).unwrap().to_le_bytes());
        data.extend_from_slice(&factoid);

        data.extend_from_slice(&0x000cu16.to_le_bytes());
        data.extend_from_slice(&0x0100u16.to_le_bytes());
        data.extend_from_slice(&0x1234_5678u32.to_le_bytes());
        data.extend_from_slice(&2u32.to_le_bytes());
        data.extend_from_slice(&pb_unicode("City"));
        data.extend_from_slice(&pb_ansi(b"M\xfcnchen"));

        data.extend_from_slice(&bag_type.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&reserved.to_le_bytes());
        data.extend_from_slice(&key_index.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> PptRecord {
        let tag_name: Vec<u8> = format!("___PPT{version}")
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

    fn root(version: u8, payload: &[u8]) -> PptRecord {
        PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![prog_tags_record(version, payload)],
        }
    }

    fn smart_tag_record(payload: &[u8]) -> Vec<u8> {
        record_bytes(0x0f, 0, 0x36b3, payload)
    }

    #[test]
    fn parses_and_resolves_powerpoint11_smart_tags() {
        let record = smart_tag_record(&store_payload(2, 0, 0));
        let store = PowerPointSmartTagStore::parse(&root(11, &record))
            .unwrap()
            .unwrap();

        assert_eq!(store.ansi, Ansi::WINDOWS_1252);
        let tag = store.get(0).unwrap();
        let kind = store.tag_type(tag).unwrap();
        assert_eq!(kind.namespace_uri, "urn:example:smarttags");
        assert_eq!(kind.tag_name, "place");
        assert_eq!(tag.properties[0].key, "City");
        assert_eq!(tag.properties[0].value, "München");
    }

    #[test]
    fn rejects_malformed_or_duplicate_smart_tag_stores() {
        let mut oversized_count = 0u32.to_le_bytes().to_vec();
        oversized_count.extend_from_slice(&u32::MAX.to_le_bytes());
        let malformed = [
            smart_tag_record(&store_payload(3, 0, 0)),
            smart_tag_record(&store_payload(2, 2, 0)),
            smart_tag_record(&store_payload(2, 0, 1)),
            smart_tag_record(&oversized_count),
        ];
        for record in malformed {
            assert!(PowerPointSmartTagStore::parse(&root(11, &record)).is_err());
        }

        let record = smart_tag_record(&store_payload(2, 0, 0));
        let mut duplicate = record.clone();
        duplicate.extend_from_slice(&record);
        assert!(PowerPointSmartTagStore::parse(&root(11, &duplicate)).is_err());
    }

    #[test]
    fn isolates_smart_tags_by_version_and_validates_codepages() {
        let record = smart_tag_record(&store_payload(2, 0, 0));

        assert!(
            PowerPointSmartTagStore::parse(&root(10, &record))
                .unwrap()
                .is_none()
        );
        assert!(PowerPointSmartTagStore::parse_page(&root(11, &record), 99_999).is_err());
        assert!(PowerPointSmartTagStore::parse_page(&root(11, &record), 65001).is_err());
    }
}
