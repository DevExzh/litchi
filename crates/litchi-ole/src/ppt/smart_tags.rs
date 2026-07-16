//! PowerPoint 11 smart-tag store parsing.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

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
    pub ansi_codepage: u32,
    pub types: Vec<PowerPointSmartTagType>,
    pub string_table: Vec<String>,
    pub tags: Vec<PowerPointSmartTag>,
}

impl PowerPointSmartTagStore {
    /// Parse the optional `___PPT11` smart-tag store using Windows-1252 for ANSI strings.
    pub fn parse(root: &PptRecord) -> Result<Option<Self>> {
        Self::parse_with_ansi_codepage(root, 1252)
    }

    /// Parse the optional `___PPT11` smart-tag store with an explicit ANSI code page.
    pub fn parse_with_ansi_codepage(root: &PptRecord, codepage: u32) -> Result<Option<Self>> {
        if litchi_core::encoding::codepage_to_encoding(codepage).is_none() {
            return Err(PptError::Corrupted(format!(
                "Unsupported smart-tag ANSI code page {codepage}"
            )));
        }
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
            store = Some(parse_store(&record, codepage)?);
        }
        Ok(store)
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

fn parse_store(record: &PptRecord, codepage: u32) -> Result<PowerPointSmartTagStore> {
    if record.record_type != PptRecordType::SmartTagStore11
        || record.version != 0x0f
        || record.instance != 0
    {
        return Err(PptError::Corrupted(
            "SmartTagStore11Container has an invalid record header".to_string(),
        ));
    }
    let data = &record.data;
    let mut offset = 0usize;
    let bag_count = read_u32(data, &mut offset, "smart-tag bag count")?;
    let type_count = read_u32(data, &mut offset, "smart-tag type count")?;
    let type_count = bounded_count(
        type_count,
        data.len().saturating_sub(offset),
        14,
        "smart-tag type count",
    )?;
    let mut types = Vec::with_capacity(type_count);
    for _ in 0..type_count {
        let size = usize::try_from(read_u32(data, &mut offset, "factoid size")?)
            .map_err(|_| PptError::Corrupted("FactoidType size overflows usize".to_string()))?;
        let end = offset
            .checked_add(size)
            .ok_or_else(|| PptError::Corrupted("FactoidType size overflow".to_string()))?;
        if end > data.len() {
            return Err(PptError::Corrupted("FactoidType is truncated".to_string()));
        }
        let id = read_u32(data, &mut offset, "factoid type id")?;
        let id = u16::try_from(id)
            .map_err(|_| PptError::Corrupted("FactoidType id exceeds 0xFFFF".to_string()))?;
        let namespace_uri = parse_pb_string(data, &mut offset, codepage)?;
        let tag_name = parse_pb_string(data, &mut offset, codepage)?;
        let download_url = parse_pb_string(data, &mut offset, codepage)?;
        if offset != end {
            return Err(PptError::Corrupted(
                "FactoidType byte count does not match its contents".to_string(),
            ));
        }
        if types
            .iter()
            .any(|kind: &PowerPointSmartTagType| kind.id == id)
        {
            return Err(PptError::Corrupted(
                "PropertyBagStore has duplicate smart-tag type ids".to_string(),
            ));
        }
        types.push(PowerPointSmartTagType {
            id,
            namespace_uri,
            tag_name,
            download_url,
        });
    }

    let header_size = read_u16(data, &mut offset, "property-bag header size")?;
    let version = read_u16(data, &mut offset, "property-bag version")?;
    if header_size != 0x000c || version != 0x0100 {
        return Err(PptError::Corrupted(
            "PropertyBagStore has an invalid header size or version".to_string(),
        ));
    }
    let _reserved = read_u32(data, &mut offset, "property-bag reserved value")?;
    let string_count = read_u32(data, &mut offset, "smart-tag string count")?;
    let string_count = bounded_count(
        string_count,
        data.len().saturating_sub(offset),
        2,
        "smart-tag string count",
    )?;
    let mut string_table = Vec::with_capacity(string_count);
    for _ in 0..string_count {
        string_table.push(parse_pb_string(data, &mut offset, codepage)?);
    }

    let bag_count = bounded_count(
        bag_count,
        data.len().saturating_sub(offset),
        6,
        "smart-tag bag count",
    )?;
    let mut tags = Vec::with_capacity(bag_count);
    for _ in 0..bag_count {
        let type_id = read_u16(data, &mut offset, "smart-tag type id")?;
        let property_count = read_u16(data, &mut offset, "smart-tag property count")?;
        let reserved = read_u16(data, &mut offset, "smart-tag reserved value")?;
        if reserved != 0 {
            return Err(PptError::Corrupted(
                "PropertyBag has a nonzero reserved field".to_string(),
            ));
        }
        if !types.iter().any(|kind| kind.id == type_id) {
            return Err(PptError::Corrupted(
                "PropertyBag references an unknown smart-tag type".to_string(),
            ));
        }
        let property_count = bounded_count(
            u32::from(property_count),
            data.len().saturating_sub(offset),
            8,
            "smart-tag property count",
        )?;
        let mut properties = Vec::with_capacity(property_count);
        for _ in 0..property_count {
            let key_index = read_u32(data, &mut offset, "smart-tag property key index")?;
            let value_index = read_u32(data, &mut offset, "smart-tag property value index")?;
            let key = resolve_string(&string_table, key_index)?.to_string();
            let value = resolve_string(&string_table, value_index)?.to_string();
            properties.push(PowerPointSmartTagProperty {
                key_index,
                value_index,
                key,
                value,
            });
        }
        tags.push(PowerPointSmartTag {
            type_id,
            properties,
        });
    }
    if offset != data.len() {
        return Err(PptError::Corrupted(
            "SmartTagStore11Container has trailing bytes".to_string(),
        ));
    }
    Ok(PowerPointSmartTagStore {
        ansi_codepage: codepage,
        types,
        string_table,
        tags,
    })
}

fn parse_pb_string(data: &[u8], offset: &mut usize, codepage: u32) -> Result<String> {
    let flags = read_u16(data, offset, "PBString header")?;
    let count = usize::from(flags & 0x7fff);
    let ansi = flags & 0x8000 != 0;
    let byte_count = if ansi {
        count
    } else {
        count
            .checked_mul(2)
            .ok_or_else(|| PptError::Corrupted("PBString size overflow".to_string()))?
    };
    let end = offset
        .checked_add(byte_count)
        .ok_or_else(|| PptError::Corrupted("PBString offset overflow".to_string()))?;
    let bytes = data
        .get(*offset..end)
        .ok_or_else(|| PptError::Corrupted("PBString is truncated".to_string()))?;
    *offset = end;
    if ansi {
        litchi_core::encoding::decode_bytes(bytes, Some(codepage))
            .ok_or_else(|| PptError::Corrupted("PBString ANSI decoding failed".to_string()))
    } else {
        let units: Vec<u16> = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .take_while(|unit| *unit != 0)
            .collect();
        String::from_utf16(&units)
            .map_err(|_| PptError::Corrupted("PBString contains invalid UTF-16".to_string()))
    }
}

fn resolve_string(strings: &[String], index: u32) -> Result<&str> {
    let index = usize::try_from(index)
        .map_err(|_| PptError::Corrupted("Smart-tag string index overflow".to_string()))?;
    strings
        .get(index)
        .map(String::as_str)
        .ok_or_else(|| PptError::Corrupted("Smart-tag string index is out of range".to_string()))
}

fn bounded_count(value: u32, remaining: usize, item_minimum: usize, name: &str) -> Result<usize> {
    let value = usize::try_from(value)
        .map_err(|_| PptError::Corrupted(format!("{name} overflows usize")))?;
    if value > remaining / item_minimum {
        return Err(PptError::Corrupted(format!(
            "{name} exceeds the remaining record data"
        )));
    }
    Ok(value)
}

fn read_u16(data: &[u8], offset: &mut usize, name: &str) -> Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| PptError::Corrupted(format!("{name} offset overflow")))?;
    let bytes = data
        .get(*offset..end)
        .ok_or_else(|| PptError::Corrupted(format!("{name} is truncated")))?;
    *offset = end;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: &mut usize, name: &str) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| PptError::Corrupted(format!("{name} offset overflow")))?;
    let bytes = data
        .get(*offset..end)
        .ok_or_else(|| PptError::Corrupted(format!("{name} is truncated")))?;
    *offset = end;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
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

        assert_eq!(store.ansi_codepage, 1252);
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
        assert!(
            PowerPointSmartTagStore::parse_with_ansi_codepage(&root(11, &record), 99_999).is_err()
        );
    }
}
