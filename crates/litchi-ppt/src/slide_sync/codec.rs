//! Binary framing for synchronization records.

use super::model::{LibraryUrl, Limits, ServerId, Synchronization, SystemTime};
use super::validation;
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

pub(crate) fn read(root: &Record) -> Result<Option<Synchronization>> {
    validation::validate(root, Limits::default())?;
    let Some(index) = root.children.iter().position(validation::is_sync_record) else {
        return Ok(None);
    };
    let record = &root.children[index];
    let children = children(record)?;
    let server_slide_id = ServerId::from_wire(&children[0].data)?;
    let slide_library_url = LibraryUrl::from_wire(&children[1].data)?;
    let server_modified = SystemTime::from_wire(&children[2].data[..16], "dateTimeModified")?;
    let client_inserted = SystemTime::from_wire(&children[2].data[16..], "dateTimeInserted")?;
    Synchronization::from_parts(
        server_slide_id,
        slide_library_url,
        server_modified,
        client_inserted,
    )
    .map(Some)
}

pub(crate) fn encode(root: &Record, limits: Limits) -> Result<Vec<u8>> {
    let mut encoder = Encoder { limits, records: 0 };
    encoder.record(root, 1)
}

pub(crate) fn parse(bytes: &[u8], limits: Limits) -> Result<Record> {
    if bytes.len() > limits.max_bytes {
        return invalid("slide synchronization snapshot exceeds the byte limit");
    }
    let (record, consumed) = Record::parse_strict(bytes, 0)?;
    if consumed != bytes.len() {
        return Err(Error::Corrupted(
            "slide synchronization input contains trailing bytes".into(),
        ));
    }
    Ok(record)
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "the timestamp payload is exactly 32 bytes and the container payload is bounded to 64 MiB by `encode_children`, so both lengths fit in u32"
)]
pub(crate) fn encode_sync(value: &Synchronization) -> Result<Record> {
    let server = atom(0, value.server_slide_id().wire()?);
    let url = atom(1, value.slide_library_url().wire()?);
    let mut timestamp_bytes = Vec::with_capacity(32);
    value.server_modified().write_wire(&mut timestamp_bytes);
    value.client_inserted().write_wire(&mut timestamp_bytes);
    let timestamps = Record {
        record_type: RecordType::RoundTripSlideSyncInfoAtom12,
        record_type_raw: RecordType::RoundTripSlideSyncInfoAtom12.as_u16(),
        version: 0,
        instance: 0,
        data_length: timestamp_bytes.len() as u32,
        data: timestamp_bytes,
        children: Vec::new(),
    };
    let children = vec![server, url, timestamps];
    let payload = encode_children(&children, Limits::default())?;
    Ok(Record {
        record_type: RecordType::RoundTripSlideSyncInfo12,
        record_type_raw: RecordType::RoundTripSlideSyncInfo12.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: payload.len() as u32,
        data: payload,
        children,
    })
}

struct Encoder {
    limits: Limits,
    records: usize,
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "wire payloads are validated to at most MAX_TEXT_BYTES (16 KiB), well below u32::MAX"
)]
fn atom(instance: u16, data: Vec<u8>) -> Record {
    Record {
        record_type: RecordType::CString,
        record_type_raw: RecordType::CString.as_u16(),
        version: 0,
        instance,
        data_length: data.len() as u32,
        data,
        children: Vec::new(),
    }
}

impl Encoder {
    fn record(&mut self, record: &Record, depth: usize) -> Result<Vec<u8>> {
        if depth > self.limits.max_depth {
            return invalid("slide synchronization record nesting exceeds the depth limit");
        }
        self.records = self.records.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("slide synchronization record count overflow".into())
        })?;
        if self.records > self.limits.max_records {
            return invalid("slide synchronization record count exceeds the limit");
        }
        if record.version > 0x0f || record.instance > 0x0fff {
            return invalid("PPT record header exceeds its bit fields");
        }

        let payload = if record.children.is_empty() {
            record.data.clone()
        } else {
            let mut payload = Vec::new();
            for child in &record.children {
                let encoded = self.record(child, depth + 1)?;
                let next = payload
                    .len()
                    .checked_add(encoded.len())
                    .ok_or_else(|| Error::InvalidFormat("PPT record size overflow".into()))?;
                if next > self.limits.max_bytes {
                    return invalid("PPT record payload exceeds the byte limit");
                }
                payload.extend_from_slice(&encoded);
            }
            payload
        };
        let total = payload
            .len()
            .checked_add(8)
            .ok_or_else(|| Error::InvalidFormat("PPT record size overflow".into()))?;
        if total > self.limits.max_bytes {
            return invalid("PPT record exceeds the byte limit");
        }
        let length = u32::try_from(payload.len())
            .map_err(|_err| Error::InvalidFormat("PPT record payload exceeds u32".into()))?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(total)
            .map_err(|_err| Error::InvalidFormat("PPT record allocation failed".into()))?;
        bytes.extend_from_slice(&(record.version | (record.instance << 4)).to_le_bytes());
        bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(&payload);
        Ok(bytes)
    }
}

fn children(record: &Record) -> Result<Vec<Record>> {
    validation::validate_sync_record(record)?;
    if !record.children.is_empty() {
        return Ok(record.children.clone());
    }
    Record::parse_sequence_strict(&record.data, "RoundTripSlideSyncInfo12")
}

fn encode_children(children: &[Record], limits: Limits) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    for child in children {
        let bytes = encode(child, limits)?;
        let next = output
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| Error::InvalidFormat("PPT record size overflow".into()))?;
        if next > limits.max_bytes {
            return invalid("PPT record payload exceeds the byte limit");
        }
        output.extend_from_slice(&bytes);
    }
    Ok(output)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
