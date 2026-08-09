//! Structural validation for diagram-build records.

use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

pub(super) const HEADER_LEN: usize = 8;
pub(super) const BUILD_PAYLOAD_LEN: usize = 16;
pub(super) const ATOM_PAYLOAD_LEN: usize = 4;
pub(super) const CONTAINER_PAYLOAD_LEN: usize = 36;

#[allow(
    clippy::cast_possible_truncation,
    reason = "callers pass only the fixed 4- or 16-byte payload constants, so `payload_len` always fits in `u32`"
)]
pub(super) fn validate_atom(
    record: &Record,
    expected: RecordType,
    payload_len: usize,
    name: &str,
) -> Result<()> {
    if record.record_type != expected || record.record_type_raw != expected.as_u16() {
        return Err(Error::InvalidFormat(format!("expected {name} record type")));
    }
    if record.version != 0 || record.instance != 0 {
        return Err(Error::Corrupted(format!(
            "{name} requires record version 0 and instance 0"
        )));
    }
    if record.data_length != payload_len as u32 || record.data.len() != payload_len {
        return Err(Error::Corrupted(format!(
            "{name} requires a {payload_len}-byte payload"
        )));
    }
    if !record.children.is_empty() {
        return Err(Error::Corrupted(format!(
            "{name} is an atom and cannot contain child records"
        )));
    }
    Ok(())
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "`CONTAINER_PAYLOAD_LEN` is a compile-time 36-byte constant, so it always fits in `u32`"
)]
pub(super) fn validate_container(record: &Record) -> Result<()> {
    if record.record_type != RecordType::DiagramBuild
        || record.record_type_raw != RecordType::DiagramBuild.as_u16()
    {
        return Err(Error::InvalidFormat(
            "expected DiagramBuild container record".to_string(),
        ));
    }
    if record.version != 0x0F || record.instance != 0 {
        return Err(Error::Corrupted(
            "DiagramBuild container requires version 0xF and instance 0".to_string(),
        ));
    }
    if record.data_length != CONTAINER_PAYLOAD_LEN as u32
        || record.data.len() != CONTAINER_PAYLOAD_LEN
    {
        return Err(Error::Corrupted(
            "DiagramBuild container requires a 36-byte payload".to_string(),
        ));
    }
    if record.children.len() != 2 {
        return Err(Error::Corrupted(
            "DiagramBuild container requires BuildAtom and DiagramBuildAtom".to_string(),
        ));
    }

    let mut offset = 0usize;
    for child in &record.children {
        let child_len = HEADER_LEN
            .checked_add(child.data.len())
            .ok_or_else(|| Error::Corrupted("diagram build child length overflow".to_string()))?;
        let end = offset
            .checked_add(child_len)
            .ok_or_else(|| Error::Corrupted("diagram build payload length overflow".to_string()))?;
        let encoded = record.data.get(offset..end).ok_or_else(|| {
            Error::Corrupted("DiagramBuild child records exceed the container payload".to_string())
        })?;
        if encoded != encode_record(child).as_slice() {
            return Err(Error::Corrupted(
                "DiagramBuild child records do not match the container payload".to_string(),
            ));
        }
        offset = end;
    }
    if offset != record.data.len() {
        return Err(Error::Corrupted(
            "DiagramBuild child records do not cover the container payload".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn encode_record(record: &Record) -> Vec<u8> {
    let version_instance = (record.version & 0x000F) | ((record.instance & 0x0FFF) << 4);
    let mut bytes = Vec::with_capacity(HEADER_LEN + record.data.len());
    bytes.extend_from_slice(&version_instance.to_le_bytes());
    bytes.extend_from_slice(&record.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&record.data_length.to_le_bytes());
    bytes.extend_from_slice(&record.data);
    bytes
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "callers pass only the fixed 4- or 36-byte payload constants, so `payload_len` always fits in `u32`"
)]
pub(super) fn encode_header(
    version: u16,
    instance: u16,
    record_type: RecordType,
    payload_len: usize,
) -> [u8; HEADER_LEN] {
    let version_instance = (version & 0x000F) | ((instance & 0x0FFF) << 4);
    let mut header = [0; HEADER_LEN];
    header[0..2].copy_from_slice(&version_instance.to_le_bytes());
    header[2..4].copy_from_slice(&record_type.as_u16().to_le_bytes());
    header[4..8].copy_from_slice(&(payload_len as u32).to_le_bytes());
    header
}

pub(super) fn parse_bool(value: u8, name: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        other => Err(Error::InvalidFormat(format!(
            "{name} has invalid bool1 value {other}"
        ))),
    }
}
