//! Shared record-header and lossless raw-record serialization helpers.

use crate::consts::RecordType;
use crate::package::{Error, Result};

pub(super) fn wrap_record(
    record_type: RecordType,
    version: u16,
    instance: u16,
    data: Vec<u8>,
) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len()).map_err(|_err| {
        Error::InvalidFormat(format!("{record_type:?} data exceeds 4 GiB record limit"))
    })?;
    let mut result = create_record_header(record_type, version, instance, length);
    result.extend(data);
    Ok(result)
}

/// Create a PPT record header.
pub(super) fn create_record_header(
    record_type: RecordType,
    version: u16,
    instance: u16,
    data_length: u32,
) -> Vec<u8> {
    create_record_header_raw(record_type.as_u16(), version, instance, data_length)
}

pub(super) fn create_record_header_raw(
    record_type: u16,
    version: u16,
    instance: u16,
    data_length: u32,
) -> Vec<u8> {
    let mut header = Vec::with_capacity(8);

    let version_instance = version | (instance << 4);
    header.extend(&version_instance.to_le_bytes());

    header.extend(&record_type.to_le_bytes());

    header.extend(&data_length.to_le_bytes());

    header
}

/// Serialize raw record (for preserving unknown/complex records).
pub(super) fn serialize_raw_record(record: &crate::records::Record) -> Vec<u8> {
    let mut data = Vec::new();

    #[allow(
        clippy::cast_possible_truncation,
        reason = "raw record payloads are parsed from a 32-bit record length, so the length always fits in u32"
    )]
    let header = create_record_header_raw(
        record.record_type_raw,
        record.version,
        record.instance,
        record.data.len() as u32,
    );
    data.extend(header);
    data.extend(&record.data);

    data
}
