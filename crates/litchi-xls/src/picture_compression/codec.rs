//! Bounded BIFF framing and `CompressPictures` payload codec.

use crate::{Error, Result};

use super::model::{
    HEADER_LEN, MAX_RECORD_PAYLOAD, MAX_RECORDS, MAX_STREAM_BYTES, MIN_PAYLOAD_LEN, Record,
    Settings, Snapshot, Unknown,
};
use super::{RECORD_TYPE, validation};

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn parse(input: &[u8]) -> Result<Snapshot> {
    if input.len() > MAX_STREAM_BYTES {
        return Err(invalid("record stream exceeds 1 MiB"));
    }
    let mut offset = 0usize;
    let mut records = Vec::new();
    while offset < input.len() {
        if records.len() == MAX_RECORDS {
            return Err(invalid("record count exceeds 1024"));
        }
        let header_end = offset
            .checked_add(4)
            .ok_or_else(|| invalid("record header offset overflows"))?;
        if header_end > input.len() {
            return Err(Error::UnexpectedEndOfStream(
                "CompressPictures record header is truncated".to_string(),
            ));
        }
        let record_type = u16::from_le_bytes([input[offset], input[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([input[offset + 2], input[offset + 3]]));
        if payload_len > MAX_RECORD_PAYLOAD {
            return Err(invalid("BIFF payload exceeds 8224 bytes"));
        }
        let payload_start = header_end;
        let payload_end = payload_start
            .checked_add(payload_len)
            .ok_or_else(|| invalid("record payload offset overflows"))?;
        if payload_end > input.len() {
            return Err(Error::UnexpectedEndOfStream(
                "CompressPictures record payload is truncated".to_string(),
            ));
        }
        let payload = &input[payload_start..payload_end];
        records.push(if record_type == RECORD_TYPE {
            Record::Settings(parse_settings(payload)?)
        } else {
            Record::Unknown(Unknown::try_new(record_type, payload.to_vec())?)
        });
        offset = payload_end;
    }
    Snapshot::try_new(records)
}

/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn write(snapshot: &Snapshot) -> Result<Vec<u8>> {
    validation::validate(snapshot)?;
    let mut output = Vec::new();
    for record in &snapshot.records {
        match record {
            Record::Settings(value) => {
                let payload = value.payload()?;
                append_record(&mut output, RECORD_TYPE, &payload)?;
            },
            Record::Unknown(value) => {
                append_record(&mut output, value.record_type(), value.payload())?;
            },
        }
    }
    Ok(output)
}

fn append_record(output: &mut Vec<u8>, record_type: u16, payload: &[u8]) -> Result<()> {
    let payload_len = u16::try_from(payload.len())
        .map_err(|_error| invalid("record payload length does not fit BIFF framing"))?;
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&payload_len.to_le_bytes());
    output.extend_from_slice(payload);
    Ok(())
}

pub(crate) fn parse_settings(payload: &[u8]) -> Result<Settings> {
    if payload.len() < MIN_PAYLOAD_LEN {
        return Err(Error::InvalidLength {
            expected: MIN_PAYLOAD_LEN,
            found: payload.len(),
        });
    }
    let mut header = [0; HEADER_LEN];
    header.copy_from_slice(&payload[..HEADER_LEN]);
    if u16::from_le_bytes([header[0], header[1]]) != RECORD_TYPE {
        return Err(invalid("FrtHeader.rt does not match CompressPictures"));
    }
    if u16::from_le_bytes([header[2], header[3]]) != 0 || header[4..].iter().any(|byte| *byte != 0)
    {
        return Err(invalid(
            "CompressPictures FrtHeader has nonzero reserved fields",
        ));
    }
    let recommendation = match u32::from_le_bytes(payload[12..16].try_into().unwrap()) {
        0 => false,
        1 => true,
        _ => return Err(invalid("fAutoCompressPictures must be 0 or 1")),
    };
    Ok(Settings::from_wire(
        recommendation,
        header,
        payload[MIN_PAYLOAD_LEN..].to_vec().into_boxed_slice(),
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: RECORD_TYPE,
        message: message.into(),
    }
}
