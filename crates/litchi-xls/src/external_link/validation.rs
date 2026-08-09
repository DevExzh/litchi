//! Framing and resource validation for source-preserving external-link edits.

use crate::{Error, Result};

/// The maximum payload carried by one BIFF8 record.
pub(super) const MAX_RECORD_PAYLOAD: usize = 8_224;
/// A bounded workbook-global stream accepted by the external-link facade.
pub(super) const MAX_STREAM_BYTES: usize = 64 * 1024 * 1024;
/// A malformed or adversarial stream must not create an unbounded span table.
const MAX_RECORDS: usize = 1_048_576;

#[derive(Debug, Clone, Copy)]
pub(super) struct RecordSpan {
    pub(super) record_type: u16,
    pub(super) record_start: usize,
    pub(super) payload_start: usize,
    pub(super) payload_end: usize,
}

impl RecordSpan {
    pub(super) fn payload(self, bytes: &[u8]) -> &[u8] {
        &bytes[self.payload_start..self.payload_end]
    }
}

pub(super) fn scan_records(bytes: &[u8]) -> Result<Vec<RecordSpan>> {
    if bytes.len() > MAX_STREAM_BYTES {
        return Err(Error::UnsafeEdit(format!(
            "external-link stream exceeds the {MAX_STREAM_BYTES}-byte resource bound"
        )));
    }
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        if records.len() >= MAX_RECORDS {
            return Err(Error::UnsafeEdit(
                "external-link stream record count exceeds the resource bound".to_string(),
            ));
        }
        let header_end = offset
            .checked_add(4)
            .ok_or(Error::Allocation("framing external-link records"))?;
        let header = bytes.get(offset..header_end).ok_or_else(|| {
            Error::UnexpectedEndOfStream("truncated BIFF8 external-link record header".to_string())
        })?;
        let record_type = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        if payload_len > MAX_RECORD_PAYLOAD {
            return Err(Error::InvalidRecord {
                record_type,
                message: format!(
                    "BIFF8 payload length {payload_len} exceeds {MAX_RECORD_PAYLOAD} bytes"
                ),
            });
        }
        let payload_start = header_end;
        let payload_end = payload_start
            .checked_add(payload_len)
            .ok_or(Error::Allocation("framing external-link record payload"))?;
        if payload_end > bytes.len() {
            return Err(Error::InvalidRecord {
                record_type,
                message: format!(
                    "record payload of {payload_len} bytes is truncated ({} bytes available)",
                    bytes.len().saturating_sub(payload_start)
                ),
            });
        }
        records.push(RecordSpan {
            record_type,
            record_start: offset,
            payload_start,
            payload_end,
        });
        offset = payload_end;
    }
    Ok(records)
}

pub(super) fn replace_record(
    bytes: &[u8],
    span: RecordSpan,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if replacement.len() > MAX_RECORD_PAYLOAD {
        return Err(Error::UnsafeEdit(format!(
            "replacement BIFF8 payload exceeds {MAX_RECORD_PAYLOAD} bytes"
        )));
    }
    let replacement_len = u16::try_from(replacement.len()).map_err(|_error| {
        Error::UnsafeEdit("replacement BIFF8 payload length exceeds u16".to_string())
    })?;
    let old_end = span.payload_end;
    let mut output = Vec::with_capacity(
        bytes
            .len()
            .saturating_sub(old_end.saturating_sub(span.record_start))
            .saturating_add(4 + replacement.len()),
    );
    output.extend_from_slice(&bytes[..span.record_start]);
    output.extend_from_slice(&span.record_type.to_le_bytes());
    output.extend_from_slice(&replacement_len.to_le_bytes());
    output.extend_from_slice(replacement);
    output.extend_from_slice(&bytes[old_end..]);
    if output.len() > MAX_STREAM_BYTES {
        return Err(Error::UnsafeEdit(format!(
            "edited external-link stream exceeds the {MAX_STREAM_BYTES}-byte resource bound"
        )));
    }
    Ok(output)
}
