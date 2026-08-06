//! BIFF8 framing and bounded source-preserving replacement helpers.

use crate::{Error, Result};

/// The byte length of a BIFF record header (`rt` plus `cb`), in bytes.
pub(super) const RECORD_HEADER_LEN: usize = 4;
/// The maximum payload carried by one BIFF8 record.
pub(super) const MAX_RECORD_PAYLOAD: usize = 8_224;
/// The maximum source stream retained by the real-time-data owner.
pub(super) const MAX_STREAM_BYTES: usize = 64 * 1024 * 1024;
/// A malformed stream must not create an unbounded record-span table.
const MAX_RECORDS: usize = 1_048_576;

#[derive(Debug, Clone, Copy)]
pub(super) struct RecordSpan {
    pub(super) record_type: u16,
    pub(super) record_start: usize,
    pub(super) payload_start: usize,
    pub(super) payload_end: usize,
}

impl RecordSpan {
    pub(super) fn payload<'a>(self, bytes: &'a [u8]) -> &'a [u8] {
        &bytes[self.payload_start..self.payload_end]
    }
}

/// Scan one complete BIFF8 record stream without interpreting its records.
pub(super) fn scan_records(bytes: &[u8]) -> Result<Vec<RecordSpan>> {
    if bytes.len() > MAX_STREAM_BYTES {
        return Err(Error::UnsafeEdit(format!(
            "RealTimeData stream exceeds the {MAX_STREAM_BYTES}-byte resource bound"
        )));
    }

    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        if records.len() >= MAX_RECORDS {
            return Err(Error::UnsafeEdit(
                "RealTimeData stream record count exceeds the resource bound".to_string(),
            ));
        }
        let header_end = offset
            .checked_add(RECORD_HEADER_LEN)
            .ok_or(Error::Allocation("framing RealTimeData records"))?;
        let header = bytes.get(offset..header_end).ok_or_else(|| {
            Error::UnexpectedEndOfStream(
                "truncated BIFF8 RealTimeData stream record header".to_string(),
            )
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
            .ok_or(Error::Allocation("framing RealTimeData record payload"))?;
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

/// Replace a complete source range while keeping every byte outside the range
/// unchanged.
pub(super) fn replace_range(
    bytes: &[u8],
    start: usize,
    end: usize,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    if start > end || end > bytes.len() {
        return Err(Error::UnsafeEdit(
            "RealTimeData replacement range is outside the source stream".to_string(),
        ));
    }
    let retained = bytes
        .len()
        .checked_sub(end - start)
        .ok_or(Error::Allocation("sizing RealTimeData replacement"))?;
    let output_len = retained
        .checked_add(replacement.len())
        .ok_or(Error::Allocation("sizing RealTimeData replacement"))?;
    if output_len > MAX_STREAM_BYTES {
        return Err(Error::UnsafeEdit(format!(
            "edited RealTimeData stream exceeds the {MAX_STREAM_BYTES}-byte resource bound"
        )));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| Error::Allocation("allocating RealTimeData replacement"))?;
    output.extend_from_slice(&bytes[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&bytes[end..]);
    Ok(output)
}

/// Frame one logical `RealTimeData` payload as a primary record followed by
/// `ContinueFrt` records.
pub(super) fn frame_logical_payload(
    payload: &[u8],
    real_time_data_record_type: u16,
    continue_frt_record_type: u16,
) -> Result<Vec<u8>> {
    if payload.is_empty() {
        return Err(Error::InvalidData(
            "RealTimeData logical payload cannot be empty".to_string(),
        ));
    }
    if payload.len() > crate::real_time_data::codec::MAX_LOGICAL_PAYLOAD_BYTES {
        return Err(Error::UnsafeEdit(
            "RealTimeData logical payload exceeds its resource bound".to_string(),
        ));
    }

    let record_count = payload.len().div_ceil(MAX_RECORD_PAYLOAD);
    let framed_len = payload
        .len()
        .checked_add(
            record_count
                .checked_mul(RECORD_HEADER_LEN)
                .ok_or(Error::Allocation("sizing RealTimeData record framing"))?,
        )
        .ok_or(Error::Allocation("sizing RealTimeData record framing"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(framed_len)
        .map_err(|_| Error::Allocation("allocating RealTimeData record framing"))?;

    for (index, chunk) in payload.chunks(MAX_RECORD_PAYLOAD).enumerate() {
        let record_type = if index == 0 {
            real_time_data_record_type
        } else {
            continue_frt_record_type
        };
        let payload_len = u16::try_from(chunk.len()).map_err(|_| {
            Error::UnsafeEdit("RealTimeData record payload length exceeds u16".to_string())
        })?;
        output.extend_from_slice(&record_type.to_le_bytes());
        output.extend_from_slice(&payload_len.to_le_bytes());
        output.extend_from_slice(chunk);
    }
    Ok(output)
}
