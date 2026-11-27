//! Common `PowerPoint` record framing and bounded primitive decoding.

use crate::consts::RecordType;
use crate::package::{Error, Result, SourceLimits};
use crate::records::Record;

pub(crate) fn parse_bool(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => corrupted(format!("ExOleEmbedAtom {field} is not a bool1")),
    }
}

pub(crate) fn u32_at(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

pub(crate) fn require_atom(
    record: &Record,
    version: u16,
    instance: u16,
    kind: RecordType,
    length: usize,
    context: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.record_type_raw != kind.as_u16()
        || record.data.len() != length
        || usize::try_from(record.data_length).ok() != Some(length)
    {
        return corrupted(format!("{context} has an invalid header or size"));
    }
    Ok(())
}

pub(crate) fn record_bytes(
    version: u16,
    instance: u16,
    kind: RecordType,
    data: &[u8],
) -> Result<Vec<u8>> {
    record_bytes_raw(version, instance, kind.as_u16(), data)
}

pub(crate) fn record_bytes_raw(
    version: u16,
    instance: u16,
    kind: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    if version > 0x000f || instance > 0x0fff {
        return corrupted("PowerPoint record header exceeds its encoded domain");
    }
    let length = u32::try_from(data.len())
        .map_err(|_err| Error::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

/// Check the selected source profile before allocating a serialized variable
/// record.  Record headers count toward `max_record_bytes`, while the payload
/// dimension is checked separately so a caller cannot bypass its narrower
/// copied-payload budget by selecting a larger complete-record budget.
pub(crate) fn ensure_record_size(
    payload_len: usize,
    limits: &SourceLimits,
    context: &str,
) -> Result<()> {
    if payload_len > limits.max_record_payload_bytes()
        || payload_len
            .checked_add(8)
            .is_none_or(|size| size > limits.max_record_bytes())
    {
        return corrupted(format!(
            "{context} exceeds the configured record size limit"
        ));
    }
    Ok(())
}

/// Admit an aggregate serialized payload before a container or sequence
/// allocates its next child buffer.  The aggregate ceiling is the same
/// selected copied-payload budget used by record parsing; counting framed
/// child bytes is intentionally conservative because those bytes are also
/// allocated by the writer.
pub(crate) fn ensure_output_total(
    total: usize,
    limits: &SourceLimits,
    context: &str,
) -> Result<()> {
    if total > limits.max_aggregate_payload_bytes() {
        return Err(Error::Limit {
            kind: crate::SourceLimitKind::AggregatePayloadBytes,
            observed: total,
            maximum: limits.max_aggregate_payload_bytes(),
        });
    }
    if total > limits.max_record_payload_bytes() {
        return corrupted(format!(
            "{context} exceeds the configured record payload limit"
        ));
    }
    Ok(())
}

pub(crate) fn record_bytes_with_limits(
    version: u16,
    instance: u16,
    kind: RecordType,
    data: &[u8],
    limits: &SourceLimits,
    context: &str,
) -> Result<Vec<u8>> {
    ensure_record_size(data.len(), limits, context)?;
    record_bytes(version, instance, kind, data)
}

pub(crate) fn record_bytes_raw_with_limits(
    version: u16,
    instance: u16,
    kind: u16,
    data: &[u8],
    limits: &SourceLimits,
    context: &str,
) -> Result<Vec<u8>> {
    ensure_record_size(data.len(), limits, context)?;
    record_bytes_raw(version, instance, kind, data)
}

pub(crate) fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
