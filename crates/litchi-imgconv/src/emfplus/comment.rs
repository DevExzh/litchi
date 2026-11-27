use litchi_core::error::Result;

use super::types::{
    EMFPLUS_COMMENT_IDENTIFIER, EMFPLUS_RECORD_HEADER_SIZE, EMR_COMMENT, ParserLimits, parse_error,
};

const EMR_HEADER_SIZE: usize = 8;
const COMMENT_DATA_SIZE_FIELD: usize = 4;
const COMMENT_IDENTIFIER_SIZE: usize = 4;
const COMMENT_MIN_RECORD_SIZE: usize = EMR_HEADER_SIZE + COMMENT_DATA_SIZE_FIELD;

/// Extract EMF+ bytes from one complete `EMR_COMMENT_EMFPLUS` record.
///
/// `record` must contain exactly one EMF record, including its 8-byte EMF
/// header. Use [`extract_emfplus_comment_body`] with `EmfRecordRef::data`,
/// whose slice starts at the `EMR_COMMENT` `DataSize` field.
pub fn extract_emfplus_comment_record(record: &[u8], limits: ParserLimits) -> Result<&[u8]> {
    try_extract_emfplus_comment_record(record, limits)?
        .ok_or_else(|| parse_error("record is not an EMR_COMMENT_EMFPLUS record"))
}

/// Like [`extract_emfplus_comment_record`], but return `None` for a different
/// EMF record type or comment identifier.
pub fn try_extract_emfplus_comment_record(
    record: &[u8],
    limits: ParserLimits,
) -> Result<Option<&[u8]>> {
    limits.validate()?;
    if record.len() < EMR_HEADER_SIZE {
        return Err(parse_error("truncated EMF record header"));
    }

    let record_type = read_u32(record, 0)?;
    if record_type != EMR_COMMENT {
        return Ok(None);
    }

    let declared_size = usize::try_from(read_u32(record, 4)?)
        .map_err(|_| parse_error("EMR_COMMENT size does not fit usize"))?;
    if declared_size < COMMENT_MIN_RECORD_SIZE {
        return Err(parse_error(format!(
            "EMR_COMMENT size {declared_size} is smaller than {COMMENT_MIN_RECORD_SIZE}"
        )));
    }
    if declared_size % 4 != 0 {
        return Err(parse_error("EMR_COMMENT size is not 32-bit aligned"));
    }
    if declared_size != record.len() {
        return Err(parse_error(format!(
            "EMR_COMMENT declares {declared_size} bytes but slice has {}",
            record.len()
        )));
    }

    try_extract_emfplus_comment_body(&record[EMR_HEADER_SIZE..], limits)
}

/// Extract EMF+ bytes from the body of one `EMR_COMMENT_EMFPLUS` record.
///
/// `body` starts with the 4-byte `EMR_COMMENT` `DataSize` field, followed by the
/// comment identifier, payload, and optional alignment padding. This is the
/// layout exposed as `EmfRecordRef::data` by the existing EMF parser.
pub fn extract_emfplus_comment_body(body: &[u8], limits: ParserLimits) -> Result<&[u8]> {
    try_extract_emfplus_comment_body(body, limits)?
        .ok_or_else(|| parse_error("comment signature is not the EMF+ identifier"))
}

/// Like [`extract_emfplus_comment_body`], but return `None` for another valid
/// comment identifier.
pub fn try_extract_emfplus_comment_body(
    body: &[u8],
    limits: ParserLimits,
) -> Result<Option<&[u8]>> {
    limits.validate()?;
    if body.len() < COMMENT_DATA_SIZE_FIELD {
        return Err(parse_error("truncated EMR_COMMENT body"));
    }

    let data_size = usize::try_from(read_u32(body, 0)?)
        .map_err(|_| parse_error("EMR_COMMENT DataSize does not fit usize"))?;
    let unpadded_record_size = EMR_HEADER_SIZE
        .checked_add(COMMENT_DATA_SIZE_FIELD)
        .and_then(|value| value.checked_add(data_size))
        .ok_or_else(|| parse_error("EMR_COMMENT DataSize arithmetic overflow"))?;
    let padding = (4 - (unpadded_record_size % 4)) % 4;
    let expected_body_size = COMMENT_DATA_SIZE_FIELD
        .checked_add(data_size)
        .and_then(|value| value.checked_add(padding))
        .ok_or_else(|| parse_error("EMR_COMMENT body size arithmetic overflow"))?;
    if body.len() != expected_body_size {
        return Err(parse_error(format!(
            "EMR_COMMENT body requires {expected_body_size} bytes but slice has {}",
            body.len()
        )));
    }

    // Ordinary comments may contain fewer than four data bytes and therefore
    // cannot carry an EMF+ identifier. They are valid non-EMF+ comments.
    if data_size < COMMENT_IDENTIFIER_SIZE {
        return Ok(None);
    }

    let identifier = read_u32(body, COMMENT_DATA_SIZE_FIELD)?;
    if identifier != EMFPLUS_COMMENT_IDENTIFIER {
        return Ok(None);
    }

    let payload_start = COMMENT_DATA_SIZE_FIELD + COMMENT_IDENTIFIER_SIZE;
    let payload_end = COMMENT_DATA_SIZE_FIELD
        .checked_add(data_size)
        .ok_or_else(|| parse_error("EMF+ payload end overflow"))?;
    let payload = body
        .get(payload_start..payload_end)
        .ok_or_else(|| parse_error("truncated EMF+ comment payload"))?;
    if payload.len() < EMFPLUS_RECORD_HEADER_SIZE {
        return Err(parse_error(
            "EMR_COMMENT_EMFPLUS must contain at least one complete record header",
        ));
    }
    if payload.len() > limits.max_bytes {
        return Err(parse_error(format!(
            "EMF+ payload has {} bytes, exceeding limit {}",
            payload.len(),
            limits.max_bytes
        )));
    }

    Ok(Some(payload))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32> {
    let end = offset
        .checked_add(4)
        .ok_or_else(|| parse_error("32-bit field offset overflow"))?;
    let bytes: [u8; 4] = data
        .get(offset..end)
        .ok_or_else(|| parse_error("truncated 32-bit field"))?
        .try_into()
        .map_err(|_| parse_error("invalid 32-bit field"))?;
    Ok(u32::from_le_bytes(bytes))
}
