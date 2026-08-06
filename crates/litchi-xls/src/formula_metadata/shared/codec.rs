//! BIFF8 `ShrFmla` payload codec.

use crate::{Error, Result};

use super::model::{Cell, Range};
use super::validation::{FIXED_PAYLOAD_SIZE, MAX_RECORD_PAYLOAD, invalid};

/// Parsed fields of a ShrFmla record before worksheet Formula cells supply
/// the anchor and participating-cell consistency context.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Record {
    pub range: Range,
    pub reserved: u8,
    pub count: u8,
    pub tokens: Vec<u8>,
}

/// Parse one complete ShrFmla payload.
pub fn parse(data: &[u8]) -> Result<Record> {
    if data.len() < FIXED_PAYLOAD_SIZE {
        return Err(Error::InvalidLength {
            expected: FIXED_PAYLOAD_SIZE,
            found: data.len(),
        });
    }
    if data.len() > MAX_RECORD_PAYLOAD {
        return Err(invalid(format!(
            "ShrFmla payload exceeds the BIFF8 limit of {MAX_RECORD_PAYLOAD} bytes"
        )));
    }

    let range = Range::new(
        Cell::new(read_u16(data, 0), data[4]),
        Cell::new(read_u16(data, 2), data[5]),
    )?;
    let reserved = data[6];
    if reserved != 0 {
        return Err(invalid("ShrFmla reserved byte must be zero"));
    }
    let count = data[7];
    if count == 0 {
        return Err(invalid("ShrFmla cUse must include its anchor cell"));
    }
    let token_len = usize::from(read_u16(data, 8));
    let end = FIXED_PAYLOAD_SIZE
        .checked_add(token_len)
        .ok_or_else(|| invalid("ShrFmla token length overflows"))?;
    if end != data.len() {
        return Err(Error::InvalidLength {
            expected: end,
            found: data.len(),
        });
    }
    if token_len == 0 {
        return Err(invalid("ShrFmla shared parsed formula cannot be empty"));
    }
    if data[FIXED_PAYLOAD_SIZE..]
        .first()
        .is_some_and(|opcode| opcode & 0x7F == 0x01)
    {
        return Err(Error::UnsupportedFeature(
            "ShrFmla shared parsed formula cannot begin with PtgExp".to_string(),
        ));
    }

    Ok(Record {
        range,
        reserved,
        count,
        tokens: data[FIXED_PAYLOAD_SIZE..].to_vec(),
    })
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}
