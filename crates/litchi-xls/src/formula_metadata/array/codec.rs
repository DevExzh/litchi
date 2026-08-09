//! Strict BIFF8 `Array` payload codec.

use crate::{Error, Result};

use super::super::{Cell, Range};
use super::validation::{FIXED_BYTES, MAX_RECORD_BYTES, invalid};
use super::{Limits, Owner};

/// Parse one complete BIFF8 `Array` record payload.
pub(crate) fn parse_payload(data: &[u8], limits: Limits) -> Result<Owner> {
    if data.len() < FIXED_BYTES + 1 {
        return Err(Error::InvalidLength {
            expected: FIXED_BYTES + 1,
            found: data.len(),
        });
    }
    if data.len() > MAX_RECORD_BYTES || data.len() > limits.max_record_bytes() {
        return Err(invalid(
            "Array payload exceeds the BIFF8 or configured limit",
        ));
    }
    let range = Range::new(
        Cell::new(read_u16(data, 0), data[4]),
        Cell::new(read_u16(data, 2), data[5]),
    )?;
    let flags = read_u16(data, 6);
    let mut unused = [0; 4];
    unused.copy_from_slice(&data[8..12]);
    let token_len = usize::from(read_u16(data, 12));
    let token_end = FIXED_BYTES
        .checked_add(token_len)
        .ok_or_else(|| invalid("ArrayParsedFormula cce overflows"))?;
    if token_end > data.len() {
        return Err(Error::InvalidLength {
            expected: token_end,
            found: data.len(),
        });
    }

    let mut tokens = Vec::new();
    tokens
        .try_reserve_exact(token_len)
        .map_err(|_error| Error::Allocation("retaining ArrayParsedFormula rgce"))?;
    tokens.extend_from_slice(&data[FIXED_BYTES..token_end]);
    let extra_len = data.len() - token_end;
    let mut extra = Vec::new();
    extra
        .try_reserve_exact(extra_len)
        .map_err(|_error| Error::Allocation("retaining ArrayParsedFormula rgcb"))?;
    extra.extend_from_slice(&data[token_end..]);

    let owner = Owner::from_wire(
        range,
        flags & 1 != 0,
        flags & !1,
        unused,
        tokens,
        extra,
        limits.max_cells(),
    );
    super::validation::validate(&owner, limits, false)?;
    Ok(owner)
}

pub(crate) fn write_payload(owner: &Owner) -> Result<Vec<u8>> {
    owner.validate()?;
    let length = FIXED_BYTES
        .checked_add(owner.tokens().len())
        .and_then(|value| value.checked_add(owner.extra().len()))
        .ok_or_else(|| invalid("Array payload length overflows"))?;
    let token_len = u16::try_from(owner.tokens().len())
        .map_err(|_error| invalid("ArrayParsedFormula cce exceeds u16"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|_error| Error::Allocation("serializing Array record"))?;
    let first = owner.range().first();
    let last = owner.range().last();
    output.extend_from_slice(&first.row().to_le_bytes());
    output.extend_from_slice(&last.row().to_le_bytes());
    output.push(first.col());
    output.push(last.col());
    let flags = owner.reserved() | u16::from(owner.always_calculate());
    output.extend_from_slice(&flags.to_le_bytes());
    output.extend_from_slice(&owner.unused());
    output.extend_from_slice(&token_len.to_le_bytes());
    output.extend_from_slice(owner.tokens());
    output.extend_from_slice(owner.extra());
    Ok(output)
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}
