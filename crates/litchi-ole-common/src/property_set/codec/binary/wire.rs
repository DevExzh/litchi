//! Binary property-set wire primitives and bounded readers.

use super::super::super::model::{
    Guid, UNICODE_CODEPAGE, checked_add, invalid, try_clone_string, try_vec_with_capacity,
};
use super::super::support::allocation;
use litchi_cfb::OleError;
use litchi_codepage::Mbcs;
use std::borrow::Cow;

pub(super) struct ValueReader<'a> {
    data: &'a [u8],
    position: usize,
    alignment_base: usize,
}

impl<'a> ValueReader<'a> {
    pub(super) const fn new(data: &'a [u8], alignment_base: usize) -> Self {
        Self {
            data,
            position: 0,
            alignment_base,
        }
    }

    pub(super) fn remaining_len(&self) -> usize {
        self.data.len() - self.position
    }

    pub(super) fn take(&mut self, length: usize, description: &str) -> Result<&'a [u8], OleError> {
        let start = self.position;
        let end = start
            .checked_add(length)
            .filter(|end| *end <= self.data.len())
            .ok_or_else(|| invalid(format!("{description} exceeds its property range")))?;
        self.position = end;
        Ok(&self.data[start..end])
    }

    pub(super) fn take_remaining(&mut self) -> &'a [u8] {
        let remaining = &self.data[self.position..];
        self.position = self.data.len();
        remaining
    }

    pub(super) fn read_u8(&mut self, description: &str) -> Result<u8, OleError> {
        Ok(self.take(1, description)?[0])
    }

    pub(super) fn read_i8(&mut self, description: &str) -> Result<i8, OleError> {
        Ok(i8::from_ne_bytes([self.read_u8(description)?]))
    }

    pub(super) fn read_u16(&mut self, description: &str) -> Result<u16, OleError> {
        let bytes = self.take(2, description)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    pub(super) fn read_i16(&mut self, description: &str) -> Result<i16, OleError> {
        let bytes = self.take(2, description)?;
        Ok(i16::from_le_bytes([bytes[0], bytes[1]]))
    }

    pub(super) fn read_u32(&mut self, description: &str) -> Result<u32, OleError> {
        let bytes = self.take(4, description)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    pub(super) fn read_i32(&mut self, description: &str) -> Result<i32, OleError> {
        let bytes = self.take(4, description)?;
        Ok(i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    pub(super) fn read_u64(&mut self, description: &str) -> Result<u64, OleError> {
        let bytes = self.take(8, description)?;
        Ok(u64::from_le_bytes([
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
        ]))
    }

    pub(super) fn read_i64(&mut self, description: &str) -> Result<i64, OleError> {
        let bytes = self.take(8, description)?;
        Ok(i64::from_le_bytes([
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
        ]))
    }

    pub(super) fn align4(&mut self, top_level: bool, description: &str) -> Result<(), OleError> {
        let absolute_position = self
            .alignment_base
            .checked_add(self.position)
            .ok_or_else(|| invalid(format!("{description} position overflow")))?;
        let padding = (4 - (absolute_position & 3)) & 3;
        let available = padding.min(self.remaining_len());
        let end = self
            .position
            .checked_add(available)
            .ok_or_else(|| invalid(format!("{description} range overflow")))?;
        let candidate = &self.data[self.position..end];
        let consumed = if top_level {
            // Top-level property offsets are authoritative. Several Office
            // producers omit filler or write nonzero filler between values.
            available
        } else {
            // Inside a vector there is no offset table. Match Office readers:
            // skip zero filler only and stop before the next nonzero field.
            candidate.iter().take_while(|byte| **byte == 0).count()
        };
        self.take(consumed, description)?;
        Ok(())
    }

    pub(super) fn finish_zero_padding(&mut self, description: &str) -> Result<(), OleError> {
        let remaining = &self.data[self.position..];
        if remaining.iter().any(|byte| *byte != 0) {
            return Err(invalid(format!("{description} must be zero")));
        }
        self.position = self.data.len();
        Ok(())
    }
}

pub(super) fn try_zeroed_vec(len: usize, resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut values = try_vec_with_capacity(len, resource)?;
    values.resize(len, 0);
    Ok(values)
}

pub(super) fn reserve_bytes(
    output: &mut Vec<u8>,
    additional: usize,
    resource: &'static str,
) -> Result<(), OleError> {
    output
        .len()
        .checked_add(additional)
        .ok_or_else(|| invalid("serialized property size overflow"))?;
    output
        .try_reserve(additional)
        .map_err(|source| allocation(resource, source))?;
    Ok(())
}

pub(super) fn append_bytes(
    output: &mut Vec<u8>,
    bytes: &[u8],
    resource: &'static str,
) -> Result<(), OleError> {
    reserve_bytes(output, bytes.len(), resource)?;
    output.extend_from_slice(bytes);
    Ok(())
}

pub(super) fn append_u16(
    output: &mut Vec<u8>,
    value: u16,
    resource: &'static str,
) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

pub(super) fn append_u32(
    output: &mut Vec<u8>,
    value: u32,
    resource: &'static str,
) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

pub(super) fn append_u64(
    output: &mut Vec<u8>,
    value: u64,
    resource: &'static str,
) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

pub(super) fn pad4(out: &mut Vec<u8>) -> Result<(), OleError> {
    let padding = (4 - (out.len() & 3)) & 3;
    reserve_bytes(out, padding, "serialized property padding")?;
    for _ in 0..padding {
        out.push(0);
    }
    Ok(())
}

pub(super) fn encode_ansi(value: &str, codepage: u16) -> Result<Vec<u8>, OleError> {
    let page = Mbcs::require(u32::from(codepage)).map_err(|error| invalid(error.to_string()))?;
    page.encode(value)
        .map(Cow::into_owned)
        .map_err(|error| invalid(error.to_string()))
}

pub(super) fn read_codepage_string(
    reader: &mut ValueReader<'_>,
    codepage: u16,
    description: &str,
    top_level: bool,
) -> Result<String, OleError> {
    let size = usize::try_from(reader.read_u32(description)?)
        .map_err(|_conversion_error| invalid(format!("{description} is too large")))?;
    let raw = reader.take(size, description)?;
    let value = if size == 0 {
        String::new()
    } else if codepage == UNICODE_CODEPAGE {
        if size % 2 != 0 || !raw.ends_with(&[0, 0]) {
            return Err(invalid(format!("{description} is not terminated UTF-16LE")));
        }
        let end = raw
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .map_or(raw.len(), |terminator_index| terminator_index * 2);
        decode_utf16(&raw[..end], description)?
    } else {
        if raw.last() != Some(&0) {
            return Err(invalid(format!("{description} is not NUL-terminated")));
        }
        let end = raw.iter().position(|byte| *byte == 0).unwrap_or(raw.len());
        decode_ansi(&raw[..end], codepage, description)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

pub(super) fn read_unicode_string(
    reader: &mut ValueReader<'_>,
    description: &str,
    top_level: bool,
) -> Result<String, OleError> {
    let units = usize::try_from(reader.read_u32(description)?)
        .map_err(|_conversion_error| invalid(format!("{description} is too large")))?;
    let byte_len = units
        .checked_mul(2)
        .ok_or_else(|| invalid(format!("{description} length overflow")))?;
    let raw = reader.take(byte_len, description)?;
    let value = if units == 0 {
        String::new()
    } else {
        if !raw.ends_with(&[0, 0]) {
            return Err(invalid(format!("{description} is not NUL-terminated")));
        }
        let end = raw
            .chunks_exact(2)
            .position(|pair| pair == [0, 0])
            .map_or(raw.len(), |terminator_index| terminator_index * 2);
        decode_utf16(&raw[..end], description)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

pub(super) fn decode_utf16(data: &[u8], description: &str) -> Result<String, OleError> {
    if !data.len().is_multiple_of(2) {
        return Err(invalid(format!("{description} has an odd byte length")));
    }
    let mut utf8_len = 0usize;
    for decoded in std::char::decode_utf16(
        data.chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
    ) {
        let character = decoded
            .map_err(|_utf16_error| invalid(format!("{description} contains invalid UTF-16")))?;
        utf8_len = checked_add(utf8_len, character.len_utf8(), description)?;
    }
    let mut value = String::new();
    value
        .try_reserve_exact(utf8_len)
        .map_err(|source| allocation("decoded UTF-16 string", source))?;
    for decoded in std::char::decode_utf16(
        data.chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
    ) {
        let character = decoded
            .map_err(|_utf16_error| invalid(format!("{description} contains invalid UTF-16")))?;
        value.push(character);
    }
    Ok(value)
}

pub(super) fn decode_ansi(
    data: &[u8],
    codepage: u16,
    description: &str,
) -> Result<String, OleError> {
    let page = Mbcs::require(u32::from(codepage))
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))?;
    let decoded = page
        .decode(data)
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))?;
    match decoded {
        Cow::Borrowed(value) => try_clone_string(value, "decoded ANSI string"),
        Cow::Owned(value) => Ok(value),
    }
}

pub(super) fn checked_range<'a>(
    data: &'a [u8],
    offset: usize,
    length: usize,
    description: &str,
) -> Result<&'a [u8], OleError> {
    let end = offset
        .checked_add(length)
        .filter(|end| *end <= data.len())
        .ok_or_else(|| invalid(format!("{description} exceeds its enclosing range")))?;
    Ok(&data[offset..end])
}

pub(super) fn read_u16(data: &[u8], offset: usize, description: &str) -> Result<u16, OleError> {
    let bytes = checked_range(data, offset, 2, description)?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

pub(super) fn read_u32(data: &[u8], offset: usize, description: &str) -> Result<u32, OleError> {
    let bytes = checked_range(data, offset, 4, description)?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

pub(super) fn read_guid(data: &[u8], offset: usize, description: &str) -> Result<Guid, OleError> {
    let bytes = checked_range(data, offset, 16, description)?;
    let mut guid = [0u8; 16];
    guid.copy_from_slice(bytes);
    Ok(Guid::from_bytes(guid))
}
