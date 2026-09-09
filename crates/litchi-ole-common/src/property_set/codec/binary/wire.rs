//! Binary property-set wire primitives and bounded readers.

use super::super::super::model::{
    Guid, UNICODE_CODEPAGE, checked_add, invalid, try_clone_string, try_vec_with_capacity,
};
use super::super::support::allocation;
use litchi_cfb::OleError;
use litchi_codepage::Mbcs;
use std::borrow::Cow;

/// A fallible byte sink shared by the canonical and bounded encoders.
///
/// Keeping the append operations behind one sink makes the bounded path run
/// the same value encoder as the ordinary path.  The bounded sink checks the
/// caller's ceiling before reserving or copying bytes; the counting sink is
/// used for the bounded preflight and never allocates payload storage.
pub(super) trait ByteSink {
    fn len(&self) -> usize;

    fn reserve(&mut self, additional: usize, resource: &'static str) -> Result<(), OleError>;

    fn append_reserved(&mut self, bytes: &[u8]);

    fn push_zero(&mut self, resource: &'static str) -> Result<(), OleError> {
        self.reserve(1, resource)?;
        self.append_reserved(&[0]);
        Ok(())
    }

    fn append_zeroes(&mut self, length: usize, resource: &'static str) -> Result<(), OleError> {
        self.reserve(length, resource)?;
        // The default implementation is intentionally small and is used only
        // for padding.  Larger payloads go through append_bytes below.
        for _ in 0..length {
            self.append_reserved(&[0]);
        }
        Ok(())
    }
}

/// A serialized byte vector with an optional exact retained-byte ceiling.
pub(super) struct ByteWriter {
    bytes: Vec<u8>,
    maximum: Option<u64>,
}

impl ByteWriter {
    pub(super) fn new(
        initial_capacity: usize,
        maximum: Option<u64>,
        resource: &'static str,
    ) -> Result<Self, OleError> {
        let capacity = maximum
            .and_then(|maximum| usize::try_from(maximum).ok())
            .map_or(initial_capacity, |maximum| initial_capacity.min(maximum));
        Ok(Self {
            bytes: try_vec_with_capacity(capacity, resource)?,
            maximum,
        })
    }

    pub(super) fn zeroed(
        length: usize,
        maximum: Option<u64>,
        resource: &'static str,
    ) -> Result<Self, OleError> {
        if let Some(maximum) = maximum {
            let observed = u64::try_from(length).unwrap_or(u64::MAX);
            if observed > maximum {
                return Err(OleError::LimitExceeded {
                    resource,
                    observed,
                    maximum,
                });
            }
        }
        let mut output = Self::new(length, maximum, resource)?;
        output.bytes.resize(length, 0);
        Ok(output)
    }

    pub(super) fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    pub(super) fn as_mut_slice(&mut self) -> &mut [u8] {
        &mut self.bytes
    }
}

impl ByteSink for ByteWriter {
    fn len(&self) -> usize {
        self.bytes.len()
    }

    fn reserve(&mut self, additional: usize, resource: &'static str) -> Result<(), OleError> {
        let required = self
            .bytes
            .len()
            .checked_add(additional)
            .ok_or_else(|| invalid("serialized property size overflow"))?;
        if let Some(maximum) = self.maximum {
            let observed = u64::try_from(required).unwrap_or(u64::MAX);
            if observed > maximum {
                return Err(OleError::LimitExceeded {
                    resource,
                    observed,
                    maximum,
                });
            }
        }
        self.bytes
            .try_reserve_exact(additional)
            .map_err(|source| allocation(resource, source))?;
        Ok(())
    }

    fn append_reserved(&mut self, bytes: &[u8]) {
        self.bytes.extend_from_slice(bytes);
    }
}

/// Allocation-free sink used to determine the exact canonical output length.
pub(super) struct CountingWriter {
    length: usize,
}

impl CountingWriter {
    pub(super) const fn new() -> Self {
        Self { length: 0 }
    }
}

impl ByteSink for CountingWriter {
    fn len(&self) -> usize {
        self.length
    }

    fn reserve(&mut self, additional: usize, _resource: &'static str) -> Result<(), OleError> {
        self.length
            .checked_add(additional)
            .ok_or_else(|| invalid("serialized property size overflow"))?;
        Ok(())
    }

    fn append_reserved(&mut self, bytes: &[u8]) {
        // `reserve` is called by every shared append helper before this
        // method.  This second checked addition keeps the sink correct if a
        // future caller appends directly.
        self.length = self.length.saturating_add(bytes.len());
    }
}

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

pub(super) fn reserve_bytes<S: ByteSink + ?Sized>(
    output: &mut S,
    additional: usize,
    resource: &'static str,
) -> Result<(), OleError> {
    output.reserve(additional, resource)
}

pub(super) fn append_bytes<S: ByteSink + ?Sized>(
    output: &mut S,
    bytes: &[u8],
    resource: &'static str,
) -> Result<(), OleError> {
    reserve_bytes(output, bytes.len(), resource)?;
    output.append_reserved(bytes);
    Ok(())
}

pub(super) fn append_u16<S: ByteSink + ?Sized>(
    output: &mut S,
    value: u16,
    resource: &'static str,
) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

pub(super) fn append_u32<S: ByteSink + ?Sized>(
    output: &mut S,
    value: u32,
    resource: &'static str,
) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

pub(super) fn append_u64<S: ByteSink + ?Sized>(
    output: &mut S,
    value: u64,
    resource: &'static str,
) -> Result<(), OleError> {
    append_bytes(output, &value.to_le_bytes(), resource)
}

pub(super) fn pad4<S: ByteSink + ?Sized>(out: &mut S) -> Result<(), OleError> {
    let padding = (4 - (out.len() & 3)) & 3;
    out.append_zeroes(padding, "serialized property padding")
}

/// Return the exact encoded byte count without retaining the complete result.
///
/// `litchi-codepage` exposes a strict whole-string conversion.  Encoding one
/// Unicode scalar at a time keeps the bounded preflight's scratch storage
/// constant while preserving the same strict conversion errors for all of the
/// supported state-free MBCS pages.
pub(super) fn encoded_ansi_len(value: &str, codepage: u16) -> Result<usize, OleError> {
    let page = Mbcs::require(u32::from(codepage)).map_err(|error| invalid(error.to_string()))?;
    let mut length = 0usize;
    for character in value.chars() {
        let mut utf8 = [0u8; 4];
        let fragment = character.encode_utf8(&mut utf8);
        let bytes = page
            .encode(fragment)
            .map_err(|error| invalid(error.to_string()))?;
        length = length
            .checked_add(bytes.len())
            .ok_or_else(|| invalid("encoded ANSI string length overflow"))?;
    }
    Ok(length)
}

/// Append a strictly encoded MBCS string without creating a whole-string
/// temporary buffer.
pub(super) fn append_ansi<S: ByteSink + ?Sized>(
    output: &mut S,
    value: &str,
    codepage: u16,
    resource: &'static str,
) -> Result<(), OleError> {
    let page = Mbcs::require(u32::from(codepage)).map_err(|error| invalid(error.to_string()))?;
    let encoded_length = encoded_ansi_len(value, codepage)?;
    output.reserve(encoded_length, resource)?;
    for character in value.chars() {
        let mut utf8 = [0u8; 4];
        let fragment = character.encode_utf8(&mut utf8);
        let bytes = page
            .encode(fragment)
            .map_err(|error| invalid(error.to_string()))?;
        append_bytes(output, &bytes, resource)?;
    }
    Ok(())
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
            .as_chunks::<2>()
            .0
            .iter()
            .position(|pair| *pair == [0, 0])
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
            .as_chunks::<2>()
            .0
            .iter()
            .position(|pair| *pair == [0, 0])
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
        data.as_chunks::<2>()
            .0
            .iter()
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
        data.as_chunks::<2>()
            .0
            .iter()
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
