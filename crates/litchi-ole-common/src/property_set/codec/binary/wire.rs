//! Binary property-set wire primitives and bounded readers.

use super::super::super::model::{
    Guid, UNICODE_CODEPAGE, checked_add, invalid, try_clone_string, try_vec_with_capacity,
};
use super::super::support::allocation;
use crate::leniency::{StringDefect, ToleranceReport};
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

/// The reference-compatible repairs one property's parse applied.
///
/// Readers set these while decoding; the section parser records them in the
/// stream's [`ToleranceReport`] against the property identifier, so one
/// property contributes at most one entry per repair class.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Repairs {
    /// A string held malformed MBCS or unpaired UTF-16 (F7.8).
    pub(crate) lossy_decode: bool,
    /// A string was not NUL-terminated within its declared size (F7.10).
    pub(crate) missing_terminator: bool,
}

impl Repairs {
    pub(crate) const fn is_empty(self) -> bool {
        !(self.lossy_decode || self.missing_terminator)
    }

    /// Records every repair class set here against `property_identifier`.
    pub(crate) fn record(self, report: &mut ToleranceReport, property_identifier: u32) {
        if self.lossy_decode {
            report.record(StringDefect::LossyDecode, property_identifier);
        }
        if self.missing_terminator {
            report.record(StringDefect::MissingTerminator, property_identifier);
        }
    }

    /// Refuses any repair, for callers that cannot record one.
    pub(crate) fn require_none(self, description: &str) -> Result<(), OleError> {
        if self.is_empty() {
            Ok(())
        } else {
            Err(invalid(format!(
                "{description} needs a repair this path cannot record"
            )))
        }
    }
}

pub(super) struct ValueReader<'a> {
    data: &'a [u8],
    position: usize,
    alignment_base: usize,
    /// Repairs applied while reading through this reader.
    pub(super) repairs: Repairs,
}

impl<'a> ValueReader<'a> {
    pub(super) const fn new(data: &'a [u8], alignment_base: usize) -> Self {
        Self {
            data,
            position: 0,
            alignment_base,
            repairs: Repairs {
                lossy_decode: false,
                missing_terminator: false,
            },
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
        if size % 2 != 0 {
            return Err(invalid(format!("{description} is not terminated UTF-16LE")));
        }
        let end = utf16_text_end(raw, &mut reader.repairs);
        decode_utf16_lossy(&raw[..end], description, &mut reader.repairs)?
    } else {
        let end = ansi_text_end(raw, &mut reader.repairs);
        decode_ansi_lossy(&raw[..end], codepage, description, &mut reader.repairs)?
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
        let end = utf16_text_end(raw, &mut reader.repairs);
        decode_utf16_lossy(&raw[..end], description, &mut reader.repairs)?
    };
    reader.align4(top_level, &format!("{description} padding"))?;
    Ok(value)
}

/// The byte length of the text in a sized ANSI string: up to its first NUL.
///
/// MS-OLEPS 2.5 requires the declared size to end in a terminator, but the
/// size alone bounds the string and producers disagree on whether it counts
/// the terminator. Like POI (`CodePageString`) the string is read to its
/// declared size and a missing terminator is noted in `repairs` rather than
/// rejected (leniency audit F7.10).
pub(super) fn ansi_text_end(raw: &[u8], repairs: &mut Repairs) -> usize {
    if raw.last() != Some(&0) {
        repairs.missing_terminator = true;
    }
    raw.iter().position(|byte| *byte == 0).unwrap_or(raw.len())
}

/// [`ansi_text_end`] for a sized UTF-16LE string of whole code units.
pub(super) fn utf16_text_end(raw: &[u8], repairs: &mut Repairs) -> usize {
    if !raw.ends_with(&[0, 0]) {
        repairs.missing_terminator = true;
    }
    raw.as_chunks::<2>()
        .0
        .iter()
        .position(|pair| *pair == [0, 0])
        .map_or(raw.len() & !1, |terminator_index| terminator_index * 2)
}

/// Decodes UTF-16LE, substituting U+FFFD for unpaired surrogates (leniency
/// audit F7.8) and noting the substitution in `repairs`.
///
/// Windows' property reader and POI's replacing decoders do the same; one
/// malformed string must not cost the caller every other property.
pub(super) fn decode_utf16_lossy(
    data: &[u8],
    description: &str,
    repairs: &mut Repairs,
) -> Result<String, OleError> {
    if !data.len().is_multiple_of(2) {
        return Err(invalid(format!("{description} has an odd byte length")));
    }
    let units = || {
        data.as_chunks::<2>()
            .0
            .iter()
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
    };
    let mut utf8_len = 0usize;
    let mut lossy = false;
    for decoded in std::char::decode_utf16(units()) {
        let character = decoded.unwrap_or_else(|_unpaired| {
            lossy = true;
            char::REPLACEMENT_CHARACTER
        });
        utf8_len = checked_add(utf8_len, character.len_utf8(), description)?;
    }
    let mut value = String::new();
    value
        .try_reserve_exact(utf8_len)
        .map_err(|source| allocation("decoded UTF-16 string", source))?;
    value.extend(
        std::char::decode_utf16(units())
            .map(|decoded| decoded.unwrap_or(char::REPLACEMENT_CHARACTER)),
    );
    repairs.lossy_decode |= lossy;
    Ok(value)
}

/// Decodes an MBCS string, substituting U+FFFD for malformed sequences
/// (leniency audit F7.8) and noting the substitution in `repairs`.
///
/// `MultiByteToWideChar`, which ole32's property reader uses, substitutes a
/// default character the same way. An unsupported code page stays an error.
pub(super) fn decode_ansi_lossy(
    data: &[u8],
    codepage: u16,
    description: &str,
    repairs: &mut Repairs,
) -> Result<String, OleError> {
    let page = Mbcs::require(u32::from(codepage))
        .map_err(|error| invalid(format!("Could not decode {description}: {error}")))?;
    let (decoded, lossy) = page.recover(data);
    repairs.lossy_decode |= lossy;
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
