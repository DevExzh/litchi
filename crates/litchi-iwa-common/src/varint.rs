//! Bounded Protocol Buffers-style variable-length integers.
//!
//! IWA streams use the same unsigned and zigzag-encoded integer representation
//! as Protocol Buffers for archive framing.  The decoder is deliberately
//! bounded to the ten bytes representable by a `u64`; malformed input fails
//! without probing an unbounded continuation stream.

#![allow(
    clippy::module_name_repetitions,
    reason = "The public names explicitly identify the varint encoding they operate on"
)]

use std::fmt;
use std::io::{self, Read};

/// Maximum number of bytes in a `u64` Protocol Buffers varint.
pub const MAX_BYTES: usize = 10;

/// Failure while decoding an in-memory varint.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeError {
    /// The input ended before a terminating byte was found.
    Truncated,
    /// The encoding cannot be represented as a `u64`.
    Overflow,
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Truncated => formatter.write_str("truncated variable-length integer"),
            Self::Overflow => formatter.write_str("variable-length integer overflow"),
        }
    }
}

impl std::error::Error for DecodeError {}

/// Encode a `u64` as a Protocol Buffers-style variable-length integer.
#[must_use]
pub fn encode_varint(value: u64) -> Vec<u8> {
    let mut output = Vec::with_capacity(encoded_len(value));
    encode_varint_into(&mut output, value);
    output
}

/// Return the number of bytes required for a canonical `u64` varint.
#[must_use]
pub const fn encoded_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

/// Append a canonical `u64` varint without allocating a temporary vector.
pub fn encode_varint_into(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

/// Encode a canonical `u64` varint into caller-owned stack storage.
#[must_use]
pub fn encode_varint_to_buffer(value: u64, buffer: &mut [u8; MAX_BYTES]) -> &[u8] {
    let length = encoded_len(value);
    let mut remaining = value;
    for (index, slot) in buffer.iter_mut().enumerate().take(length) {
        let mut byte = (remaining & 0x7f) as u8;
        remaining >>= 7;
        if index + 1 < length {
            byte |= 0x80;
        }
        *slot = byte;
    }
    &buffer[..length]
}

/// Decode a variable-length integer from a reader.
///
/// # Errors
///
/// Returns an I/O error when the reader is truncated or contains an integer
/// that cannot be represented as a `u64`.
pub fn decode_varint<R: Read>(reader: &mut R) -> io::Result<u64> {
    let mut value = 0u64;
    let mut byte_buffer = [0u8; 1];

    for index in 0..MAX_BYTES {
        reader.read_exact(&mut byte_buffer)?;
        let byte = byte_buffer[0];
        if index == MAX_BYTES - 1 && byte > 1 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                DecodeError::Overflow.to_string(),
            ));
        }

        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok(value);
        }
    }

    Err(io::Error::new(
        io::ErrorKind::InvalidData,
        DecodeError::Overflow.to_string(),
    ))
}

/// Decode a variable-length integer from a byte slice.
///
/// # Errors
///
/// Returns [`DecodeError::Truncated`] when no terminating byte is available,
/// or [`DecodeError::Overflow`] when the representation exceeds `u64`.
pub fn decode_varint_from_bytes(data: &[u8]) -> Result<(u64, usize), DecodeError> {
    let mut value = 0u64;

    for (index, &byte) in data.iter().enumerate().take(MAX_BYTES) {
        if index == MAX_BYTES - 1 && byte > 1 {
            return Err(DecodeError::Overflow);
        }

        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            return Ok((value, index + 1));
        }
    }

    if data.len() >= MAX_BYTES {
        Err(DecodeError::Overflow)
    } else {
        Err(DecodeError::Truncated)
    }
}

/// Encode a signed integer using Protocol Buffers zigzag encoding.
#[must_use]
pub fn encode_svarint(value: i64) -> Vec<u8> {
    // Convert to the unsigned bit pattern before shifting.  Shifting a
    // negative signed value can overflow in debug builds, notably for
    // `i64::MIN`.
    let zigzag = (value.cast_unsigned() << 1) ^ (value >> 63).cast_unsigned();
    encode_varint(zigzag)
}

/// Append a signed zigzag varint without allocating a temporary vector.
pub fn encode_svarint_into(output: &mut Vec<u8>, value: i64) {
    let zigzag = (value.cast_unsigned() << 1) ^ (value >> 63).cast_unsigned();
    encode_varint_into(output, zigzag);
}

/// Encode a signed zigzag varint into caller-owned stack storage.
#[must_use]
pub fn encode_svarint_to_buffer(value: i64, buffer: &mut [u8; MAX_BYTES]) -> &[u8] {
    let zigzag = (value.cast_unsigned() << 1) ^ (value >> 63).cast_unsigned();
    encode_varint_to_buffer(zigzag, buffer)
}

/// Decode a signed integer from Protocol Buffers zigzag encoding.
///
/// # Errors
///
/// Returns an I/O error when the encoded varint is truncated or overflows.
pub fn decode_svarint<R: Read>(reader: &mut R) -> io::Result<i64> {
    let unsigned = decode_varint(reader)?;
    Ok((unsigned >> 1).cast_signed() ^ -((unsigned & 1).cast_signed()))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Varint tests intentionally panic on an unexpected fixture failure"
)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn unsigned_boundaries_round_trip() {
        for value in [0, 1, 127, 128, 300, 16_384, u64::MAX] {
            let encoded = encode_varint(value);
            assert_eq!(
                decode_varint_from_bytes(&encoded),
                Ok((value, encoded.len()))
            );
            assert_eq!(decode_varint(&mut Cursor::new(encoded)).unwrap(), value);
        }
    }

    #[test]
    fn append_encoding_matches_owned_encoding() {
        for value in [0, 1, 127, 128, 16_384, u64::MAX] {
            let mut output = vec![0xaa];
            encode_varint_into(&mut output, value);
            assert_eq!(&output[1..], encode_varint(value));
            assert_eq!(encoded_len(value), output.len() - 1);
        }
        let mut output = Vec::new();
        encode_svarint_into(&mut output, i64::MIN);
        assert_eq!(
            decode_svarint(&mut Cursor::new(output)).ok(),
            Some(i64::MIN)
        );

        let mut buffer = [0u8; MAX_BYTES];
        let encoded_max = encode_varint(u64::MAX);
        assert_eq!(encode_varint_to_buffer(u64::MAX, &mut buffer), encoded_max);
        assert_eq!(encode_svarint_to_buffer(i64::MIN, &mut buffer), encoded_max);
    }

    #[test]
    fn signed_boundaries_round_trip_without_signed_shift_overflow() {
        for value in [i64::MIN, -64, -2, -1, 0, 1, 63, 64, i64::MAX] {
            let encoded = encode_svarint(value);
            assert_eq!(decode_svarint(&mut Cursor::new(encoded)).unwrap(), value);
        }
    }

    #[test]
    fn rejects_truncated_and_overflowing_slice_inputs() {
        assert_eq!(decode_varint_from_bytes(&[]), Err(DecodeError::Truncated));
        assert_eq!(
            decode_varint_from_bytes(&[0x80]),
            Err(DecodeError::Truncated)
        );
        assert_eq!(
            decode_varint_from_bytes(&[0xff; 10]),
            Err(DecodeError::Overflow)
        );
        assert_eq!(
            decode_varint_from_bytes(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 2]),
            Err(DecodeError::Overflow)
        );
        assert_eq!(
            decode_varint_from_bytes(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x80]),
            Err(DecodeError::Overflow)
        );
        assert_eq!(
            decode_varint_from_bytes(&[0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 1]),
            Ok((u64::MAX, 10))
        );
        assert_eq!(decode_varint_from_bytes(&[0x80, 0x00]), Ok((0, 2)));
        assert_eq!(decode_varint_from_bytes(&[1, 2]), Ok((1, 1)));
    }

    #[test]
    fn reader_rejects_tenth_byte_values_above_one() {
        let error = decode_varint(&mut Cursor::new([0xff; 10])).unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);
        assert_eq!(error.to_string(), DecodeError::Overflow.to_string());
    }
}
