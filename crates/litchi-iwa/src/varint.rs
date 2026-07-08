//! Variable-length integer encoding/decoding for IWA format
//!
//! iWork IWA files use Protocol Buffers variable-length encoding
//! for integers, which encodes values in 7-bit chunks with the
//! most significant bit indicating continuation.

use std::io::{self, Read};

/// Encode a u64 value as a variable-length integer
pub fn encode_varint(mut value: u64) -> Vec<u8> {
    let mut buf = Vec::new();
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        buf.push(byte);
        if value == 0 {
            break;
        }
    }
    buf
}

/// Decode a variable-length integer from a reader
pub fn decode_varint<R: Read>(reader: &mut R) -> io::Result<u64> {
    let mut value: u64 = 0;
    let mut shift = 0;
    let mut buf = [0u8; 1];

    loop {
        reader.read_exact(&mut buf)?;
        let byte = buf[0];

        // Check for overflow
        if shift >= 64 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "Variable-length integer overflow",
            ));
        }

        value |= ((byte & 0x7F) as u64) << shift;

        if (byte & 0x80) == 0 {
            break;
        }

        shift += 7;
    }

    Ok(value)
}

/// Decode a variable-length integer from a byte slice
pub fn decode_varint_from_bytes(data: &[u8]) -> Result<(u64, usize), &'static str> {
    let mut value: u64 = 0;
    let mut shift = 0;
    let mut consumed = 0;

    for &byte in data {
        // Check for overflow
        if shift >= 64 {
            return Err("Variable-length integer overflow");
        }

        value |= ((byte & 0x7F) as u64) << shift;
        consumed += 1;

        if (byte & 0x80) == 0 {
            break;
        }

        shift += 7;
    }

    Ok((value, consumed))
}

/// Encode a signed integer using zigzag encoding (as used in Protocol Buffers)
pub fn encode_svarint(value: i64) -> Vec<u8> {
    let zigzag = ((value << 1) ^ (value >> 63)) as u64;
    encode_varint(zigzag)
}

/// Decode a signed integer from zigzag encoding
pub fn decode_svarint<R: Read>(reader: &mut R) -> io::Result<i64> {
    let unsigned = decode_varint(reader)?;
    // Zigzag decoding: (unsigned >> 1) ^ -(unsigned & 1)
    let decoded = (unsigned >> 1) as i64 ^ -((unsigned & 1) as i64);
    Ok(decoded)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_varint_encoding() {
        // Test various values
        let test_cases = vec![
            (0u64, vec![0x00]),
            (1u64, vec![0x01]),
            (127u64, vec![0x7F]),
            (128u64, vec![0x80, 0x01]),
            (300u64, vec![0xAC, 0x02]),
            (16384u64, vec![0x80, 0x80, 0x01]),
        ];

        for (value, expected) in test_cases {
            let encoded = encode_varint(value);
            assert_eq!(encoded, expected, "Encoding failed for value {}", value);

            let (decoded, consumed) = decode_varint_from_bytes(&encoded).expect("Decoding failed");
            assert_eq!(decoded, value, "Decoding failed for value {}", value);
            assert_eq!(
                consumed,
                encoded.len(),
                "Wrong consumed bytes for value {}",
                value
            );
        }
    }

    #[test]
    fn test_svarint_encoding() {
        // Test signed integers
        let test_cases = vec![
            (0i64, vec![0x00]),
            (-1i64, vec![0x01]),
            (1i64, vec![0x02]),
            (-2i64, vec![0x03]),
            (63i64, vec![0x7E]),
            (-64i64, vec![0x7F]),
            (64i64, vec![0x80, 0x01]),
        ];

        for (value, expected) in test_cases {
            let encoded = encode_svarint(value);
            assert_eq!(encoded, expected, "Encoding failed for value {}", value);

            // Decode using zigzag decoding
            let unsigned = decode_varint_from_bytes(&encoded)
                .expect("Decoding failed")
                .0;
            let decoded = (unsigned >> 1) as i64 ^ -((unsigned & 1) as i64);
            assert_eq!(decoded, value, "Decoding failed for value {}", value);
        }
    }
}
