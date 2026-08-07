//! MS-OVBA compressed-container encoding and decoding.

use super::{Error, Limits, check_limit, invalid};
use std::collections::HashMap;

const CONTAINER_SIGNATURE: u8 = 0x01;
const CHUNK_SIGNATURE_MASK: u16 = 0x7000;
const CHUNK_SIGNATURE: u16 = 0x3000;
const COMPRESSED_CHUNK_FLAG: u16 = 0x8000;
const CHUNK_SIZE_MASK: u16 = 0x0fff;
const CHUNK_HEADER_BYTES: usize = 2;
const RAW_CHUNK_BYTES: usize = 4096;
const MIN_COPY_LENGTH_BITS: u32 = 4;
const MIN_COPY_BYTES: usize = 3;

/// Compress bytes into one complete MS-OVBA `CompressedContainer`.
///
/// The encoder follows the greedy matching algorithm in [MS-OVBA] section
/// 2.4.1.3.19.4 and emits deterministic output. Both the source and encoded
/// sizes are checked against [`Limits`].
///
/// A raw final chunk is padded with zero bytes to 4096 bytes, as required by
/// [MS-OVBA]. Consequently, decompressing an incompressible final partial
/// chunk can produce trailing zero bytes beyond the original input.
///
/// # Errors
///
/// Returns an error if `decompressed` or the encoded container exceeds the
/// configured [`Limits`], or if an internal size conversion overflows.
pub fn encode(decompressed: &[u8], limits: &Limits) -> Result<Vec<u8>, Error> {
    check_limit(
        "decompressed VBA stream bytes",
        decompressed.len(),
        limits.max_decompressed_stream_bytes,
    )?;

    let mut output = Vec::with_capacity(
        decompressed
            .len()
            .saturating_add(1)
            .min(limits.max_compressed_stream_bytes),
    );
    append_compressed_checked(&mut output, &[CONTAINER_SIGNATURE], limits)?;
    for chunk in decompressed.chunks(RAW_CHUNK_BYTES) {
        let encoded = compress_chunk(chunk)?;
        append_compressed_checked(&mut output, &encoded, limits)?;
    }
    Ok(output)
}

fn compress_chunk(chunk: &[u8]) -> Result<Vec<u8>, Error> {
    if chunk.is_empty() || chunk.len() > RAW_CHUNK_BYTES {
        return Err(invalid(
            "compression chunk size must be between 1 and 4096 bytes",
        ));
    }

    let mut data = Vec::with_capacity(RAW_CHUNK_BYTES);
    let mut positions = HashMap::<[u8; MIN_COPY_BYTES], Vec<u16>>::new();
    let mut current = 0usize;

    while current < chunk.len() && data.len() < RAW_CHUNK_BYTES {
        let flags_index = data.len();
        data.push(0);
        let mut flags = 0u8;

        for token_index in 0..8u32 {
            if current >= chunk.len() || data.len() >= RAW_CHUNK_BYTES {
                break;
            }

            let previous = current;
            if let Some((length, offset)) = find_match(chunk, current, &positions) {
                if data.len() + 2 > RAW_CHUNK_BYTES {
                    return Ok(raw_chunk(chunk));
                }
                let length_bits = copy_length_bits(current);
                let Ok(encoded_offset) = u16::try_from(offset - 1) else {
                    return Err(invalid("compression copy offset exceeds u16"));
                };
                let Ok(encoded_length) = u16::try_from(length - MIN_COPY_BYTES) else {
                    return Err(invalid("compression copy length exceeds u16"));
                };
                let token = (encoded_offset << length_bits) | encoded_length;
                data.extend_from_slice(&token.to_le_bytes());
                flags |= 1u8 << token_index;
                current += length;
            } else {
                data.push(chunk[current]);
                current += 1;
            }

            for position in previous..current {
                insert_match_position(chunk, position, &mut positions)?;
            }
        }
        let flags_slot = data
            .get_mut(flags_index)
            .ok_or_else(|| invalid("compression flags offset is out of bounds"))?;
        *flags_slot = flags;
    }

    if current < chunk.len() {
        return Ok(raw_chunk(chunk));
    }

    let total_size = data.len() + CHUNK_HEADER_BYTES;
    let Ok(encoded_size) = u16::try_from(total_size - 3) else {
        return Err(invalid("compressed chunk size exceeds u16"));
    };
    let header = COMPRESSED_CHUNK_FLAG | CHUNK_SIGNATURE | encoded_size;
    let mut output = Vec::with_capacity(total_size);
    output.extend_from_slice(&header.to_le_bytes());
    output.extend_from_slice(&data);
    Ok(output)
}

fn find_match(
    chunk: &[u8],
    current: usize,
    positions: &HashMap<[u8; MIN_COPY_BYTES], Vec<u16>>,
) -> Option<(usize, usize)> {
    let bytes = chunk.get(current..current + MIN_COPY_BYTES)?;
    let key = [bytes[0], bytes[1], bytes[2]];
    let candidates = positions.get(&key)?;
    let length_bits = copy_length_bits(current);
    let maximum_length = ((1usize << length_bits) - 1 + MIN_COPY_BYTES).min(chunk.len() - current);
    let mut best_length = 0usize;
    let mut best_candidate = 0usize;

    // Candidates are stored in increasing order. Searching backwards and
    // replacing only on a longer match preserves the spec's nearest-match
    // tie behavior.
    for &candidate in candidates.iter().rev() {
        let candidate_position = usize::from(candidate);
        let mut length = MIN_COPY_BYTES;
        while length < maximum_length
            && chunk[current + length] == chunk[candidate_position + length]
        {
            length += 1;
        }
        if length > best_length {
            best_length = length;
            best_candidate = candidate_position;
            if length == maximum_length {
                break;
            }
        }
    }

    (best_length >= MIN_COPY_BYTES).then(|| (best_length, current - best_candidate))
}

fn insert_match_position(
    chunk: &[u8],
    position: usize,
    positions: &mut HashMap<[u8; MIN_COPY_BYTES], Vec<u16>>,
) -> Result<(), Error> {
    let Some(bytes) = chunk.get(position..position + MIN_COPY_BYTES) else {
        return Ok(());
    };
    let key = [bytes[0], bytes[1], bytes[2]];
    let Ok(encoded_position) = u16::try_from(position) else {
        return Err(invalid("compression dictionary position exceeds u16"));
    };
    positions.entry(key).or_default().push(encoded_position);
    Ok(())
}

fn raw_chunk(chunk: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(CHUNK_HEADER_BYTES + RAW_CHUNK_BYTES);
    output.extend_from_slice(&(CHUNK_SIGNATURE | CHUNK_SIZE_MASK).to_le_bytes());
    output.extend_from_slice(chunk);
    output.resize(CHUNK_HEADER_BYTES + RAW_CHUNK_BYTES, 0);
    output
}

fn append_compressed_checked(
    output: &mut Vec<u8>,
    bytes: &[u8],
    limits: &Limits,
) -> Result<(), Error> {
    let new_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("compressed VBA size overflow"))?;
    check_limit(
        "compressed VBA stream bytes",
        new_len,
        limits.max_compressed_stream_bytes,
    )?;
    output.extend_from_slice(bytes);
    Ok(())
}

/// Decompress one complete MS-OVBA `CompressedContainer`.
///
/// The decoder validates chunk signatures, raw-chunk sizes, copy-token
/// back-references, chunk output size, truncation, and the configured input
/// and output limits.
///
/// # Errors
///
/// Returns an error if the container is malformed or truncated, or if the
/// input or decompressed output exceeds the configured [`Limits`].
pub fn decode(compressed: &[u8], limits: &Limits) -> Result<Vec<u8>, Error> {
    check_limit(
        "compressed VBA stream bytes",
        compressed.len(),
        limits.max_compressed_stream_bytes,
    )?;
    if compressed.first().copied() != Some(CONTAINER_SIGNATURE) {
        return Err(invalid("compressed container signature must be 0x01"));
    }

    let mut cursor = 1usize;
    let mut output = Vec::new();
    while cursor < compressed.len() {
        let header_end = cursor
            .checked_add(CHUNK_HEADER_BYTES)
            .ok_or_else(|| invalid("compressed chunk offset overflow"))?;
        let header_bytes = compressed
            .get(cursor..header_end)
            .ok_or_else(|| invalid("truncated compressed chunk header"))?;
        let header = u16::from_le_bytes([header_bytes[0], header_bytes[1]]);
        if header & CHUNK_SIGNATURE_MASK != CHUNK_SIGNATURE {
            return Err(invalid(format!(
                "compressed chunk at offset {cursor} has an invalid signature"
            )));
        }

        let chunk_size = usize::from(header & CHUNK_SIZE_MASK)
            .checked_add(3)
            .ok_or_else(|| invalid("compressed chunk size overflow"))?;
        let chunk_end = cursor
            .checked_add(chunk_size)
            .ok_or_else(|| invalid("compressed chunk end overflow"))?;
        if chunk_end > compressed.len() {
            return Err(invalid(format!(
                "compressed chunk at offset {cursor} extends past the container"
            )));
        }
        let chunk_data = &compressed[header_end..chunk_end];
        let decompressed_before = output.len();

        if header & COMPRESSED_CHUNK_FLAG == 0 {
            if header & CHUNK_SIZE_MASK != CHUNK_SIZE_MASK || chunk_data.len() != RAW_CHUNK_BYTES {
                return Err(invalid("raw chunk must contain exactly 4096 bytes"));
            }
            append_checked(&mut output, chunk_data, limits)?;
        } else {
            decompress_chunk(chunk_data, &mut output, decompressed_before, limits)?;
        }
        cursor = chunk_end;
    }
    Ok(output)
}

fn decompress_chunk(
    chunk: &[u8],
    output: &mut Vec<u8>,
    chunk_start: usize,
    limits: &Limits,
) -> Result<(), Error> {
    if chunk.is_empty() {
        return Err(invalid("compressed chunk data must not be empty"));
    }

    let mut cursor = 0usize;
    while cursor < chunk.len() {
        let flags = chunk[cursor];
        cursor += 1;
        for token_index in 0..8u32 {
            if cursor >= chunk.len() {
                break;
            }
            if flags & (1u8 << token_index) == 0 {
                append_byte_checked(output, chunk[cursor], chunk_start, limits)?;
                cursor += 1;
                continue;
            }

            let token_end = cursor
                .checked_add(2)
                .ok_or_else(|| invalid("copy-token offset overflow"))?;
            let token_bytes = chunk
                .get(cursor..token_end)
                .ok_or_else(|| invalid("truncated VBA copy token"))?;
            let token = u16::from_le_bytes([token_bytes[0], token_bytes[1]]);
            cursor = token_end;

            let chunk_position = output.len() - chunk_start;
            if chunk_position == 0 {
                return Err(invalid("copy token precedes all chunk output"));
            }
            let length_bits = copy_length_bits(chunk_position);
            let length_mask = (1u16 << length_bits) - 1;
            let copy_length = usize::from(token & length_mask) + 3;
            let copy_offset = usize::from(token >> length_bits) + 1;
            if copy_offset > chunk_position {
                return Err(invalid(format!(
                    "copy token offset {copy_offset} exceeds chunk position {chunk_position}"
                )));
            }
            let new_chunk_size = chunk_position
                .checked_add(copy_length)
                .ok_or_else(|| invalid("copy-token output size overflow"))?;
            if new_chunk_size > RAW_CHUNK_BYTES {
                return Err(invalid("compressed chunk expands past 4096 bytes"));
            }
            let new_total = output
                .len()
                .checked_add(copy_length)
                .ok_or_else(|| invalid("decompressed VBA size overflow"))?;
            check_limit(
                "decompressed VBA stream bytes",
                new_total,
                limits.max_decompressed_stream_bytes,
            )?;

            let copy_start = output.len() - copy_offset;
            for index in 0..copy_length {
                let value = output
                    .get(copy_start + index)
                    .copied()
                    .ok_or_else(|| invalid("copy-token source is out of bounds"))?;
                output.push(value);
            }
        }
    }
    Ok(())
}

fn copy_length_bits(chunk_position: usize) -> u32 {
    let offset_bits = usize::BITS - chunk_position.saturating_sub(1).leading_zeros();
    16 - offset_bits.max(MIN_COPY_LENGTH_BITS)
}

fn append_checked(output: &mut Vec<u8>, bytes: &[u8], limits: &Limits) -> Result<(), Error> {
    let new_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("decompressed VBA size overflow"))?;
    check_limit(
        "decompressed VBA stream bytes",
        new_len,
        limits.max_decompressed_stream_bytes,
    )?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_byte_checked(
    output: &mut Vec<u8>,
    byte: u8,
    chunk_start: usize,
    limits: &Limits,
) -> Result<(), Error> {
    if output.len() - chunk_start >= RAW_CHUNK_BYTES {
        return Err(invalid("compressed chunk expands past 4096 bytes"));
    }
    let new_len = output
        .len()
        .checked_add(1)
        .ok_or_else(|| invalid("decompressed VBA size overflow"))?;
    check_limit(
        "decompressed VBA stream bytes",
        new_len,
        limits.max_decompressed_stream_bytes,
    )?;
    output.push(byte);
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test fixtures and assertions panic intentionally on failure"
    )]

    use super::*;

    const NO_COMPRESSION: &[u8] = &[
        0x01, 0x19, 0xb0, 0x00, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x00, 0x69, 0x6a,
        0x6b, 0x6c, 0x6d, 0x6e, 0x6f, 0x70, 0x00, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76, 0x2e,
    ];
    const NORMAL_COMPRESSION: &[u8] = &[
        0x01, 0x2f, 0xb0, 0x00, 0x23, 0x61, 0x61, 0x61, 0x62, 0x63, 0x64, 0x65, 0x82, 0x66, 0x00,
        0x70, 0x61, 0x67, 0x68, 0x69, 0x6a, 0x01, 0x38, 0x08, 0x61, 0x6b, 0x6c, 0x00, 0x30, 0x6d,
        0x6e, 0x6f, 0x70, 0x06, 0x71, 0x02, 0x70, 0x04, 0x10, 0x72, 0x73, 0x74, 0x75, 0x76, 0x10,
        0x77, 0x78, 0x79, 0x7a, 0x00, 0x3c,
    ];
    const MAX_COMPRESSION: &[u8] = &[0x01, 0x03, 0xb0, 0x02, 0x61, 0x45, 0x00];

    #[test]
    fn decodes_normative_spec_examples() {
        let limits = Limits::default();
        assert_eq!(
            decode(NO_COMPRESSION, &limits).unwrap(),
            b"abcdefghijklmnopqrstuv."
        );
        assert_eq!(
            decode(NORMAL_COMPRESSION, &limits).unwrap(),
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"
        );
        assert_eq!(decode(MAX_COMPRESSION, &limits).unwrap(), vec![b'a'; 73]);
    }

    #[test]
    fn encodes_normative_spec_examples() {
        let limits = Limits::default();
        assert_eq!(
            encode(b"abcdefghijklmnopqrstuv.", &limits).unwrap(),
            NO_COMPRESSION
        );
        let normal = encode(
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
            &limits,
        )
        .unwrap();
        assert_eq!(
            decode(&normal, &limits).unwrap(),
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"
        );
        assert!(normal.len() <= NORMAL_COMPRESSION.len());
        assert_eq!(encode(&[b'a'; 73], &limits).unwrap(), MAX_COMPRESSION);
    }

    #[test]
    fn round_trips_empty_boundary_and_multiple_chunks() {
        let limits = Limits::default();
        for input in [
            Vec::new(),
            vec![b'x'],
            vec![b'x'; 16],
            vec![b'x'; 17],
            vec![b'x'; RAW_CHUNK_BYTES],
            vec![b'x'; RAW_CHUNK_BYTES + 1],
            (0..RAW_CHUNK_BYTES * 3 + 271)
                .map(|index| b'a' + u8::try_from(index % 23).unwrap())
                .collect(),
        ] {
            let compressed = encode(&input, &limits).unwrap();
            assert_eq!(decode(&compressed, &limits).unwrap(), input);
        }
    }

    #[test]
    fn round_trips_across_copy_token_bit_partition_boundaries() {
        let limits = Limits::default();
        for size in [
            15usize, 16, 17, 31, 32, 33, 63, 64, 65, 127, 128, 129, 255, 256, 257, 511, 512, 513,
            1023, 1024, 1025, 2047, 2048, 2049, 4095, 4096,
        ] {
            let input: Vec<u8> = (0..size)
                .map(|index| b'a' + u8::try_from(index % 11).unwrap())
                .collect();
            let compressed = encode(&input, &limits).unwrap();
            assert_eq!(
                decode(&compressed, &limits).unwrap(),
                input,
                "copy-token partition failed at {size} bytes"
            );
        }
    }

    #[test]
    fn raw_partial_chunk_has_spec_required_zero_padding() {
        let mut state = 0x1234_5678u32;
        let input: Vec<u8> = (0..4000)
            .map(|_| {
                state ^= state << 13;
                state ^= state >> 17;
                state ^= state << 5;
                state.to_le_bytes()[0]
            })
            .collect();

        let compressed = encode(&input, &Limits::default()).unwrap();
        assert_eq!(&compressed[..3], &[CONTAINER_SIGNATURE, 0xff, 0x3f]);
        let decoded = decode(&compressed, &Limits::default()).unwrap();
        assert_eq!(&decoded[..input.len()], input);
        assert_eq!(decoded.len(), RAW_CHUNK_BYTES);
        assert!(decoded[input.len()..].iter().all(|byte| *byte == 0));
    }

    #[test]
    fn encoding_is_deterministic_and_prefers_nearest_equal_match() {
        let input = b"abcXabcYabcYabcY";
        let first = encode(input, &Limits::default()).unwrap();
        let second = encode(input, &Limits::default()).unwrap();
        assert_eq!(first, second);
        assert_eq!(decode(&first, &Limits::default()).unwrap(), input);

        // At position 8, position 4 supplies the nearest overlapping match
        // through the end. The token encodes offset 4 and length 8.
        assert_eq!(&first[first.len() - 2..], &[0x05, 0x30]);
    }

    #[test]
    fn decodes_raw_chunk() {
        let mut bytes = vec![CONTAINER_SIGNATURE, 0xff, 0x3f];
        bytes.extend((0..RAW_CHUNK_BYTES).map(|value| value.to_le_bytes()[0]));
        let decoded = decode(&bytes, &Limits::default()).unwrap();
        assert_eq!(decoded.len(), RAW_CHUNK_BYTES);
        assert_eq!(decoded[257], 1);
    }

    #[test]
    fn rejects_invalid_back_reference_and_truncation() {
        let invalid_copy = [0x01, 0x02, 0xb0, 0x01, 0x00, 0x00];
        assert!(decode(&invalid_copy, &Limits::default()).is_err());
        assert!(decode(&[0x01, 0x00], &Limits::default()).is_err());
        assert!(decode(&[0x01, 0x00, 0x30], &Limits::default()).is_err());
    }

    #[test]
    fn enforces_output_limit_before_copy_expansion() {
        let limits = Limits {
            max_decompressed_stream_bytes: 10,
            ..Limits::default()
        };
        assert!(matches!(
            decode(MAX_COMPRESSION, &limits),
            Err(Error::LimitExceeded { .. })
        ));
    }

    #[test]
    fn encoder_enforces_input_and_output_limits() {
        let input_limits = Limits {
            max_decompressed_stream_bytes: 3,
            ..Limits::default()
        };
        assert!(matches!(
            encode(b"four", &input_limits),
            Err(Error::LimitExceeded { .. })
        ));

        let output_limits = Limits {
            max_compressed_stream_bytes: 1,
            ..Limits::default()
        };
        assert_eq!(encode(&[], &output_limits).unwrap(), [CONTAINER_SIGNATURE]);
        assert!(matches!(
            encode(b"x", &output_limits),
            Err(Error::LimitExceeded { .. })
        ));
    }
}
