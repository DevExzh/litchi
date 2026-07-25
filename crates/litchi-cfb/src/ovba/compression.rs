//! MS-OVBA compressed-container encoding and decoding.

use super::{VbaError, VbaLimits, check_limit, invalid};
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
/// sizes are checked against [`VbaLimits`].
///
/// A raw final chunk is padded with zero bytes to 4096 bytes, as required by
/// [MS-OVBA]. Consequently, decompressing an incompressible final partial
/// chunk can produce trailing zero bytes beyond the original input.
pub fn compress_container(decompressed: &[u8], limits: &VbaLimits) -> Result<Vec<u8>, VbaError> {
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
        let encoded = compress_chunk(chunk);
        append_compressed_checked(&mut output, &encoded, limits)?;
    }
    Ok(output)
}

fn compress_chunk(chunk: &[u8]) -> Vec<u8> {
    debug_assert!(!chunk.is_empty());
    debug_assert!(chunk.len() <= RAW_CHUNK_BYTES);

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
                    return raw_chunk(chunk);
                }
                let length_bits = copy_length_bits(current);
                let token = (((offset - 1) as u16) << length_bits)
                    | u16::try_from(length - MIN_COPY_BYTES)
                        .expect("copy length is bounded by a 12-bit field");
                data.extend_from_slice(&token.to_le_bytes());
                flags |= 1u8 << token_index;
                current += length;
            } else {
                data.push(chunk[current]);
                current += 1;
            }

            for position in previous..current {
                insert_match_position(chunk, position, &mut positions);
            }
        }
        data[flags_index] = flags;
    }

    if current < chunk.len() {
        return raw_chunk(chunk);
    }

    let total_size = data.len() + CHUNK_HEADER_BYTES;
    let header = COMPRESSED_CHUNK_FLAG
        | CHUNK_SIGNATURE
        | u16::try_from(total_size - 3).expect("compressed chunk is at most 4098 bytes");
    let mut output = Vec::with_capacity(total_size);
    output.extend_from_slice(&header.to_le_bytes());
    output.extend_from_slice(&data);
    output
}

fn find_match(
    chunk: &[u8],
    current: usize,
    positions: &HashMap<[u8; MIN_COPY_BYTES], Vec<u16>>,
) -> Option<(usize, usize)> {
    let key: [u8; MIN_COPY_BYTES] = chunk
        .get(current..current + MIN_COPY_BYTES)?
        .try_into()
        .expect("slice length is checked");
    let candidates = positions.get(&key)?;
    let length_bits = copy_length_bits(current);
    let maximum_length = ((1usize << length_bits) - 1 + MIN_COPY_BYTES).min(chunk.len() - current);
    let mut best_length = 0usize;
    let mut best_candidate = 0usize;

    // Candidates are stored in increasing order. Searching backwards and
    // replacing only on a longer match preserves the spec's nearest-match
    // tie behavior.
    for &candidate in candidates.iter().rev() {
        let candidate = usize::from(candidate);
        let mut length = MIN_COPY_BYTES;
        while length < maximum_length && chunk[current + length] == chunk[candidate + length] {
            length += 1;
        }
        if length > best_length {
            best_length = length;
            best_candidate = candidate;
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
) {
    let Some(bytes) = chunk.get(position..position + MIN_COPY_BYTES) else {
        return;
    };
    let key: [u8; MIN_COPY_BYTES] = bytes.try_into().expect("slice length is checked");
    positions
        .entry(key)
        .or_default()
        .push(u16::try_from(position).expect("chunk positions are at most 4095"));
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
    limits: &VbaLimits,
) -> Result<(), VbaError> {
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
pub fn decompress_container(compressed: &[u8], limits: &VbaLimits) -> Result<Vec<u8>, VbaError> {
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
    limits: &VbaLimits,
) -> Result<(), VbaError> {
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
                let value = output[copy_start + index];
                output.push(value);
            }
        }
    }
    Ok(())
}

fn copy_length_bits(chunk_position: usize) -> u32 {
    let offset_bits = usize::BITS - (chunk_position - 1).leading_zeros();
    16 - offset_bits.max(MIN_COPY_LENGTH_BITS)
}

fn append_checked(output: &mut Vec<u8>, bytes: &[u8], limits: &VbaLimits) -> Result<(), VbaError> {
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
    limits: &VbaLimits,
) -> Result<(), VbaError> {
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
        let limits = VbaLimits::default();
        assert_eq!(
            decompress_container(NO_COMPRESSION, &limits).unwrap(),
            b"abcdefghijklmnopqrstuv."
        );
        assert_eq!(
            decompress_container(NORMAL_COMPRESSION, &limits).unwrap(),
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"
        );
        assert_eq!(
            decompress_container(MAX_COMPRESSION, &limits).unwrap(),
            vec![b'a'; 73]
        );
    }

    #[test]
    fn encodes_normative_spec_examples() {
        let limits = VbaLimits::default();
        assert_eq!(
            compress_container(b"abcdefghijklmnopqrstuv.", &limits).unwrap(),
            NO_COMPRESSION
        );
        let normal = compress_container(
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa",
            &limits,
        )
        .unwrap();
        assert_eq!(
            decompress_container(&normal, &limits).unwrap(),
            b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"
        );
        assert!(normal.len() <= NORMAL_COMPRESSION.len());
        assert_eq!(
            compress_container(&[b'a'; 73], &limits).unwrap(),
            MAX_COMPRESSION
        );
    }

    #[test]
    fn round_trips_empty_boundary_and_multiple_chunks() {
        let limits = VbaLimits::default();
        for input in [
            Vec::new(),
            vec![b'x'],
            vec![b'x'; 16],
            vec![b'x'; 17],
            vec![b'x'; RAW_CHUNK_BYTES],
            vec![b'x'; RAW_CHUNK_BYTES + 1],
            (0..RAW_CHUNK_BYTES * 3 + 271)
                .map(|index| b'a' + (index % 23) as u8)
                .collect(),
        ] {
            let compressed = compress_container(&input, &limits).unwrap();
            assert_eq!(decompress_container(&compressed, &limits).unwrap(), input);
        }
    }

    #[test]
    fn round_trips_across_copy_token_bit_partition_boundaries() {
        let limits = VbaLimits::default();
        for size in [
            15usize, 16, 17, 31, 32, 33, 63, 64, 65, 127, 128, 129, 255, 256, 257, 511, 512, 513,
            1023, 1024, 1025, 2047, 2048, 2049, 4095, 4096,
        ] {
            let input: Vec<u8> = (0..size).map(|index| b'a' + (index % 11) as u8).collect();
            let compressed = compress_container(&input, &limits).unwrap();
            assert_eq!(
                decompress_container(&compressed, &limits).unwrap(),
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
                state as u8
            })
            .collect();

        let compressed = compress_container(&input, &VbaLimits::default()).unwrap();
        assert_eq!(&compressed[..3], &[CONTAINER_SIGNATURE, 0xff, 0x3f]);
        let decoded = decompress_container(&compressed, &VbaLimits::default()).unwrap();
        assert_eq!(&decoded[..input.len()], input);
        assert_eq!(decoded.len(), RAW_CHUNK_BYTES);
        assert!(decoded[input.len()..].iter().all(|byte| *byte == 0));
    }

    #[test]
    fn encoding_is_deterministic_and_prefers_nearest_equal_match() {
        let input = b"abcXabcYabcYabcY";
        let first = compress_container(input, &VbaLimits::default()).unwrap();
        let second = compress_container(input, &VbaLimits::default()).unwrap();
        assert_eq!(first, second);
        assert_eq!(
            decompress_container(&first, &VbaLimits::default()).unwrap(),
            input
        );

        // At position 8, position 4 supplies the nearest overlapping match
        // through the end. The token encodes offset 4 and length 8.
        assert_eq!(&first[first.len() - 2..], &[0x05, 0x30]);
    }

    #[test]
    fn decodes_raw_chunk() {
        let mut bytes = vec![CONTAINER_SIGNATURE, 0xff, 0x3f];
        bytes.extend((0..RAW_CHUNK_BYTES).map(|value| value as u8));
        let decoded = decompress_container(&bytes, &VbaLimits::default()).unwrap();
        assert_eq!(decoded.len(), RAW_CHUNK_BYTES);
        assert_eq!(decoded[257], 1);
    }

    #[test]
    fn rejects_invalid_back_reference_and_truncation() {
        let invalid_copy = [0x01, 0x02, 0xb0, 0x01, 0x00, 0x00];
        assert!(decompress_container(&invalid_copy, &VbaLimits::default()).is_err());
        assert!(decompress_container(&[0x01, 0x00], &VbaLimits::default()).is_err());
        assert!(decompress_container(&[0x01, 0x00, 0x30], &VbaLimits::default()).is_err());
    }

    #[test]
    fn enforces_output_limit_before_copy_expansion() {
        let limits = VbaLimits {
            max_decompressed_stream_bytes: 10,
            ..VbaLimits::default()
        };
        assert!(matches!(
            decompress_container(MAX_COMPRESSION, &limits),
            Err(VbaError::LimitExceeded { .. })
        ));
    }

    #[test]
    fn encoder_enforces_input_and_output_limits() {
        let input_limits = VbaLimits {
            max_decompressed_stream_bytes: 3,
            ..VbaLimits::default()
        };
        assert!(matches!(
            compress_container(b"four", &input_limits),
            Err(VbaError::LimitExceeded { .. })
        ));

        let output_limits = VbaLimits {
            max_compressed_stream_bytes: 1,
            ..VbaLimits::default()
        };
        assert_eq!(
            compress_container(&[], &output_limits).unwrap(),
            [CONTAINER_SIGNATURE]
        );
        assert!(matches!(
            compress_container(b"x", &output_limits),
            Err(VbaError::LimitExceeded { .. })
        ));
    }
}
