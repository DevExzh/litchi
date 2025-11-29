/// Compute the CRC32 (IEEE) of a byte slice.
///
/// Uses `crc32fast` which provides hardware-accelerated CRC32 using
/// SIMD/PCLMULQDQ instructions when available, falling back to a fast
/// software implementation otherwise.
#[inline]
pub fn crc32(data: &[u8]) -> u32 {
    crc32fast::hash(data)
}

/// Compute CRC32 incrementally, combining with a previous CRC value.
///
/// This is useful for streaming CRC32 computation where data arrives in chunks.
#[inline]
pub fn crc32_chunk(data: &[u8], prev: u32) -> u32 {
    let mut hasher = crc32fast::Hasher::new_with_initial(prev);
    hasher.update(data);
    hasher.finalize()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_crc() {
        // Test known CRC32 value
        let data = b"EU4txt\nchecksum=\"ced5411e2d4a5ec724595c2c4f1b7347\"";
        assert_eq!(crc32(data), 1702863696);

        // Test incremental CRC32
        let full = crc32(b"hello world");
        let incremental = crc32_chunk(b" world", crc32(b"hello"));
        assert_eq!(full, incremental);
    }
}
