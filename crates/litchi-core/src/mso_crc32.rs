//! MS-OSHARED `MsoCrc32Compute` checksum primitives.
//!
//! This is not the common IEEE CRC-32 algorithm. Microsoft Office specifies
//! polynomial `0xAF`, most-significant-bit-first byte ordering, and a cache
//! whose entries are masked to 16 bits.

const CACHE_MASK: u32 = 0xFFFF;
const POLYNOMIAL: u32 = 0xAF;
const CACHE: [u32; 256] = build_cache();

/// Streaming state for the MS-OSHARED `MsoCrc32Compute` algorithm.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[repr(transparent)]
pub struct MsoCrc32(u32);

impl MsoCrc32 {
    /// Create a checksum with the protocol's conventional zero seed.
    #[must_use]
    pub const fn new() -> Self {
        Self(0)
    }

    /// Create a checksum continuing from `initial`.
    #[must_use]
    pub const fn with_initial(initial: u32) -> Self {
        Self(initial)
    }

    /// Incorporate the next bytes of the stream.
    pub fn update(&mut self, data: &[u8]) {
        self.0 = update(self.0, data);
    }

    /// Return the checksum accumulated so far without consuming the state.
    #[must_use]
    pub const fn value(&self) -> u32 {
        self.0
    }

    /// Finish the stream and return its checksum.
    #[must_use]
    pub const fn finalize(self) -> u32 {
        self.0
    }
}

/// Continue `MsoCrc32Compute` from `initial` with another byte slice.
///
/// Passing the returned value into a later call is equivalent to hashing the
/// concatenated slices.
#[must_use]
pub fn update(initial: u32, data: &[u8]) -> u32 {
    data.iter().fold(initial, |crc, byte| {
        let index = ((crc >> 24) as u8 ^ byte) as usize;
        crc.wrapping_shl(8) ^ CACHE[index]
    })
}

/// Compute `MsoCrc32Compute` using the protocol's conventional zero seed.
#[must_use]
pub fn compute(data: &[u8]) -> u32 {
    update(0, data)
}

const fn build_cache() -> [u32; 256] {
    let mut cache = [0; 256];
    let mut index = 0_u8;
    loop {
        let mut value = (index as u32) << 24;
        let mut bit = 0;
        while bit < 8 {
            value = if value & 0x8000_0000 != 0 {
                value.wrapping_shl(1) ^ POLYNOMIAL
            } else {
                value.wrapping_shl(1)
            };
            bit += 1;
        }
        cache[index as usize] = value & CACHE_MASK;
        if index == u8::MAX {
            break;
        }
        index += 1;
    }
    cache
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn independent_known_vectors() {
        assert_eq!(compute(b""), 0x0000_0000);
        assert_eq!(compute(&[0x01]), 0x0000_00AF);
        assert_eq!(compute(b"123456789"), 0xBD0B_E338);
    }

    #[test]
    fn cache_covers_the_full_byte_domain() {
        assert_eq!(CACHE.len(), usize::from(u8::MAX) + 1);
        assert_eq!(compute(&[0]), CACHE[0]);
        assert_eq!(compute(&[u8::MAX]), CACHE[usize::from(u8::MAX)]);
    }

    #[test]
    fn exact_xml_byte_vector() {
        const XML: &[u8] =
            br#"<?xml version="1.0" encoding="UTF-8"?><root><value>litchi</value></root>"#;
        assert_eq!(compute(XML), 0xEE34_2778);
        assert_ne!(
            compute(b"<?xml version=\"1.0\"?><root/>\n"),
            compute(b"<?xml version=\"1.0\"?><root/>")
        );
    }

    #[test]
    fn chunked_and_one_shot_computation_are_equivalent() {
        let bytes = b"SummaryInformation bytes";
        let expected = compute(bytes);

        let mut checksum = MsoCrc32::new();
        checksum.update(&bytes[..7]);
        checksum.update(&[]);
        checksum.update(&bytes[7..18]);
        let intermediate = checksum.value();
        checksum.update(&bytes[18..]);

        assert_eq!(checksum.finalize(), expected);
        assert_eq!(update(intermediate, &bytes[18..]), expected);
        assert_eq!(MsoCrc32::with_initial(expected).finalize(), expected);
    }
}
