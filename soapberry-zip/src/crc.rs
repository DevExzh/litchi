const fn gen_crc_table() -> [[u32; 256]; 16] {
    let mut table: [[u32; 256]; 16] = [[0; 256]; 16];
    let poly = 0xEDB88320; // Polynomial used in CRC-32

    let mut i = 0;
    while i < 256 {
        let mut crc = i as u32;
        let mut j = 0;
        while j < 8 {
            if crc & 1 != 0 {
                crc = (crc >> 1) ^ poly;
            } else {
                crc >>= 1;
            }
            j += 1;
        }

        table[0][i] = crc;
        i += 1;
    }

    i = 1;
    while i < 16 {
        let mut j = 0;
        while j < 256 {
            table[i][j] = (table[i - 1][j] >> 8) ^ table[0][(table[i - 1][j] & 0xFF) as usize];
            j += 1;
        }
        i += 1;
    }

    table
}

// Prefer static over const to cut test times in half
// ref: https://github.com/srijs/rust-crc32fast/commit/e61ce6a39bbe9da495198a4037292ec299e8970f
static CRC_TABLE: [[u32; 256]; 16] = gen_crc_table();

/// Compute the CRC32 (IEEE) of a byte slice
///
/// Typically this function is used only to compute the CRC32 of data that is
/// held entirely in memory. When decompressing, a
/// [`ZipVerifier`](crate::ZipVerifier) is suitable to streaming computations.
///
/// Benchmarks showed that function should be fast enough for all uses, only
/// losing to `crc32fast` at the largest payload size and even then eking out a
/// single digit performance improvement.
pub fn crc32(data: &[u8]) -> u32 {
    crc32_chunk(data, 0)
}

#[inline]
pub fn crc32_chunk(data: &[u8], prev: u32) -> u32 {
    let mut chunks = data.chunks_exact(16);
    let mut crc = chunks.by_ref().fold(!prev, |crc, data| {
        CRC_TABLE[0x0][data[0xf] as usize]
            ^ CRC_TABLE[0x1][data[0xe] as usize]
            ^ CRC_TABLE[0x2][data[0xd] as usize]
            ^ CRC_TABLE[0x3][data[0xc] as usize]
            ^ CRC_TABLE[0x4][data[0xb] as usize]
            ^ CRC_TABLE[0x5][data[0xa] as usize]
            ^ CRC_TABLE[0x6][data[0x9] as usize]
            ^ CRC_TABLE[0x7][data[0x8] as usize]
            ^ CRC_TABLE[0x8][data[0x7] as usize]
            ^ CRC_TABLE[0x9][data[0x6] as usize]
            ^ CRC_TABLE[0xa][data[0x5] as usize]
            ^ CRC_TABLE[0xb][data[0x4] as usize]
            ^ CRC_TABLE[0xc][data[0x3] as usize ^ ((crc >> 0x18) & 0xFF) as usize]
            ^ CRC_TABLE[0xd][data[0x2] as usize ^ ((crc >> 0x10) & 0xFF) as usize]
            ^ CRC_TABLE[0xe][data[0x1] as usize ^ ((crc >> 0x08) & 0xFF) as usize]
            ^ CRC_TABLE[0xf][data[0x0] as usize ^ (crc & 0xFF) as usize]
    });

    crc = chunks.remainder().iter().fold(crc, |crc, &x| {
        (crc >> 8) ^ CRC_TABLE[0][(u32::from(x) ^ (crc & 0xFF)) as usize]
    });

    !crc
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_crc() {
        let table = gen_crc_table();
        assert_eq!(table[0][0], 0x0000_0000);
        assert_eq!(table[0][1], 0x77073096);
        assert_eq!(table[0][2], 0xee0e612c);
        assert_eq!(table[1][1], 0x191B3141);
        assert_eq!(table[1][2], 0x32366282);

        let abc = b"EU4txt\nchecksum=\"ced5411e2d4a5ec724595c2c4f1b7347\"";
        assert_eq!(crc32(abc), 1702863696);
    }
}
