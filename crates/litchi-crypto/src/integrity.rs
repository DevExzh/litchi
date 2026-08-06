//! MS-OFFCRYPTO integrity side streams for public OLE property sets.

use std::fmt;
use std::sync::LazyLock;

const STREAM_ID: u8 = 0xAB;
const CURRENT_VERSION: u8 = 0;
const HEADER_BYTES: usize = 6;
const CRC_CACHE_MASK: u32 = 0xFFFF;
const CRC_POLYNOMIAL: u32 = 0xAF;
static CRC_CACHE: LazyLock<[u32; 256]> = LazyLock::new(build_crc_cache);

pub const SUMMARY_HASH_STREAM: &str = "EncryptedSIHash";
pub const DOCUMENT_SUMMARY_HASH_STREAM: &str = "EncryptedDSIHash";
pub const SUMMARY_STREAM: &str = "\u{0005}SummaryInformation";
pub const DOCUMENT_SUMMARY_STREAM: &str = "\u{0005}DocumentSummaryInformation";

/// Parsed Info structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Info {
    /// Version 0, whose checksum and reserved bytes are understood.
    Version0 { checksum: u32, reserved: Vec<u8> },
    /// A future version that readers are required to ignore.
    UnsupportedVersion { version: u8 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    Truncated,
    InvalidStreamId(u8),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Truncated => write!(formatter, "encrypted property hash stream is truncated"),
            Self::InvalidStreamId(value) => write!(
                formatter,
                "encrypted property hash stream ID is {value:#04X}, expected {STREAM_ID:#04X}"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Parse an encrypted property hash stream into its [`Info`] structure.
///
/// # Errors
///
/// Returns [`Error::Truncated`] when `data` ends before the fixed header, and
/// [`Error::InvalidStreamId`] when the leading byte is not the expected
/// stream identifier.
pub fn parse(data: &[u8]) -> Result<Info, Error> {
    let Some((&stream_id, tail)) = data.split_first() else {
        return Err(Error::Truncated);
    };
    if stream_id != STREAM_ID {
        return Err(Error::InvalidStreamId(stream_id));
    }
    let Some((&version, payload)) = tail.split_first() else {
        return Err(Error::Truncated);
    };
    if version != CURRENT_VERSION {
        return Ok(Info::UnsupportedVersion { version });
    }
    let checksum = payload
        .get(..4)
        .ok_or(Error::Truncated)?
        .try_into()
        .map_err(|_err| Error::Truncated)?;
    Ok(Info::Version0 {
        checksum: u32::from_le_bytes(checksum),
        reserved: payload[4..].to_vec(),
    })
}

/// Serialize a version 0 stream from `checksum` and its reserved bytes.
#[must_use]
pub fn write(checksum: u32, reserved: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(HEADER_BYTES + reserved.len());
    output.push(STREAM_ID);
    output.push(CURRENT_VERSION);
    output.extend_from_slice(&checksum.to_le_bytes());
    output.extend_from_slice(reserved);
    output
}

/// Continue the MS-OSHARED `MsoCrc32Compute` algorithm from `initial`.
///
/// Passing the returned value into a later call is equivalent to hashing the
/// concatenated slices. The value normally used for a new stream is zero.
#[must_use]
pub fn update(initial: u32, data: &[u8]) -> u32 {
    data.iter().fold(initial, |crc, byte| {
        let index = ((crc >> 24) as u8 ^ byte) as usize;
        crc.wrapping_shl(8) ^ CRC_CACHE[index]
    })
}

/// Compute the property-stream checksum from the protocol's zero seed.
#[must_use]
pub fn crc32(data: &[u8]) -> u32 {
    update(0, data)
}

/// Check `property_stream` against the checksum recorded in `info`.
///
/// Returns `None` for [`Info::UnsupportedVersion`], whose payload readers are
/// required to ignore.
#[must_use]
pub fn verify(info: &Info, property_stream: &[u8]) -> Option<bool> {
    match info {
        Info::Version0 { checksum, .. } => Some(*checksum == crc32(property_stream)),
        Info::UnsupportedVersion { .. } => None,
    }
}

fn build_crc_cache() -> [u32; 256] {
    let mut cache = [0; 256];
    for (index, slot) in (0u32..256).zip(cache.iter_mut()) {
        let mut value = index << 24;
        for _ in 0..8 {
            value = if value & 0x8000_0000 != 0 {
                value.wrapping_shl(1) ^ CRC_POLYNOMIAL
            } else {
                value.wrapping_shl(1)
            };
        }
        *slot = value & CRC_CACHE_MASK;
    }
    cache
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test code panics on failure; unwrap keeps assertions concise"
    )]
    use super::*;

    #[test]
    fn version_zero_round_trips_with_reserved_bytes() {
        let bytes = write(0x1234_5678, &[9, 8, 7]);
        assert_eq!(
            parse(&bytes).unwrap(),
            Info::Version0 {
                checksum: 0x1234_5678,
                reserved: vec![9, 8, 7],
            }
        );
    }

    #[test]
    fn future_versions_are_ignored_without_interpreting_payload() {
        assert_eq!(
            parse(&[STREAM_ID, 3]).unwrap(),
            Info::UnsupportedVersion { version: 3 }
        );
    }

    #[test]
    fn invalid_header_is_rejected() {
        assert_eq!(parse(&[]).unwrap_err(), Error::Truncated);
        assert_eq!(
            parse(&[1, 0, 0, 0, 0, 0]).unwrap_err(),
            Error::InvalidStreamId(1)
        );
        assert_eq!(parse(&[STREAM_ID, 0, 1]).unwrap_err(), Error::Truncated);
    }

    #[test]
    fn crc_is_incremental_and_can_verify_an_info_stream() {
        let whole = crc32(b"SummaryInformation bytes");
        let first = update(0, b"Summary");
        assert_eq!(whole, update(first, b"Information bytes"));
        assert_eq!(crc32(b"123456789"), 0xBD0B_E338);
        let info = Info::Version0 {
            checksum: whole,
            reserved: Vec::new(),
        };
        assert_eq!(verify(&info, b"SummaryInformation bytes"), Some(true));
        assert_eq!(verify(&info, b"changed"), Some(false));
    }
}
