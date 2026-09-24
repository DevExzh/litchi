//! Bounded reads for inert `MS-OFFCRYPTO` DataSpaces streams.
//!
//! DataSpaces metadata is stored in CFB streams whose directory-declared
//! length is untrusted input.  Callers must inspect that length before asking
//! the CFB reader to materialize a payload.  The helpers here retain the
//! existing `OleError` surface while using the positional range reader, so a
//! rejected stream performs no payload I/O and an accepted stream allocates
//! only after its declared length has passed the ceiling.

use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

/// Maximum serialized size of one DataSpaces stream read by the shared codec.
pub const MAX_STREAM_BYTES: usize = 16 * 1024 * 1024;

/// Reads one DataSpaces stream after checking its declared length.
///
/// The directory entry is inspected before any payload allocation or read.
/// Streams at or below [`MAX_STREAM_BYTES`] retain their ordinary CFB
/// semantics; oversized streams return [`OleError::LimitExceeded`] without
/// following their FAT or MiniFAT chain.
///
/// # Errors
///
/// Returns the underlying CFB lookup, allocation, range-read, or limit error.
pub fn read_stream<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[&str],
) -> Result<Vec<u8>, OleError> {
    read_stream_with_limit(ole, path, MAX_STREAM_BYTES)
}

/// Reads one DataSpaces stream with an explicit byte ceiling.
///
/// This is the parameterized form used by bounded callers and tests.  The
/// declared stream length is checked through [`OleFile::stream_len`] before a
/// destination is allocated or [`OleFile::read_stream_range`] is called.
/// The limit is inclusive: a stream exactly equal to `max_bytes` is accepted.
/// Callers may lower the shared ceiling, but may not raise
/// [`MAX_STREAM_BYTES`].
///
/// # Errors
///
/// Returns [`OleError::InvalidLimit`] for a zero or above-ceiling limit, or
/// the underlying CFB lookup, allocation, range-read, or limit error.
pub fn read_stream_with_limit<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[&str],
    max_bytes: usize,
) -> Result<Vec<u8>, OleError> {
    let hard_maximum = MAX_STREAM_BYTES as u64;
    let value = u64::try_from(max_bytes).unwrap_or(u64::MAX);
    if max_bytes == 0 || max_bytes > MAX_STREAM_BYTES {
        return Err(OleError::InvalidLimit {
            resource: "DataSpaces stream bytes",
            value,
            maximum: hard_maximum,
        });
    }

    let declared_len = ole.stream_len(path)?;
    if declared_len > value {
        return Err(OleError::LimitExceeded {
            resource: "DataSpaces stream bytes",
            observed: declared_len,
            maximum: value,
        });
    }

    let length = usize::try_from(declared_len).map_err(|_error| {
        OleError::InvalidData("DataSpaces stream length does not fit usize".to_string())
    })?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(length)
        .map_err(|source| OleError::Allocation {
            resource: "DataSpaces stream bytes",
            source,
        })?;
    bytes.resize(length, 0);
    ole.read_stream_range(path, 0, &mut bytes)?;
    Ok(bytes)
}
