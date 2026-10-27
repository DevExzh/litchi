//! Package helpers of the hidden-axis engine: the exact physical source,
//! one component parsed for editing and its recompression, archive-header
//! metadata checks, the source fingerprint and compressed-size bounds.
//!
//! These are the Pages page-layout and table-lock helpers the engine used
//! before it moved out of `litchi-pages` (record 0379), with the same error
//! mapping onto the engine's [`Error`].

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len};
use litchi_iwa_core::{Archive, ArchiveObject, SnappyStream};

use super::{Error, LimitKind, Refusal};
use crate::TableSource;

const SNAPPY_RAW_MAX_OVERHEAD: usize = 32;
const SNAPPY_RAW_MAX_EXPANSION_DIVISOR: usize = 6;
const SNAPPY_FRAME_HEADER_BYTES: usize = 4;
const DEFLATE_MAX_LITERAL_BITS: usize = 9;
const DEFLATE_BLOCK_OVERHEAD_BYTES: usize = 3;

/// The exact physical source of `package`.
pub(super) fn physical(package: &impl TableSource) -> Result<&SourceCatalog, Error> {
    package
        .physical()
        .ok_or(Error::Refused(Refusal::UnsupportedSource))
}

/// A page-layout helper failure, mapped as the Pages owner mapped it: every
/// limit is a payload-bytes limit.
fn layout_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => Error::Refused(Refusal::UnsupportedSource),
        litchi_iwa_archive::Error::Iwa(core_error) => layout_core_error(core_error),
        _ => Error::InvalidSource,
    }
}

fn layout_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::PayloadBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

/// Parse `component_name` from its exact compressed member for editing.
pub(super) fn editable_archive(
    package: &impl TableSource,
    component_name: &str,
) -> Result<(Archive, litchi_iwa_core::Limits), Error> {
    let source = physical(package)?;
    let entry = source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let limits = source
        .limits()
        .effective_archive_limits()
        .map_err(layout_archive_error)?;
    let stream = SnappyStream::decompress(
        entry.data(),
        source
            .limits()
            .snappy_limits()
            .map_err(layout_archive_error)?,
    )
    .map_err(layout_core_error)?;
    let archive = Archive::parse(stream.as_bytes(), limits).map_err(layout_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    Ok((archive, limits))
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), Error> {
    for object in &archive.objects {
        let offset =
            usize::try_from(object.header_offset).map_err(|_error| Error::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(Error::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| Error::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(Error::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_error| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(Error::InvalidSource)?
                != object.data_offset
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

/// Encode and compress `archive` under the source's retained limits.
pub(super) fn compress_archive(
    source: &impl TableSource,
    archive: Archive,
    limits: litchi_iwa_core::Limits,
) -> Result<Vec<u8>, Error> {
    let bytes = archive.to_bytes(limits).map_err(layout_core_error)?;
    drop(archive);
    let snappy_limits = physical(source)?
        .limits()
        .snappy_limits()
        .map_err(layout_archive_error)?;
    SnappyStream::compress(&bytes, snappy_limits).map_err(layout_core_error)
}

/// Refuse merge or diff archive-header metadata on the selected message.
pub(super) fn validate_selected_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), Error> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

/// Require exactly one aggregate and at most one field-local declaration of
/// `identifier`, the latter at `accepted_path`.
pub(super) fn validate_reference_metadata(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), Error> {
    validate_selected_metadata(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    let aggregate_occurrences = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate_occurrences != 1 {
        return Err(Error::InvalidSource);
    }
    let mut field_declaration_seen = false;
    for field in &info.field_infos {
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if occurrences != 1
            || std::mem::replace(&mut field_declaration_seen, true)
            || field.path.as_slice() != accepted_path
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

/// The FNV-1a fingerprint of a package's exact bytes.
#[must_use]
pub fn fingerprint(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// A checked upper bound for the IWA Snappy stream of `input_len` bytes.
pub(super) fn snappy_compressed_bound(input_len: usize) -> Option<usize> {
    let full_chunk_bound = SnappyStream::WRITE_CHUNK_SIZE
        .checked_add(SnappyStream::WRITE_CHUNK_SIZE / SNAPPY_RAW_MAX_EXPANSION_DIVISOR)?
        .checked_add(SNAPPY_RAW_MAX_OVERHEAD)?
        .checked_add(SNAPPY_FRAME_HEADER_BYTES)?;
    let full_chunks = input_len / SnappyStream::WRITE_CHUNK_SIZE;
    let remainder = input_len % SnappyStream::WRITE_CHUNK_SIZE;
    let mut bound = full_chunks.checked_mul(full_chunk_bound)?;
    if remainder != 0 {
        let remainder_bound = remainder
            .checked_add(remainder / SNAPPY_RAW_MAX_EXPANSION_DIVISOR)?
            .checked_add(SNAPPY_RAW_MAX_OVERHEAD)?
            .checked_add(SNAPPY_FRAME_HEADER_BYTES)?;
        bound = bound.checked_add(remainder_bound)?;
    }
    Some(bound)
}

/// A checked upper bound for the raw Deflate output of `input_len` bytes.
pub(super) fn deflate_compressed_bound(input_len: usize) -> Option<usize> {
    let literal_bits = input_len
        .checked_mul(DEFLATE_MAX_LITERAL_BITS - 8)?
        .checked_add(7)?;
    let literal_bytes = literal_bits / 8;
    let small_input_overhead = usize::from(input_len == 0) + usize::from(input_len < 9);
    input_len
        .checked_add(literal_bytes)
        .and_then(|value| value.checked_add(small_input_overhead))
        .and_then(|value| value.checked_add(DEFLATE_BLOCK_OVERHEAD_BYTES))
}
