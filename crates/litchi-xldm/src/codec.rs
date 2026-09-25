//! MS-XLDM outer-storage codec: bounded XML projection and exact snapshots.

use super::identity::Xldm140FileReplacement;
use super::model::{
    BOM, CRC_SIZE, Compression, FileEntry, FileKind, Header, MAX_DIRECTORY_BYTES, MAX_FILES,
    MAX_PARTITIONS, MAX_PATH_BYTES, MAX_STORAGE_BYTES, MAX_XML_DEPTH, MAX_XML_NODES,
    MAX_XML_TEXT_BYTES, Node, Offset, PartitionMarker, Size, Storage, StorageProfile,
    XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE, XmlEncoding,
};
use super::semantic::parse_backup_log;
use super::validation::{validate_allocations, validate_backup_log, validate_paths};
use crate::error::{Error, Result, allocation};
use crate::xml_attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

/// Validate and inspect the outer MS-XLDM virtual storage.
pub fn inspect(bytes: &[u8]) -> Result<Storage<'_>> {
    if bytes.len() > MAX_STORAGE_BYTES {
        return Err(limit("storage bytes"));
    }
    if bytes.len() < XLDM_PAGE_SIZE * 3 || !bytes.len().is_multiple_of(XLDM_PAGE_SIZE) {
        return Err(invalid(
            "MS-XLDM storage must contain at least three complete 4096-byte pages",
        ));
    }
    if bytes[..2] != BOM {
        return Err(invalid("MS-XLDM header byte-order mark is missing"));
    }
    let signature = utf16le(XLDM_STREAM_SIGNATURE);
    if bytes.get(2..2 + signature.len()) != Some(signature.as_slice()) {
        return Err(invalid("MS-XLDM stream storage signature is invalid"));
    }
    let header_xml_start = 2 + signature.len();
    let close = utf16le("</BackupLog>");
    let relative_end = memchr::memmem::rfind(&bytes[header_xml_start..XLDM_PAGE_SIZE], &close)
        .ok_or_else(|| invalid("MS-XLDM header BackupLog closing element is missing"))?;
    let header_xml_end = header_xml_start
        .checked_add(relative_end)
        .and_then(|value| value.checked_add(close.len()))
        .ok_or_else(|| limit("header XML range"))?;
    if bytes[header_xml_end..XLDM_PAGE_SIZE]
        .iter()
        .any(|byte| *byte != 0)
    {
        return Err(invalid("MS-XLDM header padding must be zero"));
    }
    let (header_xml, header_encoding) = decode_xml(&bytes[header_xml_start..header_xml_end], true)?;
    let header = parse_header(&parse_xml(&header_xml)?)?;
    let profile = if header.backup_restore_sync_version == 140 {
        StorageProfile::Xldm140
    } else {
        StorageProfile::Tabular150
    };
    let data_offset = checked_usize(header.data_offset.0, "data offset")?;
    let directory_offset = checked_usize(header.directory_offset.0, "directory offset")?;
    let directory_size = checked_usize(header.directory_size.0, "directory size")?;
    if data_offset != XLDM_PAGE_SIZE {
        return Err(invalid(
            "MS-XLDM data offset must equal the 4096-byte header size",
        ));
    }
    if directory_offset < XLDM_PAGE_SIZE * 2 || directory_offset % XLDM_PAGE_SIZE != 0 {
        return Err(invalid(
            "MS-XLDM directory offset must be page aligned after the files section",
        ));
    }
    if directory_size == 0 || directory_size > MAX_DIRECTORY_BYTES {
        return Err(limit("directory bytes"));
    }
    let directory_end = directory_offset
        .checked_add(directory_size)
        .ok_or_else(|| limit("directory range"))?;
    if directory_end > bytes.len() {
        return Err(invalid("MS-XLDM virtual directory extends beyond storage"));
    }
    if bytes[directory_end..].iter().any(|byte| *byte != 0) {
        return Err(invalid("MS-XLDM directory padding must be zero"));
    }
    if bytes.get(data_offset..data_offset + 2) != Some(&BOM) {
        return Err(invalid("MS-XLDM files-section byte-order mark is missing"));
    }
    if bytes.get(directory_offset..directory_offset + 2) == Some(&BOM) {
        return Err(invalid(
            "MS-XLDM virtual directory must not have a byte-order mark",
        ));
    }
    let (directory_xml, directory_encoding) =
        decode_xml(&bytes[directory_offset..directory_end], false)?;
    let mut files = parse_directory(&parse_xml(&directory_xml)?)?;
    if files.len() != header.file_count as usize {
        return Err(invalid(format!(
            "header file count {} does not match directory count {}",
            header.file_count,
            files.len()
        )));
    }
    if files.len() < 2 {
        return Err(invalid(
            "MS-XLDM storage requires partition and backup-log allocations",
        ));
    }
    validate_paths(&files)?;
    let mut order: Vec<usize> = (0..files.len()).collect();
    order.sort_by_key(|index| files[*index].offset);
    validate_allocations(
        bytes,
        data_offset,
        directory_offset,
        &mut files,
        &order,
        profile,
    )?;
    let first = order[0];
    let partition_bytes = payload_slice(bytes, &files[first])?;
    let (partitions_xml, partition_encoding) = decode_marker_xml(partition_bytes, profile)?;
    let partition_count = parse_partitions(&parse_xml(&partitions_xml)?, profile)?;
    files[first].kind = FileKind::Partitions;
    let last = *order.last().unwrap_or_else(|| {
        crate::error::panic_missing_invariant("required value was checked before extraction")
    });
    files[last].kind = FileKind::BackupLog;
    let (backup_xml, backup_encoding) =
        decode_marker_xml(payload_slice(bytes, &files[last])?, profile)?;
    let backup_log = parse_backup_log(&parse_xml(&backup_xml)?, backup_encoding, profile)?;
    validate_backup_log(&backup_log, &files, first, last, bytes, profile)?;
    Ok(Storage {
        header,
        header_encoding,
        directory_encoding,
        partition_marker: PartitionMarker {
            partition_count,
            encoding: partition_encoding,
            encoded_xml: partition_bytes,
        },
        backup_log,
        files,
        bytes,
        profile,
    })
}

/// Revalidate and return the original byte stream exactly.
pub fn write(storage: &Storage<'_>) -> Result<Vec<u8>> {
    inspect(storage.bytes)?;
    Ok(storage.bytes.to_vec())
}

/// Rewrite already-admitted allocation payloads without changing their
/// lengths, directory records, or serial allocation layout.
pub(super) fn rewrite_same_size_payloads(
    storage: &Storage<'_>,
    replacements: &[Xldm140FileReplacement<'_>],
) -> Result<Vec<u8>> {
    if storage.profile != StorageProfile::Xldm140 {
        return Err(Error::Unsupported {
            feature: "same-size inner XLDM rewrites require the version-140 profile",
        });
    }
    if storage.bytes.len() > MAX_STORAGE_BYTES {
        return Err(limit("storage bytes"));
    }
    if replacements.len() > MAX_FILES {
        return Err(limit("replacement count"));
    }
    let mut indexes = Vec::new();
    indexes
        .try_reserve_exact(replacements.len())
        .map_err(|source| allocation("MS-XLDM replacement indexes", source))?;
    let mut paths = std::collections::HashSet::new();
    paths
        .try_reserve(replacements.len())
        .map_err(|source| allocation("MS-XLDM replacement paths", source))?;
    for replacement in replacements {
        if !paths.insert(replacement.storage_path) {
            return Err(invalid(format!(
                "duplicate XLDM replacement path {}",
                replacement.storage_path
            )));
        }
        let index = storage
            .files
            .iter()
            .position(|entry| entry.path == replacement.storage_path)
            .ok_or_else(|| {
                invalid(format!(
                    "XLDM replacement path {} is absent from the source directory",
                    replacement.storage_path
                ))
            })?;
        let source = storage.file_payload(index).ok_or_else(|| {
            invalid(format!(
                "cannot resolve XLDM source payload {}",
                replacement.storage_path
            ))
        })?;
        if source.len() != replacement.payload.len() {
            return Err(invalid(format!(
                "XLDM replacement {} changes allocation size",
                replacement.storage_path
            )));
        }
        indexes.push(index);
    }

    let mut result = Vec::new();
    result
        .try_reserve_exact(storage.bytes.len())
        .map_err(|source| allocation("MS-XLDM rewritten storage bytes", source))?;
    result.extend_from_slice(storage.bytes);
    for (replacement, index) in replacements.iter().zip(indexes) {
        let entry = storage
            .files
            .get(index)
            .ok_or_else(|| invalid("XLDM replacement index is out of range"))?;
        let start = checked_usize(entry.offset.0, "file offset")?;
        let payload_len = checked_usize(entry.stored_size.0, "file size")?
            .checked_sub(CRC_SIZE)
            .ok_or_else(|| limit("file size"))?;
        let end = start
            .checked_add(payload_len)
            .ok_or_else(|| limit("payload range"))?;
        if end > result.len() {
            return Err(invalid("XLDM replacement payload range is outside storage"));
        }
        result[start..end].copy_from_slice(replacement.payload);
        let marker = crc32(replacement.payload);
        let marker = marker.to_le_bytes();
        result[end..end + CRC_SIZE].copy_from_slice(&marker);
    }
    // This reparses the candidate before the caller inspects inner sections;
    // no partially modified buffer escapes if the outer CRC or directory is
    // invalid.
    inspect(&result)?;
    Ok(result)
}

/// Rewrite already-admitted inner allocations while rebuilding the serial
/// files section and virtual directory for changed payload lengths. The
/// partition marker, backup-log marker, directory order, and member paths are
/// retained; callers use this primitive only after proving the typed closure
/// and must re-inspect every nested section before publishing the result.
#[cfg(test)]
pub(super) fn rewrite_variable_size_payloads(
    storage: &Storage<'_>,
    replacements: &[Xldm140FileReplacement<'_>],
) -> Result<Vec<u8>> {
    rewrite_variable_size_payloads_with_limit(storage, replacements, MAX_STORAGE_BYTES)
}

/// The length-only input to the variable-size rewrite planner.
///
/// Keeping this separate from [`Xldm140FileReplacement`] lets a caller prove
/// the exact outer-stream size before it materializes any changed member
/// payload.  The path is source-bound and the length is the encoded payload
/// length, excluding the allocation CRC marker.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) struct VariablePayloadLength<'a> {
    pub storage_path: &'a str,
    pub payload_len: usize,
}

/// Errors from the length-only variable-storage preflight.  The codec keeps
/// its historical [`crate::error::Error`] for ordinary malformed input and
/// allocation failures, while the caller cap gets a distinct value so the
/// identity owner can preserve that diagnostic through its public seam.
#[derive(Debug)]
pub(super) enum VariableRewriteError {
    Codec(Error),
    CallerLimit { actual: usize, maximum: usize },
}

impl From<Error> for VariableRewriteError {
    fn from(error: Error) -> Self {
        Self::Codec(error)
    }
}

/// Preflight a variable-size rewrite without allocating changed member
/// payloads or the rewritten storage buffer.  The returned length is the
/// exact page-aligned outer-stream length that the writer will emit.
pub(super) fn preflight_variable_size_payloads(
    storage: &Storage<'_>,
    replacements: &[VariablePayloadLength<'_>],
    max_output_bytes: usize,
) -> std::result::Result<usize, VariableRewriteError> {
    if replacements.is_empty() {
        return Ok(storage.bytes.len());
    }
    let output_len = plan_variable_size_payloads(storage, replacements)?.output_len;
    if output_len > max_output_bytes {
        return Err(VariableRewriteError::CallerLimit {
            actual: output_len,
            maximum: max_output_bytes,
        });
    }
    Ok(output_len)
}

/// Rewrite already-admitted variable-size allocations under an explicit
/// caller output cap.  The exact plan is computed before backup-log,
/// directory, or full-storage output buffers are materialized.
pub(super) fn rewrite_variable_size_payloads_with_limit(
    storage: &Storage<'_>,
    replacements: &[Xldm140FileReplacement<'_>],
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    if storage.profile != StorageProfile::Xldm140 {
        return Err(Error::Unsupported {
            feature: "variable-size inner XLDM rewrites require the version-140 profile",
        });
    }
    if storage.bytes.len() > MAX_STORAGE_BYTES {
        return Err(limit("storage bytes"));
    }
    if replacements.is_empty() {
        return Ok(storage.bytes.to_vec());
    }
    if replacements.len() > MAX_FILES {
        return Err(limit("replacement count"));
    }

    let mut lengths = Vec::new();
    lengths
        .try_reserve_exact(replacements.len())
        .map_err(|source| allocation("MS-XLDM replacement lengths", source))?;
    for replacement in replacements {
        lengths.push(VariablePayloadLength {
            storage_path: replacement.storage_path,
            payload_len: replacement.payload.len(),
        });
    }
    let plan = plan_variable_size_payloads(storage, &lengths)?;
    if plan.output_len > max_output_bytes {
        return Err(limit("rewritten storage caller output bytes"));
    }

    let mut replacement_paths = std::collections::HashSet::new();
    replacement_paths
        .try_reserve(replacements.len())
        .map_err(|source| allocation("MS-XLDM replacement paths", source))?;
    let mut replacement_payloads: Vec<Option<&[u8]>> = Vec::new();
    replacement_payloads
        .try_reserve_exact(storage.files.len())
        .map_err(|source| allocation("MS-XLDM replacement map", source))?;
    replacement_payloads.resize(storage.files.len(), None);
    for replacement in replacements {
        if !replacement_paths.insert(replacement.storage_path) {
            return Err(invalid(format!(
                "duplicate XLDM replacement path {}",
                replacement.storage_path
            )));
        }
        let index = storage
            .files
            .iter()
            .position(|entry| entry.path == replacement.storage_path)
            .ok_or_else(|| {
                invalid(format!(
                    "XLDM replacement path {} is absent from the source directory",
                    replacement.storage_path
                ))
            })?;
        if index == plan.first || index == plan.last {
            return Err(invalid(format!(
                "XLDM replacement {} would rewrite an outer marker allocation",
                replacement.storage_path
            )));
        }
        replacement_payloads[index] = Some(replacement.payload);
    }

    let generated_backup_payload = if plan.generated_backup_len.is_some() {
        let source_backup = storage
            .file_payload(plan.last)
            .ok_or_else(|| invalid("cannot resolve XLDM source backup-log payload"))?;
        Some(rewrite_backup_log_sizes(
            source_backup,
            storage,
            &replacement_payloads,
        )?)
    } else {
        None
    };
    replacement_payloads[plan.last] = generated_backup_payload.as_deref();

    let data_offset = checked_usize(storage.header.data_offset.0, "data offset")?;
    if data_offset != XLDM_PAGE_SIZE {
        return Err(invalid(
            "XLDM variable rewrite requires a canonical data offset",
        ));
    }
    let updated = &plan.updated;
    let directory_offset = plan.directory_offset;
    let directory_xml_len = plan.directory_xml_len;
    let directory_bytes_len = plan.directory_bytes_len;
    let output_len = plan.output_len;
    let header_xml = rewritten_header_xml(storage, directory_offset, directory_bytes_len)?;
    let header_bytes = encode_xml(&header_xml, storage.header_encoding)?;
    let header_start = 2usize
        .checked_add(encoded_utf16_len(XLDM_STREAM_SIGNATURE)?)
        .ok_or_else(|| limit("header XML offset"))?;
    let header_end = header_start
        .checked_add(header_bytes.len())
        .ok_or_else(|| limit("header XML range"))?;
    if header_end > data_offset {
        return Err(limit("rewritten header page bytes"));
    }
    let directory_xml = build_directory_xml(storage, updated, directory_xml_len)?;
    let directory_bytes = encode_xml(&directory_xml, storage.directory_encoding)?;
    if directory_bytes.len() != directory_bytes_len {
        return Err(invalid(
            "rewritten directory size preflight disagrees with encoding",
        ));
    }

    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("MS-XLDM rewritten storage bytes", source))?;
    result.resize(data_offset, 0);
    result[..BOM.len()].copy_from_slice(&BOM);
    let signature = encode_xml(XLDM_STREAM_SIGNATURE, XmlEncoding::Utf16Le)?;
    let signature_end = 2usize
        .checked_add(signature.len())
        .ok_or_else(|| limit("header signature range"))?;
    result[2..signature_end].copy_from_slice(&signature);
    result[header_start..header_end].copy_from_slice(&header_bytes);
    result.extend_from_slice(&BOM);
    for (position, index) in plan.order.iter().copied().enumerate() {
        if position == plan.order.len() - 1 {
            result.extend_from_slice(&BOM);
        }
        let (expected_start, stored_size) =
            updated[index].ok_or_else(|| invalid("missing rewritten directory entry"))?;
        let expected_start = checked_usize(expected_start, "rewritten file offset")?;
        if result.len() != expected_start {
            return Err(invalid("rewritten allocation order is not serial"));
        }
        let payload = replacement_payloads[index]
            .or_else(|| storage.file_payload(index))
            .ok_or_else(|| {
                invalid(format!(
                    "cannot resolve XLDM source payload {}",
                    storage.files[index].path
                ))
            })?;
        result.extend_from_slice(payload);
        result.extend_from_slice(&crc32(payload).to_le_bytes());
        if result.len()
            != expected_start
                .checked_add(checked_usize(stored_size, "rewritten file size")?)
                .ok_or_else(|| limit("rewritten allocation range"))?
        {
            return Err(invalid(
                "rewritten allocation size disagrees with its preflight",
            ));
        }
    }
    if result.len() > directory_offset {
        return Err(invalid(
            "rewritten allocation section exceeds directory offset",
        ));
    }
    result.resize(directory_offset, 0);
    result.extend_from_slice(&directory_bytes);
    result.resize(output_len, 0);
    inspect(&result)?;
    Ok(result)
}

struct VariableRewritePlan {
    order: Vec<usize>,
    first: usize,
    last: usize,
    generated_backup_len: Option<usize>,
    updated: Vec<Option<(u64, u64)>>,
    directory_offset: usize,
    directory_xml_len: usize,
    directory_bytes_len: usize,
    output_len: usize,
}

fn plan_variable_size_payloads(
    storage: &Storage<'_>,
    replacements: &[VariablePayloadLength<'_>],
) -> Result<VariableRewritePlan> {
    if storage.profile != StorageProfile::Xldm140 {
        return Err(Error::Unsupported {
            feature: "variable-size inner XLDM rewrites require the version-140 profile",
        });
    }
    if storage.bytes.len() > MAX_STORAGE_BYTES {
        return Err(limit("storage bytes"));
    }
    if replacements.is_empty() {
        return Err(invalid("variable rewrite plan has no replacements"));
    }
    if replacements.len() > MAX_FILES {
        return Err(limit("replacement count"));
    }

    let mut order = Vec::new();
    order
        .try_reserve_exact(storage.files.len())
        .map_err(|source| allocation("MS-XLDM allocation order", source))?;
    order.extend(0..storage.files.len());
    order.sort_by_key(|index| storage.files[*index].offset);
    if order.len() < 2 {
        return Err(invalid(
            "XLDM storage has no partition and backup-log allocation pair",
        ));
    }
    let first = order[0];
    let last = *order
        .last()
        .ok_or_else(|| invalid("XLDM allocation order is empty"))?;

    let mut paths = std::collections::HashSet::new();
    paths
        .try_reserve(replacements.len())
        .map_err(|source| allocation("MS-XLDM replacement paths", source))?;
    let mut payload_lengths: Vec<Option<usize>> = Vec::new();
    payload_lengths
        .try_reserve_exact(storage.files.len())
        .map_err(|source| allocation("MS-XLDM replacement length map", source))?;
    payload_lengths.resize(storage.files.len(), None);
    for replacement in replacements {
        if !paths.insert(replacement.storage_path) {
            return Err(invalid(format!(
                "duplicate XLDM replacement path {}",
                replacement.storage_path
            )));
        }
        if replacement.payload_len > MAX_STORAGE_BYTES {
            return Err(limit("replacement payload bytes"));
        }
        let index = storage
            .files
            .iter()
            .position(|entry| entry.path == replacement.storage_path)
            .ok_or_else(|| {
                invalid(format!(
                    "XLDM replacement path {} is absent from the source directory",
                    replacement.storage_path
                ))
            })?;
        if index == first || index == last {
            return Err(invalid(format!(
                "XLDM replacement {} would rewrite an outer marker allocation",
                replacement.storage_path
            )));
        }
        payload_lengths[index] = Some(replacement.payload_len);
    }

    let has_size_change = replacements.iter().any(|replacement| {
        storage
            .files
            .iter()
            .position(|entry| entry.path == replacement.storage_path)
            .and_then(|index| storage.file_payload(index))
            .is_some_and(|source| source.len() != replacement.payload_len)
    });
    let generated_backup_len = if has_size_change {
        let source_backup = storage
            .file_payload(last)
            .ok_or_else(|| invalid("cannot resolve XLDM source backup-log payload"))?;
        Some(rewrite_backup_log_sizes_len(
            source_backup,
            storage,
            &payload_lengths,
        )?)
    } else {
        None
    };
    payload_lengths[last] = generated_backup_len;

    let data_offset = checked_usize(storage.header.data_offset.0, "data offset")?;
    if data_offset != XLDM_PAGE_SIZE {
        return Err(invalid(
            "XLDM variable rewrite requires a canonical data offset",
        ));
    }
    let mut updated: Vec<Option<(u64, u64)>> = Vec::new();
    updated
        .try_reserve_exact(storage.files.len())
        .map_err(|source| allocation("MS-XLDM updated directory entries", source))?;
    updated.resize(storage.files.len(), None);
    let mut cursor = data_offset
        .checked_add(BOM.len())
        .ok_or_else(|| limit("rewritten data range"))?;
    for (position, index) in order.iter().copied().enumerate() {
        if position == order.len() - 1 {
            cursor = cursor
                .checked_add(BOM.len())
                .ok_or_else(|| limit("rewritten backup-log marker"))?;
        }
        let payload_len = payload_lengths[index]
            .or_else(|| storage.file_payload(index).map(<[u8]>::len))
            .ok_or_else(|| {
                invalid(format!(
                    "cannot resolve XLDM source payload {}",
                    storage.files[index].path
                ))
            })?;
        let stored_size = payload_len
            .checked_add(CRC_SIZE)
            .ok_or_else(|| limit("rewritten allocation size"))?;
        let start = cursor;
        cursor = cursor
            .checked_add(stored_size)
            .ok_or_else(|| limit("rewritten data range"))?;
        updated[index] = Some((
            u64::try_from(start).map_err(|_source| limit("rewritten file offset"))?,
            u64::try_from(stored_size).map_err(|_source| limit("rewritten file size"))?,
        ));
    }
    let directory_offset = align_page(cursor)?;
    let (directory_xml_len, directory_bytes_len) =
        directory_xml_lengths(storage, storage.directory_encoding, &updated)?;
    if directory_bytes_len == 0 || directory_bytes_len > MAX_DIRECTORY_BYTES {
        return Err(limit("rewritten directory bytes"));
    }
    let directory_end = directory_offset
        .checked_add(directory_bytes_len)
        .ok_or_else(|| limit("rewritten directory range"))?;
    let output_len = align_page(directory_end)?;
    if output_len > MAX_STORAGE_BYTES {
        return Err(limit("rewritten storage bytes"));
    }
    Ok(VariableRewritePlan {
        order,
        first,
        last,
        generated_backup_len,
        updated,
        directory_offset,
        directory_xml_len,
        directory_bytes_len,
        output_len,
    })
}

fn align_page(value: usize) -> Result<usize> {
    let remainder = value % XLDM_PAGE_SIZE;
    if remainder == 0 {
        Ok(value)
    } else {
        value
            .checked_add(XLDM_PAGE_SIZE - remainder)
            .ok_or_else(|| limit("page-aligned size"))
    }
}

fn encoded_utf16_len(value: &str) -> Result<usize> {
    value
        .encode_utf16()
        .count()
        .checked_mul(2)
        .ok_or_else(|| limit("UTF-16 XML bytes"))
}

fn directory_xml_lengths(
    storage: &Storage<'_>,
    encoding: XmlEncoding,
    updated: &[Option<(u64, u64)>],
) -> Result<(usize, usize)> {
    // The directory consists solely of the deterministic fields emitted by
    // build_directory_xml. Count both UTF-8 bytes (the String construction
    // capacity) and UTF-16 code units (the encoded member size) from the same
    // scalar stream. This keeps the preflight exact without constructing a
    // scratch directory String for UTF-16 sources.
    let fields = [
        "<BackupFile><Path>",
        "</Path><Size>",
        "</Size><m_cbOffsetHeader>",
        "</m_cbOffsetHeader><Delete>",
        "</Delete><CreatedTimestamp>",
        "</CreatedTimestamp><Access>",
        "</Access><LastWriteTime>",
        "</LastWriteTime></BackupFile>",
    ];
    let mut utf8_len = "<VirtualDirectory>".len();
    let mut utf16_units = "<VirtualDirectory>".encode_utf16().count();
    for (index, entry) in storage.files.iter().enumerate() {
        let (offset, stored_size) =
            updated[index].ok_or_else(|| invalid("missing rewritten directory entry"))?;
        for field in fields {
            let field_len = field.len();
            utf8_len = utf8_len
                .checked_add(field_len)
                .ok_or_else(|| limit("directory XML bytes"))?;
            utf16_units = utf16_units
                .checked_add(field.encode_utf16().count())
                .ok_or_else(|| limit("directory XML UTF-16 units"))?;
        }
        add_directory_value_lengths(
            &mut utf8_len,
            &mut utf16_units,
            xml_text_len(entry.path.as_str())?,
            xml_text_utf16_len(entry.path.as_str())?,
        )?;
        for value in [
            stored_size,
            offset,
            entry.created_timestamp.unsigned_abs(),
            entry.access_timestamp.unsigned_abs(),
            entry.last_write_timestamp.unsigned_abs(),
        ] {
            let length = decimal_len_u64(value);
            add_directory_value_lengths(&mut utf8_len, &mut utf16_units, length, length)?;
        }
        let delete_len = if entry.delete { 4 } else { 5 };
        add_directory_value_lengths(&mut utf8_len, &mut utf16_units, delete_len, delete_len)?;
        for value in [
            entry.created_timestamp,
            entry.access_timestamp,
            entry.last_write_timestamp,
        ] {
            if value < 0 {
                add_directory_value_lengths(&mut utf8_len, &mut utf16_units, 1, 1)?;
            }
        }
    }
    add_directory_value_lengths(
        &mut utf8_len,
        &mut utf16_units,
        "</VirtualDirectory>".len(),
        "</VirtualDirectory>".encode_utf16().count(),
    )?;
    let encoded_len = match encoding {
        XmlEncoding::Utf8 => utf8_len,
        XmlEncoding::Utf16Le => utf16_units
            .checked_mul(2)
            .ok_or_else(|| limit("directory XML UTF-16 bytes"))?,
    };
    Ok((utf8_len, encoded_len))
}

fn add_directory_value_lengths(
    utf8_len: &mut usize,
    utf16_units: &mut usize,
    utf8_add: usize,
    utf16_add: usize,
) -> Result<()> {
    *utf8_len = (*utf8_len)
        .checked_add(utf8_add)
        .ok_or_else(|| limit("directory XML bytes"))?;
    *utf16_units = (*utf16_units)
        .checked_add(utf16_add)
        .ok_or_else(|| limit("directory XML UTF-16 units"))?;
    Ok(())
}

fn decimal_len_u64(value: u64) -> usize {
    if value == 0 {
        return 1;
    }
    let mut value = value;
    let mut length = 0;
    while value != 0 {
        value /= 10;
        length += 1;
    }
    length
}

fn xml_text_utf16_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let addition = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf16(),
        };
        length
            .checked_add(addition)
            .ok_or_else(|| limit("XML escaped UTF-16 units"))
    })
}

fn xml_text_len(value: &str) -> Result<usize> {
    let mut length = 0usize;
    for character in value.chars() {
        let addition = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length = length
            .checked_add(addition)
            .ok_or_else(|| limit("XML escaped text bytes"))?;
    }
    Ok(length)
}

fn append_xml_text(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn build_directory_xml(
    storage: &Storage<'_>,
    updated: &[Option<(u64, u64)>],
    expected_len: usize,
) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve_exact(expected_len)
        .map_err(|source| allocation("MS-XLDM rewritten directory XML", source))?;
    output.push_str("<VirtualDirectory>");
    for (index, entry) in storage.files.iter().enumerate() {
        let (offset, stored_size) =
            updated[index].ok_or_else(|| invalid("missing rewritten directory entry"))?;
        output.push_str("<BackupFile><Path>");
        append_xml_text(&mut output, &entry.path);
        output.push_str("</Path><Size>");
        output.push_str(&stored_size.to_string());
        output.push_str("</Size><m_cbOffsetHeader>");
        output.push_str(&offset.to_string());
        output.push_str("</m_cbOffsetHeader><Delete>");
        output.push_str(if entry.delete { "true" } else { "false" });
        output.push_str("</Delete><CreatedTimestamp>");
        output.push_str(&entry.created_timestamp.to_string());
        output.push_str("</CreatedTimestamp><Access>");
        output.push_str(&entry.access_timestamp.to_string());
        output.push_str("</Access><LastWriteTime>");
        output.push_str(&entry.last_write_timestamp.to_string());
        output.push_str("</LastWriteTime></BackupFile>");
    }
    output.push_str("</VirtualDirectory>");
    if output.len() < expected_len {
        return Err(invalid(
            "rewritten directory XML size preflight undercounted",
        ));
    }
    Ok(output)
}

fn rewritten_header_xml(
    storage: &Storage<'_>,
    directory_offset: usize,
    directory_size: usize,
) -> Result<String> {
    let header = &storage.header;
    let output = format!(
        "<BackupLog><BackupRestoreSyncVersion>{}</BackupRestoreSyncVersion><Fault>false</Fault><faultcode>{}</faultcode><ErrorCode>true</ErrorCode><EncryptionFlag>false</EncryptionFlag><EncryptionKey>{}</EncryptionKey><ApplyCompression>true</ApplyCompression><m_cbOffsetHeader>{directory_offset}</m_cbOffsetHeader><DataSize>{directory_size}</DataSize><Files>{}</Files><ObjectID>{}</ObjectID><m_cbOffsetData>{}</m_cbOffsetData></BackupLog>",
        header.backup_restore_sync_version,
        header.fault_code,
        header.encryption_key_version,
        header.file_count,
        header.object_id,
        header.data_offset.0,
    );
    Ok(output)
}

fn encode_xml(value: &str, encoding: XmlEncoding) -> Result<Vec<u8>> {
    let length = match encoding {
        XmlEncoding::Utf8 => value.len(),
        XmlEncoding::Utf16Le => encoded_utf16_len(value)?,
    };
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| allocation("MS-XLDM XML bytes", source))?;
    match encoding {
        XmlEncoding::Utf8 => output.extend_from_slice(value.as_bytes()),
        XmlEncoding::Utf16Le => {
            for unit in value.encode_utf16() {
                output.extend_from_slice(&unit.to_le_bytes());
            }
        },
    }
    Ok(output)
}

fn rewrite_backup_log_sizes_len(
    source: &[u8],
    storage: &Storage<'_>,
    replacements: &[Option<usize>],
) -> Result<usize> {
    let (xml, encoding) = decode_marker_xml(source, StorageProfile::Xldm140)?;
    let mut length = xml.len();
    let mut changed = false;
    for (index, payload_len) in replacements.iter().enumerate() {
        let Some(payload_len) = payload_len else {
            continue;
        };
        if index >= storage.files.len() {
            return Err(invalid("backup-log size rewrite index is out of range"));
        }
        let source_payload = storage
            .file_payload(index)
            .ok_or_else(|| invalid("cannot resolve source payload for backup-log size rewrite"))?;
        if source_payload.len() == *payload_len {
            continue;
        }
        let path = storage.files[index].path.as_str();
        let size = i32::try_from(*payload_len)
            .map_err(|_source| limit("backup-log logged file size"))?
            .to_string();
        let (value_start, value_end) = backup_file_size_span(&xml, path)?.ok_or_else(|| {
            invalid(format!(
                "backup log has no unique FileList member for rewritten path {path}"
            ))
        })?;
        length = length
            .checked_sub(value_end - value_start)
            .and_then(|value| value.checked_add(size.len()))
            .ok_or_else(|| limit("backup-log XML bytes"))?;
        changed = true;
    }
    if !changed {
        return Ok(source.len());
    }
    match encoding {
        XmlEncoding::Utf8 => Ok(length),
        XmlEncoding::Utf16Le => {
            let delta = length as isize - xml.len() as isize;
            let delta_bytes = delta
                .checked_mul(2)
                .ok_or_else(|| limit("backup-log UTF-16 XML bytes"))?;
            if delta_bytes.is_negative() {
                source
                    .len()
                    .checked_sub(delta_bytes.unsigned_abs())
                    .ok_or_else(|| limit("backup-log UTF-16 XML bytes"))
            } else {
                source
                    .len()
                    .checked_add(delta_bytes as usize)
                    .ok_or_else(|| limit("backup-log UTF-16 XML bytes"))
            }
        },
    }
}

fn rewrite_backup_log_sizes(
    source: &[u8],
    storage: &Storage<'_>,
    replacements: &[Option<&[u8]>],
) -> Result<Vec<u8>> {
    let (mut xml, encoding) = decode_marker_xml(source, StorageProfile::Xldm140)?;
    let mut changed = false;
    for (index, payload) in replacements.iter().enumerate() {
        let Some(payload) = payload else {
            continue;
        };
        if index >= storage.files.len() {
            return Err(invalid("backup-log size rewrite index is out of range"));
        }
        let source_payload = storage
            .file_payload(index)
            .ok_or_else(|| invalid("cannot resolve source payload for backup-log size rewrite"))?;
        if source_payload.len() == payload.len() {
            continue;
        }
        let path = storage.files[index].path.as_str();
        let size = i32::try_from(payload.len())
            .map_err(|_source| limit("backup-log logged file size"))?
            .to_string();
        if !replace_backup_file_size(&mut xml, path, &size)? {
            return Err(invalid(format!(
                "backup log has no unique FileList member for rewritten path {path}"
            )));
        }
        changed = true;
    }
    if !changed {
        return Ok(source.to_vec());
    }
    encode_xml(&xml, encoding)
}

fn replace_backup_file_size(xml: &mut String, path: &str, size: &str) -> Result<bool> {
    let Some((value_start, value_end)) = backup_file_size_span(xml, path)? else {
        return Ok(false);
    };
    xml.replace_range(value_start..value_end, size);
    Ok(true)
}

fn backup_file_size_span(xml: &str, path: &str) -> Result<Option<(usize, usize)>> {
    let mut cursor = 0usize;
    let mut found = 0usize;
    let mut replacement = None;
    while let Some(relative) = xml[cursor..].find("<BackupFile") {
        let start = cursor + relative;
        let start_end = xml[start..]
            .find('>')
            .map(|offset| start + offset)
            .ok_or_else(|| invalid("backup-log BackupFile start tag is unclosed"))?;
        let close = xml[start_end + 1..]
            .find("</BackupFile>")
            .map(|offset| start_end + 1 + offset)
            .ok_or_else(|| invalid("backup-log BackupFile element is unclosed"))?;
        let block = &xml[start..close];
        let storage_start = block
            .find("<StoragePath>")
            .ok_or_else(|| invalid("backup-log BackupFile has no StoragePath"))?;
        let storage_value_start = start + storage_start + "<StoragePath>".len();
        let storage_value_end = xml[storage_value_start..]
            .find("</StoragePath>")
            .map(|offset| storage_value_start + offset)
            .ok_or_else(|| invalid("backup-log StoragePath is unclosed"))?;
        let decoded = quick_xml::escape::unescape(&xml[storage_value_start..storage_value_end])
            .map_err(|error| invalid(format!("backup-log StoragePath is not XML: {error}")))?;
        if decoded == path {
            found = found
                .checked_add(1)
                .ok_or_else(|| limit("backup-log size matches"))?;
            let size_relative = block
                .find("<Size>")
                .ok_or_else(|| invalid("backup-log BackupFile has no Size"))?;
            let value_start = start + size_relative + "<Size>".len();
            let value_end = xml[value_start..]
                .find("</Size>")
                .map(|offset| value_start + offset)
                .ok_or_else(|| invalid("backup-log Size is unclosed"))?;
            if replacement.is_some() {
                return Err(invalid(format!(
                    "backup log contains duplicate FileList member {path}"
                )));
            }
            replacement = Some((value_start, value_end));
        }
        cursor = close + "</BackupFile>".len();
    }
    let Some((value_start, value_end)) = replacement else {
        return Ok(None);
    };
    if found != 1 {
        return Err(invalid(format!(
            "backup log contains ambiguous FileList member {path}"
        )));
    }
    let raw_size = &xml[value_start..value_end];
    if raw_size.trim() != raw_size {
        return Err(invalid(
            "backup-log Size value has unsupported surrounding whitespace",
        ));
    }
    Ok(Some((value_start, value_end)))
}

pub(super) fn bounded_leaf(node: &Node, label: &str) -> Result<String> {
    let value = leaf_text(node)?;
    if value.len() > MAX_PATH_BYTES {
        return Err(limit(label));
    }
    Ok(value.to_owned())
}

fn parse_header(root: &Node) -> Result<Header> {
    let names = [
        "BackupRestoreSyncVersion",
        "Fault",
        "faultcode",
        "ErrorCode",
        "EncryptionFlag",
        "EncryptionKey",
        "ApplyCompression",
        "m_cbOffsetHeader",
        "DataSize",
        "Files",
        "ObjectID",
        "m_cbOffsetData",
    ];
    let values = exact_children(root, "BackupLog", &names)?;
    let backup_restore_sync_version = i32_value(values[0])?;
    if !matches!(backup_restore_sync_version, 140 | 150) {
        return Err(Error::Unsupported {
            feature: "XLDM header version",
        });
    }
    if bool_value(values[1])? {
        return Err(invalid("header Fault must be false"));
    }
    let fault_code = u32_value(values[2])?;
    if !bool_value(values[3])? {
        return Err(invalid("header ErrorCode must be true"));
    }
    if bool_value(values[4])? {
        return Err(invalid("header EncryptionFlag must be false"));
    }
    let encryption_key_version = i32_value(values[5])?;
    if !bool_value(values[6])? {
        return Err(invalid("header ApplyCompression must be true"));
    }
    let directory_offset = Offset(u64_value(values[7])?);
    let directory_size = Size(u64_value(values[8])?);
    let file_count = u32_value(values[9])?;
    if file_count as usize > MAX_FILES {
        return Err(limit("file count"));
    }
    let object_id = leaf_text(values[10])?.to_owned();
    if !valid_upper_guid(&object_id) {
        return Err(invalid("header ObjectID must be an uppercase UUID"));
    }
    let data_offset = Offset(u64_value(values[11])?);
    Ok(Header {
        backup_restore_sync_version,
        fault_code,
        encryption_key_version,
        compression: Compression::Xpress,
        directory_offset,
        directory_size,
        file_count,
        object_id,
        data_offset,
    })
}

fn parse_directory(root: &Node) -> Result<Vec<FileEntry>> {
    if root.name != "VirtualDirectory" || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid("expected attribute-free VirtualDirectory root"));
    }
    if root.children.len() > MAX_FILES {
        return Err(limit("file count"));
    }
    let names = [
        "Path",
        "Size",
        "m_cbOffsetHeader",
        "Delete",
        "CreatedTimestamp",
        "Access",
        "LastWriteTime",
    ];
    root.children
        .iter()
        .map(|child| {
            let values = exact_children(child, "BackupFile", &names)?;
            let path = leaf_text(values[0])?.to_owned();
            let stored_size = Size(u64_value(values[1])?);
            let offset = Offset(u64_value(values[2])?);
            let delete = bool_value(values[3])?;
            let created_timestamp = i64_value(values[4])?;
            let access_timestamp = i64_value(values[5])?;
            let last_write_timestamp = i64_value(values[6])?;
            let lower = path.to_ascii_lowercase();
            let kind = if lower.ends_with("cryptkey.bin") {
                FileKind::CryptographicKey
            } else if lower.ends_with(".xml") {
                FileKind::XmlMetadata
            } else {
                FileKind::OpaqueBinary
            };
            Ok(FileEntry {
                path,
                kind,
                offset,
                stored_size,
                crc32: 0,
                delete,
                created_timestamp,
                access_timestamp,
                last_write_timestamp,
            })
        })
        .collect()
}

fn parse_partitions(root: &Node, profile: StorageProfile) -> Result<usize> {
    if root.name != "Partitions" || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid("expected attribute-free Partitions root"));
    }
    if root.children.len() > MAX_PARTITIONS {
        return Err(limit("partition count"));
    }
    let names: &[&str] = match profile {
        StorageProfile::Xldm140 => &[
            "ObjectPath",
            "Name",
            "DataSize",
            "Location",
            "DataSourceID",
            "ConnectionString",
        ],
        StorageProfile::Tabular150 => &[
            "ObjectPath",
            "Name",
            "DataSize",
            "Location",
            "DataSourceID",
            "DataSourceName",
            "ConnectionString",
        ],
    };
    for partition in &root.children {
        let values = exact_children(partition, "Partition", names)?;
        let _ = i64_value(values[2])?;
        for value in &values {
            if leaf_text(value)?.len() > MAX_XML_TEXT_BYTES {
                return Err(limit("partition field bytes"));
            }
        }
    }
    Ok(root.children.len())
}

fn decode_marker_xml(bytes: &[u8], profile: StorageProfile) -> Result<(String, XmlEncoding)> {
    if profile == StorageProfile::Tabular150 {
        let bytes = bytes
            .strip_prefix(&BOM)
            .ok_or_else(|| invalid("tabular XML marker BOM is missing"))?;
        decode_xml(bytes, true)
    } else {
        decode_xml(bytes, false)
    }
}

fn payload_slice<'a>(bytes: &'a [u8], entry: &FileEntry) -> Result<&'a [u8]> {
    let start = checked_usize(entry.offset.0, "file offset")?;
    let size = checked_usize(entry.stored_size.0, "file size")?;
    let end = start
        .checked_add(size)
        .and_then(|value| value.checked_sub(CRC_SIZE))
        .ok_or_else(|| limit("payload range"))?;
    bytes
        .get(start..end)
        .ok_or_else(|| invalid("payload range is outside storage"))
}

fn parse_xml(xml: &str) -> Result<Node> {
    if xml.len() > MAX_DIRECTORY_BYTES {
        return Err(limit("XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    loop {
        let event = reader.read_event().map_err(xml_error)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_XML_NODES || stack.len() >= MAX_XML_DEPTH {
                    return Err(limit("XML structure"));
                }
                let empty = matches!(&event, Event::Empty(_));
                let node = make_node(element)?;
                if empty {
                    attach(node, &mut stack, &mut root)?;
                } else {
                    stack.push(node);
                }
            },
            Event::End(_) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML closing element"))?;
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                text_bytes = text_bytes
                    .checked_add(decoded.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_XML_TEXT_BYTES {
                    return Err(limit("XML text bytes"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .map(|value| value.to_string())
                    .or_else(|| match name.as_ref() {
                        "amp" => Some("&".into()),
                        "lt" => Some("<".into()),
                        "gt" => Some(">".into()),
                        "apos" => Some("'".into()),
                        "quot" => Some("\"".into()),
                        _ => None,
                    })
                    .ok_or_else(|| invalid("custom XML entity is rejected"))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| limit("XML text bytes"))?;
                if text_bytes > MAX_XML_TEXT_BYTES {
                    return Err(limit("XML text bytes"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity outside XML root"));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::CData(_) => {
                return Err(invalid(
                    "DTDs, processing instructions, and CDATA are rejected",
                ));
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated XML"));
    }
    root.ok_or_else(|| invalid("missing XML root"))
}

fn make_node(element: &BytesStart<'_>) -> Result<Node> {
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(xml_error)?;
        let key = attribute.key.as_ref();
        if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            attributes += 1;
        }
    }
    Ok(Node {
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}
fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}
pub(super) fn exact_children<'a>(
    root: &'a Node,
    root_name: &str,
    names: &[&str],
) -> Result<Vec<&'a Node>> {
    if root.name != root_name || root.attributes != 0 || !root.text.trim().is_empty() {
        return Err(invalid(format!(
            "expected attribute-free {root_name} element"
        )));
    }
    if root.children.len() != names.len() {
        return Err(invalid(format!("{root_name} has an invalid child count")));
    }
    for (child, expected) in root.children.iter().zip(names) {
        if child.name != *expected {
            return Err(invalid(format!("expected {expected} in {root_name}")));
        }
    }
    Ok(root.children.iter().collect())
}
pub(super) fn leaf_text(node: &Node) -> Result<&str> {
    if node.attributes != 0 || !node.children.is_empty() {
        return Err(invalid(format!(
            "{} must be an attribute-free leaf",
            node.name
        )));
    }
    Ok(&node.text)
}
pub(super) fn bool_value(node: &Node) -> Result<bool> {
    match leaf_text(node)?.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{} is not an XML boolean", node.name))),
    }
}
pub(super) fn u64_value(node: &Node) -> Result<u64> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not an unsigned 64-bit integer", node.name)))
}
pub(super) fn u32_value(node: &Node) -> Result<u32> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not an unsigned 32-bit integer", node.name)))
}
pub(super) fn i64_value(node: &Node) -> Result<i64> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not a signed 64-bit integer", node.name)))
}
pub(super) fn i32_value(node: &Node) -> Result<i32> {
    leaf_text(node)?
        .trim()
        .parse()
        .map_err(|_source| invalid(format!("{} is not a signed 32-bit integer", node.name)))
}

pub(super) fn decode_xml(bytes: &[u8], require_utf16: bool) -> Result<(String, XmlEncoding)> {
    if bytes.is_empty() {
        return Err(invalid("empty XML allocation"));
    }
    if bytes.starts_with(&BOM) {
        return Err(invalid("unexpected XML byte-order mark"));
    }
    let utf16 = require_utf16 || bytes.get(1) == Some(&0);
    if utf16 {
        let encoded_limit = MAX_DIRECTORY_BYTES
            .checked_mul(2)
            .ok_or_else(|| limit("XML bytes"))?;
        if bytes.len() > encoded_limit {
            return Err(limit("XML bytes"));
        }
        if !bytes.len().is_multiple_of(2) {
            return Err(invalid("odd-length UTF-16LE XML"));
        }
        let words = || {
            bytes
                .as_chunks::<2>()
                .0
                .iter()
                .map(|pair| u16::from_le_bytes(*pair))
        };
        let mut decoded_bytes = 0usize;
        for value in std::char::decode_utf16(words()) {
            let character = value.map_err(xml_error)?;
            decoded_bytes = decoded_bytes
                .checked_add(character.len_utf8())
                .ok_or_else(|| limit("XML bytes"))?;
            if decoded_bytes > MAX_DIRECTORY_BYTES {
                return Err(limit("XML bytes"));
            }
        }
        let mut output = String::new();
        output
            .try_reserve_exact(decoded_bytes)
            .map_err(|source| allocation("MS-XLDM XML bytes", source))?;
        for value in std::char::decode_utf16(words()) {
            output.push(value.map_err(xml_error)?);
        }
        Ok((output, XmlEncoding::Utf16Le))
    } else {
        if bytes.len() > MAX_DIRECTORY_BYTES {
            return Err(limit("XML bytes"));
        }
        Ok((
            std::str::from_utf8(bytes).map_err(xml_error)?.to_owned(),
            XmlEncoding::Utf8,
        ))
    }
}
pub(super) fn utf16le(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}
pub(super) fn valid_upper_guid(value: &str) -> bool {
    if value.len() != 36 {
        return false;
    }
    value.bytes().enumerate().all(|(index, byte)| {
        if matches!(index, 8 | 13 | 18 | 23) {
            byte == b'-'
        } else {
            byte.is_ascii_digit() || (b'A'..=b'F').contains(&byte)
        }
    })
}
pub(super) fn checked_usize(value: u64, name: &str) -> Result<usize> {
    usize::try_from(value).map_err(|_source| limit(name))
}
pub(super) fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for byte in bytes {
        let mut index = ((crc >> 24) ^ u32::from(*byte)) & 0xFF;
        let mut table = index << 24;
        for _ in 0..8 {
            table = if table & 0x8000_0000 != 0 {
                (table << 1) ^ 0x04C1_1DB7
            } else {
                table << 1
            };
        }
        index = table;
        crc = (crc << 8) ^ index;
    }
    crc
}
pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
pub(super) fn limit(name: &str) -> Error {
    invalid(format!("MS-XLDM {name} limit exceeded"))
}
