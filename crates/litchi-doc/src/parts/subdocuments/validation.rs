//! Candidate and wire-limit validation for the master-document owner.

use super::codec::{
    FNIF_LEN, FNPI_NIL_IDENTIFIER, FNPI_TYPE_MAIL_MERGE, FNPI_TYPE_SUBDOCUMENT, MAX_TABLE_BYTES,
    WKB_FLAGS_IGNORED, WKB_FLAGS_REQUIRED, WKB_OUTLINE_LEVEL,
};
use super::model::Collection;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};
use std::collections::HashSet;

/// Word 97+ FIB pointer indexes owned by this package layer.
pub(super) const PLCF_WKB: usize = 54;
pub(super) const STTB_FNM: usize = 72;

/// Validate the FIB/table boundary before package publication.
pub(super) fn package_fib(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<()> {
    if fib.version() < WORD_97_NFIB {
        return Err(PackageError::UnsupportedVersion {
            nfib: fib.version(),
            name: fib.version_name(),
        });
    }
    if fib.is_encrypted() {
        return Err(corrupted(
            "encrypted DOC packages cannot be edited by the subdocument owner",
        ));
    }
    package_fib_shape(fib)?;
    if document_protected(fib, table_stream)? {
        return Err(corrupted(
            "protected DOC packages cannot be edited by the subdocument owner",
        ));
    }
    if crate::parts::protection::Ranges::parse(fib, table_stream)?.is_some() {
        return Err(corrupted(
            "DOC range-level protection is not bypassed by the subdocument owner",
        ));
    }
    Ok(())
}

pub(super) fn pointer_location(fib: &FileInformationBlock, index: usize) -> Result<usize> {
    package_fib_shape(fib)?;
    let offset = 154usize
        .checked_add(
            index
                .checked_mul(8)
                .ok_or_else(|| corrupted("FIB pointer offset overflows"))?,
        )
        .ok_or_else(|| corrupted("FIB pointer offset overflows"))?;
    let end = offset
        .checked_add(8)
        .ok_or_else(|| corrupted("FIB pointer range overflows"))?;
    if end > fib.raw_data().len() {
        return Err(corrupted("FIB pointer range exceeds WordDocument"));
    }
    Ok(offset)
}

fn package_fib_shape(fib: &FileInformationBlock) -> Result<()> {
    if fib.version() < WORD_97_NFIB {
        return Err(PackageError::UnsupportedVersion {
            nfib: fib.version(),
            name: fib.version_name(),
        });
    }
    let Some(count) = fib.table_pointer_count() else {
        return Err(corrupted(
            "WordDocument FIB table-pointer array is truncated",
        ));
    };
    if count <= STTB_FNM || count <= PLCF_WKB {
        return Err(corrupted(
            "WordDocument FIB does not expose the subdocument table pointers",
        ));
    }
    Ok(())
}

fn document_protected(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<bool> {
    // DOP is FibRgFcLcb97 pair 31. These are the protection indicators used
    // by the existing revision owner; this owner never authenticates or
    // bypasses any of them.
    const DOP: usize = 31;
    let Some((offset, length)) = fib.get_table_pointer(DOP) else {
        return Ok(false);
    };
    if length == 0 {
        return Ok(false);
    }
    let start = usize::try_from(offset).map_err(|_| corrupted("DOP offset exceeds usize"))?;
    let length = usize::try_from(length).map_err(|_| corrupted("DOP length exceeds usize"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("DOP range overflows"))?;
    let dop = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("DOP extends beyond the table stream"))?;
    if dop.len() < 84 {
        return Err(corrupted("DOP is truncated before its protection fields"));
    }
    Ok(dop[6] & 0x10 != 0
        || dop[7] & (0x02 | 0x20 | 0x40) != 0
        || dop[78..82].iter().any(|byte| *byte != 0))
}

pub(super) fn collection(value: &Collection, main_document_chars: u32) -> Result<()> {
    if value.referenced_files.len() > usize::from(u16::MAX) {
        return Err(corrupted("SttbFnm count exceeds its u16 field"));
    }
    let mut identifiers = HashSet::with_capacity(value.referenced_files.len());
    for file in &value.referenced_files {
        let fnpi = file.fnpi;
        if fnpi.file_type() != FNPI_TYPE_MAIL_MERGE && fnpi.file_type() != FNPI_TYPE_SUBDOCUMENT {
            return Err(corrupted(
                "SttbFnm FNIF fnpt is not a defined file name type",
            ));
        }
        if fnpi.identifier() == FNPI_NIL_IDENTIFIER {
            return Err(corrupted("SttbFnm FNIF fnpd is the reserved nil value"));
        }
        if !identifiers.insert((fnpi.file_type(), fnpi.identifier())) {
            return Err(corrupted("SttbFnm FNIF fnpi values must be unique"));
        }
        let units = fnpi_units(&file.path);
        if units > usize::from(u16::MAX) {
            return Err(corrupted(
                "SttbFnm file name exceeds its u16 UTF-16 length field",
            ));
        }
        if let Some(offset) = file.relative_path_offset
            && (offset >= units || offset > usize::from(u8::MAX) - 1)
        {
            return Err(corrupted(
                "SttbFnm FNIF ichRelative exceeds the file name length",
            ));
        }
        if file.is_non_file_system_path && (file.valid_on_fat || file.valid_on_ntfs) {
            return Err(corrupted(
                "SttbFnm FNIF fnfb marks a non-file-system path as FAT/NTFS valid",
            ));
        }
    }

    if value.subdocuments.len() > max_wkb_entries() {
        return Err(corrupted("PlcfWKB count exceeds its u32 byte-length field"));
    }
    let terminal_cp = main_document_chars
        .checked_add(2)
        .ok_or_else(|| corrupted("PlcfWKB terminal CP overflows"))?;
    let mut previous = None;
    for reference in &value.subdocuments {
        if reference.start >= main_document_chars
            || previous.is_some_and(|start| reference.start <= start)
        {
            return Err(corrupted("PlcfWKB CPs must be unique and increasing"));
        }
        previous = Some(reference.start);
        if reference.outline_level != WKB_OUTLINE_LEVEL {
            return Err(corrupted("WKB lvl is not the mandated outline level"));
        }
        if reference.file_name.file_type() != FNPI_TYPE_SUBDOCUMENT {
            return Err(corrupted("WKB fnpi does not reference a subdocument"));
        }
        let Some(file) = value
            .referenced_files
            .iter()
            .find(|file| file.fnpi == reference.file_name)
        else {
            return Err(corrupted("WKB fnpi has no matching SttbFnm entry"));
        };
        if file.kind() != super::model::Kind::Subdocument {
            return Err(corrupted("WKB fnpi resolves to a non-subdocument file"));
        }
        let raw_flags = reference.raw_flags;
        if raw_flags & !WKB_FLAGS_IGNORED != WKB_FLAGS_REQUIRED || raw_flags & 0xFF00 != 0 {
            return Err(corrupted("WKB reserved flags have invalid values"));
        }
    }
    let wkb_length = value
        .subdocuments
        .len()
        .checked_mul(16)
        .and_then(|length| length.checked_add(4))
        .ok_or_else(|| corrupted("PlcfWKB encoded length overflows"))?;
    if wkb_length > u32::MAX as usize || wkb_length > MAX_TABLE_BYTES {
        return Err(corrupted("PlcfWKB exceeds its bounded byte length"));
    }
    if value.subdocuments.is_empty() && terminal_cp < 2 {
        return Err(corrupted("PlcfWKB terminal CP is invalid"));
    }

    let fnm_length = value
        .referenced_files
        .iter()
        .try_fold(6usize, |size, file| {
            let string_bytes = fnpi_units(&file.path)
                .checked_mul(2)
                .and_then(|length| length.checked_add(2))
                .and_then(|length| length.checked_add(FNIF_LEN))
                .ok_or_else(|| corrupted("SttbFnm encoded length overflows"))?;
            size.checked_add(string_bytes)
                .ok_or_else(|| corrupted("SttbFnm encoded length overflows"))
        })?;
    if fnm_length > u32::MAX as usize || fnm_length > MAX_TABLE_BYTES {
        return Err(corrupted("SttbFnm exceeds its bounded byte length"));
    }
    Ok(())
}

fn fnpi_units(value: &str) -> usize {
    value.encode_utf16().count()
}

fn max_wkb_entries() -> usize {
    ((u32::MAX as usize) - 4) / 16
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
