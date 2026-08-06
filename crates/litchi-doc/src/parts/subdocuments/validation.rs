//! Candidate and wire-limit validation for the master-document owner.

use super::codec::{
    FNIF_LEN, FNPI_NIL_IDENTIFIER, FNPI_TYPE_MAIL_MERGE, FNPI_TYPE_SUBDOCUMENT, MAX_TABLE_BYTES,
    WKB_FLAGS_IGNORED, WKB_FLAGS_REQUIRED, WKB_OUTLINE_LEVEL,
};
use super::model::Collection;
use crate::package::{Error as PackageError, Result};
use std::collections::HashSet;

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
        if let Some(offset) = file.relative_path_offset {
            if offset >= units || offset > usize::from(u8::MAX) - 1 {
                return Err(corrupted(
                    "SttbFnm FNIF ichRelative exceeds the file name length",
                ));
            }
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
