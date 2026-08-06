//! Wire and package-boundary validation for `SttbfAtnBkmk`.

use super::model::Tags;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};
use std::collections::HashSet;

/// FIB index of `fcSttbfAtnBkmk`/`lcbSttbfAtnBkmk`.
pub(crate) const FIB_INDEX: usize = 37;
/// Byte offset of the selected FIB pair in WordDocument.
pub(crate) const POINTER_OFFSET: usize = 154 + FIB_INDEX * 8;
/// `ATNBE.bmc` is the annotation bookmark class.
pub(crate) const BMC_ANNOTATION: u16 = 0x0100;
/// `ATNBE` size appended to each zero-length STTB string.
pub(crate) const ATNBE_SIZE: usize = 10;
/// Maximum `cData` permitted by MS-DOC §2.9.277.
pub(crate) const MAX_ENTRIES: usize = 0x3FFC;
/// Maximum complete payload size under the format count bound.
pub(crate) const MAX_TABLE_BYTES: usize = 6 + MAX_ENTRIES * (2 + ATNBE_SIZE);

pub(crate) fn tags(value: &Tags) -> Result<()> {
    if value.entries().len() > MAX_ENTRIES {
        return Err(corrupted("SttbfAtnBkmk cData exceeds 0x3FFC entries"));
    }
    let mut ids = HashSet::with_capacity(value.entries().len());
    for entry in value.entries() {
        if !ids.insert(entry.id()) {
            return Err(corrupted("SttbfAtnBkmk lTag values must be unique"));
        }
    }
    Ok(())
}

pub(crate) fn package_fib(fib: &FileInformationBlock) -> Result<()> {
    if fib.version() < WORD_97_NFIB {
        return Err(PackageError::UnsupportedVersion {
            nfib: fib.version(),
            name: fib.version_name(),
        });
    }
    if fib.is_encrypted() {
        return Err(corrupted(
            "encrypted DOC packages cannot be edited by the annotation-bookmark owner",
        ));
    }
    if fib.table_pointer_count().is_none() {
        return Err(corrupted(
            "WordDocument FIB table-pointer array is truncated",
        ));
    }
    if fib.table_pointer_count().unwrap_or(0) <= FIB_INDEX {
        return Err(corrupted(
            "WordDocument FIB does not expose fcSttbfAtnBkmk/lcbSttbfAtnBkmk",
        ));
    }
    Ok(())
}

pub(crate) fn pointer_location(fib: &FileInformationBlock) -> Result<usize> {
    package_fib(fib)?;
    let end = POINTER_OFFSET
        .checked_add(8)
        .ok_or_else(|| corrupted("SttbfAtnBkmk pointer range overflows"))?;
    if end > fib.raw_data().len() {
        return Err(corrupted(
            "WordDocument FIB does not contain fcSttbfAtnBkmk/lcbSttbfAtnBkmk",
        ));
    }
    Ok(POINTER_OFFSET)
}

pub(crate) fn table_range(table: &[u8], offset: u32, length: u32) -> Result<&[u8]> {
    let start =
        usize::try_from(offset).map_err(|_| corrupted("fcSttbfAtnBkmk offset exceeds usize"))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted("lcbSttbfAtnBkmk length exceeds usize"))?;
    if length > MAX_TABLE_BYTES {
        return Err(corrupted(
            "SttbfAtnBkmk exceeds its specification-derived size cap",
        ));
    }
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("SttbfAtnBkmk table range overflows"))?;
    table
        .get(start..end)
        .ok_or_else(|| corrupted("SttbfAtnBkmk extends beyond the table stream"))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
