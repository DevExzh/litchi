//! Validation for caption semantic values and bounded table inputs.

use super::model::{AutoEntry, AutoTable, Definition, Heading, LabelTable, Location, Separator};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};

/// Word limits one caption label to 40 UTF-16 code units.
pub(crate) const MAX_LABEL_UNITS: usize = 40;
/// STTB stores each string length in an unsigned 16-bit field.
pub(crate) const MAX_STRING_UNITS: usize = u16::MAX as usize;
/// A hostile table must not cause an unbounded allocation during inspection.
pub(crate) const MAX_TABLE_BYTES: usize = 16 * 1024 * 1024;
/// Both caption STTB variants use a two-byte cData count.
pub(crate) const MAX_ENTRIES: usize = u16::MAX as usize;
pub(crate) const CAPTION_POINTER_BASE: usize = 154;

pub(crate) fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

pub(crate) fn location_from_raw(value: u8) -> Result<Location> {
    match value {
        0x0 => Ok(Location::Below),
        0x1 => Ok(Location::Above),
        _ => Err(corrupted(format!("invalid caption location 0x{value:X}"))),
    }
}

pub(crate) fn heading_from_raw(value: u8) -> Result<Heading> {
    match value {
        0x1 => Ok(Heading::Level1),
        0x2 => Ok(Heading::Level2),
        0x3 => Ok(Heading::Level3),
        0x4 => Ok(Heading::Level4),
        0x5 => Ok(Heading::Level5),
        0x6 => Ok(Heading::Level6),
        0x7 => Ok(Heading::Level7),
        0x8 => Ok(Heading::Level8),
        0x9 => Ok(Heading::Level9),
        _ => Err(corrupted(format!("invalid chapter heading 0x{value:X}"))),
    }
}

pub(crate) fn separator_from_raw(value: u16) -> Result<Separator> {
    match value {
        0x001E => Ok(Separator::Hyphen),
        0x002E => Ok(Separator::Period),
        0x003A => Ok(Separator::Colon),
        0x2013 => Ok(Separator::EnDash),
        0x2014 => Ok(Separator::EmDash),
        _ => Err(corrupted(format!(
            "invalid chapter number separator 0x{value:04X}"
        ))),
    }
}

pub(crate) fn validate_definition(value: &Definition) -> Result<()> {
    let units = value.label().encode_utf16().count();
    if units > MAX_LABEL_UNITS {
        return Err(corrupted("caption label exceeds 40 UTF-16 code units"));
    }
    Ok(())
}

pub(crate) fn validate_definitions(values: &[Definition]) -> Result<()> {
    if values.len() > MAX_ENTRIES {
        return Err(corrupted("SttbfCaption count exceeds 65535 entries"));
    }
    values.iter().try_for_each(validate_definition)
}

pub(crate) fn validate_auto_entry(value: &AutoEntry) -> Result<()> {
    if value.prog_id().encode_utf16().count() > MAX_STRING_UNITS {
        return Err(corrupted(
            "AutoCaption ProgID exceeds the STTB 16-bit string length",
        ));
    }
    Ok(())
}

pub(crate) fn validate_auto_entries(values: &[AutoEntry]) -> Result<()> {
    if values.len() > MAX_ENTRIES {
        return Err(corrupted("SttbfAutoCaption count exceeds 65535 entries"));
    }
    values.iter().try_for_each(validate_auto_entry)
}

/// Ensure every automatic-caption rule points into the label table it names.
pub(crate) fn validate_references(
    labels: Option<&LabelTable>,
    auto: Option<&AutoTable>,
) -> Result<()> {
    let Some(auto) = auto else {
        return Ok(());
    };
    let label_count = labels.map_or(0, LabelTable::len);
    for (index, entry) in auto.entries().iter().enumerate() {
        if usize::from(entry.caption_index()) >= label_count {
            return Err(corrupted(format!(
                "SttbfAutoCaption entry {index} references a missing SttbfCaption label"
            )));
        }
    }
    Ok(())
}

pub(crate) fn validate_table_size(data: &[u8], name: &str) -> Result<()> {
    if data.len() > MAX_TABLE_BYTES {
        return Err(corrupted(format!("{name} exceeds the table size cap")));
    }
    Ok(())
}

/// Validate the FIB/profile assumptions required by package edits.
pub(crate) fn package_fib(fib: &FileInformationBlock) -> Result<()> {
    if fib.version() < WORD_97_NFIB {
        return Err(PackageError::UnsupportedVersion {
            nfib: fib.version(),
            name: fib.version_name(),
        });
    }
    if fib.is_encrypted() {
        return Err(corrupted(
            "encrypted DOC packages cannot be edited by the caption owner",
        ));
    }
    if !fib.is_template() {
        return Err(corrupted(
            "caption tables are only writable in a Normal template",
        ));
    }
    if fib.table_pointer_count().is_none() {
        return Err(corrupted(
            "WordDocument FIB table-pointer array is truncated",
        ));
    }
    if fib.table_pointer_count().unwrap_or(0) <= super::codec::AUTO_CAPTION_FIB_INDEX {
        return Err(corrupted(
            "WordDocument FIB does not expose the caption table pointers",
        ));
    }
    Ok(())
}

/// Locate one caption FIB pointer pair in the `WordDocument` stream.
pub(crate) fn pointer_location(fib: &FileInformationBlock, index: usize) -> Result<usize> {
    package_fib(fib)?;
    let offset = CAPTION_POINTER_BASE
        .checked_add(
            index
                .checked_mul(8)
                .ok_or_else(|| corrupted("caption FIB pointer index multiplication overflows"))?,
        )
        .ok_or_else(|| corrupted("caption FIB pointer offset overflows"))?;
    let end = offset
        .checked_add(8)
        .ok_or_else(|| corrupted("caption FIB pointer range overflows"))?;
    if end > fib.raw_data().len() {
        return Err(corrupted(
            "WordDocument FIB does not contain the caption table pointer",
        ));
    }
    Ok(offset)
}
