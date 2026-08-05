//! MS-DOC `SttbFnm` and `PlcfWKB` codecs.

use super::model::{Collection, Name, Reference};
use crate::package::{DocError, Result};
use crate::parts::fib::{FileInformationBlock, WORD_97_NFIB};
use crate::parts::mail_merge::Fnpi;
use std::collections::HashSet;

/// Table-pointer index of `fcPlcfWkb`/`lcbPlcfWkb`.
pub(super) const PLCF_WKB: usize = 54;
/// Table-pointer index of `fcSttbFnm`/`lcbSttbFnm`.
pub(super) const STTB_FNM: usize = 72;

/// Size in bytes of one `WKB` element (MS-DOC 2.9.346).
const WKB_LEN: usize = 12;
/// `WKB.fn`: the mandated value.
pub(super) const WKB_FN: u16 = 0x0000;
/// `WKB` flag bits that are undefined and MUST be ignored: `fReserved3`
/// (bit 2) and `fReserved8` (bit 7).
const WKB_FLAGS_IGNORED: u16 = 0x0084;
/// The mandated value of the defined `WKB` flag bits: `fReserved6` (bit 5)
/// MUST be 1 and every other defined bit MUST be 0.
pub(super) const WKB_FLAGS_REQUIRED: u16 = 0x0020;
/// `WKB.fReserved9` occupies the high byte of the flags field and MUST be 0.
const WKB_RESERVED9_MASK: u16 = 0xFF00;
/// `WKB.lvl`: the mandated outline level.
pub(super) const WKB_OUTLINE_LEVEL: u16 = 0x0002;

/// `fExtend` marker of an extended STTB.
pub(super) const STTB_F_EXTEND: u16 = 0xFFFF;
/// `SttbFnm.cbExtra`: one `FNIF` per entry (MS-DOC 2.9.288).
pub(super) const STTB_FNM_CB_EXTRA: u16 = 8;
/// `FNIF.ichRelative` value meaning the file name carries no relative path.
const ICH_RELATIVE_NONE: u8 = 0xFF;
/// Size in bytes of the `FNIF` extra data (MS-DOC 2.9.92).
const FNIF_LEN: usize = 8;

/// `FNPI.fnpt` value for a mail merge data source file (MS-DOC 2.9.93).
const FNPI_TYPE_MAIL_MERGE: u8 = 0x3;
/// `FNPI.fnpt` value for a subdocument file (MS-DOC 2.9.93).
const FNPI_TYPE_SUBDOCUMENT: u8 = 0x5;
/// `FNPI.fnpd` value that is not a valid file name identifier.
const FNPI_NIL_IDENTIFIER: u16 = 0xFFF;

/// `FNFB` bit layout (MS-DOC 2.9.91).
pub(super) const FNFB_FAT: u8 = 0x01;
pub(super) const FNFB_NTFS: u8 = 0x08;
pub(super) const FNFB_NON_FILE_SYS: u8 = 0x10;

impl Collection {
    /// Parse the `PlcfWKB` and `SttbFnm` tables addressed by the FIB, or
    /// `None` when the document carries neither.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Collection>> {
        // The Word 6/95 FIB table-pointer layout assigns these indices to
        // unrelated structures, so they only carry this meaning from Word 97
        // on.
        if fib.version() < WORD_97_NFIB {
            return Ok(None);
        }
        let wkb_data = optional_slice(fib, table_stream, PLCF_WKB, "PlcfWKB")?;
        let fnm_data = optional_slice(fib, table_stream, STTB_FNM, "SttbFnm")?;
        if wkb_data.is_none() && fnm_data.is_none() {
            return Ok(None);
        }
        let referenced_files = fnm_data
            .map(parse_sttb_fnm)
            .transpose()?
            .unwrap_or_default();
        let subdocuments = match wkb_data {
            Some(data) => {
                if fnm_data.is_none() {
                    return Err(corrupted(
                        "PlcfWKB is present but the SttbFnm it references is missing",
                    ));
                }
                parse_plcf_wkb(data, fib.get_main_doc_range().1, &referenced_files)?
            },
            None => Vec::new(),
        };
        Ok(Some(Collection::from_parts(referenced_files, subdocuments)))
    }
}

/// Parse `SttbFnm` (MS-DOC 2.9.288): an extended STTB of full file paths,
/// each followed by an 8-byte `FNIF` (MS-DOC 2.9.92).
pub(super) fn parse_sttb_fnm(data: &[u8]) -> Result<Vec<Name>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbFnm fExtend")? != STTB_F_EXTEND
        || read_u16(data, 4, "SttbFnm cbExtra")? != STTB_FNM_CB_EXTRA
    {
        return Err(corrupted("SttbFnm has an invalid header"));
    }
    let count = usize::from(read_u16(data, 2, "SttbFnm cData")?);
    let payload_length = data.len() - 6;
    // Every entry has at least its two-byte cchData and eight-byte FNIF. This
    // rejects impossible counts before reserving and bounds allocation by the
    // encoded table length as well as cData.
    if count > payload_length / (2 + FNIF_LEN) {
        return Err(corrupted("SttbFnm cData exceeds the table length"));
    }
    let mut files = Vec::with_capacity(count);
    let mut identifiers = HashSet::with_capacity(count);
    let mut offset = 6usize;
    for _ in 0..count {
        let chars = usize::from(read_u16(data, offset, "SttbFnm cchData")?);
        offset = offset
            .checked_add(2)
            .ok_or_else(|| corrupted("SttbFnm file name range overflows"))?;
        let byte_length = chars
            .checked_mul(2)
            .ok_or_else(|| corrupted("SttbFnm file name range overflows"))?;
        let end = offset
            .checked_add(byte_length)
            .ok_or_else(|| corrupted("SttbFnm file name range overflows"))?;
        let bytes = data
            .get(offset..end)
            .ok_or_else(|| corrupted("SttbFnm file name is truncated"))?;
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        let path = String::from_utf16(&units)
            .map_err(|_| corrupted("SttbFnm file name is invalid UTF-16"))?;
        let fnif_end = end
            .checked_add(FNIF_LEN)
            .ok_or_else(|| corrupted("SttbFnm FNIF range overflows"))?;
        let fnif = data
            .get(end..fnif_end)
            .ok_or_else(|| corrupted("SttbFnm FNIF is truncated"))?;
        let fnpi = Fnpi::from_raw(u16::from_le_bytes([fnif[0], fnif[1]]));
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
        let relative_path_offset = match fnif[2] {
            ICH_RELATIVE_NONE => None,
            offset if usize::from(offset) < chars => Some(usize::from(offset)),
            _ => {
                return Err(corrupted(
                    "SttbFnm FNIF ichRelative exceeds the file name length",
                ));
            },
        };
        let fnfb = fnif[3];
        let is_non_file_system_path = fnfb & FNFB_NON_FILE_SYS != 0;
        if is_non_file_system_path && fnfb & (FNFB_FAT | FNFB_NTFS) != 0 {
            return Err(corrupted(
                "SttbFnm FNIF fnfb marks a non-file-system path as FAT/NTFS valid",
            ));
        }
        // FNIF `unused` is undefined and MUST be ignored.
        files.push(Name {
            fnpi,
            path,
            relative_path_offset,
            valid_on_fat: fnfb & FNFB_FAT != 0,
            valid_on_ntfs: fnfb & FNFB_NTFS != 0,
            is_non_file_system_path,
        });
        offset = fnif_end;
    }
    if offset != data.len() {
        return Err(corrupted("SttbFnm contains trailing bytes"));
    }
    Ok(files)
}

/// Parse `PlcfWKB` (MS-DOC 2.8.34): a PLC of 12-byte `WKB` elements
/// (MS-DOC 2.9.346), resolving each `fnpi` against the `SttbFnm` entries.
pub(super) fn parse_plcf_wkb(
    data: &[u8],
    main_document_chars: u32,
    referenced_files: &[Name],
) -> Result<Vec<Reference>> {
    if data.len() < 4 || !(data.len() - 4).is_multiple_of(4 + WKB_LEN) {
        return Err(corrupted("PlcfWKB has an invalid byte length"));
    }
    let count = (data.len() - 4) / (4 + WKB_LEN);
    let terminal_cp = main_document_chars
        .checked_add(2)
        .ok_or_else(|| corrupted("PlcfWKB terminal CP overflows"))?;
    let wkbs = (count + 1)
        .checked_mul(4)
        .ok_or_else(|| corrupted("PlcfWKB position bytes overflow"))?;
    let mut subdocuments = Vec::with_capacity(count);
    let mut previous = None;
    for index in 0..count {
        let start = read_u32(data, index * 4, "PlcfWKB CP")?;
        if start >= main_document_chars {
            return Err(corrupted("PlcfWKB CP is not within the main document"));
        }
        if previous.is_some_and(|cp| start <= cp) {
            return Err(corrupted("PlcfWKB CPs must be unique and increasing"));
        }
        previous = Some(start);

        let wkb = wkbs + index * WKB_LEN;
        if read_u16(data, wkb, "WKB fn")? != WKB_FN {
            return Err(corrupted("WKB fn is not zero"));
        }
        let flags = read_u16(data, wkb + 2, "WKB flags")?;
        if flags & WKB_RESERVED9_MASK != 0 || flags & !WKB_FLAGS_IGNORED != WKB_FLAGS_REQUIRED {
            return Err(corrupted("WKB reserved flags have invalid values"));
        }
        let outline_level = read_u16(data, wkb + 4, "WKB lvl")?;
        if outline_level != WKB_OUTLINE_LEVEL {
            return Err(corrupted("WKB lvl is not the mandated outline level"));
        }
        let file_name = Fnpi::from_raw(read_u16(data, wkb + 6, "WKB fnpi")?);
        if file_name.file_type() != FNPI_TYPE_SUBDOCUMENT {
            return Err(corrupted("WKB fnpi does not reference a subdocument"));
        }
        let file_name_index = referenced_files
            .iter()
            .position(|file| file.fnpi == file_name)
            .ok_or_else(|| corrupted("WKB fnpi has no matching SttbFnm entry"))?;
        if read_u32(data, wkb + 8, "WKB pdod")? != 0 {
            return Err(corrupted("WKB pdod is not zero"));
        }
        subdocuments.push(Reference {
            start,
            outline_level,
            file_name,
            file_name_index,
        });
    }
    if read_u32(data, count * 4, "PlcfWKB terminal CP")? != terminal_cp {
        return Err(corrupted(
            "PlcfWKB terminal CP is not the main document length plus two",
        ));
    }
    Ok(subdocuments)
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

fn read_u16(data: &[u8], offset: usize, name: &str) -> Result<u16> {
    let bytes = data
        .get(offset..offset.saturating_add(2))
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(data: &[u8], offset: usize, name: &str) -> Result<u32> {
    let bytes = data
        .get(offset..offset.saturating_add(4))
        .ok_or_else(|| corrupted(format!("{name} is truncated")))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}
