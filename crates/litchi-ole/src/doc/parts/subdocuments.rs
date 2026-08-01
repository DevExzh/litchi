//! Master-document subdocument metadata: the `PlcfWKB` subdocument directory
//! and the `SttbFnm` referenced-file name table (MS-DOC 2.8.34, 2.9.288,
//! 2.9.346, 2.9.92, and 2.9.93).
//!
//! The `PlcfWKB` lists where each subdocument begins in the main document and
//! references its file through an `FNPI`; the `SttbFnm` stores the full paths
//! of all external files the document references (subdocuments and mail merge
//! data sources) together with per-file `FNIF` metadata.
//!
//! Everything here is inert: file paths are stored verbatim and are never
//! opened, resolved, or followed, and no subdocument content is ever loaded.

use super::fib::{FileInformationBlock, WORD_97_NFIB};
use super::mail_merge::Fnpi;
use crate::doc::package::{DocError, Result};
use std::collections::HashSet;

/// Table-pointer index of `fcPlcfWkb`/`lcbPlcfWkb`.
const PLCF_WKB: usize = 54;
/// Table-pointer index of `fcSttbFnm`/`lcbSttbFnm`.
const STTB_FNM: usize = 72;

/// Size in bytes of one `WKB` element (MS-DOC 2.9.346).
const WKB_LEN: usize = 12;
/// `WKB.fn`: the mandated value.
const WKB_FN: u16 = 0x0000;
/// `WKB` flag bits that are undefined and MUST be ignored: `fReserved3`
/// (bit 2) and `fReserved8` (bit 7).
const WKB_FLAGS_IGNORED: u16 = 0x0084;
/// The mandated value of the defined `WKB` flag bits: `fReserved6` (bit 5)
/// MUST be 1 and every other defined bit MUST be 0.
const WKB_FLAGS_REQUIRED: u16 = 0x0020;
/// `WKB.fReserved9` occupies the high byte of the flags field and MUST be 0.
const WKB_RESERVED9_MASK: u16 = 0xFF00;
/// `WKB.lvl`: the mandated outline level.
const WKB_OUTLINE_LEVEL: u16 = 0x0002;

/// `fExtend` marker of an extended STTB.
const STTB_F_EXTEND: u16 = 0xFFFF;
/// `SttbFnm.cbExtra`: one `FNIF` per entry (MS-DOC 2.9.288).
const STTB_FNM_CB_EXTRA: u16 = 8;
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
const FNFB_FAT: u8 = 0x01;
const FNFB_NTFS: u8 = 0x08;
const FNFB_NON_FILE_SYS: u8 = 0x10;

/// The kind of an externally referenced file (`FNPI.fnpt`, MS-DOC 2.9.93).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferencedFileKind {
    /// A mail merge data source file.
    MailMergeDataSource,
    /// A subdocument of a master document.
    Subdocument,
}

/// One external file referenced by the document: an `SttbFnm` string plus its
/// appended `FNIF` metadata (MS-DOC 2.9.288 and 2.9.92).
///
/// The path is stored verbatim and is never opened, resolved, or followed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferencedFileName {
    fnpi: Fnpi,
    /// The full path of the referenced file, including name and extension.
    pub path: String,
    /// `FNIF.ichRelative`: the character offset into `path` at which the
    /// document-relative path segment starts, or `None` when the file name
    /// carries no such segment.
    pub relative_path_offset: Option<usize>,
    /// Whether the path is valid on FAT file systems (`FNFB.fFAT`).
    pub valid_on_fat: bool,
    /// Whether the path is valid on NTFS file systems (`FNFB.fNTFS`).
    pub valid_on_ntfs: bool,
    /// Whether the path is not a native file system path and requires an
    /// external file I/O protocol (`FNFB.fNonFileSys`).
    pub is_non_file_system_path: bool,
}

impl ReferencedFileName {
    /// The type and identifier of this file name (`FNPI`, MS-DOC 2.9.93).
    pub fn fnpi(&self) -> Fnpi {
        self.fnpi
    }

    /// The kind of the referenced file.
    pub fn kind(&self) -> ReferencedFileKind {
        match self.fnpi.file_type() {
            FNPI_TYPE_SUBDOCUMENT => ReferencedFileKind::Subdocument,
            _ => ReferencedFileKind::MailMergeDataSource,
        }
    }

    /// The path segment relative to the folder containing the document, when
    /// the file name carries one. Never resolved against the file system.
    pub fn relative_path(&self) -> Option<&str> {
        self.relative_path_offset
            .and_then(|offset| self.path.get(offset..))
    }
}

/// One subdocument of a master document (`WKB`, MS-DOC 2.9.346).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Subdocument {
    /// Character position in the main document where the subdocument begins.
    pub start: u32,
    /// The outline level of the subdocument (`WKB.lvl`).
    pub outline_level: u16,
    /// The type and identifier of the subdocument file name (`WKB.fnpi`).
    pub file_name: Fnpi,
    /// Index of the resolved `SttbFnm` entry in
    /// `DocumentSubdocuments::referenced_files`.
    file_name_index: usize,
}

/// The master-document subdocument directory and the referenced-file name
/// table, addressed by `fcPlcfWkb` and `fcSttbFnm`.
///
/// All data is inert: paths are exposed verbatim and never opened, resolved,
/// or followed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentSubdocuments {
    referenced_files: Vec<ReferencedFileName>,
    subdocuments: Vec<Subdocument>,
}

impl DocumentSubdocuments {
    /// Parse the `PlcfWKB` and `SttbFnm` tables addressed by the FIB, or
    /// `None` when the document carries neither.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentSubdocuments>> {
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
        Ok(Some(DocumentSubdocuments {
            referenced_files,
            subdocuments,
        }))
    }

    /// All externally referenced files in `SttbFnm` table order.
    pub fn referenced_files(&self) -> &[ReferencedFileName] {
        &self.referenced_files
    }

    /// The subdocuments in start-CP order (empty unless this is a master
    /// document).
    pub fn subdocuments(&self) -> &[Subdocument] {
        &self.subdocuments
    }

    /// Resolve an `FNPI` reference to its `SttbFnm` entry.
    pub fn file_name(&self, fnpi: Fnpi) -> Option<&ReferencedFileName> {
        self.referenced_files.iter().find(|file| file.fnpi == fnpi)
    }

    /// The referenced file of a subdocument. Always resolves: entries are
    /// validated against the `SttbFnm` during parsing.
    pub fn file_name_of(&self, subdocument: &Subdocument) -> &ReferencedFileName {
        &self.referenced_files[subdocument.file_name_index]
    }
}

/// Parse `SttbFnm` (MS-DOC 2.9.288): an extended STTB of full file paths,
/// each followed by an 8-byte `FNIF` (MS-DOC 2.9.92).
fn parse_sttb_fnm(data: &[u8]) -> Result<Vec<ReferencedFileName>> {
    if data.len() < 6
        || read_u16(data, 0, "SttbFnm fExtend")? != STTB_F_EXTEND
        || read_u16(data, 4, "SttbFnm cbExtra")? != STTB_FNM_CB_EXTRA
    {
        return Err(corrupted("SttbFnm has an invalid header"));
    }
    let count = usize::from(read_u16(data, 2, "SttbFnm cData")?);
    let mut files = Vec::with_capacity(count);
    let mut identifiers = HashSet::with_capacity(count);
    let mut offset = 6usize;
    for _ in 0..count {
        let chars = usize::from(read_u16(data, offset, "SttbFnm cchData")?);
        offset += 2;
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
        let fnif = data
            .get(
                end..end
                    .checked_add(FNIF_LEN)
                    .ok_or_else(|| corrupted("SttbFnm FNIF range overflows"))?,
            )
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
        files.push(ReferencedFileName {
            fnpi,
            path,
            relative_path_offset,
            valid_on_fat: fnfb & FNFB_FAT != 0,
            valid_on_ntfs: fnfb & FNFB_NTFS != 0,
            is_non_file_system_path,
        });
        offset = end + FNIF_LEN;
    }
    if offset != data.len() {
        return Err(corrupted("SttbFnm contains trailing bytes"));
    }
    Ok(files)
}

/// Parse `PlcfWKB` (MS-DOC 2.8.34): a PLC of 12-byte `WKB` elements
/// (MS-DOC 2.9.346), resolving each `fnpi` against the `SttbFnm` entries.
fn parse_plcf_wkb(
    data: &[u8],
    main_document_chars: u32,
    referenced_files: &[ReferencedFileName],
) -> Result<Vec<Subdocument>> {
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
        subdocuments.push(Subdocument {
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

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a minimal FIB whose table-pointer array covers indexes 0..73,
    /// with a main-document length of `document_end` characters.
    fn fib_bytes(document_end: u32) -> Vec<u8> {
        let pointer_count = 73usize;
        let mut bytes = vec![0u8; 154 + pointer_count * 8];
        bytes[..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
        bytes[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        bytes[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
        bytes[76..80].copy_from_slice(&document_end.to_le_bytes());
        bytes[152..154].copy_from_slice(&(pointer_count as u16).to_le_bytes());
        bytes
    }

    fn set_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let base = 154 + index * 8;
        fib[base..base + 4].copy_from_slice(&offset.to_le_bytes());
        fib[base + 4..base + 8].copy_from_slice(&length.to_le_bytes());
    }

    fn utf16(text: &str) -> Vec<u8> {
        text.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    /// One `SttbFnm` entry: (path, fnpt, fnpd, ichRelative, fnfb).
    type FileEntry = (&'static str, u8, u16, u8, u8);

    fn sttb_fnm(entries: &[FileEntry]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
        data.extend_from_slice(&(entries.len() as u16).to_le_bytes());
        data.extend_from_slice(&STTB_FNM_CB_EXTRA.to_le_bytes());
        for (path, fnpt, fnpd, ich_relative, fnfb) in entries {
            let encoded = utf16(path);
            data.extend_from_slice(&((encoded.len() / 2) as u16).to_le_bytes());
            data.extend_from_slice(&encoded);
            let fnpi = u16::from(*fnpt) | (fnpd << 4);
            data.extend_from_slice(&fnpi.to_le_bytes());
            data.push(*ich_relative);
            data.push(*fnfb);
            data.extend_from_slice(&[0; 4]); // unused
        }
        data
    }

    /// Build a `PlcfWKB` from (start CP, fnpi) entries.
    fn plcf_wkb(entries: &[(u32, u16)], terminal_cp: u32) -> Vec<u8> {
        let mut data = Vec::new();
        for (cp, _) in entries {
            data.extend_from_slice(&cp.to_le_bytes());
        }
        data.extend_from_slice(&terminal_cp.to_le_bytes());
        for (_, fnpi) in entries {
            data.extend_from_slice(&WKB_FN.to_le_bytes());
            data.extend_from_slice(&WKB_FLAGS_REQUIRED.to_le_bytes());
            data.extend_from_slice(&WKB_OUTLINE_LEVEL.to_le_bytes());
            data.extend_from_slice(&fnpi.to_le_bytes());
            data.extend_from_slice(&0u32.to_le_bytes()); // pdod
        }
        data
    }

    fn fnpi(fnpt: u8, fnpd: u16) -> u16 {
        u16::from(fnpt) | (fnpd << 4)
    }

    /// A typical master document: two subdocuments at CPs 2 and 5 in a
    /// 10-character main document, plus one mail merge data source.
    struct Tables {
        fnm: Vec<u8>,
        wkb: Vec<u8>,
    }

    impl Tables {
        fn typical() -> Self {
            Self {
                fnm: sttb_fnm(&[
                    ("C:\\docs\\intro.doc", 5, 0, 8, FNFB_FAT | FNFB_NTFS),
                    ("C:\\docs\\body.doc", 5, 1, 0xFF, FNFB_NTFS),
                    ("D:\\data\\list.csv", 3, 2, 0xFF, FNFB_FAT),
                ]),
                wkb: plcf_wkb(&[(2, fnpi(5, 0)), (5, fnpi(5, 1))], 12),
            }
        }

        fn assemble(&self) -> (Vec<u8>, Vec<u8>) {
            let mut fib = fib_bytes(10);
            let mut table = Vec::new();
            for (index, data) in [(STTB_FNM, &self.fnm), (PLCF_WKB, &self.wkb)] {
                if !data.is_empty() {
                    set_pointer(&mut fib, index, table.len() as u32, data.len() as u32);
                    table.extend_from_slice(data);
                }
            }
            (fib, table)
        }

        fn parse(&self) -> Result<Option<DocumentSubdocuments>> {
            let (fib, table) = self.assemble();
            let fib = FileInformationBlock::parse(&fib).unwrap();
            DocumentSubdocuments::parse(&fib, &table)
        }
    }

    #[test]
    fn parses_master_document_tables() {
        let parsed = Tables::typical().parse().unwrap().unwrap();
        let files = parsed.referenced_files();
        assert_eq!(files.len(), 3);
        assert_eq!(files[0].path, "C:\\docs\\intro.doc");
        assert_eq!(files[0].kind(), ReferencedFileKind::Subdocument);
        assert_eq!(files[0].relative_path_offset, Some(8));
        assert_eq!(files[0].relative_path(), Some("intro.doc"));
        assert!(files[0].valid_on_fat && files[0].valid_on_ntfs);
        assert!(!files[0].is_non_file_system_path);
        assert_eq!(files[1].relative_path_offset, None);
        assert_eq!(files[1].relative_path(), None);
        assert_eq!(files[2].kind(), ReferencedFileKind::MailMergeDataSource);

        let subdocuments = parsed.subdocuments();
        assert_eq!(subdocuments.len(), 2);
        assert_eq!(subdocuments[0].start, 2);
        assert_eq!(subdocuments[0].outline_level, 2);
        assert_eq!(
            parsed.file_name_of(&subdocuments[0]).path,
            "C:\\docs\\intro.doc"
        );
        assert_eq!(
            parsed.file_name_of(&subdocuments[1]).path,
            "C:\\docs\\body.doc"
        );
        assert!(parsed.file_name(Fnpi::from_raw(fnpi(5, 7))).is_none());
    }

    #[test]
    fn reports_absent_tables_as_none() {
        let fib = fib_bytes(10);
        let fib = FileInformationBlock::parse(&fib).unwrap();
        assert!(DocumentSubdocuments::parse(&fib, &[]).unwrap().is_none());
    }

    #[test]
    fn parses_file_name_table_without_subdocuments() {
        let tables = Tables {
            wkb: Vec::new(),
            ..Tables::typical()
        };
        let parsed = tables.parse().unwrap().unwrap();
        assert_eq!(parsed.referenced_files().len(), 3);
        assert!(parsed.subdocuments().is_empty());
    }

    #[test]
    fn rejects_subdocuments_without_file_name_table() {
        let tables = Tables {
            fnm: Vec::new(),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn rejects_invalid_sttb_fnm_framing() {
        // Wrong fExtend.
        let mut fnm = sttb_fnm(&[("a.doc", 5, 0, 0xFF, FNFB_NTFS)]);
        fnm[0..2].copy_from_slice(&0u16.to_le_bytes());
        assert!(parse_sttb_fnm(&fnm).is_err());
        // Wrong cbExtra.
        let mut fnm = sttb_fnm(&[("a.doc", 5, 0, 0xFF, FNFB_NTFS)]);
        fnm[4..6].copy_from_slice(&4u16.to_le_bytes());
        assert!(parse_sttb_fnm(&fnm).is_err());
        // Trailing bytes.
        let mut fnm = sttb_fnm(&[("a.doc", 5, 0, 0xFF, FNFB_NTFS)]);
        fnm.extend_from_slice(&[0, 0]);
        assert!(parse_sttb_fnm(&fnm).is_err());
        // Truncated string.
        let fnm = sttb_fnm(&[("a.doc", 5, 0, 0xFF, FNFB_NTFS)]);
        assert!(parse_sttb_fnm(&fnm[..fnm.len() - 4]).is_err());
    }

    #[test]
    fn rejects_invalid_fnif_values() {
        // Undefined fnpt.
        assert!(parse_sttb_fnm(&sttb_fnm(&[("a.doc", 4, 0, 0xFF, 0)])).is_err());
        // Reserved nil fnpd.
        assert!(parse_sttb_fnm(&sttb_fnm(&[("a.doc", 5, 0xFFF, 0xFF, 0)])).is_err());
        // Duplicate fnpi.
        assert!(
            parse_sttb_fnm(&sttb_fnm(&[
                ("a.doc", 5, 0, 0xFF, 0),
                ("b.doc", 5, 0, 0xFF, 0),
            ]))
            .is_err()
        );
        // Same fnpd under a different fnpt is allowed.
        assert!(
            parse_sttb_fnm(&sttb_fnm(&[
                ("a.doc", 5, 0, 0xFF, 0),
                ("b.csv", 3, 0, 0xFF, 0),
            ]))
            .is_ok()
        );
        // ichRelative beyond the string.
        assert!(parse_sttb_fnm(&sttb_fnm(&[("a.doc", 5, 0, 5, 0)])).is_err());
        // A non-file-system path must not be marked FAT/NTFS valid.
        assert!(
            parse_sttb_fnm(&sttb_fnm(&[(
                "a.doc",
                5,
                0,
                0xFF,
                FNFB_NON_FILE_SYS | FNFB_FAT
            )]))
            .is_err()
        );
        assert!(
            parse_sttb_fnm(&sttb_fnm(&[(
                "http://x/a.doc",
                5,
                0,
                0xFF,
                FNFB_NON_FILE_SYS
            )]))
            .is_ok()
        );
    }

    #[test]
    fn rejects_invalid_plcf_wkb_framing() {
        // Byte length that does not fit a whole number of WKB elements.
        assert!(parse_plcf_wkb(&[0u8; 15], 10, &[]).is_err());
        // Missing terminal CP.
        assert!(parse_plcf_wkb(&[0u8; 4], 10, &[]).is_err());
    }

    #[test]
    fn rejects_invalid_cps() {
        // CP at or beyond the main document length.
        let tables = Tables {
            wkb: plcf_wkb(&[(10, fnpi(5, 0))], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
        // Non-increasing CPs.
        let tables = Tables {
            wkb: plcf_wkb(&[(5, fnpi(5, 0)), (5, fnpi(5, 1))], 12),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
        // Wrong terminal CP.
        let tables = Tables {
            wkb: plcf_wkb(&[(2, fnpi(5, 0)), (5, fnpi(5, 1))], 11),
            ..Tables::typical()
        };
        assert!(tables.parse().is_err());
    }

    #[test]
    fn rejects_invalid_wkb_fields() {
        let files = parse_sttb_fnm(&Tables::typical().fnm).unwrap();
        let valid = plcf_wkb(&[(2, fnpi(5, 0))], 12);

        // fn MUST be 0.
        let mut wkb = valid.clone();
        wkb[8..10].copy_from_slice(&1u16.to_le_bytes());
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
        // fReserved6 MUST be 1.
        let mut wkb = valid.clone();
        wkb[10..12].copy_from_slice(&0u16.to_le_bytes());
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
        // fReserved9 MUST be 0.
        let mut wkb = valid.clone();
        wkb[11] = 0x01;
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
        // fReserved3 and fReserved8 are ignored.
        let mut wkb = valid.clone();
        wkb[10] |= 0x84;
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_ok());
        // lvl MUST be 0x0002.
        let mut wkb = valid.clone();
        wkb[12..14].copy_from_slice(&1u16.to_le_bytes());
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
        // fnpi MUST reference a subdocument, not a mail merge source.
        let mut wkb = valid.clone();
        wkb[14..16].copy_from_slice(&fnpi(3, 2).to_le_bytes());
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
        // fnpi MUST resolve against the SttbFnm.
        let mut wkb = valid.clone();
        wkb[14..16].copy_from_slice(&fnpi(5, 9).to_le_bytes());
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
        // pdod MUST be 0.
        let mut wkb = valid;
        wkb[16..20].copy_from_slice(&1u32.to_le_bytes());
        assert!(parse_plcf_wkb(&wkb, 10, &files).is_err());
    }

    #[test]
    fn rejects_pre_word97_fibs() {
        let (mut fib, table) = Tables::typical().assemble();
        fib[2..4].copy_from_slice(&0x0065u16.to_le_bytes());
        let fib = FileInformationBlock::parse(&fib).unwrap();
        assert!(DocumentSubdocuments::parse(&fib, &table).unwrap().is_none());
    }
}
