use super::codec::{
    FNFB_FAT, FNFB_NON_FILE_SYS, FNFB_NTFS, PLCF_WKB, STTB_F_EXTEND, STTB_FNM, STTB_FNM_CB_EXTRA,
    WKB_FLAGS_REQUIRED, WKB_FN, WKB_OUTLINE_LEVEL, parse_plcf_wkb, parse_sttb_fnm,
};
use super::model::{Collection, Kind};
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;
use crate::parts::mail_merge::Fnpi;

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

    fn parse(&self) -> Result<Option<Collection>> {
        let (fib, table) = self.assemble();
        let fib = FileInformationBlock::parse(&fib).unwrap();
        Collection::parse(&fib, &table)
    }
}

#[test]
fn parses_master_document_tables() {
    let parsed = Tables::typical().parse().unwrap().unwrap();
    let files = parsed.referenced_files();
    assert_eq!(files.len(), 3);
    assert_eq!(files[0].path, "C:\\docs\\intro.doc");
    assert_eq!(files[0].kind(), Kind::Subdocument);
    assert_eq!(files[0].relative_path_offset, Some(8));
    assert_eq!(files[0].relative_path(), Some("intro.doc"));
    assert!(files[0].valid_on_fat && files[0].valid_on_ntfs);
    assert!(!files[0].is_non_file_system_path);
    assert_eq!(files[1].relative_path_offset, None);
    assert_eq!(files[1].relative_path(), None);
    assert_eq!(files[2].kind(), Kind::MailMergeDataSource);

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
    assert!(Collection::parse(&fib, &[]).unwrap().is_none());
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
    assert!(Collection::parse(&fib, &table).unwrap().is_none());
}
