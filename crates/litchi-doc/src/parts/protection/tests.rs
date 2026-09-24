use super::codec::{
    BKC_F_COL, BKC_F_NATIVE, BKC_ITC_LIM_SHIFT, PLCF_BKF_PROT, PLCF_BKL_PROT, PRTI_SIZE,
    STTB_F_EXTEND, STTB_PROT_USER, STTBF_BKMK_PROT, USER_ROLE_SIZE, parse_assignments, parse_ends,
    parse_starts, parse_users,
};
use super::model::{Mode, Ranges, Role, Selector, User};
use super::policy::{EditProtection, ProtectionAuthorization, ProtectionPolicy, classify};
use crate::package::Error as PackageError;
use crate::package::Result;
use crate::parts::document_properties::DocumentProperties;
use crate::parts::fib::FileInformationBlock;
use litchi_cfb::OleFile;
use std::io::Cursor;

const BKC_F_PUB: u16 = 0x0080;

/// Build a minimal FIB whose table-pointer array covers indexes 0..144,
/// with a main-document length of `document_end` characters.
fn fib_bytes(document_end: u32) -> Vec<u8> {
    // Use the Word 2010 counted layout so the range-protection pointers at
    // indexes 141..144 are inside the specification-defined array.  The base
    // nFib remains Word 97; cswNew/nFibNew selects the effective Word 2010
    // layout as real modern DOC files do.
    let pointer_count = 183usize;
    let pointer_end = 154 + pointer_count * 8;
    let mut bytes = vec![0u8; pointer_end + 12];
    bytes[..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    bytes[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    bytes[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    bytes[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    bytes[6..8].copy_from_slice(&0x0409u16.to_le_bytes());
    bytes[76..80].copy_from_slice(&document_end.to_le_bytes());
    bytes[152..154].copy_from_slice(&(pointer_count as u16).to_le_bytes());
    bytes[pointer_end..pointer_end + 2].copy_from_slice(&5u16.to_le_bytes());
    bytes[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0112u16.to_le_bytes());
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

/// The `WordDocument` and selected table streams of a checked-in DOC fixture.
fn fixture_streams(name: &str) -> (Vec<u8>, Vec<u8>) {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name);
    let bytes = std::fs::read(path).expect("DOC fixture");
    let mut ole = OleFile::open(Cursor::new(bytes)).expect("DOC fixture CFB");
    let word = ole.open_stream(&["WordDocument"]).expect("WordDocument");
    let table_name = if u16::from_le_bytes([word[10], word[11]]) & 0x0200 != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table = ole.open_stream(&[table_name]).expect("table stream");
    (word, table)
}

/// The exact DOP bytes a checked-in DOC fixture carries.
fn fixture_dop(name: &str) -> Vec<u8> {
    let (word, table) = fixture_streams(name);
    let count = usize::from(u16::from_le_bytes([word[152], word[153]]));
    let pointer = 154 + 31 * 8;
    assert!(count > 31);
    let offset = usize::try_from(u32::from_le_bytes(
        word[pointer..pointer + 4].try_into().expect("DOP offset"),
    ))
    .expect("DOP offset");
    let length = usize::try_from(u32::from_le_bytes(
        word[pointer + 4..pointer + 8]
            .try_into()
            .expect("DOP length"),
    ))
    .expect("DOP length");
    table[offset..offset + length].to_vec()
}

fn valid_modern_dop() -> Vec<u8> {
    fixture_dop("NoHeadFoot.doc")
}

/// A counted FIB with the given `FibBase.nFib`, `cbRgFcLcb` and `cswNew`
/// (followed by `nFibNew` and a complete `FibRgCswNew` when `cswNew` is
/// nonzero), whose DOP pointer covers `dop_length` bytes at table offset 0.
fn counted_fib(
    base_nfib: u16,
    pointer_count: usize,
    csw_new: u16,
    nfib_new: u16,
    dop_length: usize,
) -> Vec<u8> {
    let pointer_end = 154 + pointer_count * 8;
    let mut fib = vec![0u8; pointer_end + 2 + usize::from(csw_new) * 2];
    fib[..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    fib[2..4].copy_from_slice(&base_nfib.to_le_bytes());
    fib[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    fib[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    fib[152..154].copy_from_slice(&u16::try_from(pointer_count).unwrap().to_le_bytes());
    fib[pointer_end..pointer_end + 2].copy_from_slice(&csw_new.to_le_bytes());
    if csw_new != 0 {
        fib[pointer_end + 2..pointer_end + 4].copy_from_slice(&nfib_new.to_le_bytes());
    }
    set_pointer(&mut fib, 31, 0, u32::try_from(dop_length).unwrap());
    fib
}

fn classify_bytes(fib: &[u8], table: &[u8]) -> EditProtection {
    classify(&FileInformationBlock::parse(fib).unwrap(), table).unwrap()
}

/// Sets `Dop2003.fEnforceDocProt` and `iDocProtCur` in byte 598.
fn enforced(mut dop: Vec<u8>, mode: u8) -> Vec<u8> {
    dop[598] = (dop[598] & !0x78) | 0x08 | (mode << 4);
    dop
}

fn sttb_users(users: &[(&str, u16)]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
    data.extend_from_slice(&(users.len() as u16).to_le_bytes());
    data.extend_from_slice(&USER_ROLE_SIZE.to_le_bytes());
    for (name, role) in users {
        let encoded = utf16(name);
        data.extend_from_slice(&((encoded.len() / 2) as u16).to_le_bytes());
        data.extend_from_slice(&encoded);
        data.extend_from_slice(&role.to_le_bytes());
    }
    data
}

fn sttbf_ranges(editors: &[Selector]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
    data.extend_from_slice(&(editors.len() as u32).to_le_bytes());
    data.extend_from_slice(&(PRTI_SIZE as u16).to_le_bytes());
    for editor in editors {
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&editor.raw().to_le_bytes());
        data.extend_from_slice(&Mode::ReadWrite.raw().to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
    }
    data
}

/// (start CP, ibkl, bkc) entries.
fn plcf_bkf(entries: &[(u32, u32, u16)], terminal_cp: u32) -> Vec<u8> {
    let mut data = Vec::new();
    for (cp, _, _) in entries {
        data.extend_from_slice(&cp.to_le_bytes());
    }
    data.extend_from_slice(&terminal_cp.to_le_bytes());
    for (_, ibkl, bkc) in entries {
        data.extend_from_slice(&ibkl.to_le_bytes());
        data.extend_from_slice(&bkc.to_le_bytes());
    }
    data
}

fn plcf_bkl(end_cps: &[u32], terminal_cp: u32) -> Vec<u8> {
    let mut data = Vec::new();
    for cp in end_cps {
        data.extend_from_slice(&cp.to_le_bytes());
    }
    data.extend_from_slice(&terminal_cp.to_le_bytes());
    data
}

struct Tables {
    users: Vec<u8>,
    assignments: Vec<u8>,
    starts: Vec<u8>,
    ends: Vec<u8>,
}

impl Tables {
    fn typical() -> Self {
        Self {
            users: sttb_users(&[
                ("CONTOSO\\alice", Role::Editor.raw()),
                ("bob@example.com", Role::Owner.raw()),
            ]),
            assignments: sttbf_ranges(&[Selector::User(1), Selector::Everyone]),
            starts: plcf_bkf(&[(2, 0, BKC_F_NATIVE), (4, 1, 0)], 12),
            ends: plcf_bkl(&[7, 9], 12),
        }
    }

    fn assemble(&self) -> (Vec<u8>, Vec<u8>) {
        let mut fib = fib_bytes(10);
        // A nonzero lcbDop is mandatory in MS-DOC.  Keep a valid base DOP in
        // every protection fixture so range-only classification exercises the
        // range state rather than the malformed-host state.
        let mut table = valid_modern_dop();
        set_pointer(&mut fib, 31, 0, table.len() as u32);
        for (index, data) in [
            (STTBF_BKMK_PROT, &self.assignments),
            (PLCF_BKF_PROT, &self.starts),
            (PLCF_BKL_PROT, &self.ends),
            (STTB_PROT_USER, &self.users),
        ] {
            if !data.is_empty() {
                set_pointer(&mut fib, index, table.len() as u32, data.len() as u32);
                table.extend_from_slice(data);
            }
        }
        (fib, table)
    }

    fn parse(&self) -> Result<Option<Ranges>> {
        let (fib, table) = self.assemble();
        let fib = FileInformationBlock::parse(&fib).unwrap();
        Ranges::parse(&fib, &table)
    }
}

#[test]
fn parses_typed_range_level_protection() {
    let parsed = Tables::typical().parse().unwrap().unwrap();
    assert_eq!(
        parsed.users(),
        &[
            User {
                name: "CONTOSO\\alice".to_string(),
                role: Role::Editor,
            },
            User {
                name: "bob@example.com".to_string(),
                role: Role::Owner,
            },
        ]
    );
    assert_eq!(parsed.ranges().len(), 2);
    assert_eq!(
        (
            parsed.ranges()[0].start,
            parsed.ranges()[0].end,
            parsed.ranges()[0].is_native,
            parsed.ranges()[0].column,
            parsed.ranges()[0].editor,
            parsed.ranges()[0].mode,
        ),
        (2, 7, true, None, Selector::User(1), Mode::ReadWrite)
    );
    assert_eq!(
        (
            parsed.ranges()[1].start,
            parsed.ranges()[1].end,
            parsed.ranges()[1].is_native,
            parsed.ranges()[1].column,
            parsed.ranges()[1].editor,
            parsed.ranges()[1].mode,
        ),
        (4, 9, false, None, Selector::Everyone, Mode::ReadWrite)
    );
    let range = &parsed.ranges()[0];
    assert_eq!(
        parsed.editor_for(range),
        Some(&User {
            name: "CONTOSO\\alice".to_string(),
            role: Role::Editor,
        })
    );
    assert!(parsed.editor_for(&parsed.ranges()[1]).is_none());
    assert!(parsed.user(3).is_none());
}

#[test]
fn preserves_unknown_selectors_modes_roles_and_reserved_words() {
    let mut assignments = sttbf_ranges(&[Selector::Unknown(0xFFFA)]);
    assignments[12..14].copy_from_slice(&0x1234u16.to_le_bytes());
    assignments[14..16].copy_from_slice(&0x5678u16.to_le_bytes());
    assignments[16..18].copy_from_slice(&0x9ABCu16.to_le_bytes());
    assignments[8..10].copy_from_slice(&1u16.to_le_bytes());
    assignments.splice(10..10, [0x41, 0x00]);

    let tables = Tables {
        users: sttb_users(&[("alice", 0x4242)]),
        assignments,
        starts: plcf_bkf(&[(2, 0, BKC_F_PUB | BKC_F_NATIVE)], 12),
        ends: plcf_bkl(&[7], 12),
    };
    let parsed = tables.parse().unwrap().unwrap();
    let range = &parsed.ranges()[0];
    assert_eq!(range.editor, Selector::Unknown(0xFFFA));
    assert_eq!(range.mode, Mode::Unknown(0x1234));
    assert_eq!(range.reserved().bkc(), BKC_F_PUB | BKC_F_NATIVE);
    assert_eq!(range.reserved().prti_i(), 0x5678);
    assert_eq!(range.reserved().prti_use_me(), 0x9ABC);
    assert_eq!(range.reserved().bookmark_data(), &[0x41, 0x00]);
    assert_eq!(parsed.users()[0].role, Role::Unknown(0x4242));
}

#[test]
fn reports_absent_tables_as_none() {
    let fib = FileInformationBlock::parse(&fib_bytes(10)).unwrap();
    assert!(Ranges::parse(&fib, &[]).unwrap().is_none());
}

#[test]
fn parses_username_table_without_bookmark_tables() {
    let tables = Tables {
        assignments: Vec::new(),
        starts: Vec::new(),
        ends: Vec::new(),
        ..Tables::typical()
    };
    let parsed = tables.parse().unwrap().unwrap();
    assert_eq!(parsed.users().len(), 2);
    assert!(parsed.ranges().is_empty());
}

#[test]
fn rejects_partially_present_bookmark_tables() {
    let tables = Tables {
        ends: Vec::new(),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
}

#[test]
fn rejects_mismatched_parallel_counts() {
    let tables = Tables {
        assignments: sttbf_ranges(&[Selector::Everyone]),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
}

#[test]
fn rejects_dangling_user_indexes_and_invalid_columns() {
    let tables = Tables {
        assignments: sttbf_ranges(&[Selector::User(3), Selector::Everyone]),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());

    let invalid = BKC_F_COL | 2 | (4 << BKC_ITC_LIM_SHIFT);
    let tables = Tables {
        starts: plcf_bkf(&[(2, 1, invalid), (4, 0, 0)], 12),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
}

#[test]
fn rejects_duplicate_or_dangling_end_indexes() {
    let tables = Tables {
        starts: plcf_bkf(&[(2, 0, 0), (4, 0, 0)], 12),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
    let tables = Tables {
        starts: plcf_bkf(&[(2, 5, 0), (4, 0, 0)], 12),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
}

#[test]
fn rejects_reversed_and_out_of_range_cps() {
    let tables = Tables {
        starts: plcf_bkf(&[(9, 1, 0), (4, 0, 0)], 12),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
    let tables = Tables {
        ends: plcf_bkl(&[11, 7], 12),
        ..Tables::typical()
    };
    assert!(tables.parse().is_err());
}

#[test]
fn accepts_reserved_public_bit_while_retaining_it() {
    let tables = Tables {
        starts: plcf_bkf(&[(2, 1, BKC_F_PUB), (4, 0, 0)], 12),
        ..Tables::typical()
    };
    let parsed = tables.parse().unwrap().unwrap();
    assert_eq!(parsed.ranges()[0].reserved().bkc(), BKC_F_PUB);
}

#[test]
fn rejects_invalid_table_framing_and_truncation() {
    let mut assignments = sttbf_ranges(&[Selector::Everyone]);
    assignments[12..14].copy_from_slice(&0x0004u16.to_le_bytes());
    // The non-read/write value is a typed unknown/known mode, not framing.
    assert!(parse_assignments(&assignments).is_ok());

    let mut wrong_extra = sttbf_ranges(&[Selector::Everyone]);
    wrong_extra[6..8].copy_from_slice(&4u16.to_le_bytes());
    assert!(parse_assignments(&wrong_extra).is_err());

    let mut count_mismatch = sttbf_ranges(&[Selector::Everyone]);
    count_mismatch[2..6].copy_from_slice(&2u32.to_le_bytes());
    assert!(parse_assignments(&count_mismatch).is_err());

    let mut users = sttb_users(&[("a", 0x0000)]);
    users.extend_from_slice(&[0, 0]);
    assert!(parse_users(&users).is_err());

    let duplicate = sttb_users(&[("a", 0x0000), ("a", 0x0000)]);
    assert!(parse_users(&duplicate).is_err());

    assert!(parse_assignments(&sttbf_ranges(&[Selector::Everyone])[..12]).is_err());
    assert!(parse_users(&sttb_users(&[("alice", 0x0000)])[..8]).is_err());
    assert!(parse_starts(&[0u8; 9]).is_err());
    assert!(parse_ends(&[0u8; 6]).is_err());
}

#[test]
fn bounds_large_encoded_counts_before_allocating() {
    let mut assignments = Vec::new();
    assignments.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
    assignments.extend_from_slice(&0x00007FF0u32.to_le_bytes());
    assignments.extend_from_slice(&(PRTI_SIZE as u16).to_le_bytes());
    assert!(parse_assignments(&assignments).is_err());

    let mut users = Vec::new();
    users.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
    users.extend_from_slice(&u16::MAX.to_le_bytes());
    users.extend_from_slice(&USER_ROLE_SIZE.to_le_bytes());
    assert!(parse_users(&users).is_err());
}

#[test]
fn classifies_document_and_range_protection_independently() {
    let mut fib = fib_bytes(10);
    let mut table = valid_modern_dop();
    table[6] = 0x10;
    set_pointer(&mut fib, 31, 0, table.len() as u32);
    let fib = FileInformationBlock::parse(&fib).unwrap();
    assert_eq!(classify(&fib, &table).unwrap(), EditProtection::Document);

    let range_tables = Tables::typical();
    let (range_fib, range_table) = range_tables.assemble();
    let range_fib = FileInformationBlock::parse(&range_fib).unwrap();
    assert_eq!(
        classify(&range_fib, &range_table).unwrap(),
        EditProtection::Ranges
    );

    table.extend_from_slice(&range_table);
    let mut combined_fib = range_tables.assemble().0;
    let combined_offset = table.len() - range_table.len();
    set_pointer(&mut combined_fib, 31, 0, 674);
    for index in [
        STTBF_BKMK_PROT,
        PLCF_BKF_PROT,
        PLCF_BKL_PROT,
        STTB_PROT_USER,
    ] {
        let base = 154 + index * 8;
        let offset = u32::from_le_bytes(combined_fib[base..base + 4].try_into().unwrap());
        let relocated = u32::try_from(combined_offset)
            .unwrap()
            .checked_add(offset)
            .unwrap();
        combined_fib[base..base + 4].copy_from_slice(&relocated.to_le_bytes());
    }
    let combined_fib = FileInformationBlock::parse(&combined_fib).unwrap();
    assert_eq!(
        classify(&combined_fib, &table).unwrap(),
        EditProtection::DocumentAndRanges
    );

    let authorization = ProtectionAuthorization::audited("alice", "approved").unwrap();
    assert!(
        ProtectionPolicy::default()
            .authorize(EditProtection::Document)
            .is_err_and(|error| matches!(error, PackageError::ProtectionDenied(_)))
    );
    assert!(
        ProtectionPolicy::allow_protected(authorization)
            .authorize(EditProtection::DocumentAndRanges)
            .is_ok()
    );
    assert!(
        ProtectionPolicy::default()
            .authorize(EditProtection::Unrecognized)
            .is_err()
    );
    let authorization = ProtectionAuthorization::audited("alice", "repair legacy DOP").unwrap();
    assert!(
        ProtectionPolicy::allow_protected(authorization)
            .authorize(EditProtection::Unrecognized)
            .is_err()
    );
    let authorization = ProtectionAuthorization::audited("alice", "repair malformed host").unwrap();
    assert!(
        ProtectionPolicy::allow_protected(authorization)
            .authorize(EditProtection::Unknown)
            .is_err()
    );
}

/// Word 2002 FIB and DOP shapes are classified by the protection fields they
/// carry (MS-DOC 2.5.14, 2.7.2, 2.7.7), not by an exact producer length.
#[test]
fn word_2002_shapes_are_classified_by_their_protection_fields() {
    let authorization =
        ProtectionAuthorization::audited("test-suite", "classify Word 2002 shapes").unwrap();
    let allow = ProtectionPolicy::allow_protected(authorization);

    // The conforming Word 2002 FIB (cbRgFcLcb 0x88, cswNew 2, nFibNew 0x0101)
    // with the complete 594-byte Dop2002 the writer emits.
    let mut dop2002 = DocumentProperties::word97_writer_bytes(false, false, false);
    dop2002.resize(594, 0);
    let word2002 = counted_fib(0x00C1, 0x88, 2, 0x0101, 594);
    assert_eq!(classify_bytes(&word2002, &dop2002), EditProtection::None);

    // The same FIB with the 500-byte DOP the pre-0759 writer emitted. The
    // DopBase is complete and Dop2000/Dop2002 add no protection field.
    let short_writer = counted_fib(0x00C1, 0x88, 2, 0x0101, 500);
    assert_eq!(
        classify_bytes(&short_writer, &dop2002[..500]),
        EditProtection::None
    );

    // LibreOffice's shape: FibBase.nFib 0x0101 with a zero cswNew, so 2.5.14
    // selects FibBase.nFib, which cbRgFcLcb 0x88 confirms; its 610-byte DOP
    // carries wvkoSaved 7 and 0x0080 in the Dop2003 word at 598.
    let libreoffice = fixture_dop("documentProperties.doc");
    assert_eq!(libreoffice.len(), 610);
    assert_eq!(libreoffice[82] & 0x07, 7);
    assert_eq!(&libreoffice[598..600], &[0x80, 0x00]);
    let libreoffice_fib = counted_fib(0x0101, 0x88, 0, 0, 610);
    assert_eq!(
        classify_bytes(&libreoffice_fib, &libreoffice),
        EditProtection::None
    );

    // An enforced restriction in that 610-byte DOP is found at byte 598.
    for mode in 0..=3 {
        assert_eq!(
            classify_bytes(&libreoffice_fib, &enforced(libreoffice.clone(), mode)),
            EditProtection::Document,
            "enforced iDocProtCur {mode}"
        );
    }
    assert!(
        ProtectionPolicy::default()
            .authorize(EditProtection::Document)
            .is_err()
    );
    assert!(allow.authorize(EditProtection::Document).is_ok());
    assert_eq!(
        classify_bytes(&libreoffice_fib, &enforced(libreoffice.clone(), 7)),
        EditProtection::None
    );
    for mode in 4..=6 {
        assert_eq!(
            classify_bytes(&libreoffice_fib, &enforced(libreoffice.clone(), mode)),
            EditProtection::Unrecognized
        );
        let mut reserved = libreoffice.clone();
        reserved[598] = (reserved[598] & !0x78) | (mode << 4);
        assert_eq!(
            classify_bytes(&libreoffice_fib, &reserved),
            EditProtection::Unrecognized,
            "reserved iDocProtCur {mode} without enforcement"
        );
    }
    assert!(allow.authorize(EditProtection::Unrecognized).is_err());

    // A DopBase lock in the LibreOffice DOP is a document restriction.
    let mut forms = libreoffice.clone();
    forms[7] |= 0x02;
    forms[78..82].copy_from_slice(&0x0BAD_F00Du32.to_le_bytes());
    assert_eq!(
        classify_bytes(&libreoffice_fib, &forms),
        EditProtection::Document
    );

    // A DOP that ends inside DopBase cannot be classified.
    let truncated_base = counted_fib(0x0101, 0x88, 0, 0, 83);
    assert_eq!(
        classify_bytes(&truncated_base, &libreoffice[..83]),
        EditProtection::Unrecognized
    );

    // The counted FIB must still describe the generation 2.5.14 selects.
    assert_eq!(
        classify_bytes(&counted_fib(0x0101, 0x6C, 0, 0, 610), &libreoffice),
        EditProtection::Unknown,
        "cbRgFcLcb of another generation"
    );
    assert_eq!(
        classify_bytes(&counted_fib(0x00C1, 0x88, 5, 0x0101, 610), &libreoffice),
        EditProtection::Unknown,
        "nonzero cswNew that is wrong for nFibNew"
    );
    let mut truncated_nfib_new = counted_fib(0x00C1, 0x88, 2, 0x0101, 594);
    truncated_nfib_new.truncate(truncated_nfib_new.len() - 3);
    assert_eq!(
        classify_bytes(&truncated_nfib_new, &dop2002),
        EditProtection::Unknown,
        "nFibNew truncated"
    );
    let mut missing_csw_new = counted_fib(0x0101, 0x88, 0, 0, 610);
    missing_csw_new.truncate(missing_csw_new.len() - 1);
    assert_eq!(
        classify_bytes(&missing_csw_new, &libreoffice),
        EditProtection::Unknown,
        "cswNew truncated"
    );
    assert!(allow.authorize(EditProtection::Unknown).is_err());
}

/// Appendix A note <11>: `FibBase.nFib` 0x00C0 (the shell's empty document)
/// and 0x00C2 (the BiDi build of Word 97) are read as 0x00C1.
#[test]
fn note_11_reads_shell_and_bidi_nfib_values_as_word_97() {
    let word97 = DocumentProperties::word97_writer_bytes(false, false, false);
    assert_eq!(word97.len(), 500);
    for nfib in [0x00C0, 0x00C1, 0x00C2] {
        assert_eq!(
            classify_bytes(&counted_fib(nfib, 0x5D, 0, 0, 500), &word97),
            EditProtection::None,
            "nFib 0x{nfib:04X}"
        );
        assert_eq!(
            classify_bytes(&counted_fib(nfib, 0x88, 0, 0, 500), &word97),
            EditProtection::Unknown,
            "nFib 0x{nfib:04X} with a Word 2002 pointer count"
        );
    }
    for nfib in [0x00BF, 0x00C3, 0x00D8, 0x0100, 0x0113] {
        assert_eq!(
            classify_bytes(&counted_fib(nfib, 0x5D, 0, 0, 500), &word97),
            EditProtection::Unknown,
            "unknown nFib 0x{nfib:04X}"
        );
    }
    // A nonzero cswNew supersedes FibBase.nFib, and note <11> does not make
    // 0x00C0 a valid nFibNew.
    assert_eq!(
        classify_bytes(&counted_fib(0x00C1, 0x5D, 2, 0x00C0, 500), &word97),
        EditProtection::Unknown
    );
}

/// From Word 2003 on the DOP must reach the enforcement unit at 598..600;
/// shorter DOPs of that generation cannot be proven unprotected.
#[test]
fn word_2003_dops_must_reach_the_enforcement_unit() {
    let dop = fixture_dop("FloatingPictures.doc");
    assert_eq!(dop.len(), 616);
    for length in 84..600 {
        let fib = counted_fib(0x00C1, 0xA4, 2, 0x010C, length);
        assert_eq!(
            classify_bytes(&fib, &dop[..length]),
            EditProtection::Unrecognized,
            "Word 2003 DOP of {length} bytes"
        );
    }
    for length in 600..=616 {
        let fib = counted_fib(0x00C1, 0xA4, 2, 0x010C, length);
        assert_eq!(
            classify_bytes(&fib, &dop[..length]),
            EditProtection::None,
            "Word 2003 DOP of {length} bytes"
        );
        let protected = enforced(dop[..length].to_vec(), 1);
        assert_eq!(
            classify_bytes(&fib, &protected),
            EditProtection::Document,
            "enforced Word 2003 DOP of {length} bytes"
        );
    }
    // A longer DOP under the Word 2003 FIB is read the same way.
    let mut longer = dop.clone();
    longer.resize(674, 0);
    let fib = counted_fib(0x00C1, 0xA4, 2, 0x010C, 674);
    assert_eq!(classify_bytes(&fib, &longer), EditProtection::None);
    assert_eq!(
        classify_bytes(&fib, &enforced(longer, 2)),
        EditProtection::Document
    );
}

/// Only the protection-bearing DOP fields decide the verdict; MS-DOC's
/// requirements on those fields stay enforced.
#[test]
fn dop_protection_fields_decide_the_document_verdict() {
    let classify_word2007 = |dop: &[u8]| {
        let mut fib = fib_bytes(10);
        set_pointer(&mut fib, 31, 0, u32::try_from(dop.len()).unwrap());
        classify_bytes(&fib, dop)
    };
    let base = valid_modern_dop();
    assert_eq!(base.len(), 674);
    assert_eq!(classify_word2007(&base), EditProtection::None);
    let with = |bits: &[(usize, u8)]| {
        let mut dop = base.clone();
        for &(byte, mask) in bits {
            dop[byte] |= mask;
        }
        dop
    };
    const REVISION_MARKING: (usize, u8) = (5, 0x80);
    const FORM_NO_FIELDS: (usize, u8) = (5, 0x20);
    const LOCK_ANNOTATIONS: (usize, u8) = (6, 0x10);
    const PROTECT_FORMS: (usize, u8) = (7, 0x02);
    const LOCK_VBA_PROJECT: (usize, u8) = (7, 0x20);
    const LOCK_REVISIONS: (usize, u8) = (7, 0x40);

    // Each DopBase lock, and a password hash, restricts editing.
    for bits in [
        &[LOCK_ANNOTATIONS][..],
        &[PROTECT_FORMS],
        &[LOCK_REVISIONS, REVISION_MARKING],
        &[PROTECT_FORMS, FORM_NO_FIELDS],
        // SHOULD-level combinations that Word 97-2003 writes (Appendix A
        // notes <164>, <165> and <167>) are protected, not malformed.
        &[PROTECT_FORMS, LOCK_ANNOTATIONS],
        &[PROTECT_FORMS, LOCK_REVISIONS, REVISION_MARKING],
    ] {
        assert_eq!(
            classify_word2007(&with(bits)),
            EditProtection::Document,
            "{bits:?}"
        );
    }
    let mut keyed = base.clone();
    keyed[78..82].copy_from_slice(&0x1234_5678u32.to_le_bytes());
    assert_eq!(classify_word2007(&keyed), EditProtection::Document);
    assert_eq!(
        classify_word2007(&enforced(base.clone(), 0)),
        EditProtection::Document
    );

    // MS-DOC MUSTs on the protection fields themselves.
    for bits in [
        &[LOCK_ANNOTATIONS, LOCK_REVISIONS, REVISION_MARKING][..],
        &[LOCK_REVISIONS],
        &[FORM_NO_FIELDS],
    ] {
        assert_eq!(
            classify_word2007(&with(bits)),
            EditProtection::Unrecognized,
            "{bits:?}"
        );
    }

    // Revision tracking and a locked VBA project are not editing restrictions.
    assert_eq!(
        classify_word2007(&with(&[REVISION_MARKING])),
        EditProtection::None
    );
    assert_eq!(
        classify_word2007(&with(&[LOCK_VBA_PROJECT])),
        EditProtection::None
    );

    // Fields without protection meaning are not validated by the classifier,
    // including those MS-DOC says to ignore.
    let mut unrelated = base.clone();
    unrelated[0] |= 0x60; // DopBase.fpc reserved value 3
    unrelated[18..20].copy_from_slice(&0xFFFFu16.to_le_bytes()); // wSpare2
    unrelated[20..24].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes()); // dttmCreated
    unrelated[54..56].copy_from_slice(&0x4000u16.to_le_bytes()); // reserved2
    unrelated[82..84].copy_from_slice(&(0x0007u16 | (9 << 3)).to_le_bytes()); // wvkoSaved 7, pctWwdSaved 9
    unrelated[0x190..0x19A].fill(0); // Dop97.dogrid display multiples 0
    unrelated[599] = 0xFF; // Dop2003.empty2
    assert_eq!(classify_word2007(&unrelated), EditProtection::None);
}

#[test]
fn rejects_wrong_fixed_fib_counts_and_generation_csw_new() {
    let dop = valid_modern_dop();
    let mut wrong_csw = fib_bytes(10);
    wrong_csw[32..34].copy_from_slice(&0x000du16.to_le_bytes());
    set_pointer(&mut wrong_csw, 31, 0, dop.len() as u32);
    let wrong_csw = FileInformationBlock::parse(&wrong_csw).unwrap();
    assert_eq!(classify(&wrong_csw, &dop).unwrap(), EditProtection::Unknown);

    let mut wrong_cslw = fib_bytes(10);
    wrong_cslw[62..64].copy_from_slice(&0x0015u16.to_le_bytes());
    set_pointer(&mut wrong_cslw, 31, 0, dop.len() as u32);
    let wrong_cslw = FileInformationBlock::parse(&wrong_cslw).unwrap();
    assert_eq!(
        classify(&wrong_cslw, &dop).unwrap(),
        EditProtection::Unknown
    );

    let mut wrong_csw_new = fib_bytes(10);
    let pointer_end = 154 + 183 * 8;
    wrong_csw_new[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    set_pointer(&mut wrong_csw_new, 31, 0, dop.len() as u32);
    let wrong_csw_new = FileInformationBlock::parse(&wrong_csw_new).unwrap();
    assert_eq!(
        classify(&wrong_csw_new, &dop).unwrap(),
        EditProtection::Unknown
    );
}

#[test]
fn malformed_range_tables_remain_unrecognized_under_an_explicit_capability() {
    let (mut fib, table) = Tables::typical().assemble();
    let pointer = 154 + PLCF_BKL_PROT * 8;
    fib[pointer + 4..pointer + 8].copy_from_slice(&1u32.to_le_bytes());
    let parsed = FileInformationBlock::parse(&fib).unwrap();
    assert_eq!(
        classify(&parsed, &table).unwrap(),
        EditProtection::Unrecognized
    );
    let authorization = ProtectionAuthorization::audited("test-suite", "inspect ranges").unwrap();
    assert!(
        ProtectionPolicy::allow_protected(authorization)
            .authorize(EditProtection::Unrecognized)
            .is_err()
    );
}

#[test]
fn preserves_nonstandard_and_truncated_dops_as_unrecognized() {
    for length in 595..=615 {
        let mut fib = fib_bytes(10);
        let table = vec![0u8; length];
        set_pointer(&mut fib, 31, 0, u32::try_from(length).unwrap());
        let fib = FileInformationBlock::parse(&fib).unwrap();
        assert_eq!(
            classify(&fib, &table).unwrap(),
            EditProtection::Unrecognized
        );
    }

    let mut fib = fib_bytes(10);
    let mut table = valid_modern_dop();
    table[0x190..0x19a].copy_from_slice(&[0xa5, 0x06, 0xc0, 0x07, 0xb4, 0, 0xb4, 0, 1, 0x81]);
    // Dop2003 iDocProtCur values 4..6 are reserved by MS-DOC and must not
    // become an unprotected state merely because fEnforceDocProt is set.
    table[598..600].copy_from_slice(&(0x0008u16 | (4 << 4)).to_le_bytes());
    set_pointer(&mut fib, 31, 0, 674);
    let fib = FileInformationBlock::parse(&fib).unwrap();
    assert_eq!(
        classify(&fib, &table).unwrap(),
        EditProtection::Unrecognized
    );

    table[598..600].copy_from_slice(&(0x0008u16 | (7 << 4)).to_le_bytes());
    assert_eq!(classify(&fib, &table).unwrap(), EditProtection::None);
}

#[test]
fn malformed_dops_and_incomplete_fibs_fail_closed_for_both_policies() {
    let authorization = ProtectionAuthorization::audited("test-suite", "inspect malformed DOC")
        .expect("test authorization");
    let allow = ProtectionPolicy::allow_protected(authorization);

    for length in 595..=615 {
        let mut fib = fib_bytes(10);
        let table = vec![0u8; length];
        set_pointer(&mut fib, 31, 0, u32::try_from(length).unwrap());
        let fib = FileInformationBlock::parse(&fib).unwrap();
        assert_eq!(
            classify(&fib, &table).unwrap(),
            EditProtection::Unrecognized
        );
        assert!(
            ProtectionPolicy::default()
                .authorize(EditProtection::Unrecognized)
                .is_err()
        );
        assert!(allow.authorize(EditProtection::Unrecognized).is_err());
    }

    // A zero or absent lcbDop is host ambiguity, not an unprotected document;
    // even the explicit protected-edit capability cannot make the host shape
    // safe to interpret.
    let fib = FileInformationBlock::parse(&fib_bytes(10)).unwrap();
    assert_eq!(classify(&fib, &[]).unwrap(), EditProtection::Unknown);
    assert!(
        ProtectionPolicy::default()
            .authorize(EditProtection::Unknown)
            .is_err()
    );
    assert!(allow.authorize(EditProtection::Unknown).is_err());

    let mut zero_dop = fib_bytes(10);
    set_pointer(&mut zero_dop, 31, 12, 0);
    let zero_dop = FileInformationBlock::parse(&zero_dop).unwrap();
    assert_eq!(
        classify(&zero_dop, &[0u8; 12]).unwrap(),
        EditProtection::Unknown
    );
    assert!(allow.authorize(EditProtection::Unknown).is_err());

    // A DOP that does not lie inside the table stream is never read as
    // unprotected.
    let dop = valid_modern_dop();
    for offset in [1, 16, u32::MAX - 673, u32::MAX] {
        let mut outside = fib_bytes(10);
        set_pointer(&mut outside, 31, offset, 674);
        let outside = FileInformationBlock::parse(&outside).unwrap();
        assert_eq!(
            classify(&outside, &dop).unwrap(),
            EditProtection::Unrecognized,
            "DOP at offset {offset}"
        );
        assert!(allow.authorize(EditProtection::Unrecognized).is_err());
    }

    let mut truncated_fib = fib_bytes(10);
    truncated_fib.truncate(154 + 136 * 8);
    let truncated_fib = FileInformationBlock::parse(&truncated_fib).unwrap();
    assert_eq!(
        classify(&truncated_fib, &[]).unwrap(),
        EditProtection::Unknown
    );
    assert!(allow.authorize(EditProtection::Unknown).is_err());

    let mut wrong_count = fib_bytes(10);
    wrong_count[152..154].copy_from_slice(&117u16.to_le_bytes());
    let wrong_count = FileInformationBlock::parse(&wrong_count).unwrap();
    assert_eq!(
        classify(&wrong_count, &[]).unwrap(),
        EditProtection::Unknown
    );
    assert!(allow.authorize(EditProtection::Unknown).is_err());
}

/// The 35 readable DOC fixtures. None carries a DopBase lock, a password
/// hash, an enforced Dop2003 mode or range-protection tables, although 24 of
/// them were refused as `Unknown` or `Unrecognized` before change 0768.
const UNPROTECTED_FIXTURES: [&str; 35] = [
    "3endnotes.doc",
    "DiffFirstPageHeadFoot.doc",
    "FancyFoot.doc",
    "FloatingPictures.doc",
    "HeaderFooterProblematic.doc",
    "HeaderFooterUnicode.doc",
    "Lists.doc",
    "NoHeadFoot.doc",
    "PngPicture.doc",
    "ThreeColFoot.doc",
    "ThreeColHead.doc",
    "ThreeColHeadFoot.doc",
    "cfb-truncated-final-sector.doc",
    "cjklist30.doc",
    "cjklist31.doc",
    "cjklist34.doc",
    "cjklist35.doc",
    "commented-table.doc",
    "documentProperties.doc",
    "duplicate-style-names.doc",
    "empty.doc",
    "endingnote.doc",
    "equation.doc",
    "first-header-footer.doc",
    "footnote.doc",
    "hyperlink.doc",
    "image-comment-at-char.doc",
    "inline-endnote-and-footnote.doc",
    "lists-margins.doc",
    "picture.doc",
    "pictures_escher.doc",
    "table-merged-cells.doc",
    "tdf71749_with_footnote.doc",
    "testPictures.doc",
    "watermark.doc",
];

#[test]
fn every_readable_fixture_classifies_as_unprotected() {
    for name in UNPROTECTED_FIXTURES {
        let (word, table) = fixture_streams(name);
        let fib = FileInformationBlock::parse(&word).unwrap();
        assert_eq!(
            classify(&fib, &table).unwrap(),
            EditProtection::None,
            "{name}"
        );
    }
}

#[test]
fn authorization_requires_actor_and_reason() {
    assert!(ProtectionAuthorization::audited("", "reason").is_err());
    assert!(ProtectionAuthorization::audited("actor", " ").is_err());
}
