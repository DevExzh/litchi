use super::codec::{
    BKC_F_COL, BKC_F_NATIVE, BKC_ITC_LIM_SHIFT, PLCF_BKF_PROT, PLCF_BKL_PROT, PRTI_SIZE,
    STTB_F_EXTEND, STTB_PROT_USER, STTBF_BKMK_PROT, USER_ROLE_SIZE, parse_assignments, parse_ends,
    parse_starts, parse_users,
};
use super::model::{Mode, Ranges, Role, Selector, User};
use super::policy::{EditProtection, ProtectionAuthorization, ProtectionPolicy, classify};
use crate::package::Error as PackageError;
use crate::package::Result;
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

fn valid_modern_dop() -> Vec<u8> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/NoHeadFoot.doc");
    let bytes = std::fs::read(path).expect("modern DOC fixture");
    let mut ole = OleFile::open(Cursor::new(bytes)).expect("modern DOC CFB");
    let word = ole.open_stream(&["WordDocument"]).expect("WordDocument");
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
    let table_name = if u16::from_le_bytes([word[10], word[11]]) & 0x0200 != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table = ole.open_stream(&[table_name]).expect("table stream");
    table[offset..offset + length].to_vec()
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

#[test]
fn word_2002_requires_its_counted_fib_and_dop_shapes() {
    const POINTER_COUNT: usize = 136;
    let pointer_end = 154 + POINTER_COUNT * 8;
    let mut fib = vec![0u8; pointer_end + 4];
    fib[..2].copy_from_slice(&0xa5ecu16.to_le_bytes());
    fib[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    fib[32..34].copy_from_slice(&0x000eu16.to_le_bytes());
    fib[62..64].copy_from_slice(&0x0016u16.to_le_bytes());
    fib[152..154].copy_from_slice(&(POINTER_COUNT as u16).to_le_bytes());
    fib[pointer_end..pointer_end + 2].copy_from_slice(&2u16.to_le_bytes());
    fib[pointer_end + 2..pointer_end + 4].copy_from_slice(&0x0101u16.to_le_bytes());
    let mut dop = crate::parts::document_properties::DocumentProperties::word97_writer_bytes(
        false, false, false,
    );
    dop.resize(594, 0);
    set_pointer(&mut fib, 31, 0, dop.len() as u32);
    let parsed = FileInformationBlock::parse(&fib).unwrap();
    assert_eq!(classify(&parsed, &dop).unwrap(), EditProtection::None);

    let authorization =
        ProtectionAuthorization::audited("test-suite", "reject malformed Word 2002 shape").unwrap();
    let allow = ProtectionPolicy::allow_protected(authorization);

    let mut wrong_csw_new = fib.clone();
    wrong_csw_new[pointer_end..pointer_end + 2].copy_from_slice(&0u16.to_le_bytes());
    let word97_dop = &dop[..500];
    set_pointer(&mut wrong_csw_new, 31, 0, word97_dop.len() as u32);
    let wrong_csw_new = FileInformationBlock::parse(&wrong_csw_new).unwrap();
    assert_eq!(
        classify(&wrong_csw_new, word97_dop).unwrap(),
        EditProtection::Unknown
    );
    assert!(
        ProtectionPolicy::default()
            .authorize(EditProtection::Unknown)
            .is_err()
    );
    assert!(allow.authorize(EditProtection::Unknown).is_err());

    let mut wrong_dop_length = fib;
    let word97_dop = &dop[..500];
    set_pointer(&mut wrong_dop_length, 31, 0, word97_dop.len() as u32);
    let wrong_dop_length = FileInformationBlock::parse(&wrong_dop_length).unwrap();
    assert_eq!(
        classify(&wrong_dop_length, word97_dop).unwrap(),
        EditProtection::Unrecognized
    );
    assert!(
        ProtectionPolicy::default()
            .authorize(EditProtection::Unrecognized)
            .is_err()
    );
    assert!(allow.authorize(EditProtection::Unrecognized).is_err());
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

#[test]
fn authorization_requires_actor_and_reason() {
    assert!(ProtectionAuthorization::audited("", "reason").is_err());
    assert!(ProtectionAuthorization::audited("actor", " ").is_err());
}
