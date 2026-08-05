use super::codec::{
    BKC_F_COL, BKC_F_NATIVE, BKC_ITC_LIM_SHIFT, PLCF_BKF_PROT, PLCF_BKL_PROT, PRTI_SIZE,
    STTB_F_EXTEND, STTB_PROT_USER, STTBF_BKMK_PROT, USER_ROLE_SIZE, parse_assignments, parse_ends,
    parse_starts, parse_users,
};
use super::model::{Mode, Ranges, Role, Selector, User};
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;

const BKC_F_PUB: u16 = 0x0080;

/// Build a minimal FIB whose table-pointer array covers indexes 0..144,
/// with a main-document length of `document_end` characters.
fn fib_bytes(document_end: u32) -> Vec<u8> {
    let pointer_count = 145usize;
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
        let mut table = Vec::new();
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
