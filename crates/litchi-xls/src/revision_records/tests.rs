//! Focused regression coverage for BIFF8 revision-record codecs.

use super::{codec::*, model::*, validation::*};

/// Build a valid 14-byte RRD structure (cbMemory = 26, no flags).
fn rrd(revt: u16, revid: i32, tabid: u16) -> [u8; RRD_LEN] {
    let mut data = [0u8; RRD_LEN];
    data[0..4].copy_from_slice(&26u32.to_le_bytes());
    data[4..8].copy_from_slice(&revid.to_le_bytes());
    data[8..10].copy_from_slice(&revt.to_le_bytes());
    data[12..14].copy_from_slice(&tabid.to_le_bytes());
    data
}

fn short_dtr_bytes() -> [u8; SHORT_DTR_LEN] {
    let mut data = [0u8; SHORT_DTR_LEN];
    data[0..2].copy_from_slice(&2024u16.to_le_bytes());
    data[2] = 1;
    data[3] = 15;
    data[4] = 10;
    data[5] = 30;
    data[6] = 5;
    data[7] = 1;
    data
}

/// Fixed-size XLUnicodeStringNoCch field holding a compressed string.
fn string_field(field_len: usize, text: &str) -> Vec<u8> {
    let mut field = vec![0u8; field_len];
    field[1..1 + text.len()].copy_from_slice(text.as_bytes());
    field
}

#[test]
fn revision_type_round_trips_all_spec_values() {
    for value in [
        0x0000u16, 0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0007, 0x0008, 0x0009, 0x000A, 0x000B,
        0x000C, 0x000D, 0x0020, 0x0025, 0x002B, 0x002C, 0x002E,
    ] {
        let kind = RevisionType::from_u16(RRD_HEAD_RECORD_TYPE, value).unwrap();
        assert_eq!(kind.to_u16(), value);
    }
    assert!(RevisionType::from_u16(RRD_HEAD_RECORD_TYPE, 0x0006).is_err());
    assert!(RevisionType::from_u16(RRD_HEAD_RECORD_TYPE, 0xFFFF).is_err());
}

#[test]
fn short_dtr_validates_calendar_ranges() {
    let dtr = ShortDtr::parse(RRD_HEAD_RECORD_TYPE, &short_dtr_bytes()).unwrap();
    assert_eq!(dtr.year(), 2024);
    assert_eq!(dtr.month(), 1);
    assert_eq!(dtr.day(), 15);
    assert_eq!(dtr.hour(), 10);
    assert_eq!(dtr.minute(), 30);
    assert_eq!(dtr.second(), 5);
    assert_eq!(dtr.weekday(), 1);

    assert!(ShortDtr::parse(RRD_HEAD_RECORD_TYPE, &[0; 7]).is_err());
    for (index, value) in [
        (0usize, 0u8), // year low byte -> year 0
        (2, 0),        // month 0
        (2, 13),       // month 13
        (3, 0),        // day 0
        (3, 32),       // day 32
        (4, 24),       // hour 24
        (5, 60),       // minute 60
        (6, 60),       // second 60
        (7, 8),        // weekday 8
    ] {
        let mut data = short_dtr_bytes();
        data[index] = value;
        if index == 0 {
            data[1] = 0;
        }
        assert!(
            ShortDtr::parse(RRD_HEAD_RECORD_TYPE, &data).is_err(),
            "index {index} value {value} must be rejected"
        );
    }
}

#[test]
fn ref8u_validates_ordering_and_column_cap() {
    let mut data = [0u8; REF8U_LEN];
    data[0..2].copy_from_slice(&5u16.to_le_bytes());
    data[2..4].copy_from_slice(&9u16.to_le_bytes());
    data[4..6].copy_from_slice(&2u16.to_le_bytes());
    data[6..8].copy_from_slice(&0x00FFu16.to_le_bytes());
    let range = RevisionCellRange::parse(RRD_INS_DEL_RECORD_TYPE, &data).unwrap();
    assert_eq!(range.first_row(), 5);
    assert_eq!(range.last_row(), 9);
    assert_eq!(range.first_column(), 2);
    assert_eq!(range.last_column(), 0x00FF);

    let mut swapped_rows = data;
    swapped_rows[0..2].copy_from_slice(&10u16.to_le_bytes());
    assert!(RevisionCellRange::parse(RRD_INS_DEL_RECORD_TYPE, &swapped_rows).is_err());
    let mut over_column = data;
    over_column[6..8].copy_from_slice(&0x0100u16.to_le_bytes());
    assert!(RevisionCellRange::parse(RRD_INS_DEL_RECORD_TYPE, &over_column).is_err());
    assert!(RevisionCellRange::parse(RRD_INS_DEL_RECORD_TYPE, &data[..7]).is_err());
}

#[test]
fn rrd_header_validates_memory_flags_and_revid() {
    let header =
        RevisionRecordHeader::parse(RRD_CHG_CELL_RECORD_TYPE, &rrd(0x0008, 3, 1), false).unwrap();
    assert_eq!(header.memory_size(), 26);
    assert_eq!(header.revision_id(), 3);
    assert_eq!(header.revision_type(), RevisionType::ChangeCell);
    assert_eq!(header.tab_id(), Some(1));
    assert!(!header.is_accepted());

    let mut small_memory = rrd(0x0008, 3, 1);
    small_memory[0..4].copy_from_slice(&25u32.to_le_bytes());
    assert!(RevisionRecordHeader::parse(RRD_CHG_CELL_RECORD_TYPE, &small_memory, false).is_err());
    // RRDHead requires the sentinel instead of the minimum.
    assert!(
        RevisionRecordHeader::parse(RRD_HEAD_RECORD_TYPE, &rrd(0x0020, 0, 0xFFFF), true).is_err()
    );
    let mut head = rrd(0x0020, 0, 0xFFFF);
    head[0..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    let header = RevisionRecordHeader::parse(RRD_HEAD_RECORD_TYPE, &head, true).unwrap();
    assert_eq!(header.tab_id(), None);

    let mut flags = rrd(0x0008, 3, 1);
    flags[10] = 0x10; // reserved bit
    assert!(RevisionRecordHeader::parse(RRD_CHG_CELL_RECORD_TYPE, &flags, false).is_err());
    let mut negative = rrd(0x0008, 3, 1);
    negative[4..8].copy_from_slice(&(-1i32).to_le_bytes());
    assert!(RevisionRecordHeader::parse(RRD_CHG_CELL_RECORD_TYPE, &negative, false).is_err());
    assert!(
        RevisionRecordHeader::parse(RRD_CHG_CELL_RECORD_TYPE, &rrd(0x0008, 3, 1)[..13], false)
            .is_err()
    );
}

/// Build a valid 50-byte RRDInfo payload (shared, tracked, 30-day history).
fn rrd_info_payload() -> Vec<u8> {
    let mut data = vec![0u8; RRD_INFO_PAYLOAD_LEN];
    data[0..2].copy_from_slice(&8u16.to_le_bytes()); // wXLVer = BIFF8
    data[4..6].copy_from_slice(&0x000Bu16.to_le_bytes()); // shared|diskHasRev|revTrack
    data[6] = 0xAA; // guid marker
    data[22] = 0xBB; // guidRoot marker
    data[38..42].copy_from_slice(&7i32.to_le_bytes()); // revid
    data[42..46].copy_from_slice(&1u32.to_le_bytes()); // version
    data[46..48].copy_from_slice(&30u16.to_le_bytes()); // interval
    data
}

#[test]
fn rrd_info_parses_flags_and_guids() {
    let info = RrdInfo::parse_payload(&rrd_info_payload()).unwrap();
    assert_eq!(info.biff_version(), 8);
    assert!(info.is_shared());
    assert!(info.disk_has_revisions());
    assert!(info.track_revisions());
    assert!(!info.is_exclusive());
    assert!(!info.auto_delete_revisions());
    assert_eq!(info.guid()[0], 0xAA);
    assert_eq!(info.root_guid()[0], 0xBB);
    assert_eq!(info.revision_id(), 7);
    assert_eq!(info.version(), 1);
    assert!(!info.history_disabled());
    assert!(!info.history_protected());
    assert_eq!(info.history_interval_days(), 30);
}

#[test]
fn rrd_info_enforces_flag_consistency() {
    let assert_flag_error = |flags: u16, flags2: u16, interval: u16| {
        let mut data = rrd_info_payload();
        data[4..6].copy_from_slice(&flags.to_le_bytes());
        data[44..46].copy_from_slice(&flags2.to_le_bytes());
        data[46..48].copy_from_slice(&interval.to_le_bytes());
        assert!(
            RrdInfo::parse_payload(&data).is_err(),
            "flags {flags:#06X} flags2 {flags2:#06X} interval {interval} must fail"
        );
    };
    assert!(RrdInfo::parse_payload(&rrd_info_payload()[..49]).is_err());
    assert_flag_error(0x0011, 0, 30); // shared + exclusive
    assert_flag_error(0x0008, 0, 30); // revTrack without shared
    assert_flag_error(0x0003, 0, 30); // diskHasRev without revTrack
    assert_flag_error(0x0005, 0, 30); // revHist without revTrack
    assert_flag_error(0x000B, 0, 0); // preserved history with zero interval
    assert_flag_error(0x000B, 0x0001, 30); // fNoRevHist with nonzero interval
    assert_flag_error(0x0000, 0x0001, 0); // fNoRevHist without shared
    assert_flag_error(0x0000, 0x0002, 0); // fProtRev without shared
    assert_flag_error(0x000B, 0, 0x8000); // interval above 0x7FFF
    assert_flag_error(0x002B, 0, 30); // reserved flag bit
    assert_flag_error(0x000B, 0x0004, 30); // reserved flag bit in flags2
    let mut reserved1 = rrd_info_payload();
    reserved1[2] = 1;
    assert!(RrdInfo::parse_payload(&reserved1).is_err());
    let mut negative_revid = rrd_info_payload();
    negative_revid[38..42].copy_from_slice(&(-1i32).to_le_bytes());
    assert!(RrdInfo::parse_payload(&negative_revid).is_err());

    // Exclusive workbooks ignore the history interval entirely.
    let mut exclusive = rrd_info_payload();
    exclusive[4..6].copy_from_slice(&0x0010u16.to_le_bytes());
    exclusive[46..48].copy_from_slice(&0u16.to_le_bytes());
    assert!(RrdInfo::parse_payload(&exclusive).unwrap().is_exclusive());
}

#[test]
fn file_lock_parses_fixed_envelope() {
    let mut data = vec![0u8; FILE_LOCK_PAYLOAD_LEN];
    data[0..4].copy_from_slice(&0x0001_0002u32.to_le_bytes());
    data[4..6].copy_from_slice(&4u16.to_le_bytes());
    data[7..11].copy_from_slice(b"Yves");
    let lock = FileLock::parse_payload(&data).unwrap();
    assert_eq!(lock.purpose(), FileLockPurpose::MergingRevisions);
    assert_eq!(lock.user_name(), "Yves");
    assert_eq!(lock.unused_bytes().len(), FILE_LOCK_PAYLOAD_LEN - 6 - 1 - 4);

    assert!(FileLock::parse_payload(&data[..161]).is_err());
    let mut bad_purpose = data.clone();
    bad_purpose[0..4].copy_from_slice(&0x0001_0003u32.to_le_bytes());
    assert!(FileLock::parse_payload(&bad_purpose).is_err());
    let mut long_name = data.clone();
    long_name[4..6].copy_from_slice(&53u16.to_le_bytes());
    assert!(FileLock::parse_payload(&long_name).is_err());
    let mut reserved_flags = data;
    reserved_flags[6] = 0x02;
    assert!(FileLock::parse_payload(&reserved_flags).is_err());
}

fn usr_excl_payload(wide: bool) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&1u32.to_le_bytes()); // fExclusive
    data.extend_from_slice(&short_dtr_bytes());
    data.extend_from_slice(&4u16.to_le_bytes()); // cchUser
    let field_len = 1 + USR_EXCL_USER_FIELD_CHARS * if wide { 2 } else { 1 };
    let mut field = vec![0u8; field_len];
    if wide {
        field[0] = 1;
        for (index, unit) in "Lock".encode_utf16().enumerate() {
            field[1 + index * 2..3 + index * 2].copy_from_slice(&unit.to_le_bytes());
        }
    } else {
        field[1..5].copy_from_slice(b"Lock");
    }
    data.extend_from_slice(&field);
    data
}

#[test]
fn usr_excl_parses_both_string_widths() {
    let lock = UsrExcl::parse_payload(&usr_excl_payload(false)).unwrap();
    assert!(lock.is_exclusive());
    assert_eq!(lock.user_name(), "Lock");
    assert_eq!(lock.date_time().year(), 2024);

    let wide = UsrExcl::parse_payload(&usr_excl_payload(true)).unwrap();
    assert_eq!(wide.user_name(), "Lock");

    let mut bad_bool = usr_excl_payload(false);
    bad_bool[0..4].copy_from_slice(&2u32.to_le_bytes());
    assert!(UsrExcl::parse_payload(&bad_bool).is_err());
    let mut bad_size = usr_excl_payload(false);
    bad_size.pop();
    assert!(UsrExcl::parse_payload(&bad_size).is_err());
    let mut long_name = usr_excl_payload(false);
    long_name[12..14].copy_from_slice(&55u16.to_le_bytes());
    assert!(UsrExcl::parse_payload(&long_name).is_err());
}

/// Build a valid 158-byte RRDHead payload.
fn rrd_head_payload() -> Vec<u8> {
    let mut data = Vec::with_capacity(158);
    let mut header = rrd(0x0020, 0, 0xFFFF);
    header[0..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    data.extend_from_slice(&header);
    data.extend_from_slice(&[0xCC; GUID_LEN]);
    data.extend_from_slice(&1200u16.to_le_bytes()); // Unicode code page
    data.extend_from_slice(&5u16.to_le_bytes()); // cchUser
    data.extend_from_slice(&string_field(RRD_HEAD_USER_FIELD_LEN, "Alice"));
    data.extend_from_slice(&short_dtr_bytes());
    data.extend_from_slice(&3i16.to_le_bytes()); // tabidMac
    data
}

#[test]
fn rrd_head_parses_metadata() {
    let head = RrdHead::parse_payload(&rrd_head_payload()).unwrap();
    assert_eq!(head.guid(), &[0xCC; GUID_LEN]);
    assert_eq!(head.code_page(), 1200);
    assert_eq!(head.user_name(), "Alice");
    assert_eq!(head.saved_at().day(), 15);
    assert_eq!(head.next_tab_id(), 3);

    assert!(RrdHead::parse_payload(&rrd_head_payload()[..157]).is_err());
    let mut bad_revt = rrd_head_payload();
    bad_revt[8..10].copy_from_slice(&0x0008u16.to_le_bytes());
    assert!(RrdHead::parse_payload(&bad_revt).is_err());
    let mut bad_revid = rrd_head_payload();
    bad_revid[4..8].copy_from_slice(&1i32.to_le_bytes());
    assert!(RrdHead::parse_payload(&bad_revid).is_err());
    let mut long_name = rrd_head_payload();
    long_name[32..34].copy_from_slice(&55u16.to_le_bytes());
    assert!(RrdHead::parse_payload(&long_name).is_err());
    let mut bad_tabid_mac = rrd_head_payload();
    bad_tabid_mac[156..158].copy_from_slice(&(-2i16).to_le_bytes());
    assert!(RrdHead::parse_payload(&bad_tabid_mac).is_err());

    // tabidMac = -1 is legal.
    let mut minus_one = rrd_head_payload();
    minus_one[156..158].copy_from_slice(&(-1i16).to_le_bytes());
    assert_eq!(
        RrdHead::parse_payload(&minus_one).unwrap().next_tab_id(),
        -1
    );
}

#[test]
fn rr_tab_id_parses_identifier_array() {
    let mut data = Vec::new();
    for id in [1u16, 7, 42] {
        data.extend_from_slice(&id.to_le_bytes());
    }
    assert_eq!(
        RrTabId::parse_payload(&data).unwrap().sheet_ids(),
        &[1, 7, 42]
    );
    assert!(RrTabId::parse_payload(&data[..5]).is_err());
    let mut too_many = Vec::new();
    for id in 0..4113u16 {
        too_many.extend_from_slice(&id.to_le_bytes());
    }
    assert!(RrTabId::parse_payload(&too_many).is_err());
}

/// Build a valid 528-byte RRDRenSheet payload.
fn ren_sheet_payload() -> Vec<u8> {
    let mut data = Vec::with_capacity(528);
    data.extend_from_slice(&rrd(0x0009, 11, 2));
    data.extend_from_slice(&4u16.to_le_bytes());
    data.extend_from_slice(&string_field(REN_SHEET_NAME_FIELD_LEN, "Old1"));
    data.extend_from_slice(&4u16.to_le_bytes());
    data.extend_from_slice(&string_field(REN_SHEET_NAME_FIELD_LEN, "New1"));
    data
}

#[test]
fn ren_sheet_parses_old_and_new_names() {
    let sheet = RrdRenSheet::parse_payload(&ren_sheet_payload()).unwrap();
    assert_eq!(sheet.old_name(), "Old1");
    assert_eq!(sheet.new_name(), "New1");
    assert_eq!(sheet.header().tab_id(), Some(2));

    assert!(RrdRenSheet::parse_payload(&ren_sheet_payload()[..527]).is_err());
    let mut bad_revt = ren_sheet_payload();
    bad_revt[8..10].copy_from_slice(&0x0008u16.to_le_bytes());
    assert!(RrdRenSheet::parse_payload(&bad_revt).is_err());
    let mut no_sheet = ren_sheet_payload();
    no_sheet[12..14].copy_from_slice(&0xFFFFu16.to_le_bytes());
    assert!(RrdRenSheet::parse_payload(&no_sheet).is_err());
    let mut zero_revid = ren_sheet_payload();
    zero_revid[4..8].copy_from_slice(&0i32.to_le_bytes());
    assert!(RrdRenSheet::parse_payload(&zero_revid).is_err());
    let mut long_name = ren_sheet_payload();
    long_name[14..16].copy_from_slice(&228u16.to_le_bytes());
    assert!(RrdRenSheet::parse_payload(&long_name).is_err());
    // UTF-16 names are limited to 127 characters.
    let mut wide = ren_sheet_payload();
    wide[14..16].copy_from_slice(&128u16.to_le_bytes());
    wide[16] = 1;
    assert!(RrdRenSheet::parse_payload(&wide).is_err());
}

fn ins_del_payload(revt: u16, flags: u16) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&rrd(revt, 12, 3));
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&5u16.to_le_bytes()); // rwFirst
    data.extend_from_slice(&9u16.to_le_bytes()); // rwLast
    data.extend_from_slice(&0u16.to_le_bytes()); // colFirst
    data.extend_from_slice(&3u16.to_le_bytes()); // colLast
    data.extend_from_slice(&0u32.to_le_bytes()); // cUcr
    data
}

#[test]
fn ins_del_validates_revision_kind_and_undo() {
    let insert = RrdInsDel::parse_payload(&ins_del_payload(0x0000, 0x0001)).unwrap();
    assert_eq!(insert.header().revision_type(), RevisionType::InsertRow);
    assert!(insert.is_end_of_list());
    assert_eq!(insert.range().first_row(), 5);
    assert_eq!(insert.range().last_column(), 3);
    assert_eq!(insert.undo_count(), 0);
    assert!(insert.undo_data().is_empty());

    // fEndOfList is only meaningful for row inserts.
    assert!(RrdInsDel::parse_payload(&ins_del_payload(0x0002, 0x0001)).is_err());
    // Sort revisions are not insert/delete revisions.
    assert!(RrdInsDel::parse_payload(&ins_del_payload(0x0007, 0)).is_err());
    // Reserved flags.
    assert!(RrdInsDel::parse_payload(&ins_del_payload(0x0000, 0x0002)).is_err());

    // Undo bytes require a matching count.
    let mut stray_undo = ins_del_payload(0x0000, 0);
    stray_undo.extend_from_slice(&[1, 2, 3]);
    assert!(RrdInsDel::parse_payload(&stray_undo).is_err());
    let mut missing_undo = ins_del_payload(0x0000, 0);
    let len = missing_undo.len();
    missing_undo[len - 4..].copy_from_slice(&2u32.to_le_bytes());
    assert!(RrdInsDel::parse_payload(&missing_undo).is_err());

    // Raw Ducr bytes are preserved when the count is nonzero.
    let mut with_undo = ins_del_payload(0x0000, 0);
    let len = with_undo.len();
    with_undo[len - 4..].copy_from_slice(&1u32.to_le_bytes());
    with_undo.extend_from_slice(&[0xDE, 0xAD, 0xBE, 0xEF]);
    let parsed = RrdInsDel::parse_payload(&with_undo).unwrap();
    assert_eq!(parsed.undo_count(), 1);
    assert_eq!(parsed.undo_data(), &[0xDE, 0xAD, 0xBE, 0xEF]);
}

fn move_payload() -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&rrd(0x0004, 13, 1));
    for (first, last) in [(1u16, 2u16), (10, 12)] {
        data.extend_from_slice(&first.to_le_bytes());
        data.extend_from_slice(&last.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes());
    }
    data.extend_from_slice(&1u16.to_le_bytes()); // tabidSrc
    data.extend_from_slice(&0u32.to_le_bytes()); // cUcr
    data
}

#[test]
fn move_parses_source_and_destination() {
    let moved = RrdMove::parse_payload(&move_payload()).unwrap();
    assert_eq!(moved.source().first_row(), 1);
    assert_eq!(moved.destination().first_row(), 10);
    assert_eq!(moved.source_tab_id(), 1);

    let mut bad_revt = move_payload();
    bad_revt[8..10].copy_from_slice(&0x0000u16.to_le_bytes());
    assert!(RrdMove::parse_payload(&bad_revt).is_err());
    let mut no_sheet = move_payload();
    no_sheet[12..14].copy_from_slice(&0xFFFFu16.to_le_bytes());
    assert!(RrdMove::parse_payload(&no_sheet).is_err());
    assert!(RrdMove::parse_payload(&move_payload()[..20]).is_err());
}

fn insert_sh_payload() -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&rrd(0x0005, 14, 4));
    data.extend_from_slice(&1u16.to_le_bytes()); // itabPos
    data.extend_from_slice(&0u16.to_le_bytes()); // reserved
    data.extend_from_slice(&5u16.to_le_bytes()); // cch
    data.extend_from_slice(&string_field(INSERT_SH_NAME_FIELD_LEN, "Added"));
    data
}

#[test]
fn insert_sh_parses_position_and_name() {
    let sheet = RrInsertSh::parse_payload(&insert_sh_payload()).unwrap();
    assert_eq!(sheet.position(), 1);
    assert_eq!(sheet.name(), "Added");
    assert_eq!(sheet.header().tab_id(), Some(4));

    let mut reserved = insert_sh_payload();
    reserved[16..18].copy_from_slice(&1u16.to_le_bytes());
    assert!(RrInsertSh::parse_payload(&reserved).is_err());
    let mut bad_revt = insert_sh_payload();
    bad_revt[8..10].copy_from_slice(&0x0009u16.to_le_bytes());
    assert!(RrInsertSh::parse_payload(&bad_revt).is_err());
    assert!(RrInsertSh::parse_payload(&insert_sh_payload()[..275]).is_err());
}

/// Build an RRDChgCell payload: blank old value, Xnum new value.
fn chg_cell_payload() -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&rrd(0x0008, 15, 1));
    // vt = Xnum (2), vtOld = Blank (0), ifmtDisp = 0, fPhShow set.
    let flags: u32 = 0x0002 | 0x0100_0000;
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&7u16.to_le_bytes()); // row
    data.extend_from_slice(&2u16.to_le_bytes()); // column
    data.extend_from_slice(&0u32.to_le_bytes()); // cbOldVal
    data.extend_from_slice(&1u16.to_le_bytes()); // cetxpRst
    data.extend_from_slice(&3.5f64.to_le_bytes()); // num
    data
}

#[test]
fn chg_cell_parses_flags_and_location() {
    let cell = RrdChgCell::parse_payload(&chg_cell_payload()).unwrap();
    assert_eq!(cell.new_content(), RevisionCellContent::Xnum);
    assert_eq!(cell.old_content(), RevisionCellContent::Blank);
    assert!(cell.phonetic_shown());
    assert!(!cell.old_phonetic_shown());
    assert!(!cell.has_old_format());
    assert_eq!(cell.location().row(), 7);
    assert_eq!(cell.location().column(), 2);
    assert!(!cell.location().is_row_relative());
    assert_eq!(cell.old_value_size(), 0);
    assert_eq!(cell.formatting_run_count(), 1);
    assert_eq!(cell.tail(), &3.5f64.to_le_bytes());

    let mut bad_revt = chg_cell_payload();
    bad_revt[8..10].copy_from_slice(&0x0004u16.to_le_bytes());
    assert!(RrdChgCell::parse_payload(&bad_revt).is_err());
    let mut edge_of_sort = chg_cell_payload();
    edge_of_sort[10] = 0x08;
    assert!(RrdChgCell::parse_payload(&edge_of_sort).is_err());
    let mut no_sheet = chg_cell_payload();
    no_sheet[12..14].copy_from_slice(&0xFFFFu16.to_le_bytes());
    assert!(RrdChgCell::parse_payload(&no_sheet).is_err());
    let mut reserved = chg_cell_payload();
    reserved[17] |= 0x80; // reserved2 bit
    assert!(RrdChgCell::parse_payload(&reserved).is_err());
    assert!(RrdChgCell::parse_payload(&chg_cell_payload()[..25]).is_err());
}

#[test]
fn chg_cell_validates_old_value_size_table() {
    // vtOld = RkNumber (1) requires cbOldVal = 4.
    let mut data = Vec::new();
    data.extend_from_slice(&rrd(0x0008, 15, 1));
    data.extend_from_slice(&(1u32 << 3).to_le_bytes()); // vt = Blank, vtOld = Rk
    data.extend_from_slice(&7u16.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    data.extend_from_slice(&4u32.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&42u32.to_le_bytes()); // rkOld
    let cell = RrdChgCell::parse_payload(&data).unwrap();
    assert_eq!(cell.old_content(), RevisionCellContent::RkNumber);
    assert_eq!(cell.old_value_size(), 4);

    let mut wrong_size = data.clone();
    wrong_size[22..26].copy_from_slice(&5u32.to_le_bytes());
    assert!(RrdChgCell::parse_payload(&wrong_size).is_err());

    // An old formula must be at least 24 bytes.
    let mut formula = data.clone();
    formula[14..18].copy_from_slice(&(5u32 << 3).to_le_bytes());
    formula[22..26].copy_from_slice(&8u32.to_le_bytes());
    assert!(RrdChgCell::parse_payload(&formula).is_err());

    // Unknown content type bits (6/7) are rejected.
    let mut unknown = data;
    unknown[14..18].copy_from_slice(&7u32.to_le_bytes());
    assert!(RrdChgCell::parse_payload(&unknown).is_err());
}

#[test]
fn conflict_and_user_view_validate_rrd_invariants() {
    let conflict = RrdConflict::parse_payload(&rrd(0x0025, 16, 0xFFFF)).unwrap();
    assert_eq!(conflict.header().revision_type(), RevisionType::Conflict);
    assert!(RrdConflict::parse_payload(&rrd(0x0025, 0, 0xFFFF)).is_err());
    assert!(RrdConflict::parse_payload(&rrd(0x0008, 16, 0xFFFF)).is_err());
    assert!(RrdConflict::parse_payload(&rrd(0x0025, 16, 0xFFFF)[..13]).is_err());

    let mut view = Vec::new();
    view.extend_from_slice(&rrd(0x002B, 0, 0xFFFF));
    view.extend_from_slice(&[0x42; GUID_LEN]);
    let view = RrdUserView::parse_payload(&view).unwrap();
    assert_eq!(view.header().revision_type(), RevisionType::AddView);
    assert_eq!(view.guid(), &[0x42; GUID_LEN]);

    let mut sheet_scoped = Vec::new();
    sheet_scoped.extend_from_slice(&rrd(0x002B, 0, 1));
    sheet_scoped.extend_from_slice(&[0x42; GUID_LEN]);
    assert!(RrdUserView::parse_payload(&sheet_scoped).is_err());
    let mut bad_revid = Vec::new();
    bad_revid.extend_from_slice(&rrd(0x002B, 5, 0xFFFF));
    bad_revid.extend_from_slice(&[0x42; GUID_LEN]);
    assert!(RrdUserView::parse_payload(&bad_revid).is_err());
}

#[test]
fn empty_markers_reject_payloads() {
    assert!(validate_empty_marker(RRD_MOVE_BEGIN_RECORD_TYPE, &[], "RRDMoveBegin").is_ok());
    assert!(validate_empty_marker(RRD_MOVE_END_RECORD_TYPE, &[0], "RRDMoveEnd").is_err());
}
