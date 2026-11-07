//! Native BIFF8 revision records and truthful legacy-export boundaries.
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::revision_records::{
    RevisionCellContent as Content, RevisionType, RrInsertSh, RrTabId, RrdChgCell, RrdHead, RrdInfo,
};
use litchi_xls::user_names::{Limits, Snapshot};
use litchi_xls::{Error, Package, SourceLimits, raw::Reader};
use std::{collections::BTreeMap, io::Cursor};
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-revisions.xls"),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ["Revision Log", "User Names", "Workbook"]
        .into_iter()
        .map(|n| (n.into(), ole.open_stream(&[n]).unwrap()))
        .collect()
}
fn records(bytes: &[u8]) -> Vec<(u16, &[u8])> {
    let mut out = Vec::new();
    let mut at = 0;
    while at < bytes.len() {
        let kind = u16::from_le_bytes(bytes[at..at + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            bytes[at + 2..at + 4].try_into().unwrap(),
        ));
        out.push((kind, &bytes[at + 4..at + 4 + len]));
        at += 4 + len;
    }
    assert_eq!(at, bytes.len());
    out
}
fn mapping_error<T>(r: Result<T, Error>) {
    assert!(
        matches!(r,Err(Error::InvalidRecord{record_type:0x13d,message}) if message=="RRTabId contains duplicate sheet identifier 0")
    );
}

#[test]
fn native_revision_records_preserve_history_fields_changes_and_header_dates() {
    let bytes = native();
    let parts = streams(&bytes);
    let raw = &parts["Revision Log"];
    assert_eq!(raw.len(), 1922);
    let recs = records(raw);
    assert_eq!(recs.len(), 24);
    let info = RrdInfo::parse_payload(recs[0].1).unwrap();
    assert_eq!(info.biff_version(), 6);
    assert!(
        info.is_shared()
            && info.track_revisions()
            && info.disk_has_revisions()
            && info.auto_delete_revisions()
    );
    assert!(!info.is_exclusive() && !info.history_disabled() && !info.history_protected());
    assert_eq!(info.revision_id(), 6);
    assert_eq!(info.version(), 2);
    assert_eq!(info.history_interval_days(), 30);
    assert_eq!(info.guid(), info.root_guid());
    let heads = recs
        .iter()
        .filter(|(k, _)| *k == 0x138)
        .map(|(_, p)| RrdHead::parse_payload(p).unwrap())
        .collect::<Vec<_>>();
    assert_eq!(heads.len(), 5);
    assert_eq!(heads.last().unwrap().guid(), info.guid());
    for (head, (minute, second, next)) in heads.iter().zip([
        (56, 27, 3),
        (54, 12, 2),
        (54, 52, 3),
        (55, 15, 3),
        (56, 27, 3),
    ]) {
        assert_eq!(head.user_name(), "Ryker Zhu");
        assert_eq!(head.code_page(), 1200);
        let t = head.saved_at();
        assert_eq!(
            (
                t.year(),
                t.month(),
                t.day(),
                t.hour(),
                t.minute(),
                t.second(),
                t.weekday()
            ),
            (2026, 10, 15, 18, minute, second, 4)
        );
        assert_eq!(head.next_tab_id(), next);
    }
    let changes = recs
        .iter()
        .filter(|(k, _)| *k == 0x13b)
        .map(|(_, p)| RrdChgCell::parse_payload(p).unwrap())
        .collect::<Vec<_>>();
    assert_eq!(changes.len(), 5);
    for (c, (id, sheet, row, col, old, new, size)) in changes.iter().zip([
        (1, 1, 1, 1, Content::Xnum, Content::Xnum, 8),
        (2, 1, 2, 1, Content::Formula, Content::Formula, 24),
        (
            3,
            1,
            3,
            1,
            Content::RichExtendedString,
            Content::RichExtendedString,
            30,
        ),
        (5, 2, 0, 0, Content::Blank, Content::RichExtendedString, 0),
        (6, 2, 0, 1, Content::Blank, Content::RichExtendedString, 0),
    ]) {
        assert_eq!(c.header().revision_id(), id);
        assert_eq!(c.header().tab_id(), Some(sheet));
        assert_eq!(c.header().revision_type(), RevisionType::ChangeCell);
        assert!(!c.header().is_accepted() && !c.header().is_undo_action());
        assert_eq!((c.location().row(), c.location().column()), (row, col));
        assert_eq!(
            (c.old_content(), c.new_content(), c.old_value_size()),
            (old, new, size)
        );
        assert_eq!(c.formatting_run_count(), 0);
    }
    assert_eq!(
        changes[0].tail(),
        [10f64.to_le_bytes(), 25f64.to_le_bytes()].concat()
    );
    assert_eq!(
        changes[2].tail(),
        [b"\x0c\0\0BASELINE_014".as_slice(), b"\x0b\0\0UPDATED_014"].concat()
    );
    assert_eq!(
        changes[2].tail().len(),
        29,
        "normalized size 30 is not a wire slice length"
    );
    assert_eq!(changes[3].tail(), b"\x0d\0\0NEW_SHEET_014");
    assert_eq!(changes[4].tail(), b"\x14\0\0REVISION_CREATED_014");
    let inserted =
        RrInsertSh::parse_payload(recs.iter().find(|(k, _)| *k == 0x14d).unwrap().1).unwrap();
    assert_eq!(inserted.header().revision_id(), 4);
    assert!(inserted.header().is_accepted());
    assert_eq!(inserted.header().tab_id(), Some(2));
    assert_eq!(inserted.position(), 1);
    assert_eq!(inserted.name(), "[excel-revisions.xlsx]Review_014");
}

#[test]
fn native_user_names_version_is_outside_the_typed_profile() {
    let bytes = native();
    let parts = streams(&bytes);
    let raw = &parts["User Names"];
    assert_eq!(raw.len(), 536);
    let recs = records(raw);
    assert_eq!(
        recs.iter().map(|(kind, _)| *kind).collect::<Vec<_>>(),
        [0x191, 0x198, 0x192, 0x197]
    );
    assert_eq!(recs[0].1, [0, 0]);
    assert_eq!(recs[1].1, [6, 0, 0, 0]);
    assert_eq!(&recs[2].1[..2], &[41, 0]);
    assert!(recs[2].1[2..].iter().all(|b| *b == 0));
    assert_eq!(recs[3].1, [0, 0]);
    // MS-XLS 2.4.338 requires 0200/0300/0400/0500/0600, not 0006.
    // Keep the native artifact intact and record this fixture-profile gap.
    assert!(matches!(Snapshot::parse(raw, Limits::default()),
        Err(Error::InvalidRecord { record_type: 0x198, message })
            if message == "UsrChk version 0x0006 is not a BIFF2-BIFF8 value"));
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().to_bytes().unwrap(),
        bytes
    );
}

#[test]
fn native_legacy_history_with_duplicate_sheet_ids_refuses_typed_publication_without_repair() {
    let bytes = native();
    let parts = streams(&bytes);
    let raw = &parts["Revision Log"];
    let maps = records(raw)
        .into_iter()
        .filter(|(k, _)| *k == 0x13d)
        .map(|(_, p)| p)
        .collect::<Vec<_>>();
    assert_eq!(
        maps,
        [
            b"\x01\0\x02\0".as_slice(),
            b"\xff\xff",
            b"\0\0\0\0",
            b"\0\0\0\0",
            b"\x01\0\x02\0"
        ]
    );
    assert_eq!(RrTabId::parse_payload(maps[0]).unwrap().sheet_ids(), [1, 2]);
    mapping_error(RrTabId::parse_payload(maps[2]));
    let mut r = Reader::new(Cursor::new(&bytes)).unwrap();
    assert_eq!(r.sheets().len(), 2);
    let sheet = r.xls_worksheet(0).unwrap();
    assert_eq!(
        sheet.get_cell(1, 1).unwrap().value(),
        &CellValue::Float(25.)
    );
    assert_eq!(
        sheet.get_cell(2, 1).unwrap().value(),
        &CellValue::Float(75.)
    );
    assert_eq!(sheet.get_cell(2, 1).unwrap().formula(), Some("=(B2*3)"));
    mapping_error(r.revision_log(&SourceLimits::default()));
    let p = Package::from_bytes(&bytes).unwrap();
    mapping_error(p.user_names());
    mapping_error(p.edit_user_names().unwrap().commit());
    assert_eq!(p.to_bytes().unwrap(), bytes);
    let f = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    p.save(f.path()).unwrap();
    assert_eq!(Package::open(f.path()).unwrap().to_bytes().unwrap(), bytes);
}
