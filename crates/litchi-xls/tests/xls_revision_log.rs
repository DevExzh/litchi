//! Integration tests for the shared-workbook `Revision Log` stream: a real
//! CFB container carrying both the `Workbook` stream and a `Revision Log`
//! stream is opened through `Workbook`.

use std::io::Cursor;

use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::{REVISION_LOG_STREAM_NAME, Revision, RevisionType, Workbook, Writer};

// BIFF8 record types (MS-XLS 2.3 record enumeration values).
const EOF_RECORD_TYPE: u16 = 0x000A;
const RRD_INFO_RECORD_TYPE: u16 = 0x0196;
const RRD_HEAD_RECORD_TYPE: u16 = 0x0138;
const RR_TAB_ID_RECORD_TYPE: u16 = 0x013D;
const RRD_REN_SHEET_RECORD_TYPE: u16 = 0x013E;

// MS-XLS 2.5.212 RevisionType values used by the fixture.
const REVT_RENAME_SHEET: u16 = 0x0009;
const REVT_HEADER: u16 = 0x0020;

// RRD structure invariants (MS-XLS 2.5.220).
const RRD_MIN_MEMORY_SIZE: u32 = 26;
const RRD_HEAD_MEMORY_SENTINEL: u32 = 0xFFFF_FFFF;
const NO_SHEET_TAB_ID: u16 = 0xFFFF;

// RRDInfo fixture values.
const BIFF8_VERSION: u16 = 8;
const INFO_SHARED_AND_TRACKED: u16 = 0x000B;
const HISTORY_INTERVAL_DAYS: u16 = 60;
const UNICODE_CODE_PAGE: u16 = 1200;

fn record(record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut data = Vec::with_capacity(4 + payload.len());
    data.extend_from_slice(&record_type.to_le_bytes());
    data.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    data.extend_from_slice(payload);
    data
}

fn rrd(revision_type: u16, revision_id: i32, tab_id: u16) -> [u8; 14] {
    let mut data = [0u8; 14];
    data[0..4].copy_from_slice(&RRD_MIN_MEMORY_SIZE.to_le_bytes());
    data[4..8].copy_from_slice(&revision_id.to_le_bytes());
    data[8..10].copy_from_slice(&revision_type.to_le_bytes());
    data[12..14].copy_from_slice(&tab_id.to_le_bytes());
    data
}

fn string_field(field_len: usize, text: &str) -> Vec<u8> {
    let mut field = vec![0u8; field_len];
    field[1..1 + text.len()].copy_from_slice(text.as_bytes());
    field
}

fn rrd_info() -> Vec<u8> {
    let mut data = vec![0u8; 50];
    data[0..2].copy_from_slice(&BIFF8_VERSION.to_le_bytes());
    data[4..6].copy_from_slice(&INFO_SHARED_AND_TRACKED.to_le_bytes());
    data[38..42].copy_from_slice(&41i32.to_le_bytes()); // revid
    data[42..46].copy_from_slice(&2u32.to_le_bytes()); // version
    data[46..48].copy_from_slice(&HISTORY_INTERVAL_DAYS.to_le_bytes());
    record(RRD_INFO_RECORD_TYPE, &data)
}

fn rrd_head() -> Vec<u8> {
    let mut data = Vec::with_capacity(158);
    let mut header = rrd(REVT_HEADER, 0, NO_SHEET_TAB_ID);
    header[0..4].copy_from_slice(&RRD_HEAD_MEMORY_SENTINEL.to_le_bytes());
    data.extend_from_slice(&header);
    data.extend_from_slice(&[0x5A; 16]); // revision-set GUID
    data.extend_from_slice(&UNICODE_CODE_PAGE.to_le_bytes());
    data.extend_from_slice(&7u16.to_le_bytes()); // cchUser
    data.extend_from_slice(&string_field(114, "Shared1"));
    // ShortDTR: 2024-02-29 08:45:10, weekday unspecified.
    data.extend_from_slice(&2024u16.to_le_bytes());
    data.extend_from_slice(&[2, 29, 8, 45, 10, 0]);
    data.extend_from_slice(&2i16.to_le_bytes()); // tabidMac
    record(RRD_HEAD_RECORD_TYPE, &data)
}

fn rr_tab_id() -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&1u16.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    record(RR_TAB_ID_RECORD_TYPE, &data)
}

fn rrd_ren_sheet() -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&rrd(REVT_RENAME_SHEET, 41, 1));
    data.extend_from_slice(&6u16.to_le_bytes());
    data.extend_from_slice(&string_field(255, "Sheet1"));
    data.extend_from_slice(&7u16.to_le_bytes());
    data.extend_from_slice(&string_field(255, "Summary"));
    record(RRD_REN_SHEET_RECORD_TYPE, &data)
}

fn revision_log_stream() -> Vec<u8> {
    let mut stream = Vec::new();
    stream.extend_from_slice(&rrd_info());
    stream.extend_from_slice(&rrd_head());
    stream.extend_from_slice(&rr_tab_id());
    stream.extend_from_slice(&rrd_ren_sheet());
    stream.extend_from_slice(&record(EOF_RECORD_TYPE, &[]));
    stream
}

/// Write a minimal workbook with `Writer`, then move its `Workbook`
/// stream into a fresh container that also holds the supplied extra streams.
fn workbook_container(extra_streams: &[(&str, Vec<u8>)]) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_string(sheet, 0, 0, "shared").unwrap();
    let mut workbook_bytes = Cursor::new(Vec::new());
    writer.write_to(&mut workbook_bytes).unwrap();

    let mut ole = OleFile::open(Cursor::new(workbook_bytes.into_inner())).unwrap();
    let workbook_stream = ole.open_stream(&["Workbook"]).unwrap();

    let mut container = OleWriter::new();
    container
        .create_stream(&["Workbook"], &workbook_stream)
        .unwrap();
    for (name, data) in extra_streams {
        container.create_stream(&[name], data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    container.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn workbook_without_revision_log_reports_none() {
    let bytes = workbook_container(&[]);
    let mut workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(!workbook.has_revision_log());
    assert!(workbook.revision_log().unwrap().is_none());
}

#[test]
fn workbook_exposes_revision_log_stream() {
    let bytes = workbook_container(&[(REVISION_LOG_STREAM_NAME, revision_log_stream())]);
    let mut workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.has_revision_log());

    let log = workbook
        .revision_log()
        .unwrap()
        .expect("revision log present");
    assert!(log.info().is_shared());
    assert!(log.info().track_revisions());
    assert_eq!(log.info().revision_id(), 41);
    assert_eq!(log.info().history_interval_days(), HISTORY_INTERVAL_DAYS);
    assert!(log.file_lock().is_none());
    assert!(log.exclusive_lock().is_none());

    assert_eq!(log.headers().len(), 1);
    let header = &log.headers()[0];
    assert_eq!(header.head().user_name(), "Shared1");
    assert_eq!(header.head().code_page(), UNICODE_CODE_PAGE);
    assert_eq!(header.head().saved_at().year(), 2024);
    assert_eq!(header.head().saved_at().day(), 29);
    assert_eq!(header.head().next_tab_id(), 2);
    assert_eq!(header.sheet_ids().unwrap().sheet_ids(), &[1, 2]);

    assert_eq!(header.revisions().len(), 1);
    match &header.revisions()[0] {
        Revision::RenSheet(sheet) => {
            assert_eq!(sheet.header().revision_type(), RevisionType::RenameSheet);
            assert_eq!(sheet.header().revision_id(), 41);
            assert_eq!(sheet.old_name(), "Sheet1");
            assert_eq!(sheet.new_name(), "Summary");
        },
        other => panic!("expected a sheet-rename revision, got {other:?}"),
    }
}

#[test]
fn malformed_revision_log_surfaces_an_error() {
    // A `Revision Log` stream that does not start with RRDInfo.
    let malformed = record(EOF_RECORD_TYPE, &[]);
    let bytes = workbook_container(&[(REVISION_LOG_STREAM_NAME, malformed)]);
    let mut workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.has_revision_log());
    assert!(workbook.revision_log().is_err());
}
