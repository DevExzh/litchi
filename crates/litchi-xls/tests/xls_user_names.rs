//! Package-level reopen coverage for the MS-XLS User Names stream owner.

use std::io::Cursor;

use litchi_cfb::{DirectoryEntry, OleFile, OleWriter};
use litchi_xls::{UserNamesLimits, UserNamesPackageSnapshot, Workbook, Writer};

const USER_NAMES_STREAM_NAME: &str = "User Names";
const REVISION_LOG_STREAM_NAME: &str = "Revision Log";
const C_USR: u16 = 401;
const USR_CHK: u16 = 408;
const CB_USR: u16 = 402;
const BC_USRS: u16 = 407;
const USR_INFO: u16 = 403;
const RRD_INFO: u16 = 0x0196;
const RRD_HEAD: u16 = 0x0138;
const RR_TAB_ID: u16 = 0x013D;
const FILE_LOCK: u16 = 0x0195;
const USR_EXCL: u16 = 0x0194;
const RRD_INS_DEL_BEGIN: u16 = 0x0150;
const EOF: u16 = 0x000A;
const BOUND_SHEET8: u16 = 0x0085;
const BOF: u16 = 0x0809;
const WORKSHEET_BOF: u16 = 0x0010;
const WINDOW1: u16 = 0x003D;
const WINDOW2: u16 = 0x023E;
const WINDOW2_SELECTED: u16 = 0x0200;
const CODE_PAGE: u16 = 0x0042;
const FILE_PASS: u16 = 0x002F;

fn record(record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(4 + payload.len());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn user_info(guid: [u8; 16], name: &str) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&17i32.to_le_bytes());
    payload.extend_from_slice(&guid);
    payload.extend_from_slice(&[0xE8, 0x07, 1, 2, 3, 4, 5, 1]);
    payload.extend_from_slice(&(name.len() as u16).to_le_bytes());
    payload.push(0);
    payload.extend_from_slice(name.as_bytes());
    payload.push(0xA7);
    payload
}

fn user_names_stream(guid: [u8; 16], name: &str) -> Vec<u8> {
    let user = user_info(guid, name);
    let mut cbusr = vec![0u8; 512];
    cbusr[..2].copy_from_slice(&(user.len() as u16).to_le_bytes());
    let mut stream = Vec::new();
    stream.extend_from_slice(&record(C_USR, &1u16.to_le_bytes()));
    stream.extend_from_slice(&record(USR_CHK, &[0x00, 0x06, 0xAA, 0xBB]));
    stream.extend_from_slice(&record(CB_USR, &cbusr));
    stream.extend_from_slice(&record(BC_USRS, &1u16.to_le_bytes()));
    stream.extend_from_slice(&record(USR_INFO, &user));
    stream
}

fn revision_log_stream(guid: [u8; 16]) -> Vec<u8> {
    revision_log_stream_with(guid, 1200, true, 1)
}

fn revision_log_stream_with(
    guid: [u8; 16],
    code_page: u16,
    include_rr_tab_id: bool,
    next_tab_id: i16,
) -> Vec<u8> {
    let tab_ids = include_rr_tab_id.then_some(&[1u16][..]);
    revision_log_stream_with_tab_ids(guid, code_page, tab_ids, next_tab_id)
}

fn revision_log_stream_with_tab_ids(
    guid: [u8; 16],
    code_page: u16,
    tab_ids: Option<&[u16]>,
    next_tab_id: i16,
) -> Vec<u8> {
    let mut info = vec![0u8; 50];
    info[..2].copy_from_slice(&8u16.to_le_bytes());
    info[4..6].copy_from_slice(&0x000Bu16.to_le_bytes());
    info[46..48].copy_from_slice(&60u16.to_le_bytes());
    let mut head = vec![0u8; 158];
    head[..4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
    head[8..10].copy_from_slice(&0x0020u16.to_le_bytes());
    head[12..14].copy_from_slice(&0xFFFFu16.to_le_bytes());
    head[14..30].copy_from_slice(&guid);
    head[30..32].copy_from_slice(&code_page.to_le_bytes());
    head[32..34].copy_from_slice(&5u16.to_le_bytes());
    head[34] = 0;
    head[35..40].copy_from_slice(b"Alice");
    head[148..150].copy_from_slice(&2024u16.to_le_bytes());
    head[150..156].copy_from_slice(&[1, 2, 3, 4, 5, 1]);
    head[156..158].copy_from_slice(&next_tab_id.to_le_bytes());
    let mut file_lock = vec![0u8; 162];
    file_lock[0..4].copy_from_slice(&0x0001_0001u32.to_le_bytes());
    file_lock[4..6].copy_from_slice(&3u16.to_le_bytes());
    file_lock[7..10].copy_from_slice(b"Bob");
    let mut usr_excl = Vec::new();
    usr_excl.extend_from_slice(&1u32.to_le_bytes());
    usr_excl.extend_from_slice(&[0xE8, 0x07, 1, 2, 3, 4, 5, 1]);
    usr_excl.extend_from_slice(&3u16.to_le_bytes());
    usr_excl.extend_from_slice(&{
        let mut field = vec![0u8; 148];
        field[1..4].copy_from_slice(b"Bob");
        field
    });
    let mut stream = Vec::new();
    stream.extend_from_slice(&record(RRD_INFO, &info));
    stream.extend_from_slice(&record(FILE_LOCK, &file_lock));
    stream.extend_from_slice(&record(USR_EXCL, &usr_excl));
    stream.extend_from_slice(&record(RRD_HEAD, &head));
    if let Some(tab_ids) = tab_ids {
        let mut payload = Vec::with_capacity(tab_ids.len() * 2);
        for tab_id in tab_ids {
            payload.extend_from_slice(&tab_id.to_le_bytes());
        }
        stream.extend_from_slice(&record(RR_TAB_ID, &payload));
    }
    stream.extend_from_slice(&record(EOF, &[]));
    stream
}

fn workbook_container(user_names: Vec<u8>, revision_log: Vec<u8>) -> Vec<u8> {
    workbook_container_from_stream(authored_workbook_stream(), user_names, revision_log, |_| {})
}

fn authored_workbook_stream() -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_string(sheet, 0, 0, "shared").unwrap();
    let mut workbook_bytes = Cursor::new(Vec::new());
    writer.write_to(&mut workbook_bytes).unwrap();
    let mut source = OleFile::open(Cursor::new(workbook_bytes.into_inner())).unwrap();
    source.open_stream(&["Workbook"]).unwrap()
}

fn workbook_container_from_stream<F>(
    workbook_stream: Vec<u8>,
    user_names: Vec<u8>,
    revision_log: Vec<u8>,
    configure: F,
) -> Vec<u8>
where
    F: FnOnce(&mut OleWriter),
{
    let mut container = OleWriter::new();
    container
        .create_stream(&["Workbook"], &workbook_stream)
        .unwrap();
    container
        .create_stream(&[USER_NAMES_STREAM_NAME], &user_names)
        .unwrap();
    container
        .create_stream(&[REVISION_LOG_STREAM_NAME], &revision_log)
        .unwrap();
    configure(&mut container);
    let mut output = Cursor::new(Vec::new());
    container.write_to(&mut output).unwrap();
    output.into_inner()
}

fn replace_workbook_record_kind(mut workbook: Vec<u8>, from: u16, to: u16) -> Vec<u8> {
    let mut offset = 0usize;
    while offset + 4 <= workbook.len() {
        let record_type = u16::from_le_bytes([workbook[offset], workbook[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([
            workbook[offset + 2],
            workbook[offset + 3],
        ]));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(payload_len))
            .expect("BIFF record end");
        assert!(end <= workbook.len(), "truncated workbook record");
        if record_type == from {
            workbook[offset..offset + 2].copy_from_slice(&to.to_le_bytes());
            return workbook;
        }
        offset = end;
    }
    panic!("workbook did not contain record 0x{from:04X}");
}

fn strip_workbook_record_type(data: &[u8], omitted: u16) -> Vec<u8> {
    let mut retained = Vec::with_capacity(data.len());
    let mut offset = 0usize;
    while offset + 4 <= data.len() {
        let record_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(payload_len))
            .expect("BIFF record end");
        assert!(end <= data.len(), "truncated workbook record");
        if record_type != omitted {
            retained.extend_from_slice(&data[offset..end]);
        }
        offset = end;
    }
    assert_eq!(offset, data.len(), "trailing workbook bytes");
    retained
}

fn insert_workbook_record_before(data: &[u8], before: u16, inserted: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(data.len() + inserted.len());
    let mut offset = 0usize;
    let mut found = false;
    while offset + 4 <= data.len() {
        let record_type = u16::from_le_bytes([data[offset], data[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([data[offset + 2], data[offset + 3]]));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(payload_len))
            .expect("BIFF record end");
        assert!(end <= data.len(), "truncated workbook record");
        if record_type == before {
            output.extend_from_slice(inserted);
            found = true;
        }
        output.extend_from_slice(&data[offset..end]);
        offset = end;
    }
    assert_eq!(offset, data.len(), "trailing workbook bytes");
    assert!(found, "workbook did not contain record 0x{before:04X}");
    output
}

fn workbook_with_sheet_count(sheet_count: usize, include_rr_tab_id: bool) -> Vec<u8> {
    assert!(sheet_count > 1);
    let source = authored_workbook_stream();
    let mut offset = 0usize;
    let mut bound_offset = None;
    let mut bound_end = None;
    let mut worksheet_offset = None;
    while offset + 4 <= source.len() {
        let record_type = u16::from_le_bytes([source[offset], source[offset + 1]]);
        let payload_len = usize::from(u16::from_le_bytes([source[offset + 2], source[offset + 3]]));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(payload_len))
            .expect("BIFF record end");
        assert!(end <= source.len(), "truncated authored workbook record");
        if record_type == BOUND_SHEET8 && bound_offset.is_none() {
            bound_offset = Some(offset);
            bound_end = Some(end);
        }
        if record_type == BOF
            && payload_len >= 4
            && source[offset + 6..offset + 8] == WORKSHEET_BOF.to_le_bytes()
        {
            worksheet_offset = Some(offset);
            break;
        }
        offset = end;
    }
    let bound_offset = bound_offset.expect("authored workbook BoundSheet8");
    let bound_end = bound_end.expect("authored workbook BoundSheet8 end");
    let worksheet_offset = worksheet_offset.expect("authored workbook worksheet BOF");
    let bound_payload = source[bound_offset + 4..bound_end].to_vec();
    assert_eq!(bound_payload.len(), 14, "fixed six-character BoundSheet8");
    let sheet_stream = source[worksheet_offset..].to_vec();
    let globals_prefix = if include_rr_tab_id {
        let mut tab_ids = Vec::with_capacity(sheet_count * 2);
        for tab_id in 1..=sheet_count {
            tab_ids.extend_from_slice(
                &u16::try_from(tab_id)
                    .expect("RRTabId test identifier")
                    .to_le_bytes(),
            );
        }
        let prefix = strip_workbook_record_type(&source[..bound_offset], RR_TAB_ID);
        insert_workbook_record_before(&prefix, WINDOW1, &record(RR_TAB_ID, &tab_ids))
    } else {
        strip_workbook_record_type(&source[..bound_offset], RR_TAB_ID)
    };
    let globals_suffix =
        strip_workbook_record_type(&source[bound_end..worksheet_offset], RR_TAB_ID);
    let bound_record_len = 4 + bound_payload.len();
    let sheet_stream_offset = globals_prefix
        .len()
        .checked_add(
            sheet_count
                .checked_mul(bound_record_len)
                .expect("BoundSheet8 size"),
        )
        .and_then(|value| value.checked_add(globals_suffix.len()))
        .expect("first worksheet offset");

    let mut expanded = Vec::with_capacity(
        sheet_stream_offset
            .checked_add(
                sheet_count
                    .checked_mul(sheet_stream.len())
                    .expect("worksheet stream size"),
            )
            .expect("expanded Workbook size"),
    );
    expanded.extend_from_slice(&globals_prefix);
    for index in 0..sheet_count {
        let mut payload = bound_payload.clone();
        let sheet_offset = sheet_stream_offset
            .checked_add(index.checked_mul(sheet_stream.len()).expect("sheet offset"))
            .expect("sheet offset");
        payload[..4].copy_from_slice(
            &u32::try_from(sheet_offset)
                .expect("Workbook stream offset")
                .to_le_bytes(),
        );
        let name = format!("S{index:05}");
        assert_eq!(name.len(), 6);
        payload[6] = 6;
        payload[7] = 0;
        payload[8..14].copy_from_slice(name.as_bytes());
        expanded.extend_from_slice(&record(BOUND_SHEET8, &payload));
    }
    expanded.extend_from_slice(&globals_suffix);
    for index in 0..sheet_count {
        let mut sheet = sheet_stream.clone();
        if index > 0 {
            let mut sheet_offset = 0usize;
            while sheet_offset + 4 <= sheet.len() {
                let record_type =
                    u16::from_le_bytes([sheet[sheet_offset], sheet[sheet_offset + 1]]);
                let payload_len = usize::from(u16::from_le_bytes([
                    sheet[sheet_offset + 2],
                    sheet[sheet_offset + 3],
                ]));
                let end = sheet_offset
                    .checked_add(4)
                    .and_then(|value| value.checked_add(payload_len))
                    .expect("worksheet BIFF record end");
                assert!(end <= sheet.len(), "truncated worksheet record");
                if record_type == WINDOW2 {
                    assert!(payload_len >= 2, "truncated Window2 record");
                    let mut flags =
                        u16::from_le_bytes([sheet[sheet_offset + 4], sheet[sheet_offset + 5]]);
                    flags &= !WINDOW2_SELECTED;
                    sheet[sheet_offset + 4..sheet_offset + 6].copy_from_slice(&flags.to_le_bytes());
                    break;
                }
                sheet_offset = end;
            }
        }
        expanded.extend_from_slice(&sheet);
    }
    expanded
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DirectoryMetadata {
    name: String,
    entry_type: u8,
    clsid: String,
    state_bits: u32,
    creation_time: u64,
    modified_time: u64,
    size: u64,
}

fn directory_metadata(entry: &DirectoryEntry) -> DirectoryMetadata {
    DirectoryMetadata {
        name: entry.name.clone(),
        entry_type: entry.entry_type,
        clsid: entry.clsid.clone(),
        state_bits: entry.state_bits,
        creation_time: entry.creation_time,
        modified_time: entry.modified_time,
        size: entry.size,
    }
}

#[test]
fn workbook_reads_user_names_lazily_and_reopens_after_stream_edit() {
    let guid = [0x4D; 16];
    let bytes = workbook_container(user_names_stream(guid, "Alice"), revision_log_stream(guid));
    let mut workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.has_user_names());
    let source = workbook.user_names().unwrap().unwrap();
    assert_eq!(source.users()[0].user_name(), "Alice");
    assert!(source.has_revision_guid_closure());

    let mut edit = source.edit();
    edit.set_user_name(0, "Élodie").unwrap();
    let target = edit.commit().unwrap().into_snapshot();
    let reopened_bytes = workbook_container(target.finish(), revision_log_stream(guid));
    let mut reopened_workbook = Workbook::new(Cursor::new(reopened_bytes)).unwrap();
    let reopened = reopened_workbook.user_names().unwrap().unwrap();
    assert_eq!(reopened.users()[0].user_name(), "Élodie");
}

#[test]
fn workbook_rejects_unclosed_user_guid() {
    let source_guid = [0x11; 16];
    let other_guid = [0x22; 16];
    let bytes = workbook_container(
        user_names_stream(source_guid, "Alice"),
        revision_log_stream(other_guid),
    );
    let mut workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.user_names().is_err());
}

#[test]
fn workbook_user_names_requires_rrtabid_from_bound_sheet_count() {
    let guid = [0x21; 16];
    let revision = revision_log_stream_with(guid, 1200, false, 4113);
    let bytes = workbook_container(user_names_stream(guid, "Alice"), revision.clone());
    let mut workbook = Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert!(workbook.user_names().is_err());
    assert!(UserNamesPackageSnapshot::from_bytes(bytes).is_err());
}

#[test]
fn package_owner_rejects_rrtabid_cardinality_mismatch_for_one_sheet() {
    for (index, tab_ids) in [Vec::new(), vec![1u16, 2]].into_iter().enumerate() {
        let guid = [0x23 + u8::try_from(index).unwrap(); 16];
        let bytes = workbook_container(
            user_names_stream(guid, "Alice"),
            revision_log_stream_with_tab_ids(guid, 1200, Some(&tab_ids), 1),
        );
        assert!(
            UserNamesPackageSnapshot::from_bytes(bytes.clone()).is_err(),
            "RRTabId cardinality {:?} was accepted for one sheet",
            tab_ids
        );
        let mut workbook = Workbook::new(Cursor::new(bytes)).unwrap();
        assert!(workbook.user_names().is_err());
    }
}

#[test]
fn package_owner_accepts_rrtabid_omission_for_more_than_4112_sheets() {
    let guid = [0x22; 16];
    let workbook_stream = workbook_with_sheet_count(4_113, false);
    let bytes = workbook_container_from_stream(
        workbook_stream,
        user_names_stream(guid, "Alice"),
        revision_log_stream_with(guid, 1200, false, 1),
        |_| {},
    );
    let source = UserNamesPackageSnapshot::from_bytes(bytes.clone()).unwrap();
    assert_eq!(source.user_names().users()[0].user_name(), "Alice");
    let mut workbook = Workbook::new(Cursor::new(source.finish())).unwrap();
    assert_eq!(workbook.sheets().len(), 4_113);
    assert!(workbook.user_names().unwrap().is_some());
}

#[test]
fn package_owner_accepts_exact_rrtabid_cardinality_at_4112_sheets() {
    let guid = [0x24; 16];
    let tab_ids = (1..=4_112u16).collect::<Vec<_>>();
    let bytes = workbook_container_from_stream(
        workbook_with_sheet_count(4_112, true),
        user_names_stream(guid, "Alice"),
        revision_log_stream_with_tab_ids(guid, 1200, Some(&tab_ids), 4_113),
        |_| {},
    );
    let source = UserNamesPackageSnapshot::from_bytes(bytes).unwrap();
    assert_eq!(source.user_names().users()[0].user_name(), "Alice");
}

#[test]
fn workbook_rechecks_revision_log_when_user_names_bytes_are_unchanged() {
    let guid = [0x31; 16];
    let names = user_names_stream(guid, "Alice");
    let valid = workbook_container(names.clone(), revision_log_stream(guid));
    let mut workbook = Workbook::new(Cursor::new(valid)).unwrap();
    let source = workbook.user_names().unwrap().unwrap();

    // The User Names stream is byte-for-byte identical, but the current
    // Revision Log has changed. The package facade must bind the actual
    // current Revision Log source instead of trusting a stale caller-supplied
    // GUID set or a User Names-only byte identity.
    let mut changed_revision = revision_log_stream(guid);
    changed_revision[4 + 50 + 4 + 35] = b'B'; // RRDHead.stUser[0]
    let changed = workbook_container(names, changed_revision);
    let mut current = Workbook::new(Cursor::new(changed)).unwrap();
    let current_snapshot = current.user_names().unwrap().unwrap();
    assert_eq!(source.finish(), current_snapshot.finish());
    assert_ne!(source, current_snapshot);

    let mut edit = source.edit();
    edit.set_user_name(0, "Bob").unwrap();
    let commit = edit.commit().unwrap();
    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(commit.patch().revert(&applied).unwrap(), source);
    assert!(commit.patch().apply(&current_snapshot).is_err());
}

#[test]
fn package_owner_is_exact_for_noop_and_reversible_for_changed_stream() {
    let guid = [0x41; 16];
    let source_bytes =
        workbook_container(user_names_stream(guid, "Alice"), revision_log_stream(guid));
    let source = UserNamesPackageSnapshot::from_bytes(source_bytes.clone()).unwrap();
    assert_eq!(source.bytes(), source_bytes.as_slice());

    let noop = source.edit().commit().unwrap();
    assert!(!noop.changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.snapshot().bytes(), source_bytes.as_slice());

    let mut edit = source.edit();
    edit.set_user_name(0, "Bob").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().user_names().users()[0].user_name(), "Bob");
    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(commit.patch().revert(commit.snapshot()).unwrap(), source);
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );
}

#[test]
fn package_owner_accepts_official_rrdhead_code_pages_and_keeps_unicode_user_name() {
    for (index, code_page) in [437u16, 850, 20_127, 1_201, 12_000].into_iter().enumerate() {
        let guid = [0x80 + u8::try_from(index).unwrap(); 16];
        let source_bytes = workbook_container(
            user_names_stream(guid, "Alice"),
            revision_log_stream_with(guid, code_page, true, 1),
        );
        let source = UserNamesPackageSnapshot::from_bytes(source_bytes).unwrap();
        let mut edit = source.edit();
        edit.set_user_name(0, "Élodie").unwrap();
        let target = edit.commit().unwrap().into_snapshot();
        assert_eq!(target.user_names().users()[0].user_name(), "Élodie");

        let mut workbook = Workbook::new(Cursor::new(target.finish())).unwrap();
        let revision = workbook.revision_log().unwrap().unwrap();
        assert_eq!(revision.headers()[0].head().code_page(), code_page);
        assert_eq!(revision.headers()[0].head().user_name(), "Alice");
    }
}

#[test]
fn package_owner_rejects_undefined_rrdhead_code_pages() {
    for code_page in [436u16, 65_535] {
        let guid = [0x90; 16];
        let bytes = workbook_container(
            user_names_stream(guid, "Alice"),
            revision_log_stream_with(guid, code_page, true, 1),
        );
        assert!(
            UserNamesPackageSnapshot::from_bytes(bytes).is_err(),
            "undefined CODEPG {code_page} was accepted"
        );
    }
}

#[test]
fn package_owner_refuses_protected_cfb_markers_and_filepass() {
    let guid = [0xA0; 16];
    for marker in ["DigitalSignature", "EncryptedPackage"] {
        let bytes = workbook_container_from_stream(
            authored_workbook_stream(),
            user_names_stream(guid, "Alice"),
            revision_log_stream(guid),
            |container| {
                container.create_stream(&[marker], b"protected").unwrap();
            },
        );
        assert!(
            UserNamesPackageSnapshot::from_bytes(bytes).is_err(),
            "protected marker {marker} was accepted"
        );
    }

    let bytes = workbook_container_from_stream(
        replace_workbook_record_kind(authored_workbook_stream(), CODE_PAGE, FILE_PASS),
        user_names_stream(guid, "Alice"),
        revision_log_stream(guid),
        |_| {},
    );
    assert!(UserNamesPackageSnapshot::from_bytes(bytes).is_err());
}

#[test]
fn package_owner_preserves_unrelated_streams_and_directory_metadata() {
    let guid = [0xB0; 16];
    let source_bytes = workbook_container_from_stream(
        authored_workbook_stream(),
        user_names_stream(guid, "Alice"),
        revision_log_stream(guid),
        |container| {
            container.set_root_clsid([
                0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46,
            ]);
            container.set_root_state_bits(0x1020_3040);
            container.set_root_creation_time_from_source(0x0102_0304_0506_0708);
            container.set_root_modified_time(0x1112_1314_1516_1718);
            container.create_storage(&["Meta"]).unwrap();
            container
                .set_storage_metadata(
                    &["Meta"],
                    0xA5A5_5A5A,
                    0x2122_2324_2526_2728,
                    0x3132_3334_3536_3738,
                )
                .unwrap();
            container
                .create_stream(&["Unrelated"], b"unrelated stream")
                .unwrap();
            container
                .set_stream_metadata(
                    &["Unrelated"],
                    0x0BAD_CAFE,
                    0x4142_4344_4546_4748,
                    0x5152_5354_5556_5758,
                )
                .unwrap();
            container
                .create_stream(&["Meta", "Nested"], b"nested stream")
                .unwrap();
            container
                .set_stream_metadata(
                    &["Meta", "Nested"],
                    0x1234_5678,
                    0x6162_6364_6566_6768,
                    0x7172_7374_7576_7778,
                )
                .unwrap();
        },
    );
    let source = UserNamesPackageSnapshot::from_bytes(source_bytes.clone()).unwrap();
    let mut edit = source.edit();
    edit.set_user_name(0, "Bob").unwrap();
    let target = edit.commit().unwrap().into_snapshot();

    let mut source_ole = OleFile::open(Cursor::new(source_bytes)).unwrap();
    let mut target_ole = OleFile::open(Cursor::new(target.finish())).unwrap();
    assert_eq!(
        source_ole.open_stream(&["Unrelated"]).unwrap(),
        target_ole.open_stream(&["Unrelated"]).unwrap()
    );
    assert_eq!(
        source_ole.open_stream(&["Meta", "Nested"]).unwrap(),
        target_ole.open_stream(&["Meta", "Nested"]).unwrap()
    );

    let source_root = source_ole.root_entry().unwrap();
    let target_root = target_ole.root_entry().unwrap();
    let source_meta = directory_metadata(
        source_ole
            .list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Meta")
            .unwrap(),
    );
    let target_meta = directory_metadata(
        target_ole
            .list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Meta")
            .unwrap(),
    );
    let source_unrelated = directory_metadata(
        source_ole
            .list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Unrelated")
            .unwrap(),
    );
    let target_unrelated = directory_metadata(
        target_ole
            .list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Unrelated")
            .unwrap(),
    );
    let source_nested = directory_metadata(
        source_ole
            .list_directory_entries(&["Meta"])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Nested")
            .unwrap(),
    );
    let target_nested = directory_metadata(
        target_ole
            .list_directory_entries(&["Meta"])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Nested")
            .unwrap(),
    );
    assert_eq!(source_root.clsid, target_root.clsid);
    assert_eq!(source_root.state_bits, target_root.state_bits);
    assert_eq!(source_root.creation_time, target_root.creation_time);
    assert_eq!(source_root.modified_time, target_root.modified_time);
    assert_eq!(source_meta, target_meta);
    assert_eq!(source_unrelated, target_unrelated);
    assert_eq!(source_nested, target_nested);
}

#[test]
fn package_owner_rejects_stream_and_user_bounds_before_publication() {
    let guid = [0x51; 16];
    let names = user_names_stream(guid, "Alice");
    let bytes = workbook_container(names.clone(), revision_log_stream(guid));

    let too_small = UserNamesLimits::default().with_max_stream_bytes(names.len() - 1);
    assert!(UserNamesPackageSnapshot::from_bytes_with_limits(bytes.clone(), too_small).is_err());

    let no_users = UserNamesLimits::default().with_max_users(0);
    assert!(UserNamesPackageSnapshot::from_bytes_with_limits(bytes, no_users).is_err());
}

#[test]
fn package_owner_rejects_revision_scan_limits_before_guid_publication() {
    let guid = [0x61; 16];
    let bytes = workbook_container(user_names_stream(guid, "Alice"), revision_log_stream(guid));

    let too_few_records = UserNamesLimits::default().with_max_revision_records(2);
    assert!(
        UserNamesPackageSnapshot::from_bytes_with_limits(bytes.clone(), too_few_records).is_err()
    );

    let no_revision_guids = UserNamesLimits::default().with_max_revision_guids(0);
    assert!(UserNamesPackageSnapshot::from_bytes_with_limits(bytes, no_revision_guids).is_err());
}

#[test]
fn package_owner_rejects_unknown_revision_before_materializing_record_table() {
    let guid = [0x71; 16];
    let names = user_names_stream(guid, "Alice");
    let mut revision = revision_log_stream(guid);
    let unknown = record(0x7777, &[]);
    let eof_offset = revision.len() - 4;
    revision.splice(eof_offset..eof_offset, unknown);
    let bytes = workbook_container(names, revision);
    assert!(UserNamesPackageSnapshot::from_bytes(bytes).is_err());
}

#[test]
fn package_owner_rejects_malformed_known_revision_production() {
    let guid = [0x72; 16];
    let names = user_names_stream(guid, "Alice");
    let mut revision = revision_log_stream(guid);
    // RRDInsDelBegin must be followed by RRDInsDel and RRDInsDelEnd. The
    // closure scanner must enforce that nested production before a package
    // User Names snapshot can be published.
    let marker = record(RRD_INS_DEL_BEGIN, &[]);
    let eof_offset = revision.len() - 4;
    revision.splice(eof_offset..eof_offset, marker);
    let bytes = workbook_container(names, revision);
    assert!(UserNamesPackageSnapshot::from_bytes(bytes).is_err());
}
