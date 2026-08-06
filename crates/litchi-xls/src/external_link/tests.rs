//! Regression tests for the layered external-link owner.

use super::codec::{parse_cache_row, parse_sup_book};
use super::edit::Snapshot;
use super::model::{CachedValue, NameBody};
use super::package::ExternalLinkCollector;
use super::{
    CONTINUE_RECORD_TYPE, EXTERN_NAME_RECORD_TYPE, EXTERN_SHEET_RECORD_TYPE, SUP_BOOK_RECORD_TYPE,
    XCT_RECORD_TYPE,
};

#[test]
fn malformed_supbook_xct_and_crn_are_rejected() {
    assert!(parse_sup_book(&[1, 0, 2, 4]).is_err());
    assert!(parse_cache_row(&[0, 1, 0, 0]).is_err());
    assert!(parse_cache_row(&[0, 0, 0, 0, 4, 2, 0, 0, 0, 0, 0, 0, 0]).is_err());

    let mut collector = ExternalLinkCollector::new();
    assert!(
        collector
            .feed_record(XCT_RECORD_TYPE, &[1, 0, 0, 0])
            .is_err()
    );
}

#[test]
fn xct_cardinality_and_xti_bounds_are_strict() {
    let external = [1, 0, 2, 0, 0, 1, b'A', 1, 0, 0, b'S'];
    let mut collector = ExternalLinkCollector::new();
    collector
        .feed_record(SUP_BOOK_RECORD_TYPE, &external)
        .unwrap();
    collector
        .feed_record(XCT_RECORD_TYPE, &[1, 0, 0, 0])
        .unwrap();
    assert!(
        collector
            .feed_record(EXTERN_SHEET_RECORD_TYPE, &[0, 0])
            .is_err()
    );

    let mut collector = ExternalLinkCollector::new();
    collector
        .feed_record(SUP_BOOK_RECORD_TYPE, &external)
        .unwrap();
    collector
        .feed_record(EXTERN_SHEET_RECORD_TYPE, &[1, 0, 1, 0, 0, 0, 0, 0])
        .unwrap();
    assert!(collector.finish(1).is_err());
}

#[test]
fn extern_names_are_contextual_bounded_and_continue_complete_moper_values() {
    let addin = [1, 0, 1, 0x3a];
    let addin_name = [0, 0, 0, 0, 0, 0, 1, 0, b'F', 2, 0, 0x1c, 0x17];
    let mut collector = ExternalLinkCollector::new();
    assert!(
        collector
            .feed_record(EXTERN_NAME_RECORD_TYPE, &addin_name)
            .is_err()
    );
    collector.feed_record(SUP_BOOK_RECORD_TYPE, &addin).unwrap();
    collector
        .feed_record(EXTERN_NAME_RECORD_TYPE, &addin_name)
        .unwrap();
    assert!(collector.feed_record(CONTINUE_RECORD_TYPE, &[0]).is_err());

    let dde = [0, 0, 3, 0, 0, b'A', 3, b'B'];
    let dde_name = [2, 0, 0, 0, 0, 0, 1, 0, b'X', 0, 0, 0];
    let mut collector = ExternalLinkCollector::new();
    collector.feed_record(SUP_BOOK_RECORD_TYPE, &dde).unwrap();
    collector
        .feed_record(EXTERN_NAME_RECORD_TYPE, &dde_name)
        .unwrap();
    collector
        .feed_record(CONTINUE_RECORD_TYPE, &[0; 9])
        .unwrap();
    let links = collector.finish(1).unwrap();
    let NameBody::DdeOrOle {
        matrix: Some(matrix),
        ..
    } = links.external_names()[0].body()
    else {
        panic!("expected DDE/OLE body")
    };
    assert_eq!(matrix.last_column, 0);
    assert_eq!(matrix.last_row, 0);
    assert_eq!(matrix.values, [CachedValue::Blank]);
}

fn record(record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut result = Vec::with_capacity(4 + payload.len());
    result.extend_from_slice(&record_type.to_le_bytes());
    result.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    result.extend_from_slice(payload);
    result
}

fn no_cch_unicode(output: &mut Vec<u8>, value: &str) {
    let compressed = value.chars().all(|character| u32::from(character) <= 0xFF);
    output.push(if compressed { 0 } else { 1 });
    if compressed {
        output.extend(value.chars().map(|character| character as u8));
    } else {
        for unit in value.encode_utf16() {
            output.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

fn counted_unicode(output: &mut Vec<u8>, value: &str) {
    output.extend_from_slice(&(value.encode_utf16().count() as u16).to_le_bytes());
    no_cch_unicode(output, value);
}

fn external_supbook(path: &str, sheets: &[&str]) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&(sheets.len() as u16).to_le_bytes());
    counted_unicode(&mut payload, path);
    for sheet in sheets {
        counted_unicode(&mut payload, sheet);
    }
    payload
}

fn external_name(name: &str) -> Vec<u8> {
    let mut payload = vec![0; 6];
    payload.push(name.encode_utf16().count() as u8);
    no_cch_unicode(&mut payload, name);
    payload.extend_from_slice(&2u16.to_le_bytes());
    payload.extend_from_slice(&[0x3A, 0x00]);
    payload
}

fn cache_records(valid: bool) -> (Vec<u8>, Vec<u8>) {
    let mut xct = Vec::from(if valid { 1i16 } else { -1i16 }.to_le_bytes());
    xct.extend_from_slice(&[0, 0]);
    let mut crn = vec![0, 0, 7, 0];
    crn.extend_from_slice(&[0; 9]);
    (xct, crn)
}

fn source_stream() -> Vec<u8> {
    let (xct, crn) = cache_records(true);
    let mut stream = record(0x7777, &[0xA1, 0xB2, 0xC3]);
    stream.extend_from_slice(&record(
        SUP_BOOK_RECORD_TYPE,
        &external_supbook("remote.xls", &["Inputs"]),
    ));
    stream.extend_from_slice(&record(EXTERN_NAME_RECORD_TYPE, &external_name("Sales")));
    stream.extend_from_slice(&record(XCT_RECORD_TYPE, &xct));
    stream.extend_from_slice(&record(super::CRN_RECORD_TYPE, &crn));
    stream.extend_from_slice(&record(0x8888, &[0xD4, 0xE5]));
    stream.extend_from_slice(&record(0x000A, &[]));
    stream
}

fn dde_stream() -> Vec<u8> {
    let mut supbook = vec![0, 0];
    counted_unicode(&mut supbook, "dde-server");
    let mut name = vec![0; 6];
    name.push(4);
    no_cch_unicode(&mut name, "Item");
    name.extend_from_slice(&[0, 0, 0]);
    let mut stream = record(SUP_BOOK_RECORD_TYPE, &supbook);
    stream.extend_from_slice(&record(EXTERN_NAME_RECORD_TYPE, &name));
    let mut continuation = vec![0x00];
    continuation.extend_from_slice(&[0; 8]);
    stream.extend_from_slice(&record(CONTINUE_RECORD_TYPE, &continuation));
    stream.extend_from_slice(&record(0x9999, &[0xF1, 0xF2, 0xF3]));
    stream.extend_from_slice(&record(0x000A, &[]));
    stream
}

#[test]
fn external_link_snapshot_noop_preserves_exact_stream() {
    let bytes = source_stream();
    let snapshot = Snapshot::parse(bytes.clone()).unwrap();

    assert_eq!(snapshot.finish(), bytes);
    assert_eq!(snapshot.links().supporting_books().len(), 1);
    assert_eq!(snapshot.links().external_names().len(), 1);
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().source_fingerprint(), snapshot.fingerprint());
    assert_eq!(commit.snapshot().finish(), bytes);
}

#[test]
fn metadata_edits_are_contextual_and_preserve_unknown_records() {
    let bytes = source_stream();
    let snapshot = Snapshot::parse(bytes.clone()).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .set_supporting_book_target(0, "renamed.xls")
        .unwrap()
        .set_sheet_name(0, 0, "InputData")
        .unwrap()
        .set_external_name(0, "Revenue")
        .unwrap()
        .set_cache_valid(0, 0, false)
        .unwrap();
    let commit = transaction.commit().unwrap();

    assert!(commit.changed());
    let links = commit.snapshot().links();
    let super::model::SupportingBook::ExternalWorkbook(book) = &links.supporting_books()[0] else {
        panic!("expected external workbook")
    };
    assert_eq!(book.encoded_virtual_path(), "renamed.xls");
    assert_eq!(book.sheets()[0].name(), "InputData");
    assert!(!book.sheets()[0].cache_valid());
    assert!(matches!(
        links.external_names()[0].body(),
        NameBody::ExternalDefinedName { name, .. } if name == "Revenue"
    ));
    assert!(
        commit
            .snapshot()
            .bytes()
            .windows(3)
            .any(|window| window == [0xA1, 0xB2, 0xC3])
    );
    assert!(
        commit
            .snapshot()
            .bytes()
            .windows(2)
            .any(|window| window == [0xD4, 0xE5])
    );

    let applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(applied, *commit.snapshot());
    let reverted = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(reverted.bytes(), snapshot.bytes());
}

#[test]
fn external_name_edit_preserves_owned_continue_payload() {
    let bytes = dde_stream();
    let snapshot = Snapshot::parse(bytes.clone()).unwrap();
    let mut transaction = snapshot.edit();
    transaction.set_external_name(0, "RenamedItem").unwrap();
    let commit = transaction.commit().unwrap();

    assert!(matches!(
        commit.snapshot().links().external_names()[0].body(),
        NameBody::DdeOrOle {
            name,
            matrix: Some(matrix),
            ..
        } if name == "RenamedItem" && matrix.values == [CachedValue::Blank]
    ));
    let continuation = record(CONTINUE_RECORD_TYPE, &[0x00, 0, 0, 0, 0, 0, 0, 0, 0]);
    assert!(
        commit
            .snapshot()
            .bytes()
            .windows(continuation.len())
            .any(|window| window == continuation)
    );
}

#[test]
fn invalid_edits_are_failure_atomic_and_stale_patches_are_rejected() {
    let bytes = source_stream();
    let snapshot = Snapshot::parse(bytes).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();

    assert!(
        transaction
            .set_supporting_book_target(3, "other.xls")
            .is_err()
    );
    assert!(transaction.set_supporting_book_target(0, "\0").is_err());
    assert!(transaction.set_sheet_name(0, 5, "Missing").is_err());
    assert!(transaction.set_external_name(0, &"x".repeat(256)).is_err());
    assert!(transaction.set_cache_valid(0, 5, true).is_err());
    assert!(transaction.set_cache_valid(0, 0, true).is_ok());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);

    let mut edit = snapshot.edit();
    edit.set_external_name(0, "Changed").unwrap();
    let commit = edit.commit().unwrap();
    let mut stale_bytes = snapshot.bytes().to_vec();
    let marker = stale_bytes
        .windows(3)
        .position(|window| window == [0xA1, 0xB2, 0xC3])
        .unwrap();
    stale_bytes[marker] ^= 1;
    let stale = Snapshot::parse(stale_bytes).unwrap();
    let stale_before = stale.finish();
    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(stale.finish(), stale_before);
}

#[test]
fn malformed_source_framing_and_link_ownership_are_rejected() {
    assert!(Snapshot::parse([0xAE, 0x01, 0x01]).is_err());

    let mut stream = record(
        SUP_BOOK_RECORD_TYPE,
        &external_supbook("remote.xls", &["Inputs"]),
    );
    stream.extend_from_slice(&record(0x4444, &[1]));
    stream.extend_from_slice(&record(EXTERN_NAME_RECORD_TYPE, &external_name("Late")));
    assert!(Snapshot::parse(stream).is_err());
}
