#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "validation fixtures intentionally panic on construction failure"
)]

use litchi_cfb::OleWriter;
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_ppt::{
    DocumentAtom, DocumentDimensions, PptValidationError, PptValidationLimits, Ratio, RecordLimits,
    SlideSizeType, validate_source, validate_source_with_limits,
};
use std::{
    io::{self, Cursor},
    sync::{
        Arc,
        atomic::{AtomicU64, AtomicUsize, Ordering},
    },
};

fn record(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&(u32::try_from(payload.len()).unwrap()).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn document_atom() -> Vec<u8> {
    DocumentAtom::new(
        DocumentDimensions::new(5760, 4320).unwrap(),
        DocumentDimensions::new(5760, 4320).unwrap(),
        Ratio::new(1, 1).unwrap(),
        0,
        0,
        1,
        SlideSizeType::Screen,
        false,
        false,
        false,
        true,
    )
    .unwrap()
    .to_record_bytes()
    .unwrap()
    .to_vec()
}

fn current_user(offset: u32, encrypted: bool) -> Vec<u8> {
    current_user_with_names(offset, encrypted, &[], &[])
}

fn current_user_with_names(
    offset: u32,
    encrypted: bool,
    ansi_name: &[u8],
    unicode_name: &[u16],
) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&0x0FF6u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(&20u32.to_le_bytes());
    bytes.extend_from_slice(
        &(if encrypted {
            0xF3D1_C4DFu32
        } else {
            0xE391_C05Fu32
        })
        .to_le_bytes(),
    );
    bytes.extend_from_slice(&offset.to_le_bytes());
    bytes.extend_from_slice(&u16::try_from(ansi_name.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(&0x03F4u16.to_le_bytes());
    bytes.extend_from_slice(&[3, 0, 0, 0]);
    bytes.extend_from_slice(ansi_name);
    bytes.extend_from_slice(&8u32.to_le_bytes());
    for code_unit in unicode_name {
        bytes.extend_from_slice(&code_unit.to_le_bytes());
    }
    let record_len = u32::try_from(bytes.len() - 8).unwrap();
    bytes[4..8].copy_from_slice(&record_len.to_le_bytes());
    bytes
}

const USER_EDIT_RECORD_SIZE: usize = 8 + 28;

fn persist_directory_atom() -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&((1_u32 << 20) | 1).to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    record(0, 0, 6002, &payload)
}

fn user_edit_payload(
    offset_last_edit: u32,
    offset_persist_directory: u32,
    doc_persist_id_ref: u32,
) -> [u8; 28] {
    let mut payload = [0u8; 28];
    payload[6] = 0;
    payload[7] = 3;
    payload[8..12].copy_from_slice(&offset_last_edit.to_le_bytes());
    payload[12..16].copy_from_slice(&offset_persist_directory.to_le_bytes());
    payload[16..20].copy_from_slice(&doc_persist_id_ref.to_le_bytes());
    payload
}

fn user_edit_atom(offset_persist_directory: u32) -> Vec<u8> {
    record(
        0,
        0,
        4085,
        &user_edit_payload(0, offset_persist_directory, 1),
    )
}

fn valid_stream(_encrypted: bool, extra_records: &[Vec<u8>]) -> Vec<u8> {
    let mut document = record(0x0F, 0, 1000, &document_atom());
    document.extend(extra_records.iter().flatten().copied());
    let persist_offset = u32::try_from(document.len()).unwrap();
    document.extend(persist_directory_atom());
    document.extend(user_edit_atom(persist_offset));
    document
}

fn package_with_streams(
    dual_storage: bool,
    document: &[u8],
    current: &[u8],
    extra_streams: &[(&[&str], &[u8])],
) -> Vec<u8> {
    let mut writer = OleWriter::new();
    let (document_path, current_path) = if dual_storage {
        writer.create_storage(&["PP97_DUALSTORAGE"]).unwrap();
        (
            &["PP97_DUALSTORAGE", "PowerPoint Document"][..],
            &["PP97_DUALSTORAGE", "Current User"][..],
        )
    } else {
        (&["PowerPoint Document"][..], &["Current User"][..])
    };
    writer.create_stream(document_path, document).unwrap();
    writer.create_stream(current_path, current).unwrap();
    for (path, bytes) in extra_streams {
        writer.create_stream(path, bytes).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn package(
    dual_storage: bool,
    encrypted: bool,
    extra_records: &[Vec<u8>],
    extra_streams: &[(&[&str], &[u8])],
) -> Vec<u8> {
    let document = valid_stream(encrypted, extra_records);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let current = current_user(edit_offset, encrypted);
    package_with_streams(dual_storage, &document, &current, extra_streams)
}

fn check<'a>(
    report: &'a litchi_core::ValidateReport,
    id: &str,
) -> &'a litchi_core::ValidationCheck {
    report
        .checks()
        .iter()
        .find(|check| check.id().as_str() == id)
        .unwrap()
}

fn fat_stream_range_and_size(bytes: &[u8], name: &str) -> ((u64, u64), u64) {
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    let entry = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == name)
        .unwrap();
    assert!(!entry.is_minifat, "fixture stream must use FAT sectors");
    let sector_size = 1usize << usize::from(u16::from_le_bytes([bytes[30], bytes[31]]));
    let start = (u64::from(entry.start_sector) + 1) * sector_size as u64;
    let sectors = entry.size.div_ceil(sector_size as u64);
    let end = start + sectors * sector_size as u64;
    ((start, end), entry.size)
}

#[test]
fn valid_root_and_dual_storage_hierarchies_are_complete() {
    for dual in [false, true] {
        let bytes = package(dual, false, &[], &[]);
        let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert!(report.is_complete(), "{report:?}");
        assert!(!report.has_errors(), "{report:?}");
        assert!(check(&report, "ppt.record.parse").status().is_complete());
    }
}

#[test]
fn pictures_must_remain_root_level_with_dual_storage() {
    let bytes = package(
        true,
        false,
        &[],
        &[(&["PP97_DUALSTORAGE", "Pictures"][..], b"picture")],
    );
    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(!check(&report, "ppt.pictures.stream").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.pictures.stream_noncanonical")
    );
}

#[test]
fn pictures_under_another_storage_are_noncanonical() {
    let bytes = package(
        false,
        false,
        &[],
        &[(&["OtherStorage", "Pictures"][..], b"picture")],
    );
    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(!check(&report, "ppt.pictures.stream").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.pictures.stream_noncanonical")
    );
}

#[test]
fn cross_storage_native_streams_are_rejected_before_payload_reads() {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let current = current_user(edit_offset, false);
    let mut writer = OleWriter::new();
    writer.create_storage(&["PP97_DUALSTORAGE"]).unwrap();
    writer
        .create_stream(&["PowerPoint Document"], &document)
        .unwrap();
    writer
        .create_stream(&["PP97_DUALSTORAGE", "Current User"], &current)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let (document_range, _) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let source = Arc::new(RejectRangesSource::new(bytes, [document_range]));
    let report = validate_source(source).unwrap();
    assert!(
        !check(&report, "ppt.storage.hierarchy")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.document.stream").status().is_complete());
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.storage.hierarchy_invalid")
    );
}

#[test]
fn directory_depth_limit_stops_native_stream_checks() {
    let document = valid_stream(false, &[]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let current = current_user(edit_offset, false);
    let bytes = package_with_streams(
        false,
        &document,
        &current,
        &[(&["Nested", "TooDeep"][..], b"marker")],
    );
    let defaults = PptValidationLimits::default();
    let limits = PptValidationLimits::new(
        defaults.max_input_bytes(),
        defaults.max_document_stream_bytes(),
        defaults.max_current_user_stream_bytes(),
        defaults.max_pictures_stream_bytes(),
        defaults.max_aggregate_stream_bytes(),
        defaults.max_directory_entries(),
        1,
        defaults.record(),
        defaults.report(),
    );
    let report = validate_source_with_limits(Arc::new(OwnedSource::new(bytes)), limits).unwrap();
    assert!(
        !check(&report, "ppt.storage.hierarchy")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.document.stream").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.storage.depth_limit")
    );
}

#[test]
fn validation_is_non_mutating_and_uses_positional_reads() {
    let bytes = package(false, false, &[], &[]);
    let original = bytes.clone();
    let source = Arc::new(CountingSource::new(bytes));
    let report = validate_source(source.clone()).unwrap();
    assert!(!report.has_errors(), "{report:?}");
    assert!(source.reads.load(Ordering::Relaxed) > 0);
    assert_eq!(source.bytes.as_slice(), original.as_slice());
}

#[test]
fn hostile_topology_and_short_source_are_reported_without_panic() {
    let document = valid_stream(false, &[]);
    let edit_offset = u32::try_from(document.len() - 8).unwrap();
    let current = current_user(edit_offset, false);
    let mut writer = OleWriter::new();
    writer.create_storage(&["PowerPoint Document"]).unwrap();
    writer.create_stream(&["Current User"], &current).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let report = validate_source(Arc::new(OwnedSource::new(output.into_inner()))).unwrap();
    assert!(report.has_errors());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.document.stream_ambiguous")
    );

    let short = validate_source(Arc::new(OwnedSource::new(vec![0u8; 64]))).unwrap();
    assert!(short.has_errors());
    assert!(
        short
            .issues()
            .iter()
            .any(|issue| issue.code().starts_with("ppt.cfb."))
    );
}

#[test]
fn security_macro_protection_and_external_markers_are_presence_only() {
    let extra = vec![
        record(0x0F, 0, 1023, &[]),
        record(0x00, 3, 4026, &[]),
        record(0x00, 0, 3009, &[]),
    ];
    let bytes = package(
        false,
        true,
        &extra,
        &[
            (&["DigitalSignature"][..], b"signature"),
            (&["EncryptionInfo"][..], b"encryption"),
            (&["\u{0009}DRMContent"][..], b"drm"),
            (&["VBA"][..], b"macro"),
        ],
    );
    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.encryption.password_to_open_present")
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.signature.infrastructure_present")
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.macro.storage_present")
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
    assert!(check(&report, "ppt.macro.presence").status().is_complete());
    assert!(
        !report
            .issues()
            .iter()
            .any(|issue| issue.repair().repair_id().is_some())
    );
}

#[test]
fn encrypted_validation_does_not_read_powerpoint_document_payload() {
    let bytes = package(
        false,
        true,
        &[record(0, 0, 0x2222, &vec![0xA5; 8 * 1024])],
        &[],
    );
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes.clone())).unwrap();
    let document = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "PowerPoint Document")
        .unwrap();
    assert!(
        !document.is_minifat,
        "fixture must keep document in FAT sectors"
    );
    let sector_size = 1usize << usize::from(u16::from_le_bytes([bytes[30], bytes[31]]));
    let start = (u64::from(document.start_sector) + 1) * sector_size as u64;
    let sectors = document.size.div_ceil(sector_size as u64);
    let end = start + sectors * sector_size as u64;
    let source = Arc::new(RejectRangesSource::new(bytes, [(start, end)]));
    let report = validate_source(source).unwrap();
    assert!(!report.has_errors(), "{report:?}");
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.encryption.password_to_open_present")
    );
}

#[test]
fn malformed_current_user_never_reads_powerpoint_document_payload() {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let bytes = package_with_streams(false, &document, &[0u8; 28], &[]);
    let (document_range, _) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let source = Arc::new(RejectRangesSource::new(bytes, [document_range]));
    let report = validate_source(source).unwrap();
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.current_user.invalid")
    );
}

fn assert_current_user_rejected_without_document_read(current: &[u8]) {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let bytes = package_with_streams(false, &document, current, &[]);
    let (document_range, _) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let source = Arc::new(RejectRangesSource::new(bytes, [document_range]));
    let report = validate_source(source).unwrap();
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.current_user.invalid")
    );
}

#[test]
fn current_user_length_and_trailing_layout_are_fail_closed() {
    let document = valid_stream(false, &[]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let base = current_user(edit_offset, false);

    let mut zero_length = base.clone();
    zero_length[4..8].copy_from_slice(&0_u32.to_le_bytes());
    assert_current_user_rejected_without_document_read(&zero_length);

    let mut huge_length = base.clone();
    huge_length[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
    assert_current_user_rejected_without_document_read(&huge_length);

    let mut truncated = base.clone();
    truncated.pop();
    assert_current_user_rejected_without_document_read(&truncated);

    let mut unexpected_trailing = base;
    unexpected_trailing[4..8].copy_from_slice(&25_u32.to_le_bytes());
    unexpected_trailing.push(0xA5);
    assert_current_user_rejected_without_document_read(&unexpected_trailing);
}

#[test]
fn current_user_must_have_supported_versions_and_printable_names() {
    let document = valid_stream(false, &[]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let base = current_user(edit_offset, false);

    let mut bad_document_version = base.clone();
    bad_document_version[22..24].copy_from_slice(&0x03F5_u16.to_le_bytes());
    assert_current_user_rejected_without_document_read(&bad_document_version);

    let mut bad_major_version = base.clone();
    bad_major_version[24..26].copy_from_slice(&4_u16.to_le_bytes());
    assert_current_user_rejected_without_document_read(&bad_major_version);

    let mut bad_minor_version = base.clone();
    bad_minor_version[25] = 1;
    assert_current_user_rejected_without_document_read(&bad_minor_version);

    let mut bad_release_version = base.clone();
    bad_release_version[28..32].copy_from_slice(&7_u32.to_le_bytes());
    assert_current_user_rejected_without_document_read(&bad_release_version);

    let bad_ansi = current_user_with_names(edit_offset, false, &[0], &[]);
    assert_current_user_rejected_without_document_read(&bad_ansi);

    let bad_unicode = current_user_with_names(edit_offset, false, b"A", &[0xD800]);
    assert_current_user_rejected_without_document_read(&bad_unicode);

    let mut nonzero_unused = base.clone();
    nonzero_unused[26..28].copy_from_slice(&0xBEEF_u16.to_le_bytes());
    let bytes = package_with_streams(false, &document, &nonzero_unused, &[]);
    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(
        check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
}

#[test]
fn oversized_current_user_never_reads_powerpoint_document_payload() {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let mut current = current_user(edit_offset, false);
    current.extend_from_slice(&[0u8; 4096]);
    let bytes = package_with_streams(false, &document, &current, &[]);
    let (document_range, _) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let defaults = PptValidationLimits::default();
    let limits = PptValidationLimits::new(
        defaults.max_input_bytes(),
        defaults.max_document_stream_bytes(),
        32,
        defaults.max_pictures_stream_bytes(),
        defaults.max_aggregate_stream_bytes(),
        defaults.max_directory_entries(),
        defaults.max_directory_depth(),
        defaults.record(),
        defaults.report(),
    );
    let source = Arc::new(RejectRangesSource::new(bytes, [document_range]));
    let report = validate_source_with_limits(source, limits).unwrap();
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
}

#[test]
fn out_of_bounds_current_edit_never_reads_powerpoint_document_payload() {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let current = current_user(u32::try_from(document.len()).unwrap(), false);
    let bytes = package_with_streams(false, &document, &current, &[]);
    let (document_range, _) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let source = Arc::new(RejectRangesSource::new(bytes, [document_range]));

    let report = validate_source(source).unwrap();
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.current_user.edit_offset_out_of_bounds"),
        "{report:?}"
    );
}

#[test]
fn current_user_parser_resource_limit_blocks_without_document_read() {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let mut current = current_user(edit_offset, false);
    current.extend_from_slice(&[0u8; 4096]);
    let bytes = package_with_streams(false, &document, &current, &[]);
    let (document_range, _) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let defaults = PptValidationLimits::default();
    let record_limits = RecordLimits {
        max_input_bytes: 32,
        ..defaults.record()
    };
    let limits = PptValidationLimits::new(
        defaults.max_input_bytes(),
        defaults.max_document_stream_bytes(),
        defaults.max_current_user_stream_bytes(),
        defaults.max_pictures_stream_bytes(),
        defaults.max_aggregate_stream_bytes(),
        defaults.max_directory_entries(),
        defaults.max_directory_depth(),
        record_limits,
        defaults.report(),
    );
    let source = Arc::new(RejectRangesSource::new(bytes, [document_range]));
    let report = validate_source_with_limits(source, limits).unwrap();
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.current_user.limit")
    );
}

#[test]
fn aggregate_stream_limit_refuses_native_payload_reads() {
    let document = valid_stream(false, &[record(0, 0, 0x2222, &[0xA5; 8 * 1024])]);
    let edit_offset = u32::try_from(document.len() - USER_EDIT_RECORD_SIZE).unwrap();
    let mut current = current_user(edit_offset, false);
    current.extend_from_slice(&[0u8; 4096]);
    let bytes = package_with_streams(false, &document, &current, &[]);
    let (document_range, document_size) = fat_stream_range_and_size(&bytes, "PowerPoint Document");
    let (current_range, current_size) = fat_stream_range_and_size(&bytes, "Current User");
    let aggregate = document_size.checked_add(current_size).unwrap();
    let defaults = PptValidationLimits::default();
    let limits = PptValidationLimits::new(
        defaults.max_input_bytes(),
        defaults.max_document_stream_bytes(),
        defaults.max_current_user_stream_bytes(),
        defaults.max_pictures_stream_bytes(),
        aggregate - 1,
        defaults.max_directory_entries(),
        defaults.max_directory_depth(),
        defaults.record(),
        defaults.report(),
    );
    let source = Arc::new(RejectRangesSource::new(
        bytes,
        [document_range, current_range],
    ));
    let report = validate_source_with_limits(source, limits).unwrap();
    assert!(!check(&report, "ppt.stream.budget").status().is_complete());
    assert!(!check(&report, "ppt.document.stream").status().is_complete());
    assert!(
        !check(&report, "ppt.current_user.stream")
            .status()
            .is_complete()
    );
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
}

#[test]
fn user_edit_bootstrap_rejects_empty_wrong_version_and_null_persist_records() {
    let mut document_prefix = record(0x0F, 0, 1000, &document_atom());
    let persist_offset = u32::try_from(document_prefix.len()).unwrap();
    document_prefix.extend(persist_directory_atom());
    let valid_payload = user_edit_payload(0, persist_offset, 1);
    let null_persist_payload = user_edit_payload(0, persist_offset, 0);
    let edits = [
        record(0, 0, 4085, &[]),
        record(1, 0, 4085, &valid_payload),
        record(0, 0, 4085, &null_persist_payload),
    ];
    for edit in edits {
        let mut document = document_prefix.clone();
        let edit_offset = u32::try_from(document.len()).unwrap();
        document.extend(edit);
        let current = current_user(edit_offset, false);
        let bytes = package_with_streams(false, &document, &current, &[]);
        let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == "ppt.current_user.edit_target_invalid"),
            "{report:?}"
        );
    }
}

#[test]
fn user_edit_bootstrap_rejects_invalid_versions_ids_and_offsets() {
    let mut document_prefix = record(0x0F, 0, 1000, &document_atom());
    let persist_offset = u32::try_from(document_prefix.len()).unwrap();
    document_prefix.extend(persist_directory_atom());
    let edit_offset = u32::try_from(document_prefix.len()).unwrap();
    let mut payloads = Vec::new();

    let mut bad_minor = user_edit_payload(0, persist_offset, 1);
    bad_minor[6] = 1;
    payloads.push(bad_minor);
    let mut bad_major = user_edit_payload(0, persist_offset, 1);
    bad_major[7] = 2;
    payloads.push(bad_major);
    payloads.push(user_edit_payload(0, persist_offset, 2));
    payloads.push(user_edit_payload(0, 0, 1));
    payloads.push(user_edit_payload(0, edit_offset, 1));
    payloads.push(user_edit_payload(0, edit_offset + 1, 1));
    payloads.push(user_edit_payload(edit_offset, persist_offset, 1));
    payloads.push(user_edit_payload(edit_offset + 1, persist_offset, 1));

    for payload in payloads {
        let mut document = document_prefix.clone();
        document.extend(record(0, 0, 4085, &payload));
        let current = current_user(edit_offset, false);
        let bytes = package_with_streams(false, &document, &current, &[]);
        let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == "ppt.current_user.edit_target_invalid"),
            "{report:?}"
        );
    }
}

#[test]
fn nested_document_container_does_not_satisfy_top_level_owner() {
    let nested = record(0x0F, 0, 1000, &document_atom());
    let mut document = record(0x0F, 0, 1006, &nested);
    let persist_offset = u32::try_from(document.len()).unwrap();
    document.extend(persist_directory_atom());
    let edit_offset = u32::try_from(document.len()).unwrap();
    document.extend(user_edit_atom(persist_offset));
    let current = current_user(edit_offset, false);
    let bytes = package_with_streams(false, &document, &current, &[]);
    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.record.document_missing"),
        "{report:?}"
    );
}

#[test]
fn top_level_document_container_header_is_checked() {
    let mut document = record(0x0E, 0, 1000, &document_atom());
    let persist_offset = u32::try_from(document.len()).unwrap();
    document.extend(persist_directory_atom());
    let edit_offset = u32::try_from(document.len()).unwrap();
    document.extend(user_edit_atom(persist_offset));
    let current = current_user(edit_offset, false);
    let bytes = package_with_streams(false, &document, &current, &[]);
    let report = validate_source(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.record.document_missing")
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "ppt.record.document_header_invalid")
    );
}

#[test]
fn stream_and_record_limits_stop_work_before_unbounded_materialization() {
    let bytes = package(false, false, &[], &[]);
    let defaults = PptValidationLimits::default();
    let too_small = PptValidationLimits::new(
        defaults.max_input_bytes(),
        1,
        defaults.max_current_user_stream_bytes(),
        defaults.max_pictures_stream_bytes(),
        defaults.max_aggregate_stream_bytes(),
        defaults.max_directory_entries(),
        defaults.max_directory_depth(),
        defaults.record(),
        defaults.report(),
    );
    let report =
        validate_source_with_limits(Arc::new(OwnedSource::new(bytes.clone())), too_small).unwrap();
    assert!(!check(&report, "ppt.document.stream").status().is_complete());

    let limited_record = RecordLimits {
        max_records: 1,
        ..defaults.record()
    };
    let limits = PptValidationLimits::new(
        defaults.max_input_bytes(),
        defaults.max_document_stream_bytes(),
        defaults.max_current_user_stream_bytes(),
        defaults.max_pictures_stream_bytes(),
        defaults.max_aggregate_stream_bytes(),
        defaults.max_directory_entries(),
        defaults.max_directory_depth(),
        limited_record,
        defaults.report(),
    );
    let report = validate_source_with_limits(Arc::new(OwnedSource::new(bytes)), limits).unwrap();
    assert!(!check(&report, "ppt.record.parse").status().is_complete());
}

#[test]
fn source_version_changes_are_errors() {
    let bytes = package(false, false, &[], &[]);
    let source = Arc::new(ChangingSource::new(bytes));
    let error = validate_source(source).unwrap_err();
    assert!(matches!(
        error,
        PptValidationError::Ingress(litchi_cfb::OleError::SourceChanged { .. })
    ));
}

struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: AtomicUsize,
    version: SourceVersion,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            reads: AtomicUsize::new(0),
            version: SourceVersion::new(0xCAFE, 0),
        }
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
        let start = usize::try_from(offset).unwrap();
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

struct ChangingSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
}

struct RejectRangesSource {
    bytes: Arc<Vec<u8>>,
    ranges: Vec<(u64, u64)>,
}

impl RejectRangesSource {
    fn new<const N: usize>(bytes: Vec<u8>, ranges: [(u64, u64); N]) -> Self {
        Self {
            bytes: Arc::new(bytes),
            ranges: ranges.into_iter().collect(),
        }
    }
}

impl ReadAt for RejectRangesSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let end = offset.saturating_add(output.len() as u64);
        if self
            .ranges
            .iter()
            .any(|(start, stop)| offset < *stop && end > *start)
        {
            return Err(io::Error::other("PowerPoint Document payload was read"));
        }
        let start = usize::try_from(offset).unwrap();
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0xFACE, 0))
    }
}

impl ChangingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
        }
    }
}

impl ReadAt for ChangingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.revision.fetch_add(1, Ordering::Relaxed);
        let start = usize::try_from(offset).unwrap();
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0xBEEF,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}
