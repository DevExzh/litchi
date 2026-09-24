#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
//! Directory state bits and FILETIMEs under the reuse sector-layout policy.
//!
//! A reused layout emits the adopted source's directory image. Directory
//! metadata supplied through the writer's metadata setters is authoritative
//! for every entry, including the all-zero default of an entry the caller did
//! not set, so the reused image may be emitted only when it already holds the
//! model's values; otherwise the writer declines to the from-scratch
//! serializer, which publishes the model exactly. A caller that supplies no
//! metadata keeps the source's values under a reused layout.

use litchi_cfb::{OleFile, OleWriter, SectorLayoutFallback};
use std::io::Cursor;
use std::sync::Arc;

const STATE: u32 = 0x1122_3344;
const CREATED: u64 = 0x01D0_0000_0000_0001;
const MODIFIED: u64 = 0x01D0_0000_0000_0002;

fn serialize(writer: &mut OleWriter) -> Vec<u8> {
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    out.into_inner()
}

fn source() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_storage(&["Storage"]).unwrap();
    writer
        .set_storage_metadata(&["Storage"], STATE, CREATED, MODIFIED)
        .unwrap();
    writer
        .create_stream(&["Storage", "Payload"], &[7u8; 5000])
        .unwrap();
    writer.create_stream(&["Small"], &[3u8; 100]).unwrap();
    serialize(&mut writer)
}

/// A writer over the source's directory shape with one edited payload.
fn edited(source: &[u8]) -> OleWriter {
    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(source).unwrap());
    writer.create_storage(&["Storage"]).unwrap();
    writer
        .create_stream(&["Storage", "Payload"], &[9u8; 5000])
        .unwrap();
    writer.create_stream(&["Small"], &[3u8; 100]).unwrap();
    writer
}

fn storage_metadata(bytes: &[u8]) -> (u32, u64, u64) {
    let ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let entry = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "Storage")
        .unwrap();
    (entry.state_bits, entry.creation_time, entry.modified_time)
}

fn payload(bytes: &[u8]) -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    ole.open_stream(&["Storage", "Payload"]).unwrap()
}

#[test]
fn reuse_without_supplied_metadata_keeps_the_source_values() {
    let source = source();
    let mut writer = edited(&source);
    let output = serialize(&mut writer);

    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert_eq!(storage_metadata(&output), (STATE, CREATED, MODIFIED));
    assert_eq!(payload(&output), vec![9u8; 5000]);
}

#[test]
fn reuse_with_matching_supplied_metadata_keeps_the_layout() {
    let source = source();
    let mut writer = edited(&source);
    writer.set_root_state_bits(0);
    writer
        .set_storage_metadata(&["Storage"], STATE, CREATED, MODIFIED)
        .unwrap();
    let output = serialize(&mut writer);

    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert_eq!(storage_metadata(&output), (STATE, CREATED, MODIFIED));
    assert_eq!(payload(&output), vec![9u8; 5000]);
}

#[test]
fn changed_supplied_metadata_declines_to_the_from_scratch_writer() {
    let source = source();
    let mut writer = edited(&source);
    writer
        .set_storage_metadata(&["Storage"], 7, CREATED + 10, MODIFIED + 10)
        .unwrap();
    let output = serialize(&mut writer);

    let report = writer.last_sector_layout().unwrap();
    assert!(!report.reused_source_layout(), "{report:?}");
    assert_eq!(
        report.fallback(),
        Some(SectorLayoutFallback::DirectoryMetadataChanged)
    );
    assert_eq!(storage_metadata(&output), (7, CREATED + 10, MODIFIED + 10));
    assert_eq!(payload(&output), vec![9u8; 5000]);
}

#[test]
fn supplied_metadata_is_authoritative_for_entries_left_unset() {
    // Supplying only root metadata makes the whole model authoritative, so
    // the storage that the caller left unset publishes the zero default
    // instead of the source's values.
    let source = source();
    let mut writer = edited(&source);
    writer.set_root_state_bits(0);
    let output = serialize(&mut writer);

    let report = writer.last_sector_layout().unwrap();
    assert_eq!(
        report.fallback(),
        Some(SectorLayoutFallback::DirectoryMetadataChanged)
    );
    assert_eq!(storage_metadata(&output), (0, 0, 0));
}

#[test]
fn changed_root_metadata_declines_to_the_from_scratch_writer() {
    let source = source();
    let mut writer = edited(&source);
    writer
        .set_storage_metadata(&["Storage"], STATE, CREATED, MODIFIED)
        .unwrap();
    writer.set_root_modified_time(MODIFIED);
    let output = serialize(&mut writer);

    let report = writer.last_sector_layout().unwrap();
    assert_eq!(
        report.fallback(),
        Some(SectorLayoutFallback::DirectoryMetadataChanged)
    );
    let ole = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(ole.root_entry().unwrap().modified_time, MODIFIED);
}

#[test]
fn shared_metadata_registration_matches_the_copying_form() {
    let bytes = vec![0x5Au8; 6000];
    let mut copied = OleWriter::new();
    copied
        .create_stream_with_metadata(&["Stream"], &bytes, 1, 2, 3)
        .unwrap();
    let mut shared = OleWriter::new();
    shared
        .create_stream_shared_with_metadata(&["Stream"], Arc::from(bytes.as_slice()), 1, 2, 3)
        .unwrap();
    let copied = serialize(&mut copied);
    assert_eq!(serialize(&mut shared), copied);

    let ole = OleFile::open(Cursor::new(copied)).unwrap();
    let entry = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "Stream")
        .unwrap();
    assert_eq!(
        (entry.state_bits, entry.creation_time, entry.modified_time),
        (1, 2, 3)
    );
}
