#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The fixtures below are small, deterministic, and bounded."
)]

//! Differential coverage for the owned-source direct-read path.
//!
//! `IndexedArchive<Arc<Vec<u8>>>::read_owned` has the same framing, decoder,
//! verifier, output reservation, and error contract as `read`, but deliberately
//! bypasses the generic `MemberSpan` copy. These tests compare the two routes
//! over stored and Deflate members, a descriptor, ZIP64 metadata, payload
//! corruption, a truncated compressed stream, and index-time resource limits.

use soapberry_zip::office::{ArchiveLimits, IndexedArchive, StreamingArchiveWriter};
use soapberry_zip::{Error, ErrorKind, LimitResource, ZipArchive, ZipOperationAccounting};
use std::sync::Arc;

const ZIP64_BYTES: &[u8] = include_bytes!("../assets/zip64.zip");

fn open(bytes: Vec<u8>) -> IndexedArchive<Arc<Vec<u8>>> {
    let source = Arc::new(bytes);
    let length = source.len() as u64;
    IndexedArchive::from_reader(source, length).expect("fixture archive opens")
}

fn open_with_limits(
    bytes: Vec<u8>,
    limits: ArchiveLimits,
) -> Result<IndexedArchive<Arc<Vec<u8>>>, Error> {
    let source = Arc::new(bytes);
    let length = source.len() as u64;
    IndexedArchive::from_reader_with_limits(source, length, limits)
}

fn descriptor_archive() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("stored.bin", b"stored payload")
        .expect("stored fixture");
    writer
        .write_deflated("deflated.bin", b"deflated payload with a data descriptor")
        .expect("deflated fixture");
    writer.finish_to_bytes().expect("fixture archive finishes")
}

fn canonical_deflate_archive() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated_sized(
            "truncated.bin",
            b"a Deflate payload whose compressed stream can be shortened",
        )
        .expect("canonical Deflate fixture");
    writer.finish_to_bytes().expect("fixture archive finishes")
}

fn entry_offsets(bytes: &[u8], name: &str) -> (usize, usize, usize, usize) {
    let archive = ZipArchive::from_slice(bytes).expect("fixture archive parses");
    let mut entries = archive.entries();
    while let Some(record) = entries.next_entry().expect("fixture entry parses") {
        if record.file_path().as_ref() == name.as_bytes() {
            let wayfinder = record.wayfinder();
            let entry = archive
                .get_entry(wayfinder)
                .expect("fixture local record parses");
            let (payload_start, payload_end) = entry.compressed_data_range();
            let local = record.local_header_offset() as usize;
            let central = find_central_header(bytes, name);
            return (local, central, payload_start as usize, payload_end as usize);
        }
    }
    panic!("fixture member {name:?} is absent")
}

fn find_central_header(bytes: &[u8], name: &str) -> usize {
    let name = name.as_bytes();
    bytes
        .windows(4)
        .enumerate()
        .filter(|(_, signature)| *signature == [0x50, 0x4b, 0x01, 0x02])
        .find_map(|(offset, _)| {
            let name_len =
                u16::from_le_bytes(bytes[offset + 28..offset + 30].try_into().ok()?) as usize;
            let name_end = offset.checked_add(46)?.checked_add(name_len)?;
            (bytes.get(offset + 46..name_end)? == name).then_some(offset)
        })
        .unwrap_or_else(|| panic!("fixture central member {name:?} is absent"))
}

fn family(error: &Error) -> &'static str {
    match error.kind() {
        ErrorKind::InvalidChecksum { .. } => "checksum",
        ErrorKind::InvalidSize { .. } => "size",
        ErrorKind::LimitExceeded { .. } => "limit",
        ErrorKind::Eof => "eof",
        ErrorKind::IO(source) | ErrorKind::Io(source) => source
            .get_ref()
            .and_then(|nested| nested.downcast_ref::<Error>())
            .map_or("io", family),
        ErrorKind::InvalidInput { .. } => "invalid-input",
        ErrorKind::InvalidSignature { .. } => "signature",
        ErrorKind::FileNotFound(_) => "not-found",
        _ => "other",
    }
}

fn assert_same_failure(materialized: &Error, owned: &Error) {
    assert_eq!(
        family(materialized),
        family(owned),
        "materialized={materialized:?}; owned={owned:?}"
    );
    assert_eq!(
        format!("{materialized:?}"),
        format!("{owned:?}"),
        "Debug diagnostics differ"
    );
    assert_eq!(
        materialized.to_string(),
        owned.to_string(),
        "Display diagnostics differ"
    );
}

fn assert_limit(error: &Error, actual: u64, maximum: u64) {
    match error.kind() {
        ErrorKind::LimitExceeded {
            resource,
            actual: found_actual,
            maximum: found_maximum,
        } => {
            assert_eq!(*resource, LimitResource::EntrySize);
            assert_eq!((*found_actual, *found_maximum), (actual, maximum));
        },
        other => panic!("expected an entry-size limit refusal, got {other:?}"),
    }
}

#[test]
fn owned_direct_reads_match_member_span_for_store_deflate_and_descriptor() {
    let bytes = descriptor_archive();
    let materialized = open(bytes.clone());
    let owned = open(bytes);

    for name in ["stored.bin", "deflated.bin"] {
        let mut materialized_accounting = ZipOperationAccounting::default();
        let mut owned_accounting = ZipOperationAccounting::default();
        let expected = materialized
            .read_with_accounting(name, &mut materialized_accounting)
            .expect("member-span read succeeds");
        let actual = owned
            .read_owned_with_accounting(name, &mut owned_accounting)
            .expect("direct owned read succeeds");
        assert_eq!(actual, expected, "payload mismatch for {name}");
        assert_eq!(
            owned_accounting, materialized_accounting,
            "accounting drift for {name}"
        );
    }

    let descriptor_bytes = descriptor_archive();
    let archive = ZipArchive::from_slice(&descriptor_bytes).expect("fixture parses");
    let mut entries = archive.entries();
    let descriptor = entries
        .next_entry()
        .expect("first entry parses")
        .expect("first entry exists");
    let deflated = entries
        .next_entry()
        .expect("second entry parses")
        .expect("second entry exists");
    assert!(!descriptor.has_data_descriptor());
    assert!(deflated.has_data_descriptor());
}

#[test]
fn owned_direct_reads_match_zip64_member_metadata() {
    let materialized = open(ZIP64_BYTES.to_vec());
    let owned = open(ZIP64_BYTES.to_vec());
    assert!(materialized.archive_is_zip64());
    assert!(materialized.has_zip64_metadata());
    assert_eq!(
        owned.read_owned("README").expect("direct ZIP64 read"),
        materialized.read("README").expect("member-span ZIP64 read")
    );
}

#[test]
fn owned_direct_preserves_crc_failure_and_missing_member_identity() {
    let mut bytes = descriptor_archive();
    let (_, _, payload_start, _) = entry_offsets(&bytes, "stored.bin");
    bytes[payload_start] ^= 0x80;

    let materialized = open(bytes.clone());
    let owned = open(bytes);
    let materialized_error = materialized
        .read("stored.bin")
        .expect_err("corrupt Store payload must fail");
    let owned_error = owned
        .read_owned("stored.bin")
        .expect_err("corrupt Store payload must fail");
    assert_eq!(family(&materialized_error), "checksum");
    assert_same_failure(&materialized_error, &owned_error);

    let materialized = open(descriptor_archive());
    let owned = open(descriptor_archive());
    let materialized_error = materialized
        .read("missing.bin")
        .expect_err("missing member must fail");
    let owned_error = owned
        .read_owned("missing.bin")
        .expect_err("missing member must fail");
    assert_same_failure(&materialized_error, &owned_error);
    assert_eq!(family(&owned_error), "not-found");
}

#[test]
fn owned_direct_preserves_truncated_deflate_failure() {
    let mut bytes = canonical_deflate_archive();
    let (local, central, _, payload_end) = entry_offsets(&bytes, "truncated.bin");
    let compressed_size = u32::from_le_bytes(bytes[local + 18..local + 22].try_into().unwrap());
    assert!(
        compressed_size > 1,
        "fixture must have a truncatable stream"
    );
    let shortened = compressed_size - 1;
    bytes[local + 18..local + 22].copy_from_slice(&shortened.to_le_bytes());
    bytes[central + 20..central + 24].copy_from_slice(&shortened.to_le_bytes());
    assert!(payload_end > 0);

    let materialized = open(bytes.clone());
    let owned = open(bytes);
    let materialized_error = materialized
        .read("truncated.bin")
        .expect_err("shortened Deflate stream must fail");
    let owned_error = owned
        .read_owned("truncated.bin")
        .expect_err("shortened Deflate stream must fail");
    assert_same_failure(&materialized_error, &owned_error);
}

#[test]
fn owned_direct_keeps_index_resource_limit_refusal_typed() {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("limited.bin", b"1234")
        .expect("limit fixture");
    let bytes = writer.finish_to_bytes().expect("fixture archive finishes");
    let limits = ArchiveLimits {
        max_entry_size: 3,
        ..ArchiveLimits::UNBOUNDED
    };

    let materialized_error = open_with_limits(bytes.clone(), limits)
        .expect_err("member-span index must enforce entry limit");
    let owned_error =
        open_with_limits(bytes, limits).expect_err("owned index must enforce entry limit");
    assert_limit(&materialized_error, 4, 3);
    assert_limit(&owned_error, 4, 3);
    assert_same_failure(&materialized_error, &owned_error);
}
