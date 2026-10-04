#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The adversarial ZIP fixtures are small and intentionally bounded."
)]

//! Adversarial differential coverage for the owned-source direct-read path.
//!
//! The ordinary owned-read tests use the repository's ZIP64 EOCD asset. This
//! file adds a forced local-ZIP64 member and malformed local/member framing so
//! the direct route is checked against the generic `MemberSpan` route at the
//! same validation boundaries.

use flate2::{Compression, write::DeflateEncoder};
use soapberry_zip::office::{ArchiveLimits, IndexedArchive, StreamingArchiveWriter};
use soapberry_zip::{CompressionMethod, Error, ErrorKind, ZipArchive, ZipArchiveWriter};
use std::io::{Cursor, Write};
use std::sync::Arc;

fn open(bytes: Vec<u8>) -> IndexedArchive<Arc<Vec<u8>>> {
    let source = Arc::new(bytes);
    let length = source.len() as u64;
    IndexedArchive::from_reader_with_limits(source, length, ArchiveLimits::UNBOUNDED)
        .expect("fixture archive opens")
}

fn sized_store_archive() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("member.bin", b"declared-size payload")
        .expect("stored fixture");
    writer.finish_to_bytes().expect("fixture archive finishes")
}

fn descriptor_archive() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated("member.bin", b"descriptor payload with stable bytes")
        .expect("Deflate fixture");
    writer.finish_to_bytes().expect("fixture archive finishes")
}

/// A small member whose local header uses ZIP64 size placeholders and a
/// signed 64-bit data descriptor despite its ordinary-sized payload.
fn forced_local_zip64_archive() -> Vec<u8> {
    let payload = b"forced local ZIP64 payload";
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::new(&mut output);
    let (mut entry, config) = archive
        .new_file("forced.bin")
        .compression_method(CompressionMethod::Deflate)
        .zip64(true)
        .start()
        .expect("forced local ZIP64 entry starts");
    let encoder = DeflateEncoder::new(&mut entry, Compression::default());
    let mut writer = config.wrap(encoder);
    writer.write_all(payload).expect("payload write");
    let (encoder, descriptor) = writer.finish().expect("Deflate finish");
    encoder.finish().expect("encoder finish");
    entry.finish(descriptor).expect("entry finish");
    archive.finish().expect("archive finish");
    output.into_inner()
}

fn local_offset(bytes: &[u8], name: &str) -> usize {
    let archive = ZipArchive::from_slice(bytes).expect("fixture archive parses");
    let mut entries = archive.entries();
    while let Some(record) = entries.next_entry().expect("fixture entry parses") {
        if record.file_path().as_ref() == name.as_bytes() {
            return record.local_header_offset() as usize;
        }
    }
    panic!("missing local member {name:?}")
}

fn central_offset(bytes: &[u8], name: &str) -> usize {
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
        .unwrap_or_else(|| panic!("missing central member {name:?}"))
}

fn compressed_range(bytes: &[u8], name: &str) -> (usize, usize) {
    let archive = ZipArchive::from_slice(bytes).expect("fixture archive parses");
    let mut entries = archive.entries();
    while let Some(record) = entries.next_entry().expect("fixture entry parses") {
        if record.file_path().as_ref() == name.as_bytes() {
            let entry = archive
                .get_entry(record.wayfinder())
                .expect("fixture local record parses");
            let (start, end) = entry.compressed_data_range();
            return (start as usize, end as usize);
        }
    }
    panic!("missing payload member {name:?}")
}

fn family(error: &Error) -> &'static str {
    match error.kind() {
        ErrorKind::InvalidChecksum { .. } => "checksum",
        ErrorKind::InvalidSize { .. } => "size",
        ErrorKind::InvalidInput { .. } => "invalid-input",
        ErrorKind::InvalidSignature { .. } => "signature",
        ErrorKind::Eof => "eof",
        ErrorKind::FileNotFound(_) => "not-found",
        ErrorKind::IO(source) | ErrorKind::Io(source) => source
            .get_ref()
            .and_then(|nested| nested.downcast_ref::<Error>())
            .map_or("io", family),
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

fn assert_read_failure(bytes: Vec<u8>, expected_family: Option<&str>) {
    let materialized = open(bytes.clone());
    let owned = open(bytes);
    let materialized = materialized
        .read("member.bin")
        .expect_err("the malformed member must fail");
    let owned = owned
        .read_owned("member.bin")
        .expect_err("the malformed member must fail");
    assert_same_failure(&materialized, &owned);
    assert!(materialized.is_content_fault());
    if let Some(expected) = expected_family {
        assert_eq!(family(&materialized), expected);
    }
}

#[test]
fn owned_direct_reads_forced_local_zip64_with_64_bit_descriptor() {
    let bytes = forced_local_zip64_archive();
    let archive = ZipArchive::from_slice(&bytes).expect("forced ZIP64 fixture parses");
    assert!(archive.is_zip64());
    let mut entries = archive.entries();
    let record = entries
        .next_entry()
        .expect("forced ZIP64 entry parses")
        .expect("forced ZIP64 entry exists");
    assert!(record.is_zip64());
    assert!(record.has_data_descriptor());
    assert_eq!(record.compression_method(), CompressionMethod::Deflate);

    let local = record.local_header_offset() as usize;
    assert_eq!(
        u16::from_le_bytes(bytes[local + 4..local + 6].try_into().unwrap()),
        45
    );
    assert_eq!(
        u32::from_le_bytes(bytes[local + 18..local + 22].try_into().unwrap()),
        u32::MAX
    );
    assert_eq!(
        u32::from_le_bytes(bytes[local + 22..local + 26].try_into().unwrap()),
        u32::MAX
    );
    let name_len = u16::from_le_bytes(bytes[local + 26..local + 28].try_into().unwrap()) as usize;
    let extra_len = u16::from_le_bytes(bytes[local + 28..local + 30].try_into().unwrap()) as usize;
    assert_eq!(extra_len, 20);
    assert_eq!(
        &bytes[local + 30 + name_len..local + 34 + name_len],
        &[1, 0, 16, 0]
    );

    let (_, payload_end) = compressed_range(&bytes, "forced.bin");
    assert_eq!(
        u32::from_le_bytes(bytes[payload_end..payload_end + 4].try_into().unwrap()),
        0x0807_4b50
    );
    assert_eq!(archive.directory_offset() as usize - payload_end, 24);

    let materialized = open(bytes.clone());
    let owned = open(bytes);
    assert_eq!(
        owned.read_owned("forced.bin").expect("direct ZIP64 read"),
        materialized
            .read("forced.bin")
            .expect("member-span ZIP64 read")
    );
}

#[test]
fn owned_direct_rejects_corrupt_descriptor_with_same_typed_error() {
    let mut bytes = descriptor_archive();
    let (_, payload_end) = compressed_range(&bytes, "member.bin");
    assert_eq!(
        u32::from_le_bytes(bytes[payload_end..payload_end + 4].try_into().unwrap()),
        0x0807_4b50,
        "fixture uses a signed data descriptor"
    );
    bytes[payload_end + 4] ^= 1;
    assert_read_failure(bytes, Some("checksum"));
}

#[test]
fn owned_direct_rejects_under_and_over_declared_sizes_identically() {
    let baseline = sized_store_archive();
    let local = local_offset(&baseline, "member.bin");
    let central = central_offset(&baseline, "member.bin");
    let actual = u32::from_le_bytes(baseline[local + 22..local + 26].try_into().unwrap());
    assert!(actual > 1);

    for declared in [actual - 1, actual + 1] {
        let mut bytes = baseline.clone();
        bytes[local + 22..local + 26].copy_from_slice(&declared.to_le_bytes());
        bytes[central + 24..central + 28].copy_from_slice(&declared.to_le_bytes());
        assert_read_failure(bytes, Some("size"));
    }
}

#[test]
fn owned_direct_preserves_local_header_and_source_bound_errors() {
    let baseline = sized_store_archive();
    let local = local_offset(&baseline, "member.bin");
    let central = central_offset(&baseline, "member.bin");

    let mut truncated_local = baseline.clone();
    truncated_local[local + 28..local + 30].copy_from_slice(&u16::MAX.to_le_bytes());
    assert_read_failure(truncated_local, None);

    let mut out_of_bounds = baseline;
    let invalid_offset = (out_of_bounds.len() - 1) as u32;
    out_of_bounds[central + 42..central + 46].copy_from_slice(&invalid_offset.to_le_bytes());
    assert_read_failure(out_of_bounds, None);
}
