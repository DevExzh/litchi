#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "the checked-in interoperability fixtures are bounded and deterministic"
)]

//! Public compatibility gates for descriptor-bearing archives emitted by
//! Python's standard-library `zipfile` writer.
//!
//! `ZipFile.open(..., force_zip64=True)` on a non-seekable sink writes a
//! version-4.5 local header with ZIP64 size placeholders even when the final
//! central record uses ordinary ZIP32 sizes.  These tests exercise that shape
//! through both owned and borrowed slice reads, plus positional `ReaderAt`
//! reads.  The unsigned case includes a member whose CRC is itself the
//! optional descriptor signature, so the first four bytes cannot identify the
//! descriptor framing by themselves.

use std::collections::HashMap;

use soapberry_zip::office::{ArchiveLimits, ArchiveReader, IndexedArchive};
use soapberry_zip::{CompressionMethod, PreservationIndex, RECOMMENDED_BUFFER_SIZE, ZipArchive};

const ZIP32_MAX: u64 = u32::MAX as u64;
const DESCRIPTOR_SIGNATURE: u32 = 0x0807_4b50;

const ZIP32: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip32-signed-store-deflate.zip"
);
const LOCAL_ONLY_SIGNED: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-local-only-signed.zip"
);
const LOCAL_ONLY_UNSIGNED: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-local-only-unsigned-crc-marker.zip"
);
const CENTRAL_LOCAL_SIGNED: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-central-local-signed.zip"
);
const SEEKABLE_LOCAL_ONLY: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-local-only-seekable-no-descriptor.zip"
);
const MANY_SMALL: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/many-small-local-only-signed.zip"
);
const MANY_SMALL_ZIP32: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/many-small-zip32-signed.zip"
);
const MANY_SMALL_CENTRAL_LOCAL: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/many-small-central-local-signed.zip"
);
const EMPTY_SIGNED: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-local-only-empty-signed.zip"
);
const EMPTY_UNSIGNED: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-local-only-empty-unsigned.zip"
);
const CENTRAL_LOCAL_ZIP32_TAIL: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-central-local-zip32-tail-signed.zip"
);
const MANY_SMALL_CENTRAL_LOCAL_ZIP32_TAIL: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/many-small-central-local-zip32-tail-signed.zip"
);

const STORE_PAYLOAD: &[u8] = b"stored payload with a deterministic ZIP descriptor\n";
const CRC_MARKER_PAYLOAD: &[u8] = b"unsigned descriptor CRC marker | PK\x07\x08 | \xd3v\xd5\xcb";

fn deflate_payload() -> Vec<u8> {
    let mut payload = b"deflated payload: ".to_vec();
    payload.extend(0_u8..32);
    payload.extend_from_slice(b"\n");
    let repeated = b"repeatable deflate text; ".repeat(11);
    payload.extend_from_slice(&repeated);
    payload
}

#[derive(Debug, Clone, Copy)]
struct Fixture<'a> {
    name: &'static str,
    expected: &'a [u8],
    method: CompressionMethod,
}

fn u16_at(data: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes(
        data[offset..offset + 2].try_into().unwrap(),
    ))
}

fn u32_at(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn u64_at(data: &[u8], offset: usize) -> u64 {
    u64::from_le_bytes(data[offset..offset + 8].try_into().unwrap())
}

#[derive(Debug, Clone)]
struct RawCentral {
    name: Vec<u8>,
    local_offset: usize,
    flags: u16,
    method: u16,
    crc: u32,
    compressed: u64,
    uncompressed: u64,
    central_zip64: bool,
    central_extra: Vec<u8>,
}

fn zip64_extra_pair(extra: &[u8], compressed32: u32, uncompressed32: u32) -> (u64, u64) {
    let mut cursor = 0;
    while cursor + 4 <= extra.len() {
        let field_id = u16_at(extra, cursor);
        let field_len = u16_at(extra, cursor + 2);
        let field_start = cursor + 4;
        let field_end = field_start + field_len;
        assert!(field_end <= extra.len(), "truncated ZIP64 extra field");
        if field_id == 1 {
            let mut field_cursor = 0;
            let uncompressed = if uncompressed32 == u32::MAX {
                let value = u64_at(extra, field_start + field_cursor);
                field_cursor += 8;
                value
            } else {
                u64::from(uncompressed32)
            };
            let compressed = if compressed32 == u32::MAX {
                u64_at(extra, field_start + field_cursor)
            } else {
                u64::from(compressed32)
            };
            return (compressed, uncompressed);
        }
        cursor = field_end;
    }
    (u64::from(compressed32), u64::from(uncompressed32))
}

fn raw_central_records(data: &[u8]) -> Vec<RawCentral> {
    let archive = ZipArchive::from_slice(data).unwrap();
    let mut cursor = usize::try_from(archive.directory_offset()).unwrap();
    let central_end = usize::try_from(archive.head_eocd_offset()).unwrap();
    let count = usize::try_from(archive.entries_hint()).unwrap();
    let mut records = Vec::with_capacity(count);

    for _ in 0..count {
        assert_eq!(u32_at(data, cursor), 0x0201_4b50);
        let name_len = u16_at(data, cursor + 28);
        let extra_len = u16_at(data, cursor + 30);
        let comment_len = u16_at(data, cursor + 32);
        let central_size = 46 + name_len + extra_len + comment_len;
        let name_start = cursor + 46;
        let name_end = name_start + name_len;
        let extra_start = name_end;
        let extra_end = extra_start + extra_len;
        let compressed32 = u32_at(data, cursor + 20);
        let uncompressed32 = u32_at(data, cursor + 24);
        let (compressed, uncompressed) =
            zip64_extra_pair(&data[extra_start..extra_end], compressed32, uncompressed32);
        records.push(RawCentral {
            name: data[name_start..name_end].to_vec(),
            local_offset: usize::try_from(u32_at(data, cursor + 42)).unwrap(),
            flags: u16::try_from(u16_at(data, cursor + 8)).unwrap(),
            method: u16::try_from(u16_at(data, cursor + 10)).unwrap(),
            crc: u32_at(data, cursor + 16),
            compressed,
            uncompressed,
            central_zip64: compressed32 == u32::MAX || uncompressed32 == u32::MAX,
            central_extra: data[extra_start..extra_end].to_vec(),
        });
        cursor += central_size;
    }
    assert_eq!(cursor, central_end, "central directory record framing");
    records
}

fn raw_record<'a>(records: &'a [RawCentral], wanted: &str) -> &'a RawCentral {
    records
        .iter()
        .find(|record| record.name == wanted.as_bytes())
        .unwrap_or_else(|| panic!("missing ZIP entry {wanted}"))
}

fn descriptor_offset(data: &[u8], wanted: &str) -> usize {
    let records = raw_central_records(data);
    let record = raw_record(&records, wanted);
    let local = record.local_offset;
    let name_len = u16_at(data, local + 26);
    let extra_len = u16_at(data, local + 28);
    local + 30 + name_len + extra_len + usize::try_from(record.compressed).unwrap()
}

fn assert_local_zip64_placeholder(data: &[u8], record: &RawCentral) {
    let local = record.local_offset;
    assert_eq!(u16_at(data, local + 4), 45);
    assert_eq!(u16_at(data, local + 6) & 8, 8);
    assert_eq!(u32_at(data, local + 18), u32::MAX);
    assert_eq!(u32_at(data, local + 22), u32::MAX);

    let name_len = u16_at(data, local + 26);
    let extra_len = u16_at(data, local + 28);
    assert_eq!(extra_len, 20);
    let extra_start = local + 30 + name_len;
    assert_eq!(&data[extra_start..extra_start + 4], &[1, 0, 16, 0]);
    assert_eq!(&data[extra_start + 4..extra_start + 20], &[0; 16]);
}

fn assert_signed_descriptor(data: &[u8], record: &RawCentral, width: usize) {
    let descriptor = descriptor_offset(data, std::str::from_utf8(&record.name).unwrap());
    assert!(
        descriptor + width <= ZipArchive::from_slice(data).unwrap().directory_offset() as usize
    );
    assert_eq!(u32_at(data, descriptor), DESCRIPTOR_SIGNATURE);
    assert_eq!(u32_at(data, descriptor + 4), record.crc);
    if width == 16 {
        assert_eq!(u32_at(data, descriptor + 8), record.compressed as u32);
        assert_eq!(u32_at(data, descriptor + 12), record.uncompressed as u32);
    } else {
        assert_eq!(u64_at(data, descriptor + 8), record.compressed);
        assert_eq!(u64_at(data, descriptor + 16), record.uncompressed);
    }
}

fn assert_unsigned_descriptor(data: &[u8], record: &RawCentral, width: usize) {
    let descriptor = descriptor_offset(data, std::str::from_utf8(&record.name).unwrap());
    assert!(
        descriptor + width <= ZipArchive::from_slice(data).unwrap().directory_offset() as usize
    );
    assert_eq!(u32_at(data, descriptor), record.crc);
    if width == 12 {
        assert_eq!(u32_at(data, descriptor + 4), record.compressed as u32);
        assert_eq!(u32_at(data, descriptor + 8), record.uncompressed as u32);
    } else {
        assert_eq!(u64_at(data, descriptor + 4), record.compressed);
        assert_eq!(u64_at(data, descriptor + 12), record.uncompressed);
    }
}

fn assert_local_only_shape(data: &[u8], signed: bool) {
    let archive = ZipArchive::from_slice(data).unwrap();
    assert!(
        !archive.is_zip64(),
        "local-only fixture must keep a ZIP32 tail"
    );
    let records = raw_central_records(data);
    for record in &records {
        assert!(matches!(record.method, 0 | 8));
        assert_eq!(record.flags & 1, 0, "fixtures are unencrypted");
        assert_eq!(record.flags & 8, 8);
        assert!(!record.central_zip64, "central sizes remain ordinary ZIP32");
        assert!(record.compressed < ZIP32_MAX);
        assert!(record.uncompressed < ZIP32_MAX);
        assert_local_zip64_placeholder(data, record);
        if signed {
            assert_signed_descriptor(data, record, 24);
        } else {
            assert_unsigned_descriptor(data, record, 20);
        }
    }
}

fn assert_zip32_shape(data: &[u8]) {
    let archive = ZipArchive::from_slice(data).unwrap();
    assert!(!archive.is_zip64());
    for record in raw_central_records(data) {
        assert_eq!(record.flags & 8, 8);
        assert!(!record.central_zip64);
        let local = record.local_offset;
        assert_eq!(u16_at(data, local + 4), 20);
        assert_eq!(u32_at(data, local + 18), 0);
        assert_eq!(u32_at(data, local + 22), 0);
        assert_signed_descriptor(data, &record, 16);
    }
}

fn assert_central_local_shape(data: &[u8], expect_archive_zip64: bool) {
    let archive = ZipArchive::from_slice(data).unwrap();
    assert_eq!(archive.is_zip64(), expect_archive_zip64);
    for record in raw_central_records(data) {
        assert_eq!(record.flags & 8, 8);
        assert!(record.central_zip64);
        assert_local_zip64_placeholder(data, &record);
        assert_eq!(record.central_extra.len(), 20);
        assert_eq!(&record.central_extra[..4], &[1, 0, 16, 0]);
        let (compressed, uncompressed) =
            zip64_extra_pair(&record.central_extra, u32::MAX, u32::MAX);
        assert_eq!(compressed, record.compressed);
        assert_eq!(uncompressed, record.uncompressed);
        assert_signed_descriptor(data, &record, 24);
    }
}

fn assert_seekable_shape(data: &[u8]) {
    let archive = ZipArchive::from_slice(data).unwrap();
    assert!(!archive.is_zip64(), "the terminal remains ordinary ZIP32");
    for record in raw_central_records(data) {
        assert_eq!(record.flags & 8, 0, "seekable writes need no descriptor");
        assert!(!record.central_zip64);
        let local = record.local_offset;
        assert_eq!(u16_at(data, local + 4), 45);
        assert_eq!(u32_at(data, local + 18), u32::MAX);
        assert_eq!(u32_at(data, local + 22), u32::MAX);
        let name_len = u16_at(data, local + 26);
        let extra_len = u16_at(data, local + 28);
        assert_eq!(extra_len, 20);
        let extra_start = local + 30 + name_len;
        assert_eq!(&data[extra_start..extra_start + 4], &[1, 0, 16, 0]);
        let local_uncompressed = u64_at(data, extra_start + 4);
        let local_compressed = u64_at(data, extra_start + 12);
        assert_eq!(local_compressed, record.compressed);
        assert_eq!(local_uncompressed, record.uncompressed);
    }
}

fn assert_read_paths(data: &[u8], entries: &[Fixture<'_>], expect_archive_zip64: bool) {
    let archive = ArchiveReader::new(data).unwrap();
    assert_eq!(archive.len(), entries.len());
    let indexed = IndexedArchive::from_reader(data, data.len() as u64).unwrap();
    assert_eq!(indexed.len(), entries.len());
    assert_eq!(indexed.archive_is_zip64(), expect_archive_zip64);

    for fixture in entries {
        assert_eq!(
            archive.is_stored(fixture.name).unwrap(),
            fixture.method == CompressionMethod::Store
        );
        let output = archive.read(fixture.name).unwrap();
        assert_eq!(output.as_slice(), fixture.expected);

        let mut output_to = Vec::new();
        assert_eq!(
            archive.read_to(fixture.name, &mut output_to).unwrap(),
            fixture.expected.len() as u64
        );
        assert_eq!(output_to.as_slice(), fixture.expected);

        let borrowed = archive.read_stored_borrowed(fixture.name).unwrap();
        if fixture.method == CompressionMethod::Store && !expect_archive_zip64 {
            assert_eq!(borrowed, Some(fixture.expected));
        } else {
            assert!(borrowed.is_none());
        }

        let entry_id = indexed.entry_id(fixture.name).unwrap();
        let indexed_output = indexed.read(fixture.name).unwrap();
        assert_eq!(indexed_output.as_slice(), fixture.expected);
        let mut indexed_to = Vec::new();
        assert_eq!(
            indexed.read_entry_to(entry_id, &mut indexed_to).unwrap(),
            fixture.expected.len() as u64
        );
        assert_eq!(indexed_to.as_slice(), fixture.expected);
    }
}

#[test]
fn python_zip32_signed_store_and_deflate_are_read_and_verified() {
    assert_zip32_shape(ZIP32);
    let deflate = deflate_payload();
    let entries = [
        Fixture {
            name: "store.bin",
            expected: STORE_PAYLOAD,
            method: CompressionMethod::Store,
        },
        Fixture {
            name: "deflate.bin",
            expected: &deflate,
            method: CompressionMethod::Deflate,
        },
    ];
    assert_read_paths(ZIP32, &entries, false);
}

#[test]
fn python_force_zip64_local_placeholders_use_64_bit_signed_descriptors() {
    assert_local_only_shape(LOCAL_ONLY_SIGNED, true);
    let deflate = deflate_payload();
    let entries = [
        Fixture {
            name: "store.bin",
            expected: STORE_PAYLOAD,
            method: CompressionMethod::Store,
        },
        Fixture {
            name: "deflate.bin",
            expected: &deflate,
            method: CompressionMethod::Deflate,
        },
    ];
    assert_read_paths(LOCAL_ONLY_SIGNED, &entries, false);
}

#[test]
fn python_force_zip64_unsigned_crc_signature_ambiguity_is_verified() {
    assert_local_only_shape(LOCAL_ONLY_UNSIGNED, false);
    let records = raw_central_records(LOCAL_ONLY_UNSIGNED);
    let marker = raw_record(&records, "crc-marker.bin");
    assert_eq!(marker.crc, DESCRIPTOR_SIGNATURE);
    let descriptor = descriptor_offset(LOCAL_ONLY_UNSIGNED, "crc-marker.bin");
    assert_eq!(
        u32_at(LOCAL_ONLY_UNSIGNED, descriptor),
        DESCRIPTOR_SIGNATURE
    );

    let deflate = deflate_payload();
    let entries = [
        Fixture {
            name: "crc-marker.bin",
            expected: CRC_MARKER_PAYLOAD,
            method: CompressionMethod::Store,
        },
        Fixture {
            name: "deflate.bin",
            expected: &deflate,
            method: CompressionMethod::Deflate,
        },
    ];
    assert_read_paths(LOCAL_ONLY_UNSIGNED, &entries, false);
}

#[test]
fn python_central_and_local_zip64_signed_descriptor_is_verified() {
    assert_central_local_shape(CENTRAL_LOCAL_SIGNED, true);
    let archive = ZipArchive::from_slice(CENTRAL_LOCAL_SIGNED).unwrap();
    assert!(archive.is_zip64());
    let deflate = deflate_payload();
    let entries = [
        Fixture {
            name: "store.bin",
            expected: STORE_PAYLOAD,
            method: CompressionMethod::Store,
        },
        Fixture {
            name: "deflate.bin",
            expected: &deflate,
            method: CompressionMethod::Deflate,
        },
    ];
    assert_read_paths(CENTRAL_LOCAL_SIGNED, &entries, true);
}

#[test]
fn python_central_zip64_entries_with_zip32_tail_use_borrowed_and_reader_at_paths() {
    assert_central_local_shape(CENTRAL_LOCAL_ZIP32_TAIL, false);
    let deflate = deflate_payload();
    let entries = [
        Fixture {
            name: "store.bin",
            expected: STORE_PAYLOAD,
            method: CompressionMethod::Store,
        },
        Fixture {
            name: "deflate.bin",
            expected: &deflate,
            method: CompressionMethod::Deflate,
        },
    ];
    assert_read_paths(CENTRAL_LOCAL_ZIP32_TAIL, &entries, false);
}

#[test]
fn seekable_force_zip64_actual_local_sizes_match_ordinary_central_records() {
    assert_seekable_shape(SEEKABLE_LOCAL_ONLY);
    let deflate = deflate_payload();
    let entries = [
        Fixture {
            name: "seekable-store.bin",
            expected: STORE_PAYLOAD,
            method: CompressionMethod::Store,
        },
        Fixture {
            name: "seekable-deflate.bin",
            expected: &deflate,
            method: CompressionMethod::Deflate,
        },
    ];
    assert_read_paths(SEEKABLE_LOCAL_ONLY, &entries, false);
}

#[test]
fn empty_local_only_zip64_signed_and_unsigned_descriptors_keep_width_eight() {
    assert_local_only_shape(EMPTY_SIGNED, true);
    assert_local_only_shape(EMPTY_UNSIGNED, false);
    let entries = [Fixture {
        name: "empty.bin",
        expected: &[],
        method: CompressionMethod::Store,
    }];
    assert_read_paths(EMPTY_SIGNED, &entries, false);
    assert_read_paths(EMPTY_UNSIGNED, &entries, false);
}

#[test]
fn many_small_controls_cover_zip32_local_only_and_central_local_zip64() {
    assert_eq!(
        ZipArchive::from_slice(MANY_SMALL).unwrap().entries_hint(),
        256
    );
    assert_eq!(
        ZipArchive::from_slice(MANY_SMALL_ZIP32)
            .unwrap()
            .entries_hint(),
        256
    );
    assert_eq!(
        ZipArchive::from_slice(MANY_SMALL_CENTRAL_LOCAL)
            .unwrap()
            .entries_hint(),
        256
    );
    assert_eq!(
        ZipArchive::from_slice(MANY_SMALL_CENTRAL_LOCAL_ZIP32_TAIL)
            .unwrap()
            .entries_hint(),
        256
    );
    assert_local_only_shape(MANY_SMALL, true);
    assert_zip32_shape(MANY_SMALL_ZIP32);
    assert_central_local_shape(MANY_SMALL_CENTRAL_LOCAL, true);
    assert_central_local_shape(MANY_SMALL_CENTRAL_LOCAL_ZIP32_TAIL, false);

    for (data, global_zip64) in [
        (MANY_SMALL, false),
        (MANY_SMALL_ZIP32, false),
        (MANY_SMALL_CENTRAL_LOCAL, true),
        (MANY_SMALL_CENTRAL_LOCAL_ZIP32_TAIL, false),
    ] {
        let reader = ArchiveReader::new(data).unwrap();
        let indexed = IndexedArchive::from_reader(data, data.len() as u64).unwrap();
        assert_eq!(indexed.archive_is_zip64(), global_zip64);
        for index in 0_u8..=u8::MAX {
            let name = format!("parts/p{index:03}.bin");
            let expected: [u8; 4] = [index, 0xA5, 0x15, 0x41];
            assert_eq!(reader.read(&name).unwrap().as_slice(), expected.as_slice());
            let mut output = Vec::new();
            reader.read_to(&name, &mut output).unwrap();
            assert_eq!(output.as_slice(), expected.as_slice());
            let borrowed = reader.read_stored_borrowed(&name).unwrap();
            if global_zip64 {
                assert!(borrowed.is_none());
            } else {
                assert_eq!(borrowed, Some(&expected[..]));
            }
            assert_eq!(indexed.read(&name).unwrap().as_slice(), expected.as_slice());
        }
    }
}

fn assert_strict_read_rejected(data: &[u8], name: &str) {
    let reader = ArchiveReader::new(data).unwrap();
    let mut output = Vec::new();
    assert!(reader.read_to(name, &mut output).is_err());
    let indexed = IndexedArchive::from_reader(data, data.len() as u64).unwrap();
    assert!(indexed.read(name).is_err());
}

#[test]
fn malformed_local_zip64_descriptor_and_size_fields_are_rejected_strictly() {
    let descriptor = descriptor_offset(LOCAL_ONLY_SIGNED, "store.bin");

    let mut bad_crc = LOCAL_ONLY_SIGNED.to_vec();
    bad_crc[descriptor + 4] ^= 1;
    assert_strict_read_rejected(&bad_crc, "store.bin");

    let mut bad_size = LOCAL_ONLY_SIGNED.to_vec();
    bad_size[descriptor + 8] ^= 1;
    assert_strict_read_rejected(&bad_size, "store.bin");

    let local_records = raw_central_records(LOCAL_ONLY_SIGNED);
    let record = raw_record(&local_records, "store.bin");
    let mut bad_placeholder = LOCAL_ONLY_SIGNED.to_vec();
    let extra_start = record.local_offset + 30 + u16_at(&bad_placeholder, record.local_offset + 26);
    bad_placeholder[extra_start + 4] = 1;
    assert_strict_read_rejected(&bad_placeholder, "store.bin");

    let seekable_records = raw_central_records(SEEKABLE_LOCAL_ONLY);
    let seekable_record = raw_record(&seekable_records, "seekable-store.bin");
    let mut bad_seekable_size = SEEKABLE_LOCAL_ONLY.to_vec();
    let extra_start = seekable_record.local_offset
        + 30
        + u16_at(&bad_seekable_size, seekable_record.local_offset + 26);
    bad_seekable_size[extra_start + 4] ^= 1;
    assert_strict_read_rejected(&bad_seekable_size, "seekable-store.bin");
}

#[derive(Debug, Clone)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
    let archive = ZipArchive::from_slice(data).unwrap().into_zip_archive();
    let mut scratch = vec![0; RECOMMENDED_BUFFER_SIZE];
    let index =
        PreservationIndex::new_with_limits(&archive, &mut scratch, ArchiveLimits::UNBOUNDED)
            .unwrap();
    index
        .entries()
        .iter()
        .map(|entry| {
            let local = entry.local_span();
            let central = entry.central_record();
            (
                entry.raw_name_bytes().to_vec(),
                RawRecord {
                    local: data[local.start as usize..local.end as usize].to_vec(),
                    central: data[central.start as usize..central.end as usize].to_vec(),
                },
            )
        })
        .collect()
}

#[test]
fn local_only_zip64_preservation_index_retains_descriptor_and_central_record_spans() {
    let records = raw_records(LOCAL_ONLY_SIGNED);
    assert_eq!(records.len(), 2);
    for record in records.values() {
        assert!(
            record
                .local
                .windows(4)
                .any(|window| window == b"PK\x07\x08")
        );
        assert_eq!(u32_at(&record.central, 0), 0x0201_4b50);
        let expected_len = 46
            + u16_at(&record.central, 28)
            + u16_at(&record.central, 30)
            + u16_at(&record.central, 32);
        assert_eq!(record.central.len(), expected_len);
    }
}
