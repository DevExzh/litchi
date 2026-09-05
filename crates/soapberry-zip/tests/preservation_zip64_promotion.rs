#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The ZIP64 integration fixtures are fixed, bounded metadata probes."
)]

//! Public preservation coverage for generated members whose local offsets
//! cross the ZIP32 offset sentinel.  The source and output images are sparse:
//! the large stored member is represented as a zero run, while every framing
//! byte is retained.  This keeps the test's memory use bounded while still
//! exercising the real sequential copy loop and the public reopen path.

use crc32fast::Hasher;
use soapberry_zip::office::ArchiveLimits;
use soapberry_zip::{
    CompressionMethod, ErrorKind, PreservationIndex, PreservationPlan, ReaderAt, RegeneratedEntry,
    ZipLocator, ZipOperationAccounting,
};
use std::io::{self, Read, Write};
use std::ops::Range;
use std::sync::{Arc, Mutex};

const ZIP64_EOCD_SIGNATURE: u32 = 0x0606_4b50;
const ZIP64_LOCATOR_SIGNATURE: u32 = 0x0706_4b50;
const EOCD_SIGNATURE: u32 = 0x0605_4b50;
const CENTRAL_SIGNATURE: u32 = 0x0201_4b50;
const LOCAL_SIGNATURE: u32 = 0x0403_4b50;
const DATA_DESCRIPTOR_SIGNATURE: u32 = 0x0807_4b50;
const ZIP64_EXTRA_ID: u16 = 1;
const ZIP64_VERSION: u16 = 45;
const DATA_DESCRIPTOR_FLAGS: u16 = 1 << 3;

const SOURCE_NAME: &[u8] = b"huge.bin";
const GENERATED_NAME: &str = "generated.bin";
const GENERATED_PAYLOAD: &[u8] = b"generated after the ZIP64 offset";
const ZIP64_EXTENSIBLE_DATA: &[u8] = &[0xa1, 0xb2, 0xc3, 0xd4];
const ARCHIVE_COMMENT: &[u8] = b"sparse ZIP64 preservation comment";
const COPY_CHUNK_SIZE: usize = 32 * 1024;
const MAX_LOCATOR_PROBE: u64 = 2 * 1024 * 1024;

#[derive(Debug, Clone)]
struct Segment {
    start: u64,
    bytes: Arc<[u8]>,
}

impl Segment {
    fn end(&self) -> u64 {
        self.start + self.bytes.len() as u64
    }
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
struct ReadStats {
    calls: u64,
    bytes: u64,
    synthetic_zero_bytes: u64,
    max_request: usize,
}

/// A byte-stable, positional image with small finite segments and explicit
/// synthetic zero ranges.  Missing bytes are errors, so reopening a captured
/// output cannot accidentally hide an omitted framing region.
#[derive(Clone)]
struct SparseReader {
    length: u64,
    segments: Arc<Vec<Segment>>,
    zero_ranges: Arc<Vec<Range<u64>>>,
    stats: Arc<Mutex<ReadStats>>,
}

impl SparseReader {
    fn new(length: u64, segments: Vec<Segment>, zero_ranges: Vec<Range<u64>>) -> Self {
        Self {
            length,
            segments: Arc::new(segments),
            zero_ranges: Arc::new(zero_ranges),
            stats: Arc::new(Mutex::new(ReadStats::default())),
        }
    }

    fn stats(&self) -> ReadStats {
        *self.stats.lock().unwrap()
    }

    fn snapshot(&self) -> (u64, Vec<(u64, Vec<u8>)>, Vec<Range<u64>>) {
        (
            self.length,
            self.segments
                .iter()
                .map(|segment| (segment.start, segment.bytes.to_vec()))
                .collect(),
            self.zero_ranges.as_ref().clone(),
        )
    }

    fn segment_at(&self, offset: u64) -> Option<&Segment> {
        self.segments
            .iter()
            .find(|segment| segment.start <= offset && offset < segment.end())
    }

    fn zero_range_at(&self, offset: u64) -> Option<&Range<u64>> {
        self.zero_ranges
            .iter()
            .find(|range| range.start <= offset && offset < range.end)
    }

    fn next_known_offset(&self, offset: u64, end: u64) -> u64 {
        self.segments
            .iter()
            .map(|segment| segment.start)
            .chain(self.zero_ranges.iter().map(|range| range.start))
            .filter(|start| *start > offset)
            .min()
            .unwrap_or(end)
    }
}

impl ReaderAt for SparseReader {
    fn read_at(&self, output: &mut [u8], offset: u64) -> io::Result<usize> {
        let mut stats = self.stats.lock().unwrap();
        stats.calls = stats.calls.saturating_add(1);
        stats.max_request = stats.max_request.max(output.len());
        if offset >= self.length || output.is_empty() {
            return Ok(0);
        }

        let available = (self.length - offset).min(output.len() as u64) as usize;
        let end = offset + available as u64;
        let mut cursor = offset;
        while cursor < end {
            let output_offset = (cursor - offset) as usize;
            if let Some(range) = self.zero_range_at(cursor) {
                let amount = (range.end.min(end) - cursor) as usize;
                output[output_offset..output_offset + amount].fill(0);
                stats.synthetic_zero_bytes += amount as u64;
                cursor += amount as u64;
                continue;
            }
            if let Some(segment) = self.segment_at(cursor) {
                let amount = (segment.end().min(end) - cursor) as usize;
                let segment_offset = (cursor - segment.start) as usize;
                output[output_offset..output_offset + amount]
                    .copy_from_slice(&segment.bytes[segment_offset..segment_offset + amount]);
                cursor += amount as u64;
                continue;
            }
            let missing_end = self.next_known_offset(cursor, end).min(end);
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                format!("sparse fixture has no bytes at {cursor}..{missing_end}"),
            ));
        }
        stats.bytes += available as u64;
        Ok(available)
    }
}

/// A forward-only sink that retains every byte outside one explicitly
/// discarded source-payload range.  Discarded bytes are checked as they pass,
/// which proves the entire sparse payload was copied byte-for-byte without
/// retaining a multi-gigabyte output vector.
#[derive(Debug)]
struct RegionSink {
    length: u64,
    writes: u64,
    max_write: usize,
    discarded: Range<u64>,
    discarded_bytes: u64,
    segments: Vec<Segment>,
}

impl RegionSink {
    fn new(discarded: Range<u64>) -> Self {
        Self {
            length: 0,
            writes: 0,
            max_write: 0,
            discarded,
            discarded_bytes: 0,
            segments: Vec::new(),
        }
    }

    fn append_segment(&mut self, start: u64, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        if self
            .segments
            .last()
            .is_some_and(|previous| previous.end() == start)
        {
            let previous = self.segments.last_mut().unwrap();
            let mut joined = previous.bytes.to_vec();
            joined.extend_from_slice(bytes);
            previous.bytes = Arc::from(joined);
            return;
        }
        self.segments.push(Segment {
            start,
            bytes: Arc::from(bytes.to_vec()),
        });
    }

    fn retained_bytes(&self) -> u64 {
        self.segments
            .iter()
            .map(|segment| segment.bytes.len() as u64)
            .sum()
    }

    fn into_reader(self) -> SparseReader {
        SparseReader::new(self.length, self.segments, vec![self.discarded])
    }
}

impl Write for RegionSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        let start = self.length;
        let end = start
            .checked_add(input.len() as u64)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "sink length overflow"))?;

        let discarded_start = start.max(self.discarded.start);
        let discarded_end = end.min(self.discarded.end);
        if discarded_start < discarded_end {
            let input_start = (discarded_start - start) as usize;
            let input_end = (discarded_end - start) as usize;
            if input[input_start..input_end].iter().any(|byte| *byte != 0) {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "sparse source payload was not copied byte-for-byte",
                ));
            }
            self.discarded_bytes += discarded_end - discarded_start;
        }

        let retained_prefix_end = end.min(self.discarded.start);
        if start < retained_prefix_end {
            self.append_segment(start, &input[..(retained_prefix_end - start) as usize]);
        }
        let retained_suffix_start = start.max(self.discarded.end);
        if retained_suffix_start < end {
            self.append_segment(
                retained_suffix_start,
                &input[(retained_suffix_start - start) as usize..],
            );
        }

        self.length = end;
        self.writes += 1;
        self.max_write = self.max_write.max(input.len());
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct SparseZipFixture {
    source: SparseReader,
    source_length: u64,
    payload_start: u64,
    payload_end: u64,
    central_start: u64,
    central_bytes: Vec<u8>,
    tail_extensible_data: Vec<u8>,
    comment: Vec<u8>,
}

fn push_u16(bytes: &mut Vec<u8>, value: u16) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

fn push_u32(bytes: &mut Vec<u8>, value: u32) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

fn push_u64(bytes: &mut Vec<u8>, value: u64) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

/// Computes CRC32(0 repeated `length` times) using logarithmic-size CRC
/// composition.  The small direct comparison test below guards the helper's
/// use of `Hasher::combine`.
fn crc32_zeroes(length: u64) -> u32 {
    if length == 0 {
        return Hasher::new().finalize();
    }

    let mut powers = Vec::new();
    let mut one = Hasher::new();
    one.update(&[0]);
    powers.push(one);
    let mut block_length = 1_u64;
    while block_length < length {
        let previous = powers.last().unwrap();
        let mut doubled = previous.clone();
        doubled.combine(previous);
        powers.push(doubled);
        block_length = block_length.saturating_mul(2);
    }

    let mut remaining = length;
    let mut bit = 0_usize;
    let mut combined = Hasher::new();
    while remaining != 0 {
        if remaining & 1 == 1 {
            combined.combine(&powers[bit]);
        }
        remaining >>= 1;
        bit += 1;
    }
    combined.finalize()
}

fn sparse_zip_fixture(generated_offset: u64) -> SparseZipFixture {
    let local_prefix_len = 30_u64 + SOURCE_NAME.len() as u64;
    let descriptor_len = 24_u64;
    assert!(generated_offset > local_prefix_len + descriptor_len);
    let payload_len = generated_offset - local_prefix_len - descriptor_len;
    let payload_crc = crc32_zeroes(payload_len);

    let mut local = Vec::with_capacity(local_prefix_len as usize);
    push_u32(&mut local, LOCAL_SIGNATURE);
    push_u16(&mut local, 20);
    push_u16(&mut local, DATA_DESCRIPTOR_FLAGS);
    push_u16(&mut local, CompressionMethod::Store.as_id().as_u16());
    push_u16(&mut local, 0);
    push_u16(&mut local, 0);
    push_u32(&mut local, 0);
    push_u32(&mut local, 0);
    push_u32(&mut local, 0);
    push_u16(&mut local, SOURCE_NAME.len() as u16);
    push_u16(&mut local, 0);
    local.extend_from_slice(SOURCE_NAME);
    assert_eq!(local.len() as u64, local_prefix_len);

    let mut descriptor = Vec::with_capacity(descriptor_len as usize);
    push_u32(&mut descriptor, DATA_DESCRIPTOR_SIGNATURE);
    push_u32(&mut descriptor, payload_crc);
    push_u64(&mut descriptor, payload_len);
    push_u64(&mut descriptor, payload_len);
    assert_eq!(descriptor.len() as u64, descriptor_len);

    let mut central = Vec::new();
    push_u32(&mut central, CENTRAL_SIGNATURE);
    push_u16(&mut central, ZIP64_VERSION);
    push_u16(&mut central, ZIP64_VERSION);
    push_u16(&mut central, DATA_DESCRIPTOR_FLAGS);
    push_u16(&mut central, CompressionMethod::Store.as_id().as_u16());
    push_u16(&mut central, 0);
    push_u16(&mut central, 0);
    push_u32(&mut central, payload_crc);
    push_u32(&mut central, u32::MAX);
    push_u32(&mut central, u32::MAX);
    push_u16(&mut central, SOURCE_NAME.len() as u16);
    push_u16(&mut central, 27);
    push_u16(&mut central, 0);
    push_u16(&mut central, 0);
    push_u16(&mut central, 0);
    push_u32(&mut central, 0xcafe_babe);
    push_u32(&mut central, 0);
    central.extend_from_slice(SOURCE_NAME);
    push_u16(&mut central, ZIP64_EXTRA_ID);
    push_u16(&mut central, 16);
    push_u64(&mut central, payload_len);
    push_u64(&mut central, payload_len);
    push_u16(&mut central, 0xaaaa);
    push_u16(&mut central, 3);
    central.extend_from_slice(&[0xde, 0xad, 0xbe]);
    assert_eq!(central.len(), 46 + SOURCE_NAME.len() + 27);
    let central_len = central.len() as u64;

    let zip64_eocd_offset = generated_offset + central_len;
    let mut zip64_eocd = Vec::new();
    push_u32(&mut zip64_eocd, ZIP64_EOCD_SIGNATURE);
    push_u64(&mut zip64_eocd, 44 + ZIP64_EXTENSIBLE_DATA.len() as u64);
    push_u16(&mut zip64_eocd, ZIP64_VERSION);
    push_u16(&mut zip64_eocd, ZIP64_VERSION);
    push_u32(&mut zip64_eocd, 0);
    push_u32(&mut zip64_eocd, 0);
    push_u64(&mut zip64_eocd, 1);
    push_u64(&mut zip64_eocd, 1);
    push_u64(&mut zip64_eocd, central_len);
    push_u64(&mut zip64_eocd, generated_offset);
    zip64_eocd.extend_from_slice(ZIP64_EXTENSIBLE_DATA);
    assert_eq!(zip64_eocd.len(), 56 + ZIP64_EXTENSIBLE_DATA.len());

    let locator_offset = zip64_eocd_offset + zip64_eocd.len() as u64;
    let mut locator = Vec::new();
    push_u32(&mut locator, ZIP64_LOCATOR_SIGNATURE);
    push_u32(&mut locator, 0);
    push_u64(&mut locator, zip64_eocd_offset);
    push_u32(&mut locator, 1);
    assert_eq!(locator.len(), 20);

    let eocd_offset = locator_offset + locator.len() as u64;
    let mut eocd = Vec::new();
    push_u32(&mut eocd, EOCD_SIGNATURE);
    push_u16(&mut eocd, 0);
    push_u16(&mut eocd, 0);
    push_u16(&mut eocd, u16::MAX);
    push_u16(&mut eocd, u16::MAX);
    push_u32(&mut eocd, u32::MAX);
    push_u32(&mut eocd, u32::MAX);
    push_u16(&mut eocd, ARCHIVE_COMMENT.len() as u16);
    assert_eq!(eocd.len(), 22);

    let comment_offset = eocd_offset + eocd.len() as u64;
    let source_length = comment_offset + ARCHIVE_COMMENT.len() as u64;
    let central_bytes = central.clone();
    let mut tail = central;
    tail.extend_from_slice(&zip64_eocd);
    tail.extend_from_slice(&locator);
    tail.extend_from_slice(&eocd);
    tail.extend_from_slice(ARCHIVE_COMMENT);

    let source = SparseReader::new(
        source_length,
        vec![
            Segment {
                start: 0,
                bytes: Arc::from(local),
            },
            Segment {
                start: generated_offset - descriptor_len,
                bytes: Arc::from(descriptor),
            },
            Segment {
                start: generated_offset,
                bytes: Arc::from(tail),
            },
        ],
        std::iter::once(
            generated_offset - descriptor_len - payload_len..generated_offset - descriptor_len,
        )
        .collect(),
    );

    SparseZipFixture {
        source,
        source_length,
        payload_start: local_prefix_len,
        payload_end: generated_offset - descriptor_len,
        central_start: generated_offset,
        central_bytes,
        tail_extensible_data: ZIP64_EXTENSIBLE_DATA.to_vec(),
        comment: ARCHIVE_COMMENT.to_vec(),
    }
}

fn read_range(reader: &SparseReader, range: Range<u64>) -> Vec<u8> {
    let length = usize::try_from(range.end - range.start).unwrap();
    let mut bytes = vec![0; length];
    reader.read_exact_at(&mut bytes, range.start).unwrap();
    bytes
}

fn read_entries<R: ReaderAt>(
    archive: &soapberry_zip::ZipArchive<R>,
) -> Vec<(
    Vec<u8>,
    u64,
    bool,
    Option<Range<u64>>,
    soapberry_zip::ZipArchiveEntryWayfinder,
)> {
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let mut entries = archive.entries(&mut buffer);
    let mut result = Vec::new();
    while let Some(record) = entries.next_entry().unwrap() {
        result.push((
            record.file_path().as_ref().to_vec(),
            record.local_header_offset(),
            record.is_zip64(),
            record.zip64_local_header_offset_range(),
            record.wayfinder(),
        ));
    }
    result
}

fn open_sparse_archive(
    reader: SparseReader,
    length: u64,
) -> soapberry_zip::ZipArchive<SparseReader> {
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    ZipLocator::new()
        .locate_in_reader(reader, &mut buffer, length)
        .map_err(|(_, error)| error)
        .unwrap()
}

fn assert_generated_offset_case(expected_offset: u64) {
    let fixture = sparse_zip_fixture(expected_offset);
    let source_snapshot = fixture.source.snapshot();
    let source_archive = open_sparse_archive(fixture.source.clone(), fixture.source_length);
    assert!(source_archive.is_zip64());
    assert_eq!(source_archive.directory_offset(), fixture.central_start);
    assert_eq!(source_archive.entries_hint(), 1);

    let reads_after_open = fixture.source.stats();
    assert!(reads_after_open.calls > 0);
    assert!(reads_after_open.bytes > 0);
    assert!(
        reads_after_open.synthetic_zero_bytes <= MAX_LOCATOR_PROBE,
        "archive location may inspect a bounded tail window, never the giant member"
    );

    let mut index_buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new_with_limits(
        &source_archive,
        &mut index_buffer,
        ArchiveLimits::UNBOUNDED,
    )
    .unwrap();
    assert_eq!(index.entries().len(), 1);
    assert_eq!(index.entries()[0].local_span(), 0..fixture.central_start);
    assert_eq!(
        index.entries()[0].central_record().start,
        fixture.central_start
    );
    assert_eq!(
        index.entries()[0].central_record().end - fixture.central_start,
        fixture.central_bytes.len() as u64
    );
    assert_eq!(
        fixture.source.stats().synthetic_zero_bytes,
        reads_after_open.synthetic_zero_bytes,
        "indexing must remain metadata-only after archive location"
    );

    let generated_local_len = 30_u64 + GENERATED_NAME.len() as u64 + GENERATED_PAYLOAD.len() as u64;
    let mut plan = PreservationPlan::copy_all(&index);
    plan.try_append(RegeneratedEntry::new(
        GENERATED_NAME,
        GENERATED_PAYLOAD.to_vec(),
    ))
    .unwrap();

    let mut sink = RegionSink::new(fixture.payload_start..fixture.payload_end);
    let mut accounting = ZipOperationAccounting::default();
    index
        .write_to_with_accounting(&plan, &mut sink, &mut accounting)
        .unwrap();

    let source_after_write = fixture.source.stats();
    let payload_len = fixture.payload_end - fixture.payload_start;
    assert_eq!(
        source_after_write.synthetic_zero_bytes,
        reads_after_open.synthetic_zero_bytes + payload_len
    );
    assert!(source_after_write.max_request <= MAX_LOCATOR_PROBE as usize);
    assert_eq!(sink.discarded_bytes, payload_len);
    assert!(sink.length > u64::from(u32::MAX));
    assert!(sink.writes > 1);
    assert!(sink.max_write <= COPY_CHUNK_SIZE.max(4096));
    assert!(
        sink.retained_bytes() < 1024 * 1024,
        "region sink must not retain the multi-gigabyte source payload"
    );
    assert_eq!(fixture.source.snapshot(), source_snapshot);

    let expected_raw = fixture.central_start + fixture.central_bytes.len() as u64 - 4
        + 24
        + fixture.tail_extensible_data.len() as u64
        + 12
        + 10
        + fixture.comment.len() as u64;
    assert_eq!(
        accounting.raw_unchanged_source_bytes_accepted(),
        expected_raw,
        "raw accounting must include the untouched sparse local span and framing"
    );
    assert_eq!(
        accounting.stored_payload_bytes_emitted(),
        GENERATED_PAYLOAD.len() as u64
    );
    assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
    assert_eq!(accounting.precompressed_payload_bytes_emitted(), 0);

    let output_reader = sink.into_reader();
    let output_length = output_reader.length;
    let output_archive = open_sparse_archive(output_reader.clone(), output_length);
    assert!(output_archive.is_zip64());
    assert_eq!(output_archive.end_offset(), output_length);
    assert_eq!(
        output_archive.directory_offset(),
        fixture.central_start + generated_local_len
    );
    assert_eq!(output_archive.entries_hint(), 2);

    let entries = read_entries(&output_archive);
    assert_eq!(entries.len(), 2);
    assert_eq!(entries[0].0, SOURCE_NAME);
    assert_eq!(entries[0].1, 0);
    assert_eq!(
        entries[0].3, None,
        "copied source offset remains fixed-width"
    );
    assert_eq!(entries[1].0, GENERATED_NAME.as_bytes());
    assert_eq!(entries[1].1, expected_offset);
    assert_eq!(entries[1].2, expected_offset >= u64::from(u32::MAX));
    if expected_offset >= u64::from(u32::MAX) {
        let range = entries[1].3.clone().unwrap();
        assert_eq!(range.end - range.start, 8);
        assert_eq!(
            read_range(&output_reader, range),
            expected_offset.to_le_bytes().as_slice()
        );
    } else {
        assert_eq!(entries[1].3, None);
    }

    let source_local_prefix = read_range(&fixture.source, 0..fixture.payload_start);
    let output_local_prefix = read_range(&output_reader, 0..fixture.payload_start);
    assert_eq!(output_local_prefix, source_local_prefix);
    let source_descriptor = read_range(&fixture.source, fixture.payload_end..fixture.central_start);
    let output_descriptor = read_range(&output_reader, fixture.payload_end..fixture.central_start);
    assert_eq!(output_descriptor, source_descriptor);
    assert_eq!(
        read_range(
            &output_reader,
            fixture.central_start + generated_local_len
                ..fixture.central_start + generated_local_len + fixture.central_bytes.len() as u64,
        ),
        fixture.central_bytes,
        "the copied central record must retain every source metadata byte"
    );

    let output_zip64_eocd = output_archive.head_eocd_offset();
    assert_eq!(
        read_range(
            &output_reader,
            output_zip64_eocd + 56..output_archive.eocd_offset() - 20,
        ),
        fixture.tail_extensible_data
    );
    assert_eq!(
        read_range(
            &output_reader,
            output_archive.eocd_offset() + 22
                ..output_archive.eocd_offset() + 22 + fixture.comment.len() as u64,
        ),
        fixture.comment
    );

    let generated_entry = output_archive.get_entry(entries[1].4).unwrap();
    let mut generated_reader = generated_entry.reader();
    let mut generated_bytes = Vec::new();
    generated_reader.read_to_end(&mut generated_bytes).unwrap();
    assert_eq!(generated_bytes, GENERATED_PAYLOAD);
}

#[test]
fn generated_store_offset_promotes_at_each_zip32_boundary() {
    for expected_offset in [
        u64::from(u32::MAX) - 1,
        u64::from(u32::MAX),
        u64::from(u32::MAX) + 1,
    ] {
        assert_generated_offset_case(expected_offset);
    }
}

#[test]
fn zero_run_crc_composition_matches_direct_reference() {
    for length in [0_u64, 1, 2, 7, 32, 4096] {
        let direct = vec![0_u8; length as usize];
        assert_eq!(crc32_zeroes(length), soapberry_zip::crc32(&direct));
    }
}

#[test]
fn unsupported_generated_compression_refuses_before_sparse_output() {
    let fixture = sparse_zip_fixture(u64::from(u32::MAX) + 1);
    let source_archive = open_sparse_archive(fixture.source.clone(), fixture.source_length);
    let mut index_buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new_with_limits(
        &source_archive,
        &mut index_buffer,
        ArchiveLimits::UNBOUNDED,
    )
    .unwrap();

    let mut plan = PreservationPlan::copy_all(&index);
    plan.try_append(
        RegeneratedEntry::new("unsupported.bin", Vec::new())
            .compression_method(CompressionMethod::Bzip2),
    )
    .unwrap();

    let mut sink = RegionSink::new(fixture.payload_start..fixture.payload_end);
    let error = index.write_to(&plan, &mut sink).unwrap_err();
    assert!(matches!(
        error.kind(),
        ErrorKind::UnsupportedPreservation { .. }
    ));
    assert_eq!(sink.length, 0);
    assert_eq!(sink.writes, 0);
    assert!(fixture.source.stats().synthetic_zero_bytes <= MAX_LOCATOR_PROBE);
}
