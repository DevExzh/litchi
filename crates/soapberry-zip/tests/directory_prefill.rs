#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "bounded, deterministic generated ZIP fixtures and explicit index arithmetic"
)]

//! Change 0632: one bounded read of the head of the central directory serves
//! both the locator's first-central-record probe and the index's
//! central-directory scan.
//!
//! These tests pin what that read is allowed to change — the number of
//! positional requests an index build issues and the size of the buffer those
//! bytes land in — and what it is not allowed to change: which members are
//! found, in which order, with which metadata, and which typed error a
//! malformed directory reaches.

use std::cell::RefCell;
use std::io;

use soapberry_zip::office::{ArchiveLimits, IndexedArchive, StreamingArchiveWriter};
use soapberry_zip::{ErrorKind, RECOMMENDED_BUFFER_SIZE, ReaderAt, ZipArchive, ZipLocator};

const ZIP64_CENTRAL_LOCAL: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/zip64-central-local-signed.zip"
);
const ZIP64_MANY_SMALL: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/many-small-central-local-signed.zip"
);

/// A positional source that records every request it is asked to serve.
#[derive(Debug)]
struct Counting<'a> {
    bytes: &'a [u8],
    /// `(offset, requested, returned)` for every `read_at`.
    log: RefCell<Vec<(u64, usize, usize)>>,
    /// Maximum bytes any single request may return, to force short reads.
    cap: usize,
}

impl<'a> Counting<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self {
            bytes,
            log: RefCell::new(Vec::new()),
            cap: usize::MAX,
        }
    }

    fn capped(bytes: &'a [u8], cap: usize) -> Self {
        Self {
            bytes,
            log: RefCell::new(Vec::new()),
            cap,
        }
    }

    fn requests(&self) -> usize {
        self.log.borrow().len()
    }

    fn bytes_returned(&self) -> usize {
        self.log.borrow().iter().map(|entry| entry.2).sum()
    }

    fn log(&self) -> Vec<(u64, usize, usize)> {
        self.log.borrow().clone()
    }
}

impl ReaderAt for Counting<'_> {
    fn read_at(&self, output: &mut [u8], offset: u64) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        let remaining = self.bytes.get(start..).unwrap_or_default();
        let count = remaining.len().min(output.len()).min(self.cap);
        output[..count].copy_from_slice(&remaining[..count]);
        self.log.borrow_mut().push((offset, output.len(), count));
        Ok(count)
    }
}

/// Every member name an index reports, in index order.
fn names<R: ReaderAt>(archive: &IndexedArchive<R>) -> Vec<String> {
    let mut found: Vec<String> = archive
        .file_names()
        .map(std::string::ToString::to_string)
        .collect();
    found.sort();
    found
}

fn stored_archive(count: usize) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    for index in 0..count {
        writer
            .write_stored(&format!("member-{index:08}"), b"payload")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn declared_directory_size(bytes: &[u8]) -> u64 {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == [0x50, 0x4b, 0x05, 0x06])
        .expect("an EOCD signature");
    u64::from(u32::from_le_bytes(
        bytes[eocd + 12..eocd + 16].try_into().unwrap(),
    ))
}

/// Append a ZIP archive comment, rewriting the EOCD comment length.
fn with_comment(bytes: &[u8], comment: &[u8]) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == [0x50, 0x4b, 0x05, 0x06])
        .expect("an EOCD signature");
    let mut out = bytes.to_vec();
    let len = u16::try_from(comment.len()).expect("a bounded comment");
    out[eocd + 20..eocd + 22].copy_from_slice(&len.to_le_bytes());
    out.extend_from_slice(comment);
    out
}

/// Prepend `prefix` bytes, so the declared directory offset is relative to a
/// ZIP that no longer starts at zero. This is the shape
/// `finish_locate_in_reader`'s base-offset fallback exists for.
fn with_prefix(bytes: &[u8], prefix: usize) -> Vec<u8> {
    let mut out = vec![0xA5_u8; prefix];
    out.extend_from_slice(bytes);
    out
}

#[test]
fn one_request_serves_the_probe_and_the_whole_directory_scan() {
    let bytes = stored_archive(8);
    let source = Counting::new(&bytes);
    let archive =
        IndexedArchive::from_reader(&source, bytes.len() as u64).expect("the archive opens");

    assert_eq!(archive.len(), 8);
    let log = source.log();
    assert_eq!(
        log.len(),
        2,
        "the locate costs one EOCD probe and one directory read: {log:?}"
    );
    // The EOCD probe, at `len - 22`.
    assert_eq!(log[0].0, bytes.len() as u64 - 22);
    assert_eq!(log[0].1, 22);
    // One read of the whole declared directory, at the directory offset.
    let located = ZipArchive::from_slice(bytes.as_slice()).unwrap();
    assert_eq!(log[1].0, located.directory_offset());
    assert_eq!(log[1].1 as u64, declared_directory_size(&bytes));
    assert_eq!(
        source.bytes_returned() as u64,
        22 + declared_directory_size(&bytes)
    );
}

#[test]
fn the_prefill_finds_exactly_what_a_separate_scan_finds() {
    for count in [0_usize, 1, 2, 17, 64] {
        let bytes = stored_archive(count);
        let source = Counting::new(&bytes);
        let indexed =
            IndexedArchive::from_reader(&source, bytes.len() as u64).expect("the archive opens");

        // The same archive scanned through the untouched public entry
        // iterator, which takes no prefill.
        let located = ZipLocator::new()
            .locate_in_reader(&source, &mut [0_u8; 512], bytes.len() as u64)
            .map_err(|(_, error)| error)
            .expect("the archive locates");
        let mut scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let mut entries = located.entries(&mut scratch);
        let mut expected = Vec::new();
        while let Some(entry) = entries.next_entry().expect("a well-formed directory") {
            expected.push(String::from_utf8_lossy(entry.file_path().as_ref()).into_owned());
        }
        expected.sort();

        assert_eq!(names(&indexed), expected, "count={count}");
    }
}

#[test]
fn a_directory_larger_than_the_window_is_read_in_two_requests() {
    // 1,600 members at 46 + 17 bytes of central record each is about 100 KiB
    // of central directory: more than the 64 KiB prefill window, so the
    // window covers its head and the scan reads the tail for itself.
    let bytes = stored_archive(1_600);
    let directory = declared_directory_size(&bytes);
    assert!(
        directory > RECOMMENDED_BUFFER_SIZE as u64,
        "the fixture must exceed the window: {directory}"
    );

    let source = Counting::new(&bytes);
    let archive =
        IndexedArchive::from_reader(&source, bytes.len() as u64).expect("the archive opens");
    assert_eq!(archive.len(), 1_600);
    assert_eq!(
        names(&archive).first().map(String::as_str),
        Some("member-00000000")
    );

    let log = source.log();
    // One EOCD probe, one 64 KiB window, then refills for the remainder.
    assert_eq!(log[0].1, 22);
    assert_eq!(log[1].1, RECOMMENDED_BUFFER_SIZE);
    assert_eq!(
        log.len(),
        3,
        "the EOCD probe, the window, and one read for the rest: {log:?}"
    );
    // Every read after the window continues forward inside the directory and
    // no byte of the directory is read twice.
    let start = log[1].0;
    let mut cursor = start + RECOMMENDED_BUFFER_SIZE as u64;
    for entry in &log[2..] {
        assert_eq!(entry.0, cursor, "reads must not overlap or skip: {log:?}");
        cursor += entry.2 as u64;
    }
    assert_eq!(cursor, start + directory);
}

#[test]
fn a_zip64_archive_keeps_its_records_and_its_members() {
    for (label, bytes) in [
        ("zip64-central-local-signed", ZIP64_CENTRAL_LOCAL),
        ("many-small-central-local-signed", ZIP64_MANY_SMALL),
    ] {
        let source = Counting::new(bytes);
        let archive =
            IndexedArchive::from_reader(&source, bytes.len() as u64).expect("the archive opens");
        assert!(archive.archive_is_zip64(), "{label} is a ZIP64 archive");

        let mut scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let located = ZipLocator::new()
            .locate_in_reader(bytes, &mut scratch, bytes.len() as u64)
            .map_err(|(_, error)| error)
            .expect("the archive locates");
        let mut scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let mut entries = located.entries(&mut scratch);
        let mut expected = Vec::new();
        while let Some(entry) = entries.next_entry().expect("a well-formed directory") {
            expected.push(String::from_utf8_lossy(entry.file_path().as_ref()).into_owned());
        }
        expected.sort();
        assert_eq!(names(&archive), expected, "{label}");
        assert!(source.requests() >= 2, "{label}: {:?}", source.log());
    }
}

#[test]
fn a_comment_bearing_archive_keeps_the_backwards_search_and_its_members() {
    let plain = stored_archive(12);
    let commented = with_comment(&plain, b"a deliberate ZIP archive comment");

    let source = Counting::new(&commented);
    let archive =
        IndexedArchive::from_reader(&source, commented.len() as u64).expect("the archive opens");
    let plain_source = Counting::new(&plain);
    let plain_archive =
        IndexedArchive::from_reader(&plain_source, plain.len() as u64).expect("the archive opens");
    assert_eq!(names(&archive), names(&plain_archive));

    // The comment pushes the EOCD off `len - 22`, so the locator falls back to
    // its backwards search. The directory is still read exactly once.
    let log = source.log();
    let directory_offset = ZipArchive::from_slice(commented.as_slice())
        .unwrap()
        .directory_offset();
    let directory_reads = log
        .iter()
        .filter(|entry| entry.0 == directory_offset)
        .count();
    assert_eq!(directory_reads, 1, "{log:?}");
    assert_eq!(
        log.iter()
            .filter(|entry| entry.0 == directory_offset && entry.1 == 46)
            .count(),
        0,
        "the separate 46-byte central-record probe is gone: {log:?}"
    );
}

#[test]
fn the_stack_probe_keeps_the_search_window_only_on_a_missed_probe() {
    let plain = stored_archive(1_600);
    assert!(declared_directory_size(&plain) > RECOMMENDED_BUFFER_SIZE as u64);

    let plain_source = Counting::new(&plain);
    let plain_archive =
        IndexedArchive::from_reader(&plain_source, plain.len() as u64).expect("plain opens");
    assert_eq!(plain_archive.len(), 1_600);
    let plain_log = plain_source.log();
    assert_eq!(plain_log[0].1, 22, "the fast path probes only the EOCD");
    assert_eq!(plain_log[1].1, RECOMMENDED_BUFFER_SIZE);

    let commented = with_comment(&plain, b"a deliberate ZIP archive comment");
    let comment_source = Counting::new(&commented);
    let comment_archive = IndexedArchive::from_reader(&comment_source, commented.len() as u64)
        .expect("commented archive opens");
    assert_eq!(comment_archive.len(), 1_600);
    let comment_log = comment_source.log();
    assert_eq!(comment_log[0].1, 22, "the missed fast probe is preserved");
    assert_eq!(
        comment_log[1].1, RECOMMENDED_BUFFER_SIZE,
        "the backwards search keeps the historical bounded window"
    );
}

#[test]
fn a_comment_that_contains_an_eocd_signature_resolves_where_it_always_did() {
    // A false EOCD signature inside the comment is the case the locator's
    // backwards search documents: it finds the *last* signature. The prefill
    // must not move which one that is.
    let plain = stored_archive(6);
    let mut comment = b"junk".to_vec();
    comment.extend_from_slice(&[0x50, 0x4b, 0x05, 0x06]);
    comment.extend_from_slice(&[0_u8; 18]);
    let commented = with_comment(&plain, &comment);

    let source = Counting::new(&commented);
    let archive =
        IndexedArchive::from_reader(&source, commented.len() as u64).expect("the archive opens");

    // The untouched public path, which takes no prefill, over the same bytes.
    let mut scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
    let located = ZipLocator::new()
        .locate_in_reader(commented.as_slice(), &mut scratch, commented.len() as u64)
        .map_err(|(_, error)| error)
        .expect("the archive locates");
    assert_eq!(archive.archive_end_offset(), located.end_offset());

    let mut scan_scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
    let mut entries = located.entries(&mut scan_scratch);
    let mut expected = Vec::new();
    while let Some(entry) = entries.next_entry().expect("a well-formed directory") {
        expected.push(String::from_utf8_lossy(entry.file_path().as_ref()).into_owned());
    }
    expected.sort();
    assert_eq!(names(&archive), expected);
}

#[test]
fn a_prefixed_archive_still_reaches_the_base_offset_fallback() {
    let plain = stored_archive(9);
    let prefixed = with_prefix(&plain, 4_096);

    let source = Counting::new(&prefixed);
    let archive =
        IndexedArchive::from_reader(&source, prefixed.len() as u64).expect("the archive opens");
    assert_eq!(archive.len(), 9);

    let plain_source = Counting::new(&plain);
    let plain_archive =
        IndexedArchive::from_reader(&plain_source, plain.len() as u64).expect("the archive opens");
    assert_eq!(names(&archive), names(&plain_archive));

    // The declared offset misses, so two directory-head reads happen: one at
    // the declared offset and one at `eocd - central_dir_size`. The scan adds
    // none of its own.
    let log = source.log();
    let declared = ZipArchive::from_slice(plain.as_slice())
        .unwrap()
        .directory_offset();
    assert!(
        log.iter().any(|entry| entry.0 == declared),
        "the declared offset is probed first: {log:?}"
    );
    assert!(
        log.iter().any(|entry| entry.0 == declared + 4_096),
        "the fallback offset is probed second: {log:?}"
    );
    assert_eq!(log.len(), 3, "{log:?}");
}

#[test]
fn a_short_reading_source_is_served_the_same_members() {
    let bytes = stored_archive(24);
    let expected = {
        let source = Counting::new(&bytes);
        names(&IndexedArchive::from_reader(&source, bytes.len() as u64).expect("opens"))
    };
    for cap in [1_usize, 7, 46, 64, 97, 512] {
        let source = Counting::capped(&bytes, cap);
        let archive = IndexedArchive::from_reader(&source, bytes.len() as u64)
            .unwrap_or_else(|error| panic!("cap={cap}: {error}"));
        assert_eq!(names(&archive), expected, "cap={cap}");
    }
}

#[test]
fn a_directory_shorter_than_one_central_record_takes_no_prefill() {
    let bytes = stored_archive(3);
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == [0x50, 0x4b, 0x05, 0x06])
        .unwrap();
    // Declare a 20-byte directory holding no records: shorter than one fixed
    // central record, so no prefill can be taken and the 46-byte probe runs
    // exactly as it always has.
    let mut truncated = bytes.clone();
    truncated[eocd + 8..eocd + 10].copy_from_slice(&0_u16.to_le_bytes());
    truncated[eocd + 10..eocd + 12].copy_from_slice(&0_u16.to_le_bytes());
    truncated[eocd + 12..eocd + 16].copy_from_slice(&20_u32.to_le_bytes());
    let source = Counting::new(&truncated);
    let error = IndexedArchive::from_reader(&source, truncated.len() as u64)
        .expect_err("a 20-byte directory cannot hold a record");
    assert!(
        matches!(error.kind(), ErrorKind::Eof),
        "unexpected refusal: {:?}",
        error.kind()
    );

    let directory_offset = ZipLocator::new()
        .locate_in_reader(
            truncated.as_slice(),
            &mut [0_u8; 512],
            truncated.len() as u64,
        )
        .map_err(|(_, error)| error)
        .expect("the archive locates")
        .directory_offset();
    let log = source.log();
    assert!(
        log.iter()
            .all(|entry| entry.0 != directory_offset || entry.1 <= 46),
        "no prefill may be taken for a sub-record directory: {log:?}"
    );
}

#[test]
fn an_oversized_variable_field_keeps_its_typed_refusal() {
    // A record whose name length runs past the end of the declared directory.
    // The scan buffer is now sized to the directory rather than to the 64 KiB
    // recommendation, so the spill threshold is pinned separately; without
    // that pin this record would move from `BufferTooSmall` to `Eof`.
    let bytes = stored_archive(3);
    let located = ZipArchive::from_slice(bytes.as_slice()).unwrap();
    let start = usize::try_from(located.directory_offset()).unwrap();
    let directory = usize::try_from(declared_directory_size(&bytes)).unwrap();
    assert!(directory < RECOMMENDED_BUFFER_SIZE);

    let mut widened = bytes.clone();
    // Name length 4,000: larger than the whole declared directory, smaller
    // than the 64 KiB recommendation.
    widened[start + 28..start + 30].copy_from_slice(&4_000_u16.to_le_bytes());

    // What the untouched public iterator with a 64 KiB buffer reports.
    let expected = {
        let mut locate_scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let located = ZipLocator::new()
            .locate_in_reader(
                widened.as_slice(),
                &mut locate_scratch,
                widened.len() as u64,
            )
            .map_err(|(_, error)| error)
            .expect("the archive locates");
        let mut scratch = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
        let mut entries = located.entries(&mut scratch);
        format!("{:?}", entries.next_entry().expect_err("must fail").kind())
    };

    let source = Counting::new(&widened);
    let error = IndexedArchive::from_reader_with_limits(
        &source,
        widened.len() as u64,
        ArchiveLimits::default(),
    )
    .expect_err("an oversized variable field must fail closed");
    assert_eq!(format!("{:?}", error.kind()), expected);
    assert!(
        matches!(error.kind(), ErrorKind::BufferTooSmall),
        "unexpected refusal: {:?}",
        error.kind()
    );
}

#[test]
fn an_empty_archive_takes_no_prefill_and_opens() {
    let bytes = stored_archive(0);
    assert_eq!(declared_directory_size(&bytes), 0);
    let source = Counting::new(&bytes);
    let archive =
        IndexedArchive::from_reader(&source, bytes.len() as u64).expect("an empty archive opens");
    assert_eq!(archive.len(), 0);
    // Only the EOCD probe and the 46-byte central-record probe, which cannot
    // be satisfied and is not expected to be.
    let log = source.log();
    assert_eq!(log[0].1, 22, "{log:?}");
    assert!(log.iter().all(|entry| entry.1 <= 46), "{log:?}");
}
