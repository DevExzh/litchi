#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The fixtures below use fixed, bounded test buffers."
)]

//! One data-descriptor read per verified entry read.
//!
//! `Read::read_to_end` must observe one `Ok(0)` after a member's payload is
//! complete. The verifier's completion branch fires on both the read that
//! completes the payload and that terminating read, so without memoization it
//! resolves the data descriptor twice at the same offset for one member read.
//! These tests count positional reads through a wrapper to pin the behaviour.

use soapberry_zip::{
    CompressionMethod, RECOMMENDED_BUFFER_SIZE, ReaderAt, ZipArchive, ZipArchiveEntryWayfinder,
    ZipArchiveWriter, ZipLocator,
};
use std::io::{self, Cursor, Read, Write};
use std::sync::Mutex;

/// Records every positional read so a test can count exact repeats.
#[derive(Debug)]
struct CountingReaderAt {
    bytes: Vec<u8>,
    reads: Mutex<Vec<(u64, usize)>>,
}

impl CountingReaderAt {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: Mutex::new(Vec::new()),
        }
    }

    fn reads(&self) -> Vec<(u64, usize)> {
        self.reads.lock().unwrap().clone()
    }

    fn clear(&self) {
        self.reads.lock().unwrap().clear();
    }

    /// Counts payload-bearing reads that start at `offset`. A zero-length read
    /// is a reader terminating, not a fetch, so it is excluded.
    fn reads_at(&self, offset: u64) -> usize {
        self.reads()
            .iter()
            .filter(|(start, length)| *start == offset && *length > 0)
            .count()
    }

    fn immediate_duplicate_count(&self) -> usize {
        let reads = self.reads();
        reads
            .windows(2)
            .filter(|window| window[0] == window[1])
            .count()
    }
}

impl ReaderAt for CountingReaderAt {
    fn read_at(&self, buf: &mut [u8], offset: u64) -> io::Result<usize> {
        self.reads.lock().unwrap().push((offset, buf.len()));
        let start = usize::try_from(offset).unwrap_or(self.bytes.len());
        if start >= self.bytes.len() || buf.is_empty() {
            return Ok(0);
        }
        let count = buf.len().min(self.bytes.len() - start);
        buf[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }
}

/// Writes one Deflate member through the streaming path, which frames it with
/// a data descriptor, and returns the archive bytes.
fn deflate_archive_with_descriptor(name: &str, payload: &[u8]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::new(&mut output);
    let (mut entry, config) = archive
        .new_file(name)
        .compression_method(CompressionMethod::Deflate)
        .start()
        .unwrap();
    let encoder = flate2::write::DeflateEncoder::new(&mut entry, flate2::Compression::default());
    let mut data_writer = config.wrap(encoder);
    data_writer.write_all(payload).unwrap();
    let (encoder, descriptor) = data_writer.finish().unwrap();
    encoder.finish().unwrap();
    entry.finish(descriptor).unwrap();
    archive.finish().unwrap();
    output.into_inner()
}

fn open(raw: Vec<u8>) -> ZipArchive<CountingReaderAt> {
    let length = raw.len() as u64;
    let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
    ZipLocator::new()
        .locate_in_reader(CountingReaderAt::new(raw), &mut buffer, length)
        .map_err(|(_, error)| error)
        .unwrap()
}

fn only_entry<R: ReaderAt>(archive: &ZipArchive<R>) -> (ZipArchiveEntryWayfinder, u64, bool) {
    let mut buffer = vec![0_u8; RECOMMENDED_BUFFER_SIZE];
    let mut entries = archive.entries(&mut buffer);
    let record = entries.next_entry().unwrap().unwrap();
    let wayfinder = record.wayfinder();
    let compressed = record.compressed_size_hint();
    let has_descriptor = record.has_data_descriptor();
    assert!(entries.next_entry().unwrap().is_none());
    (wayfinder, compressed, has_descriptor)
}

fn read_verified<R: ReaderAt>(
    archive: &ZipArchive<R>,
    wayfinder: ZipArchiveEntryWayfinder,
) -> Vec<u8> {
    let entry = archive.get_entry(wayfinder).unwrap();
    let decoder = flate2::read::DeflateDecoder::new(entry.reader());
    let mut verifier = entry.verifying_reader(decoder);
    let mut output = Vec::new();
    verifier.read_to_end(&mut output).unwrap();
    output
}

#[test]
fn one_verified_read_resolves_the_data_descriptor_once() {
    let name = "payload.bin";
    let payload: Vec<u8> = (0..4096_u32).map(|index| (index % 251) as u8).collect();
    let raw = deflate_archive_with_descriptor(name, &payload);

    let archive = open(raw.clone());
    let (wayfinder, compressed, has_descriptor) = only_entry(&archive);
    assert!(has_descriptor, "the streaming writer frames a descriptor");

    let descriptor_offset = 30 + name.len() as u64 + compressed;
    assert_eq!(
        &raw[descriptor_offset as usize..descriptor_offset as usize + 4],
        &0x0807_4b50_u32.to_le_bytes(),
        "the computed offset must name the descriptor signature"
    );

    archive.get_ref().clear();
    assert_eq!(read_verified(&archive, wayfinder), payload);

    assert_eq!(
        archive.get_ref().reads_at(descriptor_offset),
        1,
        "one verified read must resolve the descriptor once, not once per \
         completion-branch entry; reads were {:?}",
        archive.get_ref().reads()
    );
    assert_eq!(
        archive.get_ref().immediate_duplicate_count(),
        0,
        "no positional read may repeat the range that immediately preceded it"
    );
}

#[test]
fn repeated_verified_reads_each_resolve_the_descriptor_once() {
    let name = "payload.bin";
    let payload: Vec<u8> = (0..1024_u32).map(|index| (index % 97) as u8).collect();
    let raw = deflate_archive_with_descriptor(name, &payload);

    let archive = open(raw.clone());
    let (wayfinder, compressed, _) = only_entry(&archive);
    let descriptor_offset = 30 + name.len() as u64 + compressed;

    archive.get_ref().clear();
    for _ in 0..3 {
        assert_eq!(read_verified(&archive, wayfinder), payload);
    }
    assert_eq!(
        archive.get_ref().reads_at(descriptor_offset),
        3,
        "each verified read resolves the descriptor exactly once"
    );
}

#[test]
fn a_corrupt_descriptor_still_fails_the_verified_read() {
    let name = "payload.bin";
    let payload: Vec<u8> = (0..512_u32).map(|index| (index % 31) as u8).collect();
    let mut raw = deflate_archive_with_descriptor(name, &payload);

    let compressed = {
        let archive = open(raw.clone());
        only_entry(&archive).1
    };
    let descriptor_offset = (30 + name.len() as u64 + compressed) as usize;
    // Corrupt the descriptor's CRC field, which follows its signature.
    raw[descriptor_offset + 4] ^= 0xFF;

    let archive = open(raw);
    let (wayfinder, _, _) = only_entry(&archive);
    let entry = archive.get_entry(wayfinder).unwrap();
    let decoder = flate2::read::DeflateDecoder::new(entry.reader());
    let mut verifier = entry.verifying_reader(decoder);
    let mut output = Vec::new();
    let error = verifier
        .read_to_end(&mut output)
        .expect_err("a corrupt descriptor must fail the verified read");
    assert_eq!(error.kind(), io::ErrorKind::InvalidData);
}
