#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The fixtures below use fixed, bounded test buffers."
)]

//! One bounded positional read per indexed member first-read.
//!
//! `IndexedArchive::read`/`read_entry` used to cost two or three positional
//! requests for a member's first read: the 30-byte fixed local header, the
//! payload, and a data descriptor when one is declared. One read of the
//! member's whole local record replaces them. These tests count the requests a
//! source actually receives, and pin the properties that make the substitution
//! invisible to a caller: identical bytes, identical errors, independence of
//! read order, and an untouched grammar for every member the window does not
//! admit.

use soapberry_zip::extra_fields::ExtraFieldId;
use soapberry_zip::office::{ArchiveLimits, IndexedArchive};
use soapberry_zip::{CompressionMethod, Header, ReaderAt, ZipArchiveWriter};
use std::io::{self, Cursor, Write};
use std::sync::{Arc, Mutex};

/// Records every positional read a source receives.
#[derive(Debug)]
struct CountingReaderAt {
    bytes: Vec<u8>,
    /// When set, every read longer than this fails instead of being served.
    read_ceiling: Option<usize>,
    /// When set, no read returns more than this many bytes.
    short_reads: Option<usize>,
    reads: Mutex<Vec<(u64, usize)>>,
}

impl CountingReaderAt {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            read_ceiling: None,
            short_reads: None,
            reads: Mutex::new(Vec::new()),
        }
    }

    fn with_read_ceiling(bytes: Vec<u8>, ceiling: usize) -> Self {
        Self {
            bytes,
            read_ceiling: Some(ceiling),
            short_reads: None,
            reads: Mutex::new(Vec::new()),
        }
    }

    fn with_short_reads(bytes: Vec<u8>, most: usize) -> Self {
        Self {
            bytes,
            read_ceiling: None,
            short_reads: Some(most),
            reads: Mutex::new(Vec::new()),
        }
    }

    fn reads(&self) -> Vec<(u64, usize)> {
        self.reads.lock().unwrap().clone()
    }

    fn clear(&self) {
        self.reads.lock().unwrap().clear();
    }

    /// Reads that actually fetched bytes. A zero-length read is a reader
    /// terminating, not a request.
    fn fetches(&self) -> usize {
        self.reads()
            .iter()
            .filter(|(_, length)| *length > 0)
            .count()
    }

    fn fetched_bytes(&self) -> usize {
        self.reads().iter().map(|(_, length)| *length).sum()
    }
}

impl ReaderAt for CountingReaderAt {
    fn read_at(&self, buf: &mut [u8], offset: u64) -> io::Result<usize> {
        self.reads.lock().unwrap().push((offset, buf.len()));
        if let Some(ceiling) = self.read_ceiling {
            if buf.len() > ceiling {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "read above the source's ceiling",
                ));
            }
        }
        let start = usize::try_from(offset).unwrap_or(self.bytes.len());
        if start >= self.bytes.len() || buf.is_empty() {
            return Ok(0);
        }
        let mut count = buf.len().min(self.bytes.len() - start);
        if let Some(most) = self.short_reads {
            count = count.min(most);
        }
        buf[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }
}

/// A member of a fixture archive.
struct Member {
    name: &'static str,
    payload: Vec<u8>,
    method: CompressionMethod,
    /// Bytes of local extra field to pad the record with.
    extra: usize,
}

impl Member {
    fn stored(name: &'static str, payload: Vec<u8>) -> Self {
        Self {
            name,
            payload,
            method: CompressionMethod::Store,
            extra: 0,
        }
    }

    fn deflated(name: &'static str, payload: Vec<u8>) -> Self {
        Self {
            name,
            payload,
            method: CompressionMethod::Deflate,
            extra: 0,
        }
    }

    fn with_extra(mut self, extra: usize) -> Self {
        self.extra = extra;
        self
    }
}

/// Writes a seekable archive, which frames every member without a data
/// descriptor.
fn seekable_archive(members: &[Member]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::new(&mut output);
    for member in members {
        let mut builder = archive
            .new_file(member.name)
            .compression_method(member.method);
        if member.extra != 0 {
            builder = builder
                .extra_field(
                    ExtraFieldId::MICROSOFT_OPEN_PACKAGING_GROWTH_HINT,
                    &vec![0x5a_u8; member.extra],
                    Header::LOCAL,
                )
                .unwrap();
        }
        let (mut entry, config) = builder.start().unwrap();
        match member.method {
            CompressionMethod::Deflate => {
                let encoder =
                    flate2::write::DeflateEncoder::new(&mut entry, flate2::Compression::default());
                let mut data_writer = config.wrap(encoder);
                data_writer.write_all(&member.payload).unwrap();
                let (encoder, descriptor) = data_writer.finish().unwrap();
                encoder.finish().unwrap();
                entry.finish(descriptor).unwrap();
            },
            _ => {
                let mut data_writer = config.wrap(&mut entry);
                data_writer.write_all(&member.payload).unwrap();
                let (_, descriptor) = data_writer.finish().unwrap();
                entry.finish(descriptor).unwrap();
            },
        }
    }
    archive.finish().unwrap();
    output.into_inner()
}

/// Writes a streaming archive, which frames every Deflate member with a data
/// descriptor.
fn streaming_archive(members: &[Member]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut archive = ZipArchiveWriter::new(&mut output);
    for member in members {
        let (mut entry, config) = archive
            .new_file(member.name)
            .compression_method(CompressionMethod::Deflate)
            .start()
            .unwrap();
        let encoder =
            flate2::write::DeflateEncoder::new(&mut entry, flate2::Compression::default());
        let mut data_writer = config.wrap(encoder);
        data_writer.write_all(&member.payload).unwrap();
        let (encoder, descriptor) = data_writer.finish().unwrap();
        encoder.finish().unwrap();
        entry.finish(descriptor).unwrap();
    }
    archive.finish().unwrap();
    output.into_inner()
}

type Opened = (IndexedArchive<Arc<CountingReaderAt>>, Arc<CountingReaderAt>);

fn open_source(source: CountingReaderAt) -> Opened {
    let length = source.bytes.len() as u64;
    let source = Arc::new(source);
    let archive = IndexedArchive::from_reader_with_limits(
        Arc::clone(&source),
        length,
        ArchiveLimits::default(),
    )
    .unwrap();
    (archive, source)
}

fn open(raw: Vec<u8>) -> Opened {
    open_source(CountingReaderAt::new(raw))
}

/// The located central-directory offset, read straight out of the archive's
/// end-of-central-directory record.
fn directory_offset(raw: &[u8]) -> u64 {
    let eocd = raw
        .windows(4)
        .enumerate()
        .filter(|(_, window)| *window == [0x50, 0x4b, 0x05, 0x06])
        .map(|(index, _)| index)
        .next_back()
        .expect("fixture archives carry an end-of-central-directory record");
    u64::from(u32::from_le_bytes(
        raw[eocd + 16..eocd + 20].try_into().unwrap(),
    ))
}

fn compressible(seed: u8, length: usize) -> Vec<u8> {
    (0..length)
        .map(|index| b"abcdefghij"[(index + usize::from(seed)) % 10])
        .collect()
}

#[test]
fn a_deflate_member_first_read_is_one_request() {
    let raw = seekable_archive(&[
        Member::deflated("word/document.xml", compressible(0, 4_096)),
        Member::deflated("word/styles.xml", compressible(1, 2_048)),
    ]);
    let (archive, source) = open(raw);
    source.clear();

    assert_eq!(
        archive.read("word/document.xml").unwrap(),
        compressible(0, 4_096)
    );
    assert_eq!(source.fetches(), 1, "reads were {:?}", source.reads());

    source.clear();
    assert_eq!(
        archive.read("word/styles.xml").unwrap(),
        compressible(1, 2_048)
    );
    assert_eq!(source.fetches(), 1, "reads were {:?}", source.reads());
}

#[test]
fn a_stored_member_first_read_is_one_request() {
    let raw = seekable_archive(&[Member::stored("mimetype", b"application/zip".to_vec())]);
    let (archive, source) = open(raw);
    source.clear();

    assert_eq!(archive.read("mimetype").unwrap(), b"application/zip");
    assert_eq!(source.fetches(), 1, "reads were {:?}", source.reads());
}

#[test]
fn a_descriptor_bearing_member_first_read_is_one_request() {
    let raw = streaming_archive(&[
        Member::deflated("a.xml", compressible(2, 1_024)),
        Member::deflated("b.xml", compressible(3, 1_024)),
    ]);
    let (archive, source) = open(raw);
    assert!(
        archive.has_data_descriptor_entries(),
        "the streaming writer must frame these members with data descriptors"
    );

    for (name, seed) in [("a.xml", 2_u8), ("b.xml", 3)] {
        source.clear();
        assert_eq!(archive.read(name).unwrap(), compressible(seed, 1_024));
        assert_eq!(
            source.fetches(),
            1,
            "{name} reads were {:?}",
            source.reads()
        );
    }
}

#[test]
fn every_member_of_one_archive_costs_one_request() {
    let members: Vec<Member> = (0..8)
        .map(|index| {
            let name: &'static str = Box::leak(format!("part{index}.xml").into_boxed_str());
            Member::deflated(name, compressible(index as u8, 512 + index * 64))
        })
        .collect();
    let raw = seekable_archive(&members);
    let (archive, source) = open(raw);
    source.clear();

    for member in &members {
        assert_eq!(archive.read(member.name).unwrap(), member.payload);
    }
    assert_eq!(
        source.fetches(),
        members.len(),
        "reads were {:?}",
        source.reads()
    );
}

#[test]
fn member_reads_are_order_independent() {
    let members: Vec<Member> = (0..6)
        .map(|index| {
            let name: &'static str = Box::leak(format!("order{index}.xml").into_boxed_str());
            Member::deflated(name, compressible(index as u8, 700 + index * 33))
        })
        .collect();
    let raw = seekable_archive(&members);

    let forward: Vec<Vec<u8>> = {
        let (archive, _source) = open(raw.clone());
        members
            .iter()
            .map(|member| archive.read(member.name).unwrap())
            .collect()
    };
    let reverse: Vec<Vec<u8>> = {
        let (archive, _source) = open(raw.clone());
        let mut read: Vec<Vec<u8>> = members
            .iter()
            .rev()
            .map(|member| archive.read(member.name).unwrap())
            .collect();
        read.reverse();
        read
    };
    let interleaved: Vec<Vec<u8>> = {
        let (archive, _source) = open(raw.clone());
        let mut read = vec![Vec::new(); members.len()];
        for index in (0..members.len()).step_by(2) {
            read[index] = archive.read(members[index].name).unwrap();
        }
        for index in (1..members.len()).step_by(2) {
            read[index] = archive.read(members[index].name).unwrap();
        }
        read
    };
    // One shared session must agree with a fresh read of each member, and a
    // member read twice must agree with itself.
    let sessioned: Vec<Vec<u8>> = {
        let (archive, _source) = open(raw);
        let mut session = archive.read_session();
        members
            .iter()
            .map(|member| {
                let first = session.read(member.name).unwrap();
                let second = archive.read(member.name).unwrap();
                assert_eq!(first, second);
                second
            })
            .collect()
    };

    assert_eq!(forward, reverse);
    assert_eq!(forward, interleaved);
    assert_eq!(forward, sessioned);
    for (member, bytes) in members.iter().zip(&forward) {
        assert_eq!(bytes, &member.payload);
    }
}

#[test]
fn a_member_above_the_ceiling_keeps_the_historical_grammar() {
    // 96 KiB of incompressible payload is above the 64 KiB span ceiling, so no
    // speculative read is taken and the first request is the 30-byte fixed
    // local header, exactly as before.
    let payload: Vec<u8> = (0..96 * 1024)
        .map(|index: usize| (index.wrapping_mul(2_654_435_761) >> 13) as u8)
        .collect();
    let raw = seekable_archive(&[Member::stored("big.bin", payload.clone())]);
    let (archive, source) = open(raw);
    source.clear();

    assert_eq!(archive.read("big.bin").unwrap(), payload);
    let reads = source.reads();
    assert_eq!(
        reads.first().map(|(_, length)| *length),
        Some(30),
        "reads were {reads:?}"
    );
    assert!(source.fetches() > 1, "reads were {reads:?}");
}

#[test]
fn a_local_variable_region_above_the_allowance_still_reads_the_member() {
    // A 2 KiB local extra field is longer than the 610-byte allowance the
    // declared bound reserves, and it is carried only by the local header, so
    // the central record cannot predict it. The member must still read.
    let payload = compressible(9, 1_500);
    let raw = seekable_archive(&[
        Member::deflated("padded.xml", payload.clone()).with_extra(2_048),
        Member::deflated("plain.xml", compressible(10, 256)),
    ]);
    let (archive, _source) = open(raw);

    assert_eq!(archive.read("padded.xml").unwrap(), payload);
    assert_eq!(archive.read("plain.xml").unwrap(), compressible(10, 256));
}

#[test]
fn a_failing_source_read_is_the_member_read_s_failure_and_is_tried_once() {
    // The span read is the member read's first contact with the source, so a
    // source failure is reported rather than retried behind the caller's back:
    // one member read stays one refusal, one cancellation observation and one
    // resource reservation. A source that refuses every read longer than 64
    // bytes refuses the historical payload read too, so this member is refused
    // either way; what is pinned here is that exactly one read is attempted.
    let payload = compressible(4, 1_024);
    let raw = seekable_archive(&[Member::deflated("refused.xml", payload)]);
    let (archive, source) = open_source(CountingReaderAt::with_read_ceiling(raw, 64));
    source.clear();

    let error = archive.read("refused.xml").unwrap_err();
    assert!(
        format!("{error}").contains("ceiling"),
        "the source's own error must reach the caller, got {error}"
    );
    assert_eq!(
        source.fetches(),
        1,
        "a refused member read must attempt one read, got {:?}",
        source.reads()
    );
}

#[test]
fn a_short_source_read_is_served_from_what_arrived() {
    // A source that answers at most 48 bytes per read never fills the window.
    // The member must still read, byte for byte, from the reads that follow.
    let payload = compressible(12, 900);
    let raw = seekable_archive(&[Member::deflated("short.xml", payload.clone())]);
    let (archive, _source) = open_source(CountingReaderAt::with_short_reads(raw.clone(), 48));

    assert_eq!(archive.read("short.xml").unwrap(), payload);

    let (control, _control_source) = open(raw);
    assert_eq!(
        archive.read("short.xml").unwrap(),
        control.read("short.xml").unwrap()
    );
}

#[test]
fn a_truncated_member_keeps_its_error_identity() {
    let raw = seekable_archive(&[
        Member::deflated("head.xml", compressible(5, 2_048)),
        Member::deflated("tail.xml", compressible(6, 2_048)),
    ]);
    let length = raw.len() as u64;
    // Keep the central directory and the located tail intact, and zero the
    // second member's local record so its payload cannot verify.
    let directory = usize::try_from(directory_offset(&raw)).unwrap();
    let mut damaged = raw.clone();
    let second = directory / 2;
    damaged[second..directory].fill(0);

    let (archive, _source) = {
        let source = Arc::new(CountingReaderAt::new(damaged));
        let archive = IndexedArchive::from_reader_with_limits(
            Arc::clone(&source),
            length,
            ArchiveLimits::default(),
        )
        .unwrap();
        (archive, source)
    };

    let error = archive.read("tail.xml").unwrap_err();
    assert!(
        !format!("{error}").is_empty(),
        "a refused member must carry a message"
    );
    // The undamaged first member still reads.
    assert_eq!(archive.read("head.xml").unwrap(), compressible(5, 2_048));
}

#[test]
fn the_span_read_does_not_reach_past_the_central_directory() {
    let raw = seekable_archive(&[
        Member::deflated("one.xml", compressible(7, 300)),
        Member::deflated("two.xml", compressible(8, 300)),
    ]);
    let directory = directory_offset(&raw);
    let (archive, source) = open(raw);
    source.clear();

    assert_eq!(archive.read("two.xml").unwrap(), compressible(8, 300));
    for (offset, length) in source.reads() {
        let end = offset + length as u64;
        assert!(
            end <= directory,
            "read {offset}..{end} reached past the central directory at {directory}"
        );
    }
}

#[test]
fn one_request_reads_no_more_bytes_than_the_member_record() {
    let payload = compressible(11, 3_000);
    let raw = seekable_archive(&[Member::deflated("exact.xml", payload.clone())]);
    let (archive, source) = open(raw);
    let compressed = archive.metadata("exact.xml").unwrap().compressed_size();
    source.clear();

    assert_eq!(archive.read("exact.xml").unwrap(), payload);
    let fetched = source.fetched_bytes() as u64;
    // The whole local record is 30 bytes of fixed header, the name, the extra
    // field and the payload. The read may not exceed that plus the allowance
    // and the widest descriptor.
    let ceiling = 30 + "exact.xml".len() as u64 + 610 + compressed + 24;
    assert!(
        fetched <= ceiling,
        "one member read fetched {fetched} bytes, above {ceiling}"
    );
}

/// Patch the central record for `name`: set its declared compressed size and
/// its CRC. Both live in the fixed part of the central file header.
fn patch_central_record(raw: &mut [u8], name: &str, compressed: u32, crc: u32) {
    let directory = usize::try_from(directory_offset(raw)).unwrap();
    let mut cursor = directory;
    while cursor + 46 <= raw.len() && raw[cursor..cursor + 4] == [0x50, 0x4b, 0x01, 0x02] {
        let name_len = usize::from(u16::from_le_bytes([raw[cursor + 28], raw[cursor + 29]]));
        let extra_len = usize::from(u16::from_le_bytes([raw[cursor + 30], raw[cursor + 31]]));
        let comment_len = usize::from(u16::from_le_bytes([raw[cursor + 32], raw[cursor + 33]]));
        if &raw[cursor + 46..cursor + 46 + name_len] == name.as_bytes() {
            raw[cursor + 16..cursor + 20].copy_from_slice(&crc.to_le_bytes());
            raw[cursor + 20..cursor + 24].copy_from_slice(&compressed.to_le_bytes());
            return;
        }
        cursor += 46 + name_len + extra_len + comment_len;
    }
    panic!("no central record for {name}");
}

#[test]
fn a_payload_the_window_cannot_cover_is_read_from_the_source() {
    // The central record declares a compressed size that reaches far past the
    // member's own local record, so the bounded window — which stops where the
    // next local header begins — cannot hold the payload the reader will ask
    // for. Serving a prefix of it from the buffer would hand the decoder
    // different-sized pieces than the source would have, and the verifier
    // completes a member on a per-read size test, so a member whose declared
    // size and declared CRC are both wrong could be refused by the other of
    // those two checks. The payload must come from the source instead.
    let mut raw = seekable_archive(&[
        Member::deflated("a.xml", compressible(13, 32)),
        Member::stored("pad.bin", vec![0x7e; 8_192]),
    ]);
    patch_central_record(&mut raw, "a.xml", 4_096, u32::MAX);

    let (archive, source) = open(raw);
    // The payload begins after the fixed header and the local variable region
    // of the first member, whose local header is at offset zero.
    let payload_start = {
        let name_len = usize::from(u16::from_le_bytes([source.bytes[26], source.bytes[27]]));
        let extra_len = usize::from(u16::from_le_bytes([source.bytes[28], source.bytes[29]]));
        (30 + name_len + extra_len) as u64
    };
    source.clear();

    let error = archive.read("a.xml").unwrap_err();
    assert!(
        !format!("{error}").is_empty(),
        "a member whose central record over-declares its payload stays refused"
    );
    assert!(
        source
            .reads()
            .iter()
            .any(|(offset, length)| *offset == payload_start && *length > 0),
        "the payload must be fetched from the source, reads were {:?}",
        source.reads()
    );
}

// ---------------------------------------------------------------------------
// Change 0623: the read-side local-span accessor.
//
// `local_span_hint` reports the span this archive's own first read covers, so a
// caller can fetch those bytes by another route. These tests pin the two
// properties that make such a fetch sound: the hint agrees with the read the
// archive actually issues, and adjoining hints leave no hole between them.
// ---------------------------------------------------------------------------

#[test]
fn the_local_span_hint_is_the_read_the_member_actually_issues() {
    let raw = seekable_archive(&[
        Member::deflated("one.xml", compressible(1, 900)),
        Member::deflated("two.xml", compressible(2, 900)),
        Member::stored("three.bin", compressible(3, 400)),
    ]);
    let (archive, source) = open(raw);
    for name in ["one.xml", "two.xml", "three.bin"] {
        let entry = archive.entry_id(name).expect("indexed member");
        let hint = archive.local_span_hint(entry).expect("a hinted member");
        source.clear();
        archive.read(name).expect("member reads");
        let reads = source.reads();
        assert_eq!(reads[0].0, hint.offset(), "{name}: first read offset");
        assert_eq!(
            reads[0].1 as u64,
            hint.length(),
            "{name}: first read length"
        );
        assert_eq!(hint.end(), hint.offset() + hint.length());
    }
}

#[test]
fn adjoining_local_span_hints_leave_no_hole() {
    let raw = seekable_archive(&[
        Member::deflated("a.xml", compressible(4, 700)),
        Member::deflated("b.xml", compressible(5, 700)),
        Member::deflated("c.xml", compressible(6, 700)),
    ]);
    let directory = directory_offset(&raw);
    let (archive, _source) = open(raw);
    let mut hints = Vec::new();
    for name in ["a.xml", "b.xml", "c.xml"] {
        let entry = archive.entry_id(name).expect("indexed member");
        hints.push(archive.local_span_hint(entry).expect("a hinted member"));
    }
    for window in hints.windows(2) {
        assert!(
            window[0].end() >= window[1].offset(),
            "hints {:?} and {:?} leave a hole",
            window[0],
            window[1]
        );
    }
    for hint in &hints {
        assert!(
            hint.end() <= directory + 24,
            "hint {hint:?} reaches past the central directory at {directory}"
        );
    }
}

#[test]
fn a_member_the_window_does_not_admit_has_no_local_span_hint() {
    // A member whose own local record is larger than the ceiling keeps the
    // historical grammar, and the accessor must say so rather than describe a
    // read that will not happen.
    let raw = seekable_archive(&[Member::stored("big.bin", vec![b'z'; 96 * 1024])]);
    let (archive, source) = open(raw);
    let entry = archive.entry_id("big.bin").expect("indexed member");
    assert!(archive.local_span_hint(entry).is_none());
    source.clear();
    archive.read("big.bin").expect("member reads");
    assert_eq!(
        source.reads()[0].1,
        30,
        "an unhinted member still begins with the 30-byte fixed header"
    );
}
