#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The reusable Deflate regression corpus is fixed and bounded."
)]

//! Differential coverage for the owned streaming Deflate transport.
//!
//! The owned entry API is deliberately compared with the older borrowed
//! `DeflateEncoder` construction.  The comparison includes the complete ZIP
//! bytes, rather than only decompressed payloads, so a reusable codec must
//! preserve raw-stream framing, input chunking, flush points, descriptors, and
//! central-directory layout.  The failure cases keep the sink small and
//! deterministic so they exercise progress and limit handling without a large
//! allocation or a long-running traversal.

use flate2::{Compression, read::DeflateDecoder, write::DeflateEncoder};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveLimits, StreamingArchiveWriter};
use soapberry_zip::{CompressionMethod, ErrorKind, ZipArchive, ZipArchiveWriter};
use std::cell::RefCell;
use std::io::{self, Read, Write};
use std::rc::Rc;

#[derive(Clone, Debug)]
struct Member {
    name: String,
    method: CompressionMethod,
    payload: Vec<u8>,
    zip64: bool,
}

#[derive(Clone, Copy, Debug)]
struct FeedPlan {
    chunk_sizes: &'static [usize],
    flush_every: Option<usize>,
}

const PLANS: &[FeedPlan] = &[
    FeedPlan {
        chunk_sizes: &[1, 2, 7, 31, 97],
        flush_every: Some(3),
    },
    FeedPlan {
        chunk_sizes: &[4093, 17, 6553, 3],
        flush_every: Some(2),
    },
    FeedPlan {
        chunk_sizes: &[16 * 1024],
        flush_every: None,
    },
];

fn patterned_payload(length: usize, seed: u64) -> Vec<u8> {
    let mut output = Vec::with_capacity(length);
    for index in 0..length {
        let position = index as u64;
        let block = (position / 257) ^ seed;
        let byte = if position % 19 < 13 {
            (block as u8).wrapping_add(0x31)
        } else {
            let mut value = position.wrapping_add(seed.wrapping_mul(0x9e37_79b9));
            value ^= value >> 30;
            value = value.wrapping_mul(0xbf58_476d_1ce4_e5b9);
            value ^= value >> 27;
            value = value.wrapping_mul(0x94d0_49bb_1331_11eb);
            value ^= value >> 31;
            value as u8
        };
        output.push(byte);
    }
    output
}

fn corpus() -> Vec<Member> {
    vec![
        Member {
            name: "empty.deflate".to_string(),
            method: CompressionMethod::Deflate,
            payload: Vec::new(),
            zip64: false,
        },
        Member {
            name: "first.deflate".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(96 * 1024 + 19, 1),
            zip64: false,
        },
        Member {
            name: "middle.store".to_string(),
            method: CompressionMethod::Store,
            payload: patterned_payload(8193, 2),
            zip64: false,
        },
        Member {
            name: "unicode/é.deflate".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(37 * 1024 + 5, 3),
            zip64: true,
        },
        Member {
            name: "tail.store".to_string(),
            method: CompressionMethod::Store,
            payload: patterned_payload(257, 4),
            zip64: false,
        },
        Member {
            name: "last.deflate".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(64 * 1024 + 1, 5),
            zip64: false,
        },
    ]
}

fn feed<W: Write>(writer: &mut W, payload: &[u8], plan: FeedPlan) -> io::Result<()> {
    if payload.is_empty() {
        return Ok(());
    }

    let mut offset = 0_usize;
    let mut chunk_number = 0_usize;
    while offset < payload.len() {
        let requested = plan.chunk_sizes[chunk_number % plan.chunk_sizes.len()].max(1);
        let end = offset
            .checked_add(requested.min(payload.len() - offset))
            .expect("bounded feed offset");
        writer.write_all(&payload[offset..end])?;
        offset = end;
        chunk_number += 1;
        if let Some(period) = plan.flush_every {
            if period != 0 && chunk_number % period == 0 {
                writer.flush()?;
            }
        }
    }
    Ok(())
}

fn build_owned(members: &[Member], plan: FeedPlan) -> Vec<u8> {
    let mut archive = ZipArchiveWriter::new(Vec::new());
    for member in members {
        let mut entry = if member.zip64 {
            archive
                .start_file_owned_zip64(&member.name, member.method)
                .expect("owned ZIP64 entry")
        } else {
            archive
                .start_file_owned(&member.name, member.method)
                .expect("owned entry")
        };
        feed(&mut entry, &member.payload, plan).expect("owned payload");
        archive = entry.finish().expect("owned entry finish");
    }
    archive.finish().expect("owned archive finish")
}

fn build_reference(members: &[Member], plan: FeedPlan, compression: Compression) -> Vec<u8> {
    let mut output = Vec::new();
    {
        let mut archive = ZipArchiveWriter::new(&mut output);
        for member in members {
            let builder = archive
                .new_file(&member.name)
                .compression_method(member.method);
            let builder = if member.zip64 {
                builder.zip64(true)
            } else {
                builder
            };
            let (mut entry, config) = builder.start().expect("reference entry");

            let descriptor = match member.method {
                CompressionMethod::Store => {
                    let mut writer = config.wrap(&mut entry);
                    feed(&mut writer, &member.payload, plan).expect("reference stored payload");
                    let (_, descriptor) = writer.finish().expect("stored data finish");
                    descriptor
                },
                CompressionMethod::Deflate => {
                    let encoder = DeflateEncoder::new(&mut entry, compression);
                    let mut writer = config.wrap(encoder);
                    feed(&mut writer, &member.payload, plan).expect("reference payload");
                    let (encoder, descriptor) = writer.finish().expect("reference data finish");
                    encoder.finish().expect("reference Deflate finish");
                    descriptor
                },
                other => panic!("unsupported test method: {other:?}"),
            };
            entry.finish(descriptor).expect("reference entry finish");
        }
        archive.finish().expect("reference archive finish");
    }
    output
}

fn assert_round_trip(bytes: &[u8], members: &[Member]) {
    let reader = ArchiveReader::new(bytes).expect("archive reader");
    for member in members {
        assert_eq!(
            reader.read(&member.name).expect("member read"),
            member.payload,
            "payload mismatch for {}",
            member.name
        );
    }
}

#[test]
fn owned_deflate_matches_fresh_entries_across_chunking_flushes_and_store_members() {
    let members = corpus();
    for (plan_number, plan) in PLANS.iter().copied().enumerate() {
        let owned = build_owned(&members, plan);
        let reference = build_reference(&members, plan, Compression::default());
        assert_eq!(
            owned, reference,
            "owned reusable transport changed complete ZIP bytes for feed plan {plan_number}"
        );
        assert_round_trip(&owned, &members);
    }
}

#[derive(Clone, Debug)]
struct LevelMember {
    name: &'static str,
    method: CompressionMethod,
    level: Option<Compression>,
    payload: Vec<u8>,
}

fn level_corpus() -> Vec<LevelMember> {
    vec![
        LevelMember {
            name: "fast.deflate",
            method: CompressionMethod::Deflate,
            level: Some(Compression::fast()),
            payload: patterned_payload(12 * 1024 + 3, 9),
        },
        LevelMember {
            name: "between.store",
            method: CompressionMethod::Store,
            level: None,
            payload: patterned_payload(4097, 10),
        },
        LevelMember {
            name: "best.deflate",
            method: CompressionMethod::Deflate,
            level: Some(Compression::best()),
            payload: patterned_payload(12 * 1024 + 5, 11),
        },
        LevelMember {
            name: "level-one.deflate",
            method: CompressionMethod::Deflate,
            level: Some(Compression::new(1)),
            payload: patterned_payload(12 * 1024 + 7, 12),
        },
        LevelMember {
            name: "level-nine.deflate",
            method: CompressionMethod::Deflate,
            level: Some(Compression::new(9)),
            payload: patterned_payload(12 * 1024 + 11, 13),
        },
    ]
}

fn build_level_reference(members: &[LevelMember]) -> Vec<u8> {
    let mut output = Vec::new();
    {
        let mut archive = ZipArchiveWriter::new(&mut output);
        for member in members {
            let builder = archive
                .new_file(member.name)
                .compression_method(member.method);
            let (mut entry, config) = builder.start().expect("level reference entry");
            match member.method {
                CompressionMethod::Store => {
                    let mut writer = config.wrap(&mut entry);
                    writer
                        .write_all(&member.payload)
                        .expect("stored level payload");
                    let (_, descriptor) = writer.finish().expect("stored level finish");
                    entry.finish(descriptor).expect("stored level entry");
                },
                CompressionMethod::Deflate => {
                    let level = member.level.expect("Deflate level");
                    let encoder = DeflateEncoder::new(&mut entry, level);
                    let mut writer = config.wrap(encoder);
                    writer.write_all(&member.payload).expect("level payload");
                    let (encoder, descriptor) = writer.finish().expect("level data finish");
                    encoder.finish().expect("level Deflate finish");
                    entry.finish(descriptor).expect("level entry finish");
                },
                other => panic!("unsupported level method: {other:?}"),
            }
        }
        archive.finish().expect("level archive finish");
    }
    output
}

fn fresh_deflate(payload: &[u8], compression: Compression) -> Vec<u8> {
    let mut encoder = DeflateEncoder::new(Vec::new(), compression);
    encoder.write_all(payload).expect("fresh payload");
    // `ZipDataWriter::finish` flushes its encoder before the encoder itself is
    // finished.  Mirror that sync-flush boundary in the independent oracle.
    encoder.flush().expect("fresh flush");
    encoder.finish().expect("fresh finish")
}

#[test]
fn fresh_levels_and_store_interleaving_keep_each_stream_independent() {
    let members = level_corpus();
    let bytes = build_level_reference(&members);
    let archive = ZipArchive::from_slice(&bytes).expect("level archive");
    let mut wayfinders = Vec::new();
    {
        let mut entries = archive.entries();
        for member in &members {
            let record = entries
                .next_entry()
                .expect("level entry result")
                .expect("level record");
            assert_eq!(record.compression_method(), member.method);
            wayfinders.push((record.wayfinder(), record.compressed_size_hint()));
        }
        assert!(entries.next_entry().expect("level end").is_none());
    }

    for (member, (wayfinder, compressed_size)) in members.iter().zip(wayfinders) {
        let entry = archive.get_entry(wayfinder).expect("level member lookup");
        match member.method {
            CompressionMethod::Store => assert_eq!(entry.data(), member.payload),
            CompressionMethod::Deflate => {
                let expected = fresh_deflate(
                    &member.payload,
                    member.level.expect("level for Deflate member"),
                );
                assert_eq!(entry.data(), expected.as_slice());
                assert_eq!(compressed_size as usize, expected.len());
                let mut decoder = DeflateDecoder::new(entry.data());
                let mut decoded = Vec::new();
                decoder
                    .read_to_end(&mut decoded)
                    .expect("decode level member");
                assert_eq!(decoded, member.payload);
            },
            other => panic!("unsupported level method: {other:?}"),
        }
    }
}

#[derive(Debug)]
struct ShortSink {
    bytes: Vec<u8>,
    maximum: usize,
}

impl ShortSink {
    fn new(maximum: usize) -> Self {
        assert!(maximum > 0);
        Self {
            bytes: Vec::new(),
            maximum,
        }
    }
}

impl Write for ShortSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        let amount = input.len().min(self.maximum);
        self.bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn build_owned_short(members: &[Member], plan: FeedPlan, maximum: usize) -> ShortSink {
    let mut archive = ZipArchiveWriter::new(ShortSink::new(maximum));
    for member in members {
        let mut entry = if member.zip64 {
            archive
                .start_file_owned_zip64(&member.name, member.method)
                .expect("short owned ZIP64 entry")
        } else {
            archive
                .start_file_owned(&member.name, member.method)
                .expect("short owned entry")
        };
        feed(&mut entry, &member.payload, plan).expect("short owned payload");
        archive = entry.finish().expect("short owned finish");
    }
    archive.finish().expect("short owned archive finish")
}

fn build_reference_short(members: &[Member], plan: FeedPlan, maximum: usize) -> ShortSink {
    let mut archive = ZipArchiveWriter::new(ShortSink::new(maximum));
    for member in members {
        let builder = archive
            .new_file(&member.name)
            .compression_method(member.method);
        let builder = if member.zip64 {
            builder.zip64(true)
        } else {
            builder
        };
        let (mut entry, config) = builder.start().expect("short reference entry");
        let descriptor = match member.method {
            CompressionMethod::Store => {
                let mut writer = config.wrap(&mut entry);
                feed(&mut writer, &member.payload, plan).expect("short stored payload");
                let (_, descriptor) = writer.finish().expect("short stored finish");
                descriptor
            },
            CompressionMethod::Deflate => {
                let encoder = DeflateEncoder::new(&mut entry, Compression::default());
                let mut writer = config.wrap(encoder);
                feed(&mut writer, &member.payload, plan).expect("short Deflate payload");
                let (encoder, descriptor) = writer.finish().expect("short data finish");
                encoder.finish().expect("short Deflate finish");
                descriptor
            },
            other => panic!("unsupported short method: {other:?}"),
        };
        entry
            .finish(descriptor)
            .expect("short reference entry finish");
    }
    archive.finish().expect("short reference archive finish")
}

#[test]
fn owned_deflate_matches_reference_through_short_sink_writes() {
    let members = corpus();
    let plan = FeedPlan {
        chunk_sizes: &[5, 4097, 13, 819],
        flush_every: Some(2),
    };
    let owned = build_owned_short(&members, plan, 3);
    let reference = build_reference_short(&members, plan, 3);
    assert_eq!(owned.bytes, reference.bytes);
    assert_round_trip(&owned.bytes, &members);
}

/// The old `flate2::write::DeflateEncoder` accepts input into its 32 KiB
/// output buffer before it attempts to write that buffer to the wrapped sink.
/// Keep a separately inspectable sink so direct `Write::write` calls can
/// compare that accepted-input and output timing with the reusable transport.
#[derive(Clone, Debug)]
struct SharedProbeSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    stop_after: Option<usize>,
    reject_oversize: bool,
    error_kind: io::ErrorKind,
    error_message: &'static str,
}

impl SharedProbeSink {
    fn new(
        stop_after: Option<usize>,
        error_kind: io::ErrorKind,
        error_message: &'static str,
    ) -> Self {
        Self {
            bytes: Rc::new(RefCell::new(Vec::new())),
            stop_after,
            reject_oversize: false,
            error_kind,
            error_message,
        }
    }

    fn atomic_limit(
        stop_after: usize,
        error_kind: io::ErrorKind,
        error_message: &'static str,
    ) -> Self {
        let mut sink = Self::new(Some(stop_after), error_kind, error_message);
        sink.reject_oversize = true;
        sink
    }

    fn snapshot(&self) -> Vec<u8> {
        self.bytes.borrow().clone()
    }
}

impl Write for SharedProbeSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        let mut bytes = self.bytes.borrow_mut();
        let amount = match self.stop_after {
            Some(stop_after) if bytes.len() >= stop_after => {
                return Err(io::Error::new(self.error_kind, self.error_message));
            },
            Some(stop_after) => {
                let remaining = stop_after - bytes.len();
                if self.reject_oversize && input.len() > remaining {
                    return Err(io::Error::new(self.error_kind, self.error_message));
                }
                input.len().min(remaining)
            },
            None => input.len(),
        };
        bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug, PartialEq, Eq)]
struct WriteStep {
    accepted: Result<usize, io::ErrorKind>,
    output: Vec<u8>,
}

fn capture_write<W: Write>(writer: &mut W, sink: &SharedProbeSink, input: &[u8]) -> WriteStep {
    let accepted = match writer.write(input) {
        Ok(bytes) => Ok(bytes),
        Err(error) => Err(error.kind()),
    };
    WriteStep {
        accepted,
        output: sink.snapshot(),
    }
}

fn random_payload(length: usize, seed: u64) -> Vec<u8> {
    let mut state = seed;
    (0..length)
        .map(|_| {
            state ^= state << 7;
            state ^= state >> 9;
            state ^= state << 8;
            state as u8
        })
        .collect()
}

fn direct_write_steps<W: Write>(
    writer: &mut W,
    sink: &SharedProbeSink,
    payload: &[u8],
    maximum_steps: usize,
) -> (Vec<WriteStep>, usize) {
    let mut steps = Vec::new();
    let mut offset = 0_usize;
    for _ in 0..maximum_steps {
        if offset == payload.len() {
            break;
        }
        let step = capture_write(writer, sink, &payload[offset..]);
        match step.accepted {
            Ok(accepted) if accepted > 0 => {
                offset = offset.checked_add(accepted).expect("direct write offset");
            },
            Ok(_) | Err(_) => {
                steps.push(step);
                break;
            },
        }
        steps.push(step);
    }
    (steps, offset)
}

#[derive(Debug)]
struct DirectRun {
    steps: Vec<WriteStep>,
    bytes: Vec<u8>,
}

fn run_owned_direct(payload: &[u8]) -> DirectRun {
    let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
    let mut archive = ZipArchiveWriter::new(sink.clone());
    let mut entry = archive
        .start_file_owned("direct.deflate", CompressionMethod::Deflate)
        .expect("owned direct entry");
    let (steps, accepted) = direct_write_steps(&mut entry, &sink, payload, 32);
    assert_eq!(
        accepted,
        payload.len(),
        "owned direct writes did not consume input"
    );
    archive = entry.finish().expect("owned direct finish");
    let _sink = archive.finish().expect("owned direct archive finish");
    DirectRun {
        steps,
        bytes: sink.snapshot(),
    }
}

fn run_reference_direct(payload: &[u8]) -> DirectRun {
    let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
    let mut archive = ZipArchiveWriter::new(sink.clone());
    let (mut entry, config) = archive
        .new_file("direct.deflate")
        .compression_method(CompressionMethod::Deflate)
        .start()
        .expect("reference direct entry");
    let encoder = DeflateEncoder::new(&mut entry, Compression::default());
    let mut writer = config.wrap(encoder);
    let (steps, accepted) = direct_write_steps(&mut writer, &sink, payload, 32);
    assert_eq!(
        accepted,
        payload.len(),
        "reference direct writes did not consume input"
    );
    let (encoder, descriptor) = writer.finish().expect("reference direct data finish");
    encoder.finish().expect("reference direct Deflate finish");
    entry
        .finish(descriptor)
        .expect("reference direct entry finish");
    archive.finish().expect("reference direct archive finish");
    DirectRun {
        steps,
        bytes: sink.snapshot(),
    }
}

#[test]
fn direct_write_acceptance_and_output_timing_matches_fresh_encoder() {
    let payload = random_payload(128 * 1024 + 37, 0x8e37_79b9_7f4a_7c15);
    let owned = run_owned_direct(&payload);
    let reference = run_reference_direct(&payload);

    assert_eq!(owned.steps, reference.steps);
    assert_eq!(owned.bytes, reference.bytes);
    assert_round_trip(
        &owned.bytes,
        &[Member {
            name: "direct.deflate".to_string(),
            method: CompressionMethod::Deflate,
            payload,
            zip64: false,
        }],
    );
}

fn run_owned_partial_failure(payload: &[u8], stop_after: usize) -> Vec<WriteStep> {
    let sink = SharedProbeSink::new(
        Some(stop_after),
        io::ErrorKind::BrokenPipe,
        "partial probe sink failed",
    );
    let archive = ZipArchiveWriter::new(sink.clone());
    let mut entry = archive
        .start_file_owned("failure-timing.deflate", CompressionMethod::Deflate)
        .expect("owned partial-failure entry");
    let first = capture_write(&mut entry, &sink, payload);
    let mut steps = vec![first];
    if let Ok(accepted) = steps[0].accepted {
        if accepted > 0 && accepted < payload.len() {
            steps.push(capture_write(&mut entry, &sink, &payload[accepted..]));
        }
    }
    steps
}

fn run_reference_partial_failure(payload: &[u8], stop_after: usize) -> Vec<WriteStep> {
    let sink = SharedProbeSink::new(
        Some(stop_after),
        io::ErrorKind::BrokenPipe,
        "partial probe sink failed",
    );
    let mut archive = ZipArchiveWriter::new(sink.clone());
    let (mut entry, config) = archive
        .new_file("failure-timing.deflate")
        .compression_method(CompressionMethod::Deflate)
        .start()
        .expect("reference partial-failure entry");
    let encoder = DeflateEncoder::new(&mut entry, Compression::default());
    let mut writer = config.wrap(encoder);
    let first = capture_write(&mut writer, &sink, payload);
    let mut steps = vec![first];
    if let Ok(accepted) = steps[0].accepted {
        if accepted > 0 && accepted < payload.len() {
            steps.push(capture_write(&mut writer, &sink, &payload[accepted..]));
        }
    }
    steps
}

#[test]
fn direct_write_reports_partial_sink_failure_on_the_same_call_as_fresh_encoder() {
    let payload = random_payload(128 * 1024 + 37, 0x1357_9bdf_2468_ace0);
    let prefix = 30 + "failure-timing.deflate".len();
    let stop_after = prefix + 7;
    let owned = run_owned_partial_failure(&payload, stop_after);
    let reference = run_reference_partial_failure(&payload, stop_after);
    assert_eq!(owned, reference);
    assert!(owned.iter().any(|step| step.accepted.is_err()));
}

fn run_owned_compressed_limit(payload: &[u8], maximum: u64) -> Vec<WriteStep> {
    let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
    let archive = ZipArchiveWriter::new(sink.clone());
    let mut entry = archive
        .start_file_owned("limit-timing.deflate", CompressionMethod::Deflate)
        .expect("owned limit-timing entry")
        .with_compressed_limit(maximum);
    let (steps, _) = direct_write_steps(&mut entry, &sink, payload, 2);
    steps
}

fn run_reference_compressed_limit(payload: &[u8], maximum: usize) -> Vec<WriteStep> {
    let name = "limit-timing.deflate";
    let prefix = 30 + name.len();
    let sink =
        SharedProbeSink::atomic_limit(prefix + maximum, io::ErrorKind::Other, "compressed limit");
    let mut archive = ZipArchiveWriter::new(sink.clone());
    let (mut entry, config) = archive
        .new_file(name)
        .compression_method(CompressionMethod::Deflate)
        .start()
        .expect("reference limit-timing entry");
    let encoder = DeflateEncoder::new(&mut entry, Compression::default());
    let mut writer = config.wrap(encoder);
    let (steps, _) = direct_write_steps(&mut writer, &sink, payload, 2);
    steps
}

#[test]
fn direct_write_compressed_limit_preserves_fresh_encoder_timing() {
    let payload = random_payload(128 * 1024 + 37, 0xa5a5_5a5a_1234_5678);
    // The first codec call produces more than seven bytes.  The old writer
    // buffers those bytes. The reference sink rejects the over-limit slice
    // when buffered output is drained by the next direct write.
    let maximum = 7_u64;
    let owned = run_owned_compressed_limit(&payload, maximum);
    let reference = run_reference_compressed_limit(&payload, maximum as usize);
    assert_eq!(owned, reference);
    assert!(owned.iter().any(|step| step.accepted.is_err()));
}

#[derive(Debug)]
struct StopSink {
    bytes: Vec<u8>,
    stop_at: usize,
    mode: StopMode,
}

#[derive(Clone, Copy, Debug)]
enum StopMode {
    Zero,
    Error,
}

impl StopSink {
    fn new(stop_at: usize, mode: StopMode) -> Self {
        Self {
            bytes: Vec::new(),
            stop_at,
            mode,
        }
    }
}

impl Write for StopSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.stop_at {
            return match self.mode {
                StopMode::Zero => Ok(0),
                StopMode::Error => Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "owned Deflate test sink stopped",
                )),
            };
        }
        let amount = input.len().min(self.stop_at - self.bytes.len());
        self.bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn assert_io_kind(error: &ErrorKind, expected: io::ErrorKind) {
    match error {
        ErrorKind::IO(error) | ErrorKind::Io(error) => assert_eq!(error.kind(), expected),
        other => panic!("expected I/O error, got {other:?}"),
    }
}

#[test]
fn owned_deflate_reports_zero_progress_and_sink_failure_at_finish() {
    let name = "stop.deflate";
    let prefix = 30 + name.len();
    let payload = b"small payload held until Deflate finish";

    for (mode, expected) in [
        (StopMode::Zero, io::ErrorKind::WriteZero),
        (StopMode::Error, io::ErrorKind::BrokenPipe),
    ] {
        let mut sink = StopSink::new(prefix, mode);
        {
            let archive = ZipArchiveWriter::new(&mut sink);
            let mut entry = archive
                .start_file_owned(name, CompressionMethod::Deflate)
                .expect("stopped owned entry");
            entry.write_all(payload).expect("payload stays buffered");
            let error = entry
                .finish()
                .expect_err("stopped sink unexpectedly succeeded");
            assert_io_kind(error.kind(), expected);
        }
        assert_eq!(sink.bytes.len(), prefix);
        assert!(ZipArchive::from_slice(&sink.bytes).is_err());
    }
}

#[test]
fn owned_deflate_compressed_limit_does_not_publish_a_partial_member() {
    let name = "limited.deflate";
    let payload = patterned_payload(64 * 1024 + 17, 31);
    let mut output = Vec::new();
    {
        let archive = ZipArchiveWriter::new(&mut output);
        let mut entry = archive
            .start_file_owned(name, CompressionMethod::Deflate)
            .expect("limited entry")
            .with_compressed_limit(0);
        let write_error = entry.write_all(&payload).err();
        let error = match write_error {
            Some(error) => {
                assert!(error.to_string().contains("compressed limit"));
                None
            },
            None => Some(entry.finish().expect_err("zero compressed limit succeeded")),
        };
        if let Some(error) = error {
            assert!(matches!(error.kind(), ErrorKind::IO(_) | ErrorKind::Io(_)));
            assert!(error.to_string().contains("compressed limit"));
        }
    }
    assert!(ZipArchive::from_slice(&output).is_err());
}

#[test]
fn streaming_writer_maps_owned_deflate_limit_to_typed_failure() {
    let limits = StreamingArchiveLimits::new(4, 128, 1024)
        .with_byte_limits(1 << 20, 1 << 20, 1 << 20)
        .with_compressed_size_limit(0);
    let writer = StreamingArchiveWriter::with_limits(limits);
    let mut entry = writer
        .start_entry("limited.bin", CompressionMethod::Deflate)
        .expect("bounded Deflate entry");
    let payload = patterned_payload(48 * 1024 + 9, 41);
    let write_error = entry.write_all(&payload).err();
    if let Some(error) = write_error {
        assert_eq!(error.kind(), io::ErrorKind::Other);
    }
    let failure = entry
        .finish()
        .err()
        .expect("bounded Deflate limit unexpectedly succeeded");
    assert!(matches!(
        failure.error().kind(),
        ErrorKind::LimitExceeded {
            resource: soapberry_zip::LimitResource::CompressedSize,
            ..
        }
    ));
    assert!(failure.progress().is_poisoned());
}

#[test]
fn dropping_incomplete_owned_deflate_does_not_publish_a_complete_archive() {
    let mut output = Vec::new();
    {
        let archive = ZipArchiveWriter::new(&mut output);
        let mut entry = archive
            .start_file_owned("dropped.deflate", CompressionMethod::Deflate)
            .expect("dropped entry");
        entry
            .write_all(b"this entry is intentionally incomplete")
            .expect("dropped payload");
        drop(entry);
    }
    assert!(ZipArchive::from_slice(&output).is_err());
}
