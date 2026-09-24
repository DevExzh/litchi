#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The reusable Deflate regression corpus is fixed and bounded."
)]

//! Differential coverage for the owned streaming Deflate transport.
//!
//! Owned (authored) Deflate entries follow the staged protocol of change
//! 0762: the member is cut into 16 KiB chunks at absolute member offsets, the
//! codec sees each chunk only once it is complete (or at an explicit flush or
//! the finish), always with an empty output buffer, and the stream ends with
//! one `Finish` and no sync flush. The owned entry is compared with a model
//! of that protocol written against flate2's raw codec, the model is checked
//! against flate2's own `DeflateEncoder` fed the same chunks, and complete
//! archives are compared across write splittings and short or interrupted
//! sinks. Borrowed entries keep the one-codec-call-per-write transport and are
//! still compared byte for byte with fresh `DeflateEncoder`s. The failure
//! cases keep the sink small and deterministic so they exercise progress and
//! limit handling without a large allocation or a long-running traversal.

use flate2::{
    Compress, Compression, FlushCompress, Status, read::DeflateDecoder, write::DeflateEncoder,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveLimits, StreamingArchiveWriter};
use soapberry_zip::{CompressionMethod, ErrorKind, ZipArchive, ZipArchiveWriter};
use std::cell::{Cell, RefCell};
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
    // Streaming-writer shaped: many short writes staged for one CRC-32 pass,
    // interleaved with writes on each side of the stage's 1,024-byte copy
    // threshold and one that exceeds its 4,096-byte capacity.
    FeedPlan {
        chunk_sizes: &[5, 31, 56, 12, 6, 1024, 1025, 3, 4096, 4097, 700],
        flush_every: None,
    },
    // Writes that straddle the 16 KiB Deflate chunk boundaries.
    FeedPlan {
        chunk_sizes: &[16 * 1024 - 1, 2, 40 * 1024, 16 * 1024 + 1],
        flush_every: Some(5),
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

/// Office-shaped XML text. Unlike the patterned payloads, zlib-rs encodes
/// it differently when its input is split differently, so it tells a staged
/// transport from one codec call per write.
fn xml_payload(paragraphs: usize) -> Vec<u8> {
    let mut output = Vec::new();
    for index in 0..paragraphs {
        output.extend_from_slice(
            format!(
                "<w:p><w:r><w:t xml:space=\"preserve\">paragraph {index:05} caf\u{e9} &amp; &lt;tags&gt; {}</w:t></w:r></w:p>",
                "lorem ipsum ".repeat(index % 5)
            )
            .as_bytes(),
        );
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
            name: "word/document.xml".to_string(),
            method: CompressionMethod::Deflate,
            payload: xml_payload(1_200),
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

/// Feeds `payload` by `plan` and returns the member offsets of its flushes.
fn feed<W: Write>(writer: &mut W, payload: &[u8], plan: FeedPlan) -> io::Result<Vec<usize>> {
    let mut flushes = Vec::new();
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
                flushes.push(offset);
            }
        }
    }
    Ok(flushes)
}

fn start_owned<W: Write>(
    archive: ZipArchiveWriter<W>,
    member: &Member,
) -> soapberry_zip::ZipOwnedEntryWriter<W> {
    if member.zip64 {
        archive
            .start_file_owned_zip64(&member.name, member.method)
            .expect("owned ZIP64 entry")
    } else {
        archive
            .start_file_owned(&member.name, member.method)
            .expect("owned entry")
    }
}

/// The archive and, per member, the offsets at which it was flushed.
fn build_owned(members: &[Member], plan: FeedPlan) -> (Vec<u8>, Vec<Vec<usize>>) {
    let mut archive = ZipArchiveWriter::new(Vec::new());
    let mut flushes = Vec::new();
    for member in members {
        let mut entry = start_owned(archive, member);
        flushes.push(feed(&mut entry, &member.payload, plan).expect("owned payload"));
        archive = entry.finish().expect("owned entry finish");
    }
    (archive.finish().expect("owned archive finish"), flushes)
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

/// Each member's method and raw stored data, in archive order.
fn member_data(bytes: &[u8]) -> Vec<(CompressionMethod, Vec<u8>)> {
    let archive = ZipArchive::from_slice(bytes).expect("member-data archive");
    let mut located = Vec::new();
    {
        let mut entries = archive.entries();
        while let Some(record) = entries.next_entry().expect("member-data entry") {
            located.push((record.compression_method(), record.wayfinder()));
        }
    }
    located
        .into_iter()
        .map(|(method, wayfinder)| {
            let entry = archive.get_entry(wayfinder).expect("member-data lookup");
            (method, entry.data().to_vec())
        })
        .collect()
}

// ---------------------------------------------------------------------------
// Change 0762: the staged protocol of owned (authored) Deflate entries.
// ---------------------------------------------------------------------------

/// Bytes of one owned Deflate chunk: soapberry-zip's frozen
/// `OWNED_DEFLATE_STAGE_BYTES`.
const STAGE: usize = 16 * 1024;

/// soapberry-zip's reusable Deflate output buffer.
const OUTPUT_BUFFER: usize = 32 * 1024;

/// The level of owned Deflate members.
const OWNED_LEVEL: u32 = 5;

/// Change 0762's staged protocol for one owned Deflate member, written
/// against flate2's raw codec: staged bytes reach the codec only when their
/// chunk is complete, at an explicit flush, or at the finish; every codec
/// call gets the whole output buffer; a flush is the sync flush of flate2's
/// writer; the finish is one `Finish` with no sync flush before it.
struct StagedModel {
    compress: Compress,
    buffer: Vec<u8>,
    stage: Vec<u8>,
    consumed: usize,
    output: Vec<u8>,
    /// Compressed length after each hand-off of staged bytes to the codec.
    marks: Vec<usize>,
}

impl StagedModel {
    fn new() -> Self {
        Self {
            compress: Compress::new(Compression::new(OWNED_LEVEL), false),
            buffer: vec![0; OUTPUT_BUFFER],
            stage: Vec::new(),
            consumed: 0,
            output: Vec::new(),
            marks: Vec::new(),
        }
    }

    fn codec(&mut self, from_stage: bool, flush: FlushCompress) -> (Status, usize, usize) {
        let input: &[u8] = if from_stage {
            &self.stage[self.consumed..]
        } else {
            &[]
        };
        let (before_in, before_out) = (self.compress.total_in(), self.compress.total_out());
        let status = self
            .compress
            .compress(input, &mut self.buffer, flush)
            .expect("model codec call");
        let consumed = usize::try_from(self.compress.total_in() - before_in).unwrap();
        let produced = usize::try_from(self.compress.total_out() - before_out).unwrap();
        self.output.extend_from_slice(&self.buffer[..produced]);
        (status, consumed, produced)
    }

    fn hand_off(&mut self) {
        while self.consumed < self.stage.len() {
            let (_, consumed, produced) = self.codec(true, FlushCompress::None);
            assert!(consumed > 0 || produced > 0, "model codec made no progress");
            self.consumed += consumed;
        }
        self.marks.push(self.output.len());
        if self.stage.len() == STAGE {
            self.stage.clear();
            self.consumed = 0;
        }
    }

    fn write(&mut self, mut bytes: &[u8]) {
        while !bytes.is_empty() {
            if self.stage.len() == STAGE {
                self.hand_off();
            }
            let take = (STAGE - self.stage.len()).min(bytes.len());
            self.stage.extend_from_slice(&bytes[..take]);
            bytes = &bytes[take..];
        }
    }

    fn flush(&mut self) {
        self.hand_off();
        self.codec(false, FlushCompress::Sync);
        while self.codec(false, FlushCompress::None).2 != 0 {}
    }

    fn finish(mut self) -> Vec<u8> {
        self.hand_off();
        while self.codec(false, FlushCompress::Finish).0 != Status::StreamEnd {}
        self.output
    }
}

/// One member through the model, flushed at the member offsets `flushes`.
fn staged_model(payload: &[u8], flushes: &[usize]) -> Vec<u8> {
    let mut model = StagedModel::new();
    let mut position = 0;
    for &flush in flushes {
        model.write(&payload[position..flush]);
        position = flush;
        model.flush();
    }
    model.write(&payload[position..]);
    model.finish()
}

/// flate2's own streaming encoder fed the same chunks and finished without a
/// flush: an oracle for the unflushed model that shares none of its code.
fn chunked_encoder_oracle(payload: &[u8]) -> Vec<u8> {
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::new(OWNED_LEVEL));
    for chunk in payload.chunks(STAGE) {
        encoder.write_all(chunk).expect("oracle chunk");
    }
    encoder.finish().expect("oracle finish")
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

#[test]
fn staged_model_agrees_with_flate2_encoder_fed_the_same_chunks() {
    for (length, seed) in [
        (0, 1),
        (1, 2),
        (STAGE - 1, 3),
        (STAGE, 4),
        (STAGE + 1, 5),
        (3 * STAGE + 777, 6),
        (96 * 1024 + 19, 7),
    ] {
        let patterned = patterned_payload(length, seed);
        assert_eq!(
            staged_model(&patterned, &[]),
            chunked_encoder_oracle(&patterned),
            "patterned length {length}"
        );
        let random = random_payload(length, seed | 1);
        assert_eq!(
            staged_model(&random, &[]),
            chunked_encoder_oracle(&random),
            "random length {length}"
        );
    }
}

#[test]
fn owned_deflate_members_follow_the_staged_protocol_across_plans_and_store_members() {
    let members = corpus();
    for (plan_number, plan) in PLANS.iter().copied().enumerate() {
        let (bytes, flushes) = build_owned(&members, plan);
        assert_round_trip(&bytes, &members);
        let data = member_data(&bytes);
        assert_eq!(data.len(), members.len());
        for ((member, member_flushes), (method, stored)) in members.iter().zip(&flushes).zip(&data)
        {
            assert_eq!(*method, member.method, "{}", member.name);
            match member.method {
                CompressionMethod::Store => assert_eq!(stored, &member.payload),
                CompressionMethod::Deflate => assert_eq!(
                    stored,
                    &staged_model(&member.payload, member_flushes),
                    "plan {plan_number}, member {}",
                    member.name
                ),
                other => panic!("unsupported test method: {other:?}"),
            }
        }
    }
}

/// A deterministic xorshift64* source of write sizes and flush points.
///
/// An odd seed draws only Office-writer-sized writes of at most 64 bytes, the
/// splitting under which zlib-rs's output differs most from one codec call
/// per member; an even seed mixes tiny, short, long and chunk-straddling
/// writes.
struct Splitter {
    state: u64,
    tiny: bool,
}

impl Splitter {
    fn new(seed: u64) -> Self {
        Self {
            state: seed | 0x100,
            tiny: seed % 2 == 1,
        }
    }

    fn next(&mut self) -> u64 {
        self.state ^= self.state >> 12;
        self.state ^= self.state << 25;
        self.state ^= self.state >> 27;
        self.state.wrapping_mul(0x2545_f491_4f6c_dd1d)
    }

    fn below(&mut self, bound: u64) -> usize {
        usize::try_from(self.next() % bound).unwrap()
    }

    /// The next write size; zero-length writes are included.
    fn write_size(&mut self) -> usize {
        if self.tiny {
            return self.below(65);
        }
        match self.below(6) {
            0 => self.below(9),
            1 => 1 + self.below(100),
            2 => 1 + self.below(5_000),
            3 => STAGE - 2 + self.below(5),
            4 => 1 + self.below(40_000),
            _ => 1 + self.below(3 * STAGE as u64),
        }
    }
}

/// One member's payload written in random pieces, never across a flush
/// offset, flushing at each of `flushes`.
fn write_split<W: Write>(
    entry: &mut W,
    payload: &[u8],
    flushes: &[usize],
    splitter: &mut Splitter,
) -> io::Result<()> {
    let mut position = 0;
    let mut flush_points = flushes.iter().copied().peekable();
    loop {
        while flush_points.peek() == Some(&position) {
            flush_points.next();
            entry.flush()?;
        }
        if position == payload.len() {
            return Ok(());
        }
        let limit = flush_points.peek().copied().unwrap_or(payload.len());
        let end = (position + splitter.write_size()).min(limit);
        entry.write_all(&payload[position..end])?;
        position = end;
    }
}

/// A sink that accepts at most `maximum` bytes per call.
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

/// A sink that interrupts every `period`-th call while armed, accepting at
/// most seven bytes per call otherwise.
#[derive(Debug)]
struct InterruptingSink {
    bytes: Vec<u8>,
    calls: usize,
    period: usize,
    armed: Rc<Cell<bool>>,
    interruptions: Rc<Cell<usize>>,
}

impl Write for InterruptingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.calls += 1;
        if self.armed.get() && self.calls % self.period == 0 {
            self.interruptions.set(self.interruptions.get() + 1);
            return Err(io::ErrorKind::Interrupted.into());
        }
        let amount = input.len().min(7);
        self.bytes.extend_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// The corpus written with random splits (and, per member, the given flush
/// offsets) into `sink`.
fn build_split<W: Write>(sink: W, members: &[Member], flushes: &[Vec<usize>], seed: u64) -> W {
    let mut splitter = Splitter::new(seed);
    let mut archive = ZipArchiveWriter::new(sink);
    for (member, member_flushes) in members.iter().zip(flushes) {
        let mut entry = start_owned(archive, member);
        write_split(&mut entry, &member.payload, member_flushes, &mut splitter)
            .expect("split payload");
        archive = entry.finish().expect("split entry finish");
    }
    archive.finish().expect("split archive finish")
}

#[test]
fn owned_deflate_bytes_do_not_depend_on_write_sizes_or_sink_behaviour() {
    let members = corpus();
    let unflushed = vec![Vec::new(); members.len()];
    let (reference, _) = build_owned(
        &members,
        FeedPlan {
            chunk_sizes: &[usize::MAX],
            flush_every: None,
        },
    );
    for ((method, stored), member) in member_data(&reference).iter().zip(&members) {
        if *method == CompressionMethod::Deflate {
            assert_eq!(stored, &chunked_encoder_oracle(&member.payload));
        }
    }
    for seed in 1..=24_u64 {
        let seed = seed.wrapping_mul(0x9e37_79b9_7f4a_7c15);
        assert_eq!(
            build_split(Vec::new(), &members, &unflushed, seed),
            reference,
            "seed {seed:#x}"
        );
    }
    for (seed, maximum) in [(11_u64, 1_usize), (12, 3), (13, 4097)] {
        assert_eq!(
            build_split(ShortSink::new(maximum), &members, &unflushed, seed).bytes,
            reference,
            "short sink of {maximum}"
        );
    }
    for (seed, period) in [(21_u64, 2_usize), (22, 3), (23, 7)] {
        let armed = Rc::new(Cell::new(true));
        let interruptions = Rc::new(Cell::new(0));
        let sink = InterruptingSink {
            bytes: Vec::new(),
            calls: 0,
            period,
            armed: Rc::clone(&armed),
            interruptions: Rc::clone(&interruptions),
        };
        // `write_all` retries an interrupted write; the entry's `finish` does
        // not, so the sink stops interrupting before the members finish.
        let mut splitter = Splitter::new(seed);
        let mut archive = ZipArchiveWriter::new(sink);
        for member in &members {
            let mut entry = start_owned(archive, member);
            armed.set(true);
            write_split(&mut entry, &member.payload, &[], &mut splitter).expect("interrupted");
            armed.set(false);
            archive = entry.finish().expect("interrupted entry finish");
        }
        let sink = archive.finish().expect("interrupted archive finish");
        assert!(interruptions.get() > 0);
        assert_eq!(
            sink.bytes, reference,
            "interrupting sink of period {period}"
        );
    }
}

#[test]
fn explicit_flushes_cut_chunks_but_keep_the_absolute_boundaries() {
    let members = corpus();
    let mut splitter = Splitter::new(0x5eed);
    let flushes: Vec<Vec<usize>> = members
        .iter()
        .map(|member| {
            let mut offsets: Vec<usize> = (0..4)
                .map(|_| splitter.below(member.payload.len() as u64 + 1))
                .collect();
            // A chunk boundary, and the same offset twice.
            if member.payload.len() > STAGE {
                offsets.push(STAGE);
            }
            offsets.push(member.payload.len() / 2);
            offsets.push(member.payload.len() / 2);
            offsets.sort_unstable();
            offsets
        })
        .collect();
    let reference = build_split(Vec::new(), &members, &flushes, 1);
    assert_round_trip(&reference, &members);
    for ((method, stored), (member, member_flushes)) in member_data(&reference)
        .iter()
        .zip(members.iter().zip(&flushes))
    {
        if *method == CompressionMethod::Deflate {
            assert_eq!(stored, &staged_model(&member.payload, member_flushes));
        }
    }
    for seed in 2..=12_u64 {
        assert_eq!(
            build_split(Vec::new(), &members, &flushes, seed),
            reference,
            "seed {seed}"
        );
    }
    assert_eq!(
        build_split(ShortSink::new(5), &members, &flushes, 99).bytes,
        reference
    );
}

/// A sink whose output can be inspected while the writer still owns it, and
/// which fails once it holds `stop_after` bytes.
#[derive(Clone, Debug)]
struct SharedProbeSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    stop_after: Option<usize>,
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
            error_kind,
            error_message,
        }
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
            Some(stop_after) => input.len().min(stop_after - bytes.len()),
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

/// Direct `write` calls of the rest of `payload`, until an error, a zero
/// acceptance or `maximum_steps`.
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

/// The model's compressed length after each chunk hand-off of a payload
/// written in one piece.
fn model_marks(payload: &[u8]) -> Vec<usize> {
    let mut model = StagedModel::new();
    model.write(payload);
    model.hand_off();
    model.marks
}

#[test]
fn owned_deflate_accepts_input_up_to_each_chunk_boundary() {
    let payload = random_payload(128 * 1024 + 37, 0x8e37_79b9_7f4a_7c15);
    let name = "direct.deflate";
    let header = 30 + name.len();
    let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
    let archive = ZipArchiveWriter::new(sink.clone());
    let mut entry = archive
        .start_file_owned(name, CompressionMethod::Deflate)
        .expect("owned direct entry");
    let (steps, accepted) = direct_write_steps(&mut entry, &sink, &payload, 32);
    assert_eq!(accepted, payload.len());
    let marks = model_marks(&payload);
    let compressed = staged_model(&payload, &[]);
    for (index, step) in steps.iter().enumerate() {
        // Each write accepts up to the next chunk boundary.
        let expected = STAGE.min(payload.len() - index * STAGE);
        assert_eq!(step.accepted, Ok(expected), "step {index}");
        // A write hands the chunk the previous one completed to the codec and
        // sends its output to the sink before accepting more input.
        let sent = if index == 0 { 0 } else { marks[index - 1] };
        assert_eq!(step.output.len(), header + sent, "step {index}");
        assert_eq!(&step.output[header..], &compressed[..sent], "step {index}");
    }
    let _sink = entry
        .finish()
        .expect("owned direct finish")
        .finish()
        .expect("owned direct archive finish");
    let bytes = sink.snapshot();
    assert_eq!(member_data(&bytes)[0].1, compressed);
    assert_round_trip(
        &bytes,
        &[Member {
            name: name.to_string(),
            method: CompressionMethod::Deflate,
            payload,
            zip64: false,
        }],
    );
}

/// The index of the first direct write whose chunk hand-off sends more than
/// `allowance` compressed bytes, for a payload written by
/// `direct_write_steps`.
fn first_step_sending_more_than(payload: &[u8], allowance: usize) -> usize {
    let marks = model_marks(payload);
    1 + marks
        .iter()
        .position(|&sent| sent > allowance)
        .expect("the payload produces output before its finish")
}

#[test]
fn owned_deflate_reports_a_sink_failure_when_its_chunk_reaches_the_codec() {
    let payload = random_payload(128 * 1024 + 37, 0x1357_9bdf_2468_ace0);
    let name = "failure-timing.deflate";
    let stop_after = 30 + name.len() + 7;
    let sink = SharedProbeSink::new(
        Some(stop_after),
        io::ErrorKind::BrokenPipe,
        "partial probe sink failed",
    );
    let archive = ZipArchiveWriter::new(sink.clone());
    let mut entry = archive
        .start_file_owned(name, CompressionMethod::Deflate)
        .expect("owned partial-failure entry");
    let (steps, _) = direct_write_steps(&mut entry, &sink, &payload, 32);
    let failing = first_step_sending_more_than(&payload, 7);
    assert_eq!(steps.len(), failing + 1);
    for (index, step) in steps.iter().enumerate() {
        if index < failing {
            assert_eq!(step.accepted, Ok(STAGE), "step {index}");
        } else {
            // The failure surfaces on the write that sends the output, which
            // accepts none of its input.
            assert_eq!(step.accepted, Err(io::ErrorKind::BrokenPipe));
        }
        assert!(step.output.len() <= stop_after);
    }
    assert_eq!(sink.snapshot().len(), stop_after);
    assert!(entry.finish().is_err());
}

#[test]
fn owned_deflate_compressed_limit_refuses_at_the_codec_and_is_never_exceeded() {
    let payload = random_payload(128 * 1024 + 37, 0xa5a5_5a5a_1234_5678);
    let name = "limit-timing.deflate";
    let header = 30 + name.len();
    let maximum = 7_u64;
    let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
    let archive = ZipArchiveWriter::new(sink.clone());
    let mut entry = archive
        .start_file_owned(name, CompressionMethod::Deflate)
        .expect("owned limit-timing entry")
        .with_compressed_limit(maximum);
    let (steps, _) = direct_write_steps(&mut entry, &sink, &payload, 32);
    let failing = first_step_sending_more_than(&payload, 7);
    assert_eq!(steps.len(), failing + 1);
    let last = steps.last().expect("a refused step");
    assert_eq!(last.accepted, Err(io::ErrorKind::Other));
    // The refusal is decided before the output reaches the sink.
    assert!(sink.snapshot().len() <= header + 7);

    // Whatever the write splitting, an entry succeeds exactly when its whole
    // compressed stream fits, and the sink never receives more than the limit.
    let total = staged_model(&payload, &[]).len();
    let total_u64 = u64::try_from(total).unwrap();
    for (maximum, seed) in [
        (0, 1),
        (1, 2),
        (7, 3),
        (total_u64 / 2, 4),
        (total_u64 - 1, 5),
        (total_u64, 6),
        (total_u64 + 1, 7),
    ] {
        let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
        let archive = ZipArchiveWriter::new(sink.clone());
        let mut entry = archive
            .start_file_owned(name, CompressionMethod::Deflate)
            .expect("owned limit-sweep entry")
            .with_compressed_limit(maximum);
        let written = write_split(&mut entry, &payload, &[], &mut Splitter::new(seed));
        let finished = match written {
            Ok(()) => entry.finish().is_ok(),
            Err(_) => false,
        };
        assert_eq!(finished, total_u64 <= maximum, "limit {maximum}");
        if finished {
            // The finished member is the unlimited stream.
            let bytes = sink.snapshot();
            assert_eq!(
                &bytes[header..header + total],
                &staged_model(&payload, &[])[..]
            );
        } else {
            let sent = sink.snapshot().len() - header;
            assert!(
                u64::try_from(sent).unwrap() <= maximum,
                "limit {maximum}: {sent} compressed bytes reached the sink"
            );
        }
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

// ---------------------------------------------------------------------------
// Change 0618: the Office streaming writer reuses one Deflate state across the
// members of one archive. Every member must still be the byte-for-byte output
// of a freshly constructed encoder, including after a refused member.
// ---------------------------------------------------------------------------

/// The streaming writer's own zip64 selection for default limits, mirrored so
/// the reference archive opens its entries the same way.
fn streaming_zip64_for_default_limits() -> bool {
    let limits = StreamingArchiveLimits::default();
    limits.max_compressed_size >= u64::from(u32::MAX)
        || limits.max_entry_size >= u64::from(u32::MAX)
}

fn streaming_corpus() -> Vec<Member> {
    vec![
        Member {
            name: "first.xml".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(3_800, 21),
            zip64: false,
        },
        Member {
            name: "empty.xml".to_string(),
            method: CompressionMethod::Deflate,
            payload: Vec::new(),
            zip64: false,
        },
        Member {
            name: "second.xml".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(96 * 1024 + 19, 22),
            zip64: false,
        },
        Member {
            name: "unicode/é.xml".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(1_021, 23),
            zip64: false,
        },
        Member {
            name: "third.xml".to_string(),
            method: CompressionMethod::Deflate,
            payload: patterned_payload(37 * 1024 + 5, 24),
            zip64: false,
        },
    ]
}

/// One fresh `DeflateEncoder` per member, through the same borrowed entry API
/// the streaming writer used before the state became reusable.
fn fresh_streaming_reference(members: &[Member], zip64: bool) -> Vec<u8> {
    let mut output = Vec::new();
    {
        let mut archive = ZipArchiveWriter::new(&mut output);
        for member in members {
            let (mut entry, config) = archive
                .new_file(&member.name)
                .compression_method(CompressionMethod::Deflate)
                .zip64(zip64)
                .start()
                .expect("reference streaming entry");
            let encoder = DeflateEncoder::new(&mut entry, Compression::default());
            let mut writer = config.wrap(encoder);
            writer
                .write_all(&member.payload)
                .expect("reference payload");
            let (encoder, descriptor) = writer.finish().expect("reference data finish");
            encoder.finish().expect("reference Deflate finish");
            entry.finish(descriptor).expect("reference entry finish");
        }
        archive.finish().expect("reference archive finish");
    }
    output
}

#[test]
fn streaming_deflate_members_match_one_fresh_encoder_each() {
    let members = streaming_corpus();
    let mut writer = StreamingArchiveWriter::new();
    for member in &members {
        writer
            .write_deflated(&member.name, &member.payload)
            .expect("streaming Deflate member");
    }
    let reused = writer.finish_to_bytes().expect("streaming archive finish");
    let reference = fresh_streaming_reference(&members, streaming_zip64_for_default_limits());
    assert_eq!(
        reused, reference,
        "reusing one Deflate state changed the streaming archive bytes"
    );
    assert_round_trip(&reused, &members);
}

#[test]
fn streaming_sized_deflate_members_match_one_fresh_encoder_each() {
    let members = streaming_corpus();
    let mut writer = StreamingArchiveWriter::new();
    for member in &members {
        writer
            .write_deflated_sized(&member.name, &member.payload)
            .expect("streaming sized Deflate member");
    }
    let reused = writer.finish_to_bytes().expect("sized archive finish");

    let mut reference_writer = ZipArchiveWriter::new(Vec::new());
    for member in &members {
        // The sized route compresses into scratch with no intervening flush,
        // then publishes the member with upfront sizes.
        let mut encoder = DeflateEncoder::new(Vec::new(), Compression::default());
        encoder.write_all(&member.payload).expect("sized payload");
        let compressed = encoder.finish().expect("sized Deflate finish");
        reference_writer
            .write_precompressed_file(
                &member.name,
                CompressionMethod::Deflate,
                soapberry_zip::crc32(&member.payload),
                member.payload.len() as u64,
                &compressed,
            )
            .expect("sized reference member");
    }
    let reference = reference_writer.finish().expect("sized reference finish");
    assert_eq!(
        reused, reference,
        "reusing one Deflate state changed the sized streaming archive bytes"
    );
    assert_round_trip(&reused, &members);
}

/// The reader-fed route copies through a fixed stream buffer, so its encoder
/// sees the payload in fixed-size pieces. Deflate output is not invariant
/// under that regrouping (the streaming and whole-payload routes have always
/// produced slightly different bytes for the same member), so the reference
/// for this route feeds the same pieces.
fn fresh_chunked_reference(members: &[Member], zip64: bool, chunk: usize) -> Vec<u8> {
    let mut output = Vec::new();
    {
        let mut archive = ZipArchiveWriter::new(&mut output);
        for member in members {
            let (mut entry, config) = archive
                .new_file(&member.name)
                .compression_method(CompressionMethod::Deflate)
                .zip64(zip64)
                .start()
                .expect("reference chunked entry");
            let encoder = DeflateEncoder::new(&mut entry, Compression::default());
            let mut writer = config.wrap(encoder);
            for piece in member.payload.chunks(chunk) {
                writer.write_all(piece).expect("reference chunked payload");
            }
            let (encoder, descriptor) = writer.finish().expect("reference chunked data finish");
            encoder.finish().expect("reference chunked Deflate finish");
            entry
                .finish(descriptor)
                .expect("reference chunked entry finish");
        }
        archive.finish().expect("reference chunked archive finish");
    }
    output
}

/// A reader that never returns more than `chunk` bytes per call, so the copy
/// loop's writes — and therefore the encoder's input pieces — are known.
struct ChunkedReader<'a> {
    payload: &'a [u8],
    chunk: usize,
}

impl Read for ChunkedReader<'_> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let take = self.payload.len().min(self.chunk).min(buffer.len());
        buffer[..take].copy_from_slice(&self.payload[..take]);
        self.payload = &self.payload[take..];
        Ok(take)
    }
}

#[test]
fn streaming_deflate_stream_members_match_one_fresh_encoder_each() {
    const CHUNK: usize = 4_096;
    let members = streaming_corpus();
    let mut writer = StreamingArchiveWriter::new();
    for member in &members {
        writer
            .write_deflated_stream(
                &member.name,
                ChunkedReader {
                    payload: &member.payload,
                    chunk: CHUNK,
                },
            )
            .expect("streaming Deflate reader member");
    }
    let reused = writer.finish_to_bytes().expect("reader archive finish");
    let reference = fresh_chunked_reference(&members, streaming_zip64_for_default_limits(), CHUNK);
    assert_eq!(
        reused, reference,
        "reusing one Deflate state changed the reader-fed archive bytes"
    );
    assert_round_trip(&reused, &members);
}

#[test]
fn a_refused_sized_member_leaves_the_next_member_byte_identical() {
    // The sized route refuses before any archive byte is written, so the
    // writer stays usable. The refused member's Deflate stream is abandoned
    // part-way; the member after it must not inherit any of it.
    let limits = StreamingArchiveLimits {
        max_compressed_size: 4 * 1024,
        ..StreamingArchiveLimits::default()
    };
    let kept = streaming_corpus();
    let oversized = patterned_payload(256 * 1024, 25);

    let mut writer = StreamingArchiveWriter::with_limits(limits);
    writer
        .write_deflated_sized(&kept[3].name, &kept[3].payload)
        .expect("small member fits the compressed-size limit");
    let refused = writer
        .write_deflated_sized("refused.bin", &oversized)
        .expect_err("an oversized member must be refused");
    assert!(
        matches!(refused.kind(), ErrorKind::LimitExceeded { .. }),
        "expected a typed compressed-size refusal, got {refused:?}"
    );
    writer
        .write_deflated_sized(&kept[1].name, &kept[1].payload)
        .expect("the writer stays usable after a refused member");
    let after_refusal = writer.finish_to_bytes().expect("archive finish");

    let mut clean = StreamingArchiveWriter::with_limits(limits);
    clean
        .write_deflated_sized(&kept[3].name, &kept[3].payload)
        .expect("small member");
    clean
        .write_deflated_sized(&kept[1].name, &kept[1].payload)
        .expect("small member");
    let without_refusal = clean.finish_to_bytes().expect("archive finish");

    assert_eq!(
        after_refusal, without_refusal,
        "an abandoned Deflate stream leaked into the member that followed it"
    );
}

// ---------------------------------------------------------------------------
// Change 0762: owned and borrowed members share one reusable Deflate state.
// An owned member's stage, level and final block must not reach the borrowed
// member that follows it, nor the other way round.
// ---------------------------------------------------------------------------

#[test]
fn owned_and_borrowed_members_of_one_archive_keep_their_own_protocols() {
    let members = streaming_corpus();
    let mut writer = StreamingArchiveWriter::new();
    for (index, member) in members.iter().enumerate() {
        if index % 2 == 0 {
            let mut entry = writer
                .start_entry(&member.name, CompressionMethod::Deflate)
                .expect("owned streaming entry");
            for piece in member.payload.chunks(1 + 97 * index) {
                entry.write_all(piece).expect("owned streaming payload");
            }
            writer = entry.finish().expect("owned streaming finish");
        } else {
            writer
                .write_deflated(&member.name, &member.payload)
                .expect("borrowed streaming member");
        }
    }
    let bytes = writer.finish_to_bytes().expect("mixed archive finish");
    assert_round_trip(&bytes, &members);
    for (index, ((method, stored), member)) in member_data(&bytes).iter().zip(&members).enumerate()
    {
        assert_eq!(*method, CompressionMethod::Deflate);
        let expected = if index % 2 == 0 {
            staged_model(&member.payload, &[])
        } else {
            // `write_deflated` keeps one codec call per write and the sync
            // flush `ZipDataWriter::finish` performs, at the default level.
            fresh_deflate(&member.payload, Compression::default())
        };
        assert_eq!(stored, &expected, "{}", member.name);
    }
}

/// Writes three owned Deflate members in Office-sized pieces through a
/// bounded streaming archive into `sink`.
fn write_bounded_streaming(
    sink: SharedProbeSink,
    limits: StreamingArchiveLimits,
) -> Result<(), soapberry_zip::office::StreamingArchiveFailure> {
    let payloads = [
        ("word/document.xml", xml_payload(900)),
        ("docProps/app.xml", xml_payload(3)),
        (
            "xl/worksheets/sheet1.xml",
            patterned_payload(70 * 1024 + 3, 61),
        ),
    ];
    let mut writer = StreamingArchiveWriter::with_writer_and_limits(sink, limits);
    let mut splitter = Splitter::new(3);
    for (name, payload) in &payloads {
        let mut entry = writer.start_entry(name, CompressionMethod::Deflate)?;
        // A refused write poisons the entry; its `finish` reports the typed
        // failure.
        let _ = write_split(&mut entry, payload, &[], &mut splitter);
        writer = entry.finish()?;
    }
    writer.finish_with_progress().map(|_| ())
}

#[test]
fn streaming_output_and_compressed_limits_stay_exact_on_batched_members() {
    // Metadata limits small enough for any output limit tried below.
    let unlimited = StreamingArchiveLimits::new(8, 64, 256);
    let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
    write_bounded_streaming(sink.clone(), unlimited).expect("unlimited archive");
    let bytes = sink.snapshot();
    let size = u64::try_from(bytes.len()).unwrap();
    let largest = member_data(&bytes)
        .iter()
        .map(|(_, data)| u64::try_from(data.len()).unwrap())
        .max()
        .unwrap();

    for limit in [size - 1, size, size + 1] {
        let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
        let result = write_bounded_streaming(
            sink.clone(),
            unlimited.with_byte_limits(unlimited.max_entry_size, unlimited.max_total_size, limit),
        );
        let received = u64::try_from(sink.snapshot().len()).unwrap();
        assert!(received <= limit, "output limit {limit}: {received} bytes");
        match result {
            Ok(()) => {
                assert!(limit >= size);
                assert_eq!(sink.snapshot(), bytes);
            },
            Err(failure) => {
                assert!(limit < size);
                assert_eq!(
                    failure.limit().map(|limit| limit.resource()),
                    Some(soapberry_zip::office::StreamingLimitResource::OutputBytes)
                );
            },
        }
    }

    for limit in [largest - 1, largest, largest + 1] {
        let sink = SharedProbeSink::new(None, io::ErrorKind::Other, "unused probe failure");
        let result =
            write_bounded_streaming(sink.clone(), unlimited.with_compressed_size_limit(limit));
        let written = sink.snapshot();
        match result {
            Ok(()) => {
                assert!(limit >= largest);
                assert_eq!(written, bytes);
            },
            Err(failure) => {
                assert!(limit < largest);
                assert!(matches!(
                    failure.error().kind(),
                    ErrorKind::LimitExceeded {
                        resource: soapberry_zip::LimitResource::CompressedSize,
                        ..
                    }
                ));
                // Everything written before the refused member is the
                // unlimited archive's prefix, and no member's data reached
                // the sink beyond the limit.
                assert_eq!(&bytes[..written.len()], &written[..]);
            },
        }
    }
}
