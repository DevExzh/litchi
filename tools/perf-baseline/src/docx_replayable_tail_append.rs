//! Reproducible lifecycle measurements for replayable DOCX tail appends.
//!
//! The harness keeps source and authored paragraph counts as independent axes.
//! Corpus construction, XML/semantic/ZIP oracles, and exact inverse checks run
//! outside the timed lifecycle.  A timed sample owns source admission,
//! replayable stream preparation, publication to a short sequential sink, and
//! all drops.  The authored provider emits borrowed chunks from one bounded
//! cursor buffer; it never builds a complete authored XML stream.

#![allow(clippy::module_name_repetitions)]

use std::{
    collections::BTreeMap,
    ffi::OsString,
    fmt::Write as FmtWrite,
    fs::OpenOptions,
    io::{self, Cursor, Write},
    mem::size_of,
    ops::Range,
    path::{Path, PathBuf},
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::Instant,
};

use litchi_core::{ReadAt, SourceVersion};
use litchi_docx::source_backed::tail_append::Limits as TailAppendLimits;
use litchi_docx::source_backed::tail_append_stream::{
    AuthoredStreamProof, ParagraphCursor, ParagraphStreamLimits, PlainParagraphEvent,
    ReplayableParagraphSource,
};
use litchi_docx::{Package as OwnedDocxPackage, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use serde::Serialize;
use sha2::{Digest as _, Sha256};
use soapberry_zip::ZipArchive;
use soapberry_zip::office::ArchiveReader;

use super::{allocation_metrics, process_metrics};

const SCHEMA: &str = "docx-replayable-tail-append-v1";
const DEFAULT_SOURCE_COUNTS: [usize; 3] = [64, 8_192, 131_072];
const DEFAULT_AUTHORED_COUNTS: [usize; 4] = [64, 256, 4_096, 16_384];
const DEFAULT_CHUNK_BYTES: [usize; 3] = [0, 64, 8 * 1024];
const DEFAULT_SAMPLES: usize = 15;
const DEFAULT_WARMUPS: usize = 3;
const OPAQUE_PATH: &str = "word/perf-opaque.bin";
const OPAQUE_BYTES: usize = 32 * 1024;
const MAIN_PATH: &str = "word/document.xml";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HASH_SINK_MAX_WRITE: usize = 4 * 1024;
const MAX_CURSOR_TEXT_BYTES: usize = 60 * 1024;
const MAX_STREAM_XML_DEPTH: u64 = 16;
const EXPECTED_AUTHORED_OPENS: u64 = 5;

type BenchResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq, Hash)]
#[serde(rename_all = "snake_case")]
pub enum ChunkMode {
    /// Emit each paragraph's text in one borrowed chunk.
    One,
    /// Partition text into independent 64-byte borrowed chunks.
    Fixed64,
    /// Partition text into chunks close to the replay window.
    ReplayWindow,
}

impl ChunkMode {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "one" | "single" => Ok(Self::One),
            "64" | "fixed64" => Ok(Self::Fixed64),
            "window" | "replay_window" => Ok(Self::ReplayWindow),
            _ => Err(format!("invalid chunk mode {value:?}; expected one, 64, or window").into()),
        }
    }

    const fn chunk_bytes(self) -> usize {
        match self {
            Self::One => DEFAULT_CHUNK_BYTES[0],
            Self::Fixed64 => DEFAULT_CHUNK_BYTES[1],
            Self::ReplayWindow => DEFAULT_CHUNK_BYTES[2],
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::One => "one",
            Self::Fixed64 => "64",
            Self::ReplayWindow => "window",
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq, Hash)]
#[serde(rename_all = "snake_case")]
pub enum TextMode {
    /// Empty text events, retaining the paragraph envelope only.
    Empty,
    /// A short deterministic text payload containing XML-sensitive characters.
    Short,
    /// A fixed bounded payload that is independent of authored count.
    NearLimit,
}

impl TextMode {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "empty" => Ok(Self::Empty),
            "short" => Ok(Self::Short),
            "near" | "near_limit" => Ok(Self::NearLimit),
            _ => Err(format!("invalid text mode {value:?}; expected empty, short, or near").into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Empty => "empty",
            Self::Short => "short",
            Self::NearLimit => "near",
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    source_counts: Vec<usize>,
    authored_counts: Vec<usize>,
    chunk_modes: Vec<ChunkMode>,
    text_modes: Vec<TextMode>,
    samples: usize,
    warmups: usize,
    sink_write_bytes: usize,
    json_path: Option<PathBuf>,
    fixture_dir: Option<PathBuf>,
}

#[derive(Clone, Debug, Serialize)]
struct BinaryRecord {
    binary: &'static str,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Clone, Debug, Serialize)]
struct ConfigRecord {
    source_counts: Vec<usize>,
    authored_counts: Vec<usize>,
    chunk_modes: Vec<ChunkMode>,
    text_modes: Vec<TextMode>,
    samples: usize,
    warmups: usize,
    sink_write_bytes: usize,
    lifecycle: [&'static str; 4],
    expected_authored_opens: u64,
    source: &'static str,
    authored_provider: &'static str,
    sink: &'static str,
    fixture_dir: Option<String>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct MemberIdentity {
    path: String,
    compression_method: String,
    data_descriptor: bool,
    crc32: u32,
    decoded_bytes: usize,
    decoded_sha256: String,
    compressed_bytes: u64,
    compressed_sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct SemanticRecord {
    paragraph_count: usize,
    order_sha256: String,
    text_sha256: String,
    text_bytes: usize,
}

#[derive(Clone, Debug, Serialize)]
struct SourceRecord {
    archive_bytes: usize,
    archive_sha256: String,
    main_xml_bytes: usize,
    main_xml_sha256: String,
    member_count: usize,
    semantic: SemanticRecord,
    members: Vec<MemberIdentity>,
    unchanged_oracle: bool,
    opaque_member_exact: bool,
}

#[derive(Clone, Debug, Serialize)]
struct AuthoredRecord {
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
    max_chunk_bytes: u64,
    replay_window_bytes: u64,
    max_encoded_paragraph_bytes: u64,
    xml_entity_reference_count: u64,
    text_bytes: u64,
    encoded_xml_bytes: u64,
    event_count: u64,
    expected_event_sha256: String,
    expected_encoded_sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct StreamLimitRecord {
    parser_event_limit: u64,
    parser_token_bytes: u64,
    parser_workspace_bytes: u64,
    max_xml_depth: u64,
    max_authored_chunk_bytes: u64,
    replay_window_bytes: u64,
}

#[derive(Clone, Debug, Serialize)]
struct ProofRecord {
    source_len: u64,
    source_sha256: String,
    source_paragraph_count: u64,
    source_event_count: u64,
    insertion_offset: u64,
    candidate_len: u64,
    candidate_sha256: String,
    candidate_paragraph_count: u64,
    candidate_event_count: u64,
    generated_offset: u64,
    generated_once: bool,
    authored: AuthoredRecord,
}

#[derive(Clone, Debug, Serialize)]
struct OracleRecord {
    candidate_archive_bytes: usize,
    candidate_archive_sha256: String,
    candidate_main_xml_bytes: usize,
    candidate_main_xml_sha256: String,
    candidate_semantic: SemanticRecord,
    candidate_member_count: usize,
    candidate_xml_exact: bool,
    candidate_semantic_exact: bool,
    untouched_member_metadata_exact: bool,
    untouched_raw_members_preserved: bool,
    physical_order_exact: bool,
    opaque_member_exact: bool,
    source_unchanged: bool,
    inverse_exact: bool,
}

#[derive(Clone, Debug, Serialize)]
struct Sample {
    sample: usize,
    source_count: usize,
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
    elapsed_ns: u64,
    source_reads: ReadObservation,
    authored: AuthoredObservation,
    sink: SinkRecord,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Clone, Debug, Serialize)]
struct CaseRecord {
    source_count: usize,
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
    source: SourceRecord,
    authored: AuthoredRecord,
    limits: StreamLimitRecord,
    oracle: OracleRecord,
    proof: ProofRecord,
    samples: Vec<Sample>,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    version: u32,
    binary: BinaryRecord,
    config: ConfigRecord,
    cases: Vec<CaseRecord>,
}

#[derive(Debug, Clone, Copy, Default, Serialize, PartialEq, Eq)]
struct ReadObservation {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    request_histogram: RequestHistogram,
    returned_histogram: RequestHistogram,
}

#[derive(Clone, Copy, Debug, Default, Serialize, PartialEq, Eq)]
struct RequestHistogram {
    bytes_0: u64,
    bytes_1_to_512: u64,
    bytes_513_to_4096: u64,
    bytes_4097_to_16384: u64,
    bytes_16385_to_65536: u64,
    bytes_over_65536: u64,
}

#[derive(Debug, Default)]
struct AtomicRequestHistogram {
    bytes_0: AtomicU64,
    bytes_1_to_512: AtomicU64,
    bytes_513_to_4096: AtomicU64,
    bytes_4097_to_16384: AtomicU64,
    bytes_16385_to_65536: AtomicU64,
    bytes_over_65536: AtomicU64,
}

impl AtomicRequestHistogram {
    fn record(&self, bytes: usize) {
        let counter = match bytes {
            0 => &self.bytes_0,
            1..=512 => &self.bytes_1_to_512,
            513..=4_096 => &self.bytes_513_to_4096,
            4_097..=16_384 => &self.bytes_4097_to_16384,
            16_385..=65_536 => &self.bytes_16385_to_65536,
            _ => &self.bytes_over_65536,
        };
        counter.fetch_add(1, Ordering::Relaxed);
    }

    fn snapshot(&self) -> RequestHistogram {
        RequestHistogram {
            bytes_0: self.bytes_0.load(Ordering::Relaxed),
            bytes_1_to_512: self.bytes_1_to_512.load(Ordering::Relaxed),
            bytes_513_to_4096: self.bytes_513_to_4096.load(Ordering::Relaxed),
            bytes_4097_to_16384: self.bytes_4097_to_16384.load(Ordering::Relaxed),
            bytes_16385_to_65536: self.bytes_16385_to_65536.load(Ordering::Relaxed),
            bytes_over_65536: self.bytes_over_65536.load(Ordering::Relaxed),
        }
    }
}

#[derive(Debug, Default)]
struct SourceCounters {
    read_calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    request_histogram: AtomicRequestHistogram,
    returned_histogram: AtomicRequestHistogram,
}

impl SourceCounters {
    fn snapshot(&self) -> ReadObservation {
        ReadObservation {
            calls: self.read_calls.load(Ordering::Relaxed),
            requested_bytes: self.requested_bytes.load(Ordering::Relaxed),
            returned_bytes: self.returned_bytes.load(Ordering::Relaxed),
            request_histogram: self.request_histogram.snapshot(),
            returned_histogram: self.returned_histogram.snapshot(),
        }
    }
}

#[derive(Debug)]
struct MeasureSource {
    bytes: Arc<[u8]>,
    id: u64,
    revision: AtomicU64,
    counters: Arc<SourceCounters>,
}

static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

impl MeasureSource {
    fn new(bytes: Arc<[u8]>, counters: Arc<SourceCounters>) -> Self {
        Self {
            bytes,
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            revision: AtomicU64::new(0),
            counters,
        }
    }

    fn observation(&self) -> ReadObservation {
        self.counters.snapshot()
    }
}

impl ReadAt for MeasureSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source length overflow"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.read_calls.fetch_add(1, Ordering::Relaxed);
        self.counters.requested_bytes.fetch_add(
            u64::try_from(output.len())
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "request overflow"))?,
            Ordering::Relaxed,
        );
        self.counters.request_histogram.record(output.len());
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        if output.is_empty() || offset >= self.bytes.len() {
            self.counters.returned_histogram.record(0);
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        self.counters.returned_bytes.fetch_add(
            u64::try_from(count)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "read overflow"))?,
            Ordering::Relaxed,
        );
        self.counters.returned_histogram.record(count);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            self.id,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[derive(Debug, Default)]
struct AuthoredCounters {
    opens: AtomicU64,
    events: AtomicU64,
    text_chunks: AtomicU64,
    text_bytes: AtomicU64,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct AuthoredObservation {
    opens: u64,
    events: u64,
    text_chunks: u64,
    text_bytes: u64,
}

impl AuthoredCounters {
    fn snapshot(&self) -> AuthoredObservation {
        AuthoredObservation {
            opens: self.opens.load(Ordering::Relaxed),
            events: self.events.load(Ordering::Relaxed),
            text_chunks: self.text_chunks.load(Ordering::Relaxed),
            text_bytes: self.text_bytes.load(Ordering::Relaxed),
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct AuthoredSpec {
    count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
}

#[derive(Debug)]
struct GeneratedParagraphSource {
    spec: AuthoredSpec,
    counters: Arc<AuthoredCounters>,
}

impl GeneratedParagraphSource {
    fn new(spec: AuthoredSpec, counters: Arc<AuthoredCounters>) -> Self {
        Self { spec, counters }
    }
}

#[derive(Debug)]
struct GeneratedCursor {
    spec: AuthoredSpec,
    counters: Arc<AuthoredCounters>,
    paragraph: usize,
    phase: CursorPhase,
    text_offset: usize,
    text_length: usize,
    text: [u8; MAX_CURSOR_TEXT_BYTES],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CursorPhase {
    Start,
    Text,
    End,
    Done,
}

#[derive(Debug)]
struct CursorError;

impl std::fmt::Display for CursorError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("generated authored cursor failed")
    }
}

impl std::error::Error for CursorError {}

impl ReplayableParagraphSource for GeneratedParagraphSource {
    type Error = CursorError;
    type Cursor<'source> = GeneratedCursor;

    fn open<'source>(&'source self) -> Result<Self::Cursor<'source>, Self::Error> {
        self.counters.opens.fetch_add(1, Ordering::Relaxed);
        Ok(GeneratedCursor {
            spec: self.spec,
            counters: Arc::clone(&self.counters),
            paragraph: 0,
            phase: if self.spec.count == 0 {
                CursorPhase::Done
            } else {
                CursorPhase::Start
            },
            text_offset: 0,
            text_length: 0,
            text: [0; MAX_CURSOR_TEXT_BYTES],
        })
    }
}

impl ParagraphCursor for GeneratedCursor {
    type Error = CursorError;

    fn next<'event>(&'event mut self) -> Result<Option<PlainParagraphEvent<'event>>, Self::Error> {
        let event = match self.phase {
            CursorPhase::Start => {
                self.phase = CursorPhase::Text;
                self.text_offset = 0;
                self.text_length = fill_text(
                    self.paragraph,
                    self.spec.count,
                    self.spec.text_mode,
                    &mut self.text,
                );
                PlainParagraphEvent::ParagraphStart
            },
            CursorPhase::Text => {
                if self.text_length == 0 {
                    self.phase = if self.paragraph.saturating_add(1) >= self.spec.count {
                        CursorPhase::Done
                    } else {
                        self.paragraph += 1;
                        CursorPhase::Start
                    };
                    PlainParagraphEvent::ParagraphEnd
                } else {
                    let chunk = self.spec.chunk_mode.chunk_bytes();
                    let end = self.text_offset.saturating_add(if chunk == 0 {
                        self.text_length
                    } else {
                        chunk.min(self.text_length.saturating_sub(self.text_offset))
                    });
                    let text = std::str::from_utf8(&self.text[self.text_offset..end])
                        .map_err(|_| CursorError)?;
                    self.text_offset = end;
                    self.counters.events.fetch_add(1, Ordering::Relaxed);
                    self.counters.text_chunks.fetch_add(1, Ordering::Relaxed);
                    self.counters.text_bytes.fetch_add(
                        u64::try_from(text.len()).map_err(|_| CursorError)?,
                        Ordering::Relaxed,
                    );
                    if end >= self.text_length {
                        self.phase = CursorPhase::End;
                    }
                    PlainParagraphEvent::TextChunk(text)
                }
            },
            CursorPhase::End => {
                self.phase = if self.paragraph.saturating_add(1) >= self.spec.count {
                    CursorPhase::Done
                } else {
                    self.paragraph += 1;
                    CursorPhase::Start
                };
                PlainParagraphEvent::ParagraphEnd
            },
            CursorPhase::Done => return Ok(None),
        };
        if !matches!(event, PlainParagraphEvent::TextChunk(_)) {
            self.counters.events.fetch_add(1, Ordering::Relaxed);
        }
        Ok(Some(event))
    }
}

fn fill_text(
    index: usize,
    count: usize,
    mode: TextMode,
    output: &mut [u8; MAX_CURSOR_TEXT_BYTES],
) -> usize {
    match mode {
        TextMode::Empty => 0,
        TextMode::Short | TextMode::NearLimit => {
            let mut position = 0;
            position += copy_bytes(output, position, b"authored-");
            position += write_decimal(output, position, index, 8);
            position += copy_bytes(output, position, b" <&>");
            if matches!(mode, TextMode::NearLimit) {
                let target = text_bytes(index, count, mode);
                let filler = b" near<&> plain";
                while position < target {
                    let take = filler.len().min(target - position);
                    position += copy_bytes(output, position, &filler[..take]);
                }
            }
            position
        },
    }
}

fn copy_bytes(output: &mut [u8], position: usize, bytes: &[u8]) -> usize {
    let available = output.len().saturating_sub(position);
    let take = available.min(bytes.len());
    output[position..position + take].copy_from_slice(&bytes[..take]);
    take
}

fn write_decimal(output: &mut [u8], position: usize, value: usize, width: usize) -> usize {
    let mut digits = [b'0'; 20];
    let mut end = digits.len();
    let mut value = value;
    loop {
        end -= 1;
        digits[end] = b'0' + u8::try_from(value % 10).unwrap_or(0);
        value /= 10;
        if value == 0 {
            break;
        }
    }
    let digit_len = digits.len() - end;
    let padding = width.saturating_sub(digit_len);
    let mut written = 0;
    for _ in 0..padding {
        written += copy_bytes(output, position + written, b"0");
    }
    written + copy_bytes(output, position + written, &digits[end..])
}

fn decimal_digits(mut value: usize) -> usize {
    let mut digits = 1;
    while value >= 10 {
        value /= 10;
        digits += 1;
    }
    digits
}

fn authored_prefix_bytes(index: usize) -> usize {
    9 + decimal_digits(index).max(8) + 4
}

fn near_limit_text_bytes() -> usize {
    // Keep the text-size axis fixed; authored count changes the number of
    // paragraphs, never the payload selected for one paragraph.
    MAX_CURSOR_TEXT_BYTES
}

fn text_bytes(index: usize, _count: usize, mode: TextMode) -> usize {
    match mode {
        TextMode::Empty => 0,
        TextMode::Short => authored_prefix_bytes(index),
        TextMode::NearLimit => near_limit_text_bytes(),
    }
}

fn is_xml_entity_reference_byte(byte: u8) -> bool {
    matches!(byte, b'&' | b'<' | b'>' | b'"' | b'\'')
}

fn escaped_text_bytes(index: usize, count: usize, mode: TextMode) -> usize {
    let length = text_bytes(index, count, mode);
    match mode {
        TextMode::Empty => 0,
        TextMode::Short | TextMode::NearLimit => {
            // Every `<` and `&` in the deterministic payload is escaped.  The
            // near-limit filler repeats `near<&> plain` and may end partway
            // through one repetition, so count it directly from the bounded
            // generator rather than relying on a fragile closed form.
            let mut buffer = [0_u8; MAX_CURSOR_TEXT_BYTES];
            let actual = fill_text(index, count, mode, &mut buffer);
            buffer[..actual]
                .iter()
                .map(|byte| match byte {
                    b'&' => 5,
                    b'<' | b'>' | b'"' | b'\'' => 4,
                    _ => 1,
                })
                .sum()
        },
    }
    .max(length)
}

fn authored_record(spec: AuthoredSpec) -> BenchResult<AuthoredRecord> {
    let mut total_text_bytes = 0_u64;
    let mut encoded_xml_bytes = 0_u64;
    let mut max_encoded_paragraph_bytes = 0_u64;
    let mut xml_entity_reference_count = 0_u64;
    let mut event_count = 0_u64;
    let mut event_hash = Sha256::new();
    let mut encoded_hash = Sha256::new();
    for index in 0..spec.count {
        event_count = event_count
            .checked_add(1)
            .ok_or("authored event overflow")?;
        event_hash.update([0_u8]);
        let bytes = text_bytes(index, spec.count, spec.text_mode);
        let escaped = escaped_text_bytes(index, spec.count, spec.text_mode);
        total_text_bytes = total_text_bytes
            .checked_add(u64::try_from(bytes)?)
            .ok_or("authored text-byte overflow")?;
        let mut text = [0_u8; MAX_CURSOR_TEXT_BYTES];
        let actual = fill_text(index, spec.count, spec.text_mode, &mut text);
        let entity_references = text[..actual]
            .iter()
            .filter(|byte| is_xml_entity_reference_byte(**byte))
            .count();
        xml_entity_reference_count = xml_entity_reference_count
            .checked_add(u64::try_from(entity_references)?)
            .ok_or("authored XML entity-reference overflow")?;
        if actual != 0 {
            let chunk = spec.chunk_mode.chunk_bytes();
            let mut offset = 0;
            while offset < actual {
                let width = if chunk == 0 {
                    actual.saturating_sub(offset)
                } else {
                    chunk.min(actual.saturating_sub(offset))
                };
                let end = offset.saturating_add(width);
                let part = &text[offset..end];
                event_count = event_count
                    .checked_add(1)
                    .ok_or("authored event overflow")?;
                event_hash.update([1_u8]);
                event_hash.update(u64::try_from(part.len())?.to_le_bytes());
                event_hash.update(part);
                offset = end;
            }
        }
        event_count = event_count
            .checked_add(1)
            .ok_or("authored event overflow")?;
        event_hash.update([2_u8]);
        let paragraph_bytes = paragraph_xml_bytes(index, spec.count, spec.text_mode)?;
        encoded_xml_bytes = encoded_xml_bytes
            .checked_add(u64::try_from(paragraph_bytes.len())?)
            .ok_or("authored XML-byte overflow")?;
        max_encoded_paragraph_bytes =
            max_encoded_paragraph_bytes.max(u64::try_from(paragraph_bytes.len())?);
        encoded_hash.update(&paragraph_bytes);
        if actual != bytes || escaped != paragraph_bytes.len().saturating_sub(envelope_bytes()) {
            return Err("authored generator accounting disagrees with XML oracle".into());
        }
    }
    Ok(AuthoredRecord {
        authored_count: spec.count,
        chunk_mode: spec.chunk_mode,
        text_mode: spec.text_mode,
        max_chunk_bytes: u64::try_from(spec_chunk_limit(spec))?,
        replay_window_bytes: u64::try_from(replay_window_for(spec))?,
        max_encoded_paragraph_bytes,
        xml_entity_reference_count,
        text_bytes: total_text_bytes,
        encoded_xml_bytes,
        event_count,
        expected_event_sha256: hex_digest(&event_hash.finalize()),
        expected_encoded_sha256: hex_digest(&encoded_hash.finalize()),
    })
}

fn spec_chunk_limit(spec: AuthoredSpec) -> usize {
    let maximum_text = (0..spec.count)
        .map(|index| text_bytes(index, spec.count, spec.text_mode))
        .max()
        .unwrap_or(0);
    let configured = match spec.chunk_mode {
        ChunkMode::One => maximum_text,
        ChunkMode::Fixed64 => DEFAULT_CHUNK_BYTES[1].min(maximum_text),
        ChunkMode::ReplayWindow => DEFAULT_CHUNK_BYTES[2].min(maximum_text),
    };
    configured.max(1)
}

fn replay_window_for(spec: AuthoredSpec) -> usize {
    let chunk = spec_chunk_limit(spec);
    let required = chunk.saturating_mul(6).saturating_add(1024);
    required.max(64 * 1024)
}

fn paragraph_tag_count(xml: &[u8]) -> usize {
    xml.windows(b"<w:p".len())
        .filter(|window| *window == b"<w:p")
        .count()
}

fn envelope_bytes() -> usize {
    b"<w:p xmlns:w=\"".len()
        + WORD_NAMESPACE.len()
        + b"\"><w:r><w:t xml:space=\"preserve\"></w:t></w:r></w:p>".len()
}

fn paragraph_xml_bytes(index: usize, count: usize, mode: TextMode) -> BenchResult<Vec<u8>> {
    let mut text = [0_u8; MAX_CURSOR_TEXT_BYTES];
    let text_len = fill_text(index, count, mode, &mut text);
    let escaped_len = escaped_text_bytes(index, count, mode);
    let mut output = Vec::with_capacity(envelope_bytes().saturating_add(escaped_len));
    output.extend_from_slice(b"<w:p xmlns:w=\"");
    output.extend_from_slice(WORD_NAMESPACE.as_bytes());
    output.extend_from_slice(b"\"><w:r><w:t xml:space=\"preserve\">");
    for byte in &text[..text_len] {
        match *byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            value => output.push(value),
        }
    }
    output.extend_from_slice(b"</w:t></w:r></w:p>");
    Ok(output)
}

fn source_main_xml(count: usize) -> BenchResult<Vec<u8>> {
    let mut xml = String::with_capacity(count.saturating_mul(58).saturating_add(192));
    write!(
        &mut xml,
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{WORD_NAMESPACE}"><w:body>"#
    )?;
    for index in 0..count {
        write!(
            &mut xml,
            "<w:p><w:r><w:t>source-{index:08}</w:t></w:r></w:p>"
        )?;
    }
    xml.push_str("</w:body></w:document>");
    Ok(xml.into_bytes())
}

fn independent_candidate_xml(source_xml: &[u8], authored: AuthoredSpec) -> BenchResult<Vec<u8>> {
    let close = b"</w:body>";
    let body_end = source_xml
        .windows(close.len())
        .position(|window| window == close)
        .ok_or("source XML is missing body close")?;
    let authored_bytes = usize::try_from(authored_record(authored)?.encoded_xml_bytes)?;
    let mut output = Vec::with_capacity(source_xml.len().saturating_add(authored_bytes));
    output.extend_from_slice(&source_xml[..body_end]);
    for index in 0..authored.count {
        output.extend_from_slice(&paragraph_xml_bytes(
            index,
            authored.count,
            authored.text_mode,
        )?);
    }
    output.extend_from_slice(&source_xml[body_end..]);
    Ok(output)
}

fn opaque_payload(variant: u8) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(OPAQUE_BYTES);
    for index in 0..OPAQUE_BYTES {
        let low = (index & 0xff) as u8;
        let page = ((index >> 8) & 0xff) as u8;
        bytes.push(
            low.wrapping_mul(37)
                .wrapping_add(page)
                .wrapping_add(variant),
        );
    }
    bytes
}

fn source_archive(main_xml: &[u8], opaque_variant: u8) -> BenchResult<Vec<u8>> {
    let mut package = OpcPackage::new();
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(format!("/{MAIN_PATH}"))?,
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main_xml.to_vec(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(format!("/{OPAQUE_PATH}"))?,
        "application/octet-stream".to_owned(),
        opaque_payload(opaque_variant),
    )))?;
    package.relate_to(MAIN_PATH, rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

/// Mirror the source-backed scanner's checked owner envelope for one parser
/// pass.  The token window is deliberately derived from the generated
/// paragraph, while the depth remains the finite stream policy; keeping this
/// value explicit prevents near-limit text from tripping an unrelated fixed
/// two-megabyte workspace ceiling.
fn scanner_workspace_limit(max_token_bytes: u64, max_depth: u64) -> BenchResult<u64> {
    const MAX_NAMESPACE_DECLARATIONS_PER_LEVEL: usize = 256;
    const INITIAL_NAMESPACE_BYTES: usize = 73;
    let token = usize::try_from(max_token_bytes)?
        .checked_add(1)
        .ok_or("scanner token workspace overflow")?;
    let levels = usize::try_from(max_depth)?
        .checked_add(1)
        .ok_or("scanner depth workspace overflow")?;
    let namespace_declarations = MAX_NAMESPACE_DECLARATIONS_PER_LEVEL.min(token);
    let attributes = token;
    let usize_bytes = size_of::<usize>();
    let binding_bytes = usize_bytes
        .checked_mul(4)
        .ok_or("scanner binding workspace overflow")?;
    let range_bytes = size_of::<Range<usize>>();
    let frame_bytes = size_of::<u8>();
    let hash_entry_bytes = size_of::<u64>()
        .checked_add(size_of::<u8>())
        .ok_or("scanner hash workspace overflow")?;
    let seen_attribute_bytes = size_of::<(&[u8], &[u8])>();
    let token_windows = token
        .checked_mul(4)
        .ok_or("scanner token workspace overflow")?;
    let opened_names = levels
        .checked_mul(token)
        .and_then(|value| value.checked_mul(2))
        .ok_or("scanner opened-name workspace overflow")?
        .max(8);
    let opened_indexes = 4usize
        .checked_mul(usize_bytes)
        .ok_or("scanner index workspace overflow")?
        .max(
            2usize
                .checked_mul(levels)
                .and_then(|value| value.checked_mul(usize_bytes))
                .ok_or("scanner index workspace overflow")?,
        );
    let namespace_bytes = INITIAL_NAMESPACE_BYTES
        .checked_add(
            levels
                .checked_mul(token)
                .ok_or("scanner namespace workspace overflow")?,
        )
        .and_then(|value| value.checked_mul(2))
        .ok_or("scanner namespace workspace overflow")?
        .max(128);
    let namespace_binding_floor = 8usize
        .checked_mul(binding_bytes)
        .ok_or("scanner namespace binding overflow")?;
    let namespace_bindings = 2usize
        .checked_add(
            levels
                .checked_mul(namespace_declarations)
                .ok_or("scanner namespace binding overflow")?,
        )
        .and_then(|value| value.checked_mul(2))
        .and_then(|value| value.checked_mul(binding_bytes))
        .ok_or("scanner namespace binding overflow")?
        .max(namespace_binding_floor);
    let attribute_ranges = attributes
        .checked_add(1)
        .and_then(|value| value.checked_mul(range_bytes))
        .and_then(|value| value.checked_mul(2))
        .ok_or("scanner attribute-range workspace overflow")?
        .max(
            4usize
                .checked_mul(range_bytes)
                .ok_or("scanner attribute-range workspace overflow")?,
        );
    let attribute_hash = attributes
        .checked_add(1)
        .and_then(|value| value.checked_mul(hash_entry_bytes))
        .and_then(|value| value.checked_mul(4))
        .ok_or("scanner attribute-hash workspace overflow")?
        .max(
            8usize
                .checked_mul(hash_entry_bytes)
                .ok_or("scanner attribute-hash workspace overflow")?,
        );
    let attribute_seen = attributes
        .checked_add(1)
        .and_then(|value| value.checked_mul(seen_attribute_bytes))
        .and_then(|value| value.checked_mul(2))
        .ok_or("scanner attribute-seen workspace overflow")?
        .max(
            8usize
                .checked_mul(seen_attribute_bytes)
                .ok_or("scanner attribute-seen workspace overflow")?,
        );
    let scope_stack = levels
        .checked_mul(frame_bytes)
        .and_then(|value| value.checked_mul(2))
        .ok_or("scanner scope workspace overflow")?
        .max(
            8usize
                .checked_mul(frame_bytes)
                .ok_or("scanner scope workspace overflow")?,
        );
    let namespace_scope_stack = levels
        .checked_mul(size_of::<(usize, usize)>())
        .and_then(|value| value.checked_mul(2))
        .ok_or("scanner namespace-scope workspace overflow")?;
    let total = [
        token_windows,
        opened_names,
        opened_indexes,
        namespace_bytes,
        namespace_bindings,
        attribute_ranges,
        attribute_hash,
        attribute_seen,
        scope_stack,
        namespace_scope_stack,
        3,
    ]
    .into_iter()
    .try_fold(0usize, usize::checked_add)
    .ok_or("scanner workspace overflow")?;
    Ok(u64::try_from(total)?)
}

fn stream_limits(
    source_xml_bytes: usize,
    source_count: usize,
    authored: &AuthoredRecord,
) -> BenchResult<ParagraphStreamLimits> {
    let source_xml = u64::try_from(source_xml_bytes)?;
    let authored_xml = authored.encoded_xml_bytes;
    let authored_text = authored.text_bytes;
    let source_text = u64::try_from(source_count)?
        .checked_mul(15)
        .ok_or("source text-byte limit overflow")?;
    let source_ceiling = source_xml
        .checked_add(2 * 1024 * 1024)
        .ok_or("source XML limit overflow")?;
    let candidate_xml = source_xml
        .checked_add(authored_xml)
        .and_then(|value| value.checked_add(2 * 1024 * 1024))
        .ok_or("candidate XML limit overflow")?;
    let source_plus_authored = source_count
        .checked_add(authored.authored_count)
        .and_then(|value| value.checked_add(16))
        .ok_or("paragraph limit overflow")?;
    let paragraph_limit = u64::try_from(source_plus_authored)?;
    let source_events = u64::try_from(source_count)?
        .checked_mul(8)
        .and_then(|value| value.checked_add(128))
        .ok_or("source event limit overflow")?;
    // `source.max_events` counts quick-xml events in the source and the
    // candidate stream, while `authored.event_count` counts borrowed input
    // events.  Each generated paragraph has six structural events.  For each
    // XML-sensitive input byte, quick-xml may emit one GeneralRef and a Text
    // event on either side of it, so `7 + 2 * references` bounds one non-empty
    // paragraph (and also bounds an empty one after rounding the base to
    // eight).  Count references from the bounded raw payload, independently
    // of authored chunk framing, then retain a fixed envelope margin.
    let authored_xml_events = u64::try_from(authored.authored_count)?
        .checked_mul(8)
        .and_then(|value| {
            authored
                .xml_entity_reference_count
                .checked_mul(2)
                .and_then(|references| value.checked_add(references))
        })
        .and_then(|value| value.checked_add(128))
        .ok_or("authored XML event limit overflow")?;
    let event_limit = source_events
        .checked_add(authored_xml_events)
        .ok_or("combined event limit overflow")?;
    let output_limit = source_xml
        .checked_add(authored_xml)
        .and_then(|value| value.checked_add(8 * 1024 * 1024))
        .ok_or("output limit overflow")?;
    let token_limit = authored
        .max_encoded_paragraph_bytes
        .checked_add(1024)
        .ok_or("XML token limit overflow")?
        .max(64 * 1024);
    let workspace_limit = scanner_workspace_limit(token_limit, MAX_STREAM_XML_DEPTH)?;
    let mut source = TailAppendLimits::new(
        source_ceiling,
        source_text
            .checked_add(authored_text)
            .and_then(|value| value.checked_add(1024))
            .ok_or("source text-byte limit overflow")?,
        authored_xml.saturating_add(1),
        candidate_xml,
        event_limit,
        MAX_STREAM_XML_DEPTH,
        paragraph_limit,
        4 * 1024 * 1024,
        workspace_limit,
        output_limit,
        token_limit,
    );
    // The source limit's text/fragment names predate the replay route.  Keep
    // source text admission independent from the authored-only ceiling while
    // retaining a separate authored proof limit below.
    source.max_fragment_bytes = authored_xml.saturating_add(1);
    let mut limits = ParagraphStreamLimits::new(source);
    limits.max_authored_paragraphs = u64::try_from(authored.authored_count)?;
    limits.max_authored_events = authored.event_count;
    limits.max_authored_chunk_bytes = authored.max_chunk_bytes;
    limits.max_authored_text_bytes = authored_text.saturating_add(1);
    limits.max_authored_xml_bytes = authored_xml.saturating_add(1);
    limits.max_replay_bytes = 1;
    limits.max_replay_window_bytes = authored.replay_window_bytes;
    limits.max_patch_bytes = 64 * 1024;
    limits.validate()?;
    Ok(limits)
}

fn hex_digest(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        let _ = write!(&mut output, "{byte:02x}");
    }
    output
}

fn sha256_hex(bytes: &[u8]) -> String {
    hex_digest(&Sha256::digest(bytes))
}

fn member_identities(bytes: &[u8]) -> BenchResult<Vec<MemberIdentity>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let reader = ArchiveReader::new(bytes)?;
    let mut records = Vec::new();
    for result in archive.entries() {
        let entry = result?;
        if entry.is_dir() {
            continue;
        }
        let path = entry.file_path().try_normalize()?.as_str().to_owned();
        let decoded = reader.read(&path)?;
        let zip_entry = archive.get_entry(entry.wayfinder())?;
        let (start, end) = zip_entry.compressed_data_range();
        let start = usize::try_from(start)?;
        let end = usize::try_from(end)?;
        let compressed = bytes
            .get(start..end)
            .ok_or_else(|| format!("compressed range for {path} is outside archive"))?;
        records.push(MemberIdentity {
            path,
            compression_method: format!("{:?}", entry.compression_method()),
            data_descriptor: entry.has_data_descriptor(),
            crc32: entry.crc32(),
            decoded_bytes: decoded.len(),
            decoded_sha256: sha256_hex(&decoded),
            compressed_bytes: u64::try_from(compressed.len())?,
            compressed_sha256: sha256_hex(compressed),
        });
    }
    Ok(records)
}

/// Capture each member's complete raw local and central-directory records.
/// The caller excludes the changed main member; every other local header,
/// payload, descriptor, timestamp, vendor field, and comment remains part of
/// the preservation oracle.
fn raw_member_records(bytes: &[u8]) -> BenchResult<BTreeMap<String, (Vec<u8>, Vec<u8>)>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut records = BTreeMap::new();
    for result in archive.entries() {
        let entry = result?;
        if entry.is_dir() {
            continue;
        }
        let path = entry.file_path().try_normalize()?.as_str().to_owned();
        let zip_entry = archive.get_entry(entry.wayfinder())?;
        let local_start = usize::try_from(entry.local_header_offset())?;
        let (_, compressed_end) = zip_entry.compressed_data_range();
        let local_payload_end = usize::try_from(compressed_end)?;
        let descriptor_bytes = if entry.has_data_descriptor() {
            let descriptor = bytes
                .get(local_payload_end..)
                .ok_or("ZIP data descriptor starts outside archive")?;
            let with_signature = descriptor.starts_with(b"PK\x07\x08");
            if entry.is_zip64() {
                if with_signature { 24 } else { 20 }
            } else if with_signature {
                16
            } else {
                12
            }
        } else {
            0
        };
        let local_end = local_payload_end
            .checked_add(descriptor_bytes)
            .ok_or("ZIP local record end overflows usize")?;
        let central_start = usize::try_from(entry.central_directory_offset())?;
        let central_size = usize::try_from(entry.metadata_size_hint())?
            .checked_add(46)
            .ok_or("ZIP central record size overflows usize")?;
        let central_end = central_start
            .checked_add(central_size)
            .ok_or("ZIP central record end overflows usize")?;
        let local = bytes
            .get(local_start..local_end)
            .ok_or_else(|| format!("raw local ZIP record for {path} is outside archive"))?;
        let central = bytes
            .get(central_start..central_end)
            .ok_or_else(|| format!("raw central ZIP record for {path} is outside archive"))?;
        records.insert(path, (local.to_vec(), central.to_vec()));
    }
    Ok(records)
}

fn central_record_without_relocation(record: &[u8]) -> Vec<u8> {
    let mut normalized = record.to_vec();
    // Growth of the changed main member moves later local headers.  The
    // central-directory local-header offset is the sole relocation field that
    // may differ; every other central byte remains independently authenticated.
    if normalized.len() >= 46 {
        normalized[42..46].fill(0);
    }
    normalized
}

fn untouched_raw_members_equal(
    source: &BTreeMap<String, (Vec<u8>, Vec<u8>)>,
    candidate: &BTreeMap<String, (Vec<u8>, Vec<u8>)>,
) -> bool {
    source
        .iter()
        .filter(|(path, _)| path.as_str() != MAIN_PATH)
        .all(|(path, (source_local, source_central))| {
            candidate
                .get(path)
                .is_some_and(|(candidate_local, candidate_central)| {
                    source_local == candidate_local
                        && central_record_without_relocation(source_central)
                            == central_record_without_relocation(candidate_central)
                })
        })
}

fn find_member<'a>(members: &'a [MemberIdentity], path: &str) -> BenchResult<&'a MemberIdentity> {
    members
        .iter()
        .find(|member| member.path == path)
        .ok_or_else(|| format!("DOCX member {path:?} is missing").into())
}

fn semantic_record(bytes: &[u8]) -> BenchResult<(SemanticRecord, Vec<String>)> {
    let package = OwnedDocxPackage::from_reader(Cursor::new(bytes.to_vec()))?;
    let document = package.document_snapshot()?;
    let paragraphs = document.paragraphs();
    let mut texts = Vec::with_capacity(paragraphs.len());
    let mut order = Sha256::new();
    let mut text = Sha256::new();
    order.update(b"litchi-docx-replayable-tail-order-v1\0");
    text.update(b"litchi-docx-replayable-tail-text-v1\0");
    order.update(u64::try_from(paragraphs.len())?.to_le_bytes());
    text.update(u64::try_from(paragraphs.len())?.to_le_bytes());
    let mut text_bytes = 0usize;
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let value = paragraph.text()?;
        text_bytes = text_bytes
            .checked_add(value.len())
            .ok_or("DOCX semantic text-byte count overflow")?;
        order.update(u64::try_from(index)?.to_le_bytes());
        order.update(u64::try_from(value.len())?.to_le_bytes());
        order.update(value.as_bytes());
        text.update(u64::try_from(value.len())?.to_le_bytes());
        text.update(value.as_bytes());
        texts.push(value);
    }
    Ok((
        SemanticRecord {
            paragraph_count: texts.len(),
            order_sha256: hex_digest(&order.finalize()),
            text_sha256: hex_digest(&text.finalize()),
            text_bytes,
        },
        texts,
    ))
}

fn source_record(
    source_archive: &[u8],
    source_xml: &[u8],
    source_count: usize,
) -> BenchResult<SourceRecord> {
    let members = member_identities(source_archive)?;
    let reader = ArchiveReader::new(source_archive)?;
    let main = reader.read(MAIN_PATH)?;
    let opaque = reader.read(OPAQUE_PATH)?;
    let (semantic, texts) = semantic_record(source_archive)?;
    let source_guard = source_archive.to_vec();
    let unchanged_oracle = source_archive == source_guard.as_slice()
        && main == source_xml
        && semantic.paragraph_count == source_count
        && texts
            .iter()
            .enumerate()
            .all(|(index, value)| value == &format!("source-{index:08}"));
    let opaque_member = find_member(&members, OPAQUE_PATH)?;
    let opaque_member_exact =
        opaque_member.decoded_bytes == OPAQUE_BYTES && opaque == opaque_payload(0);
    Ok(SourceRecord {
        archive_bytes: source_archive.len(),
        archive_sha256: sha256_hex(source_archive),
        main_xml_bytes: source_xml.len(),
        main_xml_sha256: sha256_hex(source_xml),
        member_count: members.len(),
        semantic,
        members,
        unchanged_oracle,
        opaque_member_exact,
    })
}

#[derive(Clone, Copy, Debug, Default, Serialize, PartialEq, Eq)]
struct SinkHistogram {
    bytes_0: u64,
    bytes_1_to_512: u64,
    bytes_513_to_4096: u64,
    bytes_4097_to_16384: u64,
    bytes_16385_to_65536: u64,
    bytes_over_65536: u64,
}

#[derive(Debug)]
struct HashingSink {
    max_write: usize,
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    histogram: SinkHistogram,
    digest: Sha256,
}

#[derive(Clone, Copy, Debug)]
struct SinkObservation {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    histogram: SinkHistogram,
    digest: [u8; 32],
}

impl HashingSink {
    fn new(max_write: usize) -> BenchResult<Self> {
        if max_write == 0 {
            return Err("sink write size must be nonzero".into());
        }
        Ok(Self {
            max_write,
            accepted_bytes: 0,
            write_calls: 0,
            largest_write: 0,
            histogram: SinkHistogram::default(),
            digest: Sha256::new(),
        })
    }

    fn finish(self) -> SinkObservation {
        SinkObservation {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            largest_write: self.largest_write,
            histogram: self.histogram,
            digest: self.digest.finalize().into(),
        }
    }
}

impl Write for HashingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        let accepted = bytes.len().min(self.max_write);
        self.accepted_bytes = self
            .accepted_bytes
            .checked_add(
                u64::try_from(accepted)
                    .map_err(|_| io::Error::other("DOCX replay sink byte count overflow"))?,
            )
            .ok_or_else(|| io::Error::other("DOCX replay sink byte count overflow"))?;
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("DOCX replay sink call count overflow"))?;
        self.largest_write = self.largest_write.max(
            u64::try_from(accepted)
                .map_err(|_| io::Error::other("DOCX replay sink write size overflow"))?,
        );
        match accepted {
            0 => self.histogram.bytes_0 += 1,
            1..=512 => self.histogram.bytes_1_to_512 += 1,
            513..=4_096 => self.histogram.bytes_513_to_4096 += 1,
            4_097..=16_384 => self.histogram.bytes_4097_to_16384 += 1,
            16_385..=65_536 => self.histogram.bytes_16385_to_65536 += 1,
            _ => self.histogram.bytes_over_65536 += 1,
        }
        self.digest.update(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl SinkObservation {
    fn record(self) -> SinkRecord {
        SinkRecord {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            largest_write: self.largest_write,
            histogram: self.histogram,
            sha256: hex_digest(&self.digest),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct SinkRecord {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    histogram: SinkHistogram,
    sha256: String,
}

#[derive(Clone, Debug)]
struct Fixture {
    authored: AuthoredSpec,
    authored_record: AuthoredRecord,
    limits: ParagraphStreamLimits,
    limit_record: StreamLimitRecord,
    source_archive: Arc<[u8]>,
    candidate_archive: Arc<[u8]>,
    source_record: SourceRecord,
    oracle: OracleRecord,
    proof: ProofRecord,
}

#[derive(Clone, Debug, Serialize)]
struct FixtureArtifactRecord {
    file: String,
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct FixtureManifest {
    schema: &'static str,
    version: u32,
    source_count: usize,
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
    source: FixtureArtifactRecord,
    candidate: FixtureArtifactRecord,
    source_main_xml_sha256: String,
    candidate_main_xml_sha256: String,
}

fn fixture_stem(
    source_count: usize,
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
) -> String {
    format!(
        "source{source_count}-authored{authored_count}-{}-{}",
        chunk_mode.name(),
        text_mode.name()
    )
}

fn write_fixture_file(path: &Path, bytes: &[u8]) -> BenchResult<()> {
    let mut output = OpenOptions::new().write(true).create_new(true).open(path)?;
    output.write_all(bytes)?;
    output.flush()?;
    Ok(())
}

fn export_fixture(dir: &Path, fixture: &Fixture) -> BenchResult<()> {
    std::fs::create_dir_all(dir)?;
    let source_count = fixture.source_record.semantic.paragraph_count;
    let stem = fixture_stem(
        source_count,
        fixture.authored.count,
        fixture.authored.chunk_mode,
        fixture.authored.text_mode,
    );
    let source_file = format!("{stem}-source.docx");
    let candidate_file = format!("{stem}-candidate.docx");
    let manifest_file = format!("{stem}-hashes.json");
    let source_path = dir.join(&source_file);
    let candidate_path = dir.join(&candidate_file);
    let manifest_path = dir.join(&manifest_file);
    let source_sha256 = sha256_hex(fixture.source_archive.as_ref());
    let candidate_sha256 = sha256_hex(fixture.candidate_archive.as_ref());
    if source_sha256 != fixture.source_record.archive_sha256
        || candidate_sha256 != fixture.oracle.candidate_archive_sha256
    {
        return Err("fixture export archive hash disagrees with its oracle".into());
    }
    write_fixture_file(&source_path, fixture.source_archive.as_ref())?;
    write_fixture_file(&candidate_path, fixture.candidate_archive.as_ref())?;
    let manifest = FixtureManifest {
        schema: "docx-replayable-tail-append-fixture-v1",
        version: 1,
        source_count,
        authored_count: fixture.authored.count,
        chunk_mode: fixture.authored.chunk_mode,
        text_mode: fixture.authored.text_mode,
        source: FixtureArtifactRecord {
            file: source_file,
            bytes: fixture.source_archive.len(),
            sha256: source_sha256,
        },
        candidate: FixtureArtifactRecord {
            file: candidate_file,
            bytes: fixture.candidate_archive.len(),
            sha256: candidate_sha256,
        },
        source_main_xml_sha256: fixture.source_record.main_xml_sha256.clone(),
        candidate_main_xml_sha256: fixture.oracle.candidate_main_xml_sha256.clone(),
    };
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(manifest_path)?;
    serde_json::to_writer_pretty(&mut output, &manifest)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}

fn proof_record(
    source: litchi_docx::source_backed::tail_append::SourceProof,
    candidate: litchi_docx::source_backed::tail_append::CandidateProof,
    authored: AuthoredStreamProof,
    authored_spec: AuthoredSpec,
    expected_authored: &AuthoredRecord,
) -> ProofRecord {
    ProofRecord {
        source_len: source.source_len,
        source_sha256: hex_digest(&source.source_sha256),
        source_paragraph_count: source.paragraph_count,
        source_event_count: source.event_count,
        insertion_offset: source.insertion_offset,
        candidate_len: candidate.candidate_len,
        candidate_sha256: hex_digest(&candidate.candidate_sha256),
        candidate_paragraph_count: candidate.paragraph_count,
        candidate_event_count: candidate.event_count,
        generated_offset: candidate.generated_offset,
        generated_once: candidate.generated_once,
        authored: AuthoredRecord {
            authored_count: usize::try_from(authored.paragraph_count).unwrap_or(usize::MAX),
            chunk_mode: authored_spec.chunk_mode,
            text_mode: authored_spec.text_mode,
            max_chunk_bytes: u64::try_from(spec_chunk_limit(authored_spec)).unwrap_or(u64::MAX),
            replay_window_bytes: u64::try_from(replay_window_for(authored_spec))
                .unwrap_or(u64::MAX),
            max_encoded_paragraph_bytes: expected_authored.max_encoded_paragraph_bytes,
            xml_entity_reference_count: expected_authored.xml_entity_reference_count,
            text_bytes: authored.text_bytes,
            encoded_xml_bytes: authored.encoded_xml_bytes,
            event_count: authored.event_count,
            expected_event_sha256: hex_digest(&authored.event_sha256),
            expected_encoded_sha256: hex_digest(&authored.encoded_sha256),
        },
    }
}

fn verify_proof(
    proof: &ProofRecord,
    source_xml: &[u8],
    candidate_xml: &[u8],
    authored: &AuthoredRecord,
) -> BenchResult<()> {
    let body_end = source_xml
        .windows(b"</w:body>".len())
        .position(|window| window == b"</w:body>")
        .ok_or("source XML is missing body close")?;
    let expected_source_hash = sha256_hex(source_xml);
    let expected_candidate_hash = sha256_hex(candidate_xml);
    if proof.source_len != u64::try_from(source_xml.len())?
        || proof.source_sha256 != expected_source_hash
        || proof.insertion_offset != u64::try_from(body_end)?
        || proof.source_paragraph_count != u64::try_from(paragraph_tag_count(source_xml))?
        || proof.candidate_len != u64::try_from(candidate_xml.len())?
        || proof.candidate_sha256 != expected_candidate_hash
        || proof.candidate_paragraph_count
            != proof
                .source_paragraph_count
                .checked_add(u64::try_from(authored.authored_count)?)
                .ok_or("candidate paragraph count overflow")?
        || proof.generated_offset != u64::try_from(body_end)?
        || !proof.generated_once
        || proof.authored.authored_count != authored.authored_count
        || proof.authored.chunk_mode != authored.chunk_mode
        || proof.authored.text_mode != authored.text_mode
        || proof.authored.max_chunk_bytes != authored.max_chunk_bytes
        || proof.authored.replay_window_bytes != authored.replay_window_bytes
        || proof.authored.max_encoded_paragraph_bytes != authored.max_encoded_paragraph_bytes
        || proof.authored.xml_entity_reference_count != authored.xml_entity_reference_count
        || proof.authored.event_count != authored.event_count
        || proof.authored.text_bytes != authored.text_bytes
        || proof.authored.encoded_xml_bytes != authored.encoded_xml_bytes
        || proof.authored.expected_event_sha256 != authored.expected_event_sha256
        || proof.authored.expected_encoded_sha256 != authored.expected_encoded_sha256
    {
        return Err("DOCX replay proof disagrees with independent oracle".into());
    }
    Ok(())
}

fn verify_candidate_oracles(
    source_archive: &[u8],
    source_xml: &[u8],
    candidate_archive: &[u8],
    expected_candidate_xml: &[u8],
    source_count: usize,
    authored: AuthoredSpec,
    inverse: &[u8],
) -> BenchResult<OracleRecord> {
    let source_members = member_identities(source_archive)?;
    let candidate_members = member_identities(candidate_archive)?;
    let source_reader = ArchiveReader::new(source_archive)?;
    let candidate_reader = ArchiveReader::new(candidate_archive)?;
    let source_main = source_reader.read(MAIN_PATH)?;
    let candidate_main = candidate_reader.read(MAIN_PATH)?;
    let (candidate_semantic, candidate_texts) = semantic_record(candidate_archive)?;
    let expected_texts = {
        let (_, source_texts) = semantic_record(source_archive)?;
        let mut texts = source_texts;
        for index in 0..authored.count {
            let mut text = [0_u8; MAX_CURSOR_TEXT_BYTES];
            let length = fill_text(index, authored.count, authored.text_mode, &mut text);
            texts.push(String::from_utf8(text[..length].to_vec())?);
        }
        texts
    };
    let candidate_semantic_exact = candidate_texts == expected_texts;
    let candidate_xml_exact = source_main == source_xml && candidate_main == expected_candidate_xml;
    let untouched_member_metadata_exact = source_members
        .iter()
        .filter(|member| member.path != MAIN_PATH)
        .all(|member| {
            candidate_members
                .iter()
                .find(|value| value.path == member.path)
                == Some(member)
        });
    let source_raw_members = raw_member_records(source_archive)?;
    let candidate_raw_members = raw_member_records(candidate_archive)?;
    let untouched_raw_members_preserved =
        untouched_raw_members_equal(&source_raw_members, &candidate_raw_members);
    let physical_order_exact = source_members
        .iter()
        .map(|member| member.path.as_str())
        .eq(candidate_members.iter().map(|member| member.path.as_str()));
    let source_opaque = find_member(&source_members, OPAQUE_PATH)?;
    let candidate_opaque = find_member(&candidate_members, OPAQUE_PATH)?;
    let opaque_member_exact = source_opaque == candidate_opaque
        && source_reader.read(OPAQUE_PATH)? == candidate_reader.read(OPAQUE_PATH)?;
    let (_, source_texts) = semantic_record(source_archive)?;
    let source_unchanged = source_main == source_xml
        && source_texts.len() == source_count
        && source_texts
            .iter()
            .enumerate()
            .all(|(index, value)| value == &format!("source-{index:08}"));
    let inverse_exact = inverse == source_archive;
    if !candidate_xml_exact
        || !candidate_semantic_exact
        || !untouched_member_metadata_exact
        || !untouched_raw_members_preserved
        || !physical_order_exact
        || !opaque_member_exact
        || !source_unchanged
        || !inverse_exact
    {
        return Err("DOCX replay candidate or inverse oracle failed".into());
    }
    Ok(OracleRecord {
        candidate_archive_bytes: candidate_archive.len(),
        candidate_archive_sha256: sha256_hex(candidate_archive),
        candidate_main_xml_bytes: candidate_main.len(),
        candidate_main_xml_sha256: sha256_hex(&candidate_main),
        candidate_semantic,
        candidate_member_count: candidate_members.len(),
        candidate_xml_exact,
        candidate_semantic_exact,
        untouched_member_metadata_exact,
        untouched_raw_members_preserved,
        physical_order_exact,
        opaque_member_exact,
        source_unchanged,
        inverse_exact,
    })
}

fn build_fixture(source_count: usize, authored: AuthoredSpec) -> BenchResult<Fixture> {
    let source_xml = source_main_xml(source_count)?;
    let source_archive = source_archive(&source_xml, 0)?;
    let authored_record = authored_record(authored)?;
    let limits = stream_limits(source_xml.len(), source_count, &authored_record)?;
    let limit_record = StreamLimitRecord {
        parser_event_limit: limits.source.max_events,
        parser_token_bytes: limits.source.max_token_bytes,
        parser_workspace_bytes: limits.source.max_workspace_bytes,
        max_xml_depth: limits.source.max_depth,
        max_authored_chunk_bytes: limits.max_authored_chunk_bytes,
        replay_window_bytes: limits.max_replay_window_bytes,
    };
    let expected_candidate_xml = independent_candidate_xml(&source_xml, authored)?;
    let source_bytes: Arc<[u8]> = Arc::from(source_archive.clone());
    let counters = Arc::new(SourceCounters::default());
    let authored_counters = Arc::new(AuthoredCounters::default());
    let source = Arc::new(MeasureSource::new(Arc::clone(&source_bytes), counters));
    let package = source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
    let generated = GeneratedParagraphSource::new(authored, authored_counters);
    let edit = package.tail_append_plain_paragraphs(generated, limits);
    let plan = edit.prepare()?;
    let source_proof = plan.source_proof();
    let candidate_proof = plan.candidate_proof();
    let authored_proof = plan.authored_proof();
    let proof = proof_record(
        source_proof,
        candidate_proof,
        authored_proof,
        authored,
        &authored_record,
    );
    let mut output = Vec::new();
    let publication = plan.write_to_stream(&mut output)?;
    let candidate_archive = output;
    let candidate_bytes: Arc<[u8]> = Arc::from(candidate_archive.clone());
    let candidate_source = Arc::new(MeasureSource::new(
        Arc::clone(&candidate_bytes),
        Arc::new(SourceCounters::default()),
    ));
    let candidate_package =
        source_backed::Package::from_read_at(Arc::clone(&candidate_source) as Arc<dyn ReadAt>)?;
    let mut inverse = Vec::new();
    publication.write_inverse_to_stream(&candidate_package, &mut inverse)?;
    let oracle = verify_candidate_oracles(
        &source_archive,
        &source_xml,
        &candidate_archive,
        &expected_candidate_xml,
        source_count,
        authored,
        &inverse,
    )?;
    verify_proof(
        &proof,
        &source_xml,
        &expected_candidate_xml,
        &authored_record,
    )?;
    let source_record = source_record(&source_archive, &source_xml, source_count)?;
    if !source_record.unchanged_oracle || !source_record.opaque_member_exact {
        return Err("DOCX replay source oracle failed".into());
    }
    Ok(Fixture {
        authored,
        authored_record,
        limits,
        limit_record,
        source_archive: source_bytes,
        candidate_archive: candidate_bytes,
        source_record,
        oracle,
        proof,
    })
}

fn elapsed_ns(start: Instant) -> BenchResult<u64> {
    u64::try_from(start.elapsed().as_nanos())
        .map_err(|_| "DOCX replay elapsed time overflows u64 nanoseconds".into())
}

fn run_iteration(
    fixture: &Fixture,
    sink_write_bytes: usize,
) -> BenchResult<(
    u64,
    SinkObservation,
    ReadObservation,
    AuthoredObservation,
    Option<allocation_metrics::Sample>,
    Option<process_metrics::Delta>,
)> {
    let source_counters = Arc::new(SourceCounters::default());
    let authored_counters = Arc::new(AuthoredCounters::default());
    let process_before = process_metrics::Snapshot::read().ok();
    let region = allocation_metrics::begin();
    let start = Instant::now();
    let (sink, reads, authored) = {
        let source = Arc::new(MeasureSource::new(
            Arc::clone(&fixture.source_archive),
            Arc::clone(&source_counters),
        ));
        let package = source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
        let generated =
            GeneratedParagraphSource::new(fixture.authored, Arc::clone(&authored_counters));
        let edit = package.tail_append_plain_paragraphs(generated, fixture.limits);
        let plan = edit.prepare()?;
        let mut output = HashingSink::new(sink_write_bytes)?;
        let _publication = plan.write_to_stream(&mut output)?;
        let sink = output.finish();
        let reads = source.observation();
        let authored = authored_counters.snapshot();
        (sink, reads, authored)
    };
    let elapsed = elapsed_ns(start)?;
    let allocation = region.finish();
    let process = process_before
        .zip(process_metrics::Snapshot::read().ok())
        .map(|(before, after)| after.delta(before));
    Ok((elapsed, sink, reads, authored, allocation, process))
}

fn check_authored_counters(
    expected: &AuthoredRecord,
    authored: AuthoredObservation,
) -> BenchResult<()> {
    if authored.opens != EXPECTED_AUTHORED_OPENS {
        return Err(format!(
            "DOCX replay authored-provider opened {} passes; expected {EXPECTED_AUTHORED_OPENS}",
            authored.opens
        )
        .into());
    }
    let expected_events = expected
        .event_count
        .checked_mul(authored.opens)
        .ok_or("DOCX replay authored event counter overflow")?;
    let expected_text_bytes = expected
        .text_bytes
        .checked_mul(authored.opens)
        .ok_or("DOCX replay authored text counter overflow")?;
    if authored.events != expected_events || authored.text_bytes != expected_text_bytes {
        return Err(format!(
            "DOCX replay authored-provider counters disagree with proof: events {} != {expected_events} or text {} != {expected_text_bytes}",
            authored.events, authored.text_bytes
        )
        .into());
    }
    Ok(())
}

fn check_runtime(
    fixture: &Fixture,
    sink: SinkObservation,
    reads: ReadObservation,
    authored: AuthoredObservation,
) -> BenchResult<()> {
    if sink.accepted_bytes != u64::try_from(fixture.oracle.candidate_archive_bytes)?
        || sink.digest.as_slice() != Sha256::digest(fixture.candidate_archive.as_ref()).as_slice()
    {
        return Err("DOCX replay sink output differs from candidate archive oracle".into());
    }
    if reads.calls == 0 || reads.returned_bytes == 0 {
        return Err("DOCX replay lifecycle performed no positional source reads".into());
    }
    check_authored_counters(&fixture.authored_record, authored)?;
    Ok(())
}

fn run_case(
    source_count: usize,
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
    config: &Config,
) -> BenchResult<CaseRecord> {
    let authored_spec = AuthoredSpec {
        count: authored_count,
        chunk_mode,
        text_mode,
    };
    let fixture = build_fixture(source_count, authored_spec)?;
    if let Some(dir) = config.fixture_dir.as_deref() {
        // Fixture export is deliberate, opt-in, and outside every timed
        // iteration.  The files are the same bytes used by the independent
        // archive/oracle checks above.
        export_fixture(dir, &fixture)?;
    }
    let mut samples = Vec::with_capacity(config.samples);
    for iteration in 0..config.warmups.saturating_add(config.samples) {
        let (elapsed, sink, reads, authored, allocation, process) =
            run_iteration(&fixture, config.sink_write_bytes)?;
        check_runtime(&fixture, sink, reads, authored)?;
        if iteration >= config.warmups {
            samples.push(Sample {
                sample: iteration - config.warmups,
                source_count,
                authored_count,
                chunk_mode,
                text_mode,
                elapsed_ns: elapsed,
                source_reads: reads,
                authored,
                sink: sink.record(),
                allocation,
                process,
            });
        }
    }
    Ok(CaseRecord {
        source_count,
        authored_count,
        chunk_mode,
        text_mode,
        source: fixture.source_record,
        authored: fixture.authored_record,
        limits: fixture.limit_record,
        oracle: fixture.oracle,
        proof: fixture.proof,
        samples,
    })
}

fn parse_positive(value: &OsString, flag: &str) -> BenchResult<usize> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let parsed = text
        .parse::<usize>()
        .map_err(|error| format!("invalid {flag} value {text:?}: {error}"))?;
    if parsed == 0 {
        return Err(format!("{flag} must be positive").into());
    }
    Ok(parsed)
}

fn parse_counts(
    value: &OsString,
    flag: &str,
    allowed: Option<&[usize]>,
) -> BenchResult<Vec<usize>> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let mut values = Vec::new();
    for part in text.split(',') {
        let parsed = part
            .parse::<usize>()
            .map_err(|error| format!("invalid {flag} count {part:?}: {error}"))?;
        if parsed == 0 {
            return Err(format!("{flag} counts must be positive").into());
        }
        if let Some(allowed) = allowed
            && !allowed.contains(&parsed)
        {
            return Err(
                format!("unsupported {flag} count {parsed}; expected one of {allowed:?}").into(),
            );
        }
        if !values.contains(&parsed) {
            values.push(parsed);
        }
    }
    if values.is_empty() {
        return Err(format!("{flag} must contain at least one count").into());
    }
    Ok(values)
}

fn parse_list<T>(
    value: &OsString,
    flag: &str,
    parser: impl Fn(&str) -> BenchResult<T>,
) -> BenchResult<Vec<T>> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let mut values = Vec::new();
    for part in text.split(',') {
        values.push(parser(part)?);
    }
    if values.is_empty() {
        return Err(format!("{flag} must contain at least one value").into());
    }
    Ok(values)
}

fn next_value(values: &mut impl Iterator<Item = OsString>, flag: &str) -> BenchResult<OsString> {
    values
        .next()
        .ok_or_else(|| format!("missing value for {flag}").into())
}

fn parse_args<I>(args: I) -> BenchResult<Config>
where
    I: IntoIterator<Item = OsString>,
{
    let mut source_counts = DEFAULT_SOURCE_COUNTS.to_vec();
    let mut authored_counts = DEFAULT_AUTHORED_COUNTS.to_vec();
    let mut chunk_modes = vec![ChunkMode::One, ChunkMode::Fixed64, ChunkMode::ReplayWindow];
    let mut text_modes = vec![TextMode::Empty, TextMode::Short, TextMode::NearLimit];
    let mut samples = DEFAULT_SAMPLES;
    let mut warmups = DEFAULT_WARMUPS;
    let mut sink_write_bytes = HASH_SINK_MAX_WRITE;
    let mut json_path = None;
    let mut fixture_dir = None;
    let mut values = args.into_iter();
    while let Some(argument) = values.next() {
        let flag = argument.to_string_lossy();
        match flag.as_ref() {
            "--source-counts" | "--counts" => {
                source_counts = parse_counts(
                    &next_value(&mut values, "--source-counts")?,
                    "--source-counts",
                    Some(&DEFAULT_SOURCE_COUNTS),
                )?
            },
            "--authored-counts" => {
                authored_counts = parse_counts(
                    &next_value(&mut values, "--authored-counts")?,
                    "--authored-counts",
                    None,
                )?
            },
            "--chunks" => {
                chunk_modes = parse_list(
                    &next_value(&mut values, "--chunks")?,
                    "--chunks",
                    ChunkMode::parse,
                )?
            },
            "--text" | "--text-modes" => {
                text_modes = parse_list(
                    &next_value(&mut values, "--text")?,
                    "--text",
                    TextMode::parse,
                )?
            },
            "--samples" => {
                samples = parse_positive(&next_value(&mut values, "--samples")?, "--samples")?
            },
            "--warmups" => {
                warmups = parse_positive(&next_value(&mut values, "--warmups")?, "--warmups")?
            },
            "--sink-write" => {
                sink_write_bytes =
                    parse_positive(&next_value(&mut values, "--sink-write")?, "--sink-write")?
            },
            "--json" => json_path = Some(PathBuf::from(next_value(&mut values, "--json")?)),
            "--fixture-dir" => {
                fixture_dir = Some(PathBuf::from(next_value(&mut values, "--fixture-dir")?))
            },
            "--help" | "-h" => return Err(usage().into()),
            unknown => return Err(format!("unknown argument {unknown}; {}", usage()).into()),
        }
    }
    if source_counts.is_empty()
        || authored_counts.is_empty()
        || chunk_modes.is_empty()
        || text_modes.is_empty()
    {
        return Err("source/authored counts, chunk modes, and text modes must be nonempty".into());
    }
    Ok(Config {
        source_counts,
        authored_counts,
        chunk_modes,
        text_modes,
        samples,
        warmups,
        sink_write_bytes,
        json_path,
        fixture_dir,
    })
}

fn usage() -> &'static str {
    "usage: docx_replayable_tail_append [--source-counts 64,8192,131072] [--authored-counts 64,256,4096,16384] [--chunks one,64,window] [--text empty,short,near] [--samples N] [--warmups N] [--sink-write BYTES] [--json PATH] [--fixture-dir DIR]"
}

/// Run the replayable DOCX lifecycle benchmark using command-line arguments.
pub fn run_from_args<I>(args: I) -> BenchResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let config = parse_args(args)?;
    let capacity = config
        .source_counts
        .len()
        .checked_mul(config.authored_counts.len())
        .and_then(|value| value.checked_mul(config.chunk_modes.len()))
        .and_then(|value| value.checked_mul(config.text_modes.len()))
        .ok_or("DOCX replay case count overflow")?;
    let mut cases = Vec::with_capacity(capacity);
    for &source_count in &config.source_counts {
        for &authored_count in &config.authored_counts {
            for &chunk_mode in &config.chunk_modes {
                for &text_mode in &config.text_modes {
                    cases.push(run_case(
                        source_count,
                        authored_count,
                        chunk_mode,
                        text_mode,
                        &config,
                    )?);
                }
            }
        }
    }
    let report = Report {
        schema: SCHEMA,
        version: 1,
        binary: BinaryRecord {
            binary: allocation_metrics::binary_identity(),
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
        config: ConfigRecord {
            source_counts: config.source_counts.clone(),
            authored_counts: config.authored_counts.clone(),
            chunk_modes: config.chunk_modes.clone(),
            text_modes: config.text_modes.clone(),
            samples: config.samples,
            warmups: config.warmups,
            sink_write_bytes: config.sink_write_bytes,
            lifecycle: ["source_admission", "prepare", "publish", "drop"],
            expected_authored_opens: EXPECTED_AUTHORED_OPENS,
            source: "caller_owned_arc_positional_read_at_requested_returned_fixed_histograms",
            authored_provider: "deterministic_replayable_bounded_cursor",
            sink: "non_seek_hashing_sha256_short_write_no_archive_retention",
            fixture_dir: config
                .fixture_dir
                .as_ref()
                .map(|path| path.display().to_string()),
        },
        cases,
    };
    if let Some(path) = config.json_path {
        let mut output = OpenOptions::new().write(true).create_new(true).open(path)?;
        serde_json::to_writer_pretty(&mut output, &report)?;
        output.write_all(b"\n")?;
        output.flush()?;
    } else {
        serde_json::to_writer_pretty(io::stdout().lock(), &report)?;
        println!();
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn generated_text_and_proof_accounting_are_independent_of_chunking() {
        let spec_one = AuthoredSpec {
            count: 64,
            chunk_mode: ChunkMode::One,
            text_mode: TextMode::NearLimit,
        };
        let spec_chunks = AuthoredSpec {
            chunk_mode: ChunkMode::Fixed64,
            ..spec_one
        };
        let one = authored_record(spec_one).expect("one chunk accounting");
        let chunks = authored_record(spec_chunks).expect("fixed chunk accounting");
        assert_eq!(one.text_bytes, chunks.text_bytes);
        assert_eq!(one.encoded_xml_bytes, chunks.encoded_xml_bytes);
        assert_eq!(
            one.xml_entity_reference_count,
            chunks.xml_entity_reference_count
        );
        assert_ne!(one.event_count, chunks.event_count);
        assert_ne!(one.expected_event_sha256, chunks.expected_event_sha256);
        assert_eq!(one.expected_encoded_sha256, chunks.expected_encoded_sha256);
        assert!(one.max_chunk_bytes.saturating_mul(6) <= one.replay_window_bytes);
        assert!(chunks.max_chunk_bytes.saturating_mul(6) <= chunks.replay_window_bytes);
    }

    #[test]
    fn cursor_event_framing_matches_the_independent_authored_proof() {
        for text_mode in [TextMode::Empty, TextMode::Short, TextMode::NearLimit] {
            for chunk_mode in [ChunkMode::One, ChunkMode::Fixed64, ChunkMode::ReplayWindow] {
                let spec = AuthoredSpec {
                    count: 7,
                    chunk_mode,
                    text_mode,
                };
                let expected = authored_record(spec).expect("authored proof");
                let source =
                    GeneratedParagraphSource::new(spec, Arc::new(AuthoredCounters::default()));
                let mut cursor = source.open().expect("cursor");
                let mut event_hash = Sha256::new();
                let mut events = 0_u64;
                let mut text_bytes = 0_u64;
                while let Some(event) = cursor.next().expect("event") {
                    events += 1;
                    match event {
                        PlainParagraphEvent::ParagraphStart => event_hash.update([0_u8]),
                        PlainParagraphEvent::ParagraphEnd => event_hash.update([2_u8]),
                        PlainParagraphEvent::TextChunk(text) => {
                            event_hash.update([1_u8]);
                            event_hash.update(
                                u64::try_from(text.len())
                                    .expect("text length")
                                    .to_le_bytes(),
                            );
                            event_hash.update(text.as_bytes());
                            text_bytes += u64::try_from(text.len()).expect("text length");
                        },
                    }
                }
                assert_eq!(events, expected.event_count);
                assert_eq!(text_bytes, expected.text_bytes);
                assert_eq!(
                    hex_digest(&event_hash.finalize()),
                    expected.expected_event_sha256
                );
            }
        }
    }

    #[test]
    fn candidate_xml_oracle_rejects_a_wrong_byte() {
        let source = source_main_xml(64).expect("source XML");
        let expected = independent_candidate_xml(
            &source,
            AuthoredSpec {
                count: 64,
                chunk_mode: ChunkMode::One,
                text_mode: TextMode::Short,
            },
        )
        .expect("candidate XML");
        let mut mutated = expected.clone();
        let body = mutated
            .windows(b"</w:body>".len())
            .position(|window| window == b"</w:body>")
            .expect("body close");
        mutated[body - 1] ^= 1;
        assert_ne!(mutated, expected);
        assert_ne!(sha256_hex(&mutated), sha256_hex(&expected));
        assert!(verify_candidate_xml_only(&mutated, &expected).is_err());
    }

    #[test]
    fn untouched_member_oracle_rejects_changed_opaque_bytes() {
        let source_xml = source_main_xml(64).expect("source XML");
        let source = source_archive(&source_xml, 0).expect("source archive");
        let candidate = source_archive(&source_xml, 1).expect("mutated archive");
        let source_members = member_identities(&source).expect("source members");
        let candidate_members = member_identities(&candidate).expect("candidate members");
        let equal = source_members
            .iter()
            .filter(|member| member.path != MAIN_PATH)
            .all(|member| {
                candidate_members
                    .iter()
                    .find(|value| value.path == member.path)
                    == Some(member)
            });
        assert!(!equal);
    }

    #[test]
    fn raw_untouched_oracle_rejects_central_metadata_mutation_with_same_payload() {
        let source_xml = source_main_xml(2).expect("source XML");
        let source = source_archive(&source_xml, 0).expect("source archive");
        let archive = ZipArchive::from_slice(&source).expect("ZIP archive");
        let mut central_start = None;
        for result in archive.entries() {
            let entry = result.expect("ZIP entry");
            let path = entry
                .file_path()
                .try_normalize()
                .expect("normalized ZIP path")
                .as_str()
                .to_owned();
            if path == OPAQUE_PATH {
                central_start = Some(
                    usize::try_from(entry.central_directory_offset()).expect("central offset"),
                );
                break;
            }
        }
        let central_start = central_start.expect("opaque central record");
        let mut mutated = source.clone();
        assert!(
            central_start
                .checked_add(46)
                .is_some_and(|end| end <= mutated.len())
        );
        // DOS modification time in the central record; compressed and decoded
        // member bytes remain untouched.
        mutated[central_start + 12] ^= 1;
        let source_members = member_identities(&source).expect("source members");
        let mutated_members = member_identities(&mutated).expect("mutated members");
        assert_eq!(source_members, mutated_members);
        let source_reader = ArchiveReader::new(&source).expect("source reader");
        let mutated_reader = ArchiveReader::new(&mutated).expect("mutated reader");
        assert_eq!(
            source_reader.read(OPAQUE_PATH).expect("source opaque"),
            mutated_reader.read(OPAQUE_PATH).expect("mutated opaque")
        );
        let source_raw = raw_member_records(&source).expect("source raw members");
        let mutated_raw = raw_member_records(&mutated).expect("mutated raw members");
        assert!(!untouched_raw_members_equal(&source_raw, &mutated_raw));
    }

    #[test]
    fn source_adapter_records_returned_bytes_and_request_histogram() {
        let counters = Arc::new(SourceCounters::default());
        let source = MeasureSource::new(Arc::<[u8]>::from(vec![1, 2, 3, 4]), counters);
        let mut buffer = [0_u8; 3];
        assert_eq!(source.read_at(1, &mut buffer).expect("in-range read"), 3);
        assert_eq!(source.read_at(4, &mut buffer).expect("empty read"), 0);
        let reads = source.observation();
        assert_eq!(reads.calls, 2);
        assert_eq!(reads.requested_bytes, 6);
        assert_eq!(reads.returned_bytes, 3);
        assert_eq!(reads.request_histogram.bytes_1_to_512, 2);
        assert_eq!(reads.request_histogram.bytes_0, 0);
        assert_eq!(reads.returned_histogram.bytes_1_to_512, 1);
        assert_eq!(reads.returned_histogram.bytes_0, 1);
    }

    #[test]
    fn sink_records_all_fixed_write_histogram_bins() {
        let mut sink = HashingSink::new(100_000).expect("sink");
        sink.write_all(&[1_u8; 3]).expect("small write");
        sink.write_all(&[2_u8; 600]).expect("medium write");
        sink.write_all(&[3_u8; 5_000]).expect("large write");
        sink.write_all(&[4_u8; 20_000]).expect("very large write");
        sink.write_all(&[5_u8; 70_000]).expect("oversize write");
        let observation = sink.finish();
        assert_eq!(observation.accepted_bytes, 95_603);
        assert_eq!(observation.write_calls, 5);
        assert_eq!(observation.histogram.bytes_1_to_512, 1);
        assert_eq!(observation.histogram.bytes_513_to_4096, 1);
        assert_eq!(observation.histogram.bytes_4097_to_16384, 1);
        assert_eq!(observation.histogram.bytes_16385_to_65536, 1);
        assert_eq!(observation.histogram.bytes_over_65536, 1);
    }

    #[test]
    fn near_limit_text_is_fixed_across_authored_counts() {
        let mut buffer = [0_u8; MAX_CURSOR_TEXT_BYTES];
        for count in [64, 256, 4_096, 16_384] {
            let actual = fill_text(0, count, TextMode::NearLimit, &mut buffer);
            assert_eq!(actual, MAX_CURSOR_TEXT_BYTES);
            assert_eq!(actual, text_bytes(0, count, TextMode::NearLimit));
        }
        let large_index = 100_000_000;
        let actual = fill_text(large_index, 64, TextMode::NearLimit, &mut buffer);
        assert_eq!(actual, MAX_CURSOR_TEXT_BYTES);
        assert_eq!(actual, text_bytes(large_index, 64, TextMode::NearLimit));
    }

    #[test]
    fn authored_counter_oracle_rejects_partial_and_extra_empty_events() {
        let expected = authored_record(AuthoredSpec {
            count: 2,
            chunk_mode: ChunkMode::Fixed64,
            text_mode: TextMode::Empty,
        })
        .expect("empty authored proof");
        let valid = AuthoredObservation {
            opens: EXPECTED_AUTHORED_OPENS,
            events: expected.event_count * EXPECTED_AUTHORED_OPENS,
            text_chunks: 0,
            text_bytes: 0,
        };
        assert!(check_authored_counters(&expected, valid).is_ok());
        let mut partial = valid;
        partial.events -= 1;
        assert!(check_authored_counters(&expected, partial).is_err());
        let mut extra = valid;
        extra.events += 1;
        assert!(check_authored_counters(&expected, extra).is_err());
    }

    #[test]
    fn build_fixture_smoke_exercises_stream_and_full_oracles() {
        let fixture = build_fixture(
            2,
            AuthoredSpec {
                count: 2,
                chunk_mode: ChunkMode::Fixed64,
                text_mode: TextMode::Short,
            },
        )
        .expect("stream fixture");
        assert!(fixture.source_record.unchanged_oracle);
        assert!(fixture.source_record.opaque_member_exact);
        assert_eq!(fixture.oracle.candidate_semantic.paragraph_count, 4);
        assert!(fixture.oracle.candidate_xml_exact);
        assert!(fixture.oracle.candidate_semantic_exact);
        assert!(fixture.oracle.untouched_member_metadata_exact);
        assert!(fixture.oracle.untouched_raw_members_preserved);
        assert!(fixture.oracle.physical_order_exact);
        assert!(fixture.oracle.opaque_member_exact);
        assert!(fixture.oracle.source_unchanged);
        assert!(fixture.oracle.inverse_exact);
        assert!(fixture.proof.generated_once);
        assert_eq!(fixture.proof.authored.authored_count, 2);
        assert_eq!(
            fixture.proof.authored.text_bytes,
            fixture.authored_record.text_bytes
        );
    }

    #[test]
    fn near_limit_fixture_uses_dynamic_parser_workspace() {
        let fixture = build_fixture(
            2,
            AuthoredSpec {
                count: 1,
                chunk_mode: ChunkMode::One,
                text_mode: TextMode::NearLimit,
            },
        )
        .expect("near-limit stream fixture");
        assert_eq!(
            fixture.authored_record.text_bytes,
            MAX_CURSOR_TEXT_BYTES as u64
        );
        assert!(fixture.limit_record.parser_token_bytes > 64 * 1024);
        assert!(fixture.limit_record.parser_workspace_bytes > 2 * 1024 * 1024);
        assert!(fixture.authored_record.xml_entity_reference_count > 128);
        assert!(fixture.limit_record.parser_event_limit > 280);
        assert!(fixture.oracle.candidate_xml_exact);
        assert!(fixture.oracle.inverse_exact);
    }
}

#[cfg(test)]
fn verify_candidate_xml_only(actual: &[u8], expected: &[u8]) -> BenchResult<()> {
    if actual != expected {
        return Err("candidate XML differs from independent oracle".into());
    }
    Ok(())
}
