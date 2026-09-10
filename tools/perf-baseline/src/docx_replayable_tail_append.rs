//! Reproducible lifecycle measurements for replayable DOCX tail appends.
//!
//! The harness keeps source and authored paragraph counts as independent axes.
//! Corpus construction, XML/semantic/ZIP oracles, and exact inverse checks run
//! outside the timed lifecycle.  The historical timed sample owns source
//! admission, replayable stream preparation, publication to a short
//! sequential sink, and all drops.  The after-only counting arm keeps the sink
//! nonretaining and authenticates its emitted bytes with the production
//! publication proof.  The after-only atomic arm times the production path
//! through temporary-file sync, rename, and parent-directory sync; destination
//! byte/hash/semantic checks and cleanup run after the clock.  The authored
//! provider emits borrowed chunks from one bounded cursor buffer; it never
//! builds a complete authored XML stream.

#![allow(clippy::module_name_repetitions)]

use std::{
    collections::BTreeMap,
    ffi::OsString,
    fmt::Write as FmtWrite,
    fs::OpenOptions,
    io::{self, Cursor, Read, Write},
    mem::size_of,
    ops::Range,
    path::{Path, PathBuf},
    sync::{
        Arc, Mutex,
        atomic::{AtomicU64, Ordering},
    },
    time::Instant,
};

use litchi_core::{ReadAt, SourceVersion};
use litchi_docx::source_backed::tail_append::Limits as TailAppendLimits;
use litchi_docx::source_backed::tail_append_stream::{
    AuthoredPassProof, AuthoredReplayError, AuthoredReplayHandle, AuthoredReplayReader,
    AuthoredReplayReference, AuthoredReplayStore, AuthoredStreamProof, MemoryReplayHandle,
    MemoryReplayStore, OneShotParagraphProducer, ParagraphCursor, ParagraphEventSink,
    ParagraphStreamLimits, ParagraphStreamPlan, PlainParagraphEvent, ReplayableParagraphSource,
};
use litchi_docx::{Package as OwnedDocxPackage, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use serde::{Serialize, Serializer};
use sha2::{Digest as _, Sha256};
use soapberry_zip::ZipArchive;
use soapberry_zip::office::ArchiveReader;

use file_store::{
    FileReplayHandle, FileReplayMonitor, FileReplayObservation, FileReplayStore, FileSyncPolicy,
};
use input_profiles::{InputMode, InputProfile, InputSourceCapability};

use compression_profiles::CompressionProfile;

use super::{allocation_metrics, process_metrics};

mod compression_profiles;
mod file_store;
mod input_profiles;
#[cfg(test)]
mod publication_route_tests;
#[cfg(test)]
mod route_failure_tests;
#[cfg(test)]
mod route_smoke_tests;

const SCHEMA: &str = "docx-replayable-tail-append-v1";
const PUBLICATION_SCHEMA: &str = "docx-replayable-tail-append-publication-v1";
const DEFAULT_SOURCE_COUNTS: [usize; 3] = [64, 8_192, 131_072];
const DEFAULT_AUTHORED_COUNTS: [usize; 4] = [64, 256, 4_096, 16_384];
const DEFAULT_CHUNK_BYTES: [usize; 3] = [0, 64, 8 * 1024];
const DEFAULT_SAMPLES: usize = 15;
const DEFAULT_WARMUPS: usize = 3;
const SINK_WRITE_BYTES: [usize; 3] = [512, 4 * 1024, 65_536];
const OPAQUE_PATH: &str = "word/perf-opaque.bin";
const OPAQUE_BYTES: usize = 32 * 1024;
const MAIN_PATH: &str = "word/document.xml";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HASH_SINK_MAX_WRITE: usize = 4 * 1024;
const MAX_CURSOR_TEXT_BYTES: usize = 60 * 1024;
const MAX_STREAM_XML_DEPTH: u64 = 16;
const EXPECTED_AUTHORED_OPENS: u64 = 5;
const EXPECTED_STORE_REPLAY_OPENS: u64 = 4;

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

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq, Hash)]
#[serde(rename_all = "snake_case")]
enum AuthoredProvider {
    Deterministic,
    MemoryStore,
    FileStore,
}

impl AuthoredProvider {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "deterministic" | "cursor" => Ok(Self::Deterministic),
            "memory-store" | "memory_store" | "memory" => Ok(Self::MemoryStore),
            "file-store" | "file_store" | "file" => Ok(Self::FileStore),
            _ => Err(format!(
                "invalid authored provider {value:?}; expected deterministic, memory-store, or file-store"
            )
            .into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Deterministic => "deterministic_replayable_bounded_cursor",
            Self::MemoryStore => "memory_explicit_replay_store",
            Self::FileStore => "file_explicit_replay_store",
        }
    }

    const fn authored_opens(self) -> u64 {
        match self {
            Self::Deterministic => EXPECTED_AUTHORED_OPENS,
            Self::MemoryStore | Self::FileStore => 0,
        }
    }

    const fn replay_opens(self) -> u64 {
        match self {
            Self::Deterministic => 0,
            Self::MemoryStore | Self::FileStore => EXPECTED_STORE_REPLAY_OPENS,
        }
    }

    const fn is_store(self) -> bool {
        !matches!(self, Self::Deterministic)
    }
}

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq, Hash)]
#[serde(rename_all = "snake_case")]
enum ReplaySync {
    None,
    Data,
}

impl ReplaySync {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "none" => Ok(Self::None),
            "data" => Ok(Self::Data),
            _ => Err(format!("invalid replay sync {value:?}; expected none or data").into()),
        }
    }
}

/// Selects the output sink used by the timed publication lifecycle.
///
/// `HashingSink` is the historical default and remains byte-for-byte and
/// schema compatible with the before baseline.  `CountingSink` deliberately
/// does not compute a local digest or retain archive bytes: the production
/// publication's returned artifact proof authenticates the bytes accepted by
/// the sink.  `AtomicPath` is an after-only capability because the production
/// `write_to_path` method did not exist in the before baseline.
#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq, Hash)]
#[serde(rename_all = "snake_case")]
enum PublicationMode {
    HashingSink,
    CountingSink,
    AtomicPath,
}

impl PublicationMode {
    const fn is_default(self) -> bool {
        matches!(self, Self::HashingSink)
    }

    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "hashing" | "hashing-sink" | "hashing_sink" => Ok(Self::HashingSink),
            "counting" | "counting-sink" | "counting_sink" => Ok(Self::CountingSink),
            "atomic" | "atomic-path" | "atomic_path" => {
                Ok(Self::AtomicPath)
            },
            _ => Err(format!(
                "invalid publication mode {value:?}; expected hashing-sink, counting-sink, or atomic-path"
            )
            .into()),
        }
    }
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
    authored_provider: AuthoredProvider,
    replay_dir: Option<PathBuf>,
    replay_max_bytes: Option<u64>,
    replay_sync: ReplaySync,
    compression: CompressionProfile,
    input_mode: Option<InputMode>,
    input_file: Option<PathBuf>,
    input_max_range_bytes: Option<usize>,
    input_delay_us: u64,
    input_overhead_us: u64,
    input_bytes_per_second: Option<u64>,
    publication: PublicationMode,
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
    expected_replay_opens: u64,
    source: &'static str,
    authored_provider: &'static str,
    provider: AuthoredProvider,
    replay_max_bytes: Option<u64>,
    replay_dir: Option<String>,
    replay_sync: ReplaySync,
    compression: CompressionProfile,
    input_mode: &'static str,
    input_storage_kind: &'static str,
    input_identity_validation: &'static str,
    input_max_range_bytes: Option<usize>,
    input_delay_us: u64,
    input_overhead_us: u64,
    input_bytes_per_second: Option<u64>,
    sink: &'static str,
    fixture_dir: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    publication: Option<PublicationMode>,
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
    text_chunk_count: u64,
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
    replay: Option<ReplayObservation>,
    sink: SinkRecord,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
    #[serde(skip_serializing_if = "Option::is_none")]
    publication: Option<PublicationRecord>,
}

#[derive(Clone, Debug, Serialize)]
struct CaseRecord {
    source_count: usize,
    authored_count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
    provider: AuthoredProvider,
    replay_max_bytes: Option<u64>,
    compression: CompressionProfile,
    input_mode: &'static str,
    input_storage_kind: &'static str,
    input_identity_validation: &'static str,
    sink_write_bytes: usize,
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

    #[cfg(test)]
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

struct ProfiledMeasureSource {
    inner: Arc<dyn ReadAt>,
    counters: Arc<SourceCounters>,
}

impl ProfiledMeasureSource {
    fn new(inner: Arc<dyn ReadAt>, counters: Arc<SourceCounters>) -> Self {
        Self { inner, counters }
    }
}

impl ReadAt for ProfiledMeasureSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.read_calls.fetch_add(1, Ordering::Relaxed);
        self.counters.requested_bytes.fetch_add(
            u64::try_from(output.len())
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "request overflow"))?,
            Ordering::Relaxed,
        );
        self.counters.request_histogram.record(output.len());
        let returned = self.inner.read_at(offset, output)?;
        if returned > output.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "profiled source returned more bytes than requested",
            ));
        }
        self.counters.returned_bytes.fetch_add(
            u64::try_from(returned)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "read overflow"))?,
            Ordering::Relaxed,
        );
        self.counters.returned_histogram.record(returned);
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
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

#[derive(Clone, Copy, Debug, Default, Serialize, PartialEq, Eq)]
struct ReplayObservation {
    route: Option<AuthoredProvider>,
    producer_invocations: u64,
    #[serde(rename = "store_prepare_calls")]
    prepare_calls: u64,
    #[serde(rename = "store_append_calls")]
    append_calls: u64,
    #[serde(rename = "store_appended_bytes")]
    appended_bytes: u64,
    store_finish_calls: u64,
    replay_opens: u64,
    replay_read_calls: u64,
    replay_requested_bytes: u64,
    replay_returned_bytes: u64,
    replay_finish_calls: u64,
    replay_sha256_checks: u64,
    request_histogram: RequestHistogram,
    returned_histogram: RequestHistogram,
    retained_logical_bytes: Option<u64>,
    retained_capacity_bytes: Option<u64>,
    retained_capacity_provenance: Option<&'static str>,
    file_logical_bytes: Option<u64>,
    file_allocated_bytes: Option<u64>,
    file_write_calls: Option<u64>,
    file_sync_calls: Option<u64>,
    file_cleanup_verified: Option<bool>,
    seal_sha256_checks: Option<u64>,
    cleanup_sha256_checks: Option<u64>,
    durable_reference_kind: Option<&'static str>,
    durable_reference_bytes: u64,
    #[serde(serialize_with = "serialize_optional_digest")]
    durable_reference_sha256: Option<[u8; 32]>,
    /// Kept for internal post-timer validation; the public report uses the
    /// fixed-size flattened fields above.
    #[serde(skip_serializing)]
    file: Option<FileStoreObservation>,
}

#[derive(Clone, Copy, Debug, Default, Serialize, PartialEq, Eq)]
struct FileStoreObservation {
    write_calls: u64,
    sync_calls: u64,
    read_calls: u64,
    returned_bytes: u64,
    replay_sha256_checks: u64,
    seal_sha256_checks: u64,
    cleanup_sha256_checks: u64,
    logical_bytes: Option<u64>,
    allocated_bytes: Option<u64>,
    cleanup_verified: bool,
}

#[derive(Debug, Default)]
struct ReplayCounters {
    producer_invocations: AtomicU64,
    prepare_calls: AtomicU64,
    append_calls: AtomicU64,
    appended_bytes: AtomicU64,
    store_finish_calls: AtomicU64,
    replay_opens: AtomicU64,
    replay_read_calls: AtomicU64,
    replay_requested_bytes: AtomicU64,
    replay_returned_bytes: AtomicU64,
    replay_finish_calls: AtomicU64,
    replay_sha256_checks: AtomicU64,
    request_histogram: AtomicRequestHistogram,
    returned_histogram: AtomicRequestHistogram,
    durable_reference: Mutex<Option<DurableReferenceObservation>>,
}

#[derive(Clone, Copy, Debug)]
struct DurableReferenceObservation {
    bytes: u64,
    sha256: [u8; 32],
}

fn serialize_optional_digest<S>(digest: &Option<[u8; 32]>, serializer: S) -> Result<S::Ok, S::Error>
where
    S: Serializer,
{
    match digest {
        Some(digest) => serializer.serialize_some(&hex_digest(digest)),
        None => serializer.serialize_none(),
    }
}

impl ReplayCounters {
    fn snapshot(&self) -> ReplayObservation {
        let durable_reference = self
            .durable_reference
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        ReplayObservation {
            producer_invocations: self.producer_invocations.load(Ordering::Relaxed),
            prepare_calls: self.prepare_calls.load(Ordering::Relaxed),
            append_calls: self.append_calls.load(Ordering::Relaxed),
            appended_bytes: self.appended_bytes.load(Ordering::Relaxed),
            store_finish_calls: self.store_finish_calls.load(Ordering::Relaxed),
            replay_opens: self.replay_opens.load(Ordering::Relaxed),
            replay_read_calls: self.replay_read_calls.load(Ordering::Relaxed),
            replay_requested_bytes: self.replay_requested_bytes.load(Ordering::Relaxed),
            replay_returned_bytes: self.replay_returned_bytes.load(Ordering::Relaxed),
            replay_finish_calls: self.replay_finish_calls.load(Ordering::Relaxed),
            replay_sha256_checks: self.replay_sha256_checks.load(Ordering::Relaxed),
            request_histogram: self.request_histogram.snapshot(),
            returned_histogram: self.returned_histogram.snapshot(),
            durable_reference_bytes: durable_reference
                .as_ref()
                .map_or(0, |reference| reference.bytes),
            durable_reference_sha256: durable_reference.as_ref().map(|reference| reference.sha256),
            file: None,
            ..ReplayObservation::default()
        }
    }
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

/// Count the explicit replay-store lifecycle without retaining authored output
/// in the harness.  The wrapper deliberately forwards both operation hooks so
/// a managed provider can charge the current package context on every reopen.
struct CountingReplayStore<R> {
    inner: R,
    counters: Arc<ReplayCounters>,
}

impl<R> CountingReplayStore<R> {
    fn new(inner: R, counters: Arc<ReplayCounters>) -> Self {
        Self { inner, counters }
    }
}

struct CountingReplayHandle<H> {
    inner: H,
    counters: Arc<ReplayCounters>,
}

struct CountingReplayReader<'a> {
    inner: Box<dyn AuthoredReplayReader + 'a>,
    counters: Arc<ReplayCounters>,
}

impl<R> AuthoredReplayStore for CountingReplayStore<R>
where
    R: AuthoredReplayStore,
    R::Handle: 'static,
{
    type Handle = CountingReplayHandle<R::Handle>;

    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&litchi_core::ExecutionContext>,
        cancellation: Option<&litchi_core::CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        self.counters.prepare_calls.fetch_add(1, Ordering::Relaxed);
        self.inner
            .prepare_for_operation(limits, context, cancellation)
    }

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        self.counters.append_calls.fetch_add(1, Ordering::Relaxed);
        self.counters.appended_bytes.fetch_add(
            u64::try_from(chunk.len()).unwrap_or(u64::MAX),
            Ordering::Relaxed,
        );
        self.inner.append(chunk)
    }

    fn finish(self, proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        self.counters
            .store_finish_calls
            .fetch_add(1, Ordering::Relaxed);
        let inner = self.inner.finish(proof)?;
        if let Some(reference) = inner.durable_reference() {
            let bytes = u64::try_from(reference.len()).map_err(|_| {
                AuthoredReplayError::Store("durable replay reference length overflow")
            })?;
            let digest: [u8; 32] = Sha256::digest(reference.as_bytes()).into();
            *self
                .counters
                .durable_reference
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner()) =
                Some(DurableReferenceObservation {
                    bytes,
                    sha256: digest,
                });
        }
        Ok(CountingReplayHandle {
            inner,
            counters: self.counters,
        })
    }
}

impl<H> AuthoredReplayHandle for CountingReplayHandle<H>
where
    H: AuthoredReplayHandle,
{
    fn proof(&self) -> AuthoredStreamProof {
        self.inner.proof()
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        self.counters.replay_opens.fetch_add(1, Ordering::Relaxed);
        let inner = self.inner.open()?;
        Ok(Box::new(CountingReplayReader {
            inner,
            counters: Arc::clone(&self.counters),
        }))
    }

    fn open_for_package<'a>(
        &'a self,
        context: Option<&'a litchi_core::ExecutionContext>,
        cancellation: Option<&'a litchi_core::CancellationToken>,
    ) -> Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        self.counters.replay_opens.fetch_add(1, Ordering::Relaxed);
        let inner = self.inner.open_for_package(context, cancellation)?;
        Ok(Box::new(CountingReplayReader {
            inner,
            counters: Arc::clone(&self.counters),
        }))
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        self.inner.durable_reference()
    }
}

impl Read for CountingReplayReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        self.counters
            .replay_read_calls
            .fetch_add(1, Ordering::Relaxed);
        self.counters.replay_requested_bytes.fetch_add(
            u64::try_from(output.len()).unwrap_or(u64::MAX),
            Ordering::Relaxed,
        );
        self.counters.request_histogram.record(output.len());
        let count = self.inner.read(output)?;
        self.counters
            .replay_returned_bytes
            .fetch_add(u64::try_from(count).unwrap_or(u64::MAX), Ordering::Relaxed);
        self.counters.returned_histogram.record(count);
        Ok(count)
    }
}

impl AuthoredReplayReader for CountingReplayReader<'_> {
    fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError> {
        let CountingReplayReader { inner, counters } = *self;
        counters.replay_finish_calls.fetch_add(1, Ordering::Relaxed);
        let result = inner.finish();
        if result.is_ok() {
            counters
                .replay_sha256_checks
                .fetch_add(1, Ordering::Relaxed);
        }
        result
    }
}

#[derive(Clone, Debug)]
struct FileReplayCleanupRecord {
    observation: FileReplayObservation,
}

struct FileReplayRouteStore {
    inner: FileReplayStore,
    cleanup: Arc<Mutex<Option<FileReplayCleanupRecord>>>,
}

impl FileReplayRouteStore {
    fn new(
        path: PathBuf,
        maximum: u64,
        sync_policy: FileSyncPolicy,
        cleanup: Arc<Mutex<Option<FileReplayCleanupRecord>>>,
        monitor: &FileReplayMonitor,
    ) -> BenchResult<Self> {
        Ok(Self {
            inner: FileReplayStore::new_with_monitor(path, maximum, sync_policy, monitor)?,
            cleanup,
        })
    }
}

impl AuthoredReplayStore for FileReplayRouteStore {
    type Handle = FileReplayHandle;

    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&litchi_core::ExecutionContext>,
        cancellation: Option<&litchi_core::CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        self.inner
            .prepare_for_operation(limits, context, cancellation)
    }

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        self.inner.append(chunk)
    }

    fn finish(self, proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        let Self { inner, cleanup } = self;
        let handle = inner.finish(proof)?;
        let observation = handle.observation();
        *cleanup
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner()) =
            Some(FileReplayCleanupRecord { observation });
        Ok(handle)
    }
}

enum ReplayRouteStore {
    Memory(CountingReplayStore<MemoryReplayStore>),
    File(CountingReplayStore<FileReplayRouteStore>),
}

enum ReplayRouteHandle {
    Memory(CountingReplayHandle<MemoryReplayHandle>),
    File(CountingReplayHandle<FileReplayHandle>),
}

impl AuthoredReplayStore for ReplayRouteStore {
    type Handle = ReplayRouteHandle;

    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&litchi_core::ExecutionContext>,
        cancellation: Option<&litchi_core::CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        match self {
            Self::Memory(store) => store.prepare_for_operation(limits, context, cancellation),
            Self::File(store) => store.prepare_for_operation(limits, context, cancellation),
        }
    }

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        match self {
            Self::Memory(store) => store.append(chunk),
            Self::File(store) => store.append(chunk),
        }
    }

    fn finish(self, proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        match self {
            Self::Memory(store) => Ok(ReplayRouteHandle::Memory(store.finish(proof)?)),
            Self::File(store) => Ok(ReplayRouteHandle::File(store.finish(proof)?)),
        }
    }
}

impl AuthoredReplayHandle for ReplayRouteHandle {
    fn proof(&self) -> AuthoredStreamProof {
        match self {
            Self::Memory(handle) => handle.proof(),
            Self::File(handle) => handle.proof(),
        }
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        match self {
            Self::Memory(handle) => handle.open(),
            Self::File(handle) => handle.open(),
        }
    }

    fn open_for_package<'a>(
        &'a self,
        context: Option<&'a litchi_core::ExecutionContext>,
        cancellation: Option<&'a litchi_core::CancellationToken>,
    ) -> Result<Box<dyn AuthoredReplayReader + 'a>, AuthoredReplayError> {
        match self {
            Self::Memory(handle) => handle.open_for_package(context, cancellation),
            Self::File(handle) => handle.open_for_package(context, cancellation),
        }
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        match self {
            Self::Memory(handle) => handle.durable_reference(),
            Self::File(handle) => handle.durable_reference(),
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct AuthoredSpec {
    count: usize,
    chunk_mode: ChunkMode,
    text_mode: TextMode,
}

struct GeneratedParagraphProducer {
    spec: AuthoredSpec,
    counters: Arc<AuthoredCounters>,
    replay_counters: Arc<ReplayCounters>,
}

impl GeneratedParagraphProducer {
    fn new(
        spec: AuthoredSpec,
        counters: Arc<AuthoredCounters>,
        replay_counters: Arc<ReplayCounters>,
    ) -> Self {
        Self {
            spec,
            counters,
            replay_counters,
        }
    }
}

impl OneShotParagraphProducer for GeneratedParagraphProducer {
    fn produce(&mut self, sink: &mut dyn ParagraphEventSink) -> Result<(), AuthoredReplayError> {
        self.replay_counters
            .producer_invocations
            .fetch_add(1, Ordering::Relaxed);
        let mut cursor = GeneratedCursor::new(self.spec, Arc::clone(&self.counters));
        while let Some(event) = cursor
            .next()
            .map_err(|error| AuthoredReplayError::Provider(Box::new(error)))?
        {
            sink.push(event)?;
        }
        Ok(())
    }
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
        Ok(GeneratedCursor::new(self.spec, Arc::clone(&self.counters)))
    }
}

impl GeneratedCursor {
    fn new(spec: AuthoredSpec, counters: Arc<AuthoredCounters>) -> Self {
        Self {
            spec,
            counters,
            paragraph: 0,
            phase: if spec.count == 0 {
                CursorPhase::Done
            } else {
                CursorPhase::Start
            },
            text_offset: 0,
            text_length: 0,
            text: [0; MAX_CURSOR_TEXT_BYTES],
        }
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
    let mut text_chunk_count = 0_u64;
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
                text_chunk_count = text_chunk_count
                    .checked_add(1)
                    .ok_or("authored text-chunk overflow")?;
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
        text_chunk_count,
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

fn source_archive_with_compression(
    main_xml: &[u8],
    opaque_variant: u8,
    compression: CompressionProfile,
) -> BenchResult<Vec<u8>> {
    let archive = source_archive(main_xml, opaque_variant)?;
    Ok(compression_profiles::apply(archive, compression)?)
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
    digest: Option<[u8; 32]>,
    observed: bool,
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
            digest: Some(self.digest.finalize().into()),
            observed: true,
        }
    }
}

/// A bounded, nonretaining sequential sink used by the after-only counting
/// arm.  It intentionally has no digest state: the timed
/// `ParagraphStreamPublication` artifact proof is the authoritative byte/hash
/// evidence for this arm.
#[derive(Debug)]
struct CountingSink {
    max_write: usize,
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    histogram: SinkHistogram,
}

impl CountingSink {
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
        })
    }

    fn finish(self) -> SinkObservation {
        SinkObservation {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            largest_write: self.largest_write,
            histogram: self.histogram,
            digest: None,
            observed: true,
        }
    }
}

impl Write for CountingSink {
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
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
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
    fn atomic(length: u64) -> Self {
        Self {
            accepted_bytes: length,
            write_calls: 0,
            largest_write: 0,
            histogram: SinkHistogram::default(),
            digest: None,
            observed: false,
        }
    }

    fn record(self) -> SinkRecord {
        SinkRecord {
            accepted_bytes: self.observed.then_some(self.accepted_bytes),
            write_calls: self.observed.then_some(self.write_calls),
            largest_write: self.observed.then_some(self.largest_write),
            histogram: self.observed.then_some(self.histogram),
            sha256: self.digest.map(|digest| hex_digest(&digest)),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct SinkRecord {
    accepted_bytes: Option<u64>,
    write_calls: Option<u64>,
    largest_write: Option<u64>,
    histogram: Option<SinkHistogram>,
    sha256: Option<String>,
}

#[derive(Clone, Copy, Debug)]
struct TimedArtifactProof {
    length: u64,
    sha256: [u8; 32],
}

impl TimedArtifactProof {
    fn from_publication(
        publication: &litchi_docx::source_backed::tail_append_stream::ParagraphStreamPublication,
    ) -> Self {
        Self {
            length: publication.candidate_artifact_length(),
            sha256: publication.candidate_artifact_fingerprint().into_sha256(),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct PublicationRecord {
    schema: &'static str,
    route: PublicationMode,
    timing_scope: &'static str,
    timed_candidate_artifact_bytes: u64,
    timed_candidate_artifact_sha256: String,
    timed_candidate_matches_oracle: bool,
    verification_scope: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    atomic: Option<AtomicOutputRecord>,
}

#[derive(Clone, Debug, Serialize)]
struct AtomicPathState {
    exists: bool,
    regular_file: bool,
    bytes: Option<u64>,
}

#[derive(Clone, Debug, Serialize)]
struct AtomicCleanupRecord {
    destination_removed: bool,
    parent_removed: bool,
}

#[derive(Clone, Debug, Serialize)]
struct AtomicOutputRecord {
    destination_path: String,
    private_parent_path: String,
    before: AtomicPathState,
    after: AtomicPathState,
    post_timer_archive_bytes: u64,
    post_timer_archive_sha256: String,
    output_bytes_exact: bool,
    output_sha256_exact: bool,
    inverse_oracle_scope: &'static str,
    post_timer_oracle: OracleRecord,
    cleanup: AtomicCleanupRecord,
}

#[derive(Clone, Debug)]
struct Fixture {
    authored: AuthoredSpec,
    compression: CompressionProfile,
    authored_record: AuthoredRecord,
    limits: ParagraphStreamLimits,
    limit_record: StreamLimitRecord,
    source_archive: Arc<[u8]>,
    candidate_archive: Arc<[u8]>,
    inverse_archive: Arc<[u8]>,
    source_record: SourceRecord,
    oracle: OracleRecord,
    proof: ProofRecord,
}

struct PreparedInput {
    profile: InputProfile,
    capability: InputSourceCapability,
}

static NEXT_ATOMIC_DESTINATION: AtomicU64 = AtomicU64::new(0);

struct AtomicDestination {
    parent: PathBuf,
    path: PathBuf,
    before: AtomicPathState,
    cleaned: bool,
}

impl AtomicDestination {
    fn prepare() -> BenchResult<Self> {
        let parent = match (0..32_u8).find_map(|_| {
            let serial = NEXT_ATOMIC_DESTINATION.fetch_add(1, Ordering::Relaxed);
            let path = std::env::temp_dir().join(format!(
                "litchi-docx-replay-atomic-{}-{serial}",
                std::process::id()
            ));
            match std::fs::create_dir(&path) {
                Ok(()) => Some(Ok(path)),
                Err(error) if error.kind() == io::ErrorKind::AlreadyExists => None,
                Err(error) => Some(Err(error)),
            }
        }) {
            Some(parent) => parent?,
            None => return Err("unable to allocate a private atomic publication directory".into()),
        };
        let path = parent.join("published.docx");
        let before = atomic_path_state(&path)?;
        if before.exists {
            return Err("atomic publication destination unexpectedly exists during setup".into());
        }
        Ok(Self {
            parent,
            path,
            before,
            cleaned: false,
        })
    }

    fn path(&self) -> &Path {
        &self.path
    }

    fn verify_and_cleanup(
        &mut self,
        fixture: &Fixture,
        proof: TimedArtifactProof,
    ) -> BenchResult<AtomicOutputRecord> {
        let bytes = std::fs::read(&self.path)?;
        let after = atomic_path_state(&self.path)?;
        let byte_len = u64::try_from(bytes.len())?;
        let archive_sha256 = sha256_hex(&bytes);
        let output_bytes_exact = bytes.len() == fixture.oracle.candidate_archive_bytes
            && proof.length == byte_len
            && after.bytes == Some(byte_len);
        let output_sha256_exact = archive_sha256 == fixture.oracle.candidate_archive_sha256
            && hex_digest(&proof.sha256) == fixture.oracle.candidate_archive_sha256;
        let source_xml = source_main_xml(fixture.source_record.semantic.paragraph_count)?;
        let expected_candidate_xml = independent_candidate_xml(&source_xml, fixture.authored)?;
        let post_timer_oracle = verify_candidate_oracles(
            fixture.source_archive.as_ref(),
            &source_xml,
            &bytes,
            &expected_candidate_xml,
            fixture.source_record.semantic.paragraph_count,
            fixture.authored,
            fixture.inverse_archive.as_ref(),
        )?;
        if !output_bytes_exact || !output_sha256_exact || !after.exists || !after.regular_file {
            return Err(
                "atomic publication destination failed its post-timer artifact oracle".into(),
            );
        }
        let cleanup = self.cleanup()?;
        Ok(AtomicOutputRecord {
            destination_path: self.path.display().to_string(),
            private_parent_path: self.parent.display().to_string(),
            before: self.before.clone(),
            after,
            post_timer_archive_bytes: byte_len,
            post_timer_archive_sha256: archive_sha256,
            output_bytes_exact,
            output_sha256_exact,
            inverse_oracle_scope: "untimed_fixture_publication_inverse_exact; timed_atomic_publication_inverse_not_reexecuted",
            post_timer_oracle,
            cleanup,
        })
    }

    fn cleanup(&mut self) -> BenchResult<AtomicCleanupRecord> {
        let destination_removed = match std::fs::remove_file(&self.path) {
            Ok(()) => true,
            Err(error) if error.kind() == io::ErrorKind::NotFound => false,
            Err(error) => return Err(error.into()),
        };
        let parent_removed = match std::fs::remove_dir(&self.parent) {
            Ok(()) => true,
            Err(error) if error.kind() == io::ErrorKind::NotFound => false,
            Err(error) => return Err(error.into()),
        };
        self.cleaned = true;
        Ok(AtomicCleanupRecord {
            destination_removed,
            parent_removed,
        })
    }
}

impl Drop for AtomicDestination {
    fn drop(&mut self) {
        if !self.cleaned {
            // A failed atomic write may have returned `OpcError::Committed`:
            // the destination can already be durable even though the parent
            // directory sync failed.  Keep that artifact and its private
            // parent for the driver to audit.  If no destination was created,
            // only remove the empty setup directory.
            if std::fs::symlink_metadata(&self.path).is_err() {
                let _ = std::fs::remove_dir(&self.parent);
            }
        }
    }
}

fn atomic_path_state(path: &Path) -> BenchResult<AtomicPathState> {
    match std::fs::symlink_metadata(path) {
        Ok(metadata) => Ok(AtomicPathState {
            exists: true,
            regular_file: metadata.is_file(),
            bytes: metadata.is_file().then_some(metadata.len()),
        }),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(AtomicPathState {
            exists: false,
            regular_file: false,
            bytes: None,
        }),
        Err(error) => Err(error.into()),
    }
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
    provider: AuthoredProvider,
    compression: CompressionProfile,
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
    compression: CompressionProfile,
) -> String {
    let mut stem = format!(
        "source{source_count}-authored{authored_count}-{}-{}",
        chunk_mode.name(),
        text_mode.name()
    );
    if compression != CompressionProfile::Current {
        stem.push('-');
        stem.push_str(compression.name());
    }
    stem
}

fn write_fixture_file(path: &Path, bytes: &[u8]) -> BenchResult<()> {
    let mut output = OpenOptions::new().write(true).create_new(true).open(path)?;
    output.write_all(bytes)?;
    output.flush()?;
    Ok(())
}

fn export_fixture(
    dir: &Path,
    fixture: &Fixture,
    provider: AuthoredProvider,
    candidate_override: Option<&[u8]>,
) -> BenchResult<()> {
    std::fs::create_dir_all(dir)?;
    let source_count = fixture.source_record.semantic.paragraph_count;
    let mut stem = fixture_stem(
        source_count,
        fixture.authored.count,
        fixture.authored.chunk_mode,
        fixture.authored.text_mode,
        fixture.compression,
    );
    if provider != AuthoredProvider::Deterministic {
        stem.push('-');
        stem.push_str(provider.name());
    }
    let source_file = format!("{stem}-source.docx");
    let candidate_file = format!("{stem}-candidate.docx");
    let manifest_file = format!("{stem}-hashes.json");
    let source_path = dir.join(&source_file);
    let candidate_path = dir.join(&candidate_file);
    let manifest_path = dir.join(&manifest_file);
    let source_sha256 = sha256_hex(fixture.source_archive.as_ref());
    let candidate = candidate_override.unwrap_or(fixture.candidate_archive.as_ref());
    let candidate_sha256 = sha256_hex(candidate);
    if source_sha256 != fixture.source_record.archive_sha256
        || candidate_sha256 != fixture.oracle.candidate_archive_sha256
    {
        return Err("fixture export archive hash disagrees with its oracle".into());
    }
    write_fixture_file(&source_path, fixture.source_archive.as_ref())?;
    write_fixture_file(&candidate_path, candidate)?;
    let candidate_reader = ArchiveReader::new(candidate)?;
    let candidate_main = candidate_reader.read(MAIN_PATH)?;
    let manifest = FixtureManifest {
        schema: "docx-replayable-tail-append-fixture-v1",
        version: 1,
        source_count,
        authored_count: fixture.authored.count,
        chunk_mode: fixture.authored.chunk_mode,
        text_mode: fixture.authored.text_mode,
        provider,
        compression: fixture.compression,
        source: FixtureArtifactRecord {
            file: source_file,
            bytes: fixture.source_archive.len(),
            sha256: source_sha256,
        },
        candidate: FixtureArtifactRecord {
            file: candidate_file,
            bytes: candidate.len(),
            sha256: candidate_sha256,
        },
        source_main_xml_sha256: fixture.source_record.main_xml_sha256.clone(),
        candidate_main_xml_sha256: sha256_hex(&candidate_main),
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
            text_chunk_count: expected_authored.text_chunk_count,
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
        || proof.authored.text_chunk_count != authored.text_chunk_count
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

fn build_fixture(
    source_count: usize,
    authored: AuthoredSpec,
    compression: CompressionProfile,
) -> BenchResult<Fixture> {
    let source_xml = source_main_xml(source_count)?;
    let source_archive = source_archive_with_compression(&source_xml, 0, compression)?;
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
    // The independent inverse oracle above proves that the restored archive is
    // byte-identical to the already-owned source fixture.  Reuse that source
    // owner for post-timer oracle comparison instead of retaining another
    // full archive allocation.
    let inverse_archive = Arc::clone(&source_bytes);
    Ok(Fixture {
        authored,
        compression,
        authored_record,
        limits,
        limit_record,
        source_archive: source_bytes,
        candidate_archive: candidate_bytes,
        inverse_archive,
        source_record,
        oracle,
        proof,
    })
}

fn prepare_input(fixture: &Fixture, config: &Config) -> BenchResult<Option<PreparedInput>> {
    let Some(mode) = config.input_mode else {
        return Ok(None);
    };
    let capability = if matches!(mode, InputMode::File) {
        let path = config
            .input_file
            .as_ref()
            .ok_or("file input mode requires --input-file")?;
        InputSourceCapability::prepare_file(path)?
    } else {
        InputSourceCapability::owned(Arc::clone(&fixture.source_archive))
    };
    capability.verify_fingerprint()?;
    let fingerprint = capability.fingerprint();
    if fingerprint.len() != fixture.source_archive.len() as u64
        || fingerprint.sha256_hex() != fixture.source_record.archive_sha256
    {
        return Err("input capability does not match the generated source archive".into());
    }
    let profile = match mode {
        InputMode::Owned => InputProfile::owned(),
        InputMode::File => InputProfile::file(),
        InputMode::ShortRead => InputProfile::try_new(
            mode,
            config.input_max_range_bytes,
            std::time::Duration::ZERO,
            std::time::Duration::ZERO,
            None,
        )?,
        InputMode::Latency => InputProfile::latency(
            config
                .input_max_range_bytes
                .ok_or("latency input mode requires --input-max-range")?,
            std::time::Duration::from_micros(config.input_delay_us),
            std::time::Duration::from_micros(config.input_overhead_us),
            config.input_bytes_per_second,
        )?,
    };
    Ok(Some(PreparedInput {
        profile,
        capability,
    }))
}

fn input_mode_name(input: Option<&PreparedInput>) -> &'static str {
    input.map_or("native_owned", |prepared| prepared.profile.mode().as_str())
}

fn input_storage_kind(input: Option<&PreparedInput>) -> &'static str {
    input.map_or("owned", |prepared| {
        prepared.capability.storage_kind().as_str()
    })
}

fn input_identity_validation(input: Option<&PreparedInput>) -> &'static str {
    match input {
        None => "native_source_archive_oracle",
        Some(prepared) if prepared.capability.is_file() => {
            "setup_and_post_sample_fingerprint_outside_timing"
        },
        Some(_) => "setup_fingerprint_outside_timing",
    }
}

fn source_description(input_mode: Option<InputMode>) -> &'static str {
    match input_mode {
        None => "caller_owned_arc_positional_read_at_requested_returned_fixed_histograms",
        Some(InputMode::Owned) => {
            "profile_owned_positional_read_at_requested_returned_fixed_histograms"
        },
        Some(InputMode::File) => {
            "profile_file_source_positional_read_at_requested_returned_fixed_histograms"
        },
        Some(InputMode::ShortRead) => {
            "profile_short_read_positional_read_at_requested_returned_fixed_histograms"
        },
        Some(InputMode::Latency) => {
            "profile_latency_positional_read_at_requested_returned_fixed_histograms"
        },
    }
}

fn limits_for_provider(fixture: &Fixture, config: &Config) -> BenchResult<ParagraphStreamLimits> {
    if !config.authored_provider.is_store() {
        return Ok(fixture.limits);
    }
    let maximum = config
        .replay_max_bytes
        .ok_or("store provider requires --replay-max-bytes")?;
    let mut limits = fixture.limits;
    limits.max_replay_bytes = maximum;
    limits.validate()?;
    if maximum < fixture.authored_record.encoded_xml_bytes {
        return Err(format!(
            "--replay-max-bytes {maximum} is below authored XML bytes {}",
            fixture.authored_record.encoded_xml_bytes
        )
        .into());
    }
    Ok(limits)
}

fn file_replay_path(fixture: &Fixture, config: &Config) -> BenchResult<PathBuf> {
    let directory = config
        .replay_dir
        .as_ref()
        .ok_or("file-store provider requires --replay-dir")?;
    Ok(directory.join(format!(
        "{}.replay",
        fixture_stem(
            fixture.source_record.semantic.paragraph_count,
            fixture.authored.count,
            fixture.authored.chunk_mode,
            fixture.authored.text_mode,
            fixture.compression,
        )
    )))
}

fn cleanup_file_route(
    cleanup: Arc<Mutex<Option<FileReplayCleanupRecord>>>,
    path: &Path,
    monitor: &FileReplayMonitor,
) -> BenchResult<Option<FileStoreObservation>> {
    let record = cleanup
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .take();
    let Some(record) = record else {
        return Ok(None);
    };
    let cleanup_stats = FileReplayStore::cleanup(path, record.observation)?;
    let stats = monitor.stats();
    Ok(Some(FileStoreObservation {
        write_calls: stats.write_calls,
        sync_calls: stats.sync_calls,
        read_calls: stats.replay_read_calls,
        returned_bytes: stats.replay_returned_bytes,
        replay_sha256_checks: stats.replay_sha256_checks,
        seal_sha256_checks: stats.seal_sha256_checks,
        cleanup_sha256_checks: cleanup_stats.cleanup_sha256_checks,
        logical_bytes: cleanup_stats.file_logical_bytes,
        allocated_bytes: cleanup_stats.file_allocated_bytes,
        cleanup_verified: cleanup_stats.file_cleanup_verified,
    }))
}

fn verify_store_preflight(
    fixture: &Fixture,
    config: &Config,
    limits: ParagraphStreamLimits,
) -> BenchResult<Vec<u8>> {
    let replay_counters = Arc::new(ReplayCounters::default());
    if matches!(config.authored_provider, AuthoredProvider::MemoryStore) {
        let store = ReplayRouteStore::Memory(CountingReplayStore::new(
            MemoryReplayStore::new(limits.max_replay_bytes)?,
            Arc::clone(&replay_counters),
        ));
        return verify_store_preflight_with_store(fixture, config, limits, store, replay_counters);
    }
    let directory = config
        .replay_dir
        .as_ref()
        .ok_or("file-store provider requires --replay-dir")?;
    std::fs::create_dir_all(directory)?;
    let cleanup = Arc::new(Mutex::new(None));
    let monitor = FileReplayMonitor::new();
    let store = FileReplayRouteStore::new(
        file_replay_path(fixture, config)?,
        limits.max_replay_bytes,
        match config.replay_sync {
            ReplaySync::None => FileSyncPolicy::None,
            ReplaySync::Data => FileSyncPolicy::Data,
        },
        Arc::clone(&cleanup),
        &monitor,
    )?;
    let counted = ReplayRouteStore::File(CountingReplayStore::new(
        store,
        Arc::clone(&replay_counters),
    ));
    let result =
        verify_store_preflight_with_store(fixture, config, limits, counted, replay_counters);
    let cleanup_result = cleanup_file_route(cleanup, &file_replay_path(fixture, config)?, &monitor);
    let candidate = result?;
    cleanup_result?.ok_or("file-store preflight did not produce a cleanup record")?;
    Ok(candidate)
}

fn verify_store_preflight_with_store<R>(
    fixture: &Fixture,
    config: &Config,
    limits: ParagraphStreamLimits,
    store: R,
    replay_counters: Arc<ReplayCounters>,
) -> BenchResult<Vec<u8>>
where
    R: AuthoredReplayStore + 'static,
    R::Handle: 'static,
{
    let source_counters = Arc::new(SourceCounters::default());
    let authored_counters = Arc::new(AuthoredCounters::default());
    let (candidate_archive, inverse, proof) = {
        let source = Arc::new(MeasureSource::new(
            Arc::clone(&fixture.source_archive),
            Arc::clone(&source_counters),
        ));
        let package = source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
        let producer = GeneratedParagraphProducer::new(
            fixture.authored,
            Arc::clone(&authored_counters),
            Arc::clone(&replay_counters),
        );
        let edit = package.tail_append_plain_paragraphs_from_producer(producer, store, limits)?;
        let plan = edit.prepare()?;
        let proof = proof_record(
            plan.source_proof(),
            plan.candidate_proof(),
            plan.authored_proof(),
            fixture.authored,
            &fixture.authored_record,
        );
        let mut output = Vec::new();
        let publication = plan.write_to_stream(&mut output)?;
        let candidate_bytes: Arc<[u8]> = Arc::from(output.clone());
        let candidate_source = Arc::new(MeasureSource::new(
            Arc::clone(&candidate_bytes),
            Arc::new(SourceCounters::default()),
        ));
        let candidate_package =
            source_backed::Package::from_read_at(Arc::clone(&candidate_source) as Arc<dyn ReadAt>)?;
        let mut inverse = Vec::new();
        publication.write_inverse_to_stream(&candidate_package, &mut inverse)?;
        (output, inverse, proof)
    };
    let candidate_reader = ArchiveReader::new(&candidate_archive)?;
    let candidate_xml = candidate_reader.read(MAIN_PATH)?;
    let expected_candidate_xml = independent_candidate_xml(
        &source_main_xml(fixture.source_record.semantic.paragraph_count)?,
        fixture.authored,
    )?;
    let oracle = verify_candidate_oracles(
        fixture.source_archive.as_ref(),
        &source_main_xml(fixture.source_record.semantic.paragraph_count)?,
        &candidate_archive,
        &expected_candidate_xml,
        fixture.source_record.semantic.paragraph_count,
        fixture.authored,
        &inverse,
    )?;
    verify_proof(
        &proof,
        &source_main_xml(fixture.source_record.semantic.paragraph_count)?,
        &candidate_xml,
        &fixture.authored_record,
    )?;
    if candidate_archive.as_slice() != fixture.candidate_archive.as_ref()
        || oracle.candidate_archive_sha256 != fixture.oracle.candidate_archive_sha256
    {
        return Err("explicit replay-store preflight differs from deterministic candidate".into());
    }
    let authored = authored_counters.snapshot();
    let replay = replay_counters.snapshot();
    check_authored_counters(
        &fixture.authored_record,
        authored,
        config.authored_provider.authored_opens(),
    )?;
    check_replay_counters(&fixture.authored_record, replay, config.authored_provider)?;
    Ok(candidate_archive)
}

fn elapsed_ns(start: Instant) -> BenchResult<u64> {
    u64::try_from(start.elapsed().as_nanos())
        .map_err(|_| "DOCX replay elapsed time overflows u64 nanoseconds".into())
}

/// Execute one of the after-only publication arms.  The returned publication
/// proof is extracted and the owning publication is dropped before the caller
/// stops its lifecycle clock.  No candidate archive bytes are retained by the
/// counting sink or by this helper.
fn publish_alternative_plan(
    plan: ParagraphStreamPlan<'_>,
    publication: PublicationMode,
    sink_write_bytes: usize,
    atomic_path: Option<&Path>,
) -> BenchResult<(SinkObservation, TimedArtifactProof)> {
    match publication {
        PublicationMode::CountingSink => {
            let mut output = CountingSink::new(sink_write_bytes)?;
            let publication = plan.write_to_stream(&mut output)?;
            let proof = TimedArtifactProof::from_publication(&publication);
            drop(publication);
            Ok((output.finish(), proof))
        },
        PublicationMode::AtomicPath => {
            let path = atomic_path.ok_or("atomic publication has no prepared destination")?;
            let publication = plan.commit().write_to_path(path)?;
            let proof = TimedArtifactProof::from_publication(&publication);
            drop(publication);
            Ok((SinkObservation::atomic(proof.length), proof))
        },
        PublicationMode::HashingSink => {
            Err("hashing-sink must use the historical publication path".into())
        },
    }
}

fn run_iteration(
    fixture: &Fixture,
    config: &Config,
    input: Option<&PreparedInput>,
    sink_write_bytes: usize,
) -> BenchResult<(
    u64,
    SinkObservation,
    ReadObservation,
    AuthoredObservation,
    Option<ReplayObservation>,
    Option<PublicationRecord>,
    Option<allocation_metrics::Sample>,
    Option<process_metrics::Delta>,
)> {
    let source_counters = Arc::new(SourceCounters::default());
    let authored_counters = Arc::new(AuthoredCounters::default());
    let replay_counters = Arc::new(ReplayCounters::default());
    let limits = limits_for_provider(fixture, config)?;
    let file_path = if matches!(config.authored_provider, AuthoredProvider::FileStore) {
        Some(file_replay_path(fixture, config)?)
    } else {
        None
    };
    let file_cleanup = file_path
        .as_ref()
        .map(|_| Arc::new(Mutex::new(None::<FileReplayCleanupRecord>)));
    let file_monitor = file_path.as_ref().map(|_| FileReplayMonitor::new());
    let mut atomic_destination = if matches!(config.publication, PublicationMode::AtomicPath) {
        Some(AtomicDestination::prepare()?)
    } else {
        None
    };
    let process_before = process_metrics::Snapshot::read().ok();
    let region = allocation_metrics::begin();
    let start = Instant::now();
    let (sink, reads, authored, replay, timed_proof) = {
        let opened_input = input
            .map(|prepared| prepared.profile.open(&prepared.capability))
            .transpose()?;
        let source: Arc<dyn ReadAt> = if let Some(opened) = opened_input.as_ref() {
            Arc::new(ProfiledMeasureSource::new(
                Arc::clone(opened),
                Arc::clone(&source_counters),
            ))
        } else {
            Arc::new(MeasureSource::new(
                Arc::clone(&fixture.source_archive),
                Arc::clone(&source_counters),
            ))
        };
        let package = source_backed::Package::from_read_at(Arc::clone(&source))?;
        let atomic_path = atomic_destination.as_ref().map(AtomicDestination::path);
        let (replay, sink, timed_proof) =
            if matches!(config.authored_provider, AuthoredProvider::Deterministic) {
                let generated =
                    GeneratedParagraphSource::new(fixture.authored, Arc::clone(&authored_counters));
                let edit = package.tail_append_plain_paragraphs(generated, limits);
                let plan = edit.prepare()?;
                if config.publication.is_default() {
                    let mut output = HashingSink::new(sink_write_bytes)?;
                    let _publication = plan.write_to_stream(&mut output)?;
                    (None, output.finish(), None)
                } else {
                    let (output, proof) = publish_alternative_plan(
                        plan,
                        config.publication,
                        sink_write_bytes,
                        atomic_path,
                    )?;
                    (None, output, Some(proof))
                }
            } else {
                let producer = GeneratedParagraphProducer::new(
                    fixture.authored,
                    Arc::clone(&authored_counters),
                    Arc::clone(&replay_counters),
                );
                let store = match config.authored_provider {
                    AuthoredProvider::MemoryStore => {
                        ReplayRouteStore::Memory(CountingReplayStore::new(
                            MemoryReplayStore::new(limits.max_replay_bytes)?,
                            Arc::clone(&replay_counters),
                        ))
                    },
                    AuthoredProvider::FileStore => {
                        let cleanup = file_cleanup
                            .as_ref()
                            .ok_or("file-store route has no cleanup owner")?;
                        let path = file_path
                            .as_ref()
                            .ok_or("file-store route has no replay path")?;
                        let monitor = file_monitor
                            .as_ref()
                            .ok_or("file-store route has no monitor")?;
                        ReplayRouteStore::File(CountingReplayStore::new(
                            FileReplayRouteStore::new(
                                path.clone(),
                                limits.max_replay_bytes,
                                match config.replay_sync {
                                    ReplaySync::None => FileSyncPolicy::None,
                                    ReplaySync::Data => FileSyncPolicy::Data,
                                },
                                Arc::clone(cleanup),
                                monitor,
                            )?,
                            Arc::clone(&replay_counters),
                        ))
                    },
                    AuthoredProvider::Deterministic => {
                        return Err(
                            "deterministic provider selected an explicit store route".into()
                        );
                    },
                };
                let edit =
                    package.tail_append_plain_paragraphs_from_producer(producer, store, limits)?;
                let plan = edit.prepare()?;
                if config.publication.is_default() {
                    let mut output = HashingSink::new(sink_write_bytes)?;
                    let _publication = plan.write_to_stream(&mut output)?;
                    (Some(replay_counters.snapshot()), output.finish(), None)
                } else {
                    let (output, proof) = publish_alternative_plan(
                        plan,
                        config.publication,
                        sink_write_bytes,
                        atomic_path,
                    )?;
                    (Some(replay_counters.snapshot()), output, Some(proof))
                }
            };
        let reads = source_counters.snapshot();
        let authored = authored_counters.snapshot();
        (sink, reads, authored, replay, timed_proof)
    };
    let elapsed = elapsed_ns(start)?;
    let allocation = region.finish();
    let process = process_before
        .zip(process_metrics::Snapshot::read().ok())
        .map(|(before, after)| after.delta(before));
    let publication = match (config.publication, timed_proof) {
        (PublicationMode::HashingSink, None) => None,
        (PublicationMode::CountingSink, Some(proof)) => Some(PublicationRecord {
            schema: PUBLICATION_SCHEMA,
            route: PublicationMode::CountingSink,
            timing_scope: "source_admission_prepare_sequential_sink_publication_drop",
            timed_candidate_artifact_bytes: proof.length,
            timed_candidate_artifact_sha256: hex_digest(&proof.sha256),
            timed_candidate_matches_oracle: proof.length
                == u64::try_from(fixture.oracle.candidate_archive_bytes)?
                && hex_digest(&proof.sha256) == fixture.oracle.candidate_archive_sha256,
            verification_scope: "timed_production_artifact_proof_plus_untimed_candidate_oracle",
            atomic: None,
        }),
        (PublicationMode::AtomicPath, Some(proof)) => {
            let destination = atomic_destination
                .as_mut()
                .ok_or("atomic publication has no destination owner")?;
            let atomic = destination.verify_and_cleanup(fixture, proof)?;
            Some(PublicationRecord {
                schema: PUBLICATION_SCHEMA,
                route: PublicationMode::AtomicPath,
                timing_scope: "source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop",
                timed_candidate_artifact_bytes: proof.length,
                timed_candidate_artifact_sha256: hex_digest(&proof.sha256),
                timed_candidate_matches_oracle: proof.length
                    == u64::try_from(fixture.oracle.candidate_archive_bytes)?
                    && hex_digest(&proof.sha256) == fixture.oracle.candidate_archive_sha256,
                verification_scope: "timed_production_artifact_proof_plus_post_timer_path_oracle",
                atomic: Some(atomic),
            })
        },
        (_, None) => {
            return Err("publication route did not return its production artifact proof".into());
        },
        (PublicationMode::HashingSink, Some(_)) => {
            return Err("historical hashing sink unexpectedly returned an alternate proof".into());
        },
    };
    let replay = match (replay, file_cleanup) {
        (Some(mut replay), Some(cleanup)) => {
            let file = cleanup_file_route(
                cleanup,
                file_path
                    .as_ref()
                    .ok_or("file-store route has no replay path")?,
                file_monitor
                    .as_ref()
                    .ok_or("file-store route has no monitor")?,
            )?
            .ok_or("file-store route did not produce a cleanup record")?;
            replay.route = Some(config.authored_provider);
            replay.file_logical_bytes = file.logical_bytes;
            replay.file_allocated_bytes = file.allocated_bytes;
            replay.file_write_calls = Some(file.write_calls);
            replay.file_sync_calls = Some(file.sync_calls);
            replay.file_cleanup_verified = Some(file.cleanup_verified);
            replay.seal_sha256_checks = Some(file.seal_sha256_checks);
            replay.cleanup_sha256_checks = Some(file.cleanup_sha256_checks);
            replay.durable_reference_kind = Some("file");
            replay.file = Some(file);
            Some(replay)
        },
        (Some(mut replay), None) => {
            replay.route = Some(config.authored_provider);
            replay.retained_logical_bytes = Some(fixture.authored_record.encoded_xml_bytes);
            replay.retained_capacity_bytes = Some(limits.max_replay_bytes);
            replay.retained_capacity_provenance = Some("exact_reserve_equal_to_ceiling");
            replay.durable_reference_kind = Some("none");
            replay.seal_sha256_checks = Some(0);
            replay.cleanup_sha256_checks = Some(0);
            Some(replay)
        },
        (None, None) => None,
        (None, Some(cleanup)) => {
            let _ = cleanup_file_route(
                cleanup,
                file_path
                    .as_ref()
                    .ok_or("file-store route has no replay path")?,
                file_monitor
                    .as_ref()
                    .ok_or("file-store route has no monitor")?,
            )?;
            None
        },
    };
    Ok((
        elapsed,
        sink,
        reads,
        authored,
        replay,
        publication,
        allocation,
        process,
    ))
}

fn check_authored_counters(
    expected: &AuthoredRecord,
    authored: AuthoredObservation,
    expected_opens: u64,
) -> BenchResult<()> {
    if authored.opens != expected_opens {
        return Err(format!(
            "DOCX replay authored-provider opened {} passes; expected {expected_opens}",
            authored.opens,
        )
        .into());
    }
    // A one-shot producer emits one pass without opening a replay cursor.
    let event_passes = if expected_opens == 0 {
        1
    } else {
        expected_opens
    };
    let expected_events = expected
        .event_count
        .checked_mul(event_passes)
        .ok_or("DOCX replay authored event counter overflow")?;
    let expected_text_bytes = expected
        .text_bytes
        .checked_mul(event_passes)
        .ok_or("DOCX replay authored text counter overflow")?;
    let expected_text_chunks = expected
        .text_chunk_count
        .checked_mul(event_passes)
        .ok_or("DOCX replay authored text-chunk counter overflow")?;
    if authored.events != expected_events
        || authored.text_chunks != expected_text_chunks
        || authored.text_bytes != expected_text_bytes
    {
        return Err(format!(
            "DOCX replay authored-provider counters disagree with proof: events {} != {expected_events}, chunks {} != {expected_text_chunks}, or text {} != {expected_text_bytes}",
            authored.events, authored.text_chunks, authored.text_bytes
        )
        .into());
    }
    Ok(())
}

fn check_replay_counters(
    expected: &AuthoredRecord,
    replay: ReplayObservation,
    provider: AuthoredProvider,
) -> BenchResult<()> {
    check_replay_counters_with_file(expected, replay, provider, false)
}

fn check_replay_counters_with_file(
    expected: &AuthoredRecord,
    replay: ReplayObservation,
    provider: AuthoredProvider,
    require_file_observation: bool,
) -> BenchResult<()> {
    if !provider.is_store() {
        return Ok(());
    }
    let expected_replayed_bytes = expected
        .encoded_xml_bytes
        .checked_mul(provider.replay_opens())
        .ok_or("DOCX replay returned-byte counter overflow")?;
    if replay.producer_invocations != 1
        || replay.prepare_calls != 1
        || replay.append_calls == 0
        || replay.appended_bytes != expected.encoded_xml_bytes
        || replay.store_finish_calls != 1
        || replay.replay_opens != provider.replay_opens()
        || replay.replay_read_calls == 0
        || replay.replay_returned_bytes != expected_replayed_bytes
        || replay.replay_finish_calls != provider.replay_opens()
        || replay.replay_sha256_checks != provider.replay_opens()
    {
        return Err(format!(
            "DOCX replay store counters disagree with proof: producer {}, prepare {}, appended {}, finish {}, opens {}, reads {}, returned {}, reader_finish {}, hash_checks {}; expected appended {}, opens {}, returned {}, reader_finish {}, hash_checks {}",
            replay.producer_invocations,
            replay.prepare_calls,
            replay.appended_bytes,
            replay.store_finish_calls,
            replay.replay_opens,
            replay.replay_read_calls,
            replay.replay_returned_bytes,
            replay.replay_finish_calls,
            replay.replay_sha256_checks,
            expected.encoded_xml_bytes,
            provider.replay_opens(),
            expected_replayed_bytes,
            provider.replay_opens(),
            provider.replay_opens(),
        )
        .into());
    }
    match provider {
        AuthoredProvider::FileStore => {
            if require_file_observation {
                let Some(file) = replay.file else {
                    return Err("file-store route omitted file observations".into());
                };
                if file.write_calls == 0
                    || file.read_calls == 0
                    || file.returned_bytes != expected_replayed_bytes
                    || file.replay_sha256_checks != provider.replay_opens()
                    || file.seal_sha256_checks != 1
                    || file.cleanup_sha256_checks != 1
                    || file.logical_bytes != Some(expected.encoded_xml_bytes)
                    || !file.cleanup_verified
                {
                    return Err("file-store retention or cleanup observations failed".into());
                }
            }
        },
        AuthoredProvider::MemoryStore => {
            if replay.file.is_some() {
                return Err("memory-store route reported file observations".into());
            }
        },
        AuthoredProvider::Deterministic => {},
    }
    Ok(())
}

fn check_runtime(
    fixture: &Fixture,
    provider: AuthoredProvider,
    sink: SinkObservation,
    reads: ReadObservation,
    authored: AuthoredObservation,
    replay: Option<ReplayObservation>,
    publication: Option<&PublicationRecord>,
) -> BenchResult<()> {
    let expected_bytes = u64::try_from(fixture.oracle.candidate_archive_bytes)?;
    let expected_sha256 = Sha256::digest(fixture.candidate_archive.as_ref());
    if sink.observed {
        if sink.accepted_bytes != expected_bytes {
            return Err(
                "DOCX replay sink output length differs from candidate archive oracle".into(),
            );
        }
        if let Some(digest) = sink.digest {
            if digest.as_slice() != expected_sha256.as_slice() {
                return Err(
                    "DOCX replay hashing sink output differs from candidate archive oracle".into(),
                );
            }
        } else if publication.is_none_or(|value| !value.timed_candidate_matches_oracle) {
            return Err(
                "DOCX replay counting sink lacks a matching production artifact proof".into(),
            );
        }
    } else {
        let publication = publication.ok_or("DOCX replay publication route omitted its proof")?;
        if !publication.timed_candidate_matches_oracle {
            return Err(
                "DOCX replay atomic publication proof differs from candidate archive oracle".into(),
            );
        }
        let atomic = publication
            .atomic
            .as_ref()
            .ok_or("DOCX replay atomic route omitted its post-timer path oracle")?;
        if !atomic.output_bytes_exact
            || !atomic.output_sha256_exact
            || !atomic.post_timer_oracle.candidate_xml_exact
            || !atomic.post_timer_oracle.candidate_semantic_exact
            || !atomic.post_timer_oracle.untouched_member_metadata_exact
            || !atomic.post_timer_oracle.untouched_raw_members_preserved
            || !atomic.post_timer_oracle.physical_order_exact
            || !atomic.post_timer_oracle.opaque_member_exact
            || !atomic.post_timer_oracle.source_unchanged
            || !atomic.post_timer_oracle.inverse_exact
            || !atomic.cleanup.destination_removed
            || !atomic.cleanup.parent_removed
        {
            return Err("DOCX replay atomic destination post-timer oracle failed".into());
        }
    }
    if reads.calls == 0 || reads.returned_bytes == 0 {
        return Err("DOCX replay lifecycle performed no positional source reads".into());
    }
    check_authored_counters(
        &fixture.authored_record,
        authored,
        provider.authored_opens(),
    )?;
    if provider.is_store() {
        check_replay_counters_with_file(
            &fixture.authored_record,
            replay.ok_or("store route did not report replay counters")?,
            provider,
            true,
        )?;
    } else if replay.is_some() {
        return Err("deterministic route unexpectedly reported replay counters".into());
    }
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
    let fixture = build_fixture(source_count, authored_spec, config.compression)?;
    let limits = limits_for_provider(&fixture, config)?;
    let preflight_candidate = if config.authored_provider.is_store() {
        // Execute one complete route with a materialized output and inverse
        // before timing.  This keeps all semantic/raw ZIP checks outside the
        // measurement while proving that the selected store has equivalent
        // candidate bytes and authenticated replay facts.
        Some(verify_store_preflight(&fixture, config, limits)?)
    } else {
        None
    };
    let prepared_input = prepare_input(&fixture, config)?;
    if let Some(dir) = config.fixture_dir.as_deref() {
        // Fixture export is deliberate, opt-in, and outside every timed
        // iteration.  The files are the same bytes used by the independent
        // archive/oracle checks above.
        export_fixture(
            dir,
            &fixture,
            config.authored_provider,
            preflight_candidate.as_deref(),
        )?;
    }
    drop(preflight_candidate);
    let mut samples = Vec::with_capacity(config.samples);
    for iteration in 0..config.warmups.saturating_add(config.samples) {
        let (elapsed, sink, reads, authored, replay, publication, allocation, process) =
            run_iteration(
                &fixture,
                config,
                prepared_input.as_ref(),
                config.sink_write_bytes,
            )?;
        if let Some(input) = prepared_input.as_ref()
            && input.capability.is_file()
        {
            // Revalidate the pinned descriptor after the timed operation.  A
            // full hash here is setup/post-sample evidence and is deliberately
            // outside the timed lifecycle; native and owned input paths do not
            // acquire this extra pass.
            input.capability.verify_fingerprint()?;
        }
        check_runtime(
            &fixture,
            config.authored_provider,
            sink,
            reads,
            authored,
            replay,
            publication.as_ref(),
        )?;
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
                replay,
                sink: sink.record(),
                allocation,
                process,
                publication,
            });
        }
    }
    Ok(CaseRecord {
        source_count,
        authored_count,
        chunk_mode,
        text_mode,
        provider: config.authored_provider,
        replay_max_bytes: config.replay_max_bytes,
        compression: config.compression,
        input_mode: input_mode_name(prepared_input.as_ref()),
        input_storage_kind: input_storage_kind(prepared_input.as_ref()),
        input_identity_validation: input_identity_validation(prepared_input.as_ref()),
        sink_write_bytes: config.sink_write_bytes,
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

fn parse_u64_positive(value: &OsString, flag: &str) -> BenchResult<u64> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let parsed = text
        .parse::<u64>()
        .map_err(|error| format!("invalid {flag} value {text:?}: {error}"))?;
    if parsed == 0 || parsed == u64::MAX {
        return Err(format!("{flag} must be finite and positive").into());
    }
    Ok(parsed)
}

fn parse_u64_nonnegative(value: &OsString, flag: &str) -> BenchResult<u64> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    text.parse::<u64>()
        .map_err(|error| format!("invalid {flag} value {text:?}: {error}").into())
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
    let mut authored_provider = AuthoredProvider::Deterministic;
    let mut replay_dir = None;
    let mut replay_max_bytes = None;
    let mut replay_sync = ReplaySync::None;
    let mut compression = CompressionProfile::Current;
    let mut input_mode = None;
    let mut input_file = None;
    let mut input_max_range_bytes = None;
    let mut input_delay_us = 0_u64;
    let mut input_overhead_us = 0_u64;
    let mut input_bytes_per_second = None;
    let mut publication = PublicationMode::HashingSink;
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
                    parse_positive(&next_value(&mut values, "--sink-write")?, "--sink-write")?;
                if !SINK_WRITE_BYTES.contains(&sink_write_bytes) {
                    return Err(format!(
                        "unsupported --sink-write value {sink_write_bytes}; expected one of {SINK_WRITE_BYTES:?}"
                    )
                    .into());
                }
            },
            "--publication" | "--publication-mode" => {
                publication = PublicationMode::parse(
                    next_value(&mut values, "--publication")?
                        .to_str()
                        .ok_or("--publication must be UTF-8")?,
                )?
            },
            "--json" => json_path = Some(PathBuf::from(next_value(&mut values, "--json")?)),
            "--fixture-dir" => {
                fixture_dir = Some(PathBuf::from(next_value(&mut values, "--fixture-dir")?))
            },
            "--authored-provider" => {
                authored_provider = AuthoredProvider::parse(
                    next_value(&mut values, "--authored-provider")?
                        .to_str()
                        .ok_or("--authored-provider must be UTF-8")?,
                )?
            },
            "--replay-dir" => {
                replay_dir = Some(PathBuf::from(next_value(&mut values, "--replay-dir")?))
            },
            "--replay-max-bytes" => {
                replay_max_bytes = Some(parse_u64_positive(
                    &next_value(&mut values, "--replay-max-bytes")?,
                    "--replay-max-bytes",
                )?)
            },
            "--replay-sync" => {
                replay_sync = ReplaySync::parse(
                    next_value(&mut values, "--replay-sync")?
                        .to_str()
                        .ok_or("--replay-sync must be UTF-8")?,
                )?
            },
            "--compression" | "--compression-mode" => {
                compression = CompressionProfile::parse(
                    next_value(&mut values, "--compression")?
                        .to_str()
                        .ok_or("--compression must be UTF-8")?,
                )?
            },
            "--input-mode" => {
                input_mode = Some(InputMode::parse(
                    next_value(&mut values, "--input-mode")?
                        .to_str()
                        .ok_or("--input-mode must be UTF-8")?,
                )?)
            },
            "--input-file" => {
                input_file = Some(PathBuf::from(next_value(&mut values, "--input-file")?))
            },
            "--input-max-range" => {
                input_max_range_bytes = Some(parse_positive(
                    &next_value(&mut values, "--input-max-range")?,
                    "--input-max-range",
                )?)
            },
            "--input-delay-us" => {
                input_delay_us = parse_u64_nonnegative(
                    &next_value(&mut values, "--input-delay-us")?,
                    "--input-delay-us",
                )?
            },
            "--input-overhead-us" => {
                input_overhead_us = parse_u64_nonnegative(
                    &next_value(&mut values, "--input-overhead-us")?,
                    "--input-overhead-us",
                )?
            },
            "--input-bytes-per-second" => {
                input_bytes_per_second = Some(parse_u64_positive(
                    &next_value(&mut values, "--input-bytes-per-second")?,
                    "--input-bytes-per-second",
                )?)
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
    if authored_provider.is_store() && replay_max_bytes.is_none() {
        return Err("store provider requires --replay-max-bytes".into());
    }
    if !matches!(authored_provider, AuthoredProvider::FileStore) && replay_dir.is_some() {
        return Err("--replay-dir is only valid with --authored-provider file-store".into());
    }
    if !authored_provider.is_store() && !matches!(replay_sync, ReplaySync::None) {
        return Err("--replay-sync is only valid with a store provider".into());
    }
    if input_mode.is_none()
        && (input_file.is_some()
            || input_max_range_bytes.is_some()
            || input_delay_us != 0
            || input_overhead_us != 0
            || input_bytes_per_second.is_some())
    {
        return Err("input settings require --input-mode".into());
    }
    if !matches!(input_mode, Some(InputMode::File)) && input_file.is_some() {
        return Err("--input-file is only valid with --input-mode file".into());
    }
    if !matches!(input_mode, Some(InputMode::ShortRead | InputMode::Latency))
        && input_max_range_bytes.is_some()
    {
        return Err("--input-max-range is only valid with short-read or latency input".into());
    }
    if !matches!(input_mode, Some(InputMode::Latency))
        && (input_delay_us != 0 || input_overhead_us != 0 || input_bytes_per_second.is_some())
    {
        return Err("latency settings require --input-mode latency".into());
    }
    if matches!(input_mode, Some(InputMode::File)) && input_file.is_none() {
        return Err("file input mode requires --input-file".into());
    }
    if matches!(input_mode, Some(InputMode::ShortRead | InputMode::Latency))
        && input_max_range_bytes.is_none()
    {
        return Err("short-read and latency input require --input-max-range".into());
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
        authored_provider,
        replay_dir,
        replay_max_bytes,
        replay_sync,
        compression,
        input_mode,
        input_file,
        input_max_range_bytes,
        input_delay_us,
        input_overhead_us,
        input_bytes_per_second,
        publication,
    })
}

fn usage() -> &'static str {
    "usage: docx_replayable_tail_append [--source-counts 64,8192,131072] [--authored-counts 64,256,4096,16384] [--chunks one,64,window] [--text empty,short,near] [--samples N] [--warmups N] [--sink-write 512|4096|65536] [--publication hashing-sink|counting-sink|atomic-path] [--authored-provider deterministic|memory-store|file-store] [--replay-dir DIR] [--replay-max-bytes BYTES] [--replay-sync none|data] [--compression-mode current|store|deflate] [--input-mode owned|file|short-read|latency] [--input-file PATH] [--input-max-range BYTES] [--input-delay-us N] [--input-overhead-us N] [--input-bytes-per-second N] [--json PATH] [--fixture-dir DIR]"
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
            expected_authored_opens: config.authored_provider.authored_opens(),
            expected_replay_opens: config.authored_provider.replay_opens(),
            source: source_description(config.input_mode),
            authored_provider: config.authored_provider.name(),
            provider: config.authored_provider,
            replay_max_bytes: config.replay_max_bytes,
            replay_dir: config
                .replay_dir
                .as_ref()
                .map(|path| path.display().to_string()),
            replay_sync: config.replay_sync,
            compression: config.compression,
            input_mode: config.input_mode.map_or("native_owned", InputMode::as_str),
            input_storage_kind: config.input_mode.map_or("owned", |mode| match mode {
                InputMode::File => "file",
                InputMode::Owned | InputMode::ShortRead | InputMode::Latency => "owned",
            }),
            input_identity_validation: config.input_mode.map_or(
                "native_source_archive_oracle",
                |mode| match mode {
                    InputMode::File => "setup_and_post_sample_fingerprint_outside_timing",
                    InputMode::Owned | InputMode::ShortRead | InputMode::Latency => {
                        "setup_fingerprint_outside_timing"
                    },
                },
            ),
            input_max_range_bytes: config.input_max_range_bytes,
            input_delay_us: config.input_delay_us,
            input_overhead_us: config.input_overhead_us,
            input_bytes_per_second: config.input_bytes_per_second,
            sink: match config.publication {
                PublicationMode::HashingSink => {
                    "non_seek_hashing_sha256_short_write_no_archive_retention"
                },
                PublicationMode::CountingSink => {
                    "non_seek_counting_short_write_no_archive_retention_production_artifact_proof"
                },
                PublicationMode::AtomicPath => {
                    "production_atomic_path_output_proof_post_timer_destination_oracle"
                },
            },
            fixture_dir: config
                .fixture_dir
                .as_ref()
                .map(|path| path.display().to_string()),
            publication: (!config.publication.is_default()).then_some(config.publication),
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
fn verify_candidate_xml_only(actual: &[u8], expected: &[u8]) -> BenchResult<()> {
    if actual != expected {
        return Err("candidate XML differs from independent oracle".into());
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
        assert!(check_authored_counters(&expected, valid, EXPECTED_AUTHORED_OPENS).is_ok());
        let mut partial = valid;
        partial.events -= 1;
        assert!(check_authored_counters(&expected, partial, EXPECTED_AUTHORED_OPENS).is_err());
        let mut extra = valid;
        extra.events += 1;
        assert!(check_authored_counters(&expected, extra, EXPECTED_AUTHORED_OPENS).is_err());
    }

    #[test]
    fn replay_counter_oracle_rejects_partial_and_extra_empty_store_passes() {
        let expected = authored_record(AuthoredSpec {
            count: 2,
            chunk_mode: ChunkMode::Fixed64,
            text_mode: TextMode::Empty,
        })
        .expect("empty authored proof");
        let valid = ReplayObservation {
            producer_invocations: 1,
            prepare_calls: 1,
            append_calls: 1,
            appended_bytes: expected.encoded_xml_bytes,
            store_finish_calls: 1,
            replay_opens: EXPECTED_STORE_REPLAY_OPENS,
            replay_read_calls: 4,
            replay_requested_bytes: expected.encoded_xml_bytes * 4,
            replay_returned_bytes: expected.encoded_xml_bytes * 4,
            replay_finish_calls: EXPECTED_STORE_REPLAY_OPENS,
            replay_sha256_checks: EXPECTED_STORE_REPLAY_OPENS,
            ..ReplayObservation::default()
        };
        assert!(check_replay_counters(&expected, valid, AuthoredProvider::MemoryStore).is_ok());
        let mut partial = valid;
        partial.replay_opens -= 1;
        assert!(check_replay_counters(&expected, partial, AuthoredProvider::MemoryStore).is_err());
        let mut extra = valid;
        extra.replay_returned_bytes += 1;
        assert!(check_replay_counters(&expected, extra, AuthoredProvider::MemoryStore).is_err());
        let mut missing_proof = valid;
        missing_proof.replay_sha256_checks -= 1;
        assert!(
            check_replay_counters(&expected, missing_proof, AuthoredProvider::MemoryStore).is_err()
        );
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
            CompressionProfile::Current,
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
            CompressionProfile::Current,
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

    #[test]
    fn explicit_store_and_deflate_profiles_preserve_archive_semantics() {
        let xml = source_main_xml(2).expect("source XML");
        let current = source_archive_with_compression(&xml, 0, CompressionProfile::Current)
            .expect("current source archive");
        let current_members = member_identities(&current).expect("current member identities");
        let current_opaque = current_members
            .iter()
            .find(|member| member.path == OPAQUE_PATH)
            .expect("current opaque member")
            .clone();
        for profile in [CompressionProfile::Store, CompressionProfile::Deflate] {
            let archive =
                source_archive_with_compression(&xml, 0, profile).expect("profiled source archive");
            let reader = ArchiveReader::new(&archive).expect("archive reader");
            assert_eq!(reader.read(MAIN_PATH).expect("main XML"), xml);
            assert_eq!(
                reader.read(OPAQUE_PATH).expect("opaque payload").len(),
                OPAQUE_BYTES
            );
            let identities = member_identities(&archive).expect("member identities");
            let main = identities
                .iter()
                .find(|member| member.path == MAIN_PATH)
                .expect("profiled main member");
            assert_eq!(
                main.compression_method,
                match profile {
                    CompressionProfile::Store => "Store",
                    CompressionProfile::Deflate => "Deflate",
                    CompressionProfile::Current => unreachable!(),
                }
            );
            let opaque = identities
                .iter()
                .find(|member| member.path == OPAQUE_PATH)
                .expect("profiled opaque member");
            assert_eq!(opaque, &current_opaque);
        }
    }
}
