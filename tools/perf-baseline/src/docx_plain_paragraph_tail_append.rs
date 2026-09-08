//! Source-backed DOCX plain-paragraph tail-append evidence.
//!
//! This harness measures the existing public `source_backed` paragraph-copy
//! transaction.  It deliberately keeps corpus construction, complete output
//! reopen, ZIP member identity, semantic, and reversible-patch checks outside
//! the measured operation.  The measured total owns the source adapter,
//! source-backed package, snapshot/edit/commit/publication, and hash-only
//! non-seek sink for the complete lifecycle, including their drops.
//!
//! `--mode phases` is a separate execution of the same lifecycle.  It uses
//! one non-nested allocator region per phase so the six phase high-water marks
//! are useful attribution evidence without being mistaken for the total
//! operation high-water mark.

use std::{
    ffi::OsString,
    fmt::Write as FmtWrite,
    fs::OpenOptions,
    io::{self, Cursor, Write},
    path::PathBuf,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::Instant,
};

use litchi_core::{OwnedSource, Position, ReadAt, SourceVersion};
use litchi_docx::source_backed::paragraph_copy::{
    Commit, Edit, Error as CopyError, Limits, Patch, Publication, Snapshot,
};
use litchi_docx::{Package as OwnedDocxPackage, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use serde::Serialize;
use sha2::{Digest as _, Sha256};
use soapberry_zip::ZipArchive;
use soapberry_zip::office::ArchiveReader;

use super::{allocation_metrics, process_metrics};

const SCHEMA: &str = "docx-plain-paragraph-tail-append-v1";
const DEFAULT_COUNTS: [usize; 3] = [64, 8_192, 131_072];
const DEFAULT_SAMPLES: usize = 30;
const DEFAULT_WARMUPS: usize = 3;
const OPAQUE_PATH: &str = "word/perf-opaque.bin";
const OPAQUE_BYTES: usize = 32 * 1024;
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MAIN_PATH: &str = "word/document.xml";
const OPAQUE_MEDIA_TYPE: &str = "application/octet-stream";
const GENERATOR: &str = "litchi-docx-plain-paragraph-tail-append-v1";
const HASH_SINK_MAX_WRITE: usize = 16 * 1024;

type BenchResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum Mode {
    Total,
    Phases,
}

impl Mode {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "total" => Ok(Self::Total),
            "phases" => Ok(Self::Phases),
            _ => Err(format!("invalid --mode {value:?}; expected total or phases").into()),
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    counts: Vec<usize>,
    samples: usize,
    warmups: usize,
    mode: Mode,
    json_path: Option<PathBuf>,
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
    counts: Vec<usize>,
    samples: usize,
    warmups: usize,
    mode: Mode,
    lifecycle_phases: [&'static str; 6],
    sink: &'static str,
    source: &'static str,
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
struct LimitsRecord {
    max_xml_bytes: usize,
    max_paragraphs: usize,
    max_events: usize,
    max_depth: usize,
    max_output_bytes: usize,
    max_durable_bytes: usize,
}

#[derive(Clone, Debug, Serialize)]
struct PatchOracle {
    copied_source_position: usize,
    copied_before_position: usize,
    copied_paragraphs: usize,
    copied_bytes: usize,
    durable_bytes: usize,
    durable_canonical_verified: bool,
    replay_verified: bool,
    inverse_verified: bool,
    stale_source_refusal_verified: bool,
    publication_inverse_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct CorpusRecord {
    generator: &'static str,
    format: &'static str,
    count: usize,
    opaque_path: &'static str,
    opaque_bytes: usize,
    source_archive_bytes: usize,
    source_archive_sha256: String,
    candidate_archive_bytes: usize,
    candidate_archive_sha256: String,
    source_main_xml_bytes: usize,
    source_main_xml_sha256: String,
    candidate_main_xml_bytes: usize,
    candidate_main_xml_sha256: String,
    source_main_xml_archive_verified: bool,
    candidate_main_xml_archive_verified: bool,
    candidate_main_xml_oracle_verified: bool,
    source_members: Vec<MemberIdentity>,
    candidate_members: Vec<MemberIdentity>,
    source_member_count: usize,
    candidate_member_count: usize,
    source_semantic: SemanticRecord,
    candidate_semantic: SemanticRecord,
    source_unchanged_verified: bool,
    source_semantic_reopen_verified: bool,
    candidate_semantic_reopen_verified: bool,
    tail_copy_exactly_one_verified: bool,
    untouched_members_verified: bool,
    opaque_member_exact_verified: bool,
    physical_order_verified: bool,
    limits: LimitsRecord,
    patch: PatchOracle,
}

#[derive(Debug)]
struct Corpus {
    count: usize,
    source_archive: Arc<[u8]>,
    candidate_archive: Arc<[u8]>,
    limits: Limits,
    record: CorpusRecord,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct SinkHistogram {
    bytes_0: u64,
    bytes_1_to_512: u64,
    bytes_513_to_4096: u64,
    bytes_4097_to_16384: u64,
    bytes_16385_to_65536: u64,
    bytes_over_65536: u64,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct ReadRequestHistogram {
    bytes_0: u64,
    bytes_1_to_512: u64,
    bytes_513_to_4096: u64,
    bytes_4097_to_16384: u64,
    bytes_16385_to_65536: u64,
    bytes_over_65536: u64,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct ReadObservation {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    request_histogram: ReadRequestHistogram,
}

#[derive(Debug, Default)]
struct AtomicReadRequestHistogram {
    bytes_0: AtomicU64,
    bytes_1_to_512: AtomicU64,
    bytes_513_to_4096: AtomicU64,
    bytes_4097_to_16384: AtomicU64,
    bytes_16385_to_65536: AtomicU64,
    bytes_over_65536: AtomicU64,
}

impl AtomicReadRequestHistogram {
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

    fn snapshot(&self) -> ReadRequestHistogram {
        ReadRequestHistogram {
            bytes_0: self.bytes_0.load(Ordering::Relaxed),
            bytes_1_to_512: self.bytes_1_to_512.load(Ordering::Relaxed),
            bytes_513_to_4096: self.bytes_513_to_4096.load(Ordering::Relaxed),
            bytes_4097_to_16384: self.bytes_4097_to_16384.load(Ordering::Relaxed),
            bytes_16385_to_65536: self.bytes_16385_to_65536.load(Ordering::Relaxed),
            bytes_over_65536: self.bytes_over_65536.load(Ordering::Relaxed),
        }
    }
}

impl ReadObservation {
    fn delta(self, before: Self) -> Self {
        Self {
            calls: self.calls.saturating_sub(before.calls),
            requested_bytes: self.requested_bytes.saturating_sub(before.requested_bytes),
            returned_bytes: self.returned_bytes.saturating_sub(before.returned_bytes),
            request_histogram: ReadRequestHistogram {
                bytes_0: self
                    .request_histogram
                    .bytes_0
                    .saturating_sub(before.request_histogram.bytes_0),
                bytes_1_to_512: self
                    .request_histogram
                    .bytes_1_to_512
                    .saturating_sub(before.request_histogram.bytes_1_to_512),
                bytes_513_to_4096: self
                    .request_histogram
                    .bytes_513_to_4096
                    .saturating_sub(before.request_histogram.bytes_513_to_4096),
                bytes_4097_to_16384: self
                    .request_histogram
                    .bytes_4097_to_16384
                    .saturating_sub(before.request_histogram.bytes_4097_to_16384),
                bytes_16385_to_65536: self
                    .request_histogram
                    .bytes_16385_to_65536
                    .saturating_sub(before.request_histogram.bytes_16385_to_65536),
                bytes_over_65536: self
                    .request_histogram
                    .bytes_over_65536
                    .saturating_sub(before.request_histogram.bytes_over_65536),
            },
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct SinkObservation {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    histogram: SinkHistogram,
    digest: [u8; 32],
}

#[derive(Debug)]
struct HashingSink {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    histogram: SinkHistogram,
    digest: Sha256,
}

impl HashingSink {
    fn new() -> Self {
        Self {
            accepted_bytes: 0,
            write_calls: 0,
            largest_write: 0,
            histogram: SinkHistogram::default(),
            digest: Sha256::new(),
        }
    }

    fn finish(self) -> SinkObservation {
        let digest = self.digest.finalize();
        let mut output = [0u8; 32];
        output.copy_from_slice(&digest);
        SinkObservation {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            largest_write: self.largest_write,
            histogram: self.histogram,
            digest: output,
        }
    }
}

impl Write for HashingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let accepted = bytes.len().min(HASH_SINK_MAX_WRITE);
        let length = u64::try_from(accepted)
            .map_err(|_| io::Error::other("sink write length does not fit u64"))?;
        self.accepted_bytes = self
            .accepted_bytes
            .checked_add(length)
            .ok_or_else(|| io::Error::other("sink byte count overflow"))?;
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("sink write count overflow"))?;
        self.largest_write = self.largest_write.max(length);
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

#[derive(Debug)]
struct MeasureSource {
    bytes: Arc<[u8]>,
    id: u64,
    revision: AtomicU64,
    read_calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    request_histogram: AtomicReadRequestHistogram,
}

static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

impl MeasureSource {
    fn new(bytes: Arc<[u8]>) -> Self {
        Self {
            bytes,
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            revision: AtomicU64::new(0),
            read_calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            request_histogram: AtomicReadRequestHistogram::default(),
        }
    }

    fn reads(&self) -> ReadObservation {
        ReadObservation {
            calls: self.read_calls.load(Ordering::Relaxed),
            requested_bytes: self.requested_bytes.load(Ordering::Relaxed),
            returned_bytes: self.returned_bytes.load(Ordering::Relaxed),
            request_histogram: self.request_histogram.snapshot(),
        }
    }
}

impl ReadAt for MeasureSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source length overflow"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        self.requested_bytes.fetch_add(
            u64::try_from(output.len()).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidInput, "request size overflow")
            })?,
            Ordering::Relaxed,
        );
        self.request_histogram.record(output.len());
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source offset overflow"))?;
        if output.is_empty() || offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        self.returned_bytes.fetch_add(
            u64::try_from(count)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "read size overflow"))?,
            Ordering::Relaxed,
        );
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            self.id,
            self.revision.load(Ordering::Acquire),
        ))
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
struct TotalSample {
    sample: usize,
    elapsed_ns: u64,
    source_reads: ReadObservation,
    sink: SinkRecord,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Clone, Debug, Serialize)]
struct PhaseSample {
    elapsed_ns: u64,
    source_reads: ReadObservation,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Clone, Debug, Serialize)]
struct PhaseLifecycleSample {
    sample: usize,
    open: PhaseSample,
    snapshot: PhaseSample,
    stage: PhaseSample,
    commit: PhaseSample,
    publish: PhaseSample,
    drop: PhaseSample,
    source_reads: ReadObservation,
    sink: SinkRecord,
}

#[derive(Clone, Debug, Serialize)]
struct CaseReport {
    count: usize,
    corpus: CorpusRecord,
    total_samples: Option<Vec<TotalSample>>,
    phase_samples: Option<Vec<PhaseLifecycleSample>>,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    version: u32,
    binary: BinaryRecord,
    config: ConfigRecord,
    cases: Vec<CaseReport>,
}

#[derive(Debug)]
struct TimedPhase {
    elapsed_ns: u64,
    source_reads: ReadObservation,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

fn hex_digest(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(&mut output, "{byte:02x}");
    }
    output
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut digest = Sha256::new();
    digest.update(bytes);
    hex_digest(&digest.finalize())
}

fn source_main_xml(count: usize) -> Vec<u8> {
    let mut xml = String::with_capacity(count.saturating_mul(56).saturating_add(128));
    let _ = write!(
        &mut xml,
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{namespace}"><w:body>"#,
        namespace = WORD_NAMESPACE,
    );
    for index in 0..count {
        let _ = write!(
            &mut xml,
            "<w:p><w:r><w:t>paragraph-{index:06}</w:t></w:r></w:p>"
        );
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn independent_tail_xml(source: &[u8]) -> BenchResult<Vec<u8>> {
    let fragment = b"<w:p><w:r><w:t>paragraph-000000</w:t></w:r></w:p>";
    let body_close = b"</w:body>";
    let body_end = source
        .windows(body_close.len())
        .position(|window| window == body_close)
        .ok_or("generated DOCX source is missing body close")?;
    let fragment_start = source
        .windows(fragment.len())
        .position(|window| window == fragment)
        .ok_or("generated DOCX source is missing first paragraph")?;
    if fragment_start >= body_end {
        return Err("generated DOCX first paragraph is outside body".into());
    }
    let mut output = Vec::with_capacity(source.len() + fragment.len());
    output.extend_from_slice(&source[..body_end]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&source[body_end..]);
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
        OPAQUE_MEDIA_TYPE.to_owned(),
        opaque_payload(opaque_variant),
    )))?;
    package.relate_to(MAIN_PATH, rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn limits_for(count: usize, source_xml_bytes: usize) -> BenchResult<Limits> {
    let max_xml_bytes = source_xml_bytes
        .checked_add(1024 * 1024)
        .ok_or("DOCX XML limit overflows usize")?;
    let max_output_bytes = max_xml_bytes
        .checked_add(1024 * 1024)
        .ok_or("DOCX output limit overflows usize")?;
    let max_events = count
        .checked_add(1)
        .and_then(|value| value.checked_mul(8))
        .and_then(|value| value.checked_add(128))
        .ok_or("DOCX event limit overflows usize")?;
    Ok(Limits::new(
        max_xml_bytes,
        count
            .checked_add(1)
            .ok_or("DOCX paragraph limit overflow")?,
        max_events,
        16,
        max_output_bytes,
        64 * 1024 * 1024,
    )?)
}

fn limits_record(limits: Limits) -> LimitsRecord {
    LimitsRecord {
        max_xml_bytes: limits.max_xml_bytes(),
        max_paragraphs: limits.max_paragraphs(),
        max_events: limits.max_events(),
        max_depth: limits.max_depth(),
        max_output_bytes: limits.max_output_bytes(),
        max_durable_bytes: limits.max_durable_bytes(),
    }
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

fn semantic_record(bytes: &[u8]) -> BenchResult<(SemanticRecord, Vec<String>)> {
    let package = OwnedDocxPackage::from_reader(Cursor::new(bytes.to_vec()))?;
    let document = package.document_snapshot()?;
    let paragraphs = document.paragraphs();
    let mut texts = Vec::with_capacity(paragraphs.len());
    let mut order = Sha256::new();
    let mut text = Sha256::new();
    order.update(b"litchi-docx-tail-order-v1\0");
    text.update(b"litchi-docx-tail-text-v1\0");
    order.update(u64::try_from(paragraphs.len())?.to_le_bytes());
    text.update(u64::try_from(paragraphs.len())?.to_le_bytes());
    let mut text_bytes = 0usize;
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let value = paragraph.text()?;
        text_bytes = text_bytes
            .checked_add(value.len())
            .ok_or("DOCX semantic text byte count overflow")?;
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

fn find_member<'a>(members: &'a [MemberIdentity], path: &str) -> BenchResult<&'a MemberIdentity> {
    members
        .iter()
        .find(|member| member.path == path)
        .ok_or_else(|| format!("DOCX member {path:?} is missing").into())
}

fn build_candidate(
    source_archive_bytes: &[u8],
    source_xml_bytes: &[u8],
    count: usize,
    limits: Limits,
) -> BenchResult<(Vec<u8>, Vec<u8>, PatchOracle)> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source_archive_bytes.to_vec()));
    let package = source_backed::Package::from_read_at(source)?;
    let snapshot = package.plain_paragraph_copy_snapshot_with_limits(limits)?;
    if snapshot.paragraph_count() != count || snapshot.xml_bytes() != source_xml_bytes {
        return Err("DOCX source snapshot differs from generated main XML".into());
    }
    let mut edit = snapshot.edit();
    edit.copy_plain_paragraph(Position::new(0), Position::new(count))?;
    let commit = edit.commit();
    let candidate_xml = commit.projected().xml_bytes().to_vec();
    let effect = commit.effect_report();
    let wire = commit.patch().to_bytes()?;
    let durable = Patch::from_bytes(&wire)?;
    let replayed = durable.apply(&snapshot)?;
    let replay_verified = replayed.xml_bytes() == candidate_xml;
    let inverse_verified = durable.inverse().apply(&replayed)?.xml_bytes() == source_xml_bytes;
    let canonical = durable.to_bytes()? == wire;

    let stale_archive = source_archive(source_xml_bytes, 1)?;
    let stale_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(stale_archive));
    let stale_package = source_backed::Package::from_read_at(stale_source)?;
    let stale_snapshot = stale_package.plain_paragraph_copy_snapshot_with_limits(limits)?;
    let stale_source_refusal_verified =
        matches!(durable.apply(&stale_snapshot), Err(CopyError::StaleSource));

    let mut candidate_archive = Vec::new();
    let publication =
        package.publish_plain_paragraph_copy_to_stream(&mut candidate_archive, &commit)?;
    let candidate_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(candidate_archive.clone()));
    let candidate_package = source_backed::Package::from_read_at(candidate_source)?;
    let mut restored = Vec::new();
    candidate_package
        .publish_plain_paragraph_copy_patch_to_stream(&mut restored, publication.inverse_patch())?;
    let publication_inverse_verified = restored == source_archive_bytes;
    if !replay_verified
        || !inverse_verified
        || !canonical
        || !stale_source_refusal_verified
        || !publication_inverse_verified
    {
        return Err("DOCX plain paragraph patch oracle failed".into());
    }
    Ok((
        candidate_archive,
        candidate_xml,
        PatchOracle {
            copied_source_position: 0,
            copied_before_position: count,
            copied_paragraphs: effect.copied_paragraphs,
            copied_bytes: effect.copied_bytes,
            durable_bytes: wire.len(),
            durable_canonical_verified: canonical,
            replay_verified,
            inverse_verified,
            stale_source_refusal_verified,
            publication_inverse_verified,
        },
    ))
}

fn build_corpus(count: usize) -> BenchResult<Corpus> {
    if !DEFAULT_COUNTS.contains(&count) {
        return Err(format!("unsupported DOCX paragraph count {count}").into());
    }
    let source_main_xml = source_main_xml(count);
    let limits = limits_for(count, source_main_xml.len())?;
    let source_archive_vec = source_archive(&source_main_xml, 0)?;
    let source_archive_guard = source_archive_vec.clone();
    let (candidate_archive_vec, candidate_main_xml, patch) =
        build_candidate(&source_archive_vec, &source_main_xml, count, limits)?;
    let source_unchanged_verified = source_archive_vec == source_archive_guard;
    let (source_semantic, source_texts) = semantic_record(&source_archive_vec)?;
    let (candidate_semantic, candidate_texts) = semantic_record(&candidate_archive_vec)?;
    let expected_candidate_xml = independent_tail_xml(&source_main_xml)?;
    let source_members = member_identities(&source_archive_vec)?;
    let candidate_members = member_identities(&candidate_archive_vec)?;
    let source_main = find_member(&source_members, MAIN_PATH)?;
    let candidate_main = find_member(&candidate_members, MAIN_PATH)?;
    let source_opaque = find_member(&source_members, OPAQUE_PATH)?;
    let candidate_opaque = find_member(&candidate_members, OPAQUE_PATH)?;
    let source_archive_reader = ArchiveReader::new(&source_archive_vec)?;
    let candidate_archive_reader = ArchiveReader::new(&candidate_archive_vec)?;
    let source_main_xml_archive = source_archive_reader.read(MAIN_PATH)?;
    let candidate_main_xml_archive = candidate_archive_reader.read(MAIN_PATH)?;
    let source_main_xml_archive_verified = source_main_xml_archive == source_main_xml;
    let candidate_main_xml_archive_verified = candidate_main_xml_archive == candidate_main_xml;
    let candidate_main_xml_oracle_verified = candidate_main_xml == expected_candidate_xml;
    let untouched_members_verified = source_members
        .iter()
        .filter(|member| member.path != MAIN_PATH)
        .all(|member| {
            candidate_members
                .iter()
                .find(|candidate| candidate.path == member.path)
                == Some(member)
        });
    let opaque_member_exact_verified = source_opaque == candidate_opaque;
    let physical_order_verified = source_members
        .iter()
        .map(|member| member.path.as_str())
        .eq(candidate_members.iter().map(|member| member.path.as_str()));
    let source_semantic_reopen_verified = source_semantic.paragraph_count == count
        && source_texts
            .iter()
            .enumerate()
            .all(|(index, text)| text == &format!("paragraph-{index:06}"));
    let candidate_semantic_reopen_verified = candidate_semantic.paragraph_count == count + 1
        && candidate_texts.get(..count) == Some(source_texts.as_slice())
        && candidate_texts.last() == source_texts.first();
    let tail_copy_exactly_one_verified = candidate_texts.len() == source_texts.len() + 1
        && candidate_texts.get(..count) == Some(source_texts.as_slice())
        && candidate_texts.last() == source_texts.first();
    if source_members.len() != 4
        || candidate_members.len() != 4
        || source_main.decoded_bytes != source_main_xml.len()
        || candidate_main.decoded_bytes != candidate_main_xml.len()
        || !source_main_xml_archive_verified
        || !candidate_main_xml_archive_verified
        || !candidate_main_xml_oracle_verified
        || !source_semantic_reopen_verified
        || !candidate_semantic_reopen_verified
        || !tail_copy_exactly_one_verified
        || !untouched_members_verified
        || !opaque_member_exact_verified
        || !physical_order_verified
        || !source_unchanged_verified
    {
        return Err("DOCX plain paragraph corpus oracle failed".into());
    }
    let source_archive = Arc::<[u8]>::from(source_archive_vec.clone());
    let candidate_archive = Arc::<[u8]>::from(candidate_archive_vec.clone());
    let record = CorpusRecord {
        generator: GENERATOR,
        format: "DOCX/OOXML/OPC/ZIP",
        count,
        opaque_path: OPAQUE_PATH,
        opaque_bytes: OPAQUE_BYTES,
        source_archive_bytes: source_archive_vec.len(),
        source_archive_sha256: sha256_hex(&source_archive_vec),
        candidate_archive_bytes: candidate_archive_vec.len(),
        candidate_archive_sha256: sha256_hex(&candidate_archive_vec),
        source_main_xml_bytes: source_main_xml.len(),
        source_main_xml_sha256: sha256_hex(&source_main_xml),
        candidate_main_xml_bytes: candidate_main_xml.len(),
        candidate_main_xml_sha256: sha256_hex(&candidate_main_xml),
        source_main_xml_archive_verified,
        candidate_main_xml_archive_verified,
        candidate_main_xml_oracle_verified,
        source_member_count: source_members.len(),
        candidate_member_count: candidate_members.len(),
        source_members: source_members.clone(),
        candidate_members: candidate_members.clone(),
        source_semantic: source_semantic.clone(),
        candidate_semantic: candidate_semantic.clone(),
        source_unchanged_verified,
        source_semantic_reopen_verified,
        candidate_semantic_reopen_verified,
        tail_copy_exactly_one_verified,
        untouched_members_verified,
        opaque_member_exact_verified,
        physical_order_verified,
        limits: limits_record(limits),
        patch: patch.clone(),
    };
    Ok(Corpus {
        count,
        source_archive,
        candidate_archive,
        limits,
        record,
    })
}

fn elapsed_ns(start: Instant) -> BenchResult<u64> {
    u64::try_from(start.elapsed().as_nanos())
        .map_err(|_| "elapsed duration does not fit u64 nanoseconds".into())
}

fn phase_result(
    start: Instant,
    region: allocation_metrics::Region,
    before: Option<process_metrics::Snapshot>,
    source_before: ReadObservation,
    source_after: ReadObservation,
) -> BenchResult<TimedPhase> {
    let elapsed_ns = elapsed_ns(start)?;
    let allocation = region.finish();
    let process = before
        .zip(process_metrics::Snapshot::read().ok())
        .map(|(before, after)| after.delta(before));
    Ok(TimedPhase {
        elapsed_ns,
        source_reads: source_after.delta(source_before),
        allocation,
        process,
    })
}

fn run_total_iteration(
    corpus: &Corpus,
) -> BenchResult<(
    u64,
    SinkObservation,
    ReadObservation,
    Option<allocation_metrics::Sample>,
    Option<process_metrics::Delta>,
)> {
    let process_before = process_metrics::Snapshot::read().ok();
    let region = allocation_metrics::begin();
    let start = Instant::now();
    let (sink, reads) = {
        let source = Arc::new(MeasureSource::new(Arc::clone(&corpus.source_archive)));
        let package = source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
        let snapshot = package.plain_paragraph_copy_snapshot_with_limits(corpus.limits)?;
        let mut edit = snapshot.edit();
        edit.copy_plain_paragraph(Position::new(0), Position::new(corpus.count))?;
        let commit = edit.commit();
        let mut output = HashingSink::new();
        let _publication = package.publish_plain_paragraph_copy_to_stream(&mut output, &commit)?;
        let sink = output.finish();
        let reads = source.reads();
        (sink, reads)
    };
    let elapsed = elapsed_ns(start)?;
    let allocation = region.finish();
    let process = process_before
        .zip(process_metrics::Snapshot::read().ok())
        .map(|(before, after)| after.delta(before));
    Ok((elapsed, sink, reads, allocation, process))
}

struct PhaseState {
    source: Option<Arc<MeasureSource>>,
    package: Option<source_backed::Package>,
    snapshot: Option<Snapshot>,
    edit: Option<Edit>,
    commit: Option<Commit>,
    publication: Option<Publication>,
    sink: Option<SinkObservation>,
    reads: ReadObservation,
}

impl PhaseState {
    fn new() -> Self {
        Self {
            source: None,
            package: None,
            snapshot: None,
            edit: None,
            commit: None,
            publication: None,
            sink: None,
            reads: ReadObservation::default(),
        }
    }
}

fn run_phase<F>(source_before: ReadObservation, action: F) -> BenchResult<TimedPhase>
where
    F: FnOnce() -> BenchResult<ReadObservation>,
{
    let process_before = process_metrics::Snapshot::read().ok();
    let region = allocation_metrics::begin();
    let start = Instant::now();
    let source_after = action()?;
    phase_result(start, region, process_before, source_before, source_after)
}

fn run_phase_iteration(corpus: &Corpus) -> BenchResult<(PhaseLifecycleSample, PhaseState)> {
    let mut state = PhaseState::new();
    let open = run_phase(ReadObservation::default(), || {
        let source = Arc::new(MeasureSource::new(Arc::clone(&corpus.source_archive)));
        let package = source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
        state.source = Some(source);
        state.package = Some(package);
        Ok(state
            .source
            .as_ref()
            .map_or(ReadObservation::default(), |source| source.reads()))
    })?;
    let snapshot = run_phase(
        state
            .source
            .as_ref()
            .map_or(ReadObservation::default(), |source| source.reads()),
        || {
            let snapshot = state
                .package
                .as_ref()
                .ok_or("phase snapshot has no package")?
                .plain_paragraph_copy_snapshot_with_limits(corpus.limits)?;
            state.snapshot = Some(snapshot);
            Ok(state
                .source
                .as_ref()
                .map_or(ReadObservation::default(), |source| source.reads()))
        },
    )?;
    let stage = run_phase(
        state
            .source
            .as_ref()
            .map_or(ReadObservation::default(), |source| source.reads()),
        || {
            let mut edit = state
                .snapshot
                .as_ref()
                .ok_or("phase stage has no snapshot")?
                .edit();
            edit.copy_plain_paragraph(Position::new(0), Position::new(corpus.count))?;
            state.edit = Some(edit);
            Ok(state
                .source
                .as_ref()
                .map_or(ReadObservation::default(), |source| source.reads()))
        },
    )?;
    let commit = run_phase(
        state
            .source
            .as_ref()
            .map_or(ReadObservation::default(), |source| source.reads()),
        || {
            let edit = state.edit.take().ok_or("phase commit has no edit")?;
            state.commit = Some(edit.commit());
            Ok(state
                .source
                .as_ref()
                .map_or(ReadObservation::default(), |source| source.reads()))
        },
    )?;
    let publish = run_phase(
        state
            .source
            .as_ref()
            .map_or(ReadObservation::default(), |source| source.reads()),
        || {
            let package = state.package.take().ok_or("phase publish has no package")?;
            let commit = state.commit.as_ref().ok_or("phase publish has no commit")?;
            let mut output = HashingSink::new();
            let publication =
                package.publish_plain_paragraph_copy_to_stream(&mut output, commit)?;
            state.sink = Some(output.finish());
            state.publication = Some(publication);
            state.reads = state
                .source
                .as_ref()
                .map_or(ReadObservation::default(), |source| source.reads());
            Ok(state.reads)
        },
    )?;
    let drop_phase = run_phase(state.reads, || {
        let reads = state.reads;
        state.publication.take();
        state.commit.take();
        state.edit.take();
        state.snapshot.take();
        state.package.take();
        state.source.take();
        Ok(reads)
    })?;
    let sink = state.sink.take().ok_or("phase lifecycle lost sink")?;
    Ok((
        PhaseLifecycleSample {
            sample: 0,
            open: PhaseSample {
                elapsed_ns: open.elapsed_ns,
                source_reads: open.source_reads,
                allocation: open.allocation,
                process: open.process,
            },
            snapshot: PhaseSample {
                elapsed_ns: snapshot.elapsed_ns,
                source_reads: snapshot.source_reads,
                allocation: snapshot.allocation,
                process: snapshot.process,
            },
            stage: PhaseSample {
                elapsed_ns: stage.elapsed_ns,
                source_reads: stage.source_reads,
                allocation: stage.allocation,
                process: stage.process,
            },
            commit: PhaseSample {
                elapsed_ns: commit.elapsed_ns,
                source_reads: commit.source_reads,
                allocation: commit.allocation,
                process: commit.process,
            },
            publish: PhaseSample {
                elapsed_ns: publish.elapsed_ns,
                source_reads: publish.source_reads,
                allocation: publish.allocation,
                process: publish.process,
            },
            drop: PhaseSample {
                elapsed_ns: drop_phase.elapsed_ns,
                source_reads: drop_phase.source_reads,
                allocation: drop_phase.allocation,
                process: drop_phase.process,
            },
            source_reads: state.reads,
            sink: sink.record(),
        },
        state,
    ))
}

fn check_sink(corpus: &Corpus, sink: SinkObservation) -> BenchResult<()> {
    if sink.accepted_bytes != u64::try_from(corpus.candidate_archive.len())?
        || sink.digest.as_slice() != Sha256::digest(corpus.candidate_archive.as_ref()).as_slice()
    {
        return Err("DOCX runtime output does not match candidate archive oracle".into());
    }
    Ok(())
}

fn check_sink_record(corpus: &Corpus, sink: &SinkRecord) -> BenchResult<()> {
    if sink.accepted_bytes != u64::try_from(corpus.candidate_archive.len())?
        || sink.sha256 != corpus.record.candidate_archive_sha256
    {
        return Err("DOCX phase output does not match candidate archive oracle".into());
    }
    Ok(())
}

fn run_case(corpus: &Corpus, config: &Config) -> BenchResult<CaseReport> {
    match config.mode {
        Mode::Total => {
            let mut samples = Vec::with_capacity(config.samples);
            for iteration in 0..config.warmups.saturating_add(config.samples) {
                let (elapsed, sink, reads, allocation, process) = run_total_iteration(corpus)?;
                check_sink(corpus, sink)?;
                if iteration >= config.warmups {
                    samples.push(TotalSample {
                        sample: iteration - config.warmups,
                        elapsed_ns: elapsed,
                        source_reads: reads,
                        sink: sink.record(),
                        allocation,
                        process,
                    });
                }
            }
            Ok(CaseReport {
                count: corpus.count,
                corpus: corpus.record.clone(),
                total_samples: Some(samples),
                phase_samples: None,
            })
        },
        Mode::Phases => {
            let mut samples = Vec::with_capacity(config.samples);
            for iteration in 0..config.warmups.saturating_add(config.samples) {
                let (mut sample, _state) = run_phase_iteration(corpus)?;
                check_sink_record(corpus, &sample.sink)?;
                if iteration >= config.warmups {
                    sample.sample = iteration - config.warmups;
                    samples.push(sample);
                }
            }
            Ok(CaseReport {
                count: corpus.count,
                corpus: corpus.record.clone(),
                total_samples: None,
                phase_samples: Some(samples),
            })
        },
    }
}

fn parse_args<I>(args: I) -> BenchResult<Config>
where
    I: IntoIterator<Item = OsString>,
{
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut samples = DEFAULT_SAMPLES;
    let mut warmups = DEFAULT_WARMUPS;
    let mut mode = Mode::Total;
    let mut json_path = None;
    let mut values = args.into_iter();
    while let Some(argument) = values.next() {
        let flag = argument.to_string_lossy();
        match flag.as_ref() {
            "--counts" => counts = parse_counts(&next_value(&mut values, "--counts")?)?,
            "--samples" => {
                samples = parse_positive(&next_value(&mut values, "--samples")?, "--samples")?
            },
            "--warmups" => {
                warmups = parse_positive(&next_value(&mut values, "--warmups")?, "--warmups")?
            },
            "--mode" => {
                mode = Mode::parse(
                    next_value(&mut values, "--mode")?
                        .to_str()
                        .ok_or("--mode must be UTF-8")?,
                )?
            },
            "--json" => json_path = Some(PathBuf::from(next_value(&mut values, "--json")?)),
            "--help" | "-h" => return Err(usage().into()),
            unknown => return Err(format!("unknown argument {unknown}; {}", usage()).into()),
        }
    }
    if counts.is_empty() || samples == 0 || warmups == 0 {
        return Err("counts, samples, and warmups must be non-zero".into());
    }
    Ok(Config {
        counts,
        samples,
        warmups,
        mode,
        json_path,
    })
}

fn next_value(values: &mut impl Iterator<Item = OsString>, flag: &str) -> BenchResult<OsString> {
    values
        .next()
        .ok_or_else(|| format!("missing value for {flag}").into())
}

fn parse_counts(value: &OsString) -> BenchResult<Vec<usize>> {
    let text = value.to_str().ok_or("--counts must be UTF-8")?;
    let mut counts = Vec::new();
    for part in text.split(',') {
        let count = part
            .parse::<usize>()
            .map_err(|error| format!("invalid count {part:?}: {error}"))?;
        if !DEFAULT_COUNTS.contains(&count) {
            return Err(format!(
                "unsupported DOCX paragraph count {count}; expected 64, 8192, or 131072"
            )
            .into());
        }
        if !counts.contains(&count) {
            counts.push(count);
        }
    }
    Ok(counts)
}

fn parse_positive(value: &OsString, flag: &str) -> BenchResult<usize> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let parsed = text
        .parse::<usize>()
        .map_err(|error| format!("invalid {flag} {text:?}: {error}"))?;
    if parsed == 0 {
        return Err(format!("{flag} must be positive").into());
    }
    Ok(parsed)
}

fn usage() -> &'static str {
    "usage: docx_plain_paragraph_tail_append [--counts 64,8192,131072] [--samples N] [--warmups N] [--mode total|phases] [--json PATH]"
}

pub fn run_from_args<I>(args: I) -> BenchResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let config = parse_args(args)?;
    let mut cases = Vec::with_capacity(config.counts.len());
    for &count in &config.counts {
        let corpus = build_corpus(count)?;
        cases.push(run_case(&corpus, &config)?);
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
            counts: config.counts.clone(),
            samples: config.samples,
            warmups: config.warmups,
            mode: config.mode,
            lifecycle_phases: ["open", "snapshot", "stage", "commit", "publish", "drop"],
            sink: "non_seek_hashing_scalar_sha256_no_archive_retention_shortwrite_16k",
            source: "caller_owned_arc_positional_read_at_scalar_counters_requested_returned_fixed_histogram",
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
    fn generated_corpus_has_distinct_tail_and_opaque_identity() {
        let corpus = build_corpus(64).expect("small DOCX corpus");
        let (source_semantic, source_texts) = semantic_record(&corpus.source_archive).unwrap();
        let (candidate_semantic, candidate_texts) =
            semantic_record(&corpus.candidate_archive).unwrap();
        assert_eq!(source_semantic.paragraph_count, 64);
        assert_eq!(candidate_semantic.paragraph_count, 65);
        assert_eq!(source_texts.first(), candidate_texts.last());
        assert_eq!(candidate_texts.get(..64), Some(source_texts.as_slice()));
        assert!(corpus.record.opaque_member_exact_verified);
        assert!(corpus.record.untouched_members_verified);
        assert_eq!(corpus.record.patch.copied_source_position, 0);
        assert_eq!(corpus.record.patch.copied_before_position, 64);
        assert_eq!(corpus.record.patch.copied_paragraphs, 1);
    }

    #[test]
    fn large_policy_is_explicitly_above_default_paragraph_ceiling() {
        let limits = limits_for(131_072, 1_000).expect("finite limits");
        assert!(limits.max_paragraphs() > Limits::default().max_paragraphs());
        assert!(limits.max_events() > 131_072);
        assert!(Limits::new(0, 1, 1, 1, 1, 1).is_err());
    }

    #[test]
    fn copy_errors_remain_typed_and_atomic() {
        let xml = source_main_xml(2);
        let archive = source_archive(&xml, 0).expect("archive");
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(archive));
        let package = source_backed::Package::from_read_at(source).expect("package");
        let snapshot = package
            .plain_paragraph_copy_snapshot_with_limits(limits_for(2, xml.len()).unwrap())
            .expect("snapshot");
        let mut edit = snapshot.edit();
        edit.copy_plain_paragraph(Position::new(0), Position::new(2))
            .expect("first operation");
        assert!(matches!(
            edit.copy_plain_paragraph(Position::new(1), Position::new(2)),
            Err(CopyError::Limit {
                resource: "operations",
                max: 1,
                actual: 2
            })
        ));
        assert!(matches!(
            snapshot
                .edit()
                .copy_plain_paragraph(Position::new(2), Position::new(2)),
            Err(CopyError::OutOfBounds { .. })
        ));
    }

    #[test]
    fn section_properties_are_a_typed_refusal() {
        let xml = format!(
            r#"<?xml version="1.0"?><w:document xmlns:w="{WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#
        )
        .into_bytes();
        let archive = source_archive(&xml, 0).expect("archive");
        let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(archive));
        let package = source_backed::Package::from_read_at(source).expect("package");
        assert!(matches!(
            package.plain_paragraph_copy_snapshot_with_limits(limits_for(1, xml.len()).unwrap()),
            Err(CopyError::Refused(_))
        ));
    }

    #[test]
    fn source_adapter_counts_logical_reads_without_ranges() {
        let source = MeasureSource::new(Arc::<[u8]>::from(vec![1, 2, 3, 4]));
        let mut buffer = [0u8; 3];
        assert_eq!(source.read_at(1, &mut buffer).unwrap(), 3);
        assert_eq!(source.read_at(4, &mut buffer).unwrap(), 0);
        let reads = source.reads();
        assert_eq!(reads.calls, 2);
        assert_eq!(reads.requested_bytes, 6);
        assert_eq!(reads.returned_bytes, 3);
        assert_eq!(reads.request_histogram.bytes_1_to_512, 2);
    }

    #[test]
    fn sink_keeps_only_scalar_counters_and_fixed_histogram() {
        let mut sink = HashingSink::new();
        sink.write_all(&[1u8; 3]).unwrap();
        sink.write_all(&[2u8; 600]).unwrap();
        let observation = sink.finish();
        assert_eq!(observation.accepted_bytes, 603);
        assert_eq!(observation.write_calls, 2);
        assert_eq!(observation.histogram.bytes_1_to_512, 1);
        assert_eq!(observation.histogram.bytes_513_to_4096, 1);
    }

    #[test]
    fn cli_accepts_only_frozen_shapes_and_modes() {
        let config = parse_args([
            OsString::from("--counts"),
            OsString::from("64,131072,64"),
            OsString::from("--samples"),
            OsString::from("2"),
            OsString::from("--warmups"),
            OsString::from("1"),
            OsString::from("--mode"),
            OsString::from("phases"),
        ])
        .unwrap();
        assert_eq!(config.counts, [64, 131_072]);
        assert_eq!(config.mode, Mode::Phases);
        assert!(parse_args([OsString::from("--counts"), OsString::from("65")]).is_err());
        assert!(parse_args([OsString::from("--mode"), OsString::from("total")]).is_ok());
    }
}
