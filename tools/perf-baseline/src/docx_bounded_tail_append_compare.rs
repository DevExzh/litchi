//! Same-source DOCX tail-append comparison.
//!
//! This binary is intentionally separate from
//! [`super::docx_plain_paragraph_tail_append`].  The older binary is a frozen
//! materialized paragraph-copy schema used by the 0481 evidence.  This runner
//! adds a second route, the bounded caller-text tail append, without changing
//! that schema or its command-line behavior.
//!
//! The control source's first paragraph is the explicit caller text.  The
//! materialized route copies that paragraph to the tail; the bounded route is
//! given the same text.  Therefore both routes publish the same logical
//! paragraph sequence while the benchmark still records which route authored
//! new text.  Corpus construction and all independent archive/semantic
//! oracles are outside the measured lifecycle.

use std::{
    collections::BTreeMap,
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
use litchi_docx::source_backed::paragraph_copy::Limits as CopyLimits;
// The DOCX owner follows the source-backed ODF append shape: an edit retains
// only caller text and the indexed source, `plan` proves the source/candidate,
// and publication writes the preserved package to a sequential sink.  Keep
// this import narrow so the benchmark does not reach into OPC internals.
use litchi_docx::Package as OwnedDocxPackage;
use litchi_docx::source_backed::tail_append::{
    CandidateProof as TailAppendCandidateProof, Limits as TailAppendLimits,
    SourceProof as TailAppendSourceProof,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use serde::Serialize;
use sha2::{Digest as _, Sha256};
use soapberry_zip::ZipArchive;
use soapberry_zip::office::ArchiveReader;

use super::{allocation_metrics, process_metrics};

const SCHEMA: &str = "docx-bounded-tail-append-comparison-v1";
const DEFAULT_COUNTS: [usize; 3] = [64, 8_192, 131_072];
const DEFAULT_SAMPLES: usize = 30;
const DEFAULT_WARMUPS: usize = 3;
const OPAQUE_PATH: &str = "word/perf-opaque.bin";
const OPAQUE_BYTES: usize = 32 * 1024;
const MAIN_PATH: &str = "word/document.xml";
const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OPAQUE_MEDIA_TYPE: &str = "application/octet-stream";
const GENERATOR: &str = "litchi-docx-bounded-tail-append-comparison-v1";
const HASH_SINK_MAX_WRITE: usize = 16 * 1024;

type BenchResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum Route {
    MaterializedParagraphCopy,
    BoundedPlainTextTailAppend,
}

impl Route {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "materialized" | "materialized_paragraph_copy" => Ok(Self::MaterializedParagraphCopy),
            "bounded" | "bounded_plain_text_tail_append" => Ok(Self::BoundedPlainTextTailAppend),
            _ => Err(format!("invalid --route {value:?}; expected materialized or bounded").into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::MaterializedParagraphCopy => "materialized_paragraph_copy",
            Self::BoundedPlainTextTailAppend => "bounded_plain_text_tail_append",
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    counts: Vec<usize>,
    samples: usize,
    warmups: usize,
    route: Route,
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
    route: Route,
    lifecycle: &'static str,
    sink: &'static str,
    source: &'static str,
    text_authoring: &'static str,
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
struct ScalarProofRecord {
    source_version_id: u64,
    source_version_revision: u64,
    source_len: u64,
    source_sha256: String,
    insertion_offset: u64,
    source_paragraph_count: u64,
    source_event_count: u64,
    source_max_depth: u64,
    source_strict_namespace: bool,
    source_sect_pr_len: u64,
    source_sect_pr_sha256: String,
    candidate_len: u64,
    candidate_sha256: String,
    candidate_paragraph_count: u64,
    candidate_event_count: u64,
    candidate_max_depth: u64,
    generated_offset: u64,
    generated_once: bool,
    candidate_sect_pr_len: u64,
    candidate_sect_pr_sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct RouteCorpusRecord {
    route: Route,
    output_archive_bytes: usize,
    output_archive_sha256: String,
    output_main_xml_bytes: usize,
    output_main_xml_sha256: String,
    output_main_xml_expected_semantic_verified: bool,
    output_main_xml_expected_raw_verified: bool,
    scalar_proof: Option<ScalarProofRecord>,
    output_semantic: SemanticRecord,
    output_members: Vec<MemberIdentity>,
    output_member_count: usize,
    output_untouched_members_verified: bool,
    output_untouched_raw_members_verified: bool,
    output_opaque_member_exact_verified: bool,
    output_physical_order_verified: bool,
    output_main_compressed_equal_source: bool,
    route_output_replay_verified: bool,
    route_inverse_verified: bool,
    stale_source_refusal_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct CorpusRecord {
    generator: &'static str,
    format: &'static str,
    count: usize,
    append_text: String,
    append_text_sha256: String,
    source_archive_bytes: usize,
    source_archive_sha256: String,
    source_main_xml_bytes: usize,
    source_main_xml_sha256: String,
    source_semantic: SemanticRecord,
    source_members: Vec<MemberIdentity>,
    source_member_count: usize,
    source_unchanged_verified: bool,
    source_main_xml_archive_verified: bool,
    source_opaque_member_exact_verified: bool,
    route_semantics_equal_verified: bool,
    route_untouched_members_equal_verified: bool,
    route_outputs_physical_equal: bool,
    route_untouched_raw_members_equal_verified: bool,
    bounded_noop_verified: bool,
    materialized: RouteCorpusRecord,
    bounded: RouteCorpusRecord,
    limits: LimitsRecord,
}

#[derive(Clone, Debug)]
struct Corpus {
    count: usize,
    append_text: String,
    source_archive: Arc<[u8]>,
    materialized_archive: Arc<[u8]>,
    bounded_archive: Arc<[u8]>,
    copy_limits: CopyLimits,
    bounded_limits: TailAppendLimits,
    record: CorpusRecord,
}

#[derive(Clone, Debug, Serialize)]
struct LimitsRecord {
    copy_max_xml_bytes: usize,
    copy_max_paragraphs: usize,
    copy_max_events: usize,
    copy_max_depth: usize,
    copy_max_output_bytes: usize,
    bounded: String,
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

#[derive(Debug)]
struct HashingSink {
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
        self.accepted_bytes =
            self.accepted_bytes
                .checked_add(u64::try_from(accepted).map_err(|_| {
                    io::Error::other("DOCX comparison sink byte count overflows u64")
                })?)
                .ok_or_else(|| io::Error::other("DOCX comparison sink byte count overflows u64"))?;
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("DOCX comparison sink call count overflows u64"))?;
        self.largest_write = self.largest_write.max(
            u64::try_from(accepted)
                .map_err(|_| io::Error::other("DOCX comparison sink write size overflows u64"))?,
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

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
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
struct Sample {
    sample: usize,
    route: Route,
    elapsed_ns: u64,
    source_reads: ReadObservation,
    sink: SinkRecord,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Clone, Debug, Serialize)]
struct RouteReport {
    route: Route,
    corpus: RouteCorpusRecord,
    samples: Vec<Sample>,
}

#[derive(Clone, Debug, Serialize)]
struct CaseReport {
    count: usize,
    corpus: CorpusRecord,
    routes: Vec<RouteReport>,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    version: u32,
    binary: BinaryRecord,
    config: ConfigRecord,
    cases: Vec<CaseReport>,
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

fn scalar_proof_record(
    source: TailAppendSourceProof,
    candidate: TailAppendCandidateProof,
) -> ScalarProofRecord {
    ScalarProofRecord {
        source_version_id: source.source_version.id(),
        source_version_revision: source.source_version.revision(),
        source_len: source.source_len,
        source_sha256: hex_digest(&source.source_sha256),
        insertion_offset: source.insertion_offset,
        source_paragraph_count: source.paragraph_count,
        source_event_count: source.event_count,
        source_max_depth: source.max_depth,
        source_strict_namespace: source.strict_namespace,
        source_sect_pr_len: source.sect_pr_len,
        source_sect_pr_sha256: hex_digest(&source.sect_pr_sha256),
        candidate_len: candidate.candidate_len,
        candidate_sha256: hex_digest(&candidate.candidate_sha256),
        candidate_paragraph_count: candidate.paragraph_count,
        candidate_event_count: candidate.event_count,
        candidate_max_depth: candidate.max_depth,
        generated_offset: candidate.generated_offset,
        generated_once: candidate.generated_once,
        candidate_sect_pr_len: candidate.sect_pr_len,
        candidate_sect_pr_sha256: hex_digest(&candidate.sect_pr_sha256),
    }
}

fn verify_scalar_proof(
    proof: &ScalarProofRecord,
    count: usize,
    source_xml: &[u8],
    candidate_xml: &[u8],
) -> BenchResult<()> {
    let body_end = source_xml
        .windows(b"</w:body>".len())
        .position(|window| window == b"</w:body>")
        .ok_or("generated DOCX source is missing body close")?;
    let empty_hash = sha256_hex(&[]);
    if proof.source_len != u64::try_from(source_xml.len())?
        || proof.source_sha256 != sha256_hex(source_xml)
        || proof.insertion_offset != u64::try_from(body_end)?
        || proof.source_paragraph_count != u64::try_from(count)?
        || proof.source_strict_namespace
        || proof.source_sect_pr_len != 0
        || proof.source_sect_pr_sha256 != empty_hash
        || proof.candidate_len != u64::try_from(candidate_xml.len())?
        || proof.candidate_sha256 != sha256_hex(candidate_xml)
        || proof.candidate_paragraph_count
            != u64::try_from(
                count
                    .checked_add(1)
                    .ok_or("proof paragraph count overflow")?,
            )?
        || proof.generated_offset != u64::try_from(body_end)?
        || !proof.generated_once
        || proof.candidate_sect_pr_len != 0
        || proof.candidate_sect_pr_sha256 != empty_hash
        || proof.source_event_count == 0
        || proof.candidate_event_count <= proof.source_event_count
        || proof.source_max_depth == 0
        || proof.candidate_max_depth < proof.source_max_depth
    {
        return Err("bounded scalar proof disagrees with independent XML oracle".into());
    }
    Ok(())
}

fn xml_escape(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
    output
}

fn append_text(count: usize) -> String {
    format!("tail-text-{count:06}-café <&> plain")
}

fn paragraph_xml(text: &str) -> String {
    format!("<w:p><w:r><w:t>{}</w:t></w:r></w:p>", xml_escape(text))
}

fn bounded_paragraph_xml(text: &str) -> String {
    format!(
        "<w:p xmlns:w=\"{WORD_NAMESPACE}\"><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        xml_escape(text)
    )
}

fn source_main_xml(count: usize) -> Vec<u8> {
    let mut xml = String::with_capacity(count.saturating_mul(56).saturating_add(256));
    let _ = write!(
        &mut xml,
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{WORD_NAMESPACE}"><w:body>"#
    );
    xml.push_str(&paragraph_xml(&append_text(count)));
    for index in 1..count {
        let _ = write!(
            &mut xml,
            "<w:p><w:r><w:t>paragraph-{index:06}</w:t></w:r></w:p>"
        );
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn independent_candidate_xml(source: &[u8], text: &str) -> BenchResult<Vec<u8>> {
    independent_candidate_xml_with_fragment(source, &paragraph_xml(text))
}

fn independent_bounded_candidate_xml(source: &[u8], text: &str) -> BenchResult<Vec<u8>> {
    independent_candidate_xml_with_fragment(source, &bounded_paragraph_xml(text))
}

fn independent_candidate_xml_with_fragment(source: &[u8], fragment: &str) -> BenchResult<Vec<u8>> {
    let close = b"</w:body>";
    let body_end = source
        .windows(close.len())
        .position(|window| window == close)
        .ok_or("generated DOCX source is missing body close")?;
    let mut output = Vec::with_capacity(source.len().saturating_add(fragment.len()));
    output.extend_from_slice(&source[..body_end]);
    output.extend_from_slice(fragment.as_bytes());
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

fn copy_limits_for(count: usize, source_xml_bytes: usize) -> BenchResult<CopyLimits> {
    let max_xml_bytes = source_xml_bytes
        .checked_add(1024 * 1024)
        .ok_or("DOCX copy XML limit overflows usize")?;
    let max_output_bytes = max_xml_bytes
        .checked_add(1024 * 1024)
        .ok_or("DOCX copy output limit overflows usize")?;
    let max_events = count
        .checked_add(1)
        .and_then(|value| value.checked_mul(8))
        .and_then(|value| value.checked_add(128))
        .ok_or("DOCX copy event limit overflows usize")?;
    Ok(CopyLimits::new(
        max_xml_bytes,
        count
            .checked_add(1)
            .ok_or("DOCX copy paragraph limit overflow")?,
        max_events,
        16,
        max_output_bytes,
        64 * 1024 * 1024,
    )?)
}

fn bounded_limits_for(count: usize, source_xml_bytes: usize) -> BenchResult<TailAppendLimits> {
    // These values are finite and scale only the source/candidate/document
    // ceilings; the authored paragraph remains a small fixed fragment.
    let source_xml = u64::try_from(source_xml_bytes)?;
    let max_source_xml = source_xml
        .checked_add(1024 * 1024)
        .ok_or("DOCX bounded source XML limit overflows u64")?;
    let max_candidate_xml = source_xml
        .checked_add(2 * 1024 * 1024)
        .ok_or("DOCX bounded candidate XML limit overflows u64")?;
    let paragraphs = u64::try_from(
        count
            .checked_add(1)
            .ok_or("DOCX bounded paragraph limit overflows usize")?,
    )?;
    let events = paragraphs
        .checked_mul(8)
        .and_then(|value| value.checked_add(128))
        .ok_or("DOCX bounded event limit overflows u64")?;
    let output = source_xml
        .checked_add(4 * 1024 * 1024)
        .ok_or("DOCX bounded output limit overflows u64")?;
    // The corpus has a deliberately tiny direct-story grammar (the largest
    // lexical token is the authored text node, under 4 KiB, and the deepest
    // admitted scope is document/body/p/r/t). Keep the audit envelope fixed
    // across source sizes so its workspace reservation does not scale with
    // paragraph count; the source, candidate, event, paragraph, and output
    // ceilings above are the only size-scaled limits.
    const MAX_TOKEN_BYTES: u64 = 4 * 1024;
    const MAX_DEPTH: u64 = 16;
    const MAX_WORKSPACE_BYTES: u64 = 2 * 1024 * 1024;
    Ok(TailAppendLimits::new(
        max_source_xml,
        64 * 1024,
        64 * 1024,
        max_candidate_xml,
        events,
        MAX_DEPTH,
        paragraphs,
        4 * 1024 * 1024,
        MAX_WORKSPACE_BYTES,
        output,
        MAX_TOKEN_BYTES,
    ))
}

fn limits_record(copy: CopyLimits, bounded: TailAppendLimits) -> LimitsRecord {
    LimitsRecord {
        copy_max_xml_bytes: copy.max_xml_bytes(),
        copy_max_paragraphs: copy.max_paragraphs(),
        copy_max_events: copy.max_events(),
        copy_max_depth: copy.max_depth(),
        copy_max_output_bytes: copy.max_output_bytes(),
        bounded: format!("{bounded:?}"),
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

/// Retain each member's raw local record and central-directory record for the
/// independent untouched-member oracle.  The changed main member is excluded
/// by the caller; all other records must match even when the changed member
/// shifts later offsets in the output archive.
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

fn central_record_without_relocation(record: &[u8]) -> Vec<u8> {
    let mut normalized = record.to_vec();
    // The central record's fixed local-header offset necessarily changes when
    // an earlier main member grows.  All other raw central bytes remain part
    // of this oracle; this four-byte relocation field is the sole normalized
    // location for the generated non-ZIP64 corpus.
    if normalized.len() >= 46 {
        normalized[42..46].fill(0);
    }
    normalized
}

fn semantic_record(bytes: &[u8]) -> BenchResult<(SemanticRecord, Vec<String>)> {
    let package = OwnedDocxPackage::from_reader(Cursor::new(bytes.to_vec()))?;
    let document = package.document_snapshot()?;
    let paragraphs = document.paragraphs();
    let mut texts = Vec::with_capacity(paragraphs.len());
    let mut order = Sha256::new();
    let mut text = Sha256::new();
    order.update(b"litchi-docx-bounded-tail-order-v1\0");
    text.update(b"litchi-docx-bounded-tail-text-v1\0");
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

fn build_materialized_candidate(
    source_archive_bytes: &[u8],
    source_xml_bytes: &[u8],
    count: usize,
    limits: CopyLimits,
) -> BenchResult<(Vec<u8>, bool, bool)> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(source_archive_bytes.to_vec()));
    let package = litchi_docx::source_backed::Package::from_read_at(source)?;
    let snapshot = package.plain_paragraph_copy_snapshot_with_limits(limits)?;
    if snapshot.paragraph_count() != count || snapshot.xml_bytes() != source_xml_bytes {
        return Err("materialized source snapshot differs from generated XML".into());
    }
    let mut edit = snapshot.edit();
    edit.copy_plain_paragraph(Position::new(0), Position::new(count))?;
    let commit = edit.commit();
    let expected = independent_candidate_xml(source_xml_bytes, &append_text(count))?;
    let output_xml = commit.projected().xml_bytes();
    let candidate_xml_expected = output_xml == expected.as_slice();
    let wire = commit.patch().to_bytes()?;
    let durable = litchi_docx::source_backed::paragraph_copy::Patch::from_bytes(&wire)?;
    let replayed = durable.apply(&snapshot)?;
    let replay_verified = replayed.xml_bytes() == output_xml;
    if !candidate_xml_expected || !replay_verified {
        return Err("materialized candidate XML or patch replay oracle failed".into());
    }
    let mut output = Vec::new();
    let publication = package.publish_plain_paragraph_copy_to_stream(&mut output, &commit)?;
    let candidate_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(output.clone()));
    let candidate_package = litchi_docx::source_backed::Package::from_read_at(candidate_source)?;
    let mut restored = Vec::new();
    candidate_package
        .publish_plain_paragraph_copy_inverse_to_stream(&mut restored, &publication)?;
    if restored != source_archive_bytes {
        return Err("materialized candidate inverse did not restore source".into());
    }
    Ok((output, replay_verified, restored == source_archive_bytes))
}

fn build_bounded_candidate(
    source_archive_bytes: &[u8],
    append_text: &str,
    limits: TailAppendLimits,
) -> BenchResult<(Vec<u8>, bool, bool, ScalarProofRecord)> {
    // Use the format-owned source-backed DOCX package so admission, topology
    // checks, and the bounded XML proof are part of this route.  Calling OPC
    // directly here would under-measure the format operation and could accept
    // a package that the DOCX closure must refuse.
    let source = litchi_docx::source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        source_archive_bytes.to_vec(),
    )))?;
    let edit = source.tail_append_plain_paragraph_with_limits(append_text, limits)?;
    let plan = edit.prepare()?;
    let source_proof = plan.source_proof();
    let candidate_proof = plan.candidate_proof();
    if source_proof.source_len == 0
        || candidate_proof.candidate_len <= source_proof.source_len
        || !candidate_proof.generated_once
    {
        return Err("bounded candidate scalar proof is incomplete".into());
    }
    let mut output = Vec::new();
    let publication = plan.write_to_stream(&mut output)?;
    let replay_verified = publication.candidate_proof() == candidate_proof;
    let candidate_source = litchi_docx::source_backed::Package::from_read_at(Arc::new(
        OwnedSource::new(output.clone()),
    ))?;
    let mut restored = Vec::new();
    publication.write_inverse_to_stream(&candidate_source, &mut restored)?;
    if restored != source_archive_bytes {
        return Err("bounded candidate inverse did not restore source".into());
    }
    Ok((
        output,
        replay_verified,
        restored == source_archive_bytes,
        scalar_proof_record(source_proof, candidate_proof),
    ))
}

fn materialized_stale_source_refusal(
    source_archive_bytes: &[u8],
    count: usize,
    limits: CopyLimits,
) -> BenchResult<bool> {
    let source = Arc::new(MeasureSource::new(Arc::from(source_archive_bytes.to_vec())));
    let package =
        litchi_docx::source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
    let snapshot = package.plain_paragraph_copy_snapshot_with_limits(limits)?;
    let mut edit = snapshot.edit();
    edit.copy_plain_paragraph(Position::new(0), Position::new(count))?;
    let commit = edit.commit();
    source.bump_revision();
    let mut output = Vec::new();
    let result = package.publish_plain_paragraph_copy_to_stream(&mut output, &commit);
    Ok(result.is_err() && output.is_empty())
}

fn bounded_stale_source_refusal(
    source_archive_bytes: &[u8],
    append_text: &str,
    limits: TailAppendLimits,
) -> BenchResult<bool> {
    let source = Arc::new(MeasureSource::new(Arc::from(source_archive_bytes.to_vec())));
    let package =
        litchi_docx::source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)?;
    let edit = package.tail_append_plain_paragraph_with_limits(append_text, limits)?;
    let plan = edit.prepare()?;
    source.bump_revision();
    let mut output = Vec::new();
    let result = plan.write_to_stream(&mut output);
    Ok(result.is_err() && output.is_empty())
}

fn bounded_noop_oracle(source_archive_bytes: &[u8], limits: TailAppendLimits) -> BenchResult<bool> {
    let package = litchi_docx::source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        source_archive_bytes.to_vec(),
    )))?;
    let plan = package.tail_append_noop().with_limits(limits).prepare()?;
    if !plan.is_noop() {
        return Ok(false);
    }
    let mut output = Vec::new();
    let publication = plan.write_to_stream(&mut output)?;
    if !publication.is_noop() || output != source_archive_bytes {
        return Ok(false);
    }
    let mut restored = Vec::new();
    publication.write_inverse_to_stream(&package, &mut restored)?;
    Ok(restored == source_archive_bytes)
}

fn route_record(
    route: Route,
    source_texts: &[String],
    source_members: &[MemberIdentity],
    source_opaque: &MemberIdentity,
    output: &[u8],
    expected_main_xml: &[u8],
    append_text: &str,
    scalar_proof: Option<ScalarProofRecord>,
    untouched_raw_members_verified: bool,
    replay_verified: bool,
    inverse_verified: bool,
    stale_source_refusal_verified: bool,
) -> BenchResult<RouteCorpusRecord> {
    let output_members = member_identities(output)?;
    let output_reader = ArchiveReader::new(output)?;
    let output_xml = output_reader.read(MAIN_PATH)?;
    let output_main_xml_expected_raw_verified = output_xml == expected_main_xml;
    let (output_semantic, output_texts) = semantic_record(output)?;
    let output_tail_verified = output_texts.len() == source_texts.len() + 1
        && output_texts.get(..source_texts.len()) == Some(source_texts)
        && output_texts
            .last()
            .is_some_and(|value| value == append_text);
    let output_untouched_members_verified = source_members
        .iter()
        .filter(|member| member.path != MAIN_PATH)
        .all(|member| {
            find_member(&output_members, &member.path).is_ok_and(|candidate| candidate == member)
        });
    let output_opaque_member_exact_verified =
        find_member(&output_members, OPAQUE_PATH)? == source_opaque;
    let output_physical_order_verified = source_members
        .iter()
        .map(|member| member.path.as_str())
        .eq(output_members.iter().map(|member| member.path.as_str()));
    let source_main = find_member(source_members, MAIN_PATH)?;
    let output_main = find_member(&output_members, MAIN_PATH)?;
    let output_main_compressed_equal_source =
        output_main.compressed_sha256 == source_main.compressed_sha256;
    Ok(RouteCorpusRecord {
        route,
        output_archive_bytes: output.len(),
        output_archive_sha256: sha256_hex(output),
        output_main_xml_bytes: output_xml.len(),
        output_main_xml_sha256: sha256_hex(&output_xml),
        output_main_xml_expected_semantic_verified: output_tail_verified,
        output_main_xml_expected_raw_verified,
        scalar_proof,
        output_semantic,
        output_member_count: output_members.len(),
        output_members,
        output_untouched_members_verified,
        output_untouched_raw_members_verified: untouched_raw_members_verified,
        output_opaque_member_exact_verified,
        output_physical_order_verified,
        output_main_compressed_equal_source,
        route_output_replay_verified: replay_verified,
        route_inverse_verified: inverse_verified,
        stale_source_refusal_verified,
    })
}

fn build_corpus(count: usize) -> BenchResult<Corpus> {
    if !DEFAULT_COUNTS.contains(&count) {
        return Err(format!("unsupported DOCX paragraph count {count}").into());
    }
    let append_text = append_text(count);
    let source_xml = source_main_xml(count);
    let copy_limits = copy_limits_for(count, source_xml.len())?;
    let bounded_limits = bounded_limits_for(count, source_xml.len())?;
    let source_archive_vec = source_archive(&source_xml, 0)?;
    let source_guard = source_archive_vec.clone();
    let source_members = member_identities(&source_archive_vec)?;
    let source_reader = ArchiveReader::new(&source_archive_vec)?;
    let source_main_xml_archive = source_reader.read(MAIN_PATH)?;
    let source_main_xml_archive_verified = source_main_xml_archive == source_xml;
    let source_opaque = find_member(&source_members, OPAQUE_PATH)?;
    let source_opaque_bytes = source_reader.read(OPAQUE_PATH)?;
    let source_opaque_member_exact_verified = source_opaque_bytes == opaque_payload(0);
    let (source_semantic, source_texts) = semantic_record(&source_archive_vec)?;
    if source_texts
        .first()
        .is_none_or(|value| value != &append_text)
        || source_semantic.paragraph_count != count
    {
        return Err("source semantic authoring template does not match append text".into());
    }
    let (materialized_vec, materialized_replay, materialized_inverse) =
        build_materialized_candidate(&source_archive_vec, &source_xml, count, copy_limits)?;
    let (bounded_vec, bounded_replay, bounded_inverse, bounded_scalar_proof) =
        build_bounded_candidate(&source_archive_vec, &append_text, bounded_limits)?;
    let materialized_stale =
        materialized_stale_source_refusal(&source_archive_vec, count, copy_limits)?;
    let bounded_stale =
        bounded_stale_source_refusal(&source_archive_vec, &append_text, bounded_limits)?;
    let bounded_noop = bounded_noop_oracle(&source_archive_vec, bounded_limits)?;
    let materialized_expected_xml = independent_candidate_xml(&source_xml, &append_text)?;
    let bounded_expected_xml = independent_bounded_candidate_xml(&source_xml, &append_text)?;
    verify_scalar_proof(
        &bounded_scalar_proof,
        count,
        &source_xml,
        &bounded_expected_xml,
    )?;
    let materialized_members = member_identities(&materialized_vec)?;
    let bounded_members = member_identities(&bounded_vec)?;
    let source_raw_members = raw_member_records(&source_archive_vec)?;
    let materialized_raw_members = raw_member_records(&materialized_vec)?;
    let bounded_raw_members = raw_member_records(&bounded_vec)?;
    let materialized_untouched_raw =
        untouched_raw_members_equal(&source_raw_members, &materialized_raw_members);
    let bounded_untouched_raw =
        untouched_raw_members_equal(&source_raw_members, &bounded_raw_members);
    let materialized = route_record(
        Route::MaterializedParagraphCopy,
        &source_texts,
        &source_members,
        source_opaque,
        &materialized_vec,
        &materialized_expected_xml,
        &append_text,
        None,
        materialized_untouched_raw,
        materialized_replay,
        materialized_inverse,
        materialized_stale,
    )?;
    let bounded = route_record(
        Route::BoundedPlainTextTailAppend,
        &source_texts,
        &source_members,
        source_opaque,
        &bounded_vec,
        &bounded_expected_xml,
        &append_text,
        Some(bounded_scalar_proof),
        bounded_untouched_raw,
        bounded_replay,
        bounded_inverse,
        bounded_stale,
    )?;
    let route_semantics_equal_verified = materialized.output_semantic.paragraph_count
        == bounded.output_semantic.paragraph_count
        && materialized.output_semantic.order_sha256 == bounded.output_semantic.order_sha256
        && materialized.output_semantic.text_sha256 == bounded.output_semantic.text_sha256;
    let route_untouched_members_equal_verified = materialized_members
        .iter()
        .filter(|member| member.path != MAIN_PATH)
        .eq(bounded_members
            .iter()
            .filter(|member| member.path != MAIN_PATH));
    let source_unchanged_verified = source_archive_vec == source_guard;
    let route_outputs_physical_equal = materialized_vec == bounded_vec;
    if !source_unchanged_verified
        || !source_main_xml_archive_verified
        || !source_opaque_member_exact_verified
        || !materialized.output_main_xml_expected_semantic_verified
        || !bounded.output_main_xml_expected_semantic_verified
        || !materialized.output_main_xml_expected_raw_verified
        || !bounded.output_main_xml_expected_raw_verified
        || !materialized.output_untouched_members_verified
        || !bounded.output_untouched_members_verified
        || !materialized.output_untouched_raw_members_verified
        || !bounded.output_untouched_raw_members_verified
        || !materialized.output_opaque_member_exact_verified
        || !bounded.output_opaque_member_exact_verified
        || !materialized.output_physical_order_verified
        || !bounded.output_physical_order_verified
        || !materialized.stale_source_refusal_verified
        || !bounded.stale_source_refusal_verified
        || !bounded_noop
        || !route_semantics_equal_verified
        || !route_untouched_members_equal_verified
        || !materialized_untouched_raw
        || !bounded_untouched_raw
    {
        return Err("DOCX bounded tail comparison corpus oracle failed".into());
    }
    let append_text_sha256 = sha256_hex(append_text.as_bytes());
    let record_append_text = append_text.clone();
    Ok(Corpus {
        count,
        append_text,
        source_archive: Arc::from(source_archive_vec.clone()),
        materialized_archive: Arc::from(materialized_vec.clone()),
        bounded_archive: Arc::from(bounded_vec.clone()),
        copy_limits,
        bounded_limits,
        record: CorpusRecord {
            generator: GENERATOR,
            format: "DOCX/OOXML/OPC/ZIP",
            count,
            append_text: record_append_text,
            append_text_sha256,
            source_archive_bytes: source_archive_vec.len(),
            source_archive_sha256: sha256_hex(&source_archive_vec),
            source_main_xml_bytes: source_xml.len(),
            source_main_xml_sha256: sha256_hex(&source_xml),
            source_semantic,
            source_members: source_members.clone(),
            source_member_count: source_members.len(),
            source_unchanged_verified,
            source_main_xml_archive_verified,
            source_opaque_member_exact_verified,
            route_semantics_equal_verified,
            route_untouched_members_equal_verified,
            route_outputs_physical_equal,
            route_untouched_raw_members_equal_verified: materialized_untouched_raw
                && bounded_untouched_raw,
            bounded_noop_verified: bounded_noop,
            materialized,
            bounded,
            limits: limits_record(copy_limits, bounded_limits),
        },
    })
}

fn elapsed_ns(start: Instant) -> BenchResult<u64> {
    u64::try_from(start.elapsed().as_nanos())
        .map_err(|_| "DOCX comparison elapsed duration overflows u64".into())
}

fn check_sink(expected_archive: &[u8], sink: SinkObservation) -> BenchResult<()> {
    if sink.accepted_bytes != u64::try_from(expected_archive.len())?
        || sink.digest.as_slice() != Sha256::digest(expected_archive).as_slice()
    {
        return Err("DOCX comparison runtime output differs from route archive oracle".into());
    }
    Ok(())
}

fn run_materialized_iteration(
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
        let package = litchi_docx::source_backed::Package::from_read_at(
            Arc::clone(&source) as Arc<dyn ReadAt>
        )?;
        let snapshot = package.plain_paragraph_copy_snapshot_with_limits(corpus.copy_limits)?;
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

fn run_bounded_iteration(
    corpus: &Corpus,
) -> BenchResult<(
    u64,
    SinkObservation,
    ReadObservation,
    Option<allocation_metrics::Sample>,
    Option<process_metrics::Delta>,
)> {
    // The caller owns the deterministic text in the corpus. Borrow it for
    // the complete timed lifecycle so the route measures admission, bounded
    // encoding, scan, edit, and publication without cloning or dropping a
    // caller-text String inside the sample.
    let process_before = process_metrics::Snapshot::read().ok();
    let region = allocation_metrics::begin();
    let start = Instant::now();
    let (sink, reads) = {
        let source = Arc::new(MeasureSource::new(Arc::clone(&corpus.source_archive)));
        let package = litchi_docx::source_backed::Package::from_read_at(
            Arc::clone(&source) as Arc<dyn ReadAt>
        )?;
        let edit = package.tail_append_plain_paragraph_with_limits(
            corpus.append_text.as_str(),
            corpus.bounded_limits,
        )?;
        let plan = edit.prepare()?;
        let mut output = HashingSink::new();
        let _publication = plan.write_to_stream(&mut output)?;
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

fn run_route(corpus: &Corpus, route: Route, config: &Config) -> BenchResult<RouteReport> {
    let expected = match route {
        Route::MaterializedParagraphCopy => corpus.materialized_archive.as_ref(),
        Route::BoundedPlainTextTailAppend => corpus.bounded_archive.as_ref(),
    };
    let route_corpus = match route {
        Route::MaterializedParagraphCopy => corpus.record.materialized.clone(),
        Route::BoundedPlainTextTailAppend => corpus.record.bounded.clone(),
    };
    let mut samples = Vec::with_capacity(config.samples);
    for iteration in 0..config.warmups.saturating_add(config.samples) {
        let (elapsed, sink, reads, allocation, process) = match route {
            Route::MaterializedParagraphCopy => run_materialized_iteration(corpus)?,
            Route::BoundedPlainTextTailAppend => run_bounded_iteration(corpus)?,
        };
        check_sink(expected, sink)?;
        if !reads.calls.eq(&0) && !reads.returned_bytes.eq(&0) {
            // The comparison requires actual source reads.  The branch is
            // deliberately only an observation; the explicit error below
            // keeps a silent source/cache regression from becoming evidence.
        } else {
            return Err(format!(
                "{} route performed no positional source reads",
                route.name()
            )
            .into());
        }
        if iteration >= config.warmups {
            samples.push(Sample {
                sample: iteration - config.warmups,
                route,
                elapsed_ns: elapsed,
                source_reads: reads,
                sink: sink.record(),
                allocation,
                process,
            });
        }
    }
    Ok(RouteReport {
        route,
        corpus: route_corpus,
        samples,
    })
}

fn parse_args<I>(args: I) -> BenchResult<Config>
where
    I: IntoIterator<Item = OsString>,
{
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut samples = DEFAULT_SAMPLES;
    let mut warmups = DEFAULT_WARMUPS;
    let mut route = None;
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
            "--route" => {
                route = Some(Route::parse(
                    next_value(&mut values, "--route")?
                        .to_str()
                        .ok_or("--route must be UTF-8")?,
                )?)
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
        route: route.unwrap_or(Route::MaterializedParagraphCopy),
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
    "usage: docx_bounded_tail_append_compare [--route materialized|bounded] [--counts 64,8192,131072] [--samples N] [--warmups N] [--json PATH]"
}

pub fn run_from_args<I>(args: I) -> BenchResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let config = parse_args(args)?;
    let mut cases = Vec::with_capacity(config.counts.len());
    for &count in &config.counts {
        let corpus = build_corpus(count)?;
        let routes = vec![run_route(&corpus, config.route, &config)?];
        cases.push(CaseReport {
            count,
            corpus: corpus.record.clone(),
            routes,
        });
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
            route: config.route,
            lifecycle: "source-backed open + route preparation + commit/plan + sequential publication + sink finalization + operation drops; corpus construction and independent oracles outside",
            sink: "non_seek_hashing_scalar_sha256_shortwrite_16k",
            source: "caller_owned_arc_positional_read_at_scalar_counters_requested_returned_fixed_histogram",
            text_authoring: "bounded route borrows caller-owned UTF-8 text matching source paragraph zero and authors a locally xmlns:w-bound xml:space=preserve run; text value/storage outside clock; materialized route copies paragraph zero",
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
    fn source_text_is_explicit_and_escaped() {
        let text = append_text(64);
        let xml = String::from_utf8(source_main_xml(64)).unwrap();
        assert!(xml.contains("&lt;&amp;&gt;"));
        assert_eq!(xml.matches("<w:p>").count(), 64);
        assert_eq!(
            xml_escape(&text),
            "tail-text-000064-café &lt;&amp;&gt; plain"
        );
    }

    #[test]
    fn independent_candidate_adds_one_exact_fragment() {
        let source = source_main_xml(64);
        let text = append_text(64);
        let candidate = independent_candidate_xml(&source, &text).unwrap();
        let candidate = String::from_utf8(candidate).unwrap();
        assert_eq!(candidate.matches("<w:p>").count(), 65);
        assert!(candidate.ends_with("</w:body></w:document>"));
    }

    #[test]
    fn bounded_candidate_oracle_tracks_local_namespace_and_space_policy() {
        let source = source_main_xml(64);
        let candidate = String::from_utf8(
            independent_bounded_candidate_xml(&source, &append_text(64)).unwrap(),
        )
        .unwrap();
        assert!(candidate.contains(
            "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t xml:space=\"preserve\">"
        ));
    }

    #[test]
    fn cli_accepts_frozen_counts_and_routes() {
        let config = parse_args([
            OsString::from("--counts"),
            OsString::from("64,131072,64"),
            OsString::from("--samples"),
            OsString::from("2"),
            OsString::from("--warmups"),
            OsString::from("1"),
            OsString::from("--route"),
            OsString::from("bounded"),
        ])
        .unwrap();
        assert_eq!(config.counts, [64, 131_072]);
        assert_eq!(config.route, Route::BoundedPlainTextTailAppend);
        assert!(parse_args([OsString::from("--counts"), OsString::from("65")]).is_err());
    }

    #[test]
    fn source_adapter_counts_logical_reads() {
        let source = MeasureSource::new(Arc::<[u8]>::from(vec![1, 2, 3, 4]));
        let mut buffer = [0u8; 3];
        assert_eq!(source.read_at(1, &mut buffer).unwrap(), 3);
        assert_eq!(source.read_at(4, &mut buffer).unwrap(), 0);
        let reads = source.reads();
        assert_eq!(reads.calls, 2);
        assert_eq!(reads.requested_bytes, 6);
        assert_eq!(reads.returned_bytes, 3);
    }
}
