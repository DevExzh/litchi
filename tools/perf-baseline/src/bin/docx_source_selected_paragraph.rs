//! Bounded resource evidence for one selected source-backed DOCX paragraph.
//!
//! This binary deliberately does not measure elapsed time. It builds a fixed
//! media-rich package, runs eager and facade semantic oracles, and classifies
//! logical ReadAt ranges for unmanaged and managed source-backed owners.

use std::collections::{BTreeMap, BTreeSet};
use std::env;
use std::error::Error;
use std::fmt::Write as _;
use std::fs::{self, File, OpenOptions};
use std::io::{self, Read, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, Resource,
    SourceVersion,
};
use litchi_docx::source_backed;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, SourceCacheDiagnostics,
    TargetMode,
};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.docx.source-selected-paragraph.v1";
const CORPUS_GENERATOR: &str = "litchi-docx-source-selected-paragraph-media-v2";
const MAIN_MEMBER: &str = "word/document.xml";
const CORE_MEMBER: &str = "docProps/core.xml";
const TARGET_INDEX: usize = 100;
const TARGET_TEXT: &str = "source-selected-paragraph-0100";
const EXPECTED_PARAGRAPH_COUNT: usize = 201;
const MEDIA_COUNT: usize = 8;
const MEDIA_PAYLOAD_BYTES: usize = 2 * 1024 * 1024;
const MIN_ARCHIVE_BYTES: usize = 15 * 1024 * 1024;
const MAX_ARCHIVE_BYTES: usize = 20 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 1024 * 1024;
const MAX_RANGE_RECORDS: usize = 16 * 1024;
const MAX_SOURCE_READ_BYTES: usize = 16 * 1024 * 1024;
const MAX_COMMAND_OUTPUT_BYTES: usize = 64 * 1024;
const HASH_BUFFER_BYTES: usize = 64 * 1024;
const MAX_EXECUTABLE_HASH_BYTES: u64 = 512 * 1024 * 1024;
const MAX_CARGO_LOCK_HASH_BYTES: u64 = 16 * 1024 * 1024;
const SOURCE_ID: u64 = 0x4458_5053_5250_5631;
const EXPECTED_ARCHIVE_SHA256: Option<&str> =
    Some("a4384c2c249ef87bac6150f92b1a839d4555872f5c9b6ffe6b3d849f47bb7fef");
const EXPECTED_ARCHIVE_BYTES: Option<usize> = Some(16_786_572);
const EXPECTED_MAIN_SHA256: Option<&str> =
    Some("e113d449b7d0a59409405c923e11d3fedc32778cb5eaa91cf003270a3807b3cd");
const EXPECTED_MAIN_BYTES: Option<usize> = Some(12_776);
const EXPECTED_MEMBER_IDENTITIES: &[(&str, usize, &str)] = &[
    (
        "[Content_Types].xml",
        818,
        "849f47a7ef8b85296273d04133f31573d64c9d71de84feba05e5dc20a475fdb6",
    ),
    (
        "_rels/.rels",
        442,
        "b7ae6e1849778a0fa08ce4ad648050804c83b7a055e97ad92c7c234156896480",
    ),
    (
        "docProps/core.xml",
        218,
        "6126834278f255cf3d5774fd2d2609dd745d5192977feb8d636cfc91590e7ad9",
    ),
    (
        "word/_rels/document.xml.rels",
        1475,
        "59751058203e36f24ef75fa098982ff721520dd9e2df0322645a9f5410be2a8f",
    ),
    (
        "word/document.xml",
        12_776,
        "e113d449b7d0a59409405c923e11d3fedc32778cb5eaa91cf003270a3807b3cd",
    ),
    (
        "word/media/source-selected-00.png",
        2_097_152,
        "2429ba44c99c1d4813c334764a892b21f080ec004ec299517ee61e6007e3fefe",
    ),
    (
        "word/media/source-selected-01.png",
        2_097_152,
        "c74a391bd89d586bdbde2490325f9a9cbb0a4c219fa79c69c232e567016d5acb",
    ),
    (
        "word/media/source-selected-02.png",
        2_097_152,
        "a71bb0f89f880442c0e0fad2a0583ca79a9660c9d1c3eb691fd1386e55648364",
    ),
    (
        "word/media/source-selected-03.png",
        2_097_152,
        "b5660cc9edc7eb6b2662c0ab945601cca5f3f4095c57854fb40c986be42fcfdd",
    ),
    (
        "word/media/source-selected-04.png",
        2_097_152,
        "c06d5134493e56f256c47e8d2a5b8e63840032676c6db0bf69ce2a67ea5f275a",
    ),
    (
        "word/media/source-selected-05.png",
        2_097_152,
        "e41b5a7457e9a7ac7db866cb237f44708768786314d46e29e4aa378c508ce3e5",
    ),
    (
        "word/media/source-selected-06.png",
        2_097_152,
        "b374568e8768e6874d6d7be8fbbd06b80ce6242819606b1871752f97b4a1d783",
    ),
    (
        "word/media/source-selected-07.png",
        2_097_152,
        "c49830169bb2cdf662ca986fd15072ba8c980b00c98ff522cff7e92749d7bd07",
    ),
    (
        "word/opaque/source-selected.bin",
        2048,
        "285ce05337c55fe794fb758e94bf644419e885b58b5d01fdc0c6727216896906",
    ),
    (
        "word/settings.xml",
        84,
        "23bbc7637f0ee8acd3365be9b554ffab913e6457f089e77fe909bb659b7c222f",
    ),
];
const EXPECTED_TARGET_TEXT: &str = "source-selected-paragraph-0100";
const EXPECTED_GIT_REVISION: Option<&str> = option_env!("LITCHI_DOCX_SELECTED_GIT_REVISION");
const EXPECTED_CARGO_LOCK_SHA256: Option<&str> =
    option_env!("LITCHI_DOCX_SELECTED_CARGO_LOCK_SHA256");
const EXPECTED_PROFILE: Option<&str> = option_env!("LITCHI_DOCX_SELECTED_PROFILE");
const EXPECTED_FEATURES: Option<&str> = option_env!("LITCHI_DOCX_SELECTED_FEATURES");
const EXPECTED_BUILD_COMMAND: Option<&str> = option_env!("LITCHI_DOCX_SELECTED_BUILD_COMMAND");
const EXPECTED_RUSTC_VV_SHA256: Option<&str> = option_env!("LITCHI_DOCX_SELECTED_RUSTC_VV_SHA256");
const EXECUTABLE_IDENTITY_POLICY: &str = "runtime self-identity only; authoritative acceptance requires an independently retained artifact-manifest comparison";

const CLAIM_SCOPE: &str = "selected paragraph access after one main-document materialization; logical ReadAt/cache/managed-budget evidence only";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RunMode {
    Authoritative,
    DiscoverCorpus,
    Smoke,
}

impl RunMode {
    fn as_str(self) -> &'static str {
        match self {
            Self::Authoritative => "authoritative",
            Self::DiscoverCorpus => "discover-corpus",
            Self::Smoke => "smoke",
        }
    }

    fn claim_authorized(self) -> bool {
        matches!(self, Self::Authoritative)
    }
}

#[derive(Debug, Serialize)]
struct ExpectedMemberIdentity {
    name: &'static str,
    bytes: usize,
    sha256: &'static str,
}

#[derive(Debug, Serialize)]
struct ExpectedIdentityPlaceholders {
    archive_sha256: Option<&'static str>,
    archive_bytes: Option<usize>,
    main_document_sha256: Option<&'static str>,
    main_document_bytes: Option<usize>,
    member_identities: Vec<ExpectedMemberIdentity>,
    target_text: &'static str,
    git_revision: Option<&'static str>,
    cargo_lock_sha256: Option<&'static str>,
    profile: Option<&'static str>,
    features: Option<&'static str>,
    build_command: Option<&'static str>,
    rustc_vv_sha256: Option<&'static str>,
}

#[derive(Debug)]
struct Corpus {
    archive: Arc<[u8]>,
    archive_sha256: String,
    members: BTreeMap<String, MemberIdentity>,
    ranges: Arc<Vec<ZipMemberRange>>,
}

#[derive(Clone, Debug, Serialize)]
struct MemberIdentity {
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug)]
struct ZipMemberRange {
    name: String,
    data: Range<u64>,
}

#[derive(Debug, Serialize)]
struct Evidence {
    schema: &'static str,
    run_mode: &'static str,
    claim_authorized: bool,
    claim_scope: &'static str,
    performance_claim: &'static str,
    sha256_sidecar_path: Option<String>,
    corpus: CorpusReport,
    provenance: Provenance,
    target: TargetReport,
    oracles: OracleReport,
    source_runs: [SourceRun; 2],
    unavailable: BTreeMap<&'static str, &'static str>,
}

#[derive(Debug, Serialize)]
struct CorpusReport {
    generator: &'static str,
    archive_bytes: usize,
    archive_sha256: String,
    archive_min_bytes: usize,
    archive_max_bytes: usize,
    member_count: usize,
    media_member_count: usize,
    media_payload_bytes: usize,
    direct_body_paragraph_count: usize,
    target_index: usize,
    members: BTreeMap<String, MemberIdentity>,
}

#[derive(Debug, Serialize)]
struct TargetReport {
    direct_body_index: usize,
    expected_paragraph_count: usize,
    expected_text: &'static str,
}

#[derive(Debug, Serialize)]
struct OracleReport {
    eager: OracleStatus,
    facade: FacadeStatus,
}

#[derive(Debug, Serialize)]
struct OracleStatus {
    selected_text: String,
    out_of_bounds_none: bool,
}

#[derive(Debug, Serialize)]
struct FacadeStatus {
    selected_text: String,
    out_of_bounds_none: bool,
    stale_source_changed: bool,
}

#[derive(Debug, Serialize)]
struct SourceRun {
    mode: &'static str,
    selected_text: String,
    out_of_bounds_none: bool,
    source_version_before: VersionReport,
    source_version_after_queries: VersionReport,
    source_version_unchanged_before_stale: bool,
    open_without_semantic_payload_member_reads: PhaseEvidence,
    document_materialization_main_document_only_semantic_payload: PhaseEvidence,
    paragraph_query: PhaseEvidence,
    out_of_bounds_query: PhaseEvidence,
    stale_refusal: StaleEvidence,
    cache_after_document_materialization: CacheStats,
    cache_after_queries: CacheStats,
    budget: Option<BudgetEvidence>,
}

#[derive(Debug, Serialize)]
struct VersionReport {
    id: u64,
    revision: u64,
}

#[derive(Debug, Serialize)]
struct PhaseEvidence {
    source: ReadDelta,
    cache: CacheDelta,
}

#[derive(Debug, Serialize)]
struct StaleEvidence {
    snapshot_stable_selected_text: bool,
    snapshot_query_no_source_work: bool,
    package_reentry_typed_source_changed: bool,
    snapshot_semantics: &'static str,
    observed_revision: u64,
    source: ReadDelta,
    cache: CacheDelta,
}

#[derive(Debug, Serialize)]
struct BudgetEvidence {
    memory_before_open: u64,
    memory_after_document: u64,
    memory_after_drop: u64,
    released_after_drop: bool,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct CacheStats {
    hits: u64,
    cold_loads: u64,
    waiter_joins: u64,
    successful_loads: u64,
    failed_loads: u64,
    evictions: u64,
    bypasses: u64,
    oversized_bypasses: u64,
    allocation_bypasses: u64,
    retained_entries: usize,
    retained_bytes: usize,
    in_flight_loads: usize,
    budget_managed: bool,
    budget_reservation_failures: u64,
    budget_memory_used: u64,
    budget_cache_reserved_bytes: u64,
    budget_memory_limit: Option<u64>,
    budget_input_bytes_used: u64,
    budget_input_bytes_limit: Option<u64>,
    budget_output_bytes_used: u64,
    budget_output_bytes_limit: Option<u64>,
    budget_work_used: u64,
    budget_work_limit: Option<u64>,
    budget_objects_used: u64,
    budget_objects_limit: Option<u64>,
    budget_catalog_reserved_objects: u64,
    budget_cache_reserved_objects: u64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct CacheDelta {
    hits: u64,
    cold_loads: u64,
    waiter_joins: u64,
    successful_loads: u64,
    failed_loads: u64,
    evictions: u64,
    bypasses: u64,
    oversized_bypasses: u64,
    allocation_bypasses: u64,
    budget_reservation_failures: u64,
}

#[derive(Clone, Debug, Serialize)]
struct ReadDelta {
    read_calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    len_calls: u64,
    version_calls: u64,
    range_records: usize,
    zip_structural_requested_bytes: u64,
    zip_structural_returned_bytes: u64,
    unclassified_requested_bytes: u64,
    unclassified_returned_bytes: u64,
    class_requested_bytes: BTreeMap<String, u64>,
    class_returned_bytes: BTreeMap<String, u64>,
    member_reads: BTreeMap<String, MemberRead>,
}

impl ReadDelta {
    fn reconciles(&self) -> bool {
        let requested_classes = self
            .class_requested_bytes
            .values()
            .try_fold(0_u64, |total, value| total.checked_add(*value));
        let returned_classes = self
            .class_returned_bytes
            .values()
            .try_fold(0_u64, |total, value| total.checked_add(*value));
        requested_classes.and_then(|total| total.checked_add(self.unclassified_requested_bytes))
            == Some(self.requested_bytes)
            && returned_classes
                .and_then(|total| total.checked_add(self.unclassified_returned_bytes))
                == Some(self.returned_bytes)
            && self
                .class_requested_bytes
                .get("zip_structural")
                .copied()
                .unwrap_or(0)
                == self.zip_structural_requested_bytes
            && self
                .class_returned_bytes
                .get("zip_structural")
                .copied()
                .unwrap_or(0)
                == self.zip_structural_returned_bytes
            && self.unclassified_requested_bytes == 0
            && self.unclassified_returned_bytes == 0
    }
}

#[derive(Clone, Debug, Serialize)]
struct MemberRead {
    class: &'static str,
    calls: u64,
    range_count: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    first_offset: Option<u64>,
    last_offset: Option<u64>,
}

#[derive(Debug, Clone, Copy)]
struct ReadSnapshot {
    read_calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    len_calls: u64,
    version_calls: u64,
    range_records: usize,
}

#[derive(Debug, Clone, Copy)]
struct RangeRecord {
    offset: u64,
    requested: u64,
    returned: u64,
}

struct SequenceGuard<'a> {
    sequence: &'a AtomicU64,
}

impl<'a> SequenceGuard<'a> {
    fn new(sequence: &'a AtomicU64) -> Self {
        sequence.fetch_add(1, Ordering::SeqCst);
        Self { sequence }
    }
}

impl Drop for SequenceGuard<'_> {
    fn drop(&mut self) {
        self.sequence.fetch_add(1, Ordering::SeqCst);
    }
}

struct CountingSource {
    bytes: Arc<[u8]>,
    ranges: Arc<Vec<ZipMemberRange>>,
    read_calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    len_calls: AtomicU64,
    version_calls: AtomicU64,
    revision: AtomicU64,
    sequence: AtomicU64,
    metrics_failed: AtomicBool,
    records: Mutex<Vec<RangeRecord>>,
}

impl CountingSource {
    fn new(bytes: Arc<[u8]>, ranges: Arc<Vec<ZipMemberRange>>) -> Self {
        Self {
            bytes,
            ranges,
            read_calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            len_calls: AtomicU64::new(0),
            version_calls: AtomicU64::new(0),
            revision: AtomicU64::new(0),
            sequence: AtomicU64::new(0),
            metrics_failed: AtomicBool::new(false),
            records: Mutex::new(Vec::new()),
        }
    }

    fn increment(counter: &AtomicU64, amount: u64, label: &'static str) -> io::Result<()> {
        counter
            .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |value| {
                value.checked_add(amount)
            })
            .map(|_| ())
            .map_err(|_| io::Error::other(format!("source counter overflow: {label}")))
    }

    fn snapshot(&self) -> AnyResult<ReadSnapshot> {
        if self.metrics_failed.load(Ordering::SeqCst) {
            return Err("source metrics became unavailable".into());
        }
        loop {
            let before = self.sequence.load(Ordering::SeqCst);
            if before % 2 != 0 {
                std::hint::spin_loop();
                continue;
            }
            let snapshot = ReadSnapshot {
                read_calls: self.read_calls.load(Ordering::SeqCst),
                requested_bytes: self.requested_bytes.load(Ordering::SeqCst),
                returned_bytes: self.returned_bytes.load(Ordering::SeqCst),
                len_calls: self.len_calls.load(Ordering::SeqCst),
                version_calls: self.version_calls.load(Ordering::SeqCst),
                range_records: self
                    .records
                    .lock()
                    .map_err(|_| "source range metrics mutex was poisoned")?
                    .len(),
            };
            let after = self.sequence.load(Ordering::SeqCst);
            if before == after && after % 2 == 0 {
                return Ok(snapshot);
            }
        }
    }

    fn delta(&self, before: ReadSnapshot) -> AnyResult<ReadDelta> {
        let (after, new_records) = loop {
            let sequence_before = self.sequence.load(Ordering::SeqCst);
            if sequence_before % 2 != 0 {
                std::hint::spin_loop();
                continue;
            }
            let after = self.snapshot()?;
            if after.range_records < before.range_records {
                return Err("source range-record counter decreased".into());
            }
            let new_records = self
                .records
                .lock()
                .map_err(|_| "source range metrics mutex was poisoned")?
                .get(before.range_records..)
                .ok_or("source range-record interval was invalid")?
                .to_vec();
            let sequence_after = self.sequence.load(Ordering::SeqCst);
            if sequence_before == sequence_after && sequence_after % 2 == 0 {
                break (after, new_records);
            }
        };
        let read_calls = after
            .read_calls
            .checked_sub(before.read_calls)
            .ok_or("source read-call counter decreased")?;
        let requested_bytes = after
            .requested_bytes
            .checked_sub(before.requested_bytes)
            .ok_or("source requested-byte counter decreased")?;
        let returned_bytes = after
            .returned_bytes
            .checked_sub(before.returned_bytes)
            .ok_or("source returned-byte counter decreased")?;
        let len_calls = after
            .len_calls
            .checked_sub(before.len_calls)
            .ok_or("source len-call counter decreased")?;
        let version_calls = after
            .version_calls
            .checked_sub(before.version_calls)
            .ok_or("source version-call counter decreased")?;
        let mut class_requested_bytes = BTreeMap::new();
        let mut class_returned_bytes = BTreeMap::new();
        let mut member_reads = BTreeMap::new();
        let mut classified_requested_bytes = 0_u64;
        let mut classified_returned_bytes = 0_u64;
        for record in &new_records {
            let requested_end = record
                .offset
                .checked_add(record.requested)
                .ok_or("source requested range overflow")?;
            let returned_end = record
                .offset
                .checked_add(record.returned)
                .ok_or("source returned range overflow")?;
            let mut record_classified_requested_bytes = 0_u64;
            let mut record_classified_returned_bytes = 0_u64;
            for member in self.ranges.iter() {
                let requested_overlap = overlap(record.offset..requested_end, member.data.clone());
                let returned_overlap = overlap(record.offset..returned_end, member.data.clone());
                if requested_overlap.is_empty() && returned_overlap.is_empty() {
                    continue;
                }
                let requested_len = requested_overlap.end - requested_overlap.start;
                let returned_len = returned_overlap.end - returned_overlap.start;
                let class = member_class(&member.name);
                add_u64(&mut class_requested_bytes, class, requested_len)?;
                add_u64(&mut class_returned_bytes, class, returned_len)?;
                classified_requested_bytes = classified_requested_bytes
                    .checked_add(requested_len)
                    .ok_or("classified requested-byte counter overflow")?;
                classified_returned_bytes = classified_returned_bytes
                    .checked_add(returned_len)
                    .ok_or("classified returned-byte counter overflow")?;
                record_classified_requested_bytes = record_classified_requested_bytes
                    .checked_add(requested_len)
                    .ok_or("record classified requested-byte counter overflow")?;
                record_classified_returned_bytes = record_classified_returned_bytes
                    .checked_add(returned_len)
                    .ok_or("record classified returned-byte counter overflow")?;
                let entry = member_reads
                    .entry(member.name.clone())
                    .or_insert_with(|| MemberRead {
                        class,
                        calls: 0,
                        range_count: 0,
                        requested_bytes: 0,
                        returned_bytes: 0,
                        first_offset: None,
                        last_offset: None,
                    });
                entry.calls = entry
                    .calls
                    .checked_add(1)
                    .ok_or("member read-call counter overflow")?;
                entry.range_count = entry
                    .range_count
                    .checked_add(1)
                    .ok_or("member range counter overflow")?;
                entry.requested_bytes = entry
                    .requested_bytes
                    .checked_add(requested_len)
                    .ok_or("member requested-byte counter overflow")?;
                entry.returned_bytes = entry
                    .returned_bytes
                    .checked_add(returned_len)
                    .ok_or("member returned-byte counter overflow")?;
                entry.first_offset = Some(
                    entry
                        .first_offset
                        .map_or(record.offset, |value| value.min(record.offset)),
                );
                entry.last_offset = Some(
                    entry
                        .last_offset
                        .map_or(record.offset, |value| value.max(record.offset)),
                );
            }
            let structural_other_requested = record
                .requested
                .checked_sub(record_classified_requested_bytes)
                .ok_or("record classified requested bytes exceeded the source request")?;
            let structural_other_returned = record
                .returned
                .checked_sub(record_classified_returned_bytes)
                .ok_or("record classified returned bytes exceeded the source result")?;
            if structural_other_requested != 0 || structural_other_returned != 0 {
                add_u64(
                    &mut class_requested_bytes,
                    "zip_structural",
                    structural_other_requested,
                )?;
                add_u64(
                    &mut class_returned_bytes,
                    "zip_structural",
                    structural_other_returned,
                )?;
                classified_requested_bytes = classified_requested_bytes
                    .checked_add(structural_other_requested)
                    .ok_or("structural residual requested-byte counter overflow")?;
                classified_returned_bytes = classified_returned_bytes
                    .checked_add(structural_other_returned)
                    .ok_or("structural residual returned-byte counter overflow")?;
                let entry = member_reads
                    .entry("<zip-structural-other>".to_owned())
                    .or_insert_with(|| MemberRead {
                        class: "zip_structural",
                        calls: 0,
                        range_count: 0,
                        requested_bytes: 0,
                        returned_bytes: 0,
                        first_offset: None,
                        last_offset: None,
                    });
                entry.calls = entry
                    .calls
                    .checked_add(1)
                    .ok_or("structural residual read-call counter overflow")?;
                entry.range_count = entry
                    .range_count
                    .checked_add(1)
                    .ok_or("structural residual range counter overflow")?;
                entry.requested_bytes = entry
                    .requested_bytes
                    .checked_add(structural_other_requested)
                    .ok_or("structural residual requested-byte counter overflow")?;
                entry.returned_bytes = entry
                    .returned_bytes
                    .checked_add(structural_other_returned)
                    .ok_or("structural residual returned-byte counter overflow")?;
                entry.first_offset = Some(
                    entry
                        .first_offset
                        .map_or(record.offset, |value| value.min(record.offset)),
                );
                entry.last_offset = Some(
                    entry
                        .last_offset
                        .map_or(record.offset, |value| value.max(record.offset)),
                );
            }
        }
        if classified_requested_bytes != requested_bytes
            || classified_returned_bytes != returned_bytes
        {
            return Err("source byte partition did not cover the complete source delta".into());
        }
        let unclassified_requested_bytes = 0;
        let unclassified_returned_bytes = 0;
        let delta = ReadDelta {
            read_calls,
            requested_bytes,
            returned_bytes,
            len_calls,
            version_calls,
            range_records: new_records.len(),
            zip_structural_requested_bytes: class_requested_bytes
                .get("zip_structural")
                .copied()
                .unwrap_or(0),
            zip_structural_returned_bytes: class_returned_bytes
                .get("zip_structural")
                .copied()
                .unwrap_or(0),
            unclassified_requested_bytes,
            unclassified_returned_bytes,
            class_requested_bytes,
            class_returned_bytes,
            member_reads,
        };
        if !delta.reconciles() {
            return Err("source byte-accounting reconciliation failed".into());
        }
        Ok(delta)
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn current_version(&self) -> SourceVersion {
        SourceVersion::new(SOURCE_ID, self.revision.load(Ordering::SeqCst))
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        let _sequence = SequenceGuard::new(&self.sequence);
        Self::increment(&self.len_calls, 1, "len calls")?;
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source length overflow"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let _sequence = SequenceGuard::new(&self.sequence);
        if output.len() > MAX_SOURCE_READ_BYTES {
            self.metrics_failed.store(true, Ordering::SeqCst);
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "source read request exceeded runner bound",
            ));
        }
        Self::increment(&self.read_calls, 1, "read calls")?;
        let requested = u64::try_from(output.len())
            .map_err(|_| io::Error::other("source request length overflow"))?;
        Self::increment(&self.requested_bytes, requested, "requested bytes")?;
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source offset overflow"))?;
        let returned = self.bytes.get(start..).map_or(0, |tail| {
            let count = tail.len().min(output.len());
            output[..count].copy_from_slice(&tail[..count]);
            count
        });
        let returned_u64 = u64::try_from(returned)
            .map_err(|_| io::Error::other("source returned length overflow"))?;
        Self::increment(&self.returned_bytes, returned_u64, "returned bytes")?;
        let mut records = self
            .records
            .lock()
            .map_err(|_| io::Error::other("source range metrics mutex was poisoned"))?;
        if records.len() >= MAX_RANGE_RECORDS {
            self.metrics_failed.store(true, Ordering::SeqCst);
            return Err(io::Error::other("source range record bound exceeded"));
        }
        records.push(RangeRecord {
            offset,
            requested,
            returned: returned_u64,
        });
        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let _sequence = SequenceGuard::new(&self.sequence);
        Self::increment(&self.version_calls, 1, "version calls")?;
        Ok(self.current_version())
    }
}

fn add_u64(map: &mut BTreeMap<String, u64>, key: &str, value: u64) -> AnyResult<()> {
    let entry = map.entry(key.to_owned()).or_default();
    *entry = entry
        .checked_add(value)
        .ok_or("classified source-byte counter overflow")?;
    Ok(())
}

fn overlap(left: Range<u64>, right: Range<u64>) -> Range<u64> {
    let start = left.start.max(right.start);
    let end = left.end.min(right.end);
    start..end.max(start)
}

fn member_class(name: &str) -> &'static str {
    if name.starts_with("<zip-structural-") {
        "zip_structural"
    } else if name == MAIN_MEMBER {
        "main_document"
    } else if name.starts_with("word/media/") {
        "media"
    } else if name == CORE_MEMBER {
        "core"
    } else if name == "[Content_Types].xml" || name.ends_with(".rels") {
        "catalog_relationship"
    } else {
        "unrelated"
    }
}

impl CacheStats {
    fn from_diagnostics(value: SourceCacheDiagnostics) -> Self {
        Self {
            hits: value.hits,
            cold_loads: value.cold_loads,
            waiter_joins: value.waiter_joins,
            successful_loads: value.successful_loads,
            failed_loads: value.failed_loads,
            evictions: value.evictions,
            bypasses: value.bypasses,
            oversized_bypasses: value.oversized_bypasses,
            allocation_bypasses: value.allocation_bypasses,
            retained_entries: value.retained_entries,
            retained_bytes: value.retained_bytes,
            in_flight_loads: value.in_flight_loads,
            budget_managed: value.budget_managed,
            budget_reservation_failures: value.budget_reservation_failures,
            budget_memory_used: value.budget_memory_used,
            budget_cache_reserved_bytes: value.budget_cache_reserved_bytes,
            budget_memory_limit: value.budget_memory_limit,
            budget_input_bytes_used: value.budget_input_bytes_used,
            budget_input_bytes_limit: value.budget_input_bytes_limit,
            budget_output_bytes_used: value.budget_output_bytes_used,
            budget_output_bytes_limit: value.budget_output_bytes_limit,
            budget_work_used: value.budget_work_used,
            budget_work_limit: value.budget_work_limit,
            budget_objects_used: value.budget_objects_used,
            budget_objects_limit: value.budget_objects_limit,
            budget_catalog_reserved_objects: value.budget_catalog_reserved_objects,
            budget_cache_reserved_objects: value.budget_cache_reserved_objects,
        }
    }
}

impl CacheDelta {
    fn from_snapshots(
        before: SourceCacheDiagnostics,
        after: SourceCacheDiagnostics,
    ) -> AnyResult<Self> {
        let value = SourceCacheDiagnostics::checked_counter_delta(before, after)?;
        Ok(Self {
            hits: value.hits,
            cold_loads: value.cold_loads,
            waiter_joins: value.waiter_joins,
            successful_loads: value.successful_loads,
            failed_loads: value.failed_loads,
            evictions: value.evictions,
            bypasses: value.bypasses,
            oversized_bypasses: value.oversized_bypasses,
            allocation_bypasses: value.allocation_bypasses,
            budget_reservation_failures: value.budget_reservation_failures,
        })
    }

    fn is_zero(self) -> bool {
        self.hits == 0
            && self.cold_loads == 0
            && self.waiter_joins == 0
            && self.successful_loads == 0
            && self.failed_loads == 0
            && self.evictions == 0
            && self.bypasses == 0
            && self.oversized_bypasses == 0
            && self.allocation_bypasses == 0
            && self.budget_reservation_failures == 0
    }
}

fn version_report(value: SourceVersion) -> VersionReport {
    VersionReport {
        id: value.id(),
        revision: value.revision(),
    }
}

fn build_corpus() -> AnyResult<Corpus> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new(format!("/{MAIN_MEMBER}"))?,
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main_xml(),
    );
    for index in 0..MEDIA_COUNT {
        main.rels_mut().try_add_relationship(
            rt::IMAGE.to_owned(),
            format!("media/source-selected-{index:02}.png"),
            format!("rMedia{index:02}"),
            TargetMode::Internal,
        )?;
    }
    main.rels_mut().try_add_relationship(
        rt::SETTINGS.to_owned(),
        "settings.xml".to_owned(),
        "rSettings".to_owned(),
        TargetMode::Internal,
    )?;
    for index in 0..MEDIA_COUNT {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/word/media/source-selected-{index:02}.png"))?,
            ct::PNG.to_owned(),
            media_payload(index),
        )))?;
    }
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/settings.xml")?,
        ct::WML_SETTINGS.to_owned(),
        b"<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"
            .to_vec(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/opaque/source-selected.bin")?,
        "application/octet-stream".to_owned(),
        opaque_payload(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/docProps/core.xml")?,
        ct::OPC_CORE_PROPERTIES.to_owned(),
        core_xml(),
    )))?;
    package.try_add_part(Box::new(main))?;
    package.relate_to(MAIN_MEMBER, rt::OFFICE_DOCUMENT);
    package.relate_to(CORE_MEMBER, rt::CORE_PROPERTIES);
    let archive = PackageWriter::to_bytes(&package)?;
    if archive.len() < MIN_ARCHIVE_BYTES || archive.len() > MAX_ARCHIVE_BYTES {
        return Err(format!(
            "deterministic corpus archive size {} is outside {}..{} bytes",
            archive.len(),
            MIN_ARCHIVE_BYTES,
            MAX_ARCHIVE_BYTES
        )
        .into());
    }
    let ranges = parse_zip_ranges(&archive)?;
    let archive_reader = ArchiveReader::new(&archive)?;
    let mut members = BTreeMap::new();
    for name in archive_reader.file_names() {
        let bytes = archive_reader.read(name)?;
        members.insert(
            name.to_owned(),
            MemberIdentity {
                bytes: bytes.len(),
                sha256: sha256_hex(&bytes),
            },
        );
    }
    let media_member_count = members
        .keys()
        .filter(|name| name.starts_with("word/media/"))
        .count();
    if media_member_count != MEDIA_COUNT {
        return Err(format!(
            "corpus media count mismatch: expected {MEDIA_COUNT}, got {media_member_count}"
        )
        .into());
    }
    for index in 0..MEDIA_COUNT {
        let name = format!("word/media/source-selected-{index:02}.png");
        let media = members
            .get(&name)
            .ok_or_else(|| format!("corpus omitted media member {name}"))?;
        if media.bytes != MEDIA_PAYLOAD_BYTES {
            return Err(format!(
                "corpus media member {name} has {} bytes, expected {MEDIA_PAYLOAD_BYTES}",
                media.bytes
            )
            .into());
        }
    }
    let main_payload = archive_reader.read(MAIN_MEMBER)?;
    let paragraph_count = main_payload
        .windows(b"<w:p>".len())
        .filter(|window| (*window).eq(b"<w:p>"))
        .count();
    if paragraph_count != EXPECTED_PARAGRAPH_COUNT {
        return Err(format!(
            "corpus direct-body paragraph count mismatch: expected {EXPECTED_PARAGRAPH_COUNT}, got {paragraph_count}"
        )
        .into());
    }
    for index in 0..EXPECTED_PARAGRAPH_COUNT {
        let text = expected_paragraph_text(index);
        if main_payload
            .windows(text.len())
            .filter(|window| (*window).eq(text.as_bytes()))
            .count()
            != 1
        {
            return Err(format!("corpus paragraph text identity mismatch at index {index}").into());
        }
    }
    if !members.contains_key(MAIN_MEMBER)
        || !members.contains_key(CORE_MEMBER)
        || !members.contains_key("word/settings.xml")
        || !members.contains_key("word/opaque/source-selected.bin")
    {
        return Err("corpus omitted a required auxiliary member".into());
    }
    Ok(Corpus {
        archive: Arc::from(archive.as_slice()),
        archive_sha256: sha256_hex(&archive),
        members,
        ranges: Arc::new(ranges),
    })
}

fn corpus_report(corpus: &Corpus) -> CorpusReport {
    CorpusReport {
        generator: CORPUS_GENERATOR,
        archive_bytes: corpus.archive.len(),
        archive_sha256: corpus.archive_sha256.clone(),
        archive_min_bytes: MIN_ARCHIVE_BYTES,
        archive_max_bytes: MAX_ARCHIVE_BYTES,
        member_count: corpus.members.len(),
        media_member_count: MEDIA_COUNT,
        media_payload_bytes: MEDIA_PAYLOAD_BYTES,
        direct_body_paragraph_count: EXPECTED_PARAGRAPH_COUNT,
        target_index: TARGET_INDEX,
        members: corpus.members.clone(),
    }
}

fn expected_identity_placeholders() -> ExpectedIdentityPlaceholders {
    ExpectedIdentityPlaceholders {
        archive_sha256: EXPECTED_ARCHIVE_SHA256,
        archive_bytes: EXPECTED_ARCHIVE_BYTES,
        main_document_sha256: EXPECTED_MAIN_SHA256,
        main_document_bytes: EXPECTED_MAIN_BYTES,
        member_identities: EXPECTED_MEMBER_IDENTITIES
            .iter()
            .map(|&(name, bytes, sha256)| ExpectedMemberIdentity {
                name,
                bytes,
                sha256,
            })
            .collect(),
        target_text: EXPECTED_TARGET_TEXT,
        git_revision: EXPECTED_GIT_REVISION,
        cargo_lock_sha256: EXPECTED_CARGO_LOCK_SHA256,
        profile: EXPECTED_PROFILE,
        features: EXPECTED_FEATURES,
        build_command: EXPECTED_BUILD_COMMAND,
        rustc_vv_sha256: EXPECTED_RUSTC_VV_SHA256,
    }
}

fn verify_corpus_identity(corpus: &Corpus) -> AnyResult<()> {
    let expected_archive_sha256 = EXPECTED_ARCHIVE_SHA256
        .ok_or("authoritative corpus archive hash is unpinned; run --discover-corpus")?;
    let expected_archive_bytes = EXPECTED_ARCHIVE_BYTES
        .ok_or("authoritative corpus archive length is unpinned; run --discover-corpus")?;
    let expected_main_sha256 = EXPECTED_MAIN_SHA256
        .ok_or("authoritative main-document hash is unpinned; run --discover-corpus")?;
    let expected_main_bytes = EXPECTED_MAIN_BYTES
        .ok_or("authoritative main-document length is unpinned; run --discover-corpus")?;
    if corpus.archive_sha256 != expected_archive_sha256
        || corpus.archive.len() != expected_archive_bytes
    {
        return Err("generated archive identity differs from the pinned corpus".into());
    }
    let main = corpus
        .members
        .get(MAIN_MEMBER)
        .ok_or("pinned corpus omitted the main document member")?;
    if main.sha256 != expected_main_sha256 || main.bytes != expected_main_bytes {
        return Err("generated main-document identity differs from the pinned corpus".into());
    }
    if EXPECTED_PARAGRAPH_COUNT != 201
        || TARGET_INDEX != 100
        || MEDIA_COUNT != 8
        || MEDIA_PAYLOAD_BYTES != 2 * 1024 * 1024
        || corpus
            .members
            .keys()
            .filter(|name| name.starts_with("word/media/"))
            .count()
            != MEDIA_COUNT
    {
        return Err("compiled corpus shape differs from the pinned v2 shape".into());
    }
    if EXPECTED_TARGET_TEXT != TARGET_TEXT {
        return Err("target text constant is inconsistent with the pinned corpus".into());
    }
    if EXPECTED_MEMBER_IDENTITIES.is_empty() {
        return Err("authoritative member identities are unpinned; run --discover-corpus".into());
    }
    if corpus.members.len() != EXPECTED_MEMBER_IDENTITIES.len() {
        return Err("pinned corpus member count differs from the expected manifest".into());
    }
    for &(name, expected_bytes, expected_hash) in EXPECTED_MEMBER_IDENTITIES {
        let member = corpus
            .members
            .get(name)
            .ok_or_else(|| format!("pinned corpus omitted member {name}"))?;
        if member.bytes != expected_bytes || member.sha256 != expected_hash {
            return Err(format!("pinned member identity differs for {name}").into());
        }
    }
    Ok(())
}

fn main_xml() -> Vec<u8> {
    let mut xml = String::from(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#,
    );
    for index in 0..EXPECTED_PARAGRAPH_COUNT {
        let _ = write!(
            xml,
            "<w:p><w:r><w:t>{}</w:t></w:r></w:p>",
            expected_paragraph_text(index)
        );
    }
    xml.push_str("</w:body></w:document>");
    xml.into_bytes()
}

fn expected_paragraph_text(index: usize) -> String {
    format!("source-selected-paragraph-{index:04}")
}

fn core_xml() -> Vec<u8> {
    br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>source-selected-resource-corpus</dc:title></cp:coreProperties>"#.to_vec()
}

fn media_payload(index: usize) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(MEDIA_PAYLOAD_BYTES);
    bytes.extend_from_slice(b"\x89PNG\r\n\x1a\n");
    let mut state = 0x9e37_79b9_7f4a_7c15_u64
        ^ (u64::try_from(index)
            .unwrap_or(u64::MAX)
            .wrapping_mul(0xd1b5_4a32_d192_ed03_u64));
    while bytes.len() < MEDIA_PAYLOAD_BYTES {
        state ^= state << 7;
        state ^= state >> 9;
        state ^= state << 8;
        bytes.push((state >> 24) as u8);
    }
    bytes
}

fn opaque_payload() -> Vec<u8> {
    (0..2048)
        .map(|offset| ((offset * 13 + 7) % 251) as u8)
        .collect()
}

fn parse_zip_ranges(bytes: &[u8]) -> AnyResult<Vec<ZipMemberRange>> {
    let eocd = bytes
        .windows(22)
        .rposition(|window| window.starts_with(b"PK\x05\x06"))
        .ok_or("ZIP end-of-central-directory record was not found")?;
    let entries = u16::from_le_bytes(bytes[eocd + 10..eocd + 12].try_into()?) as usize;
    let central_size = u32::from_le_bytes(bytes[eocd + 12..eocd + 16].try_into()?) as usize;
    let central_start = u32::from_le_bytes(bytes[eocd + 16..eocd + 20].try_into()?) as usize;
    let central_end = central_start
        .checked_add(central_size)
        .ok_or("ZIP central directory overflow")?;
    if central_end > bytes.len() {
        return Err("ZIP central directory is outside the archive".into());
    }
    let mut cursor = central_start;
    let mut ranges = Vec::with_capacity(entries.saturating_mul(2).saturating_add(2));
    ranges.push(ZipMemberRange {
        name: "<zip-structural-central>".to_owned(),
        data: u64::try_from(central_start)?..u64::try_from(central_end)?,
    });
    let mut names = BTreeSet::new();
    let mut payload_ranges = Vec::with_capacity(entries);
    for _ in 0..entries {
        if bytes
            .get(cursor..cursor + 46)
            .map_or(true, |header| !header.starts_with(b"PK\x01\x02"))
        {
            return Err("invalid ZIP central directory entry".into());
        }
        let compressed = u32::from_le_bytes(bytes[cursor + 20..cursor + 24].try_into()?) as usize;
        let uncompressed = u32::from_le_bytes(bytes[cursor + 24..cursor + 28].try_into()?) as usize;
        let flags = u16::from_le_bytes(bytes[cursor + 8..cursor + 10].try_into()?);
        let name_len = u16::from_le_bytes(bytes[cursor + 28..cursor + 30].try_into()?) as usize;
        let extra_len = u16::from_le_bytes(bytes[cursor + 30..cursor + 32].try_into()?) as usize;
        let comment_len = u16::from_le_bytes(bytes[cursor + 32..cursor + 34].try_into()?) as usize;
        let local_offset = u32::from_le_bytes(bytes[cursor + 42..cursor + 46].try_into()?) as usize;
        let name_end = cursor
            .checked_add(46 + name_len)
            .ok_or("ZIP member name overflow")?;
        let name = std::str::from_utf8(
            bytes
                .get(cursor + 46..name_end)
                .ok_or("ZIP member name is outside archive")?,
        )?
        .to_owned();
        if !names.insert(name.clone()) {
            return Err(format!("duplicate ZIP member name: {name}").into());
        }
        let central_next = cursor
            .checked_add(46 + name_len + extra_len + comment_len)
            .ok_or("ZIP central directory cursor overflow")?;
        if central_next > central_end
            || bytes
                .get(local_offset..local_offset + 30)
                .map_or(true, |header| !header.starts_with(b"PK\x03\x04"))
        {
            return Err("invalid ZIP local header".into());
        }
        let local_name_len =
            u16::from_le_bytes(bytes[local_offset + 26..local_offset + 28].try_into()?) as usize;
        let local_extra_len =
            u16::from_le_bytes(bytes[local_offset + 28..local_offset + 30].try_into()?) as usize;
        let data_start = local_offset
            .checked_add(30 + local_name_len + local_extra_len)
            .ok_or("ZIP payload offset overflow")?;
        let data_end = data_start
            .checked_add(compressed)
            .ok_or("ZIP payload range overflow")?;
        if data_end > bytes.len() {
            return Err("ZIP payload is outside archive".into());
        }
        let local_name_end = local_offset
            .checked_add(30 + local_name_len)
            .ok_or("ZIP local member name overflow")?;
        let local_name = std::str::from_utf8(
            bytes
                .get(local_offset + 30..local_name_end)
                .ok_or("ZIP local member name is outside archive")?,
        )?;
        if local_name != name {
            return Err(format!("ZIP central/local member name mismatch: {name}").into());
        }
        ranges.push(ZipMemberRange {
            name: format!("<zip-structural-local-{name}>"),
            data: u64::try_from(local_offset)?..u64::try_from(data_start)?,
        });
        if flags & 0x0008 == 0 {
            let local_compressed =
                u32::from_le_bytes(bytes[local_offset + 18..local_offset + 22].try_into()?)
                    as usize;
            let local_uncompressed =
                u32::from_le_bytes(bytes[local_offset + 22..local_offset + 26].try_into()?)
                    as usize;
            if local_compressed != compressed || local_uncompressed != uncompressed {
                return Err(format!("ZIP central/local size mismatch: {name}").into());
            }
        }
        if data_start < central_start && data_end > central_start {
            return Err(format!("ZIP payload overlaps central directory: {name}").into());
        }
        if payload_ranges
            .iter()
            .any(|range: &Range<usize>| data_start < range.end && range.start < data_end)
        {
            return Err(format!("ZIP payload ranges overlap: {name}").into());
        }
        payload_ranges.push(data_start..data_end);
        ranges.push(ZipMemberRange {
            name,
            data: u64::try_from(data_start)?..u64::try_from(data_end)?,
        });
        cursor = central_next;
    }
    if cursor != central_end {
        return Err("ZIP central directory length did not match its entries".into());
    }
    ranges.push(ZipMemberRange {
        name: "<zip-structural-eocd>".to_owned(),
        data: u64::try_from(eocd)?..u64::try_from(bytes.len())?,
    });
    let mut sorted_ranges = ranges
        .iter()
        .map(|range| range.data.clone())
        .collect::<Vec<_>>();
    sorted_ranges.sort_by_key(|range| range.start);
    if sorted_ranges
        .windows(2)
        .any(|ranges| ranges[0].end > ranges[1].start)
    {
        return Err("ZIP structural and payload ranges overlap".into());
    }
    Ok(ranges)
}

fn lowercase_hex(bytes: &[u8]) -> String {
    let alphabet = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len().saturating_mul(2));
    for &byte in bytes {
        output.push(alphabet[(byte >> 4) as usize] as char);
        output.push(alphabet[(byte & 0x0f) as usize] as char);
    }
    output
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    let digest = hasher.finalize();
    lowercase_hex(&digest)
}

struct FileHash {
    bytes: usize,
    sha256: String,
}

fn hash_file_bounded(path: &Path, maximum_bytes: u64) -> AnyResult<FileHash> {
    let mut file = File::open(path)?;
    if file.metadata()?.len() > maximum_bytes {
        return Err(format!(
            "file exceeds bounded hashing limit: {} bytes",
            maximum_bytes
        )
        .into());
    }
    let mut hasher = Sha256::new();
    let mut buffer = [0_u8; HASH_BUFFER_BYTES];
    let mut bytes = 0_u64;
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        bytes = bytes
            .checked_add(u64::try_from(count)?)
            .ok_or("streaming hash byte count overflow")?;
        if bytes > maximum_bytes {
            return Err("file exceeded bounded hashing limit while being read".into());
        }
        hasher.update(&buffer[..count]);
    }
    Ok(FileHash {
        bytes: usize::try_from(bytes)?,
        sha256: lowercase_hex(&hasher.finalize()),
    })
}

fn class_bytes(delta: &ReadDelta, class: &str) -> u64 {
    delta.class_returned_bytes.get(class).copied().unwrap_or(0)
}

fn class_requested_bytes(delta: &ReadDelta, class: &str) -> u64 {
    delta.class_requested_bytes.get(class).copied().unwrap_or(0)
}

fn require_no_semantic_payload_overlap(
    delta: &ReadDelta,
    phase: &str,
    class: &str,
) -> AnyResult<()> {
    let requested = class_requested_bytes(delta, class);
    let returned = class_bytes(delta, class);
    if requested != 0 || returned != 0 {
        return Err(format!(
            "{phase} had semantic payload overlap with {class}: {requested} requested and {returned} returned logical bytes"
        )
        .into());
    }
    Ok(())
}

fn require_zero_query(phase: &PhaseEvidence) -> AnyResult<()> {
    if !phase.source.reconciles() {
        return Err("selected paragraph source byte-accounting did not reconcile".into());
    }
    if phase.source.read_calls != 0
        || phase.source.requested_bytes != 0
        || phase.source.returned_bytes != 0
        || phase.source.len_calls != 0
        || phase.source.version_calls != 0
        || phase.source.range_records != 0
        || !phase.cache.is_zero()
    {
        return Err("selected paragraph query performed incremental source work".into());
    }
    Ok(())
}

fn typed_source_changed(error: &litchi_docx::Error) -> bool {
    matches!(
        error,
        litchi_docx::Error::Opc(OpcError::SourceChanged { .. })
    )
}

fn run_eager_oracle(corpus: &Corpus) -> AnyResult<OracleStatus> {
    let opc = OpcPackage::from_bytes(&corpus.archive)?;
    let package = litchi_docx::Package::from_opc_package(opc)?;
    let document = package.document()?;
    let count = document.paragraph_count()?;
    if count != EXPECTED_PARAGRAPH_COUNT {
        return Err(format!("eager paragraph count mismatch: {count}").into());
    }
    let selected = document
        .paragraph(TARGET_INDEX)?
        .map(|paragraph| paragraph.text())
        .transpose()?;
    let out_of_bounds = document.paragraph(usize::MAX)?.is_none();
    if selected.as_deref() != Some(TARGET_TEXT) || !out_of_bounds {
        return Err("eager semantic oracle failed".into());
    }
    Ok(OracleStatus {
        selected_text: selected.ok_or("eager selected paragraph was absent")?,
        out_of_bounds_none: out_of_bounds,
    })
}

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(1);

struct OwnedTempFile {
    path: PathBuf,
}

impl Drop for OwnedTempFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}

fn create_temp_corpus(bytes: &[u8]) -> AnyResult<OwnedTempFile> {
    for attempt in 0..64_u64 {
        let serial = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let path = env::temp_dir().join(format!(
            "litchi-docx-source-selected-{}-{serial}-{attempt}.docx",
            std::process::id()
        ));
        let mut file = match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(file) => file,
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        };
        let owner = OwnedTempFile { path };
        if let Err(error) = file.write_all(bytes).and_then(|_| file.flush()) {
            drop(file);
            drop(owner);
            return Err(error.into());
        }
        drop(file);
        return Ok(owner);
    }
    Err("could not create an owned temporary corpus path after bounded retries".into())
}

fn run_facade_oracle(corpus: &Corpus) -> AnyResult<FacadeStatus> {
    let owner = create_temp_corpus(&corpus.archive)?;
    let path = owner.path.clone();
    let result = (|| {
        let document = litchi::Document::open(&path)?;
        let selected = document.paragraph_text(TARGET_INDEX)?;
        let out_of_bounds = document.paragraph_text(usize::MAX)?.is_none();
        if selected.as_deref() != Some(TARGET_TEXT) || !out_of_bounds {
            return Err("facade semantic oracle failed".into());
        }
        let mut append = OpenOptions::new().append(true).open(&path)?;
        append.write_all(b"source mutation")?;
        append.flush()?;
        let stale = matches!(
            document.paragraph_text(TARGET_INDEX),
            Err(litchi_core::Error::SourceChanged { .. })
        );
        if !stale {
            return Err("facade stale-source oracle failed".into());
        }
        drop(document);
        Ok(FacadeStatus {
            selected_text: selected.ok_or("facade selected paragraph was absent")?,
            out_of_bounds_none: out_of_bounds,
            stale_source_changed: stale,
        })
    })();
    drop(owner);
    result
}

fn managed_context(
    archive_bytes: usize,
) -> AnyResult<(Budget, CancellationSource, ExecutionContext)> {
    let memory = u64::try_from(archive_bytes)
        .ok()
        .and_then(|value| value.checked_mul(16))
        .unwrap_or(64 * 1024 * 1024)
        .max(1024 * 1024);
    let budget = Budget::root(
        "docx-source-selected-paragraph",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).ok_or("invalid worker limit")?,
        NonZeroUsize::new(1).ok_or("invalid in-flight task limit")?,
        NonZeroU64::new(memory).ok_or("invalid in-flight byte limit")?,
        0,
    )?;
    Ok((
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, limits),
    ))
}

fn run_source(corpus: &Corpus, managed: bool) -> AnyResult<SourceRun> {
    let source = Arc::new(CountingSource::new(
        Arc::clone(&corpus.archive),
        Arc::clone(&corpus.ranges),
    ));
    let source_for_package: Arc<dyn ReadAt> = source.clone();
    let mut budget = None;
    let mut cancellation_source = None;
    let memory_before_open;
    let open_before = source.snapshot()?;
    let package = if managed {
        let (managed_budget, managed_cancellation, context) =
            managed_context(corpus.archive.len())?;
        memory_before_open = managed_budget.used(Resource::Memory);
        let package = source_backed::Package::from_read_at_with_execution_context(
            source_for_package,
            litchi_docx::ReadLimits::default(),
            context,
        )?;
        budget = Some(managed_budget);
        cancellation_source = Some(managed_cancellation);
        package
    } else {
        memory_before_open = 0;
        source_backed::Package::from_read_at(source_for_package)?
    };

    let open_delta = source.delta(open_before)?;
    let open_cache_raw = package.cache_diagnostics();
    require_no_semantic_payload_overlap(
        &open_delta,
        "open_without_semantic_payload_member_reads",
        "main_document",
    )?;
    require_no_semantic_payload_overlap(
        &open_delta,
        "open_without_semantic_payload_member_reads",
        "media",
    )?;
    require_no_semantic_payload_overlap(
        &open_delta,
        "open_without_semantic_payload_member_reads",
        "core",
    )?;
    require_no_semantic_payload_overlap(
        &open_delta,
        "open_without_semantic_payload_member_reads",
        "unrelated",
    )?;
    if open_delta.read_calls == 0 {
        return Err("open_without_semantic_payload_member_reads produced no source reads".into());
    }

    let version_before_raw = package.source_version()?;
    let document_before = source.snapshot()?;
    let document_cache_before = package.cache_diagnostics();
    let document = package.document()?;
    let _document_after = source.snapshot()?;
    let document_cache_after = package.cache_diagnostics();
    let document_cache_after_stats = CacheStats::from_diagnostics(document_cache_after);
    let document_delta = source.delta(document_before)?;
    let document_cache_delta =
        CacheDelta::from_snapshots(document_cache_before, package.cache_diagnostics())?;
    require_no_semantic_payload_overlap(
        &document_delta,
        "document_materialization_main_document_only_semantic_payload",
        "media",
    )?;
    require_no_semantic_payload_overlap(
        &document_delta,
        "document_materialization_main_document_only_semantic_payload",
        "core",
    )?;
    require_no_semantic_payload_overlap(
        &document_delta,
        "document_materialization_main_document_only_semantic_payload",
        "unrelated",
    )?;
    require_no_semantic_payload_overlap(
        &document_delta,
        "document_materialization_main_document_only_semantic_payload",
        "catalog_relationship",
    )?;
    if class_bytes(&document_delta, "main_document") == 0 || document_cache_delta.cold_loads != 1 {
        return Err("document materialization was not exactly one main payload load".into());
    }
    let memory_after_document = budget.as_ref().map(|value| value.used(Resource::Memory));

    let paragraph_before = source.snapshot()?;
    let paragraph_cache_before = package.cache_diagnostics();
    let selected = document.paragraph_text(TARGET_INDEX)?;
    let paragraph_cache_after = package.cache_diagnostics();
    let paragraph_phase = PhaseEvidence {
        source: source.delta(paragraph_before)?,
        cache: CacheDelta::from_snapshots(paragraph_cache_before, paragraph_cache_after)?,
    };
    require_zero_query(&paragraph_phase)?;
    if selected.as_deref() != Some(TARGET_TEXT) {
        return Err("source selected paragraph oracle failed".into());
    }

    let oob_before = source.snapshot()?;
    let oob_cache_before = package.cache_diagnostics();
    let out_of_bounds = document.paragraph_text(usize::MAX)?;
    let oob_phase = PhaseEvidence {
        source: source.delta(oob_before)?,
        cache: CacheDelta::from_snapshots(oob_cache_before, package.cache_diagnostics())?,
    };
    require_zero_query(&oob_phase)?;
    if out_of_bounds.is_some() {
        return Err("source out-of-bounds paragraph oracle failed".into());
    }

    let version_after_queries_raw = package.source_version()?;
    if version_before_raw != version_after_queries_raw {
        return Err("source version changed during semantic queries".into());
    }

    let cache_after_queries = CacheStats::from_diagnostics(package.cache_diagnostics());
    source.bump_revision();
    let stale_before = source.snapshot()?;
    let stale_cache_before = package.cache_diagnostics();
    let held_result = document.paragraph_text(TARGET_INDEX);
    let stale_source = source.delta(stale_before)?;
    let stale_cache = CacheDelta::from_snapshots(stale_cache_before, package.cache_diagnostics())?;
    let snapshot_stable_selected_text = matches!(
        &held_result,
        Ok(Some(value)) if value == TARGET_TEXT
    );
    let snapshot_query_no_source_work = stale_source.read_calls == 0
        && stale_source.requested_bytes == 0
        && stale_source.returned_bytes == 0
        && stale_source.len_calls == 0
        && stale_source.version_calls == 0
        && stale_source.range_records == 0
        && stale_cache.is_zero();
    if !snapshot_stable_selected_text || !snapshot_query_no_source_work {
        return Err(
            "held source-backed document snapshot was not stable without incremental source work"
                .into(),
        );
    }
    let package_reentry_typed_source_changed = match package.document() {
        Err(error) if typed_source_changed(&error) => true,
        Err(error) => return Err(format!("stale source returned the wrong error: {error}").into()),
        Ok(_) => false,
    };
    if !package_reentry_typed_source_changed {
        return Err("package re-entry did not return typed SourceChanged".into());
    }
    let stale_phase = StaleEvidence {
        snapshot_stable_selected_text,
        snapshot_query_no_source_work,
        package_reentry_typed_source_changed,
        snapshot_semantics: "immutable document snapshot intentionally remains stable after source revision",
        observed_revision: source.current_version().revision(),
        source: stale_source,
        cache: stale_cache,
    };

    drop(document);
    drop(package);
    drop(cancellation_source);
    let memory_after_drop = budget
        .as_ref()
        .map_or(0, |value| value.used(Resource::Memory));
    if managed && memory_after_drop != 0 {
        return Err(format!(
            "managed source budget did not release after drop: {memory_after_drop}"
        )
        .into());
    }

    let open_phase = PhaseEvidence {
        source: open_delta,
        cache: CacheDelta::from_snapshots(SourceCacheDiagnostics::default(), open_cache_raw)?,
    };
    Ok(SourceRun {
        mode: if managed { "managed" } else { "unmanaged" },
        selected_text: selected.ok_or("source selected paragraph was absent")?,
        out_of_bounds_none: out_of_bounds.is_none(),
        source_version_before: version_report(version_before_raw),
        source_version_after_queries: version_report(version_after_queries_raw),
        source_version_unchanged_before_stale: version_before_raw == version_after_queries_raw,
        open_without_semantic_payload_member_reads: open_phase,
        document_materialization_main_document_only_semantic_payload: PhaseEvidence {
            source: document_delta,
            cache: document_cache_delta,
        },
        paragraph_query: paragraph_phase,
        out_of_bounds_query: oob_phase,
        stale_refusal: stale_phase,
        cache_after_document_materialization: document_cache_after_stats,
        cache_after_queries,
        budget: budget.map(|_| BudgetEvidence {
            memory_before_open,
            memory_after_document: memory_after_document.unwrap_or(0),
            memory_after_drop,
            released_after_drop: memory_after_drop == 0,
        }),
    })
}

fn provenance() -> AnyResult<Provenance> {
    let executable = env::current_exe()?;
    let executable_identity = hash_file_bounded(&executable, MAX_EXECUTABLE_HASH_BYTES)?;
    let cargo_lock = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../Cargo.lock");
    let cargo_lock_identity = hash_file_bounded(&cargo_lock, MAX_CARGO_LOCK_HASH_BYTES).ok();
    let repository = repository_root();
    let rustc_vv = command_text("rustc", &["-Vv"]);
    let rustc_vv_sha256 = rustc_vv
        .as_deref()
        .filter(|value| !value.is_empty())
        .map(|value| sha256_hex(value.as_bytes()));
    let mut environment = BTreeMap::new();
    for key in [
        "RUSTFLAGS",
        "RUSTC_WRAPPER",
        "CARGO_PROFILE_RELEASE_DEBUG",
        "LITCHI_PERF_PROFILE",
        "LITCHI_PERF_FEATURES",
        "LITCHI_PERF_BUILD_COMMAND",
    ] {
        if let Ok(value) = env::var(key) {
            environment.insert(key.to_owned(), value);
        }
    }
    Ok(Provenance {
        executable_path: executable.display().to_string(),
        executable_bytes: executable_identity.bytes,
        executable_sha256: executable_identity.sha256,
        git_revision: command_text_in_dir("git", &["rev-parse", "HEAD"], Some(&repository)),
        git_dirty: command_text_in_dir(
            "git",
            &["status", "--porcelain=v1", "--untracked-files=all"],
            Some(&repository),
        )
        .map(|value| git_status_is_dirty(&value)),
        rustc_vv,
        rustc_vv_sha256,
        executable_identity_policy: EXECUTABLE_IDENTITY_POLICY,
        cargo_lock_bytes: cargo_lock_identity.as_ref().map(|value| value.bytes),
        cargo_lock_sha256: cargo_lock_identity.map(|value| value.sha256),
        profile: EXPECTED_PROFILE.map(|value| value.to_owned()),
        features: EXPECTED_FEATURES.map(|value| value.to_owned()),
        build_command: EXPECTED_BUILD_COMMAND.map(|value| value.to_owned()),
        os: env::consts::OS.to_owned(),
        arch: env::consts::ARCH.to_owned(),
        cpu: cpu_model(),
        memory: memory_total(),
        environment,
    })
}

fn repository_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..")
}

fn git_status_is_dirty(status: &str) -> bool {
    status.lines().any(|line| !line.trim().is_empty())
}

#[derive(Debug, Serialize)]
struct Provenance {
    executable_path: String,
    executable_bytes: usize,
    executable_sha256: String,
    git_revision: Option<String>,
    git_dirty: Option<bool>,
    rustc_vv: Option<String>,
    rustc_vv_sha256: Option<String>,
    executable_identity_policy: &'static str,
    cargo_lock_bytes: Option<usize>,
    cargo_lock_sha256: Option<String>,
    profile: Option<String>,
    features: Option<String>,
    build_command: Option<String>,
    os: String,
    arch: String,
    cpu: Option<String>,
    memory: Option<String>,
    environment: BTreeMap<String, String>,
}

struct CappedCommandOutput {
    bytes: Vec<u8>,
    truncated: bool,
}

fn read_capped<R: Read>(mut reader: R) -> io::Result<CappedCommandOutput> {
    let mut bytes = Vec::with_capacity(MAX_COMMAND_OUTPUT_BYTES.min(4096));
    let mut buffer = [0_u8; 8192];
    let mut truncated = false;
    loop {
        let count = reader.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        let remaining = MAX_COMMAND_OUTPUT_BYTES.saturating_sub(bytes.len());
        let keep = remaining.min(count);
        bytes.extend_from_slice(&buffer[..keep]);
        if keep != count {
            truncated = true;
        }
    }
    Ok(CappedCommandOutput { bytes, truncated })
}

fn command_text(program: &str, arguments: &[&str]) -> Option<String> {
    command_text_in_dir(program, arguments, None)
}

fn command_text_in_dir(
    program: &str,
    arguments: &[&str],
    directory: Option<&Path>,
) -> Option<String> {
    let mut command = Command::new(program);
    command.args(arguments);
    if let Some(directory) = directory {
        command.current_dir(directory);
    }
    let mut child = command
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .ok()?;
    let stdout = child.stdout.take()?;
    let stderr = child.stderr.take()?;
    let (stdout, stderr) = std::thread::scope(|scope| {
        let stdout_handle = scope.spawn(move || read_capped(stdout));
        let stderr_handle = scope.spawn(move || read_capped(stderr));
        let stdout = stdout_handle.join().ok()?.ok()?;
        let stderr = stderr_handle.join().ok()?.ok()?;
        Some((stdout, stderr))
    })?;
    if !child.wait().ok()?.success() || stdout.truncated || stderr.truncated {
        return None;
    }
    let text = String::from_utf8(stdout.bytes).ok()?.trim().to_owned();
    Some(text)
}

fn cpu_model() -> Option<String> {
    let text = fs::read_to_string("/proc/cpuinfo").ok()?;
    text.lines()
        .find_map(|line| line.strip_prefix("model name\t: ").map(str::to_owned))
}

fn memory_total() -> Option<String> {
    let text = fs::read_to_string("/proc/meminfo").ok()?;
    text.lines().find_map(|line| {
        line.strip_prefix("MemTotal:")
            .map(str::trim)
            .map(str::to_owned)
    })
}

fn exact_pin_matches(actual: Option<&str>, expected: Option<&str>) -> bool {
    match (actual, expected) {
        (Some(actual), Some(expected)) => actual == expected,
        _ => false,
    }
}

fn authorize_provenance(value: &Provenance) -> AnyResult<()> {
    if value.executable_bytes == 0 || value.executable_sha256.len() != 64 {
        return Err("authoritative executable identity is unavailable".into());
    }
    let expected_revision = EXPECTED_GIT_REVISION
        .ok_or("authoritative git revision is unpinned; set LITCHI_DOCX_SELECTED_GIT_REVISION")?;
    if expected_revision.len() != 40 {
        return Err("authoritative git revision pin must be a full 40-character SHA".into());
    }
    if value.git_revision.as_deref() != Some(expected_revision) {
        return Err("authoritative git revision does not match the pinned revision".into());
    }
    if value.git_dirty != Some(false) {
        return Err("authoritative evidence requires clean tracked and untracked status".into());
    }
    let expected_lock = EXPECTED_CARGO_LOCK_SHA256.ok_or(
        "authoritative Cargo.lock hash is unpinned; set LITCHI_DOCX_SELECTED_CARGO_LOCK_SHA256",
    )?;
    if expected_lock.len() != 64 {
        return Err("authoritative Cargo.lock pin must be a 64-character SHA-256".into());
    }
    if value.cargo_lock_sha256.as_deref() != Some(expected_lock) {
        return Err("authoritative Cargo.lock hash does not match the pinned hash".into());
    }
    let expected_profile = EXPECTED_PROFILE
        .ok_or("authoritative profile is unpinned; set LITCHI_DOCX_SELECTED_PROFILE")?;
    if expected_profile != "release" {
        return Err("authoritative profile pin must be exactly release".into());
    }
    if value.profile.as_deref() != Some(expected_profile) {
        return Err("authoritative build profile does not match the pinned profile".into());
    }
    let expected_features = EXPECTED_FEATURES
        .ok_or("authoritative features are unpinned; set LITCHI_DOCX_SELECTED_FEATURES")?;
    if expected_features.is_empty() {
        return Err("authoritative enabled-feature descriptor must be non-empty".into());
    }
    if value.features.as_deref() != Some(expected_features) {
        return Err("authoritative features do not match the pinned features".into());
    }
    let expected_command = EXPECTED_BUILD_COMMAND
        .ok_or("authoritative build command is unpinned; set LITCHI_DOCX_SELECTED_BUILD_COMMAND")?;
    if expected_command.is_empty() {
        return Err("authoritative build command pin must be non-empty".into());
    }
    if value.build_command.as_deref() != Some(expected_command) {
        return Err("authoritative build command does not match the pinned command".into());
    }
    if !matches!(value.rustc_vv.as_deref(), Some(text) if !text.is_empty()) {
        return Err("authoritative rustc -Vv output is unavailable or empty".into());
    }
    let expected_rustc_vv_sha256 = EXPECTED_RUSTC_VV_SHA256
        .ok_or("authoritative rustc hash is unpinned; set LITCHI_DOCX_SELECTED_RUSTC_VV_SHA256")?;
    if expected_rustc_vv_sha256.len() != 64 {
        return Err("authoritative rustc pin must be a 64-character SHA-256".into());
    }
    if !exact_pin_matches(
        value.rustc_vv_sha256.as_deref(),
        Some(expected_rustc_vv_sha256),
    ) {
        return Err("authoritative rustc -Vv hash does not match the pinned hash".into());
    }
    Ok(())
}

#[derive(Debug, Serialize)]
struct DiscoveryReport {
    schema: &'static str,
    run_mode: &'static str,
    claim_authorized: bool,
    claim_scope: &'static str,
    performance_claim: &'static str,
    sha256_sidecar_path: Option<String>,
    corpus: CorpusReport,
    provenance: Provenance,
    expected_identity_placeholders: ExpectedIdentityPlaceholders,
}

#[derive(Debug, Serialize)]
struct OutputEnvelope {
    bytes: usize,
    sha256: String,
}

fn atomic_write(path: &Path, bytes: &[u8]) -> AnyResult<()> {
    if bytes.len() > MAX_OUTPUT_BYTES {
        return Err("runner output exceeded the bounded write limit".into());
    }
    let parent = path.parent().filter(|value| !value.as_os_str().is_empty());
    if let Some(parent) = parent {
        fs::create_dir_all(parent)?;
    }
    let file_name = path
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("output");
    for attempt in 0..64_u64 {
        let serial = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
        let temporary = path.with_file_name(format!(
            ".{file_name}.{}.{}.tmp",
            std::process::id(),
            serial + attempt
        ));
        let mut file = match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
        {
            Ok(file) => file,
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => return Err(error.into()),
        };
        let owner = OwnedTempFile {
            path: temporary.clone(),
        };
        let result = file
            .write_all(bytes)
            .and_then(|_| file.flush())
            .and_then(|_| file.sync_all());
        drop(file);
        if let Err(error) = result {
            drop(owner);
            return Err(error.into());
        }
        if let Err(error) = fs::rename(&temporary, path) {
            drop(owner);
            return Err(error.into());
        }
        drop(owner);
        return Ok(());
    }
    Err("could not allocate an output temporary path after bounded retries".into())
}

fn sha256_sidecar_path(path: &Path) -> PathBuf {
    PathBuf::from(format!("{}.sha256.json", path.display()))
}

fn reported_sidecar_path(output: Option<&Path>) -> Option<String> {
    output.map(|path| sha256_sidecar_path(path).display().to_string())
}

fn write_serialized(output: Option<PathBuf>, serialized: &[u8]) -> AnyResult<()> {
    if serialized.len() > MAX_OUTPUT_BYTES {
        return Err("evidence JSON exceeded the runner output bound".into());
    }
    if let Some(path) = output {
        atomic_write(&path, serialized)?;
        let envelope = OutputEnvelope {
            bytes: serialized.len(),
            sha256: sha256_hex(serialized),
        };
        let envelope_bytes = serde_json::to_vec(&envelope)?;
        let sidecar = sha256_sidecar_path(&path);
        atomic_write(&sidecar, &envelope_bytes)?;
    } else {
        io::stdout().write_all(serialized)?;
        io::stdout().write_all(b"\n")?;
    }
    Ok(())
}

fn parse_output_path() -> AnyResult<(Option<PathBuf>, RunMode)> {
    let mut arguments = env::args().skip(1);
    let mut output = None;
    let mut mode = RunMode::Authoritative;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--out" => {
                let value = arguments.next().ok_or("--out requires a path")?;
                output = Some(PathBuf::from(value));
            },
            "--discover-corpus" => {
                if mode != RunMode::Authoritative {
                    return Err("--discover-corpus and --smoke are mutually exclusive".into());
                }
                mode = RunMode::DiscoverCorpus;
            },
            "--smoke" => {
                if mode != RunMode::Authoritative {
                    return Err("--discover-corpus and --smoke are mutually exclusive".into());
                }
                mode = RunMode::Smoke;
            },
            "--help" | "-h" => {
                println!(
                    "usage: docx_source_selected_paragraph [--out PATH] [--discover-corpus|--smoke]"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument: {other}").into()),
        }
    }
    Ok((output, mode))
}

fn run(output: Option<PathBuf>, mode: RunMode) -> AnyResult<()> {
    let corpus = build_corpus()?;
    let provenance = provenance()?;
    if mode == RunMode::DiscoverCorpus {
        let report = DiscoveryReport {
            schema: SCHEMA,
            run_mode: mode.as_str(),
            claim_authorized: false,
            claim_scope: CLAIM_SCOPE,
            performance_claim: "none",
            sha256_sidecar_path: reported_sidecar_path(output.as_deref()),
            corpus: corpus_report(&corpus),
            provenance,
            expected_identity_placeholders: expected_identity_placeholders(),
        };
        let serialized = serde_json::to_vec_pretty(&report)?;
        return write_serialized(output, &serialized);
    }
    if mode == RunMode::Authoritative {
        verify_corpus_identity(&corpus)?;
        authorize_provenance(&provenance)?;
    }
    let eager = run_eager_oracle(&corpus)?;
    let facade = run_facade_oracle(&corpus)?;
    let unmanaged = run_source(&corpus, false)?;
    let managed = run_source(&corpus, true)?;
    if unmanaged.selected_text != TARGET_TEXT
        || managed.selected_text != TARGET_TEXT
        || !unmanaged.out_of_bounds_none
        || !managed.out_of_bounds_none
        || !unmanaged.stale_refusal.package_reentry_typed_source_changed
        || !managed.stale_refusal.package_reentry_typed_source_changed
    {
        return Err("one or more source-backed oracle statuses failed".into());
    }
    let mut unavailable = BTreeMap::new();
    unavailable.insert("physical_io", "not_available: logical ReadAt only");
    unavailable.insert(
        "elapsed_time",
        "not_available: this runner has no timing path",
    );
    unavailable.insert(
        "rss",
        "not_available: this runner has no process-memory sampler",
    );
    unavailable.insert(
        "allocation",
        "not_available: this runner has no allocator sampler",
    );
    unavailable.insert(
        "general_docx_claims",
        "not_available: this runner covers only one DOCX selected paragraph",
    );
    let evidence = Evidence {
        schema: SCHEMA,
        run_mode: mode.as_str(),
        claim_authorized: mode.claim_authorized(),
        claim_scope: CLAIM_SCOPE,
        performance_claim: "none",
        sha256_sidecar_path: reported_sidecar_path(output.as_deref()),
        corpus: corpus_report(&corpus),
        provenance,
        target: TargetReport {
            direct_body_index: TARGET_INDEX,
            expected_paragraph_count: EXPECTED_PARAGRAPH_COUNT,
            expected_text: TARGET_TEXT,
        },
        oracles: OracleReport { eager, facade },
        source_runs: [unmanaged, managed],
        unavailable,
    };
    let serialized = serde_json::to_vec_pretty(&evidence)?;
    write_serialized(output, &serialized)
}

fn main() {
    match parse_output_path().and_then(|(output, mode)| run(output, mode)) {
        Ok(()) => {},
        Err(error) => {
            eprintln!("docx source selected paragraph evidence failed: {error}");
            std::process::exit(1);
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn overlap_is_bounded_and_empty_at_edges() {
        assert_eq!(overlap(0..10, 10..20), 10..10);
        assert_eq!(overlap(2..8, 4..6), 4..6);
        assert_eq!(overlap(4..6, 2..8), 4..6);
    }

    #[test]
    fn checked_read_delta_rejects_counter_regression() {
        let before = ReadSnapshot {
            read_calls: 2,
            requested_bytes: 4,
            returned_bytes: 4,
            len_calls: 1,
            version_calls: 1,
            range_records: 1,
        };
        let after = ReadSnapshot {
            read_calls: 1,
            ..before
        };
        assert!(after.read_calls.checked_sub(before.read_calls).is_none());
    }

    #[test]
    fn schema_and_scope_are_statistics_free() {
        assert_eq!(SCHEMA, "litchi.docx.source-selected-paragraph.v1");
        assert_eq!(
            CORPUS_GENERATOR,
            "litchi-docx-source-selected-paragraph-media-v2"
        );
        assert_eq!(EXPECTED_PARAGRAPH_COUNT, 201);
        assert_eq!(TARGET_INDEX, 100);
        assert_eq!(TARGET_TEXT, "source-selected-paragraph-0100");
        assert_eq!(EXPECTED_ARCHIVE_BYTES, Some(16_786_572));
        assert_eq!(
            EXPECTED_ARCHIVE_SHA256,
            Some("a4384c2c249ef87bac6150f92b1a839d4555872f5c9b6ffe6b3d849f47bb7fef")
        );
        assert_eq!(EXPECTED_MEMBER_IDENTITIES.len(), 15);
        assert_eq!(
            CLAIM_SCOPE,
            "selected paragraph access after one main-document materialization; logical ReadAt/cache/managed-budget evidence only"
        );
    }

    #[test]
    fn deterministic_corpus_shape_is_pinned_without_hashes() {
        assert_eq!(MEDIA_COUNT, 8);
        assert_eq!(MEDIA_PAYLOAD_BYTES, 2 * 1024 * 1024);
        assert!(MIN_ARCHIVE_BYTES < MAX_ARCHIVE_BYTES);
        assert_eq!(expected_paragraph_text(0), "source-selected-paragraph-0000");
        assert_eq!(expected_paragraph_text(TARGET_INDEX), TARGET_TEXT);
        assert_eq!(
            expected_paragraph_text(EXPECTED_PARAGRAPH_COUNT - 1),
            "source-selected-paragraph-0200"
        );
        let xml = main_xml();
        assert_eq!(
            xml.windows(b"<w:p>".len())
                .filter(|window| (*window).eq(b"<w:p>"))
                .count(),
            EXPECTED_PARAGRAPH_COUNT
        );
        assert_ne!(expected_paragraph_text(0), expected_paragraph_text(1));
        assert_ne!(media_payload(0), media_payload(1));
        assert_eq!(media_payload(0).len(), MEDIA_PAYLOAD_BYTES);
        assert_eq!(media_payload(MEDIA_COUNT - 1).len(), MEDIA_PAYLOAD_BYTES);
    }

    #[test]
    fn pinned_identity_rejects_archive_mismatch() {
        let corpus = Corpus {
            archive: Arc::<[u8]>::from(&b"bad"[..]),
            archive_sha256: "bad".to_owned(),
            members: BTreeMap::new(),
            ranges: Arc::new(Vec::new()),
        };
        assert!(verify_corpus_identity(&corpus).is_err());
    }

    #[test]
    fn git_status_parser_reports_actual_dirtiness() {
        assert!(!git_status_is_dirty(""));
        assert!(!git_status_is_dirty("\n  \n"));
        assert!(git_status_is_dirty(
            "?? tools/perf-baseline/src/bin/runner.rs\n"
        ));
    }

    #[test]
    fn byte_accounting_reconciles_structural_residual_ranges() {
        let mut class_requested_bytes = BTreeMap::new();
        class_requested_bytes.insert("main_document".to_owned(), 3);
        class_requested_bytes.insert("zip_structural".to_owned(), 7);
        let mut class_returned_bytes = BTreeMap::new();
        class_returned_bytes.insert("main_document".to_owned(), 4);
        class_returned_bytes.insert("zip_structural".to_owned(), 4);
        let mut delta = ReadDelta {
            read_calls: 1,
            requested_bytes: 10,
            returned_bytes: 8,
            len_calls: 0,
            version_calls: 0,
            range_records: 1,
            zip_structural_requested_bytes: 7,
            zip_structural_returned_bytes: 4,
            unclassified_requested_bytes: 0,
            unclassified_returned_bytes: 0,
            class_requested_bytes,
            class_returned_bytes,
            member_reads: BTreeMap::new(),
        };
        assert!(delta.reconciles());
        delta.unclassified_requested_bytes = 1;
        assert!(!delta.reconciles());
    }

    #[test]
    fn bounded_streaming_hash_counts_and_enforces_limits() {
        let owner = create_temp_corpus(b"streamed").expect("temporary corpus");
        let identity = hash_file_bounded(&owner.path, 64).expect("streaming hash");
        assert_eq!(identity.bytes, 8);
        assert_eq!(identity.sha256, sha256_hex(b"streamed"));
        assert!(hash_file_bounded(&owner.path, 3).is_err());
    }

    #[test]
    fn toolchain_pin_mismatch_is_not_authorized() {
        assert!(exact_pin_matches(Some("abc"), Some("abc")));
        assert!(!exact_pin_matches(Some("abc"), Some("def")));
        assert!(!exact_pin_matches(None, Some("abc")));
    }

    #[test]
    fn non_authoritative_modes_cannot_authorize_claims() {
        assert!(!RunMode::DiscoverCorpus.claim_authorized());
        assert!(!RunMode::Smoke.claim_authorized());
        assert!(RunMode::Authoritative.claim_authorized());
    }

    #[test]
    fn owned_temp_file_is_removed_only_after_successful_creation() {
        let owner = create_temp_corpus(b"owned temporary corpus").expect("temporary corpus");
        let path = owner.path.clone();
        assert!(path.exists());
        drop(owner);
        assert!(!path.exists());
    }
}
