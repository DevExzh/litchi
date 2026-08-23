//! Controlled filesystem evidence for the OPC and OLE2 CRUD paths.
//!
//! Every measured operation runs in a fresh child process. This scopes procfs
//! counters and parser state to one sample. The parent records child operation
//! time separately from parent-observed process wall time.

use std::{
    env,
    error::Error,
    fs::{self, File, OpenOptions},
    io::{self, Write},
    ops::Range,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, Ordering},
    },
    time::{Instant, SystemTime, UNIX_EPOCH},
};

use litchi_cfb::{OleFile, OverlayLimits, SameLengthStreamOverlay, SharedOleFile};
use litchi_core::{FileSource, ReadAt, SourceVersion};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceBackedPackage};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use soapberry_zip::ZipArchive;

use crate::{cold_verified, process_metrics};

const OPC_FILE_SHAPE: super::CorpusShape = super::CorpusShape::FewLarge;
const OPC_FILE_PAYLOAD: super::PayloadKind = super::PayloadKind::Incompressible;
const OPC_FILE_ENTRY_BYTES: usize = 4 * 1024 * 1024;
const OPC_FILE_TARGET_INDEX: usize = 2;
const OPC_FILE_SOURCE_SHA256: &str =
    "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6";
const OPC_FILE_EXPECTED_OUTPUT_SHA256: &str =
    "f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009";
const CFB_FILE_SOURCE_SHA256: &str =
    "7ffbd37c3e472a21b382bcbb02e430a62164e58d2270bbee0deaa584ff47a94d";
const CFB_FILE_EXPECTED_OUTPUT_SHA256: &str =
    "7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc";
const CFB_FILE_OWNED_SOURCE_VERSION_ID: u64 = 0xcfb0_0184;
const FILESYSTEM_OLE_COMMON_REPLACEMENT: &[u8] = b"litchi-ole-common-modified-stream-v1";
const PPTX_FILE_SELECTED_POSITION: usize = super::PPTX_SOURCE_SLIDE_COUNT / 2;
const PPTX_FILE_CORPUS_GENERATOR: &str = super::PPTX_SOURCE_EDIT_CORPUS_GENERATOR;
const DOCX_FILE_CORPUS_GENERATOR: &str = super::DOCX_SOURCE_EDIT_CORPUS_GENERATOR;
const DOCX_FILE_SOURCE_SHA256: &str =
    "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4";
// The unified XLSX file selectors use one fixed, media-rich cell-CRUD input.
// Keep these literals independent of the builder's computed manifest so a
// generator or archive-shape drift cannot silently redefine the evidence.
const XLSX_FILE_CORPUS_GENERATOR: &str = "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1";
const XLSX_FILE_SOURCE_SHAPE: &str = "medium";
const XLSX_FILE_SOURCE_SHA256: &str =
    "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036";
const XLSX_FILE_SOURCE_ARCHIVE_BYTES: usize = 4_226_429;
const XLSX_FILE_SOURCE_ENTRY_COUNT: usize = 9_216;
const XLSX_FILE_SOURCE_ARCHIVE_MEMBER_COUNT: usize = 17;
const XLSX_FILE_SOURCE_UNCOMPRESSED_PAYLOAD_BYTES: usize = 4_231_168;
const XLSX_FILE_SOURCE_SHEET_COUNT: usize = 4;
const XLSX_FILE_SOURCE_ROWS_PER_SHEET: usize = 48;
const XLSX_FILE_SOURCE_COLUMNS_PER_SHEET: usize = 48;
static NEXT_PPTX_REPLAY_SOURCE_ID: AtomicU64 = AtomicU64::new(1);
static NEXT_DOCX_REPLAY_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

/// Logical order observed by a positional source wrapper.
///
/// `sequential` means every completed non-empty range after the first began
/// exactly where the previous completed range ended. `random` means at least
/// one observed transition was non-contiguous. Both labels describe the
/// caller's logical `ReadAt` ranges only; they do not describe kernel, device,
/// filesystem, or remote physical I/O. A concurrent, empty, short, or
/// otherwise insufficient observation is `unknown`; invalid range arithmetic
/// fails closed.
#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub(crate) enum ReadPattern {
    Sequential,
    Random,
    Unknown,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct CacheSelection {
    warm: bool,
    cold_requested: bool,
    cold_verified: bool,
}

#[derive(Clone, Debug, Default, Serialize)]
pub(crate) struct HostEvidence {
    pub os: Option<String>,
    pub kernel: Option<String>,
    pub cpu_model: Option<String>,
    pub total_memory_bytes: Option<u64>,
    pub page_size_bytes: Option<u64>,
    pub filesystem_type: Option<String>,
    pub source_destination_same_device: Option<bool>,
    pub cpu_affinity: Option<String>,
    pub storage_identifier: Option<String>,
}

/// Best-effort host facts for interpreting controlled filesystem samples.
/// Absolute paths and device identifiers are intentionally never serialized.
pub(crate) fn host_evidence(
    requested_root: Option<&Path>,
    filesystem_measured: bool,
) -> HostEvidence {
    let probe_root = filesystem_measured.then(|| {
        requested_root
            .map(Path::to_path_buf)
            .unwrap_or_else(env::temp_dir)
    });
    HostEvidence {
        os: Some(env::consts::OS.to_owned()),
        kernel: command_output("uname", &["-sr"]),
        cpu_model: proc_cpu_model(),
        total_memory_bytes: proc_memory_bytes(),
        page_size_bytes: command_output("getconf", &["PAGESIZE"])
            .and_then(|value| value.parse::<u64>().ok()),
        filesystem_type: probe_root
            .as_deref()
            .and_then(|path| command_output("stat", &["-f", "-c", "%T", &path.to_string_lossy()])),
        source_destination_same_device: probe_root.as_deref().and_then(same_device_probe),
        cpu_affinity: proc_status_value("Cpus_allowed_list"),
        storage_identifier: None,
    }
}

fn command_output(program: &str, arguments: &[&str]) -> Option<String> {
    let output = Command::new(program).args(arguments).output().ok()?;
    if !output.status.success() {
        return None;
    }
    String::from_utf8(output.stdout)
        .ok()
        .map(|value| value.trim().to_owned())
        .filter(|value| !value.is_empty())
}

fn proc_status_value(key: &str) -> Option<String> {
    fs::read_to_string("/proc/self/status")
        .ok()?
        .lines()
        .find_map(|line| {
            let (name, value) = line.split_once(':')?;
            (name.trim() == key).then(|| value.trim().to_owned())
        })
}

fn proc_cpu_model() -> Option<String> {
    let text = fs::read_to_string("/proc/cpuinfo").ok()?;
    text.lines().find_map(|line| {
        let (name, value) = line.split_once(':')?;
        matches!(name.trim(), "model name" | "Hardware")
            .then(|| value.trim().to_owned())
            .filter(|value| !value.is_empty())
    })
}

fn proc_memory_bytes() -> Option<u64> {
    let text = fs::read_to_string("/proc/meminfo").ok()?;
    let kib = text.lines().find_map(|line| {
        let (name, value) = line.split_once(':')?;
        (name.trim() == "MemTotal")
            .then(|| value.split_whitespace().next()?.parse::<u64>().ok())
            .flatten()
    })?;
    kib.checked_mul(1024)
}

#[cfg(unix)]
fn same_device_probe(root: &Path) -> Option<bool> {
    use std::os::unix::fs::MetadataExt;

    let root_device = fs::metadata(root).ok()?.dev();
    let probe = root.join(format!(
        ".litchi-perf-device-probe-{}-{}",
        std::process::id(),
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .ok()?
            .as_nanos()
    ));
    let file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&probe)
        .ok()?;
    let probe_device = file.metadata().ok().map(|metadata| metadata.dev());
    drop(file);
    let _ = fs::remove_file(probe);
    probe_device.map(|device| device == root_device)
}

#[cfg(not(unix))]
fn same_device_probe(_root: &Path) -> Option<bool> {
    None
}

impl Default for CacheSelection {
    fn default() -> Self {
        Self {
            warm: true,
            cold_requested: true,
            cold_verified: false,
        }
    }
}

impl CacheSelection {
    pub(crate) fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        if value.is_empty() {
            return Err("--filesystem-cache selection must not be empty".into());
        }
        let mut selection = Self {
            warm: false,
            cold_requested: false,
            cold_verified: false,
        };
        for state in value.split(',') {
            match state {
                "warm" => selection.warm = true,
                "cold-requested" => selection.cold_requested = true,
                "cold-verified" => selection.cold_verified = true,
                _ => {
                    return Err(format!(
                        "invalid --filesystem-cache state {state:?}; expected warm, cold-requested, or cold-verified"
                    )
                    .into());
                },
            }
        }
        if !selection.warm && !selection.cold_requested && !selection.cold_verified {
            return Err(
                "--filesystem-cache selection must include warm, cold-requested, or cold-verified"
                    .into(),
            );
        }
        Ok(selection)
    }

    pub(crate) const fn warm(self) -> bool {
        self.warm
    }

    pub(crate) const fn cold_requested(self) -> bool {
        self.cold_requested
    }

    pub(crate) const fn cold_verified(self) -> bool {
        self.cold_verified
    }

    pub(crate) fn names(self) -> Vec<&'static str> {
        let mut names = Vec::with_capacity(3);
        if self.warm {
            names.push("warm");
        }
        if self.cold_requested {
            names.push("cold-requested");
        }
        if self.cold_verified {
            names.push("cold-verified");
        }
        names
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct ChildResult {
    elapsed_ns: u64,
    logical_read_counter_scope: String,
    logical_read_calls: u64,
    logical_read_requested_bytes: u64,
    logical_read_bytes: u64,
    logical_read_largest_requested_bytes: u64,
    logical_read_largest_returned_bytes: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    logical_read_pattern: Option<ReadPattern>,
    max_concurrent_reads: u64,
    logical_read_request_sizes: Vec<u64>,
    logical_read_request_size_buckets: ReadSizeBuckets,
    cold_advice: ColdAdvice,
    #[serde(skip_serializing_if = "Option::is_none")]
    cold_verified: Option<cold_verified::Sample>,
    process_metrics: Option<process_metrics::Delta>,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation_metrics: Option<crate::allocation_metrics::Sample>,
    output_sha256: Option<String>,
    output_bytes: Option<u64>,
    opc_materialized_parts: Option<u64>,
    cfb_changed_spans: Option<u64>,
    cfb_published_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    cfb_phases: Option<CfbPhaseEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    cfb_owned: Option<CfbOwnedEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pptx_source_replay: Option<PptxSourceReplayEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    docx_source_replay: Option<DocxSourceReplayEvidence>,
    /// Source identity and semantic projection for the unified XLSX facade
    /// path. These are collected after the timed operation so correctness I/O
    /// does not enter measured latency.
    #[serde(skip_serializing_if = "Option::is_none")]
    xlsx_source_sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    xlsx_semantic_sha256: Option<String>,
}

impl ChildResult {
    fn ineligible_cold_verified(proof: cold_verified::Sample) -> Self {
        Self {
            elapsed_ns: 0,
            logical_read_counter_scope: "cold_verified_ineligible".to_owned(),
            logical_read_calls: 0,
            logical_read_requested_bytes: 0,
            logical_read_bytes: 0,
            logical_read_largest_requested_bytes: 0,
            logical_read_largest_returned_bytes: 0,
            logical_read_pattern: None,
            max_concurrent_reads: 0,
            logical_read_request_sizes: Vec::new(),
            logical_read_request_size_buckets: ReadSizeBuckets::default(),
            cold_advice: ColdAdvice::NotRequested,
            cold_verified: Some(proof),
            process_metrics: None,
            allocation_metrics: None,
            output_sha256: None,
            output_bytes: None,
            opc_materialized_parts: None,
            cfb_changed_spans: None,
            cfb_published_bytes: None,
            cfb_phases: None,
            cfb_owned: None,
            pptx_source_replay: None,
            docx_source_replay: None,
            xlsx_source_sha256: None,
            xlsx_semantic_sha256: None,
        }
    }
}

/// Untimed source-backed replay evidence for the ordinary-root PPTX cases.
///
/// The replay is intentionally separate from the timed root facade operation.
/// Its counters describe only compressed ZIP payload-range overlap; central
/// directory, local-header, relationship, and mandatory catalog reads remain
/// in the unclassified totals. Eager controls leave this field absent and mark
/// their generic filesystem counter scope as not applicable.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct PptxSourceReplayEvidence {
    pub implementation: String,
    pub operation: String,
    pub source_bytes: u64,
    pub source_sha256: String,
    pub slide_count: usize,
    pub selected_position: Option<usize>,
    pub read_calls: u64,
    pub read_bytes: u64,
    pub read_return_sizes: Vec<u64>,
    pub slide_payload_read_calls: u64,
    pub slide_payload_read_bytes: u64,
    pub slide_payload_covered_bytes: u64,
    pub slide_payload_ranges_fully_covered: u64,
    pub selected_slide_payload_read_calls: u64,
    pub selected_slide_payload_read_bytes: u64,
    pub selected_slide_payload_covered_bytes: u64,
    pub selected_slide_payload_fully_covered: bool,
    pub unselected_slide_payload_read_calls: u64,
    pub unselected_slide_payload_read_bytes: u64,
    pub unselected_slide_payload_covered_bytes: u64,
    pub media_payload_read_calls: u64,
    pub media_payload_read_bytes: u64,
    pub media_payload_covered_bytes: u64,
    pub semantic_sha256: String,
    pub classification: String,
}

/// Untimed source-backed replay evidence for the ordinary-root DOCX cases.
///
/// The replay is independent from the timed `Document::open` facade
/// operation. It classifies logical `ReadAt` ranges against the compressed
/// main-document, media, unselected ordinary-part, and core-properties
/// ranges. Catalog/XML relationship reads remain in generic totals and are
/// not presented as payload I/O.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct DocxSourceReplayEvidence {
    pub implementation: String,
    pub operation: String,
    pub source_bytes: u64,
    pub source_sha256: String,
    pub paragraph_count: usize,
    pub open_read_calls: u64,
    pub open_read_bytes: u64,
    pub open_read_return_sizes: Vec<u64>,
    pub open_main_payload_overlap_bytes: u64,
    pub open_media_payload_overlap_bytes: u64,
    pub open_unselected_payload_overlap_bytes: u64,
    pub open_core_payload_overlap_bytes: u64,
    pub open_main_payload_covered_bytes: u64,
    pub open_media_payload_covered_bytes: u64,
    pub open_unselected_payload_covered_bytes: u64,
    pub open_core_payload_covered_bytes: u64,
    pub preparation_read_calls: u64,
    pub preparation_read_bytes: u64,
    pub preparation_read_return_sizes: Vec<u64>,
    pub preparation_main_payload_overlap_bytes: u64,
    pub preparation_main_payload_covered_bytes: u64,
    pub preparation_main_payload_fully_covered: bool,
    pub preparation_media_payload_overlap_bytes: u64,
    pub preparation_unselected_payload_overlap_bytes: u64,
    pub preparation_core_payload_overlap_bytes: u64,
    pub preparation_media_payload_covered_bytes: u64,
    pub preparation_unselected_payload_covered_bytes: u64,
    pub preparation_core_payload_covered_bytes: u64,
    pub query_read_calls: u64,
    pub query_read_bytes: u64,
    pub query_read_return_sizes: Vec<u64>,
    pub query_main_payload_overlap_bytes: u64,
    pub query_media_payload_overlap_bytes: u64,
    pub query_unselected_payload_overlap_bytes: u64,
    pub query_core_payload_overlap_bytes: u64,
    pub query_main_payload_covered_bytes: u64,
    pub query_media_payload_covered_bytes: u64,
    pub query_unselected_payload_covered_bytes: u64,
    pub query_core_payload_covered_bytes: u64,
    pub materializations: u64,
    pub semantic_sha256: String,
    pub classification: String,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "snake_case")]
pub(crate) enum ColdAdvice {
    NotRequested,
    Requested,
    Unsupported,
    Failed,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ChildMode {
    Prime,
    VerifiedPrime,
    Warm,
    Cold,
    ColdVerified,
}

impl ChildMode {
    const fn name(self) -> &'static str {
        match self {
            Self::Prime => "prime",
            Self::VerifiedPrime => "verified-prime",
            Self::Warm => "warm",
            Self::Cold => "cold",
            Self::ColdVerified => "cold-verified",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "prime" => Some(Self::Prime),
            "verified-prime" => Some(Self::VerifiedPrime),
            "warm" => Some(Self::Warm),
            "cold" => Some(Self::Cold),
            "cold-verified" => Some(Self::ColdVerified),
            _ => None,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct SampleEvidence {
    pub sample_index: usize,
    pub cache_state: &'static str,
    pub elapsed_ns: u64,
    pub parent_wall_ns: u64,
    pub cold_advice: ColdAdvice,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cold_verified: Option<cold_verified::Sample>,
    pub logical_read_counter_scope: String,
    pub logical_read_calls: u64,
    pub logical_read_requested_bytes: u64,
    pub logical_read_bytes: u64,
    pub logical_read_largest_requested_bytes: u64,
    pub logical_read_largest_returned_bytes: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub logical_read_pattern: Option<ReadPattern>,
    pub max_concurrent_reads: u64,
    pub logical_read_request_sizes: Vec<u64>,
    pub logical_read_request_size_buckets: ReadSizeBuckets,
    pub process_metrics: Option<process_metrics::Delta>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub allocation_metrics: Option<crate::allocation_metrics::Sample>,
    pub output_sha256: Option<String>,
    pub output_bytes: Option<u64>,
    pub opc_materialized_parts: Option<u64>,
    pub cfb_changed_spans: Option<u64>,
    pub cfb_published_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cfb_phases: Option<CfbPhaseEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pptx_source_replay: Option<PptxSourceReplayEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub docx_source_replay: Option<DocxSourceReplayEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xlsx_source_sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xlsx_semantic_sha256: Option<String>,
}

#[derive(Clone, Copy, Debug, Default, Deserialize, Serialize)]
pub(crate) struct ReadSizeBuckets {
    pub bytes_0: u64,
    pub bytes_1_to_512: u64,
    pub bytes_513_to_4096: u64,
    pub bytes_4097_to_16384: u64,
    pub bytes_16385_to_65536: u64,
    pub bytes_over_65536: u64,
}

impl ReadSizeBuckets {
    fn observe(&mut self, bytes: u64) -> io::Result<()> {
        let bucket = match bytes {
            0 => &mut self.bytes_0,
            1..=512 => &mut self.bytes_1_to_512,
            513..=4096 => &mut self.bytes_513_to_4096,
            4097..=16384 => &mut self.bytes_4097_to_16384,
            16385..=65536 => &mut self.bytes_16385_to_65536,
            _ => &mut self.bytes_over_65536,
        };
        *bucket = bucket
            .checked_add(1)
            .ok_or_else(|| io::Error::other("source request-size bucket count overflows u64"))?;
        Ok(())
    }

    const fn is_empty(&self) -> bool {
        self.bytes_0 == 0
            && self.bytes_1_to_512 == 0
            && self.bytes_513_to_4096 == 0
            && self.bytes_4097_to_16384 == 0
            && self.bytes_16385_to_65536 == 0
            && self.bytes_over_65536 == 0
    }
}

#[derive(Debug, Serialize)]
pub(crate) struct Evidence {
    pub case: &'static str,
    pub corpus: super::CorpusManifest,
    pub warmup_iterations: usize,
    pub sample_count: usize,
    pub cache_states: Vec<&'static str>,
    pub fresh_child_per_sample: bool,
    pub samples: Vec<SampleEvidence>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cold_verified_status: Option<cold_verified::Status>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cold_verified_samples: Option<Vec<cold_verified::Sample>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cold_verified_claim_scope: Option<&'static str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cold_verified_fincore_command: Option<&'static str>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cfb_owned: Option<Vec<CfbOwnedSampleEvidence>>,
}

/// Per-sample immutable-owned CFB provenance kept outside the generic
/// `SampleEvidence`/operation-metrics shape, whose positional-read vectors do
/// not apply to this ingress path.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct CfbOwnedSampleEvidence {
    pub sample_index: usize,
    pub cache_state: &'static str,
    pub evidence: CfbOwnedEvidence,
}

pub(crate) struct Run {
    pub warm_result: Option<super::CaseResult>,
    pub cold_result: Option<super::CaseResult>,
    pub cold_verified_result: Option<super::CaseResult>,
    pub evidence: Evidence,
}

#[derive(Default)]
struct OperationDetails {
    opc_materialized_parts: Option<u64>,
    cfb_changed_spans: Option<u64>,
    cfb_published_bytes: Option<u64>,
    cfb_phases: Option<CfbPhaseEvidence>,
    cfb_owned: Option<CfbOwnedEvidence>,
}

/// Operation-local attribution for the three sequential stages of the CFB
/// same-length atomic-save selector.
///
/// The counters are logical `ReadAt` deltas. They do not describe physical
/// device I/O, copied bytes, or allocation work.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct CfbPhaseEvidence {
    pub open: CfbPhaseSample,
    pub plan: CfbPhaseSample,
    pub atomic_publication: CfbPhaseSample,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
pub(crate) struct CfbPhaseSample {
    pub elapsed_ns: u64,
    pub logical_read_calls: u64,
    pub logical_read_requested_bytes: u64,
    pub logical_read_returned_bytes: u64,
}

/// Provenance and phase attribution for the opt-in immutable-owned CFB
/// filesystem selector.
///
/// This evidence deliberately has no logical `ReadAt` counters. The source
/// is read from the filesystem and sealed as `Arc<[u8]>` before the timed
/// operation, so reporting positional-read events for the timed phases would
/// misrepresent the ingress path. The generic logical counter fields remain
/// zero only as a legacy wire-shape placeholder and are paired with the
/// explicit not-applicable scope.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct CfbOwnedEvidence {
    pub implementation: String,
    pub ingress: String,
    pub ownership: String,
    pub logical_read_counter_scope: String,
    pub source_ingress_bytes: u64,
    pub source_sha256: String,
    pub source_version_id: u64,
    pub source_version_revision: u64,
    pub phases: CfbOwnedPhaseEvidence,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
pub(crate) struct CfbOwnedPhaseEvidence {
    pub open: CfbOwnedPhaseSample,
    pub plan: CfbOwnedPhaseSample,
    pub atomic_publication: CfbOwnedPhaseSample,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
pub(crate) struct CfbOwnedPhaseSample {
    pub elapsed_ns: u64,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    OpcEagerOpen,
    OpcSourceOpen,
    OpcEagerSave,
    OpcSourceSave,
    CfbOverlaySave,
    CfbOwnedOverlaySave,
    PptxEagerOpen,
    PptxSourceOpen,
    PptxEagerListSlides,
    PptxSourceListSlides,
    PptxEagerSlideCount,
    PptxSourceSlideCount,
    PptxEagerSelectedSlide,
    PptxSourceSelectedSlide,
    PptxEagerOpenSlideCountLifecycle,
    PptxSourceOpenSlideCountLifecycle,
    PptxEagerOpenSelectedSlideLifecycle,
    PptxSourceOpenSelectedSlideLifecycle,
    DocxEagerOpen,
    DocxSourceOpen,
    DocxEagerParagraphCount,
    DocxSourceParagraphCount,
    DocxEagerListParagraphs,
    DocxSourceListParagraphs,
    DocxEagerFullText,
    DocxSourceFullText,
    DocxEagerOpenParagraphCountLifecycle,
    DocxSourceOpenParagraphCountLifecycle,
    DocxEagerOpenFullTextLifecycle,
    DocxSourceOpenFullTextLifecycle,
    XlsxFileOpen,
    XlsxFileOpenLifecycle,
}

impl Operation {
    fn parse(name: &str) -> Option<Self> {
        match name {
            "opc_file_eager_open" => Some(Self::OpcEagerOpen),
            "opc_file_source_open" => Some(Self::OpcSourceOpen),
            "opc_file_eager_one_part_atomic_save" => Some(Self::OpcEagerSave),
            "opc_file_source_one_part_atomic_save" => Some(Self::OpcSourceSave),
            "cfb_file_same_length_overlay_atomic_save" => Some(Self::CfbOverlaySave),
            "cfb_file_owned_same_length_overlay_atomic_save" => Some(Self::CfbOwnedOverlaySave),
            "pptx_file_eager_open" => Some(Self::PptxEagerOpen),
            "pptx_file_source_open" => Some(Self::PptxSourceOpen),
            "pptx_file_eager_list_slides" => Some(Self::PptxEagerListSlides),
            "pptx_file_source_list_slides" => Some(Self::PptxSourceListSlides),
            "pptx_file_eager_slide_count" => Some(Self::PptxEagerSlideCount),
            "pptx_file_source_slide_count" => Some(Self::PptxSourceSlideCount),
            "pptx_file_eager_selected_slide" => Some(Self::PptxEagerSelectedSlide),
            "pptx_file_source_selected_slide" => Some(Self::PptxSourceSelectedSlide),
            "pptx_file_eager_open_slide_count_lifecycle" => {
                Some(Self::PptxEagerOpenSlideCountLifecycle)
            },
            "pptx_file_source_open_slide_count_lifecycle" => {
                Some(Self::PptxSourceOpenSlideCountLifecycle)
            },
            "pptx_file_eager_open_selected_slide_lifecycle" => {
                Some(Self::PptxEagerOpenSelectedSlideLifecycle)
            },
            "pptx_file_source_open_selected_slide_lifecycle" => {
                Some(Self::PptxSourceOpenSelectedSlideLifecycle)
            },
            "docx_file_eager_open" => Some(Self::DocxEagerOpen),
            "docx_file_source_open" => Some(Self::DocxSourceOpen),
            "docx_file_eager_paragraph_count" => Some(Self::DocxEagerParagraphCount),
            "docx_file_source_paragraph_count" => Some(Self::DocxSourceParagraphCount),
            "docx_file_eager_list_paragraphs" => Some(Self::DocxEagerListParagraphs),
            "docx_file_source_list_paragraphs" => Some(Self::DocxSourceListParagraphs),
            "docx_file_eager_full_text" => Some(Self::DocxEagerFullText),
            "docx_file_source_full_text" => Some(Self::DocxSourceFullText),
            "docx_file_eager_open_paragraph_count_lifecycle" => {
                Some(Self::DocxEagerOpenParagraphCountLifecycle)
            },
            "docx_file_source_open_paragraph_count_lifecycle" => {
                Some(Self::DocxSourceOpenParagraphCountLifecycle)
            },
            "docx_file_eager_open_full_text_lifecycle" => {
                Some(Self::DocxEagerOpenFullTextLifecycle)
            },
            "docx_file_source_open_full_text_lifecycle" => {
                Some(Self::DocxSourceOpenFullTextLifecycle)
            },
            "xlsx_file_open" => Some(Self::XlsxFileOpen),
            "xlsx_file_open_lifecycle" => Some(Self::XlsxFileOpenLifecycle),
            _ => None,
        }
    }

    const fn case(self) -> super::Case {
        match self {
            Self::OpcEagerOpen => super::Case::OpcFileEagerOpen,
            Self::OpcSourceOpen => super::Case::OpcFileSourceOpen,
            Self::OpcEagerSave => super::Case::OpcFileEagerOnePartAtomicSave,
            Self::OpcSourceSave => super::Case::OpcFileSourceOnePartAtomicSave,
            Self::CfbOverlaySave => super::Case::CfbFileSameLengthOverlayAtomicSave,
            Self::CfbOwnedOverlaySave => super::Case::CfbFileOwnedSameLengthOverlayAtomicSave,
            Self::PptxEagerOpen => super::Case::PptxFileEagerOpen,
            Self::PptxSourceOpen => super::Case::PptxFileSourceOpen,
            Self::PptxEagerListSlides => super::Case::PptxFileEagerListSlides,
            Self::PptxSourceListSlides => super::Case::PptxFileSourceListSlides,
            Self::PptxEagerSlideCount => super::Case::PptxFileEagerSlideCount,
            Self::PptxSourceSlideCount => super::Case::PptxFileSourceSlideCount,
            Self::PptxEagerSelectedSlide => super::Case::PptxFileEagerSelectedSlide,
            Self::PptxSourceSelectedSlide => super::Case::PptxFileSourceSelectedSlide,
            Self::PptxEagerOpenSlideCountLifecycle => {
                super::Case::PptxFileEagerOpenSlideCountLifecycle
            },
            Self::PptxSourceOpenSlideCountLifecycle => {
                super::Case::PptxFileSourceOpenSlideCountLifecycle
            },
            Self::PptxEagerOpenSelectedSlideLifecycle => {
                super::Case::PptxFileEagerOpenSelectedSlideLifecycle
            },
            Self::PptxSourceOpenSelectedSlideLifecycle => {
                super::Case::PptxFileSourceOpenSelectedSlideLifecycle
            },
            Self::DocxEagerOpen => super::Case::DocxFileEagerOpen,
            Self::DocxSourceOpen => super::Case::DocxFileSourceOpen,
            Self::DocxEagerParagraphCount => super::Case::DocxFileEagerParagraphCount,
            Self::DocxSourceParagraphCount => super::Case::DocxFileSourceParagraphCount,
            Self::DocxEagerListParagraphs => super::Case::DocxFileEagerListParagraphs,
            Self::DocxSourceListParagraphs => super::Case::DocxFileSourceListParagraphs,
            Self::DocxEagerFullText => super::Case::DocxFileEagerFullText,
            Self::DocxSourceFullText => super::Case::DocxFileSourceFullText,
            Self::DocxEagerOpenParagraphCountLifecycle => {
                super::Case::DocxFileEagerOpenParagraphCountLifecycle
            },
            Self::DocxSourceOpenParagraphCountLifecycle => {
                super::Case::DocxFileSourceOpenParagraphCountLifecycle
            },
            Self::DocxEagerOpenFullTextLifecycle => super::Case::DocxFileEagerOpenFullTextLifecycle,
            Self::DocxSourceOpenFullTextLifecycle => {
                super::Case::DocxFileSourceOpenFullTextLifecycle
            },
            Self::XlsxFileOpen => super::Case::XlsxFileOpen,
            Self::XlsxFileOpenLifecycle => super::Case::XlsxFileOpenLifecycle,
        }
    }

    const fn is_save(self) -> bool {
        matches!(
            self,
            Self::OpcEagerSave
                | Self::OpcSourceSave
                | Self::CfbOverlaySave
                | Self::CfbOwnedOverlaySave
        )
    }

    const fn is_cfb(self) -> bool {
        matches!(self, Self::CfbOverlaySave | Self::CfbOwnedOverlaySave)
    }

    const fn is_cfb_owned(self) -> bool {
        matches!(self, Self::CfbOwnedOverlaySave)
    }

    const fn is_pptx(self) -> bool {
        matches!(
            self,
            Self::PptxEagerOpen
                | Self::PptxSourceOpen
                | Self::PptxEagerListSlides
                | Self::PptxSourceListSlides
                | Self::PptxEagerSlideCount
                | Self::PptxSourceSlideCount
                | Self::PptxEagerSelectedSlide
                | Self::PptxSourceSelectedSlide
                | Self::PptxEagerOpenSlideCountLifecycle
                | Self::PptxSourceOpenSlideCountLifecycle
                | Self::PptxEagerOpenSelectedSlideLifecycle
                | Self::PptxSourceOpenSelectedSlideLifecycle
        )
    }

    const fn is_docx(self) -> bool {
        matches!(
            self,
            Self::DocxEagerOpen
                | Self::DocxSourceOpen
                | Self::DocxEagerParagraphCount
                | Self::DocxSourceParagraphCount
                | Self::DocxEagerListParagraphs
                | Self::DocxSourceListParagraphs
                | Self::DocxEagerFullText
                | Self::DocxSourceFullText
                | Self::DocxEagerOpenParagraphCountLifecycle
                | Self::DocxSourceOpenParagraphCountLifecycle
                | Self::DocxEagerOpenFullTextLifecycle
                | Self::DocxSourceOpenFullTextLifecycle
        )
    }

    const fn is_xlsx(self) -> bool {
        matches!(self, Self::XlsxFileOpen | Self::XlsxFileOpenLifecycle)
    }

    const fn is_docx_lifecycle(self) -> bool {
        matches!(
            self,
            Self::DocxEagerOpenParagraphCountLifecycle
                | Self::DocxSourceOpenParagraphCountLifecycle
                | Self::DocxEagerOpenFullTextLifecycle
                | Self::DocxSourceOpenFullTextLifecycle
        )
    }

    const fn is_source_docx(self) -> bool {
        matches!(
            self,
            Self::DocxSourceOpen
                | Self::DocxSourceParagraphCount
                | Self::DocxSourceListParagraphs
                | Self::DocxSourceFullText
                | Self::DocxSourceOpenParagraphCountLifecycle
                | Self::DocxSourceOpenFullTextLifecycle
        )
    }

    const fn is_docx_query(self) -> bool {
        matches!(
            self,
            Self::DocxEagerParagraphCount
                | Self::DocxSourceParagraphCount
                | Self::DocxEagerListParagraphs
                | Self::DocxSourceListParagraphs
                | Self::DocxEagerFullText
                | Self::DocxSourceFullText
        )
    }

    const fn docx_query_name(self) -> Option<&'static str> {
        match self {
            Self::DocxEagerOpen | Self::DocxSourceOpen => None,
            Self::DocxEagerParagraphCount | Self::DocxSourceParagraphCount => {
                Some("paragraph_count")
            },
            Self::DocxEagerListParagraphs | Self::DocxSourceListParagraphs => {
                Some("list_paragraphs")
            },
            Self::DocxEagerFullText | Self::DocxSourceFullText => Some("full_text"),
            Self::DocxEagerOpenParagraphCountLifecycle
            | Self::DocxSourceOpenParagraphCountLifecycle => Some("open_paragraph_count_lifecycle"),
            Self::DocxEagerOpenFullTextLifecycle | Self::DocxSourceOpenFullTextLifecycle => {
                Some("open_full_text_lifecycle")
            },
            _ => None,
        }
    }

    const fn is_source_pptx(self) -> bool {
        matches!(
            self,
            Self::PptxSourceOpen
                | Self::PptxSourceListSlides
                | Self::PptxSourceSlideCount
                | Self::PptxSourceSelectedSlide
                | Self::PptxSourceOpenSlideCountLifecycle
                | Self::PptxSourceOpenSelectedSlideLifecycle
        )
    }

    const fn is_pptx_query(self) -> bool {
        matches!(
            self,
            Self::PptxEagerListSlides
                | Self::PptxSourceListSlides
                | Self::PptxEagerSlideCount
                | Self::PptxSourceSlideCount
                | Self::PptxEagerSelectedSlide
                | Self::PptxSourceSelectedSlide
        )
    }

    /// Prepared query controls have already opened/materialized their source
    /// before the timed interval, so a page-cache proof would not describe the
    /// operation being measured.  Open and lifecycle controls remain eligible.
    const fn supports_cold_verified(self) -> bool {
        !self.is_pptx_query() && !self.is_docx_query()
    }

    const fn pptx_query_name(self) -> Option<&'static str> {
        match self {
            Self::PptxEagerOpen | Self::PptxSourceOpen => None,
            Self::PptxEagerListSlides | Self::PptxSourceListSlides => Some("list_slides"),
            Self::PptxEagerSlideCount | Self::PptxSourceSlideCount => Some("slide_count"),
            Self::PptxEagerSelectedSlide | Self::PptxSourceSelectedSlide => Some("selected_slide"),
            Self::PptxEagerOpenSlideCountLifecycle | Self::PptxSourceOpenSlideCountLifecycle => {
                Some("open_slide_count_lifecycle")
            },
            Self::PptxEagerOpenSelectedSlideLifecycle
            | Self::PptxSourceOpenSelectedSlideLifecycle => Some("open_selected_slide_lifecycle"),
            _ => None,
        }
    }
}

#[derive(Debug)]
struct Invocation {
    child: ChildResult,
    parent_wall_ns: u64,
}

/// Runs the selected filesystem cases in fresh child processes.
pub(crate) fn run_selected(
    cases: &[super::Case],
    warmup_iterations: usize,
    samples: usize,
    cache_selection: CacheSelection,
    requested_root: Option<&Path>,
) -> Result<Vec<Run>, Box<dyn Error>> {
    let selected = cases
        .iter()
        .copied()
        .filter_map(|case| Operation::parse(case.name()).map(|operation| (case, operation)))
        .collect::<Vec<_>>();
    if selected.is_empty() {
        return Ok(Vec::new());
    }
    if samples == 0 {
        return Err("filesystem evidence requires at least one measured sample".into());
    }

    let root = filesystem_root(requested_root)?;
    let result = (|| {
        let opc = super::build_opc_corpus(OPC_FILE_SHAPE, OPC_FILE_PAYLOAD)?;
        let cfb = super::build_ole_common_corpus(&opc)?;
        let pptx = selected
            .iter()
            .any(|(_, operation)| operation.is_pptx())
            .then(super::build_pptx_source_edit_corpus)
            .transpose()?;
        let docx = selected
            .iter()
            .any(|(_, operation)| operation.is_docx())
            .then(super::build_docx_source_edit_corpus)
            .transpose()?;
        let xlsx = selected
            .iter()
            .any(|(_, operation)| operation.is_xlsx())
            .then(|| super::build_xlsx_cell_crud_corpus(super::XlsxCellCrudShape::Medium))
            .transpose()?;
        assert_pinned_corpora(&opc, &cfb)?;
        if let Some(docx) = docx.as_ref() {
            assert_pinned_docx_corpus(docx)?;
        }
        if let Some(xlsx) = xlsx.as_ref() {
            assert_pinned_xlsx_corpus(xlsx)?;
        }
        let mut runs = Vec::with_capacity(selected.len());
        let mut opc_save_hashes: Option<Vec<(String, String)>> = None;
        for (case, operation) in selected {
            let corpus = if operation.is_cfb() {
                &cfb
            } else if operation.is_pptx() {
                pptx.as_ref()
                    .ok_or("PPTX filesystem corpus was not prepared")?
            } else if operation.is_docx() {
                docx.as_ref()
                    .ok_or("DOCX filesystem corpus was not prepared")?
            } else if operation.is_xlsx() {
                xlsx.as_ref()
                    .ok_or("XLSX filesystem corpus was not prepared")?
            } else {
                &opc
            };
            let run = run_one(
                case,
                operation,
                corpus,
                &root,
                warmup_iterations,
                samples,
                cache_selection,
            )?;
            if matches!(
                operation,
                Operation::OpcEagerSave | Operation::OpcSourceSave
            ) {
                let current = run
                    .evidence
                    .samples
                    .iter()
                    .map(|sample| {
                        Ok((
                            format!("{}:{}", sample.sample_index, sample.cache_state),
                            sample
                                .output_sha256
                                .clone()
                                .ok_or("filesystem OPC save emitted no output digest")?,
                        ))
                    })
                    .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
                if let Some(previous) = opc_save_hashes.as_ref() {
                    if previous != &current {
                        return Err(
                            "eager and source-backed OPC filesystem save samples differ".into()
                        );
                    }
                } else {
                    opc_save_hashes = Some(current);
                }
            }
            runs.push(run);
        }
        Ok(runs)
    })();
    let cleanup = fs::remove_dir_all(&root);
    match (result, cleanup) {
        (Ok(runs), Ok(())) => Ok(runs),
        (Ok(_), Err(error)) => Err(error.into()),
        (Err(error), _) => Err(error),
    }
}

fn assert_pinned_docx_corpus(corpus: &super::Corpus) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != DOCX_FILE_CORPUS_GENERATOR {
        return Err("DOCX filesystem corpus has the wrong generator".into());
    }
    if corpus.manifest.archive_sha256 != DOCX_FILE_SOURCE_SHA256 {
        return Err(format!(
            "DOCX filesystem source hash drifted: expected {DOCX_FILE_SOURCE_SHA256}, got {}",
            corpus.manifest.archive_sha256
        )
        .into());
    }
    Ok(())
}

fn assert_pinned_xlsx_corpus(corpus: &super::Corpus) -> Result<(), Box<dyn Error>> {
    let manifest = &corpus.manifest;
    if manifest.generator != XLSX_FILE_CORPUS_GENERATOR {
        return Err(format!(
            "XLSX filesystem corpus has the wrong generator: expected {XLSX_FILE_CORPUS_GENERATOR}, got {}",
            manifest.generator
        )
        .into());
    }
    if manifest.shape != XLSX_FILE_SOURCE_SHAPE {
        return Err(format!(
            "XLSX filesystem corpus has the wrong shape: expected {XLSX_FILE_SOURCE_SHAPE}, got {}",
            manifest.shape
        )
        .into());
    }
    if manifest.archive_bytes != XLSX_FILE_SOURCE_ARCHIVE_BYTES
        || corpus.archive.len() != XLSX_FILE_SOURCE_ARCHIVE_BYTES
    {
        return Err(format!(
            "XLSX filesystem source byte count drifted: expected {XLSX_FILE_SOURCE_ARCHIVE_BYTES}, manifest {}, archive {}",
            manifest.archive_bytes,
            corpus.archive.len()
        )
        .into());
    }
    if manifest.archive_sha256 != XLSX_FILE_SOURCE_SHA256
        || super::sha256_hex(&corpus.archive) != XLSX_FILE_SOURCE_SHA256
    {
        return Err(format!(
            "XLSX filesystem source hash drifted: expected {XLSX_FILE_SOURCE_SHA256}, manifest {}, archive {}",
            manifest.archive_sha256,
            super::sha256_hex(&corpus.archive)
        )
        .into());
    }
    if manifest.entry_count != XLSX_FILE_SOURCE_ENTRY_COUNT
        || manifest.archive_member_count != XLSX_FILE_SOURCE_ARCHIVE_MEMBER_COUNT
        || manifest.uncompressed_payload_bytes != XLSX_FILE_SOURCE_UNCOMPRESSED_PAYLOAD_BYTES
    {
        return Err(format!(
            "XLSX filesystem corpus topology drifted: entry_count {}, member_count {}, uncompressed_payload_bytes {}",
            manifest.entry_count,
            manifest.archive_member_count,
            manifest.uncompressed_payload_bytes
        )
        .into());
    }
    let xlsx = manifest
        .xlsx
        .as_ref()
        .ok_or("XLSX filesystem corpus omitted its typed shape manifest")?;
    if xlsx.sheet_count != XLSX_FILE_SOURCE_SHEET_COUNT
        || xlsx.rows_per_sheet != XLSX_FILE_SOURCE_ROWS_PER_SHEET
        || xlsx.columns_per_sheet != XLSX_FILE_SOURCE_COLUMNS_PER_SHEET
        || corpus.xlsx.as_ref().is_none_or(|spec| {
            spec.sheet_count != XLSX_FILE_SOURCE_SHEET_COUNT
                || spec.row_count != XLSX_FILE_SOURCE_ROWS_PER_SHEET
                || spec.column_count != XLSX_FILE_SOURCE_COLUMNS_PER_SHEET
        })
    {
        return Err(format!(
            "XLSX filesystem corpus typed shape drifted: manifest {}x{}x{}, expected {}x{}x{}",
            xlsx.sheet_count,
            xlsx.rows_per_sheet,
            xlsx.columns_per_sheet,
            XLSX_FILE_SOURCE_SHEET_COUNT,
            XLSX_FILE_SOURCE_ROWS_PER_SHEET,
            XLSX_FILE_SOURCE_COLUMNS_PER_SHEET
        )
        .into());
    }
    Ok(())
}

fn source_sha256_for_operation<'a>(operation: Operation, corpus: &'a super::Corpus) -> &'a str {
    if operation.is_xlsx() {
        XLSX_FILE_SOURCE_SHA256
    } else {
        &corpus.manifest.archive_sha256
    }
}

fn assert_pinned_corpora(opc: &super::Corpus, cfb: &super::Corpus) -> Result<(), Box<dyn Error>> {
    if opc.manifest.archive_sha256 != OPC_FILE_SOURCE_SHA256 {
        return Err(format!(
            "filesystem OPC source hash drifted: expected {OPC_FILE_SOURCE_SHA256}, got {}",
            opc.manifest.archive_sha256
        )
        .into());
    }
    if cfb.manifest.archive_sha256 != CFB_FILE_SOURCE_SHA256 {
        return Err(format!(
            "filesystem CFB source hash drifted: expected {CFB_FILE_SOURCE_SHA256}, got {}",
            cfb.manifest.archive_sha256
        )
        .into());
    }
    let replacement = filesystem_opc_replacement()?;
    let opc_output = super::sha256_hex(&super::expected_opc_overlay_output(opc, &replacement)?);
    if opc_output != OPC_FILE_EXPECTED_OUTPUT_SHA256 {
        return Err(format!(
            "filesystem OPC expected output hash drifted: expected {OPC_FILE_EXPECTED_OUTPUT_SHA256}, got {opc_output}"
        )
        .into());
    }
    let cfb_output = expected_cfb_output_digest(cfb)?;
    if cfb_output != CFB_FILE_EXPECTED_OUTPUT_SHA256 {
        return Err(format!(
            "filesystem CFB expected output hash drifted: expected {CFB_FILE_EXPECTED_OUTPUT_SHA256}, got {cfb_output}"
        )
        .into());
    }
    Ok(())
}

fn expected_cfb_output_digest(corpus: &super::Corpus) -> Result<String, Box<dyn Error>> {
    let shared = SharedOleFile::open(Arc::new(super::OwnedSource::new(corpus.archive.clone())))?;
    let overlay = SameLengthStreamOverlay::new(
        vec![super::OLE_COMMON_TARGET.to_owned()],
        Arc::from(FILESYSTEM_OLE_COMMON_REPLACEMENT.to_vec()),
    );
    let plan = shared.plan_same_length_stream_overlays(vec![overlay], OverlayLimits::default())?;
    let mut output = Vec::new();
    plan.write_to(&mut output)?;
    Ok(super::sha256_hex(&output))
}

fn run_one(
    case: super::Case,
    operation: Operation,
    corpus: &super::Corpus,
    root: &Path,
    warmup_iterations: usize,
    samples: usize,
    cache_selection: CacheSelection,
) -> Result<Run, Box<dyn Error>> {
    let stem = case.name();
    let source_path = if operation.is_pptx() {
        root.join(format!("{stem}.pptx"))
    } else if operation.is_docx() {
        root.join(format!("{stem}.docx"))
    } else if operation.is_xlsx() {
        root.join(format!("{stem}.xlsx"))
    } else {
        root.join(format!("{stem}.source"))
    };
    let expected_source_sha256 = source_sha256_for_operation(operation, corpus);
    write_synced(&source_path, &corpus.archive)?;
    assert_source_sha256(&source_path, expected_source_sha256)?;
    let (verified_source_path, verified_source_sha256, mut cold_verified_status) =
        if cache_selection.cold_verified() && operation.supports_cold_verified() {
            match cold_verified::page_size_for_harness() {
                Ok(page_size_bytes) => {
                    let aligned = cold_verified::page_aligned_archive(
                        &corpus.archive,
                        page_size_bytes,
                        !operation.is_cfb(),
                    );
                    match aligned {
                        Ok(aligned) => {
                            let path = root.join(format!("{stem}.cold-verified"));
                            match write_synced(&path, &aligned) {
                                Ok(()) => {
                                    let hash = super::sha256_hex(&aligned);
                                    let preflight = cold_verified::prepare(&path);
                                    let status = preflight.status;
                                    (Some(path), Some(hash), Some(status))
                                },
                                Err(_) => (
                                    None,
                                    None,
                                    Some(cold_verified::Status::IneligibleSourceWriteFailed),
                                ),
                            }
                        },
                        Err(status) => (None, None, Some(status)),
                    }
                },
                Err(status) => (None, None, Some(status)),
            }
        } else if cache_selection.cold_verified() {
            (
                None,
                None,
                Some(cold_verified::Status::IneligiblePreparedQueryControl),
            )
        } else {
            (None, None, None)
        };
    let mut cold_verified_samples =
        if cold_verified_status.is_some_and(|status| !status.is_eligible()) {
            verified_source_path
                .as_deref()
                .map(cold_verified::prepare)
                .filter(|sample| !sample.status.is_eligible())
                .map(|sample| vec![sample])
        } else {
            None
        };
    let destination_path = root.join(format!("{stem}.destination"));
    let expected_digest = operation
        .is_save()
        .then(|| expected_digest(operation, corpus))
        .transpose()?;
    let verified_expected_digest = verified_expected_digest(
        operation,
        cache_selection,
        cold_verified_status,
        verified_source_path.as_deref(),
    )?;
    let expected_xlsx_semantic_sha256 = operation
        .is_xlsx()
        .then(|| xlsx_semantic_sha256(corpus))
        .transpose()?;

    // Untimed warmups are independent children. Each measured sample then
    // primes with a separate child immediately before the selected measured
    // state. Save cases are reset after priming so every operation has the
    // same pre-existing destination.
    for _ in 0..warmup_iterations {
        if operation.is_save() {
            seed_destination(&destination_path, &corpus.archive)?;
        }
        let _ = spawn_checked_child(
            operation,
            &source_path,
            &destination_path,
            ChildMode::Prime,
            expected_source_sha256,
        )?;
    }

    let mut warm_elapsed = Vec::with_capacity(samples);
    let mut cold_elapsed = Vec::with_capacity(samples);
    let mut cold_verified_elapsed = Vec::with_capacity(samples);
    let mut sample_evidence = Vec::with_capacity(samples * 2);
    let mut cfb_owned_evidence = Vec::new();
    for sample_index in 0..samples {
        if cache_selection.warm() {
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let _ = spawn_checked_child(
                operation,
                &source_path,
                &destination_path,
                ChildMode::Prime,
                expected_source_sha256,
            )?;
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let warm = spawn_checked_child(
                operation,
                &source_path,
                &destination_path,
                ChildMode::Warm,
                expected_source_sha256,
            )?;
            verify_child_output(operation, &source_path, &destination_path, corpus, false)?;
            warm_elapsed.push(warm.child.elapsed_ns);
            record_sample(
                &mut sample_evidence,
                sample_index,
                "warm",
                warm,
                operation,
                expected_digest.as_deref(),
                expected_source_sha256,
                expected_xlsx_semantic_sha256.as_deref(),
                stem,
                &mut cfb_owned_evidence,
            )?;
        }

        if cache_selection.cold_requested() {
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let _ = spawn_checked_child(
                operation,
                &source_path,
                &destination_path,
                ChildMode::Prime,
                expected_source_sha256,
            )?;
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let cold = spawn_checked_child(
                operation,
                &source_path,
                &destination_path,
                ChildMode::Cold,
                expected_source_sha256,
            )?;
            verify_child_output(operation, &source_path, &destination_path, corpus, false)?;
            cold_elapsed.push(cold.child.elapsed_ns);
            record_sample(
                &mut sample_evidence,
                sample_index,
                "cold-requested",
                cold,
                operation,
                expected_digest.as_deref(),
                expected_source_sha256,
                expected_xlsx_semantic_sha256.as_deref(),
                stem,
                &mut cfb_owned_evidence,
            )?;
        }

        if cache_selection.cold_verified()
            && operation.supports_cold_verified()
            && cold_verified_status.is_some_and(|status| status.is_eligible())
        {
            let verified_source = verified_source_path
                .as_deref()
                .ok_or("cold-verified source was not prepared")?;
            let verified_sha256 = verified_source_sha256
                .as_deref()
                .ok_or("cold-verified source hash was not prepared")?;
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let _ = spawn_checked_child(
                operation,
                verified_source,
                &destination_path,
                ChildMode::VerifiedPrime,
                verified_sha256,
            )?;
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let verified = spawn_checked_child(
                operation,
                verified_source,
                &destination_path,
                ChildMode::ColdVerified,
                verified_sha256,
            )?;
            let proof = verified
                .child
                .cold_verified
                .clone()
                .ok_or("cold-verified child omitted its proof")?;
            if proof.aligned_source_sha256.as_deref() != Some(verified_sha256) {
                return Err(
                    "cold-verified proof source hash differs from the prepared aligned source"
                        .into(),
                );
            }
            let aligned_source_bytes = fs::metadata(verified_source)?.len();
            if proof.aligned_source_bytes != Some(aligned_source_bytes) {
                return Err(
                    "cold-verified proof source size differs from the prepared aligned source"
                        .into(),
                );
            }
            cold_verified_samples
                .get_or_insert_with(Vec::new)
                .push(proof.clone());
            if !proof.status.is_eligible() {
                cold_verified_status = Some(proof.status);
                continue;
            }
            verify_child_output(operation, verified_source, &destination_path, corpus, true)?;
            cold_verified_elapsed.push(verified.child.elapsed_ns);
            record_sample(
                &mut sample_evidence,
                sample_index,
                "cold-verified",
                verified,
                operation,
                verified_expected_digest.as_deref(),
                verified_sha256,
                expected_xlsx_semantic_sha256.as_deref(),
                stem,
                &mut cfb_owned_evidence,
            )?;
        }
    }

    let warm_result = if cache_selection.warm() {
        Some(filesystem_result(
            case,
            corpus,
            warm_elapsed,
            "warm",
            expected_digest.clone(),
            &sample_evidence,
        )?)
    } else {
        None
    };
    let cold_result = if cache_selection.cold_requested() {
        Some(filesystem_result(
            case,
            corpus,
            cold_elapsed,
            "cold-requested",
            expected_digest.clone(),
            &sample_evidence,
        )?)
    } else {
        None
    };
    let cold_verified_result = if cache_selection.cold_verified()
        && operation.supports_cold_verified()
        && cold_verified_status == Some(cold_verified::Status::Eligible)
        && cold_verified_elapsed.len() == samples
    {
        Some(filesystem_result(
            case,
            corpus,
            cold_verified_elapsed,
            "cold-verified",
            verified_expected_digest,
            &sample_evidence,
        )?)
    } else {
        None
    };

    Ok(Run {
        warm_result,
        cold_result,
        evidence: Evidence {
            case: case.name(),
            corpus: corpus.manifest.clone(),
            warmup_iterations,
            sample_count: samples,
            cache_states: cache_selection.names(),
            fresh_child_per_sample: true,
            samples: sample_evidence,
            cold_verified_status: cache_selection
                .cold_verified()
                .then_some(cold_verified_status)
                .flatten(),
            cold_verified_samples: cold_verified_samples,
            cold_verified_claim_scope: cache_selection
                .cold_verified()
                .then_some(cold_verified::CLAIM_SCOPE),
            cold_verified_fincore_command: cache_selection
                .cold_verified()
                .then_some(cold_verified::FINCORE_COMMAND),
            cfb_owned: (!cfb_owned_evidence.is_empty()).then_some(cfb_owned_evidence),
        },
        cold_verified_result,
    })
}

fn verified_expected_digest(
    operation: Operation,
    cache_selection: CacheSelection,
    cold_verified_status: Option<cold_verified::Status>,
    verified_source_path: Option<&Path>,
) -> Result<Option<String>, Box<dyn Error>> {
    if !operation.is_save()
        || !cache_selection.cold_verified()
        || !operation.supports_cold_verified()
        || cold_verified_status != Some(cold_verified::Status::Eligible)
    {
        return Ok(None);
    }
    verified_source_path
        .map(|path| expected_digest_for_source(operation, path))
        .transpose()
}

fn record_sample(
    samples: &mut Vec<SampleEvidence>,
    sample_index: usize,
    cache_state: &'static str,
    invocation: Invocation,
    operation: Operation,
    expected_digest: Option<&str>,
    expected_source_sha256: &str,
    expected_xlsx_semantic_sha256: Option<&str>,
    stem: &str,
    cfb_owned_evidence: &mut Vec<CfbOwnedSampleEvidence>,
) -> Result<(), Box<dyn Error>> {
    if cache_state == "cold-verified" {
        let proof = invocation
            .child
            .cold_verified
            .as_ref()
            .ok_or("cold-verified sample omitted its proof")?;
        if !proof.status.is_eligible() || proof.read_bytes_delta.unwrap_or(0) == 0 {
            return Err("cold-verified sample did not prove a positive storage read".into());
        }
    }
    if operation.is_xlsx() {
        if invocation.child.xlsx_source_sha256.as_deref() != Some(expected_source_sha256) {
            return Err(format!(
                "{stem} {cache_state} child source hash differs from the expected source"
            )
            .into());
        }
        if invocation.child.xlsx_semantic_sha256.as_deref() != expected_xlsx_semantic_sha256 {
            return Err(format!(
                "{stem} {cache_state} child semantic projection differs from the deterministic corpus"
            )
            .into());
        }
    }
    validate_cfb_owned_evidence(operation, &invocation.child)?;
    if !operation.is_cfb_owned() {
        validate_cfb_phase_evidence(operation, &invocation.child)?;
    }
    if operation.is_save() {
        if invocation.child.output_sha256.as_deref() != expected_digest {
            return Err(format!(
                "{stem} {cache_state} output digest differs from deterministic expectation"
            )
            .into());
        }
    } else if invocation.child.output_sha256.is_some() {
        return Err(
            format!("{stem} {cache_state} emitted an output hash for an open operation").into(),
        );
    }
    if let Some(evidence) = invocation.child.cfb_owned.as_ref() {
        cfb_owned_evidence.push(CfbOwnedSampleEvidence {
            sample_index,
            cache_state,
            evidence: evidence.clone(),
        });
    }
    samples.push(SampleEvidence {
        sample_index,
        cache_state,
        elapsed_ns: invocation.child.elapsed_ns,
        parent_wall_ns: invocation.parent_wall_ns,
        cold_advice: invocation.child.cold_advice,
        cold_verified: invocation.child.cold_verified,
        logical_read_counter_scope: invocation.child.logical_read_counter_scope,
        logical_read_calls: invocation.child.logical_read_calls,
        logical_read_requested_bytes: invocation.child.logical_read_requested_bytes,
        logical_read_bytes: invocation.child.logical_read_bytes,
        logical_read_largest_requested_bytes: invocation.child.logical_read_largest_requested_bytes,
        logical_read_largest_returned_bytes: invocation.child.logical_read_largest_returned_bytes,
        logical_read_pattern: invocation.child.logical_read_pattern,
        max_concurrent_reads: invocation.child.max_concurrent_reads,
        logical_read_request_sizes: invocation.child.logical_read_request_sizes,
        logical_read_request_size_buckets: invocation.child.logical_read_request_size_buckets,
        process_metrics: invocation.child.process_metrics,
        allocation_metrics: invocation.child.allocation_metrics,
        output_sha256: invocation.child.output_sha256,
        output_bytes: invocation.child.output_bytes,
        opc_materialized_parts: invocation.child.opc_materialized_parts,
        cfb_changed_spans: invocation.child.cfb_changed_spans,
        cfb_published_bytes: invocation.child.cfb_published_bytes,
        cfb_phases: invocation.child.cfb_phases,
        pptx_source_replay: invocation.child.pptx_source_replay,
        docx_source_replay: invocation.child.docx_source_replay,
        xlsx_source_sha256: invocation.child.xlsx_source_sha256,
        xlsx_semantic_sha256: invocation.child.xlsx_semantic_sha256,
    });
    Ok(())
}

fn validate_cfb_phase_evidence(
    operation: Operation,
    child: &ChildResult,
) -> Result<(), Box<dyn Error>> {
    let Some(phases) = child.cfb_phases.as_ref() else {
        if operation.is_cfb() {
            return Err("CFB filesystem sample omitted phase evidence".into());
        }
        return Ok(());
    };
    if !operation.is_cfb() {
        return Err("non-CFB filesystem sample unexpectedly reported CFB phases".into());
    }
    let phase_values = [phases.open, phases.plan, phases.atomic_publication];
    let sum = |name: &str, value: fn(&CfbPhaseSample) -> u64| {
        phase_values.iter().try_fold(0_u64, |total, phase| {
            total
                .checked_add(value(phase))
                .ok_or_else(|| io::Error::other(format!("CFB phase {name} total overflows u64")))
        })
    };
    let elapsed = sum("elapsed", |phase| phase.elapsed_ns)?;
    if elapsed > child.elapsed_ns {
        return Err(format!(
            "CFB phase elapsed total {elapsed} exceeds operation elapsed {}",
            child.elapsed_ns
        )
        .into());
    }
    for (name, observed, expected) in [
        (
            "logical read calls",
            sum("logical read calls", |phase| phase.logical_read_calls)?,
            child.logical_read_calls,
        ),
        (
            "logical requested bytes",
            sum("logical requested bytes", |phase| {
                phase.logical_read_requested_bytes
            })?,
            child.logical_read_requested_bytes,
        ),
        (
            "logical returned bytes",
            sum("logical returned bytes", |phase| {
                phase.logical_read_returned_bytes
            })?,
            child.logical_read_bytes,
        ),
    ] {
        if observed != expected {
            return Err(format!(
                "CFB phase {name} total {observed} differs from operation total {expected}"
            )
            .into());
        }
    }
    Ok(())
}

fn validate_cfb_owned_evidence(
    operation: Operation,
    child: &ChildResult,
) -> Result<(), Box<dyn Error>> {
    if !operation.is_cfb_owned() {
        if child.cfb_owned.is_some() {
            return Err(
                "non-owned filesystem sample unexpectedly reported owned CFB evidence".into(),
            );
        }
        return Ok(());
    }
    if child.cfb_phases.is_some() {
        return Err("owned CFB filesystem sample unexpectedly reported logical-read phases".into());
    }
    let owned = child
        .cfb_owned
        .as_ref()
        .ok_or("owned CFB filesystem sample omitted ownership evidence")?;
    if owned.implementation != "SharedOleFile::open_owned" {
        return Err("owned CFB filesystem sample reported an unexpected implementation".into());
    }
    if owned.ingress != "filesystem_read_all_before_cfb_phase_timers"
        || owned.ownership != "Arc<[u8]>"
        || owned.logical_read_counter_scope != "not_applicable_immutable_owned_slice"
    {
        return Err("owned CFB filesystem sample reported unexpected provenance".into());
    }
    if owned.source_ingress_bytes == 0 || owned.source_sha256.len() != 64 {
        return Err("owned CFB filesystem sample reported invalid source provenance".into());
    }
    let phases = [
        owned.phases.open,
        owned.phases.plan,
        owned.phases.atomic_publication,
    ];
    let elapsed = phases.iter().try_fold(0_u64, |total, phase| {
        total
            .checked_add(phase.elapsed_ns)
            .ok_or_else(|| io::Error::other("owned CFB phase elapsed total overflows u64"))
    })?;
    if elapsed > child.elapsed_ns {
        return Err(format!(
            "owned CFB phase elapsed total {elapsed} exceeds operation elapsed {}",
            child.elapsed_ns
        )
        .into());
    }
    if child.logical_read_counter_scope != "not_applicable_immutable_owned_slice"
        || child.logical_read_calls != 0
        || child.logical_read_requested_bytes != 0
        || child.logical_read_bytes != 0
        || child.logical_read_largest_requested_bytes != 0
        || child.logical_read_largest_returned_bytes != 0
        || child.logical_read_pattern.is_some()
        || child.max_concurrent_reads != 0
        || !child.logical_read_request_sizes.is_empty()
        || !child.logical_read_request_size_buckets.is_empty()
    {
        return Err("owned CFB filesystem sample exposed fabricated logical-read counters".into());
    }
    Ok(())
}

fn filesystem_result(
    case: super::Case,
    corpus: &super::Corpus,
    elapsed: Vec<u64>,
    cache_state: &'static str,
    output_sha256: Option<String>,
    samples: &[SampleEvidence],
) -> Result<super::CaseResult, Box<dyn Error>> {
    let mut result = super::result(case, corpus, elapsed, None);
    result.cache_state = Some(cache_state);
    result.output_sha256 = output_sha256;
    result.operation_metrics = Some(crate::operation_metrics::aggregate(
        samples,
        cache_state,
        &result.elapsed_ns.samples,
    )?);
    Ok(result)
}

fn expected_digest(operation: Operation, corpus: &super::Corpus) -> Result<String, Box<dyn Error>> {
    match operation {
        Operation::OpcEagerSave | Operation::OpcSourceSave => {
            let replacement = filesystem_opc_replacement()?;
            Ok(super::sha256_hex(&super::expected_opc_overlay_output(
                corpus,
                &replacement,
            )?))
        },
        Operation::CfbOverlaySave | Operation::CfbOwnedOverlaySave => {
            expected_cfb_output_digest(corpus)
        },
        Operation::OpcEagerOpen
        | Operation::OpcSourceOpen
        | Operation::PptxEagerOpen
        | Operation::PptxSourceOpen
        | Operation::PptxEagerListSlides
        | Operation::PptxSourceListSlides
        | Operation::PptxEagerSlideCount
        | Operation::PptxSourceSlideCount
        | Operation::PptxEagerSelectedSlide
        | Operation::PptxSourceSelectedSlide
        | Operation::PptxEagerOpenSlideCountLifecycle
        | Operation::PptxSourceOpenSlideCountLifecycle
        | Operation::PptxEagerOpenSelectedSlideLifecycle
        | Operation::PptxSourceOpenSelectedSlideLifecycle
        | Operation::DocxEagerOpen
        | Operation::DocxSourceOpen
        | Operation::DocxEagerParagraphCount
        | Operation::DocxSourceParagraphCount
        | Operation::DocxEagerListParagraphs
        | Operation::DocxSourceListParagraphs
        | Operation::DocxEagerFullText
        | Operation::DocxSourceFullText
        | Operation::DocxEagerOpenParagraphCountLifecycle
        | Operation::DocxSourceOpenParagraphCountLifecycle
        | Operation::DocxEagerOpenFullTextLifecycle
        | Operation::DocxSourceOpenFullTextLifecycle => {
            Err("open operation has no output digest".into())
        },
        Operation::XlsxFileOpen | Operation::XlsxFileOpenLifecycle => {
            Err("open operation has no output digest".into())
        },
    }
}

/// Computes the deterministic save digest for the private page-aligned
/// verifier source.  ZIP source-backed publication intentionally preserves
/// the verifier's EOCD comment, and a padded CFB source has a different
/// physical length, so the ordinary corpus digest is not the right sink
/// expectation for a verified sample.
fn expected_digest_for_source(
    operation: Operation,
    source: &Path,
) -> Result<String, Box<dyn Error>> {
    match operation {
        Operation::OpcEagerSave => {
            let target_uri =
                PackURI::new(format!("/{}", super::entry_name(OPC_FILE_TARGET_INDEX)))?;
            let mut package = OpcPackage::from_bytes(&fs::read(source)?)?;
            package
                .get_part_mut(&target_uri)?
                .set_blob(filesystem_opc_replacement()?);
            Ok(super::sha256_hex(&PackageWriter::to_bytes(&package)?))
        },
        Operation::OpcSourceSave => {
            let target_uri =
                PackURI::new(format!("/{}", super::entry_name(OPC_FILE_TARGET_INDEX)))?;
            let package = SourceBackedPackage::from_read_at(Arc::new(FileSource::open(source)?))?;
            let mut output = Vec::new();
            package.write_part_overlay_to_stream(
                &mut output,
                &target_uri,
                filesystem_opc_replacement()?,
            )?;
            Ok(super::sha256_hex(&output))
        },
        Operation::CfbOverlaySave | Operation::CfbOwnedOverlaySave => {
            let shared = SharedOleFile::open(Arc::new(super::OwnedSource::new(fs::read(source)?)))?;
            let overlay = SameLengthStreamOverlay::new(
                vec![super::OLE_COMMON_TARGET.to_owned()],
                Arc::from(FILESYSTEM_OLE_COMMON_REPLACEMENT.to_vec()),
            );
            let plan =
                shared.plan_same_length_stream_overlays(vec![overlay], OverlayLimits::default())?;
            let mut output = Vec::new();
            plan.write_to(&mut output)?;
            Ok(super::sha256_hex(&output))
        },
        _ => Err("verified source digest requested for a non-save operation".into()),
    }
}

fn spawn_child(
    operation: Operation,
    source: &Path,
    destination: &Path,
    mode: ChildMode,
) -> Result<Invocation, Box<dyn Error>> {
    let executable = env::current_exe()?;
    let started = Instant::now();
    let output = Command::new(executable)
        .arg("--filesystem-child")
        .arg(operation.case().name())
        .arg(source)
        .arg(destination)
        .arg(mode.name())
        .stdin(Stdio::null())
        .output()?;
    let parent_wall_ns = u64::try_from(started.elapsed().as_nanos())?;
    if !output.status.success() {
        return Err(format!(
            "filesystem child {} failed with {}: {}",
            operation.case().name(),
            output.status,
            String::from_utf8_lossy(&output.stderr)
        )
        .into());
    }
    let child = serde_json::from_slice(&output.stdout).map_err(|error| {
        format!(
            "filesystem child {} emitted invalid JSON: {error}; stdout={:?}",
            operation.case().name(),
            String::from_utf8_lossy(&output.stdout)
        )
    })?;
    Ok(Invocation {
        child,
        parent_wall_ns,
    })
}

fn spawn_checked_child(
    operation: Operation,
    source: &Path,
    destination: &Path,
    mode: ChildMode,
    expected_source_sha256: &str,
) -> Result<Invocation, Box<dyn Error>> {
    let invocation = spawn_child(operation, source, destination, mode)?;
    // This check is deliberately parent-side and outside the child timer. It
    // protects every prime and measured sample from a source file changing
    // between samples or while the isolated child was running.
    assert_source_sha256(source, expected_source_sha256)?;
    Ok(invocation)
}

fn assert_source_sha256(source: &Path, expected_sha256: &str) -> Result<(), Box<dyn Error>> {
    let actual_sha256 = super::sha256_hex(&fs::read(source)?);
    if actual_sha256 != expected_sha256 {
        return Err(format!(
            "filesystem source changed: expected SHA-256 {expected_sha256}, got {actual_sha256}"
        )
        .into());
    }
    Ok(())
}

/// Handles the private child protocol. `true` means the normal report path is
/// skipped.
pub(crate) fn run_child_if_requested() -> Result<bool, Box<dyn Error>> {
    run_child_arguments(env::args_os().skip(1))
}

fn run_child_arguments<I>(arguments: I) -> Result<bool, Box<dyn Error>>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut arguments = arguments.into_iter();
    let Some(first) = arguments.next() else {
        return Ok(false);
    };
    if first != "--filesystem-child" {
        return Ok(false);
    }
    let case_name = arguments
        .next()
        .ok_or("filesystem child is missing its case")?
        .to_string_lossy()
        .into_owned();
    let source = PathBuf::from(
        arguments
            .next()
            .ok_or("filesystem child is missing its source path")?,
    );
    let destination = PathBuf::from(
        arguments
            .next()
            .ok_or("filesystem child is missing its destination path")?,
    );
    let mode = ChildMode::parse(
        &arguments
            .next()
            .ok_or("filesystem child is missing its mode")?
            .to_string_lossy(),
    )
    .ok_or("filesystem child mode must be prime, verified-prime, warm, cold, or cold-verified")?;
    if arguments.next().is_some() {
        return Err("filesystem child received unexpected arguments".into());
    }
    let operation = Operation::parse(&case_name).ok_or("unknown filesystem child case")?;
    if mode == ChildMode::ColdVerified && !operation.supports_cold_verified() {
        serde_json::to_writer(
            io::stdout().lock(),
            &ChildResult::ineligible_cold_verified(cold_verified::Sample::ineligible(
                cold_verified::Status::IneligiblePreparedQueryControl,
            )),
        )?;
        return Ok(true);
    }
    let opc_replacement = matches!(
        operation,
        Operation::OpcEagerSave | Operation::OpcSourceSave
    )
    .then(filesystem_opc_replacement)
    .transpose()?;
    let cold_advice = if mode == ChildMode::Cold {
        request_cold(&source)
    } else {
        ColdAdvice::NotRequested
    };
    let (cold_verified_preparation, verified_before) = if mode == ChildMode::ColdVerified {
        let preparation = cold_verified::prepare(&source);
        if !preparation.status.is_eligible() {
            serde_json::to_writer(
                io::stdout().lock(),
                &ChildResult::ineligible_cold_verified(preparation),
            )?;
            return Ok(true);
        }
        let before = match process_metrics::Snapshot::read() {
            Ok(before) => before,
            Err(_) => {
                let mut ineligible = preparation;
                ineligible.status = cold_verified::Status::IneligibleProcIoUnavailable;
                serde_json::to_writer(
                    io::stdout().lock(),
                    &ChildResult::ineligible_cold_verified(ineligible),
                )?;
                return Ok(true);
            },
        };
        (Some(preparation), Some(before))
    } else {
        (None, None)
    };
    // Existing query controls compare prepared eager/source roots, so root
    // construction is outside their query timer. Lifecycle controls leave
    // the root unprepared and include fresh open plus the selected query in
    // the timed operation.
    let prepared_pptx = if operation.is_pptx_query() {
        if operation.is_source_pptx() {
            Some(litchi::Presentation::open(&source)?)
        } else {
            Some(litchi::Presentation::from_bytes(fs::read(&source)?)?)
        }
    } else {
        None
    };
    let prepared_docx = if operation.is_docx_query() {
        Some(if operation.is_source_docx() {
            PreparedDocx::source(&source)?
        } else {
            PreparedDocx::eager(fs::read(&source)?)?
        })
    } else {
        None
    };
    let before = verified_before.or_else(|| process_metrics::Snapshot::read().ok());
    let allocation_region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let mut details = OperationDetails::default();
    let mut deferred_source_open_package = None;
    let mut deferred_xlsx_operation = None;
    let counter_result = (|| -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
        Ok(match operation {
            Operation::OpcEagerOpen => run_opc_eager_open(&source, &mut details)?,
            Operation::OpcSourceOpen => {
                let (counter, package) = run_opc_source_open(&source)?;
                deferred_source_open_package = Some(package);
                counter
            },
            Operation::OpcEagerSave => run_opc_eager_save(
                &source,
                &destination,
                opc_replacement
                    .as_deref()
                    .ok_or("missing OPC replacement")?,
                &mut details,
            )?,
            Operation::OpcSourceSave => run_opc_source_save(
                &source,
                &destination,
                opc_replacement
                    .as_deref()
                    .ok_or("missing OPC replacement")?,
                &mut details,
            )?,
            Operation::CfbOverlaySave => run_cfb_overlay_save(&source, &destination, &mut details)?,
            Operation::CfbOwnedOverlaySave => {
                run_cfb_owned_overlay_save(&source, &destination, &mut details)?
            },
            Operation::PptxEagerOpen
            | Operation::PptxSourceOpen
            | Operation::PptxEagerListSlides
            | Operation::PptxSourceListSlides
            | Operation::PptxEagerSlideCount
            | Operation::PptxSourceSlideCount
            | Operation::PptxEagerSelectedSlide
            | Operation::PptxSourceSelectedSlide
            | Operation::PptxEagerOpenSlideCountLifecycle
            | Operation::PptxSourceOpenSlideCountLifecycle
            | Operation::PptxEagerOpenSelectedSlideLifecycle
            | Operation::PptxSourceOpenSelectedSlideLifecycle => {
                run_pptx_operation(operation, &source, prepared_pptx.as_ref())?;
                None
            },
            Operation::DocxEagerOpen
            | Operation::DocxSourceOpen
            | Operation::DocxEagerParagraphCount
            | Operation::DocxSourceParagraphCount
            | Operation::DocxEagerListParagraphs
            | Operation::DocxSourceListParagraphs
            | Operation::DocxEagerFullText
            | Operation::DocxSourceFullText
            | Operation::DocxEagerOpenParagraphCountLifecycle
            | Operation::DocxSourceOpenParagraphCountLifecycle
            | Operation::DocxEagerOpenFullTextLifecycle
            | Operation::DocxSourceOpenFullTextLifecycle => {
                run_docx_operation(operation, &source, prepared_docx.as_ref())?;
                None
            },
            Operation::XlsxFileOpen | Operation::XlsxFileOpenLifecycle => {
                run_xlsx_operation(operation, &source, &mut deferred_xlsx_operation)?;
                None
            },
        })
    })();
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos())?;
    let allocation_metrics = allocation_region.finish();
    let counter = counter_result?;
    let after = process_metrics::Snapshot::read().ok();
    let process_delta = before.zip(after).map(|(before, after)| after.delta(before));
    let cold_verified = cold_verified_preparation
        .map(|preparation| cold_verified::complete(preparation, before, after));
    let snapshot =
        counter.map_or_else(|| Ok(ReadMetrics::default()), |counter| counter.snapshot())?;

    // All operation-only evidence is now captured. Any package diagnostics,
    // source replay, semantic projection, and source hashing below are
    // deliberately untimed correctness work.
    if let Some(package) = deferred_source_open_package {
        details.opc_materialized_parts = Some(package.try_cache_diagnostics()?.successful_loads);
        std::hint::black_box(package);
    }
    let logical_read_counter_scope = if operation.is_cfb_owned() {
        "not_applicable_immutable_owned_slice"
    } else if matches!(operation, Operation::OpcEagerOpen | Operation::OpcEagerSave) {
        "not_applicable_eager_opc"
    } else if operation.is_pptx() {
        if operation.is_source_pptx() {
            "untimed_source_replay_only"
        } else {
            "not_applicable_eager_pptx"
        }
    } else if operation.is_docx() {
        if operation.is_source_docx() {
            "untimed_source_replay_only"
        } else {
            "not_applicable_eager_docx"
        }
    } else if operation.is_xlsx() {
        "not_applicable_filesystem_xlsx"
    } else {
        "timed_read_at"
    }
    .to_owned();
    let pptx_source_replay = operation
        .is_source_pptx()
        .then(|| replay_pptx_source(&source, operation))
        .transpose()?;
    let docx_source_replay = operation
        .is_source_docx()
        .then(|| replay_docx_source(&source, operation))
        .transpose()?;

    // Correctness and hashing are intentionally after the timed operation and
    // after the operation-only counters have been sampled.
    let corpus = filesystem_corpus(operation)?;
    let xlsx_evidence = if operation.is_xlsx() {
        let deferred = deferred_xlsx_operation
            .as_ref()
            .ok_or("XLSX child operation did not retain its timed workbook")?;
        Some(verify_xlsx_operation(
            operation,
            &source,
            &corpus,
            matches!(mode, ChildMode::ColdVerified | ChildMode::VerifiedPrime),
            deferred,
        )?)
    } else {
        None
    };
    verify_child_output(
        operation,
        &source,
        &destination,
        &corpus,
        matches!(mode, ChildMode::ColdVerified | ChildMode::VerifiedPrime),
    )?;
    if let Some(deferred) = deferred_xlsx_operation.take() {
        // Only release the exact timed workbook and lifecycle projection after
        // their correctness validation and every operation-only snapshot.
        std::hint::black_box(deferred.timed_projection);
        std::hint::black_box(deferred.workbook);
    }
    let output = operation
        .is_save()
        .then(|| fs::read(&destination))
        .transpose()?;
    let output_sha256 = output.as_deref().map(super::sha256_hex);
    let output_bytes = output
        .as_ref()
        .map(|bytes| u64::try_from(bytes.len()))
        .transpose()?;
    let result = ChildResult {
        elapsed_ns,
        logical_read_counter_scope,
        logical_read_calls: snapshot.calls,
        logical_read_requested_bytes: snapshot.requested_bytes,
        logical_read_bytes: snapshot.returned_bytes,
        logical_read_largest_requested_bytes: snapshot.largest_requested_bytes,
        logical_read_largest_returned_bytes: snapshot.largest_returned_bytes,
        logical_read_pattern: snapshot.pattern,
        max_concurrent_reads: snapshot.max_concurrent,
        logical_read_request_sizes: snapshot.request_sizes,
        logical_read_request_size_buckets: snapshot.request_size_buckets,
        cold_advice,
        cold_verified,
        process_metrics: process_delta,
        allocation_metrics,
        output_sha256,
        output_bytes,
        opc_materialized_parts: details.opc_materialized_parts,
        cfb_changed_spans: details.cfb_changed_spans,
        cfb_published_bytes: details.cfb_published_bytes,
        cfb_phases: details.cfb_phases,
        cfb_owned: details.cfb_owned,
        pptx_source_replay,
        docx_source_replay,
        xlsx_source_sha256: xlsx_evidence.as_ref().map(|value| value.0.clone()),
        xlsx_semantic_sha256: xlsx_evidence.map(|value| value.1),
    };
    serde_json::to_writer(io::stdout().lock(), &result)?;
    Ok(true)
}

/// Rebuilds the deterministic oracle only after the measured operation and
/// procfs snapshots have completed. The child therefore does not retain the
/// multi-megabyte synthetic corpus while collecting VmHWM.
fn filesystem_corpus(operation: Operation) -> Result<super::Corpus, Box<dyn Error>> {
    if operation.is_pptx() {
        return super::build_pptx_source_edit_corpus();
    }
    if operation.is_docx() {
        return super::build_docx_source_edit_corpus();
    }
    if operation.is_xlsx() {
        return super::build_xlsx_cell_crud_corpus(super::XlsxCellCrudShape::Medium);
    }
    let opc = super::build_opc_corpus(OPC_FILE_SHAPE, OPC_FILE_PAYLOAD)?;
    if operation.is_cfb() {
        super::build_ole_common_corpus(&opc)
    } else {
        Ok(opc)
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct PptxReplayCounters {
    read_calls: u64,
    read_bytes: u64,
    slide_payload_read_calls: u64,
    slide_payload_read_bytes: u64,
    selected_slide_payload_read_calls: u64,
    selected_slide_payload_read_bytes: u64,
    unselected_slide_payload_read_calls: u64,
    unselected_slide_payload_read_bytes: u64,
    media_payload_read_calls: u64,
    media_payload_read_bytes: u64,
}

#[derive(Clone, Debug, Default)]
struct PptxReplayCoverage {
    slide: Vec<Range<u64>>,
    selected: Vec<Range<u64>>,
    media: Vec<Range<u64>>,
}

#[derive(Debug)]
struct PptxReplaySource {
    bytes: Arc<Vec<u8>>,
    version: SourceVersion,
    slide_ranges: Vec<Range<u64>>,
    media_ranges: Vec<Range<u64>>,
    selected_slide_range: Option<Range<u64>>,
    counters: Mutex<PptxReplayCounters>,
    coverage: Mutex<PptxReplayCoverage>,
    return_sizes: Mutex<Vec<u64>>,
}

impl PptxReplaySource {
    fn new(
        bytes: Arc<Vec<u8>>,
        slide_ranges: Vec<Range<u64>>,
        media_ranges: Vec<Range<u64>>,
        selected_slide_range: Option<Range<u64>>,
    ) -> Self {
        Self {
            bytes,
            version: SourceVersion::new(
                NEXT_PPTX_REPLAY_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
                0,
            ),
            slide_ranges,
            media_ranges,
            selected_slide_range,
            counters: Mutex::new(PptxReplayCounters::default()),
            coverage: Mutex::new(PptxReplayCoverage::default()),
            return_sizes: Mutex::new(Vec::new()),
        }
    }

    fn record(&self, offset: u64, count: usize) -> io::Result<()> {
        let end = offset
            .checked_add(
                u64::try_from(count)
                    .map_err(|_| io::Error::other("PPTX replay read length does not fit u64"))?,
            )
            .ok_or_else(|| io::Error::other("PPTX replay read range overflows u64"))?;
        let request = offset..end;
        let mut counters = self
            .counters
            .lock()
            .map_err(|_| io::Error::other("PPTX replay counters are poisoned"))?;
        let count_u64 = u64::try_from(count)
            .map_err(|_| io::Error::other("PPTX replay read length does not fit u64"))?;
        checked_counter_add(&mut counters.read_calls, 1, "read calls")?;
        checked_counter_add(&mut counters.read_bytes, count_u64, "read bytes")?;
        let slide_overlap = overlap_with_ranges(&request, &self.slide_ranges);
        let selected_overlap = self
            .selected_slide_range
            .as_ref()
            .map_or(0, |range| overlap_len(&request, range));
        let unselected_overlap = slide_overlap.saturating_sub(selected_overlap);
        let media_overlap = overlap_with_ranges(&request, &self.media_ranges);
        if slide_overlap != 0 {
            checked_counter_add(
                &mut counters.slide_payload_read_calls,
                1,
                "slide payload read calls",
            )?;
            checked_counter_add(
                &mut counters.slide_payload_read_bytes,
                slide_overlap,
                "slide payload read bytes",
            )?;
        }
        if selected_overlap != 0 {
            checked_counter_add(
                &mut counters.selected_slide_payload_read_calls,
                1,
                "selected slide payload read calls",
            )?;
            checked_counter_add(
                &mut counters.selected_slide_payload_read_bytes,
                selected_overlap,
                "selected slide payload read bytes",
            )?;
        }
        if unselected_overlap != 0 {
            checked_counter_add(
                &mut counters.unselected_slide_payload_read_calls,
                1,
                "unselected slide payload read calls",
            )?;
            checked_counter_add(
                &mut counters.unselected_slide_payload_read_bytes,
                unselected_overlap,
                "unselected slide payload read bytes",
            )?;
        }
        if media_overlap != 0 {
            checked_counter_add(
                &mut counters.media_payload_read_calls,
                1,
                "media payload read calls",
            )?;
            checked_counter_add(
                &mut counters.media_payload_read_bytes,
                media_overlap,
                "media payload read bytes",
            )?;
        }
        drop(counters);
        let mut coverage = self
            .coverage
            .lock()
            .map_err(|_| io::Error::other("PPTX replay coverage is poisoned"))?;
        append_overlaps(&request, &self.slide_ranges, &mut coverage.slide);
        if let Some(range) = self.selected_slide_range.as_ref()
            && let Some(overlap) = overlap_range(&request, range)
        {
            coverage.selected.push(overlap);
        }
        append_overlaps(&request, &self.media_ranges, &mut coverage.media);
        drop(coverage);
        self.return_sizes
            .lock()
            .map_err(|_| io::Error::other("PPTX replay return sizes are poisoned"))?
            .push(count_u64);
        Ok(())
    }

    fn snapshot(&self) -> io::Result<(PptxReplayCounters, PptxReplayCoverage, Vec<u64>)> {
        let counters = *self
            .counters
            .lock()
            .map_err(|_| io::Error::other("PPTX replay counters are poisoned"))?;
        let coverage = self
            .coverage
            .lock()
            .map_err(|_| io::Error::other("PPTX replay coverage is poisoned"))?
            .clone();
        let mut return_sizes = self
            .return_sizes
            .lock()
            .map_err(|_| io::Error::other("PPTX replay return sizes are poisoned"))?
            .clone();
        return_sizes.sort_unstable();
        Ok((counters, coverage, return_sizes))
    }
}

impl ReadAt for PptxReplaySource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("PPTX replay source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = match usize::try_from(offset) {
            Ok(start) => start,
            Err(_) => {
                self.record(offset, 0)?;
                return Ok(0);
            },
        };
        let count = self
            .bytes
            .get(start..)
            .map_or(0, |remaining| remaining.len().min(output.len()));
        if count != 0 {
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
        }
        self.record(offset, count)?;
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

fn checked_counter_add(value: &mut u64, amount: u64, label: &str) -> io::Result<()> {
    *value = value
        .checked_add(amount)
        .ok_or_else(|| io::Error::other(format!("source replay {label} overflow")))?;
    Ok(())
}

fn overlap_range(left: &Range<u64>, right: &Range<u64>) -> Option<Range<u64>> {
    let start = left.start.max(right.start);
    let end = left.end.min(right.end);
    (start < end).then_some(start..end)
}

fn append_overlaps(request: &Range<u64>, ranges: &[Range<u64>], output: &mut Vec<Range<u64>>) {
    for range in ranges {
        if let Some(overlap) = overlap_range(request, range) {
            output.push(overlap);
        }
    }
}

fn overlap_len(left: &Range<u64>, right: &Range<u64>) -> u64 {
    overlap_range(left, right).map_or(0, |range| range.end - range.start)
}

fn overlap_with_ranges(request: &Range<u64>, ranges: &[Range<u64>]) -> u64 {
    ranges.iter().map(|range| overlap_len(request, range)).sum()
}

fn merged_ranges(mut ranges: Vec<Range<u64>>) -> Vec<Range<u64>> {
    ranges.sort_by_key(|range| (range.start, range.end));
    let mut merged: Vec<Range<u64>> = Vec::with_capacity(ranges.len());
    for range in ranges {
        if let Some(previous) = merged.last_mut()
            && range.start <= previous.end
        {
            previous.end = previous.end.max(range.end);
            continue;
        }
        merged.push(range);
    }
    merged
}

fn covered_bytes(expected: &[Range<u64>], observed: &[Range<u64>]) -> Result<u64, Box<dyn Error>> {
    let observed = merged_ranges(observed.to_vec());
    expected
        .iter()
        .flat_map(|expected| {
            observed
                .iter()
                .filter_map(move |observed| overlap_range(expected, observed))
        })
        .try_fold(0_u64, |total, overlap| {
            total
                .checked_add(overlap.end - overlap.start)
                .ok_or_else(|| "PPTX replay coverage byte count overflows u64".into())
        })
}

fn range_fully_covered(target: &Range<u64>, observed: &[Range<u64>]) -> bool {
    let mut cursor = target.start;
    for range in merged_ranges(observed.to_vec()) {
        if range.end <= cursor {
            continue;
        }
        if range.start > cursor {
            return false;
        }
        cursor = range.end;
        if cursor >= target.end {
            return true;
        }
    }
    cursor >= target.end
}

fn fully_covered_range_count(expected: &[Range<u64>], observed: &[Range<u64>]) -> u64 {
    u64::try_from(
        expected
            .iter()
            .filter(|range| range_fully_covered(range, observed))
            .count(),
    )
    .expect("PPTX replay range count fits u64")
}

fn pptx_slide_part_position(name: &str) -> Option<usize> {
    let number = name
        .strip_prefix("ppt/slides/slide")?
        .strip_suffix(".xml")?
        .parse::<usize>()
        .ok()?;
    number.checked_sub(1)
}

fn pptx_payload_ranges(bytes: &[u8]) -> Result<(Vec<Range<u64>>, Vec<Range<u64>>), Box<dyn Error>> {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes)?;
    // The source facade's position is relationship order. This fixed corpus
    // writes that order as slide1.xml .. slide200.xml, so retain that mapping
    // by part name and never substitute physical ZIP offset order for it.
    let mut slide_slots: Vec<Option<Range<u64>>> =
        (0..super::PPTX_SOURCE_SLIDE_COUNT).map(|_| None).collect();
    let mut media = Vec::new();
    for header in archive.entries() {
        let header = header?;
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        let range = start..end;
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            let position = pptx_slide_part_position(&name)
                .ok_or_else(|| format!("PPTX replay found an invalid slide part name: {name}"))?;
            let slot = slide_slots
                .get_mut(position)
                .ok_or_else(|| format!("PPTX replay found out-of-range slide part name: {name}"))?;
            if slot.replace(range).is_some() {
                return Err(format!("PPTX replay found duplicate slide part name: {name}").into());
            }
        } else if name.starts_with("ppt/media/") {
            media.push(range);
        }
    }
    let slides = slide_slots
        .into_iter()
        .enumerate()
        .map(|(position, range)| {
            range.ok_or_else(|| {
                format!(
                    "PPTX replay is missing slide{}.xml for source position {position}",
                    position + 1
                )
                .into()
            })
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    media.sort_by_key(|range| range.start);
    if slides.len() != super::PPTX_SOURCE_SLIDE_COUNT {
        return Err(format!(
            "PPTX replay found {} slide payload ranges, expected {}",
            slides.len(),
            super::PPTX_SOURCE_SLIDE_COUNT
        )
        .into());
    }
    if media.len() != super::PPTX_SOURCE_MEDIA_ENTRY_COUNT {
        return Err(format!(
            "PPTX replay found {} media payload ranges, expected {}",
            media.len(),
            super::PPTX_SOURCE_MEDIA_ENTRY_COUNT
        )
        .into());
    }
    Ok((slides, media))
}

fn replay_pptx_source(
    source: &Path,
    operation: Operation,
) -> Result<PptxSourceReplayEvidence, Box<dyn Error>> {
    let bytes = Arc::new(fs::read(source)?);
    let (slide_ranges, media_ranges) = pptx_payload_ranges(&bytes)?;
    let selected_slide_range = slide_ranges
        .get(PPTX_FILE_SELECTED_POSITION)
        .cloned()
        .ok_or("PPTX replay selected slide position is outside the corpus")?;
    let replay = Arc::new(PptxReplaySource::new(
        Arc::clone(&bytes),
        slide_ranges.clone(),
        media_ranges.clone(),
        Some(selected_slide_range.clone()),
    ));
    let presentation = litchi_pptx::SourceBackedPresentation::from_read_at(replay.clone())?;
    let slide_count = presentation.slide_count();
    let mut semantic = Sha256::new();
    match operation {
        Operation::PptxSourceOpen => {},
        Operation::PptxSourceSlideCount | Operation::PptxSourceOpenSlideCountLifecycle => {
            if slide_count != super::PPTX_SOURCE_SLIDE_COUNT {
                return Err("PPTX source replay slide count differs from corpus".into());
            }
        },
        Operation::PptxSourceSelectedSlide | Operation::PptxSourceOpenSelectedSlideLifecycle => {
            let slide = presentation
                .slide(PPTX_FILE_SELECTED_POSITION)
                .ok_or("PPTX source replay selected slide is missing")?;
            let (text, name) = slide.text_and_name()?;
            semantic.update(text.as_bytes());
            semantic.update([0]);
            semantic.update(name.as_bytes());
        },
        Operation::PptxSourceListSlides => {
            for slide in presentation.slides() {
                let (text, name) = slide.text_and_name()?;
                semantic.update(text.as_bytes());
                semantic.update([0]);
                semantic.update(name.as_bytes());
            }
        },
        _ => return Err("non-source PPTX operation passed to source replay".into()),
    }
    let (counters, coverage, return_sizes) = replay.snapshot()?;
    let unselected_slide_ranges = slide_ranges
        .iter()
        .enumerate()
        .filter(|(position, _)| *position != PPTX_FILE_SELECTED_POSITION)
        .map(|(_, range)| range.clone())
        .collect::<Vec<_>>();
    let slide_payload_covered_bytes = covered_bytes(&slide_ranges, &coverage.slide)?;
    let selected_slide_payload_covered_bytes = covered_bytes(
        std::slice::from_ref(&selected_slide_range),
        &coverage.selected,
    )?;
    let unselected_slide_payload_covered_bytes =
        covered_bytes(&unselected_slide_ranges, &coverage.slide)?;
    let media_payload_covered_bytes = covered_bytes(&media_ranges, &coverage.media)?;
    let slide_payload_ranges_fully_covered =
        fully_covered_range_count(&slide_ranges, &coverage.slide);
    let selected_slide_payload_fully_covered =
        range_fully_covered(&selected_slide_range, &coverage.selected);
    let slide_payload_total_bytes = covered_bytes(&slide_ranges, &slide_ranges)?;
    let slide_range_count = u64::try_from(slide_ranges.len())?;
    let classification = match operation {
        Operation::PptxSourceOpen
        | Operation::PptxSourceSlideCount
        | Operation::PptxSourceOpenSlideCountLifecycle
            if counters.slide_payload_read_bytes == 0
                && counters.media_payload_read_bytes == 0
                && slide_payload_covered_bytes == 0
                && media_payload_covered_bytes == 0 =>
        {
            "catalog-only:zero-slide-and-media-overlap"
        },
        Operation::PptxSourceSelectedSlide | Operation::PptxSourceOpenSelectedSlideLifecycle
            if selected_slide_payload_fully_covered
                && selected_slide_payload_covered_bytes
                    == selected_slide_range.end - selected_slide_range.start
                && counters.selected_slide_payload_read_bytes != 0
                && unselected_slide_payload_covered_bytes == 0
                && media_payload_covered_bytes == 0
                && counters.unselected_slide_payload_read_bytes == 0
                && counters.media_payload_read_bytes == 0 =>
        {
            "selected-slide-only:target-slide-no-unselected-or-media-overlap"
        },
        Operation::PptxSourceListSlides
            if slide_payload_ranges_fully_covered == slide_range_count
                && slide_payload_covered_bytes == slide_payload_total_bytes
                && media_payload_covered_bytes == 0
                && counters.media_payload_read_bytes == 0 =>
        {
            "list-slides:all-slide-payloads-no-media-overlap"
        },
        _ => "classification-failed",
    }
    .to_owned();
    if classification == "classification-failed" {
        return Err(format!(
            "PPTX source replay violated {} payload-range classification",
            operation.case().name()
        )
        .into());
    }
    Ok(PptxSourceReplayEvidence {
        implementation: "litchi_pptx::SourceBackedPresentation".to_owned(),
        operation: operation.pptx_query_name().unwrap_or("open").to_owned(),
        source_bytes: u64::try_from(bytes.len())?,
        source_sha256: super::sha256_hex(&bytes),
        slide_count,
        selected_position: (operation.is_pptx_query()
            || matches!(operation, Operation::PptxSourceOpenSelectedSlideLifecycle))
        .then_some(PPTX_FILE_SELECTED_POSITION),
        read_calls: counters.read_calls,
        read_bytes: counters.read_bytes,
        read_return_sizes: return_sizes,
        slide_payload_read_calls: counters.slide_payload_read_calls,
        slide_payload_read_bytes: counters.slide_payload_read_bytes,
        slide_payload_covered_bytes,
        slide_payload_ranges_fully_covered,
        selected_slide_payload_read_calls: counters.selected_slide_payload_read_calls,
        selected_slide_payload_read_bytes: counters.selected_slide_payload_read_bytes,
        selected_slide_payload_covered_bytes,
        selected_slide_payload_fully_covered,
        unselected_slide_payload_read_calls: counters.unselected_slide_payload_read_calls,
        unselected_slide_payload_read_bytes: counters.unselected_slide_payload_read_bytes,
        unselected_slide_payload_covered_bytes,
        media_payload_read_calls: counters.media_payload_read_calls,
        media_payload_read_bytes: counters.media_payload_read_bytes,
        media_payload_covered_bytes,
        semantic_sha256: super::sha256_hex(semantic.finalize().as_slice()),
        classification,
    })
}

#[derive(Clone, Copy, Debug, Default)]
struct DocxReplayCounters {
    read_calls: u64,
    read_bytes: u64,
    main_payload_overlap_bytes: u64,
    media_payload_overlap_bytes: u64,
    unselected_payload_overlap_bytes: u64,
    core_payload_overlap_bytes: u64,
}

#[derive(Clone, Debug)]
struct DocxReplaySnapshot {
    counters: DocxReplayCounters,
    return_sizes: Vec<u64>,
    read_ranges: Vec<Range<u64>>,
}

#[derive(Clone, Debug)]
struct DocxReplayRanges {
    main: Range<u64>,
    media: Vec<Range<u64>>,
    unselected: Vec<Range<u64>>,
    core: Vec<Range<u64>>,
}

#[derive(Debug)]
struct DocxReplaySource {
    bytes: Arc<Vec<u8>>,
    version: SourceVersion,
    ranges: DocxReplayRanges,
    counters: Mutex<DocxReplayCounters>,
    return_sizes: Mutex<Vec<u64>>,
    read_ranges: Mutex<Vec<Range<u64>>>,
}

impl DocxReplaySource {
    fn new(bytes: Arc<Vec<u8>>, ranges: DocxReplayRanges) -> Self {
        Self {
            bytes,
            version: SourceVersion::new(
                NEXT_DOCX_REPLAY_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
                0,
            ),
            ranges,
            counters: Mutex::new(DocxReplayCounters::default()),
            return_sizes: Mutex::new(Vec::new()),
            read_ranges: Mutex::new(Vec::new()),
        }
    }

    fn record(&self, offset: u64, count: usize) -> io::Result<()> {
        let count = u64::try_from(count)
            .map_err(|_| io::Error::other("DOCX replay read length does not fit u64"))?;
        let end = offset
            .checked_add(count)
            .ok_or_else(|| io::Error::other("DOCX replay read range overflows u64"))?;
        let request = offset..end;
        let mut counters = self
            .counters
            .lock()
            .map_err(|_| io::Error::other("DOCX replay counters are poisoned"))?;
        checked_counter_add(&mut counters.read_calls, 1, "DOCX read calls")?;
        checked_counter_add(&mut counters.read_bytes, count, "DOCX read bytes")?;
        checked_counter_add(
            &mut counters.main_payload_overlap_bytes,
            overlap_len(&request, &self.ranges.main),
            "DOCX main payload overlap",
        )?;
        checked_counter_add(
            &mut counters.media_payload_overlap_bytes,
            overlap_with_ranges(&request, &self.ranges.media),
            "DOCX media payload overlap",
        )?;
        checked_counter_add(
            &mut counters.unselected_payload_overlap_bytes,
            overlap_with_ranges(&request, &self.ranges.unselected),
            "DOCX unselected payload overlap",
        )?;
        checked_counter_add(
            &mut counters.core_payload_overlap_bytes,
            overlap_with_ranges(&request, &self.ranges.core),
            "DOCX core payload overlap",
        )?;
        drop(counters);
        self.return_sizes
            .lock()
            .map_err(|_| io::Error::other("DOCX replay return sizes are poisoned"))?
            .push(count);
        self.read_ranges
            .lock()
            .map_err(|_| io::Error::other("DOCX replay read ranges are poisoned"))?
            .push(request);
        Ok(())
    }

    fn snapshot(&self) -> io::Result<DocxReplaySnapshot> {
        let counters = *self
            .counters
            .lock()
            .map_err(|_| io::Error::other("DOCX replay counters are poisoned"))?;
        let return_sizes = self
            .return_sizes
            .lock()
            .map_err(|_| io::Error::other("DOCX replay return sizes are poisoned"))?
            .clone();
        let read_ranges = self
            .read_ranges
            .lock()
            .map_err(|_| io::Error::other("DOCX replay read ranges are poisoned"))?
            .clone();
        Ok(DocxReplaySnapshot {
            counters,
            return_sizes,
            read_ranges,
        })
    }
}

impl ReadAt for DocxReplaySource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("DOCX replay source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = match usize::try_from(offset) {
            Ok(start) => start,
            Err(_) => {
                self.record(offset, 0)?;
                return Ok(0);
            },
        };
        let count = self
            .bytes
            .get(start..)
            .map_or(0, |remaining| remaining.len().min(output.len()));
        if count != 0 {
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
        }
        self.record(offset, count)?;
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

fn docx_replay_phase(
    before: &DocxReplaySnapshot,
    after: &DocxReplaySnapshot,
    ranges: &DocxReplayRanges,
) -> DocxReplayPhase {
    let counters = DocxReplayCounters {
        read_calls: after.counters.read_calls - before.counters.read_calls,
        read_bytes: after.counters.read_bytes - before.counters.read_bytes,
        main_payload_overlap_bytes: after.counters.main_payload_overlap_bytes
            - before.counters.main_payload_overlap_bytes,
        media_payload_overlap_bytes: after.counters.media_payload_overlap_bytes
            - before.counters.media_payload_overlap_bytes,
        unselected_payload_overlap_bytes: after.counters.unselected_payload_overlap_bytes
            - before.counters.unselected_payload_overlap_bytes,
        core_payload_overlap_bytes: after.counters.core_payload_overlap_bytes
            - before.counters.core_payload_overlap_bytes,
    };
    let return_sizes = after.return_sizes[before.return_sizes.len()..].to_vec();
    let read_ranges = &after.read_ranges[before.read_ranges.len()..];
    let main_payload_covered_bytes =
        covered_bytes(std::slice::from_ref(&ranges.main), read_ranges).unwrap_or_default();
    let media_payload_covered_bytes = covered_bytes(&ranges.media, read_ranges).unwrap_or_default();
    let unselected_payload_covered_bytes =
        covered_bytes(&ranges.unselected, read_ranges).unwrap_or_default();
    let core_payload_covered_bytes = covered_bytes(&ranges.core, read_ranges).unwrap_or_default();
    let main_payload_fully_covered = range_fully_covered(&ranges.main, read_ranges);
    DocxReplayPhase {
        counters,
        return_sizes,
        main_payload_covered_bytes,
        main_payload_fully_covered,
        media_payload_covered_bytes,
        unselected_payload_covered_bytes,
        core_payload_covered_bytes,
    }
}

#[derive(Clone, Debug)]
struct DocxReplayPhase {
    counters: DocxReplayCounters,
    return_sizes: Vec<u64>,
    main_payload_covered_bytes: u64,
    main_payload_fully_covered: bool,
    media_payload_covered_bytes: u64,
    unselected_payload_covered_bytes: u64,
    core_payload_covered_bytes: u64,
}

fn docx_replay_ranges(bytes: &[u8]) -> Result<DocxReplayRanges, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut main = None;
    let mut media = Vec::new();
    let mut unselected = Vec::new();
    let mut core = Vec::new();
    for header in archive.entries() {
        let header = header?;
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        let range = start..end;
        if name == "word/document.xml" {
            if main.replace(range).is_some() {
                return Err("DOCX replay found duplicate main-document member".into());
            }
        } else if name.starts_with("word/media/") {
            media.push(range);
        } else if name == "docProps/core.xml" {
            core.push(range);
        } else if name != "[Content_Types].xml" && !name.ends_with(".rels") {
            unselected.push(range);
        }
    }
    Ok(DocxReplayRanges {
        main: main.ok_or("DOCX replay is missing word/document.xml")?,
        media,
        unselected,
        core,
    })
}

fn replay_docx_source(
    source: &Path,
    operation: Operation,
) -> Result<DocxSourceReplayEvidence, Box<dyn Error>> {
    let bytes = Arc::new(fs::read(source)?);
    let ranges = docx_replay_ranges(&bytes)?;
    let replay = Arc::new(DocxReplaySource::new(Arc::clone(&bytes), ranges.clone()));
    let package = litchi_docx::source_backed::Package::from_read_at(replay.clone())?;
    let open = replay.snapshot()?;
    let mut semantic = Sha256::new();
    let mut paragraph_count = super::SemanticShape::Medium.docx_paragraphs();
    let (preparation, query) = if operation.is_docx_query() || operation.is_docx_lifecycle() {
        let document = package.document()?;
        let prepared = replay.snapshot()?;
        match operation {
            Operation::DocxSourceParagraphCount
            | Operation::DocxSourceOpenParagraphCountLifecycle => {
                paragraph_count = document.paragraph_count()?;
                semantic.update(paragraph_count.to_le_bytes());
            },
            Operation::DocxSourceListParagraphs => {
                let paragraphs = document.paragraphs()?;
                paragraph_count = paragraphs.len();
                for paragraph in paragraphs {
                    let text = paragraph.text()?;
                    semantic.update(text.as_bytes());
                    semantic.update([0]);
                }
            },
            Operation::DocxSourceFullText | Operation::DocxSourceOpenFullTextLifecycle => {
                let text = document.extract_text()?;
                paragraph_count = document.paragraph_count()?;
                semantic.update(text.as_bytes());
            },
            _ => return Err("non-query DOCX operation passed to source replay".into()),
        }
        let queried = replay.snapshot()?;
        (
            docx_replay_phase(&open, &prepared, &ranges),
            docx_replay_phase(&prepared, &queried, &ranges),
        )
    } else {
        (
            DocxReplayPhase {
                counters: DocxReplayCounters::default(),
                return_sizes: Vec::new(),
                main_payload_covered_bytes: 0,
                main_payload_fully_covered: false,
                media_payload_covered_bytes: 0,
                unselected_payload_covered_bytes: 0,
                core_payload_covered_bytes: 0,
            },
            DocxReplayPhase {
                counters: DocxReplayCounters::default(),
                return_sizes: Vec::new(),
                main_payload_covered_bytes: 0,
                main_payload_fully_covered: false,
                media_payload_covered_bytes: 0,
                unselected_payload_covered_bytes: 0,
                core_payload_covered_bytes: 0,
            },
        )
    };
    let open_phase = docx_replay_phase(
        &DocxReplaySnapshot {
            counters: DocxReplayCounters::default(),
            return_sizes: Vec::new(),
            read_ranges: Vec::new(),
        },
        &open,
        &ranges,
    );
    let diagnostics = package.cache_diagnostics();
    let classification = if open_phase.counters.main_payload_overlap_bytes == 0
        && open_phase.counters.media_payload_overlap_bytes == 0
        && open_phase.counters.unselected_payload_overlap_bytes == 0
        && open_phase.counters.core_payload_overlap_bytes == 0
        && (!(operation.is_docx_query() || operation.is_docx_lifecycle())
            || (preparation.counters.main_payload_overlap_bytes != 0
                && preparation.main_payload_fully_covered
                && preparation.counters.media_payload_overlap_bytes == 0
                && preparation.counters.unselected_payload_overlap_bytes == 0
                && preparation.counters.core_payload_overlap_bytes == 0
                && query.counters.main_payload_overlap_bytes == 0
                && query.counters.media_payload_overlap_bytes == 0
                && query.counters.unselected_payload_overlap_bytes == 0
                && query.counters.core_payload_overlap_bytes == 0))
    {
        if operation.is_docx_query() || operation.is_docx_lifecycle() {
            "semantic-query:one-complete-main-range-preparation-zero-query-unselected-media-core"
        } else {
            "catalog-only:zero-main-media-unselected-core-overlap"
        }
    } else {
        "classification-failed"
    }
    .to_owned();
    if classification == "classification-failed" {
        return Err(format!(
            "DOCX source replay violated {} payload-range classification",
            operation.case().name()
        )
        .into());
    }
    let semantic_digest = semantic.finalize();
    Ok(DocxSourceReplayEvidence {
        implementation: "litchi_docx::source_backed::Package".to_owned(),
        operation: operation.docx_query_name().unwrap_or("open").to_owned(),
        source_bytes: u64::try_from(bytes.len())?,
        source_sha256: super::sha256_hex(&bytes),
        paragraph_count,
        open_read_calls: open_phase.counters.read_calls,
        open_read_bytes: open_phase.counters.read_bytes,
        open_read_return_sizes: open_phase.return_sizes,
        open_main_payload_overlap_bytes: open_phase.counters.main_payload_overlap_bytes,
        open_media_payload_overlap_bytes: open_phase.counters.media_payload_overlap_bytes,
        open_unselected_payload_overlap_bytes: open_phase.counters.unselected_payload_overlap_bytes,
        open_core_payload_overlap_bytes: open_phase.counters.core_payload_overlap_bytes,
        open_main_payload_covered_bytes: open_phase.main_payload_covered_bytes,
        open_media_payload_covered_bytes: open_phase.media_payload_covered_bytes,
        open_unselected_payload_covered_bytes: open_phase.unselected_payload_covered_bytes,
        open_core_payload_covered_bytes: open_phase.core_payload_covered_bytes,
        preparation_read_calls: preparation.counters.read_calls,
        preparation_read_bytes: preparation.counters.read_bytes,
        preparation_read_return_sizes: preparation.return_sizes,
        preparation_main_payload_overlap_bytes: preparation.counters.main_payload_overlap_bytes,
        preparation_main_payload_covered_bytes: preparation.main_payload_covered_bytes,
        preparation_main_payload_fully_covered: preparation.main_payload_fully_covered,
        preparation_media_payload_overlap_bytes: preparation.counters.media_payload_overlap_bytes,
        preparation_unselected_payload_overlap_bytes: preparation
            .counters
            .unselected_payload_overlap_bytes,
        preparation_core_payload_overlap_bytes: preparation.counters.core_payload_overlap_bytes,
        preparation_media_payload_covered_bytes: preparation.media_payload_covered_bytes,
        preparation_unselected_payload_covered_bytes: preparation.unselected_payload_covered_bytes,
        preparation_core_payload_covered_bytes: preparation.core_payload_covered_bytes,
        query_read_calls: query.counters.read_calls,
        query_read_bytes: query.counters.read_bytes,
        query_read_return_sizes: query.return_sizes,
        query_main_payload_overlap_bytes: query.counters.main_payload_overlap_bytes,
        query_media_payload_overlap_bytes: query.counters.media_payload_overlap_bytes,
        query_unselected_payload_overlap_bytes: query.counters.unselected_payload_overlap_bytes,
        query_core_payload_overlap_bytes: query.counters.core_payload_overlap_bytes,
        query_main_payload_covered_bytes: query.main_payload_covered_bytes,
        query_media_payload_covered_bytes: query.media_payload_covered_bytes,
        query_unselected_payload_covered_bytes: query.unselected_payload_covered_bytes,
        query_core_payload_covered_bytes: query.core_payload_covered_bytes,
        materializations: diagnostics.successful_loads,
        semantic_sha256: super::sha256_hex(&semantic_digest[..]),
        classification,
    })
}

#[derive(Clone, Debug, Default)]
struct ReadMetrics {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    largest_requested_bytes: u64,
    largest_returned_bytes: u64,
    pattern: Option<ReadPattern>,
    max_concurrent: u64,
    request_sizes: Vec<u64>,
    request_size_buckets: ReadSizeBuckets,
}

#[derive(Clone, Copy, Debug, Default)]
struct ReadTotals {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

/// Harness-only logical source wrapper.
///
/// A call and its requested output length are recorded before delegating to
/// the wrapped source. Returned bytes are recorded only after the source's
/// successful result is validated. An underlying I/O error propagates after
/// recording a requested-only attempt; an over-returning source is rejected
/// and marks the metric snapshot unavailable. Either error fails the sample,
/// and neither attempt contributes returned bytes. These counters describe
/// the wrapper boundary, not physical I/O.
struct CountingReadAt {
    inner: Arc<dyn ReadAt>,
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    largest_requested_bytes: AtomicU64,
    largest_returned_bytes: AtomicU64,
    in_flight: AtomicU64,
    max_concurrent: AtomicU64,
    metrics_failed: AtomicBool,
    request_sizes: Mutex<Vec<u64>>,
    pattern: Mutex<ReadPatternState>,
}

#[derive(Debug, Default)]
struct ReadPatternState {
    previous_end: Option<u64>,
    observations: u64,
    non_contiguous: bool,
    unknown: bool,
}

impl ReadPatternState {
    fn observe(&mut self, offset: u64, requested: u64, returned: usize) -> io::Result<()> {
        let returned = u64::try_from(returned)
            .map_err(|_| io::Error::other("returned source range does not fit u64"))?;
        if returned == 0 || returned != requested {
            self.unknown = true;
            return Ok(());
        }
        if let Some(previous_end) = self.previous_end {
            if previous_end != offset {
                self.non_contiguous = true;
            }
        }
        self.previous_end = Some(
            offset
                .checked_add(returned)
                .ok_or_else(|| io::Error::other("source range end overflows u64"))?,
        );
        self.observations = self
            .observations
            .checked_add(1)
            .ok_or_else(|| io::Error::other("source range observation count overflows u64"))?;
        Ok(())
    }

    fn classify(&self, max_concurrent: u64) -> ReadPattern {
        if max_concurrent > 1 || self.unknown || self.observations < 2 {
            ReadPattern::Unknown
        } else if self.non_contiguous {
            ReadPattern::Random
        } else {
            ReadPattern::Sequential
        }
    }
}

impl CountingReadAt {
    fn new(inner: Arc<dyn ReadAt>) -> Self {
        Self {
            inner,
            calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            largest_requested_bytes: AtomicU64::new(0),
            largest_returned_bytes: AtomicU64::new(0),
            in_flight: AtomicU64::new(0),
            max_concurrent: AtomicU64::new(0),
            metrics_failed: AtomicBool::new(false),
            request_sizes: Mutex::new(Vec::new()),
            pattern: Mutex::new(ReadPatternState::default()),
        }
    }

    fn fail_metrics(&self, error: io::Error) -> io::Error {
        self.metrics_failed.store(true, Ordering::SeqCst);
        error
    }

    fn checked_add(
        &self,
        counter: &AtomicU64,
        amount: u64,
        label: &'static str,
    ) -> io::Result<u64> {
        checked_atomic_add(counter, amount, label).map_err(|error| self.fail_metrics(error))
    }

    fn snapshot(&self) -> io::Result<ReadMetrics> {
        if self.metrics_failed.load(Ordering::SeqCst) {
            return Err(io::Error::other(
                "filesystem source metrics are unavailable after a counter failure",
            ));
        }
        let mut request_sizes = self
            .request_sizes
            .lock()
            .map_err(|_| {
                self.fail_metrics(io::Error::other(
                    "filesystem source request sizes are poisoned",
                ))
            })?
            .clone();
        request_sizes.sort_unstable();
        let mut request_size_buckets = ReadSizeBuckets::default();
        for &size in &request_sizes {
            request_size_buckets
                .observe(size)
                .map_err(|error| self.fail_metrics(error))?;
        }
        let max_concurrent = self.max_concurrent.load(Ordering::SeqCst);
        let pattern = self
            .pattern
            .lock()
            .map_err(|_| {
                self.fail_metrics(io::Error::other(
                    "filesystem source pattern metrics are poisoned",
                ))
            })?
            .classify(max_concurrent);
        if self.metrics_failed.load(Ordering::SeqCst) {
            return Err(io::Error::other(
                "filesystem source metrics failed while taking a snapshot",
            ));
        }
        Ok(ReadMetrics {
            calls: self.calls.load(Ordering::SeqCst),
            requested_bytes: self.requested_bytes.load(Ordering::SeqCst),
            returned_bytes: self.returned_bytes.load(Ordering::SeqCst),
            largest_requested_bytes: self.largest_requested_bytes.load(Ordering::SeqCst),
            largest_returned_bytes: self.largest_returned_bytes.load(Ordering::SeqCst),
            pattern: Some(pattern),
            max_concurrent,
            request_sizes,
            request_size_buckets,
        })
    }

    fn totals(&self) -> ReadTotals {
        ReadTotals {
            calls: self.calls.load(Ordering::SeqCst),
            requested_bytes: self.requested_bytes.load(Ordering::SeqCst),
            returned_bytes: self.returned_bytes.load(Ordering::SeqCst),
        }
    }
}

fn checked_atomic_add(counter: &AtomicU64, amount: u64, label: &str) -> io::Result<u64> {
    let mut current = counter.load(Ordering::SeqCst);
    loop {
        let next = current
            .checked_add(amount)
            .ok_or_else(|| io::Error::other(format!("filesystem source {label} overflow")))?;
        match counter.compare_exchange_weak(current, next, Ordering::SeqCst, Ordering::SeqCst) {
            Ok(_) => return Ok(next),
            Err(observed) => current = observed,
        }
    }
}

fn checked_atomic_sub(counter: &AtomicU64, amount: u64, label: &str) -> io::Result<u64> {
    let mut current = counter.load(Ordering::SeqCst);
    loop {
        let next = current
            .checked_sub(amount)
            .ok_or_else(|| io::Error::other(format!("filesystem source {label} underflow")))?;
        match counter.compare_exchange_weak(current, next, Ordering::SeqCst, Ordering::SeqCst) {
            Ok(_) => return Ok(next),
            Err(observed) => current = observed,
        }
    }
}

struct InFlight<'a> {
    counter: &'a AtomicU64,
    metrics_failed: &'a AtomicBool,
}

impl Drop for InFlight<'_> {
    fn drop(&mut self) {
        if checked_atomic_sub(self.counter, 1, "in-flight reads").is_err() {
            self.metrics_failed.store(true, Ordering::SeqCst);
        }
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let requested = u64::try_from(output.len())
            .map_err(|error| self.fail_metrics(io::Error::other(error.to_string())))?;
        self.checked_add(&self.calls, 1, "read calls")?;
        self.checked_add(&self.requested_bytes, requested, "requested bytes")?;
        self.largest_requested_bytes
            .fetch_max(requested, Ordering::SeqCst);
        self.request_sizes
            .lock()
            .map_err(|_| {
                self.fail_metrics(io::Error::other(
                    "filesystem source request sizes are poisoned",
                ))
            })?
            .push(requested);
        let in_flight = self.checked_add(&self.in_flight, 1, "in-flight reads")?;
        self.max_concurrent.fetch_max(in_flight, Ordering::SeqCst);
        let _guard = InFlight {
            counter: &self.in_flight,
            metrics_failed: &self.metrics_failed,
        };
        let read = self.inner.read_at(offset, output)?;
        let returned = u64::try_from(read)
            .map_err(|error| self.fail_metrics(io::Error::other(error.to_string())))?;
        if returned > requested {
            return Err(self.fail_metrics(io::Error::other(
                "filesystem source returned more bytes than requested",
            )));
        }
        self.checked_add(&self.returned_bytes, returned, "returned bytes")?;
        self.largest_returned_bytes
            .fetch_max(returned, Ordering::SeqCst);
        let mut pattern = self.pattern.lock().map_err(|_| {
            self.fail_metrics(io::Error::other(
                "filesystem source pattern metrics are poisoned",
            ))
        })?;
        pattern
            .observe(offset, requested, read)
            .map_err(|error| self.fail_metrics(error))?;
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn request_cold(path: &Path) -> ColdAdvice {
    let file = match File::open(path) {
        Ok(file) => file,
        Err(_) => return ColdAdvice::Failed,
    };
    #[cfg(target_os = "linux")]
    {
        match rustix::fs::fadvise(&file, 0, None, rustix::fs::Advice::DontNeed) {
            Ok(()) => ColdAdvice::Requested,
            Err(_) => ColdAdvice::Failed,
        }
    }
    #[cfg(not(target_os = "linux"))]
    {
        let _ = file;
        ColdAdvice::Unsupported
    }
}

fn run_opc_eager_open(
    source: &Path,
    details: &mut OperationDetails,
) -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
    let package = OpcPackage::from_bytes(&fs::read(source)?)?;
    details.opc_materialized_parts = Some(u64::try_from(package.part_count())?);
    std::hint::black_box(&package);
    Ok(None)
}

fn run_opc_source_open(
    source: &Path,
) -> Result<(Option<Arc<CountingReadAt>>, SourceBackedPackage), Box<dyn Error>> {
    let counter = Arc::new(CountingReadAt::new(Arc::new(FileSource::open(source)?)));
    let package = SourceBackedPackage::from_read_at(counter.clone())?;
    std::hint::black_box(&package);
    Ok((Some(counter), package))
}

fn run_pptx_operation(
    operation: Operation,
    source: &Path,
    prepared: Option<&litchi::Presentation>,
) -> Result<(), Box<dyn Error>> {
    match operation {
        Operation::PptxEagerOpen => {
            let presentation = litchi::Presentation::from_bytes(fs::read(source)?)?;
            std::hint::black_box(presentation);
        },
        Operation::PptxSourceOpen => {
            // This is the candidate path under measurement. The root facade
            // must open the filesystem path directly rather than replaying a
            // caller-owned byte buffer.
            let presentation = litchi::Presentation::open(source)?;
            std::hint::black_box(presentation);
        },
        Operation::PptxEagerListSlides | Operation::PptxSourceListSlides => {
            let presentation = prepared.ok_or("PPTX list-slides operation has no prepared root")?;
            // `slides()` returns owned Slide values. This control deliberately
            // materializes the complete vector and never uses a lazy iterator.
            let slides = presentation.slides()?;
            if slides.len() != super::PPTX_SOURCE_SLIDE_COUNT {
                return Err("PPTX list-slides count differs from fixed corpus".into());
            }
            std::hint::black_box(slides);
        },
        Operation::PptxEagerSlideCount | Operation::PptxSourceSlideCount => {
            let presentation = prepared.ok_or("PPTX slide-count operation has no prepared root")?;
            let count = presentation.slide_count()?;
            if count != super::PPTX_SOURCE_SLIDE_COUNT {
                return Err("PPTX slide-count differs from fixed corpus".into());
            }
            std::hint::black_box(count);
        },
        Operation::PptxEagerSelectedSlide | Operation::PptxSourceSelectedSlide => {
            let presentation =
                prepared.ok_or("PPTX selected-slide operation has no prepared root")?;
            let slide = presentation
                .slide(PPTX_FILE_SELECTED_POSITION)?
                .ok_or("PPTX selected slide is missing")?;
            // Use the selector-first public primitive. In particular, do not
            // replace this with `slides().nth(...)`, which would materialize
            // every slide on the source-backed path.
            std::hint::black_box(slide);
        },
        Operation::PptxEagerOpenSlideCountLifecycle
        | Operation::PptxSourceOpenSlideCountLifecycle => {
            let presentation = if operation.is_source_pptx() {
                litchi::Presentation::open(source)?
            } else {
                litchi::Presentation::from_bytes(fs::read(source)?)?
            };
            let count = presentation.slide_count()?;
            if count != super::PPTX_SOURCE_SLIDE_COUNT {
                return Err("PPTX lifecycle slide-count differs from fixed corpus".into());
            }
            std::hint::black_box((presentation, count));
        },
        Operation::PptxEagerOpenSelectedSlideLifecycle
        | Operation::PptxSourceOpenSelectedSlideLifecycle => {
            let presentation = if operation.is_source_pptx() {
                litchi::Presentation::open(source)?
            } else {
                litchi::Presentation::from_bytes(fs::read(source)?)?
            };
            let slide = presentation
                .slide(PPTX_FILE_SELECTED_POSITION)?
                .ok_or("PPTX lifecycle selected slide is missing")?;
            std::hint::black_box(slide);
            std::hint::black_box(&presentation);
        },
        _ => return Err("non-PPTX operation passed to run_pptx_operation".into()),
    }
    Ok(())
}

#[allow(
    clippy::large_enum_variant,
    reason = "the harness compares the typed eager owner with the unified source owner"
)]
enum PreparedDocx {
    Eager(litchi_docx::Package, litchi_core::Metadata),
    Source(litchi::Document),
}

impl PreparedDocx {
    fn eager(bytes: Vec<u8>) -> Result<Self, Box<dyn Error>> {
        let detected = litchi::detection_smart::detect_format_smart_with_limits(
            bytes,
            litchi_docx::ReadLimits::default(),
        )
        .ok_or("DOCX eager detector did not identify a supported package")?;
        let package = match detected {
            litchi::detection_smart::DetectedFormat::Docx(opc) => {
                litchi_docx::Package::from_opc_package(opc)?
            },
            _ => return Err("DOCX eager detector returned a non-DOCX package".into()),
        };
        package.document()?.text()?;
        let metadata = package
            .props()
            .cloned()
            .map(litchi_core::Metadata::from)
            .unwrap_or_default();
        Ok(Self::Eager(package, metadata))
    }

    fn source(source: &Path) -> Result<Self, Box<dyn Error>> {
        Ok(Self::Source(litchi::Document::open(source)?))
    }

    fn paragraph_count(&self) -> Result<usize, Box<dyn Error>> {
        match self {
            Self::Eager(package, _) => Ok(package.document()?.paragraph_count()?),
            Self::Source(document) => Ok(document.paragraph_count()?),
        }
    }

    fn list_paragraphs(&self) -> Result<(), Box<dyn Error>> {
        match self {
            Self::Eager(package, _) => {
                let paragraphs = package
                    .document()?
                    .paragraphs()?
                    .into_iter()
                    .map(litchi::document::Paragraph::Docx)
                    .collect::<Vec<_>>();
                std::hint::black_box(paragraphs);
            },
            Self::Source(document) => {
                let paragraphs = document.paragraphs()?;
                std::hint::black_box(paragraphs);
            },
        }
        Ok(())
    }

    fn text(&self) -> Result<String, Box<dyn Error>> {
        match self {
            Self::Eager(package, _) => Ok(package.document()?.text()?),
            Self::Source(document) => Ok(document.text()?),
        }
    }

    fn semantic_signature(&self) -> Result<String, Box<dyn Error>> {
        match self {
            Self::Eager(package, metadata) => docx_package_signature(package, metadata),
            Self::Source(document) => docx_document_signature(document),
        }
    }
}

fn run_docx_operation(
    operation: Operation,
    source: &Path,
    prepared: Option<&PreparedDocx>,
) -> Result<(), Box<dyn Error>> {
    match operation {
        Operation::DocxEagerOpen => {
            let document = PreparedDocx::eager(fs::read(source)?)?;
            std::hint::black_box(document);
        },
        Operation::DocxSourceOpen => {
            // This is the candidate path under measurement. The root facade
            // must adopt the filesystem source rather than replaying a byte
            // buffer through the eager compatibility path.
            let document = litchi::Document::open(source)?;
            std::hint::black_box(document);
        },
        Operation::DocxEagerParagraphCount | Operation::DocxSourceParagraphCount => {
            let document = prepared.ok_or("DOCX paragraph-count operation has no prepared root")?;
            let count = document.paragraph_count()?;
            std::hint::black_box(count);
        },
        Operation::DocxEagerListParagraphs | Operation::DocxSourceListParagraphs => {
            let document = prepared.ok_or("DOCX list-paragraphs operation has no prepared root")?;
            document.list_paragraphs()?;
        },
        Operation::DocxEagerFullText | Operation::DocxSourceFullText => {
            let document = prepared.ok_or("DOCX full-text operation has no prepared root")?;
            let text = document.text()?;
            std::hint::black_box(text);
        },
        Operation::DocxEagerOpenParagraphCountLifecycle
        | Operation::DocxSourceOpenParagraphCountLifecycle => {
            let document = if operation.is_source_docx() {
                PreparedDocx::source(source)?
            } else {
                PreparedDocx::eager(fs::read(source)?)?
            };
            let count = document.paragraph_count()?;
            if count != super::SemanticShape::Medium.docx_paragraphs() {
                return Err("DOCX lifecycle paragraph count differs from fixed corpus".into());
            }
            std::hint::black_box((document, count));
        },
        Operation::DocxEagerOpenFullTextLifecycle | Operation::DocxSourceOpenFullTextLifecycle => {
            let document = if operation.is_source_docx() {
                PreparedDocx::source(source)?
            } else {
                PreparedDocx::eager(fs::read(source)?)?
            };
            let text = document.text()?;
            std::hint::black_box((document, text));
        },
        _ => return Err("non-DOCX operation passed to run_docx_operation".into()),
    }
    Ok(())
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct XlsxRootSemanticProjection {
    worksheet_names: Vec<String>,
    worksheet_count: usize,
    full_text: String,
    metadata_sha256: String,
}

/// Builds the expected XLSX projection through the typed owner and an
/// independently opened OPC/property package. The facade never manufactures
/// this oracle from bytes or from its own metadata implementation.
fn xlsx_typed_semantic_oracle(
    corpus: &super::Corpus,
) -> Result<XlsxRootSemanticProjection, Box<dyn Error>> {
    assert_pinned_xlsx_corpus(corpus)?;
    let spec = corpus
        .xlsx
        .as_ref()
        .ok_or("XLSX filesystem corpus omitted its typed specification")?;
    let typed = litchi_xlsx::Workbook::from_bytes(corpus.archive.clone())?;
    super::verify_xlsx_cells(&typed, spec, &[])?;
    let worksheet_names = typed
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    let worksheet_count = typed.len();
    let full_text = super::xlsx_typed_full_text(&typed, spec)?;
    if worksheet_names.len() != XLSX_FILE_SOURCE_SHEET_COUNT
        || worksheet_count != XLSX_FILE_SOURCE_SHEET_COUNT
    {
        return Err("typed XLSX filesystem oracle worksheet shape differs from its pin".into());
    }

    let package = OpcPackage::from_bytes(&corpus.archive)?;
    let metadata = litchi_ooxml_common::properties::read(&package)
        .map_err(|error| error.to_string())?
        .map(litchi_core::Metadata::from)
        .unwrap_or_default();
    let metadata_sha256 = super::xlsx_root_metadata_digest(&metadata)?;
    Ok(XlsxRootSemanticProjection {
        worksheet_names,
        worksheet_count,
        full_text,
        metadata_sha256,
    })
}

fn xlsx_semantic_sha256(corpus: &super::Corpus) -> Result<String, Box<dyn Error>> {
    let oracle = xlsx_typed_semantic_oracle(corpus)?;
    xlsx_semantic_projection_sha256(&oracle)
}

fn xlsx_semantic_projection_sha256(
    projection: &XlsxRootSemanticProjection,
) -> Result<String, Box<dyn Error>> {
    Ok(super::sha256_hex(&serde_json::to_vec(&(
        &projection.worksheet_names,
        projection.worksheet_count,
        &projection.full_text,
        &projection.metadata_sha256,
    ))?))
}

fn xlsx_timed_names_count_text(
    workbook: &litchi::Workbook,
) -> Result<(Vec<String>, usize, String), Box<dyn Error>> {
    let worksheet_names = workbook
        .worksheet_names()
        .map_err(|error| error.to_string())?;
    let worksheet_count = workbook
        .worksheet_count()
        .map_err(|error| error.to_string())?;
    let full_text = workbook.text().map_err(|error| error.to_string())?;
    Ok((worksheet_names, worksheet_count, full_text))
}

struct DeferredXlsxOperation {
    workbook: litchi::Workbook,
    timed_projection: Option<(Vec<String>, usize, String)>,
}

fn xlsx_deferred_semantic_projection(
    operation: Operation,
    deferred: &DeferredXlsxOperation,
) -> Result<XlsxRootSemanticProjection, Box<dyn Error>> {
    let (worksheet_names, worksheet_count, full_text) =
        match (operation, deferred.timed_projection.as_ref()) {
            (Operation::XlsxFileOpen, None) => xlsx_timed_names_count_text(&deferred.workbook)?,
            (Operation::XlsxFileOpenLifecycle, Some((names, count, text))) => {
                // Clone only after the timer and evidence snapshots: these are
                // the exact names/count/text values produced by the timed scope.
                (names.clone(), *count, text.clone())
            },
            (Operation::XlsxFileOpen, Some(_)) => {
                return Err("XLSX open retained an unexpected lifecycle projection".into());
            },
            (Operation::XlsxFileOpenLifecycle, None) => {
                return Err("XLSX lifecycle omitted its timed projection".into());
            },
            _ => return Err("non-XLSX operation passed to deferred XLSX projection".into()),
        };
    let metadata = deferred
        .workbook
        .metadata()
        .map_err(|error| error.to_string())?;
    let metadata_sha256 = super::xlsx_root_metadata_digest(&metadata)?;
    Ok(XlsxRootSemanticProjection {
        worksheet_names,
        worksheet_count,
        full_text,
        metadata_sha256,
    })
}

fn run_xlsx_operation(
    operation: Operation,
    source: &Path,
    deferred: &mut Option<DeferredXlsxOperation>,
) -> Result<(), Box<dyn Error>> {
    match operation {
        Operation::XlsxFileOpen => {
            // The open-only selector measures only path open and root
            // construction. Its full semantic projection is consumed after
            // the timer in `run_child_arguments`.
            *deferred = Some(DeferredXlsxOperation {
                workbook: litchi::Workbook::open(source).map_err(|error| error.to_string())?,
                timed_projection: None,
            });
        },
        Operation::XlsxFileOpenLifecycle => {
            // Keep this timed boundary to exactly path open plus the declared
            // names/count/full-text projection. Semantic hashing, JSON
            // serialization, and metadata I/O are deferred until after the
            // operation-only evidence snapshots.
            let workbook = litchi::Workbook::open(source).map_err(|error| error.to_string())?;
            let projection = xlsx_timed_names_count_text(&workbook)?;
            *deferred = Some(DeferredXlsxOperation {
                workbook,
                timed_projection: Some(projection),
            });
        },
        _ => return Err("non-XLSX operation passed to run_xlsx_operation".into()),
    }
    Ok(())
}

fn run_opc_eager_save(
    source: &Path,
    destination: &Path,
    replacement: &[u8],
    details: &mut OperationDetails,
) -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
    let target_uri = PackURI::new(format!("/{}", super::entry_name(OPC_FILE_TARGET_INDEX)))?;
    let mut package = OpcPackage::from_bytes(&fs::read(source)?)?;
    details.opc_materialized_parts = Some(u64::try_from(package.part_count())?);
    package
        .get_part_mut(&target_uri)?
        .set_blob(replacement.to_vec());
    package.save(destination)?;
    Ok(None)
}

fn run_opc_source_save(
    source: &Path,
    destination: &Path,
    replacement: &[u8],
    details: &mut OperationDetails,
) -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
    let target_uri = PackURI::new(format!("/{}", super::entry_name(OPC_FILE_TARGET_INDEX)))?;
    let counter = Arc::new(CountingReadAt::new(Arc::new(FileSource::open(source)?)));
    let package = SourceBackedPackage::from_read_at(counter.clone())?;
    litchi_opc::atomic::replace_with(destination, |file| {
        package.write_part_overlay_to_stream(file, &target_uri, replacement.to_vec())
    })?;
    // The production overlay publisher consumes the source-backed package.
    // Its raw-copy contract never materializes an ordinary Part, so record
    // that post-operation fact explicitly rather than sampling diagnostics
    // before the timed publisher.
    details.opc_materialized_parts = Some(0);
    Ok(Some(counter))
}

fn filesystem_opc_replacement() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut replacement = super::payload_bytes(
        OPC_FILE_PAYLOAD,
        OPC_FILE_TARGET_INDEX,
        OPC_FILE_ENTRY_BYTES,
    );
    let first = replacement
        .first_mut()
        .ok_or("filesystem OPC replacement target is empty")?;
    *first ^= 0xff;
    Ok(replacement)
}

fn run_cfb_overlay_save(
    source: &Path,
    destination: &Path,
    details: &mut OperationDetails,
) -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
    let open_started = Instant::now();
    let counter = Arc::new(CountingReadAt::new(Arc::new(FileSource::open(source)?)));
    let shared = SharedOleFile::open(counter.clone())?;
    let open_elapsed_ns = u64::try_from(open_started.elapsed().as_nanos())?;
    let after_open = counter.totals();

    let plan_started = Instant::now();
    let overlay = SameLengthStreamOverlay::new(
        vec![super::OLE_COMMON_TARGET.to_owned()],
        Arc::from(FILESYSTEM_OLE_COMMON_REPLACEMENT.to_vec()),
    );
    let plan = shared.plan_same_length_stream_overlays(vec![overlay], OverlayLimits::default())?;
    let plan_elapsed_ns = u64::try_from(plan_started.elapsed().as_nanos())?;
    let after_plan = counter.totals();

    let publication_started = Instant::now();
    let report = plan.save(destination)?;
    let publication_elapsed_ns = u64::try_from(publication_started.elapsed().as_nanos())?;
    let after_publication = counter.totals();
    details.cfb_changed_spans = Some(u64::try_from(report.changed_spans())?);
    details.cfb_published_bytes = Some(report.bytes());
    details.cfb_phases = Some(CfbPhaseEvidence {
        open: cfb_phase_sample(ReadTotals::default(), after_open, open_elapsed_ns)?,
        plan: cfb_phase_sample(after_open, after_plan, plan_elapsed_ns)?,
        atomic_publication: cfb_phase_sample(
            after_plan,
            after_publication,
            publication_elapsed_ns,
        )?,
    });
    Ok(Some(counter))
}

fn run_cfb_owned_overlay_save(
    source: &Path,
    destination: &Path,
    details: &mut OperationDetails,
) -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
    // The filesystem ingress is intentionally outside all three phase timers.
    // The resulting immutable slice is the source owned by `SharedOleFile`;
    // no logical ReadAt adapter is involved in the timed operation.
    let source_bytes = fs::read(source)?;
    let source_ingress_bytes = u64::try_from(source_bytes.len())?;
    let source_sha256 = super::sha256_hex(&source_bytes);
    let owned_source: Arc<[u8]> = Arc::from(source_bytes.into_boxed_slice());
    let version = SourceVersion::new(CFB_FILE_OWNED_SOURCE_VERSION_ID, 0);

    let open_started = Instant::now();
    let shared = SharedOleFile::open_owned(Arc::clone(&owned_source), version)?;
    let open_elapsed_ns = u64::try_from(open_started.elapsed().as_nanos())?;

    let plan_started = Instant::now();
    let overlay = SameLengthStreamOverlay::new(
        vec![super::OLE_COMMON_TARGET.to_owned()],
        Arc::from(FILESYSTEM_OLE_COMMON_REPLACEMENT.to_vec()),
    );
    let plan = shared.plan_same_length_stream_overlays(vec![overlay], OverlayLimits::default())?;
    let plan_elapsed_ns = u64::try_from(plan_started.elapsed().as_nanos())?;

    let publication_started = Instant::now();
    let report = plan.save(destination)?;
    let publication_elapsed_ns = u64::try_from(publication_started.elapsed().as_nanos())?;
    details.cfb_changed_spans = Some(u64::try_from(report.changed_spans())?);
    details.cfb_published_bytes = Some(report.bytes());
    details.cfb_owned = Some(CfbOwnedEvidence {
        implementation: "SharedOleFile::open_owned".to_owned(),
        ingress: "filesystem_read_all_before_cfb_phase_timers".to_owned(),
        ownership: "Arc<[u8]>".to_owned(),
        logical_read_counter_scope: "not_applicable_immutable_owned_slice".to_owned(),
        source_ingress_bytes,
        source_sha256,
        source_version_id: version.id(),
        source_version_revision: version.revision(),
        phases: CfbOwnedPhaseEvidence {
            open: CfbOwnedPhaseSample {
                elapsed_ns: open_elapsed_ns,
            },
            plan: CfbOwnedPhaseSample {
                elapsed_ns: plan_elapsed_ns,
            },
            atomic_publication: CfbOwnedPhaseSample {
                elapsed_ns: publication_elapsed_ns,
            },
        },
    });
    Ok(None)
}

fn cfb_phase_sample(
    before: ReadTotals,
    after: ReadTotals,
    elapsed_ns: u64,
) -> Result<CfbPhaseSample, Box<dyn Error>> {
    let delta = |name: &str, before: u64, after: u64| {
        after.checked_sub(before).ok_or_else(|| {
            io::Error::other(format!(
                "CFB phase counter {name} moved backwards from {before} to {after}"
            ))
        })
    };
    Ok(CfbPhaseSample {
        elapsed_ns,
        logical_read_calls: delta("calls", before.calls, after.calls)?,
        logical_read_requested_bytes: delta(
            "requested bytes",
            before.requested_bytes,
            after.requested_bytes,
        )?,
        logical_read_returned_bytes: delta(
            "returned bytes",
            before.returned_bytes,
            after.returned_bytes,
        )?,
    })
}

fn verify_child_output(
    operation: Operation,
    source: &Path,
    destination: &Path,
    corpus: &super::Corpus,
    allow_page_aligned_source: bool,
) -> Result<(), Box<dyn Error>> {
    match operation {
        Operation::OpcEagerOpen => {
            let package = OpcPackage::from_bytes(&fs::read(source)?)?;
            verify_opc_package(&package, corpus)
        },
        Operation::OpcSourceOpen => {
            let file_source = FileSource::open(source)?;
            let package = SourceBackedPackage::from_read_at(Arc::new(file_source))?;
            if package.iter_parts().count() != corpus.manifest.entry_count {
                return Err("source-backed OPC filesystem reopen part count differs".into());
            }
            let main = package.main_document_part()?;
            let payload = main.data()?.into_arc()?;
            if main.partname().membername() != corpus.target_name
                || payload.as_slice() != corpus.target_payload
            {
                return Err("source-backed OPC filesystem reopen differs from corpus".into());
            }
            Ok(())
        },
        Operation::OpcEagerSave | Operation::OpcSourceSave => {
            let output = fs::read(destination)?;
            let replacement = filesystem_opc_replacement()?;
            super::verify_opc_overlay_output(corpus, &output, &replacement)
        },
        Operation::CfbOverlaySave | Operation::CfbOwnedOverlaySave => {
            let mut ole = OleFile::open(File::open(destination)?)?;
            if ole.list_streams().len() != corpus.manifest.entry_count
                || ole.open_stream(&[super::OLE_COMMON_TARGET])?
                    != FILESYSTEM_OLE_COMMON_REPLACEMENT
            {
                return Err("filesystem CFB overlay semantic reopen differs".into());
            }
            for index in 0..corpus.manifest.entry_count.saturating_sub(1) {
                let name = super::cfb_entry_name(index);
                if ole.open_stream(&[name.as_str()])?
                    != super::payload_bytes(
                        super::PayloadKind::Incompressible,
                        index,
                        corpus.manifest.entry_bytes,
                    )
                {
                    return Err("filesystem CFB overlay changed an untouched stream".into());
                }
            }
            Ok(())
        },
        Operation::PptxEagerOpen
        | Operation::PptxSourceOpen
        | Operation::PptxEagerListSlides
        | Operation::PptxSourceListSlides
        | Operation::PptxEagerSlideCount
        | Operation::PptxSourceSlideCount
        | Operation::PptxEagerSelectedSlide
        | Operation::PptxSourceSelectedSlide
        | Operation::PptxEagerOpenSlideCountLifecycle
        | Operation::PptxSourceOpenSlideCountLifecycle
        | Operation::PptxEagerOpenSelectedSlideLifecycle
        | Operation::PptxSourceOpenSelectedSlideLifecycle => {
            verify_pptx_operation(source, corpus, allow_page_aligned_source)
        },
        Operation::DocxEagerOpen
        | Operation::DocxSourceOpen
        | Operation::DocxEagerParagraphCount
        | Operation::DocxSourceParagraphCount
        | Operation::DocxEagerListParagraphs
        | Operation::DocxSourceListParagraphs
        | Operation::DocxEagerFullText
        | Operation::DocxSourceFullText
        | Operation::DocxEagerOpenParagraphCountLifecycle
        | Operation::DocxSourceOpenParagraphCountLifecycle
        | Operation::DocxEagerOpenFullTextLifecycle
        | Operation::DocxSourceOpenFullTextLifecycle => {
            verify_docx_operation(source, corpus, allow_page_aligned_source)
        },
        Operation::XlsxFileOpen | Operation::XlsxFileOpenLifecycle => {
            // XLSX correctness is validated once against the exact timed
            // workbook before this output-only dispatch reaches this arm.
            Ok(())
        },
    }
}

fn verify_pptx_operation(
    source: &Path,
    corpus: &super::Corpus,
    allow_page_aligned_source: bool,
) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != PPTX_FILE_CORPUS_GENERATOR {
        return Err("PPTX filesystem source has the wrong corpus generator".into());
    }
    let bytes = fs::read(source)?;
    if !allow_page_aligned_source {
        if super::sha256_hex(&bytes) != corpus.manifest.archive_sha256 {
            return Err("PPTX filesystem source hash differs from corpus manifest".into());
        }
        if bytes.len() != corpus.manifest.archive_bytes {
            return Err("PPTX filesystem source length differs from corpus manifest".into());
        }
    }
    let eager = litchi::Presentation::from_bytes(bytes)?;
    let source_backed = litchi::Presentation::open(source)?;
    let eager_signature = pptx_presentation_signature(&eager)?;
    let source_signature = pptx_presentation_signature(&source_backed)?;
    if eager_signature != source_signature {
        return Err("PPTX eager/source ordinary-root semantic signatures differ".into());
    }
    if eager.slide_count()? != super::PPTX_SOURCE_SLIDE_COUNT {
        return Err("PPTX filesystem corpus slide count differs from specification".into());
    }
    Ok(())
}

fn verify_docx_operation(
    source: &Path,
    corpus: &super::Corpus,
    allow_page_aligned_source: bool,
) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != DOCX_FILE_CORPUS_GENERATOR {
        return Err("DOCX filesystem source has the wrong corpus generator".into());
    }
    let bytes = fs::read(source)?;
    let source_sha256 = super::sha256_hex(&bytes);
    if !allow_page_aligned_source {
        if bytes.len() != corpus.manifest.archive_bytes {
            return Err("DOCX filesystem source length differs from corpus manifest".into());
        }
        if source_sha256 != corpus.manifest.archive_sha256 {
            return Err("DOCX filesystem source hash differs from corpus manifest".into());
        }
    }
    let eager = PreparedDocx::eager(bytes.clone())?;
    let source_backed = PreparedDocx::source(source)?;
    if eager.semantic_signature()? != source_backed.semantic_signature()? {
        return Err("DOCX eager/source ordinary-root semantic signatures differ".into());
    }
    if eager.paragraph_count()? != super::SemanticShape::Medium.docx_paragraphs() {
        return Err("DOCX filesystem corpus paragraph count differs from specification".into());
    }
    let aligned_signature = docx_archive_signature(&bytes)?;
    let corpus_signature = docx_archive_signature(&corpus.archive)?;
    if aligned_signature.semantic_sha256 != corpus_signature.semantic_sha256 {
        return Err(format!(
            "DOCX filesystem archive topology or payload hashes differ (aligned bytes {}, corpus bytes {})",
            aligned_signature.physical_bytes, corpus_signature.physical_bytes
        )
        .into());
    }
    assert_source_sha256(source, &source_sha256)?;
    Ok(())
}

fn verify_xlsx_operation(
    operation: Operation,
    source: &Path,
    corpus: &super::Corpus,
    allow_page_aligned_source: bool,
    deferred: &DeferredXlsxOperation,
) -> Result<(String, String), Box<dyn Error>> {
    assert_pinned_xlsx_corpus(corpus)?;
    let initial_bytes = fs::read(source)?;
    if !allow_page_aligned_source {
        if initial_bytes.len() != XLSX_FILE_SOURCE_ARCHIVE_BYTES {
            return Err("XLSX filesystem source length differs from corpus manifest".into());
        }
    }
    // The typed owner and independent OPC package are both opened only after
    // the operation evidence has been sampled. The exact facade workbook from
    // the timed operation is checked against that oracle rather than a second
    // path open or a facade-from-bytes oracle.
    let expected = xlsx_typed_semantic_oracle(corpus)?;
    let observed = xlsx_deferred_semantic_projection(operation, deferred)?;
    if observed != expected {
        return Err("XLSX filesystem semantic projection differs from deterministic corpus".into());
    }
    let semantic_sha256 = xlsx_semantic_projection_sha256(&observed)?;

    // Read and hash the source only after semantic verification. This is the
    // final exact source identity, including the aligned cold-verifier copy;
    // the parent additionally checks it against the prepared aligned hash.
    let final_bytes = fs::read(source)?;
    let source_sha256 = super::sha256_hex(&final_bytes);
    if !allow_page_aligned_source {
        if final_bytes.len() != XLSX_FILE_SOURCE_ARCHIVE_BYTES
            || source_sha256 != XLSX_FILE_SOURCE_SHA256
        {
            return Err("XLSX filesystem source hash or length differs from its fixed pin".into());
        }
    }
    Ok((source_sha256, semantic_sha256))
}

fn docx_document_signature(document: &litchi::Document) -> Result<String, Box<dyn Error>> {
    let paragraphs = document
        .paragraphs()?
        .into_iter()
        .map(|paragraph| paragraph.text())
        .collect::<Result<Vec<_>, _>>()?;
    let tables = document
        .tables()?
        .into_iter()
        .map(|table| docx_table_projection(&table))
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    let elements = document
        .elements()?
        .into_iter()
        .map(|element| {
            if let Some(paragraph) = element.as_paragraph() {
                return Ok(("paragraph", paragraph.text()?));
            }
            if let Some(table) = element.as_table() {
                return Ok((
                    "table",
                    serde_json::to_string(&docx_table_projection(table)?)?,
                ));
            }
            Err("DOCX element has no supported projection".into())
        })
        .collect::<Result<Vec<(&str, String)>, Box<dyn Error>>>()?;
    let metadata = document.metadata()?;
    Ok(super::sha256_hex(&serde_json::to_vec(&(
        document.paragraph_count()?,
        document.text()?,
        paragraphs,
        tables,
        elements,
        metadata,
    ))?))
}

fn docx_package_signature(
    package: &litchi_docx::Package,
    metadata: &litchi_core::Metadata,
) -> Result<String, Box<dyn Error>> {
    let document = package.document()?;
    let paragraphs = document
        .paragraphs()?
        .into_iter()
        .map(|paragraph| paragraph.text())
        .collect::<Result<Vec<_>, _>>()?;
    let tables = document
        .tables()?
        .into_iter()
        .map(|table| docx_package_table_projection(&table))
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    let elements = document.elements()?.into_iter().try_fold(
        Vec::new(),
        |mut elements, element| -> Result<Vec<(&str, String)>, Box<dyn Error>> {
            match element {
                litchi_docx::Element::Paragraph(paragraph) => {
                    elements.push(("paragraph", paragraph.text()?));
                },
                litchi_docx::Element::Table(table) => {
                    let projection = docx_package_table_projection(&table)?;
                    elements.push(("table", serde_json::to_string(&projection)?));
                },
                litchi_docx::Element::Unknown(block)
                    if docx_unknown_is_section_properties(&block) => {},
                litchi_docx::Element::Unknown(_) => {
                    return Err("DOCX element has no supported projection".into());
                },
            }
            Ok(elements)
        },
    )?;
    Ok(super::sha256_hex(&serde_json::to_vec(&(
        document.paragraph_count()?,
        document.text()?,
        paragraphs,
        tables,
        elements,
        metadata,
    ))?))
}

fn docx_table_projection(
    table: &litchi::document::Table,
) -> Result<(usize, Vec<(usize, Vec<String>)>), Box<dyn Error>> {
    let row_count = table.row_count()?;
    let rows = table.rows()?;
    if rows.len() != row_count {
        return Err("DOCX table row count disagrees with its row projection".into());
    }
    let rows = rows
        .into_iter()
        .map(|row| {
            let cell_count = row.cell_count()?;
            let cells = row
                .cells()?
                .into_iter()
                .map(|cell| cell.text())
                .collect::<Result<Vec<_>, _>>()?;
            if cells.len() != cell_count {
                return Err("DOCX table cell count disagrees with its cell projection".into());
            }
            Ok((cell_count, cells))
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    Ok((row_count, rows))
}

fn docx_package_table_projection(
    table: &litchi_docx::Table,
) -> Result<(usize, Vec<(usize, Vec<String>)>), Box<dyn Error>> {
    let row_count = table.row_count()?;
    let rows = table.rows()?;
    if rows.len() != row_count {
        return Err("DOCX package table row count disagrees with its row projection".into());
    }
    let rows = rows
        .into_iter()
        .map(|row| {
            let cell_count = row.cell_count()?;
            let cells = row
                .cells()?
                .into_iter()
                .map(|cell| cell.text())
                .collect::<Result<Vec<_>, _>>()?;
            if cells.len() != cell_count {
                return Err(
                    "DOCX package table cell count disagrees with its cell projection".into(),
                );
            }
            Ok((cell_count, cells))
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    Ok((row_count, rows))
}

fn docx_unknown_is_section_properties(block: &litchi_docx::OpaqueBlock) -> bool {
    let bytes = block.xml_bytes();
    let Some(open) = bytes.iter().position(|byte| *byte == b'<') else {
        return false;
    };
    let name_start = open.saturating_add(1);
    let Some(name_end) = bytes[name_start..]
        .iter()
        .position(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
        .map(|offset| name_start.saturating_add(offset))
    else {
        return false;
    };
    let name = &bytes[name_start..name_end];
    name == b"sectPr" || name.strip_prefix(b"w:") == Some(b"sectPr")
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct DocxArchiveSignature {
    physical_bytes: usize,
    semantic_sha256: String,
}

fn docx_archive_signature(bytes: &[u8]) -> Result<DocxArchiveSignature, Box<dyn Error>> {
    let package = OpcPackage::from_bytes(bytes)?;
    let mut parts = package
        .iter_parts()
        .map(|part| {
            let mut relationships = part
                .rels()
                .iter()
                .map(|relationship| {
                    (
                        relationship.r_id().to_owned(),
                        relationship.reltype().to_owned(),
                        relationship.target_ref().to_owned(),
                        relationship.is_external(),
                    )
                })
                .collect::<Vec<_>>();
            relationships.sort();
            Ok((
                part.partname().to_string(),
                part.content_type().to_owned(),
                relationships,
                super::sha256_hex(part.blob()),
            ))
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    let mut package_relationships = package
        .rels()
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            )
        })
        .collect::<Vec<_>>();
    package_relationships.sort();
    Ok(DocxArchiveSignature {
        physical_bytes: bytes.len(),
        semantic_sha256: super::sha256_hex(&serde_json::to_vec(&(
            package.part_count(),
            package_relationships,
            parts,
        ))?),
    })
}

fn pptx_presentation_signature(
    presentation: &litchi::Presentation,
) -> Result<(usize, Option<i64>, Option<i64>, String, String, String), Box<dyn Error>> {
    let metadata = serde_json::to_string(&presentation.metadata()?)?;
    let mut slides_hasher = Sha256::new();
    for (position, slide) in presentation.slides()?.into_iter().enumerate() {
        let text = slide.text()?;
        let name = slide.name()?.unwrap_or_default();
        slides_hasher.update(position.to_le_bytes());
        slides_hasher.update((text.len() as u64).to_le_bytes());
        slides_hasher.update(text.as_bytes());
        slides_hasher.update((name.len() as u64).to_le_bytes());
        slides_hasher.update(name.as_bytes());
    }
    let selected = presentation
        .slide(PPTX_FILE_SELECTED_POSITION)?
        .ok_or("PPTX signature selected slide is missing")?;
    let selected_text = selected.text()?;
    let selected_name = selected.name()?.unwrap_or_default();
    let mut selected_hasher = Sha256::new();
    selected_hasher.update(selected_text.as_bytes());
    selected_hasher.update([0]);
    selected_hasher.update(selected_name.as_bytes());
    Ok((
        presentation.slide_count()?,
        presentation.slide_width()?,
        presentation.slide_height()?,
        metadata,
        super::sha256_hex(slides_hasher.finalize().as_slice()),
        super::sha256_hex(selected_hasher.finalize().as_slice()),
    ))
}

fn verify_opc_package(package: &OpcPackage, corpus: &super::Corpus) -> Result<(), Box<dyn Error>> {
    if package.part_count() != corpus.manifest.entry_count {
        return Err("eager OPC filesystem reopen part count differs".into());
    }
    let main = package.main_document_part()?;
    if main.partname().membername() != corpus.target_name || main.blob() != corpus.target_payload {
        return Err("eager OPC filesystem reopen differs from corpus".into());
    }
    Ok(())
}

fn seed_destination(path: &Path, bytes: &[u8]) -> io::Result<()> {
    write_synced(path, bytes)
}

fn write_synced(path: &Path, bytes: &[u8]) -> io::Result<()> {
    let mut file = OpenOptions::new()
        .create(true)
        .truncate(true)
        .write(true)
        .open(path)?;
    file.write_all(bytes)?;
    file.flush()?;
    file.sync_all()
}

fn filesystem_root(requested_root: Option<&Path>) -> io::Result<PathBuf> {
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    let parent = requested_root.map_or_else(env::temp_dir, Path::to_path_buf);
    fs::create_dir_all(&parent)?;
    let root = parent.join(format!(
        "litchi-perf-filesystem-{}-{timestamp}",
        std::process::id()
    ));
    fs::create_dir(&root)?;
    Ok(root)
}

#[cfg(test)]
mod tests {
    use std::{
        env,
        ffi::OsString,
        fs,
        panic::{AssertUnwindSafe, catch_unwind},
        path::Path,
        process::Command,
        sync::{Arc, Barrier, atomic::Ordering},
        thread,
    };

    use litchi_core::{OwnedSource, ReadAt, SourceVersion};

    use super::{
        CacheSelection, ChildMode, ColdAdvice, CountingReadAt, Operation, ReadPattern,
        ReadPatternState, ReadSizeBuckets, checked_atomic_add, checked_atomic_sub,
        xlsx_semantic_sha256,
    };

    struct OverReturningSource;

    impl ReadAt for OverReturningSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(16)
        }

        fn read_at(&self, _offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            Ok(output.len() + 1)
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(1, 0))
        }
    }

    struct FailingSource;

    impl ReadAt for FailingSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(16)
        }

        fn read_at(&self, _offset: u64, _output: &mut [u8]) -> std::io::Result<usize> {
            Err(std::io::Error::other("synthetic source failure"))
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(2, 0))
        }
    }

    struct ExactReturningSource;

    impl ReadAt for ExactReturningSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(u64::MAX)
        }

        fn read_at(&self, _offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            output.fill(0);
            Ok(output.len())
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(3, 0))
        }
    }

    struct ConcurrentExactSource {
        barrier: Arc<Barrier>,
    }

    impl ReadAt for ConcurrentExactSource {
        fn len(&self) -> std::io::Result<u64> {
            Ok(64)
        }

        fn read_at(&self, _offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.barrier.wait();
            output.fill(0);
            Ok(output.len())
        }

        fn version(&self) -> std::io::Result<SourceVersion> {
            Ok(SourceVersion::new(4, 0))
        }
    }

    #[test]
    fn filesystem_case_names_are_explicit_and_parseable() {
        for name in [
            "opc_file_eager_open",
            "opc_file_source_open",
            "opc_file_eager_one_part_atomic_save",
            "opc_file_source_one_part_atomic_save",
            "cfb_file_same_length_overlay_atomic_save",
            "cfb_file_owned_same_length_overlay_atomic_save",
            "pptx_file_eager_open",
            "pptx_file_source_open",
            "pptx_file_eager_list_slides",
            "pptx_file_source_list_slides",
            "pptx_file_eager_slide_count",
            "pptx_file_source_slide_count",
            "pptx_file_eager_selected_slide",
            "pptx_file_source_selected_slide",
            "pptx_file_eager_open_slide_count_lifecycle",
            "pptx_file_source_open_slide_count_lifecycle",
            "pptx_file_eager_open_selected_slide_lifecycle",
            "pptx_file_source_open_selected_slide_lifecycle",
            "docx_file_eager_open",
            "docx_file_source_open",
            "docx_file_eager_paragraph_count",
            "docx_file_source_paragraph_count",
            "docx_file_eager_list_paragraphs",
            "docx_file_source_list_paragraphs",
            "docx_file_eager_full_text",
            "docx_file_source_full_text",
            "docx_file_eager_open_paragraph_count_lifecycle",
            "docx_file_source_open_paragraph_count_lifecycle",
            "docx_file_eager_open_full_text_lifecycle",
            "docx_file_source_open_full_text_lifecycle",
            "xlsx_file_open",
            "xlsx_file_open_lifecycle",
        ] {
            assert!(Operation::parse(name).is_some(), "{name}");
        }
        assert!(Operation::parse("ole2_same_length_overlay_atomic_save").is_none());
    }

    #[test]
    fn owned_cfb_operation_marks_immutable_ingress_without_read_counters() {
        let operation = Operation::parse("cfb_file_owned_same_length_overlay_atomic_save")
            .expect("owned CFB selector parses");
        assert!(operation.is_cfb());
        assert!(operation.is_cfb_owned());
        assert!(operation.is_save());
        assert_eq!(
            operation.case().name(),
            "cfb_file_owned_same_length_overlay_atomic_save"
        );
    }

    #[test]
    fn pptx_root_operation_scopes_are_explicit() {
        assert!(Operation::PptxEagerOpen.is_pptx());
        assert!(Operation::PptxSourceSelectedSlide.is_source_pptx());
        assert!(Operation::PptxSourceSelectedSlide.is_pptx_query());
        assert!(!Operation::PptxSourceOpen.is_pptx_query());
        assert!(!Operation::PptxSourceOpenSlideCountLifecycle.is_pptx_query());
        assert!(Operation::PptxSourceOpenSlideCountLifecycle.is_source_pptx());
        assert_eq!(
            Operation::PptxSourceListSlides.pptx_query_name(),
            Some("list_slides")
        );
        assert_eq!(
            Operation::PptxEagerSlideCount.pptx_query_name(),
            Some("slide_count")
        );
    }

    #[test]
    fn docx_root_operation_scopes_are_explicit() {
        assert!(Operation::DocxEagerOpen.is_docx());
        assert!(Operation::DocxSourceFullText.is_source_docx());
        assert!(Operation::DocxSourceFullText.is_docx_query());
        assert!(!Operation::DocxSourceOpen.is_docx_query());
        assert!(Operation::DocxSourceOpenFullTextLifecycle.is_docx_lifecycle());
        assert!(!Operation::DocxSourceOpenFullTextLifecycle.is_docx_query());
        assert!(Operation::DocxSourceOpenFullTextLifecycle.is_source_docx());
        assert_eq!(
            Operation::DocxSourceParagraphCount.docx_query_name(),
            Some("paragraph_count")
        );
        assert_eq!(
            Operation::DocxSourceListParagraphs.docx_query_name(),
            Some("list_paragraphs")
        );
        assert_eq!(
            Operation::DocxSourceFullText.docx_query_name(),
            Some("full_text")
        );
    }

    #[test]
    fn pptx_replay_overlap_accounting_is_exact_and_disjoint() {
        let request = 10..30;
        assert_eq!(super::overlap_len(&request, &(0..10)), 0);
        assert_eq!(super::overlap_len(&request, &(20..40)), 10);
        assert_eq!(super::overlap_with_ranges(&request, &[0..12, 24..28]), 6);
        let disjoint = 30..40;
        assert_eq!(
            super::overlap_with_ranges(&request, std::slice::from_ref(&disjoint)),
            0
        );
        assert_eq!(
            super::pptx_slide_part_position("ppt/slides/slide101.xml"),
            Some(100)
        );
        assert_eq!(
            super::pptx_slide_part_position("ppt/slides/slide1.xml"),
            Some(0)
        );
        assert_eq!(
            super::pptx_slide_part_position("ppt/slides/slide0.xml"),
            None
        );
        assert_eq!(
            super::pptx_slide_part_position("ppt/slideLayouts/slideLayout1.xml"),
            None
        );
        assert!(super::range_fully_covered(&(10..20), &[0..12, 12..20]));
        assert!(!super::range_fully_covered(&(10..20), &[0..12, 13..20]));
        assert_eq!(
            super::fully_covered_range_count(&[0..10, 20..30], &[0..10, 20..25]),
            1
        );
    }

    #[test]
    fn docx_replay_coverage_tracks_each_payload_class() {
        let ranges = super::DocxReplayRanges {
            main: 100..120,
            media: std::iter::once(200..220).collect(),
            unselected: std::iter::once(300..330).collect(),
            core: std::iter::once(400..410).collect(),
        };
        let source = super::DocxReplaySource::new(Arc::new(vec![0; 512]), ranges.clone());
        let before = source.snapshot().unwrap();
        let mut main = [0_u8; 20];
        let mut media = [0_u8; 20];
        source.read_at(100, &mut main).unwrap();
        source.read_at(200, &mut media).unwrap();
        let after = source.snapshot().unwrap();
        let phase = super::docx_replay_phase(&before, &after, &ranges);
        assert_eq!(phase.counters.read_calls, 2);
        assert_eq!(phase.counters.read_bytes, 40);
        assert_eq!(phase.counters.main_payload_overlap_bytes, 20);
        assert_eq!(phase.counters.media_payload_overlap_bytes, 20);
        assert_eq!(phase.counters.unselected_payload_overlap_bytes, 0);
        assert_eq!(phase.counters.core_payload_overlap_bytes, 0);
        assert_eq!(phase.main_payload_covered_bytes, 20);
        assert!(phase.main_payload_fully_covered);
        assert_eq!(phase.media_payload_covered_bytes, 20);
        assert_eq!(phase.unselected_payload_covered_bytes, 0);
        assert_eq!(phase.core_payload_covered_bytes, 0);
        assert_eq!(phase.return_sizes, vec![20, 20]);
    }

    #[test]
    fn read_size_buckets_are_fixed_and_boundary_stable() {
        let mut buckets = ReadSizeBuckets::default();
        for size in [0, 1, 512, 513, 4096, 4097, 16384, 16385, 65536, 65537] {
            buckets.observe(size).unwrap();
        }
        assert_eq!(buckets.bytes_0, 1);
        assert_eq!(buckets.bytes_1_to_512, 2);
        assert_eq!(buckets.bytes_513_to_4096, 2);
        assert_eq!(buckets.bytes_4097_to_16384, 2);
        assert_eq!(buckets.bytes_16385_to_65536, 2);
        assert_eq!(buckets.bytes_over_65536, 1);
    }

    #[test]
    fn counting_source_records_exact_sizes_and_sequential_pattern() {
        let source = CountingReadAt::new(Arc::new(OwnedSource::new(vec![0; 64])));
        let mut first = [0_u8; 8];
        let mut second = [0_u8; 4];
        assert_eq!(source.read_at(12, &mut first).unwrap(), 8);
        assert_eq!(source.read_at(20, &mut second).unwrap(), 4);

        let metrics = source.snapshot().unwrap();
        assert_eq!(metrics.calls, 2);
        assert_eq!(metrics.requested_bytes, 12);
        assert_eq!(metrics.returned_bytes, 12);
        assert_eq!(metrics.largest_requested_bytes, 8);
        assert_eq!(metrics.largest_returned_bytes, 8);
        assert_eq!(metrics.pattern, Some(ReadPattern::Sequential));
        assert_eq!(metrics.request_sizes, vec![4, 8]);
    }

    #[test]
    fn counting_source_marks_noncontiguous_ranges_random_and_short_reads_unknown() {
        let source = CountingReadAt::new(Arc::new(OwnedSource::new(vec![0; 8])));
        let mut first = [0_u8; 2];
        let mut second = [0_u8; 2];
        assert_eq!(source.read_at(0, &mut first).unwrap(), 2);
        assert_eq!(source.read_at(5, &mut second).unwrap(), 2);
        assert_eq!(
            source.snapshot().unwrap().pattern,
            Some(ReadPattern::Random)
        );

        let short = CountingReadAt::new(Arc::new(OwnedSource::new(vec![0; 1])));
        let mut output = [0_u8; 2];
        assert_eq!(short.read_at(0, &mut output).unwrap(), 1);
        assert_eq!(short.read_at(1, &mut output).unwrap(), 0);
        assert_eq!(
            short.snapshot().unwrap().pattern,
            Some(ReadPattern::Unknown)
        );
    }

    #[test]
    fn counting_source_does_not_call_one_range_sequential() {
        let source = CountingReadAt::new(Arc::new(OwnedSource::new(vec![0; 8])));
        let mut output = [0_u8; 2];
        assert_eq!(source.read_at(3, &mut output).unwrap(), 2);
        assert_eq!(
            source.snapshot().unwrap().pattern,
            Some(ReadPattern::Unknown)
        );
    }

    #[test]
    fn counting_source_rejects_over_return_before_recording_returned_bytes() {
        let source = CountingReadAt::new(Arc::new(OverReturningSource));
        let mut output = [0_u8; 4];
        let error = source.read_at(0, &mut output).unwrap_err();
        assert!(error.to_string().contains("more bytes than requested"));
        assert_eq!(source.calls.load(Ordering::SeqCst), 1);
        assert_eq!(source.requested_bytes.load(Ordering::SeqCst), 4);
        assert_eq!(source.returned_bytes.load(Ordering::SeqCst), 0);
        assert_eq!(source.largest_returned_bytes.load(Ordering::SeqCst), 0);
        assert!(source.snapshot().is_err());
    }

    #[test]
    fn counting_source_records_failed_attempt_as_requested_only() {
        let source = CountingReadAt::new(Arc::new(FailingSource));
        let mut output = [0_u8; 4];
        assert!(source.read_at(0, &mut output).is_err());
        assert_eq!(source.calls.load(Ordering::SeqCst), 1);
        assert_eq!(source.requested_bytes.load(Ordering::SeqCst), 4);
        assert_eq!(source.returned_bytes.load(Ordering::SeqCst), 0);
        assert_eq!(
            source.snapshot().unwrap().pattern,
            Some(ReadPattern::Unknown)
        );
    }

    #[test]
    fn counting_source_marks_counter_overflow_unavailable() {
        let source = CountingReadAt::new(Arc::new(OwnedSource::new(vec![0; 8])));
        source.calls.store(u64::MAX, Ordering::SeqCst);
        let mut output = [0_u8; 1];
        let error = source.read_at(0, &mut output).unwrap_err();
        assert!(error.to_string().contains("read calls overflow"));
        assert!(source.snapshot().is_err());

        let counter = std::sync::atomic::AtomicU64::new(u64::MAX);
        assert!(checked_atomic_add(&counter, 1, "test counter").is_err());
        assert_eq!(counter.load(Ordering::SeqCst), u64::MAX);

        let in_flight = std::sync::atomic::AtomicU64::new(0);
        assert!(checked_atomic_sub(&in_flight, 1, "test in-flight").is_err());
        assert_eq!(in_flight.load(Ordering::SeqCst), 0);

        let mut buckets = ReadSizeBuckets {
            bytes_0: u64::MAX,
            ..ReadSizeBuckets::default()
        };
        assert!(buckets.observe(0).is_err());
        assert_eq!(buckets.bytes_0, u64::MAX);
    }

    #[test]
    fn counting_source_rejects_invalid_range_arithmetic() {
        let mut state = ReadPatternState::default();
        assert!(state.observe(u64::MAX, 1, 1).is_err());

        let source = CountingReadAt::new(Arc::new(ExactReturningSource));
        let mut output = [0_u8; 1];
        assert!(source.read_at(u64::MAX, &mut output).is_err());
        assert_eq!(source.returned_bytes.load(Ordering::SeqCst), 1);
        assert!(source.snapshot().is_err());
    }

    #[test]
    fn counting_source_marks_concurrent_observations_unknown() {
        let source = Arc::new(CountingReadAt::new(Arc::new(ConcurrentExactSource {
            barrier: Arc::new(Barrier::new(2)),
        })));
        let first_source = Arc::clone(&source);
        let first = thread::spawn(move || {
            let mut output = [0_u8; 2];
            first_source.read_at(0, &mut output)
        });
        let second_source = Arc::clone(&source);
        let second = thread::spawn(move || {
            let mut output = [0_u8; 2];
            second_source.read_at(2, &mut output)
        });
        assert_eq!(first.join().unwrap().unwrap(), 2);
        assert_eq!(second.join().unwrap().unwrap(), 2);
        assert_eq!(
            source.snapshot().unwrap().pattern,
            Some(ReadPattern::Unknown)
        );
        assert!(source.max_concurrent.load(Ordering::SeqCst) >= 2);
    }

    #[test]
    fn counting_source_poisoned_metrics_fail_closed() {
        let source = CountingReadAt::new(Arc::new(OwnedSource::new(vec![0; 8])));
        let poisoned = catch_unwind(AssertUnwindSafe(|| {
            let _guard = source.request_sizes.lock().unwrap();
            panic!("poison request-size metrics");
        }));
        assert!(poisoned.is_err());
        assert!(source.snapshot().is_err());
    }

    #[test]
    fn cold_state_labels_are_distinct_from_warm_and_unsupported() {
        assert_ne!(ColdAdvice::NotRequested as u8, ColdAdvice::Requested as u8);
        assert_eq!(ChildMode::parse("cold"), Some(ChildMode::Cold));
        assert_eq!(
            ChildMode::parse("verified-prime"),
            Some(ChildMode::VerifiedPrime)
        );
        assert_eq!(
            ChildMode::parse("cold-verified"),
            Some(ChildMode::ColdVerified)
        );
        assert_eq!(ChildMode::parse("warm"), Some(ChildMode::Warm));
        assert_eq!(ChildMode::parse("prime"), Some(ChildMode::Prime));
    }

    #[test]
    fn cache_selection_is_explicit_and_additive() {
        let default = CacheSelection::default();
        assert!(default.warm());
        assert!(default.cold_requested());
        assert!(!default.cold_verified());
        assert_eq!(CacheSelection::parse("warm").unwrap().names(), ["warm"]);
        assert_eq!(
            CacheSelection::parse("warm,cold-requested")
                .unwrap()
                .names(),
            ["warm", "cold-requested"]
        );
        assert_eq!(
            CacheSelection::parse("cold-requested").unwrap().names(),
            ["cold-requested"]
        );
        assert_eq!(
            CacheSelection::parse("cold-verified").unwrap().names(),
            ["cold-verified"]
        );
        assert_eq!(
            CacheSelection::parse("warm,cold-requested,cold-verified")
                .unwrap()
                .names(),
            ["warm", "cold-requested", "cold-verified"]
        );
        assert!(CacheSelection::parse("hot").is_err());
        assert!(CacheSelection::parse("").is_err());
    }

    #[test]
    fn prepared_query_controls_are_not_cold_verified() {
        assert!(!Operation::PptxSourceSelectedSlide.supports_cold_verified());
        assert!(!Operation::DocxSourceFullText.supports_cold_verified());
        assert!(Operation::PptxSourceOpen.supports_cold_verified());
        assert!(Operation::DocxSourceOpenFullTextLifecycle.supports_cold_verified());
        assert!(Operation::OpcSourceOpen.supports_cold_verified());
        assert!(Operation::XlsxFileOpen.supports_cold_verified());
        assert!(Operation::XlsxFileOpenLifecycle.supports_cold_verified());
    }

    #[test]
    fn xlsx_filesystem_operations_preserve_open_vs_lifecycle_scopes() {
        let open = Operation::parse("xlsx_file_open").expect("XLSX open selector parses");
        let lifecycle =
            Operation::parse("xlsx_file_open_lifecycle").expect("XLSX lifecycle selector parses");
        assert!(open.is_xlsx());
        assert!(lifecycle.is_xlsx());
        assert_eq!(open.case().name(), "xlsx_file_open");
        assert_eq!(lifecycle.case().name(), "xlsx_file_open_lifecycle");
        assert_eq!(Operation::parse(open.case().name()), Some(open));
        assert_eq!(Operation::parse(lifecycle.case().name()), Some(lifecycle));
        assert!(!open.is_save());
        assert!(!lifecycle.is_save());
    }

    #[test]
    fn xlsx_semantic_oracle_hash_is_deterministic() {
        let corpus = crate::build_xlsx_cell_crud_corpus(crate::XlsxCellCrudShape::Medium)
            .expect("deterministic XLSX corpus builds");
        let first = xlsx_semantic_sha256(&corpus).expect("XLSX semantic oracle hashes");
        let second = xlsx_semantic_sha256(&corpus).expect("XLSX semantic oracle rehashes");
        assert_eq!(first, second);
        assert_eq!(first.len(), 64);
        assert!(first.bytes().all(|byte| byte.is_ascii_hexdigit()));
    }

    #[test]
    fn verified_digest_is_requested_only_for_eligible_save_operations() {
        let path = std::env::temp_dir().join(format!(
            "litchi-perf-verified-digest-test-{}-{}",
            std::process::id(),
            super::SystemTime::now()
                .duration_since(super::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::write(&path, []).unwrap();
        let cache_selection = CacheSelection::parse("cold-verified").unwrap();
        let eligible = Some(super::cold_verified::Status::Eligible);

        for operation in [
            Operation::OpcEagerOpen,
            Operation::OpcSourceOpen,
            Operation::PptxEagerOpen,
            Operation::PptxSourceOpen,
            Operation::PptxEagerOpenSlideCountLifecycle,
            Operation::PptxSourceOpenSlideCountLifecycle,
            Operation::PptxEagerOpenSelectedSlideLifecycle,
            Operation::PptxSourceOpenSelectedSlideLifecycle,
            Operation::DocxEagerOpen,
            Operation::DocxSourceOpen,
            Operation::DocxEagerOpenParagraphCountLifecycle,
            Operation::DocxSourceOpenParagraphCountLifecycle,
            Operation::DocxEagerOpenFullTextLifecycle,
            Operation::DocxSourceOpenFullTextLifecycle,
        ] {
            assert!(operation.supports_cold_verified());
            assert!(!operation.is_save());
            assert_eq!(
                super::verified_expected_digest(
                    operation,
                    cache_selection,
                    eligible,
                    Some(Path::new(&path)),
                )
                .unwrap(),
                None,
                "{operation:?} must not request an output digest"
            );
        }

        for operation in [
            Operation::OpcEagerSave,
            Operation::OpcSourceSave,
            Operation::CfbOverlaySave,
            Operation::CfbOwnedOverlaySave,
        ] {
            assert!(operation.is_save());
            assert!(
                super::verified_expected_digest(
                    operation,
                    cache_selection,
                    eligible,
                    Some(Path::new(&path)),
                )
                .is_err(),
                "{operation:?} must request its output digest"
            );
        }

        fs::remove_file(path).unwrap();
    }

    #[test]
    fn padded_docx_archive_keeps_semantic_corpus_identity() {
        let corpus = crate::build_docx_source_edit_corpus().unwrap();
        let aligned =
            super::cold_verified::page_aligned_archive(&corpus.archive, 4096, true).unwrap();
        assert_eq!(aligned.len() % 4096, 0);
        let corpus_signature = super::docx_archive_signature(&corpus.archive).unwrap();
        let aligned_signature = super::docx_archive_signature(&aligned).unwrap();
        assert_eq!(
            aligned_signature.semantic_sha256,
            corpus_signature.semantic_sha256
        );
        assert_eq!(aligned_signature.physical_bytes, aligned.len());
    }

    #[test]
    fn padded_xlsx_archive_keeps_typed_semantics_and_reports_aligned_hash() {
        let corpus = crate::build_xlsx_cell_crud_corpus(crate::XlsxCellCrudShape::Medium).unwrap();
        super::assert_pinned_xlsx_corpus(&corpus).unwrap();
        let aligned =
            super::cold_verified::page_aligned_archive(&corpus.archive, 4096, true).unwrap();
        assert_eq!(aligned.len() % 4096, 0);
        assert_ne!(aligned, corpus.archive);

        let package = super::OpcPackage::from_bytes(&aligned).unwrap();
        assert!(package.part_count() > 0);
        let spec = corpus.xlsx.as_ref().unwrap();
        let typed = litchi_xlsx::Workbook::from_bytes(aligned.clone()).unwrap();
        crate::verify_xlsx_cells(&typed, spec, &[]).unwrap();
        let expected = super::xlsx_semantic_sha256(&corpus).unwrap();

        let path = env::temp_dir().join(format!(
            "litchi-perf-xlsx-padding-{}-{}",
            std::process::id(),
            super::SystemTime::now()
                .duration_since(super::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::write(&path, &aligned).unwrap();
        let deferred = super::DeferredXlsxOperation {
            workbook: litchi::Workbook::open(&path).unwrap(),
            timed_projection: None,
        };
        let (source_sha256, semantic_sha256) = super::verify_xlsx_operation(
            super::Operation::XlsxFileOpen,
            &path,
            &corpus,
            true,
            &deferred,
        )
        .unwrap();
        assert_eq!(source_sha256, crate::sha256_hex(&aligned));
        assert_eq!(semantic_sha256, expected);
        drop(deferred);
        fs::remove_file(path).unwrap();
    }

    #[test]
    fn pinned_filesystem_hash_literals_are_complete() {
        for hash in [
            super::OPC_FILE_SOURCE_SHA256,
            super::OPC_FILE_EXPECTED_OUTPUT_SHA256,
            super::CFB_FILE_SOURCE_SHA256,
            super::CFB_FILE_EXPECTED_OUTPUT_SHA256,
            super::DOCX_FILE_SOURCE_SHA256,
            super::XLSX_FILE_SOURCE_SHA256,
        ] {
            assert_eq!(hash.len(), 64);
            assert!(hash.bytes().all(|byte| byte.is_ascii_hexdigit()));
        }
        assert_eq!(super::XLSX_FILE_SOURCE_ARCHIVE_BYTES, 4_226_429);
        assert_eq!(super::XLSX_FILE_SOURCE_SHAPE, "medium");
        assert_eq!(
            super::XLSX_FILE_CORPUS_GENERATOR,
            "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1"
        );
    }

    #[test]
    fn host_evidence_is_non_path_and_reports_same_device_probe() {
        let root = std::env::temp_dir();
        let evidence = super::host_evidence(Some(&root), true);
        assert_eq!(evidence.storage_identifier, None);
        #[cfg(unix)]
        assert_eq!(evidence.source_destination_same_device, Some(true));
        #[cfg(not(unix))]
        assert_eq!(evidence.source_destination_same_device, None);
        let serialized = serde_json::to_string(&evidence).unwrap();
        assert!(!serialized.contains(root.to_string_lossy().as_ref()));
        assert!(evidence.os.is_some());
    }

    #[test]
    fn child_output_keeps_allocation_sample_additive_and_omits_error_path_absence() {
        let child = |allocation_metrics| super::ChildResult {
            elapsed_ns: 1,
            logical_read_counter_scope: "timed_read_at".to_owned(),
            logical_read_calls: 0,
            logical_read_requested_bytes: 0,
            logical_read_bytes: 0,
            logical_read_largest_requested_bytes: 0,
            logical_read_largest_returned_bytes: 0,
            logical_read_pattern: None,
            max_concurrent_reads: 0,
            logical_read_request_sizes: Vec::new(),
            logical_read_request_size_buckets: ReadSizeBuckets::default(),
            cold_advice: ColdAdvice::NotRequested,
            cold_verified: None,
            process_metrics: None,
            allocation_metrics,
            output_sha256: None,
            output_bytes: None,
            opc_materialized_parts: None,
            cfb_changed_spans: None,
            cfb_published_bytes: None,
            cfb_phases: None,
            cfb_owned: None,
            pptx_source_replay: None,
            docx_source_replay: None,
            xlsx_source_sha256: None,
            xlsx_semantic_sha256: None,
        };
        let sample = crate::allocation_metrics::Sample {
            status: crate::allocation_metrics::Status::Measured,
            scope: crate::allocation_metrics::Scope::OperationGlobalSystemAllocator,
            allocation_calls: Some(1),
            deallocation_calls: Some(1),
            reallocation_calls: Some(0),
            failed_allocation_calls: Some(0),
            allocated_bytes: Some(16),
            deallocated_bytes: Some(16),
            live_bytes_before: Some(100),
            live_bytes_after: Some(100),
            peak_live_bytes_before: Some(128),
            peak_live_bytes_after: Some(128),
        };
        let value = serde_json::to_value(child(Some(sample))).unwrap();
        assert_eq!(value["allocation_metrics"]["status"], "measured");
        assert_eq!(
            value["allocation_metrics"]["live_bytes_before"],
            serde_json::Value::from(100_u64)
        );
        let absent = serde_json::to_value(child(None)).unwrap();
        assert!(absent.get("allocation_metrics").is_none());
    }

    #[test]
    fn executing_failing_child_exits_without_success_json() {
        let output = Command::new(env::current_exe().unwrap())
            .args([
                "--exact",
                "filesystem::tests::failing_child_probe",
                "--nocapture",
            ])
            .env("LITCHI_PERF_FAILING_CHILD_PROBE", "1")
            .output()
            .unwrap();
        assert!(!output.status.success());
        let parsed = serde_json::from_slice::<serde_json::Value>(&output.stdout);
        let success_protocol = parsed
            .as_ref()
            .ok()
            .and_then(|value| value.get("elapsed_ns"))
            .is_some();
        assert!(!success_protocol, "failing child emitted success protocol");
    }

    #[test]
    fn failing_child_probe() {
        if env::var_os("LITCHI_PERF_FAILING_CHILD_PROBE").is_none() {
            return;
        }
        let error = super::run_child_arguments([
            OsString::from("--filesystem-child"),
            OsString::from("opc_file_eager_open"),
            OsString::from("/definitely/missing/litchi-perf-source"),
            OsString::from("/definitely/missing/litchi-perf-destination"),
            OsString::from("warm"),
        ])
        .unwrap_err();
        panic!("{error}");
    }

    #[test]
    fn source_hash_guard_rejects_mutation_between_samples() {
        let path = std::env::temp_dir().join(format!(
            "litchi-perf-source-hash-test-{}-{}",
            std::process::id(),
            super::SystemTime::now()
                .duration_since(super::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::write(&path, b"source").unwrap();
        let expected = "41cf6794ba4200b839c53531555f0f3998df4cbb01a4d5cb0b94e3ca5e23947d";
        super::assert_source_sha256(&path, expected).unwrap();
        fs::write(&path, b"mutated").unwrap();
        let error = super::assert_source_sha256(&path, expected).unwrap_err();
        assert!(error.to_string().contains("filesystem source changed"));
        fs::remove_file(path).unwrap();
    }
}
