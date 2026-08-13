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
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::{
        Arc, Mutex,
        atomic::{AtomicU64, Ordering},
    },
    time::{Instant, SystemTime, UNIX_EPOCH},
};

use litchi_cfb::{OleFile, OverlayLimits, SameLengthStreamOverlay, SharedOleFile};
use litchi_core::{FileSource, ReadAt, SourceVersion};
use litchi_opc::{OpcPackage, PackURI, SourceBackedPackage};
use serde::{Deserialize, Serialize};

use crate::process_metrics;

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
const FILESYSTEM_OLE_COMMON_REPLACEMENT: &[u8] = b"litchi-ole-common-modified-stream-v1";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct CacheSelection {
    warm: bool,
    cold_requested: bool,
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
        };
        for state in value.split(',') {
            match state {
                "warm" => selection.warm = true,
                "cold-requested" => selection.cold_requested = true,
                _ => {
                    return Err(format!(
                        "invalid --filesystem-cache state {state:?}; expected warm or cold-requested"
                    )
                    .into());
                },
            }
        }
        if !selection.warm && !selection.cold_requested {
            return Err("--filesystem-cache selection must include warm or cold-requested".into());
        }
        Ok(selection)
    }

    pub(crate) const fn warm(self) -> bool {
        self.warm
    }

    pub(crate) const fn cold_requested(self) -> bool {
        self.cold_requested
    }

    pub(crate) fn names(self) -> Vec<&'static str> {
        let mut names = Vec::with_capacity(2);
        if self.warm {
            names.push("warm");
        }
        if self.cold_requested {
            names.push("cold-requested");
        }
        names
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct ChildResult {
    elapsed_ns: u64,
    logical_read_calls: u64,
    logical_read_requested_bytes: u64,
    logical_read_bytes: u64,
    max_concurrent_reads: u64,
    logical_read_request_sizes: Vec<u64>,
    logical_read_request_size_buckets: ReadSizeBuckets,
    cold_advice: ColdAdvice,
    process_metrics: Option<process_metrics::Delta>,
    output_sha256: Option<String>,
    output_bytes: Option<u64>,
    opc_materialized_parts: Option<u64>,
    cfb_changed_spans: Option<u64>,
    cfb_published_bytes: Option<u64>,
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
    Warm,
    Cold,
}

impl ChildMode {
    const fn name(self) -> &'static str {
        match self {
            Self::Prime => "prime",
            Self::Warm => "warm",
            Self::Cold => "cold",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "prime" => Some(Self::Prime),
            "warm" => Some(Self::Warm),
            "cold" => Some(Self::Cold),
            _ => None,
        }
    }
}

#[derive(Debug, Serialize)]
pub(crate) struct SampleEvidence {
    pub sample_index: usize,
    pub cache_state: &'static str,
    pub elapsed_ns: u64,
    pub parent_wall_ns: u64,
    pub cold_advice: ColdAdvice,
    pub logical_read_calls: u64,
    pub logical_read_requested_bytes: u64,
    pub logical_read_bytes: u64,
    pub max_concurrent_reads: u64,
    pub logical_read_request_sizes: Vec<u64>,
    pub logical_read_request_size_buckets: ReadSizeBuckets,
    pub process_metrics: Option<process_metrics::Delta>,
    pub output_sha256: Option<String>,
    pub output_bytes: Option<u64>,
    pub opc_materialized_parts: Option<u64>,
    pub cfb_changed_spans: Option<u64>,
    pub cfb_published_bytes: Option<u64>,
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
    fn observe(&mut self, bytes: u64) {
        match bytes {
            0 => self.bytes_0 += 1,
            1..=512 => self.bytes_1_to_512 += 1,
            513..=4096 => self.bytes_513_to_4096 += 1,
            4097..=16384 => self.bytes_4097_to_16384 += 1,
            16385..=65536 => self.bytes_16385_to_65536 += 1,
            _ => self.bytes_over_65536 += 1,
        }
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
}

pub(crate) struct Run {
    pub warm_result: Option<super::CaseResult>,
    pub cold_result: Option<super::CaseResult>,
    pub evidence: Evidence,
}

#[derive(Default)]
struct OperationDetails {
    opc_materialized_parts: Option<u64>,
    cfb_changed_spans: Option<u64>,
    cfb_published_bytes: Option<u64>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    OpcEagerOpen,
    OpcSourceOpen,
    OpcEagerSave,
    OpcSourceSave,
    CfbOverlaySave,
}

impl Operation {
    fn parse(name: &str) -> Option<Self> {
        match name {
            "opc_file_eager_open" => Some(Self::OpcEagerOpen),
            "opc_file_source_open" => Some(Self::OpcSourceOpen),
            "opc_file_eager_one_part_atomic_save" => Some(Self::OpcEagerSave),
            "opc_file_source_one_part_atomic_save" => Some(Self::OpcSourceSave),
            "cfb_file_same_length_overlay_atomic_save" => Some(Self::CfbOverlaySave),
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
        }
    }

    const fn is_save(self) -> bool {
        matches!(
            self,
            Self::OpcEagerSave | Self::OpcSourceSave | Self::CfbOverlaySave
        )
    }

    const fn is_cfb(self) -> bool {
        matches!(self, Self::CfbOverlaySave)
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
        assert_pinned_corpora(&opc, &cfb)?;
        let mut runs = Vec::with_capacity(selected.len());
        let mut opc_save_hashes: Option<Vec<(String, String)>> = None;
        for (case, operation) in selected {
            let corpus = if operation.is_cfb() { &cfb } else { &opc };
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
    let source_path = root.join(format!("{stem}.source"));
    write_synced(&source_path, &corpus.archive)?;
    assert_source_sha256(&source_path, &corpus.manifest.archive_sha256)?;
    let destination_path = root.join(format!("{stem}.destination"));
    let expected_digest = operation
        .is_save()
        .then(|| expected_digest(operation, corpus))
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
            &corpus.manifest.archive_sha256,
        )?;
    }

    let mut warm_elapsed = Vec::with_capacity(samples);
    let mut cold_elapsed = Vec::with_capacity(samples);
    let mut sample_evidence = Vec::with_capacity(samples * 2);
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
                &corpus.manifest.archive_sha256,
            )?;
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let warm = spawn_checked_child(
                operation,
                &source_path,
                &destination_path,
                ChildMode::Warm,
                &corpus.manifest.archive_sha256,
            )?;
            verify_child_output(operation, &source_path, &destination_path, corpus)?;
            warm_elapsed.push(warm.child.elapsed_ns);
            record_sample(
                &mut sample_evidence,
                sample_index,
                "warm",
                warm,
                operation,
                expected_digest.as_deref(),
                stem,
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
                &corpus.manifest.archive_sha256,
            )?;
            if operation.is_save() {
                seed_destination(&destination_path, &corpus.archive)?;
            }
            let cold = spawn_checked_child(
                operation,
                &source_path,
                &destination_path,
                ChildMode::Cold,
                &corpus.manifest.archive_sha256,
            )?;
            verify_child_output(operation, &source_path, &destination_path, corpus)?;
            cold_elapsed.push(cold.child.elapsed_ns);
            record_sample(
                &mut sample_evidence,
                sample_index,
                "cold-requested",
                cold,
                operation,
                expected_digest.as_deref(),
                stem,
            )?;
        }
    }

    Ok(Run {
        warm_result: cache_selection.warm().then(|| {
            filesystem_result(case, corpus, warm_elapsed, "warm", expected_digest.clone())
        }),
        cold_result: cache_selection.cold_requested().then(|| {
            filesystem_result(
                case,
                corpus,
                cold_elapsed,
                "cold-requested",
                expected_digest,
            )
        }),
        evidence: Evidence {
            case: case.name(),
            corpus: corpus.manifest.clone(),
            warmup_iterations,
            sample_count: samples,
            cache_states: cache_selection.names(),
            fresh_child_per_sample: true,
            samples: sample_evidence,
        },
    })
}

fn record_sample(
    samples: &mut Vec<SampleEvidence>,
    sample_index: usize,
    cache_state: &'static str,
    invocation: Invocation,
    operation: Operation,
    expected_digest: Option<&str>,
    stem: &str,
) -> Result<(), Box<dyn Error>> {
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
    samples.push(SampleEvidence {
        sample_index,
        cache_state,
        elapsed_ns: invocation.child.elapsed_ns,
        parent_wall_ns: invocation.parent_wall_ns,
        cold_advice: invocation.child.cold_advice,
        logical_read_calls: invocation.child.logical_read_calls,
        logical_read_requested_bytes: invocation.child.logical_read_requested_bytes,
        logical_read_bytes: invocation.child.logical_read_bytes,
        max_concurrent_reads: invocation.child.max_concurrent_reads,
        logical_read_request_sizes: invocation.child.logical_read_request_sizes,
        logical_read_request_size_buckets: invocation.child.logical_read_request_size_buckets,
        process_metrics: invocation.child.process_metrics,
        output_sha256: invocation.child.output_sha256,
        output_bytes: invocation.child.output_bytes,
        opc_materialized_parts: invocation.child.opc_materialized_parts,
        cfb_changed_spans: invocation.child.cfb_changed_spans,
        cfb_published_bytes: invocation.child.cfb_published_bytes,
    });
    Ok(())
}

fn filesystem_result(
    case: super::Case,
    corpus: &super::Corpus,
    elapsed: Vec<u64>,
    cache_state: &'static str,
    output_sha256: Option<String>,
) -> super::CaseResult {
    let mut result = super::result(case, corpus, elapsed, None);
    result.cache_state = Some(cache_state);
    result.output_sha256 = output_sha256;
    result
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
        Operation::CfbOverlaySave => expected_cfb_output_digest(corpus),
        Operation::OpcEagerOpen | Operation::OpcSourceOpen => {
            Err("open operation has no output digest".into())
        },
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
    let mut arguments = env::args_os().skip(1);
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
    .ok_or("filesystem child mode must be prime, warm, or cold")?;
    if arguments.next().is_some() {
        return Err("filesystem child received unexpected arguments".into());
    }
    let operation = Operation::parse(&case_name).ok_or("unknown filesystem child case")?;
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
    let before = process_metrics::Snapshot::read().ok();
    let started = Instant::now();
    let mut details = OperationDetails::default();
    let counter = match operation {
        Operation::OpcEagerOpen => run_opc_eager_open(&source, &mut details)?,
        Operation::OpcSourceOpen => run_opc_source_open(&source, &mut details)?,
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
    };
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos())?;
    let after = process_metrics::Snapshot::read().ok();
    let process_delta = before.zip(after).map(|(before, after)| after.delta(before));
    let snapshot = counter.map_or_else(ReadMetrics::default, |counter| counter.snapshot());

    // Correctness and hashing are intentionally after the timed operation and
    // after the operation-only counters have been sampled.
    let corpus = filesystem_corpus(operation)?;
    verify_child_output(operation, &source, &destination, &corpus)?;
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
        logical_read_calls: snapshot.calls,
        logical_read_requested_bytes: snapshot.requested_bytes,
        logical_read_bytes: snapshot.returned_bytes,
        max_concurrent_reads: snapshot.max_concurrent,
        logical_read_request_sizes: snapshot.request_sizes,
        logical_read_request_size_buckets: snapshot.request_size_buckets,
        cold_advice,
        process_metrics: process_delta,
        output_sha256,
        output_bytes,
        opc_materialized_parts: details.opc_materialized_parts,
        cfb_changed_spans: details.cfb_changed_spans,
        cfb_published_bytes: details.cfb_published_bytes,
    };
    serde_json::to_writer(io::stdout().lock(), &result)?;
    Ok(true)
}

/// Rebuilds the deterministic oracle only after the measured operation and
/// procfs snapshots have completed. The child therefore does not retain the
/// multi-megabyte synthetic corpus while collecting VmHWM.
fn filesystem_corpus(operation: Operation) -> Result<super::Corpus, Box<dyn Error>> {
    let opc = super::build_opc_corpus(OPC_FILE_SHAPE, OPC_FILE_PAYLOAD)?;
    if operation.is_cfb() {
        super::build_ole_common_corpus(&opc)
    } else {
        Ok(opc)
    }
}

#[derive(Clone, Debug, Default)]
struct ReadMetrics {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    max_concurrent: u64,
    request_sizes: Vec<u64>,
    request_size_buckets: ReadSizeBuckets,
}

struct CountingReadAt {
    inner: Arc<dyn ReadAt>,
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    in_flight: AtomicU64,
    max_concurrent: AtomicU64,
    request_sizes: Mutex<Vec<u64>>,
}

impl CountingReadAt {
    fn new(inner: Arc<dyn ReadAt>) -> Self {
        Self {
            inner,
            calls: AtomicU64::new(0),
            requested_bytes: AtomicU64::new(0),
            returned_bytes: AtomicU64::new(0),
            in_flight: AtomicU64::new(0),
            max_concurrent: AtomicU64::new(0),
            request_sizes: Mutex::new(Vec::new()),
        }
    }

    fn snapshot(&self) -> ReadMetrics {
        let mut request_sizes = self
            .request_sizes
            .lock()
            .map(|sizes| sizes.clone())
            .unwrap_or_default();
        request_sizes.sort_unstable();
        let mut request_size_buckets = ReadSizeBuckets::default();
        for &size in &request_sizes {
            request_size_buckets.observe(size);
        }
        ReadMetrics {
            calls: self.calls.load(Ordering::SeqCst),
            requested_bytes: self.requested_bytes.load(Ordering::SeqCst),
            returned_bytes: self.returned_bytes.load(Ordering::SeqCst),
            max_concurrent: self.max_concurrent.load(Ordering::SeqCst),
            request_sizes,
            request_size_buckets,
        }
    }
}

struct InFlight<'a>(&'a AtomicU64);

impl Drop for InFlight<'_> {
    fn drop(&mut self) {
        self.0.fetch_sub(1, Ordering::SeqCst);
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let requested = u64::try_from(output.len()).unwrap_or(u64::MAX);
        self.calls.fetch_add(1, Ordering::SeqCst);
        self.requested_bytes.fetch_add(requested, Ordering::SeqCst);
        if let Ok(mut sizes) = self.request_sizes.lock() {
            sizes.push(requested);
        }
        let in_flight = self.in_flight.fetch_add(1, Ordering::SeqCst) + 1;
        self.max_concurrent.fetch_max(in_flight, Ordering::SeqCst);
        let _guard = InFlight(&self.in_flight);
        let read = self.inner.read_at(offset, output)?;
        self.returned_bytes
            .fetch_add(u64::try_from(read).unwrap_or(u64::MAX), Ordering::SeqCst);
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
    details: &mut OperationDetails,
) -> Result<Option<Arc<CountingReadAt>>, Box<dyn Error>> {
    let counter = Arc::new(CountingReadAt::new(Arc::new(FileSource::open(source)?)));
    let package = SourceBackedPackage::from_read_at(counter.clone())?;
    details.opc_materialized_parts = Some(package.cache_diagnostics().successful_loads);
    std::hint::black_box(&package);
    Ok(Some(counter))
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
    let counter = Arc::new(CountingReadAt::new(Arc::new(FileSource::open(source)?)));
    let shared = SharedOleFile::open(counter.clone())?;
    let overlay = SameLengthStreamOverlay::new(
        vec![super::OLE_COMMON_TARGET.to_owned()],
        Arc::from(FILESYSTEM_OLE_COMMON_REPLACEMENT.to_vec()),
    );
    let plan = shared.plan_same_length_stream_overlays(vec![overlay], OverlayLimits::default())?;
    let report = plan.save(destination)?;
    details.cfb_changed_spans = Some(u64::try_from(report.changed_spans())?);
    details.cfb_published_bytes = Some(report.bytes());
    Ok(Some(counter))
}

fn verify_child_output(
    operation: Operation,
    source: &Path,
    destination: &Path,
    corpus: &super::Corpus,
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
        Operation::CfbOverlaySave => {
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
    }
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
    use std::fs;

    use super::{CacheSelection, ChildMode, ColdAdvice, Operation, ReadSizeBuckets};

    #[test]
    fn filesystem_case_names_are_explicit_and_parseable() {
        for name in [
            "opc_file_eager_open",
            "opc_file_source_open",
            "opc_file_eager_one_part_atomic_save",
            "opc_file_source_one_part_atomic_save",
            "cfb_file_same_length_overlay_atomic_save",
        ] {
            assert!(Operation::parse(name).is_some(), "{name}");
        }
        assert!(Operation::parse("ole2_same_length_overlay_atomic_save").is_none());
    }

    #[test]
    fn read_size_buckets_are_fixed_and_boundary_stable() {
        let mut buckets = ReadSizeBuckets::default();
        for size in [0, 1, 512, 513, 4096, 4097, 16384, 16385, 65536, 65537] {
            buckets.observe(size);
        }
        assert_eq!(buckets.bytes_0, 1);
        assert_eq!(buckets.bytes_1_to_512, 2);
        assert_eq!(buckets.bytes_513_to_4096, 2);
        assert_eq!(buckets.bytes_4097_to_16384, 2);
        assert_eq!(buckets.bytes_16385_to_65536, 2);
        assert_eq!(buckets.bytes_over_65536, 1);
    }

    #[test]
    fn cold_state_labels_are_distinct_from_warm_and_unsupported() {
        assert_ne!(ColdAdvice::NotRequested as u8, ColdAdvice::Requested as u8);
        assert_eq!(ChildMode::parse("cold"), Some(ChildMode::Cold));
        assert_eq!(ChildMode::parse("warm"), Some(ChildMode::Warm));
        assert_eq!(ChildMode::parse("prime"), Some(ChildMode::Prime));
    }

    #[test]
    fn cache_selection_is_explicit_and_additive() {
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
        assert!(CacheSelection::parse("hot").is_err());
        assert!(CacheSelection::parse("").is_err());
    }

    #[test]
    fn pinned_filesystem_hash_literals_are_complete() {
        for hash in [
            super::OPC_FILE_SOURCE_SHA256,
            super::OPC_FILE_EXPECTED_OUTPUT_SHA256,
            super::CFB_FILE_SOURCE_SHA256,
            super::CFB_FILE_EXPECTED_OUTPUT_SHA256,
        ] {
            assert_eq!(hash.len(), 64);
            assert!(hash.bytes().all(|byte| byte.is_ascii_hexdigit()));
        }
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
