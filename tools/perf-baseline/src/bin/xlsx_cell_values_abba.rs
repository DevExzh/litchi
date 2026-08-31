use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::fs::{self, File, OpenOptions};
use std::io::{BufWriter, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

use serde::{Deserialize, Serialize};
use serde_json::{Map, Value, json};
use sha2::{Digest, Sha256};

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx.cell-values-abba.v1";
const BENCHMARK_NAME: &str = "litchi-perf-baseline";
const CORPUS_GENERATOR: &str = "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1";
const EVIDENCE_MIN_SAMPLES: usize = 20;
const DEFAULT_WARMUP: usize = 20;
const DEFAULT_SAMPLES: usize = 20;
const MAX_WARMUP: usize = 100;
const MAX_SAMPLES: usize = 100_000;
const MAX_CHILD_JSON_BYTES: u64 = 64 * 1024 * 1024;
const MAX_ROW_BYTES: usize = 1 * 1024 * 1024;
const MAX_FAILURE_ROWS: usize = 10_000;
const MAX_WRITE_BYTES: u64 = 64 * 1024;
const SAME_SIDE_GATE_PERCENT: f64 = 5.0;
const DIRECTIONAL_IMPROVEMENT_PERCENT: f64 = 1.0;
const DIRECTIONAL_IMPROVEMENT_NS: u64 = 50_000;
const DIRECTIONAL_ADVERSE_PERCENT: f64 = 5.0;

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(1);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Workload {
    OneEdit,
    OnePercent,
    Batch,
}

impl Workload {
    fn name(self) -> &'static str {
        match self {
            Self::OneEdit => "one_edit",
            Self::OnePercent => "one_percent",
            Self::Batch => "batch",
        }
    }

    const fn case_suffix(self) -> &'static str {
        match self {
            Self::OneEdit => "one_edit_save",
            Self::OnePercent => "one_percent_edit_save",
            Self::Batch => "batch_edit_save",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Implementation {
    Eager,
    Source,
}

impl Implementation {
    const fn name(self) -> &'static str {
        match self {
            Self::Eager => "eager",
            Self::Source => "source-backed",
        }
    }

    const fn case_prefix(self) -> &'static str {
        match self {
            Self::Eager => "xlsx_eager_cell_values_",
            Self::Source => "xlsx_source_backed_cell_values_",
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct CellSpec {
    shape: &'static str,
    workload: Workload,
}

impl CellSpec {
    fn name(self) -> &'static str {
        match (self.shape, self.workload) {
            ("medium", Workload::OneEdit) => "medium/one_edit",
            ("medium", Workload::OnePercent) => "medium/one_percent",
            ("medium", Workload::Batch) => "medium/batch",
            ("dense-sparse", Workload::OneEdit) => "dense-sparse/one_edit",
            ("dense-sparse", Workload::OnePercent) => "dense-sparse/one_percent",
            ("dense-sparse", Workload::Batch) => "dense-sparse/batch",
            _ => "unknown",
        }
    }
}

const CELLS: [CellSpec; 6] = [
    CellSpec {
        shape: "medium",
        workload: Workload::OneEdit,
    },
    CellSpec {
        shape: "medium",
        workload: Workload::OnePercent,
    },
    CellSpec {
        shape: "medium",
        workload: Workload::Batch,
    },
    CellSpec {
        shape: "dense-sparse",
        workload: Workload::OneEdit,
    },
    CellSpec {
        shape: "dense-sparse",
        workload: Workload::OnePercent,
    },
    CellSpec {
        shape: "dense-sparse",
        workload: Workload::Batch,
    },
];

const LEGS: [(&str, Implementation); 4] = [
    ("A1", Implementation::Eager),
    ("B1", Implementation::Source),
    ("B2", Implementation::Source),
    ("A2", Implementation::Eager),
];

#[derive(Debug)]
enum Arguments {
    Driver {
        benchmark_bin: Option<PathBuf>,
        benchmark_sha256: Option<String>,
        corpus_identities: Option<PathBuf>,
        warmup: usize,
        samples: usize,
        out_dir: PathBuf,
    },
    PrintProtocol,
}

#[derive(Clone, Debug, Serialize)]
struct BinaryIdentity {
    path: String,
    sha256: String,
    bytes: u64,
    profile: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct RunnerIdentity {
    executable_sha256: String,
    git_revision: Option<String>,
    git_dirty: Option<bool>,
    rustc_vv: Option<String>,
    os: String,
    arch: String,
    cpu: Option<String>,
    memory: Option<String>,
    rustflags: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct CellProtocol {
    cell: &'static str,
    shape: &'static str,
    workload: &'static str,
    eager_case: String,
    source_case: String,
    expected_update_count: usize,
    expected_touched_worksheets: usize,
}

#[derive(Clone, Debug, Serialize)]
struct Protocol {
    schema: &'static str,
    tool: &'static str,
    package_version: &'static str,
    mode: &'static str,
    evidence_min_samples: usize,
    statistical_gates_enforced: bool,
    warmup: usize,
    samples: usize,
    fresh_child_per_leg: bool,
    leg_order: [&'static str; 4],
    benchmark_bin: BinaryIdentity,
    runner_identity: RunnerIdentity,
    revision_scope: &'static str,
    claim_scope: &'static str,
    timing_scope: &'static str,
    cells: Vec<CellProtocol>,
    corpus_generator: &'static str,
    validations: Vec<&'static str>,
    unavailable_field_policy: &'static str,
    exclusions: Vec<&'static str>,
    corpora: BTreeMap<String, CorpusIdentity>,
    claim_blockers: Vec<String>,
    performance_claim: &'static str,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct SourceCounters {
    values: BTreeMap<String, Value>,
    complete: bool,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct SinkEvidence {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    retained_output_bytes: Option<u64>,
}

#[derive(Clone, Debug, Deserialize, Serialize, PartialEq, Eq)]
struct CorpusIdentity {
    name: String,
    generator: String,
    package_format: String,
    shape: String,
    payload_kind: String,
    compression: String,
    entry_count: usize,
    archive_member_count: usize,
    entry_bytes: usize,
    uncompressed_payload_bytes: usize,
    archive_bytes: usize,
    archive_sha256: String,
    target_entry: String,
    target_payload_bytes: usize,
    target_payload_sha256: String,
    workbook_member: String,
    worksheet_members: Vec<String>,
    shared_strings_member: Option<String>,
    styles_member: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct ChildObservation {
    case: String,
    shape: String,
    workload: String,
    implementation: String,
    primary_ns: u64,
    output_sha256: String,
    semantic_sha256: Option<String>,
    untouched_member_count: Option<usize>,
    untouched_member_sha256: Option<String>,
    sink: SinkEvidence,
    source_counters: Option<SourceCounters>,
    corpus: CorpusIdentity,
    binary_identity: BinaryIdentity,
    git_revision: Option<String>,
    git_dirty: Option<bool>,
}

#[derive(Clone, Debug, Serialize)]
struct ObservedRow {
    schema: &'static str,
    sample: usize,
    cell: &'static str,
    shape: &'static str,
    workload: &'static str,
    leg: &'static str,
    implementation: &'static str,
    case: String,
    primary_ns: u64,
    output_sha256: String,
    semantic_sha256: Option<String>,
    untouched_member_count: Option<usize>,
    untouched_member_sha256: Option<String>,
    sink: SinkEvidence,
    source_counters: Option<SourceCounters>,
}

#[derive(Clone, Debug, Serialize)]
struct FailureRecord {
    sample: Option<usize>,
    warmup: bool,
    cell: String,
    leg: String,
    implementation: String,
    error: String,
}

#[derive(Clone, Debug, Default)]
struct CellSampleState {
    eager_output: Option<String>,
    source_output: Option<String>,
    source_semantic: Option<String>,
    source_untouched_count: Option<usize>,
    source_untouched_hash: Option<String>,
    source_counters: Option<SourceCounters>,
}

#[derive(Debug)]
struct StatisticsAccumulator {
    values: Vec<u64>,
}

impl StatisticsAccumulator {
    fn new(samples: usize) -> AnyResult<Self> {
        let mut values = Vec::new();
        values.try_reserve_exact(samples)?;
        Ok(Self { values })
    }
}

#[derive(Debug)]
struct CellAccumulator {
    legs: [StatisticsAccumulator; 4],
    counts: [usize; 4],
}

impl CellAccumulator {
    fn new(samples: usize) -> AnyResult<Self> {
        let legs = [
            StatisticsAccumulator::new(samples)?,
            StatisticsAccumulator::new(samples)?,
            StatisticsAccumulator::new(samples)?,
            StatisticsAccumulator::new(samples)?,
        ];
        Ok(Self {
            legs,
            counts: [0; 4],
        })
    }
}

#[derive(Clone, Debug, Serialize)]
struct StatisticsReport {
    count: usize,
    min_ns: u64,
    p50_ns: u64,
    mean_ns_floor: u64,
    p95_ns: u64,
    p99_ns: u64,
    max_ns: u64,
}

#[derive(Clone, Debug, Serialize)]
struct LegSummary {
    leg: &'static str,
    implementation: &'static str,
    statistics: Option<StatisticsReport>,
}

#[derive(Clone, Debug, Serialize)]
struct GateResult {
    name: String,
    threshold_percent: f64,
    observed_percent: Option<f64>,
    absolute_delta_ns: Option<u64>,
    passed: bool,
    enforced: bool,
    rationale: &'static str,
}

#[derive(Clone, Debug, Serialize)]
struct CellSummary {
    cell: &'static str,
    shape: &'static str,
    workload: &'static str,
    requested_samples: usize,
    successful_samples: usize,
    legs: Vec<LegSummary>,
    same_side_gates: Vec<GateResult>,
    directional_gates: Vec<GateResult>,
    all_required_gates_passed: bool,
}

#[derive(Clone, Debug, Serialize)]
struct Summary {
    schema: &'static str,
    tool: &'static str,
    package_version: &'static str,
    mode: &'static str,
    evidence_min_samples: usize,
    statistical_gates_enforced: bool,
    performance_claim: &'static str,
    cross_revision_evidence: bool,
    requested_samples: usize,
    expected_cells: usize,
    expected_rows: usize,
    successful_rows: usize,
    failure_rows: usize,
    complete_rows: bool,
    timing_gates_passed: bool,
    all_required_gates_passed: bool,
    claim_authorized: bool,
    claim_blockers: Vec<String>,
    corpora: BTreeMap<String, CorpusIdentity>,
    cells: Vec<CellSummary>,
}

#[derive(Debug, Serialize)]
struct ArtifactManifest {
    schema: &'static str,
    self_excluded: bool,
    files: Vec<ArtifactEntry>,
}

#[derive(Debug, Serialize)]
struct ArtifactEntry {
    path: String,
    bytes: u64,
    sha256: String,
}

fn main() {
    let result = match parse_arguments() {
        Ok(Arguments::PrintProtocol) => print_protocol(),
        Ok(Arguments::Driver {
            benchmark_bin,
            benchmark_sha256,
            corpus_identities,
            warmup,
            samples,
            out_dir,
        }) => driver_main(
            benchmark_bin,
            benchmark_sha256.as_deref(),
            corpus_identities.as_deref(),
            warmup,
            samples,
            &out_dir,
        ),
        Err(error) => Err(error),
    };
    if let Err(error) = result {
        eprintln!("xlsx_cell_values_abba: {error}");
        std::process::exit(1);
    }
}

fn parse_arguments() -> AnyResult<Arguments> {
    let mut benchmark_bin = None;
    let mut benchmark_sha256 = None;
    let mut corpus_identities = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    let mut out_dir = None;
    let mut print = false;
    let mut arguments = env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--benchmark-bin" => {
                benchmark_bin = Some(PathBuf::from(
                    arguments.next().ok_or("--benchmark-bin requires PATH")?,
                ));
            },
            "--benchmark-sha256" => {
                benchmark_sha256 = Some(
                    arguments
                        .next()
                        .ok_or("--benchmark-sha256 requires SHA256")?,
                );
            },
            "--corpus-identities" => {
                corpus_identities = Some(PathBuf::from(
                    arguments
                        .next()
                        .ok_or("--corpus-identities requires PATH")?,
                ));
            },
            "--warmup" => {
                warmup = arguments.next().ok_or("--warmup requires N")?.parse()?;
            },
            "--samples" => {
                samples = arguments.next().ok_or("--samples requires N")?.parse()?;
            },
            "--out-dir" => {
                out_dir = Some(PathBuf::from(
                    arguments.next().ok_or("--out-dir requires PATH")?,
                ));
            },
            "--print-protocol" => print = true,
            "--help" | "-h" => {
                println!(
                    "usage: xlsx_cell_values_abba [--benchmark-bin PATH] [--benchmark-sha256 SHA256] [--corpus-identities PATH] [--warmup N] [--samples N] --out-dir PATH\n       xlsx_cell_values_abba --print-protocol"
                );
                return Ok(Arguments::PrintProtocol);
            },
            other => return Err(format!("unknown argument: {other}").into()),
        }
    }
    if print {
        return Ok(Arguments::PrintProtocol);
    }
    if warmup > MAX_WARMUP {
        return Err(format!("--warmup must be at most {MAX_WARMUP}").into());
    }
    if samples == 0 || samples > MAX_SAMPLES {
        return Err(format!("--samples must be in 1..={MAX_SAMPLES}").into());
    }
    let out_dir = out_dir.ok_or("--out-dir is required")?;
    Ok(Arguments::Driver {
        benchmark_bin,
        benchmark_sha256,
        corpus_identities,
        warmup,
        samples,
        out_dir,
    })
}

fn print_protocol() -> AnyResult<()> {
    let cells = CELLS
        .into_iter()
        .map(|cell| {
            json!({
                "cell": cell.name(),
                "shape": cell.shape,
                "workload": cell.workload.name(),
                "eager_case": case_name(cell, Implementation::Eager),
                "source_case": case_name(cell, Implementation::Source),
                "expected_update_count": expected_update_count(cell),
                "expected_touched_worksheets": expected_touched_worksheets(cell.workload),
            })
        })
        .collect::<Vec<_>>();
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": SCHEMA,
            "tool": "xlsx_cell_values_abba",
            "evidence_min_samples": EVIDENCE_MIN_SAMPLES,
            "fresh_child_per_leg": true,
            "leg_order": ["A1 eager", "B1 source-backed", "B2 source-backed", "A2 eager"],
            "cells": cells,
            "timing_scope": "existing child elapsed_ns.samples[0] total open/edit/save elapsed; phase vectors are not compared",
            "claim_scope": "same-revision total eager-owned open/edit/save versus source-backed open/edit/one-overlay publication on the pinned synthetic corpus",
            "exclusions": exclusions(),
        }))?
    );
    Ok(())
}

fn driver_main(
    benchmark_bin: Option<PathBuf>,
    expected_benchmark_sha256: Option<&str>,
    corpus_identities: Option<&Path>,
    warmup: usize,
    samples: usize,
    out_dir: &Path,
) -> AnyResult<()> {
    let started = SystemTime::now();
    fs::create_dir_all(out_dir)?;
    let smoke = samples < EVIDENCE_MIN_SAMPLES;
    let benchmark_path = resolve_benchmark_binary(benchmark_bin)?;
    let benchmark_profile = inferred_profile(&benchmark_path);
    let benchmark_identity = binary_identity(&benchmark_path, Some(benchmark_profile))?;
    match expected_benchmark_sha256 {
        Some(expected) => {
            validate_sha256(expected, "--benchmark-sha256")?;
            if benchmark_identity.sha256 != expected {
                return Err("--benchmark-sha256 does not match --benchmark-bin".into());
            }
        },
        None if !smoke => {
            return Err("--benchmark-sha256 is required for evidence runs".into());
        },
        None => {},
    }
    if !smoke && benchmark_profile != "release" {
        return Err("claim-bearing evidence requires a release-profile benchmark binary".into());
    }
    let expected_corpora = load_expected_corpora(corpus_identities)?;
    if !smoke
        && (expected_corpora.len() != 2
            || !expected_corpora.contains_key("medium")
            || !expected_corpora.contains_key("dense-sparse"))
    {
        return Err(
            "claim-bearing evidence requires predeclared medium and dense-sparse corpus identities"
                .into(),
        );
    }
    let runner_identity = runner_identity()?;
    let mode = if smoke { "smoke" } else { "evidence" };
    let initial_claim = if smoke {
        "none: smoke correctness and diagnostic timing only"
    } else {
        "none: evidence authorization pending all cells, gates, and provenance"
    };
    let mut claim_blockers = Vec::new();
    if runner_identity.git_revision.is_none() {
        add_blocker(&mut claim_blockers, "runner git revision unavailable");
    }
    if runner_identity.git_dirty != Some(false) {
        add_blocker(
            &mut claim_blockers,
            "clean git worktree required for evidence claim",
        );
    }
    let cells = CELLS
        .into_iter()
        .map(|cell| CellProtocol {
            cell: cell.name(),
            shape: cell.shape,
            workload: cell.workload.name(),
            eager_case: case_name(cell, Implementation::Eager),
            source_case: case_name(cell, Implementation::Source),
            expected_update_count: expected_update_count(cell),
            expected_touched_worksheets: expected_touched_worksheets(cell.workload),
        })
        .collect::<Vec<_>>();
    let mut protocol = Protocol {
        schema: SCHEMA,
        tool: "xlsx_cell_values_abba",
        package_version: env!("CARGO_PKG_VERSION"),
        mode,
        evidence_min_samples: EVIDENCE_MIN_SAMPLES,
        statistical_gates_enforced: !smoke,
        warmup,
        samples,
        fresh_child_per_leg: true,
        leg_order: ["A1", "B1", "B2", "A2"],
        benchmark_bin: benchmark_identity.clone(),
        runner_identity: runner_identity.clone(),
        revision_scope: "same runner revision, benchmark executable identity, and child configuration for every fresh leg; evidence additionally requires a clean git worktree",
        claim_scope: "same-revision total eager-owned open/edit/save versus source-backed open/edit/one-overlay publication on the pinned synthetic corpus; directional A1->B1 and A2->B2 comparisons are claim-bearing within this scope",
        timing_scope: "primary_ns is exactly the existing child elapsed_ns.samples[0] total; eager/source phase subvectors are not compared",
        cells,
        corpus_generator: CORPUS_GENERATOR,
        validations: vec![
            "exact child case, shape, workload, update count, and touched-worksheet semantics",
            "complete corpus manifest, archive hash/bytes, and workbook/worksheet/member identity",
            "output, source semantic, and untouched-member hashes where the child exposes them",
            "bounded sequential sink accepted bytes, write calls, and <=64 KiB largest write",
            "source xlsx_cell_values logical counters, materializations, cache, and budget fields",
            "B1/B2 source counter neutrality and source output/semantic identity",
            "fresh-child provenance, executable identity, and clean revision for evidence",
        ],
        unavailable_field_policy: "missing fields are recorded as unavailable and block only the affected claim; no value is inferred",
        exclusions: exclusions(),
        corpora: expected_corpora.clone(),
        claim_blockers,
        performance_claim: initial_claim,
    };
    write_json(&out_dir.join("protocol.json"), &protocol)?;

    let sample_file = File::create(out_dir.join("samples.jsonl"))?;
    let failure_file = File::create(out_dir.join("failures.jsonl"))?;
    let mut state = RunState::new(
        &benchmark_path,
        &benchmark_identity,
        &runner_identity,
        expected_corpora,
        samples,
        sample_file,
        failure_file,
    )?;
    for _ in 0..warmup {
        state.run_iteration(None, true)?;
    }
    for sample in 0..samples {
        state.run_iteration(Some(sample), false)?;
    }
    state.sample_writer.flush()?;
    state.failure_writer.flush()?;

    let complete_rows = state.rows
        == samples
            .saturating_mul(CELLS.len())
            .saturating_mul(LEGS.len())
        && state
            .accumulators
            .iter()
            .all(|cell| cell.counts.iter().all(|count| *count == samples));
    if !complete_rows {
        record_failure(
            &mut state.failure_writer,
            &mut state.failure_count,
            FailureRecord {
                sample: None,
                warmup: false,
                cell: "all".to_owned(),
                leg: "validation".to_owned(),
                implementation: "driver".to_owned(),
                error: "expected exactly samples * 6 cells * 4 legs valid rows".to_owned(),
            },
        )?;
    }

    let mut summaries = Vec::new();
    summaries.try_reserve_exact(CELLS.len())?;
    let mut timing_gates_passed = true;
    let mut gates_available = true;
    for (index, cell) in CELLS.into_iter().enumerate() {
        let summary = summarize_cell(cell, &state.accumulators[index], samples, !smoke);
        timing_gates_passed &= summary.all_required_gates_passed;
        gates_available &= summary.legs.iter().all(|leg| leg.statistics.is_some());
        summaries.push(summary);
    }
    if !gates_available {
        add_blocker(
            &mut state.claim_blockers,
            "one or more aggregate timing statistics are unavailable",
        );
    }
    if !smoke && gates_available && !timing_gates_passed {
        record_failure(
            &mut state.failure_writer,
            &mut state.failure_count,
            FailureRecord {
                sample: None,
                warmup: false,
                cell: "all".to_owned(),
                leg: "gates".to_owned(),
                implementation: "driver".to_owned(),
                error: "one or more strict same-side or directional timing gates failed".to_owned(),
            },
        )?;
    }
    state.failure_writer.flush()?;
    for blocker in state.claim_blockers.drain(..) {
        add_blocker(&mut protocol.claim_blockers, &blocker);
    }
    protocol.corpora = state.corpora.clone();
    let clean_children = state.child_git_dirty == Some(false)
        && state.child_git_revision.is_some()
        && state.child_identity.is_some();
    if !clean_children {
        add_blocker(
            &mut protocol.claim_blockers,
            "child executable or clean git provenance unavailable",
        );
    }
    let claim_authorized = !smoke
        && complete_rows
        && state.failure_count == 0
        && timing_gates_passed
        && gates_available
        && protocol.claim_blockers.is_empty();
    let final_claim = if claim_authorized {
        "same-revision total eager-owned open/edit/save versus source-backed open/edit/one-overlay publication on the pinned synthetic corpus only"
    } else {
        initial_claim
    };
    protocol.performance_claim = final_claim;
    write_json(&out_dir.join("protocol.json"), &protocol)?;
    let summary = Summary {
        schema: SCHEMA,
        tool: "xlsx_cell_values_abba",
        package_version: env!("CARGO_PKG_VERSION"),
        mode,
        evidence_min_samples: EVIDENCE_MIN_SAMPLES,
        statistical_gates_enforced: !smoke,
        performance_claim: final_claim,
        cross_revision_evidence: false,
        requested_samples: samples,
        expected_cells: CELLS.len(),
        expected_rows: samples
            .checked_mul(CELLS.len())
            .and_then(|value| value.checked_mul(LEGS.len()))
            .ok_or("summary row count overflow")?,
        successful_rows: state.rows,
        failure_rows: state.failure_count,
        complete_rows,
        timing_gates_passed,
        all_required_gates_passed: timing_gates_passed && gates_available,
        claim_authorized,
        claim_blockers: protocol.claim_blockers.clone(),
        corpora: protocol.corpora.clone(),
        cells: summaries,
    };
    write_json(&out_dir.join("summary.json"), &summary)?;
    fs::write(
        out_dir.join("sha256.txt"),
        format!(
            "benchmark_sha256 {}\nbenchmark_bytes {}\nrunner_sha256 {}\n",
            benchmark_identity.sha256, benchmark_identity.bytes, runner_identity.executable_sha256,
        ),
    )?;
    fs::write(
        out_dir.join("process-time.txt"),
        process_time_text(started)?,
    )?;
    write_artifact_manifest(out_dir)?;
    if state.failure_count != 0 {
        return Err(format!(
            "ABBA run failed with {} child or gate failure(s)",
            state.failure_count
        )
        .into());
    }
    if !smoke && !claim_authorized {
        return Err(
            "evidence authorization was blocked; artifacts retained without a claim".into(),
        );
    }
    Ok(())
}

struct RunState<'a> {
    benchmark_bin: &'a Path,
    benchmark_identity: &'a BinaryIdentity,
    runner_identity: &'a RunnerIdentity,
    expected_corpora: BTreeMap<String, CorpusIdentity>,
    sample_writer: BufWriter<File>,
    failure_writer: BufWriter<File>,
    accumulators: Vec<CellAccumulator>,
    corpora: BTreeMap<String, CorpusIdentity>,
    child_identity: Option<BinaryIdentity>,
    child_git_revision: Option<String>,
    child_git_dirty: Option<bool>,
    rows: usize,
    failure_count: usize,
    claim_blockers: Vec<String>,
}

impl<'a> RunState<'a> {
    fn new(
        benchmark_bin: &'a Path,
        benchmark_identity: &'a BinaryIdentity,
        runner_identity: &'a RunnerIdentity,
        expected_corpora: BTreeMap<String, CorpusIdentity>,
        samples: usize,
        sample_file: File,
        failure_file: File,
    ) -> AnyResult<Self> {
        let mut accumulators = Vec::new();
        accumulators.try_reserve_exact(CELLS.len())?;
        for _ in CELLS {
            accumulators.push(CellAccumulator::new(samples)?);
        }
        Ok(Self {
            benchmark_bin,
            benchmark_identity,
            runner_identity,
            expected_corpora,
            sample_writer: BufWriter::new(sample_file),
            failure_writer: BufWriter::new(failure_file),
            accumulators,
            corpora: BTreeMap::new(),
            child_identity: None,
            child_git_revision: None,
            child_git_dirty: None,
            rows: 0,
            failure_count: 0,
            claim_blockers: Vec::new(),
        })
    }

    fn run_iteration(&mut self, sample: Option<usize>, warmup: bool) -> AnyResult<()> {
        for (cell_index, cell) in CELLS.into_iter().enumerate() {
            let mut sample_state = CellSampleState::default();
            for (leg_index, (leg_name, implementation)) in LEGS.into_iter().enumerate() {
                let case = case_name(cell, implementation);
                let observation = match invoke_child(self.benchmark_bin, &case, cell.shape) {
                    Ok(report) => match parse_child_report(
                        &report,
                        cell,
                        implementation,
                        self.benchmark_identity,
                        self.runner_identity,
                        &mut self.claim_blockers,
                    ) {
                        Ok(observation) => Some(observation),
                        Err(error) => {
                            self.failure(sample, warmup, cell, leg_name, implementation, error)?;
                            None
                        },
                    },
                    Err(error) => {
                        self.failure(
                            sample,
                            warmup,
                            cell,
                            leg_name,
                            implementation,
                            error.to_string(),
                        )?;
                        None
                    },
                };
                let Some(observation) = observation else {
                    continue;
                };
                self.observe_provenance(&observation, cell, leg_name, sample, warmup)?;
                self.observe_corpus(&observation, cell, leg_name, sample, warmup)?;
                self.observe_pair(
                    &mut sample_state,
                    &observation,
                    cell,
                    leg_name,
                    sample,
                    warmup,
                )?;
                if let Some(sample) = sample {
                    let row = ObservedRow {
                        schema: SCHEMA,
                        sample,
                        cell: cell.name(),
                        shape: cell.shape,
                        workload: cell.workload.name(),
                        leg: leg_name,
                        implementation: implementation.name(),
                        case: observation.case.clone(),
                        primary_ns: observation.primary_ns,
                        output_sha256: observation.output_sha256.clone(),
                        semantic_sha256: observation.semantic_sha256.clone(),
                        untouched_member_count: observation.untouched_member_count,
                        untouched_member_sha256: observation.untouched_member_sha256.clone(),
                        sink: observation.sink.clone(),
                        source_counters: observation.source_counters.clone(),
                    };
                    write_json_line(&mut self.sample_writer, &row)?;
                    self.accumulators[cell_index].legs[leg_index]
                        .values
                        .push(observation.primary_ns);
                    self.accumulators[cell_index].counts[leg_index] += 1;
                    self.rows += 1;
                }
            }
        }
        Ok(())
    }

    fn observe_provenance(
        &mut self,
        observation: &ChildObservation,
        cell: CellSpec,
        leg: &'static str,
        sample: Option<usize>,
        warmup: bool,
    ) -> AnyResult<()> {
        if observation.binary_identity.sha256 != self.benchmark_identity.sha256
            || observation.binary_identity.bytes != self.benchmark_identity.bytes
        {
            self.failure(
                sample,
                warmup,
                cell,
                leg,
                implementation_from_name(&observation.implementation),
                "child executable identity differs from --benchmark-bin".to_owned(),
            )?;
        }
        if let Some(previous) = &self.child_identity {
            if previous.sha256 != observation.binary_identity.sha256
                || previous.bytes != observation.binary_identity.bytes
                || previous.profile != observation.binary_identity.profile
            {
                self.failure(
                    sample,
                    warmup,
                    cell,
                    leg,
                    implementation_from_name(&observation.implementation),
                    "fresh child executable identities are not invariant".to_owned(),
                )?;
            }
        } else {
            self.child_identity = Some(observation.binary_identity.clone());
        }
        if let Some(previous) = &self.child_git_revision {
            if Some(previous) != observation.git_revision.as_ref() {
                self.failure(
                    sample,
                    warmup,
                    cell,
                    leg,
                    implementation_from_name(&observation.implementation),
                    "fresh child git revisions are not invariant".to_owned(),
                )?;
            }
        } else if let Some(revision) = &observation.git_revision {
            self.child_git_revision = Some(revision.clone());
        }
        if let Some(previous) = self.child_git_dirty {
            if Some(previous) != observation.git_dirty {
                self.failure(
                    sample,
                    warmup,
                    cell,
                    leg,
                    implementation_from_name(&observation.implementation),
                    "fresh child git dirty states are not invariant".to_owned(),
                )?;
            }
        } else if let Some(dirty) = observation.git_dirty {
            self.child_git_dirty = Some(dirty);
        }
        Ok(())
    }

    fn observe_corpus(
        &mut self,
        observation: &ChildObservation,
        cell: CellSpec,
        leg: &'static str,
        sample: Option<usize>,
        warmup: bool,
    ) -> AnyResult<()> {
        match self.expected_corpora.get(cell.shape) {
            Some(expected) if expected != &observation.corpus => {
                self.failure(
                    sample,
                    warmup,
                    cell,
                    leg,
                    implementation_from_name(&observation.implementation),
                    "corpus identity differs from the predeclared manifest".to_owned(),
                )?;
            },
            None if !self.expected_corpora.is_empty() => {
                self.failure(
                    sample,
                    warmup,
                    cell,
                    leg,
                    implementation_from_name(&observation.implementation),
                    "requested corpus shape is absent from the predeclared manifest".to_owned(),
                )?;
            },
            _ => {},
        }
        if let Some(previous) = self.corpora.get(cell.shape) {
            if previous != &observation.corpus {
                self.failure(
                    sample,
                    warmup,
                    cell,
                    leg,
                    implementation_from_name(&observation.implementation),
                    "corpus manifest or archive identity differs across fresh children".to_owned(),
                )?;
            }
        } else {
            self.corpora
                .insert(cell.shape.to_owned(), observation.corpus.clone());
        }
        Ok(())
    }

    fn observe_pair(
        &mut self,
        state: &mut CellSampleState,
        observation: &ChildObservation,
        cell: CellSpec,
        leg: &'static str,
        sample: Option<usize>,
        warmup: bool,
    ) -> AnyResult<()> {
        match leg {
            "A1" => state.eager_output = Some(observation.output_sha256.clone()),
            "B1" => {
                let expected_untouched_member_count = expected_cell_untouched_member_count(
                    usize::try_from(observation.corpus.archive_member_count).map_err(|_| {
                        "source corpus member count does not fit usize".to_owned()
                    })?,
                    cell.workload,
                )
                .ok_or_else(|| "source untouched-member count underflows closure".to_owned())?;
                if observation.untouched_member_count != Some(expected_untouched_member_count) {
                    self.failure(
                        sample,
                        warmup,
                        cell,
                        leg,
                        Implementation::Source,
                        format!(
                            "source untouched-member count differs: expected {expected_untouched_member_count}, observed {:?}",
                            observation.untouched_member_count
                        ),
                    )?;
                }
                state.source_output = Some(observation.output_sha256.clone());
                state.source_semantic = observation.semantic_sha256.clone();
                state.source_untouched_count = observation.untouched_member_count;
                state.source_untouched_hash = observation.untouched_member_sha256.clone();
                state.source_counters = observation.source_counters.clone();
            },
            "B2" => {
                if state.source_output.as_deref() != Some(observation.output_sha256.as_str()) {
                    self.failure(
                        sample,
                        warmup,
                        cell,
                        leg,
                        Implementation::Source,
                        "B1 and B2 source output hashes differ".to_owned(),
                    )?;
                }
                if state.source_semantic.as_deref() != observation.semantic_sha256.as_deref() {
                    self.failure(
                        sample,
                        warmup,
                        cell,
                        leg,
                        Implementation::Source,
                        "B1 and B2 source semantic hashes differ".to_owned(),
                    )?;
                }
                if state.source_untouched_count != observation.untouched_member_count
                    || state.source_untouched_hash.as_ref()
                        != observation.untouched_member_sha256.as_ref()
                {
                    self.failure(
                        sample,
                        warmup,
                        cell,
                        leg,
                        Implementation::Source,
                        "B1 and B2 untouched-member evidence differs".to_owned(),
                    )?;
                }
                if state.source_counters.as_ref() != observation.source_counters.as_ref() {
                    self.failure(
                        sample,
                        warmup,
                        cell,
                        leg,
                        Implementation::Source,
                        "B1 and B2 source logical counters/cache fields are not neutral".to_owned(),
                    )?;
                }
            },
            "A2" => {
                if state.eager_output.as_deref() != Some(observation.output_sha256.as_str()) {
                    self.failure(
                        sample,
                        warmup,
                        cell,
                        leg,
                        Implementation::Eager,
                        "A1 and A2 eager output hashes differ".to_owned(),
                    )?;
                }
            },
            _ => unreachable!("fixed ABBA leg"),
        }
        Ok(())
    }

    fn failure(
        &mut self,
        sample: Option<usize>,
        warmup: bool,
        cell: CellSpec,
        leg: &'static str,
        implementation: Implementation,
        error: String,
    ) -> AnyResult<()> {
        record_failure(
            &mut self.failure_writer,
            &mut self.failure_count,
            FailureRecord {
                sample,
                warmup,
                cell: cell.name().to_owned(),
                leg: leg.to_owned(),
                implementation: implementation.name().to_owned(),
                error,
            },
        )
    }
}

fn invoke_child(benchmark_bin: &Path, case: &str, shape: &str) -> AnyResult<Value> {
    let temporary = TemporaryJson::new()?;
    let output = Command::new(benchmark_bin)
        .arg("--warmup")
        .arg("0")
        .arg("--samples")
        .arg("1")
        .arg("--case")
        .arg(case)
        .arg("--xlsx-cell-crud-shape")
        .arg(shape)
        .arg("--json")
        .arg(&temporary.path)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .output()?;
    if !output.status.success() {
        return Err(format!("child {case} {shape} exited with {}", output.status).into());
    }
    let length = fs::metadata(&temporary.path)?.len();
    if length == 0 || length > MAX_CHILD_JSON_BYTES {
        return Err(
            format!("child JSON length {length} is outside 1..={MAX_CHILD_JSON_BYTES}").into(),
        );
    }
    let mut file = File::open(&temporary.path)?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(usize::try_from(length)?)?;
    file.read_to_end(&mut bytes)?;
    Ok(serde_json::from_slice(&bytes)?)
}

struct TemporaryJson {
    path: PathBuf,
}

impl TemporaryJson {
    fn new() -> AnyResult<Self> {
        let directory = env::temp_dir();
        for _ in 0..1000 {
            let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
            let path = directory.join(format!(
                "litchi-xlsx-cell-values-abba-{}-{counter}.json",
                std::process::id()
            ));
            match OpenOptions::new().write(true).create_new(true).open(&path) {
                Ok(file) => {
                    drop(file);
                    fs::remove_file(&path)?;
                    return Ok(Self { path });
                },
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => return Err(error.into()),
            }
        }
        Err("could not allocate a unique temporary child JSON path".into())
    }
}

impl Drop for TemporaryJson {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}

fn parse_child_report(
    report: &Value,
    cell: CellSpec,
    implementation: Implementation,
    benchmark_identity: &BinaryIdentity,
    runner_identity: &RunnerIdentity,
    blockers: &mut Vec<String>,
) -> Result<ChildObservation, String> {
    let report_object = report
        .as_object()
        .ok_or_else(|| "child report root is not an object".to_owned())?;
    let schema_version = required_u64(report_object, "schema_version")?;
    if schema_version != 1 {
        return Err(format!("unsupported child schema_version {schema_version}"));
    }
    let tool = required_object(report, "tool")?;
    if required_string(tool, "name")? != BENCHMARK_NAME {
        return Err("child tool name is not litchi-perf-baseline".to_owned());
    }
    let binary_value = required_object(report, "binary_identity")?;
    let child_binary = parse_binary_identity(binary_value)?;
    if child_binary.sha256 != benchmark_identity.sha256
        || child_binary.bytes != benchmark_identity.bytes
    {
        return Err("child binary identity does not match requested benchmark binary".to_owned());
    }
    let environment = required_object(report, "environment")?;
    let git_revision =
        optional_string(environment, "git_revision", blockers, "child git_revision")?;
    let git_dirty = optional_bool(
        environment,
        "git_worktree_dirty",
        blockers,
        "child git_dirty",
    )?;
    if let (Some(expected), Some(observed)) = (&runner_identity.git_revision, &git_revision) {
        if expected != observed {
            return Err("child git revision differs from runner revision".to_owned());
        }
    } else {
        add_blocker(blockers, "child or runner git revision unavailable");
    }
    if let Some(dirty) = git_dirty {
        if dirty {
            add_blocker(
                blockers,
                "child git worktree is dirty; evidence claim blocked",
            );
        }
    } else {
        add_blocker(blockers, "child git dirty state unavailable");
    }

    let configuration = required_object(report, "configuration")?;
    if required_u64(configuration, "samples_per_case")? != 1
        || required_u64(configuration, "warmup_iterations_per_case")? != 0
    {
        return Err("child configuration is not one exact warmup-zero sample".to_owned());
    }
    let cases = required_string_array(configuration, "cases")?;
    let expected_case = case_name(cell, implementation);
    if cases.len() != 1 || cases.first() != Some(&expected_case) {
        return Err("child configuration contains more than the exact requested case".to_owned());
    }
    let shapes = required_string_array(configuration, "xlsx_cell_crud_shapes")?;
    if shapes.len() != 1 || shapes.first().map(String::as_str) != Some(cell.shape) {
        return Err("child configuration contains more than the exact requested shape".to_owned());
    }

    let corpus = parse_corpus(report, cell)?;
    let (expected_archive_bytes, expected_archive_sha256) =
        pinned_cell_corpus_identity(cell.shape)?;
    if corpus.archive_member_count != 17
        || usize::try_from(corpus.archive_bytes).ok() != Some(expected_archive_bytes)
        || corpus.archive_sha256 != expected_archive_sha256
    {
        return Err(format!(
            "xlsx cell-values ABBA v1 requires the pinned 17-member four-sheet no-calcChain {} corpus",
            cell.shape
        ));
    }
    let results = required_array(report, "results")?;
    if results.len() != 1 {
        return Err("child report did not contain exactly one result".to_owned());
    }
    let result = results
        .first()
        .ok_or_else(|| "child result is missing".to_owned())?
        .as_object()
        .ok_or_else(|| "child result is not an object".to_owned())?;
    if required_string(result, "case")? != expected_case {
        return Err("child result case differs from requested case".to_owned());
    }
    let elapsed = required_map_object(result, "elapsed_ns")?;
    if required_string(elapsed, "unit")? != "ns" {
        return Err("child elapsed_ns unit is not ns".to_owned());
    }
    let elapsed_samples = required_map_array(elapsed, "samples")?;
    if elapsed_samples.len() != 1 {
        return Err("child elapsed_ns did not contain exactly one total sample".to_owned());
    }
    let primary_ns = value_as_u64(
        elapsed_samples
            .first()
            .ok_or_else(|| "child elapsed sample is missing".to_owned())?,
    )?;
    let output_sha256 = required_string(result, "output_sha256")?.to_owned();
    validate_sha256(&output_sha256, "child output_sha256")?;
    let sink = parse_sink(result)?;
    let mut semantic_sha256 = None;
    let mut untouched_member_count = None;
    let mut untouched_member_sha256 = None;
    let mut source_counters = None;
    if implementation == Implementation::Source {
        let source = required_map_object(result, "source")?;
        let xlsx = required_map_object(source, "xlsx_cell_values")?;
        let evidence = parse_source_evidence(source, xlsx, cell, &output_sha256, blockers)?;
        if evidence.update_count != Some(expected_update_count(cell)) {
            if evidence.update_count.is_some() {
                return Err("source update_count differs from exact workload".to_owned());
            }
            add_blocker(blockers, "source update_count unavailable");
        }
        if evidence.selected_worksheet_count != Some(expected_touched_worksheets(cell.workload)) {
            if evidence.selected_worksheet_count.is_some() {
                return Err(
                    "source selected_worksheet_count differs from exact workload".to_owned(),
                );
            }
            add_blocker(blockers, "source selected_worksheet_count unavailable");
        }
        semantic_sha256 = evidence.semantic_sha256;
        untouched_member_count = evidence.untouched_member_count;
        untouched_member_sha256 = evidence.untouched_member_sha256;
        source_counters = Some(evidence.counters);
    }
    Ok(ChildObservation {
        case: expected_case,
        shape: cell.shape.to_owned(),
        workload: cell.workload.name().to_owned(),
        implementation: implementation.name().to_owned(),
        primary_ns,
        output_sha256,
        semantic_sha256,
        untouched_member_count,
        untouched_member_sha256,
        sink,
        source_counters,
        corpus,
        binary_identity: child_binary,
        git_revision,
        git_dirty,
    })
}

struct SourceEvidence {
    update_count: Option<usize>,
    selected_worksheet_count: Option<usize>,
    semantic_sha256: Option<String>,
    untouched_member_count: Option<usize>,
    untouched_member_sha256: Option<String>,
    counters: SourceCounters,
}

fn parse_source_evidence(
    source: &Map<String, Value>,
    xlsx: &Map<String, Value>,
    _cell: CellSpec,
    output_sha256: &str,
    blockers: &mut Vec<String>,
) -> Result<SourceEvidence, String> {
    if required_string(xlsx, "implementation")? != "source-backed"
        || required_string(xlsx, "cache_mode")? != "unmanaged-control"
    {
        return Err("source child is not the unmanaged source-backed control".to_owned());
    }
    let update_count = optional_usize(xlsx, "update_count", blockers, "source update_count")?;
    let selected_worksheet_count = optional_usize(
        xlsx,
        "selected_worksheet_count",
        blockers,
        "source selected_worksheet_count",
    )?;
    let mut values = BTreeMap::new();
    let mut complete = true;
    for key in [
        "source_read_calls",
        "source_read_bytes",
        "workbook_read_calls",
        "workbook_read_bytes",
        "selected_worksheet_read_calls",
        "selected_worksheet_read_bytes",
        "unselected_worksheet_read_calls",
        "unselected_worksheet_read_bytes",
        "payload_materializations",
        "cache_hits",
        "cache_cold_loads",
        "cache_waiter_joins",
        "cache_successful_loads",
        "cache_failed_loads",
        "cache_evictions",
        "cache_bypasses",
        "cache_oversized_bypasses",
        "cache_allocation_bypasses",
        "cache_in_flight_loads",
        "cache_retained_entries",
        "cache_retained_bytes",
        "cache_budget_memory_used",
        "cache_budget_reserved_bytes",
        "cache_budget_reservation_failures",
        "budget_used_after_package_drop",
        "budget_used_after_handles_drop",
        "budget_objects_used_after_handles_drop",
    ] {
        match optional_vector_u64(xlsx, key, blockers, &format!("source {key}"))? {
            Some(value) => {
                values.insert(key.to_owned(), json!(value));
            },
            None => complete = false,
        }
    }
    match optional_bool(
        xlsx,
        "cache_budget_managed",
        blockers,
        "source cache_budget_managed",
    )? {
        Some(value) => {
            if value {
                return Err("unmanaged source control reported managed cache budget".to_owned());
            }
            values.insert("cache_budget_managed".to_owned(), json!(value));
        },
        None => complete = false,
    }
    match optional_nullable_u64(
        xlsx,
        "cache_budget_memory_limit",
        blockers,
        "source cache_budget_memory_limit",
    )? {
        Some(value) => {
            values.insert("cache_budget_memory_limit".to_owned(), value);
        },
        None => complete = false,
    }
    for key in ["pre_publication_budget", "post_publication_budget"] {
        match optional_budget_snapshot(xlsx, key, blockers)? {
            Some(snapshot) => {
                for (name, value) in snapshot {
                    values.insert(format!("{key}.{name}"), value);
                }
            },
            None => complete = false,
        }
    }
    let top_read_calls = optional_vector_u64(source, "read_calls", blockers, "source read_calls")?;
    let top_read_bytes = optional_vector_u64(source, "read_bytes", blockers, "source read_bytes")?;
    if top_read_calls.is_none() || top_read_bytes.is_none() {
        complete = false;
    }
    if let Some(read_calls) = top_read_calls {
        values.insert("top_read_calls".to_owned(), json!(read_calls));
    }
    if let Some(read_bytes) = top_read_bytes {
        values.insert("top_read_bytes".to_owned(), json!(read_bytes));
    }
    let output_vector =
        optional_string_vector(xlsx, "output_sha256", blockers, "source output_sha256")?;
    if output_vector
        .as_ref()
        .and_then(|values| values.first())
        .map(String::as_str)
        != Some(output_sha256)
    {
        if output_vector.is_some() {
            return Err("source output_sha256 differs from top-level output_sha256".to_owned());
        }
        complete = false;
    }
    let semantic_sha256 =
        optional_string_vector(xlsx, "semantic_sha256", blockers, "source semantic_sha256")?
            .and_then(|values| values.into_iter().next());
    if semantic_sha256.is_none() {
        complete = false;
    } else if let Some(hash) = &semantic_sha256 {
        validate_sha256(hash, "source semantic_sha256")?;
    }
    let untouched_member_count = optional_usize(
        xlsx,
        "untouched_member_count",
        blockers,
        "source untouched_member_count",
    )?;
    let untouched_member_sha256 = optional_string_vector(
        xlsx,
        "untouched_member_sha256",
        blockers,
        "source untouched_member_sha256",
    )?
    .and_then(|values| values.into_iter().next());
    if untouched_member_count.is_none() || untouched_member_sha256.is_none() {
        complete = false;
    }
    if let Some(count) = untouched_member_count {
        if count == 0 {
            return Err("source untouched_member_count is zero".to_owned());
        }
    }
    if let Some(hash) = &untouched_member_sha256 {
        validate_sha256(hash, "source untouched_member_sha256")?;
    }
    if let Some(reads) = values.get("unselected_worksheet_read_calls") {
        if reads != &json!(0) {
            return Err("source unselected worksheet logical reads are nonzero".to_owned());
        }
    }
    if let Some(bytes) = values.get("unselected_worksheet_read_bytes") {
        if bytes != &json!(0) {
            return Err("source unselected worksheet logical bytes are nonzero".to_owned());
        }
    }
    Ok(SourceEvidence {
        update_count,
        selected_worksheet_count,
        semantic_sha256,
        untouched_member_count,
        untouched_member_sha256,
        counters: SourceCounters { values, complete },
    })
}

fn parse_corpus(report: &Value, cell: CellSpec) -> Result<CorpusIdentity, String> {
    let corpus = required_array(report, "results")?
        .first()
        .ok_or_else(|| "child result is missing for corpus parsing".to_owned())
        .and_then(|result| required_object(result, "corpus"))?;
    let shape = required_string(corpus, "shape")?;
    if shape != cell.shape {
        return Err("child corpus shape differs from requested shape".to_owned());
    }
    let generator = required_string(corpus, "generator")?;
    if generator != CORPUS_GENERATOR {
        return Err(
            "child corpus generator is not the pinned XLSX cell-values generator".to_owned(),
        );
    }
    let expected_name = format!("xlsx-cell-values-{}", cell.shape);
    if required_string(corpus, "name")? != expected_name.as_str() {
        return Err("child corpus name differs from the pinned shape".to_owned());
    }
    let package_format = required_string(corpus, "package_format")?;
    let payload_kind = required_string(corpus, "payload_kind")?;
    let compression = required_string(corpus, "compression")?;
    let entry_count = required_usize(corpus, "entry_count")?;
    let archive_member_count = required_usize(corpus, "archive_member_count")?;
    let entry_bytes = required_usize(corpus, "entry_bytes")?;
    let uncompressed_payload_bytes = required_usize(corpus, "uncompressed_payload_bytes")?;
    let archive_bytes = required_usize(corpus, "archive_bytes")?;
    let archive_sha256 = required_string(corpus, "archive_sha256")?.to_owned();
    validate_sha256(&archive_sha256, "corpus archive_sha256")?;
    let target_entry = required_string(corpus, "target_entry")?.to_owned();
    let target_payload_bytes = required_usize(corpus, "target_payload_bytes")?;
    let target_payload_sha256 = required_string(corpus, "target_payload_sha256")?.to_owned();
    validate_sha256(&target_payload_sha256, "corpus target_payload_sha256")?;
    let expected_cells = if cell.shape == "medium" {
        4 * 48 * 48
    } else {
        17_792
    };
    let expected_uncompressed = expected_cells * 4 + 8 * 512 * 1024;
    if entry_count != expected_cells
        || entry_bytes != 4
        || package_format != "XLSX/OPC/ZIP"
        || payload_kind != "deterministic-multi-sheet-scalar-grid-with-media"
        || compression != "deflate"
        || uncompressed_payload_bytes != expected_uncompressed
        || target_entry != "Sheet1!A1"
        || target_payload_bytes != 1
        || target_payload_sha256 != sha256_hex(b"0")
        || archive_member_count < 8
        || archive_bytes == 0
    {
        return Err("child corpus manifest fields differ from the deterministic corpus".to_owned());
    }
    let xlsx = required_map_object(corpus, "xlsx")?;
    if required_usize(xlsx, "sheet_count")? != 4
        || required_usize(xlsx, "rows_per_sheet")? != if cell.shape == "medium" { 48 } else { 128 }
        || required_usize(xlsx, "columns_per_sheet")?
            != if cell.shape == "medium" { 48 } else { 128 }
    {
        return Err("child XLSX corpus dimensions differ from the deterministic shape".to_owned());
    }
    let source_members = required_map_object(xlsx, "source_members")?;
    let workbook_member = required_string(source_members, "workbook")?.to_owned();
    if workbook_member != "xl/workbook.xml" {
        return Err("child workbook source member differs from the fixed corpus".to_owned());
    }
    let worksheet_members = required_string_array(source_members, "worksheets")?;
    let expected_worksheets = (1..=4)
        .map(|index| format!("xl/worksheets/sheet{index}.xml"))
        .collect::<Vec<_>>();
    if worksheet_members != expected_worksheets {
        return Err("child worksheet source members differ from the fixed corpus".to_owned());
    }
    let shared_strings_member = optional_nullable_string(source_members, "shared_strings")?;
    let styles_member = optional_nullable_string(source_members, "styles")?;
    Ok(CorpusIdentity {
        name: required_string(corpus, "name")?.to_owned(),
        generator: generator.to_owned(),
        package_format: package_format.to_owned(),
        shape: shape.to_owned(),
        payload_kind: payload_kind.to_owned(),
        compression: compression.to_owned(),
        entry_count,
        archive_member_count,
        entry_bytes,
        uncompressed_payload_bytes,
        archive_bytes,
        archive_sha256,
        target_entry,
        target_payload_bytes,
        target_payload_sha256,
        workbook_member,
        worksheet_members,
        shared_strings_member,
        styles_member,
    })
}

fn parse_sink(result: &Map<String, Value>) -> Result<SinkEvidence, String> {
    let sink = required_map_object(result, "sink")?;
    let accepted_bytes = required_u64(sink, "accepted_bytes")?;
    let write_calls = required_u64(sink, "write_calls")?;
    let largest_write = required_u64(sink, "largest_write")?;
    if accepted_bytes == 0 || write_calls == 0 || largest_write > MAX_WRITE_BYTES {
        return Err("child sink evidence is empty or exceeds the 64 KiB write bound".to_owned());
    }
    let retained_output_bytes = match sink.get("retained_output_bytes") {
        None | Some(Value::Null) => None,
        Some(value) => Some(value_as_u64(value)?),
    };
    Ok(SinkEvidence {
        accepted_bytes,
        write_calls,
        largest_write,
        retained_output_bytes,
    })
}

fn summarize_cell(
    cell: CellSpec,
    accumulator: &CellAccumulator,
    requested_samples: usize,
    enforced: bool,
) -> CellSummary {
    let legs = LEGS
        .into_iter()
        .enumerate()
        .map(|(index, (leg, implementation))| LegSummary {
            leg,
            implementation: implementation.name(),
            statistics: statistics(&accumulator.legs[index].values),
        })
        .collect::<Vec<_>>();
    let eager_a1 = statistics(&accumulator.legs[0].values);
    let source_b1 = statistics(&accumulator.legs[1].values);
    let source_b2 = statistics(&accumulator.legs[2].values);
    let eager_a2 = statistics(&accumulator.legs[3].values);
    let mut same_side = Vec::new();
    same_side.extend(symmetric_gates(
        "eager A1/A2",
        eager_a1.as_ref(),
        eager_a2.as_ref(),
        enforced,
    ));
    same_side.extend(symmetric_gates(
        "source B1/B2",
        source_b1.as_ref(),
        source_b2.as_ref(),
        enforced,
    ));
    let mut directional = Vec::new();
    directional.extend(directional_gates(
        "A1 eager -> B1 source",
        eager_a1.as_ref(),
        source_b1.as_ref(),
        enforced,
    ));
    directional.extend(directional_gates(
        "A2 eager -> B2 source",
        eager_a2.as_ref(),
        source_b2.as_ref(),
        enforced,
    ));
    let all_required_gates_passed = same_side
        .iter()
        .chain(directional.iter())
        .all(|gate| gate.passed);
    let successful_samples = accumulator.counts.iter().copied().min().unwrap_or(0);
    CellSummary {
        cell: cell.name(),
        shape: cell.shape,
        workload: cell.workload.name(),
        requested_samples,
        successful_samples,
        legs,
        same_side_gates: same_side,
        directional_gates: directional,
        all_required_gates_passed,
    }
}

fn statistics(values: &[u64]) -> Option<StatisticsReport> {
    if values.is_empty() {
        return None;
    }
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    let sum = values.iter().map(|value| u128::from(*value)).sum::<u128>();
    Some(StatisticsReport {
        count: values.len(),
        min_ns: sorted[0],
        p50_ns: floor_midpoint(sorted[(sorted.len() - 1) / 2], sorted[sorted.len() / 2]),
        mean_ns_floor: u64::try_from(sum / values.len() as u128).unwrap_or(u64::MAX),
        p95_ns: nearest_rank(&sorted, 95),
        p99_ns: nearest_rank(&sorted, 99),
        max_ns: *sorted.last().unwrap_or(&0),
    })
}

fn nearest_rank(values: &[u64], percentile: usize) -> u64 {
    let rank = (values.len() * percentile).div_ceil(100).max(1);
    values[rank.saturating_sub(1).min(values.len() - 1)]
}

fn floor_midpoint(left: u64, right: u64) -> u64 {
    left / 2 + right / 2 + (left % 2 + right % 2) / 2
}

fn symmetric_gates(
    prefix: &str,
    left: Option<&StatisticsReport>,
    right: Option<&StatisticsReport>,
    enforced: bool,
) -> Vec<GateResult> {
    ["p50", "mean", "p95", "p99"]
        .into_iter()
        .map(|name| {
            let (observed_percent, absolute_delta_ns) = match (left, right) {
                (Some(left), Some(right)) => {
                    let left_value = metric(left, name);
                    let right_value = metric(right, name);
                    (
                        symmetric_percent(left_value, right_value),
                        Some(left_value.abs_diff(right_value)),
                    )
                },
                _ => (None, None),
            };
            GateResult {
                name: format!("{prefix} {name} symmetric aggregate delta"),
                threshold_percent: SAME_SIDE_GATE_PERCENT,
                observed_percent,
                absolute_delta_ns,
                passed: observed_percent.is_some_and(|value| value <= SAME_SIDE_GATE_PERCENT),
                enforced,
                rationale: "symmetric delta of per-leg aggregate p50/mean/p95/p99 must be <=5%",
            }
        })
        .collect()
}

fn directional_gates(
    prefix: &str,
    eager: Option<&StatisticsReport>,
    source: Option<&StatisticsReport>,
    enforced: bool,
) -> Vec<GateResult> {
    [
        ("p50", true, "source must improve by >=1% or >=50 us"),
        ("mean", true, "source must improve by >=1% or >=50 us"),
        ("p95", false, "source may not be more than 5% adverse"),
        ("p99", false, "source may not be more than 5% adverse"),
    ]
    .into_iter()
    .map(|(name, improvement_required, rationale)| {
        let (observed_percent, absolute_delta_ns, passed) = match (eager, source) {
            (Some(eager), Some(source)) => {
                let eager_value = metric(eager, name);
                let source_value = metric(source, name);
                let change = relative_change(eager_value, source_value);
                let delta = eager_value.abs_diff(source_value);
                let passed = change.is_some_and(|change| {
                    if improvement_required {
                        source_value < eager_value
                            && ((-change) >= DIRECTIONAL_IMPROVEMENT_PERCENT
                                || delta >= DIRECTIONAL_IMPROVEMENT_NS)
                    } else {
                        change <= DIRECTIONAL_ADVERSE_PERCENT
                    }
                });
                (change, Some(delta), passed)
            },
            _ => (None, None, false),
        };
        GateResult {
            name: format!("{prefix} {name}"),
            threshold_percent: if improvement_required {
                DIRECTIONAL_IMPROVEMENT_PERCENT
            } else {
                DIRECTIONAL_ADVERSE_PERCENT
            },
            observed_percent,
            absolute_delta_ns,
            passed,
            enforced,
            rationale,
        }
    })
    .collect()
}

fn metric(statistics: &StatisticsReport, name: &str) -> u64 {
    match name {
        "p50" => statistics.p50_ns,
        "mean" => statistics.mean_ns_floor,
        "p95" => statistics.p95_ns,
        "p99" => statistics.p99_ns,
        _ => 0,
    }
}

fn symmetric_percent(left: u64, right: u64) -> Option<f64> {
    if left == 0 && right == 0 {
        return Some(0.0);
    }
    let denominator = left.min(right);
    (denominator != 0).then(|| left.abs_diff(right) as f64 / denominator as f64 * 100.0)
}

fn relative_change(eager: u64, source: u64) -> Option<f64> {
    if eager == 0 && source == 0 {
        return Some(0.0);
    }
    (eager != 0).then(|| (source as f64 / eager as f64 - 1.0) * 100.0)
}

fn expected_update_count(cell: CellSpec) -> usize {
    match cell.workload {
        Workload::OneEdit => 1,
        Workload::OnePercent => {
            let total: usize = if cell.shape == "medium" {
                4 * 48 * 48
            } else {
                17_792
            };
            total.div_ceil(100)
        },
        Workload::Batch => 256,
    }
}

const fn expected_touched_worksheets(workload: Workload) -> usize {
    match workload {
        Workload::OneEdit => 1,
        Workload::OnePercent | Workload::Batch => 4,
    }
}

fn pinned_cell_corpus_identity(shape: &str) -> Result<(usize, &'static str), String> {
    match shape {
        "medium" => Ok((
            4_226_429,
            "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
        )),
        "dense-sparse" => Ok((
            4_251_863,
            "893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a",
        )),
        _ => Err(format!(
            "xlsx cell-values ABBA v1 does not claim corpus shape '{shape}'"
        )),
    }
}

fn expected_cell_untouched_member_count(
    archive_member_count: usize,
    workload: Workload,
) -> Option<usize> {
    archive_member_count
        .checked_sub(expected_touched_worksheets(workload))
        .and_then(|count| count.checked_sub(1))
}

#[cfg(test)]
mod xlsx_cell_values_abba_identity_tests {
    use super::*;

    #[test]
    fn pinned_current_shape_identities_are_exact() {
        assert_eq!(
            pinned_cell_corpus_identity("medium").unwrap(),
            (
                4_226_429,
                "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036"
            )
        );
        assert_eq!(
            pinned_cell_corpus_identity("dense-sparse").unwrap(),
            (
                4_251_863,
                "893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a"
            )
        );
        assert!(pinned_cell_corpus_identity("unknown").is_err());
    }

    #[test]
    fn pinned_current_shape_untouched_counts_are_exact() {
        assert_eq!(
            expected_cell_untouched_member_count(17, Workload::OneEdit),
            Some(15)
        );
        assert_eq!(
            expected_cell_untouched_member_count(17, Workload::OnePercent),
            Some(12)
        );
        assert_eq!(
            expected_cell_untouched_member_count(17, Workload::Batch),
            Some(12)
        );
    }
}

fn case_name(cell: CellSpec, implementation: Implementation) -> String {
    format!(
        "{}{}",
        implementation.case_prefix(),
        cell.workload.case_suffix()
    )
}

fn implementation_from_name(name: &str) -> Implementation {
    if name == Implementation::Source.name() {
        Implementation::Source
    } else {
        Implementation::Eager
    }
}

fn resolve_benchmark_binary(requested: Option<PathBuf>) -> AnyResult<PathBuf> {
    if let Some(path) = requested {
        if path.is_file() {
            return Ok(path.canonicalize()?);
        }
        return Err(format!("benchmark binary does not exist: {}", path.display()).into());
    }
    let current = env::current_exe()?;
    let mut candidates = Vec::new();
    if let Some(parent) = current.parent() {
        candidates.push(parent.join(BENCHMARK_NAME));
        if let Some(parent) = parent.parent() {
            candidates.push(parent.join(BENCHMARK_NAME));
        }
    }
    candidates.push(PathBuf::from(format!(
        "tools/perf-baseline/target/release/{BENCHMARK_NAME}"
    )));
    candidates.push(PathBuf::from(format!(
        "tools/perf-baseline/target/debug/{BENCHMARK_NAME}"
    )));
    candidates
        .into_iter()
        .find(|path| path.is_file())
        .map(|path| path.canonicalize())
        .transpose()?
        .ok_or_else(|| "could not find sibling litchi-perf-baseline; pass --benchmark-bin".into())
}

fn binary_identity(path: &Path, profile: Option<&str>) -> AnyResult<BinaryIdentity> {
    let bytes = fs::read(path)?;
    Ok(BinaryIdentity {
        path: path.canonicalize()?.display().to_string(),
        sha256: sha256_hex(&bytes),
        bytes: u64::try_from(bytes.len())?,
        profile: profile.map(str::to_owned),
    })
}

fn inferred_profile(path: &Path) -> &'static str {
    if path
        .components()
        .any(|component| component.as_os_str() == "release")
    {
        "release"
    } else if path
        .components()
        .any(|component| component.as_os_str() == "debug")
    {
        "debug"
    } else {
        "unknown"
    }
}

fn load_expected_corpora(path: Option<&Path>) -> AnyResult<BTreeMap<String, CorpusIdentity>> {
    let Some(path) = path else {
        return Ok(BTreeMap::new());
    };
    let metadata = fs::metadata(path)?;
    if metadata.len() == 0 || metadata.len() > 1024 * 1024 {
        return Err("corpus identity manifest must be in 1..=1048576 bytes".into());
    }
    let root: Value = serde_json::from_slice(&fs::read(path)?)?;
    let corpora = root.get("corpora").unwrap_or(&root).clone();
    Ok(serde_json::from_value(corpora)?)
}

fn runner_identity() -> AnyResult<RunnerIdentity> {
    let executable = env::current_exe()?;
    Ok(RunnerIdentity {
        executable_sha256: sha256_hex(&fs::read(executable)?),
        git_revision: command_text("git", &["rev-parse", "HEAD"]),
        git_dirty: command_text("git", &["status", "--porcelain"]).map(|output| !output.is_empty()),
        rustc_vv: command_text("rustc", &["-Vv"]),
        os: env::consts::OS.to_owned(),
        arch: env::consts::ARCH.to_owned(),
        cpu: env::var("PROCESSOR_IDENTIFIER")
            .ok()
            .or_else(|| env::var("HOSTTYPE").ok()),
        memory: fs::read_to_string("/proc/meminfo").ok().and_then(|value| {
            value
                .lines()
                .find(|line| line.starts_with("MemTotal:"))
                .map(str::to_owned)
        }),
        rustflags: env::var("RUSTFLAGS").ok(),
    })
}

fn command_text(command: &str, arguments: &[&str]) -> Option<String> {
    let output = Command::new(command).args(arguments).output().ok()?;
    output
        .status
        .success()
        .then(|| String::from_utf8_lossy(&output.stdout).trim().to_owned())
}

fn exclusions() -> Vec<&'static str> {
    vec![
        "phase subvectors and phase-attribution comparisons",
        "cross-implementation ZIP byte identity; eager/source children validate their own semantic result and only same-side hashes are compared",
        "cross-revision behavior",
        "physical I/O, allocation, RSS, CPU utilization, and cache warmth",
        "managed-budget behavior",
        "formula, date, structural, and unsupported XLSX edits",
        "general XLSX CRUD or real-producer interoperability",
    ]
}

fn write_json<T: Serialize>(path: &Path, value: &T) -> AnyResult<()> {
    fs::write(path, serde_json::to_vec_pretty(value)?)?;
    Ok(())
}

fn write_json_line<T: Serialize>(writer: &mut BufWriter<File>, value: &T) -> AnyResult<()> {
    let line = serde_json::to_vec(value)?;
    if line.len() > MAX_ROW_BYTES {
        return Err("serialized sample row exceeds bounded row size".into());
    }
    writer.write_all(&line)?;
    writer.write_all(b"\n")?;
    Ok(())
}

fn record_failure(
    writer: &mut BufWriter<File>,
    count: &mut usize,
    failure: FailureRecord,
) -> AnyResult<()> {
    if *count >= MAX_FAILURE_ROWS {
        return Err("failure row bound exceeded".into());
    }
    write_json_line(writer, &failure)?;
    *count += 1;
    Ok(())
}

fn write_artifact_manifest(out_dir: &Path) -> AnyResult<()> {
    let names = [
        "protocol.json",
        "samples.jsonl",
        "summary.json",
        "failures.jsonl",
        "sha256.txt",
        "process-time.txt",
    ];
    let mut files = Vec::new();
    files.try_reserve_exact(names.len())?;
    for name in names {
        let bytes = fs::read(out_dir.join(name))?;
        files.push(ArtifactEntry {
            path: name.to_owned(),
            bytes: u64::try_from(bytes.len())?,
            sha256: sha256_hex(&bytes),
        });
    }
    write_json(
        &out_dir.join("artifact-manifest.json"),
        &ArtifactManifest {
            schema: SCHEMA,
            self_excluded: true,
            files,
        },
    )
}

fn process_time_text(started: SystemTime) -> AnyResult<String> {
    let start_ms = started.duration_since(UNIX_EPOCH)?.as_millis();
    let finish_ms = SystemTime::now().duration_since(UNIX_EPOCH)?.as_millis();
    Ok(format!(
        "pid={}\nstart_unix_ms={start_ms}\nfinish_unix_ms={finish_ms}\n",
        std::process::id()
    ))
}

fn add_blocker(blockers: &mut Vec<String>, blocker: &str) {
    if !blockers.iter().any(|value| value == blocker) {
        blockers.push(blocker.to_owned());
    }
}

fn required_object<'a>(value: &'a Value, name: &str) -> Result<&'a Map<String, Value>, String> {
    value
        .get(name)
        .ok_or_else(|| format!("missing JSON object field {name}"))?
        .as_object()
        .ok_or_else(|| format!("JSON field {name} is not an object"))
}

fn required_map_object<'a>(
    value: &'a Map<String, Value>,
    name: &str,
) -> Result<&'a Map<String, Value>, String> {
    value
        .get(name)
        .ok_or_else(|| format!("missing JSON object field {name}"))?
        .as_object()
        .ok_or_else(|| format!("JSON field {name} is not an object"))
}

fn required_map_array<'a>(
    value: &'a Map<String, Value>,
    name: &str,
) -> Result<&'a Vec<Value>, String> {
    value
        .get(name)
        .ok_or_else(|| format!("missing JSON array field {name}"))?
        .as_array()
        .ok_or_else(|| format!("JSON field {name} is not an array"))
}

fn required_array<'a>(value: &'a Value, name: &str) -> Result<&'a Vec<Value>, String> {
    value
        .get(name)
        .ok_or_else(|| format!("missing JSON array field {name}"))?
        .as_array()
        .ok_or_else(|| format!("JSON field {name} is not an array"))
}

fn required_string<'a>(value: &'a Map<String, Value>, name: &str) -> Result<&'a str, String> {
    value
        .get(name)
        .ok_or_else(|| format!("missing JSON string field {name}"))?
        .as_str()
        .ok_or_else(|| format!("JSON field {name} is not a string"))
}

fn required_u64(value: &Map<String, Value>, name: &str) -> Result<u64, String> {
    value_as_u64(
        value
            .get(name)
            .ok_or_else(|| format!("missing JSON integer field {name}"))?,
    )
}

fn required_usize(value: &Map<String, Value>, name: &str) -> Result<usize, String> {
    usize::try_from(required_u64(value, name)?)
        .map_err(|_| format!("JSON field {name} overflows usize"))
}

fn value_as_u64(value: &Value) -> Result<u64, String> {
    value
        .as_u64()
        .ok_or_else(|| "JSON value is not a non-negative integer".to_owned())
}

fn required_string_array(value: &Map<String, Value>, name: &str) -> Result<Vec<String>, String> {
    value
        .get(name)
        .ok_or_else(|| format!("missing JSON array field {name}"))?
        .as_array()
        .ok_or_else(|| format!("JSON field {name} is not an array"))?
        .iter()
        .map(|value| {
            value
                .as_str()
                .map(str::to_owned)
                .ok_or_else(|| format!("JSON array {name} contains a non-string"))
        })
        .collect()
}

fn optional_string(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
    label: &str,
) -> Result<Option<String>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, label);
        return Ok(None);
    };
    if field.is_null() {
        return Ok(None);
    }
    field
        .as_str()
        .map(|value| Some(value.to_owned()))
        .ok_or_else(|| format!("JSON field {name} is not a string or null"))
}

fn optional_bool(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
    label: &str,
) -> Result<Option<bool>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, label);
        return Ok(None);
    };
    if field.is_null() {
        return Ok(None);
    }
    field
        .as_bool()
        .map(Some)
        .ok_or_else(|| format!("JSON field {name} is not a boolean or null"))
}

fn optional_usize(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
    label: &str,
) -> Result<Option<usize>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, label);
        return Ok(None);
    };
    if field.is_null() {
        return Ok(None);
    }
    usize::try_from(value_as_u64(field)?)
        .map(Some)
        .map_err(|_| format!("JSON field {name} overflows usize"))
}

fn optional_vector_u64(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
    label: &str,
) -> Result<Option<u64>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, label);
        return Ok(None);
    };
    let values = field
        .as_array()
        .ok_or_else(|| format!("JSON field {name} is not an array"))?;
    if values.len() != 1 {
        return Err(format!(
            "JSON field {name} does not contain exactly one child sample"
        ));
    }
    values.first().map(value_as_u64).transpose()
}

fn optional_string_vector(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
    label: &str,
) -> Result<Option<Vec<String>>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, label);
        return Ok(None);
    };
    let values = field
        .as_array()
        .ok_or_else(|| format!("JSON field {name} is not an array"))?;
    if values.len() != 1 {
        return Err(format!(
            "JSON field {name} does not contain exactly one child sample"
        ));
    }
    Ok(Some(
        values
            .iter()
            .map(|value| {
                value
                    .as_str()
                    .map(str::to_owned)
                    .ok_or_else(|| format!("JSON field {name} contains a non-string"))
            })
            .collect::<Result<Vec<_>, _>>()?,
    ))
}

fn optional_nullable_u64(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
    label: &str,
) -> Result<Option<Value>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, label);
        return Ok(None);
    };
    if field.is_null() || field.as_u64().is_some() {
        return Ok(Some(field.clone()));
    }
    Err(format!("JSON field {name} is not an integer or null"))
}

fn optional_nullable_string(
    value: &Map<String, Value>,
    name: &str,
) -> Result<Option<String>, String> {
    let field = value
        .get(name)
        .ok_or_else(|| format!("missing JSON member field {name}"))?;
    if field.is_null() {
        Ok(None)
    } else {
        field
            .as_str()
            .map(|value| Some(value.to_owned()))
            .ok_or_else(|| format!("JSON member field {name} is not a string or null"))
    }
}

fn optional_budget_snapshot(
    value: &Map<String, Value>,
    name: &str,
    blockers: &mut Vec<String>,
) -> Result<Option<BTreeMap<String, Value>>, String> {
    let Some(field) = value.get(name) else {
        add_blocker(blockers, name);
        return Ok(None);
    };
    let array = field
        .as_array()
        .ok_or_else(|| format!("JSON budget field {name} is not an array"))?;
    if array.len() != 1 {
        return Err(format!(
            "JSON budget field {name} does not contain one sample"
        ));
    }
    let object = array
        .first()
        .and_then(Value::as_object)
        .ok_or_else(|| format!("JSON budget field {name} sample is not an object"))?;
    let mut result = BTreeMap::new();
    for key in [
        "input_bytes_used",
        "input_bytes_limit",
        "output_bytes_used",
        "output_bytes_limit",
        "work_used",
        "work_limit",
        "objects_used",
        "objects_limit",
        "catalog_reserved_objects",
        "cache_reserved_objects",
    ] {
        let field = object
            .get(key)
            .ok_or_else(|| format!("JSON budget snapshot {name}.{key} is missing"))?;
        if !field.is_null() && field.as_u64().is_none() {
            return Err(format!("JSON budget snapshot {name}.{key} is invalid"));
        }
        result.insert(key.to_owned(), field.clone());
    }
    Ok(Some(result))
}

fn validate_sha256(value: &str, label: &str) -> Result<(), String> {
    if value.len() != 64 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(format!("{label} is not a 64-character hexadecimal digest"));
    }
    Ok(())
}

fn parse_binary_identity(value: &Map<String, Value>) -> Result<BinaryIdentity, String> {
    let sha256 = required_string(value, "binary_sha256")?.to_owned();
    validate_sha256(&sha256, "child binary_sha256")?;
    Ok(BinaryIdentity {
        path: required_string(value, "path")?.to_owned(),
        sha256,
        bytes: required_u64(value, "binary_bytes")?,
        profile: optional_nullable_string(value, "profile")?,
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push(HEX[usize::from(byte >> 4)] as char);
        output.push(HEX[usize::from(byte & 0x0f)] as char);
    }
    output
}
