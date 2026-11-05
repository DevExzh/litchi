//! Timed opt-in replay of one qualified XLSX managed-compression case.
//!
//! The qualification module owns the fixed corpus, output oracles, finite
//! context construction, and inverse proof. This child keeps those helpers
//! private to the qualification subtree while placing every correctness
//! operation outside the wall-clock intervals.

use super::{
    ContextLimits, Corpus, ResourceSnapshot, Workload, XlsxCellCrudShape, XlsxCoordinate,
    build_xlsx_cell_crud_corpus, cache_limits, expected_output_sha256, expected_source_sha256,
    managed_context, require_clean_resources, sha256_hex, verify_xlsx_cell_crud_output,
    xlsx_cell_values_output_semantic_hash, xlsx_cell_values_payload_budget, xlsx_output_ceiling,
    xlsx_spec, xlsx_update_sheet_selectors, xlsx_value,
};
use crate::xlsx_source_editors::SourceBackedEditor;
use litchi_core::{Budget, OwnedSource, ReadAt, Resource};
use litchi_opc::{OpcPackage, PackageWriter, SourceCacheLimits, SourceLimits};
use litchi_xlsx::OpenOptions;
use litchi_xlsx::cell_values::Patch as MultiPatch;
use serde::Serialize;
use std::{
    error::Error,
    ffi::OsString,
    fs,
    io::{self, Write},
    path::{Path, PathBuf},
    sync::Arc,
    time::Instant,
};

#[path = "xlsx_managed_compression_allocation.rs"]
pub(crate) mod allocation;

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx-managed-compression-benchmark.v1";
const MAX_WARMUP: usize = 100_000;
const MAX_SAMPLES: usize = 100_000;
const MAX_SINK_WRITE_BYTES: usize = 64 * 1024;
const REPORT_FILE: &str = "report.json";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Arguments {
    shape: XlsxCellCrudShape,
    workload: Workload,
    workers: usize,
    warmup: usize,
    samples: usize,
}

#[derive(Clone, Debug, Serialize)]
struct TimingSample {
    sample: usize,
    ingress_ns: u64,
    edit_commit_ns: u64,
    publication_ns: u64,
    commit_drop_ns: u64,
    total_ns: u64,
    output_bytes: usize,
    output_bytes_used: u64,
    cpu_tasks_used: u64,
    sink_write_calls: u64,
    sink_largest_write: usize,
    resources_after_drop: ResourceSnapshot,
    output_sha256: String,
    semantic_sha256: String,
    fixture_inverse_restored_exact_source: bool,
}

#[derive(Clone, Debug, Serialize)]
struct CaseReport {
    id: String,
    shape: String,
    workload: String,
    workers: usize,
    update_count: usize,
    source_bytes: usize,
    source_sha256: String,
    output_bytes: usize,
    output_sha256: String,
    semantic_sha256: String,
    expected_cpu_tasks: u64,
    admission_expected_worker_permits: u64,
    admission_observed_worker_permits: u64,
    admission_cpu_tasks_used: u64,
    preflight_actual_commit_inverse_verified: bool,
    context_limits: ContextLimits,
    sink_capacity_bytes: usize,
    sink_max_write_bytes: usize,
    warmup_iterations: usize,
    measured_samples: usize,
    samples: Vec<TimingSample>,
}

#[derive(Debug, Serialize)]
struct BenchmarkReport {
    schema: &'static str,
    status: &'static str,
    shape: String,
    workload: String,
    workers: usize,
    warmup_iterations: usize,
    measured_samples: usize,
    timing_scope: &'static str,
    phase_scope: &'static str,
    setup_scope: &'static str,
    verification_scope: &'static str,
    performance_claim: &'static str,
    case: CaseReport,
}

#[derive(Debug, Serialize)]
struct FailureReport {
    schema: &'static str,
    status: &'static str,
    shape: String,
    workload: String,
    workers: usize,
    warmup_iterations: usize,
    measured_samples: usize,
    error: String,
}

#[derive(Debug)]
struct BoundedSink {
    bytes: Vec<u8>,
    max_output_bytes: usize,
    max_write_bytes: usize,
    write_calls: u64,
    largest_write: usize,
}

impl BoundedSink {
    fn new(max_output_bytes: usize) -> AnyResult<Self> {
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(max_output_bytes)
            .map_err(|source| format!("managed benchmark sink reservation failed: {source}"))?;
        Ok(Self {
            bytes,
            max_output_bytes,
            max_write_bytes: MAX_SINK_WRITE_BYTES,
            write_calls: 0,
            largest_write: 0,
        })
    }

    fn reset(&mut self) {
        self.bytes.clear();
        self.write_calls = 0;
        self.largest_write = 0;
    }
}

impl Write for BoundedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.len() > self.max_write_bytes {
            return Err(io::Error::other(format!(
                "managed benchmark sink write {} exceeds {}",
                bytes.len(),
                self.max_write_bytes,
            )));
        }
        let new_len = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("managed benchmark sink length overflowed"))?;
        if new_len > self.max_output_bytes {
            return Err(io::Error::other(format!(
                "managed benchmark sink exceeded configured output cap {}",
                self.max_output_bytes,
            )));
        }
        if new_len > self.bytes.capacity() {
            return Err(io::Error::other(format!(
                "managed benchmark sink exceeded preallocated capacity {}",
                self.bytes.capacity(),
            )));
        }
        self.bytes.extend_from_slice(bytes);
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("managed benchmark sink write count overflowed"))?;
        self.largest_write = self.largest_write.max(bytes.len());
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone, Copy, Debug)]
struct Timing {
    ingress_ns: u64,
    edit_commit_ns: u64,
    publication_ns: u64,
    commit_drop_ns: u64,
    total_ns: u64,
}

struct TimedIteration {
    budget: Budget,
    context_limits: ContextLimits,
    timing: Timing,
    changed: bool,
    changed_cells: usize,
    cpu_tasks_before_publication: u64,
    output_bytes_before_publication: u64,
}

struct CasePlan<'a> {
    corpus: &'a Corpus,
    spec: &'a crate::XlsxCorpus,
    shape: XlsxCellCrudShape,
    workload: Workload,
    workers: usize,
    warmup: usize,
    samples: usize,
    updates: Vec<XlsxCoordinate>,
    expected_output: Vec<u8>,
    source_hash: String,
    output_hash: String,
    semantic_hash: String,
    inverse_patch: MultiPatch,
    expected_cpu_tasks: u64,
    payload_budget: usize,
    output_budget: u64,
    cache_limits: SourceCacheLimits,
    admission_expected_worker_permits: u64,
    admission_observed_worker_permits: u64,
    admission_cpu_tasks_used: u64,
    preflight_actual_commit_inverse_verified: bool,
}

fn usage() -> &'static str {
    "usage: xlsx_managed_compression_benchmark --shape <medium|dense-sparse> --workload <one-percent|one-edit> --workers <1|2|4> --warmup N --samples N --output <new-directory>"
}

fn parse_count(value: &str, label: &str, maximum: usize) -> AnyResult<usize> {
    let count = value
        .parse::<usize>()
        .map_err(|source| format!("invalid {label} '{value}': {source}"))?;
    if count > maximum {
        return Err(format!("{label} {count} exceeds maximum {maximum}").into());
    }
    Ok(count)
}

fn parse_shape(value: &str) -> AnyResult<XlsxCellCrudShape> {
    match value {
        "medium" => Ok(XlsxCellCrudShape::Medium),
        "dense-sparse" | "dense_sparse" => Ok(XlsxCellCrudShape::DenseSparse),
        other => Err(format!("unknown XLSX managed benchmark shape '{other}'").into()),
    }
}

fn parse_workload(value: &str) -> AnyResult<Workload> {
    match value {
        "one-percent" | "one_percent" => Ok(Workload::OnePercent),
        "one-edit" | "one_edit" => Ok(Workload::OneEdit),
        other => Err(format!("unknown XLSX managed benchmark workload '{other}'").into()),
    }
}

fn parse_workers(value: &str) -> AnyResult<usize> {
    match value {
        "1" => Ok(1),
        "2" => Ok(2),
        "4" => Ok(4),
        other => Err(format!("managed benchmark workers must be 1, 2, or 4, got '{other}'").into()),
    }
}

fn parse_args<I>(arguments: I) -> AnyResult<(Arguments, PathBuf)>
where
    I: IntoIterator<Item = OsString>,
{
    let mut shape = None;
    let mut workload = None;
    let mut workers = None;
    let mut warmup = 1;
    let mut samples = 5;
    let mut warmup_seen = false;
    let mut samples_seen = false;
    let mut output = None;
    let mut values = arguments.into_iter();
    while let Some(argument) = values.next() {
        let argument = argument.to_string_lossy().into_owned();
        match argument.as_str() {
            "--shape" => {
                if shape.is_some() {
                    return Err("--shape was specified more than once".into());
                }
                let value = values.next().ok_or("missing value for --shape")?;
                shape = Some(parse_shape(&value.to_string_lossy())?);
            },
            "--workload" => {
                if workload.is_some() {
                    return Err("--workload was specified more than once".into());
                }
                let value = values.next().ok_or("missing value for --workload")?;
                workload = Some(parse_workload(&value.to_string_lossy())?);
            },
            "--workers" => {
                if workers.is_some() {
                    return Err("--workers was specified more than once".into());
                }
                let value = values.next().ok_or("missing value for --workers")?;
                workers = Some(parse_workers(&value.to_string_lossy())?);
            },
            "--warmup" => {
                if warmup_seen {
                    return Err("--warmup was specified more than once".into());
                }
                warmup_seen = true;
                let value = values.next().ok_or("missing value for --warmup")?;
                warmup = parse_count(&value.to_string_lossy(), "warmup", MAX_WARMUP)?;
            },
            "--samples" => {
                if samples_seen {
                    return Err("--samples was specified more than once".into());
                }
                samples_seen = true;
                let value = values.next().ok_or("missing value for --samples")?;
                samples = parse_count(&value.to_string_lossy(), "samples", MAX_SAMPLES)?;
            },
            "--output" | "--output-dir" => {
                if output.is_some() {
                    return Err("--output was specified more than once".into());
                }
                output = Some(PathBuf::from(
                    values.next().ok_or("missing value for --output")?,
                ));
            },
            "--help" | "-h" => return Err(usage().into()),
            other => return Err(format!("unknown argument '{other}'\n{}", usage()).into()),
        }
    }
    let shape = shape.ok_or_else(|| format!("missing --shape\n{}", usage()))?;
    let workload = workload.ok_or_else(|| format!("missing --workload\n{}", usage()))?;
    let workers = workers.ok_or_else(|| format!("missing --workers\n{}", usage()))?;
    let output = output.ok_or_else(|| format!("missing --output\n{}", usage()))?;
    if samples == 0 {
        return Err("samples must be positive".into());
    }
    if shape == XlsxCellCrudShape::DenseSparse && workload == Workload::OneEdit {
        return Err("one-edit is available only for the medium XLSX corpus".into());
    }
    Ok((
        Arguments {
            shape,
            workload,
            workers,
            warmup,
            samples,
        },
        output,
    ))
}

fn elapsed_ns(started: Instant) -> AnyResult<u64> {
    u64::try_from(started.elapsed().as_nanos())
        .map_err(|_| "managed benchmark elapsed duration exceeds u64 nanoseconds".into())
}

fn prepare_inverse_patch(
    corpus: &Corpus,
    updates: &[XlsxCoordinate],
    cache_limits: SourceCacheLimits,
) -> AnyResult<MultiPatch> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let editor = SourceBackedEditor::from_read_at_with_options(
        source,
        OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(cache_limits),
    )?;
    let commit = super::stage_commit(&editor, updates)?;
    Ok(commit.patch().clone())
}

fn stage_timed_commit(
    editor: &SourceBackedEditor,
    updates: &[XlsxCoordinate],
) -> AnyResult<(litchi_xlsx::cell_values::Commit, bool, usize)> {
    let selectors = xlsx_update_sheet_selectors(updates);
    let mut edit = editor.edit_sheets(selectors)?;
    for coordinate in updates {
        edit.set(
            coordinate.sheet,
            litchi_xlsx::Address::at(
                u32::try_from(coordinate.row)?,
                u32::try_from(coordinate.column)?,
            )?,
            xlsx_value(*coordinate) + 1,
        )?;
    }
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let changed_cells = commit.diagnostics().changed_cells();
    Ok((commit, changed, changed_cells))
}

fn timed_iteration(plan: &CasePlan<'_>, sink: &mut BoundedSink) -> AnyResult<TimedIteration> {
    // Context, source cloning, cache policy, and sink reset are setup work.
    let (budget, cancellation_source, context) = managed_context(
        plan.payload_budget,
        plan.output_budget,
        plan.workers,
        plan.expected_cpu_tasks,
    )?;
    drop(cancellation_source);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(plan.corpus.archive.clone()));
    sink.reset();

    let total_started = Instant::now();
    let ingress_started = Instant::now();
    let editor = SourceBackedEditor::from_read_at_with_options(
        source,
        OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(plan.cache_limits)
            .execution_context(context),
    )?;
    let ingress_ns = elapsed_ns(ingress_started)?;

    let edit_started = Instant::now();
    let (commit, changed, changed_cells) = stage_timed_commit(&editor, &plan.updates)?;
    let edit_commit_ns = elapsed_ns(edit_started)?;
    let cpu_tasks_before_publication = budget.used(Resource::CpuTasks);
    let output_bytes_before_publication = budget.used(Resource::OutputBytes);

    let publication_started = Instant::now();
    let published = editor.publish_multi_commit_to_stream(sink, &commit)?;
    drop(published);
    let publication_ns = elapsed_ns(publication_started)?;
    let commit_drop_started = Instant::now();
    drop(commit);
    let commit_drop_ns = elapsed_ns(commit_drop_started)?;
    let total_ns = elapsed_ns(total_started)?;
    let phase_sum = ingress_ns
        .checked_add(edit_commit_ns)
        .and_then(|value| value.checked_add(publication_ns))
        .and_then(|value| value.checked_add(commit_drop_ns))
        .ok_or("managed benchmark phase duration overflowed")?;
    if phase_sum > total_ns {
        return Err(
            format!("managed benchmark phase sum {phase_sum} exceeds total {total_ns}").into(),
        );
    }
    Ok(TimedIteration {
        context_limits: super::context_limits(&budget, plan.workers),
        budget,
        timing: Timing {
            ingress_ns,
            edit_commit_ns,
            publication_ns,
            commit_drop_ns,
            total_ns,
        },
        changed,
        changed_cells,
        cpu_tasks_before_publication,
        output_bytes_before_publication,
    })
}

fn verify_iteration(
    plan: &CasePlan<'_>,
    sink: &BoundedSink,
    iteration: usize,
    timed: TimedIteration,
) -> AnyResult<TimingSample> {
    if !timed.changed || timed.changed_cells != plan.updates.len() {
        return Err(format!(
            "managed benchmark commit changed={} cells={}, expected {}",
            timed.changed,
            timed.changed_cells,
            plan.updates.len()
        )
        .into());
    }
    let output_bytes_used = timed
        .budget
        .used(Resource::OutputBytes)
        .checked_sub(timed.output_bytes_before_publication)
        .ok_or("managed benchmark OutputBytes counter moved backwards")?;
    let cpu_tasks_used = timed
        .budget
        .used(Resource::CpuTasks)
        .checked_sub(timed.cpu_tasks_before_publication)
        .ok_or("managed benchmark CpuTasks counter moved backwards")?;
    let resources_after_drop = require_clean_resources(&timed.budget, "managed benchmark")?;
    if output_bytes_used != u64::try_from(sink.bytes.len())? {
        return Err(format!(
            "managed benchmark charged {output_bytes_used} OutputBytes for {} sink bytes",
            sink.bytes.len()
        )
        .into());
    }
    if cpu_tasks_used != plan.expected_cpu_tasks {
        return Err(format!(
            "managed benchmark charged {} CpuTasks, expected {}",
            cpu_tasks_used, plan.expected_cpu_tasks,
        )
        .into());
    }
    if sink.bytes != plan.expected_output {
        return Err("managed benchmark publication bytes differ from the frozen output".into());
    }
    if sink.write_calls == 0 || sink.largest_write > MAX_SINK_WRITE_BYTES {
        return Err("managed benchmark sink boundary was not exercised as bounded".into());
    }
    let output_sha256 = sha256_hex(&sink.bytes);
    if output_sha256 != plan.output_hash {
        return Err(format!(
            "managed benchmark output hash {output_sha256} differs from {}",
            plan.output_hash,
        )
        .into());
    }
    verify_xlsx_cell_crud_output(plan.corpus, &sink.bytes, &plan.updates)?;
    let semantic_sha256 = xlsx_cell_values_output_semantic_hash(&sink.bytes, plan.spec)?;
    if semantic_sha256 != plan.semantic_hash {
        return Err("managed benchmark semantic output hash changed".into());
    }
    let mut package = OpcPackage::from_bytes(&sink.bytes, litchi_opc::SourceLimits::default())?;
    crate::xlsx_source_editors::replay_patch(&mut package, &plan.inverse_patch.inverse())?;
    let restored = PackageWriter::to_bytes(&package)?;
    if restored != plan.corpus.archive {
        return Err("managed benchmark inverse did not restore exact source ZIP".into());
    }
    if sha256_hex(&plan.corpus.archive) != plan.source_hash {
        return Err("managed benchmark source hash changed".into());
    }
    drop(timed.budget);
    Ok(TimingSample {
        sample: iteration,
        ingress_ns: timed.timing.ingress_ns,
        edit_commit_ns: timed.timing.edit_commit_ns,
        publication_ns: timed.timing.publication_ns,
        commit_drop_ns: timed.timing.commit_drop_ns,
        total_ns: timed.timing.total_ns,
        output_bytes: sink.bytes.len(),
        output_bytes_used,
        cpu_tasks_used,
        sink_write_calls: sink.write_calls,
        sink_largest_write: sink.largest_write,
        resources_after_drop,
        output_sha256,
        semantic_sha256,
        fixture_inverse_restored_exact_source: true,
    })
}

fn build_case_plan<'a>(
    corpus: &'a Corpus,
    spec: &'a crate::XlsxCorpus,
    arguments: Arguments,
) -> AnyResult<CasePlan<'a>> {
    let updates = match arguments.workload {
        Workload::OnePercent => spec.one_percent_updates.clone(),
        Workload::OneEdit => vec![
            spec.cell_inventory
                .as_ref()
                .and_then(|inventory| inventory.first())
                .and_then(|cells| cells.first())
                .copied()
                .ok_or("managed benchmark corpus has no one-edit target")?,
        ],
    };
    let expected_output = super::super::xlsx_cell_crud_eager_output(corpus, &updates)?;
    let source_hash = sha256_hex(&corpus.archive);
    let output_hash = sha256_hex(&expected_output);
    let semantic_hash = xlsx_cell_values_output_semantic_hash(&expected_output, spec)?;
    if source_hash != expected_source_sha256(arguments.shape) {
        return Err("managed benchmark source differs from the frozen source hash".into());
    }
    if output_hash != expected_output_sha256(arguments.shape, arguments.workload) {
        return Err("managed benchmark output differs from the frozen output hash".into());
    }
    verify_xlsx_cell_crud_output(corpus, &expected_output, &updates)?;

    let payload_budget = xlsx_cell_values_payload_budget(corpus, &updates)?;
    let output_budget = xlsx_output_ceiling(expected_output.len())?;
    let cache_limits = cache_limits(corpus, &updates)?;
    let inverse_patch = prepare_inverse_patch(corpus, &updates, cache_limits)?;
    let preflight = super::run_case(
        corpus,
        spec,
        arguments.shape,
        arguments.workload,
        &updates,
        arguments.workers,
        &expected_output,
    )?;
    if preflight.observed_worker_permits != preflight.expected_worker_permits {
        return Err(format!(
            "managed benchmark admission preflight observed {} worker permits, expected {}",
            preflight.observed_worker_permits, preflight.expected_worker_permits,
        )
        .into());
    }
    if preflight.cpu_tasks_used != arguments.workload.expected_cpu_tasks() {
        return Err(format!(
            "managed benchmark admission preflight observed {} CpuTasks, expected {}",
            preflight.cpu_tasks_used,
            arguments.workload.expected_cpu_tasks(),
        )
        .into());
    }
    if !preflight.inverse_restored_exact_source {
        return Err("managed benchmark admission preflight did not verify the actual managed-commit inverse".into());
    }
    Ok(CasePlan {
        corpus,
        spec,
        shape: arguments.shape,
        workload: arguments.workload,
        workers: arguments.workers,
        warmup: arguments.warmup,
        samples: arguments.samples,
        updates,
        expected_output,
        source_hash,
        output_hash,
        semantic_hash,
        inverse_patch,
        expected_cpu_tasks: arguments.workload.expected_cpu_tasks(),
        payload_budget,
        output_budget,
        cache_limits,
        admission_expected_worker_permits: preflight.expected_worker_permits,
        admission_observed_worker_permits: preflight.observed_worker_permits,
        admission_cpu_tasks_used: preflight.cpu_tasks_used,
        preflight_actual_commit_inverse_verified: preflight.inverse_restored_exact_source,
    })
}

fn run_case(plan: CasePlan<'_>) -> AnyResult<CaseReport> {
    let sink_capacity = usize::try_from(plan.output_budget)?;
    let mut context_limits = None;
    for iteration in 0..plan.warmup {
        let mut sink = BoundedSink::new(sink_capacity)?;
        let timed = timed_iteration(&plan, &mut sink)?;
        context_limits.get_or_insert_with(|| timed.context_limits.clone());
        verify_iteration(&plan, &sink, iteration, timed)?;
    }
    let mut measured = Vec::with_capacity(plan.samples);
    for sample in 0..plan.samples {
        let mut sink = BoundedSink::new(sink_capacity)?;
        let timed = timed_iteration(&plan, &mut sink)?;
        context_limits.get_or_insert_with(|| timed.context_limits.clone());
        measured.push(verify_iteration(&plan, &sink, sample, timed)?);
    }
    let context_limits = context_limits.ok_or("managed benchmark recorded no context limits")?;
    Ok(CaseReport {
        id: format!(
            "{}-{}-workers{}",
            plan.shape.name(),
            plan.workload.name(),
            plan.workers
        ),
        shape: plan.shape.name().to_owned(),
        workload: plan.workload.name().to_owned(),
        workers: plan.workers,
        update_count: plan.updates.len(),
        source_bytes: plan.corpus.archive.len(),
        source_sha256: plan.source_hash,
        output_bytes: plan.expected_output.len(),
        output_sha256: plan.output_hash,
        semantic_sha256: plan.semantic_hash,
        expected_cpu_tasks: plan.expected_cpu_tasks,
        admission_expected_worker_permits: plan.admission_expected_worker_permits,
        admission_observed_worker_permits: plan.admission_observed_worker_permits,
        admission_cpu_tasks_used: plan.admission_cpu_tasks_used,
        preflight_actual_commit_inverse_verified: plan.preflight_actual_commit_inverse_verified,
        context_limits,
        sink_capacity_bytes: usize::try_from(plan.output_budget)?,
        sink_max_write_bytes: MAX_SINK_WRITE_BYTES,
        warmup_iterations: plan.warmup,
        measured_samples: plan.samples,
        samples: measured,
    })
}

fn run(arguments: Arguments, directory: &Path) -> AnyResult<()> {
    let corpus = build_xlsx_cell_crud_corpus(arguments.shape)?;
    let spec = xlsx_spec(&corpus)?;
    let plan = build_case_plan(&corpus, spec, arguments)?;
    let case = run_case(plan)?;
    let report = BenchmarkReport {
        schema: SCHEMA,
        status: "completed",
        shape: arguments.shape.name().to_owned(),
        workload: arguments.workload.name().to_owned(),
        workers: arguments.workers,
        warmup_iterations: arguments.warmup,
        measured_samples: arguments.samples,
        timing_scope: "with fresh finite managed context prepared outside the clock, the timer begins at SourceBackedEditor construction and covers edit/commit, source-backed publication, returned snapshot drop, and commit/package worker-pool teardown; precomputed source/cache/context setup and bounded sink allocation are outside",
        phase_scope: "ingress_ns is the direct SourceBackedEditor constructor; edit_commit_ns is selector staging, cell sets, and MultiCommit::commit; publication_ns is publish_multi_commit_to_stream plus returned MultiSnapshot drop; commit_drop_ns is MultiCommit drop; total_ns starts at the constructor and ends after commit teardown",
        setup_scope: "deterministic corpus and expected output construction, frozen source/output/semantic oracles, SourceCacheLimits and output ceiling, one untimed admission preflight, inverse patch, fresh OwnedSource clone, finite managed context, and sink capacity reservation are outside measured phases",
        verification_scope: "exact output bytes, immutable fixture source hash, output/semantic hashes, workbook verification, sink bounds, CpuTasks/OutputBytes counters, and post-drop Memory/Workers/IoConcurrency/Objects checks run outside timers; the untimed admission preflight verifies an actual managed-commit inverse (preflight_actual_commit_inverse_verified=true), while each timed sample applies the precomputed inverse patch only as a fixture to its published bytes (fixture_inverse_restored_exact_source)",
        performance_claim: "none; raw phase timings describe this explicit in-memory harness and do not claim OS thread counts, speedup, allocation behavior, or physical I/O",
        case,
    };
    fs::write(
        directory.join(REPORT_FILE),
        serde_json::to_vec_pretty(&report)?,
    )?;
    Ok(())
}

/// Run one explicit shape/workload/worker benchmark case.
pub(crate) fn benchmark<I>(arguments: I) -> AnyResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let (arguments, directory) = parse_args(arguments)?;
    let parent = directory
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    let directory_created = match fs::create_dir(&directory) {
        Ok(()) => true,
        Err(source) if source.kind() == io::ErrorKind::AlreadyExists => {
            return Err(format!(
                "XLSX managed benchmark output directory already exists: {}",
                directory.display()
            )
            .into());
        },
        Err(source) => {
            return Err(format!(
                "could not create XLSX managed benchmark output directory {}: {source}",
                directory.display()
            )
            .into());
        },
    };
    let result = run(arguments, &directory);
    if let Err(error) = result {
        if directory_created {
            let failure = FailureReport {
                schema: SCHEMA,
                status: "failed",
                shape: arguments.shape.name().to_owned(),
                workload: arguments.workload.name().to_owned(),
                workers: arguments.workers,
                warmup_iterations: arguments.warmup,
                measured_samples: arguments.samples,
                error: error.to_string(),
            };
            if let Ok(bytes) = serde_json::to_vec_pretty(&failure) {
                let _ = fs::write(directory.join("failure.json"), bytes);
            }
        }
        return Err(error);
    }
    Ok(())
}
