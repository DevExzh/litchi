//! Untimed qualification of the source-backed XLSX compression scheduler.
//!
//! This module uses the existing deterministic cell-value corpora and the
//! public source-backed editor boundary.  It deliberately records admission
//! counters rather than elapsed time: a sink observes the worker permits held
//! when publication first emits bytes, while the managed budget records the
//! cumulative Deflate task count.  The observations prove the admitted wave
//! width; they do not claim an operating-system thread count or a speedup.

use super::{
    Corpus, XlsxCellCrudShape, XlsxCoordinate, build_xlsx_cell_crud_corpus, sha256_hex,
    verify_xlsx_cell_crud_output, xlsx_cell_values_output_semantic_hash,
    xlsx_cell_values_payload_budget, xlsx_output_ceiling, xlsx_spec, xlsx_update_sheet_selectors,
    xlsx_value,
};
use crate::xlsx_source_editors::SourceBackedEditor;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource,
};
use litchi_opc::{OpcPackage, PackageWriter, SourceCacheLimits, SourceLimits};
use litchi_xlsx::OpenOptions;
use litchi_xlsx::cell_values::Commit as MultiCommit;
use serde::Serialize;
use std::{
    error::Error,
    fs,
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
    path::Path,
    sync::Arc,
};

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx-managed-compression.v1";
const GENERATOR: &str = "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1";
const WORKERS: [usize; 3] = [1, 2, 4];
const MAX_WORKERS: u64 = 4;
const PLANNING_HEADROOM: u64 = 64 * 1024;
const INPUT_LIMIT: u64 = 64 * 1024 * 1024;
const OBJECT_LIMIT: u64 = 1_000_000;
const DEPTH_LIMIT: u64 = 256;
const WORK_LIMIT: u64 = 1_000_000_000;
const IO_CONCURRENCY_LIMIT: u64 = 1;

const MEDIUM_SOURCE_SHA256: &str =
    "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036";
const DENSE_SOURCE_SHA256: &str =
    "893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a";
const MEDIUM_ONE_PERCENT_SHA256: &str =
    "fc742a7ad139e2ce2e75fabfff8d2a33408dd5eeb10997717a2e2656d8423e6c";
const MEDIUM_ONE_EDIT_SHA256: &str =
    "9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1";
const DENSE_ONE_PERCENT_SHA256: &str =
    "07eee3f980949f031776087ed61f06125edd82cb29394a608a3d8d5a4b28374d";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Workload {
    OnePercent,
    OneEdit,
}

impl Workload {
    const fn name(self) -> &'static str {
        match self {
            Self::OnePercent => "one-percent",
            Self::OneEdit => "one-edit",
        }
    }

    const fn expected_cpu_tasks(self) -> u64 {
        match self {
            Self::OnePercent => 5,
            Self::OneEdit => 2,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct ResourceSnapshot {
    memory: u64,
    workers: u64,
    io_concurrency: u64,
    objects: u64,
}

#[derive(Clone, Debug, Serialize)]
struct ContextLimits {
    memory: u64,
    input_bytes: u64,
    output_bytes: u64,
    objects: u64,
    depth: u64,
    work: u64,
    execution_workers: usize,
    execution_in_flight_tasks: usize,
    execution_in_flight_bytes: u64,
    execution_min_parallel_bytes: u64,
    execution_min_task_bytes: u64,
    root_workers: u64,
    root_io_concurrency: u64,
    root_cpu_tasks: u64,
}

#[derive(Clone, Debug, Serialize)]
struct CaseRecord {
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
    output_bytes_used: u64,
    expected_cpu_tasks: u64,
    cpu_tasks_used: u64,
    expected_worker_permits: u64,
    observed_worker_permits: u64,
    sink_write_calls: u64,
    context_limits: ContextLimits,
    resources_after_drop: ResourceSnapshot,
    inverse_restored_exact_source: bool,
}

#[derive(Clone, Debug, Serialize)]
struct GateRecord {
    shape: String,
    managed_noop_exact_source: bool,
    managed_inverse_exact_source: bool,
    pre_cancelled_zero_output: bool,
    output_budget_zero_output: bool,
    output_refusal: OutputRefusalRecord,
}

#[derive(Clone, Debug, Serialize)]
struct OutputRefusalRecord {
    output_budget_limit: u64,
    cpu_tasks_used: u64,
    worker_permits_after_drop: u64,
    output_bytes_used_after_drop: u64,
    sink_write_calls: u64,
    zero_output: bool,
    typed_output_resource_refusal: bool,
    context_limits: ContextLimits,
    resources_after_drop: ResourceSnapshot,
}

#[derive(Debug, Serialize)]
struct Manifest {
    schema: &'static str,
    generator: &'static str,
    source_publication: &'static str,
    execution: &'static str,
    interpretation: &'static str,
    cases: Vec<CaseRecord>,
    gates: Vec<GateRecord>,
}

/// A sink that observes managed worker permits at the publication boundary.
/// The first write happens after the Deflate wave has completed, and all
/// writes are sampled so a transiently narrower first write cannot hide the
/// maximum admitted width.
#[derive(Debug)]
struct PermitObservingSink {
    bytes: Vec<u8>,
    budget: Budget,
    observed_worker_permits: u64,
    write_calls: u64,
}

impl PermitObservingSink {
    fn new(budget: Budget) -> Self {
        Self {
            bytes: Vec::new(),
            budget,
            observed_worker_permits: 0,
            write_calls: 0,
        }
    }
}

impl Write for PermitObservingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.observed_worker_permits = self
            .observed_worker_permits
            .max(self.budget.used(Resource::Workers));
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("sink write count overflowed"))?;
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn expected_source_sha256(shape: XlsxCellCrudShape) -> &'static str {
    match shape {
        XlsxCellCrudShape::Medium => MEDIUM_SOURCE_SHA256,
        XlsxCellCrudShape::DenseSparse => DENSE_SOURCE_SHA256,
        XlsxCellCrudShape::Noncompact | XlsxCellCrudShape::VendorExtension => {
            unreachable!("the managed compression qualification has two fixed shapes")
        },
    }
}

fn expected_output_sha256(shape: XlsxCellCrudShape, workload: Workload) -> &'static str {
    match (shape, workload) {
        (XlsxCellCrudShape::Medium, Workload::OnePercent) => MEDIUM_ONE_PERCENT_SHA256,
        (XlsxCellCrudShape::Medium, Workload::OneEdit) => MEDIUM_ONE_EDIT_SHA256,
        (XlsxCellCrudShape::DenseSparse, Workload::OnePercent) => DENSE_ONE_PERCENT_SHA256,
        (XlsxCellCrudShape::Noncompact | XlsxCellCrudShape::VendorExtension, _) => {
            unreachable!("the managed compression qualification has two fixed shapes")
        },
        (XlsxCellCrudShape::DenseSparse, Workload::OneEdit) => {
            unreachable!("dense-sparse one-edit is outside the frozen qualification")
        },
    }
}

fn expected_worker_permits(shape: XlsxCellCrudShape, workload: Workload, workers: usize) -> u64 {
    if shape == XlsxCellCrudShape::Medium && workload == Workload::OnePercent && workers > 1 {
        workers as u64
    } else {
        0
    }
}

fn managed_context(
    payload_budget: usize,
    output_budget: u64,
    workers: usize,
    cpu_tasks: u64,
) -> AnyResult<(Budget, CancellationSource, ExecutionContext)> {
    let payload_budget = u64::try_from(payload_budget)?;
    let worker_memory = soapberry_zip::office::WORKER_STATE_BYTES
        .checked_mul(MAX_WORKERS)
        .ok_or("managed compression worker memory bound overflowed")?;
    let memory = payload_budget
        .checked_add(worker_memory)
        .and_then(|value| value.checked_add(PLANNING_HEADROOM))
        .ok_or("managed compression memory bound overflowed")?;
    let worker_count =
        NonZeroUsize::new(workers).ok_or("managed compression worker count is zero")?;
    let execution_limits = ExecutionLimits::new(
        worker_count,
        worker_count,
        NonZeroU64::new(memory.max(1)).ok_or("managed compression byte bound is zero")?,
        0,
    )?;
    let budget = Budget::root(
        "litchi-perf-xlsx-managed-compression",
        Limits::new(
            memory,
            INPUT_LIMIT,
            output_budget,
            OBJECT_LIMIT,
            DEPTH_LIMIT,
            WORK_LIMIT,
        )
        .with_execution_io(workers as u64, IO_CONCURRENCY_LIMIT, cpu_tasks),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    Ok((budget, cancellation_source, context))
}

fn context_limits(budget: &Budget, workers: usize) -> ContextLimits {
    ContextLimits {
        memory: budget.limit(Resource::Memory),
        input_bytes: budget.limit(Resource::InputBytes),
        output_bytes: budget.limit(Resource::OutputBytes),
        objects: budget.limit(Resource::Objects),
        depth: budget.limit(Resource::Depth),
        work: budget.limit(Resource::Work),
        execution_workers: workers,
        execution_in_flight_tasks: workers,
        execution_in_flight_bytes: budget.limit(Resource::Memory),
        execution_min_parallel_bytes: 0,
        execution_min_task_bytes: 0,
        root_workers: budget.limit(Resource::Workers),
        root_io_concurrency: budget.limit(Resource::IoConcurrency),
        root_cpu_tasks: budget.limit(Resource::CpuTasks),
    }
}

fn cache_limits(corpus: &Corpus, updates: &[XlsxCoordinate]) -> AnyResult<SourceCacheLimits> {
    let payload_budget = xlsx_cell_values_payload_budget(corpus, updates)?;
    let selected = xlsx_update_sheet_selectors(updates).len();
    let max_entries = selected
        .checked_add(4)
        .ok_or("managed compression cache entry bound overflowed")?;
    Ok(SourceCacheLimits::new(payload_budget.max(1), max_entries)?)
}

fn open_editor(
    corpus: &Corpus,
    updates: &[XlsxCoordinate],
    context: ExecutionContext,
) -> AnyResult<SourceBackedEditor> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    Ok(SourceBackedEditor::from_read_at_with_options(
        source,
        OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(cache_limits(corpus, updates)?)
            .execution_context(context),
    )?)
}

fn stage_commit(editor: &SourceBackedEditor, updates: &[XlsxCoordinate]) -> AnyResult<MultiCommit> {
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
    if !commit.diagnostics().changed() || commit.diagnostics().changed_cells() != updates.len() {
        return Err("managed compression commit changed an unexpected cell count".into());
    }
    Ok(commit)
}

fn resources(budget: &Budget) -> ResourceSnapshot {
    ResourceSnapshot {
        memory: budget.used(Resource::Memory),
        workers: budget.used(Resource::Workers),
        io_concurrency: budget.used(Resource::IoConcurrency),
        objects: budget.used(Resource::Objects),
    }
}

fn require_clean_resources(budget: &Budget, label: &str) -> AnyResult<ResourceSnapshot> {
    let snapshot = resources(budget);
    if snapshot.memory != 0
        || snapshot.workers != 0
        || snapshot.io_concurrency != 0
        || snapshot.objects != 0
    {
        return Err(
            format!("{label} retained managed resources after publication: {snapshot:?}").into(),
        );
    }
    Ok(snapshot)
}

fn inverse_round_trip(
    corpus: &Corpus,
    commit: &MultiCommit,
    published_output: &[u8],
) -> AnyResult<bool> {
    let mut package =
        OpcPackage::from_bytes(published_output, litchi_opc::SourceLimits::default())?;
    crate::xlsx_source_editors::replay_patch(&mut package, &commit.patch().inverse())?;
    let restored = PackageWriter::to_bytes(&package)?;
    if restored != corpus.archive {
        return Err("managed-origin inverse patch did not restore exact source ZIP".into());
    }
    Ok(true)
}

fn run_case(
    corpus: &Corpus,
    spec: &super::XlsxCorpus,
    shape: XlsxCellCrudShape,
    workload: Workload,
    updates: &[XlsxCoordinate],
    workers: usize,
    expected_output: &[u8],
) -> AnyResult<CaseRecord> {
    let expected_tasks = workload.expected_cpu_tasks();
    let output_budget = xlsx_output_ceiling(expected_output.len())?;
    let payload_budget = xlsx_cell_values_payload_budget(corpus, updates)?;
    let (budget, _cancellation, context) =
        managed_context(payload_budget, output_budget, workers, expected_tasks)?;
    let editor = open_editor(corpus, updates, context)?;
    let commit = stage_commit(&editor, updates)?;
    let cpu_tasks_before = budget.used(Resource::CpuTasks);
    let mut sink = PermitObservingSink::new(budget.clone());
    editor.publish_multi_commit_to_stream(&mut sink, &commit)?;
    let output_bytes_used = budget.used(Resource::OutputBytes);
    let observed_worker_permits = sink.observed_worker_permits;
    let sink_write_calls = sink.write_calls;
    let output = sink.bytes;
    let cpu_tasks_after = budget.used(Resource::CpuTasks);
    let cpu_tasks_used = cpu_tasks_after
        .checked_sub(cpu_tasks_before)
        .ok_or("managed compression CpuTasks counter moved backwards")?;
    if output_bytes_used != u64::try_from(output.len())? {
        return Err(format!(
            "{} {} workers {} charged {output_bytes_used} OutputBytes for {} sink bytes",
            shape.name(),
            workload.name(),
            workers,
            output.len(),
        )
        .into());
    }
    let inverse_restored_exact_source = inverse_round_trip(corpus, &commit, &output)?;
    drop(commit);
    let context_limits = context_limits(&budget, workers);
    let resources_after_drop = require_clean_resources(&budget, "managed compression case")?;
    if cpu_tasks_used != expected_tasks {
        return Err(format!(
            "{} {} workers {} admitted {cpu_tasks_used} CpuTasks, expected {expected_tasks}",
            shape.name(),
            workload.name(),
            workers,
        )
        .into());
    }
    let expected_workers = expected_worker_permits(shape, workload, workers);
    if observed_worker_permits != expected_workers {
        return Err(format!(
            "{} {} workers {} admitted {} worker permits, expected {expected_workers}",
            shape.name(),
            workload.name(),
            workers,
            observed_worker_permits,
        )
        .into());
    }
    if output != expected_output {
        return Err(format!(
            "{} {} workers {} changed deterministic publication bytes",
            shape.name(),
            workload.name(),
            workers,
        )
        .into());
    }
    verify_xlsx_cell_crud_output(corpus, &output, updates)?;
    let output_sha256 = sha256_hex(&output);
    if output_sha256 != expected_output_sha256(shape, workload) {
        return Err(format!(
            "{} {} output hash {output_sha256} differs from frozen qualification hash",
            shape.name(),
            workload.name(),
        )
        .into());
    }
    let source_sha256 = sha256_hex(&corpus.archive);
    if source_sha256 != expected_source_sha256(shape) {
        return Err(format!(
            "{} source hash {source_sha256} differs from frozen qualification hash",
            shape.name()
        )
        .into());
    }
    let semantic_sha256 = xlsx_cell_values_output_semantic_hash(&output, spec)?;
    Ok(CaseRecord {
        id: format!("{}-{}-workers{workers}", shape.name(), workload.name()),
        shape: shape.name().to_owned(),
        workload: workload.name().to_owned(),
        workers,
        update_count: updates.len(),
        source_bytes: corpus.archive.len(),
        source_sha256,
        output_bytes: output.len(),
        output_sha256,
        semantic_sha256,
        output_bytes_used,
        expected_cpu_tasks: expected_tasks,
        cpu_tasks_used,
        expected_worker_permits: expected_workers,
        observed_worker_permits,
        sink_write_calls,
        context_limits,
        resources_after_drop,
        inverse_restored_exact_source,
    })
}

fn managed_noop(corpus: &Corpus, spec: &super::XlsxCorpus) -> AnyResult<bool> {
    let selectors = (0..spec.sheet_count)
        .map(litchi_xlsx::Selector::from)
        .collect::<Vec<_>>();
    let updates: Vec<XlsxCoordinate> = spec
        .cell_inventory
        .as_ref()
        .map_or_else(Vec::new, |inventory| {
            inventory.iter().flatten().copied().collect()
        });
    let payload_budget = xlsx_cell_values_payload_budget(corpus, &updates)?;
    let output_budget = xlsx_output_ceiling(corpus.archive.len())?;
    let (budget, _cancellation, context) = managed_context(payload_budget, output_budget, 4, 1)?;
    let editor = open_editor(corpus, &updates, context)?;
    let commit = editor.edit_sheets(selectors)?.commit()?;
    if commit.diagnostics().changed() || !commit.patch().is_noop() {
        return Err("managed exact no-op produced a non-empty patch".into());
    }
    let mut sink = PermitObservingSink::new(budget.clone());
    editor.publish_multi_commit_to_stream(&mut sink, &commit)?;
    let output_bytes_used = budget.used(Resource::OutputBytes);
    let cpu_tasks_used = budget.used(Resource::CpuTasks);
    drop(commit);
    let _resources_after_drop = require_clean_resources(&budget, "managed no-op")?;
    if output_bytes_used != u64::try_from(sink.bytes.len())?
        || cpu_tasks_used != 0
        || sink.bytes != corpus.archive
        || sink.observed_worker_permits != 0
    {
        return Err(format!(
            "managed exact no-op changed bytes or counters: output_bytes_used={output_bytes_used}, sink_bytes={}, cpu_tasks_used={cpu_tasks_used}, worker_permits={}",
            sink.bytes.len(),
            sink.observed_worker_permits,
        )
        .into());
    }
    Ok(true)
}

fn managed_cancellation(corpus: &Corpus, updates: &[XlsxCoordinate]) -> AnyResult<bool> {
    let payload_budget = xlsx_cell_values_payload_budget(corpus, updates)?;
    let output_budget = xlsx_output_ceiling(corpus.archive.len())?;
    let (budget, cancellation_source, context) = managed_context(
        payload_budget,
        output_budget,
        4,
        Workload::OnePercent.expected_cpu_tasks(),
    )?;
    let editor = open_editor(corpus, updates, context)?;
    let commit = stage_commit(&editor, updates)?;
    cancellation_source.cancel();
    let mut output = Vec::new();
    let error = editor
        .publish_multi_commit_to_stream(&mut output, &commit)
        .err()
        .ok_or("pre-cancelled managed publication unexpectedly succeeded")?;
    let output_bytes_used = budget.used(Resource::OutputBytes);
    let cpu_tasks_used = budget.used(Resource::CpuTasks);
    let error_text = error.to_string();
    let cancelled = matches!(
        &error,
        litchi_xlsx::Error::Package(
            litchi_opc::OpcError::Cancelled
                | litchi_opc::OpcError::Execution(ExecutionError::Cancelled)
        )
    );
    drop(commit);
    let _resources_after_drop = require_clean_resources(&budget, "managed cancellation")?;
    if !cancelled || !output.is_empty() || output_bytes_used != 0 || cpu_tasks_used != 0 {
        return Err(format!(
            "pre-cancelled managed publication returned {error_text}, emitted {} bytes, charged {output_bytes_used} OutputBytes and {cpu_tasks_used} CpuTasks",
            output.len(),
        )
        .into());
    }
    Ok(true)
}

fn managed_output_refusal(
    corpus: &Corpus,
    updates: &[XlsxCoordinate],
) -> AnyResult<OutputRefusalRecord> {
    let payload_budget = xlsx_cell_values_payload_budget(corpus, updates)?;
    let (budget, _cancellation, context) = managed_context(
        payload_budget,
        0,
        4,
        Workload::OnePercent.expected_cpu_tasks(),
    )?;
    let context_limits = context_limits(&budget, 4);
    let editor = open_editor(corpus, updates, context)?;
    let commit = stage_commit(&editor, updates)?;
    let cpu_tasks_before = budget.used(Resource::CpuTasks);
    let mut sink = PermitObservingSink::new(budget.clone());
    let error = editor
        .publish_multi_commit_to_stream(&mut sink, &commit)
        .err()
        .ok_or("zero-output managed publication unexpectedly succeeded")?;
    let cpu_tasks_after = budget.used(Resource::CpuTasks);
    let cpu_tasks_used = cpu_tasks_after
        .checked_sub(cpu_tasks_before)
        .ok_or("managed output refusal CpuTasks counter moved backwards")?;
    let typed_output_resource_refusal = super::xlsx_cell_values_is_output_resource_refusal(&error);
    let zero_output = sink.bytes.is_empty();
    let sink_write_calls = sink.write_calls;
    drop(commit);
    let resources_after_drop = require_clean_resources(&budget, "managed output refusal")?;
    let worker_permits_after_drop = budget.used(Resource::Workers);
    let output_bytes_used_after_drop = budget.used(Resource::OutputBytes);
    if cpu_tasks_used != Workload::OnePercent.expected_cpu_tasks()
        || !typed_output_resource_refusal
        || !zero_output
        || sink_write_calls != 0
        || worker_permits_after_drop != 0
        || output_bytes_used_after_drop != 0
    {
        return Err(format!(
            "managed zero-output refusal failed its typed/counter gates: error={error:?}, cpu_tasks_used={cpu_tasks_used}, typed_output_resource_refusal={typed_output_resource_refusal}, zero_output={zero_output}, sink_write_calls={sink_write_calls}, worker_permits_after_drop={worker_permits_after_drop}, output_bytes_used_after_drop={output_bytes_used_after_drop}"
        )
        .into());
    }
    Ok(OutputRefusalRecord {
        output_budget_limit: 0,
        cpu_tasks_used,
        worker_permits_after_drop,
        output_bytes_used_after_drop,
        sink_write_calls,
        zero_output,
        typed_output_resource_refusal,
        context_limits,
        resources_after_drop,
    })
}

/// Run the untimed, deterministic managed compression qualification.
///
/// `directory` must not exist.  Only a small JSON manifest is written; source
/// and output ZIP bytes remain in memory and are identified by their hashes.
pub(crate) fn qualify(directory: &Path) -> AnyResult<()> {
    if directory.exists() {
        return Err(format!(
            "XLSX managed compression directory already exists: {}",
            directory.display()
        )
        .into());
    }
    fs::create_dir_all(directory)?;

    let mut cases = Vec::new();
    let mut gates = Vec::new();
    for shape in XlsxCellCrudShape::ALL {
        let corpus = build_xlsx_cell_crud_corpus(shape)?;
        let spec = xlsx_spec(&corpus)?;
        let one_percent = spec.one_percent_updates.clone();
        let expected_one_percent = super::xlsx_cell_crud_eager_output(&corpus, &one_percent)?;
        verify_xlsx_cell_crud_output(&corpus, &expected_one_percent, &one_percent)?;

        let gate_noop = managed_noop(&corpus, spec)?;
        let gate_cancellation = managed_cancellation(&corpus, &one_percent)?;
        let gate_refusal = managed_output_refusal(&corpus, &one_percent)?;

        let case_start = cases.len();
        let mut workloads = vec![(Workload::OnePercent, one_percent, expected_one_percent)];
        if shape == XlsxCellCrudShape::Medium {
            let update = spec
                .cell_inventory
                .as_ref()
                .and_then(|inventory| inventory.first())
                .and_then(|cells| cells.first())
                .copied()
                .ok_or("managed compression corpus has no one-edit target")?;
            let updates = vec![update];
            let expected_output = super::xlsx_cell_crud_eager_output(&corpus, &updates)?;
            workloads.push((Workload::OneEdit, updates, expected_output));
        }
        for (workload, updates, expected_output) in workloads {
            for workers in WORKERS {
                cases.push(run_case(
                    &corpus,
                    spec,
                    shape,
                    workload,
                    &updates,
                    workers,
                    &expected_output,
                )?);
            }
        }
        let managed_inverse_exact_source = cases[case_start..]
            .iter()
            .all(|case| case.inverse_restored_exact_source);
        gates.push(GateRecord {
            shape: shape.name().to_owned(),
            managed_noop_exact_source: gate_noop,
            managed_inverse_exact_source,
            pre_cancelled_zero_output: gate_cancellation,
            output_budget_zero_output: gate_refusal.zero_output,
            output_refusal: gate_refusal,
        });
    }

    let manifest = Manifest {
        schema: SCHEMA,
        generator: GENERATOR,
        source_publication: "SourceBackedEditor::edit_sheets -> set -> commit -> publish_multi_commit_to_stream",
        execution: "fresh managed context and source package for each case; workers 1, 2, and 4",
        interpretation: "sink worker permits prove admitted compression width, not operating-system thread counts or speedup; inverse checks use bounded owning MultiPatch::apply_materialized outside managed stream publication",
        cases,
        gates,
    };
    fs::write(
        directory.join("manifest.json"),
        serde_json::to_vec_pretty(&manifest)?,
    )?;
    Ok(())
}

#[path = "xlsx_managed_compression_benchmark.rs"]
pub(crate) mod benchmark;
