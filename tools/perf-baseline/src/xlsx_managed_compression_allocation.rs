//! Allocator-only attribution for the managed XLSX compression route.
//!
//! This module is a child of the timed benchmark module so it can reuse the
//! fixed corpus, case plan, bounded sink, context construction, commit
//! staging, and semantic/inverse oracles without changing the native timing
//! path.  It deliberately records allocator regions only.  It does not emit
//! elapsed phase values, RSS, throughput, native timing, or adoption claims.

use super::{
    Arguments, BoundedSink, CasePlan, ContextLimits, ResourceSnapshot, build_case_plan,
    build_xlsx_cell_crud_corpus, managed_context, require_clean_resources, sha256_hex,
    stage_timed_commit, verify_xlsx_cell_crud_output, xlsx_cell_values_output_semantic_hash,
    xlsx_spec,
};
use crate::allocation_metrics::{self, Sample as AllocationSample, Status as AllocationStatus};
use crate::xlsx_source_editors::SourceBackedEditor;
use litchi_core::{Budget, OwnedSource, ReadAt, Resource};
use litchi_opc::{OpcPackage, PackageWriter, SourceLimits};
use litchi_xlsx::OpenOptions;
use serde::Serialize;
use std::{
    error::Error,
    ffi::OsString,
    fs::{self, File},
    io::{Read, Seek, SeekFrom, Write},
    path::Path,
    sync::Arc,
    thread,
    time::{Duration, Instant},
};

type AllocationResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx-managed-compression-allocation.v1";
const REPORT_FILE: &str = "report.json";
const FAILURE_FILE: &str = "failure.json";
const CALIBRATION_POLLS: usize = 100;
const THREAD_BUFFER_BYTES: usize = 16 * 1024;
const THREAD_DEADLINE: Duration = Duration::from_secs(2);
const THREAD_POLL_LIMIT: usize = 10_000;
const THREAD_POLL_SLEEP: Duration = Duration::from_millis(1);
const THREAD_STATUS_PATH: &str = "/proc/self/status";
const ANALYZER_SCOPE: &str = "global system-allocator callbacks serialized by allocation_metrics; process-wide across threads; observer peaks are not RSS, unique heap, per-thread bytes, OS-thread counts, or timing";
const THREAD_SCOPE: &str = "calibration uses a separate measured allocator region; pre-sample /proc/self/status probes are outside operation regions; the post-core fixed-buffer drain is the fifth in-region segment; observed process thread counts do not establish admitted-worker/thread mapping or task overlap";
const PEAK_SEMANTICS: &str = "observer-ordered process-live high-water evidence, not RSS, unique heap, per-thread bytes, or OS-thread width";

#[derive(Clone, Debug, Serialize)]
struct BinaryIdentity {
    path: String,
    bytes: u64,
    sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct ThreadDrain {
    first_thread_count: usize,
    final_thread_count: usize,
    polls: usize,
    drained: bool,
    deadline_seconds: u64,
    poll_limit: usize,
    poll_sleep_milliseconds: u64,
}

#[derive(Clone, Debug, Serialize)]
struct Calibration {
    status: &'static str,
    polls: usize,
    thread_count_first: usize,
    thread_count_last: usize,
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_allocation_calls: u64,
    no_allocator_callbacks: bool,
    poll_sleep_milliseconds: u64,
    thread_scope: &'static str,
}

#[derive(Clone, Debug, Serialize)]
struct AllocationRow {
    sample: usize,
    setup: AllocationSample,
    ingress: AllocationSample,
    edit_commit: AllocationSample,
    publication: AllocationSample,
    commit_drop: AllocationSample,
    worker_drain: AllocationSample,
    core_combined: AllocationSample,
    core_combined_with_drain: AllocationSample,
    postdrop: AllocationSample,
    max_setup_core_peak_above_setup_live_before: u64,
    max_setup_drained_core_peak_above_setup_live_before: u64,
    context_limits: ContextLimits,
    workers_requested: usize,
    expected_worker_permits: u64,
    admitted_worker_permits: u64,
    expected_cpu_tasks: u64,
    cpu_tasks_used: u64,
    output_bytes: usize,
    output_bytes_used: u64,
    sink_write_calls: u64,
    sink_largest_write: usize,
    resources_after_core: ResourceSnapshot,
    source_sha256: String,
    output_sha256: String,
    semantic_sha256: String,
    fixture_inverse_restored_exact_source: bool,
    pre_sample_threads: ThreadDrain,
    post_core_threads: ThreadDrain,
}

#[derive(Clone, Debug, Serialize)]
struct AllocationCaseReport {
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
    rows: Vec<AllocationRow>,
}

#[derive(Debug, Serialize)]
struct AllocationReport {
    schema: &'static str,
    status: &'static str,
    binary_identity: BinaryIdentity,
    allocator_identity: &'static str,
    instrumentation_identity: &'static str,
    counter_revision: &'static str,
    allocation_scope: &'static str,
    thread_observation_scope: &'static str,
    timing_scope: &'static str,
    performance_claim: &'static str,
    peak_semantics: &'static str,
    calibration: Calibration,
    case: AllocationCaseReport,
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
    scope: &'static str,
}

#[derive(Debug)]
struct Verification {
    output_bytes: usize,
    output_bytes_used: u64,
    cpu_tasks_used: u64,
    sink_write_calls: u64,
    sink_largest_write: usize,
    resources_after_core: ResourceSnapshot,
    output_sha256: String,
    semantic_sha256: String,
    fixture_inverse_restored_exact_source: bool,
}

#[derive(Debug)]
struct ThreadObserver {
    status: File,
    buffer: [u8; THREAD_BUFFER_BYTES],
}

#[derive(Clone, Copy, Debug)]
enum ThreadProbeError {
    Io,
    Missing,
    Malformed,
    Overflow,
}

impl std::fmt::Display for ThreadProbeError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Io => "procfs status read failed",
            Self::Missing => "procfs Threads field is missing",
            Self::Malformed => "procfs Threads field is malformed",
            Self::Overflow => "procfs Threads field overflowed",
        })
    }
}

impl Error for ThreadProbeError {}

impl ThreadObserver {
    fn open() -> AllocationResult<Self> {
        Ok(Self {
            status: File::open(THREAD_STATUS_PATH)?,
            buffer: [0; THREAD_BUFFER_BYTES],
        })
    }

    /// Reads `Threads:` without allocating.  The file is opened before any
    /// measured region, and this method uses only the fixed stack buffer.
    fn read_threads(&mut self) -> Result<usize, ThreadProbeError> {
        self.status
            .seek(SeekFrom::Start(0))
            .map_err(|_| ThreadProbeError::Io)?;
        let length = self
            .status
            .read(&mut self.buffer)
            .map_err(|_| ThreadProbeError::Io)?;
        let key = b"Threads:";
        let mut start = 0;
        while start + key.len() <= length {
            if &self.buffer[start..start + key.len()] == key {
                let mut cursor = start + key.len();
                while cursor < length
                    && (self.buffer[cursor] == b' ' || self.buffer[cursor] == b'\t')
                {
                    cursor += 1;
                }
                if cursor == length {
                    return Err(ThreadProbeError::Malformed);
                }
                let mut value = 0usize;
                let mut digits = 0usize;
                while cursor < length
                    && self.buffer[cursor] != b'\n'
                    && self.buffer[cursor] != b'\r'
                {
                    let digit = self.buffer[cursor].wrapping_sub(b'0');
                    if digit > 9 {
                        return Err(ThreadProbeError::Malformed);
                    }
                    value = value.checked_mul(10).ok_or(ThreadProbeError::Overflow)?;
                    value = value
                        .checked_add(usize::from(digit))
                        .ok_or(ThreadProbeError::Overflow)?;
                    digits += 1;
                    cursor += 1;
                }
                return if digits == 0 {
                    Err(ThreadProbeError::Malformed)
                } else {
                    Ok(value)
                };
            }
            start += 1;
        }
        Err(ThreadProbeError::Missing)
    }

    /// Waits for exactly one process thread with a bounded, authoritative
    /// deadline. Pre-sample calls are outside operation regions; the post-core
    /// call is the measured worker-drain allocation segment, not native timing.
    fn drain_to_one(&mut self) -> Result<ThreadDrain, ThreadProbeError> {
        let started = Instant::now();
        let mut first = None;
        let mut last = 0;
        let mut polls = 0;
        for _ in 0..THREAD_POLL_LIMIT {
            polls += 1;
            let count = self.read_threads()?;
            first.get_or_insert(count);
            last = count;
            if started.elapsed() >= THREAD_DEADLINE {
                break;
            }
            if count == 1 {
                return Ok(ThreadDrain {
                    first_thread_count: first.unwrap_or(count),
                    final_thread_count: count,
                    polls,
                    drained: true,
                    deadline_seconds: THREAD_DEADLINE.as_secs(),
                    poll_limit: THREAD_POLL_LIMIT,
                    poll_sleep_milliseconds: THREAD_POLL_SLEEP.as_millis() as u64,
                });
            }
            thread::sleep(THREAD_POLL_SLEEP);
        }
        Ok(ThreadDrain {
            first_thread_count: first.unwrap_or(last),
            final_thread_count: last,
            polls,
            drained: false,
            deadline_seconds: THREAD_DEADLINE.as_secs(),
            poll_limit: THREAD_POLL_LIMIT,
            poll_sleep_milliseconds: THREAD_POLL_SLEEP.as_millis() as u64,
        })
    }
}

fn measured_sample(
    sample: Option<AllocationSample>,
    label: &'static str,
) -> AllocationResult<AllocationSample> {
    match sample {
        Some(sample)
            if sample.status == AllocationStatus::Measured
                && sample.failed_allocation_calls == Some(0) =>
        {
            Ok(sample)
        },
        Some(sample) if sample.status == AllocationStatus::Measured => {
            Err(format!("{label} allocator region recorded a failed allocation").into())
        },
        Some(sample) => Err(format!(
            "{label} allocator region was not measured: {:?}",
            sample.status
        )
        .into()),
        None => Err(format!("{label} allocator region did not publish a sample").into()),
    }
}

fn required_sample_field(
    sample: &AllocationSample,
    value: Option<u64>,
    label: &'static str,
) -> AllocationResult<u64> {
    if sample.status != AllocationStatus::Measured {
        return Err(format!("{label} sample was not measured").into());
    }
    value.ok_or_else(|| format!("{label} sample omitted a measured field").into())
}

fn sum_sample_field<F>(
    samples: &[&AllocationSample],
    label: &'static str,
    field: F,
) -> AllocationResult<u64>
where
    F: Fn(&AllocationSample) -> Option<u64>,
{
    samples.iter().try_fold(0u64, |total, sample| {
        let value = required_sample_field(sample, field(sample), label)?;
        total
            .checked_add(value)
            .ok_or_else(|| format!("{label} sample sum overflowed").into())
    })
}

fn combine_samples(
    samples: &[&AllocationSample],
    label: &'static str,
) -> AllocationResult<AllocationSample> {
    let first = samples
        .first()
        .copied()
        .ok_or_else(|| format!("{label} has no phase samples"))?;
    let last = samples
        .last()
        .copied()
        .ok_or_else(|| format!("{label} has no final phase sample"))?;
    if samples
        .iter()
        .any(|sample| sample.status != AllocationStatus::Measured)
    {
        return Err(format!("{label} contains an unavailable allocator phase").into());
    }
    for pair in samples.windows(2) {
        let left_after = required_sample_field(pair[0], pair[0].live_bytes_after, label)?;
        let right_before = required_sample_field(pair[1], pair[1].live_bytes_before, label)?;
        if left_after != right_before {
            return Err(format!("{label} live-byte boundaries are not adjacent").into());
        }
    }
    let region_peak_live_bytes = samples.iter().try_fold(None, |maximum, sample| {
        let peak = required_sample_field(sample, sample.region_peak_live_bytes, label)?;
        Ok::<_, Box<dyn Error>>(Some(maximum.map_or(peak, |value: u64| value.max(peak))))
    })?;
    let combined = AllocationSample {
        status: AllocationStatus::Measured,
        scope: first.scope,
        allocation_calls: Some(sum_sample_field(samples, label, |sample| {
            sample.allocation_calls
        })?),
        deallocation_calls: Some(sum_sample_field(samples, label, |sample| {
            sample.deallocation_calls
        })?),
        reallocation_calls: Some(sum_sample_field(samples, label, |sample| {
            sample.reallocation_calls
        })?),
        failed_allocation_calls: Some(sum_sample_field(samples, label, |sample| {
            sample.failed_allocation_calls
        })?),
        allocated_bytes: Some(sum_sample_field(samples, label, |sample| {
            sample.allocated_bytes
        })?),
        deallocated_bytes: Some(sum_sample_field(samples, label, |sample| {
            sample.deallocated_bytes
        })?),
        live_bytes_before: Some(required_sample_field(
            first,
            first.live_bytes_before,
            label,
        )?),
        live_bytes_after: Some(required_sample_field(last, last.live_bytes_after, label)?),
        peak_live_bytes_before: Some(required_sample_field(
            first,
            first.peak_live_bytes_before,
            label,
        )?),
        peak_live_bytes_after: Some(required_sample_field(
            last,
            last.peak_live_bytes_after,
            label,
        )?),
        region_peak_live_bytes,
    };
    Ok(combined)
}

fn verify_core_conservation(
    ingress: &AllocationSample,
    edit_commit: &AllocationSample,
    publication: &AllocationSample,
    commit_drop: &AllocationSample,
    worker_drain: &AllocationSample,
    core_combined_with_drain: &AllocationSample,
) -> AllocationResult<AllocationSample> {
    let four_phases = [ingress, edit_commit, publication, commit_drop];
    let core_combined = combine_samples(&four_phases, "core")?;
    let five_phases = [ingress, edit_commit, publication, commit_drop, worker_drain];
    let expected_with_drain = combine_samples(&five_phases, "core plus worker drain")?;
    if *core_combined_with_drain != expected_with_drain {
        return Err("core plus worker-drain allocator counters failed conservation".into());
    }
    Ok(core_combined)
}

fn derive_max_setup_core_peak_above_setup_live_before(
    setup: &AllocationSample,
    core: &AllocationSample,
) -> AllocationResult<u64> {
    let setup_live_before = required_sample_field(setup, setup.live_bytes_before, "setup")?;
    let setup_peak = required_sample_field(setup, setup.region_peak_live_bytes, "setup")?;
    let core_peak = required_sample_field(core, core.region_peak_live_bytes, "core")?;
    setup_peak
        .max(core_peak)
        .checked_sub(setup_live_before)
        .ok_or_else(|| "setup/core peak was below setup live baseline".into())
}

fn verify_instrumentation() -> AllocationResult<()> {
    if allocation_metrics::instrumentation_identity() != "system_allocator_operation_scoped"
        || allocation_metrics::counter_revision() != Some("serialized_region_peak_v3")
        || allocation_metrics::allocator_identity() != "CountingSystemAllocator(std::alloc::System)"
    {
        return Err("allocator instrumentation identity is not serialized_region_peak_v3".into());
    }
    Ok(())
}

fn calibrate(observer: &mut ThreadObserver) -> AllocationResult<Calibration> {
    let region = allocation_metrics::begin();
    let mut first = None;
    let mut last = None;
    let mut probe_error = None;
    for _ in 0..CALIBRATION_POLLS {
        match observer.read_threads() {
            Ok(count) => {
                first.get_or_insert(count);
                last = Some(count);
                if count != 1 {
                    probe_error = Some(ThreadProbeError::Malformed);
                    break;
                }
                thread::sleep(THREAD_POLL_SLEEP);
            },
            Err(error) => {
                probe_error = Some(error);
                break;
            },
        }
    }
    let sample = region.finish();
    if let Some(error) = probe_error {
        return Err(format!("allocator observer calibration failed: {error}").into());
    }
    if first.is_none() || last.is_none() {
        return Err("allocator observer calibration performed no polls".into());
    }
    let sample = measured_sample(sample, "calibration")?;
    let live_before =
        required_sample_field(&sample, sample.live_bytes_before, "calibration live-before")?;
    let live_after =
        required_sample_field(&sample, sample.live_bytes_after, "calibration live-after")?;
    let region_peak = required_sample_field(
        &sample,
        sample.region_peak_live_bytes,
        "calibration region peak",
    )?;
    let peak_before = required_sample_field(
        &sample,
        sample.peak_live_bytes_before,
        "calibration process peak before",
    )?;
    let peak_after = required_sample_field(
        &sample,
        sample.peak_live_bytes_after,
        "calibration process peak after",
    )?;
    if live_before != live_after {
        return Err("allocator observer calibration changed live bytes during idle polling".into());
    }
    if region_peak != live_before {
        return Err(
            "allocator observer calibration region peak changed during idle polling".into(),
        );
    }
    if peak_before != peak_after {
        return Err(
            "allocator observer calibration process peak changed during idle polling".into(),
        );
    }
    let no_callbacks = sample.allocation_calls == Some(0)
        && sample.deallocation_calls == Some(0)
        && sample.reallocation_calls == Some(0)
        && sample.failed_allocation_calls == Some(0)
        && sample.allocated_bytes == Some(0)
        && sample.deallocated_bytes == Some(0);
    if !no_callbacks {
        return Err("allocator observer calibration saw a callback during idle polls".into());
    }
    Ok(Calibration {
        status: "pass",
        polls: CALIBRATION_POLLS,
        thread_count_first: first.unwrap_or(0),
        thread_count_last: last.unwrap_or(0),
        allocation_calls: sample.allocation_calls.unwrap_or(0),
        deallocation_calls: sample.deallocation_calls.unwrap_or(0),
        reallocation_calls: sample.reallocation_calls.unwrap_or(0),
        failed_allocation_calls: sample.failed_allocation_calls.unwrap_or(0),
        no_allocator_callbacks: no_callbacks,
        poll_sleep_milliseconds: THREAD_POLL_SLEEP.as_millis() as u64,
        thread_scope: THREAD_SCOPE,
    })
}

fn binary_identity() -> AllocationResult<BinaryIdentity> {
    let path = std::env::current_exe()?;
    let bytes = fs::read(&path)?;
    Ok(BinaryIdentity {
        path: path.display().to_string(),
        bytes: u64::try_from(bytes.len())?,
        sha256: sha256_hex(&bytes),
    })
}

fn verify_before_cleanup(
    plan: &CasePlan<'_>,
    sink: &BoundedSink,
    budget: &Budget,
    changed: bool,
    changed_cells: usize,
    output_bytes_before_publication: u64,
    cpu_tasks_before_publication: u64,
) -> AllocationResult<Verification> {
    if !changed || changed_cells != plan.updates.len() {
        return Err(format!(
            "allocation diagnostic commit changed={changed} cells={changed_cells}, expected {}",
            plan.updates.len()
        )
        .into());
    }
    let output_bytes_used = budget
        .used(Resource::OutputBytes)
        .checked_sub(output_bytes_before_publication)
        .ok_or("allocation diagnostic OutputBytes counter moved backwards")?;
    let cpu_tasks_used = budget
        .used(Resource::CpuTasks)
        .checked_sub(cpu_tasks_before_publication)
        .ok_or("allocation diagnostic CpuTasks counter moved backwards")?;
    let resources_after_core = require_clean_resources(budget, "allocation diagnostic")?;
    if output_bytes_used != u64::try_from(sink.bytes.len())? {
        return Err(format!(
            "allocation diagnostic charged {output_bytes_used} OutputBytes for {} sink bytes",
            sink.bytes.len()
        )
        .into());
    }
    if cpu_tasks_used != plan.expected_cpu_tasks {
        return Err(format!(
            "allocation diagnostic charged {cpu_tasks_used} CpuTasks, expected {}",
            plan.expected_cpu_tasks
        )
        .into());
    }
    if sink.bytes != plan.expected_output {
        return Err("allocation diagnostic publication bytes differ from the frozen output".into());
    }
    if sink.write_calls == 0 || sink.largest_write > super::MAX_SINK_WRITE_BYTES {
        return Err("allocation diagnostic sink boundary was not exercised as bounded".into());
    }
    let output_sha256 = sha256_hex(&sink.bytes);
    if output_sha256 != plan.output_hash {
        return Err(format!(
            "allocation diagnostic output hash {output_sha256} differs from {}",
            plan.output_hash
        )
        .into());
    }
    verify_xlsx_cell_crud_output(plan.corpus, &sink.bytes, &plan.updates)?;
    let semantic_sha256 = xlsx_cell_values_output_semantic_hash(&sink.bytes, plan.spec)?;
    if semantic_sha256 != plan.semantic_hash {
        return Err("allocation diagnostic semantic output hash changed".into());
    }
    let mut package = OpcPackage::from_bytes(&sink.bytes, litchi_opc::SourceLimits::default())?;
    crate::xlsx_source_editors::replay_patch(&mut package, &plan.inverse_patch.inverse())?;
    let restored = PackageWriter::to_bytes(&package)?;
    if restored != plan.corpus.archive {
        return Err("allocation diagnostic inverse did not restore exact source ZIP".into());
    }
    if sha256_hex(&plan.corpus.archive) != plan.source_hash {
        return Err("allocation diagnostic source hash changed".into());
    }
    let output_bytes = sink.bytes.len();
    let output_sha256 = sha256_hex(&sink.bytes);
    Ok(Verification {
        output_bytes,
        output_bytes_used,
        cpu_tasks_used,
        sink_write_calls: sink.write_calls,
        sink_largest_write: sink.largest_write,
        resources_after_core,
        output_sha256,
        semantic_sha256,
        fixture_inverse_restored_exact_source: true,
    })
}

fn run_iteration(
    plan: &CasePlan<'_>,
    observer: &mut ThreadObserver,
    sample: usize,
) -> AllocationResult<AllocationRow> {
    let pre_sample_threads = observer
        .drain_to_one()
        .map_err(|error| format!("pre-sample thread drain failed: {error}"))?;
    if !pre_sample_threads.drained {
        return Err("pre-sample thread drain did not reach Threads==1".into());
    }

    let setup_region = allocation_metrics::begin();
    let (budget, cancellation_source, context) = managed_context(
        plan.payload_budget,
        plan.output_budget,
        plan.workers,
        plan.expected_cpu_tasks,
    )?;
    drop(cancellation_source);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(plan.corpus.archive.clone()));
    let mut sink = BoundedSink::new(usize::try_from(plan.output_budget)?)?;
    let setup = measured_sample(setup_region.finish(), "setup")?;
    let limits = super::super::context_limits(&budget, plan.workers);

    let mut core_region = allocation_metrics::begin();
    let editor = SourceBackedEditor::from_read_at_with_options(
        source,
        OpenOptions::new()
            .source_limits(SourceLimits::default())
            .cache_limits(plan.cache_limits)
            .execution_context(context),
    )?;
    let ingress = core_region.split();
    let (commit, changed, changed_cells) = stage_timed_commit(&editor, &plan.updates)?;
    let edit_commit = core_region.split();
    let cpu_tasks_before_publication = budget.used(Resource::CpuTasks);
    let output_bytes_before_publication = budget.used(Resource::OutputBytes);
    let published = editor.publish_multi_commit_to_stream(&mut sink, &commit)?;
    drop(published);
    let publication = core_region.split();
    drop(commit);
    let commit_drop = core_region.split();
    let post_core_threads = match observer.drain_to_one() {
        Ok(drain) if drain.drained => drain,
        Ok(_) => {
            drop(core_region);
            return Err("post-core thread drain did not reach Threads==1".into());
        },
        Err(_) => {
            drop(core_region);
            return Err("post-core thread drain probe failed".into());
        },
    };
    let (worker_drain, core_combined_with_drain) = core_region.finish_split();
    let ingress = measured_sample(ingress, "ingress")?;
    let edit_commit = measured_sample(edit_commit, "edit_commit")?;
    let publication = measured_sample(publication, "publication")?;
    let commit_drop = measured_sample(commit_drop, "commit_drop")?;
    let worker_drain = measured_sample(worker_drain, "worker_drain")?;
    let core_combined_with_drain =
        measured_sample(core_combined_with_drain, "core_combined_with_drain")?;
    let core_combined = verify_core_conservation(
        &ingress,
        &edit_commit,
        &publication,
        &commit_drop,
        &worker_drain,
        &core_combined_with_drain,
    )?;
    let max_setup_core_peak_above_setup_live_before =
        derive_max_setup_core_peak_above_setup_live_before(&setup, &core_combined)?;
    let max_setup_drained_core_peak_above_setup_live_before =
        derive_max_setup_core_peak_above_setup_live_before(&setup, &core_combined_with_drain)?;
    let verification = verify_before_cleanup(
        plan,
        &sink,
        &budget,
        changed,
        changed_cells,
        output_bytes_before_publication,
        cpu_tasks_before_publication,
    )?;

    let postdrop_region = allocation_metrics::begin();
    drop(budget);
    drop(sink);
    let postdrop = measured_sample(postdrop_region.finish(), "postdrop")?;
    let source_sha256 = plan.source_hash.clone();
    let output_sha256 = verification.output_sha256.clone();
    let semantic_sha256 = verification.semantic_sha256.clone();
    Ok(AllocationRow {
        sample,
        setup,
        ingress,
        edit_commit,
        publication,
        commit_drop,
        worker_drain,
        core_combined,
        core_combined_with_drain,
        postdrop,
        max_setup_core_peak_above_setup_live_before,
        max_setup_drained_core_peak_above_setup_live_before,
        context_limits: limits,
        workers_requested: plan.workers,
        expected_worker_permits: plan.admission_expected_worker_permits,
        admitted_worker_permits: plan.admission_observed_worker_permits,
        expected_cpu_tasks: plan.expected_cpu_tasks,
        cpu_tasks_used: verification.cpu_tasks_used,
        output_bytes: verification.output_bytes,
        output_bytes_used: verification.output_bytes_used,
        sink_write_calls: verification.sink_write_calls,
        sink_largest_write: verification.sink_largest_write,
        resources_after_core: verification.resources_after_core,
        source_sha256,
        output_sha256,
        semantic_sha256,
        fixture_inverse_restored_exact_source: verification.fixture_inverse_restored_exact_source,
        pre_sample_threads,
        post_core_threads,
    })
}

fn run_case(
    plan: &CasePlan<'_>,
    observer: &mut ThreadObserver,
) -> AllocationResult<AllocationCaseReport> {
    for _ in 0..plan.warmup {
        let _ = run_iteration(plan, observer, 0)?;
    }
    let mut rows = Vec::with_capacity(plan.samples);
    for sample in 0..plan.samples {
        rows.push(run_iteration(plan, observer, sample)?);
    }
    let context_limits = rows
        .first()
        .map(|row| row.context_limits.clone())
        .ok_or("allocation diagnostic recorded no measured rows")?;
    Ok(AllocationCaseReport {
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
        source_sha256: plan.source_hash.clone(),
        output_bytes: plan.expected_output.len(),
        output_sha256: plan.output_hash.clone(),
        semantic_sha256: plan.semantic_hash.clone(),
        expected_cpu_tasks: plan.expected_cpu_tasks,
        admission_expected_worker_permits: plan.admission_expected_worker_permits,
        admission_observed_worker_permits: plan.admission_observed_worker_permits,
        admission_cpu_tasks_used: plan.admission_cpu_tasks_used,
        preflight_actual_commit_inverse_verified: plan.preflight_actual_commit_inverse_verified,
        context_limits,
        sink_capacity_bytes: usize::try_from(plan.output_budget)?,
        sink_max_write_bytes: super::MAX_SINK_WRITE_BYTES,
        warmup_iterations: plan.warmup,
        measured_samples: plan.samples,
        rows,
    })
}

fn create_output_directory(directory: &Path) -> AllocationResult<()> {
    let parent = directory
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    fs::create_dir_all(parent)?;
    fs::create_dir(directory).map_err(|error| {
        if error.kind() == std::io::ErrorKind::AlreadyExists {
            format!(
                "allocator diagnostic output directory already exists: {}",
                directory.display()
            )
        } else {
            format!(
                "could not create allocator diagnostic output directory {}: {error}",
                directory.display()
            )
        }
    })?;
    Ok(())
}

fn write_atomic_json(path: &Path, bytes: &[u8]) -> AllocationResult<()> {
    let temporary = path.with_extension("json.tmp");
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&temporary)?;
    file.write_all(bytes)?;
    file.sync_all()?;
    drop(file);
    fs::rename(temporary, path)?;
    Ok(())
}

fn run(arguments: Arguments, directory: &Path) -> AllocationResult<()> {
    verify_instrumentation()?;
    let binary_identity = binary_identity()?;
    let mut observer = ThreadObserver::open()?;
    let calibration = calibrate(&mut observer)?;
    let corpus = build_xlsx_cell_crud_corpus(arguments.shape)?;
    let spec = xlsx_spec(&corpus)?;
    let plan = build_case_plan(&corpus, spec, arguments)?;
    let case = run_case(&plan, &mut observer)?;
    let report = AllocationReport {
        schema: SCHEMA,
        status: "completed",
        binary_identity,
        allocator_identity: allocation_metrics::allocator_identity(),
        instrumentation_identity: allocation_metrics::instrumentation_identity(),
        counter_revision: allocation_metrics::counter_revision()
            .ok_or("counter revision disappeared")?,
        allocation_scope: ANALYZER_SCOPE,
        thread_observation_scope: THREAD_SCOPE,
        timing_scope: "none; this report contains allocator regions and diagnostic quiescence observations only",
        performance_claim: "none; allocator counters do not claim latency, throughput, RSS, OS-thread width, worker/task overlap, speedup, or adoption; procfs Threads values are quiescence observations only",
        peak_semantics: PEAK_SEMANTICS,
        calibration,
        case,
    };
    let bytes = serde_json::to_vec_pretty(&report)?;
    write_atomic_json(&directory.join(REPORT_FILE), &bytes)
}

/// Run one explicit allocator-only managed-compression case.
pub(crate) fn measure<I>(arguments: I) -> AllocationResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let (arguments, directory) = super::parse_args(arguments)?;
    create_output_directory(&directory)?;
    let result = run(arguments, &directory);
    if let Err(error) = result {
        let failure = FailureReport {
            schema: SCHEMA,
            status: "failed",
            shape: arguments.shape.name().to_owned(),
            workload: arguments.workload.name().to_owned(),
            workers: arguments.workers,
            warmup_iterations: arguments.warmup,
            measured_samples: arguments.samples,
            error: error.to_string(),
            scope: "No allocation report is admitted after any calibration, oracle, region, or quiescence failure; this failure file is retained in the exclusive output directory.",
        };
        let bytes = serde_json::to_vec_pretty(&failure)?;
        write_atomic_json(&directory.join(FAILURE_FILE), &bytes)?;
        return Err(error);
    }
    Ok(())
}
