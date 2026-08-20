//! Fail-closed metrics for explicit benchmark parallelism.
//!
//! The baseline harness already records two bounded execution boundaries:
//! scaling cases retain the worker width and deterministic logical task/byte
//! totals, while the OPC cache contention case creates a private worker team
//! and records its width.  This module gives those facts a small, explicit
//! envelope without turning them into process-wide observations.
//!
//! In particular, this module deliberately does *not* read `/proc`, inspect a
//! process thread list, infer a worker count from CPU utilization, or convert
//! waiter counts into lock time.  The process may contain unrelated runtime,
//! profiler, or harness threads.  A process-wide thread count and lock wait
//! therefore remain unavailable until an exact instrumented boundary is added
//! to the producer schema.

use std::{error::Error, fmt};

use serde::Serialize;
use serde_json::Value;

use crate::{CaseResult, Configuration, SourceSummary};

const SCHEMA_VERSION: u32 = 1;
const SCOPE: &str = "explicit_local_execution_only";
const CLAIM: &str = "descriptive";
const CONFIGURED_WORKER_SCOPE: &str = "configuration.execution_workers";
const CONFIGURED_CASE_WORKER_SCOPE: &str = "result.execution.worker_count";
const CONFIGURED_OPC_WORKER_SCOPE: &str = "result.source.opc_cache.worker_count";
const OBSERVED_WORKER_SCOPE: &str =
    "result.source.opc_cache.worker_count_with_one_created_local_worker_team";
const TASK_SCOPE: &str = "result.execution.logical_tasks";
const CHUNK_SCOPE: &str = "result.execution.deterministic_chunk_count";
const SIMULATION_CHUNK_SCOPE: &str = "result.source.simulation.physical_request_count";
const CFB_SELECTIVE_CHUNK_SCOPE: &str =
    "result.source.cfb_selective.simulation.read.physical_request_count";
const CFB_OPEN_STREAM_CHUNK_SCOPE: &str =
    "result.source.cfb_open_stream.simulation.samples.per_operation.physical_request_count_sum";
const PROCESS_THREAD_SCOPE: &str = "process_thread_count";
const LOCK_WAIT_SCOPE: &str = "lock_wait_ns";

/// Why a metric is present in the report.
#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub(crate) enum MetricStatus {
    /// The producer exposed an exact value at the stated scope.
    Measured,
    /// The operation has no corresponding bounded metric.
    NotApplicable,
    /// The metric could be meaningful, but no sound producer counter exists.
    Unavailable,
}

/// A scalar or small structured value with an explicit measurement scope.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct Metric {
    pub status: MetricStatus,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub value: Option<Value>,
    pub scope: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub reason: Option<&'static str>,
}

impl Metric {
    fn measured(value: Value, scope: &'static str) -> Self {
        Self {
            status: MetricStatus::Measured,
            value: Some(value),
            scope,
            reason: None,
        }
    }

    fn measured_u64(value: u64, scope: &'static str) -> Self {
        Self::measured(Value::from(value), scope)
    }

    fn measured_u64s(values: Vec<u64>, scope: &'static str) -> Self {
        Self::measured(
            Value::Array(values.into_iter().map(Value::from).collect()),
            scope,
        )
    }

    fn not_applicable(scope: &'static str, reason: &'static str) -> Self {
        Self {
            status: MetricStatus::NotApplicable,
            value: None,
            scope,
            reason: Some(reason),
        }
    }

    fn unavailable(scope: &'static str, reason: &'static str) -> Self {
        Self {
            status: MetricStatus::Unavailable,
            value: None,
            scope,
            reason: Some(reason),
        }
    }
}

/// Per-result facts that can be attributed to an explicit local execution
/// boundary.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct CaseMetrics {
    pub case: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub cache_state: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub corpus_sha256: Option<String>,
    pub configured_worker_count: Metric,
    pub observed_local_worker_count: Metric,
    pub deterministic_task_count: Metric,
    pub deterministic_chunk_count: Metric,
    pub lock_wait_ns: Metric,
}

/// Additive report-level parallelism envelope.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct ReportMetrics {
    pub schema_version: u32,
    pub scope: &'static str,
    pub claim: &'static str,
    pub configured_worker_budget: Metric,
    pub observed_process_thread_count: Metric,
    pub cases: Vec<CaseMetrics>,
}

/// Input errors are kept distinct from benchmark operation errors so a
/// malformed or ambiguous metric producer cannot silently become a zero.
#[derive(Debug)]
struct InputError(String);

impl fmt::Display for InputError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for InputError {}

#[derive(Debug)]
struct ChunkVector {
    scope: &'static str,
    values: Vec<u64>,
}

#[derive(Debug)]
enum ChunkEvidence {
    None,
    One(ChunkVector),
    Ambiguous,
}

#[derive(Debug)]
struct CaseInput {
    case: String,
    cache_state: Option<String>,
    corpus_sha256: Option<String>,
    execution_worker_count: Option<usize>,
    logical_tasks: Option<usize>,
    opc_cache_worker_count: Option<usize>,
    persistent_worker_teams_created: Option<usize>,
    chunk_evidence: ChunkEvidence,
    elapsed_sample_count: usize,
    elapsed_sample_order: Vec<usize>,
}

/// Build the envelope from the harness's typed report pieces.
///
/// This intentionally reads the producer structs directly.  The benchmark
/// report is serialized only once, by the existing report writer, so metric
/// collection cannot add a second JSON materialization of every result.
pub(crate) fn collect(
    configuration: &Configuration,
    results: &[CaseResult],
) -> Result<ReportMetrics, Box<dyn Error>> {
    let configured_worker_values = configured_worker_values(&configuration.execution_workers)?;
    let configured_worker_budget =
        Metric::measured_u64s(configured_worker_values.clone(), CONFIGURED_WORKER_SCOPE);
    let cases = results
        .iter()
        .enumerate()
        .map(|(index, result)| {
            let input = case_input(result)?;
            case_metrics(input, index, &configured_worker_values)
        })
        .collect::<Result<Vec<_>, InputError>>()?;
    Ok(ReportMetrics {
        schema_version: SCHEMA_VERSION,
        scope: SCOPE,
        claim: CLAIM,
        configured_worker_budget,
        observed_process_thread_count: Metric::unavailable(
            PROCESS_THREAD_SCOPE,
            "no process-global thread counter is collected; local worker boundaries only",
        ),
        cases,
    })
}

fn case_input(result: &CaseResult) -> Result<CaseInput, InputError> {
    let source = result.source.as_ref();
    let execution = result.execution.as_ref();
    let (opc_cache_worker_count, persistent_worker_teams_created) = source
        .and_then(|source| source.opc_cache.as_ref())
        .map(|opc_cache| {
            (
                Some(opc_cache.worker_count),
                Some(opc_cache.persistent_worker_teams_created),
            )
        })
        .unwrap_or((None, None));
    Ok(CaseInput {
        case: result.case.to_owned(),
        cache_state: result.cache_state.map(str::to_owned),
        corpus_sha256: (!result.corpus.archive_sha256.is_empty())
            .then(|| result.corpus.archive_sha256.clone()),
        execution_worker_count: execution.map(|execution| execution.worker_count),
        logical_tasks: execution.map(|execution| execution.logical_tasks),
        opc_cache_worker_count,
        persistent_worker_teams_created,
        chunk_evidence: source_chunk_evidence(source)?,
        elapsed_sample_count: result.elapsed_ns.samples.len(),
        elapsed_sample_order: result.elapsed_ns.sample_order.clone(),
    })
}

fn source_chunk_evidence(source: Option<&SourceSummary>) -> Result<ChunkEvidence, InputError> {
    let Some(source) = source else {
        return Ok(ChunkEvidence::None);
    };
    let mut candidates = Vec::new();
    if let Some(simulation) = source.simulation.as_ref() {
        candidates.push(ChunkVector {
            scope: SIMULATION_CHUNK_SCOPE,
            values: simulation.physical_request_count.clone(),
        });
    }
    if let Some(cfb_selective) = source.cfb_selective.as_ref() {
        if let Some(simulation) = cfb_selective.simulation.as_ref() {
            candidates.push(ChunkVector {
                scope: CFB_SELECTIVE_CHUNK_SCOPE,
                values: simulation
                    .read
                    .iter()
                    .map(|phase| phase.physical_request_count)
                    .collect(),
            });
        }
    }
    if let Some(cfb_open_stream) = source.cfb_open_stream.as_ref() {
        if let Some(simulation) = cfb_open_stream.simulation.as_ref() {
            candidates.push(ChunkVector {
                scope: CFB_OPEN_STREAM_CHUNK_SCOPE,
                values: simulation
                    .samples
                    .iter()
                    .enumerate()
                    .map(|(sample_index, sample)| {
                        sample.per_operation.iter().try_fold(0_u64, |total, phase| {
                            total
                                .checked_add(phase.physical_request_count)
                                .ok_or_else(|| {
                                    InputError(format!(
                                        "source.cfb_open_stream.simulation.samples[{sample_index}] \
                                         per-operation request count overflows u64"
                                    ))
                                })
                        })
                    })
                    .collect::<Result<Vec<_>, _>>()?,
            });
        }
    }
    Ok(match candidates.len() {
        0 => ChunkEvidence::None,
        1 => match candidates.pop() {
            Some(candidate) => ChunkEvidence::One(candidate),
            None => ChunkEvidence::None,
        },
        _ => ChunkEvidence::Ambiguous,
    })
}

fn configured_worker_budget(workers: &[usize]) -> Result<Metric, InputError> {
    let values = configured_worker_values(workers)?;
    Ok(Metric::measured_u64s(values, CONFIGURED_WORKER_SCOPE))
}

fn configured_worker_values(workers: &[usize]) -> Result<Vec<u64>, InputError> {
    if workers.is_empty() {
        return Err(InputError(
            "configuration.execution_workers must not be empty".to_owned(),
        ));
    }
    let values = workers
        .iter()
        .enumerate()
        .map(|(index, &worker)| positive_usize(worker, &format!("execution_workers[{index}]")))
        .collect::<Result<Vec<_>, _>>()?;
    if !values.windows(2).all(|pair| pair[0] < pair[1]) {
        return Err(InputError(
            "configuration.execution_workers must be sorted and unique".to_owned(),
        ));
    }
    Ok(values)
}

fn case_metrics(
    input: CaseInput,
    index: usize,
    configured_workers: &[u64],
) -> Result<CaseMetrics, InputError> {
    let location = format!("results[{index}]");
    validate_case_worker_counts(&input, configured_workers, &location)?;
    let configured_worker_count = configured_worker_count(&input, &location)?;
    let observed_local_worker_count = observed_local_worker_count(&input, &location)?;
    let deterministic_task_count = match input.logical_tasks {
        Some(tasks) if input.execution_worker_count.is_some() => {
            Metric::measured_u64(usize_to_u64(tasks, "logical task count")?, TASK_SCOPE)
        },
        Some(_) => Metric::unavailable(
            TASK_SCOPE,
            "logical task count has no explicit execution context",
        ),
        None => Metric::not_applicable(
            TASK_SCOPE,
            "result does not use an explicit execution context",
        ),
    };
    let deterministic_chunk_count = deterministic_chunk_count(&input, &location)?;
    Ok(CaseMetrics {
        case: input.case,
        cache_state: input.cache_state,
        corpus_sha256: input.corpus_sha256,
        configured_worker_count,
        observed_local_worker_count,
        deterministic_task_count,
        deterministic_chunk_count,
        lock_wait_ns: Metric::unavailable(
            LOCK_WAIT_SCOPE,
            "no exact instrumented lock boundary is present; waiter counts are not timed",
        ),
    })
}

fn configured_worker_count(input: &CaseInput, location: &str) -> Result<Metric, InputError> {
    let execution = input
        .execution_worker_count
        .map(|worker| positive_usize(worker, &format!("{location}.execution.worker_count")))
        .transpose()?;
    let opc_cache = input
        .opc_cache_worker_count
        .map(|worker| positive_usize(worker, &format!("{location}.source.opc_cache.worker_count")))
        .transpose()?;
    match (execution, opc_cache) {
        (Some(execution), Some(opc_cache)) if execution != opc_cache => Ok(Metric::unavailable(
            CONFIGURED_CASE_WORKER_SCOPE,
            "multiple explicit execution domains report different worker widths",
        )),
        (Some(value), _) => Ok(Metric::measured_u64(value, CONFIGURED_CASE_WORKER_SCOPE)),
        (_, Some(value)) => Ok(Metric::measured_u64(value, CONFIGURED_OPC_WORKER_SCOPE)),
        (None, None) => Ok(Metric::not_applicable(
            CONFIGURED_CASE_WORKER_SCOPE,
            "result has no explicit worker-budget field",
        )),
    }
}

fn validate_case_worker_counts(
    input: &CaseInput,
    configured_workers: &[u64],
    location: &str,
) -> Result<(), InputError> {
    for (field, worker) in [
        ("execution.worker_count", input.execution_worker_count),
        (
            "source.opc_cache.worker_count",
            input.opc_cache_worker_count,
        ),
    ] {
        let Some(worker) = worker else {
            continue;
        };
        let worker = positive_usize(worker, &format!("{location}.{field}"))?;
        if !configured_workers.contains(&worker) {
            return Err(InputError(format!(
                "{location}.{field}={worker} is absent from configuration.execution_workers"
            )));
        }
    }
    Ok(())
}

fn observed_local_worker_count(input: &CaseInput, location: &str) -> Result<Metric, InputError> {
    let Some(worker_count) = input.opc_cache_worker_count else {
        return Ok(Metric::not_applicable(
            OBSERVED_WORKER_SCOPE,
            "result does not create an explicit local worker team",
        ));
    };
    let Some(teams) = input.persistent_worker_teams_created else {
        return Ok(Metric::not_applicable(
            OBSERVED_WORKER_SCOPE,
            "local worker-team creation count is not exposed",
        ));
    };
    let worker_count = positive_usize(
        worker_count,
        &format!("{location}.source.opc_cache.worker_count"),
    )?;
    match teams {
        1 => Ok(Metric::measured_u64(worker_count, OBSERVED_WORKER_SCOPE)),
        0 => Ok(Metric::not_applicable(
            OBSERVED_WORKER_SCOPE,
            "serial result created no local worker team",
        )),
        _ => Ok(Metric::unavailable(
            OBSERVED_WORKER_SCOPE,
            "multiple local worker teams were created; one bounded team is required",
        )),
    }
}

fn deterministic_chunk_count(input: &CaseInput, location: &str) -> Result<Metric, InputError> {
    match &input.chunk_evidence {
        ChunkEvidence::None if input.execution_worker_count.is_some() => Ok(Metric::unavailable(
            CHUNK_SCOPE,
            "no deterministic chunk counter is exposed; byte totals are not used as a proxy",
        )),
        ChunkEvidence::None => Ok(Metric::not_applicable(
            CHUNK_SCOPE,
            "result has neither an explicit execution context nor range simulation",
        )),
        ChunkEvidence::Ambiguous => Ok(Metric::unavailable(
            CHUNK_SCOPE,
            "multiple chunk counters are present without one unambiguous source",
        )),
        ChunkEvidence::One(vector) if vector.values.is_empty() => Ok(Metric::unavailable(
            vector.scope,
            "chunk producer exposed no retained samples",
        )),
        ChunkEvidence::One(vector) => {
            if input.elapsed_sample_count != vector.values.len() {
                return Err(InputError(format!(
                    "{location}.{} has {} values but elapsed_ns.samples has {}",
                    vector.scope,
                    vector.values.len(),
                    input.elapsed_sample_count
                )));
            }
            let values =
                reorder_to_elapsed_samples(&vector.values, &input.elapsed_sample_order, location)?;
            Ok(Metric::measured_u64s(values, vector.scope))
        },
    }
}

fn reorder_to_elapsed_samples(
    values: &[u64],
    sample_order: &[usize],
    location: &str,
) -> Result<Vec<u64>, InputError> {
    if values.len() != sample_order.len() {
        return Err(InputError(format!(
            "{location} sample-order length {} does not match chunk vector length {}",
            sample_order.len(),
            values.len()
        )));
    }
    let mut seen = vec![false; values.len()];
    let mut reordered = Vec::with_capacity(values.len());
    for (sorted_index, &original_index) in sample_order.iter().enumerate() {
        if original_index >= values.len() {
            return Err(InputError(format!(
                "{location} sample_order[{sorted_index}]={original_index} is outside \
                 the retained sample vector"
            )));
        }
        if seen[original_index] {
            return Err(InputError(format!(
                "{location} sample_order repeats original sample {original_index}"
            )));
        }
        seen[original_index] = true;
        reordered.push(values[original_index]);
    }
    if seen.iter().any(|observed| !observed) {
        return Err(InputError(format!(
            "{location} sample_order is not a complete retained-sample permutation"
        )));
    }
    Ok(reordered)
}

fn positive_usize(value: usize, location: &str) -> Result<u64, InputError> {
    let value = usize_to_u64(value, location)?;
    if value == 0 {
        return Err(InputError(format!("{location} must be positive")));
    }
    Ok(value)
}

fn usize_to_u64(value: usize, location: &str) -> Result<u64, InputError> {
    u64::try_from(value).map_err(|_| InputError(format!("{location} does not fit in u64")))
}

#[cfg(test)]
mod tests {
    #![allow(clippy::expect_used, reason = "test assertions panic by design")]

    use super::{
        CaseInput, ChunkEvidence, ChunkVector, MetricStatus, case_metrics, collect,
        configured_worker_budget, source_chunk_evidence,
    };

    fn configuration() -> super::super::Configuration {
        super::super::Configuration {
            samples_per_case: 2,
            warmup_iterations_per_case: 0,
            filesystem_cache_states: vec!["warm"],
            filesystem_fresh_child_per_sample: true,
            filesystem_process_isolated: true,
            filesystem_root_selected: false,
            cases: vec!["test"],
            corpus_shapes: Vec::new(),
            payload_kinds: Vec::new(),
            writer_shapes: Vec::new(),
            xlsx_shapes: Vec::new(),
            xlsx_cell_crud_shapes: Vec::new(),
            xlsx_row_visibility_shapes: Vec::new(),
            semantic_shapes: Vec::new(),
            rtf_variants: Vec::new(),
            range_simulation: super::super::RangeSimulationConfig::default(),
            execution_workers: vec![1, 2, 4],
        }
    }

    fn corpus() -> super::super::CorpusManifest {
        super::super::CorpusManifest {
            name: "test".to_owned(),
            generator: "test",
            package_format: "opc",
            shape: "tiny",
            payload_kind: "compressible",
            compression: "stored",
            entry_count: 1,
            archive_member_count: 1,
            entry_bytes: 1,
            uncompressed_payload_bytes: 1,
            archive_bytes: 1,
            archive_sha256: "archive".to_owned(),
            target_entry: "part".to_owned(),
            target_payload_bytes: 1,
            target_payload_sha256: "payload".to_owned(),
            rtf_variant: None,
            xlsx: None,
        }
    }

    fn typed_result() -> super::super::CaseResult {
        super::super::CaseResult {
            case: "test",
            cache_state: None,
            corpus: corpus(),
            elapsed_ns: super::super::statistics(vec![20, 10]),
            sink: None,
            source: Some(super::super::SourceSummary {
                simulation: Some(super::super::RangeSimulationSummary {
                    physical_request_count: vec![11, 22],
                    ..Default::default()
                }),
                ..Default::default()
            }),
            execution: Some(super::super::ExecutionSummary {
                worker_count: 2,
                logical_tasks: 3,
                logical_bytes: 30,
            }),
            output_sha256: None,
            operation_metrics: None,
        }
    }

    fn input() -> CaseInput {
        CaseInput {
            case: "test".to_owned(),
            cache_state: None,
            corpus_sha256: Some("archive".to_owned()),
            execution_worker_count: None,
            logical_tasks: None,
            opc_cache_worker_count: None,
            persistent_worker_teams_created: None,
            chunk_evidence: ChunkEvidence::None,
            elapsed_sample_count: 0,
            elapsed_sample_order: Vec::new(),
        }
    }

    #[test]
    fn collect_validates_typed_shape_and_reorders_chunk_evidence() {
        let report = collect(&configuration(), &[typed_result()]).expect("valid report");
        assert_eq!(report.claim, "descriptive");
        assert_eq!(
            report.configured_worker_budget.value,
            Some(serde_json::json!([1, 2, 4]))
        );
        let case = &report.cases[0];
        assert_eq!(
            case.configured_worker_count.value,
            Some(serde_json::json!(2))
        );
        assert_eq!(
            case.deterministic_task_count.value,
            Some(serde_json::json!(3))
        );
        assert_eq!(
            case.deterministic_chunk_count.value,
            Some(serde_json::json!([22, 11]))
        );
        assert_eq!(
            case.observed_local_worker_count.status,
            MetricStatus::NotApplicable
        );
        assert_eq!(
            report.observed_process_thread_count.status,
            MetricStatus::Unavailable
        );
    }

    #[test]
    fn records_explicit_worker_budget_and_logical_tasks_only() {
        let budget = configured_worker_budget(&[1, 2, 4]).expect("valid worker budget");
        assert_eq!(budget.status, MetricStatus::Measured);
        let mut input = input();
        input.execution_worker_count = Some(2);
        input.logical_tasks = Some(6);
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("valid metrics");
        assert_eq!(case.configured_worker_count.status, MetricStatus::Measured);
        assert_eq!(case.deterministic_task_count.status, MetricStatus::Measured);
        assert_eq!(
            case.deterministic_chunk_count.status,
            MetricStatus::Unavailable
        );
        assert_eq!(
            case.observed_local_worker_count.status,
            MetricStatus::NotApplicable
        );
        assert_eq!(case.lock_wait_ns.status, MetricStatus::Unavailable);
        assert_eq!(
            case.deterministic_task_count.value,
            Some(serde_json::json!(6))
        );
    }

    #[test]
    fn records_one_explicit_opc_worker_team_as_local_observation() {
        let mut input = input();
        input.opc_cache_worker_count = Some(4);
        input.persistent_worker_teams_created = Some(1);
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("valid metrics");
        assert_eq!(case.configured_worker_count.status, MetricStatus::Measured);
        assert_eq!(
            case.observed_local_worker_count.status,
            MetricStatus::Measured
        );
        assert_eq!(
            case.observed_local_worker_count.value,
            Some(serde_json::json!(4))
        );
        assert_eq!(case.lock_wait_ns.status, MetricStatus::Unavailable);
    }

    #[test]
    fn treats_serial_zero_team_as_not_applicable() {
        let mut input = input();
        input.opc_cache_worker_count = Some(1);
        input.persistent_worker_teams_created = Some(0);
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("serial metrics");
        assert_eq!(
            case.observed_local_worker_count.status,
            MetricStatus::NotApplicable
        );
        assert_eq!(case.observed_local_worker_count.value, None);
    }

    #[test]
    fn treats_multiple_worker_teams_as_ambiguous() {
        let mut input = input();
        input.opc_cache_worker_count = Some(2);
        input.persistent_worker_teams_created = Some(2);
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("ambiguous metrics");
        assert_eq!(
            case.observed_local_worker_count.status,
            MetricStatus::Unavailable
        );
        assert_eq!(case.observed_local_worker_count.value, None);
    }

    #[test]
    fn rejects_result_worker_width_outside_configured_budget() {
        let mut input = input();
        input.execution_worker_count = Some(8);
        input.logical_tasks = Some(1);
        assert!(case_metrics(input, 0, &[1, 2, 4]).is_err());

        let mut input = input();
        input.opc_cache_worker_count = Some(8);
        input.persistent_worker_teams_created = Some(1);
        assert!(case_metrics(input, 0, &[1, 2, 4]).is_err());
    }

    #[test]
    fn never_converts_waiter_counts_or_bytes_into_lock_or_chunk_metrics() {
        let mut input = input();
        input.execution_worker_count = Some(1);
        input.logical_tasks = Some(8);
        input.opc_cache_worker_count = Some(1);
        input.persistent_worker_teams_created = Some(1);
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("valid metrics");
        assert_eq!(case.lock_wait_ns.value, None);
        assert_eq!(case.deterministic_chunk_count.value, None);
    }

    #[test]
    fn records_range_simulator_request_chunks_per_retained_sample() {
        let mut input = input();
        input.elapsed_sample_count = 2;
        input.elapsed_sample_order = vec![0, 1];
        input.chunk_evidence = ChunkEvidence::One(ChunkVector {
            scope: "result.source.simulation.physical_request_count",
            values: vec![3, 4],
        });
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("valid metrics");
        assert_eq!(
            case.deterministic_chunk_count.status,
            MetricStatus::Measured
        );
        assert_eq!(
            case.deterministic_chunk_count.value,
            Some(serde_json::json!([3, 4]))
        );
    }

    #[test]
    fn extracts_typed_range_simulation_counter() {
        let source = super::super::SourceSummary {
            simulation: Some(super::super::RangeSimulationSummary {
                physical_request_count: vec![5, 6],
                ..Default::default()
            }),
            ..Default::default()
        };
        let ChunkEvidence::One(vector) =
            source_chunk_evidence(Some(&source)).expect("typed range simulation should parse")
        else {
            panic!("typed range simulation should produce one chunk vector");
        };
        assert_eq!(
            vector.scope,
            "result.source.simulation.physical_request_count"
        );
        assert_eq!(vector.values, vec![5, 6]);
    }

    #[test]
    fn rejects_chunk_vectors_that_cannot_align_to_retained_samples() {
        let mut input = input();
        input.elapsed_sample_count = 2;
        input.elapsed_sample_order = vec![0, 1];
        input.chunk_evidence = ChunkEvidence::One(ChunkVector {
            scope: "result.source.simulation.physical_request_count",
            values: vec![3],
        });
        assert!(case_metrics(input, 0, &[1, 2, 4]).is_err());
    }

    #[test]
    fn leaves_ambiguous_chunk_sources_unavailable() {
        let mut input = input();
        input.elapsed_sample_count = 1;
        input.elapsed_sample_order = vec![0];
        input.chunk_evidence = ChunkEvidence::Ambiguous;
        let case = case_metrics(input, 0, &[1, 2, 4]).expect("ambiguous evidence is reportable");
        assert_eq!(
            case.deterministic_chunk_count.status,
            MetricStatus::Unavailable
        );
        assert_eq!(case.deterministic_chunk_count.value, None);
    }

    #[test]
    fn rejects_unsorted_or_duplicate_worker_budget() {
        for workers in [vec![], vec![2, 1], vec![1, 1], vec![0, 1]] {
            assert!(configured_worker_budget(&workers).is_err());
        }
    }
}
