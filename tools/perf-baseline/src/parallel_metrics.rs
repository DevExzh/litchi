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
const CONFIGURED_WORKER_SCOPE: &str = "configuration.execution_workers";
const CONFIGURED_CASE_WORKER_SCOPE: &str = "result.execution.worker_count";
const CONFIGURED_OPC_WORKER_SCOPE: &str = "result.source.opc_cache.worker_count";
const OBSERVED_WORKER_SCOPE: &str =
    "result.source.opc_cache.worker_count_with_one_created_local_worker_team";
const TASK_SCOPE: &str = "result.execution.logical_tasks";
const CHUNK_SCOPE: &str = "result.execution.deterministic_chunk_count";
const SIMULATION_CHUNK_SCOPE: &str = "result.source.simulation.physical_request_count";
const CFB_SELECTIVE_CHUNK_SCOPE: &str =
    "result.source.cfb_selective.simulation.total.physical_request_count";
const CFB_OPEN_STREAM_CHUNK_SCOPE: &str =
    "result.source.cfb_open_stream.simulation.samples.aggregate.physical_request_count";
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
    let configured_worker_budget = configured_worker_budget(&configuration.execution_workers)?;
    let cases = results
        .iter()
        .enumerate()
        .map(|(index, result)| case_metrics(case_input(result), index))
        .collect::<Result<Vec<_>, InputError>>()?;
    Ok(ReportMetrics {
        schema_version: SCHEMA_VERSION,
        scope: SCOPE,
        configured_worker_budget,
        observed_process_thread_count: Metric::unavailable(
            PROCESS_THREAD_SCOPE,
            "no process-global thread counter is collected; local worker boundaries only",
        ),
        cases,
    })
}

fn case_input(result: &CaseResult) -> CaseInput {
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
    CaseInput {
        case: result.case.to_owned(),
        cache_state: result.cache_state.map(str::to_owned),
        corpus_sha256: (!result.corpus.archive_sha256.is_empty())
            .then(|| result.corpus.archive_sha256.clone()),
        execution_worker_count: execution.map(|execution| execution.worker_count),
        logical_tasks: execution.map(|execution| execution.logical_tasks),
        opc_cache_worker_count,
        persistent_worker_teams_created,
        chunk_evidence: source_chunk_evidence(source),
        elapsed_sample_count: result.elapsed_ns.samples.len(),
    }
}

fn source_chunk_evidence(source: Option<&SourceSummary>) -> ChunkEvidence {
    let Some(source) = source else {
        return ChunkEvidence::None;
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
                    .total
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
                    .map(|sample| sample.aggregate.physical_request_count)
                    .collect(),
            });
        }
    }
    match candidates.len() {
        0 => ChunkEvidence::None,
        1 => match candidates.pop() {
            Some(candidate) => ChunkEvidence::One(candidate),
            None => ChunkEvidence::None,
        },
        _ => ChunkEvidence::Ambiguous,
    }
}

fn configured_worker_budget(workers: &[usize]) -> Result<Metric, InputError> {
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
    Ok(Metric::measured_u64s(values, CONFIGURED_WORKER_SCOPE))
}

fn case_metrics(input: CaseInput, index: usize) -> Result<CaseMetrics, InputError> {
    let location = format!("results[{index}]");
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

fn observed_local_worker_count(input: &CaseInput, location: &str) -> Result<Metric, InputError> {
    let Some(worker_count) = input.opc_cache_worker_count else {
        return Ok(Metric::unavailable(
            OBSERVED_WORKER_SCOPE,
            "configured worker width is not an observed process thread count",
        ));
    };
    let Some(teams) = input.persistent_worker_teams_created else {
        return Ok(Metric::unavailable(
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
        0 => Ok(Metric::unavailable(
            OBSERVED_WORKER_SCOPE,
            "no local worker team was created",
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
            Ok(Metric::measured_u64s(vector.values.clone(), vector.scope))
        },
    }
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
        CaseInput, ChunkEvidence, ChunkVector, MetricStatus, case_metrics,
        configured_worker_budget, source_chunk_evidence,
    };

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
        }
    }

    #[test]
    fn records_explicit_worker_budget_and_logical_tasks_only() {
        let budget = configured_worker_budget(&[1, 2, 4]).expect("valid worker budget");
        assert_eq!(budget.status, MetricStatus::Measured);
        let mut input = input();
        input.execution_worker_count = Some(2);
        input.logical_tasks = Some(6);
        let case = case_metrics(input, 0).expect("valid metrics");
        assert_eq!(case.configured_worker_count.status, MetricStatus::Measured);
        assert_eq!(case.deterministic_task_count.status, MetricStatus::Measured);
        assert_eq!(
            case.deterministic_chunk_count.status,
            MetricStatus::Unavailable
        );
        assert_eq!(
            case.observed_local_worker_count.status,
            MetricStatus::Unavailable
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
        let case = case_metrics(input, 0).expect("valid metrics");
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
    fn never_converts_waiter_counts_or_bytes_into_lock_or_chunk_metrics() {
        let mut input = input();
        input.execution_worker_count = Some(1);
        input.logical_tasks = Some(8);
        input.opc_cache_worker_count = Some(1);
        input.persistent_worker_teams_created = Some(1);
        let case = case_metrics(input, 0).expect("valid metrics");
        assert_eq!(case.lock_wait_ns.value, None);
        assert_eq!(case.deterministic_chunk_count.value, None);
    }

    #[test]
    fn records_range_simulator_request_chunks_per_retained_sample() {
        let mut input = input();
        input.elapsed_sample_count = 2;
        input.chunk_evidence = ChunkEvidence::One(ChunkVector {
            scope: "result.source.simulation.physical_request_count",
            values: vec![3, 4],
        });
        let case = case_metrics(input, 0).expect("valid metrics");
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
        let ChunkEvidence::One(vector) = source_chunk_evidence(Some(&source)) else {
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
        input.chunk_evidence = ChunkEvidence::One(ChunkVector {
            scope: "result.source.simulation.physical_request_count",
            values: vec![3],
        });
        assert!(case_metrics(input, 0).is_err());
    }

    #[test]
    fn leaves_ambiguous_chunk_sources_unavailable() {
        let mut input = input();
        input.elapsed_sample_count = 1;
        input.chunk_evidence = ChunkEvidence::Ambiguous;
        let case = case_metrics(input, 0).expect("ambiguous evidence is reportable");
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
