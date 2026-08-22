//! Additive, operation-scoped metrics for timed harness operations.
//!
//! The filesystem child already records the counters that can be collected
//! without changing the timed operation.  This module only aligns those
//! per-child records with the sorted `elapsed_ns.samples` vector.  It does
//! not infer allocations, copies, decompression, recompression, or a peak from
//! a process-wide measurement that cannot provide one.  In-process sink
//! summaries are promoted separately when the harness has already proved that
//! the summary is deterministic across the retained samples.  `/proc/self/io`
//! vectors are explicitly scoped to the child interval including procfs probe
//! overhead; an after-snapshot read can add `rchar` and `syscr`.

use std::error::Error;

use serde::Serialize;

use crate::filesystem::{CfbPhaseEvidence, CfbPhaseSample, ReadPattern, SampleEvidence};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub(crate) enum MetricStatus {
    /// A numeric vector was measured for every retained sample.
    Measured,
    /// The selector does not expose this metric in its timed scope.
    NotApplicable,
    /// The metric is supported in principle, but the child could not collect
    /// it for this run (for example, procfs is unavailable).
    Unavailable,
    /// A counter or checked sample difference overflowed and the numeric
    /// vector was intentionally withheld.
    Overflow,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct MetricVector {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub values: Option<Vec<u64>>,
    pub status: MetricStatus,
    pub scope: &'static str,
}

impl MetricVector {
    fn measured(values: Vec<u64>, scope: &'static str) -> Self {
        Self {
            values: Some(values),
            status: MetricStatus::Measured,
            scope,
        }
    }

    fn absent(status: MetricStatus, scope: &'static str) -> Self {
        debug_assert_ne!(status, MetricStatus::Measured);
        Self {
            values: None,
            status,
            scope,
        }
    }
}

/// A categorical vector aligned with `elapsed_ns.samples`.
///
/// Pattern labels are descriptive source-observation evidence, not numeric
/// performance metrics. The wrapper still carries the same explicit status
/// and scope so an unavailable classification cannot be represented as a
/// fabricated string or zero.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct PatternVector {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub values: Option<Vec<ReadPattern>>,
    pub status: MetricStatus,
    pub scope: &'static str,
}

impl PatternVector {
    fn measured(values: Vec<ReadPattern>, scope: &'static str) -> Self {
        Self {
            values: Some(values),
            status: MetricStatus::Measured,
            scope,
        }
    }

    fn absent(status: MetricStatus, scope: &'static str) -> Self {
        debug_assert_ne!(status, MetricStatus::Measured);
        Self {
            values: None,
            status,
            scope,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct SourceMetrics {
    pub status: MetricStatus,
    pub counter_scope: String,
    pub logical_read_calls: MetricVector,
    pub logical_read_requested_bytes: MetricVector,
    pub logical_read_returned_bytes: MetricVector,
    pub logical_read_largest_requested_bytes: MetricVector,
    pub logical_read_largest_returned_bytes: MetricVector,
    pub logical_read_pattern: PatternVector,
    /// Exact compressed member bytes are not exposed by the generic `ReadAt`
    /// wrapper. This remains unavailable instead of treating raw source bytes
    /// as compressed payload bytes.
    pub compressed_bytes: MetricVector,
    /// Exact decompressed bytes are not exposed by the timed source boundary.
    pub decompressed_bytes: MetricVector,
    /// Exact recompressed bytes are not exposed by the atomic save callback.
    pub recompressed_bytes: MetricVector,
    pub max_concurrent_reads: MetricVector,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct ProcessMetrics {
    pub status: MetricStatus,
    pub user_cpu_ticks: MetricVector,
    pub system_cpu_ticks: MetricVector,
    pub clock_ticks_per_second: MetricVector,
    pub minor_faults: MetricVector,
    pub major_faults: MetricVector,
    pub voluntary_context_switches: MetricVector,
    pub nonvoluntary_context_switches: MetricVector,
    pub rss_delta_bytes: MetricVector,
    pub peak_rss_bytes: MetricVector,
    /// `/proc/self/io` bytes returned through read-like calls.
    pub rchar: MetricVector,
    /// `/proc/self/io` bytes accepted by write-like calls.
    pub wchar: MetricVector,
    /// `/proc/self/io` storage bytes read.
    pub read_bytes: MetricVector,
    /// `/proc/self/io` storage bytes written.
    pub write_bytes: MetricVector,
    /// `/proc/self/io` writes cancelled before storage.
    pub cancelled_write_bytes: MetricVector,
    /// `/proc/self/io` read-like syscall count.
    pub syscr: MetricVector,
    /// `/proc/self/io` write-like syscall count.
    pub syscw: MetricVector,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct SinkMetrics {
    /// Applicability status for `output_bytes`; retained for compatibility
    /// with the original filesystem envelope.
    pub status: MetricStatus,
    /// Final output length observed after a save.  This is not a write-call or
    /// memory-copy count; the scope says explicitly that it is post-operation.
    pub output_bytes: MetricVector,
    /// Applicability status for the logical accepted-write vectors below.
    /// These vectors are independent of `output_bytes` because a seekable sink
    /// can accept rewrites and therefore does not expose final output length.
    pub write_status: MetricStatus,
    /// Total bytes accepted by the instrumented sink's logical write calls.
    /// Requested lengths are not available in this summary and are not
    /// inferred.
    pub accepted_bytes: MetricVector,
    /// Number of logical write calls that accepted their reported length.
    pub write_calls: MetricVector,
    /// Largest accepted logical write length.
    pub largest_write: MetricVector,
    pub write_size_buckets: WriteSizeBucketMetrics,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct WriteSizeBucketMetrics {
    pub status: MetricStatus,
    pub bytes_0: MetricVector,
    pub bytes_1_to_512: MetricVector,
    pub bytes_513_to_4096: MetricVector,
    pub bytes_4097_to_16384: MetricVector,
    pub bytes_16385_to_65536: MetricVector,
    pub bytes_over_65536: MetricVector,
}

/// A per-case logical sink summary that has already been checked for
/// determinism across the retained samples.
///
/// All fields describe accepted lengths at the harness sink boundary.  The
/// sink summary does not retain requested lengths for short writes, rejected
/// calls, or operating-system syscall and storage-I/O counters.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct SinkObservation {
    pub accepted_bytes: u64,
    pub write_calls: u64,
    pub largest_write: u64,
    pub bytes_0: u64,
    pub bytes_1_to_512: u64,
    pub bytes_513_to_4096: u64,
    pub bytes_4097_to_16384: u64,
    pub bytes_16385_to_65536: u64,
    pub bytes_over_65536: u64,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct PublicationMetrics {
    pub status: MetricStatus,
    pub changed_spans: MetricVector,
    pub published_bytes: MetricVector,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct MaterializationMetrics {
    pub status: MetricStatus,
    pub opc_parts: MetricVector,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct CfbPhaseMetricSet {
    pub elapsed_ns: MetricVector,
    pub logical_read_calls: MetricVector,
    pub logical_read_requested_bytes: MetricVector,
    pub logical_read_returned_bytes: MetricVector,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct CfbPhaseMetrics {
    pub status: MetricStatus,
    pub open: CfbPhaseMetricSet,
    pub plan: CfbPhaseMetricSet,
    pub atomic_publication: CfbPhaseMetricSet,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct AllocationMetrics {
    pub status: MetricStatus,
    pub scope: crate::allocation_metrics::Scope,
    pub allocation_calls: MetricVector,
    pub deallocation_calls: MetricVector,
    pub reallocation_calls: MetricVector,
    pub failed_allocation_calls: MetricVector,
    pub allocated_bytes: MetricVector,
    pub deallocated_bytes: MetricVector,
    /// Absolute process live bytes before the timed operation. These values
    /// are deliberately not reset or presented as an operation peak.
    pub live_bytes_before: MetricVector,
    /// Absolute process live bytes after the timed operation.
    pub live_bytes_after: MetricVector,
    /// Absolute process high-water live bytes before the timed operation.
    pub peak_live_bytes_before: MetricVector,
    /// Absolute process high-water live bytes after the timed operation.
    pub peak_live_bytes_after: MetricVector,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct OperationMetrics {
    /// Number of values in every measured vector below.
    pub sample_count: usize,
    /// Stable child identity for every aligned metric vector.
    pub sample_indices: Vec<usize>,
    /// All vectors are ordered by `(elapsed_ns, sample_index)`; the explicit
    /// identity vector resolves otherwise indistinguishable elapsed ties.
    pub alignment: &'static str,
    /// Filesystem timings are retained as evidence only. They are not a
    /// latency claim or a matched eager/source performance comparison.
    pub latency_claim: &'static str,
    pub source: SourceMetrics,
    pub process: ProcessMetrics,
    pub sink: SinkMetrics,
    pub publication: PublicationMetrics,
    pub materialization: MaterializationMetrics,
    pub cfb_phases: CfbPhaseMetrics,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub allocation: Option<AllocationMetrics>,
}

/// One timed in-process operation and its best-effort process/allocator
/// observations.  The elapsed value is the same value published in the
/// caller's `Statistics`; the constructor below sorts all resource vectors by
/// that value and the original sample index.
#[derive(Clone, Debug)]
pub(crate) struct InProcessObservation {
    pub elapsed_ns: u64,
    pub process_metrics: Option<crate::process_metrics::Delta>,
    pub allocation_metrics: Option<crate::allocation_metrics::Sample>,
}

const ALIGNMENT: &str = "elapsed_ns.samples_by_elapsed_then_sample_index";
const LATENCY_CLAIM: &str = "evidence_only_filesystem_selector";
const COMPARABLE_LATENCY_CLAIM: &str = "comparable_timed_operation";
const SOURCE_SCOPE: &str = "operation_logical_read_at";
const SOURCE_PATTERN_SCOPE: &str = "operation_logical_read_at_range_order_not_physical_io";
const SOURCE_COMPRESSED_SCOPE: &str = "unavailable_read_at_has_no_compressed_member_boundary";
const SOURCE_DECOMPRESSED_SCOPE: &str = "unavailable_read_at_has_no_decompressed_byte_boundary";
const SOURCE_RECOMPRESSED_SCOPE: &str = "unavailable_atomic_save_has_no_recompressed_byte_boundary";
const PROCESS_SCOPE: &str = "procfs_operation_delta";
const PROC_IO_SCOPE: &str = "child_process_interval_delta_including_procfs_probe_overhead";
const CLOCK_SCOPE: &str = "procfs_after_sample_unit_factor";
const RSS_SCOPE: &str = "procfs_operation_delta_not_peak";
const HWM_SCOPE: &str = "process_lifetime_high_water_after_not_operation_peak";
const OUTPUT_SCOPE: &str = "post_operation_output_length_not_sink_write_volume";
const SINK_ACCEPTED_BYTES_SCOPE: &str = "logical_sink_accepted_write_bytes";
const SINK_WRITE_CALLS_SCOPE: &str = "logical_sink_accepted_write_calls";
const SINK_LARGEST_WRITE_SCOPE: &str = "logical_sink_largest_accepted_write";
const SINK_BUCKET_SCOPE: &str = "logical_sink_accepted_write_size_bucket_counts";
const SINK_SOURCE_SCOPE: &str = "not_applicable_in_process_sink";
const PUBLICATION_SCOPE: &str = "logical_publication_counter";
const MATERIALIZATION_SCOPE: &str = "logical_materialization_counter";
const CFB_PHASE_ELAPSED_SCOPE: &str = "timed_cfb_phase_elapsed_ns";
const CFB_PHASE_SOURCE_SCOPE: &str = "timed_cfb_phase_logical_read_at";
const ALLOCATION_SCOPE: &str = "operation_global_system_allocator";
const IN_PROCESS_PROCESS_SCOPE: &str =
    "procfs_in_process_operation_delta_including_procfs_probe_overhead";
const IN_PROCESS_PROC_IO_SCOPE: &str = IN_PROCESS_PROCESS_SCOPE;

/// Builds the additive envelope for one warm or cold `CaseResult`.
///
/// `elapsed` is intentionally supplied separately: `statistics` sorts its
/// input, so the child evidence must be sorted by the same elapsed values
/// before any numeric vector is published.  Every selected sample must have a
/// corresponding elapsed value, and every optional field must be either
/// present for all selected samples or absent for all of them.
pub(crate) fn aggregate(
    samples: &[SampleEvidence],
    cache_state: &str,
    elapsed: &[u64],
) -> Result<OperationMetrics, Box<dyn Error>> {
    let mut selected = samples
        .iter()
        .filter(|sample| sample.cache_state == cache_state)
        .collect::<Vec<_>>();
    if selected.len() != elapsed.len() {
        return Err(format!(
            "operation metrics {cache_state} sample count {} does not match elapsed sample count {}",
            selected.len(),
            elapsed.len()
        )
        .into());
    }
    if elapsed.is_empty() {
        return Err(format!(
            "operation metrics {cache_state} cannot publish an empty sample vector"
        )
        .into());
    }

    let mut expected_elapsed = elapsed.to_vec();
    expected_elapsed.sort_unstable();
    selected.sort_by_key(|sample| (sample.elapsed_ns, sample.sample_index));
    let observed_elapsed = selected
        .iter()
        .map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    if observed_elapsed != expected_elapsed {
        return Err(format!(
            "operation metrics {cache_state} elapsed samples do not match CaseResult ordering"
        )
        .into());
    }
    let selected = selected.into_iter().collect::<Vec<&SampleEvidence>>();
    let sample_count = selected.len();
    let sample_indices = selected
        .iter()
        .map(|sample| sample.sample_index)
        .collect::<Vec<_>>();
    let mut sorted_sample_indices = sample_indices.clone();
    sorted_sample_indices.sort_unstable();
    if sorted_sample_indices
        .windows(2)
        .any(|pair| pair[0] == pair[1])
    {
        return Err(
            format!("operation metrics {cache_state} sample identity is not unique").into(),
        );
    }
    let first_scope = selected
        .first()
        .map(|sample| sample.logical_read_counter_scope.clone())
        .unwrap_or_else(|| "no_samples".to_owned());
    if selected
        .iter()
        .any(|sample| sample.logical_read_counter_scope != first_scope)
    {
        return Err(format!(
            "operation metrics {cache_state} logical-read scope changes between samples"
        )
        .into());
    }

    let source_status = match first_scope.as_str() {
        "timed_read_at" => MetricStatus::Measured,
        "not_applicable_in_process_sink"
        | "untimed_source_replay_only"
        | "not_applicable_eager_opc"
        | "not_applicable_eager_pptx"
        | "not_applicable_eager_docx"
        | "not_applicable_immutable_owned_slice" => MetricStatus::NotApplicable,
        _ => MetricStatus::Unavailable,
    };
    let source_values = |value: fn(&SampleEvidence) -> u64| {
        if source_status == MetricStatus::Measured {
            MetricVector::measured(
                selected.iter().map(|sample| value(sample)).collect(),
                SOURCE_SCOPE,
            )
        } else {
            MetricVector::absent(source_status, SOURCE_SCOPE)
        }
    };
    let source_pattern = if source_status == MetricStatus::Measured {
        let values = selected
            .iter()
            .map(|sample| {
                sample.logical_read_pattern.ok_or_else(|| {
                    std::io::Error::other(format!(
                        "operation metrics {cache_state} measured source sample omitted logical read pattern"
                    ))
                })
            })
            .collect::<Result<Vec<_>, _>>()?;
        PatternVector::measured(values, SOURCE_PATTERN_SCOPE)
    } else {
        PatternVector::absent(source_status, SOURCE_PATTERN_SCOPE)
    };
    let unavailable_source_value = |scope| {
        MetricVector::absent(
            if source_status == MetricStatus::NotApplicable {
                MetricStatus::NotApplicable
            } else {
                MetricStatus::Unavailable
            },
            scope,
        )
    };
    let source = SourceMetrics {
        status: source_status,
        counter_scope: first_scope,
        logical_read_calls: source_values(|sample| sample.logical_read_calls),
        logical_read_requested_bytes: source_values(|sample| sample.logical_read_requested_bytes),
        logical_read_returned_bytes: source_values(|sample| sample.logical_read_bytes),
        logical_read_largest_requested_bytes: source_values(|sample| {
            sample.logical_read_largest_requested_bytes
        }),
        logical_read_largest_returned_bytes: source_values(|sample| {
            sample.logical_read_largest_returned_bytes
        }),
        logical_read_pattern: source_pattern,
        compressed_bytes: unavailable_source_value(SOURCE_COMPRESSED_SCOPE),
        decompressed_bytes: unavailable_source_value(SOURCE_DECOMPRESSED_SCOPE),
        recompressed_bytes: unavailable_source_value(SOURCE_RECOMPRESSED_SCOPE),
        max_concurrent_reads: source_values(|sample| sample.max_concurrent_reads),
    };

    let process = process_metrics(&selected)?;
    let output_status = optional_status(
        &selected,
        |sample| sample.output_bytes.is_some(),
        "output_bytes",
    )?;
    let sink = SinkMetrics {
        status: output_status,
        output_bytes: optional_values(
            &selected,
            |sample| sample.output_bytes,
            "output_bytes",
            MetricStatus::NotApplicable,
            OUTPUT_SCOPE,
        )?,
        write_status: MetricStatus::NotApplicable,
        accepted_bytes: MetricVector::absent(
            MetricStatus::NotApplicable,
            SINK_ACCEPTED_BYTES_SCOPE,
        ),
        write_calls: MetricVector::absent(MetricStatus::NotApplicable, SINK_WRITE_CALLS_SCOPE),
        largest_write: MetricVector::absent(MetricStatus::NotApplicable, SINK_LARGEST_WRITE_SCOPE),
        write_size_buckets: WriteSizeBucketMetrics::absent(
            MetricStatus::NotApplicable,
            SINK_BUCKET_SCOPE,
        ),
    };
    let publication = PublicationMetrics {
        status: publication_status(&selected)?,
        changed_spans: optional_values(
            &selected,
            |sample| sample.cfb_changed_spans,
            "cfb_changed_spans",
            MetricStatus::NotApplicable,
            PUBLICATION_SCOPE,
        )?,
        published_bytes: optional_values(
            &selected,
            |sample| sample.cfb_published_bytes,
            "cfb_published_bytes",
            MetricStatus::NotApplicable,
            PUBLICATION_SCOPE,
        )?,
    };
    let materialization = MaterializationMetrics {
        status: optional_status(
            &selected,
            |sample| sample.opc_materialized_parts.is_some(),
            "opc_materialized_parts",
        )?,
        opc_parts: optional_values(
            &selected,
            |sample| sample.opc_materialized_parts,
            "opc_materialized_parts",
            MetricStatus::NotApplicable,
            MATERIALIZATION_SCOPE,
        )?,
    };
    let cfb_phases = cfb_phase_metrics(&selected)?;
    let allocation = allocation_metrics(&selected)?;

    Ok(OperationMetrics {
        sample_count,
        sample_indices,
        alignment: ALIGNMENT,
        latency_claim: LATENCY_CLAIM,
        source,
        process,
        sink,
        publication,
        materialization,
        cfb_phases,
        allocation,
    })
}

fn allocation_metrics(
    samples: &[&SampleEvidence],
) -> Result<Option<AllocationMetrics>, Box<dyn Error>> {
    let presence = samples
        .iter()
        .map(|sample| sample.allocation_metrics.is_some())
        .collect::<Vec<_>>();
    let all_present = presence.iter().all(|present| *present);
    let all_absent = presence.iter().all(|present| !present);
    if !all_present && !all_absent {
        return Err("operation metrics allocation option cardinality is asymmetric".into());
    }
    if all_absent {
        return Ok(None);
    }

    let observations = samples
        .iter()
        .map(|sample| {
            sample
                .allocation_metrics
                .as_ref()
                .expect("presence was checked for every sample")
        })
        .collect::<Vec<_>>();
    let first_status = observations[0].status;
    let first_scope = observations[0].scope;
    if observations
        .iter()
        .any(|sample| sample.status != first_status || sample.scope != first_scope)
    {
        return Err("operation metrics allocation status or scope changes between samples".into());
    }
    let absent = |status: MetricStatus| MetricVector::absent(status, ALLOCATION_SCOPE);
    if first_status != crate::allocation_metrics::Status::Measured {
        let status = match first_status {
            crate::allocation_metrics::Status::Unavailable => MetricStatus::Unavailable,
            crate::allocation_metrics::Status::Overflow => MetricStatus::Overflow,
            crate::allocation_metrics::Status::Measured => unreachable!(),
        };
        return Ok(Some(AllocationMetrics {
            status,
            scope: first_scope,
            allocation_calls: absent(status),
            deallocation_calls: absent(status),
            reallocation_calls: absent(status),
            failed_allocation_calls: absent(status),
            allocated_bytes: absent(status),
            deallocated_bytes: absent(status),
            live_bytes_before: absent(status),
            live_bytes_after: absent(status),
            peak_live_bytes_before: absent(status),
            peak_live_bytes_after: absent(status),
        }));
    }

    let measured = |name: &str,
                    value: fn(&crate::allocation_metrics::Sample) -> Option<u64>|
     -> Result<MetricVector, Box<dyn Error>> {
        let values = observations
            .iter()
            .map(|sample| value(sample))
            .collect::<Vec<_>>();
        if values.iter().any(Option::is_none) {
            return Err(format!("operation metrics allocation {name} value is missing").into());
        }
        Ok(MetricVector::measured(
            values
                .into_iter()
                .map(|value| value.expect("presence was checked for every sample"))
                .collect(),
            ALLOCATION_SCOPE,
        ))
    };
    Ok(Some(AllocationMetrics {
        status: MetricStatus::Measured,
        scope: first_scope,
        allocation_calls: measured("allocation_calls", |sample| sample.allocation_calls)?,
        deallocation_calls: measured("deallocation_calls", |sample| sample.deallocation_calls)?,
        reallocation_calls: measured("reallocation_calls", |sample| sample.reallocation_calls)?,
        failed_allocation_calls: measured("failed_allocation_calls", |sample| {
            sample.failed_allocation_calls
        })?,
        allocated_bytes: measured("allocated_bytes", |sample| sample.allocated_bytes)?,
        deallocated_bytes: measured("deallocated_bytes", |sample| sample.deallocated_bytes)?,
        live_bytes_before: measured("live_bytes_before", |sample| sample.live_bytes_before)?,
        live_bytes_after: measured("live_bytes_after", |sample| sample.live_bytes_after)?,
        peak_live_bytes_before: measured("peak_live_bytes_before", |sample| {
            sample.peak_live_bytes_before
        })?,
        peak_live_bytes_after: measured("peak_live_bytes_after", |sample| {
            sample.peak_live_bytes_after
        })?,
    }))
}

/// Builds operation metrics for an in-process case whose sink summary was
/// already checked to be identical for every retained elapsed sample.
pub(crate) fn from_sink_observation(
    sample_count: usize,
    observation: SinkObservation,
) -> Result<OperationMetrics, Box<dyn Error>> {
    if sample_count == 0 {
        return Err("operation metrics sink observation cannot have zero samples".into());
    }
    Ok(OperationMetrics {
        sample_count,
        sample_indices: (0..sample_count).collect(),
        alignment: ALIGNMENT,
        latency_claim: COMPARABLE_LATENCY_CLAIM,
        source: SourceMetrics {
            status: MetricStatus::NotApplicable,
            counter_scope: SINK_SOURCE_SCOPE.to_owned(),
            logical_read_calls: MetricVector::absent(MetricStatus::NotApplicable, SOURCE_SCOPE),
            logical_read_requested_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_SCOPE,
            ),
            logical_read_returned_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_SCOPE,
            ),
            logical_read_largest_requested_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_SCOPE,
            ),
            logical_read_largest_returned_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_SCOPE,
            ),
            logical_read_pattern: PatternVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_PATTERN_SCOPE,
            ),
            compressed_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_COMPRESSED_SCOPE,
            ),
            decompressed_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_DECOMPRESSED_SCOPE,
            ),
            recompressed_bytes: MetricVector::absent(
                MetricStatus::NotApplicable,
                SOURCE_RECOMPRESSED_SCOPE,
            ),
            max_concurrent_reads: MetricVector::absent(MetricStatus::NotApplicable, SOURCE_SCOPE),
        },
        process: absent_process_metrics(MetricStatus::NotApplicable),
        sink: sink_metrics_for_observation(sample_count, observation),
        publication: PublicationMetrics {
            status: MetricStatus::NotApplicable,
            changed_spans: MetricVector::absent(MetricStatus::NotApplicable, PUBLICATION_SCOPE),
            published_bytes: MetricVector::absent(MetricStatus::NotApplicable, PUBLICATION_SCOPE),
        },
        materialization: MaterializationMetrics {
            status: MetricStatus::NotApplicable,
            opc_parts: MetricVector::absent(MetricStatus::NotApplicable, MATERIALIZATION_SCOPE),
        },
        cfb_phases: absent_cfb_phase_metrics(),
        allocation: None,
    })
}

/// Builds the operation-metrics envelope for a fixed-window in-process
/// writer.  The source and publication fields are deliberately marked
/// not-applicable because this operation creates directly into a logical
/// sink; process and allocator observations remain best-effort and therefore
/// publish `unavailable` when the corresponding provider is absent.
pub(crate) fn from_in_process_observations(
    observations: &[InProcessObservation],
    sink: SinkObservation,
) -> Result<OperationMetrics, Box<dyn Error>> {
    if observations.is_empty() {
        return Err("in-process operation metrics cannot have zero observations".into());
    }

    let samples = observations
        .iter()
        .enumerate()
        .map(|(sample_index, observation)| SampleEvidence {
            sample_index,
            cache_state: "warm",
            elapsed_ns: observation.elapsed_ns,
            parent_wall_ns: observation.elapsed_ns,
            cold_advice: crate::filesystem::ColdAdvice::NotRequested,
            cold_verified: None,
            logical_read_counter_scope: SINK_SOURCE_SCOPE.to_owned(),
            logical_read_calls: 0,
            logical_read_requested_bytes: 0,
            logical_read_bytes: 0,
            logical_read_largest_requested_bytes: 0,
            logical_read_largest_returned_bytes: 0,
            logical_read_pattern: None,
            max_concurrent_reads: 0,
            logical_read_request_sizes: Vec::new(),
            logical_read_request_size_buckets: crate::filesystem::ReadSizeBuckets::default(),
            process_metrics: observation.process_metrics,
            allocation_metrics: observation.allocation_metrics.clone(),
            output_sha256: None,
            output_bytes: None,
            opc_materialized_parts: None,
            cfb_changed_spans: None,
            cfb_published_bytes: None,
            cfb_phases: None,
            pptx_source_replay: None,
            docx_source_replay: None,
        })
        .collect::<Vec<_>>();
    let elapsed = observations
        .iter()
        .map(|observation| observation.elapsed_ns)
        .collect::<Vec<_>>();
    let mut metrics = aggregate(&samples, "warm", &elapsed)?;
    metrics.latency_claim = COMPARABLE_LATENCY_CLAIM;
    relabel_in_process_scopes(&mut metrics.process);
    metrics.set_sink_observation(observations.len(), sink)?;
    Ok(metrics)
}

fn relabel_in_process_scopes(metrics: &mut ProcessMetrics) {
    for vector in [
        &mut metrics.user_cpu_ticks,
        &mut metrics.system_cpu_ticks,
        &mut metrics.clock_ticks_per_second,
        &mut metrics.minor_faults,
        &mut metrics.major_faults,
        &mut metrics.voluntary_context_switches,
        &mut metrics.nonvoluntary_context_switches,
    ] {
        vector.scope = IN_PROCESS_PROCESS_SCOPE;
    }
    for vector in [
        &mut metrics.rchar,
        &mut metrics.wchar,
        &mut metrics.read_bytes,
        &mut metrics.write_bytes,
        &mut metrics.cancelled_write_bytes,
        &mut metrics.syscr,
        &mut metrics.syscw,
    ] {
        vector.scope = IN_PROCESS_PROC_IO_SCOPE;
    }
    metrics.rss_delta_bytes.scope = "procfs_in_process_rss_delta_including_procfs_probe_overhead";
    metrics.peak_rss_bytes.scope = HWM_SCOPE;
}

impl OperationMetrics {
    /// Adds an already-proven deterministic sink summary without changing the
    /// operation's measured elapsed samples or any existing metric vectors.
    pub(crate) fn set_sink_observation(
        &mut self,
        sample_count: usize,
        observation: SinkObservation,
    ) -> Result<(), Box<dyn Error>> {
        if sample_count == 0 {
            return Err("operation metrics sink observation cannot have zero samples".into());
        }
        if self.sample_count != sample_count {
            return Err(format!(
                "operation metrics sink observation sample count {sample_count} does not match envelope sample count {}",
                self.sample_count
            )
            .into());
        }
        let observed = sink_metrics_for_observation(sample_count, observation);
        self.sink.write_status = observed.write_status;
        self.sink.accepted_bytes = observed.accepted_bytes;
        self.sink.write_calls = observed.write_calls;
        self.sink.largest_write = observed.largest_write;
        self.sink.write_size_buckets = observed.write_size_buckets;
        Ok(())
    }
}

fn sink_metrics_for_observation(sample_count: usize, observation: SinkObservation) -> SinkMetrics {
    SinkMetrics {
        status: MetricStatus::NotApplicable,
        output_bytes: MetricVector::absent(MetricStatus::NotApplicable, OUTPUT_SCOPE),
        write_status: MetricStatus::Measured,
        accepted_bytes: repeated_metric(
            observation.accepted_bytes,
            sample_count,
            SINK_ACCEPTED_BYTES_SCOPE,
        ),
        write_calls: repeated_metric(
            observation.write_calls,
            sample_count,
            SINK_WRITE_CALLS_SCOPE,
        ),
        largest_write: repeated_metric(
            observation.largest_write,
            sample_count,
            SINK_LARGEST_WRITE_SCOPE,
        ),
        write_size_buckets: WriteSizeBucketMetrics::measured(sample_count, observation),
    }
}

fn repeated_metric(value: u64, sample_count: usize, scope: &'static str) -> MetricVector {
    MetricVector::measured(vec![value; sample_count], scope)
}

impl WriteSizeBucketMetrics {
    fn absent(status: MetricStatus, scope: &'static str) -> Self {
        Self {
            status,
            bytes_0: MetricVector::absent(status, scope),
            bytes_1_to_512: MetricVector::absent(status, scope),
            bytes_513_to_4096: MetricVector::absent(status, scope),
            bytes_4097_to_16384: MetricVector::absent(status, scope),
            bytes_16385_to_65536: MetricVector::absent(status, scope),
            bytes_over_65536: MetricVector::absent(status, scope),
        }
    }

    fn measured(sample_count: usize, observation: SinkObservation) -> Self {
        Self {
            status: MetricStatus::Measured,
            bytes_0: repeated_metric(observation.bytes_0, sample_count, SINK_BUCKET_SCOPE),
            bytes_1_to_512: repeated_metric(
                observation.bytes_1_to_512,
                sample_count,
                SINK_BUCKET_SCOPE,
            ),
            bytes_513_to_4096: repeated_metric(
                observation.bytes_513_to_4096,
                sample_count,
                SINK_BUCKET_SCOPE,
            ),
            bytes_4097_to_16384: repeated_metric(
                observation.bytes_4097_to_16384,
                sample_count,
                SINK_BUCKET_SCOPE,
            ),
            bytes_16385_to_65536: repeated_metric(
                observation.bytes_16385_to_65536,
                sample_count,
                SINK_BUCKET_SCOPE,
            ),
            bytes_over_65536: repeated_metric(
                observation.bytes_over_65536,
                sample_count,
                SINK_BUCKET_SCOPE,
            ),
        }
    }
}

fn absent_process_metrics(status: MetricStatus) -> ProcessMetrics {
    let process = || MetricVector::absent(status, PROCESS_SCOPE);
    let proc_io = || MetricVector::absent(status, PROC_IO_SCOPE);
    ProcessMetrics {
        status,
        rchar: proc_io(),
        wchar: proc_io(),
        read_bytes: proc_io(),
        write_bytes: proc_io(),
        cancelled_write_bytes: proc_io(),
        syscr: proc_io(),
        syscw: proc_io(),
        user_cpu_ticks: process(),
        system_cpu_ticks: process(),
        clock_ticks_per_second: process(),
        minor_faults: process(),
        major_faults: process(),
        voluntary_context_switches: process(),
        nonvoluntary_context_switches: process(),
        rss_delta_bytes: MetricVector::absent(status, RSS_SCOPE),
        peak_rss_bytes: MetricVector::absent(status, HWM_SCOPE),
    }
}

fn absent_cfb_phase_metrics() -> CfbPhaseMetrics {
    let metric_set = || CfbPhaseMetricSet {
        elapsed_ns: MetricVector::absent(MetricStatus::NotApplicable, CFB_PHASE_ELAPSED_SCOPE),
        logical_read_calls: MetricVector::absent(
            MetricStatus::NotApplicable,
            CFB_PHASE_SOURCE_SCOPE,
        ),
        logical_read_requested_bytes: MetricVector::absent(
            MetricStatus::NotApplicable,
            CFB_PHASE_SOURCE_SCOPE,
        ),
        logical_read_returned_bytes: MetricVector::absent(
            MetricStatus::NotApplicable,
            CFB_PHASE_SOURCE_SCOPE,
        ),
    };
    CfbPhaseMetrics {
        status: MetricStatus::NotApplicable,
        open: metric_set(),
        plan: metric_set(),
        atomic_publication: metric_set(),
    }
}

fn cfb_phase_metrics(samples: &[&SampleEvidence]) -> Result<CfbPhaseMetrics, Box<dyn Error>> {
    let status = optional_status(samples, |sample| sample.cfb_phases.is_some(), "cfb_phases")?;
    let phase = |name: &str, select: fn(&CfbPhaseEvidence) -> &CfbPhaseSample| {
        cfb_phase_metric_set(samples, status, name, select)
    };
    Ok(CfbPhaseMetrics {
        status,
        open: phase("open", |phases| &phases.open)?,
        plan: phase("plan", |phases| &phases.plan)?,
        atomic_publication: phase("atomic_publication", |phases| &phases.atomic_publication)?,
    })
}

fn cfb_phase_metric_set(
    samples: &[&SampleEvidence],
    status: MetricStatus,
    name: &str,
    select: fn(&CfbPhaseEvidence) -> &CfbPhaseSample,
) -> Result<CfbPhaseMetricSet, Box<dyn Error>> {
    let value = |suffix: &str, field: fn(&CfbPhaseSample) -> u64, scope: &'static str| {
        optional_values(
            samples,
            |sample| sample.cfb_phases.as_ref().map(select).map(field),
            &format!("cfb_phases.{name}.{suffix}"),
            MetricStatus::NotApplicable,
            scope,
        )
    };
    let metrics = CfbPhaseMetricSet {
        elapsed_ns: value(
            "elapsed_ns",
            |phase| phase.elapsed_ns,
            CFB_PHASE_ELAPSED_SCOPE,
        )?,
        logical_read_calls: value(
            "logical_read_calls",
            |phase| phase.logical_read_calls,
            CFB_PHASE_SOURCE_SCOPE,
        )?,
        logical_read_requested_bytes: value(
            "logical_read_requested_bytes",
            |phase| phase.logical_read_requested_bytes,
            CFB_PHASE_SOURCE_SCOPE,
        )?,
        logical_read_returned_bytes: value(
            "logical_read_returned_bytes",
            |phase| phase.logical_read_returned_bytes,
            CFB_PHASE_SOURCE_SCOPE,
        )?,
    };
    if status == MetricStatus::Measured && metrics.elapsed_ns.status != MetricStatus::Measured {
        return Err(format!("operation metrics CFB phase {name} is unexpectedly absent").into());
    }
    Ok(metrics)
}

fn process_metrics(samples: &[&SampleEvidence]) -> Result<ProcessMetrics, Box<dyn Error>> {
    let presence = samples
        .iter()
        .map(|sample| sample.process_metrics.is_some())
        .collect::<Vec<_>>();
    let all_present = presence.iter().all(|present| *present);
    let all_absent = presence.iter().all(|present| !present);
    if !all_present && !all_absent {
        return Err("operation metrics process_metrics option cardinality is asymmetric".into());
    }
    if all_absent {
        let unavailable = || MetricVector::absent(MetricStatus::Unavailable, PROCESS_SCOPE);
        let unavailable_proc_io = || MetricVector::absent(MetricStatus::Unavailable, PROC_IO_SCOPE);
        return Ok(ProcessMetrics {
            status: MetricStatus::Unavailable,
            rchar: unavailable_proc_io(),
            wchar: unavailable_proc_io(),
            read_bytes: unavailable_proc_io(),
            write_bytes: unavailable_proc_io(),
            cancelled_write_bytes: unavailable_proc_io(),
            syscr: unavailable_proc_io(),
            syscw: unavailable_proc_io(),
            user_cpu_ticks: unavailable(),
            system_cpu_ticks: unavailable(),
            clock_ticks_per_second: unavailable(),
            minor_faults: unavailable(),
            major_faults: unavailable(),
            voluntary_context_switches: unavailable(),
            nonvoluntary_context_switches: unavailable(),
            rss_delta_bytes: MetricVector::absent(MetricStatus::Unavailable, RSS_SCOPE),
            peak_rss_bytes: MetricVector::absent(MetricStatus::Unavailable, HWM_SCOPE),
        });
    }

    let values = samples
        .iter()
        .map(|sample| {
            sample
                .process_metrics
                .as_ref()
                .expect("presence was checked for every sample")
        })
        .collect::<Vec<_>>();
    let measured = |value: fn(&crate::process_metrics::Delta) -> u64| {
        MetricVector::measured(
            values.iter().map(|metrics| value(metrics)).collect(),
            PROCESS_SCOPE,
        )
    };
    let measured_proc_io = |value: fn(&crate::process_metrics::Delta) -> u64| {
        MetricVector::measured(
            values.iter().map(|metrics| value(metrics)).collect(),
            PROC_IO_SCOPE,
        )
    };
    Ok(ProcessMetrics {
        status: MetricStatus::Measured,
        rchar: measured_proc_io(|metrics| metrics.rchar),
        wchar: measured_proc_io(|metrics| metrics.wchar),
        read_bytes: measured_proc_io(|metrics| metrics.read_bytes),
        write_bytes: measured_proc_io(|metrics| metrics.write_bytes),
        cancelled_write_bytes: measured_proc_io(|metrics| metrics.cancelled_write_bytes),
        syscr: measured_proc_io(|metrics| metrics.syscr),
        syscw: measured_proc_io(|metrics| metrics.syscw),
        user_cpu_ticks: measured(|metrics| metrics.user_cpu_ticks),
        system_cpu_ticks: measured(|metrics| metrics.system_cpu_ticks),
        clock_ticks_per_second: MetricVector::measured(
            values
                .iter()
                .map(|metrics| metrics.clock_ticks_per_second)
                .collect(),
            CLOCK_SCOPE,
        ),
        minor_faults: measured(|metrics| metrics.minor_faults),
        major_faults: measured(|metrics| metrics.major_faults),
        voluntary_context_switches: measured(|metrics| metrics.voluntary_context_switches),
        nonvoluntary_context_switches: measured(|metrics| metrics.nonvoluntary_context_switches),
        rss_delta_bytes: MetricVector::measured(
            values.iter().map(|metrics| metrics.rss_bytes).collect(),
            RSS_SCOPE,
        ),
        peak_rss_bytes: MetricVector::measured(
            values
                .iter()
                .map(|metrics| metrics.peak_rss_bytes)
                .collect(),
            HWM_SCOPE,
        ),
    })
}

fn optional_status(
    samples: &[&SampleEvidence],
    present: impl Fn(&SampleEvidence) -> bool,
    name: &str,
) -> Result<MetricStatus, Box<dyn Error>> {
    let mut seen_present = false;
    let mut seen_absent = false;
    for sample in samples {
        if present(sample) {
            seen_present = true;
        } else {
            seen_absent = true;
        }
    }
    if seen_present && seen_absent {
        return Err(format!("operation metrics {name} option cardinality is asymmetric").into());
    }
    Ok(if seen_present {
        MetricStatus::Measured
    } else {
        MetricStatus::NotApplicable
    })
}

fn publication_status(samples: &[&SampleEvidence]) -> Result<MetricStatus, Box<dyn Error>> {
    let changed = optional_status(
        samples,
        |sample| sample.cfb_changed_spans.is_some(),
        "cfb_changed_spans",
    )?;
    let published = optional_status(
        samples,
        |sample| sample.cfb_published_bytes.is_some(),
        "cfb_published_bytes",
    )?;
    if changed != published {
        return Err("operation metrics publication counters have asymmetric applicability".into());
    }
    Ok(changed)
}

fn optional_values(
    samples: &[&SampleEvidence],
    value: impl Fn(&SampleEvidence) -> Option<u64>,
    name: &str,
    absent_status: MetricStatus,
    scope: &'static str,
) -> Result<MetricVector, Box<dyn Error>> {
    let values = samples
        .iter()
        .map(|sample| value(sample))
        .collect::<Vec<_>>();
    let present = values.iter().filter(|value| value.is_some()).count();
    if present != 0 && present != values.len() {
        return Err(format!("operation metrics {name} option cardinality is asymmetric").into());
    }
    if present == 0 {
        return Ok(MetricVector::absent(absent_status, scope));
    }
    Ok(MetricVector::measured(
        values
            .into_iter()
            .map(|value| value.expect("presence was checked for every sample"))
            .collect(),
        scope,
    ))
}

#[cfg(test)]
mod tests {
    use serde_json::Value;

    use super::{
        InProcessObservation, MetricStatus, MetricVector, SinkObservation, aggregate,
        from_in_process_observations, from_sink_observation,
    };
    use crate::filesystem::{
        CfbPhaseEvidence, CfbPhaseSample, ColdAdvice, ReadPattern, ReadSizeBuckets, SampleEvidence,
    };

    fn sample(
        index: usize,
        cache_state: &'static str,
        elapsed_ns: u64,
        process_metrics: Option<crate::process_metrics::Delta>,
    ) -> SampleEvidence {
        SampleEvidence {
            sample_index: index,
            cache_state,
            elapsed_ns,
            parent_wall_ns: elapsed_ns,
            cold_advice: ColdAdvice::NotRequested,
            cold_verified: None,
            logical_read_counter_scope: "timed_read_at".to_owned(),
            logical_read_calls: index as u64,
            logical_read_requested_bytes: 10,
            logical_read_bytes: 8,
            logical_read_largest_requested_bytes: 10,
            logical_read_largest_returned_bytes: 8,
            logical_read_pattern: Some(ReadPattern::Sequential),
            max_concurrent_reads: 1,
            logical_read_request_sizes: vec![10],
            logical_read_request_size_buckets: ReadSizeBuckets::default(),
            process_metrics,
            allocation_metrics: None,
            output_sha256: None,
            output_bytes: None,
            opc_materialized_parts: None,
            cfb_changed_spans: None,
            cfb_published_bytes: None,
            cfb_phases: None,
            pptx_source_replay: None,
            docx_source_replay: None,
        }
    }

    fn metrics() -> crate::process_metrics::Delta {
        crate::process_metrics::Delta {
            rchar: 11,
            wchar: 12,
            read_bytes: 13,
            write_bytes: 14,
            cancelled_write_bytes: 15,
            syscr: 16,
            syscw: 17,
            user_cpu_ticks: 2,
            system_cpu_ticks: 3,
            clock_ticks_per_second: 100,
            minor_faults: 5,
            major_faults: 6,
            voluntary_context_switches: 7,
            nonvoluntary_context_switches: 8,
            rss_bytes: 9,
            peak_rss_bytes: 10,
            ..crate::process_metrics::Delta::default()
        }
    }

    fn measured_allocation_sample(calls: u64) -> crate::allocation_metrics::Sample {
        crate::allocation_metrics::Sample {
            status: crate::allocation_metrics::Status::Measured,
            scope: crate::allocation_metrics::Scope::OperationGlobalSystemAllocator,
            allocation_calls: Some(calls),
            deallocation_calls: Some(2),
            reallocation_calls: Some(1),
            failed_allocation_calls: Some(0),
            allocated_bytes: Some(100),
            deallocated_bytes: Some(64),
            live_bytes_before: Some(1_000),
            live_bytes_after: Some(1_036),
            peak_live_bytes_before: Some(1_010),
            peak_live_bytes_after: Some(1_046),
        }
    }

    #[test]
    fn warm_and_cold_partitioning_is_exact_and_sorted_like_elapsed() {
        let samples = vec![
            sample(1, "cold-requested", 30, Some(metrics())),
            sample(0, "warm", 20, Some(metrics())),
            sample(1, "warm", 10, Some(metrics())),
            sample(0, "cold-requested", 40, Some(metrics())),
        ];
        let warm = aggregate(&samples, "warm", &[20, 10]).unwrap();
        assert_eq!(warm.sample_count, 2);
        assert_eq!(warm.sample_indices, vec![1, 0]);
        assert_eq!(warm.latency_claim, "evidence_only_filesystem_selector");
        assert_eq!(warm.source.logical_read_calls.values, Some(vec![1, 0]));
        assert_eq!(
            warm.source.logical_read_largest_requested_bytes.values,
            Some(vec![10, 10])
        );
        assert_eq!(
            warm.source.logical_read_pattern.values,
            Some(vec![ReadPattern::Sequential, ReadPattern::Sequential])
        );
        let cold = aggregate(&samples, "cold-requested", &[40, 30]).unwrap();
        assert_eq!(cold.source.logical_read_calls.values, Some(vec![1, 0]));
    }

    #[test]
    fn tied_elapsed_samples_keep_stable_sample_identity_alignment() {
        let samples = vec![
            sample(2, "warm", 10, Some(metrics())),
            sample(1, "warm", 10, Some(metrics())),
            sample(0, "warm", 20, Some(metrics())),
        ];
        let envelope = aggregate(&samples, "warm", &[10, 10, 20]).unwrap();
        assert_eq!(
            envelope.sample_indices,
            vec![1, 2, 0],
            "ties must use sample index rather than unstable elapsed sorting"
        );
        assert_eq!(
            envelope.source.logical_read_calls.values,
            Some(vec![1, 2, 0])
        );
    }

    #[test]
    fn allocation_vectors_are_aligned_and_absolute_live_counts_are_retained() {
        let allocation = |calls, allocated, before, after| crate::allocation_metrics::Sample {
            status: crate::allocation_metrics::Status::Measured,
            scope: crate::allocation_metrics::Scope::OperationGlobalSystemAllocator,
            allocation_calls: Some(calls),
            deallocation_calls: Some(2),
            reallocation_calls: Some(1),
            failed_allocation_calls: Some(0),
            allocated_bytes: Some(allocated),
            deallocated_bytes: Some(64),
            live_bytes_before: Some(before),
            live_bytes_after: Some(after),
            peak_live_bytes_before: Some(before + 10),
            peak_live_bytes_after: Some(after + 10),
        };
        let mut first = sample(0, "warm", 20, Some(metrics()));
        first.allocation_metrics = Some(allocation(7, 100, 1_000, 1_036));
        let mut second = sample(1, "warm", 10, Some(metrics()));
        second.allocation_metrics = Some(allocation(5, 80, 1_036, 1_052));
        let envelope = aggregate(&[first, second], "warm", &[20, 10]).unwrap();
        let allocation = envelope.allocation.unwrap();
        assert_eq!(allocation.status, MetricStatus::Measured);
        assert_eq!(allocation.allocation_calls.values, Some(vec![5, 7]));
        assert_eq!(
            allocation.live_bytes_before.values,
            Some(vec![1_036, 1_000])
        );
        let json = serde_json::to_value(&allocation).unwrap();
        assert_eq!(json["status"], "measured");
        assert_eq!(
            json["allocation_calls"]["values"],
            Value::from(vec![5_u64, 7])
        );
        assert_eq!(
            json["live_bytes_after"]["values"],
            Value::from(vec![1_052_u64, 1_036])
        );
        assert_eq!(
            allocation.peak_live_bytes_before.values,
            Some(vec![1_046, 1_010])
        );
        assert_eq!(
            allocation.peak_live_bytes_after.values,
            Some(vec![1_062, 1_046])
        );
    }

    #[test]
    fn allocation_status_scope_and_cardinality_are_strict() {
        let mut measured = sample(0, "warm", 10, Some(metrics()));
        measured.allocation_metrics = Some(measured_allocation_sample(1));
        let mut overflow = sample(1, "warm", 20, Some(metrics()));
        overflow.allocation_metrics = Some(crate::allocation_metrics::Sample {
            status: crate::allocation_metrics::Status::Overflow,
            scope: crate::allocation_metrics::Scope::OperationGlobalSystemAllocator,
            allocation_calls: None,
            deallocation_calls: None,
            reallocation_calls: None,
            failed_allocation_calls: None,
            allocated_bytes: None,
            deallocated_bytes: None,
            live_bytes_before: None,
            live_bytes_after: None,
            peak_live_bytes_before: None,
            peak_live_bytes_after: None,
        });
        let error = aggregate(&[measured.clone(), overflow], "warm", &[10, 20]).unwrap_err();
        assert!(error.to_string().contains("status or scope"));

        measured
            .allocation_metrics
            .as_mut()
            .unwrap()
            .peak_live_bytes_after = None;
        let error = aggregate(&[measured.clone()], "warm", &[10]).unwrap_err();
        assert!(error.to_string().contains("peak_live_bytes_after"));

        measured.allocation_metrics = None;
        let mut present = sample(1, "warm", 20, Some(metrics()));
        present.allocation_metrics = Some(measured_allocation_sample(2));
        let error = aggregate(&[measured, present], "warm", &[10, 20]).unwrap_err();
        assert!(error.to_string().contains("option cardinality"));
    }

    #[test]
    fn allocation_overflow_omits_numeric_vectors() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.allocation_metrics = Some(crate::allocation_metrics::Sample {
            status: crate::allocation_metrics::Status::Overflow,
            scope: crate::allocation_metrics::Scope::OperationGlobalSystemAllocator,
            allocation_calls: None,
            deallocation_calls: None,
            reallocation_calls: None,
            failed_allocation_calls: None,
            allocated_bytes: None,
            deallocated_bytes: None,
            live_bytes_before: None,
            live_bytes_after: None,
            peak_live_bytes_before: None,
            peak_live_bytes_after: None,
        });
        let envelope = aggregate(&[sample], "warm", &[10]).unwrap();
        let allocation = envelope.allocation.unwrap();
        assert_eq!(allocation.status, MetricStatus::Overflow);
        assert!(allocation.allocation_calls.values.is_none());
        let json = serde_json::to_value(allocation).unwrap();
        assert_eq!(json["status"], "overflow");
        assert!(json["allocation_calls"].get("values").is_none());
    }

    #[test]
    fn exact_cardinality_is_required() {
        let samples = vec![sample(0, "warm", 10, Some(metrics()))];
        let error = aggregate(&samples, "warm", &[10, 20]).unwrap_err();
        assert!(error.to_string().contains("sample count"));
    }

    #[test]
    fn duplicate_sample_identity_fails_closed() {
        let samples = vec![
            sample(0, "warm", 10, Some(metrics())),
            sample(0, "warm", 20, Some(metrics())),
        ];
        let error = aggregate(&samples, "warm", &[10, 20]).unwrap_err();
        assert!(error.to_string().contains("sample identity is not unique"));
    }

    #[test]
    fn measured_zero_is_not_unavailable() {
        let mut measured = sample(0, "warm", 10, Some(metrics()));
        measured.logical_read_calls = 0;
        let envelope = aggregate(&[measured], "warm", &[10]).unwrap();
        assert_eq!(
            envelope.source.logical_read_calls.status,
            MetricStatus::Measured
        );
        assert_eq!(envelope.source.logical_read_calls.values, Some(vec![0]));
        assert_eq!(
            envelope.source.logical_read_pattern.values,
            Some(vec![ReadPattern::Sequential])
        );
        assert_eq!(
            envelope.source.compressed_bytes.status,
            MetricStatus::Unavailable
        );
        assert_eq!(
            envelope.source.decompressed_bytes.status,
            MetricStatus::Unavailable
        );
        assert_eq!(
            envelope.source.recompressed_bytes.status,
            MetricStatus::Unavailable
        );
        assert_eq!(envelope.process.status, MetricStatus::Measured);
        assert_eq!(envelope.process.rchar.values, Some(vec![11]));
        assert_eq!(
            envelope.process.rchar.scope,
            "child_process_interval_delta_including_procfs_probe_overhead"
        );
        assert_eq!(envelope.process.wchar.values, Some(vec![12]));
        assert_eq!(envelope.process.read_bytes.values, Some(vec![13]));
        assert_eq!(envelope.process.write_bytes.values, Some(vec![14]));
        assert_eq!(
            envelope.process.cancelled_write_bytes.values,
            Some(vec![15])
        );
        assert_eq!(envelope.process.syscr.values, Some(vec![16]));
        assert_eq!(envelope.process.syscw.values, Some(vec![17]));
        let json = serde_json::to_value(envelope).unwrap();
        assert_eq!(
            json["source"]["logical_read_calls"]["values"],
            Value::from(vec![0])
        );
    }

    #[test]
    fn unavailable_procfs_metrics_omit_values_and_keep_status() {
        let envelope = aggregate(&[sample(0, "warm", 10, None)], "warm", &[10]).unwrap();
        assert_eq!(envelope.process.status, MetricStatus::Unavailable);
        assert!(envelope.process.user_cpu_ticks.values.is_none());
        assert!(envelope.process.rchar.values.is_none());
        assert!(envelope.process.wchar.values.is_none());
        assert!(envelope.process.read_bytes.values.is_none());
        assert!(envelope.process.write_bytes.values.is_none());
        assert!(envelope.process.cancelled_write_bytes.values.is_none());
        assert!(envelope.process.syscr.values.is_none());
        assert!(envelope.process.syscw.values.is_none());
        let json = serde_json::to_value(envelope).unwrap();
        assert_eq!(json["process"]["status"], "unavailable");
        assert!(json["process"]["user_cpu_ticks"].get("values").is_none());
    }

    #[test]
    fn in_process_observations_publish_aligned_resource_vectors_and_json() {
        let observations = vec![
            InProcessObservation {
                elapsed_ns: 30,
                process_metrics: Some(metrics()),
                allocation_metrics: Some(measured_allocation_sample(7)),
            },
            InProcessObservation {
                elapsed_ns: 10,
                process_metrics: Some(metrics()),
                allocation_metrics: Some(measured_allocation_sample(5)),
            },
        ];
        let envelope = from_in_process_observations(
            &observations,
            SinkObservation {
                accepted_bytes: 100,
                write_calls: 4,
                largest_write: 64,
                bytes_0: 0,
                bytes_1_to_512: 4,
                bytes_513_to_4096: 0,
                bytes_4097_to_16384: 0,
                bytes_16385_to_65536: 0,
                bytes_over_65536: 0,
            },
        )
        .unwrap();

        assert_eq!(envelope.sample_count, 2);
        assert_eq!(envelope.sample_indices, vec![1, 0]);
        assert_eq!(envelope.latency_claim, "comparable_timed_operation");
        assert_eq!(
            envelope.source.counter_scope,
            "not_applicable_in_process_sink"
        );
        assert_eq!(envelope.source.status, MetricStatus::NotApplicable);
        assert_eq!(envelope.process.status, MetricStatus::Measured);
        assert_eq!(
            envelope.process.user_cpu_ticks.scope,
            "procfs_in_process_operation_delta_including_procfs_probe_overhead"
        );
        assert_eq!(
            envelope.process.rchar.scope,
            "procfs_in_process_operation_delta_including_procfs_probe_overhead"
        );
        assert_eq!(
            envelope
                .allocation
                .as_ref()
                .unwrap()
                .allocation_calls
                .values,
            Some(vec![5, 7])
        );
        assert_eq!(envelope.sink.accepted_bytes.values, Some(vec![100, 100]));
        let json = serde_json::to_value(&envelope).unwrap();
        assert_eq!(json["process"]["status"], "measured");
        assert_eq!(
            json["allocation"]["allocation_calls"]["values"],
            serde_json::json!([5, 7])
        );
        assert_eq!(json["sink"]["write_status"], "measured");
    }

    #[test]
    fn in_process_observations_mark_procfs_unavailable_without_values() {
        let observations = vec![
            InProcessObservation {
                elapsed_ns: 20,
                process_metrics: None,
                allocation_metrics: None,
            },
            InProcessObservation {
                elapsed_ns: 10,
                process_metrics: None,
                allocation_metrics: None,
            },
        ];
        let envelope = from_in_process_observations(
            &observations,
            SinkObservation {
                accepted_bytes: 1,
                write_calls: 1,
                largest_write: 1,
                bytes_0: 0,
                bytes_1_to_512: 1,
                bytes_513_to_4096: 0,
                bytes_4097_to_16384: 0,
                bytes_16385_to_65536: 0,
                bytes_over_65536: 0,
            },
        )
        .unwrap();

        assert_eq!(envelope.process.status, MetricStatus::Unavailable);
        assert!(envelope.process.user_cpu_ticks.values.is_none());
        assert!(envelope.process.rchar.values.is_none());
        assert!(envelope.allocation.is_none());
        let json = serde_json::to_value(envelope).unwrap();
        assert_eq!(json["process"]["status"], "unavailable");
        assert!(json["process"]["user_cpu_ticks"].get("values").is_none());
    }

    #[test]
    fn measured_source_pattern_is_required_for_each_sample() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.logical_read_pattern = None;
        let error = aggregate(&[sample], "warm", &[10]).unwrap_err();
        assert!(error.to_string().contains("logical read pattern"));
    }

    #[test]
    fn asymmetric_option_fails_closed() {
        let samples = vec![
            sample(0, "warm", 10, Some(metrics())),
            sample(1, "warm", 20, None),
        ];
        let error = aggregate(&samples, "warm", &[10, 20]).unwrap_err();
        assert!(error.to_string().contains("process_metrics"));
    }

    #[test]
    fn status_and_max_value_serialize_without_null_numeric_placeholders() {
        let measured = MetricVector::measured(vec![u64::MAX], "test");
        let unavailable = MetricVector::absent(MetricStatus::Unavailable, "test");
        let json = serde_json::to_value((measured, unavailable)).unwrap();
        assert_eq!(json[0]["values"][0], u64::MAX);
        assert_eq!(json[1]["status"], "unavailable");
        assert!(json[1].get("values").is_none());
    }

    #[test]
    fn not_applicable_scope_omits_zero_read_vectors() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.logical_read_counter_scope = "not_applicable_eager_docx".to_owned();
        sample.logical_read_calls = 0;
        let envelope = aggregate(&[sample], "warm", &[10]).unwrap();
        assert_eq!(envelope.source.status, MetricStatus::NotApplicable);
        assert!(envelope.source.logical_read_calls.values.is_none());
        assert!(envelope.source.logical_read_pattern.values.is_none());
    }

    #[test]
    fn immutable_owned_slice_scope_is_not_applicable() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.logical_read_counter_scope = "not_applicable_immutable_owned_slice".to_owned();
        sample.logical_read_calls = 0;
        sample.logical_read_requested_bytes = 0;
        sample.logical_read_bytes = 0;
        let envelope = aggregate(&[sample], "warm", &[10]).unwrap();
        assert_eq!(envelope.source.status, MetricStatus::NotApplicable);
        assert!(envelope.source.logical_read_calls.values.is_none());
        assert!(envelope.source.logical_read_pattern.values.is_none());
        assert_eq!(
            envelope.source.compressed_bytes.status,
            MetricStatus::NotApplicable
        );
    }

    #[test]
    fn sink_observation_is_accepted_boundary_and_aligned() {
        let observation = SinkObservation {
            accepted_bytes: 70_000,
            write_calls: 6,
            largest_write: 65_537,
            bytes_0: 1,
            bytes_1_to_512: 1,
            bytes_513_to_4096: 1,
            bytes_4097_to_16384: 1,
            bytes_16385_to_65536: 1,
            bytes_over_65536: 1,
        };
        let envelope = from_sink_observation(2, observation).unwrap();
        assert_eq!(envelope.sample_count, 2);
        assert_eq!(envelope.sample_indices, vec![0, 1]);
        assert_eq!(envelope.latency_claim, "comparable_timed_operation");
        assert_eq!(envelope.sink.write_status, MetricStatus::Measured);
        assert_eq!(
            envelope.sink.accepted_bytes.values,
            Some(vec![70_000, 70_000])
        );
        assert_eq!(envelope.sink.write_calls.values, Some(vec![6, 6]));
        assert_eq!(
            envelope.sink.largest_write.values,
            Some(vec![65_537, 65_537])
        );
        assert_eq!(
            envelope.sink.write_size_buckets.bytes_over_65536.values,
            Some(vec![1, 1])
        );
        assert_eq!(
            envelope.sink.output_bytes.status,
            MetricStatus::NotApplicable
        );
        let json = serde_json::to_value(envelope).unwrap();
        assert!(
            json["sink"]["accepted_bytes"]["scope"]
                .as_str()
                .is_some_and(|scope| scope.contains("accepted"))
        );
        assert!(json["sink"].get("requested_bytes").is_none());
        assert!(json["sink"]["output_bytes"].get("values").is_none());
    }

    #[test]
    fn sink_observation_requires_nonempty_samples() {
        let observation = SinkObservation {
            accepted_bytes: 1,
            write_calls: 1,
            largest_write: 1,
            bytes_0: 0,
            bytes_1_to_512: 1,
            bytes_513_to_4096: 0,
            bytes_4097_to_16384: 0,
            bytes_16385_to_65536: 0,
            bytes_over_65536: 0,
        };
        let error = from_sink_observation(0, observation).unwrap_err();
        assert!(error.to_string().contains("zero samples"));
    }

    #[test]
    fn eager_opc_scope_omits_uninstrumented_fs_read_zeroes() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.logical_read_counter_scope = "not_applicable_eager_opc".to_owned();
        sample.logical_read_calls = 0;
        sample.logical_read_requested_bytes = 0;
        sample.logical_read_bytes = 0;
        let envelope = aggregate(&[sample], "warm", &[10]).unwrap();
        assert_eq!(envelope.source.status, MetricStatus::NotApplicable);
        assert!(envelope.source.logical_read_calls.values.is_none());
        assert!(
            envelope
                .source
                .logical_read_requested_bytes
                .values
                .is_none()
        );
        assert!(envelope.source.logical_read_returned_bytes.values.is_none());
        assert!(envelope.source.logical_read_pattern.values.is_none());
        assert_eq!(
            envelope.source.decompressed_bytes.status,
            MetricStatus::NotApplicable
        );
    }

    #[test]
    fn unknown_scope_is_unavailable_instead_of_a_measured_zero() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.logical_read_counter_scope = "future_scope".to_owned();
        sample.logical_read_calls = 0;
        let envelope = aggregate(&[sample], "warm", &[10]).unwrap();
        assert_eq!(envelope.source.status, MetricStatus::Unavailable);
        assert!(envelope.source.logical_read_calls.values.is_none());
        assert!(envelope.source.logical_read_pattern.values.is_none());
    }

    #[test]
    fn cfb_phase_metrics_are_aligned_and_fail_closed() {
        let phase = |elapsed_ns, calls, bytes| CfbPhaseSample {
            elapsed_ns,
            logical_read_calls: calls,
            logical_read_requested_bytes: bytes,
            logical_read_returned_bytes: bytes,
        };
        let mut first = sample(0, "warm", 20, Some(metrics()));
        first.cfb_phases = Some(CfbPhaseEvidence {
            open: phase(2, 3, 4),
            plan: phase(5, 6, 7),
            atomic_publication: phase(8, 9, 10),
        });
        let mut second = sample(1, "warm", 10, Some(metrics()));
        second.cfb_phases = Some(CfbPhaseEvidence {
            open: phase(11, 12, 13),
            plan: phase(14, 15, 16),
            atomic_publication: phase(17, 18, 19),
        });

        let mut samples = vec![first, second];
        let envelope = aggregate(&samples, "warm", &[20, 10]).unwrap();
        assert_eq!(envelope.cfb_phases.status, MetricStatus::Measured);
        assert_eq!(
            envelope.cfb_phases.open.elapsed_ns.values,
            Some(vec![11, 2])
        );
        assert_eq!(
            envelope
                .cfb_phases
                .atomic_publication
                .logical_read_returned_bytes
                .values,
            Some(vec![19, 10])
        );

        samples[1].cfb_phases = None;
        let error = aggregate(&samples, "warm", &[20, 10]).unwrap_err();
        assert!(error.to_string().contains("cfb_phases"));
    }
}
