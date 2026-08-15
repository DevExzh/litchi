//! Additive, operation-scoped metrics for the isolated filesystem selectors.
//!
//! The filesystem child already records the counters that can be collected
//! without changing the timed operation.  This module only aligns those
//! per-child records with the sorted `elapsed_ns.samples` vector.  It does
//! not infer allocations, copies, decompression, recompression, or a peak
//! from a process-wide measurement that cannot provide one.

use std::error::Error;

use serde::Serialize;

use crate::filesystem::SampleEvidence;

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

#[derive(Clone, Debug, Serialize)]
pub(crate) struct SourceMetrics {
    pub status: MetricStatus,
    pub counter_scope: String,
    pub logical_read_calls: MetricVector,
    pub logical_read_requested_bytes: MetricVector,
    pub logical_read_returned_bytes: MetricVector,
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
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct SinkMetrics {
    pub status: MetricStatus,
    /// Final output length observed after a save.  This is not a write-call or
    /// memory-copy count; the scope says explicitly that it is post-operation.
    pub output_bytes: MetricVector,
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
pub(crate) struct OperationMetrics {
    /// Number of values in every measured vector below.
    pub sample_count: usize,
    /// All vectors are ordered exactly like the case result's
    /// `elapsed_ns.samples` vector, after its deterministic sort.
    pub alignment: &'static str,
    pub source: SourceMetrics,
    pub process: ProcessMetrics,
    pub sink: SinkMetrics,
    pub publication: PublicationMetrics,
    pub materialization: MaterializationMetrics,
}

const ALIGNMENT: &str = "elapsed_ns.samples";
const SOURCE_SCOPE: &str = "operation_logical_read_at";
const PROCESS_SCOPE: &str = "procfs_operation_delta";
const CLOCK_SCOPE: &str = "procfs_after_sample_unit_factor";
const RSS_SCOPE: &str = "procfs_operation_delta_not_peak";
const HWM_SCOPE: &str = "process_lifetime_high_water_after_not_operation_peak";
const OUTPUT_SCOPE: &str = "post_operation_output_length_not_sink_write_volume";
const PUBLICATION_SCOPE: &str = "logical_publication_counter";
const MATERIALIZATION_SCOPE: &str = "logical_materialization_counter";

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
        "untimed_source_replay_only"
        | "not_applicable_eager_opc"
        | "not_applicable_eager_pptx"
        | "not_applicable_eager_docx" => MetricStatus::NotApplicable,
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
    let source = SourceMetrics {
        status: source_status,
        counter_scope: first_scope,
        logical_read_calls: source_values(|sample| sample.logical_read_calls),
        logical_read_requested_bytes: source_values(|sample| sample.logical_read_requested_bytes),
        logical_read_returned_bytes: source_values(|sample| sample.logical_read_bytes),
        max_concurrent_reads: source_values(|sample| sample.max_concurrent_reads),
    };

    let process = process_metrics(&selected)?;
    let sink = SinkMetrics {
        status: optional_status(
            &selected,
            |sample| sample.output_bytes.is_some(),
            "output_bytes",
        )?,
        output_bytes: optional_values(
            &selected,
            |sample| sample.output_bytes,
            "output_bytes",
            MetricStatus::NotApplicable,
            OUTPUT_SCOPE,
        )?,
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

    Ok(OperationMetrics {
        sample_count,
        alignment: ALIGNMENT,
        source,
        process,
        sink,
        publication,
        materialization,
    })
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
        return Ok(ProcessMetrics {
            status: MetricStatus::Unavailable,
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
    Ok(ProcessMetrics {
        status: MetricStatus::Measured,
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

    use super::{MetricStatus, MetricVector, aggregate};
    use crate::filesystem::{ColdAdvice, ReadSizeBuckets, SampleEvidence};

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
            logical_read_counter_scope: "timed_read_at".to_owned(),
            logical_read_calls: index as u64,
            logical_read_requested_bytes: 10,
            logical_read_bytes: 8,
            max_concurrent_reads: 1,
            logical_read_request_sizes: vec![10],
            logical_read_request_size_buckets: ReadSizeBuckets::default(),
            process_metrics,
            output_sha256: None,
            output_bytes: None,
            opc_materialized_parts: None,
            cfb_changed_spans: None,
            cfb_published_bytes: None,
            pptx_source_replay: None,
            docx_source_replay: None,
        }
    }

    fn metrics() -> crate::process_metrics::Delta {
        crate::process_metrics::Delta {
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
        assert_eq!(warm.source.logical_read_calls.values, Some(vec![1, 0]));
        let cold = aggregate(&samples, "cold-requested", &[40, 30]).unwrap();
        assert_eq!(cold.source.logical_read_calls.values, Some(vec![1, 0]));
    }

    #[test]
    fn exact_cardinality_is_required() {
        let samples = vec![sample(0, "warm", 10, Some(metrics()))];
        let error = aggregate(&samples, "warm", &[10, 20]).unwrap_err();
        assert!(error.to_string().contains("sample count"));
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
        assert_eq!(envelope.process.status, MetricStatus::Measured);
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
        let json = serde_json::to_value(envelope).unwrap();
        assert_eq!(json["process"]["status"], "unavailable");
        assert!(json["process"]["user_cpu_ticks"].get("values").is_none());
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
    }

    #[test]
    fn unknown_scope_is_unavailable_instead_of_a_measured_zero() {
        let mut sample = sample(0, "warm", 10, Some(metrics()));
        sample.logical_read_counter_scope = "future_scope".to_owned();
        sample.logical_read_calls = 0;
        let envelope = aggregate(&[sample], "warm", &[10]).unwrap();
        assert_eq!(envelope.source.status, MetricStatus::Unavailable);
        assert!(envelope.source.logical_read_calls.values.is_none());
    }
}
