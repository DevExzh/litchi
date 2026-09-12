//! Process-isolated receipt driver for the bounded stylesWithEffects smoke.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::print_stdout,
    reason = "the opt-in evidence harness emits machine-readable receipts"
)]

mod adapter;
mod support;

#[global_allocator]
static GLOBAL_ALLOCATOR: support::CountingAllocator = support::CountingAllocator;

use std::env;
use std::process::ExitCode;

use adapter::{CapEvidence, ErrorReceipt, PackageMetrics, RunResult};
use serde_json::{Value, json};
use std::time::Instant;
use support::{AllocSnapshot, begin_window};

const DEFAULT_WARMUP: usize = 0;
const DEFAULT_SAMPLES: usize = 1;
const SOURCE_COMMIT: &str = "d000d977b99e03f8542c7dae74acf767a91b1feb";
const OPC_SOURCE_LABEL: &str = "styles-effects-source-committed";

fn usage() {
    println!(
        "usage: docx-styles-effects-smoke --lane NAME [--warmup N] [--samples N]\n\n{} bounded public-API lanes",
        adapter::LANES.len()
    );
}

fn parse_args() -> Result<(String, usize, usize), String> {
    let mut arguments = env::args().skip(1);
    let mut lane = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                usage();
                return Err(String::from("help"));
            },
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
            "--warmup" => {
                warmup = arguments
                    .next()
                    .ok_or("--warmup requires a value")?
                    .parse()
                    .map_err(|_| String::from("--warmup must be an integer"))?;
            },
            "--samples" => {
                samples = arguments
                    .next()
                    .ok_or("--samples requires a value")?
                    .parse()
                    .map_err(|_| String::from("--samples must be an integer"))?;
            },
            unknown => return Err(format!("unknown argument: {unknown}")),
        }
    }
    let lane = lane.ok_or("--lane is required")?;
    if !adapter::is_known_lane(&lane) {
        return Err(format!("unknown lane: {lane}"));
    }
    if samples == 0 {
        return Err(String::from("--samples must be nonzero"));
    }
    Ok((lane, warmup, samples))
}

fn metrics_json(metrics: PackageMetrics) -> Value {
    json!({
        "parts": metrics.parts,
        "total_part_bytes": metrics.total_part_bytes,
        "total_relationships": metrics.total_relationships,
        "relationship_parts": metrics.relationship_parts,
        "relationship_graph_nodes": metrics.relationship_graph_nodes,
        "relationship_xml_bytes": metrics.relationship_xml_bytes,
        "relationship_xml_events": metrics.relationship_xml_events,
    })
}

fn phases_json(result: &RunResult) -> Value {
    json!({
        "capture_ns": result.phases.capture_ns,
        "snapshot_ns": result.phases.snapshot_ns,
        "stage_ns": result.phases.stage_ns,
        "commit_ns": result.phases.commit_ns,
        "publish_ns": result.phases.publish_ns,
        "reopen_ns": result.phases.reopen_ns,
        "inverse_reopen_ns": result.phases.inverse_reopen_ns,
        "inverse_ns": result.phases.inverse_ns,
        "projection_ns": result.phases.projection_ns,
        "opaque_ns": result.phases.opaque_ns,
        "graph_ns": result.phases.graph_ns,
        "readback_ns": result.phases.readback_ns,
        "validation_ns": result.phases.validation_ns,
    })
}

fn error_json(error: &ErrorReceipt) -> Value {
    json!({
        "class": error.class,
        "variant": error.variant,
        "message": error.message,
        "typed_match": error.typed_match,
        "resource": error.resource,
        "actual": error.actual,
        "maximum": error.maximum,
    })
}

fn cap_evidence_json(value: &CapEvidence) -> Value {
    json!({
        "applicable": value.applicable,
        "source_metrics": metrics_json(value.source_metrics),
        "projected_metrics": metrics_json(value.projected_metrics),
        "exact_fit_ok": value.exact_fit_ok,
        "exact_opaque_ok": value.exact_opaque_ok,
        "source_unrelated_member_digest": value.source_unrelated_member_digest,
        "exact_unrelated_member_digest": value.exact_unrelated_member_digest,
        "under_refused_ok": value.under_refused_ok,
        "commit_stage_checked": value.commit_stage_checked,
        "refusal": value.refusal.as_ref().map(error_json),
        "commit_refusal": value.commit_refusal.as_ref().map(error_json),
    })
}

fn sample_json(result: &RunResult, elapsed_ns: u64, delta: support::AllocDelta) -> Value {
    json!({
        "elapsed_ns": elapsed_ns,
        "phases": phases_json(result),
        "actual_success": result.actual_success,
        "ingress_refusal": result.ingress_refusal,
        "no_output_ok": result.no_output_ok,
        "semantic_ok": result.semantic_ok,
        "opaque_ok": result.opaque_ok,
        "exact_inverse_ok": result.exact_inverse_ok,
        "source_readback_physical_ok": result.source_readback_physical_ok,
        "source_readback_metadata_ok": result.source_readback_metadata_ok,
        "output_bytes": result.output_bytes,
        "input_sha256": result.input_sha256,
        "output_sha256": result.output_sha256,
        "input_metrics": metrics_json(result.input_metrics),
        "output_metrics": result.output_metrics.map(metrics_json),
        "input_member_digest": result.input_member_digest,
        "output_member_digest": result.output_member_digest,
        "cap_exact_fit_ok": result.cap_exact_fit_ok,
        "cap_under_refused_ok": result.cap_under_refused_ok,
        "cap_refusal": result.cap_refusal.as_ref().map(error_json),
        "cap_commit_refusal": result.cap_commit_refusal.as_ref().map(error_json),
        "cap_source_metrics": result.cap_source_metrics.map(metrics_json),
        "cap_projected_metrics": result.cap_projected_metrics.map(metrics_json),
        "cap_existing": result.cap_existing.as_ref().map(cap_evidence_json),
        "direct_allocated_bytes": delta.direct,
        "realloc_old_bytes": delta.realloc_old,
        "realloc_new_bytes": delta.realloc_new,
        "deallocated_bytes": delta.deallocated,
        "requested_alloc_bytes": delta.requested(),
        "live_before": delta.live_before,
        "live_after": delta.live_after,
        "peak_live_delta": delta.peak_delta,
        "allocation_calls": delta.calls,
        "reallocation_calls": delta.realloc_calls,
        "deallocation_calls": delta.dealloc_calls,
        "allocation_failed": delta.failed,
        "alloc_balance_ok": delta.balanced(),
        "alloc_invalid": delta.invalid,
        "error": result.error.as_ref().map(error_json),
    })
}

fn verify_result(result: &RunResult, lane: &str, expected_success: bool) -> Result<(), String> {
    if result.actual_success != expected_success {
        return Err(format!(
            "lane {lane} actual_success={} expected_success={expected_success}",
            result.actual_success
        ));
    }
    if !result.semantic_ok || !result.opaque_ok {
        let error = result
            .error
            .as_ref()
            .map(|value| {
                format!(
                    " variant={} typed={} message={}",
                    value.variant, value.typed_match, value.message
                )
            })
            .unwrap_or_default();
        return Err(format!(
            "lane {lane} semantic={} opaque={} physical={:?} metadata={:?} smoke gate failed{error}",
            result.semantic_ok,
            result.opaque_ok,
            result.source_readback_physical_ok,
            result.source_readback_metadata_ok,
        ));
    }
    if expected_success && result.error.is_some() {
        return Err(format!(
            "lane {lane} unexpectedly returned an error receipt"
        ));
    }
    if !expected_success {
        let Some(error) = result.error.as_ref() else {
            return Err(format!("lane {lane} expected a typed refusal"));
        };
        if !error.typed_match {
            return Err(format!("lane {lane} returned an unexpected typed refusal"));
        }
        if let Some(physical) = result.source_readback_physical_ok
            && (!physical || result.source_readback_metadata_ok != Some(true))
        {
            return Err(format!("lane {lane} refusal source readback failed"));
        }
    }
    if lane.starts_with("cap_")
        && (result.cap_exact_fit_ok != Some(true) || result.cap_under_refused_ok != Some(true))
    {
        return Err(format!("lane {lane} cap boundary gate failed"));
    }
    Ok(())
}

fn run(lane: &str, warmup: usize, samples: usize) -> Result<Value, String> {
    let fixture = adapter::fixture_for_lane(lane).map_err(|error| error.to_string())?;
    let expected_success = adapter::expected_success(lane);
    support::reset_process_counters();
    for _ in 0..warmup {
        let result = adapter::run_once(lane, &fixture).map_err(|error| error.to_string())?;
        verify_result(&result, lane, expected_success)?;
    }
    let mut measured = Vec::with_capacity(samples);
    for _ in 0..samples {
        begin_window();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let result = adapter::run_once(lane, &fixture).map_err(|error| error.to_string())?;
        let elapsed_ns = started.elapsed().as_nanos().try_into().unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let delta = before.delta(after);
        verify_result(&result, lane, expected_success)?;
        if !delta.balanced() || delta.invalid || delta.failed != 0 {
            return Err(format!(
                "lane {lane} allocator equation/failure gate failed"
            ));
        }
        measured.push(sample_json(&result, elapsed_ns, delta));
    }
    Ok(json!({
        "schema": "docx-styles-effects-smoke-v1",
        "source_commit": SOURCE_COMMIT,
        "opc_source_label": OPC_SOURCE_LABEL,
        "lane": lane,
        "warmup": warmup,
        "sample_count": samples,
        "expected_success": expected_success,
        "fixture": fixture.name,
        "fixture_native": fixture.native,
        "fixture_signed": fixture.signed,
        "fixture_main_present": fixture.main_present,
        "fixture_glossary_present": fixture.glossary_present,
        "fixture_expected_package_sha256": fixture.expected_package_sha256,
        "source_backed_api": true,
        "samples": measured,
    }))
}

fn main() -> ExitCode {
    let arguments = match parse_args() {
        Ok(arguments) => arguments,
        Err(error) if error == "help" => return ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{error}");
            return ExitCode::from(2);
        },
    };
    let (lane, warmup, samples) = arguments;
    match run(&lane, warmup, samples) {
        Ok(receipt) => {
            println!(
                "{}",
                serde_json::to_string_pretty(&receipt).expect("serialize smoke receipt")
            );
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("smoke fixture or correctness setup failed: {error}");
            ExitCode::from(1)
        },
    }
}
