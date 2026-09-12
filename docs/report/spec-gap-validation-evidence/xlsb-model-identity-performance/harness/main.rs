//! Process-isolated smoke runner for the source-backed XLSB identity profile.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::print_stdout,
    reason = "the opt-in profile emits machine-readable receipts"
)]

mod adapter;
mod support;

#[global_allocator]
static GLOBAL_ALLOCATOR: support::CountingAllocator = support::CountingAllocator;

use std::env;
use std::process::ExitCode;
use std::time::Instant;

use adapter::RunResult;
use serde_json::{Value, json};
use support::{Snapshot, reset};

const BASELINE_COMMIT: &str = "1b0d4864804d666aaf4ae5039bd6300996ca32a5";
const NEUTRAL_COMMIT: &str = "4f53be2d31e14b215d6eeaee56d11eefeae94107";

fn parse_args() -> Result<(String, usize, usize), String> {
    let mut arguments = env::args().skip(1);
    let mut lane = None;
    let mut warmup = 0;
    let mut samples = 1;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => return Err(String::from("help")),
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

fn usage() {
    eprintln!("usage: xlsb-model-identity-profile --lane NAME [--warmup N] [--samples N]");
    eprintln!("lanes: {}", adapter::LANES.join(", "));
}

fn phase_json(result: &RunResult) -> Value {
    json!({
        "open_ns": result.phases.open_ns,
        "stage_ns": result.phases.stage_ns,
        "commit_ns": result.phases.commit_ns,
        "save_ns": result.phases.save_ns,
        "reopen_ns": result.phases.reopen_ns,
        "inverse_ns": result.phases.inverse_ns,
        "validation_ns": result.phases.validation_ns,
    })
}

fn sample_json(result: &RunResult, elapsed_ns: u64, delta: support::Delta) -> Value {
    json!({
        "elapsed_ns": elapsed_ns,
        "phases": phase_json(result),
        "actual_success": result.actual_success,
        "semantic_ok": result.semantic_ok,
        "opaque_ok": result.opaque_ok,
        "preservation": result.preservation.as_ref().map(|check| json!({
            "all_parts_equal": check.all_parts_equal,
            "unchanged_parts_equal": check.unchanged_parts_equal,
            "relationships_equal": check.relationships_equal,
            "content_types_equal": check.content_types_equal,
            "inner_all_equal": check.inner_all_equal,
            "inner_unchanged_equal": check.inner_unchanged_equal,
            "inner_member_count": check.inner_member_count,
            "relationship_owner_count": check.relationship_owner_count,
            "content_types_bytes": check.content_types_bytes,
        })),
        "source_unchanged": result.source_unchanged,
        "exact_inverse_ok": result.exact_inverse_ok,
        "exact_cap_ok": result.exact_cap_ok,
        "one_under_cap_refused": result.one_under_cap_refused,
        "source_bytes": result.source_bytes,
        "candidate_bytes": result.candidate_bytes,
        "output_bytes": result.output_bytes,
        "direct_allocated_bytes": delta.direct_bytes,
        "realloc_old_bytes": delta.realloc_old_bytes,
        "realloc_new_bytes": delta.realloc_new_bytes,
        "deallocated_bytes": delta.deallocated_bytes,
        "requested_alloc_bytes": delta.requested_bytes(),
        "live_before": delta.live_before,
        "live_after": delta.live_after,
        "peak_live_delta": delta.peak_live_delta,
        "allocation_calls": delta.allocation_calls,
        "reallocation_calls": delta.reallocation_calls,
        "deallocation_calls": delta.deallocation_calls,
        "allocation_failed": delta.allocation_failed,
        "alloc_balance_ok": delta.balanced(),
        "alloc_invalid": delta.invalid,
        "error": result.error.as_ref().map(|error| json!({
            "class": error.class,
            "message": error.message,
            "typed_match": error.typed_match,
        })),
    })
}

fn run(lane: &str, warmup: usize, samples: usize) -> Result<Value, String> {
    let fixture = adapter::fixture_for_lane(lane)?;
    let expected_success = adapter::expected_success(lane);
    reset();
    for _ in 0..warmup {
        let result = adapter::run_once(lane, &fixture)?;
        validate_result(lane, expected_success, &result)?;
    }
    let mut measured = Vec::with_capacity(samples);
    for _ in 0..samples {
        support::reset();
        let before = Snapshot::now();
        let started = Instant::now();
        let result = adapter::run_once(lane, &fixture)?;
        let elapsed_ns = started.elapsed().as_nanos().try_into().unwrap_or(u64::MAX);
        let delta = before.delta(Snapshot::now());
        validate_result(lane, expected_success, &result)?;
        if !delta.balanced() || delta.invalid || delta.allocation_failed != 0 {
            return Err(format!("allocator gate failed for lane {lane}"));
        }
        measured.push(sample_json(&result, elapsed_ns, delta));
    }
    Ok(json!({
        "schema": "xlsb-model-identity-profile-v1-smoke",
        "baseline_commit": BASELINE_COMMIT,
        "neutral_baseline_commit": NEUTRAL_COMMIT,
        "lane": lane,
        "warmup": warmup,
        "sample_count": samples,
        "expected_success": expected_success,
        "fixture_kind": "synthetic_complete_xldm140",
        "scale": fixture.scale,
        "table_count": fixture.table_count,
        "relationship_count": fixture.relationship_count,
        "input_bytes": fixture.bytes.len(),
        "input_sha256": fixture.input_sha256,
        "input_hash_fnv1a64": fixture.input_fnv1a64,
        "source_backed_api": true,
        "native_acceptance_claim": false,
        "samples": measured,
    }))
}

fn validate_result(lane: &str, expected_success: bool, result: &RunResult) -> Result<(), String> {
    if result.actual_success != expected_success {
        return Err(format!("unexpected success state for lane {lane}"));
    }
    if !result.source_unchanged || result.opaque_ok == Some(false) || !result.semantic_ok {
        return Err(format!("correctness gate failed for lane {lane}"));
    }
    let neutral = matches!(lane, "neutral_open_tiny" | "neutral_open_relationship");
    if !neutral && result.preservation.is_none() {
        return Err(format!(
            "preservation manifest was not checked for lane {lane}"
        ));
    }
    if !expected_success {
        let Some(error) = result.error.as_ref() else {
            return Err(format!("refusal lane {lane} has no typed error"));
        };
        if !error.typed_match {
            return Err(format!("refusal lane {lane} had {}", error.class));
        }
    } else if result.error.is_some() {
        return Err(format!("successful lane {lane} returned an error"));
    }
    Ok(())
}

fn main() -> ExitCode {
    let (lane, warmup, samples) = match parse_args() {
        Ok(arguments) => arguments,
        Err(error) if error == "help" => {
            usage();
            return ExitCode::SUCCESS;
        },
        Err(error) => {
            usage();
            eprintln!("{error}");
            return ExitCode::from(2);
        },
    };
    match run(&lane, warmup, samples) {
        Ok(receipt) => {
            println!(
                "{}",
                serde_json::to_string_pretty(&receipt).expect("serialize smoke receipt")
            );
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("smoke correctness failure: {error}");
            ExitCode::from(1)
        },
    }
}
