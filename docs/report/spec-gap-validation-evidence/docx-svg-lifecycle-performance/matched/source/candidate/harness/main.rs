//! Process-isolated allocator and runtime evidence for the source-backed DOCX
//! SVG lifecycle.  The binary lives under the evidence tree and is never a
//! production dependency.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
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
use support::{AllocSnapshot, begin_window};

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;
const BASELINE_COMMIT: &str = "892441d95db29da4351390716ef5c65b4c7c97de";
const OPC_BASELINE_LABEL: &str = "OPC57680dc86";

fn usage() {
    println!(
        "usage: docx-svg-lifecycle-profile --lane NAME [--warmup N] [--samples N]\n\n{} lanes are wired to the committed DOCX source-backed API",
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
            }
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
            "--warmup" => {
                warmup = arguments
                    .next()
                    .ok_or("--warmup requires a value")?
                    .parse()
                    .map_err(|_| String::from("--warmup must be an integer"))?;
            }
            "--samples" => {
                samples = arguments
                    .next()
                    .ok_or("--samples requires a value")?
                    .parse()
                    .map_err(|_| String::from("--samples must be an integer"))?;
            }
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

fn phase_json(result: &RunResult) -> Value {
    json!({
        "capture_ns": result.phases.capture_ns,
        "stage_ns": result.phases.stage_ns,
        "commit_ns": result.phases.commit_ns,
        "publish_ns": result.phases.publish_ns,
        "reopen_ns": result.phases.reopen_ns,
        "inverse_reopen_ns": result.phases.inverse_reopen_ns,
        "inverse_ns": result.phases.inverse_ns,
        "payload_ns": result.phases.payload_ns,
        "validation_ns": result.phases.validation_ns,
        "readback_ns": result.phases.readback_ns,
    })
}

fn sample_json(
    result: &RunResult,
    elapsed_ns: u64,
    delta: support::AllocDelta,
) -> Value {
    json!({
        "elapsed_ns": elapsed_ns,
        "phases": phase_json(result),
        "actual_success": result.actual_success,
        "semantic_ok": result.semantic_ok,
        "opaque_ok": result.opaque_ok,
        "exact_inverse_ok": result.exact_inverse_ok,
        "lazy_media_cold_ok": result.lazy_media_cold_ok,
        "source_readback_physical_ok": result.source_readback_physical_ok,
        "source_readback_metadata_ok": result.source_readback_metadata_ok,
        "output_bytes": result.output_bytes,
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
        "error": result.error.as_ref().map(|error| json!({
            "class": error.class,
            "message": error.message,
            "typed_match": error.typed_match,
        })),
    })
}

fn run(lane: &str, warmup: usize, samples: usize) -> Result<Value, String> {
    let fixture = adapter::fixture_for_lane(lane).map_err(|error| error.to_string())?;
    let expected_success = adapter::expected_success(lane);
    support::reset_process_counters();
    for _ in 0..warmup {
        let result = adapter::run_once(lane, &fixture).map_err(|error| error.to_string())?;
        if result.actual_success != expected_success
            || (!expected_success && result.error.is_none())
        {
            return Err(format!("warm-up correctness gate failed for lane {lane}"));
        }
    }
    let mut measured = Vec::with_capacity(samples);
    let mut failed = false;
    for _ in 0..samples {
        begin_window();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let result = adapter::run_once(lane, &fixture).map_err(|error| error.to_string())?;
        let elapsed_ns = started.elapsed().as_nanos().try_into().unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let delta = before.delta(after);
        if !delta.balanced() || delta.invalid || delta.failed != 0 {
            failed = true;
        }
        if result.actual_success != expected_success
            || (!expected_success && result.error.is_none())
        {
            failed = true;
        }
        measured.push(sample_json(&result, elapsed_ns, delta));
    }
    let input_hash = fixture.input_hash_fnv1a64;
    let value = json!({
        "schema": "docx-svg-lifecycle-profile-v1",
        "baseline_commit": BASELINE_COMMIT,
        "opc_baseline_label": OPC_BASELINE_LABEL,
        "lane": lane,
        "warmup": warmup,
        "sample_count": samples,
        "expected_success": expected_success,
        "owners": fixture.owners,
        "input_bytes": fixture.input_bytes,
        "input_sha256": fixture.input_sha256,
        "input_hash_fnv1a64": input_hash,
        "fixture_kind": if fixture.native { "native_submodule_copy" } else { "synthetic" },
        "source_backed_api": true,
        "samples": measured,
    });
    if failed {
        return Err(format!("correctness or allocator gate failed for lane {lane}"));
    }
    Ok(value)
}

fn main() -> ExitCode {
    let arguments = match parse_args() {
        Ok(arguments) => arguments,
        Err(error) if error == "help" => return ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{error}");
            return ExitCode::from(2);
        }
    };
    let (lane, warmup, samples) = arguments;
    match run(&lane, warmup, samples) {
        Ok(receipt) => {
            println!(
                "{}",
                serde_json::to_string_pretty(&receipt).expect("serialize profile receipt")
            );
            ExitCode::SUCCESS
        }
        Err(error) => {
            eprintln!("profile fixture or correctness setup failed: {error}");
            ExitCode::from(1)
        }
    }
}
