//! Bounded process-isolated `stylesWithEffects` profile scaffold.
//!
//! The correctness adapter remains the authority for the complete 52-lane
//! smoke. This binary exposes only the reviewed initial success matrix and
//! records one real public API operation per sample. The shell runner is
//! freeze-gated; this binary is not a native Office renderer or an A/B tool.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in evidence harness emits machine-readable receipts"
)]
#![recursion_limit = "512"]

#[allow(dead_code)]
#[path = "profile_adapter.rs"]
mod smoke_adapter;
#[path = "../harness/support.rs"]
mod support;
mod synthetic;

#[global_allocator]
static GLOBAL_ALLOCATOR: support::CountingAllocator = support::CountingAllocator;

use std::env;
use std::process::ExitCode;
use std::time::Instant;

use serde_json::{Value, json};
use smoke_adapter::{Fixture, RunResult};
use support::{AllocSnapshot, begin_window};

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;
const SOURCE_COMMIT: &str = "d000d977b99e03f8542c7dae74acf767a91b1feb";
const PROFILE_SCHEMA: &str = "docx-styles-effects-profile-scaffold-v2";

const LANES: &[&str] = &[
    "capture_bug_main",
    "capture_bug_glossary",
    "capture_signed_main",
    "capture_complex_main",
    "capture_glossary_main",
    "capture_glossary_glossary",
    "capture_synthetic_main",
    "capture_synthetic_glossary",
    "projection_main",
    "projection_glossary",
    "noop_main",
    "noop_glossary",
    "replace_main",
    "replace_glossary",
    "remove_main",
    "remove_glossary",
    "add_main_absent",
    "inverse_replace_main",
    "inverse_remove_main",
    "independent_main",
    "independent_glossary",
];

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Scale {
    Native,
    KiB64,
    MiB1,
}

impl Scale {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "native" => Ok(Self::Native),
            "64k" => Ok(Self::KiB64),
            "1m" => Ok(Self::MiB1),
            _ => Err(format!("unknown scale {value}; use native, 64k, or 1m")),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Native => "native",
            Self::KiB64 => "64k",
            Self::MiB1 => "1m",
        }
    }
}

fn usage() {
    println!(
        "usage: docx-styles-effects-profile --lane NAME --scale native|64k|1m [--warmup N] [--samples N]\n\n{} initial success lanes; the runner supplies the freeze gate",
        LANES.len()
    );
}

fn parse_args() -> Result<(String, Scale, usize, usize), String> {
    let mut arguments = env::args().skip(1);
    let mut lane = None;
    let mut scale = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => {
                usage();
                return Err(String::from("help"));
            },
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
            "--scale" => {
                scale = Some(Scale::parse(
                    &arguments.next().ok_or("--scale requires a value")?,
                )?);
            },
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
    if !LANES.contains(&lane.as_str()) {
        return Err(format!("unknown initial profile lane: {lane}"));
    }
    let scale = scale.ok_or("--scale is required")?;
    if samples == 0 {
        return Err(String::from("--samples must be nonzero"));
    }
    if scale != Scale::Native
        && !matches!(
            lane.as_str(),
            "capture_synthetic_main"
                | "capture_synthetic_glossary"
                | "projection_main"
                | "projection_glossary"
                | "replace_main"
                | "replace_glossary"
        )
    {
        return Err(format!(
            "scale {} is not in the initial matrix for lane {lane}",
            scale.label()
        ));
    }
    Ok((lane, scale, warmup, samples))
}

fn smoke_lane(lane: &str) -> &'static str {
    match lane {
        "capture_bug_main" | "capture_synthetic_main" => "native_capture_bug_main",
        "capture_bug_glossary" | "capture_synthetic_glossary" => "native_capture_bug_glossary",
        "capture_signed_main" => "native_capture_signed_main",
        "capture_complex_main" => "native_capture_complex_main",
        "capture_glossary_main" => "native_capture_glossary_main",
        "capture_glossary_glossary" => "native_capture_glossary_glossary",
        "noop_main" => "source_noop_main",
        "noop_glossary" => "source_noop_glossary",
        "projection_main" => "projection_main",
        "projection_glossary" => "projection_glossary",
        "replace_main" => "replace_main",
        "replace_glossary" => "replace_glossary",
        "remove_main" => "remove_main",
        "remove_glossary" => "remove_glossary",
        "add_main_absent" => "add_main_absent",
        "inverse_replace_main" => "inverse_replace_main",
        "inverse_remove_main" => "inverse_remove_main",
        "independent_main" => "independent_main",
        "independent_glossary" => "independent_glossary",
        _ => unreachable!("lane was checked by parse_args"),
    }
}

fn fixture(lane: &str, scale: Scale) -> Result<Fixture, String> {
    let smoke = smoke_lane(lane);
    if scale == Scale::Native {
        return smoke_adapter::fixture_for_lane(smoke).map_err(|error| error.to_string());
    }
    if !matches!(
        lane,
        "capture_synthetic_main"
            | "capture_synthetic_glossary"
            | "projection_main"
            | "projection_glossary"
            | "replace_main"
            | "replace_glossary"
    ) {
        return Err(format!("synthetic scale is unsupported for lane {lane}"));
    }
    let manifest = env::var_os("DOCX_PROFILE_GENERATED_FIXTURES").ok_or_else(|| {
        String::from("synthetic profile requires DOCX_PROFILE_GENERATED_FIXTURES")
    })?;
    synthetic::fixture_from_manifest(
        std::path::Path::new(&manifest),
        smoke,
        lane.ends_with("glossary"),
        scale,
    )
    .map_err(|error| error.to_string())
}

fn metrics_json(metrics: smoke_adapter::PackageMetrics) -> Value {
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

fn member_digest_json(value: Option<&smoke_adapter::MemberDigest>) -> Value {
    value.map_or(Value::Null, |digest| {
        json!({
            "bytes": digest.length,
            "sha256": digest.sha256,
        })
    })
}

fn resource_metadata_json(value: Option<&smoke_adapter::ResourceMetadata>) -> Value {
    value.map_or(Value::Null, |metadata| {
        json!({
            "owner": metadata.owner,
            "member": metadata.member,
            "resource_bytes": metadata.resource_bytes,
            "resource_sha256": metadata.resource_sha256,
            "conformance": metadata.conformance,
            "style_count": metadata.style_count,
            "xml_events": metadata.xml_events,
            "xml_depth": metadata.xml_depth,
        })
    })
}

fn phases_json(result: &RunResult) -> Value {
    json!({
        "capture_ns": result.phases.capture_ns,
        "snapshot_ns": result.phases.snapshot_ns,
        "stage_ns": result.phases.stage_ns,
        "commit_ns": result.phases.commit_ns,
        "publish_ns": result.phases.publish_ns,
        "apply_ns": result.phases.apply_ns,
        "serialize_ns": result.phases.serialize_ns,
        "reopen_ns": result.phases.reopen_ns,
        "inverse_reopen_ns": result.phases.inverse_reopen_ns,
        "inverse_ns": result.phases.inverse_ns,
        "inverse_apply_ns": result.phases.inverse_apply_ns,
        "inverse_serialize_ns": result.phases.inverse_serialize_ns,
        "projection_ns": result.phases.projection_ns,
        "opaque_ns": result.phases.opaque_ns,
        "graph_ns": result.phases.graph_ns,
        "readback_ns": result.phases.readback_ns,
        "validation_ns": result.phases.validation_ns,
    })
}

fn allocation_json(value: Option<support::AllocDelta>) -> Value {
    value.map_or(Value::Null, |delta| {
        json!({
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
        })
    })
}

fn phase_sum(result: &RunResult) -> u64 {
    let phases = result.phases;
    phases
        .capture_ns
        .saturating_add(phases.snapshot_ns)
        .saturating_add(phases.stage_ns)
        .saturating_add(phases.commit_ns)
        .saturating_add(phases.publish_ns)
        .saturating_add(phases.reopen_ns)
        .saturating_add(phases.inverse_reopen_ns)
        .saturating_add(phases.inverse_ns)
        .saturating_add(phases.projection_ns)
        .saturating_add(phases.opaque_ns)
        .saturating_add(phases.graph_ns)
        .saturating_add(phases.readback_ns)
        .saturating_add(phases.validation_ns)
}

fn validate_attribution_delta(
    label: &str,
    value: Option<support::AllocDelta>,
) -> Result<(), String> {
    let Some(delta) = value else {
        return Ok(());
    };
    if !delta.balanced() || delta.invalid || delta.failed != 0 {
        return Err(format!("{label} allocator gate failed"));
    }
    Ok(())
}

fn sample_json(
    lane: &str,
    scale: Scale,
    smoke: &str,
    fixture: &Fixture,
    result: &RunResult,
    elapsed_ns: u64,
    delta: support::AllocDelta,
) -> Result<Value, String> {
    if !result.actual_success || !result.semantic_ok || !result.opaque_ok {
        return Err(format!(
            "lane {lane} scale {} correctness gate failed: success={} semantic={} opaque={}",
            scale.label(),
            result.actual_success,
            result.semantic_ok,
            result.opaque_ok
        ));
    }
    if phase_sum(result) > elapsed_ns {
        return Err(format!(
            "lane {lane} scale {} phases exceed elapsed: {} > {}",
            scale.label(),
            phase_sum(result),
            elapsed_ns
        ));
    }
    if result
        .phases
        .apply_ns
        .saturating_add(result.phases.serialize_ns)
        > result.phases.publish_ns
    {
        return Err(format!(
            "lane {lane} scale {} forward attribution exceeds publish: {} + {} > {}",
            scale.label(),
            result.phases.apply_ns,
            result.phases.serialize_ns,
            result.phases.publish_ns
        ));
    }
    if result
        .phases
        .inverse_apply_ns
        .saturating_add(result.phases.inverse_serialize_ns)
        > result.phases.inverse_ns
    {
        return Err(format!(
            "lane {lane} scale {} inverse attribution exceeds inverse: {} + {} > {}",
            scale.label(),
            result.phases.inverse_apply_ns,
            result.phases.inverse_serialize_ns,
            result.phases.inverse_ns
        ));
    }
    validate_attribution_delta("apply", result.phases.apply_alloc)?;
    validate_attribution_delta("serialize", result.phases.serialize_alloc)?;
    validate_attribution_delta("inverse apply", result.phases.inverse_apply_alloc)?;
    validate_attribution_delta("inverse serialize", result.phases.inverse_serialize_alloc)?;
    if !delta.balanced() || delta.invalid || delta.failed != 0 {
        return Err(format!(
            "lane {lane} scale {} allocator gate failed",
            scale.label()
        ));
    }
    Ok(json!({
        "schema": PROFILE_SCHEMA,
        "source_commit": SOURCE_COMMIT,
        "lane": lane,
        "smoke_lane": smoke,
        "scale": scale.label(),
        "fixture": fixture.name,
        "fixture_native": fixture.native,
        "fixture_signed": fixture.signed,
        "fixture_main_present": fixture.main_present,
        "fixture_glossary_present": fixture.glossary_present,
        "fixture_expected_package_sha256": fixture.expected_package_sha256,
        "input_resource_present": result.input_resource.is_some(),
        "output_resource_present": result.output_resource.is_some(),
        "input_resource": resource_metadata_json(result.input_resource.as_ref()),
        "output_resource": resource_metadata_json(result.output_resource.as_ref()),
        "input_effects_member": member_digest_json(result.input_effects_member.as_ref()),
        "output_effects_member": member_digest_json(result.output_effects_member.as_ref()),
        "source_backed_api": true,
        "process_id": std::process::id(),
        "process_index": std::env::var("DOCX_PROFILE_PROCESS_INDEX")
            .ok()
            .and_then(|value| value.parse::<u64>().ok()),
        "elapsed_ns": elapsed_ns,
        "phases": phases_json(result),
        "apply_allocation": allocation_json(result.phases.apply_alloc),
        "serialize_allocation": allocation_json(result.phases.serialize_alloc),
        "inverse_apply_allocation": allocation_json(result.phases.inverse_apply_alloc),
        "inverse_serialize_allocation": allocation_json(result.phases.inverse_serialize_alloc),
        "actual_success": result.actual_success,
        "semantic_ok": result.semantic_ok,
        "opaque_ok": result.opaque_ok,
        "exact_inverse_ok": result.exact_inverse_ok,
        "output_bytes": result.output_bytes,
        "input_sha256": result.input_sha256,
        "output_sha256": result.output_sha256,
        "input_member_digest": result.input_member_digest,
        "output_member_digest": result.output_member_digest,
        "input_metrics": metrics_json(result.input_metrics),
        "output_metrics": result.output_metrics.map(metrics_json),
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
    }))
}

fn run(lane: &str, scale: Scale, warmup: usize, samples: usize) -> Result<Value, String> {
    let smoke = smoke_lane(lane);
    let fixture = fixture(lane, scale)?;
    support::reset_process_counters();
    for _ in 0..warmup {
        let prepared =
            smoke_adapter::prepare(smoke, &fixture).map_err(|error| error.to_string())?;
        let (result, prepared_after) =
            smoke_adapter::run_prepared(prepared).map_err(|error| error.to_string())?;
        if !result.actual_success || !result.semantic_ok || !result.opaque_ok {
            return Err(format!(
                "warmup correctness failed for {lane}/{}",
                scale.label()
            ));
        }
        drop(prepared_after);
    }
    let mut measured = Vec::with_capacity(samples);
    for _ in 0..samples {
        // Package open/serialization/member hashing and replacement Resource
        // construction belong to preparation, outside this operation window.
        let prepared =
            smoke_adapter::prepare(smoke, &fixture).map_err(|error| error.to_string())?;
        begin_window();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let (mut result, prepared_after) =
            smoke_adapter::run_prepared(prepared).map_err(|error| error.to_string())?;
        let elapsed_ns = started.elapsed().as_nanos().try_into().unwrap_or(u64::MAX);
        let after = AllocSnapshot::now();
        let delta = before.delta(after);
        smoke_adapter::populate_resource_metadata(
            &mut result,
            &fixture,
            smoke_adapter::profile_owner(smoke),
        )
        .map_err(|error| error.to_string())?;
        measured.push(sample_json(
            lane, scale, smoke, &fixture, &result, elapsed_ns, delta,
        )?);
        // The preparation guard owns the package/baseline/member maps and all
        // prepared resources. Explicitly drop it only after both elapsed and
        // allocation snapshots have been taken and the sample is published.
        drop(prepared_after);
    }
    Ok(json!({
        "schema": PROFILE_SCHEMA,
        "source_commit": SOURCE_COMMIT,
        "lane": lane,
        "scale": scale.label(),
        "warmup": warmup,
        "sample_count": samples,
        "samples": measured,
    }))
}

fn main() -> ExitCode {
    if env::args().any(|argument| argument == "--emit-fixture-manifest") {
        let mut output_dir = None;
        let mut arguments = env::args().skip(1);
        while let Some(argument) = arguments.next() {
            if argument == "--output-dir" {
                let Some(value) = arguments.next() else {
                    eprintln!("--output-dir requires a value");
                    return ExitCode::from(2);
                };
                output_dir = Some(std::path::PathBuf::from(value));
            }
        }
        let rows = [
            ("native_capture_bug_main", false, Scale::KiB64),
            ("native_capture_bug_main", false, Scale::MiB1),
            ("native_capture_bug_glossary", true, Scale::KiB64),
            ("native_capture_bug_glossary", true, Scale::MiB1),
        ];
        let generated = rows
            .into_iter()
            .map(|(lane, glossary, scale)| {
                synthetic::manifest_for_lane(lane, glossary, scale, output_dir.as_deref())
            })
            .collect::<Result<Vec<_>, _>>();
        return match generated {
            Ok(rows) => {
                println!(
                    "{}",
                    json!({
                        "schema": "docx-styles-effects-generated-fixtures-v1",
                        "source_commit": SOURCE_COMMIT,
                        "fixtures": rows,
                    })
                );
                ExitCode::SUCCESS
            },
            Err(error) => {
                eprintln!("fixture manifest generation failed: {error}");
                ExitCode::from(1)
            },
        };
    }
    let arguments = match parse_args() {
        Ok(arguments) => arguments,
        Err(error) if error == "help" => return ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("{error}");
            return ExitCode::from(2);
        },
    };
    let (lane, scale, warmup, samples) = arguments;
    match run(&lane, scale, warmup, samples) {
        Ok(receipt) => {
            println!(
                "{}",
                serde_json::to_string_pretty(&receipt).expect("serialize profile receipt")
            );
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("profile scaffold operation failed: {error}");
            ExitCode::from(1)
        },
    }
}

#[cfg(test)]
mod tests {
    use super::support;

    #[test]
    fn allocator_subphase_peaks_reset_without_erasing_outer_peak() {
        // The test runs with the same global allocator as the profile binary.
        // Run it as a single test when replaying so no other test thread can
        // interleave allocator traffic with these process-local counters.
        support::reset_process_counters();
        let outer_before = support::AllocSnapshot::now();

        let retained = vec![0_u8; 4 * 1024];

        support::begin_phase_window();
        let phase_one_before = support::AllocSnapshot::now();
        let mut growing = Vec::with_capacity(8);
        growing.resize(64 * 1024, 0_u8);
        growing.reserve_exact(128 * 1024);
        growing.truncate(32);
        growing.shrink_to_fit();
        drop(growing);
        let phase_one_after = support::AllocSnapshot::now();
        let phase_one = phase_one_before.phase_delta(phase_one_after);

        support::begin_phase_window();
        let phase_two_before = support::AllocSnapshot::now();
        let small = vec![0_u8; 128];
        drop(small);
        let phase_two_after = support::AllocSnapshot::now();
        let phase_two = phase_two_before.phase_delta(phase_two_after);
        let outer_after_phases = support::AllocSnapshot::now();

        drop(retained);
        let outer_after = support::AllocSnapshot::now();
        let outer = outer_before.delta(outer_after);

        assert!(
            phase_one.realloc_calls > 0,
            "growth/shrink did not exercise realloc"
        );
        assert!(phase_one.balanced());
        assert!(phase_two.balanced());
        assert!(outer.balanced());
        assert!(!phase_one.invalid);
        assert!(!phase_two.invalid);
        assert!(!outer.invalid);
        assert_eq!(phase_one.live_before, phase_one.live_after);
        assert_eq!(phase_two.live_before, phase_two.live_after);
        assert_eq!(phase_one.live_after, phase_two.live_before);
        assert_eq!(phase_two.live_after, outer_after_phases.live);
        assert_eq!(outer.live_before, outer.live_after);
        assert!(phase_two.peak_delta > 0);
        assert!(phase_one.peak_delta > phase_two.peak_delta);
        assert!(outer_after_phases.peak >= phase_one_after.peak);
        assert_eq!(outer_after.peak, outer_after_phases.peak);
        assert!(outer.peak_delta > 0);
        assert!(phase_one.calls.saturating_add(phase_two.calls) <= outer.calls);
        assert!(
            phase_one
                .realloc_calls
                .saturating_add(phase_two.realloc_calls)
                <= outer.realloc_calls
        );
        assert!(
            phase_one
                .dealloc_calls
                .saturating_add(phase_two.dealloc_calls)
                <= outer.dealloc_calls
        );
        assert!(phase_one.direct.saturating_add(phase_two.direct) <= outer.direct);
        assert!(phase_one.realloc_old.saturating_add(phase_two.realloc_old) <= outer.realloc_old);
        assert!(phase_one.realloc_new.saturating_add(phase_two.realloc_new) <= outer.realloc_new);
        assert!(phase_one.deallocated.saturating_add(phase_two.deallocated) <= outer.deallocated);
        assert!(phase_one.failed.saturating_add(phase_two.failed) <= outer.failed);
    }
}
