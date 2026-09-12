//! Process-isolated smoke runner for the source-backed XLSB identity profile.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::print_stdout,
    reason = "the opt-in profile emits machine-readable receipts"
)]

mod adapter;
mod matrix;
mod support;

#[global_allocator]
static GLOBAL_ALLOCATOR: support::CountingAllocator = support::CountingAllocator;

use std::env;
use std::process::ExitCode;
use std::time::Instant;

use adapter::RunResult;
use litchi_xlsb::data_model::Limits;
use serde_json::{Value, json};
use support::{Snapshot, reset};

const BASELINE_COMMIT: &str = "1b0d4864804d666aaf4ae5039bd6300996ca32a5";
const NEUTRAL_COMMIT: &str = "4f53be2d31e14b215d6eeaee56d11eefeae94107";

#[derive(Debug)]
struct Arguments {
    lane: Option<String>,
    matrix_correctness: bool,
    case: Option<String>,
    endpoint_layout: Option<String>,
    name_profile: Option<String>,
    warmup: usize,
    samples: usize,
}

fn parse_args() -> Result<Arguments, String> {
    let mut arguments = env::args().skip(1);
    let mut lane = None;
    let mut matrix_correctness = false;
    let mut case = None;
    let mut endpoint_layout = None;
    let mut name_profile = None;
    let mut warmup = 0;
    let mut samples = 1;
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--help" | "-h" => return Err(String::from("help")),
            "--lane" => lane = Some(arguments.next().ok_or("--lane requires a value")?),
            "--matrix-correctness" => matrix_correctness = true,
            "--case" => case = Some(arguments.next().ok_or("--case requires a value")?),
            "--endpoint-layout" => {
                endpoint_layout = Some(
                    arguments
                        .next()
                        .ok_or("--endpoint-layout requires a value")?,
                );
            },
            "--name-profile" => {
                name_profile = Some(arguments.next().ok_or("--name-profile requires a value")?);
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
    if matrix_correctness {
        if lane.is_some() {
            return Err(String::from(
                "--matrix-correctness cannot be combined with --lane",
            ));
        }
        if let Some(layout) = endpoint_layout.as_deref()
            && !matrix::ENDPOINT_LAYOUTS.contains(&layout)
        {
            return Err(format!("unknown endpoint layout: {layout}"));
        }
        if let Some(profile) = name_profile.as_deref()
            && !matrix::NAME_PROFILES.contains(&profile)
        {
            return Err(format!("unknown name profile: {profile}"));
        }
    } else {
        if case.is_some() || endpoint_layout.is_some() || name_profile.is_some() {
            return Err(String::from(
                "--case/--endpoint-layout/--name-profile require --matrix-correctness",
            ));
        }
        let selected_lane = lane
            .as_deref()
            .ok_or("--lane is required unless --matrix-correctness is selected")?;
        if !adapter::is_known_lane(selected_lane) {
            return Err(format!("unknown lane: {selected_lane}"));
        }
    }
    if samples == 0 {
        return Err(String::from("--samples must be nonzero"));
    }
    Ok(Arguments {
        lane,
        matrix_correctness,
        case,
        endpoint_layout,
        name_profile,
        warmup,
        samples,
    })
}

fn usage() {
    eprintln!("usage: xlsb-model-identity-profile --lane NAME [--warmup N] [--samples N]");
    eprintln!(
        "       xlsb-model-identity-profile --matrix-correctness [--case family:tables:relationships] [--endpoint-layout selected_table|distributed] [--name-profile PROFILE]"
    );
    eprintln!("lanes: {}", adapter::LANES.join(", "));
    eprintln!("name profiles: {}", matrix::NAME_PROFILES.join(", "));
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

fn semantic_json(identity: &adapter::SemanticIdentity) -> Value {
    json!({
        "tables": identity.tables.iter().map(|table| json!({
            "table_id": table.table_id,
            "xml_name": table.xml_name,
            "metadata_path": table.metadata_path,
            "dimension_object_id": table.dimension_object_id,
        })).collect::<Vec<_>>(),
        "relationships": identity.relationships.iter().map(|relationship| json!({
            "relationship_id": relationship.relationship_id,
            "metadata_path": relationship.metadata_path,
            "containing_table": relationship.containing_table,
            "primary_table": relationship.primary_table,
            "primary_column": relationship.primary_column,
            "foreign_column": relationship.foreign_column,
            "expected_index_key": relationship.expected_index_key,
        })).collect::<Vec<_>>(),
        "time_groupings": identity.time_groupings.iter().map(|grouping| json!({
            "table_name": grouping.table_name,
            "column_id": grouping.column_id,
            "column_ids": grouping.column_ids,
        })).collect::<Vec<_>>(),
    })
}

fn semantic_check_json(check: &adapter::SemanticCheck) -> Value {
    json!({
        "table_ids_equal": check.table_ids_equal,
        "table_names_equal": check.table_names_equal,
        "relationship_ids_equal": check.relationship_ids_equal,
        "relationship_endpoints_equal": check.relationship_endpoints_equal,
        "relationship_paths_equal": check.relationship_paths_equal,
        "time_grouping_ids_equal": check.time_grouping_ids_equal,
        "all_equal": check.all_equal,
    })
}

fn matrix_member_manifest_json(manifest: &adapter::MatrixMemberManifest) -> Value {
    json!({
        "parts": manifest.parts,
        "relationships": manifest.relationships,
        "content_types": manifest.content_types,
        "inner_members": manifest.inner_members,
    })
}

fn matrix_preservation_json(preservation: &adapter::MatrixPreservation) -> Value {
    json!({
        "source": matrix_member_manifest_json(&preservation.source),
        "candidate": matrix_member_manifest_json(&preservation.candidate),
        "mutable_inner_paths": preservation.mutable_inner_paths,
    })
}

fn scale_matrix_json() -> Value {
    json!(
        matrix::SCALE_MATRIX
            .iter()
            .map(|case| json!({
                "family": case.family,
                "tables": case.tables,
                "relationships": case.relationships,
            }))
            .collect::<Vec<_>>()
    )
}

fn caller_limits_json() -> Value {
    let limits = Limits::DEFAULT;
    json!({
        "raw_payload_bytes": limits.raw.payload(),
        "raw_string_units": limits.raw.string_units(),
        "max_tables": limits.max_tables,
        "max_relationships": limits.max_relationships,
        "max_time_groupings": limits.max_time_groupings,
        "max_time_grouping_columns": limits.max_time_grouping_columns,
        "max_rewrite_bytes": limits.max_rewrite_bytes,
        "max_records": limits.max_records,
        "max_part_bytes": limits.max_part_bytes,
        "max_graph_parts": limits.max_graph_parts,
        "max_graph_relationships": limits.max_graph_relationships,
        "max_metadata_bytes": limits.max_metadata_bytes,
        "max_connection_bytes": limits.max_connection_bytes,
        "max_connections": limits.max_connections,
    })
}

fn olap_proof_limits_json() -> Value {
    let limits = litchi_xldm::OlapProofLimits::default();
    json!({
        "max_items": limits.max_items,
        "max_string_bytes": limits.max_string_bytes,
        "max_source_bytes": limits.max_source_bytes,
        "max_work": limits.max_work,
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
        "staged_bytes": result.staged_bytes,
        "candidate_bytes": result.candidate_bytes,
        "output_bytes": result.output_bytes,
        "semantic_observed": result.semantic_observed.as_ref().map(semantic_json),
        "semantic_check": result.semantic_check.as_ref().map(semantic_check_json),
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
        "recipe": {
            "id": matrix::RECIPE_ID,
            "version": matrix::RECIPE_VERSION,
            "source": matrix::RECIPE_SOURCE,
            "generator": matrix::RECIPE_GENERATOR,
            "selected_scale": fixture.scale,
            "selected_table_id": fixture.selected_table_id,
            "selected_tables": fixture.table_count,
            "selected_relationships": fixture.relationship_count,
            "selected_endpoint_layout": fixture.endpoint_layout,
            "selected_name_profile": fixture.name_profile,
            "scale_matrix": scale_matrix_json(),
            "name_profiles": matrix::NAME_PROFILES,
            "endpoint_layouts": matrix::ENDPOINT_LAYOUTS,
            "olap_proof_limits": olap_proof_limits_json(),
        },
        "scale": fixture.scale,
        "recipe_version": fixture.recipe_version,
        "table_count": fixture.table_count,
        "relationship_count": fixture.relationship_count,
        "caller_limits": caller_limits_json(),
        "olap_proof_limits": olap_proof_limits_json(),
        "input_bytes": fixture.bytes.len(),
        "input_sha256": fixture.input_sha256,
        "input_hash_fnv1a64": fixture.input_fnv1a64,
        "semantic_source": fixture.semantic.as_ref().map(semantic_json),
        "source_backed_api": true,
        "native_acceptance_claim": false,
        "samples": measured,
    }))
}

fn select_matrix_cases(selector: Option<&str>) -> Result<Vec<matrix::ScaleCase>, String> {
    let Some(selector) = selector else {
        return Ok(matrix::SCALE_MATRIX.to_vec());
    };
    let fields = selector.split(':').collect::<Vec<_>>();
    let (family, tables, relationships) = match fields.as_slice() {
        [family, tables, relationships] => (
            Some(*family),
            tables
                .parse::<usize>()
                .map_err(|_| format!("invalid matrix table count in --case: {selector}"))?,
            relationships
                .parse::<usize>()
                .map_err(|_| format!("invalid matrix relationship count in --case: {selector}"))?,
        ),
        [tables, relationships] => (
            None,
            tables
                .parse::<usize>()
                .map_err(|_| format!("invalid matrix table count in --case: {selector}"))?,
            relationships
                .parse::<usize>()
                .map_err(|_| format!("invalid matrix relationship count in --case: {selector}"))?,
        ),
        _ => {
            return Err(format!(
                "--case must be family:tables:relationships or tables:relationships: {selector}"
            ));
        },
    };
    matrix::SCALE_MATRIX
        .iter()
        .copied()
        .find(|case| {
            case.tables == tables
                && case.relationships == relationships
                && family.is_none_or(|value| case.family == value)
        })
        .map(|case| vec![case])
        .ok_or_else(|| format!("matrix case is not in the frozen catalog: {selector}"))
}

fn matrix_result_json(result: &adapter::MatrixResult) -> Value {
    json!({
        "status": result.status,
        "expected_success": result.expected_success,
        "family": result.family,
        "tables": result.tables,
        "relationships": result.relationships,
        "endpoint_layout": result.endpoint_layout,
        "name_profile": result.name_profile,
        "source_bytes": result.source_bytes,
        "staged_bytes": result.staged_bytes,
        "candidate_bytes": result.candidate_bytes,
        "output_bytes": result.output_bytes,
        "source_sha256": result.source_sha256,
        "source_unchanged": result.source_unchanged,
        "no_op_exact": result.no_op_exact,
        "source_proof_status": result.source_proof_status,
        "source_semantic": result.source_semantic.as_ref().map(semantic_json),
        "semantic_observed": result.semantic_observed.as_ref().map(semantic_json),
        "semantic": result.semantic.as_ref().map(semantic_check_json),
        "reopened_semantic_observed": result.reopened_semantic_observed.as_ref().map(semantic_json),
        "reopened_semantic": result.reopened_semantic.as_ref().map(semantic_check_json),
        "preservation": result.preservation.as_ref().map(matrix_preservation_json),
        "inverse_exact": result.inverse_exact,
        "error": result.error.as_ref().map(|error| json!({
            "class": error.class,
            "message": error.message,
            "typed_match": error.typed_match,
        })),
    })
}

fn run_matrix_correctness(
    case_selector: Option<&str>,
    endpoint_selector: Option<&str>,
    profile_selector: Option<&str>,
) -> Result<Value, String> {
    let cases = select_matrix_cases(case_selector)?;
    let layouts = endpoint_selector
        .map(|layout| vec![layout])
        .unwrap_or_else(|| matrix::ENDPOINT_LAYOUTS.to_vec());
    let profiles = profile_selector
        .map(|profile| vec![profile])
        .unwrap_or_else(|| matrix::NAME_PROFILES.to_vec());
    let mut results = Vec::with_capacity(cases.len() * layouts.len() * profiles.len());
    for case in &cases {
        for layout in &layouts {
            for profile in &profiles {
                let result = adapter::run_matrix_case(*case, layout, profile).map_err(|error| {
                    format!(
                        "matrix correctness failed at {}:{}:{}: {error}",
                        case.family, layout, profile
                    )
                })?;
                results.push(matrix_result_json(&result));
            }
        }
    }
    let passed = results
        .iter()
        .filter(|result| result.get("status").and_then(Value::as_str) == Some("passed"))
        .count();
    let expected_refusals = results
        .iter()
        .filter(|result| result.get("status").and_then(Value::as_str) == Some("expected_refusal"))
        .count();
    Ok(json!({
        "schema": "xlsb-model-identity-profile-v1-correctness",
        "baseline_commit": BASELINE_COMMIT,
        "neutral_baseline_commit": NEUTRAL_COMMIT,
        "fixture_kind": "synthetic_complete_xldm140",
        "correctness_only": true,
        "timings_collected": false,
        "source_backed_api": true,
        "native_acceptance_claim": false,
        "caller_limits": caller_limits_json(),
        "olap_proof_limits": olap_proof_limits_json(),
        "recipe": {
            "id": matrix::RECIPE_ID,
            "version": matrix::RECIPE_VERSION,
            "source": matrix::RECIPE_SOURCE,
            "generator": matrix::RECIPE_GENERATOR,
            "limits_mode": "Limits::DEFAULT",
            "caller_limits": caller_limits_json(),
            "olap_proof_limits": olap_proof_limits_json(),
            "scale_matrix": scale_matrix_json(),
            "name_profiles": matrix::NAME_PROFILES,
            "endpoint_layouts": matrix::ENDPOINT_LAYOUTS,
        },
        "selected": {
            "case": case_selector,
            "endpoint_layout": endpoint_selector,
            "name_profile": profile_selector,
        },
        "coverage": {
            "cases": cases.len(),
            "endpoint_layouts": layouts.len(),
            "name_profiles": profiles.len(),
            "runs": results.len(),
            "passed": passed,
            "expected_refusals": expected_refusals,
        },
        "results": results,
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
    if !neutral
        && expected_success
        && result
            .semantic_check
            .as_ref()
            .is_none_or(|check| !check.all_equal)
    {
        return Err(format!(
            "semantic identity/relationship/time-grouping gate failed for lane {lane}"
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
    let arguments = match parse_args() {
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
    let receipt = if arguments.matrix_correctness {
        run_matrix_correctness(
            arguments.case.as_deref(),
            arguments.endpoint_layout.as_deref(),
            arguments.name_profile.as_deref(),
        )
    } else {
        run(
            arguments.lane.as_deref().expect("lane validated"),
            arguments.warmup,
            arguments.samples,
        )
    };
    match receipt {
        Ok(receipt) => {
            println!(
                "{}",
                serde_json::to_string_pretty(&receipt).expect("serialize profile receipt")
            );
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("smoke correctness failure: {error}");
            ExitCode::from(1)
        },
    }
}
