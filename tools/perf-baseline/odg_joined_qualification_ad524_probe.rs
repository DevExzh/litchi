//! Historical ODG adapter for the paired joined-publication qualification.
//!
//! This file is built only in the pinned `ad52437f5e5e2c71ea476eedd5ddcf61aee2021d`
//! checkout. It deliberately calls the old `Package::apply_joined` lifecycle
//! and does not claim the candidate commit-triad or replay capabilities.

#![cfg_attr(not(odg_joined_allocator), forbid(unsafe_code))]
#![allow(unexpected_cfgs)]

#[cfg(odg_joined_allocator)]
#[path = "odg_joined_qualification_allocator.rs"]
mod allocation;
#[path = "odg_joined_qualification_common.rs"]
mod common;

#[cfg(odg_joined_allocator)]
use common::allocation_once;
use common::{
    Args, Mode, PreparedRoute, Result, digest, fixture_package, foreign_package, prepare_all,
    publish, rebuild_plan, run_batch, run_once, summary, write_json,
};
use litchi_odg::{Package, edit};
use serde_json::{Value, json};

const SCHEMA: &str = "litchi.odg.joined-native-qualification.legacy-probe.v2";
const ENDPOINT: &str = "baseline-apply-joined";

fn apply_plan(source: &Package, plan: edit::JoinedEdits) -> Result<Package> {
    Ok(source
        .apply_joined(plan)
        .map_err(|error| format!("baseline Package::apply_joined failed: {error:?}"))?)
}

fn expected_route_semantics(route: &str, source: &Package, output: &Package) -> bool {
    match route {
        "package_replace_one" => {
            output.pages().len() == source.pages().len() + 1
                && output.pages().last().and_then(|page| page.name()) == Some("Replacement Page")
        },
        "semantic_two_disjoint" => {
            output
                .pages()
                .first()
                .and_then(|page| page.shapes().first())
                .map(|shape| shape.text())
                == Some("Joined page one")
                && output
                    .pages()
                    .get(1)
                    .and_then(|page| page.shapes().get(1))
                    .and_then(|shape| shape.name())
                    == Some("Joined page two")
        },
        "semantic_four_disjoint" => {
            output
                .pages()
                .first()
                .and_then(|page| page.shapes().first())
                .map(|shape| shape.text())
                == Some("Joined page one")
                && output
                    .pages()
                    .first()
                    .and_then(|page| page.shapes().get(1))
                    .and_then(|shape| shape.name())
                    == Some("Joined page one name")
                && output
                    .pages()
                    .get(1)
                    .and_then(|page| page.shapes().first())
                    .map(|shape| shape.text())
                    == Some("Joined page two")
                && output
                    .pages()
                    .get(1)
                    .and_then(|page| page.shapes().get(1))
                    .and_then(|shape| shape.name())
                    == Some("Joined page two name")
        },
        "empty_join" => output.as_bytes() == source.as_bytes(),
        _ => false,
    }
}

fn refusal_proofs(source: &Package, fixture: &[u8], routes: &[PreparedRoute]) -> Result<Value> {
    let foreign = foreign_package(fixture)?;
    let semantic = routes
        .iter()
        .find(|route| route.route == "semantic_two_disjoint")
        .ok_or("semantic route missing")?;
    let stale_error = foreign
        .apply_joined(rebuild_plan(source, semantic)?)
        .is_err();
    let first = &semantic.sub_edits[0];
    let payload = edit::DurablePatch::from_deterministic_json(&first.wire)?;
    let mut conflicts = source.joined_edits(common::LIMITS);
    let lineage = conflicts.lineage().clone();
    conflicts
        .join(edit::SubEdit::new(
            lineage.clone(),
            common::LIMITS,
            "conflict-left",
            Vec::<String>::new(),
            first.writes.clone(),
            payload.clone(),
        )?)
        .map_err(|error| format!("conflict setup failed: {error:?}"))?;
    let conflict_error = conflicts
        .join(edit::SubEdit::new(
            lineage,
            common::LIMITS,
            "conflict-right",
            Vec::<String>::new(),
            first.writes.clone(),
            payload,
        )?)
        .is_err();
    let changed_recipient_refused = foreign
        .apply_joined(rebuild_plan(source, semantic)?)
        .is_err();
    Ok(json!({
        "stale_lineage_refused": stale_error,
        "conflicting_effect_refused": conflict_error,
        "changed_recipient_refused": changed_recipient_refused,
        "changed_recipient_scope": "candidate-only",
    }))
}

fn verify(args: &Args) -> Result<Value> {
    let (fixture, source) = fixture_package(&args.fixture)?;
    let fixture_identity = digest(&fixture);
    let source_identity = digest(source.as_bytes());
    let repeated_source = common::package_from_content(&fixture)?;
    let source_fixture_deterministic = repeated_source.as_bytes() == source.as_bytes();
    let routes = prepare_all(&source)?;
    let refusals = refusal_proofs(&source, &fixture, &routes)?;
    let mut route_reports = serde_json::Map::new();
    for route in &routes {
        let before = source_identity.clone();
        let output = apply_plan(&source, rebuild_plan(&source, route)?)?;
        let repeated = apply_plan(&source, rebuild_plan(&source, route)?)?;
        let (published, readback) = publish(&output)?;
        let changed = route.route != "empty_join";
        route_reports.insert(route.route.to_owned(), json!({
            "output": digest(output.as_bytes()), "published": digest(&published),
            "semantic": summary(&output),
            "checks": {
                "source_unchanged": digest(source.as_bytes()) == before,
                "publication_readback_exact": readback.as_bytes() == published.as_slice(),
                "route_semantics": expected_route_semantics(route.route, &source, &output),
                "changed_output": output.as_bytes() != source.as_bytes() || !changed,
                "exact_noop": route.route != "empty_join" || output.as_bytes() == source.as_bytes(),
                "deterministic_repeat": repeated.as_bytes() == output.as_bytes(),
            },
            "baseline": {"apply_joined": true, "commit_triads": false, "patch_replay": false},
        }));
    }
    Ok(json!({
        "schema": SCHEMA, "endpoint": ENDPOINT, "mode": Mode::Verify.as_str(),
        "fixture": fixture_identity, "source_package": source_identity,
        "routes": Value::Object(route_reports), "refusals": refusals,
        "capabilities": {
            "apply_joined": true, "edit_join_commit": false, "commit_triads": false,
            "patch_replay": false, "inverse_replay": false,
            "content_matching_replay": false, "changed_recipient_refusal": false,
            "changed_recipient_refusal_scope": "candidate-only",
        },
        "checks": {
            "source_fixture_deterministic": source_fixture_deterministic, "all_routes_checked": true,
            "all_baseline_apply_proofs": true,
            "all_refusal_proofs": refusals["stale_lineage_refused"].as_bool() == Some(true)
                && refusals["conflicting_effect_refused"].as_bool() == Some(true)
                && refusals["changed_recipient_scope"].as_str() == Some("candidate-only"),
        },
    }))
}

fn timed(args: &Args, mode: Mode) -> Result<Value> {
    let (_fixture, source) = fixture_package(&args.fixture)?;
    let route_name = args.route.as_deref().ok_or("route is required")?;
    let route = common::prepare_route(&source, route_name)?;
    let scope = args.scope.ok_or("scope is required")?;
    for _ in 0..args.warmups {
        let _ = run_once(&source, &route, scope, |plan| apply_plan(&source, plan))?;
    }
    if matches!(mode, Mode::Calibrate) {
        let mut attempts = Vec::new();
        let mut current = 1_usize;
        loop {
            let (elapsed_ns, proof) = run_batch(&source, &route, scope, current, |plan| {
                apply_plan(&source, plan)
            })?;
            attempts.push(json!({"iterations": current, "elapsed_ns": elapsed_ns, "proof": proof}));
            if elapsed_ns >= args.target_ns {
                return Ok(json!({
                    "schema": SCHEMA, "endpoint": ENDPOINT, "mode": Mode::Calibrate.as_str(),
                    "route": route_name, "scope": scope.as_str(), "target_ns": args.target_ns,
                    "iterations": current, "elapsed_ns": elapsed_ns, "attempts": attempts,
                    "iteration_policy": "adaptive-doubling-v1", "initial_iterations": 1,
                    "max_iterations": args.max_iterations,
                }));
            }
            if current >= args.max_iterations {
                return Err("calibration reached max-iterations before its target".into());
            }
            current = current.saturating_mul(2).min(args.max_iterations);
        }
    }
    let mut samples = Vec::new();
    let mut proof = None;
    for _ in 0..args.samples {
        let (elapsed_ns, current_proof) =
            run_batch(&source, &route, scope, args.iterations, |plan| {
                apply_plan(&source, plan)
            })?;
        if let Some(previous) = &proof {
            if previous != &current_proof {
                return Err("timed output proof differs between samples".into());
            }
        } else {
            proof = Some(current_proof);
        }
        samples.push(elapsed_ns);
    }
    Ok(json!({
        "schema": SCHEMA, "endpoint": ENDPOINT, "mode": mode.as_str(),
        "route": route_name, "scope": scope.as_str(), "iterations": args.iterations,
        "warmups": args.warmups, "samples": samples, "proof": proof.ok_or("no timing samples")?,
    }))
}

#[cfg(odg_joined_allocator)]
fn allocation(args: &Args) -> Result<Value> {
    allocation::enable();
    let (_fixture, source) = fixture_package(&args.fixture)?;
    let route_name = args.route.as_deref().ok_or("route is required")?;
    let route = common::prepare_route(&source, route_name)?;
    let scope = args.scope.ok_or("scope is required")?;
    let mut samples = Vec::new();
    for _ in 0..args.warmups {
        let _ = allocation_once(&source, &route, scope, |plan| apply_plan(&source, plan))?;
    }
    for _ in 0..args.samples {
        let (metrics, proof) =
            allocation_once(&source, &route, scope, |plan| apply_plan(&source, plan))?;
        samples.push(json!({"metrics": metrics, "proof": proof}));
    }
    Ok(json!({
        "schema": SCHEMA, "endpoint": ENDPOINT, "mode": Mode::Allocation.as_str(),
        "route": route_name, "scope": scope.as_str(), "iterations": 1,
        "warmups": args.warmups, "samples": samples,
    }))
}

fn run() -> Result<()> {
    let args = Args::parse()?;
    let report = match args.mode {
        Mode::Verify => verify(&args)?,
        Mode::Calibrate | Mode::Timed => timed(&args, args.mode)?,
        #[cfg(odg_joined_allocator)]
        Mode::Allocation => allocation(&args)?,
        #[cfg(not(odg_joined_allocator))]
        Mode::Allocation => return Err("allocation mode requires the allocator probe build".into()),
    };
    write_json(&report)
}

fn main() {
    if let Err(error) = run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}
