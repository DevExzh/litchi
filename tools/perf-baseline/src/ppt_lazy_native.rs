//! Native probe for the PPT lazy-entry qualification lane (record 0362ugj).
//!
//! The package is built before the operation boundary.  The four routes then
//! isolate construction/drop and exact no-op publication for the primary and
//! metadata edit families.  The verification mode proves the route semantics;
//! timed modes only repeat those already-proved operations.

use std::error::Error;
use std::hint::black_box;
use std::time::Instant;

use litchi_ppt::Package;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

use crate::allocation_metrics;

const SCHEMA: &str = "litchi.ppt.lazy-native.probe.v1";
const ROUTES: [&str; 4] = [
    "primary_entry_drop",
    "primary_noop_commit",
    "metadata_entry_drop",
    "metadata_noop_commit",
];

#[derive(Clone, Copy)]
enum Route {
    PrimaryEntryDrop,
    PrimaryNoopCommit,
    MetadataEntryDrop,
    MetadataNoopCommit,
}

impl Route {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "primary_entry_drop" => Ok(Self::PrimaryEntryDrop),
            "primary_noop_commit" => Ok(Self::PrimaryNoopCommit),
            "metadata_entry_drop" => Ok(Self::MetadataEntryDrop),
            "metadata_noop_commit" => Ok(Self::MetadataNoopCommit),
            _ => Err(format!("unknown PPT lazy route {value:?}").into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::PrimaryEntryDrop => ROUTES[0],
            Self::PrimaryNoopCommit => ROUTES[1],
            Self::MetadataEntryDrop => ROUTES[2],
            Self::MetadataNoopCommit => ROUTES[3],
        }
    }

    const fn is_commit(self) -> bool {
        matches!(self, Self::PrimaryNoopCommit | Self::MetadataNoopCommit)
    }
}

struct OperationResult {
    changed: bool,
    patch_noop: Option<bool>,
    published: Option<Package>,
}

fn run_once(package: &Package, route: Route) -> Result<OperationResult, Box<dyn Error>> {
    match route {
        Route::PrimaryEntryDrop => {
            let edit = package.edit()?;
            drop(edit);
            Ok(OperationResult {
                changed: false,
                patch_noop: None,
                published: None,
            })
        },
        Route::PrimaryNoopCommit => {
            let commit = package.edit()?.commit()?;
            let changed = commit.diagnostics().changed();
            let patch_noop = commit.patch().is_noop();
            Ok(OperationResult {
                changed,
                patch_noop: Some(patch_noop),
                published: Some(commit.into_package()),
            })
        },
        Route::MetadataEntryDrop => {
            let edit = package.edit_metadata()?;
            drop(edit);
            Ok(OperationResult {
                changed: false,
                patch_noop: None,
                published: None,
            })
        },
        Route::MetadataNoopCommit => {
            let commit = package.edit_metadata()?.commit()?;
            let changed = commit.diagnostics().changed();
            let patch_noop = commit.patch().is_noop();
            Ok(OperationResult {
                changed,
                patch_noop: Some(patch_noop),
                published: Some(commit.into_package()),
            })
        },
    }
}

fn digest(bytes: &[u8]) -> Value {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    json!({"bytes": bytes.len(), "sha256": format!("{:x}", hasher.finalize())})
}

fn assert_noop(
    source: &[u8],
    route: Route,
    result: OperationResult,
) -> Result<Value, Box<dyn Error>> {
    if result.changed || result.patch_noop != Some(true) {
        return Err(format!("{} was not an exact no-op", route.name()).into());
    }
    let published = result
        .published
        .ok_or_else(|| format!("{} did not publish a package", route.name()))?;
    let bytes = published.to_bytes()?;
    if bytes != source {
        return Err(format!("{} changed the source bytes", route.name()).into());
    }
    Ok(json!({
        "output": digest(&bytes),
        "source_exact": true,
        "changed": false,
        "patch_noop": true,
        "entry": "commit",
    }))
}

fn verify_route(package: &Package, source: &[u8], route: Route) -> Result<Value, Box<dyn Error>> {
    let result = run_once(package, route)?;
    if route.is_commit() {
        assert_noop(source, route, result)
    } else {
        if result.changed || result.patch_noop.is_some() || result.published.is_some() {
            return Err(format!("{} unexpectedly published", route.name()).into());
        }
        Ok(json!({
            "output": digest(source),
            "source_exact": true,
            "changed": false,
            "patch_noop": Value::Null,
            "entry": "drop",
        }))
    }
}

fn allocation_delta(
    before: allocation_metrics::Snapshot,
    after: allocation_metrics::Snapshot,
) -> Result<Value, Box<dyn Error>> {
    if after.overflowed || after.observer_invalid || before.overflowed || before.observer_invalid {
        return Err("allocator counters became invalid during the route".into());
    }
    let difference = |after: u64, before: u64, name: &str| {
        after
            .checked_sub(before)
            .ok_or_else(|| format!("allocator {name} counter regressed"))
    };
    let allocated_bytes = difference(
        after.allocated_bytes,
        before.allocated_bytes,
        "allocated_bytes",
    )?;
    let deallocated_bytes = difference(
        after.deallocated_bytes,
        before.deallocated_bytes,
        "deallocated_bytes",
    )?;
    let live_delta = i128::from(after.live_bytes) - i128::from(before.live_bytes);
    if live_delta != i128::from(allocated_bytes) - i128::from(deallocated_bytes)
        || before.peak_live_bytes < before.live_bytes
        || after.peak_live_bytes < after.live_bytes
        || after.peak_live_bytes < before.peak_live_bytes
    {
        return Err("allocator live-byte conservation or peak bounds failed".into());
    }
    let allocation_calls = difference(
        after.allocation_calls,
        before.allocation_calls,
        "allocation_calls",
    )?;
    let reallocation_calls = difference(
        after.reallocation_calls,
        before.reallocation_calls,
        "reallocation_calls",
    )?;
    if reallocation_calls > allocation_calls {
        return Err("allocator reallocation calls exceed allocation calls".into());
    }
    Ok(json!({
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "allocation_calls": allocation_calls,
        "deallocation_calls": difference(after.deallocation_calls, before.deallocation_calls, "deallocation_calls")?,
        "reallocation_calls": reallocation_calls,
        "failed_allocation_calls": difference(after.failed_allocation_calls, before.failed_allocation_calls, "failed_allocation_calls")?,
        "allocated_bytes": allocated_bytes,
        "deallocated_bytes": deallocated_bytes,
        "live_bytes_before": before.live_bytes,
        "live_bytes_after": after.live_bytes,
        "peak_live_bytes_before": before.peak_live_bytes,
        "peak_live_bytes_after": after.peak_live_bytes,
    }))
}

fn unavailable_allocation() -> Value {
    json!({
        "status": "unavailable",
        "scope": "operation_global_system_allocator",
    })
}

fn run_batch(package: &Package, route: Route, iterations: usize) -> Result<(), Box<dyn Error>> {
    for _ in 0..iterations {
        let result = run_once(package, route)?;
        if result.changed || result.patch_noop == Some(false) {
            return Err(format!("{} ceased to be an exact no-op", route.name()).into());
        }
        black_box(result);
    }
    Ok(())
}

fn parse_usize(value: &str, flag: &str) -> Result<usize, Box<dyn Error>> {
    value
        .parse()
        .map_err(|error| format!("invalid {flag} value {value:?}: {error}").into())
}

fn parse_u64(value: &str, flag: &str) -> Result<u64, Box<dyn Error>> {
    value
        .parse()
        .map_err(|error| format!("invalid {flag} value {value:?}: {error}").into())
}

struct Arguments {
    mode: String,
    route: Option<Route>,
    target_ns: u64,
    max_iterations: usize,
    iterations: usize,
    samples: usize,
    warmups: usize,
}

fn arguments<I>(values: I) -> Result<Arguments, Box<dyn Error>>
where
    I: IntoIterator<Item = String>,
{
    let mut mode = None;
    let mut route = None;
    let mut target_ns = 0;
    let mut max_iterations = 1_000_000;
    let mut iterations = 1;
    let mut samples = 1;
    let mut warmups = 0;
    let mut values = values.into_iter();
    while let Some(flag) = values.next() {
        let value = values
            .next()
            .ok_or_else(|| format!("missing value for {flag}"))?;
        match flag.as_str() {
            "--mode" => mode = Some(value),
            "--route" => route = Some(Route::parse(&value)?),
            "--target-ns" => target_ns = parse_u64(&value, "--target-ns")?,
            "--max-iterations" => max_iterations = parse_usize(&value, "--max-iterations")?,
            "--iterations" => iterations = parse_usize(&value, "--iterations")?,
            "--samples" => samples = parse_usize(&value, "--samples")?,
            "--warmups" => warmups = parse_usize(&value, "--warmups")?,
            _ => return Err(format!("unknown probe option {flag:?}").into()),
        }
    }
    let mode = mode.ok_or("--mode is required")?;
    if max_iterations == 0 || iterations == 0 || samples == 0 {
        return Err("iteration and sample limits must be positive".into());
    }
    if matches!(mode.as_str(), "calibrate" | "timed") && route.is_none() {
        return Err(format!("{mode} requires --route").into());
    }
    Ok(Arguments {
        mode,
        route,
        target_ns,
        max_iterations,
        iterations,
        samples,
        warmups,
    })
}

fn timed(
    package: &Package,
    args: &Arguments,
    route: Route,
    allocator: bool,
) -> Result<Value, Box<dyn Error>> {
    run_batch(package, route, args.warmups)?;
    let mut samples = Vec::with_capacity(args.samples);
    for _ in 0..args.samples {
        let before = allocator.then(allocation_metrics::snapshot);
        let started = Instant::now();
        run_batch(package, route, args.iterations)?;
        let elapsed_ns = started.elapsed().as_nanos();
        let after = allocator.then(allocation_metrics::snapshot);
        let allocation = match (before, after) {
            (Some(before), Some(after)) => allocation_delta(before, after)?,
            _ => unavailable_allocation(),
        };
        samples.push(json!({
            "elapsed_ns": elapsed_ns,
            "iterations": args.iterations,
            "allocation": allocation,
        }));
    }
    Ok(json!({
        "schema": SCHEMA,
        "mode": "timed",
        "route": route.name(),
        "samples": samples,
    }))
}

pub fn run_from_args<I>(values: I, allocator: bool) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = String>,
{
    if allocator {
        allocation_metrics::enable();
    }
    let args = arguments(values)?;
    let package = Package::new()?;
    let source = package.to_bytes()?;
    let source_digest = digest(&source);
    let output = match args.mode.as_str() {
        "verify" => {
            let routes = ROUTES
                .iter()
                .map(|name| {
                    let route = Route::parse(name)?;
                    Ok::<_, Box<dyn Error>>((name, verify_route(&package, &source, route)?))
                })
                .collect::<Result<Vec<_>, _>>()?;
            json!({
                "schema": SCHEMA,
                "mode": "verify",
                "fixture": {"generator": "litchi-ppt-empty-package-v1", "output": source_digest},
                "routes": routes
                    .into_iter()
                    .map(|(name, proof)| ((*name).to_owned(), proof))
                    .collect::<serde_json::Map<_, _>>(),
                "allocation": if allocator { json!({"status":"measured", "scope":"operation_global_system_allocator"}) } else { unavailable_allocation() },
            })
        },
        "calibrate" => {
            let route = args.route.ok_or("calibrate requires --route")?;
            let mut iterations = 1usize;
            let started = Instant::now();
            loop {
                run_batch(&package, route, iterations)?;
                if started.elapsed().as_nanos() >= u128::from(args.target_ns.max(1)) {
                    break;
                }
                iterations = iterations
                    .checked_mul(2)
                    .ok_or("calibration iteration overflow")?;
                if iterations > args.max_iterations {
                    return Err("calibration exceeded --max-iterations".into());
                }
            }
            json!({
                "schema": SCHEMA,
                "mode": "calibrate",
                "route": route.name(),
                "target_ns": args.target_ns,
                "iterations": iterations,
                "elapsed_ns": started.elapsed().as_nanos(),
                "output": source_digest,
            })
        },
        "timed" => {
            let route = args.route.ok_or("timed requires --route")?;
            let report = timed(&package, &args, route, allocator)?;
            report
        },
        _ => return Err(format!("unknown probe mode {:?}", args.mode).into()),
    };
    println!("{}", serde_json::to_string(&output)?);
    Ok(())
}
