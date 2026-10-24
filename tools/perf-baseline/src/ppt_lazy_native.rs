//! Native probe for the PPT lazy-entry qualification lane (record 0362ugj).
//!
//! The package is opened before the operation boundary.  Every affected PPT
//! lazy edit family has an entry/drop and exact no-op publication route.  The
//! verification mode proves the route semantics; timed modes only repeat those
//! already-proved operations.

use std::error::Error;
use std::hint::black_box;
use std::time::Instant;

use litchi_core::Position;
use litchi_ppt::Package;
use litchi_ppt::master::Kind as MasterKind;
use litchi_ppt::text_edit::Target;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

use crate::allocation_metrics;

const SCHEMA: &str = "litchi.ppt.lazy-native.probe.v1";
const DEFAULT_FIXTURE_KEY: &str = "default";
const OBJECTS_FIXTURE_KEY: &str = "objects";
const EXTERNAL_MEDIA_FIXTURE_KEY: &str = "external_media";
const PICTURE_BULLETS_FIXTURE_KEY: &str = "picture_bullets";
const FIXTURE_KEYS: [&str; 4] = [
    DEFAULT_FIXTURE_KEY,
    OBJECTS_FIXTURE_KEY,
    EXTERNAL_MEDIA_FIXTURE_KEY,
    PICTURE_BULLETS_FIXTURE_KEY,
];
const DEFAULT_FIXTURE: &[u8] =
    include_bytes!("../../../test-data/unified-ops/ppt-lazy-native-fixture.ppt");
const OBJECTS_FIXTURE: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/ppt_with_embeded.ppt");
const EXTERNAL_MEDIA_FIXTURE: &[u8] =
    include_bytes!("../../../test-data/unified-ops/0294/video.ppt");
const PICTURE_BULLETS_FIXTURE: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/bug61881.ppt");
const ROUTES: [&str; 28] = [
    "primary_entry_drop",
    "primary_noop_commit",
    "metadata_entry_drop",
    "metadata_noop_commit",
    "animations_entry_drop",
    "animations_noop_commit",
    "broadcast_entry_drop",
    "broadcast_noop_commit",
    "client_data_entry_drop",
    "client_data_noop_commit",
    "diagrams_entry_drop",
    "diagrams_noop_commit",
    "document_comparison_entry_drop",
    "document_comparison_noop_commit",
    "objects_entry_drop",
    "objects_noop_commit",
    "external_media_entry_drop",
    "external_media_noop_commit",
    "fonts_entry_drop",
    "fonts_noop_commit",
    "master_metadata_entry_drop",
    "master_metadata_noop_commit",
    "master_layout_entry_drop",
    "master_layout_noop_commit",
    "picture_bullets_entry_drop",
    "picture_bullets_noop_commit",
    "slide_sync_entry_drop",
    "slide_sync_noop_commit",
];

#[derive(Clone, Copy)]
enum Family {
    Primary,
    Metadata,
    Animations,
    Broadcast,
    ClientData,
    Diagrams,
    DocumentComparison,
    Objects,
    ExternalMedia,
    Fonts,
    MasterMetadata,
    MasterLayout,
    PictureBullets,
    SlideSync,
}

impl Family {
    const ALL: [Self; 14] = [
        Self::Primary,
        Self::Metadata,
        Self::Animations,
        Self::Broadcast,
        Self::ClientData,
        Self::Diagrams,
        Self::DocumentComparison,
        Self::Objects,
        Self::ExternalMedia,
        Self::Fonts,
        Self::MasterMetadata,
        Self::MasterLayout,
        Self::PictureBullets,
        Self::SlideSync,
    ];

    const fn entry_name(self) -> &'static str {
        match self {
            Self::Primary => ROUTES[0],
            Self::Metadata => ROUTES[2],
            Self::Animations => ROUTES[4],
            Self::Broadcast => ROUTES[6],
            Self::ClientData => ROUTES[8],
            Self::Diagrams => ROUTES[10],
            Self::DocumentComparison => ROUTES[12],
            Self::Objects => ROUTES[14],
            Self::ExternalMedia => ROUTES[16],
            Self::Fonts => ROUTES[18],
            Self::MasterMetadata => ROUTES[20],
            Self::MasterLayout => ROUTES[22],
            Self::PictureBullets => ROUTES[24],
            Self::SlideSync => ROUTES[26],
        }
    }

    const fn commit_name(self) -> &'static str {
        match self {
            Self::Primary => ROUTES[1],
            Self::Metadata => ROUTES[3],
            Self::Animations => ROUTES[5],
            Self::Broadcast => ROUTES[7],
            Self::ClientData => ROUTES[9],
            Self::Diagrams => ROUTES[11],
            Self::DocumentComparison => ROUTES[13],
            Self::Objects => ROUTES[15],
            Self::ExternalMedia => ROUTES[17],
            Self::Fonts => ROUTES[19],
            Self::MasterMetadata => ROUTES[21],
            Self::MasterLayout => ROUTES[23],
            Self::PictureBullets => ROUTES[25],
            Self::SlideSync => ROUTES[27],
        }
    }

    const fn fixture_key(self) -> &'static str {
        match self {
            Self::Objects => OBJECTS_FIXTURE_KEY,
            Self::ExternalMedia => EXTERNAL_MEDIA_FIXTURE_KEY,
            Self::PictureBullets => PICTURE_BULLETS_FIXTURE_KEY,
            _ => DEFAULT_FIXTURE_KEY,
        }
    }
}

#[derive(Clone, Copy)]
enum Route {
    EntryDrop(Family),
    NoopCommit(Family),
}

impl Route {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        for family in Family::ALL {
            if value == family.entry_name() {
                return Ok(Self::EntryDrop(family));
            }
            if value == family.commit_name() {
                return Ok(Self::NoopCommit(family));
            }
        }
        Err(format!("unknown PPT lazy route {value:?}").into())
    }

    const fn name(self) -> &'static str {
        match self {
            Self::EntryDrop(family) => family.entry_name(),
            Self::NoopCommit(family) => family.commit_name(),
        }
    }

    const fn is_commit(self) -> bool {
        matches!(self, Self::NoopCommit(_))
    }

    const fn fixture_key(self) -> &'static str {
        match self {
            Self::EntryDrop(family) | Self::NoopCommit(family) => family.fixture_key(),
        }
    }
}

fn fixture_bytes(key: &str) -> Result<&'static [u8], Box<dyn Error>> {
    match key {
        DEFAULT_FIXTURE_KEY => Ok(DEFAULT_FIXTURE),
        OBJECTS_FIXTURE_KEY => Ok(OBJECTS_FIXTURE),
        EXTERNAL_MEDIA_FIXTURE_KEY => Ok(EXTERNAL_MEDIA_FIXTURE),
        PICTURE_BULLETS_FIXTURE_KEY => Ok(PICTURE_BULLETS_FIXTURE),
        _ => Err(format!("unknown PPT lazy fixture {key:?}").into()),
    }
}

fn fixture_source(key: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    Ok(Package::from_bytes(fixture_bytes(key)?)?.to_bytes()?)
}

struct OperationResult {
    changed: bool,
    patch_noop: Option<bool>,
    published: Option<Package>,
    operation_count: Option<usize>,
    patch_before: Option<Vec<u8>>,
    patch_after: Option<Vec<u8>>,
}

fn dropped() -> OperationResult {
    OperationResult {
        changed: false,
        patch_noop: None,
        published: None,
        operation_count: None,
        patch_before: None,
        patch_after: None,
    }
}

macro_rules! drop_edit {
    ($expression:expr) => {{
        let edit = $expression?;
        drop(edit);
        Ok::<_, Box<dyn Error>>(dropped())
    }};
}

macro_rules! commit_edit {
    ($expression:expr) => {{
        let commit = $expression?;
        let changed = commit.diagnostics().changed();
        let patch_noop = commit.patch().is_noop();
        Ok::<_, Box<dyn Error>>(OperationResult {
            changed,
            patch_noop: Some(patch_noop),
            published: Some(commit.into_package()),
            operation_count: None,
            patch_before: None,
            patch_after: None,
        })
    }};
}

fn run_once(package: &Package, route: Route) -> Result<OperationResult, Box<dyn Error>> {
    let position = Position::new(0);
    let target = Target::new(position, position);
    let master_kind = MasterKind::Main;
    match route {
        Route::EntryDrop(Family::Primary) => drop_edit!(package.edit()),
        Route::NoopCommit(Family::Primary) => {
            let commit = package.edit()?.commit()?;
            Ok(OperationResult {
                changed: commit.diagnostics().changed(),
                patch_noop: Some(commit.patch().is_noop()),
                operation_count: Some(commit.diagnostics().operation_count()),
                patch_before: Some(commit.patch().before().to_vec()),
                patch_after: Some(commit.patch().after().to_vec()),
                published: Some(commit.into_package()),
            })
        },
        Route::EntryDrop(Family::Metadata) => drop_edit!(package.edit_metadata()),
        Route::NoopCommit(Family::Metadata) => {
            let commit = package.edit_metadata()?.commit()?;
            Ok(OperationResult {
                changed: commit.diagnostics().changed(),
                patch_noop: Some(commit.patch().is_noop()),
                operation_count: None,
                patch_before: Some(commit.patch().before_bytes().to_vec()),
                patch_after: Some(commit.patch().after_bytes().to_vec()),
                published: Some(commit.into_package()),
            })
        },
        Route::EntryDrop(Family::Animations) => drop_edit!(package.edit_animations()),
        Route::NoopCommit(Family::Animations) => commit_edit!(package.edit_animations()?.commit()),
        Route::EntryDrop(Family::Broadcast) => drop_edit!(package.edit_broadcast(position)),
        Route::NoopCommit(Family::Broadcast) => {
            commit_edit!(package.edit_broadcast(position)?.commit())
        },
        Route::EntryDrop(Family::ClientData) => drop_edit!(package.edit_client_data(target)),
        Route::NoopCommit(Family::ClientData) => {
            commit_edit!(package.edit_client_data(target)?.commit())
        },
        Route::EntryDrop(Family::Diagrams) => drop_edit!(package.edit_diagrams(position)),
        Route::NoopCommit(Family::Diagrams) => {
            commit_edit!(package.edit_diagrams(position)?.commit())
        },
        Route::EntryDrop(Family::DocumentComparison) => {
            drop_edit!(package.edit_document_comparison())
        },
        Route::NoopCommit(Family::DocumentComparison) => {
            commit_edit!(package.edit_document_comparison()?.commit())
        },
        Route::EntryDrop(Family::Objects) => drop_edit!(package.edit_objects()),
        Route::NoopCommit(Family::Objects) => commit_edit!(package.edit_objects()?.commit()),
        Route::EntryDrop(Family::ExternalMedia) => drop_edit!(package.edit_external_media()),
        Route::NoopCommit(Family::ExternalMedia) => {
            commit_edit!(package.edit_external_media()?.commit())
        },
        Route::EntryDrop(Family::Fonts) => drop_edit!(package.edit_fonts()),
        Route::NoopCommit(Family::Fonts) => commit_edit!(package.edit_fonts()?.commit()),
        Route::EntryDrop(Family::MasterMetadata) => {
            drop_edit!(package.edit_master_metadata(master_kind, position))
        },
        Route::NoopCommit(Family::MasterMetadata) => {
            commit_edit!(
                package
                    .edit_master_metadata(master_kind, position)?
                    .commit()
            )
        },
        Route::EntryDrop(Family::MasterLayout) => {
            drop_edit!(package.edit_master_layout(master_kind, position))
        },
        Route::NoopCommit(Family::MasterLayout) => {
            commit_edit!(package.edit_master_layout(master_kind, position)?.commit())
        },
        Route::EntryDrop(Family::PictureBullets) => drop_edit!(package.edit_picture_bullets()),
        Route::NoopCommit(Family::PictureBullets) => {
            commit_edit!(package.edit_picture_bullets()?.commit())
        },
        Route::EntryDrop(Family::SlideSync) => drop_edit!(package.edit_slide_sync(position)),
        Route::NoopCommit(Family::SlideSync) => {
            commit_edit!(package.edit_slide_sync(position)?.commit())
        },
    }
}

fn digest(bytes: &[u8]) -> Value {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    let sha256 = hasher
        .finalize()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    json!({"bytes": bytes.len(), "sha256": sha256})
}

fn assert_noop(
    source: &[u8],
    route: Route,
    result: OperationResult,
) -> Result<Value, Box<dyn Error>> {
    if result.changed || result.patch_noop != Some(true) {
        return Err(format!("{} was not an exact no-op", route.name()).into());
    }
    if let Some(operation_count) = result.operation_count {
        if operation_count != 0 {
            return Err(format!("{} accepted {operation_count} operations", route.name()).into());
        }
    }
    let published = result
        .published
        .ok_or_else(|| format!("{} did not publish a package", route.name()))?;
    let bytes = published.to_bytes()?;
    if bytes != source {
        return Err(format!("{} changed the source bytes", route.name()).into());
    }
    if let Some(before) = result.patch_before.as_deref() {
        if before != source {
            return Err(format!("{} patch has the wrong source endpoint", route.name()).into());
        }
    }
    if let Some(after) = result.patch_after.as_deref() {
        if after != source {
            return Err(format!("{} patch has the wrong target endpoint", route.name()).into());
        }
    }
    Ok(json!({
        "output": digest(&bytes),
        "source_exact": true,
        "changed": false,
        "patch_noop": true,
        "operation_count": result.operation_count,
        // Some edit families do not expose byte endpoints on their no-op
        // patch.  Keep those fields absent/null rather than turning the
        // absence into a false claim; when an endpoint exists it must still
        // match the source exactly.
        "patch_before_exact": result.patch_before.as_ref().map(|before| before == source),
        "patch_after_exact": result.patch_after.as_ref().map(|after| after == source),
        "entry": "commit",
    }))
}

fn verify_route(
    package: &Package,
    source: &[u8],
    route: Route,
    fixture_key: &str,
) -> Result<Value, Box<dyn Error>> {
    let result = run_once(package, route)?;
    let mut proof = if route.is_commit() {
        assert_noop(source, route, result)?
    } else {
        if result.changed || result.patch_noop.is_some() || result.published.is_some() {
            return Err(format!("{} unexpectedly published", route.name()).into());
        }
        json!({
            "output": digest(source),
            "source_exact": true,
            "changed": false,
            "patch_noop": Value::Null,
            "operation_count": Value::Null,
            "patch_before_exact": Value::Null,
            "patch_after_exact": Value::Null,
            "entry": "drop",
        })
    };
    proof["fixture"] = json!({"key": fixture_key, "output": digest(source)});
    Ok(proof)
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
    let output = match args.mode.as_str() {
        "verify" => {
            let routes = ROUTES
                .iter()
                .map(|name| {
                    let route = Route::parse(name)?;
                    let fixture_key = route.fixture_key();
                    let source = fixture_source(fixture_key)?;
                    let package = Package::from_bytes(&source)?;
                    Ok::<_, Box<dyn Error>>((
                        name,
                        verify_route(&package, &source, route, fixture_key)?,
                    ))
                })
                .collect::<Result<Vec<_>, _>>()?;
            let fixtures = FIXTURE_KEYS
                .iter()
                .map(|key| {
                    Ok::<_, Box<dyn Error>>((
                        (*key).to_owned(),
                        json!({
                            "generator": "litchi-ppt-lazy-native-fixture-v2",
                            "output": digest(&fixture_source(key)?),
                        }),
                    ))
                })
                .collect::<Result<serde_json::Map<_, _>, _>>()?;
            json!({
                "schema": SCHEMA,
                "mode": "verify",
                "fixtures": fixtures,
                "routes": routes
                    .into_iter()
                    .map(|(name, proof)| ((*name).to_owned(), proof))
                    .collect::<serde_json::Map<_, _>>(),
                "allocation": if allocator { json!({"status":"measured", "scope":"operation_global_system_allocator"}) } else { unavailable_allocation() },
            })
        },
        "calibrate" => {
            let route = args.route.ok_or("calibrate requires --route")?;
            let fixture_key = route.fixture_key();
            let source = fixture_source(fixture_key)?;
            let package = Package::from_bytes(&source)?;
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
                "fixture": {"key": fixture_key, "output": digest(&source)},
                "postcondition": verify_route(&package, &source, route, fixture_key)?,
            })
        },
        "timed" => {
            let route = args.route.ok_or("timed requires --route")?;
            let fixture_key = route.fixture_key();
            let source = fixture_source(fixture_key)?;
            let package = Package::from_bytes(&source)?;
            let report = timed(&package, &args, route, allocator)?;
            let proof = verify_route(&package, &source, route, fixture_key)?;
            let mut report = report;
            report["fixture"] = json!({"key": fixture_key, "output": digest(&source)});
            report["postcondition"] = proof;
            report
        },
        _ => return Err(format!("unknown probe mode {:?}", args.mode).into()),
    };
    println!("{}", serde_json::to_string(&output)?);
    Ok(())
}
