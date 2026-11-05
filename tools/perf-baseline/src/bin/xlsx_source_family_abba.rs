//! Paired before/after evidence for the XLSX source-backed family routes.
//!
//! The regular performance matrix exercises the families independently, but
//! its process order is a single revision at a time.  This runner keeps the
//! operation protocol fixed while alternating two release benchmark binaries
//! in fresh child processes.  Each route therefore measures the complete
//! open, edit, commit, and publication operation on the same deterministic
//! corpus at both revisions.  The optional merge children add the compact
//! source manifest forward/inverse oracle to the same report.

use std::env;
use std::error::Error;
use std::fs::{self, File, OpenOptions};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicU64, Ordering};

use serde::Serialize;
use serde_json::{Map, Value};
use sha2::{Digest, Sha256};

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SCHEMA: &str = "litchi.xlsx.source-family-abba.v1";
const TOOL: &str = "xlsx_source_family_abba";
const MAX_CHILD_JSON_BYTES: u64 = 64 * 1024 * 1024;
const MAX_WARMUP: usize = 100;
const MAX_SAMPLES: usize = 100_000;
const SINK_WRITE_LIMIT: u64 = 64 * 1024;
const FAMILY_ROUTES: [FamilyRoute; 6] = [
    FamilyRoute::new(
        "page_margin",
        "xlsx_eager_page_margin_edit_save",
        "xlsx_source_backed_page_margin_edit_save",
        &[],
    ),
    FamilyRoute::new(
        "defined_names",
        "xlsx_eager_defined_names_edit_save",
        "xlsx_source_backed_defined_names_edit_save",
        &[],
    ),
    FamilyRoute::new(
        "data_validation",
        "xlsx_eager_data_validation_edit_save",
        "xlsx_source_backed_data_validation_edit_save",
        &[],
    ),
    FamilyRoute::new(
        "auto_filter",
        "xlsx_eager_auto_filter_edit_save",
        "xlsx_source_backed_auto_filter_edit_save",
        &[],
    ),
    FamilyRoute::new(
        "conditional_formatting",
        "xlsx_eager_conditional_formatting_edit_save",
        "xlsx_source_backed_conditional_formatting_edit_save",
        &[],
    ),
    FamilyRoute::new(
        "row_visibility",
        "xlsx_eager_row_visibility_edit_save",
        "xlsx_source_backed_row_visibility_edit_save",
        &["--xlsx-row-visibility-shape", "medium"],
    ),
];

static TEMP_COUNTER: AtomicU64 = AtomicU64::new(1);

#[derive(Clone, Copy, Debug)]
struct FamilyRoute {
    name: &'static str,
    eager_case: &'static str,
    source_case: &'static str,
    extra_args: &'static [&'static str],
}

impl FamilyRoute {
    const fn new(
        name: &'static str,
        eager_case: &'static str,
        source_case: &'static str,
        extra_args: &'static [&'static str],
    ) -> Self {
        Self {
            name,
            eager_case,
            source_case,
            extra_args,
        }
    }
}

#[derive(Debug)]
struct Arguments {
    before_bin: PathBuf,
    after_bin: PathBuf,
    merge_before_bin: Option<PathBuf>,
    merge_after_bin: Option<PathBuf>,
    warmup: usize,
    samples: usize,
    out_dir: PathBuf,
    preflight_only: bool,
    family: Option<String>,
    compact: bool,
}

#[derive(Clone, Debug, Serialize)]
struct BinaryIdentity {
    path: String,
    sha256: String,
    bytes: u64,
}

#[derive(Clone, Debug, Serialize)]
struct Observation {
    elapsed_ns: u64,
    output_sha256: Option<String>,
    accepted_bytes: Option<u64>,
    largest_write: Option<u64>,
    source_read_calls: Option<u64>,
    source_read_bytes: Option<u64>,
}

#[derive(Clone, Debug, Serialize)]
struct Statistics {
    count: usize,
    min_ns: u64,
    p50_ns: u64,
    p95_ns: u64,
    max_ns: u64,
    mean_ns: f64,
}

#[derive(Clone, Debug, Serialize)]
struct RouteSummary {
    route: &'static str,
    eager_case: &'static str,
    source_case: &'static str,
    eager_before: Statistics,
    eager_after: Statistics,
    source_before: Statistics,
    source_after: Statistics,
    eager_delta_percent: f64,
    source_delta_percent: f64,
    eager_output_sha256: String,
    source_output_sha256: String,
    source_sink_before: SinkSummary,
    source_sink_after: SinkSummary,
    source_reads: SourceReadSummary,
    eager_blocks: Vec<AbbaBlock>,
    source_blocks: Vec<AbbaBlock>,
}

/// One complete ABBA block, retained in child execution order so a report can
/// distinguish a shared revision cost from a pooled percentile shift.
#[derive(Clone, Debug, Serialize)]
struct AbbaBlock {
    sample: usize,
    legs: [LegObservation; 4],
}

#[derive(Clone, Debug, Serialize)]
struct LegObservation {
    order: usize,
    revision: &'static str,
    observation: Observation,
}

#[derive(Clone, Debug, Serialize)]
struct SourceReadSummary {
    before_calls: u64,
    before_bytes: u64,
    after_calls: u64,
    after_bytes: u64,
}

#[derive(Clone, Debug, Serialize)]
struct SinkSummary {
    accepted_bytes: u64,
    largest_write: u64,
}

#[derive(Clone, Debug, Serialize)]
struct CompactObservation {
    revision: &'static str,
    implementation: &'static str,
    operation: &'static str,
    primary_ns: u64,
    output_sha256: String,
    forward_applied: bool,
    forward_exact_one_part_overlay: bool,
    inverse_restored_complete_package: bool,
    stale_refused: bool,
    foreign_refused: bool,
}

#[derive(Clone, Debug, Serialize)]
struct CompactSummary {
    merge_before: Statistics,
    merge_after: Statistics,
    unmerge_before: Statistics,
    unmerge_after: Statistics,
    forward_output_sha256: String,
    inverse_restored_complete_package: bool,
    stale_refused: bool,
    foreign_refused: bool,
    observations: Vec<CompactObservation>,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    claim: &'static str,
    warmup: usize,
    samples: usize,
    before_binary: BinaryIdentity,
    after_binary: BinaryIdentity,
    routes: Vec<RouteSummary>,
    compact_manifest: Option<CompactSummary>,
}

#[derive(Clone, Debug)]
struct MergeChildObservation {
    primary_ns: u64,
    output_sha256: String,
    forward_applied: bool,
    forward_exact_one_part_overlay: bool,
    inverse_restored_complete_package: bool,
    stale_refused: bool,
    foreign_refused: bool,
}

pub fn run_from_args<I>(arguments: I) -> AnyResult<()>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let arguments = parse_arguments(arguments)?;
    run(arguments)
}

fn main() -> AnyResult<()> {
    run_from_args(env::args_os().skip(1))
}

fn parse_arguments<I>(arguments: I) -> AnyResult<Arguments>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut before_bin = None;
    let mut after_bin = None;
    let mut merge_before_bin = None;
    let mut merge_after_bin = None;
    let mut warmup = 2;
    let mut samples = 20;
    let mut out_dir = None;
    let mut preflight_only = false;
    let mut family = None;
    let mut compact = false;
    let mut arguments = arguments.into_iter();
    while let Some(argument) = arguments.next() {
        let argument = argument.to_string_lossy();
        match argument.as_ref() {
            "--before-bin" => {
                before_bin = Some(PathBuf::from(next_argument(
                    &mut arguments,
                    "--before-bin",
                )?))
            },
            "--after-bin" => {
                after_bin = Some(PathBuf::from(next_argument(&mut arguments, "--after-bin")?))
            },
            "--merge-before-bin" => {
                merge_before_bin = Some(PathBuf::from(next_argument(
                    &mut arguments,
                    "--merge-before-bin",
                )?))
            },
            "--merge-after-bin" => {
                merge_after_bin = Some(PathBuf::from(next_argument(
                    &mut arguments,
                    "--merge-after-bin",
                )?))
            },
            "--warmup" => {
                warmup = next_argument(&mut arguments, "--warmup")?
                    .to_string_lossy()
                    .parse()?
            },
            "--samples" => {
                samples = next_argument(&mut arguments, "--samples")?
                    .to_string_lossy()
                    .parse()?
            },
            "--out-dir" => {
                out_dir = Some(PathBuf::from(next_argument(&mut arguments, "--out-dir")?))
            },
            "--preflight-only" => preflight_only = true,
            "--family" => {
                family = Some(
                    next_argument(&mut arguments, "--family")?
                        .to_string_lossy()
                        .into_owned(),
                )
            },
            "--compact" => compact = true,
            "--help" | "-h" => {
                println!(
                    "usage: xlsx_source_family_abba --before-bin PATH --after-bin PATH [--merge-before-bin PATH --merge-after-bin PATH --compact] [--family NAME] [--warmup N] [--samples N] [--preflight-only] --out-dir PATH"
                );
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    if warmup > MAX_WARMUP || samples == 0 || samples > MAX_SAMPLES {
        return Err("warmup or samples is outside the supported range".into());
    }
    if merge_before_bin.is_some() != merge_after_bin.is_some() {
        return Err("merge-before-bin and merge-after-bin must be supplied together".into());
    }
    Ok(Arguments {
        before_bin: before_bin.ok_or("--before-bin is required")?,
        after_bin: after_bin.ok_or("--after-bin is required")?,
        merge_before_bin,
        merge_after_bin,
        warmup,
        samples,
        out_dir: out_dir.ok_or("--out-dir is required")?,
        preflight_only,
        family,
        compact,
    })
}

fn next_argument<I>(arguments: &mut I, option: &str) -> AnyResult<std::ffi::OsString>
where
    I: Iterator<Item = std::ffi::OsString>,
{
    arguments
        .next()
        .ok_or_else(|| format!("{option} requires a value").into())
}

fn run(arguments: Arguments) -> AnyResult<()> {
    fs::create_dir_all(&arguments.out_dir)?;
    let routes = selected_routes(&arguments)?;
    let before_identity = binary_identity(&arguments.before_bin)?;
    let after_identity = binary_identity(&arguments.after_bin)?;
    let preflight = match preflight_protocols(&arguments, &routes) {
        Ok(preflight) => preflight,
        Err(error) => {
            fs::write(
                arguments.out_dir.join("preflight-failure.txt"),
                error.to_string(),
            )?;
            return Err(error);
        },
    };
    fs::write(
        arguments.out_dir.join("preflight.json"),
        serde_json::to_vec_pretty(&preflight)?,
    )?;
    if arguments.preflight_only {
        return Ok(());
    }
    let mut route_summaries = Vec::with_capacity(routes.len());
    for (index, route) in routes.iter().copied().enumerate() {
        let summary = run_route(&arguments, route)?;
        write_route_checkpoint(&arguments.out_dir, index, &summary)?;
        route_summaries.push(summary);
    }
    let compact_manifest = if arguments.compact {
        arguments
            .merge_before_bin
            .as_deref()
            .zip(arguments.merge_after_bin.as_deref())
            .map(|(before, after)| run_compact(before, after, arguments.warmup, arguments.samples))
            .transpose()?
    } else {
        None
    };
    let report = Report {
        schema: SCHEMA,
        tool: TOOL,
        claim: "paired before/after complete-operation evidence; output and source-contract checks are required, and compact-manifest replay checks are included only when --compact is supplied with merge children",
        warmup: arguments.warmup,
        samples: arguments.samples,
        before_binary: before_identity,
        after_binary: after_identity,
        routes: route_summaries,
        compact_manifest,
    };
    let report_path = arguments.out_dir.join("report.json");
    let bytes = serde_json::to_vec_pretty(&report)?;
    fs::write(report_path, bytes)?;
    Ok(())
}

#[derive(Debug, Serialize)]
struct PreflightSummary {
    family_cases: usize,
    compact_children: usize,
    compact_operations: usize,
}

fn selected_routes(arguments: &Arguments) -> AnyResult<Vec<FamilyRoute>> {
    let routes = match arguments.family.as_deref() {
        None => FAMILY_ROUTES.to_vec(),
        Some(name) => FAMILY_ROUTES
            .iter()
            .copied()
            .filter(|route| route.name == name)
            .collect(),
    };
    if routes.is_empty() {
        return Err(format!(
            "unknown or empty --family selection: {:?}",
            arguments.family
        )
        .into());
    }
    Ok(routes)
}

fn preflight_protocols(
    arguments: &Arguments,
    routes: &[FamilyRoute],
) -> AnyResult<PreflightSummary> {
    let mut family_cases = 0usize;
    for route in routes.iter().copied() {
        for (binary, case, implementation) in [
            (&arguments.before_bin, route.eager_case, "eager"),
            (&arguments.after_bin, route.eager_case, "eager"),
            (&arguments.before_bin, route.source_case, "source-backed"),
            (&arguments.after_bin, route.source_case, "source-backed"),
        ] {
            let _ = invoke_family(binary, case, route.extra_args, "preflight", implementation)?;
            family_cases = family_cases
                .checked_add(1)
                .ok_or("family preflight counter overflow")?;
        }
    }

    let (compact_children, compact_operations) = if !arguments.compact {
        (0, 0)
    } else {
        match (
            arguments.merge_before_bin.as_deref(),
            arguments.merge_after_bin.as_deref(),
        ) {
            (Some(before), Some(after)) => {
                let mut children = 0usize;
                for binary in [before, after] {
                    for operation in ["merge", "unmerge"] {
                        let _ = invoke_merge(binary, operation, "preflight")?;
                        children = children
                            .checked_add(1)
                            .ok_or("compact child preflight counter overflow")?;
                    }
                }
                (children, 2)
            },
            (None, None) => (0, 0),
            _ => return Err("merge child binaries must be supplied as a pair".into()),
        }
    };

    Ok(PreflightSummary {
        family_cases,
        compact_children,
        compact_operations,
    })
}

fn write_route_checkpoint(out_dir: &Path, index: usize, summary: &RouteSummary) -> AnyResult<()> {
    let name = format!("route-{index:02}-{}.json", summary.route);
    fs::write(out_dir.join(name), serde_json::to_vec_pretty(summary)?)?;
    Ok(())
}

fn run_route(arguments: &Arguments, route: FamilyRoute) -> AnyResult<RouteSummary> {
    let eager = run_revision_pair(
        &arguments.before_bin,
        &arguments.after_bin,
        route.eager_case,
        route.extra_args,
        "eager",
        arguments.warmup,
        arguments.samples,
    )?;
    let source = run_revision_pair(
        &arguments.before_bin,
        &arguments.after_bin,
        route.source_case,
        route.extra_args,
        "source-backed",
        arguments.warmup,
        arguments.samples,
    )?;
    ensure_same_output(&eager, &format!("{} eager", route.name))?;
    ensure_same_output(&source, &format!("{} source", route.name))?;
    let eager_output = stable_output(&eager, &format!("{} eager", route.name))?;
    let source_output = stable_output(&source, &format!("{} source", route.name))?;
    if eager_output != source_output {
        return Err(format!(
            "{} eager/source output hashes differ: {eager_output} != {source_output}",
            route.name
        )
        .into());
    }
    let source_sink_before = stable_sink(&source.before, &format!("{} source before", route.name))?;
    let source_sink_after = stable_sink(&source.after, &format!("{} source after", route.name))?;
    let before_reads =
        stable_source_reads(&source.before, &format!("{} source before", route.name))?;
    let after_reads = stable_source_reads(&source.after, &format!("{} source after", route.name))?;
    Ok(RouteSummary {
        route: route.name,
        eager_case: route.eager_case,
        source_case: route.source_case,
        eager_before: statistics(&eager.before_times)?,
        eager_after: statistics(&eager.after_times)?,
        source_before: statistics(&source.before_times)?,
        source_after: statistics(&source.after_times)?,
        eager_delta_percent: delta_percent(&eager.before_times, &eager.after_times)?,
        source_delta_percent: delta_percent(&source.before_times, &source.after_times)?,
        eager_output_sha256: eager_output,
        source_output_sha256: source_output,
        source_sink_before: SinkSummary {
            accepted_bytes: source_sink_before.0,
            largest_write: source_sink_before.1,
        },
        source_sink_after: SinkSummary {
            accepted_bytes: source_sink_after.0,
            largest_write: source_sink_after.1,
        },
        source_reads: SourceReadSummary {
            before_calls: before_reads.0,
            before_bytes: before_reads.1,
            after_calls: after_reads.0,
            after_bytes: after_reads.1,
        },
        eager_blocks: eager.blocks,
        source_blocks: source.blocks,
    })
}

#[derive(Debug)]
struct RevisionPair {
    before: Vec<Observation>,
    after: Vec<Observation>,
    before_times: Vec<u64>,
    after_times: Vec<u64>,
    blocks: Vec<AbbaBlock>,
}

fn run_revision_pair(
    before_bin: &Path,
    after_bin: &Path,
    case: &'static str,
    extra_args: &[&str],
    implementation: &'static str,
    warmup: usize,
    samples: usize,
) -> AnyResult<RevisionPair> {
    let legs = [
        ("before", before_bin),
        ("after", after_bin),
        ("after", after_bin),
        ("before", before_bin),
    ];
    for _ in 0..warmup {
        for (revision, binary) in legs {
            let _ = invoke_family(binary, case, extra_args, revision, implementation)?;
        }
    }
    let mut before = Vec::with_capacity(samples * 2);
    let mut after = Vec::with_capacity(samples * 2);
    let mut blocks = Vec::with_capacity(samples);
    for _ in 0..samples {
        let mut observations = Vec::with_capacity(4);
        for (revision, binary) in legs {
            observations.push(invoke_family(
                binary,
                case,
                extra_args,
                revision,
                implementation,
            )?);
        }
        let sample = blocks.len();
        blocks.push(AbbaBlock {
            sample,
            legs: [
                LegObservation {
                    order: 0,
                    revision: legs[0].0,
                    observation: observations[0].clone(),
                },
                LegObservation {
                    order: 1,
                    revision: legs[1].0,
                    observation: observations[1].clone(),
                },
                LegObservation {
                    order: 2,
                    revision: legs[2].0,
                    observation: observations[2].clone(),
                },
                LegObservation {
                    order: 3,
                    revision: legs[3].0,
                    observation: observations[3].clone(),
                },
            ],
        });
        before.extend([observations[0].clone(), observations[3].clone()]);
        after.extend([observations[1].clone(), observations[2].clone()]);
    }
    let before_times = before.iter().map(|item| item.elapsed_ns).collect();
    let after_times = after.iter().map(|item| item.elapsed_ns).collect();
    Ok(RevisionPair {
        before,
        after,
        before_times,
        after_times,
        blocks,
    })
}

fn invoke_family(
    benchmark_bin: &Path,
    case: &str,
    extra_args: &[&str],
    revision: &'static str,
    implementation: &'static str,
) -> AnyResult<Observation> {
    let temporary = TemporaryJson::new()?;
    let mut command = Command::new(benchmark_bin);
    command
        .arg("--warmup")
        .arg("0")
        .arg("--samples")
        .arg("1")
        .arg("--case")
        .arg(case);
    for argument in extra_args {
        command.arg(argument);
    }
    let output = command
        .arg("--json")
        .arg(&temporary.path)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .output()?;
    if !output.status.success() {
        return Err(format!("child {case} ({revision}) exited with {}", output.status).into());
    }
    let report = read_json(&temporary.path)?;
    parse_family_report(&report, case, implementation)
}

fn parse_family_report(
    report: &Value,
    expected_case: &str,
    implementation: &'static str,
) -> AnyResult<Observation> {
    let results = report
        .get("results")
        .and_then(Value::as_array)
        .ok_or("child report has no results array")?;
    if results.len() != 1 {
        return Err(format!("child {expected_case} returned {} results", results.len()).into());
    }
    let result = results[0]
        .as_object()
        .ok_or("child result is not an object")?;
    if result.get("case").and_then(Value::as_str) != Some(expected_case) {
        return Err(format!("child result case is not {expected_case}").into());
    }
    let elapsed_ns = result
        .get("elapsed_ns")
        .and_then(Value::as_object)
        .and_then(|value| value.get("samples"))
        .and_then(Value::as_array)
        .and_then(|values| values.first())
        .and_then(Value::as_u64)
        .ok_or("child result has no elapsed_ns sample")?;
    let output_sha256 = result
        .get("output_sha256")
        .and_then(Value::as_str)
        .map(str::to_owned);
    if output_sha256.is_none() {
        return Err(format!("child {expected_case} has no output hash").into());
    }
    let (accepted_bytes, largest_write) = result
        .get("sink")
        .and_then(Value::as_object)
        .map(|sink| {
            (
                sink.get("accepted_bytes").and_then(Value::as_u64),
                sink.get("largest_write").and_then(Value::as_u64),
            )
        })
        .unwrap_or((None, None));
    if let Some(largest_write) = largest_write
        && largest_write > SINK_WRITE_LIMIT
    {
        return Err(format!("child {expected_case} exceeded 64 KiB sink writes").into());
    }
    let (source_read_calls, source_read_bytes) = source_reads(result);
    if implementation == "source-backed" && source_read_calls.is_none() {
        return Err(format!("source child {expected_case} has no source counters").into());
    }
    Ok(Observation {
        elapsed_ns,
        output_sha256,
        accepted_bytes,
        largest_write,
        source_read_calls,
        source_read_bytes,
    })
}

fn source_reads(result: &Map<String, Value>) -> (Option<u64>, Option<u64>) {
    let Some(source) = result.get("source").and_then(Value::as_object) else {
        return (None, None);
    };
    let calls = source
        .get("read_calls")
        .and_then(Value::as_array)
        .and_then(|values| values.first())
        .and_then(Value::as_u64)
        .or_else(|| source.get("source_read_calls").and_then(Value::as_u64));
    let bytes = source
        .get("read_bytes")
        .and_then(Value::as_array)
        .and_then(|values| values.first())
        .and_then(Value::as_u64)
        .or_else(|| source.get("source_read_bytes").and_then(Value::as_u64));
    (calls, bytes)
}

fn run_compact(
    before_bin: &Path,
    after_bin: &Path,
    warmup: usize,
    samples: usize,
) -> AnyResult<CompactSummary> {
    let mut observations = Vec::with_capacity(samples * 8);
    for _ in 0..warmup {
        for (revision, binary) in [
            ("before", before_bin),
            ("after", after_bin),
            ("after", after_bin),
            ("before", before_bin),
        ] {
            let _ = invoke_merge(binary, "merge", revision)?;
            let _ = invoke_merge(binary, "unmerge", revision)?;
        }
    }
    for _ in 0..samples {
        for (revision, binary) in [
            ("before", before_bin),
            ("after", after_bin),
            ("after", after_bin),
            ("before", before_bin),
        ] {
            observations.push(CompactObservation::from_child(
                revision,
                "merge",
                invoke_merge(binary, "merge", revision)?,
            ));
            observations.push(CompactObservation::from_child(
                revision,
                "unmerge",
                invoke_merge(binary, "unmerge", revision)?,
            ));
        }
    }
    let merge_before = observations
        .iter()
        .filter(|item| item.operation == "merge" && item.revision == "before")
        .map(|item| item.primary_ns)
        .collect::<Vec<_>>();
    let merge_after = observations
        .iter()
        .filter(|item| item.operation == "merge" && item.revision == "after")
        .map(|item| item.primary_ns)
        .collect::<Vec<_>>();
    let unmerge_before = observations
        .iter()
        .filter(|item| item.operation == "unmerge" && item.revision == "before")
        .map(|item| item.primary_ns)
        .collect::<Vec<_>>();
    let unmerge_after = observations
        .iter()
        .filter(|item| item.operation == "unmerge" && item.revision == "after")
        .map(|item| item.primary_ns)
        .collect::<Vec<_>>();
    let output = observations
        .iter()
        .find(|item| item.operation == "merge")
        .ok_or("compact route returned no merge observation")?
        .output_sha256
        .clone();
    let unmerge_output = observations
        .iter()
        .find(|item| item.operation == "unmerge")
        .ok_or("compact route returned no unmerge observation")?
        .output_sha256
        .clone();
    if observations.iter().any(|item| {
        let expected = if item.operation == "merge" {
            &output
        } else {
            &unmerge_output
        };
        &item.output_sha256 != expected
            || !item.forward_applied
            || !item.forward_exact_one_part_overlay
    }) {
        return Err("compact source forward output or exact overlay checks are unstable".into());
    }
    let inverse_restored_complete_package = observations
        .iter()
        .all(|item| item.inverse_restored_complete_package);
    let stale_refused = observations.iter().all(|item| item.stale_refused);
    let foreign_refused = observations.iter().all(|item| item.foreign_refused);
    if !inverse_restored_complete_package || !stale_refused || !foreign_refused {
        return Err("compact source replay safety checks failed".into());
    }
    Ok(CompactSummary {
        merge_before: statistics(&merge_before)?,
        merge_after: statistics(&merge_after)?,
        unmerge_before: statistics(&unmerge_before)?,
        unmerge_after: statistics(&unmerge_after)?,
        forward_output_sha256: output,
        inverse_restored_complete_package,
        stale_refused,
        foreign_refused,
        observations,
    })
}

impl CompactObservation {
    fn from_child(
        revision: &'static str,
        operation: &'static str,
        child: MergeChildObservation,
    ) -> Self {
        Self {
            revision,
            implementation: "source-backed",
            operation,
            primary_ns: child.primary_ns,
            output_sha256: child.output_sha256,
            forward_applied: child.forward_applied,
            forward_exact_one_part_overlay: child.forward_exact_one_part_overlay,
            inverse_restored_complete_package: child.inverse_restored_complete_package,
            stale_refused: child.stale_refused,
            foreign_refused: child.foreign_refused,
        }
    }
}

fn invoke_merge(
    merge_bin: &Path,
    operation: &str,
    revision: &'static str,
) -> AnyResult<MergeChildObservation> {
    let output = Command::new(merge_bin)
        .arg("--child")
        .arg("--implementation")
        .arg("source")
        .arg("--operation")
        .arg(operation)
        .output()?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!(
            "merge child {operation} ({revision}) exited with {}: {}",
            output.status,
            stderr.trim()
        )
        .into());
    }
    let value: Value = serde_json::from_slice(&output.stdout)?;
    let object = value
        .as_object()
        .ok_or("merge child output is not an object")?;
    let primary_ns = object
        .get("primary_ns")
        .and_then(Value::as_u64)
        .ok_or("merge child output has no primary_ns")?;
    let output_sha256 = object
        .get("output_sha256")
        .and_then(Value::as_str)
        .ok_or("merge child output has no output_sha256")?
        .to_owned();
    let patch = object
        .get("patch")
        .and_then(Value::as_object)
        .ok_or("merge child output has no patch checks")?;
    Ok(MergeChildObservation {
        primary_ns,
        output_sha256,
        forward_applied: required_bool(patch, "forward_applied")?,
        forward_exact_one_part_overlay: required_bool(patch, "forward_exact_one_part_overlay")?,
        inverse_restored_complete_package: required_bool(
            patch,
            "inverse_restored_complete_package",
        )?,
        stale_refused: object
            .get("stale_refused")
            .and_then(Value::as_bool)
            .ok_or("merge child output has no stale_refused")?,
        foreign_refused: required_bool(patch, "foreign_refused")?,
    })
}

fn required_bool(object: &Map<String, Value>, key: &str) -> AnyResult<bool> {
    object
        .get(key)
        .and_then(Value::as_bool)
        .ok_or_else(|| format!("merge child patch has no boolean {key}").into())
}

fn ensure_same_output(pair: &RevisionPair, name: &str) -> AnyResult<()> {
    let hashes = pair
        .before
        .iter()
        .chain(pair.after.iter())
        .filter_map(|item| item.output_sha256.as_deref())
        .collect::<Vec<_>>();
    let first = hashes
        .first()
        .ok_or_else(|| format!("{name} has no output hash"))?;
    if hashes.iter().any(|hash| hash != first) {
        return Err(format!("{name} output hash changed across ABBA legs").into());
    }
    Ok(())
}

fn stable_output(pair: &RevisionPair, name: &str) -> AnyResult<String> {
    pair.before
        .iter()
        .chain(pair.after.iter())
        .find_map(|item| item.output_sha256.clone())
        .ok_or_else(|| format!("{name} has no output hash").into())
}

fn stable_sink(observations: &[Observation], name: &str) -> AnyResult<(u64, u64)> {
    let mut values = Vec::with_capacity(observations.len());
    for item in observations {
        let accepted_bytes = item
            .accepted_bytes
            .ok_or_else(|| format!("{name} has no accepted sink bytes"))?;
        let largest_write = item
            .largest_write
            .ok_or_else(|| format!("{name} has no largest sink write"))?;
        values.push((accepted_bytes, largest_write));
    }
    let first = values
        .first()
        .ok_or_else(|| format!("{name} has no sink values"))?;
    if values.iter().any(|value| value != first) {
        return Err(format!("{name} sink evidence changed across ABBA legs").into());
    }
    Ok(*first)
}

fn stable_source_reads(pair: &[Observation], name: &str) -> AnyResult<(u64, u64)> {
    let mut values = Vec::with_capacity(pair.len());
    for item in pair {
        let calls = item
            .source_read_calls
            .ok_or_else(|| format!("{name} has no source read calls"))?;
        let bytes = item
            .source_read_bytes
            .ok_or_else(|| format!("{name} has no source read bytes"))?;
        values.push((calls, bytes));
    }
    let first = values
        .first()
        .ok_or_else(|| format!("{name} has no source values"))?;
    if values.iter().any(|value| value != first) {
        return Err(format!("{name} source counters changed across ABBA legs").into());
    }
    Ok(*first)
}

fn statistics(values: &[u64]) -> AnyResult<Statistics> {
    if values.is_empty() {
        return Err("statistics requires at least one sample".into());
    }
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    let sum = values.iter().try_fold(0u128, |sum, value| {
        sum.checked_add(u128::from(*value))
            .ok_or("timing sum overflow")
    })?;
    Ok(Statistics {
        count: values.len(),
        min_ns: sorted[0],
        p50_ns: percentile(&sorted, 0.50),
        p95_ns: percentile(&sorted, 0.95),
        max_ns: *sorted.last().expect("nonempty samples"),
        mean_ns: sum as f64 / values.len() as f64,
    })
}

fn percentile(sorted: &[u64], fraction: f64) -> u64 {
    let index = ((sorted.len() - 1) as f64 * fraction).round() as usize;
    sorted[index]
}

fn delta_percent(before: &[u64], after: &[u64]) -> AnyResult<f64> {
    let before = statistics(before)?;
    let after = statistics(after)?;
    if before.p50_ns == 0 {
        return Ok(0.0);
    }
    Ok((after.p50_ns as f64 / before.p50_ns as f64 - 1.0) * 100.0)
}

fn binary_identity(path: &Path) -> AnyResult<BinaryIdentity> {
    let bytes = fs::metadata(path)?.len();
    if bytes == 0 {
        return Err(format!("benchmark binary {} is empty", path.display()).into());
    }
    let mut file = File::open(path)?;
    let mut digest = Sha256::new();
    let mut buffer = [0; 64 * 1024];
    loop {
        let count = file.read(&mut buffer)?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    let digest = digest.finalize();
    Ok(BinaryIdentity {
        path: path.display().to_string(),
        sha256: hex_digest(&digest),
        bytes,
    })
}

fn hex_digest(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        output.push(HEX[usize::from(byte >> 4)] as char);
        output.push(HEX[usize::from(byte & 0x0f)] as char);
    }
    output
}

fn read_json(path: &Path) -> AnyResult<Value> {
    let length = fs::metadata(path)?.len();
    if length == 0 || length > MAX_CHILD_JSON_BYTES {
        return Err(format!("child JSON size {length} is outside the supported range").into());
    }
    let mut file = File::open(path)?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(usize::try_from(length)?)?;
    file.read_to_end(&mut bytes)?;
    Ok(serde_json::from_slice(&bytes)?)
}

struct TemporaryJson {
    path: PathBuf,
}

impl TemporaryJson {
    fn new() -> AnyResult<Self> {
        let directory = env::temp_dir();
        for _ in 0..1000 {
            let counter = TEMP_COUNTER.fetch_add(1, Ordering::Relaxed);
            let path = directory.join(format!(
                "litchi-xlsx-source-family-{}-{counter}.json",
                std::process::id()
            ));
            match OpenOptions::new().write(true).create_new(true).open(&path) {
                Ok(_) => {
                    fs::remove_file(&path)?;
                    return Ok(Self { path });
                },
                Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(error) => return Err(error.into()),
            }
        }
        Err("could not allocate a unique child JSON path".into())
    }
}

impl Drop for TemporaryJson {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}
