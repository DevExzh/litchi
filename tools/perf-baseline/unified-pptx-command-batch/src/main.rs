//! Paired PPTX command-batch timing: native primary edits versus `File::batch`.
//!
//! The native leg submits one homogeneous edit and one commit, splitting the
//! shape replacements at the native 256-target call boundary.  The unified leg
//! submits the equivalent distinct direct-path `set` commands in one atomic
//! batch.  Corpus generation, opening, command preparation, warm-up output
//! serialization and equivalence checks are outside the timed intervals.

use litchi_core::{OwnedSource, Position};
use litchi_ops::{BatchItem, BatchMode, File, OpenOptions, Props, SourceMode};
use litchi_pptx::{ShapePropertiesReplacement, ShapeTextReplacement};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::{collections::BTreeMap, error::Error, hint::black_box, sync::Arc, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const DEFAULT_COUNTS: [usize; 5] = [1, 10, 100, 1_000, 10_000];
const NATIVE_CHUNK: usize = 256;
const SELECTOR_PATTERN: &str = "/slide[1]/shape[i]";
const TEXT: &str = "command batch text & <escaped>";
const BOUNDS: (i64, i64, i64, i64) = (914_400, 0, 914_400, 914_400);

#[derive(Clone, Copy, PartialEq, Eq)]
enum OwnerMode {
    Eager,
    Source,
}

impl OwnerMode {
    const fn name(self) -> &'static str {
        match self {
            Self::Eager => "eager",
            Self::Source => "source",
        }
    }
}

#[derive(Clone, Copy)]
enum Operation {
    Text,
    Mixed,
}

impl Operation {
    const fn name(self) -> &'static str {
        match self {
            Self::Text => "text",
            Self::Mixed => "text_x",
        }
    }
}

// The native package types do not share a cloneable enum wrapper.  Keeping the
// owner inline also makes the timer measure the same output ownership cost on
// both native owner modes without an extra balancing allocation.
#[allow(clippy::large_enum_variant)]
enum NativeOwner {
    Eager(litchi_pptx::Package),
    Source(litchi_pptx::source::Package),
}

enum NativePrepared {
    Text(Vec<ShapeTextReplacement<'static>>),
    Mixed(Vec<ShapePropertiesReplacement<'static>>),
}

struct Workload {
    native: NativePrepared,
    // Every command vector is fully owned before timing starts.  `File::batch`
    // consumes one vector per sample, so the timed call does not clone input
    // BatchItems merely to satisfy its generic iterator boundary.
    command_sets: Vec<Vec<BatchItem<'static>>>,
}

#[derive(Clone, Copy)]
struct NativeStats {
    changed: usize,
    calls: usize,
    commits: usize,
}

struct Config {
    owners: Vec<OwnerMode>,
    counts: Vec<usize>,
    rounds: usize,
}

fn corpus(count: usize) -> Result<Vec<u8>> {
    let mut builder = litchi_pptx::builder::Builder::new()?;
    let slide = builder.presentation()?.add_slide()?;
    for index in 0..count {
        slide.add_text_box(&format!("original-{index}"), 0, 0, BOUNDS.2, BOUNDS.3);
    }
    Ok(builder.build()?.to_bytes()?)
}

fn unified_props(operation: Operation) -> Props<'static> {
    match operation {
        Operation::Text => Props::new().with("text", TEXT),
        Operation::Mixed => Props::new().with("text", TEXT).with("x", "1in"),
    }
}

fn command_set(count: usize, operation: Operation) -> Vec<BatchItem<'static>> {
    let props = unified_props(operation);
    (0..count)
        .map(|index| {
            let mut item = BatchItem::new("set");
            item.path = Some(format!("/slide[1]/shape[{}]", index + 1).into());
            item.props = Some(props.clone());
            item
        })
        .collect()
}

fn workload(count: usize, operation: Operation, samples: usize) -> Workload {
    let native = match operation {
        Operation::Text => NativePrepared::Text(
            (0..count)
                .map(|index| ShapeTextReplacement::at(index, TEXT))
                .collect(),
        ),
        Operation::Mixed => NativePrepared::Mixed(
            (0..count)
                .map(|index| {
                    ShapePropertiesReplacement::new(Position::new(index))
                        .with_text(TEXT)
                        .with_bounds(BOUNDS)
                })
                .collect(),
        ),
    };
    let command_sets = (0..samples)
        .map(|_| command_set(count, operation))
        .collect();
    Workload {
        native,
        command_sets,
    }
}

fn open_native(bytes: &[u8], mode: OwnerMode) -> Result<NativeOwner> {
    match mode {
        OwnerMode::Eager => Ok(NativeOwner::Eager(litchi_pptx::Package::from_bytes(bytes)?)),
        OwnerMode::Source => Ok(NativeOwner::Source(
            litchi_pptx::source::Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))?,
        )),
    }
}

fn open_unified(bytes: &[u8], mode: OwnerMode) -> Result<File> {
    let options = OpenOptions::new().source_mode(match mode {
        OwnerMode::Eager => SourceMode::Eager,
        OwnerMode::Source => SourceMode::Auto,
    });
    Ok(litchi_ops::open_bytes_with(bytes.to_vec(), &options)?)
}

fn native_once(
    owner: &NativeOwner,
    prepared: &NativePrepared,
) -> Result<(NativeOwner, NativeStats)> {
    match (owner, prepared) {
        (NativeOwner::Eager(base), NativePrepared::Text(replacements)) => {
            let mut edit = base.edit()?;
            let mut changed = 0;
            let mut calls = 0;
            for chunk in replacements.chunks(NATIVE_CHUNK) {
                changed += edit.set_shape_texts(Position::new(0), chunk)?;
                calls += 1;
            }
            Ok((
                NativeOwner::Eager(edit.commit()?.into_package()),
                NativeStats {
                    changed,
                    calls,
                    commits: 1,
                },
            ))
        },
        (NativeOwner::Eager(base), NativePrepared::Mixed(replacements)) => {
            let mut edit = base.edit()?;
            let mut changed = 0;
            let mut calls = 0;
            for chunk in replacements.chunks(NATIVE_CHUNK) {
                changed += edit.set_shape_properties(Position::new(0), chunk)?;
                calls += 1;
            }
            Ok((
                NativeOwner::Eager(edit.commit()?.into_package()),
                NativeStats {
                    changed,
                    calls,
                    commits: 1,
                },
            ))
        },
        (NativeOwner::Source(base), NativePrepared::Text(replacements)) => {
            let mut edit = base.edit()?;
            let mut changed = 0;
            let mut calls = 0;
            for chunk in replacements.chunks(NATIVE_CHUNK) {
                changed += edit.set_shape_texts(Position::new(0), chunk)?;
                calls += 1;
            }
            Ok((
                NativeOwner::Source(edit.commit()?.into_package()),
                NativeStats {
                    changed,
                    calls,
                    commits: 1,
                },
            ))
        },
        (NativeOwner::Source(base), NativePrepared::Mixed(replacements)) => {
            let mut edit = base.edit()?;
            let mut changed = 0;
            let mut calls = 0;
            for chunk in replacements.chunks(NATIVE_CHUNK) {
                changed += edit.set_shape_properties(Position::new(0), chunk)?;
                calls += 1;
            }
            Ok((
                NativeOwner::Source(edit.commit()?.into_package()),
                NativeStats {
                    changed,
                    calls,
                    commits: 1,
                },
            ))
        },
    }
}

fn native_bytes(owner: &mut NativeOwner) -> Result<Vec<u8>> {
    match owner {
        NativeOwner::Eager(package) => Ok(package.to_bytes()?),
        NativeOwner::Source(package) => Ok(package.to_bytes()?),
    }
}

fn unified_once(file: &File, commands: Vec<BatchItem<'static>>) -> Result<litchi_ops::BatchCommit> {
    Ok(file.batch(commands, BatchMode::Atomic)?)
}

fn validate_batch(commit: &litchi_ops::BatchCommit, count: usize) -> Result<()> {
    if commit.value().len() != count || commit.diagnostics().operations() != count {
        return Err(format!(
            "unified batch accounting mismatch: results={}, operations={}, expected={count}",
            commit.value().len(),
            commit.diagnostics().operations()
        )
        .into());
    }
    if commit.value().iter().any(
        |result| !matches!(result, litchi_ops::ItemResult::Applied(applied) if applied.changed),
    ) {
        return Err("unified batch did not apply every direct shape command".into());
    }
    if !commit.diagnostics().changed() {
        return Err("unified batch produced an unchanged archive".into());
    }
    // These workloads intentionally contain one homogeneous direct shape-set
    // family.  A batch that cannot coalesce that family would turn the 1k/10k
    // measurement into a per-command commit benchmark and would miss the
    // atomic linear-scaling contract this route is meant to exercise.
    if commit.diagnostics().commits() != 1 {
        return Err(format!(
            "homogeneous PPTX command batch did not coalesce to one native commit: {}",
            commit.diagnostics().commits()
        )
        .into());
    }
    Ok(())
}

fn decoded_members(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let archive = ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())?;
    let mut members = BTreeMap::new();
    for name in archive.file_names() {
        members.insert(name.to_owned(), archive.read(name)?);
    }
    Ok(members)
}

fn verify_equivalence(native: &[u8], unified: &[u8]) -> Result<&'static str> {
    if native == unified {
        return Ok("exact_bytes");
    }
    if decoded_members(native)? == decoded_members(unified)? {
        return Ok("decoded_members");
    }
    Err("native and unified outputs differ in bytes and decoded members".into())
}

fn sha256(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn route_name(mode: OwnerMode, leg: &str, operation: Operation) -> String {
    format!(
        "pptx_{}_{}_{}_command_batch_v1",
        mode.name(),
        leg,
        operation.name(),
    )
}

fn measure(
    corpus: &[u8],
    count: usize,
    mode: OwnerMode,
    operation: Operation,
    rounds: usize,
) -> Result<Value> {
    let mut native = open_native(corpus, mode).map_err(|error| format!("native open: {error}"))?;
    let unified = open_unified(corpus, mode).map_err(|error| format!("unified open: {error}"))?;
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows the prepared sample count")?;
    let mut workload = workload(count, operation, sample_count);

    let native_before = native_bytes(&mut native)?;
    let unified_before = unified.to_bytes()?;
    // The source and eager owners may choose different ZIP container metadata;
    // verify the prepared artifacts by exact bytes or decoded members before
    // either route enters its timed warm-up.
    verify_equivalence(corpus, &native_before)?;
    verify_equivalence(corpus, &unified_before)?;

    let warm_commands = workload
        .command_sets
        .pop()
        .ok_or("command preparation unexpectedly empty")?;
    let (mut native_warm, native_warm_stats) = native_once(&native, &workload.native)
        .map_err(|error| format!("native warm-up edit: {error}"))?;
    if native_warm_stats.changed != count
        || native_warm_stats.commits != 1
        || native_warm_stats.calls != count.div_ceil(NATIVE_CHUNK)
    {
        return Err(format!(
            "native warm-up accounting mismatch: changed={}, calls={}, commits={}, expected changed={}, calls={}, commits=1",
            native_warm_stats.changed,
            native_warm_stats.calls,
            native_warm_stats.commits,
            count,
            count.div_ceil(NATIVE_CHUNK)
        )
        .into());
    }
    let native_output = native_bytes(&mut native_warm)?;
    let unified_warm = unified_once(&unified, warm_commands)
        .map_err(|error| format!("unified warm-up batch: {error}"))?;
    validate_batch(&unified_warm, count)?;
    let unified_output = unified_warm.file().to_bytes()?;
    let equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_calls = None;
    let mut native_commits = None;
    let mut unified_commits = None;
    let mut unified_leaf_operations = None;
    let mut unified_snapshot_bytes = None;
    let mut command_sets = workload.command_sets.into_iter();
    for _ in 0..rounds {
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let (output, stats) = native_once(black_box(&native), black_box(&workload.native))?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&output);
                native_ns.push(elapsed);
                native_calls.get_or_insert(stats.calls);
                native_commits.get_or_insert(stats.commits);
                if stats.changed != count {
                    return Err(format!(
                        "native changed {} targets during timing, expected {count}",
                        stats.changed
                    )
                    .into());
                }
                drop(output);
            } else {
                let commands = command_sets.next().ok_or("command preparation exhausted")?;
                let start = Instant::now();
                let output = unified_once(black_box(&unified), commands)?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&output);
                unified_ns.push(elapsed);
                validate_batch(&output, count)?;
                unified_commits.get_or_insert(output.diagnostics().commits());
                unified_leaf_operations.get_or_insert(output.diagnostics().leaf_operations());
                unified_snapshot_bytes.get_or_insert(output.diagnostics().snapshot_bytes());
                drop(output);
            }
        }
    }
    if command_sets.next().is_some() {
        return Err("command preparation has unused samples".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "count": count,
        "owner": mode.name(),
        "operation": operation.name(),
        "path_pattern": SELECTOR_PATTERN,
        "native_route": route_name(mode, "native_primary", operation),
        "unified_route": route_name(mode, "unified_file_batch", operation),
        "batch_mode": "atomic",
        "native_chunk_size": NATIVE_CHUNK,
        "native_changed": count,
        "native_calls": native_calls,
        "native_commits": native_commits,
        "unified_batch_calls": 1,
        "unified_operations": count,
        "unified_commits": unified_commits,
        "unified_leaf_operations": unified_leaf_operations,
        "unified_snapshot_bytes": unified_snapshot_bytes,
        "equivalence": equivalence,
        "native_ns": native_ns,
        "unified_ns": unified_ns,
        "native_median_ns": native_median,
        "unified_median_ns": unified_median,
        "unified_over_native": unified_median as f64 / native_median as f64,
    }))
}

fn parse_counts(value: &str) -> Result<Vec<usize>> {
    let counts = value
        .split(',')
        .map(|part| {
            let count = part.parse::<usize>()?;
            if count == 0 {
                return Err("counts must be positive".into());
            }
            Ok(count)
        })
        .collect::<Result<Vec<_>>>()?;
    if counts.is_empty() {
        return Err("--counts requires at least one positive count".into());
    }
    Ok(counts)
}

fn parse_args() -> Result<Config> {
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut rounds = 9;
    let mut owners = vec![OwnerMode::Eager, OwnerMode::Source];
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        let (key, inline) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(key, value)| (key, Some(value)));
        match key {
            "--help" | "-h" => {
                println!(
                    "usage: unified-pptx-command-batch-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
                );
                std::process::exit(0);
            },
            "--counts" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--counts requires a comma-separated value")?;
                counts = parse_counts(&value)?;
            },
            "--owners" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--owners requires eager, source, or eager,source")?;
                owners.clear();
                for name in value.split(',') {
                    let mode = match name {
                        "eager" => OwnerMode::Eager,
                        "source" => OwnerMode::Source,
                        _ => return Err("unknown owner; expected eager or source".into()),
                    };
                    if owners.contains(&mode) {
                        return Err("duplicate owner".into());
                    }
                    owners.push(mode);
                }
            },
            "--rounds" => {
                let value = inline
                    .map(str::to_owned)
                    .or_else(|| args.next())
                    .ok_or("--rounds requires a value")?;
                rounds = value.parse()?;
            },
            _ => return Err(format!("unknown argument {argument:?}").into()),
        }
    }
    if rounds < 3 {
        return Err("at least three ABBA rounds are required".into());
    }
    Ok(Config {
        counts,
        rounds,
        owners,
    })
}

fn build_provenance() -> Value {
    json!({
        "crate": env!("CARGO_PKG_NAME"),
        "crate_version": env!("CARGO_PKG_VERSION"),
        "rustc": std::env::var("LITCHI_RUSTC_VERSION").unwrap_or_else(|_| "unspecified".to_owned()),
        "target": std::env::var("LITCHI_TARGET").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_commit": std::env::var("LITCHI_SOURCE_COMMIT").unwrap_or_else(|_| "unspecified".to_owned()),
        "build_command": std::env::var("LITCHI_BUILD_COMMAND").unwrap_or_else(|_| "unspecified".to_owned()),
    })
}

fn main() -> Result<()> {
    let config = parse_args()?;
    let mut results = Vec::new();
    let mut failures = Vec::new();
    for count in config.counts.iter().copied() {
        let archive = corpus(count)?;
        let artifact = json!({
            "generator": "litchi_pptx::builder::Builder; one slide; direct text boxes",
            "count": count,
            "bytes": archive.len(),
            "sha256": sha256(&archive),
        });
        for mode in config.owners.iter().copied() {
            for operation in [Operation::Text, Operation::Mixed] {
                let mut result = match measure(&archive, count, mode, operation, config.rounds) {
                    Ok(result) => result,
                    Err(error) => {
                        eprintln!(
                            "failed {} {}: {count} commands: {error}",
                            mode.name(),
                            operation.name()
                        );
                        failures.push(json!({
                            "count": count,
                            "owner": mode.name(),
                            "operation": operation.name(),
                            "artifact": artifact,
                            "error": error.to_string(),
                        }));
                        continue;
                    },
                };
                result["artifact"] = artifact.clone();
                eprintln!(
                    "measured {} {}: {count} commands",
                    mode.name(),
                    operation.name()
                );
                results.push(result);
            }
        }
    }

    let routes = [
        (OwnerMode::Eager, "native_primary", Operation::Text),
        (OwnerMode::Eager, "unified_file_batch", Operation::Text),
        (OwnerMode::Eager, "native_primary", Operation::Mixed),
        (OwnerMode::Eager, "unified_file_batch", Operation::Mixed),
        (OwnerMode::Source, "native_primary", Operation::Text),
        (OwnerMode::Source, "unified_file_batch", Operation::Text),
        (OwnerMode::Source, "native_primary", Operation::Mixed),
        (OwnerMode::Source, "unified_file_batch", Operation::Mixed),
    ]
    .into_iter()
    .map(|(mode, leg, operation)| route_name(mode, leg, operation))
    .collect::<Vec<_>>();
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "pptx",
            "workload": "homogeneous-direct-shape-command-batch",
            "path_pattern": SELECTOR_PATTERN,
            "command": "set",
            "batch_mode": "atomic",
            "counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": routes,
            "timing_scope": "one native primary Edit with 256-target setter chunks plus one commit, or one unified File::batch Atomic call; excludes corpus generation, opening, command/props preparation, warm-up, output serialization/save, equivalence checks and drop",
            "coalescing_contract": "each workload is one homogeneous direct-shape set family; unified diagnostics must report one native commit while preserving one result per command",
            "equivalence_scope": "warm-up outputs are compared by exact bytes, falling back to decoded ZIP members before timing",
            "build_provenance": build_provenance(),
            "results": results,
            "failures": failures,
            "performance_claim": "scoped timing evidence only; full acceptance remains pending",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("benchmark matrix contains failures; completed samples and failures are recorded in JSON".into())
    }
}
