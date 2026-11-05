//! Paired PPTX shape removal: native primary edits versus `File::remove`.
//!
//! Corpus generation, opening, target preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. Each
//! timed sample removes one known last shape through one native call and one
//! commit, or one unified path removal call.

use litchi_core::{OwnedSource, Position};
use litchi_ops::{File, OpenOptions, SourceMode, TargetArg};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::{collections::BTreeMap, error::Error, hint::black_box, sync::Arc, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const DEFAULT_COUNTS: [usize; 5] = [1, 10, 100, 1_000, 10_000];
const DEFAULT_ROUNDS: usize = 9;
const SELECTOR_PREFIX: &str = "/slide[1]/shape[";
const SELECTOR_SUFFIX: &str = "]";
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

// Keep benchmark outputs inline so the timer does not include an artificial
// heap allocation merely to balance these two native owner representations.
#[allow(clippy::large_enum_variant)]
enum NativeOwner {
    Eager(litchi_pptx::Package),
    Source(litchi_pptx::source::Package),
}

#[derive(Clone, Copy)]
struct NativeStats {
    changed: bool,
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

fn native_once(owner: &NativeOwner, shape: Position) -> Result<(NativeOwner, NativeStats)> {
    let slide = Position::new(0);
    match owner {
        NativeOwner::Eager(base) => {
            let mut edit = base.edit()?;
            let changed = edit.remove_shape(slide, shape.get())?;
            let package = edit.commit()?.into_package();
            Ok((
                NativeOwner::Eager(package),
                NativeStats {
                    changed,
                    calls: 1,
                    commits: 1,
                },
            ))
        },
        NativeOwner::Source(base) => {
            let mut edit = base.edit()?;
            let changed = edit.remove_shape(slide, shape.get())?;
            let package = edit.commit()?.into_package();
            Ok((
                NativeOwner::Source(package),
                NativeStats {
                    changed,
                    calls: 1,
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

fn unified_once<'p>(file: &File, target: TargetArg<'p>) -> Result<File> {
    Ok(file.remove(target)?)
}

fn decoded_members(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let archive = ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())?;
    let mut members = BTreeMap::new();
    for name in archive.file_names() {
        members.insert(name.to_owned(), archive.read(name)?);
    }
    Ok(members)
}

fn verify_equivalence(left: &[u8], right: &[u8]) -> Result<&'static str> {
    if left == right {
        return Ok("exact_bytes");
    }
    if decoded_members(left)? == decoded_members(right)? {
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

fn route_name(mode: OwnerMode, leg: &str) -> String {
    format!("pptx_{}_{}_shape_removal_v1", mode.name(), leg)
}

fn expected_native_stats(stats: NativeStats) -> Result<()> {
    if stats.changed && stats.calls == 1 && stats.commits == 1 {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, calls={}, commits={}, expected changed=true, calls=1, commits=1",
        stats.changed, stats.calls, stats.commits
    )
    .into())
}

fn measure(corpus: &[u8], count: usize, mode: OwnerMode, rounds: usize) -> Result<Value> {
    let mut native = open_native(corpus, mode)?;
    let unified = open_unified(corpus, mode)?;
    let shape = Position::new(count - 1);
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows the prepared target count")?;
    // Parse each bounded target before timing. The backing String lives
    // through the complete measurement for this matrix case.
    let target_text = format!("{SELECTOR_PREFIX}{count}{SELECTOR_SUFFIX}");
    let mut targets = (0..sample_count)
        .map(|_| TargetArg::parse(target_text.as_str()))
        .collect::<std::result::Result<Vec<_>, _>>()?;

    let native_before = native_bytes(&mut native)?;
    let unified_before = unified.to_bytes()?;
    verify_equivalence(corpus, &native_before)?;
    verify_equivalence(corpus, &unified_before)?;

    let warm_target = targets
        .pop()
        .ok_or("unified target preparation unexpectedly empty")?;
    let (mut native_warm, native_warm_stats) = native_once(&native, shape)?;
    expected_native_stats(native_warm_stats)?;
    let native_output = native_bytes(&mut native_warm)?;
    let unified_warm = unified_once(&unified, warm_target)?;
    let unified_output = unified_warm.to_bytes()?;
    if native_output == corpus || unified_output == corpus {
        return Err("warm-up removal produced an unchanged archive".into());
    }
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_calls = None;
    let mut native_commits = None;
    let mut targets = targets.into_iter();
    for _ in 0..rounds {
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let (output, stats) = native_once(black_box(&native), black_box(shape))?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_calls.get_or_insert(stats.calls);
                native_commits.get_or_insert(stats.commits);
                drop(output);
            } else {
                let target = targets
                    .next()
                    .ok_or("unified target preparation exhausted")?;
                let start = Instant::now();
                let output = unified_once(black_box(&unified), target)?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&output);
                unified_ns.push(elapsed);
                drop(output);
            }
        }
    }
    if targets.next().is_some() {
        return Err("unified target preparation has unused samples".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "count": count,
        "owner": mode.name(),
        "operation": "shape_remove",
        "native_shape_position": shape.get(),
        "unified_target": format!("{SELECTOR_PREFIX}{count}{SELECTOR_SUFFIX}"),
        "native_route": route_name(mode, "native_primary"),
        "unified_route": route_name(mode, "unified_file_remove"),
        "native_changed": true,
        "native_calls": native_calls,
        "native_commits": native_commits,
        "unified_remove_calls": 1,
        "unified_diagnostics": null,
        "warm_equivalence": warm_equivalence,
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
    let mut rounds = DEFAULT_ROUNDS;
    let mut owners = vec![OwnerMode::Eager, OwnerMode::Source];
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        let (key, inline) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(key, value)| (key, Some(value)));
        match key {
            "--help" | "-h" => {
                println!(
                    "usage: unified-pptx-shape-removal-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
    if owners.is_empty() {
        return Err("--owners requires at least one owner".into());
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
        let archive = match corpus(count) {
            Ok(archive) => archive,
            Err(error) => {
                eprintln!("failed corpus generation: {count} shapes: {error}");
                failures.push(json!({
                    "count": count,
                    "phase": "corpus",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let corpus_sha256 = sha256(&archive);
        let artifact = json!({
            "generator": "litchi_pptx::builder::Builder; one slide; direct text boxes",
            "count": count,
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let mut result = match measure(&archive, count, mode, config.rounds) {
                Ok(result) => result,
                Err(error) => {
                    eprintln!(
                        "failed {} shape removal: {count} shapes: {error}",
                        mode.name()
                    );
                    failures.push(json!({
                        "count": count,
                        "owner": mode.name(),
                        "operation": "shape_remove",
                        "artifact": artifact,
                        "error": error.to_string(),
                    }));
                    continue;
                },
            };
            result["artifact"] = artifact.clone();
            eprintln!("measured {} shape removal: {count} shapes", mode.name());
            results.push(result);
        }
    }

    let routes = [
        (OwnerMode::Eager, "native_primary"),
        (OwnerMode::Eager, "unified_file_remove"),
        (OwnerMode::Source, "native_primary"),
        (OwnerMode::Source, "unified_file_remove"),
    ]
    .into_iter()
    .map(|(mode, leg)| route_name(mode, leg))
    .collect::<Vec<_>>();
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "pptx",
            "workload": "one-slide-shape-removal",
            "selector_template": "/slide[1]/shape[N]",
            "counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": routes,
            "timing_scope": "one native primary Edit.remove_shape(Position(0), count-1) call plus one commit, or one prepared unified File::remove('/slide[1]/shape[N]') call; excludes corpus generation, opening, target preparation, warm-up, output serialization/save, equivalence checks and drop",
            "publication_contract": "native_calls and native_commits are one per sample; unified_remove_calls is one per sample; one top-level shape is removed per sample",
            "equivalence_scope": "warm-up outputs are compared by exact bytes, falling back to decoded ZIP members before timing",
            "build_provenance": build_provenance(),
            "results": results,
            "failure_matrix": failures,
            "performance_claim": "scoped timing evidence only; full acceptance remains pending",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("benchmark matrix contains failures; completed samples and failures are recorded in JSON".into())
    }
}
