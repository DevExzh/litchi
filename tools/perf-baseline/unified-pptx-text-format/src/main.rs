//! Paired PPTX rich run formatting (record 0252): native primary
//! `Edit::set_text_properties` versus `File::set` on one run.
//!
//! Corpus generation, opening, target and property preparation, warm-up,
//! output serialization and equivalence checks are outside the timed interval.
//! Each timed native sample performs one `set_text_properties` call in one Edit
//! and commits once; each unified sample performs one prepared File::set call.

use litchi_core::{OwnedSource, Position};
use litchi_ops::{File, OpenOptions, Props, SourceMode, TargetArg};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::{collections::BTreeMap, error::Error, hint::black_box, sync::Arc, time::Instant};

type Result<T> = std::result::Result<T, Box<dyn Error>>;

const DEFAULT_COUNTS: [usize; 5] = [1, 10, 100, 1_000, 10_000];
const DEFAULT_ROUNDS: usize = 9;
const TARGET_PREFIX: &str = "/slide[1]/shape[1]/paragraph[1]/run[1]";

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

#[allow(clippy::large_enum_variant)]
enum NativeOwner {
    Eager(litchi_pptx::Package),
    Source(litchi_pptx::source::Package),
}

#[derive(Clone, Copy)]
struct NativeStats {
    changed: bool,
    setter_calls: usize,
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
        slide.add_text_box(&format!("original-{index}"), 0, 0, 914_400, 914_400);
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

fn slide_props() -> Props<'static> {
    Props::new()
        .with("bold", true)
        .with("size", "24pt")
        .with("color", "1F4E79")
}

fn native_edit() -> [litchi_pptx::opened::TextAssignment<'static>; 1] {
    use litchi_pptx::opened::{Rgb, TextAssignment, TextProperties, TextTarget};
    [TextAssignment::new(
        TextTarget::Run {
            shape: Position::new(0),
            paragraph: Position::new(0),
            run: Position::new(0),
        },
        TextProperties::default()
            .with_bold(Some(true))
            .with_size(Some(2_400))
            .with_color(Some(Rgb([0x1F, 0x4E, 0x79]))),
    )]
}

fn native_once(owner: &NativeOwner) -> Result<(NativeOwner, NativeStats)> {
    let slide = Position::new(0);
    match owner {
        NativeOwner::Eager(base) => {
            let mut edit = base.edit()?;
            let changed = edit.set_text_properties(slide, &native_edit())? != 0;
            let package = edit.commit()?.into_package();
            Ok((
                NativeOwner::Eager(package),
                NativeStats {
                    changed,
                    setter_calls: 1,
                    commits: 1,
                },
            ))
        },
        NativeOwner::Source(base) => {
            let mut edit = base.edit()?;
            let changed = edit.set_text_properties(slide, &native_edit())? != 0;
            let package = edit.commit()?.into_package();
            Ok((
                NativeOwner::Source(package),
                NativeStats {
                    changed,
                    setter_calls: 1,
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

fn unified_once<'p, 'v>(file: &File, target: TargetArg<'p>, props: Props<'v>) -> Result<File> {
    Ok(file.set(target, props)?)
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
    format!("pptx_{}_{}_run_format_v1", mode.name(), leg)
}

fn expected_native_stats(stats: NativeStats) -> Result<()> {
    if stats.changed && stats.setter_calls == 1 && stats.commits == 1 {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, setter_calls={}, commits={}, expected true, 1, 1",
        stats.changed, stats.setter_calls, stats.commits
    )
    .into())
}

fn measure(corpus: &[u8], shape_count: usize, mode: OwnerMode, rounds: usize) -> Result<Value> {
    let mut native = open_native(corpus, mode)?;
    let unified = open_unified(corpus, mode)?;
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows prepared sample count")?;

    // Parse and prepare all borrowed target/property inputs before timing.
    let target_text = TARGET_PREFIX.to_owned();
    let mut targets: Vec<TargetArg<'_>> = (0..sample_count)
        .map(|_| TargetArg::parse(target_text.as_str()))
        .collect::<litchi_ops::Result<Vec<_>>>()?;
    let mut property_sets: Vec<Props<'static>> = (0..sample_count).map(|_| slide_props()).collect();

    let native_before = native_bytes(&mut native)?;
    let unified_before = unified.to_bytes()?;
    verify_equivalence(corpus, &native_before)?;
    verify_equivalence(corpus, &unified_before)?;

    let warm_target = targets
        .pop()
        .ok_or("unified target preparation unexpectedly empty")?;
    let warm_props = property_sets
        .pop()
        .ok_or("unified property preparation unexpectedly empty")?;
    let (mut native_warm, native_warm_stats) = native_once(&native)?;
    expected_native_stats(native_warm_stats)?;
    let native_output = native_bytes(&mut native_warm)?;
    let unified_warm = unified_once(&unified, warm_target, warm_props)?;
    let unified_output = unified_warm.to_bytes()?;
    if native_output == corpus || unified_output == corpus {
        return Err("warm-up property edit produced an unchanged archive".into());
    }
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_setter_calls = None;
    let mut native_commits = None;
    let mut targets = targets.into_iter();
    let mut property_sets = property_sets.into_iter();
    for _ in 0..rounds {
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let (output, stats) = native_once(black_box(&native))?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_setter_calls.get_or_insert(stats.setter_calls);
                native_commits.get_or_insert(stats.commits);
                drop(output);
            } else {
                let target = targets
                    .next()
                    .ok_or("unified target preparation exhausted")?;
                let props = property_sets
                    .next()
                    .ok_or("unified property preparation exhausted")?;
                let start = Instant::now();
                let output = unified_once(black_box(&unified), target, props)?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&output);
                unified_ns.push(elapsed);
                drop(output);
            }
        }
    }
    if targets.next().is_some() || property_sets.next().is_some() {
        return Err("prepared unified inputs have unused samples".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "shape_count": shape_count,
        "owner": mode.name(),
        "operation": "run_format",
        "target": TARGET_PREFIX,
        "native_route": route_name(mode, "native_primary"),
        "unified_route": route_name(mode, "unified_file_set"),
        "native_changed": true,
        "native_setter_calls": native_setter_calls,
        "native_commits": native_commits,
        "unified_set_calls": 1,
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
                    "usage: unified-pptx-text-format-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
    for shape_count in config.counts.iter().copied() {
        let archive = match corpus(shape_count) {
            Ok(archive) => archive,
            Err(error) => {
                eprintln!("failed corpus generation: {shape_count} shapes: {error}");
                failures.push(json!({
                    "shape_count": shape_count,
                    "phase": "corpus",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let corpus_sha256 = sha256(&archive);
        let artifact = json!({
            "generator": "litchi_pptx::builder::Builder; one slide; direct text boxes",
            "shape_count": shape_count,
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let mut result = match measure(&archive, shape_count, mode, config.rounds) {
                Ok(result) => result,
                Err(error) => {
                    eprintln!(
                        "failed {} run format: {shape_count} shapes: {error}",
                        mode.name()
                    );
                    failures.push(json!({
                        "shape_count": shape_count,
                        "owner": mode.name(),
                        "operation": "run_format",
                        "artifact": artifact,
                        "error": error.to_string(),
                    }));
                    continue;
                },
            };
            result["artifact"] = artifact.clone();
            eprintln!("measured {} run format: {shape_count} shapes", mode.name());
            results.push(result);
        }
    }

    let routes = [
        (OwnerMode::Eager, "native_primary"),
        (OwnerMode::Eager, "unified_file_set"),
        (OwnerMode::Source, "native_primary"),
        (OwnerMode::Source, "unified_file_set"),
    ]
    .into_iter()
    .map(|(mode, leg)| route_name(mode, leg))
    .collect::<Vec<_>>();
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "pptx",
            "workload": "one-run-format",
            "target": TARGET_PREFIX,
            "properties": { "bold": true, "size": "24pt", "color": "1F4E79" },
            "shape_counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": routes,
            "timing_scope": "one native Edit.set_text_properties call and one commit, or one prepared unified File::set(run path, props) call; excludes corpus generation, opening, target/property preparation, warm-up, output serialization/save, equivalence checks and drop",
            "publication_contract": "native_setter_calls is one and native_commits is one per sample; unified_set_calls is one per sample",
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
