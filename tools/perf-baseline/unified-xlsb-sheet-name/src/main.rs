//! Paired XLSB sheet-name publication: native primary edits versus
//! `File::set`.
//!
//! Corpus generation, opening, target/property preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. The
//! native source leg deliberately materializes its prepared
//! `SourceBackedWorkbook` into a package for every edit, matching the source
//! facade's edit boundary.

use litchi_core::OwnedSource;
use litchi_core::sheet::WorkbookTrait;
use litchi_ops::{File, OpenOptions, Props, SourceMode, TargetArg};
use litchi_xlsb::SourceBackedWorkbook;
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::hint::black_box;
use std::io::Cursor;
use std::sync::Arc;
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn Error + Send + Sync>>;

const DEFAULT_COUNTS: [usize; 4] = [1, 10, 100, 1_000];
const DEFAULT_ROUNDS: usize = 9;
const TARGET_PREFIX: &str = "/sheet[1]";
const ORIGINAL_NAME: &str = "Sheet1";
const RENAMED_NAME: &str = "Renamed";

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

enum NativeOwner {
    Eager(Box<litchi_xlsb::Package>),
    Source(SourceBackedWorkbook),
}

#[derive(Clone, Copy)]
struct NativeStats {
    changed: bool,
    rename_calls: usize,
    commits: usize,
}

struct Config {
    owners: Vec<OwnerMode>,
    counts: Vec<usize>,
    rounds: usize,
}

fn corpus(cell_count: usize) -> Result<Vec<u8>> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    let mut sheet = builder::Worksheet::new(ORIGINAL_NAME);
    for index in 0..cell_count {
        let row = u32::try_from(index / 32)?;
        let column = u32::try_from(index % 32)?;
        let value = (index as f64) + 0.5;
        sheet.set_cell(row, column, value);
    }
    writer.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    writer.build()?.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn open_native(bytes: &[u8], mode: OwnerMode) -> Result<NativeOwner> {
    match mode {
        OwnerMode::Eager => Ok(NativeOwner::Eager(Box::new(
            litchi_xlsb::Package::from_bytes(bytes)?,
        ))),
        OwnerMode::Source => Ok(NativeOwner::Source(SourceBackedWorkbook::from_read_at(
            Arc::new(OwnedSource::new(bytes.to_vec())),
        )?)),
    }
}

fn open_unified(bytes: &[u8], mode: OwnerMode) -> Result<File> {
    let source_mode = match mode {
        OwnerMode::Eager => SourceMode::Eager,
        OwnerMode::Source => SourceMode::Auto,
    };
    let options = OpenOptions::new().source_mode(source_mode);
    Ok(litchi_ops::open_bytes_with(bytes.to_vec(), &options)?)
}

fn native_once(owner: &NativeOwner) -> Result<(litchi_xlsb::Package, NativeStats)> {
    match owner {
        NativeOwner::Eager(base) => {
            let mut edit = base.edit()?;
            edit.set_sheet_name(0, RENAMED_NAME.to_owned())?
                .ok_or("native sheet selector unexpectedly missing")?;
            let commit = edit.commit()?;
            let changed = commit.diagnostics().changed();
            let package = commit.into_package();
            Ok((
                package,
                NativeStats {
                    changed,
                    rename_calls: 1,
                    commits: 1,
                },
            ))
        },
        NativeOwner::Source(base) => {
            let package = base.to_workbook()?.into_package();
            let mut edit = package.edit()?;
            edit.set_sheet_name(0, RENAMED_NAME.to_owned())?
                .ok_or("native sheet selector unexpectedly missing")?;
            let commit = edit.commit()?;
            let changed = commit.diagnostics().changed();
            let package = commit.into_package();
            Ok((
                package,
                NativeStats {
                    changed,
                    rename_calls: 1,
                    commits: 1,
                },
            ))
        },
    }
}

fn native_before(owner: &NativeOwner) -> Result<Vec<u8>> {
    match owner {
        NativeOwner::Eager(package) => Ok(package.to_bytes()?),
        NativeOwner::Source(source) => Ok(source.to_workbook()?.into_package().to_bytes()?),
    }
}

fn native_bytes(package: &litchi_xlsb::Package) -> Result<Vec<u8>> {
    Ok(package.to_bytes()?)
}

fn unified_once<'p, 'v>(file: &File, target: TargetArg<'p>, props: Props<'v>) -> Result<File> {
    Ok(file.set(target, props)?)
}

fn unified_bytes(file: &File) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    file.write_to(&mut output)?;
    Ok(output)
}

fn decoded_members(bytes: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let archive = ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())?;
    let mut members = BTreeMap::new();
    for name in archive.file_names() {
        members.insert(name.to_owned(), archive.read(name)?);
    }
    Ok(members)
}

fn sha256(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn member_description(bytes: Option<&Vec<u8>>) -> Value {
    match bytes {
        Some(bytes) => json!({
            "present": true,
            "bytes": bytes.len(),
            "sha256": sha256(bytes),
        }),
        None => json!({"present": false}),
    }
}

/// Compare archives without hiding a byte-level mismatch behind a boolean.
/// Exact bytes are preferred; otherwise every decoded member is compared and
/// the returned report identifies each missing or differing member.
fn verify_equivalence(native: &[u8], unified: &[u8]) -> Result<Value> {
    let native_sha256 = sha256(native);
    let unified_sha256 = sha256(unified);
    if native == unified {
        return Ok(json!({
            "kind": "exact_bytes",
            "native_sha256": native_sha256,
            "unified_sha256": unified_sha256,
            "differing_members": [],
        }));
    }

    let native_members = decoded_members(native)?;
    let unified_members = decoded_members(unified)?;
    let names = native_members
        .keys()
        .chain(unified_members.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let differing_members = names
        .into_iter()
        .filter_map(|name| {
            let left = native_members.get(&name);
            let right = unified_members.get(&name);
            (left != right).then(|| {
                json!({
                    "member": name,
                    "native": member_description(left),
                    "unified": member_description(right),
                })
            })
        })
        .collect::<Vec<_>>();
    if differing_members.is_empty() {
        Ok(json!({
            "kind": "decoded_member_bytes",
            "native_sha256": native_sha256,
            "unified_sha256": unified_sha256,
            "decoded_member_count": native_members.len(),
            "differing_members": [],
        }))
    } else {
        Err(format!(
            "native and unified XLSB outputs differ in decoded ZIP members: {}",
            serde_json::to_string(&differing_members)?
        )
        .into())
    }
}

fn verify_output(bytes: &[u8], expected_cells: usize) -> Result<()> {
    let workbook = litchi_xlsb::Workbook::new(Cursor::new(bytes.to_vec()))?;
    let names = WorkbookTrait::worksheet_names(&workbook);
    if names.len() != 1 || names[0] != RENAMED_NAME {
        return Err(format!("renamed output has worksheet names {names:?}").into());
    }
    let cells = workbook.cell_values(0)?.cells().count();
    if cells != expected_cells {
        return Err(format!("renamed output has {cells} cells, expected {expected_cells}").into());
    }
    Ok(())
}

fn verify_original(bytes: &[u8], expected_cells: usize) -> Result<()> {
    let workbook = litchi_xlsb::Workbook::new(Cursor::new(bytes.to_vec()))?;
    let names = WorkbookTrait::worksheet_names(&workbook);
    if names.len() != 1 || names[0] != ORIGINAL_NAME {
        return Err(format!("corpus has worksheet names {names:?}").into());
    }
    let cells = workbook.cell_values(0)?.cells().count();
    if cells != expected_cells {
        return Err(format!("corpus has {cells} cells, expected {expected_cells}").into());
    }
    Ok(())
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn native_route(mode: OwnerMode) -> String {
    format!("xlsb_{}_native_primary_sheet_rename_v1", mode.name())
}

fn unified_route(mode: OwnerMode) -> String {
    format!("xlsb_{}_unified_file_set_sheet_name_v1", mode.name())
}

fn expected_native_stats(stats: NativeStats) -> Result<()> {
    if stats.changed && stats.rename_calls == 1 && stats.commits == 1 {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, rename_calls={}, commits={}, expected true, 1, 1",
        stats.changed, stats.rename_calls, stats.commits
    )
    .into())
}

fn measure(bytes: &[u8], cell_count: usize, mode: OwnerMode, rounds: usize) -> Result<Value> {
    let native = open_native(bytes, mode)?;
    let unified = open_unified(bytes, mode)?;
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows prepared sample count")?;

    // Parse and prepare all borrowed target/property inputs before timing.
    let target_text = TARGET_PREFIX.to_owned();
    let mut targets: Vec<TargetArg<'_>> = (0..sample_count)
        .map(|_| TargetArg::parse(target_text.as_str()))
        .collect::<litchi_ops::Result<Vec<_>>>()?;
    let mut property_sets: Vec<Props<'static>> = (0..sample_count)
        .map(|_| Props::new().with("name", RENAMED_NAME))
        .collect();

    let native_before = native_before(&native)?;
    let unified_before = unified_bytes(&unified)?;
    verify_original(bytes, cell_count)?;
    verify_equivalence(bytes, &native_before)?;
    verify_equivalence(bytes, &unified_before)?;

    let warm_target = targets
        .pop()
        .ok_or("unified target preparation unexpectedly empty")?;
    let warm_props = property_sets
        .pop()
        .ok_or("unified property preparation unexpectedly empty")?;
    let (native_warm, native_warm_stats) = native_once(&native)?;
    expected_native_stats(native_warm_stats)?;
    let native_output = native_bytes(&native_warm)?;
    let unified_warm = unified_once(&unified, warm_target, warm_props)?;
    let unified_output = unified_bytes(&unified_warm)?;
    if native_output == bytes || unified_output == bytes {
        return Err("warm-up sheet rename produced an unchanged archive".into());
    }
    verify_output(&native_output, cell_count)?;
    verify_output(&unified_output, cell_count)?;
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_rename_calls = None;
    let mut native_commits = None;
    let mut targets = targets.into_iter();
    let mut property_sets = property_sets.into_iter();
    for _ in 0..rounds {
        // ABBA: native, unified, unified, native.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let (output, stats) = native_once(&native)?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_rename_calls.get_or_insert(stats.rename_calls);
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
        "cell_count": cell_count,
        "owner": mode.name(),
        "operation": "sheet_name",
        "target": TARGET_PREFIX,
        "name": RENAMED_NAME,
        "native_route": native_route(mode),
        "unified_route": unified_route(mode),
        "unified_prepared_owner": mode.name(),
        "native_changed": true,
        "native_rename_calls": native_rename_calls,
        "native_commits": native_commits,
        "unified_set_calls": 1,
        "unified_diagnostics": null,
        "warm_equivalence": warm_equivalence,
        "native_output_sha256": sha256(&native_output),
        "unified_output_sha256": sha256(&unified_output),
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
                    "usage: unified-xlsb-sheet-name-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
        "binary_sha256": std::env::var("LITCHI_BINARY_SHA256").unwrap_or_else(|_| "unspecified".to_owned()),
        "build_command": std::env::var("LITCHI_BUILD_COMMAND").unwrap_or_else(|_| "unspecified".to_owned()),
    })
}

fn main() -> Result<()> {
    let config = parse_args()?;
    let mut results = Vec::new();
    let mut failures = Vec::new();
    for cell_count in config.counts.iter().copied() {
        let archive = match corpus(cell_count) {
            Ok(archive) => archive,
            Err(error) => {
                eprintln!("failed corpus generation: {cell_count} cells: {error}");
                failures.push(json!({
                    "cell_count": cell_count,
                    "phase": "corpus",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let corpus_sha256 = sha256(&archive);
        let artifact = json!({
            "generator": "litchi_xlsb::builder::{self, Builder}; one worksheet; row-major numeric cells",
            "cell_count": cell_count,
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let mut result = match measure(&archive, cell_count, mode, config.rounds) {
                Ok(result) => result,
                Err(error) => {
                    eprintln!(
                        "failed {} sheet name: {cell_count} cells: {error}",
                        mode.name()
                    );
                    failures.push(json!({
                        "cell_count": cell_count,
                        "owner": mode.name(),
                        "operation": "sheet_name",
                        "artifact": artifact,
                        "error": error.to_string(),
                    }));
                    continue;
                },
            };
            result["artifact"] = artifact.clone();
            eprintln!("measured {} sheet name: {cell_count} cells", mode.name());
            results.push(result);
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 2,
            "format": "xlsb",
            "workload": "one-sheet-name-rename",
            "target": TARGET_PREFIX,
            "properties": { "name": RENAMED_NAME },
            "cell_counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": [
                "xlsb_eager_native_primary_sheet_rename_v1",
                "xlsb_eager_unified_file_set_sheet_name_v1",
                "xlsb_source_native_primary_sheet_rename_v1",
                "xlsb_source_unified_file_set_sheet_name_v1",
            ],
            "timing_scope": "one native Package::edit + rename_sheet + commit + Commit::into_package call; source-owner samples additionally include SourceBackedWorkbook::to_workbook().into_package() per edit; paired unified samples use a prepared File with SourceMode::Eager or SourceMode::Auto respectively and time one File::set('/sheet[1]', props) call; excludes corpus generation, opening, target/property preparation, warm-up, output serialization/write_to, equivalence checks and drop",
            "publication_contract": "native_rename_calls is one and native_commits is one per sample; unified_set_calls is one per sample",
            "equivalence_scope": "warm-up outputs are compared by exact ZIP bytes, falling back to decoded ZIP member bytes with per-member length and SHA-256 reporting",
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
