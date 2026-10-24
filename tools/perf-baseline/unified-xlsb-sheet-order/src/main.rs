//! Paired XLSB sheet-order publication: native complete-catalog permutation
//! versus one unified `File::move_to` call.
//!
//! Corpus generation, opening, permutation/path preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. The
//! source native leg deliberately promotes its prepared
//! `SourceBackedWorkbook` into a package inside every timed edit, matching the
//! source facade's promotion boundary.

use litchi_core::selector::Path;
use litchi_core::{OwnedSource, Position};
use litchi_ops::{File, OpenOptions, Placement, SourceMode};
use litchi_xlsb::SourceBackedWorkbook;
use litchi_xlsb::cell_values::Value;
use litchi_xlsb::sheet::Selector;
use serde_json::{Value as JsonValue, json};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::hint::black_box;
use std::io::Cursor;
use std::sync::Arc;
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn Error + Send + Sync>>;

const DEFAULT_COUNTS: [usize; 4] = [2, 10, 100, 1_000];
const DEFAULT_ROUNDS: usize = 9;
const TARGET_PATH: &str = "/sheet[1]";

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
    reorder_calls: usize,
    commits: usize,
}

struct PreparedMove {
    path: Path<'static>,
    placement: Placement<'static>,
}

struct Config {
    owners: Vec<OwnerMode>,
    counts: Vec<usize>,
    rounds: usize,
}

#[derive(Debug, PartialEq)]
struct CatalogSnapshot {
    names: Vec<String>,
    markers: Vec<f64>,
    active_catalog_position: Option<usize>,
    active_worksheet_index: Option<usize>,
    active_sheet_index: usize,
}

fn sheet_name(index: usize) -> String {
    format!("Sheet{index:05}")
}

fn sheet_marker(index: usize) -> f64 {
    index as f64 + 0.25
}

fn permutation(tab_count: usize) -> Vec<usize> {
    (1..tab_count).chain(std::iter::once(0)).collect()
}

fn expected_names(tab_count: usize) -> Vec<String> {
    (0..tab_count).map(sheet_name).collect()
}

fn corpus(tab_count: usize) -> Result<Vec<u8>> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    for index in 0..tab_count {
        let mut sheet = builder::Worksheet::new(sheet_name(index));
        // Every worksheet owns one distinct marker.  Reordering must move the
        // complete worksheet identity, rather than merely changing names.
        sheet.set_cell(0, 0, sheet_marker(index));
        writer.add_worksheet(sheet);
    }
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

fn native_edit(
    package: &litchi_xlsb::Package,
    selectors: &[Selector<'_>],
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    let mut edit = package.edit()?;
    // The selectors are a complete, checked semantic catalog permutation. The
    // benchmark wrapper owns one prepared vector per sample and only borrows
    // it for the leaf call, so no input clone is hidden in the timer.
    let _ = edit.reorder_sheets(selectors)?;
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let package = commit.into_package();
    Ok((
        package,
        NativeStats {
            changed,
            reorder_calls: 1,
            commits: 1,
        },
    ))
}

fn native_once(
    owner: &NativeOwner,
    selectors: Vec<Selector<'static>>,
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    match owner {
        NativeOwner::Eager(base) => native_edit(base, &selectors),
        NativeOwner::Source(base) => {
            // Source promotion is intentionally inside this function, which
            // is called inside the timer for every source-owner sample.
            let package = base.to_workbook()?.into_package();
            native_edit(&package, &selectors)
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

fn prepared_moves(sample_count: usize) -> Result<Vec<PreparedMove>> {
    (0..sample_count)
        .map(|_| {
            Ok(PreparedMove {
                path: Path::parse(TARGET_PATH)?.into_owned(),
                placement: Placement::End,
            })
        })
        .collect()
}

fn unified_once(file: &File, operation: PreparedMove) -> Result<File> {
    // `Path` is parsed and owned before the timer; the move consumes that
    // prepared path and placement without cloning either input.
    Ok(file.move_to(operation.path, operation.placement)?)
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

fn member_description(bytes: Option<&Vec<u8>>) -> JsonValue {
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
fn verify_equivalence(native: &[u8], unified: &[u8]) -> Result<JsonValue> {
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

fn marker(snapshot: &litchi_xlsb::cell_values::Snapshot) -> Result<f64> {
    let mut cells = snapshot.cells();
    let cell = cells.next().ok_or("worksheet marker cell is missing")?;
    if cells.next().is_some() {
        return Err("worksheet has more than one stored cell".into());
    }
    match cell.value() {
        Value::Number(value) | Value::RkNumber(value) | Value::FormulaNumberCache(value) => {
            Ok(*value)
        },
        value => Err(format!("worksheet marker has unexpected value {value:?}").into()),
    }
}

fn catalog_snapshot(bytes: &[u8]) -> Result<CatalogSnapshot> {
    let package = litchi_xlsb::Package::from_bytes(bytes)?;
    let workbook = package.workbook()?;
    let tabs = workbook
        .sheets()
        .map(|tab| {
            (
                tab.name().to_owned(),
                tab.worksheet_position().map(|p| p.get()),
            )
        })
        .collect::<Vec<_>>();
    let names = tabs
        .iter()
        .map(|(name, _)| name.clone())
        .collect::<Vec<_>>();
    let mut markers = Vec::with_capacity(tabs.len());
    for (_, worksheet_position) in tabs {
        let worksheet_position =
            worksheet_position.ok_or("generated catalog has a non-worksheet tab")?;
        markers.push(marker(&workbook.cell_values(worksheet_position)?)?);
    }
    Ok(CatalogSnapshot {
        names,
        markers,
        active_catalog_position: workbook.active_catalog_position(),
        active_worksheet_index: workbook.active_worksheet_index(),
        active_sheet_index: litchi_core::sheet::WorkbookTrait::active_sheet_index(workbook),
    })
}

fn verify_original(bytes: &[u8], tab_count: usize) -> Result<()> {
    let observed = catalog_snapshot(bytes)?;
    let names = expected_names(tab_count);
    let markers = (0..tab_count).map(sheet_marker).collect::<Vec<_>>();
    if observed.names != names || observed.markers != markers {
        return Err(format!("corpus catalog or markers differ: {observed:?}").into());
    }
    if observed.active_catalog_position != Some(0)
        || observed.active_worksheet_index != Some(0)
        || observed.active_sheet_index != 0
    {
        return Err(format!(
            "corpus active identity/index is {:?}; expected Sheet00000 at index 0",
            observed
        )
        .into());
    }
    Ok(())
}

fn verify_output(bytes: &[u8], tab_count: usize) -> Result<()> {
    let observed = catalog_snapshot(bytes)?;
    let expected_order = (1..tab_count)
        .chain(std::iter::once(0))
        .map(sheet_name)
        .collect::<Vec<_>>();
    let expected_markers = (1..tab_count)
        .chain(std::iter::once(0))
        .map(sheet_marker)
        .collect::<Vec<_>>();
    if observed.names != expected_order || observed.markers != expected_markers {
        return Err(format!("reordered catalog or markers differ: {observed:?}").into());
    }
    let expected_active = tab_count - 1;
    if observed.active_catalog_position != Some(expected_active)
        || observed.active_worksheet_index != Some(expected_active)
        || observed.active_sheet_index != expected_active
    {
        return Err(format!(
            "active identity/index is {:?}; expected Sheet00000 at catalog and worksheet index {expected_active}",
            observed
        )
        .into());
    }
    Ok(())
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn native_route(mode: OwnerMode) -> String {
    format!("xlsb_{}_native_sheet_order_permutation_v1", mode.name())
}

fn unified_route(mode: OwnerMode) -> String {
    format!("xlsb_{}_unified_file_move_sheet_to_end_v1", mode.name())
}

fn expected_native_stats(stats: NativeStats) -> Result<()> {
    if stats.changed && stats.reorder_calls == 1 && stats.commits == 1 {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, reorder_calls={}, commits={}, expected true, 1, 1",
        stats.changed, stats.reorder_calls, stats.commits
    )
    .into())
}

fn measure(bytes: &[u8], tab_count: usize, mode: OwnerMode, rounds: usize) -> Result<JsonValue> {
    let native = open_native(bytes, mode)?;
    let unified = open_unified(bytes, mode)?;
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows prepared sample count")?;

    // Build one owned semantic permutation and one owned parsed path for every
    // warm-up/sample slot before timing.  The per-operation ownership costs are
    // therefore visible in the result instead of being accidentally charged to
    // one side of the pair.
    let canonical_selectors = permutation(tab_count)
        .into_iter()
        .map(|position| Selector::Position(Position::new(position)))
        .collect::<Vec<Selector<'static>>>();
    let mut native_inputs = (0..sample_count)
        .map(|_| canonical_selectors.clone())
        .collect::<Vec<_>>();
    let mut unified_inputs = prepared_moves(sample_count)?;

    let native_before = native_before(&native)?;
    let unified_before = unified_bytes(&unified)?;
    verify_original(bytes, tab_count)?;
    verify_equivalence(bytes, &native_before)?;
    verify_equivalence(bytes, &unified_before)?;

    let warm_native_input = native_inputs
        .pop()
        .ok_or("native permutation preparation unexpectedly empty")?;
    let warm_unified_input = unified_inputs
        .pop()
        .ok_or("unified path preparation unexpectedly empty")?;
    let (native_warm, native_warm_stats) = native_once(&native, warm_native_input)?;
    expected_native_stats(native_warm_stats)?;
    let native_output = native_bytes(&native_warm)?;
    let unified_warm = unified_once(&unified, warm_unified_input)?;
    let unified_output = unified_bytes(&unified_warm)?;
    if native_output == bytes || unified_output == bytes {
        return Err("warm-up sheet reorder produced an unchanged archive".into());
    }
    verify_output(&native_output, tab_count)?;
    verify_output(&unified_output, tab_count)?;
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_reorder_calls = None;
    let mut native_commits = None;
    for _ in 0..rounds {
        // ABBA: native, unified, unified, native.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let input = native_inputs
                    .pop()
                    .ok_or("prepared native permutation samples exhausted")?;
                let start = Instant::now();
                let (output, stats) = native_once(black_box(&native), input)?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_reorder_calls.get_or_insert(stats.reorder_calls);
                native_commits.get_or_insert(stats.commits);
                drop(output);
            } else {
                let input = unified_inputs
                    .pop()
                    .ok_or("prepared unified path samples exhausted")?;
                let start = Instant::now();
                let output = unified_once(black_box(&unified), input)?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&output);
                unified_ns.push(elapsed);
                drop(output);
            }
        }
    }
    if !native_inputs.is_empty() || !unified_inputs.is_empty() {
        return Err("prepared order-operation samples have unused inputs".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "tab_count": tab_count,
        "owner": mode.name(),
        "operation": "sheet_order_permutation",
        "target": TARGET_PATH,
        "placement": "end",
        "permutation": permutation(tab_count),
        "native_route": native_route(mode),
        "unified_route": unified_route(mode),
        "unified_prepared_owner": mode.name(),
        "native_changed": native_warm_stats.changed,
        "native_reorder_calls": native_reorder_calls,
        "native_commits": native_commits,
        "unified_move_calls": 1,
        "active_identity_before": "Sheet00000",
        "active_catalog_position_before": 0,
        "active_identity_after": "Sheet00000",
        "active_catalog_position_after": tab_count - 1,
        "warm_equivalence": warm_equivalence,
        "native_output_sha256": sha256(&native_output),
        "unified_output_sha256": sha256(&unified_output),
        "native_ns": native_ns,
        "unified_ns": unified_ns,
        "native_median_ns": native_median,
        "unified_median_ns": unified_median,
        "unified_over_native": unified_median as f64 / native_median as f64,
        "input_cost": {
            "sample_count_including_warmup": sample_count,
            "native": {
                "permutation_selectors_prepared_outside_timing": true,
                "permutation_selector_count": tab_count,
                "per_sample_owned_input_clones_outside_timing": sample_count,
                "timed_input_clones": 0,
                "timed_call_boundary": "benchmark consumes Vec<Selector>; native Edit::reorder_sheets borrows &[Selector]"
            },
            "unified": {
                "path_prepared_outside_timing": true,
                "path": TARGET_PATH,
                "per_sample_owned_input_parses_outside_timing": sample_count,
                "timed_input_clones": 0,
                "timed_call_boundary": "benchmark consumes owned Path<'static> and Placement::End"
            },
            "pairing": "one owned operation input is prepared per sample on each leg; no permutation/path clone or parse occurs inside a timed call"
        }
    }))
}

fn parse_counts(value: &str) -> Result<Vec<usize>> {
    let counts = value
        .split(',')
        .map(|part| {
            let count = part.parse::<usize>()?;
            if count < 2 {
                return Err("counts must be at least two tabs".into());
            }
            Ok(count)
        })
        .collect::<Result<Vec<_>>>()?;
    if counts.is_empty() {
        return Err("--counts requires at least one count".into());
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
                    "usage: unified-xlsb-sheet-order-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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

fn build_provenance() -> JsonValue {
    json!({
        "crate": env!("CARGO_PKG_NAME"),
        "crate_version": env!("CARGO_PKG_VERSION"),
        "rustc": std::env::var("LITCHI_RUSTC_VERSION").unwrap_or_else(|_| "unspecified".to_owned()),
        "target": std::env::var("LITCHI_TARGET").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_commit": std::env::var("LITCHI_SOURCE_COMMIT").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_digest": std::env::var("LITCHI_SOURCE_DIGEST").unwrap_or_else(|_| "unspecified".to_owned()),
        "binary_sha256": std::env::var("LITCHI_BINARY_SHA256").unwrap_or_else(|_| "unspecified".to_owned()),
        "build_command": std::env::var("LITCHI_BUILD_COMMAND").unwrap_or_else(|_| "unspecified".to_owned()),
        "machine": {
            "os": std::env::consts::OS,
            "arch": std::env::consts::ARCH,
            "parallelism": std::thread::available_parallelism().map_or(0, |value| value.get()),
            "label": std::env::var("LITCHI_MACHINE").unwrap_or_else(|_| "unspecified".to_owned()),
        },
    })
}

fn main() -> Result<()> {
    let config = parse_args()?;
    let mut results = Vec::new();
    let mut failures = Vec::new();
    for tab_count in config.counts.iter().copied() {
        let archive = match corpus(tab_count) {
            Ok(archive) => archive,
            Err(error) => {
                eprintln!("failed corpus generation: {tab_count} tabs: {error}");
                failures.push(json!({
                    "tab_count": tab_count,
                    "phase": "corpus",
                    "operation": "sheet_order_permutation",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let corpus_sha256 = sha256(&archive);
        let artifact = json!({
            "generator": "litchi_xlsb::builder::{self, Builder}; N worksheets in catalog order; one distinct numeric marker per worksheet",
            "tab_count": tab_count,
            "permutation": permutation(tab_count),
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let mut result = match measure(&archive, tab_count, mode, config.rounds) {
                Ok(result) => result,
                Err(error) => {
                    eprintln!(
                        "failed {} sheet-order permutation: {tab_count} tabs: {error}",
                        mode.name()
                    );
                    failures.push(json!({
                        "tab_count": tab_count,
                        "owner": mode.name(),
                        "operation": "sheet_order_permutation",
                        "artifact": artifact,
                        "error": error.to_string(),
                    }));
                    continue;
                },
            };
            result["artifact"] = artifact.clone();
            eprintln!(
                "measured {} sheet-order permutation: {tab_count} tabs",
                mode.name()
            );
            results.push(result);
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "xlsb",
            "workload": "complete-sheet-catalog-order-permutation-first-to-end",
            "target": TARGET_PATH,
            "placement": "end",
            "counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": config
                .owners
                .iter()
                .flat_map(|mode| [native_route(*mode), unified_route(*mode)])
                .collect::<Vec<_>>(),
            "timing_scope": "one native Package::edit + complete semantic Edit::reorder_sheets permutation + one commit + Commit::into_package call; source-owner native samples additionally include SourceBackedWorkbook::to_workbook().into_package() inside each timed edit; one unified File::move_to(prepared /sheet[1], Placement::End) call with source promotion inside the source facade leg; excludes corpus generation, opening, permutation/path preparation, warm-up, output serialization/write_to, equivalence checks and drop",
            "publication_contract": "native reorder_calls is one and native_commits is one per sample; unified_move_calls is one per sample; each result verifies complete catalog order, one distinct preserved marker per worksheet, and active Sheet00000 identity at its final catalog and worksheet index",
            "equivalence_scope": "warm-up outputs are compared by exact ZIP bytes, falling back to decoded ZIP member bytes with per-member length and SHA-256 reporting; differing member bytes fail the matrix",
            "build_provenance": build_provenance(),
            "results": results,
            "failure_matrix": failures,
            "performance_claim": "scoped paired sheet-order timing evidence only; no batch-scaling claim and full acceptance remains pending",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("benchmark matrix contains failures; completed samples and failures are recorded in JSON".into())
    }
}
