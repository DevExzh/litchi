//! Paired XLSB blank-sheet insertion: native `Edit::insert_sheet` versus one
//! unified `File::add` call.
//!
//! Corpus generation, opening, position/path preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. The
//! source native leg deliberately promotes its prepared `SourceBackedWorkbook`
//! into a package inside every timed edit, matching the source facade's
//! promotion boundary.

use litchi_core::selector::Path;
use litchi_core::{OwnedSource, Position};
use litchi_ops::{File, NewNode, OpenOptions, Placement, SourceMode};
use litchi_xlsb::SourceBackedWorkbook;
use litchi_xlsb::cell_values::Value;
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

const DEFAULT_COUNTS: [usize; 4] = [1, 10, 100, 1_000];
const DEFAULT_ROUNDS: usize = 9;
const INSERTED_NAME: &str = "Inserted";

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
    insert_calls: usize,
    commits: usize,
}

struct PreparedAdd {
    parent: Path<'static>,
    node: NewNode<'static>,
}

struct Config {
    owners: Vec<OwnerMode>,
    counts: Vec<usize>,
    rounds: usize,
}

#[derive(Debug, PartialEq)]
struct CatalogSnapshot {
    names: Vec<String>,
    markers: Vec<Option<f64>>,
    visible: Vec<bool>,
    binary_index: Vec<bool>,
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

fn expected_names(tab_count: usize) -> Vec<String> {
    (0..tab_count).map(sheet_name).collect()
}

fn insertion_position(tab_count: usize) -> usize {
    tab_count / 2
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
    input: (Position, String),
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    let mut edit = package.edit()?;
    edit.insert_sheet(input.0, input.1, litchi_xlsb::sheet::State::Visible)?;
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let package = commit.into_package();
    Ok((
        package,
        NativeStats {
            changed,
            insert_calls: 1,
            commits: 1,
        },
    ))
}

fn native_once(
    owner: &NativeOwner,
    input: (Position, String),
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    match owner {
        NativeOwner::Eager(base) => native_edit(base, input),
        NativeOwner::Source(base) => {
            // Source promotion is intentionally inside this function, which
            // is called inside the timer for every source-owner sample.
            let package = base.to_workbook()?.into_package();
            native_edit(&package, input)
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

fn prepared_adds(sample_count: usize, position: usize) -> Result<Vec<PreparedAdd>> {
    (0..sample_count)
        .map(|_| {
            Ok(PreparedAdd {
                parent: Path::parse("/")?.into_owned(),
                node: NewNode::new("sheet")
                    .prop("name", INSERTED_NAME)
                    .prop("visibility", "visible")
                    .at(Placement::Position(Position::new(position))),
            })
        })
        .collect()
}

fn unified_once(file: &File, operation: PreparedAdd) -> Result<(File, Path<'static>)> {
    // The parent path and complete NewNode (including name, visible state and
    // insertion position) are prepared outside the timer.
    let added = file.add(operation.parent, operation.node)?;
    Ok((added.file, added.path))
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

fn marker(snapshot: &litchi_xlsb::cell_values::Snapshot) -> Result<Option<f64>> {
    let mut cells = snapshot.cells();
    let Some(cell) = cells.next() else {
        return Ok(None);
    };
    if cells.next().is_some() {
        return Err("worksheet has more than one stored cell".into());
    }
    match cell.value() {
        Value::Number(value) | Value::RkNumber(value) | Value::FormulaNumberCache(value) => {
            Ok(Some(*value))
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
                tab.state().is_visible(),
            )
        })
        .collect::<Vec<_>>();
    let names = tabs
        .iter()
        .map(|(name, _, _)| name.clone())
        .collect::<Vec<_>>();
    let visible = tabs.iter().map(|(_, _, visible)| *visible).collect();
    let mut markers = Vec::with_capacity(tabs.len());
    let mut binary_index = Vec::with_capacity(tabs.len());
    for (_, worksheet_position, _) in tabs {
        let worksheet_position =
            worksheet_position.ok_or("generated catalog has a non-worksheet tab")?;
        binary_index.push(
            workbook
                .worksheet_binary_index(worksheet_position)?
                .is_some(),
        );
        markers.push(marker(&workbook.cell_values(worksheet_position)?)?);
    }
    Ok(CatalogSnapshot {
        names,
        markers,
        visible,
        binary_index,
        active_catalog_position: workbook.active_catalog_position(),
        active_worksheet_index: workbook.active_worksheet_index(),
        active_sheet_index: litchi_core::sheet::WorkbookTrait::active_sheet_index(workbook),
    })
}

fn verify_original(bytes: &[u8], tab_count: usize) -> Result<()> {
    let observed = catalog_snapshot(bytes)?;
    let names = expected_names(tab_count);
    let markers = (0..tab_count)
        .map(|index| Some(sheet_marker(index)))
        .collect::<Vec<_>>();
    let visible = vec![true; tab_count];
    let binary_index = vec![true; tab_count];
    if observed.names != names
        || observed.markers != markers
        || observed.visible != visible
        || observed.binary_index != binary_index
    {
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
    let position = insertion_position(tab_count);
    let mut expected_order = expected_names(tab_count);
    expected_order.insert(position, INSERTED_NAME.to_owned());
    let mut expected_markers = (0..tab_count)
        .map(|index| Some(sheet_marker(index)))
        .collect::<Vec<_>>();
    expected_markers.insert(position, None);
    let expected_visible = vec![true; tab_count + 1];
    let expected_binary_index = vec![true; tab_count + 1];
    if observed.names != expected_order
        || observed.markers != expected_markers
        || observed.visible != expected_visible
        || observed.binary_index != expected_binary_index
    {
        return Err(format!("inserted catalog or markers differ: {observed:?}").into());
    }
    let expected_active = usize::from(position == 0);
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
    format!("xlsb_{}_native_sheet_insert_v1", mode.name())
}

fn unified_route(mode: OwnerMode) -> String {
    format!("xlsb_{}_unified_file_add_sheet_v1", mode.name())
}

fn expected_native_stats(stats: NativeStats) -> Result<()> {
    if stats.changed && stats.insert_calls == 1 && stats.commits == 1 {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, insert_calls={}, commits={}, expected true, 1, 1",
        stats.changed, stats.insert_calls, stats.commits
    )
    .into())
}

fn measure(bytes: &[u8], tab_count: usize, mode: OwnerMode, rounds: usize) -> Result<JsonValue> {
    let native = open_native(bytes, mode)?;
    let unified = open_unified(bytes, mode)?;
    let position = insertion_position(tab_count);
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows prepared sample count")?;

    // Both legs consume their prepared inputs, including an owned native name.
    // Parent parsing and caller-side node construction are outside the timer.
    let prepared_position = Position::new(position);
    let mut native_inputs = (0..sample_count)
        .map(|_| (prepared_position, INSERTED_NAME.to_owned()))
        .collect::<Vec<_>>();
    let mut unified_inputs = prepared_adds(sample_count, position)?;

    let native_before = native_before(&native)?;
    let unified_before = unified_bytes(&unified)?;
    verify_original(bytes, tab_count)?;
    verify_equivalence(bytes, &native_before)?;
    verify_equivalence(bytes, &unified_before)?;

    let warm_native_input = native_inputs
        .pop()
        .ok_or("native insertion preparation unexpectedly empty")?;
    let warm_unified_input = unified_inputs
        .pop()
        .ok_or("unified add preparation unexpectedly empty")?;
    let (native_warm, native_warm_stats) = native_once(&native, warm_native_input)?;
    expected_native_stats(native_warm_stats)?;
    let native_output = native_bytes(&native_warm)?;
    let (unified_warm, warm_path) = unified_once(&unified, warm_unified_input)?;
    let unified_output = unified_bytes(&unified_warm)?;
    let expected_path = format!("/sheet[{}]", position + 1);
    if warm_path.to_string() != expected_path {
        return Err(format!(
            "unified add returned canonical path {:?}; expected {expected_path}",
            warm_path
        )
        .into());
    }
    if native_output == bytes || unified_output == bytes {
        return Err("warm-up sheet insertion produced an unchanged archive".into());
    }
    verify_output(&native_output, tab_count)?;
    verify_output(&unified_output, tab_count)?;
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_insert_calls = None;
    let mut native_commits = None;
    for _ in 0..rounds {
        // ABBA: native, unified, unified, native.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let input = native_inputs
                    .pop()
                    .ok_or("prepared native insertion samples exhausted")?;
                let start = Instant::now();
                let (output, stats) = native_once(black_box(&native), input)?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_insert_calls.get_or_insert(stats.insert_calls);
                native_commits.get_or_insert(stats.commits);
                drop(output);
            } else {
                let input = unified_inputs
                    .pop()
                    .ok_or("prepared unified add samples exhausted")?;
                let start = Instant::now();
                let (output, path) = unified_once(black_box(&unified), input)?;
                let elapsed = start.elapsed().as_nanos();
                if path.to_string() != expected_path {
                    return Err(format!(
                        "unified add returned canonical path {:?}; expected {expected_path}",
                        path
                    )
                    .into());
                }
                black_box(&output);
                unified_ns.push(elapsed);
                drop(output);
            }
        }
    }
    if !native_inputs.is_empty() || !unified_inputs.is_empty() {
        return Err("prepared insertion samples have unused inputs".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    let active_after = usize::from(position == 0);
    Ok(json!({
        "tab_count": tab_count,
        "owner": mode.name(),
        "operation": "sheet_insert_blank",
        "parent": "/",
        "position": position,
        "inserted_name": INSERTED_NAME,
        "inserted_state": "visible",
        "canonical_path": expected_path,
        "native_route": native_route(mode),
        "unified_route": unified_route(mode),
        "unified_prepared_owner": mode.name(),
        "native_changed": native_warm_stats.changed,
        "native_insert_calls": native_insert_calls,
        "native_commits": native_commits,
        "unified_add_calls": 1,
        "active_identity_before": "Sheet00000",
        "active_catalog_position_before": 0,
        "active_identity_after": "Sheet00000",
        "active_catalog_position_after": active_after,
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
                "position_prepared_outside_timing": true,
                "owned_name_prepared_outside_timing": true,
                "position": position,
                "timed_input_clones": 0,
                "timed_call_boundary": "benchmark consumes checked Position scalar and owned name; native Edit::insert_sheet authors one blank worksheet"
            },
            "unified": {
                "parent_path_prepared_outside_timing": true,
                "parent": "/",
                "node_prepared_outside_timing": true,
                "per_sample_owned_node_preparations_outside_timing": sample_count,
                "timed_input_clones": 0,
                "timed_call_boundary": "benchmark consumes owned Path<'static> and NewNode<'static> through File::add"
            },
            "pairing": "one prepared insertion input is owned per sample on each leg; no caller-side path parsing or node clone occurs inside a timed call; facade property coercion is timed"
        }
    }))
}

fn parse_counts(value: &str) -> Result<Vec<usize>> {
    let counts = value
        .split(',')
        .map(|part| {
            let count = part.parse::<usize>()?;
            if count < 1 {
                return Err("counts must be at least one tab".into());
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
                    "usage: unified-xlsb-sheet-insert-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
    let source_manifest = std::env::var("LITCHI_SOURCE_MANIFEST")
        .ok()
        .and_then(|value| serde_json::from_str(&value).ok())
        .unwrap_or_else(|| json!("unspecified"));
    json!({
        "crate": env!("CARGO_PKG_NAME"),
        "crate_version": env!("CARGO_PKG_VERSION"),
        "rustc": std::env::var("LITCHI_RUSTC_VERSION").unwrap_or_else(|_| "unspecified".to_owned()),
        "target": std::env::var("LITCHI_TARGET").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_commit": std::env::var("LITCHI_SOURCE_COMMIT").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_digest": std::env::var("LITCHI_SOURCE_DIGEST").unwrap_or_else(|_| "unspecified".to_owned()),
        "source_manifest": source_manifest,
        "binary_sha256": std::env::var("LITCHI_BINARY_SHA256").unwrap_or_else(|_| "unspecified".to_owned()),
        "binary_manifest": {
            "path": std::env::var("LITCHI_BINARY_PATH").unwrap_or_else(|_| "unspecified".to_owned()),
            "bytes": std::env::var("LITCHI_BINARY_BYTES").unwrap_or_else(|_| "unspecified".to_owned()),
            "sha256": std::env::var("LITCHI_BINARY_SHA256").unwrap_or_else(|_| "unspecified".to_owned()),
        },
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
                    "operation": "sheet_insert_blank",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let corpus_sha256 = sha256(&archive);
        let artifact = json!({
            "generator": "litchi_xlsb::builder::{self, Builder}; N worksheets in catalog order; one distinct numeric marker per worksheet",
            "tab_count": tab_count,
            "insertion_position": insertion_position(tab_count),
            "inserted_name": INSERTED_NAME,
            "inserted_state": "visible",
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let mut result = match measure(&archive, tab_count, mode, config.rounds) {
                Ok(result) => result,
                Err(error) => {
                    eprintln!(
                        "failed {} sheet insertion: {tab_count} tabs: {error}",
                        mode.name()
                    );
                    failures.push(json!({
                        "tab_count": tab_count,
                        "owner": mode.name(),
                        "operation": "sheet_insert_blank",
                        "artifact": artifact,
                        "error": error.to_string(),
                    }));
                    continue;
                },
            };
            result["artifact"] = artifact.clone();
            eprintln!("measured {} sheet insertion: {tab_count} tabs", mode.name());
            results.push(result);
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "xlsb",
            "workload": "complete-sheet-catalog-blank-insertion-at-floor-half",
            "parent": "/",
            "inserted_name": INSERTED_NAME,
            "inserted_state": "visible",
            "position": "floor(N/2)",
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
            "timing_scope": "one native Package::edit + Edit::insert_sheet(Position, Inserted, Visible) + one commit + Commit::into_package call; source-owner native samples additionally include SourceBackedWorkbook::to_workbook().into_package() inside each timed edit; one unified File::add(prepared workbook parent, NewNode sheet with name/visible state/Position) call with source promotion inside the source facade leg; excludes corpus generation, opening, position/node preparation, warm-up, output serialization/write_to, equivalence checks and drop",
            "publication_contract": "native insert_calls is one and native_commits is one per sample; unified_add_calls is one per sample; each result verifies complete catalog order, one blank inserted worksheet, all distinct preserved markers, visible state and active Sheet00000 identity after insertion",
            "equivalence_scope": "warm-up outputs are compared by exact ZIP bytes, falling back to decoded ZIP member bytes with per-member length and SHA-256 reporting; differing member bytes fail the matrix",
            "build_provenance": build_provenance(),
            "results": results,
            "failure_matrix": failures,
            "performance_claim": "scoped paired single-sheet insertion timing evidence only; no batch-scaling claim and full acceptance remains pending",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("benchmark matrix contains failures; completed samples and failures are recorded in JSON".into())
    }
}
