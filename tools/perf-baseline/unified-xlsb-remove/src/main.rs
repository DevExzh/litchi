//! Paired XLSB cell-removal publication: native package-owned edits versus
//! the unified `File::remove` and transaction routes.
//!
//! Corpus generation, opening, target preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. The
//! eager native leg starts from `Package`; the eager facade starts from an
//! eager workbook. The source native leg deliberately materializes its
//! prepared `SourceBackedWorkbook` into a package for every removal, matching
//! the source facade's edit boundary. Range cases remove one prepared range
//! atomically. Command cases exercise prebuilt direct-cell `BatchItem` values;
//! one-cell case also measures the one-shot `File::remove` route.

use litchi_core::OwnedSource;
use litchi_core::selector::Path;
use litchi_core::sheet::WorkbookTrait;
use litchi_ops::{BatchItem, BatchMode, File, IntoTarget, OpenOptions, SourceMode, TargetArg};
use litchi_xlsb::SourceBackedWorkbook;
use litchi_xlsb::cell_values::Reference;
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

const DEFAULT_COUNTS: [usize; 5] = [1, 10, 100, 1_000, 10_000];
const DEFAULT_ROUNDS: usize = 9;
const ORIGINAL_NAME: &str = "Sheet1";
const TARGET_TEMPLATE: &str = "/sheet[1]/A1:A{N}";
const SENTINEL_VALUE: f64 = -9_999.75;

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
    operations: usize,
    commits: usize,
    changes: usize,
}

#[derive(Clone, Copy)]
struct UnifiedStats {
    changed: bool,
    operations: usize,
    commits: usize,
    leaf_operations: Option<usize>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum OperationShape {
    OneShot,
    Range,
    Transaction,
    CommandBatch,
}

impl OperationShape {
    const fn name(self) -> &'static str {
        match self {
            Self::OneShot => "one_shot",
            Self::Range => "range",
            Self::Transaction => "transaction",
            Self::CommandBatch => "command_batch",
        }
    }
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
        let row = u32::try_from(index)?;
        sheet.set_cell(row, 0, index as f64 + 0.5);
    }
    // Keep one untouched cell outside the removal range so each successful
    // output proves that the selected range was removed and the worksheet
    // remained usable.
    sheet.set_cell(0, 1, SENTINEL_VALUE);
    writer.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    writer.build()?.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn target_text(cell_count: usize) -> String {
    format!("/sheet[1]/A1:A{cell_count}")
}

enum PreparedOperation {
    Targets(Vec<TargetArg<'static>>),
    Batch(Vec<BatchItem<'static>>),
}

fn prepared_operation(cell_count: usize, shape: OperationShape) -> Result<PreparedOperation> {
    match shape {
        OperationShape::OneShot | OperationShape::Transaction | OperationShape::Range => {
            let text = if matches!(shape, OperationShape::OneShot | OperationShape::Transaction) {
                "/sheet[1]/A1".to_owned()
            } else {
                target_text(cell_count)
            };
            let path = Path::parse(&text)?;
            Ok(PreparedOperation::Targets(vec![
                path.into_owned().into_target()?,
            ]))
        },
        OperationShape::CommandBatch => {
            let mut items = Vec::with_capacity(cell_count);
            for row in 0..cell_count {
                let mut item = BatchItem::new("remove");
                item.path = Some(std::borrow::Cow::Owned(format!("/sheet[1]/A{}", row + 1)));
                items.push(item);
            }
            Ok(PreparedOperation::Batch(items))
        },
    }
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

fn native_once(
    owner: &NativeOwner,
    cell_count: usize,
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    let mut edit = match owner {
        NativeOwner::Eager(base) => base
            .edit_cell_values(0usize)?
            .resolve()?
            .ok_or("worksheet selector unexpectedly did not resolve")?,
        NativeOwner::Source(base) => {
            let package = base.to_workbook()?.into_package();
            package
                .edit_cell_values(0usize)?
                .resolve()?
                .ok_or("worksheet selector unexpectedly did not resolve")?
        },
    };
    let mut removed = 0;
    for row in 0..cell_count {
        let reference = Reference::new(u32::try_from(row)?, 0)?;
        if edit.remove(reference)? {
            removed += 1;
        }
    }
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let changes = commit.patch().changes().len();
    let package = commit.into_package();
    if removed != cell_count {
        return Err(format!("native removal staged {removed} cells, expected {cell_count}").into());
    }
    Ok((
        package,
        NativeStats {
            changed,
            operations: cell_count,
            commits: 1,
            changes,
        },
    ))
}

fn native_before(owner: &NativeOwner) -> Result<Vec<u8>> {
    match owner {
        NativeOwner::Eager(package) => Ok(package.to_bytes()?),
        NativeOwner::Source(source) => Ok(source.to_workbook()?.into_package().to_bytes()?),
    }
}

fn unified_once(
    file: &File,
    operation: PreparedOperation,
    shape: OperationShape,
) -> Result<(File, UnifiedStats)> {
    match shape {
        OperationShape::OneShot | OperationShape::Range => {
            let PreparedOperation::Targets(targets) = operation else {
                return Err("one-shot removal received batch commands".into());
            };
            let mut targets = targets.into_iter();
            let target = targets
                .next()
                .ok_or("one-shot removal unexpectedly has no target")?;
            if targets.next().is_some() {
                return Err("one-shot removal unexpectedly has multiple targets".into());
            }
            Ok((
                file.remove(target)?,
                UnifiedStats {
                    changed: true,
                    operations: 1,
                    commits: 1,
                    leaf_operations: None,
                },
            ))
        },
        OperationShape::Transaction => {
            let PreparedOperation::Targets(targets) = operation else {
                return Err("transaction removal received batch commands".into());
            };
            let mut targets = targets.into_iter();
            let target = targets
                .next()
                .ok_or("transaction removal unexpectedly has no target")?;
            if targets.next().is_some() {
                return Err("transaction removal unexpectedly has multiple targets".into());
            }
            let commit = file.edit(|transaction| {
                transaction.remove(target)?;
                Ok(())
            })?;
            let diagnostics = *commit.diagnostics();
            Ok((
                commit.into_file(),
                UnifiedStats {
                    changed: diagnostics.changed(),
                    operations: diagnostics.operations(),
                    commits: diagnostics.commits(),
                    leaf_operations: Some(diagnostics.leaf_operations()),
                },
            ))
        },
        OperationShape::CommandBatch => {
            let PreparedOperation::Batch(items) = operation else {
                return Err("command batch received direct targets".into());
            };
            let commit = file.batch(items, BatchMode::Atomic)?;
            let diagnostics = *commit.diagnostics();
            if !diagnostics.changed() || diagnostics.commits() != 1 {
                return Err(format!(
                    "command batch did not coalesce one changed native commit: changed={}, operations={}, commits={}, leaf_operations={}",
                    diagnostics.changed(),
                    diagnostics.operations(),
                    diagnostics.commits(),
                    diagnostics.leaf_operations(),
                )
                .into());
            }
            Ok((
                commit.into_file(),
                UnifiedStats {
                    changed: diagnostics.changed(),
                    operations: diagnostics.operations(),
                    commits: diagnostics.commits(),
                    leaf_operations: Some(diagnostics.leaf_operations()),
                },
            ))
        },
    }
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

fn cell_number(value: &litchi_xlsb::cell_values::StoredCell) -> Result<f64> {
    match value.value() {
        litchi_xlsb::cell_values::Value::Number(value)
        | litchi_xlsb::cell_values::Value::RkNumber(value) => Ok(*value),
        value => Err(format!("cell has unexpected value family {value:?}").into()),
    }
}

fn verify_original(bytes: &[u8], expected_removed: usize) -> Result<()> {
    let workbook = litchi_xlsb::Workbook::new(Cursor::new(bytes.to_vec()))?;
    let names = WorkbookTrait::worksheet_names(&workbook);
    if names.len() != 1 || names[0] != ORIGINAL_NAME {
        return Err(format!("corpus has worksheet names {names:?}").into());
    }
    let snapshot = workbook.cell_values(0)?;
    let cells = snapshot.cells().count();
    if cells != expected_removed + 1 {
        return Err(format!(
            "corpus has {cells} cells, expected {}",
            expected_removed + 1
        )
        .into());
    }
    let sentinel = snapshot
        .cell(Reference::new(0, 1)?)?
        .ok_or("corpus lost sentinel cell")?;
    let actual = cell_number(sentinel)?;
    if actual.to_bits() != SENTINEL_VALUE.to_bits() {
        return Err(format!("sentinel has value {actual}, expected {SENTINEL_VALUE}").into());
    }
    Ok(())
}

fn verify_output(bytes: &[u8], removed: usize) -> Result<()> {
    let workbook = litchi_xlsb::Workbook::new(Cursor::new(bytes.to_vec()))?;
    let names = WorkbookTrait::worksheet_names(&workbook);
    if names.len() != 1 || names[0] != ORIGINAL_NAME {
        return Err(format!("removal output has worksheet names {names:?}").into());
    }
    let snapshot = workbook.cell_values(0)?;
    let cells = snapshot.cells().count();
    if cells != 1 {
        return Err(format!("removal output has {cells} cells, expected 1").into());
    }
    for row in 0..removed {
        let reference = Reference::new(u32::try_from(row)?, 0)?;
        if snapshot.cell(reference)?.is_some() {
            return Err(format!("removal output retained cell {reference:?}").into());
        }
    }
    let sentinel = snapshot
        .cell(Reference::new(0, 1)?)?
        .ok_or("removal output lost sentinel cell")?;
    let actual = cell_number(sentinel)?;
    if actual.to_bits() != SENTINEL_VALUE.to_bits() {
        return Err(format!("sentinel has value {actual}, expected {SENTINEL_VALUE}").into());
    }
    Ok(())
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn native_route(mode: OwnerMode, shape: OperationShape) -> String {
    format!(
        "xlsb_{}_native_primary_cell_removal_{}_v1",
        mode.name(),
        shape.name()
    )
}

fn unified_route(mode: OwnerMode, shape: OperationShape) -> String {
    format!(
        "xlsb_{}_unified_file_remove_cell_{}_v1",
        mode.name(),
        shape.name()
    )
}

fn expected_native_stats(stats: NativeStats, expected_operations: usize) -> Result<()> {
    if stats.changed
        && stats.operations == expected_operations
        && stats.commits == 1
        && stats.changes == expected_operations
    {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, operations={}, commits={}, changes={}, expected true, {expected_operations}, 1, {expected_operations}",
        stats.changed, stats.operations, stats.commits, stats.changes
    )
    .into())
}

fn measure(
    bytes: &[u8],
    cell_count: usize,
    mode: OwnerMode,
    shape: OperationShape,
    rounds: usize,
) -> Result<Value> {
    let native = open_native(bytes, mode)?;
    let unified = open_unified(bytes, mode)?;
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows prepared sample count")?;

    // Parse and own every bounded target before timing. The target parser and
    // selector allocation therefore do not distort the publication clocks.
    let mut prepared = Vec::with_capacity(sample_count);
    for _ in 0..sample_count {
        prepared.push(prepared_operation(cell_count, shape)?);
    }

    let native_before = native_before(&native)?;
    let unified_before = unified_bytes(&unified)?;
    verify_original(bytes, cell_count)?;
    verify_equivalence(bytes, &native_before)?;
    verify_equivalence(bytes, &unified_before)?;

    let warm_operation = prepared
        .pop()
        .ok_or("unified operation preparation unexpectedly empty")?;
    let (native_warm, native_warm_stats) = native_once(&native, cell_count)?;
    expected_native_stats(native_warm_stats, cell_count)?;
    let native_output = native_warm.to_bytes()?;
    let (unified_warm, unified_warm_stats) = unified_once(&unified, warm_operation, shape)?;
    let unified_output = unified_bytes(&unified_warm)?;
    if native_output == bytes || unified_output == bytes {
        return Err("warm-up cell removal produced an unchanged archive".into());
    }
    verify_output(&native_output, cell_count)?;
    verify_output(&unified_output, cell_count)?;
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_operations = None;
    let mut native_commits = None;
    let mut native_changes = None;
    let mut unified_operations = None;
    let mut unified_commits = None;
    let mut unified_leaf_operations = None;
    for _ in 0..rounds {
        // ABBA: native, unified, unified, native.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let (output, stats) = native_once(black_box(&native), cell_count)?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats, cell_count)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_operations.get_or_insert(stats.operations);
                native_commits.get_or_insert(stats.commits);
                native_changes.get_or_insert(stats.changes);
                drop(output);
            } else {
                let operation = prepared
                    .pop()
                    .ok_or("unified operation preparation exhausted")?;
                let start = Instant::now();
                let (output, stats) = unified_once(black_box(&unified), operation, shape)?;
                let elapsed = start.elapsed().as_nanos();
                black_box(&output);
                unified_ns.push(elapsed);
                unified_operations.get_or_insert(stats.operations);
                unified_commits.get_or_insert(stats.commits);
                unified_leaf_operations.get_or_insert(stats.leaf_operations);
                drop(output);
            }
        }
    }
    if !prepared.is_empty() {
        return Err("prepared unified operations have unused samples".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "cell_count": cell_count,
        "owner": mode.name(),
        "operation": "cell_removal",
        "operation_shape": shape.name(),
        "target": if shape == OperationShape::OneShot || shape == OperationShape::Transaction {
            "/sheet[1]/A1".to_owned()
        } else if shape == OperationShape::Range {
            target_text(cell_count)
        } else {
            "/sheet[1]/A1, /sheet[1]/A2, ... direct-cell targets".to_owned()
        },
        "removed_cells": cell_count,
        "native_route": native_route(mode, shape),
        "unified_route": unified_route(mode, shape),
        "native_changed": native_warm_stats.changed,
        "native_operations": native_operations,
        "native_commits": native_commits,
        "native_changes": native_changes,
        "unified_changed": unified_warm_stats.changed,
        "unified_operations": unified_operations,
        "unified_commits": unified_commits,
        "unified_leaf_operations": unified_leaf_operations,
        "unified_remove_calls": if shape == OperationShape::CommandBatch { cell_count } else { 1 },
        "unified_coalesced": (shape == OperationShape::CommandBatch || shape == OperationShape::Transaction)
            .then(|| unified_commits == Some(1)),
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
                    "usage: unified-xlsb-remove-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
            "generator": "litchi_xlsb::builder::{self, Builder}; one worksheet; column-A numeric cells plus B1 sentinel",
            "cell_count": cell_count,
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let shapes = if cell_count == 1 {
                vec![
                    OperationShape::OneShot,
                    OperationShape::Range,
                    OperationShape::Transaction,
                    OperationShape::CommandBatch,
                ]
            } else {
                vec![OperationShape::Range, OperationShape::CommandBatch]
            };
            for shape in shapes {
                let mut result = match measure(&archive, cell_count, mode, shape, config.rounds) {
                    Ok(result) => result,
                    Err(error) => {
                        eprintln!(
                            "failed {} {} cell removal: {cell_count} cells: {error}",
                            mode.name(),
                            shape.name()
                        );
                        failures.push(json!({
                            "cell_count": cell_count,
                            "owner": mode.name(),
                            "operation": "cell_removal",
                            "operation_shape": shape.name(),
                            "artifact": artifact,
                            "error": error.to_string(),
                        }));
                        continue;
                    },
                };
                result["artifact"] = artifact.clone();
                eprintln!(
                    "measured {} {} cell removal: {cell_count} cells",
                    mode.name(),
                    shape.name()
                );
                results.push(result);
            }
        }
    }

    let routes = [
        (OwnerMode::Eager, OperationShape::OneShot),
        (OwnerMode::Eager, OperationShape::Range),
        (OwnerMode::Eager, OperationShape::Transaction),
        (OwnerMode::Eager, OperationShape::CommandBatch),
        (OwnerMode::Source, OperationShape::OneShot),
        (OwnerMode::Source, OperationShape::Range),
        (OwnerMode::Source, OperationShape::Transaction),
        (OwnerMode::Source, OperationShape::CommandBatch),
    ]
    .into_iter()
    .flat_map(|(mode, shape)| [native_route(mode, shape), unified_route(mode, shape)])
    .collect::<Vec<_>>();
    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 1,
            "format": "xlsb",
            "workload": "one-sheet-cell-removal",
            "selector_template": TARGET_TEMPLATE,
            "cell_counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": routes,
            "timing_scope": "one native Package::edit_cell_values(worksheet 0) + N Edit::remove(Reference) calls + one commit + Commit::into_package; source-owner native samples additionally include SourceBackedWorkbook::to_workbook().into_package() per edit; one unified File::remove range call for range samples; one-cell transaction samples time one File::edit closure with one tx.remove('/sheet[1]/A1'); command samples time one File::batch with N prebuilt direct-cell BatchItem remove commands and report retained leaf commits; one-shot samples time File::remove('/sheet[1]/A1'); includes source checks, range/direct-cell resolution, facade admission, native removal and publication; eager native starts from Package while eager unified starts from Workbook; excludes corpus generation, opening, caller target and BatchItem preparation, warm-up, output serialization/write_to, equivalence checks and drop",
            "publication_contract": "native samples remove every selected cell in one package-owned edit and commit; range unified samples use one File::remove range call; the one-cell transaction uses one File::edit and exposes diagnostics.commits; command samples use one File::batch and expose diagnostics.commits and diagnostics.leaf_operations; one-shot samples use one File::remove direct-cell call",
            "equivalence_scope": "warm-up outputs are compared by exact ZIP bytes, falling back to decoded ZIP member bytes with per-member length and SHA-256 reporting",
            "failure_matrix_scope": ["corpus", "open", "warm_up", "timed_matrix"],
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
