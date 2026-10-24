//! Paired XLSB cell-value publication: native package-owned edits versus the
//! unified `File::edit`/`Edit::set` route.
//!
//! Corpus generation, opening, target/property preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. Facade
//! property recollection, admission and coercion remain inside the clock. The
//! eager native leg starts from Package; the eager facade starts from Workbook.
//! This is a package-native comparison, not identical prepared owner types. The
//! native source leg deliberately materializes its prepared
//! `SourceBackedWorkbook` into a package for every edit, matching the source
//! facade's edit boundary. Each timed operation stages all selected cells in
//! one native package edit or one unified transaction; the result records how
//! many unified leaf commits were retained so a non-coalesced route cannot be
//! mistaken for the one-edit comparison.

use litchi_core::OwnedSource;
use litchi_core::sheet::WorkbookTrait;
use litchi_ops::{File, IntoTarget, OpenOptions, Props, SourceMode, TargetArg};
use litchi_xlsb::SourceBackedWorkbook;
use litchi_xlsb::cell_values::{Reference, Scalar};
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
const DEFAULT_ROUNDS: usize = 3;
const TARGET_PREFIX: &str = "/sheet[1]/A1";
const ORIGINAL_NAME: &str = "Sheet1";
const VALUE_OFFSET: f64 = 10_000.25;

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

struct UnifiedStats {
    changed: bool,
    operations: usize,
    commits: usize,
    leaf_operations: usize,
}

struct CellOp {
    reference: Reference,
    value: f64,
    target: TargetArg<'static>,
    props: Props<'static>,
}

struct PreparedOps {
    cells: Vec<CellOp>,
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

fn column_name(mut column: u32) -> String {
    let mut name = String::new();
    loop {
        let digit = u8::try_from(column % 26).expect("column remainder fits");
        name.push(char::from(b'A' + digit));
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    name.chars().rev().collect()
}

fn prepared_ops(cell_count: usize) -> Result<PreparedOps> {
    let mut cells = Vec::with_capacity(cell_count);
    for index in 0..cell_count {
        let row = u32::try_from(index / 32)?;
        let column = u32::try_from(index % 32)?;
        let reference = Reference::new(row, column)?;
        let value = index as f64 + VALUE_OFFSET;
        let coordinate = format!("{}{}", column_name(column), row + 1);
        let target_text = format!("/sheet[1]/{coordinate}");
        let target = litchi_core::selector::Path::parse(&target_text)?
            .into_owned()
            .into_target()?;
        let props = Props::new().with("value", value);
        cells.push(CellOp {
            reference,
            value,
            target,
            props,
        });
    }
    Ok(PreparedOps { cells })
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
    operations: &PreparedOps,
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
    for operation in &operations.cells {
        edit.set_scalar(operation.reference, Scalar::Number(operation.value))?;
    }
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let changes = commit.patch().changes().len();
    let package = commit.into_package();
    Ok((
        package,
        NativeStats {
            changed,
            operations: operations.cells.len(),
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

fn native_bytes(package: &litchi_xlsb::Package) -> Result<Vec<u8>> {
    Ok(package.to_bytes()?)
}

fn unified_once(
    file: &File,
    operations: PreparedOps,
    shape: OperationShape,
) -> Result<(File, UnifiedStats)> {
    match shape {
        OperationShape::Set => {
            let mut cells = operations.cells.into_iter();
            let operation = cells
                .next()
                .ok_or("one-shot set unexpectedly has no cell")?;
            if cells.next().is_some() {
                return Err("one-shot set unexpectedly has multiple cells".into());
            }
            let output = file.set(operation.target, operation.props)?;
            Ok((
                output,
                UnifiedStats {
                    changed: true,
                    operations: 1,
                    commits: 1,
                    leaf_operations: 1,
                },
            ))
        },
        OperationShape::Edit => {
            let commit = file.edit(|transaction| {
                for operation in operations.cells {
                    transaction.set(operation.target, operation.props)?;
                }
                Ok(())
            })?;
            let diagnostics = *commit.diagnostics();
            Ok((
                commit.into_file(),
                UnifiedStats {
                    changed: diagnostics.changed(),
                    operations: diagnostics.operations(),
                    commits: diagnostics.commits(),
                    leaf_operations: diagnostics.leaf_operations(),
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

fn verify_output(bytes: &[u8], operations: &PreparedOps) -> Result<()> {
    let workbook = litchi_xlsb::Workbook::new(Cursor::new(bytes.to_vec()))?;
    let names = WorkbookTrait::worksheet_names(&workbook);
    if names.len() != 1 || names[0] != ORIGINAL_NAME {
        return Err(format!("cell-value output has worksheet names {names:?}").into());
    }
    let snapshot = workbook.cell_values(0)?;
    let cells = snapshot.cells().count();
    if cells != operations.cells.len() {
        return Err(format!(
            "cell-value output has {cells} cells, expected {}",
            operations.cells.len()
        )
        .into());
    }
    for operation in &operations.cells {
        let cell = snapshot
            .cell(operation.reference)?
            .ok_or("cell-value output lost an edited cell")?;
        let actual = match cell.value() {
            litchi_xlsb::cell_values::Value::Number(value)
            | litchi_xlsb::cell_values::Value::RkNumber(value) => *value,
            value => {
                return Err(format!(
                    "cell {:?} has unexpected value family {value:?}",
                    operation.reference
                )
                .into());
            },
        };
        if actual.to_bits() != operation.value.to_bits() {
            return Err(format!(
                "cell {:?} has value {actual}, expected {}",
                operation.reference, operation.value
            )
            .into());
        }
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

fn native_route(mode: OwnerMode, shape: OperationShape) -> String {
    format!(
        "xlsb_{}_native_primary_cell_value_{}_v1",
        mode.name(),
        shape.name()
    )
}

fn unified_route(mode: OwnerMode, shape: OperationShape) -> String {
    format!(
        "xlsb_{}_unified_file_{}_cell_value_v1",
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

#[derive(Clone, Copy, PartialEq, Eq)]
enum OperationShape {
    Set,
    Edit,
}

impl OperationShape {
    const fn name(self) -> &'static str {
        match self {
            Self::Set => "set",
            Self::Edit => "edit",
        }
    }

    const fn staged_cells(self, corpus_cells: usize) -> usize {
        match self {
            Self::Set => 1,
            Self::Edit => corpus_cells,
        }
    }
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
    let staged_cells = shape.staged_cells(cell_count);
    if staged_cells > cell_count {
        return Err("the operation shape stages more cells than the corpus".into());
    }

    // Parse and prepare all target/property inputs before timing. Each
    // prepared sample owns its short target/property payload, so the timed
    // unified leg measures dispatch, admission, coalescing and publication.
    let native_ops = prepared_ops(staged_cells)?;
    let mut unified_samples = Vec::with_capacity(sample_count);
    for _ in 0..sample_count {
        unified_samples.push(prepared_ops(staged_cells)?);
    }

    let native_before = native_before(&native)?;
    let unified_before = unified_bytes(&unified)?;
    verify_original(bytes, cell_count)?;
    verify_equivalence(bytes, &native_before)?;
    verify_equivalence(bytes, &unified_before)?;

    let warm_operations = unified_samples
        .pop()
        .ok_or("unified operation preparation unexpectedly empty")?;
    let (native_warm, native_warm_stats) = native_once(&native, &native_ops)?;
    expected_native_stats(native_warm_stats, staged_cells)?;
    let native_output = native_bytes(&native_warm)?;
    let (unified_warm, unified_warm_stats) = unified_once(&unified, warm_operations, shape)?;
    let unified_output = unified_bytes(&unified_warm)?;
    if native_output == bytes || unified_output == bytes {
        return Err("warm-up cell-value edit produced an unchanged archive".into());
    }
    verify_output(&native_output, &native_ops)?;
    verify_output(&unified_output, &native_ops)?;
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
                let (output, stats) = native_once(&native, &native_ops)?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats, staged_cells)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_operations.get_or_insert(stats.operations);
                native_commits.get_or_insert(stats.commits);
                native_changes.get_or_insert(stats.changes);
                drop(output);
            } else {
                let operations = unified_samples
                    .pop()
                    .ok_or("unified operation preparation exhausted")?;
                let start = Instant::now();
                let (output, stats) = unified_once(black_box(&unified), operations, shape)?;
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
    if !unified_samples.is_empty() {
        return Err("prepared unified inputs have unused samples".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "cell_count": cell_count,
        "owner": mode.name(),
        "operation": "cell_value",
        "operation_shape": shape.name(),
        "target_template": TARGET_PREFIX,
        "staged_cells": staged_cells,
        "value": "index + 10000.25",
        "native_route": native_route(mode, shape),
        "unified_route": unified_route(mode, shape),
        "unified_prepared_owner": mode.name(),
        "native_changed": native_warm_stats.changed,
        "native_operations": native_operations,
        "native_commits": native_commits,
        "native_changes": native_changes,
        "unified_changed": unified_warm_stats.changed,
        "unified_operations": unified_operations,
        "unified_commits": unified_commits,
        "unified_leaf_operations": unified_leaf_operations,
        "unified_coalesced": (shape == OperationShape::Edit)
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
                    "usage: unified-xlsb-cell-value-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
            "generator": "litchi_xlsb::builder::{self, Builder}; one worksheet; row-major numeric cells",
            "cell_count": cell_count,
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let shapes = if cell_count == 1 {
                vec![OperationShape::Set, OperationShape::Edit]
            } else {
                vec![OperationShape::Edit]
            };
            for shape in shapes {
                let mut result = match measure(&archive, cell_count, mode, shape, config.rounds) {
                    Ok(result) => result,
                    Err(error) => {
                        eprintln!(
                            "failed {} {} cell value: {cell_count} cells: {error}",
                            mode.name(),
                            shape.name()
                        );
                        failures.push(json!({
                            "cell_count": cell_count,
                            "owner": mode.name(),
                            "operation": "cell_value",
                            "operation_shape": shape.name(),
                            "artifact": artifact,
                            "error": error.to_string(),
                        }));
                        continue;
                    },
                };
                result["artifact"] = artifact.clone();
                eprintln!(
                    "measured {} {} cell value: {cell_count} cells",
                    mode.name(),
                    shape.name()
                );
                results.push(result);
            }
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 3,
            "format": "xlsb",
            "workload": "one-sheet-cell-value-set",
            "target_template": TARGET_PREFIX,
            "property": "value",
            "cell_counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": [
                "xlsb_eager_native_primary_cell_value_set_v1",
                "xlsb_eager_unified_file_set_cell_value_v1",
                "xlsb_eager_native_primary_cell_value_edit_v1",
                "xlsb_eager_unified_file_edit_cell_value_v1",
                "xlsb_source_native_primary_cell_value_set_v1",
                "xlsb_source_unified_file_set_cell_value_v1",
                "xlsb_source_native_primary_cell_value_edit_v1",
                "xlsb_source_unified_file_edit_cell_value_v1",
            ],
            "timing_scope": "one native Package::edit_cell_values + N Edit::set_scalar(Reference, Scalar::Number) + commit + Commit::into_package call; source-owner native samples additionally include SourceBackedWorkbook::to_workbook().into_package() per edit; one-shot unified samples time File::set('/sheet[1]/A1', {value: number}); batch unified samples time one File::edit closure with N tx.set('/sheet[1]/A1', {value: number}) calls and report retained leaf commits; includes facade property recollection, admission and coercion; eager native starts from Package while eager unified starts from Workbook; excludes corpus generation, opening, caller target/property preparation, warm-up, output serialization/write_to, equivalence checks and drop",
            "publication_contract": "native commits stage all selected cells in one package-owned edit; one-shot unified routes use one File::set; batch unified routes use one File::edit and expose diagnostics.commits to make coalescing observable",
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
