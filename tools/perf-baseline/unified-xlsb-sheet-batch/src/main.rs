//! Paired XLSB sheet-visibility batch publication: one native package edit
//! versus one unified `File::edit` closure containing direct sheet-path sets.
//!
//! Corpus generation, opening, target/property preparation, warm-up, output
//! serialization and equivalence checks are outside the timed interval. The
//! source native leg deliberately materializes its prepared
//! `SourceBackedWorkbook` into a package inside every timed edit, matching the
//! source facade's promotion boundary. The unified leg uses one direct
//! `/sheet[N]` path command per target in one transaction; it does not use a
//! query fanout.

use litchi_core::OwnedSource;
use litchi_core::selector::Path;
use litchi_ops::{File, IntoTarget, OpenOptions, Props, SourceMode, TargetArg};
use litchi_xlsb::SourceBackedWorkbook;
use litchi_xlsb::sheet::State;
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
const SENTINEL_NAME: &str = "VisibleSentinel";
const SHEET_PATH_TEMPLATE: &str = "/sheet[2]..../sheet[N+1]";
const SENTINEL_POSITION: usize = 1;
const FIRST_TARGET_CATALOG_POSITION: usize = 1;
const FIRST_TARGET_PATH_POSITION: usize = 2;

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
    visibility_setters: usize,
    commits: usize,
}

#[derive(Clone, Copy)]
struct UnifiedStats {
    changed: bool,
    operations: usize,
    commits: usize,
    leaf_operations: usize,
}

struct SheetOperation {
    target: TargetArg<'static>,
    props: Props<'static>,
}

struct PreparedOperations {
    sheets: Vec<SheetOperation>,
}

struct Config {
    owners: Vec<OwnerMode>,
    counts: Vec<usize>,
    rounds: usize,
}

fn target_name(index: usize) -> String {
    format!("Target{index:05}")
}

fn corpus(target_count: usize) -> Result<Vec<u8>> {
    use litchi_xlsb::builder::{self, Builder};

    let mut writer = Builder::new();
    // Keep the visible sentinel first. It is the active tab in the generated
    // workbook, so every changed target is non-active and the direct batch can
    // exercise one homogeneous native publication without an active-view
    // relocation barrier.
    let mut sentinel = builder::Worksheet::new(SENTINEL_NAME);
    sentinel.set_cell(0, 0, -1.0);
    writer.add_worksheet(sentinel);
    for index in 0..target_count {
        let mut sheet = builder::Worksheet::new(target_name(index));
        // Keep each generated worksheet physically non-empty. This exercises
        // the same worksheet payload closure at every N while the operation
        // itself remains workbook-tab metadata only.
        sheet.set_cell(0, 0, index as f64);
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

/// Parse and type every direct target and build every property set before the
/// timer starts. `Path::into_owned` gives each target a self-contained
/// lifetime so the timed closure only dispatches prepared values.
fn prepared_operations(target_count: usize) -> Result<PreparedOperations> {
    let mut sheets = Vec::with_capacity(target_count);
    for index in 0..target_count {
        let target_text = format!("/sheet[{}]", index + FIRST_TARGET_PATH_POSITION);
        let target = Path::parse(&target_text)?.into_owned().into_target()?;
        let props = Props::new().with("visibility", "hidden");
        sheets.push(SheetOperation { target, props });
    }
    Ok(PreparedOperations { sheets })
}

fn native_once(
    owner: &NativeOwner,
    target_count: usize,
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    let NativeOwner::Eager(base) = owner else {
        return Err("eager native operation received a source owner".into());
    };
    let mut edit = base.edit()?;
    for sheet in FIRST_TARGET_CATALOG_POSITION..=target_count {
        let _ = edit.set_sheet_state(sheet, State::Hidden)?;
    }
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let package = commit.into_package();
    Ok((
        package,
        NativeStats {
            changed,
            visibility_setters: target_count,
            commits: 1,
        },
    ))
}

fn native_once_source(
    owner: &NativeOwner,
    target_count: usize,
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    let NativeOwner::Source(base) = owner else {
        return Err("source native operation received an eager owner".into());
    };
    // Source promotion intentionally belongs inside this function, which is
    // called inside the timer for every source-owner native sample.
    let package = base.to_workbook()?.into_package();
    let mut edit = package.edit()?;
    for sheet in FIRST_TARGET_CATALOG_POSITION..=target_count {
        let _ = edit.set_sheet_state(sheet, State::Hidden)?;
    }
    let commit = edit.commit()?;
    let changed = commit.diagnostics().changed();
    let package = commit.into_package();
    Ok((
        package,
        NativeStats {
            changed,
            visibility_setters: target_count,
            commits: 1,
        },
    ))
}

fn native_once_for_mode(
    owner: &NativeOwner,
    target_count: usize,
    mode: OwnerMode,
) -> Result<(litchi_xlsb::Package, NativeStats)> {
    match mode {
        OwnerMode::Eager => native_once(owner, target_count),
        OwnerMode::Source => native_once_source(owner, target_count),
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

fn unified_once(file: &File, operations: PreparedOperations) -> Result<(File, UnifiedStats)> {
    let commit = file.edit(|transaction| {
        for operation in operations.sheets {
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

fn tab_states(bytes: &[u8]) -> Result<Vec<(String, State)>> {
    let package = litchi_xlsb::Package::from_bytes(bytes)?;
    let workbook = package.workbook()?;
    Ok(workbook
        .sheets()
        .map(|tab| (tab.name().to_owned(), tab.state()))
        .collect())
}

fn verify_output(bytes: &[u8], target_count: usize) -> Result<()> {
    let states = tab_states(bytes)?;
    if states.len() != target_count + 1 {
        return Err(format!(
            "sheet-visibility output has {} tabs, expected {}",
            states.len(),
            target_count + 1
        )
        .into());
    }
    for (index, (name, state)) in states.iter().enumerate() {
        if index == SENTINEL_POSITION - 1 {
            if name != SENTINEL_NAME || *state != State::Visible {
                return Err("visible sentinel was not preserved".into());
            }
            continue;
        }
        let target_index = index - SENTINEL_POSITION;
        if *state != State::Hidden {
            return Err(
                format!("target tab at index {index} has unexpected state {state:?}").into(),
            );
        }
        if name != &target_name(target_index) {
            return Err(format!("target tab at index {index} has unexpected name").into());
        }
    }
    Ok(())
}

fn verify_original(bytes: &[u8], target_count: usize) -> Result<()> {
    let states = tab_states(bytes)?;
    if states.len() != target_count + 1 {
        return Err(format!(
            "corpus has {} tabs, expected {}",
            states.len(),
            target_count + 1
        )
        .into());
    }
    if states.iter().any(|(_, state)| *state != State::Visible) {
        return Err("generated corpus has a non-visible tab".into());
    }
    if states.first().map(|(name, _)| name.as_str()) != Some(SENTINEL_NAME) {
        return Err("generated corpus does not begin with visible sentinel".into());
    }
    Ok(())
}

fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}

fn native_route(mode: OwnerMode) -> String {
    format!("xlsb_{}_native_sheet_visibility_batch_v1", mode.name())
}

fn unified_route(mode: OwnerMode) -> String {
    format!(
        "xlsb_{}_unified_file_edit_direct_sheet_visibility_batch_v1",
        mode.name()
    )
}

fn expected_native_stats(stats: NativeStats, target_count: usize) -> Result<()> {
    if stats.changed && stats.visibility_setters == target_count && stats.commits == 1 {
        return Ok(());
    }
    Err(format!(
        "native accounting mismatch: changed={}, visibility_setters={}, commits={}, expected true, {target_count}, 1",
        stats.changed, stats.visibility_setters, stats.commits
    )
    .into())
}

fn expected_unified_stats(stats: UnifiedStats, target_count: usize) -> Result<()> {
    if stats.changed && stats.operations == target_count && stats.leaf_operations == target_count {
        return Ok(());
    }
    Err(format!(
        "unified accounting mismatch: changed={}, operations={}, commits={}, leaf_operations={}, expected true, {target_count}, positive, {target_count}",
        stats.changed, stats.operations, stats.commits, stats.leaf_operations
    )
    .into())
}

fn measure(bytes: &[u8], target_count: usize, mode: OwnerMode, rounds: usize) -> Result<Value> {
    // Owner opening is deliberately outside every timed sample. Source mode
    // remains immutable here; promotion occurs in each operation boundary.
    let native = open_native(bytes, mode)?;
    let unified = open_unified(bytes, mode)?;
    let sample_count = rounds
        .checked_mul(2)
        .and_then(|value| value.checked_add(1))
        .ok_or("round count overflows prepared sample count")?;

    // Parse and prepare N direct sheet paths (`/sheet[2]` through
    // `/sheet[N+1]`) and N visibility property sets for every unified sample
    // before any warm-up or timed operation.
    let mut unified_samples = (0..sample_count)
        .map(|_| prepared_operations(target_count))
        .collect::<Result<Vec<_>>>()?;

    let native_before = native_before(&native)?;
    let unified_before = unified_bytes(&unified)?;
    verify_original(bytes, target_count)?;
    verify_equivalence(bytes, &native_before)?;
    verify_equivalence(bytes, &unified_before)?;

    let warm_operations = unified_samples
        .pop()
        .ok_or("unified operation preparation unexpectedly empty")?;
    let (native_warm, native_warm_stats) = native_once_for_mode(&native, target_count, mode)?;
    expected_native_stats(native_warm_stats, target_count)?;
    let native_output = native_bytes(&native_warm)?;
    let (unified_warm, unified_warm_stats) = unified_once(&unified, warm_operations)?;
    expected_unified_stats(unified_warm_stats, target_count)?;
    let unified_output = unified_bytes(&unified_warm)?;
    if native_output == bytes || unified_output == bytes {
        return Err("warm-up sheet visibility edit produced an unchanged archive".into());
    }
    verify_output(&native_output, target_count)?;
    verify_output(&unified_output, target_count)?;
    let warm_equivalence = verify_equivalence(&native_output, &unified_output)?;
    drop((native_warm, unified_warm));

    let mut native_ns = Vec::with_capacity(rounds * 2);
    let mut unified_ns = Vec::with_capacity(rounds * 2);
    let mut native_visibility_setters = None;
    let mut native_commits = None;
    let mut unified_operations = None;
    let mut unified_commits = None;
    let mut unified_leaf_operations = None;
    for _ in 0..rounds {
        // ABBA: native, unified, unified, native.
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let (output, stats) = native_once_for_mode(black_box(&native), target_count, mode)?;
                let elapsed = start.elapsed().as_nanos();
                expected_native_stats(stats, target_count)?;
                black_box(&output);
                native_ns.push(elapsed);
                native_visibility_setters.get_or_insert(stats.visibility_setters);
                native_commits.get_or_insert(stats.commits);
                drop(output);
            } else {
                let operations = unified_samples
                    .pop()
                    .ok_or("prepared unified operation samples exhausted")?;
                let start = Instant::now();
                let (output, stats) = unified_once(black_box(&unified), operations)?;
                let elapsed = start.elapsed().as_nanos();
                expected_unified_stats(stats, target_count)?;
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
        return Err("prepared unified operation samples have unused inputs".into());
    }

    let native_median = median(&native_ns);
    let unified_median = median(&unified_ns);
    Ok(json!({
        "target_count": target_count,
        "tab_count": target_count + 1,
        "owner": mode.name(),
        "operation": "sheet_visibility_batch",
        "target_template": SHEET_PATH_TEMPLATE,
        "sentinel_position": SENTINEL_POSITION,
        "first_target_catalog_position": FIRST_TARGET_CATALOG_POSITION,
        "first_target_path_position": FIRST_TARGET_PATH_POSITION,
        "property": "visibility",
        "property_value": "hidden",
        "sentinel": SENTINEL_NAME,
        "native_route": native_route(mode),
        "unified_route": unified_route(mode),
        "unified_prepared_owner": mode.name(),
        "native_changed": native_warm_stats.changed,
        "native_visibility_setters": native_visibility_setters,
        "native_commits": native_commits,
        "unified_changed": unified_warm_stats.changed,
        "unified_operations": unified_operations,
        "unified_commits": unified_commits,
        "unified_leaf_operations": unified_leaf_operations,
        "unified_one_file_edit": true,
        "unified_direct_path_commands": target_count,
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
                    "usage: unified-xlsb-sheet-batch-baseline [--counts N[,N...]] [--rounds N] [--owners eager,source]"
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
    for target_count in config.counts.iter().copied() {
        let archive = match corpus(target_count) {
            Ok(archive) => archive,
            Err(error) => {
                eprintln!("failed corpus generation: {target_count} target sheets: {error}");
                failures.push(json!({
                    "target_count": target_count,
                    "phase": "corpus",
                    "operation": "sheet_visibility_batch",
                    "error": error.to_string(),
                }));
                continue;
            },
        };
        let corpus_sha256 = sha256(&archive);
        let artifact = json!({
            "generator": "litchi_xlsb::builder::{self, Builder}; N target worksheets plus one visible sentinel; one numeric marker per tab",
            "target_count": target_count,
            "tab_count": target_count + 1,
            "bytes": archive.len(),
            "sha256": corpus_sha256.clone(),
            "corpus_sha256": corpus_sha256,
        });
        for mode in config.owners.iter().copied() {
            let mut result = match measure(&archive, target_count, mode, config.rounds) {
                Ok(result) => result,
                Err(error) => {
                    eprintln!(
                        "failed {} direct sheet visibility batch: {target_count} target sheets: {error}",
                        mode.name()
                    );
                    failures.push(json!({
                        "target_count": target_count,
                        "owner": mode.name(),
                        "operation": "sheet_visibility_batch",
                        "artifact": artifact,
                        "error": error.to_string(),
                    }));
                    continue;
                },
            };
            result["artifact"] = artifact.clone();
            eprintln!(
                "measured {} direct sheet visibility batch: {target_count} target sheets",
                mode.name()
            );
            results.push(result);
        }
    }

    println!(
        "{}",
        serde_json::to_string_pretty(&json!({
            "schema": 3,
            "format": "xlsb",
            "workload": "sheet-visibility-direct-path-batch-with-first-visible-sentinel",
            "target_template": SHEET_PATH_TEMPLATE,
            "property": "visibility",
            "property_value": "hidden",
            "sentinel": SENTINEL_NAME,
            "target_counts": config.counts,
            "owners": config.owners.iter().map(|mode| mode.name()).collect::<Vec<_>>(),
            "rounds": config.rounds,
            "order": "ABBA",
            "samples_per_leg": config.rounds * 2,
            "routes": config
                .owners
                .iter()
                .flat_map(|mode| [native_route(*mode), unified_route(*mode)])
                .collect::<Vec<_>>(),
            "timing_scope": "one native Package::edit + N set_sheet_state(catalog positions 1..=N, State::Hidden) calls + one commit + Commit::into_package; source-owner native samples additionally include SourceBackedWorkbook::to_workbook().into_package() inside each timed edit; one unified File::edit closure with N prepared direct /sheet[2]..../sheet[N+1] path tx.set calls and diagnostics publication; source-owner facade promotion remains inside the timed edit; excludes corpus generation, opening, target/property preparation, warm-up, output serialization/write_to, equivalence checks and drop",
            "publication_contract": "the visible sentinel is catalog position 0 and remains active; native visibility_setters equals target count and native_commits is one per sample; unified_operations and unified_leaf_operations equal target count, unified_commits is recorded per sample, and every unified sample uses one File::edit closure",
            "equivalence_scope": "warm-up outputs are compared by exact ZIP bytes, falling back to decoded ZIP member bytes with per-member length and SHA-256 reporting",
            "build_provenance": build_provenance(),
            "results": results,
            "failure_matrix": failures,
            "performance_claim": "scoped direct-path batch timing evidence only; full acceptance remains pending",
        }))?
    );
    if failures.is_empty() {
        Ok(())
    } else {
        Err("benchmark matrix contains failures; completed samples and failures are recorded in JSON".into())
    }
}
