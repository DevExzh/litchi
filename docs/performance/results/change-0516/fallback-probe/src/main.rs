//! Standalone public-API fallback probe for XLSX semantic commit work.
//!
//! The probe deliberately lives outside the performance harness. It creates
//! one deterministic two-sheet numeric workbook with public worksheet default
//! descent metadata for each selected shape, then measures only `Edit::commit`
//! after opening a fresh workbook, warming its Store, and preparing the
//! requested changed edit. Serialization and all semantic/source oracles stay
//! outside the measured clock.

extern crate self as litchi_perf_baseline;

#[allow(
    dead_code,
    reason = "The shared observer includes harness-only identity helpers."
)]
#[path = "../../../../../../tools/perf-baseline/src/allocation_metrics.rs"]
pub mod allocation_metrics;
#[cfg(feature = "allocator-metrics")]
#[path = "../../../../../../tools/perf-baseline/src/bin/support/counting_allocator.rs"]
mod counting_allocator;

use std::{
    collections::HashSet,
    error::Error,
    fs::File,
    io::{self, Write},
    path::PathBuf,
    time::Instant,
};

use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_xlsx::{Cell, Rect, Value, Workbook};
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA_VERSION: u32 = 1;
const SHEET_COUNT: usize = 2;
const DEFAULT_SAMPLES: usize = 100;
const DEFAULT_WARMUPS: usize = 3;
const CORPUS_VARIANT: &str = "numeric-two-sheet-public-default-descent-x14ac-v1";
const DEFAULT_DESCENT: f64 = 0.2;
const DEFAULT_HEIGHT: f64 = 15.0;
const X14AC_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

type ProbeResult<T> = Result<T, Box<dyn Error>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Shape {
    Tiny,
    Medium,
    DenseWide,
}

impl Shape {
    const ALL: [Self; 3] = [Self::Tiny, Self::Medium, Self::DenseWide];

    const fn name(self) -> &'static str {
        match self {
            Self::Tiny => "tiny",
            Self::Medium => "medium",
            Self::DenseWide => "dense-wide",
        }
    }

    const fn side(self) -> usize {
        match self {
            Self::Tiny => 8,
            Self::Medium => 32,
            Self::DenseWide => 256,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Scenario {
    WarmChangedOneCell,
    WarmChangedOnePercent,
}

impl Scenario {
    const ALL: [Self; 2] = [Self::WarmChangedOneCell, Self::WarmChangedOnePercent];

    const fn name(self) -> &'static str {
        match self {
            Self::WarmChangedOneCell => "warm-changed-one-cell",
            Self::WarmChangedOnePercent => "warm-changed-one-percent",
        }
    }

    const fn warm_store(self) -> bool {
        true
    }

    const fn one_cell(self) -> bool {
        matches!(self, Self::WarmChangedOneCell)
    }
}

#[derive(Clone, Copy, Debug, Hash, PartialEq, Eq)]
struct Coordinate {
    sheet: usize,
    row: usize,
    column: usize,
}

struct Corpus {
    side: usize,
    bytes: Vec<u8>,
    updates: Vec<Coordinate>,
    source_markers: SourceMarkers,
    descent_readback: DescentReadback,
}

#[derive(Serialize)]
struct Report {
    schema_version: u32,
    probe: &'static str,
    corpus_variant: CorpusVariant,
    timer_scope: &'static str,
    allocation: AllocationIdentity,
    samples: usize,
    warmups: usize,
    shapes: Vec<ShapeReport>,
}

#[derive(Serialize)]
struct CorpusVariant {
    name: &'static str,
    source_generator: &'static str,
    default_height: f64,
    default_descent: f64,
    metadata_application: &'static str,
}

#[derive(Serialize)]
struct AllocationIdentity {
    binary: &'static str,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Serialize)]
struct ShapeReport {
    corpus_variant: &'static str,
    shape: &'static str,
    sheet_count: usize,
    rows_per_sheet: usize,
    columns_per_sheet: usize,
    cells: usize,
    one_percent_update_count: usize,
    corpus_bytes: usize,
    corpus_sha256: String,
    source_markers: SourceMarkers,
    descent_readback: DescentReadback,
    scenarios: Vec<ScenarioReport>,
}

#[derive(Clone, Copy, Serialize)]
struct SourceMarkers {
    worksheet_parts_checked: usize,
    x14ac_namespace_bindings: usize,
    mce_namespace_bindings: usize,
    mce_ignorable_attributes: usize,
    dy_descent_attributes: usize,
}

#[derive(Clone, Copy, Serialize)]
struct DescentReadback {
    expected: f64,
    sheets_checked: usize,
    all_sheets_match: bool,
}

#[derive(Serialize)]
struct ScenarioReport {
    scenario: &'static str,
    warm_store: bool,
    changed: bool,
    update_count: usize,
    warmup_iterations: usize,
    sample_count: usize,
    sample_indices: Vec<usize>,
    elapsed_ns: Vec<u64>,
    stats: Stats,
    allocation_samples: Vec<allocation_metrics::Sample>,
    oracle: OracleReport,
}

#[derive(Serialize)]
struct OracleReport {
    iterations_checked: usize,
    patch_empty: Option<bool>,
    source_bytes_equal: Option<bool>,
    output_serialized: Option<bool>,
    untouched_data_equal: Option<bool>,
    changed_readback: Option<bool>,
    first_cell_readback: Option<bool>,
    defaults_descent_readback: Option<bool>,
    source_markers: Option<bool>,
}

#[derive(Serialize)]
struct Stats {
    min_ns: u64,
    p50_ns: u64,
    p95_ns: u64,
    p99_ns: u64,
    max_ns: u64,
    mean_ns: f64,
}

#[derive(Default)]
struct Config {
    samples: Option<usize>,
    warmups: Option<usize>,
    shapes: Option<Vec<Shape>>,
    scenarios: Option<Vec<Scenario>>,
    json: Option<PathBuf>,
}

fn main() {
    #[cfg(feature = "allocator-metrics")]
    allocation_metrics::enable();
    if let Err(error) = run() {
        eprintln!("xlsx commit guard probe failed: {error}");
        std::process::exit(1);
    }
}

fn run() -> ProbeResult<()> {
    let config = parse_args()?;
    let samples = config.samples.unwrap_or(DEFAULT_SAMPLES);
    let warmups = config.warmups.unwrap_or(DEFAULT_WARMUPS);
    if samples == 0 {
        return Err(probe_error("--samples must be greater than zero"));
    }
    let shapes = config.shapes.unwrap_or_else(|| Shape::ALL.to_vec());
    let scenarios = config.scenarios.unwrap_or_else(|| Scenario::ALL.to_vec());
    if shapes.is_empty() || scenarios.is_empty() {
        return Err(probe_error("at least one shape and scenario is required"));
    }

    let mut shape_reports = Vec::with_capacity(shapes.len());
    for shape in shapes {
        let corpus = build_corpus(shape)?;
        let mut scenario_reports = Vec::with_capacity(scenarios.len());
        for scenario in &scenarios {
            scenario_reports.push(run_scenario(&corpus, *scenario, warmups, samples)?);
        }
        let cells = SHEET_COUNT
            .checked_mul(corpus.side)
            .and_then(|value| value.checked_mul(corpus.side))
            .ok_or_else(|| probe_error("corpus cell count overflows usize"))?;
        shape_reports.push(ShapeReport {
            corpus_variant: CORPUS_VARIANT,
            shape: shape.name(),
            sheet_count: SHEET_COUNT,
            rows_per_sheet: corpus.side,
            columns_per_sheet: corpus.side,
            cells,
            one_percent_update_count: corpus.updates.len(),
            corpus_bytes: corpus.bytes.len(),
            corpus_sha256: sha256_hex(&corpus.bytes),
            source_markers: corpus.source_markers,
            descent_readback: corpus.descent_readback,
            scenarios: scenario_reports,
        });
    }

    let report = Report {
        schema_version: SCHEMA_VERSION,
        probe: "litchi-xlsx-public-commit-fallback-guard-v1",
        corpus_variant: CorpusVariant {
            name: CORPUS_VARIANT,
            source_generator: "public Workbook/Edit two-sheet numeric generator",
            default_height: DEFAULT_HEIGHT,
            default_descent: DEFAULT_DESCENT,
            metadata_application: "public WorksheetEdit::defaults().height(15).descent(0.2) before initial commit",
        },
        timer_scope: "Edit::commit only; fresh Workbook open, complete public Store warming, edit preparation, output serialization, source-marker/readback/cell oracles, and Commit/View drop are outside the clock",
        allocation: AllocationIdentity {
            binary: if cfg!(feature = "allocator-metrics") {
                "litchi-xlsx-commit-guard-alloc"
            } else {
                "litchi-xlsx-commit-guard"
            },
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
        samples,
        warmups,
        shapes: shape_reports,
    };
    write_report(&report, config.json.as_deref())
}

fn parse_args() -> ProbeResult<Config> {
    let mut config = Config::default();
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        let (key, inline_value) = argument
            .split_once('=')
            .map_or((argument.as_str(), None), |(key, value)| (key, Some(value)));
        match key {
            "--help" | "-h" => {
                print_help();
                std::process::exit(0);
            },
            "--samples" => {
                let value = value_or_next(inline_value, &mut arguments, "--samples")?;
                config.samples = Some(parse_usize(&value)?);
            },
            "--warmup" | "--warmups" => {
                let value = value_or_next(inline_value, &mut arguments, "--warmup")?;
                config.warmups = Some(parse_usize(&value)?);
            },
            "--shape" => {
                let value = value_or_next(inline_value, &mut arguments, "--shape")?;
                config.shapes = Some(parse_list(&value, parse_shape, "shape")?);
            },
            "--scenario" => {
                let value = value_or_next(inline_value, &mut arguments, "--scenario")?;
                config.scenarios = Some(parse_list(&value, parse_scenario, "scenario")?);
            },
            "--json" | "--output" => {
                config.json = Some(PathBuf::from(value_or_next(
                    inline_value,
                    &mut arguments,
                    "--json",
                )?));
            },
            _ => return Err(probe_error(format!("unknown argument '{argument}'"))),
        }
    }
    Ok(config)
}

fn value_or_next(
    inline: Option<&str>,
    arguments: &mut impl Iterator<Item = String>,
    option: &str,
) -> ProbeResult<String> {
    if let Some(value) = inline {
        return Ok(value.to_owned());
    }
    arguments
        .next()
        .ok_or_else(|| probe_error(format!("{option} requires a value")))
}

fn parse_usize(value: &str) -> ProbeResult<usize> {
    value
        .parse::<usize>()
        .map_err(|error| probe_error(format!("invalid integer '{value}': {error}")))
}

fn parse_list<T: Copy + PartialEq + std::fmt::Debug>(
    value: &str,
    parse: fn(&str) -> Option<T>,
    kind: &str,
) -> ProbeResult<Vec<T>> {
    let mut parsed = Vec::new();
    for item in value.split(',').filter(|item| !item.is_empty()) {
        let item = parse(item).ok_or_else(|| probe_error(format!("unknown {kind} '{item}'")))?;
        if parsed.contains(&item) {
            return Err(probe_error(format!("duplicate {kind} '{item:?}'")));
        }
        parsed.push(item);
    }
    if parsed.is_empty() {
        return Err(probe_error(format!("empty {kind} list")));
    }
    Ok(parsed)
}

fn parse_shape(value: &str) -> Option<Shape> {
    match value {
        "tiny" => Some(Shape::Tiny),
        "medium" => Some(Shape::Medium),
        "dense-wide" | "dense_wide" => Some(Shape::DenseWide),
        _ => None,
    }
}

fn parse_scenario(value: &str) -> Option<Scenario> {
    match value {
        "warm-changed-one-cell" | "warm_changed_one_cell" => Some(Scenario::WarmChangedOneCell),
        "warm-changed-one-percent" | "warm_changed_one_percent" => {
            Some(Scenario::WarmChangedOnePercent)
        },
        _ => None,
    }
}

fn print_help() {
    println!(
        "Usage: litchi-xlsx-commit-guard [OPTIONS]\n\n\
         --samples N             measured samples (default: 100)\n\
         --warmup N              warmup iterations (default: 3)\n\
         --shape LIST            tiny,medium,dense-wide (default: all)\n\
         --scenario LIST         warm-changed-one-cell,warm-changed-one-percent (default: both)\n\
         --json PATH             write report to PATH; use '-' for stdout\n\
         --help                  show this help"
    );
}

fn build_corpus(shape: Shape) -> ProbeResult<Corpus> {
    let side = shape.side();
    let workbook = Workbook::new()?;
    let mut edit = workbook.edit()?;
    for sheet_index in 0..SHEET_COUNT {
        if sheet_index == 0 {
            let mut sheet = edit
                .sheet("Sheet1")?
                .ok_or_else(|| probe_error("initial worksheet is missing"))?;
            {
                let mut defaults = sheet.defaults();
                defaults.height(DEFAULT_HEIGHT)?.descent(DEFAULT_DESCENT)?;
            }
            for row in 0..side {
                for column in 0..side {
                    sheet.set(
                        (u32::try_from(row)?, u32::try_from(column)?),
                        cell_value(Coordinate {
                            sheet: sheet_index,
                            row,
                            column,
                        }),
                    )?;
                }
            }
        } else {
            let mut sheet = edit.add(format!("Sheet{}", sheet_index + 1))?;
            {
                let mut defaults = sheet.defaults();
                defaults.height(DEFAULT_HEIGHT)?.descent(DEFAULT_DESCENT)?;
            }
            for row in 0..side {
                for column in 0..side {
                    sheet.set(
                        (u32::try_from(row)?, u32::try_from(column)?),
                        cell_value(Coordinate {
                            sheet: sheet_index,
                            row,
                            column,
                        }),
                    )?;
                }
            }
        }
    }
    let commit = edit.commit()?;
    let bytes = commit.workbook().to_bytes()?;
    let updates = one_percent_updates(side)?;
    let source_markers = inspect_source_markers(&bytes)?;
    let source_workbook = Workbook::from_bytes(bytes.clone())?;
    let descent_readback = verify_descent_readback(&source_workbook)?;
    let corpus = Corpus {
        side,
        bytes,
        updates,
        source_markers,
        descent_readback,
    };
    verify_corpus(&corpus)?;
    Ok(corpus)
}

fn one_percent_updates(side: usize) -> ProbeResult<Vec<Coordinate>> {
    let total = SHEET_COUNT
        .checked_mul(side)
        .and_then(|value| value.checked_mul(side))
        .ok_or_else(|| probe_error("update corpus cell count overflows usize"))?;
    let update_count = total
        .checked_add(99)
        .ok_or_else(|| probe_error("update count overflows usize"))?
        / 100;
    let per_sheet = side
        .checked_mul(side)
        .ok_or_else(|| probe_error("per-sheet cell count overflows usize"))?;
    let mut updates = Vec::with_capacity(update_count);
    for index in 0..update_count {
        let linear = index
            .checked_mul(total)
            .ok_or_else(|| probe_error("update position overflows usize"))?
            / update_count;
        let sheet = linear / per_sheet;
        let within_sheet = linear % per_sheet;
        updates.push(Coordinate {
            sheet,
            row: within_sheet / side,
            column: within_sheet % side,
        });
    }
    Ok(updates)
}

fn verify_corpus(corpus: &Corpus) -> ProbeResult<()> {
    let workbook = Workbook::from_bytes(corpus.bytes.clone())?;
    if workbook.len() != SHEET_COUNT {
        return Err(probe_error(
            "generated workbook has an unexpected sheet count",
        ));
    }
    for sheet_index in 0..SHEET_COUNT {
        let sheet = workbook
            .sheet(sheet_name(sheet_index).as_str())?
            .ok_or_else(|| probe_error("generated worksheet is missing"))?;
        let defaults = sheet
            .defaults()?
            .ok_or_else(|| probe_error("generated worksheet defaults are missing"))?;
        let descent = defaults.descent().map(litchi_xlsx::layout::Descent::get);
        if descent != Some(DEFAULT_DESCENT) {
            return Err(probe_error(format!(
                "generated worksheet default descent is {descent:?}, expected {DEFAULT_DESCENT}"
            )));
        }
        let count = sheet.cells(Rect::ALL)?.count();
        let expected = corpus
            .side
            .checked_mul(corpus.side)
            .ok_or_else(|| probe_error("expected cell count overflows usize"))?;
        if count != expected {
            return Err(probe_error(format!(
                "generated worksheet has {count} cells, expected {expected}"
            )));
        }
        for row in 0..corpus.side {
            for column in 0..corpus.side {
                let coordinate = Coordinate {
                    sheet: sheet_index,
                    row,
                    column,
                };
                let expected = cell_value(coordinate).to_string();
                let cell = sheet
                    .cell((u32::try_from(row)?, u32::try_from(column)?))?
                    .stored()
                    .ok_or_else(|| probe_error("generated numeric cell is missing"))?;
                if !matches!(cell, Cell::Value(Value::Number(value)) if value.as_str() == expected)
                {
                    return Err(probe_error("generated numeric cell differs from its value"));
                }
            }
        }
    }
    if corpus.source_markers.worksheet_parts_checked != SHEET_COUNT
        || corpus.source_markers.x14ac_namespace_bindings != SHEET_COUNT
        || corpus.source_markers.mce_namespace_bindings != SHEET_COUNT
        || corpus.source_markers.mce_ignorable_attributes != SHEET_COUNT
        || corpus.source_markers.dy_descent_attributes != SHEET_COUNT
        || !corpus.descent_readback.all_sheets_match
    {
        return Err(probe_error(
            "generated x14ac corpus marker proof is incomplete",
        ));
    }
    Ok(())
}

fn inspect_source_markers(bytes: &[u8]) -> ProbeResult<SourceMarkers> {
    let package = OpcPackage::from_bytes(bytes)?;
    let mut markers = SourceMarkers {
        worksheet_parts_checked: 0,
        x14ac_namespace_bindings: 0,
        mce_namespace_bindings: 0,
        mce_ignorable_attributes: 0,
        dy_descent_attributes: 0,
    };
    for sheet_index in 0..SHEET_COUNT {
        let partname = PackURI::new(format!("/xl/worksheets/sheet{}.xml", sheet_index + 1))?;
        let part = package.get_part(&partname)?;
        let xml = part.blob();
        let x14ac_binding = marker(
            xml,
            b"xmlns:x14ac=\"",
            X14AC_NAMESPACE,
            "x14ac namespace binding",
        )?;
        let mce_binding = marker(
            xml,
            b"xmlns:mc=\"",
            MCE_NAMESPACE,
            "markup-compatibility namespace binding",
        )?;
        if !xml
            .windows(b"mc:Ignorable=\"x14ac\"".len())
            .any(|window| window == b"mc:Ignorable=\"x14ac\"")
        {
            return Err(probe_error(format!(
                "worksheet part '{}' lacks its x14ac MCE Ignorable token",
                partname
            )));
        }
        if !xml
            .windows(b"x14ac:dyDescent=\"0.2\"".len())
            .any(|window| window == b"x14ac:dyDescent=\"0.2\"")
        {
            return Err(probe_error(format!(
                "worksheet part '{}' lacks serialized default descent",
                partname
            )));
        }
        markers.worksheet_parts_checked += 1;
        markers.x14ac_namespace_bindings += usize::from(x14ac_binding);
        markers.mce_namespace_bindings += usize::from(mce_binding);
        markers.mce_ignorable_attributes += 1;
        markers.dy_descent_attributes += 1;
    }
    Ok(markers)
}

fn marker(xml: &[u8], prefix: &[u8], value: &[u8], description: &str) -> ProbeResult<bool> {
    let mut expected = Vec::with_capacity(prefix.len() + value.len() + 1);
    expected.extend_from_slice(prefix);
    expected.extend_from_slice(value);
    expected.push(b'\"');
    if !xml
        .windows(expected.len())
        .any(|window| window == expected.as_slice())
    {
        return Err(probe_error(format!(
            "serialized source lacks {description}"
        )));
    }
    Ok(true)
}

fn verify_descent_readback(workbook: &Workbook) -> ProbeResult<DescentReadback> {
    if workbook.len() != SHEET_COUNT {
        return Err(probe_error(
            "descent readback has an unexpected sheet count",
        ));
    }
    let mut all_sheets_match = true;
    for sheet_index in 0..SHEET_COUNT {
        let sheet = workbook
            .sheet(sheet_name(sheet_index).as_str())?
            .ok_or_else(|| probe_error("descent readback worksheet is missing"))?;
        let descent = sheet
            .defaults()?
            .and_then(|defaults| defaults.descent())
            .map(litchi_xlsx::layout::Descent::get);
        all_sheets_match &= descent == Some(DEFAULT_DESCENT);
    }
    Ok(DescentReadback {
        expected: DEFAULT_DESCENT,
        sheets_checked: SHEET_COUNT,
        all_sheets_match,
    })
}

fn verify_numeric_workbook(
    workbook: &Workbook,
    side: usize,
    changed: &HashSet<Coordinate>,
) -> ProbeResult<bool> {
    if workbook.len() != SHEET_COUNT {
        return Err(probe_error("numeric output has an unexpected sheet count"));
    }
    let mut all_match = true;
    for sheet_index in 0..SHEET_COUNT {
        let sheet = workbook
            .sheet(sheet_name(sheet_index).as_str())?
            .ok_or_else(|| probe_error("numeric output worksheet is missing"))?;
        for row in 0..side {
            for column in 0..side {
                let coordinate = Coordinate {
                    sheet: sheet_index,
                    row,
                    column,
                };
                let expected =
                    (cell_value(coordinate) + i32::from(changed.contains(&coordinate))).to_string();
                let cell = sheet
                    .cell((u32::try_from(row)?, u32::try_from(column)?))?
                    .stored()
                    .ok_or_else(|| probe_error("numeric output cell is missing"))?;
                all_match &= matches!(
                    cell,
                    Cell::Value(Value::Number(value)) if value.as_str() == expected
                );
            }
        }
    }
    Ok(all_match)
}

fn run_scenario(
    corpus: &Corpus,
    scenario: Scenario,
    warmups: usize,
    samples: usize,
) -> ProbeResult<ScenarioReport> {
    let updates = if scenario.one_cell() {
        corpus
            .updates
            .first()
            .copied()
            .map(|coordinate| vec![coordinate])
            .ok_or_else(|| probe_error("corpus has no update coordinate"))?
    } else {
        corpus.updates.clone()
    };
    let update_set = updates.iter().copied().collect::<HashSet<_>>();
    let total_iterations = warmups
        .checked_add(samples)
        .ok_or_else(|| probe_error("iteration count overflows usize"))?;
    let mut elapsed_ns = Vec::with_capacity(samples);
    let mut sample_indices = Vec::with_capacity(samples);
    let mut patch_empty = None;
    let mut source_bytes_equal = None;
    let mut output_serialized = None;
    let mut untouched_data_equal = None;
    let mut changed_readback = None;
    let mut defaults_descent_readback = None;
    let mut source_markers = None;
    let mut retained_commit = None;
    let mut allocation_samples = Vec::with_capacity(samples);

    for iteration in 0..total_iterations {
        let workbook = Workbook::from_bytes(corpus.bytes.clone())?;
        if scenario.warm_store() {
            warm_stores(&workbook, &updates, corpus.side)?;
        }
        let edit = prepare_edit(&workbook, &updates, true)?;
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let commit = edit.commit()?;
        let elapsed = started.elapsed();
        let allocation_sample = allocation_region
            .finish()
            .unwrap_or_else(allocation_metrics::unavailable_sample);

        // Keep the result live and run every oracle after the measured clock.
        std::hint::black_box(&commit);
        let is_empty = commit.patch().is_empty();
        let output = commit.workbook().to_bytes()?;
        let output_workbook = Workbook::from_bytes(output.clone())?;
        let output_markers = inspect_source_markers(&output)?;
        let marker_ok = output_markers.worksheet_parts_checked == SHEET_COUNT
            && output_markers.x14ac_namespace_bindings == SHEET_COUNT
            && output_markers.mce_namespace_bindings == SHEET_COUNT
            && output_markers.mce_ignorable_attributes == SHEET_COUNT
            && output_markers.dy_descent_attributes == SHEET_COUNT;
        let descent = verify_descent_readback(&output_workbook)?;
        let complete_values = verify_numeric_workbook(&output_workbook, corpus.side, &update_set)?;
        let readback = verify_changed_readback(&commit, &updates)?;
        if is_empty {
            return Err(probe_error(format!(
                "{} produced an empty patch for a changed edit",
                scenario.name()
            )));
        }
        if !marker_ok || !descent.all_sheets_match || !complete_values || !readback {
            return Err(probe_error(format!(
                "{} failed its post-commit oracle",
                scenario.name()
            )));
        }
        patch_empty = Some(is_empty);
        source_bytes_equal = None;
        output_serialized = Some(true);
        untouched_data_equal = Some(complete_values);
        changed_readback = Some(readback);
        defaults_descent_readback = Some(descent.all_sheets_match);
        source_markers = Some(marker_ok);

        if iteration >= warmups {
            let sample_index = iteration - warmups;
            sample_indices.push(sample_index);
            elapsed_ns.push(
                u64::try_from(elapsed.as_nanos())
                    .map_err(|_| probe_error("elapsed time does not fit u64"))?,
            );
            allocation_samples.push(allocation_sample);
        }
        // Assignment and the eventual drop happen outside the clock. Keeping
        // the last successful commit retained also prevents a dead-result
        // optimizer from changing the public-API workload.
        retained_commit = Some(commit);
    }
    std::hint::black_box(&retained_commit);
    if elapsed_ns.len() != samples
        || allocation_samples.len() != samples
        || sample_indices != (0..samples).collect::<Vec<_>>()
    {
        return Err(probe_error("measured sample indices are not contiguous"));
    }
    let stats = stats(&elapsed_ns)?;
    Ok(ScenarioReport {
        scenario: scenario.name(),
        warm_store: scenario.warm_store(),
        changed: true,
        update_count: updates.len(),
        warmup_iterations: warmups,
        sample_count: samples,
        sample_indices,
        elapsed_ns,
        stats,
        allocation_samples,
        oracle: OracleReport {
            iterations_checked: total_iterations,
            patch_empty,
            source_bytes_equal,
            output_serialized,
            untouched_data_equal,
            changed_readback,
            first_cell_readback: None,
            defaults_descent_readback,
            source_markers,
        },
    })
}

fn warm_stores(workbook: &Workbook, updates: &[Coordinate], side: usize) -> ProbeResult<()> {
    let mut warmed = [false; SHEET_COUNT];
    for coordinate in updates {
        if warmed[coordinate.sheet] {
            continue;
        }
        let sheet = workbook
            .sheet(sheet_name(coordinate.sheet).as_str())?
            .ok_or_else(|| probe_error("Store warmup target worksheet is missing"))?;
        let count = sheet.cells(Rect::ALL)?.count();
        let expected = side
            .checked_mul(side)
            .ok_or_else(|| probe_error("Store warmup cell count overflows usize"))?;
        if count != expected {
            return Err(probe_error(
                "Store warmup observed an unexpected cell count",
            ));
        }
        warmed[coordinate.sheet] = true;
    }
    Ok(())
}

fn prepare_edit(
    workbook: &Workbook,
    updates: &[Coordinate],
    changed: bool,
) -> ProbeResult<litchi_xlsx::Edit> {
    let mut edit = workbook.edit()?;
    for coordinate in updates {
        let mut sheet = edit
            .sheet(sheet_name(coordinate.sheet).as_str())?
            .ok_or_else(|| probe_error("edit target worksheet is missing"))?;
        let value = cell_value(*coordinate) + i32::from(changed);
        sheet.set(
            (
                u32::try_from(coordinate.row)?,
                u32::try_from(coordinate.column)?,
            ),
            value,
        )?;
    }
    Ok(edit)
}

fn verify_changed_readback(
    commit: &litchi_xlsx::Commit,
    updates: &[Coordinate],
) -> ProbeResult<bool> {
    for coordinate in updates {
        let sheet = commit
            .workbook()
            .sheet(sheet_name(coordinate.sheet).as_str())?
            .ok_or_else(|| probe_error("changed commit target worksheet is missing"))?;
        let cell = sheet
            .cell((
                u32::try_from(coordinate.row)?,
                u32::try_from(coordinate.column)?,
            ))?
            .stored()
            .ok_or_else(|| probe_error("changed commit target cell is missing"))?;
        let expected = (cell_value(*coordinate) + 1).to_string();
        if !matches!(cell, Cell::Value(Value::Number(value)) if value.as_str() == expected) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn stats(values: &[u64]) -> ProbeResult<Stats> {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    let first = sorted
        .first()
        .copied()
        .ok_or_else(|| probe_error("cannot summarize an empty sample vector"))?;
    let last = sorted
        .last()
        .copied()
        .ok_or_else(|| probe_error("cannot summarize an empty sample vector"))?;
    let sum = values
        .iter()
        .try_fold(0_u128, |sum, value| sum.checked_add(u128::from(*value)))
        .ok_or_else(|| probe_error("sample sum overflows u128"))?;
    let length = u128::try_from(values.len())?;
    Ok(Stats {
        min_ns: first,
        p50_ns: percentile(&sorted, 50),
        p95_ns: percentile(&sorted, 95),
        p99_ns: percentile(&sorted, 99),
        max_ns: last,
        mean_ns: (sum as f64) / (length as f64),
    })
}

fn percentile(sorted: &[u64], percentile: usize) -> u64 {
    let index = (sorted.len() - 1) * percentile / 100;
    sorted[index]
}

fn sheet_name(index: usize) -> String {
    format!("Sheet{}", index + 1)
}

fn cell_value(coordinate: Coordinate) -> i32 {
    let sheet = i32::try_from(coordinate.sheet).expect("bounded sheet index fits i32");
    let row = i32::try_from(coordinate.row).expect("bounded row index fits i32");
    let column = i32::try_from(coordinate.column).expect("bounded column index fits i32");
    sheet * 1_000_000 + row * 1_000 + column
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

fn write_report(report: &Report, path: Option<&std::path::Path>) -> ProbeResult<()> {
    if path.is_some_and(|path| path == std::path::Path::new("-")) || path.is_none() {
        let stdout = io::stdout();
        let mut output = stdout.lock();
        serde_json::to_writer_pretty(&mut output, report)?;
        output.write_all(b"\n")?;
        return Ok(());
    }
    let path = path.expect("path was checked above");
    let mut output = File::options().write(true).create_new(true).open(path)?;
    serde_json::to_writer_pretty(&mut output, report)?;
    output.write_all(b"\n")?;
    Ok(())
}

fn probe_error(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(io::ErrorKind::InvalidData, message.into()))
}
