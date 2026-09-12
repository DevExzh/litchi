//! Standalone public-API guard for XLSX semantic commit work.
//!
//! The probe deliberately lives outside the performance harness.  It creates
//! one deterministic two-sheet numeric workbook for each selected shape, then
//! measures only `Edit::commit` after opening a fresh workbook and preparing
//! the requested edit.  Its first-cell read scenario measures only the first
//! public Store-forcing cell lookup. Store warming, output/readback oracles,
//! and returned commit destruction all remain outside the measured clock.

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
    error::Error,
    fs::File,
    io::{self, Write},
    path::PathBuf,
    time::Instant,
};

use litchi_xlsx::{Cell, Rect, Value, Workbook};
use serde::Serialize;
use sha2::{Digest, Sha256};

const SCHEMA_VERSION: u32 = 1;
const SHEET_COUNT: usize = 2;
const DEFAULT_SAMPLES: usize = 100;
const DEFAULT_WARMUPS: usize = 3;

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
    ColdFirstCellRead,
    ColdSameOneCell,
    ColdSameOnePercent,
    WarmSameOneCell,
    WarmSameOnePercent,
    WarmChangedOneCell,
    WarmChangedOnePercent,
}

impl Scenario {
    const ALL: [Self; 7] = [
        Self::ColdFirstCellRead,
        Self::ColdSameOneCell,
        Self::ColdSameOnePercent,
        Self::WarmSameOneCell,
        Self::WarmSameOnePercent,
        Self::WarmChangedOneCell,
        Self::WarmChangedOnePercent,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::ColdFirstCellRead => "cold-first-cell-read",
            Self::ColdSameOneCell => "cold-same-one-cell",
            Self::ColdSameOnePercent => "cold-same-one-percent",
            Self::WarmSameOneCell => "warm-same-one-cell",
            Self::WarmSameOnePercent => "warm-same-one-percent",
            Self::WarmChangedOneCell => "warm-changed-one-cell",
            Self::WarmChangedOnePercent => "warm-changed-one-percent",
        }
    }

    const fn warm_store(self) -> bool {
        matches!(
            self,
            Self::WarmSameOneCell
                | Self::WarmSameOnePercent
                | Self::WarmChangedOneCell
                | Self::WarmChangedOnePercent
        )
    }

    const fn changed(self) -> bool {
        matches!(self, Self::WarmChangedOneCell | Self::WarmChangedOnePercent)
    }

    const fn one_cell(self) -> bool {
        matches!(
            self,
            Self::ColdSameOneCell | Self::WarmSameOneCell | Self::WarmChangedOneCell
        )
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Coordinate {
    sheet: usize,
    row: usize,
    column: usize,
}

struct Corpus {
    side: usize,
    bytes: Vec<u8>,
    updates: Vec<Coordinate>,
}

#[derive(Serialize)]
struct Report {
    schema_version: u32,
    probe: &'static str,
    timer_scope: &'static str,
    allocation: AllocationIdentity,
    samples: usize,
    warmups: usize,
    shapes: Vec<ShapeReport>,
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
    shape: &'static str,
    sheet_count: usize,
    rows_per_sheet: usize,
    columns_per_sheet: usize,
    cells: usize,
    one_percent_update_count: usize,
    corpus_bytes: usize,
    corpus_sha256: String,
    scenarios: Vec<ScenarioReport>,
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
    changed_readback: Option<bool>,
    first_cell_readback: Option<bool>,
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
            shape: shape.name(),
            sheet_count: SHEET_COUNT,
            rows_per_sheet: corpus.side,
            columns_per_sheet: corpus.side,
            cells,
            one_percent_update_count: corpus.updates.len(),
            corpus_bytes: corpus.bytes.len(),
            corpus_sha256: sha256_hex(&corpus.bytes),
            scenarios: scenario_reports,
        });
    }

    let report = Report {
        schema_version: SCHEMA_VERSION,
        probe: "litchi-xlsx-public-commit-guard-v1",
        timer_scope: "Edit::commit or first public Worksheet::cell Store load only; fresh Workbook open, Store warming, edit preparation, output/readback oracles, and Commit/View drop are outside the clock",
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
        "cold-first-cell-read" | "cold_first_cell_read" => Some(Scenario::ColdFirstCellRead),
        "cold-same-one-cell" | "cold_same_one_cell" => Some(Scenario::ColdSameOneCell),
        "cold-same-one-percent" | "cold_same_one_percent" => Some(Scenario::ColdSameOnePercent),
        "warm-same-one-cell" | "warm_same_one_cell" => Some(Scenario::WarmSameOneCell),
        "warm-same-one-percent" | "warm_same_one_percent" => Some(Scenario::WarmSameOnePercent),
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
         --scenario LIST         comma-separated scenario names (default: all)\n\
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
    let corpus = Corpus {
        side,
        bytes,
        updates,
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
    Ok(())
}

fn run_scenario(
    corpus: &Corpus,
    scenario: Scenario,
    warmups: usize,
    samples: usize,
) -> ProbeResult<ScenarioReport> {
    if scenario == Scenario::ColdFirstCellRead {
        return run_first_cell_read(corpus, warmups, samples);
    }
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
    let total_iterations = warmups
        .checked_add(samples)
        .ok_or_else(|| probe_error("iteration count overflows usize"))?;
    let mut elapsed_ns = Vec::with_capacity(samples);
    let mut sample_indices = Vec::with_capacity(samples);
    let mut patch_empty = None;
    let mut source_bytes_equal = None;
    let mut changed_readback = None;
    let mut retained_commit = None;
    let mut allocation_samples = Vec::with_capacity(samples);

    for iteration in 0..total_iterations {
        let workbook = Workbook::from_bytes(corpus.bytes.clone())?;
        if scenario.warm_store() {
            warm_stores(&workbook, &updates, corpus.side)?;
        }
        let edit = prepare_edit(&workbook, &updates, scenario.changed())?;
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
        let exact_source = if scenario.changed() {
            None
        } else {
            Some(commit.workbook().to_bytes()? == corpus.bytes)
        };
        let readback = if scenario.changed() {
            Some(verify_changed_readback(&commit, &updates)?)
        } else {
            None
        };
        let expected_empty = !scenario.changed();
        if is_empty != expected_empty {
            return Err(probe_error(format!(
                "{} produced patch_empty={is_empty}, expected {expected_empty}",
                scenario.name()
            )));
        }
        if exact_source == Some(false) || readback == Some(false) {
            return Err(probe_error(format!(
                "{} failed its post-commit oracle",
                scenario.name()
            )));
        }
        patch_empty = Some(is_empty);
        source_bytes_equal = exact_source;
        changed_readback = readback;

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
        changed: scenario.changed(),
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
            changed_readback,
            first_cell_readback: None,
        },
    })
}

fn run_first_cell_read(
    corpus: &Corpus,
    warmups: usize,
    samples: usize,
) -> ProbeResult<ScenarioReport> {
    let total_iterations = warmups
        .checked_add(samples)
        .ok_or_else(|| probe_error("iteration count overflows usize"))?;
    let mut elapsed_ns = Vec::with_capacity(samples);
    let mut sample_indices = Vec::with_capacity(samples);
    let mut first_cell_readback = None;
    let mut allocation_samples = Vec::with_capacity(samples);

    for iteration in 0..total_iterations {
        let workbook = Workbook::from_bytes(corpus.bytes.clone())?;
        let sheet = workbook
            .sheet("Sheet1")?
            .ok_or_else(|| probe_error("first-cell read worksheet is missing"))?;
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let cell = sheet.cell((0_u32, 0_u32))?;
        let elapsed = started.elapsed();
        let allocation_sample = allocation_region
            .finish()
            .unwrap_or_else(allocation_metrics::unavailable_sample);

        // The Store parse is forced by `cell`; retain the returned view until
        // after the clock and check its numeric value outside the clock.
        std::hint::black_box(&cell);
        let stored = cell
            .stored()
            .ok_or_else(|| probe_error("first-cell read returned a missing cell"))?;
        let expected = cell_value(Coordinate {
            sheet: 0,
            row: 0,
            column: 0,
        })
        .to_string();
        let readback = matches!(
            stored,
            Cell::Value(Value::Number(value)) if value.as_str() == expected
        );
        if !readback {
            return Err(probe_error("first-cell readback differs from corpus value"));
        }
        first_cell_readback = Some(true);

        if iteration >= warmups {
            let sample_index = iteration - warmups;
            sample_indices.push(sample_index);
            elapsed_ns.push(
                u64::try_from(elapsed.as_nanos())
                    .map_err(|_| probe_error("elapsed time does not fit u64"))?,
            );
            allocation_samples.push(allocation_sample);
        }
        // `cell`, `sheet`, and `workbook` are dropped after the timed region.
        std::hint::black_box(&cell);
    }
    if elapsed_ns.len() != samples
        || allocation_samples.len() != samples
        || sample_indices != (0..samples).collect::<Vec<_>>()
    {
        return Err(probe_error("measured sample indices are not contiguous"));
    }
    Ok(ScenarioReport {
        scenario: Scenario::ColdFirstCellRead.name(),
        warm_store: false,
        changed: false,
        update_count: 0,
        warmup_iterations: warmups,
        sample_count: samples,
        sample_indices,
        stats: stats(&elapsed_ns)?,
        elapsed_ns,
        allocation_samples,
        oracle: OracleReport {
            iterations_checked: total_iterations,
            patch_empty: None,
            source_bytes_equal: None,
            changed_readback: None,
            first_cell_readback,
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
