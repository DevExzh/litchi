//! Opt-in DOCX owned-ingress attribution diagnostics.
//!
//! This module is deliberately separate from the default case matrix. It
//! brackets only the owned-vector/owned-reader preparation and open phases,
//! and records the procfs observer work in an explicit gap segment. Workload
//! mode keeps the existing section-layout publication and inverse checks;
//! calibration mode is a pure black-box boundary control and never claims
//! DOCX correctness.

use std::error::Error;
use std::ffi::OsString;
use std::fs::OpenOptions;
use std::hint::black_box;
use std::io::Write;
use std::path::PathBuf;
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};
use std::time::Instant;

use litchi_core::{OwnedSource, Position};
use litchi_opc::SourceBackedPackage;
use serde::Serialize;

use crate::allocation_metrics;
use crate::process_metrics;

const SCHEMA: &str = "docx_owned_ingress_attribution_v1";
const ROUTE_VEC: &str = "owned_vec";
const ROUTE_READER: &str = "owned_reader";
const ROUTE_CALIBRATION: &str = "route_independent";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Route {
    OwnedVec,
    OwnedReader,
}

impl Route {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            ROUTE_VEC => Ok(Self::OwnedVec),
            ROUTE_READER => Ok(Self::OwnedReader),
            _ => Err("--route must be owned_vec or owned_reader".into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::OwnedVec => ROUTE_VEC,
            Self::OwnedReader => ROUTE_READER,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Case {
    Workload,
    Calibration,
}

impl Case {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "workload" => Ok(Self::Workload),
            "calibration" => Ok(Self::Calibration),
            _ => Err("--case must be workload or calibration".into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Workload => "workload",
            Self::Calibration => "calibration",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Diagnostic {
    Off,
    On,
}

impl Diagnostic {
    fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "off" => Ok(Self::Off),
            "on" => Ok(Self::On),
            _ => Err("--diagnostic must be off or on".into()),
        }
    }

    const fn enabled(self) -> bool {
        matches!(self, Self::On)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    case: Case,
    route: Option<Route>,
    diagnostic: Diagnostic,
    warmups: usize,
    samples: usize,
    source_revision: String,
    output: Option<PathBuf>,
}

#[derive(Clone, Debug, Serialize)]
struct ProcessObservation {
    before: Option<process_metrics::Snapshot>,
    after: Option<process_metrics::Snapshot>,
    before_status: &'static str,
    after_status: &'static str,
    delta: Option<process_metrics::Delta>,
    /// Unlike Delta::rss_bytes, this keeps decreases visible instead of
    /// turning them into a saturating zero.
    rss_delta_signed: Option<i64>,
    rss_delta_status: &'static str,
}

#[derive(Clone, Copy)]
struct ProcessRead {
    snapshot: Option<process_metrics::Snapshot>,
    status: &'static str,
}

#[derive(Clone, Debug, Serialize)]
struct AllocationPhases {
    prep: Vec<allocation_metrics::Sample>,
    observer_gap: Vec<allocation_metrics::Sample>,
    open: Vec<allocation_metrics::Sample>,
    combined_prep_gap_open: Vec<allocation_metrics::Sample>,
}

#[derive(Clone, Debug, Serialize)]
struct ProcessPhases {
    prep: Vec<ProcessObservation>,
    observer_gap: Vec<ProcessObservation>,
    open: Vec<ProcessObservation>,
    combined_prep_gap_open: Vec<ProcessObservation>,
}

#[derive(Clone, Debug, Serialize)]
struct SinkObservation {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
}

#[derive(Clone, Debug, Serialize)]
struct CorrectnessGates {
    snapshot_shape_verified: bool,
    changed_nonnoop_columns_verified: bool,
    patch_forward_inverse_stale_verified: bool,
    publication_inverse_applicable_verified: bool,
    output_verified: bool,
    reopened_output_verified: bool,
    inverse_source_verified: bool,
    source_immutability_verified: bool,
    cache_contract_verified: bool,
    cache_deterministic_verified: bool,
    sink_bounds_verified: bool,
    owned_reader_ingress_verified: bool,
    cache_hits: u64,
    cache_cold_loads: u64,
    cache_successful_loads: u64,
    cache_retained_entries: usize,
    cache_retained_bytes: usize,
}

#[derive(Serialize)]
struct Report {
    schema: &'static str,
    scenario: &'static str,
    case: &'static str,
    route: &'static str,
    diagnostic_enabled: bool,
    diagnostic_arm: String,
    binary_identity: &'static str,
    allocator_identity: &'static str,
    instrumentation_identity: &'static str,
    allocator_counter_revision: Option<&'static str>,
    source_revision: String,
    source_corpus_generator: &'static str,
    source_bytes: u64,
    source_sha256: String,
    expected_output_bytes: Option<u64>,
    expected_output_sha256: Option<String>,
    warmup_iterations: usize,
    sample_count: usize,
    timing_scope: &'static str,
    observer_gap_scope: &'static str,
    operation_elapsed_ns: super::Statistics,
    wall_elapsed_ns: super::Statistics,
    sample_order: Vec<usize>,
    preparation_ns: Vec<u64>,
    observer_gap_ns: Vec<u64>,
    open_ns: Vec<u64>,
    query_ns: Vec<u64>,
    edit_ns: Vec<u64>,
    commit_ns: Vec<u64>,
    publication_ns: Vec<u64>,
    correctness_status: &'static str,
    correctness_verified: Vec<bool>,
    inverse_source_verified: Vec<bool>,
    output_sha256: Vec<Option<String>>,
    correctness_gates: Vec<CorrectnessGates>,
    sink: Vec<SinkObservation>,
    owned_ingress_read_calls: Vec<Option<u64>>,
    owned_ingress_read_bytes: Vec<Option<u64>>,
    owned_input_bytes: Vec<u64>,
    allocation: Option<AllocationPhases>,
    process: Option<ProcessPhases>,
}

struct Oracle {
    output: Vec<u8>,
    output_sha256: String,
    maximum_sink_bytes: u64,
}

fn parse_usize(value: &str, flag: &str) -> Result<usize, Box<dyn Error>> {
    value
        .parse::<usize>()
        .map_err(|error| format!("{flag} must be a non-negative integer: {error}").into())
}

fn parse_source_revision(value: &str) -> Result<String, Box<dyn Error>> {
    if value.len() != 40 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err("--source-revision must be exactly 40 hexadecimal characters".into());
    }
    Ok(value.to_owned())
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut case = None;
    let mut route = None;
    let mut diagnostic = None;
    let mut warmups = None;
    let mut samples = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("attribution argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("attribution argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--case" => {
                if case.is_some() {
                    return Err("duplicate --case".into());
                }
                case = Some(Case::parse(value_text)?);
            },
            "--route" => {
                if route.is_some() {
                    return Err("duplicate --route".into());
                }
                route = Some(Route::parse(value_text)?);
            },
            "--diagnostic" => {
                if diagnostic.is_some() {
                    return Err("duplicate --diagnostic".into());
                }
                diagnostic = Some(Diagnostic::parse(value_text)?);
            },
            "--warmups" => {
                if warmups.is_some() {
                    return Err("duplicate --warmups".into());
                }
                warmups = Some(parse_usize(value_text, "--warmups")?);
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("duplicate --samples".into());
                }
                let parsed = parse_usize(value_text, "--samples")?;
                if parsed == 0 {
                    return Err("--samples must be greater than zero".into());
                }
                samples = Some(parsed);
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("duplicate --source-revision".into());
                }
                source_revision = Some(parse_source_revision(value_text)?);
            },
            "--output" => {
                if output.is_some() {
                    return Err("duplicate --output".into());
                }
                if value_text.is_empty() {
                    return Err("--output cannot be empty".into());
                }
                output = Some(PathBuf::from(value));
            },
            _ => return Err(format!("unknown attribution argument: {flag}").into()),
        }
    }
    let case = case.ok_or("missing --case")?;
    let diagnostic = diagnostic.ok_or("missing --diagnostic")?;
    if case == Case::Workload && route.is_none() {
        return Err("workload case requires --route".into());
    }
    if case == Case::Calibration && route.is_some() {
        return Err("calibration case is route-independent; omit --route".into());
    }
    if case == Case::Calibration && !diagnostic.enabled() {
        return Err("calibration case requires --diagnostic on".into());
    }
    Ok(Config {
        case,
        route,
        diagnostic,
        warmups: warmups.ok_or("missing --warmups")?,
        samples: samples.ok_or("missing --samples")?,
        source_revision: source_revision.ok_or("missing --source-revision")?,
        output,
    })
}

fn process_observation(before: ProcessRead, after: ProcessRead) -> ProcessObservation {
    let delta = match (before.snapshot, after.snapshot) {
        (Some(before), Some(after)) => Some(after.delta(before)),
        _ => None,
    };
    let rss_delta_signed = match (before.snapshot, after.snapshot) {
        (Some(before), Some(after)) => i64::try_from(after.rss_bytes).ok().and_then(|after| {
            i64::try_from(before.rss_bytes)
                .ok()
                .and_then(|before| after.checked_sub(before))
        }),
        _ => None,
    };
    let rss_delta_status = match (before.snapshot, after.snapshot) {
        (Some(before), Some(after))
            if i64::try_from(before.rss_bytes).is_ok()
                && i64::try_from(after.rss_bytes).is_ok() =>
        {
            "measured"
        },
        (Some(_), Some(_)) => "overflow",
        _ => "unavailable",
    };
    ProcessObservation {
        before: before.snapshot,
        after: after.snapshot,
        before_status: before.status,
        after_status: after.status,
        delta,
        rss_delta_signed,
        rss_delta_status,
    }
}

fn diagnostic_arm(enabled: bool) -> String {
    match (allocation_metrics::binary_identity(), enabled) {
        ("litchi-perf-baseline", false) => "normal_off".to_owned(),
        ("litchi-perf-baseline", true) => "normal_on".to_owned(),
        ("litchi-perf-baseline-alloc", true) => "allocator_on".to_owned(),
        (binary, false) => format!("{binary}_off"),
        (binary, true) => format!("{binary}_on"),
    }
}

fn read_process(enabled: bool) -> ProcessRead {
    if !enabled {
        return ProcessRead {
            snapshot: None,
            status: "disabled",
        };
    }
    match process_metrics::Snapshot::read() {
        Ok(snapshot) => ProcessRead {
            snapshot: Some(snapshot),
            status: "measured",
        },
        Err(error) => ProcessRead {
            snapshot: None,
            status: match error.kind() {
                std::io::ErrorKind::NotFound => "procfs_not_found",
                std::io::ErrorKind::PermissionDenied => "procfs_permission_denied",
                std::io::ErrorKind::InvalidData => "procfs_invalid_data",
                _ => "procfs_read_error",
            },
        },
    }
}

fn allocation_sample(sample: Option<allocation_metrics::Sample>) -> allocation_metrics::Sample {
    sample.unwrap_or_else(allocation_metrics::unavailable_sample)
}

fn prepare_oracle(corpus: &super::Corpus) -> Result<Oracle, Box<dyn Error>> {
    let package = litchi_docx::source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        corpus.archive.clone(),
    )))?;
    let snapshot = package.section_layout_snapshot()?;
    super::verify_docx_section_layout_snapshot(&snapshot, [2, 2, 2])?;
    let mut edit = package.edit_section_layout(litchi_docx::section::Selector::paragraph(
        Position::new(super::DOCX_SECTION_LAYOUT_SECTION_POSITIONS[0]),
    ))?;
    edit.set_columns(Some(litchi_docx::section::Columns {
        equal_width: true,
        count: 3,
        space: Some(litchi_docx::section::Emu::from_twips(240)),
        separator: true,
        columns: Vec::new(),
    }))?;
    let commit = edit.commit()?;
    if commit.patch().is_noop() {
        return Err("DOCX owned attribution oracle commit is not changed".into());
    }
    let mut output = Vec::new();
    let publication = package.apply_section_layout(commit.patch())?;
    publication.package().write_to(&mut output)?;
    super::verify_docx_section_layout_output(corpus, &output, [3, 2, 2])?;
    super::verify_docx_section_layout_controls(corpus, &publication)?;
    let output_sha256 = super::sha256_hex(&output);
    let maximum_sink_bytes = u64::try_from(output.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("DOCX owned attribution sink ceiling overflows u64")?;
    Ok(Oracle {
        output,
        output_sha256,
        maximum_sink_bytes,
    })
}

fn open_owned_package(
    route: Route,
    input: Vec<u8>,
    read_calls: Option<Arc<AtomicU64>>,
    read_bytes: Option<Arc<AtomicU64>>,
) -> Result<litchi_docx::source_backed::Package, Box<dyn Error>> {
    match route {
        Route::OwnedVec => {
            let opc = SourceBackedPackage::from_vec(
                input,
                litchi_opc::SourceLimits::default(),
                litchi_opc::SourceCacheLimits::default(),
                None,
            )?;
            Ok(litchi_docx::source_backed::Package::from_source_backed_package(opc)?)
        },
        Route::OwnedReader => {
            let reader = super::OwnedIngressReader::new(
                input,
                read_calls.ok_or("owned-reader read-call observer is missing")?,
                read_bytes.ok_or("owned-reader byte observer is missing")?,
            );
            Ok(litchi_docx::source_backed::Package::from_reader(reader)?)
        },
    }
}

fn reorder<T: Clone>(values: &mut Vec<T>, order: &[usize]) -> Result<(), Box<dyn Error>> {
    super::reorder_sample_vector(values, order)
}

fn reorder_if_present<T: Clone>(
    values: &mut Vec<T>,
    order: &[usize],
) -> Result<(), Box<dyn Error>> {
    if values.is_empty() {
        Ok(())
    } else {
        reorder(values, order)
    }
}

fn write_report(config: &Config, report: &Report) -> Result<(), Box<dyn Error>> {
    let bytes = serde_json::to_vec_pretty(report)?;
    if let Some(path) = &config.output {
        let mut output = OpenOptions::new().create_new(true).write(true).open(path)?;
        output.write_all(&bytes)?;
        output.write_all(b"\n")?;
        output.flush()?;
    }
    println!(
        "DOCX_OWNED_INGRESS_ATTRIBUTION_JSON={}",
        serde_json::to_string(report)?
    );
    Ok(())
}

fn run_calibration(config: &Config, corpus: &super::Corpus) -> Result<Report, Box<dyn Error>> {
    let mut operation = Vec::with_capacity(config.samples);
    let mut wall = Vec::with_capacity(config.samples);
    let mut preparation = Vec::with_capacity(config.samples);
    let mut gap = Vec::with_capacity(config.samples);
    let mut open = Vec::with_capacity(config.samples);
    let mut allocations = AllocationPhases {
        prep: Vec::with_capacity(config.samples),
        observer_gap: Vec::with_capacity(config.samples),
        open: Vec::with_capacity(config.samples),
        combined_prep_gap_open: Vec::with_capacity(config.samples),
    };
    let mut processes = ProcessPhases {
        prep: Vec::with_capacity(config.samples),
        observer_gap: Vec::with_capacity(config.samples),
        open: Vec::with_capacity(config.samples),
        combined_prep_gap_open: Vec::with_capacity(config.samples),
    };
    for iteration in 0..super::iteration_count(config.warmups, config.samples)? {
        let wall_started = Instant::now();
        let before_prep = read_process(true);
        let mut region = allocation_metrics::begin();
        let prep_started = Instant::now();
        black_box(corpus.archive.len());
        let prep_ns = super::elapsed_ns(prep_started.elapsed())?;
        let prep_sample = allocation_sample(region.split());
        let gap_started = Instant::now();
        let after_prep = read_process(true);
        let before_open = read_process(true);
        let gap_sample = allocation_sample(region.split());
        let gap_ns = super::elapsed_ns(gap_started.elapsed())?;
        let open_started = Instant::now();
        black_box(corpus.archive.as_slice());
        let open_ns = super::elapsed_ns(open_started.elapsed())?;
        let (open_sample, combined_sample) = region.finish_split();
        let after_open = read_process(true);
        let wall_ns = super::elapsed_ns(wall_started.elapsed())?;
        let open_sample = allocation_sample(open_sample);
        let combined_sample = allocation_sample(combined_sample);
        if iteration >= config.warmups {
            let operation_ns = prep_ns
                .checked_add(open_ns)
                .ok_or("calibration operation timing overflowed")?;
            operation.push(operation_ns);
            wall.push(wall_ns);
            preparation.push(prep_ns);
            gap.push(gap_ns);
            open.push(open_ns);
            if config.diagnostic.enabled() {
                allocations.prep.push(prep_sample);
                allocations.observer_gap.push(gap_sample);
                allocations.open.push(open_sample);
                allocations.combined_prep_gap_open.push(combined_sample);
                processes
                    .prep
                    .push(process_observation(before_prep, after_prep));
                processes
                    .observer_gap
                    .push(process_observation(after_prep, before_open));
                processes
                    .open
                    .push(process_observation(before_open, after_open));
                processes
                    .combined_prep_gap_open
                    .push(process_observation(before_prep, after_open));
            }
        }
    }
    finish_report(
        config,
        corpus,
        None,
        operation,
        wall,
        preparation,
        gap,
        open,
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        Vec::new(),
        allocations,
        processes,
        "not_applicable",
    )
}

#[allow(clippy::too_many_arguments)]
fn finish_report(
    config: &Config,
    corpus: &super::Corpus,
    oracle: Option<&Oracle>,
    mut operation: Vec<u64>,
    mut wall: Vec<u64>,
    mut preparation: Vec<u64>,
    mut gap: Vec<u64>,
    mut open: Vec<u64>,
    mut query: Vec<u64>,
    mut edit: Vec<u64>,
    mut commit: Vec<u64>,
    mut publication: Vec<u64>,
    mut correctness: Vec<bool>,
    mut inverse: Vec<bool>,
    mut output_sha256: Vec<Option<String>>,
    mut correctness_gates: Vec<CorrectnessGates>,
    mut sink: Vec<SinkObservation>,
    mut ingress_calls: Vec<Option<u64>>,
    mut ingress_bytes: Vec<Option<u64>>,
    mut input_bytes: Vec<u64>,
    mut allocations: AllocationPhases,
    mut processes: ProcessPhases,
    correctness_status: &'static str,
) -> Result<Report, Box<dyn Error>> {
    let operation_statistics = super::statistics(std::mem::take(&mut operation));
    let wall_statistics = super::statistics(std::mem::take(&mut wall));
    let sample_order = operation_statistics.sample_order.clone();
    if wall_statistics.sample_order.len() != sample_order.len() {
        return Err("attribution elapsed vectors have different sample counts".into());
    }
    let require_len = |name: &str, length: usize| -> Result<(), Box<dyn Error>> {
        if length == config.samples {
            Ok(())
        } else {
            Err(format!(
                "attribution {name} vector has length {length}, expected {}",
                config.samples
            )
            .into())
        }
    };
    let require_empty = |name: &str, length: usize| -> Result<(), Box<dyn Error>> {
        if length == 0 {
            Ok(())
        } else {
            Err(format!("calibration emitted workload vector {name}").into())
        }
    };
    require_len("preparation_ns", preparation.len())?;
    require_len("observer_gap_ns", gap.len())?;
    require_len("open_ns", open.len())?;
    match config.case {
        Case::Workload => {
            for (name, length) in [
                ("query_ns", query.len()),
                ("edit_ns", edit.len()),
                ("commit_ns", commit.len()),
                ("publication_ns", publication.len()),
                ("correctness_verified", correctness.len()),
                ("inverse_source_verified", inverse.len()),
                ("output_sha256", output_sha256.len()),
                ("correctness_gates", correctness_gates.len()),
                ("sink", sink.len()),
                ("owned_ingress_read_calls", ingress_calls.len()),
                ("owned_ingress_read_bytes", ingress_bytes.len()),
                ("owned_input_bytes", input_bytes.len()),
            ] {
                require_len(name, length)?;
            }
        },
        Case::Calibration => {
            for (name, length) in [
                ("query_ns", query.len()),
                ("edit_ns", edit.len()),
                ("commit_ns", commit.len()),
                ("publication_ns", publication.len()),
                ("correctness_verified", correctness.len()),
                ("inverse_source_verified", inverse.len()),
                ("output_sha256", output_sha256.len()),
                ("correctness_gates", correctness_gates.len()),
                ("sink", sink.len()),
                ("owned_ingress_read_calls", ingress_calls.len()),
                ("owned_ingress_read_bytes", ingress_bytes.len()),
                ("owned_input_bytes", input_bytes.len()),
            ] {
                require_empty(name, length)?;
            }
        },
    }
    reorder_if_present(&mut preparation, &sample_order)?;
    reorder_if_present(&mut gap, &sample_order)?;
    reorder_if_present(&mut open, &sample_order)?;
    reorder_if_present(&mut query, &sample_order)?;
    reorder_if_present(&mut edit, &sample_order)?;
    reorder_if_present(&mut commit, &sample_order)?;
    reorder_if_present(&mut publication, &sample_order)?;
    reorder_if_present(&mut correctness, &sample_order)?;
    reorder_if_present(&mut inverse, &sample_order)?;
    reorder_if_present(&mut output_sha256, &sample_order)?;
    reorder_if_present(&mut correctness_gates, &sample_order)?;
    reorder_if_present(&mut sink, &sample_order)?;
    reorder_if_present(&mut ingress_calls, &sample_order)?;
    reorder_if_present(&mut ingress_bytes, &sample_order)?;
    reorder_if_present(&mut input_bytes, &sample_order)?;
    if config.diagnostic.enabled() {
        require_len("allocation.prep", allocations.prep.len())?;
        require_len("process.prep", processes.prep.len())?;
        reorder(&mut allocations.prep, &sample_order)?;
        reorder(&mut allocations.observer_gap, &sample_order)?;
        reorder(&mut allocations.open, &sample_order)?;
        reorder(&mut allocations.combined_prep_gap_open, &sample_order)?;
        reorder(&mut processes.prep, &sample_order)?;
        reorder(&mut processes.observer_gap, &sample_order)?;
        reorder(&mut processes.open, &sample_order)?;
        reorder(&mut processes.combined_prep_gap_open, &sample_order)?;
    } else if !allocations.prep.is_empty()
        || !allocations.observer_gap.is_empty()
        || !allocations.open.is_empty()
        || !allocations.combined_prep_gap_open.is_empty()
        || !processes.prep.is_empty()
        || !processes.observer_gap.is_empty()
        || !processes.open.is_empty()
        || !processes.combined_prep_gap_open.is_empty()
    {
        return Err("diagnostic-off attribution emitted instrumentation vectors".into());
    }
    let source_bytes = u64::try_from(corpus.archive.len())?;
    let route = config.route.map_or(ROUTE_CALIBRATION, Route::name);
    Ok(Report {
        schema: SCHEMA,
        scenario: "docx-owned-ingress-attribution",
        case: config.case.name(),
        route,
        diagnostic_enabled: config.diagnostic.enabled(),
        diagnostic_arm: diagnostic_arm(config.diagnostic.enabled()),
        binary_identity: allocation_metrics::binary_identity(),
        allocator_identity: allocation_metrics::allocator_identity(),
        instrumentation_identity: allocation_metrics::instrumentation_identity(),
        allocator_counter_revision: allocation_metrics::counter_revision(),
        source_revision: config.source_revision.clone(),
        source_corpus_generator: super::DOCX_SECTION_LAYOUT_CORPUS_GENERATOR,
        source_bytes,
        source_sha256: super::sha256_hex(&corpus.archive),
        expected_output_bytes: oracle
            .map(|value| u64::try_from(value.output.len()).unwrap_or(u64::MAX)),
        expected_output_sha256: oracle.map(|value| value.output_sha256.clone()),
        warmup_iterations: config.warmups,
        sample_count: config.samples,
        timing_scope: "operation_elapsed_ns includes prep, open, and workload phases; wall_elapsed_ns is the measured outer bracket from before-prep observation through the selected body end (publication for workload, post-open observation for calibration)",
        observer_gap_scope: "observer_gap_ns starts after prep's first Region::split and ends after the second split; allocation observer_gap contains both procfs reads, allocation combined ends before after-open observation, process combined includes after-open observation, and wall includes the full outer bracket",
        operation_elapsed_ns: operation_statistics,
        wall_elapsed_ns: wall_statistics,
        sample_order,
        preparation_ns: preparation,
        observer_gap_ns: gap,
        open_ns: open,
        query_ns: query,
        edit_ns: edit,
        commit_ns: commit,
        publication_ns: publication,
        correctness_status,
        correctness_verified: correctness,
        inverse_source_verified: inverse,
        output_sha256,
        correctness_gates,
        sink,
        owned_ingress_read_calls: ingress_calls,
        owned_ingress_read_bytes: ingress_bytes,
        owned_input_bytes: input_bytes,
        allocation: config.diagnostic.enabled().then_some(allocations),
        process: config.diagnostic.enabled().then_some(processes),
    })
}

fn run_workload(
    config: &Config,
    corpus: &super::Corpus,
    route: Route,
) -> Result<Report, Box<dyn Error>> {
    validate_corpus(corpus)?;
    let oracle = prepare_oracle(corpus)?;
    let mut operation = Vec::with_capacity(config.samples);
    let mut wall = Vec::with_capacity(config.samples);
    let mut preparation = Vec::with_capacity(config.samples);
    let mut gap = Vec::with_capacity(config.samples);
    let mut open = Vec::with_capacity(config.samples);
    let mut query = Vec::with_capacity(config.samples);
    let mut edit_ns_values = Vec::with_capacity(config.samples);
    let mut commit_ns_values = Vec::with_capacity(config.samples);
    let mut publication_ns_values = Vec::with_capacity(config.samples);
    let mut correctness = Vec::with_capacity(config.samples);
    let mut inverse = Vec::with_capacity(config.samples);
    let mut output_sha256 = Vec::with_capacity(config.samples);
    let mut correctness_gates = Vec::with_capacity(config.samples);
    let mut sink_observations = Vec::with_capacity(config.samples);
    let mut ingress_calls = Vec::with_capacity(config.samples);
    let mut ingress_bytes = Vec::with_capacity(config.samples);
    let mut input_bytes = Vec::with_capacity(config.samples);
    let mut allocations = AllocationPhases {
        prep: Vec::with_capacity(config.samples),
        observer_gap: Vec::with_capacity(config.samples),
        open: Vec::with_capacity(config.samples),
        combined_prep_gap_open: Vec::with_capacity(config.samples),
    };
    let mut processes = ProcessPhases {
        prep: Vec::with_capacity(config.samples),
        observer_gap: Vec::with_capacity(config.samples),
        open: Vec::with_capacity(config.samples),
        combined_prep_gap_open: Vec::with_capacity(config.samples),
    };
    let payload_marker = super::DOCX_SECTION_LAYOUT_SECTION_POSITIONS[0];
    let mut cache_observation: Option<(u64, u64, u64, usize, usize)> = None;
    for iteration in 0..super::iteration_count(config.warmups, config.samples)? {
        let wall_started = Instant::now();
        let before_prep = read_process(config.diagnostic.enabled());
        let mut region = config.diagnostic.enabled().then(allocation_metrics::begin);
        let preparation_started = Instant::now();
        let input = corpus.archive.clone();
        let read_calls = (route == Route::OwnedReader).then(|| Arc::new(AtomicU64::new(0)));
        let read_bytes = (route == Route::OwnedReader).then(|| Arc::new(AtomicU64::new(0)));
        let mut sink = crate::CountingSink::bounded(oracle.maximum_sink_bytes, 64 * 1024);
        sink.reserve_budget()?;
        let preparation_ns = super::elapsed_ns(preparation_started.elapsed())?;
        let prep_sample = region
            .as_mut()
            .and_then(allocation_metrics::Region::split)
            .unwrap_or_else(allocation_metrics::unavailable_sample);
        let observer_gap_started = Instant::now();
        let after_prep = read_process(config.diagnostic.enabled());
        let before_open = read_process(config.diagnostic.enabled());
        let gap_sample = region
            .as_mut()
            .and_then(allocation_metrics::Region::split)
            .unwrap_or_else(allocation_metrics::unavailable_sample);
        let observer_gap_ns = super::elapsed_ns(observer_gap_started.elapsed())?;
        let open_started = Instant::now();
        let package = open_owned_package(route, input, read_calls.clone(), read_bytes.clone())?;
        black_box(&package);
        let open_ns = super::elapsed_ns(open_started.elapsed())?;
        let (open_sample, combined_sample) = region
            .map(allocation_metrics::Region::finish_split)
            .unwrap_or((None, None));
        let after_open = read_process(config.diagnostic.enabled());
        let open_sample = allocation_sample(open_sample);
        let combined_sample = allocation_sample(combined_sample);

        let query_started = Instant::now();
        let snapshot = package.section_layout_snapshot()?;
        black_box(&snapshot);
        let query_ns = super::elapsed_ns(query_started.elapsed())?;
        let edit_started = Instant::now();
        let mut layout_edit = package.edit_section_layout(
            litchi_docx::section::Selector::paragraph(Position::new(payload_marker)),
        )?;
        layout_edit.set_columns(Some(litchi_docx::section::Columns {
            equal_width: true,
            count: 3,
            space: Some(litchi_docx::section::Emu::from_twips(240)),
            separator: true,
            columns: Vec::new(),
        }))?;
        black_box(&layout_edit);
        let edit_ns = super::elapsed_ns(edit_started.elapsed())?;
        let commit_started = Instant::now();
        let layout_commit = layout_edit.commit()?;
        black_box(&layout_commit);
        let commit_ns = super::elapsed_ns(commit_started.elapsed())?;
        let before_publication_after_commit = package.cache_diagnostics();
        let publication_started = Instant::now();
        let layout_publication = package.apply_section_layout(layout_commit.patch())?;
        layout_publication.package().write_to(&mut sink)?;
        black_box(&layout_publication);
        let publication_ns = super::elapsed_ns(publication_started.elapsed())?;
        let operation_ns = preparation_ns
            .checked_add(open_ns)
            .and_then(|value| value.checked_add(query_ns))
            .and_then(|value| value.checked_add(edit_ns))
            .and_then(|value| value.checked_add(commit_ns))
            .and_then(|value| value.checked_add(publication_ns))
            .ok_or("DOCX owned attribution operation timing overflowed")?;
        let wall_ns = super::elapsed_ns(wall_started.elapsed())?;

        super::verify_docx_section_layout_snapshot(&snapshot, [2, 2, 2])?;
        let changed_nonnoop_columns_verified = !layout_commit.patch().is_noop()
            && layout_publication
                .package()
                .section_layout_snapshot()?
                .inventory()
                .sections()[0]
                .columns()
                .map(|columns| columns.count)
                == Some(3);
        if !changed_nonnoop_columns_verified {
            return Err("DOCX owned attribution commit changed an unexpected section".into());
        }
        let mut raw_tree = snapshot.raw_tree_at(
            litchi_docx::section::Selector::paragraph(Position::new(payload_marker)),
            &litchi_docx::section::Limits::default(),
        )?;
        raw_tree.set_columns(Some(litchi_docx::section::Columns {
            equal_width: true,
            count: 3,
            space: Some(litchi_docx::section::Emu::from_twips(240)),
            separator: true,
            columns: Vec::new(),
        }))?;
        let raw_patch = raw_tree.patch()?;
        let replayed = raw_patch.apply(&snapshot)?;
        let restored_snapshot = raw_patch.inverse().apply(&replayed)?;
        let patch_forward_inverse_stale_verified = restored_snapshot.inventory().sections()
            == snapshot.inventory().sections()
            && layout_commit
                .package()
                .apply_section_layout(layout_commit.patch())
                .is_err();
        if !patch_forward_inverse_stale_verified {
            return Err("DOCX owned attribution patch oracle failed".into());
        }
        let publication_inverse_applicable_verified = layout_publication
            .package()
            .apply_section_layout(&layout_publication.patch().inverse())
            .is_ok();
        if !publication_inverse_applicable_verified {
            return Err("DOCX owned attribution publication inverse is not applicable".into());
        }
        let output_digest = super::sha256_hex(&sink.bytes);
        let output_verified = sink.bytes == oracle.output;
        if !output_verified {
            return Err("DOCX owned attribution publication differs from oracle".into());
        }
        super::verify_docx_section_layout_output(corpus, &sink.bytes, [3, 2, 2])?;
        let reopened = litchi_docx::source_backed::Package::from_read_at(Arc::new(
            OwnedSource::new(sink.bytes.clone()),
        ))?;
        let mut restored = Vec::new();
        reopened
            .apply_section_layout(&layout_publication.patch().inverse())?
            .package()
            .write_to(&mut restored)?;
        let inverse_verified = restored == corpus.archive;
        if !inverse_verified {
            return Err("DOCX owned attribution inverse did not restore source".into());
        }
        let sink_summary = sink.summary();
        let sink_bounds_verified = sink_summary.accepted_bytes
            == u64::try_from(oracle.output.len())?
            && sink_summary.largest_write <= 64 * 1024;
        if !sink_bounds_verified {
            return Err("DOCX owned attribution sink bounds differ from oracle".into());
        }
        let read_calls_value = read_calls
            .as_ref()
            .map(|value| value.load(Ordering::Relaxed));
        let read_bytes_value = read_bytes
            .as_ref()
            .map(|value| value.load(Ordering::Relaxed));
        let owned_reader_ingress_verified = route != Route::OwnedReader
            || (read_calls_value.unwrap_or(0) > 0 && read_bytes_value.unwrap_or(0) > 0);
        if !owned_reader_ingress_verified {
            return Err("DOCX owned-reader attribution performed no sequential reads".into());
        }
        let cache_contract_verified = before_publication_after_commit.successful_loads == 1
            && before_publication_after_commit.cold_loads == 1
            && before_publication_after_commit.retained_entries == 1
            && before_publication_after_commit.retained_bytes > 0;
        if !cache_contract_verified {
            return Err("DOCX owned attribution source cache contract failed".into());
        }
        let observed_cache = (
            before_publication_after_commit.hits,
            before_publication_after_commit.cold_loads,
            before_publication_after_commit.successful_loads,
            before_publication_after_commit.retained_entries,
            before_publication_after_commit.retained_bytes,
        );
        let cache_deterministic_verified =
            cache_observation.is_none_or(|expected| expected == observed_cache);
        if !cache_deterministic_verified {
            return Err("DOCX owned attribution cache counters are not deterministic".into());
        }
        cache_observation = Some(observed_cache);
        if iteration >= config.warmups {
            operation.push(operation_ns);
            wall.push(wall_ns);
            preparation.push(preparation_ns);
            gap.push(observer_gap_ns);
            open.push(open_ns);
            query.push(query_ns);
            edit_ns_values.push(edit_ns);
            commit_ns_values.push(commit_ns);
            publication_ns_values.push(publication_ns);
            correctness.push(true);
            inverse.push(true);
            output_sha256.push(Some(output_digest));
            correctness_gates.push(CorrectnessGates {
                snapshot_shape_verified: true,
                changed_nonnoop_columns_verified,
                patch_forward_inverse_stale_verified,
                publication_inverse_applicable_verified,
                output_verified,
                reopened_output_verified: true,
                inverse_source_verified: inverse_verified,
                source_immutability_verified: inverse_verified,
                cache_contract_verified,
                cache_deterministic_verified,
                sink_bounds_verified,
                owned_reader_ingress_verified,
                cache_hits: observed_cache.0,
                cache_cold_loads: observed_cache.1,
                cache_successful_loads: observed_cache.2,
                cache_retained_entries: observed_cache.3,
                cache_retained_bytes: observed_cache.4,
            });
            sink_observations.push(SinkObservation {
                accepted_bytes: sink_summary.accepted_bytes,
                write_calls: sink_summary.write_calls,
                largest_write: sink_summary.largest_write,
            });
            ingress_calls.push(read_calls_value);
            ingress_bytes.push(read_bytes_value);
            input_bytes.push(u64::try_from(corpus.archive.len())?);
            if config.diagnostic.enabled() {
                allocations.prep.push(prep_sample);
                allocations.observer_gap.push(gap_sample);
                allocations.open.push(open_sample);
                allocations.combined_prep_gap_open.push(combined_sample);
                processes
                    .prep
                    .push(process_observation(before_prep, after_prep));
                processes
                    .observer_gap
                    .push(process_observation(after_prep, before_open));
                processes
                    .open
                    .push(process_observation(before_open, after_open));
                processes
                    .combined_prep_gap_open
                    .push(process_observation(before_prep, after_open));
            }
        }
    }
    finish_report(
        config,
        corpus,
        Some(&oracle),
        operation,
        wall,
        preparation,
        gap,
        open,
        query,
        edit_ns_values,
        commit_ns_values,
        publication_ns_values,
        correctness,
        inverse,
        output_sha256,
        correctness_gates,
        sink_observations,
        ingress_calls,
        ingress_bytes,
        input_bytes,
        allocations,
        processes,
        "verified",
    )
}

fn validate_corpus(corpus: &super::Corpus) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != super::DOCX_SECTION_LAYOUT_CORPUS_GENERATOR
        || corpus.manifest.archive_sha256 != super::sha256_hex(&corpus.archive)
    {
        return Err("DOCX owned attribution corpus identity is invalid".into());
    }
    Ok(())
}

/// Run the explicit owned-ingress attribution case.
pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    let args = args.into_iter().collect::<Vec<_>>();
    if args.iter().any(|arg| arg == "--help" || arg == "-h") {
        println!(
            "docx-owned-attribution --case <workload|calibration> --diagnostic <off|on> [--route <owned_vec|owned_reader>] --warmups N --samples N [--source-revision HEX40] [--output PATH]"
        );
        return Ok(());
    }
    let config = parse_config(&args)?;
    let corpus = super::build_docx_section_layout_corpus()?;
    validate_corpus(&corpus)?;
    let report = match config.case {
        Case::Calibration => run_calibration(&config, &corpus)?,
        Case::Workload => run_workload(
            &config,
            &corpus,
            config.route.ok_or("workload route is missing")?,
        )?,
    };
    write_report(&config, &report)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn args(values: &[&str]) -> Vec<OsString> {
        values.iter().map(OsString::from).collect()
    }

    #[test]
    fn parser_requires_explicit_case_route_and_revision() {
        let parsed = parse_config(&args(&[
            "--case",
            "workload",
            "--route",
            "owned_vec",
            "--diagnostic",
            "off",
            "--warmups",
            "5",
            "--samples",
            "50",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
        ]))
        .expect("valid workload arguments");
        assert_eq!(parsed.case, Case::Workload);
        assert_eq!(parsed.route, Some(Route::OwnedVec));
        assert_eq!(parsed.diagnostic, Diagnostic::Off);
        assert_eq!(parsed.samples, 50);
        assert!(
            parse_config(&args(&[
                "--case",
                "workload",
                "--diagnostic",
                "off",
                "--warmups",
                "0",
                "--samples",
                "1",
                "--source-revision",
                "0123456789abcdef0123456789abcdef01234567",
            ]))
            .is_err()
        );
        assert!(
            parse_config(&args(&[
                "--case",
                "workload",
                "--route",
                "owned_vec",
                "--diagnostic",
                "off",
                "--warmups",
                "0",
                "--samples",
                "1",
            ]))
            .is_err()
        );
    }

    #[test]
    fn calibration_is_route_independent_and_diagnostic_only() {
        let revision = "0123456789abcdef0123456789abcdef01234567";
        let valid = parse_config(&args(&[
            "--case",
            "calibration",
            "--diagnostic",
            "on",
            "--warmups",
            "5",
            "--samples",
            "50",
            "--source-revision",
            revision,
        ]))
        .expect("valid calibration arguments");
        assert_eq!(valid.route, None);
        assert_eq!(valid.case, Case::Calibration);
        assert!(
            parse_config(&args(&[
                "--case",
                "calibration",
                "--route",
                "owned_vec",
                "--diagnostic",
                "on",
                "--warmups",
                "0",
                "--samples",
                "1",
                "--source-revision",
                revision,
            ]))
            .is_err()
        );
        assert!(
            parse_config(&args(&[
                "--case",
                "calibration",
                "--diagnostic",
                "off",
                "--warmups",
                "0",
                "--samples",
                "1",
                "--source-revision",
                revision,
            ]))
            .is_err()
        );
    }

    #[test]
    fn process_observation_exposes_status_and_signed_rss() {
        let before = process_metrics::Snapshot {
            rss_bytes: 4096,
            ..process_metrics::Snapshot::default()
        };
        let after = process_metrics::Snapshot {
            rss_bytes: 2048,
            ..process_metrics::Snapshot::default()
        };
        let observation = process_observation(
            ProcessRead {
                snapshot: Some(before),
                status: "measured",
            },
            ProcessRead {
                snapshot: Some(after),
                status: "measured",
            },
        );
        assert_eq!(observation.before_status, "measured");
        assert_eq!(observation.after_status, "measured");
        assert_eq!(observation.rss_delta_signed, Some(-2048));
        assert_eq!(observation.rss_delta_status, "measured");
    }

    #[test]
    fn process_observation_preserves_unavailable_and_overflow_status() {
        let unavailable = process_observation(
            ProcessRead {
                snapshot: None,
                status: "procfs_invalid_data",
            },
            ProcessRead {
                snapshot: None,
                status: "procfs_not_found",
            },
        );
        assert_eq!(unavailable.before_status, "procfs_invalid_data");
        assert_eq!(unavailable.after_status, "procfs_not_found");
        assert_eq!(unavailable.rss_delta_signed, None);
        assert_eq!(unavailable.rss_delta_status, "unavailable");

        let overflow = process_observation(
            ProcessRead {
                snapshot: Some(process_metrics::Snapshot::default()),
                status: "measured",
            },
            ProcessRead {
                snapshot: Some(process_metrics::Snapshot {
                    rss_bytes: u64::MAX,
                    ..process_metrics::Snapshot::default()
                }),
                status: "measured",
            },
        );
        assert_eq!(overflow.rss_delta_signed, None);
        assert_eq!(overflow.rss_delta_status, "overflow");
    }
}
