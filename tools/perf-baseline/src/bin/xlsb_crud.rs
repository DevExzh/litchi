//! Opt-in XLSB semantic CRUD baseline.
//!
//! This binary intentionally lives outside the default perf matrix. It uses
//! only the public XLSB owner and umbrella-facade APIs, and keeps the complete
//! reopen/preservation checks outside the timed sample interval. The default
//! corpus is a public POI workbook with opaque members and a VBA payload; the
//! payload is treated as inert and is checked as an untouched package member.

use litchi::common::FileFormat;
use litchi::sheet::{WorkbookTrait, Worksheet};
use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_opc::SourceCacheDiagnostics;
use litchi_xlsb::SourceBackedWorkbook;
use litchi_xlsb::cell_values::{CellError, Reference, StoredCell, Value};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::env;
use std::fmt;
use std::fs;
use std::io::{self, Cursor};
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};
use std::time::Instant;

type Error = Box<dyn std::error::Error + Send + Sync>;
type Result<T> = std::result::Result<T, Error>;

const CORPUS_VERSION: &str = "litchi-xlsb-semantic-crud-poi-v1";
const DEFAULT_FIXTURE: &str = "test-data/poi/test-data/spreadsheet/testVarious.xlsb";
const DEFAULT_WARMUP: usize = 3;
const DEFAULT_SAMPLES: usize = 30;
#[cfg(test)]
const TEST_VARIUS_SHA256: &str = "8c600e97d719b0266dcfb49c1872feb8d10c6ed12bc768ff16ace7dae555ebfc";

/// One opt-in semantic operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
enum Case {
    OpenIdentify,
    WorksheetCatalog,
    SelectedWorksheetCell,
    FullStoredCellScan,
    FullText,
    NoopTransactionCommitSave,
    EditOneExistingScalarSave,
    EditCeilOnePercentExistingCellsSave,
}

/// Input implementation exercised by the harness. The source-backed mode is
/// intentionally limited to read-only open/catalog/materialization cases so
/// its positional reads and deferred Part loads can be compared directly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
enum Backend {
    Owned,
    OwnedDirect,
    OwnedWithoutDrawings,
    SourceBacked,
}

impl FromStr for Backend {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        match value {
            "owned" => Ok(Self::Owned),
            "owned_direct" => Ok(Self::OwnedDirect),
            "owned_without_drawings" => Ok(Self::OwnedWithoutDrawings),
            "source_backed" => Ok(Self::SourceBacked),
            _ => Err(format!(
                "unknown XLSB backend {value:?}; expected owned, owned_direct, owned_without_drawings, or source_backed"
            )
            .into()),
        }
    }
}

impl Case {
    const ALL: [Self; 8] = [
        Self::OpenIdentify,
        Self::WorksheetCatalog,
        Self::SelectedWorksheetCell,
        Self::FullStoredCellScan,
        Self::FullText,
        Self::NoopTransactionCommitSave,
        Self::EditOneExistingScalarSave,
        Self::EditCeilOnePercentExistingCellsSave,
    ];

    const fn as_str(self) -> &'static str {
        match self {
            Self::OpenIdentify => "open_identify",
            Self::WorksheetCatalog => "worksheet_catalog",
            Self::SelectedWorksheetCell => "selected_worksheet_cell",
            Self::FullStoredCellScan => "full_stored_cell_scan",
            Self::FullText => "full_text",
            Self::NoopTransactionCommitSave => "noop_transaction_commit_save",
            Self::EditOneExistingScalarSave => "edit_one_existing_scalar_save",
            Self::EditCeilOnePercentExistingCellsSave => {
                "edit_ceil_one_percent_existing_cells_save"
            },
        }
    }

    fn parse_case(value: &str) -> Result<Vec<Self>> {
        if value == "all" {
            return Ok(Self::ALL.to_vec());
        }
        Ok(vec![Self::from_str(value)?])
    }
}

impl FromStr for Case {
    type Err = Error;

    fn from_str(value: &str) -> Result<Self> {
        match value {
            "open_identify" => Ok(Self::OpenIdentify),
            "worksheet_catalog" => Ok(Self::WorksheetCatalog),
            "selected_worksheet_cell" => Ok(Self::SelectedWorksheetCell),
            "full_stored_cell_scan" => Ok(Self::FullStoredCellScan),
            "full_text" => Ok(Self::FullText),
            "noop_transaction_commit_save" => Ok(Self::NoopTransactionCommitSave),
            "edit_one_existing_scalar_save" => Ok(Self::EditOneExistingScalarSave),
            "edit_ceil_one_percent_existing_cells_save" => {
                Ok(Self::EditCeilOnePercentExistingCellsSave)
            },
            _ => Err(format!("unknown XLSB CRUD case {value:?}").into()),
        }
    }
}

#[derive(Debug, Clone)]
struct Args {
    backend: Backend,
    cases: Vec<Case>,
    warmup: usize,
    samples: usize,
    fixture: PathBuf,
    json: Option<PathBuf>,
}

impl Args {
    fn parse<I>(arguments: I) -> Result<Self>
    where
        I: IntoIterator<Item = String>,
    {
        let mut cases = None;
        let mut backend = Backend::Owned;
        let mut warmup = DEFAULT_WARMUP;
        let mut samples = DEFAULT_SAMPLES;
        let mut fixture = PathBuf::from(DEFAULT_FIXTURE);
        let mut json = None;
        let mut arguments = arguments.into_iter();
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--backend" => {
                    backend = Backend::from_str(&next_arg("--backend", &mut arguments)?)?
                },
                "--case" => cases = Some(Case::parse_case(&next_arg("--case", &mut arguments)?)?),
                "--warmup" => {
                    warmup = parse_positive("--warmup", &next_arg("--warmup", &mut arguments)?)?
                },
                "--samples" => {
                    samples = parse_positive("--samples", &next_arg("--samples", &mut arguments)?)?
                },
                "--fixture" => fixture = PathBuf::from(next_arg("--fixture", &mut arguments)?),
                "--json" => json = Some(PathBuf::from(next_arg("--json", &mut arguments)?)),
                "--help" | "-h" => return Err(usage().into()),
                other => return Err(format!("unknown argument {other:?}\n\n{}", usage()).into()),
            }
        }
        let cases = cases.ok_or_else(|| format!("--case is required\n\n{}", usage()))?;
        for case in &cases {
            ensure_case_supported(backend, *case)?;
        }
        Ok(Self {
            backend,
            cases,
            warmup,
            samples,
            fixture,
            json,
        })
    }
}

fn next_arg<I>(name: &str, arguments: &mut I) -> Result<String>
where
    I: Iterator<Item = String>,
{
    arguments
        .next()
        .ok_or_else(|| format!("missing value for {name}").into())
}

fn parse_positive(name: &str, value: &str) -> Result<usize> {
    let value = value
        .parse::<usize>()
        .map_err(|error| format!("{name} must be a positive integer: {error}"))?;
    if value == 0 {
        return Err(format!("{name} must be greater than zero").into());
    }
    Ok(value)
}

fn usage() -> &'static str {
    "usage: xlsb_crud --case <all|open_identify|worksheet_catalog|selected_worksheet_cell|full_stored_cell_scan|full_text|noop_transaction_commit_save|edit_one_existing_scalar_save|edit_ceil_one_percent_existing_cells_save> [--backend owned|owned_direct|owned_without_drawings|source_backed] [--fixture PATH] [--warmup N] [--samples N] [--json PATH]"
}

#[derive(Debug, Clone)]
struct Corpus {
    path: PathBuf,
    bytes: Vec<u8>,
    source_sha256: String,
    worksheet_names: Vec<String>,
    selected_sheet: usize,
    selected_sheet_name: String,
    selected_reference: Reference,
    selected_value: Value,
    selected_semantic_value: litchi_core::sheet::CellValue,
    stored_cells: Vec<StoredCell>,
    semantic_cells: Vec<(u32, u32, litchi_core::sheet::CellValue)>,
    selected_coordinate: String,
    edits: Vec<EditTarget>,
    stored_cell_count: usize,
    coordinates: Vec<String>,
    full_text_sha256: String,
    full_text_bytes: usize,
    part_digests: BTreeMap<String, String>,
}

#[derive(Debug, Clone)]
struct EditTarget {
    reference: Reference,
    coordinate: String,
    after: Value,
}

#[derive(Debug, Clone, Serialize)]
struct CorpusReport {
    generator: &'static str,
    fixture: String,
    source_sha256: String,
    input_bytes: usize,
    worksheet_count: usize,
    worksheet_names: Vec<String>,
    selected_sheet: usize,
    selected_sheet_name: String,
    selected_coordinate: String,
    selected_editable_count: usize,
    ceil_one_percent_edit_count: usize,
    stored_cell_count: usize,
    stored_cell_coordinates: Vec<String>,
    full_text_sha256: String,
    full_text_bytes: usize,
    package_part_count: usize,
}

#[derive(Debug, Clone, Serialize)]
struct Statistics {
    warmup: usize,
    samples: usize,
    samples_ns: Vec<u64>,
    p50_ns: u64,
    mean_ns: f64,
    p95_ns: u64,
    p99_ns: u64,
}

#[derive(Debug, Clone, Serialize)]
struct GateReport {
    representative_output_reopen_ok: Option<bool>,
    semantic_readback_ok: bool,
    exact_noop_patch: Option<bool>,
    output_matches_across_samples: Option<bool>,
    unchanged_parts_ok: Option<bool>,
    changed_part_names: Vec<String>,
    malformed_input_refused: bool,
    tight_limits_refused: bool,
    tight_cell_limits_refused: bool,
    fixture_cell_count_within_dimensions: bool,
}

#[derive(Debug, Clone, Serialize)]
struct CaseReport {
    backend: Backend,
    case: Case,
    timing_scope: &'static str,
    input_bytes: usize,
    output_bytes: Option<usize>,
    output_sha256: Option<String>,
    observed_count: Option<usize>,
    observed_text_bytes: Option<usize>,
    observed_text_sha256: Option<String>,
    selected_coordinates: Vec<String>,
    source_observation: Option<SourceObservation>,
    source_observation_stable: Option<bool>,
    statistics: Statistics,
    gates: GateReport,
}

/// Per-operation evidence from the source-backed XLSB path.
///
/// The open interval includes catalog construction. The operation interval
/// begins after construction and covers the selected catalog/materialization
/// action. `part_materializations` is the checked OPC cache cold-load count;
/// it is the direct proof of deferred Part payload work in this harness.
#[derive(Debug, Clone, Copy, Serialize, PartialEq, Eq)]
struct SourceObservation {
    source_read_calls: u64,
    source_read_requested_bytes: u64,
    source_read_bytes: u64,
    open_source_read_calls: u64,
    open_source_read_requested_bytes: u64,
    open_source_read_bytes: u64,
    operation_source_read_calls: u64,
    operation_source_read_requested_bytes: u64,
    operation_source_read_bytes: u64,
    part_materializations: u64,
    open_part_materializations: u64,
    operation_part_materializations: u64,
    part_cache_hits: u64,
    open_part_cache_hits: u64,
    operation_part_cache_hits: u64,
    retained_part_entries: usize,
    retained_part_bytes: usize,
}

#[derive(Debug, Clone, Copy)]
struct SourceReadSnapshot {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

#[derive(Debug, Default)]
struct SourceReadCounters {
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
}

impl SourceReadCounters {
    fn snapshot(&self) -> SourceReadSnapshot {
        SourceReadSnapshot {
            calls: self.calls.load(Ordering::Relaxed),
            requested_bytes: self.requested_bytes.load(Ordering::Relaxed),
            returned_bytes: self.returned_bytes.load(Ordering::Relaxed),
        }
    }
}

struct CountingReadAt {
    inner: OwnedSource,
    counters: Arc<SourceReadCounters>,
}

impl CountingReadAt {
    fn new(bytes: Vec<u8>, counters: Arc<SourceReadCounters>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            counters,
        }
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.calls.fetch_add(1, Ordering::Relaxed);
        self.counters.requested_bytes.fetch_add(
            u64::try_from(output.len()).unwrap_or(u64::MAX),
            Ordering::Relaxed,
        );
        let read = self.inner.read_at(offset, output)?;
        self.counters
            .returned_bytes
            .fetch_add(u64::try_from(read).unwrap_or(u64::MAX), Ordering::Relaxed);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    generator: &'static str,
    binary_identity: litchi_perf_baseline::BinaryIdentity,
    corpus: CorpusReport,
    cases: Vec<CaseReport>,
}

#[derive(Debug)]
enum RunOutcome {
    Projection { count: usize, text: Option<String> },
    Saved(Vec<u8>),
}

fn main() -> Result<()> {
    let args = match Args::parse(env::args().skip(1)) {
        Ok(args) => args,
        Err(error) if error.to_string().starts_with("usage:") => {
            println!("{error}");
            return Ok(());
        },
        Err(error) => return Err(error),
    };
    let corpus = Corpus::load(&args.fixture)?;
    let mut reports = Vec::with_capacity(args.cases.len());
    for case in args.cases {
        reports.push(benchmark_case(
            &corpus,
            args.backend,
            case,
            args.warmup,
            args.samples,
        )?);
    }
    let binary_identity = litchi_perf_baseline::current_executable_identity()?;
    let report = Report {
        schema: "xlsb-crud-v2",
        generator: CORPUS_VERSION,
        binary_identity,
        corpus: corpus.report(),
        cases: reports,
    };
    let json = serde_json::to_string_pretty(&report)?;
    if let Some(path) = args.json {
        fs::write(path, json.as_bytes())?;
    } else {
        println!("{json}");
    }
    Ok(())
}

impl Corpus {
    fn load(path: &Path) -> Result<Self> {
        let bytes = fs::read(path)?;
        if bytes.is_empty() {
            return Err("XLSB fixture is empty".into());
        }
        let source_sha256 = sha256_hex(&bytes);
        let workbook = open_direct(&bytes)?;
        let worksheet_names = workbook.worksheet_names().to_vec();
        if worksheet_names.is_empty() {
            return Err("XLSB fixture has no worksheets".into());
        }
        let mut selected = None;
        for sheet in 0..workbook.worksheet_count() {
            let snapshot = workbook.cell_values(sheet)?;
            let cells: Vec<StoredCell> = snapshot.cells().cloned().collect();
            if selected.is_none()
                && cells
                    .iter()
                    .any(|cell| replacement_for(cell, &workbook).is_some())
            {
                selected = Some((sheet, cells));
            }
        }
        let (selected_sheet, cells) = selected.ok_or_else(|| {
            "the deterministic XLSB fixture has no publicly editable scalar cell".to_string()
        })?;
        let selected_cell = cells
            .iter()
            .find(|cell| replacement_for(cell, &workbook).is_some())
            .ok_or_else(|| "selected XLSB worksheet has no editable scalar cell".to_string())?;
        let mut edits = Vec::new();
        let mut coordinates = Vec::new();
        for cell in &cells {
            coordinates.push(coordinate(cell.reference()));
            if let Some(after) = replacement_for(cell, &workbook) {
                edits.push(EditTarget {
                    reference: cell.reference(),
                    coordinate: coordinate(cell.reference()),
                    after,
                });
            }
        }
        let ceil_count = cells.len().div_ceil(100);
        if edits.len() < ceil_count {
            return Err(format!(
                "fixture supports only {} of {} required ceil(1%) scalar edits",
                edits.len(),
                ceil_count
            )
            .into());
        }
        let selected_semantic_value = workbook
            .worksheet(selected_sheet)?
            .cell(
                selected_cell.reference().row(),
                selected_cell.reference().column(),
            )?
            .value()
            .clone();
        let worksheet = workbook.worksheet(selected_sheet)?;
        let mut semantic_cells = Vec::new();
        let mut semantic_iterator = worksheet.cells();
        while let Some(cell) = semantic_iterator.next() {
            let cell = cell?;
            semantic_cells.push((cell.row(), cell.column(), cell.value().clone()));
        }
        let full_text = facade_text(&bytes)?;
        let package = litchi_xlsb::Package::from_slice(&bytes)?;
        let part_digests = package_part_digests(package.opc_package());
        Ok(Self {
            path: path.to_path_buf(),
            bytes,
            source_sha256,
            worksheet_names: worksheet_names.clone(),
            selected_sheet,
            selected_sheet_name: worksheet_names[selected_sheet].clone(),
            selected_reference: selected_cell.reference(),
            selected_value: selected_cell.value().clone(),
            selected_semantic_value,
            selected_coordinate: coordinate(selected_cell.reference()),
            edits,
            stored_cell_count: cells.len(),
            coordinates,
            full_text_sha256: sha256_hex(full_text.as_bytes()),
            full_text_bytes: full_text.len(),
            part_digests,
            stored_cells: cells,
            semantic_cells,
        })
    }

    fn report(&self) -> CorpusReport {
        CorpusReport {
            generator: CORPUS_VERSION,
            fixture: self.path.display().to_string(),
            source_sha256: self.source_sha256.clone(),
            input_bytes: self.bytes.len(),
            worksheet_count: self.worksheet_names.len(),
            worksheet_names: self.worksheet_names.clone(),
            selected_sheet: self.selected_sheet,
            selected_sheet_name: self.selected_sheet_name.clone(),
            selected_coordinate: self.selected_coordinate.clone(),
            selected_editable_count: self.edits.len(),
            ceil_one_percent_edit_count: self.stored_cell_count.div_ceil(100),
            stored_cell_count: self.stored_cell_count,
            stored_cell_coordinates: self.coordinates.clone(),
            full_text_sha256: self.full_text_sha256.clone(),
            full_text_bytes: self.full_text_bytes,
            package_part_count: self
                .part_digests
                .keys()
                .filter(|name| name.as_str() != PACKAGE_RELATIONSHIP_KEY)
                .count(),
        }
    }
}

fn benchmark_case(
    corpus: &Corpus,
    backend: Backend,
    case: Case,
    warmup: usize,
    samples: usize,
) -> Result<CaseReport> {
    ensure_case_supported(backend, case)?;
    for _ in 0..warmup {
        let (outcome, _) = run_case_with_backend(corpus, backend, case)?;
        std::hint::black_box(outcome);
    }
    let mut elapsed = Vec::with_capacity(samples);
    let mut output_identities = Vec::with_capacity(samples);
    let mut sample_source_observations = Vec::with_capacity(samples);
    for _ in 0..samples {
        let started = Instant::now();
        let (outcome, source_observation) = run_case_with_backend(corpus, backend, case)?;
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let outcome = std::hint::black_box(outcome);
        if let RunOutcome::Saved(bytes) = outcome {
            output_identities.push((bytes.len(), sha256_hex(&bytes)));
        }
        if let Some(source_observation) = source_observation {
            sample_source_observations.push(source_observation);
        }
        elapsed.push(elapsed_ns);
    }
    let (representative, source_observation) = run_case_with_backend(corpus, backend, case)?;
    let (output, observed_count, observed_text) = match representative {
        RunOutcome::Saved(bytes) => (Some(bytes), None, None),
        RunOutcome::Projection { count, text } => (None, Some(count), text),
    };
    let output_sha256 = output.as_deref().map(sha256_hex);
    let output_bytes = output.as_ref().map(Vec::len);
    let observed_text_sha256 = observed_text
        .as_deref()
        .map(|text| sha256_hex(text.as_bytes()));
    let observed_text_bytes = observed_text.as_ref().map(String::len);
    let source_observation_stable = if backend == Backend::SourceBacked {
        Some(
            sample_source_observations
                .iter()
                .all(|sample| Some(*sample) == source_observation),
        )
    } else {
        None
    };
    if source_observation_stable == Some(false) {
        return Err(
            format!("{case}: source ReadAt/cache observation changed across samples").into(),
        );
    }
    let gate = match backend {
        Backend::Owned => verify_case(
            corpus,
            case,
            open_direct,
            output.as_deref(),
            &output_identities,
            observed_count,
            observed_text_bytes,
            observed_text_sha256.as_deref(),
        )?,
        Backend::OwnedDirect => verify_case(
            corpus,
            case,
            open_direct,
            output.as_deref(),
            &output_identities,
            observed_count,
            observed_text_bytes,
            observed_text_sha256.as_deref(),
        )?,
        Backend::OwnedWithoutDrawings => verify_case(
            corpus,
            case,
            open_direct_without_drawings,
            output.as_deref(),
            &output_identities,
            observed_count,
            observed_text_bytes,
            observed_text_sha256.as_deref(),
        )?,
        Backend::SourceBacked => verify_source_case(
            corpus,
            case,
            observed_count,
            observed_text_bytes,
            observed_text_sha256.as_deref(),
        )?,
    };
    ensure_gates(backend, case, &gate)?;
    Ok(CaseReport {
        backend,
        case,
        timing_scope: timing_scope(backend, case),
        input_bytes: corpus.bytes.len(),
        output_bytes,
        output_sha256,
        observed_count,
        observed_text_bytes,
        observed_text_sha256,
        selected_coordinates: selected_coordinates(case, corpus),
        source_observation,
        source_observation_stable,
        statistics: Statistics::from_samples(warmup, elapsed),
        gates: gate,
    })
}

fn ensure_gates(backend: Backend, case: Case, gate: &GateReport) -> Result<()> {
    if gate.representative_output_reopen_ok == Some(false) {
        return Err(format!("{case}: representative reopen gate failed").into());
    }
    if !gate.semantic_readback_ok {
        return Err(format!("{case}: semantic readback/projection gate failed").into());
    }
    if case == Case::NoopTransactionCommitSave && gate.exact_noop_patch != Some(true) {
        return Err(format!("{case}: exact no-op patch identity gate failed").into());
    }
    if gate.output_matches_across_samples == Some(false) {
        return Err(format!("{case}: output identity was not deterministic across samples").into());
    }
    if gate.unchanged_parts_ok == Some(false) {
        return Err(format!("{case}: untouched package-member gate failed").into());
    }
    if !gate.malformed_input_refused {
        return Err(format!("{case}: malformed-input refusal gate failed").into());
    }
    if !gate.tight_limits_refused {
        return Err(format!("{case}: tight read-limit refusal gate failed").into());
    }
    if matches!(
        backend,
        Backend::Owned | Backend::OwnedDirect | Backend::OwnedWithoutDrawings
    ) && !gate.tight_cell_limits_refused
    {
        return Err(format!("{case}: tight cell-limit refusal gate failed").into());
    }
    if !gate.fixture_cell_count_within_dimensions {
        return Err(format!("{case}: fixture cell-count/dimensions gate failed").into());
    }
    Ok(())
}

fn ensure_case_supported(backend: Backend, case: Case) -> Result<()> {
    if backend == Backend::SourceBacked
        && !matches!(
            case,
            Case::OpenIdentify
                | Case::WorksheetCatalog
                | Case::SelectedWorksheetCell
                | Case::FullStoredCellScan
        )
    {
        return Err(format!(
            "source_backed backend does not support XLSB case {case}; choose one of open_identify, worksheet_catalog, selected_worksheet_cell, or full_stored_cell_scan"
        )
        .into());
    }
    if case == Case::FullText && backend != Backend::Owned {
        return Err(format!(
            "{backend:?} backend does not support XLSB case {case}; full_text is a facade-only workload"
        )
        .into());
    }
    Ok(())
}

fn run_case_with_backend(
    corpus: &Corpus,
    backend: Backend,
    case: Case,
) -> Result<(RunOutcome, Option<SourceObservation>)> {
    ensure_case_supported(backend, case)?;
    match backend {
        Backend::Owned => Ok((run_case(corpus, case)?, None)),
        Backend::OwnedDirect => Ok((
            run_case_with_opener(corpus, case, open_direct, false)?,
            None,
        )),
        Backend::OwnedWithoutDrawings => Ok((
            run_case_with_opener(corpus, case, open_direct_without_drawings, false)?,
            None,
        )),
        Backend::SourceBacked => {
            let (outcome, observation) = run_source_case(corpus, case)?;
            Ok((outcome, Some(observation)))
        },
    }
}

fn run_source_case(corpus: &Corpus, case: Case) -> Result<(RunOutcome, SourceObservation)> {
    let counters = Arc::new(SourceReadCounters::default());
    let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
        corpus.bytes.clone(),
        Arc::clone(&counters),
    ));
    let workbook = SourceBackedWorkbook::from_read_at(source)?;
    let open_reads = counters.snapshot();
    let open_diagnostics = workbook.cache_diagnostics();
    let (outcome, _) = match case {
        Case::OpenIdentify => (
            RunOutcome::Projection {
                count: workbook.worksheet_count()?,
                text: None,
            },
            (),
        ),
        Case::WorksheetCatalog => (
            RunOutcome::Projection {
                count: workbook.worksheet_count()?,
                text: Some(workbook.worksheet_names()?.join("\u{1f}")),
            },
            (),
        ),
        Case::SelectedWorksheetCell => {
            let worksheet = workbook
                .worksheet_by_index(corpus.selected_sheet)?
                .ok_or_else(|| "selected source-backed XLSB worksheet disappeared".to_string())?
                .materialize()?;
            let cell = worksheet.cell(
                corpus.selected_reference.row(),
                corpus.selected_reference.column(),
            )?;
            if cell.row() != corpus.selected_reference.row()
                || cell.column() != corpus.selected_reference.column()
                || cell.value() != &corpus.selected_semantic_value
            {
                return Err("selected source-backed XLSB cell changed during lookup".into());
            }
            std::hint::black_box(cell);
            (
                RunOutcome::Projection {
                    count: 1,
                    text: None,
                },
                (),
            )
        },
        Case::FullStoredCellScan => {
            let worksheet = workbook
                .worksheet_by_index(corpus.selected_sheet)?
                .ok_or_else(|| "selected source-backed XLSB worksheet disappeared".to_string())?
                .materialize()?;
            let mut count = 0usize;
            let mut cells = worksheet.cells();
            while let Some(cell) = cells.next() {
                let cell = cell?;
                count = count.saturating_add(1);
                std::hint::black_box(cell);
            }
            (RunOutcome::Projection { count, text: None }, ())
        },
        _ => return Err(format!("source-backed backend does not support XLSB case {case}").into()),
    };
    let end_reads = counters.snapshot();
    let end_diagnostics = workbook.cache_diagnostics();
    Ok((
        outcome,
        SourceObservation::from_snapshots(
            open_reads,
            end_reads,
            open_diagnostics,
            end_diagnostics,
        )?,
    ))
}

impl SourceObservation {
    fn from_snapshots(
        open_reads: SourceReadSnapshot,
        end_reads: SourceReadSnapshot,
        open_diagnostics: SourceCacheDiagnostics,
        end_diagnostics: SourceCacheDiagnostics,
    ) -> Result<Self> {
        let operation_reads = subtract_reads(end_reads, open_reads)?;
        let total_diagnostics =
            source_cache_delta(SourceCacheDiagnostics::default(), end_diagnostics)?;
        let operation_diagnostics = source_cache_delta(open_diagnostics, end_diagnostics)?;
        Ok(Self {
            source_read_calls: end_reads.calls,
            source_read_requested_bytes: end_reads.requested_bytes,
            source_read_bytes: end_reads.returned_bytes,
            open_source_read_calls: open_reads.calls,
            open_source_read_requested_bytes: open_reads.requested_bytes,
            open_source_read_bytes: open_reads.returned_bytes,
            operation_source_read_calls: operation_reads.calls,
            operation_source_read_requested_bytes: operation_reads.requested_bytes,
            operation_source_read_bytes: operation_reads.returned_bytes,
            part_materializations: total_diagnostics.cold_loads,
            open_part_materializations: open_diagnostics.cold_loads,
            operation_part_materializations: operation_diagnostics.cold_loads,
            part_cache_hits: total_diagnostics.hits,
            open_part_cache_hits: open_diagnostics.hits,
            operation_part_cache_hits: operation_diagnostics.hits,
            retained_part_entries: end_diagnostics.retained_entries,
            retained_part_bytes: end_diagnostics.retained_bytes,
        })
    }
}

fn subtract_reads(
    after: SourceReadSnapshot,
    before: SourceReadSnapshot,
) -> Result<SourceReadSnapshot> {
    Ok(SourceReadSnapshot {
        calls: after
            .calls
            .checked_sub(before.calls)
            .ok_or_else(|| "source ReadAt call counter moved backwards".to_string())?,
        requested_bytes: after
            .requested_bytes
            .checked_sub(before.requested_bytes)
            .ok_or_else(|| "source ReadAt requested-byte counter moved backwards".to_string())?,
        returned_bytes: after
            .returned_bytes
            .checked_sub(before.returned_bytes)
            .ok_or_else(|| "source ReadAt returned-byte counter moved backwards".to_string())?,
    })
}

fn source_cache_delta(
    before: SourceCacheDiagnostics,
    after: SourceCacheDiagnostics,
) -> Result<litchi_opc::SourceCacheCounterDelta> {
    SourceCacheDiagnostics::checked_counter_delta(before, after).map_err(|error| {
        format!("source cache diagnostic counter interval invalid: {error}").into()
    })
}

fn run_case(corpus: &Corpus, case: Case) -> Result<RunOutcome> {
    run_case_with_opener(corpus, case, open_direct, true)
}

fn run_case_with_opener(
    corpus: &Corpus,
    case: Case,
    opener: fn(&[u8]) -> Result<litchi_xlsb::Workbook>,
    identify_facade: bool,
) -> Result<RunOutcome> {
    match case {
        Case::OpenIdentify => {
            if identify_facade {
                let format = litchi::detect_file_format_from_bytes(&corpus.bytes)
                    .ok_or_else(|| "facade could not identify the XLSB fixture".to_string())?;
                if format != FileFormat::Xlsb {
                    return Err(format!("facade identified fixture as {format:?}, not XLSB").into());
                }
                let workbook = litchi::sheet::open_xlsb_workbook_from_bytes(&corpus.bytes)?;
                return Ok(RunOutcome::Projection {
                    count: workbook.worksheet_count(),
                    text: None,
                });
            }
            let workbook = opener(&corpus.bytes)?;
            Ok(RunOutcome::Projection {
                count: workbook.worksheet_count(),
                text: None,
            })
        },
        Case::WorksheetCatalog => {
            let workbook = opener(&corpus.bytes)?;
            Ok(RunOutcome::Projection {
                count: workbook.worksheet_count(),
                text: Some(workbook.worksheet_names().join("\u{1f}")),
            })
        },
        Case::SelectedWorksheetCell => {
            let workbook = opener(&corpus.bytes)?;
            let snapshot = workbook.cell_values(corpus.selected_sheet)?;
            let cell = snapshot
                .cell(corpus.selected_reference)?
                .ok_or_else(|| "selected XLSB cell disappeared".to_string())?;
            if cell.value() != &corpus.selected_value {
                return Err("selected XLSB cell value changed during lookup".into());
            }
            Ok(RunOutcome::Projection {
                count: if cell.reference() == corpus.selected_reference {
                    1
                } else {
                    0
                },
                text: None,
            })
        },
        Case::FullStoredCellScan => {
            let workbook = opener(&corpus.bytes)?;
            let snapshot = workbook.cell_values(corpus.selected_sheet)?;
            let mut count = 0usize;
            for cell in snapshot.cells() {
                count = count.saturating_add(1);
                std::hint::black_box((cell.reference(), cell.style(), cell.value()));
            }
            Ok(RunOutcome::Projection { count, text: None })
        },
        Case::FullText => Ok(RunOutcome::Projection {
            count: 0,
            text: Some(facade_text(&corpus.bytes)?),
        }),
        Case::NoopTransactionCommitSave => {
            let mut workbook = opener(&corpus.bytes)?;
            let snapshot = workbook.cell_values(corpus.selected_sheet)?;
            let commit = snapshot.edit().commit()?;
            if !commit.patch().is_empty() {
                return Err("public exact no-op commit unexpectedly changed bytes".into());
            }
            workbook.apply_cell_values(corpus.selected_sheet, &commit)?;
            Ok(RunOutcome::Saved(save_workbook(&workbook)?))
        },
        Case::EditOneExistingScalarSave => {
            let mut workbook = opener(&corpus.bytes)?;
            let target = corpus
                .edits
                .first()
                .ok_or_else(|| "no editable scalar target".to_string())?;
            let mut edit = workbook.edit_cell_values(corpus.selected_sheet)?;
            edit.set_value(target.reference, target.after.clone())?;
            let commit = edit.commit()?;
            workbook.apply_cell_values(corpus.selected_sheet, &commit)?;
            Ok(RunOutcome::Saved(save_workbook(&workbook)?))
        },
        Case::EditCeilOnePercentExistingCellsSave => {
            let mut workbook = opener(&corpus.bytes)?;
            let count = corpus.stored_cell_count.div_ceil(100);
            let mut edit = workbook.edit_cell_values(corpus.selected_sheet)?;
            for target in corpus.edits.iter().take(count) {
                edit.set_value(target.reference, target.after.clone())?;
            }
            let commit = edit.commit()?;
            workbook.apply_cell_values(corpus.selected_sheet, &commit)?;
            Ok(RunOutcome::Saved(save_workbook(&workbook)?))
        },
    }
}

fn verify_case(
    corpus: &Corpus,
    case: Case,
    opener: fn(&[u8]) -> Result<litchi_xlsb::Workbook>,
    output: Option<&[u8]>,
    sample_identities: &[(usize, String)],
    observed_count: Option<usize>,
    observed_text_bytes: Option<usize>,
    observed_text_sha256: Option<&str>,
) -> Result<GateReport> {
    let malformed_input_refused = litchi_xlsb::Package::from_bytes(vec![0, 1, 2, 3]).is_err();
    let tight_limits_refused = {
        let limit = litchi_xlsb::ReadLimits::builder()
            .max_input_bytes(u64::try_from(corpus.bytes.len().saturating_sub(1))?)?
            .build()?;
        litchi_xlsb::Package::from_bytes_with_limits(corpus.bytes.clone(), limit).is_err()
    };
    let tight_cell_limits_refused = {
        let workbook = opener(&corpus.bytes)?;
        let limits = litchi_xlsb::cell_values::Limits::new(1, 1, 1, 1);
        workbook
            .cell_values_with_limits(corpus.selected_sheet, limits)
            .is_err()
    };
    let fixture_cell_count_within_dimensions = sparse_gate(corpus)?;
    let catalog_text_sha256 = sha256_hex(corpus.worksheet_names.join("\u{1f}").as_bytes());
    let semantic_projection_ok = match case {
        Case::OpenIdentify => observed_count == Some(corpus.worksheet_names.len()),
        Case::WorksheetCatalog => {
            observed_count == Some(corpus.worksheet_names.len())
                && observed_text_bytes == Some(corpus.worksheet_names.join("\u{1f}").len())
                && observed_text_sha256 == Some(catalog_text_sha256.as_str())
        },
        Case::SelectedWorksheetCell => observed_count == Some(1),
        Case::FullStoredCellScan => {
            observed_count == Some(corpus.stored_cell_count)
                && opener(&corpus.bytes)?
                    .cell_values(corpus.selected_sheet)?
                    .cells()
                    .eq(corpus.stored_cells.iter())
        },
        Case::FullText => {
            observed_text_bytes == Some(corpus.full_text_bytes)
                && observed_text_sha256 == Some(corpus.full_text_sha256.as_str())
        },
        Case::NoopTransactionCommitSave
        | Case::EditOneExistingScalarSave
        | Case::EditCeilOnePercentExistingCellsSave => true,
    };
    let Some(output) = output else {
        let report = GateReport {
            representative_output_reopen_ok: None,
            semantic_readback_ok: semantic_projection_ok,
            exact_noop_patch: None,
            output_matches_across_samples: None,
            unchanged_parts_ok: None,
            changed_part_names: Vec::new(),
            malformed_input_refused,
            tight_limits_refused,
            tight_cell_limits_refused,
            fixture_cell_count_within_dimensions,
        };
        return Ok(report);
    };
    let reopened = open_direct(output)?;
    let representative_output_reopen_ok =
        reopened.worksheet_names() == corpus.worksheet_names && semantic_projection_ok;
    let output_identity = (output.len(), sha256_hex(output));
    let output_matches_across_samples = sample_identities
        .iter()
        .all(|identity| identity == &output_identity);
    let output_package = litchi_xlsb::Package::from_slice(output)?;
    let output_parts = package_part_digests(output_package.opc_package());
    let changed_part_names: Vec<String> = corpus
        .part_digests
        .iter()
        .filter_map(|(name, digest)| match output_parts.get(name) {
            Some(output_digest) if output_digest == digest => None,
            Some(_) | None => Some(name.clone()),
        })
        .chain(
            output_parts
                .keys()
                .filter(|name| !corpus.part_digests.contains_key(*name))
                .cloned(),
        )
        .collect();
    let allowed_changed = match case {
        Case::NoopTransactionCommitSave => changed_part_names.is_empty(),
        Case::EditOneExistingScalarSave | Case::EditCeilOnePercentExistingCellsSave => {
            changed_part_names.len() == 1 && changed_part_names[0].contains("/worksheets/")
        },
        Case::OpenIdentify
        | Case::WorksheetCatalog
        | Case::SelectedWorksheetCell
        | Case::FullStoredCellScan
        | Case::FullText => false,
    };
    let (semantic_readback_ok, exact_noop_patch) = match case {
        Case::NoopTransactionCommitSave => {
            let snapshot = reopened.cell_values(corpus.selected_sheet)?;
            let source = corpus.bytes_for_selected_sheet()?;
            (
                snapshot.source_bytes() == source,
                Some(snapshot.source_bytes() == source),
            )
        },
        Case::EditOneExistingScalarSave => {
            let target = corpus
                .edits
                .first()
                .ok_or_else(|| "missing edit target".to_string())?;
            let snapshot = reopened.cell_values(corpus.selected_sheet)?;
            let actual = snapshot
                .cell(target.reference)?
                .ok_or_else(|| "edited target is missing after reopen".to_string())?;
            (actual.value() == &target.after, None)
        },
        Case::EditCeilOnePercentExistingCellsSave => {
            let count = corpus.stored_cell_count.div_ceil(100);
            let snapshot = reopened.cell_values(corpus.selected_sheet)?;
            let mut ok = true;
            for target in corpus.edits.iter().take(count) {
                let actual = snapshot
                    .cell(target.reference)?
                    .ok_or_else(|| "edited target is missing after reopen".to_string())?;
                ok &= actual.value() == &target.after;
            }
            (ok, None)
        },
        Case::OpenIdentify
        | Case::WorksheetCatalog
        | Case::SelectedWorksheetCell
        | Case::FullStoredCellScan
        | Case::FullText => (representative_output_reopen_ok, None),
    };
    let report = GateReport {
        representative_output_reopen_ok: Some(representative_output_reopen_ok),
        semantic_readback_ok,
        exact_noop_patch,
        output_matches_across_samples: Some(output_matches_across_samples),
        unchanged_parts_ok: Some(allowed_changed),
        changed_part_names,
        malformed_input_refused,
        tight_limits_refused,
        tight_cell_limits_refused,
        fixture_cell_count_within_dimensions,
    };
    Ok(report)
}

// Complete semantic replay is deliberately outside the timed scan.
fn source_scan_matches(corpus: &Corpus) -> Result<bool> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.bytes.clone()));
    let workbook = SourceBackedWorkbook::from_read_at(source)?;
    let worksheet = workbook
        .worksheet_by_index(corpus.selected_sheet)?
        .ok_or("selected source-backed worksheet disappeared")?
        .materialize()?;
    let mut expected = corpus.semantic_cells.iter();
    let mut cells = worksheet.cells();
    while let Some(cell) = cells.next() {
        let cell = cell?;
        let Some((row, column, value)) = expected.next() else {
            return Ok(false);
        };
        if cell.row() != *row || cell.column() != *column || cell.value() != value {
            return Ok(false);
        }
    }
    Ok(expected.next().is_none())
}

fn verify_source_case(
    corpus: &Corpus,
    case: Case,
    observed_count: Option<usize>,
    observed_text_bytes: Option<usize>,
    observed_text_sha256: Option<&str>,
) -> Result<GateReport> {
    let catalog_text_sha256 = sha256_hex(corpus.worksheet_names.join("\u{1f}").as_bytes());
    let semantic_readback_ok = match case {
        Case::OpenIdentify => observed_count == Some(corpus.worksheet_names.len()),
        Case::WorksheetCatalog => {
            observed_count == Some(corpus.worksheet_names.len())
                && observed_text_bytes == Some(corpus.worksheet_names.join("\u{1f}").len())
                && observed_text_sha256 == Some(catalog_text_sha256.as_str())
        },
        Case::SelectedWorksheetCell => observed_count == Some(1),
        Case::FullStoredCellScan => {
            observed_count == Some(corpus.stored_cell_count) && source_scan_matches(corpus)?
        },
        _ => false,
    };
    let malformed_input_refused = {
        let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
            vec![0, 1, 2, 3],
            Arc::new(SourceReadCounters::default()),
        ));
        SourceBackedWorkbook::from_read_at(source).is_err()
    };
    let tight_limits_refused = {
        let limit = litchi_xlsb::ReadLimits::builder()
            .max_input_bytes(u64::try_from(corpus.bytes.len().saturating_sub(1))?)?
            .build()?;
        let source: Arc<dyn ReadAt> = Arc::new(CountingReadAt::new(
            corpus.bytes.clone(),
            Arc::new(SourceReadCounters::default()),
        ));
        SourceBackedWorkbook::from_read_at_with_limits(source, limit).is_err()
    };
    // SourceBackedWorksheet::materialize currently exposes the format-level
    // read limits but no separate cell-value limit parameter. Keep this gate
    // explicitly inapplicable rather than claiming a cell-limit refusal.
    let tight_cell_limits_refused = false;
    Ok(GateReport {
        representative_output_reopen_ok: None,
        semantic_readback_ok,
        exact_noop_patch: None,
        output_matches_across_samples: None,
        unchanged_parts_ok: None,
        changed_part_names: Vec::new(),
        malformed_input_refused,
        tight_limits_refused,
        tight_cell_limits_refused,
        fixture_cell_count_within_dimensions: sparse_gate(corpus)?,
    })
}

impl Corpus {
    fn bytes_for_selected_sheet(&self) -> Result<Vec<u8>> {
        let workbook = open_direct(&self.bytes)?;
        Ok(workbook
            .cell_values(self.selected_sheet)?
            .source_bytes()
            .to_vec())
    }
}

fn sparse_gate(corpus: &Corpus) -> Result<bool> {
    let workbook = open_direct(&corpus.bytes)?;
    let worksheet = workbook.worksheet(corpus.selected_sheet)?;
    let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() else {
        return Ok(true);
    };
    let area = u64::from(max_row.saturating_sub(min_row).saturating_add(1))
        .saturating_mul(u64::from(max_col.saturating_sub(min_col).saturating_add(1)));
    let area = usize::try_from(area).unwrap_or(usize::MAX);
    if corpus.stored_cell_count >= area {
        // A dense worksheet is not a sparse-expansion probe. Treat an exact
        // rectangle as valid after the full-scan gate has checked the stored
        // count; an overfull projection remains a refusal.
        return Ok(corpus.stored_cell_count == area);
    }
    Ok(true)
}

fn selected_coordinates(case: Case, corpus: &Corpus) -> Vec<String> {
    match case {
        Case::SelectedWorksheetCell | Case::EditOneExistingScalarSave => {
            vec![corpus.selected_coordinate.clone()]
        },
        Case::EditCeilOnePercentExistingCellsSave => corpus
            .edits
            .iter()
            .take(corpus.stored_cell_count.div_ceil(100))
            .map(|target| target.coordinate.clone())
            .collect(),
        Case::OpenIdentify
        | Case::WorksheetCatalog
        | Case::FullStoredCellScan
        | Case::FullText
        | Case::NoopTransactionCommitSave => Vec::new(),
    }
}

fn timing_scope(backend: Backend, case: Case) -> &'static str {
    if backend == Backend::SourceBacked {
        return match case {
            Case::OpenIdentify => {
                "source-backed XLSB catalog open plus worksheet count (facade detection excluded)"
            },
            Case::WorksheetCatalog => {
                "source-backed XLSB catalog open plus worksheet count and names"
            },
            Case::SelectedWorksheetCell => {
                "source-backed XLSB catalog open plus selected worksheet materialization and one cell lookup"
            },
            Case::FullStoredCellScan => {
                "source-backed XLSB catalog open plus complete stored-cell scan on the selected worksheet"
            },
            _ => "unsupported source-backed XLSB case",
        };
    }
    if backend == Backend::OwnedWithoutDrawings {
        return match case {
            Case::OpenIdentify => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus worksheet count"
            },
            Case::WorksheetCatalog => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus worksheet count and names"
            },
            Case::SelectedWorksheetCell => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus selected-worksheet snapshot materialization and one cell lookup"
            },
            Case::FullStoredCellScan => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus complete source-bound stored-cell scan"
            },
            Case::FullText => "unsupported: full_text is facade-only; use the owned backend",
            Case::NoopTransactionCommitSave => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus exact no-op transaction, commit, publication, and save"
            },
            Case::EditOneExistingScalarSave => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus one existing scalar edit, commit, publication, and save"
            },
            Case::EditCeilOnePercentExistingCellsSave => {
                "owned XLSB cell/catalog projection open with worksheet drawing parse skipped plus deterministic ceil(1%) existing scalar edits on the selected worksheet, commit, publication, and save"
            },
        };
    }
    if backend == Backend::OwnedDirect {
        return match case {
            Case::OpenIdentify => "direct eager XLSB open plus worksheet count",
            Case::WorksheetCatalog => "direct eager XLSB open plus worksheet count and names",
            Case::SelectedWorksheetCell => {
                "direct eager XLSB open plus selected-worksheet snapshot materialization and one cell lookup"
            },
            Case::FullStoredCellScan => {
                "direct eager XLSB open plus complete source-bound stored-cell scan on the selected worksheet"
            },
            Case::FullText => "unsupported: full_text is facade-only; use the owned backend",
            Case::NoopTransactionCommitSave => {
                "direct eager XLSB open plus exact no-op transaction, commit, publication, and save"
            },
            Case::EditOneExistingScalarSave => {
                "direct eager XLSB open plus one existing scalar edit, commit, publication, and save"
            },
            Case::EditCeilOnePercentExistingCellsSave => {
                "direct eager XLSB open plus deterministic ceil(1%) existing scalar edits on the selected worksheet, commit, publication, and save"
            },
        };
    }
    match case {
        Case::OpenIdentify => "facade format identification plus XLSB facade open",
        Case::WorksheetCatalog => "direct XLSB open plus worksheet count and names",
        Case::SelectedWorksheetCell => {
            "direct XLSB open plus selected-worksheet snapshot materialization and one cell lookup"
        },
        Case::FullStoredCellScan => {
            "direct XLSB open plus complete source-bound stored-cell scan on the selected worksheet"
        },
        Case::FullText => "facade bytes open plus all-worksheet text extraction",
        Case::NoopTransactionCommitSave => {
            "direct XLSB open plus exact no-op transaction, commit, publication, and save"
        },
        Case::EditOneExistingScalarSave => {
            "direct XLSB open plus one existing scalar edit, commit, publication, and save"
        },
        Case::EditCeilOnePercentExistingCellsSave => {
            "direct XLSB open plus deterministic ceil(1%) existing scalar edits on the selected worksheet, commit, publication, and save"
        },
    }
}

fn open_direct(bytes: &[u8]) -> Result<litchi_xlsb::Workbook> {
    Ok(litchi_xlsb::Workbook::new(Cursor::new(bytes.to_vec()))?)
}

fn open_direct_without_drawings(bytes: &[u8]) -> Result<litchi_xlsb::Workbook> {
    Ok(litchi_xlsb::Workbook::new_without_drawing_parse(
        Cursor::new(bytes.to_vec()),
    )?)
}

fn facade_text(bytes: &[u8]) -> Result<String> {
    litchi::sheet::Workbook::from_bytes(bytes.to_vec())?.text()
}

fn save_workbook(workbook: &litchi_xlsb::Workbook) -> Result<Vec<u8>> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output)?;
    Ok(output.into_inner())
}

fn replacement_for(cell: &StoredCell, workbook: &litchi_xlsb::Workbook) -> Option<Value> {
    match cell.value() {
        Value::Number(value) => Some(Value::Number(if value.to_bits() == 1.0f64.to_bits() {
            2.0
        } else {
            1.0
        })),
        Value::RkNumber(value) => Some(Value::RkNumber(if value.to_bits() == 1.0f64.to_bits() {
            2.0
        } else {
            1.0
        })),
        Value::Boolean(value) => Some(Value::Boolean(!value)),
        Value::Error(error) => Some(Value::Error(alternate_error(*error))),
        Value::InlineString(value) => Some(Value::InlineString(alternate_string(value))),
        Value::SharedStringIndex(index) => {
            let count = workbook.shared_strings().len();
            if count > 1 {
                let index = usize::try_from(*index).ok()?;
                Some(Value::SharedStringIndex(
                    u32::try_from((index + 1) % count).ok()?,
                ))
            } else {
                None
            }
        },
        Value::FormulaNumberCache(value) => Some(Value::FormulaNumberCache(
            if value.to_bits() == 1.0f64.to_bits() {
                2.0
            } else {
                1.0
            },
        )),
        Value::FormulaBooleanCache(value) => Some(Value::FormulaBooleanCache(!value)),
        Value::FormulaErrorCache(error) => Some(Value::FormulaErrorCache(alternate_error(*error))),
        Value::FormulaStringCache(value) => {
            Some(Value::FormulaStringCache(alternate_string(value)))
        },
        Value::Blank | Value::RichString(_) => None,
        _ => None,
    }
}

fn alternate_string(value: &str) -> String {
    let units = value.encode_utf16().count();
    if units == 0 {
        "X".to_string()
    } else if value == "X".repeat(units) {
        "Y".repeat(units)
    } else {
        "X".repeat(units)
    }
}

const fn alternate_error(error: CellError) -> CellError {
    if matches!(error, CellError::Reference) {
        CellError::Value
    } else {
        CellError::Reference
    }
}

fn coordinate(reference: Reference) -> String {
    let mut column = reference.column().saturating_add(1);
    let mut letters = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        letters.push(char::from(b'A' + u8::try_from(remainder).unwrap_or(0)));
        column = (column - 1) / 26;
    }
    let letters: String = letters.chars().rev().collect();
    format!("{letters}{}", reference.row().saturating_add(1))
}

const PACKAGE_RELATIONSHIP_KEY: &str = "<package>";

fn package_part_digests(package: &litchi_opc::OpcPackage) -> BTreeMap<String, String> {
    let mut digests = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .map(|part| {
            let mut hasher = Sha256::new();
            hasher.update(part.content_type().as_bytes());
            hasher.update([0]);
            hasher.update(part.blob());
            hasher.update([0]);
            let mut relationships: Vec<String> = part
                .rels()
                .iter()
                .map(|relationship| {
                    format!(
                        "{}\u{1f}{}\u{1f}{}\u{1f}{}",
                        relationship.r_id(),
                        relationship.reltype(),
                        relationship.target_ref(),
                        relationship.is_external()
                    )
                })
                .collect();
            relationships.sort();
            for relationship in relationships {
                hasher.update(relationship.as_bytes());
                hasher.update([0]);
            }
            (
                part.partname().as_str().to_string(),
                hex_bytes(&hasher.finalize()),
            )
        })
        .collect::<BTreeMap<_, _>>();

    let mut package_relationships: Vec<String> = package
        .rels()
        .iter()
        .map(|relationship| {
            format!(
                "{}\u{1f}{}\u{1f}{}\u{1f}{}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.is_external()
            )
        })
        .collect();
    package_relationships.sort();
    let mut package_hasher = Sha256::new();
    package_hasher.update(b"package-relationships");
    package_hasher.update([0]);
    for relationship in package_relationships {
        package_hasher.update(relationship.as_bytes());
        package_hasher.update([0]);
    }
    digests.insert(
        PACKAGE_RELATIONSHIP_KEY.to_owned(),
        hex_bytes(&package_hasher.finalize()),
    );
    digests
}

#[cfg(test)]
fn package_relationships(package: &litchi_opc::OpcPackage) -> BTreeMap<String, Vec<String>> {
    let mut graph: BTreeMap<String, Vec<String>> = package
        .iter_parts()
        .map(|part| {
            let mut relationships: Vec<String> = part
                .rels()
                .iter()
                .map(|relationship| {
                    format!(
                        "{}\u{1f}{}\u{1f}{}\u{1f}{}",
                        relationship.r_id(),
                        relationship.reltype(),
                        relationship.target_ref(),
                        relationship.is_external()
                    )
                })
                .collect();
            relationships.sort();
            (part.partname().as_str().to_string(), relationships)
        })
        .collect();
    let mut relationships: Vec<String> = package
        .rels()
        .iter()
        .map(|relationship| {
            format!(
                "{}\u{1f}{}\u{1f}{}\u{1f}{}",
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
                relationship.is_external()
            )
        })
        .collect();
    relationships.sort();
    graph.insert(PACKAGE_RELATIONSHIP_KEY.to_owned(), relationships);
    graph
}

fn sha256_hex(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hex_bytes(&hasher.finalize())
}

fn hex_bytes(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len().saturating_mul(2));
    for byte in bytes {
        output.push(char::from(HEX[usize::from(byte >> 4)]));
        output.push(char::from(HEX[usize::from(byte & 0x0f)]));
    }
    output
}

impl Statistics {
    fn from_samples(warmup: usize, samples: Vec<u64>) -> Self {
        let mut sorted = samples.clone();
        sorted.sort_unstable();
        let percentile = |percent: usize| {
            let index = sorted
                .len()
                .saturating_mul(percent)
                .saturating_add(99)
                .checked_div(100)
                .unwrap_or(1)
                .saturating_sub(1)
                .min(sorted.len().saturating_sub(1));
            sorted[index]
        };
        let total: u128 = samples.iter().map(|sample| u128::from(*sample)).sum();
        let mean_ns = if samples.is_empty() {
            0.0
        } else {
            total as f64 / samples.len() as f64
        };
        Self {
            warmup,
            samples: samples.len(),
            samples_ns: samples,
            p50_ns: percentile(50),
            mean_ns,
            p95_ns: percentile(95),
            p99_ns: percentile(99),
        }
    }
}

impl fmt::Display for Case {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../")
            .join(DEFAULT_FIXTURE)
    }

    #[test]
    fn parses_every_case_and_all() {
        let all = Case::parse_case("all").expect("all");
        assert_eq!(all, Case::ALL);
        for case in Case::ALL {
            assert_eq!(Case::from_str(case.as_str()).expect("case"), case);
        }
    }

    #[test]
    fn rejects_malformed_case_and_arguments() {
        assert!(Case::from_str("unknown").is_err());
        assert!(Case::parse_case("").is_err());
        assert!(
            Args::parse([
                "--case".to_string(),
                "full_text".to_string(),
                "--samples".to_string(),
                "0".to_string(),
            ])
            .is_err()
        );
        assert!(
            Args::parse([
                "--case".to_string(),
                "full_text".to_string(),
                "--nope".to_string(),
            ])
            .is_err()
        );
        let direct_error = ensure_case_supported(Backend::OwnedWithoutDrawings, Case::FullText)
            .expect_err("drawing-skipped projection must not time the facade-only text workload");
        assert!(direct_error.to_string().contains("facade-only"));
        let source_error =
            ensure_case_supported(Backend::SourceBacked, Case::NoopTransactionCommitSave)
                .expect_err("source-backed workbook has no save/edit workload");
        assert!(source_error.to_string().contains("source_backed backend"));
    }

    #[test]
    fn deterministic_public_fixture_identity() {
        let bytes = fs::read(fixture_path()).expect("public fixture");
        assert_eq!(sha256_hex(&bytes), TEST_VARIUS_SHA256);
        let first = Corpus::load(&fixture_path()).expect("corpus");
        let second = Corpus::load(&fixture_path()).expect("corpus");
        assert_eq!(first.source_sha256, second.source_sha256);
        assert_eq!(first.worksheet_names, second.worksheet_names);
        assert_eq!(first.coordinates, second.coordinates);
    }

    #[test]
    fn malformed_and_tight_limits_are_refused() {
        let corpus = Corpus::load(&fixture_path()).expect("corpus");
        let limit = litchi_xlsb::ReadLimits::builder()
            .max_input_bytes(u64::try_from(corpus.bytes.len().saturating_sub(1)).expect("u64"))
            .expect("limit")
            .build()
            .expect("limits");
        assert!(litchi_xlsb::Package::from_bytes_with_limits(corpus.bytes.clone(), limit).is_err());
        assert!(litchi_xlsb::Package::from_bytes(vec![0, 1, 2, 3]).is_err());
        let workbook = open_direct(&corpus.bytes).expect("workbook");
        let limits = litchi_xlsb::cell_values::Limits::new(1, 1, 1, 1);
        assert!(
            workbook
                .cell_values_with_limits(corpus.selected_sheet, limits)
                .is_err()
        );
    }

    #[test]
    fn exact_noop_identity_and_edit_readback_are_preserved() {
        let corpus = Corpus::load(&fixture_path()).expect("corpus");
        let noop = run_case(&corpus, Case::NoopTransactionCommitSave).expect("noop");
        let RunOutcome::Saved(noop_bytes) = noop else {
            panic!("noop must save")
        };
        let noop_again = run_case(&corpus, Case::NoopTransactionCommitSave)
            .expect("noop output")
            .saved_bytes();
        assert_eq!(sha256_hex(&noop_bytes), sha256_hex(&noop_again));
        let noop_gate = verify_case(
            &corpus,
            Case::NoopTransactionCommitSave,
            open_direct,
            Some(&noop_bytes),
            &[(noop_bytes.len(), sha256_hex(&noop_bytes))],
            None,
            None,
            None,
        )
        .expect("noop gates");
        assert_eq!(noop_gate.exact_noop_patch, Some(true));
        assert_eq!(noop_gate.unchanged_parts_ok, Some(true));
        assert!(noop_gate.tight_cell_limits_refused);
        let mut failing_gate = noop_gate.clone();
        failing_gate.output_matches_across_samples = Some(false);
        assert!(
            ensure_gates(
                Backend::Owned,
                Case::NoopTransactionCommitSave,
                &failing_gate
            )
            .is_err()
        );

        let edited = run_case(&corpus, Case::EditOneExistingScalarSave).expect("edit");
        let RunOutcome::Saved(edited_bytes) = edited else {
            panic!("edit must save")
        };
        let edit_gate = verify_case(
            &corpus,
            Case::EditOneExistingScalarSave,
            open_direct,
            Some(&edited_bytes),
            &[(edited_bytes.len(), sha256_hex(&edited_bytes))],
            None,
            None,
            None,
        )
        .expect("edit gates");
        assert_eq!(edit_gate.representative_output_reopen_ok, Some(true));
        assert!(edit_gate.semantic_readback_ok);
        assert_eq!(edit_gate.unchanged_parts_ok, Some(true));
    }

    #[test]
    fn sparse_iteration_does_not_expand_to_rectangle() {
        let corpus = Corpus::load(&fixture_path()).expect("corpus");
        assert!(sparse_gate(&corpus).expect("sparse gate"));
        let scan = run_case(&corpus, Case::FullStoredCellScan).expect("scan");
        let RunOutcome::Projection { count, .. } = scan else {
            panic!("scan projection")
        };
        assert_eq!(count, corpus.stored_cell_count);
    }

    #[test]
    fn source_backed_measurements_validate_selected_value_and_counter_intervals() {
        let mut corpus = Corpus::load(&fixture_path()).expect("corpus");
        for case in [
            Case::OpenIdentify,
            Case::WorksheetCatalog,
            Case::SelectedWorksheetCell,
            Case::FullStoredCellScan,
        ] {
            let report = benchmark_case(&corpus, Backend::SourceBacked, case, 0, 2)
                .expect("verified source-backed report");
            assert_eq!(report.source_observation_stable, Some(true));
            let reads = report.source_observation.expect("source observation");
            assert_eq!(
                reads.source_read_calls,
                reads.open_source_read_calls + reads.operation_source_read_calls
            );
            assert_eq!(
                reads.source_read_bytes,
                reads.open_source_read_bytes + reads.operation_source_read_bytes
            );
            assert_eq!(
                reads.part_materializations,
                reads.open_part_materializations + reads.operation_part_materializations
            );
        }
        assert_eq!(
            corpus.report().package_part_count,
            litchi_xlsb::Package::from_slice(&corpus.bytes)
                .expect("package")
                .opc_package()
                .iter_parts()
                .count()
        );
        corpus.semantic_cells[0].0 = u32::MAX;
        assert!(
            !source_scan_matches(&corpus).expect("fresh scan"),
            "a matching cell count cannot substitute for full scan semantics"
        );
        corpus.selected_semantic_value = litchi_core::sheet::CellValue::Empty;
        assert!(
            run_source_case(&corpus, Case::SelectedWorksheetCell).is_err(),
            "a matching count cannot substitute for matching selected cell semantics"
        );
    }

    #[test]
    fn drawing_skipped_projection_preserves_raw_drawing_graph_on_noop_save() {
        let path = fixture_path();
        let bytes = fs::read(&path).expect("fixture");
        let corpus = Corpus::load(&path).expect("corpus");
        let eager = open_direct(&bytes).expect("eager workbook");
        assert!(
            !eager.sheet_drawings().is_empty(),
            "fixture must exercise the skipped drawing projection"
        );
        let workbook = open_direct_without_drawings(&bytes).expect("projection workbook");
        assert!(workbook.sheet_drawings().is_empty());
        let output = save_workbook(&workbook).expect("projection save");
        let before = litchi_xlsb::Package::from_slice(&bytes).expect("input package");
        let after = litchi_xlsb::Package::from_slice(&output).expect("output package");
        let before_parts = package_part_digests(before.opc_package());
        assert!(
            before_parts.keys().any(|name| name.contains("/drawings/")),
            "fixture must retain a raw drawing part for the projection test"
        );
        assert!(
            package_relationships(before.opc_package())
                .values()
                .flatten()
                .any(|relationship| relationship.contains("/drawing")),
            "fixture must retain a worksheet-to-drawing relationship"
        );
        assert_eq!(before_parts, package_part_digests(after.opc_package()));

        let mut raw_noop = open_direct_without_drawings(&bytes).expect("raw no-op projection");
        raw_noop
            .edit_opc(|_| Ok(()))
            .expect("raw no-op projection publication");
        assert!(raw_noop.sheet_drawings().is_empty());
        let raw_noop_output = save_workbook(&raw_noop).expect("raw no-op projection save");
        let raw_noop_package =
            litchi_xlsb::Package::from_slice(&raw_noop_output).expect("raw no-op package");
        assert_eq!(
            before_parts,
            package_part_digests(raw_noop_package.opc_package())
        );

        let mut changed = open_direct_without_drawings(&bytes).expect("changed projection");
        let target = corpus.edits.first().expect("editable scalar");
        let mut edit = changed
            .edit_cell_values(corpus.selected_sheet)
            .expect("projection edit");
        edit.set_value(target.reference, target.after.clone())
            .expect("projection scalar edit");
        let commit = edit.commit().expect("projection commit");
        changed
            .apply_cell_values(corpus.selected_sheet, &commit)
            .expect("projection publication");
        assert!(
            changed.sheet_drawings().is_empty(),
            "drawing-skipped projection must retain its explicit typed-inventory boundary after publication"
        );
        let changed_output = save_workbook(&changed).expect("changed projection save");
        let changed_package =
            litchi_xlsb::Package::from_slice(&changed_output).expect("changed output package");
        let changed_parts = package_part_digests(changed_package.opc_package());
        assert_ne!(
            before_parts, changed_parts,
            "the scalar edit must be observable"
        );
        for (name, digest) in &before_parts {
            if !name.contains("/worksheets/") {
                assert_eq!(
                    changed_parts.get(name),
                    Some(digest),
                    "non-worksheet opaque/drawing part changed: {name}"
                );
            }
        }
        assert_eq!(
            package_relationships(before.opc_package()),
            package_relationships(changed_package.opc_package()),
            "drawing and worksheet relationship graphs changed during scalar publication"
        );

        let drawing_uri =
            litchi_opc::PackURI::new("/xl/drawings/drawing1.xml").expect("drawing URI");
        let mut raw_drawing =
            open_direct_without_drawings(&bytes).expect("raw drawing edit projection");
        let raw_before_parts = package_part_digests(raw_drawing.opc_package());
        let raw_before_relationships = package_relationships(raw_drawing.opc_package());
        raw_drawing
            .edit_opc(|package| {
                let part = package.get_part_mut(&drawing_uri)?;
                let mut blob = part.blob().to_vec();
                for whitespace in [
                    b">\r\n<".as_slice(),
                    b">\n<",
                    b">\r<",
                    b"> \t<",
                    b"> \r<",
                    b">\t<",
                ] {
                    while let Some(offset) = blob
                        .windows(whitespace.len())
                        .position(|window| window == whitespace)
                    {
                        blob.splice(offset..offset + whitespace.len(), b"><".iter().copied());
                    }
                }
                let marker = b"macro=\"\"";
                let offset = blob
                    .windows(marker.len())
                    .position(|window| window == marker)
                    .ok_or_else(|| {
                        litchi_xlsb::package::error::Error::InvalidFormat(
                            "drawing marker missing".to_owned(),
                        )
                    })?;
                blob.splice(
                    offset + marker.len() - 1..offset + marker.len() - 1,
                    b"x".iter().copied(),
                );
                part.set_blob(blob);
                Ok(())
            })
            .expect("raw drawing mutation publication");
        assert!(
            raw_drawing.sheet_drawings().is_empty(),
            "raw drawing mutation must retain the skipped typed boundary"
        );
        let raw_after_parts = package_part_digests(raw_drawing.opc_package());
        assert_ne!(
            raw_before_parts.get("/xl/drawings/drawing1.xml"),
            raw_after_parts.get("/xl/drawings/drawing1.xml"),
            "raw drawing mutation must change the targeted drawing part"
        );
        for (name, digest) in &raw_before_parts {
            if name != "/xl/drawings/drawing1.xml" {
                assert_eq!(
                    raw_after_parts.get(name),
                    Some(digest),
                    "raw drawing mutation changed unrelated package member: {name}"
                );
            }
        }
        assert_eq!(
            raw_before_relationships,
            package_relationships(raw_drawing.opc_package()),
            "raw drawing mutation changed the complete package/part relationship graph"
        );
        let raw_saved = save_workbook(&raw_drawing).expect("raw drawing mutation save");
        let raw_reopened =
            litchi_xlsb::Package::from_slice(&raw_saved).expect("raw drawing mutation reopen");
        assert_eq!(
            raw_after_parts,
            package_part_digests(raw_reopened.opc_package()),
            "raw drawing mutation must survive save/reopen"
        );

        let mut package = litchi_opc::OpcPackage::from_bytes(&bytes).expect("raw OPC package");
        let sheet_uri =
            litchi_opc::PackURI::new("/xl/worksheets/sheet1.bin").expect("worksheet URI");
        let drawing_rel_id = package
            .get_part(&sheet_uri)
            .expect("worksheet part")
            .rels()
            .iter()
            .find(|relationship| relationship.reltype().ends_with("/drawing"))
            .map(|relationship| relationship.r_id().to_owned())
            .expect("fixture drawing relationship");
        package
            .get_part_mut(&sheet_uri)
            .expect("mutable worksheet part")
            .rels_mut()
            .remove(&drawing_rel_id);
        let mut eager_empty =
            litchi_xlsb::Workbook::from_opc_package(package).expect("eager empty workbook");
        assert!(
            eager_empty.sheet_drawings().is_empty(),
            "fixture with no drawing relationship starts with an empty eager inventory"
        );
        eager_empty
            .edit_opc(|package| {
                package
                    .get_part_mut(&sheet_uri)?
                    .rels_mut()
                    .add_relationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
                        .to_owned(),
                    "../drawings/drawing1.xml".to_owned(),
                    drawing_rel_id.clone(),
                    false,
                );
                Ok(())
            })
            .expect("restore drawing relationship");
        assert_eq!(
            eager_empty.sheet_drawings().len(),
            1,
            "an eager workbook must load a drawing added after an initially empty inventory"
        );
    }

    #[test]
    fn skipped_projection_propagates_through_reparse_paths_and_ignores_opaque_drawing() {
        let path = fixture_path();
        let bytes = fs::read(&path).expect("fixture");
        let corpus = Corpus::load(&path).expect("corpus");

        let malformed = |bytes: &[u8]| {
            let mut workbook = open_direct_without_drawings(bytes).expect("skipped workbook");
            let drawing_uri = workbook
                .opc_package()
                .iter_parts()
                .find(|part| part.partname().as_str().contains("/drawings/"))
                .map(|part| part.partname().clone())
                .expect("drawing part");
            workbook
                .edit_opc(|package| {
                    package
                        .get_part_mut(&drawing_uri)?
                        .set_blob(b"<broken-drawing/>".to_vec());
                    Ok(())
                })
                .expect("opaque drawing mutation");
            assert!(workbook.sheet_drawings().is_empty());
            Ok::<_, Error>((workbook, drawing_uri))
        };

        let (mut structure_noop, structure_noop_uri) =
            malformed(&bytes).expect("malformed structure no-op fixture");
        let structure_noop_before = package_part_digests(structure_noop.opc_package());
        let structure_noop_commit = structure_noop
            .edit_workbook_structure()
            .expect("structure no-op edit")
            .commit()
            .expect("structure no-op commit");
        assert_eq!(structure_noop_commit.patch().operation_count(), 0);
        assert_eq!(
            structure_noop_commit.patch().before(),
            structure_noop_commit.patch().after()
        );
        structure_noop
            .apply_workbook_structure(&structure_noop_commit)
            .expect("structure no-op publication");
        assert!(structure_noop.sheet_drawings().is_empty());
        assert_eq!(
            package_part_digests(structure_noop.opc_package()),
            structure_noop_before,
            "structure no-op must retain every opaque part byte"
        );
        assert_eq!(
            package_part_digests(structure_noop.opc_package()).get(structure_noop_uri.as_str()),
            structure_noop_before.get(structure_noop_uri.as_str())
        );

        let (mut structure, drawing_uri) = malformed(&bytes).expect("malformed structure fixture");
        let before_parts = package_part_digests(structure.opc_package());
        let before_drawing = before_parts
            .get(drawing_uri.as_str())
            .cloned()
            .expect("malformed drawing digest");
        let mut structure_edit = structure.edit_workbook_structure().expect("structure edit");
        let new_name = format!("{}_v3", corpus.worksheet_names[0]);
        structure_edit
            .rename_sheet(0, new_name)
            .expect("structure rename");
        let structure_commit = structure_edit.commit().expect("structure commit");
        structure
            .apply_workbook_structure(&structure_commit)
            .expect("structure publication");
        assert!(structure.sheet_drawings().is_empty());
        let structure_parts = package_part_digests(structure.opc_package());
        assert_eq!(
            structure_parts.get(drawing_uri.as_str()),
            Some(&before_drawing)
        );

        let (mut calculation, calculation_drawing_uri) =
            malformed(&bytes).expect("malformed calculation fixture");
        let calculation_before = package_part_digests(calculation.opc_package());
        let calculation_commit = calculation
            .edit_calculation_chain()
            .expect("calculation edit")
            .commit()
            .expect("calculation no-op commit");
        assert!(calculation_commit.patch().is_empty());
        calculation
            .apply_calculation_chain(&calculation_commit)
            .expect("calculation no-op publication");
        assert!(calculation.sheet_drawings().is_empty());
        assert_eq!(
            package_part_digests(calculation.opc_package()),
            calculation_before,
            "calculation-chain no-op must retain every opaque part byte"
        );
        assert_eq!(
            package_part_digests(calculation.opc_package()).get(calculation_drawing_uri.as_str()),
            calculation_before.get(calculation_drawing_uri.as_str())
        );

        let (mut scalar, scalar_drawing_uri) = malformed(&bytes).expect("malformed scalar fixture");
        let scalar_before = package_part_digests(scalar.opc_package());
        let target = corpus.edits.first().expect("editable scalar");
        let mut edit = scalar
            .edit_cell_values(corpus.selected_sheet)
            .expect("scalar edit");
        edit.set_value(target.reference, target.after.clone())
            .expect("scalar value");
        let commit = edit.commit().expect("scalar commit");
        scalar
            .apply_cell_values(corpus.selected_sheet, &commit)
            .expect("scalar publication with opaque drawing");
        assert!(scalar.sheet_drawings().is_empty());
        let scalar_after = package_part_digests(scalar.opc_package());
        assert_eq!(
            scalar_after.get(scalar_drawing_uri.as_str()),
            scalar_before.get(scalar_drawing_uri.as_str()),
            "changed scalar must leave malformed opaque drawing bytes untouched"
        );
    }

    trait SavedBytes {
        fn saved_bytes(self) -> Vec<u8>;
    }

    impl SavedBytes for RunOutcome {
        fn saved_bytes(self) -> Vec<u8> {
            match self {
                RunOutcome::Saved(bytes) => bytes,
                RunOutcome::Projection { .. } => panic!("expected saved bytes"),
            }
        }
    }
}
