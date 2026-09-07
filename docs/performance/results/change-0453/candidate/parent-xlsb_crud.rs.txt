//! Opt-in XLSB semantic CRUD baseline.
//!
//! This binary intentionally lives outside the default perf matrix. It uses
//! only the public XLSB owner and umbrella-facade APIs, and keeps the complete
//! reopen/preservation checks outside the timed sample interval. The default
//! corpus is a public POI workbook with opaque members and a VBA payload; the
//! payload is treated as inert and is checked as an untouched package member.

use litchi::common::FileFormat;
use litchi::sheet::{WorkbookTrait, Worksheet};
use litchi_xlsb::cell_values::{CellError, Reference, StoredCell, Value};
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::env;
use std::fmt;
use std::fs;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::str::FromStr;
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
        let mut warmup = DEFAULT_WARMUP;
        let mut samples = DEFAULT_SAMPLES;
        let mut fixture = PathBuf::from(DEFAULT_FIXTURE);
        let mut json = None;
        let mut arguments = arguments.into_iter();
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
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
        Ok(Self {
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
    "usage: xlsb_crud --case <all|open_identify|worksheet_catalog|selected_worksheet_cell|full_stored_cell_scan|full_text|noop_transaction_commit_save|edit_one_existing_scalar_save|edit_ceil_one_percent_existing_cells_save> [--fixture PATH] [--warmup N] [--samples N] [--json PATH]"
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
    sparse_iteration_without_rectangular_expansion: bool,
}

#[derive(Debug, Clone, Serialize)]
struct CaseReport {
    case: Case,
    timing_scope: &'static str,
    input_bytes: usize,
    output_bytes: Option<usize>,
    output_sha256: Option<String>,
    observed_count: Option<usize>,
    observed_text_bytes: Option<usize>,
    observed_text_sha256: Option<String>,
    selected_coordinates: Vec<String>,
    statistics: Statistics,
    gates: GateReport,
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
        reports.push(benchmark_case(&corpus, case, args.warmup, args.samples)?);
    }
    let binary_identity = litchi_perf_baseline::current_executable_identity()?;
    let report = Report {
        schema: "xlsb-crud-v1",
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
            selected_coordinate: coordinate(selected_cell.reference()),
            edits,
            stored_cell_count: cells.len(),
            coordinates,
            full_text_sha256: sha256_hex(full_text.as_bytes()),
            full_text_bytes: full_text.len(),
            part_digests,
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
            package_part_count: self.part_digests.len(),
        }
    }
}

fn benchmark_case(
    corpus: &Corpus,
    case: Case,
    warmup: usize,
    samples: usize,
) -> Result<CaseReport> {
    for _ in 0..warmup {
        let outcome = run_case(corpus, case)?;
        std::hint::black_box(outcome);
    }
    let mut elapsed = Vec::with_capacity(samples);
    let mut output_identities = Vec::with_capacity(samples);
    for _ in 0..samples {
        let started = Instant::now();
        let outcome = run_case(corpus, case)?;
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let outcome = std::hint::black_box(outcome);
        if let RunOutcome::Saved(bytes) = outcome {
            output_identities.push((bytes.len(), sha256_hex(&bytes)));
        }
        elapsed.push(elapsed_ns);
    }
    let representative = run_case(corpus, case)?;
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
    let gate = verify_case(
        corpus,
        case,
        output.as_deref(),
        &output_identities,
        observed_count,
        observed_text_bytes,
        observed_text_sha256.as_deref(),
    )?;
    ensure_gates(case, &gate)?;
    Ok(CaseReport {
        case,
        timing_scope: timing_scope(case),
        input_bytes: corpus.bytes.len(),
        output_bytes,
        output_sha256,
        observed_count,
        observed_text_bytes,
        observed_text_sha256,
        selected_coordinates: selected_coordinates(case, corpus),
        statistics: Statistics::from_samples(warmup, elapsed),
        gates: gate,
    })
}

fn ensure_gates(case: Case, gate: &GateReport) -> Result<()> {
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
    if !gate.tight_cell_limits_refused {
        return Err(format!("{case}: tight cell-limit refusal gate failed").into());
    }
    if !gate.sparse_iteration_without_rectangular_expansion {
        return Err(format!("{case}: sparse-iteration gate failed").into());
    }
    Ok(())
}

fn run_case(corpus: &Corpus, case: Case) -> Result<RunOutcome> {
    match case {
        Case::OpenIdentify => {
            let format = litchi::detect_file_format_from_bytes(&corpus.bytes)
                .ok_or_else(|| "facade could not identify the XLSB fixture".to_string())?;
            if format != FileFormat::Xlsb {
                return Err(format!("facade identified fixture as {format:?}, not XLSB").into());
            }
            let workbook = litchi::sheet::open_xlsb_workbook_from_bytes(&corpus.bytes)?;
            Ok(RunOutcome::Projection {
                count: workbook.worksheet_count(),
                text: None,
            })
        },
        Case::WorksheetCatalog => {
            let workbook = open_direct(&corpus.bytes)?;
            Ok(RunOutcome::Projection {
                count: workbook.worksheet_count(),
                text: Some(workbook.worksheet_names().join("\u{1f}")),
            })
        },
        Case::SelectedWorksheetCell => {
            let workbook = open_direct(&corpus.bytes)?;
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
            let workbook = open_direct(&corpus.bytes)?;
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
            let mut workbook = open_direct(&corpus.bytes)?;
            let snapshot = workbook.cell_values(corpus.selected_sheet)?;
            let commit = snapshot.edit().commit()?;
            if !commit.patch().is_empty() {
                return Err("public exact no-op commit unexpectedly changed bytes".into());
            }
            workbook.apply_cell_values(corpus.selected_sheet, &commit)?;
            Ok(RunOutcome::Saved(save_workbook(&workbook)?))
        },
        Case::EditOneExistingScalarSave => {
            let mut workbook = open_direct(&corpus.bytes)?;
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
            let mut workbook = open_direct(&corpus.bytes)?;
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
        let workbook = open_direct(&corpus.bytes)?;
        let limits = litchi_xlsb::cell_values::Limits::new(1, 1, 1, 1);
        workbook
            .cell_values_with_limits(corpus.selected_sheet, limits)
            .is_err()
    };
    let sparse_iteration_without_rectangular_expansion = sparse_gate(corpus)?;
    let catalog_text_sha256 = sha256_hex(corpus.worksheet_names.join("\u{1f}").as_bytes());
    let semantic_projection_ok = match case {
        Case::OpenIdentify => observed_count == Some(corpus.worksheet_names.len()),
        Case::WorksheetCatalog => {
            observed_count == Some(corpus.worksheet_names.len())
                && observed_text_sha256 == Some(catalog_text_sha256.as_str())
        },
        Case::SelectedWorksheetCell => observed_count == Some(1),
        Case::FullStoredCellScan => observed_count == Some(corpus.stored_cell_count),
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
            sparse_iteration_without_rectangular_expansion,
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
        sparse_iteration_without_rectangular_expansion,
    };
    Ok(report)
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
    Ok(corpus.stored_cell_count < usize::try_from(area).unwrap_or(usize::MAX))
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

fn timing_scope(case: Case) -> &'static str {
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

fn facade_text(bytes: &[u8]) -> Result<String> {
    Ok(litchi::sheet::Workbook::from_bytes(bytes.to_vec())?.text()?)
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

fn package_part_digests(package: &litchi_opc::OpcPackage) -> BTreeMap<String, String> {
    package
        .iter_parts()
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
        .collect()
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
        assert!(ensure_gates(Case::NoopTransactionCommitSave, &failing_gate).is_err());

        let edited = run_case(&corpus, Case::EditOneExistingScalarSave).expect("edit");
        let RunOutcome::Saved(edited_bytes) = edited else {
            panic!("edit must save")
        };
        let edit_gate = verify_case(
            &corpus,
            Case::EditOneExistingScalarSave,
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
