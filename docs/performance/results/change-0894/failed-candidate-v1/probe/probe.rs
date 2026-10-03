//! Observable four-leg XLSX differential probe for change 0894.
//!
//! The probe is built twice from this same source, once against an immutable
//! baseline checkout and once against the candidate checkout.  It intentionally
//! uses only the public XLSX APIs.  That makes the result useful for deciding
//! whether a selected-route optimization preserves the behavior visible to a
//! caller, while avoiding any claim about an internal eligibility reason.
//!
//! Commands:
//!
//! ```text
//! xlsx0894-probe input PATH [--jsonl OUT]
//! xlsx0894-probe manifest ROOT... [--jsonl OUT]
//! xlsx0894-probe census ROOT... [--jsonl OUT]
//! ```
//!
//! `manifest` recursively discovers `.xlsx` files in sorted path order.  Every
//! discovered file gets an input record, including files that cannot be opened
//! or hashed.  An open or parse refusal is recorded as data; it is never
//! silently dropped.  The four worksheet legs are, for every sheet in source
//! workbook order: cold source `visit_cells`, cold source `cells`, warm source
//! `visit_cells` after `stored_extent`, and eager model `Workbook::cells`.

use serde::Serialize;
use sha2::{Digest, Sha256};
use std::collections::BTreeSet;
use std::error::Error as StdError;
use std::fs::{self, File};
use std::io::{self, BufWriter, Read, Write};
use std::path::{Path, PathBuf};

type BoxError = Box<dyn StdError>;

const SCHEMA: &str = "litchi.xlsx-observable-differential-0894.v1";
const CENSUS_SCHEMA: &str = "litchi.xlsx-input-census-0894.v1";
const WHOLE_SHEET: &str = "A1:XFD1048576";
const TRANSCRIPT_ALGORITHM: &str = "sha256";
const TRANSCRIPT_FRAMING: &str = "u64le-length-frame-v1(address-display,cell-debug),ordered-cells";

#[derive(Debug, Serialize, Clone)]
struct ErrorLink {
    debug: String,
    display: String,
}

#[derive(Debug, Serialize, Clone)]
struct ErrorRecord {
    debug: String,
    display: String,
    source_chain: Vec<ErrorLink>,
}

fn error_record(error: &dyn StdError) -> ErrorRecord {
    let mut source_chain = Vec::new();
    let mut next = error.source();
    while let Some(source) = next {
        source_chain.push(ErrorLink {
            debug: format!("{source:?}"),
            display: source.to_string(),
        });
        next = source.source();
    }
    ErrorRecord {
        debug: format!("{error:?}"),
        display: error.to_string(),
        source_chain,
    }
}

#[derive(Debug, Serialize, Clone)]
struct TranscriptRecord {
    algorithm: &'static str,
    framing: &'static str,
    sha256: String,
    records: u64,
}

/// A length-framed digest prevents concatenation ambiguity while preserving
/// every public `Cell` variant through its canonical derived `Debug` form.
struct Transcript {
    hasher: Sha256,
    records: u64,
    first: Option<String>,
    last: Option<String>,
}

impl Transcript {
    fn new() -> Self {
        Self {
            hasher: Sha256::new(),
            records: 0,
            first: None,
            last: None,
        }
    }

    fn frame(&mut self, bytes: &[u8]) {
        let length = u64::try_from(bytes.len()).expect("probe frame length fits u64");
        self.hasher.update(length.to_le_bytes());
        self.hasher.update(bytes);
    }

    fn record(&mut self, address: litchi_xlsx::Address, cell: &litchi_xlsx::cell::Cell) {
        let address = address.to_string();
        let debug = format!("{cell:?}");
        self.frame(address.as_bytes());
        self.frame(debug.as_bytes());
        if self.first.is_none() {
            self.first = Some(address.clone());
        }
        self.last = Some(address);
        self.records = self.records.saturating_add(1);
    }

    fn finish(self) -> (TranscriptRecord, Option<String>, Option<String>) {
        (
            TranscriptRecord {
                algorithm: TRANSCRIPT_ALGORITHM,
                framing: TRANSCRIPT_FRAMING,
                sha256: hex(self.hasher.finalize().as_slice()),
                records: self.records,
            },
            self.first,
            self.last,
        )
    }
}

fn hex(bytes: &[u8]) -> String {
    let mut text = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        use std::fmt::Write as _;
        let _ = write!(&mut text, "{byte:02x}");
    }
    text
}

#[derive(Debug, Serialize, Clone)]
struct SourceVersionRecord {
    id: u64,
    revision: u64,
}

#[derive(Debug, Serialize, Clone)]
struct ToleranceRecord {
    status: &'static str,
    clean: Option<bool>,
    total: Option<u64>,
    unrecorded: Option<u32>,
    debug: Option<String>,
    error: Option<ErrorRecord>,
}

fn tolerance_report(report: &litchi_xlsx::leniency::ToleranceReport) -> ToleranceRecord {
    ToleranceRecord {
        status: "ok",
        clean: Some(report.is_clean()),
        total: Some(report.total()),
        unrecorded: Some(report.unrecorded()),
        debug: Some(format!("{report:?}")),
        error: None,
    }
}

fn tolerance_error(error: &litchi_xlsx::Error) -> ToleranceRecord {
    ToleranceRecord {
        status: "error",
        clean: None,
        total: None,
        unrecorded: None,
        debug: None,
        error: Some(error_record(error)),
    }
}

#[derive(Debug, Serialize, Clone)]
struct GuardSheetRecord {
    position: usize,
    name: String,
    kind: String,
    report: ToleranceRecord,
}

#[derive(Debug, Serialize, Clone)]
struct GuardRecord {
    scope: &'static str,
    worksheet_reports: Vec<GuardSheetRecord>,
    aggregate: ToleranceRecord,
}

fn source_guard(
    workbook: &litchi_xlsx::SourceBackedWorkbook,
    sheets: &[litchi_xlsx::SourceWorksheet],
) -> GuardRecord {
    let worksheet_reports = sheets
        .iter()
        .map(|sheet| GuardSheetRecord {
            position: sheet.position(),
            name: sheet.name().to_owned(),
            kind: format!("{:?}", sheet.kind()),
            report: match sheet.tolerance_report() {
                Ok(report) => tolerance_report(report),
                Err(error) => tolerance_error(&error),
            },
        })
        .collect();
    GuardRecord {
        scope: "source-backed-every-worksheet-then-workbook-aggregate",
        worksheet_reports,
        aggregate: tolerance_report(&workbook.tolerance_report()),
    }
}

fn model_guard(workbook: &litchi_xlsx::Workbook, sheets: &[litchi_xlsx::Worksheet]) -> GuardRecord {
    let worksheet_reports = sheets
        .iter()
        .map(|sheet| GuardSheetRecord {
            position: sheet.position(),
            name: sheet.name().to_owned(),
            kind: format!("{:?}", sheet.kind()),
            report: match sheet.tolerance_report() {
                Ok(report) => tolerance_report(report),
                Err(error) => tolerance_error(&error),
            },
        })
        .collect();
    GuardRecord {
        scope: "model-every-worksheet-then-workbook-aggregate",
        worksheet_reports,
        aggregate: tolerance_report(&workbook.tolerance_report()),
    }
}

#[derive(Debug, Serialize, Clone)]
struct LegRecord {
    sheet_position: usize,
    sheet_name: String,
    leg: String,
    status: &'static str,
    observed_records: u64,
    callback_records: Option<u64>,
    first_observed: Option<String>,
    last_observed: Option<String>,
    transcript: TranscriptRecord,
    operation_error: Option<ErrorRecord>,
    guard: Option<GuardRecord>,
    source_version_before: Option<SourceVersionRecord>,
    source_version_before_error: Option<ErrorRecord>,
    source_version_after: Option<SourceVersionRecord>,
    source_version_after_error: Option<ErrorRecord>,
    source_version_stable: Option<bool>,
}

fn empty_leg(sheet_position: usize, sheet_name: &str, leg: &str) -> LegRecord {
    let (transcript, first_observed, last_observed) = Transcript::new().finish();
    LegRecord {
        sheet_position,
        sheet_name: sheet_name.to_owned(),
        leg: leg.to_owned(),
        status: "not_run",
        observed_records: 0,
        callback_records: None,
        first_observed,
        last_observed,
        transcript,
        operation_error: None,
        guard: None,
        source_version_before: None,
        source_version_before_error: None,
        source_version_after: None,
        source_version_after_error: None,
        source_version_stable: None,
    }
}

fn source_version(
    workbook: &litchi_xlsx::SourceBackedWorkbook,
) -> Result<SourceVersionRecord, ErrorRecord> {
    workbook
        .source_version()
        .map(|version| SourceVersionRecord {
            id: version.id(),
            revision: version.revision(),
        })
        .map_err(|error| error_record(&error))
}

fn source_operation(
    worksheet: &litchi_xlsx::SourceWorksheet,
    leg: &str,
    transcript: &mut Transcript,
) -> Result<usize, litchi_xlsx::Error> {
    match leg {
        "cold_source_visit" | "warm_source_visit" => {
            worksheet.visit_cells(WHOLE_SHEET, |address, cell| {
                transcript.record(address, cell);
                Ok(())
            })
        },
        "cold_source_cells" => {
            let values = worksheet.cells(WHOLE_SHEET)?;
            for value in &values {
                transcript.record(value.address, &value.cell);
            }
            Ok(values.len())
        },
        other => Err(litchi_xlsx::Error::Invalid(format!(
            "unknown source probe leg {other}"
        ))),
    }
}

fn finish_source_leg(
    mut result: LegRecord,
    transcript: Transcript,
    operation: Result<usize, litchi_xlsx::Error>,
    warm: bool,
    workbook: &litchi_xlsx::SourceBackedWorkbook,
    sheets: &[litchi_xlsx::SourceWorksheet],
    before: Result<SourceVersionRecord, ErrorRecord>,
) -> LegRecord {
    let (transcript, first_observed, last_observed) = transcript.finish();
    result.transcript = transcript;
    result.first_observed = first_observed;
    result.last_observed = last_observed;
    result.callback_records = if result.leg.ends_with("visit") {
        Some(result.transcript.records)
    } else {
        None
    };
    match operation {
        Ok(count) => {
            result.status = "ok";
            result.observed_records = u64::try_from(count).unwrap_or(u64::MAX);
        },
        Err(error) => {
            result.status = if warm {
                "warmup_or_operation_error"
            } else {
                "operation_error"
            };
            result.operation_error = Some(error_record(&error));
            result.observed_records = result.transcript.records;
        },
    }
    result.source_version_before = before.as_ref().ok().cloned();
    result.source_version_before_error = before.err();
    result.guard = Some(source_guard(workbook, sheets));
    match source_version(workbook) {
        Ok(after) => {
            result.source_version_after = Some(after);
            result.source_version_stable = match (
                result.source_version_before.as_ref(),
                result.source_version_after.as_ref(),
            ) {
                (Some(before), Some(after)) => {
                    Some(before.id == after.id && before.revision == after.revision)
                },
                _ => None,
            };
        },
        Err(error) => result.source_version_after_error = Some(error),
    }
    result
}

fn run_source_leg(path: &Path, sheet_position: usize, sheet_name: &str, leg: &str) -> LegRecord {
    let mut result = empty_leg(sheet_position, sheet_name, leg);
    let workbook = match litchi_xlsx::SourceBackedWorkbook::from_path(path) {
        Ok(workbook) => workbook,
        Err(error) => {
            result.status = "open_error";
            result.operation_error = Some(error_record(&error));
            return result;
        },
    };
    let sheets: Vec<_> = workbook.sheets().collect();
    let before = source_version(&workbook);
    let worksheet = match workbook.sheet(sheet_name) {
        Ok(Some(worksheet)) => worksheet,
        Ok(None) => {
            result.status = "missing_sheet";
            result.guard = Some(source_guard(&workbook, &sheets));
            result.source_version_before = before.as_ref().ok().cloned();
            result.source_version_before_error = before.err();
            match source_version(&workbook) {
                Ok(after) => {
                    result.source_version_after = Some(after);
                    result.source_version_stable = match (
                        result.source_version_before.as_ref(),
                        result.source_version_after.as_ref(),
                    ) {
                        (Some(before), Some(after)) => {
                            Some(before.id == after.id && before.revision == after.revision)
                        },
                        _ => None,
                    };
                },
                Err(error) => result.source_version_after_error = Some(error),
            }
            return result;
        },
        Err(error) => {
            result.status = "sheet_lookup_error";
            result.operation_error = Some(error_record(&error));
            result.guard = Some(source_guard(&workbook, &sheets));
            return result;
        },
    };
    let mut transcript = Transcript::new();
    let operation = if leg == "warm_source_visit" {
        match worksheet.stored_extent() {
            Ok(_) => source_operation(&worksheet, leg, &mut transcript),
            Err(error) => Err(error),
        }
    } else {
        source_operation(&worksheet, leg, &mut transcript)
    };
    finish_source_leg(
        result,
        transcript,
        operation,
        leg == "warm_source_visit",
        &workbook,
        &sheets,
        before,
    )
}

fn model_operation(
    worksheet: &litchi_xlsx::Worksheet,
    transcript: &mut Transcript,
) -> Result<usize, litchi_xlsx::Error> {
    let values = worksheet.cells(WHOLE_SHEET)?;
    let mut count = 0usize;
    for (address, cell) in values {
        transcript.record(address, cell);
        count = count.saturating_add(1);
    }
    Ok(count)
}

fn run_model_leg(path: &Path, sheet_position: usize, sheet_name: &str) -> LegRecord {
    let mut result = empty_leg(sheet_position, sheet_name, "eager_model_cells");
    let workbook = match litchi_xlsx::Workbook::open(path) {
        Ok(workbook) => workbook,
        Err(error) => {
            result.status = "open_error";
            result.operation_error = Some(error_record(&error));
            return result;
        },
    };
    let sheets: Vec<_> = workbook.sheets().collect();
    let worksheet = match workbook.sheet(sheet_name) {
        Ok(Some(worksheet)) => worksheet,
        Ok(None) => {
            result.status = "missing_sheet";
            result.guard = Some(model_guard(&workbook, &sheets));
            return result;
        },
        Err(error) => {
            result.status = "sheet_lookup_error";
            result.operation_error = Some(error_record(&error));
            result.guard = Some(model_guard(&workbook, &sheets));
            return result;
        },
    };
    let mut transcript = Transcript::new();
    let operation = model_operation(&worksheet, &mut transcript);
    let (transcript, first_observed, last_observed) = transcript.finish();
    result.transcript = transcript;
    result.first_observed = first_observed;
    result.last_observed = last_observed;
    match operation {
        Ok(count) => {
            result.status = "ok";
            result.observed_records = u64::try_from(count).unwrap_or(u64::MAX);
        },
        Err(error) => {
            result.status = "operation_error";
            result.operation_error = Some(error_record(&error));
            result.observed_records = result.transcript.records;
        },
    }
    result.guard = Some(model_guard(&workbook, &sheets));
    result
}

#[derive(Debug, Serialize, Clone)]
struct InputMeta {
    path: String,
    bytes: Option<u64>,
    sha256: Option<String>,
    metadata_error: Option<String>,
}

fn input_meta(path: &Path) -> InputMeta {
    let bytes = match fs::metadata(path) {
        Ok(metadata) => Some(metadata.len()),
        Err(error) => {
            return InputMeta {
                path: path.display().to_string(),
                bytes: None,
                sha256: None,
                metadata_error: Some(error.to_string()),
            };
        },
    };
    match file_sha256(path) {
        Ok(sha256) => InputMeta {
            path: path.display().to_string(),
            bytes,
            sha256: Some(sha256),
            metadata_error: None,
        },
        Err(error) => InputMeta {
            path: path.display().to_string(),
            bytes,
            sha256: None,
            metadata_error: Some(error.to_string()),
        },
    }
}

fn file_sha256(path: &Path) -> io::Result<String> {
    let mut file = File::open(path)?;
    let mut hasher = Sha256::new();
    let mut buffer = [0_u8; 128 * 1024];
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        hasher.update(&buffer[..read]);
    }
    Ok(hex(hasher.finalize().as_slice()))
}

#[derive(Debug, Serialize, Clone)]
struct SheetCensus {
    position: usize,
    name: String,
    kind: String,
    visibility: String,
}

#[derive(Debug, Serialize)]
struct InputRecord {
    schema: &'static str,
    kind: &'static str,
    status: &'static str,
    input: InputMeta,
    source_open_error: Option<ErrorRecord>,
    sheets: Vec<SheetCensus>,
    legs: Vec<LegRecord>,
}

fn source_census(
    path: &Path,
) -> Result<(litchi_xlsx::SourceBackedWorkbook, Vec<SheetCensus>), ErrorRecord> {
    let workbook =
        litchi_xlsx::SourceBackedWorkbook::from_path(path).map_err(|error| error_record(&error))?;
    let sheets = workbook
        .sheets()
        .map(|sheet| SheetCensus {
            position: sheet.position(),
            name: sheet.name().to_owned(),
            kind: format!("{:?}", sheet.kind()),
            visibility: format!("{:?}", sheet.visibility()),
        })
        .collect();
    Ok((workbook, sheets))
}

fn probe_input(path: &Path) -> InputRecord {
    let input = input_meta(path);
    let census = source_census(path);
    let (workbook, sheets) = match census {
        Ok(value) => value,
        Err(error) => {
            return InputRecord {
                schema: SCHEMA,
                kind: "input",
                status: "source_open_error",
                input,
                source_open_error: Some(error),
                sheets: Vec::new(),
                legs: Vec::new(),
            };
        },
    };
    let mut legs = Vec::with_capacity(sheets.len() * 4);
    for sheet in &sheets {
        legs.push(run_source_leg(
            path,
            sheet.position,
            &sheet.name,
            "cold_source_visit",
        ));
        legs.push(run_source_leg(
            path,
            sheet.position,
            &sheet.name,
            "cold_source_cells",
        ));
        legs.push(run_source_leg(
            path,
            sheet.position,
            &sheet.name,
            "warm_source_visit",
        ));
        legs.push(run_model_leg(path, sheet.position, &sheet.name));
    }
    drop(workbook);
    InputRecord {
        schema: SCHEMA,
        kind: "input",
        status: "ok",
        input,
        source_open_error: None,
        sheets,
        legs,
    }
}

#[derive(Debug, Serialize)]
struct CensusRecord {
    schema: &'static str,
    kind: &'static str,
    status: &'static str,
    input: InputMeta,
}

#[derive(Debug, Serialize)]
struct SummaryRecord {
    schema: &'static str,
    kind: &'static str,
    inputs: usize,
    worksheets: usize,
    legs: usize,
    discovery_errors: usize,
    source_open_errors: usize,
    operation_errors: usize,
    note: &'static str,
}

fn write_json_line<T: Serialize>(writer: &mut impl Write, value: &T) -> io::Result<()> {
    serde_json::to_writer(&mut *writer, value).map_err(io::Error::other)?;
    writer.write_all(b"\n")
}

fn output_writer(path: Option<&Path>) -> Result<BufWriter<Box<dyn Write>>, BoxError> {
    let writer: Box<dyn Write> = match path {
        None => Box::new(io::stdout()),
        Some(path) if path == Path::new("-") => Box::new(io::stdout()),
        Some(path) => Box::new(File::create(path)?),
    };
    Ok(BufWriter::new(writer))
}

fn normalize_path(path: &Path) -> PathBuf {
    if path.is_absolute() {
        path.to_owned()
    } else {
        std::env::current_dir()
            .map(|root| root.join(path))
            .unwrap_or_else(|_| path.to_owned())
    }
}

fn discover_paths(roots: &[String]) -> (Vec<PathBuf>, Vec<String>) {
    let mut files = BTreeSet::new();
    let mut errors = Vec::new();
    for root in roots {
        let root = normalize_path(Path::new(root));
        discover_one(&root, &mut files, &mut errors);
    }
    (files.into_iter().collect(), errors)
}

fn discover_one(path: &Path, files: &mut BTreeSet<PathBuf>, errors: &mut Vec<String>) {
    let metadata = match fs::metadata(path) {
        Ok(metadata) => metadata,
        Err(error) => {
            errors.push(format!("{}: {error}", path.display()));
            return;
        },
    };
    if metadata.is_file() {
        let is_xlsx = path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"));
        if is_xlsx {
            files.insert(path.to_owned());
        }
        return;
    }
    if !metadata.is_dir() {
        errors.push(format!(
            "{}: not a regular file or directory",
            path.display()
        ));
        return;
    }
    let entries = match fs::read_dir(path) {
        Ok(entries) => entries,
        Err(error) => {
            errors.push(format!("{}: {error}", path.display()));
            return;
        },
    };
    let mut children = Vec::new();
    for entry in entries {
        match entry {
            Ok(entry) => children.push(entry.path()),
            Err(error) => errors.push(format!("{}: {error}", path.display())),
        }
    }
    children.sort();
    for child in children {
        discover_one(&child, files, errors);
    }
}

fn parse_output(args: &[String]) -> Result<(Vec<String>, Option<PathBuf>), BoxError> {
    let mut roots = Vec::new();
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        match args[index].as_str() {
            "--jsonl" => {
                index += 1;
                let path = args.get(index).ok_or("--jsonl requires a path")?.to_owned();
                output = Some(PathBuf::from(path));
            },
            value if value.starts_with('-') => {
                return Err(format!("unknown option {value}").into());
            },
            value => roots.push(value.to_owned()),
        }
        index += 1;
    }
    if roots.is_empty() {
        return Err("at least one input root is required".into());
    }
    Ok((roots, output))
}

fn run_census(roots: &[String], output: Option<&Path>) -> Result<(), BoxError> {
    let (paths, errors) = discover_paths(roots);
    let mut writer = output_writer(output)?;
    for path in paths {
        let input = input_meta(&path);
        let status = if input.metadata_error.is_some() {
            "metadata_error"
        } else {
            "discovered"
        };
        write_json_line(
            &mut writer,
            &CensusRecord {
                schema: CENSUS_SCHEMA,
                kind: "input",
                status,
                input,
            },
        )?;
    }
    for error in errors {
        write_json_line(
            &mut writer,
            &CensusRecord {
                schema: CENSUS_SCHEMA,
                kind: "discovery_error",
                status: "error",
                input: InputMeta {
                    path: error,
                    bytes: None,
                    sha256: None,
                    metadata_error: Some("recursive discovery failed".to_owned()),
                },
            },
        )?;
    }
    writer.flush()?;
    Ok(())
}

fn run_probe(roots: &[String], output: Option<&Path>) -> Result<(), BoxError> {
    let (paths, discovery_errors) = discover_paths(roots);
    let mut writer = output_writer(output)?;
    let mut inputs = 0usize;
    let mut worksheets = 0usize;
    let mut legs = 0usize;
    let discovery_error_count = discovery_errors.len();
    let mut source_open_errors = 0usize;
    let mut operation_errors = 0usize;
    for error in discovery_errors {
        write_json_line(
            &mut writer,
            &InputRecord {
                schema: SCHEMA,
                kind: "discovery_error",
                status: "discovery_error",
                input: InputMeta {
                    path: error,
                    bytes: None,
                    sha256: None,
                    metadata_error: Some("recursive discovery failed".to_owned()),
                },
                source_open_error: None,
                sheets: Vec::new(),
                legs: Vec::new(),
            },
        )?;
    }
    for path in paths {
        let record = probe_input(&path);
        inputs += 1;
        worksheets += record.sheets.len();
        legs += record.legs.len();
        if record.source_open_error.is_some() {
            source_open_errors += 1;
        }
        operation_errors += record
            .legs
            .iter()
            .filter(|leg| leg.operation_error.is_some())
            .count();
        write_json_line(&mut writer, &record)?;
    }
    write_json_line(
        &mut writer,
        &SummaryRecord {
            schema: SCHEMA,
            kind: "summary",
            inputs,
            worksheets,
            legs,
            discovery_errors: discovery_error_count,
            source_open_errors,
            operation_errors,
            note: "Observable four-leg results; compare baseline and candidate JSONL externally and retain explicit refusals.",
        },
    )?;
    writer.flush()?;
    Ok(())
}

fn usage() -> &'static str {
    "usage: xlsx0894-probe input PATH [--jsonl OUT] | manifest ROOT... [--jsonl OUT] | census ROOT... [--jsonl OUT]"
}

fn main() -> Result<(), BoxError> {
    let mut args = std::env::args().skip(1);
    let command = args.next().ok_or(usage())?;
    let remaining: Vec<_> = args.collect();
    let (roots, output) = parse_output(&remaining)?;
    match command.as_str() {
        "input" => {
            if roots.len() != 1 {
                return Err("input accepts exactly one path".into());
            }
            let mut writer = output_writer(output.as_deref())?;
            let record = probe_input(Path::new(&roots[0]));
            write_json_line(&mut writer, &record)?;
            write_json_line(
                &mut writer,
                &SummaryRecord {
                    schema: SCHEMA,
                    kind: "summary",
                    inputs: 1,
                    worksheets: record.sheets.len(),
                    legs: record.legs.len(),
                    discovery_errors: 0,
                    source_open_errors: usize::from(record.source_open_error.is_some()),
                    operation_errors: record
                        .legs
                        .iter()
                        .filter(|leg| leg.operation_error.is_some())
                        .count(),
                    note: "Observable four-leg result for one explicit input.",
                },
            )?;
            writer.flush()?;
        },
        "manifest" => run_probe(&roots, output.as_deref())?,
        "census" => run_census(&roots, output.as_deref())?,
        _ => return Err(usage().into()),
    }
    Ok(())
}
