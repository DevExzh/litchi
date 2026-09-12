//! Standalone correctness and planning-cost guards for source-backed XLSX.
//!
//! The benchmark deliberately measures only `edit_sheets`.  Source opening,
//! fixture construction, selector construction, and every correctness oracle
//! are outside the measured interval.  The result of the call is kept alive
//! until both the clock and the allocator region have ended.

use std::env;
use std::error::Error;
use std::ffi::OsString;
use std::fs;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_core::OwnedSource;
use litchi_xlsx::cell_values::SourceBackedEditor;
use litchi_xlsx::{Address, Error as XlsxError, Number, Selector, Value};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;

use crate::allocation_metrics;

type AnyResult<T> = Result<T, Box<dyn Error>>;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPC_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WORKSHEET_PART: &str = "xl/worksheets/sheet1.xml";
const MAX_SOURCE_BYTES: usize = 32 * 1024 * 1024;
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_WARMUP: usize = 1_000;
const MAX_SAMPLES: usize = 10_000;

#[derive(Clone, Copy, Debug)]
enum Shape {
    Medium,
    DenseSparse,
}

impl Shape {
    const fn name(self) -> &'static str {
        match self {
            Self::Medium => "medium",
            Self::DenseSparse => "dense-sparse",
        }
    }

    const fn dimensions(self) -> (usize, usize) {
        match self {
            Self::Medium => (96, 96),
            Self::DenseSparse => (128, 128),
        }
    }

    fn parse(value: &str) -> AnyResult<Self> {
        match value {
            "medium" => Ok(Self::Medium),
            "dense-sparse" | "dense_sparse" => Ok(Self::DenseSparse),
            other => Err(format!("unknown XLSX planning-guard shape '{other}'").into()),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum GuardCase {
    Valid,
    LateValidator,
    LateRaw,
}

impl GuardCase {
    const fn name(self) -> &'static str {
        match self {
            Self::Valid => "valid",
            Self::LateValidator => "late-validator",
            Self::LateRaw => "late-raw",
        }
    }

    const fn expected_error(self) -> Option<&'static str> {
        match self {
            Self::LateValidator => Some("value-only edits refuse attribute 'future' on 'c'"),
            Self::LateRaw => Some("invalid worksheet boolean 'maybe'"),
            Self::Valid => None,
        }
    }

    fn parse(value: &str) -> AnyResult<Self> {
        match value {
            "valid" => Ok(Self::Valid),
            "late-validator" | "late_validator" | "late-validator-unknown-attribute" => {
                Ok(Self::LateValidator)
            },
            "late-raw" | "late_raw" | "late-raw-invalid-boolean" => Ok(Self::LateRaw),
            // A valid empty transaction is the semantic no-op guard. Keep
            // the spelling as an input compatibility alias while emitting
            // one canonical case identity.
            "semantic-noop" | "semantic_noop" | "noop" => Ok(Self::Valid),
            other => Err(format!("unknown XLSX planning-guard case '{other}'").into()),
        }
    }
}

#[derive(Debug)]
struct Arguments {
    shape: Shape,
    case: GuardCase,
    warmup: usize,
    samples: usize,
    json_path: Option<PathBuf>,
}

fn usage() -> &'static str {
    "usage: xlsx_planning_guard --shape <medium|dense-sparse> --case <valid|late-validator|late-raw|semantic-noop> [--warmup N] [--samples N] [--json [PATH]]"
}

fn parse_count(value: &str, label: &str, maximum: usize) -> AnyResult<usize> {
    let count = value
        .parse::<usize>()
        .map_err(|source| format!("invalid {label} '{value}': {source}"))?;
    if count > maximum {
        return Err(format!("{label} {count} exceeds maximum {maximum}").into());
    }
    Ok(count)
}

fn parse_args<I>(arguments: I) -> AnyResult<Arguments>
where
    I: IntoIterator<Item = OsString>,
{
    let mut shape = None;
    let mut case = None;
    let mut warmup = 5;
    let mut samples = if cfg!(feature = "allocator-metrics") {
        10
    } else {
        30
    };
    let mut json_path = None;
    let mut args = arguments
        .into_iter()
        .map(|argument| argument.to_string_lossy().into_owned())
        .peekable();
    while let Some(argument) = args.next() {
        match argument.as_str() {
            "--shape" | "--xlsx-cell-crud-shape" => {
                let value = args.next().ok_or("missing value for --shape")?;
                shape = Some(Shape::parse(&value)?);
            },
            "--case" => {
                let value = args.next().ok_or("missing value for --case")?;
                case = Some(GuardCase::parse(&value)?);
            },
            "--warmup" => {
                let value = args.next().ok_or("missing value for --warmup")?;
                warmup = parse_count(&value, "warmup", MAX_WARMUP)?;
            },
            "--samples" => {
                let value = args.next().ok_or("missing value for --samples")?;
                samples = parse_count(&value, "samples", MAX_SAMPLES)?;
            },
            "--json" => {
                if args.peek().is_some_and(|value| !value.starts_with('-')) {
                    json_path = Some(PathBuf::from(args.next().expect("peeked JSON path")));
                }
            },
            "--help" | "-h" => return Err(usage().into()),
            other => return Err(format!("unknown argument '{other}'\n{}", usage()).into()),
        }
    }
    let shape = shape.ok_or_else(|| format!("missing --shape\n{}", usage()))?;
    let case = case.ok_or_else(|| format!("missing --case\n{}", usage()))?;
    if samples == 0 {
        return Err("samples must be positive".into());
    }
    Ok(Arguments {
        shape,
        case,
        warmup,
        samples,
        json_path,
    })
}

#[derive(Clone, Debug)]
struct Fixture {
    archive: Arc<Vec<u8>>,
    worksheet: Arc<Vec<u8>>,
    source_sha256: String,
    worksheet_sha256: String,
    rows: usize,
    columns: usize,
}

fn sha256(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn column_name(mut column: usize) -> String {
    let mut output = String::new();
    column += 1;
    while column != 0 {
        let remainder = (column - 1) % 26;
        output.push(char::from(
            b'A' + u8::try_from(remainder).expect("column remainder"),
        ));
        column = (column - 1) / 26;
    }
    output.chars().rev().collect()
}

fn worksheet_xml(shape: Shape, case: GuardCase) -> AnyResult<Vec<u8>> {
    let (rows, columns) = shape.dimensions();
    let last_cell = format!("{}{}", column_name(columns - 1), rows);
    let mut xml = String::new();
    xml.try_reserve(1_024 + rows * columns * 28)
        .map_err(|source| format!("worksheet XML reservation failed: {source}"))?;
    xml.push_str(&format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:{last_cell}"/><sheetData>"#
    ));
    for row in 0..rows {
        let row_number = row + 1;
        xml.push_str(&format!(r#"<row r="{row_number}">"#));
        for column in 0..columns {
            let address = format!("{}{}", column_name(column), row_number);
            let ordinal = row
                .checked_mul(columns)
                .and_then(|value| value.checked_add(column))
                .and_then(|value| value.checked_add(1))
                .ok_or("worksheet value ordinal overflow")?;
            let final_cell = row + 1 == rows && column + 1 == columns;
            match (case, final_cell) {
                (GuardCase::LateValidator, true) => xml.push_str(&format!(
                    r#"<c r="{address}" future="1"><v>{ordinal}</v></c>"#
                )),
                (GuardCase::LateRaw, true) => {
                    xml.push_str(&format!(r#"<c r="{address}" t="b"><v>maybe</v></c>"#));
                },
                _ => xml.push_str(&format!(r#"<c r="{address}"><v>{ordinal}</v></c>"#)),
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    let bytes = xml.into_bytes();
    if bytes.len() > MAX_XML_BYTES {
        return Err(format!("worksheet XML exceeds {MAX_XML_BYTES} bytes").into());
    }
    Ok(bytes)
}

fn archive_for(shape: Shape, worksheet: &[u8]) -> AnyResult<Vec<u8>> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let workbook = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#
    );
    let root_rels = format!(
        r#"<Relationships xmlns="{OPC_REL}"><Relationship Id="rIdRoot" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook_rels = format!(
        r#"<Relationships xmlns="{OPC_REL}"><Relationship Id="rIdSheet" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", content_types.as_bytes())?;
    writer.write_stored("_rels/.rels", root_rels.as_bytes())?;
    writer.write_stored("xl/workbook.xml", workbook.as_bytes())?;
    writer.write_stored("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes())?;
    writer.write_stored(WORKSHEET_PART, worksheet)?;
    let archive = writer.finish_to_bytes()?;
    if archive.len() > MAX_SOURCE_BYTES {
        return Err(format!("{shape:?} source exceeds {MAX_SOURCE_BYTES} bytes").into());
    }
    Ok(archive)
}

fn fixture(shape: Shape, case: GuardCase) -> AnyResult<Fixture> {
    let worksheet = worksheet_xml(shape, case)?;
    let archive = archive_for(shape, &worksheet)?;
    let (rows, columns) = shape.dimensions();
    Ok(Fixture {
        source_sha256: sha256(&archive),
        worksheet_sha256: sha256(&worksheet),
        archive: Arc::new(archive),
        worksheet: Arc::new(worksheet),
        rows,
        columns,
    })
}

#[derive(Clone, Debug, Serialize)]
struct LogicalError {
    status: &'static str,
    variant: Option<&'static str>,
    message: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct PhaseSample {
    order: usize,
    duration_ns: u128,
    allocation_metrics: allocation_metrics::Sample,
}

#[derive(Clone, Debug)]
struct Iteration {
    duration_ns: u128,
    allocation: allocation_metrics::Sample,
    logical_error: LogicalError,
    source_unchanged: bool,
    retry_preserved_error: bool,
    valid_snapshot_values: bool,
    empty_commit_is_noop: bool,
    expected_error_exact: bool,
    commit_outside_timing: bool,
}

#[derive(Serialize)]
struct BinaryIdentity {
    path: String,
    sha256: String,
    bytes: u64,
}

#[derive(Serialize)]
struct RunnerIdentity {
    git_revision: Option<String>,
    git_dirty: Option<bool>,
    rustc_vv: Option<String>,
    os: String,
    arch: String,
    profile: &'static str,
}

#[derive(Serialize)]
struct SourceIdentity {
    generator: &'static str,
    shape: &'static str,
    rows: usize,
    columns: usize,
    archive_bytes: usize,
    worksheet_bytes: usize,
    source_sha256: String,
    worksheet_sha256: String,
    worksheet_member: &'static str,
    compression: &'static str,
    fixture_kind: &'static str,
}

#[derive(Serialize)]
struct Correctness {
    source_unchanged: bool,
    retry_preserved_error: bool,
    valid_snapshot_values: bool,
    empty_commit_is_noop: bool,
    expected_error_exact: bool,
    commit_outside_timing: bool,
}

#[derive(Serialize)]
struct Report {
    schema: &'static str,
    tool: &'static str,
    case: &'static str,
    shape: &'static str,
    warmup_iterations: usize,
    samples: usize,
    source: SourceIdentity,
    source_sha256: String,
    logical_error: LogicalError,
    phase: PhaseReport,
    correctness: Correctness,
    binary: BinaryIdentity,
    runner: RunnerIdentity,
    allocation_scope: &'static str,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
    timing_scope: &'static str,
    performance_claim: &'static str,
}

#[derive(Serialize)]
struct PhaseReport {
    name: &'static str,
    sample_order: Vec<usize>,
    duration_ns: Vec<u128>,
    allocation_metrics: Vec<allocation_metrics::Sample>,
    samples: Vec<PhaseSample>,
}

fn run_iteration(
    fixture: &Fixture,
    selectors: &[Selector<'static>],
    case: GuardCase,
) -> AnyResult<Iteration> {
    let source: Arc<dyn litchi_core::ReadAt> =
        Arc::new(OwnedSource::from_arc(Arc::clone(&fixture.archive)));
    // Opening is deliberately outside the planning interval.  This also
    // ensures every sample starts with an independent deferred-package cache.
    let editor = SourceBackedEditor::from_read_at(source)?;

    let allocation_region = allocation_metrics::begin();
    let started = Instant::now();
    let planned = editor.edit_sheets(selectors.iter().cloned());
    let duration_ns = started.elapsed().as_nanos();
    // Keep `planned` alive until the region has been closed. This makes the
    // boundary include result retention while excluding all later oracles.
    let allocation = allocation_region
        .finish()
        .unwrap_or_else(allocation_metrics::unavailable_sample);
    if cfg!(feature = "allocator-metrics")
        && allocation.status != allocation_metrics::Status::Measured
    {
        return Err(format!(
            "allocator planning region was not measured: {:?}",
            allocation.status
        )
        .into());
    }

    let logical_error = match planned {
        Ok(transaction) => {
            if case.expected_error().is_some() {
                return Err(format!(
                    "{} unexpectedly accepted invalid planning guard",
                    case.name()
                )
                .into());
            }
            validate_valid_transaction(transaction)?;
            if sha256(&fixture.archive) != fixture.source_sha256 {
                return Err("valid planning guard changed source bytes".into());
            }
            LogicalError {
                status: "accepted",
                variant: None,
                message: None,
            }
        },
        Err(error) => {
            let expected = case
                .expected_error()
                .ok_or_else(|| format!("{} unexpectedly refused valid source", case.name()))?;
            let (variant, message) = classify_error(&error);
            if variant != "Invalid" || message != expected {
                return Err(format!(
                    "{} error mismatch: expected Invalid({expected:?}), got {variant}({message:?})",
                    case.name()
                )
                .into());
            }
            let retry = editor.edit_sheets(selectors.iter().cloned());
            let retry_message = match retry {
                Ok(_) => return Err("invalid planning guard succeeded on retry".into()),
                Err(retry_error) => {
                    let (retry_variant, retry_message) = classify_error(&retry_error);
                    if retry_variant != variant || retry_message != message {
                        return Err("invalid planning guard changed its error on retry".into());
                    }
                    true
                },
            };
            if !retry_message {
                return Err("invalid planning guard retry was not observed".into());
            }
            LogicalError {
                status: "expected_failure",
                variant: Some("Invalid"),
                message: Some(message),
            }
        },
    };

    let source_unchanged = sha256(&fixture.archive) == fixture.source_sha256;
    if !source_unchanged {
        return Err("planning guard changed source bytes".into());
    }

    Ok(Iteration {
        duration_ns,
        allocation,
        logical_error,
        source_unchanged,
        retry_preserved_error: case.expected_error().is_some(),
        valid_snapshot_values: case.expected_error().is_none(),
        empty_commit_is_noop: case.expected_error().is_none(),
        expected_error_exact: case.expected_error().is_some(),
        commit_outside_timing: case.expected_error().is_none(),
    })
}

fn validate_valid_transaction(
    transaction: litchi_xlsx::cell_values::MultiSourceEdit,
) -> AnyResult<()> {
    if transaction.worksheet_count() != 1 {
        return Err("valid guard selected an unexpected worksheet count".into());
    }
    let expected = Value::Number(Number::new("1").map_err(|error| error.to_string())?);
    let address = Address::from_a1("A1")?;
    if transaction.before().value(0, address) != Some(&expected) {
        return Err("valid guard snapshot did not retain A1".into());
    }
    // This is intentionally outside the phase clock: the empty commit is the
    // exact no-op oracle for both `valid` and `semantic-noop`.
    let commit = transaction.commit()?;
    if commit.changed() || !commit.patch().is_empty() {
        return Err("empty planning guard commit was not an exact no-op".into());
    }
    if commit.snapshot().value(0, address) != Some(&expected) {
        return Err("empty planning guard commit lost A1".into());
    }
    Ok(())
}

fn classify_error(error: &XlsxError) -> (&'static str, String) {
    match error {
        XlsxError::Invalid(message) => ("Invalid", message.clone()),
        XlsxError::Xml(error) => ("Xml", error.to_string()),
        XlsxError::MarkupCompatibility(error) => ("MarkupCompatibility", error.to_string()),
        XlsxError::Package(error) => ("Package", error.to_string()),
        _ => ("Other", error.to_string()),
    }
}

fn run_command(command: &str, args: &[&str]) -> Option<String> {
    std::process::Command::new(command)
        .args(args)
        .output()
        .ok()
        .filter(|output| output.status.success())
        .and_then(|output| String::from_utf8(output.stdout).ok())
        .map(|value| value.trim().to_owned())
        .filter(|value| !value.is_empty())
}

fn runner_identity() -> RunnerIdentity {
    let git_revision = run_command("git", &["rev-parse", "HEAD"]);
    let git_dirty = std::process::Command::new("git")
        .args(["diff", "--quiet"])
        .status()
        .ok()
        .map(|status| !status.success());
    RunnerIdentity {
        git_revision,
        git_dirty,
        rustc_vv: run_command("rustc", &["-Vv"]),
        os: env::consts::OS.to_owned(),
        arch: env::consts::ARCH.to_owned(),
        profile: if cfg!(debug_assertions) {
            "debug"
        } else {
            "release"
        },
    }
}

fn binary_identity() -> AnyResult<BinaryIdentity> {
    let path = env::current_exe()?.canonicalize()?;
    let bytes = fs::read(&path)?;
    Ok(BinaryIdentity {
        path: path.to_string_lossy().into_owned(),
        sha256: sha256(&bytes),
        bytes: u64::try_from(bytes.len())?,
    })
}

pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let arguments = parse_args(args)?;
    let fixture = fixture(arguments.shape, arguments.case)?;
    let selectors = [Selector::from("Sheet1")];
    let source_sha256 = fixture.source_sha256.clone();
    let mut observations = Vec::new();
    observations.try_reserve_exact(arguments.samples)?;

    for _ in 0..arguments.warmup {
        let _ = run_iteration(&fixture, &selectors, arguments.case)?;
    }
    for _ in 0..arguments.samples {
        observations.push(run_iteration(&fixture, &selectors, arguments.case)?);
    }

    if sha256(&fixture.archive) != source_sha256 {
        return Err("source archive changed while running planning guard".into());
    }
    let first_error = observations
        .first()
        .ok_or("planning guard produced no observations")?
        .logical_error
        .clone();
    if observations.iter().any(|observation| {
        observation.logical_error.status != first_error.status
            || observation.logical_error.variant != first_error.variant
            || observation.logical_error.message != first_error.message
    }) {
        return Err("planning guard outcome changed between samples".into());
    }

    let phase_samples = observations
        .iter()
        .enumerate()
        .map(|(order, observation)| PhaseSample {
            order,
            duration_ns: observation.duration_ns,
            allocation_metrics: observation.allocation.clone(),
        })
        .collect::<Vec<_>>();
    let report = Report {
        schema: "litchi.xlsx.planning-refusal-guard.v1",
        tool: "xlsx_planning_guard",
        case: arguments.case.name(),
        shape: arguments.shape.name(),
        warmup_iterations: arguments.warmup,
        samples: arguments.samples,
        source: SourceIdentity {
            generator: "litchi-xlsx-planning-refusal-guard-fixed-grid-v1",
            shape: arguments.shape.name(),
            rows: fixture.rows,
            columns: fixture.columns,
            archive_bytes: fixture.archive.len(),
            worksheet_bytes: fixture.worksheet.len(),
            source_sha256: fixture.source_sha256.clone(),
            worksheet_sha256: fixture.worksheet_sha256,
            worksheet_member: WORKSHEET_PART,
            compression: "stored",
            fixture_kind: "raw_zip_input_no_opc_authored_xml_validation",
        },
        source_sha256,
        logical_error: first_error,
        phase: PhaseReport {
            name: "edit_sheets",
            sample_order: phase_samples.iter().map(|sample| sample.order).collect(),
            duration_ns: phase_samples
                .iter()
                .map(|sample| sample.duration_ns)
                .collect(),
            allocation_metrics: phase_samples
                .iter()
                .map(|sample| sample.allocation_metrics.clone())
                .collect(),
            samples: phase_samples,
        },
        correctness: Correctness {
            source_unchanged: observations.iter().all(|o| o.source_unchanged),
            retry_preserved_error: observations.iter().all(|o| o.retry_preserved_error),
            valid_snapshot_values: observations.iter().all(|o| o.valid_snapshot_values),
            empty_commit_is_noop: observations.iter().all(|o| o.empty_commit_is_noop),
            expected_error_exact: observations.iter().all(|o| o.expected_error_exact),
            commit_outside_timing: observations.iter().all(|o| o.commit_outside_timing),
        },
        binary: binary_identity()?,
        runner: runner_identity(),
        allocation_scope: "operation_global_system_allocator",
        allocator: allocation_metrics::allocator_identity(),
        instrumentation: allocation_metrics::instrumentation_identity(),
        counter_revision: allocation_metrics::counter_revision(),
        timing_scope: "source_open_and_fixture_and_selector_setup_excluded; edit_sheets_result_retained_until_clock_and_allocation_end; inspection_commit_and_drop_excluded",
        performance_claim: "none: correctness and planning diagnostic only; no speedup claim",
    };
    let encoded = serde_json::to_vec_pretty(&report)?;
    if let Some(path) = arguments.json_path {
        fs::write(path, &encoded)?;
    }
    println!("{}", String::from_utf8(encoded)?);
    Ok(())
}
