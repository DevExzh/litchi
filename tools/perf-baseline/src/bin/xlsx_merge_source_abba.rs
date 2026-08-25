//! Same-revision ABBA evidence for the guarded XLSX merge editor.
//!
//! The driver deliberately starts a fresh child for every leg.  It measures
//! only commit plus publication in the child and keeps process creation,
//! fixture construction, and the semantic/reversibility oracles outside that
//! primary interval.  This is a diagnostic path, not a cross-revision claim.

use std::cmp::Ordering;
use std::collections::BTreeMap;
use std::env;
use std::error::Error;
use std::fs;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering as AtomicOrdering},
};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter};
use litchi_xlsx::workbook::{SourceBackedMergeEditor, Workbook};
use litchi_xlsx::{Cell, Error as XlsxError, Value};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

type AnyError = Box<dyn Error>;
type AnyResult<T> = Result<T, AnyError>;

const SCHEMA: &str = "litchi.xlsx.source-merge-abba.v1";
const SINK_LIMIT: usize = 64 * 1024;
const WORKSHEET_PART: &str = "/xl/worksheets/sheet1.xml";
const GATE_PERCENT: f64 = 5.0;
const EVIDENCE_MIN_SAMPLES: usize = 20;
const EVIDENCE_PERFORMANCE_CLAIM: &str =
    "same-revision eager-vs-source path comparison only; no cross-revision evidence";
const NO_PERFORMANCE_CLAIM: &str =
    "none: correctness and diagnostic timing only; no performance claim";
const SMOKE_PERFORMANCE_CLAIM: &str =
    "none: explicit smoke correctness run; timing gates are diagnostic only";
const EXPECTED_MERGE_ARCHIVE_SHA256: Option<&str> =
    Some("151fed9651e6f88a1e7e17183c8dac1f4885b6a922756214295b0d7c828a589e");
const EXPECTED_MERGE_TARGET_SHA256: Option<&str> =
    Some("692d1a1b71bd6af7bffc8d008d28ba75c5068dc913e0452363950fbe09d5b605");
const EXPECTED_UNMERGE_ARCHIVE_SHA256: Option<&str> =
    Some("6329afb234f9f1ea073e37baa4ca9ab6a0bb559fd40ff46a94320d416296a03f");
const EXPECTED_UNMERGE_TARGET_SHA256: Option<&str> =
    Some("467a33b4a4635f43d4d1a582ed8f31390a0eb6b71c94fc13059de9e5c436798e");
static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

#[derive(Clone, Copy, Debug, Deserialize, Eq, Ord, PartialEq, PartialOrd)]
enum Implementation {
    Eager,
    Source,
}

impl Implementation {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Eager => "eager",
            Self::Source => "source",
        }
    }

    fn parse(value: &str) -> AnyResult<Self> {
        match value {
            "eager" => Ok(Self::Eager),
            "source" => Ok(Self::Source),
            other => Err(format!("unknown implementation '{other}'").into()),
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Eq, Ord, PartialEq, PartialOrd)]
enum Operation {
    Merge,
    Unmerge,
}

impl Operation {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Merge => "merge",
            Self::Unmerge => "unmerge",
        }
    }

    fn parse(value: &str) -> AnyResult<Self> {
        match value {
            "merge" => Ok(Self::Merge),
            "unmerge" => Ok(Self::Unmerge),
            other => Err(format!("unknown operation '{other}'").into()),
        }
    }
}

#[derive(Clone, Copy)]
struct Leg {
    name: &'static str,
    implementation: Implementation,
}

const LEGS: [Leg; 4] = [
    Leg {
        name: "A1",
        implementation: Implementation::Eager,
    },
    Leg {
        name: "B1",
        implementation: Implementation::Source,
    },
    Leg {
        name: "B2",
        implementation: Implementation::Source,
    },
    Leg {
        name: "A2",
        implementation: Implementation::Eager,
    },
];

#[derive(Debug)]
enum Arguments {
    Child {
        implementation: Implementation,
        operation: Operation,
    },
    Driver {
        warmup: usize,
        samples: usize,
        out_dir: PathBuf,
    },
    PrintFixtureIdentity,
}

#[derive(Debug, Default)]
struct CountingCounters {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    len_calls: AtomicU64,
    version_calls: AtomicU64,
}

#[derive(Clone, Copy, Debug, Default, Deserialize, Serialize)]
struct SourceMetrics {
    read_calls: u64,
    read_bytes: u64,
    len_calls: u64,
    version_calls: u64,
}

impl CountingCounters {
    fn snapshot(&self) -> SourceMetrics {
        SourceMetrics {
            read_calls: self.read_calls.load(AtomicOrdering::Relaxed),
            read_bytes: self.read_bytes.load(AtomicOrdering::Relaxed),
            len_calls: self.len_calls.load(AtomicOrdering::Relaxed),
            version_calls: self.version_calls.load(AtomicOrdering::Relaxed),
        }
    }
}

impl SourceMetrics {
    fn checked_delta(self, before: Self) -> Option<Self> {
        Some(Self {
            read_calls: self.read_calls.checked_sub(before.read_calls)?,
            read_bytes: self.read_bytes.checked_sub(before.read_bytes)?,
            len_calls: self.len_calls.checked_sub(before.len_calls)?,
            version_calls: self.version_calls.checked_sub(before.version_calls)?,
        })
    }
}

#[derive(Debug)]
struct CountingSource {
    bytes: Arc<[u8]>,
    counters: Arc<CountingCounters>,
    source_id: u64,
    revision: AtomicU64,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::from(bytes),
            counters: Arc::new(CountingCounters::default()),
            source_id: NEXT_SOURCE_ID.fetch_add(1, AtomicOrdering::Relaxed),
            revision: AtomicU64::new(0),
        }
    }

    fn counters(&self) -> SourceMetrics {
        self.counters.snapshot()
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, AtomicOrdering::Relaxed);
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        self.counters
            .len_calls
            .fetch_add(1, AtomicOrdering::Relaxed);
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("fixture exceeds u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let count = self
            .bytes
            .get(start..)
            .map_or(0, |remaining| remaining.len().min(output.len()));
        if count != 0 {
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
        }
        self.counters
            .read_calls
            .fetch_add(1, AtomicOrdering::Relaxed);
        self.counters.read_bytes.fetch_add(
            u64::try_from(count).unwrap_or(u64::MAX),
            AtomicOrdering::Relaxed,
        );
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.counters
            .version_calls
            .fetch_add(1, AtomicOrdering::Relaxed);
        Ok(SourceVersion::new(
            self.source_id,
            self.revision.load(AtomicOrdering::Relaxed),
        ))
    }
}

#[derive(Debug)]
struct RetainingSink {
    bytes: Vec<u8>,
    capacity: usize,
    write_calls: u64,
    write_bytes: u64,
    largest_write: usize,
}

impl RetainingSink {
    fn new(expected_output: usize) -> AnyResult<Self> {
        let capacity = expected_output
            .checked_mul(2)
            .and_then(|value| value.checked_add(SINK_LIMIT))
            .ok_or("bounded sink capacity overflows usize")?;
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(capacity)?;
        Ok(Self {
            bytes,
            capacity,
            write_calls: 0,
            write_bytes: 0,
            largest_write: 0,
        })
    }

    fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    fn capacity(&self) -> usize {
        self.capacity
    }

    fn write_calls(&self) -> u64 {
        self.write_calls
    }

    fn write_bytes(&self) -> u64 {
        self.write_bytes
    }

    fn largest_write(&self) -> usize {
        self.largest_write
    }
}

impl Write for RetainingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.len() > SINK_LIMIT {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "one benchmark sink write exceeds 64 KiB",
            ));
        }
        let remaining = self.capacity.saturating_sub(self.bytes.len());
        if bytes.len() > remaining {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "bounded benchmark sink exhausted",
            ));
        }
        self.bytes.extend_from_slice(bytes);
        self.write_calls = self.write_calls.saturating_add(1);
        self.write_bytes = self
            .write_bytes
            .saturating_add(u64::try_from(bytes.len()).unwrap_or(u64::MAX));
        self.largest_write = self.largest_write.max(bytes.len());
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone, Debug, Deserialize, Serialize)]
struct SemanticCheck {
    reopened: bool,
    expected_merge: bool,
    observed_merge: bool,
    anchor_text_preserved: bool,
    unrelated_text_preserved: bool,
    merge_followers_covered: bool,
    unmerge_followers_missing: bool,
    c2_missing: bool,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
struct PatchChecks {
    forward_applied: Option<bool>,
    forward_exact_one_part_overlay: Option<bool>,
    inverse_restored_complete_package: Option<bool>,
    stale_refused: Option<bool>,
    foreign_refused: Option<bool>,
}

impl PatchChecks {
    fn unavailable() -> Self {
        Self {
            forward_applied: None,
            forward_exact_one_part_overlay: None,
            inverse_restored_complete_package: None,
            stale_refused: None,
            foreign_refused: None,
        }
    }
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
struct SampleIdentity {
    executable_sha256: String,
    git_revision: String,
    git_dirty: bool,
    rustc_vv: Option<String>,
    os: String,
    arch: String,
    cpu: Option<String>,
    memory: Option<String>,
    rustflags: Option<String>,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
struct SampleRecord {
    schema: String,
    implementation: String,
    operation: String,
    identity: SampleIdentity,
    fixture_sha256: String,
    fixture_target_sha256: String,
    fixture_bytes: u64,
    primary_ns: u64,
    preparation_ns: u64,
    lifecycle_ns: u64,
    semantic_reopen_ns: u64,
    output_bytes: u64,
    output_sha256: String,
    sink_capacity_bytes: u64,
    retained_output_bytes: u64,
    sink_write_calls: u64,
    sink_write_bytes: u64,
    sink_largest_write_bytes: u64,
    source_preparation: Option<SourceMetrics>,
    source_primary: Option<SourceMetrics>,
    semantic: SemanticCheck,
    untouched_part_preserved: Option<bool>,
    patch: PatchChecks,
    stale_refused: Option<bool>,
    materialized_source_bytes: Option<u64>,
    materialized_selected_part_bytes: Option<u64>,
}

#[derive(Clone, Debug, Serialize)]
struct ObservedSample {
    sample: usize,
    leg: String,
    #[serde(flatten)]
    record: SampleRecord,
}

#[derive(Debug, Serialize)]
struct FailureRecord {
    sample: Option<usize>,
    warmup: bool,
    leg: String,
    implementation: String,
    operation: String,
    error: String,
}

#[derive(Debug, Serialize)]
struct Protocol {
    schema: &'static str,
    tool: &'static str,
    package_version: &'static str,
    revision_scope: &'static str,
    directional_claim_scope: &'static str,
    performance_claim: &'static str,
    mode: &'static str,
    evidence_min_samples: usize,
    statistical_gates_enforced: bool,
    warmup: usize,
    samples: usize,
    fresh_child_per_leg: bool,
    leg_order: [&'static str; 4],
    operations: [&'static str; 2],
    sink_limit_bytes: usize,
    fixture: FixtureIdentity,
    identity: SampleIdentity,
    timing: TimingIdentity,
    gates: GatePolicy,
    eager_durable_patch_available: bool,
    durable_patch_limitation: &'static str,
    materialization_claim: bool,
}

#[derive(Clone, Debug, Serialize)]
struct FixtureIdentity {
    merge_sha256: String,
    merge_target_sha256: String,
    merge_bytes: u64,
    unmerge_sha256: String,
    unmerge_target_sha256: String,
    unmerge_bytes: u64,
    sheet: &'static str,
    populated_cells: [&'static str; 2],
    merge_range: &'static str,
}

#[derive(Debug, Serialize)]
struct TimingIdentity {
    primary: &'static str,
    preparation: &'static str,
    lifecycle: &'static str,
    source_counters: &'static str,
}

#[derive(Debug, Serialize)]
struct GatePolicy {
    same_side_symmetric_delta_threshold_percent: f64,
    statistics: [&'static str; 4],
    directional_improvement_threshold_percent: f64,
    directional_improvement_threshold_ns: u64,
    directional_adverse_threshold_percent: f64,
    cross_side_deltas_are_descriptive: bool,
}

#[derive(Debug, Serialize)]
struct Summary {
    schema: &'static str,
    tool: &'static str,
    package_version: &'static str,
    performance_claim: &'static str,
    mode: &'static str,
    evidence_min_samples: usize,
    statistical_gates_enforced: bool,
    cross_revision_evidence: bool,
    warmup: usize,
    requested_samples: usize,
    successful_rows: usize,
    failure_rows: usize,
    operations: Vec<OperationSummary>,
}

#[derive(Debug, Serialize)]
struct OperationSummary {
    operation: String,
    requested_samples: usize,
    successful_samples: usize,
    eager: EagerSummary,
    source: SourceSummary,
    pairwise_sample_drift_diagnostic: PairwiseSampleDriftDiagnostic,
    eager_vs_source_delta_percent: DeltaSummary,
    same_side_gates: SameSideGateSummary,
    directional_gates: DirectionalGateSummary,
    gates: Vec<GateResult>,
    all_required_gates_passed: bool,
}

#[derive(Debug, Serialize)]
struct EagerSummary {
    combined: Option<Statistics>,
    a1: Option<Statistics>,
    a2: Option<Statistics>,
}

#[derive(Debug, Serialize)]
struct SourceSummary {
    combined: Option<Statistics>,
    b1: Option<Statistics>,
    b2: Option<Statistics>,
}

#[derive(Debug, Serialize)]
struct Statistics {
    count: usize,
    min_ns: u64,
    p50_ns: u64,
    mean_ns_floor: u64,
    p95_ns: u64,
    p99_ns: u64,
    max_ns: u64,
}

#[derive(Debug, Serialize)]
struct PairwiseSampleDriftDiagnostic {
    basis: &'static str,
    eager: Option<PercentStatistics>,
    source: Option<PercentStatistics>,
}

#[derive(Debug, Serialize)]
struct PercentStatistics {
    count: usize,
    min_percent: f64,
    p50_percent: f64,
    mean_percent: f64,
    p95_percent: f64,
    p99_percent: f64,
    max_percent: f64,
}

#[derive(Debug, Serialize)]
struct DeltaSummary {
    p50_percent: Option<f64>,
    mean_percent: Option<f64>,
    p95_percent: Option<f64>,
    p99_percent: Option<f64>,
}

#[derive(Debug, Serialize)]
struct SameSideGateSummary {
    basis: &'static str,
    eager_a1_a2: Vec<GateResult>,
    source_b1_b2: Vec<GateResult>,
}

#[derive(Debug, Serialize)]
struct DirectionalGateSummary {
    eager_a1_to_source_b1: Vec<GateResult>,
    eager_a2_to_source_b2: Vec<GateResult>,
}

#[derive(Clone, Debug, Serialize)]
struct GateResult {
    name: String,
    threshold_percent: f64,
    observed_percent: Option<f64>,
    absolute_delta_ns: Option<u64>,
    passed: bool,
    rationale: &'static str,
}

#[derive(Debug, Serialize)]
struct ArtifactManifest {
    schema: &'static str,
    self_excluded: bool,
    files: Vec<ArtifactEntry>,
}

#[derive(Debug, Serialize)]
struct ArtifactEntry {
    path: String,
    bytes: Option<u64>,
    sha256: Option<String>,
}

fn main() {
    match parse_arguments().and_then(|arguments| match arguments {
        Arguments::Child {
            implementation,
            operation,
        } => child_main(implementation, operation),
        Arguments::Driver {
            warmup,
            samples,
            out_dir,
        } => driver_main(warmup, samples, &out_dir),
        Arguments::PrintFixtureIdentity => print_fixture_identity_main(),
    }) {
        Ok(()) => {},
        Err(error) => {
            eprintln!("xlsx_merge_source_abba: {error}");
            std::process::exit(1);
        },
    }
}

fn parse_arguments() -> AnyResult<Arguments> {
    let mut arguments = env::args().skip(1).peekable();
    if matches!(
        arguments.peek().map(String::as_str),
        Some("--print-fixture-identity")
    ) {
        arguments.next();
        if arguments.next().is_some() {
            return Err("--print-fixture-identity does not accept additional arguments".into());
        }
        return Ok(Arguments::PrintFixtureIdentity);
    }
    let mut child = false;
    let mut child_implementation = None;
    let mut child_operation = None;
    let mut positional = Vec::new();
    let mut warmup = 3usize;
    let mut samples = 15usize;
    let mut out_dir = PathBuf::from("xlsx-merge-source-abba");

    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--child" => child = true,
            "--implementation" | "--impl" => {
                let value = arguments
                    .next()
                    .ok_or("--implementation requires eager or source")?;
                child_implementation = Some(Implementation::parse(&value)?);
            },
            "--operation" | "--op" => {
                let value = arguments
                    .next()
                    .ok_or("--operation requires merge or unmerge")?;
                child_operation = Some(Operation::parse(&value)?);
            },
            "--warmup" => {
                warmup = arguments
                    .next()
                    .ok_or("--warmup requires a non-negative integer")?
                    .parse()?;
            },
            "--samples" => {
                samples = arguments
                    .next()
                    .ok_or("--samples requires a positive integer")?
                    .parse()?;
            },
            "--out-dir" => {
                out_dir = PathBuf::from(arguments.next().ok_or("--out-dir requires a path")?);
            },
            value if value.starts_with('-') => {
                return Err(format!("unknown argument '{value}'").into());
            },
            value => positional.push(value.to_owned()),
        }
    }

    if child {
        let implementation = child_implementation
            .or_else(|| {
                positional
                    .first()
                    .and_then(|value| Implementation::parse(value).ok())
            })
            .ok_or("--child requires eager or source")?;
        let operation = child_operation
            .or_else(|| {
                positional
                    .get(1)
                    .and_then(|value| Operation::parse(value).ok())
            })
            .ok_or("--child requires merge or unmerge")?;
        return Ok(Arguments::Child {
            implementation,
            operation,
        });
    }

    if !positional.is_empty() {
        return Err(format!("unexpected positional argument '{}", positional[0]).into());
    }
    if samples == 0 {
        return Err("--samples must be at least 1".into());
    }
    if warmup > 100 {
        return Err("--warmup must be at most 100".into());
    }
    if samples > 100_000 {
        return Err("--samples must be at most 100000".into());
    }
    Ok(Arguments::Driver {
        warmup,
        samples,
        out_dir,
    })
}

fn print_fixture_identity_main() -> AnyResult<()> {
    let merge = fixture_for(Operation::Merge)?;
    let unmerge = fixture_for(Operation::Unmerge)?;
    let identity = FixtureIdentity {
        merge_sha256: sha256_hex(&merge),
        merge_target_sha256: fixture_target_sha256(&merge)?,
        merge_bytes: u64::try_from(merge.len())?,
        unmerge_sha256: sha256_hex(&unmerge),
        unmerge_target_sha256: fixture_target_sha256(&unmerge)?,
        unmerge_bytes: u64::try_from(unmerge.len())?,
        sheet: "Sheet1",
        populated_cells: ["A1", "C1"],
        merge_range: "A1:B2",
    };
    println!("{}", serde_json::to_string_pretty(&identity)?);
    Ok(())
}

fn child_main(implementation: Implementation, operation: Operation) -> AnyResult<()> {
    let mut record = match implementation {
        Implementation::Eager => run_eager(operation)?,
        Implementation::Source => run_source(operation)?,
    };
    record.identity = sample_identity()?;
    println!("{}", serde_json::to_string(&record)?);
    Ok(())
}

fn sample_identity_placeholder() -> SampleIdentity {
    SampleIdentity {
        executable_sha256: String::new(),
        git_revision: String::new(),
        git_dirty: false,
        rustc_vv: None,
        os: String::new(),
        arch: String::new(),
        cpu: None,
        memory: None,
        rustflags: None,
    }
}

fn sample_identity() -> AnyResult<SampleIdentity> {
    let executable = env::current_exe()?;
    let revision =
        command_text("git", &["rev-parse", "HEAD"]).unwrap_or_else(|| "unavailable".to_owned());
    let dirty = command_text("git", &["status", "--porcelain", "--untracked-files=no"])
        .is_some_and(|value| !value.is_empty());
    Ok(SampleIdentity {
        executable_sha256: sha256_hex(&fs::read(executable)?),
        git_revision: revision,
        git_dirty: dirty,
        rustc_vv: command_text("rustc", &["-Vv"]),
        os: env::consts::OS.to_owned(),
        arch: env::consts::ARCH.to_owned(),
        cpu: env::var("PROCESSOR_IDENTIFIER")
            .ok()
            .or_else(|| env::var("HOSTTYPE").ok()),
        memory: fs::read_to_string("/proc/meminfo").ok().and_then(|value| {
            value
                .lines()
                .find(|line| line.starts_with("MemTotal:"))
                .map(str::to_owned)
        }),
        rustflags: env::var("RUSTFLAGS").ok(),
    })
}

fn command_text(command: &str, arguments: &[&str]) -> Option<String> {
    let output = Command::new(command).args(arguments).output().ok()?;
    output
        .status
        .success()
        .then(|| String::from_utf8_lossy(&output.stdout).trim().to_owned())
}

fn run_eager(operation: Operation) -> AnyResult<SampleRecord> {
    let fixture = fixture_for(operation)?;
    let fixture_sha256 = sha256_hex(&fixture);
    let fixture_target_sha256 = fixture_target_sha256(&fixture)?;
    let lifecycle_started = Instant::now();
    let preparation_started = Instant::now();
    let workbook = Workbook::from_slice(&fixture)?;
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet("Sheet1")?
            .ok_or("eager fixture worksheet disappeared")?;
        stage_eager(&mut sheet, operation)?;
    }
    let preparation_ns = elapsed_ns(preparation_started.elapsed())?;

    let mut sink = RetainingSink::new(fixture.len())?;
    let sink_capacity = sink.capacity();
    let primary_started = Instant::now();
    let committed = edit.commit()?;
    let changed = committed.into_workbook();
    changed.write_to(&mut sink)?;
    let primary_ns = elapsed_ns(primary_started.elapsed())?;
    let sink_write_calls = sink.write_calls();
    let sink_write_bytes = sink.write_bytes();
    let sink_largest_write_bytes = sink.largest_write();
    let output = sink.into_bytes();
    let semantic_started = Instant::now();
    let semantic = semantic_oracle(&output, operation)?;
    let semantic_reopen_ns = elapsed_ns(semantic_started.elapsed())?;
    let untouched_part_preserved = Some(untouched_part_preserved(&fixture, &output)?);
    let lifecycle_ns = elapsed_ns(lifecycle_started.elapsed())?;

    Ok(SampleRecord {
        schema: SCHEMA.to_owned(),
        implementation: Implementation::Eager.as_str().to_owned(),
        operation: operation.as_str().to_owned(),
        identity: sample_identity_placeholder(),
        fixture_sha256,
        fixture_target_sha256,
        fixture_bytes: u64::try_from(fixture.len())?,
        primary_ns,
        preparation_ns,
        lifecycle_ns,
        semantic_reopen_ns,
        output_bytes: u64::try_from(output.len())?,
        output_sha256: sha256_hex(&output),
        sink_capacity_bytes: u64::try_from(sink_capacity)?,
        retained_output_bytes: u64::try_from(output.len())?,
        sink_write_calls,
        sink_write_bytes,
        sink_largest_write_bytes: u64::try_from(sink_largest_write_bytes)?,
        source_preparation: None,
        source_primary: None,
        semantic,
        untouched_part_preserved,
        patch: PatchChecks::unavailable(),
        stale_refused: None,
        materialized_source_bytes: None,
        materialized_selected_part_bytes: None,
    })
}

fn run_source(operation: Operation) -> AnyResult<SampleRecord> {
    let fixture = fixture_for(operation)?;
    let fixture_sha256 = sha256_hex(&fixture);
    let fixture_target_sha256 = fixture_target_sha256(&fixture)?;
    let lifecycle_started = Instant::now();
    let concrete_source = Arc::new(CountingSource::new(fixture.clone()));
    let source: Arc<dyn ReadAt> = concrete_source.clone();
    let preparation_started = Instant::now();
    let editor = SourceBackedMergeEditor::from_read_at(source)?;
    let mut edit = editor
        .edit("Sheet1")?
        .ok_or("source fixture worksheet disappeared")?;
    stage_source(&mut edit, operation)?;
    let preparation_metrics = concrete_source.counters();
    let preparation_ns = elapsed_ns(preparation_started.elapsed())?;

    let mut sink = RetainingSink::new(fixture.len())?;
    let sink_capacity = sink.capacity();
    let primary_started = Instant::now();
    let commit = edit.commit()?;
    let _published = editor.publish_commit_to_stream(&mut sink, &commit)?;
    let primary_ns = elapsed_ns(primary_started.elapsed())?;
    let primary_metrics = concrete_source
        .counters()
        .checked_delta(preparation_metrics)
        .ok_or("CountingSource counter decreased during primary timing")?;
    let sink_write_calls = sink.write_calls();
    let sink_write_bytes = sink.write_bytes();
    let sink_largest_write_bytes = sink.largest_write();
    let output = sink.into_bytes();
    let semantic_started = Instant::now();
    let semantic = semantic_oracle(&output, operation)?;
    let semantic_reopen_ns = elapsed_ns(semantic_started.elapsed())?;
    let untouched_part_preserved = Some(untouched_part_preserved(&fixture, &output)?);
    let mut patch = source_patch_checks(&fixture, &commit)?;
    let stale_refused = Some(stale_refusal(&fixture, operation)?);
    patch.stale_refused = stale_refused;
    let lifecycle_ns = elapsed_ns(lifecycle_started.elapsed())?;

    Ok(SampleRecord {
        schema: SCHEMA.to_owned(),
        implementation: Implementation::Source.as_str().to_owned(),
        operation: operation.as_str().to_owned(),
        identity: sample_identity_placeholder(),
        fixture_sha256,
        fixture_target_sha256,
        fixture_bytes: u64::try_from(fixture.len())?,
        primary_ns,
        preparation_ns,
        lifecycle_ns,
        semantic_reopen_ns,
        output_bytes: u64::try_from(output.len())?,
        output_sha256: sha256_hex(&output),
        sink_capacity_bytes: u64::try_from(sink_capacity)?,
        retained_output_bytes: u64::try_from(output.len())?,
        sink_write_calls,
        sink_write_bytes,
        sink_largest_write_bytes: u64::try_from(sink_largest_write_bytes)?,
        source_preparation: Some(preparation_metrics),
        source_primary: Some(primary_metrics),
        semantic,
        untouched_part_preserved,
        patch,
        stale_refused,
        materialized_source_bytes: None,
        materialized_selected_part_bytes: None,
    })
}

fn stage_eager(
    sheet: &mut litchi_xlsx::workbook::WorksheetEdit<'_>,
    operation: Operation,
) -> AnyResult<()> {
    match operation {
        Operation::Merge => sheet.merge("A1:B2")?,
        Operation::Unmerge => sheet.unmerge("B2")?,
    };
    Ok(())
}

fn stage_source(
    edit: &mut litchi_xlsx::workbook::SourceBackedMergeEdit,
    operation: Operation,
) -> AnyResult<()> {
    match operation {
        Operation::Merge => edit.merge("A1:B2")?,
        Operation::Unmerge => edit.unmerge("B2")?,
    };
    Ok(())
}

fn fixture_for(operation: Operation) -> AnyResult<Vec<u8>> {
    let workbook = Workbook::new()?;
    let mut edit = workbook.edit()?;
    {
        let mut sheet = edit
            .sheet("Sheet1")?
            .ok_or("minimal workbook worksheet disappeared")?;
        sheet.set("A1", "litchi-xlsx-merge-anchor-v1")?;
        sheet.set("C1", "litchi-xlsx-merge-unrelated-v1")?;
        if operation == Operation::Unmerge {
            sheet.merge("A1:B2")?;
        }
    }
    let committed = edit.commit()?.into_workbook();
    let mut output = Vec::new();
    committed.write_to(&mut output)?;
    let mut package = OpcPackage::from_bytes(&output)?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/merge-abba-aux.bin")?,
        "application/octet-stream".to_owned(),
        b"litchi-xlsx-source-merge-aux-v1".to_vec(),
    )))?;
    Ok(PackageWriter::to_bytes(&package)?)
}

fn fixture_target_sha256(bytes: &[u8]) -> AnyResult<String> {
    let package = OpcPackage::from_bytes(bytes)?;
    let uri = PackURI::new(WORKSHEET_PART)?;
    Ok(sha256_hex(package.get_part(&uri)?.blob()))
}

fn semantic_oracle(bytes: &[u8], operation: Operation) -> AnyResult<SemanticCheck> {
    let workbook = Workbook::from_slice(bytes)?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("semantic reopen lost Sheet1")?;
    let expected_merge = operation == Operation::Merge;
    let target = litchi_xlsx::Rect::from_a1("A1:B2")?;
    let ranges = sheet.merges()?.collect::<Vec<_>>();
    let observed_merge = if expected_merge {
        ranges.len() == 1 && ranges[0] == target
    } else {
        ranges.is_empty()
    };
    let anchor_text_preserved = matches!(
        sheet.cell("A1")?.stored(),
        Some(Cell::Value(Value::Text(text)))
            if text.as_str() == "litchi-xlsx-merge-anchor-v1"
    );
    let unrelated_text_preserved = matches!(
        sheet.cell("C1")?.stored(),
        Some(Cell::Value(Value::Text(text)))
            if text.as_str() == "litchi-xlsx-merge-unrelated-v1"
    );
    let merge_followers_covered = ["A2", "B1", "B2"].into_iter().all(|address| {
        sheet
            .cell(address)
            .is_ok_and(|cell| cell.merge() == Some(target))
    });
    let unmerge_followers_missing = ["A2", "B1", "B2"]
        .into_iter()
        .all(|address| sheet.cell(address).is_ok_and(|cell| cell.is_missing()));
    let c2_missing = sheet.cell("C2")?.is_missing();
    Ok(SemanticCheck {
        reopened: true,
        expected_merge,
        observed_merge,
        anchor_text_preserved,
        unrelated_text_preserved,
        merge_followers_covered,
        unmerge_followers_missing,
        c2_missing,
    })
}

#[derive(Clone, Debug, Eq, PartialEq)]
struct PartDigest {
    content_type: String,
    blob_sha256: String,
}

type PartInventory = BTreeMap<String, PartDigest>;

fn package_inventory(bytes: &[u8]) -> AnyResult<PartInventory> {
    let package = OpcPackage::from_bytes(bytes)?;
    let mut inventory = BTreeMap::new();
    for part in package.iter_parts() {
        inventory.insert(
            part.partname().to_string(),
            PartDigest {
                content_type: part.content_type().to_owned(),
                blob_sha256: sha256_hex(part.blob()),
            },
        );
    }
    Ok(inventory)
}

fn untouched_part_preserved(before: &[u8], after: &[u8]) -> AnyResult<bool> {
    exact_worksheet_overlay(&package_inventory(before)?, &package_inventory(after)?)
}

fn exact_worksheet_overlay(before: &PartInventory, after: &PartInventory) -> AnyResult<bool> {
    if before.keys().ne(after.keys()) {
        return Ok(false);
    }
    let changed = before
        .iter()
        .filter(|(name, digest)| after.get(*name) != Some(*digest))
        .map(|(name, _)| name.as_str())
        .collect::<Vec<_>>();
    Ok(changed.len() == 1 && changed[0] == WORKSHEET_PART)
}

fn source_patch_checks(
    fixture: &[u8],
    commit: &litchi_xlsx::workbook::SourceBackedMergeCommit,
) -> AnyResult<PatchChecks> {
    let before_inventory = package_inventory(fixture)?;
    let mut package = OpcPackage::from_bytes(fixture)?;
    let forward_applied = commit.patch().apply(&mut package).is_ok();
    if !forward_applied {
        return Ok(PatchChecks {
            forward_applied: Some(false),
            forward_exact_one_part_overlay: Some(false),
            inverse_restored_complete_package: Some(false),
            stale_refused: None,
            foreign_refused: None,
        });
    }
    let forward_inventory = package_inventory_from_package(&package);
    let forward_exact_one_part_overlay =
        exact_worksheet_overlay(&before_inventory, &forward_inventory)?;
    let inverse_restored_complete_package = commit
        .patch()
        .inverse()
        .apply_materialized(&mut package, SINK_LIMIT)
        .is_ok()
        && package_inventory_from_package(&package) == before_inventory;
    let foreign_refused = foreign_refusal(fixture, commit)?;
    Ok(PatchChecks {
        forward_applied: Some(true),
        forward_exact_one_part_overlay: Some(forward_exact_one_part_overlay),
        inverse_restored_complete_package: Some(inverse_restored_complete_package),
        stale_refused: None,
        foreign_refused: Some(foreign_refused),
    })
}

fn package_inventory_from_package(package: &OpcPackage) -> PartInventory {
    package
        .iter_parts()
        .map(|part| {
            (
                part.partname().to_string(),
                PartDigest {
                    content_type: part.content_type().to_owned(),
                    blob_sha256: sha256_hex(part.blob()),
                },
            )
        })
        .collect()
}

fn stale_refusal(fixture: &[u8], operation: Operation) -> AnyResult<bool> {
    let concrete_source = Arc::new(CountingSource::new(fixture.to_vec()));
    let source: Arc<dyn ReadAt> = concrete_source.clone();
    let editor = SourceBackedMergeEditor::from_read_at(source)?;
    let mut edit = editor
        .edit("Sheet1")?
        .ok_or("stale fixture worksheet disappeared")?;
    stage_source(&mut edit, operation)?;
    let commit = edit.commit()?;
    concrete_source.bump_revision();
    Ok(matches!(
        editor.publish_commit_to_stream(io::sink(), &commit),
        Err(XlsxError::Package(OpcError::SourceChanged { .. }))
    ))
}

fn foreign_refusal(
    fixture: &[u8],
    commit: &litchi_xlsx::workbook::SourceBackedMergeCommit,
) -> AnyResult<bool> {
    let foreign_source = Arc::new(CountingSource::new(fixture.to_vec()));
    let foreign: Arc<dyn ReadAt> = foreign_source;
    let editor = SourceBackedMergeEditor::from_read_at(foreign)?;
    Ok(matches!(
        editor.publish_commit_to_stream(io::sink(), commit),
        Err(XlsxError::PatchConflict { .. })
    ))
}

fn driver_main(warmup: usize, samples: usize, out_dir: &Path) -> AnyResult<()> {
    let process_started = SystemTime::now();
    fs::create_dir_all(out_dir)?;
    let merge_fixture = fixture_for(Operation::Merge)?;
    let unmerge_fixture = fixture_for(Operation::Unmerge)?;
    let fixture = FixtureIdentity {
        merge_sha256: sha256_hex(&merge_fixture),
        merge_target_sha256: fixture_target_sha256(&merge_fixture)?,
        merge_bytes: u64::try_from(merge_fixture.len())?,
        unmerge_sha256: sha256_hex(&unmerge_fixture),
        unmerge_target_sha256: fixture_target_sha256(&unmerge_fixture)?,
        unmerge_bytes: u64::try_from(unmerge_fixture.len())?,
        sheet: "Sheet1",
        populated_cells: ["A1", "C1"],
        merge_range: "A1:B2",
    };
    verify_pinned_fixture_identity(&fixture)?;
    let identity = sample_identity()?;
    let smoke = samples < EVIDENCE_MIN_SAMPLES;
    let mode = if smoke { "smoke" } else { "evidence" };
    let initial_claim = if smoke {
        SMOKE_PERFORMANCE_CLAIM
    } else {
        NO_PERFORMANCE_CLAIM
    };
    let mut protocol = Protocol {
        schema: SCHEMA,
        tool: "xlsx_merge_source_abba",
        package_version: env!("CARGO_PKG_VERSION"),
        revision_scope: "same workspace revision and same executable path for all legs; evidence requires a clean revision",
        directional_claim_scope: "A1->B1 and A2->B2 are claim-bearing directional comparisons within the same-revision eager-vs-source path scope; cross-revision evidence is excluded",
        performance_claim: initial_claim,
        mode,
        evidence_min_samples: EVIDENCE_MIN_SAMPLES,
        statistical_gates_enforced: !smoke,
        warmup,
        samples,
        fresh_child_per_leg: true,
        leg_order: ["A1 eager", "B1 source", "B2 source", "A2 eager"],
        operations: ["merge", "unmerge"],
        sink_limit_bytes: SINK_LIMIT,
        fixture: fixture.clone(),
        identity,
        timing: TimingIdentity {
            primary: "commit plus publication into a bounded retaining 64 KiB sink",
            preparation: "fixture-independent implementation open plus edit staging",
            lifecycle: "child lifecycle through semantic and source-contract diagnostics",
            source_counters: "logical ReadAt, len, and version calls from a local CountingSource",
        },
        gates: GatePolicy {
            same_side_symmetric_delta_threshold_percent: GATE_PERCENT,
            statistics: ["p50", "mean", "p95", "p99"],
            directional_improvement_threshold_percent: 1.0,
            directional_improvement_threshold_ns: 50_000,
            directional_adverse_threshold_percent: GATE_PERCENT,
            cross_side_deltas_are_descriptive: true,
        },
        eager_durable_patch_available: false,
        durable_patch_limitation: "eager durable Patch apply/inverse is not exposed by this runner; no durable-patch claim is made",
        materialization_claim: false,
    };
    write_json(&out_dir.join("protocol.json"), &protocol)?;

    let mut rows = Vec::new();
    let mut failures = Vec::new();
    let expected_rows = samples
        .checked_mul(LEGS.len())
        .and_then(|value| value.checked_mul(2))
        .ok_or("sample row reservation overflows usize")?;
    rows.try_reserve_exact(expected_rows)?;
    let expected_warmup_failures = warmup
        .checked_mul(LEGS.len())
        .and_then(|value| value.checked_mul(2))
        .ok_or("warmup failure reservation overflows usize")?;
    failures.try_reserve(expected_rows.saturating_add(expected_warmup_failures))?;
    for operation in [Operation::Merge, Operation::Unmerge] {
        run_phase(
            operation,
            warmup,
            samples,
            &fixture,
            &protocol.identity,
            &mut rows,
            &mut failures,
        )?;
    }

    let validation = validate_rows(samples, &rows);
    if let Err(error) = validation {
        failures.push(FailureRecord {
            sample: None,
            warmup: false,
            leg: "validation".to_owned(),
            implementation: "driver".to_owned(),
            operation: "merge+unmerge".to_owned(),
            error: error.to_string(),
        });
    }
    write_jsonl(&out_dir.join("samples.jsonl"), &rows)?;
    let diagnostic_summary = build_summary(
        warmup,
        samples,
        &rows,
        failures.len(),
        mode,
        !smoke,
        initial_claim,
    );
    let gates_passed = diagnostic_summary
        .operations
        .iter()
        .all(|operation| operation.all_required_gates_passed);
    if !smoke && !gates_passed {
        failures.push(FailureRecord {
            sample: None,
            warmup: false,
            leg: "gates".to_owned(),
            implementation: "driver".to_owned(),
            operation: "merge+unmerge".to_owned(),
            error: "one or more strict same-side or directional 5% gates failed".to_owned(),
        });
    }
    let evidence_claim =
        !smoke && !protocol.identity.git_dirty && failures.is_empty() && gates_passed;
    let final_claim = if evidence_claim {
        EVIDENCE_PERFORMANCE_CLAIM
    } else if smoke {
        SMOKE_PERFORMANCE_CLAIM
    } else {
        NO_PERFORMANCE_CLAIM
    };
    if protocol.performance_claim != final_claim {
        protocol.performance_claim = final_claim;
        write_json(&out_dir.join("protocol.json"), &protocol)?;
    }
    let summary = build_summary(
        warmup,
        samples,
        &rows,
        failures.len(),
        mode,
        !smoke,
        final_claim,
    );
    write_failure_jsonl(&out_dir.join("failures.jsonl"), &failures)?;
    write_json(&out_dir.join("summary.json"), &summary)?;
    let sha256_text = format!(
        "merge_fixture_archive_sha256 {}\nmerge_fixture_target_sha256 {}\nmerge_fixture_bytes {}\nunmerge_fixture_archive_sha256 {}\nunmerge_fixture_target_sha256 {}\nunmerge_fixture_bytes {}\n",
        fixture.merge_sha256,
        fixture.merge_target_sha256,
        fixture.merge_bytes,
        fixture.unmerge_sha256,
        fixture.unmerge_target_sha256,
        fixture.unmerge_bytes,
    );
    fs::write(out_dir.join("sha256.txt"), sha256_text.as_bytes())?;
    fs::write(
        out_dir.join("process-time.txt"),
        process_time_text(process_started)?,
    )?;
    write_artifact_manifest(out_dir)?;
    if !failures.is_empty() {
        return Err(format!("ABBA run failed with {} failure(s)", failures.len()).into());
    }
    Ok(())
}

fn process_time_text(started: SystemTime) -> AnyResult<String> {
    let start_ms = started.duration_since(UNIX_EPOCH)?.as_millis();
    let finish_ms = SystemTime::now().duration_since(UNIX_EPOCH)?.as_millis();
    Ok(format!(
        "pid={}\nstart_unix_ms={start_ms}\nfinish_unix_ms={finish_ms}\n",
        std::process::id()
    ))
}

fn verify_pinned_fixture_identity(fixture: &FixtureIdentity) -> AnyResult<()> {
    verify_pinned_hash(
        "merge archive",
        &fixture.merge_sha256,
        EXPECTED_MERGE_ARCHIVE_SHA256,
    )?;
    verify_pinned_hash(
        "merge worksheet target",
        &fixture.merge_target_sha256,
        EXPECTED_MERGE_TARGET_SHA256,
    )?;
    verify_pinned_hash(
        "unmerge archive",
        &fixture.unmerge_sha256,
        EXPECTED_UNMERGE_ARCHIVE_SHA256,
    )?;
    verify_pinned_hash(
        "unmerge worksheet target",
        &fixture.unmerge_target_sha256,
        EXPECTED_UNMERGE_TARGET_SHA256,
    )?;
    Ok(())
}

fn verify_pinned_hash(label: &str, observed: &str, expected: Option<&str>) -> AnyResult<()> {
    let expected = expected.ok_or_else(|| {
        format!(
            "{label} identity is an unpinned placeholder; run --print-fixture-identity and pin the printed hash"
        )
    })?;
    if observed != expected {
        return Err(
            format!("{label} hash mismatch: expected {expected}, observed {observed}").into(),
        );
    }
    Ok(())
}

fn run_phase(
    operation: Operation,
    warmup: usize,
    samples: usize,
    fixture: &FixtureIdentity,
    identity: &SampleIdentity,
    rows: &mut Vec<ObservedSample>,
    failures: &mut Vec<FailureRecord>,
) -> AnyResult<()> {
    for _ in 0..warmup {
        for leg in LEGS {
            match invoke_child(leg.implementation, operation) {
                Ok(record) => {
                    if let Err(error) = validate_child_record(
                        &record,
                        leg,
                        operation,
                        expected_fixture_hash(fixture, operation),
                        expected_fixture_target_hash(fixture, operation),
                        identity,
                    ) {
                        failures.push(FailureRecord {
                            sample: None,
                            warmup: true,
                            leg: leg.name.to_owned(),
                            implementation: leg.implementation.as_str().to_owned(),
                            operation: operation.as_str().to_owned(),
                            error: error.to_string(),
                        });
                    }
                },
                Err(error) => failures.push(FailureRecord {
                    sample: None,
                    warmup: true,
                    leg: leg.name.to_owned(),
                    implementation: leg.implementation.as_str().to_owned(),
                    operation: operation.as_str().to_owned(),
                    error: error.to_string(),
                }),
            }
        }
    }
    for sample in 0..samples {
        for leg in LEGS {
            match invoke_child(leg.implementation, operation) {
                Ok(record) => match validate_child_record(
                    &record,
                    leg,
                    operation,
                    expected_fixture_hash(fixture, operation),
                    expected_fixture_target_hash(fixture, operation),
                    identity,
                ) {
                    Ok(()) => rows.push(ObservedSample {
                        sample,
                        leg: leg.name.to_owned(),
                        record,
                    }),
                    Err(error) => failures.push(FailureRecord {
                        sample: Some(sample),
                        warmup: false,
                        leg: leg.name.to_owned(),
                        implementation: leg.implementation.as_str().to_owned(),
                        operation: operation.as_str().to_owned(),
                        error: error.to_string(),
                    }),
                },
                Err(error) => failures.push(FailureRecord {
                    sample: Some(sample),
                    warmup: false,
                    leg: leg.name.to_owned(),
                    implementation: leg.implementation.as_str().to_owned(),
                    operation: operation.as_str().to_owned(),
                    error: error.to_string(),
                }),
            }
        }
    }
    Ok(())
}

fn expected_fixture_hash(fixture: &FixtureIdentity, operation: Operation) -> &str {
    match operation {
        Operation::Merge => &fixture.merge_sha256,
        Operation::Unmerge => &fixture.unmerge_sha256,
    }
}

fn expected_fixture_target_hash(fixture: &FixtureIdentity, operation: Operation) -> &str {
    match operation {
        Operation::Merge => &fixture.merge_target_sha256,
        Operation::Unmerge => &fixture.unmerge_target_sha256,
    }
}

fn validate_child_record(
    record: &SampleRecord,
    leg: Leg,
    operation: Operation,
    expected_archive_sha256: &str,
    expected_target_sha256: &str,
    expected_identity: &SampleIdentity,
) -> AnyResult<()> {
    if record.schema != SCHEMA {
        return Err(format!("child schema mismatch: {}", record.schema).into());
    }
    if record.implementation != leg.implementation.as_str()
        || record.operation != operation.as_str()
    {
        return Err("child implementation or operation identity mismatch".into());
    }
    if record.fixture_sha256 != expected_archive_sha256
        || record.fixture_target_sha256 != expected_target_sha256
    {
        return Err("child fixture archive or target hash mismatch".into());
    }
    if &record.identity != expected_identity {
        return Err("child provenance identity differs from protocol identity".into());
    }
    if !record.semantic.reopened
        || record.semantic.expected_merge != (operation == Operation::Merge)
        || !record.semantic.observed_merge
        || !record.semantic.anchor_text_preserved
        || !record.semantic.unrelated_text_preserved
        || (operation == Operation::Merge && !record.semantic.merge_followers_covered)
        || (operation == Operation::Unmerge && !record.semantic.unmerge_followers_missing)
        || !record.semantic.c2_missing
    {
        return Err("child semantic oracle did not prove the complete fixture state".into());
    }
    if record.untouched_part_preserved != Some(true) {
        return Err("child did not prove exactly one changed worksheet part".into());
    }
    if record.materialized_source_bytes.is_some()
        || record.materialized_selected_part_bytes.is_some()
    {
        return Err("materialization fields must remain explicitly unavailable".into());
    }
    if record.sink_largest_write_bytes > u64::try_from(SINK_LIMIT)?
        || record.output_bytes != record.retained_output_bytes
        || record.sink_write_bytes != record.output_bytes
    {
        return Err("child bounded sink evidence is invalid".into());
    }
    match leg.implementation {
        Implementation::Eager => {
            if record.source_preparation.is_some()
                || record.source_primary.is_some()
                || record.patch.forward_applied.is_some()
                || record.patch.forward_exact_one_part_overlay.is_some()
                || record.patch.inverse_restored_complete_package.is_some()
                || record.patch.stale_refused.is_some()
                || record.patch.foreign_refused.is_some()
            {
                return Err(
                    "eager child reported unavailable source or durable-patch evidence".into(),
                );
            }
        },
        Implementation::Source => {
            if record.source_preparation.is_none()
                || record.source_primary.is_none()
                || record.patch.forward_applied != Some(true)
                || record.patch.forward_exact_one_part_overlay != Some(true)
                || record.patch.inverse_restored_complete_package != Some(true)
                || record.patch.stale_refused != Some(true)
                || record.patch.foreign_refused != Some(true)
                || record.stale_refused != Some(true)
            {
                return Err("source child did not prove patch, stale, and foreign refusals".into());
            }
        },
    }
    Ok(())
}

fn validate_rows(samples: usize, rows: &[ObservedSample]) -> AnyResult<()> {
    let rows_per_operation = samples
        .checked_mul(LEGS.len())
        .ok_or("expected row count overflows usize")?;
    let expected_total = rows_per_operation
        .checked_mul(2)
        .ok_or("expected total row count overflows usize")?;
    if rows.len() != expected_total {
        return Err(format!(
            "expected exactly {expected_total} valid rows, observed {}",
            rows.len()
        )
        .into());
    }
    let mut seen = std::collections::BTreeSet::new();
    for row in rows {
        if row.sample >= samples {
            return Err("row sample index exceeds requested samples".into());
        }
        let key = (
            row.sample,
            row.record.operation.clone(),
            row.leg.clone(),
            row.record.implementation.clone(),
        );
        if !seen.insert(key) {
            return Err("duplicate sample/operation/leg identity".into());
        }
    }
    for operation in [Operation::Merge, Operation::Unmerge] {
        for sample in 0..samples {
            for leg in LEGS {
                let count = rows
                    .iter()
                    .filter(|row| {
                        row.sample == sample
                            && row.record.operation == operation.as_str()
                            && row.leg == leg.name
                            && row.record.implementation == leg.implementation.as_str()
                    })
                    .count();
                if count != 1 {
                    return Err(format!(
                        "missing or duplicate row for sample {sample}, {}, {}, {}",
                        operation.as_str(),
                        leg.name,
                        leg.implementation.as_str()
                    )
                    .into());
                }
            }
        }
    }
    Ok(())
}

fn invoke_child(implementation: Implementation, operation: Operation) -> AnyResult<SampleRecord> {
    let output = Command::new(env::current_exe()?)
        .arg("--child")
        .arg(implementation.as_str())
        .arg(operation.as_str())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!(
            "child {} {} exited with {}: {}",
            implementation.as_str(),
            operation.as_str(),
            output.status,
            stderr.trim()
        )
        .into());
    }
    let stdout = String::from_utf8(output.stdout)?;
    let line = stdout
        .lines()
        .last()
        .ok_or("child returned no JSON record")?;
    Ok(serde_json::from_str(line)?)
}

fn build_summary(
    warmup: usize,
    requested_samples: usize,
    rows: &[ObservedSample],
    failure_rows: usize,
    mode: &'static str,
    statistical_gates_enforced: bool,
    performance_claim: &'static str,
) -> Summary {
    let operations = [Operation::Merge, Operation::Unmerge]
        .into_iter()
        .map(|operation| operation_summary(operation, requested_samples, rows))
        .collect();
    Summary {
        schema: SCHEMA,
        tool: "xlsx_merge_source_abba",
        package_version: env!("CARGO_PKG_VERSION"),
        performance_claim,
        mode,
        evidence_min_samples: EVIDENCE_MIN_SAMPLES,
        statistical_gates_enforced,
        cross_revision_evidence: false,
        warmup,
        requested_samples,
        successful_rows: rows.len(),
        failure_rows,
        operations,
    }
}

fn operation_summary(
    operation: Operation,
    requested_samples: usize,
    rows: &[ObservedSample],
) -> OperationSummary {
    let eager_a1 = leg_values(rows, operation, "A1");
    let eager_a2 = leg_values(rows, operation, "A2");
    let source_b1 = leg_values(rows, operation, "B1");
    let source_b2 = leg_values(rows, operation, "B2");
    let eager_combined = eager_a1
        .iter()
        .chain(eager_a2.iter())
        .copied()
        .collect::<Vec<_>>();
    let source_combined = source_b1
        .iter()
        .chain(source_b2.iter())
        .copied()
        .collect::<Vec<_>>();
    let eager_drift = paired_drift(rows, operation, "A1", "A2");
    let source_drift = paired_drift(rows, operation, "B1", "B2");
    let eager_stats = statistics(&eager_combined);
    let source_stats = statistics(&source_combined);
    let eager_a1_stats = statistics(&eager_a1);
    let eager_a2_stats = statistics(&eager_a2);
    let source_b1_stats = statistics(&source_b1);
    let source_b2_stats = statistics(&source_b2);
    let deltas = DeltaSummary {
        p50_percent: paired_stat_delta(eager_stats.as_ref(), source_stats.as_ref(), |s| s.p50_ns),
        mean_percent: paired_stat_delta(eager_stats.as_ref(), source_stats.as_ref(), |s| {
            s.mean_ns_floor
        }),
        p95_percent: paired_stat_delta(eager_stats.as_ref(), source_stats.as_ref(), |s| s.p95_ns),
        p99_percent: paired_stat_delta(eager_stats.as_ref(), source_stats.as_ref(), |s| s.p99_ns),
    };
    let eager_same_side = same_side_gates(
        "eager A1/A2",
        eager_a1_stats.as_ref(),
        eager_a2_stats.as_ref(),
    );
    let source_same_side = same_side_gates(
        "source B1/B2",
        source_b1_stats.as_ref(),
        source_b2_stats.as_ref(),
    );
    let a1_to_b1 = directional_gates(
        "A1 eager -> B1 source",
        eager_a1_stats.as_ref(),
        source_b1_stats.as_ref(),
    );
    let a2_to_b2 = directional_gates(
        "A2 eager -> B2 source",
        eager_a2_stats.as_ref(),
        source_b2_stats.as_ref(),
    );
    let mut gates = Vec::new();
    gates.extend(eager_same_side.iter().cloned());
    gates.extend(source_same_side.iter().cloned());
    gates.extend(a1_to_b1.iter().cloned());
    gates.extend(a2_to_b2.iter().cloned());
    let successful_samples = rows
        .iter()
        .filter(|row| row.record.operation == operation.as_str())
        .map(|row| row.sample)
        .collect::<std::collections::BTreeSet<_>>()
        .len();
    let all_required_gates_passed = gates.iter().all(|result| result.passed);
    OperationSummary {
        operation: operation.as_str().to_owned(),
        requested_samples,
        successful_samples,
        eager: EagerSummary {
            combined: eager_stats,
            a1: eager_a1_stats,
            a2: eager_a2_stats,
        },
        source: SourceSummary {
            combined: source_stats,
            b1: source_b1_stats,
            b2: source_b2_stats,
        },
        pairwise_sample_drift_diagnostic: PairwiseSampleDriftDiagnostic {
            basis: "diagnostic only: per-sample paired drift is not used by any gate",
            eager: eager_drift,
            source: source_drift,
        },
        eager_vs_source_delta_percent: deltas,
        same_side_gates: SameSideGateSummary {
            basis: "gates use symmetric deltas of per-leg aggregate p50/mean/p95/p99 values; per-leg values are reported in eager/source statistics",
            eager_a1_a2: eager_same_side,
            source_b1_b2: source_same_side,
        },
        directional_gates: DirectionalGateSummary {
            eager_a1_to_source_b1: a1_to_b1,
            eager_a2_to_source_b2: a2_to_b2,
        },
        gates,
        all_required_gates_passed,
    }
}

fn leg_values(rows: &[ObservedSample], operation: Operation, leg: &str) -> Vec<u64> {
    rows.iter()
        .filter(|row| row.record.operation == operation.as_str() && row.leg == leg)
        .map(|row| row.record.primary_ns)
        .collect()
}

fn paired_drift(
    rows: &[ObservedSample],
    operation: Operation,
    first_leg: &str,
    second_leg: &str,
) -> Option<PercentStatistics> {
    let first = rows
        .iter()
        .filter(|row| row.record.operation == operation.as_str() && row.leg == first_leg)
        .map(|row| (row.sample, row.record.primary_ns))
        .collect::<BTreeMap<_, _>>();
    let second = rows
        .iter()
        .filter(|row| row.record.operation == operation.as_str() && row.leg == second_leg)
        .map(|row| (row.sample, row.record.primary_ns))
        .collect::<BTreeMap<_, _>>();
    let values = first
        .into_iter()
        .filter_map(|(sample, left)| {
            second
                .get(&sample)
                .and_then(|right| relative_drift(left, *right))
        })
        .collect::<Vec<_>>();
    percent_statistics(&values)
}

fn relative_drift(left: u64, right: u64) -> Option<f64> {
    if left == 0 && right == 0 {
        return Some(0.0);
    }
    let denominator = left.min(right);
    (denominator != 0).then(|| (left.abs_diff(right) as f64 / denominator as f64) * 100.0)
}

fn paired_stat_delta(
    eager: Option<&Statistics>,
    source: Option<&Statistics>,
    select: impl Fn(&Statistics) -> u64,
) -> Option<f64> {
    match (eager, source) {
        (Some(eager), Some(source)) => relative_change(select(eager), select(source)),
        _ => None,
    }
}

fn relative_change(eager: u64, source: u64) -> Option<f64> {
    if eager == 0 && source == 0 {
        return Some(0.0);
    }
    (eager != 0).then(|| ((source as f64 / eager as f64) - 1.0) * 100.0)
}

fn statistics(values: &[u64]) -> Option<Statistics> {
    if values.is_empty() {
        return None;
    }
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    let sum = sorted.iter().fold(0u128, |total, value| {
        total.saturating_add(u128::from(*value))
    });
    Some(Statistics {
        count: sorted.len(),
        min_ns: sorted[0],
        p50_ns: floor_midpoint(sorted[(sorted.len() - 1) / 2], sorted[sorted.len() / 2]),
        mean_ns_floor: u64::try_from(sum / u128::try_from(sorted.len()).unwrap_or(u128::MAX))
            .unwrap_or(u64::MAX),
        p95_ns: nearest_rank(&sorted, 95),
        p99_ns: nearest_rank(&sorted, 99),
        max_ns: sorted[sorted.len() - 1],
    })
}

fn percent_statistics(values: &[f64]) -> Option<PercentStatistics> {
    if values.is_empty() {
        return None;
    }
    let mut sorted = values.to_vec();
    sorted.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
    let sum = sorted.iter().sum::<f64>();
    Some(PercentStatistics {
        count: sorted.len(),
        min_percent: sorted[0],
        p50_percent: floor_midpoint_f64(sorted[(sorted.len() - 1) / 2], sorted[sorted.len() / 2]),
        mean_percent: sum / sorted.len() as f64,
        p95_percent: nearest_rank_f64(&sorted, 95),
        p99_percent: nearest_rank_f64(&sorted, 99),
        max_percent: sorted[sorted.len() - 1],
    })
}

fn nearest_rank(values: &[u64], percentile: usize) -> u64 {
    let rank = percentile
        .saturating_mul(values.len())
        .saturating_add(99)
        .saturating_div(100)
        .max(1);
    values[rank.saturating_sub(1).min(values.len() - 1)]
}

fn nearest_rank_f64(values: &[f64], percentile: usize) -> f64 {
    let rank = percentile
        .saturating_mul(values.len())
        .saturating_add(99)
        .saturating_div(100)
        .max(1);
    values[rank.saturating_sub(1).min(values.len() - 1)]
}

fn floor_midpoint(left: u64, right: u64) -> u64 {
    left / 2 + right / 2 + (left % 2 + right % 2) / 2
}

fn floor_midpoint_f64(left: f64, right: f64) -> f64 {
    (left + right).floor() / 2.0
}

fn same_side_gates(
    prefix: &str,
    first: Option<&Statistics>,
    second: Option<&Statistics>,
) -> Vec<GateResult> {
    ["p50", "mean", "p95", "p99"]
        .into_iter()
        .map(|metric| {
            let (observed_percent, absolute_delta_ns) = match (first, second) {
                (Some(first), Some(second)) => {
                    let left = metric_value(first, metric);
                    let right = metric_value(second, metric);
                    (relative_drift(left, right), Some(left.abs_diff(right)))
                },
                _ => (None, None),
            };
            GateResult {
                name: format!("{prefix} {metric} symmetric delta"),
                threshold_percent: GATE_PERCENT,
                observed_percent,
                absolute_delta_ns,
                passed: observed_percent.is_some_and(|value| value <= GATE_PERCENT),
                rationale: "symmetric delta of per-leg aggregate statistics must be <=5%; paired-sample drift is diagnostic only",
            }
        })
        .collect()
}

fn directional_gates(
    prefix: &str,
    eager: Option<&Statistics>,
    source: Option<&Statistics>,
) -> Vec<GateResult> {
    [
        ("p50", true),
        ("mean", true),
        ("p95", false),
        ("p99", false),
    ]
    .into_iter()
    .map(|(metric, requires_improvement)| {
        let (observed_percent, absolute_delta_ns, passed) = match (eager, source) {
            (Some(eager), Some(source)) => {
                let eager_value = metric_value(eager, metric);
                let source_value = metric_value(source, metric);
                let improvement = directional_improvement(eager_value, source_value);
                let absolute_delta_ns = Some(eager_value.abs_diff(source_value));
                let passed = improvement.is_some_and(|value| {
                    if requires_improvement {
                        source_value < eager_value
                            && (value >= 1.0 || eager_value.abs_diff(source_value) >= 50_000)
                    } else {
                        (-value).max(0.0) <= GATE_PERCENT
                    }
                });
                (improvement, absolute_delta_ns, passed)
            },
            _ => (None, None, false),
        };
        GateResult {
            name: format!("{prefix} {metric}"),
            threshold_percent: if requires_improvement {
                1.0
            } else {
                GATE_PERCENT
            },
            observed_percent,
            absolute_delta_ns,
            passed,
            rationale: if requires_improvement {
                "source must improve by >=1% or >=50 us for p50/mean"
            } else {
                "source may not be more than 5% adverse for p95/p99"
            },
        }
    })
    .collect()
}

fn metric_value(statistics: &Statistics, metric: &str) -> u64 {
    match metric {
        "p50" => statistics.p50_ns,
        "mean" => statistics.mean_ns_floor,
        "p95" => statistics.p95_ns,
        "p99" => statistics.p99_ns,
        _ => 0,
    }
}

fn directional_improvement(eager: u64, source: u64) -> Option<f64> {
    if eager == 0 && source == 0 {
        return Some(0.0);
    }
    (eager != 0).then(|| ((eager as f64 - source as f64) / eager as f64) * 100.0)
}

fn write_json<T: Serialize>(path: &Path, value: &T) -> AnyResult<()> {
    fs::write(path, serde_json::to_vec_pretty(value)?)?;
    Ok(())
}

fn write_jsonl(path: &Path, rows: &[ObservedSample]) -> AnyResult<()> {
    let mut output = String::new();
    for row in rows {
        let line = serde_json::to_string(row)?;
        output.try_reserve(line.len().saturating_add(1))?;
        output.push_str(&line);
        output.push('\n');
    }
    fs::write(path, output.as_bytes())?;
    Ok(())
}

fn write_failure_jsonl(path: &Path, failures: &[FailureRecord]) -> AnyResult<()> {
    let mut output = String::new();
    for failure in failures {
        let line = serde_json::to_string(failure)?;
        output.try_reserve(line.len().saturating_add(1))?;
        output.push_str(&line);
        output.push('\n');
    }
    fs::write(path, output.as_bytes())?;
    Ok(())
}

fn write_artifact_manifest(out_dir: &Path) -> AnyResult<()> {
    let names = [
        "protocol.json",
        "samples.jsonl",
        "summary.json",
        "failures.jsonl",
        "sha256.txt",
        "process-time.txt",
    ];
    let mut files = Vec::new();
    for name in names {
        let bytes = fs::read(out_dir.join(name))?;
        files.push(ArtifactEntry {
            path: name.to_owned(),
            bytes: Some(u64::try_from(bytes.len())?),
            sha256: Some(sha256_hex(&bytes)),
        });
    }
    write_json(
        &out_dir.join("artifact-manifest.json"),
        &ArtifactManifest {
            schema: SCHEMA,
            self_excluded: true,
            files,
        },
    )
}

fn elapsed_ns(duration: std::time::Duration) -> AnyResult<u64> {
    u64::try_from(duration.as_nanos()).map_err(|_| "duration exceeds u64 nanoseconds".into())
}

fn sha256_hex(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push(HEX[usize::from(byte >> 4)] as char);
        output.push(HEX[usize::from(byte & 0x0f)] as char);
    }
    output
}
