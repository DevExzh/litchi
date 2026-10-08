//! Matched native/allocator controls for the 0967 revision-scan candidate.
//!
//! The fixed input is loaded and checked before warmups. The only operation
//! inside a measured region is the selected public DOC lifecycle. Output
//! validation, hashing, and JSON reporting happen after that region returns.

use std::error::Error;
use std::fmt::{Display, Formatter};
use std::fs;
use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::time::Instant;

use litchi_core::Position;
use litchi_doc::Package;
use litchi_doc::body_text::{Projection, Snapshot, TransactionLimits};
use litchi_doc::tracked_revision::Limits;
use serde::Serialize;
use sha2::{Digest, Sha256};

pub mod alloc_metrics;

pub type BoxError = Box<dyn Error>;

const SCHEMA: &str = "litchi.performance.0967.control_profile.v1";
const DEFAULT_WARMUPS: usize = 5;
const DEFAULT_SAMPLES: usize = 80;
const ALLOC_WARMUPS: usize = 0;
const ALLOC_SAMPLES: usize = 1;
const MAX_WARMUPS: usize = 10_000;
const MAX_SAMPLES: usize = 10_000;
const RETAINED_RENDER_BYTES: usize = 8 * 1024 * 1024;
const FIRST_TEXT: &str = "plain-target";
const SECOND_TEXT: &str = "plain-second";

const DENSE1024_BYTES: usize = 93_696;
const DENSE1024_SHA256: &str = "f29dc1e3e1c05dcf37ce0843551a21ca9423cb283d54599813b21d52fcd76def";
const FALSE_MARKER_BYTES: usize = 15_360;
const FALSE_MARKER_SHA256: &str =
    "67cc08d9c17b3e8336327b51ecd78e6f44b14929037e877559b451918159f4ee";
const REAL_POSITIVE_BYTES: usize = 66_048;
const REAL_POSITIVE_SHA256: &str =
    "f8c19b5712e12244af5e243ab6b996af6bc32193dad208948a155c4a4d5bb3dd";
const REAL_POSITIVE_REVISIONS: usize = 20;

#[derive(Debug)]
struct ProbeError(String);

impl Display for ProbeError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for ProbeError {}

fn failure<T>(message: impl Into<String>) -> Result<T, BoxError> {
    Err(Box::new(ProbeError(message.into())))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Control {
    GeneratedPositive,
    Dense1024,
    FalseMarker,
    RealPositive,
}

impl Control {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "generated-positive" => Ok(Self::GeneratedPositive),
            "dense1024" => Ok(Self::Dense1024),
            "false-marker" => Ok(Self::FalseMarker),
            "real-positive" => Ok(Self::RealPositive),
            other => failure(format!(
                "unknown --control {other:?}; expected generated-positive, dense1024, false-marker, or real-positive"
            )),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::GeneratedPositive => "generated-positive",
            Self::Dense1024 => "dense1024",
            Self::FalseMarker => "false-marker",
            Self::RealPositive => "real-positive",
        }
    }

    const fn scope(self) -> &'static str {
        match self {
            Self::GeneratedPositive | Self::Dense1024 | Self::FalseMarker => "edit",
            Self::RealPositive => "query",
        }
    }

    const fn expected_identity(self) -> (usize, &'static str) {
        match self {
            Self::GeneratedPositive => (
                15_360,
                "47e6f76319a8086ebc36719bc9f1d894ad907a02856cfb842ac198d7879d8737",
            ),
            Self::Dense1024 => (DENSE1024_BYTES, DENSE1024_SHA256),
            Self::FalseMarker => (FALSE_MARKER_BYTES, FALSE_MARKER_SHA256),
            Self::RealPositive => (REAL_POSITIVE_BYTES, REAL_POSITIVE_SHA256),
        }
    }

    const fn expected_revisions(self) -> usize {
        match self {
            Self::GeneratedPositive => 2,
            Self::Dense1024 => 1_024,
            Self::FalseMarker => 1,
            Self::RealPositive => REAL_POSITIVE_REVISIONS,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Retention {
    Default,
}

impl Retention {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "default" => Ok(Self::Default),
            other => failure(format!(
                "unsupported --retention {other:?}; expected default"
            )),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Default => "default_8mib",
        }
    }
}

#[derive(Clone, Debug)]
struct Args {
    variant: String,
    control: Control,
    input: PathBuf,
    edits: Option<usize>,
    retention: Option<Retention>,
    warmups: usize,
    samples: usize,
}

fn parse_count(flag: &str, value: &str, maximum: usize) -> Result<usize, BoxError> {
    let count = value.parse::<usize>().map_err(|error| {
        Box::new(ProbeError(format!(
            "invalid {flag} value {value:?}: {error}"
        ))) as BoxError
    })?;
    if count > maximum {
        return failure(format!("{flag} exceeds bound {maximum}"));
    }
    Ok(count)
}

fn parse_args(allocator_lane: bool) -> Result<Args, BoxError> {
    let mut variant = None;
    let mut control = None;
    let mut input = None;
    let mut edits = None;
    let mut retention = None;
    let mut warmups = if allocator_lane {
        ALLOC_WARMUPS
    } else {
        DEFAULT_WARMUPS
    };
    let mut samples = if allocator_lane {
        ALLOC_SAMPLES
    } else {
        DEFAULT_SAMPLES
    };
    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut value = || -> Result<String, BoxError> {
            arguments.next().ok_or_else(|| {
                Box::new(ProbeError(format!("missing value for {flag}"))) as BoxError
            })
        };
        match flag.as_str() {
            "--variant" => variant = Some(value()?),
            "--control" => control = Some(Control::parse(&value()?)?),
            "--input" => input = Some(PathBuf::from(value()?)),
            "--edits" => {
                let count = value()?.parse::<usize>().map_err(|error| {
                    Box::new(ProbeError(format!("invalid --edits value: {error}"))) as BoxError
                })?;
                if !matches!(count, 1 | 32) {
                    return failure("--edits must be 1 or 32");
                }
                edits = Some(count);
            }
            "--retention" => retention = Some(Retention::parse(&value()?)?),
            "--warmup" => warmups = parse_count("--warmup", &value()?, MAX_WARMUPS)?,
            "--samples" => samples = parse_count("--samples", &value()?, MAX_SAMPLES)?,
            other => return failure(format!("unknown flag {other:?}")),
        }
    }
    let variant =
        variant.ok_or_else(|| Box::new(ProbeError("missing --variant".to_owned())) as BoxError)?;
    if !matches!(variant.as_str(), "baseline" | "candidate") {
        return failure("--variant must be baseline or candidate");
    }
    let control =
        control.ok_or_else(|| Box::new(ProbeError("missing --control".to_owned())) as BoxError)?;
    let input =
        input.ok_or_else(|| Box::new(ProbeError("missing --input".to_owned())) as BoxError)?;
    match control {
        Control::GeneratedPositive | Control::Dense1024 | Control::FalseMarker => {
            if edits.is_none() {
                return failure("edit controls require --edits 1 or 32");
            }
            if retention != Some(Retention::Default) {
                return failure("edit controls require --retention default");
            }
        }
        Control::RealPositive => {
            if edits.is_some() || retention.is_some() {
                return failure("real-positive query does not accept edit or retention flags");
            }
        }
    }
    if allocator_lane && (warmups != ALLOC_WARMUPS || samples != ALLOC_SAMPLES) {
        return failure("allocator lane requires --warmup 0 --samples 1");
    }
    Ok(Args {
        variant,
        control,
        input,
        edits,
        retention,
        warmups,
        samples,
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn stage_text(stage: usize) -> &'static str {
    if stage.is_multiple_of(2) {
        FIRST_TEXT
    } else {
        SECOND_TEXT
    }
}

fn transaction_limits() -> TransactionLimits {
    TransactionLimits::default().with_max_retained_render_bytes(RETAINED_RENDER_BYTES)
}

#[derive(Debug)]
struct EditResult {
    output: Vec<u8>,
    logical_vec_capacity_bytes: usize,
    semantic_change_count: usize,
}

#[derive(Debug)]
struct QueryResult {
    output: Vec<u8>,
    revision_count: usize,
    paragraph_count: usize,
    public_reopen_succeeded: bool,
}

#[derive(Debug)]
enum OperationResult {
    Edit(EditResult),
    Query(QueryResult),
}

fn run_edit_case(source: &[u8], args: &Args) -> Result<OperationResult, BoxError> {
    let edits = args.edits.expect("edit count validated");
    let snapshot =
        Snapshot::open_bounded(source.to_vec(), Limits::default(), transaction_limits())?;
    let mut edit = snapshot.edit()?;
    for stage in 0..edits {
        edit.replace_paragraph(Position::new(0), stage_text(stage))?;
    }
    let commit = edit.commit()?;
    let output = commit.snapshot().bytes().to_vec();
    let logical_vec_capacity_bytes = output.capacity();
    let semantic_change_count = commit.patch().changes().len();
    Ok(OperationResult::Edit(EditResult {
        output,
        logical_vec_capacity_bytes,
        semantic_change_count,
    }))
}

fn public_reader_reopens(bytes: &[u8]) -> Result<bool, BoxError> {
    let mut package = Package::from_reader(Cursor::new(bytes.to_vec()))?;
    package.document()?;
    Ok(true)
}

fn run_query_case(source: &[u8], _args: &Args) -> Result<OperationResult, BoxError> {
    let snapshot = Snapshot::open_bounded(
        source.to_vec(),
        Limits::default(),
        TransactionLimits::default(),
    )?;
    let revision_count = snapshot.revisions()?.len();
    let paragraph_count = snapshot.paragraphs(Projection::All)?.len();
    let output = snapshot.bytes().to_vec();
    let public_reopen_succeeded = public_reader_reopens(&output)?;
    Ok(OperationResult::Query(QueryResult {
        output,
        revision_count,
        paragraph_count,
        public_reopen_succeeded,
    }))
}

fn run_operation(source: &[u8], args: &Args) -> Result<OperationResult, BoxError> {
    match args.control.scope() {
        "edit" => run_edit_case(source, args),
        "query" => run_query_case(source, args),
        _ => failure("unknown control scope"),
    }
}

/// The only native/allocator measured closure. Do not add hashing, JSON,
/// fixture generation, source checks, or correctness readers here.
#[inline(never)]
fn profiled_measured_case(source: &[u8], args: &Args) -> Result<OperationResult, BoxError> {
    let result = run_operation(source, args)?;
    match &result {
        OperationResult::Edit(edit) => {
            black_box(edit.logical_vec_capacity_bytes);
        }
        OperationResult::Query(query) => {
            black_box(query.output.len());
        }
    }
    Ok(result)
}

#[derive(Clone, Copy, Debug, Serialize)]
struct AllocationReport {
    allocated_bytes: u64,
    deallocated_bytes: u64,
    allocation_calls: u64,
    peak_live_bytes: u64,
    retained_bytes: i64,
}

impl From<alloc_metrics::Region> for AllocationReport {
    fn from(region: alloc_metrics::Region) -> Self {
        Self {
            allocated_bytes: region.allocated_bytes,
            deallocated_bytes: region.deallocated_bytes,
            allocation_calls: region.allocation_calls,
            peak_live_bytes: region.peak_live_bytes,
            retained_bytes: region.retained_bytes,
        }
    }
}

struct MeasuredResult {
    operation: OperationResult,
    elapsed_ns: Option<u128>,
    allocation: Option<AllocationReport>,
}

fn measure_one(source: &[u8], args: &Args) -> Result<MeasuredResult, BoxError> {
    if alloc_metrics::instrumented() {
        let mut operation = None;
        let region = alloc_metrics::region(|| {
            operation = Some(profiled_measured_case(source, args)?);
            Ok::<(), BoxError>(())
        })?;
        let operation = operation.ok_or_else(|| {
            Box::new(ProbeError("measured operation was not returned".to_owned())) as BoxError
        })?;
        Ok(MeasuredResult {
            operation,
            elapsed_ns: None,
            allocation: Some(region.into()),
        })
    } else {
        let start = Instant::now();
        let operation = profiled_measured_case(source, args)?;
        let elapsed_ns = start.elapsed().as_nanos();
        black_box(operation_output_len(&operation));
        Ok(MeasuredResult {
            operation,
            elapsed_ns: Some(elapsed_ns),
            allocation: None,
        })
    }
}

fn operation_output(operation: &OperationResult) -> &[u8] {
    match operation {
        OperationResult::Edit(edit) => &edit.output,
        OperationResult::Query(query) => &query.output,
    }
}

fn operation_output_len(operation: &OperationResult) -> usize {
    operation_output(operation).len()
}

fn operation_logical_vec_capacity_bytes(operation: &OperationResult) -> usize {
    match operation {
        OperationResult::Edit(edit) => edit.logical_vec_capacity_bytes,
        OperationResult::Query(query) => query.output.capacity(),
    }
}

#[derive(Clone, Debug)]
struct Reference {
    output: Vec<u8>,
    output_sha256: String,
    output_bytes: usize,
    logical_vec_capacity_bytes: usize,
    revision_count: Option<usize>,
    paragraph_count: Option<usize>,
    semantic_change_count: Option<usize>,
}

fn prepare_reference(source: &[u8], args: &Args) -> Result<Reference, BoxError> {
    let operation = run_operation(source, args)?;
    let output = operation_output(&operation).to_vec();
    let (revision_count, paragraph_count, semantic_change_count) = match &operation {
        OperationResult::Edit(edit) => {
            let snapshot =
                Snapshot::open_bounded(output.clone(), Limits::default(), transaction_limits())?;
            let revisions = snapshot.revisions()?.len();
            let paragraphs = snapshot.paragraphs(Projection::All)?.len();
            if !public_reader_reopens(&output)? {
                return failure("reference edit output did not reopen publicly");
            }
            (
                Some(revisions),
                Some(paragraphs),
                Some(edit.semantic_change_count),
            )
        }
        OperationResult::Query(query) => {
            if !query.public_reopen_succeeded || query.output != source {
                return failure("reference query did not preserve source bytes and reopen");
            }
            (
                Some(query.revision_count),
                Some(query.paragraph_count),
                None,
            )
        }
    };
    Ok(Reference {
        output_sha256: sha256_hex(&output),
        output_bytes: output.len(),
        logical_vec_capacity_bytes: operation_logical_vec_capacity_bytes(&operation),
        output,
        revision_count,
        paragraph_count,
        semantic_change_count,
    })
}

#[derive(Clone, Debug, Serialize)]
struct SampleReport {
    index: usize,
    elapsed_ns: Option<u128>,
    output_sha256: String,
    output_bytes: usize,
    logical_vec_capacity_bytes: usize,
    exact_output_equal_reference: bool,
    revision_count: Option<usize>,
    paragraph_count: Option<usize>,
    semantic_change_count: Option<usize>,
    paragraph_zero_sha256: Option<String>,
    public_reopen_succeeded: bool,
    source_unchanged: bool,
    allocation: Option<AllocationReport>,
}

fn sample_report(
    source: &[u8],
    index: usize,
    measured: MeasuredResult,
    args: &Args,
    reference: &Reference,
) -> Result<SampleReport, BoxError> {
    let output = operation_output(&measured.operation);
    let exact_output_equal_reference = output == reference.output;
    let output_sha256 = sha256_hex(output);
    let output_bytes = output.len();
    let logical_vec_capacity_bytes = operation_logical_vec_capacity_bytes(&measured.operation);
    let mut paragraph_zero_sha256 = None;
    let mut public_reopen_succeeded = false;
    let mut source_unchanged = false;
    let (revision_count, paragraph_count, semantic_change_count) = match &measured.operation {
        OperationResult::Edit(edit) => {
            let snapshot =
                Snapshot::open_bounded(output.to_vec(), Limits::default(), transaction_limits())?;
            let paragraphs = snapshot.paragraphs(Projection::All)?;
            let paragraph_zero = paragraphs
                .iter()
                .find(|paragraph| paragraph.position() == Position::new(0))
                .ok_or_else(|| {
                    Box::new(ProbeError("paragraph zero is unavailable".to_owned())) as BoxError
                })?;
            paragraph_zero_sha256 = Some(sha256_hex(paragraph_zero.text().as_bytes()));
            let revisions = snapshot.revisions()?.len();
            public_reopen_succeeded = public_reader_reopens(output)?;
            (
                Some(revisions),
                Some(paragraphs.len()),
                Some(edit.semantic_change_count),
            )
        }
        OperationResult::Query(query) => {
            source_unchanged = output == source;
            (
                Some(query.revision_count),
                Some(query.paragraph_count),
                None,
            )
        }
    };
    let valid = match args.control.scope() {
        "edit" => {
            let edits = args.edits.expect("validated edit count");
            let expected_text = stage_text(edits - 1);
            let snapshot =
                Snapshot::open_bounded(output.to_vec(), Limits::default(), transaction_limits())?;
            let paragraph_zero = snapshot
                .paragraphs(Projection::All)?
                .into_iter()
                .find(|paragraph| paragraph.position() == Position::new(0))
                .is_some_and(|paragraph| paragraph.text() == expected_text);
            exact_output_equal_reference
                && paragraph_zero
                && public_reopen_succeeded
                && revision_count == reference.revision_count
                && paragraph_count == reference.paragraph_count
                && semantic_change_count == Some(edits)
        }
        "query" => {
            exact_output_equal_reference
                && source_unchanged
                && public_reopen_succeeded_or_query(&measured.operation)
                && revision_count == Some(args.control.expected_revisions())
        }
        _ => false,
    };
    if !valid {
        return failure(format!(
            "control sample {index} failed semantic or output gate"
        ));
    }
    Ok(SampleReport {
        index,
        elapsed_ns: measured.elapsed_ns,
        output_sha256,
        output_bytes,
        logical_vec_capacity_bytes,
        exact_output_equal_reference,
        revision_count,
        paragraph_count,
        semantic_change_count,
        paragraph_zero_sha256,
        public_reopen_succeeded: if args.control.scope() == "query" {
            public_reopen_succeeded_or_query(&measured.operation)
        } else {
            public_reopen_succeeded
        },
        source_unchanged,
        allocation: measured.allocation,
    })
}

fn public_reopen_succeeded_or_query(operation: &OperationResult) -> bool {
    match operation {
        OperationResult::Query(query) => query.public_reopen_succeeded,
        OperationResult::Edit(_) => false,
    }
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema_version: u32,
    schema: &'static str,
    lane: &'static str,
    variant: String,
    control: &'static str,
    scope: &'static str,
    input: String,
    source_sha256: String,
    source_bytes: usize,
    edits: Option<usize>,
    retention: Option<&'static str>,
    timing_claim: bool,
    allocator_instrumented: bool,
    allocator_timing_claim: bool,
    region_definition: &'static str,
    warmups: usize,
    samples_requested: usize,
    reference_output_sha256: String,
    reference_output_bytes: usize,
    reference_logical_vec_capacity_bytes: usize,
    reference_revision_count: Option<usize>,
    reference_paragraph_count: Option<usize>,
    reference_semantic_change_count: Option<usize>,
    elapsed_ns: Vec<u128>,
    samples: Vec<SampleReport>,
    allocation: Option<AllocationReport>,
}

fn verify_input(args: &Args, source: &[u8]) -> Result<(), BoxError> {
    let (expected_bytes, expected_hash) = args.control.expected_identity();
    let actual_hash = sha256_hex(source);
    if source.len() != expected_bytes || actual_hash != expected_hash {
        return failure(format!(
            "{} input identity mismatch: got {} bytes / {}, expected {} bytes / {}",
            args.control.name(),
            source.len(),
            actual_hash,
            expected_bytes,
            expected_hash
        ));
    }
    Ok(())
}

pub fn run(allocator_lane: bool) -> Result<(), BoxError> {
    let args = parse_args(allocator_lane)?;
    let source = fs::read(&args.input)?;
    verify_input(&args, &source)?;
    let source_sha256 = sha256_hex(&source);

    // Reference preparation is deliberately outside every warmup and sample.
    let reference = prepare_reference(&source, &args)?;
    let mut samples = Vec::with_capacity(args.samples);
    for _ in 0..args.warmups {
        drop(measure_one(&source, &args)?);
    }
    for index in 0..args.samples {
        let measured = measure_one(&source, &args)?;
        samples.push(sample_report(&source, index, measured, &args, &reference)?);
    }
    let elapsed_ns = samples
        .iter()
        .filter_map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    let allocation = samples.last().and_then(|sample| sample.allocation);
    let output = Report {
        schema_version: 1,
        schema: SCHEMA,
        lane: if alloc_metrics::instrumented() {
            "allocator"
        } else {
            "native"
        },
        variant: args.variant,
        control: args.control.name(),
        scope: args.control.scope(),
        input: args.input.display().to_string(),
        source_sha256,
        source_bytes: source.len(),
        edits: args.edits,
        retention: args.retention.map(Retention::name),
        timing_claim: !alloc_metrics::instrumented(),
        allocator_instrumented: alloc_metrics::instrumented(),
        allocator_timing_claim: false,
        region_definition: match args.control {
            Control::GeneratedPositive | Control::Dense1024 | Control::FalseMarker => {
                "public Snapshot::open_bounded source clone, one Edit, all paragraph-0 replacements, commit, output Vec copy, and local snapshot/edit/commit destruction; input read, source checks, reference preparation, correctness readers, hashing, and JSON excluded"
            }
            Control::RealPositive => {
                "public Snapshot::open_bounded source clone, revisions query, paragraph query, output Vec copy, and public Package reopen; input read, source checks, reference preparation, correctness readers, hashing, and JSON excluded"
            }
        },
        warmups: args.warmups,
        samples_requested: args.samples,
        reference_output_sha256: reference.output_sha256,
        reference_output_bytes: reference.output_bytes,
        reference_logical_vec_capacity_bytes: reference.logical_vec_capacity_bytes,
        reference_revision_count: reference.revision_count,
        reference_paragraph_count: reference.paragraph_count,
        reference_semantic_change_count: reference.semantic_change_count,
        elapsed_ns,
        samples,
        allocation,
    };
    println!("{}", serde_json::to_string(&output)?);
    Ok(())
}
