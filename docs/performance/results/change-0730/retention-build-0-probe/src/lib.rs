//! Candidate-only bounded DOC retained-render ownership probe.
//!
//! The public API is intentionally exercised without the main timing probe's
//! diagnostics feature. Each process measures one retention choice and route,
//! then compares its output bytes directly with a zero-retention reference.

use litchi_core::Position;
use litchi_doc::body_text::{Snapshot, TransactionLimits};
use litchi_doc::tracked_revision::Limits;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::path::PathBuf;

pub mod alloc_metrics;

pub type BoxError = Box<dyn Error>;

const DEFAULT_RETAINED_RENDER_BYTES: usize = 8 * 1024 * 1024;
const FIRST_TEXT: &str = "litchi copy-through baseline replacement text";
const SECOND_TEXT: &str = "litchi retained render second replacement text";

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

#[derive(Clone, Copy, Debug)]
enum Case {
    DocFloat,
    DocNoHf,
}

impl Case {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "docfloat" => Ok(Self::DocFloat),
            "docnohf" => Ok(Self::DocNoHf),
            other => failure(format!("unknown --case {other:?}")),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::DocFloat => "docfloat",
            Self::DocNoHf => "docnohf",
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum EditRoute {
    One,
    Two,
}

impl EditRoute {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "one" | "1" => Ok(Self::One),
            "two" | "2" => Ok(Self::Two),
            other => failure(format!("unknown --edits {other:?}")),
        }
    }

    const fn count(self) -> usize {
        match self {
            Self::One => 1,
            Self::Two => 2,
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Retention {
    Zero,
    Default,
    Release,
}

impl Retention {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "zero" => Ok(Self::Zero),
            "default" | "8mib" => Ok(Self::Default),
            "release" => Ok(Self::Release),
            other => failure(format!("unknown --retention {other:?}")),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Zero => "zero",
            Self::Default => "default_8mib",
            Self::Release => "release_8mib",
        }
    }

    const fn ceiling(self) -> usize {
        match self {
            Self::Zero => 0,
            Self::Default | Self::Release => DEFAULT_RETAINED_RENDER_BYTES,
        }
    }

    const fn releases_after_stage(self) -> bool {
        matches!(self, Self::Release)
    }
}

#[derive(Clone, Debug)]
struct Args {
    case: Case,
    input: PathBuf,
    route: EditRoute,
    retention: Retention,
}

fn parse_args() -> Result<Args, BoxError> {
    let mut case = None;
    let mut input = None;
    let mut route = None;
    let mut retention = None;
    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let value = || -> Result<String, BoxError> {
            arguments.next().ok_or_else(|| {
                Box::new(ProbeError(format!("missing value for {flag}"))) as BoxError
            })
        };
        match flag.as_str() {
            "--case" => case = Some(Case::parse(&value()?)?),
            "--input" => input = Some(PathBuf::from(value()?)),
            "--edits" => route = Some(EditRoute::parse(&value()?)?),
            "--retention" => retention = Some(Retention::parse(&value()?)?),
            other => return failure(format!("unknown flag {other:?}")),
        }
    }
    Ok(Args {
        case: case.ok_or_else(|| Box::new(ProbeError("missing --case".into())) as BoxError)?,
        input: input.ok_or_else(|| Box::new(ProbeError("missing --input".into())) as BoxError)?,
        route: route.ok_or_else(|| Box::new(ProbeError("missing --edits".into())) as BoxError)?,
        retention: retention
            .ok_or_else(|| Box::new(ProbeError("missing --retention".into())) as BoxError)?,
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

#[derive(Clone, Copy, Debug, Serialize)]
struct StageObservation {
    stage: usize,
    before: Option<usize>,
    after: Option<usize>,
    after_release: Option<usize>,
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

struct CaseResult {
    output: Vec<u8>,
    stages: [StageObservation; 2],
    final_retained_before_commit: Option<usize>,
}

fn transaction_limits(retention: Retention) -> TransactionLimits {
    TransactionLimits::default().with_max_retained_render_bytes(retention.ceiling())
}

fn run_case(source: &[u8], args: &Args, retention: Retention) -> Result<CaseResult, BoxError> {
    let snapshot = Snapshot::open_bounded(
        source.to_vec(),
        Limits::default(),
        transaction_limits(retention),
    )?;
    let mut edit = snapshot.edit()?;
    let mut stages = [StageObservation {
        stage: 0,
        before: None,
        after: None,
        after_release: None,
    }; 2];
    for (stage, text) in [FIRST_TEXT, SECOND_TEXT]
        .into_iter()
        .enumerate()
        .take(args.route.count())
    {
        let before = edit.retained_render_bytes();
        edit.replace_paragraph(Position::new(0), text)?;
        let after = edit.retained_render_bytes();
        let after_release = if retention.releases_after_stage() {
            edit.release_retained_render();
            edit.retained_render_bytes()
        } else {
            None
        };
        stages[stage] = StageObservation {
            stage: stage + 1,
            before,
            after,
            after_release,
        };
    }
    let final_retained_before_commit = edit.retained_render_bytes();
    let commit = edit.commit()?;
    Ok(CaseResult {
        output: commit.snapshot().bytes().to_vec(),
        stages,
        final_retained_before_commit,
    })
}

#[derive(Serialize)]
struct ProbeOutput {
    schema_version: u32,
    case: &'static str,
    route: &'static str,
    edits: usize,
    retention: &'static str,
    retention_ceiling_bytes: usize,
    allocator_instrumented: bool,
    input: String,
    source_sha256: String,
    source_bytes: usize,
    output_sha256: String,
    output_bytes: usize,
    reference_output_sha256: String,
    direct_output_equal_reference: bool,
    final_retained_before_commit: Option<usize>,
    stages: [StageObservation; 2],
    allocation: Option<AllocationReport>,
}

fn measure(source: &[u8], args: &Args) -> Result<(CaseResult, Option<AllocationReport>), BoxError> {
    if !alloc_metrics::instrumented() {
        return Ok((run_case(source, args, args.retention)?, None));
    }
    let mut result = None;
    let region = alloc_metrics::region(|| {
        result = Some(run_case(source, args, args.retention)?);
        Ok::<(), BoxError>(())
    })?;
    let result =
        result.ok_or_else(|| Box::new(ProbeError("missing measured result".into())) as BoxError)?;
    Ok((result, Some(region.into())))
}

pub fn run(_allocator_lane: bool) -> Result<(), BoxError> {
    let args = parse_args()?;
    let source = std::fs::read(&args.input)?;
    let source_sha256 = sha256_hex(&source);
    let (measured, allocation) = measure(&source, &args)?;
    let reference = run_case(&source, &args, Retention::Zero)?;
    let direct_output_equal_reference = measured.output == reference.output;
    if !direct_output_equal_reference {
        return failure("retention choice changed the committed output bytes");
    }
    let result = ProbeOutput {
        schema_version: 1,
        case: args.case.name(),
        route: match args.route {
            EditRoute::One => "one",
            EditRoute::Two => "two",
        },
        edits: args.route.count(),
        retention: args.retention.name(),
        retention_ceiling_bytes: args.retention.ceiling(),
        allocator_instrumented: alloc_metrics::instrumented(),
        input: args.input.display().to_string(),
        source_sha256,
        source_bytes: source.len(),
        output_sha256: sha256_hex(&measured.output),
        output_bytes: measured.output.len(),
        reference_output_sha256: sha256_hex(&reference.output),
        direct_output_equal_reference,
        final_retained_before_commit: measured.final_retained_before_commit,
        stages: measured.stages,
        allocation,
    };
    println!("{}", serde_json::to_string(&result)?);
    Ok(())
}
