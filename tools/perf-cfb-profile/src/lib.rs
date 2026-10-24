//! Current-source repeated-edit DOC retention harness for the 0942 CFB profile.
//!
//! This crate is an evidence harness. It does not add production code or
//! call a private optimization seam. The public lifecycle under observation
//! is `Snapshot::open_bounded` -> one `Edit` with repeated paragraph-zero
//! replacements -> `commit` -> output `Vec` copy. The ordinary binary reports
//! elapsed clocks for that complete lifecycle. The allocator binary uses the
//! same lifecycle with the counting system allocator and makes no timing
//! claim. The measured path alone goes through `profiled_measured_case`, a
//! non-inlined callchain marker for operation-local native profiling.

use std::error::Error;
use std::fmt::{Display, Formatter};
use std::hint::black_box;
use std::io::Cursor;
use std::path::PathBuf;
use std::sync::Arc;
use std::time::Instant;

use litchi_core::Position;
use litchi_doc::body_text::{Error as BodyError, Projection, Refusal, TransactionLimits};
use litchi_doc::raw::Reader;
use litchi_doc::tracked_revision::Limits;
use serde::Serialize;
use sha2::{Digest, Sha256};

mod oracle_0730;

/// The package route this harness measures (record 0359a): the body-text
/// engine (`body_text::Snapshot`) is crate-private, so the historical
/// `Snapshot::open_bounded -> Edit -> commit` shape runs over
/// `Package::from_bytes` + `edit_with_options` with the same revision and
/// transaction limits.
#[derive(Clone)]
pub(crate) struct Snapshot {
    package: litchi_doc::Package,
    options: litchi_doc::EditOptions,
}

impl Snapshot {
    pub(crate) fn open_bounded(
        bytes: Vec<u8>,
        limits: Limits,
        transaction_limits: TransactionLimits,
    ) -> Result<Self, BoxError> {
        Ok(Self {
            package: litchi_doc::Package::from_bytes(&bytes)?,
            options: litchi_doc::EditOptions::default()
                .revision_limits(limits)
                .transaction_limits(transaction_limits),
        })
    }

    pub(crate) fn open(bytes: Vec<u8>, limits: Limits) -> Result<Self, BoxError> {
        Self::open_bounded(bytes, limits, TransactionLimits::default())
    }

    pub(crate) fn edit(&self) -> Result<litchi_doc::Edit<'static>, BoxError> {
        Ok(self.package.edit_with_options(self.options.clone())?)
    }

    pub(crate) fn paragraphs(
        &self,
        projection: Projection,
    ) -> Result<Vec<litchi_doc::body_text::Paragraph>, BoxError> {
        Ok(self.package.body_paragraphs(projection)?)
    }

    pub(crate) fn package(&self) -> &litchi_doc::Package {
        &self.package
    }

    pub(crate) fn bytes(&self) -> Vec<u8> {
        package_bytes(&self.package)
    }
}

/// A package's exact bytes (the harness compares artifacts by value).
pub(crate) fn package_bytes(package: &litchi_doc::Package) -> Vec<u8> {
    package.to_bytes().expect("package bytes")
}

pub mod alloc_metrics;

pub type BoxError = Box<dyn Error>;

const MAX_EDIT_STAGES: usize = 32;
const DEFAULT_RETAINED_RENDER_BYTES: usize = 8 * 1024 * 1024;
const DEFAULT_WARMUPS: usize = 3;
const DEFAULT_SAMPLES: usize = 30;
const MAX_WARMUPS: usize = 10_000;
const MAX_SAMPLES: usize = 10_000;

// These are intentionally the exact two texts retained by the 0730 probe.
const FIRST_TEXT: &str = "litchi copy-through baseline replacement text";
const SECOND_TEXT: &str = "litchi retained render second replacement text";

const NO_HEAD_FOOT_SHA256: &str =
    "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48";
const FLOATING_PICTURES_SHA256: &str =
    "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d";

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

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
enum Case {
    #[serde(rename = "docnohf")]
    DocNoHf,
    #[serde(rename = "docfloat")]
    DocFloat,
}

impl Case {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "docnohf" => Ok(Self::DocNoHf),
            "docfloat" => Ok(Self::DocFloat),
            other => failure(format!("unknown --case {other:?}")),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::DocNoHf => "docnohf",
            Self::DocFloat => "docfloat",
        }
    }

    const fn expected_bytes(self) -> usize {
        match self {
            Self::DocNoHf => 26_112,
            Self::DocFloat => 335_360,
        }
    }

    const fn expected_sha256(self) -> &'static str {
        match self {
            Self::DocNoHf => NO_HEAD_FOOT_SHA256,
            Self::DocFloat => FLOATING_PICTURES_SHA256,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
enum EditCount {
    #[serde(rename = "1")]
    One,
    #[serde(rename = "2")]
    Two,
    #[serde(rename = "8")]
    Eight,
    #[serde(rename = "32")]
    ThirtyTwo,
}

impl EditCount {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "1" => Ok(Self::One),
            "2" => Ok(Self::Two),
            "8" => Ok(Self::Eight),
            "32" => Ok(Self::ThirtyTwo),
            other => failure(format!(
                "unknown --edits {other:?}; expected 1, 2, 8, or 32"
            )),
        }
    }

    const fn count(self) -> usize {
        match self {
            Self::One => 1,
            Self::Two => 2,
            Self::Eight => 8,
            Self::ThirtyTwo => 32,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
enum Retention {
    #[serde(rename = "zero")]
    Zero,
    #[serde(rename = "default_8mib")]
    Default,
    #[serde(rename = "release_8mib")]
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
    edits: EditCount,
    retention: Retention,
    warmup: usize,
    samples: usize,
}

fn parse_count(flag: &str, value: &str, maximum: usize) -> Result<usize, BoxError> {
    let count = value.parse::<usize>().map_err(|error| {
        Box::new(ProbeError(format!(
            "invalid {flag} value {value:?}: {error}"
        ))) as BoxError
    })?;
    if count > maximum {
        return failure(format!("{flag} exceeds probe bound {maximum}"));
    }
    Ok(count)
}

fn parse_args() -> Result<Args, BoxError> {
    let mut case = None;
    let mut input = None;
    let mut edits = None;
    let mut retention = None;
    let mut warmup = DEFAULT_WARMUPS;
    let mut samples = DEFAULT_SAMPLES;
    let mut arguments = std::env::args().skip(1);
    while let Some(flag) = arguments.next() {
        let mut value = || -> Result<String, BoxError> {
            arguments.next().ok_or_else(|| {
                Box::new(ProbeError(format!("missing value for {flag}"))) as BoxError
            })
        };
        match flag.as_str() {
            "--case" => case = Some(Case::parse(&value()?)?),
            "--input" => input = Some(PathBuf::from(value()?)),
            "--edits" => edits = Some(EditCount::parse(&value()?)?),
            "--retention" => retention = Some(Retention::parse(&value()?)?),
            "--warmup" => warmup = parse_count("--warmup", &value()?, MAX_WARMUPS)?,
            "--samples" => samples = parse_count("--samples", &value()?, MAX_SAMPLES)?,
            other => return failure(format!("unknown flag {other:?}")),
        }
    }
    Ok(Args {
        case: case.ok_or_else(|| Box::new(ProbeError("missing --case".into())) as BoxError)?,
        input: input.ok_or_else(|| Box::new(ProbeError("missing --input".into())) as BoxError)?,
        edits: edits.ok_or_else(|| Box::new(ProbeError("missing --edits".into())) as BoxError)?,
        retention: retention
            .ok_or_else(|| Box::new(ProbeError("missing --retention".into())) as BoxError)?,
        warmup,
        samples,
    })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}

const fn stage_text(stage: usize) -> &'static str {
    if stage.is_multiple_of(2) {
        FIRST_TEXT
    } else {
        SECOND_TEXT
    }
}

const fn final_text(edits: EditCount) -> &'static str {
    stage_text(edits.count() - 1)
}

#[derive(Clone, Copy, Debug, Serialize)]
struct StageObservation {
    stage: usize,
    requested_text: &'static str,
    before: Option<usize>,
    after: Option<usize>,
    after_release: Option<usize>,
}

impl StageObservation {
    const fn empty() -> Self {
        Self {
            stage: 0,
            requested_text: "",
            before: None,
            after: None,
            after_release: None,
        }
    }
}

#[derive(Debug)]
struct CaseResult {
    output: Vec<u8>,
    output_capacity: usize,
    semantic_change_count: usize,
    stages: [StageObservation; MAX_EDIT_STAGES],
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
    let mut stages = [StageObservation::empty(); MAX_EDIT_STAGES];
    for (stage, observation) in stages.iter_mut().enumerate().take(args.edits.count()) {
        let requested_text = stage_text(stage);
        let before = edit.retained_render_bytes();
        edit.replace_paragraph(Position::new(0), requested_text)?;
        let after = edit.retained_render_bytes();
        let after_release = if retention.releases_after_stage() {
            edit.release_retained_render()?;
            edit.retained_render_bytes()
        } else {
            None
        };
        *observation = StageObservation {
            stage: stage + 1,
            requested_text,
            before,
            after,
            after_release,
        };
    }
    let final_retained_before_commit = edit.retained_render_bytes();
    let commit = edit.commit()?;
    let semantic_change_count = commit.patch().changes().len();
    let output = package_bytes(commit.package());
    let output_capacity = output.capacity();
    black_box(output.len());
    Ok(CaseResult {
        output,
        output_capacity,
        semantic_change_count,
        stages,
        final_retained_before_commit,
    })
}

/// Keeps a unique, stable callchain marker around the measured lifecycle.
///
/// Preparation, correctness routes, and per-sample validation call
/// `run_case` directly. The post-call barrier prevents the marker from
/// collapsing into a tail call while preserving the measured result.
#[inline(never)]
fn profiled_measured_case(
    source: &[u8],
    args: &Args,
    retention: Retention,
) -> Result<CaseResult, BoxError> {
    let result = run_case(source, args, retention)?;
    black_box(result.output_capacity);
    Ok(result)
}

struct TimedResult {
    result: CaseResult,
    elapsed_ns: u128,
}

fn timed_case(source: &[u8], args: &Args) -> Result<TimedResult, BoxError> {
    let start = Instant::now();
    let result = profiled_measured_case(source, args, args.retention)?;
    let elapsed_ns = start.elapsed().as_nanos();
    black_box(result.output.len());
    Ok(TimedResult { result, elapsed_ns })
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
    result: CaseResult,
    elapsed_ns: Option<u128>,
    allocation: Option<AllocationReport>,
}

fn measure_one(source: &[u8], args: &Args) -> Result<MeasuredResult, BoxError> {
    if alloc_metrics::instrumented() {
        let mut result = None;
        let region = alloc_metrics::region(|| {
            result = Some(profiled_measured_case(source, args, args.retention)?);
            Ok::<(), BoxError>(())
        })?;
        let result = result
            .ok_or_else(|| Box::new(ProbeError("missing measured result".into())) as BoxError)?;
        Ok(MeasuredResult {
            result,
            elapsed_ns: None,
            allocation: Some(region.into()),
        })
    } else {
        let timed = timed_case(source, args)?;
        Ok(MeasuredResult {
            result: timed.result,
            elapsed_ns: Some(timed.elapsed_ns),
            allocation: None,
        })
    }
}

#[derive(Clone, Debug, Serialize)]
struct RouteReport {
    retention: &'static str,
    retention_ceiling_bytes: usize,
    output_sha256: String,
    output_bytes: usize,
    output_capacity: usize,
    semantic_change_count: usize,
    final_retained_before_commit: Option<usize>,
    stages: [StageObservation; MAX_EDIT_STAGES],
    oracle: oracle_0730::Check,
}

#[derive(Debug)]
struct RouteResult {
    retention: Retention,
    result: CaseResult,
    oracle: oracle_0730::Check,
}

fn route_result(
    source: &[u8],
    args: &Args,
    retention: Retention,
    expected: Option<&[u8]>,
) -> Result<RouteResult, BoxError> {
    let result = run_case(source, args, retention)?;
    let oracle_expected = expected.unwrap_or(&result.output);
    let oracle = oracle_0730::check_doc_output(
        source,
        oracle_expected,
        &result.output,
        final_text(args.edits),
    )?;
    Ok(RouteResult {
        retention,
        result,
        oracle,
    })
}

#[derive(Clone, Debug, Serialize)]
struct CorrectnessGates {
    all_stages_change_text: bool,
    exact_cross_route_outputs: bool,
    final_paragraph_semantic_readback: bool,
    final_public_reader_reopen: bool,
    exact_noop_source: bool,
    noop_shares_source_allocation: bool,
    forward_patch_exact_output: bool,
    inverse_patch_exact_source: bool,
    foreign_patch_refused: bool,
    stale_patch_refused: bool,
    same_value_preserves_retention: bool,
    operation_limit_refusal_atomic: bool,
    retained_capacities_within_ceiling: bool,
    zero_route_clears_retention: bool,
    release_route_clears_every_stage: bool,
    untouched_streams_preserved: bool,
    directory_metadata_preserved: bool,
    semantic_oracles_passed: bool,
    semantic_change_count_matches_edit_count: bool,
    allocator_retained_matches_output_capacity: bool,
}

struct Prepared {
    routes: [RouteResult; 3],
    gates: CorrectnessGates,
}

fn paragraph_zero_text(snapshot: &Snapshot) -> Result<String, BoxError> {
    let paragraphs = snapshot.paragraphs(Projection::All)?;
    paragraphs
        .into_iter()
        .find(|paragraph| paragraph.position() == Position::new(0))
        .map(|paragraph| paragraph.text().to_owned())
        .ok_or_else(|| Box::new(ProbeError("DOC paragraph zero is unavailable".into())) as BoxError)
}

fn all_stages_change_text(original: &str, edits: EditCount) -> bool {
    let mut previous = original;
    for stage in 0..edits.count() {
        let next = stage_text(stage);
        if next == previous {
            return false;
        }
        previous = next;
    }
    true
}

fn route_capacities_within_ceiling(route: &RouteResult, edits: EditCount) -> bool {
    let ceiling = route.retention.ceiling();
    route.result.stages.iter().take(edits.count()).all(|stage| {
        stage.after.is_none_or(|capacity| capacity <= ceiling)
            && stage
                .after_release
                .is_none_or(|capacity| capacity <= ceiling)
    }) && route
        .result
        .final_retained_before_commit
        .is_none_or(|capacity| capacity <= ceiling)
}

fn semantic_change_count_matches(route: &RouteResult, edits: EditCount) -> bool {
    route.result.semantic_change_count == edits.count()
}

fn zero_route_clears(route: &RouteResult, edits: EditCount) -> bool {
    route.retention == Retention::Zero
        && route.result.final_retained_before_commit.is_none()
        && route
            .result
            .stages
            .iter()
            .take(edits.count())
            .all(|stage| stage.before.is_none() && stage.after.is_none())
}

fn release_route_clears(route: &RouteResult, edits: EditCount) -> bool {
    route.retention == Retention::Release
        && route.result.final_retained_before_commit.is_none()
        && route
            .result
            .stages
            .iter()
            .take(edits.count())
            .all(|stage| stage.after_release.is_none())
}

fn final_paragraph_matches(result: &CaseResult, edits: EditCount) -> Result<bool, BoxError> {
    let snapshot = Snapshot::open(result.output.clone(), Limits::default())?;
    Ok(paragraph_zero_text(&snapshot)? == final_text(edits))
}

fn public_reader_reopens(result: &CaseResult) -> Result<bool, BoxError> {
    let mut package = Reader::from_reader(Cursor::new(result.output.clone()))?;
    package.document()?;
    Ok(true)
}

fn source_for_other_case(case: Case, input: &std::path::Path) -> Result<Vec<u8>, BoxError> {
    let other_name = match case {
        Case::DocNoHf => "FloatingPictures.doc",
        Case::DocFloat => "NoHeadFoot.doc",
    };
    let other = input
        .parent()
        .unwrap_or_else(|| std::path::Path::new("."))
        .join(other_name);
    Ok(std::fs::read(other)?)
}

fn patch_and_refusal_gates(
    source: &[u8],
    args: &Args,
) -> Result<(bool, bool, bool, bool), BoxError> {
    let source_snapshot = Snapshot::open_bounded(
        source.to_vec(),
        Limits::default(),
        transaction_limits(Retention::Zero),
    )?;
    let mut edit = source_snapshot.edit()?;
    edit.replace_paragraph(Position::new(0), FIRST_TEXT)?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    let forward = source_snapshot.package().apply(&patch)?.into_package();
    let forward_exact = package_bytes(&forward) == package_bytes(commit.package());
    let inverse = commit.package().apply(&patch.inverse())?.into_package();
    let inverse_exact = package_bytes(&inverse) == source_snapshot.bytes();

    let foreign_bytes = source_for_other_case(args.case, &args.input)?;
    let foreign = Snapshot::open_bounded(
        foreign_bytes,
        Limits::default(),
        transaction_limits(Retention::Zero),
    )?;
    let foreign_refused = foreign.package().apply(&patch).is_err();
    let stale_refused = commit.package().apply(&patch).is_err();
    Ok((forward_exact, inverse_exact, foreign_refused, stale_refused))
}

fn exact_noop_gate(
    source: &[u8],
    original: &str,
    retention: Retention,
) -> Result<(bool, bool), BoxError> {
    let source_snapshot = Snapshot::open_bounded(
        source.to_vec(),
        Limits::default(),
        transaction_limits(retention),
    )?;
    let mut edit = source_snapshot.edit()?;
    edit.replace_paragraph(Position::new(0), original)?;
    let no_retained_token = edit.retained_render_bytes().is_none();
    let commit = edit.commit()?;
    let exact_source = !commit.diagnostics().changed()
        && commit.patch().is_noop()
        && commit.patch().changes().len() == 0
        && package_bytes(commit.package()) == source
        && no_retained_token;
    let shares_source_allocation = Arc::ptr_eq(
        &source_snapshot.package().shared_source(),
        &commit.package().shared_source(),
    );
    Ok((exact_source, shares_source_allocation))
}

fn atomic_edit_controls(source: &[u8]) -> Result<(bool, bool), BoxError> {
    let source_snapshot = Snapshot::open_bounded(
        source.to_vec(),
        Limits::default(),
        transaction_limits(Retention::Default),
    )?;
    let mut one_operation = source_snapshot.edit()?;
    one_operation.replace_paragraph(Position::new(0), FIRST_TEXT)?;
    let one_operation_commit = one_operation.commit()?;
    let one_operation_bytes = package_bytes(one_operation_commit.package());
    let one_operation_patch_count = one_operation_commit.patch().changes().len();

    let mut same_value = source_snapshot.edit()?;
    same_value.replace_paragraph(Position::new(0), FIRST_TEXT)?;
    let held_before_same_value = same_value.retained_render_bytes();
    same_value.replace_paragraph(Position::new(0), FIRST_TEXT)?;
    let same_value_retention_unchanged =
        same_value.retained_render_bytes() == held_before_same_value;
    let same_value_commit = same_value.commit()?;
    let same_value_snapshot = Snapshot::open(
        package_bytes(same_value_commit.package()),
        Limits::default(),
    )?;
    let same_value_matches_one_operation = same_value_commit.diagnostics().changed()
        && !same_value_commit.patch().is_noop()
        && same_value_commit.patch().changes().len() == one_operation_patch_count
        && package_bytes(same_value_commit.package()) == one_operation_bytes
        && paragraph_zero_text(&same_value_snapshot)? == FIRST_TEXT;

    let limited = TransactionLimits::new(1, 4096, 4096)
        .with_max_retained_render_bytes(DEFAULT_RETAINED_RENDER_BYTES);
    let limited_snapshot = Snapshot::open_bounded(source.to_vec(), Limits::default(), limited)?;
    let mut limited_control = limited_snapshot.edit()?;
    limited_control.replace_paragraph(Position::new(0), FIRST_TEXT)?;
    let limited_control_commit = limited_control.commit()?;
    let limited_control_bytes = package_bytes(limited_control_commit.package());
    let limited_control_patch_count = limited_control_commit.patch().changes().len();

    let mut refused = limited_snapshot.edit()?;
    refused.replace_paragraph(Position::new(0), FIRST_TEXT)?;
    let held_before_refusal = refused.retained_render_bytes();
    let error = refused.replace_paragraph(Position::new(0), SECOND_TEXT);
    let typed_operation_refusal = matches!(
        error,
        Err(litchi_doc::package::edit::Error::Body(BodyError::Refused(
            Refusal::OperationLimit { .. }
        )))
    );
    let refusal_retention_unchanged = refused.retained_render_bytes() == held_before_refusal;
    let refusal_commit = refused.commit()?;
    let refusal_snapshot =
        Snapshot::open(package_bytes(refusal_commit.package()), Limits::default())?;
    let refusal_output_matches_one_operation = refusal_commit.diagnostics().changed()
        && !refusal_commit.patch().is_noop()
        && refusal_commit.patch().changes().len() == limited_control_patch_count
        && package_bytes(refusal_commit.package()) == limited_control_bytes
        && refusal_snapshot.bytes() == limited_control_bytes
        && paragraph_zero_text(&refusal_snapshot)? == FIRST_TEXT;
    Ok((
        same_value_retention_unchanged && same_value_matches_one_operation,
        typed_operation_refusal
            && refusal_retention_unchanged
            && refusal_output_matches_one_operation,
    ))
}

fn prepare(source: &[u8], args: &Args) -> Result<Prepared, BoxError> {
    let source_snapshot = Snapshot::open_bounded(
        source.to_vec(),
        Limits::default(),
        transaction_limits(Retention::Zero),
    )?;
    let original = paragraph_zero_text(&source_snapshot)?;
    let all_stages_change = all_stages_change_text(&original, args.edits);

    let zero = route_result(source, args, Retention::Zero, None)?;
    let default = route_result(source, args, Retention::Default, Some(&zero.result.output))?;
    let release = route_result(source, args, Retention::Release, Some(&zero.result.output))?;
    let routes = [zero, default, release];

    let exact_cross_route_outputs = routes
        .iter()
        .skip(1)
        .all(|route| route.result.output == routes[0].result.output);
    let final_paragraph_semantic_readback = routes
        .iter()
        .map(|route| final_paragraph_matches(&route.result, args.edits))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .all(|match_| match_);
    let final_public_reader_reopen = routes
        .iter()
        .map(|route| public_reader_reopens(&route.result))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .all(|reopened| reopened);
    let mut exact_noop_source = true;
    let mut noop_shares_source_allocation = true;
    for retention in [Retention::Zero, Retention::Default, Retention::Release] {
        let (exact, shares) = exact_noop_gate(source, &original, retention)?;
        exact_noop_source &= exact;
        noop_shares_source_allocation &= shares;
    }
    let (
        forward_patch_exact_output,
        inverse_patch_exact_source,
        foreign_patch_refused,
        stale_patch_refused,
    ) = patch_and_refusal_gates(source, args)?;
    let (same_value_preserves_retention, operation_limit_refusal_atomic) =
        atomic_edit_controls(source)?;
    let retained_capacities_within_ceiling = routes
        .iter()
        .all(|route| route_capacities_within_ceiling(route, args.edits));
    let zero_route_clears_retention = zero_route_clears(&routes[0], args.edits);
    let release_route_clears_every_stage = release_route_clears(&routes[2], args.edits);
    let untouched_streams_preserved = routes.iter().all(|route| {
        route.oracle.unchanged_stream_bytes_match_expected
            && route.oracle.unchanged_source_output_stream_bytes_match
            && route.oracle.output_stream_bytes_match_expected
    });
    let directory_metadata_preserved = routes.iter().all(|route| {
        route.oracle.semantic_directory_metadata_matches_expected
            && route.oracle.root_clsid_preserved
            && route.oracle.storage_clsids_preserved
    });
    let semantic_oracles_passed = routes
        .iter()
        .all(|route| route.oracle.oracle_ok && route.oracle.semantic_reopen_ok);
    let semantic_change_count_matches_edit_count = routes
        .iter()
        .all(|route| semantic_change_count_matches(route, args.edits));
    let gates = CorrectnessGates {
        all_stages_change_text: all_stages_change,
        exact_cross_route_outputs,
        final_paragraph_semantic_readback,
        final_public_reader_reopen,
        exact_noop_source,
        noop_shares_source_allocation,
        forward_patch_exact_output,
        inverse_patch_exact_source,
        foreign_patch_refused,
        stale_patch_refused,
        same_value_preserves_retention,
        operation_limit_refusal_atomic,
        retained_capacities_within_ceiling,
        zero_route_clears_retention,
        release_route_clears_every_stage,
        untouched_streams_preserved,
        directory_metadata_preserved,
        semantic_oracles_passed,
        semantic_change_count_matches_edit_count,
        allocator_retained_matches_output_capacity: true,
    };
    if !gates.all_stages_change_text
        || !gates.exact_cross_route_outputs
        || !gates.final_paragraph_semantic_readback
        || !gates.final_public_reader_reopen
        || !gates.exact_noop_source
        || !gates.noop_shares_source_allocation
        || !gates.forward_patch_exact_output
        || !gates.inverse_patch_exact_source
        || !gates.foreign_patch_refused
        || !gates.stale_patch_refused
        || !gates.same_value_preserves_retention
        || !gates.operation_limit_refusal_atomic
        || !gates.retained_capacities_within_ceiling
        || !gates.zero_route_clears_retention
        || !gates.release_route_clears_every_stage
        || !gates.untouched_streams_preserved
        || !gates.directory_metadata_preserved
        || !gates.semantic_oracles_passed
        || !gates.semantic_change_count_matches_edit_count
    {
        return failure(format!("functional correctness gates failed: {gates:?}"));
    }
    Ok(Prepared { routes, gates })
}

fn route_report(route: &RouteResult) -> RouteReport {
    RouteReport {
        retention: route.retention.name(),
        retention_ceiling_bytes: route.retention.ceiling(),
        output_sha256: sha256_hex(&route.result.output),
        output_bytes: route.result.output.len(),
        output_capacity: route.result.output_capacity,
        semantic_change_count: route.result.semantic_change_count,
        final_retained_before_commit: route.result.final_retained_before_commit,
        stages: route.result.stages,
        oracle: route.oracle.clone(),
    }
}

#[derive(Clone, Debug, Serialize)]
struct SampleReport {
    index: usize,
    elapsed_ns: Option<u128>,
    output_sha256: String,
    output_bytes: usize,
    output_capacity: usize,
    semantic_change_count: usize,
    exact_output_equal_reference: bool,
    final_paragraph_semantic_readback: bool,
    final_public_reader_reopen: bool,
    retained_capacities_within_ceiling: bool,
    allocator_retained_matches_output_capacity: bool,
    oracle: oracle_0730::Check,
    stages: [StageObservation; MAX_EDIT_STAGES],
    allocation: Option<AllocationReport>,
}

fn sample_report_with_source(
    source: &[u8],
    index: usize,
    measured: MeasuredResult,
    args: &Args,
    prepared: &Prepared,
) -> Result<SampleReport, BoxError> {
    let expected = &prepared.routes[0].result.output;
    let exact_output_equal_reference = measured.result.output == *expected;
    let final_paragraph_semantic_readback = final_paragraph_matches(&measured.result, args.edits)?;
    let final_public_reader_reopen = public_reader_reopens(&measured.result)?;
    let oracle = oracle_0730::check_doc_output(
        source,
        expected,
        &measured.result.output,
        final_text(args.edits),
    )?;
    let route = RouteResult {
        retention: args.retention,
        result: measured.result,
        oracle: oracle.clone(),
    };
    let retained_capacities_within_ceiling = route_capacities_within_ceiling(&route, args.edits);
    let semantic_change_count_matches_edit_count =
        semantic_change_count_matches(&route, args.edits);
    let allocator_retained_matches_output_capacity = measured
        .allocation
        .is_none_or(|allocation| allocation.retained_bytes == route.result.output_capacity as i64);
    if !exact_output_equal_reference
        || !final_paragraph_semantic_readback
        || !final_public_reader_reopen
        || !retained_capacities_within_ceiling
        || !allocator_retained_matches_output_capacity
        || !semantic_change_count_matches_edit_count
        || !oracle.oracle_ok
    {
        return failure(format!(
            "measured sample {index} failed output, semantic, retention, or oracle gate"
        ));
    }
    Ok(SampleReport {
        index,
        elapsed_ns: measured.elapsed_ns,
        output_sha256: sha256_hex(&route.result.output),
        output_bytes: route.result.output.len(),
        output_capacity: route.result.output_capacity,
        semantic_change_count: route.result.semantic_change_count,
        exact_output_equal_reference,
        final_paragraph_semantic_readback,
        final_public_reader_reopen,
        retained_capacities_within_ceiling,
        allocator_retained_matches_output_capacity,
        oracle,
        stages: route.result.stages,
        allocation: measured.allocation,
    })
}

#[derive(Debug, Serialize)]
struct ProbeOutput {
    schema_version: u32,
    schema: &'static str,
    change: &'static str,
    case: &'static str,
    input: String,
    source_sha256: String,
    source_bytes: usize,
    selected_paragraph: usize,
    edits: usize,
    alternating_texts: [&'static str; 2],
    retention: &'static str,
    retention_ceiling_bytes: usize,
    timing_claim: bool,
    allocator_instrumented: bool,
    allocator_timing_claim: bool,
    region_definition: &'static str,
    warmups: usize,
    samples_requested: usize,
    output_sha256: String,
    output_bytes: usize,
    output_capacity: usize,
    reference_output_sha256: String,
    reference_output_bytes: usize,
    exact_output_equal_reference: bool,
    route_outputs: [RouteReport; 3],
    correctness: CorrectnessGates,
    elapsed_ns: Vec<u128>,
    samples: Vec<SampleReport>,
    allocation: Option<AllocationReport>,
}

fn verify_fixture(args: &Args, source: &[u8]) -> Result<(), BoxError> {
    let actual = sha256_hex(source);
    if source.len() != args.case.expected_bytes() || actual != args.case.expected_sha256() {
        return failure(format!(
            "fixture identity mismatch for {}: got {} bytes / {}, expected {} bytes / {}",
            args.case.name(),
            source.len(),
            actual,
            args.case.expected_bytes(),
            args.case.expected_sha256()
        ));
    }
    Ok(())
}

pub fn run(_allocator_lane: bool) -> Result<(), BoxError> {
    let args = parse_args()?;
    if alloc_metrics::instrumented() && (args.warmup != 0 || args.samples != 1) {
        return failure("allocator lane requires --warmup 0 --samples 1");
    }
    let source = std::fs::read(&args.input)?;
    verify_fixture(&args, &source)?;
    let source_sha256 = sha256_hex(&source);
    let prepared = prepare(&source, &args)?;

    // Expected outputs and all correctness controls are complete before any
    // warmup or measured native sample. Their allocations never enter the
    // timed or allocator regions.
    let mut samples = Vec::with_capacity(args.samples);
    for _ in 0..args.warmup {
        let warmup = measure_one(&source, &args)?;
        drop(warmup);
    }
    for index in 0..args.samples {
        let measured = measure_one(&source, &args)?;
        samples.push(sample_report_with_source(
            &source, index, measured, &args, &prepared,
        )?);
    }
    let selected_route = &prepared.routes[match args.retention {
        Retention::Zero => 0,
        Retention::Default => 1,
        Retention::Release => 2,
    }];
    let measured_sample = samples.last();
    let output_sha256 = measured_sample
        .map(|sample| sample.output_sha256.clone())
        .unwrap_or_else(|| sha256_hex(&selected_route.result.output));
    let output_bytes = measured_sample.map_or(selected_route.result.output.len(), |sample| {
        sample.output_bytes
    });
    let output_capacity = measured_sample.map_or(selected_route.result.output_capacity, |sample| {
        sample.output_capacity
    });
    let elapsed_ns = samples
        .iter()
        .filter_map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    let allocation = samples.last().and_then(|sample| sample.allocation);
    let route_outputs = [
        route_report(&prepared.routes[0]),
        route_report(&prepared.routes[1]),
        route_report(&prepared.routes[2]),
    ];
    let output = ProbeOutput {
        schema_version: 1,
        schema: "litchi.performance.0917.doc_retention.v1",
        change: "0917",
        case: args.case.name(),
        input: args.input.display().to_string(),
        source_sha256,
        source_bytes: source.len(),
        selected_paragraph: 0,
        edits: args.edits.count(),
        alternating_texts: [FIRST_TEXT, SECOND_TEXT],
        retention: args.retention.name(),
        retention_ceiling_bytes: args.retention.ceiling(),
        timing_claim: !alloc_metrics::instrumented(),
        allocator_instrumented: alloc_metrics::instrumented(),
        allocator_timing_claim: false,
        region_definition: "public Package::from_bytes + edit_with_options source (record 0359a; was Snapshot::open_bounded), one Edit with all paragraph-0 replacements and retention releases, commit, output Vec copy, and local snapshot/edit/commit destruction; source file read, references, correctness checks, hashing and JSON excluded",
        warmups: args.warmup,
        samples_requested: args.samples,
        output_sha256,
        output_bytes,
        output_capacity,
        reference_output_sha256: sha256_hex(&prepared.routes[0].result.output),
        reference_output_bytes: prepared.routes[0].result.output.len(),
        exact_output_equal_reference: measured_sample
            .is_none_or(|sample| sample.exact_output_equal_reference),
        route_outputs,
        correctness: prepared.gates,
        elapsed_ns,
        samples,
        allocation,
    };
    println!("{}", serde_json::to_string(&output)?);
    Ok(())
}
