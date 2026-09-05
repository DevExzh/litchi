//! Bounded, allocator-only PPTX cross-copy retention diagnostics.
//!
//! This target deliberately reports absolute process allocator counters at a
//! small ownership journal. It has no timers and makes no claim about RSS,
//! object ownership, allocator-internal overlap, or latency. All corpus
//! construction and correctness gates run before the first observed
//! checkpoint; the lifecycle itself only compares against those prevalidated
//! bytes and metadata.

use std::error::Error;
use std::ffi::OsString;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::PathBuf;
use std::sync::Arc;

use litchi_core::ReadAt;
use serde::Serialize;

use crate::allocation_metrics::{self, Sample as AllocationSample, Snapshot as AllocationSnapshot};

const SCHEMA: &str = "pptx_retention_v1";
const COUNTER_REVISION: &str = "serialized_region_peak_v3";
const MAX_SAMPLES: usize = 1_000;
const MAX_WARMUP: usize = 1_000;
const MAX_WRITE: u64 = 64 * 1024;

const CALLBACK_SCOPE: &str = "global allocator callbacks after System returns; process-wide across threads; excludes RSS, object-owned memory, allocator-internal realloc overlap, and latency";
const OWNERSHIP_SCOPE: &str = "phase labels describe retained handles at callback-order boundaries; live bytes are process-global allocator counters and are not object-retention measurements";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Api {
    Owned,
    SourceBacked,
}

impl Api {
    const fn name(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::SourceBacked => "source-backed",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CorpusKind {
    Plain,
    MediaRich,
}

impl CorpusKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Plain => "plain",
            Self::MediaRich => "media-rich",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    api: Api,
    corpus: CorpusKind,
    samples: usize,
    warmup: usize,
    source_revision: String,
    output: PathBuf,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct Checkpoint {
    allocation_calls: u64,
    deallocation_calls: u64,
    reallocation_calls: u64,
    failed_allocation_calls: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    overflowed: bool,
    observer_invalid: bool,
}

#[derive(Clone, Copy, Debug, Default, Serialize)]
struct ArcCounts {
    source: usize,
    destination: usize,
}

#[derive(Clone, Debug, Serialize)]
struct RetentionRow {
    sample_index: usize,
    baseline_before_inputs: Checkpoint,
    prepared_inputs_and_sink: Checkpoint,
    opened_documents: Checkpoint,
    planned: Checkpoint,
    published: Checkpoint,
    drop_result: Checkpoint,
    drop_plan: Checkpoint,
    drop_document_handles: Checkpoint,
    #[serde(skip_serializing_if = "Option::is_none")]
    drop_caller_source_arcs: Option<Checkpoint>,
    drop_sink: Checkpoint,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_arc_counts_after_document_drop: Option<ArcCounts>,
    retention_probe: RetentionProbe,
}

#[derive(Clone, Debug, Serialize)]
struct RetentionProbe {
    status: allocation_metrics::Status,
    scope: allocation_metrics::Scope,
    allocation_calls: Option<u64>,
    deallocation_calls: Option<u64>,
    reallocation_calls: Option<u64>,
    failed_allocation_calls: Option<u64>,
    allocated_bytes: Option<u64>,
    deallocated_bytes: Option<u64>,
    live_bytes_before: Option<u64>,
    live_bytes_after: Option<u64>,
    peak_live_bytes_before: Option<u64>,
    peak_live_bytes_after: Option<u64>,
    #[serde(rename = "retention_probe_region_peak_live_bytes")]
    region_peak_live_bytes: Option<u64>,
}

impl From<AllocationSample> for RetentionProbe {
    fn from(sample: AllocationSample) -> Self {
        Self {
            status: sample.status,
            scope: sample.scope,
            allocation_calls: sample.allocation_calls,
            deallocation_calls: sample.deallocation_calls,
            reallocation_calls: sample.reallocation_calls,
            failed_allocation_calls: sample.failed_allocation_calls,
            allocated_bytes: sample.allocated_bytes,
            deallocated_bytes: sample.deallocated_bytes,
            live_bytes_before: sample.live_bytes_before,
            live_bytes_after: sample.live_bytes_after,
            peak_live_bytes_before: sample.peak_live_bytes_before,
            peak_live_bytes_after: sample.peak_live_bytes_after,
            region_peak_live_bytes: sample.region_peak_live_bytes,
        }
    }
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDescription {
    label: &'static str,
    live_owners: &'static str,
}

const PHASES_OWNED: [PhaseDescription; 9] = [
    PhaseDescription {
        label: "baseline_before_inputs",
        live_owners: "region acquired; corpus bytes remain outside the lifecycle; no input clones or sink",
    },
    PhaseDescription {
        label: "prepared_inputs_and_sink",
        live_owners: "owned source and destination input Vec clones, bounded sink, and reserved sink capacity",
    },
    PhaseDescription {
        label: "opened_documents",
        live_owners: "owned source and destination Packages plus their opened presentation handles",
    },
    PhaseDescription {
        label: "planned",
        live_owners: "opened source/destination handles and the validated cross-copy plan",
    },
    PhaseDescription {
        label: "published",
        live_owners: "opened packages, plan, publication result, and sink containing the exact output",
    },
    PhaseDescription {
        label: "drop_result",
        live_owners: "opened packages, plan, and exact output sink after the publication result is released",
    },
    PhaseDescription {
        label: "drop_plan",
        live_owners: "opened packages and exact output sink after the plan is released",
    },
    PhaseDescription {
        label: "drop_document_handles",
        live_owners: "exact output sink remains after source and destination Packages and opened handles are released",
    },
    PhaseDescription {
        label: "drop_sink",
        live_owners: "no lifecycle-owned input, document, plan, publication result, or sink handles; process baseline may differ",
    },
];

const PHASES_SOURCE_BACKED: [PhaseDescription; 10] = [
    PhaseDescription {
        label: "baseline_before_inputs",
        live_owners: "region acquired; corpus bytes remain outside the lifecycle; no source Arcs or sink",
    },
    PhaseDescription {
        label: "prepared_inputs_and_sink",
        live_owners: "source-backed caller InstrumentedSource Arcs, temporary ReadAt Arcs, and reserved bounded sink",
    },
    PhaseDescription {
        label: "opened_documents",
        live_owners: "source-backed presentation view and editor, caller source Arcs, and sink",
    },
    PhaseDescription {
        label: "planned",
        live_owners: "source-backed view, editor, source-retaining plan, caller source Arcs, and sink",
    },
    PhaseDescription {
        label: "published",
        live_owners: "source-backed view, plan, publication result, caller source Arcs, and exact output sink; editor is consumed by publication",
    },
    PhaseDescription {
        label: "drop_result",
        live_owners: "source-backed view, plan, caller source Arcs, and exact output sink",
    },
    PhaseDescription {
        label: "drop_plan",
        live_owners: "source-backed view, caller source Arcs, and exact output sink",
    },
    PhaseDescription {
        label: "drop_document_handles",
        live_owners: "source InstrumentedSource Arcs and exact output sink; source-backed view is released",
    },
    PhaseDescription {
        label: "drop_caller_source_arcs",
        live_owners: "exact output sink only; caller InstrumentedSource Arcs are released",
    },
    PhaseDescription {
        label: "drop_sink",
        live_owners: "no lifecycle-owned source, document, plan, publication result, or sink handles; process baseline may differ",
    },
];

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    api: &'static str,
    corpus: &'static str,
    corpus_manifest: crate::CorpusManifest,
    samples: usize,
    warmup: usize,
    source_revision: String,
    source_archive_sha256: String,
    source_archive_bytes: usize,
    destination_archive_sha256: String,
    destination_archive_bytes: usize,
    expected_output_sha256: String,
    expected_output_bytes: usize,
    corpus_gates_verified: bool,
    all_iteration_output_bytes_verified: bool,
    checked_iteration_count: usize,
    binary_sha256: String,
    current_exe: String,
    binary_bytes: u64,
    allocator: &'static str,
    instrumentation: &'static str,
    allocator_counter_revision: &'static str,
    callback_scope: &'static str,
    ownership_scope: &'static str,
    phases: Vec<PhaseDescription>,
    samples_raw: Vec<RetentionRow>,
}

#[derive(Clone, Copy, Debug)]
struct RawCheckpoints {
    baseline: Checkpoint,
    inputs_and_sink_ready: Checkpoint,
    opened: Checkpoint,
    planned: Checkpoint,
    published: Checkpoint,
    publication_result_dropped: Checkpoint,
    plan_dropped: Checkpoint,
    document_handles_dropped: Checkpoint,
    caller_source_arcs_dropped: Option<Checkpoint>,
    sink_dropped: Checkpoint,
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut api = None;
    let mut corpus = None;
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("retention argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value_text = value
            .to_str()
            .ok_or("retention argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--api" => {
                if api.is_some() {
                    return Err("duplicate --api".into());
                }
                api = Some(match value_text {
                    "owned" => Api::Owned,
                    "source-backed" => Api::SourceBacked,
                    _ => return Err("--api must be owned or source-backed".into()),
                });
            },
            "--corpus" => {
                if corpus.is_some() {
                    return Err("duplicate --corpus".into());
                }
                corpus = Some(match value_text {
                    "plain" => CorpusKind::Plain,
                    "media-rich" => CorpusKind::MediaRich,
                    _ => return Err("--corpus must be plain or media-rich".into()),
                });
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("duplicate --samples".into());
                }
                samples = Some(parse_bounded(value_text, "samples", 1, MAX_SAMPLES)?);
            },
            "--warmup" => {
                if warmup.is_some() {
                    return Err("duplicate --warmup".into());
                }
                warmup = Some(parse_bounded(value_text, "warmup", 0, MAX_WARMUP)?);
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("duplicate --source-revision".into());
                }
                if value_text.len() != 40
                    || !value_text.bytes().all(|byte| byte.is_ascii_hexdigit())
                    || value_text.to_ascii_lowercase() != value_text
                {
                    return Err(
                        "--source-revision must contain exactly 40 hexadecimal characters".into(),
                    );
                }
                source_revision = Some(value_text.to_owned());
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
            _ => return Err(format!("unknown retention argument: {flag}").into()),
        }
    }
    Ok(Config {
        api: api.ok_or("missing --api")?,
        corpus: corpus.ok_or("missing --corpus")?,
        samples: samples.ok_or("missing --samples")?,
        warmup: warmup.ok_or("missing --warmup")?,
        source_revision: source_revision.ok_or("missing --source-revision")?,
        output: output.ok_or("missing --output")?,
    })
}

fn parse_bounded(
    value: &str,
    name: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("--{name} must be an unsigned decimal integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("--{name} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn lifecycle_case(api: Api, corpus: CorpusKind) -> crate::Case {
    match (api, corpus) {
        (Api::Owned, CorpusKind::Plain) => crate::Case::PptxCrossCopyPlainLifecycle,
        (Api::Owned, CorpusKind::MediaRich) => crate::Case::PptxCrossCopyMediaRichLifecycle,
        (Api::SourceBacked, CorpusKind::Plain) => {
            crate::Case::PptxSourceBackedCrossCopyPlainLifecycle
        },
        (Api::SourceBacked, CorpusKind::MediaRich) => {
            crate::Case::PptxSourceBackedCrossCopyMediaRichLifecycle
        },
    }
}

fn snapshot_checkpoint(snapshot: AllocationSnapshot) -> Result<Checkpoint, Box<dyn Error>> {
    if snapshot.overflowed || snapshot.observer_invalid {
        return Err("allocator snapshot is overflowed or observer-invalid".into());
    }
    let expected_live = snapshot
        .allocated_bytes
        .checked_sub(snapshot.deallocated_bytes)
        .ok_or("allocator snapshot live-byte balance underflowed")?;
    if snapshot.live_bytes != expected_live || snapshot.peak_live_bytes < snapshot.live_bytes {
        return Err("allocator snapshot live-byte balance or high-water is invalid".into());
    }
    Ok(Checkpoint {
        allocation_calls: snapshot.allocation_calls,
        deallocation_calls: snapshot.deallocation_calls,
        reallocation_calls: snapshot.reallocation_calls,
        failed_allocation_calls: snapshot.failed_allocation_calls,
        allocated_bytes: snapshot.allocated_bytes,
        deallocated_bytes: snapshot.deallocated_bytes,
        live_bytes: snapshot.live_bytes,
        peak_live_bytes: snapshot.peak_live_bytes,
        overflowed: snapshot.overflowed,
        observer_invalid: snapshot.observer_invalid,
    })
}

fn capture_checkpoint() -> Result<Checkpoint, Box<dyn Error>> {
    snapshot_checkpoint(allocation_metrics::snapshot())
}

fn validate_transition(previous: Checkpoint, current: Checkpoint) -> Result<(), Box<dyn Error>> {
    if current.allocation_calls < previous.allocation_calls
        || current.deallocation_calls < previous.deallocation_calls
        || current.reallocation_calls < previous.reallocation_calls
        || current.failed_allocation_calls < previous.failed_allocation_calls
        || current.allocated_bytes < previous.allocated_bytes
        || current.deallocated_bytes < previous.deallocated_bytes
        || current.peak_live_bytes < previous.peak_live_bytes
    {
        return Err("allocator absolute counters are not monotonic".into());
    }
    let expected_live = previous
        .live_bytes
        .checked_add(current.allocated_bytes - previous.allocated_bytes)
        .and_then(|live| live.checked_sub(current.deallocated_bytes - previous.deallocated_bytes))
        .ok_or("allocator live-byte balance overflowed")?;
    if expected_live != current.live_bytes {
        return Err("allocator live bytes do not balance successful callbacks".into());
    }
    Ok(())
}

fn validate_all_transitions(checkpoints: RawCheckpoints) -> Result<(), Box<dyn Error>> {
    let sequence = [
        checkpoints.baseline,
        checkpoints.inputs_and_sink_ready,
        checkpoints.opened,
        checkpoints.planned,
        checkpoints.published,
        checkpoints.publication_result_dropped,
        checkpoints.plan_dropped,
        checkpoints.document_handles_dropped,
    ];
    for pair in sequence.windows(2) {
        validate_transition(pair[0], pair[1])?;
    }
    let mut previous = checkpoints.document_handles_dropped;
    if let Some(caller) = checkpoints.caller_source_arcs_dropped {
        validate_transition(previous, caller)?;
        previous = caller;
    }
    validate_transition(previous, checkpoints.sink_dropped)
}

fn validate_region(
    sample: &AllocationSample,
    checkpoints: RawCheckpoints,
) -> Result<(), Box<dyn Error>> {
    if sample.status != allocation_metrics::Status::Measured {
        return Err("allocator lifecycle region is unavailable or overflowed".into());
    }
    let before = sample
        .live_bytes_before
        .ok_or("measured region omitted its starting live bytes")?;
    let after = sample
        .live_bytes_after
        .ok_or("measured region omitted its ending live bytes")?;
    let peak = sample
        .region_peak_live_bytes
        .ok_or("measured region omitted its observer-ordered peak")?;
    if before != checkpoints.baseline.live_bytes || after != checkpoints.sink_dropped.live_bytes {
        return Err("region endpoints differ from the lifecycle checkpoints".into());
    }
    let all = [
        checkpoints.baseline,
        checkpoints.inputs_and_sink_ready,
        checkpoints.opened,
        checkpoints.planned,
        checkpoints.published,
        checkpoints.publication_result_dropped,
        checkpoints.plan_dropped,
        checkpoints.document_handles_dropped,
        checkpoints
            .caller_source_arcs_dropped
            .unwrap_or(checkpoints.document_handles_dropped),
        checkpoints.sink_dropped,
    ];
    if all.iter().any(|checkpoint| checkpoint.live_bytes > peak) {
        return Err("a lifecycle live-byte point exceeds the region peak".into());
    }
    if sample.peak_live_bytes_before != Some(checkpoints.baseline.peak_live_bytes)
        || sample.peak_live_bytes_after != Some(checkpoints.sink_dropped.peak_live_bytes)
        || peak < before
        || peak < after
        || peak > checkpoints.sink_dropped.peak_live_bytes
    {
        return Err("region high-water endpoints are inconsistent".into());
    }
    let fields = [
        sample.allocation_calls,
        sample.deallocation_calls,
        sample.reallocation_calls,
        sample.failed_allocation_calls,
        sample.allocated_bytes,
        sample.deallocated_bytes,
    ];
    if fields.iter().any(Option::is_none) {
        return Err("measured region omitted counter deltas".into());
    }
    let expected = [
        checkpoints
            .sink_dropped
            .allocation_calls
            .checked_sub(checkpoints.baseline.allocation_calls),
        checkpoints
            .sink_dropped
            .deallocation_calls
            .checked_sub(checkpoints.baseline.deallocation_calls),
        checkpoints
            .sink_dropped
            .reallocation_calls
            .checked_sub(checkpoints.baseline.reallocation_calls),
        checkpoints
            .sink_dropped
            .failed_allocation_calls
            .checked_sub(checkpoints.baseline.failed_allocation_calls),
        checkpoints
            .sink_dropped
            .allocated_bytes
            .checked_sub(checkpoints.baseline.allocated_bytes),
        checkpoints
            .sink_dropped
            .deallocated_bytes
            .checked_sub(checkpoints.baseline.deallocated_bytes),
    ];
    if fields
        .iter()
        .zip(expected)
        .any(|(actual, expected)| *actual != expected)
    {
        return Err("region counter deltas differ from lifecycle endpoints".into());
    }
    Ok(())
}

fn make_sink(maximum: usize) -> Result<crate::CountingSink, Box<dyn Error>> {
    let maximum = sink_ceiling(maximum)?;
    let mut sink = crate::CountingSink::bounded(maximum, MAX_WRITE);
    sink.reserve_budget()?;
    Ok(sink)
}

fn sink_ceiling(maximum: usize) -> Result<u64, Box<dyn Error>> {
    u64::try_from(maximum)?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or_else(|| "sink ceiling overflows u64".into())
}

fn validate_owned_plan(
    plan: &litchi_pptx::opened::CrossSlideCopyPlan,
    corpus: &crate::PptxCrossCopyCorpus,
) -> Result<(), Box<dyn Error>> {
    if plan.source().name() != corpus.source_slide_name
        || plan.destination().name() != corpus.destination_slide_name
        || plan.position() != corpus.insertion_position
        || plan.parts().len() != corpus.plan_parts
        || plan.planned_bytes() != corpus.planned_bytes
        || plan.external_relationship_count() != corpus.external_relationships
        || plan.source_layout() != plan.destination_layout()
        || plan
            .parts()
            .iter()
            .filter(|part| part.source() != part.target())
            .count()
            != corpus.collision_remapped_parts
    {
        return Err("owned cross-copy plan metadata differs from the prevalidated corpus".into());
    }
    Ok(())
}

fn validate_source_plan(
    plan: &litchi_pptx::SourceBackedCrossSlideCopyPlan,
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
) -> Result<(), Box<dyn Error>> {
    if plan.source_position() != corpus.source_slide
        || plan.destination_slide_position() != corpus.destination_slide
        || plan.insertion_position() != corpus.insertion_position
        || plan.destination_slide_count() != corpus.destination_slide_count + 1
    {
        return Err(
            "source-backed cross-copy plan metadata differs from the prevalidated corpus".into(),
        );
    }
    Ok(())
}

fn validate_sink(sink: &crate::CountingSink, expected: &[u8]) -> Result<(), Box<dyn Error>> {
    if sink.bytes != expected
        || sink.summary().accepted_bytes != u64::try_from(expected.len())?
        || sink.summary().largest_write > MAX_WRITE
    {
        return Err("sequential sink output differs from the prevalidated exact output".into());
    }
    Ok(())
}

fn run_owned_iteration(
    corpus: &crate::PptxCrossCopyCorpus,
) -> Result<(RawCheckpoints, AllocationSample, Option<ArcCounts>), Box<dyn Error>> {
    let region = allocation_metrics::begin();
    let baseline = capture_checkpoint()?;
    let source_input = corpus.source_archive.clone();
    let destination_input = corpus.destination_archive.clone();
    let mut sink = make_sink(corpus.expected_output.len())?;
    let inputs_and_sink_ready = capture_checkpoint()?;

    let source = litchi_pptx::Package::from_vec(source_input)?;
    let mut destination = litchi_pptx::Package::from_vec(destination_input)?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination.opened_presentation()?;
    let opened = capture_checkpoint()?;

    let plan = destination_snapshot.plan_cross_slide_copy(
        &source_snapshot,
        corpus.source_slide,
        corpus.destination_slide,
        corpus.insertion_position,
    )?;
    validate_owned_plan(&plan, corpus)?;
    let planned = capture_checkpoint()?;

    let published = destination.apply_cross_slide_copy_plan(&source, &plan)?;
    if published.slides().len() != corpus.destination_slide_count + 1 {
        return Err("owned publication changed the wrong slide count".into());
    }
    {
        let destination_opc = destination.opc()?;
        destination_opc.to_stream(&mut sink)?;
    }
    validate_sink(&sink, &corpus.expected_output)?;
    let published_checkpoint = capture_checkpoint()?;

    drop(published);
    let publication_result_dropped = capture_checkpoint()?;
    drop(plan);
    let plan_dropped = capture_checkpoint()?;
    drop(source_snapshot);
    drop(destination_snapshot);
    drop(source);
    drop(destination);
    let document_handles_dropped = capture_checkpoint()?;
    drop(sink);
    let sink_dropped = capture_checkpoint()?;

    let checkpoints = RawCheckpoints {
        baseline,
        inputs_and_sink_ready,
        opened,
        planned,
        published: published_checkpoint,
        publication_result_dropped,
        plan_dropped,
        document_handles_dropped,
        caller_source_arcs_dropped: None,
        sink_dropped,
    };
    validate_all_transitions(checkpoints)?;
    let region_sample = region.finish().ok_or("allocator region did not publish")?;
    validate_region(&region_sample, checkpoints)?;
    Ok((checkpoints, region_sample, None))
}

fn run_source_backed_iteration(
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
) -> Result<(RawCheckpoints, AllocationSample, Option<ArcCounts>), Box<dyn Error>> {
    let region = allocation_metrics::begin();
    let baseline = capture_checkpoint()?;
    let source = Arc::new(crate::InstrumentedSource::new(
        corpus.source_archive.clone(),
        Vec::new(),
    ));
    let destination = Arc::new(crate::InstrumentedSource::new(
        corpus.destination_archive.clone(),
        Vec::new(),
    ));
    let source_read: Arc<dyn ReadAt> = source.clone();
    let destination_read: Arc<dyn ReadAt> = destination.clone();
    let ceiling = sink_ceiling(corpus.matched_owned_output_bytes)?;
    if corpus.source_backed_expected_output.len() > usize::try_from(ceiling)? {
        return Err("source-backed expected output exceeds the matched owned sink ceiling".into());
    }
    let mut sink = make_sink(corpus.matched_owned_output_bytes)?;
    let inputs_and_sink_ready = capture_checkpoint()?;

    let source_view = litchi_pptx::SourceBackedPresentation::from_read_at(source_read)?;
    let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at(destination_read)?;
    if source_view.slide_count() != crate::PPTX_CROSS_COPY_SOURCE_SLIDE_COUNT
        || editor.slide_count() != corpus.destination_slide_count
    {
        return Err("source-backed setup changed the prevalidated slide counts".into());
    }
    let opened = capture_checkpoint()?;

    let plan = editor.plan_cross_slide_copy(
        &source_view,
        corpus.source_slide,
        corpus.destination_slide,
        corpus.insertion_position,
    )?;
    validate_source_plan(&plan, corpus)?;
    let planned = capture_checkpoint()?;

    let published = editor.publish_cross_slide_copy_to_stream(&mut sink, &plan)?;
    if published.destination_slide_count() != corpus.destination_slide_count + 1
        || published.insertion_position() != corpus.insertion_position
        || published.name() != corpus.source_slide_name
    {
        return Err(
            "source-backed publication metadata differs from the prevalidated corpus".into(),
        );
    }
    validate_sink(&sink, &corpus.source_backed_expected_output)?;
    let published_checkpoint = capture_checkpoint()?;

    drop(published);
    let publication_result_dropped = capture_checkpoint()?;
    drop(plan);
    let plan_dropped = capture_checkpoint()?;
    drop(source_view);
    let document_handles_dropped = capture_checkpoint()?;
    let arc_counts_after_document_drop = ArcCounts {
        source: Arc::strong_count(&source),
        destination: Arc::strong_count(&destination),
    };
    if arc_counts_after_document_drop.source != 1 || arc_counts_after_document_drop.destination != 1
    {
        return Err("source-backed document drop retained an unexpected caller Arc owner".into());
    }
    drop(source);
    drop(destination);
    let caller_source_arcs_dropped = capture_checkpoint()?;
    drop(sink);
    let sink_dropped = capture_checkpoint()?;

    let checkpoints = RawCheckpoints {
        baseline,
        inputs_and_sink_ready,
        opened,
        planned,
        published: published_checkpoint,
        publication_result_dropped,
        plan_dropped,
        document_handles_dropped,
        caller_source_arcs_dropped: Some(caller_source_arcs_dropped),
        sink_dropped,
    };
    validate_all_transitions(checkpoints)?;
    let region_sample = region.finish().ok_or("allocator region did not publish")?;
    validate_region(&region_sample, checkpoints)?;
    Ok((
        checkpoints,
        region_sample,
        Some(arc_counts_after_document_drop),
    ))
}

fn row_from_parts(
    index: usize,
    checkpoints: RawCheckpoints,
    region: AllocationSample,
    arc_counts: Option<ArcCounts>,
) -> RetentionRow {
    RetentionRow {
        sample_index: index,
        baseline_before_inputs: checkpoints.baseline,
        prepared_inputs_and_sink: checkpoints.inputs_and_sink_ready,
        opened_documents: checkpoints.opened,
        planned: checkpoints.planned,
        published: checkpoints.published,
        drop_result: checkpoints.publication_result_dropped,
        drop_plan: checkpoints.plan_dropped,
        drop_document_handles: checkpoints.document_handles_dropped,
        drop_caller_source_arcs: checkpoints.caller_source_arcs_dropped,
        drop_sink: checkpoints.sink_dropped,
        source_arc_counts_after_document_drop: arc_counts,
        retention_probe: region.into(),
    }
}

fn validate_owned_gates(corpus: &crate::PptxCrossCopyCorpus) -> Result<(), Box<dyn Error>> {
    let gates = &corpus.gates;
    if !gates.semantic_output_verified
        || !gates.package_topology_verified
        || !gates.dependency_closure_verified
        || !gates.source_immutability_verified
        || !gates.collision_remap_verified
        || !gates.durable_patch_round_trip_verified
        || !gates.borrowed_provenance_refusal_verified
        || !gates.stale_source_refusal_verified
        || !gates.stale_destination_refusal_verified
        || !gates.foreign_source_refusal_verified
    {
        return Err("owned PPTX lifecycle corpus has an incomplete gate set".into());
    }
    Ok(())
}

fn validate_source_backed_gates(
    corpus: &crate::PptxSourceBackedCrossCopyCorpus,
) -> Result<(), Box<dyn Error>> {
    let gates = corpus
        .lifecycle_gates
        .as_ref()
        .ok_or("source-backed PPTX lifecycle corpus has no lifecycle gate set")?;
    if !gates.matched_owned_corpus_verified
        || !gates.semantic_output_verified
        || !gates.package_topology_verified
        || !gates.dependency_boundary_verified
        || !gates.layout_reuse_verified
        || !gates.untouched_destination_members_verified
        || !gates.deterministic_output_verified
        || !gates.source_version_stability_verified
        || !gates.source_revision_refusal_verified
        || !gates.destination_revision_refusal_verified
        || !gates.foreign_destination_refusal_verified
        || !gates.added_opc_parts_verified
        || !gates.added_zip_members_verified
        || !gates.media_leaf_payloads_verified
        || !gates.media_leaf_content_types_verified
        || !gates.media_relationships_verified
    {
        return Err("source-backed PPTX lifecycle corpus has an incomplete gate set".into());
    }
    Ok(())
}

enum PreparedCorpus {
    Owned(crate::PptxCrossCopyCorpus),
    SourceBacked(crate::PptxSourceBackedCrossCopyCorpus),
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    let counter_revision = allocation_metrics::counter_revision()
        .ok_or("allocator instrumentation is disabled; use litchi-perf-baseline-alloc")?;
    if counter_revision != COUNTER_REVISION {
        return Err("allocator counter revision is unsupported by this diagnostic".into());
    }
    let case = lifecycle_case(config.api, config.corpus);
    let binary = crate::current_executable_identity()?;

    // Build and validate one complete fixed corpus before observations. Keep
    // this object alive for every iteration so no corpus construction or hash
    // enters the lifecycle region.
    let prepared = match config.api {
        Api::Owned => {
            let corpus = crate::build_pptx_cross_copy_corpus(case)?;
            validate_owned_gates(&corpus)?;
            PreparedCorpus::Owned(corpus)
        },
        Api::SourceBacked => {
            let corpus = crate::build_pptx_source_backed_cross_copy_corpus(case)?;
            validate_source_backed_gates(&corpus)?;
            PreparedCorpus::SourceBacked(corpus)
        },
    };
    let (
        source_digest,
        destination_digest,
        expected_digest,
        source_bytes,
        destination_bytes,
        expected_bytes,
        corpus_manifest,
    ) = match &prepared {
        PreparedCorpus::Owned(corpus) => (
            corpus.source_archive_sha256.clone(),
            corpus.manifest.archive_sha256.clone(),
            crate::sha256_hex(&corpus.expected_output),
            corpus.source_archive.len(),
            corpus.destination_archive.len(),
            corpus.expected_output.len(),
            corpus.manifest.clone(),
        ),
        PreparedCorpus::SourceBacked(corpus) => (
            crate::sha256_hex(&corpus.source_archive),
            crate::sha256_hex(&corpus.destination_archive),
            crate::sha256_hex(&corpus.source_backed_expected_output),
            corpus.source_archive.len(),
            corpus.destination_archive.len(),
            corpus.source_backed_expected_output.len(),
            corpus.manifest.clone(),
        ),
    };
    let total = config
        .warmup
        .checked_add(config.samples)
        .ok_or("warmup and samples overflow usize")?;
    let mut rows: Vec<Option<RetentionRow>> = (0..config.samples).map(|_| None).collect();
    let mut checked_iteration_count = 0;
    for iteration in 0..total {
        let (checkpoints, region, arc_counts) = match &prepared {
            PreparedCorpus::Owned(corpus) => run_owned_iteration(corpus)?,
            PreparedCorpus::SourceBacked(corpus) => run_source_backed_iteration(corpus)?,
        };
        checked_iteration_count += 1;
        if iteration >= config.warmup {
            let index = iteration - config.warmup;
            rows[index] = Some(row_from_parts(index, checkpoints, region, arc_counts));
        }
    }
    let samples_raw = rows
        .into_iter()
        .map(|row| row.ok_or("retention sample storage has a missing row"))
        .collect::<Result<Vec<_>, _>>()?;
    let phases = match config.api {
        Api::Owned => PHASES_OWNED.to_vec(),
        Api::SourceBacked => PHASES_SOURCE_BACKED.to_vec(),
    };
    Ok(Report {
        schema: SCHEMA,
        api: config.api.name(),
        corpus: config.corpus.name(),
        corpus_manifest,
        samples: config.samples,
        warmup: config.warmup,
        source_revision: config.source_revision.clone(),
        source_archive_sha256: source_digest,
        source_archive_bytes: source_bytes,
        destination_archive_sha256: destination_digest,
        destination_archive_bytes: destination_bytes,
        expected_output_sha256: expected_digest,
        expected_output_bytes: expected_bytes,
        corpus_gates_verified: true,
        all_iteration_output_bytes_verified: checked_iteration_count == total,
        checked_iteration_count,
        binary_sha256: binary.binary_sha256.clone(),
        current_exe: binary.path.clone(),
        binary_bytes: binary.binary_bytes,
        allocator: allocation_metrics::allocator_identity(),
        instrumentation: allocation_metrics::instrumentation_identity(),
        allocator_counter_revision: counter_revision,
        callback_scope: CALLBACK_SCOPE,
        ownership_scope: OWNERSHIP_SCOPE,
        phases,
        samples_raw,
    })
}

/// Run the bounded retention diagnostic from allocator-binary arguments after
/// the caller has installed and enabled the allocator wrapper.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let args = args.into_iter().collect::<Vec<_>>();
    let config = parse_config(&args)?;
    let report = run_capture(&config)?;
    let bytes = serde_json::to_vec_pretty(&report)?;
    let mut file = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&config.output)?;
    file.write_all(&bytes)?;
    file.write_all(b"\n")?;
    file.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn args(values: &[&str]) -> Vec<OsString> {
        values.iter().map(OsString::from).collect()
    }

    #[test]
    fn config_requires_each_flag_and_rejects_duplicates_and_unknowns() {
        let valid = args(&[
            "--api",
            "owned",
            "--corpus",
            "plain",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "retention.json",
        ]);
        assert!(parse_config(&valid).is_ok());
        let mut duplicate = valid.clone();
        duplicate.extend(args(&["--samples", "2"]));
        assert!(parse_config(&duplicate).is_err());
        let mut unknown = valid;
        unknown.extend(args(&["--unexpected", "value"]));
        assert!(parse_config(&unknown).is_err());
    }

    #[test]
    fn config_rejects_missing_required_flags_and_numeric_overflow() {
        let valid = args(&[
            "--api",
            "owned",
            "--corpus",
            "plain",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "retention.json",
        ]);
        for pair in [[0, 1], [2, 3], [4, 5], [6, 7], [8, 9], [10, 11]] {
            let missing = valid
                .iter()
                .enumerate()
                .filter(|(index, _)| *index != pair[0] && *index != pair[1])
                .map(|(_, value)| value.clone())
                .collect::<Vec<_>>();
            assert!(parse_config(&missing).is_err());
        }
        let mut samples_too_large = valid.clone();
        samples_too_large[5] = OsString::from("1001");
        assert!(parse_config(&samples_too_large).is_err());
        let mut warmup_too_large = valid.clone();
        warmup_too_large[7] = OsString::from("1001");
        assert!(parse_config(&warmup_too_large).is_err());
        let mut decimal_overflow = valid;
        decimal_overflow[5] = OsString::from("184467440737095516161");
        assert!(parse_config(&decimal_overflow).is_err());
    }

    #[test]
    fn config_bounds_source_revision_and_output_are_strict() {
        let mut values = args(&[
            "--api",
            "source-backed",
            "--corpus",
            "media-rich",
            "--samples",
            "0",
            "--warmup",
            "0",
            "--source-revision",
            "0123456789abcdef0123456789abcdef01234567",
            "--output",
            "retention.json",
        ]);
        assert!(parse_config(&values).is_err());
        values[5] = OsString::from("1");
        values[9] = OsString::from("not-40-hex");
        assert!(parse_config(&values).is_err());
        values[9] = OsString::from("0123456789abcdef0123456789abcdef01234567");
        values[11] = OsString::from("");
        assert!(parse_config(&values).is_err());
    }

    #[test]
    fn counter_validation_rejects_nonmonotonic_and_unbalanced_points() {
        let first = Checkpoint {
            allocated_bytes: 10,
            live_bytes: 10,
            ..Checkpoint::default()
        };
        let lower = Checkpoint {
            allocated_bytes: 9,
            live_bytes: 9,
            ..Checkpoint::default()
        };
        assert!(validate_transition(first, lower).is_err());
        let unbalanced = Checkpoint {
            allocated_bytes: 20,
            live_bytes: 11,
            ..Checkpoint::default()
        };
        assert!(validate_transition(first, unbalanced).is_err());
        let balanced = Checkpoint {
            allocated_bytes: 20,
            live_bytes: 20,
            ..Checkpoint::default()
        };
        assert!(validate_transition(first, balanced).is_ok());
    }

    #[test]
    fn snapshot_validation_rejects_flags_underflow_and_low_peak() {
        let valid = AllocationSnapshot {
            allocated_bytes: 10,
            live_bytes: 10,
            peak_live_bytes: 10,
            ..AllocationSnapshot::default()
        };
        assert!(snapshot_checkpoint(valid).is_ok());
        assert!(
            snapshot_checkpoint(AllocationSnapshot {
                overflowed: true,
                ..valid
            })
            .is_err()
        );
        assert!(
            snapshot_checkpoint(AllocationSnapshot {
                observer_invalid: true,
                ..valid
            })
            .is_err()
        );
        assert!(
            snapshot_checkpoint(AllocationSnapshot {
                allocated_bytes: 0,
                deallocated_bytes: 1,
                live_bytes: 0,
                ..valid
            })
            .is_err()
        );
        assert!(
            snapshot_checkpoint(AllocationSnapshot {
                peak_live_bytes: 9,
                ..valid
            })
            .is_err()
        );
    }

    fn measured_region_sample() -> AllocationSample {
        AllocationSample {
            status: allocation_metrics::Status::Measured,
            scope: allocation_metrics::Scope::OperationGlobalSystemAllocator,
            allocation_calls: Some(0),
            deallocation_calls: Some(0),
            reallocation_calls: Some(0),
            failed_allocation_calls: Some(0),
            allocated_bytes: Some(0),
            deallocated_bytes: Some(0),
            live_bytes_before: Some(10),
            live_bytes_after: Some(10),
            peak_live_bytes_before: Some(10),
            peak_live_bytes_after: Some(10),
            region_peak_live_bytes: Some(10),
        }
    }

    fn stable_checkpoints() -> RawCheckpoints {
        let point = Checkpoint {
            allocated_bytes: 10,
            live_bytes: 10,
            peak_live_bytes: 10,
            ..Checkpoint::default()
        };
        RawCheckpoints {
            baseline: point,
            inputs_and_sink_ready: point,
            opened: point,
            planned: point,
            published: point,
            publication_result_dropped: point,
            plan_dropped: point,
            document_handles_dropped: point,
            caller_source_arcs_dropped: None,
            sink_dropped: point,
        }
    }

    #[test]
    fn region_validation_rejects_missing_numeric_endpoint_and_intermediate_peak() {
        let checkpoints = stable_checkpoints();
        let mut missing = measured_region_sample();
        missing.allocated_bytes = None;
        assert!(validate_region(&missing, checkpoints).is_err());

        let mut endpoint = measured_region_sample();
        endpoint.live_bytes_after = Some(9);
        assert!(validate_region(&endpoint, checkpoints).is_err());

        let mut intermediate = checkpoints;
        intermediate.inputs_and_sink_ready = Checkpoint {
            live_bytes: 11,
            allocated_bytes: 11,
            peak_live_bytes: 11,
            ..Checkpoint::default()
        };
        let mut low_peak = measured_region_sample();
        low_peak.region_peak_live_bytes = Some(10);
        assert!(validate_region(&low_peak, intermediate).is_err());
    }
}
