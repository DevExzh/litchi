//! Standalone phase attribution for the ordinary owned ODP append lifecycle.
//!
//! The existing `odp_existing_append_lifecycle` selector remains the exact
//! whole-lifecycle control. This opt-in command runs the same deterministic
//! corpus through either that control clock (`--mode lifecycle`) or a set of
//! five harness-only phase clocks (`--mode phases`). No production clock or
//! public API is added: the commit implementation is intentionally measured as
//! one public call because its internal stages have no instrumentation seam.

use crate::{
    CorpusManifest, HashingDiscardSink, SemanticShape, SinkSummary, WriteSizeBuckets,
    allocation_metrics, elapsed_ns, iteration_count, odp_existing_append, sha256_hex,
};
use serde::Serialize;
use std::{error::Error, ffi::OsString, fs::OpenOptions, io::Write, path::PathBuf, time::Instant};

const SCHEMA: &str = "litchi-odp-append-attribution-v1";
const MAX_SAMPLES: usize = 10_000;
const MAX_WARMUP: usize = 10_000;
const PHASE_NAMES: [&str; 5] = [
    "snapshot_open",
    "transaction",
    "add",
    "commit",
    "publication",
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Mode {
    Lifecycle,
    Phases,
}

impl Mode {
    const fn name(self) -> &'static str {
        match self {
            Self::Lifecycle => "lifecycle",
            Self::Phases => "phases",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Config {
    mode: Mode,
    shape: SemanticShape,
    warmup: usize,
    samples: usize,
    repeat: &'static str,
    output: PathBuf,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct CorpusIdentity {
    source_archive_sha256: &'static str,
    source_archive_bytes: usize,
    expected_output_sha256: &'static str,
    expected_output_bytes: usize,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct AllocatorIdentity {
    allocator: &'static str,
    instrumentation: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    counter_revision: Option<&'static str>,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct PhaseDescription {
    name: &'static str,
    scope: &'static str,
}

const PHASES: [PhaseDescription; 5] = [
    PhaseDescription {
        name: "snapshot_open",
        scope: "Snapshot::from_bytes(input Vec)",
    },
    PhaseDescription {
        name: "transaction",
        scope: "Snapshot::transaction()",
    },
    PhaseDescription {
        name: "add",
        scope: "Transaction::add(title, body)",
    },
    PhaseDescription {
        name: "commit",
        scope: "Transaction::commit() including Patch construction and readback",
    },
    PhaseDescription {
        name: "publication",
        scope: "HashingDiscardSink::write_all(committed snapshot bytes)",
    },
];

#[derive(Clone, Debug, Serialize)]
struct PreflightGates {
    source_manifest_bindings_verified: bool,
    output_manifest_bindings_verified: bool,
    untouched_members_verified: bool,
    opaque_member_compressed_identity_verified: bool,
    patch_replay_verified: bool,
    inverse_patch_verified: bool,
    stale_source_refusal_verified: bool,
    exact_noop_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct RuntimeGates {
    source_bytes_identity_verified: bool,
    candidate_bytes_identity_verified: bool,
    commit_changed_verified: bool,
    patch_non_noop_verified: bool,
    sink_digest_verified: bool,
    sink_length_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct PhaseMeasurement {
    name: &'static str,
    elapsed_ns: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation_metrics: Option<allocation_metrics::Sample>,
}

/// The attribution contract retains only logical sink counters.  The shared
/// sink also carries optional retention annotations used by other baselines;
/// those are deliberately projected out here so this report cannot imply a
/// retention claim.
#[derive(Clone, Copy, Debug, Serialize)]
struct SinkObservation {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    write_size_buckets: WriteSizeBuckets,
}

impl From<SinkSummary> for SinkObservation {
    fn from(summary: SinkSummary) -> Self {
        Self {
            accepted_bytes: summary.accepted_bytes,
            write_calls: summary.write_calls,
            largest_write: summary.largest_write,
            write_size_buckets: summary.write_size_buckets,
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct Row {
    sample_index: usize,
    lifecycle_ns: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    phase_sum_ns: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    boundary_gap_ns: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    phases: Option<Vec<PhaseMeasurement>>,
    source_sha256: String,
    source_bytes: usize,
    candidate_sha256: String,
    candidate_bytes: usize,
    sink_sha256: String,
    sink: SinkObservation,
    runtime_gates: RuntimeGates,
    #[serde(skip_serializing_if = "Option::is_none")]
    lifecycle_allocation_metrics: Option<allocation_metrics::Sample>,
}

struct TimingMeasurement {
    lifecycle_ns: u64,
    phases: Option<Vec<PhaseMeasurement>>,
    lifecycle_allocation_metrics: Option<allocation_metrics::Sample>,
}

#[derive(Clone, Debug, Serialize)]
struct Report {
    schema: &'static str,
    mode: &'static str,
    /// Alias for consumers that key lanes by the experiment scope.
    scope: &'static str,
    repeat: &'static str,
    shape: &'static str,
    warmup: usize,
    samples: usize,
    checked_iteration_count: usize,
    corpus_generator: &'static str,
    corpus: CorpusManifest,
    identity: CorpusIdentity,
    allocator: AllocatorIdentity,
    timing_scope: &'static str,
    phase_order: [&'static str; 5],
    phases: [PhaseDescription; 5],
    preflight_gates: PreflightGates,
    rows: Vec<Row>,
}

#[derive(Clone, Copy)]
struct ExpectedIdentity {
    source_sha256: &'static str,
    source_bytes: usize,
    output_sha256: &'static str,
    output_bytes: usize,
}

fn expected_identity(shape: SemanticShape) -> ExpectedIdentity {
    match shape {
        SemanticShape::Tiny => ExpectedIdentity {
            source_sha256: "92321a679c82b333416a478f7b1ab52b5876abad071a2298a1417e1ff0213e75",
            source_bytes: 4_684,
            output_sha256: "43533f3c8464983121760e20537585652afe6d2cbe4bab53136e9de19ea1b3f5",
            output_bytes: 4_697,
        },
        SemanticShape::Medium => ExpectedIdentity {
            source_sha256: "62f45803d9363c0a0afd9c35600487a9ac6f287bedd9c0aba37dbd9b7aa3a593",
            source_bytes: 54_646,
            output_sha256: "743d0d0a02c6316a7ccb4da15de6d8359173d815f8fc9496d53ace6a948b091b",
            output_bytes: 54_658,
        },
        SemanticShape::Large => ExpectedIdentity {
            source_sha256: "bda1bffb24b312cd872ee36b1c72e6da886f22fdf93086b727aecb2262c5563c",
            source_bytes: 105_112,
            output_sha256: "a7dbfb18387dc546beb567c231bc7762e9efd440d72883e5e2ac2dcd24294bf1",
            output_bytes: 105_124,
        },
    }
}

fn parse_usize(
    value: &str,
    flag: &str,
    minimum: usize,
    maximum: usize,
) -> Result<usize, Box<dyn Error>> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| format!("{flag} must be an unsigned decimal integer"))?;
    if !(minimum..=maximum).contains(&parsed) {
        return Err(format!("{flag} must be between {minimum} and {maximum}").into());
    }
    Ok(parsed)
}

fn parse_config(args: &[OsString]) -> Result<Config, Box<dyn Error>> {
    let mut mode = None;
    let mut shape = None;
    let mut warmup = None;
    let mut samples = None;
    let mut repeat = None;
    let mut output = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("odp-append-attribution argument is not valid UTF-8")?;
        let value = args
            .get(index + 1)
            .ok_or_else(|| format!("missing value for {flag}"))?;
        let value = value
            .to_str()
            .ok_or("odp-append-attribution argument value is not valid UTF-8")?;
        index += 2;
        match flag {
            "--mode" => {
                if mode.is_some() {
                    return Err("duplicate --mode".into());
                }
                mode = Some(match value {
                    "lifecycle" => Mode::Lifecycle,
                    "phases" => Mode::Phases,
                    _ => return Err("--mode must be lifecycle or phases".into()),
                });
            },
            "--shape" => {
                if shape.is_some() {
                    return Err("duplicate --shape".into());
                }
                shape = Some(match value {
                    "tiny" => SemanticShape::Tiny,
                    "medium" => SemanticShape::Medium,
                    "large" => SemanticShape::Large,
                    _ => return Err("--shape must be tiny, medium, or large".into()),
                });
            },
            "--warmup" => {
                if warmup.is_some() {
                    return Err("duplicate --warmup".into());
                }
                warmup = Some(parse_usize(value, "--warmup", 0, MAX_WARMUP)?);
            },
            "--samples" => {
                if samples.is_some() {
                    return Err("duplicate --samples".into());
                }
                samples = Some(parse_usize(value, "--samples", 1, MAX_SAMPLES)?);
            },
            "--repeat" => {
                if repeat.is_some() {
                    return Err("duplicate --repeat".into());
                }
                repeat = Some(match value {
                    "R1" => "R1",
                    "R2" => "R2",
                    "diagnostic" => "diagnostic",
                    _ => return Err("--repeat must be R1, R2, or diagnostic".into()),
                });
            },
            "--output" => {
                if output.is_some() {
                    return Err("duplicate --output".into());
                }
                if value.is_empty() || value == "-" {
                    return Err("--output must be a create_new file path".into());
                }
                output = Some(PathBuf::from(value));
            },
            _ => return Err(format!("unknown odp-append-attribution argument: {flag}").into()),
        }
    }
    Ok(Config {
        mode: mode.ok_or("missing --mode")?,
        shape: shape.ok_or("missing --shape")?,
        warmup: warmup.ok_or("missing --warmup")?,
        samples: samples.ok_or("missing --samples")?,
        repeat: repeat.ok_or("missing --repeat")?,
        output: output.ok_or("missing --output")?,
    })
}

fn preflight_gates(corpus: &odp_existing_append::OdpExistingAppendCorpus) -> PreflightGates {
    let gates = corpus.preflight_gates();
    PreflightGates {
        source_manifest_bindings_verified: gates.source_manifest_bindings_verified,
        output_manifest_bindings_verified: gates.output_manifest_bindings_verified,
        untouched_members_verified: gates.untouched_members_verified,
        opaque_member_compressed_identity_verified: gates
            .opaque_member_compressed_identity_verified,
        patch_replay_verified: gates.patch_replay_verified,
        inverse_patch_verified: gates.inverse_patch_verified,
        stale_source_refusal_verified: gates.stale_source_refusal_verified,
        exact_noop_verified: gates.exact_noop_verified,
    }
}

fn verify_corpus_identity(
    corpus: &odp_existing_append::OdpExistingAppendCorpus,
) -> Result<CorpusIdentity, Box<dyn Error>> {
    let expected = expected_identity(corpus.shape);
    let observed_source = &corpus.corpus.manifest;
    if observed_source.archive_sha256 != expected.source_sha256
        || observed_source.archive_bytes != expected.source_bytes
        || corpus.corpus.archive.len() != expected.source_bytes
        || sha256_hex(corpus.corpus.archive.as_slice()) != expected.source_sha256
        || corpus.expected_output_sha256() != expected.output_sha256
        || corpus.expected_output().len() != expected.output_bytes
        || sha256_hex(corpus.expected_output()) != expected.output_sha256
    {
        return Err(format!(
            "ODP attribution {} corpus identity differs from the frozen 0457 corpus",
            corpus.shape.name()
        )
        .into());
    }
    Ok(CorpusIdentity {
        source_archive_sha256: expected.source_sha256,
        source_archive_bytes: expected.source_bytes,
        expected_output_sha256: expected.output_sha256,
        expected_output_bytes: expected.output_bytes,
    })
}

#[inline(never)]
fn phase_snapshot_open(
    input: Vec<u8>,
) -> litchi_core::Result<litchi_odp::authoring::edit::Snapshot> {
    let result = litchi_odp::authoring::edit::Snapshot::from_bytes(input);
    std::hint::black_box(&result);
    result
}

#[inline(never)]
fn phase_transaction(
    source: &litchi_odp::authoring::edit::Snapshot,
) -> litchi_core::Result<litchi_odp::authoring::edit::Transaction> {
    let result = source.transaction();
    std::hint::black_box(&result);
    result
}

#[inline(never)]
fn phase_add(
    transaction: &mut litchi_odp::authoring::edit::Transaction,
    title: &str,
    body: &str,
) -> litchi_core::Result<()> {
    let result = transaction.add(title, body);
    std::hint::black_box(&result);
    result
}

#[inline(never)]
fn phase_commit(
    transaction: litchi_odp::authoring::edit::Transaction,
) -> litchi_core::Result<litchi_odp::authoring::edit::Commit> {
    let result = transaction.commit();
    std::hint::black_box(&result);
    result
}

#[inline(never)]
fn phase_publication(sink: &mut HashingDiscardSink, bytes: &[u8]) -> std::io::Result<()> {
    let result = sink.write_all(bytes);
    std::hint::black_box(&result);
    result
}

fn measure_phase<T, E, F>(
    name: &'static str,
    call: F,
) -> Result<(T, PhaseMeasurement), Box<dyn Error>>
where
    E: Error + 'static,
    F: FnOnce() -> std::result::Result<T, E>,
{
    let allocation_region = allocation_metrics::begin();
    let started = Instant::now();
    let result = call();
    let elapsed = elapsed_ns(started.elapsed())?;
    let allocation_metrics = allocation_region.finish();
    let value = result.map_err(|error| Box::new(error) as Box<dyn Error>)?;
    Ok((
        value,
        PhaseMeasurement {
            name,
            elapsed_ns: elapsed,
            allocation_metrics,
        },
    ))
}

fn runtime_row(
    sample_index: usize,
    corpus: &odp_existing_append::OdpExistingAppendCorpus,
    source: &litchi_odp::authoring::edit::Snapshot,
    commit: &litchi_odp::authoring::edit::Commit,
    sink: HashingDiscardSink,
    measurement: TimingMeasurement,
) -> Result<Row, Box<dyn Error>> {
    let TimingMeasurement {
        lifecycle_ns,
        phases,
        lifecycle_allocation_metrics,
    } = measurement;
    let candidate = commit.snapshot().bytes();
    let source_bytes = source.bytes();
    let candidate_sha256 = sha256_hex(candidate);
    let (sink_summary, sink_sha256) = sink.finish();
    let sink = SinkObservation::from(sink_summary);
    let source_bytes_identity_verified = source_bytes == corpus.corpus.archive.as_slice();
    let candidate_bytes_identity_verified = candidate == corpus.expected_output();
    let commit_changed_verified = commit.changed();
    let patch_non_noop_verified = !commit.patch().is_noop();
    let sink_digest_verified = sink_sha256 == candidate_sha256;
    let sink_length_verified =
        sink.accepted_bytes == u64::try_from(corpus.expected_output().len())?;
    let runtime_gates = RuntimeGates {
        source_bytes_identity_verified,
        candidate_bytes_identity_verified,
        commit_changed_verified,
        patch_non_noop_verified,
        sink_digest_verified,
        sink_length_verified,
    };
    if !source_bytes_identity_verified
        || !candidate_bytes_identity_verified
        || !commit_changed_verified
        || !patch_non_noop_verified
        || !sink_digest_verified
        || !sink_length_verified
    {
        return Err("ODP attribution runtime correctness gate failed".into());
    }
    let (phase_sum_ns, boundary_gap_ns) = if let Some(phases) = &phases {
        let sum = phases.iter().try_fold(0_u64, |total, phase| {
            total
                .checked_add(phase.elapsed_ns)
                .ok_or("ODP attribution phase sum overflows u64")
        })?;
        let gap = lifecycle_ns
            .checked_sub(sum)
            .ok_or("ODP attribution phase sum exceeds the enclosing lifecycle span")?;
        (Some(sum), Some(gap))
    } else {
        (None, None)
    };
    Ok(Row {
        sample_index,
        lifecycle_ns,
        phase_sum_ns,
        boundary_gap_ns,
        phases,
        source_sha256: corpus.corpus.manifest.archive_sha256.clone(),
        source_bytes: source_bytes.len(),
        candidate_sha256,
        candidate_bytes: candidate.len(),
        sink_sha256,
        sink,
        runtime_gates,
        lifecycle_allocation_metrics,
    })
}

#[inline(never)]
fn run_lifecycle_iteration(
    corpus: &odp_existing_append::OdpExistingAppendCorpus,
    sample_index: usize,
) -> Result<Row, Box<dyn Error>> {
    let input = corpus.corpus.archive.clone();
    let title = odp_existing_append::appended_title(corpus.shape);
    let body = odp_existing_append::appended_body(corpus.shape);
    let mut sink = HashingDiscardSink::without_authoring_window(u64::try_from(
        corpus.expected_output().len(),
    )?);
    let allocation_region = allocation_metrics::begin();
    let started = Instant::now();
    let source = litchi_odp::authoring::edit::Snapshot::from_bytes(input)?;
    let mut transaction = source.transaction()?;
    transaction.add(&title, &body)?;
    let commit = transaction.commit()?;
    sink.write_all(commit.snapshot().bytes())?;
    let lifecycle_ns = elapsed_ns(started.elapsed())?;
    let lifecycle_allocation_metrics = allocation_region.finish();
    let row = runtime_row(
        sample_index,
        corpus,
        &source,
        &commit,
        sink,
        TimingMeasurement {
            lifecycle_ns,
            phases: None,
            lifecycle_allocation_metrics,
        },
    )?;
    std::hint::black_box(row.candidate_sha256.as_str());
    drop(commit);
    drop(source);
    Ok(row)
}

#[inline(never)]
fn run_phases_iteration(
    corpus: &odp_existing_append::OdpExistingAppendCorpus,
    sample_index: usize,
) -> Result<Row, Box<dyn Error>> {
    let input = corpus.corpus.archive.clone();
    let title = odp_existing_append::appended_title(corpus.shape);
    let body = odp_existing_append::appended_body(corpus.shape);
    let mut sink = HashingDiscardSink::without_authoring_window(u64::try_from(
        corpus.expected_output().len(),
    )?);
    let lifecycle_started = Instant::now();
    let (source, open) = measure_phase("snapshot_open", || phase_snapshot_open(input))?;
    let (mut transaction, transaction_phase) =
        measure_phase("transaction", || phase_transaction(&source))?;
    let (_, add_phase) = measure_phase("add", || phase_add(&mut transaction, &title, &body))?;
    let (commit, commit_phase) = measure_phase("commit", || phase_commit(transaction))?;
    let (_, publication_phase) = measure_phase("publication", || {
        phase_publication(&mut sink, commit.snapshot().bytes())
    })?;
    let lifecycle_ns = elapsed_ns(lifecycle_started.elapsed())?;
    let row = runtime_row(
        sample_index,
        corpus,
        &source,
        &commit,
        sink,
        TimingMeasurement {
            lifecycle_ns,
            phases: Some(vec![
                open,
                transaction_phase,
                add_phase,
                commit_phase,
                publication_phase,
            ]),
            lifecycle_allocation_metrics: None,
        },
    )?;
    std::hint::black_box(row.candidate_sha256.as_str());
    drop(commit);
    drop(source);
    Ok(row)
}

fn run_capture(config: &Config) -> Result<Report, Box<dyn Error>> {
    let corpus = odp_existing_append::build_odp_existing_append_corpus(config.shape)?;
    let identity = verify_corpus_identity(&corpus)?;
    let checked_iteration_count = iteration_count(config.warmup, config.samples)?;
    let mut rows = Vec::with_capacity(config.samples);
    for iteration in 0..checked_iteration_count {
        let row = match config.mode {
            Mode::Lifecycle => run_lifecycle_iteration(&corpus, iteration)?,
            Mode::Phases => run_phases_iteration(&corpus, iteration)?,
        };
        if iteration >= config.warmup {
            rows.push(Row {
                sample_index: iteration - config.warmup,
                ..row
            });
        }
        // The warm-up row is fully gated but is deliberately discarded. Keep
        // the result alive only until all post-clock checks have completed.
        if iteration < config.warmup {
            std::hint::black_box(iteration);
        }
    }
    Ok(Report {
        schema: SCHEMA,
        mode: config.mode.name(),
        scope: config.mode.name(),
        repeat: config.repeat,
        shape: config.shape.name(),
        warmup: config.warmup,
        samples: config.samples,
        checked_iteration_count,
        corpus_generator: odp_existing_append::ODP_EXISTING_APPEND_CORPUS_GENERATOR,
        corpus: corpus.corpus.manifest.clone(),
        identity,
        allocator: AllocatorIdentity {
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
        timing_scope: match config.mode {
            Mode::Lifecycle => {
                "input/title/body/sink construction outside; one Instant includes Snapshot::from_bytes, Snapshot::transaction, Transaction::add, Transaction::commit, and sink write_all; all hashes, gates, report assembly, and drops outside"
            },
            Mode::Phases => {
                "input/title/body/sink construction outside; lifecycle_ns encloses five harness phase clocks and their boundary instrumentation; each phase clock includes only its named public call; hashes, gates, report assembly, and drops outside"
            },
        },
        phase_order: PHASE_NAMES,
        phases: PHASES,
        preflight_gates: preflight_gates(&corpus),
        rows,
    })
}

/// Run the standalone ODP attribution command after `main` consumes its selector.
///
/// # Errors
///
/// Returns an error when command-line controls are invalid, corpus construction
/// or a correctness gate fails, or the create-new report cannot be written.
pub fn run_from_args<I>(args: I) -> Result<(), Box<dyn Error>>
where
    I: IntoIterator<Item = OsString>,
{
    let args = args.into_iter().collect::<Vec<_>>();
    let config = parse_config(&args)?;
    let report = run_capture(&config)?;
    let bytes = serde_json::to_vec_pretty(&report)?;
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&config.output)?;
    output.write_all(&bytes)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn args(values: &[&str]) -> Vec<OsString> {
        values.iter().map(OsString::from).collect()
    }

    fn config(mode: Mode) -> Config {
        Config {
            mode,
            shape: SemanticShape::Tiny,
            warmup: 0,
            samples: 1,
            repeat: "R1",
            output: PathBuf::from("unused.json"),
        }
    }

    #[test]
    fn parser_rejects_missing_and_invalid_controls() {
        assert!(parse_config(&[]).is_err());
        let invalid_mode = args(&[
            "--mode",
            "other",
            "--shape",
            "tiny",
            "--warmup",
            "0",
            "--samples",
            "1",
            "--repeat",
            "R1",
            "--output",
            "out.json",
        ]);
        assert!(parse_config(&invalid_mode).is_err());
        let duplicate = args(&[
            "--mode",
            "lifecycle",
            "--mode",
            "phases",
            "--shape",
            "tiny",
            "--warmup",
            "0",
            "--samples",
            "1",
            "--repeat",
            "R1",
            "--output",
            "out.json",
        ]);
        assert!(parse_config(&duplicate).is_err());
        let invalid_repeat = args(&[
            "--mode",
            "lifecycle",
            "--shape",
            "tiny",
            "--warmup",
            "0",
            "--samples",
            "1",
            "--repeat",
            "R3",
            "--output",
            "out.json",
        ]);
        assert!(parse_config(&invalid_repeat).is_err());
        let diagnostic = args(&[
            "--mode",
            "phases",
            "--shape",
            "large",
            "--warmup",
            "3",
            "--samples",
            "100",
            "--repeat",
            "diagnostic",
            "--output",
            "out.json",
        ]);
        assert_eq!(parse_config(&diagnostic).unwrap().repeat, "diagnostic");
    }

    #[test]
    fn phase_order_and_frozen_identities_are_stable() {
        assert_eq!(
            PHASE_NAMES,
            [
                "snapshot_open",
                "transaction",
                "add",
                "commit",
                "publication",
            ]
        );
        let tiny = expected_identity(SemanticShape::Tiny);
        assert_eq!(tiny.source_bytes, 4_684);
        assert_eq!(tiny.output_bytes, 4_697);
        assert_eq!(config(Mode::Lifecycle).mode.name(), "lifecycle");
    }

    #[test]
    fn phases_capture_has_checked_output_and_sum_gap() {
        let report = run_capture(&config(Mode::Phases)).unwrap();
        assert_eq!(report.schema, SCHEMA);
        assert_eq!(report.rows.len(), 1);
        let row = &report.rows[0];
        let phases = row.phases.as_ref().unwrap();
        assert_eq!(phases.len(), PHASE_NAMES.len());
        assert_eq!(
            phases.iter().map(|phase| phase.name).collect::<Vec<_>>(),
            PHASE_NAMES.to_vec()
        );
        assert_eq!(row.candidate_sha256, report.identity.expected_output_sha256);
        assert!(row.runtime_gates.candidate_bytes_identity_verified);
        assert!(row.runtime_gates.sink_digest_verified);
        assert!(row.phase_sum_ns.unwrap() <= row.lifecycle_ns);
        assert_eq!(
            row.boundary_gap_ns.unwrap(),
            row.lifecycle_ns - row.phase_sum_ns.unwrap()
        );
    }

    #[test]
    fn lifecycle_capture_has_no_phase_rows_and_exact_output() {
        let report = run_capture(&config(Mode::Lifecycle)).unwrap();
        assert_eq!(report.rows.len(), 1);
        let row = &report.rows[0];
        assert!(row.phases.is_none());
        assert!(row.phase_sum_ns.is_none());
        assert_eq!(row.candidate_sha256, report.identity.expected_output_sha256);
        assert!(row.runtime_gates.source_bytes_identity_verified);
    }
}
