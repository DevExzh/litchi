//! Matched public PPTX creation through ordinary and explicit metadata
//! storage.
//!
//! The control and spool lanes use the same deterministic plain-text PPTX
//! producer and the same finite writer limits.  Corpus construction and the
//! complete physical/semantic/relationship oracle are outside the timed
//! operation.  A timed spool lane opens a caller-selected file, runs the
//! public metadata-spool constructor, finishes the package, and closes that
//! file before the elapsed interval ends.  The output sink retains only a
//! SHA-256 state and scalar counters.
//!
//! The caller must keep the selected scratch directory exclusively accessible
//! during the run. Scratch files are removed after successful operations;
//! errors leave them for caller inspection, including existing files refused
//! by `create_new`.

use std::{
    collections::BTreeSet,
    ffi::OsString,
    fs::{self, File, OpenOptions},
    io::{self, Write},
    path::{Path, PathBuf},
    time::Instant,
};

use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_pptx::{
    StreamingPresentationOptions, StreamingPresentationScratchLimits, StreamingPresentationWriter,
};
use serde::Serialize;
use sha2::{Digest as _, Sha256};

use super::{allocation_metrics, pptx_streaming_create, process_metrics};
use pptx_streaming_create::PptxStreamingCorpus;

const SCHEMA: &str = "pptx-metadata-spool-v1";
const DEFAULT_COUNTS: &[usize] = &[8, 256, 8_192];
const DEFAULT_SAMPLES: usize = 30;
const DEFAULT_WARMUPS: usize = 3;
const DEFAULT_REPEATS: usize = 2;
const DEFAULT_SPOOL_MAX_BYTES: u64 = 64 * 1024 * 1024;
const DEFAULT_SPOOL_BUFFER_BYTES: usize = 16 * 1024;
const FIXED_MEMBER_COUNT: usize = 37;
const MEMBERS_PER_SLIDE: usize = 2;

type BenchResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum Mode {
    Control,
    Spool,
}

impl Mode {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "control" => Ok(Self::Control),
            "spool" => Ok(Self::Spool),
            _ => Err(format!("invalid --mode {value:?}; expected control or spool").into()),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Control => "control",
            Self::Spool => "spool",
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct ScratchLimits {
    max_bytes: u64,
    buffer_bytes: usize,
}

#[derive(Clone, Debug)]
struct Config {
    modes: Vec<Mode>,
    counts: Vec<usize>,
    samples: usize,
    warmups: usize,
    repeats: usize,
    spool_dir: Option<PathBuf>,
    scratch_limits: ScratchLimits,
    json_path: Option<PathBuf>,
}

#[derive(Clone, Debug)]
struct PreparedCase {
    corpus_index: usize,
    mode: Mode,
    expected_sha256: String,
    expected_bytes: u64,
}

#[derive(Debug)]
struct HashingSink {
    accepted_bytes: u64,
    write_calls: u64,
    digest: Sha256,
}

#[derive(Clone, Debug, Serialize)]
struct SinkSummary {
    accepted_bytes: u64,
    write_calls: u64,
    sha256: String,
}

impl HashingSink {
    fn new() -> Self {
        Self {
            accepted_bytes: 0,
            write_calls: 0,
            digest: Sha256::new(),
        }
    }

    /// Finalize the fixed sink outside the timed operation.  The timed
    /// operation includes every digest update made by `write`, but not this
    /// endpoint allocation/finalization work.
    fn finish(self) -> BenchResult<SinkSummary> {
        let digest = self.digest.finalize();
        Ok(SinkSummary {
            accepted_bytes: self.accepted_bytes,
            write_calls: self.write_calls,
            sha256: hex_digest(&digest),
        })
    }
}

impl Write for HashingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.accepted_bytes = self
            .accepted_bytes
            .checked_add(u64::try_from(bytes.len()).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "output length does not fit u64",
                )
            })?)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "output length overflow"))?;
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "write-call overflow"))?;
        self.digest.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct OracleBytes {
    bytes: Vec<u8>,
    summary: SinkSummary,
    scratch_bytes: Option<u64>,
}

#[derive(Debug)]
struct TimedOperation {
    elapsed_ns: u64,
    sink: SinkSummary,
    scratch_bytes: Option<u64>,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Debug, Serialize)]
struct CorpusRecord {
    slide_count: usize,
    entry_count: usize,
    input_text_bytes: usize,
    source_archive_bytes: usize,
    source_archive_sha256: String,
    semantic_sha256: String,
    full_text_sha256: String,
}

#[derive(Debug, Serialize)]
struct OracleRecord {
    mode: Mode,
    slide_count: usize,
    entry_count: usize,
    source_archive_bytes: usize,
    source_archive_sha256: String,
    output_bytes: u64,
    output_sha256: String,
    scratch_bytes: Option<u64>,
    byte_exact_control_match: bool,
    every_physical_member_verified: bool,
    every_slide_semantic_verified: bool,
    presentation_graph_verified: bool,
    slide_geometry_verified: bool,
    text_digest_verified: bool,
}

#[derive(Debug, Serialize)]
struct OperationRecord {
    mode: Mode,
    slide_count: usize,
    entry_count: usize,
    repeat: usize,
    sample: usize,
    source_archive_bytes: usize,
    source_archive_sha256: String,
    input_text_bytes: usize,
    elapsed_ns: u64,
    output_bytes: u64,
    output_write_calls: u64,
    output_sha256: String,
    output_matches_oracle: bool,
    scratch_bytes: Option<u64>,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    instrumentation: &'static str,
    allocator: &'static str,
    timing_scope: &'static str,
    oracle_scope: &'static str,
    cleanup_scope: &'static str,
    control_storage_policy: &'static str,
    spool_storage_policy: &'static str,
    samples: usize,
    warmups: usize,
    repeats: usize,
    counts: Vec<usize>,
    modes: Vec<Mode>,
    spool_max_bytes: u64,
    spool_buffer_bytes: usize,
    corpora: Vec<CorpusRecord>,
    cases: Vec<OracleRecord>,
    operations: Vec<OperationRecord>,
}

/// Run the matched public PPTX metadata-storage diagnostic.
pub fn run_from_args<I>(args: I) -> BenchResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let config = parse_args(args)?;
    if config.modes.contains(&Mode::Spool) {
        let spool_dir = config.spool_dir.as_ref().ok_or(
            "--spool-dir is required for spool mode; the benchmark never selects an ambient path",
        )?;
        fs::create_dir_all(spool_dir)?;
    } else if let Some(spool_dir) = config.spool_dir.as_ref() {
        // Supplying an explicit directory for a control-only lane requests
        // the same paired oracle preflight used by formal captures.
        fs::create_dir_all(spool_dir)?;
    }

    let corpora = config
        .counts
        .iter()
        .copied()
        .map(pptx_streaming_create::build_corpus_for_slide_count)
        .collect::<BenchResult<Vec<_>>>()?;
    let corpus_records = corpora
        .iter()
        .map(corpus_record)
        .collect::<BenchResult<Vec<_>>>()?;

    let mut prepared = Vec::new();
    let mut cases = Vec::new();
    for (corpus_index, corpus) in corpora.iter().enumerate() {
        let count = slide_count(corpus);
        let expected_entries = expected_entry_count(count)?;
        if pptx_streaming_create::corpus_archive_member_count(corpus) != expected_entries {
            return Err(format!(
                "PPTX corpus has the wrong physical member count for {count} slides"
            )
            .into());
        }

        let control = oracle_archive(corpus, Mode::Control, None, config.scratch_limits)?;
        cases.push(oracle_record(corpus, Mode::Control, &control, true));
        let control_sha = control.summary.sha256.clone();
        let control_bytes = control.summary.accepted_bytes;

        if config.spool_dir.is_some() {
            let spool_path = config
                .spool_dir
                .as_deref()
                .expect("spool directory was checked above")
                .join(format!("pptx-metadata-spool-{count}-oracle-spool.tmp"));
            let spool = oracle_archive(
                corpus,
                Mode::Spool,
                Some(&spool_path),
                config.scratch_limits,
            )?;
            // A failed create_new may refer to a caller's existing file.
            // On any operation error, retain scratch for caller inspection.
            if spool.bytes != control.bytes {
                return Err(format!(
                    "PPTX metadata-spool oracle differs from control for {count} slides"
                )
                .into());
            }
            remove_owned_file(&spool_path)?;
            cases.push(oracle_record(corpus, Mode::Spool, &spool, true));
        }

        for mode in &config.modes {
            prepared.push(PreparedCase {
                corpus_index,
                mode: *mode,
                expected_sha256: control_sha.clone(),
                expected_bytes: control_bytes,
            });
        }
    }

    let mut operations = Vec::new();
    let mut order: Vec<usize> = (0..prepared.len()).collect();
    for repeat in 0..config.repeats {
        if repeat % 2 == 1 {
            order.reverse();
        }
        for &case_index in &order {
            let case = &prepared[case_index];
            let corpus = &corpora[case.corpus_index];
            for sample in 0..config.warmups {
                let path = operation_spool_path(&config, case, repeat, sample, true);
                let timed =
                    run_timed_operation(corpus, case.mode, path.as_deref(), config.scratch_limits)?;
                if !matches_expected(case, &timed.sink) {
                    return Err(format!(
                        "PPTX metadata-spool warmup output failed oracle for {} slides / {}",
                        slide_count(corpus),
                        case.mode.label()
                    )
                    .into());
                }
                cleanup_operation_path(path.as_deref())?;
            }
            for sample in 0..config.samples {
                let path = operation_spool_path(&config, case, repeat, sample, false);
                let timed =
                    run_timed_operation(corpus, case.mode, path.as_deref(), config.scratch_limits)?;
                let matches = matches_expected(case, &timed.sink);
                if !matches {
                    return Err(format!(
                        "PPTX metadata-spool sample output failed oracle for {} slides / {}",
                        slide_count(corpus),
                        case.mode.label()
                    )
                    .into());
                }
                cleanup_operation_path(path.as_deref())?;
                operations.push(operation_record(
                    corpus, case, repeat, sample, timed, matches,
                ));
            }
        }
    }

    let report = Report {
        schema: SCHEMA,
        instrumentation: allocation_metrics::instrumentation_identity(),
        allocator: allocation_metrics::allocator_identity(),
        timing_scope: "deterministic corpus references, limits, expected output, and semantic/physical oracles are prepared before the clock; the timed operation constructs the public writer and its name/metadata plan, creates fresh deterministic text Strings, opens/creates the caller-selected scratch File for the spool lane, writes every slide and relationship, completes writer.finish including central-directory publication, updates the fixed SHA-256 sink, flushes and closes the scratch File; process and allocator endpoint observations, sink digest finalization, oracle reopening, and unlink are outside",
        oracle_scope: "the control oracle is generated for every selected slide count; whenever --spool-dir is supplied an explicit-spool oracle is generated in every mode lane, byte-compared with control, and both are passed through the existing exact 37+2N physical-member, per-slide text/geometry, and presentation relationship-graph oracle",
        cleanup_scope: "the caller-selected spool directory is created before capture; each oracle and timed spool uses create_new and is measured through File metadata after the writer closes, then removed after successful endpoint observations; errors retain scratch files for caller inspection, including pre-existing files refused by create_new; the caller must keep the scratch directory exclusively accessible during the run; no ambient temporary path is selected",
        control_storage_policy: "StreamingPresentationWriter::with_options with the existing ordinary name and metadata owners",
        spool_storage_policy: "StreamingPresentationWriter::with_options_and_metadata_spool with caller-owned read/write/seek std::fs::File and explicit StreamingPresentationScratchLimits",
        samples: config.samples,
        warmups: config.warmups,
        repeats: config.repeats,
        counts: config.counts,
        modes: config.modes,
        spool_max_bytes: config.scratch_limits.max_bytes,
        spool_buffer_bytes: config.scratch_limits.buffer_bytes,
        corpora: corpus_records,
        cases,
        operations,
    };
    let json = serde_json::to_string_pretty(&report)?;
    if let Some(path) = config.json_path {
        fs::write(path, json.as_bytes())?;
    } else {
        println!("{json}");
    }
    Ok(())
}

fn corpus_record(corpus: &PptxStreamingCorpus) -> BenchResult<CorpusRecord> {
    let slides = slide_count(corpus);
    Ok(CorpusRecord {
        slide_count: slides,
        entry_count: expected_entry_count(slides)?,
        input_text_bytes: pptx_streaming_create::corpus_source_bytes(corpus),
        source_archive_bytes: pptx_streaming_create::corpus_archive_bytes(corpus),
        source_archive_sha256: pptx_streaming_create::corpus_archive_sha256(corpus).to_owned(),
        semantic_sha256: pptx_streaming_create::corpus_semantic_sha256(corpus).to_owned(),
        full_text_sha256: pptx_streaming_create::corpus_full_text_sha256(corpus).to_owned(),
    })
}

fn oracle_record(
    corpus: &PptxStreamingCorpus,
    mode: Mode,
    oracle: &OracleBytes,
    byte_exact_control_match: bool,
) -> OracleRecord {
    OracleRecord {
        mode,
        slide_count: slide_count(corpus),
        entry_count: expected_entry_count(slide_count(corpus)).expect("validated corpus count"),
        source_archive_bytes: pptx_streaming_create::corpus_archive_bytes(corpus),
        source_archive_sha256: pptx_streaming_create::corpus_archive_sha256(corpus).to_owned(),
        output_bytes: oracle.summary.accepted_bytes,
        output_sha256: oracle.summary.sha256.clone(),
        scratch_bytes: oracle.scratch_bytes,
        byte_exact_control_match,
        every_physical_member_verified: true,
        every_slide_semantic_verified: true,
        presentation_graph_verified: true,
        slide_geometry_verified: true,
        text_digest_verified: true,
    }
}

fn operation_record(
    corpus: &PptxStreamingCorpus,
    case: &PreparedCase,
    repeat: usize,
    sample: usize,
    timed: TimedOperation,
    output_matches_oracle: bool,
) -> OperationRecord {
    OperationRecord {
        mode: case.mode,
        slide_count: slide_count(corpus),
        entry_count: expected_entry_count(slide_count(corpus)).expect("validated corpus count"),
        repeat,
        sample,
        source_archive_bytes: pptx_streaming_create::corpus_archive_bytes(corpus),
        source_archive_sha256: pptx_streaming_create::corpus_archive_sha256(corpus).to_owned(),
        input_text_bytes: pptx_streaming_create::corpus_source_bytes(corpus),
        elapsed_ns: timed.elapsed_ns,
        output_bytes: timed.sink.accepted_bytes,
        output_write_calls: timed.sink.write_calls,
        output_sha256: timed.sink.sha256,
        output_matches_oracle,
        scratch_bytes: timed.scratch_bytes,
        allocation: timed.allocation,
        process: timed.process,
    }
}

fn run_timed_operation(
    corpus: &PptxStreamingCorpus,
    mode: Mode,
    spool_path: Option<&Path>,
    scratch_limits: ScratchLimits,
) -> BenchResult<TimedOperation> {
    let process_before = process_metrics::Snapshot::read().ok();
    let allocation_region = allocation_metrics::begin();
    let started = Instant::now();
    let output_result = create_archive(corpus, mode, spool_path, scratch_limits);
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos())
        .map_err(|_| "PPTX operation elapsed time does not fit u64")?;

    // The writer owns the spool File and returns only after `finish` has
    // released it.  Metadata length is intentionally observed after the
    // clock, while sink digest finalization is also outside the interval.
    let allocation = allocation_region.finish();
    let process = process_metrics::Snapshot::read()
        .ok()
        .zip(process_before)
        .map(|(after, before)| after.delta(before));
    let output = output_result?;
    let scratch_bytes = match mode {
        Mode::Control => None,
        Mode::Spool => {
            let path = spool_path.ok_or("spool mode did not receive an explicit path")?;
            Some(fs::metadata(path)?.len())
        },
    };
    let sink = output.finish()?;
    Ok(TimedOperation {
        elapsed_ns,
        sink,
        scratch_bytes,
        allocation,
        process,
    })
}

fn create_archive(
    corpus: &PptxStreamingCorpus,
    mode: Mode,
    spool_path: Option<&Path>,
    scratch_limits: ScratchLimits,
) -> BenchResult<HashingSink> {
    let limits = pptx_streaming_create::corpus_limits(corpus)?;
    let slides = slide_count(corpus);
    let options = StreamingPresentationOptions::standard();
    let writer = match mode {
        Mode::Control => {
            StreamingPresentationWriter::with_options(HashingSink::new(), slides, options, limits)?
        },
        Mode::Spool => {
            let path = spool_path.ok_or("spool mode requires a caller-selected spool path")?;
            let spool = open_spool(path)?;
            let scratch = StreamingPresentationScratchLimits {
                max_bytes: scratch_limits.max_bytes,
                buffer_bytes: scratch_limits.buffer_bytes,
            };
            StreamingPresentationWriter::with_options_and_metadata_spool(
                HashingSink::new(),
                slides,
                options,
                limits,
                spool,
                scratch,
            )?
        },
    };
    pptx_streaming_create::finish_pptx_stream(writer, pptx_streaming_create::corpus_spec(corpus))
}

fn oracle_archive(
    corpus: &PptxStreamingCorpus,
    mode: Mode,
    spool_path: Option<&Path>,
    scratch_limits: ScratchLimits,
) -> BenchResult<OracleBytes> {
    let output = match mode {
        Mode::Control => create_materialized_control(corpus, scratch_limits)?,
        Mode::Spool => create_materialized_spool(corpus, spool_path, scratch_limits)?,
    };
    validate_oracle(corpus, &output.0)?;
    let summary = summary_for_bytes(&output.0)?;
    Ok(OracleBytes {
        bytes: output.0,
        summary,
        scratch_bytes: output.1,
    })
}

fn create_materialized_control(
    corpus: &PptxStreamingCorpus,
    _scratch_limits: ScratchLimits,
) -> BenchResult<(Vec<u8>, Option<u64>)> {
    let limits = pptx_streaming_create::corpus_limits(corpus)?;
    let writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        slide_count(corpus),
        StreamingPresentationOptions::standard(),
        limits,
    )?;
    Ok((
        pptx_streaming_create::finish_pptx_stream(
            writer,
            pptx_streaming_create::corpus_spec(corpus),
        )?,
        None,
    ))
}

fn create_materialized_spool(
    corpus: &PptxStreamingCorpus,
    spool_path: Option<&Path>,
    scratch_limits: ScratchLimits,
) -> BenchResult<(Vec<u8>, Option<u64>)> {
    let path = spool_path.ok_or("spool oracle requires a caller-selected spool path")?;
    let limits = pptx_streaming_create::corpus_limits(corpus)?;
    let scratch = StreamingPresentationScratchLimits {
        max_bytes: scratch_limits.max_bytes,
        buffer_bytes: scratch_limits.buffer_bytes,
    };
    let writer = StreamingPresentationWriter::with_options_and_metadata_spool(
        Vec::new(),
        slide_count(corpus),
        StreamingPresentationOptions::standard(),
        limits,
        open_spool(path)?,
        scratch,
    )?;
    let bytes = pptx_streaming_create::finish_pptx_stream(
        writer,
        pptx_streaming_create::corpus_spec(corpus),
    )?;
    let scratch_bytes = fs::metadata(path)?.len();
    Ok((bytes, Some(scratch_bytes)))
}

fn validate_oracle(corpus: &PptxStreamingCorpus, bytes: &[u8]) -> BenchResult<()> {
    let expected = expected_entry_count(slide_count(corpus))?;
    if pptx_streaming_create::corpus_archive_member_count(corpus) != expected {
        return Err("PPTX corpus member-count identity is inconsistent".into());
    }
    reopen_every_physical_member(bytes, expected)?;
    pptx_streaming_create::validate_materialized_archive(bytes, corpus)?;
    Ok(())
}

/// Reopen and decompress every physical member before the richer PPTX oracle
/// runs.  `inspect_materialized_archive` proves the exact member names and
/// graph semantics; this independent pass makes the physical payload gate
/// explicit for every fixed and per-slide member.
fn reopen_every_physical_member(bytes: &[u8], expected: usize) -> BenchResult<()> {
    let package = PhysPkgReader::new(bytes)?;
    let names = package.member_names()?;
    if names.len() != expected {
        return Err(format!(
            "PPTX physical oracle reopened {} members, expected {expected}",
            names.len()
        )
        .into());
    }
    let mut seen = BTreeSet::new();
    for name in names {
        if !seen.insert(name.clone()) {
            return Err(format!("PPTX physical oracle found duplicate member {name}").into());
        }
        let _payload = package.read_member(&name)?;
    }
    Ok(())
}

fn summary_for_bytes(bytes: &[u8]) -> BenchResult<SinkSummary> {
    let mut sink = HashingSink::new();
    sink.write_all(bytes)?;
    sink.finish()
}

fn matches_expected(case: &PreparedCase, sink: &SinkSummary) -> bool {
    sink.accepted_bytes == case.expected_bytes && sink.sha256 == case.expected_sha256
}

fn slide_count(corpus: &PptxStreamingCorpus) -> usize {
    pptx_streaming_create::corpus_manifest_slide_count(corpus)
}

fn expected_entry_count(slides: usize) -> BenchResult<usize> {
    FIXED_MEMBER_COUNT
        .checked_add(
            slides
                .checked_mul(MEMBERS_PER_SLIDE)
                .ok_or("PPTX physical member count overflows usize")?,
        )
        .ok_or_else(|| "PPTX physical member count overflows usize".into())
}

fn open_spool(path: &Path) -> io::Result<File> {
    OpenOptions::new()
        .create_new(true)
        .read(true)
        .write(true)
        .open(path)
}

fn operation_spool_path(
    config: &Config,
    case: &PreparedCase,
    repeat: usize,
    sample: usize,
    warmup: bool,
) -> Option<PathBuf> {
    (case.mode == Mode::Spool).then(|| {
        let count = config.counts[case.corpus_index];
        config
            .spool_dir
            .as_deref()
            .expect("spool directory was checked above")
            .join(format!(
                "pptx-metadata-spool-{count}-{}-{repeat}-{sample}-{}.tmp",
                case.mode.label(),
                if warmup { "warmup" } else { "sample" }
            ))
    })
}

fn cleanup_operation_path(path: Option<&Path>) -> io::Result<()> {
    let Some(path) = path else { return Ok(()) };
    remove_owned_file(path)
}

fn remove_owned_file(path: &Path) -> io::Result<()> {
    match fs::remove_file(path) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error),
    }
}

fn parse_args<I>(args: I) -> BenchResult<Config>
where
    I: IntoIterator<Item = OsString>,
{
    let mut modes = vec![Mode::Control];
    let mut counts = DEFAULT_COUNTS.to_vec();
    let mut samples = DEFAULT_SAMPLES;
    let mut warmups = DEFAULT_WARMUPS;
    let mut repeats = DEFAULT_REPEATS;
    let mut spool_dir = None;
    let mut scratch_limits = ScratchLimits {
        max_bytes: DEFAULT_SPOOL_MAX_BYTES,
        buffer_bytes: DEFAULT_SPOOL_BUFFER_BYTES,
    };
    let mut json_path = None;
    let mut values = args.into_iter();
    while let Some(argument) = values.next() {
        let flag = argument.to_string_lossy();
        match flag.as_ref() {
            "--mode" => modes = parse_modes(&next_value(&mut values, "--mode")?)?,
            "--counts" => counts = parse_counts(&next_value(&mut values, "--counts")?)?,
            "--samples" => {
                samples = parse_positive(&next_value(&mut values, "--samples")?, "--samples")?
            },
            "--warmups" => {
                warmups = parse_positive(&next_value(&mut values, "--warmups")?, "--warmups")?
            },
            "--repeats" => {
                repeats = parse_positive(&next_value(&mut values, "--repeats")?, "--repeats")?
            },
            "--spool-dir" => {
                spool_dir = Some(PathBuf::from(next_value(&mut values, "--spool-dir")?))
            },
            "--max-spool-bytes" => {
                scratch_limits.max_bytes = parse_positive_u64(
                    &next_value(&mut values, "--max-spool-bytes")?,
                    "--max-spool-bytes",
                )?
            },
            "--spool-buffer-bytes" => {
                scratch_limits.buffer_bytes = parse_positive(
                    &next_value(&mut values, "--spool-buffer-bytes")?,
                    "--spool-buffer-bytes",
                )?
            },
            "--json" => json_path = Some(PathBuf::from(next_value(&mut values, "--json")?)),
            "--help" | "-h" => return Err(usage().into()),
            unknown => return Err(format!("unknown argument {unknown}; {}", usage()).into()),
        }
    }
    if samples == 0 || warmups == 0 || repeats == 0 || counts.is_empty() {
        return Err("samples, warmups, repeats, and counts must be non-zero".into());
    }
    if modes.is_empty() {
        return Err("at least one mode is required".into());
    }
    if modes.contains(&Mode::Spool) && spool_dir.is_none() {
        return Err("--spool-dir is required for spool or both mode".into());
    }
    Ok(Config {
        modes,
        counts,
        samples,
        warmups,
        repeats,
        spool_dir,
        scratch_limits,
        json_path,
    })
}

fn next_value(values: &mut impl Iterator<Item = OsString>, flag: &str) -> BenchResult<OsString> {
    values
        .next()
        .ok_or_else(|| format!("missing value for {flag}").into())
}

fn parse_modes(value: &OsString) -> BenchResult<Vec<Mode>> {
    let text = value.to_str().ok_or("--mode must be UTF-8")?;
    let mut modes = Vec::new();
    for part in text.split(',') {
        if part == "both" {
            for mode in [Mode::Control, Mode::Spool] {
                if !modes.contains(&mode) {
                    modes.push(mode);
                }
            }
            continue;
        }
        let mode = Mode::parse(part)?;
        if !modes.contains(&mode) {
            modes.push(mode);
        }
    }
    Ok(modes)
}

fn parse_counts(value: &OsString) -> BenchResult<Vec<usize>> {
    let text = value.to_str().ok_or("--counts must be UTF-8")?;
    let mut counts = Vec::new();
    for part in text.split(',') {
        let count = part
            .parse::<usize>()
            .map_err(|error| format!("invalid count {part:?}: {error}"))?;
        if !DEFAULT_COUNTS.contains(&count) {
            return Err(
                format!("unsupported PPTX slide count {count}; expected 8, 256, or 8192").into(),
            );
        }
        if !counts.contains(&count) {
            counts.push(count);
        }
    }
    if counts.is_empty() {
        return Err("--counts cannot be empty".into());
    }
    Ok(counts)
}

fn parse_positive(value: &OsString, flag: &str) -> BenchResult<usize> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let parsed = text
        .parse::<usize>()
        .map_err(|error| format!("invalid {flag} {text:?}: {error}"))?;
    if parsed == 0 {
        return Err(format!("{flag} must be positive").into());
    }
    Ok(parsed)
}

fn parse_positive_u64(value: &OsString, flag: &str) -> BenchResult<u64> {
    let text = value
        .to_str()
        .ok_or_else(|| format!("{flag} must be UTF-8"))?;
    let parsed = text
        .parse::<u64>()
        .map_err(|error| format!("invalid {flag} {text:?}: {error}"))?;
    if parsed == 0 {
        return Err(format!("{flag} must be positive").into());
    }
    Ok(parsed)
}

fn hex_digest(digest: &[u8]) -> String {
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn usage() -> &'static str {
    "usage: pptx_metadata_spool [--mode control|spool|both] [--counts 8,256,8192] [--samples N] [--warmups N] [--repeats N] [--spool-dir PATH] [--max-spool-bytes N] [--spool-buffer-bytes N] [--json PATH]"
}

#[cfg(test)]
mod tests {
    use super::*;
    use soapberry_zip::office::StreamingArchiveWriter;

    struct TestDirectory(PathBuf);

    impl TestDirectory {
        fn new() -> Self {
            use std::sync::atomic::{AtomicU64, Ordering};
            static NEXT: AtomicU64 = AtomicU64::new(0);
            loop {
                let path = std::env::temp_dir().join(format!(
                    "litchi-pptx-spool-test-{}-{}",
                    std::process::id(),
                    NEXT.fetch_add(1, Ordering::Relaxed)
                ));
                match fs::create_dir(&path) {
                    Ok(()) => return Self(path),
                    Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
                    Err(error) => panic!("create test directory: {error}"),
                }
            }
        }

        fn args(&self) -> Vec<OsString> {
            [
                "--mode",
                "spool",
                "--counts",
                "8",
                "--samples",
                "1",
                "--warmups",
                "1",
                "--repeats",
                "1",
                "--spool-dir",
            ]
            .into_iter()
            .map(OsString::from)
            .chain([
                self.0.clone().into_os_string(),
                OsString::from("--json"),
                self.0.join("report.json").into_os_string(),
            ])
            .collect()
        }
    }

    impl Drop for TestDirectory {
        fn drop(&mut self) {
            let _ = fs::remove_dir_all(&self.0);
        }
    }

    #[test]
    fn cli_preserves_preexisting_oracle_warmup_and_sample_files() {
        for name in [
            "pptx-metadata-spool-8-oracle-spool.tmp",
            "pptx-metadata-spool-8-spool-0-0-warmup.tmp",
            "pptx-metadata-spool-8-spool-0-0-sample.tmp",
        ] {
            let directory = TestDirectory::new();
            let path = directory.0.join(name);
            fs::write(&path, b"caller-owned sentinel").unwrap();
            let error = run_from_args(directory.args()).expect_err("create_new must refuse");
            assert_eq!(
                error.downcast_ref::<io::Error>().unwrap().kind(),
                io::ErrorKind::AlreadyExists
            );
            assert_eq!(fs::read(path).unwrap(), b"caller-owned sentinel");
            assert!(!directory.0.join("report.json").exists());
        }
    }

    #[test]
    fn cli_removes_scratch_after_success_and_retains_it_after_failure() {
        let directory = TestDirectory::new();
        run_from_args(directory.args()).expect("successful spool run");
        let names = fs::read_dir(&directory.0)
            .unwrap()
            .map(|entry| entry.unwrap().file_name())
            .collect::<Vec<_>>();
        assert_eq!(names, vec![OsString::from("report.json")]);

        let failed = TestDirectory::new();
        let mut args = failed.args();
        args.extend([OsString::from("--max-spool-bytes"), OsString::from("1")]);
        assert!(run_from_args(args).is_err());
        assert!(
            failed
                .0
                .join("pptx-metadata-spool-8-oracle-spool.tmp")
                .exists()
        );
        assert!(!failed.0.join("report.json").exists());
    }

    #[test]
    fn default_cli_is_control_only_with_existing_shapes() {
        let config = parse_args(Vec::<OsString>::new()).expect("default config");
        assert_eq!(config.modes, vec![Mode::Control]);
        assert_eq!(config.counts, DEFAULT_COUNTS);
        assert_eq!(config.samples, DEFAULT_SAMPLES);
        assert_eq!(config.warmups, DEFAULT_WARMUPS);
        assert_eq!(config.repeats, DEFAULT_REPEATS);
    }

    #[test]
    fn both_mode_deduplicates_and_requires_spool_directory() {
        let error = parse_args([OsString::from("--mode"), OsString::from("both,control")])
            .expect_err("paired mode must require an explicit directory");
        assert!(error.to_string().contains("--spool-dir"));

        let config = parse_args([
            OsString::from("--mode"),
            OsString::from("both,control"),
            OsString::from("--spool-dir"),
            OsString::from("/tmp/pptx-metadata-spool-test"),
            OsString::from("--counts"),
            OsString::from("8,8,256"),
        ])
        .expect("paired config");
        assert_eq!(config.modes, vec![Mode::Control, Mode::Spool]);
        assert_eq!(config.counts, vec![8, 256]);
    }

    #[test]
    fn physical_entry_count_is_fixed_members_plus_two_per_slide() {
        assert_eq!(expected_entry_count(8).unwrap(), 37 + 16);
        assert_eq!(expected_entry_count(256).unwrap(), 37 + 512);
        assert_eq!(expected_entry_count(8_192).unwrap(), 37 + 16_384);
    }

    #[test]
    fn hashing_sink_is_fixed_and_deterministic() {
        let mut sink = HashingSink::new();
        sink.write_all(b"abc").unwrap();
        sink.write_all(b"def").unwrap();
        let summary = sink.finish().unwrap();
        assert_eq!(summary.accepted_bytes, 6);
        assert_eq!(summary.write_calls, 2);
        assert_eq!(
            summary.sha256,
            "bef57ec7f53a6d40beb640a780a639c83bc29ac8a9816f1fc6c5c6dcd93c4721"
        );
    }

    #[test]
    fn physical_oracle_rejects_corrupted_static_payload() {
        let corpus = pptx_streaming_create::build_corpus_for_slide_count(8)
            .expect("deterministic PPTX corpus");
        let (archive, _) = create_materialized_control(
            &corpus,
            ScratchLimits {
                max_bytes: DEFAULT_SPOOL_MAX_BYTES,
                buffer_bytes: DEFAULT_SPOOL_BUFFER_BYTES,
            },
        )
        .expect("control archive");
        let source = PhysPkgReader::new(&archive).expect("source package");
        let static_payload = source.read_member("ppt/presProps.xml").unwrap();
        let mut rewritten = StreamingArchiveWriter::new();
        for name in source.member_names().unwrap() {
            let payload = source.read_member(&name).unwrap();
            rewritten.write_stored(&name, &payload).unwrap();
        }
        let mut corrupted = rewritten.finish_to_bytes().unwrap();
        let count = expected_entry_count(8).unwrap();
        reopen_every_physical_member(&corrupted, count).unwrap();
        let offset = corrupted
            .windows(static_payload.len())
            .position(|bytes| bytes == static_payload)
            .expect("stored static payload");
        corrupted[offset + static_payload.len() / 2] ^= 1;
        assert!(reopen_every_physical_member(&corrupted, count).is_err());
    }
}
