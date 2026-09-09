//! A bounded XML-audit comparison harness.
//!
//! The materialized leg first consumes the deterministic [`XmlGenerator`] into
//! a `Vec<u8>` and then calls the existing slice auditor.  The streaming leg
//! gives the same generator directly to the bounded reader auditor.  Corpus
//! construction and the independent source digest oracle happen before each
//! timed region; generator construction, source consumption, audit, and all
//! owned-source drops happen inside it.

use std::{
    error::Error,
    ffi::OsString,
    fmt::Write as _,
    fs,
    io::{self, BufRead, Read, Write},
    path::PathBuf,
    time::Instant,
};

use serde::Serialize;
use sha2::{Digest as _, Sha256};
use xml_minifier::audit::{self, Error as AuditError, Limits, Report, StreamError};

use crate::allocation_metrics;

const SCHEMA: &str = "litchi.xml-stream-audit.v1";
const GENERATOR: &str = "litchi-xml-repetitive-root-items-v1";
const DEFAULT_SIZES: [usize; 3] = [64 * 1024, 8 * 1024 * 1024, 128 * 1024 * 1024];
const DEFAULT_SAMPLES: usize = 30;
const DEFAULT_WARMUPS: usize = 3;
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUPS: usize = 1_000;
const GENERATOR_CHUNK_BYTES: usize = 16 * 1024;
const RECORD_BYTES: usize = 1_024;
const PREFIX: &[u8] = b"<root>";
const SUFFIX: &[u8] = b"</root>";
const ITEM_OPEN: &[u8] = b"<item>";
const ITEM_CLOSE: &[u8] = b"</item>";
const MIN_RECORD_BYTES: usize = ITEM_OPEN.len() + ITEM_CLOSE.len();

type BenchResult<T> = Result<T, Box<dyn Error>>;

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum Mode {
    Materialized,
    Streaming,
}

impl Mode {
    const fn name(self) -> &'static str {
        match self {
            Self::Materialized => "materialized",
            Self::Streaming => "streaming",
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    modes: Vec<Mode>,
    sizes: Vec<usize>,
    samples: usize,
    warmups: usize,
    json: Option<PathBuf>,
}

#[derive(Clone, Debug, Serialize)]
struct BinaryRecord {
    binary: &'static str,
    allocator: &'static str,
    instrumentation: &'static str,
    counter_revision: Option<&'static str>,
}

#[derive(Clone, Debug, Serialize)]
struct LimitRecord {
    max_bytes: usize,
    max_depth: usize,
    max_events: usize,
    max_attributes: usize,
    max_token_bytes: usize,
    max_text_bytes: usize,
    streaming_memory_upper_bound: Option<usize>,
}

#[derive(Clone, Debug, Serialize)]
struct LayoutRecord {
    prefix_bytes: usize,
    suffix_bytes: usize,
    full_record_bytes: usize,
    full_record_count: usize,
    tail_record_bytes: usize,
    record_count: usize,
    text_bytes: usize,
    item_open_bytes: usize,
    item_close_bytes: usize,
}

#[derive(Clone, Debug, Serialize)]
struct SourceRecord {
    requested_bytes: usize,
    generated_bytes: usize,
    sha256: String,
    actual_generator_bytes: usize,
    actual_generator_sha256: String,
    actual_generator_oracle_verified: bool,
    generator: &'static str,
    chunk_bytes: usize,
    source_materialization: &'static str,
    layout: LayoutRecord,
}

#[derive(Clone, Copy, Debug)]
struct Layout {
    total_bytes: usize,
    full_record_count: usize,
    tail_record_bytes: usize,
    record_count: usize,
}

#[derive(Clone, Debug)]
struct ActualGeneratorRecord {
    bytes: usize,
    sha256: String,
}

impl Layout {
    fn for_size(total_bytes: usize) -> BenchResult<Self> {
        let framing = PREFIX
            .len()
            .checked_add(SUFFIX.len())
            .ok_or("XML framing length overflow")?;
        let body = total_bytes
            .checked_sub(framing)
            .ok_or("XML size is smaller than root framing")?;
        if body < MIN_RECORD_BYTES {
            return Err(format!("XML size {total_bytes} is too small for one item record").into());
        }

        let mut full_record_count = body / RECORD_BYTES;
        let remainder = body % RECORD_BYTES;
        let tail_record_bytes = if remainder == 0 {
            0
        } else if remainder >= MIN_RECORD_BYTES {
            remainder
        } else {
            full_record_count = full_record_count
                .checked_sub(1)
                .ok_or("XML size cannot borrow a full record for its tail")?;
            RECORD_BYTES
                .checked_add(remainder)
                .ok_or("XML tail record length overflow")?
        };
        let record_count = full_record_count
            .checked_add(usize::from(tail_record_bytes != 0))
            .ok_or("XML record count overflow")?;
        if record_count == 0 {
            return Err("XML generator produced no item records".into());
        }
        Ok(Self {
            total_bytes,
            full_record_count,
            tail_record_bytes,
            record_count,
        })
    }

    fn full_record_bytes(self) -> usize {
        self.full_record_count * RECORD_BYTES
    }

    fn record_len(self, record: usize) -> usize {
        if record < self.full_record_count {
            RECORD_BYTES
        } else {
            self.tail_record_bytes
        }
    }

    fn text_bytes(self) -> usize {
        self.total_bytes - PREFIX.len() - SUFFIX.len() - self.record_count * MIN_RECORD_BYTES
    }

    fn record_count(self) -> usize {
        self.record_count
    }

    fn text_event_count(self) -> usize {
        (0..self.record_count())
            .filter(|&record| self.record_len(record) > MIN_RECORD_BYTES)
            .count()
    }

    fn record_byte(self, record: usize, offset: usize) -> u8 {
        let length = self.record_len(record);
        debug_assert!(offset < length);
        if offset < ITEM_OPEN.len() {
            ITEM_OPEN[offset]
        } else if offset < length - ITEM_CLOSE.len() {
            b'a' + (record as u8 % 26)
        } else {
            ITEM_CLOSE[offset - (length - ITEM_CLOSE.len())]
        }
    }

    fn byte_at(self, position: usize) -> u8 {
        debug_assert!(position < self.total_bytes);
        if position < PREFIX.len() {
            return PREFIX[position];
        }
        let body_position = position - PREFIX.len();
        let full_bytes = self.full_record_bytes();
        if body_position < full_bytes {
            let record = body_position / RECORD_BYTES;
            return self.record_byte(record, body_position % RECORD_BYTES);
        }
        if self.tail_record_bytes != 0 {
            let tail_position = body_position - full_bytes;
            if tail_position < self.tail_record_bytes {
                return self.record_byte(self.full_record_count, tail_position);
            }
        }
        let suffix_position = position - (self.total_bytes - SUFFIX.len());
        SUFFIX[suffix_position]
    }

    fn record_layout(self) -> LayoutRecord {
        LayoutRecord {
            prefix_bytes: PREFIX.len(),
            suffix_bytes: SUFFIX.len(),
            full_record_bytes: RECORD_BYTES,
            full_record_count: self.full_record_count,
            tail_record_bytes: self.tail_record_bytes,
            record_count: self.record_count,
            text_bytes: self.text_bytes(),
            item_open_bytes: ITEM_OPEN.len(),
            item_close_bytes: ITEM_CLOSE.len(),
        }
    }
}

/// Deterministic XML source with a fixed-size read buffer and no payload-sized
/// allocation.  The generated bytes are all ASCII and contain one root with
/// repeated `<item>...</item>` children.
#[derive(Debug)]
struct XmlGenerator {
    layout: Layout,
    position: usize,
    buffer: Vec<u8>,
    offset: usize,
}

impl XmlGenerator {
    fn new(total_bytes: usize) -> BenchResult<Self> {
        let layout = Layout::for_size(total_bytes)?;
        let mut buffer = Vec::new();
        buffer
            .try_reserve_exact(GENERATOR_CHUNK_BYTES)
            .map_err(|_| "could not reserve the fixed XML generator buffer")?;
        buffer.resize(GENERATOR_CHUNK_BYTES, 0);
        Ok(Self {
            layout,
            position: 0,
            buffer,
            offset: GENERATOR_CHUNK_BYTES,
        })
    }

    fn position(&self) -> usize {
        self.position
    }
}

impl Read for XmlGenerator {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl BufRead for XmlGenerator {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.offset < self.buffer.len() {
            return Ok(&self.buffer[self.offset..]);
        }
        if self.position >= self.layout.total_bytes {
            self.buffer.clear();
            self.offset = 0;
            return Ok(&[]);
        }
        let count = (self.layout.total_bytes - self.position).min(GENERATOR_CHUNK_BYTES);
        self.buffer.clear();
        self.buffer.resize(count, 0);
        for (index, byte) in self.buffer.iter_mut().enumerate() {
            *byte = self.layout.byte_at(self.position + index);
        }
        self.offset = 0;
        Ok(&self.buffer)
    }

    fn consume(&mut self, amount: usize) {
        let available = self.buffer.len().saturating_sub(self.offset);
        let count = amount.min(available);
        self.offset += count;
        self.position += count;
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize)]
struct ReportRecord {
    attributes: usize,
    bytes: usize,
    events: usize,
    max_depth: usize,
    text_bytes: usize,
}

impl From<Report> for ReportRecord {
    fn from(report: Report) -> Self {
        Self {
            attributes: report.attributes(),
            bytes: report.bytes(),
            events: report.events(),
            max_depth: report.max_depth(),
            text_bytes: report.text_bytes(),
        }
    }
}

#[derive(Clone, Debug, Serialize)]
struct AuditFailure {
    kind: &'static str,
    display: String,
    debug: String,
}

impl AuditFailure {
    fn from_audit(error: AuditError) -> Self {
        let kind = match &error {
            AuditError::Limit { .. } => "limit",
            AuditError::Encoding { .. } => "encoding",
            AuditError::Malformed { .. } => "malformed",
            AuditError::NotCompact(_) => "not_compact",
            AuditError::Doctype { .. } => "doctype",
            AuditError::Allocation => "allocation",
            _ => "unknown",
        };
        Self {
            kind,
            display: error.to_string(),
            debug: format!("{error:?}"),
        }
    }

    fn from_io(error: io::Error) -> Self {
        Self {
            kind: "io",
            display: error.to_string(),
            debug: format!("{error:?}"),
        }
    }

    fn from_stream(error: StreamError) -> Self {
        match error {
            StreamError::Input(source) => Self::from_io(source),
            StreamError::Audit(source) => Self::from_audit(source),
            other => Self {
                kind: "unknown_stream",
                display: other.to_string(),
                debug: format!("{other:?}"),
            },
        }
    }
}

#[derive(Clone, Debug, Serialize)]
#[serde(tag = "status", rename_all = "snake_case")]
enum AuditOutcome {
    Success { report: ReportRecord },
    Failure { error: AuditFailure },
}

#[derive(Debug)]
enum RawOutcome {
    Success(Report),
    Failure(AuditFailure),
}

#[derive(Debug)]
struct RawOperation {
    outcome: RawOutcome,
    actual_bytes: usize,
    source_materialization_bytes: Option<usize>,
    generator_position: usize,
}

#[derive(Clone, Debug, Serialize)]
struct OperationRecord {
    source_materialization_bytes: Option<usize>,
    generator_buffer_bytes: usize,
    generator_position: usize,
    entry_live_bytes: Option<u64>,
    exit_live_bytes: Option<u64>,
    net_live_bytes: Option<i128>,
    zero_net_live: Option<bool>,
}

#[derive(Debug, Serialize)]
struct SampleRecord {
    sample: usize,
    elapsed_ns: u64,
    actual_bytes: usize,
    source_bytes_match: bool,
    outcome: AuditOutcome,
    allocation: Option<allocation_metrics::Sample>,
    operation: OperationRecord,
}

#[derive(Debug, Serialize)]
struct CaseRecord {
    mode: Mode,
    size_bytes: usize,
    source: SourceRecord,
    limits: LimitRecord,
    warmups: usize,
    samples: Vec<SampleRecord>,
}

#[derive(Debug, Serialize)]
struct BenchmarkReport {
    schema: &'static str,
    benchmark: &'static str,
    binary: BinaryRecord,
    modes: Vec<Mode>,
    sizes: Vec<usize>,
    samples: usize,
    warmups: usize,
    timed_scope: &'static str,
    corpus_scope: &'static str,
    cases: Vec<CaseRecord>,
}

fn limits_for(size: usize) -> BenchResult<Limits> {
    Limits::new(size, 8, Limits::EVENT_CEILING, 16, 16 * 1024, size)
        .map_err(|error| format!("cannot configure XML audit limits: {error}").into())
}

fn limit_record(limits: Limits) -> LimitRecord {
    LimitRecord {
        max_bytes: limits.max_bytes(),
        max_depth: limits.max_depth(),
        max_events: limits.max_events(),
        max_attributes: limits.max_attributes(),
        max_token_bytes: limits.max_token_bytes(),
        max_text_bytes: limits.max_text_bytes(),
        streaming_memory_upper_bound: limits.streaming_memory_upper_bound(),
    }
}

fn source_identity(size: usize, layout: Layout) -> SourceRecord {
    let mut digest = Sha256::new();
    digest.update(PREFIX);
    let mut text = [0u8; RECORD_BYTES];
    for record in 0..layout.record_count() {
        let record_length = layout.record_len(record);
        let text_length = record_length - MIN_RECORD_BYTES;
        let byte = b'a' + (record as u8 % 26);
        text[..text_length].fill(byte);
        digest.update(ITEM_OPEN);
        digest.update(&text[..text_length]);
        digest.update(ITEM_CLOSE);
    }
    digest.update(SUFFIX);
    SourceRecord {
        requested_bytes: size,
        generated_bytes: layout.total_bytes,
        sha256: hex_digest(&digest.finalize()),
        actual_generator_bytes: 0,
        actual_generator_sha256: String::new(),
        actual_generator_oracle_verified: false,
        generator: GENERATOR,
        chunk_bytes: GENERATOR_CHUNK_BYTES,
        source_materialization: "materialized_leg_only",
        layout: layout.record_layout(),
    }
}

fn actual_generator_record(size: usize) -> BenchResult<ActualGeneratorRecord> {
    let mut source = XmlGenerator::new(size)?;
    let mut digest = Sha256::new();
    let mut scratch = [0u8; GENERATOR_CHUNK_BYTES];
    let mut bytes = 0usize;
    loop {
        let count = source.read(&mut scratch)?;
        if count == 0 {
            break;
        }
        bytes = bytes
            .checked_add(count)
            .ok_or("generated XML byte count overflow")?;
        digest.update(&scratch[..count]);
    }
    Ok(ActualGeneratorRecord {
        bytes,
        sha256: hex_digest(&digest.finalize()),
    })
}

fn run_materialized(size: usize, limits: Limits) -> BenchResult<RawOperation> {
    let mut source = XmlGenerator::new(size)?;
    let mut materialized = Vec::new();
    let result = source
        .read_to_end(&mut materialized)
        .map_err(AuditFailure::from_io)
        .and_then(|_| {
            audit::verify_authored(&materialized, limits).map_err(AuditFailure::from_audit)
        });
    let actual_bytes = materialized.len();
    let generator_position = source.position();
    drop(source);
    drop(materialized);
    Ok(RawOperation {
        outcome: match result {
            Ok(report) => RawOutcome::Success(report),
            Err(error) => RawOutcome::Failure(error),
        },
        actual_bytes,
        source_materialization_bytes: Some(actual_bytes),
        generator_position,
    })
}

fn run_streaming(size: usize, limits: Limits) -> BenchResult<RawOperation> {
    let mut source = XmlGenerator::new(size)?;
    let result =
        audit::verify_authored_reader(&mut source, limits).map_err(AuditFailure::from_stream);
    let generator_position = source.position();
    let actual_bytes = generator_position;
    drop(source);
    Ok(RawOperation {
        outcome: match result {
            Ok(report) => RawOutcome::Success(report),
            Err(error) => RawOutcome::Failure(error),
        },
        actual_bytes,
        source_materialization_bytes: None,
        generator_position,
    })
}

fn operation_record(
    raw: &RawOperation,
    allocation: &Option<allocation_metrics::Sample>,
) -> OperationRecord {
    let (entry_live_bytes, exit_live_bytes, net_live_bytes, zero_net_live) = allocation
        .as_ref()
        .map_or((None, None, None, None), |sample| {
            let entry = sample.live_bytes_before;
            let exit = sample.live_bytes_after;
            let net = entry
                .zip(exit)
                .map(|(before, after)| i128::from(after) - i128::from(before));
            (entry, exit, net, net.map(|value| value == 0))
        });
    OperationRecord {
        source_materialization_bytes: raw.source_materialization_bytes,
        generator_buffer_bytes: GENERATOR_CHUNK_BYTES,
        generator_position: raw.generator_position,
        entry_live_bytes,
        exit_live_bytes,
        net_live_bytes,
        zero_net_live,
    }
}

fn outcome_record(outcome: RawOutcome) -> AuditOutcome {
    match outcome {
        RawOutcome::Success(report) => AuditOutcome::Success {
            report: report.into(),
        },
        RawOutcome::Failure(error) => AuditOutcome::Failure { error },
    }
}

fn measure_one(
    mode: Mode,
    size: usize,
    limits: Limits,
    sample: usize,
) -> BenchResult<SampleRecord> {
    let allocation_region = allocation_metrics::begin();
    let start = Instant::now();
    let raw = match mode {
        Mode::Materialized => run_materialized(size, limits)?,
        Mode::Streaming => run_streaming(size, limits)?,
    };
    let elapsed_ns = u64::try_from(start.elapsed().as_nanos())
        .map_err(|_| "elapsed duration does not fit in u64 nanoseconds")?;
    let allocation = allocation_region.finish();
    let actual_bytes = raw.actual_bytes;
    let source_bytes_match = actual_bytes == size && raw.generator_position == size;
    let operation = operation_record(&raw, &allocation);
    let outcome = outcome_record(raw.outcome);
    Ok(SampleRecord {
        sample,
        elapsed_ns,
        actual_bytes,
        source_bytes_match,
        outcome,
        allocation,
        operation,
    })
}

fn run_case(mode: Mode, size: usize, samples: usize, warmups: usize) -> BenchResult<CaseRecord> {
    let layout = Layout::for_size(size)?;
    let mut source = source_identity(size, layout);
    let actual = actual_generator_record(size)?;
    if actual.bytes != source.generated_bytes || actual.sha256 != source.sha256 {
        return Err(format!(
            "deterministic XML generator oracle mismatch for {}: expected {} bytes/hash {}, observed {} bytes/hash {}",
            size,
            source.generated_bytes,
            source.sha256,
            actual.bytes,
            actual.sha256,
        )
        .into());
    }
    source.actual_generator_bytes = actual.bytes;
    source.actual_generator_sha256 = actual.sha256;
    source.actual_generator_oracle_verified = true;
    let limits = limits_for(size)?;
    for _ in 0..warmups {
        let warmup = measure_one(mode, size, limits, 0)?;
        validate_sample(mode, size, layout, &warmup)?;
    }
    let mut timed = Vec::new();
    timed
        .try_reserve_exact(samples)
        .map_err(|_| "could not reserve timed XML audit sample records")?;
    for sample in 0..samples {
        let record = measure_one(mode, size, limits, sample)?;
        validate_sample(mode, size, layout, &record)?;
        timed.push(record);
    }
    Ok(CaseRecord {
        mode,
        size_bytes: size,
        source,
        limits: limit_record(limits),
        warmups,
        samples: timed,
    })
}

fn expected_report(size: usize, layout: Layout) -> ReportRecord {
    ReportRecord {
        attributes: 0,
        bytes: size,
        events: layout.record_count() * 2 + layout.text_event_count() + 3,
        max_depth: 2,
        text_bytes: layout.text_bytes(),
    }
}

fn validate_sample(
    mode: Mode,
    size: usize,
    layout: Layout,
    sample: &SampleRecord,
) -> BenchResult<()> {
    if !sample.source_bytes_match {
        return Err(format!(
            "{} sample {} generated {} bytes for requested {}",
            mode.name(),
            sample.sample,
            sample.actual_bytes,
            size
        )
        .into());
    }
    match &sample.outcome {
        AuditOutcome::Success { report } => {
            let expected = expected_report(size, layout);
            if *report != expected {
                return Err(format!(
                    "{} sample {} returned report {:?}, expected {:?}",
                    mode.name(),
                    sample.sample,
                    report,
                    expected
                )
                .into());
            }
            Ok(())
        },
        AuditOutcome::Failure { error } => Err(format!(
            "{} sample {} failed XML audit with typed {} error: {} ({})",
            mode.name(),
            sample.sample,
            error.kind,
            error.display,
            error.debug
        )
        .into()),
    }
}

fn parse_usize(value: &str, flag: &str, maximum: usize) -> BenchResult<usize> {
    let parsed = value
        .parse::<usize>()
        .map_err(|error| format!("invalid {flag} value {value:?}: {error}"))?;
    if parsed > maximum {
        return Err(format!("{flag} value {parsed} exceeds maximum {maximum}").into());
    }
    Ok(parsed)
}

fn parse_sizes(value: &str) -> BenchResult<Vec<usize>> {
    let mut sizes = Vec::new();
    for part in value.split(',') {
        let size = part
            .parse::<usize>()
            .map_err(|error| format!("invalid --sizes value {part:?}: {error}"))?;
        if size > Limits::BYTE_CEILING {
            return Err(format!(
                "--sizes value {size} exceeds XML audit hard byte ceiling {}",
                Limits::BYTE_CEILING
            )
            .into());
        }
        Layout::for_size(size)?;
        sizes.push(size);
    }
    if sizes.is_empty() {
        return Err("--sizes must contain at least one size".into());
    }
    Ok(sizes)
}

fn parse_mode(value: &str) -> BenchResult<Vec<Mode>> {
    match value {
        "materialized" => Ok(vec![Mode::Materialized]),
        "streaming" => Ok(vec![Mode::Streaming]),
        "both" => Ok(vec![Mode::Materialized, Mode::Streaming]),
        _ => Err(
            format!("invalid --mode {value:?}; expected materialized, streaming, or both").into(),
        ),
    }
}

fn usage() -> &'static str {
    "usage: xml_stream_audit [--mode materialized|streaming|both] [--sizes BYTES,...] [--samples N] [--warmup N] [--json PATH]"
}

fn parse_args<I>(args: I) -> BenchResult<Option<Config>>
where
    I: IntoIterator<Item = OsString>,
{
    let mut modes = vec![Mode::Materialized, Mode::Streaming];
    let mut sizes = DEFAULT_SIZES.to_vec();
    let mut samples = DEFAULT_SAMPLES;
    let mut warmups = DEFAULT_WARMUPS;
    let mut json = None;
    let mut args = args.into_iter();
    while let Some(argument) = args.next() {
        let argument = argument
            .into_string()
            .map_err(|_| "command-line argument is not valid UTF-8")?;
        if argument == "--help" || argument == "-h" {
            println!("{}", usage());
            return Ok(None);
        }
        let (flag, value) = argument.split_once('=').map_or_else(
            || (argument.as_str(), None),
            |(flag, value)| (flag, Some(value)),
        );
        let value = match value {
            Some(value) => value.to_owned(),
            None => args
                .next()
                .ok_or_else(|| format!("{flag} requires a value; {}", usage()))?
                .into_string()
                .map_err(|_| "command-line value is not valid UTF-8")?,
        };
        match flag {
            "--mode" => modes = parse_mode(&value)?,
            "--sizes" => sizes = parse_sizes(&value)?,
            "--samples" => {
                samples = parse_usize(&value, "--samples", MAX_SAMPLES)?;
                if samples == 0 {
                    return Err("--samples must be at least one".into());
                }
            },
            "--warmup" | "--warmups" => warmups = parse_usize(&value, "--warmup", MAX_WARMUPS)?,
            "--json" => json = Some(PathBuf::from(value)),
            _ => return Err(format!("unknown option {flag}; {}", usage()).into()),
        }
    }
    Ok(Some(Config {
        modes,
        sizes,
        samples,
        warmups,
        json,
    }))
}

fn hex_digest(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        let _ = write!(&mut output, "{byte:02x}");
    }
    output
}

/// Runs the benchmark and writes its JSON report to stdout or `--json`.
pub fn run_from_args<I>(args: I) -> BenchResult<()>
where
    I: IntoIterator<Item = OsString>,
{
    let Some(config) = parse_args(args)? else {
        return Ok(());
    };
    let mut cases = Vec::new();
    for mode in &config.modes {
        for &size in &config.sizes {
            cases.push(run_case(*mode, size, config.samples, config.warmups)?);
        }
    }
    let report = BenchmarkReport {
        schema: SCHEMA,
        benchmark: GENERATOR,
        binary: BinaryRecord {
            binary: allocation_metrics::binary_identity(),
            allocator: allocation_metrics::allocator_identity(),
            instrumentation: allocation_metrics::instrumentation_identity(),
            counter_revision: allocation_metrics::counter_revision(),
        },
        modes: config.modes,
        sizes: config.sizes,
        samples: config.samples,
        warmups: config.warmups,
        timed_scope: "generator_create_read_materialize_or_stream_audit_and_drop",
        corpus_scope: "deterministic_generator_hash_and_layout_are_precomputed_outside_timed_regions",
        cases,
    };
    let output = serde_json::to_vec_pretty(&report)?;
    if let Some(path) = config.json {
        fs::write(path, &output)?;
    } else {
        let mut stdout = io::BufWriter::new(io::stdout().lock());
        stdout.write_all(&output)?;
        stdout.write_all(b"\n")?;
        stdout.flush()?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limits(size: usize) -> Limits {
        limits_for(size).unwrap()
    }

    #[test]
    fn generator_is_exact_and_auditable_at_small_size() {
        let size = 64 * 1024;
        let layout = Layout::for_size(size).unwrap();
        let mut source = XmlGenerator::new(size).unwrap();
        let mut bytes = Vec::new();
        source.read_to_end(&mut bytes).unwrap();
        assert_eq!(bytes.len(), size);
        assert_eq!(source.position(), size);
        let identity = source_identity(size, layout);
        assert_eq!(identity.sha256, hex_digest(Sha256::digest(&bytes).as_ref()));
        let actual = actual_generator_record(size).unwrap();
        assert_eq!(actual.bytes, identity.generated_bytes);
        assert_eq!(actual.sha256, identity.sha256);
        let report = audit::verify_authored(&bytes, limits(size)).unwrap();
        assert_eq!(report.bytes(), size);
        assert_eq!(report.max_depth(), 2);
        assert_eq!(report.text_bytes(), layout.text_bytes());
    }

    #[test]
    fn materialized_and_streaming_reports_match() {
        let size = 64 * 1024;
        let limits = limits(size);
        let materialized = run_materialized(size, limits).unwrap();
        let streaming = run_streaming(size, limits).unwrap();
        match (materialized.outcome, streaming.outcome) {
            (RawOutcome::Success(left), RawOutcome::Success(right)) => {
                assert_eq!(left, right);
            },
            _ => panic!("deterministic valid XML must audit successfully"),
        }
        assert_eq!(materialized.actual_bytes, streaming.actual_bytes);
        assert_eq!(
            materialized.generator_position,
            streaming.generator_position
        );
    }

    #[test]
    fn generator_tail_borrows_a_record_when_remainder_is_short() {
        let size = PREFIX.len() + SUFFIX.len() + RECORD_BYTES + 1;
        let layout = Layout::for_size(size).unwrap();
        assert_eq!(layout.tail_record_bytes, RECORD_BYTES + 1);
        let mut source = XmlGenerator::new(size).unwrap();
        let mut bytes = Vec::new();
        source.read_to_end(&mut bytes).unwrap();
        assert_eq!(bytes.len(), size);
        let report = audit::verify_authored(&bytes, limits(size)).unwrap();
        assert_eq!(report.bytes(), size);
    }

    fn reference_xml(layout: Layout) -> Vec<u8> {
        let mut output = Vec::with_capacity(layout.total_bytes);
        output.extend_from_slice(PREFIX);
        for record in 0..layout.record_count() {
            let length = layout.record_len(record);
            let text_length = length - MIN_RECORD_BYTES;
            output.extend_from_slice(ITEM_OPEN);
            output.extend(std::iter::repeat_n(b'a' + (record as u8 % 26), text_length));
            output.extend_from_slice(ITEM_CLOSE);
        }
        output.extend_from_slice(SUFFIX);
        output
    }

    #[test]
    fn generator_short_remainders_match_independent_reference_and_reader() {
        let framing = PREFIX.len() + SUFFIX.len();
        let bodies = [
            MIN_RECORD_BYTES,
            RECORD_BYTES - 1,
            RECORD_BYTES,
            RECORD_BYTES + 1,
            RECORD_BYTES + 12,
            RECORD_BYTES + 13,
            RECORD_BYTES * 2 - 1,
            RECORD_BYTES * 2,
            RECORD_BYTES * 2 + 1,
        ];
        for body in bodies {
            let size = framing + body;
            let layout = Layout::for_size(size).unwrap();
            let expected = reference_xml(layout);
            let mut source = XmlGenerator::new(size).unwrap();
            let mut actual = Vec::new();
            source.read_to_end(&mut actual).unwrap();
            assert_eq!(actual, expected, "size {size}");
            let report = audit::verify_authored(&actual, limits(size)).unwrap();
            assert_eq!(
                ReportRecord::from(report),
                expected_report(size, layout),
                "size {size}"
            );
            let streamed = run_streaming(size, limits(size)).unwrap();
            match streamed.outcome {
                RawOutcome::Success(streamed) => {
                    assert_eq!(streamed, report, "size {size}");
                },
                RawOutcome::Failure(error) => {
                    panic!("streaming audit failed at size {size}: {error:?}");
                },
            }
        }
    }
}
