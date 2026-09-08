//! A deterministic ZIP writer storage-policy benchmark.
//!
//! The benchmark compares the existing in-memory central-directory policy with
//! the explicit caller-owned directory spool.  Corpus construction, expected
//! archive generation, and all reopen oracles are outside the timed interval.
//! A timed operation constructs a fixed hashing sink and owns the complete ZIP writer lifetime;
//! for the spool lane it also opens the caller-selected `File`, writes the
//! archive, flushes the writer, and closes the spool.  Removing that file is
//! benchmark-artifact cleanup after the process snapshots and is reported as
//! excluded from the operation boundary.  No implicit temporary path is used.

use std::{
    collections::BTreeSet,
    ffi::OsString,
    fs::{self, File, OpenOptions},
    io::{self, Write},
    path::{Path, PathBuf},
    time::Instant,
};

use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::{CompressionMethod, DirectorySpoolLimits, ZipArchiveWriter};

use crate::{allocation_metrics, process_metrics};

const SCHEMA: &str = "zip-directory-spool-v1";
const DEFAULT_COUNTS: &[usize] = &[8, 256, 8_192];
const DEFAULT_METHODS: &[Method] = &[Method::Store, Method::Deflate];
const DEFAULT_SAMPLES: usize = 30;
const DEFAULT_WARMUPS: usize = 3;
const DEFAULT_REPEATS: usize = 2;
const DEFAULT_SPOOL_MAX_BYTES: u64 = 64 * 1024 * 1024;
const DEFAULT_SPOOL_BUFFER_BYTES: usize = 16 * 1024;
const PAYLOAD_BYTES: usize = 256;
const MAX_COUNT: usize = 100_000;

type BenchResult<T> = Result<T, Box<dyn std::error::Error>>;

/// Selects the writer storage policy.
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
}

/// Selects the ZIP member compression method.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Serialize)]
#[serde(rename_all = "snake_case")]
pub enum Method {
    Store,
    Deflate,
}

impl Method {
    fn parse(value: &str) -> BenchResult<Self> {
        match value {
            "store" => Ok(Self::Store),
            "deflate" => Ok(Self::Deflate),
            _ => Err(
                format!("invalid compression method {value:?}; expected store or deflate").into(),
            ),
        }
    }

    fn compression_method(self) -> CompressionMethod {
        match self {
            Self::Store => CompressionMethod::Store,
            Self::Deflate => CompressionMethod::Deflate,
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Store => "store",
            Self::Deflate => "deflate",
        }
    }
}

#[derive(Clone, Debug)]
struct Config {
    modes: Vec<Mode>,
    counts: Vec<usize>,
    methods: Vec<Method>,
    samples: usize,
    warmups: usize,
    repeats: usize,
    spool_dir: Option<PathBuf>,
    spool_limits: DirectorySpoolLimits,
    json_path: Option<PathBuf>,
}

#[derive(Debug)]
struct MemberSpec {
    name: String,
    payload: Vec<u8>,
}

#[derive(Debug)]
struct Corpus {
    count: usize,
    members: Vec<MemberSpec>,
    source_sha256: String,
    source_bytes: u64,
}

#[derive(Clone, Debug)]
struct PreparedCase {
    corpus_index: usize,
    method: Method,
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

#[derive(Debug, Serialize)]
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

    fn finish(self) -> BenchResult<SinkSummary> {
        let accepted_bytes = self.accepted_bytes;
        let write_calls = self.write_calls;
        let digest = self.digest.finalize();
        Ok(SinkSummary {
            accepted_bytes,
            write_calls,
            sha256: hex_digest(&digest),
        })
    }
}

impl Write for HashingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = u64::try_from(bytes.len()).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "output length does not fit u64",
            )
        })?;
        self.accepted_bytes = self
            .accepted_bytes
            .checked_add(length)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "output length overflow"))?;
        self.write_calls = self.write_calls.checked_add(1).ok_or_else(|| {
            io::Error::new(io::ErrorKind::InvalidInput, "write-call count overflow")
        })?;
        self.digest.update(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct TimedOperation {
    elapsed_ns: u64,
    sink: SinkSummary,
    spool_bytes: Option<u64>,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Debug, Serialize)]
struct OperationRecord {
    mode: Mode,
    method: Method,
    member_count: usize,
    repeat: usize,
    sample: usize,
    source_sha256: String,
    source_bytes: u64,
    elapsed_ns: u64,
    output_bytes: u64,
    output_write_calls: u64,
    output_sha256: String,
    output_matches_oracle: bool,
    spool_bytes: Option<u64>,
    allocation: Option<allocation_metrics::Sample>,
    process: Option<process_metrics::Delta>,
}

#[derive(Debug, Serialize)]
struct OracleRecord {
    mode: Mode,
    method: Method,
    member_count: usize,
    output_bytes: u64,
    output_sha256: String,
    every_member_reopened: bool,
    byte_exact_control_match: bool,
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
    storage_capability: &'static str,
    samples: usize,
    warmups: usize,
    repeats: usize,
    counts: Vec<usize>,
    methods: Vec<Method>,
    modes: Vec<Mode>,
    spool_max_bytes: u64,
    spool_buffer_bytes: usize,
    cases: Vec<OracleRecord>,
    operations: Vec<OperationRecord>,
}

#[derive(Debug)]
struct OracleBytes {
    bytes: Vec<u8>,
    summary: SinkSummary,
}

/// Runs the benchmark CLI.
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
        // A caller may pass the explicit directory while timing only the
        // control lane.  Keeping the spool oracle in this process makes the
        // control and spool process preflights identical for paired captures.
        fs::create_dir_all(spool_dir)?;
    }

    let corpora = config
        .counts
        .iter()
        .copied()
        .map(build_corpus)
        .collect::<BenchResult<Vec<_>>>()?;

    let spool_limits = config.spool_limits;
    let mut prepared = Vec::new();
    let mut oracle_records = Vec::new();
    let mut oracle_keys = BTreeSet::new();

    for (corpus_index, corpus) in corpora.iter().enumerate() {
        for method in config.methods.iter().copied() {
            let control = oracle_archive(corpus, method, Mode::Control, None, spool_limits)?;
            let control_sha = control.summary.sha256.clone();
            validate_oracle(corpus, &control.bytes)?;
            insert_oracle_record(
                &mut oracle_records,
                &mut oracle_keys,
                Mode::Control,
                method,
                corpus,
                &control,
                true,
            );
            prepared.extend(config.modes.iter().copied().map(|mode| PreparedCase {
                corpus_index,
                method,
                mode,
                expected_sha256: control_sha.clone(),
                expected_bytes: control.summary.accepted_bytes,
            }));

            if config.spool_dir.is_some() {
                let oracle_path = spool_path(
                    config.spool_dir.as_deref().expect("validated above"),
                    corpus.count,
                    method,
                    usize::MAX,
                    usize::MAX,
                    "oracle",
                );
                let spool_result = oracle_archive(
                    corpus,
                    method,
                    Mode::Spool,
                    Some(&oracle_path),
                    spool_limits,
                );
                let cleanup_result = remove_owned_file(&oracle_path);
                let spool = match spool_result {
                    Ok(spool) => spool,
                    Err(error) => {
                        cleanup_result?;
                        return Err(error);
                    },
                };
                cleanup_result?;
                let byte_exact = spool.bytes == control.bytes;
                if !byte_exact {
                    return Err(format!(
                        "spool oracle differs from control for {} members / {}",
                        corpus.count,
                        method.label()
                    )
                    .into());
                }
                validate_oracle(corpus, &spool.bytes)?;
                insert_oracle_record(
                    &mut oracle_records,
                    &mut oracle_keys,
                    Mode::Spool,
                    method,
                    corpus,
                    &spool,
                    byte_exact,
                );
            }
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
            for warmup in 0..config.warmups {
                let path = operation_spool_path(&config, case, repeat, warmup, true);
                let timed_result = run_timed_operation(corpus, case, path.as_deref(), spool_limits);
                let cleanup_result = cleanup_operation_path(path.as_deref());
                let timed = match timed_result {
                    Ok(timed) => timed,
                    Err(error) => {
                        cleanup_result?;
                        return Err(error);
                    },
                };
                cleanup_result?;
                if !matches_expected(case, &timed.sink) {
                    return Err(format!(
                        "warmup output failed oracle for {} members / {} / {:?}",
                        corpus.count,
                        case.method.label(),
                        case.mode
                    )
                    .into());
                }
            }
            for sample in 0..config.samples {
                let path = operation_spool_path(&config, case, repeat, sample, false);
                let timed_result = run_timed_operation(corpus, case, path.as_deref(), spool_limits);
                let cleanup_result = cleanup_operation_path(path.as_deref());
                let timed = match timed_result {
                    Ok(timed) => timed,
                    Err(error) => {
                        cleanup_result?;
                        return Err(error);
                    },
                };
                cleanup_result?;
                let matches = matches_expected(case, &timed.sink);
                let record = OperationRecord {
                    mode: case.mode,
                    method: case.method,
                    member_count: corpus.count,
                    repeat,
                    sample,
                    source_sha256: corpus.source_sha256.clone(),
                    source_bytes: corpus.source_bytes,
                    elapsed_ns: timed.elapsed_ns,
                    output_bytes: timed.sink.accepted_bytes,
                    output_write_calls: timed.sink.write_calls,
                    output_sha256: timed.sink.sha256,
                    output_matches_oracle: matches,
                    spool_bytes: timed.spool_bytes,
                    allocation: timed.allocation,
                    process: timed.process,
                };
                if !matches {
                    return Err(format!(
                        "measured output failed oracle for {} members / {} / {:?}",
                        corpus.count,
                        case.method.label(),
                        case.mode
                    )
                    .into());
                }
                operations.push(record);
            }
        }
    }

    let report = Report {
        schema: SCHEMA,
        instrumentation: if cfg!(feature = "allocator-metrics") {
            "system_allocator_operation_scoped"
        } else {
            "none"
        },
        allocator: if cfg!(feature = "allocator-metrics") {
            "CountingSystemAllocator(std::alloc::System)"
        } else {
            "Rust system allocator"
        },
        timing_scope: "corpus references, expected output, and metric observers are prepared before the clock; a fixed hashing sink is constructed at operation entry, and the clock includes explicit spool File creation/open, every local header/member payload/data descriptor, complete writer finish including central-directory publication, sink acceptance and digest updates, writer flush, and spool File close; process and allocator endpoint snapshots plus digest finalization, reopen oracles, and artifact unlink are outside",
        oracle_scope: "one control oracle and, when selected, one explicit-spool oracle per member-count/method pair; each oracle is byte-compared and reopened through ArchiveReader, then every generated member is decompressed and compared to the deterministic source payload",
        cleanup_scope: "the caller-selected spool directory is created before capture; each per-operation File uses create_new and is removed after endpoint snapshots; unlink is excluded from elapsed_ns and no ambient temporary path is selected",
        control_storage_policy: "ZipArchiveWriterBuilder::with_capacity(member_count).build uses the existing in-memory Vec<FileHeader> plus contiguous member-name buffer; this is the measured control policy",
        storage_capability: "explicit caller-owned std::fs::File opened read/write/seek and passed to DirectorySpoolLimits; no sync_all is issued by the harness",
        samples: config.samples,
        warmups: config.warmups,
        repeats: config.repeats,
        counts: config.counts,
        methods: config.methods,
        modes: config.modes,
        spool_max_bytes: config.spool_limits.max_bytes,
        spool_buffer_bytes: config.spool_limits.buffer_bytes,
        cases: oracle_records,
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

fn insert_oracle_record(
    records: &mut Vec<OracleRecord>,
    keys: &mut BTreeSet<String>,
    mode: Mode,
    method: Method,
    corpus: &Corpus,
    oracle: &OracleBytes,
    byte_exact_control_match: bool,
) {
    let key = format!("{:?}-{:?}-{}", mode, method, corpus.count);
    if keys.insert(key) {
        records.push(OracleRecord {
            mode,
            method,
            member_count: corpus.count,
            output_bytes: oracle.summary.accepted_bytes,
            output_sha256: oracle.summary.sha256.clone(),
            every_member_reopened: true,
            byte_exact_control_match,
        });
    }
}

fn run_timed_operation(
    corpus: &Corpus,
    case: &PreparedCase,
    spool_path: Option<&Path>,
    spool_limits: DirectorySpoolLimits,
) -> BenchResult<TimedOperation> {
    let process_before = process_metrics::Snapshot::read().ok();
    let allocation_region = allocation_metrics::begin();
    let started = Instant::now();
    let output = create_archive(corpus, case.method, case.mode, spool_path, spool_limits);
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos())
        .map_err(|_| "operation elapsed time does not fit u64")?;
    let allocation = allocation_region.finish();
    let process = process_metrics::Snapshot::read()
        .ok()
        .zip(process_before)
        .map(|(after, before)| after.delta(before));

    let spool_bytes = if case.mode == Mode::Spool {
        let path = spool_path.ok_or("spool mode did not receive an explicit path")?;
        Some(
            fs::metadata(path)
                .map_err(|error| {
                    format!("cannot inspect completed spool {}: {error}", path.display())
                })?
                .len(),
        )
    } else {
        None
    };
    let sink = output?.finish()?;
    Ok(TimedOperation {
        elapsed_ns,
        sink,
        spool_bytes,
        allocation,
        process,
    })
}

fn create_archive(
    corpus: &Corpus,
    method: Method,
    mode: Mode,
    spool_path: Option<&Path>,
    spool_limits: DirectorySpoolLimits,
) -> BenchResult<HashingSink> {
    let sink = HashingSink::new();
    let mut archive = match mode {
        Mode::Control => ZipArchiveWriter::builder()
            .with_capacity(corpus.members.len())
            .build(sink),
        Mode::Spool => {
            let path = spool_path.ok_or("spool mode requires a caller-selected spool path")?;
            let spool = open_spool(path)?;
            ZipArchiveWriter::builder()
                .with_capacity(corpus.members.len())
                .build_with_spool(sink, spool, spool_limits)?
        },
    };

    for member in &corpus.members {
        let mut entry =
            archive.start_file_owned(member.name.as_str(), method.compression_method())?;
        entry.write_all(&member.payload)?;
        archive = entry.finish()?;
    }
    Ok(archive.finish()?)
}

fn oracle_archive(
    corpus: &Corpus,
    method: Method,
    mode: Mode,
    spool_path: Option<&Path>,
    spool_limits: DirectorySpoolLimits,
) -> BenchResult<OracleBytes> {
    let output = Vec::new();
    let mut archive = match mode {
        Mode::Control => ZipArchiveWriter::builder()
            .with_capacity(corpus.members.len())
            .build(output),
        Mode::Spool => {
            let path = spool_path.ok_or("spool oracle requires a caller-selected spool path")?;
            let spool = open_spool(path)?;
            ZipArchiveWriter::builder()
                .with_capacity(corpus.members.len())
                .build_with_spool(output, spool, spool_limits)?
        },
    };
    for member in &corpus.members {
        let mut entry =
            archive.start_file_owned(member.name.as_str(), method.compression_method())?;
        entry.write_all(&member.payload)?;
        archive = entry.finish()?;
    }
    let bytes = archive.finish()?;
    let mut hasher = Sha256::new();
    hasher.update(&bytes);
    let summary = SinkSummary {
        accepted_bytes: u64::try_from(bytes.len())?,
        write_calls: 0,
        sha256: hex_digest(&hasher.finalize()),
    };
    Ok(OracleBytes { bytes, summary })
}

fn validate_oracle(corpus: &Corpus, bytes: &[u8]) -> BenchResult<()> {
    let reader = soapberry_zip::office::ArchiveReader::new(bytes)?;
    if reader.len() != corpus.count {
        return Err(format!(
            "oracle has {} files, expected {}",
            reader.len(),
            corpus.count
        )
        .into());
    }
    let names: Vec<&str> = reader.file_names().collect();
    if names.len() != corpus.count {
        return Err(format!(
            "oracle has {} indexed names, expected {}",
            names.len(),
            corpus.count
        )
        .into());
    }
    for member in &corpus.members {
        if !reader.contains(&member.name) {
            return Err(format!("oracle is missing {}", member.name).into());
        }
        let actual = reader.read(&member.name)?;
        if actual != member.payload {
            return Err(format!("oracle payload mismatch for {}", member.name).into());
        }
    }
    Ok(())
}

fn matches_expected(case: &PreparedCase, sink: &SinkSummary) -> bool {
    sink.accepted_bytes == case.expected_bytes && sink.sha256 == case.expected_sha256
}

fn build_corpus(count: usize) -> BenchResult<Corpus> {
    if count == 0 || count > MAX_COUNT {
        return Err(format!("member count must be in 1..={MAX_COUNT}, got {count}").into());
    }
    let mut members = Vec::new();
    members.try_reserve_exact(count)?;
    let mut source_hasher = Sha256::new();
    let mut source_bytes = 0u64;
    for index in 0..count {
        let name = format!("ppt/slides/slide-{index:05}.xml");
        let payload = deterministic_payload(index);
        source_hasher.update((name.len() as u64).to_le_bytes());
        source_hasher.update(name.as_bytes());
        source_hasher.update((payload.len() as u64).to_le_bytes());
        source_hasher.update(&payload);
        source_bytes = source_bytes
            .checked_add(u64::try_from(payload.len())?)
            .ok_or("source byte count overflow")?;
        members.push(MemberSpec { name, payload });
    }
    Ok(Corpus {
        count,
        members,
        source_sha256: hex_digest(&source_hasher.finalize()),
        source_bytes,
    })
}

fn deterministic_payload(index: usize) -> Vec<u8> {
    let mut value = (index as u64).wrapping_add(0x9e37_79b9_7f4a_7c15);
    let mut payload = vec![0u8; PAYLOAD_BYTES];
    for byte in &mut payload {
        value ^= value >> 30;
        value = value.wrapping_mul(0xbf58_476d_1ce4_e5b9);
        value ^= value >> 27;
        value = value.wrapping_mul(0x94d0_49bb_1331_11eb);
        value ^= value >> 31;
        *byte = value as u8;
    }
    payload
}

fn hex_digest(digest: &[u8]) -> String {
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn open_spool(path: &Path) -> io::Result<File> {
    OpenOptions::new()
        .create_new(true)
        .read(true)
        .write(true)
        .open(path)
}

fn spool_path(
    root: &Path,
    count: usize,
    method: Method,
    repeat: usize,
    sample: usize,
    phase: &str,
) -> PathBuf {
    root.join(format!(
        "zip-directory-spool-{count}-{}-{repeat}-{sample}-{phase}.tmp",
        method.label()
    ))
}

fn operation_spool_path(
    config: &Config,
    case: &PreparedCase,
    repeat: usize,
    sample: usize,
    warmup: bool,
) -> Option<PathBuf> {
    (case.mode == Mode::Spool).then(|| {
        spool_path(
            config.spool_dir.as_deref().expect("validated above"),
            case.corpus_index,
            case.method,
            repeat,
            sample,
            if warmup { "warmup" } else { "sample" },
        )
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
    let mut methods = DEFAULT_METHODS.to_vec();
    let mut samples = DEFAULT_SAMPLES;
    let mut warmups = DEFAULT_WARMUPS;
    let mut repeats = DEFAULT_REPEATS;
    let mut spool_dir = None;
    let mut spool_limits = DirectorySpoolLimits {
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
            "--methods" => methods = parse_methods(&next_value(&mut values, "--methods")?)?,
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
                spool_limits.max_bytes = parse_positive_u64(
                    &next_value(&mut values, "--max-spool-bytes")?,
                    "--max-spool-bytes",
                )?
            },
            "--spool-buffer-bytes" => {
                spool_limits.buffer_bytes = parse_positive(
                    &next_value(&mut values, "--spool-buffer-bytes")?,
                    "--spool-buffer-bytes",
                )?
            },
            "--json" => json_path = Some(PathBuf::from(next_value(&mut values, "--json")?)),
            "--help" | "-h" => return Err(usage().into()),
            unknown => return Err(format!("unknown argument {unknown}; {}", usage()).into()),
        }
    }
    if samples == 0 || warmups == 0 || repeats == 0 || counts.is_empty() || methods.is_empty() {
        return Err("samples, warmups, repeats, counts, and methods must be non-zero".into());
    }
    if modes.is_empty() {
        return Err("at least one mode is required".into());
    }
    Ok(Config {
        modes,
        counts,
        methods,
        samples,
        warmups,
        repeats,
        spool_dir,
        spool_limits,
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
        let mode = if part == "both" {
            for mode in [Mode::Control, Mode::Spool] {
                if !modes.contains(&mode) {
                    modes.push(mode);
                }
            }
            continue;
        } else {
            Mode::parse(part)?
        };
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
        if count == 0 || count > MAX_COUNT {
            return Err(format!("count must be in 1..={MAX_COUNT}, got {count}").into());
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

fn parse_methods(value: &OsString) -> BenchResult<Vec<Method>> {
    let text = value.to_str().ok_or("--methods must be UTF-8")?;
    let mut methods = Vec::new();
    for part in text.split(',') {
        let method = if part == "both" {
            for method in [Method::Store, Method::Deflate] {
                if !methods.contains(&method) {
                    methods.push(method);
                }
            }
            continue;
        } else {
            Method::parse(part)?
        };
        if !methods.contains(&method) {
            methods.push(method);
        }
    }
    if methods.is_empty() {
        return Err("--methods cannot be empty".into());
    }
    Ok(methods)
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

fn usage() -> &'static str {
    "usage: zip_directory_spool [--mode control|spool|both] [--counts 8,256,8192] [--methods store,deflate] [--samples N] [--warmups N] [--repeats N] [--spool-dir PATH] [--max-spool-bytes N] [--spool-buffer-bytes N] [--json PATH]"
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_cli_is_control_only() {
        let config = parse_args(Vec::<OsString>::new()).unwrap();
        assert_eq!(config.modes, vec![Mode::Control]);
        assert_eq!(config.counts, DEFAULT_COUNTS);
        assert_eq!(config.methods, DEFAULT_METHODS);
        assert_eq!(config.samples, DEFAULT_SAMPLES);
        assert_eq!(config.warmups, DEFAULT_WARMUPS);
        assert_eq!(config.repeats, DEFAULT_REPEATS);
    }

    #[test]
    fn both_mode_and_method_lists_deduplicate() {
        let args = [
            OsString::from("--mode"),
            OsString::from("both,control"),
            OsString::from("--methods"),
            OsString::from("both,store"),
            OsString::from("--counts"),
            OsString::from("8,8,256"),
        ];
        let config = parse_args(args).unwrap();
        assert_eq!(config.modes, vec![Mode::Control, Mode::Spool]);
        assert_eq!(config.methods, vec![Method::Store, Method::Deflate]);
        assert_eq!(config.counts, vec![8, 256]);
    }

    #[test]
    fn corpus_generation_is_deterministic() {
        let left = build_corpus(8).unwrap();
        let right = build_corpus(8).unwrap();
        assert_eq!(left.source_sha256, right.source_sha256);
        assert_eq!(left.source_bytes, right.source_bytes);
        assert_eq!(
            left.members
                .iter()
                .map(|member| (&member.name, &member.payload))
                .collect::<Vec<_>>(),
            right
                .members
                .iter()
                .map(|member| (&member.name, &member.payload))
                .collect::<Vec<_>>()
        );
    }

    #[test]
    fn hash_sink_counts_every_accepted_byte() {
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
}
