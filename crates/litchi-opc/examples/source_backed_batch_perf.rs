//! Small, process-friendly benchmark for source-backed OPC Part loads.
//!
//! The harness deliberately keeps package construction and Part selection out
//! of the timed region.  Each sample opens a fresh package, selects the same
//! deterministic Part set, resets the source observer, and then times only
//! the selected `PartView::data` calls.  The default API is the existing
//! serial read path; the worker option is a bounded benchmark harness for
//! disjoint reads and does not imply a production scheduler.

#![allow(
    clippy::cast_precision_loss,
    clippy::cast_possible_truncation,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::missing_errors_doc,
    clippy::print_stdout,
    clippy::unwrap_used,
    reason = "this is an opt-in benchmark executable; malformed CLI and fixture failures abort"
)]

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, FileSource, Limits, OwnedSource,
    ReadAt, Resource, SourceVersion,
};
use litchi_opc::{
    OpcOperationAccounting, PackURI, PartBatch, PartData, ReadLimits, SourceBackedPackage,
    SourceCacheDiagnostics, SourceCacheLimits,
};
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;
use std::fmt::Write as _;
use std::fs;
use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::{Path, PathBuf};
use std::sync::{
    Arc,
    atomic::{AtomicU64, AtomicUsize, Ordering},
};
use std::thread;
use std::time::{Duration, Instant};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const MIB: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum CorpusKind {
    FewLarge,
    ManySmall,
}

impl CorpusKind {
    fn as_str(self) -> &'static str {
        match self {
            Self::FewLarge => "few-large",
            Self::ManySmall => "many-small",
        }
    }

    fn selected_count(self) -> usize {
        match self {
            Self::FewLarge => 4,
            Self::ManySmall => 64,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SourceKind {
    Owned,
    File,
    ShortDelay,
}

impl SourceKind {
    fn as_str(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::File => "file",
            Self::ShortDelay => "instrumented-short-delay",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Capability {
    Normal,
    ManagedAfterOnly,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ApiKind {
    Serial,
    Batch,
}

impl ApiKind {
    fn as_str(self) -> &'static str {
        match self {
            Self::Serial => "serial-part-data",
            Self::Batch => "read-parts-ordered",
        }
    }
}

impl Capability {
    fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::ManagedAfterOnly => "managed-after-only",
        }
    }
}

#[derive(Debug)]
struct Config {
    corpora: Vec<CorpusKind>,
    sources: Vec<SourceKind>,
    apis: Vec<ApiKind>,
    capabilities: Vec<Capability>,
    workers: Vec<usize>,
    warmups: usize,
    samples: usize,
    repeats: usize,
    short_read_bytes: usize,
    delay_us: u64,
    accounting: bool,
    artifact_dir: PathBuf,
    output: Option<PathBuf>,
    keep_artifacts: bool,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            corpora: vec![CorpusKind::FewLarge, CorpusKind::ManySmall],
            sources: vec![SourceKind::Owned, SourceKind::File, SourceKind::ShortDelay],
            apis: vec![ApiKind::Serial],
            capabilities: vec![Capability::Normal],
            workers: vec![1, 2, 4, 8],
            warmups: 3,
            samples: 30,
            repeats: 2,
            short_read_bytes: 4096,
            delay_us: 25,
            accounting: false,
            artifact_dir: std::env::temp_dir()
                .join(format!("litchi-change-0498-{}", std::process::id())),
            output: None,
            keep_artifacts: false,
        }
    }
}

#[derive(Debug)]
struct Corpus {
    kind: CorpusKind,
    bytes: Arc<Vec<u8>>,
    sha256: String,
    part_names: Vec<String>,
    path: PathBuf,
}

#[derive(Debug, Default)]
struct ProbeCounters {
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
    max_request: AtomicUsize,
    le_4k: AtomicU64,
    le_16k: AtomicU64,
    le_64k: AtomicU64,
    le_256k: AtomicU64,
    gt_256k: AtomicU64,
}

#[derive(Clone, Copy, Debug, Default)]
struct ProbeSnapshot {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
    max_request: usize,
    le_4k: u64,
    le_16k: u64,
    le_64k: u64,
    le_256k: u64,
    gt_256k: u64,
}

impl ProbeCounters {
    fn reset(&self) {
        self.calls.store(0, Ordering::Relaxed);
        self.requested_bytes.store(0, Ordering::Relaxed);
        self.returned_bytes.store(0, Ordering::Relaxed);
        self.max_request.store(0, Ordering::Relaxed);
        self.le_4k.store(0, Ordering::Relaxed);
        self.le_16k.store(0, Ordering::Relaxed);
        self.le_64k.store(0, Ordering::Relaxed);
        self.le_256k.store(0, Ordering::Relaxed);
        self.gt_256k.store(0, Ordering::Relaxed);
    }

    fn record(&self, requested: usize, returned: usize) {
        self.calls.fetch_add(1, Ordering::Relaxed);
        self.requested_bytes
            .fetch_add(requested as u64, Ordering::Relaxed);
        self.returned_bytes
            .fetch_add(returned as u64, Ordering::Relaxed);
        update_max(&self.max_request, requested);
        match requested {
            0..=4_096 => {
                self.le_4k.fetch_add(1, Ordering::Relaxed);
            },
            4_097..=16_384 => {
                self.le_16k.fetch_add(1, Ordering::Relaxed);
            },
            16_385..=65_536 => {
                self.le_64k.fetch_add(1, Ordering::Relaxed);
            },
            65_537..=262_144 => {
                self.le_256k.fetch_add(1, Ordering::Relaxed);
            },
            _ => {
                self.gt_256k.fetch_add(1, Ordering::Relaxed);
            },
        }
    }

    fn snapshot(&self) -> ProbeSnapshot {
        ProbeSnapshot {
            calls: self.calls.load(Ordering::Relaxed),
            requested_bytes: self.requested_bytes.load(Ordering::Relaxed),
            returned_bytes: self.returned_bytes.load(Ordering::Relaxed),
            max_request: self.max_request.load(Ordering::Relaxed),
            le_4k: self.le_4k.load(Ordering::Relaxed),
            le_16k: self.le_16k.load(Ordering::Relaxed),
            le_64k: self.le_64k.load(Ordering::Relaxed),
            le_256k: self.le_256k.load(Ordering::Relaxed),
            gt_256k: self.gt_256k.load(Ordering::Relaxed),
        }
    }
}

struct TrackedSource {
    inner: Arc<dyn ReadAt>,
    counters: Arc<ProbeCounters>,
}

impl ReadAt for TrackedSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let requested = output.len();
        let result = self.inner.read_at(offset, output);
        if let Ok(returned) = result {
            self.counters.record(requested, returned);
        } else {
            self.counters.record(requested, 0);
        }
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct ShortDelaySource {
    inner: Arc<dyn ReadAt>,
    max_read: usize,
    delay: Duration,
}

impl ReadAt for ShortDelaySource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if !self.delay.is_zero() {
            thread::sleep(self.delay);
        }
        let amount = output.len().min(self.max_read);
        self.inner.read_at(offset, &mut output[..amount])
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn update_max(slot: &AtomicUsize, candidate: usize) {
    let mut current = slot.load(Ordering::Relaxed);
    while candidate > current {
        match slot.compare_exchange_weak(current, candidate, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => return,
            Err(next) => current = next,
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct Sample {
    elapsed_ns: u128,
    logical_bytes: u64,
    digest: u64,
    source: ProbeSnapshot,
    cache: SourceCacheDiagnostics,
    cache_delta: litchi_opc::SourceCacheCounterDelta,
    accounting: AccountingSnapshot,
    retained_before_drop: usize,
    budget_memory_after_read: u64,
    budget_memory_after_release: u64,
    budget_objects_after_read: u64,
    budget_objects_after_release: u64,
    budget_input_bytes: u64,
    budget_work: u64,
}

#[derive(Clone, Copy, Debug, Default)]
struct AccountingSnapshot {
    deflated_read: u64,
    stored_read: u64,
    deflated_produced: u64,
    stored_accepted: u64,
}

impl From<&OpcOperationAccounting> for AccountingSnapshot {
    fn from(value: &OpcOperationAccounting) -> Self {
        Self {
            deflated_read: value.compressed_deflate_payload_bytes_read(),
            stored_read: value.stored_payload_bytes_read(),
            deflated_produced: value.deflate_bytes_produced(),
            stored_accepted: value.stored_payload_bytes_accepted(),
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct Summary {
    p50_ns: u128,
    p95_ns: u128,
    p99_ns: u128,
    mean_ns: u128,
    throughput_bytes_s: f64,
    source_calls: f64,
    source_requested_bytes: f64,
    source_returned_bytes: f64,
    source_max_request: usize,
    cache_cold_loads: f64,
    cache_hits: f64,
    retained_bytes: usize,
    budget_memory_after_release: u64,
    budget_objects_after_release: u64,
}

enum Loaded {
    Serial(Vec<PartData>),
    Batch(PartBatch),
}

fn verify_loaded(loaded: &Loaded) -> (u64, u64) {
    match loaded {
        Loaded::Serial(payloads) => verify_payloads(payloads),
        Loaded::Batch(batch) => batch.iter().fold((0_u64, 0_u64), |(bytes, digest), data| {
            (
                bytes.saturating_add(data.as_bytes().len() as u64),
                fold_digest(digest, data.as_bytes()),
            )
        }),
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = parse_args()?;
    validate_config(&config)?;
    fs::create_dir_all(&config.artifact_dir)?;
    let corpora = build_corpora(&config.artifact_dir)?;
    let output_path = config.output.clone();
    let mut output: Box<dyn io::Write> = match output_path {
        Some(path) => Box::new(io::BufWriter::new(fs::File::create(path)?)),
        None => Box::new(io::BufWriter::new(io::stdout())),
    };
    write_header(&mut output)?;

    for corpus in &corpora {
        if !config.corpora.contains(&corpus.kind) {
            continue;
        }
        for &source in &config.sources {
            for &api in &config.apis {
                for &capability in &config.capabilities {
                    for &workers in &config.workers {
                        let mut all_samples = Vec::new();
                        for repeat in 0..config.repeats {
                            for warmup in 0..config.warmups {
                                let sample =
                                    run_sample(corpus, source, api, capability, workers, &config)?;
                                write_sample(
                                    &mut output,
                                    corpus,
                                    source,
                                    api,
                                    capability,
                                    workers,
                                    repeat,
                                    warmup,
                                    sample,
                                )?;
                            }
                            for sample_index in 0..config.samples {
                                let sample =
                                    run_sample(corpus, source, api, capability, workers, &config)?;
                                write_sample(
                                    &mut output,
                                    corpus,
                                    source,
                                    api,
                                    capability,
                                    workers,
                                    repeat,
                                    config.warmups + sample_index,
                                    sample,
                                )?;
                                all_samples.push(sample);
                            }
                        }
                        let summary = summarize(&all_samples);
                        write_summary(
                            &mut output,
                            corpus,
                            source,
                            api,
                            capability,
                            workers,
                            summary,
                        )?;
                        output.flush()?;
                        eprintln!(
                            "completed corpus={} source={} api={} capability={} workers={} samples={} p50={}us p95={}us p99={}us",
                            corpus.kind.as_str(),
                            source.as_str(),
                            api.as_str(),
                            capability.as_str(),
                            workers,
                            all_samples.len(),
                            summary.p50_ns / 1_000,
                            summary.p95_ns / 1_000,
                            summary.p99_ns / 1_000
                        );
                    }
                }
            }
        }
    }

    if !config.keep_artifacts {
        for corpus in &corpora {
            let _ = fs::remove_file(&corpus.path);
        }
        let _ = fs::remove_dir(&config.artifact_dir);
    }
    Ok(())
}

fn validate_config(config: &Config) -> Result<(), Box<dyn std::error::Error>> {
    if config.samples == 0 || config.repeats == 0 {
        return Err("samples and repeats must be positive".into());
    }
    if config.workers.contains(&0) {
        return Err("worker counts must be positive".into());
    }
    if config.short_read_bytes == 0 {
        return Err("short-read size must be positive".into());
    }
    Ok(())
}

fn build_corpora(artifact_dir: &Path) -> io::Result<Vec<Corpus>> {
    [CorpusKind::FewLarge, CorpusKind::ManySmall]
        .into_iter()
        .map(|kind| {
            let bytes = Arc::new(make_archive(kind));
            let sha256 = hex_digest(&bytes);
            let part_count = match kind {
                CorpusKind::FewLarge => 4,
                CorpusKind::ManySmall => 512,
            };
            let part_names = (0..part_count)
                .map(|index| format!("/parts/{}/{:04}.bin", kind.as_str(), index))
                .collect::<Vec<_>>();
            let path = artifact_dir.join(format!("{}.opc", kind.as_str()));
            fs::write(&path, bytes.as_slice())?;
            Ok(Corpus {
                kind,
                bytes,
                sha256,
                part_names,
                path,
            })
        })
        .collect()
}

fn make_archive(kind: CorpusKind) -> Vec<u8> {
    let part_count = match kind {
        CorpusKind::FewLarge => 4,
        CorpusKind::ManySmall => 512,
    };
    let payload_bytes = match kind {
        CorpusKind::FewLarge => MIB,
        CorpusKind::ManySmall => 16 * 1024,
    };
    let first_target = format!("parts/{}/0000.bin", kind.as_str());
    let content_types = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_relationships = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{first_target}"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("benchmark manifest must be writable");
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .expect("benchmark relationships must be writable");
    for index in 0..part_count {
        let name = format!("parts/{}/{:04}.bin", kind.as_str(), index);
        let payload = deterministic_payload(kind, index, payload_bytes);
        writer
            .write_stored(&name, &payload)
            .expect("benchmark Part must be writable");
    }
    writer
        .finish_to_bytes()
        .expect("benchmark archive must finish")
}

fn deterministic_payload(kind: CorpusKind, part: usize, length: usize) -> Vec<u8> {
    let mut payload = vec![0_u8; length];
    let seed = (part as u64).wrapping_mul(0x9e37_79b9_7f4a_7c15)
        ^ match kind {
            CorpusKind::FewLarge => 0x4645_572d_4c41_5247,
            CorpusKind::ManySmall => 0x4d41_4e59_2d53_4d41,
        };
    for (offset, byte) in payload.iter_mut().enumerate() {
        let value = seed
            .wrapping_add(offset as u64)
            .rotate_left((offset & 31) as u32);
        *byte = (value ^ (value >> 17) ^ (value >> 41)) as u8;
    }
    payload
}

fn hex_digest(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest
        .iter()
        .fold(String::with_capacity(64), |mut text, byte| {
            let _ = write!(text, "{byte:02x}");
            text
        })
}

fn run_sample(
    corpus: &Corpus,
    source_kind: SourceKind,
    api: ApiKind,
    capability: Capability,
    workers: usize,
    config: &Config,
) -> Result<Sample, Box<dyn std::error::Error>> {
    let counters = Arc::new(ProbeCounters::default());
    let source_inner: Arc<dyn ReadAt> = match source_kind {
        SourceKind::Owned | SourceKind::ShortDelay => {
            let owned: Arc<dyn ReadAt> = Arc::new(OwnedSource::from_arc(Arc::clone(&corpus.bytes)));
            if source_kind == SourceKind::ShortDelay {
                Arc::new(ShortDelaySource {
                    inner: owned,
                    max_read: config.short_read_bytes,
                    delay: Duration::from_micros(config.delay_us),
                })
            } else {
                owned
            }
        },
        SourceKind::File => Arc::new(FileSource::open(&corpus.path)?),
    };
    let source: Arc<dyn ReadAt> = Arc::new(TrackedSource {
        inner: source_inner,
        counters: Arc::clone(&counters),
    });
    let (context, budget) = if capability == Capability::ManagedAfterOnly {
        let budget = Budget::root(
            "change-0498-benchmark",
            Limits::new(
                1024 * 1024 * 1024,
                8 * 1024 * 1024 * 1024,
                8 * 1024 * 1024 * 1024,
                2_000_000,
                1024,
                8 * 1024 * 1024 * 1024,
            ),
        );
        let (_cancel_source, cancellation) = CancellationSource::pair();
        let execution_limits = ExecutionLimits::new(
            NonZeroUsize::new(workers.max(1)).expect("workers are positive"),
            NonZeroUsize::new(workers.max(1)).expect("workers are positive"),
            NonZeroU64::new(1024 * 1024 * 1024).expect("in-flight bytes are positive"),
            0,
        )?;
        (
            Some(ExecutionContext::new(
                budget.clone(),
                cancellation,
                execution_limits,
            )),
            Some(budget),
        )
    } else {
        (None, None)
    };
    let cache_limits = SourceCacheLimits::new(512 * MIB, 1024)?;
    let package = match context {
        Some(context) => {
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                ReadLimits::default(),
                cache_limits,
                context,
            )?
        },
        None => SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
            source,
            ReadLimits::default(),
            cache_limits,
        )?,
    };
    let selected = selected_names(corpus)
        .iter()
        .map(PackURI::new)
        .collect::<Result<Vec<PackURI>, _>>()?;
    counters.reset();
    let before_cache = package.cache_diagnostics();
    let mut accounting = OpcOperationAccounting::new();
    let started = Instant::now();
    let loaded = match api {
        ApiKind::Serial => Loaded::Serial(read_uris(
            &package,
            &selected,
            workers,
            config.accounting.then_some(&mut accounting),
        )?),
        ApiKind::Batch => Loaded::Batch(package.read_parts_ordered(&selected)?),
    };
    let elapsed_ns = started.elapsed().as_nanos();
    let after_cache = package.cache_diagnostics();
    let cache_delta = SourceCacheDiagnostics::checked_counter_delta(before_cache, after_cache)?;
    let source = counters.snapshot();
    let accounting_snapshot = AccountingSnapshot::from(&accounting);
    let retained_before_drop = after_cache.retained_bytes;
    let budget_input_bytes = after_cache.budget_input_bytes_used;
    let budget_work = after_cache.budget_work_used;
    // Keep returned payload handles live through the post-read cache/budget
    // observation. Verification is deliberately outside the timed interval,
    // matching a future batch result that returns a retained collection.
    let (logical_bytes, digest) = verify_loaded(&loaded);
    let budget_memory_after_read = budget
        .as_ref()
        .map_or(0, |budget| budget.used(Resource::Memory));
    let budget_objects_after_read = budget
        .as_ref()
        .map_or(0, |budget| budget.used(Resource::Objects));
    drop(loaded);
    drop(package);
    let budget_memory_after_release = budget
        .as_ref()
        .map_or(0, |budget| budget.used(Resource::Memory));
    let budget_objects_after_release = budget
        .as_ref()
        .map_or(0, |budget| budget.used(Resource::Objects));
    Ok(Sample {
        elapsed_ns,
        logical_bytes,
        digest,
        source,
        cache: after_cache,
        cache_delta,
        accounting: accounting_snapshot,
        retained_before_drop,
        budget_memory_after_read,
        budget_memory_after_release,
        budget_objects_after_read,
        budget_objects_after_release,
        budget_input_bytes,
        budget_work,
    })
}

fn selected_names(corpus: &Corpus) -> Vec<String> {
    let selected = corpus.kind.selected_count();
    corpus
        .part_names
        .iter()
        .step_by((corpus.part_names.len() / selected).max(1))
        .take(selected)
        .cloned()
        .collect()
}

fn read_uris(
    package: &SourceBackedPackage,
    uris: &[PackURI],
    workers: usize,
    accounting: Option<&mut OpcOperationAccounting>,
) -> Result<Vec<PartData>, Box<dyn std::error::Error>> {
    let workers = workers.min(uris.len().max(1));
    if workers == 1 {
        let mut payloads = Vec::with_capacity(uris.len());
        let mut accounting = accounting;
        for uri in uris {
            let data = match accounting.as_deref_mut() {
                Some(accounting) => package.part(uri)?.data_with_accounting(accounting)?,
                None => package.part(uri)?.data()?,
            };
            payloads.push(data);
        }
        return Ok(payloads);
    }

    let values = thread::scope(|scope| {
        let mut handles = Vec::with_capacity(workers);
        for worker in 0..workers {
            handles.push(scope.spawn(move || {
                let mut values = Vec::new();
                for index in (worker..uris.len()).step_by(workers) {
                    let data = package
                        .part(&uris[index])
                        .expect("benchmark Part lookup must succeed")
                        .data()
                        .expect("benchmark Part read must succeed");
                    values.push((index, data));
                }
                values
            }));
        }
        let mut values = Vec::with_capacity(uris.len());
        for handle in handles {
            values.extend(handle.join().expect("benchmark worker must not panic"));
        }
        values.sort_unstable_by_key(|(index, _)| *index);
        values.into_iter().map(|(_, value)| value).collect()
    });
    // The parallel harness has independent operation reports by design; its
    // accounting values remain zero rather than being fabricated by merging
    // private reports through a production API that does not exist yet.
    Ok(values)
}

fn verify_payloads(payloads: &[PartData]) -> (u64, u64) {
    payloads
        .iter()
        .fold((0_u64, 0_u64), |(bytes, digest), data| {
            (
                bytes.saturating_add(data.as_bytes().len() as u64),
                fold_digest(digest, data.as_bytes()),
            )
        })
}

fn fold_digest(prior: u64, bytes: &[u8]) -> u64 {
    bytes.iter().fold(prior.rotate_left(7), |hash, byte| {
        hash.wrapping_mul(1_099_511_628_211)
            .wrapping_add(*byte as u64)
    })
}

fn summarize(samples: &[Sample]) -> Summary {
    let mut elapsed = samples
        .iter()
        .map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    elapsed.sort_unstable();
    let p = |rank: usize| elapsed[rank.min(elapsed.len() - 1)];
    let total_ns = elapsed.iter().sum::<u128>();
    let total_bytes = samples
        .iter()
        .map(|sample| sample.logical_bytes)
        .sum::<u64>();
    let mean = total_ns / elapsed.len() as u128;
    Summary {
        p50_ns: p((elapsed.len() * 50).div_ceil(100).saturating_sub(1)),
        p95_ns: p((elapsed.len() * 95).div_ceil(100).saturating_sub(1)),
        p99_ns: p((elapsed.len() * 99).div_ceil(100).saturating_sub(1)),
        mean_ns: mean,
        throughput_bytes_s: (total_bytes as f64) * 1e9 / (total_ns as f64),
        source_calls: samples
            .iter()
            .map(|sample| sample.source.calls as f64)
            .sum::<f64>()
            / samples.len() as f64,
        source_requested_bytes: samples
            .iter()
            .map(|sample| sample.source.requested_bytes as f64)
            .sum::<f64>()
            / samples.len() as f64,
        source_returned_bytes: samples
            .iter()
            .map(|sample| sample.source.returned_bytes as f64)
            .sum::<f64>()
            / samples.len() as f64,
        source_max_request: samples
            .iter()
            .map(|sample| sample.source.max_request)
            .max()
            .unwrap_or(0),
        cache_cold_loads: samples
            .iter()
            .map(|sample| sample.cache_delta.cold_loads as f64)
            .sum::<f64>()
            / samples.len() as f64,
        cache_hits: samples
            .iter()
            .map(|sample| sample.cache_delta.hits as f64)
            .sum::<f64>()
            / samples.len() as f64,
        retained_bytes: samples
            .iter()
            .map(|sample| sample.retained_before_drop)
            .max()
            .unwrap_or(0),
        budget_memory_after_release: samples
            .iter()
            .map(|sample| sample.budget_memory_after_release)
            .max()
            .unwrap_or(0),
        budget_objects_after_release: samples
            .iter()
            .map(|sample| sample.budget_objects_after_release)
            .max()
            .unwrap_or(0),
    }
}

fn write_header(output: &mut dyn io::Write) -> io::Result<()> {
    writeln!(
        output,
        "record,corpus,corpus_sha256,source,api,capability,workers,repeat,index,elapsed_ns,logical_bytes,digest,source_calls,source_requested_bytes,source_returned_bytes,source_max_request,source_le_4k,source_le_16k,source_le_64k,source_le_256k,source_gt_256k,cache_cold_loads,cache_hits,cache_retained_bytes,cache_retained_entries,cache_budget_memory_used,cache_budget_objects_used,budget_memory_after_read,budget_memory_after_release,budget_objects_after_read,budget_objects_after_release,budget_input_bytes,budget_work,accounting_deflated_read,accounting_stored_read,accounting_deflated_produced,accounting_stored_accepted,p50_ns,p95_ns,p99_ns,mean_ns,throughput_bytes_s,summary_source_calls,summary_source_requested_bytes,summary_source_returned_bytes,summary_source_max_request,summary_cache_cold_loads,summary_cache_hits,summary_retained_bytes,summary_budget_memory_after_release,summary_budget_objects_after_release"
    )
}

fn write_sample(
    output: &mut dyn io::Write,
    corpus: &Corpus,
    source: SourceKind,
    api: ApiKind,
    capability: Capability,
    workers: usize,
    repeat: usize,
    index: usize,
    sample: Sample,
) -> io::Result<()> {
    let fields = vec![
        "sample".to_string(),
        corpus.kind.as_str().to_string(),
        corpus.sha256.clone(),
        source.as_str().to_string(),
        api.as_str().to_string(),
        capability.as_str().to_string(),
        workers.to_string(),
        repeat.to_string(),
        index.to_string(),
        sample.elapsed_ns.to_string(),
        sample.logical_bytes.to_string(),
        sample.digest.to_string(),
        sample.source.calls.to_string(),
        sample.source.requested_bytes.to_string(),
        sample.source.returned_bytes.to_string(),
        sample.source.max_request.to_string(),
        sample.source.le_4k.to_string(),
        sample.source.le_16k.to_string(),
        sample.source.le_64k.to_string(),
        sample.source.le_256k.to_string(),
        sample.source.gt_256k.to_string(),
        sample.cache_delta.cold_loads.to_string(),
        sample.cache_delta.hits.to_string(),
        sample.cache.retained_bytes.to_string(),
        sample.cache.retained_entries.to_string(),
        sample.cache.budget_memory_used.to_string(),
        sample.cache.budget_objects_used.to_string(),
        sample.budget_memory_after_read.to_string(),
        sample.budget_memory_after_release.to_string(),
        sample.budget_objects_after_read.to_string(),
        sample.budget_objects_after_release.to_string(),
        sample.budget_input_bytes.to_string(),
        sample.budget_work.to_string(),
        sample.accounting.deflated_read.to_string(),
        sample.accounting.stored_read.to_string(),
        sample.accounting.deflated_produced.to_string(),
        sample.accounting.stored_accepted.to_string(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
        String::new(),
    ];
    writeln!(output, "{}", fields.join(","))
}

fn write_summary(
    output: &mut dyn io::Write,
    corpus: &Corpus,
    source: SourceKind,
    api: ApiKind,
    capability: Capability,
    workers: usize,
    summary: Summary,
) -> io::Result<()> {
    let mut fields = vec![
        "summary".to_string(),
        corpus.kind.as_str().to_string(),
        corpus.sha256.clone(),
        source.as_str().to_string(),
        api.as_str().to_string(),
        capability.as_str().to_string(),
        workers.to_string(),
    ];
    fields.extend(std::iter::repeat_n(String::new(), 30));
    fields.extend([
        summary.p50_ns.to_string(),
        summary.p95_ns.to_string(),
        summary.p99_ns.to_string(),
        summary.mean_ns.to_string(),
        summary.throughput_bytes_s.to_string(),
        summary.source_calls.to_string(),
        summary.source_requested_bytes.to_string(),
        summary.source_returned_bytes.to_string(),
        summary.source_max_request.to_string(),
        summary.cache_cold_loads.to_string(),
        summary.cache_hits.to_string(),
        summary.retained_bytes.to_string(),
        summary.budget_memory_after_release.to_string(),
        summary.budget_objects_after_release.to_string(),
    ]);
    writeln!(output, "{}", fields.join(","))
}

fn parse_args() -> Result<Config, Box<dyn std::error::Error>> {
    let mut config = Config::default();
    let mut args = std::env::args().skip(1);
    while let Some(argument) = args.next() {
        let value = |name: &str, args: &mut std::iter::Skip<std::env::Args>| {
            args.next()
                .ok_or_else(|| format!("missing value for {name}"))
        };
        match argument.as_str() {
            "--corpus" => config.corpora = parse_corpora(&value("--corpus", &mut args)?)?,
            "--source" => config.sources = parse_sources(&value("--source", &mut args)?)?,
            "--api" => config.apis = parse_apis(&value("--api", &mut args)?)?,
            "--capability" => {
                config.capabilities = parse_capabilities(&value("--capability", &mut args)?)?
            },
            "--workers" => config.workers = parse_usizes(&value("--workers", &mut args)?)?,
            "--warmups" => config.warmups = value("--warmups", &mut args)?.parse()?,
            "--samples" => config.samples = value("--samples", &mut args)?.parse()?,
            "--repeats" => config.repeats = value("--repeats", &mut args)?.parse()?,
            "--short-read" => {
                config.short_read_bytes = value("--short-read", &mut args)?.parse()?
            },
            "--delay-us" => config.delay_us = value("--delay-us", &mut args)?.parse()?,
            "--accounting" => config.accounting = true,
            "--artifact-dir" => config.artifact_dir = value("--artifact-dir", &mut args)?.into(),
            "--output" => config.output = Some(value("--output", &mut args)?.into()),
            "--keep-artifacts" => config.keep_artifacts = true,
            "--help" | "-h" => {
                print_help();
                std::process::exit(0);
            },
            unknown => return Err(format!("unknown argument {unknown}").into()),
        }
    }
    Ok(config)
}

fn parse_corpora(value: &str) -> Result<Vec<CorpusKind>, Box<dyn std::error::Error>> {
    value
        .split(',')
        .map(|item| match item {
            "few-large" => Ok(CorpusKind::FewLarge),
            "many-small" => Ok(CorpusKind::ManySmall),
            "all" => Err("all must be used alone".into()),
            other => Err(format!("unknown corpus {other}").into()),
        })
        .collect()
}

fn parse_sources(value: &str) -> Result<Vec<SourceKind>, Box<dyn std::error::Error>> {
    value
        .split(',')
        .map(|item| match item {
            "owned" => Ok(SourceKind::Owned),
            "file" => Ok(SourceKind::File),
            "instrumented" | "short-delay" => Ok(SourceKind::ShortDelay),
            "all" => Err("all must be used alone".into()),
            other => Err(format!("unknown source {other}").into()),
        })
        .collect()
}

fn parse_apis(value: &str) -> Result<Vec<ApiKind>, Box<dyn std::error::Error>> {
    value
        .split(',')
        .map(|item| match item {
            "serial" | "serial-part-data" => Ok(ApiKind::Serial),
            "batch" | "read-parts-ordered" => Ok(ApiKind::Batch),
            other => Err(format!("unknown API {other}").into()),
        })
        .collect()
}

fn parse_capabilities(value: &str) -> Result<Vec<Capability>, Box<dyn std::error::Error>> {
    value
        .split(',')
        .map(|item| match item {
            "normal" => Ok(Capability::Normal),
            "managed" | "managed-after-only" => Ok(Capability::ManagedAfterOnly),
            other => Err(format!("unknown capability {other}").into()),
        })
        .collect()
}

fn parse_usizes(value: &str) -> Result<Vec<usize>, Box<dyn std::error::Error>> {
    value
        .split(',')
        .map(|item| item.parse::<usize>().map_err(|error| error.into()))
        .collect()
}

fn print_help() {
    println!(
        "source_backed_batch_perf [options]\n\
         --corpus few-large,many-small   deterministic corpus set\n\
         --source owned,file,instrumented source adapter set\n\
         --api serial,batch                  API path to measure\n\
         --capability normal,managed     normal or explicit managed context\n\
         --workers 1,2,4,8               bounded harness worker widths\n\
         --warmups N --samples N --repeats N\n\
         --short-read BYTES --delay-us MICROSECONDS --accounting\n\
         --artifact-dir DIR --output CSV --keep-artifacts"
    );
}
