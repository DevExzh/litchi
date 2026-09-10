//! Profile the XLSB DrawingML theme owner.
//!
//! Each invocation measures one lane in one fresh process and emits one JSON
//! report. The `coder_api` module is the only adapter to the XLSB theme API;
//! the process isolation, allocator accounting, digest gates, and
//! machine-readable report are deliberately independent of that API shape.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    clippy::similar_names,
    reason = "this opt-in profile owns a process allocator wrapper and emits machine-readable stdout"
)]

use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_drawingml::theme::{Color, Palette, Slot, Theme, codec};
use litchi_xlsb::theme;
use litchi_xlsb::{SourceBackedWorkbook, Workbook};
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::io;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

type BoxError = Box<dyn Error + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_FIXTURE: &str = "test-data/poi/test-data/spreadsheet/testVarious.xlsb";
const DEFAULT_WARMUP: usize = 3;
const DEFAULT_SAMPLES: usize = 30;

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

// SAFETY: each method forwards the allocator contract unchanged to System;
// counters only observe successful operations and never affect ownership.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair belongs to the caller.
        unsafe { System.dealloc(pointer, layout) };
        ALLOC_DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_DEALLOCATED_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            ALLOC_REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOC_ALLOCATED_BYTES.fetch_add(as_u64(new_size), Ordering::Relaxed);
            ALLOC_DEALLOCATED_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            replace_live(layout.size(), new_size);
        }
        result
    }
}

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static ALLOC_ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_DEALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_INVALID: AtomicBool = AtomicBool::new(false);

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn observe_alloc(size: usize) {
    let size = as_u64(size);
    ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
    ALLOC_ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
    let live = ALLOC_LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    update_peak(live);
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = ALLOC_LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        ALLOC_INVALID.store(true, Ordering::Release);
    }
}

fn replace_live(old_size: usize, new_size: usize) {
    subtract_live(old_size);
    let size = as_u64(new_size);
    let live = ALLOC_LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    update_peak(live);
}

fn update_peak(live: u64) {
    let mut peak = ALLOC_PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while live > peak {
        match ALLOC_PEAK_LIVE_BYTES.compare_exchange_weak(
            peak,
            live,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => peak = observed,
        }
    }
}

fn reset_peak_to_live() {
    ALLOC_PEAK_LIVE_BYTES.store(ALLOC_LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
}

#[derive(Clone, Copy)]
struct AllocationSnapshot {
    calls: u64,
    dealloc_calls: u64,
    realloc_calls: u64,
    failed: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    invalid: bool,
}

impl AllocationSnapshot {
    fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: ALLOC_DEALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: ALLOC_REALLOC_CALLS.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            allocated_bytes: ALLOC_ALLOCATED_BYTES.load(Ordering::Acquire),
            deallocated_bytes: ALLOC_DEALLOCATED_BYTES.load(Ordering::Acquire),
            live_bytes: ALLOC_LIVE_BYTES.load(Ordering::Acquire),
            peak_live_bytes: ALLOC_PEAK_LIVE_BYTES.load(Ordering::Acquire),
            invalid: ALLOC_INVALID.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> AllocationDelta {
        AllocationDelta {
            calls: after.calls.saturating_sub(self.calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            failed: after.failed.saturating_sub(self.failed),
            allocated_bytes: after.allocated_bytes.saturating_sub(self.allocated_bytes),
            deallocated_bytes: after
                .deallocated_bytes
                .saturating_sub(self.deallocated_bytes),
            live_before: self.live_bytes,
            live_after: after.live_bytes,
            peak_before: self.peak_live_bytes,
            peak_after: after.peak_live_bytes,
            invalid: self.invalid || after.invalid,
        }
    }
}

#[derive(Clone, Copy)]
struct AllocationDelta {
    calls: u64,
    dealloc_calls: u64,
    realloc_calls: u64,
    failed: u64,
    allocated_bytes: u64,
    deallocated_bytes: u64,
    live_before: u64,
    live_after: u64,
    peak_before: u64,
    peak_after: u64,
    invalid: bool,
}

#[derive(Clone, Copy, Default)]
struct ReadCounters {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

/// Source-reader instrumentation is kept independent from the report's plain
/// counter value. The source-backed adapter can pass the returned `Arc` to
/// `SourceBackedWorkbook::from_read_at` and snapshot it around the timed call.
/// This remains deliberately small so the adapter is the only code that needs
/// to change when the owner API's source constructor settles.
#[allow(dead_code)]
#[derive(Default)]
struct ReadCounterState {
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
}

#[allow(dead_code)]
struct CountingReadAt {
    inner: OwnedSource,
    counters: Arc<ReadCounterState>,
}

#[allow(dead_code)]
impl CountingReadAt {
    fn new(bytes: Vec<u8>, counters: Arc<ReadCounterState>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            counters,
        }
    }
}

#[allow(dead_code)]
impl ReadCounterState {
    fn snapshot(&self) -> ReadCounters {
        ReadCounters {
            calls: self.calls.load(Ordering::Acquire),
            requested_bytes: self.requested_bytes.load(Ordering::Acquire),
            returned_bytes: self.returned_bytes.load(Ordering::Acquire),
        }
    }
}

#[allow(dead_code)]
impl ReadCounters {
    fn delta(self, after: Self) -> Self {
        Self {
            calls: after.calls.saturating_sub(self.calls),
            requested_bytes: after.requested_bytes.saturating_sub(self.requested_bytes),
            returned_bytes: after.returned_bytes.saturating_sub(self.returned_bytes),
        }
    }

    fn add(self, other: Self) -> Self {
        Self {
            calls: self.calls.saturating_add(other.calls),
            requested_bytes: self.requested_bytes.saturating_add(other.requested_bytes),
            returned_bytes: self.returned_bytes.saturating_add(other.returned_bytes),
        }
    }
}

#[allow(dead_code)]
impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.calls.fetch_add(1, Ordering::Relaxed);
        self.counters
            .requested_bytes
            .fetch_add(as_u64(output.len()), Ordering::Relaxed);
        let read = self.inner.read_at(offset, output)?;
        self.counters
            .returned_bytes
            .fetch_add(as_u64(read), Ordering::Relaxed);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

/// A compact semantic observation that is stable across typed owner
/// implementations. The digest is over the complete typed theme re-encoding;
/// the metadata digest exercises the repeated lookup path without allocating
/// strings for every sample.
#[derive(Clone, Copy)]
struct ThemeObservation {
    semantic_digest: u64,
    metadata_digest: u64,
}

struct Fixture {
    path: PathBuf,
    bytes: Vec<u8>,
    theme_bytes: Vec<u8>,
    expected_observation: ThemeObservation,
    input_digest: u64,
    theme_digest: u64,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Case {
    SourceCold,
    EagerCold,
    SourceWarm,
    EagerWarm,
    Noop,
    ChangeInverse,
}

impl Case {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "source_cold" => Ok(Self::SourceCold),
            "eager_cold" => Ok(Self::EagerCold),
            "source_warm" => Ok(Self::SourceWarm),
            "eager_warm" => Ok(Self::EagerWarm),
            "noop" => Ok(Self::Noop),
            "change_inverse" => Ok(Self::ChangeInverse),
            _ => Err(format!(
                "unknown case {value:?}; expected source_cold, eager_cold, source_warm, eager_warm, noop, or change_inverse"
            )
            .into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::SourceCold => "source_cold",
            Self::EagerCold => "eager_cold",
            Self::SourceWarm => "source_warm",
            Self::EagerWarm => "eager_warm",
            Self::Noop => "noop",
            Self::ChangeInverse => "change_inverse",
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::SourceCold => {
                "source-backed workbook open + typed theme read + source-view construction"
            },
            Self::EagerCold => {
                "owned XLSB package open + typed theme read + typed snapshot construction"
            },
            Self::SourceWarm => "repeated typed theme metadata lookups on one source-backed handle",
            Self::EagerWarm => "repeated typed theme metadata lookups on one owned package handle",
            Self::Noop => "typed theme transaction no-op commit and source-checked publication",
            Self::ChangeInverse => {
                "typed theme color change, readback, inverse publication, and exact restoration"
            },
        }
    }
}

struct Args {
    case: Case,
    fixture: PathBuf,
    warmup: usize,
    samples: usize,
}

impl Args {
    fn parse() -> Result<Self> {
        let mut case = None;
        let mut fixture = PathBuf::from(DEFAULT_FIXTURE);
        let mut warmup = DEFAULT_WARMUP;
        let mut samples = DEFAULT_SAMPLES;
        let mut arguments = env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--case" => case = Some(Case::parse(&next_arg(&mut arguments, "--case")?)?),
                "--fixture" => fixture = PathBuf::from(next_arg(&mut arguments, "--fixture")?),
                "--warmup" => {
                    warmup = parse_positive(&next_arg(&mut arguments, "--warmup")?, "--warmup")?
                },
                "--samples" => {
                    samples = parse_positive(&next_arg(&mut arguments, "--samples")?, "--samples")?
                },
                "--help" | "-h" => return Err(Usage.into()),
                other => return Err(format!("unknown argument {other:?}\n\n{Usage}").into()),
            }
        }
        Ok(Self {
            case: case.ok_or_else(|| format!("--case is required\n\n{Usage}"))?,
            fixture,
            warmup,
            samples,
        })
    }
}

fn next_arg<I>(arguments: &mut I, name: &str) -> Result<String>
where
    I: Iterator<Item = String>,
{
    arguments
        .next()
        .ok_or_else(|| format!("missing value for {name}").into())
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(parsed)
}

struct Usage;

impl std::fmt::Display for Usage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(
            "usage: theme_profile --case <source_cold|eager_cold|source_warm|eager_warm|noop|change_inverse> [--fixture PATH] [--warmup N] [--samples N]",
        )
    }
}

impl std::fmt::Debug for Usage {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Display::fmt(self, formatter)
    }
}

impl Error for Usage {}

fn main() -> Result<()> {
    let args = Args::parse()?;
    let fixture = Fixture::load(&args.fixture)?;
    let report = match args.case {
        Case::SourceCold | Case::EagerCold => {
            run_cold(&fixture, args.case, args.warmup, args.samples)?
        },
        Case::SourceWarm | Case::EagerWarm => {
            run_warm(&fixture, args.case, args.warmup, args.samples)?
        },
        Case::Noop | Case::ChangeInverse => {
            run_transaction(&fixture, args.case, args.warmup, args.samples)?
        },
    };
    println!("{}", report.to_json());
    Ok(())
}

impl Fixture {
    fn load(path: &Path) -> Result<Self> {
        let bytes = fs::read(path)?;
        let workbook = Workbook::new(Cursor::new(bytes.clone()))?;
        let snapshot = workbook.theme()?.ok_or("workbook has no Theme part")?;
        let theme_bytes = snapshot.source_xml().to_vec();
        let theme = snapshot.theme().clone();
        let expected_observation = observe(&theme)?;
        Ok(Self {
            path: path.to_path_buf(),
            input_digest: digest_bytes(0xcbf29ce484222325, &bytes),
            theme_digest: digest_bytes(0xcbf29ce484222325, &theme_bytes),
            bytes,
            theme_bytes,
            expected_observation,
        })
    }
}

struct Prepared {
    snapshot: Option<theme::Snapshot>,
    source_view: Option<theme::View>,
    source_digest: u64,
    semantic_digest: u64,
    setup_reads: ReadCounters,
    read_counters: Option<Arc<ReadCounterState>>,
}

impl Prepared {
    fn theme(&self) -> Result<&Theme> {
        self.source_view
            .as_ref()
            .map(theme::View::theme)
            .or_else(|| self.snapshot.as_ref().map(theme::Snapshot::theme))
            .ok_or_else(|| "prepared theme payload is missing".into())
    }

    fn read_snapshot(&self) -> ReadCounters {
        self.read_counters
            .as_ref()
            .map_or_else(ReadCounters::default, |counters| counters.snapshot())
    }
}

fn run_cold(fixture: &Fixture, case: Case, warmup: usize, samples: usize) -> Result<Report> {
    let mut values = Vec::with_capacity(samples);
    for _ in 0..warmup {
        let _ = run_cold_sample(fixture, case)?;
    }
    for _ in 0..samples {
        values.push(run_cold_sample(fixture, case)?);
    }
    Report::new(fixture, case, warmup, values)
}

fn run_cold_sample(fixture: &Fixture, case: Case) -> Result<Sample> {
    reset_peak_to_live();
    let allocation_before = AllocationSnapshot::now();
    let started = Instant::now();
    let (prepared, source_digest, reads, preservation_ok, inverse_ok) = match case {
        Case::SourceCold => {
            let mut counters = ReadCounters::default();
            let (prepared, observed_reads) = coder_api::source_cold(&fixture.bytes, &mut counters)?;
            let source_digest = prepared.source_digest;
            (prepared, source_digest, observed_reads, true, true)
        },
        Case::EagerCold => {
            let prepared = coder_api::eager_cold(&fixture.bytes)?;
            let source_digest = prepared.source_digest;
            (prepared, source_digest, ReadCounters::default(), true, true)
        },
        _ => return Err("invalid cold case".into()),
    };
    let elapsed_ns = started.elapsed().as_nanos();
    let allocation = allocation_before.delta(AllocationSnapshot::now());
    let observation = observe(prepared.theme()?)?;
    Ok(Sample {
        elapsed_ns: u64::try_from(elapsed_ns).unwrap_or(u64::MAX),
        observation,
        source_digest,
        reads,
        reads_before: ReadCounters::default(),
        reads_total: reads,
        allocation,
        preservation_ok,
        inverse_ok,
        change_ok: true,
    })
}

fn run_warm(fixture: &Fixture, case: Case, warmup: usize, samples: usize) -> Result<Report> {
    reset_peak_to_live();
    let setup_allocation_before = AllocationSnapshot::now();
    let mut prepared = match case {
        Case::SourceWarm => coder_api::source_prepare(&fixture.bytes)?,
        Case::EagerWarm => coder_api::eager_prepare(&fixture.bytes)?,
        _ => return Err("invalid warm case".into()),
    };
    let setup_allocation = setup_allocation_before.delta(AllocationSnapshot::now());
    prepared.semantic_digest = observe(prepared.theme()?)?.semantic_digest;
    let mut values = Vec::with_capacity(samples);
    for _ in 0..warmup {
        let _ = run_warm_sample(&prepared, case)?;
    }
    for _ in 0..samples {
        values.push(run_warm_sample(&prepared, case)?);
    }
    Report::with_setup(
        fixture,
        case,
        warmup,
        values,
        prepared.setup_reads,
        setup_allocation,
    )
}

fn run_warm_sample(prepared: &Prepared, case: Case) -> Result<Sample> {
    reset_peak_to_live();
    let allocation_before = AllocationSnapshot::now();
    let started = Instant::now();
    let (observation, reads_before, reads) = match case {
        Case::SourceWarm => coder_api::source_query(prepared)?,
        Case::EagerWarm => coder_api::eager_query(prepared)?,
        _ => return Err("invalid warm case".into()),
    };
    let elapsed_ns = started.elapsed().as_nanos();
    let allocation = allocation_before.delta(AllocationSnapshot::now());
    let reads_total = reads_before.add(reads);
    Ok(Sample {
        elapsed_ns: u64::try_from(elapsed_ns).unwrap_or(u64::MAX),
        observation,
        source_digest: prepared.source_digest,
        reads,
        reads_before,
        reads_total,
        allocation,
        preservation_ok: true,
        inverse_ok: true,
        change_ok: true,
    })
}

fn run_transaction(fixture: &Fixture, case: Case, warmup: usize, samples: usize) -> Result<Report> {
    if case != Case::Noop && case != Case::ChangeInverse {
        return Err("invalid transaction case".into());
    }
    let mut values = Vec::with_capacity(samples);
    for _ in 0..warmup {
        let _ = run_transaction_sample(fixture, case)?;
    }
    for _ in 0..samples {
        values.push(run_transaction_sample(fixture, case)?);
    }
    Report::new(fixture, case, warmup, values)
}

fn run_transaction_sample(fixture: &Fixture, case: Case) -> Result<Sample> {
    reset_peak_to_live();
    let allocation_before = AllocationSnapshot::now();
    let started = Instant::now();
    let result = match case {
        Case::Noop => coder_api::typed_noop(&fixture.bytes)?,
        Case::ChangeInverse => coder_api::typed_change_inverse(&fixture.bytes)?,
        _ => return Err("invalid transaction case".into()),
    };
    let elapsed_ns = started.elapsed().as_nanos();
    let allocation = allocation_before.delta(AllocationSnapshot::now());
    Ok(Sample {
        elapsed_ns: u64::try_from(elapsed_ns).unwrap_or(u64::MAX),
        observation: result.observation,
        source_digest: result.source_digest,
        reads: result.reads,
        reads_before: ReadCounters::default(),
        reads_total: result.reads,
        allocation,
        preservation_ok: result.preservation_ok,
        inverse_ok: result.inverse_ok,
        change_ok: result.change_ok,
    })
}

fn observe(theme: &Theme) -> Result<ThemeObservation> {
    let encoded = codec::encode_part(&theme.name, &theme.colors, &theme.fonts)?;
    Ok(ThemeObservation {
        semantic_digest: digest_bytes(0xcbf29ce484222325, &encoded),
        metadata_digest: metadata_digest(theme),
    })
}

fn metadata_digest(theme: &Theme) -> u64 {
    let mut metadata_digest = 0xcbf29ce484222325;
    metadata_digest = digest_bytes(metadata_digest, theme.name.as_bytes());
    metadata_digest = digest_bytes(metadata_digest, theme.colors.name().as_bytes());
    metadata_digest = digest_bytes(metadata_digest, theme.fonts.name().as_bytes());
    metadata_digest = digest_bytes(metadata_digest, theme.fonts.major().latin.as_bytes());
    metadata_digest = digest_bytes(metadata_digest, theme.fonts.minor().latin.as_bytes());
    if let Some(color) = theme.colors.color(Slot::Accent1) {
        match color {
            Color::Rgb(value) => metadata_digest = digest_bytes(metadata_digest, value.as_bytes()),
            Color::System { kind, last } => {
                metadata_digest = digest_bytes(metadata_digest, kind.token().as_bytes());
                if let Some(value) = last {
                    metadata_digest = digest_bytes(metadata_digest, value.as_bytes());
                }
            },
        }
    }
    metadata_digest
}

fn digest_bytes(mut digest: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        digest ^= u64::from(*byte);
        digest = digest.wrapping_mul(0x100000001b3);
    }
    digest
}

// ---------------------------------------------------------------------------
// The coder-facing seam. Keep this section small: production API names and
// source-backed handle types are intentionally isolated here while the rest of
// the profiler remains frozen and reviewable.
// ---------------------------------------------------------------------------

mod coder_api {
    use super::*;

    pub(super) struct TransactionResult {
        pub(super) observation: ThemeObservation,
        pub(super) source_digest: u64,
        pub(super) reads: ReadCounters,
        pub(super) preservation_ok: bool,
        pub(super) inverse_ok: bool,
        pub(super) change_ok: bool,
    }

    pub(super) fn eager_cold(bytes: &[u8]) -> Result<Prepared> {
        eager_prepare(bytes)
    }

    pub(super) fn eager_prepare(bytes: &[u8]) -> Result<Prepared> {
        let workbook = Workbook::new(Cursor::new(bytes.to_vec()))?;
        let snapshot = workbook
            .theme()?
            .ok_or("eager workbook has no Theme part")?;
        let theme_xml = snapshot.source_xml();
        let source_digest = digest_bytes(0xcbf29ce484222325, theme_xml);
        Ok(Prepared {
            snapshot: Some(snapshot),
            source_view: None,
            source_digest,
            semantic_digest: 0,
            setup_reads: ReadCounters::default(),
            read_counters: None,
        })
    }

    pub(super) fn eager_query(
        prepared: &Prepared,
    ) -> Result<(ThemeObservation, ReadCounters, ReadCounters)> {
        let reads_before = prepared.read_snapshot();
        let theme = prepared.theme()?;
        Ok((
            ThemeObservation {
                semantic_digest: prepared.semantic_digest,
                metadata_digest: metadata_digest(theme),
            },
            reads_before,
            ReadCounters::default(),
        ))
    }

    pub(super) fn source_cold(
        bytes: &[u8],
        _counters: &mut ReadCounters,
    ) -> Result<(Prepared, ReadCounters)> {
        let counters = Arc::new(ReadCounterState::default());
        let source: Arc<dyn ReadAt> =
            Arc::new(CountingReadAt::new(bytes.to_vec(), Arc::clone(&counters)));
        let workbook = SourceBackedWorkbook::from_read_at(source)?;
        let view = workbook
            .theme()?
            .ok_or("source-backed workbook has no Theme part")?;
        let source_digest = digest_bytes(0xcbf29ce484222325, view.source_xml());
        let setup_reads = counters.snapshot();
        let prepared = Prepared {
            source_digest,
            semantic_digest: 0,
            snapshot: None,
            source_view: Some(view),
            setup_reads,
            read_counters: Some(counters),
        };
        Ok((prepared, setup_reads))
    }

    pub(super) fn source_prepare(bytes: &[u8]) -> Result<Prepared> {
        let counters = Arc::new(ReadCounterState::default());
        let source: Arc<dyn ReadAt> =
            Arc::new(CountingReadAt::new(bytes.to_vec(), Arc::clone(&counters)));
        let workbook = SourceBackedWorkbook::from_read_at(source)?;
        let view = workbook
            .theme()?
            .ok_or("source-backed workbook has no Theme part")?;
        let source_digest = digest_bytes(0xcbf29ce484222325, view.source_xml());
        let setup_reads = counters.snapshot();
        Ok(Prepared {
            source_digest,
            semantic_digest: 0,
            snapshot: None,
            source_view: Some(view),
            setup_reads,
            read_counters: Some(counters),
        })
    }

    pub(super) fn source_query(
        prepared: &Prepared,
    ) -> Result<(ThemeObservation, ReadCounters, ReadCounters)> {
        let reads_before = prepared.read_snapshot();
        let theme = prepared.theme()?;
        let observation = ThemeObservation {
            semantic_digest: prepared.semantic_digest,
            metadata_digest: metadata_digest(theme),
        };
        let reads_after = prepared.read_snapshot();
        Ok((observation, reads_before, reads_before.delta(reads_after)))
    }

    /// Typed Theme no-op transaction and exact source-preservation gate.
    pub(super) fn typed_noop(bytes: &[u8]) -> Result<TransactionResult> {
        let mut workbook = Workbook::new(Cursor::new(bytes.to_vec()))?;
        let before = workbook.theme()?.ok_or("workbook has no Theme part")?;
        let transaction = workbook.edit_theme()?;
        let commit = transaction.commit()?;
        let after = workbook.apply_theme(&commit)?;
        let observation = observe(after.theme())?;
        let source_digest = digest_bytes(0xcbf29ce484222325, after.source_xml());
        Ok(TransactionResult {
            observation,
            source_digest,
            reads: ReadCounters::default(),
            preservation_ok: after.source_xml() == before.source_xml(),
            inverse_ok: true,
            change_ok: !commit.changed(),
        })
    }

    pub(super) fn typed_change_inverse(bytes: &[u8]) -> Result<TransactionResult> {
        let mut workbook = Workbook::new(Cursor::new(bytes.to_vec()))?;
        let before = workbook.theme()?.ok_or("workbook has no Theme part")?;
        let original_theme = before.theme().clone();
        let mut transaction = workbook.edit_theme()?;
        let changed = change_theme(&original_theme)?;
        let staged_changed = transaction.replace(changed)?;
        let commit = transaction.commit()?;
        if !staged_changed || !commit.changed() {
            return Err("typed Theme change unexpectedly produced a no-op commit".into());
        }
        let changed_snapshot = workbook.apply_theme(&commit)?;
        let changed_observation = observe(changed_snapshot.theme())?;
        let inverse = commit.patch().inverse();
        let restored = workbook.apply_theme_patch(&inverse)?;
        let restored_theme_xml = restored.source_xml();
        let source_digest = digest_bytes(0xcbf29ce484222325, restored_theme_xml);
        Ok(TransactionResult {
            observation: changed_observation,
            source_digest,
            reads: ReadCounters::default(),
            preservation_ok: restored_theme_xml == before.source_xml(),
            inverse_ok: *restored.theme() == original_theme,
            change_ok: changed_observation.semantic_digest
                != observe(&original_theme)?.semantic_digest,
        })
    }

    fn change_theme(theme: &Theme) -> Result<Theme> {
        let mut colors = Palette::new(theme.colors.name());
        for slot in Slot::ALL {
            let color = theme
                .colors
                .color(slot)
                .ok_or("theme palette is missing a required slot")?
                .clone();
            colors = colors.with(slot, color);
        }
        colors = colors.with(Slot::Accent1, Color::rgb("010203")?);
        Ok(Theme {
            name: theme.name.clone(),
            colors,
            fonts: theme.fonts.clone(),
        })
    }
}

struct Sample {
    elapsed_ns: u64,
    observation: ThemeObservation,
    source_digest: u64,
    reads: ReadCounters,
    reads_before: ReadCounters,
    reads_total: ReadCounters,
    allocation: AllocationDelta,
    preservation_ok: bool,
    inverse_ok: bool,
    change_ok: bool,
}

struct Report {
    fixture: PathBuf,
    case: Case,
    input_bytes: usize,
    theme_bytes: usize,
    input_digest: u64,
    theme_digest: u64,
    expected_semantic_digest: u64,
    expected_metadata_digest: u64,
    warmup: usize,
    samples: Vec<Sample>,
    setup_reads: Option<ReadCounters>,
    setup_allocation: Option<AllocationDelta>,
}

impl Report {
    fn new(fixture: &Fixture, case: Case, warmup: usize, samples: Vec<Sample>) -> Result<Self> {
        if samples.is_empty() {
            return Err("profile requires at least one sample".into());
        }
        Ok(Self {
            fixture: fixture.path.clone(),
            case,
            input_bytes: fixture.bytes.len(),
            theme_bytes: fixture.theme_bytes.len(),
            input_digest: fixture.input_digest,
            theme_digest: fixture.theme_digest,
            expected_semantic_digest: fixture.expected_observation.semantic_digest,
            expected_metadata_digest: fixture.expected_observation.metadata_digest,
            warmup,
            samples,
            setup_reads: None,
            setup_allocation: None,
        })
    }

    fn with_setup(
        fixture: &Fixture,
        case: Case,
        warmup: usize,
        samples: Vec<Sample>,
        setup_reads: ReadCounters,
        setup_allocation: AllocationDelta,
    ) -> Result<Self> {
        let mut report = Self::new(fixture, case, warmup, samples)?;
        report.setup_reads = Some(setup_reads);
        report.setup_allocation = Some(setup_allocation);
        Ok(report)
    }

    fn to_json(&self) -> String {
        let mut output = String::new();
        output.push('{');
        json_str(&mut output, "schema", "xlsb-theme-profile-v1");
        json_str(&mut output, "case", self.case.name());
        json_str(&mut output, "timing_scope", self.case.timing_scope());
        json_str(&mut output, "fixture", &self.fixture.to_string_lossy());
        json_num(&mut output, "input_bytes", self.input_bytes as u64);
        json_num(&mut output, "theme_bytes", self.theme_bytes as u64);
        json_num(&mut output, "input_digest", self.input_digest);
        json_num(&mut output, "theme_digest", self.theme_digest);
        json_num(
            &mut output,
            "expected_semantic_digest",
            self.expected_semantic_digest,
        );
        json_num(
            &mut output,
            "expected_metadata_digest",
            self.expected_metadata_digest,
        );
        json_num(&mut output, "warmup", self.warmup as u64);
        json_num(&mut output, "sample_count", self.samples.len() as u64);
        if let Some(reads) = self.setup_reads {
            output.push_str(",\"setup_reads\":{");
            json_num(&mut output, "calls", reads.calls);
            json_num(&mut output, "requested_bytes", reads.requested_bytes);
            json_num(&mut output, "returned_bytes", reads.returned_bytes);
            output.push('}');
        }
        if let Some(allocation) = self.setup_allocation {
            output.push_str(",\"setup_allocation\":{");
            json_num(&mut output, "calls", allocation.calls);
            json_num(&mut output, "dealloc_calls", allocation.dealloc_calls);
            json_num(&mut output, "realloc_calls", allocation.realloc_calls);
            json_num(&mut output, "failed", allocation.failed);
            json_num(&mut output, "allocated_bytes", allocation.allocated_bytes);
            json_num(
                &mut output,
                "deallocated_bytes",
                allocation.deallocated_bytes,
            );
            json_num(&mut output, "live_before", allocation.live_before);
            json_num(&mut output, "live_after", allocation.live_after);
            json_num(&mut output, "peak_before", allocation.peak_before);
            json_num(&mut output, "peak_after", allocation.peak_after);
            json_bool(&mut output, "invalid", allocation.invalid);
            output.push('}');
        }
        json_bool(&mut output, "semantic_ok", self.semantic_ok());
        json_bool(&mut output, "digest_stable", self.digest_stable());
        json_bool(&mut output, "preservation_ok", self.preservation_ok());
        json_bool(&mut output, "inverse_ok", self.inverse_ok());
        json_bool(&mut output, "change_ok", self.change_ok());
        json_bool(
            &mut output,
            "source_observation_stable",
            self.source_observation_stable(),
        );
        json_num(&mut output, "p50_ns", percentile(&self.samples, 50));
        json_num(&mut output, "p95_ns", percentile(&self.samples, 95));
        json_num(&mut output, "p99_ns", percentile(&self.samples, 99));
        let mean = self
            .samples
            .iter()
            .map(|sample| u128::from(sample.elapsed_ns))
            .sum::<u128>()
            / self.samples.len() as u128;
        json_num(
            &mut output,
            "mean_ns",
            u64::try_from(mean).unwrap_or(u64::MAX),
        );
        output.push_str(",\"samples\":[");
        for (index, sample) in self.samples.iter().enumerate() {
            if index > 0 {
                output.push(',');
            }
            output.push('{');
            json_num(&mut output, "elapsed_ns", sample.elapsed_ns);
            json_num(
                &mut output,
                "semantic_digest",
                sample.observation.semantic_digest,
            );
            json_num(
                &mut output,
                "metadata_digest",
                sample.observation.metadata_digest,
            );
            json_num(&mut output, "source_digest", sample.source_digest);
            json_bool(&mut output, "preservation_ok", sample.preservation_ok);
            json_bool(&mut output, "inverse_ok", sample.inverse_ok);
            json_bool(&mut output, "change_ok", sample.change_ok);
            output.push_str(",\"reads_before\":{");
            json_num(&mut output, "calls", sample.reads_before.calls);
            json_num(
                &mut output,
                "requested_bytes",
                sample.reads_before.requested_bytes,
            );
            json_num(
                &mut output,
                "returned_bytes",
                sample.reads_before.returned_bytes,
            );
            output.push('}');
            output.push_str(",\"reads\":{");
            json_num(&mut output, "calls", sample.reads.calls);
            json_num(&mut output, "requested_bytes", sample.reads.requested_bytes);
            json_num(&mut output, "returned_bytes", sample.reads.returned_bytes);
            output.push('}');
            output.push_str(",\"reads_total\":{");
            json_num(&mut output, "calls", sample.reads_total.calls);
            json_num(
                &mut output,
                "requested_bytes",
                sample.reads_total.requested_bytes,
            );
            json_num(
                &mut output,
                "returned_bytes",
                sample.reads_total.returned_bytes,
            );
            output.push('}');
            output.push_str(",\"allocation\":{");
            json_num(&mut output, "calls", sample.allocation.calls);
            json_num(
                &mut output,
                "dealloc_calls",
                sample.allocation.dealloc_calls,
            );
            json_num(
                &mut output,
                "realloc_calls",
                sample.allocation.realloc_calls,
            );
            json_num(&mut output, "failed", sample.allocation.failed);
            json_num(
                &mut output,
                "allocated_bytes",
                sample.allocation.allocated_bytes,
            );
            json_num(
                &mut output,
                "deallocated_bytes",
                sample.allocation.deallocated_bytes,
            );
            json_num(&mut output, "live_before", sample.allocation.live_before);
            json_num(&mut output, "live_after", sample.allocation.live_after);
            json_num(&mut output, "peak_before", sample.allocation.peak_before);
            json_num(&mut output, "peak_after", sample.allocation.peak_after);
            json_bool(&mut output, "invalid", sample.allocation.invalid);
            output.push('}');
            output.push('}');
        }
        output.push_str("]}");
        output
    }

    fn semantic_ok(&self) -> bool {
        self.samples.iter().all(|sample| {
            let observed = sample.observation.semantic_digest;
            observed != 0
                && match self.case {
                    Case::ChangeInverse => observed != self.expected_semantic_digest,
                    _ => observed == self.expected_semantic_digest,
                }
        })
    }

    fn digest_stable(&self) -> bool {
        self.samples.iter().all(|sample| {
            sample.source_digest == self.theme_digest
                && sample.observation.metadata_digest != 0
                && match self.case {
                    Case::ChangeInverse => sample.observation.semantic_digest != 0,
                    _ => sample.observation.metadata_digest == self.expected_metadata_digest,
                }
        }) && self.samples.windows(2).all(|pair| {
            pair[0].observation.semantic_digest == pair[1].observation.semantic_digest
                && pair[0].observation.metadata_digest == pair[1].observation.metadata_digest
                && pair[0].source_digest == pair[1].source_digest
        })
    }

    fn preservation_ok(&self) -> bool {
        self.samples.iter().all(|sample| sample.preservation_ok)
    }

    fn inverse_ok(&self) -> bool {
        self.samples.iter().all(|sample| sample.inverse_ok)
    }

    fn change_ok(&self) -> bool {
        self.samples.iter().all(|sample| sample.change_ok)
    }

    fn source_observation_stable(&self) -> bool {
        self.samples.windows(2).all(|pair| {
            pair[0].reads.calls == pair[1].reads.calls
                && pair[0].reads.requested_bytes == pair[1].reads.requested_bytes
                && pair[0].reads.returned_bytes == pair[1].reads.returned_bytes
                && pair[0].reads_before.calls == pair[1].reads_before.calls
                && pair[0].reads_before.requested_bytes == pair[1].reads_before.requested_bytes
                && pair[0].reads_before.returned_bytes == pair[1].reads_before.returned_bytes
                && pair[0].reads_total.calls == pair[1].reads_total.calls
                && pair[0].reads_total.requested_bytes == pair[1].reads_total.requested_bytes
                && pair[0].reads_total.returned_bytes == pair[1].reads_total.returned_bytes
        })
    }
}

fn percentile(samples: &[Sample], percentile: u64) -> u64 {
    let mut values = samples
        .iter()
        .map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    values.sort_unstable();
    let rank = (percentile
        .saturating_mul(values.len() as u64)
        .saturating_add(99))
        / 100;
    values[rank.saturating_sub(1) as usize]
}

fn json_str(output: &mut String, key: &str, value: &str) {
    json_separator(output);
    write!(output, "\"{key}\":\"").expect("String cannot fail");
    for character in value.chars() {
        match character {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            character if character.is_control() => {
                write!(output, "\\u{:04x}", character as u32).expect("String cannot fail")
            },
            character => output.push(character),
        }
    }
    output.push('"');
}

fn json_num(output: &mut String, key: &str, value: u64) {
    json_separator(output);
    write!(output, "\"{key}\":{value}").expect("String cannot fail");
}

fn json_bool(output: &mut String, key: &str, value: bool) {
    json_separator(output);
    write!(output, "\"{key}\":{value}").expect("String cannot fail");
}

fn json_separator(output: &mut String) {
    if !matches!(output.as_bytes().last(), Some(b'{') | Some(b'[')) {
        output.push(',');
    }
}
