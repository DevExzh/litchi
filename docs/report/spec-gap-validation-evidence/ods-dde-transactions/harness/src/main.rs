use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    hint::black_box,
    sync::atomic::{AtomicU64, Ordering},
    time::{Duration, Instant},
};

use litchi_ods::dde::Snapshot;

type AnyResult<T> = Result<T, Box<dyn std::error::Error>>;

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let live =
            LIVE_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed) + layout.size() as u64;
        record_peak(live);
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(size as u64, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        let old = layout.size() as u64;
        let new = size as u64;
        let live = if new >= old {
            LIVE_BYTES.fetch_add(new - old, Ordering::Relaxed) + (new - old)
        } else {
            LIVE_BYTES.fetch_sub(old - new, Ordering::Relaxed) - (old - new)
        };
        record_peak(live);
        unsafe { System.realloc(pointer, layout, size) }
    }
}

#[global_allocator]
static ALLOCATOR: CountingAllocator = CountingAllocator;

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

#[derive(Clone, Copy, Debug)]
enum Workload {
    Parse,
    Read,
}

impl Workload {
    fn parse(value: &str) -> Self {
        match value {
            "parse" => Self::Parse,
            "read" => Self::Read,
            other => panic!("unknown workload {other:?}; baseline supports parse or read"),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Case {
    Small,
    LargeCache,
    ManyLinks,
}

impl Case {
    fn parse(value: &str) -> Self {
        match value {
            "small" => Self::Small,
            "large-cache" => Self::LargeCache,
            "many-links" => Self::ManyLinks,
            other => panic!("unknown case {other:?}; expected small, large-cache, or many-links"),
        }
    }

    const fn shape(self) -> Shape {
        match self {
            Self::Small => Shape {
                sheet_sources: 1,
                links: 2,
                cache_rows: 8,
                cache_columns: 8,
            },
            Self::LargeCache => Shape {
                sheet_sources: 1,
                links: 1,
                cache_rows: 256,
                cache_columns: 256,
            },
            Self::ManyLinks => Shape {
                sheet_sources: 4,
                links: 2_048,
                cache_rows: 1,
                cache_columns: 1,
            },
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct Shape {
    sheet_sources: usize,
    links: usize,
    cache_rows: usize,
    cache_columns: usize,
}

#[derive(Clone, Copy, Debug)]
struct Config {
    workload: Workload,
    case: Case,
    warmups: usize,
    iterations: usize,
}

impl Config {
    fn from_args() -> Self {
        let mut workload = Workload::Parse;
        let mut case = Case::Small;
        let mut warmups = 3;
        let mut iterations = 15;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {arg}"));
            match arg.as_str() {
                "--workload" => workload = Workload::parse(&value),
                "--case" => case = Case::parse(&value),
                "--warmups" => warmups = value.parse().expect("warmups must be an integer"),
                "--iterations" => {
                    iterations = value.parse().expect("iterations must be an integer")
                },
                other => panic!("unknown option {other}"),
            }
        }
        assert!(iterations > 0);
        Self {
            workload,
            case,
            warmups,
            iterations,
        }
    }
}

fn content(shape: Shape) -> String {
    let estimate = 512
        + shape.sheet_sources * 260
        + shape.links * 320
        + shape.links * shape.cache_rows * (shape.cache_columns * 76 + 36);
    let mut source = String::with_capacity(estimate);
    source.push_str(&format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"{OFFICE}\" xmlns:table=\"{TABLE}\" office:version=\"1.4\"><office:body><office:spreadsheet>"
    ));
    for sheet in 0..shape.sheet_sources {
        source.push_str(&format!(
            "<table:table table:name=\"Sheet{sheet}\"><office:dde-source office:dde-application=\"soffice\" office:dde-topic=\"file:///never/contacted-{sheet}.ods\" office:dde-item=\"Sheet{sheet}.A1:B2\" office:name=\"SheetSource{sheet}\" office:conversion-mode=\"keep-text\" office:automatic-update=\"true\"/><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table>"
        ));
    }
    source.push_str("<table:dde-links>");
    for link in 0..shape.links {
        source.push_str(&format!(
            "<table:dde-link><office:dde-source office:dde-application=\"calc\" office:dde-topic=\"file:///never/opened-{link}.ods\" office:dde-item=\"Prices{link}.A1\" office:name=\"Link{link}\" office:conversion-mode=\"keep-text\" office:automatic-update=\"false\"/><table:table table:name=\"Cache{link}\"><table:table-column/>"
        ));
        for row in 0..shape.cache_rows {
            source.push_str("<table:table-row>");
            for column in 0..shape.cache_columns {
                let value = (link + row + column) % 97;
                source.push_str(&format!(
                    "<table:table-cell office:value-type=\"float\" office:value=\"{value}\"/>"
                ));
            }
            source.push_str("</table:table-row>");
        }
        source.push_str("</table:table></table:dde-link>");
    }
    source.push_str(
        "</table:dde-links></office:spreadsheet></office:body></office:document-content>",
    );
    source
}

#[derive(Clone, Copy, Debug)]
struct Sample {
    elapsed: Duration,
    alloc_calls: u64,
    dealloc_calls: u64,
    requested_bytes: u64,
    released_bytes: u64,
    live_before_bytes: u64,
    live_after_bytes: u64,
    peak_live_delta: u64,
    probes: usize,
}

fn record_peak(value: u64) {
    let mut current = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while value > current {
        match PEAK_LIVE_BYTES.compare_exchange_weak(
            current,
            value,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => current = observed,
        }
    }
}

fn reset_observer() -> u64 {
    ALLOC_CALLS.store(0, Ordering::Relaxed);
    DEALLOC_CALLS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
    let live = LIVE_BYTES.load(Ordering::Relaxed);
    PEAK_LIVE_BYTES.store(live, Ordering::Relaxed);
    live
}

fn read_snapshot(snapshot: &Snapshot) -> usize {
    let mut probes = 0usize;
    for source in snapshot.sheet_sources() {
        black_box(source.sheet().len());
        black_box(source.source().application().len());
        black_box(source.source().topic().len());
        black_box(source.source().item().len());
        black_box(source.source().name().map(str::len).unwrap_or(0));
        probes += 1;
    }
    for link in snapshot.links() {
        black_box(link.source().application().len());
        black_box(link.source().topic().len());
        black_box(link.source().item().len());
        black_box(link.source().name().map(str::len).unwrap_or(0));
        black_box(link.cached_table_xml().len());
        probes += 1;
    }
    probes
}

fn profile_once(source: &str, workload: Workload) -> AnyResult<Sample> {
    match workload {
        Workload::Parse => {
            let live_before = reset_observer();
            let started = Instant::now();
            let snapshot = Snapshot::parse(source)?;
            black_box(snapshot.source_xml().len());
            black_box(snapshot.links().len());
            black_box(snapshot.sheet_sources().len());
            let elapsed = started.elapsed();
            let live_after = LIVE_BYTES.load(Ordering::Relaxed);
            let peak = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
            Ok(Sample {
                elapsed,
                alloc_calls: ALLOC_CALLS.load(Ordering::Relaxed),
                dealloc_calls: DEALLOC_CALLS.load(Ordering::Relaxed),
                requested_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
                released_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
                live_before_bytes: live_before,
                live_after_bytes: live_after,
                peak_live_delta: peak.saturating_sub(live_before),
                probes: 0,
            })
        },
        Workload::Read => {
            let snapshot = Snapshot::parse(source)?;
            let live_before = reset_observer();
            let started = Instant::now();
            let probes = read_snapshot(&snapshot);
            let elapsed = started.elapsed();
            let live_after = LIVE_BYTES.load(Ordering::Relaxed);
            let peak = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
            Ok(Sample {
                elapsed,
                alloc_calls: ALLOC_CALLS.load(Ordering::Relaxed),
                dealloc_calls: DEALLOC_CALLS.load(Ordering::Relaxed),
                requested_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
                released_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
                live_before_bytes: live_before,
                live_after_bytes: live_after,
                peak_live_delta: peak.saturating_sub(live_before),
                probes,
            })
        },
    }
}

fn percentile(values: &mut [u128], numerator: usize, denominator: usize) -> u128 {
    values.sort_unstable();
    values[(values.len() * numerator)
        .div_ceil(denominator)
        .saturating_sub(1)]
}

fn percentile_u64(values: &mut [u64], numerator: usize, denominator: usize) -> u64 {
    values.sort_unstable();
    values[(values.len() * numerator)
        .div_ceil(denominator)
        .saturating_sub(1)]
}

fn main() -> AnyResult<()> {
    let config = Config::from_args();
    let shape = config.case.shape();
    let source = content(shape);
    println!(
        "config workload={:?} case={:?} sheet_sources={} links={} cache_rows={} cache_columns={} source_bytes={} warmups={} iterations={} work_units=unavailable memory_budget_bytes=unavailable",
        config.workload,
        config.case,
        shape.sheet_sources,
        shape.links,
        shape.cache_rows,
        shape.cache_columns,
        source.len(),
        config.warmups,
        config.iterations,
    );
    for _ in 0..config.warmups {
        let _ = profile_once(&source, config.workload)?;
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(profile_once(&source, config.workload)?);
    }
    let mut elapsed: Vec<u128> = samples
        .iter()
        .map(|sample| sample.elapsed.as_nanos())
        .collect();
    let p50 = percentile(&mut elapsed, 50, 100);
    let p95 = percentile(&mut elapsed, 95, 100);
    let p99 = percentile(&mut elapsed, 99, 100);
    let mean = samples
        .iter()
        .map(|sample| sample.elapsed.as_nanos())
        .sum::<u128>()
        / samples.len() as u128;
    let mut alloc_calls: Vec<u64> = samples.iter().map(|sample| sample.alloc_calls).collect();
    let mut requested: Vec<u64> = samples
        .iter()
        .map(|sample| sample.requested_bytes)
        .collect();
    let mut peak_live: Vec<u64> = samples
        .iter()
        .map(|sample| sample.peak_live_delta)
        .collect();
    let mut live_after: Vec<u64> = samples
        .iter()
        .map(|sample| sample.live_after_bytes)
        .collect();
    let mut dealloc_calls: Vec<u64> = samples.iter().map(|sample| sample.dealloc_calls).collect();
    let mut released: Vec<u64> = samples.iter().map(|sample| sample.released_bytes).collect();
    let mut live_before: Vec<u64> = samples
        .iter()
        .map(|sample| sample.live_before_bytes)
        .collect();
    let probes: usize = samples.iter().map(|sample| sample.probes).sum();
    let p50_u64 = |values: &mut Vec<u64>| percentile_u64(values, 50, 100);
    let max_u64 = |values: &Vec<u64>| values.iter().copied().max().unwrap_or(0);
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} alloc_calls_p50={} alloc_calls_max={} dealloc_calls_p50={} dealloc_calls_max={} requested_bytes_p50={} requested_bytes_max={} released_bytes_p50={} released_bytes_max={} live_before_p50={} live_after_p50={} live_after_max={} peak_live_delta_p50={} peak_live_delta_max={} work_delta_p50=unavailable memory_delta_p50=unavailable probes={}",
        mean,
        p50,
        p95,
        p99,
        p50_u64(&mut alloc_calls),
        max_u64(&alloc_calls),
        p50_u64(&mut dealloc_calls),
        max_u64(&dealloc_calls),
        p50_u64(&mut requested),
        max_u64(&requested),
        p50_u64(&mut released),
        max_u64(&released),
        p50_u64(&mut live_before),
        p50_u64(&mut live_after),
        max_u64(&live_after),
        p50_u64(&mut peak_live),
        max_u64(&peak_live),
        probes,
    );
    Ok(())
}
