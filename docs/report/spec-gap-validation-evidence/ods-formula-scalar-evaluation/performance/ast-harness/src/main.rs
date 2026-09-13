use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    hint::black_box,
    sync::atomic::{AtomicU64, Ordering},
    time::{Duration, Instant},
};

use litchi_ods::codec::formula::expression::{Expression, Limits};

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

#[derive(Clone, Copy, Debug)]
enum Case {
    Flat64,
    Flat256,
    Flat1024,
    Flat4096,
    Array64,
    Array256,
    Array1024,
    Array4096,
    Name64,
    Name256,
    Name1024,
    Name4096,
    String64,
    String256,
    String1024,
    String4096,
    Reference64,
    Reference256,
    Reference1024,
    Reference4096,
    DepthLimit256,
    MalformedTrailing4096,
    MalformedUnclosedArray4096,
    MalformedUnclosedString64k,
}

impl Case {
    fn parse(value: &str) -> Self {
        match value {
            "expr-flat-64" => Self::Flat64,
            "expr-flat-256" => Self::Flat256,
            "expr-flat-1024" => Self::Flat1024,
            "expr-flat-4096" => Self::Flat4096,
            "expr-array-64" => Self::Array64,
            "expr-array-256" => Self::Array256,
            "expr-array-1024" => Self::Array1024,
            "expr-array-4096" => Self::Array4096,
            "expr-name-64" => Self::Name64,
            "expr-name-256" => Self::Name256,
            "expr-name-1024" => Self::Name1024,
            "expr-name-4096" => Self::Name4096,
            "expr-string-64" => Self::String64,
            "expr-string-256" => Self::String256,
            "expr-string-1024" => Self::String1024,
            "expr-string-4096" => Self::String4096,
            "expr-reference-64" => Self::Reference64,
            "expr-reference-256" => Self::Reference256,
            "expr-reference-1024" => Self::Reference1024,
            "expr-reference-4096" => Self::Reference4096,
            "expr-depth-limit-256" => Self::DepthLimit256,
            "expr-malformed-trailing-4096" => Self::MalformedTrailing4096,
            "expr-malformed-unclosed-array-4096" => Self::MalformedUnclosedArray4096,
            "expr-malformed-unclosed-string-64k" => Self::MalformedUnclosedString64k,
            other => panic!("unknown case {other:?}"),
        }
    }

    fn label(self) -> &'static str {
        match self {
            Self::Flat64 => "expr-flat-64",
            Self::Flat256 => "expr-flat-256",
            Self::Flat1024 => "expr-flat-1024",
            Self::Flat4096 => "expr-flat-4096",
            Self::Array64 => "expr-array-64",
            Self::Array256 => "expr-array-256",
            Self::Array1024 => "expr-array-1024",
            Self::Array4096 => "expr-array-4096",
            Self::Name64 => "expr-name-64",
            Self::Name256 => "expr-name-256",
            Self::Name1024 => "expr-name-1024",
            Self::Name4096 => "expr-name-4096",
            Self::String64 => "expr-string-64",
            Self::String256 => "expr-string-256",
            Self::String1024 => "expr-string-1024",
            Self::String4096 => "expr-string-4096",
            Self::Reference64 => "expr-reference-64",
            Self::Reference256 => "expr-reference-256",
            Self::Reference1024 => "expr-reference-1024",
            Self::Reference4096 => "expr-reference-4096",
            Self::DepthLimit256 => "expr-depth-limit-256",
            Self::MalformedTrailing4096 => "expr-malformed-trailing-4096",
            Self::MalformedUnclosedArray4096 => "expr-malformed-unclosed-array-4096",
            Self::MalformedUnclosedString64k => "expr-malformed-unclosed-string-64k",
        }
    }

    fn scale(self) -> Option<usize> {
        match self {
            Self::Flat64 | Self::Array64 | Self::Name64 | Self::String64 | Self::Reference64 => {
                Some(64)
            },
            Self::Flat256
            | Self::Array256
            | Self::Name256
            | Self::String256
            | Self::Reference256 => Some(256),
            Self::Flat1024
            | Self::Array1024
            | Self::Name1024
            | Self::String1024
            | Self::Reference1024 => Some(1024),
            Self::Flat4096
            | Self::Array4096
            | Self::Name4096
            | Self::String4096
            | Self::Reference4096 => Some(4096),
            Self::DepthLimit256
            | Self::MalformedTrailing4096
            | Self::MalformedUnclosedArray4096 => Some(256),
            Self::MalformedUnclosedString64k => Some(64 * 1024),
        }
    }

    fn default_repeat(self) -> usize {
        match self.scale() {
            Some(64) => 256,
            Some(256) => 128,
            Some(1024) => 32,
            Some(4096) => 8,
            Some(65536) => 8,
            _ => 8,
        }
    }

    fn input(self) -> String {
        match self {
            Self::Flat64 => flat_chain(64),
            Self::Flat256 => flat_chain(256),
            Self::Flat1024 => flat_chain(1024),
            Self::Flat4096 => flat_chain(4096),
            Self::Array64 => array(64),
            Self::Array256 => array(256),
            Self::Array1024 => array(1024),
            Self::Array4096 => array(4096),
            Self::Name64 => name(64),
            Self::Name256 => name(256),
            Self::Name1024 => name(1024),
            Self::Name4096 => name(4096),
            Self::String64 => string_literal(64),
            Self::String256 => string_literal(256),
            Self::String1024 => string_literal(1024),
            Self::String4096 => string_literal(4096),
            Self::Reference64 => reference(64),
            Self::Reference256 => reference(256),
            Self::Reference1024 => reference(1024),
            Self::Reference4096 => reference(4096),
            Self::DepthLimit256 => nested_parentheses(256),
            Self::MalformedTrailing4096 => {
                let mut value = flat_chain(4096);
                value.push('@');
                value
            },
            Self::MalformedUnclosedArray4096 => unclosed_array(4096),
            Self::MalformedUnclosedString64k => {
                let mut value = String::with_capacity(65_538);
                value.push_str("=\"");
                value.extend(std::iter::repeat_n('a', 65_536));
                value
            },
        }
    }

    fn limits(self) -> Limits {
        match self {
            Self::DepthLimit256 => Limits::default().with_max_depth(64),
            _ => Limits::default(),
        }
    }

    fn expected_success(self) -> bool {
        !matches!(
            self,
            Self::DepthLimit256
                | Self::MalformedTrailing4096
                | Self::MalformedUnclosedArray4096
                | Self::MalformedUnclosedString64k
        )
    }
}

fn flat_chain(operands: usize) -> String {
    let mut value = String::with_capacity(operands.saturating_mul(2).saturating_add(1));
    value.push('=');
    for index in 0..operands {
        if index != 0 {
            value.push('+');
        }
        value.push('1');
    }
    value
}

fn array(cells: usize) -> String {
    let mut value = String::with_capacity(cells.saturating_mul(2).saturating_add(3));
    value.push_str("={");
    for index in 0..cells {
        if index != 0 {
            value.push(';');
        }
        value.push('1');
    }
    value.push('}');
    value
}

fn unclosed_array(cells: usize) -> String {
    let mut value = String::with_capacity(cells.saturating_mul(2).saturating_add(2));
    value.push_str("={");
    for index in 0..cells {
        if index != 0 {
            value.push(';');
        }
        value.push('1');
    }
    value
}

fn name(length: usize) -> String {
    let mut value = String::with_capacity(length.saturating_add(1));
    value.push('=');
    value.extend(std::iter::repeat_n('N', length));
    value
}

fn string_literal(length: usize) -> String {
    let mut value = String::with_capacity(length.saturating_add(3));
    value.push_str("=\"");
    value.extend(std::iter::repeat_n('a', length));
    value.push('"');
    value
}

fn reference(source_length: usize) -> String {
    const PREFIX: &str = "file:///";
    let mut source = String::with_capacity(source_length.max(PREFIX.len()));
    source.push_str(PREFIX);
    source.extend(std::iter::repeat_n(
        'a',
        source_length.saturating_sub(PREFIX.len()),
    ));
    format!("=['{source}'#.A1]")
}

fn nested_parentheses(depth: usize) -> String {
    let mut value = String::with_capacity(depth.saturating_mul(2).saturating_add(2));
    value.push('=');
    value.extend(std::iter::repeat_n('(', depth));
    value.push('1');
    value.extend(std::iter::repeat_n(')', depth));
    value
}

#[derive(Clone, Copy, Debug)]
struct Config {
    case: Case,
    warmups: usize,
    iterations: usize,
    repeat: Option<usize>,
}

impl Config {
    fn from_args() -> Self {
        let mut case = None;
        let mut warmups = 3;
        let mut iterations = 15;
        let mut repeat = None;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {arg}"));
            match arg.as_str() {
                "--workload" if value == "expression" => {},
                "--workload" => panic!("workload must be expression"),
                "--case" => case = Some(Case::parse(&value)),
                "--warmups" => warmups = value.parse().expect("warmups must be an integer"),
                "--iterations" => {
                    iterations = value.parse().expect("iterations must be an integer")
                },
                "--repeat" => repeat = Some(value.parse().expect("repeat must be an integer")),
                other => panic!("unknown option {other}"),
            }
        }
        let case = case.unwrap_or_else(|| panic!("--case is required"));
        assert!(iterations > 0);
        Self {
            case,
            warmups,
            iterations,
            repeat,
        }
    }
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
    successes: u64,
    checksum: u64,
}

fn measure(input: &str, limits: &Limits, repeat: usize) -> Sample {
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0_u64;
    let mut checksum = 0_u64;
    for _ in 0..repeat {
        match Expression::parse_with_limits(black_box(input), limits) {
            Ok(expression) => {
                let expression = black_box(expression);
                successes += 1;
                checksum = checksum.wrapping_add(expression.source().len() as u64);
                checksum = checksum.wrapping_add(expression.node_count() as u64);
                checksum = checksum.wrapping_add(expression.edge_count() as u64);
                black_box(checksum);
            },
            Err(error) => {
                black_box(error);
            },
        }
    }
    black_box((successes, checksum));
    Sample {
        elapsed: started.elapsed(),
        alloc_calls: ALLOC_CALLS.load(Ordering::Relaxed),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Relaxed),
        requested_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
        released_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
        live_before_bytes: live_before,
        live_after_bytes: LIVE_BYTES.load(Ordering::Relaxed),
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Relaxed)
            .saturating_sub(live_before),
        successes,
        checksum,
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
    let input = config.case.input();
    let limits = config.case.limits();
    let repeat = config
        .repeat
        .unwrap_or_else(|| config.case.default_repeat());
    assert!(repeat > 0);
    println!(
        "config workload=expression case={} input_bytes={} repeat={} warmups={} iterations={} expected_success={}",
        config.case.label(),
        input.len(),
        repeat,
        config.warmups,
        config.iterations,
        config.case.expected_success(),
    );
    for _ in 0..config.warmups {
        let _ = measure(&input, &limits, repeat);
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(&input, &limits, repeat));
    }
    let mut elapsed: Vec<u128> = samples
        .iter()
        .map(|sample| sample.elapsed.as_nanos())
        .collect();
    let mean = samples
        .iter()
        .map(|sample| sample.elapsed.as_nanos())
        .sum::<u128>()
        / samples.len() as u128;
    let p50 = percentile(&mut elapsed, 50, 100);
    let p95 = percentile(&mut elapsed, 95, 100);
    let p99 = percentile(&mut elapsed, 99, 100);
    let mut alloc_calls: Vec<u64> = samples.iter().map(|sample| sample.alloc_calls).collect();
    let mut dealloc_calls: Vec<u64> = samples.iter().map(|sample| sample.dealloc_calls).collect();
    let mut requested: Vec<u64> = samples
        .iter()
        .map(|sample| sample.requested_bytes)
        .collect();
    let mut released: Vec<u64> = samples.iter().map(|sample| sample.released_bytes).collect();
    let mut live_before: Vec<u64> = samples
        .iter()
        .map(|sample| sample.live_before_bytes)
        .collect();
    let mut live_after: Vec<u64> = samples
        .iter()
        .map(|sample| sample.live_after_bytes)
        .collect();
    let mut peak_live: Vec<u64> = samples
        .iter()
        .map(|sample| sample.peak_live_delta)
        .collect();
    let mut successes: Vec<u64> = samples.iter().map(|sample| sample.successes).collect();
    let mut checksums: Vec<u64> = samples.iter().map(|sample| sample.checksum).collect();
    let p50_u64 = |values: &mut Vec<u64>| percentile_u64(values, 50, 100);
    let max_u64 = |values: &Vec<u64>| values.iter().copied().max().unwrap_or(0);
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} alloc_calls_p50={} alloc_calls_max={} dealloc_calls_p50={} dealloc_calls_max={} requested_bytes_p50={} requested_bytes_max={} released_bytes_p50={} released_bytes_max={} live_before_p50={} live_after_p50={} live_after_max={} peak_live_delta_p50={} peak_live_delta_max={} successes_p50={} successes_max={} checksum_p50={} checksum_max={}",
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
        p50_u64(&mut successes),
        max_u64(&successes),
        p50_u64(&mut checksums),
        max_u64(&checksums),
    );
    Ok(())
}
