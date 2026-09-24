use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    fmt::Write as _,
    hint::black_box,
    sync::atomic::{AtomicU64, Ordering},
    time::{Duration, Instant},
};

use litchi_ods::codec::formula::{is_valid_function, FormulaParser};

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
enum Workload {
    Lookup,
    Parse,
}

impl Workload {
    fn parse(value: &str) -> Self {
        match value {
            "lookup" => Self::Lookup,
            "parse" => Self::Parse,
            other => panic!("unknown workload {other:?}; expected lookup or parse"),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Case {
    LookupSumUpper,
    LookupSumLower,
    LookupSumMixed,
    LookupVlookupUpper,
    LookupVlookupLower,
    LookupVlookupMixed,
    LookupInvalidShort,
    LookupInvalidAscii4k,
    LookupInvalidAscii64k,
    LookupInvalidUnicode4k,
    LookupInvalidUnicode64k,
    LookupNewName,
    ParseSumUpper,
    ParseSumLower,
    ParseSumMixed,
    ParseVlookupUpper,
    ParseVlookupLower,
    ParseVlookupMixed,
    ParseAbsolute,
    ParseInvalidAscii4k,
    ParseInvalidAscii64k,
    ParseInvalidUnicode4k,
    ParseInvalidUnicode64k,
    ParseCellHeavy,
    ParseNewName,
}

impl Case {
    fn parse(value: &str) -> Self {
        match value {
            "lookup-sum-upper" => Self::LookupSumUpper,
            "lookup-sum-lower" => Self::LookupSumLower,
            "lookup-sum-mixed" => Self::LookupSumMixed,
            "lookup-vlookup-upper" => Self::LookupVlookupUpper,
            "lookup-vlookup-lower" => Self::LookupVlookupLower,
            "lookup-vlookup-mixed" => Self::LookupVlookupMixed,
            "lookup-invalid-short" => Self::LookupInvalidShort,
            "lookup-invalid-ascii-4k" => Self::LookupInvalidAscii4k,
            "lookup-invalid-ascii-64k" => Self::LookupInvalidAscii64k,
            "lookup-invalid-unicode-4k" => Self::LookupInvalidUnicode4k,
            "lookup-invalid-unicode-64k" => Self::LookupInvalidUnicode64k,
            "lookup-new-name" => Self::LookupNewName,
            "parse-sum-upper" => Self::ParseSumUpper,
            "parse-sum-lower" => Self::ParseSumLower,
            "parse-sum-mixed" => Self::ParseSumMixed,
            "parse-vlookup-upper" => Self::ParseVlookupUpper,
            "parse-vlookup-lower" => Self::ParseVlookupLower,
            "parse-vlookup-mixed" => Self::ParseVlookupMixed,
            "parse-absolute" => Self::ParseAbsolute,
            "parse-invalid-ascii-4k" => Self::ParseInvalidAscii4k,
            "parse-invalid-ascii-64k" => Self::ParseInvalidAscii64k,
            "parse-invalid-unicode-4k" => Self::ParseInvalidUnicode4k,
            "parse-invalid-unicode-64k" => Self::ParseInvalidUnicode64k,
            "parse-cell-heavy" => Self::ParseCellHeavy,
            "parse-new-name" => Self::ParseNewName,
            other => panic!("unknown case {other:?}"),
        }
    }

    fn default_repeat(self) -> usize {
        match self {
            Self::LookupInvalidAscii4k
            | Self::LookupInvalidUnicode4k
            | Self::ParseInvalidAscii4k
            | Self::ParseInvalidUnicode4k => 256,
            Self::LookupInvalidAscii64k | Self::LookupInvalidUnicode64k => 16,
            Self::ParseInvalidAscii64k | Self::ParseInvalidUnicode64k => 4,
            Self::LookupSumUpper
            | Self::LookupSumLower
            | Self::LookupSumMixed
            | Self::LookupVlookupUpper
            | Self::LookupVlookupLower
            | Self::LookupVlookupMixed
            | Self::LookupInvalidShort
            | Self::LookupNewName => 10_000,
            Self::ParseSumUpper
            | Self::ParseSumLower
            | Self::ParseSumMixed
            | Self::ParseVlookupUpper
            | Self::ParseVlookupLower
            | Self::ParseVlookupMixed
            | Self::ParseAbsolute
            | Self::ParseCellHeavy
            | Self::ParseNewName => 1_000,
        }
    }

    fn is_lookup(self) -> bool {
        matches!(
            self,
            Self::LookupSumUpper
                | Self::LookupSumLower
                | Self::LookupSumMixed
                | Self::LookupVlookupUpper
                | Self::LookupVlookupLower
                | Self::LookupVlookupMixed
                | Self::LookupInvalidShort
                | Self::LookupInvalidAscii4k
                | Self::LookupInvalidAscii64k
                | Self::LookupInvalidUnicode4k
                | Self::LookupInvalidUnicode64k
                | Self::LookupNewName
        )
    }

    fn label(self) -> &'static str {
        match self {
            Self::LookupSumUpper => "lookup-sum-upper",
            Self::LookupSumLower => "lookup-sum-lower",
            Self::LookupSumMixed => "lookup-sum-mixed",
            Self::LookupVlookupUpper => "lookup-vlookup-upper",
            Self::LookupVlookupLower => "lookup-vlookup-lower",
            Self::LookupVlookupMixed => "lookup-vlookup-mixed",
            Self::LookupInvalidShort => "lookup-invalid-short",
            Self::LookupInvalidAscii4k => "lookup-invalid-ascii-4k",
            Self::LookupInvalidAscii64k => "lookup-invalid-ascii-64k",
            Self::LookupInvalidUnicode4k => "lookup-invalid-unicode-4k",
            Self::LookupInvalidUnicode64k => "lookup-invalid-unicode-64k",
            Self::LookupNewName => "lookup-new-name",
            Self::ParseSumUpper => "parse-sum-upper",
            Self::ParseSumLower => "parse-sum-lower",
            Self::ParseSumMixed => "parse-sum-mixed",
            Self::ParseVlookupUpper => "parse-vlookup-upper",
            Self::ParseVlookupLower => "parse-vlookup-lower",
            Self::ParseVlookupMixed => "parse-vlookup-mixed",
            Self::ParseAbsolute => "parse-absolute",
            Self::ParseInvalidAscii4k => "parse-invalid-ascii-4k",
            Self::ParseInvalidAscii64k => "parse-invalid-ascii-64k",
            Self::ParseInvalidUnicode4k => "parse-invalid-unicode-4k",
            Self::ParseInvalidUnicode64k => "parse-invalid-unicode-64k",
            Self::ParseCellHeavy => "parse-cell-heavy",
            Self::ParseNewName => "parse-new-name",
        }
    }

    fn input(self, name_override: Option<&str>, formula_override: Option<&str>) -> String {
        if self.is_lookup() {
            if let Some(name) = name_override {
                return name.to_string();
            }
            return match self {
                Self::LookupSumUpper => "SUM".to_string(),
                Self::LookupSumLower => "sum".to_string(),
                Self::LookupSumMixed => "SuM".to_string(),
                Self::LookupVlookupUpper => "VLOOKUP".to_string(),
                Self::LookupVlookupLower => "vlookup".to_string(),
                Self::LookupVlookupMixed => "vLoOkUp".to_string(),
                Self::LookupInvalidShort => "INVALID_FUNCTION".to_string(),
                Self::LookupInvalidAscii4k => invalid_name(4_096, "A"),
                Self::LookupInvalidAscii64k => invalid_name(65_536, "A"),
                Self::LookupInvalidUnicode4k => invalid_name(4_096, "é"),
                Self::LookupInvalidUnicode64k => invalid_name(65_536, "é"),
                Self::LookupNewName => "BITAND".to_string(),
                _ => unreachable!(),
            };
        }
        if let Some(formula) = formula_override {
            return formula.to_string();
        }
        match self {
            Self::ParseSumUpper => "of:=SUM(A1:A10)".to_string(),
            Self::ParseSumLower => "of:=sum(A1:A10)".to_string(),
            Self::ParseSumMixed => "of:=SuM(A1:A10)".to_string(),
            Self::ParseVlookupUpper => "of:=VLOOKUP(A1;B1:C10;2;0)".to_string(),
            Self::ParseVlookupLower => "of:=vlookup(A1;B1:C10;2;0)".to_string(),
            Self::ParseVlookupMixed => "of:=vLoOkUp(A1;B1:C10;2;0)".to_string(),
            Self::ParseAbsolute => "of:=SUM([$Inputs.$A$1:.$B$2])".to_string(),
            Self::ParseInvalidAscii4k => invalid_formula(4_096, "A"),
            Self::ParseInvalidAscii64k => invalid_formula(65_536, "A"),
            Self::ParseInvalidUnicode4k => invalid_formula(4_096, "é"),
            Self::ParseInvalidUnicode64k => invalid_formula(65_536, "é"),
            Self::ParseCellHeavy => {
                let mut formula = String::from("of:=");
                for row in 1..=256 {
                    if row > 1 {
                        formula.push('+');
                    }
                    let _ = write!(formula, "A{row}");
                }
                formula
            },
            Self::ParseNewName => {
                let name = name_override.unwrap_or("BITAND");
                format!("of:={name}(A1)")
            },
            _ => unreachable!(),
        }
    }
}

fn invalid_name(target_bytes: usize, unit: &str) -> String {
    let mut name = String::from("INVALID_FUNCTION_");
    while name.len() + unit.len() <= target_bytes {
        name.push_str(unit);
    }
    name
}

fn invalid_formula(target_bytes: usize, unit: &str) -> String {
    let mut formula = String::from("of:=INVALID_FUNCTION_");
    while formula.len() + unit.len() + 4 <= target_bytes {
        formula.push_str(unit);
    }
    formula.push_str("(A1)");
    formula
}

#[derive(Clone, Copy, Debug)]
struct Config {
    workload: Workload,
    case: Case,
    warmups: usize,
    iterations: usize,
    repeat: Option<usize>,
}

impl Config {
    fn from_args() -> Self {
        let mut workload = None;
        let mut case = None;
        let mut warmups = 3;
        let mut iterations = 15;
        let mut repeat = None;
        let mut name_override = None;
        let mut formula_override = None;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {arg}"));
            match arg.as_str() {
                "--workload" => workload = Some(Workload::parse(&value)),
                "--case" => case = Some(Case::parse(&value)),
                "--warmups" => warmups = value.parse().expect("warmups must be an integer"),
                "--iterations" => {
                    iterations = value.parse().expect("iterations must be an integer")
                },
                "--repeat" => repeat = Some(value.parse().expect("repeat must be an integer")),
                "--name" => name_override = Some(value),
                "--formula" => formula_override = Some(value),
                other => panic!("unknown option {other}"),
            }
        }
        let workload = workload.unwrap_or(Workload::Lookup);
        let case = case.unwrap_or_else(|| panic!("--case is required"));
        if matches!(workload, Workload::Lookup) != case.is_lookup() {
            panic!("workload does not match case {}", case.label());
        }
        assert!(iterations > 0);
        let _ = (name_override, formula_override);
        Self {
            workload,
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

fn measure(input: &str, workload: Workload, repeat: usize) -> AnyResult<Sample> {
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0u64;
    let mut checksum = 0u64;
    for _ in 0..repeat {
        let input = black_box(input);
        match workload {
            Workload::Lookup => {
                let valid = is_valid_function(black_box(input));
                successes += u64::from(valid);
                checksum = checksum.wrapping_add(u64::from(valid));
            },
            Workload::Parse => match FormulaParser::new(black_box(input)).parse() {
                Ok(formula) => {
                    successes += 1;
                    checksum = checksum.wrapping_add(formula.tokens.len() as u64);
                    checksum = checksum.wrapping_add(formula.text.len() as u64);
                },
                Err(error) => {
                    checksum = checksum.wrapping_add(error.to_string().len() as u64);
                },
            },
        }
    }
    black_box((successes, checksum));
    Ok(Sample {
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
    })
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
    let name_override = env::args().skip(1).collect::<Vec<_>>();
    let name = name_override
        .windows(2)
        .find(|window| window[0] == "--name")
        .map(|window| window[1].as_str());
    let formula = name_override
        .windows(2)
        .find(|window| window[0] == "--formula")
        .map(|window| window[1].as_str());
    let input = config.case.input(name, formula);
    let repeat = config.repeat.unwrap_or_else(|| config.case.default_repeat());
    assert!(repeat > 0);
    println!(
        "config workload={:?} case={} input_bytes={} repeat={} warmups={} iterations={}",
        config.workload,
        config.case.label(),
        input.len(),
        repeat,
        config.warmups,
        config.iterations,
    );
    for _ in 0..config.warmups {
        let _ = measure(&input, config.workload, repeat)?;
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(&input, config.workload, repeat)?);
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
    let mut released: Vec<u64> = samples
        .iter()
        .map(|sample| sample.released_bytes)
        .collect();
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
