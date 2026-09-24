use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    fmt::Write as _,
    hint::black_box,
    sync::atomic::{AtomicU64, Ordering},
    time::{Duration, Instant},
};

use litchi_ods::codec::formula::FormulaParser;

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
    BracketCurrentCell,
    BracketSheetCell,
    BracketRange,
    QuotedDoubledSheet,
    LocalRefs256,
    SumUnbracketed,
    VlookupUnbracketed,
    MalformedZeroRow,
    MalformedMissingSeparator,
    MalformedUnclosedBracket,
    MalformedMissingRow,
    MalformedBadQuote,
    CoverageSourceCell,
    CoverageSourceRange,
    CoverageEmptySource,
    CoverageUnicodeEscapedSource,
    CoverageWholeColumns,
    CoverageWholeRows,
    CoverageCrossSheetRange,
    CoverageNestedInherited,
    CoverageRefError,
    CoverageColonSheet1k,
    CoverageColonSheet4k,
    CoverageColonSheet16k,
    CoverageSourceIri1k,
    CoverageSourceIri4k,
    CoverageSourceIri16k,
    CoverageSourceIriOver16k,
}

impl Case {
    fn parse(value: &str) -> Self {
        match value {
            "parse-bracket-current-cell" => Self::BracketCurrentCell,
            "parse-bracket-sheet-cell" => Self::BracketSheetCell,
            "parse-bracket-range" => Self::BracketRange,
            "parse-quoted-doubled-sheet" => Self::QuotedDoubledSheet,
            "parse-local-refs-256" => Self::LocalRefs256,
            "parse-sum-unbracketed" => Self::SumUnbracketed,
            "parse-vlookup-unbracketed" => Self::VlookupUnbracketed,
            "parse-malformed-zero-row" => Self::MalformedZeroRow,
            "parse-malformed-missing-separator" => Self::MalformedMissingSeparator,
            "parse-malformed-unclosed-bracket" => Self::MalformedUnclosedBracket,
            "parse-malformed-missing-row" => Self::MalformedMissingRow,
            "parse-malformed-bad-quote" => Self::MalformedBadQuote,
            "coverage-source-cell" => Self::CoverageSourceCell,
            "coverage-source-range" => Self::CoverageSourceRange,
            "coverage-empty-source" => Self::CoverageEmptySource,
            "coverage-unicode-escaped-source" => Self::CoverageUnicodeEscapedSource,
            "coverage-whole-columns" => Self::CoverageWholeColumns,
            "coverage-whole-rows" => Self::CoverageWholeRows,
            "coverage-cross-sheet-range" => Self::CoverageCrossSheetRange,
            "coverage-nested-inherited" => Self::CoverageNestedInherited,
            "coverage-ref-error" => Self::CoverageRefError,
            "coverage-colon-sheet-1k" => Self::CoverageColonSheet1k,
            "coverage-colon-sheet-4k" => Self::CoverageColonSheet4k,
            "coverage-colon-sheet-16k" => Self::CoverageColonSheet16k,
            "coverage-source-iri-1k" => Self::CoverageSourceIri1k,
            "coverage-source-iri-4k" => Self::CoverageSourceIri4k,
            "coverage-source-iri-16k" => Self::CoverageSourceIri16k,
            "coverage-source-iri-over-16k" => Self::CoverageSourceIriOver16k,
            other => panic!("unknown case {other:?}"),
        }
    }

    fn default_repeat(self) -> usize {
        match self {
            Self::LocalRefs256
            | Self::CoverageColonSheet16k
            | Self::CoverageSourceIri16k
            | Self::CoverageSourceIriOver16k => 128,
            _ => 1_000,
        }
    }

    fn label(self) -> &'static str {
        match self {
            Self::BracketCurrentCell => "parse-bracket-current-cell",
            Self::BracketSheetCell => "parse-bracket-sheet-cell",
            Self::BracketRange => "parse-bracket-range",
            Self::QuotedDoubledSheet => "parse-quoted-doubled-sheet",
            Self::LocalRefs256 => "parse-local-refs-256",
            Self::SumUnbracketed => "parse-sum-unbracketed",
            Self::VlookupUnbracketed => "parse-vlookup-unbracketed",
            Self::MalformedZeroRow => "parse-malformed-zero-row",
            Self::MalformedMissingSeparator => "parse-malformed-missing-separator",
            Self::MalformedUnclosedBracket => "parse-malformed-unclosed-bracket",
            Self::MalformedMissingRow => "parse-malformed-missing-row",
            Self::MalformedBadQuote => "parse-malformed-bad-quote",
            Self::CoverageSourceCell => "coverage-source-cell",
            Self::CoverageSourceRange => "coverage-source-range",
            Self::CoverageEmptySource => "coverage-empty-source",
            Self::CoverageUnicodeEscapedSource => "coverage-unicode-escaped-source",
            Self::CoverageWholeColumns => "coverage-whole-columns",
            Self::CoverageWholeRows => "coverage-whole-rows",
            Self::CoverageCrossSheetRange => "coverage-cross-sheet-range",
            Self::CoverageNestedInherited => "coverage-nested-inherited",
            Self::CoverageRefError => "coverage-ref-error",
            Self::CoverageColonSheet1k => "coverage-colon-sheet-1k",
            Self::CoverageColonSheet4k => "coverage-colon-sheet-4k",
            Self::CoverageColonSheet16k => "coverage-colon-sheet-16k",
            Self::CoverageSourceIri1k => "coverage-source-iri-1k",
            Self::CoverageSourceIri4k => "coverage-source-iri-4k",
            Self::CoverageSourceIri16k => "coverage-source-iri-16k",
            Self::CoverageSourceIriOver16k => "coverage-source-iri-over-16k",
        }
    }

    fn input(self) -> String {
        match self {
            Self::BracketCurrentCell => "of:=[.A1]".to_string(),
            Self::BracketSheetCell => "of:=[$Inputs.$A$1]".to_string(),
            Self::BracketRange => "of:=SUM([$Inputs.$A$1:.$B$2])".to_string(),
            Self::QuotedDoubledSheet => "OF:=['Bob''s'.$A$1]".to_string(),
            Self::LocalRefs256 => {
                let mut formula = String::from("of:=");
                for row in 1..=256 {
                    if row > 1 {
                        formula.push('+');
                    }
                    let _ = write!(formula, "A{row}");
                }
                formula
            },
            Self::SumUnbracketed => "of:=SUM(A1:A10)".to_string(),
            Self::VlookupUnbracketed => "of:=VLOOKUP(A1;B1:C10;2;0)".to_string(),
            Self::MalformedZeroRow => "of:=[.A0]".to_string(),
            Self::MalformedMissingSeparator => "of:=[A1]".to_string(),
            Self::MalformedUnclosedBracket => "of:=[.A1".to_string(),
            Self::MalformedMissingRow => "of:=[.A]".to_string(),
            Self::MalformedBadQuote => "OF:=['Bob's'.$A$1]".to_string(),
            Self::CoverageSourceCell => "of:=['file:///book.ods'#.A1]".to_string(),
            Self::CoverageSourceRange => "of:=['file:///book.ods'#.A1:.B2]".to_string(),
            Self::CoverageEmptySource => "of:=[''#.A1]".to_string(),
            Self::CoverageUnicodeEscapedSource => {
                "of:=['https://例え.テスト/O''Brien.ods'#.A1]".to_string()
            },
            Self::CoverageWholeColumns => "of:=[.A:.C]".to_string(),
            Self::CoverageWholeRows => "of:=[.1:.3]".to_string(),
            Self::CoverageCrossSheetRange => "of:=[Sheet1.A1:Sheet2.B2]".to_string(),
            Self::CoverageNestedInherited => "of:=[Sheet.A1.B2:.C3]".to_string(),
            Self::CoverageRefError => "of:=[#REF!]".to_string(),
            Self::CoverageColonSheet1k => colon_sheet(1_024),
            Self::CoverageColonSheet4k => colon_sheet(4_096),
            Self::CoverageColonSheet16k => colon_sheet(16_384),
            Self::CoverageSourceIri1k => source_iri(1_024),
            Self::CoverageSourceIri4k => source_iri(4_096),
            Self::CoverageSourceIri16k => source_iri(16_384),
            Self::CoverageSourceIriOver16k => source_iri(16_385),
        }
    }

    fn expected_success(self) -> bool {
        !matches!(
            self,
            Self::MalformedZeroRow
                | Self::MalformedMissingSeparator
                | Self::MalformedUnclosedBracket
                | Self::MalformedMissingRow
                | Self::MalformedBadQuote
                | Self::CoverageSourceIriOver16k
        )
    }
}

fn colon_sheet(target_bytes: usize) -> String {
    let mut sheet = String::with_capacity(target_bytes);
    sheet.push('S');
    sheet.extend(std::iter::repeat_n(':', target_bytes.saturating_sub(1)));
    format!("of:=[{sheet}.A1]")
}

fn source_iri(target_bytes: usize) -> String {
    const PREFIX: &str = "file:///";
    let mut iri = String::with_capacity(target_bytes);
    iri.push_str(PREFIX);
    iri.extend(std::iter::repeat_n('a', target_bytes.saturating_sub(PREFIX.len())));
    format!("of:=['{iri}'#.A1]")
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
                "--workload" if value == "parse" => {},
                "--workload" => panic!("workload must be parse"),
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

fn measure(input: &str, repeat: usize) -> Sample {
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0_u64;
    let mut checksum = 0_u64;
    for _ in 0..repeat {
        let input = black_box(input);
        match FormulaParser::new(black_box(input)).parse() {
            Ok(formula) => {
                successes += 1;
                checksum = checksum.wrapping_add(formula.tokens.len() as u64);
                checksum = checksum.wrapping_add(formula.text.len() as u64);
            },
            Err(error) => {
                checksum = checksum.wrapping_add(error.to_string().len() as u64);
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
    let repeat = config.repeat.unwrap_or_else(|| config.case.default_repeat());
    assert!(repeat > 0);
    println!(
        "config workload=parse case={} input_bytes={} repeat={} warmups={} iterations={} expected_success={}",
        config.case.label(),
        input.len(),
        repeat,
        config.warmups,
        config.iterations,
        config.case.expected_success(),
    );
    for _ in 0..config.warmups {
        let _ = measure(&input, repeat);
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(&input, repeat));
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
