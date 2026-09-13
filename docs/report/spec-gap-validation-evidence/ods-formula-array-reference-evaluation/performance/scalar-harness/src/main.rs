use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
    error::Error,
    hint::black_box,
    num::{NonZeroU64, NonZeroUsize},
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits as BudgetLimits, Resource,
};
use litchi_ods::codec::formula::{
    evaluation::{
        EvaluationContext, EvaluationFailure, EvaluationLimits, ScalarError, ScalarValue,
        UnsupportedKind, evaluate_scalar,
    },
    expression::Expression,
};

type AnyResult<T> = Result<T, Box<dyn Error>>;

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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Revision {
    Baseline,
    Candidate,
}

impl Revision {
    fn parse(value: &str) -> Self {
        match value {
            "baseline" => Self::Baseline,
            "candidate" => Self::Candidate,
            other => panic!("unknown revision {other:?}"),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Baseline => "baseline",
            Self::Candidate => "candidate",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Phase {
    Parse,
    Evaluate,
    ParseEvaluate,
}

impl Phase {
    fn parse(value: &str) -> Self {
        match value {
            "parse" => Self::Parse,
            "evaluate" => Self::Evaluate,
            "parse-evaluate" => Self::ParseEvaluate,
            other => panic!("unknown phase {other:?}"),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Parse => "parse",
            Self::Evaluate => "evaluate",
            Self::ParseEvaluate => "parse-evaluate",
        }
    }

    const fn evaluates(self) -> bool {
        !matches!(self, Self::Parse)
    }
}

fn scale(case: &str) -> Option<usize> {
    for value in [4096, 1024, 256, 64] {
        if case.ends_with(&format!("-{value}")) {
            return Some(value);
        }
    }
    None
}

fn default_repeat(case: &str) -> usize {
    match scale(case) {
        Some(64) => 64,
        Some(256) => 32,
        Some(1024) => 8,
        Some(4096) => 2,
        Some(_) => 2,
        None => 128,
    }
}

fn expected_success(case: &str, phase: Phase) -> bool {
    if !phase.evaluates() {
        return true;
    }
    if case.starts_with("failure-") {
        return false;
    }
    if case.starts_with("roman-refusal-") {
        return false;
    }
    true
}

enum ExpectedValue {
    Number(f64),
    Logical(bool),
    Text(String),
    FormulaError(ScalarError),
}

fn roman_format_case(case: &str) -> Option<(usize, usize)> {
    let mut parts = case.split('-');
    if parts.next() != Some("roman") {
        return None;
    }
    let number = parts.next()?.parse().ok()?;
    if parts.next() != Some("format") {
        return None;
    }
    let format = parts.next()?.parse().ok()?;
    (format <= 4).then_some((number, format))
}

fn roman_expected(number: usize, format: usize) -> String {
    // These vectors follow the ODF 1.4 §6.19.17 format table. In particular,
    // format 4 uses the shortest valid signed-residue spelling, rather than
    // importing an Excel-specific ROMAN convention.
    match (number, format) {
        (3888, 0) => "MMMDCCCLXXXVIII".to_owned(),
        (3888, 1) => "MMMDCCCLXXXVIII".to_owned(),
        (3888, 2) => "MMMDCCCLXXXVIII".to_owned(),
        (3888, 3) => "MMMDCCCLXXXVIII".to_owned(),
        (3888, 4) => "IIXCMMMM".to_owned(),
        (499, 0) => "CDXCIX".to_owned(),
        (499, 1) => "LDVLIV".to_owned(),
        (499, 2) => "ID".to_owned(),
        (499, 3) => "ID".to_owned(),
        (499, 4) => "ID".to_owned(),
        (998, 0) => "CMXCVIII".to_owned(),
        (998, 1) => "CMXCVIII".to_owned(),
        (998, 2) => "XMVIII".to_owned(),
        (998, 3) => "VMIII".to_owned(),
        (998, 4) => "IIM".to_owned(),
        _ => panic!("missing reviewed ROMAN vector for {number} format {format}"),
    }
}

fn expected_value(case: &str) -> Option<ExpectedValue> {
    if let Some((number, format)) = roman_format_case(case) {
        return Some(ExpectedValue::Text(roman_expected(number, format)));
    }
    if let Some(value) = scale(case) {
        if case.starts_with("control-flat-") {
            return Some(ExpectedValue::Number(value as f64));
        }
        if case.starts_with("control-coerce-") {
            return Some(ExpectedValue::Number(value as f64 + 1.0));
        }
        if case.starts_with("control-utf8-left-") {
            return Some(ExpectedValue::Text(format!("1{}", utf8_text(value))));
        }
        if case.starts_with("control-escaped-") {
            return Some(ExpectedValue::Text(escaped_decoded(value)));
        }
        if case.starts_with("logical-and-")
            || case.starts_with("logical-or-")
            || case.starts_with("logical-not-")
            || case.starts_with("logical-if-")
            || case.starts_with("logical-iferror-")
            || case.starts_with("logical-ifna-")
        {
            return Some(ExpectedValue::Logical(true));
        }
        if case.starts_with("logical-xor-") {
            return Some(ExpectedValue::Logical(false));
        }
        if case.starts_with("lazy-if-true-heavy-") {
            return Some(ExpectedValue::Logical(true));
        }
        if case.starts_with("bitwise-and-") {
            return Some(ExpectedValue::Number(value as f64 * 2.0));
        }
        if case.starts_with("bitwise-or-") {
            return Some(ExpectedValue::Number(value as f64 * 7.0));
        }
        if case.starts_with("bitwise-xor-") || case.starts_with("bitwise-rshift-") {
            return Some(ExpectedValue::Number(value as f64 * 5.0));
        }
        if case.starts_with("bitwise-lshift-") {
            return Some(ExpectedValue::Number(value as f64 * 20.0));
        }
        if case.starts_with("arabic-input-") {
            return Some(ExpectedValue::Number(value as f64 * 1000.0));
        }
        if case.starts_with("roman-concat-") {
            return Some(ExpectedValue::Text("I".repeat(value)));
        }
        if case.starts_with("radix-base-input-") {
            return Some(ExpectedValue::Text("FF".repeat(value)));
        }
        if case.starts_with("radix-decimal-input-") {
            return Some(ExpectedValue::Number(value as f64 * 255.0));
        }
        if case.starts_with("radix-dec2hex-padding-") {
            return Some(ExpectedValue::Text("00000FF".repeat(value)));
        }
        if case.starts_with("radix-base-padding-") {
            return Some(ExpectedValue::Text(format!("{}FF", "0".repeat(value - 2))));
        }
    }
    match case {
        "control-true" | "logical-true" => Some(ExpectedValue::Logical(true)),
        "control-false" | "logical-false" => Some(ExpectedValue::Logical(false)),
        "lazy-if-false-reference"
        | "lazy-iferror-success-reference"
        | "lazy-ifna-success-reference"
        | "lazy-iferror-error" => Some(ExpectedValue::Logical(true)),
        "lazy-ifna-non-na" => Some(ExpectedValue::FormulaError(ScalarError::DivisionByZero)),
        "text-if-true" | "text-if-false" | "text-iferror" | "text-ifna" => {
            Some(ExpectedValue::Text("selected".to_owned()))
        },
        "text-if-escaped" => Some(ExpectedValue::Text("é\"".to_owned())),
        "text-if-concat" => Some(ExpectedValue::Text("étail".to_owned())),
        "bitwise-coerce-text"
        | "bitwise-coerce-fraction"
        | "bitwise-shift-negative"
        | "bitwise-lazy" => Some(ExpectedValue::Number(2.0)),
        "bitwise-lazy-error" | "bitwise-lazy-selected" => Some(ExpectedValue::Number(7.0)),
        "bitwise-error-negative" | "bitwise-error-overflow" | "bitwise-error-shift" => {
            Some(ExpectedValue::FormulaError(ScalarError::Number))
        },
        "bitwise-error-numeric" | "bitwise-error-arity" => {
            Some(ExpectedValue::FormulaError(ScalarError::Value))
        },
        "radix-base" | "radix-dec2hex" | "radix-base-small" => {
            Some(ExpectedValue::Text("FF".to_owned()))
        },
        "radix-bin2hex" | "radix-oct2hex" => Some(ExpectedValue::Text("A".to_owned())),
        "radix-bin2dec" => Some(ExpectedValue::Number(10.0)),
        "radix-bin2oct" => Some(ExpectedValue::Text("12".to_owned())),
        "radix-dec2bin" | "radix-hex2bin" | "radix-oct2bin" => {
            Some(ExpectedValue::Text("1010".to_owned()))
        },
        "radix-dec2oct" => Some(ExpectedValue::Text("77".to_owned())),
        "radix-hex2oct" => Some(ExpectedValue::Text("12".to_owned())),
        "radix-decimal" | "radix-hex2dec" | "radix-decimal-small" => {
            Some(ExpectedValue::Number(255.0))
        },
        "radix-oct2dec" => Some(ExpectedValue::Number(63.0)),
        "radix-decimal-space"
        | "radix-decimal-tab"
        | "radix-decimal-prefix"
        | "radix-decimal-h" => Some(ExpectedValue::Number(255.0)),
        "radix-decimal-b" => Some(ExpectedValue::Number(10.0)),
        "radix-base-truncate" => Some(ExpectedValue::Text("F".to_owned())),
        "radix-error-fraction" | "radix-fraction-error" => {
            Some(ExpectedValue::FormulaError(ScalarError::Value))
        },
        "radix-error-radix-low" | "radix-error-radix-high" | "radix-error-large-dec2bin" | "radix-error-large-dec2hex" => Some(ExpectedValue::FormulaError(ScalarError::Number)),
        "radix-error-invalid-digit"
        | "radix-error-empty"
        | "radix-error-arity"
        | "radix-error-text-number" => Some(ExpectedValue::FormulaError(ScalarError::Value)),
        "radix-big-base" => {
            Some(ExpectedValue::Text("1FFFFFFFFFFFFF".to_owned()))
        },
        "radix-max-base" | "radix-base-max" => Some(ExpectedValue::Text("FFFFFFFFFFFFF800000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000".to_owned())),
        "radix-max-decimal" | "radix-decimal-max" => Some(ExpectedValue::Number(f64::MAX)),
        "radix-big-decimal" => Some(ExpectedValue::Number(9007199254740991.0)),
        "radix-direct-negative" => Some(ExpectedValue::Text("FFFFFFFFFF".to_owned())),
        "arabic-uppercase" => Some(ExpectedValue::Number(3888.0)),
        "arabic-lowercase" => Some(ExpectedValue::Number(3888.0)),
        "arabic-indirect" => Some(ExpectedValue::Number(3888.0)),
        "roman-zero" => Some(ExpectedValue::Text(String::new())),
        "roman-truncate" => Some(ExpectedValue::Text(roman_expected(3888, 0))),
        "roman-format-logical-true" => Some(ExpectedValue::Text(roman_expected(499, 0))),
        "roman-format-logical-false" => Some(ExpectedValue::Text(roman_expected(499, 4))),
        "arabic-empty" => Some(ExpectedValue::Number(0.0)),
        "roman-bound-max" => Some(ExpectedValue::Text("MMMCMXCIX".to_owned())),
        "roman-error-fraction" | "roman-error-text" | "arabic-error-invalid" => {
            Some(ExpectedValue::FormulaError(ScalarError::Value))
        },
        "roman-error-arity" => Some(ExpectedValue::FormulaError(ScalarError::Value)),
        "roman-error-low" | "roman-error-high" => {
            Some(ExpectedValue::FormulaError(ScalarError::Number))
        },
        _ => None,
    }
}

fn expected_failure(case: &str) -> Option<&'static str> {
    match case {
        "failure-reference" => Some("unsupported-reference"),
        "failure-array" => Some("unsupported-array"),
        "failure-name" => Some("unsupported-name"),
        "failure-work" => Some("resource-work"),
        "failure-memory" => Some("resource-memory"),
        "failure-stack" => Some("resource-objects"),
        "failure-cancelled" => Some("cancelled"),
        "radix-refusal-work" => Some("resource-work"),
        "radix-refusal-memory" => Some("resource-memory"),
        "radix-refusal-stack" => Some("resource-objects"),
        "radix-refusal-cancelled" => Some("cancelled"),
        "roman-refusal-work" => Some("resource-work"),
        "roman-refusal-memory" => Some("resource-memory"),
        "roman-refusal-stack" => Some("resource-objects"),
        "roman-refusal-cancelled" => Some("cancelled"),
        _ => None,
    }
}

fn input(case: &str) -> String {
    if let Some((number, format)) = roman_format_case(case) {
        return format!("=ROMAN({number};{format})");
    }
    match case {
        "radix-base-small" => return input("radix-base"),
        "radix-decimal-small" => return input("radix-decimal"),
        "radix-base-max" => return input("radix-max-base"),
        "radix-decimal-max" => return input("radix-max-decimal"),
        "radix-fraction-error" => return input("radix-error-fraction"),
        _ => {},
    }
    if let Some(value) = scale(case) {
        if case.starts_with("arabic-input-") {
            return format!("=ARABIC(\"{}\")", "M".repeat(value));
        }
        if case.starts_with("roman-concat-") {
            return radix_chain("ROMAN(1)", value, '&');
        }

        if case.starts_with("control-flat-") {
            return flat_chain(value);
        }
        if case.starts_with("control-coerce-") {
            return numeric_coercion(value);
        }
        if case.starts_with("control-utf8-left-") {
            return utf8_number_left(value);
        }
        if case.starts_with("control-escaped-") {
            return escaped_literal(&escaped_body(value));
        }
        if case.starts_with("logical-and-") {
            return variadic("AND", value, |_| "TRUE()".to_owned());
        }
        if case.starts_with("logical-or-") {
            return variadic("OR", value, |index| {
                if index + 1 == value {
                    "TRUE()".to_owned()
                } else {
                    "FALSE()".to_owned()
                }
            });
        }
        if case.starts_with("logical-xor-") {
            return variadic("XOR", value, |index| {
                if index % 2 == 0 {
                    "TRUE()".to_owned()
                } else {
                    "FALSE()".to_owned()
                }
            });
        }
        if case.starts_with("logical-not-") {
            return variadic("AND", value, |_| "NOT(FALSE())".to_owned());
        }
        if case.starts_with("logical-if-") {
            return variadic("AND", value, |_| "IF(TRUE();TRUE();FALSE())".to_owned());
        }
        if case.starts_with("logical-iferror-") {
            return variadic("AND", value, |_| "IFERROR(TRUE();FALSE())".to_owned());
        }
        if case.starts_with("logical-ifna-") {
            return variadic("AND", value, |_| "IFNA(TRUE();FALSE())".to_owned());
        }
        if case.starts_with("radix-base-input-") {
            return radix_chain("BASE(255;16)", value, '&');
        }
        if case.starts_with("radix-decimal-input-") {
            return radix_chain("DECIMAL(\"FF\";16)", value, '+');
        }
        if case.starts_with("radix-dec2hex-padding-") {
            return radix_chain("DEC2HEX(255;7)", value, '&');
        }
        if case.starts_with("radix-base-padding-") {
            return format!("=BASE(255;16;{value})");
        }
        if case.starts_with("lazy-if-true-heavy-") {
            let heavy = "x".repeat(value);
            return format!("=IF(TRUE();TRUE();\"{heavy}\")");
        }
        if case.starts_with("bitwise-") {
            let name = if case.starts_with("bitwise-and-") {
                "BITAND"
            } else if case.starts_with("bitwise-or-") {
                "BITOR"
            } else if case.starts_with("bitwise-xor-") {
                "BITXOR"
            } else if case.starts_with("bitwise-lshift-") {
                "BITLSHIFT"
            } else if case.starts_with("bitwise-rshift-") {
                "BITRSHIFT"
            } else {
                panic!("unknown scaled bitwise case {case:?}")
            };
            return bitwise_chain(name, value);
        }
    }
    match case {
        "radix-direct-negative" => "=DEC2HEX(-1)".to_owned(),
        "arabic-uppercase" => "=ARABIC(\"MMMDCCCLXXXVIII\")".to_owned(),
        "arabic-lowercase" => "=ARABIC(\"mmmdccclxxxviii\")".to_owned(),
        "arabic-indirect" => "=ARABIC(ROMAN(3888;4))".to_owned(),
        "roman-zero" => "=ROMAN(0)".to_owned(),
        "roman-truncate" => "=ROMAN(3888.9)".to_owned(),
        "roman-format-logical-true" => "=ROMAN(499;TRUE())".to_owned(),
        "roman-format-logical-false" => "=ROMAN(499;FALSE())".to_owned(),
        "arabic-empty" => "=ARABIC(\"\")".to_owned(),
        "roman-bound-max" => "=ROMAN(3999)".to_owned(),
        "roman-error-text" => "=ROMAN(\"bad\")".to_owned(),
        "roman-error-arity" => "=ROMAN(1;0;2)".to_owned(),
        "arabic-error-invalid" => "=ARABIC(\"A\")".to_owned(),
        "roman-error-low" => "=ROMAN(-1)".to_owned(),
        "roman-error-high" => "=ROMAN(4000)".to_owned(),
        "roman-refusal-work" => "=ROMAN(3888;4)".to_owned(),
        "roman-refusal-memory" => "=ROMAN(3888;4)".to_owned(),
        "roman-refusal-stack" => "=ROMAN(3888;4)".to_owned(),
        "roman-refusal-cancelled" => "=ROMAN(3888;4)".to_owned(),
        "control-true" | "logical-true" => "=TRUE()".to_owned(),
        "control-false" | "logical-false" => "=FALSE()".to_owned(),
        "lazy-if-false-reference" => "=IF(FALSE();[.A1];TRUE())".to_owned(),
        "lazy-iferror-success-reference" => "=IFERROR(TRUE();[.A1])".to_owned(),
        "lazy-ifna-success-reference" => "=IFNA(TRUE();[.A1])".to_owned(),
        "lazy-iferror-error" => "=IFERROR(#N/A;TRUE())".to_owned(),
        "lazy-ifna-non-na" => "=IFNA(#DIV/0!;TRUE())".to_owned(),
        "text-if-true" => "=IF(TRUE();\"selected\";\"unused\")".to_owned(),
        "text-if-false" => "=IF(FALSE();\"unused\";\"selected\")".to_owned(),
        "text-if-escaped" => "=IF(TRUE();\"é\"\"\";\"unused\")".to_owned(),
        "text-if-concat" => "=IF(TRUE();\"é\";\"x\")&\"tail\"".to_owned(),
        "text-iferror" => "=IFERROR(\"selected\";\"fallback\")".to_owned(),
        "text-ifna" => "=IFNA(\"selected\";\"fallback\")".to_owned(),
        "radix-base" => "=BASE(255;16)".to_owned(),
        "radix-bin2dec" => "=BIN2DEC(\"1010\")".to_owned(),
        "radix-bin2hex" => "=BIN2HEX(\"1010\")".to_owned(),
        "radix-bin2oct" => "=BIN2OCT(\"1010\")".to_owned(),
        "radix-dec2bin" => "=DEC2BIN(10)".to_owned(),
        "radix-dec2hex" => "=DEC2HEX(255)".to_owned(),
        "radix-dec2oct" => "=DEC2OCT(63)".to_owned(),
        "radix-decimal" => "=DECIMAL(\"FF\";16)".to_owned(),
        "radix-hex2bin" => "=HEX2BIN(\"A\")".to_owned(),
        "radix-hex2dec" => "=HEX2DEC(\"FF\")".to_owned(),
        "radix-hex2oct" => "=HEX2OCT(\"A\")".to_owned(),
        "radix-oct2bin" => "=OCT2BIN(\"12\")".to_owned(),
        "radix-oct2dec" => "=OCT2DEC(\"77\")".to_owned(),
        "radix-oct2hex" => "=OCT2HEX(\"12\")".to_owned(),
        "radix-decimal-space" => "=DECIMAL(\"  FF\";16)".to_owned(),
        "radix-decimal-tab" => "=DECIMAL(\"\tFF\";16)".to_owned(),
        "radix-decimal-prefix" => "=DECIMAL(\"0xFF\";16)".to_owned(),
        "radix-decimal-h" => "=DECIMAL(\"FFH\";16)".to_owned(),
        "radix-decimal-b" => "=DECIMAL(\"1010B\";2)".to_owned(),
        "radix-base-truncate" => "=BASE(15.9;16)".to_owned(),
        "radix-error-fraction" => "=DEC2HEX(15.9)".to_owned(),
        "radix-error-invalid-digit" => "=DECIMAL(\"1G\";16)".to_owned(),
        "radix-error-radix-low" => "=DECIMAL(\"10\";1)".to_owned(),
        "radix-error-radix-high" => "=DECIMAL(\"10\";37)".to_owned(),
        "radix-error-empty" => "=DECIMAL(\"\";16)".to_owned(),
        "radix-error-arity" => "=BASE(1)".to_owned(),
        "radix-error-text-number" => "=DEC2HEX(\"x\")".to_owned(),
        "radix-refusal-work" => "=BASE(255;16)".to_owned(),
        "radix-refusal-memory" => "=BASE(255;16)".to_owned(),
        "radix-refusal-stack" => "=BASE(255;16)".to_owned(),
        "radix-refusal-cancelled" => "=BASE(255;16)".to_owned(),
        "radix-max-base" => "=BASE(1.7976931348623157e308;16)".to_owned(),
        "radix-max-decimal" => "=DECIMAL(\"FFFFFFFFFFFFF800000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000\";16)".to_owned(),
        "radix-big-base" => "=BASE(9007199254740991;16)".to_owned(),
        "radix-error-large-dec2bin" => "=DEC2BIN(9007199254740991)".to_owned(),
        "radix-error-large-dec2hex" => "=DEC2HEX(9007199254740991)".to_owned(),
        "radix-big-decimal" => "=DECIMAL(\"1FFFFFFFFFFFFF\";16)".to_owned(),
        "bitwise-coerce-text" => "=BITAND(\"6\";\"3\")".to_owned(),
        "bitwise-coerce-fraction" => "=BITAND(6.5;3)".to_owned(),
        "bitwise-error-negative" => "=BITAND(-1;3)".to_owned(),
        "bitwise-error-overflow" => "=BITLSHIFT(281474976710655;1)".to_owned(),
        "bitwise-error-shift" => "=BITLSHIFT(1;54)".to_owned(),
        "bitwise-shift-negative" => "=BITRSHIFT(1;-1)".to_owned(),
        "bitwise-error-numeric" => "=BITAND(\"x\";3)".to_owned(),
        "bitwise-error-arity" => "=BITOR(1)".to_owned(),
        "bitwise-lazy" => "=IF(TRUE();BITAND(6;3);[.A1])".to_owned(),
        "bitwise-lazy-error" => "=IFERROR(BITAND(\"x\";3);7)".to_owned(),
        "bitwise-lazy-selected" => "=IF(FALSE();BITAND(6;3);BITOR(6;3))".to_owned(),
        "failure-reference" => "=[.A1]".to_owned(),
        "failure-array" => "={1;2}".to_owned(),
        "failure-name" => "=Named".to_owned(),
        "failure-work" => "=1+2".to_owned(),
        "failure-memory" => "=\"abcdef\"".to_owned(),
        "failure-stack" => "=((((1))))".to_owned(),
        "failure-cancelled" => "=IF(TRUE();TRUE();FALSE())".to_owned(),
        other => panic!("unknown case {other:?}"),
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

fn numeric_coercion(numbers: usize) -> String {
    let mut value = String::with_capacity(numbers.saturating_mul(2).saturating_add(5));
    value.push_str("=\"1\"");
    for _ in 0..numbers {
        value.push_str("+1");
    }
    value
}

fn utf8_text(length: usize) -> String {
    std::iter::repeat_n('é', length).collect()
}

fn utf8_number_left(length: usize) -> String {
    let text = utf8_text(length);
    let mut value = String::with_capacity(text.len().saturating_add(7));
    value.push_str("=1&\"");
    value.push_str(&text);
    value.push('"');
    value
}

fn escaped_body(length: usize) -> String {
    let mut value = String::with_capacity(length.saturating_mul(4));
    for _ in 0..length {
        value.push('é');
        value.push('"');
        value.push('"');
    }
    value
}

fn escaped_literal(body: &str) -> String {
    let mut value = String::with_capacity(body.len().saturating_add(3));
    value.push_str("=\"");
    value.push_str(body);
    value.push('"');
    value
}

fn escaped_decoded(length: usize) -> String {
    std::iter::repeat_n("é\"", length).collect()
}

fn variadic<F>(name: &str, count: usize, mut operand: F) -> String
where
    F: FnMut(usize) -> String,
{
    let mut value = String::with_capacity(count.saturating_mul(12).saturating_add(name.len() + 3));
    value.push('=');
    value.push_str(name);
    value.push('(');
    for index in 0..count {
        if index != 0 {
            value.push(';');
        }
        value.push_str(&operand(index));
    }
    value.push(')');
    value
}

fn bitwise_chain(name: &str, count: usize) -> String {
    let call = match name {
        "BITAND" => "BITAND(6;3)",
        "BITOR" => "BITOR(6;3)",
        "BITXOR" => "BITXOR(6;3)",
        "BITLSHIFT" => "BITLSHIFT(5;2)",
        "BITRSHIFT" => "BITRSHIFT(20;2)",
        _ => panic!("unknown bitwise function {name:?}"),
    };
    let mut value = String::with_capacity(count.saturating_mul(call.len() + 1) + 1);
    value.push('=');
    for index in 0..count {
        if index != 0 {
            value.push('+');
        }
        value.push_str(call);
    }
    value
}

fn radix_chain(call: &str, count: usize, separator: char) -> String {
    let mut value = String::with_capacity(count.saturating_mul(call.len() + 1) + 1);
    value.push('=');
    for index in 0..count {
        if index != 0 {
            value.push(separator);
        }
        value.push_str(call);
    }
    value
}

#[derive(Clone, Copy, Debug)]
struct Config {
    revision: Revision,
    phase: Phase,
    case: &'static str,
    group: &'static str,
    warmups: usize,
    iterations: usize,
    repeat: Option<usize>,
}

impl Config {
    fn from_args() -> Self {
        let mut revision = None;
        let mut phase = None;
        let mut case = None;
        let mut group = "all";
        let mut warmups = 3;
        let mut iterations = 15;
        let mut repeat = None;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {arg}"));
            match arg.as_str() {
                "--workload" if value == "scalar-evaluation" => {},
                "--workload" => panic!("workload must be scalar-evaluation"),
                "--revision" => revision = Some(Revision::parse(&value)),
                "--group" => {
                    group = match value.as_str() {
                        "comparable" => "comparable",
                        "roman" => "roman",
                        "all" => "all",
                        other => panic!("unknown group {other:?}"),
                    }
                },
                "--phase" => phase = Some(Phase::parse(&value)),
                "--case" => {
                    // The runner owns the finite corpus. Leak one short command-line
                    // copy so the timed code can keep an immutable &'static str.
                    case = Some(Box::leak(value.into_boxed_str()) as &'static str)
                },
                "--warmups" => warmups = value.parse().expect("warmups must be an integer"),
                "--iterations" => {
                    iterations = value.parse().expect("iterations must be an integer")
                },
                "--repeat" => repeat = Some(value.parse().expect("repeat must be an integer")),
                other => panic!("unknown option {other}"),
            }
        }
        let revision = revision.unwrap_or_else(|| panic!("--revision is required"));
        let phase = phase.unwrap_or_else(|| panic!("--phase is required"));
        let case = case.unwrap_or_else(|| panic!("--case is required"));
        assert!(iterations > 0);
        Self {
            revision,
            phase,
            case,
            group,
            warmups,
            iterations,
            repeat,
        }
    }
}

fn limits(case: &str) -> EvaluationLimits {
    match case {
        "failure-work" | "radix-refusal-work" | "roman-refusal-work" => {
            EvaluationLimits::default().with_max_steps(1)
        },
        "failure-memory" | "radix-refusal-memory" | "roman-refusal-memory" => {
            EvaluationLimits::default().with_max_text_bytes(0)
        },
        "failure-stack" | "radix-refusal-stack" | "roman-refusal-stack" => {
            EvaluationLimits::default().with_max_stack_entries(0)
        },
        _ => EvaluationLimits::default(),
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

fn execution(cancelled: bool) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        Arc::<str>::from("ods-formula-array-reference-evaluation-scalar-profile"),
        BudgetLimits::new(
            1_u64 << 50,
            1_u64 << 50,
            1_u64 << 50,
            1_u64 << 50,
            1_u64 << 30,
            1_u64 << 50,
        ),
    );
    let (source, token) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero worker count"),
        NonZeroUsize::new(1).expect("nonzero task count"),
        NonZeroU64::new(1_u64 << 40).expect("nonzero in-flight bytes"),
        1 << 20,
    )
    .expect("profile execution limits are valid");
    if cancelled {
        source.cancel();
    }
    (
        source,
        ExecutionContext::new(budget, token, execution_limits),
    )
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
    refusals: u64,
    checksum: u64,
    output_reserved_bytes: u64,
    failure: &'static str,
}

fn note_failure(slot: &mut &'static str, value: &'static str) {
    if *slot == "none" {
        *slot = value;
    } else if *slot != value {
        *slot = "mixed";
    }
}

fn scalar_checksum(value: &ScalarValue<'_>) -> u64 {
    match value {
        ScalarValue::Number(value) => value.to_bits().rotate_left(17) ^ 0x4e_u64,
        ScalarValue::Logical(value) => u64::from(*value) ^ 0x6c_u64,
        ScalarValue::Text(value) => checksum_bytes(value.as_bytes()) ^ 0x74_u64,
        ScalarValue::Error(error) => scalar_error_code(*error),
        _ => 0x7363_616c_6172_u64,
    }
}

fn scalar_error_code(error: ScalarError) -> u64 {
    match error {
        ScalarError::NotAvailable => 1,
        ScalarError::Name => 2,
        ScalarError::Value => 3,
        ScalarError::DivisionByZero => 4,
        ScalarError::Reference => 5,
        ScalarError::Number => 6,
        ScalarError::Null => 7,
        _ => 255,
    }
}

fn checksum_bytes(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn failure_label(error: &EvaluationFailure) -> &'static str {
    match error {
        EvaluationFailure::Unsupported(kind) => match kind {
            UnsupportedKind::Reference => "unsupported-reference",
            UnsupportedKind::ReferenceOperator => "unsupported-reference-operator",
            UnsupportedKind::Array => "unsupported-array",
            UnsupportedKind::NamedExpression => "unsupported-name",
            UnsupportedKind::Label => "unsupported-label",
            UnsupportedKind::MissingArgument => "unsupported-missing-argument",
            UnsupportedKind::Function => "unsupported-function",
            _ => "unsupported-other",
        },
        EvaluationFailure::InvalidExpression(_) => "invalid-expression",
        EvaluationFailure::Cancelled => "cancelled",
        EvaluationFailure::ResourceLimit(limit) => resource_label(limit.resource),
        EvaluationFailure::Allocation { .. } => "allocation",
        EvaluationFailure::Execution(error) => execution_label(error),
        _ => "evaluation-failure",
    }
}

fn resource_label(resource: Resource) -> &'static str {
    match resource {
        Resource::Memory => "resource-memory",
        Resource::InputBytes => "resource-input-bytes",
        Resource::OutputBytes => "resource-output-bytes",
        Resource::Objects => "resource-objects",
        Resource::Depth => "resource-depth",
        Resource::Work => "resource-work",
        _ => "resource-other",
    }
}

fn execution_label(error: &ExecutionError) -> &'static str {
    match error {
        ExecutionError::Cancelled => "execution-cancelled",
        ExecutionError::ResourceLimit(limit) => resource_label(limit.resource),
        _ => "execution-error",
    }
}

fn validate_value(case: &str, value: &ScalarValue<'_>) -> AnyResult<()> {
    match (expected_value(case), value) {
        (Some(ExpectedValue::Number(expected)), ScalarValue::Number(actual))
            if *actual == expected =>
        {
            Ok(())
        },
        (Some(ExpectedValue::Logical(expected)), ScalarValue::Logical(actual))
            if *actual == expected =>
        {
            Ok(())
        },
        (Some(ExpectedValue::Text(expected)), ScalarValue::Text(actual))
            if actual.as_ref() == expected.as_str() =>
        {
            Ok(())
        },
        (Some(ExpectedValue::FormulaError(expected)), ScalarValue::Error(actual))
            if *actual == expected =>
        {
            Ok(())
        },
        (Some(expected), actual) => Err(format!(
            "{case} returned {actual:?}, expected {}",
            expected_label(&expected)
        )
        .into()),
        (None, actual) => Err(format!("{case} returned unexpected scalar value {actual:?}").into()),
    }
}

fn expected_label(expected: &ExpectedValue) -> &'static str {
    match expected {
        ExpectedValue::Number(_) => "Number",
        ExpectedValue::Logical(_) => "Logical",
        ExpectedValue::Text(_) => "Text",
        ExpectedValue::FormulaError(_) => "formula Error",
    }
}

fn validate_evaluation(
    case: &str,
    revision: Revision,
    result: Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'_>, EvaluationFailure>,
) -> AnyResult<()> {
    match result {
        Ok(value) if expected_success(case, Phase::Evaluate) => {
            validate_value(case, value.value())
        },
        Ok(value) => Err(format!(
            "{revision:?} {case} unexpectedly evaluated successfully: {:?}",
            value.value()
        )
        .into()),
        Err(error) if !expected_success(case, Phase::Evaluate) => {
            let expected = expected_failure(case)
                .unwrap_or_else(|| panic!("refusal case {case} lacks an expected failure"));
            let actual = failure_label(&error);
            if actual == expected {
                Ok(())
            } else {
                Err(format!(
                    "{revision:?} {case} expected failure {expected}, got {actual}: {error}"
                )
                .into())
            }
        },
        Err(error) => {
            Err(format!("{revision:?} {case} unexpectedly refused evaluation: {error}").into())
        },
    }
}

fn preflight(
    case: &str,
    revision: Revision,
    phase: Phase,
    source: &str,
    parsed: Option<&Expression>,
) -> AnyResult<()> {
    match phase {
        Phase::Parse => Expression::parse(source)
            .map(|_| ())
            .map_err(|error| format!("{case} parse preflight failed: {error}").into()),
        Phase::Evaluate => {
            let expression = parsed.expect("evaluation preflight requires a parsed expression");
            let (_source, execution) = execution(matches!(
                case,
                "failure-cancelled" | "radix-refusal-cancelled" | "roman-refusal-cancelled"
            ));
            let context = EvaluationContext::new(&execution);
            validate_evaluation(
                case,
                revision,
                evaluate_scalar(expression, &context, &limits(case)),
            )
        },
        Phase::ParseEvaluate => {
            let expression = Expression::parse(source)
                .map_err(|error| format!("{case} parse preflight failed: {error}"))?;
            let (_source, execution) = execution(matches!(
                case,
                "failure-cancelled" | "radix-refusal-cancelled" | "roman-refusal-cancelled"
            ));
            let context = EvaluationContext::new(&execution);
            validate_evaluation(
                case,
                revision,
                evaluate_scalar(&expression, &context, &limits(case)),
            )
        },
    }
}

fn measure(
    case: &str,
    phase: Phase,
    source: &str,
    parsed: Option<&Expression>,
    repeat: usize,
) -> Sample {
    let limits = limits(case);
    // Context construction is outside the timed region. Keep both handles
    // alive through live-after and counter reads so setup teardown is excluded.
    let execution_state = phase.evaluates().then(|| {
        execution(matches!(
            case,
            "failure-cancelled" | "radix-refusal-cancelled" | "roman-refusal-cancelled"
        ))
    });
    let eval_context = execution_state
        .as_ref()
        .map(|(_, execution)| EvaluationContext::new(execution));
    let live_before = reset_observer();
    let started = Instant::now();
    let mut successes = 0_u64;
    let mut refusals = 0_u64;
    let mut checksum = 0_u64;
    let mut output_reserved_bytes = 0_u64;
    let mut failure = "none";

    for _ in 0..repeat {
        match phase {
            Phase::Parse => match Expression::parse(black_box(source)) {
                Ok(expression) => {
                    successes += 1;
                    checksum = checksum
                        .wrapping_add(expression.source().len() as u64)
                        .wrapping_add(expression.node_count() as u64)
                        .wrapping_add(expression.edge_count() as u64);
                    drop(black_box(expression));
                },
                Err(error) => {
                    refusals += 1;
                    note_failure(&mut failure, "parse-error");
                    black_box(error);
                },
            },
            Phase::Evaluate => {
                let expression = parsed.expect("evaluation phase requires a parsed expression");
                let context = eval_context.expect("evaluation phase requires a context");
                match evaluate_scalar(expression, &context, &limits) {
                    Ok(value) => {
                        successes += 1;
                        output_reserved_bytes =
                            output_reserved_bytes.max(value.reserved_output_bytes() as u64);
                        checksum = checksum.wrapping_add(scalar_checksum(value.value()));
                        black_box(checksum);
                        drop(value);
                    },
                    Err(error) => {
                        refusals += 1;
                        note_failure(&mut failure, failure_label(&error));
                        black_box(error);
                    },
                }
            },
            Phase::ParseEvaluate => {
                let context = eval_context.expect("parse-evaluate phase requires a context");
                match Expression::parse(black_box(source)) {
                    Ok(expression) => match evaluate_scalar(&expression, &context, &limits) {
                        Ok(value) => {
                            successes += 1;
                            output_reserved_bytes =
                                output_reserved_bytes.max(value.reserved_output_bytes() as u64);
                            checksum = checksum.wrapping_add(scalar_checksum(value.value()));
                            black_box(checksum);
                            drop(value);
                        },
                        Err(error) => {
                            refusals += 1;
                            note_failure(&mut failure, failure_label(&error));
                            black_box(error);
                        },
                    },
                    Err(error) => {
                        refusals += 1;
                        note_failure(&mut failure, "parse-error");
                        black_box(error);
                    },
                }
            },
        }
    }
    black_box((successes, refusals, checksum));
    let elapsed = started.elapsed();
    let live_after = LIVE_BYTES.load(Ordering::Relaxed);
    let sample = Sample {
        elapsed,
        alloc_calls: ALLOC_CALLS.load(Ordering::Relaxed),
        dealloc_calls: DEALLOC_CALLS.load(Ordering::Relaxed),
        requested_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
        released_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
        live_before_bytes: live_before,
        live_after_bytes: live_after,
        peak_live_delta: PEAK_LIVE_BYTES
            .load(Ordering::Relaxed)
            .saturating_sub(live_before),
        successes,
        refusals,
        checksum,
        output_reserved_bytes,
        failure,
    };
    let _ = eval_context;
    drop(execution_state);
    sample
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

fn aggregate_failure(samples: &[Sample]) -> &'static str {
    let mut result = "none";
    for sample in samples {
        if result == "none" {
            result = sample.failure;
        } else if result != sample.failure {
            return "mixed";
        }
    }
    result
}

fn main() -> AnyResult<()> {
    let config = Config::from_args();
    let source = input(config.case);
    let parsed = if config.phase == Phase::Evaluate {
        Some(Expression::parse(&source)?)
    } else {
        None
    };
    let repeat = config.repeat.unwrap_or_else(|| default_repeat(config.case));
    assert!(repeat > 0);
    println!(
        "config workload=scalar-evaluation revision={} group={} phase={} case={} input_bytes={} repeat={} warmups={} iterations={} expected_success={}",
        config.revision.label(),
        config.group,
        config.phase.label(),
        config.case,
        source.len(),
        repeat,
        config.warmups,
        config.iterations,
        expected_success(config.case, config.phase),
    );
    preflight(
        config.case,
        config.revision,
        config.phase,
        &source,
        parsed.as_ref(),
    )?;
    for _ in 0..config.warmups {
        let _ = measure(config.case, config.phase, &source, parsed.as_ref(), repeat);
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(
            config.case,
            config.phase,
            &source,
            parsed.as_ref(),
            repeat,
        ));
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
    let alloc_calls_max = alloc_calls.clone();
    let mut dealloc_calls: Vec<u64> = samples.iter().map(|sample| sample.dealloc_calls).collect();
    let dealloc_calls_max = dealloc_calls.clone();
    let mut requested: Vec<u64> = samples
        .iter()
        .map(|sample| sample.requested_bytes)
        .collect();
    let requested_max = requested.clone();
    let mut released: Vec<u64> = samples.iter().map(|sample| sample.released_bytes).collect();
    let released_max = released.clone();
    let mut live_before: Vec<u64> = samples
        .iter()
        .map(|sample| sample.live_before_bytes)
        .collect();
    let mut live_after: Vec<u64> = samples
        .iter()
        .map(|sample| sample.live_after_bytes)
        .collect();
    let live_after_max = live_after.clone();
    let mut peak_live: Vec<u64> = samples
        .iter()
        .map(|sample| sample.peak_live_delta)
        .collect();
    let peak_live_max = peak_live.clone();
    let mut successes: Vec<u64> = samples.iter().map(|sample| sample.successes).collect();
    let successes_max = successes.clone();
    let mut refusals: Vec<u64> = samples.iter().map(|sample| sample.refusals).collect();
    let refusals_max = refusals.clone();
    let mut checksums: Vec<u64> = samples.iter().map(|sample| sample.checksum).collect();
    let checksums_max = checksums.clone();
    let mut output_reserved: Vec<u64> = samples
        .iter()
        .map(|sample| sample.output_reserved_bytes)
        .collect();
    let output_reserved_max = output_reserved.clone();
    let p50_u64 = |values: &mut Vec<u64>| percentile_u64(values, 50, 100);
    let max_u64 = |values: &Vec<u64>| values.iter().copied().max().unwrap_or(0);
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} alloc_calls_p50={} alloc_calls_max={} dealloc_calls_p50={} dealloc_calls_max={} requested_bytes_p50={} requested_bytes_max={} released_bytes_p50={} released_bytes_max={} live_before_p50={} live_after_p50={} live_after_max={} peak_live_delta_p50={} peak_live_delta_max={} successes_p50={} successes_max={} refusals_p50={} refusals_max={} checksum_p50={} checksum_max={} output_reserved_bytes_p50={} output_reserved_bytes_max={} failure={}",
        mean,
        p50,
        p95,
        p99,
        p50_u64(&mut alloc_calls),
        max_u64(&alloc_calls_max),
        p50_u64(&mut dealloc_calls),
        max_u64(&dealloc_calls_max),
        p50_u64(&mut requested),
        max_u64(&requested_max),
        p50_u64(&mut released),
        max_u64(&released_max),
        p50_u64(&mut live_before),
        p50_u64(&mut live_after),
        max_u64(&live_after_max),
        p50_u64(&mut peak_live),
        max_u64(&peak_live_max),
        p50_u64(&mut successes),
        max_u64(&successes_max),
        p50_u64(&mut refusals),
        max_u64(&refusals_max),
        p50_u64(&mut checksums),
        max_u64(&checksums_max),
        p50_u64(&mut output_reserved),
        max_u64(&output_reserved_max),
        aggregate_failure(&samples),
    );
    Ok(())
}
