use std::{
    alloc::{GlobalAlloc, Layout, System},
    env,
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

#[derive(Clone, Copy, Debug)]
enum Case {
    Flat64,
    Flat256,
    Flat1024,
    Flat4096,
    Utf8TextText64,
    Utf8TextText256,
    Utf8TextText1024,
    Utf8TextText4096,
    Utf8NumberLeft64,
    Utf8NumberLeft256,
    Utf8NumberLeft1024,
    Utf8NumberLeft4096,
    Utf8TextNumberRight64,
    Utf8TextNumberRight256,
    Utf8TextNumberRight1024,
    Utf8TextNumberRight4096,
    EscapedTextText64,
    EscapedTextText256,
    EscapedTextText1024,
    EscapedTextText4096,
    EscapedTextNumberRight64,
    EscapedTextNumberRight256,
    EscapedTextNumberRight1024,
    EscapedTextNumberRight4096,
    Coerce64,
    Coerce256,
    Coerce1024,
    Coerce4096,
    LongName4096,
    Array4096,
    Reference,
    ExactStep,
    WorkLimit,
    TextLimit,
    Cancelled,
}

impl Case {
    fn parse(value: &str) -> Self {
        match value {
            "eval-flat-64" => Self::Flat64,
            "eval-flat-256" => Self::Flat256,
            "eval-flat-1024" => Self::Flat1024,
            "eval-flat-4096" => Self::Flat4096,
            "eval-utf8-text-text-64" => Self::Utf8TextText64,
            "eval-utf8-text-text-256" => Self::Utf8TextText256,
            "eval-utf8-text-text-1024" => Self::Utf8TextText1024,
            "eval-utf8-text-text-4096" => Self::Utf8TextText4096,
            "eval-utf8-number-left-64" => Self::Utf8NumberLeft64,
            "eval-utf8-number-left-256" => Self::Utf8NumberLeft256,
            "eval-utf8-number-left-1024" => Self::Utf8NumberLeft1024,
            "eval-utf8-number-left-4096" => Self::Utf8NumberLeft4096,
            "eval-utf8-text-number-right-64" => Self::Utf8TextNumberRight64,
            "eval-utf8-text-number-right-256" => Self::Utf8TextNumberRight256,
            "eval-utf8-text-number-right-1024" => Self::Utf8TextNumberRight1024,
            "eval-utf8-text-number-right-4096" => Self::Utf8TextNumberRight4096,
            "eval-escaped-text-text-64" => Self::EscapedTextText64,
            "eval-escaped-text-text-256" => Self::EscapedTextText256,
            "eval-escaped-text-text-1024" => Self::EscapedTextText1024,
            "eval-escaped-text-text-4096" => Self::EscapedTextText4096,
            "eval-escaped-text-number-right-64" => Self::EscapedTextNumberRight64,
            "eval-escaped-text-number-right-256" => Self::EscapedTextNumberRight256,
            "eval-escaped-text-number-right-1024" => Self::EscapedTextNumberRight1024,
            "eval-escaped-text-number-right-4096" => Self::EscapedTextNumberRight4096,
            "eval-coerce-64" => Self::Coerce64,
            "eval-coerce-256" => Self::Coerce256,
            "eval-coerce-1024" => Self::Coerce1024,
            "eval-coerce-4096" => Self::Coerce4096,
            "eval-long-name-4096" => Self::LongName4096,
            "eval-array-4096" => Self::Array4096,
            "eval-reference" => Self::Reference,
            "eval-exact-step" => Self::ExactStep,
            "eval-work-limit" => Self::WorkLimit,
            "eval-text-limit" => Self::TextLimit,
            "eval-cancelled" => Self::Cancelled,
            other => panic!("unknown case {other:?}"),
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Flat64 => "eval-flat-64",
            Self::Flat256 => "eval-flat-256",
            Self::Flat1024 => "eval-flat-1024",
            Self::Flat4096 => "eval-flat-4096",
            Self::Utf8TextText64 => "eval-utf8-text-text-64",
            Self::Utf8TextText256 => "eval-utf8-text-text-256",
            Self::Utf8TextText1024 => "eval-utf8-text-text-1024",
            Self::Utf8TextText4096 => "eval-utf8-text-text-4096",
            Self::Utf8NumberLeft64 => "eval-utf8-number-left-64",
            Self::Utf8NumberLeft256 => "eval-utf8-number-left-256",
            Self::Utf8NumberLeft1024 => "eval-utf8-number-left-1024",
            Self::Utf8NumberLeft4096 => "eval-utf8-number-left-4096",
            Self::Utf8TextNumberRight64 => "eval-utf8-text-number-right-64",
            Self::Utf8TextNumberRight256 => "eval-utf8-text-number-right-256",
            Self::Utf8TextNumberRight1024 => "eval-utf8-text-number-right-1024",
            Self::Utf8TextNumberRight4096 => "eval-utf8-text-number-right-4096",
            Self::EscapedTextText64 => "eval-escaped-text-text-64",
            Self::EscapedTextText256 => "eval-escaped-text-text-256",
            Self::EscapedTextText1024 => "eval-escaped-text-text-1024",
            Self::EscapedTextText4096 => "eval-escaped-text-text-4096",
            Self::EscapedTextNumberRight64 => "eval-escaped-text-number-right-64",
            Self::EscapedTextNumberRight256 => "eval-escaped-text-number-right-256",
            Self::EscapedTextNumberRight1024 => "eval-escaped-text-number-right-1024",
            Self::EscapedTextNumberRight4096 => "eval-escaped-text-number-right-4096",
            Self::Coerce64 => "eval-coerce-64",
            Self::Coerce256 => "eval-coerce-256",
            Self::Coerce1024 => "eval-coerce-1024",
            Self::Coerce4096 => "eval-coerce-4096",
            Self::LongName4096 => "eval-long-name-4096",
            Self::Array4096 => "eval-array-4096",
            Self::Reference => "eval-reference",
            Self::ExactStep => "eval-exact-step",
            Self::WorkLimit => "eval-work-limit",
            Self::TextLimit => "eval-text-limit",
            Self::Cancelled => "eval-cancelled",
        }
    }

    const fn scale(self) -> Option<usize> {
        match self {
            Self::Flat64
            | Self::Utf8TextText64
            | Self::Utf8NumberLeft64
            | Self::Utf8TextNumberRight64
            | Self::EscapedTextText64
            | Self::EscapedTextNumberRight64
            | Self::Coerce64 => Some(64),
            Self::Flat256
            | Self::Utf8TextText256
            | Self::Utf8NumberLeft256
            | Self::Utf8TextNumberRight256
            | Self::EscapedTextText256
            | Self::EscapedTextNumberRight256
            | Self::Coerce256 => Some(256),
            Self::Flat1024
            | Self::Utf8TextText1024
            | Self::Utf8NumberLeft1024
            | Self::Utf8TextNumberRight1024
            | Self::EscapedTextText1024
            | Self::EscapedTextNumberRight1024
            | Self::Coerce1024 => Some(1024),
            Self::Flat4096
            | Self::Utf8TextText4096
            | Self::Utf8NumberLeft4096
            | Self::Utf8TextNumberRight4096
            | Self::EscapedTextText4096
            | Self::EscapedTextNumberRight4096
            | Self::Coerce4096
            | Self::LongName4096
            | Self::Array4096 => Some(4096),
            Self::Reference
            | Self::ExactStep
            | Self::WorkLimit
            | Self::TextLimit
            | Self::Cancelled => None,
        }
    }

    fn default_repeat(self) -> usize {
        match self.scale() {
            Some(64) => 64,
            Some(256) => 32,
            Some(1024) => 8,
            Some(4096) => 2,
            Some(_) => 2,
            None => 128,
        }
    }

    fn input(self) -> String {
        match self {
            Self::Flat64 => flat_chain(64),
            Self::Flat256 => flat_chain(256),
            Self::Flat1024 => flat_chain(1024),
            Self::Flat4096 => flat_chain(4096),
            Self::Utf8TextText64 => utf8_text_text(64),
            Self::Utf8TextText256 => utf8_text_text(256),
            Self::Utf8TextText1024 => utf8_text_text(1024),
            Self::Utf8TextText4096 => utf8_text_text(4096),
            Self::Utf8NumberLeft64 => utf8_number_left(64),
            Self::Utf8NumberLeft256 => utf8_number_left(256),
            Self::Utf8NumberLeft1024 => utf8_number_left(1024),
            Self::Utf8NumberLeft4096 => utf8_number_left(4096),
            Self::Utf8TextNumberRight64 => utf8_text_number_right(64),
            Self::Utf8TextNumberRight256 => utf8_text_number_right(256),
            Self::Utf8TextNumberRight1024 => utf8_text_number_right(1024),
            Self::Utf8TextNumberRight4096 => utf8_text_number_right(4096),
            Self::EscapedTextText64 => escaped_text_text(64),
            Self::EscapedTextText256 => escaped_text_text(256),
            Self::EscapedTextText1024 => escaped_text_text(1024),
            Self::EscapedTextText4096 => escaped_text_text(4096),
            Self::EscapedTextNumberRight64 => escaped_text_number_right(64),
            Self::EscapedTextNumberRight256 => escaped_text_number_right(256),
            Self::EscapedTextNumberRight1024 => escaped_text_number_right(1024),
            Self::EscapedTextNumberRight4096 => escaped_text_number_right(4096),
            Self::Coerce64 => numeric_coercion(64),
            Self::Coerce256 => numeric_coercion(256),
            Self::Coerce1024 => numeric_coercion(1024),
            Self::Coerce4096 => numeric_coercion(4096),
            Self::LongName4096 => long_name(4096),
            Self::Array4096 => array(4096),
            Self::Reference => "[.A1]".to_owned(),
            Self::ExactStep => "=1".to_owned(),
            Self::TextLimit => "=\"abcdef\"".to_owned(),
            Self::WorkLimit | Self::Cancelled => "=1+2".to_owned(),
        }
    }

    fn limits(self) -> EvaluationLimits {
        match self {
            Self::WorkLimit => EvaluationLimits::default().with_max_steps(1),
            Self::ExactStep => EvaluationLimits::default().with_max_steps(2),
            Self::TextLimit => EvaluationLimits::default().with_max_text_bytes(0),
            _ => EvaluationLimits::default(),
        }
    }

    const fn expected_success(self, phase: Phase) -> bool {
        if !phase.evaluates() {
            return true;
        }
        !matches!(
            self,
            Self::LongName4096
                | Self::Array4096
                | Self::Reference
                | Self::WorkLimit
                | Self::TextLimit
                | Self::Cancelled
        )
    }

    const fn cancelled(self) -> bool {
        matches!(self, Self::Cancelled)
    }

    fn expected_value(self) -> Option<ExpectedValue> {
        match self {
            Self::Flat64 | Self::Flat256 | Self::Flat1024 | Self::Flat4096 => Some(
                ExpectedValue::Number(self.scale().expect("flat scale") as f64),
            ),
            Self::Utf8TextText64
            | Self::Utf8TextText256
            | Self::Utf8TextText1024
            | Self::Utf8TextText4096 => {
                let text = utf8_text(self.scale().expect("UTF-8 text scale"));
                Some(ExpectedValue::Text(format!("{text}{text}")))
            },
            Self::Utf8NumberLeft64
            | Self::Utf8NumberLeft256
            | Self::Utf8NumberLeft1024
            | Self::Utf8NumberLeft4096 => {
                let text = utf8_text(self.scale().expect("UTF-8 left scale"));
                Some(ExpectedValue::Text(format!("1{text}")))
            },
            Self::Utf8TextNumberRight64
            | Self::Utf8TextNumberRight256
            | Self::Utf8TextNumberRight4096
            | Self::Utf8TextNumberRight1024 => {
                let text = utf8_text(self.scale().expect("UTF-8 right scale"));
                Some(ExpectedValue::Text(format!("{text}1")))
            },
            Self::EscapedTextText64
            | Self::EscapedTextText256
            | Self::EscapedTextText1024
            | Self::EscapedTextText4096 => {
                let text = escaped_decoded_text(self.scale().expect("escaped text scale"));
                Some(ExpectedValue::Text(format!("{text}{text}")))
            },
            Self::EscapedTextNumberRight64
            | Self::EscapedTextNumberRight256
            | Self::EscapedTextNumberRight1024
            | Self::EscapedTextNumberRight4096 => {
                let text = escaped_decoded_text(self.scale().expect("escaped right scale"));
                Some(ExpectedValue::Text(format!("{text}1")))
            },
            Self::Coerce64 | Self::Coerce256 | Self::Coerce1024 | Self::Coerce4096 => Some(
                ExpectedValue::Number(self.scale().expect("coercion scale") as f64 + 1.0),
            ),
            Self::ExactStep => Some(ExpectedValue::Number(1.0)),
            Self::LongName4096
            | Self::Array4096
            | Self::Reference
            | Self::WorkLimit
            | Self::TextLimit
            | Self::Cancelled => None,
        }
    }

    const fn expected_failure(self) -> Option<&'static str> {
        match self {
            Self::LongName4096 => Some("unsupported-name"),
            Self::Array4096 => Some("unsupported-array"),
            Self::Reference => Some("unsupported-reference"),
            Self::WorkLimit => Some("resource-work"),
            Self::TextLimit => Some("resource-memory"),
            Self::Cancelled => Some("cancelled"),
            _ => None,
        }
    }
}

enum ExpectedValue {
    Number(f64),
    Text(String),
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

fn utf8_text(length: usize) -> String {
    std::iter::repeat_n('é', length).collect()
}

fn utf8_text_text(length: usize) -> String {
    let text = utf8_text(length);
    let mut value = String::with_capacity(text.len().saturating_mul(2).saturating_add(5));
    value.push_str("=\"");
    value.push_str(&text);
    value.push_str("\"&\"");
    value.push_str(&text);
    value.push('"');
    value
}

fn utf8_number_left(length: usize) -> String {
    let text = utf8_text(length);
    let mut value = String::with_capacity(text.len().saturating_add(7));
    value.push_str("=1&\"");
    value.push_str(&text);
    value.push('"');
    value
}

fn utf8_text_number_right(length: usize) -> String {
    let text = utf8_text(length);
    let mut value = String::with_capacity(text.len().saturating_add(7));
    value.push_str("=\"");
    value.push_str(&text);
    value.push_str("\"&1");
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

fn escaped_text_text(length: usize) -> String {
    let body = escaped_body(length);
    let mut value = String::with_capacity(body.len().saturating_mul(2).saturating_add(5));
    value.push_str("=\"");
    value.push_str(&body);
    value.push_str("\"&\"");
    value.push_str(&body);
    value.push('"');
    value
}

fn escaped_text_number_right(length: usize) -> String {
    let body = escaped_body(length);
    let mut value = String::with_capacity(body.len().saturating_add(7));
    value.push_str("=\"");
    value.push_str(&body);
    value.push_str("\"&1");
    value
}

fn escaped_decoded_text(length: usize) -> String {
    std::iter::repeat_n("é\"", length).collect()
}

fn numeric_coercion(numbers: usize) -> String {
    let mut value = String::with_capacity(numbers.saturating_mul(2).saturating_add(5));
    value.push_str("=\"1\"");
    for _ in 0..numbers {
        value.push_str("+1");
    }
    value
}

fn long_name(length: usize) -> String {
    let mut value = String::with_capacity(length.saturating_add(1));
    value.push('=');
    value.extend(std::iter::repeat_n('N', length));
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

#[derive(Clone, Copy, Debug)]
struct Config {
    case: Case,
    phase: Phase,
    warmups: usize,
    iterations: usize,
    repeat: Option<usize>,
}

impl Config {
    fn from_args() -> Self {
        let mut case = None;
        let mut phase = None;
        let mut warmups = 3;
        let mut iterations = 15;
        let mut repeat = None;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let value = args
                .next()
                .unwrap_or_else(|| panic!("missing value for {arg}"));
            match arg.as_str() {
                "--workload" if value == "evaluation" => {},
                "--workload" => panic!("workload must be evaluation"),
                "--phase" => phase = Some(Phase::parse(&value)),
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
        let phase = phase.unwrap_or_else(|| panic!("--phase is required"));
        assert!(iterations > 0);
        Self {
            case,
            phase,
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

fn execution(cancelled: bool) -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        Arc::<str>::from("ods-formula-scalar-evaluation-profile"),
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
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero worker count"),
        NonZeroUsize::new(1).expect("nonzero task count"),
        NonZeroU64::new(1_u64 << 40).expect("nonzero in-flight bytes"),
        1 << 20,
    )
    .expect("profile execution limits are valid");
    if cancelled {
        source.cancel();
    }
    (source, ExecutionContext::new(budget, token, limits))
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

fn validate_value(case: Case, value: &ScalarValue<'_>) -> AnyResult<()> {
    match (case.expected_value(), value) {
        (Some(ExpectedValue::Number(expected)), ScalarValue::Number(actual))
            if *actual == expected =>
        {
            Ok(())
        },
        (Some(ExpectedValue::Text(expected)), ScalarValue::Text(actual))
            if actual.as_ref() == expected.as_str() =>
        {
            Ok(())
        },
        (Some(ExpectedValue::Number(expected)), actual) => Err(format!(
            "{} expected Number({expected}), got {actual:?}",
            case.label()
        )
        .into()),
        (Some(ExpectedValue::Text(expected)), actual) => Err(format!(
            "{} expected exact text of {} bytes, got {actual:?}",
            case.label(),
            expected.len()
        )
        .into()),
        (None, actual) => Err(format!(
            "{} returned unexpected scalar value {actual:?}",
            case.label()
        )
        .into()),
    }
}

fn validate_evaluation(
    case: Case,
    result: Result<litchi_ods::codec::formula::evaluation::EvaluatedScalar<'_>, EvaluationFailure>,
) -> AnyResult<()> {
    match result {
        Ok(value) if case.expected_success(Phase::Evaluate) => validate_value(case, value.value()),
        Ok(value) => Err(format!(
            "{} unexpectedly evaluated successfully: {:?}",
            case.label(),
            value.value()
        )
        .into()),
        Err(error) if !case.expected_success(Phase::Evaluate) => {
            let expected = case
                .expected_failure()
                .expect("refusal case has a failure label");
            let actual = failure_label(&error);
            if actual == expected {
                Ok(())
            } else {
                Err(format!(
                    "{} expected failure {expected}, got {actual}: {error}",
                    case.label()
                )
                .into())
            }
        },
        Err(error) => {
            Err(format!("{} unexpectedly refused evaluation: {error}", case.label()).into())
        },
    }
}

fn preflight(case: Case, phase: Phase, input: &str, parsed: Option<&Expression>) -> AnyResult<()> {
    match phase {
        Phase::Parse => Expression::parse(input)
            .map(|_| ())
            .map_err(|error| format!("{} parse preflight failed: {error}", case.label()).into()),
        Phase::Evaluate => {
            let expression = parsed.expect("evaluation preflight requires a parsed expression");
            let (_source, execution) = execution(case.cancelled());
            let context = EvaluationContext::new(&execution);
            validate_evaluation(case, evaluate_scalar(expression, &context, &case.limits()))
        },
        Phase::ParseEvaluate => {
            let expression = Expression::parse(input)
                .map_err(|error| format!("{} parse preflight failed: {error}", case.label()))?;
            let (_source, execution) = execution(case.cancelled());
            let context = EvaluationContext::new(&execution);
            validate_evaluation(case, evaluate_scalar(&expression, &context, &case.limits()))
        },
    }
}

fn measure(
    case: Case,
    phase: Phase,
    input: &str,
    parsed: Option<&Expression>,
    repeat: usize,
) -> Sample {
    let limits = case.limits();
    // Context construction is intentionally outside the timed region. Keep
    // both handles alive through live-after and allocation-counter reads so
    // setup teardown cannot contaminate the operation measurements.
    let execution_state = phase.evaluates().then(|| execution(case.cancelled()));
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
            Phase::Parse => match Expression::parse(black_box(input)) {
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
                match Expression::parse(black_box(input)) {
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
    let input = config.case.input();
    let parsed = if config.phase == Phase::Evaluate {
        Some(Expression::parse(&input)?)
    } else {
        None
    };
    let repeat = config
        .repeat
        .unwrap_or_else(|| config.case.default_repeat());
    assert!(repeat > 0);
    println!(
        "config workload=evaluation phase={} case={} input_bytes={} repeat={} warmups={} iterations={} expected_success={}",
        config.phase.label(),
        config.case.label(),
        input.len(),
        repeat,
        config.warmups,
        config.iterations,
        config.case.expected_success(config.phase),
    );
    preflight(config.case, config.phase, &input, parsed.as_ref())?;
    for _ in 0..config.warmups {
        let _ = measure(
            config.case,
            config.phase,
            &input,
            parsed.as_ref(),
            repeat,
        );
    }
    let mut samples = Vec::with_capacity(config.iterations);
    for _ in 0..config.iterations {
        samples.push(measure(
            config.case,
            config.phase,
            &input,
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
    let mut refusals: Vec<u64> = samples.iter().map(|sample| sample.refusals).collect();
    let mut checksums: Vec<u64> = samples.iter().map(|sample| sample.checksum).collect();
    let mut output_reserved: Vec<u64> = samples
        .iter()
        .map(|sample| sample.output_reserved_bytes)
        .collect();
    let p50_u64 = |values: &mut Vec<u64>| percentile_u64(values, 50, 100);
    let max_u64 = |values: &Vec<u64>| values.iter().copied().max().unwrap_or(0);
    println!(
        "result mean_ns={} p50_ns={} p95_ns={} p99_ns={} alloc_calls_p50={} alloc_calls_max={} dealloc_calls_p50={} dealloc_calls_max={} requested_bytes_p50={} requested_bytes_max={} released_bytes_p50={} released_bytes_max={} live_before_p50={} live_after_p50={} live_after_max={} peak_live_delta_p50={} peak_live_delta_max={} successes_p50={} successes_max={} refusals_p50={} refusals_max={} checksum_p50={} checksum_max={} output_reserved_bytes_p50={} output_reserved_bytes_max={} failure={}",
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
        p50_u64(&mut refusals),
        max_u64(&refusals),
        p50_u64(&mut checksums),
        max_u64(&checksums),
        p50_u64(&mut output_reserved),
        max_u64(&output_reserved),
        aggregate_failure(&samples),
    );
    Ok(())
}
